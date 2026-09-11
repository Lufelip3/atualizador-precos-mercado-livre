"""Atualiza preço e estoque do Mercado Livre diretamente pelo SKU.

Entradas em ``inputs/``:
- CSV do Bling com Código, Preço e Estoque.
- XLSX oficial do Mercado Livre com SKU, PRICE e QUANTITY.

Somente células editáveis de PRICE e QUANTITY com valores diferentes são
alteradas. O restante da estrutura do XLSX é preservado.
"""

import csv
import glob
import os
import posixpath
import re
import shutil
import sys
import tempfile
import unicodedata
import zipfile
import xml.etree.ElementTree as ET
from datetime import datetime

import openpyxl


sys.stdout.reconfigure(encoding="utf-8")
os.chdir(os.path.dirname(os.path.abspath(__file__)))

PASTA_INPUTS = "inputs"
PASTA_SAIDA = "saidas"

COR_VERMELHO = "\033[91m"
COR_VERDE = "\033[92m"
COR_LARANJA = "\033[33m"
COR_RESET = "\033[0m"
PADRAO_ITEM_ID = re.compile(r"^MLB\d+$")


def remover_acentos(valor) -> str:
    texto = unicodedata.normalize("NFD", str(valor))
    return "".join(c for c in texto if unicodedata.category(c) != "Mn").lower()


def normalizar_coluna(valor) -> str:
    return re.sub(r"[^a-z0-9]+", " ", remover_acentos(valor)).strip()


def limpar_sku(valor) -> str:
    texto = str(valor or "").strip().strip("\t")
    return texto[:-2] if texto.endswith(".0") else texto


def converter_numero_br(valor) -> float:
    texto = str(valor or "").strip().replace("R$", "").replace(" ", "")
    if not texto:
        raise ValueError("valor numérico vazio")
    if "," in texto:
        texto = texto.replace(".", "").replace(",", ".")
    return float(texto)


def valores_numericos_iguais(atual, novo, tolerancia: float = 0.0) -> bool:
    try:
        return abs(converter_numero_br(atual) - float(novo)) <= tolerancia
    except (TypeError, ValueError):
        return False


def celula_editavel(celula) -> bool:
    """Segue a instrução do ML: editar somente células brancas/desbloqueadas."""
    if celula.protection.locked:
        return False
    preenchimento = celula.fill
    if preenchimento.fill_type in (None, "none"):
        return True
    cor = preenchimento.fgColor
    if cor.type == "rgb" and cor.rgb:
        return cor.rgb.upper().endswith("FFFFFF")
    return cor.type == "theme" and cor.theme == 0 and cor.tint == 0


def _cabecalho_csv(caminho: str) -> list[str]:
    with open(caminho, "r", encoding="utf-8-sig", newline="") as arquivo:
        return next(csv.reader(arquivo, delimiter=";"), [])


def _encontrar_coluna(colunas: list[str], termos: tuple[str, ...], excluir=()) -> str | None:
    for coluna in colunas:
        nome = normalizar_coluna(coluna)
        if all(t in nome for t in termos) and not any(t in nome for t in excluir):
            return coluna
    return None


def localizar_arquivo_bling(exibir_log: bool = True) -> str:
    candidatos = []
    for caminho in glob.glob(os.path.join(PASTA_INPUTS, "*.csv")):
        try:
            colunas = _cabecalho_csv(caminho)
            codigo = _encontrar_coluna(colunas, ("codigo",)) or _encontrar_coluna(colunas, ("sku",))
            preco = _encontrar_coluna(colunas, ("preco",), ("custo", "promocional", "compra"))
            estoque = _encontrar_coluna(colunas, ("estoque",), ("minimo", "maximo"))
            if codigo and preco and estoque:
                candidatos.append(caminho)
        except (OSError, UnicodeError):
            continue
    if not candidatos:
        raise FileNotFoundError("CSV do Bling com Código, Preço e Estoque não encontrado em inputs/.")
    escolhido = max(candidatos, key=os.path.getmtime)
    if exibir_log:
        print(f"Catálogo Bling       : {escolhido}")
    return escolhido


def _localizar_cabecalho_ml(ws):
    for linha in ws.iter_rows(min_row=1, max_row=min(10, ws.max_row)):
        valores = [celula.value for celula in linha]
        if all(campo in valores for campo in ("SKU", "QUANTITY", "PRICE")):
            return linha[0].row, {valor: i + 1 for i, valor in enumerate(valores) if valor}
    return None, {}


def localizar_arquivo_anuncios_ml(exibir_log: bool = True) -> tuple[str, str]:
    candidatos = []
    for caminho in glob.glob(os.path.join(PASTA_INPUTS, "*.xlsx")):
        if os.path.basename(caminho).startswith("~$"):
            continue
        try:
            wb = openpyxl.load_workbook(caminho, read_only=True, data_only=False)
            for ws in wb.worksheets:
                linha, _ = _localizar_cabecalho_ml(ws)
                if linha:
                    candidatos.append((caminho, ws.title))
                    break
            wb.close()
        except Exception:
            continue
    if not candidatos:
        raise FileNotFoundError(
            "Planilha oficial do Mercado Livre com SKU, QUANTITY e PRICE não encontrada em inputs/."
        )
    caminho, aba = max(candidatos, key=lambda item: os.path.getmtime(item[0]))
    if exibir_log:
        print(f"Anúncios Mercado Livre: {caminho}  [aba: {aba}]")
    return caminho, aba


def carregar_catalogo_bling(caminho: str) -> dict[str, dict]:
    with open(caminho, "r", encoding="utf-8-sig", newline="") as arquivo:
        leitor = csv.DictReader(arquivo, delimiter=";")
        colunas = leitor.fieldnames or []
        col_sku = _encontrar_coluna(colunas, ("codigo",)) or _encontrar_coluna(colunas, ("sku",))
        col_preco = _encontrar_coluna(colunas, ("preco",), ("custo", "promocional", "compra"))
        col_estoque = _encontrar_coluna(colunas, ("estoque",), ("minimo", "maximo"))
        if not col_sku or not col_preco or not col_estoque:
            raise ValueError("Não foi possível identificar Código, Preço e Estoque no CSV do Bling.")

        catalogo = {}
        duplicados = set()
        for linha in leitor:
            sku = limpar_sku(linha.get(col_sku))
            if not sku:
                continue
            try:
                preco = round(converter_numero_br(linha.get(col_preco)), 2)
                estoque = int(round(converter_numero_br(linha.get(col_estoque))))
            except ValueError:
                print(f"{COR_VERMELHO}SKU {sku} ignorado: preço ou estoque inválido no CSV.{COR_RESET}")
                continue
            if sku in catalogo:
                duplicados.add(sku)
            catalogo[sku] = {
                "preco": preco,
                "estoque": estoque,
            }

    if duplicados:
        print(
            f"{COR_LARANJA}{len(duplicados)} SKU(s) duplicado(s) no CSV; "
            f"foi usada a última ocorrência.{COR_RESET}"
        )
    return catalogo


def _caminho_xml_da_aba(arquivo_zip: zipfile.ZipFile, nome_aba: str) -> str:
    ns_planilha = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    ns_relacao = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    ns_pacote = "http://schemas.openxmlformats.org/package/2006/relationships"
    workbook = ET.fromstring(arquivo_zip.read("xl/workbook.xml"))
    id_relacao = next(
        (
            aba.get(f"{{{ns_relacao}}}id")
            for aba in workbook.findall(f".//{{{ns_planilha}}}sheet")
            if aba.get("name") == nome_aba
        ),
        None,
    )
    if not id_relacao:
        raise ValueError(f"Aba '{nome_aba}' não encontrada no XLSX.")
    relacoes = ET.fromstring(arquivo_zip.read("xl/_rels/workbook.xml.rels"))
    for relacao in relacoes.findall(f"{{{ns_pacote}}}Relationship"):
        if relacao.get("Id") == id_relacao:
            destino = relacao.get("Target", "")
            return destino.lstrip("/") if destino.startswith("/") else posixpath.normpath(
                posixpath.join("xl", destino)
            )
    raise ValueError(f"XML da aba '{nome_aba}' não encontrado no XLSX.")


def _alterar_valor_numerico_no_xml(xml: str, referencia: str, valor) -> str:
    padrao = re.compile(
        rf'(<c\b(?=[^>]*\br="{re.escape(referencia)}")[^>]*>)(.*?)(</c>)', re.DOTALL
    )
    encontrado = padrao.search(xml)
    if not encontrado:
        raise ValueError(f"Célula {referencia} não encontrada no XLSX.")
    abertura, conteudo, fechamento = encontrado.groups()
    if "<f" in conteudo:
        raise ValueError(f"A célula {referencia} contém fórmula e não será substituída.")
    abertura = re.sub(r'\s+t="(?:s|str|inlineStr)"', "", abertura)
    numero = str(int(valor)) if float(valor).is_integer() else format(float(valor), ".15g")
    if re.search(r"<v(?:\s[^>]*)?>.*?</v>", conteudo, flags=re.DOTALL):
        conteudo = re.sub(
            r"(<v(?:\s[^>]*)?>).*?(</v>)",
            rf"\g<1>{numero}\g<2>",
            conteudo,
            count=1,
            flags=re.DOTALL,
        )
    else:
        conteudo = re.sub(r"<is>.*?</is>", "", conteudo, count=1, flags=re.DOTALL)
        conteudo += f"<v>{numero}</v>"
    substituto = abertura + conteudo + fechamento
    return xml[: encontrado.start()] + substituto + xml[encontrado.end() :]


def salvar_alterando_somente_celulas(
    origem: str, destino: str, nome_aba: str, alteracoes: dict[str, float | int]
):
    os.makedirs(os.path.dirname(os.path.abspath(destino)), exist_ok=True)
    if not alteracoes:
        shutil.copy2(origem, destino)
        return
    temporario = tempfile.NamedTemporaryFile(
        prefix="ml_", suffix=".xlsx", dir=os.path.dirname(os.path.abspath(destino)), delete=False
    )
    caminho_temporario = temporario.name
    temporario.close()
    try:
        with zipfile.ZipFile(origem, "r") as arquivo_origem:
            caminho_aba = _caminho_xml_da_aba(arquivo_origem, nome_aba)
            xml_aba = arquivo_origem.read(caminho_aba).decode("utf-8")
            for referencia, valor in alteracoes.items():
                xml_aba = _alterar_valor_numerico_no_xml(xml_aba, referencia, valor)
            with zipfile.ZipFile(caminho_temporario, "w") as arquivo_destino:
                for info in arquivo_origem.infolist():
                    dados = xml_aba.encode("utf-8") if info.filename == caminho_aba else arquivo_origem.read(info.filename)
                    arquivo_destino.writestr(info, dados)
        os.replace(caminho_temporario, destino)
    except Exception:
        if os.path.exists(caminho_temporario):
            os.remove(caminho_temporario)
        raise


def atualizar_planilha_ml(caminho_xlsx: str, nome_aba: str, catalogo: dict[str, dict]):
    wb = openpyxl.load_workbook(caminho_xlsx, read_only=True, data_only=False)
    ws = wb[nome_aba]
    linha_cabecalho, colunas = _localizar_cabecalho_ml(ws)
    if not linha_cabecalho:
        wb.close()
        raise ValueError("Cabeçalho SKU/QUANTITY/PRICE não encontrado na planilha do ML.")

    alteracoes = {}
    anuncios_alterados = set()
    sem_sku_no_bling = {}
    campos_bloqueados = []
    detalhes_alteracoes = []
    relacoes = []
    linhas_analisadas = 0
    linhas_encontradas_no_bling = 0
    col_item = colunas.get("ITEM_ID")

    for linha in ws.iter_rows(min_row=linha_cabecalho + 1, max_row=ws.max_row):
        item_id = str(linha[col_item - 1].value or "").strip() if col_item else ""
        if col_item and not PADRAO_ITEM_ID.match(item_id):
            continue
        sku = limpar_sku(linha[colunas["SKU"] - 1].value)
        celula_preco = linha[colunas["PRICE"] - 1]
        celula_quantidade = linha[colunas["QUANTITY"] - 1]
        relacao = {
            "item_id": item_id or "sem ITEM_ID",
            "sku": sku or "sem SKU",
            "celula_price": celula_preco.coordinate,
            "price_ml": celula_preco.value,
            "price_bling": "",
            "status_price": "SEM SKU NA PLANILHA",
            "celula_quantity": celula_quantidade.coordinate,
            "quantity_ml": celula_quantidade.value,
            "quantity_bling": "",
            "status_quantity": "SEM SKU NA PLANILHA",
        }
        if not sku:
            relacoes.append(relacao)
            continue
        linhas_analisadas += 1
        produto = catalogo.get(sku)
        if not produto:
            sem_sku_no_bling.setdefault(sku, set()).add(item_id or "sem ITEM_ID")
            relacao["status_price"] = "SKU AUSENTE NO BLING"
            relacao["status_quantity"] = "SKU AUSENTE NO BLING"
            relacoes.append(relacao)
            continue
        linhas_encontradas_no_bling += 1
        relacao["price_bling"] = produto["preco"]
        relacao["quantity_bling"] = produto["estoque"]

        item_id = item_id or sku
        campos = (
            ("PRICE", produto["preco"], 0.001, "status_price"),
            ("QUANTITY", produto["estoque"], 0.0, "status_quantity"),
        )
        for campo, valor_novo, tolerancia, chave_status in campos:
            celula = linha[colunas[campo] - 1]
            if valores_numericos_iguais(celula.value, valor_novo, tolerancia):
                relacao[chave_status] = "SEM ALTERAÇÃO"
                continue
            if celula_editavel(celula):
                alteracoes[celula.coordinate] = valor_novo
                anuncios_alterados.add(item_id)
                relacao[chave_status] = "ALTERADO"
                detalhes_alteracoes.append(
                    (item_id, sku, campo, celula.coordinate, celula.value, valor_novo)
                )
            else:
                relacao[chave_status] = "BLOQUEADO PELO MERCADO LIVRE"
                campos_bloqueados.append(
                    (item_id, sku, campo, celula.coordinate, celula.value, valor_novo)
                )
        relacoes.append(relacao)
    wb.close()

    nome_base = os.path.splitext(os.path.basename(caminho_xlsx))[0]
    caminho_saida = os.path.join(PASTA_SAIDA, f"{nome_base}_atualizado.xlsx")
    salvar_alterando_somente_celulas(caminho_xlsx, caminho_saida, nome_aba, alteracoes)
    return {
        "anuncios_alterados": anuncios_alterados,
        "celulas_alteradas": len(alteracoes),
        "sem_sku_no_bling": sem_sku_no_bling,
        "campos_bloqueados": campos_bloqueados,
        "detalhes_alteracoes": detalhes_alteracoes,
        "relacoes": relacoes,
        "linhas_analisadas": linhas_analisadas,
        "linhas_encontradas_no_bling": linhas_encontradas_no_bling,
    }, caminho_saida


def gerar_relatorio_txt(resumo: dict, arquivo_bling: str, arquivo_ml: str, caminho_saida: str) -> str:
    """Grava um histórico legível e completo de todas as relações analisadas."""
    caminho_relatorio = os.path.join(PASTA_SAIDA, "relatorio_atualizacao.txt")
    cabecalho_colunas = (
        "ITEM_ID\tSKU\tCÉLULA_QUANTITY\tQUANTITY_ML\tESTOQUE_BLING\tRESULTADO_QUANTITY\t"
        "CÉLULA_PRICE\tPRICE_ML\tPREÇO_BLING\tRESULTADO_PRICE"
    )

    with open(caminho_relatorio, "w", encoding="utf-8-sig", newline="") as arquivo:
        arquivo.write("RELATÓRIO DE ATUALIZAÇÃO — MERCADO LIVRE\n")
        arquivo.write("=" * 80 + "\n")
        arquivo.write(f"Data e hora: {datetime.now():%d/%m/%Y %H:%M:%S}\n")
        arquivo.write(f"Catálogo Bling: {os.path.abspath(arquivo_bling)}\n")
        arquivo.write(f"Planilha Mercado Livre: {os.path.abspath(arquivo_ml)}\n")
        arquivo.write(f"Planilha gerada: {os.path.abspath(caminho_saida)}\n\n")
        arquivo.write("RESUMO\n")
        arquivo.write("-" * 80 + "\n")
        arquivo.write(f"Relações registradas: {len(resumo['relacoes'])}\n")
        arquivo.write(f"Linhas com SKU analisadas: {resumo['linhas_analisadas']}\n")
        arquivo.write(f"Linhas com SKU encontrado no Bling: {resumo['linhas_encontradas_no_bling']}\n")
        arquivo.write(f"Anúncios alterados: {len(resumo['anuncios_alterados'])}\n")
        arquivo.write(f"Células alteradas: {resumo['celulas_alteradas']}\n")
        arquivo.write(f"SKUs ausentes no Bling: {len(resumo['sem_sku_no_bling'])}\n")
        arquivo.write(f"Campos bloqueados: {len(resumo['campos_bloqueados'])}\n\n")
        arquivo.write("RELAÇÕES DETALHADAS\n")
        arquivo.write("-" * 80 + "\n")
        arquivo.write(cabecalho_colunas + "\n")

        for relacao in resumo["relacoes"]:
            valores = (
                relacao["item_id"],
                relacao["sku"],
                relacao["celula_quantity"],
                relacao["quantity_ml"],
                relacao["quantity_bling"],
                relacao["status_quantity"],
                relacao["celula_price"],
                relacao["price_ml"],
                relacao["price_bling"],
                relacao["status_price"],
            )
            arquivo.write("\t".join(str(valor if valor is not None else "") for valor in valores) + "\n")

    return caminho_relatorio


def limpar_pasta_saida():
    os.makedirs(PASTA_SAIDA, exist_ok=True)
    for nome in os.listdir(PASTA_SAIDA):
        if nome == ".gitkeep":
            continue
        caminho = os.path.join(PASTA_SAIDA, nome)
        try:
            if os.path.isfile(caminho):
                os.remove(caminho)
            elif os.path.isdir(caminho):
                shutil.rmtree(caminho)
        except OSError:
            pass


def main():
    print("=" * 60)
    print("ATUALIZADOR DE PREÇO E ESTOQUE — MERCADO LIVRE")
    print("=" * 60)
    limpar_pasta_saida()
    arquivo_bling = localizar_arquivo_bling()
    arquivo_ml, aba_ml = localizar_arquivo_anuncios_ml()
    catalogo = carregar_catalogo_bling(arquivo_bling)
    print(f"{len(catalogo)} produto(s) carregado(s) do Bling.")
    print("\nAtualizando pelo SKU da própria planilha do Mercado Livre...")

    resumo, caminho_saida = atualizar_planilha_ml(arquivo_ml, aba_ml, catalogo)
    caminho_relatorio = gerar_relatorio_txt(resumo, arquivo_bling, arquivo_ml, caminho_saida)
    print(f"Linhas de anúncios analisadas: {resumo['linhas_analisadas']}")
    print(f"Linhas com SKU encontrado no Bling: {resumo['linhas_encontradas_no_bling']}")
    print(
        f"\n{len(resumo['anuncios_alterados'])} anúncio(s) atualizado(s) "
        f"em {resumo['celulas_alteradas']} célula(s)."
    )
    print(f"Saída: {caminho_saida}")
    print(f"Relatório: {caminho_relatorio}")

    if resumo["detalhes_alteracoes"]:
        print("\nALTERAÇÕES REALIZADAS:")
        for item_id, sku, campo, celula, valor_antigo, valor_novo in resumo["detalhes_alteracoes"]:
            print(
                f"{COR_VERDE}  {item_id} | SKU {sku} | {campo} {celula}: "
                f"{valor_antigo} -> {valor_novo}{COR_RESET}"
            )

    if resumo["sem_sku_no_bling"]:
        print(
            f"\n{COR_VERMELHO}{len(resumo['sem_sku_no_bling'])} SKU(s) da planilha ML "
            f"não aparecem no CSV e foram mantidos sem alteração.{COR_RESET}"
        )
        limite = 20
        faltantes = sorted(resumo["sem_sku_no_bling"].items())
        for sku, item_ids in faltantes[:limite]:
            anuncios = ", ".join(sorted(item_ids))
            print(f"{COR_VERMELHO}  SKU {sku} | anúncio(s): {anuncios}{COR_RESET}")
        if len(faltantes) > limite:
            print(
                f"{COR_VERMELHO}  ... e mais {len(faltantes) - limite} SKU(s) "
                f"não exibido(s) para manter o log legível.{COR_RESET}"
            )
    if resumo["campos_bloqueados"]:
        print(f"\n{COR_LARANJA}CAMPOS NÃO EDITÁVEIS IGNORADOS:{COR_RESET}")
        for item_id, sku, campo, celula, valor_antigo, valor_novo in resumo["campos_bloqueados"]:
            print(
                f"{COR_LARANJA}  {item_id} | SKU {sku} | {campo} {celula} "
                f"permaneceu {valor_antigo} (Bling: {valor_novo}){COR_RESET}"
            )
    print("\n" + "=" * 60)
    print(
        f"CONCLUÍDO: {len(resumo['anuncios_alterados'])} anúncio(s) atualizado(s) | "
        f"{len(resumo['sem_sku_no_bling'])} SKU(s) não encontrado(s) no Bling | "
        f"{len(resumo['campos_bloqueados'])} campo(s) bloqueado(s)"
    )


if __name__ == "__main__":
    try:
        main()
    except Exception as erro:
        print(f"\n{COR_VERMELHO}ERRO: {erro}{COR_RESET}")
        raise
