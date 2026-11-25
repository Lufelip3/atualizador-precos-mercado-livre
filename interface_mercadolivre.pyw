import os
import re
import sys
import threading
import subprocess
import tkinter as tk
from datetime import datetime
from tkinter import ttk, messagebox

# Reaproveita toda a logica de deteccao/leitura ja escrita no motor —
# evita duplicar regras de negocio entre a interface e o script.
import script_mercadolivre as motor

PASTA_BASE = os.path.dirname(os.path.abspath(__file__))
PASTA_INPUTS = os.path.join(PASTA_BASE, "inputs")
PASTA_SAIDAS = os.path.join(PASTA_BASE, "saidas")
CAMINHO_RELATORIO = os.path.join(PASTA_SAIDAS, "relatorio_atualizacao.txt")


class AtualizadorMercadoLivreApp:
    # ── PALETA AMARELA DO MERCADO LIVRE ─────────────────────────────────
    COR_BG_RAIZ     = "#202020"
    COR_BG_HEADER   = "#FFE600"
    COR_BG_PAINEL   = "#202020"
    COR_BG_LOG      = "#101010"
    COR_TITULO      = "#242424"
    COR_BTN_ACAO    = "#FFE600"
    COR_BTN_ACAO_H  = "#E6CF00"

    COR_SUBTITULO   = "#555555"
    COR_LABEL       = "#F2F2F2"
    COR_LOG_LABEL   = "#BDBDBD"
    COR_LOG_TEXT    = "#E6E6E6"
    COR_BTN_ABRIR   = "#3A3A3A"
    COR_BTN_ABRIR_H = "#505050"

    CORES_ANSI = {
        "91": "#ff6b6b",  # vermelho — SKU da planilha ausente no Bling / erros
        "92": "#4ec94e",  # verde    — sucesso
        "93": "#f0c040",  # amarelo
        "33": "#ff8c00",  # laranja  — campo não editável no ML
        "0":  "#ccccdd",  # reset    — cor padrao do log
    }

    def __init__(self, root):
        self.root = root
        self.root.title("Atualizador de Preco e Estoque — Mercado Livre")
        self.root.geometry("700x690")
        self.root.resizable(True, True)
        self.root.minsize(680, 580)
        self.root.configure(bg=self.COR_BG_RAIZ)

        style = ttk.Style()
        if "clam" in style.theme_names():
            style.theme_use("clam")
        style.configure(
            "Dark.Horizontal.TProgressbar",
            background=self.COR_BTN_ACAO,
            troughcolor=self.COR_BG_LOG,
            bordercolor=self.COR_BG_RAIZ,
            lightcolor=self.COR_BTN_ACAO,
            darkcolor=self.COR_BTN_ACAO,
        )

        self.modo = tk.StringVar(value=motor.MODOS["estoque"])
        self._criar_widgets()
        self.atualizar_painel_de_status()

    # ── CONSTRUCAO DA JANELA ───────────────────────────────────────────
    def _criar_widgets(self):
        header = tk.Frame(self.root, bg=self.COR_BG_HEADER, pady=12)
        header.pack(fill=tk.X)

        tk.Label(
            header,
            text="Atualizador Mercado Livre  🛒",
            font=("Helvetica", 17, "bold"),
            bg=self.COR_BG_HEADER,
            fg=self.COR_TITULO,
        ).pack()

        tk.Label(
            header,
            text="Bling (CSV)  →  SKU  →  Anúncios Mercado Livre (XLSX)",
            font=("Helvetica", 9),
            bg=self.COR_BG_HEADER,
            fg=self.COR_SUBTITULO,
        ).pack()

        frame_modo = tk.Frame(self.root, bg=self.COR_BG_RAIZ)
        frame_modo.pack(fill=tk.X, padx=24, pady=(10, 0))
        tk.Label(frame_modo, text="O que atualizar?", bg=self.COR_BG_RAIZ,
                 fg=self.COR_LABEL).pack(side=tk.LEFT, padx=(0, 12))
        self.seletor_modo = ttk.Combobox(frame_modo, textvariable=self.modo,
                                        values=list(motor.MODOS.values()), state="readonly", width=22)
        self.seletor_modo.pack(side=tk.LEFT)
        self.seletor_modo.bind("<<ComboboxSelected>>", lambda evento: self.atualizar_painel_de_status())

        # ── PAINEL DE STATUS ───────────
        self.frame_status = tk.Frame(self.root, bg=self.COR_BG_PAINEL, pady=10)
        self.frame_status.pack(fill=tk.X, padx=24)

        ttk.Separator(self.root, orient="horizontal").pack(fill=tk.X, padx=24, pady=4)

        # ── BOTOES ───────────
        frame_btn = tk.Frame(self.root, bg=self.COR_BG_RAIZ, pady=12)
        frame_btn.pack(fill=tk.X, padx=24)

        self.btn_executar = tk.Button(
            frame_btn,
            text="🚀  Atualizar Mercado Livre",
            bg=self.COR_BTN_ACAO,
            fg="#202020",
            font=("Arial", 12, "bold"),
            relief="flat",
            cursor="hand2",
            padx=14, pady=8,
            activebackground=self.COR_BTN_ACAO_H,
            activeforeground="#202020",
            command=self.executar_atualizacao,
        )
        self.btn_executar.pack(fill=tk.X, pady=(0, 8))

        self.btn_abrir = tk.Button(
            frame_btn,
            text="📂  Abrir pasta saidas/",
            bg=self.COR_BTN_ABRIR,
            fg="white",
            font=("Arial", 10),
            relief="flat",
            cursor="hand2",
            padx=14, pady=5,
            activebackground=self.COR_BTN_ABRIR_H,
            command=self.abrir_pasta_saidas,
        )
        self.btn_abrir.pack(fill=tk.X, pady=(0, 8))

        self.btn_relatorio = tk.Button(
            frame_btn,
            text="📄  Abrir relatório TXT",
            bg=self.COR_BTN_ABRIR,
            fg=self.COR_BTN_ACAO,
            font=("Arial", 10, "bold"),
            relief="flat",
            cursor="hand2",
            padx=14, pady=5,
            activebackground=self.COR_BTN_ABRIR_H,
            activeforeground=self.COR_BTN_ACAO,
            command=self.abrir_relatorio,
            state="normal" if os.path.exists(CAMINHO_RELATORIO) else "disabled",
        )
        self.btn_relatorio.pack(fill=tk.X)

        # ── BARRA DE PROGRESSO ───────────
        self.barra_progresso = ttk.Progressbar(
            self.root, mode="indeterminate", style="Dark.Horizontal.TProgressbar"
        )
        self.barra_progresso.pack(fill=tk.X, padx=24, pady=(0, 8))

        # ── LEGENDA DE CORES DO RELATORIO ───────────
        frame_legenda = tk.Frame(self.root, bg=self.COR_BG_RAIZ)
        frame_legenda.pack(fill=tk.X, padx=24, pady=(0, 6))

        legenda_itens = [
            ("SKU do ML ausente no Bling", self.CORES_ANSI["91"]),
            ("Campo bloqueado pelo ML", self.CORES_ANSI["33"]),
        ]
        for texto, cor in legenda_itens:
            tk.Label(
                frame_legenda,
                text=f"● {texto}",
                font=("Segoe UI", 8),
                bg=self.COR_BG_RAIZ,
                fg=cor,
            ).pack(side=tk.LEFT, padx=(0, 12))

        # ── AREA DE LOG ───────────
        frame_log = tk.Frame(self.root, bg=self.COR_BG_RAIZ)
        frame_log.pack(fill=tk.BOTH, expand=True, padx=24, pady=(0, 16))

        tk.Label(
            frame_log,
            text="Registro detalhado da execução:",
            font=("Segoe UI", 9, "bold"),
            bg=self.COR_BG_RAIZ,
            fg=self.COR_LOG_LABEL,
        ).pack(anchor="w")

        container = tk.Frame(frame_log, bg=self.COR_BG_LOG)
        container.pack(fill=tk.BOTH, expand=True, pady=(4, 0))

        self.area_log = tk.Text(
            container,
            font=("Consolas", 10),
            bg=self.COR_BG_LOG,
            fg=self.COR_LOG_TEXT,
            insertbackground="white",
            relief="flat",
            bd=0,
            state="disabled",
            padx=8, pady=6,
        )
        sb = ttk.Scrollbar(container, command=self.area_log.yview)
        self.area_log.configure(yscrollcommand=sb.set)

        self.area_log.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        sb.pack(side=tk.RIGHT, fill=tk.Y)

    # ── DETECCAO DE ARQUIVOS (para o painel de status) ──────────────────
    def modo_selecionado(self):
        return next(chave for chave, titulo in motor.MODOS.items() if titulo == self.modo.get())

    def detectar_arquivos(self):
        """Detecta as duas entradas necessárias usando o motor do programa."""
        arquivo_bling = None
        arquivo_anuncios, aba_anuncios = None, None

        try:
            arquivo_anuncios, aba_anuncios = motor.localizar_arquivo_anuncios_ml(exibir_log=False, modo=self.modo_selecionado())
        except Exception:
            pass

        try:
            arquivo_bling = motor.localizar_arquivo_bling(exibir_log=False, modo=self.modo_selecionado())
        except Exception:
            pass

        return arquivo_bling, (arquivo_anuncios, aba_anuncios)

    def atualizar_painel_de_status(self):
        for widget in self.frame_status.winfo_children():
            widget.destroy()

        tk.Label(
            self.frame_status,
            text="Arquivos detectados na pasta inputs/:",
            font=("Segoe UI", 10, "bold"),
            bg=self.COR_BG_PAINEL,
            fg=self.COR_LABEL,
        ).pack(anchor="w", pady=(0, 4))

        arquivo_bling, (arquivo_anuncios, aba_anuncios) = self.detectar_arquivos()

        def linha(texto, cor):
            tk.Label(
                self.frame_status,
                text=texto,
                font=("Segoe UI", 10),
                bg=self.COR_BG_PAINEL,
                fg=cor,
            ).pack(anchor="w")

        if arquivo_bling:
            linha(f"  OK   Bling (catalogo): {os.path.basename(arquivo_bling)}", "#4ec94e")
        else:
            linha(f"  FALTA   CSV do Bling compatível com {self.modo.get()}", "#ff6b6b")

        if arquivo_anuncios:
            linha(f"  OK   Anuncios Mercado Livre: {os.path.basename(arquivo_anuncios)}  [aba: {aba_anuncios}]", "#4ec94e")
        else:
            linha(f"  SEM PLANILHA COMPATÍVEL: exige {motor.colunas_exigidas(self.modo_selecionado())}", "#ff6b6b")

    # ── PASTA DE SAIDA ───────────────────────────────────────────────────
    def abrir_pasta_saidas(self):
        if os.path.exists(PASTA_SAIDAS):
            os.startfile(PASTA_SAIDAS)
        else:
            messagebox.showinfo("Aviso", "A pasta 'saidas' ainda nao foi criada.")

    def abrir_relatorio(self):
        if os.path.exists(CAMINHO_RELATORIO):
            os.startfile(CAMINHO_RELATORIO)
        else:
            messagebox.showinfo(
                "Relatório ainda não gerado",
                "Execute uma atualização para gerar o relatório TXT.",
            )

    # ── LOG COLORIDO (interpreta os codigos ANSI que o script imprime) ───
    def escrever_no_log(self, texto: str, cor_padrao: str = "#d4d4d4"):
        self.area_log.configure(state="normal")
        segmentos = re.split(r"(\033\[[0-9;]+m)", texto)
        cor_atual = cor_padrao
        contador_tags = 0

        for segmento in segmentos:
            match_ansi = re.match(r"\033\[([0-9;]+)m", segmento)
            if match_ansi:
                codigo = match_ansi.group(1)
                cor_atual = self.CORES_ANSI.get(codigo, cor_padrao)
            elif segmento:
                nome_tag = f"cor_{id(segmento)}_{contador_tags}"
                contador_tags += 1
                self.area_log.tag_configure(nome_tag, foreground=cor_atual)
                self.area_log.insert("end", segmento, nome_tag)

        self.area_log.insert("end", "\n")
        self.area_log.configure(state="disabled")
        self.area_log.see("end")

    # ── EXECUCAO ───────────────────────────────────────────────────────
    def executar_atualizacao(self):
        modo = self.modo_selecionado()
        self.atualizar_painel_de_status()
        arquivo_bling, (arquivo_anuncios, _) = self.detectar_arquivos()

        faltando = []
        if not arquivo_bling:
            faltando.append("catalogo do Bling (CSV)")
        if not arquivo_anuncios:
            faltando.append(f"planilha do Mercado Livre com {motor.colunas_exigidas(modo)} (XLSX)")

        if faltando:
            messagebox.showerror(
                "Arquivos ausentes ou incompatíveis",
                f"Não foi possível iniciar no modo {motor.MODOS[modo]}.\nConfira em inputs/:\n\n- " + "\n- ".join(faltando),
            )
            return

        self.btn_executar.config(state="disabled")
        self.seletor_modo.config(state="disabled")
        self.btn_abrir.config(state="disabled")
        self.btn_relatorio.config(state="disabled")

        self.area_log.configure(state="normal")
        self.area_log.delete("1.0", "end")
        self.area_log.configure(state="disabled")

        self.escrever_no_log(f"Pasta de trabalho: {PASTA_INPUTS}")
        self.escrever_no_log(f"Início: {datetime.now():%d/%m/%Y %H:%M:%S}")
        self.escrever_no_log("Aguardando o processamento...\n")
        self.barra_progresso.start(10)

        def rodar_em_segundo_plano():
            try:
                caminho_script = os.path.join(PASTA_BASE, "script_mercadolivre.py")
                processo = subprocess.Popen(
                    [sys.executable, "-u", caminho_script, "--modo", modo],
                    stdout=subprocess.PIPE,
                    stderr=subprocess.STDOUT,
                    text=True,
                    encoding="utf-8",
                    errors="replace",
                    bufsize=1,
                    cwd=PASTA_BASE,
                    creationflags=subprocess.CREATE_NO_WINDOW if os.name == "nt" else 0,
                )
                for linha in processo.stdout:
                    texto = linha.rstrip("\r\n")
                    self.root.after(0, lambda conteudo=texto: self.escrever_no_log(conteudo))
                codigo_retorno = processo.wait()
                self.root.after(0, lambda: self.finalizar_execucao(codigo_retorno))
            except Exception as erro:
                mensagem = str(erro)
                self.root.after(0, lambda conteudo=mensagem: self.finalizar_com_erro(conteudo))

        threading.Thread(target=rodar_em_segundo_plano, daemon=True).start()

    def finalizar_execucao(self, codigo_retorno):
        self.seletor_modo.config(state="readonly")
        self.barra_progresso.stop()
        self.btn_executar.config(state="normal")
        self.btn_abrir.config(state="normal")
        self.btn_relatorio.config(
            state="normal" if os.path.exists(CAMINHO_RELATORIO) else "disabled"
        )
        self.atualizar_painel_de_status()

        if codigo_retorno == 0:
            self.escrever_no_log("\n" + "-" * 37, "#555")
            self.escrever_no_log("Arquivo salvo em:  saidas/", self.CORES_ANSI["92"])
            self.escrever_no_log("  - Anuncios-MercadoLivre_atualizado.xlsx", self.CORES_ANSI["92"])
            self.escrever_no_log("  - relatorio_atualizacao.txt", self.CORES_ANSI["92"])
            messagebox.showinfo(
                "Concluído",
                "Processo finalizado!\n\nConfira o log, a planilha atualizada e o relatório TXT.",
            )
        else:
            self.escrever_no_log(
                f"\nO processo terminou com o código de erro {codigo_retorno}.",
                self.CORES_ANSI["91"],
            )
            messagebox.showerror("Erro", "Ocorreu um erro.\nVeja o log para detalhes.")

    def finalizar_com_erro(self, mensagem_erro: str):
        self.seletor_modo.config(state="readonly")
        self.barra_progresso.stop()
        self.btn_executar.config(state="normal")
        self.btn_abrir.config(state="normal")
        self.btn_relatorio.config(
            state="normal" if os.path.exists(CAMINHO_RELATORIO) else "disabled"
        )
        self.escrever_no_log(f"Erro inesperado: {mensagem_erro}", self.CORES_ANSI["91"])
        messagebox.showerror("Erro", f"Erro inesperado:\n{mensagem_erro}")


if __name__ == "__main__":
    root = tk.Tk()
    app = AtualizadorMercadoLivreApp(root)
    root.mainloop()
