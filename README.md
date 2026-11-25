# 🏷️ Atualizador de Preços e Estoque do Mercado Livre — Sincronização direta por SKU

Programa para atualizar preço e estoque dos anúncios do Mercado Livre usando o
catálogo exportado do Bling. O cruzamento é feito diretamente entre o `SKU` da
planilha do Mercado Livre e o `Código` do CSV do Bling.

> **Versão atual** — escolha entre Só estoque, Só preço ou Preço e estoque. O relatório registra os valores anteriores e novos de cada alteração aplicada.

---

## 📋 Índice

- [Como funciona](#-como-funciona)
- [Estrutura do projeto](#-estrutura-do-projeto)
- [Requisitos](#-requisitos)
- [Como baixar a planilha do Mercado Livre](#-como-baixar-a-planilha-do-mercado-livre)
- [Como usar](#-como-usar)
- [Novidades da versão atual](#-novidades-da-versão-atual)
- [Problemas comuns](#-problemas-comuns)
- [Notas de segurança](#-notas-de-segurança)

---

## 🧠 Como funciona

1. Escolha o modo na interface: **Só estoque** (inicial), **Só preço** ou **Preço e estoque**.
   O CSV do Bling precisa de `Código` e apenas dos campos selecionados (`Estoque` e/ou `Preço`).
2. Localiza em `inputs/` uma planilha oficial compatível:
   estoque exige `SKU` e `QUANTITY`; preço exige `SKU` e `PRICE`; ambos exige as três colunas.
3. Cruza o `SKU` da planilha com o `Código` do Bling.
4. Atualiza somente as células editáveis dos campos selecionados quando o valor estiver diferente.
   Os campos não selecionados permanecem intactos e aparecem assim identificados no relatório.
5. Mantém todas as outras colunas, abas, formatações e informações originais.
6. Salva o resultado em `saidas/Anuncios-MercadoLivre_atualizado.xlsx`.

O programa não utiliza planilhas de vínculo. O `ITEM_ID` continua na planilha,
mas a correspondência com o Bling é feita exclusivamente pelo SKU.

---

## 🗂️ Estrutura do projeto

```text
Automação de preços para Mercado Livre/
├── inputs/
│   ├── Anuncios-MercadoLivre.xlsx
│   └── produtos_sincronizados.csv
├── saidas/
│   ├── Anuncios-MercadoLivre_atualizado.xlsx
│   └── relatorio_atualizacao.txt
├── interface_mercadolivre.pyw
├── script_mercadolivre.py
├── tests/
│   └── test_modos.py
└── README.md
```

| Arquivo ou pasta | Descrição |
|---|---|
| `inputs/` | Recebe o CSV do Bling e o XLSX baixado do Mercado Livre. |
| `saidas/` | Pasta de saída do programa: recebe a planilha atualizada e o relatório TXT. |
| `saidas/relatorio_atualizacao.txt` | Registra todas as relações entre anúncios, SKUs e dados do Bling. |
| `script_mercadolivre.py` | Faz a identificação dos arquivos, o cruzamento por SKU e a atualização. |
| `interface_mercadolivre.pyw` | Interface gráfica usada para executar e acompanhar o processo. |
| `erros.xlsx` | Relatório que o Mercado Livre pode gerar depois do envio. Não é uma entrada do programa. |

O fluxo do atualizador é **`inputs/` → processamento → `saidas/`**.
O nome do XLSX gerado usa o nome do arquivo de entrada seguido de `_atualizado.xlsx`.

`tests/` contém verificações locais de desenvolvimento e é ignorada pelo Git.
O programa não importa nem executa esses arquivos: funciona normalmente sem
essa pasta. Arquivos auxiliares e dependências locais (`outputs/` e
`node_modules/`) também são ignorados, assim como os resultados em `saidas/`.

A cada execução, após validar as entradas, o programa limpa `saidas/` e gera
o resultado e o relatório da execução atual. Para guardar um histórico, copie
os resultados desejados para outra pasta antes de executar novamente.

---

## 📋 Requisitos

### Organização e GitHub

Versione o código (`interface_mercadolivre.pyw` e `script_mercadolivre.py`),
o `README.md`, `.gitignore`, `.gitattributes` e
`inputs/.gitkeep`. Não envie os CSVs, planilhas, relatórios ou arquivos pessoais.
Mantenha entradas em `inputs/`, resultados em `saidas/` e testes locais em `tests/`.

Se a pasta local `tests/` estiver disponível, você pode executar as verificações
manualmente, sem alterar os arquivos de entrada ou saída:

```bash
python -B -m unittest discover -s tests -v
```

Os testes usam arquivos temporários e verificam os modos, as células bloqueadas
e o registro das alterações no relatório.

- Windows com Python 3.10 ou superior.
- Biblioteca `openpyxl`.
- `tkinter`, normalmente incluído na instalação do Python para Windows.

Instale a dependência necessária:

```bash
pip install openpyxl
```

---

## ⬇️ Como baixar a planilha do Mercado Livre

### 1. Abrir a edição por planilha

1. Entre na Central de vendedores do Mercado Livre.
2. Acesse a área `Anúncios`.
3. Abra `Alteração na planilha Excel`.
4. Clique em `Baixe a planilha Excel`.

Também é possível abrir diretamente a página de
[edição por planilha Excel](https://vendedores.mercadolivre.com.br/anuncios/edicao-em-excel?downloadPurpose=edition)
depois de entrar na conta do Mercado Livre.

Essa mesma tela possui a opção `Envie a planilha que você baixou`, usada no
final do processo.

### 2. Selecionar os dados da planilha

Escolha os campos conforme o modo que será usado no programa:

| Modo | Campos necessários no XLSX | Colunas necessárias no CSV do Bling |
|---|---|---|
| Só estoque (inicial na interface) | `SKU` e `QUANTITY` (estoque no depósito) | `Código` e `Estoque` |
| Só preço | `SKU` e `PRICE` | `Código` e `Preço` |
| Preço e estoque | `SKU`, `PRICE` e `QUANTITY` | `Código`, `Preço` e `Estoque` |

Para **Só estoque**, inclua o SKU e o estoque no seu depósito; não é necessário
incluir preços. Para **Só preço**, inclua SKU e preços. Para atualizar ambos,
inclua os três campos. Não é necessário selecionar características do produto.

Os demais campos não são necessários para este programa. Selecionar apenas o
que será alterado deixa a planilha menor e reduz o risco de enviar dados que não
fazem parte da atualização.

### 3. Escolher quais anúncios baixar

Na seção `Anúncios que você quer incluir na planilha`:

1. Escolha a categoria ou as categorias que deseja atualizar.
2. A categoria `Perfumes` é apenas um exemplo e não é obrigatória.
3. Se quiser atualizar outras categorias, marque-as no lugar de `Perfumes`.
4. Use os filtros de status, forma de entrega, custo de envio, preço ou
   quantidade somente quando quiser limitar os anúncios baixados.
5. Sem filtros adicionais, o Mercado Livre inclui os anúncios pertencentes às
   categorias selecionadas conforme a disponibilidade para edição.

O Mercado Livre informa que variações de produtos e filtros podem fazer a
quantidade de linhas baixadas ser diferente da quantidade exibida na seleção.

### 4. Baixar o arquivo

1. Revise as opções selecionadas.
2. Clique em `Baixar`.
3. Mova o XLSX baixado para a pasta `inputs/` do programa.
4. Mantenha apenas a versão mais recente da planilha em `inputs/`.

Sempre baixe uma planilha nova antes de iniciar outra atualização. O Mercado
Livre permite enviar cada planilha somente uma vez e recomenda trabalhar com a
versão mais recente. A planilha pode conter no máximo 10.000 anúncios.

---

## 🚀 Como usar

### 1. Preparar os arquivos

1. Exporte do Bling um CSV atualizado com `Código` e os campos exigidos pelo modo escolhido.
2. Baixe do Mercado Livre o XLSX seguindo a configuração descrita acima.
3. Coloque os dois arquivos em `inputs/`.
4. Feche no Excel qualquer arquivo existente na pasta `saidas/`.

### 2. Executar a atualização

1. Dê duplo clique em `interface_mercadolivre.pyw`.
2. Escolha **Só estoque**, **Só preço** ou **Preço e estoque** e confirme no painel
   que os dois arquivos foram identificados. A interface inicia em **Só estoque**.
3. Clique no botão de execução da interface.
4. Aguarde a mensagem de conclusão.
5. Abra `saidas/Anuncios-MercadoLivre_atualizado.xlsx` somente depois que o
   processamento terminar.
6. Use `Abrir relatório TXT` para consultar todas as comparações e decisões do
   processamento.

O relatório começa com uma seção **ALTERAÇÕES REALIZADAS NA PLANILHA GERADA**,
com SKU, anúncio, campo, valor anterior, valor novo e célula. Por exemplo:
`SKU 001 | Anúncio MLB1 | Estoque: 10 -> 0 | Célula C2`.
Somente alterações aplicadas entram nessa seção. Valores iguais, campos bloqueados
e produtos sem correspondência continuam discriminados nas relações detalhadas.
O relatório em `saidas/relatorio_atualizacao.txt` corresponde à última execução.

Para conferir uma alteração, localize o SKU no CSV e o anúncio/célula no XLSX
original e no gerado. O valor novo deve corresponder ao CSV. Estoque `0,00`
no CSV é um zero válido e pode zerar o estoque do anúncio na planilha gerada.
SKU ausente no CSV é mantido sem alteração; ausência não é tratada como zero.
O programa usa os arquivos locais e não consulta o estoque atual do Bling pela internet.

Pelo terminal, use `python script_mercadolivre.py --modo estoque`,
`--modo preco` ou `--modo ambos`. Sem argumento, o terminal usa `ambos`.

### 3. Enviar a planilha atualizada

1. Volte à tela `Alteração na planilha Excel` no Mercado Livre.
2. Clique em `Envie a planilha que você baixou`.
3. Selecione `saidas/Anuncios-MercadoLivre_atualizado.xlsx`.
4. Aguarde o processamento e confira o resultado informado pelo Mercado Livre.
5. Se for gerado um `erros.xlsx`, leia a coluna `ERRORS` para identificar a causa
   de cada recusa.

---

## 🆕 Novidades da versão atual

| Funcionalidade | Descrição |
|---|---|
| Cruzamento direto por SKU | Usa `SKU` no XLSX e `Código` no CSV, sem planilha de vínculo. |
| Seletor de modo | Só estoque, Só preço ou Preço e estoque, com validação das colunas necessárias. |
| Atualização limitada | Altera apenas os campos do modo escolhido: preço, estoque ou ambos. |
| Respeito às permissões | Ignora células cinzas, bloqueadas ou não editáveis. |
| Preservação do XLSX | Mantém estrutura, formatação, abas ocultas e demais campos. |
| Interface Mercado Livre | Usa o amarelo característico do Mercado Livre como cor principal. |
| Log detalhado em tempo real | Mostra arquivos, comparações, células alteradas, SKUs ausentes, campos bloqueados e erros durante o processamento. |
| Relatório TXT automático | Lista as alterações aplicadas com valor anterior e novo, além das comparações da última execução. |

---

## 🩹 Problemas comuns

| Situação | Solução |
|---|---|
| O programa não encontra o CSV do Bling | Confirme se está em `inputs/` e possui `Código` e os campos selecionados no modo. |
| O programa não encontra o XLSX | Confira as colunas exigidas pelo modo. Para uma planilha com `SKU` e `QUANTITY`, escolha **Só estoque**. |
| Aparece `PermissionError` ou acesso negado | Feche a planilha de saída no Excel e execute novamente. |
| Um SKU do Mercado Livre não aparece no Bling | O anúncio é mantido sem alteração. Confira o cadastro ou a exportação do Bling. |
| O log mostra campo não editável | O programa respeitou a proteção do Mercado Livre e não alterou a célula. |
| O Mercado Livre gera `erros.xlsx` para um anúncio paralisado | A recusa é do Mercado Livre. Anúncios paralisados permanentemente não aceitam atualizações pela planilha. |
| O Mercado Livre informa que o anúncio está em revisão | Aguarde ou verifique o anúncio na Central de vendedores. Esse bloqueio ocorre no envio, não na geração do XLSX. |
| A quantidade baixada difere da quantidade selecionada | Variações e filtros aplicados podem alterar a quantidade final de linhas. |

---

## 🔒 Notas de segurança

- O programa nunca altera título, descrição, categoria, condição, moeda,
  tipo de anúncio ou identificadores.
- Somente os campos selecionados (`PRICE` e/ou `QUANTITY`) podem ser modificados.
- Células bloqueadas ou fora do padrão editável do Mercado Livre são
  ignoradas.
- A ordem das colunas e das linhas é preservada.
- A planilha original em `inputs/` não é sobrescrita.
- Não renomeie colunas, não exclua linhas e não reorganize a planilha antes
  de enviá-la ao Mercado Livre.
- Confira o log e mantenha uma cópia do arquivo original até confirmar o
  processamento no Mercado Livre.

---

*Desenvolvido para atualizar preços e estoques de anúncios com segurança e compatibilidade com o Mercado Livre.*
