# 🏷️ Atualizador de Preços e Estoque do Mercado Livre — Sincronização direta por SKU

Programa para atualizar preço e estoque dos anúncios do Mercado Livre usando o
catálogo exportado do Bling. O cruzamento é feito diretamente entre o `SKU` da
planilha do Mercado Livre e o `Código` do CSV do Bling.

> **Versão atual** — atualiza somente preço e estoque e preserva toda a estrutura da planilha oficial.

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

1. O programa localiza, na pasta `inputs/`, um CSV do Bling que contenha as
   colunas `Código`, `Preço` e `Estoque`.
2. Localiza a planilha oficial do Mercado Livre que contenha `SKU`, `PRICE` e
   `QUANTITY`.
3. Cruza o `SKU` da planilha com o `Código` do Bling.
4. Atualiza somente as células editáveis de `PRICE` e `QUANTITY` quando o valor
   estiver diferente.
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
│   └── produtos_2026-09-02-15-03-08.csv
├── saidas/
│   ├── Anuncios-MercadoLivre_atualizado.xlsx
│   └── relatorio_atualizacao.txt
├── interface_mercadolivre.pyw
├── script_mercadolivre.py
└── README.md
```

| Arquivo ou pasta | Descrição |
|---|---|
| `inputs/` | Recebe o CSV do Bling e o XLSX baixado do Mercado Livre. |
| `saidas/` | Recebe a planilha atualizada que será enviada ao Mercado Livre. |
| `saidas/relatorio_atualizacao.txt` | Registra todas as relações entre anúncios, SKUs e dados do Bling. |
| `script_mercadolivre.py` | Faz a identificação dos arquivos, o cruzamento por SKU e a atualização. |
| `interface_mercadolivre.pyw` | Interface gráfica usada para executar e acompanhar o processo. |
| `erros.xlsx` | Relatório que o Mercado Livre pode gerar depois do envio. Não é uma entrada do programa. |

---

## 📋 Requisitos

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

Na tela `Personalize a planilha que você vai baixar`, use esta configuração:

1. Em `Dados para alterar`, selecione `Condições gerais`.
2. Em `Condições de venda`, marque somente `Preços`.
3. Em `Dados do produto`, marque `SKU`.
4. Em `Estoque`, marque `No seu depósito`.
5. Não selecione `Características do produto`.

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

1. Exporte do Bling um CSV com `Código`, `Preço` e `Estoque`.
2. Baixe do Mercado Livre o XLSX seguindo a configuração descrita acima.
3. Coloque os dois arquivos em `inputs/`.
4. Feche no Excel qualquer arquivo existente na pasta `saidas/`.

### 2. Executar a atualização

1. Dê duplo clique em `interface_mercadolivre.pyw`.
2. Confirme no painel que o CSV do Bling e o XLSX do Mercado Livre foram
   identificados.
3. Clique no botão de execução da interface.
4. Aguarde a mensagem de conclusão.
5. Abra `saidas/Anuncios-MercadoLivre_atualizado.xlsx` somente depois que o
   processamento terminar.
6. Use `Abrir relatório TXT` para consultar todas as comparações e decisões do
   processamento.

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
| Atualização limitada | Altera somente `PRICE` e `QUANTITY`. |
| Respeito às permissões | Ignora células cinzas, bloqueadas ou não editáveis. |
| Preservação do XLSX | Mantém estrutura, formatação, abas ocultas e demais campos. |
| Interface Mercado Livre | Usa o amarelo característico do Mercado Livre como cor principal. |
| Log detalhado em tempo real | Mostra arquivos, comparações, células alteradas, SKUs ausentes, campos bloqueados e erros durante o processamento. |
| Relatório TXT automático | Gera um histórico completo com `ITEM_ID`, SKU, células, valores do Mercado Livre, valores do Bling e resultado de cada comparação. |

---

## 🩹 Problemas comuns

| Situação | Solução |
|---|---|
| O programa não encontra o CSV do Bling | Confirme se o arquivo está em `inputs/` e possui `Código`, `Preço` e `Estoque`. |
| O programa não encontra o XLSX | Baixe novamente a planilha oficial e confirme se ela possui `SKU`, `PRICE` e `QUANTITY`. |
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
- Somente `PRICE` e `QUANTITY` podem ser modificados.
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
