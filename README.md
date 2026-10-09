# 📦 Relatório de Estoque e Pedidos

![Testes](https://github.com/Eliel-cyber/relatorio-estoque/actions/workflows/testes.yml/badge.svg)
![Python](https://img.shields.io/badge/Python-3.10%2B-3776AB?style=flat-square&logo=python&logoColor=white)
![pandas](https://img.shields.io/badge/pandas-150458?style=flat-square&logo=pandas&logoColor=white)
![Excel](https://img.shields.io/badge/Sa%C3%ADda-Excel-217346?style=flat-square&logo=microsoftexcel&logoColor=white)

Script em Python que lê as planilhas de **estoque** e **pedidos** exportadas do ERP e gera, em segundos, um **relatório em Excel** com os pontos que precisam de atenção no dia a dia de um e-commerce.

## 🎯 Por que criei este projeto

No meu trabalho como Assistente de E-commerce, faço o controle de estoque e o faturamento de pedidos em ERP (Bling e TOTVS) e monto os controles em Excel com PROCV e tabela dinâmica. Algumas conferências se repetem toda semana:

- o saldo do sistema bate com a contagem física?
- quais produtos estão abaixo do estoque mínimo e quanto preciso repor?
- tem pedido pendente que vai atrasar por falta de produto?
- quanto faturamos em cada canal de venda?

Este projeto automatiza essas conferências: em vez de cruzar planilhas manualmente, o script faz o cruzamento e entrega o resultado pronto em abas separadas.

> Os dados da pasta `dados/` são **fictícios**, criados apenas para demonstração.

## 📊 O que o relatório mostra

| Aba | O que tem | Para que serve |
|---|---|---|
| **Divergencias** | Produtos em que o saldo do sistema ≠ contagem física, com o impacto em R$ | Corrigir o estoque no ERP e investigar perdas |
| **Abaixo do minimo** | Produtos abaixo do estoque mínimo, quantidade a repor e custo | Planejar a compra |
| **Pedidos em risco** | Pedidos pendentes sem estoque suficiente | Avisar o cliente antes de atrasar |
| **Faturamento** | Pedidos, itens, faturamento e ticket médio por canal | Acompanhar vendas |

### Exemplo de saída no terminal

```
Relatório gerado: relatorio_estoque.xlsx
  Divergências de estoque: 4 produto(s)
  Abaixo do mínimo:        3 produto(s)
  Pedidos em risco:        2 pedido(s)
  Faturamento total:       R$ 1.862,70
```

### Exemplo da aba "Pedidos em risco"

| pedido | data | sku | produto | quantidade | contagem_fisica | faltam | canal |
|---|---|---|---|---|---|---|---|
| 1007 | 04/09/2026 | VIS-060 | Viseira Fume Universal | 2 | 1 | 1 | Mercado Livre |
| 1009 | 05/09/2026 | LUV-011 | Luva de Couro Tam. M | 5 | 4 | 1 | Mercado Livre |

## ▶️ Como rodar

Pré-requisito: Python 3.10 ou superior.

```bash
git clone https://github.com/Eliel-cyber/relatorio-estoque.git
cd relatorio-estoque
pip install -r requirements.txt
python src/relatorio.py
```

O arquivo `relatorio_estoque.xlsx` será criado na pasta. Para usar outras planilhas:

```bash
python src/relatorio.py --estoque minha_pasta/estoque.csv --pedidos minha_pasta/pedidos.csv --saida relatorio_setembro.xlsx
```

### Formato esperado dos arquivos

**estoque.csv:** `sku, produto, saldo_sistema, contagem_fisica, estoque_minimo, custo_unitario`

**pedidos.csv:** `pedido, data, sku, quantidade, valor_unitario, canal, status`
(status: `Faturado`, `Pendente`, `Cancelado` ou `Devolvido`)

Se faltar alguma coluna ou houver SKU repetido no estoque, o script avisa qual é o problema em vez de gerar um relatório errado.

## 🗂️ Estrutura

```
relatorio-estoque/
├── dados/
│   ├── estoque.csv          # exemplo de estoque (fictício)
│   └── pedidos.csv          # exemplo de pedidos (fictício)
├── src/
│   └── relatorio.py         # leitura, análises e geração do Excel
├── tests/
│   └── test_relatorio.py    # testes automáticos de cada análise
└── requirements.txt
```

## ✅ Testes

```bash
python -m unittest discover -s tests -v
```

Os testes também rodam automaticamente no GitHub Actions a cada atualização.

## 🧠 O que aprendi

- Ler e cruzar tabelas com **pandas** (`merge`, equivalente ao PROCV do Excel, e `groupby`, equivalente à tabela dinâmica)
- Gerar planilhas Excel com várias abas usando **openpyxl**
- Validar os dados de entrada e mostrar mensagens de erro claras
- Escrever testes automáticos e configurar integração contínua

---

Feito por **Eliel Dias** · [LinkedIn](https://www.linkedin.com/in/elieldias) · [Portfólio](https://eliel-cyber.github.io)
