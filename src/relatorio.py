"""
Relatório de Estoque e Pedidos
==============================

Lê duas planilhas exportadas do ERP (estoque e pedidos, em CSV) e gera um
relatório em Excel com:

1. Divergências entre o saldo do sistema e a contagem física
2. Produtos abaixo do estoque mínimo (com quantidade para repor)
3. Pedidos pendentes sem estoque suficiente (risco de atraso na entrega)
4. Resumo do faturamento por canal de venda

Uso:
    python src/relatorio.py
    python src/relatorio.py --estoque dados/estoque.csv --pedidos dados/pedidos.csv --saida relatorio.xlsx
"""

import argparse
import sys
from pathlib import Path

import pandas as pd

COLUNAS_ESTOQUE = ["sku", "produto", "saldo_sistema", "contagem_fisica", "estoque_minimo", "custo_unitario"]
COLUNAS_PEDIDOS = ["pedido", "data", "sku", "quantidade", "valor_unitario", "canal", "status"]


# ---------------------------------------------------------------------------
# Leitura dos arquivos
# ---------------------------------------------------------------------------

def _validar_colunas(tabela: pd.DataFrame, obrigatorias: list, nome_arquivo: str) -> None:
    """Confere se o arquivo tem todas as colunas esperadas e avisa quais faltam."""
    faltando = [coluna for coluna in obrigatorias if coluna not in tabela.columns]
    if faltando:
        raise ValueError(f"O arquivo '{nome_arquivo}' está sem as colunas: {', '.join(faltando)}")


def carregar_estoque(caminho) -> pd.DataFrame:
    estoque = pd.read_csv(caminho)
    _validar_colunas(estoque, COLUNAS_ESTOQUE, str(caminho))
    if estoque["sku"].duplicated().any():
        repetidos = estoque.loc[estoque["sku"].duplicated(), "sku"].tolist()
        raise ValueError(f"SKUs repetidos no estoque: {', '.join(repetidos)}")
    return estoque


def carregar_pedidos(caminho) -> pd.DataFrame:
    pedidos = pd.read_csv(caminho, parse_dates=["data"])
    _validar_colunas(pedidos, COLUNAS_PEDIDOS, str(caminho))
    return pedidos


# ---------------------------------------------------------------------------
# Análises
# ---------------------------------------------------------------------------

def divergencias_de_estoque(estoque: pd.DataFrame) -> pd.DataFrame:
    """Produtos em que o saldo do sistema não bate com a contagem física."""
    resultado = estoque.copy()
    resultado["diferenca"] = resultado["contagem_fisica"] - resultado["saldo_sistema"]
    resultado = resultado[resultado["diferenca"] != 0]
    resultado["impacto_rs"] = (resultado["diferenca"] * resultado["custo_unitario"]).round(2)
    colunas = ["sku", "produto", "saldo_sistema", "contagem_fisica", "diferenca", "impacto_rs"]
    return resultado[colunas].sort_values("impacto_rs").reset_index(drop=True)


def abaixo_do_minimo(estoque: pd.DataFrame) -> pd.DataFrame:
    """Produtos cuja contagem física está abaixo do estoque mínimo."""
    resultado = estoque[estoque["contagem_fisica"] < estoque["estoque_minimo"]].copy()
    resultado["repor"] = resultado["estoque_minimo"] - resultado["contagem_fisica"]
    resultado["custo_reposicao_rs"] = (resultado["repor"] * resultado["custo_unitario"]).round(2)
    colunas = ["sku", "produto", "contagem_fisica", "estoque_minimo", "repor", "custo_reposicao_rs"]
    return resultado[colunas].sort_values("repor", ascending=False).reset_index(drop=True)


def pendentes_sem_estoque(pedidos: pd.DataFrame, estoque: pd.DataFrame) -> pd.DataFrame:
    """Pedidos pendentes cujo SKU não tem quantidade suficiente na contagem física."""
    pendentes = pedidos[pedidos["status"] == "Pendente"]
    juntos = pendentes.merge(estoque[["sku", "produto", "contagem_fisica"]], on="sku", how="left")
    juntos["contagem_fisica"] = juntos["contagem_fisica"].fillna(0).astype(int)
    em_risco = juntos[juntos["quantidade"] > juntos["contagem_fisica"]].copy()
    em_risco["faltam"] = em_risco["quantidade"] - em_risco["contagem_fisica"]
    em_risco["data"] = em_risco["data"].dt.strftime("%d/%m/%Y")
    colunas = ["pedido", "data", "sku", "produto", "quantidade", "contagem_fisica", "faltam", "canal"]
    return em_risco[colunas].reset_index(drop=True)


def resumo_faturamento(pedidos: pd.DataFrame) -> pd.DataFrame:
    """Total faturado por canal de venda (considera apenas pedidos com status Faturado)."""
    faturados = pedidos[pedidos["status"] == "Faturado"].copy()
    faturados["valor_total"] = faturados["quantidade"] * faturados["valor_unitario"]
    resumo = (
        faturados.groupby("canal")
        .agg(pedidos=("pedido", "nunique"), itens=("quantidade", "sum"), faturamento_rs=("valor_total", "sum"))
        .reset_index()
    )
    resumo["ticket_medio_rs"] = (resumo["faturamento_rs"] / resumo["pedidos"]).round(2)
    resumo["faturamento_rs"] = resumo["faturamento_rs"].round(2)
    return resumo.sort_values("faturamento_rs", ascending=False).reset_index(drop=True)


# ---------------------------------------------------------------------------
# Geração do Excel
# ---------------------------------------------------------------------------

def _ajustar_largura(planilha) -> None:
    for coluna in planilha.columns:
        maior = max(len(str(celula.value)) if celula.value is not None else 0 for celula in coluna)
        planilha.column_dimensions[coluna[0].column_letter].width = min(maior + 3, 45)


def gerar_relatorio(estoque: pd.DataFrame, pedidos: pd.DataFrame, saida) -> dict:
    abas = {
        "Divergencias": divergencias_de_estoque(estoque),
        "Abaixo do minimo": abaixo_do_minimo(estoque),
        "Pedidos em risco": pendentes_sem_estoque(pedidos, estoque),
        "Faturamento": resumo_faturamento(pedidos),
    }
    with pd.ExcelWriter(saida, engine="openpyxl") as arquivo:
        for nome, tabela in abas.items():
            tabela.to_excel(arquivo, sheet_name=nome, index=False)
            planilha = arquivo.sheets[nome]
            planilha.freeze_panes = "A2"
            _ajustar_largura(planilha)
    return abas


def main(argumentos=None) -> int:
    pasta = Path(__file__).resolve().parent.parent
    parser = argparse.ArgumentParser(description="Gera relatório de estoque e pedidos em Excel.")
    parser.add_argument("--estoque", default=pasta / "dados" / "estoque.csv", help="CSV de estoque")
    parser.add_argument("--pedidos", default=pasta / "dados" / "pedidos.csv", help="CSV de pedidos")
    parser.add_argument("--saida", default="relatorio_estoque.xlsx", help="Arquivo Excel de saída")
    args = parser.parse_args(argumentos)

    try:
        estoque = carregar_estoque(args.estoque)
        pedidos = carregar_pedidos(args.pedidos)
    except (FileNotFoundError, ValueError) as erro:
        print(f"Erro: {erro}", file=sys.stderr)
        return 1

    abas = gerar_relatorio(estoque, pedidos, args.saida)

    print(f"Relatório gerado: {args.saida}")
    print(f"  Divergências de estoque: {len(abas['Divergencias'])} produto(s)")
    print(f"  Abaixo do mínimo:        {len(abas['Abaixo do minimo'])} produto(s)")
    print(f"  Pedidos em risco:        {len(abas['Pedidos em risco'])} pedido(s)")
    total = abas["Faturamento"]["faturamento_rs"].sum()
    print(f"  Faturamento total:       R$ {total:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
    return 0


if __name__ == "__main__":
    sys.exit(main())
