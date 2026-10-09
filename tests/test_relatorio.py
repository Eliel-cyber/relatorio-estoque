import sys
import tempfile
import unittest
from pathlib import Path

import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))

import relatorio  # noqa: E402

ESTOQUE = pd.DataFrame({
    "sku": ["A", "B", "C"],
    "produto": ["Produto A", "Produto B", "Produto C"],
    "saldo_sistema": [10, 5, 2],
    "contagem_fisica": [10, 3, 2],
    "estoque_minimo": [4, 4, 5],
    "custo_unitario": [10.0, 20.0, 30.0],
})

PEDIDOS = pd.DataFrame({
    "pedido": [1, 2, 3, 4],
    "data": pd.to_datetime(["2026-09-01", "2026-09-01", "2026-09-02", "2026-09-02"]),
    "sku": ["A", "B", "C", "A"],
    "quantidade": [2, 1, 3, 1],
    "valor_unitario": [15.0, 30.0, 40.0, 15.0],
    "canal": ["Mercado Livre", "Loja Propria", "Mercado Livre", "Mercado Livre"],
    "status": ["Faturado", "Faturado", "Pendente", "Cancelado"],
})


class TestRelatorio(unittest.TestCase):
    def test_divergencias_mostra_apenas_produtos_com_diferenca(self):
        resultado = relatorio.divergencias_de_estoque(ESTOQUE)
        self.assertEqual(resultado["sku"].tolist(), ["B"])
        self.assertEqual(resultado.loc[0, "diferenca"], -2)
        self.assertEqual(resultado.loc[0, "impacto_rs"], -40.0)

    def test_abaixo_do_minimo_calcula_quanto_repor(self):
        resultado = relatorio.abaixo_do_minimo(ESTOQUE)
        self.assertEqual(set(resultado["sku"]), {"B", "C"})
        self.assertEqual(resultado.set_index("sku").loc["C", "repor"], 3)

    def test_pedido_pendente_sem_estoque_entra_em_risco(self):
        resultado = relatorio.pendentes_sem_estoque(PEDIDOS, ESTOQUE)
        self.assertEqual(resultado["pedido"].tolist(), [3])
        self.assertEqual(resultado.loc[0, "faltam"], 1)

    def test_faturamento_ignora_cancelados_e_pendentes(self):
        resultado = relatorio.resumo_faturamento(PEDIDOS).set_index("canal")
        self.assertEqual(resultado.loc["Mercado Livre", "faturamento_rs"], 30.0)
        self.assertEqual(resultado.loc["Loja Propria", "faturamento_rs"], 30.0)

    def test_arquivo_sem_coluna_obrigatoria_gera_erro_claro(self):
        with tempfile.TemporaryDirectory() as pasta:
            caminho = Path(pasta) / "estoque.csv"
            ESTOQUE.drop(columns=["contagem_fisica"]).to_csv(caminho, index=False)
            with self.assertRaises(ValueError) as erro:
                relatorio.carregar_estoque(caminho)
            self.assertIn("contagem_fisica", str(erro.exception))

    def test_gera_excel_com_quatro_abas(self):
        with tempfile.TemporaryDirectory() as pasta:
            saida = Path(pasta) / "relatorio.xlsx"
            relatorio.gerar_relatorio(ESTOQUE, PEDIDOS, saida)
            abas = pd.ExcelFile(saida).sheet_names
            self.assertEqual(abas, ["Divergencias", "Abaixo do minimo", "Pedidos em risco", "Faturamento"])


if __name__ == "__main__":
    unittest.main()
