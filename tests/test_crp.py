import unittest
import os
import pandas as pd
from openpyxl import Workbook
from crp import CRP


class TestCRPInicializacao(unittest.TestCase):
    """
    Classe de teste para verificar a inicialização do CRP.
    """

    def setUp(self):
        """
        Método que é executado antes de cada teste.
        Configura o ambiente de teste, criando uma instância da classe CRP.
        """
        self.pasta_testes = "test_data_crp"

        # Cria o diretório de testes se não existir
        if not os.path.exists(self.pasta_testes):
            os.makedirs(self.pasta_testes)

        # Cria os arquivos necessários para o teste
        self.criar_arquivo_planejamento_teste()

        # Inicializa o CRP com o arquivo de planejamento
        self.crp = CRP(self.pasta_testes)

    def tearDown(self):
        """
        Método que é executado após cada teste.
        Limpa o ambiente de teste, removendo os arquivos e diretórios criados.
        """
        # Remove o diretório de testes e seu conteúdo
        for arquivo in os.listdir(self.pasta_testes):
            caminho_arquivo = os.path.join(self.pasta_testes, arquivo)
            os.remove(caminho_arquivo)
        os.rmdir(self.pasta_testes)

    def criar_arquivo_planejamento_teste(self):
        """
        Cria um arquivo de planejamento de teste para ser usado nos testes.
        """
        # Criar um novo workbook
        wb = Workbook()
        ws = wb.active

        # Adicionar cabeçalhos
        ws.append(["Material", "Estoque Atual", "2023-10-01", "2023-10-15", "2023-10-30"])

        # Adicionar dados
        ws.append(["ETF", 5, 100, 50, 75])
        ws.append(["ETI", 10, 80, 60, 90])
        ws.append(["JOKER", 15, 200, 100, 150])
        ws.append(["DAQ", 20, 150, 75, 125])

        # Salvar o arquivo
        wb.save(os.path.join(self.pasta_testes, "planejamento_mrp.xlsx"))

    def test_inicializar_crp(self):
        """
        Testa se o CRP é inicializado corretamente.
        """
        # Verifica se o CRP foi inicializado
        self.assertIsNotNone(self.crp)

        # Verifica se o diretório de arquivos foi definido corretamente
        self.assertEqual(self.crp.pasta_arquivos, self.pasta_testes)

        # Verifica se o estado inicial é "Não Inicializado"
        self.assertEqual(self.crp.estado, "Não Inicializado")

    def test_carregar_planejamento_mrp(self):
        """
        Testa se o planejamento do MRP é carregado corretamente.
        """
        # Carrega o planejamento do MRP
        resultado = self.crp.carregar_planejamento_mrp("planejamento_mrp.xlsx")

        # Verifica se o carregamento foi bem-sucedido
        self.assertTrue(resultado)

        # Verifica se o estado mudou para "Inicializado"
        self.assertEqual(self.crp.estado, "Inicializado")

        # Verifica se o dicionário de planejamento foi preenchido corretamente
        planejamento = self.crp.planejamento_mrp
        self.assertIn("ETF", planejamento)
        self.assertIn("ETI", planejamento)
        self.assertIn("JOKER", planejamento)
        self.assertIn("DAQ", planejamento)

        # Verifica os valores do estoque atual
        self.assertEqual(planejamento["ETF"]["Estoque Atual"], 5)
        self.assertEqual(planejamento["ETI"]["Estoque Atual"], 10)

        # Verifica os valores das datas de entrega
        self.assertEqual(planejamento["ETF"]["2023-10-01"], 100)
        self.assertEqual(planejamento["ETF"]["2023-10-15"], 50)
        self.assertEqual(planejamento["ETF"]["2023-10-30"], 75)

        # Verifica se as datas foram extraídas corretamente
        self.assertIn("2023-10-01", self.crp.datas_entrega)
        self.assertIn("2023-10-15", self.crp.datas_entrega)
        self.assertIn("2023-10-30", self.crp.datas_entrega)

        # Verifica se os produtos foram extraídos corretamente
        self.assertIn("ETF", self.crp.produtos)
        self.assertIn("ETI", self.crp.produtos)
        self.assertIn("JOKER", self.crp.produtos)
        self.assertIn("DAQ", self.crp.produtos)


if __name__ == '__main__':
    unittest.main()
