import unittest
import os
from mrp import MRP  # Importa a classe MRP


class TestMRPInicializacao(unittest.TestCase):
    """
    Classe de teste para verificar a inicialização do MRP (carregamento das BOMs).
    """

    def setUp(self):
        """
        Método que é executado antes de cada teste.
        Configura o ambiente de teste, criando uma instância da classe MRP.
        """
        self.diretorio_testes = os.getcwd()  # Diretório com os arquivos de teste
        self.mrp = MRP(self.diretorio_testes)

        # Cria o diretório de testes se não existir
        if not os.path.exists(self.diretorio_testes):
            os.makedirs(self.diretorio_testes)

        # Cria arquivos de BOM de exemplo para teste
        self.criar_arquivos_bom_teste()

    def tearDown(self):
        """
        Método que é executado após cada teste.
        Limpa o ambiente de teste, removendo os arquivos e diretórios criados.
        """
        # Remove o diretório de testes e seu conteúdo
        for arquivo in os.listdir(self.diretorio_testes):
            caminho_arquivo = os.path.join(self.diretorio_testes, arquivo)
            os.remove(caminho_arquivo)
        os.rmdir(self.diretorio_testes)

    def criar_arquivos_bom_teste(self):
        """
        Cria arquivos de BOM de exemplo para serem usados nos testes.
        """
        # Conteúdo dos arquivos de BOM de exemplo
        bom_eti = """
        Material,Quantidade
        ETI,1
        JOKER,1
        DAQ,1
        """

        bom_etf = """
        Material,Quantidade
        ETF,1
        ETI,1
        JOKER,1
        """

        # Cria os arquivos no diretório de testes
        with open(os.path.join(self.diretorio_testes, "ETI_BOM.xlsx"), "w") as f:
            f.write(bom_eti)

        with open(os.path.join(self.diretorio_testes, "ETF_BOM.xlsx"), "w") as f:
            f.write(bom_etf)

    def test_carregar_bom(self):
        """
        Testa se as BOMs são carregadas corretamente.
        """
        self.mrp.inicializar_dados()

        # Verifica se as BOMs foram carregadas corretamente
        self.assertIn("ETI", self.mrp.bom)
        self.assertIn("ETF", self.mrp.bom)
        self.assertEqual(self.mrp.bom["ETI"]["JOKER"], 1)
        self.assertEqual(self.mrp.bom["ETF"]["ETI"], 1)


if __name__ == '__main__':
    unittest.main()