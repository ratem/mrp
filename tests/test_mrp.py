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
        self.pasta_testes = "test_data"  # Diretório com os arquivos de teste
        self.mrp = MRP(self.pasta_testes)

        # Cria o diretório de testes se não existir
        if not os.path.exists(self.pasta_testes):
            os.makedirs(self.pasta_testes)

        # Cria arquivos de BOM de exemplo para teste
        self.criar_arquivos_bom_teste()

        # Cria arquivo de Estoque de exemplo para teste
        self.criar_arquivo_estoque_teste()


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

    def criar_arquivos_bom_teste(self):
        """
        Cria arquivos de BOM de exemplo (XLSX) para serem usados nos testes.
        """
        from openpyxl import Workbook

        # Criar arquivo ETI_BOM.xlsx
        workbook_eti = Workbook()
        sheet_eti = workbook_eti.active
        sheet_eti.append(["Material", "Quantidade"])  # Cabeçalho
        sheet_eti.append(["ETI", 1])
        sheet_eti.append(["JOKER", 1])
        sheet_eti.append(["DAQ", 1])
        workbook_eti.save(os.path.join(self.pasta_testes, "ETI_BOM.xlsx"))

        # Criar arquivo ETF_BOM.xlsx
        workbook_etf = Workbook()
        sheet_etf = workbook_etf.active
        sheet_etf.append(["Material", "Quantidade"])  # Cabeçalho
        sheet_etf.append(["ETF", 1])
        sheet_etf.append(["ETI", 1])
        sheet_etf.append(["JOKER", 1])
        workbook_etf.save(os.path.join(self.pasta_testes, "ETF_BOM.xlsx"))

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

    def criar_arquivo_estoque_teste(self):
        """
        Cria arquivo de Estoque de exemplo (XLSX) para ser usado nos testes.
        """
        from openpyxl import Workbook

        # Criar arquivo Estoque.xlsx
        workbook_estoque = Workbook()
        sheet_estoque = workbook_estoque.active

        # Dados de Estoque (obtidos da planilha anexa)
        dados_estoque = [
            ["Material", "Em Estoque", "Mínimo", "Custo Medio Unitario", "Imposto Medio Unitario", "Frete Medio Lote",
             "Leadtime Medio Lote"],
            ["ETF", 0, 5, 100, 0, 0, 5],
            ["JOKER", 10, 10, 10, 1, 5, 10],
            ["DAQ", 10, 10, 11, 2, 6, 11],
            ["Módulo 3G/4G", 10, 10, 12, 3, 7, 12],
            ["Chip Internet Móvel", 5, 5, 13, 4, 8, 13],
            ["SX1276 (LoRa)", 20, 20, 14, 5, 9, 14],
            ["ADS1115", 10, 5, 15, 6, 10, 15],
            ["Conector DB25 fêmea", 30, 10, 16, 7, 11, 16],
            ["Soquetes 2x6p", 20, 20, 17, 8, 12, 17],
            ["Invólucro", 5, 5, 18, 9, 13, 18],
            ["Antena LoRa", 15, 10, 19, 10, 14, 19],
            ["D-SUB25 macho", 12, 50, 20, 11, 15, 20],
            ["ETI", 5, 5, 80, 0, 0, 5]
        ]

        # Escrever os dados na planilha
        for linha in dados_estoque:
            sheet_estoque.append(linha)

        # Salvar o arquivo
        workbook_estoque.save(os.path.join(self.pasta_testes, "Estoque.xlsx"))

    def test_calcular_quantidades_producao_aquisicao(self):
        """
        Testa o cálculo das quantidades de produção e aquisição.
        """
        # Define a demanda
        demanda = {"ETI": 100, "ETF": 100}

        # Calcula as quantidades de produção e aquisição
        quantidades, ordens_planejamento = self.mrp.calcular_quantidades_producao_aquisicao(demanda)

        # Verificações (ajustadas com os dados da planilha)
        # Produto ETI
        self.assertEqual(quantidades["ETI"]["quantidade_a_produzir"], 95)
        self.assertEqual(ordens_planejamento["ETI"]["Produção"], 95)
        self.assertEqual(quantidades["JOKER"]["quantidade_a_adquirir"], 190)
        self.assertEqual(ordens_planejamento["JOKER"]["Aquisição"], 190)
        self.assertEqual(quantidades["DAQ"]["quantidade_a_adquirir"], 190)
        self.assertEqual(ordens_planejamento["DAQ"]["Aquisição"], 190)
        self.assertEqual(quantidades["SX1276 (LoRa)"]["quantidade_a_adquirir"], 190)
        self.assertEqual(ordens_planejamento["SX1276 (LoRa)"]["Aquisição"], 190)
        self.assertEqual(quantidades["ADS1115"]["quantidade_a_adquirir"], 380)
        self.assertEqual(ordens_planejamento["ADS1115"]["Aquisição"], 380)
        self.assertEqual(quantidades["Conector DB25 fêmea"]["quantidade_a_adquirir"], 190)
        self.assertEqual(ordens_planejamento["Conector DB25 fêmea"]["Aquisição"], 190)
        self.assertEqual(quantidades["Soquetes 2x6p"]["quantidade_a_adquirir"], 190)
        self.assertEqual(ordens_planejamento["Soquetes 2x6p"]["Aquisição"], 190)
        self.assertEqual(quantidades["Invólucro"]["quantidade_a_adquirir"], 190)
        self.assertEqual(ordens_planejamento["Invólucro"]["Aquisição"], 190)
        self.assertEqual(quantidades["Antena LoRa"]["quantidade_a_adquirir"], 190)
        self.assertEqual(ordens_planejamento["Antena LoRa"]["Aquisição"], 190)
        self.assertEqual(quantidades["D-SUB25 macho"]["quantidade_a_adquirir"], 190)
        self.assertEqual(ordens_planejamento["D-SUB25 macho"]["Aquisição"], 190)

        # Produto ETF
        self.assertEqual(quantidades["ETF"]["quantidade_a_produzir"], 105)
        self.assertEqual(ordens_planejamento["ETF"]["Produção"], 105)
        self.assertEqual(quantidades["Módulo 3G/4G"]["quantidade_a_adquirir"], 100)
        self.assertEqual(ordens_planejamento["Módulo 3G/4G"]["Aquisição"], 100)
        self.assertEqual(quantidades["Chip Internet Móvel"]["quantidade_a_adquirir"], 100)
        self.assertEqual(ordens_planejamento["Chip Internet Móvel"]["Aquisição"], 100)

if __name__ == '__main__':
    unittest.main()