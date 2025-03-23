import unittest
import os
from mrp import MRP  # Importa a classe MRP
from openpyxl import Workbook


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
        # Criar arquivo ETI_BOM.xlsx
        workbook_eti = Workbook()
        sheet_eti = workbook_eti.active
        sheet_eti.append(["Material", "Quantidade"])  # Cabeçalho
        sheet_eti.append(["JOKER", 1])
        sheet_eti.append(["DAQ", 1])
        sheet_eti.append(["SX1276 (LoRa)", 1])
        sheet_eti.append(["ADS1115", 2])
        sheet_eti.append(["Conector DB25 fêmea", 1])
        sheet_eti.append(["Soquetes 2x6p", 1])
        sheet_eti.append(["Invólucro", 1])
        sheet_eti.append(["Antena LoRa", 1])
        sheet_eti.append(["D-SUB25 macho", 1])
        workbook_eti.save(os.path.join(self.pasta_testes, "ETI_BOM.xlsx"))

        # Criar arquivo ETF_BOM.xlsx
        workbook_etf = Workbook()
        sheet_etf = workbook_etf.active
        sheet_etf.append(["Material", "Quantidade"])  # Cabeçalho
        sheet_etf.append(["JOKER", 1])
        sheet_etf.append(["DAQ", 1])
        sheet_etf.append(["Módulo 3G/4G", 1])
        sheet_etf.append(["Chip Internet Móvel", 1])
        sheet_etf.append(["SX1276 (LoRa)", 1])
        sheet_etf.append(["ADS1115", 2])
        sheet_etf.append(["Conector DB25 fêmea", 1])
        sheet_etf.append(["Soquetes 2x6p", 1])
        sheet_etf.append(["Invólucro", 1])
        sheet_etf.append(["Antena LoRa", 1])
        sheet_etf.append(["D-SUB25 macho", 1])
        workbook_etf.save(os.path.join(self.pasta_testes, "ETF_BOM.xlsx"))

    def test_carregar_bom(self):
        """
        Testa se as BOMs são carregadas corretamente.
        """
        self.mrp.inicializar_dados()

        # Verifica se as BOMs foram carregadas corretamente
        self.assertIn("ETI", self.mrp.boms)
        self.assertIn("ETF", self.mrp.boms)
        self.assertEqual(self.mrp.boms["ETI"]["JOKER"], 1)
        self.assertEqual(self.mrp.boms["ETF"]["JOKER"], 1)

    def criar_arquivo_estoque_teste(self):
        """
        Cria arquivo de Estoque de exemplo (XLSX) para ser usado nos testes.
        """
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

        # Inicializa os dados do MRP
        self.mrp.inicializar_dados()

        # Calcula as quantidades de produção e aquisição
        quantidades, ordens_planejamento = self.mrp.calcular_quantidades_producao_aquisicao(demanda)

        # Verificações
        # Produto ETI
        self.assertEqual(quantidades["ETI"]["quantidade_a_produzir"], 100)
        self.assertEqual(ordens_planejamento["ETI"]["Produção"], 100)

        # Produto ETF
        self.assertEqual(quantidades["ETF"]["quantidade_a_produzir"], 105)
        self.assertEqual(ordens_planejamento["ETF"]["Produção"], 105)

        # Componentes (CORREÇÃO AQUI)
        # Estoque mínimo de JOKER é 10, estoque inicial é 10. Demanda é 205. Precisa adquirir 205.
        self.assertEqual(quantidades["JOKER"]["quantidade_a_adquirir"], 205)
        self.assertEqual(ordens_planejamento["JOKER"]["Aquisição"], 205)

        # Estoque mínimo de DAQ é 10, estoque inicial é 10. Demanda é 205. Precisa adquirir 205.
        self.assertEqual(quantidades["DAQ"]["quantidade_a_adquirir"], 205)
        self.assertEqual(ordens_planejamento["DAQ"]["Aquisição"], 205)

        # Estoque mínimo de SX1276 (LoRa) é 20, estoque inicial é 20. Demanda é 205. Precisa adquirir 205.
        self.assertEqual(quantidades["SX1276 (LoRa)"]["quantidade_a_adquirir"], 205)
        self.assertEqual(ordens_planejamento["SX1276 (LoRa)"]["Aquisição"], 205)

        # Estoque mínimo de ADS1115 é 5, estoque inicial é 10. Demanda é 410. Precisa adquirir 405.
        self.assertEqual(quantidades["ADS1115"]["quantidade_a_adquirir"], 405)
        self.assertEqual(ordens_planejamento["ADS1115"]["Aquisição"], 405)

        # Estoque mínimo de Conector DB25 fêmea é 10, estoque inicial é 30. Demanda é 205. Precisa adquirir 185.
        self.assertEqual(quantidades["Conector DB25 fêmea"]["quantidade_a_adquirir"], 185)
        self.assertEqual(ordens_planejamento["Conector DB25 fêmea"]["Aquisição"], 185)

        # Estoque mínimo de Soquetes 2x6p é 20, estoque inicial é 20. Demanda é 205. Precisa adquirir 205.
        self.assertEqual(quantidades["Soquetes 2x6p"]["quantidade_a_adquirir"], 205)
        self.assertEqual(ordens_planejamento["Soquetes 2x6p"]["Aquisição"], 205)

        # Estoque mínimo de Invólucro é 5, estoque inicial é 5. Demanda é 205. Precisa adquirir 205.
        self.assertEqual(quantidades["Invólucro"]["quantidade_a_adquirir"], 205)
        self.assertEqual(ordens_planejamento["Invólucro"]["Aquisição"], 205)

        # Estoque mínimo de Antena LoRa é 10, estoque inicial é 15. Demanda é 205. Precisa adquirir 200.
        self.assertEqual(quantidades["Antena LoRa"]["quantidade_a_adquirir"], 200)
        self.assertEqual(ordens_planejamento["Antena LoRa"]["Aquisição"], 200)

        # Estoque mínimo de D-SUB25 macho é 50, estoque inicial é 12. Demanda é 205. Precisa adquirir 248.
        self.assertEqual(quantidades["D-SUB25 macho"]["quantidade_a_adquirir"], 243)
        self.assertEqual(ordens_planejamento["D-SUB25 macho"]["Aquisição"], 243)

        # Estoque mínimo de Módulo 3G/4G é 10, estoque inicial é 10. Demanda é 105. Precisa adquirir 105.
        self.assertEqual(quantidades["Módulo 3G/4G"]["quantidade_a_adquirir"], 105)
        self.assertEqual(ordens_planejamento["Módulo 3G/4G"]["Aquisição"], 105)

        # Estoque mínimo de Chip Internet Móvel é 5, estoque inicial é 5. Demanda é 105. Precisa adquirir 105.
        self.assertEqual(quantidades["Chip Internet Móvel"]["quantidade_a_adquirir"], 105)
        self.assertEqual(ordens_planejamento["Chip Internet Móvel"]["Aquisição"], 105)

    def test_calcular_fc_lt_esperados(self):
        """
        Testa o cálculo do fluxo de caixa e lead times esperados.
        """
        # Define a demanda
        demanda = {"ETI": 100, "ETF": 100}

        # Inicializa os dados do MRP
        self.mrp.inicializar_dados()

        # Calcula as quantidades de produção e aquisição
        self.mrp.calcular_quantidades_producao_aquisicao(
            demanda)  # Garante que self.ordens_planejamento esteja preenchido
        fc_lt_esperados = self.mrp.calcular_fc_lt_esperados()

        # Verificações
        # Produto ETI
        self.assertIn("ETI", fc_lt_esperados)
        self.assertIn("Leadtime", fc_lt_esperados["ETI"])
        self.assertIn("Custo", fc_lt_esperados["ETI"])
        self.assertEqual(fc_lt_esperados["ETI"]["Leadtime"], 25)
        self.assertAlmostEqual(fc_lt_esperados["ETI"]["Custo"], 8000)  # Verifique o custo

        # Produto ETF
        self.assertIn("ETF", fc_lt_esperados)
        self.assertIn("Leadtime", fc_lt_esperados["ETF"])
        self.assertIn("Custo", fc_lt_esperados["ETF"])
        self.assertEqual(fc_lt_esperados["ETF"]["Leadtime"], 25)
        self.assertAlmostEqual(fc_lt_esperados["ETF"]["Custo"], 10500)

        # Componente JOKER
        self.assertIn("JOKER", fc_lt_esperados)
        self.assertIn("Leadtime", fc_lt_esperados["JOKER"])
        self.assertIn("Custo", fc_lt_esperados["JOKER"])
        self.assertEqual(fc_lt_esperados["JOKER"]["Leadtime"], 10)
        self.assertAlmostEqual(fc_lt_esperados["JOKER"]["Custo"], 2260)

        # Componente DAQ
        self.assertIn("DAQ", fc_lt_esperados)
        self.assertIn("Leadtime", fc_lt_esperados["DAQ"])
        self.assertIn("Custo", fc_lt_esperados["DAQ"])
        self.assertEqual(fc_lt_esperados["DAQ"]["Leadtime"], 11)
        self.assertAlmostEqual(fc_lt_esperados["DAQ"]["Custo"], 2671)

        # Componente Módulo 3G/4G
        self.assertIn("Módulo 3G/4G", fc_lt_esperados)
        self.assertEqual(fc_lt_esperados["Módulo 3G/4G"]["Leadtime"], 12)
        self.assertAlmostEqual(fc_lt_esperados["Módulo 3G/4G"]["Custo"], 1582)

        # Componente Chip Internet Móvel
        self.assertIn("Chip Internet Móvel", fc_lt_esperados)
        self.assertEqual(fc_lt_esperados["Chip Internet Móvel"]["Leadtime"], 13)
        self.assertAlmostEqual(fc_lt_esperados["Chip Internet Móvel"]["Custo"], 1793)

        # Componente SX1276 (LoRa)
        self.assertIn("SX1276 (LoRa)", fc_lt_esperados)
        self.assertEqual(fc_lt_esperados["SX1276 (LoRa)"]["Leadtime"], 14)
        self.assertAlmostEqual(fc_lt_esperados["SX1276 (LoRa)"]["Custo"], 3904)

        # Componente ADS1115
        self.assertIn("ADS1115", fc_lt_esperados)
        self.assertEqual(fc_lt_esperados["ADS1115"]["Leadtime"], 15)
        self.assertAlmostEqual(fc_lt_esperados["ADS1115"]["Custo"], 8515)

        # Componente Conector DB25 fêmea
        self.assertIn("Conector DB25 fêmea", fc_lt_esperados)
        self.assertEqual(fc_lt_esperados["Conector DB25 fêmea"]["Leadtime"], 16)
        self.assertAlmostEqual(fc_lt_esperados["Conector DB25 fêmea"]["Custo"], 4266)

        # Componente Soquetes 2x6p
        self.assertIn("Soquetes 2x6p", fc_lt_esperados)
        self.assertEqual(fc_lt_esperados["Soquetes 2x6p"]["Leadtime"], 17)
        self.assertAlmostEqual(fc_lt_esperados["Soquetes 2x6p"]["Custo"], 5137)

        # Componente Invólucro
        self.assertIn("Invólucro", fc_lt_esperados)
        self.assertEqual(fc_lt_esperados["Invólucro"]["Leadtime"], 18)
        self.assertAlmostEqual(fc_lt_esperados["Invólucro"]["Custo"], 5548)

        # Componente Antena LoRa
        self.assertIn("Antena LoRa", fc_lt_esperados)
        self.assertEqual(fc_lt_esperados["Antena LoRa"]["Leadtime"], 19)
        self.assertAlmostEqual(fc_lt_esperados["Antena LoRa"]["Custo"], 5814)

        # Componente D-SUB25 macho
        self.assertIn("D-SUB25 macho", fc_lt_esperados)
        self.assertEqual(fc_lt_esperados["D-SUB25 macho"]["Leadtime"], 20)
        self.assertAlmostEqual(fc_lt_esperados["D-SUB25 macho"]["Custo"], 7548)


if __name__ == '__main__':
    unittest.main()