import os
import openpyxl


class MRP:
    def __init__(self, diretorio_planilhas):
        """
        Inicializa o objeto MRP com o diretório onde as planilhas estão localizadas.

        Args:
            diretorio_planilhas (str): Caminho para o diretório contendo as planilhas.
        """
        self.diretorio_planilhas = diretorio_planilhas
        self.bom = {}  # Dicionário para armazenar as BOMs
        self.estoque = {}  # Dicionário para armazenar os dados de estoque
        self.estado = "Não Inicializado"

    def inicializar_dados(self):
        """
        Lê as planilhas de BOM e armazena os dados em dicionários.
        """

        # 1. Ler todas as planilhas BOM no diretório
        for arquivo in os.listdir(self.diretorio_planilhas):
            if arquivo.endswith("_BOM.xlsx"):
                # 2. Extrair o código do produto do nome do arquivo
                codigo_produto = arquivo.replace("_BOM.xlsx", "")

                # Inicializar dicionário para armazenar os componentes
                self.bom[codigo_produto] = {}

                # Ler os dados da planilha BOM
                caminho_arquivo = os.path.join(self.diretorio_planilhas, arquivo)
                workbook = openpyxl.load_workbook(caminho_arquivo)
                sheet = workbook.active

                # Ler os dados da planilha, considerando a codificação
                primeira_linha = True
                for row in sheet.iter_rows(values_only=True):
                    if primeira_linha:
                        primeira_linha = False
                        continue  # Pular o cabeçalho

                    # Verificar se a linha contém dados de componentes
                    try:
                        quantidade = int(row[1])  # Tentar converter para inteiro
                        componente = row[0]
                        self.bom[codigo_produto][componente] = quantidade
                    except (ValueError, TypeError):
                        # Se não for possível converter para inteiro, é o nome do produto
                        continue

        self.estado = "Inicializado"

    def planejar_producao(self, demanda):
        """
        Realiza o planejamento da produção com base na demanda e nos dados inicializados.

        Args:
            demanda (dict): Dicionário contendo os produtos finais e suas respectivas quantidades demandadas.
                            Exemplo: {"ETF": 10, "ETI": 15}
        """
        if self.estado != "Inicializado":
            print("Erro: Os dados ainda não foram inicializados.")
            return

        self.quantidades = {}
        self.ordens_planejamento = {}
        self.fc_lt_esperados = {}
        self.planejamento = {}

        for produto, quantidade_demandada in demanda.items():
            # 1) Calcular a quantidade do produto final multiplicada pelas quantidades obtidas na BOM
            if produto not in self.bom:
                print(f"Aviso: Produto {produto} não encontrado na BOM. Ignorando.")
                continue

            self.quantidades[produto] = self.quantidades.get(produto, 0) + quantidade_demandada

            for componente, quantidade_por_produto in self.bom[produto].items():
                quantidade_necessaria = quantidade_demandada * quantidade_por_produto
                self.quantidades[componente] = self.quantidades.get(componente, 0) + quantidade_necessaria

            # 2) Verificar no estoque quando de cada material (produto e componentes) está registrado como disponível no estoque.
            # 3) Lógica de produção e aquisição (considerando o estoque mínimo)
            # 4) Salvar as quantidades num dicionário denominado quantidades
            # 5) Um segundo dicionário, denominado ordens_planejamento
            # 6) Um terceiro dicionário, denominado fc_lt_esperados
            # 7) Um quarto dicionário, denominado planejamento
            # 8) Após finalizado, mudar o estado do MRP para Planejado.
            # Implementar a lógica de planejamento conforme os requisitos

        self.estado = "Planejado"

    def executar_controle(self, cotações):
        """
        Executa e controla as ordens de produção e aquisição.

        Args:
            cotações (dict): Dicionário contendo os valores atualizados dos materiais.
        """
        if self.estado != "Planejado":
            print("Erro: O planejamento ainda não foi realizado.")
            return

        # 1) Iniciada a execução, mudar o estado do MRP para Em Execução.
        # 2) O dicionário ordens_planejamento deve ser copiado para o dicionário ordens_controle
        # 3) Os valores médios são substituídos, quando houverem, pelos valores obtidos das cotações
        # 4) Deve possuir funcionalidade de listar em tela ou exportar para planilhas as ordens.
        # 5) Deve possuir funcionalidade de editar o dicionário de ordens
        # 6) Após o ciclo de produção ser fechado, o MRP deve mudar seu estado para Encerrado.
        # Implementar a lógica de execução e controle conforme os requisitos

        self.estado = "Em Execução"

    def analisar_resultados(self):
        """
        Analisa os resultados do ciclo de produção.
        """
        if self.estado != "Encerrado":
            print("Erro: O ciclo de produção ainda não foi encerrado.")
            return

        # 1) Ser capaz de ler N planilhas de MRP e mostrar insights como alterações em ordens de produção, atrasos, variação de custos, flutuação de estoque etc.
        # 2) A análise é feita sobre ciclos encerrados ou em execução.
        # 3) É interessante para o usuário manter na mesma pasta versões das planilhas, a medida que as ordens vão sendo executadas, para que seja possível ter um histórico da evolução da execução e aprimorar valores médios empregados, bem como detectar gargalos e outros problemas.
        # Implementar a lógica de análise conforme os requisitos
        pass