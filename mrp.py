import os
import openpyxl


class MRP:
    def __init__(self, pasta_planilhas):
        """
        Inicializa a classe MRP.

        Args:
            pasta_planilhas (str): O caminho para a pasta que contém as planilhas de dados.
        """
        self.pasta_planilhas = pasta_planilhas
        self.estoque = {}
        self.boms = {}  # Inicializa o dicionário boms
        self.ordens_planejamento = {}
        self.ordens_controle = {}
        self.planejamento = {}
        self.fc_lt_esperados = {}

    def inicializar_dados(self):
        """
        Lê as planilhas de BOM e armazena os dados em dicionários.
        """

        # 1. Ler todas as planilhas BOM no diretório
        for arquivo in os.listdir(self.pasta_planilhas):
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

    def calcular_quantidades_producao_aquisicao(self, demanda):
        """
        Calcula as quantidades de produção e aquisição necessárias para atender à demanda.

        Args:
            demanda (dict): Um dicionário contendo a demanda para cada produto final.

        Returns:
            tuple: Um dicionário contendo as quantidades calculadas para cada material
                   e um dicionário contendo as ordens de planejamento.
        """
        quantidades = {}
        ordens_planejamento = {}

        # Processa cada produto na demanda
        for produto, quantidade_demandada in demanda.items():
            # Inicializa as entradas para o produto
            quantidades[produto] = {"quantidade_a_produzir": 0, "quantidade_a_adquirir": 0}
            ordens_planejamento[produto] = {"Produção": 0, "Aquisição": 0, "Retirada de Estoque": 0}

            # Obtém a BOM do produto
            bom = self.boms.get(produto)

            # Se o produto não tiver BOM, considera que não há componentes
            if bom is None:
                continue

            # Calcula a necessidade do produto final
            estoque_atual = self.estoque.get(produto, {}).get("Em Estoque", 0)
            estoque_minimo = self.estoque.get(produto, {}).get("Mínimo", 0)
            necessidade_produto = quantidade_demandada

            quantidade_a_produzir = 0
            if estoque_atual < estoque_minimo:
                quantidade_a_produzir = estoque_minimo - estoque_atual

            quantidade_a_produzir += demanda

            quantidades[produto]["quantidade_a_produzir"] = quantidade_a_produzir
            ordens_planejamento[produto]["Produção"] = quantidade_a_produzir

            # Calcula a necessidade de cada componente
            for componente, quantidade_na_bom in bom.items():
                if componente == "Material":
                    continue  # Ignora a linha de cabeçalho

                quantidade_necessaria_componente = quantidade_demandada * float(quantidade_na_bom)
                estoque_atual_componente = self.estoque.get(componente, {}).get("Em Estoque", 0)
                estoque_minimo_componente = self.estoque.get(componente, {}).get("Mínimo", 0)
                necessidade_componente = quantidade_necessaria_componente + estoque_minimo_componente

                if estoque_atual_componente < necessidade_componente:
                    quantidades[componente]["quantidade_a_adquirir"] = max(0,
                                                                           necessidade_componente - estoque_atual_componente)
                    ordens_planejamento[componente]["Aquisição"] = max(0,
                                                                       necessidade_componente - estoque_atual_componente)
                else:
                    ordens_planejamento[componente]["Retirada de Estoque"] = quantidade_necessaria_componente

        return quantidades, ordens_planejamento

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