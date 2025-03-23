import pandas as pd
import os
from openpyxl import load_workbook

class MRP:
    """
    Classe que implementa o Material Requirements Planning (MRP).
    """

    def __init__(self, pasta_arquivos):
        """
        Inicializa o objeto MRP com a pasta contendo os arquivos de entrada.
        """
        self.pasta_arquivos = pasta_arquivos
        self.estoque = {}
        self.boms = {}  # Inicializa o atributo boms como um dicionário
        self.ordens_planejamento = {}

    def carregar_estoque(self, df_estoque):
        """
        Carrega os dados de estoque a partir de um DataFrame do pandas.
        """
        for _, row in df_estoque.iterrows():
            codigo_material = row['Material']
            self.estoque[codigo_material] = {
                'em_estoque': row['Em Estoque'],
                'minimo': row['Mínimo'],
                'custo_medio_unitario': row['Custo Medio Unitario'],
                'imposto_medio_unitario': row['Imposto Medio Unitario'],
                'frete_medio_lote': row['Frete Medio Lote'],
                'leadtime_medio_lote': row['Leadtime Medio Lote']
            }

    def carregar_bom(self, codigo_produto, df_bom):
        """
        Carrega a BOM (Bill of Materials) a partir de um DataFrame do pandas.
        """
        self.boms[codigo_produto] = {}
        for _, row in df_bom.iterrows():
            material = row['Material']
            quantidade = row['Quantidade']
            self.boms[codigo_produto][material] = quantidade

    def inicializar_dados(self):
        """
        Carrega os dados de estoque e BOMs a partir dos arquivos XLSX.
        """
        # Carrega o estoque
        estoque_path = os.path.join(self.pasta_arquivos, "Estoque.xlsx")
        estoque_wb = load_workbook(estoque_path)
        estoque_sheet = estoque_wb.active
        df_estoque = pd.DataFrame(estoque_sheet.values)
        df_estoque.columns = df_estoque.iloc[0]
        df_estoque = df_estoque[1:]
        self.carregar_estoque(df_estoque)

        # Carrega as BOMs
        for arquivo in os.listdir(self.pasta_arquivos):
            if "_BOM.xlsx" in arquivo:
                codigo_produto = arquivo.replace("_BOM.xlsx", "")
                bom_path = os.path.join(self.pasta_arquivos, arquivo)
                bom_wb = load_workbook(bom_path)
                bom_sheet = bom_wb.active
                df_bom = pd.DataFrame(bom_sheet.values)
                df_bom.columns = df_bom.iloc[0]
                df_bom = df_bom[1:]
                self.carregar_bom(codigo_produto, df_bom)

    def calcular_quantidades_producao_aquisicao(self, demanda):
        """
        Calcula as quantidades de produção e aquisição para atender à demanda.
        """
        quantidades = {}
        self.ordens_planejamento = {}

        # 1. Calcular a quantidade a produzir para os produtos finais
        for produto, quantidade_demanda in demanda.items():
            # Inicializa as entradas para o produto
            if produto not in quantidades:
                quantidades[produto] = {'quantidade_a_produzir': 0, 'quantidade_a_adquirir': 0}
            if produto not in self.ordens_planejamento:
                self.ordens_planejamento[produto] = {'Produção': 0, 'Aquisição': 0}

            # Calcula a quantidade a produzir do produto final
            estoque_atual = self.estoque.get(produto, {'em_estoque': 0})['em_estoque']
            estoque_minimo = self.estoque.get(produto, {'minimo': 0})['minimo']
            quantidade_necessaria = max(0, quantidade_demanda + estoque_minimo - estoque_atual)
            quantidades[produto]['quantidade_a_produzir'] = quantidade_necessaria
            self.ordens_planejamento[produto]['Produção'] = quantidade_necessaria

        # 2. Calcular a quantidade a adquirir para os componentes
        componentes_necessarios = {}  # Dicionário para acumular as necessidades dos componentes

        for produto, quantidade_demanda in demanda.items():
            if produto in self.boms:
                bom = self.boms[produto]
                # Calcula a quantidade necessária para o produto
                estoque_atual_produto = self.estoque.get(produto, {'em_estoque': 0})['em_estoque']
                estoque_minimo_produto = self.estoque.get(produto, {'minimo': 0})['minimo']
                quantidade_necessaria = max(0, quantidade_demanda + estoque_minimo_produto - estoque_atual_produto)

                for componente, quantidade_por_produto in bom.items():
                    # Calcula a necessidade total do componente
                    necessidade_componente = quantidade_necessaria * quantidade_por_produto

                    # Acumula as necessidades dos componentes
                    if componente not in componentes_necessarios:
                        componentes_necessarios[componente] = 0
                    componentes_necessarios[componente] += necessidade_componente

        # Calcula a quantidade a adquirir para cada componente, considerando o estoque
        for componente, necessidade_total in componentes_necessarios.items():
            # Inicializa o componente no dicionário quantidades, se necessário
            if componente not in quantidades:
                quantidades[componente] = {'quantidade_a_produzir': 0, 'quantidade_a_adquirir': 0}
            if componente not in self.ordens_planejamento:
                self.ordens_planejamento[componente] = {'Produção': 0, 'Aquisição': 0}

            # Calcula a quantidade a adquirir (CORREÇÃO AQUI)
            estoque_componente = self.estoque.get(componente, {'em_estoque': 0})['em_estoque']
            estoque_minimo_componente = self.estoque.get(componente, {'minimo': 0})['minimo']
            quantidade_a_adquirir = max(0, necessidade_total + estoque_minimo_componente - estoque_componente)

            # Define a quantidade a adquirir
            quantidades[componente]['quantidade_a_adquirir'] = quantidade_a_adquirir
            self.ordens_planejamento[componente]['Aquisição'] = quantidade_a_adquirir

        return quantidades, self.ordens_planejamento

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