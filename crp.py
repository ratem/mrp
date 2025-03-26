import os
import pandas as pd
from datetime import datetime


class CRP:
    """
    Classe que implementa o Capacity Requirements Planning (CRP).

    O CRP é responsável por ajustar o planejamento do MRP considerando
    as restrições de capacidade dos recursos produtivos.
    """

    def __init__(self, pasta_arquivos):
        """
        Inicializa o CRP.

        Args:
            pasta_arquivos (str): Caminho para a pasta onde os arquivos serão lidos/salvos.
        """
        self.pasta_arquivos = pasta_arquivos
        self.estado = "Não Inicializado"
        self.planejamento_mrp = {}
        self.datas_entrega = []
        self.produtos = []

    def carregar_planejamento_mrp(self, nome_arquivo):
        """
        Carrega o planejamento do MRP a partir de uma planilha Excel.

        Args:
            nome_arquivo (str): Nome do arquivo Excel contendo o planejamento do MRP.

        Returns:
            bool: True se o planejamento foi carregado com sucesso, False caso contrário.
        """
        try:
            # Constrói o caminho completo do arquivo
            caminho_arquivo = os.path.join(self.pasta_arquivos, nome_arquivo)

            # Lê a planilha Excel
            df = pd.read_excel(caminho_arquivo)

            # Inicializa o dicionário de planejamento
            self.planejamento_mrp = {}

            # Extrai as datas de entrega (todas as colunas exceto 'Material' e 'Estoque Atual')
            self.datas_entrega = [col for col in df.columns if col not in ['Material', 'Estoque Atual']]

            # Extrai os produtos (todos os valores da coluna 'Material')
            self.produtos = df['Material'].tolist()

            # Itera sobre as linhas do DataFrame
            for _, row in df.iterrows():
                material = row['Material']
                self.planejamento_mrp[material] = {'Estoque Atual': row['Estoque Atual']}

                # Adiciona as datas e quantidades ao dicionário
                for data in self.datas_entrega:
                    quantidade = row[data]
                    if pd.notna(quantidade) and quantidade > 0:
                        self.planejamento_mrp[material][data] = quantidade

            # Atualiza o estado para "Inicializado"
            self.estado = "Inicializado"

            print(f"Planejamento do MRP carregado com sucesso de: {caminho_arquivo}")
            return True

        except Exception as e:
            print(f"Erro ao carregar o planejamento do MRP: {str(e)}")
            return False

    def carregar_demanda_recursos(self, nome_arquivo):
        """
        Carrega a planilha de demanda por recursos.

        Args:
            nome_arquivo (str): Nome do arquivo Excel contendo a demanda por recursos.

        Returns:
            bool: True se a demanda foi carregada com sucesso, False caso contrário.
        """
        try:
            caminho_arquivo = os.path.join(self.pasta_arquivos, nome_arquivo)
            df = pd.read_excel(caminho_arquivo)

            self.demanda_recursos = {}
            for _, row in df.iterrows():
                produto = row['Produto']
                self.demanda_recursos[produto] = {}
                for col in df.columns:
                    if col != 'Produto' and pd.notna(row[col]):
                        self.demanda_recursos[produto][col] = row[col]

            print(f"Demanda por recursos carregada com sucesso de: {caminho_arquivo}")
            return True
        except Exception as e:
            print(f"Erro ao carregar a demanda por recursos: {str(e)}")
            return False

    def calcular_demanda_por_operacao(self):
        """
        Calcula a demanda por operação com base no planejamento do MRP e na demanda por recursos.

        Returns:
            dict: Demanda por operação, organizada por produto e operação.
        """
        if not hasattr(self, 'demanda_recursos'):
            print("Erro: Demanda por recursos não foi carregada.")
            return None

        demanda_por_operacao = {}

        for produto, planejamento in self.planejamento_mrp.items():
            if produto not in self.demanda_recursos:
                continue

            total_unidades = sum(
                quantidade for data, quantidade in planejamento.items()
                if data != "Estoque Atual"
            )

            for operacao, minutos_por_unidade in self.demanda_recursos[produto].items():
                if operacao not in demanda_por_operacao:
                    demanda_por_operacao[operacao] = {}

                demanda_por_operacao[operacao][produto] = total_unidades * minutos_por_unidade

        return demanda_por_operacao

    def carregar_capacidade_recursos(self, nome_arquivo):
        """
        Carrega a planilha de capacidade de produção nominal por recurso e operação.

        Args:
            nome_arquivo (str): Nome do arquivo Excel contendo a capacidade dos recursos.

        Returns:
            bool: True se a capacidade foi carregada com sucesso, False caso contrário.
        """
        try:
            caminho_arquivo = os.path.join(self.pasta_arquivos, nome_arquivo)
            df = pd.read_excel(caminho_arquivo)

            # Inicializa o dicionário de capacidade de recursos
            self.capacidade_recursos = {}

            # A primeira coluna deve conter os recursos
            recursos = df.iloc[:, 0].tolist()

            # As demais colunas são as operações
            operacoes = df.columns[1:].tolist()

            # Preenche o dicionário de capacidade
            for i, recurso in enumerate(recursos):
                self.capacidade_recursos[recurso] = {}
                for operacao in operacoes:
                    self.capacidade_recursos[recurso][operacao] = df.iloc[i][operacao]

            print(f"Capacidade de recursos carregada com sucesso de: {caminho_arquivo}")
            return True

        except Exception as e:
            print(f"Erro ao carregar a capacidade de recursos: {str(e)}")
            return False


    def carregar_excecoes_capacidade(self, nome_arquivo):
        """
        Carrega a planilha de exceções de capacidade, que contém reduções na capacidade
        nominal dos recursos em datas específicas.

        Args:
            nome_arquivo (str): Nome do arquivo Excel contendo as exceções de capacidade.

        Returns:
            bool: True se as exceções foram carregadas com sucesso, False caso contrário.
        """
        try:
            caminho_arquivo = os.path.join(self.pasta_arquivos, nome_arquivo)
            df = pd.read_excel(caminho_arquivo, header=None, parse_dates=False)

            self.excecoes_capacidade = {}
            recurso_atual = None
            operacoes = []

            for _, row in df.iterrows():
                # Se a primeira célula contém um recurso
                if pd.notna(row.iloc[0]) and isinstance(row.iloc[0], str) and not row.iloc[0].startswith('20'):
                    recurso_atual = row.iloc[0]
                    operacoes = [op for op in row.iloc[1:].tolist() if pd.notna(op)]
                    self.excecoes_capacidade[recurso_atual] = {}
                    for op in operacoes:
                        self.excecoes_capacidade[recurso_atual][op] = {}
                # Se a primeira célula contém uma data
                elif recurso_atual and pd.notna(row.iloc[0]):
                    # Converter a data para string no formato YYYY-MM-DD
                    if isinstance(row.iloc[0], pd.Timestamp):
                        data = row.iloc[0].strftime('%Y-%m-%d')
                    else:
                        try:
                            data = pd.to_datetime(row.iloc[0]).strftime('%Y-%m-%d')
                        except:
                            data = str(row.iloc[0])

                    # Processar cada operação
                    for i, operacao in enumerate(operacoes):
                        if i + 1 < len(row) and pd.notna(row.iloc[i + 1]):
                            self.excecoes_capacidade[recurso_atual][operacao][data] = row.iloc[i + 1]

            print(f"Exceções de capacidade carregadas com sucesso de: {caminho_arquivo}")
            return True

        except Exception as e:
            print(f"Erro ao carregar as exceções de capacidade: {str(e)}")
            return False