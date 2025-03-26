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
