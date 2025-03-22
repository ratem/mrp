import os
from mrp import MRP

if __name__ == '__main__':
    mrp = MRP(os.getcwd())
    mrp.inicializar_dados()
    print(os.listdir(mrp.diretorio_planilhas))
    print(mrp.bom)