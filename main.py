import os
from mrp import MRP

if __name__ == '__main__':
    print("MRP - Planejamento de Necessidades de Materiais")
    mrp = MRP(os.getcwd())
    demanda = {"ETI": 100, "ETF": 100}
    # Inicializa os dados do MRP
    mrp.inicializar_dados()
    # Calcula as quantidades de produção e aquisição
    mrp.calcular_quantidades_producao_aquisicao(demanda)
    # Calcula leadtimes
    mrp.calcular_fc_lt_esperados()
    # Monta o quadro de planejamento
    quadro_planejamento = mrp.montar_quadro_planejamento()
    mrp.imprimir_quadro_planejamento()