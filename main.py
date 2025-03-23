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
    mrp.montar_quadro_planejamento()
    #mrp.imprimir_quadro_planejamento()
    #mrp.exportar_quadro_planejamento("quadro.xlsx")
    #mrp.exportar_ordens_producao("ordens.xlsx")
    mrp.iniciar_execucao()
    mrp.listar_ordens_controle()
    mrp.atualizar_status_ordem("JOKER", 'Executada')
    # Edita a ordem novamente
    nova_quantidade = 80
    resultado = mrp.editar_ordem("JOKER", 'Aquisição', 80)

