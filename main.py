import os
from mrp import MRP

if __name__ == '__main__':
    '''
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
    mrp.listar_custos_materiais()
    mrp.exportar_custos_materiais("custos.xlsx")
    mrp.atualizar_status_ordem("JOKER", 'Executada')
    # Edita a ordem novamente
    nova_quantidade = 80
    resultado = mrp.editar_ordem("JOKER", 'Aquisição', 80)
    '''
import os
from datetime import datetime
from mrp import MRP

if __name__ == '__main__':
    print("=" * 80)
    print("SIMULAÇÃO DO SISTEMA MRP - MATERIAL REQUIREMENTS PLANNING")
    print("=" * 80)

    # Definir o diretório de trabalho e criar uma pasta para o ciclo atual
    diretorio_base = os.getcwd()
    data_ciclo = datetime.now().strftime("%Y%m%d_%H%M")
    pasta_ciclo = os.path.join(diretorio_base, f"ciclo_{data_ciclo}")

    if not os.path.exists(pasta_ciclo):
        os.makedirs(pasta_ciclo)

    # Criar instância do MRP
    mrp = MRP(pasta_ciclo)

    # Copiar arquivos necessários para a pasta do ciclo
    for arquivo in ["Estoque.xlsx", "ETI_BOM.xlsx", "ETF_BOM.xlsx", "Cotacoes.xlsx"]:
        if os.path.exists(os.path.join(diretorio_base, arquivo)):
            import shutil

            shutil.copy(os.path.join(diretorio_base, arquivo), os.path.join(pasta_ciclo, arquivo))

    print("\n" + "=" * 80)
    print("FASE 1: INICIALIZAÇÃO")
    print("=" * 80)
    print("Carregando dados de estoque e BOMs...")
    mrp.inicializar_dados()
    print("Dados inicializados com sucesso.")

    # Definir a demanda
    demanda = {"ETI": 100, "ETF": 150}
    print(f"\nDemanda definida: {demanda}")

    print("\n" + "=" * 80)
    print("FASE 2: PLANEJAMENTO INICIAL")
    print("=" * 80)
    print("Realizando planejamento inicial...")

    # Realizar o planejamento
    mrp.planejar_producao(demanda)

    # Listar o quadro de planejamento
    mrp.imprimir_quadro_planejamento()

    # Exportar o quadro de planejamento inicial
    arquivo_planejamento_inicial = "planejamento_inicial.xlsx"
    mrp.exportar_quadro_planejamento(arquivo_planejamento_inicial)
    print(f"Quadro de planejamento inicial exportado para: {arquivo_planejamento_inicial}")

    # Exportar as ordens iniciais
    arquivo_ordens_inicial = "ordens_inicial.xlsx"
    mrp.exportar_ordens_producao(arquivo_ordens_inicial)
    print(f"Ordens de produção iniciais exportadas para: {arquivo_ordens_inicial}")

    # Mostrar custos estimados
    print("\nCustos estimados no planejamento inicial:")
    custo_total_inicial = mrp.listar_custos_materiais()
    mrp.exportar_custos_materiais("custos_inicial.xlsx")

    print("\n" + "=" * 80)
    print("FASE 3: EXECUÇÃO E CONTROLE")
    print("=" * 80)
    print("Iniciando execução do planejamento...")

    # Iniciar a execução
    mrp.iniciar_execucao()

    # Listar ordens de controle
    print("\nOrdens de controle no início da execução:")
    mrp.listar_ordens_controle()

    # Atualizar o status de algumas ordens para "Executada"
    print("\nAtualizando status de algumas ordens para 'Executada'...")
    materiais_em_execucao = ["JOKER", "DAQ", "ADS1115"]
    for material in materiais_em_execucao:
        mrp.atualizar_status_ordem(material, "Executada")

    print("\nOrdens de controle após atualização de status:")
    mrp.listar_ordens_controle()

    # Atualizar custos e leadtimes com base em novas cotações
    print("\nAtualizando custos e leadtimes com base em novas cotações...")
    sucesso, alertas = mrp.atualizar_custos_leadtimes("Cotacoes.xlsx")

    # Verificar se há necessidade de replanejamento
    necessidade_replanejar = any("Necessidade de replanejar" in alerta for alerta in alertas)

    if necessidade_replanejar:
        print("\n" + "=" * 80)
        print("REPLANEJAMENTO NECESSÁRIO")
        print("=" * 80)
        print("Detectado aumento significativo em leadtimes. Realizando replanejamento...")

        # Salvar o planejamento atual
        mrp.exportar_quadro_planejamento("planejamento_pre_atualizacao.xlsx")

        # Realizar novo planejamento com os valores atualizados
        mrp.planejar_producao(demanda)

        # Exportar o novo quadro de planejamento
        arquivo_planejamento_atualizado = "planejamento_atualizado.xlsx"
        mrp.exportar_quadro_planejamento(arquivo_planejamento_atualizado)
        print(f"Novo quadro de planejamento exportado para: {arquivo_planejamento_atualizado}")

        # Exportar as novas ordens
        arquivo_ordens_atualizado = "ordens_atualizado.xlsx"
        mrp.exportar_ordens_producao(arquivo_ordens_atualizado)
        print(f"Novas ordens de produção exportadas para: {arquivo_ordens_atualizado}")

        # Mostrar novos custos estimados
        print("\nCustos estimados após replanejamento:")
        custo_total_atualizado = mrp.listar_custos_materiais()
        mrp.exportar_custos_materiais("custos_atualizado.xlsx")

        # Calcular a diferença de custo
        diferenca_custo = custo_total_atualizado - custo_total_inicial
        print(f"\nDiferença de custo após replanejamento: R$ {diferenca_custo:.2f}")

        # Reiniciar a execução com o novo planejamento
        print("\nReiniciando execução com o novo planejamento...")
        mrp.iniciar_execucao()

    # Continuar a execução
    print("\nContinuando a execução...")

    # Atualizar o status de algumas ordens para "Executada"
    materiais_em_execucao = ["JOKER", "DAQ", "ADS1115", "Invólucro", "Antena LoRa"]
    for material in materiais_em_execucao:
        mrp.atualizar_status_ordem(material, "Executada")

    # Finalizar algumas ordens (status "Pronta")
    print("\nFinalizando algumas ordens (status 'Pronta')...")
    materiais_prontos = ["JOKER", "DAQ"]
    for material in materiais_prontos:
        mrp.atualizar_status_ordem(material, "Pronta")

    # Verificar a atualização do estoque
    print("\nOrdens de controle após finalização de algumas ordens:")
    mrp.listar_ordens_controle()

    # Editar uma ordem
    print("\nEditando a ordem de 'ADS1115'...")
    mrp.editar_ordem("ADS1115", "Aquisição", 450)

    # Cancelar uma ordem
    print("\nCancelando a ordem de 'Invólucro'...")
    mrp.cancelar_ordem("Invólucro")

    # Listar ordens após edições
    print("\nOrdens de controle após edições:")
    mrp.listar_ordens_controle()

    # Exportar o estado final das ordens
    arquivo_ordens_final = "ordens_final.xlsx"
    mrp.exportar_ordens_producao(arquivo_ordens_final)
    print(f"Estado final das ordens exportado para: {arquivo_ordens_final}")

    print("\n" + "=" * 80)
    print("FASE 4: ANÁLISE")
    print("=" * 80)
    print("Análise do ciclo de produção:")
    print("- Foram detectadas alterações significativas em custos e leadtimes")
    print("- Um replanejamento foi necessário devido ao aumento de leadtimes")
    print(f"- A diferença de custo após o replanejamento foi de R$ {diferenca_custo:.2f}")
    print("- Algumas ordens foram concluídas com sucesso")
    print("- Algumas ordens foram editadas ou canceladas durante a execução")

    print("\nCiclo de produção concluído com sucesso!")


