# Roteiro de Uso do Sistema MRP-CRP

O sistema integrado MRP-CRP implementado segue um fluxo de trabalho estruturado que combina o planejamento de necessidades de materiais (MRP) com o planejamento de capacidade de recursos (CRP). Este roteiro apresenta as fases principais e instruções detalhadas para utilização do sistema.

## Fase 1: Inicialização do MRP

Nesta fase, o sistema carrega os dados necessários para o planejamento:

- Carrega as BOMs (Bill of Materials) para cada produto final
- Carrega os dados de estoque, incluindo custos médios e leadtimes médios
- Prepara as estruturas de dados para as próximas fases

```python
# Definir o diretório de trabalho e criar pastas para os ciclos
diretorio_base = os.getcwd()
data_ciclo = datetime.now().strftime("%Y%m%d_%H%M")
pasta_ciclo_mrp = os.path.join(diretorio_base, f"ciclo_mrp_{data_ciclo}")
pasta_ciclo_crp = os.path.join(diretorio_base, f"ciclo_crp_{data_ciclo}")

# Criar diretórios se não existirem
if not os.path.exists(pasta_ciclo_mrp):
    os.makedirs(pasta_ciclo_mrp)
if not os.path.exists(pasta_ciclo_crp):
    os.makedirs(pasta_ciclo_crp)

# Criar instâncias do MRP e CRP
mrp = MRP(pasta_ciclo_mrp)
crp = CRP(pasta_ciclo_crp)

# Copiar arquivos necessários para as pastas dos ciclos
arquivos_mrp = ["Estoque.xlsx", "ETI_BOM.xlsx", "ETF_BOM.xlsx", "Cotacoes.xlsx"]
for arquivo in arquivos_mrp:
    if os.path.exists(os.path.join(diretorio_base, arquivo)):
        import shutil
        shutil.copy(os.path.join(diretorio_base, arquivo), os.path.join(pasta_ciclo_mrp, arquivo))

# Inicializar dados do MRP
print("Inicializando dados do MRP...")
mrp.inicializar_dados()
```


## Fase 2: Planejamento MRP

Com base na demanda informada, o sistema:

- Calcula as quantidades a serem produzidas (produtos finais) e adquiridas (componentes)
- Calcula o fluxo de caixa e leadtimes esperados
- Monta o quadro de planejamento com datas de entrega previstas

```python
# Definir a demanda
demanda = {"ETI": 100, "ETF": 155}
print(f"Demanda definida: {demanda}")

# Realizar o planejamento MRP
print("\nRealizando planejamento MRP...")
mrp.planejar_producao(demanda)

# Exibir resultados do MRP em formato tabular
print("\nQuadro de planejamento MRP:")
mrp.imprimir_quadro_planejamento()

# Iniciar execução para poder listar ordens
mrp.iniciar_execucao()

print("\nOrdens de produção e aquisição:")
mrp.listar_ordens_controle()

print("\nCustos estimados:")
mrp.listar_custos_materiais()

# Exportar o quadro de planejamento
arquivo_planejamento = "planejamento_atualizado.xlsx"
caminho_planejamento = mrp.exportar_quadro_planejamento(arquivo_planejamento)
print(f"\nQuadro de planejamento exportado para: {arquivo_planejamento}")

# Copiar o planejamento para a pasta do CRP
import shutil
shutil.copy(caminho_planejamento, os.path.join(pasta_ciclo_crp, arquivo_planejamento))
```


## Fase 3: Inicialização do CRP

Nesta fase, o sistema prepara os dados necessários para o planejamento de capacidade:

- Carrega o planejamento do MRP
- Carrega dados de demanda de recursos por produto
- Carrega dados de capacidade nominal dos recursos
- Carrega exceções de capacidade para datas específicas

```python
# Criar arquivos necessários para o CRP
print("Criando arquivos de configuração para o CRP...")

# Criar arquivo de demanda por recursos
from openpyxl import Workbook

# Demanda por recursos
wb_demanda = Workbook()
ws_demanda = wb_demanda.active
ws_demanda.append(["Produto", "OP1", "OP2"])
ws_demanda.append(["ETI", 20, 40])
ws_demanda.append(["ETF", 30, 30])
caminho_demanda = os.path.join(pasta_ciclo_crp, "demanda_recursos.xlsx")
wb_demanda.save(caminho_demanda)
print("Arquivo de demanda por recursos criado.")

# Capacidade de recursos
wb_capacidade = Workbook()
ws_capacidade = wb_capacidade.active
ws_capacidade.append(["Recurso", "OP1", "OP2", "OP3"])
ws_capacidade.append(["RE1", 480, 0, 480])
ws_capacidade.append(["RE2", 0, 480, 480])
caminho_capacidade = os.path.join(pasta_ciclo_crp, "capacidade_recursos.xlsx")
wb_capacidade.save(caminho_capacidade)
print("Arquivo de capacidade de recursos criado.")

# Exceções de capacidade
wb_excecoes = Workbook()
ws_excecoes = wb_excecoes.active

# Adicionar dados para RE1
ws_excecoes.append(["RE1", "OP1", "OP3"])
ws_excecoes.append(["2025-04-01", 120, 120])  # Redução significativa no primeiro dia
ws_excecoes.append(["2025-04-02", 60, 60])    # Redução média no segundo dia
ws_excecoes.append(["2025-04-03", 30, 30])    # Redução pequena no terceiro dia

# Adicionar dados para RE2
ws_excecoes.append(["RE2", "OP2", "OP3"])
ws_excecoes.append(["2025-04-01", 30, 30])    # Redução pequena no primeiro dia
ws_excecoes.append(["2025-04-02", 180, 180])  # Redução grande no segundo dia
ws_excecoes.append(["2025-04-03", 90, 90])    # Redução média no terceiro dia

caminho_excecoes = os.path.join(pasta_ciclo_crp, "excecoes_capacidade.xlsx")
wb_excecoes.save(caminho_excecoes)
print("Arquivo de exceções de capacidade criado com reduções significativas nos primeiros 3 dias.")

# Inicializar o CRP com os dados do MRP
print("\nInicializando o CRP com o planejamento do MRP...")
crp.carregar_planejamento_mrp(arquivo_planejamento)

# Carregar dados de demanda por recursos
print("Carregando demanda por recursos...")
crp.carregar_demanda_recursos("demanda_recursos.xlsx")

# Carregar dados de capacidade de recursos
print("Carregando capacidade de recursos...")
crp.carregar_capacidade_recursos("capacidade_recursos.xlsx")

# Carregar exceções de capacidade
print("Carregando exceções de capacidade...")
crp.carregar_excecoes_capacidade("excecoes_capacidade.xlsx")
```


## Fase 4: Planejamento CRP

Nesta fase, o sistema:

- Calcula a demanda por operação para cada produto
- Gera uma planilha interativa para planejamento de capacidade
- Permite a alocação de produtos por dia, considerando restrições de capacidade

```python
# Calcular demanda por operação
print("\nCalculando demanda por operação...")
demanda_por_operacao = crp.calcular_demanda_por_operacao()

# Exibir demanda por operação
print("\nDemanda por operação:")
for operacao, produtos in demanda_por_operacao.items():
    print(f"  {operacao}:")
    for produto, minutos in produtos.items():
        print(f"    {produto}: {minutos} minutos")

# Criar planilha CRP para planejamento interativo
print("\nCriando planilha CRP para planejamento interativo...")
data_atual = datetime.now().strftime("%Y-%m-%d")
arquivo_crp = "crp_planejamento.xlsx"
dias_planejamento = 5
caminho_crp = crp.criar_planilha_crp(arquivo_crp, data_atual, dias_planejamento)
```


## Fase 5: Execução e Controle do MRP

Durante esta fase:

- As ordens são iniciadas, executadas e finalizadas
- Os custos e leadtimes são atualizados com base em cotações reais
- Quando necessário, um replanejamento é realizado

```python
# Atualizar o status de algumas ordens para "Executada"
print("\nAtualizando status de algumas ordens para 'Executada'...")
materiais_em_execucao = ["JOKER", "DAQ", "ADS1115"]
for material in materiais_em_execucao:
    mrp.atualizar_status_ordem(material, "Executada")

# Atualizar custos e leadtimes com base em novas cotações
print("\nAtualizando custos e leadtimes com base em novas cotações...")
sucesso, alertas = mrp.atualizar_custos_leadtimes("Cotacoes.xlsx")

# Verificar se há necessidade de replanejamento
necessidade_replanejar = any("Necessidade de replanejar" in alerta for alerta in alertas)
if necessidade_replanejar:
    print("Detectado aumento significativo em leadtimes. Realizando replanejamento...")
    mrp.planejar_producao(demanda)
    
    # Exportar o novo quadro de planejamento
    arquivo_planejamento_atualizado = "planejamento_atualizado.xlsx"
    mrp.exportar_quadro_planejamento(arquivo_planejamento_atualizado)
    
    # Reiniciar a execução com o novo planejamento
    mrp.iniciar_execucao()
```


## Fase 6: Ajuste do Planejamento CRP

Após um replanejamento do MRP, pode ser necessário ajustar o planejamento de capacidade:

- Gerar uma nova planilha CRP com base no planejamento atualizado
- Alocar produtos considerando as novas datas e quantidades

```python
# Após replanejamento, atualizar o CRP
if necessidade_replanejar:
    # Carregar o novo planejamento do MRP
    crp.carregar_planejamento_mrp(arquivo_planejamento_atualizado)
    
    # Calcular nova demanda por operação
    demanda_por_operacao = crp.calcular_demanda_por_operacao()
    
    # Criar nova planilha CRP
    arquivo_crp_atualizado = "crp_planejamento_atualizado.xlsx"
    caminho_crp = crp.criar_planilha_crp(arquivo_crp_atualizado, data_atual, dias_planejamento)
    print(f"Nova planilha CRP criada em: {arquivo_crp_atualizado}")
```


## Instruções para Uso da Planilha CRP

A planilha CRP gerada é uma ferramenta interativa para planejamento de capacidade:

1. Abra a planilha CRP gerada
2. Na aba "Demanda Total", verifique a demanda total por produto
3. Em cada aba de alocação diária:
    - Insira as quantidades de produtos a serem produzidos no dia
    - Observe na tabela de recursos o impacto nas capacidades
    - Ajuste as quantidades para evitar sobrecargas (células em vermelho)
4. Verifique na aba "Demanda Total" o que já foi alocado e o que ainda está pendente
5. Continue o processo até alocar toda a demanda ou até identificar a necessidade de mais dias de produção
```python
print("\nInstruções para uso da planilha CRP:")
print("1. Abra a planilha CRP gerada")
print(f"2. Nas {dias_planejamento} abas de alocação diária, insira as quantidades de produtos a serem produzidos")
print("3. Observe o impacto nas capacidades dos recursos em cada dia")
print("4. Ajuste as quantidades para otimizar a utilização dos recursos")
print("5. Verifique na aba 'Demanda Total' o que já foi alocado e o que ainda está pendente")
```


## Fase 7: Análise

Após a conclusão do ciclo integrado MRP-CRP:

- Os resultados são analisados para identificar desvios
- Insights são gerados para melhorar ciclos futuros
- As planilhas geradas são arquivadas para referência

```python
print("\n" + "=" * 80)
print("ANÁLISE DO CICLO INTEGRADO MRP-CRP")
print("=" * 80)
print("- O planejamento MRP gerou ordens de produção e aquisição")
print("- O planejamento CRP distribuiu a produção considerando restrições de capacidade")
if necessidade_replanejar:
    print("- Foi necessário realizar um replanejamento devido a alterações em leadtimes")
print("- As planilhas geradas foram salvas para referência futura")
print("\nCiclo integrado MRP-CRP concluído com sucesso!")
```

Este roteiro fornece uma visão completa do fluxo de trabalho do sistema integrado MRP-CRP, desde a inicialização até a análise final. O código Python apresentado pode ser usado como base para implementar cada fase do processo.


