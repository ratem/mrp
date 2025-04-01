## Fase 1: Inicialização do MRP

- **Entradas**:
    - Arquivos de BOM (ETI_BOM.xlsx, ETF_BOM.xlsx)
    - Dados de estoque (Estoque.xlsx)
- **Processamento**: Carregamento dos dados em dicionários
- **Saídas**:
    - Dicionário de BOMs (self.boms)
    - Dicionário de estoque (self.estoque)
    - Estado do MRP = "Inicializado"


## Fase 2: Planejamento MRP

- **Entradas**:
    - Demanda de produtos finais (ETI, ETF)
    - Dicionários de BOMs e estoque
- **Processamento**:
    - Cálculo de quantidades a produzir/adquirir
    - Cálculo de fluxo de caixa e leadtimes
- **Saídas**:
    - Dicionário de ordens de planejamento
    - Dicionário de fluxo de caixa e leadtimes
    - Quadro de planejamento (planejamento_atualizado.xlsx)
    - Estado do MRP = "Planejado"


## Fase 3: Execução e Controle do MRP

- **Entradas**:
    - Dicionário de ordens de planejamento
    - Cotações atualizadas (Cotacoes.xlsx)
- **Processamento**:
    - Atualização de status de ordens
    - Atualização de custos e leadtimes
- **Saídas**:
    - Dicionário de ordens de controle atualizado
    - Alertas de necessidade de replanejamento
    - Estado do MRP = "Em Execução"


## Fase 4: Inicialização do CRP

- **Entradas**:
    - Quadro de planejamento do MRP
    - Dados de demanda por recursos (demanda_recursos.xlsx)
    - Dados de capacidade de recursos (capacidade_recursos.xlsx)
    - Exceções de capacidade (excecoes_capacidade.xlsx)
- **Processamento**: Carregamento dos dados em dicionários
- **Saídas**:
    - Dicionário de planejamento MRP
    - Dicionário de demanda por recursos
    - Dicionário de capacidade de recursos
    - Dicionário de exceções de capacidade
    - Estado do CRP = "Inicializado"


## Fase 5: Planejamento CRP

- **Entradas**:
    - Dicionários carregados na inicialização do CRP
- **Processamento**:
    - Cálculo da demanda por operação
    - Geração de planilha interativa
- **Saídas**:
    - Dicionário de demanda por operação
    - Planilha CRP interativa (crp_planejamento.xlsx)
    - Estado do CRP = "Planejado"


## Fase 6: Ajuste do Planejamento CRP

- **Entradas**:
    - Planilha CRP interativa
    - Feedback do usuário (alocações diárias)
- **Processamento**:
    - Alocação de produtos por dia
    - Verificação de capacidade disponível
- **Saídas**:
    - Planilha CRP atualizada
    - Plano de produção viável considerando restrições de capacidade


## Fase 7: Análise

- **Entradas**:
    - Resultados do MRP e CRP
    - Histórico de planilhas e atualizações
- **Processamento**:
    - Identificação de desvios
    - Análise de eficiência
- **Saídas**:
    - Insights para melhoria contínua
    - Documentação do ciclo produtivo


## Fluxos de Realimentação

- Da Fase 3 para Fase 2: Quando há necessidade de replanejamento devido a alterações em leadtimes ou custos
- Da Fase 6 para Fase 5: Quando a alocação de recursos revela impossibilidade de atender à demanda no prazo
- Da Fase 7 para Fase 1: Insights para ajustar parâmetros em ciclos futuros

Esta representação detalhada do fluxo de informações mostra como o MRP e o CRP trabalham em conjunto, com o MRP determinando "o que" e "quando" produzir, e o CRP ajustando esse plano considerando "como" e "com quais recursos" produzir, respeitando as limitações reais de capacidade.

