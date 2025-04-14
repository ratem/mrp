# Sistema Integrado MRP-CRP

Este é um sistema integrado de Planejamento de Necessidades de Materiais (MRP) e Planejamento de Capacidade de Recursos (CRP) desenvolvido em Python.

## Descrição

O sistema MRP-CRP é uma ferramenta de planejamento de produção que combina:

- Material Requirements Planning (MRP): para calcular as necessidades de materiais e gerar ordens de produção/aquisição.
- Capacity Requirements Planning (CRP): para verificar e ajustar o planejamento considerando as restrições de capacidade dos recursos produtivos.


## Funcionalidades Principais

### MRP

- Carregamento de dados de estoque e BOMs (Bill of Materials)
- Cálculo de necessidades de materiais baseado na demanda
- Geração de ordens de produção e aquisição
- Cálculo de fluxo de caixa e leadtimes esperados
- Exportação de resultados para planilhas Excel


### CRP

- Carregamento de dados de capacidade de recursos
- Análise de demanda por operação
- Ajuste do planejamento considerando restrições de capacidade
- Geração de planilha interativa para alocação de recursos


### Interface Gráfica

- GUI intuitiva desenvolvida com Tkinter
- Fluxo de trabalho guiado para operações MRP e CRP
- Visualização e edição de ordens de produção/aquisição
- Exportação de resultados e relatórios


## Requisitos

- Python 3.7+
- Bibliotecas: pandas, openpyxl, tkinter


## Instalação

1. Clone o repositório:

```
git clone https://github.com/seu-usuario/mrp-crp-system.git
```

2. Navegue até o diretório do projeto:

```
cd mrp-crp-system
```

3. Instale as dependências:

```
pip install -r requirements.txt
```


## Uso

1. Execute o arquivo principal da interface gráfica:

```
python gui_mrp_crp.py
```

2. Siga as instruções na interface para:
    - Definir a pasta de trabalho
    - Inicializar o MRP
    - Planejar a produção
    - Executar o controle de produção
    - Inicializar o CRP
    - Realizar o planejamento de capacidade

## Estrutura de Arquivos

- `mrp.py`: Implementação da classe MRP
- `crp.py`: Implementação da classe CRP
- `gui_mrp_crp.py`: Interface gráfica do sistema
- `test_mrp.py`: Testes unitários para o MRP
- `test_crp.py`: Testes unitários para o CRP


## Testes

Execute os testes unitários com:

```
python -m unittest test_mrp.py
python -m unittest test_crp.py
```


## Contribuições

Contribuições são bem-vindas! Por favor, abra uma issue para discutir mudanças propostas ou envie um pull request com suas melhorias.

## Licença

Este projeto está licenciado sob a [MIT License](LICENSE).

