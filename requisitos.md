**REQUISITOS MRP**

| VERSÃO | COMENTÁRIOS | RESPONSÁVEL | DATA |
| :---: | ----- | ----- | ----- |
| **1.0** | Versão inicial após prototipação | Rogério Atem | 07/03/2025 |
| **1.01** | \-Versão para GitHub\-Correções na redação dos requisitos | Rogério Atem | 13/03/2025 |

1) Trata-se do módulo de Materials Requirements Planning (MRP) de uma empresa que produz  dispositivos eletrônicos. Tal sistema servirá como base para uma versão mais complexa que levará em conta também os processos de produção, além dos materiais. Em termos de processo de negócio, o módulo MRP se baseia em quatro fases operacionais:  
   1) Inicialização;  
   2) Planejamento;  
   3) Execução e Controle;  
   4) Análise.  
2) Os conceitos mais básicos são de Produto Final e seus Componentes. Os únicos a serem produzidos são os produtos finais, enquanto que os componentes são adquiridos. Ambos tipos são mantidos no estoque,  representado por uma planilha que define também outras informações relativas às operações de movimentação de estoques e estimativas de custos. É realizado um planejamento inicial, com quantidades, tempos e custos levando-se em conta a demanda informada e os valores médios de tempos e custos, que gerarão então um conjunto de ordens de retirada de estoque, produção e /ou aquisição. À medida que as ordens são efetivamente executadas na linha de produção, quantidades, tempos e custos vão sendo atualizados \- através de confirmações, cancelamentos e alterações nas ordens originais. Desta forma, o planejamento vai sendo transformado em execução e, potencialmente, pode ser necessário um novo planejamento. As ordens de produção são relativas a produtos individuais, porém o MRP em si é feito para N produtos finais, que potencialmente possuem componentes em comum. As quantidades dos componentes comuns devem ser somadas, obviamente. A seguir são explicadas as fases e outros detalhes de implementação.

3) Fase de Inicialização:  
   1) Para cada Produto Final, obter suas planilhas de dados estáticos e dinâmicos:  
      1) Estáticos: Não mudam entre uma ordem e outra, apenas em ciclos de reprojeto do produto. Representados pela BOM, Bill Of Materials, que descreve para o primeiro elemento da lista (o produto final) quais e quantos componentes ele possui. Por exemplo, para cada ETF eu tenho uma Joker, 2 ADS1115 etc.   
      2) Dinâmicos: Mudam entre diferentes ordens e/ou ciclos de planejamento. Representados pelo Estoque, que descreve a Quantidade em Estoque, o Estoque Mínimo, o Lead Time Médio, o Custo Médio, Imposto Médio (quando aplicável) e Frete por Lote (quando aplicável), do produto final e dos componentes.  
      3) Para um dado ciclo MRP, podem ser demandados um ou mais produtos finais, que por sua vez possuem componentes em comum, sendo portanto necessário calcular as ordens considerando tal fator.   
   2) Os dados lidos dessas planilhas devem ser armazenados em dicionários, usando o código do produto e dos componentes como chaves.   
   3) Concluída a inicialização, mudar o estado do MRP para Inicializado.  
4) Fase de Planejamento:  
   1) Usar como entrada os dicionários preenchidos na Inicialização.  
      2) A lógica de cálculo do MRP é a seguinte: para cada produto final, calcular a quantidade do produto final multiplicada pelas quantidades obtidas na BOM para saber quantos componentes são necessários, por exemplo, se são demandados 30 ETF(produto final), serão necessários 30 DAQs e 60 AD1115. Verificar no estoque quando de cada material (produto e componentes) está registrado como disponível no estoque. Exemplificando, se existem 5 ETF e 10 DAQ no estoque, então será necessário obter 25 ETF e 20 DAQ. Deve ser considerado também o estoque mínimo, que são 5 e 10, respectivamente, no exemplo. Como a obtenção reduziria o estoque para abaixo do mínimo, deve-se calcular para também manter o mínimo, somando este valor ao de obtenção, quando necessário. Finalmente, o seguinte deve ser considerado: se o estoque for suficiente para atender à demanda e não baixar do mínimo, apenas retirar do estoque. Se o estoque baixar do mínimo, produzir (produto final) e/ou adquirir (componentes) na quantidade que atenda à demanda e reponha o estoque mínimo simultaneamente.  
      3) O cálculo intermediário descrito (quantidade a ser produzida ou adquirida sem considerar o estoque mínimo) anteriormente deve ser salvo, bem como a quantidade final, num dicionário denominado quantidades, que armazena as quantidades para cada produto ou componente.   
      4) Um segundo dicionário, denominado ordens\_planejamento, deve conter as quantidades, por produto ou componente, a serem produzidos (produto), adquiridos (componente) e/ou retirados do estoque (produto ou componente). Este dicionário descreve as ordens de produção (apenas para produto final), compra (apenas para componentes) ou de retirada do estoque (para ambos tipos).   
      5) Um terceiro dicionário deve ser criado, denominado fc\_lt\_esperados (fluxo de caixa e lead times esperados), com os custos e impostos médios de produção (produto) e obtenção (componentes) que deverão ser produzidos ou adquiridos (sem contar os retirados do estoque), bem como seus leadtimes médios. Os lead-times dos componentes são copiados do dicionário que os mantém, o do produto final é calculado como o maior lead-time entre seus componentes a adquirir somado ao seu próprio  lead-time de produção. A Lógica de Custos é a seguinte: Multiplicar o número de itens a ser produzido (produto) ou adquirido (componentes) pelo seus valores unitários e impostos unitários e somar o frete por lote, obtendo seu custo total. Atenção: o que for apenas retirado do estoque não entra na contagem de custo, pois seu custo já incidiu em ciclos anteriores.  
      6) Um quarto dicionário, denominado planejamento, deve ser alimentado pelas ordens e representar uma tabela, que tem por linhas cada produto ou componente, e por coluna a data e quantidades que ele é esperado de acordo com as ordens. A data é calculada em função do dia de execução do planejamento somado ao leadtime médio do produto ou componente. Se houver, para um mesmo produto ou componente, ordens múltiplas, fazer uma linha para cada ordem do material. A primeira coluna do dicionário planejamento deve se chamar estoque atual e conter os valores lidos do Estoque para fazer o planejamento.  
      7) Após finalizado, mudar o estado do MRP para Planejado.  
5) Fase de Execução e Controle  
   1) Iniciada a execução, mudar o estado do MRP para Em Execução.  
   2) O dicionário ordens\_planejamento deve ser copiado para o dicionário ordens\_controle, que contém, além de todos os campos do anterior e a primeira coluna com os valores atuais de estoque e, para cada ordem, uma flag que diz que se uma ordem de um determinado produto ou componente está ainda como planejada, foi iniciada ou já está pronta. Planejada significa que a ordem ainda está aguardando ser executada, Executada que foi executada, mas o material (produto ou componente) ainda não está fisicamente disponível  e Pronta que o material está disponível na linha de produção ou para ser despachado.  
   3) Os valores médios são substituídos, quando houverem, pelos valores obtidos das cotações, na planilha Cotacoes.xlsx. Quando não estiverem listados nesta planilha, são mantidos os médios. Toda vez que essa planilha for entrada durante o controle, os valores são atualizados por aqueles presentes nela.  
   4) Deve possuir funcionalidade de listar em tela ou exportar para planilhas as ordens. A planilha com todas as ordens deve ser única, contendo como linhas o produto ou componente e colunas as quantidades a retirar do estoque, produzir (apenas produtos) ou adquirir (apenas componentes).  
   5) Deve possuir funcionalidade de editar o dicionário de ordens, através de seleção de qual ordem a editar ou cancelar. Ao alterar quantidades de materiais, através da alteração ou cancelamento das ordens, dar a alterar o dicionário de estoque e dar opção de exportar o dicionário atualizado para planilha.  
      1) Caso haja alteração de quantidades ou do leadtime do componente com maior leadtime, é necessário refazer as fases de inicialização e planejamento.  
   6) Após o ciclo de produção ser fechado, o MRP deve mudar seu estado para Encerrado.  
6) Fase de Análise (Futuro):  
   1) Ser capaz de ler N planilhas de MRP e mostrar insights como alterações em ordens de produção, atrasos, variação de custos, flutuação de estoque etc.  
   2) A análise é feita sobre ciclos encerrados ou em execução.  
   3) É interessante para o usuário manter na mesma pasta versões das planilhas, a medida que as ordens vão sendo executadas, para que seja possível ter um histórico da evolução da execução e aprimorar valores médios empregados, bem como detectar gargalos e outros problemas.  
7) Formato das Planilhas  
   1) BOM  
      1) Cabeçalho (colunas): Material, Quantidade  
      2) Linhas: produto final e abaixo deste seus componentes  
   2) Estoque  
      1) Cabeçalho (colunas): Material, Em Estoque, Mínimo, Custo Médio Unitário, Imposto Médio Unitário, Custo Frete Médio Lote, Leadtime Médio Lote.  
      2) Linhas: produto final e abaixo deste seus componentes  
   3) Ordens  
      1) Cabeçalho (colunas): Material, Retirada de Estoque, Produção\*, Aquisição\*, Custo Total Estimado\*\*, Leadtime Total Estimado\*\*, Custo Total Real\*\*\*, Leadtime Total Real\*\*\*.  
         1) \*: Produção para produto final, Aquisição para componentes.  
         2) \*\*: Baseados nos valores médios.  
         3) \*\*\*: Baseados na execução, informados pelo usuário através de modificações nas ordens.  
8) Implementação do código:  
   1) Implementado inicialmente como um módulo Python rodando na linha de comando, em Linux ou Windows.   
   2) A classe MRP deve possuir um método principal para cada fase e métodos auxiliares para cálculos, criação e preenchimento de estruturas de dados e output e input em planilhas (e futuramente em outras fontes de dados), quando se aplicar. Os dicionários devem ser todos atributos do objeto MRP e não variáveis que são passadas como parâmetro.  
   3) As  planilhas devem ser carregadas (entradas) e/ou salvas (saídas) em pasta definida como parâmetro para o Construtor da classe MRP.   
   4) O nome da planilha de BOM deve ser no formato XXX\_BOM, onde XXX é o código do produto. Podem ser empregados prefixos com 3 a 6 caracteres para o código do produto. Exemplo simples de input para planejamento:   
      1) Produto: ETF  
         Quantidade: 10  
         (Considerar BOM em  ETF\_BOM)  
      2) Produto: ETI  
         Quantidade: 15  
         (Considerar BOM em  ETI\_BOM)  
      3) Supondo que cada um dos produtos usem 02 componentes do tipo AD1115, então o MRP deve prever a demanda por 50 AD1115, ou seja  (10+25)\*2.  
   5) A planilha de Estoque é única, com nome a ser informado pelo usuário.  
   6) Para exportar em planilhas, o usuário deve fornecer o nome das mesmas.  
   7) A planilha de ordens deve ser única para todos os produtos e componentes envolvidos no ciclo de produção. Porém, as alterações e cancelamentos podem ser feitas ordem a ordem, assim, versões atualizadas da planilha de ordens devem poder ser exportadas também.  
   8) Toda busca por e associação de materiais deve ser feita pelos códigos dos materiais, nunca usar a sequência nas planilhas, exceto na BOM, cuja primeira linha (excetuando-se cabeçalho) é sempre referente a um produto final.  
   9) Considerar possíveis erros de planilhas faltantes e demais erros de input por parte do usuário.  
   10) Sugere-se ao usuário manter uma pasta por ciclo produtivo, de maneira a ser possível rastrear as mudanças de estado do ciclo, permitindo real controle da linha.  
       

