<h1>Descrição do Projeto e Funcionalidades</h1>
<p>Este projeto tem como objetivo principal a padronização e análise de dados de produtos presentes em notas fiscais eletrônicas (NFes). Ele é dividido em duas etapas principais: a padronização de dados e a identificação de erros, seguida pela análise estatística e cálculo de variações. As funcionalidades do projeto incluem:</p>

<strong>Carregamento e validação de dados</strong>: Leitura de arquivos Excel contendo dados de produtos e tabela de erros.
<strong>Padronização de dados</strong>: Aplicação de intervalos de aceitação para valores unitários e filtragem de valores discrepantes.
<strong>Análise estatística</strong>: Cálculo de variâncias, desvios padrão e outras métricas estatísticas para cada grupo de produtos.
<strong>Geração de relatórios</strong>: Criação de relatórios detalhados com análise de erros, variações de preços e distribuição de vendas por setor.
<strong>Salvamento de resultados</strong>: Armazenamento dos dados padronizados e analisados em arquivos Excel para posterior revisão e utilização.


<h2>Detalhamento das Etapas</h2>
<h3>Primeira Etapa: Padronização de Dados e Identificação de Erros</h3>

<strong>Carregamento de Dados</strong>: Carregamento de arquivos Excel contendo dados padronizados e tabela de erros. Verificação e carregamento de arquivo de configuração <code>Gtin.xlsx</code>.
<strong>Inicialização de Estruturas</strong>: Criação de DataFrames para armazenar resultados, dados de desvio padrão, fornecedores e resultados finais. Inicialização de dicionários para armazenar variações e intervalos de aceitação.
<strong>Cálculo de Intervalos de Aceitação</strong>: Cálculo de limites inferior e superior de aceitação para valores unitários de produtos agrupados por código GTIN.
<strong>Filtragem de Dados</strong>: Filtragem de dados com base nos intervalos de aceitação. Valores dentro dos intervalos são considerados corretos, enquanto valores fora são classificados como erros.
<strong>Análise Estatística</strong>: Cálculo de variâncias e desvios padrão para cada grupo de produtos. Cálculo do valor total das notas fiscais para posterior análise percentual.
<strong>Compilação dos Resultados</strong>: Cálculo de várias métricas para cada produto e armazenamento em tabelas de resultados e fornecedores.
<strong>Salvamento dos Resultados</strong>: Salvamento de DataFrames resultantes em um arquivo Excel, organizados em várias abas, incluindo dados padronizados, tabela de erros, dados tratados, relatório e análise de erros por fornecedores.


<h3>Segunda Etapa: Análise Estatística e Cálculo de Variações</h3>
<strong>Preparação dos Dados</strong>: Carregamento dos dados previamente padronizados e filtrados. Inicialização de estruturas de dados para armazenar resultados da análise estatística
<strong>Cálculo de Variações</strong>: Agrupamento de dados por código GTIN e cálculo de soma de unidades comerciais e valor total do produto. Cálculo de variâncias e desvios padrão para cada grupo de produtos
<strong>Análise de Desvios Padrão</strong>: Identificação de produtos com variações fora dos intervalos de aceitação. Remoção de valores discrepantes e recalculação de métricas
<strong>Compilação de Relatórios</strong>: Criação de relatórios detalhados contendo análise de quantidade de erros, variações de preços e distribuições de vendas por setor. Preenchimento de tabelas com dados consolidados e estatísticas calculadas
<strong>Salvamento Final dos Resultados</strong>: Salvamento de todos os resultados e análises em um arquivo Excel, organizados em abas específicas para fácil acesso e revisão

<h2>Tecnologias Utilizadas</h2>
<strong>Python</strong>: Linguagem principal de programação utilizada para todo o desenvolvimento do projeto.
<strong>Pandas</strong>: Biblioteca essencial para manipulação e análise de dados, usada para operações com DataFrames.
<strong>Openpyxl</strong>: Biblioteca usada para leitura e escrita de arquivos Excel.
<strong>NumPy</strong>: Biblioteca usada para cálculos numéricos e operações matemáticas.
<strong>PyQt5</strong>: Biblioteca usada para criar a interface gráfica do usuário (GUI) que facilita a interação com o sistema.
<strong>Os</strong>: Módulo usado para operações relacionadas ao sistema operacional, como verificação da existência de arquivos.
<strong>Bibliotecas de estatística</strong>: Ferramentas para cálculos estatísticos como variância e desvio padrão.
