<!DOCTYPE html>
<html lang="pt-BR">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Descrição do Projeto</title>
    <style>
        body {
            font-family: Arial, sans-serif;
            line-height: 1.6;
            margin: 20px;
        }
        h1, h2, h3 {
            color: #333;
        }
        ul {
            list-style-type: disc;
            margin-left: 20px;
        }
        code {
            background: #f4f4f4;
            padding: 2px 4px;
            border-radius: 4px;
        }
    </style>
</head>
<body>
    <h1>Descrição do Projeto e Funcionalidades</h1>
    <p>Este projeto tem como objetivo principal a padronização e análise de dados de produtos presentes em notas fiscais eletrônicas (NFes). Ele é dividido em duas etapas principais: a padronização de dados e a identificação de erros, seguida pela análise estatística e cálculo de variações. As funcionalidades do projeto incluem:</p>
    <ul>
        <li><strong>Carregamento e validação de dados</strong>: Leitura de arquivos Excel contendo dados de produtos e tabela de erros.</li>
        <li><strong>Padronização de dados</strong>: Aplicação de intervalos de aceitação para valores unitários e filtragem de valores discrepantes.</li>
        <li><strong>Análise estatística</strong>: Cálculo de variâncias, desvios padrão e outras métricas estatísticas para cada grupo de produtos.</li>
        <li><strong>Geração de relatórios</strong>: Criação de relatórios detalhados com análise de erros, variações de preços e distribuição de vendas por setor.</li>
        <li><strong>Salvamento de resultados</strong>: Armazenamento dos dados padronizados e analisados em arquivos Excel para posterior revisão e utilização.</li>
    </ul>

    <h2>Detalhamento das Etapas</h2>
    <h3>Primeira Etapa: Padronização de Dados e Identificação de Erros</h3>
    <ol>
        <li><strong>Carregamento de Dados</strong>: Carregamento de arquivos Excel contendo dados padronizados e tabela de erros. Verificação e carregamento de arquivo de configuração <code>Gtin.xlsx</code>.</li>
        <li><strong>Inicialização de Estruturas</strong>: Criação de DataFrames para armazenar resultados, dados de desvio padrão, fornecedores e resultados finais. Inicialização de dicionários para armazenar variações e intervalos de aceitação.</li>
        <li><strong>Cálculo de Intervalos de Aceitação</strong>: Cálculo de limites inferior e superior de aceitação para valores unitários de produtos agrupados por código GTIN.</li>
        <li><strong>Filtragem de Dados</strong>: Filtragem de dados com base nos intervalos de aceitação. Valores dentro dos intervalos são considerados corretos, enquanto valores fora são classificados como erros.</li>
        <li><strong>Análise Estatística</strong>: Cálculo de variâncias e desvios padrão para cada grupo de produtos. Cálculo do valor total das notas fiscais para posterior análise percentual.</li>
        <li><strong>Compilação dos Resultados</strong>: Cálculo de várias métricas para cada produto e armazenamento em tabelas de resultados e fornecedores.</li>
        <li><strong>Salvamento dos Resultados</strong>: Salvamento de DataFrames resultantes em um arquivo Excel, organizados em várias abas, incluindo dados padronizados, tabela de erros, dados tratados, relatório e análise de erros por fornecedores.</li>
    </ol>

    <h3>Segunda Etapa: Análise Estatística e Cálculo de Variações</h3>
    <ol>
        <li><strong>Preparação dos Dados</strong>: Carregamento dos dados previamente padronizados e filtrados. Inicialização de estruturas de dados para armazenar resultados da análise estatística.</li>
        <li><strong>Cálculo de Variações</strong>: Agrupamento de dados por código GTIN e cálculo de soma de unidades comerciais e valor total do produto. Cálculo de variâncias e desvios padrão para cada grupo de produtos.</li>
        <li><strong>Análise de Desvios Padrão</strong>: Identificação de produtos com variações fora dos intervalos de aceitação. Remoção de valores discrepantes e recalculação de métricas.</li>
        <li><strong>Compilação de Relatórios</strong>: Criação de relatórios detalhados contendo análise de quantidade de erros, variações de preços e distribuições de vendas por setor. Preenchimento de tabelas com dados consolidados e estatísticas calculadas.</li>
        <li><strong>Salvamento Final dos Resultados</strong>: Salvamento de todos os resultados e análises em um arquivo Excel, organizados em abas específicas para fácil acesso e revisão.</li>
    </ol>

    <h2>Tecnologias Utilizadas</h2>
    <ul>
        <li><strong>Python</strong>: Linguagem principal de programação utilizada para todo o desenvolvimento do projeto.</li>
        <li><strong>Pandas</strong>: Biblioteca essencial para manipulação e análise de dados, usada para operações com DataFrames.</li>
        <li><strong>Openpyxl</strong>: Biblioteca usada para leitura e escrita de arquivos Excel.</li>
        <li><strong>NumPy</strong>: Biblioteca usada para cálculos numéricos e operações matemáticas.</li>
        <li><strong>PyQt5</strong>: Biblioteca usada para criar a interface gráfica do usuário (GUI) que facilita a interação com o sistema.</li>
        <li><strong>Os</strong>: Módulo usado para operações relacionadas ao sistema operacional, como verificação da existência de arquivos.</li>
        <li><strong>Bibliotecas de estatística</strong>: Ferramentas para cálculos estatísticos como variância e desvio padrão.</li>
    </ul>
</body>
</html>
