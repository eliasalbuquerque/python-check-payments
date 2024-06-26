<!--
title: 'requisicao-tecnica.md'
author: 'Elias Albuquerque'
created: '2024-06-20'
update: '2024-06-20'
-->


<!-- omit from toc -->
## Requisição Técnica: Atualização de Planilha com Dados da Web

- [1. Introdução](#1-introdução)
- [2. Objetivo](#2-objetivo)
- [3. Descrição do Processo](#3-descrição-do-processo)
- [4. Tecnologias](#4-tecnologias)
- [5. Considerações](#5-considerações)
- [6. Restrições](#6-restrições)
- [7. Próximos Passos](#7-próximos-passos)
- [8. Versão do Documento](#8-versão-do-documento)
- [9. Conclusão](#9-conclusão)


### 1. Introdução

Este documento descreve a requisição técnica para o desenvolvimento de um 
script Python que automatiza a atualização de uma planilha Excel com dados 
coletados de um website. O script visa otimizar o processo manual de consulta e 
atualização de informações, garantindo maior eficiência e precisão.

### 2. Objetivo

O objetivo principal é automatizar o processo de atualização de uma planilha 
Excel com dados de status de pagamento coletados a partir de um website 
específico. O script deve:

* Ler dados de CPFs da planilha Excel.
* Acessar o website e consultar o status de pagamento para cada CPF.
* Coletar dados relevantes (status, data de pagamento, método de pagamento).
* Inserir os dados coletados na planilha Excel nas colunas correspondentes.

### 3. Descrição do Processo

O script Python irá executar as seguintes etapas:

1. **Leitura da Planilha:**
   * Carregar a planilha Excel "python-status_pagamento.xlsx" utilizando a biblioteca `openpyxl`.
   * Iterar pelas linhas da planilha, buscando os CPFs na coluna C.
   * Armazenar os CPFs em uma lista.

2. **Coleta de Dados da Web:**
   * Abrir o website "https://consultcpf-devaprender.netlify.app/" utilizando a biblioteca `selenium`.
   * Iterar pela lista de CPFs:
     * Inserir cada CPF no campo de pesquisa do website.
     * Simular o clique no botão de pesquisa.
     * Extrair os dados de status, data de pagamento e método de pagamento do website.
     * Armazenar os dados coletados em um dicionário, onde a chave é o CPF e o valor é um dicionário com as informações coletadas.

3. **Atualização da Planilha:**
   * Verificar se a coluna E possui o cabeçalho "Status" e, caso contrário, inseri-lo, juntamente com os cabeçalhos "Data de Pagamento" (coluna F) e "Método de Pagamento" (coluna G).
   * Iterar pelas linhas da planilha e, para cada CPF encontrado, inserir os dados coletados nas colunas E, F e G.

4. **Gravação da Planilha:**
   * Salvar as alterações na planilha Excel.

### 4. Tecnologias

* **Python:** Linguagem de programação utilizada para o desenvolvimento do script.
* **openpyxl:** Biblioteca Python para manipulação de planilhas Excel.
* **selenium:** Biblioteca Python para automatizar ações em navegadores web.
* **re:** Biblioteca Python para trabalhar com expressões regulares (para limpeza dos dados de CPF).
* **time:** Biblioteca Python para utilizar funções de tempo (para espera e controle de fluxo).

### 5. Considerações

* O script assume que a planilha Excel possui um cabeçalho na primeira linha.
* O script espera que os dados de CPF na planilha estejam na coluna C (terceira coluna).
* O script espera que os elementos do website sejam acessíveis através dos seus IDs ou XPATHs especificados.
* O script utiliza tempos de espera para permitir que o website carregue completamente e os elementos sejam renderizados.
* O script foi testado utilizando o navegador Chrome. Outros navegadores podem requerer configurações adicionais.

### 6. Restrições

* O script depende do funcionamento do website "https://consultcpf-devaprender.netlify.app/".
* O script pode ser afetado por alterações no layout ou estrutura do website.
* O script pode ser afetado por alterações no layout ou estrutura da planilha.

### 7. Próximos Passos

* Implementar o script Python de acordo com a descrição técnica.
* Testar o script com dados reais para garantir a funcionalidade e precisão.
* Ajustar o script, se necessário, para atender a quaisquer requisitos adicionais.

### 8. Versão do Documento

* **Versão:** 1.0
* **Data de Revisão:** 20/06/2024

### 9. Conclusão

A automatização da atualização da planilha Excel com dados da web utilizando o 
script Python proporcionará uma solução eficiente e precisa para o processo 
manual de consulta e atualização de informações. O script economizará tempo e 
esforço, além de reduzir a possibilidade de erros humanos.
