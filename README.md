# Automatização de atualização de planilha com dados da web

Este projeto utiliza Python para automatizar a atualização de uma planilha Excel com dados de status de pagamento coletados de um website.

## Funcionalidades:

- Lê dados de CPFs da planilha Excel.
- Acessa o website e consulta o status de pagamento para cada CPF.
- Coleta dados relevantes (status, data de pagamento, método de pagamento).
- Insere os dados coletados na planilha Excel nas colunas correspondentes.

## Como usar:

1. **Instale as dependências:**
   ```bash
   pip install openpyxl selenium
   ```
2. **Renomeie a planilha "python-status_pagamento.xlsx" com os seus dados.**
3. **Execute o script:**
   ```bash
   python script.py
   ```

## Detalhes:

- A planilha deve conter um cabeçalho na primeira linha.
- Os CPFs devem estar na coluna C da planilha.
- O script utiliza o website `https://consultcpf-devaprender.netlify.app/` para consultar os dados.
- Ajuste o código, se necessário, para outras configurações do website ou da planilha.

## Versão:

- Versão 1.0

## Próximos Passos:

- Implementar tratamento de erros e validações.
- Criar interface gráfica para facilitar o uso.
- Adaptar para outros websites.

## Contribuições:

Contribuições são bem-vindas! Abra um issue para reportar problemas ou sugestões.
<!-- 
## Licença:

Este projeto está sob a licença [MIT](LICENSE). -->
