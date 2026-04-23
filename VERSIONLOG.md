# VERSIONLOG

## v0.2

### Novidades

- Adicionado suporte a parâmetros `-CsvPath` e `-OutDir`
- Adicionada validação de existência do arquivo CSV
- Adicionada verificação da disponibilidade do cmdlet `Add-DistributionGroupMember`
- Implementado suporte a `-WhatIf`
- Adicionado resumo final com totais de sucesso e falha
- Adicionado `timestamp` nos logs
- Melhorada a estrutura geral do script
- Melhorada a documentação para uso em GitHub

### Observações

Esta versão mantém a proposta original do script, mas torna o uso mais seguro, mais flexível e mais adequado para repositórios técnicos e ambientes administrativos.

## v0.1

### Base inicial

- Leitura de arquivo CSV fixo na área de trabalho
- Criação automática da pasta de logs
- Inclusão de membros via `Add-DistributionGroupMember`
- Separação de sucessos e falhas em arquivos CSV
- Remoção de duplicados por linha do CSV
