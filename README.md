# Exchange DG Member Importer

Script em PowerShell para importação em massa de membros em Grupos de Distribuição do Exchange via CSV, com geração de logs de sucesso e falha.

## Visão geral

Este projeto foi criado para facilitar a administração de grupos de distribuição no Exchange, permitindo adicionar vários usuários em lote a partir de um arquivo CSV.

A versão `0.2` melhora a base inicial com:

- validação do arquivo CSV antes da execução;
- verificação da disponibilidade do cmdlet `Add-DistributionGroupMember`;
- suporte a `-WhatIf` via `SupportsShouldProcess`;
- resumo final da execução;
- logs com `timestamp`;
- maior organização do código para uso em ambiente corporativo.

## Estrutura do projeto

```text
.
├── importar-membros-dg.ps1
├── importar-membros-dg-refatorado.ps1
├── secret
├── README.md
├── VERSIONLOG.md
├── .gitignore
└── examples
    └── import_membros_exemplo.csv
```

## Requisitos

- Windows PowerShell 5.1+ ou PowerShell 7+
- Permissões para gerenciar grupos no Exchange
- Sessão autenticada no Exchange Online ou ambiente compatível
- Cmdlet disponível:

```powershell
Add-DistributionGroupMember
```

## Formato do CSV

O arquivo precisa conter as colunas abaixo:

- `Group`
- `Users`

Exemplo:

```csv
Group,Users
grupoti@empresa.com,user1@empresa.com;user2@empresa.com;user3@empresa.com
gestores@empresa.com,gestor1@empresa.com;gestor2@empresa.com
```

### Regras do campo `Users`

- Separar múltiplos usuários com `;`
- Espaços extras são removidos automaticamente
- Usuários duplicados na mesma linha são eliminados
- Os valores são normalizados para minúsculas

## Como executar

### 1. Conecte-se ao Exchange

Exemplo para Exchange Online:

```powershell
Connect-ExchangeOnline
```

### 2. Execute o script com o caminho padrão

```powershell
.\importar-membros-dg.ps1
```

Por padrão, o script procura o arquivo em:

```powershell
$env:USERPROFILE\Desktop\import_membros.csv
```

E gera logs em:

```powershell
$env:USERPROFILE\Desktop\Import-Logs
```

### 3. Execute informando caminhos personalizados

```powershell
.\importar-membros-dg.ps1 -CsvPath "C:\Temp\import_membros.csv" -OutDir "C:\Temp\Logs"
```

### 4. Simule sem alterar nada

```powershell
.\importar-membros-dg.ps1 -WhatIf
```

## Saídas geradas

O script cria dois arquivos CSV de log:

- `DG_ok.csv` → usuários adicionados com sucesso
- `DG_falhas.csv` → falhas e mensagens de erro

Os logs possuem as colunas:

- `Timestamp`
- `Group`
- `User`
- `Status`
- `Error`

## Fluxo de funcionamento

1. Valida se o arquivo CSV existe
2. Verifica se o cmdlet do Exchange está disponível
3. Cria a pasta de logs
4. Importa o CSV
5. Valida `Group` e `Users`
6. Separa e higieniza os usuários
7. Remove duplicados
8. Executa a inclusão no grupo
9. Registra sucesso ou falha
10. Exporta os logs e exibe um resumo final

## Exemplo de saída no console

```text
OK: user1@empresa.com -> grupoti@empresa.com
OK: user2@empresa.com -> grupoti@empresa.com
FALHA: user3@empresa.com -> grupoti@empresa.com | The operation couldn't be performed because object wasn't found.

Resumo da execução:
Total OK: 2
Total Falhas: 1
Logs em: C:\Users\SeuUsuario\Desktop\Import-Logs
```

## Melhorias futuras

- validar previamente se o grupo existe;
- validar previamente se o usuário existe;
- verificar se o usuário já é membro do grupo antes de tentar adicionar;
- permitir leitura de delimitadores customizados;
- gerar relatório consolidado em HTML;
- adicionar parâmetros para nome dos arquivos de log.

## Boas práticas de uso

- Teste primeiro com `-WhatIf`
- Valide o CSV antes de rodar em produção
- Execute com conta que tenha permissão adequada
- Guarde os logs para auditoria e conferência

## Licença

Defina a licença conforme a política do seu ambiente, instituição ou repositório.
