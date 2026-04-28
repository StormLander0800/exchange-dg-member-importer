<#
.SYNOPSIS
    Importa membros para Grupos de Distribuição do Exchange a partir de um CSV.
    Formato esperado do CSV: Group, Users (separados por ';')
#>
[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory = $false, HelpMessage = "Caminho para o arquivo CSV de entrada.")]
    [ValidateScript({ Test-Path $_ -PathType Leaf })]
    [string]$CsvPath = (Join-Path $env:USERPROFILE "Desktop\import_membros.csv"),

    [Parameter(Mandatory = $false, HelpMessage = "Diretório onde os logs serão salvos.")]
    [string]$OutDir = (Join-Path $env:USERPROFILE "Desktop\Import-Logs")
)

process {
    $ErrorActionPreference = 'Stop'
    
    # 1. Validações Iniciais
    if (-not (Get-Command Add-DistributionGroupMember -ErrorAction SilentlyContinue)) {
        Write-Error "Comando do Exchange não encontrado. Por favor, conecte-se ao módulo do Exchange Online."
        return
    }

    # Garantir criação do diretório de log
    if (-not (Test-Path $OutDir)) {
        New-Item -ItemType Directory -Path $OutDir -Force | Out-Null
    }

    # 2. Importação e Preparação
    try {
        $rows = Import-Csv -Path $CsvPath -Delimiter ',' # Ajuste o delimitador se necessário
        if (-not $rows) { throw "O CSV está vazio." }
    }
    catch {
        Write-Error "Erro ao ler CSV: $($_.Exception.Message)"
        return
    }

    $okLogs   = [System.Collections.Generic.List[pscustomobject]]::new()
    $failLogs = [System.Collections.Generic.List[pscustomobject]]::new()

    # 3. Processamento
    $totalRows = $rows.Count
    for ($i = 0; $i -lt $totalRows; $i++) {
        $r = $rows[$i]
        $group = $r.Group?.Trim()
        $usersRaw = $r.Users?.Trim()

        Write-Progress -Activity "Processando Grupos" -Status "Grupo: $group" -PercentComplete (($i / $totalRows) * 100)

        # Validação de campos vazios
        if ([string]::IsNullOrWhiteSpace($group) -or [string]::IsNullOrWhiteSpace($usersRaw)) {
            $failLogs.Add([pscustomobject]@{
                Timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
                Group     = $group
                User      = "N/A"
                Status    = "Falha"
                Error     = "Linha com campos obrigatórios vazios"
            })
            continue
        }

        # Tratamento da lista de usuários
        $users = $usersRaw -split ';' | ForEach-Object { $_.Trim() } | Where-Object { $_ } | Sort-Object -Unique

        foreach ($u in $users) {
            $logEntry = [pscustomobject]@{
                Timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
                Group     = $group
                User      = $u
                Status    = ""
                Error     = ""
            }

            try {
                if ($PSCmdlet.ShouldProcess("Grupo '$group'", "Adicionar membro '$u'")) {
                    Add-DistributionGroupMember -Identity $group -Member $u -ErrorAction Stop
                }
                
                $logEntry.Status = "Sucesso"
                $okLogs.Add($logEntry)
                Write-Host " [OK] $u -> $group" -ForegroundColor Green
            }
            catch {
                # Trata erro comum de membro já existente para não poluir o log de falhas críticas
                if ($_.Exception.Message -match "is already a member") {
                    $logEntry.Status = "Aviso"
                    $logEntry.Error  = "Usuário já pertence ao grupo"
                    $okLogs.Add($logEntry)
                    Write-Host " [!] $u já existe em $group" -ForegroundColor Yellow
                }
                else {
                    $logEntry.Status = "Falha"
                    $logEntry.Error  = $_.Exception.Message
                    $failLogs.Add($logEntry)
                    Write-Host " [X] $u -> $group | Erro: $($_.Exception.Message)" -ForegroundColor Red
                }
            }
        }
    }

    # 4. Exportação dos Logs
    $timestampFile = Get-Date -Format "yyyyMMdd_HHmmss"
    if ($okLogs.Count -gt 0) {
        $okLogs | Export-Csv -Path (Join-Path $OutDir "Sucesso_$timestampFile.csv") -NoTypeInformation -Encoding UTF8
    }
    if ($failLogs.Count -gt 0) {
        $failLogs | Export-Csv -Path (Join-Path $OutDir "Falhas_$timestampFile.csv") -NoTypeInformation -Encoding UTF8
    }

    # 5. Resumo Final
    Write-Host "`n" + ("=" * 30)
    Write-Host " RELATÓRIO FINAL" -ForegroundColor Cyan
    Write-Host " Sucessos/Avisos: $($okLogs.Count)"
    Write-Host " Falhas:          $($failLogs.Count)" -ForegroundColor Red
    Write-Host " Logs salvos em:  $OutDir"
    Write-Host ("=" * 30)
}
