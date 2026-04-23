[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory = $false)]
    [string]$CsvPath = "$env:USERPROFILE\Desktop\import_membros.csv",

    [Parameter(Mandatory = $false)]
    [string]$OutDir = (Join-Path $env:USERPROFILE "Desktop\Import-Logs")
)

$ErrorActionPreference = 'Stop'

function New-LogObject {
    param(
        [string]$Group,
        [string]$User,
        [string]$Status,
        [string]$Error
    )

    [pscustomobject]@{
        Timestamp = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
        Group     = $Group
        User      = $User
        Status    = $Status
        Error     = $Error
    }
}

try {
    if (-not (Test-Path -Path $CsvPath)) {
        throw "Arquivo CSV não encontrado em: $CsvPath"
    }

    if (-not (Get-Command Add-DistributionGroupMember -ErrorAction SilentlyContinue)) {
        throw "O cmdlet Add-DistributionGroupMember não está disponível na sessão atual. Conecte-se ao Exchange antes de executar o script."
    }

    New-Item -ItemType Directory -Path $OutDir -Force | Out-Null

    $OkLog   = Join-Path $OutDir "DG_ok.csv"
    $FailLog = Join-Path $OutDir "DG_falhas.csv"

    $ok   = New-Object System.Collections.Generic.List[object]
    $fail = New-Object System.Collections.Generic.List[object]

    $rows = Import-Csv -Path $CsvPath

    if (-not $rows) {
        throw "O CSV está vazio ou não contém dados válidos."
    }

    foreach ($r in $rows) {
        $group = "$($r.Group)".Trim()
        $usersRaw = "$($r.Users)".Trim()

        if ([string]::IsNullOrWhiteSpace($group)) {
            $fail.Add((New-LogObject -Group "" -User "" -Status "Falha" -Error "Group vazio no CSV")) | Out-Null
            continue
        }

        if ([string]::IsNullOrWhiteSpace($usersRaw)) {
            $fail.Add((New-LogObject -Group $group -User "" -Status "Falha" -Error "Users vazio no CSV")) | Out-Null
            continue
        }

        $users = $usersRaw -split ';' |
            ForEach-Object { $_.Trim().ToLower() } |
            Where-Object { $_ -ne "" } |
            Sort-Object -Unique

        foreach ($u in $users) {
            try {
                if ($PSCmdlet.ShouldProcess("Grupo '$group'", "Adicionar membro '$u'")) {
                    Add-DistributionGroupMember -Identity $group -Member $u -ErrorAction Stop
                }

                $ok.Add((New-LogObject -Group $group -User $u -Status "Adicionado" -Error "")) | Out-Null
                Write-Host "OK: $u -> $group"
            }
            catch {
                $fail.Add((New-LogObject -Group $group -User $u -Status "Falha" -Error $_.Exception.Message)) | Out-Null
                Write-Host "FALHA: $u -> $group | $($_.Exception.Message)"
            }
        }
    }

    $ok   | Export-Csv -Path $OkLog -NoTypeInformation -Encoding UTF8
    $fail | Export-Csv -Path $FailLog -NoTypeInformation -Encoding UTF8

    Write-Host ""
    Write-Host "Resumo da execução:"
    Write-Host "Total OK: $($ok.Count)"
    Write-Host "Total Falhas: $($fail.Count)"
    Write-Host "Logs em: $OutDir"
}
catch {
    Write-Error $_.Exception.Message
    exit 1
}
