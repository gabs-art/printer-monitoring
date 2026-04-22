#Requires -RunAsAdministrator
<#
.SYNOPSIS
    Monitor de impressoras - varredura a cada 5 minutos.
    Remove jobs com erro ou em impressoras com problema que estejam parados ha mais de 30 minutos.
.NOTES
    Para registrar como Tarefa Agendada: .\printers.ps1 -Registrar
#>

param(
    [switch]$Registrar,
    [int]$MinutosParaExpirar = 30   # Jobs parados ha mais desse tempo serao removidos
)

# ─── Configuracoes ────────────────────────────────────────────────────────────
$Config = @{
    LogDir          = "C:\Logs\PrinterMonitor"
    SpoolDir        = "$env:SystemRoot\System32\spool\PRINTERS"
    TaskName        = "PrinterMonitor"
    TaskDescription = "Monitor de filas de impressao - limpa erros e reinicia Spooler"
}
# ─────────────────────────────────────────────────────────────────────────────

$LogFile = Join-Path $Config.LogDir "printer_monitor_$(Get-Date -Format 'yyyy-MM').log"

function Write-Log {
    param(
        [string]$Mensagem,
        [ValidateSet("INFO","AVISO","ERRO")][string]$Nivel = "INFO"
    )
    $linha = "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] [$Nivel] $Mensagem"
    Write-Host $linha -ForegroundColor $(switch ($Nivel) {
        "AVISO" { "Yellow" }
        "ERRO"  { "Red"    }
        default { "Cyan"   }
    })
    Add-Content -Path $LogFile -Value $linha -ErrorAction SilentlyContinue
}

function Initialize-LogDir {
    if (-not (Test-Path $Config.LogDir)) {
        New-Item -ItemType Directory -Path $Config.LogDir -Force | Out-Null
    }
}

function Remove-OldLogs {
    Get-ChildItem -Path $Config.LogDir -Filter "*.log" -ErrorAction SilentlyContinue |
        Where-Object { $_.LastWriteTime -lt (Get-Date).AddDays(-90) } |
        Remove-Item -Force -ErrorAction SilentlyContinue
}

function Format-Idade {
    param([TimeSpan]$ts)
    if ($ts.TotalDays -ge 1) { return "$([int]$ts.TotalDays)d $($ts.Hours)h $($ts.Minutes)min" }
    if ($ts.TotalHours -ge 1) { return "$([int]$ts.TotalHours)h $($ts.Minutes)min" }
    return "$([int]$ts.TotalMinutes)min"
}

# Retorna lista de jobs que devem ser removidos:
# - Job com status de erro E parado ha mais de $MinutosParaExpirar minutos
# - OU impressora Offline/Erro E job parado ha mais de $MinutosParaExpirar minutos
function Get-JobsParaRemover {
    $limite     = (Get-Date).AddMinutes(-$MinutosParaExpirar)
    $resultado  = [System.Collections.Generic.List[PSCustomObject]]::new()

    try {
        $impressoras = Get-Printer -ErrorAction Stop
    } catch {
        Write-Log "Nao foi possivel listar impressoras: $($_.Exception.Message)" "ERRO"
        return $resultado
    }

    Write-Log "Total de impressoras encontradas: $($impressoras.Count)"

    foreach ($impressora in $impressoras) {
        $statusImpressora = $impressora.PrinterStatus.ToString()
        $impressoraComProblema = $statusImpressora -match "Offline|Error|PaperOut|UserIntervention|BlockedDeviceQueue"

        if ($impressoraComProblema) {
            Write-Log "PROBLEMA: '$($impressora.Name)' | Status: $statusImpressora" "AVISO"
        } else {
            Write-Log "Verificando: '$($impressora.Name)' | Status: $statusImpressora"
        }

        $jobs = Get-PrintJob -PrinterName $impressora.Name -ErrorAction SilentlyContinue
        if (-not $jobs) { continue }

        foreach ($job in $jobs) {
            $statusJob  = $job.JobStatus.ToString()
            $jobComErro = $statusJob -match "Error|UserIntervention|Offline|PaperOut|BlockedDeviceQueue|Paused|Restart"

            # Ignora jobs que estao ativamente imprimindo e saudaveis
            if (-not $jobComErro -and -not $impressoraComProblema) { continue }

            # Verifica se esta parado ha mais do tempo limite
            if ($job.SubmittedTime -ge $limite) {
                $idadeAtual = Format-Idade -ts ((Get-Date) - $job.SubmittedTime)
                Write-Log ("  Aguardando (ainda no prazo - $idadeAtual): ID={0} | '{1}'" -f $job.Id, $job.Document)
                continue
            }

            $idade = Format-Idade -ts ((Get-Date) - $job.SubmittedTime)
            Write-Log ("  EXPIRADO [{5}]: ID={0} | Doc='{1}' | Usuario='{2}' | JobStatus='{3}' | Fila='{4}'" -f `
                $job.Id, $job.Document, $job.UserName, $statusJob, $impressora.Name, $idade) "AVISO"

            $resultado.Add([PSCustomObject]@{
                Id         = $job.Id
                Impressora = $impressora.Name
                Documento  = $job.Document
                Idade      = $idade
            })
        }
    }

    return $resultado
}

function Wait-ServiceStatus {
    param([string]$Nome, [string]$Status, [int]$TimeoutSeg = 30)
    $elapsed = 0
    while ((Get-Service -Name $Nome).Status -ne $Status -and $elapsed -lt $TimeoutSeg) {
        Start-Sleep -Seconds 2
        $elapsed += 2
    }
    return (Get-Service -Name $Nome).Status -eq $Status
}

# Para o Spooler, remove os arquivos .SPL/.SHD de cada job informado e sobe o servico novamente.
function Invoke-LimpezaERestart {
    param([System.Collections.Generic.List[PSCustomObject]]$Jobs)

    Write-Log "Parando servico Spooler..."
    try {
        Stop-Service -Name Spooler -Force -ErrorAction Stop

        if (-not (Wait-ServiceStatus -Nome Spooler -Status Stopped)) {
            Write-Log "Spooler nao parou dentro do tempo esperado." "ERRO"
            return
        }
        Write-Log "Spooler parado."

        if ($Jobs.Count -gt 0) {
            Write-Log "Removendo $($Jobs.Count) arquivo(s) de spool expirado(s)..." "AVISO"

            foreach ($job in $Jobs) {
                $baseName = "{0:D8}" -f [int]$job.Id
                $removido = $false

                foreach ($ext in @("SPL", "SHD")) {
                    $arquivo = Join-Path $Config.SpoolDir "$baseName.$ext"
                    if (Test-Path $arquivo) {
                        Remove-Item -Path $arquivo -Force -ErrorAction SilentlyContinue
                        if (-not (Test-Path $arquivo)) {
                            Write-Log ("  OK: {0}.{1} | ID={2} | '{3}' | Fila='{4}'" -f `
                                $baseName, $ext, $job.Id, $job.Documento, $job.Impressora)
                            $removido = $true
                        } else {
                            Write-Log ("  FALHA ao excluir: {0}.{1}" -f $baseName, $ext) "ERRO"
                        }
                    }
                }

                if (-not $removido) {
                    Write-Log ("  Arquivos nao encontrados no spool para ID={0} (ja removidos ou ID divergente)" -f $job.Id) "AVISO"
                }
            }
        } else {
            Write-Log "Nenhum job expirado. Nenhum arquivo de spool removido."
        }

        Write-Log "Iniciando servico Spooler..."
        Start-Service -Name Spooler -ErrorAction Stop

        if (Wait-ServiceStatus -Nome Spooler -Status Running) {
            Write-Log "Spooler reiniciado com sucesso."
        } else {
            Write-Log "Spooler nao subiu dentro do tempo esperado." "AVISO"
        }
    } catch {
        Write-Log "Erro critico no ciclo do Spooler: $($_.Exception.Message)" "ERRO"
        Start-Service -Name Spooler -ErrorAction SilentlyContinue
    }
}

function Register-TarefaAgendada {
    Write-Host "`nRegistrando Tarefa Agendada '$($Config.TaskName)'..." -ForegroundColor Green

    $scriptPath = $MyInvocation.ScriptName
    if (-not $scriptPath) { $scriptPath = $PSCommandPath }

    $action  = New-ScheduledTaskAction -Execute "powershell.exe" `
                   -Argument "-NonInteractive -NoProfile -ExecutionPolicy Bypass -File `"$scriptPath`""

    $trigger = New-ScheduledTaskTrigger -RepetitionInterval (New-TimeSpan -Minutes 5) `
                   -Once -At (Get-Date)

    $settings = New-ScheduledTaskSettingsSet `
                    -ExecutionTimeLimit (New-TimeSpan -Minutes 4) `
                    -RestartCount 3 `
                    -RestartInterval (New-TimeSpan -Minutes 1) `
                    -StartWhenAvailable

    $principal = New-ScheduledTaskPrincipal -UserId "SYSTEM" -RunLevel Highest

    Register-ScheduledTask `
        -TaskName    $Config.TaskName `
        -Action      $action `
        -Trigger     $trigger `
        -Settings    $settings `
        -Principal   $principal `
        -Description $Config.TaskDescription `
        -Force | Out-Null

    Write-Host "Tarefa '$($Config.TaskName)' registrada com sucesso!" -ForegroundColor Green
    Write-Host "Execucao automatica a cada 5 minutos via Task Scheduler.`n" -ForegroundColor Green
}

# ─── Entrypoint ───────────────────────────────────────────────────────────────

Initialize-LogDir

if ($Registrar) {
    Register-TarefaAgendada
    exit 0
}

Remove-OldLogs

Write-Log "════════════════════════════════════════"
Write-Log "Iniciando ciclo | Limite de expiracao: $MinutosParaExpirar minutos"

$jobsParaRemover = Get-JobsParaRemover
Write-Log "Jobs expirados encontrados: $($jobsParaRemover.Count)"

Invoke-LimpezaERestart -Jobs $jobsParaRemover

Write-Log "Ciclo concluido."
