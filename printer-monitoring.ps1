#Requires -RunAsAdministrator
<#
.SYNOPSIS
    Monitor de impressoras - varredura a cada 5 minutos.
    Limpa apenas jobs com erro, preserva jobs aguardando impressao.
.NOTES
    Requer execucao como Administrador no servidor de impressao.
    Para registrar como Tarefa Agendada, execute com o parametro -Registrar.
    Exemplo: .\printers.ps1 -Registrar
#>

param(
    [switch]$Registrar
)

# ─── Configuracoes ────────────────────────────────────────────────────────────
$Config = @{
    LogDir          = "C:\Logs\PrinterMonitor"
    SpoolDir        = "$env:SystemRoot\System32\spool\PRINTERS"
    IntervalSeconds = 300
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

function Get-JobsComErro {
    param([string]$NomeImpressora)

    $jobs = Get-PrintJob -PrinterName $NomeImpressora -ErrorAction SilentlyContinue
    if (-not $jobs) { return @() }

    return @($jobs | Where-Object {
        $_.JobStatus -match "Error|UserIntervention|Offline|PaperOut|BlockedDeviceQueue|Paused|Restart"
    })
}

function Test-JobAindaExiste {
    param([string]$NomeImpressora, [int]$JobId)

    $jobs = Get-PrintJob -PrinterName $NomeImpressora -ErrorAction SilentlyContinue
    return $null -ne ($jobs | Where-Object { $_.Id -eq $JobId })
}

# Tenta remover jobs com erro via cmdlet.
# Retorna os IDs dos jobs que continuaram travados apos a tentativa.
function Clear-FilaComErro {
    param([string]$NomeImpressora)

    $jobsErro = Get-JobsComErro -NomeImpressora $NomeImpressora
    if ($jobsErro.Count -eq 0) { return @() }

    Write-Log "  $($jobsErro.Count) job(s) com erro detectado(s) em '$NomeImpressora'" "AVISO"

    $idsTravados = [System.Collections.Generic.List[int]]::new()

    foreach ($job in $jobsErro) {
        Write-Log ("  Tentando remover: ID={0} | Doc='{1}' | Usuario='{2}' | Status='{3}'" -f `
            $job.Id, $job.Document, $job.UserName, $job.JobStatus) "AVISO"

        Remove-PrintJob -PrinterName $NomeImpressora -Id $job.Id -ErrorAction SilentlyContinue
        Start-Sleep -Milliseconds 800

        if (Test-JobAindaExiste -NomeImpressora $NomeImpressora -JobId $job.Id) {
            Write-Log "  Job ID=$($job.Id) continua travado no spool. Sera removido forcadamente." "AVISO"
            $idsTravados.Add($job.Id)
        } else {
            Write-Log "  Job ID=$($job.Id) removido com sucesso."
        }
    }

    return $idsTravados.ToArray()
}

function Invoke-VarredurasImpressoras {
    Write-Log "════════════════════════════════════════"
    Write-Log "Iniciando varredura de impressoras..."

    try {
        $impressoras = Get-Printer -ErrorAction Stop
    } catch {
        Write-Log "Nao foi possivel listar impressoras: $($_.Exception.Message)" "ERRO"
        return @()
    }

    Write-Log "Total de impressoras: $($impressoras.Count)"

    $todosIdsTravados = [System.Collections.Generic.List[int]]::new()
    $totalOffline     = 0
    $totalSemErro     = 0

    foreach ($impressora in $impressoras) {
        $status = $impressora.PrinterStatus.ToString()

        if ($status -eq "Offline") {
            $totalOffline++
            Write-Log "OFFLINE: '$($impressora.Name)'" "AVISO"
        } else {
            Write-Log "Verificando: '$($impressora.Name)' | Status: $status"
        }

        $idsTravados = Clear-FilaComErro -NomeImpressora $impressora.Name

        if ($idsTravados.Count -eq 0) {
            $totalSemErro++
        } else {
            foreach ($id in $idsTravados) { $todosIdsTravados.Add($id) }
        }
    }

    Write-Log ("Varredura concluida. OK: {0} | Offline: {1} | Jobs travados (forcado): {2}" -f `
        $totalSemErro, $totalOffline, $todosIdsTravados.Count)

    return $todosIdsTravados.ToArray()
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

# Remove apenas os arquivos .SPL e .SHD dos job IDs informados.
# Jobs saudaveis (aguardando impressao) nao sao tocados.
function Remove-SpoolFilesPorId {
    param([int[]]$JobIds)

    foreach ($id in $JobIds) {
        $baseName = "{0:D8}" -f $id
        foreach ($ext in @("SPL", "SHD")) {
            $arquivo = Join-Path $Config.SpoolDir "$baseName.$ext"
            if (Test-Path $arquivo) {
                Remove-Item -Path $arquivo -Force -ErrorAction SilentlyContinue
                if (Test-Path $arquivo) {
                    Write-Log "  Nao foi possivel excluir: $baseName.$ext" "ERRO"
                } else {
                    Write-Log "  Arquivo de spool excluido: $baseName.$ext"
                }
            }
        }
    }
}

function Restart-PrintSpooler {
    param([int[]]$JobsTravados = @())

    Write-Log "Parando servico Spooler..."
    try {
        Stop-Service -Name Spooler -Force -ErrorAction Stop

        if (-not (Wait-ServiceStatus -Nome Spooler -Status Stopped)) {
            Write-Log "Spooler nao parou dentro do tempo esperado." "ERRO"
            return
        }

        if ($JobsTravados.Count -gt 0) {
            Write-Log "Removendo arquivos de spool apenas dos $($JobsTravados.Count) job(s) travado(s)..." "AVISO"
            Remove-SpoolFilesPorId -JobIds $JobsTravados
        } else {
            Write-Log "Nenhum arquivo de spool para remover. Jobs aguardando impressao preservados."
        }

        Write-Log "Iniciando servico Spooler..."
        Start-Service -Name Spooler -ErrorAction Stop

        if (Wait-ServiceStatus -Nome Spooler -Status Running) {
            Write-Log "Spooler reiniciado com sucesso."
        } else {
            Write-Log "Spooler nao subiu dentro do tempo esperado." "AVISO"
        }
    } catch {
        Write-Log "Falha ao reiniciar Spooler: $($_.Exception.Message)" "ERRO"
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
    Write-Host "O script sera executado automaticamente a cada 5 minutos pelo Task Scheduler.`n" -ForegroundColor Green
}

# ─── Entrypoint ───────────────────────────────────────────────────────────────

Initialize-LogDir

if ($Registrar) {
    Register-TarefaAgendada
    exit 0
}

Remove-OldLogs
$idsTravados = Invoke-VarredurasImpressoras
Restart-PrintSpooler -JobsTravados $idsTravados
Write-Log "Ciclo concluido. Proxima execucao em $($Config.IntervalSeconds / 60) minuto(s)."
