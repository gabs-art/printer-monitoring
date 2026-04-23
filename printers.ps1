#Requires -RunAsAdministrator
<#
.SYNOPSIS
    Monitor de impressoras - replica exatamente o fluxo manual:
    "Abrir fila > botao direito > Cancelar > Sim"
    Remove jobs expirados (erro ou impressora offline) via API Win32,
    com fallback de exclusao fisica dos arquivos .SPL/.SHD.
.NOTES
    Para registrar como Tarefa Agendada: .\printers.ps1 -Registrar
#>

param(
    [switch]$Registrar,
    [int]$MinutosParaExpirar = 1440   # 1440 min = 24h. Apenas jobs com mais de 1 dia serao removidos.
)

# ─── Configuracoes ────────────────────────────────────────────────────────────
$Config = @{
    LogDir          = "C:\Logs\PrinterMonitor"
    SpoolDir        = "$env:SystemRoot\System32\spool\PRINTERS"
    TaskName        = "PrinterMonitor"
    TaskDescription = "Monitor de filas de impressao - cancela jobs expirados e reinicia Spooler"
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
    if ($ts.TotalDays -ge 1)   { return "$([int]$ts.TotalDays)d $($ts.Hours)h $($ts.Minutes)min" }
    if ($ts.TotalHours -ge 1)  { return "$([int]$ts.TotalHours)h $($ts.Minutes)min" }
    return "$([int]$ts.TotalMinutes)min"
}

# ─── API Win32 - mesmo caminho que o botao "Cancelar" da UI usa ───────────────
Add-Type @"
using System;
using System.Runtime.InteropServices;
public class WinPrint {
    [DllImport("winspool.drv", CharSet=CharSet.Unicode, SetLastError=true)]
    public static extern bool OpenPrinter(string pPrinterName, out IntPtr phPrinter, IntPtr pDefault);

    [DllImport("winspool.drv", SetLastError=true)]
    public static extern bool SetJob(IntPtr hPrinter, int JobId, int Level, IntPtr pJob, int Command);

    [DllImport("winspool.drv", SetLastError=true)]
    public static extern bool ClosePrinter(IntPtr hPrinter);

    public const int JOB_CONTROL_DELETE = 5;
}
"@
# ─────────────────────────────────────────────────────────────────────────────

function Invoke-CancelarJobAPI {
    param(
        [string]$NomeImpressora,
        [int]$JobId,
        [int]$Comando = 5   # 5 = JOB_CONTROL_DELETE | 3 = JOB_CONTROL_CANCEL
    )

    $hPrinter = [IntPtr]::Zero
    try {
        if ([WinPrint]::OpenPrinter($NomeImpressora, [ref]$hPrinter, [IntPtr]::Zero)) {
            $ok = [WinPrint]::SetJob($hPrinter, $JobId, 0, [IntPtr]::Zero, $Comando)
            [WinPrint]::ClosePrinter($hPrinter) | Out-Null
            return $ok
        }
    } catch {
        if ($hPrinter -ne [IntPtr]::Zero) {
            [WinPrint]::ClosePrinter($hPrinter) | Out-Null
        }
    }
    return $false
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

# Retorna todos os jobs parados ha mais de $MinutosParaExpirar minutos.
# Impressao deve acontecer na hora - qualquer job acima do limite esta travado.
function Get-JobsExpirados {
    $limite    = (Get-Date).AddMinutes(-$MinutosParaExpirar)
    $resultado = [System.Collections.Generic.List[PSCustomObject]]::new()

    try {
        $impressoras = Get-Printer -ErrorAction Stop
    } catch {
        Write-Log "Nao foi possivel listar impressoras: $($_.Exception.Message)" "ERRO"
        return $resultado
    }

    Write-Log "Total de impressoras: $($impressoras.Count)"

    foreach ($impressora in $impressoras) {
        $statusImpressora = $impressora.PrinterStatus.ToString()
        Write-Log "Verificando: '$($impressora.Name)' | Status: $statusImpressora"

        $jobs = Get-PrintJob -PrinterName $impressora.Name -ErrorAction SilentlyContinue
        if (-not $jobs) { continue }

        foreach ($job in $jobs) {
            $idade = (Get-Date) - $job.SubmittedTime

            if ($job.SubmittedTime -ge $limite) {
                Write-Log ("  Recente ({0}): ID={1} | '{2}' | Status='{3}'" -f `
                    (Format-Idade -ts $idade), $job.Id, $job.Document, $job.JobStatus)
                continue
            }

            Write-Log ("  EXPIRADO [{0}]: ID={1} | Doc='{2}' | Usuario='{3}' | Status='{4}' | Fila='{5}'" -f `
                (Format-Idade -ts $idade), $job.Id, $job.Document, $job.UserName, $job.JobStatus, $impressora.Name) "AVISO"

            $resultado.Add([PSCustomObject]@{
                Id         = [int]$job.Id
                Impressora = $impressora.Name
                Documento  = $job.Document
                Idade      = Format-Idade -ts $idade
            })
        }
    }

    return $resultado
}

# Tenta cancelar cada job via API Win32 (mesmo caminho do botao "Cancelar" da UI).
# Verifica se o job realmente saiu da fila apos a chamada.
# Retorna lista dos que continuam na fila para tratamento forcado.
function Invoke-CancelarJobs {
    param([System.Collections.Generic.List[PSCustomObject]]$Jobs)

    $naoRemovidos = [System.Collections.Generic.List[PSCustomObject]]::new()

    foreach ($job in $Jobs) {
        Write-Log ("  Cancelando via API: ID={0} | '{1}' | Fila='{2}' | {3}" -f `
            $job.Id, $job.Documento, $job.Impressora, $job.Idade) "AVISO"

        # Tenta DELETE; se falhar, tenta CANCEL antes de DELETE (necessario para status Complete)
        $ok = Invoke-CancelarJobAPI -NomeImpressora $job.Impressora -JobId $job.Id
        if (-not $ok) {
            # JOB_CONTROL_CANCEL = 3, depois DELETE = 5
            Invoke-CancelarJobAPI -NomeImpressora $job.Impressora -JobId $job.Id -Comando 3 | Out-Null
            Start-Sleep -Milliseconds 300
            $ok = Invoke-CancelarJobAPI -NomeImpressora $job.Impressora -JobId $job.Id
        }

        # Aguarda o Spooler processar e confirma se o job realmente saiu da fila
        Start-Sleep -Milliseconds 800
        $aindaExiste = Get-PrintJob -PrinterName $job.Impressora -ErrorAction SilentlyContinue |
                       Where-Object { $_.Id -eq $job.Id }

        if (-not $aindaExiste) {
            Write-Log "  OK: Job ID=$($job.Id) removido da fila com sucesso."
        } else {
            Write-Log "  Job ID=$($job.Id) ainda na fila apos API. Encaminhando para remocao fisica." "AVISO"
            $naoRemovidos.Add($job)
        }
    }

    return $naoRemovidos
}

# Fallback: para o Spooler, localiza os arquivos .SHD/.SPL das impressoras
# com jobs travados lendo o cabecalho binario do SHD, e os exclui.
function Invoke-RemocaoFisica {
    param([System.Collections.Generic.List[PSCustomObject]]$Jobs)

    if ($Jobs.Count -eq 0) { return }

    $impressorasAfetadas = @($Jobs | Select-Object -ExpandProperty Impressora -Unique)

    Write-Log "Parando Spooler para remocao fisica de $($Jobs.Count) job(s) travado(s)..." "AVISO"
    try {
        Stop-Service -Name Spooler -Force -ErrorAction Stop
        if (-not (Wait-ServiceStatus -Nome Spooler -Status Stopped)) {
            Write-Log "Spooler nao parou no tempo esperado." "ERRO"
            return
        }
        Write-Log "Spooler parado."

        $shdFiles = Get-ChildItem -Path $Config.SpoolDir -Filter "*.SHD" -ErrorAction SilentlyContinue
        $removidos = 0

        foreach ($shd in $shdFiles) {
            try {
                # Le o arquivo binario e converte para Unicode para encontrar o nome da impressora
                $bytes   = [System.IO.File]::ReadAllBytes($shd.FullName)
                $conteudo = [System.Text.Encoding]::Unicode.GetString($bytes)

                foreach ($impressora in $impressorasAfetadas) {
                    if ($conteudo -like "*$impressora*") {
                        $spl = [System.IO.Path]::ChangeExtension($shd.FullName, ".SPL")

                        Remove-Item -Path $shd.FullName -Force -ErrorAction SilentlyContinue
                        if (Test-Path $spl) {
                            Remove-Item -Path $spl -Force -ErrorAction SilentlyContinue
                        }

                        Write-Log ("  Fisico removido: {0} | Impressora='{1}'" -f `
                            [System.IO.Path]::GetFileNameWithoutExtension($shd.Name), $impressora)
                        $removidos++
                        break
                    }
                }
            } catch {
                # arquivo pode estar em uso por outro processo, ignora
            }
        }

        Write-Log "$removidos arquivo(s) de spool removido(s) fisicamente."
    } catch {
        Write-Log "Erro critico durante remocao fisica: $($_.Exception.Message)" "ERRO"
    } finally {
        Write-Log "Iniciando Spooler..."
        Start-Service -Name Spooler -ErrorAction SilentlyContinue
        if (Wait-ServiceStatus -Nome Spooler -Status Running) {
            Write-Log "Spooler reiniciado com sucesso."
        } else {
            Write-Log "Spooler nao subiu no tempo esperado." "AVISO"
        }
    }
}

# Reinicio preventivo do Spooler quando nao ha jobs travados
function Invoke-RestartSpooler {
    Write-Log "Reiniciando Spooler preventivamente..."
    try {
        Stop-Service -Name Spooler -Force -ErrorAction Stop
        if (Wait-ServiceStatus -Nome Spooler -Status Stopped) {
            Start-Service -Name Spooler -ErrorAction Stop
            if (Wait-ServiceStatus -Nome Spooler -Status Running) {
                Write-Log "Spooler reiniciado com sucesso."
            } else {
                Write-Log "Spooler nao subiu no tempo esperado." "AVISO"
            }
        } else {
            Write-Log "Spooler nao parou no tempo esperado." "ERRO"
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
$limiteLabel = if ($MinutosParaExpirar -ge 1440) { "$([int]($MinutosParaExpirar/1440))d ($MinutosParaExpirar min)" } else { "$MinutosParaExpirar min" }
Write-Log "Iniciando ciclo | Remove jobs parados ha mais de: $limiteLabel"

$jobsExpirados = Get-JobsExpirados
Write-Log "Jobs expirados encontrados: $($jobsExpirados.Count)"

if ($jobsExpirados.Count -gt 0) {
    # Passo 1: tenta cancelar via API Win32 (igual ao botao Cancelar da UI)
    $jobsTravados = Invoke-CancelarJobs -Jobs $jobsExpirados

    # Passo 2: para os que a API nao conseguiu, remove fisicamente os arquivos de spool
    Invoke-RemocaoFisica -Jobs $jobsTravados
} else {
    # Sem jobs expirados: apenas reinicia o Spooler preventivamente
    Invoke-RestartSpooler
}

Write-Log "Ciclo concluido."
