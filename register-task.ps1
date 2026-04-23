    #Requires -RunAsAdministrator

    # Resolve o diretorio do script com fallbacks para execucao via console ou ISE
    $scriptDir = $PSScriptRoot
    if (-not $scriptDir) { $scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path }
    if (-not $scriptDir) { $scriptDir = "C:\TrustIt" }

    $ScriptPrinters = Join-Path $scriptDir "printers.ps1"

    if (-not (Test-Path $ScriptPrinters)) {
        Write-Host "Arquivo nao encontrado: $ScriptPrinters" -ForegroundColor Red
        Write-Host "Certifique-se de que printers.ps1 esta na mesma pasta que este script." -ForegroundColor Yellow
        exit 1
    }

    $TaskName = "PrinterMonitor"

    # Remove a tarefa anterior se existir, evitando duplicatas
    if (Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue) {
        Write-Host "Tarefa existente encontrada. Removendo antes de recriar..." -ForegroundColor DarkYellow
        Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false
    }

    $action = New-ScheduledTaskAction `
        -Execute "powershell.exe" `
        -Argument "-NonInteractive -NoProfile -ExecutionPolicy Bypass -File `"$ScriptPrinters`""

    $trigger = New-ScheduledTaskTrigger `
        -RepetitionInterval (New-TimeSpan -Hours 2) `
        -Once -At (Get-Date)

    $settings = New-ScheduledTaskSettingsSet `
        -ExecutionTimeLimit (New-TimeSpan -Minutes 30) `
        -RestartCount 2 `
        -RestartInterval (New-TimeSpan -Minutes 5) `
        -StartWhenAvailable

    $principal = New-ScheduledTaskPrincipal -UserId "SYSTEM" -RunLevel Highest

    Register-ScheduledTask `
        -TaskName    $TaskName `
        -Action      $action `
        -Trigger     $trigger `
        -Settings    $settings `
        -Principal   $principal `
        -Description "Monitor de filas de impressao - executa a cada 2 horas" `
        -Force | Out-Null

    Write-Host "Tarefa '$TaskName' registrada com sucesso!" -ForegroundColor Green
    Write-Host "Script : $ScriptPrinters" -ForegroundColor Cyan
    Write-Host "Intervalo: a cada 2 horas" -ForegroundColor Cyan
    Write-Host "Conta : SYSTEM (privilégio total)" -ForegroundColor Cyan
