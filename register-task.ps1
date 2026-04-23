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

$action = New-ScheduledTaskAction `
    -Execute "powershell.exe" `
    -Argument "-NonInteractive -NoProfile -ExecutionPolicy Bypass -File `"$ScriptPrinters`""

$trigger = New-ScheduledTaskTrigger `
    -RepetitionInterval (New-TimeSpan -Minutes 30) `
    -Once -At (Get-Date)

$settings = New-ScheduledTaskSettingsSet `
    -ExecutionTimeLimit (New-TimeSpan -Minutes 25) `
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
    -Description "Monitor de filas de impressao - executa a cada 30 minutos" `
    -Force | Out-Null

Write-Host "Tarefa '$TaskName' registrada com sucesso!" -ForegroundColor Green
Write-Host "Script : $ScriptPrinters" -ForegroundColor Cyan
Write-Host "Intervalo: a cada 30 minutos" -ForegroundColor Cyan
Write-Host "Conta : SYSTEM (privilégio total)" -ForegroundColor Cyan
