# ============================================================
#  TrustIT - Printer Monitoring Installer
#  Baixa os arquivos crus direto do GitHub para C:\TrustIT
# ============================================================

$ErrorActionPreference = "Stop"
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

$destRoot = "C:\TrustIT"

$arquivos = @(
    @{
        Url    = "https://raw.githubusercontent.com/gabs-art/printer-monitoring/main/printers.ps1"
        Destino = Join-Path $destRoot "printers.ps1"
    },
    @{
        Url    = "https://raw.githubusercontent.com/gabs-art/printer-monitoring/main/register-task.ps1"
        Destino = Join-Path $destRoot "register-task.ps1"
    }
)

Write-Host ""
Write-Host "================================================" -ForegroundColor Cyan
Write-Host "   TrustIT - Printer Monitoring Setup" -ForegroundColor Cyan
Write-Host "================================================" -ForegroundColor Cyan
Write-Host ""

# ── 1. Criar pasta C:\TrustIT se nao existir ────────────────
if (-not (Test-Path $destRoot)) {
    Write-Host "[1/2] Criando pasta $destRoot..." -ForegroundColor Yellow
    New-Item -ItemType Directory -Path $destRoot -Force | Out-Null
    Write-Host "      Pasta criada com sucesso." -ForegroundColor Green
} else {
    Write-Host "[1/2] Pasta $destRoot ja existe." -ForegroundColor Gray
}

# ── 2. Baixar cada arquivo diretamente do GitHub ────────────
Write-Host "[2/2] Baixando arquivos do GitHub..." -ForegroundColor Yellow

foreach ($arquivo in $arquivos) {
    $nome = Split-Path $arquivo.Destino -Leaf
    Write-Host "      -> $nome" -ForegroundColor Yellow

    try {
        Invoke-WebRequest -Uri $arquivo.Url -OutFile $arquivo.Destino -UseBasicParsing

        $info = Get-Item $arquivo.Destino -ErrorAction Stop
        if ($info.Length -lt 10) {
            throw "Arquivo muito pequeno ($($info.Length) bytes). Verifique se o arquivo existe no repositorio."
        }

        Write-Host "         OK ($([math]::Round($info.Length/1KB, 1)) KB)" -ForegroundColor Green
    } catch {
        Write-Host "         ERRO ao baixar $nome`: $_" -ForegroundColor Red
        exit 1
    }
}

Write-Host ""
Write-Host "================================================" -ForegroundColor Cyan
Write-Host "   Instalacao concluida!" -ForegroundColor Green
Write-Host "   Arquivos em: $destRoot" -ForegroundColor White
Write-Host "================================================" -ForegroundColor Cyan
Write-Host ""

# ── 3. Executar register-task.ps1 para registrar a tarefa agendada ──
$registerScript = Join-Path $destRoot "register-task.ps1"

Write-Host "[3/3] Registrando tarefa agendada..." -ForegroundColor Yellow
try {
    & powershell.exe -NonInteractive -NoProfile -ExecutionPolicy Bypass -File $registerScript
    if ($LASTEXITCODE -ne 0) {
        throw "register-task.ps1 encerrou com codigo $LASTEXITCODE"
    }
} catch {
    Write-Host "ERRO ao registrar a tarefa: $_" -ForegroundColor Red
    exit 1
}
