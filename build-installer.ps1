# =============================================================================
# CSV.Convertor — скрипт сборки инсталлятора
# Выполняет: dotnet publish → компиляция Inno Setup → копирование в _zip
# =============================================================================

param(
    [string]$Configuration = "Release",
    [string]$InnoSetupPath = ""
)

$ErrorActionPreference = "Stop"
$ProjectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$PublishDir  = Join-Path $ProjectRoot "bin\$Configuration\net10.0-windows\win-x64\publish"
$InstallerDir = Join-Path $ProjectRoot "installer"
$OutputDir   = Join-Path $InstallerDir "Output"
$ZipDir      = Join-Path $ProjectRoot "_zip"
$SetupIss    = Join-Path $InstallerDir "setup.iss"

Write-Host "========================================" -ForegroundColor Cyan
Write-Host " CSV.Convertor — сборка инсталлятора" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# ---------------------------------------------------------------------------
# Шаг 1: dotnet publish (single-file, self-contained)
# ---------------------------------------------------------------------------
Write-Host "[1/3] Публикация приложения (dotnet publish -c $Configuration)..." -ForegroundColor Yellow
Push-Location $ProjectRoot
try {
    # Сначала восстанавливаем пакеты
    Write-Host "  Восстановление пакетов..." -ForegroundColor Gray
    dotnet restore 2>&1 | Out-Host
    if ($LASTEXITCODE -ne 0) {
        throw "Ошибка dotnet restore (код $LASTEXITCODE)"
    }

    # Публикация single-file
    Write-Host "  Публикация single-file..." -ForegroundColor Gray
    dotnet publish -c $Configuration --no-restore 2>&1 | Out-Host
    if ($LASTEXITCODE -ne 0) {
        throw "Ошибка dotnet publish (код $LASTEXITCODE)"
    }
    Write-Host "  Готово. Выходная папка: $PublishDir" -ForegroundColor Green
} finally {
    Pop-Location
}

# Проверяем, что exe существует
$ExePath = Join-Path $PublishDir "CSV.Convertor.exe"
if (-not (Test-Path $ExePath)) {
    throw "Не найден $ExePath после публикации"
}
Write-Host "  Размер exe: $([math]::Round((Get-Item $ExePath).Length / 1MB, 1)) МБ" -ForegroundColor Gray
Write-Host ""

# ---------------------------------------------------------------------------
# Шаг 2: Компиляция Inno Setup
# ---------------------------------------------------------------------------
Write-Host "[2/3] Компиляция инсталлятора (Inno Setup)..." -ForegroundColor Yellow

# Если путь не указан явно — ищем автоматически
if (-not $InnoSetupPath) {
    $InnoSetupPath = "${env:ProgramFiles(x86)}\Inno Setup 6\ISCC.exe"
}

if (-not (Test-Path $InnoSetupPath)) {
    Write-Host "  Inno Setup не найден по пути: $InnoSetupPath" -ForegroundColor Red
    Write-Host "  Пытаюсь найти через реестр..." -ForegroundColor Yellow
    
    # Поиск через реестр
    $regPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\Inno Setup 6_is1"
    if (Test-Path $regPath) {
        $installLocation = (Get-ItemProperty $regPath).InstallLocation
        if ($installLocation) {
            $InnoSetupPath = Join-Path $installLocation "ISCC.exe"
        }
    }
}

if (-not (Test-Path $InnoSetupPath)) {
    throw "Не удалось найти Inno Setup (ISCC.exe).`nСкачайте: https://jrsoftware.org/isinfo.php"
}

Write-Host "  Компилятор: $InnoSetupPath" -ForegroundColor Gray
$issDir = Split-Path $SetupIss -Parent
Push-Location $issDir
try {
    & $InnoSetupPath $SetupIss 2>&1 | Out-Host
    if ($LASTEXITCODE -ne 0) {
        throw "Ошибка компиляции Inno Setup (код $LASTEXITCODE)"
    }
} finally {
    Pop-Location
}

$SetupExe = Join-Path $OutputDir "CSV.Convertor.Setup.exe"
if (-not (Test-Path $SetupExe)) {
    throw "Не найден $SetupExe после компиляции"
}
Write-Host "  Готово: $SetupExe" -ForegroundColor Green
Write-Host ""

# ---------------------------------------------------------------------------
# Шаг 3: Копирование в _zip
# ---------------------------------------------------------------------------
Write-Host "[3/3] Копирование инсталлятора в _zip ..." -ForegroundColor Yellow
New-Item -ItemType Directory -Force -Path $ZipDir | Out-Null
Copy-Item -Force $SetupExe -Destination (Join-Path $ZipDir "CSV.Convertor.Setup.exe")
Write-Host "  Скопировано в: $ZipDir\CSV.Convertor.Setup.exe" -ForegroundColor Green
Write-Host ""

# ---------------------------------------------------------------------------
Write-Host "========================================" -ForegroundColor Cyan
Write-Host " Инсталлятор успешно собран!" -ForegroundColor Green
Write-Host " $(Join-Path $ZipDir 'CSV.Convertor.Setup.exe')" -ForegroundColor White
Write-Host "========================================" -ForegroundColor Cyan
