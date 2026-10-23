$ErrorActionPreference = "Continue"
$nodeDir = "C:\Program Files\nodejs"
if (-not (Test-Path "$nodeDir\npm.cmd")) {
    $npmCmd = (Get-Command npm.cmd -ErrorAction SilentlyContinue)
    if ($npmCmd) { $nodeDir = Split-Path $npmCmd.Source }
}
$env:PATH = "$nodeDir;" + $env:PATH

Write-Host "========================================================================" -ForegroundColor Cyan
Write-Host "       COMPILADOR PORTABLE DE SISTEMA OCCESS CON ELECTRON" -ForegroundColor Yellow
Write-Host "========================================================================" -ForegroundColor Cyan
Write-Host ""

# 1. Configurar cache de winCodeSign para evitar fallo de symlink en Windows
$cacheDir = "$env:LOCALAPPDATA\electron-builder\Cache\winCodeSign"
$targetWinCodeSign = "$cacheDir\winCodeSign-2.6.0"
if (-not (Test-Path "$targetWinCodeSign\windows-10")) {
    Write-Host "[1/3] Configurando herramientas de empaquetado Windows..." -ForegroundColor Gray
    New-Item -ItemType Directory -Path $targetWinCodeSign -Force | Out-Null
    $first7z = Get-ChildItem -Path $cacheDir -Filter "*.7z" -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($first7z) {
        $7za = "node_modules\7zip-bin\win\x64\7za.exe"
        if (Test-Path $7za) {
            & $7za x -y $first7z.FullName "-o$targetWinCodeSign" | Out-Null
        }
    }
}

# 2. Configurar variables de entorno para compilar portable sin certificados remotos
$env:CSC_IDENTITY_AUTO_DISCOVERY = "false"
$env:ELECTRON_BUILDER_ALLOW_UNRESOLVED_DEPENDENCIES = "true"

Write-Host "[2/3] Empaquetando aplicacion a ejecutable Portable (.exe)..." -ForegroundColor Gray
& "$nodeDir\npm.cmd" run build:portable

Write-Host ""
$distExe = "dist\Occess.exe"
if (Test-Path $distExe) {
    $sizeMB = [Math]::Round(((Get-Item $distExe).Length / 1MB), 2)
    Write-Host "========================================================================" -ForegroundColor Green
    Write-Host " [EXITO] ¡Ejecutable portable Occess generado correctamente!" -ForegroundColor Green
    Write-Host " Ruta: $distExe ($sizeMB MB)" -ForegroundColor Green
    Write-Host "========================================================================" -ForegroundColor Green
} else {
    Write-Host "[INFO] Verificando carpeta dist..." -ForegroundColor Yellow
    Get-ChildItem "dist" | Format-Table Name, Length, LastWriteTime
}
