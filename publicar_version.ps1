$ErrorActionPreference = "Continue"

Write-Host "========================================================================" -ForegroundColor Cyan
Write-Host "       PUBLICADOR DE VERSIONES DE OCCESS EN GITHUB RELEASES" -ForegroundColor Yellow
Write-Host "========================================================================" -ForegroundColor Cyan
Write-Host ""

$nodeDir = "C:\Program Files\nodejs"
if (-not (Test-Path "$nodeDir\npm.cmd")) {
    $npmCmd = (Get-Command npm.cmd -ErrorAction SilentlyContinue)
    if ($npmCmd) { $nodeDir = Split-Path $npmCmd.Source }
}
$env:PATH = "$nodeDir;" + $env:PATH

# Cargar GH_TOKEN desde github_token.txt si existe
if (Test-Path "github_token.txt") {
    $token = (Get-Content "github_token.txt").Trim()
    if ($token) {
        $env:GH_TOKEN = $token
        Write-Host "[OK] Token de GitHub cargado correctamente." -ForegroundColor Green
    }
}

if (-not $env:GH_TOKEN) {
    Write-Host "[ERROR] No se encontro GH_TOKEN en github_token.txt ni en las variables de entorno." -ForegroundColor Red
    pause
    exit 1
}

# 1. Configurar cache de winCodeSign para evitar fallo de symlink en Windows
$cacheDir = "$env:LOCALAPPDATA\electron-builder\Cache\winCodeSign"
$targetWinCodeSign = "$cacheDir\winCodeSign-2.6.0"
if (-not (Test-Path "$targetWinCodeSign\windows-10")) {
    Write-Host "[1/2] Configurando herramientas de empaquetado Windows..." -ForegroundColor Gray
    New-Item -ItemType Directory -Path $targetWinCodeSign -Force | Out-Null
    $first7z = Get-ChildItem -Path $cacheDir -Filter "*.7z" -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($first7z) {
        $7za = "node_modules\7zip-bin\win\x64\7za.exe"
        if (Test-Path $7za) {
            & $7za x -y $first7z.FullName "-o$targetWinCodeSign" | Out-Null
        }
    }
}

$env:CSC_IDENTITY_AUTO_DISCOVERY = "false"
$env:ELECTRON_BUILDER_ALLOW_UNRESOLVED_DEPENDENCIES = "true"

# Limpieza de instaladores anteriores para ahorrar espacio en disco
Write-Host "[*] Limpiando instaladores anteriores en dist para ahorrar espacio..." -ForegroundColor Gray
Remove-Item "dist\Occess-Setup-*.exe", "dist\*.blockmap" -Force -ErrorAction SilentlyContinue

Write-Host "[2/2] Compilando y subiendo a GitHub Releases (version $( (Get-Content package.json | ConvertFrom-Json).version ))..." -ForegroundColor Yellow
Write-Host ""

& "$nodeDir\npm.cmd" run release

if ($LASTEXITCODE -eq 0) {
    # Publicar automáticamente sacando de borrador (draft: false)
    $currVersion = "v" + ((Get-Content package.json | ConvertFrom-Json).version)
    Write-Host "[3/3] Publicando version $currVersion como lanzamiento oficial visible..." -ForegroundColor Cyan
    try {
        $headers = @{
            Authorization = "Bearer $env:GH_TOKEN"
            "User-Agent" = "Occess-App"
            "Accept" = "application/vnd.github+json"
        }
        $releases = Invoke-RestMethod -Uri "https://api.github.com/repos/fuegoleon/Occess/releases" -Headers $headers
        foreach ($r in $releases) {
            if ($r.tag_name -eq $currVersion -and $r.draft) {
                $body = @{ draft = $false; make_latest = "true" } | ConvertTo-Json
                Invoke-RestMethod -Uri "https://api.github.com/repos/fuegoleon/Occess/releases/$($r.id)" -Method Patch -Headers $headers -Body $body -ContentType "application/json" | Out-Null
            }
        }
    } catch {
        Write-Host "[AVISO] No se pudo cambiar el estado de draft automaticamente: $($_.Exception.Message)" -ForegroundColor Yellow
    }

    Write-Host ""
    Write-Host "========================================================================" -ForegroundColor Green
    Write-Host " [EXITO] ¡Version publicada exitosamente en GitHub Releases!" -ForegroundColor Green
    Write-Host " URL: https://github.com/fuegoleon/Occess/releases" -ForegroundColor Green
    Write-Host "========================================================================" -ForegroundColor Green
} else {
    Write-Host ""
    Write-Host "[ERROR] Hubo un problema al compilar o publicar. Codigo de salida: $LASTEXITCODE" -ForegroundColor Red
}
