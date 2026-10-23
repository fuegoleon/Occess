$ErrorActionPreference = "Continue"

Write-Host "========================================================================" -ForegroundColor Cyan
Write-Host "       SUBIR CAMBIOS Y ACTUALIZAR PORTABLE DE OCCESS (GITHUB)           " -ForegroundColor Yellow
Write-Host "========================================================================" -ForegroundColor Cyan
Write-Host ""

# 1. Configurar rutas de Node y Git
$nodeDir = "C:\Program Files\nodejs"
if (-not (Test-Path "$nodeDir\npm.cmd")) {
    $npmCmd = (Get-Command npm.cmd -ErrorAction SilentlyContinue)
    if ($npmCmd) { $nodeDir = Split-Path $npmCmd.Source }
}
$env:PATH = "$nodeDir;" + $env:PATH

$gitCmd = (Get-Command git.exe -ErrorAction SilentlyContinue)
if (-not $gitCmd) {
    if (Test-Path "C:\Program Files\Git\cmd\git.exe") {
        $env:PATH = "C:\Program Files\Git\cmd;" + $env:PATH
    }
}

# 2. Verificar estado de Git
Write-Host "[1/5] Verificando cambios locales en Git..." -ForegroundColor Gray
$gitStatus = & git status --porcelain
if (-not $gitStatus) {
    Write-Host "[INFO] No se detectaron archivos modificados pendientes." -ForegroundColor Yellow
    $continuar = Read-Host "¿Deseas forzar la publicacion/compilacion de todos modos? (S/N) [N]"
    if ($continuar -notmatch '^[sSyY]') {
        Write-Host "Operacion cancelada por el usuario." -ForegroundColor Gray
        exit 0
    }
} else {
    Write-Host "[OK] Se detectaron archivos modificados:" -ForegroundColor Green
    & git status --short
    Write-Host ""
}

# 3. Leer versión actual y calcular siguiente versión
$pkg = Get-Content "package.json" -Raw | ConvertFrom-Json
$currVer = $pkg.version
$parts = $currVer.Split('.')
if ($parts.Length -eq 3) {
    $nextVerAuto = "$($parts[0]).$($parts[1]).$([int]$parts[2] + 1)"
} else {
    $nextVerAuto = "1.0.6"
}

Write-Host "------------------------------------------------------------------------" -ForegroundColor Cyan
Write-Host " Version actual: " -NoNewline
Write-Host "v$currVer" -ForegroundColor Yellow
Write-Host " Siguiente version recomendada: " -NoNewline
Write-Host "v$nextVerAuto" -ForegroundColor Green
Write-Host "------------------------------------------------------------------------" -ForegroundColor Cyan
Write-Host ""

$inputVer = Read-Host "Ingresa la version a publicar (Presiona ENTER para v$nextVerAuto)"
$finalVer = if ([string]::IsNullOrWhiteSpace($inputVer)) { $nextVerAuto } else { $inputVer.Trim().TrimStart('v') }

$defaultMsg = "feat: actualizacion de sistema v$finalVer"
$inputMsg = Read-Host "Mensaje del cambio (Presiona ENTER para '$defaultMsg')"
$finalMsg = if ([string]::IsNullOrWhiteSpace($inputMsg)) { $defaultMsg } else { $inputMsg.Trim() }

Write-Host ""
Write-Host "¿Que accion deseas realizar?" -ForegroundColor Cyan
Write-Host " [1] Subir a GitHub Y PUBLICAR ejecutable portable en GitHub Releases (Recomendado)" -ForegroundColor Yellow
Write-Host " [2] Solo subir codigo a GitHub (sin compilar instaladores)" -ForegroundColor Gray
$opcion = Read-Host "Elige una opcion (1 o 2) [Presiona ENTER para 1]"
if ([string]::IsNullOrWhiteSpace($opcion)) { $opcion = "1" }

# 4. Actualizar números de versión en archivos
Write-Host ""
Write-Host "[2/5] Actualizando version a v$finalVer en archivos..." -ForegroundColor Gray

# package.json
$pkgContent = Get-Content "package.json" -Raw
$pkgContent = $pkgContent -replace '"version":\s*"[^"]+"', "`"version`": `"$finalVer`""
[System.IO.File]::WriteAllText("$PWD\package.json", $pkgContent, [System.Text.Encoding]::UTF8)

# index.html
if (Test-Path "index.html") {
    $indexContent = Get-Content "index.html" -Raw
    $indexContent = $indexContent -replace 'Occess v\d+\.\d+\.\d+', "Occess v$finalVer"
    $indexContent = $indexContent -replace '>v\d+\.\d+\.\d+<', ">v$finalVer<"
    [System.IO.File]::WriteAllText("$PWD\index.html", $indexContent, [System.Text.Encoding]::UTF8)
}

Write-Host "[OK] package.json e index.html actualizados a v$finalVer" -ForegroundColor Green

# 5. Git commit y push
Write-Host ""
Write-Host "[3/5] Guardando cambios y subiendo a GitHub (git push)..." -ForegroundColor Gray
& git add -A
& git commit -m "$finalMsg"
& git push origin main

if ($LASTEXITCODE -ne 0) {
    Write-Host "[AVISO] Hubo un aviso al subir a git o no habia cambios nuevos para pushear." -ForegroundColor Yellow
} else {
    Write-Host "[OK] Codigo subido correctamente a GitHub (rama main)." -ForegroundColor Green
}

# Si eligió solo git push, terminamos aquí
if ($opcion -eq "2") {
    Write-Host ""
    Write-Host "========================================================================" -ForegroundColor Green
    Write-Host " [EXITO] Cambios guardados y subidos a GitHub correctamente." -ForegroundColor Green
    Write-Host "========================================================================" -ForegroundColor Green
    exit 0
}

# 6. Compilar y publicar Release en GitHub
Write-Host ""
Write-Host "[4/5] Preparando compilacion y publicacion en GitHub Releases..." -ForegroundColor Gray

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
    exit 1
}

# Configurar herramientas de winCodeSign
$cacheDir = "$env:LOCALAPPDATA\electron-builder\Cache\winCodeSign"
$targetWinCodeSign = "$cacheDir\winCodeSign-2.6.0"
if (-not (Test-Path "$targetWinCodeSign\windows-10")) {
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

# Limpieza de instaladores anteriores en dist
Remove-Item "dist\Occess-Setup-*.exe", "dist\*.blockmap" -Force -ErrorAction SilentlyContinue

Write-Host "[5/5] Compilando ejecutables y subiendo a GitHub Releases (v$finalVer)..." -ForegroundColor Yellow
Write-Host "Esto tomara aproximadamente 1 minuto..." -ForegroundColor Gray
Write-Host ""

& "$nodeDir\npm.cmd" run release

if ($LASTEXITCODE -eq 0) {
    # Publicar automáticamente sacando de borrador (draft: false)
    $currTag = "v$finalVer"
    Write-Host "[*] Publicando $currTag como lanzamiento visible..." -ForegroundColor Cyan
    try {
        $headers = @{
            Authorization = "Bearer $env:GH_TOKEN"
            "User-Agent" = "Occess-App"
            "Accept" = "application/vnd.github+json"
        }
        $releases = Invoke-RestMethod -Uri "https://api.github.com/repos/fuegoleon/Occess/releases" -Headers $headers
        foreach ($r in $releases) {
            if ($r.tag_name -eq $currTag -and $r.draft) {
                $body = @{ draft = $false; make_latest = "true" } | ConvertTo-Json
                Invoke-RestMethod -Uri "https://api.github.com/repos/fuegoleon/Occess/releases/$($r.id)" -Method Patch -Headers $headers -Body $body -ContentType "application/json" | Out-Null
            }
        }
    } catch {
        Write-Host "[AVISO] Nota sobre estado draft: $($_.Exception.Message)" -ForegroundColor Yellow
    }

    Write-Host ""
    Write-Host "========================================================================" -ForegroundColor Green
    Write-Host " [EXITO] ¡Version v$finalVer publicada exitosamente en GitHub Releases!" -ForegroundColor Green
    Write-Host " El ejecutable portable detectara la actualizacion automaticamente." -ForegroundColor Green
    Write-Host " URL: https://github.com/fuegoleon/Occess/releases" -ForegroundColor Green
    Write-Host "========================================================================" -ForegroundColor Green
} else {
    Write-Host ""
    Write-Host "[ERROR] Hubo un problema al compilar o publicar. Codigo de salida: $LASTEXITCODE" -ForegroundColor Red
}
