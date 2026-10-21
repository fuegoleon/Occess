@echo off
title Publicador Automatico de Occess a GitHub Releases
setlocal enabledelayedexpansion

echo ========================================================================
echo       PUBLICADOR DE VERSIONES DE OCCESS EN GITHUB RELEASES
echo ========================================================================
echo.

set "nodeDir=C:\Users\ext_rleon\AppData\Local\Microsoft\WinGet\Packages\OpenJS.NodeJS.LTS_Microsoft.Winget.Source_8wekyb3d8bbwe\node-v24.20.0-win-x64"
set "PATH=%nodeDir%;%PATH%"

REM Verificar si existe archivo con el token o variable de entorno
if "%GH_TOKEN%"=="" (
    if exist "github_token.txt" (
        set /p GH_TOKEN=<github_token.txt
    )
)

if "%GH_TOKEN%"=="" (
    echo [AVISO] Se necesita un GitHub Personal Access Token para publicar.
    echo Si no lo tienes, puedes crearlo en: https://github.com/settings/tokens/new
    echo (Solo marca la casilla de 'repo').
    echo.
    set /p GH_TOKEN="Pega tu GH_TOKEN aqui: "
    echo !GH_TOKEN!> github_token.txt
    echo Token guardado localmente en github_token.txt (protegido por .gitignore).
    echo.
)

REM Configurar cache de winCodeSign para evitar fallo de symlink en Windows
powershell -NoProfile -Command "$cacheDir = \"$env:LOCALAPPDATA\electron-builder\Cache\winCodeSign\"; $targetWinCodeSign = \"$cacheDir\winCodeSign-2.6.0\"; if (-not (Test-Path \"$targetWinCodeSign\windows-10\")) { New-Item -ItemType Directory -Path $targetWinCodeSign -Force | Out-Null; $first7z = Get-ChildItem -Path $cacheDir -Filter \"*.7z\" -ErrorAction SilentlyContinue | Select-Object -First 1; if ($first7z) { $7za = \"node_modules\7zip-bin\win\x64\7za.exe\"; if (Test-Path $7za) { & $7za x -y $first7z.FullName \"-o$targetWinCodeSign\" | Out-Null } } }"

set "CSC_IDENTITY_AUTO_DISCOVERY=false"
set "ELECTRON_BUILDER_ALLOW_UNRESOLVED_DEPENDENCIES=true"

echo [1/2] Compilando Occess (Instalador NSIS One-Click y Portable)...
echo [2/2] Subiendo version y latest.yml a GitHub Releases...
echo.

call npm run release

if %ERRORLEVEL% EQU 0 (
    echo.
    echo ========================================================================
    echo  [EXITO] ¡Version publicada exitosamente en GitHub Releases!
    echo  Los usuarios recibiran la notificacion de actualizacion automatica.
    echo  Revisalo en: https://github.com/fuegoleon/Occess/releases
    echo ========================================================================
) else (
    echo.
    echo [ERROR] Hubo un error al compilar o publicar. Revisa el mensaje arriba.
)

echo.
pause
