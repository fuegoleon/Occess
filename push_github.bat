@echo off
title Subir Cambios a GitHub - Occess
cd /d "%~dp0"
echo ====================================================
echo             SUBIR CAMBIOS A GITHUB
echo ====================================================
echo.
echo  [1] Subir cambios Y publicar actualizacion para el Portable (Recomendado)
echo  [2] Solo hacer git push simple
echo.
set /p Opc="Elige una opcion (1 o 2) [Presiona ENTER para 1]: "
if "%Opc%"=="" set Opc=1
if "%Opc%"=="1" (
    powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0subir_cambios.ps1"
    exit /b
)

echo.
echo Conectando y subiendo codigo a GitHub...
where git >nul 2>nul
if %ERRORLEVEL% EQU 0 (
    set "GIT_EXEC=git"
) else (
    set "GIT_EXEC=C:\Program Files\Git\cmd\git.exe"
)
"%GIT_EXEC%" push -u origin main
echo.
if %ERRORLEVEL% EQU 0 (
    echo ====================================================
    echo  [EXITO] Codigo subido correctamente a GitHub!
    echo ====================================================
) else (
    echo [AVISO] Revisa si se abrio una ventana de inicio de sesion de GitHub.
)
echo.
pause
