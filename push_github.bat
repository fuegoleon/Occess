@echo off
title Subiendo Occess a GitHub...
echo ====================================================
echo      CONECTANDO Y SUBIENDO CODIGO A GITHUB
echo ====================================================
echo.
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
