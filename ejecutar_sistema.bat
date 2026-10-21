@echo off
chcp 65001 >nul
title Sistema de Gestion y Filtrado de Servicios (SAR)
cls

echo ========================================================================
echo       SISTEMA DE GESTION Y VINCULACION DE TABLAS EXCEL (SAR)
echo ========================================================================
echo.
echo [1/2] Procesando y vinculando los archivos Excel...
if "%~1"=="" (
    echo       - Archivo Base: Base a Cargar prueba.xlsx (por defecto)
) else (
    echo       - Archivo Base especificado: %~nx1
)
echo       - Tablas vinculadas auxiliares (Tiempos, Agrupamiento, Reguladores, etc.)
echo.

powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0actualizar_datos.ps1" "%~1"
if %ERRORLEVEL% NEQ 0 (
    echo.
    echo [ERROR] Ocurrio un error al leer o vincular los archivos Excel.
    echo Asegurese de que los archivos no esten bloqueados o danados.
    echo Presione cualquier tecla para salir...
    pause >nul
    exit /b %ERRORLEVEL%
)

echo.
echo [2/2] Abriendo sistema interactivo en su navegador web...
start "" "%~dp0index.html"

echo.
echo ========================================================================
echo [EXITO] El sistema esta listo y abierto en su navegador.
echo Puede usar los filtros tipo Excel por columna, buscar y exportar.
echo ========================================================================
timeout /t 3 >nul
