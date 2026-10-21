@echo off
chcp 65001 >nul
title Filtrado Inteligente de Fecha Anterior (Corte 9:00 AM)
cls
echo ========================================================================
echo       FILTRADO INTELIGENTE DE FECHA ANTERIOR (CORTE 9:00 AM)
echo ========================================================================
echo.
echo [1/3] Evaluando dia actual y reglas de negocio...
echo       - Si es Lunes: Jala Viernes, Sabado y Domingo (3 dias antes)
echo       - Si es Martes a Viernes: Jala 1 dia antes
echo       - Si es Sabado o Domingo: No jala informacion
echo.
echo [2/3] Procesando Pruebas2026.xlsx...
echo       - Asignando Fecha Registro: Fecha de hoy a las 09:00:00
echo       - Asignando Mes Registro: Mes actual
echo       - Asignando Hora Registro: 09:00:00
echo.

set ARCHIVO_ORIGEN=%~1
if "%ARCHIVO_ORIGEN%"=="" set ARCHIVO_ORIGEN=Pruebas2026.xlsx

powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0filtrar_fecha_anterior.ps1" "%ARCHIVO_ORIGEN%"
if %ERRORLEVEL% NEQ 0 (
    echo.
    echo [ERROR] Ocurrio un error al filtrar el archivo.
    pause
    exit /b %ERRORLEVEL%
)

echo.
echo [3/3] Actualizando datos en el Sistema Web y abriendo navegador...
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0actualizar_datos.ps1" "Pruebas2026_Filtrado_Lunes.xlsx"
start "" "%~dp0index.html"

echo.
echo ========================================================================
echo [EXITO] Proceso completado exitosamente.
echo - Archivo generado: Pruebas2026_Filtrado_Lunes.xlsx
echo - Sistema web actualizado y abierto en su navegador con el nuevo boton.
echo ========================================================================
timeout /t 4 >nul
