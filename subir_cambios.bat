@echo off
title Subir Cambios y Publicar Occess
cd /d "%~dp0"
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0subir_cambios.ps1"
echo.
pause
