@echo off
title Publicador Automatico de Occess a GitHub Releases
cd /d "%~dp0"
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0publicar_version.ps1"
echo.
pause
