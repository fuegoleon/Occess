@echo off
chcp 65001 >nul
title Compilador Portable - Sistema SAR (Electron)
echo Iniciando proceso de compilacion portable...
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0build_portable.ps1"
pause
