@echo off
setlocal
title Inventario Legislativo
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0Iniciar_Inventario.ps1"
if errorlevel 1 pause
