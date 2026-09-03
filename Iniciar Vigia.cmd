@echo off
title Vigia PED - Publicacao automatica do dashboard
echo Iniciando o vigia da PED...
powershell -NoProfile -ExecutionPolicy Bypass -WindowStyle Minimized -File "%~dp0Vigiar_PED.ps1"
