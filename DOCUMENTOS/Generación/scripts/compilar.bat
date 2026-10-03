@echo off
chcp 65001 > nul
echo ==========================================================
echo  Compilando Proyecto LaTeX - GRESLY / UNSCH
echo ==========================================================
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0compilar.ps1" %*
pause
