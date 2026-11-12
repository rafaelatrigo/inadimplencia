@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo ==== Reiniciando o servico do painel de inadimplencia ====
set "ME=%~dp0scripts\servico_local.py"
powershell -NoProfile -Command "Get-CimInstance Win32_Process -Filter \"Name like 'python%%'\" | Where-Object { $_.CommandLine -like ('*' + $env:ME + '*') } | ForEach-Object { Stop-Process -Id $_.ProcessId -Force }"
python "%~dp0scripts\instalar_servico.py"
if errorlevel 1 (echo. & echo ERRO ao reiniciar. Mande um print desta tela. & pause & exit /b 1)
echo.
echo Pronto! Servico reiniciado com a versao nova.
pause
