@echo off
chcp 65001 >nul
cd /d "%~dp0"
del "%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup\PainelInadimplencia.vbs" 2>nul
set "ME=%~dp0scripts\servico_local.py"
powershell -NoProfile -Command "Get-CimInstance Win32_Process -Filter \"Name like 'python%%'\" | Where-Object { $_.CommandLine -like ('*' + $env:ME + '*') } | ForEach-Object { Stop-Process -Id $_.ProcessId -Force }"
echo Servico do painel de inadimplencia parado e removido da inicializacao do Windows.
echo (O servico do painel de despesas nao e afetado.)
pause
