@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo ==== Instalando o servico do painel de inadimplencia ====
set "PYW="
for /f "delims=" %%P in ('where pythonw 2^>nul') do if not defined PYW set "PYW=%%P"
if not defined PYW (echo Nao encontrei o pythonw. Reinstale o Python marcando "Add python.exe to PATH". & pause & exit /b 1)
set "STARTUP=%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup"
> "%STARTUP%\PainelInadimplencia.vbs" echo CreateObject("WScript.Shell").Run """%PYW%"" ""%~dp0scripts\servico_local.py""", 0, False
start "" wscript "%STARTUP%\PainelInadimplencia.vbs"
echo.
echo Pronto! O servico esta rodando em segundo plano e vai iniciar sozinho junto com o Windows.
echo Abra o painel neste computador: aparecem o selo "modo edicao" e o botao "Publicar agora".
echo Registro de tudo que ele faz: atualizacao.log
pause
