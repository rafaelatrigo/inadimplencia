@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo ==== Instalando o servico do painel de inadimplencia ====
echo.
python "%~dp0scripts\instalar_servico.py"
if errorlevel 1 (echo. & echo ERRO na instalacao do servico. Mande um print desta tela. & pause & exit /b 1)
echo.
echo Pronto! O servico esta rodando e vai iniciar sozinho junto com o Windows.
echo Abra o painel neste computador e aperte F5: aparecem o selo "modo edicao" e o botao "Publicar agora".
pause
