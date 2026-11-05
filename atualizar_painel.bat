@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo ==== Painel de inadimplência ====
echo.
echo 1/3 Gerando dados (relatórios do Sienge mais recentes da pasta entrada)...
python scripts\gerar_dados.py
if errorlevel 1 (echo. & echo ERRO ao gerar os dados. Nada foi publicado. & pause & exit /b 1)
echo.
echo 2/3 Registrando no Git...
git pull --quiet --rebase origin main 2>nul
git add -A
git diff --cached --quiet
if not errorlevel 1 (echo Sem mudancas nos dados. Nada a publicar. & pause & exit /b 0)
git commit -m "Atualiza dados do painel %date%"
echo.
echo 3/3 Enviando para o GitHub (o GitHub Pages publica sozinho em 1-2 min)...
git push -u origin HEAD:main
if errorlevel 1 (echo ERRO no envio para o GitHub. & pause & exit /b 1)
echo.
echo Pronto! Painel atualizado.
pause
