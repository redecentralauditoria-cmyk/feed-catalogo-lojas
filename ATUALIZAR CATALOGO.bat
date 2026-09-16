@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo ============================================
echo   ATUALIZAR CATALOGO DO WHATSAPP
echo ============================================
echo.

python gerar_catalogo.py
if errorlevel 1 goto erro

echo Enviando para a nuvem...
git add catalogo.csv
git diff --cached --quiet && (echo Nenhuma mudanca de preco hoje. Nada a enviar. & goto fim)
git commit -m "atualiza catalogo %date%"
git push
if errorlevel 1 goto erro

echo.
echo PRONTO. A Meta vai reler o catalogo no proximo horario programado (09:26).
goto fim

:erro
echo.
echo *** DEU ERRO. Tire um print desta tela e mande para o Claude. ***

:fim
echo.
pause
