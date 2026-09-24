@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo ============================================ >> log_atualizacao.txt
echo %date% %time% - ATUALIZACAO AUTOMATICA >> log_atualizacao.txt

python gerar_catalogo.py >> log_atualizacao.txt 2>&1
if errorlevel 1 (
    echo ERRO no gerar_catalogo.py >> log_atualizacao.txt
    exit /b 1
)

git add catalogo.csv catalogo_top500.csv >> log_atualizacao.txt 2>&1
git diff --cached --quiet
if errorlevel 1 (
    git commit -m "atualiza catalogo automatico %date%" >> log_atualizacao.txt 2>&1
    git push >> log_atualizacao.txt 2>&1
    if errorlevel 1 (
        echo ERRO no git push >> log_atualizacao.txt
        exit /b 1
    )
    echo OK - catalogo publicado >> log_atualizacao.txt
) else (
    echo Nenhuma mudanca de preco hoje. >> log_atualizacao.txt
)
echo. >> log_atualizacao.txt
