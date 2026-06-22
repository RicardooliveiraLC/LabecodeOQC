@echo off
chcp 65001 >nul
title Atualizar Dashboard OQC - Label Code
cd /d "%~dp0"

echo ========================================================
echo   Atualizando o Dashboard OQC a partir da planilha...
echo ========================================================
echo.

REM tenta 'python', depois 'py' como alternativa
python gerar_dashboard.py
if errorlevel 1 (
    py gerar_dashboard.py
    if errorlevel 1 (
        echo.
        echo [ERRO] Python nao encontrado ou falha ao gerar.
        echo Verifique se o Python esta instalado e no PATH.
        pause
        exit /b 1
    )
)

echo.
echo Abrindo o dashboard no navegador...
start "" "dashboard.html"
timeout /t 3 >nul
