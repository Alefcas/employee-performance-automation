@echo off
chcp 65001 >nul
title TOP AMBASSADORS - FC locale (DEMO / PORTFOLIO)
cd /d "%~dp0"

set "LAUNCHER=2. RECURSOS\TOP AMBASSADORS.py"

REM Abre o painel de boas-vindas e depois o processo completo roda no terminal.
REM Por isso usamos python (com console), e nao pythonw.
where python >nul 2>nul
if %errorlevel%==0 (
    python "%LAUNCHER%"
    goto fim
)

where py >nul 2>nul
if %errorlevel%==0 (
    py "%LAUNCHER%"
    goto fim
)

REM Python nao encontrado
echo.
echo  ============================================================
echo    Python nao foi encontrado neste computador.
echo    Instale o Python em: https://www.python.org/downloads/
echo    (marque a opcao "Add Python to PATH" na instalacao)
echo  ============================================================
echo.
pause

:fim
