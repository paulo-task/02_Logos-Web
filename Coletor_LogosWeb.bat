@echo off
title Executor de Scripts NS - Engelmig
cls

echo =======================================================
echo   INICIANDO PROCESSAMENTO DE NOTAS DE SERVICO (NS)
echo =======================================================
echo.

:: 1. Executando o Script de Paulista
echo [1/2] Rodando Script PAULISTA...
python "C:\Users\paulo.janio\ENGELMIG ENERGIA LTDA\LEC ENGELMIG - Workspace\03 Repository\02_Logos-Web\NS_LoWeb_PAULISTA.py"
if %errorlevel% neq 0 (
    echo.
    echo [AVISO] O script de Paulista encontrou um erro ou atingiu o limite de tentativas.
) else (
    echo [OK] Script Paulista finalizado com sucesso.
)

echo.
echo -------------------------------------------------------
echo.

:: 2. Executando o Script de Piratininga
echo [2/2] Rodando Script PIRATININGA...
python "C:\Users\paulo.janio\ENGELMIG ENERGIA LTDA\LEC ENGELMIG - Workspace\03 Repository\02_Logos-Web\NS_LoWeb_PIRATININGA.py"
if %errorlevel% neq 0 (
    echo.
    echo [AVISO] O script de Piratininga encontrou um erro ou atingiu o limite de tentativas.
) else (
    echo [OK] Script Piratininga finalizado com sucesso.
)

echo.
echo =======================================================
echo   TODOS OS PROCESSOS FORAM CONCLUIDOS!
echo =======================================================
pause