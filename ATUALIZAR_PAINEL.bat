@echo off
setlocal EnableExtensions EnableDelayedExpansion
title ATUALIZAR PAINEL COMERCIAL

echo =====================================
echo Atualizando painel a partir do Excel
echo =====================================
echo.

REM Vai para a pasta do projeto (onde está o .bat)
cd /d "%~dp0"

REM ===============================
REM 1) Rodar Python principal
REM ===============================
echo [1/4] Gerando JSON (Python)...
python python\atualizar_painel_completo.py
set PY_ERR=%ERRORLEVEL%

IF NOT "%PY_ERR%"=="0" (
    echo.
    echo ❌ ERRO: Python falhou (codigo %PY_ERR%).
    echo Verifique a mensagem acima.
    echo.
    pause
    exit /b %PY_ERR%
)

echo.
echo ✅ Python finalizado com sucesso!
echo.

REM ===============================
REM 2) Mostrar status antes
REM ===============================
echo [2/4] Conferindo alteracoes (git status)...
git status
echo.

REM ===============================
REM 3) Git add
REM ===============================
echo [3/4] Executando git add...
git add .
set GIT_ADD_ERR=%ERRORLEVEL%

IF NOT "%GIT_ADD_ERR%"=="0" (
    echo.
    echo ❌ ERRO: git add falhou (codigo %GIT_ADD_ERR%).
    echo.
    pause
    exit /b %GIT_ADD_ERR%
)

REM ===============================
REM 4) Commit + Push
REM ===============================
echo [4/4] Commit + Push...

REM Monta data e hora sem caracteres invalidos
REM Ex: 2026-02-17 11-32-10
for /f "tokens=1-3 delims=/- " %%a in ("%date%") do (
    set D1=%%a
    set D2=%%b
    set D3=%%c
)

REM Ajuste automático (PT-BR normalmente vem DD/MM/AAAA)
REM Se seu Windows estiver diferente, ainda funciona pois pegamos 3 tokens
set DIA=%D1%
set MES=%D2%
set ANO=%D3%

for /f "tokens=1-3 delims=:." %%h in ("%time%") do (
    set H=%%h
    set M=%%i
    set S=%%j
)

REM Remove espaços do H (quando vem " 9" por exemplo)
set H=%H: =0%

set MSG=Atualizacao painel %ANO%-%MES%-%DIA% %H%-%M%-%S%

git commit -m "%MSG%"
set GIT_COMMIT_ERR=%ERRORLEVEL%

REM Se nao tiver nada pra commitar, nao é erro critico
REM git commit retorna 1 quando "nothing to commit"
IF "%GIT_COMMIT_ERR%"=="1" (
    echo.
    echo ⚠️ Nada novo para commitar (working tree clean).
) ELSE (
    IF NOT "%GIT_COMMIT_ERR%"=="0" (
        echo.
        echo ❌ ERRO: git commit falhou (codigo %GIT_COMMIT_ERR%).
        echo.
        pause
        exit /b %GIT_COMMIT_ERR%
    )
)

echo.
echo Enviando para o GitHub (git push)...
git push
set GIT_PUSH_ERR=%ERRORLEVEL%

IF NOT "%GIT_PUSH_ERR%"=="0" (
    echo.
    echo ❌ ERRO: git push falhou (codigo %GIT_PUSH_ERR%).
    echo Verifique se o remote esta correto e se voce tem permissao.
    echo.
    pause
    exit /b %GIT_PUSH_ERR%
)

echo.
echo =====================================
echo ✅ Painel atualizado com sucesso!
echo =====================================
echo.

REM Status final
git status
echo.
pause
endlocal
