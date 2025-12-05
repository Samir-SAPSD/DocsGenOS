@echo off
chcp 65001 >nul
title Configurar Ambiente - DocsGen

echo ============================================
echo    Configurador de Ambiente - DocsGen
echo ============================================
echo.

:: Detecta o diretório do script .bat
cd /d "%~dp0"

:: Tenta encontrar o Python no PATH ou em locais comuns
where python >nul 2>&1
if %ERRORLEVEL% EQU 0 (
    echo Python encontrado no PATH
    python "%~dp0ambiente_config.py"
) else (
    :: Tenta localizar Python em caminhos comuns
    if exist "%LOCALAPPDATA%\Programs\Python\Python314\python.exe" (
        echo Usando Python 3.14
        "%LOCALAPPDATA%\Programs\Python\Python314\python.exe" "%~dp0ambiente_config.py"
    ) else if exist "%LOCALAPPDATA%\Programs\Python\Python313\python.exe" (
        echo Usando Python 3.13
        "%LOCALAPPDATA%\Programs\Python\Python313\python.exe" "%~dp0ambiente_config.py"
    ) else if exist "%LOCALAPPDATA%\Programs\Python\Python312\python.exe" (
        echo Usando Python 3.12
        "%LOCALAPPDATA%\Programs\Python\Python312\python.exe" "%~dp0ambiente_config.py"
    ) else if exist "%LOCALAPPDATA%\Programs\Python\Python311\python.exe" (
        echo Usando Python 3.11
        "%LOCALAPPDATA%\Programs\Python\Python311\python.exe" "%~dp0ambiente_config.py"
    ) else if exist "%LOCALAPPDATA%\Programs\Python\Python310\python.exe" (
        echo Usando Python 3.10
        "%LOCALAPPDATA%\Programs\Python\Python310\python.exe" "%~dp0ambiente_config.py"
    ) else (
        echo.
        echo ERRO: Python nao foi encontrado!
        echo Por favor, instale o Python em: https://www.python.org/downloads/
        echo.
    )
)

echo.
echo ============================================
echo    Configuracao concluida!
echo ============================================
echo.
pause
