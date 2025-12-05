@echo off
chcp 65001 >nul
title Configurar Ambiente - DocsGen

echo ============================================
echo    Configurador de Ambiente - DocsGen
echo ============================================
echo.

:: Detecta o diretório do script .bat
cd /d "%~dp0"

:: Primeiro tenta usar o Python Launcher (py)
where py >nul 2>&1
if %ERRORLEVEL% EQU 0 (
    echo Python Launcher encontrado
    py "%~dp0ambiente_config.py"
    goto :fim
)

:: Tenta encontrar python no PATH
where python >nul 2>&1
if %ERRORLEVEL% EQU 0 (
    echo Python encontrado no PATH
    python "%~dp0ambiente_config.py"
    goto :fim
)

:: Tenta localizar Python em caminhos comuns do usuário
if exist "%LOCALAPPDATA%\Programs\Python\Python314\python.exe" (
    echo Usando Python 3.14
    "%LOCALAPPDATA%\Programs\Python\Python314\python.exe" "%~dp0ambiente_config.py"
    goto :fim
)
if exist "%LOCALAPPDATA%\Programs\Python\Python313\python.exe" (
    echo Usando Python 3.13
    "%LOCALAPPDATA%\Programs\Python\Python313\python.exe" "%~dp0ambiente_config.py"
    goto :fim
)
if exist "%LOCALAPPDATA%\Programs\Python\Python312\python.exe" (
    echo Usando Python 3.12
    "%LOCALAPPDATA%\Programs\Python\Python312\python.exe" "%~dp0ambiente_config.py"
    goto :fim
)
if exist "%LOCALAPPDATA%\Programs\Python\Python311\python.exe" (
    echo Usando Python 3.11
    "%LOCALAPPDATA%\Programs\Python\Python311\python.exe" "%~dp0ambiente_config.py"
    goto :fim
)
if exist "%LOCALAPPDATA%\Programs\Python\Python310\python.exe" (
    echo Usando Python 3.10
    "%LOCALAPPDATA%\Programs\Python\Python310\python.exe" "%~dp0ambiente_config.py"
    goto :fim
)

:: Tenta localizar Python em caminhos de instalação global
if exist "C:\Python314\python.exe" (
    echo Usando Python 3.14 (instalacao global)
    "C:\Python314\python.exe" "%~dp0ambiente_config.py"
    goto :fim
)
if exist "C:\Python313\python.exe" (
    echo Usando Python 3.13 (instalacao global)
    "C:\Python313\python.exe" "%~dp0ambiente_config.py"
    goto :fim
)
if exist "C:\Python312\python.exe" (
    echo Usando Python 3.12 (instalacao global)
    "C:\Python312\python.exe" "%~dp0ambiente_config.py"
    goto :fim
)

:: Tenta Program Files
if exist "%ProgramFiles%\Python314\python.exe" (
    echo Usando Python 3.14 (Program Files)
    "%ProgramFiles%\Python314\python.exe" "%~dp0ambiente_config.py"
    goto :fim
)
if exist "%ProgramFiles%\Python313\python.exe" (
    echo Usando Python 3.13 (Program Files)
    "%ProgramFiles%\Python313\python.exe" "%~dp0ambiente_config.py"
    goto :fim
)

echo.
echo ERRO: Python nao foi encontrado!
echo Por favor, instale o Python em: https://www.python.org/downloads/
echo Certifique-se de marcar "Add Python to PATH" durante a instalacao.
echo.

:fim
echo.
echo ============================================
echo    Configuracao concluida!
echo ============================================
echo.
pause
