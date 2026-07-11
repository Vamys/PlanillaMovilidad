@echo off
title Sistema de Movilidad - FFVV
echo =======================================================================
echo   SISTEMA DE PLANILLA DE MOVILIDAD Y CONSOLIDADOR DE PASAJES
echo =======================================================================
echo.

cd /d "%~dp0"

:: Check if virtual environment exists
if not exist .venv (
    echo [INFO] Creando entorno virtual de Python (.venv)...
    python -m venv .venv
    if errorlevel 1 (
        echo [ERROR] No se pudo crear el entorno virtual. Asegurate de tener Python instalado.
        pause
        exit /b 1
    )
)

:: Install dependencies
echo [INFO] Verificando e instalando dependencias (Flask, openpyxl)...
.venv\Scripts\pip install -r requirements.txt
if errorlevel 1 (
    echo [ERROR] No se pudieron instalar las dependencias.
    pause
    exit /b 1
)

echo.
echo [INFO] Servidor web iniciado en http://localhost:5000
echo [INFO] Para detener el servidor, cierra esta ventana o presiona Ctrl+C.
echo.

:: Open the default web browser automatically
start http://localhost:5000

:: Execute the Flask server
.venv\Scripts\python app.py
