@echo off
chcp 65001 > nul
setlocal enabledelayedexpansion
title Gestor de Reembolsos

echo ==================================================================
echo         GESTOR INTELIGENTE DE REEMBOLSOS (WINDOWS)
echo ==================================================================
echo.

cd /d "%~dp0"

:: 1. Verificar si Python está instalado en el sistema
where python >nul 2>nul
if %errorlevel% neq 0 (
    where py >nul 2>nul
    if %errorlevel% neq 0 (
        echo [ERROR] No se ha encontrado Python instalado en tu computadora.
        echo.
        echo Para usar esta aplicacion necesitas instalar Python 3.10 o superior:
        echo    1. Descargalo gratis desde: https://www.python.org/downloads/
        echo    2. IMPORTANTE: Durante la instalacion, marca la casilla que dice:
        echo       "[X] Add Python to PATH" (Agregar Python al PATH).
        echo.
        pause
        exit /b 1
    ) else (
        set "PYTHON_CMD=py"
    )
) else (
    set "PYTHON_CMD=python"
)

:: 2. Crear y configurar el entorno virtual en la primera ejecucion
if not exist ".venv\Scripts\python.exe" (
    echo [INFO] Configurando la aplicacion por primera vez...
    echo Esto solo se hace una vez y puede tardar 1 o 2 minutos.
    echo.
    echo Creando entorno virtual (.venv)...
    %PYTHON_CMD% -m venv .venv
    if %errorlevel% neq 0 (
        echo [ERROR] No se pudo crear el entorno virtual.
        pause
        exit /b 1
    )
    
    echo Instalando librerias necesarias (OCR, PDF, Excel)...
    .venv\Scripts\python.exe -m pip install --upgrade pip --quiet
    .venv\Scripts\python.exe -m pip install -r requirements.txt
    if %errorlevel% neq 0 (
        echo [ADVERTENCIA] Algunas dependencias reportaron advertencias. Continuando...
    )
    echo.
    echo [OK] Configuracion inicial completada con exito.
    echo.
)

:: 3. Iniciar el servidor de la aplicacion
echo [INFO] Iniciando el Gestor de Reembolsos...
echo [INFO] Abriendo la aplicacion en tu navegador...
echo.
echo Presiona Ctrl + C en esta ventana cuando desees cerrar la aplicacion.
echo.

.venv\Scripts\python.exe app.py 5050

if %errorlevel% neq 0 (
    echo.
    echo Hubo un problema al ejecutar la aplicacion.
    pause
)
