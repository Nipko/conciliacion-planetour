@echo off
chcp 65001 > nul
setlocal EnableDelayedExpansion
title PLANETOUR S.A.S. - Actualizador de Cambios de Git (Windows 11)

echo ==============================================================================
echo        PLANETOUR S.A.S. - ACTUALIZADOR DE SISTEMA DESDE GITHUB
echo ==============================================================================
echo.

set "SCRIPT_DIR=%~dp0"
cd /d "%SCRIPT_DIR%"

REM 1. Descargar últimos cambios del repositorio
echo [1/3] Descargando últimos cambios desde GitHub...
where git >nul 2>nul
if %errorlevel% neq 0 (
    echo [ERROR] Git no está disponible en la terminal.
    pause
    exit /b 1
)

git fetch origin
git pull origin main

if %errorlevel% neq 0 (
    echo.
    echo [ADVERTENCIA] Hubo un problema al hacer git pull.
    echo Si tienes cambios locales no guardados, por favor revísalos.
    pause
    exit /b 1
)
echo [✓] Código actualizado a la última versión disponible en GitHub.

REM 2. Actualizar dependencias de Python
echo.
echo [2/3] Verificando dependencias en venv...
if not exist "venv" (
    echo Creando entorno virtual venv...
    python -m venv venv
)

call venv\Scripts\activate.bat
pip install -r requirements.txt --quiet
echo [✓] Dependencias verificadas y al día.

REM 3. Iniciar o reiniciar servicio
echo.
echo ==============================================================================
echo   ¡ACTUALIZACIÓN COMPLETADA CON ÉXITO!
echo ==============================================================================
echo.

REM Obtener IP local del equipo
for /f "tokens=2 delims=:" %%a in ('ipconfig ^| findstr /c:"IPv4" ^| findstr /v "127.0.0.1"') do (
    set "LOCAL_IP=%%a"
    set "LOCAL_IP=!LOCAL_IP: =!"
    goto :ip_found_up
)
:ip_found_up

echo 🌐 Enlaces de Acceso:
echo    - En este equipo:  http://localhost:8501
if defined LOCAL_IP (
echo    - En la red local: http://!LOCAL_IP!:8501
)
echo.
echo Iniciando servicio actualizado...
echo.

python -m streamlit run app.py --server.port=8501 --server.address=0.0.0.0 --server.headless=true
pause
