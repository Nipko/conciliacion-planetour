@echo off
chcp 65001 > nul
title Conciliador Bancario - Planetour SAS
setlocal EnableDelayedExpansion

echo ============================================================
echo   PLANETOUR S.A.S. - CONCILIACIÓN BANCARIA AUTOMATIZADA
echo ============================================================
echo.

set "SCRIPT_DIR=%~dp0"
cd /d "%SCRIPT_DIR%"

REM -------------------------------------------------------------
REM 1. DETECTAR ENTORNO PYTHON O VENV
REM -------------------------------------------------------------
set "FOUND_PY="

REM 1.1 Revisar entorno virtual local
if exist "venv\Scripts\python.exe" (
    set "FOUND_PY=venv\Scripts\python.exe"
    goto :py_configured
)

REM 1.2 Revisar si python está en el PATH
where python >nul 2>nul
if %errorlevel% equ 0 (
    set "FOUND_PY=python"
    goto :build_venv
)

REM 1.3 Revisar lanzador oficial de Windows (py.exe)
where py >nul 2>nul
if %errorlevel% equ 0 (
    set "FOUND_PY=py"
    goto :build_venv
)

REM 1.4 Buscar en carpetas estándar de instalación de Windows
for /d %%D in ("%LocalAppData%\Programs\Python\Python*") do (
    if exist "%%D\python.exe" (
        set "FOUND_PY=%%D\python.exe"
        goto :build_venv
    )
)
for /d %%D in ("C:\Program Files\Python*") do (
    if exist "%%D\python.exe" (
        set "FOUND_PY=%%D\python.exe"
        goto :build_venv
    )
)
for /d %%D in ("C:\Program Files (x86)\Python*") do (
    if exist "%%D\python.exe" (
        set "FOUND_PY=%%D\python.exe"
        goto :build_venv
    )
)

:build_venv
REM Si no se encontró Python en ningún lugar, iniciar instalador automático
if not defined FOUND_PY (
    echo ============================================================
    echo   [!] Python no está instalado en este equipo.
    echo   Iniciando instalador automático para Windows 11...
    echo ============================================================
    echo.
    if exist "instalar_windows11.bat" (
        call instalar_windows11.bat
        exit /b 0
    ) else (
        echo [ERROR] No se encontró instalar_windows11.bat.
        echo Por favor instale Python 3 desde https://www.python.org/downloads/
        pause
        exit /b 1
    )
)

REM Si Python existe pero venv no ha sido creado, prepararlo automáticamente
if not exist "venv\Scripts\python.exe" (
    echo ============================================================
    echo   Configurando entorno por primera vez...
    echo   (Creando venv e instalando librerías requeridas)
    echo ============================================================
    echo.
    "%FOUND_PY%" -m venv venv
    if not exist "venv\Scripts\python.exe" (
        echo [ERROR] No se pudo crear el entorno virtual venv.
        pause
        exit /b 1
    )
    call venv\Scripts\activate.bat
    python -m pip install --upgrade pip --quiet
    pip install -r requirements.txt --quiet
    echo [✓] Entorno virtual preparado con éxito.
    echo.
)

:py_configured
set "RUN_PY=venv\Scripts\python.exe"
if not exist "!RUN_PY!" (
    if defined FOUND_PY (
        set "RUN_PY=%FOUND_PY%"
    ) else (
        set "RUN_PY=python"
    )
)

REM -------------------------------------------------------------
REM 2. OBTENER IP LOCAL E INICIAR SERVICIO
REM -------------------------------------------------------------
for /f "tokens=2 delims=:" %%a in ('ipconfig ^| findstr /c:"IPv4" ^| findstr /v "127.0.0.1"') do (
    set "LOCAL_IP=%%a"
    set "LOCAL_IP=!LOCAL_IP: =!"
    goto :ip_ready
)
:ip_ready

echo Iniciando panel interactivo en su navegador...
echo   - Local:    http://localhost:8501
if defined LOCAL_IP (
echo   - Red local: http://!LOCAL_IP!:8501
)
echo.

start "" "http://localhost:8501"
"!RUN_PY!" -m streamlit run app.py --server.port=8501 --server.address=0.0.0.0 --server.headless=true
pause
