@echo off
chcp 65001 > nul
setlocal EnableDelayedExpansion
title PLANETOUR S.A.S. - Instalación Automatizada para Windows 11

echo ==============================================================================
echo        PLANETOUR S.A.S. - SISTEMA DE CONCILIACIÓN BANCARIA
echo                 Instalador Automático para Windows 11
echo ==============================================================================
echo.

set "SCRIPT_DIR=%~dp0"
cd /d "%SCRIPT_DIR%"

REM -------------------------------------------------------------
REM 1. VERIFICAR O INSTALAR GIT (WINGET EN WINDOWS 11)
REM -------------------------------------------------------------
echo [1/5] Verificando Git...
where git >nul 2>nul
if %errorlevel% neq 0 (
    echo [!] Git no está instalado en este equipo.
    echo     Instalando Git automáticamente mediante winget...
    winget install --id Git.Git -e --source winget --accept-package-agreements --accept-source-agreements --silent
    
    set "PATH=%PATH%;C:\Program Files\Git\cmd;C:\Program Files\Git\bin"
    where git >nul 2>nul
    if %errorlevel% neq 0 (
        echo [!] Git fue instalado. Continuando con la configuración...
    ) else (
        echo [✓] Git instalado con éxito.
    )
) else (
    echo [✓] Git ya se encuentra instalado.
)

REM -------------------------------------------------------------
REM 2. VERIFICAR O INSTALAR PYTHON
REM -------------------------------------------------------------
echo.
echo [2/5] Verificando Python...
set "FOUND_PY="

where python >nul 2>nul && set "FOUND_PY=python"
if not defined FOUND_PY (
    where py >nul 2>nul && set "FOUND_PY=py"
)
if not defined FOUND_PY (
    for /d %%D in ("%LocalAppData%\Programs\Python\Python*") do (
        if exist "%%D\python.exe" set "FOUND_PY=%%D\python.exe"
    )
)
if not defined FOUND_PY (
    for /d %%D in ("C:\Program Files\Python*") do (
        if exist "%%D\python.exe" set "FOUND_PY=%%D\python.exe"
    )
)
if not defined FOUND_PY (
    for /d %%D in ("C:\Program Files (x86)\Python*") do (
        if exist "%%D\python.exe" set "FOUND_PY=%%D\python.exe"
    )
)

if not defined FOUND_PY (
    echo [!] Python no está instalado en este equipo.
    echo     Instalando Python 3 automáticamente mediante winget...
    winget install --id Python.Python.3.12 -e --source winget --accept-package-agreements --accept-source-agreements --silent
    
    where python >nul 2>nul && set "FOUND_PY=python"
    if not defined FOUND_PY (
        where py >nul 2>nul && set "FOUND_PY=py"
    )
    if not defined FOUND_PY (
        for /d %%D in ("%LocalAppData%\Programs\Python\Python*") do (
            if exist "%%D\python.exe" set "FOUND_PY=%%D\python.exe"
        )
    )
    if not defined FOUND_PY (
        for /d %%D in ("C:\Program Files\Python*") do (
            if exist "%%D\python.exe" set "FOUND_PY=%%D\python.exe"
        )
    )
)

if not defined FOUND_PY (
    echo [ERROR] No se pudo encontrar Python en el sistema.
    echo Por favor instale Python 3.11 o 3.12 desde: https://www.python.org/downloads/
    echo Recuerde marcar la opción: "Add python.exe to PATH"
    pause
    exit /b 1
)

echo [✓] Python detectado: "!FOUND_PY!"

REM -------------------------------------------------------------
REM 3. CREAR ENTORNO VIRTUAL E INSTALAR DEPENDENCIAS
REM -------------------------------------------------------------
echo.
echo [3/5] Configurando entorno virtual aislado (venv)...
if not exist "venv\Scripts\python.exe" (
    echo Creando carpeta venv con "!FOUND_PY!"...
    "!FOUND_PY!" -m venv venv
)

if not exist "venv\Scripts\python.exe" (
    echo [ERROR] Falló la creación del entorno virtual venv.
    pause
    exit /b 1
)

echo Instalando librerías requeridas (Streamlit, Pandas, OpenpyXL, PDFPlumber)...
call venv\Scripts\activate.bat
python -m pip install --upgrade pip --quiet
pip install -r requirements.txt --quiet

if %errorlevel% neq 0 (
    echo [AVISO] Reintentando instalación con salida detallada...
    pip install -r requirements.txt
)
echo [✓] Todas las librerías instaladas con éxito en venv.

REM -------------------------------------------------------------
REM 4. CONFIGURAR REGLA DE FIREWALL (RED LOCAL)
REM -------------------------------------------------------------
echo.
echo [4/5] Habilitando puerto 8501 en el Firewall de Windows para red de oficina...
powershell -Command "if (-not (Get-NetFirewallRule -DisplayName 'Conciliador Planetour (8501)' -ErrorAction SilentlyContinue)) { New-NetFirewallRule -DisplayName 'Conciliador Planetour (8501)' -Direction Inbound -LocalPort 8501 -Protocol TCP -Action Allow -ErrorAction SilentlyContinue | Out-Null }" >nul 2>nul

for /f "tokens=2 delims=:" %%a in ('ipconfig ^| findstr /c:"IPv4" ^| findstr /v "127.0.0.1"') do (
    set "LOCAL_IP=%%a"
    set "LOCAL_IP=!LOCAL_IP: =!"
    goto :ip_found
)
:ip_found

REM -------------------------------------------------------------
REM 5. INICIAR EL SERVICIO WEB
REM -------------------------------------------------------------
echo.
echo ==============================================================================
echo   ¡INSTALACIÓN COMPLETADA! INICIANDO SERVICIO WEB DE CONCILIACIÓN
echo ==============================================================================
echo.
echo 🌐 Enlaces de Acceso:
echo    - En este mismo equipo:  http://localhost:8501
if defined LOCAL_IP (
echo    - Desde otros computadores de la oficina: http://!LOCAL_IP!:8501
)
echo.
echo (Mantenga esta ventana abierta para que el servicio permanezca disponible)
echo.

start "" "http://localhost:8501"
venv\Scripts\python.exe -m streamlit run app.py --server.port=8501 --server.address=0.0.0.0 --server.headless=true
pause
