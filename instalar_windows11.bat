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
    echo [!] Git no está instalado en este servidor.
    echo     Instalando Git automáticamente con winget...
    winget install --id Git.Git -e --source winget --accept-package-agreements --accept-source-agreements --silent
    set "PATH=%PATH%;C:\Program Files\Git\cmd;C:\Program Files\Git\bin"
)
echo [✓] Git verificado.

REM -------------------------------------------------------------
REM 2. VERIFICAR SI PYTHON ESTÁ REALMENTE INSTALADO (DESCARTANDO ALIAS FALSO DE WINDOWS)
REM -------------------------------------------------------------
echo.
echo [2/5] Verificando Python funcional en este servidor...
set "REAL_PY="

REM 2.1 Probar si ya existe venv local funcional
if exist "venv\Scripts\python.exe" (
    venv\Scripts\python.exe -c "import sys" >nul 2>nul
    if !errorlevel! equ 0 (
        set "REAL_PY=venv\Scripts\python.exe"
        goto :py_validated
    )
)

REM 2.2 Probar py launcher oficial de Python
where py >nul 2>nul
if %errorlevel% equ 0 (
    py -3 -c "import sys" >nul 2>nul
    if !errorlevel! equ 0 (
        set "REAL_PY=py -3"
        goto :py_validated
    )
)

REM 2.3 Probar python en PATH descartando expresamente el alias falso de Microsoft WindowsApps
where python >nul 2>nul
if %errorlevel% equ 0 (
    for /f "delims=" %%i in ('where python') do (
        echo %%i | findstr /i "WindowsApps" >nul
        if errorlevel 1 (
            "%%i" -c "import sys" >nul 2>nul
            if !errorlevel! equ 0 (
                set "REAL_PY=%%i"
                goto :py_validated
            )
        )
    )
)

REM 2.4 Buscar en rutas estándar de instalación de Python en Windows
for /d %%D in ("%LocalAppData%\Programs\Python\Python*") do (
    if exist "%%D\python.exe" (
        "%%D\python.exe" -c "import sys" >nul 2>nul
        if !errorlevel! equ 0 (
            set "REAL_PY=%%D\python.exe"
            goto :py_validated
        )
    )
)
for /d %%D in ("C:\Program Files\Python*") do (
    if exist "%%D\python.exe" (
        "%%D\python.exe" -c "import sys" >nul 2>nul
        if !errorlevel! equ 0 (
            set "REAL_PY=%%D\python.exe"
            goto :py_validated
        )
    )
)
for /d %%D in ("C:\Python*") do (
    if exist "%%D\python.exe" (
        "%%D\python.exe" -c "import sys" >nul 2>nul
        if !errorlevel! equ 0 (
            set "REAL_PY=%%D\python.exe"
            goto :py_validated
        )
    )
)

REM 2.5 Si Python NO está instalado de verdad, instalarlo automáticamente
if not defined REAL_PY (
    echo.
    echo [!] Python no está instalado en este equipo.
    echo     (El comando detectado anteriormente era solo un acceso directo de Windows Store).
    echo     Descargando e instalando Python 3.12 oficial de forma desatendida...
    echo.

    powershell -Command "[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12; Write-Host 'Descargando instalador oficial de python.org...'; $f=\"$env:TEMP\python-3.12.8-amd64.exe\"; Invoke-WebRequest -Uri 'https://www.python.org/ftp/python/3.12.8/python-3.12.8-amd64.exe' -OutFile $f; Write-Host 'Instalando Python 3.12 en segundo plano...'; Start-Process -FilePath $f -ArgumentList '/quiet InstallAllUsers=0 PrependPath=1 Include_test=0 Include_pip=1' -Wait; Remove-Item -Force $f"

    REM Volver a buscar el Python recién instalado
    for /d %%D in ("%LocalAppData%\Programs\Python\Python*") do (
        if exist "%%D\python.exe" (
            set "REAL_PY=%%D\python.exe"
        )
    )
    if not defined REAL_PY (
        for /d %%D in ("C:\Program Files\Python*") do (
            if exist "%%D\python.exe" (
                set "REAL_PY=%%D\python.exe"
            )
        )
    )
    if not defined REAL_PY (
        for /d %%D in ("C:\Python*") do (
            if exist "%%D\python.exe" (
                set "REAL_PY=%%D\python.exe"
            )
        )
    )
)

if not defined REAL_PY (
    echo [ERROR] No se pudo completar la instalación automática de Python.
    echo Por favor instale Python 3.11 o 3.12 manualmente desde: https://www.python.org/downloads/
    echo (Recuerde marcar la casilla: "Add python.exe to PATH")
    pause
    exit /b 1
)

:py_validated
echo [✓] Python funcional listo: !REAL_PY!

REM -------------------------------------------------------------
REM 3. CREAR ENTORNO VIRTUAL E INSTALAR DEPENDENCIAS
REM -------------------------------------------------------------
echo.
echo [3/5] Configurando entorno virtual aislado (venv)...
if not exist "venv\Scripts\python.exe" (
    echo Creando entorno venv...
    !REAL_PY! -m venv venv
)

if not exist "venv\Scripts\python.exe" (
    echo [ERROR] Falló la creación del entorno virtual venv.
    pause
    exit /b 1
)

echo [✓] Entorno virtual preparado.
echo Instalando dependencias requeridas (Streamlit, Pandas, OpenpyXL, PDFPlumber)...
call venv\Scripts\activate.bat
python -m pip install --upgrade pip --quiet
pip install -r requirements.txt --quiet
echo [✓] Todas las librerías instaladas con éxito.

REM -------------------------------------------------------------
REM 4. CONFIGURAR REGLA DE FIREWALL (RED LOCAL)
REM -------------------------------------------------------------
echo.
echo [4/5] Habilitando puerto 8501 en el Firewall de Windows para red local...
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
