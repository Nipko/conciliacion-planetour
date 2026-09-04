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
REM 1. VERIFICAR O INSTALAR GIT (NATIVO CON WINGET EN WINDOWS 11)
REM -------------------------------------------------------------
echo [1/5] Verificando Git...
where git >nul 2>nul
if %errorlevel% neq 0 (
    echo [!] Git no está instalado en este equipo.
    echo     Instalando Git automáticamente mediante Windows 11 Package Manager (winget)...
    winget install --id Git.Git -e --source winget --accept-package-agreements --accept-source-agreements --silent
    
    REM Recargar PATH en la sesión actual
    set "PATH=%PATH%;C:\Program Files\Git\cmd;C:\Program Files\Git\bin"
    where git >nul 2>nul
    if %errorlevel% neq 0 (
        echo [ERROR] No se pudo encontrar Git tras la instalación. Reinicie esta ventana.
        pause
        exit /b 1
    )
    echo [✓] Git instalado con éxito.
) else (
    echo [✓] Git ya se encuentra instalado.
)

REM -------------------------------------------------------------
REM 2. VERIFICAR O INSTALAR PYTHON (NATIVO CON WINGET)
REM -------------------------------------------------------------
echo.
echo [2/5] Verificando Python...
where python >nul 2>nul
if %errorlevel% neq 0 (
    echo [!] Python no está instalado en este equipo.
    echo     Instalando Python 3 automáticamente mediante winget...
    winget install --id Python.Python.3.12 -e --source winget --accept-package-agreements --accept-source-agreements --silent
    
    REM Actualizar variables de entorno temporales
    for /f "tokens=*" %%p in ('powershell -Command "[System.Environment]::GetEnvironmentVariable('Path', 'Machine') + ';' + [System.Environment]::GetEnvironmentVariable('Path', 'User')"') do set "PATH=%%p"
    
    where python >nul 2>nul
    if %errorlevel% neq 0 (
        echo [ERROR] Python fue instalado pero requiere reiniciar la ventana de comandos o sesión.
        echo Por favor cierre esta ventana y vuelva a ejecutar este archivo.
        pause
        exit /b 1
    )
    echo [✓] Python instalado con éxito.
) else (
    echo [✓] Python ya se encuentra instalado.
)

REM -------------------------------------------------------------
REM 3. CREAR ENTORNO VIRTUAL E INSTALAR DEPENDENCIAS
REM -------------------------------------------------------------
echo.
echo [3/5] Configurando entorno virtual aislado (venv)...
if not exist "venv" (
    echo Creando carpeta venv...
    python -m venv venv
)

echo Instalando dependencias requeridas (Streamlit, Pandas, OpenpyXL, PDFPlumber)...
call venv\Scripts\activate.bat
python -m pip install --upgrade pip --quiet
pip install -r requirements.txt --quiet

if %errorlevel% neq 0 (
    echo [ERROR] Ocurrió un error instalando librerías. Reintentando con detalle...
    pip install -r requirements.txt
    pause
    exit /b 1
)
echo [✓] Todas las librerías se instalaron correctamente.

REM -------------------------------------------------------------
REM 4. CONFIGURAR REGLA DE FIREWALL (OPCIONAL PARA RED LOCAL)
REM -------------------------------------------------------------
echo.
echo [4/5] Habilitando puerto 8501 en el Firewall de Windows para red de oficina...
powershell -Command "if (-not (Get-NetFirewallRule -DisplayName 'Conciliador Planetour (8501)' -ErrorAction SilentlyContinue)) { New-NetFirewallRule -DisplayName 'Conciliador Planetour (8501)' -Direction Inbound -LocalPort 8501 -Protocol TCP -Action Allow -ErrorAction SilentlyContinue | Out-Null }" >nul 2>nul

REM Obtener IP local del equipo
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

python -m streamlit run app.py --server.port=8501 --server.address=0.0.0.0 --server.headless=true
pause
