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

REM Usar venv si existe, o python global
if exist "venv\Scripts\python.exe" (
    set "PY_CMD=venv\Scripts\python.exe"
) else (
    set "PY_CMD=python"
)

REM Obtener IP local
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
!PY_CMD! -m streamlit run app.py --server.port=8501 --server.address=0.0.0.0 --server.headless=true
pause
