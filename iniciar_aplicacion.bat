@echo off
chcp 65001 > nul
title Sistema de Gestion Academica - UNAD (ECISA)
cd /d "%~dp0"

echo ============================================================================
echo   SISTEMA DE GESTION ACADEMICA (DOCENTES, MATRICULAS Y HOMOLOGACIONES)
echo ============================================================================
echo.

:: 1. Verificar si existe el entorno virtual .venv; si no existe (ej. en otro equipo), lo crea e instala todo
if not exist ".venv\Scripts\streamlit.exe" (
    echo [1/2] No se detecto el entorno virtual local. Creando .venv e instalando librerias...
    python -m venv .venv
    if errorlevel 1 (
        echo [ERROR] No se encontro Python instalado en este equipo.
        echo Descarga Python 3.11+ desde https://www.python.org/downloads/ marcando "Add Python to PATH".
        pause
        exit /b 1
    )
    .\.venv\Scripts\python.exe -m pip install --upgrade pip
    .\.venv\Scripts\python.exe -m pip install -r requirements.txt
)

echo [OK] Iniciando servidor Streamlit...
echo.
echo  - Acceso en este equipo:       http://localhost:8501
echo  - Acceso desde otros equipos en la misma red Wi-Fi / LAN:
for /f "tokens=2 delims=:" %%a in ('ipconfig ^| findstr /c:"IPv4"') do (
    echo    http://%%a:8501
)
echo.
echo (Para cerrar la aplicacion, cierra esta ventana o presiona Ctrl + C)
echo ============================================================================
echo.

.\.venv\Scripts\streamlit.exe run app.py --server.address 0.0.0.0 --server.port 8501

pause
