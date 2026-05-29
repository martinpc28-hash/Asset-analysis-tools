@echo off
REM ============================================================
REM  Atajo para ejecutar el analisis GLD vs SP500
REM  Doble click a este archivo y listo.
REM ============================================================

REM Ubica este script en su propia carpeta
cd /d "%~dp0"

echo.
echo === Activando entorno Anaconda ===
call conda activate analisis 2>nul
if errorlevel 1 (
    echo  - Entorno "analisis" no encontrado, usando "base"
    call conda activate base
)

echo.
echo === Ejecutando analizar_gld.py ===
python analizar_gld.py

echo.
echo === Terminado ===
pause
