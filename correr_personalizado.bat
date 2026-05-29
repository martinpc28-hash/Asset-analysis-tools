@echo off
REM ============================================================
REM  Atajo interactivo: pide ticker(s) y los analiza vs SP500
REM ============================================================

cd /d "%~dp0"

echo.
echo === Activando entorno Anaconda ===
call conda activate analisis 2>nul
if errorlevel 1 call conda activate base

echo.
set /p TICKERS="Escribe uno o mas tickers separados por espacio (ej: GLD TLT QQQ): "

if "%TICKERS%"=="" (
    echo No escribiste nada, usando GLD por defecto.
    set TICKERS=GLD
)

echo.
echo === Analizando: %TICKERS%  vs  SP500 ===
python analizador.py %TICKERS%

echo.
pause
