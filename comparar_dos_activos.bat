@echo off
setlocal EnableDelayedExpansion
REM ============================================================
REM  Compara DOS activos cualesquiera (tickers Yahoo o archivos)
REM  Ejemplo: AAPL vs MSFT, GLD vs QQQ, fondo vs ETF, etc.
REM ============================================================

cd /d "%~dp0"

echo.
echo === Buscando Anaconda ===
set "CONDA_PATHS=C:\ProgramData\Anaconda3;C:\ProgramData\Miniconda3;%USERPROFILE%\Anaconda3;%USERPROFILE%\Miniconda3;%USERPROFILE%\AppData\Local\anaconda3;%USERPROFILE%\AppData\Local\miniconda3;C:\Anaconda3;C:\Miniconda3"

set "CONDA_BAT="
for %%P in (%CONDA_PATHS:;= %) do (
    if exist "%%P\Scripts\activate.bat" (
        set "CONDA_BAT=%%P\Scripts\activate.bat"
        echo  - Anaconda encontrada en: %%P
        goto :found
    )
)
echo  - Anaconda NO encontrada; intentando python del PATH.
goto :runpy

:found
call "%CONDA_BAT%" analisis 2>nul
if errorlevel 1 call "%CONDA_BAT%" base

:runpy
where python >nul 2>&1
if errorlevel 1 (
    echo [ERROR] No encuentro python. Abre Anaconda Prompt y ejecuta: conda init cmd.exe
    pause
    exit /b 1
)

echo.
echo === COMPARAR DOS ACTIVOS ===
echo (puedes usar tickers Yahoo como GLD, AAPL, ^^GSPC, BTC-USD, o rutas a archivos .xlsx/.csv^)
echo.
set /p ACTIVO1="Activo 1 (el que quieres analizar): "
set /p ACTIVO2="Activo 2 (benchmark / referencia): "

if "%ACTIVO1%"=="" (
    echo [ERROR] No escribiste Activo 1
    pause
    exit /b 1
)
if "%ACTIVO2%"=="" (
    echo [ERROR] No escribiste Activo 2
    pause
    exit /b 1
)

REM Quita comillas si las trae
set ACTIVO1=%ACTIVO1:"=%
set ACTIVO2=%ACTIVO2:"=%

echo.
echo === Analizando: %ACTIVO1%  vs  %ACTIVO2% ===
python analizador.py "%ACTIVO1%" --benchmark "%ACTIVO2%"

echo.
pause
