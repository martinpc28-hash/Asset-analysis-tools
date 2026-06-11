@echo off
setlocal EnableDelayedExpansion
REM ============================================================
REM  Lanza la interfaz grafica (Tkinter) del analizador.
REM  Usa "python" (no pythonw) para que cualquier error sea VISIBLE
REM  en la ventana negra. Si la ventana se cierra rapido o ves un
REM  error, podras leerlo: el cmd se queda abierto con pause al final.
REM ============================================================

cd /d "%~dp0"

echo === Buscando Anaconda ===

REM --- Buscar Anaconda ---
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
if errorlevel 1 (
    echo  - Entorno "analisis" no existe, usando "base"
    call "%CONDA_BAT%" base
)

:runpy
where python >nul 2>&1
if errorlevel 1 (
    echo.
    echo [ERROR] No encuentro python en el PATH.
    echo Abre Anaconda Prompt y ejecuta:  conda init cmd.exe
    echo Cierra cmd, vuelve a abrir y prueba otra vez.
    pause
    exit /b 1
)

echo.
echo === Lanzando interfaz.py ===
echo (Si hay un error, aparecera abajo y la ventana se quedara abierta^)
echo.

python interfaz.py
set EXITCODE=%ERRORLEVEL%

echo.
echo === Programa terminado con codigo %EXITCODE% ===
if not "%EXITCODE%"=="0" (
    echo [ATENCION] Hubo un error - lee el mensaje arriba.
)
pause
