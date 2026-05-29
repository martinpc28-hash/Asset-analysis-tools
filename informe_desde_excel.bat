@echo off
setlocal EnableDelayedExpansion
REM ============================================================
REM  Genera el INFORME HTML interactivo a partir de un Excel
REM  (plantilla rellenada o reporte de analizador.py).
REM  Busca Anaconda automaticamente; si no la encuentra,
REM  usa el python que este en el PATH.
REM ============================================================

cd /d "%~dp0"

echo.
echo === Buscando Anaconda ===

REM Posibles ubicaciones de Anaconda/Miniconda
set "CONDA_PATHS=C:\ProgramData\Anaconda3;C:\ProgramData\Miniconda3;%USERPROFILE%\Anaconda3;%USERPROFILE%\Miniconda3;%USERPROFILE%\AppData\Local\anaconda3;%USERPROFILE%\AppData\Local\miniconda3;C:\Anaconda3;C:\Miniconda3"

set "CONDA_BAT="
for %%P in (%CONDA_PATHS:;= %) do (
    if exist "%%P\Scripts\activate.bat" (
        set "CONDA_BAT=%%P\Scripts\activate.bat"
        echo  - Anaconda encontrada en: %%P
        goto :found
    )
)

echo  - Anaconda NO encontrada en ubicaciones tipicas.
echo  - Intentando con python que este en el PATH...
goto :runpy

:found
call "%CONDA_BAT%" analisis 2>nul
if errorlevel 1 (
    echo  - Entorno "analisis" no existe, usando "base"
    call "%CONDA_BAT%" base
)

:runpy
echo.
where python >nul 2>&1
if errorlevel 1 (
    echo [ERROR] No encuentro python. Abre Anaconda Prompt y ejecuta:
    echo         conda init cmd.exe
    echo         (despues cierra y vuelve a abrir cmd^)
    pause
    exit /b 1
)

echo.
set /p ARCHIVO="Arrastra aqui tu Excel y pulsa Enter (vacio = Plantilla_Analisis_vs_SP500.xlsx): "

REM Quita comillas si las trae al arrastrar
set ARCHIVO=%ARCHIVO:"=%

if "%ARCHIVO%"=="" (
    python informe_html.py
) else (
    python informe_html.py "%ARCHIVO%"
)

echo.
pause
