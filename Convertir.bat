@echo off
REM Arrastra el Excel de invitados encima de este archivo y suelta.
REM Genera el .glf al lado del Excel, listo para abrir en Guest List.

setlocal
cd /d "%~dp0"

if "%~1"=="" (
    echo.
    echo   Arrastra el archivo Excel de invitados encima de "Convertir.bat".
    echo.
    pause
    exit /b 1
)

set "EVENTO=%~n1"
set /p EVENTO=Nombre del evento [%EVENTO%]:

echo.
python excel_a_glf.py "%~1" --evento "%EVENTO%"
echo.
pause
