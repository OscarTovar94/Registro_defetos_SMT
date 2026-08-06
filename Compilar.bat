@echo off
setlocal

title Registros de defectos SMT Builder
color 0A

cd /d "%~dp0"

echo ================================================
echo        Registros de defectos SMT - COMPILADOR
echo ================================================
echo.

set /p "VERSION=Ingrese la nueva version (Ejemplo 1.0): "

if not defined VERSION (
    echo.
    echo ERROR: No se ingreso una version.
    pause
    exit /b 1
)

set "APPNAME=Registros_defectos_SMT_Rev%VERSION%"
set "ORIGEN=%~dp0dist\%APPNAME%"
set "DESTINO=%~dp0"

echo.
echo Nueva compilacion:
echo %APPNAME%.exe
echo.

echo Eliminando compilaciones anteriores...

if exist "%~dp0build" rmdir /s /q "%~dp0build"
if exist "%~dp0dist" rmdir /s /q "%~dp0dist"
if exist "%~dp0_internal" rmdir /s /q "%~dp0_internal"

del /q "%~dp0Registros_defectos_SMT_Rev*.exe" 2>nul
del /q "%~dp0*.spec" 2>nul
del /q "%~dp0^" 2>nul

echo.
echo Iniciando PyInstaller...
echo.

python -m PyInstaller --collect-all babel --collect-all tkcalendar --onedir --noconsole --clean --icon="C:\Registro_defetos_SMT\Image\elrad.ico" --name "%APPNAME%" "main.py"

if errorlevel 1 (
    echo.
    echo ================================================
    echo       ERROR DURANTE LA COMPILACION
    echo ================================================
    echo.
    pause
    exit /b 1
)

if not exist "%ORIGEN%\%APPNAME%.exe" (
    echo.
    echo ERROR: No se encontro el ejecutable generado:
    echo "%ORIGEN%\%APPNAME%.exe"
    echo.
    pause
    exit /b 1
)

echo.
echo Copiando ejecutable y carpeta _internal...
echo.

xcopy "%ORIGEN%\*" "%DESTINO%\" /E /I /Y

if errorlevel 8 (
    echo.
    echo ERROR: No fue posible copiar los archivos.
    pause
    exit /b 1
)

echo.
echo Eliminando archivos temporales...

if exist "%~dp0build" rmdir /s /q "%~dp0build"
if exist "%~dp0dist" rmdir /s /q "%~dp0dist"
del /q "%~dp0*.spec" 2>nul

echo.
echo ================================================
echo        COMPILACION FINALIZADA
echo ================================================
echo.
echo Ejecutable generado:
echo "%~dp0%APPNAME%.exe"
echo.
echo Carpeta de dependencias:
echo "%~dp0_internal"
echo.

pause
endlocal