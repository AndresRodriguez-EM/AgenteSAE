@echo off
setlocal enabledelayedexpansion
REM AgenteSAE - INSTALADOR (Setup.exe con accesos directos y desinstalador)
REM Requiere Inno Setup 6:  https://jrsoftware.org/isdl.php

echo Instalando dependencias...
py -m pip install --upgrade pyinstaller pyside6 python-docx lxml
echo.
echo Generando aplicacion (carpeta)...
py -m PyInstaller --noconfirm --windowed --name AgenteSAE --icon assets\icon.ico --add-data "assets;assets" --distpath dist_app app.py
if errorlevel 1 (
  echo [ERROR] Fallo la generacion de la aplicacion con PyInstaller.
  pause & exit /b 1
)
echo.

REM --- Localizar Inno Setup (ISCC.exe) ---
set "ISCC="
if exist "%ProgramFiles(x86)%\Inno Setup 6\ISCC.exe" set "ISCC=%ProgramFiles(x86)%\Inno Setup 6\ISCC.exe"
if not defined ISCC if exist "%ProgramFiles%\Inno Setup 6\ISCC.exe" set "ISCC=%ProgramFiles%\Inno Setup 6\ISCC.exe"
if not defined ISCC for /f "delims=" %%i in ('where ISCC 2^>nul') do set "ISCC=%%i"

if not defined ISCC (
  echo ============================================================
  echo  No se encontro Inno Setup ^(ISCC.exe^).
  echo.
  echo  La APLICACION ya quedo lista y funcional en:
  echo      dist_app\AgenteSAE\AgenteSAE.exe
  echo.
  echo  Para generar el instalador "Setup.exe" instala Inno Setup 6:
  echo      https://jrsoftware.org/isdl.php
  echo  y vuelve a ejecutar este archivo.
  echo.
  echo  ^(Si solo quieres un .exe portable, usa build_exe.bat;
  echo   ese NO necesita Inno Setup.^)
  echo ============================================================
  pause & exit /b 1
)

echo Empaquetando instalador con Inno Setup...
echo   Usando: "!ISCC!"
"!ISCC!" installer.iss
if errorlevel 1 (
  echo [ERROR] Inno Setup no pudo generar el instalador.
  pause & exit /b 1
)
echo.
echo Listo: Output\AgenteSAE-Setup.exe
pause
