@echo off
REM AgenteSAE - INSTALADOR (Setup.exe con accesos directos y desinstalador)
REM Requiere Inno Setup 6 instalado: https://jrsoftware.org/isdl.php
echo Instalando dependencias...
py -m pip install --upgrade pyinstaller pyside6 python-docx lxml
echo.
echo Generando aplicacion (carpeta)...
py -m PyInstaller --noconfirm --windowed --name AgenteSAE --distpath dist_app app.py
echo.
echo Empaquetando instalador con Inno Setup...
"C:\Program Files (x86)\Inno Setup 6\ISCC.exe" installer.iss
echo.
echo Listo: Output\AgenteSAE-Setup.exe
pause
