@echo off
REM AgenteSAE - version PORTABLE (un solo .exe, no requiere Python)
echo Instalando dependencias...
py -m pip install --upgrade pyinstaller pyside6 python-docx lxml
echo.
echo Generando ejecutable portable...
py -m PyInstaller --noconfirm --onefile --windowed --name AgenteSAE --distpath dist_portable app.py
echo.
echo Listo: dist_portable\AgenteSAE.exe  (copialo a cualquier PC Windows)
pause
