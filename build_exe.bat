@echo off
REM AgenteSAE - genera el ejecutable de escritorio (Windows)
echo Instalando dependencias...
py -m pip install --upgrade pyinstaller python-docx lxml
echo.
echo Generando ejecutable...
py -m PyInstaller --onefile --windowed --name AgenteSAE app.py
echo.
echo Listo. El ejecutable esta en:  dist\AgenteSAE.exe
pause
