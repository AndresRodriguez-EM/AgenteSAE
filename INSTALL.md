# AgenteSAE — Instalación y empaquetado (aplicación de escritorio)

La aplicación tiene una interfaz gráfica (`app.py`) construida con **Tkinter**
(incluido en Python en Windows y macOS). Funciona para las 6 sociedades
(IRCA, SANTA, MONTOYA, ZARLHA, INVERMAP, CIA) y **no modifica la Nota 3**.

## Opción A — Ejecutar con Python (rápido)

Requiere Python 3.10+ instalado (en Windows, marcar *"Add Python to PATH"*).

```bat
py -m pip install python-docx lxml
py app.py
```

Se abre la ventana: eliges la sociedad, los tres archivos (PDF del período,
PDF comparativo del año anterior y el Word de notas), y pulsas **Procesar**.
Genera el Word actualizado (`..._ACTUALIZADO.docx`) y un informe (`..._informe.md`).

## Opción B — Generar un ejecutable de escritorio (.exe, sin instalar Python)

En una máquina **Windows** (PyInstaller no hace compilación cruzada: el `.exe`
se genera en Windows):

```bat
py -m pip install --upgrade pyinstaller python-docx lxml
py -m PyInstaller --onefile --windowed --name AgenteSAE app.py
```

El ejecutable queda en **`dist\AgenteSAE.exe`** — se puede copiar a cualquier PC
Windows y abrir con doble clic, sin necesidad de Python.

> Atajo: ejecutar el archivo **`build_exe.bat`** incluido en el repositorio.

### macOS

```bash
pip3 install pyinstaller python-docx lxml
pyinstaller --onefile --windowed --name AgenteSAE app.py
# resultado: dist/AgenteSAE.app
```

## Uso mes a mes

1. Exportar del ERP el **balance de comprobación (PDF)** del mes (período actual)
   y el del **mismo mes del año anterior** (comparativo).
2. Abrir AgenteSAE, elegir la sociedad y los tres archivos.
3. **Procesar** → revisar el panel de resultado (cambios, altas/bajas y
   observaciones) → el Word actualizado queda guardado con su informe.

La app recuerda la última carpeta usada y sugiere automáticamente el nombre de
salida y la sociedad a partir del nombre del Word.

## Notas

- La **Nota 3** (Propiedad, planta y equipo) se deja intacta a propósito.
- Si una sociedad usa un template de notas distinto, el informe lo avisa
  ("no se encontró la tabla …"); en ese caso se ajusta `NOTE_CONFIG` en
  `agentesae/docx_updater.py`.
