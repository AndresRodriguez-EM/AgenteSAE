# AgenteSAE — Instalación y empaquetado

La aplicación tiene una **interfaz profesional en Qt** (`app.py`, PySide6): barra
de menú, barra de herramientas, panel de archivos y panel de resultado, con
**detección automática por carpeta**. Funciona para las 6 sociedades (IRCA,
SANTA, MONTOYA, ZARLHA, INVERMAP, CIA) y **no modifica la Nota 3**.

> `app_tk.py` es una versión alternativa ligera (Tkinter), por si se necesita.

## Opción A — Ejecutar con Python (para probar/desarrollar)

```bat
py -m pip install -r requirements.txt
py app.py
```

## Opción B — Instalar SIN Python (ejecutable de escritorio)

El `.exe` se compila en **Windows** (PyInstaller no hace compilación cruzada).
Se generan los **dos formatos**:

### B.1 Portable (un solo archivo)

Ejecuta **`build_exe.bat`** → produce **`dist_portable\AgenteSAE.exe`**.
Es un único archivo: doble clic y funciona, sin instalar nada. Se puede copiar a
cualquier PC Windows (por USB o red).

### B.2 Instalador con accesos directos (recomendado para el usuario final)

1. Instala **Inno Setup 6** (gratis): https://jrsoftware.org/isdl.php  ← **hazlo antes**
2. Ejecuta **`build_installer.bat`** → produce **`Output\AgenteSAE-Setup.exe`**.

Ese `Setup.exe` instala la app, crea acceso directo en el **escritorio** y el
**menú inicio**, y agrega **desinstalador**. El equipo final no necesita Python.

> **Si aún no instalas Inno Setup**, el `.bat` te avisará (no falla la app): la
> aplicación igual queda lista y ejecutable en `dist_app\AgenteSAE\AgenteSAE.exe`.
> Y si solo quieres un único archivo portable, usa **`build_exe.bat`**, que **no
> requiere Inno Setup**.

### B.3 Compilación automática en la nube (sin Python ni compilar)

El repositorio incluye un flujo de **GitHub Actions**
(`.github/workflows/build-windows.yml`) que, en cada cambio, compila en un
runner Windows tanto el **portable** como el **instalador** y los publica como
*artifacts* descargables. Así nadie del equipo necesita instalar Python ni
PyInstaller: solo descargar el `.exe` desde la pestaña **Actions** del repo.
*(Queda operativo cuando se habilite el acceso de escritura al repositorio.)*

## Uso mes a mes (modo "carpeta por sociedad")

Ten una **carpeta por sociedad** y deja dentro, cada mes, los dos PDF y el Word:

```
IRCA\
   IRCA_2026.pdf                 (balance del período actual)
   IRCA_2025.pdf                 (balance del mismo mes del año anterior)
   NOTAS_IRCA_ABRIL_2026.docx    (Word de notas)
```

1. Abre AgenteSAE → **Abrir carpeta** y selecciona la carpeta de la sociedad.
   La app detecta sola cuál PDF es el período y cuál el comparativo (por el
   "Período" impreso en cada PDF) y cuál es el Word.
2. **Procesar** → revisa el panel de resultado y el informe (botón *Ver informe*).
   El Word actualizado (`..._ACTUALIZADO.docx`) y el informe quedan en la carpeta.

## Notas

- La **Nota 3** se deja intacta a propósito (se hace manual).
- Si una sociedad usa un template distinto, el informe avisa ("no se encontró la
  tabla …"); se ajusta `NOTE_CONFIG` en `agentesae/docx_updater.py`.
