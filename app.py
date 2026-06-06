"""
AgenteSAE — Aplicación de escritorio (interfaz profesional Qt / PySide6).

Actualiza las NOTAS (Word) de cada sociedad a partir de los auxiliares (PDF),
conservando el formato del documento y sin tocar la Nota 3.

Detección automática por carpeta: al elegir la carpeta de una sociedad, la app
identifica el PDF del período, el comparativo y el Word.

Ejecutar:   python app.py
Empacar:    ver INSTALL.md (PyInstaller -> .exe; Inno Setup -> instalador)
"""
from __future__ import annotations
import json
import os
import sys
import traceback

from PySide6 import QtCore, QtGui, QtWidgets

from agentesae import workflow

CONFIG_PATH = os.path.join(os.path.expanduser("~"), ".agentesae.json")
OK, NO = "✔", "—"


def resource_path(rel: str) -> str:
    """Ruta a un recurso, funcione en código o empaquetado con PyInstaller."""
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, rel)


def app_icon() -> "QtGui.QIcon":
    p = resource_path(os.path.join("assets", "icon.png"))
    return QtGui.QIcon(p) if os.path.exists(p) else QtGui.QIcon()


def _load_cfg():
    try:
        with open(CONFIG_PATH, encoding="utf-8") as fh:
            return json.load(fh)
    except Exception:
        return {}


def _save_cfg(cfg):
    try:
        with open(CONFIG_PATH, "w", encoding="utf-8") as fh:
            json.dump(cfg, fh)
    except Exception:
        pass


class Worker(QtCore.QObject):
    """Ejecuta una función en un hilo y emite el resultado."""
    done = QtCore.Signal(object)
    error = QtCore.Signal(str)

    def __init__(self, fn, *a, **kw):
        super().__init__()
        self.fn, self.a, self.kw = fn, a, kw

    @QtCore.Slot()
    def run(self):
        try:
            self.done.emit(self.fn(*self.a, **self.kw))
        except Exception:
            self.error.emit(traceback.format_exc())


class MainWindow(QtWidgets.QMainWindow):
    def __init__(self):
        super().__init__()
        self.cfg = _load_cfg()
        self.ruta_informe = ""
        self._threads = []
        self.setWindowTitle("AgenteSAE")
        self.setWindowIcon(app_icon())
        self.resize(960, 720)
        self._build_actions()
        self._build_menu_toolbar()
        self._build_central()
        self.statusBar().showMessage("Listo.")

    # --------------------------------------------------------------- UI
    def _build_actions(self):
        st = self.style()
        self.act_abrir = QtGui.QAction(st.standardIcon(QtWidgets.QStyle.SP_DirOpenIcon), "Abrir carpeta…", self)
        self.act_abrir.setShortcut("Ctrl+O")
        self.act_abrir.triggered.connect(self.abrir_carpeta)
        self.act_proc = QtGui.QAction(st.standardIcon(QtWidgets.QStyle.SP_MediaPlay), "Procesar", self)
        self.act_proc.setShortcut("Ctrl+R")
        self.act_proc.triggered.connect(self.procesar)
        self.act_inf = QtGui.QAction(st.standardIcon(QtWidgets.QStyle.SP_FileIcon), "Ver informe", self)
        self.act_inf.triggered.connect(self.ver_informe)
        self.act_inf.setEnabled(False)
        self.act_salir = QtGui.QAction("Salir", self); self.act_salir.triggered.connect(self.close)
        self.act_reset = QtGui.QAction("Restablecer campos", self); self.act_reset.triggered.connect(self.reset)
        self.act_about = QtGui.QAction("Acerca de…", self); self.act_about.triggered.connect(self.acerca)

    def _build_menu_toolbar(self):
        m = self.menuBar()
        arch = m.addMenu("&Archivo")
        arch.addActions([self.act_abrir, self.act_inf])
        arch.addSeparator(); arch.addAction(self.act_salir)
        m.addMenu("&Editar").addAction(self.act_reset)
        m.addMenu("A&yuda").addAction(self.act_about)

        tb = self.addToolBar("Principal")
        tb.setToolButtonStyle(QtCore.Qt.ToolButtonTextBesideIcon)
        tb.setMovable(False)
        tb.addAction(self.act_abrir); tb.addAction(self.act_proc); tb.addSeparator(); tb.addAction(self.act_inf)

    def _build_central(self):
        central = QtWidgets.QWidget(); self.setCentralWidget(central)
        outer = QtWidgets.QVBoxLayout(central)

        # --- fila superior: sociedad + período + carpeta ---
        top = QtWidgets.QGridLayout()
        top.addWidget(QtWidgets.QLabel("Sociedad:"), 0, 0)
        self.cmb = QtWidgets.QComboBox(); self.cmb.addItems(workflow.SOCIEDADES); self.cmb.setEditable(True)
        top.addWidget(self.cmb, 0, 1)
        self.lbl_periodo = QtWidgets.QLabel("Período: —")
        self.lbl_periodo.setStyleSheet("font-weight:600;")
        top.addWidget(self.lbl_periodo, 0, 2, 1, 2)
        top.addWidget(QtWidgets.QLabel("Carpeta de la sociedad:"), 1, 0)
        self.ed_carpeta = QtWidgets.QLineEdit()
        top.addWidget(self.ed_carpeta, 1, 1, 1, 2)
        b = QtWidgets.QPushButton("Examinar…"); b.clicked.connect(self.abrir_carpeta)
        top.addWidget(b, 1, 3)
        top.setColumnStretch(1, 1); top.setColumnStretch(2, 1)
        outer.addLayout(top)

        # --- columnas: archivos | resultado ---
        split = QtWidgets.QSplitter(QtCore.Qt.Horizontal)

        gb_files = QtWidgets.QGroupBox("Archivos")
        gl = QtWidgets.QGridLayout(gb_files)
        self.chk, self.eds = {}, {}
        for i, (key, etq) in enumerate([("actual", "PDF período"),
                                        ("comparativo", "PDF comparativo año anterior"),
                                        ("notas", "Notas (Word)")]):
            lbl = QtWidgets.QLabel(NO); lbl.setStyleSheet("color:#999;font-size:14px;")
            self.chk[key] = lbl
            ed = QtWidgets.QLineEdit(); ed.setReadOnly(False); self.eds[key] = ed
            btn = QtWidgets.QPushButton("…"); btn.setFixedWidth(32)
            btn.clicked.connect(lambda _=False, k=key: self.examinar(k))
            gl.addWidget(lbl, i, 0)
            gl.addWidget(QtWidgets.QLabel(etq), i, 1)
            gl.addWidget(ed, i, 2)
            gl.addWidget(btn, i, 3)
            gl.setColumnStretch(2, 1)
        gl.addWidget(QtWidgets.QLabel("Guardar en:"), 3, 1)
        self.ed_salida = QtWidgets.QLineEdit(); gl.addWidget(self.ed_salida, 3, 2)
        bs = QtWidgets.QPushButton("…"); bs.setFixedWidth(32); bs.clicked.connect(self.elegir_salida)
        gl.addWidget(bs, 3, 3)
        gl.setRowStretch(4, 1)
        split.addWidget(gb_files)

        gb_res = QtWidgets.QGroupBox("Resultado")
        rl = QtWidgets.QVBoxLayout(gb_res)
        self.lbl_resumen = QtWidgets.QLabel("Sin procesar todavía.")
        self.lbl_resumen.setStyleSheet("font-weight:600;")
        rl.addWidget(self.lbl_resumen)
        self.txt = QtWidgets.QTextBrowser(); self.txt.setOpenExternalLinks(True)
        rl.addWidget(self.txt, 1)
        split.addWidget(gb_res)
        split.setSizes([420, 540])
        outer.addWidget(split, 1)

        # --- botón procesar grande ---
        self.btn_proc = QtWidgets.QPushButton("▶  Procesar")
        self.btn_proc.setMinimumHeight(38)
        self.btn_proc.setStyleSheet("font-size:14px;font-weight:600;")
        self.btn_proc.clicked.connect(self.procesar)
        outer.addWidget(self.btn_proc)

    # ------------------------------------------------------------- helpers
    def _set_path(self, key, path):
        self.eds[key].setText(path or "")
        ok = bool(path) and os.path.isfile(path)
        self.chk[key].setText(OK if ok else NO)
        self.chk[key].setStyleSheet(("color:#16a34a;" if ok else "color:#999;") + "font-size:14px;")

    def _run_async(self, fn, on_done, *a):
        # on_done debe ser un método de esta ventana (hilo principal): así Qt usa
        # conexión en cola y la UI se actualiza en el hilo correcto.
        th = QtCore.QThread(self)
        wk = Worker(fn, *a)
        wk.moveToThread(th)
        th.started.connect(wk.run)
        wk.done.connect(on_done)
        wk.error.connect(self._on_error)
        wk.done.connect(th.quit)
        wk.error.connect(th.quit)
        th.finished.connect(wk.deleteLater)
        th.finished.connect(th.deleteLater)
        self._threads.append((th, wk))
        th.start()

    # ------------------------------------------------------------- acciones
    def abrir_carpeta(self):
        init = self.cfg.get("last_dir", "")
        d = QtWidgets.QFileDialog.getExistingDirectory(self, "Carpeta de la sociedad", init)
        if not d:
            return
        self.ed_carpeta.setText(d)
        self.cfg["last_dir"] = d; _save_cfg(self.cfg)
        soc = workflow.sociedad_de(os.path.basename(d))
        if soc:
            self.cmb.setCurrentText(soc)
        self.statusBar().showMessage("Detectando archivos…")
        self._run_async(workflow.detectar_en_carpeta, self._on_detected, d)

    def _on_detected(self, d):
        self._set_path("actual", d["actual"])
        self._set_path("comparativo", d["comparativo"])
        self._set_path("notas", d["notas"])
        if d["notas"]:
            self.ed_salida.setText(workflow.salida_sugerida(d["notas"]))
        if d["sociedad"]:
            self.cmb.setCurrentText(d["sociedad"])
        per = []
        if d["actual_periodo"]:
            per.append(d["actual_periodo"])
        if d["comparativo_periodo"]:
            per.append("vs " + d["comparativo_periodo"])
        self.lbl_periodo.setText("Período: " + (" ".join(per) if per else "—"))
        if d["faltan"]:
            self.statusBar().showMessage("Detección parcial — falta: " + ", ".join(d["faltan"]) + ". Complétalo manualmente.")
        else:
            self.statusBar().showMessage("Archivos detectados correctamente.")

    def examinar(self, key):
        init = self.cfg.get("last_dir", "")
        filt = "Word (*.docx)" if key == "notas" else "PDF (*.pdf)"
        p, _ = QtWidgets.QFileDialog.getOpenFileName(self, "Seleccionar archivo", init, filt)
        if p:
            self._set_path(key, p)
            self.cfg["last_dir"] = os.path.dirname(p); _save_cfg(self.cfg)
            if key == "notas" and not self.ed_salida.text():
                self.ed_salida.setText(workflow.salida_sugerida(p))

    def elegir_salida(self):
        p, _ = QtWidgets.QFileDialog.getSaveFileName(self, "Guardar Word actualizado", self.ed_salida.text(), "Word (*.docx)")
        if p:
            self.ed_salida.setText(p)

    def reset(self):
        for k in self.eds:
            self._set_path(k, "")
        self.ed_carpeta.clear(); self.ed_salida.clear()
        self.lbl_periodo.setText("Período: —")
        self.lbl_resumen.setText("Sin procesar todavía."); self.txt.clear()
        self.act_inf.setEnabled(False); self.statusBar().showMessage("Campos restablecidos.")

    def procesar(self):
        actual, comp, notas = self.eds["actual"].text(), self.eds["comparativo"].text(), self.eds["notas"].text()
        for ruta, etq in [(actual, "PDF del período"), (comp, "PDF comparativo"), (notas, "Word de notas")]:
            if not ruta or not os.path.isfile(ruta):
                QtWidgets.QMessageBox.critical(self, "Falta archivo", f"Selecciona un archivo válido para: {etq}.")
                return
        salida = self.ed_salida.text() or workflow.salida_sugerida(notas)
        self.ed_salida.setText(salida)
        self.btn_proc.setEnabled(False); self.act_proc.setEnabled(False)
        self.statusBar().showMessage("Procesando…")
        self._run_async(workflow.procesar, self._on_done, actual, comp, notas, salida, self.cmb.currentText())

    def _on_done(self, res):
        rep = res["rep"]
        swap = "  (se intercambiaron actual/comparativo)" if res["swap"] else ""
        self.lbl_resumen.setText(
            f"{res['periodo_actual']} vs {res['periodo_comparativo']}{swap}  ·  "
            f"{len(rep.changes)} cambios · {len(rep.added)} altas · {len(rep.removed)} bajas · {len(rep.flags)} observaciones")
        self.txt.setMarkdown(res["informe"])
        self.ruta_informe = res["ruta_informe"]
        self.act_inf.setEnabled(True)
        self.btn_proc.setEnabled(True); self.act_proc.setEnabled(True)
        self.statusBar().showMessage("Listo ✓  Word guardado: " + os.path.basename(res["salida"]))
        QtWidgets.QMessageBox.information(self, "Completado",
                                         "Word actualizado y guardado:\n" + res["salida"])

    def _on_error(self, err):
        self.btn_proc.setEnabled(True); self.act_proc.setEnabled(True)
        self.statusBar().showMessage("Error.")
        self.txt.setPlainText("ERROR:\n\n" + err)
        QtWidgets.QMessageBox.critical(self, "Error al procesar",
                                       "Ocurrió un error. Revisa el detalle en el panel de resultado.")

    def ver_informe(self):
        if self.ruta_informe and os.path.isfile(self.ruta_informe):
            QtGui.QDesktopServices.openUrl(QtCore.QUrl.fromLocalFile(self.ruta_informe))

    def acerca(self):
        box = QtWidgets.QMessageBox(self)
        box.setWindowTitle("Acerca de AgenteSAE")
        box.setText("<b>AgenteSAE</b>")
        box.setInformativeText(
            "Actualiza las Notas a los Estados Financieros (Word) de cada sociedad "
            "a partir de los auxiliares (Balance de Comprobación en PDF), conservando "
            "el formato y sin modificar la Nota 3.")
        pix = QtGui.QPixmap(resource_path(os.path.join("assets", "icon.png")))
        if not pix.isNull():
            box.setIconPixmap(pix.scaled(72, 72, QtCore.Qt.KeepAspectRatio, QtCore.Qt.SmoothTransformation))
        box.exec()


def main():
    app = QtWidgets.QApplication(sys.argv)
    app.setStyle("Fusion")
    app.setWindowIcon(app_icon())
    w = MainWindow()
    w.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
