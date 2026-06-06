"""
AgenteSAE — Aplicación de escritorio para actualizar las NOTAS (Word) de cada
sociedad a partir de los auxiliares (Balance de Comprobación en PDF).

Interfaz gráfica (Tkinter) sobre el motor `agentesae`. No modifica la Nota 3
(se actualiza manualmente). Conserva intacto el formato del documento Word.

Ejecutar:   python app.py
Empacar:    ver INSTALL.md (PyInstaller -> ejecutable de escritorio)
"""
from __future__ import annotations
import glob
import json
import os
import re
import threading
import traceback
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

import docx

from agentesae.pdf_parser import parse_pdf
from agentesae.docx_updater import update_document
from agentesae.cli import generar_informe

SOCIEDADES = ["IRCA", "SANTA", "MONTOYA", "ZARLHA", "INVERMAP", "CIA"]
CONFIG_PATH = os.path.join(os.path.expanduser("~"), ".agentesae.json")


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


MESES = {"ENERO": 1, "FEBRERO": 2, "MARZO": 3, "ABRIL": 4, "MAYO": 5, "JUNIO": 6,
         "JULIO": 7, "AGOSTO": 8, "SEPTIEMBRE": 9, "OCTUBRE": 10, "NOVIEMBRE": 11, "DICIEMBRE": 12}


def _anio(periodo):
    m = re.search(r"(\d{4})", periodo or "")
    return int(m.group(1)) if m else 0


def _mes(periodo):
    for nombre, n in MESES.items():
        if nombre in (periodo or "").upper():
            return n
    return 0


class App(ttk.Frame):
    def __init__(self, master):
        super().__init__(master, padding=12)
        self.master = master
        self.cfg = _load_cfg()
        self.grid(sticky="nsew")
        master.columnconfigure(0, weight=1)
        master.rowconfigure(0, weight=1)
        self.columnconfigure(1, weight=1)

        self.v_carpeta = tk.StringVar()
        self.v_actual = tk.StringVar()
        self.v_comp = tk.StringVar()
        self.v_notas = tk.StringVar()
        self.v_salida = tk.StringVar()
        self.v_soc = tk.StringVar(value=SOCIEDADES[0])
        self.v_status = tk.StringVar(value="Listo.")

        self._build()

    # ------------------------------------------------------------------ UI
    def _build(self):
        r = 0
        ttk.Label(self, text="AgenteSAE — Actualización de Notas a los Estados Financieros",
                  font=("Segoe UI", 13, "bold")).grid(row=r, column=0, columnspan=3, sticky="w", pady=(0, 10))
        r += 1

        ttk.Label(self, text="Sociedad:").grid(row=r, column=0, sticky="w", pady=3)
        cb = ttk.Combobox(self, textvariable=self.v_soc, values=SOCIEDADES, width=20)
        cb.grid(row=r, column=1, sticky="w", pady=3)
        r += 1

        # Carpeta de la sociedad: autodetecta los PDFs y el Word
        ttk.Label(self, text="Carpeta de la sociedad:").grid(row=r, column=0, sticky="w", pady=3)
        ttk.Entry(self, textvariable=self.v_carpeta, width=64).grid(row=r, column=1, sticky="we", pady=3, padx=(0, 6))
        ttk.Button(self, text="Elegir carpeta…", command=self._pick_folder).grid(row=r, column=2, sticky="w", pady=3)
        r += 1
        ttk.Separator(self, orient="horizontal").grid(row=r, column=0, columnspan=3, sticky="we", pady=(2, 6))
        r += 1

        self._file_row(r, "Auxiliar del período (PDF):", self.v_actual,
                       [("PDF", "*.pdf")], self._on_pick_pdf); r += 1
        self._file_row(r, "Auxiliar comparativo año anterior (PDF):", self.v_comp,
                       [("PDF", "*.pdf")], self._on_pick_pdf); r += 1
        self._file_row(r, "Notas (Word .docx):", self.v_notas,
                       [("Word", "*.docx")], self._on_pick_notas); r += 1
        self._file_row(r, "Guardar Word actualizado en:", self.v_salida,
                       [("Word", "*.docx")], None, save=True); r += 1

        bar = ttk.Frame(self)
        bar.grid(row=r, column=0, columnspan=3, sticky="we", pady=(8, 6))
        self.btn = ttk.Button(bar, text="Procesar", command=self._procesar)
        self.btn.pack(side="left")
        ttk.Label(bar, textvariable=self.v_status, foreground="#0a6").pack(side="left", padx=10)
        r += 1

        ttk.Label(self, text="Resultado:").grid(row=r, column=0, sticky="w"); r += 1
        wrap = ttk.Frame(self)
        wrap.grid(row=r, column=0, columnspan=3, sticky="nsew")
        self.rowconfigure(r, weight=1)
        wrap.columnconfigure(0, weight=1); wrap.rowconfigure(0, weight=1)
        self.txt = tk.Text(wrap, height=18, wrap="word", font=("Consolas", 9))
        self.txt.grid(row=0, column=0, sticky="nsew")
        sb = ttk.Scrollbar(wrap, command=self.txt.yview)
        sb.grid(row=0, column=1, sticky="ns")
        self.txt.config(yscrollcommand=sb.set, state="disabled")

    def _file_row(self, r, label, var, types, on_pick, save=False):
        ttk.Label(self, text=label).grid(row=r, column=0, sticky="w", pady=3)
        ttk.Entry(self, textvariable=var, width=64).grid(row=r, column=1, sticky="we", pady=3, padx=(0, 6))

        def browse():
            init = self.cfg.get("last_dir", "")
            if save:
                p = filedialog.asksaveasfilename(defaultextension=".docx", filetypes=types, initialdir=init)
            else:
                p = filedialog.askopenfilename(filetypes=types, initialdir=init)
            if p:
                var.set(p)
                self.cfg["last_dir"] = os.path.dirname(p)
                _save_cfg(self.cfg)
                if on_pick:
                    on_pick()
        ttk.Button(self, text="Examinar…", command=browse).grid(row=r, column=2, sticky="w", pady=3)

    def _pick_folder(self):
        init = self.cfg.get("last_dir", "")
        d = filedialog.askdirectory(initialdir=init)
        if not d:
            return
        self.v_carpeta.set(d)
        self.cfg["last_dir"] = d
        _save_cfg(self.cfg)
        # detectar sociedad por el nombre de la carpeta
        up = os.path.basename(d).upper()
        for s in SOCIEDADES:
            if s in up:
                self.v_soc.set(s)
                break
        self._status("Detectando archivos…")
        threading.Thread(target=self._scan_folder, args=(d,), daemon=True).start()

    def _scan_folder(self, folder):
        pdfs = []
        for p in glob.glob(os.path.join(folder, "*.pdf")):
            try:
                b = parse_pdf(p)
                pdfs.append((_anio(b.periodo), _mes(b.periodo), p, b.periodo))
            except Exception:
                pass
        pdfs.sort(key=lambda x: (x[0], x[1]), reverse=True)
        actual = pdfs[0] if pdfs else None
        comp = None
        if actual:
            for cand in pdfs[1:]:                 # mismo mes, año anterior
                if cand[0] == actual[0] - 1 and cand[1] == actual[1]:
                    comp = cand
                    break
            if comp is None:                       # si no, el siguiente más antiguo
                comp = next((c for c in pdfs[1:] if c[0] < actual[0]), None)
            if comp is None and len(pdfs) > 1:
                comp = pdfs[1]
        docs = [d for d in glob.glob(os.path.join(folder, "*.docx"))
                if "ACTUALIZADO" not in os.path.basename(d).upper()
                and not os.path.basename(d).startswith("~$")]
        word = None
        if docs:
            pref = [d for d in docs if "NOTA" in os.path.basename(d).upper()]
            word = (pref or sorted(docs, key=os.path.getmtime, reverse=True))[0]
        self.master.after(0, lambda: self._apply_scan(actual, comp, word))

    def _apply_scan(self, actual, comp, word):
        msg = []
        if actual:
            self.v_actual.set(actual[2]); msg.append(f"actual={actual[3]}")
        if comp:
            self.v_comp.set(comp[2]); msg.append(f"comparativo={comp[3]}")
        if word:
            self.v_notas.set(word); self._on_pick_notas()
        faltan = [n for n, v in [("PDF período", actual), ("PDF comparativo", comp), ("Word", word)] if not v]
        if faltan:
            self._status("Detección parcial — falta: " + ", ".join(faltan) + ". Completa manualmente.")
        else:
            self._status("Archivos detectados: " + "  |  ".join(msg))

    def _on_pick_pdf(self):
        self._status("Auxiliar seleccionado.")

    def _on_pick_notas(self):
        # sugerir nombre de salida y sociedad
        p = self.v_notas.get()
        if p and not self.v_salida.get():
            base, ext = os.path.splitext(p)
            self.v_salida.set(base + "_ACTUALIZADO" + ext)
        up = os.path.basename(p).upper()
        for s in SOCIEDADES:
            if s in up:
                self.v_soc.set(s)
                break

    # -------------------------------------------------------------- proceso
    def _status(self, msg, color="#0a6"):
        self.v_status.set(msg)
        self.update_idletasks()

    def _log(self, text):
        self.txt.config(state="normal")
        self.txt.delete("1.0", "end")
        self.txt.insert("1.0", text)
        self.txt.config(state="disabled")

    def _procesar(self):
        actual, comp, notas = self.v_actual.get(), self.v_comp.get(), self.v_notas.get()
        for ruta, etq in [(actual, "auxiliar del período"), (comp, "auxiliar comparativo"), (notas, "Word de notas")]:
            if not ruta or not os.path.isfile(ruta):
                messagebox.showerror("Falta archivo", f"Selecciona un archivo válido para: {etq}.")
                return
        salida = self.v_salida.get() or os.path.splitext(notas)[0] + "_ACTUALIZADO.docx"
        self.v_salida.set(salida)
        self.btn.config(state="disabled")
        self._status("Procesando…")
        threading.Thread(target=self._run, args=(actual, comp, notas, salida), daemon=True).start()

    def _run(self, actual, comp, notas, salida):
        try:
            b_a = parse_pdf(actual)
            b_c = parse_pdf(comp)
            # Auto-corrección: el "actual" debe ser el período más reciente
            swap = ""
            if _anio(b_c.periodo) > _anio(b_a.periodo):
                b_a, b_c = b_c, b_a
                swap = "  (se intercambiaron: el comparativo era más reciente)"
            doc = docx.Document(notas)
            rep = update_document(doc, b_a, b_c)
            os.makedirs(os.path.dirname(salida) or ".", exist_ok=True)
            doc.save(salida)
            meta = dict(actual=actual, comparativo=comp, notas=notas, salida=salida,
                        sociedad=self.v_soc.get())
            informe = generar_informe(rep, b_a, b_c, meta)
            ruta_inf = os.path.splitext(salida)[0] + "_informe.md"
            with open(ruta_inf, "w", encoding="utf-8") as fh:
                fh.write(informe)
            resumen = (f"Período {b_a.periodo}  vs  {b_c.periodo}{swap}\n"
                       f"{len(rep.changes)} valores actualizados · {len(rep.added)} terceros agregados · "
                       f"{len(rep.removed)} eliminados · {len(rep.flags)} observaciones\n"
                       f"\nWord actualizado: {salida}\nInforme: {ruta_inf}\n"
                       + "=" * 70 + "\n\n" + informe)
            self.master.after(0, lambda: self._done(resumen, len(rep.flags)))
        except Exception:
            err = traceback.format_exc()
            self.master.after(0, lambda: self._error(err))

    def _done(self, resumen, n_flags):
        self._log(resumen)
        self._status(f"Listo ✓  ({n_flags} observaciones)")
        self.btn.config(state="normal")
        messagebox.showinfo("Completado", "Word actualizado y guardado correctamente.")

    def _error(self, err):
        self._log("ERROR:\n\n" + err)
        self._status("Error", color="#c00")
        self.btn.config(state="normal")
        messagebox.showerror("Error al procesar", "Ocurrió un error. Revisa el detalle en el panel de resultado.")


def main():
    root = tk.Tk()
    root.title("AgenteSAE")
    root.geometry("900x680")
    try:
        ttk.Style().theme_use("clam")
    except Exception:
        pass
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
