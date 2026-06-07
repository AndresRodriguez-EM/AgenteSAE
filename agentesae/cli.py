"""
AgenteSAE — Actualización de NOTAS (Word) a partir de los auxiliares (PDF).

Uso:
    python -m agentesae.cli --actual AUX_2026.pdf --comparativo AUX_2025.pdf \
        --notas NOTAS.docx --salida NOTAS_ACTUALIZADO.docx [--informe informe.md]

Reutilizable para cualquier sociedad que use el mismo template de notas.
"""
from __future__ import annotations
import argparse
import datetime as _dt
import os

import docx

from .pdf_parser import parse_pdf
from .docx_updater import update_document, _patrimonio


# Cuadres: VALOR TOTAL de cada nota = total de la(s) clase(s) del auxiliar.
NOTA_TOTALES = {
    "1 Efectivo y equivalente": ["11"],
    "2 Cuentas por cobrar": ["13"],
    "3 Propiedad planta y equipo": ["15"],
    "4 Acreedores y otras CxP": ["2"],
    "6 Ingresos operacionales": ["41"],
    "7 Gastos de administración": ["51"],
    "8 Gastos no operacionales": ["53"],
    "9 Ingresos no operacionales": ["42"],
}


def _fmt(v):
    return f"{v:,.0f}"


def conciliacion(b26, b25):
    """Devuelve líneas de verificación de cuadre (auxiliar 2026 vs 2025)."""
    out = []
    eq = sum(b26.saldo(c) for c in "123456789")
    out.append(("Ecuación contable auxiliar 2026 (Σ clases = 0)", _fmt(eq), "OK" if abs(eq) < 1 else "REVISAR"))
    for nombre, clases in NOTA_TOTALES.items():
        v26 = sum(abs(b26.saldo(c)) for c in clases)
        v25 = sum(abs(b25.saldo(c)) for c in clases)
        out.append((f"Nota {nombre}", f"{_fmt(v26)} | {_fmt(v25)}", ""))
    p26 = _patrimonio(b26)[1]
    p25 = _patrimonio(b25)[1]
    out.append(("Nota 5 Patrimonio (total)", f"{_fmt(p26)} | {_fmt(p25)}", ""))
    return out


def generar_informe(rep, b26, b25, meta):
    L = []
    L.append(f"# Informe de actualización de NOTAS — {meta['sociedad']}")
    L.append("")
    L.append(f"- **Periodo (actual):** {b26.periodo}  ·  **Comparativo:** {b25.periodo}")
    L.append(f"- **Auxiliar actual:** `{meta['actual']}`")
    L.append(f"- **Auxiliar comparativo:** `{meta['comparativo']}`")
    L.append(f"- **Notas (entrada):** `{meta['notas']}`")
    L.append(f"- **Notas (salida):** `{meta['salida']}`")
    L.append(f"- **Generado:** {_dt.datetime.now():%Y-%m-%d %H:%M}")
    L.append("")
    L.append(f"**Resumen:** {len(rep.changes)} valores actualizados · "
             f"{len(rep.added)} terceros agregados · {len(rep.removed)} terceros eliminados · "
             f"{len(getattr(rep, 'added_tables', []))} tablas creadas · "
             f"{len(rep.flags)} observaciones.")
    L.append("")

    L.append("## Verificación de cuadre (auxiliar)")
    L.append("")
    L.append("| Concepto | Abril 2026 \\| Abril 2025 | Estado |")
    L.append("|---|---:|:--:|")
    for concepto, val, estado in conciliacion(b26, b25):
        L.append(f"| {concepto} | {val} | {estado} |")
    L.append("")

    if rep.changes:
        L.append("## Valores actualizados")
        L.append("")
        L.append("| Nota | Tabla | Concepto | Columna | Antes | Después |")
        L.append("|:--:|---|---|---|---:|---:|")
        for nota, tabla, concepto, col, antes, desp in rep.changes:
            L.append(f"| {nota} | {tabla} | {str(concepto).strip()} | {col} | {antes} | {desp} |")
        L.append("")

    if rep.added:
        L.append("## Terceros agregados (presentes en el auxiliar, ausentes en el Word)")
        L.append("")
        L.append("| Tabla | NIT | Tercero | Valores |")
        L.append("|---|---|---|---|")
        for tabla, nit, nombre, vals in rep.added:
            v = ", ".join(f"{k}={_fmt(x)}" for k, x in vals.items())
            L.append(f"| {tabla} | {nit} | {nombre} | {v} |")
        L.append("")

    if getattr(rep, "added_tables", None):
        L.append("## Tablas creadas (faltaban en el Word y se generaron del auxiliar)")
        L.append("")
        L.append("| Nota | Tabla | Terceros |")
        L.append("|:--:|---|:--:|")
        for nota, titulo, n in rep.added_tables:
            L.append(f"| {nota} | {titulo} | {n} |")
        L.append("")

    if rep.removed:
        L.append("## Terceros eliminados (sin saldo en ningún auxiliar)")
        L.append("")
        L.append("| Tabla | NIT | Tercero |")
        L.append("|---|---|---|")
        for tabla, nit, nombre in rep.removed:
            L.append(f"| {tabla} | {nit} | {nombre} |")
        L.append("")

    if rep.flags:
        L.append("## Observaciones (requieren revisión / decisión)")
        L.append("")
        for f in rep.flags:
            L.append(f"- {f}")
        L.append("")

    return "\n".join(L)


def run(actual, comparativo, notas, salida, informe=None, sociedad=None):
    b26 = parse_pdf(actual)
    b25 = parse_pdf(comparativo)
    doc = docx.Document(notas)
    rep = update_document(doc, b26, b25)
    os.makedirs(os.path.dirname(salida) or ".", exist_ok=True)
    doc.save(salida)
    meta = dict(actual=actual, comparativo=comparativo, notas=notas, salida=salida,
                sociedad=sociedad or os.path.basename(notas))
    texto = generar_informe(rep, b26, b25, meta)
    if informe:
        with open(informe, "w", encoding="utf-8") as fh:
            fh.write(texto)
    return rep, texto


def main():
    ap = argparse.ArgumentParser(description="Actualiza las NOTAS (Word) desde los auxiliares (PDF).")
    ap.add_argument("--actual", required=True, help="PDF del balance de comprobación del periodo")
    ap.add_argument("--comparativo", required=True, help="PDF del balance comparativo (año anterior)")
    ap.add_argument("--notas", required=True, help="Documento Word de notas a actualizar")
    ap.add_argument("--salida", required=True, help="Ruta del Word actualizado de salida")
    ap.add_argument("--informe", help="Ruta opcional para el informe (.md)")
    ap.add_argument("--sociedad", help="Nombre de la sociedad (para el informe)")
    a = ap.parse_args()
    rep, _ = run(a.actual, a.comparativo, a.notas, a.salida, a.informe, a.sociedad)
    print(f"OK -> {a.salida}")
    print(f"   {len(rep.changes)} cambios · {len(rep.added)} altas · {len(rep.removed)} bajas · {len(rep.flags)} observaciones")
    if a.informe:
        print(f"   Informe: {a.informe}")


if __name__ == "__main__":
    main()
