"""
Lógica de flujo independiente de la interfaz: detección automática de archivos
por carpeta de sociedad y procesamiento (parseo + actualización + informe).

La usan tanto la app de escritorio (Qt) como la versión Tkinter y el CLI.
"""
from __future__ import annotations
import glob
import os
import re

import docx

from .pdf_parser import parse_pdf
from .docx_updater import update_document
from .cli import generar_informe

SOCIEDADES = ["IRCA", "SANTA", "MONTOYA", "ZARLHA", "INVERMAP", "CIA"]

MESES = {"ENERO": 1, "FEBRERO": 2, "MARZO": 3, "ABRIL": 4, "MAYO": 5, "JUNIO": 6,
         "JULIO": 7, "AGOSTO": 8, "SEPTIEMBRE": 9, "OCTUBRE": 10, "NOVIEMBRE": 11, "DICIEMBRE": 12}


def anio(periodo: str) -> int:
    m = re.search(r"(\d{4})", periodo or "")
    return int(m.group(1)) if m else 0


def mes(periodo: str) -> int:
    for nombre, n in MESES.items():
        if nombre in (periodo or "").upper():
            return n
    return 0


_ALIAS = {"INVERMARP": "INVERMAP"}   # variantes de escritura en los archivos


def sociedad_de(nombre: str) -> str | None:
    up = (nombre or "").upper()
    for alias, s in _ALIAS.items():
        if alias in up:
            return s
    for s in SOCIEDADES:
        if s in up:
            return s
    return None


# Palabras de otros informes que NO son el balance de comprobación (auxiliar).
_NEG_PDF = ("BALANCE", "ESTADO", "RESULTADO", "INFORME", "GESTION", "OFICIO",
            "RADICAC", "EXTRACTO", "INDICADOR", "PLAN", "BIENES", "FLUJO",
            "CAMBIOS", "INGRESO", "AVANCE", "LIQUIDA", "PROPOSITO", "CONTRATO",
            "ARREND", "ARRIEND", "SITUACION", "NOTAS", "NOTA", "ANEXO")


def _score_pdf_name(path: str, company: str) -> int:
    """Qué tan probable es que un PDF sea el auxiliar (balance), por su nombre."""
    name = os.path.splitext(os.path.basename(path))[0].upper()
    toks = re.findall(r"[A-ZÁÉÍÓÚÑ0-9]+", name)
    s = 0
    if re.search(r"\b\d{4}\b", name):
        s += 2
    if company and company.upper() in name:
        s += 1
    if len([t for t in toks if not t.isdigit()]) <= 2:
        s += 2
    if re.fullmatch(r"[A-ZÁÉÍÓÚÑ .]*\d{4}", name.strip()):
        s += 3
    if any(w in name for w in _NEG_PDF):
        s -= 5
    return s


def _score_word_name(path: str, company: str) -> int:
    name = os.path.splitext(os.path.basename(path))[0].upper()
    s = 0
    if "NOTAS" in name or "NOTA" in name:
        s += 2
    if "ESTADOS FINANCIEROS" in name:
        s += 3
    if company and company.upper() in name:
        s += 1
    for w in ("BALANCE", "PROPOSITO", "OFICIO", "RADICAC", "GESTION", "ANEXO"):
        if w in name:
            s -= 3
    return s


def detectar_en_carpeta(folder: str) -> dict:
    """Detecta en una carpeta el PDF del período, el comparativo y el Word.

    Devuelve dict con rutas y períodos: actual, actual_periodo, comparativo,
    comparativo_periodo, notas, sociedad, faltan (lista de lo que no se encontró).
    """
    company = sociedad_de(os.path.basename(folder)) or ""

    # Parsear solo los PDF más probables (por nombre) y confirmar por contenido
    # que sean balances reales (>=15 cuentas y con período). Corte temprano.
    pdf_paths = sorted(glob.glob(os.path.join(folder, "*.pdf")),
                       key=lambda p: -_score_pdf_name(p, company))
    balances, years, parsed = [], set(), 0
    for p in pdf_paths:
        if _score_pdf_name(p, company) < 0 and balances:
            break
        try:
            b = parse_pdf(p)
        except Exception:
            b = None
        parsed += 1
        if b and len(b.cuentas) >= 15 and anio(b.periodo):
            balances.append((anio(b.periodo), mes(b.periodo), p, b.periodo))
            years.add(anio(b.periodo))
        if len(balances) >= 2 and len(years) >= 2:
            break
        if parsed >= 8:
            break
    balances.sort(key=lambda x: (x[0], x[1]), reverse=True)

    actual = balances[0] if balances else None
    comp = None
    if actual:
        for cand in balances[1:]:                   # mismo mes, año anterior
            if cand[0] == actual[0] - 1 and cand[1] == actual[1]:
                comp = cand
                break
        if comp is None:
            comp = next((c for c in balances[1:] if c[0] < actual[0]), None)
        if comp is None and len(balances) > 1:
            comp = balances[1]

    docs = [d for d in glob.glob(os.path.join(folder, "*.docx"))
            if "ACTUALIZADO" not in os.path.basename(d).upper()
            and not os.path.basename(d).startswith("~$")]
    word = max(docs, key=lambda d: _score_word_name(d, company)) if docs else None
    if word and _score_word_name(word, company) <= -3:
        word = None

    faltan = [etq for etq, v in [("PDF del período", actual),
                                 ("PDF comparativo", comp),
                                 ("Word de notas", word)] if not v]
    return dict(
        actual=actual[2] if actual else "",
        actual_periodo=actual[3] if actual else "",
        comparativo=comp[2] if comp else "",
        comparativo_periodo=comp[3] if comp else "",
        notas=word or "",
        sociedad=sociedad_de(os.path.basename(folder)) or "",
        faltan=faltan,
    )


def salida_sugerida(notas: str) -> str:
    base, ext = os.path.splitext(notas)
    return base + "_ACTUALIZADO" + (ext or ".docx")


def procesar(actual: str, comparativo: str, notas: str, salida: str, sociedad: str = "") -> dict:
    """Ejecuta la actualización y guarda el Word y el informe.

    Devuelve dict con: rep (Report), informe (texto), ruta_informe, salida,
    periodo_actual, periodo_comparativo, swap (bool).
    """
    b_a = parse_pdf(actual)
    b_c = parse_pdf(comparativo)
    swap = False
    if anio(b_c.periodo) > anio(b_a.periodo):          # asegurar que "actual" sea el más reciente
        b_a, b_c = b_c, b_a
        actual, comparativo = comparativo, actual
        swap = True

    doc = docx.Document(notas)
    rep = update_document(doc, b_a, b_c)

    salida = salida or salida_sugerida(notas)
    os.makedirs(os.path.dirname(salida) or ".", exist_ok=True)
    doc.save(salida)

    meta = dict(actual=actual, comparativo=comparativo, notas=notas, salida=salida,
                sociedad=sociedad or sociedad_de(os.path.basename(notas)) or os.path.basename(notas))
    informe = generar_informe(rep, b_a, b_c, meta)
    ruta_informe = os.path.splitext(salida)[0] + "_informe.md"
    with open(ruta_informe, "w", encoding="utf-8") as fh:
        fh.write(informe)

    return dict(rep=rep, informe=informe, ruta_informe=ruta_informe, salida=salida,
                periodo_actual=b_a.periodo, periodo_comparativo=b_c.periodo, swap=swap)
