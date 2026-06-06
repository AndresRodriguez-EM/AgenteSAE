"""
Actualizador de las NOTAS (Word) a partir de los auxiliares (Balance de
Comprobación en PDF), preservando el formato del documento.

Reglas:
  * Solo se modifican CELDAS DE TABLA (nunca párrafos narrativos ni estilos).
  * Columna "saldo/acumulado mes actual"  -> Nuevo Saldo del auxiliar del periodo.
  * Columna "... 2025" (mes anterior/año) -> Nuevo Saldo del auxiliar comparativo.
  * Columna "mov / movimiento mes"        -> |Débito - Crédito| del periodo.
  * Terceros: se AGREGAN los que estén en el auxiliar y falten; se ELIMINAN las
    filas cuyo tercero no exista en ningún auxiliar (ni actual ni comparativo).
  * Patrimonio (Nota 5) usa convención de signo propia (aportes en positivo,
    pérdidas en negativo).

El emparejamiento Word<->auxiliar se hace por NIT y, cuando un NIT aparece en
varias cuentas del alcance, se desempata por la DES CUENTA del Word.
"""
from __future__ import annotations
import copy
import math
import re

from .pdf_parser import Balance

# --------------------------------------------------------------------------- #
#  Utilidades de formato
# --------------------------------------------------------------------------- #

def parse_money(text: str):
    t = (text or "").replace("$", "").replace(".", "").replace(" ", "")
    t = t.replace(",", "")
    if not re.search(r"\d", t):
        return None
    neg = "-" in (text or "")
    t = t.replace("-", "")
    try:
        v = float(t)
        return -v if neg else v
    except ValueError:
        return None


def _round(value: float) -> int:
    return int(math.floor(abs(value) + 0.5)) * (-1 if value < 0 else 1)


def format_money(value: float, dollar: bool = False) -> str:
    n = _round(value)
    s = f"{abs(n):,}"
    if dollar:
        s = "$ " + s
    if n < 0:
        s = "-" + s
    return s


def format_like(value: float, cell) -> str:
    """Formatea preservando el estilo de la celda: símbolo '$' y separador de
    miles (algunas celdas usan punto en vez de coma)."""
    txt = cell.text
    s = format_money(value, "$" in txt)
    if re.search(r"\d\.\d{3}", txt) and "," not in txt:
        s = s.replace(",", ".")
    return s


def _toks(s: str) -> set[str]:
    return set(re.findall(r"[A-ZÁÉÍÓÚÑ]{3,}", (s or "").upper()))


def _norm_nit(s: str) -> str:
    return re.sub(r"\D", "", s or "")


def _detect_nit_sep(sample: str) -> str:
    if "." in (sample or ""):
        return "."
    return ","


def _fmt_nit(nit: str, sep: str) -> str:
    nit = _norm_nit(nit)
    return f"{int(nit):,}".replace(",", sep) if nit else nit


# --------------------------------------------------------------------------- #
#  Acceso a celdas preservando formato
# --------------------------------------------------------------------------- #

def set_cell_value(cell, text: str):
    """Escribe `text` en la celda conservando el formato del primer run."""
    target = None
    for p in cell.paragraphs:
        if any(ch.isdigit() for ch in p.text):
            target = p
            break
    if target is None:
        target = cell.paragraphs[-1] if cell.paragraphs else cell.add_paragraph()
    if target.runs:
        target.runs[0].text = text
        for r in target.runs[1:]:
            r.text = ""
    else:
        target.add_run(text)


def cell_has_dollar(cell) -> bool:
    return "$" in cell.text


def _unique_grid(row):
    """Devuelve [(grid_idx, cell)] saltando celdas combinadas repetidas."""
    out, seen = [], set()
    for i, c in enumerate(row.cells):
        if id(c._tc) in seen:
            continue
        seen.add(id(c._tc))
        out.append((i, c))
    return out


# --------------------------------------------------------------------------- #
#  Detección de roles de columnas a partir del encabezado
# --------------------------------------------------------------------------- #

def _role_of(header: str):
    """Clasifica el encabezado de una columna de valor. Tolera variantes:
    'SALDO MES ACTUAL', 'ACUMULADO MES (ACTUAL)', 'MOV(IMIENTO) MES',
    'SALDO/ACUMULADO MES ANTERIOR 2025', 'ACUMULADO AÑO ANTERIOR 2025', 'SALDO AÑO 2025'."""
    h = (header or "").upper()
    if "MOV" in h:
        return "mov"
    comparativo = ("2025" in h) or ("ANTERIOR" in h) or bool(re.search(r"\bA[NÑ]O\b", h))
    if "ACUMULAD" in h:
        return "comparativo" if comparativo else "acum_actual"
    if "SALDO" in h:
        if comparativo:
            return "comparativo"
        if "ACTUAL" in h:
            return "actual"
    return None


def _find_header_row(table):
    for r, row in enumerate(table.rows):
        texts = [c.text.upper() for c in row.cells]
        if any("TERCERO" in t for t in texts) or any(_role_of(t) for t in texts):
            if any(_role_of(t) for t in texts):
                return r
    return None


def _col_index(row, *keywords):
    for i, c in enumerate(row.cells):
        u = c.text.upper()
        if any(k in u for k in keywords):
            return i
    return None


# --------------------------------------------------------------------------- #
#  Registros de alcance (auxiliar) y emparejamiento
# --------------------------------------------------------------------------- #

def _anc_names(bal: Balance, code: str):
    names = []
    for L in (2, 4, 6, 8, 10):
        pre = code[:L]
        if pre in bal.cuentas:
            names.append(bal.cuentas[pre].nombre)
    return names


def scope_records(bal: Balance, prefixes):
    """Construye registros del alcance: a nivel de TERCERO (kind='t') y a nivel de
    CUENTA de 4/6 dígitos (kind='acc'). Las cuentas permiten emparejar filas que
    el Word presenta consolidadas/netas (p. ej. retención, IVA, impuestos)."""
    in_scope = [c for c in bal.cuentas.values() if any(c.codigo.startswith(p) for p in prefixes)]
    recs = []
    for c in in_scope:
        if c.terceros:
            des = _anc_names(bal, c.codigo) + [c.nombre]
            for t in c.terceros:
                recs.append(dict(kind="t", nit=t.nit, code=c.codigo, value=t.nuevo_saldo,
                                 mov=abs(t.debito - t.credito), des=des, t=t, tid=id(t)))
    for c in in_scope:
        if len(c.codigo) in (4, 6):
            children = {id(t) for cc in bal.cuentas.values()
                        if cc.codigo.startswith(c.codigo) and cc.terceros for t in cc.terceros}
            recs.append(dict(kind="acc", nit=None, code=c.codigo, value=c.nuevo_saldo,
                             mov=abs(c.debito - c.credito),
                             des=_anc_names(bal, c.codigo) + [c.nombre], children=children))
    return recs


def _pick(cands, des, hint):
    dw = _toks(des)
    ov = lambda r: max((len(dw & _toks(x)) for x in r["des"]), default=0)
    best = max(ov(r) for r in cands)
    top = [r for r in cands if ov(r) == best]
    if len(top) == 1 or not hint:
        return top[0]
    return min(top, key=lambda r: abs(abs(r["value"]) - abs(hint)))


def match_record(recs, nit, des, hint=None, allowed_nits=None, tname=None):
    """Empareja una fila del Word con un registro del auxiliar. Prioridad:
    1) tercero del mismo NIT con valor exacto; 2) cuenta con valor exacto;
    3) tercero del mismo NIT por descripción/proximidad; 4) cuenta por descripción;
    5) tercero por NOMBRE (tolera errores de digitación en el NIT).

    El emparejamiento a nivel de CUENTA (filas que el Word presenta consolidadas)
    solo se permite si el NIT de la fila corresponde a un tercero real del
    auxiliar; así una fila cuyo tercero ya no existe queda sin match (se elimina)."""
    tc = [r for r in recs if r["kind"] == "t" and r["nit"] == nit]
    ac = [r for r in recs if r["kind"] == "acc"]
    allow_acc = allowed_nits is None or nit in allowed_nits
    rh = _round(abs(hint)) if hint else None
    if rh is not None:
        et = [r for r in tc if _round(abs(r["value"])) == rh]
        if et:
            return _pick(et, des, hint)
        if allow_acc:
            ea = [r for r in ac if _round(abs(r["value"])) == rh]
            if ea:
                return _pick(ea, des, hint)
    if tc:
        return _pick(tc, des, hint)
    if allow_acc:
        dw = _toks(des)
        ad = [r for r in ac if any(len(dw & _toks(x)) >= 1 for x in r["des"])]
        if ad:
            return _pick(ad, des, hint)
    if tname:
        tn = _toks(tname)
        nt = [r for r in recs if r["kind"] == "t" and len(tn & _toks(r["t"].nombre)) >= 2]
        if nt:
            return _pick(nt, des, hint)
    return None


def find_same25(recs25, r26):
    """Registro 2025 con la MISMA identidad (mismo nivel y cuenta) que r26."""
    if r26["kind"] == "t":
        for r in recs25:
            if r["kind"] == "t" and r["nit"] == r26["nit"] and r["code"] == r26["code"]:
                return r
    else:
        for r in recs25:
            if r["kind"] == "acc" and r["code"] == r26["code"]:
                return r
    return None


def _val(role, r26, r25):
    if role == "comparativo":
        return abs(r25["value"]) if r25 else 0.0
    if role == "mov":
        return r26["mov"] if r26 else 0.0
    return abs(r26["value"]) if r26 else 0.0


# Tolerancia de redondeo (los auxiliares traen 2 decimales; el Word muestra
# enteros y el redondeo manual no siempre es consistente). Cambios <= esto se
# consideran inmateriales y no se aplican.
ROUND_TOL = 1


def _accept_set(recs_year, prim, nit, role, famcode):
    """Conjunto de cifras VÁLIDAS del auxiliar para una celda (a nivel tercero o
    a nivel de cuenta de la misma familia). Si el valor actual del Word ya está
    en este conjunto, la celda es correcta y no se modifica ('do no harm')."""
    out = set()
    for r in recs_year:
        same_t = r["kind"] == "t" and r["nit"] == nit
        same_a = r["kind"] == "acc" and famcode and famcode.startswith(r["code"])
        if same_t or same_a:
            out.add(_round(r["mov"] if role == "mov" else abs(r["value"])))
    if prim:
        out.add(_round(prim["mov"] if role == "mov" else abs(prim["value"])))
    return out


# --------------------------------------------------------------------------- #
#  Procesadores por tipo de tabla
# --------------------------------------------------------------------------- #

class Report:
    def __init__(self):
        self.changes = []   # (nota, tabla, concepto, columna, antes, despues)
        self.added = []     # (tabla, nit, nombre, valores)
        self.removed = []   # (tabla, nit, nombre)
        self.flags = []     # observaciones

    def chg(self, *a):
        self.changes.append(a)


def _data_rows(table, header_r, col_nit):
    """Filas de datos (con NIT) y fila de total (sin NIT, con números)."""
    data, total = [], None
    for r in range(header_r + 1, len(table.rows)):
        row = table.rows[r]
        nit = _norm_nit(row.cells[col_nit].text) if col_nit is not None else ""
        money = any(parse_money(c.text) is not None for _, c in _unique_grid(row))
        if nit:
            data.append(r)
        elif money and total is None:
            total = r
    return data, total


def process_terceros(table, cfg, b26, b25, rep: Report):
    nota = cfg.get("nota", "")
    header_r = _find_header_row(table)
    hdr = table.rows[header_r]
    roles = {}
    for i, c in enumerate(hdr.cells):
        role = _role_of(c.text)
        if role and i not in roles:
            roles[i] = role
    col_nit = _col_index(hdr, "CEDULA", "CÉDULA", "ID,")
    col_ter = _col_index(hdr, "TERCERO")
    col_des = _col_index(hdr, "DES CUENTA", "DESCRIPCION", "DES CUE")

    recs26 = scope_records(b26, cfg["scope"])
    recs25 = scope_records(b25, cfg["scope"])
    allowed = {r["nit"] for r in recs26 + recs25 if r["kind"] == "t"}
    matched_tids = set()
    covered = set()

    data_idx, total_idx = _data_rows(table, header_r, col_nit)
    nit_sep = _detect_nit_sep(table.rows[data_idx[0]].cells[col_nit].text) if data_idx and col_nit is not None else ","
    dollar = cell_has_dollar(table.rows[data_idx[0]].cells[list(roles)[0]]) if data_idx and roles else True
    col_actual = next((c for c, r in roles.items() if r in ("actual", "acum_actual")), None)
    col_comp = next((c for c, r in roles.items() if r == "comparativo"), None)

    to_remove = []
    for r in data_idx:
        row = table.rows[r]
        nit = _norm_nit(row.cells[col_nit].text) if col_nit is not None else ""
        des = row.cells[col_des].text if col_des is not None else ""
        tname = row.cells[col_ter].text if col_ter is not None else ""
        hint26 = parse_money(row.cells[col_actual].text) if col_actual is not None else None
        hint25 = parse_money(row.cells[col_comp].text) if col_comp is not None else None
        r26 = match_record(recs26, nit, des, hint26, allowed, tname)
        r25 = find_same25(recs25, r26) if r26 else match_record(recs25, nit, des, hint25, allowed, tname)
        if r26 is None and r25 is None:
            to_remove.append(r)
            rep.removed.append((cfg["sig"], nit, row.cells[col_ter].text if col_ter is not None else ""))
            continue
        # Tercero emparejado por nombre con NIT distinto -> posible error de digitación
        mref = r26 or r25
        if mref and mref["kind"] == "t" and nit and mref["nit"] != nit:
            rep.flags.append(
                f"Nota {nota} · {cfg['sig']}: NIT del tercero '{tname.strip()}' difiere "
                f"(Word {row.cells[col_nit].text.strip()} vs auxiliar {mref['nit']}). Se conservó el del Word.")
        if r26:
            if r26["kind"] == "t":
                matched_tids.add(r26["tid"])
            else:
                covered.add(r26["code"])
                matched_tids |= r26["children"]
        famcode = (r26 or r25 or {}).get("code")
        for col, role in roles.items():
            cell = row.cells[col]
            new = _val(role, r26, r25)
            existing = parse_money(cell.text)
            year = recs25 if role == "comparativo" else recs26
            prim = r25 if role == "comparativo" else r26
            acc = _accept_set(year, prim, nit, role, famcode)
            if existing is not None and (_round(existing) in acc or abs(existing - new) <= ROUND_TOL):
                continue  # ya es correcto (a nivel tercero o cuenta) o es redondeo
            nuevo = format_like(new, cell)
            antes = cell.text.strip()
            if antes != nuevo:
                set_cell_value(cell, nuevo)
                rep.chg(nota, cfg["sig"], (row.cells[col_ter].text if col_ter is not None else nit), role, antes, nuevo)

    for r in sorted(to_remove, reverse=True):
        tr = table.rows[r]._tr
        tr.getparent().remove(tr)

    # Altas: terceros del auxiliar 2026 no emparejados y no cubiertos por una cuenta neta
    if data_idx and col_nit is not None:
        from docx.table import _Row
        for rec in recs26:
            if rec["kind"] != "t" or rec["tid"] in matched_tids:
                continue
            if any(rec["code"].startswith(cov) for cov in covered):
                continue
            if abs(rec["value"]) < 1 and rec["mov"] < 1:
                continue
            cur = _data_rows(table, header_r, col_nit)[0]
            des_new = rec["t"].nombre
            for ri in cur:
                rr = match_record(recs26, _norm_nit(table.rows[ri].cells[col_nit].text),
                                  table.rows[ri].cells[col_des].text if col_des is not None else "",
                                  parse_money(table.rows[ri].cells[col_actual].text) if col_actual is not None else None,
                                  allowed)
                if rr and rr["kind"] == "t" and rr["code"][:4] == rec["code"][:4] and col_des is not None:
                    des_new = table.rows[ri].cells[col_des].text.strip()
                    break
            r25 = find_same25(recs25, rec) or match_record(recs25, rec["nit"], des_new, allowed_nits=allowed)
            src = table.rows[cur[-1]]
            new_tr = copy.deepcopy(src._tr)
            src._tr.addnext(new_tr)
            nrow = _Row(new_tr, table)
            set_cell_value(nrow.cells[col_nit], _fmt_nit(rec["nit"], nit_sep))
            if col_ter is not None:
                set_cell_value(nrow.cells[col_ter], rec["t"].nombre)
            if col_des is not None:
                set_cell_value(nrow.cells[col_des], des_new)
            vals = {}
            for col, role in roles.items():
                v = _val(role, rec, r25)
                set_cell_value(nrow.cells[col], format_like(v, nrow.cells[col]))
                vals[role] = v
            rep.added.append((cfg["sig"], rec["nit"], rec["t"].nombre, vals))

    # Recalcular fila de total = suma de filas de datos mostradas
    data_idx, total_idx = _data_rows(table, header_r, col_nit)
    if total_idx is not None:
        trow = table.rows[total_idx]
        for col, role in roles.items():
            s = 0.0
            for r in data_idx:
                v = parse_money(table.rows[r].cells[col].text)
                if v is not None:
                    s += v
            cell = trow.cells[col]
            existing = parse_money(cell.text)
            if existing is not None and abs(existing - s) <= ROUND_TOL:
                continue
            antes = cell.text.strip()
            nuevo = format_like(s, cell)
            if antes != nuevo:
                set_cell_value(cell, nuevo)
                rep.chg(nota, cfg["sig"], "TOTAL", role, antes, nuevo)


def process_cuenta(table, cfg, b26, b25, rep: Report):
    nota = cfg.get("nota", "")
    header_r = _find_header_row(table)
    hdr = table.rows[header_r]
    roles = {i: _role_of(c.text) for i, c in enumerate(hdr.cells) if _role_of(c.text)}
    v26 = abs(b26.saldo(cfg["account"]))
    v25 = abs(b25.saldo(cfg["account"]))
    for r in range(header_r + 1, len(table.rows)):
        row = table.rows[r]
        if not any(parse_money(c.text) is not None for _, c in _unique_grid(row)):
            continue
        for col, role in roles.items():
            val = v25 if role == "comparativo" else v26
            cell = row.cells[col]
            existing = parse_money(cell.text)
            if existing is not None and abs(existing - val) <= ROUND_TOL:
                continue
            antes = cell.text.strip()
            nuevo = format_like(val, cell)
            if antes != nuevo:
                set_cell_value(cell, nuevo)
                rep.chg(nota, cfg["sig"], row.cells[0].text.strip() or "TOTAL", role, antes, nuevo)


def _patrimonio(b: Balance):
    cap = abs(b.saldo("31"))
    reval = abs(b.saldo("34"))
    ejercicio = -(b.saldo("4") + b.saldo("5"))
    perdidas = -(b.saldo("36") + b.saldo("37"))
    valoriz = abs(b.saldo("38"))
    lines = dict(cap=cap, reval=reval, ejercicio=ejercicio, perdidas=perdidas, valoriz=valoriz)
    total = sum(_round(v) for v in (cap, reval, ejercicio, perdidas, valoriz))
    return lines, total


def process_patrimonio(table, cfg, b26, b25, rep: Report):
    l26, t26 = _patrimonio(b26)
    l25, t25 = _patrimonio(b25)
    keymap = [
        ("CAPITAL SOCIAL", "cap"),
        ("REVALORIZ", "reval"),
        ("EJERCICIO", "ejercicio"),
        ("ACUMULADAS", "perdidas"),
        ("VALORIZ", "valoriz"),   # se evalúa después de REVALORIZ
        ("TOTAL", "total"),
    ]
    for row in table.rows:
        label = row.cells[0].text.upper()
        key = None
        for kw, k in keymap:
            if kw in label:
                key = k
                if kw == "VALORIZ" and "REVALORIZ" in label:
                    key = "reval"
                break
        if key is None:
            continue
        cells = _unique_grid(row)
        if len(cells) < 3:
            continue
        col26 = cells[1][1]
        col25 = cells[2][1]
        v26 = t26 if key == "total" else l26[key]
        v25 = t25 if key == "total" else l25[key]
        for cell, val in ((col26, v26), (col25, v25)):
            existing = parse_money(cell.text)
            if existing is not None and abs(existing - val) <= 2:
                continue  # redondeo de partidas sumadas
            antes = cell.text.strip()
            nuevo = format_like(val, cell)
            if antes != nuevo:
                set_cell_value(cell, nuevo)
                rep.chg(cfg.get("nota", "5"), cfg["sig"], label.title(), "col", antes, nuevo)


def process_vtotal(table, cfg, b26, b25, rep: Report):
    if cfg.get("patrimonio"):
        _, t26 = _patrimonio(b26)
        _, t25 = _patrimonio(b25)
        v26, v25 = t26, t25
    else:
        v26 = abs(b26.saldo(cfg["account"]))
        v25 = abs(b25.saldo(cfg["account"]))
    for row in table.rows:
        cells = _unique_grid(row)
        moneycells = [(i, c) for (i, c) in cells if parse_money(c.text) is not None]
        if not moneycells:
            continue
        first = moneycells[0][1]
        last = moneycells[-1][1]
        for cell, val in ((first, v26), (last, v25)):
            existing = parse_money(cell.text)
            if existing is not None and abs(existing - val) <= 2:
                continue  # redondeo
            antes = cell.text.strip()
            nuevo = format_like(val, cell)
            if antes != nuevo:
                set_cell_value(cell, nuevo)
                rep.chg(cfg.get("nota", ""), cfg["sig"], "VALOR TOTAL", "col", antes, nuevo)
        break  # solo la primera fila de valores


# --------------------------------------------------------------------------- #
#  Configuración del template (índices estables, validados por firma)
# --------------------------------------------------------------------------- #

TABLES = [
    dict(idx=11, kind="cuenta",     nota="1", sig="EFECTIVO Y EQUIVALENTE", account="1105"),
    dict(idx=12, kind="vtotal",     nota="1", sig="VALOR TOTAL NOTA 1", account="11"),
    dict(idx=13, kind="terceros",   nota="2", sig="DETALLE CUENTAS POR COBRAR", scope=["1345"]),
    dict(idx=14, kind="terceros",   nota="2", sig="ANTICIPO DE IMPUESTOS", scope=["1355"]),
    dict(idx=15, kind="vtotal",     nota="2", sig="VALOR TOTAL NOTA 2", account="13"),
    dict(idx=16, kind="cuenta",     nota="3", sig="INMUEBLES", account="15"),
    dict(idx=17, kind="vtotal",     nota="3", sig="VALOR TOTAL NOTA 3", account="15"),
    dict(idx=18, kind="terceros",   nota="4", sig="CUENTAS POR PAGAR", scope=["2335", "2365"]),
    dict(idx=20, kind="terceros",   nota="4", sig="MPUESTOS POR PAGAR",
         scope=["2404", "2405", "2408", "2412", "2416"]),
    dict(idx=21, kind="terceros",   nota="4", sig="IMPUESTOS MULTAS Y SANCIONES",
         scope=["2615", "2635"]),
    dict(idx=23, kind="vtotal",     nota="4", sig="VALOR TOTAL NOTA  4", account="2"),
    dict(idx=25, kind="patrimonio", nota="5", sig="PATRIMONIO"),
    dict(idx=26, kind="vtotal",     nota="5", sig="VALOR TOTAL NOTA  5", patrimonio=True),
    dict(idx=28, kind="terceros",   nota="6", sig="INGRESOS POR ARRENDAMIENTOS", scope=["4155"]),
    dict(idx=30, kind="vtotal",     nota="6", sig="VALOR TOTAL NOTA 6", account="41"),
    dict(idx=32, kind="terceros",   nota="7", sig="HONORARIOS", scope=["5110"]),
    dict(idx=33, kind="terceros",   nota="7", sig="IMPUESTOS", scope=["5115"]),
    dict(idx=34, kind="terceros",   nota="7", sig="GASTOS LEGALES", scope=["5140"]),
    dict(idx=35, kind="terceros",   nota="7", sig="SERVICIOS", scope=["5135"]),
    dict(idx=36, kind="terceros",   nota="7", sig="MANTENIMIENTO", scope=["5145"]),
    dict(idx=37, kind="terceros",   nota="7", sig="DIVERSOS", scope=["5195"]),
    dict(idx=38, kind="vtotal",     nota="7", sig="VALOR TOTAL NOTA 7", account="51"),
    dict(idx=40, kind="terceros",   nota="8", sig="GASTOS EXTRAORDINARIOS", scope=["5305"]),
    dict(idx=41, kind="terceros",   nota="8", sig="GASTOS EXTRAORDINARIOS", scope=["5315"]),
    dict(idx=42, kind="vtotal",     nota="8", sig="VALOR TOTAL NOTA 8", account="53"),
    dict(idx=44, kind="terceros",   nota="9", sig="OTROS INGRESOS", scope=["4295"]),
    dict(idx=45, kind="vtotal",     nota="9", sig="VALOR TOTAL NOTA 9", account="42"),
]

_PROCESSORS = {
    "terceros": process_terceros,
    "cuenta": process_cuenta,
    "patrimonio": process_patrimonio,
    "vtotal": process_vtotal,
}


def update_document(doc, b26: Balance, b25: Balance) -> Report:
    rep = Report()
    for cfg in TABLES:
        table = doc.tables[cfg["idx"]]
        title = table.rows[0].cells[0].text.upper()
        if cfg["sig"].upper() not in title:
            rep.flags.append(f"⚠ Tabla {cfg['idx']} no coincide con firma '{cfg['sig']}' (encontrado: '{title[:40]}'). Se omite.")
            continue
        _PROCESSORS[cfg["kind"]](table, cfg, b26, b25, rep)
    return rep
