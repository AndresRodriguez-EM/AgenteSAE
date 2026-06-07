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
            hijos = [t for cc in bal.cuentas.values()
                     if cc.codigo.startswith(c.codigo) and cc.terceros for t in cc.terceros]
            recs.append(dict(kind="acc", nit=None, code=c.codigo, value=c.nuevo_saldo,
                             mov=abs(c.debito - c.credito),
                             des=_anc_names(bal, c.codigo) + [c.nombre],
                             children={id(t) for t in hijos},
                             child_nits={t.nit for t in hijos}))
    return recs


def _target_account(word_nits, group_des, group_total, acc_recs):
    """Elige la cuenta (registro 'acc' de 4/6 dígitos) a la que pertenece un grupo
    de filas del Word con la misma DES CUENTA. Prioriza el solape de NITs (lo más
    fiable), luego la coincidencia de descripción y por último la cercanía de
    valor. Devuelve el código de cuenta o None."""
    if not acc_recs:
        return None
    dw = _toks(group_des)
    gt = _round(abs(group_total)) if group_total else None

    def feats(r):
        nit_ov = len(word_nits & r.get("child_nits", set()))
        des_ov = max((len(dw & _toks(x)) for x in r["des"]), default=0)
        exact = 1 if (gt is not None and _round(abs(r["value"])) == gt) else 0
        val_dist = abs(abs(r["value"]) - abs(group_total)) if group_total else 0
        return exact, des_ov, nit_ov, val_dist

    # Prioridad: valor exacto > descripción > NITs en común > cercanía de valor.
    best = max(acc_recs, key=lambda r: (lambda e, d, n, v: (e, d, n, -v))(*feats(r)))
    e, d, n, _ = feats(best)
    return best["code"] if (e or d or n) else None


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
    # A nivel de cuenta solo si el NIT pertenece a esa cuenta (es uno de sus terceros).
    ac = [r for r in recs if r["kind"] == "acc" and nit in r.get("child_nits", set())]
    rh = _round(abs(hint)) if hint else None
    if rh is not None:
        et = [r for r in tc if _round(abs(r["value"])) == rh]
        if et:
            return _pick(et, des, hint)
        ea = [r for r in ac if _round(abs(r["value"])) == rh]
        if ea:
            return _pick(ea, des, hint)
    if tc:
        return _pick(tc, des, hint)
    if ac:
        return _pick(ac, des, hint)
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
        self.changes = []        # (nota, tabla, concepto, columna, antes, despues)
        self.added = []          # (tabla, nit, nombre, valores)
        self.removed = []        # (tabla, nit, nombre)
        self.added_tables = []   # (nota, titulo, n_terceros)
        self.flags = []          # observaciones

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

    # Agrupar filas por DES CUENTA y mapear cada grupo a su cuenta del auxiliar
    # (por solape de NITs + descripción + valor). Así una fila se empareja SOLO
    # con terceros de su cuenta y no se "contamina" con el mismo NIT de otra.
    acc26 = [x for x in recs26 if x["kind"] == "acc"]
    grupos = {}
    for r in data_idx:
        dk = table.rows[r].cells[col_des].text.strip().upper() if col_des is not None else ""
        grupos.setdefault(dk, []).append(r)
    target_map = {}
    for dk, rows in grupos.items():
        wnits = {_norm_nit(table.rows[r].cells[col_nit].text) for r in rows} if col_nit is not None else set()
        gtot = sum(parse_money(table.rows[r].cells[col_actual].text) or 0.0 for r in rows) if col_actual is not None else 0.0
        dtext = table.rows[rows[0]].cells[col_des].text if col_des is not None else ""
        target_map[dk] = _target_account(wnits, dtext, gtot, acc26)

    to_remove = []
    for r in data_idx:
        row = table.rows[r]
        nit = _norm_nit(row.cells[col_nit].text) if col_nit is not None else ""
        des = row.cells[col_des].text if col_des is not None else ""
        tname = row.cells[col_ter].text if col_ter is not None else ""
        hint26 = parse_money(row.cells[col_actual].text) if col_actual is not None else None
        hint25 = parse_money(row.cells[col_comp].text) if col_comp is not None else None
        tgt = target_map.get(des.strip().upper())
        sub26 = [x for x in recs26 if x["code"].startswith(tgt)] if tgt else recs26
        sub25 = [x for x in recs25 if x["code"].startswith(tgt)] if tgt else recs25
        r26 = match_record(sub26, nit, des, hint26, tname=tname)
        r25 = find_same25(sub25, r26) if r26 else match_record(sub25, nit, des, hint25, tname=tname)
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
            year = sub25 if role == "comparativo" else sub26
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


def _best_sub(subs, label):
    """Subcuenta cuyo nombre mejor coincide con el rótulo de la fila."""
    lt = _toks(label)
    if not lt:
        return None
    best, bs = None, 0
    for c in subs:
        sc = len(lt & _toks(c.nombre))
        if sc > bs or (sc == bs and best is not None and len(c.codigo) < len(best.codigo)):
            bs, best = sc, c
    return best if bs > 0 else None


def process_cuenta(table, cfg, b26, b25, rep: Report):
    """Notas a nivel de cuenta (sin terceros), p. ej. Efectivo. Cada fila se
    mapea a su subcuenta por descripción; el total a la clase configurada."""
    nota = cfg.get("nota", "")
    header_r = _find_header_row(table)
    roles = {i: _role_of(c.text) for i, c in enumerate(table.rows[header_r].cells) if _role_of(c.text)}
    parent = cfg["account"]
    subs = [c for c in b26.cuentas.values() if c.codigo.startswith(parent) and c.codigo != parent]

    data, total = [], None
    for r in range(header_r + 1, len(table.rows)):
        row = table.rows[r]
        if not any(parse_money(c.text) is not None for _, c in _unique_grid(row)):
            continue
        if row.cells[0].text.strip():
            data.append(r)
        elif total is None:
            total = r

    def _set(row, concepto, getval):
        for col, role in roles.items():
            val = getval(role)
            cell = row.cells[col]
            existing = parse_money(cell.text)
            if existing is not None and abs(existing - val) <= ROUND_TOL:
                continue
            antes = cell.text.strip()
            nuevo = format_like(val, cell)
            if antes != nuevo:
                set_cell_value(cell, nuevo)
                rep.chg(nota, cfg["sig"], concepto, role, antes, nuevo)

    for r in data:
        row = table.rows[r]
        sub = _best_sub(subs, row.cells[0].text)
        code = sub.codigo if sub else parent
        _set(row, row.cells[0].text.strip() or "TOTAL",
             lambda role, code=code: abs(b25.saldo(code) if role == "comparativo" else b26.saldo(code)))
    if total is not None:
        _set(table.rows[total], "TOTAL",
             lambda role: abs(b25.saldo(parent) if role == "comparativo" else b26.saldo(parent)))


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
#  Configuración por NOTA (independiente de la posición de las tablas)
#
#  El recorrido identifica a qué nota pertenece cada tabla siguiendo los
#  encabezados "NOTA n" y "VALOR TOTAL NOTA n"; dentro de cada nota, cada tabla
#  de datos se reconoce por su título. Así funciona aunque cambien las filas o
#  el número de tablas entre sociedades.
#
#  Cada entrada de "tablas": (subcadena_de_titulo, kind, alcance)
#    - kind="terceros"  -> alcance = lista de prefijos de cuenta del PUC
#    - kind="cuenta"    -> alcance = clase padre (p. ej. "11")
#    - kind="patrimonio"-> alcance = None
# --------------------------------------------------------------------------- #

NOTE_CONFIG = {
    "1": {"vtotal": "11", "tablas": [("EFECTIVO Y EQUIVALENTE", "cuenta", "11")]},
    "2": {"vtotal": "13", "tablas": [("DETALLE CUENTAS POR COBRAR", "terceros", ["1345"]),
                                     ("ANTICIPO DE IMPUESTOS", "terceros", ["1355"])]},
    # Nota 3 (Propiedad, planta y equipo): EXCLUIDA — se actualiza manualmente.
    "4": {"vtotal": "2", "tablas": [("CUENTAS POR PAGAR", "terceros", ["2335", "2365"]),
                                    ("MPUESTOS POR PAGAR", "terceros",
                                     ["2404", "2405", "2408", "2412", "2416"]),
                                    ("IMPUESTOS MULTAS Y SANCIONES", "terceros", ["2615", "2635"])]},
    "5": {"vtotal_patrimonio": True, "tablas": [("PATRIMONIO", "patrimonio", None)]},
    "6": {"vtotal": "41", "tablas": [("INGRESOS POR ARRENDAMIENTOS", "terceros", ["4155"])]},
    "7": {"vtotal": "51", "tablas": [("HONORARIOS", "terceros", ["5110"]),
                                     ("IMPUESTOS", "terceros", ["5115"]),
                                     ("GASTOS LEGALES", "terceros", ["5140"]),
                                     ("SERVICIOS", "terceros", ["5135"]),
                                     ("MANTENIMIENTO", "terceros", ["5145"]),
                                     ("DIVERSOS", "terceros", ["5195"])]},
    "8": {"vtotal": "53", "tablas": [("GASTOS EXTRAORDINARIOS", "terceros", ["5305"]),
                                     ("GASTOS EXTRAORDINARIOS", "terceros", ["5315"])]},
    "9": {"vtotal": "42", "tablas": [("OTROS INGRESOS", "terceros", ["4295"])]},
}

SKIP_NOTES = {"3"}

_PROCESSORS = {
    "terceros": process_terceros,
    "cuenta": process_cuenta,
    "patrimonio": process_patrimonio,
}


def _nota_en_fila(table):
    """Devuelve el número de nota si la tabla contiene una fila de encabezado
    'NOTA | n' (sea tabla de encabezado o fila embebida en un VALOR TOTAL)."""
    nota = None
    for row in table.rows:
        cells = [c.text.strip() for c in _unique_grid_cells(row)]
        if cells and cells[0].upper() == "NOTA":
            for c in cells[1:]:
                if c.isdigit() and 1 <= int(c) <= 9:
                    nota = c
                    break
    return nota


def _unique_grid_cells(row):
    out, seen = [], set()
    for c in row.cells:
        if id(c._tc) in seen:
            continue
        seen.add(id(c._tc))
        out.append(c)
    return out


def _des_de_rec(bal: Balance, code: str, fallback: str) -> str:
    """Descripción de cuenta (nombre de la cuenta del tercero) para una fila nueva."""
    for L in (6, 8, 4, 10):
        c = bal.cuentas.get(code[:L])
        if c and c.nombre:
            return c.nombre
    return fallback


def crear_tabla(template_table, anchor_el, titulo, scope, b26, b25, nota, rep) -> bool:
    """Crea una tabla de terceros faltante clonando la estructura de otra del
    mismo tipo y la inserta antes del cuadro VALOR TOTAL de la nota."""
    if template_table is None or anchor_el is None:
        return False
    recs26 = scope_records(b26, scope)
    recs25 = scope_records(b25, scope)
    map26 = {(r["nit"], r["code"]): r for r in recs26 if r["kind"] == "t"}
    map25 = {(r["nit"], r["code"]): r for r in recs25 if r["kind"] == "t"}
    keys = list(map26) + [k for k in map25 if k not in map26]
    pares = []
    for k in keys:
        r26, r25 = map26.get(k), map25.get(k)
        if (abs(r26["value"]) if r26 else 0) >= 1 or (r26["mov"] if r26 else 0) >= 1 \
                or (abs(r25["value"]) if r25 else 0) >= 1:
            pares.append((r26, r25))
    if not pares:
        return False

    from docx.table import Table, _Row
    from docx.oxml import OxmlElement

    new_tbl = copy.deepcopy(template_table._tbl)
    t = Table(new_tbl, template_table._parent)
    header_r = _find_header_row(t)
    if header_r is None:
        return False
    hdr = t.rows[header_r]
    roles = {i: _role_of(c.text) for i, c in enumerate(hdr.cells) if _role_of(c.text)}
    col_nit = _col_index(hdr, "CEDULA", "CÉDULA", "ID,")
    col_ter = _col_index(hdr, "TERCERO")
    col_des = _col_index(hdr, "DES CUENTA", "DESCRIPCION", "DES CUE")
    data_idx, _ = _data_rows(t, header_r, col_nit)
    if not data_idx or col_nit is None:
        return False
    nit_sep = _detect_nit_sep(t.rows[data_idx[0]].cells[col_nit].text)

    set_cell_value(t.rows[0].cells[0], titulo)                 # título de la tabla

    orig_trs = [t.rows[r]._tr for r in data_idx]
    last = orig_trs[-1]
    for r26, r25 in pares:
        base = r26 or r25
        new_tr = copy.deepcopy(orig_trs[0])
        last.addnext(new_tr)
        last = new_tr
        nr = _Row(new_tr, t)
        set_cell_value(nr.cells[col_nit], _fmt_nit(base["nit"], nit_sep))
        if col_ter is not None:
            set_cell_value(nr.cells[col_ter], base["t"].nombre)
        if col_des is not None:
            set_cell_value(nr.cells[col_des], _des_de_rec(b26 if r26 else b25, base["code"], base["t"].nombre))
        for col, role in roles.items():
            set_cell_value(nr.cells[col], format_like(_val(role, r26, r25), nr.cells[col]))
    for tr in orig_trs:
        tr.getparent().remove(tr)

    data_idx, total_idx = _data_rows(t, header_r, col_nit)
    if total_idx is not None:
        trow = t.rows[total_idx]
        for col, role in roles.items():
            s = sum(parse_money(t.rows[r].cells[col].text) or 0.0 for r in data_idx)
            set_cell_value(trow.cells[col], format_like(s, trow.cells[col]))

    anchor_el.addprevious(new_tbl)
    new_tbl.addprevious(OxmlElement("w:p"))
    rep.added_tables.append((nota, titulo, len(pares)))
    return True


def update_document(doc, b26: Balance, b25: Balance, skip_notes=SKIP_NOTES, crear_faltantes=True) -> Report:
    rep = Report()
    queues = {n: list(cfg.get("tablas", [])) for n, cfg in NOTE_CONFIG.items()}
    current_note = None
    note_template = {}   # nota -> tabla de terceros (para clonar faltantes)
    note_anchor = {}     # nota -> elemento del cuadro VALOR TOTAL (punto de inserción)

    for table in doc.tables:
        title = table.rows[0].cells[0].text.strip()
        up = title.upper()

        m = re.search(r"VALOR TOTAL NOTA\s*0*(\d+)", up)
        if m:                                   # cuadro "VALOR TOTAL NOTA n"
            n = m.group(1)
            cfg = NOTE_CONFIG.get(n, {})
            note_anchor[n] = table._tbl
            if n not in skip_notes and cfg:
                if cfg.get("vtotal_patrimonio"):
                    process_vtotal(table, {"patrimonio": True, "nota": n, "sig": title}, b26, b25, rep)
                elif "vtotal" in cfg:
                    process_vtotal(table, {"account": cfg["vtotal"], "nota": n, "sig": title}, b26, b25, rep)
        elif _find_header_row(table) is not None and current_note and current_note not in skip_notes:
            q = queues.get(current_note, [])
            for i, (sub, kind, scope) in enumerate(q):
                if sub.upper() in up:
                    cfg = dict(nota=current_note, sig=title, kind=kind)
                    if kind == "terceros":
                        cfg["scope"] = scope
                        note_template[current_note] = table
                    elif kind == "cuenta":
                        cfg["account"] = scope
                    _PROCESSORS[kind](table, cfg, b26, b25, rep)
                    q.pop(i)
                    break

        # Actualizar la nota vigente para las tablas siguientes.
        n_aqui = _nota_en_fila(table)
        if n_aqui:
            current_note = n_aqui

    # Tablas esperadas no encontradas: crearlas desde el auxiliar (o avisar).
    for n, q in queues.items():
        if n in skip_notes:
            continue
        for sub, kind, scope in q:
            creada = False
            if crear_faltantes and kind == "terceros":
                titulo = (b26.cuenta(scope[0]).nombre if b26.cuenta(scope[0])
                          else (b25.cuenta(scope[0]).nombre if b25.cuenta(scope[0]) else sub))
                creada = crear_tabla(note_template.get(n), note_anchor.get(n),
                                     titulo, scope, b26, b25, n, rep)
            if not creada:
                rep.flags.append(f"Nota {n}: no se encontró la tabla '{sub}' "
                                 + ("(sin datos en el auxiliar)." if kind == "terceros" else "(¿template distinto?)."))
    return rep
