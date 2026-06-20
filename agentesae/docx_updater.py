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
import calendar
import copy
import math
import re

from .pdf_parser import Balance

# --------------------------------------------------------------------------- #
#  Utilidades de formato
# --------------------------------------------------------------------------- #

def parse_money(text: str):
    core = (text or "").replace("$", "").replace(" ", "").strip()
    neg = core.startswith("-") or (core.startswith("(") and core.endswith(")"))
    core = core.lstrip("-").strip("()")
    # Solo dígitos con punto/coma como separador de miles (rechaza fechas "01/08/2024"
    # y matrículas "126-403", que no son importes).
    if not core or not re.fullmatch(r"[\d.,]+", core):
        return None
    digits = core.replace(".", "").replace(",", "")
    if not digits.isdigit():
        return None
    v = float(digits)
    return -v if neg else v


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


# Palabras vacías (conectores) que NO deben contar como coincidencia de nombre:
# 'INTERESES POR MULTAS' no debe parecerse más a 'IMPUESTOS... POR PAGAR' por el 'POR'.
_STOPWORDS = {"POR", "PARA", "CON", "LAS", "LOS", "DEL", "SUS", "QUE", "SIN", "ANTE"}


def _toks(s: str) -> set[str]:
    return set(re.findall(r"[A-ZÁÉÍÓÚÑ]{3,}", (s or "").upper())) - _STOPWORDS


def _commonprefix(strs) -> str:
    if not strs:
        return ""
    lo, hi = min(strs), max(strs)
    i = 0
    while i < len(lo) and lo[i] == hi[i]:
        i += 1
    return lo[:i]


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
    año = bool(re.search(r"\bA[NÑ]O\b", h))
    # Un encabezado de columna de VALOR siempre referencia un PERÍODO (mes, año,
    # 'actual'/'anterior' o un año). Así un TÍTULO que contiene 'ACUMULADO(S)' o
    # 'MOV...' (p. ej. 'GASTOS DIVERSOS ACUMULADOS') no se confunde con una columna
    # —antes eso convertía el título en encabezado y llenaba toda la tabla.
    if not ("MES" in h or año or "ANTERIOR" in h or "ACTUAL" in h or re.search(r"20\d\d", h)):
        return None
    if "MOV" in h:
        return "mov"
    comparativo = ("2025" in h) or ("2024" in h) or ("ANTERIOR" in h) or año
    if "ACUMULAD" in h:
        return "comparativo" if comparativo else "acum_actual"
    if "SALDO" in h:
        if comparativo:
            return "comparativo"
        if "ACTUAL" in h or "2026" in h:    # 'SALDO MES 2026' = columna del periodo
            return "actual"
    return None


def _find_header_row(table):
    for r, row in enumerate(table.rows):
        texts = [c.text.upper() for c in row.cells]
        if any("TERCERO" in t for t in texts) or any(_role_of(t) for t in texts):
            if any(_role_of(t) for t in texts):
                return r
    return None


_ACCENTS = str.maketrans("ÁÉÍÓÚÜÑáéíóúüñ", "AEIOUUNaeiouun")


def _deaccent(s: str) -> str:
    return (s or "").translate(_ACCENTS)


def _col_index(row, *keywords):
    # Insensible a tildes: 'DESCRIPCIÓN'/'CÉDULA' deben casar con 'DESCRIPCION'/'CEDULA'.
    for i, c in enumerate(row.cells):
        u = _deaccent(c.text.upper())
        if any(_deaccent(k.upper()) in u for k in keywords):
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


def _merge_acc(acc26, acc25):
    """Une los registros de cuenta de ambos años por código (para poder ubicar
    una cuenta que solo tiene saldo en el año comparativo)."""
    m = {}
    for r in acc26:
        m[r["code"]] = {"code": r["code"], "des": list(r["des"]),
                        "child_nits": set(r["child_nits"]), "v26": r["value"], "v25": 0.0}
    for r in acc25:
        e = m.setdefault(r["code"], {"code": r["code"], "des": list(r["des"]),
                                     "child_nits": set(), "v26": 0.0, "v25": 0.0})
        e["v25"] = r["value"]
        e["child_nits"] |= set(r["child_nits"])
        for x in r["des"]:
            if x not in e["des"]:
                e["des"].append(x)
    return list(m.values())


def _target_account(word_nits, des26, gtot26, gtot25, acc_cands, n_rows=1):
    """Elige el código de cuenta al que pertenece un grupo de filas del Word.
    Considera ambos años (una cuenta puede tener saldo solo en el comparativo).

    - Grupos de UNA fila: valor exacto y cercanía de valor (la cifra del Word es
      fiable; la descripción puede compartir palabras como 'RENTA').
    - Grupos de VARIAS filas: valor exacto y descripción (el total puede venir
      inflado por una fila espuria, así que el nombre manda).
    """
    if not acc_cands:
        return None
    dw = _toks(des26)
    g26 = _round(abs(gtot26)) if gtot26 else None
    g25 = _round(abs(gtot25)) if gtot25 else None

    def feats(r):
        nit_ov = len(word_nits & r.get("child_nits", set()))
        des_ov = max((len(dw & _toks(x)) for x in r["des"]), default=0)
        exact = 1 if ((g26 is not None and _round(abs(r["v26"])) == g26)
                      or (g25 is not None and _round(abs(r["v25"])) == g25)) else 0
        d26 = abs(abs(r["v26"]) - abs(gtot26)) if gtot26 else None
        d25 = abs(abs(r["v25"]) - abs(gtot25)) if gtot25 else None
        val_dist = min([d for d in (d26, d25) if d is not None], default=0)
        return exact, des_ov, nit_ov, val_dist

    # Prioridad de señales. La pertenencia por NIT (la fila ES un tercero de esa
    # cuenta) manda sobre la cercanía de valor: el saldo del Word puede no cuadrar
    # con el de la cuenta cuando faltan terceros por agregar (p. ej. EMBARGOS, donde
    # la cuenta incluye un tercero ausente en el Word). El largo del código desempata
    # hacia la cuenta más específica. Grupos multi-fila: la descripción pesa más.
    if n_rows > 1:
        key = lambda r: (lambda e, d, n, v: (e, d, n, -v, len(r["code"])))(*feats(r))
    else:
        key = lambda r: (lambda e, d, n, v: (e, n, d, -v, len(r["code"])))(*feats(r))
    best = max(acc_cands, key=key)
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


def _account_code_rec(bal: Balance, code: str, prefixes):
    """Si la 'cédula' de una fila es en realidad un CÓDIGO DE CUENTA del alcance
    (algunas notas traen filas por cuenta, p. ej. 'saldo a favor renta'),
    devuelve un registro a nivel de cuenta con sus valores."""
    c = bal.cuentas.get(code)
    if c is None or not any(code.startswith(p) for p in prefixes):
        return None
    return dict(kind="acc", nit=None, code=code, value=c.nuevo_saldo,
                mov=abs(c.debito - c.credito), des=[c.nombre], child_nits=set())


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


def _es_neteo(recs, code4):
    """True si bajo la cuenta de 4 dígitos hay terceros con signos opuestos
    (la cuenta se 'netea', p. ej. IVA generado vs IVA descontable, o retención
    causada vs pago de retención). En esos casos el Word muestra el NETO."""
    vals = [r["value"] for r in recs if r["kind"] == "t" and r["code"].startswith(code4)]
    return any(v > 0 for v in vals) and any(v < 0 for v in vals)


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
    # Solo el registro emparejado (prim) y los netos de su familia de cuenta;
    # NO todos los terceros del mismo NIT (eso aceptaba valores ajenos, p. ej.
    # el movimiento 0 de otro tercero de la DIAN).
    for r in recs_year:
        if r["kind"] == "acc" and famcode and famcode.startswith(r["code"]):
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
        self.date_changes = 0    # fechas de período actualizadas
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
    col_des = _col_index(hdr, "DES CUENTA", "DESCRIPCION", "DESCRICION", "DES CUE")

    recs26 = scope_records(b26, cfg["scope"])
    recs25 = scope_records(b25, cfg["scope"])
    allowed = {r["nit"] for r in recs26 + recs25 if r["kind"] == "t"}
    matched_tids = set()
    matched_keys = set()        # (nit, code) de terceros ya representados por una fila
    covered = set()

    data_idx, total_idx = _data_rows(table, header_r, col_nit)
    nit_sep = _detect_nit_sep(table.rows[data_idx[0]].cells[col_nit].text) if data_idx and col_nit is not None else ","
    dollar = cell_has_dollar(table.rows[data_idx[0]].cells[list(roles)[0]]) if data_idx and roles else True
    # Plantilla de fila (clon limpio) para crear filas nuevas, capturada ANTES de
    # cualquier eliminación (si todas las filas se borran, igual hay de dónde clonar).
    tmpl_tr = copy.deepcopy(table.rows[data_idx[0]]._tr) if data_idx else None
    col_actual = next((c for c, r in roles.items() if r in ("actual", "acum_actual")), None)
    col_comp = next((c for c, r in roles.items() if r == "comparativo"), None)

    # Agrupar filas por DES CUENTA y mapear cada grupo a su cuenta del auxiliar
    # (por solape de NITs + descripción + valor). Así una fila se empareja SOLO
    # con terceros de su cuenta y no se "contamina" con el mismo NIT de otra.
    acc_cands = _merge_acc([x for x in recs26 if x["kind"] == "acc"],
                           [x for x in recs25 if x["kind"] == "acc"])
    # Cuenta ANCLA por el TÍTULO de la tabla: el título suele ser el nombre de la
    # cuenta (p. ej. 'DEUDORES VARIOS' -> 1380, 'ANTICIPO DE IMPUESTOS' -> 1355,
    # 'CUOTAS DE ADMINISTRACION' -> 1345950501). Es la señal MÁS FIABLE para ubicar
    # la tabla: evita que un saldo desactualizado del Word mande las filas a otra
    # cuenta de la misma clase con valor parecido (el NIT/valor no bastan cuando un
    # mismo tercero —p. ej. la DIAN— aparece en varias cuentas de la clase).
    # El ancla es el PREFIJO COMÚN de las cuentas cuyo nombre coincide con el título
    # (>=2 palabras): así una tabla como 'IMPUESTOS MULTAS Y SANCIONES', cuyo título
    # casa con dos cuentas hermanas (263505 y 263510), se ancla en el padre común
    # (2635) y el pool incluye a AMBAS —no a una sola, que dejaría sin candidato a la
    # otra fila—. Restringir el pool basta para no irse a otra cuenta de la clase con
    # saldo parecido; no se fuerza el destino (cada grupo elige su cuenta en el pool).
    title_toks = _toks(cfg.get("sig", ""))
    anchor = None
    if title_toks:
        ovs = [(max((len(title_toks & _toks(x)) for x in a["des"]), default=0), a["code"]) for a in acc_cands]
        best_ov = max((o for o, _ in ovs), default=0)
        if best_ov >= 2:
            pre = _commonprefix([code for o, code in ovs if o == best_ov])
            pre = pre[:len(pre) // 2 * 2]               # códigos PUC: longitud par
            if len(pre) >= 4:
                anchor = pre
    pool = [a for a in acc_cands if a["code"].startswith(anchor)] if anchor else acc_cands
    grupos = {}
    for r in data_idx:
        dk = table.rows[r].cells[col_des].text.strip().upper() if col_des is not None else ""
        grupos.setdefault(dk, []).append(r)
    target_map = {}
    for dk, rows in grupos.items():
        wnits = {_norm_nit(table.rows[r].cells[col_nit].text) for r in rows} if col_nit is not None else set()
        g26 = sum(parse_money(table.rows[r].cells[col_actual].text) or 0.0 for r in rows) if col_actual is not None else 0.0
        g25 = sum(parse_money(table.rows[r].cells[col_comp].text) or 0.0 for r in rows) if col_comp is not None else 0.0
        dtext = table.rows[rows[0]].cells[col_des].text if col_des is not None else ""
        target_map[dk] = _target_account(wnits, dtext, g26, g25, pool, len(rows))
    # Cuentas ESPECÍFICAS que ESTA tabla representa (código completo, 6/10 díg):
    # solo a ellas se agregan terceros nuevos. Antes se truncaba a 4 díg, lo que
    # mezclaba sub-tablas hermanas (p. ej. ARRENDAMIENTO 134530 con CUOTA DE
    # ADMINISTRACIÓN 1345950501, ambas bajo 1345). `table_accounts` (4 díg) se
    # conserva solo para la lógica de "crear tablas faltantes".
    target_codes = {t for t in target_map.values() if t}
    table_accounts = {t[:4] for t in target_codes}

    to_remove = []
    for r in data_idx:
        row = table.rows[r]
        nit = _norm_nit(row.cells[col_nit].text) if col_nit is not None else ""
        des = row.cells[col_des].text if col_des is not None else ""
        tname = row.cells[col_ter].text if col_ter is not None else ""
        hint26 = parse_money(row.cells[col_actual].text) if col_actual is not None else None
        hint25 = parse_money(row.cells[col_comp].text) if col_comp is not None else None
        dk = des.strip().upper()
        tgt = target_map.get(dk)
        # Ámbito de emparejamiento: la cuenta destino si se resolvió; si no, el ancla
        # de la tabla (no toda la clase) para no traer un tercero de otra cuenta.
        scope_pref = tgt or anchor
        sub26 = [x for x in recs26 if x["code"].startswith(scope_pref)] if scope_pref else recs26
        sub25 = [x for x in recs25 if x["code"].startswith(scope_pref)] if scope_pref else recs25
        r26 = match_record(sub26, nit, des, hint26, tname=tname)
        # Fila que representa una CUENTA sin terceros (etiquetada con un NIT
        # representativo, p. ej. 'ANTICIPO IMPUESTO DE RENTA' = 135505, que en el
        # auxiliar no tiene tercero): se usa el saldo de la propia cuenta. Solo si el
        # grupo es de UNA fila (evita duplicar el valor de la cuenta en varias).
        if r26 is None and tgt and len(grupos.get(dk, [])) == 1:
            r26 = next((x for x in recs26 if x["kind"] == "acc" and x["code"] == tgt), None)
        # Si la cuenta se netea (sub-cuentas con signos opuestos), usar el NETO de
        # la cuenta de 4 dígitos (no el de una sub-cuenta) en ambos años.
        if r26 is not None and tgt and (_es_neteo(recs26, tgt[:4]) or _es_neteo(recs25, tgt[:4])):
            acc4 = next((x for x in sub26 if x["kind"] == "acc" and x["code"] == tgt[:4]), None)
            if acc4 is not None:
                r26 = acc4
        r25 = find_same25(sub25, r26) if r26 else match_record(sub25, nit, des, hint25, tname=tname)
        if r26 is None and r25 is None and nit:
            # ¿La 'cédula' es en realidad un código de cuenta? (filas por cuenta)
            r26 = _account_code_rec(b26, nit, cfg["scope"])
            r25 = _account_code_rec(b25, nit, cfg["scope"])
        if r26 is None and r25 is None:
            # Eliminar SOLO si el NIT no existe en ninguna cuenta de la clase
            # (no borrar terceros reales que solo quedaron mal agrupados).
            if nit and nit not in allowed:
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
                matched_tids |= r26.get("children", set())
        for rr in (r26, r25):
            if rr and rr["kind"] == "t":
                matched_keys.add((rr["nit"], rr["code"]))
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

    # Altas: terceros del auxiliar (2026 y SOLO-2025) no emparejados ni cubiertos.
    # En tablas de impuestos (add=False) NO se agregan terceros: solo se actualizan
    # las filas existentes (cada subcuenta/tercero a su valor).
    if col_nit is not None and tmpl_tr is not None and cfg.get("add", True):
        from docx.table import _Row
        cur, tot = _data_rows(table, header_r, col_nit)
        last_tr = [table.rows[cur[-1]]._tr if cur else None]
        tot_tr = table.rows[tot]._tr if (not cur and tot is not None) else None
        shown_nits = {_norm_nit(table.rows[r].cells[col_nit].text) for r in cur}
        present26 = {(r["nit"], r["code"]) for r in recs26 if r["kind"] == "t"}

        def _add(rec, r25, bal):
            des_new = None
            if col_des is not None:
                for ri in _data_rows(table, header_r, col_nit)[0]:
                    rr = match_record(recs26 + recs25, _norm_nit(table.rows[ri].cells[col_nit].text),
                                      table.rows[ri].cells[col_des].text, None)
                    if rr and rr["code"][:6] == rec["code"][:6]:   # misma SUBcuenta (6 díg)
                        des_new = table.rows[ri].cells[col_des].text.strip()
                        break
            if des_new is None:
                des_new = _des_de_rec(bal, rec["code"], rec["t"].nombre)
            new_tr = copy.deepcopy(tmpl_tr)
            if last_tr[0] is not None:
                last_tr[0].addnext(new_tr)
            elif tot_tr is not None:
                tot_tr.addprevious(new_tr)
            else:
                table._tbl.append(new_tr)
            last_tr[0] = new_tr
            nrow = _Row(new_tr, table)
            set_cell_value(nrow.cells[col_nit], _fmt_nit(rec["nit"], nit_sep))
            if col_ter is not None:
                set_cell_value(nrow.cells[col_ter], rec["t"].nombre)
            if col_des is not None:
                set_cell_value(nrow.cells[col_des], des_new)
            vals = {}
            r26 = rec if bal is b26 else None
            for col, role in roles.items():
                v = _val(role, r26, r25)
                set_cell_value(nrow.cells[col], format_like(v, nrow.cells[col]))
                vals[role] = v
            shown_nits.add(rec["nit"])
            rep.added.append((cfg["sig"], rec["nit"], rec["t"].nombre, vals))

        def _addable(rec):
            return (rec["kind"] == "t"
                    and not any(rec["code"].startswith(cov) for cov in covered)
                    and any(rec["code"].startswith(tc) for tc in target_codes))

        for rec in recs26:                              # terceros del periodo (2026)
            if not _addable(rec) or rec["tid"] in matched_tids:
                continue
            if abs(rec["value"]) < 1 and rec["mov"] < 1:
                continue
            _add(rec, find_same25(recs25, rec), b26)
        for rec in recs25:                              # terceros SOLO del comparativo (2025)
            if (rec["nit"], rec["code"]) in present26 or (rec["nit"], rec["code"]) in matched_keys:
                continue
            if not _addable(rec) or abs(rec["value"]) < 1:
                continue
            _add(rec, rec, b25)

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

    return matched_tids, table_accounts


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


# Clases del PUC que se presentan a nivel de CUENTA/SUBCUENTA (no de tercero):
# impuestos por pagar (24), obligaciones laborales (25), pasivos estimados (26)
# y el anticipo de impuestos del activo (1355). En estas tablas cada fila es una
# SUBCUENTA (impuesto de renta, IVA, ICA, predial...) y su valor es el de la
# subcuenta —NO la suma de sus terceros— y NO se agregan terceros.
_IMPUESTO_CLASES = {"24", "25", "26"}


def _impuesto_scope(nota: str):
    if nota in ("4",):
        return ["24", "25", "26"]
    if nota in ("2",):
        return ["1355"]
    return []


def _subcuentas_valor(b26, b25, clases, niveles=(4, 6)):
    """Subcuentas (por defecto 4/6 díg) de esas clases con su saldo en ambos años."""
    out = {}
    for bal in (b26, b25):
        for c in bal.cuentas.values():
            if len(c.codigo) in niveles and any(c.codigo.startswith(p) for p in clases):
                out.setdefault(c.codigo, c)
    return out


def _es_tabla_impuestos(table, header_r, b26, b25, nota):
    """True si la tabla se debe tratar a nivel de subcuenta de impuestos: la mayoría
    de sus filas (por descripción) casan con una subcuenta de impuestos (24/25/26 ó
    1355). Distingue 'IMPUESTOS POR PAGAR' (subcuentas por tipo de impuesto) de
    'cuentas por pagar/honorarios' (terceros) y de 'multas y sanciones' (cuyas
    subcuentas no se nombran por tipo de impuesto)."""
    clases = _impuesto_scope(nota)
    if not clases:
        return False
    # 'Multas y sanciones' / 'provisiones' / 'contingencias' (pasivos estimados) NO
    # son la tabla de impuestos por pagar: sus subcuentas no se nombran por tipo de
    # impuesto y se emparejan por tercero/valor (no por subcuenta).
    titulo = table.rows[0].cells[0].text.upper()
    if any(k in titulo for k in ("MULTA", "SANCION", "PROVISION", "CONTINGENC")):
        return False
    subs = list(_subcuentas_valor(b26, b25, clases).values())
    if not subs:
        return False
    # Por TÍTULO: 'IMPUESTO(S)...' / 'IVA' / 'ANTICIPO DE IMPUESTOS' son tablas de
    # impuestos (también las de una sola subcuenta: 'IMPUESTO POR PAGAR - IVA').
    if "IMPUESTO" in titulo or re.search(r"\bIVA\b", titulo):
        return True
    hdr = table.rows[header_r]
    col_des = _col_index(hdr, "DESCRIPCION", "DESCRICION", "DES CUENTA", "DES CUE")
    if col_des is None:
        return False
    n_rows = n_match = 0
    for r in range(header_r + 1, len(table.rows)):
        des = table.rows[r].cells[col_des].text.strip()
        if not des:
            continue
        n_rows += 1
        sub = _best_sub(subs, des)
        if sub is not None and any(sub.codigo.startswith(p) for p in clases):
            n_match += 1
    return n_rows > 0 and n_match >= max(1, (n_rows + 1) // 2)


def _nits_distintos(table, header_r):
    """True si las filas de datos tienen NITs DISTINTOS entre sí (terceros reales,
    p. ej. predial por municipio). False si el mismo NIT se repite o hay uno solo
    (una subcuenta presentada en una/varias filas: IVA, renta, anticipo)."""
    col_nit = _col_index(table.rows[header_r], "CEDULA", "CÉDULA", "ID,")
    if col_nit is None:
        return False
    nits = []
    for r in range(header_r + 1, len(table.rows)):
        nn = _norm_nit(table.rows[r].cells[col_nit].text)
        if nn:
            nits.append(nn)
    return len(nits) > 1 and len(set(nits)) == len(nits)


def process_impuestos(table, cfg, b26, b25, rep: Report):
    """Tabla de impuestos a nivel de SUBCUENTA. Cada fila se mapea a una subcuenta
    (por descripción y, si sobran, por eliminación) y toma SU saldo —no la suma de
    terceros—. No agrega terceros. El total = suma de las filas mostradas."""
    nota = cfg.get("nota", "")
    clases = _impuesto_scope(nota) or cfg.get("scope", [])
    header_r = _find_header_row(table)
    if header_r is None:
        return
    hdr = table.rows[header_r]
    roles = {i: _role_of(c.text) for i, c in enumerate(hdr.cells) if _role_of(c.text)}
    if not roles:
        return
    col_des = _col_index(hdr, "DESCRIPCION", "DESCRICION", "DES CUENTA", "DES CUE")
    col_nit = _col_index(hdr, "CEDULA", "CÉDULA", "ID,")

    data, total = [], None
    for r in range(header_r + 1, len(table.rows)):
        row = table.rows[r]
        if not any(parse_money(c.text) is not None for _, c in _unique_grid(row)):
            continue
        des = row.cells[col_des].text.strip() if col_des is not None else ""
        nit = _norm_nit(row.cells[col_nit].text) if col_nit is not None else ""
        if des or nit:
            data.append((r, des))
        elif total is None:
            total = r

    def _val(bal, code):
        return abs(bal.saldo(code))

    def _set(cell, val):
        ex = parse_money(cell.text)
        if ex is not None and abs(ex - val) <= ROUND_TOL:
            return
        a = cell.text.strip()
        n = format_like(val, cell)
        if a != n:
            set_cell_value(cell, n)
            rep.chg(nota, cfg["sig"], "impuesto", "col", a, n)

    def _set_row(r, code):
        c26, c25 = b26.cuentas.get(code), None
        for col, role in roles.items():
            if role == "mov":
                val = abs(c26.debito - c26.credito) if c26 else 0.0
            else:
                val = _val(b25 if role == "comparativo" else b26, code)
            _set(table.rows[r].cells[col], val)

    # Subcuenta(s) que ESTA tabla representa, según su TÍTULO (sin el texto repetido
    # 'IMPUESTO(S) POR PAGAR'): p. ej. '... - IVA' -> 2408, '... - RENTA' -> 2404.
    boiler = {"IMPUESTO", "IMPUESTOS", "POR", "PAGAR", "ANTICIPO", "CONTRIBUCIONES",
              "SALDOS", "FAVOR", "PESOS", "EXPRESADO"}
    titulo = table.rows[0].cells[0].text.upper()
    if re.search(r"\bIVA\b", titulo):
        titulo += " VENTAS"           # IVA = impuesto a las VENTAS (2408)
    ttoks = _toks(titulo) - boiler
    allsubs = _subcuentas_valor(b26, b25, clases, niveles=(4, 6))
    anchor, aov = None, 0
    for c in allsubs.values():
        ov = len(ttoks & (_toks(c.nombre) - boiler))
        if ov > aov or (ov == aov and anchor and ov > 0 and len(c.codigo) < len(anchor)):
            aov, anchor = ov, c.codigo
    if aov == 0:
        anchor = None

    # Subcuentas de presentación del ámbito: nivel 6 díg con saldo y, para cuentas de
    # 4 díg que no tienen hijas de 6 díg con saldo, la propia de 4 díg.
    def _leaves(pref):
        with_val = {k for k in allsubs
                    if (pref is None or k.startswith(pref))
                    and (abs(b26.saldo(k)) > 0 or abs(b25.saldo(k)) > 0)}
        sixes = [k for k in with_val if len(k) == 6]
        fours = [k for k in with_val if len(k) == 4 and not any(s.startswith(k) for s in sixes)]
        return sorted(sixes + fours)

    leaves = _leaves(anchor)

    if anchor and len(data) == 1:
        # Tabla de UNA subcuenta (p. ej. 'IMPUESTO POR PAGAR - IVA'): se usa el saldo
        # de la cuenta de 4 díg (NETO de IVA generado/descontable, etc.).
        _set_row(data[0][0], anchor)
    elif data and len(data) == len(leaves):
        # 1-a-1: por descripción y, lo que sobre, por eliminación (resuelve
        # 'AUTORRETENCION' -> 135515 cuando hay tantas filas como subcuentas).
        libres = list(leaves)
        pend = []
        for r, des in data:
            sub = _best_sub([c for k, c in allsubs.items() if k in libres], des)
            if sub is not None:
                _set_row(r, sub.codigo)
                libres.remove(sub.codigo)
            else:
                pend.append(r)
        for r in pend:
            if libres:
                _set_row(r, libres.pop(0))
    elif anchor and data:
        # UNA subcuenta con varias filas (p. ej. vigencias de renta que ya suman la
        # subcuenta): se dejan las filas (do-no-harm) y solo se cuadra el total.
        pass
    else:
        # Genérica (varias subcuentas por nombre, p. ej. cuenta 24): cada fila a su
        # subcuenta por descripción si la coincidencia es clara.
        for r, des in data:
            sub = _best_sub([c for c in allsubs.values()], des)
            if sub is not None and len(_toks(des) & _toks(sub.nombre)) >= 1:
                _set_row(r, sub.codigo)

    if total is not None:
        for col, role in roles.items():
            if anchor and not (data and len(data) == len(leaves)):
                val = (abs(b26.cuentas[anchor].debito - b26.cuentas[anchor].credito)
                       if role == "mov" and anchor in b26.cuentas
                       else _val(b25 if role == "comparativo" else b26, anchor))
            else:
                val = sum(parse_money(table.rows[r].cells[col].text) or 0.0 for r, _ in data)
            _set(table.rows[total].cells[col], val)


def _cuenta_class(title, default):
    """Algunas tablas de tipo 'cuenta' se presentan bajo una nota pero pertenecen
    a otra clase del PUC (p. ej. VALORIZACIONES = clase 19, aunque se muestre en la
    Nota 3 junto a Propiedad, Planta y Equipo)."""
    if "VALORIZAC" in (title or "").upper():
        return "19"
    return default


def process_cuenta(table, cfg, b26, b25, rep: Report):
    """Notas a nivel de cuenta (sin terceros), p. ej. Efectivo / Inmuebles /
    Valorizaciones. Cada fila se mapea a su subcuenta por la DESCRIPCIÓN (columna
    'DESCRIPCION DE CUENTA', no la columna 'NOTA'). El total va a la clase de la
    tabla. Cuando una fila toma una cuenta de nivel superior y sus hijas también
    figuran como filas, se le RESTA lo asignado a esas hijas (evita el doble conteo
    que ponía el total de la clase en todas las filas)."""
    nota = cfg.get("nota", "")
    header_r = _find_header_row(table)
    if header_r is None:
        return
    hdr = table.rows[header_r]
    roles = {i: _role_of(c.text) for i, c in enumerate(hdr.cells) if _role_of(c.text)}
    if not roles:
        return
    col_des = _col_index(hdr, "DESCRIPCION", "DESCRICION", "DES CUENTA", "DES CUE")
    if col_des is None:                       # 1ª columna de texto que no sea NOTA ni de valor
        for i, c in enumerate(hdr.cells):
            if i not in roles and "NOTA" not in c.text.upper():
                col_des = i
                break
    parent = _cuenta_class(table.rows[0].cells[0].text, cfg["account"])
    subs = [c for c in b26.cuentas.values() if c.codigo.startswith(parent) and c.codigo != parent]
    seen = {c.codigo for c in subs}
    subs += [c for c in b25.cuentas.values()
             if c.codigo.startswith(parent) and c.codigo != parent and c.codigo not in seen]

    data, total = [], None
    for r in range(header_r + 1, len(table.rows)):
        row = table.rows[r]
        if not any(parse_money(c.text) is not None for _, c in _unique_grid(row)):
            continue
        desc = row.cells[col_des].text.strip() if col_des is not None else row.cells[0].text.strip()
        if desc:
            data.append((r, desc))
        elif total is None:
            total = r

    matched = {}
    for r, desc in data:
        sub = _best_sub(subs, desc)
        matched[r] = sub.codigo if sub else parent

    def _saldo(bal, code, others):
        v = abs(bal.saldo(code))
        for oc in others:                     # resta hijas asignadas a OTRAS filas
            if oc != code and oc.startswith(code):
                v -= abs(bal.saldo(oc))
        return v

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

    for r, desc in data:
        code = matched[r]
        others = [c for rr, c in matched.items() if rr != r]
        _set(table.rows[r], desc,
             lambda role, code=code, others=others:
                 _saldo(b25 if role == "comparativo" else b26, code, others))
    if total is not None:
        _set(table.rows[total], "TOTAL",
             lambda role: abs((b25 if role == "comparativo" else b26).saldo(parent)))


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
#  Configuración por NOTA: a cada nota del Word le corresponde una CLASE del PUC.
#
#  El recorrido sigue los encabezados "NOTA n" / "VALOR TOTAL NOTA n" y procesa
#  TODAS las sub-tablas de cada nota (sin importar su título ni cuántas haya),
#  detectando el tipo de cada una (terceros / cuenta / contratos). Cada sub-tabla
#  se asocia sola a su(s) cuenta(s) dentro de la clase de la nota. Así funciona
#  con sociedades que tienen más cuentas/terceros (p. ej. CIA).
#
#    clase  -> prefijo(s) de clase del PUC para los terceros de la nota
#    vtotal -> clase para el cuadro "VALOR TOTAL NOTA n"
# --------------------------------------------------------------------------- #

NOTE_CONFIG = {
    "1": {"clase": ["11"], "vtotal": "11"},     # Efectivo
    "2": {"clase": ["13"], "vtotal": "13"},     # Cuentas por cobrar
    "3": {"clase": ["15"], "vtotal": "15"},     # Propiedad, planta y equipo (inmuebles)
    "3.1": {"clase": ["17"], "vtotal": "17"},   # Gastos diferidos / pagados por anticipado
    "4": {"clase": ["2"],  "vtotal": "2"},      # Acreedores / cuentas por pagar
    "6": {"clase": ["41"], "vtotal": "41"},     # Ingresos operacionales
    "7": {"clase": ["51"], "vtotal": "51"},     # Gastos de administración
    "8": {"clase": ["53"], "vtotal": "53"},     # Gastos no operacionales
    "9": {"clase": ["42"], "vtotal": "42"},     # Ingresos no operacionales
}

# Nota 5 (Patrimonio = clase 3) y Nota 10 (Cuentas de orden) se hacen manual.
SKIP_NOTES = {"5", "10"}


def _detect_kind(table, header_r):
    """Tipo de una sub-tabla por sus columnas: contratos (cánon+arrendatario),
    terceros (tiene cédula/NIT con datos) o cuenta (solo descripción)."""
    hdr = table.rows[header_r]
    if _col_index(hdr, "CÁNON", "CANON") is not None and _col_index(hdr, "ARRENDATARIO") is not None:
        return "contratos"
    col_nit = _col_index(hdr, "CEDULA", "CÉDULA", "ID,")
    if col_nit is not None:
        for r in range(header_r + 1, len(table.rows)):
            if _norm_nit(table.rows[r].cells[col_nit].text):
                return "terceros"
    return "cuenta"


def _nota_en_fila(table):
    """Devuelve el número de nota si la tabla contiene una fila de encabezado
    'NOTA | n' (sea tabla de encabezado o fila embebida en un VALOR TOTAL)."""
    nota = None
    for row in table.rows:
        cells = [c.text.strip() for c in _unique_grid_cells(row)]
        if cells and cells[0].upper() == "NOTA":
            for c in cells[1:]:
                # Acepta notas enteras ("4") y sub-notas decimales ("3.1").
                if re.fullmatch(r"\d{1,2}(\.\d+)?", c):
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


def process_contratos(table, cfg, b26, b25, rep: Report):
    """Tabla de detalle de contratos de arrendamiento (Nota 6). Toma de los
    ingresos (4155): cánon mensual y mes actual = MOVIMIENTO DEL MES; total
    acumulado año = ACUMULADO MES ACTUAL.

    Empareja cada fila por NOMBRE del arrendatario primero (en algunas
    sociedades las cédulas de esta tabla están corridas/erradas) y por cédula
    como respaldo. Las filas sin pareja no se tocan, pero sí suman al total."""
    nota = cfg.get("nota", "6")
    header_r = _find_header_row(table)
    if header_r is None:
        return
    hdr = table.rows[header_r]
    col_ced = _col_index(hdr, "CEDULA", "CÉDULA")
    col_arr = _col_index(hdr, "ARRENDATARIO")
    col_canon = _col_index(hdr, "CÁNON", "CANON")
    col_total = None
    for i, c in enumerate(hdr.cells):
        u = c.text.upper()
        if "TOTAL" in u and ("ACUMULAD" in u or re.search(r"A[NÑ]O", u)):
            col_total = i
            break
    if col_canon is None or col_total is None or (col_ced is None and col_arr is None):
        return

    recs = [r for r in scope_records(b26, cfg["scope"]) if r["kind"] == "t"]
    by_nit = {r["nit"]: r for r in recs}

    def match_row(r):
        name = table.rows[r].cells[col_arr].text if col_arr is not None else ""
        nt = _toks(name)
        if nt:
            best, bs = None, 0
            for rec in recs:
                sc = len(nt & _toks(rec["t"].nombre))
                if sc > bs:
                    bs, best = sc, rec
            if best is not None and bs >= min(2, len(nt)):
                return best
        if col_ced is not None:
            return by_nit.get(_norm_nit(table.rows[r].cells[col_ced].text))
        return None

    # Filas de datos: cánon con valor y arrendatario con TEXTO (letras). Esto
    # excluye filas ajenas embebidas en la misma tabla (cuadros "VALOR TOTAL",
    # encabezados de mes, etc.).
    data, total_row = [], None
    for r in range(header_r + 1, len(table.rows)):
        rowtxt = " ".join(c.text for c in _unique_grid_cells(table.rows[r])).upper()
        if "VALOR TOTAL" in rowtxt:
            continue
        canon_ok = parse_money(table.rows[r].cells[col_canon].text) is not None
        arr = table.rows[r].cells[col_arr].text.strip() if col_arr is not None else ""
        arr_es_texto = bool(re.search(r"[A-ZÁÉÍÓÚÑ]{3,}", arr.upper()))
        if canon_ok and arr_es_texto:
            data.append(r)
        elif total_row is None and not arr_es_texto and \
                parse_money(table.rows[r].cells[col_total].text) is not None:
            total_row = r

    if not data:
        return
    valcols = [i for i, c in _unique_grid(table.rows[data[0]]) if parse_money(c.text) is not None]
    col_month = next((i for i in valcols if col_canon < i < col_total), None)
    if col_month is not None:                       # encabezado del mes actual
        set_cell_value(hdr.cells[col_month], b26.periodo)

    def _set(cell, val):
        ex = parse_money(cell.text)
        if ex is not None and abs(ex - val) <= ROUND_TOL:
            return
        a = cell.text.strip()
        n = format_like(val, cell)
        if a != n:
            set_cell_value(cell, n)
            rep.chg(nota, cfg["sig"], "contrato", "col", a, n)

    cols = [c for c in (col_canon, col_month, col_total) if c is not None]
    for r in data:
        rec = match_row(r)
        if rec is None:
            continue
        vals = {col_canon: rec["mov"], col_month: rec["mov"], col_total: abs(rec["value"])}
        for col in cols:
            _set(table.rows[r].cells[col], vals[col])

    if total_row is not None:                        # total = suma de TODAS las filas
        for col in cols:
            _set(table.rows[total_row].cells[col],
                 sum(parse_money(table.rows[dr].cells[col].text) or 0 for dr in data))


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
    col_des = _col_index(hdr, "DES CUENTA", "DESCRIPCION", "DESCRICION", "DES CUE")
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


_PROCESSORS = {
    "terceros": process_terceros,
    "cuenta": process_cuenta,
    "patrimonio": process_patrimonio,
    "contratos": process_contratos,
}


# --------------------------------------------------------------------------- #
#  Actualización de FECHAS del período (rueda el mes: marzo -> abril, etc.)
# --------------------------------------------------------------------------- #

_MESES = ["ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO", "JUNIO", "JULIO",
          "AGOSTO", "SEPTIEMBRE", "OCTUBRE", "NOVIEMBRE", "DICIEMBRE"]
_MES_NUM = {m: i + 1 for i, m in enumerate(_MESES)}
_MES_NUM["SETIEMBRE"] = 9


def _case_like(target: str, sample: str) -> str:
    if sample.isupper():
        return target.upper()
    if sample.islower():
        return target.lower()
    return target.capitalize()


def _periodo_info(periodo: str):
    m = re.search(r"([A-ZÁÉÍÓÚ]+)\s+(\d{4})", (periodo or "").upper())
    if not m or m.group(1) not in _MES_NUM:
        return None
    mn = _MES_NUM[m.group(1)]
    yr = int(m.group(2))
    pm = mn - 1 or 12
    return dict(num=mn, name=_MESES[mn - 1], year=yr,
                last=calendar.monthrange(yr, mn)[1], prior_name=_MESES[pm - 1])


def _roll_text(s: str, P) -> str:
    """Cambia SOLO fechas del período al período actual:
      * 'último-día de <mes-anterior> de <año actual>'  -> '<último-día> de <mes actual> de <año>'
        (no toca fechas históricas: distinto año, o día que no es fin de mes)
      * '<mes anterior> [de] <2025|2026>' (encabezados) -> '<mes actual> ...'
        (no toca '... de <mes> de ...' por el lookbehind)."""
    mes_re = "|".join(_MESES + ["SETIEMBRE"])

    def r1(m):
        d, mes, conn, yr = int(m.group(1)), m.group(2), m.group(3), int(m.group(4))
        mn = _MES_NUM.get(mes.upper())
        if mn and yr == P["year"] and mn < P["num"] and d == calendar.monthrange(yr, mn)[1]:
            return f"{P['last']} de {_case_like(P['name'], mes)} {conn} {P['year']}"
        return m.group(0)

    # Conector mes->año flexible: 'de' o 'del' (se conserva el del original).
    s = re.sub(rf"\b(\d{{1,2}})\s+de\s+({mes_re})\s+(del?)\s+(\d{{4}})\b", r1, s, flags=re.I)

    # R3: '<mes anterior> de <año actual>' en texto (p. ej. 'al cierre del mes de
    # marzo de 2026') -> mes actual. NUNCA si va precedido por un día ('1 de marzo
    # de 2026' = fecha histórica/contrato) ni en el año comparativo.
    def r3(m):
        return _case_like(P["name"], m.group(1)) + " " + m.group(2) + " " + m.group(3)

    s = re.sub(rf"(?<!\d de )\b({P['prior_name']}) (del?) ({P['year']})\b", r3, s, flags=re.I)

    def r2(m):
        return _case_like(P["name"], m.group(1)) + (m.group(2) or "") + m.group(3) + m.group(4)

    s = re.sub(rf"(?<!de )\b({P['prior_name']})(\s+del?)?(\s+)({P['year']}|{P['year'] - 1})\b",
               r2, s, flags=re.I)
    return s


def actualizar_fechas(doc, periodo: str, rep=None) -> int:
    """Rueda las fechas del período en narrativa y tablas, preservando formato."""
    P = _periodo_info(periodo)
    if not P:
        return 0
    n = 0

    def proc(paras):
        nonlocal n
        for p in paras:
            if not p.runs:
                continue
            full = "".join(r.text for r in p.runs)
            new = _roll_text(full, P)
            if new == full:
                continue
            # Solo se reescribe la REGIÓN mínima que cambió (prefijo/sufijo común);
            # los runs fuera de esa región conservan su texto y formato intactos, y
            # el texto nuevo se deposita en el primer run que la toca. Así el formato
            # sobrevive aunque cambie la longitud (p. ej. 'abril' -> 'mayo').
            pre = 0
            while pre < len(full) and pre < len(new) and full[pre] == new[pre]:
                pre += 1
            suf = 0
            while (suf < len(full) - pre and suf < len(new) - pre
                   and full[-1 - suf] == new[-1 - suf]):
                suf += 1
            mid_new = new[pre:len(new) - suf]
            start = len(full) - suf
            placed, pos = False, 0
            for run in p.runs:
                a, b = pos, pos + len(run.text)
                pos = b
                if b <= pre or a >= start:
                    continue                      # run fuera de la región cambiada
                keep_left = run.text[:max(0, pre - a)]
                keep_right = run.text[max(0, start - a):] if b > start else ""
                run.text = keep_left + (mid_new if not placed else "") + keep_right
                placed = True
            n += 1

    proc(doc.paragraphs)
    for t in doc.tables:
        for row in t.rows:
            seen = set()
            for cell in row.cells:
                if id(cell._tc) in seen:
                    continue
                seen.add(id(cell._tc))
                proc(cell.paragraphs)
    if rep is not None:
        rep.date_changes = n
    return n


def update_document(doc, b26: Balance, b25: Balance, skip_notes=SKIP_NOTES, crear_faltantes=False) -> Report:
    rep = Report()
    current_note = None
    info = {n: {"template": None, "anchor": None, "matched": set(), "accounts": set(),
                "titles": [], "nits": set()} for n in NOTE_CONFIG}

    for table in doc.tables:
        title = table.rows[0].cells[0].text.strip()
        up = title.upper()

        m = re.search(r"VALOR TOTAL NOTA\s*0*(\d+(?:\.\d+)?)", up)
        if m:                                   # cuadro "VALOR TOTAL NOTA n"
            n = m.group(1)
            cfg = NOTE_CONFIG.get(n, {})
            if n in info:
                info[n]["anchor"] = table._tbl
            if n not in skip_notes and cfg.get("vtotal"):
                process_vtotal(table, {"account": cfg["vtotal"], "nota": n, "sig": title}, b26, b25, rep)
        else:
            hr = _find_header_row(table)
            if hr is not None and current_note in NOTE_CONFIG and current_note not in skip_notes:
                n = current_note
                clase = NOTE_CONFIG[n]["clase"]
                kind = _detect_kind(table, hr)
                info[n]["titles"].append(up)        # título de CUALQUIER sub-tabla
                if kind == "terceros" and _es_tabla_impuestos(table, hr, b26, b25, n):
                    # Tabla de impuestos: NUNCA se agregan terceros. Si sus filas son
                    # terceros DISTINTOS (p. ej. predial por municipio) se actualiza
                    # cada uno (process_terceros sin altas); si es una subcuenta con
                    # el mismo NIT repetido / una sola fila (IVA, renta, anticipo) se
                    # toma el valor de la SUBCUENTA (process_impuestos).
                    if _nits_distintos(table, hr):
                        process_terceros(table, dict(nota=n, sig=title, scope=clase, add=False),
                                         b26, b25, rep)
                    else:
                        process_impuestos(table, dict(nota=n, sig=title, scope=clase), b26, b25, rep)
                elif kind == "terceros":
                    info[n]["template"] = table
                    cnit = _col_index(table.rows[hr], "CEDULA", "CÉDULA", "ID,")
                    if cnit is not None:               # NITs presentes en el Word
                        for rr in range(hr + 1, len(table.rows)):
                            nn = _norm_nit(table.rows[rr].cells[cnit].text)
                            if nn:
                                info[n]["nits"].add(nn)
                    matched, taccts = process_terceros(table, dict(nota=n, sig=title, scope=clase), b26, b25, rep)
                    info[n]["matched"] |= matched
                    info[n]["accounts"] |= taccts
                elif kind == "contratos":
                    process_contratos(table, dict(nota=n, sig=title, scope=clase), b26, b25, rep)
                else:
                    process_cuenta(table, dict(nota=n, sig=title, account=clase[0]), b26, b25, rep)

        n_aqui = _nota_en_fila(table)
        if n_aqui:
            current_note = n_aqui

    # Cuentas (4 díg) de la clase con terceros que NINGUNA sub-tabla cubrió:
    # se crea la tabla desde el auxiliar (deseo del usuario), SALVO que ya exista
    # en la nota una tabla con título similar (sería un duplicado: la sociedad
    # presenta esa cuenta con otra estructura) — en ese caso solo se avisa.
    for n, cfg in NOTE_CONFIG.items():
        if n in skip_notes or info[n]["template"] is None or info[n]["anchor"] is None:
            continue
        por4 = {}
        for r in scope_records(b26, cfg["clase"]):
            if r["kind"] == "t":
                por4.setdefault(r["code"][:4], []).append(r)
        for c4, terc in por4.items():
            if c4 in info[n]["accounts"]:
                continue
            if all(t["tid"] in info[n]["matched"] for t in terc):
                continue
            # Si los terceros de la cuenta YA aparecen (por NIT) en alguna tabla de
            # la nota, la cuenta ya está presentada aunque el emparejamiento por
            # valor no la haya "cubierto": no se recrea (evita tablas duplicadas).
            if terc and all(t["nit"] in info[n]["nits"] for t in terc):
                continue
            if not any(abs(t["value"]) >= 1 or t["mov"] >= 1 for t in terc):
                continue
            titulo = b26.cuentas[c4].nombre if c4 in b26.cuentas else c4
            ntoks = _toks(titulo)
            similar = any(len(ntoks & _toks(t)) >= min(2, len(ntoks)) for t in info[n]["titles"])
            # Cuentas que se netean (p. ej. retención causada vs pagada): la
            # presentación es a criterio del contador -> solo avisar.
            netea = _es_neteo(scope_records(b26, [c4]), c4) or _es_neteo(scope_records(b25, [c4]), c4)
            if similar or netea or not crear_faltantes:
                rep.flags.append(f"Nota {n}: la cuenta {c4} '{titulo}' tiene datos en el auxiliar "
                                 f"pero el Word la presenta distinto (revisar manualmente).")
            else:
                crear_tabla(info[n]["template"], info[n]["anchor"], titulo, [c4], b26, b25, n, rep)

    # Rodar las fechas del período (marzo -> abril, etc.) en narrativa y tablas.
    actualizar_fechas(doc, b26.periodo, rep)
    return rep
