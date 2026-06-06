"""
Parser de Balance de Comprobación (auxiliar) en PDF -> datos estructurados.

Los PDF "AJUSTADO" exportados por CAFESOFT-ERP tienen, por cada cuenta del PUC y
por cada tercero, las columnas:  Saldo Anterior | Débito | Crédito | Nuevo Saldo.
Donde Débito/Crédito son los MOVIMIENTOS del período y Nuevo Saldo el ACUMULADO.

El parser reconstruye el texto desde los streams del PDF (sin librerías externas),
agrupa por fila según la coordenada Y y asigna cada número a su columna según la
coordenada X (las 4 columnas de valores están alineadas a la derecha en 4 bandas).
"""
from __future__ import annotations
import re
import zlib
from dataclasses import dataclass, field

# ------------------------- extracción de texto del PDF -------------------------

def _decompress_streams(path: str) -> list[bytes]:
    data = open(path, "rb").read()
    out = []
    for s in re.findall(rb"stream[ \t]*\r?\n(.*?)endstream", data, re.DOTALL):
        for cand in (s.rstrip(b"\r\n"), s, s.strip()):
            try:
                out.append(zlib.decompress(cand))
                break
            except Exception:
                continue
    return out


def _decode_pdf_string(s: str) -> str:
    res, i = [], 0
    while i < len(s):
        c = s[i]
        if c == "\\":
            i += 1
            if i >= len(s):
                break
            n = s[i]
            if n in "nrtbf":
                res.append({"n": "\n", "r": "\r", "t": "\t", "b": "\b", "f": "\f"}[n]); i += 1
            elif n in "()\\":
                res.append(n); i += 1
            elif n in "01234567":
                octd = n; i += 1
                for _ in range(2):
                    if i < len(s) and s[i] in "01234567":
                        octd += s[i]; i += 1
                    else:
                        break
                res.append(chr(int(octd, 8)))
            else:
                res.append(n); i += 1
        else:
            res.append(c); i += 1
    return "".join(res)


def _tokenize(content: str):
    tokens, i, n = [], 0, len(content)
    while i < n:
        c = content[i]
        if c in " \t\r\n":
            i += 1; continue
        if c == "(":
            depth, j, buf = 1, i + 1, []
            while j < n and depth > 0:
                ch = content[j]
                if ch == "\\":
                    buf.append(ch)
                    if j + 1 < n:
                        buf.append(content[j + 1]); j += 2; continue
                    j += 1; continue
                if ch == "(":
                    depth += 1
                elif ch == ")":
                    depth -= 1
                    if depth == 0:
                        break
                buf.append(ch); j += 1
            tokens.append(("STR", _decode_pdf_string("".join(buf)))); i = j + 1; continue
        if c == "[":
            tokens.append(("AOPEN", None)); i += 1; continue
        if c == "]":
            tokens.append(("ACLOSE", None)); i += 1; continue
        if c == "<":
            j = content.find(">", i)
            try:
                b = bytes.fromhex(content[i + 1:j].replace(" ", "").replace("\n", ""))
                tokens.append(("STR", b.decode("latin-1", "ignore")))
            except Exception:
                tokens.append(("STR", ""))
            i = j + 1; continue
        if c == "/":
            j = i + 1
            while j < n and content[j] not in " \t\r\n/[]<>()":
                j += 1
            tokens.append(("NAME", content[i:j])); i = j; continue
        j = i
        while j < n and content[j] not in " \t\r\n/[]<>()":
            j += 1
        tok = content[i:j]
        try:
            tokens.append(("NUM", float(tok)))
        except ValueError:
            tokens.append(("OP", tok))
        i = j
    return tokens


def _page_items(content: str):
    """Devuelve [(y, x, texto)] de los operadores de texto de una página."""
    toks = _tokenize(content)
    items, stack = [], []
    tm = [1, 0, 0, 1, 0, 0]
    tlm = [1, 0, 0, 1, 0, 0]
    leading = 0.0
    for kind, val in toks:
        if kind in ("NUM", "STR", "NAME"):
            stack.append(val); continue
        if kind in ("AOPEN", "ACLOSE"):
            if kind == "AOPEN":
                stack.append("[")
            continue
        if kind != "OP":
            continue
        op = val
        if op == "BT":
            tm = [1, 0, 0, 1, 0, 0]; tlm = tm[:]; stack = []
        elif op in ("ET", "Tf"):
            stack = []
        elif op == "TL":
            if stack:
                leading = stack[-1]
            stack = []
        elif op in ("Td", "TD"):
            if len(stack) >= 2:
                tx, ty = stack[-2], stack[-1]
                tlm = [tlm[0], tlm[1], tlm[2], tlm[3], tlm[4] + tx * tlm[0], tlm[5] + ty * tlm[3]]
                tm = tlm[:]
                if op == "TD":
                    leading = -ty
            stack = []
        elif op == "Tm":
            if len(stack) >= 6:
                tlm = [x for x in stack[-6:]]; tm = tlm[:]
            stack = []
        elif op == "T*":
            tlm = [tlm[0], tlm[1], tlm[2], tlm[3], tlm[4], tlm[5] - leading]; tm = tlm[:]
            stack = []
        elif op in ("Tj", "'", '"'):
            if op in ("'", '"'):
                tlm = [tlm[0], tlm[1], tlm[2], tlm[3], tlm[4], tlm[5] - leading]; tm = tlm[:]
            s = stack[-1] if stack and isinstance(stack[-1], str) else ""
            if s:
                items.append((round(tm[5], 1), round(tm[4], 1), s))
            stack = []
        elif op == "TJ":
            parts = [x for x in stack if isinstance(x, str) and x != "["]
            s = "".join(parts)
            if s:
                items.append((round(tm[5], 1), round(tm[4], 1), s))
            stack = []
        else:
            stack = []
    return items


# ------------------------- estructura de datos -------------------------

# Bandas X (coordenada inicial del texto) de las 4 columnas de valores.
_COL_BANDS = [(285, 375), (375, 455), (455, 513), (513, 9999)]


def _num(s: str):
    s = s.strip().replace(",", "")
    if not s or not re.search(r"\d", s):
        return None
    try:
        return float(s)
    except ValueError:
        return None


@dataclass
class Tercero:
    nit: str
    nombre: str
    saldo_ant: float = 0.0
    debito: float = 0.0
    credito: float = 0.0
    nuevo_saldo: float = 0.0


@dataclass
class Cuenta:
    codigo: str
    nombre: str
    saldo_ant: float = 0.0
    debito: float = 0.0
    credito: float = 0.0
    nuevo_saldo: float = 0.0
    terceros: list[Tercero] = field(default_factory=list)


@dataclass
class Balance:
    periodo: str
    cuentas: dict[str, Cuenta]

    # ---- consultas ----
    def cuenta(self, codigo: str) -> Cuenta | None:
        return self.cuentas.get(codigo)

    def saldo(self, codigo: str) -> float:
        c = self.cuentas.get(codigo)
        return c.nuevo_saldo if c else 0.0

    def cuentas_con_prefijo(self, prefijo: str) -> list[Cuenta]:
        return [c for k, c in self.cuentas.items() if k.startswith(prefijo)]

    def hojas_con_prefijo(self, prefijo: str) -> list[Cuenta]:
        """Cuentas con ese prefijo que tienen terceros (cuentas auxiliares hoja)."""
        return [c for c in self.cuentas_con_prefijo(prefijo) if c.terceros]

    def terceros_de(self, prefijo: str) -> list[Tercero]:
        out = []
        for c in self.cuentas_con_prefijo(prefijo):
            out.extend(c.terceros)
        return out


def _norm_nit(s: str) -> str:
    return re.sub(r"\D", "", s)


def parse_pdf(path: str) -> Balance:
    streams = [s.decode("latin-1", "ignore") for s in _decompress_streams(path) if b"BT" in s]
    periodo = ""
    cuentas: dict[str, Cuenta] = {}
    last_leaf: Cuenta | None = None       # última cuenta hoja (para colgar terceros)
    last_obj = None                        # último objeto creado (cuenta o tercero) p/continuación de nombre

    for content in streams:
        items = _page_items(content)
        # agrupar por fila (Y)
        rows: dict[float, list] = {}
        for (y, x, s) in items:
            rows.setdefault(round(y / 2.5) * 2.5, []).append((x, s))

        for y in sorted(rows.keys(), reverse=True):
            cells = sorted(rows[y], key=lambda c: c[0])
            if not periodo:
                joined = " ".join(s for _, s in cells)
                m = re.search(r"Per[ií]odo:\s*([A-ZÁÉÍÓÚ]+\s*\d{4})", joined)
                if m:
                    periodo = m.group(1).strip()

            # separar columnas de valores (x>=285) del área izquierda
            valores = [None, None, None, None]
            izq = []
            for (x, s) in cells:
                v = _num(s)
                if x >= 285 and v is not None:
                    for bi, (a, b) in enumerate(_COL_BANDS):
                        if a <= x < b:
                            valores[bi] = v
                            break
                else:
                    izq.append((x, s))

            has_values = any(v is not None for v in valores)
            # Código/NIT está a la izquierda (x<110); el nombre a x>=110.
            # El marcador de tercero es el token de sucursal "000" (~x=117), que
            # comparte banda con el nombre de cuenta pero es numérico.
            code_area = [s for (x, s) in izq if x < 110]
            name_area = [s for (x, s) in izq if x >= 110 and s.strip() != "000"]
            izq_left = " ".join(s for (x, s) in izq if x < 135)
            code_join = " ".join(code_area).strip()
            name = " ".join(t.strip() for t in name_area).strip()

            if not has_values:
                # ¿continuación de nombre? (texto sin valores y sin código nuevo)
                cont = (code_join + " " + name).strip()
                if cont and last_obj is not None and not re.match(r"^\d", code_join):
                    last_obj.nombre = (last_obj.nombre + " " + cont).strip()
                continue

            sa, de, cr, ns = (v or 0.0 for v in valores)
            es_tercero = bool(re.search(r"(?<!\d)000(?!\d)", izq_left))
            digits = re.findall(r"\d+", code_join)

            if es_tercero and last_leaf is not None and digits:
                nit = _norm_nit(digits[0])
                t = Tercero(nit=nit, nombre=name, saldo_ant=sa, debito=de, credito=cr, nuevo_saldo=ns)
                last_leaf.terceros.append(t)
                last_obj = t
            elif digits and not es_tercero:
                codigo = digits[0]
                c = Cuenta(codigo=codigo, nombre=name, saldo_ant=sa, debito=de, credito=cr, nuevo_saldo=ns)
                cuentas[codigo] = c
                last_leaf = c
                last_obj = c

    return Balance(periodo=periodo or "?", cuentas=cuentas)


if __name__ == "__main__":
    import sys
    b = parse_pdf(sys.argv[1])
    print("Periodo:", b.periodo, "| cuentas:", len(b.cuentas))
    for code in ["1", "2", "3", "4", "5"]:
        c = b.cuenta(code)
        if c:
            print(f"  {code} {c.nombre:30} ant={c.saldo_ant:,.2f} db={c.debito:,.2f} cr={c.credito:,.2f} nuevo={c.nuevo_saldo:,.2f}")
