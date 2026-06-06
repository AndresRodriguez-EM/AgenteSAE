"""
Genera el logo de AgenteSAE (vectorial, con Qt) en varios tamaños:
  assets/icon.png  (256)  y  assets/icon.ico  (multi-tamaño para Windows).

Concepto: documento de "notas" con líneas y una insignia verde de verificación,
sobre el azul corporativo de los estados financieros (#07314A).

Uso:  python tools/make_logo.py
"""
import os
import struct

from PySide6 import QtCore, QtGui, QtWidgets

AZUL_TOP = "#2A8FD4"
AZUL_BASE = "#07314A"
VERDE = "#16A34A"
ACENTO = "#2A8FD4"
LINEA = "#B7C4D1"

OUT = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "assets")


def _round_rect_path(x, y, w, h, r):
    p = QtGui.QPainterPath()
    p.addRoundedRect(QtCore.QRectF(x, y, w, h), r, r)
    return p


def draw(painter: QtGui.QPainter):
    """Dibuja el icono en un lienzo lógico de 256x256."""
    painter.setRenderHint(QtGui.QPainter.Antialiasing, True)

    # Fondo redondeado con degradado azul
    grad = QtGui.QLinearGradient(0, 0, 0, 256)
    grad.setColorAt(0, QtGui.QColor(AZUL_TOP))
    grad.setColorAt(1, QtGui.QColor(AZUL_BASE))
    painter.setPen(QtCore.Qt.NoPen)
    painter.setBrush(QtGui.QBrush(grad))
    painter.drawPath(_round_rect_path(8, 8, 240, 240, 52))

    # Sombra del documento
    painter.setBrush(QtGui.QColor(0, 0, 0, 60))
    painter.drawPath(_round_rect_path(81, 57, 100, 150, 14))

    # Documento (blanco)
    painter.setBrush(QtGui.QColor("#FFFFFF"))
    painter.drawPath(_round_rect_path(76, 50, 100, 150, 14))

    # Esquina doblada (arriba-derecha)
    fold = QtGui.QPainterPath()
    fold.moveTo(150, 50); fold.lineTo(176, 50); fold.lineTo(176, 76); fold.closeSubpath()
    painter.setBrush(QtGui.QColor("#C9D6E2"))
    painter.drawPath(fold)

    # Líneas de texto
    def line(y, x2, color):
        painter.setBrush(QtGui.QColor(color))
        painter.drawPath(_round_rect_path(92, y, x2 - 92, 9, 4))
    line(82, 150, ACENTO)     # línea acento (azul)
    line(104, 160, LINEA)
    line(122, 160, LINEA)
    line(140, 150, LINEA)
    line(158, 132, LINEA)

    # Insignia verde de verificación (con aro blanco separador)
    painter.setBrush(QtGui.QColor("#FFFFFF"))
    painter.drawEllipse(QtCore.QPointF(178, 178), 42, 42)
    painter.setBrush(QtGui.QColor(VERDE))
    painter.drawEllipse(QtCore.QPointF(178, 178), 35, 35)

    # Chulo (check)
    pen = QtGui.QPen(QtGui.QColor("#FFFFFF"))
    pen.setWidthF(11)
    pen.setCapStyle(QtCore.Qt.RoundCap)
    pen.setJoinStyle(QtCore.Qt.RoundJoin)
    painter.setPen(pen)
    chk = QtGui.QPainterPath()
    chk.moveTo(161, 179); chk.lineTo(173, 192); chk.lineTo(197, 165)
    painter.drawPath(chk)


def render_png(size: int) -> QtGui.QImage:
    img = QtGui.QImage(size, size, QtGui.QImage.Format_ARGB32)
    img.fill(QtCore.Qt.transparent)
    p = QtGui.QPainter(img)
    p.scale(size / 256.0, size / 256.0)
    draw(p)
    p.end()
    return img


def png_bytes(img: QtGui.QImage) -> bytes:
    buf = QtCore.QBuffer()
    buf.open(QtCore.QIODevice.WriteOnly)
    img.save(buf, "PNG")
    return bytes(buf.data())


def write_ico(path: str, images_png: list[tuple[int, bytes]]):
    """Empaqueta varios PNG en un .ico (entradas PNG, soportado por Windows)."""
    n = len(images_png)
    header = struct.pack("<HHH", 0, 1, n)
    entries, offset = b"", 6 + 16 * n
    blobs = b""
    for size, data in images_png:
        w = h = 0 if size >= 256 else size
        entries += struct.pack("<BBBBHHII", w, h, 0, 0, 1, 32, len(data), offset)
        offset += len(data)
        blobs += data
    with open(path, "wb") as fh:
        fh.write(header + entries + blobs)


def main():
    QtWidgets.QApplication([])
    os.makedirs(OUT, exist_ok=True)
    sizes = [16, 24, 32, 48, 64, 128, 256]
    pngs = [(s, png_bytes(render_png(s))) for s in sizes]
    render_png(256).save(os.path.join(OUT, "icon.png"), "PNG")
    write_ico(os.path.join(OUT, "icon.ico"), pngs)
    print("Generado:", os.path.join(OUT, "icon.png"), "y", os.path.join(OUT, "icon.ico"))


if __name__ == "__main__":
    main()
