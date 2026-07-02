# -*- coding: utf-8 -*-
"""Taller 2 — Vectores. Resuelto en ESPAÑOL."""
import docbuild as D
from docbuild import (new_doc, cover, exercise_heading, statement, subhead,
                      para, bullet, add_eq, answer, answer_text, image,
                      r, i, sup, sub, subsup, frac, sqrt, paren, brack)

NAME = "EDWARD ANDRÉS RODRÍGUEZ MAZUERA"
CODE = "2535250-2713"

doc = new_doc()
cover(doc, "Vectores", "Capítulo 3 — Ejercicios (Clase 06)", NAME, CODE, lang='es')

para(doc, "Estrategia general de resolución: (1) pensar, (2) esquematizar, "
          "(3) investigar, (4) simplificar, (5) calcular, (6) redondear, "
          "(7) analizar. Se trabaja en el Sistema Internacional (SI) y se usa la "
          "descomposición en componentes rectangulares.",
     size=10, italic=True, space_after=8)

# ---------------------------------------------------------------- Ej 1
exercise_heading(doc, "Ejercicio 1")
statement(doc, "Una persona da un paseo siguiendo la ruta: 100 m al este, luego "
               "300 m al sur, se desvía 150 m a 30° al sur del oeste y, por último, "
               "200 m a 60° al norte del oeste. Hallar el desplazamiento resultante "
               "de la persona desde el punto de partida.")
subhead(doc, "Datos")
bullet(doc, "a = 100 m al este       b = 300 m al sur")
bullet(doc, "c = 150 m a 30° al sur del oeste   d = 200 m a 60° al norte del oeste")
image(doc, "img/w2e1.png", 9.0, "Figura. Trayectoria de los cuatro tramos y vector resultante R.")
subhead(doc, "Componentes rectangulares")
para(doc, "Tomando el eje +x hacia el este y el eje +y hacia el norte:")
add_eq(doc, sub(i('a'),i('x')), r('='), r('100 m'), r('   ,   '), sub(i('b'),i('y')), r('='), r('−300 m'))
add_eq(doc, sub(i('c'),i('x')), r('='), r('150'), r('cos',True)+paren(r('210°')), r('='), r('−129.90 m'),
            r('   ,   '), sub(i('c'),i('y')), r('='), r('150'), r('sen',True)+paren(r('210°')), r('='), r('−75 m'))
add_eq(doc, sub(i('d'),i('x')), r('='), r('200'), r('cos',True)+paren(r('120°')), r('='), r('−100 m'),
            r('   ,   '), sub(i('d'),i('y')), r('='), r('200'), r('sen',True)+paren(r('120°')), r('='), r('173.21 m'))
subhead(doc, "Vector resultante")
add_eq(doc, sub(i('R'),i('x')), r('='), sub(i('a'),i('x')), r('+'), sub(i('c'),i('x')), r('+'), sub(i('d'),i('x')),
            r('='), r('−129.90 m'))
add_eq(doc, sub(i('R'),i('y')), r('='), sub(i('b'),i('y')), r('+'), sub(i('c'),i('y')), r('+'), sub(i('d'),i('y')),
            r('='), r('−201.79 m'))
add_eq(doc, i('R'), r('='), sqrt(subsup(i('R'),i('x'),r('2'))+r('+')+subsup(i('R'),i('y'),r('2'))),
            r('='), sqrt(sup(r('(−129.90)'),r('2'))+r('+')+sup(r('(−201.79)'),r('2'))), r('='), r('240 m'))
add_eq(doc, i('θ'), r('='), sup(r('tan',True),r('−1'))+brack(frac(r('|')+sub(i('R'),i('y'))+r('|'), r('|')+sub(i('R'),i('x'))+r('|'))),
            r('='), r('57.2°'))
answer(doc, i('R'), r('≈'), r('240 m'), r('  ,  '), i('θ'), r('='), r('57.2° al sur del oeste'), label='Respuesta:')

# ---------------------------------------------------------------- Ej 2
exercise_heading(doc, "Ejercicio 2")
statement(doc, "Dados los vectores de desplazamiento A = (3î − 4ĵ + 4k̂) m y "
               "B = (2î + 3ĵ − 7k̂) m, hallar las magnitudes de: (a) C = A + B, "
               "(b) D = 2A − B. Exprese cada resultado en términos de sus "
               "componentes rectangulares.")
subhead(doc, "Datos")
bullet(doc, "A = (3î − 4ĵ + 4k̂) m       B = (2î + 3ĵ − 7k̂) m")
subhead(doc, "(a)  C = A + B")
add_eq(doc, i('C'), r('='), paren(r('3î − 4ĵ + 4k̂')), r('+'), paren(r('2î + 3ĵ − 7k̂')),
            r('='), r('5î − ĵ − 3k̂  (m)'))
add_eq(doc, i('C'), r('='), sqrt(sup(r('5'),r('2'))+r('+')+sup(r('(−1)'),r('2'))+r('+')+sup(r('(−3)'),r('2'))),
            r('='), sqrt(r('35')), r('≈'), r('5.92 m'))
answer_text(doc, "C = (5î − ĵ − 3k̂) m,   |C| ≈ 6 m.", label='Respuesta (a):')
subhead(doc, "(b)  D = 2A − B")
add_eq(doc, i('D'), r('='), r('2'), paren(r('3î − 4ĵ + 4k̂')), r('−'), paren(r('2î + 3ĵ − 7k̂')))
add_eq(doc, i('D'), r('='), paren(r('6î − 8ĵ + 8k̂')), r('−'), paren(r('2î + 3ĵ − 7k̂')),
            r('='), r('4î − 11ĵ + 15k̂  (m)'))
add_eq(doc, i('D'), r('='), sqrt(sup(r('4'),r('2'))+r('+')+sup(r('(−11)'),r('2'))+r('+')+sup(r('15'),r('2'))),
            r('='), sqrt(r('362')), r('≈'), r('19.0 m'))
answer_text(doc, "D = (4î − 11ĵ + 15k̂) m,   |D| ≈ 19.0 m.", label='Respuesta (b):')

# ---------------------------------------------------------------- Ej 3
exercise_heading(doc, "Ejercicio 3")
statement(doc, "Una marinera en un velero pequeño se topa con vientos cambiantes. "
               "Navega 2.00 km al este, luego 3.50 km a 45.0° al sur del este y "
               "después otro tramo en una dirección desconocida. Su posición final "
               "es 5.80 km directamente al este del punto de partida. Determine la "
               "magnitud y dirección del tercer tramo.")
subhead(doc, "Planteamiento por componentes")
para(doc, "Tramo 1: (2.00, 0) km.  Tramo 2 (45° al sur del este): "
          "(3.50 cos(−45°), 3.50 sen(−45°)) = (2.475, −2.475) km.")
add_eq(doc, sub(i('S'),i('x')), r('='), r('2.00 + 2.475'), r('='), r('4.475 km'), r('   ,   '),
            sub(i('S'),i('y')), r('='), r('0 − 2.475'), r('='), r('−2.475 km'))
para(doc, "La posición final debe ser (5.80, 0). El tercer tramo es la diferencia:")
add_eq(doc, sub(i('T'),i('x')), r('='), r('5.80 − 4.475'), r('='), r('1.325 km'), r('   ,   '),
            sub(i('T'),i('y')), r('='), r('0 −'), paren(r('−2.475')), r('='), r('2.475 km'))
add_eq(doc, i('T'), r('='), sqrt(sup(r('1.325'),r('2'))+r('+')+sup(r('2.475'),r('2'))), r('='), r('2.81 km'))
add_eq(doc, i('θ'), r('='), sup(r('tan',True),r('−1'))+brack(frac(r('2.475'),r('1.325'))), r('='), r('61.8°'))
answer(doc, i('T'), r('≈'), r('2.81 km'), r('  ,  '), i('θ'), r('='), r('61.8° al norte del este'), label='Respuesta:')

# ---------------------------------------------------------------- Ej 4
exercise_heading(doc, "Ejercicio 4")
statement(doc, "Una araña descansa en su red. La fuerza gravitacional sobre la araña "
               "es 0.150 N hacia abajo. La araña está sostenida por fuerzas de tensión "
               "por encima de ella, de modo que la fuerza resultante es cero. Las dos "
               "tensiones son perpendiculares entre sí y la tensión Tx = 0.127 N. "
               "Calcule (a) la tensión Ty y (b) el ángulo que forma el peso con el eje y.")
subhead(doc, "Solución")
para(doc, "Como la resultante es cero, la suma de las dos tensiones perpendiculares "
          "equilibra el peso. Por ser perpendiculares:")
add_eq(doc, i('W'), r('='), sqrt(subsup(i('T'),i('x'),r('2'))+r('+')+subsup(i('T'),i('y'),r('2'))))
add_eq(doc, sub(i('T'),i('y')), r('='), sqrt(sup(i('W'),r('2'))+r('−')+subsup(i('T'),i('x'),r('2'))),
            r('='), sqrt(sup(r('0.150'),r('2'))+r('−')+sup(r('0.127'),r('2'))), r('='), r('0.0798 N'))
answer_text(doc, "Ty = 0.0798 N.", label='Respuesta (a):')
para(doc, "(b) El ángulo que el peso (vertical) forma con el eje y se obtiene de las componentes:")
add_eq(doc, i('θ'), r('='), sup(r('tan',True),r('−1'))+brack(frac(sub(i('T'),i('x')), sub(i('T'),i('y')))),
            r('='), sup(r('tan',True),r('−1'))+brack(frac(r('0.127'),r('0.0798'))), r('='), r('57.9°'))
answer(doc, i('θ'), r('≈'), r('57.9°'), label='Respuesta (b):')

# ---------------------------------------------------------------- Ej 5
exercise_heading(doc, "Ejercicio 5")
statement(doc, "Un avión de pasajeros viaja inicialmente a 300 km/h hacia el este y, "
               "de repente, entra en una región donde el viento sopla a 100 km/h en "
               "una dirección 30.0° al norte del este. ¿Cuáles son la nueva velocidad "
               "y la dirección de la aeronave respecto al suelo?")
subhead(doc, "Suma de velocidades (avión + viento)")
add_eq(doc, sub(i('V'),i('x')), r('='), r('300 + 100'), r('cos',True)+paren(r('30°')), r('='), r('386.6 km/h'))
add_eq(doc, sub(i('V'),i('y')), r('='), r('0 + 100'), r('sen',True)+paren(r('30°')), r('='), r('50 km/h'))
add_eq(doc, i('V'), r('='), sqrt(sup(r('386.6'),r('2'))+r('+')+sup(r('50'),r('2'))), r('='), r('390 km/h'))
add_eq(doc, i('θ'), r('='), sup(r('tan',True),r('−1'))+brack(frac(r('50'),r('386.6'))), r('='), r('7.37°'))
answer(doc, i('V'), r('≈'), r('390 km/h'), r('  ,  '), i('θ'), r('='), r('7.37° al norte del este'), label='Respuesta:')

# ---------------------------------------------------------------- Ej 6
exercise_heading(doc, "Ejercicio 6")
statement(doc, "Un rinoceronte está en el origen de un sistema de coordenadas en "
               "t₁ = 0. Para el intervalo entre t₁ y t₂ = 12.0 s, la velocidad media "
               "del animal tiene componente x de 3.8 m/s y componente y de 4.9 m/s. "
               "En t₂: (a) ¿qué coordenadas x e y tiene el rinoceronte? "
               "(b) ¿qué tan lejos está del origen?")
subhead(doc, "Solución")
para(doc, "Como parte del origen, la posición en t₂ es la velocidad media por el tiempo:")
add_eq(doc, i('x'), r('='), sub(i('v'),i('x')), i('t'), r('='), r('(3.8)(12.0)'), r('='), r('45.6 m'))
add_eq(doc, i('y'), r('='), sub(i('v'),i('y')), i('t'), r('='), r('(4.9)(12.0)'), r('='), r('58.8 m'))
answer_text(doc, "x = 45.6 m,   y = 58.8 m.", label='Respuesta (a):')
add_eq(doc, i('r'), r('='), sqrt(sup(i('x'),r('2'))+r('+')+sup(i('y'),r('2'))),
            r('='), sqrt(sup(r('45.6'),r('2'))+r('+')+sup(r('58.8'),r('2'))), r('='), r('74.4 m'))
answer(doc, i('r'), r('≈'), r('74.4 m'), label='Respuesta (b):')

# ---------------------------------------------------------------- Ej 7
exercise_heading(doc, "Ejercicio 7")
statement(doc, "Un avión despega del aeropuerto A y vuela 300 km al este, luego "
               "350 km a 30.0° al oeste del norte y después 150 km al norte para "
               "llegar finalmente al aeropuerto B. Al día siguiente, otro avión vuela "
               "en línea recta desde A hasta B. (a) ¿En qué dirección debe viajar el "
               "piloto en este vuelo directo? (b) ¿Qué distancia recorrerá? "
               "Suponga que no hay viento.")
subhead(doc, "Suma vectorial de los tres tramos")
add_eq(doc, sub(i('R'),i('x')), r('='), r('300 − 350'), r('sen',True)+paren(r('30°')), r('+ 0'),
            r('='), r('125 km'))
add_eq(doc, sub(i('R'),i('y')), r('='), r('0 + 350'), r('cos',True)+paren(r('30°')), r('+ 150'),
            r('='), r('453.1 km'))
add_eq(doc, i('R'), r('='), sqrt(sup(r('125'),r('2'))+r('+')+sup(r('453.1'),r('2'))), r('='), r('470 km'))
add_eq(doc, i('θ'), r('='), sup(r('tan',True),r('−1'))+brack(frac(r('453.1'),r('125'))), r('='), r('74.6°'))
answer(doc, i('θ'), r('='), r('74.6° al norte del este'), r('  ,  '), i('R'), r('≈'), r('470 km'), label='Respuesta:')

out = "/home/user/AgenteSAE/talleres/Taller_Vectores_Espanol.docx"
doc.save(out)
print("Saved", out, "-- paragraphs:", len(doc.paragraphs))
