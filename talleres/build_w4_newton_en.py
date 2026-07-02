# -*- coding: utf-8 -*-
"""Workshop 4 — Newton's Laws of Motion. Exercises 5–8 only. ENGLISH, step by step."""
import docbuild as D
from docbuild import (new_doc, cover, exercise_heading, statement, subhead,
                      para, bullet, add_eq, answer, answer_text, image, step,
                      r, i, sup, sub, subsup, frac, sqrt, paren, brack)

NAME = "EDWARD ANDRÉS RODRÍGUEZ MAZUERA"
CODE = "2535250-2713"

doc = new_doc()
cover(doc, "Newton's Laws of Motion",
      "Chapter 5 — Problem Set (Class 10) · Exercises 5–8", NAME, CODE, lang='en')

para(doc, "General problem-solving strategy: (1) think, (2) sketch, (3) research, "
          "(4) simplify, (5) calculate, (6) round, (7) analyse. Each step is briefly "
          "explained. Free-body diagrams and Newton's second law (ΣF = ma) are "
          "applied to each body; g = 9.8 m/s².",
     size=10, italic=True, space_after=8)

# ================================================================ Ex 5
exercise_heading(doc, "Exercise 5")
statement(doc, "Mass m₁ on a frictionless horizontal table is connected to mass m₂ "
               "by means of a very light movable pulley P₁ and a light fixed pulley "
               "P₂, as shown in the figure. (a) If a₁ and a₂ are the accelerations of "
               "m₁ and m₂ respectively, what is the relationship between these "
               "accelerations? Express (b) the tensions in the strings and (c) the "
               "accelerations a₁ and a₂ in terms of the masses m₁, m₂ and g.")
image(doc, "img/w4e5.png", 9.5, "Figure. m₁ pulled by a single string; m₂ is carried by the movable pulley (two supporting strands).")
subhead(doc, "Solution")
step(doc, 1, "(a) Two strands support the movable pulley, so m₂ moves half as far as m₁; differentiate twice.")
add_eq(doc, sub(i('a'),r('1')), r('='), r('2'), sub(i('a'),r('2')))
answer(doc, sub(i('a'),r('1')), r('='), r('2'), sub(i('a'),r('2')), label='Answer (a):')
step(doc, 2, "Apply ΣF = ma to m₁ on the frictionless table (single tension T₁).")
add_eq(doc, sub(i('T'),r('1')), r('='), sub(i('m'),r('1')), sub(i('a'),r('1')))
step(doc, 3, "Apply ΣF = ma to m₂; two strands pull up, so the net upward force is 2T₁.")
add_eq(doc, sub(i('m'),r('2')), i('g'), r('−'), r('2'), sub(i('T'),r('1')), r('='),
            sub(i('m'),r('2')), sub(i('a'),r('2')))
step(doc, 4, "(c) Substitute T₁ = m₁a₁ = 2m₁a₂ into the m₂ equation and solve for a₂; then a₁ = 2a₂.")
add_eq(doc, sub(i('m'),r('2')), i('g'), r('−'), r('4'), sub(i('m'),r('1')), sub(i('a'),r('2')),
            r('='), sub(i('m'),r('2')), sub(i('a'),r('2')),
            r('  ⇒  '), sub(i('a'),r('2')), r('='), frac(sub(i('m'),r('2'))+i('g'), r('4')+sub(i('m'),r('1'))+r('+')+sub(i('m'),r('2'))))
answer(doc, sub(i('a'),r('1')), r('='), frac(r('2')+sub(i('m'),r('2'))+i('g'), r('4')+sub(i('m'),r('1'))+r('+')+sub(i('m'),r('2'))),
            r('  ,   '), sub(i('a'),r('2')), r('='), frac(sub(i('m'),r('2'))+i('g'), r('4')+sub(i('m'),r('1'))+r('+')+sub(i('m'),r('2'))),
            label='Answer (c):')
step(doc, 5, "(b) Back-substitute a₁ to get the string tension T₁; the movable pulley doubles it to T₂ = 2T₁.")
answer(doc, sub(i('T'),r('1')), r('='), frac(r('2')+sub(i('m'),r('1'))+sub(i('m'),r('2'))+i('g'), r('4')+sub(i('m'),r('1'))+r('+')+sub(i('m'),r('2'))),
            r('  ,   '), sub(i('T'),r('2')), r('='), r('2'), sub(i('T'),r('1')), r('='),
            frac(r('4')+sub(i('m'),r('1'))+sub(i('m'),r('2'))+i('g'), r('4')+sub(i('m'),r('1'))+r('+')+sub(i('m'),r('2'))),
            label='Answer (b):')

# ================================================================ Ex 6
exercise_heading(doc, "Exercise 6")
statement(doc, "Block A in the figure weighs 1.90 N, and block B weighs 4.20 N. The "
               "coefficient of kinetic friction between all surfaces is 0.30. Find the "
               "magnitude of the horizontal force F necessary to drag block B to the "
               "left at constant speed if A and B are connected by a light, flexible "
               "cord passing around a fixed, frictionless pulley.")
subhead(doc, "Data")
bullet(doc, "w_A = 1.90 N,  w_B = 4.20 N,  μ_k = 0.30,  constant speed ⇒ ΣF = 0")
image(doc, "img/w4e6.png", 9.0, "Figure. A rests on B; the cord links A to B over a wall pulley. Dragging B left makes A slide right.")
subhead(doc, "Solution")
step(doc, 1, "The cord over the pulley makes A slide right when B moves left; identify the two sliding contacts (A–B and B–floor).")
step(doc, 2, "Block A moves at constant speed, so the cord tension equals the friction from B (N_A = w_A).")
add_eq(doc, i('T'), r('='), sub(i('μ'),i('k')), sub(i('w'),i('A')), r('='), r('0.30'), r('×'),
            r('1.90'), r('='), r('0.57 N'))
step(doc, 3, "Block B moves at constant speed: F balances the tension, the friction from A on top, and the floor friction.")
add_eq(doc, i('F'), r('='), i('T'), r('+'), sub(i('μ'),i('k')), sub(i('w'),i('A')), r('+'),
            sub(i('μ'),i('k')), paren(sub(i('w'),i('A'))+r('+')+sub(i('w'),i('B'))))
step(doc, 4, "Substitute the numbers and add.")
add_eq(doc, i('F'), r('='), r('0.57'), r('+'), r('0.57'), r('+'), r('0.30'), paren(r('1.90+4.20')),
            r('='), r('0.57 + 0.57 + 1.83'), r('='), r('2.97 N'))
answer(doc, i('F'), r('≈'), r('2.97 N'), label='Answer:')
para(doc, "Note. The listed answer key of 2.52 N corresponds to the original textbook "
          "value w_A = 1.40 N, for which F = μ_k(3w_A + w_B) = 0.30(4.20 + 4.20) = 2.52 N. "
          "The value above (2.97 N) is correct for the weight printed here (w_A = 1.90 N).",
     size=9.5, italic=True)

# ================================================================ Ex 7
exercise_heading(doc, "Exercise 7")
statement(doc, "In terms of m₁, m₂ and g, find the acceleration of each block in the "
               "figure. There is no friction anywhere in the system.")
image(doc, "img/w4e7.png", 8.5, "Figure. m₁ on a frictionless table connected through a fixed pulley to a movable pulley carrying m₂.")
subhead(doc, "Solution")
step(doc, 1, "The layout is the frictionless version of Exercise 5: a single string, tension T, with two strands on the movable pulley (m₂).")
step(doc, 2, "Constraint: m₁ moves twice as fast as m₂ because m₂ hangs from the movable pulley.")
add_eq(doc, sub(i('a'),r('1')), r('='), r('2'), sub(i('a'),r('2')))
step(doc, 3, "Newton's second law for m₁ (frictionless table) and for m₂ (net upward pull 2T).")
add_eq(doc, i('T'), r('='), sub(i('m'),r('1')), sub(i('a'),r('1')), r('   ,   '),
            sub(i('m'),r('2')), i('g'), r('−'), r('2'), i('T'), r('='), sub(i('m'),r('2')), sub(i('a'),r('2')))
step(doc, 4, "Substitute T = 2m₁a₂ and solve for a₂; then a₁ = 2a₂.")
add_eq(doc, sub(i('m'),r('2')), i('g'), r('−'), r('4'), sub(i('m'),r('1')), sub(i('a'),r('2')),
            r('='), sub(i('m'),r('2')), sub(i('a'),r('2')),
            r('  ⇒  '), sub(i('a'),r('2')), r('='), frac(sub(i('m'),r('2'))+i('g'), r('4')+sub(i('m'),r('1'))+r('+')+sub(i('m'),r('2'))))
answer(doc, sub(i('a'),r('2')), r('='), frac(sub(i('m'),r('2'))+i('g'), r('4')+sub(i('m'),r('1'))+r('+')+sub(i('m'),r('2'))),
            r('  ,   '), sub(i('a'),r('1')), r('='), r('2'), sub(i('a'),r('2')), r('='),
            frac(r('2')+sub(i('m'),r('2'))+i('g'), r('4')+sub(i('m'),r('1'))+r('+')+sub(i('m'),r('2'))),
            label='Answer:')
para(doc, "Note. The movable pulley makes a₁ = 2a₂ (the answer key's “a₁ = a₂” drops "
          "the factor of 2; its value of a₂ matches the 2 : 1 pulley shown).",
     size=9.5, italic=True)

# ================================================================ Ex 8
exercise_heading(doc, "Exercise 8")
statement(doc, "A 20-kg monkey has a firm hold on a light rope that passes over a "
               "frictionless pulley and is attached to a 20-kg bunch of bananas. The "
               "monkey looks up, sees the bananas, and starts to climb the rope to "
               "get them. (a) As the monkey climbs, do the bananas move up, down, or "
               "stay at rest? (b) As the monkey climbs, does the distance between the "
               "monkey and the bananas decrease, increase, or stay the same? (c) The "
               "monkey releases her hold on the rope. What happens to the distance "
               "between the monkey and the bananas while she is falling? (d) Before "
               "reaching the ground, the monkey grabs the rope to stop her fall. What "
               "do the bananas do?")
image(doc, "img/w4e8.png", 5.0, "Figure. Equal masses (20 kg each) on a light rope over a frictionless pulley.")
subhead(doc, "Solution")
step(doc, 1, "Write ΣF = ma for the monkey and for the bananas with the same tension T (equal masses m).")
add_eq(doc, i('T'), r('−'), i('m'), i('g'), r('='), i('m'), sub(i('a'),r('mon')),
            r('   ,   '), i('T'), r('−'), i('m'), i('g'), r('='), i('m'), sub(i('a'),r('ban')))
step(doc, 2, "The two equations are identical, so both bodies always share the same acceleration and velocity.")
step(doc, 3, "(a) Climbing raises T above mg, so both accelerate upward equally — the bananas rise.")
answer_text(doc, "The bananas move up (at the same speed as the monkey).", label='Answer (a):')
step(doc, 4, "(b) Equal ground velocities mean the vertical gap does not change.")
answer_text(doc, "The distance stays the same (remains constant).", label='Answer (b):')
step(doc, 5, "(c) Releasing the rope sets T = 0, so both are in free fall at g and keep the same gap.")
answer_text(doc, "The distance stays the same (both fall freely at g).", label='Answer (c):')
step(doc, 6, "(d) Grabbing the rope restores T; equal masses decelerate identically, so both come to rest together.")
answer_text(doc, "The bananas also stop (they decelerate identically and come to rest).", label='Answer (d):')

out = "/home/user/AgenteSAE/talleres/Taller_Newtons_Laws_English.docx"
doc.save(out)
print("Saved", out, "-- paragraphs:", len(doc.paragraphs))
