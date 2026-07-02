# -*- coding: utf-8 -*-
"""Workshop 4 — Newton's Laws of Motion. Exercises 5–8 only. Solved in ENGLISH."""
import docbuild as D
from docbuild import (new_doc, cover, exercise_heading, statement, subhead,
                      para, bullet, add_eq, answer, answer_text, image,
                      r, i, sup, sub, subsup, frac, sqrt, paren, brack)

NAME = "EDWARD ANDRÉS RODRÍGUEZ MAZUERA"
CODE = "2535250-2713"

doc = new_doc()
cover(doc, "Newton's Laws of Motion",
      "Chapter 5 — Problem Set (Class 10) · Exercises 5–8", NAME, CODE, lang='en')

para(doc, "General problem-solving strategy: (1) think, (2) sketch, (3) research, "
          "(4) simplify, (5) calculate, (6) round, (7) analyse. Free-body diagrams "
          "and Newton's second law (ΣF = ma) are applied to each body; g = 9.8 m/s².",
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
subhead(doc, "(a)  Constraint between the accelerations")
para(doc, "A single string of tension T₁ runs from m₁, over the fixed pulley P₂, and "
          "around the movable pulley P₁ that carries m₂. Because two strands support "
          "the movable pulley, when m₁ advances a distance x the pulley (and m₂) moves "
          "only x/2. Differentiating twice:")
add_eq(doc, sub(i('a'),r('1')), r('='), r('2'), sub(i('a'),r('2')))
answer(doc, sub(i('a'),r('1')), r('='), r('2'), sub(i('a'),r('2')), label='Answer (a):')
subhead(doc, "Newton's second law")
para(doc, "Block m₁ (frictionless table, single tension T₁):")
add_eq(doc, sub(i('T'),r('1')), r('='), sub(i('m'),r('1')), sub(i('a'),r('1')))
para(doc, "Block m₂ hangs from the movable pulley, so two strands (each T₁) pull it "
          "up; the net upward pull is T₂ = 2T₁:")
add_eq(doc, sub(i('m'),r('2')), i('g'), r('−'), r('2'), sub(i('T'),r('1')), r('='),
            sub(i('m'),r('2')), sub(i('a'),r('2')))
subhead(doc, "(c)  Solving for the accelerations")
para(doc, "Substituting T₁ = m₁a₁ = 2m₁a₂ into the m₂ equation:")
add_eq(doc, sub(i('m'),r('2')), i('g'), r('−'), r('4'), sub(i('m'),r('1')), sub(i('a'),r('2')),
            r('='), sub(i('m'),r('2')), sub(i('a'),r('2')),
            r('  ⇒  '), sub(i('m'),r('2')), i('g'), r('='), paren(r('4')+sub(i('m'),r('1'))+r('+')+sub(i('m'),r('2'))), sub(i('a'),r('2')))
add_eq(doc, sub(i('a'),r('2')), r('='), frac(sub(i('m'),r('2'))+i('g'), r('4')+sub(i('m'),r('1'))+r('+')+sub(i('m'),r('2'))),
            r('  ,   '), sub(i('a'),r('1')), r('='), r('2'), sub(i('a'),r('2')), r('='),
            frac(r('2')+sub(i('m'),r('2'))+i('g'), r('4')+sub(i('m'),r('1'))+r('+')+sub(i('m'),r('2'))))
answer(doc, sub(i('a'),r('1')), r('='), frac(r('2')+sub(i('m'),r('2'))+i('g'), r('4')+sub(i('m'),r('1'))+r('+')+sub(i('m'),r('2'))),
            r('  ,   '), sub(i('a'),r('2')), r('='), frac(sub(i('m'),r('2'))+i('g'), r('4')+sub(i('m'),r('1'))+r('+')+sub(i('m'),r('2'))),
            label='Answer (c):')
subhead(doc, "(b)  Tensions")
add_eq(doc, sub(i('T'),r('1')), r('='), sub(i('m'),r('1')), sub(i('a'),r('1')), r('='),
            frac(r('2')+sub(i('m'),r('1'))+sub(i('m'),r('2'))+i('g'), r('4')+sub(i('m'),r('1'))+r('+')+sub(i('m'),r('2'))),
            r('  ,   '), sub(i('T'),r('2')), r('='), r('2'), sub(i('T'),r('1')))
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
subhead(doc, "Free-body analysis")
para(doc, "Because the cord runs over the fixed pulley, when B moves left, A moves "
          "right. There are two sliding contacts (A–B and B–floor), so three friction "
          "effects plus the cord tension resist the motion.")
para(doc, "Block A (constant velocity): the cord tension balances the friction from B, "
          "with normal force N_A = w_A:")
add_eq(doc, i('T'), r('='), sub(i('μ'),i('k')), sub(i('w'),i('A')), r('='), r('0.30'), r('×'),
            r('1.90'), r('='), r('0.57 N'))
para(doc, "Block B (constant velocity): F must overcome the cord tension T, the "
          "friction from A on top (μ_k·w_A), and the floor friction (μ_k·(w_A + w_B)):")
add_eq(doc, i('F'), r('='), i('T'), r('+'), sub(i('μ'),i('k')), sub(i('w'),i('A')), r('+'),
            sub(i('μ'),i('k')), paren(sub(i('w'),i('A'))+r('+')+sub(i('w'),i('B'))))
add_eq(doc, i('F'), r('='), r('0.57'), r('+'), r('0.57'), r('+'), r('0.30'), paren(r('1.90+4.20')),
            r('='), r('0.57 + 0.57 + 1.83'), r('='), r('2.97 N'))
answer(doc, i('F'), r('≈'), r('2.97 N'), label='Answer:')
para(doc, "Note. The listed answer key of 2.52 N corresponds to the original textbook "
          "value w_A = 1.40 N; with that weight F = μ_k(3w_A + w_B) = 0.30(4.20 + 4.20) "
          "= 2.52 N. The value above (2.97 N) is the correct result for the weights "
          "printed in this exercise (w_A = 1.90 N).", size=9.5, italic=True)

# ================================================================ Ex 7
exercise_heading(doc, "Exercise 7")
statement(doc, "In terms of m₁, m₂ and g, find the acceleration of each block in the "
               "figure. There is no friction anywhere in the system.")
image(doc, "img/w4e7.png", 8.5, "Figure. m₁ on a frictionless table connected through a fixed pulley to a movable pulley carrying m₂.")
subhead(doc, "Setup")
para(doc, "The arrangement is mechanically identical to Exercise 5, now completely "
          "frictionless. A single string (tension T) is attached to m₁, passes over "
          "the fixed pulley at the table edge, wraps around the movable pulley that "
          "holds m₂, and ends at a fixed hook. Two strands support the movable pulley.")
subhead(doc, "(a)  Constraint and equations")
add_eq(doc, sub(i('a'),r('1')), r('='), r('2'), sub(i('a'),r('2')))
para(doc, "Block m₁ (frictionless table):")
add_eq(doc, i('T'), r('='), sub(i('m'),r('1')), sub(i('a'),r('1')))
para(doc, "Block m₂ (carried by the movable pulley, net upward pull 2T):")
add_eq(doc, sub(i('m'),r('2')), i('g'), r('−'), r('2'), i('T'), r('='), sub(i('m'),r('2')), sub(i('a'),r('2')))
subhead(doc, "Solving")
para(doc, "With T = m₁a₁ = 2m₁a₂:")
add_eq(doc, sub(i('m'),r('2')), i('g'), r('−'), r('4'), sub(i('m'),r('1')), sub(i('a'),r('2')),
            r('='), sub(i('m'),r('2')), sub(i('a'),r('2')),
            r('  ⇒  '), sub(i('a'),r('2')), r('='), frac(sub(i('m'),r('2'))+i('g'), r('4')+sub(i('m'),r('1'))+r('+')+sub(i('m'),r('2'))))
answer(doc, sub(i('a'),r('2')), r('='), frac(sub(i('m'),r('2'))+i('g'), r('4')+sub(i('m'),r('1'))+r('+')+sub(i('m'),r('2'))),
            r('  ,   '), sub(i('a'),r('1')), r('='), r('2'), sub(i('a'),r('2')), r('='),
            frac(r('2')+sub(i('m'),r('2'))+i('g'), r('4')+sub(i('m'),r('1'))+r('+')+sub(i('m'),r('2'))),
            label='Answer:')
para(doc, "Note. The movable pulley makes m₁ move twice as fast as m₂, hence a₁ = 2a₂ "
          "(the answer key's “a₁ = a₂” drops this factor of 2; its value of a₂ is "
          "consistent with the 2 : 1 pulley shown).", size=9.5, italic=True)

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
subhead(doc, "Key idea")
para(doc, "The monkey (mass m) and the bananas (mass m) are equal. Applying Newton's "
          "second law to each body with the same rope tension T (upward positive):")
add_eq(doc, i('T'), r('−'), i('m'), i('g'), r('='), i('m'), sub(i('a'),r('monkey')),
            r('   ,   '), i('T'), r('−'), i('m'), i('g'), r('='), i('m'), sub(i('a'),r('bananas')))
para(doc, "The two equations are identical, so the monkey and the bananas always have "
          "the same acceleration (and, starting from rest, the same velocity) measured "
          "from the ground.")
subhead(doc, "(a)  Do the bananas move up, down, or stay at rest?")
para(doc, "When the monkey climbs, she increases the tension, so T > mg and both bodies "
          "accelerate upward equally. The bananas move UP, at the same rate as the monkey.")
answer_text(doc, "The bananas move up (at the same speed as the monkey).", label='Answer (a):')
subhead(doc, "(b)  Distance between monkey and bananas")
para(doc, "Since both rise with identical velocity relative to the ground, the vertical "
          "gap between them does not change.")
answer_text(doc, "The distance stays the same (remains constant).", label='Answer (b):')
subhead(doc, "(c)  After the monkey releases the rope")
para(doc, "With no one holding one end, the rope tension drops to zero, so the monkey "
          "and the bananas are both in free fall with acceleration g. Their velocities "
          "change identically, so the gap between them is unchanged.")
answer_text(doc, "The distance stays the same (both fall freely at g).", label='Answer (c):')
subhead(doc, "(d)  The monkey grabs the rope to stop her fall")
para(doc, "Grabbing the rope restores the tension. Equal masses again force equal "
          "accelerations, so the same braking that stops the monkey also acts on the "
          "bananas. Having equal downward speeds, both decelerate together and come "
          "to rest.")
answer_text(doc, "The bananas also stop (they decelerate identically and come to rest).", label='Answer (d):')

out = "/home/user/AgenteSAE/talleres/Taller_Newtons_Laws_English.docx"
doc.save(out)
print("Saved", out, "-- paragraphs:", len(doc.paragraphs))
