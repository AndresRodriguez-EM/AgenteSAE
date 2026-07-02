# -*- coding: utf-8 -*-
"""Workshop 3 — Kinematics in two dimensions (2D). Solved in ENGLISH, step by step."""
import docbuild as D
from docbuild import (new_doc, cover, exercise_heading, statement, subhead,
                      para, bullet, add_eq, answer, answer_text, image, step,
                      r, i, sup, sub, subsup, frac, sqrt, paren, brack)

NAME = "EDWARD ANDRÉS RODRÍGUEZ MAZUERA"
CODE = "2535250-2713"

doc = new_doc()
cover(doc, "Kinematics in Two Dimensions (2D)",
      "Chapter 4 — Problem Set (Class 08)", NAME, CODE, lang='en')

para(doc, "General problem-solving strategy: (1) think, (2) sketch, (3) research, "
          "(4) simplify, (5) calculate, (6) round, (7) analyse. Each step below is "
          "briefly explained. Unless stated otherwise, g = 9.8 m/s² and air "
          "resistance is neglected.",
     size=10, italic=True, space_after=8)

# ---------------------------------------------------------------- Ex 1
exercise_heading(doc, "Exercise 1")
statement(doc, "A projectile is launched so that the horizontal distance it "
               "travels (its range) is twice its maximum height. Find the launch angle.")
subhead(doc, "Data")
bullet(doc, "R = 2h,  y₀ = 0,  x₀ = 0")
subhead(doc, "Solution")
step(doc, 1, "Write the standard range and maximum-height formulas for a ground-level launch.")
add_eq(doc, i('R'), r('='), frac(subsup(i('v'),r('0'),r('2'))+r('sin',True)+paren(r('2')+i('θ')), i('g')),
            r('  ,   '), i('h'), r('='), frac(subsup(i('v'),r('0'),r('2'))+sup(r('sin',True)+i('θ'),r('2')), r('2')+i('g')))
step(doc, 2, "Impose the condition R = 2h; v₀² and g cancel.")
add_eq(doc, frac(subsup(i('v'),r('0'),r('2'))+r('sin',True)+paren(r('2')+i('θ')), i('g')), r('='),
            r('2'), frac(subsup(i('v'),r('0'),r('2'))+sup(r('sin',True)+i('θ'),r('2')), r('2')+i('g')),
            r('  ⇒  '), r('sin',True)+paren(r('2')+i('θ')), r('='), sup(r('sin',True)+i('θ'),r('2')))
step(doc, 3, "Expand sin(2θ) = 2 sinθ cosθ and cancel one sinθ to get tanθ.")
add_eq(doc, r('2'), r('sin',True)+i('θ'), r('cos',True)+i('θ'), r('='), sup(r('sin',True)+i('θ'),r('2')),
            r('  ⇒  '), r('2'), r('cos',True)+i('θ'), r('='), r('sin',True)+i('θ'),
            r('  ⇒  '), r('tan',True)+i('θ'), r('='), r('2'))
step(doc, 4, "Take the inverse tangent to obtain the launch angle.")
answer(doc, i('θ'), r('='), sup(r('tan',True),r('−1'))+paren(r('2')), r('='), r('63.4°'), label='Answer:')

# ---------------------------------------------------------------- Ex 2
exercise_heading(doc, "Exercise 2")
statement(doc, "A bomber flies horizontally over flat ground with a speed of "
               "275 m/s relative to the ground, at an altitude of 3 000 m. "
               "(a) How far will the bomb travel horizontally between release and "
               "impact? (b) If the plane keeps its original course and speed, where "
               "is it when the bomb hits the ground? (c) At what angle from the "
               "vertical, at the point of release, should the bomb-sight be set so "
               "that the bomb hits the target seen in the sight at the instant of release?")
subhead(doc, "Data")
bullet(doc, "θ = 0°  (horizontal release),  v₀ = 275 m/s,  y₀ = 3 000 m")
subhead(doc, "Solution")
step(doc, 1, "(a) The vertical motion is free fall from rest; set y = 0 at impact and solve for the fall time.")
add_eq(doc, r('0'), r('='), sub(i('y'),r('0')), r('−'), frac(r('1'),r('2')), i('g'), sup(i('t'),r('2')),
            r('  ⇒  '), r('4.9'), subsup(i('t'),r('b'),r('2')), r('−'), r('3000'), r('='), r('0'))
add_eq(doc, sub(i('t'),r('b')), r('='), sqrt(frac(r('3000'), r('4.9'))), r('='), r('24.7 s'))
step(doc, 2, "Multiply the constant horizontal speed by the fall time to get the horizontal range.")
add_eq(doc, i('x'), r('='), sub(i('v'),r('0')), sub(i('t'),r('b')), r('='),
            r('275'), r('×'), r('24.7'), r('='), r('6 805 m'))
answer_text(doc, "x ≈ 6.80 × 10³ m  (about 6 805 m).", label='Answer (a):')
step(doc, 3, "(b) Bomb and plane share the same horizontal velocity, so the plane stays directly above the bomb.")
answer_text(doc, "The plane is directly over the bomb (6 805 m from release, 3 000 m up).", label='Answer (b):')
step(doc, 4, "(c) The sight line to the target makes angle φ with the vertical; take the arctangent of x/y₀.")
add_eq(doc, r('tan',True)+i('φ'), r('='), frac(i('x'), sub(i('y'),r('0'))), r('='),
            frac(r('6 805'), r('3 000')), r('  ⇒  '), i('φ'), r('='), r('66.2°'))
answer(doc, i('φ'), r('≈'), r('66°'), r(' from the vertical'), label='Answer (c):')

# ---------------------------------------------------------------- Ex 3
exercise_heading(doc, "Exercise 3")
statement(doc, "Jimmy is at the bottom of a hill, while Billy is 30 m up the same "
               "hill. Jimmy is at the origin of an xy-system and the line of the "
               "slope is given by y = 0.4x. If Jimmy throws an apple to Billy at an "
               "angle of 50° above the horizontal, with what speed must he throw the "
               "apple so that it reaches Billy?")
subhead(doc, "Data")
bullet(doc, "d = 30 m (distance along the slope),  θ = 50°,  slope:  y = 0.4x")
subhead(doc, "Solution")
step(doc, 1, "Write the trajectory equation y(x) of the projectile.")
add_eq(doc, i('y'), r('='), r('tan',True)+i('θ'), i('x'), r('−'),
            frac(i('g'), r('2')+subsup(i('v'),r('0'),r('2'))+sup(r('cos',True)+i('θ'),r('2'))), sup(i('x'),r('2')))
step(doc, 2, "Billy lies on y = 0.4x, so replace y/x by 0.4; his horizontal coordinate is x = d·cos(tan⁻¹0.4).")
add_eq(doc, r('0.4'), r('='), r('tan',True)+i('θ'), r('−'),
            frac(i('g'), r('2')+subsup(i('v'),r('0'),r('2'))+sup(r('cos',True)+i('θ'),r('2'))), i('x'))
step(doc, 3, "Solve the equation algebraically for the launch speed v₀.")
add_eq(doc, sub(i('v'),r('0')), r('='),
            sqrt(frac(i('g')+i('d')+r('cos',True)+brack(sup(r('tan',True),r('−1'))+brack(r('0.4'))),
                      r('2')+sup(r('cos',True)+i('θ'),r('2'))+paren(r('tan',True)+i('θ')+r('−0.4')))))
step(doc, 4, "Substitute x = 27.85 m, tan50° = 1.1918, cos²50° = 0.4132 and evaluate.")
add_eq(doc, sub(i('v'),r('0')), r('='), sqrt(frac(r('(9.8)(27.85)'), r('2(0.4132)(1.1918−0.4)'))),
            r('='), r('20.4 m/s'))
answer(doc, sub(i('v'),r('0')), r('≈'), r('20.4 m/s'), label='Answer:')

# ---------------------------------------------------------------- Ex 4
exercise_heading(doc, "Exercise 4")
statement(doc, "A child slides down a frozen roof. He starts from the top of the "
               "roof, which is 8.0 m long and makes an angle of 40° with the "
               "horizontal, accelerating at 5.0 m/s². The edge of the roof is 6.0 m "
               "above a soft snowbank, on which he will land. Find the horizontal "
               "distance between the house and the landing point.")
subhead(doc, "Data")
bullet(doc, "v₀ = 0,  L = 8.0 m,  a = 5.0 m/s²,  θ = 40°,  y₀ = 6.0 m")
subhead(doc, "Solution")
step(doc, 1, "Stage 1 (on the roof): use v² = v₀² + 2aL to find the speed at the edge.")
add_eq(doc, sup(i('v'),r('2')), r('='), subsup(i('v'),r('0'),r('2'))+r('+'), r('2'), i('a'), i('L'),
            r('='), r('2(5.0)(8.0)'), r('='), r('80 m²/s²'))
step(doc, 2, "This speed becomes the launch speed for stage 2, directed 40° below the horizontal.")
step(doc, 3, "Stage 2 (projectile): place the origin at the edge and set y = 0 at the snowbank.")
add_eq(doc, r('0'), r('='), sub(i('y'),r('0')), r('−'), r('tan',True)+i('θ'), i('x'), r('−'),
            frac(i('g'), r('2')+sup(i('v'),r('2'))+sup(r('cos',True)+i('θ'),r('2'))), sup(i('x'),r('2')))
step(doc, 4, "Solve the resulting quadratic in x and keep the positive root.")
add_eq(doc, r('0.1044'), sup(i('x'),r('2')), r('+'), r('0.8391'), i('x'), r('−'), r('6.0'), r('='), r('0'),
            r('  ⇒  '), i('x'), r('='), r('4.56 m'))
answer(doc, i('x'), r('≈'), r('4.6 m'), label='Answer:')

# ---------------------------------------------------------------- Ex 5
exercise_heading(doc, "Exercise 5")
statement(doc, "A skier leaves the ramp of a ski jump with a velocity of 10.0 m/s, "
               "15.0° above the horizontal. The slope below is inclined at 50.0°, and "
               "air resistance is negligible. Find (a) the distance from the ramp to "
               "the landing point, and (b) the velocity components just before landing. "
               "(How might the results change if air resistance were included? Note "
               "that jumpers lean forward into an airfoil shape to increase their "
               "distance — why does this work?)")
subhead(doc, "Data")
bullet(doc, "v₀ = 10.0 m/s,  θ = 15.0° (above horizontal),  slope angle = 50.0°")
image(doc, "img/w3e5.png", 10.5, "Figure. Parabolic trajectory meeting the 50° slope.")
subhead(doc, "Solution")
step(doc, 1, "Landing point lies on the slope, so impose y = −x·tan50° with x = v₀cosθ·t and y = v₀sinθ·t − ½gt².")
add_eq(doc, frac(i('v')+sub(r(''),r('0'))+r('sin',True)+i('θ')+i('t')+r('−')+frac(r('1'),r('2'))+i('g')+sup(i('t'),r('2')),
                 i('v')+sub(r(''),r('0'))+r('cos',True)+i('θ')+i('t')), r('='), r('−'), r('tan',True)+r('50°'))
step(doc, 2, "Solve for the flight time t.")
add_eq(doc, i('t'), r('='), frac(i('v')+sub(r(''),r('0'))+paren(r('sin',True)+i('θ')+r('+')+r('cos',True)+i('θ')+r('tan',True)+r('50°')),
                                 frac(r('1'),r('2'))+i('g')), r('='), r('2.88 s'))
step(doc, 3, "(a) Find the horizontal reach x, then divide by cos50° to get the distance d along the slope.")
add_eq(doc, i('d'), r('='), frac(i('x'), r('cos',True)+r('50°')), r('='), frac(r('27.8'), r('0.6428')),
            r('='), r('43.2 m'))
answer(doc, i('d'), r('≈'), r('43.2 m'), label='Answer (a):')
step(doc, 4, "(b) Evaluate the velocity components at t = 2.88 s (vₓ is constant; v_y = v₀sinθ − gt).")
add_eq(doc, sub(i('v'),i('x')), r('='), r('9.7 m/s'), r('   ,   '),
            sub(i('v'),i('y')), r('='), r('2.59 − (9.8)(2.88)'), r('='), r('−25.6 m/s'))
answer_text(doc, "v = (9.7 î − 25.6 ĵ) m/s.", label='Answer (b):')
para(doc, "Discussion: with air resistance the range would shorten; leaning into an "
          "airfoil shape adds lift and cuts drag, lengthening the flight and the distance.",
     size=9.5, italic=True)

# ---------------------------------------------------------------- Ex 6
exercise_heading(doc, "Exercise 6")
statement(doc, "A daring 510 N swimmer dives off a cliff with a running horizontal "
               "leap. What must her minimum speed be as she leaves the top so that she "
               "clears the ledge at the bottom, which is 1.75 m wide and 9.00 m below "
               "the top of the cliff?")
subhead(doc, "Data")
bullet(doc, "Horizontal launch (θ = 0),  ledge width x = 1.75 m,  drop y = 9.00 m")
subhead(doc, "Solution")
step(doc, 1, "The 510 N weight is irrelevant (motion is mass-independent); find the time to fall 9.00 m from rest.")
add_eq(doc, i('t'), r('='), sqrt(frac(r('2')+i('y'), i('g'))), r('='), sqrt(frac(r('2(9.00)'),r('9.8'))),
            r('='), r('1.355 s'))
step(doc, 2, "The minimum speed just clears the ledge: cover 1.75 m horizontally in that time.")
add_eq(doc, sub(i('v'),r('0')), r('='), frac(i('x'), i('t')), r('='),
            frac(r('1.75'), r('1.355')), r('='), r('1.29 m/s'))
answer(doc, sub(i('v'),r('0')), r('≈'), r('1.3 m/s'), label='Answer:')

# ---------------------------------------------------------------- Ex 7
exercise_heading(doc, "Exercise 7")
statement(doc, "A water hose fills a large cylindrical storage tank of diameter D and "
               "height 2D. The hose shoots water at 45° above the horizontal, from the "
               "same level as the base of the tank and a distance 6D away. For what "
               "range of launch speeds v₀ will the water enter the tank? Ignore air "
               "resistance and express the answer in terms of D and g.")
subhead(doc, "Data")
bullet(doc, "θ = 45°,  near wall at x = 6D,  far wall at x = 7D,  rim height y = 2D")
image(doc, "img/w3e7.png", 10.0, "Figure. Slowest jet just clears the near rim; fastest just reaches the far rim.")
subhead(doc, "Solution")
step(doc, 1, "At θ = 45° the trajectory simplifies (2cos²45° = 1).")
add_eq(doc, i('y'), r('='), i('x'), r('−'), frac(i('g'), subsup(i('v'),r('0'),r('2'))), sup(i('x'),r('2')))
step(doc, 2, "Slowest speed: the jet just clears the near rim (x = 6D, y = 2D). Solve for v₀².")
add_eq(doc, r('2'), i('D'), r('='), r('6'), i('D'), r('−'),
            frac(i('g')+r('(6')+i('D')+sup(r(')'),r('2')), subsup(i('v'),r('0'),r('2'))),
            r('  ⇒  '), i('v')+sub(r(''),r('0')), r('='), r('3.0'), sqrt(i('g')+i('D')))
step(doc, 3, "Fastest speed: the jet just reaches the far rim (x = 7D, y = 2D). Solve for v₀².")
add_eq(doc, r('2'), i('D'), r('='), r('7'), i('D'), r('−'),
            frac(i('g')+r('(7')+i('D')+sup(r(')'),r('2')), subsup(i('v'),r('0'),r('2'))),
            r('  ⇒  '), i('v')+sub(r(''),r('0')), r('='), r('3.1'), sqrt(i('g')+i('D')))
step(doc, 4, "The water enters for any speed between these two limits.")
answer(doc, r('3.0'), sqrt(i('g')+i('D')), r('<'), i('v')+sub(r(''),r('0')), r('<'),
            r('3.1'), sqrt(i('g')+i('D')), label='Answer:')

# ---------------------------------------------------------------- Ex 8
exercise_heading(doc, "Exercise 8")
statement(doc, "A flea can jump to a vertical height h. (a) What is the maximum "
               "horizontal distance it can jump? (b) What is the time in the air in "
               "both cases (the vertical jump and the maximum-range jump)?")
subhead(doc, "Solution")
step(doc, 1, "A vertical jump to height h fixes the take-off speed v₀.")
add_eq(doc, i('h'), r('='), frac(subsup(i('v'),r('0'),r('2')), r('2')+i('g')), r('  ⇒  '),
            subsup(i('v'),r('0'),r('2')), r('='), r('2'), i('g'), i('h'))
step(doc, 2, "(a) Maximum range occurs at 45°, where R = v₀²/g; substitute v₀².")
add_eq(doc, sub(i('R'),r('max')), r('='), frac(subsup(i('v'),r('0'),r('2')), i('g')), r('='),
            frac(r('2')+i('g')+i('h'), i('g')), r('='), r('2'), i('h'))
answer(doc, sub(i('R'),r('max')), r('='), r('2'), i('h'), label='Answer (a):')
step(doc, 3, "(b) Compute the time aloft for each launch angle using t = 2v₀sinθ/g.")
add_eq(doc, r('Vertical (90°):  '), sub(i('t'),r('v')), r('='), frac(r('2')+i('v')+sub(r(''),r('0')), i('g')),
            r('='), r('2'), sqrt(frac(r('2')+i('h'), i('g'))))
add_eq(doc, r('Range (45°):  '), sub(i('t'),r('R')), r('='), frac(r('2')+i('v')+sub(r(''),r('0'))+r('sin',True)+r('45°'), i('g')),
            r('='), r('2'), sqrt(frac(i('h'), i('g'))))
answer_text(doc, "Vertical jump: t = 2√(2h/g).   45° jump: t = 2√(h/g)  (a factor √2 shorter).",
            label='Answer (b):')

out = "/home/user/AgenteSAE/talleres/Taller_Kinematics_2D_English.docx"
doc.save(out)
print("Saved", out, "-- paragraphs:", len(doc.paragraphs))
