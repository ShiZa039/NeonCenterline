"""Стенд: BoxGeo (кубик под фрезеровку, раздел @@BOXGEO) в LibreOffice против proto/box.py box().
Запуск из корня репозитория:  python tools/stands/box.py  ->  mismatches: 0

Наборы: пример автора (500 × 500 × 500, материал 5, ступенька 2, зазор 0,2, выход 2), разные стороны,
тонкий и толстый материал, без зазора и без выхода, центр не в нуле. Сравниваются детали (угол,
размер, вид) и фрезеровка ступеньки (деталь, вид, концы) с допуском 1e-6 мм.
"""
import os, sys
HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(os.path.dirname(HERE))
sys.path.insert(0, os.path.join(ROOT, 'tools'))
sys.path.insert(0, os.path.join(ROOT, 'proto'))
import lo_stand as S
import box as B

# (cx, cy, L, W, H, t, e, c, o)
CASES = {
    'example': (0, 0, 500, 500, 500, 5, 2, 0.2, 2),
    'long':    (148.5, 105, 800, 300, 450, 6, 2.5, 0.2, 3),
    'tall':    (-20, 37.25, 200, 250, 1200, 10, 4, 0.3, 5),
    'thin':    (1000, -500, 120, 90, 60, 3, 1, 0.1, 1),
    'nogap':   (0, 0, 333.3, 222.2, 111.1, 4, 1.5, 0, 0),
}

main = ['Sub Main', '    Dim g As Integer, i As Long', '    g = FreeFile', f'    Open "{S.OUT}" For Output As #g']
for name, a in CASES.items():
    main += [f'    BoxGeo {", ".join(repr(float(v)) for v in a)}',
             f'    Print #g, "C {name} " & NBx & " " & NBr',
             '    For i = 0 To NBx - 1',
             '        Print #g, Format(BxX(i), "0.000000") & " " & Format(BxY(i), "0.000000") & " " & '
             'Format(BxW(i), "0.000000") & " " & Format(BxH(i), "0.000000") & " " & BxK(i)',
             '    Next i',
             '    For i = 0 To NBr - 1',
             '        Print #g, BrP(i) & " " & BrT(i) & " " & Format(BrX1(i), "0.000000") & " " & Format(BrY1(i), "0.000000") & " " & '
             'Format(BrX2(i), "0.000000") & " " & Format(BrY2(i), "0.000000")',
             '    Next i']
main += ['    Close #g', 'End Sub']

code = S.module(S.section('NEST_DECL'), S.section('BOXGEO'), main)
L = S.run(code)

i = bad = 0
while i < len(L):
    _, name, nb, nr = L[i].split()
    nb, nr = int(nb), int(nr)
    i += 1
    vp = [tuple(map(float, L[i + j].replace(',', '.').split())) for j in range(nb)]
    i += nb
    vr = [tuple(map(float, L[i + j].replace(',', '.').split())) for j in range(nr)]
    i += nr
    pp, pr = B.box(*CASES[name])
    d = float('inf')
    if len(pp) == len(vp) and len(pr) == len(vr):
        d = max(max(abs(a - b) for a, b in zip(x, y)) for x, y in list(zip(pp, vp)) + list(zip(pr, vr)))
    ok = d < 1e-6
    bad += not ok
    print(f'{name:8} VBA {nb} деталей {nr} ступенек  Python {len(pp)} / {len(pr)}  max diff {d:.1e}  {"OK" if ok else "MISMATCH"}')
print('mismatches:', bad)
sys.exit(1 if bad else 0)
