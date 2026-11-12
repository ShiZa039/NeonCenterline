"""Пример стенда (образец для новых): GromHem — контур под подгиб у люверсов — в LibreOffice
против proto/grommets.py hem(). Запуск из корня репозитория:  python tools/stands/grom_hem.py

Схема любого стенда:
  1) входные данные — одни и те же для Python и для VBA (здесь — контуры из сегментов Безье);
  2) Main кладёт их туда же, куда их кладёт код для Corel (здесь — PutS в SX/SY, как ShapeGrommets),
     зовёт процедуру раздела и печатает результат в файл S.OUT с 6 знаками;
  3) Python считает то же прототипом и сравнивает поточечно с допуском.
"""
import math, os, sys
HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(os.path.dirname(HERE))
sys.path.insert(0, os.path.join(ROOT, 'tools'))
sys.path.insert(0, os.path.join(ROOT, 'proto'))
import lo_stand as S
import grommets as G


def lines(P):
    """Замкнутый многоугольник -> сегменты-прямые в виде Безье (как отдаёт Corel)."""
    n = len(P)
    out = []
    for k in range(n):
        a, b = P[k], P[(k + 1) % n]
        out.append((a, (a[0] + (b[0] - a[0]) / 3, a[1] + (b[1] - a[1]) / 3),
                    (a[0] + 2 * (b[0] - a[0]) / 3, a[1] + 2 * (b[1] - a[1]) / 3), b, True))
    return out


w = 1450
k = 0.5522847498 * 500
CASES = {
    'rect': lines([(0, 0), (3000, 0), (3000, 1000), (0, 1000)]),
    'rect_cw': lines([(0, 1000), (3000, 1000), (3000, 0), (0, 0)]),
    'cross': lines([(w, 0), (2 * w, 0), (2 * w, w), (3 * w, w), (3 * w, 2 * w), (2 * w, 2 * w), (2 * w, 3 * w),
                    (w, 3 * w), (w, 2 * w), (0, 2 * w), (0, w), (w, w)]),
    'tri': lines([(0, 0), (1257, 0), (0, 1980)]),
    'circle': [((500, 0), (500, k), (k, 500), (0, 500), False), ((0, 500), (-k, 500), (-500, k), (-500, 0), False),
               ((-500, 0), (-500, -k), (-k, -500), (0, -500), False), ((0, -500), (k, -500), (500, -k), (500, 0), False)],
}
HEM = 50

main = ['Sub Main', '    Dim g As Integer, i As Long', '    g = FreeFile', f'    Open "{S.OUT}" For Output As #g']
for name, segs in CASES.items():
    main += S.put_segs(segs)
    main += [f'    GromHem {HEM}', f'    Print #g, "C {name} " & GHN',
             '    For i = 0 To GHN - 1',
             '        Print #g, Format(GHX(i), "0.000000") & " " & Format(GHY(i), "0.000000")',
             '    Next i']
main += ['    Close #g', 'End Sub']

# объявления уровня модуля — вперёд (как в VBA), затем раздел; PutS/NormV/BzPt module() допишет сам
code = S.module(S.section('SIMP_DECL'), S.section('GROM_DECL'), S.section('GROMGEO'), main)
L = S.run(code)

i = bad = 0
while i < len(L):
    _, name, n = L[i].split()
    n = int(n)
    i += 1
    vb = [tuple(map(float, L[i + j].split())) for j in range(n)]
    i += n
    py = G.hem(CASES[name], HEM)
    d = max((math.dist(a, b) for a, b in zip(py, vb)), default=0) if len(py) == len(vb) else float('inf')
    ok = d < 1e-4
    bad += not ok
    print(f'{name:8} VBA {n:3}  Python {len(py):3}  max diff {d:.2e}  {"OK" if ok else "MISMATCH"}')
print('mismatches:', bad)
sys.exit(1 if bad else 0)
