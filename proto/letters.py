"""Прототип: разбить на буквы (src/NeonTools.bas, раздел @@LETTERSGEO).

Кривая из многих контуров (текст в кривых, объединённые буквы) -> по объекту на букву, дырки
остаются в своих буквах. Контуры — точки через шаг, как у «Одинаковых» (nesting.shape). Для
каждого замкнутого контура — точка на нём (середина самого длинного отрезка) и сколько других
контуров её окружают (чёт-нечет): чётное число — наружный контур буквы, нечётное — дырка; дырка
идёт к самому маленькому из окружающих контуров. Остров в дырке («®», буква в «О») — отдельная
буква. Незамкнутые контуры — каждый отдельно. Буквы — слева направо (по левому краю, затем сверху).
"""
import math
from nesting import shape, _d

NEAR_CLOSED = 0.05


def _area(P):
    a = 0.0
    n = len(P)
    for i in range(n):
        x0, y0 = P[i]
        x1, y1 = P[(i + 1) % n]
        a += x0 * y1 - x1 * y0
    return a / 2


def _inside(px, py, P):
    c = False
    n = len(P)
    for i in range(n):
        ax, ay = P[i]
        bx, by = P[(i + 1) % n]
        if (ay > py) != (by > py):
            if px < (bx - ax) * (py - ay) / (by - ay) + ax:
                c = not c
    return c


def split(subs):
    """subs: [(сегменты, замкнут)] -> список букв: номера подпутей (наружный первым, затем дырки)."""
    sh = shape(subs)
    polys = sh['polys']
    n = len(polys)
    closed = []
    for (segs, cl) in subs:
        a, b = segs[0][1], segs[-1][-1]
        closed.append(cl or _d(*a, *b) < NEAR_CLOSED)
    bb = []
    probe = []
    area = []
    for k in range(n):
        P = polys[k][0]
        bb.append((min(x for x, y in P), min(y for x, y in P), max(x for x, y in P), max(y for x, y in P)))
        if closed[k] and len(P) >= 3:
            best = -1.0
            pr = P[0]
            for i in range(len(P)):
                ax, ay = P[i]
                bx, by = P[(i + 1) % len(P)]
                L = _d(ax, ay, bx, by)
                if L > best:
                    best = L
                    pr = ((ax + bx) / 2, (ay + by) / 2)
            probe.append(pr)
            area.append(abs(_area(P)))
        else:
            probe.append(None)
            area.append(0.0)
    depth = [0] * n
    parent = [-1] * n
    for i in range(n):
        if probe[i] is None:
            continue
        px, py = probe[i]
        for j in range(n):
            if j == i or probe[j] is None:
                continue
            if bb[i][0] < bb[j][0] - 1e-9 or bb[i][1] < bb[j][1] - 1e-9 or bb[i][2] > bb[j][2] + 1e-9 or bb[i][3] > bb[j][3] + 1e-9:
                continue
            if _inside(px, py, polys[j][0]):
                depth[i] += 1
                if parent[i] < 0 or area[j] < area[parent[i]]:
                    parent[i] = j
    groups = []
    for o in range(n):
        if probe[o] is None:
            groups.append([o])
        elif depth[o] % 2 == 0:
            groups.append([o] + [i for i in range(n) if probe[i] is not None and depth[i] % 2 == 1 and parent[i] == o])
    # дырка без наружного (не должно быть) — отдельно
    used = set(i for g in groups for i in g)
    for i in range(n):
        if i not in used:
            groups.append([i])
    groups.sort(key=lambda g: (bb[g[0]][0], -bb[g[0]][3], g[0]))
    return groups
