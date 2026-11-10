"""Прототип: разбить на буквы (src/NeonTools.bas, раздел @@LETTERSGEO).

Кривая из многих контуров (текст в кривых, объединённые буквы) -> по объекту на букву, дырки
остаются в своих буквах. Контуры — точки через шаг, как у «Одинаковых» (nesting.shape). Для
каждого замкнутого контура — точка на нём (середина самого длинного отрезка) и сколько других
контуров её окружают (чёт-нечет): чётное число — наружный контур буквы, нечётное — дырка; дырка
идёт к самому маленькому из окружающих контуров. Остров в дырке («®», буква в «О») — отдельная
буква. Незамкнутые контуры — каждый отдельно. Точки и акценты (Ё, Й, i, ї, !, ?, ÷) — со своей буквой
(attach_marks). Буквы — слева направо (по левому краю, затем сверху).
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
    groups = attach_marks(groups, bb, parent)
    groups.sort(key=lambda g: (bb[g[0]][0], -bb[g[0]][3], g[0]))
    return groups


MARK_SIZE = 0.4  # знак — не больше 0.4 буквы по большей стороне габарита,
MARK_AREA = 0.3  # или не больше 0.3 её площади (по габариту) и не больше 0.7 по большей стороне
MARK_SIZE2 = 0.7
MARK_X = 0.25    # центр знака — над буквой или под ней (её габарит по горизонтали + 0.25 её размера с каждой стороны)
MARK_GAP = 0.35  # и не дальше 0.35 её размера по вертикали


def attach_marks(groups, bb, par):
    """Точки и акценты — к своей букве (Ё, Й, i, ї, !, ?, ÷): кусок, который намного меньше буквы,
    стоит над ней или под ней (центр — выше или ниже её габарита) и близко по вертикали, идёт к ней; из таких букв — к той, у которой
    меньше зазор по вертикали плюс сдвиг центров по горизонтали. Только внутри одного уровня
    вложенности (тот же окружающий контур): буквы в дырке рамки к рамке не липнут, точка в нуле и
    буква в «®» остаются отдельными. Куски одного размера (двоеточие, «=», «ы») и стоящие рядом
    (точка после буквы, кольца «%») не сливаются."""
    n = len(groups)
    if n < 2:
        return groups
    box = [bb[g[0]] for g in groups]
    size = [max(b[2] - b[0], b[3] - b[1]) for b in box]
    area = [(b[2] - b[0]) * (b[3] - b[1]) for b in box]
    lvl = [par[g[0]] for g in groups]
    up = list(range(n))

    def find(i):
        while up[i] != i:
            i = up[i]
        return i
    for i in range(n):
        x0, y0, x1, y1 = box[i]
        cx = (x0 + x1) / 2
        cy = (y0 + y1) / 2
        best = -1
        bs = 0.0
        ba = 0.0
        for j in range(n):
            if j == i or lvl[j] != lvl[i]:
                continue
            if not (size[i] <= MARK_SIZE * size[j] or
                    (area[i] <= MARK_AREA * area[j] and size[i] <= MARK_SIZE2 * size[j])):
                continue
            u0, v0, u1, v1 = box[j]
            if cx < u0 - MARK_X * size[j] or cx > u1 + MARK_X * size[j]:
                continue
            # знак — над или под буквой, а не рядом (точка после «Ш.» — не её)
            if cy >= v0 and cy <= v1:
                continue
            gap = max(0.0, v0 - y1, y0 - v1)
            if gap > MARK_GAP * size[j]:
                continue
            sc = gap + abs(cx - (u0 + u1) / 2)
            if best < 0 or sc < bs or (sc == bs and area[j] > ba):
                best = j
                bs = sc
                ba = area[j]
        if best >= 0:
            a = find(i)
            b = find(best)
            if a != b:
                up[a] = b
    out = []
    for r in range(n):
        if find(r) != r:
            continue
        # кусок-корень первым (его наружный контур задаёт место буквы), затем остальные по номеру
        g = list(groups[r])
        for i in range(n):
            if i != r and find(i) == r:
                g += groups[i]
        out.append(g)
    return out
