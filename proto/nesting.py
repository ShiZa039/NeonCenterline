"""Прототип: одинаковые формы и раскладка на листы (src/NeonTools.bas, разделы @@SAMEGEO, @@NESTGEO).

Одинаковые формы
    Сравнивается сам контур, а не узлы: точки по контуру через 0.5 мм, сдвинутые к левому нижнему
    углу габарита. Два объекта одинаковые, если у них один цвет, габариты и длины контура почти
    равны и каждая точка одного контура — не дальше tol (0.1 мм) от другого, в обе стороны. Число
    узлов, их расстановка, начальная точка и направление обхода значения не имеют (одна и та же
    цифра из разных переводов текста в кривые); поворот и зеркало — имеют.

Раскладка
    MaxRects без поворота. Деталь — габарит + зазор, лист — без полей + зазор: так между деталями
    ровно зазор, от края — ровно поле. Детали по очереди, каждая — на самый ранний лист, где есть
    место (FFD), на листе — лучшее свободное место по эвристике. Несколько порядков и эвристик,
    лучший результат: меньше листов, затем меньше занятый габарит последнего листа.
    Координаты на листе: от левого верхнего угла листа, y вниз.
    Ширина по раскладке: высота листа задана, ширина — сколько заняли детали + поле, вверх до 10 мм,
    от wmin до W; детали тогда идут ещё и «по столбцам» (левее, затем выше), чтобы лист был у́же.
"""

import math

EPS = 1e-6


# ---------------------------------------------------------------- одинаковые формы
SAME_STEP = 0.5      # шаг точек по контуру, мм
SAME_MAXPTS = 3000   # точек на объект не больше: у больших объектов шаг крупнее


def _d(ax, ay, bx, by):
    return math.sqrt((bx - ax) * (bx - ax) + (by - ay) * (by - ay))


def _bez(p0, p1, p2, p3, t):
    mt = 1 - t
    a = mt * mt * mt
    b = 3 * mt * mt * t
    c = 3 * mt * t * t
    d = t * t * t
    return (a * p0[0] + b * p1[0] + c * p2[0] + d * p3[0], a * p0[1] + b * p1[1] + c * p2[1] + d * p3[1])


def shape(subs, ck='0'):
    """Объект для same_classes из подпутей [(сегменты, замкнут)]; сегмент — ('L', p0, p3) или
    ('C', p0, p1, p2, p3). Контур — точки через шаг (у кривой — равные доли параметра), сдвинутые к
    левому нижнему углу габарита точек."""
    lc = 0.0
    for segs, cl in subs:
        for sg in segs:
            if sg[0] == 'L':
                lc += _d(*sg[1], *sg[2])
            else:
                lc += _d(*sg[1], *sg[2]) + _d(*sg[2], *sg[3]) + _d(*sg[3], *sg[4])
    step = max(SAME_STEP, lc / SAME_MAXPTS)
    polys = []
    for segs, cl in subs:
        P = []
        for sg in segs:
            if sg[0] == 'L':
                a, b = sg[1], sg[2]
                n = max(1, math.ceil(_d(*a, *b) / step))
                for i in range(n):
                    P.append((a[0] + (b[0] - a[0]) * i / n, a[1] + (b[1] - a[1]) * i / n))
            else:
                L = _d(*sg[1], *sg[2]) + _d(*sg[2], *sg[3]) + _d(*sg[3], *sg[4])
                n = max(2, math.ceil(L / step))
                for i in range(n):
                    P.append(_bez(*sg[1:], i / n))
        if not cl:
            P.append(segs[-1][-1])
        polys.append([P, cl])
    mx = min(x for P, cl in polys for x, y in P)
    my = min(y for P, cl in polys for x, y in P)
    w = h = ln = 0.0
    for q in polys:
        q[0] = [(x - mx, y - my) for x, y in q[0]]
        for x, y in q[0]:
            w = max(w, x); h = max(h, y)
        P = q[0]
        for i in range(len(P) - 1):
            ln += _d(*P[i], *P[i + 1])
        if q[1] and len(P) > 1:
            ln += _d(*P[-1], *P[0])
    return dict(ck=ck, w=w, h=h, len=ln, polys=polys)


def _near(A, B, tol):
    """Каждая точка контура A — не дальше tol от контура B (сетка по отрезкам B)."""
    segs = []
    for P, cl in B['polys']:
        for i in range(len(P) - 1):
            segs.append((P[i], P[i + 1]))
        if cl and len(P) > 1:
            segs.append((P[-1], P[0]))
    if not segs:
        return False
    cs = max(4 * tol, math.sqrt((B['w'] + 2 * tol) * (B['h'] + 2 * tol) / len(segs)))
    nx = int((B['w'] + 2 * tol) / cs) + 1
    ny = int((B['h'] + 2 * tol) / cs) + 1
    cells = [[] for _ in range(nx * ny)]
    for k, (p, q) in enumerate(segs):
        # сетка от (-tol, -tol): отрезок с запасом tol — в ячейки своего габарита
        ix0 = max(0, math.floor(min(p[0], q[0]) / cs))
        ix1 = min(nx - 1, math.floor((max(p[0], q[0]) + 2 * tol) / cs))
        iy0 = max(0, math.floor(min(p[1], q[1]) / cs))
        iy1 = min(ny - 1, math.floor((max(p[1], q[1]) + 2 * tol) / cs))
        for iy in range(iy0, iy1 + 1):
            for ix in range(ix0, ix1 + 1):
                cells[iy * nx + ix].append(k)
    t2 = tol * tol
    for P, cl in A['polys']:
        for px, py in P:
            ix = math.floor((px + tol) / cs)
            iy = math.floor((py + tol) / cs)
            if ix < 0 or ix >= nx or iy < 0 or iy >= ny:
                return False
            ok = False
            for k in cells[iy * nx + ix]:
                (ax, ay), (bx, by) = segs[k]
                dx = bx - ax; dy = by - ay
                L = dx * dx + dy * dy
                t = 0.0
                if L > 0:
                    t = ((px - ax) * dx + (py - ay) * dy) / L
                    if t < 0:
                        t = 0.0
                    elif t > 1:
                        t = 1.0
                ex = ax + t * dx - px
                ey = ay + t * dy - py
                if ex * ex + ey * ey <= t2:
                    ok = True
                    break
            if not ok:
                return False
    return True


def same_eq(a, b, tol):
    """Одинаковые: цвет, габарит (±tol), длина контура (±1% + 2tol), и каждая точка одного контура —
    не дальше tol от другого, в обе стороны. Узлы и их число значения не имеют."""
    if a['ck'] != b['ck'] or abs(a['w'] - b['w']) > tol or abs(a['h'] - b['h']) > tol:
        return False
    if abs(a['len'] - b['len']) > 0.01 * max(a['len'], b['len']) + 2 * tol:
        return False
    return _near(a, b, tol) and _near(b, a, tol)


def same_classes(objs, tol=0.1):
    """Номер класса каждого объекта (по порядку появления): объект сравнивается с первым объектом
    каждого класса."""
    reps = []
    cls = []
    for i, o in enumerate(objs):
        c = -1
        for k, r in enumerate(reps):
            if same_eq(objs[r], o, tol):
                c = k
                break
        if c < 0:
            c = len(reps)
            reps.append(i)
        cls.append(c)
    return cls


def subs_closed(subs, eps=1e-6):
    """Подпути формата proto/cleanup.py -> [(сегменты, замкнут)]."""
    out = []
    for s in subs:
        a, b = s[0][1], s[-1][-1]
        out.append((s, abs(a[0] - b[0]) <= eps and abs(a[1] - b[1]) <= eps))
    return out


# ---------------------------------------------------------------- раскладка
def _contains(b, a):
    return (a[1] >= b[1] - EPS and a[2] >= b[2] - EPS and a[1] + a[3] <= b[1] + b[3] + EPS
            and a[2] + a[4] <= b[2] + b[4] + EPS)


def _split(F, s, x, y, w, h, mw=0.0, mh=0.0):
    """Свободные прямоугольники, которые задевает деталь (x, y, w, h) на листе s, — на части;
    части внутри других — долой. Уже́ mw или ниже mh (меньше любой оставшейся детали) — тоже долой."""
    keep, new = [], []
    for r in F:
        rs, rx, ry, rw, rh = r
        if rw < mw - EPS or rh < mh - EPS:
            continue
        if rs != s or x >= rx + rw - EPS or x + w <= rx + EPS or y >= ry + rh - EPS or y + h <= ry + EPS:
            keep.append(r)
            continue
        if x > rx + EPS:
            new.append([s, rx, ry, x - rx, rh])
        if x + w < rx + rw - EPS:
            new.append([s, x + w, ry, rx + rw - (x + w), rh])
        if y > ry + EPS:
            new.append([s, rx, ry, rw, y - ry])
        if y + h < ry + rh - EPS:
            new.append([s, rx, y + h, rw, ry + rh - (y + h)])
    new = [a for a in new if a[3] >= mw - EPS and a[4] >= mh - EPS]
    out = list(keep)
    for ia, a in enumerate(new):
        drop = False
        for ib, b in enumerate(new):
            if ib != ia and _contains(b, a) and (not _contains(a, b) or ib < ia):
                drop = True
                break
        if not drop:
            for b in keep:
                if b[0] == s and _contains(b, a):
                    drop = True
                    break
        if not drop:
            out.append(a)
    F[:] = out


def _order(PW, PH, mode):
    idx = list(range(len(PW)))
    if mode == 0:
        key = lambda i: (-PH[i], -PW[i], i)
    elif mode == 1:
        key = lambda i: (-PW[i], -PH[i], i)
    elif mode == 2:
        key = lambda i: (-PW[i] * PH[i], -PH[i], i)
    else:
        key = lambda i: (-max(PW[i], PH[i]), -min(PW[i], PH[i]), i)
    idx.sort(key=key)
    return idx


def pack_run(PW, PH, BW, BH, order, heur):
    """Один проход: (лист, x, y) каждой детали (лист -1 — не помещается), число листов."""
    F = []
    nsh = 0
    res = [(-1, 0.0, 0.0)] * len(PW)
    # наименьшие ширина и высота среди деталей после q-й
    n = len(order)
    sw = [float('inf')] * (n + 1)
    sh = [float('inf')] * (n + 1)
    for q in range(n - 1, -1, -1):
        sw[q] = min(sw[q + 1], PW[order[q]])
        sh[q] = min(sh[q + 1], PH[order[q]])
    for q, i in enumerate(order):
        pw, ph = PW[i], PH[i]
        if pw > BW + EPS or ph > BH + EPS:
            continue
        best = None
        for j, (s, x, y, w, h) in enumerate(F):
            if pw <= w + EPS and ph <= h + EPS:
                if heur == 0:      # лучшая короткая сторона
                    s1 = min(w - pw, h - ph); s2 = max(w - pw, h - ph)
                elif heur == 1:    # выше и левее (по строкам)
                    s1 = y + ph; s2 = x
                elif heur == 2:    # лучшая площадь
                    s1 = w * h - pw * ph; s2 = min(w - pw, h - ph)
                else:              # левее и выше (по столбцам: раскладка как можно уже)
                    s1 = x + pw; s2 = y
                if (best is None or s < best[0] or (s == best[0] and (s1 < best[1] or (s1 == best[1] and s2 < best[2])))):
                    best = (s, s1, s2, j)
        if best is None:
            s = nsh
            nsh += 1
            F.append([s, 0.0, 0.0, BW, BH])
            x = y = 0.0
        else:
            s = best[0]
            x, y = F[best[3]][1], F[best[3]][2]
        res[i] = (s, x, y)
        _split(F, s, x, y, pw, ph, sw[q + 1], sh[q + 1])
    return res, nsh


def last_area(res, PW, PH, nsh):
    if nsh == 0:
        return 0.0
    mx = my = 0.0
    for i, (s, x, y) in enumerate(res):
        if s == nsh - 1:
            mx = max(mx, x + PW[i]); my = max(my, y + PH[i])
    return mx * my


def used_widths(res, PW, nsh):
    """Занятая ширина каждого листа (с зазором справа)."""
    u = [0.0] * nsh
    for i, (s, x, y) in enumerate(res):
        if s >= 0 and x + PW[i] > u[s]:
            u[s] = x + PW[i]
    return u


def pack(W, H, margin, gap, sizes, auto=False, wmin=0.0, fill=2):
    """sizes: [(w, h)] деталей -> (размещение [(лист, x, y)] от левого верхнего угла листа, число листов,
    порядок, эвристика, ширина каждого листа). auto: ширина листа — по раскладке, от wmin до W
    (вверх до 10 мм); тогда ещё эвристика «по столбцам», а при равном числе листов лучше меньшая
    общая ширина. fill: 0 — столбцами (сверху вниз, затем вправо), 1 — рядами (слева направо, затем
    вниз), 2 — как плотнее (все эвристики)."""
    PW = [w + gap for w, h in sizes]
    PH = [h + gap for w, h in sizes]
    BW = W - 2 * margin + gap
    BH = H - 2 * margin + gap
    best = None
    for mode in range(4):
        order = _order(PW, PH, mode)
        heurs = [3] if fill == 0 else [1] if fill == 1 else list(range(4 if auto else 3))
        for heur in heurs:
            res, nsh = pack_run(PW, PH, BW, BH, order, heur)
            la = sum(used_widths(res, PW, nsh)) if auto else last_area(res, PW, PH, nsh)
            if best is None or nsh < best[1] or (nsh == best[1] and la < best[2] - EPS):
                best = (res, nsh, la, mode, heur)
    res, nsh = best[0], best[1]
    out = [(s, x + margin, y + margin) if s >= 0 else (-1, 0.0, 0.0) for s, x, y in res]
    widths = [W] * nsh
    if auto:
        for s, u in enumerate(used_widths(res, PW, nsh)):
            w = math.ceil((2 * margin + u - gap) / 10 - 1e-9) * 10
            widths[s] = min(max(w, wmin), W)
    return out, nsh, best[3], best[4], widths


def check(W, H, margin, gap, sizes, out, widths=None):
    """Детали внутри поля листа и не ближе зазора друг к другу."""
    for i, (s, x, y) in enumerate(out):
        if s < 0:
            continue
        w, h = sizes[i]
        if widths:
            W = widths[s]
        assert x >= margin - 1e-6 and y >= margin - 1e-6, (i, x, y)
        assert x + w <= W - margin + 1e-6 and y + h <= H - margin + 1e-6, (i, x, y, w, h)
    by = {}
    for i, (s, x, y) in enumerate(out):
        if s >= 0:
            by.setdefault(s, []).append(i)
    for s, L in by.items():
        for a in range(len(L)):
            i = L[a]
            xi, yi = out[i][1], out[i][2]
            for b in range(a + 1, len(L)):
                j = L[b]
                xj, yj = out[j][1], out[j][2]
                sep = (xi + sizes[i][0] + gap <= xj + 1e-6 or xj + sizes[j][0] + gap <= xi + 1e-6
                       or yi + sizes[i][1] + gap <= yj + 1e-6 or yj + sizes[j][1] + gap <= yi + 1e-6)
                assert sep, (s, i, j)
    return True
