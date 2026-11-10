"""Прототип: одинаковые формы и раскладка на листы (src/NeonTools.bas, разделы @@SAMEGEO, @@NESTGEO).

Одинаковые формы
    Объект — набор точек кривой: узлы и управляющие точки, относительно самого левого-нижнего узла
    (минимум x и минимум y по узлам), отсортированные по x, затем y. Два объекта одинаковые, если у
    них один цвет, одинаковое число подпутей, узлов и сегментов, габариты отличаются не больше чем
    на tol и каждой точке одного найдётся своя точка другого не дальше tol по x и по y. Порядок
    узлов, начальная точка и направление обхода значения не имеют, поворот и зеркало — имеют.

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
def norm_points(anchors, ctrls):
    """Точки объекта относительно минимума узлов, по x, затем y."""
    mx = min(p[0] for p in anchors)
    my = min(p[1] for p in anchors)
    P = [(x - mx, y - my) for x, y in list(anchors) + list(ctrls)]
    P.sort()
    return P


def pts_match(A, B, tol):
    n = len(A)
    if n != len(B):
        return False
    used = [False] * n
    lo = 0
    for ax, ay in A:
        while lo < n and B[lo][0] < ax - tol:
            lo += 1
        k = lo
        found = False
        while k < n and B[k][0] <= ax + tol:
            if not used[k] and abs(B[k][1] - ay) <= tol:
                used[k] = True
                found = True
                break
            k += 1
        if not found:
            return False
    return True


def cheap_eq(a, b, tol):
    return (a['ck'] == b['ck'] and a['nsub'] == b['nsub'] and a['nn'] == b['nn'] and a['ns'] == b['ns']
            and abs(a['w'] - b['w']) <= tol and abs(a['h'] - b['h']) <= tol)


def same_classes(objs, tol=0.1):
    """objs: dict(ck, nsub, nn, ns, w, h, pts) -> номер класса каждого объекта (по порядку появления)."""
    reps = []
    cls = []
    for i, o in enumerate(objs):
        c = -1
        for k, r in enumerate(reps):
            if cheap_eq(objs[r], o, tol) and pts_match(objs[r]['pts'], o['pts'], tol):
                c = k
                break
        if c < 0:
            c = len(reps)
            reps.append(i)
        cls.append(c)
    return cls


def obj_from_subs(subs, ck='0'):
    """Объект из подпутей в формате proto/cleanup.py: ('L', p0, p3) / ('C', p0, p1, p2, p3)."""
    from cleanup import _bez
    an, ct = [], []
    ns = 0
    xs, ys = [], []
    for s in subs:
        # узлы: начала сегментов; у незамкнутого подпути ещё конец последнего
        for sg in s:
            an.append(sg[1])
            if sg[0] == 'C':
                ct += [sg[2], sg[3]]
            ns += 1
        a, b = s[0][1], s[-1][-1]
        if abs(a[0] - b[0]) > 1e-6 or abs(a[1] - b[1]) > 1e-6:
            an.append(s[-1][-1])
    nn = len(an)
    for s in subs:
        for sg in s:
            if sg[0] == 'L':
                xs += [sg[1][0], sg[2][0]]; ys += [sg[1][1], sg[2][1]]
            else:
                for i in range(17):
                    p = _bez(*sg[1:], i / 16)
                    xs.append(p[0]); ys.append(p[1])
    return dict(ck=ck, nsub=len(subs), nn=nn, ns=ns, w=max(xs) - min(xs), h=max(ys) - min(ys),
                pts=norm_points(an, ct))


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
