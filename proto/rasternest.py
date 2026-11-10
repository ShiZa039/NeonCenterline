"""Прототип: раскладка по контуру (src/NeonTools.bas, раздел @@NESTRASTER).

Деталь — растр: клетки r × r, центр которых внутри контура (чёт-нечет) или не дальше
d = g/2 + r·0.7072 от него. Это с запасом: две детали на непересекающихся клетках не ближе зазора g.
Дырки букв остаются пустыми — в них может встать мелкая деталь. Повороты на 90/180/270° — тот же
набор клеток, перевёрнутый по индексам (точно: клетки квадратные).

Лист — сетка рабочего поля (поле от края минус g/2). У каждой строки — длина свободного (R) и
занятого (O) куска от каждой клетки и самый длинный свободный кусок (M): место для детали
проверяется кусками её строк, занятое пропускается скачком. Деталь — в первое свободное место по
строкам (сверху, затем слева); «столбцами» — тот же поиск на транспонированной сетке (слева, затем
сверху). Из поворотов — тот, у которого правый (нижний) край ближе. Детали — от больших к
маленьким, каждая — на самый ранний открытый лист (открыты последние три), иначе — новый лист.
Два порядка (по площади клеток и по площади габарита) — лучший: меньше листов, затем короче.
"""
import math

RN_OPEN = 3          # открытых листов
RN_CELLS = 1000000   # клеток на лист: шаг r = sqrt(площадь / RN_CELLS), от 0.5 до 5 мм
RN_SQ = 0.7072       # половина диагонали клетки (√2/2 с запасом)


def grid_step(sw, sh, margin, gap):
    a = max(1.0, (sw - 2 * margin + gap) * (sh - 2 * margin + gap))
    return min(5.0, max(0.5, math.sqrt(a / RN_CELLS)))


def _seg_row(au, av, bu, bv, c, d):
    """Отрезок AB, раздутый на d, пересекает прямую v = c: (lo, hi) или None."""
    lo = 1e300
    hi = -1e300
    for pu, pv in ((au, av), (bu, bv)):
        e = d * d - (c - pv) * (c - pv)
        if e >= 0:
            e = math.sqrt(e)
            if pu - e < lo:
                lo = pu - e
            if pu + e > hi:
                hi = pu + e
    du = bu - au
    dv = bv - av
    L2 = du * du + dv * dv
    if L2 > 0:
        L = math.sqrt(L2)
        nu = -dv / L
        nv = du / L
        # |(x - au)·nu + (c - av)·nv| <= d  и  0 <= (x - au)·du + (c - av)·dv <= L2
        k = (c - av) * nv
        if nu != 0:
            x1 = au + (-d - k) / nu
            x2 = au + (d - k) / nu
            s0, s1 = (x1, x2) if x1 < x2 else (x2, x1)
        elif abs(k) <= d:
            s0, s1 = -1e300, 1e300
        else:
            s0, s1 = 1.0, -1.0
        m = (c - av) * dv
        if du != 0:
            x1 = au + (0 - m) / du
            x2 = au + (L2 - m) / du
            t0, t1 = (x1, x2) if x1 < x2 else (x2, x1)
        elif 0 <= m <= L2:
            t0, t1 = -1e300, 1e300
        else:
            t0, t1 = 1.0, -1.0
        a0 = s0 if s0 > t0 else t0
        a1 = s1 if s1 < t1 else t1
        if a0 <= a1:
            if a0 < lo:
                lo = a0
            if a1 > hi:
                hi = a1
    if lo > hi:
        return None
    return lo, hi


def part_cells(polys, w, h, r, gap):
    """polys: [(точки, замкнут)] в координатах детали (y вверх, габарит [0, w] × [0, h]).
    -> растр (y вниз): cols, rows, занятые клетки (bytearray), pad — сдвиг габарита от угла растра."""
    d = gap / 2 + r * RN_SQ
    pad = d + 0.5 * r
    P = [([(x, h - y) for x, y in pts], cl) for pts, cl in polys]
    cols = int(math.ceil((w + 2 * pad) / r))
    rows = int(math.ceil((h + 2 * pad) / r))
    g = bytearray(cols * rows)
    # внутри: пересечения рёбер с прямыми через центры строк — по рёбрам, в списки строк
    X = [[] for _ in range(rows)]
    for q, cl in P:
        n = len(q)
        if not cl:
            continue
        for k in range(n):
            au, av = q[k]
            bu, bv = q[(k + 1) % n]
            if av == bv:
                continue
            lo, hi = (av, bv) if av < bv else (bv, av)
            # строки j, у которых lo < c <= hi (центр c = -pad + (j + 0.5)·r); (a > c) != (b > c)
            j0 = max(0, math.floor((lo + pad) / r - 0.5) + 1)
            j1 = min(rows - 1, math.floor((hi + pad) / r - 0.5))
            for j in range(j0, j1 + 1):
                c = -pad + (j + 0.5) * r
                if (av > c) != (bv > c):
                    X[j].append(au + (c - av) * (bu - au) / (bv - av))
    for j in range(rows):
        xs = X[j]
        xs.sort()
        for k in range(0, len(xs) - 1, 2):
            i0 = max(0, math.ceil((xs[k] + pad) / r - 0.5))
            i1 = min(cols - 1, math.floor((xs[k + 1] + pad) / r - 0.5))
            for i in range(i0, i1 + 1):
                g[j * cols + i] = 1
    for q, cl in P:
        n = len(q)
        for k in range(n if cl else n - 1):
            au, av = q[k]
            bu, bv = q[(k + 1) % n]
            j0 = max(0, math.ceil((min(av, bv) - d + pad) / r - 0.5))
            j1 = min(rows - 1, math.floor((max(av, bv) + d + pad) / r - 0.5))
            for j in range(j0, j1 + 1):
                c = -pad + (j + 0.5) * r
                iv = _seg_row(au, av, bu, bv, c, d)
                if iv is None:
                    continue
                i0 = max(0, math.ceil((iv[0] + pad) / r - 0.5))
                i1 = min(cols - 1, math.floor((iv[1] + pad) / r - 0.5))
                for i in range(i0, i1 + 1):
                    g[j * cols + i] = 1
    return cols, rows, g, pad


def fill_holes(g, cols, rows):
    """Без «матрёшки»: пустые клетки, до которых не дойти от края растра по соседям (сверху, снизу,
    слева, справа), — занятыми. Дырки букв и рамок закрываются, открытые выемки («П», «С») остаются."""
    g = bytearray(g)
    st = []
    for c in list(range(cols)) + [(rows - 1) * cols + i for i in range(cols)] + \
             [j * cols for j in range(rows)] + [j * cols + cols - 1 for j in range(rows)]:
        if g[c] == 0:
            g[c] = 2; st.append(c)
    while st:
        c = st.pop()
        i, j = c % cols, c // cols
        for ok, n in ((i > 0, c - 1), (i < cols - 1, c + 1), (j > 0, c - cols), (j < rows - 1, c + cols)):
            if ok and g[n] == 0:
                g[n] = 2; st.append(n)
    return bytearray(1 if v != 2 else 0 for v in g)


def orient(cols, rows, g, pad, w, h, r, k, tr):
    """Поворот k (0 — нет, 1 — 90° по часовой на экране, 2 — 180°, 3 — 90° против) и транспонирование
    (tr — для поиска «столбцами»). -> куски строк [(j, i, длина)], правый край, нижний край, сдвиг
    габарита детали (u, v) от угла растра — в координатах листа (без транспонирования), размер
    повёрнутого габарита (w, h)."""
    if k == 0:
        C, Rw = cols, rows
        f = lambda i, j: (i, j)
        off = (pad, pad); wh = (w, h)
    elif k == 1:
        C, Rw = rows, cols
        f = lambda i, j: (rows - 1 - j, i)
        off = (rows * r - pad - h, pad); wh = (h, w)
    elif k == 2:
        C, Rw = cols, rows
        f = lambda i, j: (cols - 1 - i, rows - 1 - j)
        off = (cols * r - pad - w, rows * r - pad - h); wh = (w, h)
    else:
        C, Rw = rows, cols
        f = lambda i, j: (j, cols - 1 - i)
        off = (pad, cols * r - pad - w); wh = (h, w)
    G = bytearray(C * Rw)
    for j in range(rows):
        for i in range(cols):
            if g[j * cols + i]:
                a, b = f(i, j)
                G[b * C + a] = 1
    if tr:
        G2 = bytearray(C * Rw)
        for j in range(Rw):
            for i in range(C):
                if G[j * C + i]:
                    G2[i * Rw + j] = 1
        G, C, Rw = G2, Rw, C
    runs = []
    cnt = 0
    for j in range(Rw):
        i = 0
        while i < C:
            if G[j * C + i]:
                s = i
                while i < C and G[j * C + i]:
                    i += 1
                runs.append((j, s, i - s))
                cnt += i - s
            else:
                i += 1
    runs.sort(key=lambda t: (-t[2], t[0], t[1]))
    right = max(i + ln for j, i, ln in runs)
    low = max(j for j, i, ln in runs) + 1
    return runs, right, low, off, wh, cnt


class Sheet:
    def __init__(self, nx, ny):
        self.nx, self.ny = nx, ny
        self.R = [list(range(nx, 0, -1)) for _ in range(ny)]
        self.O = [[0] * nx for _ in range(ny)]
        self.M = [nx] * ny

    def find(self, runs, right, low):
        nx, ny = self.nx, self.ny
        j0, i0, l0 = runs[0]
        for Y in range(0, ny - low + 1):
            if self.M[Y + j0] < l0:
                continue
            X = 0
            while X + right <= nx:
                ok = True
                for j, i, ln in runs:
                    x = X + i
                    f = self.R[Y + j][x]
                    if f < ln:
                        ok = False
                        if f == 0:
                            X += self.O[Y + j][x]
                        else:
                            q = x + f
                            X += f + (self.O[Y + j][q] if q < nx else nx)
                        break
                if ok:
                    return X, Y
        return None

    def put(self, runs, X, Y):
        rows = set()
        for j, i, ln in runs:
            Rr = self.R[Y + j]
            for x in range(X + i, X + i + ln):
                Rr[x] = 0
            rows.add(Y + j)
        nx = self.nx
        for y in sorted(rows):
            Rr, Or = self.R[y], self.O[y]
            nr = no = mx = 0
            for x in range(nx - 1, -1, -1):
                if Rr[x] == 0:
                    no += 1
                    nr = 0
                    Or[x] = no
                else:
                    nr += 1
                    no = 0
                    Rr[x] = nr
                    Or[x] = 0
                    if nr > mx:
                        mx = nr
            self.M[y] = mx


def pack_run(nx, ny, P, order, r, margin, gap, tr):
    sheets = []
    res = [(-1, 0.0, 0.0, 0)] * len(P)
    nsh = 0
    for k in order:
        vs = [v for v in P[k] if v[1] <= nx and v[2] <= ny]
        if not vs:
            continue
        best = None
        for s, S in sheets:
            for v in vs:
                p = S.find(v[0], v[1], v[2])
                if p is not None:
                    sc = (p[1] + v[2], p[0])
                    if best is None or sc < best[0]:
                        best = (sc, s, S, p, v)
            if best is not None:
                break
        if best is None:
            S = Sheet(nx, ny)
            s = nsh
            nsh += 1
            sheets.append((s, S))
            if len(sheets) > RN_OPEN:
                sheets.pop(0)
            for v in vs:
                p = S.find(v[0], v[1], v[2])
                sc = (p[1] + v[2], p[0])
                if best is None or sc < best[0]:
                    best = (sc, s, S, p, v)
        sc, s, S, (X, Y), v = best
        S.put(v[0], X, Y)
        if tr:
            X, Y = Y, X
        res[k] = (s, margin - gap / 2 + X * r + v[3][0], margin - gap / 2 + Y * r + v[3][1], v[5])
    return res, nsh


def pack(sw, sh, margin, gap, parts, cols_first=True, auto=False, wmin=0.0, rots=(0,), inside=True):
    """parts: [(polys, w, h)] -> (размещение [(лист, x, y, поворот)] — левый верхний угол габарита
    повёрнутой детали от угла листа, y вниз; лист -1 — не помещается), число листов, ширины листов, шаг.
    cols_first: True — столбцами, False — рядами, None — оба (лучший). inside: False — не класть детали
    внутрь других (дырки деталей закрыты)."""
    r = grid_step(sw, sh, margin, gap)
    trs = [True, False] if cols_first is None else [cols_first]
    nxo = int(math.floor((sw - 2 * margin + gap) / r))
    nyo = int(math.floor((sh - 2 * margin + gap) / r))
    P = {tr: [] for tr in trs}
    area = []
    for polys, w, h in parts:
        cols, rows, g, pad = part_cells(polys, w, h, r, gap)
        if not inside:
            g = fill_holes(g, cols, rows)
        area.append(sum(g))
        for tr in trs:
            vs = []
            for k in rots:
                runs, right, low, off, wh, cnt = orient(cols, rows, g, pad, w, h, r, k, tr)
                vs.append((runs, right, low, off, wh, k, cnt))
            P[tr].append(vs)
    best = None
    for tr in trs:
        nx, ny = (nyo, nxo) if tr else (nxo, nyo)
        for om in (0, 1):
            if om == 0:
                key = lambda k: (-area[k], k)
            else:
                key = lambda k: (-parts[k][1] * parts[k][2], k)
            order = sorted(range(len(parts)), key=key)
            res, nsh = pack_run(nx, ny, P[tr], order, r, margin, gap, tr)
            used = [0.0] * nsh
            for k, (s, x, y, rt) in enumerate(res):
                if s >= 0:
                    ww = parts[k][1] if rt % 2 == 0 else parts[k][2]
                    used[s] = max(used[s], x + ww)
            sc = (nsh, used[-1] if nsh else 0.0)
            if best is None or sc[0] < best[0][0] or (sc[0] == best[0][0] and sc[1] < best[0][1] - 1e-9):
                best = (sc, res, nsh, used)
    sc, res, nsh, used = best
    widths = [sw] * nsh
    if auto:
        for s in range(nsh):
            wv = math.ceil((used[s] + margin) / 10 - 1e-9) * 10
            widths[s] = min(max(wv, wmin), sw)
    return res, nsh, widths, r
