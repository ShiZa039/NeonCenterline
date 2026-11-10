"""Прототип «Очистки контура»: убрать наложения и петли в контурах (шрифты, переведённые в кривые).

Контур — список замкнутых подконтуров, подконтур — список сегментов:
    ('L', p0, p1)              — прямой
    ('C', p0, c1, c2, p3)      — кубическая кривая Безье
Заливка — по правилу ненулевой обмотки (так CorelDRAW заливает текст в кривых): точка залита,
если сумма обходов контуров вокруг неё не ноль.

Шаги:
  1. Сегменты разбиваются на мелкие отрезки (допуск 0.02 мм, как AppendRings в плагине),
     у каждого отрезка — номер сегмента и отрезок параметра t.
  2. Все пересечения отрезков (сетка по габаритам) режут подконтуры на куски.
  3. Кусок остаётся, если залита ровно одна сторона от него (обмотка слева/справа у середины
     самого длинного звена):
     это граница залитого и пустого. Внутренние линии наложений уходят, дырки остаются.
  4. Куски сшиваются в замкнутые контуры (залитое — справа, как у исходных внешних контуров
     текста CorelDRAW в координатах «y вверх»), в точках касания — по самому правому повороту.
  5. Мелкие петли и лоскуты (площадь < min_area или толщина < min_width) отбрасываются. Куски
     вывернутых петель (ласточкин хвост сдвинутого контура, «ушко» обводки: залитое у куска —
     только от его же подконтура и против его хода) не остаются уже на шаге 3.
  6. Каждый кусок собирается обратно из исходных сегментов, разрезанных точно в точках
     пересечения (де Кастельжо) — форма не меняется, без перетрассировки.
"""
import math

TOL = 0.02          # допуск разбиения кривых, мм


def _bez(p0, c1, c2, p3, t):
    mt = 1.0 - t
    a, b, c, d = mt*mt*mt, 3*mt*mt*t, 3*mt*t*t, t*t*t
    return (a*p0[0] + b*c1[0] + c*c2[0] + d*p3[0], a*p0[1] + b*c1[1] + c*c2[1] + d*p3[1])


def _split(seg, t0, t1):
    """Часть сегмента от t0 до t1 (t0 > t1 — в обратную сторону)."""
    if seg[0] == 'L':
        p0, p1 = seg[1], seg[2]
        a = (p0[0] + (p1[0]-p0[0])*t0, p0[1] + (p1[1]-p0[1])*t0)
        b = (p0[0] + (p1[0]-p0[0])*t1, p0[1] + (p1[1]-p0[1])*t1)
        return ('L', a, b)
    rev = t0 > t1
    if rev:
        t0, t1 = t1, t0
    p0, c1, c2, p3 = seg[1:]

    def cut_after(P, t):        # часть [t, 1]
        p0, c1, c2, p3 = P
        lerp = lambda a, b: (a[0] + (b[0]-a[0])*t, a[1] + (b[1]-a[1])*t)
        a, b, c = lerp(p0, c1), lerp(c1, c2), lerp(c2, p3)
        d, e = lerp(a, b), lerp(b, c)
        f = lerp(d, e)
        return (f, e, c, p3)

    def cut_before(P, t):       # часть [0, t]
        p0, c1, c2, p3 = P
        lerp = lambda a, b: (a[0] + (b[0]-a[0])*t, a[1] + (b[1]-a[1])*t)
        a, b, c = lerp(p0, c1), lerp(c1, c2), lerp(c2, p3)
        d, e = lerp(a, b), lerp(b, c)
        f = lerp(d, e)
        return (p0, a, d, f)
    P = (p0, c1, c2, p3)
    if t0 > 0.0:
        P = cut_after(P, t0)
    if t1 < 1.0:
        tt = (t1 - t0) / (1.0 - t0) if t0 < 1.0 else 1.0
        P = cut_before(P, tt)
    if rev:
        P = (P[3], P[2], P[1], P[0])
    return ('C',) + P


def _flatten(subs):
    """Отрезки: (номер подконтура, номер сегмента, t0, t1, a, b)."""
    E = []
    for si, segs in enumerate(subs):
        for gi, s in enumerate(segs):
            if s[0] == 'L':
                E.append((si, gi, 0.0, 1.0, s[1], s[2]))
                continue
            p0, c1, c2, p3 = s[1:]
            dd = max(math.hypot(p0[0]-2*c1[0]+c2[0], p0[1]-2*c1[1]+c2[1]),
                     math.hypot(c1[0]-2*c2[0]+p3[0], c1[1]-2*c2[1]+p3[1]))
            n = max(2, min(256, math.ceil(math.sqrt(0.75*dd/TOL))))
            prev = p0
            for i in range(1, n + 1):
                t = i / n
                q = p3 if i == n else _bez(p0, c1, c2, p3, t)
                E.append((si, gi, (i-1)/n, t, prev, q))
                prev = q
    return E


def _seg_x(a, b, c, d):
    ex, ey = b[0]-a[0], b[1]-a[1]; fx, fy = d[0]-c[0], d[1]-c[1]
    den = ex*fy - ey*fx
    if abs(den) < 1e-14:
        return None
    u = ((c[0]-a[0])*fy - (c[1]-a[1])*fx) / den
    v = ((c[0]-a[0])*ey - (c[1]-a[1])*ex) / den
    if -1e-12 <= u <= 1 + 1e-12 and -1e-12 <= v <= 1 + 1e-12:
        return u, v
    return None


def _grid(E, cell):
    G = {}
    for k, e in enumerate(E):
        a, b = e[4], e[5]
        for gx in range(int(math.floor(min(a[0], b[0])/cell)), int(math.floor(max(a[0], b[0])/cell)) + 1):
            for gy in range(int(math.floor(min(a[1], b[1])/cell)), int(math.floor(max(a[1], b[1])/cell)) + 1):
                G.setdefault((gx, gy), []).append(k)
    return G


def _winding(E, rows, h, y0, p, sp=-1):
    """Обмотка точки p: луч вправо, отрезки из полосы по y. Вторым — вклад подконтура sp."""
    r = int(math.floor((p[1] - y0) / h))
    w = 0; wo = 0
    for k in rows.get(r, ()):
        a, b = E[k][4], E[k][5]
        if a[1] <= p[1] < b[1] or b[1] <= p[1] < a[1]:
            x = a[0] + (p[1] - a[1]) * (b[0] - a[0]) / (b[1] - a[1])
            if x > p[0]:
                d = 1 if b[1] > a[1] else -1
                w += d
                if E[k][0] == sp:
                    wo += d
    return w, wo


def _area(pts):
    return 0.5*sum(pts[i][0]*pts[i-1][1]*-1 + pts[i-1][0]*pts[i][1] for i in range(len(pts))) * -1


def _emit(src, T0, T1, fwd):
    """Сегменты исходного подконтура src на отрезке параметра T0 -> T1 (T = номер сегмента + t,
    по кругу); fwd=False — тот же отрезок в обратную сторону."""
    N = len(src)
    if T1 <= T0:
        T1 += N
    out = []
    g = int(math.floor(T0))
    while g < T1 - 1e-12:
        a = max(T0 - g, 0.0); b = min(T1 - g, 1.0)
        if b - a > 1e-12:
            out.append(_split(src[g % N], a, b))
        g += 1
    if not fwd:
        out = [_split(sg, 1.0, 0.0) if sg[0] == 'C' else ('L', sg[2], sg[1]) for sg in out[::-1]]
    return out


def clean(subs, min_area=0.3, min_width=0.15):
    """Очищенный контур: список замкнутых подконтуров из сегментов исходных кривых."""
    E = _flatten(subs)
    if not E:
        return []
    L = [math.dist(e[4], e[5]) for e in E]
    cell = max(1.0, 8*sum(L)/len(L))
    G = _grid(E, cell)
    # 1. пересечения: параметр T = сегмент + t в точке пересечения (на обоих подконтурах)
    cutT = {si: [] for si in range(len(subs))}
    seen = set()
    for ks in G.values():
        for i in range(len(ks)):
            for j in range(i + 1, len(ks)):
                a, b = ks[i], ks[j]
                if (a, b) in seen:
                    continue
                seen.add((a, b))
                ea, eb = E[a], E[b]
                if ea[0] == eb[0] and (ea[5] == eb[4] or eb[5] == ea[4]):
                    continue                     # соседние отрезки делят конец
                r = _seg_x(ea[4], ea[5], eb[4], eb[5])
                if r is None:
                    continue
                X = (ea[4][0] + (ea[5][0]-ea[4][0])*r[0], ea[4][1] + (ea[5][1]-ea[4][1])*r[0])
                for e_, u in ((ea, r[0]), (eb, r[1])):
                    N = len(subs[e_[0]])
                    T = (e_[1] + e_[2] + (e_[3]-e_[2])*u) % N
                    cutT[e_[0]].append((T, X))
    # 2. куски подконтуров между разрезами: (подконтур, T0, T1, ломаная)
    pieces = []
    for si, segs in enumerate(subs):
        N = len(segs)
        V = [((e[1] + e[2]) % N, e[4]) for e in E if e[0] == si]      # вершины ломаной
        cs = sorted(cutT[si])
        # слить совпадающие разрезы
        cc = []
        for T, X in cs:
            if cc and abs(T - cc[-1][0]) < 1e-9:
                continue
            cc.append((T, X))
        if not cc:
            pieces.append((si, 0.0, float(N), [v[1] for v in V] + [V[0][1]], True))
            continue
        for c in range(len(cc)):
            T0, X0 = cc[c]; T1, X1 = cc[(c + 1) % len(cc)]
            if len(cc) == 1:
                T1 = T0 + N
            elif T1 <= T0:
                T1 += N
            P = [X0] + [v[1] for v in V if T0 < v[0] < T1 or T0 < v[0] + N < T1] + [X1]
            # вершины по порядку параметра
            if T1 > N:
                inner = [v for v in V if v[0] > T0] + [(v[0] + N, v[1]) for v in V if v[0] + N < T1]
                P = [X0] + [v[1] for v in sorted(inner)] + [X1]
            else:
                P = [X0] + [v[1] for v in sorted(v for v in V if T0 < v[0] < T1)] + [X1]
            pieces.append((si, T0, T1, P, False))
    # знак обмотки внутри подконтура (площадь со знаком; у петляющего — по большей части)
    sig = {}
    for si in range(len(subs)):
        V = [e[4] for e in E if e[0] == si]
        a = sum(V[i-1][0]*V[i][1] - V[i][0]*V[i-1][1] for i in range(len(V)))
        sig[si] = 1 if a > 0 else -1
    # 3. обмотка слева и справа от середины куска
    ys = [p[1] for e in E for p in (e[4], e[5])]
    y0 = min(ys); h = cell
    rows = {}
    for k, e in enumerate(E):
        for r in range(int(math.floor((min(e[4][1], e[5][1]) - y0)/h)), int(math.floor((max(e[4][1], e[5][1]) - y0)/h)) + 1):
            rows.setdefault(r, []).append(k)
    kept = []
    for si, T0, T1, P, whole in pieces:
        tot = sum(math.dist(P[i], P[i+1]) for i in range(len(P)-1))
        if tot < 1e-9:
            continue
        # проба — у середины самого длинного звена (не в вершине: у острого зубца точки по обе
        # стороны от вершины обе оказались бы снаружи)
        im = max(range(len(P)-1), key=lambda i: math.dist(P[i], P[i+1]))
        d = math.dist(P[im], P[im+1])
        if d < 1e-12:
            continue
        m = ((P[im][0] + P[im+1][0])/2, (P[im][1] + P[im+1][1])/2)
        tx, ty = (P[im+1][0]-P[im][0])/d, (P[im+1][1]-P[im][1])/d
        eps = 1e-4
        wl, ol = _winding(E, rows, h, y0, (m[0] - ty*eps, m[1] + tx*eps), si)
        wr, orr = _winding(E, rows, h, y0, (m[0] + ty*eps, m[1] - tx*eps), si)
        if (wl != 0) == (wr != 0):
            continue                     # внутри залитого или снаружи — не граница
        fwd = wr != 0                    # залитое справа
        wf, wo = (wr, orr) if fwd else (wl, ol)
        # вывернутая петля (ласточкин хвост сдвинутого контура, «ушко» обводки): залитое здесь —
        # только от самого подконтура и против его хода; у дырки залитое — от внешнего контура
        if wf == wo and (1 if wf > 0 else -1) == -sig[si]:
            continue
        kept.append((si, T0, T1, P if fwd else P[::-1], fwd, whole))
    # 4. сшивка кусков в контуры
    key = lambda p: (round(p[0], 5), round(p[1], 5))
    loops = []
    starts = {}
    for n, k in enumerate(kept):
        if k[5]:
            loops.append([n])
        else:
            starts.setdefault(key(k[3][0]), []).append(n)
    used = set()
    for n, k in enumerate(kept):
        if k[5] or n in used:
            continue
        loop = []; cur = n
        while cur is not None and cur not in used:
            used.add(cur); loop.append(cur)
            P = kept[cur][3]; end = P[-1]
            cand = [c for c in starts.get(key(end), []) if c not in used]
            if not cand:
                break
            if len(cand) > 1:
                # касание: самый правый поворот (залитое справа — обход вплотную к нему)
                a = P[-2]
                din = math.atan2(end[1]-a[1], end[0]-a[0])

                def turn(c):
                    b = kept[c][3][1]
                    dout = math.atan2(b[1]-end[1], b[0]-end[0])
                    return (dout - din + math.pi) % (2*math.pi) - math.pi
                cand.sort(key=turn)
            cur = cand[0]
        loops.append(loop)
    # 5. мелкие петли и лоскуты; 6. сегменты исходных кривых
    result = []
    for loop in loops:
        P = [p for n in loop for p in kept[n][3]]
        A = abs(_area(P))
        per = sum(math.dist(P[i], P[i-1]) for i in range(len(P)))
        if A < min_area or (per > 0 and 2*A/per < min_width):
            continue
        segs = []
        for n in loop:
            si, T0, T1, _, fwd, whole = kept[n]
            src = subs[si]
            segs += _emit(src, T0, T1, True) if fwd else _emit(src, T0, T1, False)
        if segs:
            _snap(segs)
            result.append(segs)
    return result


def _snap(segs):
    """Стыки кусков разных подконтуров: точка пересечения найдена по ломаной, а кривые разрезаны
    точно — концы расходятся до допуска разбиения. Оба конца — в середину (с ручками)."""
    n = len(segs)
    for i in range(n):
        a = segs[i-1]; b = segs[i]
        pa, pb = a[-1], b[1]
        if pa == pb:
            continue
        m = ((pa[0]+pb[0])/2, (pa[1]+pb[1])/2)
        da = (m[0]-pa[0], m[1]-pa[1]); db = (m[0]-pb[0], m[1]-pb[1])
        if a[0] == 'C':
            a = ('C', a[1], a[2], (a[3][0]+da[0], a[3][1]+da[1]), m)
        else:
            a = ('L', a[1], m)
        if b[0] == 'C':
            b = ('C', m, (b[2][0]+db[0], b[2][1]+db[1]), b[3], b[4])
        else:
            b = ('L', m, b[2])
        segs[i-1] = a; segs[i] = b
