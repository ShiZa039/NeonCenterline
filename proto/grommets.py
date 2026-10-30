"""Прототип расстановки люверсов (то же делает NeonTools.PlaceGrommets в VBA).

Контур баннера (замкнутый подпуть из сегментов Безье) -> ломаная -> смещение внутрь на отступ
от края до центра люверса -> люверсы: в каждом углу обязательно, на сторонах между углами —
равномерно, расстояние между соседними не больше заданного шага. Без углов (круг, овал) —
равномерно по всему контуру, первый — в верхней точке.

Угол — узел контура, где касательная ломается больше чем на CORNER_DEG градусов, или середина
скругления угла: участка с радиусом кривизны меньше четверти меньшей стороны габарита, который
поворачивает не меньше чем на 45 град (у круга и овала без прямых участков углов нет — люверсы
идут равномерно по всему контуру).

Усиление углов (reinforce = r, мм): у каждого угла ещё по люверсу на обе соседние стороны на
расстоянии r от углового. Точка ближе r/2 к уже поставленной не ставится (короткая сторона)."""
import math

CORNER_DEG = 30.0


def _bz(p, t):
    mt = 1 - t
    return (mt**3*p[0][0] + 3*mt*mt*t*p[1][0] + 3*mt*t*t*p[2][0] + t**3*p[3][0],
            mt**3*p[0][1] + 3*mt*mt*t*p[1][1] + 3*mt*t*t*p[2][1] + t**3*p[3][1])


def _norm(dx, dy):
    L = math.hypot(dx, dy)
    return (dx/L, dy/L) if L > 1e-12 else (0.0, 0.0)


def flatten(segs, step):
    """segs: [(p0, p1, p2, p3, is_line)] замкнутого подпути. -> (точки, флаг угла у каждой)."""
    pts = []; corner = []
    n = len(segs)
    for k, (p0, p1, p2, p3, ln) in enumerate(segs):
        # излом в начальном узле: касательная входа (конец предыдущего) и выхода
        q0, q1, q2, q3, pl = segs[k-1]
        ti = _norm(q3[0]-q0[0], q3[1]-q0[1]) if pl else _tan_end(q0, q1, q2, q3)
        to = _norm(p3[0]-p0[0], p3[1]-p0[1]) if ln else _tan_start(p0, p1, p2, p3)
        c = ti[0]*to[0] + ti[1]*to[1]
        is_c = c < math.cos(math.radians(CORNER_DEG))
        if ln:
            L = math.dist(p0, p3); m = max(2, int(math.ceil(L/step)))
            ps = [(p0[0] + (p3[0]-p0[0])*i/m, p0[1] + (p3[1]-p0[1])*i/m) for i in range(m)]
        else:
            Lc = math.dist(p0, p1) + math.dist(p1, p2) + math.dist(p2, p3)
            m = max(2, int(math.ceil(Lc/step)))
            ps = [_bz((p0, p1, p2, p3), i/m) for i in range(m)]
        for i, p in enumerate(ps):
            if pts and math.dist(pts[-1], p) < 1e-6:
                corner[-1] = corner[-1] or (is_c and i == 0)
                continue
            pts.append(p); corner.append(is_c and i == 0)
    if len(pts) > 1 and math.dist(pts[0], pts[-1]) < 1e-6:
        corner[0] = corner[0] or corner[-1]
        pts.pop(); corner.pop()
    return pts, corner


def _tan_start(p0, p1, p2, p3):
    for q in (p1, p2, p3):
        if math.dist(p0, q) > 1e-9:
            return _norm(q[0]-p0[0], q[1]-p0[1])
    return (0.0, 0.0)


def _tan_end(p0, p1, p2, p3):
    for q in (p2, p1, p0):
        if math.dist(q, p3) > 1e-9:
            return _norm(p3[0]-q[0], p3[1]-q[1])
    return (0.0, 0.0)


def area(pts):
    return 0.5*sum(pts[i][0]*pts[(i+1) % len(pts)][1] - pts[(i+1) % len(pts)][0]*pts[i][1] for i in range(len(pts)))


def offset_in(pts, d):
    """Смещение замкнутой ломаной (против часовой) внутрь на d: у каждой вершины — по биссектрисе
    нормалей соседних звеньев, длина d / cos(половины угла) (острые углы — не дальше 4d)."""
    n = len(pts)
    nrm = []
    for i in range(n):
        a = pts[i]; b = pts[(i+1) % n]
        t = _norm(b[0]-a[0], b[1]-a[1])
        nrm.append((-t[1], t[0]))                        # влево = внутрь у обхода против часовой
    out = []
    for i in range(n):
        n0 = nrm[i-1]; n1 = nrm[i]
        m = _norm(n0[0]+n1[0], n0[1]+n1[1])
        c = m[0]*n1[0] + m[1]*n1[1]
        L = d/c if c > 0.25 else 4*d
        out.append((pts[i][0] + m[0]*L, pts[i][1] + m[1]*L))
    return out


def soft_corners(pts, cor):
    """Скругления углов: подряд идущие вершины с радиусом кривизны < 0.25 меньшей стороны
    габарита (без острых углов внутри), поворот участка >= 45 град; почти прямых участков
    (радиус больше 20 таких) — не меньше четверти периметра. Середина участка (по длине) становится углом."""
    n = len(pts)
    xs = [p[0] for p in pts]; ys = [p[1] for p in pts]
    Rc = 0.25*min(max(xs)-min(xs), max(ys)-min(ys))
    seg = [math.dist(pts[i], pts[(i+1) % n]) for i in range(n)]
    turn = []; curved = []
    for i in range(n):
        a = pts[i-1]; b = pts[i]; c = pts[(i+1) % n]
        t0 = _norm(b[0]-a[0], b[1]-a[1]); t1 = _norm(c[0]-b[0], c[1]-b[1])
        ang = math.atan2(t0[0]*t1[1] - t0[1]*t1[0], t0[0]*t1[0] + t0[1]*t1[1])
        L = 0.5*(seg[i-1] + seg[i])
        turn.append(ang)
        curved.append(not cor[i] and abs(ang) > 1e-9 and L/abs(ang) < Rc)
    total = sum(seg)
    # прямых участков (радиус больше 20 Rc) — не меньше четверти периметра: у овала их нет
    straight = sum(seg[i] for i in range(n) if not cor[i] and (abs(turn[i]) < 1e-9 or
                   0.5*(seg[i-1] + seg[i])/abs(turn[i]) > 20*Rc))
    if straight < 0.25*total:
        return
    # участки подряд идущих изогнутых вершин (по кругу); начать с прямой вершины
    st = next(i for i in range(n) if not curved[i])
    i = 0
    while i < n:
        k = (st + i) % n
        if not curved[k]:
            i += 1; continue
        run = []
        while i < n and curved[(st + i) % n]:
            run.append((st + i) % n); i += 1
        tot = sum(turn[k] for k in run)
        if abs(tot) < math.radians(45):
            continue
        L = sum(seg[k] for k in run[:-1]); acc = 0.0; mid = run[0]
        for k in run[:-1]:
            if acc + seg[k] >= L/2:
                mid = k if L/2 - acc < acc + seg[k] - L/2 else (k+1) % n
                break
            acc += seg[k]
        cor[mid] = True


def place(segs, edge, step, flat=None, reinforce=None):
    """Центры люверсов одного замкнутого подпути. None — отступ больше, чем позволяет контур."""
    if flat is None:
        flat = max(0.2, min(1.0, edge/4))
    pts, cor = flatten(segs, flat)
    if len(pts) < 3:
        return None
    if area(pts) < 0:
        pts = pts[::-1]; cor = cor[::-1]
    a0 = area(pts)
    soft_corners(pts, cor)
    O = offset_in(pts, edge)
    # у выпуклого угла смещённые точки сторон ближе отступа к углу уходят за смещённый угол
    # (петля «ласточкин хвост»): выбрасываются — после угла всё, что не впереди него по
    # выходящей стороне, до угла — всё, что не позади по входящей
    n = len(O); drop = [False]*n
    for c in range(n):
        if not cor[c]:
            continue
        ti = _norm(pts[c][0]-pts[c-1][0], pts[c][1]-pts[c-1][1])
        to = _norm(pts[(c+1) % n][0]-pts[c][0], pts[(c+1) % n][1]-pts[c][1])
        k = (c+1) % n
        while k != c and not cor[k] and (O[k][0]-O[c][0])*to[0] + (O[k][1]-O[c][1])*to[1] <= 1e-9:
            drop[k] = True; k = (k+1) % n
        k = (c-1) % n
        while k != c and not cor[k] and (O[k][0]-O[c][0])*ti[0] + (O[k][1]-O[c][1])*ti[1] >= -1e-9:
            drop[k] = True; k = (k-1) % n
    keep = [i for i in range(n) if not drop[i]]
    O = [O[i] for i in keep]; cor = [cor[i] for i in keep]
    if len(O) < 3 or area(O) <= 0.02*a0:
        return None
    n = len(O)
    cum = [0.0]
    for i in range(n):
        cum.append(cum[-1] + math.dist(O[i], O[(i+1) % n]))
    total = cum[-1]

    def at(i0, s):
        """Точка на расстоянии s по обходу от вершины i0."""
        s0 = cum[i0] + s
        while s0 >= total:
            s0 -= total
        lo, hi = 0, n
        while hi - lo > 1:
            mid = (lo + hi)//2
            if cum[mid] <= s0:
                lo = mid
            else:
                hi = mid
        i = lo; L = cum[i+1] - cum[i]
        t = 0.0 if L < 1e-12 else (s0 - cum[i])/L
        a = O[i]; b = O[(i+1) % n]
        return (a[0] + (b[0]-a[0])*t, a[1] + (b[1]-a[1])*t)

    cs = [i for i in range(n) if cor[i]]
    res = []
    if not cs:
        top = max(range(n), key=lambda i: (O[i][1], -O[i][0]))
        k = max(1, int(math.ceil(total/step - 1e-9)))
        return [at(top, total*r/k) for r in range(k)]
    for j, c in enumerate(cs):
        c2 = cs[(j+1) % len(cs)]
        Lr = cum[c2] - cum[c] if c2 > c else total - cum[c] + cum[c2]
        k = max(1, int(math.ceil(Lr/step - 1e-9)))
        res += [at(c, Lr*r/k) for r in range(k)]
    if reinforce:
        for c in cs:
            for s in (reinforce, total - reinforce):
                p = at(c, s)
                if all(math.dist(p, q) >= 0.5*reinforce for q in res):
                    res.append(p)
    return res


if __name__ == '__main__':
    import matplotlib; matplotlib.use('Agg'); import matplotlib.pyplot as plt

    def rect(x, y, w, h):
        P = [(x, y), (x+w, y), (x+w, y+h), (x, y+h)]
        return [(P[i], None, None, P[(i+1) % 4], True) for i in range(4)]

    def fix_lines(segs):
        return [(a, b if b else a, c if c else d, d, ln) for a, b, c, d, ln in segs]

    def circle(cx, cy, r):
        k = 0.5523; segs = []
        for q in range(4):
            a0 = q*math.pi/2; a1 = a0 + math.pi/2
            p0 = (cx + r*math.cos(a0), cy + r*math.sin(a0)); p3 = (cx + r*math.cos(a1), cy + r*math.sin(a1))
            p1 = (p0[0] - k*r*math.sin(a0), p0[1] + k*r*math.cos(a0))
            p2 = (p3[0] + k*r*math.sin(a1), p3[1] - k*r*math.cos(a1))
            segs.append((p0, p1, p2, p3, False))
        return segs

    def rrect(x, y, w, h, r):
        k = 0.5523; s = []
        A = [(x+r, y), (x+w-r, y)]; s.append((A[0], A[0], A[1], A[1], True))
        s.append(((x+w-r, y), (x+w-r+k*r, y), (x+w, y+r-k*r), (x+w, y+r), False))
        s.append(((x+w, y+r), (x+w, y+r), (x+w, y+h-r), (x+w, y+h-r), True))
        s.append(((x+w, y+h-r), (x+w, y+h-r+k*r), (x+w-r+k*r, y+h), (x+w-r, y+h), False))
        s.append(((x+w-r, y+h), (x+w-r, y+h), (x+r, y+h), (x+r, y+h), True))
        s.append(((x+r, y+h), (x+r-k*r, y+h), (x, y+h-r+k*r), (x, y+h-r), False))
        s.append(((x, y+h-r), (x, y+h-r), (x, y+r), (x, y+r), True))
        s.append(((x, y+r), (x, y+r-k*r), (x+r-k*r, y), (x+r, y), False))
        return s

    def lshape():
        P = [(0, 0), (1200, 0), (1200, 400), (500, 400), (500, 900), (0, 900)]
        return [(P[i], P[i], P[(i+1) % 6], P[(i+1) % 6], True) for i in range(6)]

    cases = [('3000×1000, отступ 20, шаг 300', fix_lines(rect(0, 0, 3000, 1000)), 20, 300),
             ('1000×1000, скругл. 100', rrect(0, 0, 1000, 1000, 100), 20, 300),
             ('круг Ø800', circle(400, 400, 400), 25, 250),
             ('овал 1500×600', [((p0[0]*1.875, p0[1]*0.75), (p1[0]*1.875, p1[1]*0.75), (p2[0]*1.875, p2[1]*0.75), (p3[0]*1.875, p3[1]*0.75), l) for p0, p1, p2, p3, l in circle(400, 400, 400)], 20, 300),
             ('Г-образный', lshape(), 20, 300)]
    fig, axs = plt.subplots(1, 5, figsize=(27, 6))
    for ax, (t, segs, e, st) in zip(axs, cases):
        pts, _ = flatten(segs, 2)
        ax.plot([p[0] for p in pts + pts[:1]], [p[1] for p in pts + pts[:1]], 'k-', lw=1)
        g = place(segs, e, st, reinforce=50)
        ax.plot([p[0] for p in g], [p[1] for p in g], 'o', color='#e3007a', ms=6)
        ax.set_title(f'{t}: {len(g)} шт. (с усилением углов 50)')
        ax.set_aspect('equal')
    plt.tight_layout(); plt.savefig('/tmp/grommets_proto.png', dpi=55)
    for t, segs, e, st in cases:
        print(t, len(place(segs, e, st)), len(place(segs, e, st, reinforce=50)))
