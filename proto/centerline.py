"""Прототип алгоритма центральной линии (написан "в стиле VBA": простые списки,
без numpy в самом алгоритме), чтобы порт в CorelDRAW был 1:1.

Вход:  список колец (замкнутые полилинии) одной фигуры, мм.
Выход: список путей; путь = список сегментов ('L', p0, p1) | ('C', p0, c1, c2, p3).
"""
import math

P = dict(
    alpha=0.80,      # хорды короче alpha*W выбрасываются (усики, концы)
    prune=1.0,       # ветки с вылетом (длина+до края) < prune*W удаляются
    tol=0.25,        # допуск аппроксимации Безье, мм
    corner=55.0,     # угол излома (град), при котором ставим острый узел
    merge_ang=40.0,  # макс. излом для сквозного прохода через узел ветвления
    tool=8.0,        # диаметр фрезы, мм: конец линии на tool/2 от края
    smooth_it=12,
    jzone=1.0,
    widen=1.05,     # у развилок узлы, где штрих шире widen*W, тоже ненадёжны
    jmerge=1.0,     # развилки ближе jmerge*W (по графу) сливаются в одну      # радиус зоны развилки (в W), где точки CAT ненадёжны
    end_trim=0.95,  # у свободных концов отрезаем участок, где хорда < end_trim*W
)

# ---------------------------------------------------------------- geometry

def ring_area(r):
    a = 0.0
    n = len(r)
    for i in range(n):
        x1, y1 = r[i]; x2, y2 = r[(i + 1) % n]
        a += x1 * y2 - x2 * y1
    return a / 2


def ring_len(r):
    return sum(math.dist(r[i], r[(i + 1) % len(r)]) for i in range(len(r)))


def resample(r, h):
    L = ring_len(r)
    n = max(8, int(math.ceil(L / h)))
    step = L / n
    out = []
    i = 0; acc = 0.0; t = 0.0
    m = len(r)
    seglen = [math.dist(r[k], r[(k + 1) % m]) for k in range(m)]
    k = 0; pos = 0.0
    for j in range(n):
        target = j * step
        while k < m - 1 and pos + seglen[k] < target:
            pos += seglen[k]; k += 1
        s = seglen[k]
        u = 0 if s == 0 else (target - pos) / s
        a = r[k]; b = r[(k + 1) % m]
        out.append((a[0] + (b[0] - a[0]) * u, a[1] + (b[1] - a[1]) * u))
    return out


def inside(rings, x, y):
    c = False
    for r in rings:
        n = len(r)
        j = n - 1
        for i in range(n):
            xi, yi = r[i]; xj, yj = r[j]
            if (yi > y) != (yj > y):
                if x < (xj - xi) * (y - yi) / (yj - yi) + xi:
                    c = not c
            j = i
    return c

# ---------------------------------------------------------------- delaunay (Bourke)

def delaunay(px, py):
    n = len(px)
    order = sorted(range(n), key=lambda i: px[i])
    xs = [px[i] for i in order] ; ys = [py[i] for i in order]
    xmin, xmax, ymin, ymax = min(xs), max(xs), min(ys), max(ys)
    dm = max(xmax - xmin, ymax - ymin) * 20
    xm = (xmin + xmax) / 2; ym = (ymin + ymax) / 2
    xs += [xm - dm, xm, xm + dm]; ys += [ym - dm, ym + dm, ym - dm]
    tri = [[n, n + 1, n + 2]]
    cc = [None]; done = [False]

    def circ(t):
        a, b, c = t
        ax, ay, bx, by, cx, cy = xs[a], ys[a], xs[b], ys[b], xs[c], ys[c]
        d = 2 * (ax * (by - cy) + bx * (cy - ay) + cx * (ay - by))
        if abs(d) < 1e-12:
            return (ax, ay, 1e30)
        ux = ((ax*ax+ay*ay)*(by-cy)+(bx*bx+by*by)*(cy-ay)+(cx*cx+cy*cy)*(ay-by))/d
        uy = ((ax*ax+ay*ay)*(cx-bx)+(bx*bx+by*by)*(ax-cx)+(cx*cx+cy*cy)*(bx-ax))/d
        return (ux, uy, (ax-ux)**2 + (ay-uy)**2)
    cc[0] = circ(tri[0])
    for i in range(n):
        x, y = xs[i], ys[i]
        edges = {}
        k = 0
        while k < len(tri):
            if done[k]:
                k += 1; continue
            ux, uy, r2 = cc[k]
            dx = x - ux
            if dx > 0 and dx * dx > r2:
                done[k] = True; k += 1; continue
            if dx * dx + (y - uy) ** 2 <= r2:
                a, b, c = tri[k]
                for e in ((a, b), (b, c), (c, a)):
                    key = (min(e), max(e))
                    edges[key] = edges.get(key, 0) + 1
                tri[k] = tri[-1]; cc[k] = cc[-1]; done[k] = done[-1]
                tri.pop(); cc.pop(); done.pop()
                continue
            k += 1
        for (a, b), cnt in edges.items():
            if cnt == 1:
                t = [a, b, i]
                tri.append(t); cc.append(circ(t)); done.append(False)
    res = []
    for a, b, c in tri:
        if a < n and b < n and c < n:
            res.append((order[a], order[b], order[c]))
    return res

# ---------------------------------------------------------------- bezier fit (Schneider)

def _bez(p, t):
    mt = 1 - t
    return (mt**3*p[0][0]+3*mt*mt*t*p[1][0]+3*mt*t*t*p[2][0]+t**3*p[3][0],
            mt**3*p[0][1]+3*mt*mt*t*p[1][1]+3*mt*t*t*p[2][1]+t**3*p[3][1])

def _bezd(p, t):
    mt = 1 - t
    return (3*mt*mt*(p[1][0]-p[0][0])+6*mt*t*(p[2][0]-p[1][0])+3*t*t*(p[3][0]-p[2][0]),
            3*mt*mt*(p[1][1]-p[0][1])+6*mt*t*(p[2][1]-p[1][1])+3*t*t*(p[3][1]-p[2][1]))

def _bezdd(p, t):
    return (6*(1-t)*(p[2][0]-2*p[1][0]+p[0][0])+6*t*(p[3][0]-2*p[2][0]+p[1][0]),
            6*(1-t)*(p[2][1]-2*p[1][1]+p[0][1])+6*t*(p[3][1]-2*p[2][1]+p[1][1]))

def _norm(v):
    l = math.hypot(*v)
    return (v[0]/l, v[1]/l) if l > 1e-12 else (0.0, 0.0)

def _gen(pts, u, t1, t2):
    p0, p3 = pts[0], pts[-1]
    c00 = c01 = c11 = x0 = x1 = 0.0
    for i, t in enumerate(u):
        mt = 1 - t
        b1 = 3*mt*mt*t; b2 = 3*mt*t*t
        a1 = (t1[0]*b1, t1[1]*b1); a2 = (t2[0]*b2, t2[1]*b2)
        c00 += a1[0]*a1[0]+a1[1]*a1[1]
        c01 += a1[0]*a2[0]+a1[1]*a2[1]
        c11 += a2[0]*a2[0]+a2[1]*a2[1]
        b0 = mt**3; b3 = t**3
        tx = pts[i][0]-(p0[0]*(b0+b1)+p3[0]*(b2+b3))
        ty = pts[i][1]-(p0[1]*(b0+b1)+p3[1]*(b2+b3))
        x0 += a1[0]*tx+a1[1]*ty
        x1 += a2[0]*tx+a2[1]*ty
    det = c00*c11-c01*c01
    seg = math.dist(p0, p3)
    if abs(det) > 1e-12:
        al1 = (x0*c11-x1*c01)/det
        al2 = (c00*x1-c01*x0)/det
    else:
        al1 = al2 = -1
    if al1 < seg*1e-3 or al2 < seg*1e-3 or al1 > seg*3 or al2 > seg*3:
        al1 = al2 = seg/3
    return [p0, (p0[0]+t1[0]*al1, p0[1]+t1[1]*al1),
            (p3[0]+t2[0]*al2, p3[1]+t2[1]*al2), p3]

def _maxerr(pts, bez, u):
    m = 0.0; idx = len(pts)//2
    for i in range(1, len(pts)-1):
        q = _bez(bez, u[i])
        d = (q[0]-pts[i][0])**2+(q[1]-pts[i][1])**2
        if d > m:
            m = d; idx = i
    return math.sqrt(m), idx

def _reparam(pts, bez, u):
    out = []
    for i, t in enumerate(u):
        q = _bez(bez, t); d1 = _bezd(bez, t); d2 = _bezdd(bez, t)
        num = (q[0]-pts[i][0])*d1[0]+(q[1]-pts[i][1])*d1[1]
        den = d1[0]**2+d1[1]**2+(q[0]-pts[i][0])*d2[0]+(q[1]-pts[i][1])*d2[1]
        out.append(min(1, max(0, t-num/den)) if abs(den) > 1e-12 else t)
    return out

def _line_dev(pts):
    a, b = pts[0], pts[-1]
    L = math.dist(a, b)
    if L < 1e-9:
        return 1e9
    return max(abs((b[0]-a[0])*(a[1]-p[1])-(a[0]-p[0])*(b[1]-a[1]))/L for p in pts)

SMOOTH_ARC = False   # не считать углом крутую гладкую дугу (режим образца)
LMIN_K = 0.8   # прямой участок >= LMIN_K*W в начале/конце куска — отдельным отрезком


def _straight_run(pts, tol, lmin):
    """Самый длинный прямой (dev <= tol) участок от pts[0]: индекс последней точки,
    если длина >= lmin и это не весь участок; иначе None."""
    n = len(pts); best = None
    for k in range(2, n-1):
        if _line_dev(pts[:k+1]) > tol:
            break
        if math.dist(pts[0], pts[k]) >= lmin:
            best = k
    return best


def fit(pts, t1, t2, tol, out, lmin=None):
    # прямая — только если участок действительно прямой (0.4 допуска), иначе
    # пологие дуги рассыпаются на цепочку коротких отрезков
    if len(pts) == 2 or _line_dev(pts) <= 0.4*tol:
        out.append(('L', pts[0], pts[-1])); return
    if lmin is not None and len(pts) > 3:
        # длинный прямой участок в начале и/или конце куска — отдельным отрезком, остаток
        # по касательной к нему (низ чаши P, перекладина B: прямая + дуга, а не одна
        # кривая). Только на уровне куска: внутри остатка обычное деление, иначе короткий
        # изгиб между двумя прямыми рассыпается на отрезки с изломами.
        k = _straight_run(pts, 0.4*tol, lmin)
        a = 0
        if k is not None:
            out.append(('L', pts[0], pts[k]))
            t1 = _norm((pts[k][0]-pts[0][0], pts[k][1]-pts[0][1])); a = k
        k = _straight_run(pts[a:][::-1], 0.4*tol, lmin)
        m = len(pts)-1
        if k is not None:
            m = len(pts)-1-k
            t2 = _norm((pts[m][0]-pts[-1][0], pts[m][1]-pts[-1][1]))
        if a > 0 or m < len(pts)-1:
            if m - a >= 1:
                fit(pts[a:m+1], t1, t2, tol, out, None)
            if m < len(pts)-1:
                out.append(('L', pts[m], pts[-1]))
            return
    d = [0.0]
    for i in range(1, len(pts)):
        d.append(d[-1]+math.dist(pts[i-1], pts[i]))
    u = [x/d[-1] for x in d]
    bez = _gen(pts, u, t1, t2)
    err, idx = _maxerr(pts, bez, u)
    if err <= tol:
        out.append(('C',)+tuple(bez)); return
    if err <= tol*4:
        for _ in range(6):
            u = _reparam(pts, bez, u)
            bez = _gen(pts, u, t1, t2)
            err, idx = _maxerr(pts, bez, u)
            if err <= tol:
                out.append(('C',)+tuple(bez)); return
    idx = max(1, min(len(pts)-2, idx))
    tc = _norm((pts[idx-1][0]-pts[idx+1][0], pts[idx-1][1]-pts[idx+1][1]))
    fit(pts[:idx+1], t1, tc, tol, out, lmin)
    fit(pts[idx:], (-tc[0], -tc[1]), t2, tol, out, lmin)

# ---------------------------------------------------------------- main

def centerline(rings_in, prm=P, debug=None):
    A = abs(sum(ring_area(r) if i == 0 else 0 for i, r in enumerate(rings_in)))
    # ширина штриха: W ~ 2*|площадь| / периметр
    area = 0.0
    for r in rings_in:
        area += ring_area(r)
    # знак колец может быть любым — считаем площадь заливки чётно-нечётно через
    # сумму |площадей| с учётом вложенности: проще — внешние минус дырки
    areas = sorted((abs(ring_area(r)) for r in rings_in), reverse=True)
    per = sum(ring_len(r) for r in rings_in)
    fillA = _evenodd_area(rings_in)
    W = 2 * fillA / per
    h = W / 5.0
    rings = [resample(r, h) for r in rings_in]
    # точки
    px, py, ring_id, idx_in_ring, ring_start, ring_n = [], [], [], [], [], []
    for ri, r in enumerate(rings):
        ring_start.append(len(px)); ring_n.append(len(r))
        for k, (x, y) in enumerate(r):
            px.append(x); py.append(y); ring_id.append(ri); idx_in_ring.append(k)
    tris = delaunay(px, py)

    def is_boundary(a, b):
        if ring_id[a] != ring_id[b]:
            return False
        n = ring_n[ring_id[a]]
        d = abs(idx_in_ring[a]-idx_in_ring[b])
        return d == 1 or d == n-1

    T = []
    for a, b, c in tris:
        cx = (px[a]+px[b]+px[c])/3; cy = (py[a]+py[b]+py[c])/3
        if inside(rings, cx, cy):
            T.append((a, b, c))

    # узлы графа: середины внутренних рёбер (ключ - ребро) и центры развилок
    node_xy = []; node_r = []; node_kind = []
    ekey = {}
    def enode(a, b):
        k = (min(a, b), max(a, b))
        if k not in ekey:
            ekey[k] = len(node_xy)
            node_xy.append(((px[a]+px[b])/2, (py[a]+py[b])/2))
            node_r.append(math.dist((px[a], py[a]), (px[b], py[b])))
            node_kind.append(0)
        return ekey[k]
    adj = {}
    def link(i, j):
        adj.setdefault(i, set()).add(j); adj.setdefault(j, set()).add(i)

    for a, b, c in T:
        ie = [e for e in ((a, b), (b, c), (c, a)) if not is_boundary(*e)]
        if len(ie) == 2:
            link(enode(*ie[0]), enode(*ie[1]))
        elif len(ie) == 3:
            ns = [enode(*e) for e in ie]
            # центр развилки: среднее середин хорд
            jx = sum(node_xy[n][0] for n in ns)/3; jy = sum(node_xy[n][1] for n in ns)/3
            j = len(node_xy); node_xy.append((jx, jy)); node_r.append(1e9); node_kind.append(1)
            for n in ns:
                link(j, n)
        # len==1: концевой треугольник - пропускаем

    # 1) фильтр по хорде
    alive = [True]*len(node_xy)
    for i in range(len(node_xy)):
        if node_kind[i] == 0 and node_r[i] < prm['alpha']*W:
            alive[i] = False
    for i in list(adj):
        if not alive[i]:
            for j in adj[i]:
                adj[j].discard(i)
            del adj[i]
    for i in list(adj):
        if not adj[i]:
            del adj[i]
    # развилки, у которых меньше 3 соседей — либо транзит, либо мусор
    _dissolve_junctions(adj, node_kind)

    # 2) обрезка коротких висячих веток (итеративно)
    for _ in range(4):
        changed = _prune(adj, node_xy, prm['prune']*W, rings_in)
        _dissolve_junctions(adj, node_kind)
        if not changed:
            break
    _merge_junctions(adj, node_xy, node_kind, node_r, prm['jmerge']*W)
    _merge_bridges(adj, node_xy, node_kind, node_r, 2.5*W, prm['merge_ang'], 1.5*W)
    # 3) мелкие компоненты
    _drop_small(adj, node_xy, W)

    # 4) цепочки
    chains, passthru = _chains(adj, node_xy, prm['merge_ang'], 1.5*W)
    jzone = prm['jzone'] * W
    jun = [i for i in adj if len(adj[i]) >= 3]
    # какая цепочка проходит сквозь развилку
    thru_of = {}
    for ci, (ch, closed) in enumerate(chains):
        for i in ch[1:-1] if not closed else ch:
            if i in passthru and passthru[i]:
                thru_of[i] = ci
    data = []   # [pts, closed, stub_start(junction id|None), stub_end]
    for ci, (ch, closed) in enumerate(chains):
        stub = [None, None]
        free = [False, False]
        if not closed:
            for end in (0, 1):
                e = ch[0] if end == 0 else ch[-1]
                if len(adj[e]) == 1:
                    free[end] = True
                elif len(adj[e]) >= 3:
                    stub[end] = e
            for end in (0, 1):
                if not free[end]:
                    continue
                ids = ch if end == 1 else ch[::-1]
                k = len(ids) - 1
                while k > 1 and node_kind[ids[k]] == 0 and node_r[ids[k]] < prm['end_trim']*W:
                    k -= 1
                ids = ids[:k+1]
                ch = ids if end == 1 else ids[::-1]
        # развилки, которых касается цепочка
        touched = [i for i in ch if i in adj and len(adj[i]) >= 3]
        # точки в зоне развилок ненадёжны: выкидываем их, а дыру внутри
        # цепочки заполняем гладкой кривой Эрмита по касательным соседей
        flags = []
        for i in ch:
            bad = len(adj.get(i, ())) >= 3
            if not bad:
                for j in touched:
                    dj = math.dist(node_xy[i], node_xy[j])
                    if dj < jzone or (dj < 2.5*W and node_r[i] > prm['widen']*W):
                        bad = True; break
            flags.append(bad)
        if closed and all(flags):
            continue
        pts = []
        n = len(ch); k = 0
        if closed and flags[0]:
            # повернём кольцо так, чтобы начиналось с хорошей точки
            r0 = flags.index(False)
            ch = ch[r0:] + ch[:r0]; flags = flags[r0:] + flags[:r0]
        while k < n:
            if not flags[k]:
                pts.append(node_xy[ch[k]]); k += 1; continue
            k2 = k
            while k2 < n and flags[k2]:
                k2 += 1
            jc = [ch[q] for q in range(k, k2) if len(adj.get(ch[q], ())) >= 4]
            jc = node_xy[jc[0]] if jc else None
            if pts and k2 < n:          # дыра в середине -> Эрмит (через центр развилки)
                pts.extend(_hermite_fill(pts, [node_xy[i] for i in ch[k2:]], W/5, jc))
            elif pts and closed:
                pts.extend(_hermite_fill(pts, pts[:], W/5, jc))
            k = k2
        if len(pts) < 2:
            # короткий "язычок" целиком в зоне развилки: прямая от развилки к листу
            if not closed and (free[0] or free[1]) and (stub[0] is not None or stub[1] is not None):
                js = stub[0] if stub[0] is not None else stub[1]
                lf = ch[-1] if free[1] else ch[0]
                pts = [node_xy[js], node_xy[lf]] if free[1] else [node_xy[lf], node_xy[js]]
                stub = [None, None]
            else:
                continue
        pts = _smooth(pts, closed, prm['smooth_it'])
        if not closed:
            for end in (0, 1):
                if not free[end]:
                    continue
                q = pts if end == 1 else pts[::-1]
                d = _tangent(q, len(q)-1, W, rev=False)
                D = _raycast(rings_in, q[-1], d)
                if D is not None:
                    ext = D - prm['tool']/2
                    if ext > 0:
                        q.append((q[-1][0]+d[0]*ext, q[-1][1]+d[1]*ext))
                    else:
                        while len(q) > 2 and math.dist(q[-1], q[-2]) <= -ext:
                            ext += math.dist(q[-1], q[-2]); q.pop()
                pts = q if end == 1 else q[::-1]
        data.append([pts, closed, stub, ci])
    # 5) заглушки: продлеваем до сквозной линии развилки или до её центра
    ci2pts = {d[3]: d[0] for d in data}
    for d in data:
        pts, closed, stub, ci = d
        for end in (0, 1):
            j = stub[end]
            if j is None:
                continue
            q = pts if end == 1 else pts[::-1]
            tgt = None
            if j in thru_of and thru_of[j] in ci2pts:
                dvec = _tangent(q, len(q)-1, W, rev=False)
                poly = ci2pts[thru_of[j]]
                if thru_of[j] == ci:
                    # сквозная линия — эта же цепочка: исключаем свой хвост
                    m = len(q) - 1; L = 0.0
                    while m > 0 and L < 3*W:
                        L += math.dist(q[m], q[m-1]); m -= 1
                    poly = q[:m+1]
                hit = _ray_poly(poly, q[-1], dvec)
                if hit is not None and hit < math.dist(q[-1], node_xy[j])*1.5 + W:
                    tgt = (q[-1][0]+dvec[0]*hit, q[-1][1]+dvec[1]*hit)
            if tgt is None:
                tgt = node_xy[j]
            q.append(tgt)
            d[0] = q if end == 1 else q[::-1]
            pts = d[0]
    paths = []
    for pts, closed, stub, ci in data:
        segs = _fit_path(pts, closed, prm['tol'], prm['corner'], W)
        paths.append((segs, closed))
    if debug is not None:
        debug.update(W=W, T=T, px=px, py=py, node_xy=node_xy, adj=adj)
    return paths


def _evenodd_area(rings):
    # площадь заливки: сумма знаковых площадей, где знак = (+1 если кольцо
    # чётной глубины вложенности)
    tot = 0.0
    for i, r in enumerate(rings):
        depth = 0
        x, y = r[0]
        for j, q in enumerate(rings):
            if j != i and inside([q], x, y):
                depth += 1
        a = abs(ring_area(r))
        tot += a if depth % 2 == 0 else -a
    return tot


def _dissolve_junctions(adj, kind):
    for i in list(adj):
        if kind[i] == 1 and i in adj and len(adj[i]) == 2:
            a, b = list(adj[i])
            adj[a].discard(i); adj[b].discard(i)
            adj[a].add(b); adj[b].add(a)
            del adj[i]
        elif kind[i] == 1 and i in adj and len(adj[i]) < 2:
            for j in adj[i]:
                adj[j].discard(i)
            del adj[i]


def _prune(adj, xy, minlen, rings=None):
    changed = False
    leaves = [i for i in adj if len(adj[i]) == 1]
    for lf in leaves:
        if lf not in adj or len(adj[lf]) != 1:
            continue
        path = [lf]; L = 0.0; prev = None; cur = lf
        while True:
            nb = [j for j in adj[cur] if j != prev]
            if len(adj[cur]) > 2 or not nb:
                break
            if len(adj[cur]) == 2 or cur == lf:
                nxt = nb[0]
                L += math.dist(xy[cur], xy[nxt])
                prev, cur = cur, nxt
                path.append(cur)
                if len(adj[cur]) != 2:
                    break
        # cur — развилка (deg>=3) или другой лист
        reach = L
        if len(path) >= 2 and rings is not None:
            # "вылет" ветки: длина + расстояние от листа до края по направлению ветки
            a = xy[path[min(len(path)-1, 3)]]; b = xy[lf]
            d = _norm((b[0]-a[0], b[1]-a[1]))
            D = _raycast(rings, b, d)
            if D is not None:
                reach += D
        if len(adj[cur]) >= 3 and reach < minlen:
            for k in path[:-1]:
                for j in list(adj.get(k, ())):
                    adj[j].discard(k)
                adj.pop(k, None)
            changed = True
    return changed


def _drop_small(adj, xy, W):
    seen = set()
    for s in list(adj):
        if s in seen:
            continue
        comp = []; st = [s]; seen.add(s)
        while st:
            i = st.pop(); comp.append(i)
            for j in adj[i]:
                if j not in seen:
                    seen.add(j); st.append(j)
        L = sum(math.dist(xy[i], xy[j]) for i in comp for j in adj[i])/2
        if L < W:
            for i in comp:
                adj.pop(i, None)


def _chains(adj, xy, merge_ang, dlen, passthru=None):
    """Разбивает граф на цепочки. В узлах ветвления сквозь проходит пара
    самых соосных веток (если излом < merge_ang)."""
    # для каждого узла deg>=3 — словарь пар "сквозного прохода"
    if passthru is None:
        passthru = _pairing(adj, xy, merge_ang, dlen)
    used_edges = set()
    def ek(a, b): return (min(a, b), max(a, b))
    chains = []
    def walk(start, nxt):
        ch = [start]; prev, cur = start, nxt
        used_edges.add(ek(prev, cur))
        while True:
            ch.append(cur)
            if cur == start and len(adj[cur]) == 2:
                return ch, True
            if cur == start and not (prev in passthru.get(cur, {})):
                return ch, False
            deg = len(adj[cur])
            if deg == 2:
                n2 = [k for k in adj[cur] if k != prev][0]
            elif deg >= 3 and prev in passthru[cur]:
                n2 = passthru[cur][prev]
            else:
                return ch, False
            if ek(cur, n2) in used_edges:
                return ch, False
            used_edges.add(ek(cur, n2))
            prev, cur = cur, n2
    # сначала от листьев и развилок
    for i in adj:
        if len(adj[i]) != 2:
            for j in adj[i]:
                if ek(i, j) not in used_edges:
                    if len(adj[i]) >= 3 and j in passthru[i]:
                        continue  # стартуем только с концов сквозных
                    chains.append(walk(i, j))
    # оставшиеся сквозные (концы которых — тоже развилки) и циклы
    for i in adj:
        for j in adj[i]:
            if ek(i, j) not in used_edges:
                ch, cl = walk(i, j)
                if not cl and len(adj[i]) >= 3 and j in passthru.get(i, {}):
                    # дойти назад через сквозной проход
                    back, _ = walk(i, passthru[i][j]) if ek(i, passthru[i][j]) not in used_edges else ([i], False)
                    ch = list(reversed(back))[:-1] + ch
                chains.append((ch, cl))
    return chains, passthru


def _tangent(pts, i, W, rev):
    # направление наружу у конца, по участку длиной ~W
    if not rev:
        j = i; L = 0
        while j > 0 and L < W*0.7:
            L += math.dist(pts[j], pts[j-1]); j -= 1
        return _norm((pts[i][0]-pts[j][0], pts[i][1]-pts[j][1]))
    j = i; L = 0
    while j < len(pts)-1 and L < W*0.7:
        L += math.dist(pts[j], pts[j+1]); j += 1
    return _norm((pts[i][0]-pts[j][0], pts[i][1]-pts[j][1]))


def _smooth(pts, closed, it, fixed=()):
    # Taubin (lambda/mu): сглаживает зубцы, не "сжимая" дуги
    p = list(pts); n = len(p)
    for k in range(it * 2):
        f = 0.5 if k % 2 == 0 else -0.53
        q = list(p)
        rng = range(n) if closed else range(1, n-1)
        for i in rng:
            if i in fixed:
                continue
            a = p[(i-1) % n]; b = p[(i+1) % n]
            q[i] = (p[i][0] + f*((a[0]+b[0])/2 - p[i][0]), p[i][1] + f*((a[1]+b[1])/2 - p[i][1]))
        p = q
    return p


def _fit_path(pts, closed, tol, corner_deg, W):
    # убрать дубли
    p = [pts[0]]
    for q in pts[1:]:
        if math.dist(q, p[-1]) > 1e-6:
            p.append(q)
    if closed and math.dist(p[0], p[-1]) < 1e-6:
        p.pop()
    n = len(p)
    # углы излома по окну ~0.5W
    def turn(i):
        a = i; L = 0
        while L < W*0.5:
            if not closed and a == 0: break
            a2 = (a-1) % n; L += math.dist(p[a], p[a2]); a = a2
            if a == i: break
        b = i; L = 0
        while L < W*0.5:
            if not closed and b == n-1: break
            b2 = (b+1) % n; L += math.dist(p[b], p[b2]); b = b2
            if b == i: break
        if a == i or b == i:
            return 0.0
        d1 = _norm((p[i][0]-p[a][0], p[i][1]-p[a][1]))
        d2 = _norm((p[b][0]-p[i][0], p[b][1]-p[i][1]))
        return math.degrees(math.acos(max(-1, min(1, d1[0]*d2[0]+d1[1]*d2[1]))))
    ang = [turn(i) for i in range(n)]
    def local(i):
        # излом в самой точке (по соседям): у настоящего угла почти весь поворот окна
        # сосредоточен в одной точке, у крутой гладкой дуги (шпилька «ы», «в») — размазан
        if not closed and (i == 0 or i == n-1):
            return 0.0
        a = p[(i-1) % n]; b = p[(i+1) % n]
        d1 = _norm((p[i][0]-a[0], p[i][1]-a[1])); d2 = _norm((b[0]-p[i][0], b[1]-p[i][1]))
        return math.degrees(math.acos(max(-1, min(1, d1[0]*d2[0]+d1[1]*d2[1]))))
    corners = []
    for i in range(n):
        if ang[i] > corner_deg and ang[i] >= ang[(i-1) % n] and ang[i] >= ang[(i+1) % n]:
            if not closed and (i == 0 or i == n-1):
                continue
            if corners and (i - corners[-1]) < 3:
                continue
            if SMOOTH_ARC and max(local(j) for j in range(max(0, i-2), min(n, i+3))) < 0.35*ang[i]:
                continue
            corners.append(i)
    if closed:
        if not corners:
            # замкнутый без углов: разрезаем в точке 0, касательные общие
            pp = p + [p[0]]
            t = _norm((p[1][0]-p[-1][0], p[1][1]-p[-1][1]))
            out = []
            fit(pp, t, (-t[0], -t[1]), tol, out, LMIN_K*W if LMIN_K else None)
            return out
        p = p[corners[0]:] + p[:corners[0]+1]
        corners = [c - corners[0] if c >= corners[0] else c - corners[0] + n for c in corners]
        corners = corners + [len(p)-1]
    else:
        corners = [0] + corners + [n-1]
    out = []
    for a, b in zip(corners[:-1], corners[1:]):
        seg = p[a:b+1]
        if len(seg) < 2:
            continue
        t1 = _endtan(seg, True, W); t2 = _endtan(seg, False, W)
        fit(seg, t1, t2, tol, out, LMIN_K*W if LMIN_K else None)
    return _merge_lines(out, tol)


def _merge_lines(segs, tol):
    out = []
    for s in segs:
        if out and s[0] == 'L' and out[-1][0] == 'L':
            a = out[-1][1]; m = out[-1][2]; b = s[2]
            if _line_dev([a, m, b]) <= tol:
                out[-1] = ('L', a, b); continue
        out.append(s)
    return out


def _endtan(seg, start, W):
    if start:
        j = 1; L = 0
        while j < len(seg)-1 and L < W*0.25:
            L += math.dist(seg[j], seg[j-1]); j += 1
        return _norm((seg[j][0]-seg[0][0], seg[j][1]-seg[0][1]))
    j = len(seg)-2; L = 0
    while j > 0 and L < W*0.25:
        L += math.dist(seg[j], seg[j+1]); j -= 1
    return _norm((seg[j][0]-seg[-1][0], seg[j][1]-seg[-1][1]))


def _raycast(rings, o, d):
    """Расстояние от точки o вдоль единичного d до ближайшего пересечения с контуром."""
    best = None
    for r in rings:
        n = len(r)
        for i in range(n):
            ax, ay = r[i]; bx, by = r[(i+1) % n]
            ex, ey = bx-ax, by-ay
            den = d[0]*ey - d[1]*ex
            if abs(den) < 1e-12:
                continue
            t = ((ax-o[0])*ey - (ay-o[1])*ex) / den
            u = ((ax-o[0])*d[1] - (ay-o[1])*d[0]) / den
            if t > 1e-9 and -1e-9 <= u <= 1+1e-9:
                if best is None or t < best:
                    best = t
    return best


def _ray_poly(poly, o, d):
    best = None
    for i in range(len(poly)-1):
        ax, ay = poly[i]; bx, by = poly[i+1]
        ex, ey = bx-ax, by-ay
        den = d[0]*ey - d[1]*ex
        if abs(den) < 1e-12:
            continue
        t = ((ax-o[0])*ey - (ay-o[1])*ex) / den
        u = ((ax-o[0])*d[1] - (ay-o[1])*d[0]) / den
        if t > -1e-9 and -1e-9 <= u <= 1+1e-9:
            if best is None or t < best:
                best = t
    return best


def _merge_junctions(adj, xy, kind, rr, maxd):
    """Сливает развилки, соединённые короткой цепочкой (типично для Т-образных
    стыков: Делоне даёт 2 треугольника-развилки рядом)."""
    changed = True
    while changed:
        changed = False
        for j in list(adj):
            if j not in adj or len(adj[j]) < 3:
                continue
            for n0 in list(adj[j]):
                path = [j]; prev, cur = j, n0; L = math.dist(xy[j], xy[n0])
                while cur in adj and len(adj[cur]) == 2 and L <= maxd:
                    nb = [k for k in adj[cur] if k != prev]
                    path.append(cur); prev, cur = cur, nb[0]
                    L += math.dist(xy[prev], xy[cur])
                if cur != j and cur in adj and len(adj[cur]) >= 3 and L <= maxd:
                    path.append(cur)
                    # новый узел — центр масс двух развилок
                    k2 = cur
                    nx = ((xy[j][0]+xy[k2][0])/2, (xy[j][1]+xy[k2][1])/2)
                    nid = len(xy); xy.append(nx); kind.append(1); rr.append(1e9)
                    ext = set()
                    for p in path:
                        for q in adj[p]:
                            if q not in path:
                                ext.add(q)
                    for p in path:
                        for q in list(adj[p]):
                            adj[q].discard(p)
                        del adj[p]
                    adj[nid] = set()
                    for q in ext:
                        adj[nid].add(q); adj[q].add(nid)
                    changed = True
                    break
            if changed:
                break


def _hermite_fill(before, after, step, center=None):
    a = before[-1]; b = after[0]
    chord = _norm((b[0]-a[0], b[1]-a[1]))
    win = step*6
    ta = _end_dir(before[::-1], win) or chord
    tb = _end_dir(after, win)
    tb = (-tb[0], -tb[1]) if tb else chord
    if center is None:
        return _herm(a, ta, b, tb, step)
    tc = _norm((b[0]-a[0], b[1]-a[1]))
    return _herm(a, ta, center, tc, step) + [center] + _herm(center, tc, b, tb, step)


def _end_dir(seq, win):
    """Направление ломаной в точке seq[0] (наружу, к seq[0]) — МНК-прямая по точкам
    в пределах win от seq[0]; None, если точек мало."""
    P = [seq[0]]
    for p in seq[1:]:
        if math.dist(p, seq[0]) > win:
            break
        P.append(p)
    if len(P) < 2 or math.dist(P[0], P[-1]) < 1e-9:
        return None
    n = len(P)
    mx = sum(p[0] for p in P)/n; my = sum(p[1] for p in P)/n
    sxx = sum((p[0]-mx)**2 for p in P); syy = sum((p[1]-my)**2 for p in P)
    sxy = sum((p[0]-mx)*(p[1]-my) for p in P)
    ang = 0.5*math.atan2(2*sxy, sxx-syy)
    d = (math.cos(ang), math.sin(ang))
    if (P[0][0]-P[-1][0])*d[0] + (P[0][1]-P[-1][1])*d[1] < 0:
        d = (-d[0], -d[1])
    return d


def _herm(a, ta, b, tb, step):
    L = math.dist(a, b)
    m = max(1, int(L/step))
    out = []
    for k in range(1, m):
        t = k/m
        h00 = 2*t**3-3*t*t+1; h10 = t**3-2*t*t+t; h01 = -2*t**3+3*t*t; h11 = t**3-t*t
        out.append((h00*a[0]+h10*L*ta[0]+h01*b[0]+h11*L*tb[0],
                    h00*a[1]+h10*L*ta[1]+h01*b[1]+h11*L*tb[1]))
    return out


def _pairing(adj, xy, merge_ang, dlen, jw=None, dirfn=None, pairfn=None):
    passthru = {}
    def dirto(i, j):
        w_ = (dlen if jw is None else 1.5 * jw.get(i, dlen)) / 1.5
        if dirfn is not None:
            return dirfn(adj, xy, i, j, w_)
        return _branch_dir(adj, xy, i, j, w_)
    for i in adj:
        if len(adj[i]) >= 3:
            nb = list(adj[i]); pr = {j: dirto(i, j) for j in nb}
            ds = {j: pr[j][0] for j in nb}; pe = {j: pr[j][1] for j in nb}
            wi = (dlen if jw is None else 1.5 * jw.get(i, dlen)) / 1.5
            cand = []
            for a in range(len(nb)):
                for b in range(a+1, len(nb)):
                    da, db = ds[nb[a]], ds[nb[b]]
                    dot = -(da[0]*db[0]+da[1]*db[1])  # 1 = прямо
                    # параллельные, но смещённые ветки (ножка Д и её пятка, ножка и хвост Ц) —
                    # не сквозные: точка одной ветки на ~1w должна лежать на касательной другой
                    # (с поправкой на кривизну обеих)
                    if not _aimed(pr[nb[a]], pr[nb[b]], wi) or not _aimed(pr[nb[b]], pr[nb[a]], wi):
                        continue
                    cand.append((dot, nb[a], nb[b]))
            cand.sort(reverse=True)
            if pairfn is not None:
                # внешний критерий пар (режим образца): совместная подгонка дальних участков
                cand = pairfn(adj, xy, i, nb, wi, pr)
            used = set(); pm = {}; segs = []
            for dot, a, b in cand:
                if a in used or b in used:
                    continue
                if math.degrees(math.acos(max(-1, min(1, dot)))) <= merge_ang:
                    # вторая и далее пары — только настоящий крест: отрезок между
                    # концами проб этой пары должен пересекать отрезки всех принятых пар
                    if all(_segs_cross(pe[a], pe[b], p, q) for p, q in segs):
                        pm[a] = b; pm[b] = a; used |= {a, b}; segs.append((pe[a], pe[b]))
            if not pm and len(nb) == 3:
                # тройник без сквозной пары: угол внутри развилки (пятка Д: диагональ в
                # перекладину, пятка сбоку) — сквозной считаем пару, оси которой
                # пересекаются ближе всего к развилке (не дальше 1w); третья ветка — заглушка
                best = None
                for a in range(3):
                    for b in range(a+1, 3):
                        X = _line_x(pr[nb[a]][2], ds[nb[a]], pr[nb[b]][2], ds[nb[b]])
                        if X is None:
                            continue
                        dX = math.dist(X, xy[i])
                        if dX > 1.0*wi:
                            continue
                        # третья ветка должна упираться в одну из двух осей ЗА углом X
                        # (по ходу ветки от X), иначе заглушке не во что упереться
                        c = nb[3 - a - b]; ok = False
                        for k in (a, b):
                            Xc = _line_x(pr[nb[k]][2], ds[nb[k]], pr[c][2], ds[c])
                            if Xc is not None and math.dist(Xc, X) <= 2.0*wi and \
                               (Xc[0]-X[0])*ds[nb[k]][0] + (Xc[1]-X[1])*ds[nb[k]][1] > 0:
                                ok = True
                        key = (0 if ok else 1, dX)
                        if best is None or key < best[0]:
                            best = (key, nb[a], nb[b])
                if best is not None:
                    pm[best[1]] = best[2]; pm[best[2]] = best[1]
            passthru[i] = pm
    return passthru


def _line_x(p1, d1, p2, d2):
    """Пересечение прямых (p1, d1) и (p2, d2) или None (параллельны)."""
    den = d1[0]*d2[1] - d1[1]*d2[0]
    if abs(den) < 1e-9:
        return None
    t = ((p2[0]-p1[0])*d2[1] - (p2[1]-p1[1])*d2[0]) / den
    return (p1[0] + d1[0]*t, p1[1] + d1[1]*t)


def _aimed(pa, pb, w):
    """Ветка b «целится» в ветку a: её точка на ~1w лежит на касательной a (в точке входа)
    с допуском 0.35w плюс стрелка дуги на 1w для кривизны обеих веток."""
    da, _, near_a, _, Ra = pa
    _, _, _, mid_b, Rb = pb
    tol = 0.35*w + (0.5*w*w/Ra if Ra else 0.0) + (0.5*w*w/Rb if Rb else 0.0)
    vx, vy = mid_b[0]-near_a[0], mid_b[1]-near_a[1]
    return abs(vx*da[1] - vy*da[0]) <= tol


def _segs_cross(a, b, c, d):
    def orient(p, q, r):
        return (q[0]-p[0])*(r[1]-p[1]) - (q[1]-p[1])*(r[0]-p[0])
    o1 = orient(a, b, c); o2 = orient(a, b, d); o3 = orient(c, d, a); o4 = orient(c, d, b)
    return (o1 > 0) != (o2 > 0) and (o3 > 0) != (o4 > 0)


def _merge_bridges(adj, xy, kind, rr, maxlen, merge_ang, dlen):
    """Короткая перемычка между двумя развилками, которая ни в одной из них
    не идёт насквозь (талия "8"), стягивается в одну крестовую развилку."""
    while True:
        pt = _pairing(adj, xy, merge_ang, dlen)
        done = False
        for j in list(adj):
            if len(adj[j]) < 3:
                continue
            for n0 in list(adj[j]):
                if n0 in pt.get(j, {}):
                    continue
                path = [j]; prev, cur = j, n0; L = math.dist(xy[j], xy[n0])
                while len(adj[cur]) == 2 and L <= maxlen:
                    nb = [k for k in adj[cur] if k != prev][0]
                    path.append(cur); prev, cur = cur, nb
                    L += math.dist(xy[prev], xy[cur])
                if cur == j or len(adj[cur]) < 3 or L > maxlen or prev in pt.get(cur, {}):
                    continue
                path.append(cur)
                nid = len(xy); xy.append(((xy[j][0]+xy[cur][0])/2, (xy[j][1]+xy[cur][1])/2))
                kind.append(1); rr.append(1e9)
                ext = set(q for p in path for q in adj[p] if q not in path)
                for p in path:
                    for q in list(adj[p]):
                        adj[q].discard(p)
                    del adj[p]
                adj[nid] = set(ext)
                for q in ext:
                    adj[q].add(nid)
                done = True
                break
            if done:
                break
        if not done:
            return


def _branch_dir(adj, xy, i, j, w):
    """Направление ветки развилки i, идущей через соседа j, в точке входа в развилку:
    касательная окружности (или прямой), подогнанной к узлам ветки на 0.3..2.0 w от
    развилки. Центр развилки может лежать не на оси, поэтому направление "от центра"
    не годится. Возвращает (единичный вектор наружу, дальняя точка пробы)."""
    prev, cur = i, j; L = math.dist(xy[i], xy[j])
    P = []; PL = []
    if L >= 0.3*w:
        P.append(xy[cur]); PL.append(L)
    while L < 2.0*w:
        nb = [k for k in adj[cur] if k != prev]
        if len(adj[cur]) != 2 or not nb:
            break
        prev, cur = cur, nb[0]; L += math.dist(xy[prev], xy[cur])
        if L >= 0.3*w:
            P.append(xy[cur]); PL.append(L)
    if len(adj[cur]) == 1 and P:
        # ветка кончается свободным концом: узлы у скруглённого торца (крючок) — выбросить
        keep = [k for k in range(len(P)) if PL[k] <= L - 0.7*w]
        if len(keep) >= 2:
            P = [P[k] for k in keep]; PL = [PL[k] for k in keep]
    # излом внутри окна (перекладина Ц -> хвост): направление меряем до излома
    if len(P) >= 3:
        d0 = None; cut = None
        for k in range(1, len(P)):
            m = k
            while m > 0 and PL[k] - PL[m] < 0.4*w:
                m -= 1
            if PL[k] - PL[m] < 0.3*w:
                continue
            dk = _norm((P[k][0]-P[m][0], P[k][1]-P[m][1]))
            if d0 is None:
                d0 = dk
            elif d0[0]*dk[0] + d0[1]*dk[1] < math.cos(math.radians(50)):
                cut = m
                break
        if cut is not None and cut >= 2:
            P = P[:cut+1]; L = PL[cut]
    far = P[-1] if P else xy[cur]
    if len(P) < 2:
        return _norm((far[0]-xy[i][0], far[1]-xy[i][1])), far, far, far, None
    near = P[0]
    mid = min(zip(P, PL), key=lambda t: abs(t[1] - 1.0*w))[0]   # точка на ~1w от развилки
    chord = _norm((far[0]-near[0], far[1]-near[1]))
    if len(P) >= 4 and L >= 1.5*w:        # короткой ветке (хвост Q) окружность не доверяем
        c = _fit_circle_pts(P)
        if c is not None and c[2] < 25*w and c[3] < 0.03*w and math.dist(P[0], P[-1])**2/(8*c[2]) > 0.03*w:
            cx, cy, R = c[0], c[1], c[2]
            t = _norm((-(near[1]-cy), near[0]-cx))
            if t[0]*chord[0] + t[1]*chord[1] < 0:
                t = (-t[0], -t[1])
            return t, far, near, mid, R
    return chord, far, near, mid, None


def _fit_circle_pts(P):
    n = len(P)
    mx = sum(p[0] for p in P)/n; my = sum(p[1] for p in P)/n
    suu = suv = svv = suuu = svvv = suvv = svuu = 0.0
    for x, y in P:
        u = x-mx; v = y-my
        suu += u*u; suv += u*v; svv += v*v
        suuu += u*u*u; svvv += v*v*v; suvv += u*v*v; svuu += v*u*u
    det = suu*svv - suv*suv
    if abs(det) < 1e-12:
        return None
    b1 = 0.5*(suuu + suvv); b2 = 0.5*(svvv + svuu)
    uc = (b1*svv - b2*suv)/det; vc = (suu*b2 - suv*b1)/det
    cx = uc+mx; cy = vc+my
    R = math.sqrt(max(0.0, uc*uc + vc*vc + (suu+svv)/n))
    rms = math.sqrt(sum((math.dist(p, (cx, cy))-R)**2 for p in P)/n)
    return cx, cy, R, rms
