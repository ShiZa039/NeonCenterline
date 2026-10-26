"""Версия 2: локальная ширина штриха (рукописные шрифты с переменной толщиной)
и перекрывающиеся контуры (соединения букв в Segoe Script и т.п.).

Вход: список полигонов; полигон = [внешнее кольцо, дырки...]. Полигоны могут
перекрываться — тогда заливка = их объединение.
"""
import math
import centerline as C
from centerline import (P as P1, ring_area, ring_len, resample, inside, delaunay,
                        _pairing, _smooth, _fit_path, _hermite_fill, _norm, _ray_poly, _herm, _line_dev)

P = dict(P1)
P.update(
    jmerge=0.5,      # развилки ближе jmerge*ширины сливаются (только дубли треугольников у Т-стыка)
    alpha=0.0,      # узел выбрасывается, если хорда < alpha * max хорды по соседству
    prune=1.0,       # ветка с вылетом < prune * ширины у развилки удаляется
    end_trim=0.92,   # у конца срезаем, пока хорда растёт быстрее этого
    hstep=7.0,       # шаг ресэмплинга = W/hstep
)


def _inside_any(polys, x, y, skip=-1):
    for k, p in enumerate(polys):
        if k != skip and inside(p, x, y):
            return True
    return False


def centerline(polys, prm=P, debug=None):
    # --- средняя ширина (только масштаб)
    area = 0.0; per = 0.0
    for p in polys:
        a = abs(ring_area(p[0])) - sum(abs(ring_area(h)) for h in p[1:])
        area += a; per += sum(ring_len(r) for r in p)
    W = 2 * area / per
    h = W / prm['hstep']
    # режим «шрифт — образец»: ref=1 — до кончика штриха, ref=2 — отступ полширины; фреза не учитывается
    if prm.get('ref'):
        prm = dict(prm); prm['tool'] = 0.05*W if prm['ref'] == 1 else W; prm['recenter'] = False
    # крошечные дырки (счётчик меньше самой фрезы: треугольник в рукописной «т») — не контур
    polys = [[p[0]] + [r for r in p[1:] if abs(ring_area(r)) >= 0.3*W*W] for p in polys]

    # --- точки: ресэмплинг, выкидываем точки, лежащие внутри других полигонов
    px, py, rings_rs = [], [], []
    for k, p in enumerate(polys):
        rp = []
        for r in p:
            rs = resample(r, h)
            rp.append(rs)
            for (x, y) in rs:
                if len(polys) > 1 and _inside_any(polys, x, y, skip=k):
                    continue
                px.append(x); py.append(y)
        rings_rs.append(rp)
    tris = delaunay(px, py)
    T = []
    for a, b, c in tris:
        cx = (px[a]+px[b]+px[c])/3; cy = (py[a]+py[b]+py[c])/3
        if any(inside(rp, cx, cy) for rp in rings_rs):
            T.append((a, b, c))
    # ребро внутреннее, если его делят два внутренних треугольника
    ecnt = {}
    for a, b, c in T:
        for e in ((a, b), (b, c), (c, a)):
            k = (min(e), max(e)); ecnt[k] = ecnt.get(k, 0) + 1

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
        ie = [e for e in ((a, b), (b, c), (c, a)) if ecnt[(min(e), max(e))] == 2]
        if len(ie) == 2:
            link(enode(*ie[0]), enode(*ie[1]))
        elif len(ie) == 3:
            ns = [enode(*e) for e in ie]
            jx = sum(node_xy[n][0] for n in ns)/3; jy = sum(node_xy[n][1] for n in ns)/3
            j = len(node_xy); node_xy.append((jx, jy)); node_r.append(1e9); node_kind.append(1)
            for n in ns:
                link(j, n)

    G = dict(adj=adj, xy=node_xy, r=node_r, kind=node_kind, polys=polys, bcorners=_boundary_corners(polys, W), W=W)

    # 1) фильтр по локальной хорде: хорда < alpha * max хорды в радиусе своей хорды
    kill = []
    for i in list(adj):
        if node_kind[i] != 0:
            continue
        c = node_r[i]
        mx = _local_max_chord(G, i, c)
        if c < prm['alpha'] * mx:
            kill.append(i)
    for i in kill:
        _kill(adj, i)
    _dissolve(G)

    # 2) висячие ветки
    for _ in range(4):
        ch = _prune(G, prm['prune'])
        _dissolve(G)
        if not ch:
            break
    if prm.get('ref'):
        # шрифт — образец: «ушки» (короткий выступ стойки над местом, где в неё
        # вливается дуга) перо не рисует — дуга поворачивает прямо в стойку
        for _ in range(2):
            ch = _prune_ears(G, prm.get('ear', 1.8))
            _dissolve(G)
            if not ch:
                break
    _merge_junctions(G, prm['jmerge'])
    if prm.get('ref'):
        # шрифт — образец: две близкие развилки, которые вместе — настоящий крест
        # (перетяжка «з», петли): сливаются в одну, штрихи проходят крест насквозь
        _merge_x(G, prm.get('xmerge', 1.0), 50.0)
    _merge_bridges(G, prm['merge_ang'])
    _drop_small(G)

    chains, passthru = _chains(G, prm['merge_ang'], prm['tool'], bool(prm.get('ref')))
    data = _process(G, chains, passthru, prm, W, h)
    paths = []
    segA = segB = None
    for pts, closed, stub, ci in data:
        if prm.get('recenter', True):
            if segA is None:
                import numpy as _np
                A = []; B = []
                for p in polys:
                    for r in p:
                        for i in range(len(r)):
                            A.append(r[i]); B.append(r[(i+1) % len(r)])
                segA = _np.array(A, float); segB = _np.array(B, float)
            pts = _recenter(pts, closed, polys, 0.5*prm['tool'] + min(0.05*W, 0.5), W, segA, segB,
                              concave=[c[:2] for c in G['bcorners'] if not c[2] and
                                       c[3][0]*c[4][0] + c[3][1]*c[4][1] < math.cos(math.radians(75))])
        if prm.get('ref') and not closed:
            pts = _fillet(pts, W, prm.get('fillet', 0.5))
        pts = _deglitch(pts, closed, W)
        C.SMOOTH_ARC = bool(prm.get('ref'))
        segs = _fit_path(pts, closed, prm['tol'], prm['corner'], _median_chord_pts(G, pts, W))
        C.SMOOTH_ARC = False
        paths.append((segs, closed))
    paths = _join_paths(paths, 0.05*W)
    if debug is not None:
        debug.update(W=W, node_xy=node_xy, adj=adj, px=px, py=py, T=T)
    return paths


def _recenter(pts, closed, polys, rad, W, segA, segB, it=3, prm_it=40, concave=()):
    """Курсив: в местах, где ось подошла к краю ближе радиуса фрезы (канал вырезал бы контур —
    узкий проход между наложенными штрихами, вершина острого разворота у самого кончика),
    ось сдвигается от ближнего края — но не дальше середины до противоположного края.
    Сдвиги сглаживаются вдоль линии (без ступенек); ломаная без таких мест не меняется."""
    import numpy as np
    def clear(P):
        P = np.asarray(P, float)
        d = segB - segA; L2 = (d*d).sum(1); L2[L2 == 0] = 1e-12
        out_d = np.empty(len(P)); out_n = np.zeros((len(P), 2))
        for k, p in enumerate(P):
            t = np.clip(((p - segA)*d).sum(1)/L2, 0, 1)
            c = segA + d*t[:, None]; v = p - c; dd = np.hypot(v[:, 0], v[:, 1])
            m = int(dd.argmin()); out_d[k] = dd[m]
            if dd[m] > 1e-9:
                out_n[k] = v[m]/dd[m]
        return out_d, out_n
    n0 = len(pts)
    if n0 < 2:
        return pts
    d0, _ = clear(pts)
    # быстрый выход: и узлы, и середины звеньев далеко от края
    mids = [((a[0]+b[0])/2, (a[1]+b[1])/2) for a, b in zip(pts[:-1], pts[1:])]
    dm, _ = clear(mids) if mids else (np.array([rad]), None)
    if d0.min() >= rad and dm.min() >= rad:
        return pts
    # уплотнение (шаг 0.15W), исходные точки сохраняются
    step = 0.15*W
    P = [pts[0]]
    seq = list(zip(pts[:-1], pts[1:])) + ([(pts[-1], pts[0])] if closed else [])
    for a, b in seq:
        m = max(1, int(math.dist(a, b)/step))
        for k in range(1, m + 1):
            P.append((a[0] + (b[0]-a[0])*k/m, a[1] + (b[1]-a[1])*k/m))
    if closed:
        P.pop()
    P = np.array(P, float); n = len(P); P0 = P.copy()
    s = np.concatenate([[0.0], np.cumsum(np.hypot(*np.diff(P, axis=0).T))])
    endz = 3.0*rad
    # триггер: заметное нарушение там, где канал помещается
    d, nv = clear(P)
    hot = np.zeros(n, bool)
    for k in range(n):
        if not closed and (s[k] < endz or s[-1] - s[k] < endz):
            continue
        if d[k] > rad - min(0.12*W, 0.3*rad) or not nv[k].any():
            continue
        # ближе всего — вогнутый угол контура (стык штрихов, Д): это нормально, не трогаем
        e = P[k] - nv[k]*d[k]
        if any(math.dist(e, c) < 0.3*W for c in concave):
            continue
        d2 = _raycast_union(polys, tuple(P[k]), tuple(nv[k]))
        if d2 is not None and d[k] + d2 >= 2.0*rad + min(0.1*W, 0.6*rad):
            hot[k] = True
    if not hot.any():
        return pts
    # зона действия: до 2W вдоль линии от нарушений; вне неё точки неподвижны
    act = np.zeros(n, bool)
    for k in np.nonzero(hot)[0]:
        ds = np.abs(s - s[k])
        if closed:
            ds = np.minimum(ds, s[-1] + math.dist(P[-1], P[0]) - ds)
        act |= ds < 2.0*W
    if not closed:
        act[0] = act[-1] = False
        act &= (s >= endz) & (s[-1] - s >= endz)
    # «змейка»: отталкивание от края (до rad) + натяжение (Лаплас), 40 итераций
    for _ in range(int(prm_it)):
        d, nv = clear(P)
        F = nv*np.clip(rad - d, 0, None)[:, None]*0.5
        Lp = np.zeros_like(P)
        if closed:
            Lp = 0.5*(np.roll(P, 1, 0) + np.roll(P, -1, 0)) - P
        else:
            Lp[1:-1] = 0.5*(P[:-2] + P[2:]) - P[1:-1]
        Pn = P + (F + 0.5*Lp)*act[:, None]
        ok = np.array([_inside_any(polys, x, y) for x, y in Pn])
        Pn[~ok] = P[~ok]
        P = Pn
    if np.hypot(*(P - P0).T).max() < 0.02*W:
        return pts                        # сдвигать нечего (штрих уже фрезы — ось и так посередине)
    return [tuple(p) for p in P]


def _deglitch(pts, closed, W, it=4):
    """Финальная зачистка мелких ступенек (стыки этапов: продление конца, проекция после развилки):
    точки с резким локальным изломом (> 8 град на соседях) сглаживаются, если на масштабе 0.5W
    поворот невелик (< 30 град) — настоящие углы и шпильки не трогаются; концы — на месте."""
    n = len(pts)
    if n < 5:
        return pts
    def turn_w(i):
        a = i; L = 0.0
        while L < 0.5*W:
            if not closed and a == 0:
                break
            a2 = (a - 1) % n; L += math.dist(pts[a], pts[a2]); a = a2
            if a == i:
                break
        b = i; L = 0.0
        while L < 0.5*W:
            if not closed and b == n - 1:
                break
            b2 = (b + 1) % n; L += math.dist(pts[b], pts[b2]); b = b2
            if b == i:
                break
        if a == i or b == i:
            return 180.0
        d1 = _norm((pts[i][0]-pts[a][0], pts[i][1]-pts[a][1])); d2 = _norm((pts[b][0]-pts[i][0], pts[b][1]-pts[i][1]))
        return math.degrees(math.acos(max(-1.0, min(1.0, d1[0]*d2[0] + d1[1]*d2[1]))))
    protect = set(i for i in range(n) if turn_w(i) > 30.0)
    rng = range(n) if closed else range(1, n - 1)
    for _ in range(it):
        new = list(pts)
        for i in rng:
            if i in protect:
                continue
            a = pts[(i - 1) % n]; b = pts[(i + 1) % n]; p = pts[i]
            d1 = _norm((p[0]-a[0], p[1]-a[1])); d2 = _norm((b[0]-p[0], b[1]-p[1]))
            if d1[0]*d2[0] + d1[1]*d2[1] < math.cos(math.radians(8)):
                new[i] = (0.5*p[0] + 0.25*(a[0] + b[0]), 0.5*p[1] + 0.25*(a[1] + b[1]))
        pts = new
    return pts


def _rev_segs(segs):
    out = []
    for s in reversed(segs):
        out.append((s[0],) + tuple(reversed(s[1:])))
    return out


def _join_paths(paths, eps):
    """Две незамкнутые линии, сходящиеся концами в одной точке (угол Z, 1, 4, 7, где
    торец перекладины и диагональ встречаются), — одна линия с углом. Точка, в которой
    сходятся три и больше линий (Т-стык), не трогается."""
    paths = [list(p) for p in paths]
    changed = True
    while changed:
        changed = False
        opens = [i for i, (segs, cl) in enumerate(paths) if not cl and segs]
        ends = {}
        for i in opens:
            segs = paths[i][0]
            for end, p in ((0, segs[0][1]), (1, segs[-1][-1])):
                key = (round(p[0]/eps), round(p[1]/eps))
                ends.setdefault(key, []).append((i, end))
        for key, lst in ends.items():
            if len(lst) != 2:
                continue
            (i, ei), (j, ej) = lst
            if i == j:
                continue
            a = paths[i][0]; b = paths[j][0]
            if ei == 0:
                a = _rev_segs(a)
            if ej == 1:
                b = _rev_segs(b)
            # стык: конец a == начало b
            if math.dist(a[-1][-1], b[0][1]) > 2*eps:
                continue
            b = [(b[0][0],) + (a[-1][-1],) + tuple(b[0][2:])] + b[1:]
            merged = a + b
            paths[i] = [merged, False]; paths[j] = [[], False]
            changed = True
            break
    return [(segs, cl) for segs, cl in paths if segs]


def _median_chord_pts(G, pts, W):
    return W  # масштаб окна углов — общий


def _kill(adj, i):
    for j in adj.get(i, ()):
        adj[j].discard(i)
    adj.pop(i, None)


def _local_max_chord(G, i, rad):
    adj, xy, r, kind = G['adj'], G['xy'], G['r'], G['kind']
    best = r[i]; seen = {i: 0.0}; st = [i]
    while st:
        a = st.pop()
        for b in adj.get(a, ()):
            d = seen[a] + math.dist(xy[a], xy[b])
            if d <= rad and (b not in seen or d < seen[b]):
                seen[b] = d; st.append(b)
                if kind[b] == 0 and r[b] > best:
                    best = r[b]
    return best


def _jwidth(G, j):
    """Ширина штриха у развилки: максимум хорд в радиусе ~ ширины."""
    adj, xy, r, kind = G['adj'], G['xy'], G['r'], G['kind']
    best = 0.0
    for n in adj.get(j, ()):
        prev, cur = j, n; k = 0
        while True:
            if kind[cur] == 0:
                best = max(best, r[cur])
            k += 1
            nb = [q for q in adj.get(cur, ()) if q != prev]
            if k >= 6 or len(adj.get(cur, ())) != 2 or not nb:
                break
            prev, cur = cur, nb[0]
    return best if best > 0 else 1.0


def _dissolve(G):
    adj, kind = G['adj'], G['kind']
    for i in list(adj):
        if kind[i] == 1 and i in adj:
            if len(adj[i]) == 2:
                a, b = list(adj[i])
                _kill(adj, i); adj[a].add(b); adj[b].add(a)
            elif len(adj[i]) < 2:
                _kill(adj, i)


def _raycast_union(polys, o, d):
    """Луч до края объединения полигонов (точки попадания внутри чужого полигона пропускаем)."""
    hits = []
    for k, p in enumerate(polys):
        for r in p:
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
                    hits.append((t, k))
    hits.sort()
    for t, k in hits:
        x = o[0]+d[0]*t; y = o[1]+d[1]*t
        if len(polys) == 1 or not _inside_any(polys, x, y, skip=k):
            return t
    return None


COLL_T = -0.75


def _leaf_reach(G, lf):
    """Путь от листа lf до развилки и вылет (длина пути + луч до контура). (path, L, reach, cur)."""
    adj, xy = G['adj'], G['xy']
    path = [lf]; L = 0.0; prev = None; cur = lf
    while True:
        nb = [j for j in adj[cur] if j != prev]
        if not nb or (cur != lf and len(adj[cur]) != 2):
            break
        nxt = nb[0]; L += math.dist(xy[cur], xy[nxt]); prev, cur = cur, nxt; path.append(cur)
        if len(adj[cur]) != 2:
            break
    a = xy[path[min(len(path)-1, 3)]]; b = xy[lf]
    d = _norm((b[0]-a[0], b[1]-a[1]))
    D = _raycast_union(G['polys'], b, d)
    return path, L, L + (D or 0.0), cur


def LEAF_KEY(G):
    # порядок обхода листьев — по координатам (снизу вверх, слева направо): одинаков в VBA
    # и в прототипе (номера узлов там разные), а итог prune от порядка зависит
    xy = G['xy']
    return lambda i: (round(xy[i][1], 3), round(xy[i][0], 3))


def _prune(G, kprune):
    adj, xy = G['adj'], G['xy']
    W_ = G.get('W', 1.0)
    changed = False
    for lf in sorted((i for i in adj if len(adj[i]) == 1), key=LEAF_KEY(G)):
        if lf not in adj or len(adj[lf]) != 1:
            continue
        path, L, reach, cur = _leaf_reach(G, lf)
        if len(adj[cur]) < 3:
            continue
        jw = _jwidth(G, cur)
        if reach < kprune * jw:
            # хвост, продолжающий другую ветку по прямой (ножка ц идёт дальше за дугу,
            # которая влилась в неё по касательной) — не усик: усик всегда торчит в сторону
            # (биссектриса угла, скругление торца)
            if reach >= 0.6 * jw and L >= 0.8 * W_ and len(path) >= 2:
                # направление хвоста — как у остальных веток (с обрезкой загиба торца)
                dl = C._branch_dir(adj, xy, cur, path[-2], jw)[0]
                dots = [(lambda dd: dd[0]*dl[0] + dd[1]*dl[1])(C._branch_dir(adj, xy, cur, o, jw)[0])
                        for o in adj[cur] if o != path[-2]]
                # соосная ветка есть, а остальные ветки идут почти вдоль неё (касательное слияние:
                # дуга ц вливается в ножку); если ветка отходит под углом (чаша p от верха ножки) —
                # это обычный усик-колпачок над развилкой
                coll = any(d_ < COLL_T for d_ in dots)
                # и у хвоста есть «тело» полной ширины (усик в угол сужается сразу)
                body = False; s = 0.0; first = None
                for k in range(len(path)-2, -1, -1):
                    s += math.dist(xy[path[k]], xy[path[k+1]])
                    if G['kind'][path[k]] == 0 and G['r'][path[k]] >= 0.9 * W_:
                        if first is None:
                            first = s
                        elif s - first >= 0.15 * W_:
                            body = True; break
                if coll and body:
                    continue
            for k in path[:-1]:
                _kill(adj, k)
            changed = True
    return changed


def _fillet(pts, W, kr):
    """Режим образца: перо поворачивает плавно — острые изломы оси (> 35 град по участкам
    ~0.6W с обеих сторон) скругляются кривой радиуса ~kr*W."""
    n = len(pts)
    if n < 5:
        return pts
    cum = [0.0]
    for i in range(1, n):
        cum.append(cum[-1] + math.dist(pts[i-1], pts[i]))
    def at(sv):
        sv = max(0.0, min(cum[-1], sv))
        import bisect
        j = max(1, bisect.bisect_left(cum, sv))
        j = min(j, n-1)
        L = cum[j] - cum[j-1]
        t = 0.0 if L < 1e-12 else (sv - cum[j-1])/L
        return (pts[j-1][0] + (pts[j][0]-pts[j-1][0])*t, pts[j-1][1] + (pts[j][1]-pts[j-1][1])*t)
    win = 0.6*W
    turns = []
    for i in range(1, n-1):
        if cum[i] < 0.3*W or cum[-1] - cum[i] < 0.3*W:
            continue
        a = at(cum[i] - win); b = at(cum[i] + win); c = pts[i]
        d1 = _norm((c[0]-a[0], c[1]-a[1])); d2 = _norm((b[0]-c[0], b[1]-c[1]))
        ang = math.acos(max(-1.0, min(1.0, d1[0]*d2[0] + d1[1]*d2[1])))
        turns.append((ang, i))
    # локальные максимумы поворота > 35 град, без перекрытия
    turns.sort(reverse=True)
    used = []; cuts = []
    for ang, i in turns:
        if ang < math.radians(35):
            break
        if any(abs(cum[i] - cum[j]) < 1.2*W for j in used):
            continue
        used.append(i)
        R = kr*W
        c = min(R*math.tan(min(ang, math.radians(150))/2), 0.45*(cum[i]), 0.45*(cum[-1]-cum[i]), 1.2*W)
        cuts.append((cum[i]-c, cum[i]+c, i))
    if not cuts:
        return pts
    cuts.sort()
    out = []; k = 0
    for s0, s1, i in cuts:
        while k < n and cum[k] < s0:
            out.append(pts[k]); k += 1
        A = at(s0); B = at(s1); Cn = pts[i]
        m = max(4, int((s1-s0)/(0.08*W)))
        for t in range(m+1):
            u = t/m
            out.append(((1-u)**2*A[0] + 2*u*(1-u)*Cn[0] + u*u*B[0], (1-u)**2*A[1] + 2*u*(1-u)*Cn[1] + u*u*B[1]))
        while k < n and cum[k] <= s1:
            k += 1
    while k < n:
        out.append(pts[k]); k += 1
    return out


def _prune_ears(G, kear):
    """Режим образца: висячая ветка короче kear*W, продолжающая одну из веток развилки
    почти по прямой (выступ стойки за место вливания дуги), удаляется."""
    adj, xy = G['adj'], G['xy']
    W_ = G.get('W', 1.0)
    changed = False
    for lf in sorted((i for i in adj if len(adj[i]) == 1), key=LEAF_KEY(G)):
        if lf not in adj or len(adj[lf]) != 1:
            continue
        path, L, reach, cur = _leaf_reach(G, lf)
        if len(adj[cur]) != 3 or L >= kear * W_ or len(path) < 2:
            continue
        jw = _jwidth(G, cur)
        dl = C._branch_dir(adj, xy, cur, path[-2], jw)[0]
        others = [o for o in adj[cur] if o != path[-2]]
        dots = [(lambda dd: dd[0]*dl[0] + dd[1]*dl[1])(C._branch_dir(adj, xy, cur, o, jw)[0]) for o in others]
        # соосна с одной веткой и не соосна с другой (иначе это не ушко, а перекрёсток/Т)
        if min(dots) > -0.75 or max(dots) < -0.9:
            continue
        # остальные две ветки — обе не висячие короткие (не срезать всё подряд)
        for k in path[:-1]:
            _kill(adj, k)
        changed = True
    return changed


def _contract(G, path):
    adj, xy, kind, r = G['adj'], G['xy'], G['kind'], G['r']
    a, b = path[0], path[-1]
    nid = len(xy); xy.append(((xy[a][0]+xy[b][0])/2, (xy[a][1]+xy[b][1])/2))
    kind.append(1); r.append(1e9)
    ext = set(q for p in path for q in adj[p] if q not in path)
    for p in path:
        _kill(adj, p)
    adj[nid] = set(ext)
    for q in ext:
        adj[q].add(nid)


def _walk_to_junction(G, j, n0, maxd):
    adj, xy = G['adj'], G['xy']
    path = [j]; prev, cur = j, n0; L = math.dist(xy[j], xy[n0])
    while len(adj[cur]) == 2 and L <= maxd:
        nb = [k for k in adj[cur] if k != prev][0]
        path.append(cur); prev, cur = cur, nb
        L += math.dist(xy[prev], xy[cur])
    return path, prev, cur, L


def _merge_junctions(G, kmerge):
    adj = G['adj']
    while True:
        done = False
        for j in list(adj):
            if j not in adj or len(adj[j]) < 3:
                continue
            maxd = kmerge * _jwidth(G, j)
            for n0 in list(adj[j]):
                path, prev, cur, L = _walk_to_junction(G, j, n0, maxd)
                if cur != j and len(adj[cur]) >= 3 and L <= maxd:
                    _contract(G, path + [cur]); done = True; break
            if done:
                break
        if not done:
            return


def _merge_x(G, kmerge, ang):
    """Две тройные развилки на перемычке короче kmerge*jw, у которых внешние ветки
    попарно (по одной от каждой развилки) идут почти по прямой (излом < ang) и пары
    пересекаются, — стягиваются в одну развилку-крест."""
    import centerline as C
    adj, xy = G['adj'], G['xy']
    lim = math.cos(math.radians(ang))
    while True:
        done = False
        for j in sorted(adj, key=LEAF_KEY(G)):
            if j not in adj or len(adj[j]) != 3:
                continue
            wj = _jwidth(G, j)
            for n0 in sorted(adj[j], key=LEAF_KEY(G)):
                path, prev, cur, L = _walk_to_junction(G, j, n0, kmerge*wj)
                if cur == j or len(adj[cur]) != 3 or L > kmerge*wj:
                    continue
                wc = _jwidth(G, cur)
                A = [o for o in adj[j] if o != n0]
                B = [o for o in adj[cur] if o != prev]
                if len(A) != 2 or len(B) != 2:
                    continue
                pa = {a: C._branch_dir(adj, xy, j, a, wj) for a in A}
                pb = {b: C._branch_dir(adj, xy, cur, b, wc) for b in B}
                ok = False
                if _DBG is not None:
                    _DBG.append(('x', xy[j], xy[cur], L/wj, [[round(-(pa[a][0][0]*pb[b][0][0] + pa[a][0][1]*pb[b][0][1]), 2) for b in B] for a in A]))
                for b0, b1 in ((B[0], B[1]), (B[1], B[0])):
                    pairs = ((A[0], b0), (A[1], b1))
                    if all(-(pa[a][0][0]*pb[b][0][0] + pa[a][0][1]*pb[b][0][1]) >= lim for a, b in pairs) and \
                            C._segs_cross(pa[A[0]][1], pb[b0][1], pa[A[1]][1], pb[b1][1]):
                        ok = True
                if ok:
                    _contract(G, path + [cur]); done = True
                    break
            if done:
                break
        if not done:
            return


def _merge_bridges(G, merge_ang):
    adj, xy = G['adj'], G['xy']
    while True:
        jw = {j: _jwidth(G, j) for j in adj if len(adj[j]) >= 3}
        pt = _pairing(adj, xy, merge_ang, 0, jw)
        done = False
        for j in list(adj):
            if len(adj[j]) < 3:
                continue
            for n0 in list(adj[j]):
                # перемычка может быть и «сквозной» парой (X под острым углом: два узла
                # степени 3 на 0.5jw и мостик между ними) — решает проверка креста ниже
                path, prev, cur, L = _walk_to_junction(G, j, n0, 2.5 * jw[j])
                if cur == j or len(adj[cur]) < 3 or L > 2.5 * jw[j]:
                    continue
                # стягиваем, только если после этого ВСЕ ветки проходят насквозь попарно
                # и пары образуют настоящий крест (отрезки проб пересекаются)
                import centerline as _C
                ds = [_C._branch_dir(adj, xy, j, o, jw[j]) for o in adj[j] if o != n0] + \
                     [_C._branch_dir(adj, xy, cur, o, jw[j]) for o in adj[cur] if o != prev]
                if not _all_paired(ds, merge_ang):
                    continue
                _contract(G, path + [cur]); done = True; break
            if done:
                break
        if not done:
            return


def _drop_small(G):
    adj, xy, r, kind = G['adj'], G['xy'], G['r'], G['kind']
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
        L = sum(math.dist(xy[i], xy[j]) for i in comp for j in adj[i]) / 2
        mc = max([r[i] for i in comp if kind[i] == 0] or [0])
        if L < mc:
            for i in comp:
                adj.pop(i, None)


def _chains(G, merge_ang, tool, ref=False):
    import centerline as C
    adj, xy = G['adj'], G['xy']
    jw = {j: _jwidth(G, j) for j in adj if len(adj[j]) >= 3}
    G['jw'] = jw
    # шрифт — образец: направления веток меряются за пределами зоны развилки (1.2..3.2 jw),
    # где срединная ось ещё не искажена острым стыком (стойка «ж» у пазухи с диагональю)
    pt = C._pairing(adj, xy, merge_ang, 0, jw, _branch_dir_far if ref else None, _pair_by_fit if ref else None)
    G['cusp'] = _cusps(G, pt, jw, tool)
    return C._chains(adj, xy, merge_ang, 0, pt)


def _far_pts(adj, xy, i, j, w, lo=0.8, hi=3.0):
    """Узлы ветки развилки i через соседа j на lo..hi w от развилки (у свободного конца —
    без последних 0.7w, крючок торца)."""
    prev, cur = i, j; L = math.dist(xy[i], xy[j])
    P = [(xy[cur], L)]; leaf = len(adj[cur]) == 1
    while L < hi*w and len(adj[cur]) == 2:
        nb = [k for k in adj[cur] if k != prev]
        if not nb:
            break
        prev, cur = cur, nb[0]; L += math.dist(xy[prev], xy[cur])
        P.append((xy[cur], L)); leaf = len(adj[cur]) == 1
    top = hi*w
    if not leaf and len(adj[cur]) >= 3:
        # ветка упёрлась в другую развилку: её зона (0.8w) тоже искажена — не берём
        top = min(top, P[-1][1] - 0.8*w)
    if leaf:
        # короткая ветка к свободному концу (верх стойки «ж» над диагональю): окно сдвигаем
        # ближе к развилке, но не ближе 0.4w, и без последних 0.5w (крючок торца)
        top = min(top, P[-1][1] - 0.5*w)
        lo = max(0.4, min(lo, top/w - 0.6))
    return [p for p, l in P if lo*w <= l <= top]


def _pair_by_fit(adj, xy, i, nb, w, pr):
    """Режим образца: кандидаты сквозных пар (score, a, b), лучшие первыми; score сравнивается
    с cos(merge_ang) в _pairing. Пара сквозная, если
      - дальние участки веток (0.8..3w от развилки) вместе ложатся на одну прямую
        (rms < 0.06w): score = 1 - rms/w;
      - или на одну дугу: score = 0.9 - rms/w;
      - или направления веток (по дальним окнам) почти встречные и ветки целятся друг в
        друга (как в обычном режиме): score = cos(излом) - 0.08.
    Прямые продолжения сильнее дуг и просто «похожих направлений»: у острого стыка
    (стойка «ж» с диагональю) направление короткой ветки искажено, а прямая через оба
    дальних участка — нет."""
    import centerline as C
    fp = {j: _far_pts(adj, xy, i, j, w) for j in nb}
    prn = {j: C._branch_dir(adj, xy, i, j, w) for j in nb}
    cand = []
    for a in range(len(nb)):
        for b in range(a+1, len(nb)):
            A, B = fp[nb[a]], fp[nb[b]]
            da, db = pr[nb[a]][0], pr[nb[b]][0]
            dot = -(da[0]*db[0] + da[1]*db[1])
            dn1, dn2 = prn[nb[a]][0], prn[nb[b]][0]
            dotn = -(dn1[0]*dn2[0] + dn1[1]*dn2[1])
            score = None
            if len(A) >= 2 and len(B) >= 2:
                pts = A + B
                fl = _fit_line(pts); rl = fl[4]
                dx, dy = fl[2], fl[3]
                pa_ = [p[0]*dx + p[1]*dy for p in A]; pb_ = [p[0]*dx + p[1]*dy for p in B]
                apart = max(pa_) <= min(pb_) + 0.2*w or max(pb_) <= min(pa_) + 0.2*w
                if apart and rl <= 0.08*w:
                    score = 1.0 - rl/w
                elif apart and len(pts) >= 5 and max(dot, dotn) >= 0.64:
                    # дуга — только при изломе направлений < 50 град: маленькая дуга легко
                    # «склеивает» стойку с отходящей от неё диагональю («т»)
                    c = C._fit_circle_pts(pts)
                    if c is not None and c[2] > 1.5*w and c[3] <= 0.06*w:
                        score = 0.9 - c[3]/w
                        # штрих, подходящий к стойке сбоку по касательной (выход «т» в стойку
                        # «ь»): дальние участки не целятся друг в друга — пара слабее, чем
                        # сама стойка (обе половины целятся)
                        if not (C._aimed(pr[nb[a]], pr[nb[b]], w) and C._aimed(pr[nb[b]], pr[nb[a]], w)):
                            score -= 0.03
            if score is None:
                # направления: по дальним окнам или по обычным (касательная у входа в развилку:
                # дуга, вливающаяся в стойку, — «т», «м»); ветки должны целиться друг в друга.
                # Только если подгонка не дала оценки: у пары с дугой оценка по дуге
                for pp in (pr, prn):
                    da2, db2 = pp[nb[a]][0], pp[nb[b]][0]
                    dot2 = -(da2[0]*db2[0] + da2[1]*db2[1])
                    am1 = C._aimed(pp[nb[a]], pp[nb[b]], w); am2 = C._aimed(pp[nb[b]], pp[nb[a]], w)
                    if am1 and am2:
                        if score is None or dot2 - 0.08 > score:
                            score = dot2 - 0.08
                    elif (am1 or am2) and dot2 >= 0.9:
                        # почти соосные (< 26 град), целится только одна: стойка «а» с изломом
                        # в выход — слабее, но сквозная (параллельные смещённые ветки — ножка и
                        # пятка «Д» — не целятся ни одной)
                        if score is None or dot2 - 0.15 > score:
                            score = dot2 - 0.15
            if score is not None:
                cand.append((score, nb[a], nb[b], dot))
    cand.sort(reverse=True)
    if _DBG is not None:
        _DBG.append(('pair', xy[i], [(round(s_, 3), a, b, round(d, 2)) for s_, a, b, d in cand]))
    return [(s_, a, b) for s_, a, b, d in cand]


def _branch_dir_far(adj, xy, i, j, w):
    """Направление ветки развилки i через соседа j по дальнему окну 1.2..3.2w от развилки
    (прямая или дуга по узлам окна); у свободного конца последние 0.7w (крючок торца)
    отбрасываются. Короткая ветка — обычным способом. Формат как у C._branch_dir:
    (единичный вектор наружу, дальняя точка, ближняя точка окна, она же, R или None)."""
    import centerline as C
    prev, cur = i, j; L = math.dist(xy[i], xy[j])
    P = [(xy[cur], L)]
    leaf = len(adj[cur]) == 1
    while L < 3.2*w and len(adj[cur]) == 2:
        nb = [k for k in adj[cur] if k != prev]
        if not nb:
            break
        prev, cur = cur, nb[0]; L += math.dist(xy[prev], xy[cur])
        P.append((xy[cur], L))
        leaf = len(adj[cur]) == 1
    Lend = P[-1][1]
    hi = 3.2*w
    if leaf:
        hi = min(hi, Lend - 0.7*w)
    W_ = [p for p, l in P if 1.2*w <= l <= hi]
    if len(W_) < 3 or math.dist(W_[0], W_[-1]) < 0.6*w:
        return C._branch_dir(adj, xy, i, j, w)
    near = W_[0]; far = W_[-1]
    chord = _norm((far[0]-near[0], far[1]-near[1]))
    if len(W_) >= 4:
        c = C._fit_circle_pts(W_)
        if c is not None and c[2] < 25*w and c[3] < 0.03*w and math.dist(near, far)**2/(8*c[2]) > 0.03*w:
            cx, cy, R = c[0], c[1], c[2]
            t = _norm((-(near[1]-cy), near[0]-cx))
            if t[0]*chord[0] + t[1]*chord[1] < 0:
                t = (-t[0], -t[1])
            return t, far, near, near, R
    fl = _fit_line(W_)
    d = (fl[2], fl[3])
    if d[0]*chord[0] + d[1]*chord[1] < 0:
        d = (-d[0], -d[1])
    return d, far, near, near, None


def _cusps(G, pt, jw, tool):
    """V-образный "возврат" пера (острый низ у м/т/п в прописи): развилка, где
    две ветки уходят вверх, а третья — короткий усик к острию. Делаем шпильку:
    ветка A -> острие -> ветка B."""
    adj, xy = G['adj'], G['xy']
    cusp = {}
    for j in [i for i in adj if len(adj[i]) == 3]:
        nb = list(adj[j])
        best = None
        for s in nb:
            # дойти до листа
            prev, cur = j, s; path = [s]; L = math.dist(xy[j], xy[s])
            while len(adj[cur]) == 2:
                nx = [k for k in adj[cur] if k != prev][0]
                L += math.dist(xy[cur], xy[nx]); prev, cur = cur, nx; path.append(cur)
            if len(adj[cur]) != 1:
                continue
            d = _norm((xy[cur][0]-xy[j][0], xy[cur][1]-xy[j][1]))
            D = _raycast_union(G['polys'], xy[cur], d) or 0.0
            if L + D > 2.5 * jw[j]:
                continue
            others = [o for o in nb if o != s]
            ok = True
            ds = []
            for o in others:
                dd = _dir_from(G, j, o, 2.5 * jw[j])
                ds.append(dd)
                if dd[0]*d[0] + dd[1]*d[1] > 0.2:
                    ok = False
            # остриё — "возврат": усик направлен против суммы двух других веток
            # (у Т-стыка ветки противоположны, сумма ~0 — это не остриё)
            if ok:
                sx, sy = ds[0][0]+ds[1][0], ds[0][1]+ds[1][1]
                sl = math.hypot(sx, sy)
                if sl < 0.6 or (sx*d[0] + sy*d[1])/sl > -0.7:
                    ok = False
            if ok:
                best = (s, path, d, others)
                break
        if not best:
            continue
        s, path, d, others = best
        Dj = _raycast_union(G['polys'], xy[j], d) or 0.0
        # остриё должно упираться в острый угол контура (скруглённый торец — не остриё)
        hit = (xy[j][0] + d[0]*Dj, xy[j][1] + d[1]*Dj)
        if not any(c[2] and math.dist(c[:2], hit) < 0.6 * jw[j] for c in G['bcorners']):
            continue
        t = max(0.0, Dj - tool / 2)
        cusp[j] = (xy[j][0] + d[0]*t, xy[j][1] + d[1]*t)
        for k in path:
            _kill(adj, k)
        a, b = others
        pt[j] = {a: b, b: a}
    return cusp


def _dir_from(G, i, j, dlen):
    adj, xy = G['adj'], G['xy']
    prev, cur = i, j; L = math.dist(xy[i], xy[j])
    while L < dlen and len(adj[cur]) == 2:
        nx = [k for k in adj[cur] if k != prev][0]
        prev, cur = cur, nx; L += math.dist(xy[prev], xy[cur])
    return _norm((xy[cur][0]-xy[i][0], xy[cur][1]-xy[i][1]))


def _end_trim(G, ids, k_trim):
    """Срезать хвост у свободного конца ids[-1] (зона скругления/торца), пока
    хорда быстро растёт, но не глубже ~0.6 ширины штриха."""
    xy, r, kind = G['xy'], G['r'], G['kind']
    k = len(ids) - 1; cut = 0.0; cmax = 0.0
    while k > 1 and kind[ids[k]] == 0:
        c = r[ids[k]]
        m = k; L = 0.0
        while m > 0 and L < 0.5 * c:
            L += math.dist(xy[ids[m]], xy[ids[m-1]]); m -= 1
        if kind[ids[m]] != 0 or r[ids[m]] * k_trim <= c:
            break
        cmax = max(cmax, r[ids[m]])
        step = math.dist(xy[ids[k]], xy[ids[k-1]])
        if cut + step > 0.6 * cmax:
            break
        cut += step
        k -= 1
    return ids[:k+1]


STUBZ = 0.5


def _process(G, chains, passthru, prm, W, h):
    adj, xy, r, kind, polys = G['adj'], G['xy'], G['r'], G['kind'], G['polys']
    jw = G['jw']; cusp = G['cusp']
    thru_of = {}
    for ci, (ch, closed) in enumerate(chains):
        for i in (ch if closed else ch[1:-1]):
            if i in passthru and passthru[i]:
                thru_of[i] = ci
    data = []
    for ci, (ch, closed) in enumerate(chains):
        stub = [None, None]; free = [False, False]
        tail_drop0 = False; tail_drop1 = False
        if not closed:
            for end in (0, 1):
                e = ch[0] if end == 0 else ch[-1]
                if len(adj[e]) == 1: free[end] = True
                elif len(adj[e]) >= 3: stub[end] = e
            for end in (0, 1):
                if free[end]:
                    ids = ch if end == 1 else ch[::-1]
                    ids = _end_trim(G, ids, prm['end_trim'])
                    ch = ids if end == 1 else ids[::-1]
        chords = sorted(r[i] for i in ch if kind[i] == 0) or [W]
        cref = chords[len(chords)//2]
        touched = [i for i in ch if len(adj.get(i, ())) >= 3 or i in cusp]
        n = len(ch)
        # дуговая длина вдоль цепочки
        arc = [0.0]
        for a_, b_ in zip(ch[:-1], ch[1:]):
            arc.append(arc[-1] + math.dist(xy[a_], xy[b_]))
        total = arc[-1] + (math.dist(xy[ch[-1]], xy[ch[0]]) if closed else 0.0)
        spec = [k for k in range(n) if len(adj.get(ch[k], ())) >= 3 or ch[k] in cusp]
        hard = [False]*n; flags = [False]*n
        for k in spec:
            hard[k] = True; flags[k] = True
        # зоны вокруг особых узлов: радиус = min(jzone*jw, треть длины участка до соседнего
        # особого узла / конца цепочки) — короткие перемычки не должны исчезать целиком
        def run_len(k, dirn):
            # длина участка цепочки от особого узла k в сторону dirn до следующего особого узла/конца
            L = 0.0; q = k; cnt = 0
            while True:
                q2 = q + dirn
                if closed:
                    q2 %= n
                elif q2 < 0 or q2 >= n:
                    return L
                L += math.dist(xy[ch[q]], xy[ch[q2]]); q = q2; cnt += 1
                if hard[q] or cnt >= n:
                    return L
        for k in spec:
            j = ch[k]
            zj = prm['jzone'] * jw.get(j, cref) * (0.5 if j in cusp else 1.0)
            sz = 1.8 * jw.get(j, cref)
            if not closed and (k == 0 or k == n-1) and j not in cusp:
                # конец цепочки — заглушка: её ось у развилки искажена меньше, чем сквозная
                # (нет «проседания»), а короткой перекладине (низ чаши P) нужны свои точки
                zj = STUBZ * jw.get(j, cref); sz = min(sz, 1.0 * jw.get(j, cref))
                if prm.get('ref'):
                    # шрифт — образец: ось заглушки у острого стыка искажена на ~1.5 ширины
                    # (уходит в пазуху) — отбрасываем, штрих достраивается по своему ходу
                    zj = prm.get('stubz', 1.5) * jw.get(j, cref); sz = max(sz, zj)
            for dirn in (-1, 1):
                Lr = run_len(k, dirn)
                cap = Lr / 3.0
                zh = min(zj, cap); zs = min(sz, cap)
                L = 0.0; q = k
                while True:
                    q2 = q + dirn
                    if closed:
                        q2 %= n
                    elif q2 < 0 or q2 >= n:
                        break
                    L += math.dist(xy[ch[q]], xy[ch[q2]]); q = q2
                    if hard[q] and q in spec:
                        break
                    if L >= zs:
                        break
                    if L < zh:
                        hard[q] = True; flags[q] = True
                    elif j not in cusp and r[ch[q]] > prm['widen']*cref:
                        flags[q] = True
        zone_ok = True
        # короткий остаток (< 1 jw) между зоной развилки и свободным концом — выбрасываем:
        # конец достроится продолжением оси сквозной линии (торец ножки у N, 4, k)
        for end in (0, 1):
            if closed or not free[end]:
                continue
            rng = range(n-1, -1, -1) if end == 1 else range(n)
            run = []; L = 0.0; prev = None
            for q in rng:
                if hard[q]:
                    break
                if prev is not None:
                    L += math.dist(xy[ch[q]], xy[ch[prev]])
                run.append(q); prev = q
            else:
                continue            # зоны нет вовсе
            jq = ch[q]
            # лист-заглушка целиком (хвост Q, короткий отросток): зона упёрлась в развилку
            # на другом конце цепочки — это не «колпачок», выбрасывать нельзя
            if (q == 0 or q == n-1) and len(adj.get(jq, ())) >= 3:
                continue
            # и только если за развилкой цепочка продолжается далеко (есть чем продолжить ось);
            # короткий сквозной хвостик Q (обе стороны ~1.5jw) оставляем как есть
            if total - L < 3.0 * jw.get(jq, cref):
                continue
            # «колпачок» — это торец штриха-заглушки (ножка N торчит над диагональю): заглушка
            # должна смотреть ровно против остатка. Иначе (рукописная n: штрих сам идёт дальше
            # за развилку и слегка гнётся) остаток — настоящая ось, его нельзя выбрасывать
            js = None
            for q2 in (range(q, -1, -1) if end == 1 else range(q, n)):
                if ch[q2] in passthru and len(adj.get(ch[q2], ())) == 3:
                    js = ch[q2]; break
                if not hard[q2]:
                    break
            if js is None or not passthru.get(js):
                continue
            pair = passthru[js]
            others = [o for o in adj[js] if o not in pair]
            if len(others) != 1:
                continue
            sd = C._branch_dir(adj, xy, js, others[0], jw.get(js, cref))[0]
            e = ch[n-1] if end == 1 else ch[0]
            ed = _norm((xy[e][0]-xy[js][0], xy[e][1]-xy[js][1]))
            if sd[0]*ed[0] + sd[1]*ed[1] > -0.85:
                continue
            if run and L < 1.0 * jw.get(jq, cref):
                for q2 in run:
                    hard[q2] = True; flags[q2] = True
                if end == 1:
                    tail_drop1 = True
                else:
                    tail_drop0 = True
        if closed and all(flags):
            continue
        if not closed and all(flags):
            continue
        if closed and flags[0]:
            r0 = flags.index(False); ch = ch[r0:]+ch[:r0]; flags = flags[r0:]+flags[:r0]
        pts = []; k = 0; fixed = set(); pend = None
        tail_drop0 = locals().get('tail_drop0', False); tail_drop1 = locals().get('tail_drop1', False)
        while k < n:
            if not flags[k]:
                q = xy[ch[k]]
                if pend is not None:
                    if math.dist(q, pend[1]) <= pend[2]:
                        q = pend[0](q)
                    else:
                        pend = None
                pts.append(q); k += 1; continue
            k2 = k
            while k2 < n and flags[k2]:
                k2 += 1
            jc = [ch[q] for q in range(k, k2) if len(passthru.get(ch[q], {})) >= 4]
            jc = xy[jc[0]] if jc else None
            cp = [cusp[ch[q]] for q in range(k, k2) if ch[q] in cusp]
            fl = None
            if cp and pts and (k2 < n or closed):
                fl = [cp[0]]                       # шпилька: прямо в острие и обратно
            elif pts and k2 < n:
                aft = [xy[i] for i in ch[k2:]]
                fl = None
                # сначала фит одной кривой/прямой через обе стороны (и для перекрестий:
                # прямая диагональ X должна остаться прямой), затем угол, затем Эрмит
                raw = [xy[ch[q]] for q in range(k, k2)]
                fl, proj, xi = _arc_fill(pts, aft, h*1.4, cref, polys, G['bcorners'], raw)
                if proj is not None:
                    pend = (proj, xy[ch[k2]], 0.8*cref)
                if xi is not None:
                    fixed.add(len(pts) + xi)        # угол оси в развилке — не сглаживать
                if fl is None:
                    fl = _hermite_fill(pts, aft, h*1.4, jc)
            elif pts and closed:
                fl = _hermite_fill(pts, pts[:], h*1.4, jc)
            if fl is None and not closed and k2 == n and free[1] and pts and not tail_drop1:
                # зона развилки дошла до свободного конца: тянемся к самому концу
                fl = [xy[ch[n-1]]]
            if fl is not None:
                nxt = xy[ch[k2]] if k2 < n else (pts[0] if closed else None)
                if not _poly_inside(polys, [pts[-1]] + fl + ([nxt] if nxt else []), 0.5*h):
                    fl = [xy[ch[q]] for q in range(k, k2)]   # обход вылез за контур — берём сырые узлы
                elif cp:
                    fixed.add(len(pts))
                pts.extend(fl)
            k = k2
        if not closed and free[0] and flags and flags[0] and pts and not tail_drop0:
            q = 0
            while q < n and flags[q]:
                q += 1
            if _poly_inside(polys, [xy[ch[0]], pts[0]], 0.5*h):
                pts.insert(0, xy[ch[0]])
            else:
                pts[0:0] = [xy[ch[t]] for t in range(q)]
        if len(pts) < 2:
            if not closed and (free[0] or free[1]) and (stub[0] is not None or stub[1] is not None):
                js = stub[0] if stub[0] is not None else stub[1]
                lf = ch[-1] if free[1] else ch[0]
                pts = [xy[js], xy[lf]] if free[1] else [xy[lf], xy[js]]
                stub = [None, None]
            elif not closed and stub[0] is not None and stub[1] is not None and stub[0] != stub[1]:
                # короткая перемычка между двумя развилками (к, ж): прямой отрезок
                pts = [xy[stub[0]], xy[stub[1]]]
                stub = [None, None]
            else:
                continue
        pts = _smooth(pts, closed, prm['smooth_it'], fixed)
        pts, fixed = _sharpen(pts, fixed, closed, cref, polys, G['bcorners'], h)
        pts = _merge_collinear(pts, fixed, closed, cref)
        def nmax_of(end):
            # точки от конца до ближайшего острого угла (включительно) — дальше не смотрим
            if not fixed:
                return None
            return (len(pts) - max(fixed)) if end == 1 else (min(fixed) + 1)
        if not closed:
            for end in (0, 1):
                if not free[end]:
                    continue
                q = pts if end == 1 else pts[::-1]
                if prm.get('ref') == 1:
                    _ref_end(q, cref, polys, prm['tool'], h)
                else:
                    _free_end(q, cref, polys, 2*prm['end_off'] if prm.get('end_off') is not None else prm['tool'], h, nmax_of(end), G['bcorners'])
                pts = q if end == 1 else q[::-1]
        data.append([pts, closed, stub, ci, cref, nmax_of(0), nmax_of(1)])
    ci2pts = {d[3]: d[0] for d in data}
    deferred = []
    for d in data:
        pts, closed, stub, ci, cref = d[:5]
        for end in (0, 1):
            j = stub[end]
            if j is None:
                continue
            q = pts if end == 1 else pts[::-1]
            poly = None; v0 = v1 = 0
            if j in thru_of and thru_of[j] in ci2pts and passthru.get(j):
                # сквозная линия развилки: кусок обработанной ломаной сквозной цепочки
                # между её узлами-соседями развилки (+2jw в обе стороны), а не вся цепочка
                P = ci2pts[thru_of[j]]
                a = next(iter(passthru[j])); b = passthru[j][a]
                ia = min(range(len(P)), key=lambda t: math.dist(P[t], xy[a]))
                ib = min(range(len(P)), key=lambda t: math.dist(P[t], xy[b]))
                lo, hi = min(ia, ib), max(ia, ib)
                L = 0.0
                while lo > 0 and L < 2*jw[j]:
                    L += math.dist(P[lo], P[lo-1]); lo -= 1
                L = 0.0
                while hi < len(P)-1 and L < 2*jw[j]:
                    L += math.dist(P[hi], P[hi+1]); hi += 1
                poly = P[lo:hi+1]
                # у свободных концов сквозной линии — виртуальное продолжение на 1jw:
                # заглушка, приходящая ровно в торец (диагональ «4» в верх ножки), не должна
                # промахнуться мимо конца ломаной из-за округления
                v0 = v1 = 0
                if lo == 0 and len(poly) >= 2:
                    dv = _norm((poly[0][0]-poly[1][0], poly[0][1]-poly[1][1]))
                    poly.insert(0, (poly[0][0] + dv[0]*0.5*jw[j], poly[0][1] + dv[1]*0.5*jw[j])); v0 = 1
                if hi == len(P)-1 and len(poly) >= 2:
                    dv = _norm((poly[-1][0]-poly[-2][0], poly[-1][1]-poly[-2][1]))
                    poly.append((poly[-1][0] + dv[0]*0.5*jw[j], poly[-1][1] + dv[1]*0.5*jw[j])); v1 = 1
            e0 = q[-1]; qsave = list(q)
            ok = False
            if prm.get('ref'):
                # шрифт — образец: штрих, отходящий от стойки, — прямой по своему же ходу до
                # сквозной линии (диагональ к дуге «т», «м»), без S-изгиба к центру развилки
                ok = _ref_stub(q, cref, poly, jw[j], xy[j], h, polys, v0, v1)
            if not ok:
                ok = _stub_extend(q, cref, poly, jw[j], xy[j], h, d[5 + end], polys, v0, v1, G['bcorners'])
            if ok and thru_of.get(j) == ci and math.dist(q[-1], e0) < 0.05*cref:
                # сквозная — эта же цепочка: «попадание» в собственный конец (з: диагональ к ножке)
                # не продление; откат и запасной вариант без сквозной
                q[:] = qsave; poly = None; ok = False
            if not ok and prm.get('ref'):
                # шрифт — образец: сквозной линии не достать — отложить, после всех заглушек
                # попробовать упереться в любую другую линию у развилки (чаша «с» в диагональ «ж»)
                deferred.append((d, end, j, poly, v0, v1))
                d[0] = q if end == 1 else q[::-1]
                pts = d[0]
                continue
            if not ok:
                # запасной вариант: плавно к точке сквозной линии у центра развилки
                rawc = [xy[i] for i in chains[ci][0]]
                if end == 0:
                    rawc = rawc[::-1]
                _stub_to_junction(q, cref, poly, xy[j], h, polys, v0, v1, jw[j], rawc)
            d[0] = q if end == 1 else q[::-1]
            pts = d[0]
    # отложенные заглушки: по кругу, пока кто-то цепляется (чаша «с» ждёт, когда диагональ
    # «ж» достроится до стойки, и упирается уже в неё)
    while deferred:
        progress = False
        for item in list(deferred):
            d, end, j, poly, v0, v1 = item
            pts, closed, stub, ci, cref = d[:5]
            q = pts if end == 1 else pts[::-1]
            others = []
            for d2 in data:
                if d2 is d:
                    continue
                P2 = d2[0]
                if any(math.dist(p_, xy[j]) <= 3.0*jw[j] for p_ in P2[::max(1, len(P2)//40)]):
                    others.append(P2)
            for P2 in sorted(others, key=lambda P2: min(math.dist(p_, xy[j]) for p_ in P2)):
                if _ref_stub(q, cref, P2, jw[j], xy[j], h, polys, 0, 0):
                    d[0] = q if end == 1 else q[::-1]
                    deferred.remove(item); progress = True
                    break
        if not progress:
            break
    for d, end, j, poly, v0, v1 in deferred:
        pts, closed, stub, ci, cref = d[:5]
        q = pts if end == 1 else pts[::-1]
        rawc = [xy[i] for i in chains[ci][0]]
        if end == 0:
            rawc = rawc[::-1]
        _stub_to_junction(q, cref, poly, xy[j], h, polys, v0, v1, jw[j], rawc)
        d[0] = q if end == 1 else q[::-1]
    return [(d[0], d[1], d[2], d[3]) for d in data]


def _tan(q, w):
    j = len(q)-1; L = 0.0
    while j > 0 and L < w*0.7:
        L += math.dist(q[j], q[j-1]); j -= 1
    return _norm((q[-1][0]-q[j][0], q[-1][1]-q[j][1]))


def _poly_inside(polys, pts, step=1.0):
    """Проверка, что ломаная целиком внутри фигуры (по точкам через ~равный шаг)."""
    for a, b in zip(pts[:-1], pts[1:]):
        L = math.dist(a, b)
        n = max(1, min(64, int(L / step)))
        for i in range(1, n + 1):
            t = i / (n + 1)
            x = a[0] + (b[0]-a[0])*t; y = a[1] + (b[1]-a[1])*t
            if not any(inside(p, x, y) for p in polys):
                return False
    return True


def _all_paired(ds, merge_ang):
    """ds: список (направление, дальняя точка). Все ветки разбиваются на соосные пары,
    и отрезки проб всех пар попарно пересекаются (крест)."""
    import centerline as _C
    if len(ds) % 2:
        return False
    cand = []
    for a in range(len(ds)):
        for b in range(a+1, len(ds)):
            cand.append((-(ds[a][0][0]*ds[b][0][0]+ds[a][0][1]*ds[b][0][1]), a, b))
    cand.sort(reverse=True)
    used = set(); lim = math.cos(math.radians(merge_ang)); segs = []
    for dot, a, b in cand:
        if a in used or b in used:
            continue
        if dot < lim:
            return False
        if not all(_C._segs_cross(ds[a][1], ds[b][1], p, q) for p, q in segs):
            return False
        used |= {a, b}; segs.append((ds[a][1], ds[b][1]))
    return len(used) == len(ds)


def _boundary_corners(polys, W):
    """Острые углы контура (поворот > 40 град на участке +-0.15W): (x, y, convex, e1, e2),
    convex — выпуклый относительно заливки (остриё), иначе вогнутый (внутренний угол);
    e1, e2 — направления хорд контура от угла на 0.5W в обе стороны (кромки штрихов)."""
    out = []
    for p in polys:
        for ri, r in enumerate(p):
            n = len(r)
            ccw = ring_area(r) > 0
            hole = ri > 0
            for i in range(n):
                def walk(step, L):
                    s = 0.0; j = i
                    for _ in range(n):
                        k = (j + step) % n
                        s += math.dist(r[j], r[k]); j = k
                        if s >= L:
                            break
                    return j
                a = walk(-1, 0.15*W); b = walk(+1, 0.15*W)
                d1 = _norm((r[i][0]-r[a][0], r[i][1]-r[a][1]))
                d2 = _norm((r[b][0]-r[i][0], r[b][1]-r[i][1]))
                c = d1[0]*d2[0] + d1[1]*d2[1]
                if c < math.cos(math.radians(40)):
                    left = (d1[0]*d2[1] - d1[1]*d2[0]) > 0
                    convex = (left == ccw) != hole
                    a = walk(-1, 0.5*W); b = walk(+1, 0.5*W)
                    e1 = _norm((r[a][0]-r[i][0], r[a][1]-r[i][1]))
                    e2 = _norm((r[b][0]-r[i][0], r[b][1]-r[i][1]))
                    # прямизна кромки: макс. отклонение контура от хорды на 0.8W от угла (в долях W)
                    def straight(step):
                        s = 0.0; j = i; pts = []
                        for _ in range(n):
                            k = (j + step) % n
                            s += math.dist(r[j], r[k]); j = k; pts.append(r[k])
                            if s >= 0.8*W:
                                break
                        ex, ey = _norm((r[j][0]-r[i][0], r[j][1]-r[i][1]))
                        return max(abs((p[0]-r[i][0])*ey - (p[1]-r[i][1])*ex) for p in pts) / W
                    out.append((r[i][0], r[i][1], convex, e1, e2, straight(-1), straight(+1)))
    return out


def _edge_lines(c, w, polys, idx=False):
    """Гипотезы осей штрихов прямо по контуру у вогнутого угла c: прямая вдоль кромки
    e (хорда 0.5w от угла) на половине местной ширины (луч от кромки поперёк штриха
    до противоположной кромки). Для коротких перекладин (В, Я), где по оси фит не
    получить. Список ('L', mx, my, dx, dy)."""
    out = []
    for ei, e in enumerate((c[3], c[4])):
        for sg in (1.0, -1.0):
            nx, ny = -e[1]*sg, e[0]*sg
            o = (c[0] + e[0]*0.2*w + nx*0.05*w, c[1] + e[1]*0.2*w + ny*0.05*w)
            if not _inside_any(polys, o[0], o[1]):
                continue
            D = _raycast_union(polys, o, (nx, ny))
            if D is None:
                continue
            Wl = D + 0.05*w
            if not (0.6*w <= Wl <= 1.6*w):
                continue
            out.append((ei, ('L', c[0] + nx*Wl/2, c[1] + ny*Wl/2, e[0], e[1])) if idx
                       else ('L', c[0] + nx*Wl/2, c[1] + ny*Wl/2, e[0], e[1]))
            break
    return out


def _edge_match(fit, c, w, e):
    """Ось fit вблизи угла контура c идёт вдоль кромки e (хорда 0.5w): угол < 7 град."""
    p = _fit_proj_c(fit, c[:2])
    cosmin = math.cos(math.radians(7))
    if fit[0] == 'L':
        return abs(fit[3]*e[0] + fit[4]*e[1]) >= cosmin
    _, cx, cy, R = fit
    a0 = math.atan2(p[1]-cy, p[0]-cx)
    for sg in (1.0, -1.0):
        a = a0 + sg*0.5*w/R
        q = (cx + R*math.cos(a), cy + R*math.sin(a))
        d = _norm((q[0]-p[0], q[1]-p[1]))
        if d[0]*e[0] + d[1]*e[1] >= cosmin:
            return True
    return False


# ---------------------------------------------------------------- концы по кривизне

def _fit_end(q, w, nmax=None):
    """q[-1] — конец. Окружность/прямая по последним точкам: окно 2w, 1.2w, 0.7w —
    берётся первое, при котором последние 0.7w точек лежат на фите (остаток < 0.025w).
    ('C', cx, cy, R, sign) | ('L', mx, my, dx, dy) (dx,dy — наружу, к концу) | None."""
    allP = []; s = 0.0; dist_from_end = []
    imin = 0 if nmax is None else max(0, len(q)-nmax)
    for i in range(len(q)-1, imin-1, -1):
        allP.append(q[i]); dist_from_end.append(s)
        if i > 0:
            s += math.dist(q[i], q[i-1])
        if s > 2.0*w:
            break
    end = q[-1]
    tail = [p for p, d in zip(allP, dist_from_end) if d <= 0.7*w]
    for win in (2.0, 1.2, 0.7):
        P = [p for p, d in zip(allP, dist_from_end) if d <= win*w]
        if len(P) < 3:
            continue
        fit = None
        if len(P) >= 4 and win > 0.7:          # на самом коротком окне окружности не верим
            c = _fit_circle(P)
            if c is not None and c[2] < 25*w and c[3] < 0.05*w and _bent(P, c[2], w):
                cx, cy, R, _ = c
                s = 0.0; ref = P[-1]
                for i in range(1, len(P)):
                    s += math.dist(P[i-1], P[i])
                    if s >= 0.5*w:
                        ref = P[i]; break
                a_end = math.atan2(end[1]-cy, end[0]-cx); a_ref = math.atan2(ref[1]-cy, ref[0]-cx)
                dlt = (a_end - a_ref + math.pi) % (2*math.pi) - math.pi
                fit = ('C', cx, cy, R, 1.0 if dlt >= 0 else -1.0)
        if fit is None:
            mx, my, dx, dy, rms = _fit_line(P)
            if rms < 0.08*w:
                if (end[0]-P[-1][0])*dx + (end[1]-P[-1][1])*dy < 0:
                    dx, dy = -dx, -dy
                fit = ('L', mx, my, dx, dy)
        if fit is None:
            continue
        res = max(math.dist(p, _fit_proj(fit, p)) for p in tail) if tail else 0.0
        if res < 0.02*w or win == 0.7:
            return fit
    return None


def _fit_proj(fit, p):
    if fit[0] == 'C':
        _, cx, cy, R, _s = fit
        d = math.dist(p, (cx, cy))
        return (cx + (p[0]-cx)*R/d, cy + (p[1]-cy)*R/d) if d > 1e-9 else p
    _, mx, my, dx, dy = fit
    t = (p[0]-mx)*dx + (p[1]-my)*dy
    return (mx + dx*t, my + dy*t)


def _fit_at(fit, start, s):
    """Точка на расстоянии s вдоль фита от start (start лежит на фите) в сторону конца."""
    if fit[0] == 'C':
        _, cx, cy, R, sg = fit
        a = math.atan2(start[1]-cy, start[0]-cx) + sg*s/R
        return (cx + R*math.cos(a), cy + R*math.sin(a))
    _, mx, my, dx, dy = fit
    return (start[0] + dx*s, start[1] + dy*s)


def _end_project(q, fit, w, nmax=None, L=None):
    """Последние точки (в пределах L, по умолчанию 1w, от конца) — на фит, чтобы убрать вихляние."""
    s = 0.0
    if L is None:
        L = 1.0*w
    imin = 0 if nmax is None else max(0, len(q)-nmax)
    for i in range(len(q)-1, imin-1, -1):
        if i > 0:
            s += math.dist(q[i], q[i-1])
        q[i] = _fit_proj(fit, q[i])
        if s > L:
            break


def _retract(q, back):
    """Укоротить ломаную с конца на длину back."""
    while len(q) > 2 and math.dist(q[-1], q[-2]) <= back:
        back -= math.dist(q[-1], q[-2]); q.pop()
    if back > 0 and len(q) >= 2:
        L = math.dist(q[-1], q[-2])
        if L > 1e-9:
            t = back / L
            q[-1] = (q[-1][0] + (q[-2][0]-q[-1][0])*t, q[-1][1] + (q[-2][1]-q[-1][1])*t)


def _centered(polys, p, d, w):
    """Точка p на оси штриха: расстояния до края поперёк направления d примерно равны."""
    n = (-d[1], d[0])
    dl = _raycast_union(polys, p, n); dr = _raycast_union(polys, p, (-n[0], -n[1]))
    if dl is None or dr is None:
        return False
    # не прижат к стенке; одна сторона может быть «бесконечной» — конец ножки u-образной
    # буквы (ц, u) упирается в сплошное дно, луч поперёк уходит через всё дно
    return min(dl, dr) >= 0.3*w and (abs(dl - dr) <= 0.4*(dl + dr) or max(dl, dr) > 2.0*w)


def _free_end_ray(q, w, polys, tool):
    """Запасной конец: прямая по направлению (в пределах +-30 град от касательной),
    по которому луч из конца идёт дальше всего до края — ось штриха; конец на tool/2."""
    d0 = _tan(q, w)
    a0 = math.atan2(d0[1], d0[0])
    best = None
    for k in range(-6, 7):
        a = a0 + math.radians(5*k)
        d = (math.cos(a), math.sin(a))
        D = _raycast_union(polys, q[-1], d)
        if D is not None and (best is None or D > best[0]):
            best = (D, d)
    if best is None:
        return
    D, d = best
    ext = D - tool/2
    if ext > 0:
        q.append((q[-1][0]+d[0]*ext, q[-1][1]+d[1]*ext))
    else:
        _retract(q, -ext)


def _free_end(q, w, polys, tool, h, nmax=None, bcorners=()):
    """Свободный конец: продолжить ось по её кривизне до края контура и остановиться
    на tool/2 от края (по длине пути); если край ближе tool/2 — отступить назад.
    nmax — сколько точек от конца можно использовать (до ближайшего острого угла).
    Если продолжение уходит в боковую стенку (конец не по центру штриха) — прямая
    по самому длинному лучу."""
    q0 = list(q)
    _free_end_fit(q, w, polys, tool, h, nmax, bcorners)
    if len(q) >= 2:
        d = _norm((q[-1][0]-q[-2][0], q[-1][1]-q[-2][1]))
        if _centered(polys, q[-1], d, w):
            return
    q[:] = q0
    _free_end_ray(q, w, polys, tool)


_DBG = None


def _bdist(polys, p):
    """Расстояние от точки до края контура."""
    best = 1e18
    for poly in polys:
        for r in poly:
            n = len(r)
            for i in range(n):
                a = r[i]; b = r[(i+1) % n]
                dx = b[0]-a[0]; dy = b[1]-a[1]; L2 = dx*dx + dy*dy
                t = 0.0 if L2 < 1e-18 else max(0.0, min(1.0, ((p[0]-a[0])*dx + (p[1]-a[1])*dy)/L2))
                ex = a[0] + dx*t - p[0]; ey = a[1] + dy*t - p[1]
                d = ex*ex + ey*ey
                if d < best:
                    best = d
    return math.sqrt(best)


def _ref_end(q, w, polys, tool, h):
    """Режим «шрифт — образец», до кончика штриха: конец оси у торца искажён формой торца
    (уходит в угол среза, крючок). Срезаем зону торца (до ~0.9 локальной ширины) и продолжаем
    ось по её же ходу до края, минус tool/2. Кандидаты: прямая по участку перед срезом и дуга
    (завиток, шарик); берётся тот, что дальше уходит вдоль штриха."""
    if len(q) < 3:
        return
    # локальная ширина: максимум 2*расстояния до края на последних 2w пути
    wl = 0.0; s = 0.0
    for i in range(len(q)-1, 0, -1):
        wl = max(wl, 2*_bdist(polys, q[i]))
        s += math.dist(q[i], q[i-1])
        if s > 2.0*w:
            break
    if wl <= 0:
        return
    Ltot = sum(math.dist(q[i], q[i+1]) for i in range(len(q)-1))
    q0 = list(q)
    found = {}
    # срез зоны торца: сначала глубокий, потом мельче — пока участок перед ним
    # хорошо ложится на прямую (дугу); иначе срез залез в изгиб штриха
    for kc in (0.9, 0.65, 0.4, 0.2):
        for kw in (1.3, 0.9, 0.6):
            if 'L' in found and 'C' in found:
                break
            cut = min(kc*wl, max(0.0, Ltot - (kw + 0.3)*wl))
            qq = list(q0)
            if cut > 0:
                _retract(qq, cut)
            P = [qq[-1]]; s = 0.0
            for i in range(len(qq)-1, 0, -1):
                s += math.dist(qq[i], qq[i-1])
                P.append(qq[i-1])
                if s >= kw*wl:
                    break
            P = P[::-1]
            if len(P) < 3 or s < 0.4*wl:
                continue
            if 'L' not in found:
                fl = _fit_line(P)
                if fl[4] < 0.03*wl:
                    d = (fl[2], fl[3])
                    if (P[-1][0]-P[0][0])*d[0] + (P[-1][1]-P[0][1])*d[1] < 0:
                        d = (-d[0], -d[1])
                    st = qq[-1]

                    def line_at(t, st=st, d=d):
                        return (st[0] + d[0]*t, st[1] + d[1]*t)
                    found['L'] = (qq, cut, line_at, 3.0*wl)
            if 'C' not in found and len(P) >= 4:
                c = _fit_circle(P)
                if c is not None and c[2] > 0.6*wl and c[3] < 0.03*wl and _bent(P, c[2], wl):
                    cx, cy, R = c[:3]
                    st = qq[-1]
                    a0 = math.atan2(st[1]-cy, st[0]-cx)
                    ap = math.atan2(P[-2][1]-cy, P[-2][0]-cx)
                    sgn = 1.0 if math.sin(a0-ap) > 0 else -1.0

                    def arc_at(t, a0=a0, sgn=sgn, cx=cx, cy=cy, R=R):
                        a = a0 + sgn*t/R
                        return (cx + R*math.cos(a), cy + R*math.sin(a))
                    found['C'] = (qq, cut, arc_at, min(3.0*wl, 2.1*R))
        if 'L' in found and 'C' in found:
            break

    def exit_len(at, tmax):
        step = h/4; t = step
        while t <= tmax:
            p = at(t)
            if not _inside_any(polys, p[0], p[1]):
                lo, hi = t-step, t
                for _ in range(12):
                    m = (lo+hi)/2
                    pm = at(m)
                    if _inside_any(polys, pm[0], pm[1]):
                        lo = m
                    else:
                        hi = m
                return (lo+hi)/2
            t += step
        return None
    # вынос за исходный конец (L - срез); дуга должна выиграть заметно — при равенстве
    # прямая к торцу надёжнее
    best = None
    for kind in ('L', 'C'):
        if kind not in found:
            continue
        qq, cut, at, tmax = found[kind]
        Lx = exit_len(at, tmax)
        if Lx is None:
            continue
        reach = Lx - cut - (0.15*wl if kind == 'C' else 0.0)
        if _DBG is not None:
            _DBG.append((kind, round(Lx, 2), round(cut, 2), round(wl, 2)))
        if best is None or reach > best[0]:
            best = (reach, qq, Lx, at)
    if best is None:
        _free_end(q, w, polys, tool, h)
        return
    _, qq, Lx, at = best
    q[:] = qq
    target = Lx - tool/2
    if target <= 0:
        _retract(q, -target)
        return
    m = int(target/h)
    for i in range(1, m+1):
        q.append(at(i*h))
    if target - m*h > 1e-6 or m == 0:
        q.append(at(target))

def _strip_hook(q, w, nmax=None):
    """Остаток скруглённого торца / крючок на самом конце: точки в последних 0.5w, не лежащие
    (> 0.08w) на прямой по участку 0.5..1.5w от конца, отбрасываются."""
    P = []; s = 0.0
    imin = 1 if nmax is None else max(1, len(q)-nmax+1)
    for i in range(len(q)-1, imin-1, -1):
        s += math.dist(q[i], q[i-1])
        if s > 1.5*w:
            break
        if s >= 0.5*w:
            P.append(q[i-1])
    if len(P) < 3:
        return
    mx, my, dx, dy, rms = _fit_line(P)
    if rms > 0.03*w:
        return
    s = 0.0; k = len(q)
    while k > 2:
        p = q[k-1]
        if abs((p[0]-mx)*dy - (p[1]-my)*dx) <= 0.04*w:
            break
        s += math.dist(q[k-1], q[k-2])
        if s > 0.5*w:
            return
        k -= 1
    del q[k:]


def _free_end_fit(q, w, polys, tool, h, nmax=None, bcorners=()):
    _strip_hook(q, w, nmax)
    nl = _near_line(q, w, nmax, polys, bcorners)
    fit = None
    if nl is not None:
        fit, L = nl
        _end_project(q, fit, w, nmax, L)   # только прямой участок, дальше дуга остаётся дугой
        _blend_arc(q, fit, L, w, nmax)
    else:
        fit = _far_fit(q, w, 1.0, 3.0, nmax)
        if fit is None or not _apply_far(q, w, fit, 1.0, 0.12*w, nmax):
            fit = _fit_end(q, w, nmax)
            if fit is not None:
                _end_project(q, fit, w, nmax)
    if fit is None:
        # запасной вариант: касательная по двум точкам
        d = _tan(q, w)
        D = _raycast_union(polys, q[-1], d)
        if D is None:
            return
        ext = D - tool/2
        if ext > 0:
            q.append((q[-1][0]+d[0]*ext, q[-1][1]+d[1]*ext))
        else:
            _retract(q, -ext)
        return
    start = q[-1]
    step = h/4; L_exit = None; smax = 4.5*w
    k = 1
    while k*step <= smax:
        p = _fit_at(fit, start, k*step)
        if not _inside_any(polys, p[0], p[1]):
            lo, hi = (k-1)*step, k*step
            for _ in range(10):
                mid = (lo+hi)/2
                pm = _fit_at(fit, start, mid)
                if _inside_any(polys, pm[0], pm[1]):
                    lo = mid
                else:
                    hi = mid
            L_exit = (lo+hi)/2
            break
        k += 1
    if L_exit is None:
        return
    target = L_exit - tool/2
    if target >= 0:
        m = int(target/h)
        for i in range(1, m+1):
            q.append(_fit_at(fit, start, i*h))
        if target - m*h > 1e-6 or m == 0:
            q.append(_fit_at(fit, start, target))
    else:
        _retract(q, -target)


def _seg_hit(poly, a, b):
    """Пересечение отрезка a-b с ломаной poly: ближайшая к a точка или None."""
    best = None; bt = 2.0
    ex, ey = b[0]-a[0], b[1]-a[1]
    for i in range(len(poly)-1):
        px, py = poly[i]; qx, qy = poly[i+1][0]-px, poly[i+1][1]-py
        den = ex*qy - ey*qx
        if abs(den) < 1e-12:
            continue
        t = ((px-a[0])*qy - (py-a[1])*qx) / den
        u = ((px-a[0])*ey - (py-a[1])*ex) / den
        if -1e-9 <= t <= 1+1e-9 and -1e-9 <= u <= 1+1e-9 and t < bt:
            bt = t; best = (a[0]+ex*t, a[1]+ey*t)
    return best


def _poly_closest(poly, p):
    """Ближайшая точка ломаной к p: (расстояние, точка)."""
    best = None; bd = float('inf')
    for i in range(len(poly)-1):
        ax, ay = poly[i]; bx, by = poly[i+1]
        ex, ey = bx-ax, by-ay; L2 = ex*ex + ey*ey
        t = 0.0 if L2 < 1e-18 else max(0.0, min(1.0, ((p[0]-ax)*ex + (p[1]-ay)*ey)/L2))
        c = (ax + ex*t, ay + ey*t); d = math.dist(p, c)
        if d < bd:
            bd = d; best = c
    return bd, best


def _far_fit(q, w, near, far, nmax=None):
    """Фит оси по точкам на near..far w от конца q[-1] — без искажённого хвоста
    (зона развилки, «крючок» скруглённого торца). Окружность (R<25w, rms<0.02w,
    заметно изогнута) или прямая (rms<0.02w), направление — к концу.
    ('C', cx, cy, R, sign) | ('L', mx, my, dx, dy) | None."""
    P = []; s = 0.0
    i0 = 1 if nmax is None else max(1, len(q)-nmax+1)
    for i in range(len(q)-1, i0-1, -1):
        s += math.dist(q[i], q[i-1])
        if s > far*w:
            break
        if s >= near*w:
            P.append(q[i-1])            # P[0] — ближе всего к концу
    if len(P) < 4:
        return None
    c = _fit_circle(P)
    if c is not None and c[2] < 25*w and c[3] < 0.02*w and _bent(P, c[2], w):
        cx, cy, R, _ = c
        a0 = math.atan2(P[0][1]-cy, P[0][0]-cx); a1 = math.atan2(P[-1][1]-cy, P[-1][0]-cx)
        dlt = (a0 - a1 + math.pi) % (2*math.pi) - math.pi
        return ('C', cx, cy, R, 1.0 if dlt >= 0 else -1.0)
    mx, my, dx, dy, rms = _fit_line(P)
    if rms < 0.02*w:
        if (P[0][0]-P[-1][0])*dx + (P[0][1]-P[-1][1])*dy < 0:
            dx, dy = -dx, -dy
        return ('L', mx, my, dx, dy)
    return None


def _fit_param(fit, p):
    """Координата вдоль фита (растёт к концу)."""
    if fit[0] == 'L':
        return (p[0]-fit[1])*fit[3] + (p[1]-fit[2])*fit[4]
    _, cx, cy, R, sg = fit
    return sg * math.atan2(p[1]-cy, p[0]-cx) * R


def _apply_far(q, w, fit, near, maxdev, nmax=None):
    """Хвост (последние near*w) лечь на фит дальнего участка: все точки хвоста
    в пределах maxdev от фита -> спроецировать, выбросить «откаты» назад; иначе False."""
    s = 0.0; i0 = len(q)-1
    imin = 1 if nmax is None else max(1, len(q)-nmax+1)
    for i in range(len(q)-1, imin-1, -1):
        s += math.dist(q[i], q[i-1])
        if s > near*w:
            break
        i0 = i-1
    tail = q[i0:]
    devs = [math.dist(p, _fit_proj(fit, p)) for p in tail]
    # «крючок» на самом конце (остаток скруглённого торца после обрезки): последние точки
    # в пределах 0.5w, не лежащие на фите, просто отбрасываем
    k = len(tail); s = 0.0
    while k > 1 and devs[k-1] > maxdev:
        s += math.dist(tail[k-1], tail[k-2])
        if s > 0.5*w:
            break
        k -= 1
    if k < len(tail) and max(devs[:k]) <= maxdev:
        del q[i0+k:]; tail = tail[:k]; devs = devs[:k]
    if max(devs) > maxdev:
        return False
    # прямой хвост, а дальний фит от него уходит (тугая дуга перед прямым концом,
    # рукописная «n»): ближнее свидетельство важнее — хвост решает сам (_fit_end)
    if len(tail) >= 3 and max(devs) > 0.04*w:
        mx, my, dx, dy, rms = _fit_line(tail)
        if rms < 0.02*w:
            # направление прямого хвоста против касательной фита в его начале
            tx, ty = _fit_tangent(fit if fit[0] == 'L' else fit[:4], _fit_proj(fit, tail[0]),
                                  (tail[-1][0]-tail[0][0], tail[-1][1]-tail[0][1]))
            if abs(tx*dx + ty*dy) < math.cos(math.radians(6)):
                return False
    proj = [_fit_proj(fit, p) for p in tail]
    ref = _fit_param(fit, proj[0]); out = [proj[0]]; tmax = ref
    for p in proj[1:]:
        t = _fit_param(fit, p)
        if fit[0] == 'C':
            t = ref + ((t - ref + math.pi*fit[3]) % (2*math.pi*fit[3])) - math.pi*fit[3]
        if t > tmax + 1e-9:
            out.append(p); tmax = t
    q[i0:] = out
    return True


def _near_line(q, w, nmax=None, polys=(), bcorners=()):
    """Прямой участок от самого конца: точки добавляются, пока ложатся на прямую
    (макс. отклонение < 0.015w). Участок не короче 1w — ('L', ...) к концу; короче
    (0.4..1w) — только если прямая вдоль кромки контура у вогнутого угла рядом подтверждает
    направление (низ чаши P, перемычка B: короткая прямая перекладина перед дугой — её
    направление по своим точкам и кромке, а не по далёкой дуге). Иначе None."""
    imin = 0 if nmax is None else max(0, len(q)-nmax)
    P = [q[-1]]; best = None; s = 0.0; bestP = None
    for i in range(len(q)-2, imin-1, -1):
        P.append(q[i]); s += math.dist(q[i], q[i+1])
        if len(P) < 3:
            continue
        mx, my, dx, dy, rms = _fit_line(P)
        dev = max(abs((p[0]-mx)*dy - (p[1]-my)*dx) for p in P)
        if dev > 0.015*w:
            break
        if s >= 0.4*w:
            if (q[-1][0]-q[i][0])*dx + (q[-1][1]-q[i][1])*dy < 0:
                dx, dy = -dx, -dy
            best = ('L', mx, my, dx, dy); bestL = s; bestP = list(P)
    if best is None:
        return None
    if bestL >= 1.0*w:
        # пологая дуга (терминалы C, S) тоже укладывается в 0.015w на 1w: по точкам
        # участка без последних 0.3w (там остаток «крючка» конца) окружность R < 8w
        # ложится вдвое лучше прямой — это дуга, не прямая
        Pin = []; sk = 0.0
        for k in range(1, len(bestP)):
            sk += math.dist(bestP[k], bestP[k-1])
            if sk >= 0.3*w:
                Pin.append(bestP[k])
        arc = False
        if len(Pin) >= 4:
            lr = _fit_line(Pin)[4]; c = _fit_circle(Pin)
            arc = c is not None and c[2] < 8*w and c[3] < 0.002*w and c[3] < 0.5*lr
        if not arc:
            return _refine_line(best, bestP, w, polys), bestL
    # короткий участок: нужна кромка контура с тем же направлением (в 12 град), на которой
    # лежат точки участка (средняя дистанция < 0.25w)
    for c in bcorners:
        if c[2] or math.dist(c[:2], q[-1]) > 2.0*w:
            continue
        for ei, e in _edge_lines(c, w, polys, idx=True):
            if c[5 + ei] > 0.02:
                continue   # кромка сама не прямая на 0.8w (арка m, n) — не подтверждение
            if abs(e[3]*best[3] + e[4]*best[4]) >= math.cos(math.radians(12)):
                d = sum(math.dist(p, _fit_proj_c(e, p)) for p in bestP) / len(bestP)
                if d < 0.25*w:
                    dx, dy = e[3], e[4]
                    if dx*best[3] + dy*best[4] < 0:
                        dx, dy = -dx, -dy
                    return ('L', e[1], e[2], dx, dy), bestL
    return None


def _blend_herm(q, kf, E, dx, dy, L, h):
    """Запасное сопряжение: дуга шрифта не окружность — от q[kf] кривая Эрмита к началу прямого
    участка A = E - d*L по касательной (d), дальше прямая до конца E."""
    A = (E[0] - dx*L, E[1] - dy*L)
    if kf < 1 or (A[0]-q[kf][0])*dx + (A[1]-q[kf][1])*dy <= 0:
        return False
    tk = _norm((q[kf][0]-q[kf-1][0], q[kf][1]-q[kf-1][1]))
    if tk[0]*dx + tk[1]*dy < 0.5:
        return False
    # одна дуга, касательная к хвосту в q[kf] и к прямой (точка касания T не дальше конца E)
    K = q[kf]
    def sd(P):
        return (P[0]-E[0])*dy - (P[1]-E[1])*dx
    for sg in (1.0, -1.0):
        nk = (-tk[1]*sg, tk[0]*sg)
        sn = nk[0]*dy - nk[1]*dx
        s0 = sd(K)
        for side in (1.0, -1.0):
            den = side - sn
            if abs(den) < 1e-9:
                continue
            r = s0 / den
            if r <= 0:
                continue
            c = (K[0] + nk[0]*r, K[1] + nk[1]*r)
            tc = (c[0]-E[0])*dx + (c[1]-E[1])*dy
            T = (E[0] + dx*tc, E[1] + dy*tc)
            if abs(math.dist(c, T) - r) > 1e-6*max(1.0, r) or -tc < -0.05*L or -tc > 3*L:
                continue
            a0 = math.atan2(K[1]-c[1], K[0]-c[0]); a1 = math.atan2(T[1]-c[1], T[0]-c[0])
            da = (a1 - a0) % (2*math.pi)
            if sg < 0:
                da -= 2*math.pi          # центр справа — по часовой
            # касательная дуги в T — по d
            at = a0 + da
            tt = _norm((-math.sin(at)*(1 if da > 0 else -1), math.cos(at)*(1 if da > 0 else -1)))
            if tt[0]*dx + tt[1]*dy < 0.98 or abs(da) > math.pi:
                continue
            na = max(2, int(abs(da)*r/h))
            new = [(c[0] + r*math.cos(a0 + da*k/na), c[1] + r*math.sin(a0 + da*k/na)) for k in range(1, na+1)]
            Ls = max(0.0, -tc); nl = max(1, int(Ls/h))
            if Ls > 1e-6:
                new += [(T[0] + dx*Ls*k/nl, T[1] + dy*Ls*k/nl) for k in range(1, nl+1)]
            q[kf+1:] = new
            return True
    new = _herm(q[kf], tk, A, (dx, dy), h) + [A]
    nl = max(1, int(L/h))
    new += [(A[0] + dx*L*k/nl, A[1] + dy*L*k/nl) for k in range(1, nl+1)]
    q[kf+1:] = new
    return True


def _blend_arc(q, fit, L, w, nmax=None):
    """Сопряжение прямой (последние L точек q лежат на fit) с дугой, идущей дальше по цепочке.
    Окружность по точкам на 1.4w от первой точки, явно сошедшей с прямой (> 0.08w).
    Касается прямой (|dist - R| < 0.015w) — участок перестраивается как дуга до точки
    касания + прямая. Пересекает прямую под малым углом (< 20 град; ось Делоне у
    квадратного угла счётчика: низ чаши P, перекладина B) — прямая до конца прямого
    участка, галтель (дуга r, касательная к обеим) и дальше большая дуга. Иначе — ничего."""
    n = len(q)
    imin = 0 if nmax is None else max(0, n-nmax)
    S = [0.0]*n; s = 0.0
    for i in range(n-2, imin-1, -1):
        s += math.dist(q[i], q[i+1]); S[i] = s
    _, mx, my, dx, dy = fit
    s1 = None
    for i in range(n-1, imin-1, -1):
        if S[i] > L and abs((q[i][0]-mx)*dy - (q[i][1]-my)*dx) > 0.08*w:
            s1 = S[i]
            break
    if s1 is None or S[imin] < s1 + 1.4*w:
        return False
    P = [q[i] for i in range(imin, n) if s1 <= S[i] <= s1 + 1.4*w]
    if len(P) < 4:
        return False
    c = _fit_circle(P)
    if c is None or c[2] > 25*w or c[3] > 0.03*w:
        return False
    cx, cy, R = c[0], c[1], c[2]
    kf = max(i for i in range(imin, n) if S[i] >= s1 + 1.4*w)
    h = S[imin] / max(1, n-1-imin)
    E = q[-1]
    dist = abs((cx-mx)*dy - (cy-my)*dx)
    if R < dist <= R + 0.15*w:
        # дуга шрифта не точно окружность: подогнанная чуть не доходит до прямой — сдвинуть
        # центр к прямой до касания (на величину зазора, < 0.15w)
        g = dist - R
        sgn = 1.0 if ((cx-mx)*dy - (cy-my)*dx) > 0 else -1.0
        cx -= sgn*dy*g; cy += sgn*dx*g
        dist = R
    arcs = []   # список (центр, R, от, до) — дуги от q[kf] к прямой
    if abs(dist - R) <= 0.015*w:
        T = _fit_proj(fit, (cx, cy))
        tT = (E[0]-T[0])*dx + (E[1]-T[1])*dy       # вдоль прямой к концу
        tF = (E[0]-q[kf][0])*dx + (E[1]-q[kf][1])*dy
        if not (tF > tT > 0.1*w):
            return _blend_herm(q, kf, E, dx, dy, L, h)
        arcs.append(((cx, cy), R, q[kf], T)); A = T
    elif dist < R:
        if math.degrees(math.acos(dist/R)) > 20:
            return False
        # пересечение прямой с окружностью: X(t) = E - d*t, ближайшее к концу (t > 0)
        vx, vy = E[0]-cx, E[1]-cy
        b = vx*dx + vy*dy; disc = b*b - (vx*vx + vy*vy - R*R)
        if disc < 0:
            return False
        roots = [t for t in (b - math.sqrt(disc), b + math.sqrt(disc)) if t > 0.1*w]
        if not roots:
            return False
        tx = min(roots)
        tA = min(L, tx - 0.1*w)
        if tA < 0.1*w:
            return False
        A = (E[0] - dx*tA, E[1] - dy*tA)
        n0 = (-dy, dx)
        if (cx-A[0])*n0[0] + (cy-A[1])*n0[1] < 0:
            n0 = (dy, -dx)
        ax, ay = A[0]-cx, A[1]-cy
        den = 2*(R + ax*n0[0] + ay*n0[1])
        if den <= 1e-9:
            return False
        r = (R*R - ax*ax - ay*ay) / den
        if not (0.1*w < r < R):
            return False
        cf = (A[0] + n0[0]*r, A[1] + n0[1]*r)
        Tb = (cx + (cf[0]-cx)*R/(R-r), cy + (cf[1]-cy)*R/(R-r))
        arcs.append(((cx, cy), R, q[kf], Tb))
        arcs.append((cf, r, Tb, A))
    else:
        return False
    new = []
    for (ccx, ccy), rr, p0, p1 in arcs:
        a0 = math.atan2(p0[1]-ccy, p0[0]-ccx); a1 = math.atan2(p1[1]-ccy, p1[0]-ccx)
        da = (a1 - a0 + math.pi) % (2*math.pi) - math.pi
        na = max(1, int(abs(da)*rr/h))
        for k in range(0 if not new else 1, na+1):
            a = a0 + da*k/na
            new.append((ccx + rr*math.cos(a), ccy + rr*math.sin(a)))
    Ls = math.dist(A, E); nl = max(1, int(Ls/h))
    for k in range(1, nl+1):
        new.append((A[0] + dx*Ls*k/nl, A[1] + dy*Ls*k/nl))
    q[kf:] = new
    return True


def _refine_line(fit, P, w, polys):
    """Уточнить прямую по контуру: середины поперечников (лучи поперёк до краёв) в точках
    проекций P; берутся только «чистые» поперечники — ширина как у штриха (+-10% от
    медианы) и центр рядом с прямой (< 0.15w). Ось по точкам у развилки/торца чуть
    наклонена (ножка ц, p), контур — нет. Две итерации; мало чистых точек — fit как есть."""
    _, mx, my, dx, dy = fit
    for _it in range(2):
        nrm = (-dy, dx); WR = []
        for p in P:
            t = (p[0]-mx)*dx + (p[1]-my)*dy
            X = (mx + dx*t, my + dy*t)
            dl = _raycast_union(polys, X, nrm); dr = _raycast_union(polys, X, (-nrm[0], -nrm[1]))
            if dl is not None and dr is not None:
                WR.append((X, dl, dr))
        if len(WR) < 3:
            return fit
        ws = sorted(a + b for _, a, b in WR); wmed = ws[len(ws)//2]
        M = [(X[0] + nrm[0]*(a-b)/2, X[1] + nrm[1]*(a-b)/2) for X, a, b in WR
             if abs(a + b - wmed) < 0.1*wmed and abs(a - b) < 0.3*w]
        if len(M) < max(3, len(P)//3):
            return fit
        nx, ny, ndx, ndy, _r = _fit_line(M)
        if ndx*dx + ndy*dy < 0:
            ndx, ndy = -ndx, -ndy
        mx, my, dx, dy = nx, ny, ndx, ndy
    return ('L', mx, my, dx, dy)


def _stub_extend(q, w, poly, jw_j, xy_j, h, nmax=None, polys=(), v0=0, v1=0, bcorners=()):
    """Заглушка: продолжить ось до сквозной линии развилки. Прямой штрих — прямой
    (по дальнему участку, хвост у развилки спроецирован), дуга — по кривизне.
    Если кривая проходит по касательной (не пересекает) — к ближайшей точке сквозной."""
    if poly is None:
        return False
    nl = _near_line(q, w, nmax, polys, bcorners)
    fit = None
    if nl is not None:
        fit, L = nl
        _end_project(q, fit, w, nmax, L)   # только прямой участок, дальше дуга остаётся дугой
        _blend_arc(q, fit, L, w, nmax)
    else:
        fit = _far_fit(q, w, 1.5, 3.5, nmax)
        if fit is None or not _apply_far(q, w, fit, 1.5, 0.12*w, nmax):
            fit = _fit_end(q, w, nmax)
            if fit is None:
                return False
            _end_project(q, fit, w, nmax)
    start = q[-1]; step = h/2
    maxhit = math.dist(start, xy_j)*1.5 + jw_j       # пересечение — только в этих пределах
    # касание прямой ножки вдоль дуги (g: ножка в кольцо) — можно чуть дальше
    maxlen = maxhit + (jw_j if fit[0] == 'L' else 0.0)
    prev = start; added = []; k = 1
    bestd, besti, bestc = float('inf'), -1, None
    while k*step <= maxlen:
        p = _fit_at(fit, start, k*step)
        hit = _seg_hit(poly, prev, p) if k*step <= maxhit else None
        if hit is None and not _inside_any(polys, p[0], p[1]):
            break                       # продолжение вышло за контур — дальше не ищем
        if hit is not None and _inside_any(polys, hit[0], hit[1]):
            q.extend(added); q.append(hit)
            return True
        d, c = _poly_closest(poly[v0:len(poly)-v1], p)   # без виртуальных концов
        if bestc is not None and bestd < 0.15*jw_j and d > bestd + 0.02*jw_j:
            # касание: расстояние до сквозной прошло через минимум и снова растёт
            # (дальше кривая идёт вдоль сквозной параллельно — петля «6», «9»)
            q.extend(added[:besti+1]); q.append(bestc)
            _land_offset(q, w)
            return True
        if d < bestd:
            bestd, besti, bestc = d, len(added), c
        lastd = d
        added.append(p); prev = p; k += 1
    def parallel_at(i, c):
        # направление кривой в точке added[i] и сквозной в точке c — почти параллельны (< 15 град)
        if not (0 < i < len(added) - 1):
            return False
        tc = _norm((added[i+1][0]-added[i-1][0], added[i+1][1]-added[i-1][1]))
        real_ = poly[v0:len(poly)-v1]
        it = min(range(len(real_)), key=lambda t: math.dist(real_[t], c))
        i1 = min(it + 1, len(real_) - 1); i0 = max(it - 1, 0)
        td = _norm((real_[i1][0]-real_[i0][0], real_[i1][1]-real_[i0][1]))
        return abs(tc[0]*td[0] + tc[1]*td[1]) >= math.cos(math.radians(15))
    if bestc is not None and besti < len(added) - 1 and lastd > bestd + 0.01*w and \
            (bestd < 0.15*jw_j or (bestd < 0.3*jw_j and parallel_at(besti, bestc))):
        # касание большой дугой (арка n, m приходит в ножку по касательной): минимум расстояния
        # внутри участка, дальше оно растёт. Дуга шрифта не точно окружность — зазор до 0.3jw
        # допустим, если в точке подхода кривая почти параллельна сквозной (зазор разводится плавно)
        q.extend(added[:besti+1]); q.append(bestc)
        _land_offset(q, w)
        return True
    # ни пересечения, ни касания. Сквозная могла поворачивать в развилке (база Д уходит в
    # диагональ, а её кусок за развилкой отрезан как усик): продолжения сторон сквозной за
    # точку у развилки (до 1jw) — луч заглушки попадает в продолжение: конец в точке
    # попадания и уголком к сквозной (нога Д стоит на базе)
    real = poly[v0:len(poly)-v1]
    # направление — прямая по участку 0.5..3w от конца (загнутый к развилке конец не в счёт)
    P = []; s_ = 0.0
    for i in range(len(q)-2, -1, -1):
        s_ += math.dist(q[i], q[i+1])
        if s_ > 3.0*w:
            break
        if s_ >= 0.5*w:
            P.append(q[i])
    if len(real) >= 3 and len(P) >= 3 and _fit_line(P)[4] < 0.02*w:
        mx_, my_, dx_, dy_, _r = _fit_line(P)
        if (q[-1][0]-P[0][0])*dx_ + (q[-1][1]-P[0][1])*dy_ < 0:
            dx_, dy_ = -dx_, -dy_
        fl_ = ('L', mx_, my_, dx_, dy_)
        # обрезать конец, сошедший с прямой (> 0.1w)
        qq = list(q)
        while len(qq) > 2 and math.dist(qq[-1], _fit_proj(fl_, qq[-1])) > 0.1*w:
            qq.pop()
        start = _fit_proj(fl_, qq[-1]); fit = fl_
        # продолжения сторон за вершины сквозной у развилки (< 1jw от центра); ближайшее попадание
        best = None
        for iv in range(1, len(real)-1):
            V = real[iv]
            if math.dist(V, xy_j) > jw_j:
                continue
            for nb in (real[iv-1], real[iv+1]):
                if math.dist(V, nb) < 1e-6:
                    continue
                d = _norm((V[0]-nb[0], V[1]-nb[1]))
                ext = [V, (V[0]+d[0]*jw_j, V[1]+d[1]*jw_j)]
                prev = start; k = 1
                while k*step <= maxhit:
                    p = _fit_at(fit, start, k*step)
                    hit = _seg_hit(ext, prev, p)
                    if hit is not None:
                        if math.dist(hit, V) > 0.02*w and _inside_any(polys, hit[0], hit[1]) \
                                and _poly_inside(polys, [start, hit, V], 0.5*h):
                            L_ = math.dist(start, hit)
                            if best is None or L_ < best[0]:
                                best = (L_, hit, V)
                        break
                    prev = p; k += 1
        if best is not None:
            L_, hit, V = best
            q[:] = qq[:-1] + [start]
            m = max(1, int(L_/step))
            q.extend(_fit_at(fit, start, L_*i/m) for i in range(1, m))
            q.append(hit); q.append(V)
            return True
    # запасной вариант (плавно к точке у центра развилки)
    return False


def _seg_dist(a, b, p):
    dx = b[0]-a[0]; dy = b[1]-a[1]; L2 = dx*dx + dy*dy
    t = 0.0 if L2 < 1e-18 else max(0.0, min(1.0, ((p[0]-a[0])*dx + (p[1]-a[1])*dy)/L2))
    return math.hypot(a[0] + dx*t - p[0], a[1] + dy*t - p[1])


def _rsf(k):
    if _DBG is not None:
        _DBG.append(('rsf', k))


def _ref_stub(q, w, poly, jw_j, xy_j, h, polys, v0=0, v1=0):
    """Режим образца: заглушку продолжить по её собственному ходу (прямая или дуга по
    последним ~1.5w пути) до пересечения с линией poly. Если продолжение не доходит
    (штрих гнётся к стойке сильнее), — отрезком прямо от конца к точке poly с наименьшим
    отклонением от касательной (< 35 град). Попадание принимается, если точка внутри контура,
    штрих подходит под углом (> 12 град к poly) и за poly выходит из контура быстро (< 1.2 jw):
    при касательном слиянии продолжение долго идёт внутри соседнего штриха."""
    if poly is None or len(q) < 3:
        _rsf(1); return False
    P = [q[-1]]; s = 0.0
    for i in range(len(q)-1, 0, -1):
        s += math.dist(q[i], q[i-1])
        P.append(q[i-1])
        if s >= 1.5*w:
            break
    if len(P) < 3 or s < 0.4*w:
        _rsf(2); return False
    start = q[-1]; step = h/2
    maxhit = math.dist(start, xy_j) + 1.5*jw_j
    fl = _fit_line(P)
    d = (fl[2], fl[3])
    if (P[0][0]-P[-1][0])*d[0] + (P[0][1]-P[-1][1])*d[1] < 0:
        d = (-d[0], -d[1])
    at = None
    if fl[4] <= 0.03*w:
        def at(t, d=d):
            return (start[0] + d[0]*t, start[1] + d[1]*t)
    elif len(P) >= 4:
        c = _fit_circle(P)
        if c is not None and c[2] > 0.8*w and c[3] <= 0.03*w:
            cx, cy, R = c[:3]
            a0 = math.atan2(start[1]-cy, start[0]-cx)
            ap = math.atan2(P[1][1]-cy, P[1][0]-cx)
            sgn = 1.0 if math.sin(a0-ap) > 0 else -1.0
            # касательная в начале — для проверки угла подхода
            d = _norm((-sgn*math.sin(a0), sgn*math.cos(a0)))
            def at(t, a0=a0, sgn=sgn, cx=cx, cy=cy, R=R):
                a = a0 + sgn*t/R
                return (cx + R*math.cos(a), cy + R*math.sin(a))

    def accept(hit, dv):
        if not _inside_any(polys, hit[0], hit[1]):
            return False
        it = min(range(len(poly)-1), key=lambda t: _seg_dist(poly[t], poly[t+1], hit))
        td = _norm((poly[it+1][0]-poly[it][0], poly[it+1][1]-poly[it][1]))
        if abs(td[0]*dv[0] + td[1]*dv[1]) > math.cos(math.radians(12)):
            return False
        # касательное примыкание (угол к линии < 30 град: чаша «а» в стойку) — садимся у самой
        # развилки, а не там, где направления совпадут лучше всего (иначе «горб» над стыком)
        if abs(td[0]*dv[0] + td[1]*dv[1]) > math.cos(math.radians(30)) and math.dist(hit, xy_j) > 0.8*jw_j:
            return False
        ex = 0.0
        while ex <= 1.2*jw_j:
            ex += step
            pe = (hit[0] + dv[0]*ex, hit[1] + dv[1]*ex)
            if not _inside_any(polys, pe[0], pe[1]):
                break
        return ex <= 1.2*jw_j
    if at is not None:
        prev = start; added = []; k = 1
        while k*step <= maxhit:
            p = at(k*step)
            hit = _seg_hit(poly, prev, p)
            if hit is not None:
                dv = _norm((p[0]-prev[0], p[1]-prev[1]))
                if accept(hit, dv):
                    q.extend(added); q.append(hit)
                    return True
                _rsf(4)
                break
            if not _inside_any(polys, p[0], p[1]):
                break
            added.append(p); prev = p; k += 1
    best = None
    for t in range(len(poly)):
        pt = poly[t]
        L = math.dist(start, pt)
        if L < 0.3*w or L > maxhit:
            continue
        dv = ((pt[0]-start[0])/L, (pt[1]-start[1])/L)
        dev = dv[0]*d[0] + dv[1]*d[1]
        if dev < math.cos(math.radians(35)):
            continue
        if best is not None and dev < best[0] - 1e-9:
            continue
        if not _poly_inside(polys, [start, pt], 0.5*h):
            continue
        if not accept(pt, dv):
            continue
        if best is None or dev > best[0] + 1e-9 or (abs(dev - best[0]) <= 1e-9 and L < best[1]):
            best = (dev, L, pt)
    if best is None:
        _rsf(6); return False
    pt = best[2]; L = best[1]
    m = max(1, int(L/step))
    q.extend((start[0] + (pt[0]-start[0])*i/m, start[1] + (pt[1]-start[1])*i/m) for i in range(1, m))
    q.append(pt)
    return True

def _merge_collinear(pts, fixed, closed, w):
    """Прямые куски оси (отклонение < 0.02w, длина >= 1w), соседние, почти соосные (< 3 град)
    и смещённые друг относительно друга меньше чем на 0.15w (половины диагонали «х» в шрифте
    чуть сдвинуты у перекрестия), — на одну общую прямую: без ступеньки в стыке."""
    n = len(pts)
    if n < 6 or closed:
        return pts
    fx = set(fixed)
    pieces = []; i = 0
    while i < n - 2:
        j = i + 2; best = None
        while j < n:
            if (j - 1) in fx and j - 1 > i:
                break
            if _line_dev(pts[i:j+1]) > 0.02*w:
                break
            best = j; j += 1
        if best is not None and math.dist(pts[i], pts[best]) >= 1.0*w:
            pieces.append([i, best, False]); i = best + 1
        else:
            i += 1
    changed = True
    while changed:
        changed = False
        for k in range(len(pieces) - 1):
            a0, a1 = pieces[k][:2]; b0, b1 = pieces[k+1][:2]
            if b0 - a1 > 3 or any(q in fx for q in range(a1, b0 + 1)):
                continue
            fa = _fit_line(pts[a0:a1+1]); fb = _fit_line(pts[b0:b1+1])
            if abs(fa[2]*fb[2] + fa[3]*fb[3]) < math.cos(math.radians(3)):
                continue
            mb = pts[(b0 + b1)//2]
            off = abs((mb[0]-fa[0])*fa[3] - (mb[1]-fa[1])*fa[2])
            if off > 0.15*w:
                continue
            pieces[k] = [a0, b1, True]; del pieces[k+1]; changed = True
            break
    for a0, a1, merged in pieces:
        if not merged:
            continue
        mx, my, dx, dy, _r = _fit_line(pts[a0:a1+1])
        def near(q):
            return abs((pts[q][0]-mx)*dy - (pts[q][1]-my)*dx) < 0.2*w
        while a0 > 0 and (a0 - 1) not in fx and near(a0 - 1):
            a0 -= 1          # прилегающие точки до конца линии — тоже на общую прямую
        while a1 < n - 1 and (a1 + 1) not in fx and near(a1 + 1):
            a1 += 1
        # только если общая прямая доходит до конца линии или до острого угла с обеих сторон
        # (диагональ «х» целиком), иначе сдвиг даёт ступеньку посреди штриха
        if not ((a0 == 0 or (a0 - 1) in fx or a0 in fx) and (a1 == n - 1 or (a1 + 1) in fx or a1 in fx)):
            continue
        for q in range(a0, a1 + 1):
            t = (pts[q][0]-mx)*dx + (pts[q][1]-my)*dy
            pts[q] = (mx + dx*t, my + dy*t)
    return pts


def _land_offset(q, w):
    """Касание сквозной: последняя точка q — точка сквозной, предпоследняя — на кривой
    заглушки на боковом зазоре off от неё. Зазор распределяется по последним ~1.5w точек
    (гладкий спад), чтобы вместо бокового «шажка» был плавный подход."""
    if len(q) < 3:
        return
    off = (q[-1][0]-q[-2][0], q[-1][1]-q[-2][1])
    m = math.hypot(*off)
    if m < 0.01*w:
        return
    L = max(1.5*w, 6.0*m); s = 0.0
    for i in range(len(q)-2, 0, -1):
        if i < len(q)-2:
            s += math.dist(q[i], q[i+1])     # путь от предпоследней (сам зазор — не путь)
        if s >= L:
            break
        u = s/L; wt = 2*u**3 - 3*u*u + 1
        q[i] = (q[i][0] + off[0]*wt, q[i][1] + off[1]*wt)


def _stub_to_junction(q, w, poly, xy_j, h, polys, v0=0, v1=0, jw_j=None, rawc=None):
    """Запасной вариант заглушки: плавная кривая (Эрмит) от конца к точке сквозной линии,
    ближайшей к центру развилки; касательная в начале — по оси заглушки. Если кривая
    вылезает за контур — прямой отрезок."""
    real = poly[v0:len(poly)-v1] if poly is not None else []
    if jw_j is None:
        jw_j = w
    a = q[-1]
    ta = _tan(q, w)
    if len(real) >= 2:
        _, tgt = _poly_closest(real, xy_j)
        # вливаемся по касательной к сквозной: цель сдвигаем вдоль неё на 0.5jw в ту
        # сторону, куда и так идём, касательная в конце — направление сквозной
        it = min(range(len(real)), key=lambda t: math.dist(real[t], tgt))
        i1 = min(it + 1, len(real) - 1); i0 = max(it - 1, 0)
        td = _norm((real[i1][0]-real[i0][0], real[i1][1]-real[i0][1]))
        if td[0]*(tgt[0]-a[0]) + td[1]*(tgt[1]-a[1]) < 0:
            td = (-td[0], -td[1]); step = -1
        else:
            step = 1
        s = 0.0; k = it
        while 0 <= k + step < len(real) and s < 0.5*jw_j:
            s += math.dist(real[k], real[k+step]); k += step
        tgt = real[k]
        td = _norm((real[min(k+1, len(real)-1)][0]-real[max(k-1, 0)][0], real[min(k+1, len(real)-1)][1]-real[max(k-1, 0)][1]))
        if td[0]*(tgt[0]-a[0]) + td[1]*(tgt[1]-a[1]) < 0:
            td = (-td[0], -td[1])
    else:
        tgt = xy_j
        td = _norm((tgt[0]-a[0], tgt[1]-a[1]))
    ch = _norm((tgt[0]-a[0], tgt[1]-a[1]))
    if ta[0]*ch[0] + ta[1]*ch[1] < 0.3:
        ta = ch
    fl = C._herm(a, ta, tgt, td, h) + [tgt]
    if _poly_inside(polys, [a] + fl, 0.5*h):
        q.extend(fl)
    else:
        fl = C._herm(a, ta, tgt, ch, h) + [tgt]
        if _poly_inside(polys, [a] + fl, 0.5*h):
            q.extend(fl)
        elif _poly_inside(polys, [a, tgt], 0.5*h):
            q.append(tgt)
        else:
            # прямая к сквозной режет контур (вырез между штрихами): через центр развилки
            cj = _norm((xy_j[0]-a[0], xy_j[1]-a[1]))
            if ta[0]*cj[0] + ta[1]*cj[1] < 0.3:
                ta = cj
            f1 = C._herm(a, ta, xy_j, _norm((tgt[0]-xy_j[0], tgt[1]-xy_j[1])), h) + [xy_j]
            if _poly_inside(polys, [a] + f1 + [tgt], 0.5*h):
                q.extend(f1); q.append(tgt)
            elif _poly_inside(polys, [a, xy_j, tgt], 0.5*h):
                q.append(xy_j); q.append(tgt)
            elif rawc:
                # по сырой оси (узлы графа всегда внутри) от ближайшего к концу узла до развилки,
                # сглаженной; затем к сквозной, если отрезок внутри
                ia = min(range(len(rawc)), key=lambda t: math.dist(rawc[t], a))
                tail = list(rawc[ia+1:])
                if tail and math.dist(tail[-1], xy_j) > 1e-6:
                    tail.append(xy_j)
                for _ in range(3):
                    for t in range(len(tail)-1):
                        pr = a if t == 0 else tail[t-1]
                        cand = ((pr[0]+tail[t+1][0])/2*0.5 + tail[t][0]*0.5, (pr[1]+tail[t+1][1])/2*0.5 + tail[t][1]*0.5)
                        if _inside_any(polys, cand[0], cand[1]):
                            tail[t] = cand
                if _poly_inside(polys, [a] + tail, 0.5*h):
                    q.extend(tail)
                    if math.dist(q[-1], tgt) > 1e-6 and _poly_inside(polys, [q[-1], tgt], 0.5*h):
                        q.append(tgt)
            # иначе конец заглушки остаётся как есть — лучше разрыв, чем линия вне контура



# ---------------------------------------------------------------- восстановленные/обновлённые блоки

def _fit_circle(P):
    """Касательная (Kasa) аппроксимация окружности; (cx, cy, R, rms) или None."""
    n = len(P)
    if n < 4:
        return None
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


def _fit_line(P):
    n = len(P)
    mx = sum(p[0] for p in P)/n; my = sum(p[1] for p in P)/n
    sxx = sum((p[0]-mx)**2 for p in P); syy = sum((p[1]-my)**2 for p in P)
    sxy = sum((p[0]-mx)*(p[1]-my) for p in P)
    ang = 0.5*math.atan2(2*sxy, sxx-syy)
    dx, dy = math.cos(ang), math.sin(ang)
    rms = math.sqrt(sum(((p[0]-mx)*(-dy)+(p[1]-my)*dx)**2 for p in P)/n)
    return mx, my, dx, dy, rms


def _poly_dist(raw, poly):
    """Средняя дистанция точек raw до ломаной poly."""
    if not raw or len(poly) < 2:
        return float('inf')
    tot = 0.0
    for p in raw:
        best = float('inf')
        for a, b in zip(poly[:-1], poly[1:]):
            vx, vy = b[0]-a[0], b[1]-a[1]; L2 = vx*vx + vy*vy
            t = 0.0 if L2 < 1e-12 else max(0.0, min(1.0, ((p[0]-a[0])*vx + (p[1]-a[1])*vy)/L2))
            best = min(best, math.dist(p, (a[0]+vx*t, a[1]+vy*t)))
        tot += best
    return tot / len(raw)


def _arc_fill(before, after, step, w, polys, bcorners=(), raw=()):
    """Дыру в зоне развилки на сквозной линии заполняем по окружности (или прямой),
    подогнанной к точкам с обеих сторон — линия не "проседает" к развилке.
    Окно подгонки уменьшается (2.5w -> 1.5w -> 0.8w), пока точки не лягут на фит."""
    def take(seq, L):
        out = [seq[0]]; s = 0.0
        for a, b in zip(seq[:-1], seq[1:]):
            s += math.dist(a, b)
            if s > L:
                break
            out.append(b)
        return out
    proj = None
    r = _short_side_line(before, after, step, w, polys)
    if r is not None:
        return r
    # стороны по отдельности: если ось подходит к развилке под заметным углом
    # (Y, V, k) — возможно, это угол, а не плавная дуга через две короткие стороны;
    # кто прав — решают сырые узлы зоны (raw): они идут через центр развилки (угол)
    # или вдоль плавной кривой (изгиб штриха в пересечении рукописной «е»)
    corner = None
    A = take(before[::-1], 2.0*w); B = take(after, 2.0*w)
    fa = _fit_pts(A, w); fb = _fit_pts(B, w)
    if fa is not None and fb is not None and len(A) > 1 and len(B) > 1:
        bc = list(before)
        r = _corner_fill(bc, after, step, w, polys, bcorners)
        if r[0] is not None:
            if r[2] is None:
                # составная кривая (прямая + дуга по касательной) — точнее одного фита
                before[:] = bc
                return r
            corner = (bc, r)
    smooth = _arc_fill_single(before, after, step, w)
    if corner is None and smooth is None:
        br = _tangent_bridge(before, after, step, w, polys)
        if br is not None:
            before[:] = br[0]
            return br[1]
    if corner is not None and smooth is not None:
        bs, rs = smooth
        dc = _poly_dist(raw, [corner[0][-1]] + corner[1][0] + [after[0]])
        ds = _poly_dist(raw, [bs[-1]] + rs[0] + [after[0]])
        if dc <= ds:
            smooth = None
        else:
            corner = None
    if corner is not None:
        before[:] = corner[0]
        return corner[1]
    if smooth is not None:
        before[:] = smooth[0]
        return smooth[1]
    # запасной вариант: сквозная прямая одной стороны проходит через начало другой (к: перекладина
    # идёт прямо через развилку, а за развилкой сразу фаска — фит «через обе стороны» не ложится)
    cb = _collinear_bridge(before, after, step, w, polys)
    if cb is not None:
        return cb, None, None
    return None, None, None


def _collinear_bridge(before, after, step, w, polys):
    """Прямая по стороне (2w, rms < 0.02w) проходит через первую точку другой стороны (< 0.05w), и
    другая сторона начинается вдоль неё (первые 0.4w — в пределах 0.05w): заполняем отрезком."""
    def take(seq, L):
        out = [seq[0]]; s_ = 0.0
        for a_, b_ in zip(seq[:-1], seq[1:]):
            s_ += math.dist(a_, b_)
            if s_ > L:
                break
            out.append(b_)
        return out
    if len(before) < 3 or len(after) < 3:
        return None
    for A, B in ((take(before[::-1], 2.0*w), take(after, 0.4*w)), (take(after, 2.0*w), take(before[::-1], 0.4*w))):
        if len(A) < 4 or len(B) < 2:
            continue
        mx, my, dx, dy, rr = _fit_line(A)
        if rr > 0.02*w:
            continue
        if max(abs((q[0]-mx)*dy - (q[1]-my)*dx) for q in B) > 0.05*w:
            continue
        P0, P1 = before[-1], after[0]
        fl = []
        n = max(1, int(math.dist(P0, P1) / step))
        for i in range(1, n):
            t = i / n
            fl.append((P0[0] + (P1[0]-P0[0])*t, P0[1] + (P1[1]-P0[1])*t))
        if polys and not _poly_inside(polys, [P0] + fl + [P1], 0.5*step):
            continue
        return fl
    return None


def _short_side_line(before, after, step, w, polys=()):
    """Одна сторона дыры короткая (< 1.5w до конца цепочки: колпачок ножки над стыком с чашей
    p, b), другая — прямая на 2w, и точки короткой лежат на её продолжении (< 0.2w): обход
    и короткая сторона — по прямой длинной стороны (без излома ножки в стыке)."""
    def plen(seq):
        return sum(math.dist(a, b) for a, b in zip(seq[:-1], seq[1:]))
    def take(seq, L):
        out = [seq[0]]; s = 0.0
        for a, b in zip(seq[:-1], seq[1:]):
            s += math.dist(a, b)
            if s > L:
                break
            out.append(b)
        return out
    lb, la = plen(before), plen(after)
    if min(lb, la) >= 1.5*w or max(lb, la) < 2.8*w:
        return None
    longB = lb >= la
    seq = before[::-1] if longB else after
    # прямая по дальнему участку длинной стороны (0.8..2.8w, иначе 0.8..1.8w): у самого стыка
    # ось проседает, а дальше может начаться зона следующего стыка
    sh0 = after if longB else before
    Lg = None
    for far in (2.8, 1.8, 1.4):
        P = []; s = 0.0
        for a_, b_ in zip(seq[:-1], seq[1:]):
            s += math.dist(a_, b_)
            if s > far*w:
                break
            if s >= 0.8*w:
                P.append(b_)
        if len(P) >= 3:
            # прямая должна лечь и на дальний участок длинной стороны, и на короткую сторону
            mx, my, dx, dy, rms = _fit_line(P + list(sh0))
            if rms <= 0.02*w:
                Lg = P
                break
    if Lg is None:
        return None
    sh = after if longB else before
    if any(abs((p[0]-mx)*dy - (p[1]-my)*dx) > 0.2*w for p in sh):
        return None
    def proj(q, mx=mx, my=my, dx=dx, dy=dy):
        t = (q[0]-mx)*dx + (q[1]-my)*dy
        return (mx + dx*t, my + dy*t)
    if longB:
        s = 0.0; i = len(before)-1
        while i >= 0:
            before[i] = proj(before[i])
            if i > 0:
                s += math.dist(before[i], before[i-1])
            if s > 0.8*w:
                break
            i -= 1
    else:
        for i in range(len(before)):
            before[i] = proj(before[i])
    a = before[-1]; b = proj(after[0])
    m = max(1, int(math.dist(a, b)/step))
    fl = [proj((a[0]+(b[0]-a[0])*k/m, a[1]+(b[1]-a[1])*k/m)) for k in range(1, m)]
    return fl, proj, None


def _tangent_bridge(before, after, step, w, polys):
    """Стороны дыры — две дуги (или дуга и прямая), между ними короткий прямой штрих (ножка g
    между хвостом и кольцом): общая касательная к обеим. Заполнение — дуга стороны «до» до
    точки касания T1, прямая T1-T2, дуга стороны «после» от T2. (before, (fl, proj, None)) | None."""
    def take(seq, L):
        out = [seq[0]]; s = 0.0
        for p, q in zip(seq[:-1], seq[1:]):
            s += math.dist(p, q)
            if s > L:
                break
            out.append(q)
        return out
    A = take(before[::-1], 2.0*w); B = take(after, 2.0*w)
    fa = _fit_pts(A, w); fb = _fit_pts(B, w)
    if fa is None or fb is None or (fa[0] == 'L' and fb[0] == 'L') or len(A) < 2 or len(B) < 2:
        return None
    a = before[-1]; b = after[0]
    ta = _norm((A[0][0]-A[1][0], A[0][1]-A[1][1]))     # направление хода в конце «до»
    tb = _norm((B[1][0]-B[0][0], B[1][1]-B[0][1]))     # направление хода в начале «после»
    def circ(f):
        return (f[1], f[2], f[3]) if f[0] == 'C' else None
    cands = []
    ca, cb = circ(fa), circ(fb)
    if ca and cb:
        D = (cb[0]-ca[0], cb[1]-ca[1]); L0 = math.hypot(*D)
        if L0 < 1e-9:
            return None
        Dn = (D[0]/L0, D[1]/L0); Dp = (-Dn[1], Dn[0])
        for sa, sb in ((1, 1), (1, -1), (-1, 1), (-1, -1)):
            # центры по нормали n от точек касания: c1 = T1 + sa*r1*n, c2 = T2 + sb*r2*n
            k = (sa*ca[2] - sb*cb[2]) / L0
            if abs(k) > 1:
                continue
            for sg in (1, -1):
                n = (k*Dn[0] + sg*math.sqrt(1-k*k)*Dp[0], k*Dn[1] + sg*math.sqrt(1-k*k)*Dp[1])
                T1 = (ca[0] - sa*ca[2]*n[0], ca[1] - sa*ca[2]*n[1])
                T2 = (cb[0] - sb*cb[2]*n[0], cb[1] - sb*cb[2]*n[1])
                cands.append((T1, T2))
    else:
        lf, cf_ = (fa, cb) if cb else (fb, ca)
        _, mx, my, dx, dy = lf
        t0 = (cf_[0]-mx)*dx + (cf_[1]-my)*dy
        F = (mx + dx*t0, my + dy*t0)
        if abs(math.dist(F, cf_[:2]) - cf_[2]) > 0.1*w:
            return None
        if ca:
            cands.append((F, _fit_proj_c(fb, a)))
        else:
            cands.append((_fit_proj_c(fa, b), F))
    best = None
    for T1, T2 in cands:
        Lt = math.dist(T1, T2)
        if Lt < 0.3*w:
            continue
        d = _norm((T2[0]-T1[0], T2[1]-T1[1]))
        # касательная к дуге «до» в T1 и к «после» в T2 — по направлению прямой
        t1 = _fit_tangent(fa, T1, ta); t2 = _fit_tangent(fb, T2, tb)
        if d[0]*t1[0] + d[1]*t1[1] < 0.98 or d[0]*t2[0] + d[1]*t2[1] < 0.98:
            continue
        # точки касания не дальше 1w за краями дыры (короткая ножка: зона стыка её почти съела)
        if (T1[0]-a[0])*ta[0] + (T1[1]-a[1])*ta[1] < -1.0*w or (b[0]-T2[0])*tb[0] + (b[1]-T2[1])*tb[1] < -1.0*w:
            continue
        if not _poly_inside(polys, [T1, T2], 0.5*step):
            continue
        sc = math.dist(a, T1) + math.dist(T2, b)
        if best is None or sc < best[0]:
            best = (sc, T1, T2)
    if best is None:
        return None
    _, T1, T2 = best
    d = _norm((T2[0]-T1[0], T2[1]-T1[1])); Lt = math.dist(T1, T2)
    def onl(p, t):
        return (T1[0] + d[0]*t, T1[1] + d[1]*t)
    def par(p):
        return (p[0]-T1[0])*d[0] + (p[1]-T1[1])*d[1]
    def projA(p):
        t = par(p)
        return _fit_proj_c(fa, p) if t < 0 else onl(p, min(t, Lt))
    def projB(p):
        t = par(p)
        return _fit_proj_c(fb, p) if t > Lt else onl(p, max(t, 0.0))
    before = list(before)
    reach = max(0.8*w, -par(before[-1]) + 0.3*w)
    s = 0.0; i = len(before)-1
    while i >= 0:
        before[i] = projA(before[i])
        if i > 0:
            s += math.dist(before[i], before[i-1])
        if s > reach:
            break
        i -= 1
    a = before[-1]; b2 = projB(b)
    ta_, tb_ = par(a), par(b2)
    fl = []
    if ta_ < 0:
        fl += _fit_between(fa, a, T1, step) + [T1]
    t0, t1 = max(ta_, 0.0), min(tb_, Lt)
    m = max(1, int((t1 - t0)/step))
    fl += [onl(None, t0 + (t1-t0)*k/m) for k in range(1, m)]
    if tb_ > Lt:
        fl += [T2] + _fit_between(fb, T2, b2, step)
    return before, (fl, projB, None)


def _arc_fill_single(before, after, step, w):
    """Одна окружность/прямая через обе стороны дыры (окна 2.5w -> 1.5w -> 0.8w).
    Возвращает (новый before, (fl, proj, None)) или None."""
    def take(seq, L):
        out = [seq[0]]; s = 0.0
        for a, b in zip(seq[:-1], seq[1:]):
            s += math.dist(a, b)
            if s > L:
                break
            out.append(b)
        return out
    before = list(before)
    proj = None
    for win in (2.5, 1.5, 0.8):
        A = take(before[::-1], win*w); B = take(after, win*w)
        if len(A) < 3 or len(B) < 3:
            continue
        P = A + B
        c = _fit_circle(P)
        if c is not None and c[2] < 25*w and c[3] < 0.02*w and _bent(P, c[2], w):
            cx, cy, R, _ = c
            def proj(q, cx=cx, cy=cy, R=R):
                d = math.dist(q, (cx, cy))
                return (cx + (q[0]-cx)*R/d, cy + (q[1]-cy)*R/d) if d > 1e-9 else q
            break
        mx, my, dx, dy, rms = _fit_line(P)
        if rms < 0.02*w:
            def proj(q, mx=mx, my=my, dx=dx, dy=dy):
                t = (q[0]-mx)*dx + (q[1]-my)*dy
                return (mx + dx*t, my + dy*t)
            break
    if proj is None:
        return None
    # подправить хвост "до" и сам обход
    s = 0.0; i = len(before)-1
    while i >= 0:
        before[i] = proj(before[i])
        if i > 0:
            s += math.dist(before[i], before[i-1])
        if s > 0.8*w:
            break
        i -= 1
    a = before[-1]; b = proj(after[0])
    m = max(1, int(math.dist(a, b)/step))
    fl = [proj((a[0]+(b[0]-a[0])*k/m, a[1]+(b[1]-a[1])*k/m)) for k in range(1, m)]
    return before, (fl, proj, None)


def _fit_tangent(fit, p, ref):
    """Единичная касательная к фиту в точке p (p на фите), ориентированная по ref."""
    if fit[0] == 'L':
        d = (fit[3], fit[4])
    else:
        d = _norm((-(p[1]-fit[2]), p[0]-fit[1]))
    if d[0]*ref[0] + d[1]*ref[1] < 0:
        d = (-d[0], -d[1])
    return d


def _fit_pts(P, w, kc=0.02, kl=0.02):
    """Строгий фит набора точек: прямая (rms<kl*w), иначе окружность (R<25w, rms<kc*w,
    заметно изогнута). Прямая — первой: прямой штрих с мелким изломом у развилки
    окружность «съедает» огромным радиусом. ('C', cx, cy, R) | ('L', mx, my, dx, dy) | None."""
    if len(P) < 3:
        return None
    mx, my, dx, dy, rms = _fit_line(P)
    if rms < kl*w:
        return ('L', mx, my, dx, dy)
    if len(P) >= 4:
        c = _fit_circle(P)
        if c is not None and c[2] < 25*w and c[3] < kc*w and _bent(P, c[2], w):
            return ('C', c[0], c[1], c[2])
    return None


def _snap_edge(fit, P, w, polys, bcorners, span):
    """Короткая прямая сторона (span < 1.2w): направление по 3-5 точкам ненадёжно —
    заменить прямой вдоль кромки контура у ближайшего вогнутого угла (если та в 12 град)."""
    if fit[0] != 'L' or span >= 1.2*w:
        return fit
    best = None
    for c in bcorners:
        if c[2] or math.dist(c[:2], P[0]) > 2.0*w:
            continue
        for e in _edge_lines(c, w, polys):
            if abs(e[3]*fit[3] + e[4]*fit[4]) >= math.cos(math.radians(12)):
                d = sum(math.dist(p, _fit_proj_c(e, p)) for p in P) / len(P)
                if d < 0.25*w and (best is None or d < best[0]):
                    best = (d, e)
    return best[1] if best else fit


def _corner_fill(before, after, step, w, polys, bcorners=()):
    """Сквозная линия ломается в развилке под углом (Y, V-стыки): стороны дыры подогнаны
    по отдельности, их пересечение X — угол оси; заполнение — по кривым сторон до X.
    Возвращает (fl, proj_after, индекс X в fl) или (None, None, None)."""
    def take(seq, L):
        out = [seq[0]]; s = 0.0
        for p, q in zip(seq[:-1], seq[1:]):
            s += math.dist(p, q)
            if s > L:
                break
            out.append(q)
        return out
    A = take(before[::-1], 2.0*w); B = take(after, 2.0*w)
    fa = _fit_pts(A, w); fb = _fit_pts(B, w)
    if fa is None or fb is None:
        return None, None, None
    fa = _snap_edge(fa, A, w, polys, bcorners, math.dist(A[0], A[-1]))
    fb = _snap_edge(fb, B, w, polys, bcorners, math.dist(B[0], B[-1]))
    a = before[-1]; b = after[0]
    X = _fit_intersect(fa, fb, ((a[0]+b[0])/2, (a[1]+b[1])/2))
    if X is None:
        X = _fit_touch(fa, fb, w)          # касание прямой и дуги (перекладина B -> чаша)
    if X is None or not any(inside(p, X[0], X[1]) for p in polys):
        return None, None, None
    XA = _fit_proj_c(fa, X); XB = _fit_proj_c(fb, X)
    gap = math.dist(a, b)
    if math.dist(a, XA) + math.dist(XB, b) > 2.0*gap + w:
        return None, None, None
    # X впереди по ходу от a и позади b
    da = _norm((XA[0]-A[1][0], XA[1]-A[1][1])) if len(A) > 1 else None
    db = _norm((B[1][0]-XB[0], B[1][1]-XB[1])) if len(B) > 1 else None
    if da is None or db is None:
        return None, None, None
    ta = _norm((a[0]-A[1][0], a[1]-A[1][1])); tb = _norm((B[1][0]-b[0], B[1][1]-b[1]))
    if da[0]*ta[0] + da[1]*ta[1] < 0.7 or db[0]*tb[0] + db[1]*tb[1] < 0.7:
        return None, None, None
    # поворот в X заметный — угол; иначе — составная кривая (прямая + дуга по касательной:
    # перекладина B переходит в чашу, ножка в дугу), но только если стороны разной природы —
    # две почти соосные прямые или две похожие дуги лучше одним фитом (без излома)
    cosang = da[0]*db[0] + da[1]*db[1]
    corner = cosang <= math.cos(math.radians(25))
    if not corner:
        if fa[0] == fb[0] == 'L':
            return None, None, None
        if fa[0] == fb[0] == 'C':
            ka, kb = 1.0/fa[3], 1.0/fb[3]
            # по разные стороны (S) или заметно разная кривизна
            sa = (fa[1]-X[0])*(-da[1]) + (fa[2]-X[1])*da[0]
            sb = (fb[1]-X[0])*(-db[1]) + (fb[2]-X[1])*db[0]
            if sa*sb > 0 and max(ka, kb) < 1.5*min(ka, kb):
                return None, None, None
    # хвост "до" — на кривую своей стороны
    s = 0.0; i = len(before)-1
    while i >= 0:
        before[i] = _fit_proj_c(fa, before[i])
        if i > 0:
            s += math.dist(before[i], before[i-1])
        if s > 0.8*w:
            break
        i -= 1
    a = before[-1]
    left = _fit_between(fa, a, XA, step)
    fl = left + [X] + _fit_between(fb, XB, _fit_proj_c(fb, b), step)
    return fl, (lambda q, fb=fb: _fit_proj_c(fb, q)), (len(left) if corner else None)


def _fit_touch(f1, f2, w):
    """Точка касания двух фитов, если они не пересекаются, но почти касаются (< 0.15w)."""
    if f1[0] == 'C' and f2[0] == 'C':
        d = math.dist(f1[1:3], f2[1:3])
        if d < 1e-9:
            return None
        ra, rb = f1[3], f2[3]
        if abs(d - (ra+rb)) < 0.15*w or abs(d - abs(ra-rb)) < 0.15*w:
            ux, uy = (f2[1]-f1[1])/d, (f2[2]-f1[2])/d
            sg = 1.0 if (abs(d - (ra+rb)) < 0.15*w or ra > rb) else -1.0
            return (f1[1] + ux*ra*sg, f1[2] + uy*ra*sg)
        return None
    if f1[0] == 'C':
        f1, f2 = f2, f1
    if f1[0] != 'L' or f2[0] != 'C':
        return None
    _, mx, my, dx, dy = f1; _, cx, cy, R = f2
    t0 = (cx-mx)*dx + (cy-my)*dy
    fx, fy = mx + dx*t0, my + dy*t0
    if abs(math.dist((fx, fy), (cx, cy)) - R) < 0.15*w:
        return (fx, fy)
    return None


# ---- острые углы оси: пересечение подогнанных кривых двух сторон

def _side_fit(pts, i, closed, w, dirn):
    """Точки стороны dirn (-1/+1) от i на дуговом расстоянии [0.9w, 2.4w]; варианты фита.
    Возвращает (fits, k_in): fits — список гипотез по убыванию предпочтения:
    строгий фит (rms<0.02w) на окне 2.4w (окружность, иначе прямая), строгий на коротком
    окне 1.6w (прямая, иначе окружность — короткая перекладина перед дугой: Я, В),
    затем те же окна с нестрогими порогами. Какая гипотеза верна — решает проверка
    по вогнутому углу контура в _sharpen. fit = ('C', cx, cy, R) | ('L', mx, my, dx, dy);
    k_in — индекс ближайшей к i сохраняемой точки стороны (на ~0.9w). None — стороны не хватает."""
    n = len(pts); s = 0.0; j = i; P = []; S = []; k_in = None; cnt = 0
    while True:
        k = j + dirn
        if not closed and (k < 0 or k >= n):
            break
        k %= n
        s += math.dist(pts[j], pts[k]); j = k; cnt += 1
        if j == i or cnt > n:
            break
        if s >= 0.9*w:
            if k_in is None:
                k_in = j
            P.append(pts[j]); S.append(s)
        if s >= 2.4*w:
            break
    if k_in is None or len(P) < 2:
        return None
    fits = []
    for kc, kl in ((0.02, 0.02), (0.05, 0.08)):
        for win in (2.4, 1.6):
            Q = [p for p, sp in zip(P, S) if sp <= win*w]
            if len(Q) < 2:
                continue
            fc = fl = None
            if len(Q) >= 4:
                c = _fit_circle(Q)
                if c is not None and c[2] < 25*w and c[3] < kc*w and _bent(Q, c[2], w):
                    fc = ('C', c[0], c[1], c[2])
            mx, my, dx, dy, rms = _fit_line(Q)
            if rms < kl*w:
                fl = ('L', mx, my, dx, dy)
            f = (fc or fl) if win == 2.4 else (fl or fc)
            if f is not None and f not in fits:
                fits.append(f)
    if not fits:
        return None
    return fits, k_in


def _fit_intersect(f1, f2, near):
    """Пересечение двух фитов (прямая/окружность); ближайшее к near решение или None."""
    cands = []
    if f1[0] == 'L' and f2[0] == 'L':
        _, x1, y1, d1x, d1y = f1; _, x2, y2, d2x, d2y = f2
        den = d1x*d2y - d1y*d2x
        if abs(den) > 1e-9:
            t = ((x2-x1)*d2y - (y2-y1)*d2x) / den
            cands.append((x1 + d1x*t, y1 + d1y*t))
    elif f1[0] == 'C' and f2[0] == 'C':
        _, ax, ay, ra = f1; _, bx, by, rb = f2
        d = math.dist((ax, ay), (bx, by))
        if d > 1e-9 and abs(ra-rb) <= d <= ra+rb:
            a = (ra*ra - rb*rb + d*d) / (2*d)
            hh = math.sqrt(max(0.0, ra*ra - a*a))
            ux, uy = (bx-ax)/d, (by-ay)/d
            px, py = ax + ux*a, ay + uy*a
            cands += [(px - uy*hh, py + ux*hh), (px + uy*hh, py - ux*hh)]
    else:
        if f1[0] == 'C':
            f1, f2 = f2, f1
        _, mx, my, dx, dy = f1; _, cx, cy, R = f2
        # ближайшая точка прямой к центру
        t0 = (cx-mx)*dx + (cy-my)*dy
        fx, fy = mx + dx*t0, my + dy*t0
        dd = math.dist((fx, fy), (cx, cy))
        if dd <= R:
            hh = math.sqrt(max(0.0, R*R - dd*dd))
            cands += [(fx + dx*hh, fy + dy*hh), (fx - dx*hh, fy - dy*hh)]
    if not cands:
        return None
    return min(cands, key=lambda p: math.dist(p, near))


def _fit_between(fit, a, b, step):
    """Промежуточные точки вдоль фита от a до b (обе на фите), без концов."""
    out = []
    if fit[0] == 'L':
        L = math.dist(a, b); m = int(L/step)
        for k in range(1, m):
            t = k/m
            out.append((a[0] + (b[0]-a[0])*t, a[1] + (b[1]-a[1])*t))
    else:
        _, cx, cy, R = fit
        a0 = math.atan2(a[1]-cy, a[0]-cx); a1 = math.atan2(b[1]-cy, b[0]-cx)
        d = (a1 - a0 + math.pi) % (2*math.pi) - math.pi
        m = int(abs(d)*R/step)
        for k in range(1, m):
            t = a0 + d*k/m
            out.append((cx + R*math.cos(t), cy + R*math.sin(t)))
    return out


def _sharpen(pts, fixed, closed, w, polys, bcorners, h):
    """Острый угол оси напротив вогнутого угла контура: участок +-0.9w вокруг максимума
    поворота заменяется пересечением кривых (окружность/прямая), подогнанных к обеим
    сторонам на 0.9..2.4w, с промежуточными точками вдоль этих кривых."""
    pts = list(pts); fixed = set(fixed)
    bcorners = [c for c in bcorners if not c[2]]
    if not bcorners:
        return pts, fixed
    def arc_idx(i, L, step):
        # индекс точки на дуговом расстоянии L от i в сторону step; None, если конец
        # ломаной или по пути встретилась фиксированная точка (уже поставленный угол)
        s = 0.0; j = i
        while True:
            k = j + step
            if not closed and (k < 0 or k >= len(pts)):
                return None
            k %= len(pts)
            s += math.dist(pts[j], pts[k]); j = k
            if j in fixed:
                return None
            if s >= L:
                return j
            if j == i:
                return None
    skip = set(); guard = 0
    while guard < 50:
        guard += 1
        n = len(pts)
        best = None
        for i in range(n):
            if i in fixed or i in skip or not any(math.dist(c[:2], pts[i]) < 1.8*w for c in bcorners):
                continue
            a = arc_idx(i, 0.5*w, -1); b = arc_idx(i, 0.5*w, +1)
            if a is None or b is None:
                continue
            d1 = _norm((pts[i][0]-pts[a][0], pts[i][1]-pts[a][1]))
            d2 = _norm((pts[b][0]-pts[i][0], pts[b][1]-pts[i][1]))
            t = math.degrees(math.acos(max(-1, min(1, d1[0]*d2[0]+d1[1]*d2[1]))))
            if t > 45 and (best is None or t > best[0]):
                best = (t, i)
        if best is None:
            return pts, fixed
        t, i = best
        sa = _side_fit(pts, i, closed, w, -1); sb = _side_fit(pts, i, closed, w, +1)
        if sa is None or sb is None:
            skip.add(i); continue
        B, C = sa[1], sb[1]
        # перебор гипотез фита сторон: принимается первая пара, у которой пересечение X
        # близко и внутри контура, а вогнутый угол контура лежит на полуширине от обеих
        # осей (это их «выемка») — так отсекаются ложные фиты (перекладина+дуга как одна окружность)
        found = None
        near = [c for c in bcorners if math.dist(c[:2], pts[i]) < 1.8*w]
        el = [f for c in near for f in _edge_lines(c, w, polys)]
        # порядок гипотез: строгие фиты по оси, прямые по кромкам контура, нестрогие фиты
        ha = sa[0][:1] + el + sa[0][1:]; hb = sb[0][:1] + el + sb[0][1:]
        for fa in ha:
            # гипотеза должна проходить через свою сторону (кромочные прямые общие для обеих)
            if math.dist(pts[B], _fit_proj_c(fa, pts[B])) > 0.3*w:
                continue
            for fb in hb:
                if fa is fb or math.dist(pts[C], _fit_proj_c(fb, pts[C])) > 0.3*w:
                    continue
                X = _fit_intersect(fa, fb, pts[i])
                if X is None or math.dist(X, pts[i]) >= 2.0*w or not any(inside(p, X[0], X[1]) for p in polys):
                    continue
                for c in near:
                    da = math.dist(c[:2], _fit_proj_c(fa, c[:2])); db = math.dist(c[:2], _fit_proj_c(fb, c[:2]))
                    if not (0.3*w <= da <= 0.8*w and 0.3*w <= db <= 0.8*w):
                        continue
                    # и оси идут вдоль кромок контура, выходящих из этого угла (разные кромки)
                    if not ((_edge_match(fa, c, w, c[3]) and _edge_match(fb, c, w, c[4])) or
                            (_edge_match(fa, c, w, c[4]) and _edge_match(fb, c, w, c[3]))):
                        continue
                    # тугой завиток (рукописная s) снаружи круглый с большим радиусом: ось режет
                    # угол, и до внешней кромки от X по биссектрисе меньше полуширины; у настоящего
                    # угла (стык штрихов) — не меньше полуширины
                    bis = _norm((X[0]-c[0], X[1]-c[1]))
                    dout = _raycast_union(polys, X, bis)
                    if dout is not None and dout < 0.75*(da + db)/2:
                        continue
                    found = (fa, fb, X); break
                if found:
                    break
            if found:
                break
        if found is None:
            skip.add(i); continue
        fa, fb, X = found
        # у острого угла ось искажена дальше 0.9w: расширяем вырезаемый участок,
        # пока точки не лягут на кривую своей стороны (не дальше 3w)
        def extend(k, fit, dirn):
            s = 0.0; j = k
            while True:
                k2 = j + dirn
                if not closed and (k2 < 0 or k2 >= n):
                    return j
                k2 %= n
                if k2 == i or k2 in fixed:
                    return j
                s += math.dist(pts[j], pts[k2])
                if s > 2.1*w or math.dist(pts[k2], _fit_proj_c(fit, pts[k2])) <= 0.03*w:
                    return j
                j = k2
        B = extend(B, fa, -1); C = extend(C, fb, +1)
        # новая ломаная: ... B, [дуга B->X], X, [дуга X->C], C ...
        if closed and C < B:
            removed = set(range(B+1, n)) | set(range(0, C))
        else:
            removed = set(range(B+1, C))
        XA = _fit_proj_c(fa, X); XB = _fit_proj_c(fb, X)
        ins = _fit_between(fa, pts[B], XA, h) + [X] + _fit_between(fb, XB, pts[C], h)
        newp = []; newf = set()
        for q in range(n):
            if q in removed:
                continue
            if q in fixed:
                newf.add(len(newp))
            newp.append(pts[q])
            if q == B:
                for k, p in enumerate(ins):
                    if k == len(_fit_between(fa, pts[B], XA, h)):
                        newf.add(len(newp))
                    newp.append(p)
        pts, fixed = newp, newf; skip = set()
    return pts, fixed


def _fit_proj_c(fit, p):
    if fit[0] == 'C':
        _, cx, cy, R = fit
        d = math.dist(p, (cx, cy))
        return (cx + (p[0]-cx)*R/d, cy + (p[1]-cy)*R/d) if d > 1e-9 else p
    _, mx, my, dx, dy = fit
    t = (p[0]-mx)*dx + (p[1]-my)*dy
    return (mx + dx*t, my + dy*t)


def bbox(p):
    xs = [q[0] for q in p[0]]; ys = [q[1] for q in p[0]]
    return min(xs), min(ys), max(xs), max(ys)


def group_polys(polys):
    """Полигоны с пересекающимися рамками обрабатываются вместе (union-find)."""
    n = len(polys); par = list(range(n)); bb = [bbox(p) for p in polys]
    def f(i):
        while par[i] != i:
            par[i] = par[par[i]]; i = par[i]
        return i
    for i in range(n):
        for j in range(i+1, n):
            a, b = bb[i], bb[j]
            if a[0] <= b[2] and b[0] <= a[2] and a[1] <= b[3] and b[1] <= a[3]:
                par[f(i)] = f(j)
    groups = {}
    for i in range(n):
        groups.setdefault(f(i), []).append(polys[i])
    return list(groups.values())


def centerline_all(polys, prm=P):
    out = []
    for g in group_polys(polys):
        out += centerline(g, prm)
    return out


def _bent(P, R, w):
    """Дуга по точкам P заметно изогнута (стрелка > 0.03w) — иначе это прямая."""
    chord = math.dist(P[0], P[-1])
    return chord*chord/(8*R) > 0.03*w
