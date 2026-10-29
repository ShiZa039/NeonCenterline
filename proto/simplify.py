"""Упрощение контура (прототип для модуля NeonSimplify.bas).

Вход — подпути кривой: список сегментов (x0,y0,x1,y1,x2,y2,x3,y3) — прямые как кубики с
контрольными точками на отрезке — и признак «замкнут». Выход — такие же подпути с меньшим числом
узлов: кривая не отходит от исходной больше чем на tol (мм), острые углы остаются углами.

1. Углы. Узел — угол, если касательные по контрольным точкам расходятся > 10 град И поворот
   хорд на масштабе d = max(4 tol, 0.5 мм) по обе стороны > 25 град (зазубрина трассировки —
   не угол, тугое скругление без излома — не угол). Концы незамкнутого подпути — всегда границы.
2. Участки между углами (замкнутый без углов — целиком, от узла 0) разворачиваются в точки с
   шагом step = 4 tol (0.2..1 мм), исходные узлы — среди точек.
3. Каждый участок — фит кубиками Безье (Шнайдер, Graphics Gems 1990): касательные на концах
   участка — по исходным сегментам; хорда-параметризация, МНК для длин ручек, до 4 шагов Ньютона
   по параметру; не ложится — деление в точке наибольшей ошибки с общей касательной (гладко).
   Точки участка в пределах 0.5 tol от хорды — прямая. Участок, где узлов не стало меньше,
   остаётся из исходных сегментов (экономная кривая шрифта не портится).
"""
import math

CORNER_LOCAL = 10.0     # град: излом касательных по контрольным точкам
CORNER_CHORD = 25.0     # град: поворот хорд на масштабе d
COS_LOCAL = math.cos(math.radians(CORNER_LOCAL))
COS_CHORD = math.cos(math.radians(CORNER_CHORD))
MAX_REPARAM = 4


def _bz(s, t):
    mt = 1 - t
    a, b, c, d = mt*mt*mt, 3*mt*mt*t, 3*mt*t*t, t*t*t
    return (a*s[0] + b*s[2] + c*s[4] + d*s[6], a*s[1] + b*s[3] + c*s[5] + d*s[7])


def _norm(dx, dy):
    L = math.hypot(dx, dy)
    return (dx/L, dy/L) if L > 1e-12 else (0.0, 0.0)


def _tan_out(s):
    """Касательная в начале сегмента (направление движения)."""
    for k in (2, 4, 6):
        dx, dy = s[k] - s[0], s[k+1] - s[1]
        if math.hypot(dx, dy) > 1e-9:
            return _norm(dx, dy)
    return (0.0, 0.0)


def _tan_in(s):
    """Касательная в конце сегмента (направление движения)."""
    for k in (4, 2, 0):
        dx, dy = s[6] - s[k], s[7] - s[k+1]
        if math.hypot(dx, dy) > 1e-9:
            return _norm(dx, dy)
    return (0.0, 0.0)


def _seg_len(s):
    L = 0.0; p = (s[0], s[1])
    for k in range(1, 17):
        q = _bz(s, k/16); L += math.dist(p, q); p = q
    return L


def _dot(a, b):
    return a[0]*b[0] + a[1]*b[1]


def flatten(segs, step):
    """Точки подпути: (точки, индексы точек исходных узлов). Узел i — начало сегмента i;
    последний узел — конец последнего сегмента."""
    pts = [(segs[0][0], segs[0][1])]; node = [0]
    for s in segs:
        L = _seg_len(s)
        m = max(1, int(math.ceil(L/step)))
        if not _is_line(s):
            m = max(m, 4)
        for k in range(1, m+1):
            q = _bz(s, k/m)
            if math.dist(q, pts[-1]) > 1e-9:
                pts.append(q)
        node.append(len(pts)-1)
    return pts, node


def _is_line(s):
    dx, dy = s[6]-s[0], s[7]-s[1]; L = math.hypot(dx, dy)
    if L < 1e-9:
        return math.hypot(s[2]-s[0], s[3]-s[1]) < 1e-9 and math.hypot(s[4]-s[0], s[5]-s[1]) < 1e-9
    for k in (2, 4):
        if abs((s[k]-s[0])*dy - (s[k+1]-s[1])*dx)/L > 1e-3:
            return False
    return True


def _walk(pts, cum, i, d, dirn, closed):
    """Точка на расстоянии d по ломаной от точки i в сторону dirn (с обходом замкнутой)."""
    n = len(pts); total = cum[-1]
    if total <= 0:
        return pts[i]
    if closed:
        target = (cum[i] + dirn*d) % total
        # бинарный поиск
        lo, hi = 0, n-1
        while hi - lo > 1:
            mid = (lo+hi)//2
            if cum[mid] <= target:
                lo = mid
            else:
                hi = mid
        seg = cum[hi] - cum[lo]
        t = (target - cum[lo])/seg if seg > 0 else 0
        return (pts[lo][0] + (pts[hi][0]-pts[lo][0])*t, pts[lo][1] + (pts[hi][1]-pts[lo][1])*t)
    target = cum[i] + dirn*d
    if target <= 0:
        return pts[0]
    if target >= total:
        return pts[-1]
    lo, hi = 0, n-1
    while hi - lo > 1:
        mid = (lo+hi)//2
        if cum[mid] <= target:
            lo = mid
        else:
            hi = mid
    seg = cum[hi] - cum[lo]
    t = (target - cum[lo])/seg if seg > 0 else 0
    return (pts[lo][0] + (pts[hi][0]-pts[lo][0])*t, pts[lo][1] + (pts[hi][1]-pts[lo][1])*t)


def corners(segs, closed, pts, node, tol):
    """Индексы узлов-углов (0..len(segs)-1 для замкнутого; 1..len(segs)-1 для открытого)."""
    n = len(segs)
    cum = [0.0]
    for a, b in zip(pts[:-1], pts[1:]):
        cum.append(cum[-1] + math.dist(a, b))
    if closed and math.dist(pts[0], pts[-1]) < 1e-9:
        P = pts[:-1]; C = cum[:-1]; total = cum[-1]
    else:
        P = pts; C = cum; total = cum[-1]
    d = max(4*tol, 0.5)
    out = []
    rng = range(n) if closed else range(1, n)
    for i in rng:
        tin = _tan_in(segs[i-1]); tout = _tan_out(segs[i])
        if _dot(tin, tout) >= COS_LOCAL:
            continue
        k = node[i] if node[i] < len(P) else 0
        if closed:
            cc = C + [total]
            a = _walk(P + [P[0]], cc, k, d, -1, True); b = _walk(P + [P[0]], cc, k, d, 1, True)
        else:
            a = _walk(P, C, k, d, -1, False); b = _walk(P, C, k, d, 1, False)
        p = P[k]
        u = _norm(p[0]-a[0], p[1]-a[1]); v = _norm(b[0]-p[0], b[1]-p[1])
        if u == (0.0, 0.0) or v == (0.0, 0.0) or _dot(u, v) < COS_CHORD:
            out.append(i)
    return out


# ---------------------------------------------------------------------- фит Шнайдера
def _chord_param(P):
    u = [0.0]
    for a, b in zip(P[:-1], P[1:]):
        u.append(u[-1] + math.dist(a, b))
    L = u[-1]
    return [x/L for x in u] if L > 0 else [i/(len(P)-1) for i in range(len(P))]


def _gen_bezier(P, u, t1, t2):
    p0, p3 = P[0], P[-1]
    c00 = c01 = c11 = x0 = x1 = 0.0
    for (px, py), t in zip(P, u):
        mt = 1 - t
        b0, b1, b2, b3 = mt*mt*mt, 3*mt*mt*t, 3*mt*t*t, t*t*t
        a1 = (t1[0]*b1, t1[1]*b1); a2 = (t2[0]*b2, t2[1]*b2)
        c00 += a1[0]*a1[0] + a1[1]*a1[1]
        c01 += a1[0]*a2[0] + a1[1]*a2[1]
        c11 += a2[0]*a2[0] + a2[1]*a2[1]
        tx = px - (p0[0]*(b0+b1) + p3[0]*(b2+b3)); ty = py - (p0[1]*(b0+b1) + p3[1]*(b2+b3))
        x0 += a1[0]*tx + a1[1]*ty
        x1 += a2[0]*tx + a2[1]*ty
    det = c00*c11 - c01*c01
    seg = math.dist(p0, p3)
    al = ar = 0.0
    if abs(det) > 1e-12:
        al = (x0*c11 - x1*c01)/det
        ar = (c00*x1 - c01*x0)/det
    eps = 1e-6*seg
    if al < eps or ar < eps:
        al = ar = seg/3
    return (p0[0], p0[1], p0[0] + t1[0]*al, p0[1] + t1[1]*al,
            p3[0] + t2[0]*ar, p3[1] + t2[1]*ar, p3[0], p3[1])


def _max_err(P, u, b):
    me = -1.0; at = len(P)//2
    for i in range(1, len(P)-1):
        q = _bz(b, u[i])
        d = (q[0]-P[i][0])**2 + (q[1]-P[i][1])**2
        if d > me:
            me = d; at = i
    return math.sqrt(max(me, 0.0)), at


def _newton(b, p, t):
    mt = 1 - t
    q = _bz(b, t)
    # первая и вторая производные
    d1 = [3*(b[2]-b[0]), 3*(b[3]-b[1]), 3*(b[4]-b[2]), 3*(b[5]-b[3]), 3*(b[6]-b[4]), 3*(b[7]-b[5])]
    d2 = [2*(d1[2]-d1[0]), 2*(d1[3]-d1[1]), 2*(d1[4]-d1[2]), 2*(d1[5]-d1[3])]
    q1 = (d1[0]*mt*mt + d1[2]*2*mt*t + d1[4]*t*t, d1[1]*mt*mt + d1[3]*2*mt*t + d1[5]*t*t)
    q2 = (d2[0]*mt + d2[2]*t, d2[1]*mt + d2[3]*t)
    num = (q[0]-p[0])*q1[0] + (q[1]-p[1])*q1[1]
    den = q1[0]*q1[0] + q1[1]*q1[1] + (q[0]-p[0])*q2[0] + (q[1]-p[1])*q2[1]
    if abs(den) < 1e-12:
        return t
    t2 = t - num/den
    return min(1.0, max(0.0, t2))


def _line_ok(P, tol):
    a, b = P[0], P[-1]; L = math.dist(a, b)
    if L < 1e-9:
        return False
    dx, dy = (b[0]-a[0])/L, (b[1]-a[1])/L
    for p in P[1:-1]:
        t = (p[0]-a[0])*dx + (p[1]-a[1])*dy
        if t < -0.5*tol or t > L + 0.5*tol:
            return False
        if abs(-(p[0]-a[0])*dy + (p[1]-a[1])*dx) > 0.5*tol:
            return False
    return True


def _center_tan(P, i):
    k = max(1, min(3, i, len(P)-1-i))
    return _norm(P[i-k][0] - P[i+k][0], P[i-k][1] - P[i+k][1])


def fit_run(P, t1, t2, tol, out):
    """Фит участка точек P (касательные t1 — из начала внутрь, t2 — из конца внутрь)."""
    if len(P) < 2:
        return
    if len(P) == 2 or _line_ok(P, tol):
        a, b = P[0], P[-1]
        out.append((a[0], a[1], a[0] + (b[0]-a[0])/3, a[1] + (b[1]-a[1])/3,
                    a[0] + 2*(b[0]-a[0])/3, a[1] + 2*(b[1]-a[1])/3, b[0], b[1], True))
        return
    u = _chord_param(P)
    b = _gen_bezier(P, u, t1, t2)
    err, at = _max_err(P, u, b)
    if err <= tol:
        out.append(b + (False,)); return
    if err <= 4*tol:
        for _ in range(MAX_REPARAM):
            u = [u[0]] + [_newton(b, P[i], u[i]) for i in range(1, len(P)-1)] + [u[-1]]
            b = _gen_bezier(P, u, t1, t2)
            err, at = _max_err(P, u, b)
            if err <= tol:
                out.append(b + (False,)); return
    at = max(1, min(len(P)-2, at))
    tc = _center_tan(P, at)
    fit_run(P[:at+1], t1, tc, tol, out)
    fit_run(P[at:], (-tc[0], -tc[1]), t2, tol, out)


def _end_tan(P, at_start, dist):
    """Касательная конца участка по хорде до точки на расстоянии dist (внутрь участка)."""
    n = len(P); L = 0.0
    if at_start:
        k = 1
        while k < n - 1:
            L += math.dist(P[k-1], P[k])
            if L >= dist:
                break
            k += 1
        k = min(k, n - 1)
        return _norm(P[k][0] - P[0][0], P[k][1] - P[0][1])
    k = n - 2
    while k > 0:
        L += math.dist(P[k+1], P[k])
        if L >= dist:
            break
        k -= 1
    k = max(k, 0)
    return _norm(P[k][0] - P[-1][0], P[k][1] - P[-1][1])


def _run_tans(P, sa, sb, step):
    """Касательные концов участка (t1 — из начала внутрь, t2 — из конца внутрь): по контрольным
    точкам крайних сегментов; если сегмент короче 2 шагов (микросегмент, «шип») — по хорде."""
    if _seg_len(sa) >= 2*step:
        t1 = _tan_out(sa)
    else:
        t1 = _end_tan(P, True, 2*step)
    if _seg_len(sb) >= 2*step:
        ti = _tan_in(sb); t2 = (-ti[0], -ti[1])
    else:
        t2 = _end_tan(P, False, 2*step)
    return t1, t2


def simplify_subpath(segs, closed, tol):
    """Сегменты (x0..y3, is_line) упрощённого подпути."""
    segs = [s for s in segs if _seg_len(s) > 1e-9] or segs
    if closed and math.hypot(segs[-1][6]-segs[0][0], segs[-1][7]-segs[0][1]) > 1e-6:
        a, b = segs[-1][6:8], segs[0][0:2]       # замыкающий отрезок
        segs = segs + [(a[0], a[1], a[0] + (b[0]-a[0])/3, a[1] + (b[1]-a[1])/3,
                        a[0] + 2*(b[0]-a[0])/3, a[1] + 2*(b[1]-a[1])/3, b[0], b[1])]
    step = min(1.0, max(0.2, 4*tol))
    pts, node = flatten(segs, step)
    n = len(segs)
    if len(pts) < 3:
        return [tuple(s[:8]) + (_is_line(s),) for s in segs]      # вырожденный подпуть
    cs = corners(segs, closed, pts, node, tol)
    tol = 0.9*tol           # запас: ошибка фита считается по точкам, между ними кривая может отойти чуть больше
    out = []
    if closed:
        if not cs:
            # без углов: от узла 0 по кругу, касательная — по хорде между точками в 2 шагах
            # по обе стороны (устойчиво к микросегментам у начала)
            P = pts[:]
            if math.dist(P[0], P[-1]) > 1e-9:
                P.append(P[0])
            cum = [0.0]
            for a_, b_ in zip(P[:-1], P[1:]):
                cum.append(cum[-1] + math.dist(a_, b_))
            a = _walk(P, cum, 0, 2*step, -1, True); b = _walk(P, cum, 0, 2*step, 1, True)
            t = _norm(b[0] - a[0], b[1] - a[1])
            if t == (0.0, 0.0):
                t = _tan_out(segs[0])
            fit_run(P, t, (-t[0], -t[1]), tol, out)
            return out
        bounds = cs + [cs[0] + n]
        for a, b in zip(bounds[:-1], bounds[1:]):
            ia = node[a]; ib = node[b] if b < n else node[b - n] + (len(pts) - 1)
            idx = [k % (len(pts) - 1) for k in range(ia, ib + 1)]
            P = [pts[k] for k in idx]
            t1, t2 = _run_tans(P, segs[a % n], segs[(b - 1) % n], step)
            _fit_or_keep(P, t1, t2, tol, out, [segs[k % n] for k in range(a, b)])
        return out
    bounds = [0] + cs + [n]
    for a, b in zip(bounds[:-1], bounds[1:]):
        P = pts[node[a]:node[b] + 1]
        t1, t2 = _run_tans(P, segs[a], segs[b - 1], step)
        _fit_or_keep(P, t1, t2, tol, out, segs[a:b])
    return out


def _fit_or_keep(P, t1, t2, tol, out, orig):
    """Фит участка между углами; если узлов не стало меньше — исходные сегменты участка."""
    r = []
    fit_run(P, t1, t2, tol, r)
    if len(r) >= len(orig):
        r = [tuple(s[:8]) + (_is_line(s),) for s in orig]
    out.extend(r)


def simplify(subpaths, tol=0.1):
    """subpaths: [(segs, closed)] -> [(segs, closed)]; подпуть, у которого узлов не стало
    меньше (уже экономная кривая: окружность из 4 узлов), остаётся как был."""
    out = []
    for segs, closed in subpaths:
        r = simplify_subpath(segs, closed, tol)
        if len(r) >= len(segs):
            r = [tuple(s[:8]) + (_is_line(s),) for s in segs]
        out.append((r, closed))
    return out


def max_dev(sub_a, sub_b, step=0.01):
    """Хаусдорфово расстояние между двумя наборами подпутей (по плотным точкам)."""
    import numpy as np
    from scipy.spatial import cKDTree

    def dense(subs):
        out = []
        for segs, cl in subs:
            for s in segs:
                m = max(2, int(math.ceil(_seg_len(s)/step)))
                out += [_bz(s, k/m) for k in range(m+1)]
        return np.array(out)
    A = dense(sub_a); B = dense(sub_b)
    return max(cKDTree(B).query(A)[0].max(), cKDTree(A).query(B)[0].max())
