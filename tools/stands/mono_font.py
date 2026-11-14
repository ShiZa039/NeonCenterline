"""«Моно» на всём тестовом шрифте: снимок до правки и сравнение после — правило «нулевой регрессии».

Ядро core/NeonCenterline.bas (раздел '@@CORE_BEGIN … '@@CORE_END и объявления перед ним) исполняется
в LibreOffice Basic; знаки font/SABNeonTest.ttf подаются ломаными, как их читает CorelDRAW.
Каждая строка алфавита — один запуск: буквы — отдельные группы ядра, как слово в CorelDRAW.

    python tools/stands/mono_font.py snap base.pkl          # до правки
    ... правка core/NeonCenterline.bas ...
    python tools/stands/mono_font.py snap new.pkl           # после
    python tools/stands/mono_font.py diff base.pkl new.pkl  # должно быть «diff 0»
    python tools/stands/mono_font.py diff base.pkl new.pkl --png diff.png   # картинка изменившихся строк

snap: --core <.bas> (по умолчанию core/NeonCenterline.bas — можно сохранить копию «до» и сравнить с ней),
      --sizes 70:6,100:8 (высота прописных, мм : диаметр фрезы, мм), --text "АБВ" (только эти знаки).
Два размера по всему шрифту — около 16 минут. Сравнение точное (до последнего знака координат).
14.11.2026: правки «Моно» для мелкого текста — diff 0 на 70:6, 100:8 и 40:4.
"""
import argparse, os, pickle, sys, tempfile, time
HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(os.path.dirname(HERE))
sys.path.insert(0, os.path.join(ROOT, 'tools'))
import lo_stand as S
from fontTools.ttLib import TTFont
from fontTools.pens.basePen import BasePen

FONT = os.path.join(ROOT, 'font', 'SABNeonTest.ttf')
TEXTS = ['ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'АБВГДЕЁЖЗИЙКЛМНОПРСТУФХЦЧШЩЪЫЬЭЮЯ', '0123456789?',
         'abcdefghijklmnopqrstuvwxyz', 'абвгдеёжзийклмнопрстуфхцчшщъыьэюя']
IN = os.path.join(tempfile.gettempdir(), 'neon_lo_stand', 'mono_in.txt').replace('\\', '/')

HARNESS = '''
Sub Main
    Dim f As Integer, ln As String, st As Long, p As Long, i As Long, k As Long, parts As Variant
    On Error GoTo EH
    CoreInitRings
    st = -1
    f = FreeFile
    Open "{IN}" For Input As #f
    Do While Not EOF(f)
        Line Input #f, ln
        If Left(ln, 1) = "T" Then
            parts = Split(ln, " ")
            CoreSetParams CDbl(parts(1)), CDbl(parts(2))
        ElseIf ln = "R" Then
            If st >= 0 Then CoreEndRing st
            st = CoreRingStart()
        ElseIf Len(ln) > 0 Then
            parts = Split(ln, " ")
            CoreAddPoint CDbl(parts(0)), CDbl(parts(1))
        End If
    Loop
    Close #f
    If st >= 0 Then CoreEndRing st
    CoreRun
    f = FreeFile
    Open "{OUT}" For Output As #f
    Print #f, "OK"{THIN}
    For p = 0 To CorePathCount() - 1
        Print #f, "P " & IIf(CorePathClosed(p), "1", "0") & " " & CorePathStart(p) & " " & CorePathGroup(p)
    Next p
    For i = 0 To CoreSegCount() - 1
        ln = "S"
        For k = 0 To 8
            ln = ln & " " & Str(CoreSeg(k, i))
        Next k
        Print #f, ln
    Next i
    Close #f
    Exit Sub
EH:
    f = FreeFile
    Open "{OUT}" For Output As #f
    Print #f, "ERR " & Err & " " & Error$ & " line " & Erl
    Close #f
End Sub
'''


class Flat(BasePen):
    """Контуры знака -> ломаные (кубики — 8 частей, квадратичные — 6), мм."""
    def __init__(self, gs, dx, sc):
        super().__init__(gs); self.rings = []; self.cur = None; self.dx = dx; self.sc = sc
    def _p(self, p): return ((p[0] + self.dx) * self.sc, p[1] * self.sc)
    def _moveTo(self, p): self.cur = [self._p(p)]; self.last = p
    def _lineTo(self, p): self.cur.append(self._p(p)); self.last = p
    def _curveToOne(self, a, b, c):
        p0 = self.last
        for i in range(1, 9):
            t = i / 8; mt = 1 - t
            self.cur.append(self._p((mt**3*p0[0]+3*mt*mt*t*a[0]+3*mt*t*t*b[0]+t**3*c[0],
                                     mt**3*p0[1]+3*mt*mt*t*a[1]+3*mt*t*t*b[1]+t**3*c[1])))
        self.last = c
    def _qCurveToOne(self, a, b):
        p0 = self.last
        for i in range(1, 7):
            t = i / 6; mt = 1 - t
            self.cur.append(self._p((mt*mt*p0[0]+2*mt*t*a[0]+t*t*b[0], mt*mt*p0[1]+2*mt*t*a[1]+t*t*b[1])))
        self.last = b
    def _closePath(self):
        if self.cur and len(self.cur) > 2:
            if self.cur[0] == self.cur[-1]:
                self.cur.pop()
            self.rings.append(self.cur)
        self.cur = None
    _endPath = _closePath


def text_rings(text, cap_mm):
    f = TTFont(FONT); gs = f.getGlyphSet(); cmap = f.getBestCmap()
    sc = cap_mm / f['OS/2'].sCapHeight
    x = 0; rings = []
    for ch in text:
        if ord(ch) not in cmap:
            continue
        g = cmap[ord(ch)]
        pen = Flat(gs, x, sc); gs[g].draw(pen)
        rings += pen.rings
        x += gs[g].width
    return rings


def core_module(core):
    s = open(core, encoding='utf-8').read().replace('\r\n', '\n')
    head = s[:s.index("'@@CORE_BEGIN")]
    body = s[s.index("'@@CORE_BEGIN"):s.index("'@@CORE_END")]
    head = '\n'.join(l for l in head.splitlines() if not l.startswith('Attribute VB_'))
    # у ядер до 14.11.2026 нет NeonMonoThinW / NeonMonoDropped — тогда их не печатаем
    thin = ' & " thinW=" & NeonMonoThinW() & " dropped=" & NeonMonoDropped()' if 'NeonMonoThinW' in s else ''
    return ('Option VBASupport 1\n' + head + '\n' + body +
            HARNESS.replace('{IN}', IN).replace('{OUT}', S.OUT).replace('{THIN}', thin))


def run_rings(code, rings, tool, tol=0.25):
    os.makedirs(os.path.dirname(IN), exist_ok=True)
    with open(IN, 'w') as f:
        f.write(f'T {tool} {tol}\n')
        for r in rings:
            f.write('R\n')
            for x, y in r:
                f.write(f'{x:.5f} {y:.5f}\n')
    L = S.run(code, timeout=1800)
    if not L or L[0].startswith('ERR'):
        raise RuntimeError(L[0] if L else 'empty output')
    starts, segs = [], []
    for l in L[1:]:
        a = l.split()
        if a[0] == 'P':
            starts.append((int(a[2]), a[1] == '1', int(a[3])))
        elif a[0] == 'S':
            segs.append(tuple(float(v) for v in a[1:]))
    paths = []
    for i, (s0, cl, g) in enumerate(starts):
        s1 = starts[i + 1][0] if i + 1 < len(starts) else len(segs)
        paths.append((g, cl, segs[s0:s1]))
    return paths, L[0]


def snap(a):
    code = core_module(a.core)
    sizes = [tuple(map(float, s.split(':'))) for s in a.sizes.split(',')]
    texts = [a.text] if a.text else TEXTS
    res = {}
    for cap, tool in sizes:
        for txt in texts:
            t0 = time.time()
            try:
                res[(cap, tool, txt)], info = run_rings(code, text_rings(txt, cap), tool)
            except Exception as e:
                res[(cap, tool, txt)], info = 'ERR ' + str(e), ''
            pickle.dump(res, open(a.out, 'wb'))
            n = len(res[(cap, tool, txt)]) if not isinstance(res[(cap, tool, txt)], str) else res[(cap, tool, txt)]
            print(f'{cap:g}:{tool:g} {txt[:4]}… paths {n} ({time.time() - t0:.0f} s) {info}', flush=True)
    print('saved', a.out)


def diff(a):
    A = pickle.load(open(a.a, 'rb')); B = pickle.load(open(a.b, 'rb'))
    keys = [k for k in A if k in B]
    bad = [k for k in keys if A[k] != B[k]]
    for k in bad:
        ga = {p[0] for p in A[k]} if not isinstance(A[k], str) else set()
        print(f'{k[0]:g}:{k[1]:g} {k[2][:6]}…  groups changed:',
              sorted(g for g in ga | {p[0] for p in B[k]} if [p for p in A[k] if p[0] == g] != [p for p in B[k] if p[0] == g])
              if not isinstance(A[k], str) and not isinstance(B[k], str) else (A[k] if isinstance(A[k], str) else B[k]))
    print(f'compared {len(keys)}, diff {len(bad)}')
    if bad and a.png:
        import matplotlib; matplotlib.use('Agg'); import matplotlib.pyplot as plt
        fig, axs = plt.subplots(2 * len(bad), 1, figsize=(24, 3.2 * 2 * len(bad)), squeeze=False)
        for i, k in enumerate(bad):
            for j, (lab, R) in enumerate((('before', A[k]), ('after', B[k]))):
                ax = axs[2 * i + j][0]
                for r in text_rings(k[2], k[0]):
                    ax.plot([p[0] for p in r + r[:1]], [p[1] for p in r + r[:1]], color='0.6', lw=0.6)
                for g, cl, segs in (R if not isinstance(R, str) else []):
                    for s in segs:
                        if s[0] == 0:
                            ax.plot([s[1], s[7]], [s[2], s[8]], 'r-', lw=1)
                        else:
                            ts = [t / 16 for t in range(17)]
                            xs = [(1-t)**3*s[1]+3*(1-t)**2*t*s[3]+3*(1-t)*t*t*s[5]+t**3*s[7] for t in ts]
                            ys = [(1-t)**3*s[2]+3*(1-t)**2*t*s[4]+3*(1-t)*t*t*s[6]+t**3*s[8] for t in ts]
                            ax.plot(xs, ys, 'b-', lw=1)
                ax.set_aspect('equal'); ax.axis('off'); ax.set_title(f'{lab} {k[0]:g}:{k[1]:g}', loc='left')
        plt.tight_layout(); plt.savefig(a.png, dpi=50)
        print('picture', a.png)
    sys.exit(1 if bad else 0)


if __name__ == '__main__':
    ap = argparse.ArgumentParser()
    sp = ap.add_subparsers(dest='cmd', required=True)
    p1 = sp.add_parser('snap'); p1.add_argument('out')
    p1.add_argument('--core', default=os.path.join(ROOT, 'core', 'NeonCenterline.bas'))
    p1.add_argument('--sizes', default='70:6,100:8'); p1.add_argument('--text', default='')
    p2 = sp.add_parser('diff'); p2.add_argument('a'); p2.add_argument('b'); p2.add_argument('--png', default='')
    a = ap.parse_args()
    snap(a) if a.cmd == 'snap' else diff(a)
