"""VBA: число аргументов в вызовах процедур проекта (VBA: «Wrong number of arguments»; LibreOffice
такое пропускает, пока строка не выполнится), ключи строк NeonL/NeonLF и число подстановок NeonLF."""
import re, glob, os, sys
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(ROOT, 'tools'))
from strings import STRINGS
files = glob.glob(os.path.join(ROOT, 'src', '*.bas')) + glob.glob(os.path.join(ROOT, 'src', '*.txt')) + [os.path.join(ROOT, 'build', 'NeonBuild.bas')]
def clean(src):
    src = re.sub(r' _\r?\n\s*', ' ', src)
    out = []
    for l in src.splitlines():
        r = ''; ins = False
        for ch in l:
            if ch == '"': ins = not ins; r += ch; continue
            if ins: r += 'x' if ch not in '"' else ch; continue
            if ch == "'": break
            r += ch
        out.append(r)
    return out
def split_args(s):
    depth = 0; parts = ['']; ins = False
    for ch in s:
        if ch == '"': ins = not ins
        if not ins:
            if ch == '(': depth += 1
            elif ch == ')':
                if depth == 0: break
                depth -= 1
            elif ch == ',' and depth == 0:
                parts.append(''); continue
        parts[-1] += ch
    if len(parts) == 1 and parts[0].strip() == '': return []
    return parts
local = {}; public = {}
for f in files:
    for l in clean(open(f, encoding='utf-8').read()):
        m = re.match(r'\s*(Public |Private |Friend )?(?:Static )?(Sub|Function|Property (?:Get|Let|Set)) (\w+)\s*\((.*)\)', l)
        if m:
            ps = split_args(m.group(4))
            mn = sum(1 for p in ps if 'Optional' not in p and 'ParamArray' not in p)
            mx = 999 if any('ParamArray' in p for p in ps) else len(ps)
            local[(f, m.group(3).lower())] = (mn, mx, f)
            if m.group(1) != 'Private ':
                public.setdefault(m.group(3).lower(), (mn, mx, f))
class P:
    f = None
    def __contains__(self, nm): return (self.f, nm) in local or nm in public
    def __getitem__(self, nm): return local.get((self.f, nm)) or public[nm]
procs = P()
bad = 0
for f in files:
    procs.f = f
    for n, l in enumerate(clean(open(f, encoding='utf-8').read()), 1):
        if re.match(r'\s*(?:Public |Private )?(?:Sub|Function|Declare)', l): continue
        # function-style calls
        for m in re.finditer(r'(?<![.\w])(\w+)\(', l):
            nm = m.group(1).lower()
            if nm in procs:
                args = split_args(l[m.end():])
                mn, mx, _ = procs[nm]
                if not (mn <= len(args) <= mx):
                    print(f'{os.path.basename(f)}:{n}: {m.group(1)} {len(args)} args, expected {mn}..{mx}'); bad += 1
        # statement calls
        for st in l.split(':'):
            m = re.match(r'\s*(?:Call\s+)?(\w+)(\s+(?!=)(.*))?$', st)
            if m and m.group(1).lower() in procs and not re.match(r'\s*\w+\s*=', st):
                rest = (m.group(3) or '').strip()
                if rest.startswith('('):
                    continue
                args = split_args(rest) if rest else []
                mn, mx, _ = procs[m.group(1).lower()]
                if not (mn <= len(args) <= mx):
                    print(f'{os.path.basename(f)}:{n}: {m.group(1)} {len(args)} args, expected {mn}..{mx}'); bad += 1
# NeonLF: число подстановок
for f in files:
    src = re.sub(r' _\r?\n\s*', ' ', open(f, encoding='utf-8').read())
    for m in re.finditer(r'NeonLF\("(\w+)"', src):
        args = split_args(src[m.start() + 7:])
        key = m.group(1)
        if key not in STRINGS:
            print('unknown key', key); bad += 1; continue
        need = max([int(x) + 1 for x in re.findall(r'\{(\d)\}', STRINGS[key][0] + STRINGS[key][1])] or [0])
        if len(args) - 1 < need:
            print(f'{os.path.basename(f)}: NeonLF("{key}") {len(args)-1} values, needs {need}'); bad += 1
    for m in re.finditer(r'NeonL\("(\w+)"\)', src):
        if m.group(1) not in STRINGS:
            print('unknown key', m.group(1)); bad += 1
print('issues', bad)
