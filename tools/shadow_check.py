"""VBA: локальная переменная/параметр с именем процедуры (регистр не важен), которую эта же
процедура вызывает, — в VBA ломает вызов («Expected array»). Плюс дубли Public-имён между модулями."""
import re, glob, sys, os
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
files = sys.argv[1:] or glob.glob(os.path.join(ROOT, 'src', '*.bas')) + glob.glob(os.path.join(ROOT, 'src', '*.txt')) + [os.path.join(ROOT, 'build', 'NeonBuild.bas')]
procs = {}; pubs = {}
for f in files:
    s = open(f, encoding='utf-8').read()
    for m in re.finditer(r'^(Public |Private )?(Sub|Function) (\w+)', s, re.M):
        procs.setdefault(m.group(3).lower(), set()).add(f)
        if m.group(1) != 'Private ':
            pubs.setdefault(m.group(3).lower(), []).append(f.split('/')[-1])
bad = 0
for k, v in pubs.items():
    if len(v) > 1 and k not in ('userform_initialize',):
        print('PUBLIC DUP', k, v); bad += 1
for f in files:
    s = open(f, encoding='utf-8').read()
    for m in re.finditer(r'^(?:Public |Private )?(?:Sub|Function) (\w+)\((.*?)\).*?$(.*?)^End (?:Sub|Function)', s, re.M | re.S):
        name, params, body = m.groups()
        locs = set(x.lower() for x in re.findall(r'(?:ByVal |ByRef |Optional )*(\w+)(?:\(\))? As', params))
        for d in re.findall(r'^\s*Dim (.*)$', body, re.M):
            locs |= set(x.lower() for x in re.findall(r'(\w+)(?:\([^)]*\))? As', d))
        for l in locs:
            if l in procs and re.search(r'\b' + l + r'\s*\(', body, re.I):
                print('SHADOW', f.split('/')[-1], name, l); bad += 1
print('issues', bad)
