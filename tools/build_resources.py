"""addon/NeonCenterline/Resources.dll (значки + подписи кнопок) и config.xml (связь GUID кнопки ->
значок/подпись). Нужны binutils-mingw-w64 (x86_64-w64-mingw32-windres/ld). Значки — из icons/*.ico,
пересобираются в ICO с 32-битными BMP-кадрами 16/24/32."""
import os, struct, subprocess, tempfile, uuid
from PIL import Image
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
A = os.path.join(ROOT, 'addon', 'NeonCenterline')
G = lambda k: str(uuid.uuid5(uuid.NAMESPACE_URL, 'sabcore/neoncenterline/' + k))
ITEMS = [('mono', 'neon_mono', 101, 'Неон: Моно'),
         ('script', 'neon_script', 102, 'Неон: Курсив'),
         ('settings', 'neon_settings', 103, 'Неон: настройки'),
         ('update', 'neon_settings', 103, 'Неон: проверить обновления'),
         ('about', 'neon_settings', 103, 'Неон: о программе')]

def bmp_entry(im):
    im = im.convert('RGBA'); w, h = im.size; px = im.load()
    xor = bytearray()
    for y in range(h - 1, -1, -1):
        for x in range(w):
            r, g, b, a = px[x, y]; xor += bytes((b, g, r, a))
    rowand = ((w + 31) // 32) * 4; andm = bytearray()
    for y in range(h - 1, -1, -1):
        row = bytearray(rowand)
        for x in range(w):
            if px[x, y][3] < 128:
                row[x // 8] |= 0x80 >> (x % 8)
        andm += row
    return struct.pack('<IiiHHIIiiII', 40, w, h * 2, 1, 32, 0, len(xor) + len(andm), 0, 0, 0, 0) + bytes(xor) + bytes(andm)

def bmp_ico(src, dst):
    big = Image.open(src); big.size = (48, 48); base = big.convert('RGBA')
    frames = [base.resize((s, s), Image.LANCZOS) for s in (16, 24, 32)]
    data = [bmp_entry(f) for f in frames]
    out = struct.pack('<HHH', 0, 1, len(frames)); off = 6 + 16 * len(frames)
    for f, d in zip(frames, data):
        out += struct.pack('<BBBBHHII', f.size[0], f.size[1], 0, 0, 1, 32, len(d), off); off += len(d)
    open(dst, 'wb').write(out + b''.join(data))

esc = lambda t: ''.join(c if ord(c) < 128 else '\\x%04X' % ord(c) for c in t)
with tempfile.TemporaryDirectory() as td:
    rc = ['LANGUAGE 0, 0']
    seen = set()
    for k, ico, rid, cap in ITEMS:
        if rid not in seen:
            bmp_ico(os.path.join(ROOT, 'icons', ico + '.ico'), os.path.join(td, ico + '.ico'))
            rc.append(f'{rid} ICON "{ico}.ico"'); seen.add(rid)
    rc += ['STRINGTABLE', 'BEGIN', '  20000 L"NeonCenterline"']
    rc += [f'  {20001 + i} L"{esc(cap)}"' for i, (k, ico, rid, cap) in enumerate(ITEMS)]
    rc += ['END', '']
    open(os.path.join(td, 'res.rc'), 'w').write('\n'.join(rc))
    subprocess.run(['x86_64-w64-mingw32-windres', '--preprocessor=cat', 'res.rc', '-O', 'coff', '-o', 'res.o'], cwd=td, check=True)
    subprocess.run(['x86_64-w64-mingw32-ld', '--dll', '-e', '0', '--subsystem', 'windows', '-s', '-o',
                    os.path.join(A, 'Resources.dll'), 'res.o'], cwd=td, check=True)
cfg = ['<?xml version="1.0"?>', '<config>', '  <resources>',
       f'    <resource name="{G("resources")}" path="Resources.dll">', '      <resourceMap>',
       f'        <resEntry id="{G("bar")}" string="20000"/>']
cfg += [f'        <resEntry id="{G(k)}" string="{20001 + i}" icon="{rid}"/>' for i, (k, ico, rid, cap) in enumerate(ITEMS)]
cfg += ['      </resourceMap>', '    </resource>', '  </resources>', '</config>', '']
open(os.path.join(A, 'config.xml'), 'w', encoding='utf-8').write('\n'.join(cfg))
print('Resources.dll', os.path.getsize(os.path.join(A, 'Resources.dll')), 'bytes; config.xml')
