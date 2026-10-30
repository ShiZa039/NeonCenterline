"""addon/NeonCenterline/Resources<N>.dll (значки + подписи кнопок) и config.xml (связь GUID кнопки ->
значок/подпись). N — набор кнопок панели (NEON_TOOLBAR_VER в src/NeonUI.bas): у каждого набора своё
имя файла, поэтому установка нового набора не перезаписывает Resources*.dll, который держит
работающий CorelDRAW (занятый файл не заменить — значки остались бы старыми). Нужны binutils-mingw-w64 (x86_64-w64-mingw32-windres/ld). Значки — из icons/*.ico,
пересобираются в ICO с 32-битными BMP-кадрами 16/24/32."""
import os, struct, subprocess, tempfile, uuid
from PIL import Image
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
A = os.path.join(ROOT, 'addon', 'NeonCenterline')
import re, glob
VER = re.search(r'NEON_TOOLBAR_VER As String = "(\w+)"',
                open(os.path.join(ROOT, 'src', 'NeonUI.bas'), encoding='utf-8').read()).group(1)
DLL = f'Resources{VER}.dll'
for old in glob.glob(os.path.join(A, 'Resources*.dll')):
    if os.path.basename(old) != DLL:
        os.remove(old)
G = lambda k: str(uuid.uuid5(uuid.NAMESPACE_URL, 'sabcore/neoncenterline/' + k))
ITEMS = [('mono', 'neon_mono', 101, 'Неон: Моно'),
         ('script', 'neon_script', 102, 'Неон: Курсив (бета)'),
         ('settings', 'neon_settings', 103, 'Неон: настройки'),
         ('length', 'neon_length', 104, 'Неон: длина контуров и линий'),
         ('select', 'neon_select', 105, 'Неон: выделить центральные линии'),
         ('cut', 'neon_cut', 106, 'Неон: подогнать под кратность реза'),
         ('update', 'neon_update', 109, 'Неон: проверить обновления'),
         ('about', 'neon_about', 110, 'Неон: о программе'),
         ('simplify', 'neon_simplify', 107, 'Неон: упростить контур'),
         # кнопки-группы панели (выпадающие списки), с панели [NeonToolbar 3]
         ('lines-menu', 'neon_lines', 108, 'Неон: линии'),
         ('settings-menu', 'neon_settings', 103, 'Неон: настройки и помощь'),
         # с панели [NeonToolbar 5]: люверсы; с [NeonToolbar 6]: группа «Раскладка»
         ('grommets', 'neon_grommets', 111, 'Неон: люверсы'),
         ('layout-menu', 'neon_grommets', 111, 'Неон: раскладка')]

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
    rc += ['STRINGTABLE', 'BEGIN', '  20000 L"' + esc('Неон (NeonCenterline)') + '"']
    rc += [f'  {20001 + i} L"{esc(cap)}"' for i, (k, ico, rid, cap) in enumerate(ITEMS)]
    rc += ['END', '']
    open(os.path.join(td, 'res.rc'), 'w').write('\n'.join(rc))
    subprocess.run(['x86_64-w64-mingw32-windres', '--preprocessor=cat', 'res.rc', '-O', 'coff', '-o', 'res.o'], cwd=td, check=True)
    subprocess.run(['x86_64-w64-mingw32-ld', '--dll', '-e', '0', '--subsystem', 'windows', '-s', '-o',
                    os.path.join(A, DLL), 'res.o'], cwd=td, check=True)
cfg = ['<?xml version="1.0"?>', '<config>', '  <resources>',
       f'    <resource name="{G("resources")}" path="{DLL}">', '      <resourceMap>',
       f'        <resEntry id="{G("bar-v2")}" string="20000"/>']
cfg += [f'        <resEntry id="{G(k)}" string="{20001 + i}" icon="{rid}"/>' for i, (k, ico, rid, cap) in enumerate(ITEMS)]
cfg += ['      </resourceMap>', '    </resource>', '  </resources>', '</config>', '']
open(os.path.join(A, 'config.xml'), 'w', encoding='utf-8').write('\n'.join(cfg))
# ld с .rsrc больше 64 КБ пишет «.rsrc is too large» и не перекладывает секцию, а оставляет её как
# собрал windres (адреса данных — через обычные релокации); проверяем, что подписи и значки читаются
try:
    import pefile
    pe = pefile.PE(os.path.join(A, DLL))
    d = {e.id: e for e in pe.DIRECTORY_ENTRY_RESOURCE.entries}
    blk = d[6].directory.entries[0].directory.entries[0].data.struct
    raw = pe.get_data(blk.OffsetToData, blk.Size); o = 0; got = []
    for i in range(len(ITEMS) + 1):
        n = struct.unpack_from('<H', raw, o)[0]; got.append(raw[o + 2:o + 2 + 2 * n].decode('utf-16le')); o += 2 + 2 * n
    assert got == ['Неон (NeonCenterline)'] + [c for k, ico, r, c in ITEMS], got
    icons = {e.id: e for e in d[3].directory.entries}
    for g in d[14].directory.entries:
        e = g.directory.entries[0].data.struct; gd = pe.get_data(e.OffsetToData, e.Size)
        for i in range(struct.unpack_from('<H', gd, 4)[0]):
            w, iid = gd[6 + 14 * i], struct.unpack_from('<H', gd, 6 + 14 * i + 12)[0]
            ie = icons[iid].directory.entries[0].data.struct
            assert struct.unpack_from('<Ii', pe.get_data(ie.OffsetToData, 8)) == (40, w), (g.id, iid)
    print(DLL + ': подписи и значки читаются')
except ImportError:
    print('pefile не установлен — ' + DLL + ' не проверен')
print(DLL, os.path.getsize(os.path.join(A, DLL)), 'bytes; config.xml')
