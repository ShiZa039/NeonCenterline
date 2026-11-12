"""Стенд LibreOffice: прогнать геометрию плагина (разделы '@@ИМЯ_BEGIN / '@@ИМЯ_END и отдельные
процедуры из src/*.bas) в LibreOffice Basic и сравнить результат с прототипом из proto/.

Зачем: VBA в CorelDRAW здесь не запустить, а LibreOffice Basic с «Option VBASupport 1» исполняет
тот же код. Раздел вырезается из исходника как есть (без правок), к нему добавляется Sub Main
с входными данными, Main печатает результат в файл, Python читает файл и сравнивает с proto/.

Как пользоваться (пример целиком — tools/stands/grom_hem.py):

    import lo_stand as S
    main = ['Sub Main', '    Dim g As Integer', '    g = FreeFile',
            '    Open "' + S.OUT + '" For Output As #g', ...,  '    Close #g', 'End Sub']
    code = S.module(S.section('SIMP_DECL'), S.section('GROMGEO'), S.proc('PutS'), *main)
    lines = S.run(code)        # строки, которые напечатал Main

Подводные камни (подробно — docs/pitfalls.md):
  - ошибка компиляции -> LibreOffice без экрана молча висит на окне ошибки; run() обрывает его
    по таймауту и бросает StandTimeout. Почти всегда не хватает процедуры или объявления, которые
    лежат вне вырезанного раздела (PutS, NormV, BzPt, *_DECL) — добавить их через proc()/section();
  - из разделов *GEO зовутся NeonProgStep / NeonProgPack / NeonProgBox / NeonProgCompare (окно хода
    работы) — module() сам добавляет пустые заглушки (PROG_STUBS);
  - LibreOffice прощает то, чего не прощает VBA (And/Or без сокращения, объявления после процедур,
    лишние аргументы): совпадение на стенде не заменяет tools/shadow_check.py и tools/call_check.py;
  - в LibreOffice присваивание массива может не копировать его (оба имени смотрят в один массив) —
    если стенд расходится с прототипом только в LibreOffice, проверить это первым делом.

Проверено на Linux (LibreOffice 24.2). На Windows пути берутся из %APPDATA% и Program Files —
не проверялось.
"""
import html, os, re, shutil, subprocess, sys, tempfile

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if os.name == 'nt':
    LIB = os.path.join(os.environ.get('APPDATA', ''), 'LibreOffice', '4', 'user', 'basic', 'Standard')
    SOFFICE = shutil.which('soffice') or r'C:\Program Files\LibreOffice\program\soffice.exe'
else:
    LIB = os.path.expanduser('~/.config/libreoffice/4/user/basic/Standard')
    SOFFICE = shutil.which('soffice') or 'soffice'
# файл, в который печатает Main (прямые слэши LibreOffice понимает и на Windows)
OUT = os.path.join(tempfile.gettempdir(), 'neon_lo_stand', 'out.txt').replace('\\', '/')

PROG_STUBS = '''
Public Sub NeonProgStep(ByVal i As Long)
End Sub
Public Sub NeonProgPack(ByVal k As Long, ByVal n As Long)
End Sub
Public Sub NeonProgBox(ByVal k As Long, ByVal n As Long)
End Sub
Public Sub NeonProgCompare(ByVal n As Long)
End Sub
'''

_cache = {}


class StandTimeout(RuntimeError):
    pass


def source(path='src/NeonTools.bas'):
    """Текст исходника (UTF-8), путь — от корня репозитория."""
    if path not in _cache:
        _cache[path] = open(os.path.join(ROOT, path), encoding='utf-8').read()
    return _cache[path]


def section(name, path='src/NeonTools.bas'):
    """Раздел между '@@ИМЯ_BEGIN и '@@ИМЯ_END (метки включительно — это комментарии)."""
    t = source(path)
    a = t.index(f"'@@{name}_BEGIN")
    b = t.index(f"'@@{name}_END")
    return t[a:b]


def proc(name, path='src/NeonTools.bas'):
    """Одна процедура (Sub/Function) целиком — для того, что зовётся из раздела, но лежит вне его."""
    t = source(path)
    m = re.search(r'^(?:Private |Public |Friend )?(Sub|Function) ' + re.escape(name) + r'\b', t, re.M)
    if not m:
        raise KeyError(f'{name} not found in {path}')
    end = 'End ' + m.group(1)
    k = t.index(end, m.start())
    return t[m.start():k + len(end)]


def _procs(path):
    return set(re.findall(r'^(?:Private |Public |Friend )?(?:Sub|Function) (\w+)', source(path), re.M))


def _code_only(code):
    # без строк в кавычках и комментариев — чтобы не тянуть процедуры, упомянутые в тексте
    code = re.sub(r'"[^"\n]*"', '""', code)
    return re.sub(r"'.*$", '', code, flags=re.M)


def module(*parts, stubs=True, deps=True, path='src/NeonTools.bas'):
    """Модуль LibreOffice Basic: части (строки или списки строк) подряд, объявления — вперёд.
    Порядок частей важен так же, как в VBA: объявления уровня модуля (*_DECL) — до процедур.
    deps: дописать процедуры из path, которые код зовёт, но которых в нём нет (PutS, NormV, BzPt…),
    и то, что зовут они сами. Если так притянется процедура, работающая с объектами Corel, модуль
    не скомпилируется — тогда deps=False и нужные процедуры перечислить самому через proc()."""
    body = []
    for p in parts:
        body.append('\n'.join(p) if isinstance(p, (list, tuple)) else p)
    code = 'Option VBASupport 1\nOption Explicit\n' + '\n'.join(body)
    if deps:
        known = _procs(path) - set(re.findall(r'Sub (\w+)\(', PROG_STUBS))
        while True:
            have = set(re.findall(r'^(?:Private |Public |Friend )?(?:Sub|Function) (\w+)', code, re.M))
            need = sorted((set(re.findall(r'\b([A-Za-z]\w*)\b', _code_only(code))) & known) - have)
            if not need:
                break
            code += '\n' + '\n'.join(proc(n, path) for n in need) + '\n'
    if stubs:
        # заглушки — только тех процедур хода работы, которых в модуле ещё нет
        for m in re.finditer(r'Public Sub (\w+)\(.*?\)\nEnd Sub\n', PROG_STUBS):
            if not re.search(r'^(?:Private |Public )?Sub ' + m.group(1) + r'\b', code, re.M):
                code += '\n' + m.group(0)
    return code + '\n'


def run(code, timeout=180, entry='Main'):
    """Положить модуль в библиотеку Standard как Module1, запустить entry в LibreOffice без экрана,
    вернуть строки файла OUT. Прежний Module1 стенда перезаписывается."""
    os.makedirs(LIB, exist_ok=True)
    os.makedirs(os.path.dirname(OUT), exist_ok=True)
    xml = ('<?xml version="1.0" encoding="UTF-8"?>\n<!DOCTYPE script:module PUBLIC "-//OpenOffice.org//DTD OfficeDocument 1.0//EN" "module.dtd">\n'
           '<script:module xmlns:script="http://openoffice.org/2000/script" script:name="Module1" script:language="StarBasic">'
           + html.escape(code, quote=False) + '</script:module>')
    open(os.path.join(LIB, 'Module1.xba'), 'w', encoding='utf-8').write(xml)
    xlb = os.path.join(LIB, 'script.xlb')
    if not os.path.exists(xlb) or 'library:name="Module1"' not in open(xlb, encoding='utf-8').read():
        open(xlb, 'w', encoding='utf-8').write(
            '<?xml version="1.0" encoding="UTF-8"?>\n<!DOCTYPE library:library PUBLIC "-//OpenOffice.org//DTD OfficeDocument 1.0//EN" "library.dtd">\n'
            '<library:library xmlns:library="http://openoffice.org/2000/library" library:name="Standard" library:readonly="false" library:passwordprotected="false">\n'
            ' <library:element library:name="Module1"/>\n</library:library>\n')
    if os.path.exists(OUT):
        os.remove(OUT)
    try:
        subprocess.run([SOFFICE, '--headless', f'macro:///Standard.Module1.{entry}'], timeout=timeout,
                       stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    except subprocess.TimeoutExpired:
        raise StandTimeout('LibreOffice did not finish: most likely a compile error in the module '
                           '(missing procedure/declaration, or a syntax error). See docs/pitfalls.md')
    finally:
        if os.name == 'nt':
            subprocess.run('taskkill /F /IM soffice.bin >NUL 2>&1', shell=True)
        else:
            subprocess.run('pkill -9 -x soffice.bin', shell=True)
    if not os.path.exists(OUT):
        raise RuntimeError('Main did not write ' + OUT + ' (runtime error before Open, or wrong entry name)')
    return open(OUT, encoding='utf-8', errors='replace').read().splitlines()


def put_segs(segs, count='NS'):
    """Строки Main, которые кладут сегменты Безье ((x0,y0),(x1,y1),(x2,y2),(x3,y3),is_line) в SX/SY
    через PutS — как это делает PlaceGrommets / ShapeGrommets перед GromChain."""
    out = [f'    {count} = 0', '    ReDim SX(0 To 4 * ' + str(max(len(segs), 1)) + ' + 63): ReDim SY(0 To 4 * ' + str(max(len(segs), 1)) + ' + 63)']
    for (p0, p1, p2, p3, _ln) in segs:
        out.append(f'    PutS {count}, {p0[0]!r}, {p0[1]!r}, {p1[0]!r}, {p1[1]!r}, {p2[0]!r}, {p2[1]!r}, {p3[0]!r}, {p3[1]!r}: {count} = {count} + 1')
    return out


if __name__ == '__main__':
    # быстрая проверка: компилируется ли раздел(ы) с заглушками — python tools/lo_stand.py GROMGEO
    names = sys.argv[1:] or ['GROMGEO']
    decl = [section(n) for n in ('SIMP_DECL', 'GROM_DECL', 'CLEAN_DECL', 'NEST_DECL')]
    main = ['Sub Main', '    Dim g As Integer', '    g = FreeFile', f'    Open "{OUT}" For Output As #g',
            '    Print #g, "compiled"', '    Close #g', 'End Sub']
    try:
        print(run(module(*decl, *[section(n) for n in names], *main))[0], ' '.join(names))
    except StandTimeout as e:
        print('TIMEOUT', ' '.join(names), '-', e)
        sys.exit(1)
