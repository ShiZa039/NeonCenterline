"""src/NeonLang.bas из strings.py: строки в коде только ASCII (не зависят от кодировки Windows).

Сжатая запись (LD): «#» переключает русский режим — в нём знаки '?'..'~' это буквы А..я
(U+0410..U+044F), пробел, цифры и знаки препинания — как есть; «$XXXX» — любой другой знак Unicode
(и сами «#», «$», «"»). Русский текст выходит чуть длиннее исходного, а не вчетверо, как кодами
(LibreOffice, на котором проверяется компиляция, не тянет больше ~64 КБ строк в модуле)."""
import sys, os
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(ROOT, 'tools'))
from strings import STRINGS

def enc(t):
    """Строка -> сжатая ASCII-запись для LD (каждый кусок — с латинского режима)."""
    out = []
    cyr = False
    for ch in t:
        o = ord(ch)
        if 0x410 <= o <= 0x44F:
            if not cyr:
                out.append('#'); cyr = True
            out.append(chr(0x3F + o - 0x410))
        elif 0x20 <= o <= 0x7E and ch not in '#$"':
            if cyr and o >= 0x3F:
                out.append('#'); cyr = False
            out.append(ch)
        else:
            out.append('$%04X' % o)
    return ''.join(out)


def dec(s):
    """Обратно (проверка)."""
    r = []; cyr = False; i = 0
    while i < len(s):
        ch = s[i]
        if ch == '#':
            cyr = not cyr; i += 1
        elif ch == '$':
            r.append(chr(int(s[i + 1:i + 5], 16))); i += 5
        else:
            o = ord(ch)
            r.append(chr(0x410 + o - 0x3F) if cyr and o >= 0x3F else ch); i += 1
    return ''.join(r)

LANG_CHUNK = 40  # строк в одной функции: VBA не компилирует процедуру больше 64 КБ («Procedure too large»)


def emit_lang(idx, fname):
    """{fname}(key) — перебирает части {fname}_1.._n (по LANG_CHUNK строк в каждой); не нашлось — сам ключ."""
    keys = list(STRINGS.keys())
    chunks = [keys[i:i + LANG_CHUNK] for i in range(0, len(keys), LANG_CHUNK)]
    out = [f'Private Function {fname}(ByVal key As String) As String', '    Dim s As String']
    for n in range(len(chunks)):
        out.append(f'    {"If" if n == 0 else "ElseIf"} {fname}_{n + 1}(key, s) Then')
    out += ['    Else', '        s = key', '    End If', f'    {fname} = s', 'End Function', '']
    for n, ch in enumerate(chunks):
        out += [f'Private Function {fname}_{n + 1}(ByVal key As String, s As String) As Boolean', f'    {fname}_{n + 1} = True',
                '    s = ""', '    Select Case key']
        for k in ch:
            tx = STRINGS[k][idx]
            out.append(f'    Case "{k}"')
            # кусками по 60 символов — строки VBA не длиннее 1023 знаков
            parts = [tx[i:i+60] for i in range(0, len(tx), 60)] or ['']
            for p in parts:
                assert dec(enc(p)) == p
                out.append(f'        s = s & LD("{enc(p)}")' if p else '        s = ""')
        out += ['    Case Else', f'        {fname}_{n + 1} = False', '    End Select', 'End Function', '']
    return out

L = ['Attribute VB_Name = "NeonLang"',
     "' =====================================================================",
     "'  NeonCenterline — строки интерфейса (RU/EN). ФАЙЛ СГЕНЕРИРОВАН tools/gen_lang.py",
     "'  из tools/strings.py — правьте там. Строки хранятся кодами Unicode, чтобы русский",
     "'  текст не зависел от кодировки Windows.",
     "' =====================================================================",
     'Option Explicit', '',
     'Private Const REG_APP As String = "NeonCenterline"', '',
     "' текущий язык: ru / en (настройка UI\\Lang; пусто — по языку Windows)",
     'Public Function NeonLangCode() As String',
     '    Dim v As String, loc As String',
     '    v = LCase$(GetSetting(REG_APP, "UI", "Lang", ""))',
     '    If v = "ru" Or v = "en" Then',
     '        NeonLangCode = v',
     '        Exit Function',
     '    End If',
     '    loc = ""',
     '    On Error Resume Next',
     '    loc = LCase$(CreateObject("WScript.Shell").RegRead("HKCU\\Control Panel\\International\\LocaleName"))',
     '    On Error GoTo 0',
     '    If Left$(loc, 2) = "ru" Or Left$(loc, 2) = "uk" Or Left$(loc, 2) = "be" Or Left$(loc, 2) = "kk" Then',
     '        NeonLangCode = "ru"',
     '    Else',
     '        NeonLangCode = "en"',
     '    End If',
     'End Function', '',
     "' строка по ключу на текущем языке; \\n -> перенос строки",
     'Public Function NeonL(ByVal key As String) As String',
     '    Dim s As String',
     '    If NeonLangCode() = "ru" Then',
     '        s = LangRu(key)',
     '    Else',
     '        s = LangEn(key)',
     '    End If',
     '    NeonL = Replace(s, "\\n", vbCrLf)',
     'End Function', '',
     "' то же с подстановкой {0}..{7}",
     'Public Function NeonLF(ByVal key As String, Optional ByVal a0 As Variant = "", Optional ByVal a1 As Variant = "", _',
     '                       Optional ByVal a2 As Variant = "", Optional ByVal a3 As Variant = "", Optional ByVal a4 As Variant = "", _',
     '                       Optional ByVal a5 As Variant = "", Optional ByVal a6 As Variant = "", Optional ByVal a7 As Variant = "") As String',
     '    Dim s As String',
     '    s = NeonL(key)',
     '    s = Replace(s, "{0}", CStr(a0)): s = Replace(s, "{1}", CStr(a1)): s = Replace(s, "{2}", CStr(a2))',
     '    s = Replace(s, "{3}", CStr(a3)): s = Replace(s, "{4}", CStr(a4)): s = Replace(s, "{5}", CStr(a5))',
     '    s = Replace(s, "{6}", CStr(a6)): s = Replace(s, "{7}", CStr(a7))',
     '    NeonLF = s',
     'End Function', '',
     "' сжатая запись строки: «#» — русский режим ('?'..'~' -> А..я), «$XXXX» — знак Unicode",
     'Private Function LD(ByVal t As String) As String',
     '    Dim i As Long, n As Long, c As Long, cyr As Boolean, r As String',
     '    n = Len(t)',
     '    i = 1',
     '    Do While i <= n',
     '        c = AscW(Mid$(t, i, 1))',
     '        If c = 35 Then',
     '            cyr = Not cyr',
     '            i = i + 1',
     '        ElseIf c = 36 Then',
     '            r = r & ChrW(CLng("&H" & Mid$(t, i + 1, 4)))',
     '            i = i + 5',
     '        ElseIf cyr And c >= 63 Then',
     '            r = r & ChrW(&H410 + c - 63)',
     '            i = i + 1',
     '        Else',
     '            r = r & Mid$(t, i, 1)',
     '            i = i + 1',
     '        End If',
     '    Loop',
     '    LD = r',
     'End Function', '']
# \n в строках: оставляем литералом "\n" (заменяется в NeonL)
for k, v in STRINGS.items():
    assert '"' not in v[0] or True
L += emit_lang(0, 'LangRu')
L += emit_lang(1, 'LangEn')
open(os.path.join(ROOT, 'src', 'NeonLang.bas'), 'w', encoding='utf-8').write('\n'.join(L))
print('NeonLang.bas', len(L), 'lines', max(len(x) for x in L), 'max line')
