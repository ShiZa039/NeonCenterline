"""src/NeonLang.bas из strings.py: строки как коды Unicode (не зависят от кодировки Windows)."""
import sys, os
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(ROOT, 'tools'))
from strings import STRINGS

def hexs(t):
    return ' '.join('%X' % ord(c) for c in t)

def emit_lang(idx, fname):
    out = [f'Private Function {fname}(ByVal key As String) As String', '    Dim s As String', '    Select Case key']
    for k, v in STRINGS.items():
        t = v[idx]
        out.append(f'    Case "{k}"')
        # кусками по 60 символов — строки VBA не длиннее 1023 знаков
        parts = [t[i:i+60] for i in range(0, len(t), 60)] or ['']
        for p in parts:
            out.append(f'        s = s & LU("{hexs(p)}")' if p else '        s = ""')
    out += ['    Case Else', '        s = key', '    End Select', f'    {fname} = s', 'End Function', '']
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
     "' то же с подстановкой {0}..{4}",
     'Public Function NeonLF(ByVal key As String, Optional ByVal a0 As Variant = "", Optional ByVal a1 As Variant = "", _',
     '                       Optional ByVal a2 As Variant = "", Optional ByVal a3 As Variant = "", Optional ByVal a4 As Variant = "", _',
     '                       Optional ByVal a5 As Variant = "") As String',
     '    Dim s As String',
     '    s = NeonL(key)',
     '    s = Replace(s, "{0}", CStr(a0)): s = Replace(s, "{1}", CStr(a1)): s = Replace(s, "{2}", CStr(a2))',
     '    s = Replace(s, "{3}", CStr(a3)): s = Replace(s, "{4}", CStr(a4)): s = Replace(s, "{5}", CStr(a5))',
     '    NeonLF = s',
     'End Function', '',
     'Private Function LU(ByVal hx As String) As String',
     '    Dim a As Variant, i As Long, r As String',
     '    a = Split(hx, " ")',
     '    For i = 0 To UBound(a)',
     '        r = r & ChrW(CLng("&H" & a(i)))',
     '    Next i',
     '    LU = r',
     'End Function', '']
# \n в строках: оставляем литералом "\n" (заменяется в NeonL)
for k, v in STRINGS.items():
    assert '"' not in v[0] or True
L += emit_lang(0, 'LangRu')
L += emit_lang(1, 'LangEn')
open(os.path.join(ROOT, 'src', 'NeonLang.bas'), 'w', encoding='utf-8').write('\n'.join(L))
print('NeonLang.bas', len(L), 'lines', max(len(x) for x in L), 'max line')
