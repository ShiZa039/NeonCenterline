"""src/NeonAddon.bas: файлы панели (addon/NeonCenterline/*) внутри плагина + макрос InstallToolbar,
который кладёт их в папку Addons ТЕКУЩЕГО CorelDRAW (Application.Path), с запросом прав администратора."""
import os
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
A = os.path.join(ROOT, 'addon', 'NeonCenterline')
files = ['AppUI.xslt', 'UserUI.xslt', 'Coreldrw.addon', 'config.xml']
import base64
B64 = base64.b64encode(open(os.path.join(A, 'Resources.dll'), 'rb').read()).decode()
L = ['Attribute VB_Name = "NeonAddon"',
     "' =====================================================================",
     "'  NeonCenterline — установка панели инструментов в текущий CorelDRAW.",
     "'  ФАЙЛ СГЕНЕРИРОВАН tools/gen_addon.py из addon/NeonCenterline — не править вручную.",
     "'  Файлы панели хранятся здесь кодами Unicode; макрос пишет их во временную папку и",
     "'  копирует (с правами администратора) в Application.AddonPath\\NeonCenterline.",
     "' =====================================================================",
     'Option Explicit', '',
     'Public Sub InstallToolbar()',
     '    Dim tmp As String, dst As String, cmd As String, sh As Object, stage As String',
     '    On Error GoTo Fail',
     '    tmp = CreateObject("WScript.Shell").ExpandEnvironmentStrings("%TEMP%") & "\\NeonCenterline_addon"',
     '    MakeDir tmp']
for f in files:
    L.append(f'    stage = "{f}"')
    L.append(f'    WriteUtf8 tmp & "\\{f}", AddonText_{f.replace(".", "_")}()')
L.append('    stage = "Resources.dll"')
L.append('    WriteBase64 tmp & "\\Resources.dll", AddonBin_Resources_dll()')
L.append('    stage = "AddonPath"')
L += ['    dst = Application.AddonPath',
      '    If Right$(dst, 1) <> "\\" Then',
      '        dst = dst & "\\"',
      '    End If',
      '    dst = dst & "NeonCenterline"',
      '    cmd = "/c mkdir """ & dst & """ 2>nul & xcopy /Y /I """ & tmp & "\\*"" """ & dst & """"',
      '    stage = "copy"',
      '    Set sh = CreateObject("Shell.Application")',
      '    sh.ShellExecute "cmd.exe", cmd, "", "runas", 0',
      '    MsgBox NeonLF("addon_done", dst), vbInformation, NeonL("app_title")',
      '    Exit Sub',
      'Fail:',
      '    MsgBox NeonLF("err_fail", Err.Description & " [" & stage & "]"), vbExclamation, NeonL("app_title")',
      'End Sub', '',
      'Private Sub MakeDir(ByVal p As String)',
      '    On Error Resume Next',
      '    MkDir p',
      'End Sub', '',
      'Private Sub WriteUtf8(ByVal path As String, ByVal txt As String)',
      '    Dim st As Object, f As Integer',
      '    If Len(txt) = 0 Then',
      "        ' пустой файл-метка (Coreldrw.addon): ADODB.Stream не принимает пустую строку",
      '        f = FreeFile',
      '        Open path For Output As #f',
      '        Close #f',
      '        Exit Sub',
      '    End If',
      '    Set st = CreateObject("ADODB.Stream")',
      '    st.Type = 2',
      '    st.Charset = "utf-8"',
      '    st.Open',
      '    st.WriteText txt',
      '    st.SaveToFile path, 2',
      '    st.Close',
      'End Sub', '',
      'Private Sub WriteBase64(ByVal path As String, ByVal b64 As String)',
      '    Dim doc As Object, el As Object, st As Object',
      '    Set doc = CreateObject("MSXML2.DOMDocument")',
      '    Set el = doc.createElement("b")',
      '    el.DataType = "bin.base64"',
      '    el.Text = b64',
      '    Set st = CreateObject("ADODB.Stream")',
      '    st.Type = 1',
      '    st.Open',
      '    st.Write el.nodeTypedValue',
      '    st.SaveToFile path, 2',
      '    st.Close',
      'End Sub', '',
      'Private Function AU(ByVal hx As String) As String',
      '    Dim a As Variant, i As Long, r As String',
      '    If Len(hx) = 0 Then',
      '        Exit Function',
      '    End If',
      '    a = Split(hx, " ")',
      '    For i = 0 To UBound(a)',
      '        r = r & ChrW(CLng("&H" & a(i)))',
      '    Next i',
      '    AU = r',
      'End Function', '']
for f in files:
    t = open(os.path.join(A, f), encoding='utf-8').read().replace('\r\n', '\n')
    fn = 'AddonText_' + f.replace('.', '_')
    L.append(f'Private Function {fn}() As String')
    L.append('    Dim s As String')
    # кусками: ASCII-строки как литералы, остальное кодами
    for i in range(0, len(t), 70):
        p = t[i:i+70]
        if all(32 <= ord(c) < 127 or c == '\n' for c in p):
            q = p.replace('"', '""').replace('\n', '" & vbLf & "')
            L.append(f'    s = s & "{q}"')
        else:
            L.append('    s = s & AU("' + ' '.join('%X' % ord(c) for c in p) + '")')
    L += [f'    {fn} = s', 'End Function', '']
# двоичный Resources.dll — base64, по 100 символов; функция бьётся на части (лимит размера процедуры)
chunks = [B64[i:i+100] for i in range(0, len(B64), 100)]
parts = [chunks[i:i+150] for i in range(0, len(chunks), 150)]
L.append('Private Function AddonBin_Resources_dll() As String')
L.append('    AddonBin_Resources_dll = ' + ' & '.join(f'AddonBinPart{j}()' for j in range(len(parts))))
L += ['End Function', '']
for j, p in enumerate(parts):
    L.append(f'Private Function AddonBinPart{j}() As String')
    L.append('    Dim s As String')
    L += [f'    s = s & "{c}"' for c in p]
    L += [f'    AddonBinPart{j} = s', 'End Function', '']
open(os.path.join(ROOT, 'src', 'NeonAddon.bas'), 'w', encoding='utf-8').write('\n'.join(L))
print('NeonAddon.bas', len(L), 'lines, max', max(len(x) for x in L))
