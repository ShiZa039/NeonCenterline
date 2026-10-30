"""src/NeonAddon.bas: файлы панели (addon/NeonCenterline/*) внутри плагина + макрос InstallToolbar,
который кладёт их в папку Addons ТЕКУЩЕГО CorelDRAW (Application.Path), с запросом прав администратора."""
import os
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
A = os.path.join(ROOT, 'addon', 'NeonCenterline')
files = ['AppUI.xslt', 'UserUI.xslt', 'Coreldrw.addon', 'config.xml']
import base64, re
# версия набора кнопок: метка [NeonToolbar N] в AppUI.xslt должна совпадать с NEON_TOOLBAR_VER
# в src/NeonUI.bas — по ней плагин после обновления понимает, что панель надо сбросить
_ver = re.search(r'NEON_TOOLBAR_VER As String = "(\w+)"',
                 open(os.path.join(ROOT, 'src', 'NeonUI.bas'), encoding='utf-8').read()).group(1)
if f'[NeonToolbar {_ver}]' not in open(os.path.join(A, 'AppUI.xslt'), encoding='utf-8').read():
    raise SystemExit(f'AppUI.xslt: нет метки [NeonToolbar {_ver}] (NEON_TOOLBAR_VER в src/NeonUI.bas)')
# значки и подписи: Resources<N>.dll (tools/build_resources.py) — у каждого набора своё имя файла
DLL = f'Resources{_ver}.dll'
if not os.path.exists(os.path.join(A, DLL)):
    raise SystemExit(f'нет addon/NeonCenterline/{DLL}: python tools/build_resources.py')
_dll = open(os.path.join(A, DLL), 'rb').read()
B64 = base64.b64encode(_dll).decode()
_cfg = open(os.path.join(A, 'config.xml'), encoding='utf-8').read()
if f'path="{DLL}"' not in _cfg:
    raise SystemExit(f'config.xml ссылается не на {DLL}: python tools/build_resources.py')
# кнопки панели этого набора (без разделителей): по ним плагин проверяет, что сброс панели удался
_app = open(os.path.join(A, 'AppUI.xslt'), encoding='utf-8').read()
_bar = re.search(r'<commandBarData guid="3d22040b-[^"]*"[^>]*type="toolbar">(.*?)</commandBarData>', _app, re.S).group(1)
PROBE = []
for g in re.findall(r'<item guidRef="([^"]+)"', _bar):
    if g == '266435b4-6e53-460f-9fa7-f45be187d400':
        continue
    d = re.search(r'<itemData guid="%s"(.*?)/>' % g, _app, re.S).group(1)
    fl = re.search(r'flyoutBarRef="([^"]+)"', d)
    PROBE.append((g, fl.group(1) if fl else '', re.search(r'userCaption="([^"]+)"', d).group(1)))
L = ['Attribute VB_Name = "NeonAddon"',
     "' =====================================================================",
     "'  NeonCenterline — установка панели инструментов в текущий CorelDRAW.",
     "'  ФАЙЛ СГЕНЕРИРОВАН tools/gen_addon.py из addon/NeonCenterline — не править вручную.",
     "'  Файлы панели хранятся здесь кодами Unicode; макрос пишет их во временную папку и",
     "'  копирует (с правами администратора) в Application.AddonPath\\NeonCenterline.",
     "' =====================================================================",
     'Option Explicit', '',
     "' имя и размер файла значков этого набора кнопок (NeonToolbarFilesState сверяет файл в Addons)",
     'Public Function NeonAddonDll() As String',
     f'    NeonAddonDll = "{DLL}"',
     'End Function', '',
     'Public Function NeonAddonDllSize() As Long',
     f'    NeonAddonDllSize = {len(_dll)}',
     'End Function', '',
     "' кнопки панели «Неон» этого набора: GUID кнопки, GUID списка кнопки-группы, подпись — через Tab",
     "' (NeonBarProbe в NeonUI: удался ли сброс панели)",
     'Public Function NeonTbProbeItems() As Variant',
     '    NeonTbProbeItems = Array( _'] + [
     '        "%s" & vbTab & "%s" & vbTab & AU("%s")%s' % (g, fl, ' '.join('%X' % ord(c) for c in cap), ', _' if i < len(PROBE) - 1 else ')')
     for i, (g, fl, cap) in enumerate(PROBE)] + [
     'End Function', '',
     'Public Sub InstallToolbar()',
     '    Dim tmp As String, dst As String, bat As String, q As String, sh As Object, stage As String, lg As String',
     '    Dim t0 As Single, tn As Single, tc As Single, tl As Single, ok As Boolean, txt As String, pre As Long',
     '    On Error GoTo Fail',
     '    tmp = CreateObject("WScript.Shell").ExpandEnvironmentStrings("%TEMP%") & "\\NeonCenterline_addon"',
     '    MakeDir tmp',
     "    ' прежнее содержимое временной папки (файлы прошлых установок) не копировать",
     '    KillFile tmp & "\\*.*"']
for f in files:
    L.append(f'    stage = "{f}"')
    L.append(f'    WriteUtf8 tmp & "\\{f}", AddonText_{f.replace(".", "_")}()')
L.append(f'    stage = "{DLL}"')
L.append('    WriteBase64 tmp & "\\" & NeonAddonDll(), AddonBin_Resources_dll()')
L.append('    stage = "AddonPath"')
L += ['    dst = Application.AddonPath',
      '    If Right$(dst, 1) <> "\\" Then',
      '        dst = dst & "\\"',
      '    End If',
      '    dst = dst & "NeonCenterline"',
      "    ' Команды — в .cmd-файл рядом с временной папкой (не внутри: её содержимое копируется), и с",
      "    ' правами администратора запускается он сам: длинная команда через cmd /c при запуске с",
      "    ' повышением прав не выполнялась. Файл значков у каждого набора кнопок свой (Resources<N>.dll):",
      "    ' прежний занят работающим CorelDRAW и не перезаписывается, а новый ложится рядом; прежние",
      "    ' удаляются, когда не заняты (иначе — при следующей установке). Вывод — в журнал; при сбое .cmd",
      "    ' можно запустить вручную",
      '    lg = tmp & ".log"',
      '    bat = tmp & ".cmd"',
      '    KillFile lg',
      '    q = Chr$(34)',
      '    WriteUtf8NoBom bat, "@echo off" & vbCrLf & "chcp 65001 >nul" & vbCrLf & "(" & vbCrLf & _',
      '        "mkdir " & q & dst & q & " 2>nul" & vbCrLf & _',
      '        "for %%f in (" & q & dst & "\\Resources*.dll" & q & ") do if /i not " & q & "%%~nxf" & q & "==" & q & NeonAddonDll() & q & _',
      '        " del /f /q " & q & "%%f" & q & vbCrLf & _',
      '        "xcopy /Y /I /C " & q & tmp & "\\*" & q & " " & q & dst & q & vbCrLf & _',
      '        ") > " & q & lg & q & " 2>&1" & vbCrLf',
      '    stage = "copy"',
      "    ' файлы этого набора уже на месте (повторная установка) — ждать конца копирования по журналу",
      '    pre = NeonToolbarFilesState()',
      '    Set sh = CreateObject("Shell.Application")',
      '    sh.ShellExecute bat, "", "", "runas", 0',
      "    ' до перезапуска в памяти CorelDRAW старый состав панели — не сбрасывать её в этом сеансе",
      '    NeonTbInstalled = True',
      "    ' копирование идёт отдельно (сначала — запрос прав администратора): ждём, пока в Addons",
      "    ' появятся AppUI.xslt этого набора кнопок и его файл значков целиком; журнал есть, а файлов нет",
      "    ' 5 с — копирование не удалось; ничего за 60 с — запрос прав, скорее всего, отклонён",
      '    t0 = Timer',
      '    tc = -1',
      '    tl = -1',
      '    Do',
      '        DoEvents',
      '        tn = Timer',
      '        If tn < t0 Then',
      '            tn = tn + 86400',
      '        End If',
      '        If tn - tc >= 0.5 Then',
      '            tc = tn',
      '            If tl < 0 Then',
      '                If Len(Dir$(lg)) > 0 Then',
      '                    tl = tn',
      '                End If',
      '            End If',
      '            If NeonToolbarFilesState() = 2 Then',
      '                If pre <> 2 Then',
      '                    ok = True',
      '                ElseIf tl >= 0 Then',
      '                    ok = (tn - tl > 1.5)',
      '                End If',
      '                If ok Then',
      '                    Exit Do',
      '                End If',
      '            End If',
      '            If tl < 0 Then',
      "                ' журнала ещё нет — копирование не началось",
      '            ElseIf tn - tl > 5 Then',
      '                Exit Do',
      '            End If',
      '            If tn - t0 > 60 Then',
      '                Exit Do',
      '            End If',
      '        End If',
      '    Loop',
      '    If ok Then',
      '        MsgBox NeonLF("addon_done", dst), vbInformation, NeonL("app_title")',
      '    Else',
      '        txt = ReadLog(lg)',
      '        If Len(txt) = 0 Then',
      '            txt = NeonL("addon_nolog")',
      '        End If',
      '        NeonTbAskAgain',
      '        MsgBox NeonLF("addon_fail", dst, txt), vbExclamation, NeonL("app_title")',
      '    End If',
      '    Exit Sub',
      'Fail:',
      '    MsgBox NeonLF("err_fail", Err.Description & " [" & stage & "]"), vbExclamation, NeonL("app_title")',
      'End Sub', '',
      'Private Sub KillFile(ByVal p As String)',
      '    On Error Resume Next',
      '    Kill p',
      'End Sub', '',
      "' журнал установки (.cmd переключает консоль на UTF-8), не длиннее 800 знаков",
      'Private Function ReadLog(ByVal p As String) As String',
      '    Dim st As Object, t As String',
      '    On Error Resume Next',
      '    If Len(Dir$(p)) = 0 Then',
      '        Exit Function',
      '    End If',
      '    Set st = CreateObject("ADODB.Stream")',
      '    st.Type = 2',
      '    st.Charset = "utf-8"',
      '    st.Open',
      '    st.LoadFromFile p',
      '    t = st.ReadText',
      '    st.Close',
      '    t = Trim$(Replace(t, vbCr, ""))',
      '    If Len(t) > 800 Then',
      '        t = Left$(t, 800) & "..."',
      '    End If',
      '    ReadLog = t',
      'End Function', '',
      "' текст в UTF-8 без BOM (BOM в начале .cmd ломает первую строку)",
      'Private Sub WriteUtf8NoBom(ByVal path As String, ByVal txt As String)',
      '    Dim st As Object, bs As Object',
      '    Set st = CreateObject("ADODB.Stream")',
      '    st.Type = 2',
      '    st.Charset = "utf-8"',
      '    st.Open',
      '    st.WriteText txt',
      '    st.Position = 3',
      '    Set bs = CreateObject("ADODB.Stream")',
      '    bs.Type = 1',
      '    bs.Open',
      '    st.CopyTo bs',
      '    bs.SaveToFile path, 2',
      '    bs.Close',
      '    st.Close',
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
