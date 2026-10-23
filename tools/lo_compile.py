"""Проверка компиляции всех модулей плагина в LibreOffice Basic (по модулю на файл .xba,
типы/объекты Corel заменены заглушками). Модуль считается скомпилированным, если Main
дошёл до вызова функции из него."""
import re, html, os, subprocess, sys
LIB = os.path.expanduser('~/.config/libreoffice/4/user/basic/Standard')
P = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
MODS = [('NeonLang', f'{P}/src/NeonLang.bas', 'x = NeonLangCode()'),
        ('NeonUI', f'{P}/src/NeonUI.bas', 'x = FmtNum(1.5)'),
        ('NeonUpdate', f'{P}/src/NeonUpdate.bas', 'x = VerCmp("1.2", "1.10")'),
        ('NeonAddon', f'{P}/src/NeonAddon.bas', 'x = NeonAddonProbe()'),
        ('NeonTools', f'{P}/src/NeonTools.bas', 'x = NeonToolsProbe()'),
        ('NeonCut', f'{P}/src/NeonCut.bas', 'x = NeonCutStep()'),
        ('NeonMono', f'{P}/src/NeonMono.bas', 'x = NeonMonoRunProbe()'),
        ('NeonScript', f'{P}/src/NeonScript.bas', 'x = NeonScriptRunProbe()'),
        ('NeonBuild', f'{P}/build/NeonBuild.bas', 'x = NeonBuildProbe()'),
        ('frmCode', f'{P}/src/frmNeonSettings.code.txt', 'x = FrmProbe()'),
        ('frmRun', f'{P}/src/frmNeonRun.code.txt', 'x = FrmRunProbe()')]
STUB_COREL = '''
Private Const cdrGroupShape As Long = 1
Private Const cdrCurveShape As Long = 2
Private Const cdrCurveSegment As Long = 3
Private Const cdrMillimeter As Long = 4
Private ActiveDocument As Object, ActivePage As Object, ActiveWindow As Object, Application As Object
Private Optimization As Boolean, EventsEnabled As Boolean
Private GMSManager As Object
Private Function ActiveSelectionRange() As Object
End Function
Private Function CreateShapeRange() As Object
End Function
Private Function CreateRGBColor(ByVal r As Long, ByVal g As Long, ByVal b As Long) As Object
End Function
Private Function GetSetting(ByVal a As String, ByVal b As String, ByVal c As String, Optional ByVal d As String = "") As String
    GetSetting = d
End Function
Private Sub SaveSetting(ByVal a As String, ByVal b As String, ByVal c As String, ByVal d As String)
End Sub
'''
FORM_STUB = '''
Private Me As Object, lblTool As Object, cboTool As Object, lblHint As Object, chkAsk As Object, lblLang As Object
Private cboLang As Object, lblType As Object, cboType As Object, lblCut As Object, cboCut As Object, chkFit As Object, lblTrim As Object, cboTrim As Object, lblRef As Object, cboRef As Object, chkUpd As Object, lblVer As Object, btnUpdate As Object, btnAbout As Object, btnOK As Object, btnCancel As Object
'''
names = []
for name, path, call in MODS:
    s = open(path, encoding='utf-8').read()
    s = "\n".join(l for l in s.splitlines() if not l.startswith("Attribute VB_"))
    for t in ['ShapeRange', 'Layer', 'Shape', 'Collection', 'Curve', 'SubPath', 'Segment']:
        s = re.sub(r'\bAs %s\b' % t, 'As Object', s)
    s = re.sub(r'\bAs cdrUnit\b', 'As Long', s)
    extra = STUB_COREL
    if name in ('frmCode', 'frmRun'):
        s = s.replace('Unload Me', "Me = Nothing")
        extra += FORM_STUB
    probe = call.split('= ')[1].split('(')[0]
    if probe.endswith('Probe'):
        extra += f'\nPublic Function {probe}() As Long\n    {probe} = 1\nEnd Function\n'
    code = "Option VBASupport 1\n" + s.replace('Option Explicit', 'Option Explicit' + extra, 1) if 'Option Explicit' in s else "Option VBASupport 1\n" + extra + s
    xml = ('<?xml version="1.0" encoding="UTF-8"?>\n<!DOCTYPE script:module PUBLIC "-//OpenOffice.org//DTD OfficeDocument 1.0//EN" "module.dtd">\n'
           f'<script:module xmlns:script="http://openoffice.org/2000/script" script:name="{name}" script:language="StarBasic">'
           + html.escape(code, quote=False) + '</script:module>')
    open(f'{LIB}/{name}.xba', 'w', encoding='utf-8').write(xml)
    names.append(name)
main = 'Sub Main\n    Dim x As Variant\n    On Error GoTo EH\n'
for name, path, call in MODS:
    main += f'    {call}\n    Open "/tmp/lotest/ok_{name}.txt" For Output As #1\n    Print #1, CStr(x)\n    Close #1\n'
main += '    Exit Sub\nEH:\n    Open "/tmp/lotest/err.txt" For Output As #1\n    Print #1, Err & " " & Error$ & " line " & Erl\n    Close #1\nEnd Sub\n'
xml = ('<?xml version="1.0" encoding="UTF-8"?>\n<!DOCTYPE script:module PUBLIC "-//OpenOffice.org//DTD OfficeDocument 1.0//EN" "module.dtd">\n'
       '<script:module xmlns:script="http://openoffice.org/2000/script" script:name="Module1" script:language="StarBasic">'
       + html.escape("Option VBASupport 1\n" + main, quote=False) + '</script:module>')
open(f'{LIB}/Module1.xba', 'w', encoding='utf-8').write(xml)
xlb = ('<?xml version="1.0" encoding="UTF-8"?>\n<!DOCTYPE library:library PUBLIC "-//OpenOffice.org//DTD OfficeDocument 1.0//EN" "library.dtd">\n'
       '<library:library xmlns:library="http://openoffice.org/2000/library" library:name="Standard" library:readonly="false" library:passwordprotected="false">\n'
       ' <library:element library:name="Module1"/>\n' + ''.join(f' <library:element library:name="{n}"/>\n' for n in names) + '</library:library>\n')
open(f'{LIB}/script.xlb', 'w').write(xlb)
os.makedirs('/tmp/lotest', exist_ok=True)
for f in os.listdir('/tmp/lotest'):
    if f.startswith('ok_') or f == 'err.txt':
        os.remove(f'/tmp/lotest/{f}')
try:
    subprocess.run(['soffice', '--headless', 'macro:///Standard.Module1.Main'], timeout=120,
                   stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, cwd='/tmp/lotest')
except subprocess.TimeoutExpired:
    print('TIMEOUT (compile error dialog?)')
subprocess.run('pkill -9 -x soffice.bin', shell=True)
for name, _, _ in MODS:
    ok = os.path.exists(f'/tmp/lotest/ok_{name}.txt')
    print(name, 'OK ' + open(f'/tmp/lotest/ok_{name}.txt').read().strip() if ok else 'NOT REACHED')
if os.path.exists('/tmp/lotest/err.txt'):
    print('RUNTIME ERR:', open('/tmp/lotest/err.txt').read())
# вернуть библиотеку стенда в исходное состояние (только Module1)
for n in names:
    os.remove(f'{LIB}/{n}.xba')
open(f'{LIB}/script.xlb', 'w').write(xlb.replace(''.join(f' <library:element library:name="{n}"/>\n' for n in names), ''))
