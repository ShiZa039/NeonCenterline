"""Ядра плагина из рабочих модулей макроса: core/NeonCenterline.bas -> src/NeonMono.bas,
core/NeonScript.bas -> src/NeonScript.bas. Алгоритм не меняется: заменяется только макрос-точка
входа (с InputBox/MsgBox) на функцию без интерфейса, Core* делаются Private."""
import re, sys, os
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
V = os.path.join(ROOT, 'core'); S = os.path.join(ROOT, 'src')

ENTRY = r'''Public Function {fn}(ByVal tool As Double, nIn As Long, nOut As Long, tAll As Double, tCore As Double, errText As String) As Long
    ' Результат: 0 — готово, 1 — нет документа, 2 — пустое выделение, 3 — нет замкнутых контуров,
    ' 9 — ошибка (текст в errText). Интерфейс (сообщения, настройки) — в модуле NeonUI.
    Dim sr As ShapeRange, oldUnit As cdrUnit, lay As Layer, col As Collection
    Dim i As Long, t0 As Single, t1 As Single, rc As Long, grp As Boolean
    nIn = 0: nOut = 0: tAll = 0: tCore = 0: errText = ""
    If ActiveDocument Is Nothing Then
        {fn} = 1
        Exit Function
    End If
    Set sr = ActiveSelectionRange
    If sr.Count = 0 Then
        {fn} = 2
        Exit Function
    End If
    If tool <= 0 Then
        tool = 8
    End If
    CoreSetParams tool, P_TOL
    oldUnit = ActiveDocument.Unit
    ActiveDocument.Unit = cdrMillimeter
    ActiveDocument.BeginCommandGroup UStr("41D 435 43E 43D 3A 20 446 435 43D 442 440 430 43B 44C 43D 430 44F 20 43B 438 43D 438 44F")
    grp = True
    Optimization = True
    EventsEnabled = False
    t0 = Timer
    On Error GoTo Fail
    Set col = New Collection
    CollectShapes sr, col
    CoreInitRings
    For i = 1 To col.Count
        If ReadShape(col(i)) Then
            nIn = nIn + 1
        End If
    Next i
    If nIn = 0 Then
        rc = 3
        GoTo Fin
    End If
    DumpInput tool
    t1 = Timer
    CoreRun
    tCore = Timer - t1
    Set lay = GetOutLayer()
    On Error Resume Next
    lay.Visible = True: lay.Editable = True
    On Error GoTo Fail
    nOut = EmitShapes(lay)
    rc = 0
Fin:
    On Error Resume Next
    tAll = Timer - t0
    EventsEnabled = True
    Optimization = False
    If grp Then
        ActiveDocument.EndCommandGroup
    End If
    ActiveDocument.Unit = oldUnit
    ActiveWindow.Refresh
    Application.Refresh
    {fn} = rc
    Exit Function
Fail:
    errText = Err.Description
    rc = 9
    Resume Fin
End Function
'''

def conv(src, dst, vbname, oldmacro, fn):
    s = open(src, encoding='utf-8').read()
    s = re.sub(r'^Attribute VB_Name = ".*"', f'Attribute VB_Name = "{vbname}"', s, count=1, flags=re.M)
    s = re.sub(r'^Public (Sub|Function) (Core\w+)', r'Private \1 \2', s, flags=re.M)
    a = s.index(f'Public Sub {oldmacro}()')
    b = s.index('End Sub', a) + len('End Sub')
    s = s[:a] + ENTRY.format(fn=fn).rstrip('\n') + s[b:]
    pub = re.findall(r'^Public (?:Sub|Function) (\w+)', s, flags=re.M)
    assert pub == [fn], pub
    open(dst, 'w', encoding='utf-8').write(s)
    print(dst, 'public:', pub)

conv(f'{V}/NeonCenterline.bas', f'{S}/NeonMono.bas', 'NeonMono', 'NeonCenterline', 'NeonMonoRun')
conv(f'{V}/NeonScript.bas', f'{S}/NeonScript.bas', 'NeonScript', 'NeonCenterlineScript', 'NeonScriptRun')
