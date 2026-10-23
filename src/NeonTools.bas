Attribute VB_Name = "NeonTools"
' =====================================================================
'  NeonCenterline — инструменты:
'    MeasureLength   — длина выделенных контуров (для резки) и центральных линий (неон),
'                      результат можно скопировать в буфер обмена;
'    SelectNeonLines — выделить все центральные линии, созданные плагином, на странице.
' =====================================================================
Option Explicit

' длина выделенного: контуры отдельно от центральных линий плагина
Public Sub MeasureLength()
    Dim sr As ShapeRange, oldUnit As cdrUnit, msg As String
    Dim lenC As Double, nObjC As Long, nPathC As Long
    Dim lenN As Double, nObjN As Long, nPathN As Long, nmax As Double
    If ActiveDocument Is Nothing Then
        MsgBox NeonL("err_nodoc"), vbExclamation, NeonL("app_title")
        Exit Sub
    End If
    Set sr = ActiveSelectionRange
    If sr.Count = 0 Then
        MsgBox NeonL("err_nosel_any"), vbExclamation, NeonL("app_title")
        Exit Sub
    End If
    oldUnit = ActiveDocument.Unit
    ActiveDocument.Unit = cdrMillimeter
    On Error GoTo Fail
    MeasureRange sr, lenC, nObjC, nPathC, lenN, nObjN, nPathN, nmax
    ActiveDocument.Unit = oldUnit

    If nObjC > 0 Then
        msg = NeonLF("len_contours", nObjC, nPathC, FmtLen(lenC))
    End If
    If nObjN > 0 Then
        If Len(msg) > 0 Then
            msg = msg & vbCrLf & vbCrLf
        End If
        msg = msg & NeonLF("len_neon", nPathN, FmtLen(lenN), FmtLen(nmax))
    End If
    If Len(msg) = 0 Then
        MsgBox NeonL("len_none"), vbExclamation, NeonL("app_title")
        Exit Sub
    End If
    If MsgBox(msg & vbCrLf & vbCrLf & NeonL("len_copy"), vbYesNo + vbInformation, NeonL("len_title")) = vbYes Then
        CopyText msg
    End If
    Exit Sub
Fail:
    On Error Resume Next
    ActiveDocument.Unit = oldUnit
    MsgBox NeonLF("err_fail", Err.Description), vbCritical, NeonL("app_title")
End Sub

' все центральные линии плагина на текущей странице — в выделение
Public Sub SelectNeonLines()
    Dim col As Collection, i As Long, sr As ShapeRange
    If ActiveDocument Is Nothing Then
        MsgBox NeonL("err_nodoc"), vbExclamation, NeonL("app_title")
        Exit Sub
    End If
    Set col = New Collection
    FindNeon ActivePage.Shapes.All, col
    If col.Count = 0 Then
        MsgBox NeonL("sel_none"), vbInformation, NeonL("app_title")
        Exit Sub
    End If
    Set sr = CreateShapeRange
    For i = 1 To col.Count
        sr.Add col(i)
    Next i
    ActiveDocument.ClearSelection
    sr.AddToSelection
End Sub

' ---------------------------------------------------------------------
Private Sub MeasureRange(sr As ShapeRange, lenC As Double, nObjC As Long, nPathC As Long, _
                         lenN As Double, nObjN As Long, nPathN As Long, nmax As Double)
    Dim sh As Shape, crv As Curve, tmp As Shape, L As Double, sp As SubPath, isNeon As Boolean
    For Each sh In sr
        If sh.Type = cdrGroupShape Then
            MeasureRange sh.Shapes.All, lenC, nObjC, nPathC, lenN, nObjN, nPathN, nmax
        Else
            isNeon = (Left$(sh.Name, Len(NeonPrefix())) = NeonPrefix())
            Set tmp = Nothing
            Set crv = Nothing
            If sh.Type = cdrCurveShape Then
                Set crv = sh.Curve
            Else
                ' текст, прямоугольник, эллипс...: кривая отображения, иначе — временная копия в кривых
                On Error Resume Next
                Set crv = sh.DisplayCurve
                If crv Is Nothing Then
                    Set tmp = sh.Duplicate
                    tmp.ConvertToCurves
                    If tmp.Type = cdrCurveShape Then
                        Set crv = tmp.Curve
                    End If
                End If
                On Error GoTo 0
            End If
            If Not crv Is Nothing Then
                If isNeon Then
                    nObjN = nObjN + 1
                    For Each sp In crv.SubPaths
                        L = sp.Length
                        lenN = lenN + L
                        nPathN = nPathN + 1
                        If L > nmax Then
                            nmax = L
                        End If
                    Next sp
                Else
                    nObjC = nObjC + 1
                    For Each sp In crv.SubPaths
                        lenC = lenC + sp.Length
                        nPathC = nPathC + 1
                    Next sp
                End If
            End If
            If Not tmp Is Nothing Then
                tmp.Delete
            End If
        End If
    Next sh
End Sub

Private Sub FindNeon(sr As ShapeRange, col As Collection)
    Dim sh As Shape
    For Each sh In sr
        If sh.Type = cdrGroupShape Then
            FindNeon sh.Shapes.All, col
        ElseIf Left$(sh.Name, Len(NeonPrefix())) = NeonPrefix() Then
            col.Add sh
        End If
    Next sh
End Sub

' длина: мм и метры
Private Function FmtLen(ByVal mm As Double) As String
    FmtLen = Format$(mm, "0") & " " & NeonL("unit_mm") & " (" & Format$(mm / 1000, "0.00") & " " & NeonL("unit_m") & ")"
End Function

Private Sub CopyText(ByVal s As String)
    Dim d As Object
    On Error Resume Next
    Set d = CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
    d.SetText s
    d.PutInClipboard
End Sub

' «Неон: центр » — имя линий, которые создаёт плагин (как в NeonMono / NeonScript)
Private Function NeonPrefix() As String
    NeonPrefix = ChrW(&H41D) & ChrW(&H435) & ChrW(&H43E) & ChrW(&H43D) & ":" & " " & _
                 ChrW(&H446) & ChrW(&H435) & ChrW(&H43D) & ChrW(&H442) & ChrW(&H440) & " "
End Function
