Attribute VB_Name = "NeonTools"
' =====================================================================
'  NeonCenterline — инструменты:
'    MeasureLength   — длина выделенных контуров (для резки) и центральных линий (неон),
'                      результат можно скопировать в буфер обмена;
'    SelectNeonLines — выделить все центральные линии, созданные плагином, на странице;
'    NeonLineIds / NeonLinesNew — какие линии появились после расчёта (для подгонки под рез);
'    SimplifyContours — упростить выделенные кривые (меньше узлов в пределах допуска);
'    PlaceGrommets   — точки под люверсы по краю выделенных баннеров (подложек).
' =====================================================================
Option Explicit

' ---- упрощение контура: данные текущего подпути (см. раздел «Упрощение контура» ниже) ----
'@@SIMP_DECL_BEGIN
Private Const SIMP_REPARAM As Long = 4

' исходный подпуть (4 точки на сегмент)
Private SX() As Double, SY() As Double, NS As Long
' рабочие сегменты: без вырожденных, замкнутый — с замыкающим отрезком
Private WX() As Double, WY() As Double, NW As Long
' точки развёртки, индексы исходных узлов в них
Private PX() As Double, PY() As Double, NP As Long, NodeI() As Long
' точки текущего участка
Private RX() As Double, RY() As Double, NR As Long
' результат (4 точки на сегмент, признак прямой)
Private OX() As Double, OY() As Double, OL() As Boolean, NO As Long
Private CosLocal As Double, CosChord As Double
'@@SIMP_DECL_END

' ---- люверсы: ломаная контура (флаг угла), центры люверсов (см. раздел «Люверсы») ----
'@@GROM_DECL_BEGIN
Private GPX() As Double, GPY() As Double, GPC() As Boolean, GN As Long
Private GRX() As Double, GRY() As Double, GRN As Long
'@@GROM_DECL_END

' длина выделенного: контуры отдельно от центральных линий плагина
Public Sub MeasureLength()
    Dim sr As ShapeRange, oldUnit As cdrUnit, msg As String
    Dim lenC As Double, nObjC As Long, nPathC As Long
    Dim lenN As Double, nObjN As Long, nPathN As Long, nmax As Double, neonCol As Collection, cs As String
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
        Set neonCol = New Collection
        FindNeon sr, neonCol
        cs = NeonCutSummary(neonCol)
        If Len(cs) > 0 Then
            msg = msg & vbCrLf & cs
        End If
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

' StaticID всех центральных линий на странице (ключи коллекции)
Public Function NeonLineIds() As Collection
    Dim col As Collection, ids As Collection, i As Long
    Set ids = New Collection
    Set col = New Collection
    If Not ActivePage Is Nothing Then
        FindNeon ActivePage.Shapes.All, col
    End If
    For i = 1 To col.Count
        On Error Resume Next
        ids.Add CStr(col(i).StaticID), CStr(col(i).StaticID)
        On Error GoTo 0
    Next i
    Set NeonLineIds = ids
End Function

' линии, которых не было в before
Public Function NeonLinesNew(before As Collection) As Collection
    Dim col As Collection, res As Collection, i As Long, v As Variant, found As Boolean
    Set res = New Collection
    Set col = New Collection
    FindNeon ActivePage.Shapes.All, col
    For i = 1 To col.Count
        found = False
        On Error Resume Next
        v = before(CStr(col(i).StaticID))
        found = (Err.Number = 0)
        Err.Clear
        On Error GoTo 0
        If Not found Then
            res.Add col(i)
        End If
    Next i
    Set NeonLinesNew = res
End Function

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

' =====================================================================
'  NeonCenterline — упрощение контура (кнопка «Неон: упростить контур»).
'
'  Выделенные кривые меняются на месте (одна операция, откат — Ctrl+Z): меньше узлов, кривая
'  не отходит от исходной больше чем на допуск (окно перед запуском и Настройки, по умолчанию
'  0.1 мм), острые углы
'  остаются углами, гладкие места — гладкими. Подпуть, у которого узлов меньше не стало
'  (уже экономная кривая), не трогается. Текст и фигуры (не кривые) пропускаются.
'
'  Алгоритм — прототип proto/simplify.py (этот модуль повторяет его один в один):
'    1. Угол — узел, где касательные по контрольным точкам расходятся > 10 град И хорды на
'       масштабе d = max(4 допуска, 0.5 мм) поворачивают > 25 град (зазубрина трассировки —
'       не угол, тугое скругление без излома — не угол).
'    2. Участки между углами разворачиваются в точки (шаг 4 допуска, 0.2..1 мм).
'    3. Участок — кубики Безье по Шнайдеру (Graphics Gems, 1990): касательные на концах — по
'       исходным сегментам, МНК для длин ручек, до 4 шагов Ньютона по параметру; не ложится —
'       деление в точке наибольшей ошибки с общей касательной. Почти прямой участок — отрезок.
'       Участок, где узлов не стало меньше, остаётся из исходных сегментов.
' =====================================================================
'@@SIMP_BEGIN
' допуск упрощения, мм (Настройки)
Public Function NeonSimplifyTol() As Double
    Dim t As Double
    t = ParseNum(GetSetting("NeonCenterline", "Simplify", "Tol", "0.1"))
    If t <= 0 Then
        t = 0.1
    End If
    NeonSimplifyTol = t
End Function

' ---------------------------------------------------------------------
'  Кнопка
' ---------------------------------------------------------------------
Public Sub SimplifyContours()
    Dim col As Collection, tol As Double, oldUnit As cdrUnit, i As Long, nSkip As Long
    Dim n0 As Long, n1 As Long, nObj As Long, grp As Boolean, errText As String, msg As String
    If ActiveDocument Is Nothing Then
        MsgBox NeonL("err_nodoc"), vbExclamation, NeonL("app_title")
        Exit Sub
    End If
    If ActiveSelectionRange.Count = 0 Then
        MsgBox NeonL("err_nosel_any"), vbExclamation, NeonL("app_title")
        Exit Sub
    End If
    ' окно с допуском перед запуском (как у центральных линий; отключается галочкой в окне)
    If Not AskSimplifyParams() Then
        Exit Sub
    End If
    tol = NeonSimplifyTol()
    Set col = New Collection
    CollectSimp ActiveSelectionRange, col, nSkip
    If col.Count = 0 Then
        MsgBox NeonL("simp_none"), vbExclamation, NeonL("simp_title")
        Exit Sub
    End If
    oldUnit = ActiveDocument.Unit
    ActiveDocument.Unit = cdrMillimeter
    ActiveDocument.BeginCommandGroup NeonL("simp_title")
    grp = True
    Optimization = True
    EventsEnabled = False
    On Error GoTo Fail
    For i = 1 To col.Count
        SimplifyShape col(i), tol, n0, n1, nObj
    Next i
Fin:
    On Error Resume Next
    EventsEnabled = True
    Optimization = False
    If grp Then
        ActiveDocument.EndCommandGroup
    End If
    ActiveDocument.Unit = oldUnit
    ActiveWindow.Refresh
    Application.Refresh
    On Error GoTo 0
    If Len(errText) > 0 Then
        MsgBox NeonLF("err_fail", errText), vbCritical, NeonL("simp_title")
        Exit Sub
    End If
    If n0 > 0 Then
        msg = NeonLF("simp_report", FmtNum(tol), nObj, col.Count, n0, n1, Format$(100# * (n0 - n1) / n0, "0"))
    Else
        msg = NeonLF("simp_report", FmtNum(tol), nObj, col.Count, n0, n1, "0")
    End If
    If nSkip > 0 Then
        msg = msg & vbCrLf & vbCrLf & NeonLF("simp_skipped", nSkip)
    End If
    MsgBox msg, vbInformation, NeonL("simp_title")
    NeonToolbarCheck
    Exit Sub
Fail:
    errText = Err.Description
    Resume Fin
End Sub

Private Sub CollectSimp(sr As ShapeRange, col As Collection, nSkip As Long)
    Dim sh As Shape
    For Each sh In sr
        If sh.Type = cdrGroupShape Then
            CollectSimp sh.Shapes.All, col, nSkip
        ElseIf sh.Type = cdrCurveShape Then
            col.Add sh
        Else
            nSkip = nSkip + 1
        End If
    Next sh
End Sub

' упростить кривую фигуры: подпути по очереди, кривая пересобирается, если стало меньше узлов
Private Sub SimplifyShape(shp As Shape, ByVal tol As Double, n0 As Long, n1 As Long, nObj As Long)
    Dim sp As SubPath, sg As Segment, crv As Curve, nsp As Long, i As Long, s As Long, changed As Boolean
    Dim x0 As Double, y0 As Double, x1 As Double, y1 As Double, x2 As Double, y2 As Double, x3 As Double, y3 As Double
    Dim bx() As Double, by() As Double, bl() As Boolean, nb As Long
    Dim pS0() As Long, pN() As Long, pC() As Boolean, cl As Boolean
    ReDim bx(0 To 255): ReDim by(0 To 255): ReDim bl(0 To 63): nb = 0
    ReDim pS0(0 To 15): ReDim pN(0 To 15): ReDim pC(0 To 15): nsp = 0
    For Each sp In shp.Curve.SubPaths
        NS = 0
        ReDim SX(0 To 63): ReDim SY(0 To 63)
        For Each sg In sp.Segments
            x0 = sg.StartNode.PositionX: y0 = sg.StartNode.PositionY
            x3 = sg.EndNode.PositionX: y3 = sg.EndNode.PositionY
            If sg.Type = cdrCurveSegment Then
                sg.GetStartingControlPointPosition x1, y1
                sg.GetEndingControlPointPosition x2, y2
            Else
                x1 = x0 + (x3 - x0) / 3: y1 = y0 + (y3 - y0) / 3
                x2 = x0 + 2 * (x3 - x0) / 3: y2 = y0 + 2 * (y3 - y0) / 3
            End If
            PutS NS, x0, y0, x1, y1, x2, y2, x3, y3
            NS = NS + 1
        Next sg
        If NS > 0 Then
            cl = sp.Closed
            SimpSubpath cl, tol
            If NO < NS Then
                changed = True
            End If
            If cl Then
                n0 = n0 + NS: n1 = n1 + NO
            Else
                n0 = n0 + NS + 1: n1 = n1 + NO + 1
            End If
            If nsp > UBound(pS0) Then
                ReDim Preserve pS0(0 To 2 * nsp): ReDim Preserve pN(0 To 2 * nsp): ReDim Preserve pC(0 To 2 * nsp)
            End If
            pS0(nsp) = nb: pN(nsp) = NO: pC(nsp) = cl: nsp = nsp + 1
            For s = 0 To NO - 1
                If 4 * nb + 3 > UBound(bx) Then
                    ReDim Preserve bx(0 To 8 * nb + 16): ReDim Preserve by(0 To 8 * nb + 16)
                End If
                If nb > UBound(bl) Then
                    ReDim Preserve bl(0 To 2 * nb + 4)
                End If
                For i = 0 To 3
                    bx(4 * nb + i) = OX(4 * s + i): by(4 * nb + i) = OY(4 * s + i)
                Next i
                bl(nb) = OL(s)
                nb = nb + 1
            Next s
        End If
    Next sp
    If Not changed Then
        Exit Sub
    End If
    Set crv = Application.CreateCurve(ActiveDocument)
    For i = 0 To nsp - 1
        If pN(i) > 0 Then
            Set sp = crv.CreateSubPath(bx(4 * pS0(i)), by(4 * pS0(i)))
            For s = pS0(i) To pS0(i) + pN(i) - 1
                If bl(s) Then
                    sp.AppendLineSegment bx(4 * s + 3), by(4 * s + 3)
                Else
                    sp.AppendCurveSegment2 bx(4 * s + 3), by(4 * s + 3), bx(4 * s + 1), by(4 * s + 1), bx(4 * s + 2), by(4 * s + 2)
                End If
            Next s
            If pC(i) Then
                On Error Resume Next
                sp.Closed = True
                On Error GoTo 0
            End If
        End If
    Next i
    shp.Curve.CopyAssign crv
    nObj = nObj + 1
End Sub

Private Sub PutS(ByVal s As Long, ByVal x0 As Double, ByVal y0 As Double, ByVal x1 As Double, ByVal y1 As Double, _
                 ByVal x2 As Double, ByVal y2 As Double, ByVal x3 As Double, ByVal y3 As Double)
    If 4 * s + 3 > UBound(SX) Then
        ReDim Preserve SX(0 To 8 * s + 16): ReDim Preserve SY(0 To 8 * s + 16)
    End If
    SX(4 * s) = x0: SY(4 * s) = y0: SX(4 * s + 1) = x1: SY(4 * s + 1) = y1
    SX(4 * s + 2) = x2: SY(4 * s + 2) = y2: SX(4 * s + 3) = x3: SY(4 * s + 3) = y3
End Sub

' ---------------------------------------------------------------------
'  Упрощение одного подпути: вход SX/SY/NS, выход OX/OY/OL/NO.
'  Если узлов меньше не стало — на выходе исходные сегменты.
' ---------------------------------------------------------------------
Public Sub SimpSubpath(ByVal closed As Boolean, ByVal tol As Double)
    Dim s As Long, k As Long
    CosLocal = Cos(10# * Atn(1#) / 45#)
    CosChord = Cos(25# * Atn(1#) / 45#)
    NO = 0
    ReDim OX(0 To 4 * NS + 16): ReDim OY(0 To 4 * NS + 16): ReDim OL(0 To NS + 4)
    SimpCore closed, tol
    If NO >= NS Then
        ReDim OX(0 To 4 * NS + 16): ReDim OY(0 To 4 * NS + 16): ReDim OL(0 To NS + 4)
        For s = 0 To NS - 1
            For k = 0 To 3
                OX(4 * s + k) = SX(4 * s + k): OY(4 * s + k) = SY(4 * s + k)
            Next k
            OL(s) = IsLineArr(SX, SY, s)
        Next s
        NO = NS
    End If
End Sub

Private Sub SimpCore(ByVal closed As Boolean, ByVal tol As Double)
    Dim s As Long, k As Long, stp As Double, nc As Long, cs() As Long, i As Long
    Dim ax As Double, ay As Double, bx As Double, by As Double, tinx As Double, tiny As Double, toutx As Double, touty As Double
    Dim tx As Double, ty As Double, L As Double, a As Long, b As Long, ia As Long, ib As Long, m As Long, bnd() As Long
    Dim rc() As Double
    ' рабочие сегменты: без вырожденных (если все вырождены — как есть)
    ReDim WX(0 To 4 * NS + 8): ReDim WY(0 To 4 * NS + 8): NW = 0
    For s = 0 To NS - 1
        If SegLenArr(SX, SY, s) > 0.000000001 Then
            CopySeg s
        End If
    Next s
    If NW = 0 Then
        For s = 0 To NS - 1
            CopySeg s
        Next s
    End If
    If closed Then
        ax = WX(4 * (NW - 1) + 3): ay = WY(4 * (NW - 1) + 3): bx = WX(0): by = WY(0)
        If Sqr((ax - bx) ^ 2 + (ay - by) ^ 2) > 0.000001 Then
            ' замыкающий отрезок
            WX(4 * NW) = ax: WY(4 * NW) = ay
            WX(4 * NW + 1) = ax + (bx - ax) / 3: WY(4 * NW + 1) = ay + (by - ay) / 3
            WX(4 * NW + 2) = ax + 2 * (bx - ax) / 3: WY(4 * NW + 2) = ay + 2 * (by - ay) / 3
            WX(4 * NW + 3) = bx: WY(4 * NW + 3) = by
            NW = NW + 1
        End If
    End If
    stp = 4 * tol
    If stp < 0.2 Then
        stp = 0.2
    End If
    If stp > 1 Then
        stp = 1
    End If
    Flatten stp
    If NP < 3 Then
        ' вырожденный подпуть: рабочие сегменты как есть
        For s = 0 To NW - 1
            For k = 0 To 3
                OX(4 * s + k) = WX(4 * s + k): OY(4 * s + k) = WY(4 * s + k)
            Next k
            OL(s) = IsLineArr(WX, WY, s)
        Next s
        NO = NW
        Exit Sub
    End If
    FindCorners closed, tol, cs, nc
    tol = 0.9 * tol            ' запас: ошибка фита считается по точкам
    If closed Then
        If nc = 0 Then
            ' без углов: от узла 0 по кругу, касательная — по хорде между точками в 2 шагах
            ' по обе стороны (устойчиво к микросегментам у начала)
            m = NP
            If Sqr((PX(0) - PX(NP - 1)) ^ 2 + (PY(0) - PY(NP - 1)) ^ 2) > 0.000000001 Then
                m = NP + 1
            End If
            ReDim RX(0 To m - 1): ReDim RY(0 To m - 1): NR = m
            For i = 0 To NP - 1
                RX(i) = PX(i): RY(i) = PY(i)
            Next i
            If m > NP Then
                RX(NP) = PX(0): RY(NP) = PY(0)
            End If
            ReDim rc(0 To NR - 1)
            rc(0) = 0
            For i = 1 To NR - 1
                rc(i) = rc(i - 1) + Sqr((RX(i) - RX(i - 1)) ^ 2 + (RY(i) - RY(i - 1)) ^ 2)
            Next i
            WalkPt RX, RY, rc, NR, 0, 2 * stp, -1, True, ax, ay
            WalkPt RX, RY, rc, NR, 0, 2 * stp, 1, True, bx, by
            NormV bx - ax, by - ay, tx, ty
            If tx = 0 And ty = 0 Then
                TanOut 0, tx, ty
            End If
            FitRun 0, NR - 1, tx, ty, -tx, -ty, tol
            Exit Sub
        End If
        For k = 0 To nc - 1
            a = cs(k)
            If k < nc - 1 Then
                b = cs(k + 1)
            Else
                b = cs(0) + NW
            End If
            ia = NodeI(a)
            If b < NW Then
                ib = NodeI(b)
            Else
                ib = NodeI(b - NW) + (NP - 1)
            End If
            NR = ib - ia + 1
            ReDim RX(0 To NR - 1): ReDim RY(0 To NR - 1)
            For i = ia To ib
                RX(i - ia) = PX(i Mod (NP - 1)): RY(i - ia) = PY(i Mod (NP - 1))
            Next i
            RunTans a Mod NW, (b - 1) Mod NW, stp, toutx, touty, tinx, tiny
            FitOrKeep toutx, touty, tinx, tiny, tol, a, b
        Next k
        Exit Sub
    End If
    ' открытый: границы — концы и углы
    ReDim bnd(0 To nc + 1)
    bnd(0) = 0
    For k = 0 To nc - 1
        bnd(k + 1) = cs(k)
    Next k
    bnd(nc + 1) = NW
    For k = 0 To nc
        a = bnd(k): b = bnd(k + 1)
        ia = NodeI(a): ib = NodeI(b)
        NR = ib - ia + 1
        ReDim RX(0 To NR - 1): ReDim RY(0 To NR - 1)
        For i = ia To ib
            RX(i - ia) = PX(i): RY(i - ia) = PY(i)
        Next i
        RunTans a, b - 1, stp, toutx, touty, tinx, tiny
        FitOrKeep toutx, touty, tinx, tiny, tol, a, b
    Next k
End Sub

' фит участка RX/RY между углами (рабочие сегменты a..b-1, по кругу); если узлов не стало
' меньше — исходные сегменты участка (экономная кривая шрифта не портится)
Private Sub FitOrKeep(ByVal t1x As Double, ByVal t1y As Double, ByVal t2x As Double, ByVal t2y As Double, _
                      ByVal tol As Double, ByVal a As Long, ByVal b As Long)
    Dim no0 As Long, s As Long, k As Long, bz(0 To 7) As Double
    no0 = NO
    FitRun 0, NR - 1, t1x, t1y, t2x, t2y, tol
    If NO - no0 >= b - a Then
        NO = no0
        For s = a To b - 1
            For k = 0 To 3
                bz(2 * k) = WX(4 * (s Mod NW) + k): bz(2 * k + 1) = WY(4 * (s Mod NW) + k)
            Next k
            Emit bz, IsLineArr(WX, WY, s Mod NW)
        Next s
    End If
End Sub

' касательные концов участка RX/RY (t1 — из начала внутрь, t2 — из конца внутрь): по контрольным
' точкам крайних сегментов sa, sb; сегмент короче 2 шагов (микросегмент, «шип») — по хорде
Private Sub RunTans(ByVal sa As Long, ByVal sb As Long, ByVal stp As Double, t1x As Double, t1y As Double, _
                    t2x As Double, t2y As Double)
    Dim tx As Double, ty As Double
    If SegLenArr(WX, WY, sa) >= 2 * stp Then
        TanOut sa, t1x, t1y
    Else
        EndTan True, 2 * stp, t1x, t1y
    End If
    If SegLenArr(WX, WY, sb) >= 2 * stp Then
        TanIn sb, tx, ty
        t2x = -tx: t2y = -ty
    Else
        EndTan False, 2 * stp, t2x, t2y
    End If
End Sub

' касательная конца участка RX/RY по хорде до точки на расстоянии dist (внутрь участка)
Private Sub EndTan(ByVal atStart As Boolean, ByVal dist As Double, tx As Double, ty As Double)
    Dim k As Long, L As Double
    If atStart Then
        k = 1
        Do While k < NR - 1
            L = L + Sqr((RX(k) - RX(k - 1)) ^ 2 + (RY(k) - RY(k - 1)) ^ 2)
            If L >= dist Then
                Exit Do
            End If
            k = k + 1
        Loop
        If k > NR - 1 Then
            k = NR - 1
        End If
        NormV RX(k) - RX(0), RY(k) - RY(0), tx, ty
    Else
        k = NR - 2
        Do While k > 0
            L = L + Sqr((RX(k + 1) - RX(k)) ^ 2 + (RY(k + 1) - RY(k)) ^ 2)
            If L >= dist Then
                Exit Do
            End If
            k = k - 1
        Loop
        If k < 0 Then
            k = 0
        End If
        NormV RX(k) - RX(NR - 1), RY(k) - RY(NR - 1), tx, ty
    End If
End Sub

Private Sub CopySeg(ByVal s As Long)
    Dim k As Long
    For k = 0 To 3
        WX(4 * NW + k) = SX(4 * s + k): WY(4 * NW + k) = SY(4 * s + k)
    Next k
    NW = NW + 1
End Sub

' ---------------------------------------------------------------------
'  Сегменты
' ---------------------------------------------------------------------
Private Sub BzArr(ax() As Double, ay() As Double, ByVal s As Long, ByVal t As Double, x As Double, y As Double)
    Dim mt As Double, a As Double, b As Double, c As Double, d As Double
    mt = 1 - t
    a = mt * mt * mt: b = 3 * mt * mt * t: c = 3 * mt * t * t: d = t * t * t
    x = a * ax(4 * s) + b * ax(4 * s + 1) + c * ax(4 * s + 2) + d * ax(4 * s + 3)
    y = a * ay(4 * s) + b * ay(4 * s + 1) + c * ay(4 * s + 2) + d * ay(4 * s + 3)
End Sub

Private Function SegLenArr(ax() As Double, ay() As Double, ByVal s As Long) As Double
    Dim k As Long, px0 As Double, py0 As Double, x As Double, y As Double, L As Double
    px0 = ax(4 * s): py0 = ay(4 * s)
    For k = 1 To 16
        BzArr ax, ay, s, k / 16, x, y
        L = L + Sqr((x - px0) ^ 2 + (y - py0) ^ 2)
        px0 = x: py0 = y
    Next k
    SegLenArr = L
End Function

Private Function IsLineArr(ax() As Double, ay() As Double, ByVal s As Long) As Boolean
    Dim dx As Double, dy As Double, L As Double, k As Long
    dx = ax(4 * s + 3) - ax(4 * s): dy = ay(4 * s + 3) - ay(4 * s)
    L = Sqr(dx * dx + dy * dy)
    If L < 0.000000001 Then
        IsLineArr = (Sqr((ax(4 * s + 1) - ax(4 * s)) ^ 2 + (ay(4 * s + 1) - ay(4 * s)) ^ 2) < 0.000000001) And _
                    (Sqr((ax(4 * s + 2) - ax(4 * s)) ^ 2 + (ay(4 * s + 2) - ay(4 * s)) ^ 2) < 0.000000001)
        Exit Function
    End If
    For k = 1 To 2
        If Abs((ax(4 * s + k) - ax(4 * s)) * dy - (ay(4 * s + k) - ay(4 * s)) * dx) / L > 0.001 Then
            Exit Function
        End If
    Next k
    IsLineArr = True
End Function

Private Sub NormV(ByVal dx As Double, ByVal dy As Double, nx As Double, ny As Double)
    Dim L As Double
    L = Sqr(dx * dx + dy * dy)
    If L > 0.000000000001 Then
        nx = dx / L: ny = dy / L
    Else
        nx = 0: ny = 0
    End If
End Sub

' касательная в начале рабочего сегмента (направление движения)
Private Sub TanOut(ByVal s As Long, tx As Double, ty As Double)
    Dim k As Long, dx As Double, dy As Double
    For k = 1 To 3
        dx = WX(4 * s + k) - WX(4 * s): dy = WY(4 * s + k) - WY(4 * s)
        If Sqr(dx * dx + dy * dy) > 0.000000001 Then
            NormV dx, dy, tx, ty
            Exit Sub
        End If
    Next k
    tx = 0: ty = 0
End Sub

' касательная в конце рабочего сегмента (направление движения)
Private Sub TanIn(ByVal s As Long, tx As Double, ty As Double)
    Dim k As Long, dx As Double, dy As Double
    For k = 2 To 0 Step -1
        dx = WX(4 * s + 3) - WX(4 * s + k): dy = WY(4 * s + 3) - WY(4 * s + k)
        If Sqr(dx * dx + dy * dy) > 0.000000001 Then
            NormV dx, dy, tx, ty
            Exit Sub
        End If
    Next k
    tx = 0: ty = 0
End Sub

' развёртка рабочих сегментов в точки; NodeI(i) — точка узла i (начало сегмента i; NodeI(NW) — конец)
Private Sub Flatten(ByVal stp As Double)
    Dim s As Long, m As Long, k As Long, L As Double, x As Double, y As Double
    ReDim PX(0 To 255): ReDim PY(0 To 255): ReDim NodeI(0 To NW)
    PX(0) = WX(0): PY(0) = WY(0): NP = 1
    NodeI(0) = 0
    For s = 0 To NW - 1
        L = SegLenArr(WX, WY, s)
        m = -Int(-L / stp)
        If m < 1 Then
            m = 1
        End If
        If Not IsLineArr(WX, WY, s) Then
            If m < 4 Then
                m = 4
            End If
        End If
        For k = 1 To m
            BzArr WX, WY, s, k / m, x, y
            If Sqr((x - PX(NP - 1)) ^ 2 + (y - PY(NP - 1)) ^ 2) > 0.000000001 Then
                If NP > UBound(PX) Then
                    ReDim Preserve PX(0 To 2 * NP): ReDim Preserve PY(0 To 2 * NP)
                End If
                PX(NP) = x: PY(NP) = y: NP = NP + 1
            End If
        Next k
        NodeI(s + 1) = NP - 1
    Next s
End Sub

' точка на расстоянии d по ломаной (qx, qy; накопленные длины qc) от точки i в сторону dirn
Private Sub WalkPt(qx() As Double, qy() As Double, qc() As Double, ByVal n As Long, ByVal i As Long, _
                   ByVal d As Double, ByVal dirn As Long, ByVal closed As Boolean, ox As Double, oy As Double)
    Dim total As Double, target As Double, lo As Long, hi As Long, md As Long, sg As Double, t As Double
    total = qc(n - 1)
    If total <= 0 Then
        ox = qx(i): oy = qy(i)
        Exit Sub
    End If
    target = qc(i) + dirn * d
    If closed Then
        target = target - total * Int(target / total)
    Else
        If target <= 0 Then
            ox = qx(0): oy = qy(0)
            Exit Sub
        End If
        If target >= total Then
            ox = qx(n - 1): oy = qy(n - 1)
            Exit Sub
        End If
    End If
    lo = 0: hi = n - 1
    Do While hi - lo > 1
        md = (lo + hi) \ 2
        If qc(md) <= target Then
            lo = md
        Else
            hi = md
        End If
    Loop
    sg = qc(hi) - qc(lo)
    If sg > 0 Then
        t = (target - qc(lo)) / sg
    Else
        t = 0
    End If
    ox = qx(lo) + (qx(hi) - qx(lo)) * t: oy = qy(lo) + (qy(hi) - qy(lo)) * t
End Sub

' узлы-углы (по возрастанию): у замкнутого 0..NW-1, у открытого 1..NW-1
Private Sub FindCorners(ByVal closed As Boolean, ByVal tol As Double, cs() As Long, nc As Long)
    Dim cum() As Double, i As Long, n As Long, total As Double, d As Double
    Dim qx() As Double, qy() As Double, qc() As Double, nq As Long, i0 As Long, k As Long
    Dim tinx As Double, tiny As Double, toutx As Double, touty As Double
    Dim ax As Double, ay As Double, bx As Double, by As Double, ux As Double, uy As Double, vx As Double, vy As Double
    ReDim cs(0 To NW): nc = 0
    ReDim cum(0 To NP - 1)
    cum(0) = 0
    For i = 1 To NP - 1
        cum(i) = cum(i - 1) + Sqr((PX(i) - PX(i - 1)) ^ 2 + (PY(i) - PY(i - 1)) ^ 2)
    Next i
    total = cum(NP - 1)
    ' рабочая ломаная: у замкнутого с повтором начала — без повтора, затем + начало (как в прототипе)
    If closed And Sqr((PX(0) - PX(NP - 1)) ^ 2 + (PY(0) - PY(NP - 1)) ^ 2) < 0.000000001 Then
        n = NP - 1
    Else
        n = NP
    End If
    If closed Then
        nq = n + 1
    Else
        nq = n
    End If
    ReDim qx(0 To nq - 1): ReDim qy(0 To nq - 1): ReDim qc(0 To nq - 1)
    For i = 0 To n - 1
        qx(i) = PX(i): qy(i) = PY(i): qc(i) = cum(i)
    Next i
    If closed Then
        qx(n) = PX(0): qy(n) = PY(0): qc(n) = total
    End If
    d = 4 * tol
    If d < 0.5 Then
        d = 0.5
    End If
    If closed Then
        i0 = 0
    Else
        i0 = 1
    End If
    For i = i0 To NW - 1
        If i = 0 Then
            TanIn NW - 1, tinx, tiny
        Else
            TanIn i - 1, tinx, tiny
        End If
        TanOut i, toutx, touty
        If tinx * toutx + tiny * touty < CosLocal Then
            k = NodeI(i)
            If k >= n Then
                k = 0
            End If
            WalkPt qx, qy, qc, nq, k, d, -1, closed, ax, ay
            WalkPt qx, qy, qc, nq, k, d, 1, closed, bx, by
            NormV qx(k) - ax, qy(k) - ay, ux, uy
            NormV bx - qx(k), by - qy(k), vx, vy
            If (ux = 0 And uy = 0) Or (vx = 0 And vy = 0) Then
                cs(nc) = i: nc = nc + 1
            ElseIf ux * vx + uy * vy < CosChord Then
                cs(nc) = i: nc = nc + 1
            End If
        End If
    Next i
End Sub

' ---------------------------------------------------------------------
'  Фит участка RX/RY(a..b) кубиками (Шнайдер). t1 — касательная из начала внутрь,
'  t2 — из конца внутрь.
' ---------------------------------------------------------------------
Private Sub FitRun(ByVal a As Long, ByVal b As Long, ByVal t1x As Double, ByVal t1y As Double, _
                   ByVal t2x As Double, ByVal t2y As Double, ByVal tol As Double)
    Dim u() As Double, bz(0 To 7) As Double, er As Double, at As Long, it As Long, i As Long
    Dim tcx As Double, tcy As Double, L As Double, n As Long, k As Long
    n = b - a + 1
    If n < 2 Then
        Exit Sub
    End If
    If n = 2 Or LineOK(a, b, tol) Then
        bz(0) = RX(a): bz(1) = RY(a)
        bz(2) = RX(a) + (RX(b) - RX(a)) / 3: bz(3) = RY(a) + (RY(b) - RY(a)) / 3
        bz(4) = RX(a) + 2 * (RX(b) - RX(a)) / 3: bz(5) = RY(a) + 2 * (RY(b) - RY(a)) / 3
        bz(6) = RX(b): bz(7) = RY(b)
        Emit bz, True
        Exit Sub
    End If
    ReDim u(a To b)
    u(a) = 0
    For i = a + 1 To b
        u(i) = u(i - 1) + Sqr((RX(i) - RX(i - 1)) ^ 2 + (RY(i) - RY(i - 1)) ^ 2)
    Next i
    L = u(b)
    For i = a To b
        If L > 0 Then
            u(i) = u(i) / L
        Else
            u(i) = (i - a) / (n - 1)
        End If
    Next i
    GenBezier a, b, u, t1x, t1y, t2x, t2y, bz
    er = MaxErr(a, b, u, bz, at)
    If er <= tol Then
        Emit bz, False
        Exit Sub
    End If
    If er <= 4 * tol Then
        For it = 1 To SIMP_REPARAM
            For i = a + 1 To b - 1
                u(i) = NewtonT(bz, RX(i), RY(i), u(i))
            Next i
            GenBezier a, b, u, t1x, t1y, t2x, t2y, bz
            er = MaxErr(a, b, u, bz, at)
            If er <= tol Then
                Emit bz, False
                Exit Sub
            End If
        Next it
    End If
    ' деление в точке наибольшей ошибки (at — индекс в участке), общая касательная
    If at < 1 Then
        at = 1
    End If
    If at > n - 2 Then
        at = n - 2
    End If
    k = at
    If k > 3 Then
        k = 3
    End If
    If n - 1 - at < k Then
        k = n - 1 - at
    End If
    If k < 1 Then
        k = 1
    End If
    NormV RX(a + at - k) - RX(a + at + k), RY(a + at - k) - RY(a + at + k), tcx, tcy
    FitRun a, a + at, t1x, t1y, tcx, tcy, tol
    FitRun a + at, b, -tcx, -tcy, t2x, t2y, tol
End Sub

Private Function LineOK(ByVal a As Long, ByVal b As Long, ByVal tol As Double) As Boolean
    Dim L As Double, dx As Double, dy As Double, i As Long, t As Double
    L = Sqr((RX(b) - RX(a)) ^ 2 + (RY(b) - RY(a)) ^ 2)
    If L < 0.000000001 Then
        Exit Function
    End If
    dx = (RX(b) - RX(a)) / L: dy = (RY(b) - RY(a)) / L
    For i = a + 1 To b - 1
        t = (RX(i) - RX(a)) * dx + (RY(i) - RY(a)) * dy
        If t < -0.5 * tol Or t > L + 0.5 * tol Then
            Exit Function
        End If
        If Abs(-(RX(i) - RX(a)) * dy + (RY(i) - RY(a)) * dx) > 0.5 * tol Then
            Exit Function
        End If
    Next i
    LineOK = True
End Function

' МНК для длин ручек при заданных касательных (Шнайдер); вырожденный случай — треть хорды
Private Sub GenBezier(ByVal a As Long, ByVal b As Long, u() As Double, ByVal t1x As Double, ByVal t1y As Double, _
                      ByVal t2x As Double, ByVal t2y As Double, bz() As Double)
    Dim c00 As Double, c01 As Double, c11 As Double, x0 As Double, x1 As Double, i As Long
    Dim t As Double, mt As Double, b0 As Double, b1 As Double, b2 As Double, b3 As Double
    Dim a1x As Double, a1y As Double, a2x As Double, a2y As Double, tx As Double, ty As Double
    Dim det As Double, sg As Double, al As Double, ar As Double, eps As Double
    For i = a To b
        t = u(i): mt = 1 - t
        b0 = mt * mt * mt: b1 = 3 * mt * mt * t: b2 = 3 * mt * t * t: b3 = t * t * t
        a1x = t1x * b1: a1y = t1y * b1: a2x = t2x * b2: a2y = t2y * b2
        c00 = c00 + a1x * a1x + a1y * a1y
        c01 = c01 + a1x * a2x + a1y * a2y
        c11 = c11 + a2x * a2x + a2y * a2y
        tx = RX(i) - (RX(a) * (b0 + b1) + RX(b) * (b2 + b3))
        ty = RY(i) - (RY(a) * (b0 + b1) + RY(b) * (b2 + b3))
        x0 = x0 + a1x * tx + a1y * ty
        x1 = x1 + a2x * tx + a2y * ty
    Next i
    det = c00 * c11 - c01 * c01
    sg = Sqr((RX(b) - RX(a)) ^ 2 + (RY(b) - RY(a)) ^ 2)
    al = 0: ar = 0
    If Abs(det) > 0.000000000001 Then
        al = (x0 * c11 - x1 * c01) / det
        ar = (c00 * x1 - c01 * x0) / det
    End If
    eps = 0.000001 * sg
    If al < eps Or ar < eps Then
        al = sg / 3: ar = sg / 3
    End If
    bz(0) = RX(a): bz(1) = RY(a)
    bz(2) = RX(a) + t1x * al: bz(3) = RY(a) + t1y * al
    bz(4) = RX(b) + t2x * ar: bz(5) = RY(b) + t2y * ar
    bz(6) = RX(b): bz(7) = RY(b)
End Sub

Private Sub BzPt(bz() As Double, ByVal t As Double, x As Double, y As Double)
    Dim mt As Double, a As Double, b As Double, c As Double, d As Double
    mt = 1 - t
    a = mt * mt * mt: b = 3 * mt * mt * t: c = 3 * mt * t * t: d = t * t * t
    x = a * bz(0) + b * bz(2) + c * bz(4) + d * bz(6)
    y = a * bz(1) + b * bz(3) + c * bz(5) + d * bz(7)
End Sub

' наибольшее отклонение точек участка от кубика (at — индекс точки в участке)
Private Function MaxErr(ByVal a As Long, ByVal b As Long, u() As Double, bz() As Double, at As Long) As Double
    Dim i As Long, x As Double, y As Double, d As Double, dmax As Double
    dmax = -1: at = (b - a + 1) \ 2
    For i = a + 1 To b - 1
        BzPt bz, u(i), x, y
        d = (x - RX(i)) ^ 2 + (y - RY(i)) ^ 2
        If d > dmax Then
            dmax = d: at = i - a
        End If
    Next i
    If dmax < 0 Then
        dmax = 0
    End If
    MaxErr = Sqr(dmax)
End Function

' шаг Ньютона: параметр ближайшей к (px, py) точки кубика
Private Function NewtonT(bz() As Double, ByVal px As Double, ByVal py As Double, ByVal t As Double) As Double
    Dim mt As Double, qx As Double, qy As Double, d1(0 To 5) As Double, d2(0 To 3) As Double
    Dim q1x As Double, q1y As Double, q2x As Double, q2y As Double, num As Double, den As Double, t2 As Double, k As Long
    mt = 1 - t
    BzPt bz, t, qx, qy
    For k = 0 To 5
        d1(k) = 3 * (bz(k + 2) - bz(k))
    Next k
    For k = 0 To 3
        d2(k) = 2 * (d1(k + 2) - d1(k))
    Next k
    q1x = d1(0) * mt * mt + d1(2) * 2 * mt * t + d1(4) * t * t
    q1y = d1(1) * mt * mt + d1(3) * 2 * mt * t + d1(5) * t * t
    q2x = d2(0) * mt + d2(2) * t
    q2y = d2(1) * mt + d2(3) * t
    num = (qx - px) * q1x + (qy - py) * q1y
    den = q1x * q1x + q1y * q1y + (qx - px) * q2x + (qy - py) * q2y
    If Abs(den) < 0.000000000001 Then
        NewtonT = t
        Exit Function
    End If
    t2 = t - num / den
    If t2 < 0 Then
        t2 = 0
    End If
    If t2 > 1 Then
        t2 = 1
    End If
    NewtonT = t2
End Function

Private Sub Emit(bz() As Double, ByVal isLine As Boolean)
    Dim k As Long
    If 4 * NO + 3 > UBound(OX) Then
        ReDim Preserve OX(0 To 8 * NO + 16): ReDim Preserve OY(0 To 8 * NO + 16)
    End If
    If NO > UBound(OL) Then
        ReDim Preserve OL(0 To 2 * NO + 4)
    End If
    For k = 0 To 3
        OX(4 * NO + k) = bz(2 * k): OY(4 * NO + k) = bz(2 * k + 1)
    Next k
    OL(NO) = isLine
    NO = NO + 1
End Sub

' ---------------------------------------------------------------------
'  Для стенда (LibreOffice): подпуть из массивов, результат в массивы
' ---------------------------------------------------------------------
Public Function SimpTest(inX() As Double, inY() As Double, ByVal nSeg As Long, ByVal closed As Boolean, ByVal tol As Double, _
                         outX() As Double, outY() As Double, outL() As Boolean) As Long
    Dim s As Long, k As Long
    NS = 0
    ReDim SX(0 To 4 * nSeg + 16): ReDim SY(0 To 4 * nSeg + 16)
    For s = 0 To nSeg - 1
        PutS s, inX(4 * s), inY(4 * s), inX(4 * s + 1), inY(4 * s + 1), inX(4 * s + 2), inY(4 * s + 2), inX(4 * s + 3), inY(4 * s + 3)
    Next s
    NS = nSeg
    SimpSubpath closed, tol
    ReDim outX(0 To 4 * NO): ReDim outY(0 To 4 * NO): ReDim outL(0 To NO)
    For s = 0 To NO - 1
        For k = 0 To 3
            outX(4 * s + k) = OX(4 * s + k): outY(4 * s + k) = OY(4 * s + k)
        Next k
        outL(s) = OL(s)
    Next s
    SimpTest = NO
End Function
'@@SIMP_END

'@@GROM_BEGIN
' =====================================================================
'  Люверсы (как proto/grommets.py): точки под люверсы по краю баннера.
'  Контур объекта (самый большой замкнутый подпуть) -> ломаная -> смещение внутрь на отступ
'  от края до центра люверса -> люверс в каждом углу, на сторонах между углами — равномерно,
'  не реже шага. Без углов (круг, овал) — равномерно по контуру, первый в верхней точке.
'  Угол — излом касательной больше 30 град или середина скругления угла. Усиление углов —
'  ещё по точке на обе стороны угла на заданном расстоянии от угловой.
'  Точки — кружки заданного диаметра и цвета на слое «Люверсы», по группе на объект.
' =====================================================================
Public Function NeonGromEdge() As Double
    NeonGromEdge = GromNum("Edge", 20)
End Function

Public Function NeonGromStep() As Double
    NeonGromStep = GromNum("Step", 300)
End Function

Public Function NeonGromDot() As Double
    NeonGromDot = GromNum("Dot", 5)
End Function

' усиление углов: расстояние от угловой точки, 0 — выключено
Public Function NeonGromReinf() As Double
    If GetSetting("NeonCenterline", "Grommet", "Reinf", "0") = "1" Then
        NeonGromReinf = GromNum("ReinfDist", 50)
    End If
End Function

Public Function NeonGromReinfDist() As Double
    NeonGromReinfDist = GromNum("ReinfDist", 50)
End Function

' цвет точки: "P:n" — из списка (0 чёрный, 1 красный, 2 пурпурный, 3 синий, 4 зелёный, 5 белый),
' "S:строка" — свой цвет (Color.ToString)
Public Function NeonGromColorSetting() As String
    NeonGromColorSetting = GetSetting("NeonCenterline", "Grommet", "Color", "P:0")
End Function

Private Function GromNum(ByVal key As String, ByVal def As Double) As Double
    Dim v As Double
    v = ParseNum(GetSetting("NeonCenterline", "Grommet", key, FmtNum(def)))
    If v <= 0 Then
        v = def
    End If
    GromNum = v
End Function

Private Function GromColor(ByVal cs As String) As Color
    Dim c As Color
    If Left$(cs, 2) = "S:" Then
        On Error Resume Next
        Set c = CreateCMYKColor(0, 0, 0, 100)
        c.StringAssign Mid$(cs, 3)
        If Err.Number = 0 Then
            Set GromColor = c
            Exit Function
        End If
        Err.Clear
        On Error GoTo 0
    End If
    Select Case CLng(Val(Mid$(cs, 3)))
    Case 1
        Set GromColor = CreateCMYKColor(0, 100, 100, 0)
    Case 2
        Set GromColor = CreateCMYKColor(0, 100, 0, 0)
    Case 3
        Set GromColor = CreateCMYKColor(100, 100, 0, 0)
    Case 4
        Set GromColor = CreateCMYKColor(100, 0, 100, 0)
    Case 5
        Set GromColor = CreateCMYKColor(0, 0, 0, 0)
    Case Else
        Set GromColor = CreateCMYKColor(0, 0, 0, 100)
    End Select
End Function

' свой цвет окном выбора цвета CorelDRAW; "" — отмена. UserAssign в CorelDRAW 2026 — процедура
' без результата: выбор (а не отмену) видно по тому, что цвет изменился
Public Function NeonGromPickColor(ByVal cur As String) As String
    Dim c As Color, was As String
    On Error GoTo Fail
    Set c = GromColor(cur)
    was = c.ToString
    c.UserAssign
    If c.ToString <> was Then
        NeonGromPickColor = "S:" & c.ToString
    End If
Fail:
End Function

' имя своего цвета для списка в окне
Public Function NeonGromColorName(ByVal cs As String) As String
    On Error Resume Next
    NeonGromColorName = GromColor(cs).Name
End Function

' ---------------------------------------------------------------------
'  Кнопка
' ---------------------------------------------------------------------
Public Sub PlaceGrommets()
    Dim sr As ShapeRange, col As Collection, oldUnit As cdrUnit, lay As Layer, clr As Color
    Dim edge As Double, stp As Double, dia As Double, rf As Double, i As Long, k As Long
    Dim nAll As Long, nObj As Long, nSkip As Long, msg As String, per As String, errText As String
    Dim grpOn As Boolean, dots As ShapeRange, dot As Shape, g As Shape, sh As Shape
    If ActiveDocument Is Nothing Then
        MsgBox NeonL("err_nodoc"), vbExclamation, NeonL("grom_title")
        Exit Sub
    End If
    Set sr = ActiveSelectionRange
    If sr.Count = 0 Then
        MsgBox NeonL("grom_nosel"), vbExclamation, NeonL("grom_title")
        Exit Sub
    End If
    If Not AskGrommetParams() Then
        Exit Sub
    End If
    edge = NeonGromEdge(): stp = NeonGromStep(): dia = NeonGromDot(): rf = NeonGromReinf()
    Set clr = GromColor(NeonGromColorSetting())
    Set col = New Collection
    CollectGrom sr, col
    oldUnit = ActiveDocument.Unit
    ActiveDocument.Unit = cdrMillimeter
    On Error GoTo Fail
    ActiveDocument.BeginCommandGroup NeonL("grom_title")
    grpOn = True
    Set lay = Nothing
    For i = 1 To col.Count
        Set sh = col(i)
        If ShapeGrommets(sh, edge, stp, rf) Then
            If lay Is Nothing Then
                Set lay = GromLayer()
            End If
            Set dots = CreateShapeRange
            For k = 0 To GRN - 1
                Set dot = lay.CreateEllipse2(GRX(k), GRY(k), dia / 2)
                dot.Fill.ApplyUniformFill clr
                dot.Outline.SetNoOutline
                dot.Name = NeonL("grom_dot")
                dots.Add dot
            Next k
            If dots.Count > 1 Then
                Set g = dots.Group
                g.Name = NeonLF("grom_group", GRN)
            End If
            nObj = nObj + 1: nAll = nAll + GRN
            If nObj <= 12 Then
                per = per & "  " & nObj & ") " & GRN & vbCrLf
            End If
        Else
            nSkip = nSkip + 1
        End If
    Next i
Fin:
    If grpOn Then
        ActiveDocument.EndCommandGroup
    End If
    ActiveDocument.Unit = oldUnit
    If Len(errText) > 0 Then
        MsgBox NeonLF("err_fail", errText), vbCritical, NeonL("grom_title")
        Exit Sub
    End If
    If nObj = 0 Then
        MsgBox NeonL("grom_none"), vbExclamation, NeonL("grom_title")
        Exit Sub
    End If
    msg = NeonLF("grom_report", nAll, nObj, FmtNum(edge), FmtNum(stp), FmtNum(dia))
    If rf > 0 Then
        msg = msg & vbCrLf & NeonLF("grom_report_reinf", FmtNum(rf))
    End If
    If nObj > 1 Then
        msg = msg & vbCrLf & vbCrLf & per
    End If
    If nSkip > 0 Then
        msg = msg & vbCrLf & NeonLF("grom_skipped", nSkip)
    End If
    If MsgBox(msg & vbCrLf & vbCrLf & NeonL("len_copy"), vbYesNo + vbInformation, NeonL("grom_title")) = vbYes Then
        CopyText msg
    End If
    NeonToolbarCheck
    Exit Sub
Fail:
    errText = Err.Description
    Resume Fin
End Sub

' объекты под люверсы: всё выделенное (группы — насквозь), кроме линий плагина и самих точек
Private Sub CollectGrom(sr As ShapeRange, col As Collection)
    Dim sh As Shape, ln As String
    For Each sh In sr
        ln = ""
        On Error Resume Next
        ln = sh.Layer.Name
        On Error GoTo 0
        If ln = NeonL("grom_layer") Then
            ' точки, поставленные раньше
        ElseIf sh.Type = cdrGroupShape Then
            CollectGrom sh.Shapes.All, col
        ElseIf Left$(sh.Name, Len(NeonPrefix())) <> NeonPrefix() Then
            col.Add sh
        End If
    Next sh
End Sub

' слой «Люверсы» на текущей странице
Private Function GromLayer() As Layer
    Dim ly As Layer, nm As String
    nm = NeonL("grom_layer")
    For Each ly In ActivePage.Layers
        If ly.Name = nm Then
            Set GromLayer = ly
            Exit Function
        End If
    Next ly
    Set GromLayer = ActivePage.CreateLayer(nm)
End Function

' люверсы одного объекта (GRX/GRY/GRN): самый большой замкнутый подпуть его контура
Private Function ShapeGrommets(sh As Shape, ByVal edge As Double, ByVal stp As Double, ByVal rf As Double) As Boolean
    Dim crv As Curve, tmp As Shape, sp As SubPath, sg As Segment, best As Double, a As Double
    Dim x0 As Double, y0 As Double, x1 As Double, y1 As Double, x2 As Double, y2 As Double, x3 As Double, y3 As Double
    Dim bx() As Double, by() As Double, bn As Long, k As Long
    On Error Resume Next
    If sh.Type = cdrCurveShape Then
        Set crv = sh.Curve
    Else
        ' прямоугольник, эллипс, текст...: кривая отображения, иначе — временная копия в кривых
        Set crv = sh.DisplayCurve
        If crv Is Nothing Then
            Set tmp = sh.Duplicate
            tmp.ConvertToCurves
            If tmp.Type = cdrCurveShape Then
                Set crv = tmp.Curve
            End If
        End If
    End If
    Err.Clear
    On Error GoTo 0
    If crv Is Nothing Then
        GoTo Done
    End If
    For Each sp In crv.SubPaths
        If sp.Closed Then
            NS = 0
            ReDim SX(0 To 63): ReDim SY(0 To 63)
            For Each sg In sp.Segments
                x0 = sg.StartNode.PositionX: y0 = sg.StartNode.PositionY
                x3 = sg.EndNode.PositionX: y3 = sg.EndNode.PositionY
                If sg.Type = cdrCurveSegment Then
                    sg.GetStartingControlPointPosition x1, y1
                    sg.GetEndingControlPointPosition x2, y2
                Else
                    x1 = x0 + (x3 - x0) / 3: y1 = y0 + (y3 - y0) / 3
                    x2 = x0 + 2 * (x3 - x0) / 3: y2 = y0 + 2 * (y3 - y0) / 3
                End If
                PutS NS, x0, y0, x1, y1, x2, y2, x3, y3
                NS = NS + 1
            Next sg
            If NS > 0 Then
                GromFlatten GromFlatStep(edge)
                a = Abs(GromArea())
                If a > best Then
                    best = a: bn = NS
                    ReDim bx(0 To 4 * NS): ReDim by(0 To 4 * NS)
                    For k = 0 To 4 * NS - 1
                        bx(k) = SX(k): by(k) = SY(k)
                    Next k
                End If
            End If
        End If
    Next sp
    If best > 0 Then
        NS = bn
        ReDim SX(0 To 4 * bn): ReDim SY(0 To 4 * bn)
        For k = 0 To 4 * bn - 1
            SX(k) = bx(k): SY(k) = by(k)
        Next k
        ShapeGrommets = GromPlace(edge, stp, rf)
    End If
Done:
    If Not tmp Is Nothing Then
        tmp.Delete
    End If
End Function

'@@GROMGEO_BEGIN
Private Function GromFlatStep(ByVal edge As Double) As Double
    Dim f As Double
    f = edge / 4
    If f > 1 Then
        f = 1
    End If
    If f < 0.2 Then
        f = 0.2
    End If
    GromFlatStep = f
End Function

' ---------------------------------------------------------------------
'  Геометрия (без CorelDRAW): вход — подпуть SX/SY/NS, выход — GRX/GRY/GRN
' ---------------------------------------------------------------------
Private Function GromPlace(ByVal edge As Double, ByVal stp As Double, ByVal rf As Double) As Boolean
    Dim ox() As Double, oy() As Double, oc() As Boolean, n As Long, i As Long, k As Long, a0 As Double
    Dim nx() As Double, ny() As Double, mx As Double, my As Double, c As Double, L As Double
    Dim drop() As Boolean, tix As Double, tiy As Double, tox As Double, toy As Double
    Dim kx() As Double, ky() As Double, kc() As Boolean, nk As Long
    Dim cum() As Double, total As Double, cs() As Long, nc As Long, j As Long, c2 As Long, Lr As Double, m As Long
    Dim top As Long, x As Double, y As Double, s As Double, ok As Boolean, q As Long
    GRN = 0
    ReDim GRX(0 To 15): ReDim GRY(0 To 15)
    GromFlatten GromFlatStep(edge)
    If GN < 3 Then
        Exit Function
    End If
    If GromArea() < 0 Then
        GromReverse
    End If
    a0 = GromArea()
    GromSoftCorners
    n = GN
    ' смещение внутрь: у вершины — по биссектрисе нормалей соседних звеньев на d / cos
    ReDim nx(0 To n - 1): ReDim ny(0 To n - 1)
    For i = 0 To n - 1
        k = (i + 1) Mod n
        NormV GPX(k) - GPX(i), GPY(k) - GPY(i), tox, toy
        nx(i) = -toy: ny(i) = tox
    Next i
    ReDim ox(0 To n - 1): ReDim oy(0 To n - 1)
    For i = 0 To n - 1
        k = (i - 1 + n) Mod n
        NormV nx(k) + nx(i), ny(k) + ny(i), mx, my
        c = mx * nx(i) + my * ny(i)
        If c > 0.25 Then
            L = edge / c
        Else
            L = 4 * edge
        End If
        ox(i) = GPX(i) + mx * L: oy(i) = GPY(i) + my * L
    Next i
    ' у выпуклого угла смещённые точки сторон ближе отступа к углу уходят за смещённый угол
    ' (петля): после угла — всё, что не впереди него по выходящей стороне, до угла — что не позади
    ReDim drop(0 To n - 1)
    For c2 = 0 To n - 1
        If GPC(c2) Then
            k = (c2 - 1 + n) Mod n
            NormV GPX(c2) - GPX(k), GPY(c2) - GPY(k), tix, tiy
            k = (c2 + 1) Mod n
            NormV GPX(k) - GPX(c2), GPY(k) - GPY(c2), tox, toy
            k = (c2 + 1) Mod n
            Do While k <> c2
                If GPC(k) Then
                    Exit Do
                End If
                If (ox(k) - ox(c2)) * tox + (oy(k) - oy(c2)) * toy > 0.000000001 Then
                    Exit Do
                End If
                drop(k) = True: k = (k + 1) Mod n
            Loop
            k = (c2 - 1 + n) Mod n
            Do While k <> c2
                If GPC(k) Then
                    Exit Do
                End If
                If (ox(k) - ox(c2)) * tix + (oy(k) - oy(c2)) * tiy < -0.000000001 Then
                    Exit Do
                End If
                drop(k) = True: k = (k - 1 + n) Mod n
            Loop
        End If
    Next c2
    ReDim kx(0 To n - 1): ReDim ky(0 To n - 1): ReDim kc(0 To n - 1): nk = 0
    For i = 0 To n - 1
        If Not drop(i) Then
            kx(nk) = ox(i): ky(nk) = oy(i): kc(nk) = GPC(i): nk = nk + 1
        End If
    Next i
    If nk < 3 Then
        Exit Function
    End If
    If PolyArea(kx, ky, nk) <= 0.02 * a0 Then
        Exit Function
    End If
    n = nk
    ReDim cum(0 To n)
    For i = 0 To n - 1
        k = (i + 1) Mod n
        cum(i + 1) = cum(i) + Sqr((kx(k) - kx(i)) ^ 2 + (ky(k) - ky(i)) ^ 2)
    Next i
    total = cum(n)
    ReDim cs(0 To n): nc = 0
    For i = 0 To n - 1
        If kc(i) Then
            cs(nc) = i: nc = nc + 1
        End If
    Next i
    If nc = 0 Then
        top = 0
        For i = 1 To n - 1
            If ky(i) > ky(top) Then
                top = i
            ElseIf ky(i) = ky(top) And kx(i) < kx(top) Then
                top = i
            End If
        Next i
        m = GromCeil(total / stp)
        For j = 0 To m - 1
            GromAt kx, ky, cum, n, top, total * j / m, x, y
            GromAdd x, y
        Next j
        GromPlace = True
        Exit Function
    End If
    For j = 0 To nc - 1
        k = cs(j): c2 = cs((j + 1) Mod nc)
        If c2 > k Then
            Lr = cum(c2) - cum(k)
        Else
            Lr = total - cum(k) + cum(c2)
        End If
        m = GromCeil(Lr / stp)
        For q = 0 To m - 1
            GromAt kx, ky, cum, n, k, Lr * q / m, x, y
            GromAdd x, y
        Next q
    Next j
    If rf > 0 Then
        nk = GRN
        For j = 0 To nc - 1
            For q = 0 To 1
                If q = 0 Then
                    s = rf
                Else
                    s = total - rf
                End If
                GromAt kx, ky, cum, n, cs(j), s, x, y
                ok = True
                For i = 0 To GRN - 1
                    If Sqr((GRX(i) - x) ^ 2 + (GRY(i) - y) ^ 2) < 0.5 * rf Then
                        ok = False
                        Exit For
                    End If
                Next i
                If ok Then
                    GromAdd x, y
                End If
            Next q
        Next j
    End If
    GromPlace = True
End Function

' число промежутков: вверх до целого, не меньше 1
Private Function GromCeil(ByVal v As Double) As Long
    Dim m As Long
    m = -Int(-(v - 0.000000001))
    If m < 1 Then
        m = 1
    End If
    GromCeil = m
End Function

Private Sub GromAdd(ByVal x As Double, ByVal y As Double)
    If GRN > UBound(GRX) Then
        ReDim Preserve GRX(0 To 2 * GRN + 16): ReDim Preserve GRY(0 To 2 * GRN + 16)
    End If
    GRX(GRN) = x: GRY(GRN) = y: GRN = GRN + 1
End Sub

' точка на расстоянии s по обходу смещённого контура от вершины i0
Private Sub GromAt(kx() As Double, ky() As Double, cum() As Double, ByVal n As Long, ByVal i0 As Long, _
                   ByVal s As Double, x As Double, y As Double)
    Dim s0 As Double, lo As Long, hi As Long, md As Long, i As Long, L As Double, t As Double, k As Long
    s0 = cum(i0) + s
    Do While s0 >= cum(n)
        s0 = s0 - cum(n)
    Loop
    lo = 0: hi = n
    Do While hi - lo > 1
        md = (lo + hi) \ 2
        If cum(md) <= s0 Then
            lo = md
        Else
            hi = md
        End If
    Loop
    i = lo: L = cum(i + 1) - cum(i)
    t = 0
    If L >= 0.000000000001 Then
        t = (s0 - cum(i)) / L
    End If
    k = (i + 1) Mod n
    x = kx(i) + (kx(k) - kx(i)) * t: y = ky(i) + (ky(k) - ky(i)) * t
End Sub

' подпуть SX/SY/NS -> ломаная GPX/GPY, GPC — угол в узле (излом касательной больше 30 град)
Private Sub GromFlatten(ByVal stp As Double)
    Dim s As Long, q As Long, m As Long, i As Long, x As Double, y As Double, Lc As Double
    Dim tix As Double, tiy As Double, tox As Double, toy As Double, isC As Boolean
    Dim bz(0 To 7) As Double
    GN = 0
    ReDim GPX(0 To 255): ReDim GPY(0 To 255): ReDim GPC(0 To 255)
    For s = 0 To NS - 1
        q = s - 1
        If q < 0 Then
            q = NS - 1
        End If
        GromTanEnd q, tix, tiy
        GromTanStart s, tox, toy
        isC = (tix * tox + tiy * toy < Cos(30# * Atn(1#) / 45#))
        For i = 0 To 7
            If i Mod 2 = 0 Then
                bz(i) = SX(4 * s + i \ 2)
            Else
                bz(i) = SY(4 * s + i \ 2)
            End If
        Next i
        Lc = Sqr((bz(2) - bz(0)) ^ 2 + (bz(3) - bz(1)) ^ 2) + Sqr((bz(4) - bz(2)) ^ 2 + (bz(5) - bz(3)) ^ 2) + _
             Sqr((bz(6) - bz(4)) ^ 2 + (bz(7) - bz(5)) ^ 2)
        m = -Int(-(Lc / stp))
        If m < 2 Then
            m = 2
        End If
        For i = 0 To m - 1
            BzPt bz, i / m, x, y
            If GN > 0 Then
                If Sqr((GPX(GN - 1) - x) ^ 2 + (GPY(GN - 1) - y) ^ 2) < 0.000001 Then
                    If isC And i = 0 Then
                        GPC(GN - 1) = True
                    End If
                    GoTo NextI
                End If
            End If
            If GN > UBound(GPX) Then
                ReDim Preserve GPX(0 To 2 * GN + 16): ReDim Preserve GPY(0 To 2 * GN + 16): ReDim Preserve GPC(0 To 2 * GN + 16)
            End If
            GPX(GN) = x: GPY(GN) = y: GPC(GN) = (isC And i = 0): GN = GN + 1
NextI:
        Next i
    Next s
    If GN > 1 Then
        If Sqr((GPX(0) - GPX(GN - 1)) ^ 2 + (GPY(0) - GPY(GN - 1)) ^ 2) < 0.000001 Then
            GPC(0) = GPC(0) Or GPC(GN - 1)
            GN = GN - 1
        End If
    End If
End Sub

' касательная в начале сегмента s; вырожденная ручка — к следующей точке
Private Sub GromTanStart(ByVal s As Long, tx As Double, ty As Double)
    Dim k As Long, dx As Double, dy As Double
    tx = 0: ty = 0
    For k = 1 To 3
        dx = SX(4 * s + k) - SX(4 * s): dy = SY(4 * s + k) - SY(4 * s)
        If Sqr(dx * dx + dy * dy) > 0.000000001 Then
            NormV dx, dy, tx, ty
            Exit Sub
        End If
    Next k
End Sub

Private Sub GromTanEnd(ByVal s As Long, tx As Double, ty As Double)
    Dim k As Long, dx As Double, dy As Double
    tx = 0: ty = 0
    For k = 2 To 0 Step -1
        dx = SX(4 * s + 3) - SX(4 * s + k): dy = SY(4 * s + 3) - SY(4 * s + k)
        If Sqr(dx * dx + dy * dy) > 0.000000001 Then
            NormV dx, dy, tx, ty
            Exit Sub
        End If
    Next k
End Sub

Private Function GromArea() As Double
    GromArea = PolyArea(GPX, GPY, GN)
End Function

Private Function PolyArea(x() As Double, y() As Double, ByVal n As Long) As Double
    Dim i As Long, k As Long, a As Double
    For i = 0 To n - 1
        k = (i + 1) Mod n
        a = a + x(i) * y(k) - x(k) * y(i)
    Next i
    PolyArea = a / 2
End Function

Private Sub GromReverse()
    Dim i As Long, j As Long, t As Double, b As Boolean
    i = 0: j = GN - 1
    Do While i < j
        t = GPX(i): GPX(i) = GPX(j): GPX(j) = t
        t = GPY(i): GPY(i) = GPY(j): GPY(j) = t
        b = GPC(i): GPC(i) = GPC(j): GPC(j) = b
        i = i + 1: j = j - 1
    Loop
End Sub

' скругления углов: подряд идущие вершины с радиусом кривизны меньше четверти меньшей стороны
' габарита, поворот участка не меньше 45 град; почти прямых участков (радиус больше 20 таких) —
' не меньше четверти периметра (у овала их нет). Середина участка по длине становится углом.
Private Sub GromSoftCorners()
    Dim n As Long, i As Long, k As Long, a As Long, b As Long, x0 As Double, x1 As Double, y0 As Double, y1 As Double
    Dim Rc As Double, sg() As Double, tn() As Double, cv() As Boolean, L As Double, total As Double, straight As Double
    Dim t0x As Double, t0y As Double, t1x As Double, t1y As Double, st As Long, tot As Double
    Dim r0 As Long, nr As Long, acc As Double, half As Double, md As Long, j As Long
    n = GN
    x0 = GPX(0): x1 = GPX(0): y0 = GPY(0): y1 = GPY(0)
    For i = 1 To n - 1
        If GPX(i) < x0 Then
            x0 = GPX(i)
        End If
        If GPX(i) > x1 Then
            x1 = GPX(i)
        End If
        If GPY(i) < y0 Then
            y0 = GPY(i)
        End If
        If GPY(i) > y1 Then
            y1 = GPY(i)
        End If
    Next i
    Rc = x1 - x0
    If y1 - y0 < Rc Then
        Rc = y1 - y0
    End If
    Rc = 0.25 * Rc
    ReDim sg(0 To n - 1): ReDim tn(0 To n - 1): ReDim cv(0 To n - 1)
    For i = 0 To n - 1
        k = (i + 1) Mod n
        sg(i) = Sqr((GPX(k) - GPX(i)) ^ 2 + (GPY(k) - GPY(i)) ^ 2)
        total = total + sg(i)
    Next i
    For i = 0 To n - 1
        a = (i - 1 + n) Mod n: b = (i + 1) Mod n
        NormV GPX(i) - GPX(a), GPY(i) - GPY(a), t0x, t0y
        NormV GPX(b) - GPX(i), GPY(b) - GPY(i), t1x, t1y
        tn(i) = GromAtan2(t0x * t1y - t0y * t1x, t0x * t1x + t0y * t1y)
        L = 0.5 * (sg(a) + sg(i))
        If Not GPC(i) And Abs(tn(i)) > 0.000000001 Then
            cv(i) = (L / Abs(tn(i)) < Rc)
        End If
        If Not GPC(i) Then
            If Abs(tn(i)) < 0.000000001 Then
                straight = straight + sg(i)
            ElseIf L / Abs(tn(i)) > 20 * Rc Then
                straight = straight + sg(i)
            End If
        End If
    Next i
    If straight < 0.25 * total Then
        Exit Sub
    End If
    st = -1
    For i = 0 To n - 1
        If Not cv(i) Then
            st = i
            Exit For
        End If
    Next i
    If st < 0 Then
        Exit Sub
    End If
    i = 0
    Do While i < n
        If Not cv((st + i) Mod n) Then
            i = i + 1
        Else
            r0 = i: nr = 0: tot = 0
            Do While i < n
                If Not cv((st + i) Mod n) Then
                    Exit Do
                End If
                tot = tot + tn((st + i) Mod n)
                nr = nr + 1: i = i + 1
            Loop
            If Abs(tot) >= 45# * Atn(1#) / 45# Then
                ' середина участка по длине (по звеньям между его вершинами)
                L = 0
                For j = 0 To nr - 2
                    L = L + sg((st + r0 + j) Mod n)
                Next j
                half = L / 2: acc = 0: md = (st + r0) Mod n
                For j = 0 To nr - 2
                    k = (st + r0 + j) Mod n
                    If acc + sg(k) >= half Then
                        If half - acc < acc + sg(k) - half Then
                            md = k
                        Else
                            md = (k + 1) Mod n
                        End If
                        Exit For
                    End If
                    acc = acc + sg(k)
                Next j
                GPC(md) = True
            End If
        End If
    Loop
End Sub

Private Function GromAtan2(ByVal y As Double, ByVal x As Double) As Double
    Const PI_ As Double = 3.14159265358979
    If x > 0 Then
        GromAtan2 = Atn(y / x)
    ElseIf x < 0 Then
        If y >= 0 Then
            GromAtan2 = Atn(y / x) + PI_
        Else
            GromAtan2 = Atn(y / x) - PI_
        End If
    ElseIf y > 0 Then
        GromAtan2 = PI_ / 2
    ElseIf y < 0 Then
        GromAtan2 = -PI_ / 2
    End If
End Function

' стенд: подпуть из массивов (4 точки на сегмент) -> центры люверсов
Public Function GromTest(inX() As Double, inY() As Double, ByVal nSeg As Long, ByVal edge As Double, ByVal stp As Double, _
                         ByVal rf As Double, outX() As Double, outY() As Double) As Long
    Dim s As Long, k As Long
    NS = 0
    ReDim SX(0 To 4 * nSeg + 16): ReDim SY(0 To 4 * nSeg + 16)
    For s = 0 To nSeg - 1
        PutS s, inX(4 * s), inY(4 * s), inX(4 * s + 1), inY(4 * s + 1), inX(4 * s + 2), inY(4 * s + 2), inX(4 * s + 3), inY(4 * s + 3)
    Next s
    NS = nSeg
    If Not GromPlace(edge, stp, rf) Then
        GromTest = -1
        Exit Function
    End If
    ReDim outX(0 To GRN): ReDim outY(0 To GRN)
    For k = 0 To GRN - 1
        outX(k) = GRX(k): outY(k) = GRY(k)
    Next k
    GromTest = GRN
End Function
'@@GROMGEO_END
'@@GROM_END

