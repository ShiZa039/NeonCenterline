Attribute VB_Name = "NeonTools"
' =====================================================================
'  NeonCenterline — инструменты:
'    MeasureLength   — длина выделенных контуров (для резки) и центральных линий (неон),
'                      результат можно скопировать в буфер обмена;
'    SelectNeonLines — выделить все центральные линии, созданные плагином, на странице;
'    NeonLineIds / NeonLinesNew — какие линии появились после расчёта (для подгонки под рез);
'    SimplifyContours — упростить выделенные кривые (меньше узлов в пределах допуска);
'    PlaceGrommets   — точки под люверсы по краю выделенных баннеров (подложек);
'    CleanContours   — очистить контур: убрать наложения и петли (каждый объект — одна кривая);
'    GroupSame       — одинаковые объекты (форма и цвет) — в группы;
'    NestSheets      — раскладка деталей на листы: по габаритам, без поворота, с полем и зазором;
'    SplitLetters    — разбить на буквы: по объекту на букву, дырки остаются в своих буквах.
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

' ---- очистка контура: вход, рабочие данные, выход (см. раздел «Очистка контура» ниже) ----
'@@CLEAN_DECL_BEGIN
Private Const CL_TOL As Double = 0.02
Private Const CL_MIN_AREA As Double = 0.3
Private Const CL_MIN_WIDTH As Double = 0.15
' вход: сегменты (4 точки на сегмент, признак прямой), замкнутые подконтуры
Private CSX() As Double, CSY() As Double, CSL() As Boolean, NCS As Long
Private CPS0() As Long, CPN() As Long, NCP As Long
' отрезки разбиения: подконтур, сегмент в нём, t0..t1, концы; ячейки сетки; первый отрезок подконтура
Private ESp() As Long, EG() As Long, ET0() As Double, ET1() As Double, NE As Long
Private EAX() As Double, EAY() As Double, EBX() As Double, EBY() As Double
Private ECx0() As Long, ECx1() As Long, ECy0() As Long, ECy1() As Long, ES0() As Long, ENn() As Long
' знак обмотки внутри подконтура (площадь со знаком)
Private CLSig() As Long
' разрезы (подконтур, параметр T = сегмент + t, точка), полосы для обмотки
Private CuSp() As Long, CuT() As Double, CuX() As Double, CuY() As Double, NCu As Long
Private CLCell As Double, CLY0 As Double, CLRowSt() As Long, CLRowL() As Long, CLNRow As Long
' куски-границы: подконтур, T0..T1, ход, целый подконтур, начало/вторая/предпоследняя/конец, площадь, длина
Private KSp() As Long, KT0() As Double, KT1() As Double, KFw() As Boolean, KWh() As Boolean, NK As Long
Private KSx() As Double, KSy() As Double, K2x() As Double, K2y() As Double, KPx() As Double, KPy() As Double
Private KEx() As Double, KEy() As Double, KAr() As Double, KLn() As Double
Private WPx() As Double, WPy() As Double
' выход: сегменты, подконтуры, число пересечений
Private COX() As Double, COY() As Double, COL() As Boolean, NCO As Long
Private COS0() As Long, CON() As Long, NCOP As Long, CLCuts As Long
'@@CLEAN_DECL_END

' ---- одинаковые формы и раскладка на листы (см. раздел «Раскладка» ниже) ----
'@@NEST_DECL_BEGIN
Private Const NT_EPS As Double = 0.000001
' листов раскладки в ряд
Private Const NT_ROW As Long = 5
' одинаковые: допуск, шаг точек по контуру, точек на объект не больше (у больших шаг крупнее)
Private Const SAME_TOL As Double = 0.1
Private Const SAME_STEP As Double = 0.5
Private Const SAME_MAXPTS As Double = 3000
' объекты: цвет (ключ, название), габарит по CorelDRAW, подпути (первый, число; 0 — контур не читался)
Private SmCk() As String, SmCn() As String, SmW() As Double, SmH() As Double, SmQ0() As Long, SmQn() As Long, NSm As Long
' подпути: первый сегмент, число, замкнут; сегменты: 4 точки, прямая
Private SqS0() As Long, SqN() As Long, SqC() As Boolean, NSq As Long
Private SgX() As Double, SgY() As Double, SgL() As Boolean, NSg As Long
' точки контура (у подпути — первая и число), габарит и длина контура по ним
Private SmX() As Double, SmY() As Double, NSmP As Long, QP0() As Long, QPn() As Long
Private SmSW() As Double, SmSH() As Double, SmLen() As Double
' результат: класс каждого объекта, число классов; сетка сравнения (отрезки, ячейки)
Private SmCls() As Long, NSmCls As Long
Private GrA() As Long, GrB() As Long, NGr As Long, GrCnt() As Long, GrSt() As Long, GrIdx() As Long
' разбить на буквы: почти замкнутый контур (мм), число букв, у буквы — подпути в LtIdx (первый, число)
Private Const LT_NEAR As Double = 0.05
' точки и акценты — к своей букве: знак не больше 0.4 буквы по большей стороне габарита или не больше 0.3
' её площади и 0.7 по стороне; центр — над буквой или под ней (её габарит + 0.25 её размера по
' горизонтали), не дальше 0.35 её размера по вертикали
Private Const MARK_SIZE As Double = 0.4
Private Const MARK_AREA As Double = 0.3
Private Const MARK_SIZE2 As Double = 0.7
Private Const MARK_X As Double = 0.25
Private Const MARK_GAP As Double = 0.35
Private LtNG As Long, LtG0() As Long, LtGn() As Long, LtIdx() As Long
' раскладка: детали (габарит без зазора), порядок, результат (лист, x, y от левого верхнего угла листа)
Private NtW() As Double, NtH() As Double, NNt As Long, NtOrd() As Long
Private NtS() As Long, NtX() As Double, NtY() As Double, NtSheets As Long, NtMode As Long, NtHeur As Long
' ширина каждого листа
Private NtSW() As Double
' проход: габарит с зазором, размещение; свободные прямоугольники (лист, x, y, ш, в) и новые при разрезе
Private NtPW() As Double, NtPH() As Double, RnS() As Long, RnX() As Double, RnY() As Double
Private FrS() As Long, FrX() As Double, FrY() As Double, FrW() As Double, FrH() As Double, NFr As Long
Private NwX() As Double, NwY() As Double, NwW() As Double, NwH() As Double, NNw As Long
'@@NEST_DECL_END

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

'@@CLEAN_BEGIN
' =====================================================================
'  NeonCenterline — очистка контура (кнопка «Неон: очистить контур», группа «Контур ▾»).
'
'  Наложения и петли в контурах: буквы, собранные из перекрывающихся частей (стойка поверх дуг),
'  налезающие друг на друга буквы, контур, перехлёстывающий сам себя (цифры, у которых дырка —
'  петля одного контура), фигуры после «Обводка в объект». Каждый выделенный объект становится
'  одной кривой (оригинал заменяется, одна операция — откат Ctrl+Z): внутренние линии наложений
'  уходят, дырки остаются, мелкие петли и лоскуты (меньше 0.3 мм² или тоньше 0.15 мм) — убираются.
'  Залитое — по правилу ненулевой обмотки (так CorelDRAW заливает текст в кривых). Текст и
'  многоугольники переводятся в кривые; прямоугольники и эллипсы пропускаются (наложений в них
'  не бывает); группы — по объектам. Отдельные объекты, которые надо слить, сначала объединить
'  (Ctrl+L). Незамкнутые подпути кривой остаются как были.
'
'  Алгоритм — прототип proto/cleanup.py (раздел CLEANGEO повторяет его один в один):
'    1. Сегменты разбиваются на отрезки (допуск 0.02 мм), у отрезка — сегмент и отрезок t.
'    2. Пересечения отрезков (сетка) режут подконтуры на куски.
'    3. Кусок остаётся, если залита ровно одна сторона (обмотка у середины слева и справа).
'    4. Куски сшиваются в контуры (в точках касания — самый правый поворот), мелкие отбрасываются.
'    5. Контур собирается из исходных сегментов, разрезанных точно в точках пересечения
'       (де Кастельжо) — форма не меняется, без перетрассировки.
' =====================================================================
Public Sub CleanContours()
    Dim col As Collection, oldUnit As cdrUnit, i As Long, nSkip As Long, grp As Boolean, errText As String, msg As String
    Dim nObj As Long, nCuts As Long, nSp0 As Long, nSp1 As Long
    If ActiveDocument Is Nothing Then
        MsgBox NeonL("err_nodoc"), vbExclamation, NeonL("app_title")
        Exit Sub
    End If
    If ActiveSelectionRange.Count = 0 Then
        MsgBox NeonL("err_nosel_any"), vbExclamation, NeonL("app_title")
        Exit Sub
    End If
    oldUnit = ActiveDocument.Unit
    ActiveDocument.Unit = cdrMillimeter
    ActiveDocument.BeginCommandGroup NeonL("clean_title")
    grp = True
    Optimization = True
    EventsEnabled = False
    On Error GoTo Fail
    Set col = New Collection
    CollectClean ActiveSelectionRange, col, nSkip
    For i = 1 To col.Count
        CleanShape col(i), nObj, nCuts, nSp0, nSp1
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
        MsgBox NeonLF("err_fail", errText), vbCritical, NeonL("clean_title")
        Exit Sub
    End If
    If col.Count = 0 Then
        msg = NeonL("clean_none")
    ElseIf nObj = 0 Then
        msg = NeonL("clean_nochange")
    Else
        msg = NeonLF("clean_report", nObj, col.Count, nCuts, nSp0, nSp1)
    End If
    If nSkip > 0 Then
        msg = msg & vbCrLf & vbCrLf & NeonLF("clean_skipped", nSkip)
    End If
    MsgBox msg, vbInformation, NeonL("clean_title")
    NeonToolbarCheck
    Exit Sub
Fail:
    errText = Err.Description
    Resume Fin
End Sub

' кривые выделения; текст и многоугольники — в кривые (в той же операции отмены)
Private Sub CollectClean(sr As ShapeRange, col As Collection, nSkip As Long)
    Dim sh As Shape, tp As Long
    For Each sh In sr
        tp = sh.Type
        If tp = cdrGroupShape Then
            CollectClean sh.Shapes.All, col, nSkip
        ElseIf tp = cdrCurveShape Then
            col.Add sh
        ElseIf tp = cdrRectangleShape Or tp = cdrEllipseShape Then
            ' наложений не бывает — не трогать
        ElseIf tp = cdrTextShape Or tp = cdrPolygonShape Then
            On Error Resume Next
            sh.ConvertToCurves
            On Error GoTo 0
            If sh.Type = cdrCurveShape Then
                col.Add sh
            ElseIf sh.Type = cdrGroupShape Then
                CollectClean sh.Shapes.All, col, nSkip
            Else
                nSkip = nSkip + 1
            End If
        Else
            nSkip = nSkip + 1
        End If
    Next sh
End Sub

' очистить кривую фигуры: замкнутые подпути — через CleanGeo, незамкнутые — как были
Private Sub CleanShape(shp As Shape, nObj As Long, nCuts As Long, nSp0 As Long, nSp1 As Long)
    Dim sp As SubPath, sg As Segment, crv As Curve, isClosed As Boolean, i As Long, s As Long, n0 As Long
    Dim x0 As Double, y0 As Double, x1 As Double, y1 As Double, x2 As Double, y2 As Double, x3 As Double, y3 As Double
    Dim ox() As Double, oy() As Double, ol() As Boolean, nos As Long, os0() As Long, on_() As Long, nop As Long
    NCS = 0: NCP = 0: nos = 0: nop = 0
    ReDim CSX(0 To 255): ReDim CSY(0 To 255): ReDim CSL(0 To 63): ReDim CPS0(0 To 15): ReDim CPN(0 To 15)
    ReDim ox(0 To 63): ReDim oy(0 To 63): ReDim ol(0 To 15): ReDim os0(0 To 3): ReDim on_(0 To 3)
    For Each sp In shp.Curve.SubPaths
        isClosed = sp.Closed
        If Not isClosed Then
            On Error Resume Next
            isClosed = (Abs(sp.StartNode.PositionX - sp.EndNode.PositionX) < 0.05 And _
                        Abs(sp.StartNode.PositionY - sp.EndNode.PositionY) < 0.05 And sp.Segments.Count >= 2)
            On Error GoTo 0
        End If
        n0 = NCS
        For Each sg In sp.Segments
            x0 = sg.StartNode.PositionX: y0 = sg.StartNode.PositionY
            x3 = sg.EndNode.PositionX: y3 = sg.EndNode.PositionY
            If sg.Type = cdrCurveSegment Then
                sg.GetStartingControlPointPosition x1, y1
                sg.GetEndingControlPointPosition x2, y2
                AddCS x0, y0, x1, y1, x2, y2, x3, y3, False
            Else
                AddCS x0, y0, 0, 0, 0, 0, x3, y3, True
            End If
        Next sg
        If NCS > n0 And isClosed Then
            ' почти замкнутый — замкнуть отрезком
            If CSX(4 * (NCS - 1) + 3) <> CSX(4 * n0) Or CSY(4 * (NCS - 1) + 3) <> CSY(4 * n0) Then
                AddCS CSX(4 * (NCS - 1) + 3), CSY(4 * (NCS - 1) + 3), 0, 0, 0, 0, CSX(4 * n0), CSY(4 * n0), True
            End If
            If NCP > UBound(CPS0) Then
                ReDim Preserve CPS0(0 To 2 * NCP): ReDim Preserve CPN(0 To 2 * NCP)
            End If
            CPS0(NCP) = n0: CPN(NCP) = NCS - n0: NCP = NCP + 1
        ElseIf NCS > n0 Then
            ' незамкнутый — отложить как есть
            If nop > UBound(os0) Then
                ReDim Preserve os0(0 To 2 * nop): ReDim Preserve on_(0 To 2 * nop)
            End If
            os0(nop) = nos: on_(nop) = NCS - n0: nop = nop + 1
            For s = n0 To NCS - 1
                If 4 * nos + 3 > UBound(ox) Then
                    ReDim Preserve ox(0 To 8 * nos + 16): ReDim Preserve oy(0 To 8 * nos + 16)
                End If
                If nos > UBound(ol) Then
                    ReDim Preserve ol(0 To 2 * nos + 4)
                End If
                For i = 0 To 3
                    ox(4 * nos + i) = CSX(4 * s + i): oy(4 * nos + i) = CSY(4 * s + i)
                Next i
                ol(nos) = CSL(s): nos = nos + 1
            Next s
            NCS = n0
        End If
    Next sp
    If NCP = 0 Then
        Exit Sub
    End If
    CleanGeo CL_MIN_AREA, CL_MIN_WIDTH
    ' ничего не пересекалось и ничего не выброшено — кривую не трогать; всё мелкое — тоже
    If (CLCuts = 0 And NCOP = NCP) Or NCOP = 0 Then
        Exit Sub
    End If
    Set crv = Application.CreateCurve(ActiveDocument)
    For i = 0 To NCOP - 1
        AppendSubCO crv, COX, COY, COL, COS0(i), CON(i), True
    Next i
    For i = 0 To nop - 1
        AppendSubCO crv, ox, oy, ol, os0(i), on_(i), False
    Next i
    shp.Curve.CopyAssign crv
    nObj = nObj + 1: nCuts = nCuts + CLCuts: nSp0 = nSp0 + NCP: nSp1 = nSp1 + NCOP
End Sub

Private Sub AddCS(ByVal x0 As Double, ByVal y0 As Double, ByVal x1 As Double, ByVal y1 As Double, _
                  ByVal x2 As Double, ByVal y2 As Double, ByVal x3 As Double, ByVal y3 As Double, ByVal isL As Boolean)
    If 4 * NCS + 3 > UBound(CSX) Then
        ReDim Preserve CSX(0 To 8 * NCS + 16): ReDim Preserve CSY(0 To 8 * NCS + 16)
    End If
    If NCS > UBound(CSL) Then
        ReDim Preserve CSL(0 To 2 * NCS + 4)
    End If
    If isL Then
        x1 = x0 + (x3 - x0) / 3: y1 = y0 + (y3 - y0) / 3
        x2 = x0 + 2 * (x3 - x0) / 3: y2 = y0 + 2 * (y3 - y0) / 3
    End If
    CSX(4 * NCS) = x0: CSY(4 * NCS) = y0: CSX(4 * NCS + 1) = x1: CSY(4 * NCS + 1) = y1
    CSX(4 * NCS + 2) = x2: CSY(4 * NCS + 2) = y2: CSX(4 * NCS + 3) = x3: CSY(4 * NCS + 3) = y3
    CSL(NCS) = isL
    NCS = NCS + 1
End Sub

' подпуть в кривую: сегменты s0..s0+n-1 (4 точки на сегмент)
Private Sub AppendSubCO(crv As Curve, bx() As Double, by() As Double, bl() As Boolean, ByVal s0 As Long, _
                        ByVal n As Long, ByVal cl As Boolean)
    Dim sp As SubPath, s As Long
    If n <= 0 Then
        Exit Sub
    End If
    Set sp = crv.CreateSubPath(bx(4 * s0), by(4 * s0))
    For s = s0 To s0 + n - 1
        If bl(s) Then
            sp.AppendLineSegment bx(4 * s + 3), by(4 * s + 3)
        Else
            sp.AppendCurveSegment2 bx(4 * s + 3), by(4 * s + 3), bx(4 * s + 1), by(4 * s + 1), bx(4 * s + 2), by(4 * s + 2)
        End If
    Next s
    If cl Then
        On Error Resume Next
        sp.Closed = True
        On Error GoTo 0
    End If
End Sub

'@@CLEANGEO_BEGIN
' ---------------------------------------------------------------------
'  Очистка контура: геометрия (повторяет proto/cleanup.py один в один).
'  Вход: сегменты CSX/CSY (4 точки на сегмент), CSL (прямая), NCS; замкнутые подконтуры
'  CPS0 (первый сегмент), CPN (сегментов), NCP. Выход: COX/COY/COL/NCO, подконтуры COS0/CON/NCOP;
'  CLCuts — число пересечений.
' ---------------------------------------------------------------------
Public Sub CleanGeo(ByVal minArea As Double, ByVal minWidth As Double)
    Dim sp As Long, s As Long, n As Long, k As Long, i As Long, j As Long, e As Long
    Dim x0 As Double, y0 As Double, x1 As Double, y1 As Double, x2 As Double, y2 As Double, x3 As Double, y3 As Double
    Dim dd As Double, d2 As Double, t As Double, mt As Double, px As Double, py As Double, qx As Double, qy As Double
    Dim sumL As Double, cell As Double, mnX As Double, mnY As Double, mxX As Double, mxY As Double
    Dim gx As Long, gy As Long, nx As Long, ny As Long, c As Long, gx0 As Long, gx1 As Long, gy0 As Long, gy1 As Long
    Dim cnt() As Long, st() As Long, lst() As Long, a As Long, b As Long, ia As Long, ib As Long
    Dim u As Double, v As Double, hx As Double, hy As Double, NN As Long, TT As Double
    NCO = 0: NCOP = 0: CLCuts = 0
    ReDim COX(0 To 63): ReDim COY(0 To 63): ReDim COL(0 To 15): ReDim COS0(0 To 15): ReDim CON(0 To 15)
    If NCP = 0 Then
        Exit Sub
    End If
    ' 1. разбиение на отрезки
    NE = 0
    ReDim ESp(0 To 1023): ReDim EG(0 To 1023): ReDim ET0(0 To 1023): ReDim ET1(0 To 1023)
    ReDim EAX(0 To 1023): ReDim EAY(0 To 1023): ReDim EBX(0 To 1023): ReDim EBY(0 To 1023)
    ReDim ES0(0 To NCP): ReDim ENn(0 To NCP)
    For sp = 0 To NCP - 1
        ES0(sp) = NE
        For s = CPS0(sp) To CPS0(sp) + CPN(sp) - 1
            x0 = CSX(4 * s): y0 = CSY(4 * s): x3 = CSX(4 * s + 3): y3 = CSY(4 * s + 3)
            If CSL(s) Then
                PushE sp, s - CPS0(sp), 0, 1, x0, y0, x3, y3
            Else
                x1 = CSX(4 * s + 1): y1 = CSY(4 * s + 1): x2 = CSX(4 * s + 2): y2 = CSY(4 * s + 2)
                dd = Sqr((x0 - 2 * x1 + x2) ^ 2 + (y0 - 2 * y1 + y2) ^ 2)
                d2 = Sqr((x1 - 2 * x2 + x3) ^ 2 + (y1 - 2 * y2 + y3) ^ 2)
                If d2 > dd Then
                    dd = d2
                End If
                n = -Int(-Sqr(0.75 * dd / CL_TOL))
                If n < 2 Then
                    n = 2
                End If
                If n > 256 Then
                    n = 256
                End If
                px = x0: py = y0
                For k = 1 To n
                    If k = n Then
                        qx = x3: qy = y3
                    Else
                        t = k / n: mt = 1 - t
                        qx = mt * mt * mt * x0 + 3 * mt * mt * t * x1 + 3 * mt * t * t * x2 + t * t * t * x3
                        qy = mt * mt * mt * y0 + 3 * mt * mt * t * y1 + 3 * mt * t * t * y2 + t * t * t * y3
                    End If
                    PushE sp, s - CPS0(sp), (k - 1) / n, k / n, px, py, qx, qy
                    px = qx: py = qy
                Next k
            End If
        Next s
        ENn(sp) = NE - ES0(sp)
    Next sp
    If NE = 0 Then
        Exit Sub
    End If
    ' знак обмотки внутри подконтура: площадь со знаком (у петляющего — по большей части)
    ReDim CLSig(0 To NCP)
    For sp = 0 To NCP - 1
        dd = 0
        For e = ES0(sp) To ES0(sp) + ENn(sp) - 1
            If e = ES0(sp) Then
                k = ES0(sp) + ENn(sp) - 1
            Else
                k = e - 1
            End If
            dd = dd + (EAX(k) * EAY(e) - EAX(e) * EAY(k))
        Next e
        If dd > 0 Then
            CLSig(sp) = 1
        Else
            CLSig(sp) = -1
        End If
    Next sp
    ' 2. сетка по габаритам отрезков (ячейка — 8 средних длин, не меньше 1 мм)
    mnX = EAX(0): mnY = EAY(0): mxX = mnX: mxY = mnY
    For e = 0 To NE - 1
        sumL = sumL + Sqr((EBX(e) - EAX(e)) ^ 2 + (EBY(e) - EAY(e)) ^ 2)
        MinMax4 EAX(e), EBX(e), mnX, mxX
        MinMax4 EAY(e), EBY(e), mnY, mxY
    Next e
    cell = 8 * sumL / NE
    If cell < 1 Then
        cell = 1
    End If
    CLCell = cell: CLY0 = mnY
    ' слишком мелкая сетка на большом объекте — крупнее (на результат не влияет)
    Do While ((mxX - mnX) / cell + 1) * ((mxY - mnY) / cell + 1) > 4000000#
        cell = cell * 2
    Loop
    nx = Int((mxX - mnX) / cell) + 1: ny = Int((mxY - mnY) / cell) + 1
    ReDim cnt(0 To nx * ny): ReDim st(0 To nx * ny + 1)
    ReDim ECx0(0 To NE): ReDim ECx1(0 To NE): ReDim ECy0(0 To NE): ReDim ECy1(0 To NE)
    For e = 0 To NE - 1
        CellRange e, mnX, mnY, cell, gx0, gx1, gy0, gy1
        ECx0(e) = gx0: ECx1(e) = gx1: ECy0(e) = gy0: ECy1(e) = gy1
        For gy = gy0 To gy1
            For gx = gx0 To gx1
                cnt(gy * nx + gx) = cnt(gy * nx + gx) + 1
            Next gx
        Next gy
    Next e
    For c = 0 To nx * ny - 1
        st(c + 1) = st(c) + cnt(c)
        cnt(c) = 0
    Next c
    ReDim lst(0 To st(nx * ny) + 1)
    For e = 0 To NE - 1
        For gy = ECy0(e) To ECy1(e)
            For gx = ECx0(e) To ECx1(e)
                c = gy * nx + gx
                lst(st(c) + cnt(c)) = e: cnt(c) = cnt(c) + 1
            Next gx
        Next gy
    Next e
    ' 3. пересечения: каждая пара — один раз, в первой общей ячейке
    NCu = 0
    ReDim CuSp(0 To 255): ReDim CuT(0 To 255): ReDim CuX(0 To 255): ReDim CuY(0 To 255)
    For c = 0 To nx * ny - 1
        For i = st(c) To st(c + 1) - 1
            a = lst(i)
            For j = i + 1 To st(c + 1) - 1
                b = lst(j)
                If FirstCommonCell(a, b, nx) = c Then
                    If Not (ESp(a) = ESp(b) And ((EBX(a) = EAX(b) And EBY(a) = EAY(b)) Or (EBX(b) = EAX(a) And EBY(b) = EAY(a)))) Then
                        If SegCross(EAX(a), EAY(a), EBX(a), EBY(a), EAX(b), EAY(b), EBX(b), EBY(b), u, v) Then
                            hx = EAX(a) + (EBX(a) - EAX(a)) * u: hy = EAY(a) + (EBY(a) - EAY(a)) * u
                            NN = CPN(ESp(a))
                            TT = EG(a) + ET0(a) + (ET1(a) - ET0(a)) * u
                            PushCut ESp(a), FMod(TT, NN), hx, hy
                            NN = CPN(ESp(b))
                            TT = EG(b) + ET0(b) + (ET1(b) - ET0(b)) * v
                            PushCut ESp(b), FMod(TT, NN), hx, hy
                            CLCuts = CLCuts + 1
                        End If
                    End If
                End If
            Next j
        Next i
    Next c
    SortCuts
    ' 4. полосы по y для обмотки
    Dim ry0 As Long, ry1 As Long, nrw As Long, rcnt() As Long, rst() As Long, rl() As Long, ylo As Double, yhi As Double
    nrw = Int((mxY - mnY) / CLCell) + 1
    ReDim rcnt(0 To nrw): ReDim rst(0 To nrw + 1)
    For e = 0 To NE - 1
        RowRange e, ry0, ry1
        For k = ry0 To ry1
            rcnt(k) = rcnt(k) + 1
        Next k
    Next e
    For k = 0 To nrw - 1
        rst(k + 1) = rst(k) + rcnt(k): rcnt(k) = 0
    Next k
    ReDim rl(0 To rst(nrw) + 1)
    For e = 0 To NE - 1
        RowRange e, ry0, ry1
        For k = ry0 To ry1
            rl(rst(k) + rcnt(k)) = e: rcnt(k) = rcnt(k) + 1
        Next k
    Next e
    CLRowSt = rst: CLRowL = rl: CLNRow = nrw
    ' 5. куски подконтуров между разрезами; остаются границы залитого и пустого
    NK = 0
    ReDim KSp(0 To 63): ReDim KT0(0 To 63): ReDim KT1(0 To 63): ReDim KFw(0 To 63): ReDim KWh(0 To 63)
    ReDim KSx(0 To 63): ReDim KSy(0 To 63): ReDim K2x(0 To 63): ReDim K2y(0 To 63): ReDim KPx(0 To 63)
    ReDim KPy(0 To 63): ReDim KEx(0 To 63): ReDim KEy(0 To 63): ReDim KAr(0 To 63): ReDim KLn(0 To 63)
    ReDim WPx(0 To 15): ReDim WPy(0 To 15)
    Dim c0 As Long, c1 As Long, nc As Long, T0 As Double, T1 As Double
    c0 = 0
    For sp = 0 To NCP - 1
        c1 = c0
        Do While c1 < NCu
            If CuSp(c1) <> sp Then
                Exit Do
            End If
            c1 = c1 + 1
        Loop
        nc = c1 - c0
        NN = CPN(sp)
        If nc = 0 Then
            TestPiece sp, 0#, CDbl(NN), 0#, 0#, 0#, 0#, True
        Else
            For k = 0 To nc - 1
                T0 = CuT(c0 + k)
                If nc = 1 Then
                    T1 = T0 + NN
                Else
                    T1 = CuT(c0 + (k + 1) Mod nc)
                    If T1 <= T0 Then
                        T1 = T1 + NN
                    End If
                End If
                TestPiece sp, T0, T1, CuX(c0 + k), CuY(c0 + k), CuX(c0 + (k + 1) Mod nc), CuY(c0 + (k + 1) Mod nc), False
            Next k
        End If
        c0 = c1
    Next sp
    ' 6. сшивка в контуры, отброс мелких, сборка из исходных сегментов
    ChainPieces minArea, minWidth
End Sub

Private Sub PushE(ByVal sp As Long, ByVal g As Long, ByVal t0 As Double, ByVal t1 As Double, _
                  ByVal ax As Double, ByVal ay As Double, ByVal bx As Double, ByVal by As Double)
    If NE > UBound(ESp) Then
        ReDim Preserve ESp(0 To 2 * NE): ReDim Preserve EG(0 To 2 * NE)
        ReDim Preserve ET0(0 To 2 * NE): ReDim Preserve ET1(0 To 2 * NE)
        ReDim Preserve EAX(0 To 2 * NE): ReDim Preserve EAY(0 To 2 * NE)
        ReDim Preserve EBX(0 To 2 * NE): ReDim Preserve EBY(0 To 2 * NE)
    End If
    ESp(NE) = sp: EG(NE) = g: ET0(NE) = t0: ET1(NE) = t1
    EAX(NE) = ax: EAY(NE) = ay: EBX(NE) = bx: EBY(NE) = by
    NE = NE + 1
End Sub

Private Sub MinMax4(ByVal a As Double, ByVal b As Double, mn As Double, mx As Double)
    If a < mn Then
        mn = a
    End If
    If b < mn Then
        mn = b
    End If
    If a > mx Then
        mx = a
    End If
    If b > mx Then
        mx = b
    End If
End Sub

Private Sub CellRange(ByVal e As Long, ByVal mnX As Double, ByVal mnY As Double, ByVal cell As Double, _
                      gx0 As Long, gx1 As Long, gy0 As Long, gy1 As Long)
    Dim a As Double, b As Double
    a = EAX(e): b = EBX(e)
    If a > b Then
        a = EBX(e): b = EAX(e)
    End If
    gx0 = Int((a - mnX) / cell): gx1 = Int((b - mnX) / cell)
    a = EAY(e): b = EBY(e)
    If a > b Then
        a = EBY(e): b = EAY(e)
    End If
    gy0 = Int((a - mnY) / cell): gy1 = Int((b - mnY) / cell)
End Sub

' первая (нижняя левая) общая ячейка двух отрезков; -1 — общих нет
Private Function FirstCommonCell(ByVal a As Long, ByVal b As Long, ByVal nx As Long) As Long
    Dim x0 As Long, y0 As Long
    If ECx1(a) < ECx0(b) Or ECx1(b) < ECx0(a) Or ECy1(a) < ECy0(b) Or ECy1(b) < ECy0(a) Then
        FirstCommonCell = -1
        Exit Function
    End If
    x0 = ECx0(a): y0 = ECy0(a)
    If ECx0(b) > x0 Then
        x0 = ECx0(b)
    End If
    If ECy0(b) > y0 Then
        y0 = ECy0(b)
    End If
    FirstCommonCell = y0 * nx + x0
End Function

Private Sub RowRange(ByVal e As Long, r0 As Long, r1 As Long)
    Dim a As Double, b As Double
    a = EAY(e): b = EBY(e)
    If a > b Then
        a = EBY(e): b = EAY(e)
    End If
    r0 = Int((a - CLY0) / CLCell): r1 = Int((b - CLY0) / CLCell)
End Sub

' пересечение отрезков (доли u на первом, v на втором), концы включительно
Private Function SegCross(ByVal ax As Double, ByVal ay As Double, ByVal bx As Double, ByVal by As Double, _
                          ByVal cx As Double, ByVal cy As Double, ByVal dx As Double, ByVal dy As Double, _
                          u As Double, v As Double) As Boolean
    Dim ex As Double, ey As Double, fx As Double, fy As Double, den As Double
    ex = bx - ax: ey = by - ay: fx = dx - cx: fy = dy - cy
    den = ex * fy - ey * fx
    If Abs(den) < 0.00000000000001 Then
        Exit Function
    End If
    u = ((cx - ax) * fy - (cy - ay) * fx) / den
    v = ((cx - ax) * ey - (cy - ay) * ex) / den
    If u >= -0.000000000001 And u <= 1.000000000001 And v >= -0.000000000001 And v <= 1.000000000001 Then
        SegCross = True
    End If
End Function

Private Function FMod(ByVal x As Double, ByVal m As Double) As Double
    FMod = x - m * Int(x / m)
End Function

Private Sub PushCut(ByVal sp As Long, ByVal T As Double, ByVal x As Double, ByVal y As Double)
    If NCu > UBound(CuSp) Then
        ReDim Preserve CuSp(0 To 2 * NCu): ReDim Preserve CuT(0 To 2 * NCu)
        ReDim Preserve CuX(0 To 2 * NCu): ReDim Preserve CuY(0 To 2 * NCu)
    End If
    CuSp(NCu) = sp: CuT(NCu) = T: CuX(NCu) = x: CuY(NCu) = y
    NCu = NCu + 1
End Sub

' разрезы по (подконтур, T); совпадающие (|dT| < 1e-9) — один
Private Sub SortCuts()
    Dim idx() As Long, i As Long, m As Long, ns() As Long, nt() As Double, nx() As Double, ny() As Double
    If NCu = 0 Then
        Exit Sub
    End If
    ReDim idx(0 To NCu - 1)
    For i = 0 To NCu - 1
        idx(i) = i
    Next i
    QSortCuts idx, 0, NCu - 1
    ReDim ns(0 To NCu): ReDim nt(0 To NCu): ReDim nx(0 To NCu): ReDim ny(0 To NCu)
    m = 0
    For i = 0 To NCu - 1
        If m > 0 Then
            If ns(m - 1) = CuSp(idx(i)) And Abs(CuT(idx(i)) - nt(m - 1)) < 0.000000001 Then
                GoTo NextCut
            End If
        End If
        ns(m) = CuSp(idx(i)): nt(m) = CuT(idx(i)): nx(m) = CuX(idx(i)): ny(m) = CuY(idx(i))
        m = m + 1
NextCut:
    Next i
    CuSp = ns: CuT = nt: CuX = nx: CuY = ny: NCu = m
End Sub

Private Function CutLess(ByVal a As Long, ByVal b As Long) As Boolean
    If CuSp(a) <> CuSp(b) Then
        CutLess = (CuSp(a) < CuSp(b))
    ElseIf CuT(a) <> CuT(b) Then
        CutLess = (CuT(a) < CuT(b))
    Else
        CutLess = (a < b)
    End If
End Function

Private Sub QSortCuts(idx() As Long, ByVal lo As Long, ByVal hi As Long)
    Dim i As Long, j As Long, p As Long, tmp As Long
    Do While lo < hi
        i = lo: j = hi: p = idx((lo + hi) \ 2)
        Do While i <= j
            Do While CutLess(idx(i), p)
                i = i + 1
            Loop
            Do While CutLess(p, idx(j))
                j = j - 1
            Loop
            If i <= j Then
                tmp = idx(i): idx(i) = idx(j): idx(j) = tmp
                i = i + 1: j = j - 1
            End If
        Loop
        If j - lo < hi - i Then
            QSortCuts idx, lo, j
            lo = i
        Else
            QSortCuts idx, i, hi
            hi = j
        End If
    Loop
End Sub

' вершины подконтура sp на параметре (T0, T1) по порядку (T1 может быть > N — по кругу):
' возвращает их число, точки — в WPx/WPy (с X0 в начале и X1 в конце)
Private Function PieceLine(ByVal sp As Long, ByVal T0 As Double, ByVal T1 As Double, _
                           ByVal x0 As Double, ByVal y0 As Double, ByVal x1 As Double, ByVal y1 As Double, _
                           ByVal whole As Boolean) As Long
    Dim e As Long, n As Long, Tv As Double, NN As Long, pass As Long
    NN = CPN(sp)
    If UBound(WPx) < ENn(sp) + 4 Then
        ReDim WPx(0 To ENn(sp) + 4): ReDim WPy(0 To ENn(sp) + 4)
    End If
    If whole Then
        For e = ES0(sp) To ES0(sp) + ENn(sp) - 1
            WPx(n) = EAX(e): WPy(n) = EAY(e): n = n + 1
        Next e
        WPx(n) = EAX(ES0(sp)): WPy(n) = EAY(ES0(sp)): n = n + 1
        PieceLine = n
        Exit Function
    End If
    WPx(0) = x0: WPy(0) = y0: n = 1
    If T1 > NN Then
        ' по кругу: вершины после T0, затем с начала до T1 - N
        For e = ES0(sp) To ES0(sp) + ENn(sp) - 1
            Tv = FMod(EG(e) + ET0(e), NN)
            If Tv > T0 Then
                WPx(n) = EAX(e): WPy(n) = EAY(e): n = n + 1
            End If
        Next e
        For e = ES0(sp) To ES0(sp) + ENn(sp) - 1
            Tv = FMod(EG(e) + ET0(e), NN)
            If Tv + NN < T1 Then
                WPx(n) = EAX(e): WPy(n) = EAY(e): n = n + 1
            End If
        Next e
    Else
        For e = ES0(sp) To ES0(sp) + ENn(sp) - 1
            Tv = FMod(EG(e) + ET0(e), NN)
            If T0 < Tv And Tv < T1 Then
                WPx(n) = EAX(e): WPy(n) = EAY(e): n = n + 1
            End If
        Next e
    End If
    WPx(n) = x1: WPy(n) = y1: n = n + 1
    PieceLine = n
End Function

' кусок: обмотка слева и справа от середины; граница залитого и пустого — в список (залитое справа)
Private Sub TestPiece(ByVal sp As Long, ByVal T0 As Double, ByVal T1 As Double, _
                      ByVal x0 As Double, ByVal y0 As Double, ByVal x1 As Double, ByVal y1 As Double, _
                      ByVal whole As Boolean)
    Dim n As Long, i As Long, tot As Double, d As Double, mx As Double, my As Double, im As Long, dm As Double
    Dim tx As Double, ty As Double, wl As Long, wr As Long, fw As Boolean, ar As Double
    Dim ol As Long, orr As Long, wf As Long, wo As Long
    n = PieceLine(sp, T0, T1, x0, y0, x1, y1, whole)
    ' проба — у середины самого длинного звена (не в вершине: у острого зубца точки по обе стороны
    ' от вершины обе оказались бы снаружи)
    im = -1: dm = -1
    For i = 0 To n - 2
        d = Sqr((WPx(i + 1) - WPx(i)) ^ 2 + (WPy(i + 1) - WPy(i)) ^ 2)
        tot = tot + d
        If d > dm Then
            dm = d: im = i
        End If
    Next i
    If tot < 0.000000001 Or dm < 0.000000000001 Then
        Exit Sub
    End If
    mx = (WPx(im) + WPx(im + 1)) / 2: my = (WPy(im) + WPy(im + 1)) / 2
    tx = (WPx(im + 1) - WPx(im)) / dm: ty = (WPy(im + 1) - WPy(im)) / dm
    wl = Winding(mx - ty * 0.0001, my + tx * 0.0001, sp, ol)
    wr = Winding(mx + ty * 0.0001, my - tx * 0.0001, sp, orr)
    If (wl <> 0) = (wr <> 0) Then
        Exit Sub
    End If
    fw = (wr <> 0)
    If fw Then
        wf = wr: wo = orr
    Else
        wf = wl: wo = ol
    End If
    ' вывернутая петля (ласточкин хвост сдвинутого контура, «ушко» обводки): залитое здесь — только
    ' от самого подконтура и против его хода; у дырки залитое — от внешнего контура
    If wf = wo Then
        If (wf > 0 And CLSig(sp) < 0) Or (wf < 0 And CLSig(sp) > 0) Then
            Exit Sub
        End If
    End If
    For i = 0 To n - 2
        ar = ar + (WPx(i) * WPy(i + 1) - WPx(i + 1) * WPy(i))
    Next i
    If NK > UBound(KSp) Then
        ReDim Preserve KSp(0 To 2 * NK): ReDim Preserve KT0(0 To 2 * NK): ReDim Preserve KT1(0 To 2 * NK)
        ReDim Preserve KFw(0 To 2 * NK): ReDim Preserve KWh(0 To 2 * NK)
        ReDim Preserve KSx(0 To 2 * NK): ReDim Preserve KSy(0 To 2 * NK): ReDim Preserve K2x(0 To 2 * NK): ReDim Preserve K2y(0 To 2 * NK)
        ReDim Preserve KPx(0 To 2 * NK): ReDim Preserve KPy(0 To 2 * NK): ReDim Preserve KEx(0 To 2 * NK): ReDim Preserve KEy(0 To 2 * NK)
        ReDim Preserve KAr(0 To 2 * NK): ReDim Preserve KLn(0 To 2 * NK)
    End If
    KSp(NK) = sp: KT0(NK) = T0: KT1(NK) = T1: KFw(NK) = fw: KWh(NK) = whole: KLn(NK) = tot
    If fw Then
        KSx(NK) = WPx(0): KSy(NK) = WPy(0): K2x(NK) = WPx(1): K2y(NK) = WPy(1)
        KPx(NK) = WPx(n - 2): KPy(NK) = WPy(n - 2): KEx(NK) = WPx(n - 1): KEy(NK) = WPy(n - 1)
        KAr(NK) = ar
    Else
        KSx(NK) = WPx(n - 1): KSy(NK) = WPy(n - 1): K2x(NK) = WPx(n - 2): K2y(NK) = WPy(n - 2)
        KPx(NK) = WPx(1): KPy(NK) = WPy(1): KEx(NK) = WPx(0): KEy(NK) = WPy(0)
        KAr(NK) = -ar
    End If
    NK = NK + 1
End Sub

' обмотка точки: луч вправо, отрезки из полосы; wo — вклад подконтура sp
Private Function Winding(ByVal px As Double, ByVal py As Double, ByVal sp As Long, wo As Long) As Long
    Dim r As Long, k As Long, e As Long, w As Long, x As Double, d As Long
    wo = 0
    r = Int((py - CLY0) / CLCell)
    If r < 0 Or r >= CLNRow Then
        Exit Function
    End If
    For k = CLRowSt(r) To CLRowSt(r + 1) - 1
        e = CLRowL(k)
        If (EAY(e) <= py And py < EBY(e)) Or (EBY(e) <= py And py < EAY(e)) Then
            x = EAX(e) + (py - EAY(e)) * (EBX(e) - EAX(e)) / (EBY(e) - EAY(e))
            If x > px Then
                If EBY(e) > EAY(e) Then
                    d = 1
                Else
                    d = -1
                End If
                w = w + d
                If ESp(e) = sp Then
                    wo = wo + d
                End If
            End If
        End If
    Next k
    Winding = w
End Function

Private Function RKey(ByVal a As Double, ByVal b As Double) As Boolean
    RKey = (Round(a, 5) = Round(b, 5))
End Function

Private Function Atan2C(ByVal y As Double, ByVal x As Double) As Double
    Const PI_ As Double = 3.14159265358979
    If x > 0 Then
        Atan2C = Atn(y / x)
    ElseIf x < 0 Then
        If y >= 0 Then
            Atan2C = Atn(y / x) + PI_
        Else
            Atan2C = Atn(y / x) - PI_
        End If
    ElseIf y > 0 Then
        Atan2C = PI_ / 2
    ElseIf y < 0 Then
        Atan2C = -PI_ / 2
    End If
End Function

Private Sub ChainPieces(ByVal minArea As Double, ByVal minWidth As Double)
    Const PI_ As Double = 3.14159265358979
    Dim used() As Boolean, lp() As Long, nl As Long, n As Long, cur As Long, c As Long, best As Long
    Dim din As Double, dout As Double, tr As Double, bestTr As Double
    If NK = 0 Then
        Exit Sub
    End If
    ReDim used(0 To NK): ReDim lp(0 To NK)
    ' целые подконтуры — каждый сам по себе
    For n = 0 To NK - 1
        If KWh(n) Then
            lp(0) = n: nl = 1
            EmitLoop lp, nl, minArea, minWidth
        End If
    Next n
    For n = 0 To NK - 1
        If Not KWh(n) And Not used(n) Then
            nl = 0: cur = n
            Do While cur >= 0
                If used(cur) Then
                    Exit Do
                End If
                used(cur) = True: lp(nl) = cur: nl = nl + 1
                best = -1
                For c = 0 To NK - 1
                    If Not KWh(c) And Not used(c) Then
                        If RKey(KSx(c), KEx(cur)) And RKey(KSy(c), KEy(cur)) Then
                            ' касание: самый правый поворот (залитое справа — обход вплотную)
                            din = Atan2C(KEy(cur) - KPy(cur), KEx(cur) - KPx(cur))
                            dout = Atan2C(K2y(c) - KEy(cur), K2x(c) - KEx(cur))
                            tr = dout - din + PI_
                            tr = tr - 2 * PI_ * Int(tr / (2 * PI_)) - PI_
                            If best < 0 Then
                                best = c: bestTr = tr
                            ElseIf tr < bestTr Then
                                best = c: bestTr = tr
                            End If
                        End If
                    End If
                Next c
                cur = best
            Loop
            EmitLoop lp, nl, minArea, minWidth
        End If
    Next n
End Sub

' контур из кусков: мелкий (площадь, толщина) — пропустить; иначе — сегменты исходных кривых
Private Sub EmitLoop(lp() As Long, ByVal nl As Long, ByVal minArea As Double, ByVal minWidth As Double)
    Dim i As Long, A As Double, per As Double, s0 As Long
    For i = 0 To nl - 1
        A = A + KAr(lp(i)): per = per + KLn(lp(i))
    Next i
    A = A + (KEx(lp(nl - 1)) * KSy(lp(0)) - KSx(lp(0)) * KEy(lp(nl - 1)))
    per = per + Sqr((KSx(lp(0)) - KEx(lp(nl - 1))) ^ 2 + (KSy(lp(0)) - KEy(lp(nl - 1))) ^ 2)
    A = Abs(A / 2)
    If A < minArea Then
        Exit Sub
    End If
    If per > 0 Then
        If 2 * A / per < minWidth Then
            Exit Sub
        End If
    End If
    s0 = NCO
    For i = 0 To nl - 1
        EmitPiece lp(i)
    Next i
    If NCO > s0 Then
        SnapLoop s0, NCO - s0
        If NCOP > UBound(COS0) Then
            ReDim Preserve COS0(0 To 2 * NCOP): ReDim Preserve CON(0 To 2 * NCOP)
        End If
        COS0(NCOP) = s0: CON(NCOP) = NCO - s0: NCOP = NCOP + 1
    End If
End Sub

' сегменты подконтура на отрезке параметра куска (в обратную сторону — с конца, развёрнутые)
Private Sub EmitPiece(ByVal k As Long)
    Dim sp As Long, NN As Long, T0 As Double, T1 As Double, g As Long, a As Double, b As Double, s0 As Long, n As Long
    Dim i As Long, j As Long, m As Long, tx As Double, ty As Double, tl As Boolean
    sp = KSp(k): NN = CPN(sp): T0 = KT0(k): T1 = KT1(k)
    If T1 <= T0 Then
        T1 = T1 + NN
    End If
    s0 = NCO
    g = Int(T0)
    Do While g < T1 - 0.000000000001
        a = T0 - g
        If a < 0 Then
            a = 0
        End If
        b = T1 - g
        If b > 1 Then
            b = 1
        End If
        If b - a > 0.000000000001 Then
            SplitSeg CPS0(sp) + (g Mod NN), a, b
        End If
        g = g + 1
    Loop
    If Not KFw(k) Then
        ' обратный ход: порядок сегментов и каждый сегмент — наоборот
        n = NCO - s0
        For i = 0 To n \ 2 - 1
            For m = 0 To 3
                tx = COX(4 * (s0 + i) + m): ty = COY(4 * (s0 + i) + m)
                COX(4 * (s0 + i) + m) = COX(4 * (s0 + n - 1 - i) + m): COY(4 * (s0 + i) + m) = COY(4 * (s0 + n - 1 - i) + m)
                COX(4 * (s0 + n - 1 - i) + m) = tx: COY(4 * (s0 + n - 1 - i) + m) = ty
            Next m
            tl = COL(s0 + i): COL(s0 + i) = COL(s0 + n - 1 - i): COL(s0 + n - 1 - i) = tl
        Next i
        For i = s0 To NCO - 1
            For m = 0 To 1
                j = 3 - m
                tx = COX(4 * i + m): ty = COY(4 * i + m)
                COX(4 * i + m) = COX(4 * i + j): COY(4 * i + m) = COY(4 * i + j)
                COX(4 * i + j) = tx: COY(4 * i + j) = ty
            Next m
        Next i
    End If
End Sub

' часть сегмента s от a до b (0 <= a < b <= 1) — в выход (де Кастельжо; прямая — по точкам)
Private Sub SplitSeg(ByVal s As Long, ByVal a As Double, ByVal b As Double)
    Dim p0x As Double, p0y As Double, c1x As Double, c1y As Double, c2x As Double, c2y As Double, p3x As Double, p3y As Double
    Dim tt As Double
    p0x = CSX(4 * s): p0y = CSY(4 * s): c1x = CSX(4 * s + 1): c1y = CSY(4 * s + 1)
    c2x = CSX(4 * s + 2): c2y = CSY(4 * s + 2): p3x = CSX(4 * s + 3): p3y = CSY(4 * s + 3)
    If CSL(s) Then
        PushO p0x + (p3x - p0x) * a, p0y + (p3y - p0y) * a, 0, 0, 0, 0, p0x + (p3x - p0x) * b, p0y + (p3y - p0y) * b, True
        Exit Sub
    End If
    If a > 0 Then
        CutAfter p0x, p0y, c1x, c1y, c2x, c2y, p3x, p3y, a
    End If
    If b < 1 Then
        If a < 1 Then
            tt = (b - a) / (1 - a)
        Else
            tt = 1
        End If
        CutBefore p0x, p0y, c1x, c1y, c2x, c2y, p3x, p3y, tt
    End If
    PushO p0x, p0y, c1x, c1y, c2x, c2y, p3x, p3y, False
End Sub

' кривая -> её часть [t, 1]
Private Sub CutAfter(p0x As Double, p0y As Double, c1x As Double, c1y As Double, c2x As Double, c2y As Double, _
                     p3x As Double, p3y As Double, ByVal t As Double)
    Dim ax As Double, ay As Double, bx As Double, by As Double, cx As Double, cy As Double
    Dim dx As Double, dy As Double, ex As Double, ey As Double, fx As Double, fy As Double
    ax = p0x + (c1x - p0x) * t: ay = p0y + (c1y - p0y) * t
    bx = c1x + (c2x - c1x) * t: by = c1y + (c2y - c1y) * t
    cx = c2x + (p3x - c2x) * t: cy = c2y + (p3y - c2y) * t
    dx = ax + (bx - ax) * t: dy = ay + (by - ay) * t
    ex = bx + (cx - bx) * t: ey = by + (cy - by) * t
    fx = dx + (ex - dx) * t: fy = dy + (ey - dy) * t
    p0x = fx: p0y = fy: c1x = ex: c1y = ey: c2x = cx: c2y = cy
End Sub

' кривая -> её часть [0, t]
Private Sub CutBefore(p0x As Double, p0y As Double, c1x As Double, c1y As Double, c2x As Double, c2y As Double, _
                      p3x As Double, p3y As Double, ByVal t As Double)
    Dim ax As Double, ay As Double, bx As Double, by As Double, cx As Double, cy As Double
    Dim dx As Double, dy As Double, ex As Double, ey As Double, fx As Double, fy As Double
    ax = p0x + (c1x - p0x) * t: ay = p0y + (c1y - p0y) * t
    bx = c1x + (c2x - c1x) * t: by = c1y + (c2y - c1y) * t
    cx = c2x + (p3x - c2x) * t: cy = c2y + (p3y - c2y) * t
    dx = ax + (bx - ax) * t: dy = ay + (by - ay) * t
    ex = bx + (cx - bx) * t: ey = by + (cy - by) * t
    fx = dx + (ex - dx) * t: fy = dy + (ey - dy) * t
    c1x = ax: c1y = ay: c2x = dx: c2y = dy: p3x = fx: p3y = fy
End Sub

Private Sub PushO(ByVal p0x As Double, ByVal p0y As Double, ByVal c1x As Double, ByVal c1y As Double, _
                  ByVal c2x As Double, ByVal c2y As Double, ByVal p3x As Double, ByVal p3y As Double, ByVal isL As Boolean)
    If 4 * NCO + 3 > UBound(COX) Then
        ReDim Preserve COX(0 To 8 * NCO + 16): ReDim Preserve COY(0 To 8 * NCO + 16)
    End If
    If NCO > UBound(COL) Then
        ReDim Preserve COL(0 To 2 * NCO + 4)
    End If
    If isL Then
        c1x = p0x + (p3x - p0x) / 3: c1y = p0y + (p3y - p0y) / 3
        c2x = p0x + 2 * (p3x - p0x) / 3: c2y = p0y + 2 * (p3y - p0y) / 3
    End If
    COX(4 * NCO) = p0x: COY(4 * NCO) = p0y: COX(4 * NCO + 1) = c1x: COY(4 * NCO + 1) = c1y
    COX(4 * NCO + 2) = c2x: COY(4 * NCO + 2) = c2y: COX(4 * NCO + 3) = p3x: COY(4 * NCO + 3) = p3y
    COL(NCO) = isL
    NCO = NCO + 1
End Sub

' стыки кусков: концы расходятся до допуска разбиения — оба в середину (с ручками)
Private Sub SnapLoop(ByVal s0 As Long, ByVal n As Long)
    Dim i As Long, a As Long, b As Long, mx As Double, my As Double, dax As Double, day As Double, dbx As Double, dby As Double
    For i = 0 To n - 1
        b = s0 + i
        If i = 0 Then
            a = s0 + n - 1
        Else
            a = b - 1
        End If
        If COX(4 * a + 3) <> COX(4 * b) Or COY(4 * a + 3) <> COY(4 * b) Then
            mx = (COX(4 * a + 3) + COX(4 * b)) / 2: my = (COY(4 * a + 3) + COY(4 * b)) / 2
            dax = mx - COX(4 * a + 3): day = my - COY(4 * a + 3)
            dbx = mx - COX(4 * b): dby = my - COY(4 * b)
            If COL(a) Then
                COX(4 * a + 3) = mx: COY(4 * a + 3) = my
                COX(4 * a + 1) = COX(4 * a) + (mx - COX(4 * a)) / 3: COY(4 * a + 1) = COY(4 * a) + (my - COY(4 * a)) / 3
                COX(4 * a + 2) = COX(4 * a) + 2 * (mx - COX(4 * a)) / 3: COY(4 * a + 2) = COY(4 * a) + 2 * (my - COY(4 * a)) / 3
            Else
                COX(4 * a + 2) = COX(4 * a + 2) + dax: COY(4 * a + 2) = COY(4 * a + 2) + day
                COX(4 * a + 3) = mx: COY(4 * a + 3) = my
            End If
            If COL(b) Then
                COX(4 * b) = mx: COY(4 * b) = my
                COX(4 * b + 1) = mx + (COX(4 * b + 3) - mx) / 3: COY(4 * b + 1) = my + (COY(4 * b + 3) - my) / 3
                COX(4 * b + 2) = mx + 2 * (COX(4 * b + 3) - mx) / 3: COY(4 * b + 2) = my + 2 * (COY(4 * b + 3) - my) / 3
            Else
                COX(4 * b) = mx: COY(4 * b) = my
                COX(4 * b + 1) = COX(4 * b + 1) + dbx: COY(4 * b + 1) = COY(4 * b + 1) + dby
            End If
        End If
    Next i
End Sub
'@@CLEANGEO_END

'@@CLEANTEST_BEGIN
' стенд: очистка по сегментам (тесты сверяют с proto/cleanup.py)
Public Function CleanTest(inX() As Double, inY() As Double, inL() As Boolean, ByVal nSeg As Long, _
                          spS0() As Long, spN() As Long, ByVal nSp As Long, ByVal minArea As Double, ByVal minWidth As Double) As Long
    Dim s As Long, k As Long
    ReDim CSX(0 To 4 * nSeg + 4): ReDim CSY(0 To 4 * nSeg + 4): ReDim CSL(0 To nSeg + 1)
    For s = 0 To nSeg - 1
        For k = 0 To 3
            CSX(4 * s + k) = inX(4 * s + k): CSY(4 * s + k) = inY(4 * s + k)
        Next k
        CSL(s) = inL(s)
    Next s
    NCS = nSeg
    ReDim CPS0(0 To nSp + 1): ReDim CPN(0 To nSp + 1)
    For s = 0 To nSp - 1
        CPS0(s) = spS0(s): CPN(s) = spN(s)
    Next s
    NCP = nSp
    CleanGeo minArea, minWidth
    CleanTest = NCOP
End Function
'@@CLEANTEST_END
'@@CLEAN_END

'@@NEST_BEGIN
' =====================================================================
'  NeonCenterline — раскладка (группа «Раскладка ▾»):
'    GroupSame  — одинаковые объекты (один контур с точностью SAME_TOL — как бы ни стояли узлы, без
'                 поворота и зеркала, и один цвет) — в группы «Одинаковые: N шт · Ш × В мм», ведомость форм;
'    NestSheets — детали на листы: по габаритам, без поворота, с полем от края и зазором,
'                 каждый цвет — на свои листы; длина листа (по горизонтали) — заданная или по раскладке.
'  Деталь — выделенный объект. Группа одинаковых и лист прежней раскладки — поштучно (их
'  элементы); одна выделенная группа — тоже (импорт SVG/DXF приходит одной группой).
'  Цвет детали — цвет заливки, без заливки — цвет обводки.
' =====================================================================

' ---------------------------------------------------------------------
'  Настройки раскладки (окно — frmNeonRun в режиме 3, NeonUI.AskNestParams)
' ---------------------------------------------------------------------
Public Function NeonNestSize() As String
    NeonNestSize = GetSetting("NeonCenterline", "Nest", "Size", "3050x2050")
End Function

Public Function NeonNestNum(ByVal key As String, ByVal def As Double) As Double
    Dim s As String
    s = GetSetting("NeonCenterline", "Nest", key, "")
    If Len(Trim$(s)) = 0 Then
        NeonNestNum = def
    Else
        NeonNestNum = ParseNum(s)
    End If
End Function

Public Function NeonNestFlag(ByVal key As String, ByVal def As Boolean) As Boolean
    NeonNestFlag = (GetSetting("NeonCenterline", "Nest", key, IIf(def, "1", "0")) = "1")
End Function

' «3050x2050», «3050 × 2050», «3050х2050» (русская х), «3050*2050» -> ширина и высота; False — не размер
Public Function NeonNestParse(ByVal s As String, w As Double, h As Double) As Boolean
    Dim a() As String, i As Long, n As Long, v(0 To 2) As Double
    s = LCase$(Trim$(s))
    s = Replace(s, ChrW(&HD7), " "): s = Replace(s, ChrW(&H445), " "): s = Replace(s, ChrW(&H425), " ")
    s = Replace(s, "x", " "): s = Replace(s, "*", " "): s = Replace(s, ";", " "): s = Replace(s, "/", " ")
    a = Split(s, " ")
    For i = 0 To UBound(a)
        If Len(a(i)) > 0 Then
            If n < 3 Then
                v(n) = ParseNum(a(i))
            End If
            n = n + 1
        End If
    Next i
    If n = 2 Then
        w = v(0): h = v(1)
        NeonNestParse = (w >= 10 And w <= 100000 And h >= 10 And h <= 100000)
    End If
End Function

' ---------------------------------------------------------------------
'  Кнопка «Одинаковые — в группы»
' ---------------------------------------------------------------------
Public Sub GroupSame()
    Dim col As Collection, old As ShapeRange, oldUnit As cdrUnit, grp As Boolean, errText As String, ttl As String
    Dim i As Long, c As Long, k As Long, t As Long, nGrp As Long, nIn As Long, nFail As Long, nSingle As Long
    Dim cnt() As Long, first() As Long, ord() As Long, sr As ShapeRange, g As Shape, allR As ShapeRange
    Dim msg As String, lst As String
    ttl = NeonL("same_title")
    If ActiveDocument Is Nothing Then
        MsgBox NeonL("err_nodoc"), vbExclamation, ttl
        Exit Sub
    End If
    If ActiveSelectionRange.Count = 0 Then
        MsgBox NeonL("same_nosel"), vbExclamation, ttl
        Exit Sub
    End If
    oldUnit = ActiveDocument.Unit
    ActiveDocument.Unit = cdrMillimeter
    ActiveDocument.BeginCommandGroup ttl
    grp = True
    Optimization = True
    EventsEnabled = False
    On Error GoTo Fail
    Set col = New Collection
    Set old = CreateShapeRange
    NtCollect ActiveSelectionRange, col, old, True
    NSm = 0: NSmCls = 0
    If col.Count = 0 Then
        GoTo Fin
    End If
    SmLoad col
    SameGeo SAME_TOL
    ReDim cnt(0 To NSmCls): ReDim first(0 To NSmCls)
    For i = 0 To NSm - 1
        c = SmCls(i)
        If cnt(c) = 0 Then
            first(c) = i
        End If
        cnt(c) = cnt(c) + 1
    Next i
    Set allR = CreateShapeRange
    For c = 0 To NSmCls - 1
        If cnt(c) >= 2 Then
            Set sr = CreateShapeRange
            For i = 0 To NSm - 1
                If SmCls(i) = c Then
                    sr.Add col(i + 1)
                End If
            Next i
            Set g = NtGroup(sr)
            If g Is Nothing Then
                nFail = nFail + 1
                allR.AddRange sr
            Else
                g.Name = NeonLF("same_group", cnt(c), NtMM(SmW(first(c))), NtMM(SmH(first(c))))
                NtTag g, "same"
                allR.Add g
                nGrp = nGrp + 1: nIn = nIn + cnt(c)
            End If
        Else
            allR.Add col(first(c) + 1)
            nSingle = nSingle + 1
        End If
    Next c
    On Error Resume Next
    allR.CreateSelection
    On Error GoTo Fail
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
        MsgBox NeonLF("err_fail", errText), vbCritical, ttl
        Exit Sub
    End If
    If NSm = 0 Then
        MsgBox NeonL("same_nosel"), vbExclamation, ttl
        Exit Sub
    End If
    If nGrp = 0 And nFail = 0 Then
        MsgBox NeonLF("same_none", NSm), vbInformation, ttl
        Exit Sub
    End If
    msg = NeonLF("same_report", NSm, NSmCls, nGrp, nIn, nSingle)
    If nFail > 0 Then
        msg = msg & vbCrLf & NeonLF("same_failed", nFail)
    End If
    ' ведомость: формы по убыванию числа штук (при равенстве — по порядку в выделении)
    ReDim ord(0 To NSmCls)
    For c = 0 To NSmCls - 1
        k = c
        Do While k > 0
            If cnt(ord(k - 1)) >= cnt(c) Then
                Exit Do
            End If
            ord(k) = ord(k - 1)
            k = k - 1
        Loop
        ord(k) = c
    Next c
    For t = 0 To NSmCls - 1
        c = ord(t)
        lst = lst & NeonLF("same_line", cnt(c), NtMM(SmW(first(c))), NtMM(SmH(first(c))), SmCn(first(c))) & vbCrLf
    Next t
    If NSmCls <= 25 Then
        msg = msg & vbCrLf & vbCrLf & NeonL("same_list") & vbCrLf & lst
    Else
        For t = 0 To 24
            c = ord(t)
            msg = msg & IIf(t = 0, vbCrLf & vbCrLf & NeonL("same_list") & vbCrLf, "") & _
                  NeonLF("same_line", cnt(c), NtMM(SmW(first(c))), NtMM(SmH(first(c))), SmCn(first(c))) & vbCrLf
        Next t
        msg = msg & NeonLF("same_more", NSmCls - 25) & vbCrLf
    End If
    If MsgBox(msg & vbCrLf & NeonL("len_copy"), vbYesNo + vbInformation, ttl) = vbYes Then
        CopyText NeonL("same_list") & vbCrLf & lst
    End If
    NeonToolbarCheck
    Exit Sub
Fail:
    errText = Err.Description
    Resume Fin
End Sub

' объекты для сравнения: цвет, габарит; контур — только у тех, у кого есть похожий по цвету и габариту
Private Sub SmLoad(col As Collection)
    Dim i As Long, sh As Shape, x As Double, y As Double, w As Double, h As Double
    NSm = col.Count: NSq = 0: NSg = 0
    ReDim SmCk(0 To NSm): ReDim SmCn(0 To NSm): ReDim SmW(0 To NSm): ReDim SmH(0 To NSm)
    ReDim SmQ0(0 To NSm): ReDim SmQn(0 To NSm)
    ReDim SqS0(0 To 63): ReDim SqN(0 To 63): ReDim SqC(0 To 63)
    ReDim SgX(0 To 1023): ReDim SgY(0 To 1023): ReDim SgL(0 To 255)
    For i = 0 To NSm - 1
        Set sh = col(i + 1)
        sh.GetBoundingBox x, y, w, h
        SmW(i) = w: SmH(i) = h
        SmCk(i) = NtColorKey(sh, SmCn(i))
    Next i
    For i = 0 To NSm - 1
        SmQ0(i) = NSq
        If SmHasTwin(i, SAME_TOL) Then
            SmReadGeo col(i + 1)
        End If
        SmQn(i) = NSq - SmQ0(i)
    Next i
End Sub

' подпути объекта (у группы — всех элементов) в SqS0/SqN/SqC и SgX/SgY/SgL
Private Sub SmReadGeo(sh As Shape)
    Dim s As Shape, crv As Curve, tmp As Shape, sp As SubPath, sg As Segment, s0 As Long
    Dim x0 As Double, y0 As Double, x1 As Double, y1 As Double, x2 As Double, y2 As Double, x3 As Double, y3 As Double
    If sh.Type = cdrGroupShape Then
        For Each s In sh.Shapes
            SmReadGeo s
        Next s
        Exit Sub
    End If
    Set crv = NtCurve(sh, tmp)
    If Not crv Is Nothing Then
        For Each sp In crv.SubPaths
            s0 = NSg
            For Each sg In sp.Segments
                x0 = sg.StartNode.PositionX: y0 = sg.StartNode.PositionY
                x3 = sg.EndNode.PositionX: y3 = sg.EndNode.PositionY
                If sg.Type = cdrCurveSegment Then
                    sg.GetStartingControlPointPosition x1, y1
                    sg.GetEndingControlPointPosition x2, y2
                    SmAddSg x0, y0, x1, y1, x2, y2, x3, y3, False
                Else
                    SmAddSg x0, y0, x0, y0, x3, y3, x3, y3, True
                End If
            Next sg
            If NSg > s0 Then
                If NSq > UBound(SqS0) Then
                    ReDim Preserve SqS0(0 To 2 * NSq): ReDim Preserve SqN(0 To 2 * NSq): ReDim Preserve SqC(0 To 2 * NSq)
                End If
                SqS0(NSq) = s0: SqN(NSq) = NSg - s0: SqC(NSq) = sp.Closed
                NSq = NSq + 1
            End If
        Next sp
    End If
    If Not tmp Is Nothing Then
        tmp.Delete
    End If
End Sub

Private Sub SmAddSg(ByVal x0 As Double, ByVal y0 As Double, ByVal x1 As Double, ByVal y1 As Double, _
                    ByVal x2 As Double, ByVal y2 As Double, ByVal x3 As Double, ByVal y3 As Double, ByVal isL As Boolean)
    If 4 * NSg + 3 > UBound(SgX) Then
        ReDim Preserve SgX(0 To 8 * NSg + 16): ReDim Preserve SgY(0 To 8 * NSg + 16)
    End If
    If NSg > UBound(SgL) Then
        ReDim Preserve SgL(0 To 2 * NSg + 4)
    End If
    SgX(4 * NSg) = x0: SgY(4 * NSg) = y0: SgX(4 * NSg + 1) = x1: SgY(4 * NSg + 1) = y1
    SgX(4 * NSg + 2) = x2: SgY(4 * NSg + 2) = y2: SgX(4 * NSg + 3) = x3: SgY(4 * NSg + 3) = y3
    SgL(NSg) = isL
    NSg = NSg + 1
End Sub

' ---------------------------------------------------------------------
'  Кнопка «Разбить на буквы» (группа «Контур ▾»)
' ---------------------------------------------------------------------
' кривая из многих контуров (текст в кривых, объединённые буквы) -> по объекту на букву, дырки — в
' своих буквах. Каждая буква — копия исходной кривой (заливка, обводка, слой те же) с своими
' контурами, исходная удаляется. Текст переводится в кривые сам
Public Sub SplitLetters()
    Dim col As Collection, oldUnit As cdrUnit, grp As Boolean, errText As String, ttl As String, msg As String
    Dim i As Long, g As Long, m As Long, q As Long, nSkip As Long, nObj As Long, nLet As Long, nKeep As Long
    Dim shp As Shape, d As Shape, crv As Curve, res As ShapeRange
    ttl = NeonL("split_title")
    If ActiveDocument Is Nothing Then
        MsgBox NeonL("err_nodoc"), vbExclamation, ttl
        Exit Sub
    End If
    If ActiveSelectionRange.Count = 0 Then
        MsgBox NeonL("err_nosel_any"), vbExclamation, ttl
        Exit Sub
    End If
    oldUnit = ActiveDocument.Unit
    ActiveDocument.Unit = cdrMillimeter
    ActiveDocument.BeginCommandGroup ttl
    grp = True
    Optimization = True
    EventsEnabled = False
    On Error GoTo Fail
    Set col = New Collection
    CollectClean ActiveSelectionRange, col, nSkip
    Set res = CreateShapeRange
    For i = 1 To col.Count
        Set shp = col(i)
        NSq = 0: NSg = 0
        ReDim SqS0(0 To 63): ReDim SqN(0 To 63): ReDim SqC(0 To 63)
        ReDim SgX(0 To 1023): ReDim SgY(0 To 1023): ReDim SgL(0 To 255)
        SmReadGeo shp
        LtNG = 0
        If NSq >= 2 Then
            LettersGeo
        End If
        If LtNG <= 1 Then
            res.Add shp
            nKeep = nKeep + 1
        Else
            For g = 0 To LtNG - 1
                Set crv = Application.CreateCurve(ActiveDocument)
                For m = LtG0(g) To LtG0(g) + LtGn(g) - 1
                    q = LtIdx(m)
                    AppendSubCO crv, SgX, SgY, SgL, SqS0(q), SqN(q), SqC(q)
                Next m
                Set d = shp.Duplicate(0, 0)
                d.Curve.CopyAssign crv
                res.Add d
            Next g
            shp.Delete
            nObj = nObj + 1: nLet = nLet + LtNG
        End If
    Next i
    On Error Resume Next
    res.CreateSelection
    On Error GoTo Fail
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
        MsgBox NeonLF("err_fail", errText), vbCritical, ttl
        Exit Sub
    End If
    If nObj = 0 Then
        msg = NeonL("split_none")
    Else
        msg = NeonLF("split_report", nObj, nLet, nKeep)
    End If
    If nSkip > 0 Then
        msg = msg & vbCrLf & vbCrLf & NeonLF("clean_skipped", nSkip)
    End If
    MsgBox msg, vbInformation, ttl
    NeonToolbarCheck
    Exit Sub
Fail:
    errText = Err.Description
    Resume Fin
End Sub

' ---------------------------------------------------------------------
'  Кнопка «Раскладка на листы»
' ---------------------------------------------------------------------
Public Sub NestSheets()
    Dim col As Collection, old As ShapeRange, oldUnit As cdrUnit, grp As Boolean, errText As String, ttl As String
    Dim sw As Double, shh As Double, mg As Double, gp As Double, byColor As Boolean, autoW As Boolean, wmin As Double
    Dim n As Long, i As Long, j As Long, k As Long, s As Long, nKey As Long, rows As Long
    Dim bx() As Double, by() As Double, bw() As Double, bh() As Double, ck() As String, cn() As String
    Dim keys() As String, knm() As String, krep() As Long, idx() As Long
    Dim minX As Double, maxX As Double, maxY As Double, x0 As Double, yTop As Double, cx As Double, cy As Double
    Dim sp As Double, lh As Double, area As Double, kArea As Double, kSheet As Double, cnt As Long, kCnt As Long
    Dim lay As Layer, fr As Shape, sh As Shape, sr As ShapeRange, g As Shape, allR As ShapeRange, vis As ShapeRange, lc As Color
    Dim nPlaced As Long, nUnfit As Long, nSheets As Long, colRep As String, wl As String, msg As String, wo As Object
    Dim fillMode As Long, labels As Boolean, lab As ShapeRange, whole As ShapeRange
    ttl = NeonL("nest_title")
    If ActiveDocument Is Nothing Then
        MsgBox NeonL("err_nodoc"), vbExclamation, ttl
        Exit Sub
    End If
    If ActiveSelectionRange.Count = 0 Then
        MsgBox NeonL("nest_nosel"), vbExclamation, ttl
        Exit Sub
    End If
    If Not AskNestParams() Then
        Exit Sub
    End If
    If Not NeonNestParse(NeonNestSize(), sw, shh) Then
        sw = 3050: shh = 2050
    End If
    mg = NeonNestNum("Margin", 10): gp = NeonNestNum("Gap", 10)
    byColor = NeonNestFlag("ByColor", True)
    autoW = NeonNestFlag("AutoW", False)
    wmin = NeonNestNum("WMin", 0)
    If wmin > sw Then
        wmin = sw
    End If
    fillMode = CLng(NeonNestNum("Fill", 0))
    labels = NeonNestFlag("Labels", True)
    oldUnit = ActiveDocument.Unit
    ActiveDocument.Unit = cdrMillimeter
    ActiveDocument.BeginCommandGroup ttl
    grp = True
    Optimization = True
    EventsEnabled = False
    On Error GoTo Fail
    Set col = New Collection
    Set old = CreateShapeRange
    NtCollect ActiveSelectionRange, col, old, True
    n = col.Count
    If n = 0 Then
        GoTo Fin
    End If
    ReDim bx(0 To n): ReDim by(0 To n): ReDim bw(0 To n): ReDim bh(0 To n): ReDim ck(0 To n): ReDim cn(0 To n)
    For i = 0 To n - 1
        Set sh = col(i + 1)
        sh.GetBoundingBox bx(i), by(i), bw(i), bh(i)
        ck(i) = NtColorKey(sh, cn(i))
        If Not byColor Then
            ck(i) = ""
        End If
        If i = 0 Or bx(i) < minX Then
            minX = bx(i)
        End If
        If i = 0 Or bx(i) + bw(i) > maxX Then
            maxX = bx(i) + bw(i)
        End If
        If i = 0 Or by(i) + bh(i) > maxY Then
            maxY = by(i) + bh(i)
        End If
    Next i
    ' рамки и подписи прежней раскладки из выделения — долой
    If old.Count > 0 Then
        old.Delete
    End If
    ' цвета — по порядку появления
    ReDim keys(0 To n): ReDim knm(0 To n): ReDim krep(0 To n)
    For i = 0 To n - 1
        For k = 0 To nKey - 1
            If keys(k) = ck(i) Then
                Exit For
            End If
        Next k
        If k = nKey Then
            keys(nKey) = ck(i): krep(nKey) = i
            If byColor Then
                knm(nKey) = cn(i)
            Else
                knm(nKey) = NeonL("all_colors")
            End If
            nKey = nKey + 1
        End If
    Next i
    Set lay = NtLayer(col(1))
    sp = 50
    If shh * 0.05 > sp Then
        sp = Int(shh * 0.05)
    End If
    lh = 8
    If shh * 0.025 > lh Then
        lh = shh * 0.025
    End If
    x0 = maxX + 2 * sp: yTop = maxY
    Set allR = CreateShapeRange
    Set vis = CreateShapeRange
    ReDim idx(0 To n)
    For k = 0 To nKey - 1
        NNt = 0
        ReDim NtW(0 To n): ReDim NtH(0 To n)
        For i = 0 To n - 1
            If ck(i) = keys(k) Then
                NtW(NNt) = bw(i): NtH(NNt) = bh(i): idx(NNt) = i
                NNt = NNt + 1
            End If
        Next i
        NestPack sw, shh, mg, gp, autoW, wmin, fillMode
        Set lc = Nothing
        If byColor Then
            Set lc = NtColorOf(col(krep(k) + 1))
        End If
        kArea = 0: kSheet = 0: kCnt = 0: wl = ""
        cx = x0: cy = yTop - 2 * lh
        For s = 0 To NtSheets - 1
            If s > 0 And s Mod NT_ROW = 0 Then
                cx = x0: cy = cy - (shh + 2 * lh + sp)
            End If
            Set fr = lay.CreateRectangle2(cx, cy - shh, NtSW(s), shh)
            NtFrameStyle fr
            fr.Name = NeonLF("nest_frame", s + 1)
            NtTag fr, "frame"
            vis.Add fr
            Set sr = CreateShapeRange
            cnt = 0: area = 0
            For j = 0 To NNt - 1
                If NtS(j) = s Then
                    i = idx(j)
                    Set sh = col(i + 1)
                    sh.Move cx + NtX(j) - bx(i), (cy - NtY(j)) - (by(i) + bh(i))
                    sr.Add sh
                    cnt = cnt + 1: area = area + bw(i) * bh(i)
                End If
            Next j
            Set lab = CreateShapeRange
            If labels Then
                NtLabel lay, lab, cx, cy + lh * 0.5, lh, NeonLF("nest_label", s + 1, NtSheets, NtMM(NtSW(s)), NtMM(shh), knm(k), cnt, _
                        CLng(100 * area / (NtSW(s) * shh))), lc, NeonLF("nest_labname", s + 1)
                vis.AddRange lab
            End If
            ' лист — одной группой: рамка, подпись, детали; не вышло (детали на другом слое) — только детали
            Set g = Nothing
            If sr.Count > 0 Then
                Set whole = CreateShapeRange
                whole.AddRange sr
                whole.Add fr
                whole.AddRange lab
                Set g = NtGroup(whole)
                If g Is Nothing And sr.Count > 1 Then
                    Set g = NtGroup(sr)
                End If
            End If
            If g Is Nothing Then
                allR.AddRange sr
            Else
                g.Name = NeonLF("nest_group", s + 1, NtSheets, knm(k))
                NtTag g, "sheet"
                allR.Add g
            End If
            kArea = kArea + area: kSheet = kSheet + NtSW(s) * shh: kCnt = kCnt + cnt
            wl = wl & IIf(s > 0, ", ", "") & NtMM(NtSW(s))
            cx = cx + NtSW(s) + sp
        Next s
        For j = 0 To NNt - 1
            If NtS(j) < 0 Then
                nUnfit = nUnfit + 1
            End If
        Next j
        If NtSheets > 0 Then
            rows = (NtSheets + NT_ROW - 1) \ NT_ROW
            yTop = yTop - rows * (shh + 2 * lh + sp) - sp
            colRep = colRep & "  " & NeonLF("nest_rep_color", knm(k), NtSheets, kCnt, CLng(100 * kArea / kSheet))
            If autoW Then
                colRep = colRep & NeonLF("nest_rep_widths", wl)
            End If
            colRep = colRep & vbCrLf
        End If
        nPlaced = nPlaced + kCnt: nSheets = nSheets + NtSheets
    Next k
    On Error Resume Next
    allR.CreateSelection
    vis.AddRange allR
    Set wo = ActiveWindow
    wo.ActiveView.ToFitShapeRange vis
    On Error GoTo Fail
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
        MsgBox NeonLF("err_fail", errText), vbCritical, ttl
        Exit Sub
    End If
    If n = 0 Then
        MsgBox NeonL("nest_nosel"), vbExclamation, ttl
        Exit Sub
    End If
    If nPlaced = 0 Then
        MsgBox NeonL("nest_none"), vbExclamation, ttl
        Exit Sub
    End If
    If autoW Then
        msg = NeonLF("nest_rep_auto", NtMM(shh), NtMM(wmin), NtMM(sw))
    Else
        msg = NeonLF("nest_rep_sheet", NtMM(sw), NtMM(shh))
    End If
    msg = NeonLF("nest_report", nPlaced, nSheets, msg, FmtNum(mg), FmtNum(gp)) & vbCrLf & vbCrLf & colRep
    If nUnfit > 0 Then
        msg = msg & vbCrLf & NeonLF("nest_unfit", nUnfit) & vbCrLf
    End If
    msg = msg & vbCrLf & NeonL("nest_rep_tail") & vbCrLf & vbCrLf & NeonL("nest_beta")
    If MsgBox(msg & vbCrLf & vbCrLf & NeonL("len_copy"), vbYesNo + vbInformation, ttl) = vbYes Then
        CopyText msg
    End If
    NeonToolbarCheck
    Exit Sub
Fail:
    errText = Err.Description
    Resume Fin
End Sub

' рамка листа: без заливки, тонкая серая обводка
Private Sub NtFrameStyle(fr As Shape)
    On Error Resume Next
    fr.Fill.ApplyNoFill
    fr.Outline.Width = 0.5
    fr.Outline.Color.RGBAssign 128, 128, 128
End Sub

' подпись листа над рамкой: образец цвета (если есть) и текст высотой ~lh мм — в vis
Private Sub NtLabel(lay As Layer, vis As ShapeRange, ByVal x As Double, ByVal y As Double, ByVal lh As Double, _
                    ByVal txt As String, c As Color, ByVal nm As String)
    Dim t As Shape, r As Shape, pt As Double
    On Error Resume Next
    If Not c Is Nothing Then
        Set r = lay.CreateRectangle2(x, y, lh, lh)
        r.Fill.ApplyUniformFill c
        r.Outline.Width = 0.2
        r.Outline.Color.RGBAssign 128, 128, 128
        r.Name = nm
        NtTag r, "label"
        vis.Add r
        x = x + lh * 1.5
    End If
    ' кегль: заглавная ~0.7 кегля, 1 пт = 0.3528 мм
    pt = lh / 0.7 / 0.3528
    If pt > 2000 Then
        pt = 2000
    End If
    If pt < 4 Then
        pt = 4
    End If
    Set t = lay.CreateArtisticText(x, y, txt, , , , pt)
    If Not t Is Nothing Then
        t.Fill.UniformColor.RGBAssign 90, 90, 90
        t.Name = nm
        NtTag t, "label"
        vis.Add t
    End If
End Sub

' ---------------------------------------------------------------------
'  Общее: детали выделения, цвет, кривая, группы и метки
' ---------------------------------------------------------------------
' детали выделения в col; рамки и подписи прежней раскладки — в old. Группа одинаковых и лист
' раскладки — поштучно; одна выделенная группа (top) — тоже: импорт SVG/DXF приходит одной группой
Private Sub NtCollect(sr As ShapeRange, col As Collection, old As ShapeRange, ByVal top As Boolean)
    Dim sh As Shape
    If top And sr.Count = 1 Then
        Set sh = sr(1)
        If sh.Type = cdrGroupShape Then
            If Not NtIsLayout(sh) And Len(NtTagOf(sh)) = 0 Then
                NtCollect NtUngroup(sh), col, old, True
                Exit Sub
            End If
        End If
    End If
    For Each sh In sr
        If NtIsLayout(sh) Then
            old.Add sh
        ElseIf sh.Type = cdrGroupShape Then
            If Len(NtTagOf(sh)) > 0 Then
                NtCollect NtUngroup(sh), col, old, False
            Else
                col.Add sh
            End If
        Else
            col.Add sh
        End If
    Next sh
End Sub

' разгруппировать один уровень: элементы группы
Private Function NtUngroup(sh As Shape) As ShapeRange
    Dim r As ShapeRange, o As Object
    On Error Resume Next
    Set o = sh
    Set r = o.UngroupEx
    If r Is Nothing Then
        Err.Clear
        Set r = sh.Shapes.All
        sh.Ungroup
    End If
    If r Is Nothing Then
        Set r = CreateShapeRange
    End If
    Set NtUngroup = r
End Function

Private Function NtGroup(sr As ShapeRange) As Shape
    On Error Resume Next
    Set NtGroup = sr.Group
    Err.Clear
End Function

' метка объекта плагина: "same" — группа одинаковых, "sheet" — лист раскладки, "frame" — рамка листа,
' "label" — подпись листа. Без свойств (старый CorelDRAW) — по началу имени
Private Sub NtTag(sh As Shape, ByVal tg As String)
    Dim o As Object
    On Error Resume Next
    Set o = sh
    o.Properties("NeonCenterline", 1) = tg
    Err.Clear
End Sub

Private Function NtTagOf(sh As Shape) As String
    Dim v As Variant, nm As String, o As Object
    On Error Resume Next
    Set o = sh
    v = Empty
    v = o.Properties("NeonCenterline", 1)
    If Err.Number = 0 And Not IsEmpty(v) And Not IsNull(v) Then
        NtTagOf = CStr(v)
    End If
    Err.Clear
    If Len(NtTagOf) = 0 Then
        nm = sh.Name
        If NtStarts(nm, "same_group") Then
            NtTagOf = "same"
        ElseIf NtStarts(nm, "nest_group") Then
            NtTagOf = "sheet"
        ElseIf NtStarts(nm, "nest_frame") Then
            NtTagOf = "frame"
        ElseIf NtStarts(nm, "nest_labname") Then
            NtTagOf = "label"
        End If
    End If
End Function

' имя начинается с постоянной части строки key (до первой подстановки)
Private Function NtStarts(ByVal nm As String, ByVal key As String) As Boolean
    Dim p As String, q As Long
    p = NeonL(key)
    q = InStr(p, "{")
    If q > 1 Then
        p = Left$(p, q - 1)
        NtStarts = (Left$(nm, Len(p)) = p)
    End If
End Function

' рамка или подпись прежней раскладки (или объект со слоя «Раскладка» первых сборок)
Private Function NtIsLayout(sh As Shape) As Boolean
    Dim tg As String
    tg = NtTagOf(sh)
    If tg = "frame" Or tg = "label" Then
        NtIsLayout = True
        Exit Function
    End If
    On Error Resume Next
    NtIsLayout = (sh.Layer.Name = NeonL("nest_layer"))
End Function

' слой для рамок и подписей: слой деталей (обычно — текущий), иначе текущий слой документа
Private Function NtLayer(part As Shape) As Layer
    Dim ly As Layer
    On Error Resume Next
    Set ly = part.Layer
    If ly Is Nothing Then
        Set ly = ActiveDocument.ActiveLayer
    End If
    Set NtLayer = ly
End Function

' цвет детали: заливка, без неё — обводка; у группы — первого элемента с цветом
Private Function NtColorOf(sh As Shape) As Color
    Dim s As Shape, c As Color
    On Error Resume Next
    If sh.Type = cdrGroupShape Then
        For Each s In sh.Shapes
            Set c = NtColorOf(s)
            If Not c Is Nothing Then
                Set NtColorOf = c
                Exit Function
            End If
        Next s
        Exit Function
    End If
    If sh.Fill.Type = cdrUniformFill Then
        Set c = sh.Fill.UniformColor
    ElseIf sh.Outline.Type <> cdrNoOutline Then
        Set c = sh.Outline.Color
    End If
    Set NtColorOf = c
End Function

' ключ цвета (для сравнения) и название (для подписей)
Private Function NtColorKey(sh As Shape, nm As String) As String
    Dim c As Color
    Set c = NtColorOf(sh)
    If c Is Nothing Then
        nm = NeonL("no_color")
    Else
        On Error Resume Next
        NtColorKey = c.ToString
        nm = NtColorName(c)
    End If
End Function

Private Function NtColorName(c As Color) As String
    Dim s As String
    On Error Resume Next
    Select Case c.Type
    Case cdrColorRGB
        s = "#" & Right$("0" & Hex$(c.RGBRed), 2) & Right$("0" & Hex$(c.RGBGreen), 2) & Right$("0" & Hex$(c.RGBBlue), 2)
    Case cdrColorCMYK
        s = "C" & c.CMYKCyan & " M" & c.CMYKMagenta & " Y" & c.CMYKYellow & " K" & c.CMYKBlack
    End Select
    ' у CMYK и RGB — только компоненты (имя у своих цветов — «цвет без названия»), у прочих — имя
    If Len(s) = 0 Then
        s = c.Name
        If Len(s) = 0 Then
            s = c.ToString
        End If
    End If
    NtColorName = s
End Function

' кривая объекта: своя, кривая отображения, иначе — временная копия в кривых (tmp — удалить после)
Private Function NtCurve(sh As Shape, tmp As Shape) As Curve
    Dim crv As Curve
    On Error Resume Next
    Set tmp = Nothing
    If sh.Type = cdrCurveShape Then
        Set crv = sh.Curve
    Else
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
    Set NtCurve = crv
End Function

' мм для подписей: целые, у мелких — с десятыми
Private Function NtMM(ByVal v As Double) As String
    If v >= 100 Then
        NtMM = FmtNum(Round(v, 0))
    Else
        NtMM = FmtNum(Round(v, 1))
    End If
End Function

'@@SAMEGEO_BEGIN
' ---------------------------------------------------------------------
'  Одинаковые формы: геометрия (повторяет proto/nesting.py один в один).
'  Сравнивается сам контур, а не узлы: точки по контуру через SAME_STEP (у кривой — равные доли
'  параметра), сдвинутые к левому нижнему углу габарита точек. Одинаковые: один цвет, габарит и
'  длина контура почти равны, каждая точка одного контура — не дальше tol от другого (в обе
'  стороны). Число узлов, их расстановка, начало и направление обхода значения не имеют.
'  Вход: объекты SmCk, подпути SmQ0/SmQn (SmQn = 0 — контур не читался: похожих объектов нет),
'  подпути SqS0/SqN/SqC, сегменты SgX/SgY (4 точки на сегмент), SgL (прямая).
'  Выход: SmCls (класс по порядку появления), NSmCls.
' ---------------------------------------------------------------------
Public Sub SameGeo(ByVal tol As Double)
    Dim i As Long, k As Long, c As Long, reps() As Long
    NSmP = 0
    ReDim SmX(0 To 1023): ReDim SmY(0 To 1023)
    ReDim QP0(0 To NSq): ReDim QPn(0 To NSq)
    ReDim SmSW(0 To NSm): ReDim SmSH(0 To NSm): ReDim SmLen(0 To NSm)
    For i = 0 To NSm - 1
        SmSample i
    Next i
    ReDim SmCls(0 To NSm): ReDim reps(0 To NSm)
    NSmCls = 0
    For i = 0 To NSm - 1
        c = -1
        For k = 0 To NSmCls - 1
            If SmEq(reps(k), i, tol) Then
                c = k
                Exit For
            End If
        Next k
        If c < 0 Then
            c = NSmCls
            reps(NSmCls) = i
            NSmCls = NSmCls + 1
        End If
        SmCls(i) = c
    Next i
End Sub

Private Function SmDist(ByVal ax As Double, ByVal ay As Double, ByVal bx As Double, ByVal by As Double) As Double
    SmDist = Sqr((bx - ax) * (bx - ax) + (by - ay) * (by - ay))
End Function

' точки контура объекта o, габарит и длина по ним
Private Sub SmSample(ByVal o As Long)
    Dim q As Long, s As Long, i As Long, n As Long, lc As Double, stp As Double, ln As Double, p0 As Long
    Dim t As Double, mt As Double, a As Double, b As Double, c As Double, d As Double, mx As Double, my As Double
    If SmQn(o) <= 0 Then
        Exit Sub
    End If
    lc = 0
    For q = SmQ0(o) To SmQ0(o) + SmQn(o) - 1
        For s = SqS0(q) To SqS0(q) + SqN(q) - 1
            If SgL(s) Then
                lc = lc + SmDist(SgX(4 * s), SgY(4 * s), SgX(4 * s + 3), SgY(4 * s + 3))
            Else
                lc = lc + SmDist(SgX(4 * s), SgY(4 * s), SgX(4 * s + 1), SgY(4 * s + 1)) + _
                     SmDist(SgX(4 * s + 1), SgY(4 * s + 1), SgX(4 * s + 2), SgY(4 * s + 2)) + _
                     SmDist(SgX(4 * s + 2), SgY(4 * s + 2), SgX(4 * s + 3), SgY(4 * s + 3))
            End If
        Next s
    Next q
    stp = SAME_STEP
    If lc / SAME_MAXPTS > stp Then
        stp = lc / SAME_MAXPTS
    End If
    p0 = NSmP
    For q = SmQ0(o) To SmQ0(o) + SmQn(o) - 1
        QP0(q) = NSmP
        For s = SqS0(q) To SqS0(q) + SqN(q) - 1
            If SgL(s) Then
                n = -Int(-(SmDist(SgX(4 * s), SgY(4 * s), SgX(4 * s + 3), SgY(4 * s + 3)) / stp))
                If n < 1 Then
                    n = 1
                End If
                For i = 0 To n - 1
                    SmAddPt SgX(4 * s) + (SgX(4 * s + 3) - SgX(4 * s)) * i / n, SgY(4 * s) + (SgY(4 * s + 3) - SgY(4 * s)) * i / n
                Next i
            Else
                ln = SmDist(SgX(4 * s), SgY(4 * s), SgX(4 * s + 1), SgY(4 * s + 1)) + _
                     SmDist(SgX(4 * s + 1), SgY(4 * s + 1), SgX(4 * s + 2), SgY(4 * s + 2)) + _
                     SmDist(SgX(4 * s + 2), SgY(4 * s + 2), SgX(4 * s + 3), SgY(4 * s + 3))
                n = -Int(-(ln / stp))
                If n < 2 Then
                    n = 2
                End If
                For i = 0 To n - 1
                    t = i / n
                    mt = 1 - t
                    a = mt * mt * mt
                    b = 3 * mt * mt * t
                    c = 3 * mt * t * t
                    d = t * t * t
                    SmAddPt a * SgX(4 * s) + b * SgX(4 * s + 1) + c * SgX(4 * s + 2) + d * SgX(4 * s + 3), _
                            a * SgY(4 * s) + b * SgY(4 * s + 1) + c * SgY(4 * s + 2) + d * SgY(4 * s + 3)
                Next i
            End If
        Next s
        If Not SqC(q) And SqN(q) > 0 Then
            s = SqS0(q) + SqN(q) - 1
            SmAddPt SgX(4 * s + 3), SgY(4 * s + 3)
        End If
        QPn(q) = NSmP - QP0(q)
    Next q
    If NSmP = p0 Then
        Exit Sub
    End If
    mx = SmX(p0): my = SmY(p0)
    For i = p0 + 1 To NSmP - 1
        If SmX(i) < mx Then
            mx = SmX(i)
        End If
        If SmY(i) < my Then
            my = SmY(i)
        End If
    Next i
    For i = p0 To NSmP - 1
        SmX(i) = SmX(i) - mx
        SmY(i) = SmY(i) - my
        If SmX(i) > SmSW(o) Then
            SmSW(o) = SmX(i)
        End If
        If SmY(i) > SmSH(o) Then
            SmSH(o) = SmY(i)
        End If
    Next i
    ln = 0
    For q = SmQ0(o) To SmQ0(o) + SmQn(o) - 1
        For i = QP0(q) To QP0(q) + QPn(q) - 2
            ln = ln + SmDist(SmX(i), SmY(i), SmX(i + 1), SmY(i + 1))
        Next i
        If SqC(q) And QPn(q) > 1 Then
            ln = ln + SmDist(SmX(QP0(q) + QPn(q) - 1), SmY(QP0(q) + QPn(q) - 1), SmX(QP0(q)), SmY(QP0(q)))
        End If
    Next q
    SmLen(o) = ln
End Sub

Private Sub SmAddPt(ByVal x As Double, ByVal y As Double)
    If NSmP > UBound(SmX) Then
        ReDim Preserve SmX(0 To 2 * NSmP): ReDim Preserve SmY(0 To 2 * NSmP)
    End If
    SmX(NSmP) = x: SmY(NSmP) = y
    NSmP = NSmP + 1
End Sub

' объекты a и b одинаковые
Private Function SmEq(ByVal a As Long, ByVal b As Long, ByVal tol As Double) As Boolean
    Dim m As Double
    If SmQn(a) <= 0 Or SmQn(b) <= 0 Then
        Exit Function
    End If
    If SmCk(a) <> SmCk(b) Or Abs(SmSW(a) - SmSW(b)) > tol Or Abs(SmSH(a) - SmSH(b)) > tol Then
        Exit Function
    End If
    m = SmLen(a)
    If SmLen(b) > m Then
        m = SmLen(b)
    End If
    If Abs(SmLen(a) - SmLen(b)) > 0.01 * m + 2 * tol Then
        Exit Function
    End If
    If Not SmNear(a, b, tol) Then
        Exit Function
    End If
    SmEq = SmNear(b, a, tol)
End Function

' каждая точка контура a — не дальше tol от контура b (сетка по отрезкам b)
Private Function SmNear(ByVal a As Long, ByVal b As Long, ByVal tol As Double) As Boolean
    Dim q As Long, i As Long, k As Long, m As Long, nx As Long, ny As Long, nc As Long, ix As Long, iy As Long
    Dim ix0 As Long, ix1 As Long, iy0 As Long, iy1 As Long, cs As Double, v As Double, t2 As Double
    Dim ax As Double, ay As Double, bx As Double, by As Double, px As Double, py As Double
    Dim dx As Double, dy As Double, l2 As Double, t As Double, ex As Double, ey As Double, ok As Boolean
    Dim pos() As Long
    NGr = 0
    ReDim GrA(0 To 255): ReDim GrB(0 To 255)
    For q = SmQ0(b) To SmQ0(b) + SmQn(b) - 1
        For i = QP0(q) To QP0(q) + QPn(q) - 2
            SmAddGr i, i + 1
        Next i
        If SqC(q) And QPn(q) > 1 Then
            SmAddGr QP0(q) + QPn(q) - 1, QP0(q)
        End If
    Next q
    If NGr = 0 Then
        Exit Function
    End If
    cs = 4 * tol
    v = Sqr((SmSW(b) + 2 * tol) * (SmSH(b) + 2 * tol) / NGr)
    If v > cs Then
        cs = v
    End If
    nx = Int((SmSW(b) + 2 * tol) / cs) + 1
    ny = Int((SmSH(b) + 2 * tol) / cs) + 1
    nc = nx * ny
    ' сетка от (-tol, -tol): отрезок с запасом tol — в ячейки своего габарита (подсчёт, затем раскладка)
    ReDim GrCnt(0 To nc): ReDim GrSt(0 To nc + 1)
    For m = 0 To 1
        For k = 0 To NGr - 1
            ax = SmX(GrA(k)): ay = SmY(GrA(k)): bx = SmX(GrB(k)): by = SmY(GrB(k))
            If ax < bx Then
                ix0 = Int(ax / cs): ix1 = Int((bx + 2 * tol) / cs)
            Else
                ix0 = Int(bx / cs): ix1 = Int((ax + 2 * tol) / cs)
            End If
            If ay < by Then
                iy0 = Int(ay / cs): iy1 = Int((by + 2 * tol) / cs)
            Else
                iy0 = Int(by / cs): iy1 = Int((ay + 2 * tol) / cs)
            End If
            If ix0 < 0 Then
                ix0 = 0
            End If
            If iy0 < 0 Then
                iy0 = 0
            End If
            If ix1 > nx - 1 Then
                ix1 = nx - 1
            End If
            If iy1 > ny - 1 Then
                iy1 = ny - 1
            End If
            For iy = iy0 To iy1
                For ix = ix0 To ix1
                    If m = 0 Then
                        GrCnt(iy * nx + ix) = GrCnt(iy * nx + ix) + 1
                    Else
                        GrIdx(pos(iy * nx + ix)) = k
                        pos(iy * nx + ix) = pos(iy * nx + ix) + 1
                    End If
                Next ix
            Next iy
        Next k
        If m = 0 Then
            GrSt(0) = 0
            For i = 0 To nc - 1
                GrSt(i + 1) = GrSt(i) + GrCnt(i)
            Next i
            ReDim GrIdx(0 To GrSt(nc) + 1)
            ReDim pos(0 To nc)
            For i = 0 To nc - 1
                pos(i) = GrSt(i)
            Next i
        End If
    Next m
    t2 = tol * tol
    For q = SmQ0(a) To SmQ0(a) + SmQn(a) - 1
        For i = QP0(q) To QP0(q) + QPn(q) - 1
            px = SmX(i): py = SmY(i)
            ix = Int((px + tol) / cs)
            iy = Int((py + tol) / cs)
            If ix < 0 Or ix >= nx Or iy < 0 Or iy >= ny Then
                Exit Function
            End If
            ok = False
            For m = GrSt(iy * nx + ix) To GrSt(iy * nx + ix + 1) - 1
                k = GrIdx(m)
                ax = SmX(GrA(k)): ay = SmY(GrA(k)): bx = SmX(GrB(k)): by = SmY(GrB(k))
                dx = bx - ax: dy = by - ay
                l2 = dx * dx + dy * dy
                t = 0
                If l2 > 0 Then
                    t = ((px - ax) * dx + (py - ay) * dy) / l2
                    If t < 0 Then
                        t = 0
                    ElseIf t > 1 Then
                        t = 1
                    End If
                End If
                ex = ax + t * dx - px
                ey = ay + t * dy - py
                If ex * ex + ey * ey <= t2 Then
                    ok = True
                    Exit For
                End If
            Next m
            If Not ok Then
                Exit Function
            End If
        Next i
    Next q
    SmNear = True
End Function

Private Sub SmAddGr(ByVal i As Long, ByVal j As Long)
    If NGr > UBound(GrA) Then
        ReDim Preserve GrA(0 To 2 * NGr): ReDim Preserve GrB(0 To 2 * NGr)
    End If
    GrA(NGr) = i: GrB(NGr) = j
    NGr = NGr + 1
End Sub

' похожие по цвету и габариту CorelDRAW (без контура): есть ли у объекта o пара
Public Function SmHasTwin(ByVal o As Long, ByVal tol As Double) As Boolean
    Dim j As Long
    For j = 0 To NSm - 1
        If j <> o Then
            If SmCk(j) = SmCk(o) Then
                If Abs(SmW(j) - SmW(o)) <= 2 * tol And Abs(SmH(j) - SmH(o)) <= 2 * tol Then
                    SmHasTwin = True
                    Exit Function
                End If
            End If
        End If
    Next j
End Function
'@@SAMEGEO_END

'@@LETTERSGEO_BEGIN
' ---------------------------------------------------------------------
'  Разбить на буквы: геометрия (повторяет proto/letters.py один в один).
'  Вход: подпути одного объекта SqS0/SqN/SqC (NSq), сегменты SgX/SgY/SgL. Контуры — точки, как у
'  «Одинаковых» (SmSample). У замкнутого контура — точка на нём (середина самого длинного отрезка)
'  и сколько других контуров её окружают: чётное число — наружный контур буквы, нечётное — дырка;
'  дырка — к самому маленькому из окружающих. Незамкнутые — каждый отдельно. Точки и акценты — к
'  своей букве (LtAttach).
'  Выход: LtNG букв, у буквы — LtGn подпутей с LtG0 в LtIdx (наружный первым); буквы слева направо.
' ---------------------------------------------------------------------
Public Sub LettersGeo()
    Dim k As Long, i As Long, j As Long, n As Long, s As Long, m As Long, best As Double, l As Double
    Dim ax As Double, ay As Double, bx As Double, by As Double, a As Double, has() As Boolean, ng As Long
    Dim cl() As Boolean, bx0() As Double, by0() As Double, bx1() As Double, by1() As Double
    Dim prx() As Double, pry() As Double, ar() As Double, dep() As Long, par() As Long, used() As Boolean
    Dim g0() As Long, gn() As Long, ord() As Long, t As Long
    n = NSq
    NSm = 1
    ReDim SmQ0(0 To 1): ReDim SmQn(0 To 1)
    SmQ0(0) = 0: SmQn(0) = n
    NSmP = 0
    ReDim SmX(0 To 1023): ReDim SmY(0 To 1023)
    ReDim QP0(0 To n): ReDim QPn(0 To n)
    ReDim SmSW(0 To 1): ReDim SmSH(0 To 1): ReDim SmLen(0 To 1)
    SmSample 0
    ReDim cl(0 To n): ReDim has(0 To n): ReDim bx0(0 To n): ReDim by0(0 To n): ReDim bx1(0 To n): ReDim by1(0 To n)
    ReDim prx(0 To n): ReDim pry(0 To n): ReDim ar(0 To n): ReDim dep(0 To n): ReDim par(0 To n)
    For k = 0 To n - 1
        s = SqS0(k) + SqN(k) - 1
        cl(k) = SqC(k)
        If Not cl(k) Then
            cl(k) = (SmDist(SgX(4 * SqS0(k)), SgY(4 * SqS0(k)), SgX(4 * s + 3), SgY(4 * s + 3)) < LT_NEAR)
        End If
        par(k) = -1
        If QPn(k) > 0 Then
            bx0(k) = SmX(QP0(k)): by0(k) = SmY(QP0(k)): bx1(k) = bx0(k): by1(k) = by0(k)
            For i = QP0(k) + 1 To QP0(k) + QPn(k) - 1
                If SmX(i) < bx0(k) Then
                    bx0(k) = SmX(i)
                End If
                If SmY(i) < by0(k) Then
                    by0(k) = SmY(i)
                End If
                If SmX(i) > bx1(k) Then
                    bx1(k) = SmX(i)
                End If
                If SmY(i) > by1(k) Then
                    by1(k) = SmY(i)
                End If
            Next i
        End If
        If cl(k) And QPn(k) >= 3 Then
            has(k) = True
            best = -1
            a = 0
            For m = 0 To QPn(k) - 1
                i = QP0(k) + m
                j = QP0(k) + (m + 1) Mod QPn(k)
                ax = SmX(i): ay = SmY(i): bx = SmX(j): by = SmY(j)
                l = SmDist(ax, ay, bx, by)
                If l > best Then
                    best = l
                    prx(k) = (ax + bx) / 2: pry(k) = (ay + by) / 2
                End If
                a = a + (ax * by - bx * ay)
            Next m
            ar(k) = Abs(a / 2)
        End If
    Next k
    For i = 0 To n - 1
        If has(i) Then
            For j = 0 To n - 1
                If j <> i And has(j) Then
                    If Not (bx0(i) < bx0(j) - 0.000000001 Or by0(i) < by0(j) - 0.000000001 Or _
                            bx1(i) > bx1(j) + 0.000000001 Or by1(i) > by1(j) + 0.000000001) Then
                        If LtInside(prx(i), pry(i), j) Then
                            dep(i) = dep(i) + 1
                            If par(i) < 0 Then
                                par(i) = j
                            ElseIf ar(j) < ar(par(i)) Then
                                par(i) = j
                            End If
                        End If
                    End If
                End If
            Next j
        End If
    Next i
    ' буквы: наружный контур и его дырки; незамкнутый — отдельно; оставшееся (не должно быть) — отдельно
    ReDim LtIdx(0 To n): ReDim g0(0 To n): ReDim gn(0 To n): ReDim used(0 To n)
    ng = 0: t = 0
    For k = 0 To n - 1
        If Not has(k) Or dep(k) Mod 2 = 0 Then
            g0(ng) = t
            LtIdx(t) = k: t = t + 1: used(k) = True
            If has(k) Then
                For i = 0 To n - 1
                    If has(i) And dep(i) Mod 2 = 1 And par(i) = k Then
                        LtIdx(t) = i: t = t + 1: used(i) = True
                    End If
                Next i
            End If
            gn(ng) = t - g0(ng)
            ng = ng + 1
        End If
    Next k
    For k = 0 To n - 1
        If Not used(k) Then
            g0(ng) = t: LtIdx(t) = k: t = t + 1: gn(ng) = 1
            ng = ng + 1
        End If
    Next k
    LtAttach ng, g0, gn, bx0, by0, bx1, by1, par
    ' слева направо: по левому краю наружного контура, затем сверху, затем по номеру
    ReDim ord(0 To ng)
    For i = 0 To ng - 1
        k = i
        Do While k > 0
            If Not LtBefore(LtIdx(g0(i)), LtIdx(g0(ord(k - 1))), bx0, by1) Then
                Exit Do
            End If
            ord(k) = ord(k - 1)
            k = k - 1
        Loop
        ord(k) = i
    Next i
    LtNG = ng
    ReDim LtG0(0 To ng): ReDim LtGn(0 To ng)
    For i = 0 To ng - 1
        LtG0(i) = g0(ord(i)): LtGn(i) = gn(ord(i))
    Next i
End Sub

' точки и акценты — к своей букве (Ё, Й, i, ї, !, ?, ÷): кусок, который намного меньше буквы, стоит над
' ней или под ней и близко по вертикали, идёт к ней; из таких — к той, у которой меньше зазор по
' вертикали плюс сдвиг центров по горизонтали. Только внутри одного уровня вложенности (буквы в дырке
' рамки к рамке не липнут). Куски сливаются в LtIdx: буква-корень первой, затем остальные по номеру
Private Sub LtAttach(ng As Long, g0() As Long, gn() As Long, bx0() As Double, by0() As Double, _
                     bx1() As Double, by1() As Double, par() As Long)
    Dim i As Long, j As Long, k As Long, r As Long, t As Long, best As Long, a As Long, b As Long
    Dim x0 As Double, y0 As Double, x1 As Double, y1 As Double, u0 As Double, v0 As Double, u1 As Double, v1 As Double
    Dim cx As Double, cy As Double, gap As Double, sc As Double, bs As Double, ba As Double
    Dim sz() As Double, ar() As Double, up() As Long, nidx() As Long, n0() As Long, nn() As Long, nng As Long
    If ng < 2 Then
        Exit Sub
    End If
    ReDim sz(0 To ng): ReDim ar(0 To ng): ReDim up(0 To ng)
    For i = 0 To ng - 1
        k = LtIdx(g0(i))
        sz(i) = bx1(k) - bx0(k)
        If by1(k) - by0(k) > sz(i) Then
            sz(i) = by1(k) - by0(k)
        End If
        ar(i) = (bx1(k) - bx0(k)) * (by1(k) - by0(k))
        up(i) = i
    Next i
    For i = 0 To ng - 1
        k = LtIdx(g0(i))
        x0 = bx0(k): y0 = by0(k): x1 = bx1(k): y1 = by1(k)
        cx = (x0 + x1) / 2
        cy = (y0 + y1) / 2
        best = -1
        For j = 0 To ng - 1
            If j <> i Then
                r = LtIdx(g0(j))
                If par(r) = par(k) And (sz(i) <= MARK_SIZE * sz(j) Or _
                        (ar(i) <= MARK_AREA * ar(j) And sz(i) <= MARK_SIZE2 * sz(j))) Then
                    u0 = bx0(r): v0 = by0(r): u1 = bx1(r): v1 = by1(r)
                    If cx >= u0 - MARK_X * sz(j) And cx <= u1 + MARK_X * sz(j) And (cy < v0 Or cy > v1) Then
                        gap = 0
                        If v0 - y1 > gap Then
                            gap = v0 - y1
                        End If
                        If y0 - v1 > gap Then
                            gap = y0 - v1
                        End If
                        If gap <= MARK_GAP * sz(j) Then
                            sc = gap + Abs(cx - (u0 + u1) / 2)
                            If best < 0 Then
                                best = j: bs = sc: ba = ar(j)
                            ElseIf sc < bs Or (sc = bs And ar(j) > ba) Then
                                best = j: bs = sc: ba = ar(j)
                            End If
                        End If
                    End If
                End If
            End If
        Next j
        If best >= 0 Then
            a = LtFind(up, i)
            b = LtFind(up, best)
            If a <> b Then
                up(a) = b
            End If
        End If
    Next i
    ReDim nidx(0 To UBound(LtIdx)): ReDim n0(0 To ng): ReDim nn(0 To ng)
    t = 0: nng = 0
    For r = 0 To ng - 1
        If LtFind(up, r) = r Then
            n0(nng) = t
            For k = g0(r) To g0(r) + gn(r) - 1
                nidx(t) = LtIdx(k): t = t + 1
            Next k
            For i = 0 To ng - 1
                If i <> r Then
                    If LtFind(up, i) = r Then
                        For k = g0(i) To g0(i) + gn(i) - 1
                            nidx(t) = LtIdx(k): t = t + 1
                        Next k
                    End If
                End If
            Next i
            nn(nng) = t - n0(nng)
            nng = nng + 1
        End If
    Next r
    For k = 0 To t - 1
        LtIdx(k) = nidx(k)
    Next k
    For i = 0 To nng - 1
        g0(i) = n0(i): gn(i) = nn(i)
    Next i
    ng = nng
End Sub

Private Function LtFind(up() As Long, ByVal i As Long) As Long
    Do While up(i) <> i
        i = up(i)
    Loop
    LtFind = i
End Function

' буква с наружным контуром a раньше, чем с b
Private Function LtBefore(ByVal a As Long, ByVal b As Long, bx0() As Double, by1() As Double) As Boolean
    If bx0(a) <> bx0(b) Then
        LtBefore = (bx0(a) < bx0(b))
    ElseIf by1(a) <> by1(b) Then
        LtBefore = (by1(a) > by1(b))
    Else
        LtBefore = (a < b)
    End If
End Function

' точка внутри контура k (чёт-нечет по точкам контура)
Private Function LtInside(ByVal px As Double, ByVal py As Double, ByVal k As Long) As Boolean
    Dim m As Long, i As Long, j As Long, ax As Double, ay As Double, bx As Double, by As Double, c As Boolean
    For m = 0 To QPn(k) - 1
        i = QP0(k) + m
        j = QP0(k) + (m + 1) Mod QPn(k)
        ax = SmX(i): ay = SmY(i): bx = SmX(j): by = SmY(j)
        If (ay > py) <> (by > py) Then
            If px < (bx - ax) * (py - ay) / (by - ay) + ax Then
                c = Not c
            End If
        End If
    Next m
    LtInside = c
End Function
'@@LETTERSGEO_END

'@@NESTGEO_BEGIN
' ---------------------------------------------------------------------
'  Раскладка на листы: MaxRects без поворота (повторяет proto/nesting.py один в один).
'  Вход: габариты деталей NtW/NtH (NNt штук). Деталь — габарит + зазор, лист — без полей + зазор:
'  так между деталями ровно зазор, от края — ровно поле. Детали по очереди, каждая — на самый
'  ранний лист, где есть место, на листе — лучшее место по эвристике. Четыре порядка × три
'  эвристики, лучший результат: меньше листов, затем меньше занятый габарит последнего листа.
'  autoW: длина листа (по горизонтали, sw) — по раскладке (поле + детали + поле, вверх до 10 мм, от wmin до sw);
'  тогда ещё эвристика «по столбцам», а при равном числе листов лучше меньшая общая ширина.
'  fillMode: 0 — только «по столбцам» (сверху вниз, затем вправо), 1 — только «по рядам», 2 — все.
'  Выход: NtS (лист с 0, -1 — не помещается), NtX/NtY (от левого верхнего угла листа, y вниз),
'  NtSheets, ширина каждого листа NtSW.
' ---------------------------------------------------------------------
Public Sub NestPack(ByVal sw As Double, ByVal sh As Double, ByVal margin As Double, ByVal gap As Double, _
                    ByVal autoW As Boolean, ByVal wmin As Double, ByVal fillMode As Long)
    Dim bw As Double, bh As Double, mode As Long, heur As Long, nsh As Long, la As Double, q As Long
    Dim bestN As Long, bestA As Double, has As Boolean, i As Long, s As Long, w As Double, hs As Variant
    ReDim NtPW(0 To NNt): ReDim NtPH(0 To NNt)
    ReDim NtS(0 To NNt): ReDim NtX(0 To NNt): ReDim NtY(0 To NNt)
    ReDim RnS(0 To NNt): ReDim RnX(0 To NNt): ReDim RnY(0 To NNt)
    For i = 0 To NNt - 1
        NtPW(i) = NtW(i) + gap: NtPH(i) = NtH(i) + gap
        NtS(i) = -1: NtX(i) = 0: NtY(i) = 0
    Next i
    bw = sw - 2 * margin + gap
    bh = sh - 2 * margin + gap
    NtSheets = 0: NtMode = 0: NtHeur = 0
    ' заполнение: 0 — столбцами (левее, затем выше), 1 — рядами (выше, затем левее), 2 — как плотнее
    If fillMode = 0 Then
        hs = Array(3)
    ElseIf fillMode = 1 Then
        hs = Array(1)
    ElseIf autoW Then
        hs = Array(0, 1, 2, 3)
    Else
        hs = Array(0, 1, 2)
    End If
    For mode = 0 To 3
        NtOrder mode
        For q = 0 To UBound(hs)
            heur = hs(q)
            nsh = NtRun(bw, bh, heur)
            If autoW Then
                NtUsedW nsh
                la = 0
                For s = 0 To nsh - 1
                    la = la + NtSW(s)
                Next s
            Else
                la = NtLastArea(nsh)
            End If
            If Not has Or nsh < bestN Or (nsh = bestN And la < bestA - NT_EPS) Then
                has = True
                bestN = nsh: bestA = la
                NtMode = mode: NtHeur = heur
                For i = 0 To NNt - 1
                    NtS(i) = RnS(i)
                    If RnS(i) >= 0 Then
                        NtX(i) = RnX(i) + margin: NtY(i) = RnY(i) + margin
                    Else
                        NtX(i) = 0: NtY(i) = 0
                    End If
                Next i
            End If
        Next q
    Next mode
    NtSheets = bestN
    ' ширина листов: у лучшего прохода
    For i = 0 To NNt - 1
        RnS(i) = NtS(i): RnX(i) = NtX(i) - margin
    Next i
    NtUsedW NtSheets
    For s = 0 To NtSheets - 1
        If autoW Then
            w = -Int(-(2 * margin + NtSW(s) - gap) / 10 + 0.000000001) * 10
            If w < wmin Then
                w = wmin
            End If
            If w > sw Then
                w = sw
            End If
            NtSW(s) = w
        Else
            NtSW(s) = sw
        End If
    Next s
End Sub

' занятая ширина каждого листа прохода RnS/RnX (с зазором справа) — в NtSW
Private Sub NtUsedW(ByVal nsh As Long)
    Dim i As Long, s As Long
    ReDim NtSW(0 To nsh)
    For i = 0 To NNt - 1
        s = RnS(i)
        If s >= 0 Then
            If RnX(i) + NtPW(i) > NtSW(s) Then
                NtSW(s) = RnX(i) + NtPW(i)
            End If
        End If
    Next i
End Sub

' порядок деталей NtOrd: 0 — по высоте, 1 — по ширине, 2 — по площади, 3 — по большей стороне (по убыванию)
Private Sub NtOrder(ByVal mode As Long)
    Dim i As Long, j As Long, t As Long, tmp() As Long, w As Long, lo As Long, md As Long, hi As Long
    Dim a As Long, b As Long, k As Long
    ReDim NtOrd(0 To NNt): ReDim tmp(0 To NNt)
    For i = 0 To NNt - 1
        NtOrd(i) = i
    Next i
    ' сортировка слиянием снизу вверх (порядок полный: при равенстве — по номеру)
    w = 1
    Do While w < NNt
        lo = 0
        Do While lo < NNt
            md = lo + w
            If md > NNt Then
                md = NNt
            End If
            hi = lo + 2 * w
            If hi > NNt Then
                hi = NNt
            End If
            a = lo: b = md: k = lo
            Do While k < hi
                If a < md And b < hi Then
                    If NtLess(NtOrd(b), NtOrd(a), mode) Then
                        tmp(k) = NtOrd(b): b = b + 1
                    Else
                        tmp(k) = NtOrd(a): a = a + 1
                    End If
                ElseIf a < md Then
                    tmp(k) = NtOrd(a): a = a + 1
                Else
                    tmp(k) = NtOrd(b): b = b + 1
                End If
                k = k + 1
            Loop
            lo = hi
        Loop
        For j = 0 To NNt - 1
            t = tmp(j): NtOrd(j) = t
        Next j
        w = 2 * w
    Loop
End Sub

' деталь i раньше j
Private Function NtLess(ByVal i As Long, ByVal j As Long, ByVal mode As Long) As Boolean
    Dim a1 As Double, a2 As Double, b1 As Double, b2 As Double
    Select Case mode
    Case 0
        a1 = NtPH(i): a2 = NtPW(i): b1 = NtPH(j): b2 = NtPW(j)
    Case 1
        a1 = NtPW(i): a2 = NtPH(i): b1 = NtPW(j): b2 = NtPH(j)
    Case 2
        a1 = NtPW(i) * NtPH(i): a2 = NtPH(i): b1 = NtPW(j) * NtPH(j): b2 = NtPH(j)
    Case Else
        a1 = NtPW(i): a2 = NtPH(i)
        If a2 > a1 Then
            a1 = NtPH(i): a2 = NtPW(i)
        End If
        b1 = NtPW(j): b2 = NtPH(j)
        If b2 > b1 Then
            b1 = NtPH(j): b2 = NtPW(j)
        End If
    End Select
    If a1 <> b1 Then
        NtLess = (a1 > b1)
    ElseIf a2 <> b2 Then
        NtLess = (a2 > b2)
    Else
        NtLess = (i < j)
    End If
End Function

' один проход по порядку NtOrd: RnS/RnX/RnY, результат — число листов
Private Function NtRun(ByVal bw As Double, ByVal bh As Double, ByVal heur As Long) As Long
    Dim q As Long, i As Long, j As Long, nsh As Long, pw As Double, ph As Double, s As Long
    Dim x As Double, y As Double, bj As Long, bs As Long, b1 As Double, b2 As Double
    Dim s1 As Double, s2 As Double, dw As Double, dh As Double, better As Boolean
    Dim mw() As Double, mh() As Double
    NFr = 0
    ReDim FrS(0 To 63): ReDim FrX(0 To 63): ReDim FrY(0 To 63): ReDim FrW(0 To 63): ReDim FrH(0 To 63)
    For i = 0 To NNt - 1
        RnS(i) = -1: RnX(i) = 0: RnY(i) = 0
    Next i
    ' наименьшие ширина и высота среди деталей после q-й: свободное место меньше — уже никому
    ReDim mw(0 To NNt): ReDim mh(0 To NNt)
    mw(NNt) = 1E+300: mh(NNt) = 1E+300
    For q = NNt - 1 To 0 Step -1
        mw(q) = mw(q + 1): mh(q) = mh(q + 1)
        If NtPW(NtOrd(q)) < mw(q) Then
            mw(q) = NtPW(NtOrd(q))
        End If
        If NtPH(NtOrd(q)) < mh(q) Then
            mh(q) = NtPH(NtOrd(q))
        End If
    Next q
    For q = 0 To NNt - 1
        i = NtOrd(q)
        pw = NtPW(i): ph = NtPH(i)
        If pw <= bw + NT_EPS And ph <= bh + NT_EPS Then
            bj = -1
            For j = 0 To NFr - 1
                If bj >= 0 And FrS(j) > bs Then
                    ' лист позже лучшего — не лучше
                ElseIf pw <= FrW(j) + NT_EPS And ph <= FrH(j) + NT_EPS Then
                    dw = FrW(j) - pw: dh = FrH(j) - ph
                    If heur = 0 Then
                        ' лучшая короткая сторона
                        If dw < dh Then
                            s1 = dw: s2 = dh
                        Else
                            s1 = dh: s2 = dw
                        End If
                    ElseIf heur = 1 Then
                        ' выше и левее (по строкам)
                        s1 = FrY(j) + ph: s2 = FrX(j)
                    ElseIf heur = 2 Then
                        ' лучшая площадь
                        s1 = FrW(j) * FrH(j) - pw * ph
                        If dw < dh Then
                            s2 = dw
                        Else
                            s2 = dh
                        End If
                    Else
                        ' левее и выше (по столбцам: раскладка как можно уже)
                        s1 = FrX(j) + pw: s2 = FrY(j)
                    End If
                    better = False
                    If bj < 0 Then
                        better = True
                    ElseIf FrS(j) < bs Then
                        better = True
                    ElseIf FrS(j) = bs Then
                        If s1 < b1 Then
                            better = True
                        ElseIf s1 = b1 And s2 < b2 Then
                            better = True
                        End If
                    End If
                    If better Then
                        bj = j: bs = FrS(j): b1 = s1: b2 = s2
                    End If
                End If
            Next j
            If bj < 0 Then
                s = nsh
                nsh = nsh + 1
                NtAddFr s, 0, 0, bw, bh
                x = 0: y = 0
            Else
                s = bs
                x = FrX(bj): y = FrY(bj)
            End If
            RnS(i) = s: RnX(i) = x: RnY(i) = y
            NtSplit s, x, y, pw, ph, mw(q + 1), mh(q + 1)
        End If
    Next q
    NtRun = nsh
End Function

Private Sub NtAddFr(ByVal s As Long, ByVal x As Double, ByVal y As Double, ByVal w As Double, ByVal h As Double)
    If NFr > UBound(FrS) Then
        ReDim Preserve FrS(0 To 2 * NFr): ReDim Preserve FrX(0 To 2 * NFr): ReDim Preserve FrY(0 To 2 * NFr)
        ReDim Preserve FrW(0 To 2 * NFr): ReDim Preserve FrH(0 To 2 * NFr)
    End If
    FrS(NFr) = s: FrX(NFr) = x: FrY(NFr) = y: FrW(NFr) = w: FrH(NFr) = h
    NFr = NFr + 1
End Sub

Private Sub NtAddNw(ByVal x As Double, ByVal y As Double, ByVal w As Double, ByVal h As Double)
    If NNw > UBound(NwX) Then
        ReDim Preserve NwX(0 To 2 * NNw): ReDim Preserve NwY(0 To 2 * NNw)
        ReDim Preserve NwW(0 To 2 * NNw): ReDim Preserve NwH(0 To 2 * NNw)
    End If
    NwX(NNw) = x: NwY(NNw) = y: NwW(NNw) = w: NwH(NNw) = h
    NNw = NNw + 1
End Sub

' деталь (x, y, w, h) на листе s: свободные прямоугольники, которые она задевает, — на части;
' части, лежащие внутри других, — долой; уже́ mw или ниже mh (меньше любой оставшейся детали) — тоже
Private Sub NtSplit(ByVal s As Long, ByVal x As Double, ByVal y As Double, ByVal w As Double, ByVal h As Double, _
                    ByVal mw As Double, ByVal mh As Double)
    Dim j As Long, nKeep As Long, ia As Long, ib As Long, rx As Double, ry As Double, rw As Double, rh As Double
    Dim drop As Boolean, nNew As Long
    NNw = 0
    ReDim NwX(0 To 15): ReDim NwY(0 To 15): ReDim NwW(0 To 15): ReDim NwH(0 To 15)
    nKeep = 0
    For j = 0 To NFr - 1
        rx = FrX(j): ry = FrY(j): rw = FrW(j): rh = FrH(j)
        If rw < mw - NT_EPS Or rh < mh - NT_EPS Then
            ' никому не подойдёт
        ElseIf FrS(j) <> s Or x >= rx + rw - NT_EPS Or x + w <= rx + NT_EPS Or y >= ry + rh - NT_EPS Or y + h <= ry + NT_EPS Then
            FrS(nKeep) = FrS(j): FrX(nKeep) = rx: FrY(nKeep) = ry: FrW(nKeep) = rw: FrH(nKeep) = rh
            nKeep = nKeep + 1
        Else
            If x > rx + NT_EPS Then
                NtAddNw rx, ry, x - rx, rh
            End If
            If x + w < rx + rw - NT_EPS Then
                NtAddNw x + w, ry, rx + rw - (x + w), rh
            End If
            If y > ry + NT_EPS Then
                NtAddNw rx, ry, rw, y - ry
            End If
            If y + h < ry + rh - NT_EPS Then
                NtAddNw rx, y + h, rw, ry + rh - (y + h)
            End If
        End If
    Next j
    NFr = nKeep
    ' новые — только те, куда ещё что-то встанет
    nNew = 0
    For ia = 0 To NNw - 1
        If NwW(ia) >= mw - NT_EPS And NwH(ia) >= mh - NT_EPS Then
            NwX(nNew) = NwX(ia): NwY(nNew) = NwY(ia): NwW(nNew) = NwW(ia): NwH(nNew) = NwH(ia)
            nNew = nNew + 1
        End If
    Next ia
    NNw = nNew
    For ia = 0 To NNw - 1
        drop = False
        For ib = 0 To NNw - 1
            If ib <> ia Then
                If NtIn(NwX(ib), NwY(ib), NwW(ib), NwH(ib), NwX(ia), NwY(ia), NwW(ia), NwH(ia)) Then
                    If ib < ia Then
                        drop = True
                    ElseIf Not NtIn(NwX(ia), NwY(ia), NwW(ia), NwH(ia), NwX(ib), NwY(ib), NwW(ib), NwH(ib)) Then
                        drop = True
                    End If
                    If drop Then
                        Exit For
                    End If
                End If
            End If
        Next ib
        If Not drop Then
            For j = 0 To nKeep - 1
                If FrS(j) = s Then
                    If NtIn(FrX(j), FrY(j), FrW(j), FrH(j), NwX(ia), NwY(ia), NwW(ia), NwH(ia)) Then
                        drop = True
                        Exit For
                    End If
                End If
            Next j
        End If
        If Not drop Then
            NtAddFr s, NwX(ia), NwY(ia), NwW(ia), NwH(ia)
        End If
    Next ia
End Sub

' прямоугольник a внутри b
Private Function NtIn(ByVal bx As Double, ByVal by As Double, ByVal bw As Double, ByVal bh As Double, _
                      ByVal ax As Double, ByVal ay As Double, ByVal aw As Double, ByVal ah As Double) As Boolean
    NtIn = (ax >= bx - NT_EPS And ay >= by - NT_EPS And ax + aw <= bx + bw + NT_EPS And ay + ah <= by + bh + NT_EPS)
End Function

' занятый габарит последнего листа прохода
Private Function NtLastArea(ByVal nsh As Long) As Double
    Dim i As Long, mx As Double, my As Double
    If nsh = 0 Then
        Exit Function
    End If
    For i = 0 To NNt - 1
        If RnS(i) = nsh - 1 Then
            If RnX(i) + NtPW(i) > mx Then
                mx = RnX(i) + NtPW(i)
            End If
            If RnY(i) + NtPH(i) > my Then
                my = RnY(i) + NtPH(i)
            End If
        End If
    Next i
    NtLastArea = mx * my
End Function
'@@NESTGEO_END
'@@NEST_END
