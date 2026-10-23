Attribute VB_Name = "NeonCut"
' =====================================================================
'  NeonCenterline — кратность реза неона / ленты.
'
'  Настройки (Настройки плагина):
'    Cut\Type   classic — обычный светодиодный неон (лента внутри: кусок неона = длина паза),
'               neon2   — неон 2.0 (силиконовый рассеиватель в пазу, лента отдельно);
'    Cut\Step   кратность реза, мм (0 — не учитывать);
'    Cut\Fit    1 — после расчёта подгонять линии под кратность (укорачивать свободные концы);
'    Cut\Trim   both — поровну с обоих свободных концов, end — только с конца линии.
'
'  Подгонка: открытый кусок укорачивается до ближайшей меньшей кратной длины — только за счёт
'  СВОБОДНЫХ концов (конец, который упирается в другую линию, не трогается). Удлинять нельзя:
'  свободный конец уже стоит на D/2 от края контура. Замкнутые куски и куски без свободных
'  концов не меняются — остаток показывается в отчёте.
' =====================================================================
Option Explicit

Private Const REG_APP As String = "NeonCenterline"
Private Const JOIN_TOL As Double = 0.8        ' мм: конец ближе к другой линии — не свободный
Private Const FLAT_N As Long = 48             ' точек на сегмент при измерении длины

' ---- данные кусков текущей обработки ----
Private SX() As Double, SY() As Double        ' точки сегментов: 4 на сегмент (P0, P1, P2, P3)
Private PcShape() As Long, PcSub() As Long, PcS0() As Long, PcN() As Long, PcClosed() As Boolean, NPc As Long, NSeg As Long

Public Function NeonCutStep() As Double
    NeonCutStep = ParseNum(GetSetting(REG_APP, "Cut", "Step", "10"))
End Function

Public Function NeonCutType() As String
    NeonCutType = GetSetting(REG_APP, "Cut", "Type", "classic")
End Function

Public Function NeonCutFitOn() As Boolean
    NeonCutFitOn = (GetSetting(REG_APP, "Cut", "Fit", "0") = "1")
End Function

' команда: подогнать выделенные центральные линии под кратность реза
Public Sub FitToCut()
    Dim col As Collection, rep As String
    If ActiveDocument Is Nothing Then
        MsgBox NeonL("err_nodoc"), vbExclamation, NeonL("app_title")
        Exit Sub
    End If
    If NeonCutStep() <= 0 Then
        MsgBox NeonL("cut_nostep"), vbExclamation, NeonL("app_title")
        Exit Sub
    End If
    Set col = New Collection
    CollectCurves ActiveSelectionRange, col
    If col.Count = 0 Then
        MsgBox NeonL("len_none"), vbExclamation, NeonL("app_title")
        Exit Sub
    End If
    ActiveDocument.BeginCommandGroup NeonL("cut_title")
    rep = NeonCutFitShapes(col)
    ActiveDocument.EndCommandGroup
    MsgBox rep, vbInformation, NeonL("cut_title")
End Sub

' подгонка набора кривых (Shape); возвращает текст отчёта
Public Function NeonCutFitShapes(col As Collection) As String
    Dim stp As Double, i As Long, k As Long, L As Double, tgt As Double, cut As Double
    Dim f0 As Boolean, f1 As Boolean, a As Double, b As Double, oldUnit As cdrUnit
    Dim nFit As Long, nOk As Long, nSkipC As Long, nSkipJ As Long, nShort As Long, totCut As Double, trimEnd As Boolean
    Dim changed() As Boolean, rest As Double
    stp = NeonCutStep()
    If stp <= 0 Or col.Count = 0 Then
        Exit Function
    End If
    trimEnd = (GetSetting(REG_APP, "Cut", "Trim", "both") = "end")
    oldUnit = ActiveDocument.Unit
    ActiveDocument.Unit = cdrMillimeter
    LoadPieces col
    ReDim changed(0 To col.Count)
    For i = 0 To NPc - 1
        L = PieceLen(i)
        tgt = Int(L / stp + 0.000001) * stp
        rest = L - tgt
        If PcClosed(i) Then
            If rest > 0.5 And stp - rest > 0.5 Then
                nSkipC = nSkipC + 1
            Else
                nOk = nOk + 1
            End If
        ElseIf rest <= 0.5 Then
            nOk = nOk + 1
        ElseIf tgt < stp Then
            nShort = nShort + 1                       ' кусок короче одного шага
        Else
            f0 = IsFreeEnd(i, 0): f1 = IsFreeEnd(i, 1)
            cut = rest
            If Not f0 And Not f1 Then
                nSkipJ = nSkipJ + 1
            Else
                a = 0: b = 0
                If f0 And f1 Then
                    If trimEnd Then
                        b = cut
                    Else
                        a = cut / 2: b = cut / 2
                    End If
                ElseIf f0 Then
                    a = cut
                Else
                    b = cut
                End If
                TrimPiece i, a, b
                changed(PcShape(i)) = True
                nFit = nFit + 1: totCut = totCut + cut
            End If
        End If
    Next i
    For k = 1 To col.Count
        If changed(k) Then
            RebuildShape col(k), k
        End If
    Next k
    ActiveDocument.Unit = oldUnit
    NeonCutFitShapes = NeonLF("cut_report", FmtNum(stp), nFit, Format$(totCut, "0"), nOk, nSkipC + nSkipJ + nShort)
End Function

' отчёт по кратности для кусков (для кнопки «Длина»): сколько неона/ленты кратно резу
Public Function NeonCutSummary(col As Collection) As String
    Dim stp As Double, i As Long, L As Double, tot As Double, totK As Double, oldUnit As cdrUnit
    stp = NeonCutStep()
    If stp <= 0 Or col.Count = 0 Then
        Exit Function
    End If
    oldUnit = ActiveDocument.Unit
    ActiveDocument.Unit = cdrMillimeter
    LoadPieces col
    For i = 0 To NPc - 1
        L = PieceLen(i)
        tot = tot + L
        totK = totK + Int(L / stp + 0.000001) * stp
    Next i
    ActiveDocument.Unit = oldUnit
    If NeonCutType() = "neon2" Then
        NeonCutSummary = NeonLF("cut_sum_neon2", FmtNum(stp), Format$(totK, "0"), Format$(tot - totK, "0"))
    Else
        NeonCutSummary = NeonLF("cut_sum_classic", FmtNum(stp), Format$(totK, "0"), Format$(tot - totK, "0"))
    End If
End Function

' ---------------------------------------------------------------------
'  Чтение кусков (подпутей) кривых в массивы сегментов
' ---------------------------------------------------------------------
Private Sub LoadPieces(col As Collection)
    Dim k As Long, sp As SubPath, sg As Segment, si As Long
    Dim x0 As Double, y0 As Double, x1 As Double, y1 As Double, x2 As Double, y2 As Double, x3 As Double, y3 As Double
    NPc = 0: NSeg = 0
    ReDim PcShape(0 To 15): ReDim PcSub(0 To 15): ReDim PcS0(0 To 15): ReDim PcN(0 To 15): ReDim PcClosed(0 To 15)
    ReDim SX(0 To 255): ReDim SY(0 To 255)
    For k = 1 To col.Count
        si = 0
        For Each sp In col(k).Curve.SubPaths
            si = si + 1
            If NPc > UBound(PcShape) Then
                ReDim Preserve PcShape(0 To 2 * NPc): ReDim Preserve PcSub(0 To 2 * NPc): ReDim Preserve PcS0(0 To 2 * NPc)
                ReDim Preserve PcN(0 To 2 * NPc): ReDim Preserve PcClosed(0 To 2 * NPc)
            End If
            PcShape(NPc) = k: PcSub(NPc) = si: PcS0(NPc) = NSeg: PcN(NPc) = 0: PcClosed(NPc) = sp.Closed
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
                PutSeg NSeg, x0, y0, x1, y1, x2, y2, x3, y3
                NSeg = NSeg + 1: PcN(NPc) = PcN(NPc) + 1
            Next sg
            NPc = NPc + 1
        Next sp
    Next k
End Sub

Private Sub PutSeg(ByVal s As Long, ByVal x0 As Double, ByVal y0 As Double, ByVal x1 As Double, ByVal y1 As Double, _
                   ByVal x2 As Double, ByVal y2 As Double, ByVal x3 As Double, ByVal y3 As Double)
    If 4 * s + 3 > UBound(SX) Then
        ReDim Preserve SX(0 To 8 * s + 16): ReDim Preserve SY(0 To 8 * s + 16)
    End If
    SX(4 * s) = x0: SY(4 * s) = y0: SX(4 * s + 1) = x1: SY(4 * s + 1) = y1
    SX(4 * s + 2) = x2: SY(4 * s + 2) = y2: SX(4 * s + 3) = x3: SY(4 * s + 3) = y3
End Sub

Private Sub BzAt(ByVal s As Long, ByVal t As Double, x As Double, y As Double)
    Dim mt As Double
    mt = 1 - t
    x = mt * mt * mt * SX(4 * s) + 3 * mt * mt * t * SX(4 * s + 1) + 3 * mt * t * t * SX(4 * s + 2) + t * t * t * SX(4 * s + 3)
    y = mt * mt * mt * SY(4 * s) + 3 * mt * mt * t * SY(4 * s + 1) + 3 * mt * t * t * SY(4 * s + 2) + t * t * t * SY(4 * s + 3)
End Sub

Private Function CutSegLen(ByVal s As Long) As Double
    Dim k As Long, x As Double, y As Double, px As Double, py As Double, L As Double
    px = SX(4 * s): py = SY(4 * s)
    For k = 1 To FLAT_N
        BzAt s, k / FLAT_N, x, y
        L = L + Sqr((x - px) ^ 2 + (y - py) ^ 2)
        px = x: py = y
    Next k
    CutSegLen = L
End Function

Private Function PieceLen(ByVal i As Long) As Double
    Dim s As Long, L As Double
    For s = PcS0(i) To PcS0(i) + PcN(i) - 1
        L = L + CutSegLen(s)
    Next s
    PieceLen = L
End Function

' параметр t на сегменте s, где длина от начала сегмента = d
Private Function ParamAt(ByVal s As Long, ByVal d As Double) As Double
    Dim k As Long, x As Double, y As Double, px As Double, py As Double, L As Double, dl As Double
    px = SX(4 * s): py = SY(4 * s)
    For k = 1 To FLAT_N
        BzAt s, k / FLAT_N, x, y
        dl = Sqr((x - px) ^ 2 + (y - py) ^ 2)
        If L + dl >= d Then
            If dl > 0 Then
                ParamAt = (k - 1 + (d - L) / dl) / FLAT_N
            Else
                ParamAt = k / FLAT_N
            End If
            Exit Function
        End If
        L = L + dl: px = x: py = y
    Next k
    ParamAt = 1
End Function

' конец куска (0 — начало, 1 — конец) свободен: рядом (JOIN_TOL) нет других линий
Private Function IsFreeEnd(ByVal i As Long, ByVal which As Long) As Boolean
    Dim ex As Double, ey As Double, j As Long, s As Long, k As Long, x As Double, y As Double, px As Double, py As Double
    If which = 0 Then
        ex = SX(4 * PcS0(i)): ey = SY(4 * PcS0(i))
    Else
        s = PcS0(i) + PcN(i) - 1
        ex = SX(4 * s + 3): ey = SY(4 * s + 3)
    End If
    For j = 0 To NPc - 1
        If j <> i Then
            For s = PcS0(j) To PcS0(j) + PcN(j) - 1
                px = SX(4 * s): py = SY(4 * s)
                For k = 1 To FLAT_N
                    BzAt s, k / FLAT_N, x, y
                    If PtSegDist(ex, ey, px, py, x, y) < JOIN_TOL Then
                        Exit Function
                    End If
                    px = x: py = y
                Next k
            Next s
        End If
    Next j
    IsFreeEnd = True
End Function

Private Function PtSegDist(ByVal x As Double, ByVal y As Double, ByVal ax As Double, ByVal ay As Double, _
                           ByVal bx As Double, ByVal by As Double) As Double
    Dim dx As Double, dy As Double, L2 As Double, t As Double
    dx = bx - ax: dy = by - ay: L2 = dx * dx + dy * dy
    If L2 > 0 Then
        t = ((x - ax) * dx + (y - ay) * dy) / L2
        If t < 0 Then
            t = 0
        ElseIf t > 1 Then
            t = 1
        End If
    End If
    PtSegDist = Sqr((x - ax - dx * t) ^ 2 + (y - ay - dy * t) ^ 2)
End Function

' срезать a мм с начала и b мм с конца куска i (сегменты делятся по де Кастельжо)
Private Sub TrimPiece(ByVal i As Long, ByVal a As Double, ByVal b As Double)
    Dim s As Long, s0 As Long, s1 As Long, L As Double, acc As Double, t As Double
    s0 = PcS0(i): s1 = PcS0(i) + PcN(i) - 1
    If a > 0 Then
        acc = 0
        For s = s0 To s1
            L = CutSegLen(s)
            If acc + L > a Then
                t = ParamAt(s, a - acc)
                SplitKeep s, t, True
                ' сегменты до s выкидываются: сдвиг начала куска
                PcN(i) = PcN(i) - (s - PcS0(i)): PcS0(i) = s
                Exit For
            End If
            acc = acc + L
        Next s
    End If
    s0 = PcS0(i): s1 = PcS0(i) + PcN(i) - 1
    If b > 0 Then
        acc = 0
        For s = s1 To s0 Step -1
            L = CutSegLen(s)
            If acc + L > b Then
                t = ParamAt(s, L - (b - acc))
                SplitKeep s, t, False
                PcN(i) = s - PcS0(i) + 1
                Exit For
            End If
            acc = acc + L
        Next s
    End If
End Sub

' оставить у сегмента s часть после t (keepRight) или до t
Private Sub SplitKeep(ByVal s As Long, ByVal t As Double, ByVal keepRight As Boolean)
    Dim x(0 To 3) As Double, y(0 To 3) As Double, ax As Double, ay As Double, bx As Double, by As Double
    Dim cx As Double, cy As Double, dx As Double, dy As Double, ex As Double, ey As Double, fx As Double, fy As Double, k As Long
    For k = 0 To 3
        x(k) = SX(4 * s + k): y(k) = SY(4 * s + k)
    Next k
    ax = x(0) + (x(1) - x(0)) * t: ay = y(0) + (y(1) - y(0)) * t
    bx = x(1) + (x(2) - x(1)) * t: by = y(1) + (y(2) - y(1)) * t
    cx = x(2) + (x(3) - x(2)) * t: cy = y(2) + (y(3) - y(2)) * t
    dx = ax + (bx - ax) * t: dy = ay + (by - ay) * t
    ex = bx + (cx - bx) * t: ey = by + (cy - by) * t
    fx = dx + (ex - dx) * t: fy = dy + (ey - dy) * t
    If keepRight Then
        PutSeg s, fx, fy, ex, ey, cx, cy, x(3), y(3)
    Else
        PutSeg s, x(0), y(0), ax, ay, dx, dy, fx, fy
    End If
End Sub

' пересобрать кривую фигуры k из (возможно, укороченных) кусков
Private Sub RebuildShape(shp As Shape, ByVal k As Long)
    Dim crv As Curve, sp As SubPath, i As Long, s As Long, isLine As Boolean
    Set crv = Application.CreateCurve(ActiveDocument)
    For i = 0 To NPc - 1
        If PcShape(i) = k And PcN(i) > 0 Then
            Set sp = crv.CreateSubPath(SX(4 * PcS0(i)), SY(4 * PcS0(i)))
            For s = PcS0(i) To PcS0(i) + PcN(i) - 1
                isLine = IsStraight(s)
                If isLine Then
                    sp.AppendLineSegment SX(4 * s + 3), SY(4 * s + 3)
                Else
                    sp.AppendCurveSegment2 SX(4 * s + 3), SY(4 * s + 3), SX(4 * s + 1), SY(4 * s + 1), SX(4 * s + 2), SY(4 * s + 2)
                End If
            Next s
            If PcClosed(i) Then
                On Error Resume Next
                sp.Closed = True
                On Error GoTo 0
            End If
        End If
    Next i
    shp.Curve.CopyAssign crv
End Sub

' контрольные точки на отрезке P0-P3 — прямой сегмент
Private Function IsStraight(ByVal s As Long) As Boolean
    Dim dx As Double, dy As Double, L As Double, k As Long, d As Double
    dx = SX(4 * s + 3) - SX(4 * s): dy = SY(4 * s + 3) - SY(4 * s)
    L = Sqr(dx * dx + dy * dy)
    If L < 0.000001 Then
        Exit Function
    End If
    For k = 1 To 2
        d = Abs((SX(4 * s + k) - SX(4 * s)) * dy - (SY(4 * s + k) - SY(4 * s)) * dx) / L
        If d > 0.001 Then
            Exit Function
        End If
    Next k
    IsStraight = True
End Function

Private Sub CollectCurves(sr As ShapeRange, col As Collection)
    Dim sh As Shape
    For Each sh In sr
        If sh.Type = cdrGroupShape Then
            CollectCurves sh.Shapes.All, col
        ElseIf sh.Type = cdrCurveShape Then
            col.Add sh
        End If
    Next sh
End Sub
