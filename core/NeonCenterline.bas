Attribute VB_Name = "NeonCenterline"
' =====================================================================
'  NEON CENTERLINE — центральная линия для фрезеровки под гибкий неон
'  CorelDRAW, VBA (GlobalMacros.gms).                 SAB:CORE, 2026
'
'  Запуск: выделить контуры (кривые, текст, группы) -> макрос NeonCenterline.
'  Результат: кривые на слое "Неон - центр" (красная волосяная линия).
'
'  Алгоритм: контуры -> ресэмплинг -> триангуляция Делоне -> хордальная ось
'  -> обрезка "усиков" по локальной ширине штриха -> слияние развилок
'  -> цепочки со сквозным проходом через развилки и "шпильками" в острых
'  возвратах пера -> сглаживание -> концы на D/2 фрезы от края
'  -> аппроксимация кривыми Безье с минимумом узлов.
' =====================================================================
Option Explicit

' ---------- параметры ----------
Private Const P_PRUNE As Double = 1#       ' ветка с вылетом < PRUNE * ширины у развилки — усик
Private Const P_ENDTRIM As Double = 0.92   ' срез зоны скругления у свободного конца
Private Const P_HSTEP As Double = 7#       ' шаг ресэмплинга = W / HSTEP
Private Const P_SMOOTH As Long = 12        ' итераций сглаживания
Private Const P_JZONE As Double = 1#       ' радиус зоны развилки (в ширинах штриха)
Private Const P_STUBZ As Double = 0.5      ' зона развилки со стороны заглушки (конца цепочки)
Private Const P_LMIN As Double = 0.8       ' прямой участок >= LMIN*W в начале/конце куска — отдельным отрезком
Private Const P_WIDEN As Double = 1.05     ' у развилки узлы со штрихом шире WIDEN*типичного ненадёжны
Private Const P_JMERGE As Double = 0.5     ' развилки ближе JMERGE * ширины сливаются (дубли треугольников у Т-стыка)
Private Const P_CORNER As Double = 55#     ' угол (град), с которого ставится острый узел
Private Const P_MERGEANG As Double = 40#   ' макс. излом для сквозного прохода через развилку
Private Const P_TOL As Double = 0.25       ' допуск аппроксимации, мм
Private Const P_FLAT_TOL As Double = 0.02  ' допуск разбиения исходных кривых на отрезки, мм
Private Const P_DOT As Double = 0.5        ' радиус метки-точки, мм
Private Const REG_APP As String = "NeonCenterline"

Private Const MAXD As Long = 16            ' макс. степень узла графа
Private Const MAXE As Long = 32            ' макс. соседей точки в триангуляции
Private Const BIG As Double = 1E+30
Private Const PI_ As Double = 3.14159265358979

Private gW As Double, gTool As Double, gTol As Double, gH As Double

' --- все кольца объекта ---
Private AX() As Double, AY() As Double, NAP As Long
Private AS_() As Long, AN_() As Long, NAR As Long

' --- текущая группа: полигоны (внешний контур + дырки), кольца подряд ---
Private RX() As Double, RY() As Double, RS() As Long, RN() As Long, NR As Long
Private RPoly() As Long
Private PolyR0() As Long, PolyRN() As Long, NPoly As Long
Private PBx0() As Double, PBy0() As Double, PBx1() As Double, PBy1() As Double
' --- корзины рёбер по Y для быстрой проверки "точка внутри" ---
Private Const NBIN As Long = 128
Private BinS() As Long, BinN() As Long, BinA() As Long, BinB() As Long, BinP() As Long
Private BinY0 As Double, BinDY As Double

' --- острые углы контура группы ---
Private BCX() As Double, BCY() As Double, NBC As Long
Private BCC() As Boolean                      ' выпуклый (остриё) / вогнутый угол
Private BE1X() As Double, BE1Y() As Double, BE2X() As Double, BE2Y() As Double   ' кромки от угла
Private BRCX As Double, BRCY As Double, BRCR As Double   ' окружность кусочной проекции (TangentBridge)
Private BS1() As Double, BS2() As Double   ' прямизна кромок (отклонение контура от хорды на 0.8W, в долях W)

' --- точки триангуляции ---
Private PX() As Double, PY() As Double, NP As Long

' --- треугольники (только внутренние) ---
Private TA() As Long, TB() As Long, TC() As Long, NT As Long

' --- рёбра триангуляции: у точки p список соседей q с числом треугольников и узлом ---
Private EQ() As Long, ECnt() As Long, ENode() As Long, EDeg() As Long

' --- граф ---
Private GX() As Double, GY() As Double, GR() As Double, GK() As Long
Private GDeg() As Long, GAdj() As Long, GUsed() As Boolean, GLive() As Boolean, NG As Long, GCap As Long
Private GJW() As Double                       ' ширина у развилки (кэш)
Private PassA() As Long, PassB() As Long, PassN() As Long   ' пары сквозного прохода у развилки
Private CuspOn() As Boolean, CuspX() As Double, CuspY() As Double

' --- цепочки ---
Private ChI() As Variant, ChC() As Boolean, NCh As Long

' --- результат ---
Private OSeg() As Double, NOS As Long        ' (0..8, k): тип, x0,y0,x1,y1,x2,y2,x3,y3
Private OPathStart() As Long, OPathClosed() As Boolean, OPathGroup() As Long, NOP As Long
Private gGroup As Long                       ' номер текущей группы (буквы)

' --- острые узлы, поставленные Sharpen: FitPath обязан сделать в них излом ---
Private ShX() As Double, ShY() As Double, NSh As Long

'@@CORE_BEGIN
' =====================================================================
'  Ядро: кольца объекта (AX/AY/AS_/AN_) -> пути (OSeg/OPath*)
' =====================================================================
Public Sub CoreInitRings()
    NAP = 0: NAR = 0
    ReDim AX(0 To 1023): ReDim AY(0 To 1023)
    ReDim AS_(0 To 63): ReDim AN_(0 To 63)
End Sub

Public Sub CoreAddPoint(ByVal x As Double, ByVal y As Double)
    If NAP > 0 Then
        If Abs(AX(NAP - 1) - x) < 0.000001 And Abs(AY(NAP - 1) - y) < 0.000001 Then
            Exit Sub
        End If
    End If
    If NAP > UBound(AX) Then
        ReDim Preserve AX(0 To 2 * NAP): ReDim Preserve AY(0 To 2 * NAP)
    End If
    AX(NAP) = x: AY(NAP) = y: NAP = NAP + 1
End Sub

' закрыть текущее кольцо, начатое с индекса st
Public Sub CoreEndRing(ByVal st As Long)
    If NAP - st > 1 Then
        If Dist(AX(st), AY(st), AX(NAP - 1), AY(NAP - 1)) < 0.000001 Then
            NAP = NAP - 1
        End If
    End If
    If NAP - st >= 3 Then
        If NAR > UBound(AS_) Then
            ReDim Preserve AS_(0 To 2 * NAR): ReDim Preserve AN_(0 To 2 * NAR)
        End If
        AS_(NAR) = st: AN_(NAR) = NAP - st: NAR = NAR + 1
    Else
        NAP = st
    End If
End Sub

Public Function CoreRingStart() As Long
    CoreRingStart = NAP
End Function

Public Sub CoreSetParams(ByVal tool As Double, ByVal tol As Double)
    gTool = tool: gTol = tol
End Sub

Public Function CoreSegCount() As Long
    CoreSegCount = NOS
End Function

Public Function CorePathCount() As Long
    CorePathCount = NOP
End Function

Public Function CorePathStart(ByVal p As Long) As Long
    CorePathStart = OPathStart(p)
End Function

Public Function CorePathClosed(ByVal p As Long) As Boolean
    CorePathClosed = OPathClosed(p)
End Function

' номер группы (отдельной буквы / связки букв), к которой относится путь
Public Function CorePathGroup(ByVal p As Long) As Long
    CorePathGroup = OPathGroup(p)
End Function

Public Function CoreGroupCount() As Long
    CoreGroupCount = gGroup
End Function

Public Function CoreSeg(ByVal k As Long, ByVal i As Long) As Double
    CoreSeg = OSeg(k, i)
End Function

' Главная процедура ядра
Public Sub CoreRun()
    Dim r As Long, j As Long, g As Long, k As Long
    Dim depth() As Long, parent() As Long, area() As Double, best As Double
    Dim gid() As Long
    Dim bx0() As Double, by0() As Double, bx1() As Double, by1() As Double
    Dim rbx0() As Double, rby0() As Double, rbx1() As Double, rby1() As Double
    Dim lst() As Long, nl As Long, nPolys As Long, outer() As Long

    NOS = 0: NOP = 0: gGroup = 0
    ReDim OSeg(0 To 8, 0 To 255)
    ReDim OPathStart(0 To 63): ReDim OPathClosed(0 To 63): ReDim OPathGroup(0 To 63)
    If NAR = 0 Then
        Exit Sub
    End If

    ' --- вложенность колец: r внутри q, если рамка r внутри рамки q, q больше
    '     и 7 из 8 проб r лежат внутри q
    ReDim depth(0 To NAR - 1): ReDim parent(0 To NAR - 1): ReDim area(0 To NAR - 1)
    ReDim rbx0(0 To NAR - 1): ReDim rby0(0 To NAR - 1): ReDim rbx1(0 To NAR - 1): ReDim rby1(0 To NAR - 1)
    For r = 0 To NAR - 1
        area(r) = Abs(RingAreaA(r))
        RingBBoxA r, rbx0(r), rby0(r), rbx1(r), rby1(r)
    Next r
    For r = 0 To NAR - 1
        depth(r) = 0
        For j = 0 To NAR - 1
            If j <> r And area(j) > area(r) Then
                If rbx0(j) <= rbx0(r) And rbx1(j) >= rbx1(r) And rby0(j) <= rby0(r) And rby1(j) >= rby1(r) Then
                    If RingInRingA(r, j) Then
                        depth(r) = depth(r) + 1
                    End If
                End If
            End If
        Next j
    Next r
    For r = 0 To NAR - 1
        parent(r) = -1
        If depth(r) Mod 2 = 1 Then
            best = BIG
            For j = 0 To NAR - 1
                If j <> r And depth(j) = depth(r) - 1 And area(j) > area(r) And area(j) < best Then
                    If rbx0(j) <= rbx0(r) And rbx1(j) >= rbx1(r) And rby0(j) <= rby0(r) And rby1(j) >= rby1(r) Then
                        If RingInRingA(r, j) Then
                            best = area(j): parent(r) = j
                        End If
                    End If
                End If
            Next j
        End If
    Next r

    ' --- полигоны = внешние кольца; рамки
    ReDim outer(0 To NAR - 1): nPolys = 0
    ReDim bx0(0 To NAR - 1): ReDim by0(0 To NAR - 1): ReDim bx1(0 To NAR - 1): ReDim by1(0 To NAR - 1)
    For r = 0 To NAR - 1
        If depth(r) Mod 2 = 0 Then
            outer(nPolys) = r
            bx0(nPolys) = rbx0(r): by0(nPolys) = rby0(r): bx1(nPolys) = rbx1(r): by1(nPolys) = rby1(r)
            nPolys = nPolys + 1
        End If
    Next r
    If nPolys = 0 Then
        Exit Sub
    End If

    ' --- группы: полигоны с пересекающимися рамками (union-find)
    ReDim gid(0 To nPolys - 1)
    For k = 0 To nPolys - 1: gid(k) = k: Next k
    For k = 0 To nPolys - 1
        For j = k + 1 To nPolys - 1
            If bx0(k) <= bx1(j) And bx0(j) <= bx1(k) And by0(k) <= by1(j) And by0(j) <= by1(k) Then
                UFUnion gid, k, j
            End If
        Next j
    Next k

    ReDim lst(0 To nPolys - 1)
    For g = 0 To nPolys - 1
        If UFFind(gid, g) = g Then
            nl = 0
            For k = 0 To nPolys - 1
                If UFFind(gid, k) = g Then
                    lst(nl) = outer(k): nl = nl + 1
                End If
            Next k
            LoadGroup lst, nl, parent
            BuildBins
            RunGroup
            gGroup = gGroup + 1
        End If
    Next g
End Sub

Private Function UFFind(gid() As Long, ByVal i As Long) As Long
    Do While gid(i) <> i
        gid(i) = gid(gid(i))
        i = gid(i)
    Loop
    UFFind = i
End Function

Private Sub UFUnion(gid() As Long, ByVal a As Long, ByVal b As Long)
    Dim ra As Long, rb As Long
    ra = UFFind(gid, a): rb = UFFind(gid, b)
    If ra <> rb Then
        gid(ra) = rb
    End If
End Sub

Private Function RingAreaA(ByVal r As Long) As Double
    Dim i As Long, j As Long, s As Double, n As Long, b As Long
    b = AS_(r): n = AN_(r)
    For i = 0 To n - 1
        j = (i + 1) Mod n
        s = s + AX(b + i) * AY(b + j) - AX(b + j) * AY(b + i)
    Next i
    RingAreaA = s / 2
End Function

Private Sub RingBBoxA(ByVal r As Long, x0 As Double, y0 As Double, x1 As Double, y1 As Double)
    Dim i As Long
    x0 = BIG: y0 = BIG: x1 = -BIG: y1 = -BIG
    For i = AS_(r) To AS_(r) + AN_(r) - 1
        If AX(i) < x0 Then
            x0 = AX(i)
        End If
        If AX(i) > x1 Then
            x1 = AX(i)
        End If
        If AY(i) < y0 Then
            y0 = AY(i)
        End If
        If AY(i) > y1 Then
            y1 = AY(i)
        End If
    Next i
End Sub

Private Function PointInRingA(ByVal r As Long, ByVal x As Double, ByVal y As Double) As Boolean
    Dim i As Long, j As Long, n As Long, b As Long, c As Boolean
    b = AS_(r): n = AN_(r): j = n - 1
    For i = 0 To n - 1
        If (AY(b + i) > y) <> (AY(b + j) > y) Then
            If x < (AX(b + j) - AX(b + i)) * (y - AY(b + i)) / (AY(b + j) - AY(b + i)) + AX(b + i) Then
                c = Not c
            End If
        End If
        j = i
    Next i
    PointInRingA = c
End Function

' кольцо r лежит внутри кольца q (7 из 8 проб)
Private Function RingInRingA(ByVal r As Long, ByVal q As Long) As Boolean
    Dim k As Long, idx As Long, cnt As Long
    For k = 0 To 7
        idx = AS_(r) + (k * AN_(r)) \ 8
        If PointInRingA(q, AX(idx), AY(idx)) Then
            cnt = cnt + 1
        End If
    Next k
    RingInRingA = (cnt >= 7)
End Function

' загрузить группу полигонов (внешние кольца lst + их дырки) в RX/RY
Private Sub LoadGroup(lst() As Long, ByVal nl As Long, parent() As Long)
    Dim i As Long, k As Long, j As Long, tot As Long, nr_ As Long, r As Long, p As Long
    tot = 0: nr_ = 0
    For i = 0 To nl - 1
        tot = tot + AN_(lst(i)): nr_ = nr_ + 1
        For j = 0 To NAR - 1
            If parent(j) = lst(i) Then
                tot = tot + AN_(j): nr_ = nr_ + 1
            End If
        Next j
    Next i
    ReDim RX(0 To tot - 1): ReDim RY(0 To tot - 1)
    ReDim RS(0 To nr_ - 1): ReDim RN(0 To nr_ - 1): ReDim RPoly(0 To nr_ - 1)
    ReDim PolyR0(0 To nl - 1): ReDim PolyRN(0 To nl - 1)
    ReDim PBx0(0 To nl - 1): ReDim PBy0(0 To nl - 1): ReDim PBx1(0 To nl - 1): ReDim PBy1(0 To nl - 1)
    NPoly = nl: NR = 0: tot = 0
    For p = 0 To nl - 1
        PolyR0(p) = NR: PolyRN(p) = 0
        For j = -1 To NAR - 1
            If j = -1 Then
                r = lst(p)
            ElseIf parent(j) = lst(p) Then
                r = j
            Else
                r = -1
            End If
            If r >= 0 Then
                RS(NR) = tot: RN(NR) = AN_(r): RPoly(NR) = p
                For k = 0 To AN_(r) - 1
                    RX(tot) = AX(AS_(r) + k): RY(tot) = AY(AS_(r) + k): tot = tot + 1
                Next k
                NR = NR + 1: PolyRN(p) = PolyRN(p) + 1
            End If
        Next j
        RingBBoxA lst(p), PBx0(p), PBy0(p), PBx1(p), PBy1(p)
    Next p
End Sub

Private Sub BuildBins()
    Dim r As Long, i As Long, j As Long, b As Long, n As Long, k As Long, k0 As Long, k1 As Long
    Dim ya As Double, yb As Double, ymin As Double, ymax As Double, tot As Long, pos() As Long
    ymin = BIG: ymax = -BIG
    For i = 0 To UBound(RY)
        If RY(i) < ymin Then
            ymin = RY(i)
        End If
        If RY(i) > ymax Then
            ymax = RY(i)
        End If
    Next i
    BinY0 = ymin
    BinDY = (ymax - ymin) / NBIN
    If BinDY <= 0 Then
        BinDY = 1
    End If
    ReDim BinS(0 To NBIN - 1): ReDim BinN(0 To NBIN - 1): ReDim pos(0 To NBIN - 1)
    For k = 0 To 1
        For r = 0 To NR - 1
            b = RS(r): n = RN(r)
            For i = 0 To n - 1
                j = (i + 1) Mod n
                ya = RY(b + i): yb = RY(b + j)
                If ya > yb Then
                    k0 = BinOf(yb): k1 = BinOf(ya)
                Else
                    k0 = BinOf(ya): k1 = BinOf(yb)
                End If
                For tot = k0 To k1
                    If k = 0 Then
                        BinN(tot) = BinN(tot) + 1
                    Else
                        BinA(pos(tot)) = b + i: BinB(pos(tot)) = b + j: BinP(pos(tot)) = RPoly(r)
                        pos(tot) = pos(tot) + 1
                    End If
                Next tot
            Next i
        Next r
        If k = 0 Then
            tot = 0
            For i = 0 To NBIN - 1
                BinS(i) = tot: pos(i) = tot: tot = tot + BinN(i)
            Next i
            ReDim BinA(0 To tot): ReDim BinB(0 To tot): ReDim BinP(0 To tot)
        End If
    Next k
End Sub

Private Function BinOf(ByVal y As Double) As Long
    Dim k As Long
    k = Int((y - BinY0) / BinDY)
    If k < 0 Then
        k = 0
    End If
    If k > NBIN - 1 Then
        k = NBIN - 1
    End If
    BinOf = k
End Function

Private Function InPoly(ByVal p As Long, ByVal x As Double, ByVal y As Double) As Boolean
    Dim k As Long, e As Long, a As Long, b As Long, c As Boolean
    If x < PBx0(p) Or x > PBx1(p) Or y < PBy0(p) Or y > PBy1(p) Then
        Exit Function
    End If
    k = BinOf(y)
    For e = BinS(k) To BinS(k) + BinN(k) - 1
        If BinP(e) = p Then
            a = BinA(e): b = BinB(e)
            If (RY(a) > y) <> (RY(b) > y) Then
                If x < (RX(b) - RX(a)) * (y - RY(a)) / (RY(b) - RY(a)) + RX(a) Then
                    c = Not c
                End If
            End If
        End If
    Next e
    InPoly = c
End Function

' точка внутри объединения полигонов группы (кроме skip)
Private Function InUnion(ByVal x As Double, ByVal y As Double, ByVal skip As Long) As Boolean
    Dim p As Long
    For p = 0 To NPoly - 1
        If p <> skip Then
            If InPoly(p, x, y) Then
                InUnion = True: Exit Function
            End If
        End If
    Next p
End Function

' =====================================================================
'  Группа
' =====================================================================
Private Sub RunGroup()
    Dim it As Long, chg As Boolean, area As Double, per As Double, p As Long, r As Long, s0 As Long
    Dim gx0 As Double, gy0 As Double, gx1 As Double, gy1 As Double
    For r = 0 To NR - 1
        If r = PolyR0(RPoly(r)) Then
            area = area + Abs(RingArea(r))
        Else
            area = area - Abs(RingArea(r))
        End If
        per = per + RingLen(r)
    Next r
    If per <= 0 Or area <= 0 Then
        Exit Sub
    End If
    gW = 2 * area / per
    gH = gW / P_HSTEP
    s0 = NOP
    DropTinyHoles

    NSh = 0
    ReDim ShX(0 To 15): ReDim ShY(0 To 15)
    Resample gH
    FindBCorners
    If NP >= 6 Then
        Delaunay
        BuildGraph
        DissolveJunctions
        For it = 1 To 4
            chg = PruneLeaves(P_PRUNE)
            DissolveJunctions
            If Not chg Then
                Exit For
            End If
        Next it
        MergeJunctions P_JMERGE
        MergeBridges P_MERGEANG
        DropSmall
        BuildChains P_MERGEANG
        ProcessChains
        JoinPaths s0, 0.05 * gW
    End If

    ' маленький комочек (точка над i, ё, точка) — вместо линий метка в центре для сверловки
    gx0 = BIG: gy0 = BIG: gx1 = -BIG: gy1 = -BIG
    For p = 0 To NPoly - 1
        If PBx0(p) < gx0 Then
            gx0 = PBx0(p)
        End If
        If PBy0(p) < gy0 Then
            gy0 = PBy0(p)
        End If
        If PBx1(p) > gx1 Then
            gx1 = PBx1(p)
        End If
        If PBy1(p) > gy1 Then
            gy1 = PBy1(p)
        End If
    Next p
    If (gx1 - gx0 < 2.5 * gW And gy1 - gy0 < 2.5 * gW) Or NOP = s0 Then
        If NOP > s0 Then
            NOS = OPathStart(s0): NOP = s0
        End If
        If gx1 - gx0 < 4 * gW And gy1 - gy0 < 4 * gW Then
            EmitDot (gx0 + gx1) / 2, (gy0 + gy1) / 2
        End If
    End If
End Sub

' крошечные дырки (счётчик меньше самой фрезы: треугольник в рукописной «т») — не контур
Private Sub DropTinyHoles()
    Dim r As Long, p As Long, keepR() As Boolean, anyDrop As Boolean, nNew As Long, r0 As Long, rn0 As Long, lim As Double
    lim = 0.3 * gW * gW
    ReDim keepR(0 To NR - 1)
    For r = 0 To NR - 1
        keepR(r) = True
        If r <> PolyR0(RPoly(r)) Then
            If Abs(RingArea(r)) < lim Then
                keepR(r) = False: anyDrop = True
            End If
        End If
    Next r
    If Not anyDrop Then
        Exit Sub
    End If
    nNew = 0
    For p = 0 To NPoly - 1
        r0 = nNew: rn0 = 0
        For r = PolyR0(p) To PolyR0(p) + PolyRN(p) - 1
            If keepR(r) Then
                RS(nNew) = RS(r): RN(nNew) = RN(r): RPoly(nNew) = RPoly(r)
                nNew = nNew + 1: rn0 = rn0 + 1
            End If
        Next r
        PolyR0(p) = r0: PolyRN(p) = rn0
    Next p
    NR = nNew
    BuildBins
End Sub

Private Function RingArea(ByVal r As Long) As Double
    Dim i As Long, j As Long, s As Double, n As Long, b As Long
    b = RS(r): n = RN(r)
    For i = 0 To n - 1
        j = (i + 1) Mod n
        s = s + RX(b + i) * RY(b + j) - RX(b + j) * RY(b + i)
    Next i
    RingArea = s / 2
End Function

Private Function RingLen(ByVal r As Long) As Double
    Dim i As Long, j As Long, s As Double, n As Long, b As Long
    b = RS(r): n = RN(r)
    For i = 0 To n - 1
        j = (i + 1) Mod n
        s = s + Dist(RX(b + i), RY(b + i), RX(b + j), RY(b + j))
    Next i
    RingLen = s
End Function

Private Sub EmitDot(ByVal cx As Double, ByVal cy As Double)
    Dim kk As Double
    AddPath True
    kk = 0.5522847498 * P_DOT
    EmitSeg 1, cx + P_DOT, cy, cx + P_DOT, cy + kk, cx + kk, cy + P_DOT, cx, cy + P_DOT
    EmitSeg 1, cx, cy + P_DOT, cx - kk, cy + P_DOT, cx - P_DOT, cy + kk, cx - P_DOT, cy
    EmitSeg 1, cx - P_DOT, cy, cx - P_DOT, cy - kk, cx - kk, cy - P_DOT, cx, cy - P_DOT
    EmitSeg 1, cx, cy - P_DOT, cx + kk, cy - P_DOT, cx + P_DOT, cy - kk, cx + P_DOT, cy
End Sub

' ресэмплинг колец с шагом h; точки внутри чужих полигонов выкидываются
Private Sub Resample(ByVal h As Double)
    Dim r As Long, i As Long, j As Long, k As Long, n As Long, m As Long, b As Long
    Dim L As Double, stp As Double, pos As Double, target As Double, s As Double, u As Double
    Dim seglen() As Double, x As Double, y As Double, keep As Boolean
    NP = 0
    ReDim PX(0 To 1023): ReDim PY(0 To 1023)
    For r = 0 To NR - 1
        b = RS(r): m = RN(r)
        ReDim seglen(0 To m - 1)
        L = 0
        For k = 0 To m - 1
            j = (k + 1) Mod m
            seglen(k) = Dist(RX(b + k), RY(b + k), RX(b + j), RY(b + j))
            L = L + seglen(k)
        Next k
        n = -Int(-L / h)
        If n < 8 Then
            n = 8
        End If
        stp = L / n
        k = 0: pos = 0
        For i = 0 To n - 1
            target = i * stp
            Do While k < m - 1
                If pos + seglen(k) >= target Then
                    Exit Do
                End If
                pos = pos + seglen(k): k = k + 1
            Loop
            s = seglen(k)
            If s = 0 Then
                u = 0
            Else
                u = (target - pos) / s
            End If
            j = (k + 1) Mod m
            x = RX(b + k) + (RX(b + j) - RX(b + k)) * u
            y = RY(b + k) + (RY(b + j) - RY(b + k)) * u
            keep = True
            If NPoly > 1 Then
                If InUnion(x, y, RPoly(r)) Then
                    keep = False
                End If
            End If
            If keep Then
                If NP > UBound(PX) Then
                    ReDim Preserve PX(0 To 2 * NP): ReDim Preserve PY(0 To 2 * NP)
                End If
                PX(NP) = x: PY(NP) = y: NP = NP + 1
            End If
        Next i
    Next r
End Sub

' ---------------------------------------------------------------------
'  Делоне (Bowyer–Watson с сортировкой по X, алгоритм Бурка)
' ---------------------------------------------------------------------
Private Sub Delaunay()
    Dim n As Long, i As Long, k As Long, e As Long, f As Long
    Dim ord() As Long, xs() As Double, ys() As Double
    Dim xmin As Double, xmax As Double, ymin As Double, ymax As Double, dm As Double, xm As Double, ym As Double
    Dim t0() As Long, t1() As Long, t2() As Long, ccx() As Double, ccy() As Double, ccr() As Double
    Dim ntri As Long, cap As Long
    Dim e0() As Long, e1() As Long, necnt As Long, ecap As Long, dup() As Boolean
    Dim x As Double, y As Double, dx As Double, va As Long, vb As Long, vc As Long
    Dim nDone As Long, tl As Long, td As Double

    n = NP
    ReDim ord(0 To n - 1)
    For i = 0 To n - 1: ord(i) = i: Next i
    SortIdx ord, 0, n - 1
    ReDim xs(0 To n + 2): ReDim ys(0 To n + 2)
    xmin = BIG: ymin = BIG: xmax = -BIG: ymax = -BIG
    For i = 0 To n - 1
        xs(i) = PX(ord(i)): ys(i) = PY(ord(i))
        If xs(i) < xmin Then
            xmin = xs(i)
        End If
        If xs(i) > xmax Then
            xmax = xs(i)
        End If
        If ys(i) < ymin Then
            ymin = ys(i)
        End If
        If ys(i) > ymax Then
            ymax = ys(i)
        End If
    Next i
    dm = xmax - xmin
    If ymax - ymin > dm Then
        dm = ymax - ymin
    End If
    dm = dm * 20
    xm = (xmin + xmax) / 2: ym = (ymin + ymax) / 2
    xs(n) = xm - dm: ys(n) = ym - dm
    xs(n + 1) = xm: ys(n + 1) = ym + dm
    xs(n + 2) = xm + dm: ys(n + 2) = ym - dm

    cap = 4 * n + 16
    ReDim t0(0 To cap): ReDim t1(0 To cap): ReDim t2(0 To cap)
    ReDim ccx(0 To cap): ReDim ccy(0 To cap): ReDim ccr(0 To cap)
    t0(0) = n: t1(0) = n + 1: t2(0) = n + 2
    Circum xs, ys, n, n + 1, n + 2, ccx(0), ccy(0), ccr(0)
    ntri = 1
    ecap = 256
    ReDim e0(0 To ecap): ReDim e1(0 To ecap)

    For i = 0 To n - 1
        x = xs(i): y = ys(i)
        necnt = 0
        k = nDone
        Do While k < ntri
            dx = x - ccx(k)
            If dx > 0 And dx * dx > ccr(k) Then
                ' треугольник завершён (точки идут по X) — переносим в начало
                If k <> nDone Then
                    tl = t0(k): t0(k) = t0(nDone): t0(nDone) = tl
                    tl = t1(k): t1(k) = t1(nDone): t1(nDone) = tl
                    tl = t2(k): t2(k) = t2(nDone): t2(nDone) = tl
                    td = ccx(k): ccx(k) = ccx(nDone): ccx(nDone) = td
                    td = ccy(k): ccy(k) = ccy(nDone): ccy(nDone) = td
                    td = ccr(k): ccr(k) = ccr(nDone): ccr(nDone) = td
                End If
                nDone = nDone + 1
                k = k + 1
            ElseIf dx * dx + (y - ccy(k)) * (y - ccy(k)) <= ccr(k) Then
                If necnt + 3 > ecap Then
                    ecap = ecap * 2
                    ReDim Preserve e0(0 To ecap): ReDim Preserve e1(0 To ecap)
                End If
                e0(necnt) = t0(k): e1(necnt) = t1(k): necnt = necnt + 1
                e0(necnt) = t1(k): e1(necnt) = t2(k): necnt = necnt + 1
                e0(necnt) = t2(k): e1(necnt) = t0(k): necnt = necnt + 1
                ntri = ntri - 1
                t0(k) = t0(ntri): t1(k) = t1(ntri): t2(k) = t2(ntri)
                ccx(k) = ccx(ntri): ccy(k) = ccy(ntri): ccr(k) = ccr(ntri)
            Else
                k = k + 1
            End If
        Loop
        ReDim dup(0 To necnt)
        For e = 0 To necnt - 2
            For f = e + 1 To necnt - 1
                If (e0(e) = e1(f) And e1(e) = e0(f)) Or (e0(e) = e0(f) And e1(e) = e1(f)) Then
                    dup(e) = True: dup(f) = True
                End If
            Next f
        Next e
        For e = 0 To necnt - 1
            If Not dup(e) Then
                If ntri > cap Then
                    cap = cap * 2
                    ReDim Preserve t0(0 To cap): ReDim Preserve t1(0 To cap): ReDim Preserve t2(0 To cap)
                    ReDim Preserve ccx(0 To cap): ReDim Preserve ccy(0 To cap)
                    ReDim Preserve ccr(0 To cap)
                End If
                t0(ntri) = e0(e): t1(ntri) = e1(e): t2(ntri) = i
                Circum xs, ys, t0(ntri), t1(ntri), t2(ntri), ccx(ntri), ccy(ntri), ccr(ntri)
                ntri = ntri + 1
            End If
        Next e
    Next i

    ' только треугольники без вершин супертреугольника и внутри фигуры
    NT = 0
    ReDim TA(0 To ntri): ReDim TB(0 To ntri): ReDim TC(0 To ntri)
    For k = 0 To ntri - 1
        If t0(k) < n And t1(k) < n And t2(k) < n Then
            va = ord(t0(k)): vb = ord(t1(k)): vc = ord(t2(k))
            If InUnion((PX(va) + PX(vb) + PX(vc)) / 3, (PY(va) + PY(vb) + PY(vc)) / 3, -1) Then
                TA(NT) = va: TB(NT) = vb: TC(NT) = vc: NT = NT + 1
            End If
        End If
    Next k
End Sub

Private Sub Circum(xs() As Double, ys() As Double, ByVal a As Long, ByVal b As Long, ByVal c As Long, _
                   ux As Double, uy As Double, r2 As Double)
    Dim ax_ As Double, ay_ As Double, bx_ As Double, by_ As Double, cx_ As Double, cy_ As Double, d As Double
    ax_ = xs(a): ay_ = ys(a): bx_ = xs(b): by_ = ys(b): cx_ = xs(c): cy_ = ys(c)
    d = 2 * (ax_ * (by_ - cy_) + bx_ * (cy_ - ay_) + cx_ * (ay_ - by_))
    If Abs(d) < 1E-12 Then
        ux = ax_: uy = ay_: r2 = BIG
        Exit Sub
    End If
    ux = ((ax_ * ax_ + ay_ * ay_) * (by_ - cy_) + (bx_ * bx_ + by_ * by_) * (cy_ - ay_) + (cx_ * cx_ + cy_ * cy_) * (ay_ - by_)) / d
    uy = ((ax_ * ax_ + ay_ * ay_) * (cx_ - bx_) + (bx_ * bx_ + by_ * by_) * (ax_ - cx_) + (cx_ * cx_ + cy_ * cy_) * (bx_ - ax_)) / d
    r2 = (ax_ - ux) * (ax_ - ux) + (ay_ - uy) * (ay_ - uy)
End Sub

Private Sub SortIdx(ord() As Long, ByVal lo As Long, ByVal hi As Long)
    Dim i As Long, j As Long, pv As Double, t As Long
    Do While lo < hi
        i = lo: j = hi
        pv = PX(ord((lo + hi) \ 2))
        Do While i <= j
            Do While PX(ord(i)) < pv
                i = i + 1
            Loop
            Do While PX(ord(j)) > pv
                j = j - 1
            Loop
            If i <= j Then
                t = ord(i): ord(i) = ord(j): ord(j) = t
                i = i + 1: j = j - 1
            End If
        Loop
        If j - lo < hi - i Then
            If lo < j Then
                SortIdx ord, lo, j
            End If
            lo = i
        Else
            If i < hi Then
                SortIdx ord, i, hi
            End If
            hi = j
        End If
    Loop
End Sub

' ---------------------------------------------------------------------
'  Граф хордальной оси
' ---------------------------------------------------------------------
' найти/создать запись ребра (a,b) у точки a; вернуть индекс слота
Private Function ESlot(ByVal a As Long, ByVal b As Long) As Long
    Dim k As Long
    For k = 0 To EDeg(a) - 1
        If EQ(k, a) = b Then
            ESlot = k: Exit Function
        End If
    Next k
    If EDeg(a) >= MAXE Then
        ESlot = -1: Exit Function
    End If
    k = EDeg(a)
    EQ(k, a) = b: ECnt(k, a) = 0: ENode(k, a) = -1
    EDeg(a) = k + 1
    ESlot = k
End Function

Private Sub ECount(ByVal a As Long, ByVal b As Long)
    Dim k As Long, lo As Long, hi As Long
    If a < b Then
        lo = a: hi = b
    Else
        lo = b: hi = a
    End If
    k = ESlot(lo, hi)
    If k >= 0 Then
        ECnt(k, lo) = ECnt(k, lo) + 1
    End If
End Sub

Private Function EInner(ByVal a As Long, ByVal b As Long) As Boolean
    Dim k As Long, lo As Long, hi As Long
    If a < b Then
        lo = a: hi = b
    Else
        lo = b: hi = a
    End If
    k = ESlot(lo, hi)
    If k >= 0 Then
        EInner = (ECnt(k, lo) = 2)
    End If
End Function

Private Function EdgeNode(ByVal a As Long, ByVal b As Long) As Long
    Dim k As Long, lo As Long, hi As Long
    If a < b Then
        lo = a: hi = b
    Else
        lo = b: hi = a
    End If
    k = ESlot(lo, hi)
    If ENode(k, lo) < 0 Then
        ENode(k, lo) = NewNode((PX(lo) + PX(hi)) / 2, (PY(lo) + PY(hi)) / 2, Dist(PX(lo), PY(lo), PX(hi), PY(hi)), 0)
    End If
    EdgeNode = ENode(k, lo)
End Function

Private Function NewNode(ByVal x As Double, ByVal y As Double, ByVal r As Double, ByVal kind As Long) As Long
    If NG > GCap Then
        GCap = GCap * 2
        ReDim Preserve GX(0 To GCap): ReDim Preserve GY(0 To GCap): ReDim Preserve GR(0 To GCap)
        ReDim Preserve GK(0 To GCap): ReDim Preserve GDeg(0 To GCap): ReDim Preserve GLive(0 To GCap)
        ReDim Preserve GAdj(0 To MAXD - 1, 0 To GCap): ReDim Preserve GUsed(0 To MAXD - 1, 0 To GCap)
    End If
    GX(NG) = x: GY(NG) = y: GR(NG) = r: GK(NG) = kind: GDeg(NG) = 0: GLive(NG) = True
    NewNode = NG
    NG = NG + 1
End Function

Private Function Linked(ByVal i As Long, ByVal j As Long) As Boolean
    Dim k As Long
    For k = 0 To GDeg(i) - 1
        If GAdj(k, i) = j Then
            Linked = True: Exit Function
        End If
    Next k
End Function

Private Sub LinkN(ByVal i As Long, ByVal j As Long)
    If i = j Then
        Exit Sub
    End If
    If Linked(i, j) Then
        Exit Sub
    End If
    If GDeg(i) >= MAXD Or GDeg(j) >= MAXD Then
        Exit Sub
    End If
    GAdj(GDeg(i), i) = j: GUsed(GDeg(i), i) = False: GDeg(i) = GDeg(i) + 1
    GAdj(GDeg(j), j) = i: GUsed(GDeg(j), j) = False: GDeg(j) = GDeg(j) + 1
End Sub

Private Sub DelHalf(ByVal i As Long, ByVal j As Long)
    Dim k As Long, last As Long
    For k = 0 To GDeg(i) - 1
        If GAdj(k, i) = j Then
            last = GDeg(i) - 1
            GAdj(k, i) = GAdj(last, i): GUsed(k, i) = GUsed(last, i)
            GDeg(i) = last
            Exit Sub
        End If
    Next k
End Sub

Private Sub UnlinkN(ByVal i As Long, ByVal j As Long)
    DelHalf i, j
    DelHalf j, i
End Sub

Private Sub KillNode(ByVal i As Long)
    Do While GDeg(i) > 0
        UnlinkN i, GAdj(0, i)
    Loop
    GLive(i) = False
End Sub

Private Function InG(ByVal i As Long) As Boolean
    If GLive(i) Then
        InG = (GDeg(i) > 0)
    End If
End Function

Private Sub BuildGraph()
    Dim t As Long, va As Long, vb As Long, vc As Long, ns(0 To 2) As Long, nn As Long, j As Long
    ReDim EQ(0 To MAXE - 1, 0 To NP): ReDim ECnt(0 To MAXE - 1, 0 To NP)
    ReDim ENode(0 To MAXE - 1, 0 To NP): ReDim EDeg(0 To NP)
    NG = 0: GCap = 1023
    ReDim GX(0 To GCap): ReDim GY(0 To GCap): ReDim GR(0 To GCap)
    ReDim GK(0 To GCap): ReDim GDeg(0 To GCap): ReDim GLive(0 To GCap)
    ReDim GAdj(0 To MAXD - 1, 0 To GCap): ReDim GUsed(0 To MAXD - 1, 0 To GCap)
    ' ребро внутреннее, если его делят два внутренних треугольника
    For t = 0 To NT - 1
        ECount TA(t), TB(t): ECount TB(t), TC(t): ECount TC(t), TA(t)
    Next t
    For t = 0 To NT - 1
        va = TA(t): vb = TB(t): vc = TC(t)
        nn = 0
        If EInner(va, vb) Then
            ns(nn) = EdgeNode(va, vb): nn = nn + 1
        End If
        If EInner(vb, vc) Then
            ns(nn) = EdgeNode(vb, vc): nn = nn + 1
        End If
        If EInner(vc, va) Then
            ns(nn) = EdgeNode(vc, va): nn = nn + 1
        End If
        If nn = 2 Then
            LinkN ns(0), ns(1)
        ElseIf nn = 3 Then
            j = NewNode((GX(ns(0)) + GX(ns(1)) + GX(ns(2))) / 3, (GY(ns(0)) + GY(ns(1)) + GY(ns(2))) / 3, BIG, 1)
            LinkN j, ns(0): LinkN j, ns(1): LinkN j, ns(2)
        End If
    Next t
End Sub

Private Sub DissolveJunctions()
    Dim i As Long, a As Long, b As Long
    For i = 0 To NG - 1
        If GLive(i) And GK(i) = 1 Then
            If GDeg(i) = 2 Then
                a = GAdj(0, i): b = GAdj(1, i)
                KillNode i
                LinkN a, b
            ElseIf GDeg(i) < 2 Then
                KillNode i
            End If
        End If
    Next i
End Sub

' сосед cur, отличный от prv (для узлов степени 1-2), иначе -1
Private Function OtherNb(ByVal cur As Long, ByVal prv As Long) As Long
    Dim k As Long
    OtherNb = -1
    If GDeg(cur) > 2 Then
        Exit Function
    End If
    For k = 0 To GDeg(cur) - 1
        If GAdj(k, cur) <> prv Then
            OtherNb = GAdj(k, cur): Exit Function
        End If
    Next k
End Function

' ширина штриха у развилки: максимум хорд в ближайших узлах веток
Private Function JWidth(ByVal j As Long) As Double
    Dim k As Long, prv As Long, cur As Long, nxt As Long, cnt As Long, best As Double
    For k = 0 To GDeg(j) - 1
        prv = j: cur = GAdj(k, j): cnt = 0
        Do
            If GK(cur) = 0 Then
                If GR(cur) > best Then
                    best = GR(cur)
                End If
            End If
            cnt = cnt + 1
            If cnt >= 6 Or GDeg(cur) <> 2 Then
                Exit Do
            End If
            nxt = OtherNb(cur, prv)
            If nxt < 0 Then
                Exit Do
            End If
            prv = cur: cur = nxt
        Loop
    Next k
    If best <= 0 Then
        best = gW
    End If
    JWidth = best
End Function

' направление от i вдоль ветки через j на длину dlen
Private Sub DirFrom(ByVal i As Long, ByVal j As Long, ByVal dlen As Double, dx As Double, dy As Double)
    Dim prv As Long, cur As Long, nxt As Long, L As Double
    prv = i: cur = j: L = Dist(GX(i), GY(i), GX(j), GY(j))
    Do While L < dlen
        If GDeg(cur) <> 2 Then
            Exit Do
        End If
        nxt = OtherNb(cur, prv)
        If nxt < 0 Then
            Exit Do
        End If
        prv = cur: cur = nxt
        L = L + Dist(GX(prv), GY(prv), GX(cur), GY(cur))
    Loop
    dx = GX(cur) - GX(i): dy = GY(cur) - GY(i)
    Norm dx, dy
End Sub

Private Function PruneLeaves(ByVal kprune As Double) As Boolean
    Dim lf As Long, prv As Long, cur As Long, nxt As Long, L As Double, reach As Double
    Dim path() As Long, npth As Long, k As Long, D As Double, dx As Double, dy As Double, a As Long
    Dim leaves() As Long, nl As Long, i As Long, jw As Double, keep As Boolean
    Dim dlx As Double, dly As Double, coll As Boolean, body As Boolean, s As Double, first As Double, hasFirst As Boolean
    Dim o As Long, m As Long, bdx As Double, bdy As Double, fX As Double, fY As Double, nX As Double, nY As Double
    Dim mX As Double, mY As Double, rr As Double
    Dim kq As Long, tmpL As Long
    ReDim leaves(0 To NG): ReDim path(0 To NG)
    For i = 0 To NG - 1
        If GLive(i) And GDeg(i) = 1 Then
            leaves(nl) = i: nl = nl + 1
        End If
    Next i
    ' порядок обхода — по координатам (снизу вверх, слева направо), как в прототипе:
    ' итог зависит от порядка, а номера узлов в прототипе другие
    For i = 1 To nl - 1
        tmpL = leaves(i): kq = i - 1
        Do While kq >= 0
            If Not LeafBefore(tmpL, leaves(kq)) Then
                Exit Do
            End If
            leaves(kq + 1) = leaves(kq): kq = kq - 1
        Loop
        leaves(kq + 1) = tmpL
    Next i
    For i = 0 To nl - 1
        lf = leaves(i)
        If GLive(lf) And GDeg(lf) = 1 Then
            npth = 0: path(npth) = lf: npth = npth + 1
            L = 0: prv = -1: cur = lf
            Do
                nxt = OtherNb(cur, prv)
                If nxt < 0 Then
                    Exit Do
                End If
                L = L + Dist(GX(cur), GY(cur), GX(nxt), GY(nxt))
                prv = cur: cur = nxt
                path(npth) = cur: npth = npth + 1
                If GDeg(cur) <> 2 Then
                    Exit Do
                End If
            Loop
            If GDeg(cur) >= 3 Then
                k = npth - 1
                If k > 3 Then
                    k = 3
                End If
                a = path(k)
                dx = GX(lf) - GX(a): dy = GY(lf) - GY(a)
                reach = L
                If Norm(dx, dy) Then
                    D = RayUnion(GX(lf), GY(lf), dx, dy)
                    If D < BIG Then
                        reach = reach + D
                    End If
                End If
                jw = JWidth(cur)
                If reach < kprune * jw Then
                    keep = False
                    ' хвост, продолжающий другую ветку по прямой (ножка ц идёт дальше за дугу,
                    ' которая влилась в неё по касательной) — не усик: усик всегда торчит в сторону
                    ' (биссектриса угла, скругление торца)
                    If reach >= 0.6 * jw And L >= 0.8 * gW And npth >= 2 Then
                        ' направление хвоста — как у остальных веток (с обрезкой загиба торца)
                        BranchDir cur, path(npth - 2), jw, dlx, dly, fX, fY, nX, nY, mX, mY, rr
                        coll = False
                        For m = 0 To GDeg(cur) - 1
                            o = GAdj(m, cur)
                            If o <> path(npth - 2) Then
                                BranchDir cur, o, jw, bdx, bdy, fX, fY, nX, nY, mX, mY, rr
                                If bdx * dlx + bdy * dly < -0.75 Then
                                    coll = True
                                End If
                            End If
                        Next m
                        ' и у хвоста есть «тело» полной ширины (усик в угол сужается сразу)
                        body = False: s = 0: hasFirst = False
                        For k = npth - 2 To 0 Step -1
                            s = s + Dist(GX(path(k)), GY(path(k)), GX(path(k + 1)), GY(path(k + 1)))
                            If GK(path(k)) = 0 And GR(path(k)) >= 0.9 * gW Then
                                If Not hasFirst Then
                                    first = s: hasFirst = True
                                ElseIf s - first >= 0.15 * gW Then
                                    body = True
                                    Exit For
                                End If
                            End If
                        Next k
                        keep = (coll And body)
                    End If
                    If Not keep Then
                        For k = 0 To npth - 2
                            KillNode path(k)
                        Next k
                        PruneLeaves = True
                    End If
                End If
            End If
        End If
    Next i
End Function

' стянуть путь (развилка ... развилка) в одну развилку посередине
Private Sub ContractPath(path() As Long, ByVal npth As Long)
    Dim ext() As Long, ne As Long, p As Long, k As Long, q As Long, z As Long, isP As Boolean, nid As Long
    Dim a As Long, b As Long
    a = path(0): b = path(npth - 1)
    nid = NewNode((GX(a) + GX(b)) / 2, (GY(a) + GY(b)) / 2, BIG, 1)
    ReDim ext(0 To npth * MAXD)
    For p = 0 To npth - 1
        For k = 0 To GDeg(path(p)) - 1
            q = GAdj(k, path(p))
            isP = False
            For z = 0 To npth - 1
                If path(z) = q Then
                    isP = True: Exit For
                End If
            Next z
            If Not isP Then
                ext(ne) = q: ne = ne + 1
            End If
        Next k
    Next p
    For p = 0 To npth - 1
        KillNode path(p)
    Next p
    For k = 0 To ne - 1
        LinkN nid, ext(k)
    Next k
End Sub

' дойти от развилки j через n0 по узлам степени 2 (не дальше maxd)
Private Sub WalkToJ(ByVal j As Long, ByVal n0 As Long, ByVal maxd As Double, path() As Long, npth As Long, _
                    prv As Long, cur As Long, L As Double)
    Dim nxt As Long
    npth = 0: path(npth) = j: npth = npth + 1
    prv = j: cur = n0: L = Dist(GX(j), GY(j), GX(n0), GY(n0))
    Do While GDeg(cur) = 2 And L <= maxd
        nxt = OtherNb(cur, prv)
        If nxt < 0 Then
            Exit Do
        End If
        path(npth) = cur: npth = npth + 1
        prv = cur: cur = nxt
        L = L + Dist(GX(prv), GY(prv), GX(cur), GY(cur))
    Loop
End Sub

Private Sub MergeJunctions(ByVal kmerge As Double)
    Dim changed As Boolean, j As Long, kk As Long, prv As Long, cur As Long
    Dim L As Double, path() As Long, npth As Long, maxd As Double, guard As Long
    Do
        changed = False
        ReDim path(0 To NG + 16)
        For j = 0 To NG - 1
            If GLive(j) And GDeg(j) >= 3 Then
                maxd = kmerge * JWidth(j)
                For kk = 0 To GDeg(j) - 1
                    WalkToJ j, GAdj(kk, j), maxd, path, npth, prv, cur, L
                    If cur <> j And GDeg(cur) >= 3 And L <= maxd Then
                        path(npth) = cur: npth = npth + 1
                        ContractPath path, npth
                        changed = True
                        Exit For
                    End If
                Next kk
                If changed Then
                    Exit For
                End If
            End If
        Next j
        guard = guard + 1
    Loop While changed And guard < 2000
End Sub

' ---------- пары сквозного прохода ----------
Private Sub ClearPairs()
    Dim i As Long
    ReDim PassA(0 To MAXD - 1, 0 To NG): ReDim PassB(0 To MAXD - 1, 0 To NG): ReDim PassN(0 To NG)
    ReDim GJW(0 To NG)
    For i = 0 To NG - 1
        If InG(i) And GDeg(i) >= 3 Then
            GJW(i) = JWidth(i)
        End If
    Next i
End Sub

Private Function PassOf(ByVal j As Long, ByVal a As Long) As Long
    Dim k As Long
    PassOf = -1
    If j > UBound(PassN) Then
        Exit Function
    End If
    For k = 0 To PassN(j) - 1
        If PassA(k, j) = a Then
            PassOf = PassB(k, j): Exit Function
        End If
    Next k
End Function

Private Sub AddPair(ByVal j As Long, ByVal a As Long, ByVal b As Long)
    If PassN(j) + 2 > MAXD Then
        Exit Sub
    End If
    PassA(PassN(j), j) = a: PassB(PassN(j), j) = b: PassN(j) = PassN(j) + 1
    PassA(PassN(j), j) = b: PassB(PassN(j), j) = a: PassN(j) = PassN(j) + 1
End Sub

' жадно: самые соосные пары веток в каждой развилке
Private Sub ComputePairs(ByVal mergeAng As Double)
    Dim i As Long, a As Long, b As Long, c As Long, d As Long, nb(0 To MAXD - 1) As Long
    Dim ddx(0 To MAXD - 1) As Double, ddy(0 To MAXD - 1) As Double
    Dim fxA(0 To MAXD - 1) As Double, fyA(0 To MAXD - 1) As Double, nxA(0 To MAXD - 1) As Double, nyA(0 To MAXD - 1) As Double
    Dim mxA(0 To MAXD - 1) As Double, myA(0 To MAXD - 1) As Double, rdA(0 To MAXD - 1) As Double
    Dim cd() As Double, ca() As Long, cb() As Long, nc As Long, used(0 To MAXD - 1) As Boolean
    Dim p As Long, q As Long, tmpD As Double, tmpL As Long, cosLim As Double, wi As Double
    Dim sgA() As Long, sgB() As Long, nsg As Long, okc As Boolean, npair As Long
    Dim xX As Double, yX As Double, dX As Double, xc As Double, yc As Double, ok As Boolean, k As Long
    Dim bestK As Long, bestD As Double, bestA As Long, bestB As Long, kk As Long
    ClearPairs
    cosLim = Cos(mergeAng * PI_ / 180)
    ReDim cd(0 To MAXD * MAXD): ReDim ca(0 To MAXD * MAXD): ReDim cb(0 To MAXD * MAXD)
    ReDim sgA(0 To MAXD): ReDim sgB(0 To MAXD)
    For i = 0 To NG - 1
        If InG(i) And GDeg(i) >= 3 Then
            d = GDeg(i): wi = GJW(i)
            For a = 0 To d - 1
                nb(a) = GAdj(a, i)
                BranchDir i, nb(a), wi, ddx(a), ddy(a), fxA(a), fyA(a), nxA(a), nyA(a), mxA(a), myA(a), rdA(a)
                used(a) = False
            Next a
            nc = 0
            For a = 0 To d - 2
                For b = a + 1 To d - 1
                    If Aimed(ddx(a), ddy(a), nxA(a), nyA(a), rdA(a), mxA(b), myA(b), rdA(b), wi) And _
                       Aimed(ddx(b), ddy(b), nxA(b), nyA(b), rdA(b), mxA(a), myA(a), rdA(a), wi) Then
                        cd(nc) = -(ddx(a) * ddx(b) + ddy(a) * ddy(b))
                        ca(nc) = a: cb(nc) = b: nc = nc + 1
                    End If
                Next b
            Next a
            For p = 1 To nc - 1
                For q = p To 1 Step -1
                    If cd(q) > cd(q - 1) Then
                        tmpD = cd(q): cd(q) = cd(q - 1): cd(q - 1) = tmpD
                        tmpL = ca(q): ca(q) = ca(q - 1): ca(q - 1) = tmpL
                        tmpL = cb(q): cb(q) = cb(q - 1): cb(q - 1) = tmpL
                    Else
                        Exit For
                    End If
                Next q
            Next p
            nsg = 0: npair = 0
            For p = 0 To nc - 1
                If Not used(ca(p)) And Not used(cb(p)) And cd(p) >= cosLim Then
                    okc = True
                    For k = 0 To nsg - 1
                        If Not SegsCross(fxA(ca(p)), fyA(ca(p)), fxA(cb(p)), fyA(cb(p)), _
                                         fxA(sgA(k)), fyA(sgA(k)), fxA(sgB(k)), fyA(sgB(k))) Then
                            okc = False
                            Exit For
                        End If
                    Next k
                    If okc Then
                        AddPair i, nb(ca(p)), nb(cb(p))
                        used(ca(p)) = True: used(cb(p)) = True
                        sgA(nsg) = ca(p): sgB(nsg) = cb(p): nsg = nsg + 1: npair = npair + 1
                    End If
                End If
            Next p
            If npair = 0 And d = 3 Then
                ' тройник без сквозной пары: угол внутри развилки (пятка Д) — сквозной считаем
                ' пару, оси которой пересекаются ближе всего к развилке (не дальше 1w), причём
                ' третья ветка должна упираться в одну из осей за углом X; третья — заглушка
                bestK = 99: bestD = BIG: bestA = -1
                For a = 0 To 1
                    For b = a + 1 To 2
                        If LineX(nxA(a), nyA(a), ddx(a), ddy(a), nxA(b), nyA(b), ddx(b), ddy(b), xX, yX) Then
                            dX = Dist(xX, yX, GX(i), GY(i))
                            If dX <= wi Then
                                c = 3 - a - b: ok = False
                                For kk = 0 To 1
                                    If kk = 0 Then
                                        k = a
                                    Else
                                        k = b
                                    End If
                                    If LineX(nxA(k), nyA(k), ddx(k), ddy(k), nxA(c), nyA(c), ddx(c), ddy(c), xc, yc) Then
                                        If Dist(xc, yc, xX, yX) <= 2 * wi And (xc - xX) * ddx(k) + (yc - yX) * ddy(k) > 0 Then
                                            ok = True
                                        End If
                                    End If
                                Next kk
                                If ok Then
                                    k = 0
                                Else
                                    k = 1
                                End If
                                If k < bestK Or (k = bestK And dX < bestD) Then
                                    bestK = k: bestD = dX: bestA = a: bestB = b
                                End If
                            End If
                        End If
                    Next b
                Next a
                If bestA >= 0 Then
                    AddPair i, nb(bestA), nb(bestB)
                End If
            End If
        End If
    Next i
End Sub

' все направления разбиваются на соосные пары?
Private Function AllPaired(ddx() As Double, ddy() As Double, fxA() As Double, fyA() As Double, ByVal n As Long, _
                           ByVal mergeAng As Double) As Boolean
    Dim a As Long, b As Long, used() As Boolean, cnt As Long, bestD As Double, ba As Long, bb As Long, d As Double
    Dim lim As Double, sgA() As Long, sgB() As Long, nsg As Long, k As Long
    If n Mod 2 = 1 Or n = 0 Then
        Exit Function
    End If
    lim = Cos(mergeAng * PI_ / 180)
    ReDim used(0 To n - 1): ReDim sgA(0 To n): ReDim sgB(0 To n)
    Do While cnt < n
        bestD = -2: ba = -1: bb = -1
        For a = 0 To n - 2
            If Not used(a) Then
                For b = a + 1 To n - 1
                    If Not used(b) Then
                        d = -(ddx(a) * ddx(b) + ddy(a) * ddy(b))
                        If d > bestD Then
                            bestD = d: ba = a: bb = b
                        End If
                    End If
                Next b
            End If
        Next a
        If ba < 0 Or bestD < lim Then
            Exit Function
        End If
        For k = 0 To nsg - 1
            If Not SegsCross(fxA(ba), fyA(ba), fxA(bb), fyA(bb), fxA(sgA(k)), fyA(sgA(k)), fxA(sgB(k)), fyA(sgB(k))) Then
                Exit Function
            End If
        Next k
        used(ba) = True: used(bb) = True: cnt = cnt + 2
        sgA(nsg) = ba: sgB(nsg) = bb: nsg = nsg + 1
    Loop
    AllPaired = True
End Function

' короткая перемычка между развилками, не идущая насквозь ни в одной из них, и
' после стягивания все ветки проходят крест насквозь (талия "8") -> одна развилка
Private Sub MergeBridges(ByVal mergeAng As Double)
    Dim changed As Boolean, j As Long, kk As Long, n0 As Long, prv As Long, cur As Long
    Dim L As Double, path() As Long, npth As Long, guard As Long, maxl As Double
    Dim ddx(0 To 2 * MAXD) As Double, ddy(0 To 2 * MAXD) As Double, fxA(0 To 2 * MAXD) As Double, fyA(0 To 2 * MAXD) As Double
    Dim nd As Long, k As Long, o As Long, t1 As Double, t2 As Double, t3 As Double, t4 As Double, t5 As Double, t6 As Double, t7 As Double
    Do
        changed = False
        ClearPairs
        ReDim path(0 To NG + 16)
        For j = 0 To NG - 1
            If InG(j) And GDeg(j) >= 3 Then
                maxl = 2.5 * GJW(j)
                For kk = 0 To GDeg(j) - 1
                    n0 = GAdj(kk, j)
                    WalkToJ j, n0, maxl, path, npth, prv, cur, L
                    If cur <> j And GDeg(cur) >= 3 And L <= maxl Then
                        nd = 0
                        For k = 0 To GDeg(j) - 1
                            o = GAdj(k, j)
                            If o <> n0 Then
                                BranchDir j, o, GJW(j), ddx(nd), ddy(nd), fxA(nd), fyA(nd), t1, t2, t3, t4, t5: nd = nd + 1
                            End If
                        Next k
                        For k = 0 To GDeg(cur) - 1
                            o = GAdj(k, cur)
                            If o <> prv Then
                                BranchDir cur, o, GJW(j), ddx(nd), ddy(nd), fxA(nd), fyA(nd), t1, t2, t3, t4, t5: nd = nd + 1
                            End If
                        Next k
                        If AllPaired(ddx, ddy, fxA, fyA, nd, mergeAng) Then
                            path(npth) = cur: npth = npth + 1
                            ContractPath path, npth
                            changed = True
                            Exit For
                        End If
                    End If
                Next kk
                If changed Then
                    Exit For
                End If
            End If
        Next j
        guard = guard + 1
    Loop While changed And guard < 2000
End Sub

Private Sub DropSmall()
    Dim seen() As Boolean, st() As Long, ns As Long, comp() As Long, nc As Long
    Dim s As Long, i As Long, k As Long, j As Long, L As Double, mc As Double
    ReDim seen(0 To NG): ReDim st(0 To NG): ReDim comp(0 To NG)
    For s = 0 To NG - 1
        If InG(s) And Not seen(s) Then
            nc = 0: ns = 0: L = 0: mc = 0
            st(ns) = s: ns = ns + 1: seen(s) = True
            Do While ns > 0
                ns = ns - 1: i = st(ns)
                comp(nc) = i: nc = nc + 1
                If GK(i) = 0 Then
                    If GR(i) > mc Then
                        mc = GR(i)
                    End If
                End If
                For k = 0 To GDeg(i) - 1
                    j = GAdj(k, i)
                    L = L + Dist(GX(i), GY(i), GX(j), GY(j))
                    If Not seen(j) Then
                        seen(j) = True: st(ns) = j: ns = ns + 1
                    End If
                Next k
            Loop
            If L / 2 < mc Then
                For k = 0 To nc - 1
                    KillNode comp(k)
                Next k
            End If
        End If
    Next s
End Sub

' ---------------------------------------------------------------------
'  Острые возвраты пера (V): развилка, где две ветки уходят в одну сторону,
'  а третья — короткий усик к острию. Делаем шпильку A -> острие -> B.
' ---------------------------------------------------------------------
Private Sub FindCusps()
    Dim j As Long, kk As Long, s As Long, prv As Long, cur As Long, nxt As Long, L As Double
    Dim path() As Long, npth As Long, dx As Double, dy As Double, D As Double, ok As Boolean
    Dim o(0 To 1) As Long, no As Long, ox(0 To 1) As Double, oy(0 To 1) As Double, k As Long
    Dim t As Double, found As Boolean, sx As Double, sy As Double, sl As Double, hx As Double, hy As Double, cv As Boolean
    ReDim CuspOn(0 To NG): ReDim CuspX(0 To NG): ReDim CuspY(0 To NG)
    ReDim path(0 To NG + 1)
    For j = 0 To NG - 1
        If InG(j) And GDeg(j) = 3 Then
            found = False
            For kk = 0 To 2
                If found Then
                    Exit For
                End If
                s = GAdj(kk, j)
                npth = 0: path(npth) = s: npth = npth + 1
                prv = j: cur = s: L = Dist(GX(j), GY(j), GX(s), GY(s))
                Do While GDeg(cur) = 2
                    nxt = OtherNb(cur, prv)
                    If nxt < 0 Then
                        Exit Do
                    End If
                    L = L + Dist(GX(cur), GY(cur), GX(nxt), GY(nxt))
                    prv = cur: cur = nxt
                    path(npth) = cur: npth = npth + 1
                Loop
                If GDeg(cur) = 1 Then
                    dx = GX(cur) - GX(j): dy = GY(cur) - GY(j)
                    If Norm(dx, dy) Then
                        D = RayUnion(GX(cur), GY(cur), dx, dy)
                        If D >= BIG Then
                            D = 0
                        End If
                        If L + D <= 2.5 * GJW(j) Then
                            no = 0: ok = True
                            For k = 0 To 2
                                If GAdj(k, j) <> s Then
                                    o(no) = GAdj(k, j)
                                    DirFrom j, o(no), 2.5 * GJW(j), ox(no), oy(no)
                                    If ox(no) * dx + oy(no) * dy > 0.2 Then
                                        ok = False
                                    End If
                                    no = no + 1
                                End If
                            Next k
                            If ok Then
                                sx = ox(0) + ox(1): sy = oy(0) + oy(1)
                                sl = Sqr(sx * sx + sy * sy)
                                If sl < 0.6 Then
                                    ok = False
                                ElseIf (sx * dx + sy * dy) / sl > -0.7 Then
                                    ok = False
                                End If
                            End If
                            If ok Then
                                found = True
                                D = RayUnion(GX(j), GY(j), dx, dy)
                                If D >= BIG Then
                                    D = 0
                                End If
                                ' остриё должно упираться в выпуклый острый угол контура
                                hx = GX(j) + dx * D: hy = GY(j) + dy * D
                                cv = False
                                For k = 0 To NBC - 1
                                    If BCC(k) Then
                                        If Dist(BCX(k), BCY(k), hx, hy) < 0.6 * GJW(j) Then
                                            cv = True
                                            Exit For
                                        End If
                                    End If
                                Next k
                                If cv Then
                                    t = D - gTool / 2
                                    If t < 0 Then
                                        t = 0
                                    End If
                                    CuspOn(j) = True
                                    CuspX(j) = GX(j) + dx * t: CuspY(j) = GY(j) + dy * t
                                    For k = 0 To npth - 1
                                        KillNode path(k)
                                    Next k
                                    PassN(j) = 0
                                    AddPair j, o(0), o(1)
                                End If
                            End If
                        End If
                    End If
                End If
            Next kk
        End If
    Next j
End Sub

' ---------------------------------------------------------------------
'  Цепочки
' ---------------------------------------------------------------------
Private Function IsUsed(ByVal i As Long, ByVal j As Long) As Boolean
    Dim k As Long
    For k = 0 To GDeg(i) - 1
        If GAdj(k, i) = j Then
            IsUsed = GUsed(k, i): Exit Function
        End If
    Next k
End Function

Private Sub SetUsed(ByVal i As Long, ByVal j As Long)
    Dim k As Long
    For k = 0 To GDeg(i) - 1
        If GAdj(k, i) = j Then
            GUsed(k, i) = True
        End If
    Next k
    For k = 0 To GDeg(j) - 1
        If GAdj(k, j) = i Then
            GUsed(k, j) = True
        End If
    Next k
End Sub

Private Sub BuildChains(ByVal mergeAng As Double)
    Dim i As Long, k As Long, j As Long, b As Long, p As Long, m As Long
    Dim ch() As Long, chn As Long, cl As Boolean, bk() As Long, nbk As Long, cl2 As Boolean
    Dim outIds() As Long

    ComputePairs mergeAng
    FindCusps
    For i = 0 To NG - 1
        For k = 0 To MAXD - 1: GUsed(k, i) = False: Next k
    Next i
    NCh = 0
    ReDim ChI(0 To 63): ReDim ChC(0 To 63)

    ' от листьев и развилок
    For i = 0 To NG - 1
        If InG(i) And GDeg(i) <> 2 Then
            For k = 0 To GDeg(i) - 1
                j = GAdj(k, i)
                If Not IsUsed(i, j) Then
                    If Not (GDeg(i) >= 3 And PassOf(i, j) >= 0) Then
                        WalkChain i, j, ch, chn, cl
                        AddChain ch, chn, cl
                    End If
                End If
            Next k
        End If
    Next i
    ' оставшиеся: сквозные с развилками на обоих концах, циклы
    For i = 0 To NG - 1
        If InG(i) Then
            For k = 0 To GDeg(i) - 1
                j = GAdj(k, i)
                If Not IsUsed(i, j) Then
                    WalkChain i, j, ch, chn, cl
                    If Not cl And GDeg(i) >= 3 And PassOf(i, j) >= 0 Then
                        b = PassOf(i, j)
                        If Not IsUsed(i, b) Then
                            WalkChain i, b, bk, nbk, cl2
                        Else
                            ReDim bk(0 To 0): bk(0) = i: nbk = 1
                        End If
                        ReDim outIds(0 To nbk + chn)
                        m = 0
                        For p = nbk - 1 To 1 Step -1
                            outIds(m) = bk(p): m = m + 1
                        Next p
                        For p = 0 To chn - 1
                            outIds(m) = ch(p): m = m + 1
                        Next p
                        AddChain outIds, m, False
                    Else
                        AddChain ch, chn, cl
                    End If
                End If
            Next k
        End If
    Next i
End Sub

Private Sub AddChain(ids() As Long, ByVal n As Long, ByVal closed As Boolean)
    Dim a() As Long, i As Long
    If n < 2 Then
        Exit Sub
    End If
    ReDim a(0 To n - 1)
    For i = 0 To n - 1: a(i) = ids(i): Next i
    If NCh > UBound(ChI) Then
        ReDim Preserve ChI(0 To 2 * NCh): ReDim Preserve ChC(0 To 2 * NCh)
    End If
    ChI(NCh) = a: ChC(NCh) = closed
    NCh = NCh + 1
End Sub

Private Sub WalkChain(ByVal st As Long, ByVal nx As Long, ch() As Long, chn As Long, closed As Boolean)
    Dim prv As Long, cur As Long, n2 As Long
    ReDim ch(0 To 63)
    chn = 0: closed = False
    ch(chn) = st: chn = chn + 1
    prv = st: cur = nx
    SetUsed prv, cur
    Do
        If chn > UBound(ch) Then
            ReDim Preserve ch(0 To 2 * chn)
        End If
        ch(chn) = cur: chn = chn + 1
        If cur = st And GDeg(cur) = 2 Then
            closed = True
            Exit Sub
        End If
        If cur = st And PassOf(cur, prv) < 0 Then
            Exit Sub
        End If
        If GDeg(cur) = 2 Then
            n2 = OtherNb(cur, prv)
        ElseIf GDeg(cur) >= 3 And PassOf(cur, prv) >= 0 Then
            n2 = PassOf(cur, prv)
        Else
            Exit Sub
        End If
        If n2 < 0 Then
            Exit Sub
        End If
        If IsUsed(cur, n2) Then
            Exit Sub
        End If
        SetUsed cur, n2
        prv = cur: cur = n2
    Loop
End Sub

' "особая" точка цепочки: развилка (степень >= 3) или острие
Private Function IsJ(ByVal i As Long) As Boolean
    If GDeg(i) >= 3 Then
        IsJ = True
    ElseIf CuspOn(i) Then
        IsJ = True
    End If
End Function

' срезать хвост у свободного конца ids(n-1): зона скругления/торца; вернуть новую длину
Private Function EndTrim(ids() As Long, ByVal n As Long) As Long
    Dim k As Long, m As Long, L As Double, c As Double, cutL As Double, cmax As Double, stp As Double
    k = n - 1
    Do While k > 1
        If GK(ids(k)) <> 0 Then
            Exit Do
        End If
        c = GR(ids(k))
        m = k: L = 0
        Do While m > 0 And L < 0.5 * c
            L = L + Dist(GX(ids(m)), GY(ids(m)), GX(ids(m - 1)), GY(ids(m - 1))): m = m - 1
        Loop
        If GK(ids(m)) <> 0 Then
            Exit Do
        End If
        If GR(ids(m)) * P_ENDTRIM <= c Then
            Exit Do
        End If
        If GR(ids(m)) > cmax Then
            cmax = GR(ids(m))
        End If
        stp = Dist(GX(ids(k)), GY(ids(k)), GX(ids(k - 1)), GY(ids(k - 1)))
        If cutL + stp > 0.6 * cmax Then
            Exit Do
        End If
        cutL = cutL + stp
        k = k - 1
    Loop
    EndTrim = k + 1
End Function

Private Sub ReverseIds(ids() As Long, ByVal n As Long)
    Dim i As Long, j As Long, t As Long
    i = 0: j = n - 1
    Do While i < j
        t = ids(i): ids(i) = ids(j): ids(j) = t
        i = i + 1: j = j - 1
    Loop
End Sub

Private Function MedianChord(ids() As Long, ByVal n As Long) As Double
    Dim a() As Double, m As Long, i As Long, j As Long, t As Double, gap As Long
    ReDim a(0 To n)
    For i = 0 To n - 1
        If GK(ids(i)) = 0 Then
            a(m) = GR(ids(i)): m = m + 1
        End If
    Next i
    If m = 0 Then
        MedianChord = gW
        Exit Function
    End If
    gap = m \ 2
    Do While gap > 0
        For i = gap To m - 1
            t = a(i): j = i
            Do While j >= gap
                If a(j - gap) <= t Then
                    Exit Do
                End If
                a(j) = a(j - gap): j = j - gap
            Loop
            a(j) = t
        Next i
        gap = gap \ 2
    Loop
    MedianChord = a(m \ 2)
End Function

Private Sub ProcessChains()
    Dim ci As Long, ids() As Long, n As Long, closed As Boolean, i As Long, k As Long, k2 As Long, e As Long, j As Long
    Dim thru() As Long, stub0() As Long, stub1() As Long, free0 As Boolean, free1 As Boolean
    Dim cX() As Variant, cY() As Variant, dN() As Long, dCl() As Boolean, dOk() As Boolean
    Dim dRef() As Double, dNm0() As Long, dNm1() As Long
    Dim flags() As Boolean, hard() As Boolean, cref As Double, r0 As Long
    Dim xs() As Double, ys() As Double, np_ As Long, fixedF() As Boolean
    Dim fx() As Double, fy() As Double, nf As Long, jc As Long, cp As Long, hasFill As Boolean
    Dim ax_() As Double, ay_() As Double, na As Long, nxx As Double, nxy As Double, hasNxt As Boolean
    Dim js As Long, lf As Long, q As Long, q2 As Long, allBad As Boolean, isCuspFill As Boolean
    Dim okArc As Boolean, pf(0 To 5) As Double, pT As Long, pA As Double, pB As Double, pC As Double, pD As Double
    Dim pendX As Double, pendY As Double, pendR As Double, qx As Double, qy As Double
    Dim total As Double, dirn As Long, zj As Double, sz As Double, lr As Double, cap As Double, zh As Double, zs As Double, L As Double
    Dim td0 As Boolean, td1 As Boolean, side As Long, run() As Long, nrun As Long, prevQ As Long, hasHard As Boolean, jq As Long
    Dim rx() As Double, ry() As Double, nr_ As Long, xi As Long, fixI As Long, minF As Long, maxF As Long, hasF As Boolean
    Dim nm0 As Long, nm1 As Long, okDrop As Boolean, o3 As Long, sdx As Double, sdy As Double, edx As Double, edy As Double
    Dim t1 As Double, t2 As Double, t3 As Double, t4 As Double, t5 As Double, t6 As Double, t7 As Double

    If NCh = 0 Then
        Exit Sub
    End If
    ReDim thru(0 To NG)
    For i = 0 To NG: thru(i) = -1: Next i
    ReDim stub0(0 To NCh - 1): ReDim stub1(0 To NCh - 1)
    ReDim cX(0 To NCh - 1): ReDim cY(0 To NCh - 1): ReDim dN(0 To NCh - 1)
    ReDim dCl(0 To NCh - 1): ReDim dOk(0 To NCh - 1): ReDim dRef(0 To NCh - 1)
    ReDim dNm0(0 To NCh - 1): ReDim dNm1(0 To NCh - 1)

    For ci = 0 To NCh - 1
        ids = ChI(ci): closed = ChC(ci): n = UBound(ids) + 1
        For k = 0 To n - 1
            If closed Or (k > 0 And k < n - 1) Then
                If GDeg(ids(k)) >= 3 And PassN(ids(k)) > 0 Then
                    thru(ids(k)) = ci
                End If
            End If
        Next k
    Next ci

    For ci = 0 To NCh - 1
        ids = ChI(ci): closed = ChC(ci): n = UBound(ids) + 1
        stub0(ci) = -1: stub1(ci) = -1: free0 = False: free1 = False: td0 = False: td1 = False
        If Not closed Then
            If GDeg(ids(0)) = 1 Then
                free0 = True
            ElseIf GDeg(ids(0)) >= 3 Then
                stub0(ci) = ids(0)
            End If
            If GDeg(ids(n - 1)) = 1 Then
                free1 = True
            ElseIf GDeg(ids(n - 1)) >= 3 Then
                stub1(ci) = ids(n - 1)
            End If
            If free1 Then
                n = EndTrim(ids, n)
            End If
            If free0 Then
                ReverseIds ids, n
                n = EndTrim(ids, n)
                ReverseIds ids, n
            End If
        End If
        cref = MedianChord(ids, n)
        total = 0
        For k = 0 To n - 2
            total = total + Dist(GX(ids(k)), GY(ids(k)), GX(ids(k + 1)), GY(ids(k + 1)))
        Next k
        If closed Then
            total = total + Dist(GX(ids(n - 1)), GY(ids(n - 1)), GX(ids(0)), GY(ids(0)))
        End If

        ' особые узлы (развилки, острия) и зоны вокруг них: радиус = min(JZONE*jw, треть
        ' длины участка до соседнего особого узла / конца цепочки) — короткие перемычки
        ' не должны исчезать целиком; мягкая зона 1.8jw — только узлы со штрихом шире обычного
        ReDim flags(0 To n - 1): ReDim hard(0 To n - 1)
        For k = 0 To n - 1
            hard(k) = IsJ(ids(k)): flags(k) = hard(k)
        Next k
        For k = 0 To n - 1
            If IsJ(ids(k)) Then
                j = ids(k)
                zj = P_JZONE * JWOf(j, cref)
                If CuspOn(j) Then
                    zj = zj * 0.5
                End If
                sz = 1.8 * JWOf(j, cref)
                If (Not closed) And (k = 0 Or k = n - 1) And (Not CuspOn(j)) Then
                    ' конец цепочки — заглушка: её ось у развилки искажена меньше, чем сквозная
                    ' (нет «проседания»), а короткой перекладине (низ чаши P) нужны свои точки
                    zj = P_STUBZ * JWOf(j, cref)
                    If sz > JWOf(j, cref) Then
                        sz = JWOf(j, cref)
                    End If
                End If
                For dirn = -1 To 1 Step 2
                    lr = RunLen(ids, hard, n, closed, k, dirn)
                    cap = lr / 3
                    zh = zj: zs = sz
                    If cap < zh Then
                        zh = cap
                    End If
                    If cap < zs Then
                        zs = cap
                    End If
                    L = 0: q = k
                    Do
                        q2 = q + dirn
                        If closed Then
                            q2 = (q2 + n) Mod n
                        ElseIf q2 < 0 Or q2 >= n Then
                            Exit Do
                        End If
                        L = L + Dist(GX(ids(q)), GY(ids(q)), GX(ids(q2)), GY(ids(q2))): q = q2
                        If hard(q) And IsJ(ids(q)) Then
                            Exit Do
                        End If
                        If L >= zs Then
                            Exit Do
                        End If
                        If L < zh Then
                            hard(q) = True: flags(q) = True
                        ElseIf Not CuspOn(j) Then
                            If GR(ids(q)) > P_WIDEN * cref Then
                                flags(q) = True
                            End If
                        End If
                    Loop
                Next dirn
            End If
        Next k
        ' короткий остаток (< 1 jw) между зоной развилки и свободным концом — выбрасываем:
        ' конец достроится продолжением оси сквозной линии (торец ножки у N, 4, k).
        ' Но не лист-заглушку целиком (хвост Q) и только если за развилкой цепочка далеко продолжается
        If Not closed Then
            ReDim run(0 To n)
            For side = 0 To 1
                If (side = 0 And free0) Or (side = 1 And free1) Then
                    nrun = 0: L = 0: prevQ = -1: hasHard = False
                    If side = 1 Then
                        q = n - 1
                    Else
                        q = 0
                    End If
                    Do While q >= 0 And q < n
                        If hard(q) Then
                            hasHard = True
                            Exit Do
                        End If
                        If prevQ >= 0 Then
                            L = L + Dist(GX(ids(q)), GY(ids(q)), GX(ids(prevQ)), GY(ids(prevQ)))
                        End If
                        run(nrun) = q: nrun = nrun + 1: prevQ = q
                        If side = 1 Then
                            q = q - 1
                        Else
                            q = q + 1
                        End If
                    Loop
                    If hasHard Then
                        jq = ids(q)
                        okDrop = False
                        If Not ((q = 0 Or q = n - 1) And GDeg(jq) >= 3) Then
                            If total - L >= 3 * JWOf(jq, cref) Then
                                ' «колпачок» — это торец штриха-заглушки (ножка N торчит над
                                ' диагональю): заглушка должна смотреть ровно против остатка.
                                ' Иначе (рукописная n: штрих сам идёт дальше за развилку и слегка
                                ' гнётся) остаток — настоящая ось, его нельзя выбрасывать
                                js = -1: q2 = q
                                Do While q2 >= 0 And q2 < n
                                    If GDeg(ids(q2)) = 3 And PassN(ids(q2)) > 0 Then
                                        js = ids(q2)
                                        Exit Do
                                    End If
                                    If Not hard(q2) Then
                                        Exit Do
                                    End If
                                    If side = 1 Then
                                        q2 = q2 - 1
                                    Else
                                        q2 = q2 + 1
                                    End If
                                Loop
                                If js >= 0 Then
                                    o3 = -1
                                    For e = 0 To GDeg(js) - 1
                                        If PassOf(js, GAdj(e, js)) < 0 Then
                                            o3 = GAdj(e, js)
                                        End If
                                    Next e
                                    If o3 >= 0 Then
                                        BranchDir js, o3, JWOf(js, cref), sdx, sdy, t1, t2, t3, t4, t5, t6, t7
                                        If side = 1 Then
                                            e = ids(n - 1)
                                        Else
                                            e = ids(0)
                                        End If
                                        edx = GX(e) - GX(js): edy = GY(e) - GY(js): Norm edx, edy
                                        If sdx * edx + sdy * edy <= -0.85 Then
                                            okDrop = True
                                        End If
                                    End If
                                End If
                            End If
                        End If
                        If okDrop Then
                            If True Then
                                If nrun > 0 And L < JWOf(jq, cref) Then
                                    For e = 0 To nrun - 1
                                        hard(run(e)) = True: flags(run(e)) = True
                                    Next e
                                    If side = 1 Then
                                        td1 = True
                                    Else
                                        td0 = True
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            Next side
        End If
        allBad = True
        For k = 0 To n - 1
            If Not flags(k) Then
                allBad = False
                Exit For
            End If
        Next k
        If allBad Then
            GoTo NextChain
        End If
        If closed And flags(0) Then
            r0 = 0
            Do While flags(r0)
                r0 = r0 + 1
            Loop
            RotateIds ids, flags, n, r0
        End If

        ' сборка точек; дыры в зонах развилок — дуга/угол/Эрмит/в острие
        np_ = 0: ReDim xs(0 To n + 16): ReDim ys(0 To n + 16): ReDim fixedF(0 To n + 16)
        pT = 0
        k = 0
        Do While k < n
            If Not flags(k) Then
                qx = GX(ids(k)): qy = GY(ids(k))
                If pT <> 0 Then
                    If Dist(qx, qy, pendX, pendY) <= pendR Then
                        ProjPt pT, pA, pB, pC, pD, qx, qy
                    Else
                        pT = 0
                    End If
                End If
                PushPtF xs, ys, fixedF, np_, qx, qy, False
                k = k + 1
            Else
                k2 = k
                Do While k2 < n
                    If Not flags(k2) Then
                        Exit Do
                    End If
                    k2 = k2 + 1
                Loop
                jc = -1: cp = -1
                For e = k To k2 - 1
                    If PassN(ids(e)) >= 4 And jc < 0 Then
                        jc = ids(e)
                    End If
                    If CuspOn(ids(e)) And cp < 0 Then
                        cp = ids(e)
                    End If
                Next e
                hasFill = False: isCuspFill = False: nf = 0: fixI = -1
                ReDim fx(0 To 7): ReDim fy(0 To 7)
                If np_ > 0 Then
                    If cp >= 0 And (k2 < n Or closed) Then
                        PushPt fx, fy, nf, CuspX(cp), CuspY(cp)
                        hasFill = True: isCuspFill = True
                    ElseIf k2 < n Then
                        na = 0: ReDim ax_(0 To n - k2): ReDim ay_(0 To n - k2)
                        For e = k2 To n - 1
                            ax_(na) = GX(ids(e)): ay_(na) = GY(ids(e)): na = na + 1
                        Next e
                        nr_ = 0: ReDim rx(0 To k2 - k): ReDim ry(0 To k2 - k)
                        For e = k To k2 - 1
                            rx(nr_) = GX(ids(e)): ry(nr_) = GY(ids(e)): nr_ = nr_ + 1
                        Next e
                        ' сначала фит одной кривой/прямой через обе стороны (и для перекрестий:
                        ' прямая диагональ X должна остаться прямой), либо угол, затем Эрмит
                        okArc = ArcFill(xs, ys, np_, ax_, ay_, na, gH * 1.4, cref, rx, ry, nr_, fx, fy, nf, pf, xi)
                        If okArc Then
                            pT = pf(0): pA = pf(1): pB = pf(2): pC = pf(3): pD = pf(4)
                            pendX = GX(ids(k2)): pendY = GY(ids(k2)): pendR = 0.8 * cref
                            If xi >= 0 Then
                                fixI = xi
                            End If
                        Else
                            nf = 0
                            HermiteFill xs, ys, np_, ax_, ay_, na, gH * 1.4, jc, fx, fy, nf
                        End If
                        hasFill = True
                    ElseIf closed Then
                        na = np_: ReDim ax_(0 To na - 1): ReDim ay_(0 To na - 1)
                        For e = 0 To na - 1: ax_(e) = xs(e): ay_(e) = ys(e): Next e
                        HermiteFill xs, ys, np_, ax_, ay_, na, gH * 1.4, jc, fx, fy, nf
                        hasFill = True
                    ElseIf free1 And Not td1 Then
                        ' зона развилки дошла до свободного конца — тянемся к самому концу
                        PushPt fx, fy, nf, GX(ids(n - 1)), GY(ids(n - 1))
                        hasFill = True
                    End If
                End If
                If hasFill Then
                    hasNxt = False
                    If k2 < n Then
                        nxx = GX(ids(k2)): nxy = GY(ids(k2)): hasNxt = True
                    ElseIf closed Then
                        nxx = xs(0): nxy = ys(0): hasNxt = True
                    End If
                    If Not PathInside(xs(np_ - 1), ys(np_ - 1), fx, fy, nf, nxx, nxy, hasNxt) Then
                        ' обход вылез за контур — берём сырые узлы
                        For e = k To k2 - 1
                            PushPtF xs, ys, fixedF, np_, GX(ids(e)), GY(ids(e)), False
                        Next e
                    Else
                        For e = 0 To nf - 1
                            PushPtF xs, ys, fixedF, np_, fx(e), fy(e), isCuspFill Or (e = fixI)
                            If e = fixI Then
                                If NSh > UBound(ShX) Then
                                    ReDim Preserve ShX(0 To 2 * NSh): ReDim Preserve ShY(0 To 2 * NSh)
                                End If
                                ShX(NSh) = fx(e): ShY(NSh) = fy(e): NSh = NSh + 1
                            End If
                        Next e
                    End If
                End If
                k = k2
            End If
        Loop
        ' зона развилки в самом начале у свободного конца
        If Not closed And free0 And np_ > 0 And Not td0 Then
            If flags(0) Then
                If SegInside(GX(ids(0)), GY(ids(0)), xs(0), ys(0)) Then
                    InsertFront xs, ys, fixedF, np_, GX(ids(0)), GY(ids(0))
                Else
                    q = 0
                    Do While q < n
                        If Not flags(q) Then
                            Exit Do
                        End If
                        q = q + 1
                    Loop
                    For e = q - 1 To 0 Step -1
                        InsertFront xs, ys, fixedF, np_, GX(ids(e)), GY(ids(e))
                    Next e
                End If
            End If
        End If
        If np_ < 2 Then
            If Not closed And (free0 Or free1) And (stub0(ci) >= 0 Or stub1(ci) >= 0) Then
                ' короткий "язычок" целиком в зоне развилки
                If stub0(ci) >= 0 Then
                    js = stub0(ci)
                Else
                    js = stub1(ci)
                End If
                np_ = 0
                If free1 Then
                    lf = ids(n - 1)
                    PushPtF xs, ys, fixedF, np_, GX(js), GY(js), False
                    PushPtF xs, ys, fixedF, np_, GX(lf), GY(lf), False
                Else
                    lf = ids(0)
                    PushPtF xs, ys, fixedF, np_, GX(lf), GY(lf), False
                    PushPtF xs, ys, fixedF, np_, GX(js), GY(js), False
                End If
                stub0(ci) = -1: stub1(ci) = -1
            ElseIf Not closed And stub0(ci) >= 0 And stub1(ci) >= 0 And stub0(ci) <> stub1(ci) Then
                ' короткая перемычка между двумя развилками
                np_ = 0
                PushPtF xs, ys, fixedF, np_, GX(stub0(ci)), GY(stub0(ci)), False
                PushPtF xs, ys, fixedF, np_, GX(stub1(ci)), GY(stub1(ci)), False
                stub0(ci) = -1: stub1(ci) = -1
            Else
                GoTo NextChain
            End If
        End If

        Smooth xs, ys, fixedF, np_, closed, P_SMOOTH
        Sharpen xs, ys, fixedF, np_, closed, cref
        MergeCollinear xs, ys, fixedF, np_, closed, cref
        ' точки от конца до ближайшего острого угла (включительно) — дальше концы не смотрят
        hasF = False: minF = -1: maxF = -1
        For q = 0 To np_ - 1
            If fixedF(q) Then
                If Not hasF Then
                    minF = q: hasF = True
                End If
                maxF = q
            End If
        Next q
        If hasF Then
            nm0 = minF + 1: nm1 = np_ - maxF
        Else
            nm0 = -1: nm1 = -1
        End If
        If Not closed Then
            If free1 Then
                FreeEnd xs, ys, np_, cref, nm1
            End If
            If free0 Then
                ReversePts xs, ys, np_
                FreeEnd xs, ys, np_, cref, nm0
                ReversePts xs, ys, np_
            End If
        End If
        cX(ci) = CopyD(xs, np_): cY(ci) = CopyD(ys, np_): dN(ci) = np_: dCl(ci) = closed: dOk(ci) = True: dRef(ci) = cref
        dNm0(ci) = nm0: dNm1(ci) = nm1
NextChain:
    Next ci

    ' заглушки: продление до сквозной линии развилки (кусок обработанной ломаной сквозной
    ' цепочки между её узлами-соседями развилки, +2jw в обе стороны), иначе по касательной
    ' или в центр развилки
    Dim tgtX As Double, tgtY As Double, hit As Double, poly As Long, m As Long, nmx As Long
    Dim hasT As Boolean, dx As Double, dy As Double, bx() As Double, by() As Double, nb As Long
    Dim plx() As Double, ply() As Double, npl As Long, a As Long, b As Long, ia As Long, ib As Long, lo As Long, hi As Long
    Dim bestD As Double, d As Double, dvx As Double, dvy As Double, v0 As Long, v1 As Long
    For ci = 0 To NCh - 1
        If dOk(ci) Then
            For side = 0 To 1
                If side = 0 Then
                    j = stub0(ci): nmx = dNm0(ci)
                Else
                    j = stub1(ci): nmx = dNm1(ci)
                End If
                If j >= 0 Then
                    xs = cX(ci): ys = cY(ci): np_ = dN(ci)
                    If side = 0 Then
                        ReversePts xs, ys, np_
                    End If
                    hasT = False: npl = 0: v0 = 0: v1 = 0
                    ReDim plx(0 To 15): ReDim ply(0 To 15)
                    poly = thru(j)
                    If poly >= 0 And PassN(j) > 0 Then
                        If dOk(poly) Then
                            bx = cX(poly): by = cY(poly): nb = dN(poly)
                            a = PassA(0, j): b = PassB(0, j)
                            bestD = BIG: ia = 0
                            For m = 0 To nb - 1
                                d = Dist(bx(m), by(m), GX(a), GY(a))
                                If d < bestD Then
                                    bestD = d: ia = m
                                End If
                            Next m
                            bestD = BIG: ib = 0
                            For m = 0 To nb - 1
                                d = Dist(bx(m), by(m), GX(b), GY(b))
                                If d < bestD Then
                                    bestD = d: ib = m
                                End If
                            Next m
                            If ia < ib Then
                                lo = ia: hi = ib
                            Else
                                lo = ib: hi = ia
                            End If
                            L = 0
                            Do While lo > 0 And L < 2 * GJW(j)
                                L = L + Dist(bx(lo), by(lo), bx(lo - 1), by(lo - 1)): lo = lo - 1
                            Loop
                            L = 0
                            Do While hi < nb - 1 And L < 2 * GJW(j)
                                L = L + Dist(bx(hi), by(hi), bx(hi + 1), by(hi + 1)): hi = hi + 1
                            Loop
                            ' у свободных концов сквозной линии — виртуальное продолжение на 1jw:
                            ' заглушка, приходящая ровно в торец, не должна промахнуться из-за округления
                            v0 = 0: v1 = 0
                            If lo = 0 And hi > lo Then
                                dvx = bx(0) - bx(1): dvy = by(0) - by(1): Norm dvx, dvy
                                PushPt plx, ply, npl, bx(0) + dvx * 0.5 * GJW(j), by(0) + dvy * 0.5 * GJW(j): v0 = 1
                            End If
                            For m = lo To hi
                                PushPt plx, ply, npl, bx(m), by(m)
                            Next m
                            If hi = nb - 1 And hi > lo Then
                                dvx = bx(nb - 1) - bx(nb - 2): dvy = by(nb - 1) - by(nb - 2): Norm dvx, dvy
                                PushPt plx, ply, npl, bx(nb - 1) + dvx * 0.5 * GJW(j), by(nb - 1) + dvy * 0.5 * GJW(j): v1 = 1
                            End If
                        End If
                    End If
                    If Not StubExtend(xs, ys, np_, dRef(ci), plx, ply, npl, v0, v1, GJW(j), GX(j), GY(j), nmx) Then
                        ' запасной вариант: плавно к точке сквозной линии у центра развилки
                        StubToJunction xs, ys, np_, dRef(ci), plx, ply, npl, v0, v1, GX(j), GY(j), GJW(j)
                    End If
                    If side = 0 Then
                        ReversePts xs, ys, np_
                    End If
                    cX(ci) = CopyD(xs, np_): cY(ci) = CopyD(ys, np_): dN(ci) = np_
                End If
            Next side
        End If
    Next ci

    ' аппроксимация
    For ci = 0 To NCh - 1
        If dOk(ci) Then
            xs = cX(ci): ys = cY(ci)
            Deglitch xs, ys, dN(ci), dCl(ci), gW
            FitPath xs, ys, dN(ci), dCl(ci), dRef(ci)
        End If
    Next ci
End Sub

' ---------------------------------------------------------------------
'  Острые углы контура: поворот > 40 град на участке +-0.15W
' ---------------------------------------------------------------------
Private Sub FindBCorners()
    Dim r As Long, i As Long, b As Long, n As Long, a As Long, c As Long, j As Long, k As Long, s As Double
    Dim d1x As Double, d1y As Double, d2x As Double, d2y As Double, lim As Double, ccw As Boolean, hole As Boolean
    Dim lft As Boolean, ex As Double, ey As Double, side As Long, stp As Long, dev As Double, mx As Double, q As Long
    NBC = 0
    ReDim BCX(0 To 63): ReDim BCY(0 To 63): ReDim BCC(0 To 63)
    ReDim BE1X(0 To 63): ReDim BE1Y(0 To 63): ReDim BE2X(0 To 63): ReDim BE2Y(0 To 63)
    ReDim BS1(0 To 63): ReDim BS2(0 To 63)
    lim = Cos(40 * PI_ / 180)
    For r = 0 To NR - 1
        b = RS(r): n = RN(r)
        ccw = (RingArea(r) > 0)
        hole = (r <> PolyR0(RPoly(r)))
        For i = 0 To n - 1
            s = 0: j = i
            For k = 1 To n
                a = (j - 1 + n) Mod n
                s = s + Dist(RX(b + j), RY(b + j), RX(b + a), RY(b + a)): j = a
                If s >= 0.15 * gW Then
                    Exit For
                End If
            Next k
            a = j
            s = 0: j = i
            For k = 1 To n
                c = (j + 1) Mod n
                s = s + Dist(RX(b + j), RY(b + j), RX(b + c), RY(b + c)): j = c
                If s >= 0.15 * gW Then
                    Exit For
                End If
            Next k
            c = j
            d1x = RX(b + i) - RX(b + a): d1y = RY(b + i) - RY(b + a): Norm d1x, d1y
            d2x = RX(b + c) - RX(b + i): d2y = RY(b + c) - RY(b + i): Norm d2x, d2y
            If d1x * d2x + d1y * d2y < lim Then
                If NBC > UBound(BCX) Then
                    ReDim Preserve BCX(0 To 2 * NBC): ReDim Preserve BCY(0 To 2 * NBC): ReDim Preserve BCC(0 To 2 * NBC)
                    ReDim Preserve BE1X(0 To 2 * NBC): ReDim Preserve BE1Y(0 To 2 * NBC)
                    ReDim Preserve BE2X(0 To 2 * NBC): ReDim Preserve BE2Y(0 To 2 * NBC)
                    ReDim Preserve BS1(0 To 2 * NBC): ReDim Preserve BS2(0 To 2 * NBC)
                End If
                lft = ((d1x * d2y - d1y * d2x) > 0)
                BCX(NBC) = RX(b + i): BCY(NBC) = RY(b + i)
                BCC(NBC) = ((lft = ccw) <> hole)
                ' кромки: хорды на 0.5W в обе стороны; прямизна кромки — макс. отклонение
                ' контура от хорды на 0.8W от угла (в долях W)
                For side = 0 To 1
                    If side = 0 Then
                        stp = -1
                    Else
                        stp = 1
                    End If
                    s = 0: j = i
                    For k = 1 To n
                        q = (j + stp + n) Mod n
                        s = s + Dist(RX(b + j), RY(b + j), RX(b + q), RY(b + q)): j = q
                        If s >= 0.5 * gW Then
                            Exit For
                        End If
                    Next k
                    ex = RX(b + j) - RX(b + i): ey = RY(b + j) - RY(b + i): Norm ex, ey
                    If side = 0 Then
                        BE1X(NBC) = ex: BE1Y(NBC) = ey
                    Else
                        BE2X(NBC) = ex: BE2Y(NBC) = ey
                    End If
                    s = 0: j = i
                    For k = 1 To n
                        q = (j + stp + n) Mod n
                        s = s + Dist(RX(b + j), RY(b + j), RX(b + q), RY(b + q)): j = q
                        If s >= 0.8 * gW Then
                            Exit For
                        End If
                    Next k
                    ex = RX(b + j) - RX(b + i): ey = RY(b + j) - RY(b + i): Norm ex, ey
                    mx = 0: q = i
                    For k = 1 To n
                        q = (q + stp + n) Mod n
                        dev = Abs((RX(b + q) - RX(b + i)) * ey - (RY(b + q) - RY(b + i)) * ex)
                        If dev > mx Then
                            mx = dev
                        End If
                        If q = j Then
                            Exit For
                        End If
                    Next k
                    If side = 0 Then
                        BS1(NBC) = mx / gW
                    Else
                        BS2(NBC) = mx / gW
                    End If
                Next side
                NBC = NBC + 1
            End If
        Next i
    Next r
End Sub

Private Function NearBCorner(ByVal x As Double, ByVal y As Double, ByVal rad As Double) As Boolean
    Dim i As Long
    For i = 0 To NBC - 1
        If Abs(BCX(i) - x) < rad Then
            If Dist(BCX(i), BCY(i), x, y) < rad Then
                NearBCorner = True
                Exit Function
            End If
        End If
    Next i
End Function

' ---------------------------------------------------------------------
'  Обход зоны развилки по окружности (или прямой), подогнанной к точкам
'  с обеих сторон: сквозная линия не "проседает" к развилке.
'  pT: 1 — окружность (pA,pB центр, pC радиус), 2 — прямая (pA,pB точка, pC,pD направление)
' ---------------------------------------------------------------------
Private Function ArcFill(xs() As Double, ys() As Double, ByVal np_ As Long, ax_() As Double, ay_() As Double, _
                         ByVal na As Long, ByVal stp As Double, ByVal w As Double, rx() As Double, ry() As Double, _
                         ByVal nr_ As Long, fx() As Double, fy() As Double, nf As Long, pf() As Double, xi As Long) As Boolean
    Dim bx() As Double, by() As Double, nb As Long, cx_() As Double, cy_() As Double, nc As Long
    Dim fa(0 To 5) As Double, fb(0 To 5) As Double
    Dim hasC As Boolean, hasS As Boolean, i As Long
    Dim xc() As Double, yc() As Double, fcx() As Double, fcy() As Double, nfc As Long, pfc(0 To 5) As Double, xic As Long
    Dim xsm() As Double, ysm() As Double, fsx() As Double, fsy() As Double, nfs As Long, pfs(0 To 5) As Double
    Dim tx() As Double, ty() As Double, nt As Long, dc As Double, ds As Double
    xi = -1: nf = 0
    If ShortSideLine(xs, ys, np_, ax_, ay_, na, stp, w, fx, fy, nf, pf) Then
        ArcFill = True
        Exit Function
    End If
    ' стороны по отдельности: если ось подходит к развилке под заметным углом (Y, V, k) —
    ' возможно, это угол, а не плавная дуга через две короткие стороны; кто прав — решают
    ' сырые узлы зоны. Составная кривая (прямая + дуга по касательной, xic = -1) точнее
    ' одного фита — берётся сразу.
    TakePts xs, ys, np_, True, 2 * w, bx, by, nb
    TakePts ax_, ay_, na, False, 2 * w, cx_, cy_, nc
    If nb > 1 And nc > 1 Then
        If FitPts(bx, by, nb, w, 0.02, 0.02, fa) And FitPts(cx_, cy_, nc, w, 0.02, 0.02, fb) Then
            ReDim xc(0 To np_ + 8): ReDim yc(0 To np_ + 8)
            For i = 0 To np_ - 1: xc(i) = xs(i): yc(i) = ys(i): Next i
            ReDim fcx(0 To 7): ReDim fcy(0 To 7)
            hasC = CornerFill(xc, yc, np_, ax_, ay_, na, stp, w, fcx, fcy, nfc, pfc, xic)
            If hasC And xic < 0 Then
                For i = 0 To np_ - 1: xs(i) = xc(i): ys(i) = yc(i): Next i
                nf = 0
                For i = 0 To nfc - 1: PushPt fx, fy, nf, fcx(i), fcy(i): Next i
                For i = 0 To 5: pf(i) = pfc(i): Next i
                ArcFill = True
                Exit Function
            End If
        End If
    End If
    ReDim xsm(0 To np_ + 8): ReDim ysm(0 To np_ + 8)
    For i = 0 To np_ - 1: xsm(i) = xs(i): ysm(i) = ys(i): Next i
    ReDim fsx(0 To 7): ReDim fsy(0 To 7)
    hasS = ArcFillSingle(xsm, ysm, np_, ax_, ay_, na, stp, w, fsx, fsy, nfs, pfs)
    If (Not hasC) And (Not hasS) Then
        If TangentBridge(xs, ys, np_, ax_, ay_, na, stp, w, fx, fy, nf, pf) Then
            ArcFill = True
            Exit Function
        End If
    End If
    If hasC And hasS Then
        ReDim tx(0 To nfc + 2): ReDim ty(0 To nfc + 2): nt = 0
        PushPt tx, ty, nt, xc(np_ - 1), yc(np_ - 1)
        For i = 0 To nfc - 1: PushPt tx, ty, nt, fcx(i), fcy(i): Next i
        PushPt tx, ty, nt, ax_(0), ay_(0)
        dc = PolyDist(rx, ry, nr_, tx, ty, nt)
        ReDim tx(0 To nfs + 2): ReDim ty(0 To nfs + 2): nt = 0
        PushPt tx, ty, nt, xsm(np_ - 1), ysm(np_ - 1)
        For i = 0 To nfs - 1: PushPt tx, ty, nt, fsx(i), fsy(i): Next i
        PushPt tx, ty, nt, ax_(0), ay_(0)
        ds = PolyDist(rx, ry, nr_, tx, ty, nt)
        If dc <= ds Then
            hasS = False
        Else
            hasC = False
        End If
    End If
    If hasC Then
        For i = 0 To np_ - 1: xs(i) = xc(i): ys(i) = yc(i): Next i
        nf = 0
        For i = 0 To nfc - 1: PushPt fx, fy, nf, fcx(i), fcy(i): Next i
        For i = 0 To 5: pf(i) = pfc(i): Next i
        xi = xic
        ArcFill = True
    ElseIf hasS Then
        For i = 0 To np_ - 1: xs(i) = xsm(i): ys(i) = ysm(i): Next i
        nf = 0
        For i = 0 To nfs - 1: PushPt fx, fy, nf, fsx(i), fsy(i): Next i
        For i = 0 To 5: pf(i) = pfs(i): Next i
        ArcFill = True
    End If
End Function

Private Sub ProjPt(ByVal pT As Long, ByVal pA As Double, ByVal pB As Double, ByVal pC As Double, ByVal pD As Double, _
                   x As Double, y As Double)
    Dim d As Double, t As Double
    If pT = 1 Then
        d = Dist(x, y, pA, pB)
        If d > 0.000000001 Then
            x = pA + (x - pA) * pC / d: y = pB + (y - pB) * pC / d
        End If
    ElseIf pT = 2 Then
        t = (x - pA) * pC + (y - pB) * pD
        x = pA + pC * t: y = pB + pD * t
    ElseIf pT = 3 Then
        ' кусочная: до T2 (pA,pB) — прямая, после — окружность BRC*
        t = (x - pA) * pC + (y - pB) * pD
        If t > 0 Then
            d = Dist(x, y, BRCX, BRCY)
            If d > 0.000000001 Then
                x = BRCX + (x - BRCX) * BRCR / d: y = BRCY + (y - BRCY) * BRCR / d
            End If
        Else
            x = pA + pC * t: y = pB + pD * t
        End If
    End If
End Sub

' окружность по МНК (Kasa)
Private Function FitCircle(x() As Double, y() As Double, ByVal n As Long, cx As Double, cy As Double, _
                           rr As Double, rms As Double) As Boolean
    Dim i As Long, mx As Double, my As Double, u As Double, v As Double
    Dim suu As Double, suv As Double, svv As Double, suuu As Double, svvv As Double, suvv As Double, svuu As Double
    Dim det As Double, b1 As Double, b2 As Double, uc As Double, vc As Double, e As Double
    rms = 0: cx = 0: cy = 0: rr = 0
    If n < 4 Then
        Exit Function
    End If
    For i = 0 To n - 1
        mx = mx + x(i): my = my + y(i)
    Next i
    mx = mx / n: my = my / n
    For i = 0 To n - 1
        u = x(i) - mx: v = y(i) - my
        suu = suu + u * u: suv = suv + u * v: svv = svv + v * v
        suuu = suuu + u * u * u: svvv = svvv + v * v * v: suvv = suvv + u * v * v: svuu = svuu + v * u * u
    Next i
    det = suu * svv - suv * suv
    If Abs(det) < 1E-12 Then
        Exit Function
    End If
    b1 = 0.5 * (suuu + suvv): b2 = 0.5 * (svvv + svuu)
    uc = (b1 * svv - b2 * suv) / det: vc = (suu * b2 - suv * b1) / det
    cx = uc + mx: cy = vc + my
    rr = Sqr(uc * uc + vc * vc + (suu + svv) / n)
    For i = 0 To n - 1
        e = Dist(x(i), y(i), cx, cy) - rr
        rms = rms + e * e
    Next i
    rms = Sqr(rms / n)
    FitCircle = True
End Function

Private Sub FitLineLS(x() As Double, y() As Double, ByVal n As Long, mx As Double, my As Double, _
                      dx As Double, dy As Double, rms As Double)
    Dim i As Long, sxx As Double, syy As Double, sxy As Double, ang As Double, e As Double
    mx = 0: my = 0: rms = 0
    For i = 0 To n - 1
        mx = mx + x(i): my = my + y(i)
    Next i
    mx = mx / n: my = my / n
    For i = 0 To n - 1
        sxx = sxx + (x(i) - mx) ^ 2: syy = syy + (y(i) - my) ^ 2: sxy = sxy + (x(i) - mx) * (y(i) - my)
    Next i
    ang = 0.5 * Atan2(2 * sxy, sxx - syy)
    dx = Cos(ang): dy = Sin(ang)
    For i = 0 To n - 1
        e = (x(i) - mx) * (-dy) + (y(i) - my) * dx
        rms = rms + e * e
    Next i
    rms = Sqr(rms / n)
End Sub

Private Function Atan2(ByVal y As Double, ByVal x As Double) As Double
    If x > 0 Then
        Atan2 = Atn(y / x)
    ElseIf x < 0 Then
        If y >= 0 Then
            Atan2 = Atn(y / x) + PI_
        Else
            Atan2 = Atn(y / x) - PI_
        End If
    Else
        If y > 0 Then
            Atan2 = PI_ / 2
        ElseIf y < 0 Then
            Atan2 = -PI_ / 2
        Else
            Atan2 = 0
        End If
    End If
End Function

' индекс точки на расстоянии L по ломаной от i (dirn = -1/+1), -1 если не хватило
Private Function ArcIdx(xs() As Double, ys() As Double, ByVal n As Long, ByVal closed As Boolean, _
                        ByVal i As Long, ByVal L As Double, ByVal dirn As Long) As Long
    Dim s As Double, j As Long, k As Long, cnt As Long
    j = i
    Do
        k = j + dirn
        If Not closed Then
            If k < 0 Or k >= n Then
                ArcIdx = -1
                Exit Function
            End If
        Else
            k = (k + n) Mod n
        End If
        s = s + Dist(xs(j), ys(j), xs(k), ys(k)): j = k
        If s >= L Then
            ArcIdx = j
            Exit Function
        End If
        cnt = cnt + 1
        If j = i Or cnt > n Then
            ArcIdx = -1
            Exit Function
        End If
    Loop
End Function

' острый угол там, где линия резко поворачивает на коротком участке у острого угла контура:
' скругление заменяется точкой пересечения касательных
Private Sub Sharpen(xs() As Double, ys() As Double, fixedF() As Boolean, np_ As Long, ByVal closed As Boolean, ByVal w As Double)
    Dim guard As Long, i As Long, a As Long, b As Long, bi As Long, bt As Double, t As Double
    Dim d1x As Double, d1y As Double, d2x As Double, d2y As Double, lim As Double, c As Long, k As Long, q As Long
    Dim skip() As Boolean, nearC() As Long, nnc As Long, hasNear As Boolean
    Dim ha() As Double, nha As Long, hb() As Double, nhb As Long, kA As Long, kB As Long, el() As Double, nel As Long
    Dim fa(0 To 5) As Double, fb(0 To 5) As Double, xX As Double, yX As Double, found As Boolean, ia As Long, ib As Long
    Dim da As Double, db As Double, ok As Boolean, xa As Double, ya As Double, xb As Double, yb As Double
    Dim j As Long, k2 As Long, s As Double, rem_() As Boolean, nx() As Double, ny() As Double, nfx() As Boolean, nn As Long
    Dim fx() As Double, fy() As Double, nf As Long, nleft As Long, m As Long
    Dim ncc As Long, bsx As Double, bsy As Double, dout As Double
    If np_ < 5 Then
        Exit Sub
    End If
    ncc = 0
    For c = 0 To NBC - 1
        If Not BCC(c) Then
            ncc = ncc + 1
        End If
    Next c
    If ncc = 0 Then
        Exit Sub
    End If
    lim = Cos(45 * PI_ / 180)
    ReDim skip(0 To np_ + 8)
    ReDim nearC(0 To NBC)
    For guard = 1 To 50
        bi = -1: bt = 2
        For i = 0 To np_ - 1
            If Not fixedF(i) And Not skip(i) Then
                hasNear = False
                For c = 0 To NBC - 1
                    If Not BCC(c) Then
                        If Abs(BCX(c) - xs(i)) < 1.8 * w Then
                            If Dist(BCX(c), BCY(c), xs(i), ys(i)) < 1.8 * w Then
                                hasNear = True
                                Exit For
                            End If
                        End If
                    End If
                Next c
                If hasNear Then
                    a = ArcIdxF(xs, ys, fixedF, np_, closed, i, 0.5 * w, -1)
                    b = ArcIdxF(xs, ys, fixedF, np_, closed, i, 0.5 * w, 1)
                    If a >= 0 And b >= 0 Then
                        d1x = xs(i) - xs(a): d1y = ys(i) - ys(a): Norm d1x, d1y
                        d2x = xs(b) - xs(i): d2y = ys(b) - ys(i): Norm d2x, d2y
                        t = d1x * d2x + d1y * d2y
                        If t < lim And t < bt Then
                            bt = t: bi = i
                        End If
                    End If
                End If
            End If
        Next i
        If bi < 0 Then
            Exit Sub
        End If
        i = bi
        ReDim ha(0 To 5, 0 To 7): ReDim hb(0 To 5, 0 To 7)
        If Not SideFit(xs, ys, np_, closed, i, w, -1, ha, nha, kA) Then
            skip(i) = True
            GoTo NextGuard
        End If
        If Not SideFit(xs, ys, np_, closed, i, w, 1, hb, nhb, kB) Then
            skip(i) = True
            GoTo NextGuard
        End If
        ' вогнутые углы рядом и прямые по их кромкам
        nnc = 0
        ReDim el(0 To 5, 0 To 7): nel = 0
        For c = 0 To NBC - 1
            If Not BCC(c) Then
                If Dist(BCX(c), BCY(c), xs(i), ys(i)) < 1.8 * w Then
                    nearC(nnc) = c: nnc = nnc + 1
                    EdgeLines c, w, el, nel
                End If
            End If
        Next c
        ' порядок гипотез: строгий фит по оси, прямые по кромкам, остальные фиты
        found = False
        For ia = -1 To nha + nel - 1
            If found Then
                Exit For
            End If
            If ia = -1 Then
                GetFit ha, 0, fa
            ElseIf ia < nel Then
                GetFit el, ia, fa
            Else
                If ia - nel + 1 >= nha Then
                    GoTo NextA
                End If
                GetFit ha, ia - nel + 1, fa
            End If
            ' гипотеза должна проходить через свою сторону (кромочные прямые общие для обеих)
            If FitDist(fa, xs(kA), ys(kA)) > 0.3 * w Then
                GoTo NextA
            End If
            For ib = -1 To nhb + nel - 1
                If ib = -1 Then
                    GetFit hb, 0, fb
                ElseIf ib < nel Then
                    GetFit el, ib, fb
                Else
                    If ib - nel + 1 >= nhb Then
                        GoTo NextB
                    End If
                    GetFit hb, ib - nel + 1, fb
                End If
                If fa(5) = fb(5) Then
                    GoTo NextB
                End If
                If FitDist(fb, xs(kB), ys(kB)) > 0.3 * w Then
                    GoTo NextB
                End If
                If Not FitIntersect(fa, fb, xs(i), ys(i), xX, yX) Then
                    GoTo NextB
                End If
                If Dist(xX, yX, xs(i), ys(i)) >= 2 * w Then
                    GoTo NextB
                End If
                If Not InUnion(xX, yX, -1) Then
                    GoTo NextB
                End If
                For k = 0 To nnc - 1
                    c = nearC(k)
                    da = FitDist(fa, BCX(c), BCY(c)): db = FitDist(fb, BCX(c), BCY(c))
                    If da >= 0.3 * w And da <= 0.8 * w And db >= 0.3 * w And db <= 0.8 * w Then
                        ok = False
                        If EdgeMatch(fa, c, w, BE1X(c), BE1Y(c)) And EdgeMatch(fb, c, w, BE2X(c), BE2Y(c)) Then
                            ok = True
                        ElseIf EdgeMatch(fa, c, w, BE2X(c), BE2Y(c)) And EdgeMatch(fb, c, w, BE1X(c), BE1Y(c)) Then
                            ok = True
                        End If
                        If ok Then
                            ' тугой завиток (рукописная s) снаружи круглый с большим радиусом:
                            ' ось режет угол, и до внешней кромки от X по биссектрисе меньше
                            ' 0.75 полуширины; у настоящего угла (стык штрихов) — около полуширины
                            bsx = xX - BCX(c): bsy = yX - BCY(c): Norm bsx, bsy
                            dout = RayUnion(xX, yX, bsx, bsy)
                            If dout < BIG And dout < 0.75 * (da + db) / 2 Then
                                ok = False
                            End If
                        End If
                        If ok Then
                            found = True
                            Exit For
                        End If
                    End If
                Next k
                If found Then
                    Exit For
                End If
NextB:
            Next ib
NextA:
        Next ia
        If Not found Then
            skip(i) = True
            GoTo NextGuard
        End If
        ' у острого угла ось искажена дальше 0.9w: расширяем вырезаемый участок,
        ' пока точки не лягут на кривую своей стороны (не дальше 2.1w)
        a = kA: s = 0: j = kA
        Do
            k2 = j - 1
            If Not closed Then
                If k2 < 0 Then
                    Exit Do
                End If
            Else
                k2 = (k2 + np_) Mod np_
            End If
            If k2 = i Or fixedF(k2) Then
                Exit Do
            End If
            s = s + Dist(xs(j), ys(j), xs(k2), ys(k2))
            If s > 2.1 * w Or FitDist(fa, xs(k2), ys(k2)) <= 0.03 * w Then
                Exit Do
            End If
            j = k2
        Loop
        a = j
        s = 0: j = kB
        Do
            k2 = j + 1
            If Not closed Then
                If k2 >= np_ Then
                    Exit Do
                End If
            Else
                k2 = k2 Mod np_
            End If
            If k2 = i Or fixedF(k2) Then
                Exit Do
            End If
            s = s + Dist(xs(j), ys(j), xs(k2), ys(k2))
            If s > 2.1 * w Or FitDist(fb, xs(k2), ys(k2)) <= 0.03 * w Then
                Exit Do
            End If
            j = k2
        Loop
        b = j
        ' новая ломаная: ... a, [кривая a->X], X, [кривая X->b], b ...
        ReDim rem_(0 To np_ - 1)
        q = a
        Do
            q = q + 1
            If closed Then
                q = q Mod np_
            End If
            If q = b Or q >= np_ Then
                Exit Do
            End If
            rem_(q) = True
        Loop
        xa = xX: ya = yX: FitProj fa, xa, ya
        xb = xX: yb = yX: FitProj fb, xb, yb
        ReDim fx(0 To 15): ReDim fy(0 To 15): nf = 0
        FitBetween fa, xs(a), ys(a), xa, ya, gH, fx, fy, nf
        nleft = nf
        PushPt fx, fy, nf, xX, yX
        FitBetween fb, xb, yb, xs(b), ys(b), gH, fx, fy, nf
        ReDim nx(0 To np_ + nf + 8): ReDim ny(0 To np_ + nf + 8): ReDim nfx(0 To np_ + nf + 8)
        nn = 0
        For q = 0 To np_ - 1
            If Not rem_(q) Then
                nx(nn) = xs(q): ny(nn) = ys(q): nfx(nn) = fixedF(q): nn = nn + 1
                If q = a Then
                    For m = 0 To nf - 1
                        nx(nn) = fx(m): ny(nn) = fy(m): nfx(nn) = (m = nleft): nn = nn + 1
                    Next m
                    If NSh > UBound(ShX) Then
                        ReDim Preserve ShX(0 To 2 * NSh): ReDim Preserve ShY(0 To 2 * NSh)
                    End If
                    ShX(NSh) = xX: ShY(NSh) = yX: NSh = NSh + 1
                End If
            End If
        Next q
        np_ = nn
        ReDim xs(0 To nn + 8): ReDim ys(0 To nn + 8): ReDim fixedF(0 To nn + 8)
        For q = 0 To nn - 1
            xs(q) = nx(q): ys(q) = ny(q): fixedF(q) = nfx(q)
        Next q
        ReDim skip(0 To np_ + 8)
NextGuard:
    Next guard
End Sub

' копия массива (отдельный объект — важно для хранения в Variant)
Private Function CopyD(src() As Double, ByVal n As Long) As Variant
    Dim a() As Double, i As Long
    ReDim a(0 To n + 8)
    For i = 0 To n - 1
        a(i) = src(i)
    Next i
    CopyD = a
End Function

Private Sub RotateIds(ids() As Long, flags() As Boolean, ByVal n As Long, ByVal r0 As Long)
    Dim a() As Long, f() As Boolean, i As Long
    ReDim a(0 To n - 1): ReDim f(0 To n - 1)
    For i = 0 To n - 1
        a(i) = ids((i + r0) Mod n): f(i) = flags((i + r0) Mod n)
    Next i
    For i = 0 To n - 1
        ids(i) = a(i): flags(i) = f(i)
    Next i
End Sub

Private Sub InsertFront(xs() As Double, ys() As Double, fixedF() As Boolean, np_ As Long, ByVal x As Double, ByVal y As Double)
    Dim i As Long
    PushPtF xs, ys, fixedF, np_, 0, 0, False
    For i = np_ - 1 To 1 Step -1
        xs(i) = xs(i - 1): ys(i) = ys(i - 1): fixedF(i) = fixedF(i - 1)
    Next i
    xs(0) = x: ys(0) = y: fixedF(0) = False
End Sub

' ломаная (x0,y0) -> fill... -> (nx,ny) целиком внутри фигуры?
Private Function PathInside(ByVal x0 As Double, ByVal y0 As Double, fx() As Double, fy() As Double, ByVal nf As Long, _
                            ByVal nx As Double, ByVal ny As Double, ByVal hasNxt As Boolean) As Boolean
    Dim i As Long, ax0 As Double, ay0 As Double, bx0 As Double, by0 As Double
    ax0 = x0: ay0 = y0
    For i = 0 To nf
        If i < nf Then
            bx0 = fx(i): by0 = fy(i)
        ElseIf hasNxt Then
            bx0 = nx: by0 = ny
        Else
            Exit For
        End If
        If Not SegInside(ax0, ay0, bx0, by0) Then
            Exit Function
        End If
        ax0 = bx0: ay0 = by0
    Next i
    PathInside = True
End Function

Private Function SegInside(ByVal x0 As Double, ByVal y0 As Double, ByVal x1 As Double, ByVal y1 As Double) As Boolean
    Dim L As Double, n As Long, i As Long, t As Double
    L = Dist(x0, y0, x1, y1)
    n = Int(L / (0.5 * gH))
    If n < 1 Then
        n = 1
    End If
    If n > 64 Then
        n = 64
    End If
    For i = 1 To n
        t = i / (n + 1)
        If Not InUnion(x0 + (x1 - x0) * t, y0 + (y1 - y0) * t, -1) Then
            Exit Function
        End If
    Next i
    SegInside = True
End Function

Private Sub FreeEnd(xs() As Double, ys() As Double, np_ As Long, ByVal cref As Double, ByVal nmax As Long)
    Dim bx() As Double, by() As Double, nb As Long, i As Long, dx As Double, dy As Double
    nb = np_
    ReDim bx(0 To np_ + 8): ReDim by(0 To np_ + 8)
    For i = 0 To np_ - 1: bx(i) = xs(i): by(i) = ys(i): Next i
    FreeEndFit xs, ys, np_, cref, nmax
    If np_ >= 2 Then
        dx = xs(np_ - 1) - xs(np_ - 2): dy = ys(np_ - 1) - ys(np_ - 2): Norm dx, dy
        If Centered(xs(np_ - 1), ys(np_ - 1), dx, dy, cref) Then
            Exit Sub
        End If
    End If
    np_ = nb
    If np_ > UBound(xs) Then
        ReDim Preserve xs(0 To np_ + 8): ReDim Preserve ys(0 To np_ + 8)
    End If
    For i = 0 To np_ - 1: xs(i) = bx(i): ys(i) = by(i): Next i
    FreeEndRay xs, ys, np_, cref
End Sub

' направление наружу у последней точки, по участку ~0.7 ширины
Private Sub TanEnd(xs() As Double, ys() As Double, ByVal n As Long, ByVal w As Double, dx As Double, dy As Double)
    Dim j As Long, L As Double
    j = n - 1
    Do While j > 0 And L < w * 0.7
        L = L + Dist(xs(j), ys(j), xs(j - 1), ys(j - 1)): j = j - 1
    Loop
    dx = xs(n - 1) - xs(j): dy = ys(n - 1) - ys(j)
    Norm dx, dy
End Sub

Private Sub HermiteFill(xs() As Double, ys() As Double, ByVal np_ As Long, ax_() As Double, ay_() As Double, _
                       ByVal na As Long, ByVal stp As Double, ByVal jc As Long, fx() As Double, fy() As Double, nf As Long)
    Dim a0 As Double, a1 As Double, b0 As Double, b1 As Double, cx As Double, cy As Double
    Dim tax As Double, tay As Double, tbx As Double, tby As Double, win As Double
    a0 = xs(np_ - 1): a1 = ys(np_ - 1): b0 = ax_(0): b1 = ay_(0)
    cx = b0 - a0: cy = b1 - a1: Norm cx, cy
    win = stp * 6
    If Not EndDir(xs, ys, np_, True, win, tax, tay) Then
        tax = cx: tay = cy
    End If
    If EndDir(ax_, ay_, na, False, win, tbx, tby) Then
        tbx = -tbx: tby = -tby
    Else
        tbx = cx: tby = cy
    End If
    If jc >= 0 Then
        HermSeg fx, fy, nf, a0, a1, tax, tay, GX(jc), GY(jc), cx, cy, stp
        PushPt fx, fy, nf, GX(jc), GY(jc)
        HermSeg fx, fy, nf, GX(jc), GY(jc), cx, cy, b0, b1, tbx, tby, stp
    Else
        HermSeg fx, fy, nf, a0, a1, tax, tay, b0, b1, tbx, tby, stp
    End If
End Sub

Private Sub HermSeg(fx() As Double, fy() As Double, nf As Long, ByVal a0 As Double, ByVal a1 As Double, _
                    ByVal tax As Double, ByVal tay As Double, ByVal b0 As Double, ByVal b1 As Double, _
                    ByVal tbx As Double, ByVal tby As Double, ByVal stp As Double)
    Dim L As Double, m As Long, k As Long, t As Double, h00 As Double, h10 As Double, h01 As Double, h11 As Double
    L = Dist(a0, a1, b0, b1)
    m = Int(L / stp)
    If m < 1 Then
        m = 1
    End If
    For k = 1 To m - 1
        t = k / m
        h00 = 2 * t * t * t - 3 * t * t + 1: h10 = t * t * t - 2 * t * t + t
        h01 = -2 * t * t * t + 3 * t * t: h11 = t * t * t - t * t
        PushPt fx, fy, nf, h00 * a0 + h10 * L * tax + h01 * b0 + h11 * L * tbx, _
                           h00 * a1 + h10 * L * tay + h01 * b1 + h11 * L * tby
    Next k
End Sub

' сглаживание Таубина (без "усадки" дуг); fixedF — неподвижные точки (острия)
Private Sub Smooth(xs() As Double, ys() As Double, fixedF() As Boolean, ByVal n As Long, ByVal closed As Boolean, ByVal it As Long)
    Dim qx() As Double, qy() As Double, k As Long, i As Long, f As Double, a As Long, b As Long, i0 As Long, i1 As Long
    If n < 3 Then
        Exit Sub
    End If
    ReDim qx(0 To n - 1): ReDim qy(0 To n - 1)
    If closed Then
        i0 = 0: i1 = n - 1
    Else
        i0 = 1: i1 = n - 2
    End If
    For k = 0 To 2 * it - 1
        If k Mod 2 = 0 Then
            f = 0.5
        Else
            f = -0.53
        End If
        For i = 0 To n - 1: qx(i) = xs(i): qy(i) = ys(i): Next i
        For i = i0 To i1
            If Not fixedF(i) Then
                a = (i - 1 + n) Mod n: b = (i + 1) Mod n
                qx(i) = xs(i) + f * ((xs(a) + xs(b)) / 2 - xs(i))
                qy(i) = ys(i) + f * ((ys(a) + ys(b)) / 2 - ys(i))
            End If
        Next i
        For i = 0 To n - 1: xs(i) = qx(i): ys(i) = qy(i): Next i
    Next k
End Sub

' ---------------------------------------------------------------------
'  Аппроксимация Безье (Schneider) + острые углы
' ---------------------------------------------------------------------
Private Sub FitPath(xs() As Double, ys() As Double, ByVal n0 As Long, ByVal closed As Boolean, ByVal wf As Double)
    Dim px_() As Double, py_() As Double, n As Long, i As Long, ang() As Double
    Dim corners() As Long, nco As Long, k As Long
    Dim t1x As Double, t1y As Double, segStart As Long, lmin As Double
    Dim qx() As Double, qy() As Double, c0 As Long

    lmin = P_LMIN * wf
    ReDim px_(0 To n0): ReDim py_(0 To n0)
    n = 0
    For i = 0 To n0 - 1
        If n = 0 Then
            px_(n) = xs(i): py_(n) = ys(i): n = n + 1
        ElseIf Dist(xs(i), ys(i), px_(n - 1), py_(n - 1)) > 0.000001 Then
            px_(n) = xs(i): py_(n) = ys(i): n = n + 1
        End If
    Next i
    If closed And n > 1 Then
        If Dist(px_(0), py_(0), px_(n - 1), py_(n - 1)) < 0.000001 Then
            n = n - 1
        End If
    End If
    If n < 2 Then
        Exit Sub
    End If
    If closed And n < 3 Then
        closed = False
    End If

    segStart = NOS
    AddPath closed

    ReDim ang(0 To n - 1)
    For i = 0 To n - 1
        ang(i) = TurnAt(px_, py_, n, i, closed)
    Next i
    ReDim corners(0 To n + 1): nco = 0
    For i = 0 To n - 1
        If IsSharpPt(px_(i), py_(i)) Or (ang(i) > P_CORNER And ang(i) >= ang((i - 1 + n) Mod n) And ang(i) >= ang((i + 1) Mod n)) Then
            If closed Or (i > 0 And i < n - 1) Then
                If nco = 0 Then
                    corners(nco) = i: nco = nco + 1
                ElseIf i - corners(nco - 1) >= 3 Then
                    corners(nco) = i: nco = nco + 1
                End If
            End If
        End If
    Next i

    If closed Then
        If nco = 0 Then
            ReDim qx(0 To n): ReDim qy(0 To n)
            For i = 0 To n - 1: qx(i) = px_(i): qy(i) = py_(i): Next i
            qx(n) = px_(0): qy(n) = py_(0)
            t1x = px_(1) - px_(n - 1): t1y = py_(1) - py_(n - 1): Norm t1x, t1y
            FitCubic qx, qy, 0, n, t1x, t1y, -t1x, -t1y, 0, lmin
        Else
            c0 = corners(0)
            ReDim qx(0 To n): ReDim qy(0 To n)
            For i = 0 To n
                qx(i) = px_((c0 + i) Mod n): qy(i) = py_((c0 + i) Mod n)
            Next i
            For k = 0 To nco - 1
                corners(k) = corners(k) - c0
                If corners(k) < 0 Then
                    corners(k) = corners(k) + n
                End If
            Next k
            corners(nco) = n: nco = nco + 1
            For k = 0 To nco - 2
                FitPiece qx, qy, corners(k), corners(k + 1), lmin
            Next k
        End If
    Else
        For k = nco To 1 Step -1: corners(k) = corners(k - 1): Next k
        corners(0) = 0: nco = nco + 1
        corners(nco) = n - 1: nco = nco + 1
        For k = 0 To nco - 2
            FitPiece px_, py_, corners(k), corners(k + 1), lmin
        Next k
    End If
    MergeLines segStart
End Sub

' точка поставлена Sharpen'ом (острый узел)?
Private Function IsSharpPt(ByVal x As Double, ByVal y As Double) As Boolean
    Dim i As Long
    For i = 0 To NSh - 1
        If Abs(ShX(i) - x) < 0.000001 Then
            If Abs(ShY(i) - y) < 0.000001 Then
                IsSharpPt = True
                Exit Function
            End If
        End If
    Next i
End Function

Private Sub FitPiece(x() As Double, y() As Double, ByVal a As Long, ByVal b As Long, ByVal lmin As Double)
    Dim t1x As Double, t1y As Double, t2x As Double, t2y As Double, j As Long, L As Double
    If b <= a Then
        Exit Sub
    End If
    j = a + 1: L = 0
    Do While j < b
        L = L + Dist(x(j), y(j), x(j - 1), y(j - 1))
        If L >= gW * 0.25 Then
            Exit Do
        End If
        j = j + 1
    Loop
    t1x = x(j) - x(a): t1y = y(j) - y(a): Norm t1x, t1y
    j = b - 1: L = 0
    Do While j > a
        L = L + Dist(x(j), y(j), x(j + 1), y(j + 1))
        If L >= gW * 0.25 Then
            Exit Do
        End If
        j = j - 1
    Loop
    t2x = x(j) - x(b): t2y = y(j) - y(b): Norm t2x, t2y
    FitCubic x, y, a, b, t1x, t1y, t2x, t2y, 0, lmin
End Sub

Private Function TurnAt(x() As Double, y() As Double, ByVal n As Long, ByVal i As Long, ByVal closed As Boolean) As Double
    Dim a As Long, b As Long, L As Double, a2 As Long, d1x As Double, d1y As Double, d2x As Double, d2y As Double
    a = i: L = 0
    Do While L < gW * 0.5
        If Not closed And a = 0 Then
            Exit Do
        End If
        a2 = (a - 1 + n) Mod n: L = L + Dist(x(a), y(a), x(a2), y(a2)): a = a2
        If a = i Then
            Exit Do
        End If
    Loop
    b = i: L = 0
    Do While L < gW * 0.5
        If Not closed And b = n - 1 Then
            Exit Do
        End If
        a2 = (b + 1) Mod n: L = L + Dist(x(b), y(b), x(a2), y(a2)): b = a2
        If b = i Then
            Exit Do
        End If
    Loop
    If a = i Or b = i Then
        Exit Function
    End If
    d1x = x(i) - x(a): d1y = y(i) - y(a): Norm d1x, d1y
    d2x = x(b) - x(i): d2y = y(b) - y(i): Norm d2x, d2y
    TurnAt = ACosD(d1x * d2x + d1y * d2y)
End Function

Private Function LineDev(x() As Double, y() As Double, ByVal i0 As Long, ByVal i1 As Long) As Double
    Dim L As Double, i As Long, d As Double, m As Double
    L = Dist(x(i0), y(i0), x(i1), y(i1))
    If L < 0.000000001 Then
        LineDev = BIG
        Exit Function
    End If
    For i = i0 To i1
        d = Abs((x(i1) - x(i0)) * (y(i0) - y(i)) - (x(i0) - x(i)) * (y(i1) - y(i0))) / L
        If d > m Then
            m = d
        End If
    Next i
    LineDev = m
End Function

Private Sub FitCubic(x() As Double, y() As Double, ByVal i0 As Long, ByVal i1 As Long, _
                     ByVal t1x As Double, ByVal t1y As Double, ByVal t2x As Double, ByVal t2y As Double, ByVal depth As Long, _
                     ByVal lmin As Double)
    Dim n As Long, u() As Double, i As Long, bz(0 To 7) As Double, errV As Double, idx As Long, it As Long
    Dim tcx As Double, tcy As Double, k As Long, a As Long, m As Long, splitL As Boolean
    n = i1 - i0 + 1
    If n <= 2 Or depth > 40 Then
        EmitSeg 0, x(i0), y(i0), 0, 0, 0, 0, x(i1), y(i1)
        Exit Sub
    End If
    ' прямая — только если участок действительно прямой (0.4 допуска), иначе
    ' пологие дуги рассыпаются на цепочку коротких отрезков
    If LineDev(x, y, i0, i1) <= 0.4 * gTol Then
        EmitSeg 0, x(i0), y(i0), 0, 0, 0, 0, x(i1), y(i1)
        Exit Sub
    End If
    If lmin > 0 And n > 3 Then
        ' длинный прямой участок в начале и/или конце куска — отдельным отрезком, остаток
        ' по касательной к нему (низ чаши P, перекладина B: прямая + дуга, а не одна кривая).
        ' Только на уровне куска: внутри остатка обычное деление
        a = i0: m = i1
        k = StraightRun(x, y, i0, i1, 1, 0.4 * gTol, lmin)
        If k >= 0 Then
            EmitSeg 0, x(i0), y(i0), 0, 0, 0, 0, x(k), y(k)
            t1x = x(k) - x(i0): t1y = y(k) - y(i0): Norm t1x, t1y
            a = k: splitL = True
        End If
        k = StraightRun(x, y, i1, a, -1, 0.4 * gTol, lmin)
        If k >= 0 Then
            m = k
            t2x = x(m) - x(i1): t2y = y(m) - y(i1): Norm t2x, t2y
            splitL = True
        End If
        If splitL Then
            If m - a >= 1 Then
                FitCubic x, y, a, m, t1x, t1y, t2x, t2y, depth, 0
            End If
            If m < i1 Then
                EmitSeg 0, x(m), y(m), 0, 0, 0, 0, x(i1), y(i1)
            End If
            Exit Sub
        End If
    End If
    ReDim u(0 To n - 1)
    u(0) = 0
    For i = 1 To n - 1
        u(i) = u(i - 1) + Dist(x(i0 + i - 1), y(i0 + i - 1), x(i0 + i), y(i0 + i))
    Next i
    For i = 1 To n - 1: u(i) = u(i) / u(n - 1): Next i
    GenBez x, y, i0, i1, u, t1x, t1y, t2x, t2y, bz
    errV = MaxErr(x, y, i0, i1, u, bz, idx)
    If errV <= gTol Then
        EmitSeg 1, bz(0), bz(1), bz(2), bz(3), bz(4), bz(5), bz(6), bz(7)
        Exit Sub
    End If
    If errV <= gTol * 4 Then
        For it = 1 To 6
            Reparam x, y, i0, i1, u, bz
            GenBez x, y, i0, i1, u, t1x, t1y, t2x, t2y, bz
            errV = MaxErr(x, y, i0, i1, u, bz, idx)
            If errV <= gTol Then
                EmitSeg 1, bz(0), bz(1), bz(2), bz(3), bz(4), bz(5), bz(6), bz(7)
                Exit Sub
            End If
        Next it
    End If
    If idx < 1 Then
        idx = 1
    End If
    If idx > n - 2 Then
        idx = n - 2
    End If
    tcx = x(i0 + idx - 1) - x(i0 + idx + 1): tcy = y(i0 + idx - 1) - y(i0 + idx + 1): Norm tcx, tcy
    FitCubic x, y, i0, i0 + idx, t1x, t1y, tcx, tcy, depth + 1, lmin
    FitCubic x, y, i0 + idx, i1, -tcx, -tcy, t2x, t2y, depth + 1, lmin
End Sub

Private Sub GenBez(x() As Double, y() As Double, ByVal i0 As Long, ByVal i1 As Long, u() As Double, _
                   ByVal t1x As Double, ByVal t1y As Double, ByVal t2x As Double, ByVal t2y As Double, bz() As Double)
    Dim i As Long, t As Double, mt As Double, b0 As Double, b1 As Double, b2 As Double, b3 As Double
    Dim c00 As Double, c01 As Double, c11 As Double, x0 As Double, x1 As Double
    Dim a1x As Double, a1y As Double, a2x As Double, a2y As Double, tx As Double, ty As Double
    Dim p0x As Double, p0y As Double, p3x As Double, p3y As Double, det As Double, al1 As Double, al2 As Double, seg As Double
    p0x = x(i0): p0y = y(i0): p3x = x(i1): p3y = y(i1)
    For i = 0 To i1 - i0
        t = u(i): mt = 1 - t
        b0 = mt * mt * mt: b1 = 3 * mt * mt * t: b2 = 3 * mt * t * t: b3 = t * t * t
        a1x = t1x * b1: a1y = t1y * b1: a2x = t2x * b2: a2y = t2y * b2
        c00 = c00 + a1x * a1x + a1y * a1y
        c01 = c01 + a1x * a2x + a1y * a2y
        c11 = c11 + a2x * a2x + a2y * a2y
        tx = x(i0 + i) - (p0x * (b0 + b1) + p3x * (b2 + b3))
        ty = y(i0 + i) - (p0y * (b0 + b1) + p3y * (b2 + b3))
        x0 = x0 + a1x * tx + a1y * ty
        x1 = x1 + a2x * tx + a2y * ty
    Next i
    det = c00 * c11 - c01 * c01
    seg = Dist(p0x, p0y, p3x, p3y)
    If Abs(det) > 1E-12 Then
        al1 = (x0 * c11 - x1 * c01) / det
        al2 = (c00 * x1 - c01 * x0) / det
    Else
        al1 = -1: al2 = -1
    End If
    If al1 < seg * 0.001 Or al2 < seg * 0.001 Or al1 > seg * 3 Or al2 > seg * 3 Then
        al1 = seg / 3: al2 = seg / 3
    End If
    bz(0) = p0x: bz(1) = p0y
    bz(2) = p0x + t1x * al1: bz(3) = p0y + t1y * al1
    bz(4) = p3x + t2x * al2: bz(5) = p3y + t2y * al2
    bz(6) = p3x: bz(7) = p3y
End Sub

Private Sub BezPt(bz() As Double, ByVal t As Double, qx As Double, qy As Double)
    Dim mt As Double
    mt = 1 - t
    qx = mt * mt * mt * bz(0) + 3 * mt * mt * t * bz(2) + 3 * mt * t * t * bz(4) + t * t * t * bz(6)
    qy = mt * mt * mt * bz(1) + 3 * mt * mt * t * bz(3) + 3 * mt * t * t * bz(5) + t * t * t * bz(7)
End Sub

Private Function MaxErr(x() As Double, y() As Double, ByVal i0 As Long, ByVal i1 As Long, u() As Double, _
                        bz() As Double, idx As Long) As Double
    Dim i As Long, qx As Double, qy As Double, d As Double, m As Double
    idx = (i1 - i0) \ 2
    For i = 1 To i1 - i0 - 1
        BezPt bz, u(i), qx, qy
        d = (qx - x(i0 + i)) * (qx - x(i0 + i)) + (qy - y(i0 + i)) * (qy - y(i0 + i))
        If d > m Then
            m = d: idx = i
        End If
    Next i
    MaxErr = Sqr(m)
End Function

Private Sub Reparam(x() As Double, y() As Double, ByVal i0 As Long, ByVal i1 As Long, u() As Double, bz() As Double)
    Dim i As Long, t As Double, mt As Double, qx As Double, qy As Double
    Dim d1x As Double, d1y As Double, d2x As Double, d2y As Double, num As Double, den As Double
    For i = 0 To i1 - i0
        t = u(i): mt = 1 - t
        BezPt bz, t, qx, qy
        d1x = 3 * mt * mt * (bz(2) - bz(0)) + 6 * mt * t * (bz(4) - bz(2)) + 3 * t * t * (bz(6) - bz(4))
        d1y = 3 * mt * mt * (bz(3) - bz(1)) + 6 * mt * t * (bz(5) - bz(3)) + 3 * t * t * (bz(7) - bz(5))
        d2x = 6 * mt * (bz(4) - 2 * bz(2) + bz(0)) + 6 * t * (bz(6) - 2 * bz(4) + bz(2))
        d2y = 6 * mt * (bz(5) - 2 * bz(3) + bz(1)) + 6 * t * (bz(7) - 2 * bz(5) + bz(3))
        num = (qx - x(i0 + i)) * d1x + (qy - y(i0 + i)) * d1y
        den = d1x * d1x + d1y * d1y + (qx - x(i0 + i)) * d2x + (qy - y(i0 + i)) * d2y
        If Abs(den) > 1E-12 Then
            t = t - num / den
            If t < 0 Then
                t = 0
            End If
            If t > 1 Then
                t = 1
            End If
            u(i) = t
        End If
    Next i
End Sub

Private Sub AddPath(ByVal closed As Boolean)
    If NOP > UBound(OPathStart) Then
        ReDim Preserve OPathStart(0 To 2 * NOP): ReDim Preserve OPathClosed(0 To 2 * NOP)
        ReDim Preserve OPathGroup(0 To 2 * NOP)
    End If
    OPathStart(NOP) = NOS: OPathClosed(NOP) = closed: OPathGroup(NOP) = gGroup: NOP = NOP + 1
End Sub

Private Sub EmitSeg(ByVal typ As Long, ByVal x0 As Double, ByVal y0 As Double, ByVal x1 As Double, ByVal y1 As Double, _
                    ByVal x2 As Double, ByVal y2 As Double, ByVal x3 As Double, ByVal y3 As Double)
    If NOS > UBound(OSeg, 2) Then
        ReDim Preserve OSeg(0 To 8, 0 To 2 * NOS)
    End If
    OSeg(0, NOS) = typ
    OSeg(1, NOS) = x0: OSeg(2, NOS) = y0: OSeg(3, NOS) = x1: OSeg(4, NOS) = y1
    OSeg(5, NOS) = x2: OSeg(6, NOS) = y2: OSeg(7, NOS) = x3: OSeg(8, NOS) = y3
    NOS = NOS + 1
End Sub

' слить подряд идущие почти коллинеарные отрезки
Private Sub MergeLines(ByVal s0 As Long)
    Dim i As Long, w As Long, x0 As Double, y0 As Double, mx As Double, my As Double, x1 As Double, y1 As Double
    Dim L As Double, d As Double, k As Long, merged As Boolean
    w = s0
    For i = s0 To NOS - 1
        merged = False
        If w > s0 Then
            If OSeg(0, i) = 0 And OSeg(0, w - 1) = 0 Then
                x0 = OSeg(1, w - 1): y0 = OSeg(2, w - 1)
                mx = OSeg(7, w - 1): my = OSeg(8, w - 1)
                x1 = OSeg(7, i): y1 = OSeg(8, i)
                L = Dist(x0, y0, x1, y1)
                If L > 0.000000001 Then
                    d = Abs((x1 - x0) * (y0 - my) - (x0 - mx) * (y1 - y0)) / L
                    If d <= gTol Then
                        OSeg(7, w - 1) = x1: OSeg(8, w - 1) = y1
                        merged = True
                    End If
                End If
            End If
        End If
        If Not merged Then
            If w <> i Then
                For k = 0 To 8: OSeg(k, w) = OSeg(k, i): Next k
            End If
            w = w + 1
        End If
    Next i
    NOS = w
End Sub

' ---------------------------------------------------------------------
'  Геометрия
' ---------------------------------------------------------------------
Private Function Dist(ByVal x1 As Double, ByVal y1 As Double, ByVal x2 As Double, ByVal y2 As Double) As Double
    Dist = Sqr((x2 - x1) * (x2 - x1) + (y2 - y1) * (y2 - y1))
End Function

Private Function Norm(dx As Double, dy As Double) As Boolean
    Dim L As Double
    L = Sqr(dx * dx + dy * dy)
    If L > 1E-12 Then
        dx = dx / L: dy = dy / L: Norm = True
    Else
        dx = 0: dy = 0
    End If
End Function

Private Function ACosD(ByVal c As Double) As Double
    If c >= 1 Then
        ACosD = 0
        Exit Function
    End If
    If c <= -1 Then
        ACosD = 180
        Exit Function
    End If
    ACosD = (Atn(-c / Sqr(1 - c * c)) + 2 * Atn(1)) * 180 / PI_
End Function

Private Sub PushPt(xs() As Double, ys() As Double, np_ As Long, ByVal x As Double, ByVal y As Double)
    If np_ > UBound(xs) Then
        ReDim Preserve xs(0 To 2 * np_ + 8): ReDim Preserve ys(0 To 2 * np_ + 8)
    End If
    xs(np_) = x: ys(np_) = y: np_ = np_ + 1
End Sub

Private Sub PushPtF(xs() As Double, ys() As Double, fixedF() As Boolean, np_ As Long, ByVal x As Double, ByVal y As Double, ByVal fxd As Boolean)
    If np_ > UBound(xs) Then
        ReDim Preserve xs(0 To 2 * np_ + 8): ReDim Preserve ys(0 To 2 * np_ + 8)
    End If
    If np_ > UBound(fixedF) Then
        ReDim Preserve fixedF(0 To 2 * np_ + 8)
    End If
    xs(np_) = x: ys(np_) = y: fixedF(np_) = fxd: np_ = np_ + 1
End Sub

Private Sub ReversePts(xs() As Double, ys() As Double, ByVal n As Long)
    Dim i As Long, j As Long, t As Double
    i = 0: j = n - 1
    Do While i < j
        t = xs(i): xs(i) = xs(j): xs(j) = t
        t = ys(i): ys(i) = ys(j): ys(j) = t
        i = i + 1: j = j - 1
    Loop
End Sub

' луч из (ox,oy) по (dx,dy): расстояние до края объединения (BIG — нет пересечения).
' попадания в контур, лежащие внутри другого полигона, пропускаются
Private Function RayUnion(ByVal ox As Double, ByVal oy As Double, ByVal dx As Double, ByVal dy As Double) As Double
    Dim r As Long, i As Long, j As Long, b As Long, n As Long, best As Double, lastT As Double
    Dim ax0 As Double, ay0 As Double, ex As Double, ey As Double, den As Double, t As Double, u As Double
    Dim candP As Long, guard As Long
    lastT = 0.000000001
    RayUnion = BIG
    Do
        best = BIG: candP = -1
        For r = 0 To NR - 1
            b = RS(r): n = RN(r)
            For i = 0 To n - 1
                j = (i + 1) Mod n
                ax0 = RX(b + i): ay0 = RY(b + i)
                ex = RX(b + j) - ax0: ey = RY(b + j) - ay0
                den = dx * ey - dy * ex
                If Abs(den) > 1E-12 Then
                    t = ((ax0 - ox) * ey - (ay0 - oy) * ex) / den
                    u = ((ax0 - ox) * dy - (ay0 - oy) * dx) / den
                    If t > lastT And u >= -0.000000001 And u <= 1.000000001 Then
                        If t < best Then
                            best = t: candP = RPoly(r)
                        End If
                    End If
                End If
            Next i
        Next r
        If best >= BIG Then
            Exit Function
        End If
        If NPoly = 1 Then
            RayUnion = best
            Exit Function
        End If
        If Not InUnion(ox + dx * best, oy + dy * best, candP) Then
            RayUnion = best
            Exit Function
        End If
        lastT = best + 0.000001
        guard = guard + 1
    Loop While guard < 200
End Function

Private Function RayPoly(x() As Double, y() As Double, ByVal n As Long, ByVal ox As Double, ByVal oy As Double, _
                         ByVal dx As Double, ByVal dy As Double) As Double
    Dim i As Long, best As Double, ax0 As Double, ay0 As Double, ex As Double, ey As Double
    Dim den As Double, t As Double, u As Double
    best = BIG
    For i = 0 To n - 2
        ax0 = x(i): ay0 = y(i)
        ex = x(i + 1) - ax0: ey = y(i + 1) - ay0
        den = dx * ey - dy * ex
        If Abs(den) > 1E-12 Then
            t = ((ax0 - ox) * ey - (ay0 - oy) * ex) / den
            u = ((ax0 - ox) * dy - (ay0 - oy) * dx) / den
            If t > -0.000000001 And u >= -0.000000001 And u <= 1.000000001 Then
                If t < best Then
                    best = t
                End If
            End If
        End If
    Next i
    RayPoly = best
End Function
Private Sub BranchDir(ByVal i As Long, ByVal j As Long, ByVal w As Double, dx As Double, dy As Double, _
                      farX As Double, farY As Double, nearX As Double, nearY As Double, _
                      midX As Double, midY As Double, rad As Double)
    Dim prv As Long, cur As Long, nxt As Long, L As Double, px_() As Double, py_() As Double, pl() As Double, np_ As Long
    Dim k As Long, m As Long, keep As Long, d0x As Double, d0y As Double, dkx As Double, dky As Double, has0 As Boolean
    Dim cut As Long, cx As Double, cy As Double, rr As Double, rms As Double, chx As Double, chy As Double
    Dim tx As Double, ty As Double, bestD As Double, ch As Double, cosLim As Double
    ReDim px_(0 To 63): ReDim py_(0 To 63): ReDim pl(0 To 63)
    rad = 0
    prv = i: cur = j: L = Dist(GX(i), GY(i), GX(j), GY(j))
    If L >= 0.3 * w Then
        px_(0) = GX(cur): py_(0) = GY(cur): pl(0) = L: np_ = 1
    End If
    Do While L < 2# * w
        If GDeg(cur) <> 2 Then
            Exit Do
        End If
        nxt = OtherNb(cur, prv)
        If nxt < 0 Then
            Exit Do
        End If
        prv = cur: cur = nxt
        L = L + Dist(GX(prv), GY(prv), GX(cur), GY(cur))
        If L >= 0.3 * w Then
            If np_ > UBound(px_) Then
                ReDim Preserve px_(0 To 2 * np_): ReDim Preserve py_(0 To 2 * np_): ReDim Preserve pl(0 To 2 * np_)
            End If
            px_(np_) = GX(cur): py_(np_) = GY(cur): pl(np_) = L: np_ = np_ + 1
        End If
    Loop
    ' ветка кончается свободным концом: узлы у скруглённого торца (крючок) — выбросить
    If GDeg(cur) = 1 And np_ > 0 Then
        keep = 0
        For k = 0 To np_ - 1
            If pl(k) <= L - 0.7 * w Then
                keep = keep + 1
            End If
        Next k
        If keep >= 2 Then
            np_ = keep
        End If
    End If
    ' излом внутри окна (перекладина Ц -> хвост): направление меряем до излома
    If np_ >= 3 Then
        has0 = False: cut = -1
        cosLim = Cos(50 * PI_ / 180)
        For k = 1 To np_ - 1
            m = k
            Do While m > 0 And pl(k) - pl(m) < 0.4 * w
                m = m - 1
            Loop
            If pl(k) - pl(m) >= 0.3 * w Then
                dkx = px_(k) - px_(m): dky = py_(k) - py_(m): Norm dkx, dky
                If Not has0 Then
                    d0x = dkx: d0y = dky: has0 = True
                ElseIf d0x * dkx + d0y * dky < cosLim Then
                    cut = m
                    Exit For
                End If
            End If
        Next k
        If cut >= 2 Then
            np_ = cut + 1: L = pl(cut)
        End If
    End If
    If np_ > 0 Then
        farX = px_(np_ - 1): farY = py_(np_ - 1)
    Else
        farX = GX(cur): farY = GY(cur)
    End If
    If np_ < 2 Then
        dx = farX - GX(i): dy = farY - GY(i): Norm dx, dy
        nearX = farX: nearY = farY: midX = farX: midY = farY
        Exit Sub
    End If
    nearX = px_(0): nearY = py_(0)
    ' точка на ~1w от развилки
    bestD = BIG
    For k = 0 To np_ - 1
        If Abs(pl(k) - w) < bestD Then
            bestD = Abs(pl(k) - w): midX = px_(k): midY = py_(k)
        End If
    Next k
    chx = farX - nearX: chy = farY - nearY: Norm chx, chy
    If np_ >= 4 And L >= 1.5 * w Then
        If FitCircle(px_, py_, np_, cx, cy, rr, rms) Then
            ch = Dist(px_(0), py_(0), px_(np_ - 1), py_(np_ - 1))
            If rr < 25 * w And rms < 0.03 * w And ch * ch / (8 * rr) > 0.03 * w Then
                tx = -(nearY - cy): ty = nearX - cx: Norm tx, ty
                If tx * chx + ty * chy < 0 Then
                    tx = -tx: ty = -ty
                End If
                dx = tx: dy = ty: rad = rr
                Exit Sub
            End If
        End If
    End If
    dx = chx: dy = chy
End Sub


Private Function SegsCross(ByVal ax0 As Double, ByVal ay0 As Double, ByVal bx0 As Double, ByVal by0 As Double, _
                           ByVal cx0 As Double, ByVal cy0 As Double, ByVal dx0 As Double, ByVal dy0 As Double) As Boolean
    Dim o1 As Double, o2 As Double, o3 As Double, o4 As Double
    o1 = (bx0 - ax0) * (cy0 - ay0) - (by0 - ay0) * (cx0 - ax0)
    o2 = (bx0 - ax0) * (dy0 - ay0) - (by0 - ay0) * (dx0 - ax0)
    o3 = (dx0 - cx0) * (ay0 - cy0) - (dy0 - cy0) * (ax0 - cx0)
    o4 = (dx0 - cx0) * (by0 - cy0) - (dy0 - cy0) * (bx0 - cx0)
    SegsCross = ((o1 > 0) <> (o2 > 0)) And ((o3 > 0) <> (o4 > 0))
End Function


Private Function LineX(ByVal p1x As Double, ByVal p1y As Double, ByVal d1x As Double, ByVal d1y As Double, _
                       ByVal p2x As Double, ByVal p2y As Double, ByVal d2x As Double, ByVal d2y As Double, _
                       xX As Double, yX As Double) As Boolean
    Dim den As Double, t As Double
    den = d1x * d2y - d1y * d2x
    If Abs(den) < 0.000000001 Then
        Exit Function
    End If
    t = ((p2x - p1x) * d2y - (p2y - p1y) * d2x) / den
    xX = p1x + d1x * t: yX = p1y + d1y * t
    LineX = True
End Function


Private Function Aimed(ByVal dax As Double, ByVal day_ As Double, ByVal nax As Double, ByVal nay As Double, ByVal ra As Double, _
                       ByVal mbx As Double, ByVal mby As Double, ByVal rb As Double, ByVal w As Double) As Boolean
    Dim tol As Double, vx As Double, vy As Double
    tol = 0.35 * w
    If ra > 0 Then
        tol = tol + 0.5 * w * w / ra
    End If
    If rb > 0 Then
        tol = tol + 0.5 * w * w / rb
    End If
    vx = mbx - nax: vy = mby - nay
    Aimed = (Abs(vx * day_ - vy * dax) <= tol)
End Function


Private Function Bent(x() As Double, y() As Double, ByVal i0 As Long, ByVal i1 As Long, ByVal rr As Double, ByVal w As Double) As Boolean
    Dim ch As Double
    ch = Dist(x(i0), y(i0), x(i1), y(i1))
    Bent = (ch * ch / (8 * rr) > 0.03 * w)
End Function


Private Sub FitProj(f() As Double, x As Double, y As Double)
    Dim d As Double, t As Double
    If f(0) = 1 Then
        d = Dist(x, y, f(1), f(2))
        If d > 0.000000001 Then
            x = f(1) + (x - f(1)) * f(3) / d: y = f(2) + (y - f(2)) * f(3) / d
        End If
    ElseIf f(0) = 2 Then
        t = (x - f(1)) * f(3) + (y - f(2)) * f(4)
        x = f(1) + f(3) * t: y = f(2) + f(4) * t
    End If
End Sub


Private Sub FitAt(f() As Double, ByVal sx As Double, ByVal sy As Double, ByVal s As Double, x As Double, y As Double)
    Dim a As Double
    If f(0) = 1 Then
        a = Atan2(sy - f(2), sx - f(1)) + f(4) * s / f(3)
        x = f(1) + f(3) * Cos(a): y = f(2) + f(3) * Sin(a)
    Else
        x = sx + f(3) * s: y = sy + f(4) * s
    End If
End Sub


Private Function FitParam(f() As Double, ByVal x As Double, ByVal y As Double) As Double
    If f(0) = 2 Then
        FitParam = (x - f(1)) * f(3) + (y - f(2)) * f(4)
    Else
        FitParam = f(4) * Atan2(y - f(2), x - f(1)) * f(3)
    End If
End Function


Private Sub FitTangent(f() As Double, ByVal px0 As Double, ByVal py0 As Double, ByVal rx As Double, ByVal ry As Double, _
                       tx As Double, ty As Double)
    If f(0) = 2 Then
        tx = f(3): ty = f(4)
    Else
        tx = -(py0 - f(2)): ty = px0 - f(1): Norm tx, ty
    End If
    If tx * rx + ty * ry < 0 Then
        tx = -tx: ty = -ty
    End If
End Sub


Private Function FarFit(xs() As Double, ys() As Double, ByVal np_ As Long, ByVal w As Double, ByVal nearW As Double, _
                        ByVal farW As Double, ByVal nmax As Long, f() As Double) As Boolean
    Dim px_() As Double, py_() As Double, n As Long, i As Long, i0 As Long, s As Double
    Dim cx As Double, cy As Double, rr As Double, rms As Double, mx As Double, my As Double, dx As Double, dy As Double
    Dim a0 As Double, a1 As Double, dlt As Double
    f(0) = 0
    ReDim px_(0 To np_): ReDim py_(0 To np_)
    i0 = 1
    If nmax >= 0 Then
        If np_ - nmax + 1 > i0 Then
            i0 = np_ - nmax + 1
        End If
    End If
    For i = np_ - 1 To i0 Step -1
        s = s + Dist(xs(i), ys(i), xs(i - 1), ys(i - 1))
        If s > farW * w Then
            Exit For
        End If
        If s >= nearW * w Then
            px_(n) = xs(i - 1): py_(n) = ys(i - 1): n = n + 1
        End If
    Next i
    If n < 4 Then
        Exit Function
    End If
    If FitCircle(px_, py_, n, cx, cy, rr, rms) Then
        If rr < 25 * w And rms < 0.02 * w And Bent(px_, py_, 0, n - 1, rr, w) Then
            a0 = Atan2(py_(0) - cy, px_(0) - cx): a1 = Atan2(py_(n - 1) - cy, px_(n - 1) - cx)
            dlt = a0 - a1
            Do While dlt > PI_
                dlt = dlt - 2 * PI_
            Loop
            Do While dlt <= -PI_
                dlt = dlt + 2 * PI_
            Loop
            f(0) = 1: f(1) = cx: f(2) = cy: f(3) = rr
            If dlt >= 0 Then
                f(4) = 1
            Else
                f(4) = -1
            End If
            FarFit = True
            Exit Function
        End If
    End If
    FitLineLS px_, py_, n, mx, my, dx, dy, rms
    If rms < 0.02 * w Then
        If (px_(0) - px_(n - 1)) * dx + (py_(0) - py_(n - 1)) * dy < 0 Then
            dx = -dx: dy = -dy
        End If
        f(0) = 2: f(1) = mx: f(2) = my: f(3) = dx: f(4) = dy
        FarFit = True
    End If
End Function


Private Function ApplyFar(xs() As Double, ys() As Double, np_ As Long, ByVal w As Double, f() As Double, _
                          ByVal nearW As Double, ByVal maxdev As Double, ByVal nmax As Long) As Boolean
    Dim i As Long, i0 As Long, imin As Long, s As Double, x As Double, y As Double, k As Long
    Dim ox() As Double, oy() As Double, no As Long, ref As Double, tmax As Double, t As Double, d As Double, per As Double
    Dim devs() As Double, nt As Long, kk As Long, mxd As Double, tx() As Double, ty() As Double
    Dim mx As Double, my As Double, dx As Double, dy As Double, rms As Double, px0 As Double, py0 As Double
    Dim tgx As Double, tgy As Double, f4(0 To 5) As Double
    imin = 1
    If nmax >= 0 Then
        If np_ - nmax + 1 > imin Then
            imin = np_ - nmax + 1
        End If
    End If
    i0 = np_ - 1
    For i = np_ - 1 To imin Step -1
        s = s + Dist(xs(i), ys(i), xs(i - 1), ys(i - 1))
        If s > nearW * w Then
            Exit For
        End If
        i0 = i - 1
    Next i
    nt = np_ - i0
    ReDim devs(0 To nt)
    For i = i0 To np_ - 1
        x = xs(i): y = ys(i)
        FitProj f, x, y
        devs(i - i0) = Dist(x, y, xs(i), ys(i))
    Next i
    ' крючок на конце
    kk = nt: s = 0
    Do While kk > 1 And devs(kk - 1) > maxdev
        s = s + Dist(xs(i0 + kk - 1), ys(i0 + kk - 1), xs(i0 + kk - 2), ys(i0 + kk - 2))
        If s > 0.5 * w Then
            Exit Do
        End If
        kk = kk - 1
    Loop
    If kk < nt Then
        mxd = 0
        For i = 0 To kk - 1
            If devs(i) > mxd Then
                mxd = devs(i)
            End If
        Next i
        If mxd <= maxdev Then
            np_ = i0 + kk: nt = kk
        End If
    End If
    mxd = 0
    For i = 0 To nt - 1
        If devs(i) > mxd Then
            mxd = devs(i)
        End If
    Next i
    If mxd > maxdev Then
        Exit Function
    End If
    If nt >= 3 And mxd > 0.04 * w Then
        ReDim tx(0 To nt): ReDim ty(0 To nt)
        For i = 0 To nt - 1
            tx(i) = xs(i0 + i): ty(i) = ys(i0 + i)
        Next i
        FitLineLS tx, ty, nt, mx, my, dx, dy, rms
        If rms < 0.02 * w Then
            px0 = tx(0): py0 = ty(0)
            FitProj f, px0, py0
            FitTangent f, px0, py0, tx(nt - 1) - tx(0), ty(nt - 1) - ty(0), tgx, tgy
            If Abs(tgx * dx + tgy * dy) < Cos(6 * PI_ / 180) Then
                Exit Function
            End If
        End If
    End If
    ReDim ox(0 To nt): ReDim oy(0 To nt)
    For i = i0 To np_ - 1
        x = xs(i): y = ys(i)
        FitProj f, x, y
        If i = i0 Then
            ref = FitParam(f, x, y): tmax = ref
            ox(0) = x: oy(0) = y: no = 1
        Else
            t = FitParam(f, x, y)
            If f(0) = 1 Then
                per = 2 * PI_ * f(3)
                d = t - ref + PI_ * f(3)
                d = d - Int(d / per) * per
                t = ref + d - PI_ * f(3)
            End If
            If t > tmax + 0.000000001 Then
                ox(no) = x: oy(no) = y: no = no + 1: tmax = t
            End If
        End If
    Next i
    For k = 0 To no - 1
        xs(i0 + k) = ox(k): ys(i0 + k) = oy(k)
    Next k
    np_ = i0 + no
    ApplyFar = True
End Function

Private Function FitEnd(xs() As Double, ys() As Double, ByVal np_ As Long, ByVal w As Double, ByVal nmax As Long, _
                        f() As Double) As Boolean
    Dim ax0() As Double, ay0() As Double, ad() As Double, na As Long, i As Long, imin As Long, s As Double
    Dim wins(0 To 2) As Double, wi As Long, n As Long, ok As Boolean, cx As Double, cy As Double, rr As Double, rms As Double
    Dim mx As Double, my As Double, dx As Double, dy As Double, k As Long, refX As Double, refY As Double
    Dim aEnd As Double, aRef As Double, dlt As Double, res As Double, x As Double, y As Double, d As Double
    Dim ex As Double, ey As Double
    f(0) = 0
    ReDim ax0(0 To np_): ReDim ay0(0 To np_): ReDim ad(0 To np_)
    imin = 0
    If nmax >= 0 Then
        If np_ - nmax > imin Then
            imin = np_ - nmax
        End If
    End If
    For i = np_ - 1 To imin Step -1
        ax0(na) = xs(i): ay0(na) = ys(i): ad(na) = s: na = na + 1
        If i > 0 Then
            s = s + Dist(xs(i), ys(i), xs(i - 1), ys(i - 1))
        End If
        If s > 2 * w Then
            Exit For
        End If
    Next i
    ex = xs(np_ - 1): ey = ys(np_ - 1)
    wins(0) = 2: wins(1) = 1.2: wins(2) = 0.7
    For wi = 0 To 2
        n = 0
        For i = 0 To na - 1
            If ad(i) <= wins(wi) * w Then
                n = n + 1
            End If
        Next i
        If n >= 3 Then
            ok = False
            If n >= 4 And wi < 2 Then          ' на самом коротком окне окружности не верим
                If FitCircle(ax0, ay0, n, cx, cy, rr, rms) Then
                    If rr < 25 * w And rms < 0.05 * w And Bent(ax0, ay0, 0, n - 1, rr, w) Then
                        ' знак хода: от точки на 0.5w от конца к концу
                        s = 0: refX = ax0(n - 1): refY = ay0(n - 1)
                        For k = 1 To n - 1
                            s = s + Dist(ax0(k - 1), ay0(k - 1), ax0(k), ay0(k))
                            If s >= 0.5 * w Then
                                refX = ax0(k): refY = ay0(k)
                                Exit For
                            End If
                        Next k
                        aEnd = Atan2(ey - cy, ex - cx): aRef = Atan2(refY - cy, refX - cx)
                        dlt = aEnd - aRef
                        Do While dlt > PI_
                            dlt = dlt - 2 * PI_
                        Loop
                        Do While dlt <= -PI_
                            dlt = dlt + 2 * PI_
                        Loop
                        f(0) = 1: f(1) = cx: f(2) = cy: f(3) = rr
                        If dlt >= 0 Then
                            f(4) = 1
                        Else
                            f(4) = -1
                        End If
                        ok = True
                    End If
                End If
            End If
            If Not ok Then
                FitLineLS ax0, ay0, n, mx, my, dx, dy, rms
                If rms < 0.08 * w Then
                    If (ex - ax0(n - 1)) * dx + (ey - ay0(n - 1)) * dy < 0 Then
                        dx = -dx: dy = -dy
                    End If
                    f(0) = 2: f(1) = mx: f(2) = my: f(3) = dx: f(4) = dy
                    ok = True
                End If
            End If
            If ok Then
                res = 0
                For i = 0 To na - 1
                    If ad(i) <= 0.7 * w Then
                        x = ax0(i): y = ay0(i)
                        FitProj f, x, y
                        d = Dist(x, y, ax0(i), ay0(i))
                        If d > res Then
                            res = d
                        End If
                    End If
                Next i
                If res < 0.02 * w Or wi = 2 Then
                    FitEnd = True
                    Exit Function
                End If
            End If
        End If
    Next wi
    f(0) = 0
End Function


Private Sub EndProject(xs() As Double, ys() As Double, ByVal np_ As Long, f() As Double, ByVal w As Double, ByVal nmax As Long, _
                       Optional ByVal L As Double = -1)
    Dim i As Long, imin As Long, s As Double
    If L < 0 Then
        L = w
    End If
    imin = 0
    If nmax >= 0 Then
        If np_ - nmax > imin Then
            imin = np_ - nmax
        End If
    End If
    For i = np_ - 1 To imin Step -1
        If i > 0 Then
            s = s + Dist(xs(i), ys(i), xs(i - 1), ys(i - 1))
        End If
        FitProj f, xs(i), ys(i)
        If s > L Then
            Exit For
        End If
    Next i
End Sub

Private Sub Retract(xs() As Double, ys() As Double, np_ As Long, ByVal back As Double)
    Dim s As Double, t As Double
    Do While np_ > 2
        s = Dist(xs(np_ - 1), ys(np_ - 1), xs(np_ - 2), ys(np_ - 2))
        If s > back Then
            Exit Do
        End If
        back = back - s
        np_ = np_ - 1
    Loop
    If back > 0 And np_ >= 2 Then
        s = Dist(xs(np_ - 1), ys(np_ - 1), xs(np_ - 2), ys(np_ - 2))
        If s > 0.000000001 Then
            t = back / s
            xs(np_ - 1) = xs(np_ - 1) + (xs(np_ - 2) - xs(np_ - 1)) * t
            ys(np_ - 1) = ys(np_ - 1) + (ys(np_ - 2) - ys(np_ - 1)) * t
        End If
    End If
End Sub


Private Function SegHit(px_() As Double, py_() As Double, ByVal n As Long, ByVal ax0 As Double, ByVal ay0 As Double, _
                        ByVal bx0 As Double, ByVal by0 As Double, hx As Double, hy As Double) As Boolean
    Dim i As Long, ex As Double, ey As Double, qx As Double, qy As Double, den As Double, t As Double, u As Double, bt As Double
    bt = 2
    ex = bx0 - ax0: ey = by0 - ay0
    For i = 0 To n - 2
        qx = px_(i + 1) - px_(i): qy = py_(i + 1) - py_(i)
        den = ex * qy - ey * qx
        If Abs(den) > 1E-12 Then
            t = ((px_(i) - ax0) * qy - (py_(i) - ay0) * qx) / den
            u = ((px_(i) - ax0) * ey - (py_(i) - ay0) * ex) / den
            If t >= -0.000000001 And t <= 1.000000001 And u >= -0.000000001 And u <= 1.000000001 And t < bt Then
                bt = t: hx = ax0 + ex * t: hy = ay0 + ey * t
                SegHit = True
            End If
        End If
    Next i
End Function


Private Function PolyClosest(px_() As Double, py_() As Double, ByVal n As Long, ByVal x As Double, ByVal y As Double, _
                             cx As Double, cy As Double) As Double
    Dim i As Long, ex As Double, ey As Double, L2 As Double, t As Double, qx As Double, qy As Double, d As Double
    PolyClosest = BIG
    For i = 0 To n - 2
        ex = px_(i + 1) - px_(i): ey = py_(i + 1) - py_(i): L2 = ex * ex + ey * ey
        If L2 < 1E-18 Then
            t = 0
        Else
            t = ((x - px_(i)) * ex + (y - py_(i)) * ey) / L2
            If t < 0 Then
                t = 0
            ElseIf t > 1 Then
                t = 1
            End If
        End If
        qx = px_(i) + ex * t: qy = py_(i) + ey * t
        d = Dist(x, y, qx, qy)
        If d < PolyClosest Then
            PolyClosest = d: cx = qx: cy = qy
        End If
    Next i
End Function


Private Function StubExtend(xs() As Double, ys() As Double, np_ As Long, ByVal w As Double, px_() As Double, py_() As Double, _
                            ByVal npl As Long, ByVal v0 As Long, ByVal v1 As Long, ByVal jwJ As Double, ByVal jx As Double, _
                            ByVal jy As Double, ByVal nmax As Long) As Boolean
    Dim f(0 To 5) As Double, okF As Boolean, sx As Double, sy As Double, stp As Double, maxlen As Double, L As Double
    Dim pvx As Double, pvy As Double, ax0() As Double, ay0() As Double, na As Long, k As Long, x As Double, y As Double
    Dim hx As Double, hy As Double, bestD As Double, bestI As Long, bcx As Double, bcy As Double, cx As Double, cy As Double, d As Double
    Dim i As Long, rx() As Double, ry() As Double, nr_ As Long, lastD As Double
    Dim maxhit As Double, hitOk As Boolean, touch As Boolean
    If npl < 2 Then
        Exit Function
    End If
    If NearLine(xs, ys, np_, w, nmax, f, L) Then
        EndProject xs, ys, np_, f, w, nmax, L   ' только прямой участок, дальше дуга остаётся дугой
        BlendArc xs, ys, np_, f, L, w, nmax
    Else
        okF = FarFit(xs, ys, np_, w, 1.5, 3.5, nmax, f)
        If okF Then
            okF = ApplyFar(xs, ys, np_, w, f, 1.5, 0.12 * w, nmax)
        End If
        If Not okF Then
            okF = FitEnd(xs, ys, np_, w, nmax, f)
            If Not okF Then
                Exit Function
            End If
            EndProject xs, ys, np_, f, w, nmax
        End If
    End If
    ' реальная (без виртуальных концов) ломаная для касания
    nr_ = npl - v0 - v1
    ReDim rx(0 To npl): ReDim ry(0 To npl)
    For i = 0 To nr_ - 1
        rx(i) = px_(v0 + i): ry(i) = py_(v0 + i)
    Next i
    sx = xs(np_ - 1): sy = ys(np_ - 1): stp = gH / 2
    maxhit = Dist(sx, sy, jx, jy) * 1.5 + jwJ     ' пересечение — только в этих пределах
    maxlen = maxhit
    If f(0) = 2 Then
        maxlen = maxhit + jwJ   ' касание прямой ножки вдоль дуги (g: ножка в кольцо) — чуть дальше
    End If
    pvx = sx: pvy = sy
    ReDim ax0(0 To 63): ReDim ay0(0 To 63)
    bestD = BIG: bestI = -1
    k = 1
    Do While k * stp <= maxlen
        FitAt f, sx, sy, k * stp, x, y
        hitOk = False
        If k * stp <= maxhit Then
            hitOk = SegHit(px_, py_, npl, pvx, pvy, x, y, hx, hy)
        End If
        If hitOk Then
            If InUnion(hx, hy, -1) Then
                For i = 0 To na - 1
                    PushPt xs, ys, np_, ax0(i), ay0(i)
                Next i
                PushPt xs, ys, np_, hx, hy
                StubExtend = True
                Exit Function
            End If
        End If
        If nr_ >= 2 Then
            d = PolyClosest(rx, ry, nr_, x, y, cx, cy)
            If bestI >= 0 And bestD < 0.15 * jwJ And d > bestD + 0.02 * jwJ Then
                ' касание: расстояние до сквозной прошло через минимум и снова растёт
                ' (дальше кривая идёт вдоль сквозной параллельно — петля «6», «9»)
                For i = 0 To bestI
                    PushPt xs, ys, np_, ax0(i), ay0(i)
                Next i
                PushPt xs, ys, np_, bcx, bcy
                LandOffset xs, ys, np_, w
                StubExtend = True
                Exit Function
            End If
            If d < bestD Then
                bestD = d: bestI = na: bcx = cx: bcy = cy
            End If
            lastD = d
        End If
        PushPt ax0, ay0, na, x, y
        pvx = x: pvy = y: k = k + 1
    Loop
    touch = False
    If bestI >= 0 And bestI < na - 1 And lastD > bestD + 0.01 * w Then
        If bestD < 0.15 * jwJ Then
            touch = True
        ElseIf bestD < 0.3 * jwJ Then
            touch = ParallelAt(ax0, ay0, na, bestI, rx, ry, nr_, bcx, bcy)
        End If
    End If
    If touch Then
        ' касание большой дугой (арка n, m приходит в ножку по касательной): минимум расстояния
        ' внутри участка, дальше оно растёт. Дуга шрифта не точно окружность — зазор до 0.3jw
        ' допустим, если в точке подхода кривая почти параллельна сквозной (зазор разводится плавно)
        For i = 0 To bestI
            PushPt xs, ys, np_, ax0(i), ay0(i)
        Next i
        PushPt xs, ys, np_, bcx, bcy
        LandOffset xs, ys, np_, w
        StubExtend = True
        Exit Function
    End If
    ' ни пересечения, ни касания. Сквозная могла поворачивать в развилке (база Д уходит в
    ' диагональ, а её кусок за развилкой отрезан как усик): продолжения сторон сквозной за
    ' её вершины у развилки (< 1jw от центра) — луч заглушки (прямая по участку 0.5..3w от
    ' конца) попадает в продолжение: конец в точке попадания и уголком к сквозной
    StubExtend = StubToExt(xs, ys, np_, w, rx, ry, nr_, jwJ, jx, jy, stp, maxhit)
End Function

Private Function FitPts(x() As Double, y() As Double, ByVal n As Long, ByVal w As Double, ByVal kc As Double, _
                        ByVal kl As Double, f() As Double) As Boolean
    Dim mx As Double, my As Double, dx As Double, dy As Double, rms As Double, cx As Double, cy As Double, rr As Double
    f(0) = 0
    If n < 3 Then
        Exit Function
    End If
    FitLineLS x, y, n, mx, my, dx, dy, rms
    If rms < kl * w Then
        f(0) = 2: f(1) = mx: f(2) = my: f(3) = dx: f(4) = dy
        FitPts = True
        Exit Function
    End If
    If n >= 4 Then
        If FitCircle(x, y, n, cx, cy, rr, rms) Then
            If rr < 25 * w And rms < kc * w And Bent(x, y, 0, n - 1, rr, w) Then
                f(0) = 1: f(1) = cx: f(2) = cy: f(3) = rr: f(4) = 0
                FitPts = True
            End If
        End If
    End If
End Function


Private Function FitIntersect(f1() As Double, f2() As Double, ByVal nx0 As Double, ByVal ny0 As Double, _
                              xX As Double, yX As Double) As Boolean
    Dim cx() As Double, cy() As Double, nc As Long, k As Long, bestD As Double, d As Double
    Dim ax0 As Double, ay0 As Double, ra As Double, bx0 As Double, by0 As Double, rb As Double, dd As Double, a As Double, hh As Double
    Dim ux As Double, uy As Double, px0 As Double, py0 As Double, mx As Double, my As Double, dx As Double, dy As Double
    Dim t0 As Double, fxx As Double, fyy As Double, swapF As Boolean
    ReDim cx(0 To 1): ReDim cy(0 To 1)
    If f1(0) = 2 And f2(0) = 2 Then
        If LineX(f1(1), f1(2), f1(3), f1(4), f2(1), f2(2), f2(3), f2(4), cx(0), cy(0)) Then
            nc = 1
        End If
    ElseIf f1(0) = 1 And f2(0) = 1 Then
        ax0 = f1(1): ay0 = f1(2): ra = f1(3): bx0 = f2(1): by0 = f2(2): rb = f2(3)
        dd = Dist(ax0, ay0, bx0, by0)
        If dd > 0.000000001 And Abs(ra - rb) <= dd And dd <= ra + rb Then
            a = (ra * ra - rb * rb + dd * dd) / (2 * dd)
            hh = ra * ra - a * a
            If hh < 0 Then
                hh = 0
            End If
            hh = Sqr(hh)
            ux = (bx0 - ax0) / dd: uy = (by0 - ay0) / dd
            px0 = ax0 + ux * a: py0 = ay0 + uy * a
            cx(0) = px0 - uy * hh: cy(0) = py0 + ux * hh
            cx(1) = px0 + uy * hh: cy(1) = py0 - ux * hh
            nc = 2
        End If
    Else
        If f1(0) = 1 Then
            mx = f2(1): my = f2(2): dx = f2(3): dy = f2(4)
            ax0 = f1(1): ay0 = f1(2): ra = f1(3)
        Else
            mx = f1(1): my = f1(2): dx = f1(3): dy = f1(4)
            ax0 = f2(1): ay0 = f2(2): ra = f2(3)
        End If
        t0 = (ax0 - mx) * dx + (ay0 - my) * dy
        fxx = mx + dx * t0: fyy = my + dy * t0
        dd = Dist(fxx, fyy, ax0, ay0)
        If dd <= ra Then
            hh = ra * ra - dd * dd
            If hh < 0 Then
                hh = 0
            End If
            hh = Sqr(hh)
            cx(0) = fxx + dx * hh: cy(0) = fyy + dy * hh
            cx(1) = fxx - dx * hh: cy(1) = fyy - dy * hh
            nc = 2
        End If
    End If
    If nc = 0 Then
        Exit Function
    End If
    bestD = BIG
    For k = 0 To nc - 1
        d = Dist(cx(k), cy(k), nx0, ny0)
        If d < bestD Then
            bestD = d: xX = cx(k): yX = cy(k)
        End If
    Next k
    FitIntersect = True
End Function


Private Sub FitBetween(f() As Double, ByVal ax0 As Double, ByVal ay0 As Double, ByVal bx0 As Double, ByVal by0 As Double, _
                       ByVal stp As Double, fx() As Double, fy() As Double, nf As Long)
    Dim L As Double, m As Long, k As Long, t As Double, a0 As Double, a1 As Double, d As Double
    If f(0) = 2 Then
        L = Dist(ax0, ay0, bx0, by0): m = Int(L / stp)
        For k = 1 To m - 1
            t = k / m
            PushPt fx, fy, nf, ax0 + (bx0 - ax0) * t, ay0 + (by0 - ay0) * t
        Next k
    Else
        a0 = Atan2(ay0 - f(2), ax0 - f(1)): a1 = Atan2(by0 - f(2), bx0 - f(1))
        d = a1 - a0
        Do While d > PI_
            d = d - 2 * PI_
        Loop
        Do While d <= -PI_
            d = d + 2 * PI_
        Loop
        m = Int(Abs(d) * f(3) / stp)
        For k = 1 To m - 1
            t = a0 + d * k / m
            PushPt fx, fy, nf, f(1) + f(3) * Cos(t), f(2) + f(3) * Sin(t)
        Next k
    End If
End Sub


Private Sub EdgeLines(ByVal c As Long, ByVal w As Double, hf() As Double, nh As Long)
    Dim side As Long, sg As Long, ex As Double, ey As Double, nx0 As Double, ny0 As Double, ox As Double, oy As Double
    Dim D As Double, wl As Double
    For side = 0 To 1
        If side = 0 Then
            ex = BE1X(c): ey = BE1Y(c)
        Else
            ex = BE2X(c): ey = BE2Y(c)
        End If
        For sg = 1 To -1 Step -2
            nx0 = -ey * sg: ny0 = ex * sg
            ox = BCX(c) + ex * 0.2 * w + nx0 * 0.05 * w: oy = BCY(c) + ey * 0.2 * w + ny0 * 0.05 * w
            If InUnion(ox, oy, -1) Then
                D = RayUnion(ox, oy, nx0, ny0)
                If D < BIG Then
                    wl = D + 0.05 * w
                    If wl >= 0.6 * w And wl <= 1.6 * w Then
                        If nh > UBound(hf, 2) Then
                            ReDim Preserve hf(0 To 5, 0 To 2 * nh)
                        End If
                        hf(0, nh) = 2: hf(1, nh) = BCX(c) + nx0 * wl / 2: hf(2, nh) = BCY(c) + ny0 * wl / 2
                        hf(3, nh) = ex: hf(4, nh) = ey: hf(5, nh) = -(c * 2 + side + 1)
                        nh = nh + 1
                        Exit For
                    End If
                End If
            End If
        Next sg
    Next side
End Sub


Private Function EdgeMatch(f() As Double, ByVal c As Long, ByVal w As Double, ByVal ex As Double, ByVal ey As Double) As Boolean
    Dim px0 As Double, py0 As Double, cosmin As Double, a0 As Double, a As Double, sg As Long, qx As Double, qy As Double
    Dim dx As Double, dy As Double
    px0 = BCX(c): py0 = BCY(c)
    FitProj f, px0, py0
    cosmin = Cos(7 * PI_ / 180)
    If f(0) = 2 Then
        EdgeMatch = (Abs(f(3) * ex + f(4) * ey) >= cosmin)
        Exit Function
    End If
    a0 = Atan2(py0 - f(2), px0 - f(1))
    For sg = 1 To -1 Step -2
        a = a0 + sg * 0.5 * w / f(3)
        qx = f(1) + f(3) * Cos(a): qy = f(2) + f(3) * Sin(a)
        dx = qx - px0: dy = qy - py0: Norm dx, dy
        If dx * ex + dy * ey >= cosmin Then
            EdgeMatch = True
            Exit Function
        End If
    Next sg
End Function


Private Sub SnapEdge(f() As Double, x() As Double, y() As Double, ByVal n As Long, ByVal w As Double, ByVal span As Double)
    Dim c As Long, hf() As Double, nh As Long, k As Long, i As Long, d As Double, px0 As Double, py0 As Double
    Dim bestD As Double, bestK As Long, cosmin As Double
    If f(0) <> 2 Or span >= 1.2 * w Then
        Exit Sub
    End If
    cosmin = Cos(12 * PI_ / 180)
    ReDim hf(0 To 5, 0 To 7)
    For c = 0 To NBC - 1
        If Not BCC(c) Then
            If Dist(BCX(c), BCY(c), x(0), y(0)) <= 2 * w Then
                EdgeLines c, w, hf, nh
            End If
        End If
    Next c
    bestD = BIG: bestK = -1
    For k = 0 To nh - 1
        If Abs(hf(3, k) * f(3) + hf(4, k) * f(4)) >= cosmin Then
            d = 0
            For i = 0 To n - 1
                px0 = x(i): py0 = y(i)
                px0 = hf(1, k) + hf(3, k) * ((x(i) - hf(1, k)) * hf(3, k) + (y(i) - hf(2, k)) * hf(4, k))
                py0 = hf(2, k) + hf(4, k) * ((x(i) - hf(1, k)) * hf(3, k) + (y(i) - hf(2, k)) * hf(4, k))
                d = d + Dist(px0, py0, x(i), y(i))
            Next i
            d = d / n
            If d < 0.25 * w And d < bestD Then
                bestD = d: bestK = k
            End If
        End If
    Next k
    If bestK >= 0 Then
        f(1) = hf(1, bestK): f(2) = hf(2, bestK): f(3) = hf(3, bestK): f(4) = hf(4, bestK)
    End If
End Sub


Private Function PolyDist(rx() As Double, ry() As Double, ByVal nr_ As Long, px_() As Double, py_() As Double, ByVal n As Long) As Double
    Dim i As Long, tot As Double, cx As Double, cy As Double
    If nr_ = 0 Or n < 2 Then
        PolyDist = BIG
        Exit Function
    End If
    For i = 0 To nr_ - 1
        tot = tot + PolyClosest(px_, py_, n, rx(i), ry(i), cx, cy)
    Next i
    PolyDist = tot / nr_
End Function


Private Sub TakePts(src_x() As Double, src_y() As Double, ByVal n As Long, ByVal fromEnd As Boolean, ByVal L As Double, _
                    ox() As Double, oy() As Double, no As Long)
    Dim i As Long, s As Double, stp As Long
    no = 0
    ReDim ox(0 To n): ReDim oy(0 To n)
    If fromEnd Then
        i = n - 1: stp = -1
    Else
        i = 0: stp = 1
    End If
    ox(0) = src_x(i): oy(0) = src_y(i): no = 1
    Do
        If i + stp < 0 Or i + stp >= n Then
            Exit Do
        End If
        s = s + Dist(src_x(i), src_y(i), src_x(i + stp), src_y(i + stp))
        If s > L Then
            Exit Do
        End If
        i = i + stp
        ox(no) = src_x(i): oy(no) = src_y(i): no = no + 1
    Loop
End Sub


Private Function ArcFillSingle(xs() As Double, ys() As Double, ByVal np_ As Long, ax_() As Double, ay_() As Double, _
                               ByVal na As Long, ByVal stp As Double, ByVal w As Double, fx() As Double, fy() As Double, _
                               nf As Long, pf() As Double) As Boolean
    Dim wins(0 To 2) As Double, wi As Long, qx() As Double, qy() As Double, nq As Long
    Dim bx() As Double, by() As Double, nb As Long, cx_() As Double, cy_() As Double, nc As Long, i As Long, s As Double
    Dim cx As Double, cy As Double, rr As Double, rms As Double, mx As Double, my As Double, ddx As Double, ddy As Double
    Dim a0 As Double, a1 As Double, b0 As Double, b1 As Double, m As Long, k As Long, x As Double, y As Double
    wins(0) = 2.5: wins(1) = 1.5: wins(2) = 0.8
    pf(0) = 0
    For wi = 0 To 2
        TakePts xs, ys, np_, True, wins(wi) * w, bx, by, nb
        TakePts ax_, ay_, na, False, wins(wi) * w, cx_, cy_, nc
        If nb >= 3 And nc >= 3 Then
            ReDim qx(0 To nb + nc): ReDim qy(0 To nb + nc): nq = 0
            For i = 0 To nb - 1
                qx(nq) = bx(i): qy(nq) = by(i): nq = nq + 1
            Next i
            For i = 0 To nc - 1
                qx(nq) = cx_(i): qy(nq) = cy_(i): nq = nq + 1
            Next i
            If FitCircle(qx, qy, nq, cx, cy, rr, rms) Then
                If rr < 25 * w And rms < 0.02 * w And Bent(qx, qy, 0, nq - 1, rr, w) Then
                    pf(0) = 1: pf(1) = cx: pf(2) = cy: pf(3) = rr: pf(4) = 0
                    Exit For
                End If
            End If
            FitLineLS qx, qy, nq, mx, my, ddx, ddy, rms
            If rms < 0.02 * w Then
                pf(0) = 2: pf(1) = mx: pf(2) = my: pf(3) = ddx: pf(4) = ddy
                Exit For
            End If
        End If
    Next wi
    If pf(0) = 0 Then
        Exit Function
    End If
    ' подправить хвост "до"
    i = np_ - 1: s = 0
    Do While i >= 0
        FitProj pf, xs(i), ys(i)
        If i > 0 Then
            s = s + Dist(xs(i), ys(i), xs(i - 1), ys(i - 1))
        End If
        If s > 0.8 * w Then
            Exit Do
        End If
        i = i - 1
    Loop
    a0 = xs(np_ - 1): a1 = ys(np_ - 1)
    b0 = ax_(0): b1 = ay_(0)
    FitProj pf, b0, b1
    m = Int(Dist(a0, a1, b0, b1) / stp)
    If m < 1 Then
        m = 1
    End If
    nf = 0
    For k = 1 To m - 1
        x = a0 + (b0 - a0) * k / m: y = a1 + (b1 - a1) * k / m
        FitProj pf, x, y
        PushPt fx, fy, nf, x, y
    Next k
    ArcFillSingle = True
End Function


Private Function CornerFill(xs() As Double, ys() As Double, ByVal np_ As Long, ax_() As Double, ay_() As Double, _
                            ByVal na As Long, ByVal stp As Double, ByVal w As Double, fx() As Double, fy() As Double, _
                            nf As Long, pf() As Double, xi As Long) As Boolean
    Dim bx() As Double, by() As Double, nb As Long, cx_() As Double, cy_() As Double, nc As Long
    Dim fa(0 To 5) As Double, fb(0 To 5) As Double, a0 As Double, a1 As Double, b0 As Double, b1 As Double
    Dim xX As Double, yX As Double, xa As Double, ya As Double, xb As Double, yb As Double, gap As Double
    Dim dax As Double, day_ As Double, dbx As Double, dby As Double, tax As Double, tay As Double, tbx As Double, tby As Double
    Dim i As Long, s As Double, cosLim As Double, k As Long, corner As Boolean, ka As Double, kb As Double, sa As Double, sb As Double
    Dim kmax As Double, kmin As Double
    xi = -1
    TakePts xs, ys, np_, True, 2 * w, bx, by, nb
    TakePts ax_, ay_, na, False, 2 * w, cx_, cy_, nc
    If Not FitPts(bx, by, nb, w, 0.02, 0.02, fa) Then
        Exit Function
    End If
    If Not FitPts(cx_, cy_, nc, w, 0.02, 0.02, fb) Then
        Exit Function
    End If
    SnapEdge fa, bx, by, nb, w, Dist(bx(0), by(0), bx(nb - 1), by(nb - 1))
    SnapEdge fb, cx_, cy_, nc, w, Dist(cx_(0), cy_(0), cx_(nc - 1), cy_(nc - 1))
    a0 = xs(np_ - 1): a1 = ys(np_ - 1): b0 = ax_(0): b1 = ay_(0)
    If Not FitIntersect(fa, fb, (a0 + b0) / 2, (a1 + b1) / 2, xX, yX) Then
        If Not FitTouch(fa, fb, w, xX, yX) Then   ' касание прямой и дуги (перекладина B -> чаша)
            Exit Function
        End If
    End If
    If Not InUnion(xX, yX, -1) Then
        Exit Function
    End If
    xa = xX: ya = yX: FitProj fa, xa, ya
    xb = xX: yb = yX: FitProj fb, xb, yb
    gap = Dist(a0, a1, b0, b1)
    If Dist(a0, a1, xa, ya) + Dist(xb, yb, b0, b1) > 2 * gap + w Then
        Exit Function
    End If
    If nb < 2 Or nc < 2 Then
        Exit Function
    End If
    ' X впереди по ходу от a и позади b
    dax = xa - bx(1): day_ = ya - by(1): Norm dax, day_
    dbx = cx_(1) - xb: dby = cy_(1) - yb: Norm dbx, dby
    tax = a0 - bx(1): tay = a1 - by(1): Norm tax, tay
    tbx = cx_(1) - b0: tby = cy_(1) - b1: Norm tbx, tby
    If dax * tax + day_ * tay < 0.7 Or dbx * tbx + dby * tby < 0.7 Then
        Exit Function
    End If
    ' поворот в X заметный — угол; иначе — составная кривая, но только если стороны разной
    ' природы — две почти соосные прямые или две похожие дуги лучше одним фитом (без излома)
    cosLim = Cos(25 * PI_ / 180)
    corner = (dax * dbx + day_ * dby <= cosLim)
    If Not corner Then
        If fa(0) = 2 And fb(0) = 2 Then
            Exit Function
        End If
        If fa(0) = 1 And fb(0) = 1 Then
            ka = 1 / fa(3): kb = 1 / fb(3)
            ' по разные стороны (S) или заметно разная кривизна
            sa = (fa(1) - xX) * (-day_) + (fa(2) - yX) * dax
            sb = (fb(1) - xX) * (-dby) + (fb(2) - yX) * dbx
            kmax = ka: kmin = kb
            If kb > kmax Then
                kmax = kb: kmin = ka
            End If
            If sa * sb > 0 And kmax < 1.5 * kmin Then
                Exit Function
            End If
        End If
    End If
    ' хвост "до" — на кривую своей стороны
    i = np_ - 1: s = 0
    Do While i >= 0
        FitProj fa, xs(i), ys(i)
        If i > 0 Then
            s = s + Dist(xs(i), ys(i), xs(i - 1), ys(i - 1))
        End If
        If s > 0.8 * w Then
            Exit Do
        End If
        i = i - 1
    Loop
    a0 = xs(np_ - 1): a1 = ys(np_ - 1)
    nf = 0
    FitBetween fa, a0, a1, xa, ya, stp, fx, fy, nf
    If corner Then
        xi = nf
    Else
        xi = -1
    End If
    PushPt fx, fy, nf, xX, yX
    b0 = ax_(0): b1 = ay_(0): FitProj fb, b0, b1
    FitBetween fb, xb, yb, b0, b1, stp, fx, fy, nf
    For k = 0 To 5
        pf(k) = fb(k)
    Next k
    CornerFill = True
End Function

Private Function EndDir(sx() As Double, sy() As Double, ByVal n As Long, ByVal fromEnd As Boolean, ByVal win As Double, _
                        dx As Double, dy As Double) As Boolean
    Dim px_() As Double, py_() As Double, np_ As Long, i As Long, stp As Long, x0 As Double, y0 As Double
    Dim mx As Double, my As Double, rms As Double
    ReDim px_(0 To n): ReDim py_(0 To n)
    If fromEnd Then
        i = n - 1: stp = -1
    Else
        i = 0: stp = 1
    End If
    x0 = sx(i): y0 = sy(i)
    px_(0) = x0: py_(0) = y0: np_ = 1
    Do
        i = i + stp
        If i < 0 Or i >= n Then
            Exit Do
        End If
        If Dist(sx(i), sy(i), x0, y0) > win Then
            Exit Do
        End If
        px_(np_) = sx(i): py_(np_) = sy(i): np_ = np_ + 1
    Loop
    If np_ < 2 Then
        Exit Function
    End If
    If Dist(px_(0), py_(0), px_(np_ - 1), py_(np_ - 1)) < 0.000000001 Then
        Exit Function
    End If
    FitLineLS px_, py_, np_, mx, my, dx, dy, rms
    If (px_(0) - px_(np_ - 1)) * dx + (py_(0) - py_(np_ - 1)) * dy < 0 Then
        dx = -dx: dy = -dy
    End If
    EndDir = True
End Function


Private Function ArcIdxF(xs() As Double, ys() As Double, fixedF() As Boolean, ByVal n As Long, ByVal closed As Boolean, _
                         ByVal i As Long, ByVal L As Double, ByVal dirn As Long) As Long
    Dim s As Double, j As Long, k As Long, cnt As Long
    j = i
    ArcIdxF = -1
    Do
        k = j + dirn
        If Not closed Then
            If k < 0 Or k >= n Then
                Exit Function
            End If
        Else
            k = (k + n) Mod n
        End If
        s = s + Dist(xs(j), ys(j), xs(k), ys(k)): j = k
        If fixedF(j) Then
            Exit Function
        End If
        If s >= L Then
            ArcIdxF = j
            Exit Function
        End If
        cnt = cnt + 1
        If j = i Or cnt > n Then
            Exit Function
        End If
    Loop
End Function


Private Function SideFit(xs() As Double, ys() As Double, ByVal n As Long, ByVal closed As Boolean, ByVal i As Long, _
                         ByVal w As Double, ByVal dirn As Long, hf() As Double, nh As Long, kIn As Long) As Boolean
    Dim s As Double, j As Long, k As Long, cnt As Long, px_() As Double, py_() As Double, ps() As Double, np_ As Long
    Dim strict As Long, wi As Long, kc As Double, kl As Double, win As Double, qx() As Double, qy() As Double, nq As Long
    Dim fc(0 To 5) As Double, fl(0 To 5) As Double, hasC As Boolean, hasL As Boolean, f(0 To 5) As Double, hasF As Boolean
    Dim cx As Double, cy As Double, rr As Double, rms As Double, mx As Double, my As Double, dx As Double, dy As Double
    Dim dup As Boolean, m As Long
    kIn = -1: nh = 0
    ReDim px_(0 To n): ReDim py_(0 To n): ReDim ps(0 To n)
    j = i
    Do
        k = j + dirn
        If Not closed Then
            If k < 0 Or k >= n Then
                Exit Do
            End If
        Else
            k = (k + n) Mod n
        End If
        s = s + Dist(xs(j), ys(j), xs(k), ys(k)): j = k: cnt = cnt + 1
        If j = i Or cnt > n Then
            Exit Do
        End If
        If s >= 0.9 * w Then
            If kIn < 0 Then
                kIn = j
            End If
            px_(np_) = xs(j): py_(np_) = ys(j): ps(np_) = s: np_ = np_ + 1
        End If
        If s >= 2.4 * w Then
            Exit Do
        End If
    Loop
    If kIn < 0 Or np_ < 2 Then
        Exit Function
    End If
    ReDim qx(0 To n): ReDim qy(0 To n)
    For strict = 0 To 1
        If strict = 0 Then
            kc = 0.02: kl = 0.02
        Else
            kc = 0.05: kl = 0.08
        End If
        For wi = 0 To 1
            If wi = 0 Then
                win = 2.4
            Else
                win = 1.6
            End If
            nq = 0
            For k = 0 To np_ - 1
                If ps(k) <= win * w Then
                    qx(nq) = px_(k): qy(nq) = py_(k): nq = nq + 1
                End If
            Next k
            If nq >= 2 Then
                hasC = False: hasL = False
                If nq >= 4 Then
                    If FitCircle(qx, qy, nq, cx, cy, rr, rms) Then
                        If rr < 25 * w And rms < kc * w And Bent(qx, qy, 0, nq - 1, rr, w) Then
                            fc(0) = 1: fc(1) = cx: fc(2) = cy: fc(3) = rr: fc(4) = 0: hasC = True
                        End If
                    End If
                End If
                FitLineLS qx, qy, nq, mx, my, dx, dy, rms
                If rms < kl * w Then
                    fl(0) = 2: fl(1) = mx: fl(2) = my: fl(3) = dx: fl(4) = dy: hasL = True
                End If
                hasF = False
                If wi = 0 Then
                    If hasC Then
                        For m = 0 To 4: f(m) = fc(m): Next m
                        hasF = True
                    ElseIf hasL Then
                        For m = 0 To 4: f(m) = fl(m): Next m
                        hasF = True
                    End If
                Else
                    If hasL Then
                        For m = 0 To 4: f(m) = fl(m): Next m
                        hasF = True
                    ElseIf hasC Then
                        For m = 0 To 4: f(m) = fc(m): Next m
                        hasF = True
                    End If
                End If
                If hasF Then
                    dup = False
                    For k = 0 To nh - 1
                        If hf(0, k) = f(0) And Abs(hf(1, k) - f(1)) < 0.000001 And Abs(hf(2, k) - f(2)) < 0.000001 _
                           And Abs(hf(3, k) - f(3)) < 0.000001 And Abs(hf(4, k) - f(4)) < 0.000001 Then
                            dup = True
                            Exit For
                        End If
                    Next k
                    If Not dup Then
                        If nh > UBound(hf, 2) Then
                            ReDim Preserve hf(0 To 5, 0 To 2 * nh)
                        End If
                        For m = 0 To 4: hf(m, nh) = f(m): Next m
                        hf(5, nh) = 1000 + nh + 100 * (dirn + 1)
                        nh = nh + 1
                    End If
                End If
            End If
        Next wi
    Next strict
    SideFit = (nh > 0)
End Function


Private Function FitDist(f() As Double, ByVal x As Double, ByVal y As Double) As Double
    Dim px0 As Double, py0 As Double
    px0 = x: py0 = y
    FitProj f, px0, py0
    FitDist = Dist(px0, py0, x, y)
End Function


Private Sub GetFit(hf() As Double, ByVal k As Long, f() As Double)
    Dim m As Long
    For m = 0 To 5: f(m) = hf(m, k): Next m
End Sub


Private Function RunLen(ids() As Long, hard() As Boolean, ByVal n As Long, ByVal closed As Boolean, ByVal k As Long, ByVal dirn As Long) As Double
    Dim L As Double, q As Long, q2 As Long, cnt As Long
    q = k
    Do
        q2 = q + dirn
        If closed Then
            q2 = (q2 + n) Mod n
        ElseIf q2 < 0 Or q2 >= n Then
            Exit Do
        End If
        L = L + Dist(GX(ids(q)), GY(ids(q)), GX(ids(q2)), GY(ids(q2))): q = q2: cnt = cnt + 1
        If hard(q) Or cnt >= n Then
            Exit Do
        End If
    Loop
    RunLen = L
End Function


Private Function JWOf(ByVal j As Long, ByVal cref As Double) As Double
    If GDeg(j) >= 3 Or CuspOn(j) Then
        If GJW(j) > 0 Then
            JWOf = GJW(j)
            Exit Function
        End If
    End If
    JWOf = cref
End Function


Private Sub StripHook(xs() As Double, ys() As Double, np_ As Long, ByVal w As Double, ByVal nmax As Long)
    Dim px_() As Double, py_() As Double, n As Long, i As Long, imin As Long, s As Double, k As Long
    Dim mx As Double, my As Double, dx As Double, dy As Double, rms As Double
    ReDim px_(0 To np_): ReDim py_(0 To np_)
    imin = 1
    If nmax >= 0 Then
        If np_ - nmax + 1 > imin Then
            imin = np_ - nmax + 1
        End If
    End If
    For i = np_ - 1 To imin Step -1
        s = s + Dist(xs(i), ys(i), xs(i - 1), ys(i - 1))
        If s > 1.5 * w Then
            Exit For
        End If
        If s >= 0.5 * w Then
            px_(n) = xs(i - 1): py_(n) = ys(i - 1): n = n + 1
        End If
    Next i
    If n < 3 Then
        Exit Sub
    End If
    FitLineLS px_, py_, n, mx, my, dx, dy, rms
    If rms > 0.03 * w Then
        Exit Sub
    End If
    s = 0: k = np_
    Do While k > 2
        If Abs((xs(k - 1) - mx) * dy - (ys(k - 1) - my) * dx) <= 0.04 * w Then
            Exit Do
        End If
        s = s + Dist(xs(k - 1), ys(k - 1), xs(k - 2), ys(k - 2))
        If s > 0.5 * w Then
            Exit Sub
        End If
        k = k - 1
    Loop
    np_ = k
End Sub


Private Function Centered(ByVal x As Double, ByVal y As Double, ByVal dx As Double, ByVal dy As Double, ByVal w As Double) As Boolean
    Dim dl As Double, dr As Double, mn As Double, mx As Double
    dl = RayUnion(x, y, -dy, dx): dr = RayUnion(x, y, dy, -dx)
    If dl >= BIG Or dr >= BIG Then
        Exit Function
    End If
    ' не прижат к стенке; одна сторона может быть «бесконечной» — конец ножки u-образной
    ' буквы (ц, u) упирается в сплошное дно, луч поперёк уходит через всё дно
    mn = dl: mx = dr
    If dr < mn Then
        mn = dr: mx = dl
    End If
    If mn < 0.3 * w Then
        Exit Function
    End If
    Centered = (Abs(dl - dr) <= 0.4 * (dl + dr)) Or (mx > 2 * w)
End Function

Private Sub FreeEndRay(xs() As Double, ys() As Double, np_ As Long, ByVal w As Double)
    Dim dx As Double, dy As Double, a0 As Double, a As Double, k As Long, D As Double, bestD As Double
    Dim bx As Double, by As Double, ext As Double
    TanEnd xs, ys, np_, w, dx, dy
    a0 = Atan2(dy, dx)
    bestD = -1
    For k = -6 To 6
        a = a0 + 5 * k * PI_ / 180
        D = RayUnion(xs(np_ - 1), ys(np_ - 1), Cos(a), Sin(a))
        If D < BIG And D > bestD Then
            bestD = D: bx = Cos(a): by = Sin(a)
        End If
    Next k
    If bestD < 0 Then
        Exit Sub
    End If
    ext = bestD - gTool / 2
    If ext > 0 Then
        PushPt xs, ys, np_, xs(np_ - 1) + bx * ext, ys(np_ - 1) + by * ext
    Else
        Retract xs, ys, np_, -ext
    End If
End Sub


Private Sub FreeEndFit(xs() As Double, ys() As Double, np_ As Long, ByVal cref As Double, ByVal nmax As Long)
    Dim f(0 To 5) As Double, dx As Double, dy As Double, D As Double, ext As Double, s As Double, L As Double
    Dim sx As Double, sy As Double, stp As Double, lExit As Double, smax As Double, k As Long, x As Double, y As Double
    Dim lo As Double, hi As Double, midS As Double, it As Long, target As Double, m As Long, i As Long, hasExit As Boolean
    Dim okF As Boolean
    StripHook xs, ys, np_, cref, nmax
    If NearLine(xs, ys, np_, cref, nmax, f, L) Then
        EndProject xs, ys, np_, f, cref, nmax, L   ' только прямой участок, дальше дуга остаётся дугой
        BlendArc xs, ys, np_, f, L, cref, nmax
        okF = True
    Else
        okF = FarFit(xs, ys, np_, cref, 1#, 3#, nmax, f)
        If okF Then
            okF = ApplyFar(xs, ys, np_, cref, f, 1#, 0.12 * cref, nmax)
        End If
        If Not okF Then
            okF = FitEnd(xs, ys, np_, cref, nmax, f)
            If okF Then
                EndProject xs, ys, np_, f, cref, nmax
            End If
        End If
    End If
    If Not okF Then
        ' запасной вариант: касательная по двум точкам
        TanEnd xs, ys, np_, cref, dx, dy
        D = RayUnion(xs(np_ - 1), ys(np_ - 1), dx, dy)
        If D >= BIG Then
            Exit Sub
        End If
        ext = D - gTool / 2
        If ext > 0 Then
            PushPt xs, ys, np_, xs(np_ - 1) + dx * ext, ys(np_ - 1) + dy * ext
        Else
            Retract xs, ys, np_, -ext
        End If
        Exit Sub
    End If
    sx = xs(np_ - 1): sy = ys(np_ - 1)
    stp = gH / 4: smax = 4.5 * cref: hasExit = False
    k = 1
    Do While k * stp <= smax
        FitAt f, sx, sy, k * stp, x, y
        If Not InUnion(x, y, -1) Then
            lo = (k - 1) * stp: hi = k * stp
            For it = 1 To 10
                midS = (lo + hi) / 2
                FitAt f, sx, sy, midS, x, y
                If InUnion(x, y, -1) Then
                    lo = midS
                Else
                    hi = midS
                End If
            Next it
            lExit = (lo + hi) / 2: hasExit = True
            Exit Do
        End If
        k = k + 1
    Loop
    If Not hasExit Then
        Exit Sub
    End If
    target = lExit - gTool / 2
    If target >= 0 Then
        m = Int(target / gH)
        For i = 1 To m
            FitAt f, sx, sy, i * gH, x, y
            PushPt xs, ys, np_, x, y
        Next i
        If target - m * gH > 0.000001 Or m = 0 Then
            FitAt f, sx, sy, target, x, y
            PushPt xs, ys, np_, x, y
        End If
    Else
        Retract xs, ys, np_, -target
    End If
End Sub

Private Sub StubToJunction(xs() As Double, ys() As Double, np_ As Long, ByVal w As Double, px_() As Double, py_() As Double, _
                           ByVal npl As Long, ByVal v0 As Long, ByVal v1 As Long, ByVal jx As Double, ByVal jy As Double, _
                           ByVal jwJ As Double)
    Dim rx() As Double, ry() As Double, nr_ As Long, i As Long, tgx As Double, tgy As Double, d As Double
    Dim tax As Double, tay As Double, chx As Double, chy As Double, fx() As Double, fy() As Double, nf As Long
    Dim ax0 As Double, ay0 As Double, tdx As Double, tdy As Double, it As Long, i0 As Long, i1 As Long, stp As Long
    Dim s As Double, k As Long, bestD As Double
    nr_ = npl - v0 - v1
    ax0 = xs(np_ - 1): ay0 = ys(np_ - 1)
    TanEnd xs, ys, np_, w, tax, tay
    If nr_ >= 2 Then
        ReDim rx(0 To npl): ReDim ry(0 To npl)
        For i = 0 To nr_ - 1
            rx(i) = px_(v0 + i): ry(i) = py_(v0 + i)
        Next i
        d = PolyClosest(rx, ry, nr_, jx, jy, tgx, tgy)
        bestD = BIG: it = 0
        For i = 0 To nr_ - 1
            d = Dist(rx(i), ry(i), tgx, tgy)
            If d < bestD Then
                bestD = d: it = i
            End If
        Next i
        i1 = it + 1
        If i1 > nr_ - 1 Then
            i1 = nr_ - 1
        End If
        i0 = it - 1
        If i0 < 0 Then
            i0 = 0
        End If
        tdx = rx(i1) - rx(i0): tdy = ry(i1) - ry(i0): Norm tdx, tdy
        If tdx * (tgx - ax0) + tdy * (tgy - ay0) < 0 Then
            stp = -1
        Else
            stp = 1
        End If
        s = 0: k = it
        Do While k + stp >= 0 And k + stp < nr_ And s < 0.5 * jwJ
            s = s + Dist(rx(k), ry(k), rx(k + stp), ry(k + stp)): k = k + stp
        Loop
        tgx = rx(k): tgy = ry(k)
        i1 = k + 1
        If i1 > nr_ - 1 Then
            i1 = nr_ - 1
        End If
        i0 = k - 1
        If i0 < 0 Then
            i0 = 0
        End If
        tdx = rx(i1) - rx(i0): tdy = ry(i1) - ry(i0): Norm tdx, tdy
        If tdx * (tgx - ax0) + tdy * (tgy - ay0) < 0 Then
            tdx = -tdx: tdy = -tdy
        End If
    Else
        tgx = jx: tgy = jy
        tdx = tgx - ax0: tdy = tgy - ay0: Norm tdx, tdy
    End If
    chx = tgx - ax0: chy = tgy - ay0: Norm chx, chy
    If tax * chx + tay * chy < 0.3 Then
        tax = chx: tay = chy
    End If
    ReDim fx(0 To 15): ReDim fy(0 To 15): nf = 0
    HermSeg fx, fy, nf, ax0, ay0, tax, tay, tgx, tgy, tdx, tdy, gH
    PushPt fx, fy, nf, tgx, tgy
    If PathInside(ax0, ay0, fx, fy, nf, 0, 0, False) Then
        For i = 0 To nf - 1
            PushPt xs, ys, np_, fx(i), fy(i)
        Next i
        Exit Sub
    End If
    nf = 0
    HermSeg fx, fy, nf, ax0, ay0, tax, tay, tgx, tgy, chx, chy, gH
    PushPt fx, fy, nf, tgx, tgy
    If PathInside(ax0, ay0, fx, fy, nf, 0, 0, False) Then
        For i = 0 To nf - 1
            PushPt xs, ys, np_, fx(i), fy(i)
        Next i
    Else
        PushPt xs, ys, np_, tgx, tgy
    End If
End Sub

Private Function NearLine(xs() As Double, ys() As Double, ByVal np_ As Long, ByVal w As Double, ByVal nmax As Long, _
                          f() As Double, L As Double) As Boolean
    Dim imin As Long, i As Long, px() As Double, py() As Double, n As Long, s As Double
    Dim mx As Double, my As Double, dx As Double, dy As Double, rms As Double, dev As Double, e As Double, k As Long
    Dim hasBest As Boolean, bx() As Double, by() As Double, nb As Long, bestL As Double
    Dim ix() As Double, iy() As Double, ni As Long, sk As Double, lr As Double, cx As Double, cy As Double, rr As Double
    Dim arc As Boolean, c As Long, hf() As Double, nh As Long, j As Long, side As Long, straight As Double, d As Double
    Dim fmx As Double, fmy As Double, fdx As Double, fdy As Double, cosLim As Double
    imin = 0
    If nmax >= 0 Then
        If np_ - nmax > imin Then
            imin = np_ - nmax
        End If
    End If
    ReDim px(0 To np_): ReDim py(0 To np_): ReDim bx(0 To np_): ReDim by(0 To np_)
    px(0) = xs(np_ - 1): py(0) = ys(np_ - 1): n = 1
    For i = np_ - 2 To imin Step -1
        px(n) = xs(i): py(n) = ys(i): n = n + 1
        s = s + Dist(xs(i), ys(i), xs(i + 1), ys(i + 1))
        If n >= 3 Then
            FitLineLS px, py, n, mx, my, dx, dy, rms
            dev = 0
            For k = 0 To n - 1
                e = Abs((px(k) - mx) * dy - (py(k) - my) * dx)
                If e > dev Then
                    dev = e
                End If
            Next k
            If dev > 0.015 * w Then
                Exit For
            End If
            If s >= 0.4 * w Then
                If (xs(np_ - 1) - xs(i)) * dx + (ys(np_ - 1) - ys(i)) * dy < 0 Then
                    dx = -dx: dy = -dy
                End If
                fmx = mx: fmy = my: fdx = dx: fdy = dy: bestL = s: hasBest = True
                nb = n
                For k = 0 To n - 1
                    bx(k) = px(k): by(k) = py(k)
                Next k
            End If
        End If
    Next i
    If Not hasBest Then
        Exit Function
    End If
    If bestL >= w Then
        ' пологая дуга (терминалы C, S) тоже укладывается в 0.015w на 1w: по точкам участка
        ' без последних 0.3w (там остаток «крючка» конца) окружность R < 8w ложится вдвое
        ' лучше прямой — это дуга, не прямая
        ReDim ix(0 To nb): ReDim iy(0 To nb): ni = 0: sk = 0
        For k = 1 To nb - 1
            sk = sk + Dist(bx(k), by(k), bx(k - 1), by(k - 1))
            If sk >= 0.3 * w Then
                ix(ni) = bx(k): iy(ni) = by(k): ni = ni + 1
            End If
        Next k
        arc = False
        If ni >= 4 Then
            FitLineLS ix, iy, ni, mx, my, dx, dy, lr
            If FitCircle(ix, iy, ni, cx, cy, rr, rms) Then
                arc = (rr < 8 * w And rms < 0.002 * w And rms < 0.5 * lr)
            End If
        End If
        If Not arc Then
            f(0) = 2: f(1) = fmx: f(2) = fmy: f(3) = fdx: f(4) = fdy: f(5) = 0
            RefineLine f, bx, by, nb, w
            L = bestL
            NearLine = True
            Exit Function
        End If
    End If
    ' короткий участок: нужна прямая (на 0.8w) кромка контура с тем же направлением (в 12 град),
    ' на которой лежат точки участка (средняя дистанция < 0.25w)
    cosLim = Cos(12 * PI_ / 180)
    ReDim hf(0 To 5, 0 To 7)
    For c = 0 To NBC - 1
        If Not BCC(c) Then
            If Dist(BCX(c), BCY(c), xs(np_ - 1), ys(np_ - 1)) <= 2 * w Then
                nh = 0
                EdgeLines c, w, hf, nh
                For j = 0 To nh - 1
                    side = (-hf(5, j) - 1) - 2 * c
                    If side = 0 Then
                        straight = BS1(c)
                    Else
                        straight = BS2(c)
                    End If
                    If straight <= 0.02 Then
                        If Abs(hf(3, j) * fdx + hf(4, j) * fdy) >= cosLim Then
                            d = 0
                            For k = 0 To nb - 1
                                d = d + Abs((bx(k) - hf(1, j)) * hf(4, j) - (by(k) - hf(2, j)) * hf(3, j))
                            Next k
                            d = d / nb
                            If d < 0.25 * w Then
                                dx = hf(3, j): dy = hf(4, j)
                                If dx * fdx + dy * fdy < 0 Then
                                    dx = -dx: dy = -dy
                                End If
                                f(0) = 2: f(1) = hf(1, j): f(2) = hf(2, j): f(3) = dx: f(4) = dy: f(5) = 0
                                L = bestL
                                NearLine = True
                                Exit Function
                            End If
                        End If
                    End If
                Next j
            End If
        End If
    Next c
End Function


Private Sub RefineLine(f() As Double, px() As Double, py() As Double, ByVal n As Long, ByVal w As Double)
    Dim it As Long, mx As Double, my As Double, dx As Double, dy As Double, nx0 As Double, ny0 As Double
    Dim wx() As Double, wy() As Double, wl() As Double, wr() As Double, nw As Long, i As Long, t As Double
    Dim x As Double, y As Double, dl As Double, dr As Double, ws() As Double, wmed As Double, k As Long, tmp As Double
    Dim ax() As Double, ay() As Double, na As Long, need As Long, nmx As Double, nmy As Double, ndx As Double, ndy As Double, rms As Double
    mx = f(1): my = f(2): dx = f(3): dy = f(4)
    For it = 1 To 2
        nx0 = -dy: ny0 = dx
        ReDim wx(0 To n): ReDim wy(0 To n): ReDim wl(0 To n): ReDim wr(0 To n): nw = 0
        For i = 0 To n - 1
            t = (px(i) - mx) * dx + (py(i) - my) * dy
            x = mx + dx * t: y = my + dy * t
            dl = RayUnion(x, y, nx0, ny0): dr = RayUnion(x, y, -nx0, -ny0)
            If dl < BIG And dr < BIG Then
                wx(nw) = x: wy(nw) = y: wl(nw) = dl: wr(nw) = dr: nw = nw + 1
            End If
        Next i
        If nw < 3 Then
            Exit Sub
        End If
        ReDim ws(0 To nw)
        For i = 0 To nw - 1
            ws(i) = wl(i) + wr(i)
        Next i
        ' медиана ширины (сортировка вставками)
        For i = 1 To nw - 1
            tmp = ws(i): k = i - 1
            Do While k >= 0
                If ws(k) <= tmp Then
                    Exit Do
                End If
                ws(k + 1) = ws(k): k = k - 1
            Loop
            ws(k + 1) = tmp
        Next i
        wmed = ws(nw \ 2)
        ReDim ax(0 To nw): ReDim ay(0 To nw): na = 0
        For i = 0 To nw - 1
            If Abs(wl(i) + wr(i) - wmed) < 0.1 * wmed And Abs(wl(i) - wr(i)) < 0.3 * w Then
                ax(na) = wx(i) + nx0 * (wl(i) - wr(i)) / 2: ay(na) = wy(i) + ny0 * (wl(i) - wr(i)) / 2: na = na + 1
            End If
        Next i
        need = n \ 3
        If need < 3 Then
            need = 3
        End If
        If na < need Then
            Exit Sub
        End If
        FitLineLS ax, ay, na, nmx, nmy, ndx, ndy, rms
        If ndx * dx + ndy * dy < 0 Then
            ndx = -ndx: ndy = -ndy
        End If
        mx = nmx: my = nmy: dx = ndx: dy = ndy
        f(1) = mx: f(2) = my: f(3) = dx: f(4) = dy
    Next it
End Sub


Private Function BlendArc(xs() As Double, ys() As Double, np_ As Long, f() As Double, ByVal L As Double, _
                          ByVal w As Double, ByVal nmax As Long) As Boolean
    Dim imin As Long, S() As Double, i As Long, s1 As Double, hasS1 As Boolean, mx As Double, my As Double
    Dim dx As Double, dy As Double, px() As Double, py() As Double, n As Long, cx As Double, cy As Double, rr As Double, rms As Double
    Dim kf As Long, h As Double, ex As Double, ey As Double, dLn As Double, nArc As Long
    Dim acx(0 To 1) As Double, acy(0 To 1) As Double, arr(0 To 1) As Double, ap0x(0 To 1) As Double, ap0y(0 To 1) As Double
    Dim ap1x(0 To 1) As Double, ap1y(0 To 1) As Double, ax0 As Double, ay0 As Double, tx As Double, ty As Double
    Dim tT As Double, tF As Double, vx As Double, vy As Double, b As Double, disc As Double, r1 As Double, r2 As Double
    Dim tx_ As Double, tA As Double, n0x As Double, n0y As Double, vax As Double, vay As Double, den As Double, r As Double
    Dim cfx As Double, cfy As Double, tbx As Double, tby As Double, nx() As Double, ny() As Double, nn As Long
    Dim k As Long, a0 As Double, a1 As Double, da As Double, na As Long, a As Double, Ls As Double, nl As Long, k0 As Long
    Dim gp As Double, sgV As Double
    imin = 0
    If nmax >= 0 Then
        If np_ - nmax > imin Then
            imin = np_ - nmax
        End If
    End If
    If np_ - imin < 6 Then
        Exit Function
    End If
    ReDim S(0 To np_)
    For i = np_ - 2 To imin Step -1
        S(i) = S(i + 1) + Dist(xs(i), ys(i), xs(i + 1), ys(i + 1))
    Next i
    mx = f(1): my = f(2): dx = f(3): dy = f(4)
    ex = xs(np_ - 1): ey = ys(np_ - 1)
    ' окно окружности: от первой точки, явно сошедшей с прямой (> 0.08w), на 1.4w дальше
    For i = np_ - 1 To imin Step -1
        If S(i) > L And Abs((xs(i) - mx) * dy - (ys(i) - my) * dx) > 0.08 * w Then
            s1 = S(i): hasS1 = True
            Exit For
        End If
    Next i
    If Not hasS1 Then
        Exit Function
    End If
    If S(imin) < s1 + 1.4 * w Then
        Exit Function
    End If
    ReDim px(0 To np_): ReDim py(0 To np_): n = 0
    For i = imin To np_ - 1
        If S(i) >= s1 And S(i) <= s1 + 1.4 * w Then
            px(n) = xs(i): py(n) = ys(i): n = n + 1
        End If
    Next i
    If n < 4 Then
        Exit Function
    End If
    If Not FitCircle(px, py, n, cx, cy, rr, rms) Then
        Exit Function
    End If
    If rr > 25 * w Or rms > 0.03 * w Then
        Exit Function
    End If
    kf = -1
    For i = imin To np_ - 1
        If S(i) >= s1 + 1.4 * w Then
            kf = i
        End If
    Next i
    If kf < 0 Then
        Exit Function
    End If
    h = S(imin) / (np_ - 1 - imin)
    dLn = Abs((cx - mx) * dy - (cy - my) * dx)
    nArc = 0
    If dLn > rr And dLn <= rr + 0.15 * w Then
        ' дуга шрифта не точно окружность: подогнанная чуть не доходит до прямой — сдвинуть
        ' центр к прямой до касания (на величину зазора, < 0.15w)
        gp = dLn - rr
        If (cx - mx) * dy - (cy - my) * dx > 0 Then
            sgV = 1
        Else
            sgV = -1
        End If
        cx = cx - sgV * dy * gp: cy = cy + sgV * dx * gp
        dLn = rr
    End If
    If Abs(dLn - rr) <= 0.015 * w Then
        tx = cx: ty = cy: FitProj f, tx, ty
        tT = (ex - tx) * dx + (ey - ty) * dy
        tF = (ex - xs(kf)) * dx + (ey - ys(kf)) * dy
        If Not (tF > tT And tT > 0.1 * w) Then
            BlendArc = BlendHerm(xs, ys, np_, kf, ex, ey, dx, dy, L, h)
            Exit Function
        End If
        acx(0) = cx: acy(0) = cy: arr(0) = rr: ap0x(0) = xs(kf): ap0y(0) = ys(kf): ap1x(0) = tx: ap1y(0) = ty
        nArc = 1: ax0 = tx: ay0 = ty
    ElseIf dLn < rr Then
        If ACosD(dLn / rr) > 20 Then
            Exit Function
        End If
        ' пересечение прямой с окружностью: X(t) = E - d*t, ближайшее к концу (t > 0)
        vx = ex - cx: vy = ey - cy
        b = vx * dx + vy * dy: disc = b * b - (vx * vx + vy * vy - rr * rr)
        If disc < 0 Then
            Exit Function
        End If
        r1 = b - Sqr(disc): r2 = b + Sqr(disc)
        tx_ = -1
        If r1 > 0.1 * w Then
            tx_ = r1
        ElseIf r2 > 0.1 * w Then
            tx_ = r2
        End If
        If tx_ < 0 Then
            Exit Function
        End If
        tA = L
        If tx_ - 0.1 * w < tA Then
            tA = tx_ - 0.1 * w
        End If
        If tA < 0.1 * w Then
            Exit Function
        End If
        ax0 = ex - dx * tA: ay0 = ey - dy * tA
        n0x = -dy: n0y = dx
        If (cx - ax0) * n0x + (cy - ay0) * n0y < 0 Then
            n0x = dy: n0y = -dx
        End If
        vax = ax0 - cx: vay = ay0 - cy
        den = 2 * (rr + vax * n0x + vay * n0y)
        If den <= 0.000000001 Then
            Exit Function
        End If
        r = (rr * rr - vax * vax - vay * vay) / den
        If r <= 0.1 * w Or r >= rr Then
            Exit Function
        End If
        cfx = ax0 + n0x * r: cfy = ay0 + n0y * r
        tbx = cx + (cfx - cx) * rr / (rr - r): tby = cy + (cfy - cy) * rr / (rr - r)
        acx(0) = cx: acy(0) = cy: arr(0) = rr: ap0x(0) = xs(kf): ap0y(0) = ys(kf): ap1x(0) = tbx: ap1y(0) = tby
        acx(1) = cfx: acy(1) = cfy: arr(1) = r: ap0x(1) = tbx: ap0y(1) = tby: ap1x(1) = ax0: ap1y(1) = ay0
        nArc = 2
    Else
        Exit Function
    End If
    ReDim nx(0 To 63): ReDim ny(0 To 63): nn = 0
    For k = 0 To nArc - 1
        a0 = Atan2(ap0y(k) - acy(k), ap0x(k) - acx(k)): a1 = Atan2(ap1y(k) - acy(k), ap1x(k) - acx(k))
        da = a1 - a0
        Do While da > PI_
            da = da - 2 * PI_
        Loop
        Do While da <= -PI_
            da = da + 2 * PI_
        Loop
        na = Int(Abs(da) * arr(k) / h)
        If na < 1 Then
            na = 1
        End If
        If nn = 0 Then
            k0 = 0
        Else
            k0 = 1
        End If
        For i = k0 To na
            a = a0 + da * i / na
            PushPt nx, ny, nn, acx(k) + arr(k) * Cos(a), acy(k) + arr(k) * Sin(a)
        Next i
    Next k
    Ls = Dist(ax0, ay0, ex, ey)
    nl = Int(Ls / h)
    If nl < 1 Then
        nl = 1
    End If
    For k = 1 To nl
        PushPt nx, ny, nn, ax0 + dx * Ls * k / nl, ay0 + dy * Ls * k / nl
    Next k
    np_ = kf
    For i = 0 To nn - 1
        PushPt xs, ys, np_, nx(i), ny(i)
    Next i
    BlendArc = True
End Function


Private Sub LandOffset(xs() As Double, ys() As Double, ByVal np_ As Long, ByVal w As Double)
    Dim ox As Double, oy As Double, m As Double, L As Double, s As Double, i As Long, u As Double, wt As Double
    If np_ < 3 Then
        Exit Sub
    End If
    ox = xs(np_ - 1) - xs(np_ - 2): oy = ys(np_ - 1) - ys(np_ - 2)
    m = Sqr(ox * ox + oy * oy)
    If m < 0.01 * w Then
        Exit Sub
    End If
    L = 1.5 * w
    If 6 * m > L Then
        L = 6 * m
    End If
    s = 0
    For i = np_ - 2 To 1 Step -1
        If i < np_ - 2 Then
            s = s + Dist(xs(i), ys(i), xs(i + 1), ys(i + 1))
        End If
        If s >= L Then
            Exit For
        End If
        u = s / L: wt = 2 * u * u * u - 3 * u * u + 1
        xs(i) = xs(i) + ox * wt: ys(i) = ys(i) + oy * wt
    Next i
End Sub


Private Function FitTouch(f1() As Double, f2() As Double, ByVal w As Double, xX As Double, yX As Double) As Boolean
    Dim d As Double, ra As Double, rb As Double, ux As Double, uy As Double, sg As Double
    Dim mx As Double, my As Double, dx As Double, dy As Double, cx As Double, cy As Double, rr As Double, t0 As Double
    If f1(0) = 1 And f2(0) = 1 Then
        d = Dist(f1(1), f1(2), f2(1), f2(2))
        If d < 0.000000001 Then
            Exit Function
        End If
        ra = f1(3): rb = f2(3)
        If Abs(d - (ra + rb)) < 0.15 * w Or Abs(d - Abs(ra - rb)) < 0.15 * w Then
            ux = (f2(1) - f1(1)) / d: uy = (f2(2) - f1(2)) / d
            If Abs(d - (ra + rb)) < 0.15 * w Or ra > rb Then
                sg = 1
            Else
                sg = -1
            End If
            xX = f1(1) + ux * ra * sg: yX = f1(2) + uy * ra * sg
            FitTouch = True
        End If
        Exit Function
    End If
    If f1(0) = 2 And f2(0) = 1 Then
        mx = f1(1): my = f1(2): dx = f1(3): dy = f1(4): cx = f2(1): cy = f2(2): rr = f2(3)
    ElseIf f1(0) = 1 And f2(0) = 2 Then
        mx = f2(1): my = f2(2): dx = f2(3): dy = f2(4): cx = f1(1): cy = f1(2): rr = f1(3)
    Else
        Exit Function
    End If
    t0 = (cx - mx) * dx + (cy - my) * dy
    xX = mx + dx * t0: yX = my + dy * t0
    If Abs(Dist(xX, yX, cx, cy) - rr) < 0.15 * w Then
        FitTouch = True
    End If
End Function


Private Function StraightRun(x() As Double, y() As Double, ByVal i0 As Long, ByVal i1 As Long, ByVal stp As Long, _
                             ByVal tol As Double, ByVal lmin As Double) As Long
    Dim k As Long, best As Long, dev As Double
    best = -1
    k = i0 + 2 * stp
    Do While (stp > 0 And k < i1) Or (stp < 0 And k > i1)
        If stp > 0 Then
            dev = LineDev(x, y, i0, k)
        Else
            dev = LineDev(x, y, k, i0)
        End If
        If dev > tol Then
            Exit Do
        End If
        If Dist(x(i0), y(i0), x(k), y(k)) >= lmin Then
            best = k
        End If
        k = k + stp
    Loop
    StraightRun = best
End Function


Private Sub JoinPaths(ByVal s0 As Long, ByVal eps As Double)
    Dim i As Long, j As Long, e As Long, nE As Long, kx() As Long, ky() As Long, ep() As Long, ee() As Long
    Dim cnt As Long, pi_ As Long, pj As Long, ei As Long, ej As Long, changed As Boolean, k As Long
    Dim x As Double, y As Double, ax As Double, ay As Double, bx As Double, by As Double
    Dim TS() As Double, nT As Long, s As Long, m As Long, segA0 As Long, segA1 As Long, segB0 As Long, segB1 As Long
    Dim pStart() As Long, pN As Long, w As Long, q As Long, ps As Long, pe As Long, pcl As Boolean
    Dim OC() As Double, nOC As Long, oStart() As Long
    Do
        changed = False
        ' концы незамкнутых путей: ключ = округлённые координаты
        nE = 0
        ReDim kx(0 To 2 * (NOP - s0) + 1): ReDim ky(0 To 2 * (NOP - s0) + 1)
        ReDim ep(0 To 2 * (NOP - s0) + 1): ReDim ee(0 To 2 * (NOP - s0) + 1)
        For i = s0 To NOP - 1
            If Not OPathClosed(i) Then
                ps = OPathStart(i)
                If i < NOP - 1 Then
                    pe = OPathStart(i + 1) - 1
                Else
                    pe = NOS - 1
                End If
                If pe >= ps Then
                    For e = 0 To 1
                        If e = 0 Then
                            x = OSeg(1, ps): y = OSeg(2, ps)
                        Else
                            x = OSeg(7, pe): y = OSeg(8, pe)
                        End If
                        kx(nE) = CLng(Round(x / eps)): ky(nE) = CLng(Round(y / eps)): ep(nE) = i: ee(nE) = e
                        nE = nE + 1
                    Next e
                End If
            End If
        Next i
        For i = 0 To nE - 1
            cnt = 0
            For j = 0 To nE - 1
                If kx(j) = kx(i) And ky(j) = ky(i) Then
                    cnt = cnt + 1
                    If j <> i Then
                        pj = ep(j): ej = ee(j)
                    End If
                End If
            Next j
            If cnt = 2 And pj <> ep(i) Then
                pi_ = ep(i): ei = ee(i)
                ' сегменты путей
                segA0 = OPathStart(pi_)
                If pi_ < NOP - 1 Then
                    segA1 = OPathStart(pi_ + 1) - 1
                Else
                    segA1 = NOS - 1
                End If
                segB0 = OPathStart(pj)
                If pj < NOP - 1 Then
                    segB1 = OPathStart(pj + 1) - 1
                Else
                    segB1 = NOS - 1
                End If
                ' a: конец в точке стыка (ei = 0 — перевернуть), b: начало в точке стыка (ej = 1 — перевернуть)
                If ei = 0 Then
                    ax = OSeg(1, segA0): ay = OSeg(2, segA0)
                Else
                    ax = OSeg(7, segA1): ay = OSeg(8, segA1)
                End If
                If ej = 1 Then
                    bx = OSeg(7, segB1): by = OSeg(8, segB1)
                Else
                    bx = OSeg(1, segB0): by = OSeg(2, segB0)
                End If
                If Dist(ax, ay, bx, by) <= 2 * eps Then
                    ' собрать объединённый путь во временный буфер
                    nT = 0
                    ReDim TS(0 To 8, 0 To (segA1 - segA0 + 1) + (segB1 - segB0 + 1) + 1)
                    If ei = 0 Then
                        For s = segA1 To segA0 Step -1
                            RevSegInto TS, nT, s
                        Next s
                    Else
                        For s = segA0 To segA1
                            For k = 0 To 8: TS(k, nT) = OSeg(k, s): Next k
                            nT = nT + 1
                        Next s
                    End If
                    m = nT
                    If ej = 1 Then
                        For s = segB1 To segB0 Step -1
                            RevSegInto TS, nT, s
                        Next s
                    Else
                        For s = segB0 To segB1
                            For k = 0 To 8: TS(k, nT) = OSeg(k, s): Next k
                            nT = nT + 1
                        Next s
                    End If
                    TS(1, m) = ax: TS(2, m) = ay   ' стык: начало b = конец a
                    ' перестроить сегменты группы: пути по порядку, pi_ заменён объединённым, pj удалён
                    ReDim pStart(0 To NOP - s0)
                    w = OPathStart(s0)
                    ' копия всех сегментов группы
                    nOC = NOS - OPathStart(s0)
                    ReDim OC(0 To 8, 0 To nOC): ReDim oStart(0 To NOP - s0 + 1)
                    For s = 0 To nOC - 1
                        For k = 0 To 8: OC(k, s) = OSeg(k, OPathStart(s0) + s): Next k
                    Next s
                    For q = s0 To NOP - 1
                        oStart(q - s0) = OPathStart(q) - OPathStart(s0)
                    Next q
                    oStart(NOP - s0) = nOC
                    pN = 0
                    NOS = OPathStart(s0)
                    For q = s0 To NOP - 1
                        If q <> pj Then
                            pcl = OPathClosed(q)
                            OPathStart(s0 + pN) = NOS: OPathClosed(s0 + pN) = pcl: OPathGroup(s0 + pN) = gGroup
                            If q = pi_ Then
                                For s = 0 To nT - 1
                                    EmitSeg TS(0, s), TS(1, s), TS(2, s), TS(3, s), TS(4, s), TS(5, s), TS(6, s), TS(7, s), TS(8, s)
                                Next s
                            Else
                                For s = oStart(q - s0) To oStart(q - s0 + 1) - 1
                                    EmitSeg OC(0, s), OC(1, s), OC(2, s), OC(3, s), OC(4, s), OC(5, s), OC(6, s), OC(7, s), OC(8, s)
                                Next s
                            End If
                            pN = pN + 1
                        End If
                    Next q
                    NOP = s0 + pN
                    changed = True
                    Exit For
                End If
            End If
        Next i
    Loop While changed
End Sub


Private Sub RevSegInto(TS() As Double, nT As Long, ByVal s As Long)
    TS(0, nT) = OSeg(0, s)
    TS(1, nT) = OSeg(7, s): TS(2, nT) = OSeg(8, s)
    TS(3, nT) = OSeg(5, s): TS(4, nT) = OSeg(6, s)
    TS(5, nT) = OSeg(3, s): TS(6, nT) = OSeg(4, s)
    TS(7, nT) = OSeg(1, s): TS(8, nT) = OSeg(2, s)
    nT = nT + 1
End Sub


Private Function ShortSideLine(xs() As Double, ys() As Double, ByVal np_ As Long, ax_() As Double, ay_() As Double, _
                               ByVal na As Long, ByVal stp As Double, ByVal w As Double, fx() As Double, fy() As Double, _
                               nf As Long, pf() As Double) As Boolean
    Dim lb As Double, la As Double, i As Long, longB As Boolean, k As Long, s As Double, far As Double, fi As Long
    Dim px() As Double, py() As Double, n As Long, mx As Double, my As Double, dx As Double, dy As Double, rms As Double
    Dim ok As Boolean, x As Double, y As Double, t As Double, a0 As Double, a1 As Double, b0 As Double, b1 As Double, m As Long
    For i = 1 To np_ - 1
        lb = lb + Dist(xs(i), ys(i), xs(i - 1), ys(i - 1))
    Next i
    For i = 1 To na - 1
        la = la + Dist(ax_(i), ay_(i), ax_(i - 1), ay_(i - 1))
    Next i
    If lb < la Then
        If lb >= 1.5 * w Or la < 2.8 * w Then
            Exit Function
        End If
    Else
        If la >= 1.5 * w Or lb < 2.8 * w Then
            Exit Function
        End If
    End If
    longB = (lb >= la)
    ReDim px(0 To np_ + na + 2): ReDim py(0 To np_ + na + 2)
    For fi = 0 To 2
        If fi = 0 Then
            far = 2.8
        ElseIf fi = 1 Then
            far = 1.8
        Else
            far = 1.4
        End If
        n = 0: s = 0
        If longB Then
            For k = np_ - 2 To 0 Step -1
                s = s + Dist(xs(k), ys(k), xs(k + 1), ys(k + 1))
                If s > far * w Then
                    Exit For
                End If
                If s >= 0.8 * w Then
                    px(n) = xs(k): py(n) = ys(k): n = n + 1
                End If
            Next k
        Else
            For k = 1 To na - 1
                s = s + Dist(ax_(k), ay_(k), ax_(k - 1), ay_(k - 1))
                If s > far * w Then
                    Exit For
                End If
                If s >= 0.8 * w Then
                    px(n) = ax_(k): py(n) = ay_(k): n = n + 1
                End If
            Next k
        End If
        If n >= 3 Then
            If longB Then
                For k = 0 To na - 1
                    px(n) = ax_(k): py(n) = ay_(k): n = n + 1
                Next k
            Else
                For k = 0 To np_ - 1
                    px(n) = xs(k): py(n) = ys(k): n = n + 1
                Next k
            End If
            FitLineLS px, py, n, mx, my, dx, dy, rms
            If rms <= 0.02 * w Then
                ok = True
                Exit For
            End If
        End If
    Next fi
    If Not ok Then
        Exit Function
    End If
    ' короткая сторона — на прямой (< 0.2w)
    If longB Then
        For k = 0 To na - 1
            If Abs((ax_(k) - mx) * dy - (ay_(k) - my) * dx) > 0.2 * w Then
                Exit Function
            End If
        Next k
    Else
        For k = 0 To np_ - 1
            If Abs((xs(k) - mx) * dy - (ys(k) - my) * dx) > 0.2 * w Then
                Exit Function
            End If
        Next k
    End If
    pf(0) = 2: pf(1) = mx: pf(2) = my: pf(3) = dx: pf(4) = dy: pf(5) = 0
    If longB Then
        i = np_ - 1: s = 0
        Do While i >= 0
            FitProj pf, xs(i), ys(i)
            If i > 0 Then
                s = s + Dist(xs(i), ys(i), xs(i - 1), ys(i - 1))
            End If
            If s > 0.8 * w Then
                Exit Do
            End If
            i = i - 1
        Loop
    Else
        For i = 0 To np_ - 1
            FitProj pf, xs(i), ys(i)
        Next i
    End If
    a0 = xs(np_ - 1): a1 = ys(np_ - 1)
    b0 = ax_(0): b1 = ay_(0): FitProj pf, b0, b1
    m = Int(Dist(a0, a1, b0, b1) / stp)
    If m < 1 Then
        m = 1
    End If
    nf = 0
    For k = 1 To m - 1
        x = a0 + (b0 - a0) * k / m: y = a1 + (b1 - a1) * k / m
        FitProj pf, x, y
        PushPt fx, fy, nf, x, y
    Next k
    ShortSideLine = True
End Function


Private Function LeafBefore(ByVal a As Long, ByVal b As Long) As Boolean
    Dim ya As Double, yb As Double, xa As Double, xb As Double
    ya = Round(GY(a), 3): yb = Round(GY(b), 3)
    If ya < yb Then
        LeafBefore = True
    ElseIf ya = yb Then
        xa = Round(GX(a), 3): xb = Round(GX(b), 3)
        LeafBefore = (xa < xb)
    End If
End Function


Private Function StubToExt(xs() As Double, ys() As Double, np_ As Long, ByVal w As Double, rx() As Double, ry() As Double, _
                           ByVal nr_ As Long, ByVal jwJ As Double, ByVal jx As Double, ByVal jy As Double, _
                           ByVal stp As Double, ByVal maxlen As Double) As Boolean
    Dim px() As Double, py() As Double, n As Long, s As Double, i As Long, mx As Double, my As Double
    Dim dx As Double, dy As Double, rms As Double, f(0 To 5) As Double, nq As Long, sx As Double, sy As Double
    Dim iv As Long, sd As Long, nbx As Double, nby As Double, ex As Double, ey As Double, ux(0 To 1) As Double, uy(0 To 1) As Double
    Dim k As Long, pvx As Double, pvy As Double, x As Double, y As Double, hx As Double, hy As Double
    Dim hasBest As Boolean, bestL As Double, bhx As Double, bhy As Double, bvx As Double, bvy As Double, L As Double, m As Long
    If nr_ < 3 Then
        Exit Function
    End If
    ReDim px(0 To np_): ReDim py(0 To np_)
    For i = np_ - 2 To 0 Step -1
        s = s + Dist(xs(i), ys(i), xs(i + 1), ys(i + 1))
        If s > 3 * w Then
            Exit For
        End If
        If s >= 0.5 * w Then
            px(n) = xs(i): py(n) = ys(i): n = n + 1
        End If
    Next i
    If n < 3 Then
        Exit Function
    End If
    FitLineLS px, py, n, mx, my, dx, dy, rms
    If rms >= 0.02 * w Then
        Exit Function
    End If
    If (xs(np_ - 1) - px(0)) * dx + (ys(np_ - 1) - py(0)) * dy < 0 Then
        dx = -dx: dy = -dy
    End If
    f(0) = 2: f(1) = mx: f(2) = my: f(3) = dx: f(4) = dy: f(5) = 0
    ' обрезать конец, сошедший с прямой (> 0.1w)
    nq = np_
    Do While nq > 2
        x = xs(nq - 1): y = ys(nq - 1): FitProj f, x, y
        If Dist(x, y, xs(nq - 1), ys(nq - 1)) <= 0.1 * w Then
            Exit Do
        End If
        nq = nq - 1
    Loop
    sx = xs(nq - 1): sy = ys(nq - 1): FitProj f, sx, sy
    For iv = 1 To nr_ - 2
        If Dist(rx(iv), ry(iv), jx, jy) <= jwJ Then
            For sd = 0 To 1
                If sd = 0 Then
                    nbx = rx(iv - 1): nby = ry(iv - 1)
                Else
                    nbx = rx(iv + 1): nby = ry(iv + 1)
                End If
                ex = rx(iv) - nbx: ey = ry(iv) - nby
                If Norm(ex, ey) Then
                    ux(0) = rx(iv): uy(0) = ry(iv): ux(1) = rx(iv) + ex * jwJ: uy(1) = ry(iv) + ey * jwJ
                    pvx = sx: pvy = sy: k = 1
                    Do While k * stp <= maxlen
                        FitAt f, sx, sy, k * stp, x, y
                        If SegHit(ux, uy, 2, pvx, pvy, x, y, hx, hy) Then
                            If Dist(hx, hy, rx(iv), ry(iv)) > 0.02 * w And InUnion(hx, hy, -1) Then
                                If SegInside(sx, sy, hx, hy) And SegInside(hx, hy, rx(iv), ry(iv)) Then
                                    L = Dist(sx, sy, hx, hy)
                                    If (Not hasBest) Or L < bestL Then
                                        hasBest = True: bestL = L: bhx = hx: bhy = hy: bvx = rx(iv): bvy = ry(iv)
                                    End If
                                End If
                            End If
                            Exit Do
                        End If
                        pvx = x: pvy = y: k = k + 1
                    Loop
                End If
            Next sd
        End If
    Next iv
    If Not hasBest Then
        Exit Function
    End If
    np_ = nq - 1
    PushPt xs, ys, np_, sx, sy
    m = Int(bestL / stp)
    If m < 1 Then
        m = 1
    End If
    For i = 1 To m - 1
        FitAt f, sx, sy, bestL * i / m, x, y
        PushPt xs, ys, np_, x, y
    Next i
    PushPt xs, ys, np_, bhx, bhy
    PushPt xs, ys, np_, bvx, bvy
    StubToExt = True
End Function


Private Function TangentBridge(xs() As Double, ys() As Double, ByVal np_ As Long, ax_() As Double, ay_() As Double, _
                               ByVal na As Long, ByVal stp As Double, ByVal w As Double, fx() As Double, fy() As Double, _
                               nf As Long, pf() As Double) As Boolean
    Dim bx() As Double, by() As Double, nb As Long, cx_() As Double, cy_() As Double, nc As Long
    Dim fa(0 To 5) As Double, fb(0 To 5) As Double, a0 As Double, a1 As Double, b0 As Double, b1 As Double
    Dim tax As Double, tay As Double, tbx As Double, tby As Double, L0 As Double, dnx As Double, dny As Double
    Dim sa As Long, sb As Long, sg As Long, kk As Double, nx0 As Double, ny0 As Double, q As Double
    Dim c1x As Double, c1y As Double, c2x As Double, c2y As Double, t1x As Double, t1y As Double, t2x As Double, t2y As Double
    Dim cand(0 To 3, 0 To 7) As Double, ncand As Long, ci As Long, Lt As Double, dx As Double, dy As Double
    Dim u1x As Double, u1y As Double, u2x As Double, u2y As Double, sc As Double, bestSc As Double, hasBest As Boolean
    Dim B1x As Double, B1y As Double, B2x As Double, B2y As Double, mx As Double, my As Double, t0 As Double, fxx As Double, fyy As Double
    Dim i As Long, s As Double, reach As Double, ta_ As Double, tb_ As Double, tq As Double, m As Long, k As Long
    Dim tA0 As Double, tA1 As Double, bpx As Double, bpy As Double, rr As Double
    TakePts xs, ys, np_, True, 2 * w, bx, by, nb
    TakePts ax_, ay_, na, False, 2 * w, cx_, cy_, nc
    If nb < 2 Or nc < 2 Then
        Exit Function
    End If
    If Not FitPts(bx, by, nb, w, 0.02, 0.02, fa) Then
        Exit Function
    End If
    If Not FitPts(cx_, cy_, nc, w, 0.02, 0.02, fb) Then
        Exit Function
    End If
    If fa(0) = 2 And fb(0) = 2 Then
        Exit Function
    End If
    a0 = xs(np_ - 1): a1 = ys(np_ - 1): b0 = ax_(0): b1 = ay_(0)
    tax = bx(0) - bx(1): tay = by(0) - by(1): Norm tax, tay
    tbx = cx_(1) - cx_(0): tby = cy_(1) - cy_(0): Norm tbx, tby
    ncand = 0
    If fa(0) = 1 And fb(0) = 1 Then
        dnx = fb(1) - fa(1): dny = fb(2) - fa(2): L0 = Sqr(dnx * dnx + dny * dny)
        If L0 < 0.000000001 Then
            Exit Function
        End If
        dnx = dnx / L0: dny = dny / L0
        For sa = -1 To 1 Step 2
            For sb = -1 To 1 Step 2
                kk = (sa * fa(3) - sb * fb(3)) / L0
                If Abs(kk) <= 1 Then
                    q = Sqr(1 - kk * kk)
                    For sg = -1 To 1 Step 2
                        nx0 = kk * dnx + sg * q * (-dny): ny0 = kk * dny + sg * q * dnx
                        If ncand <= 3 Then
                            cand(ncand, 0) = fa(1) - sa * fa(3) * nx0: cand(ncand, 1) = fa(2) - sa * fa(3) * ny0
                            cand(ncand, 2) = fb(1) - sb * fb(3) * nx0: cand(ncand, 3) = fb(2) - sb * fb(3) * ny0
                            ncand = ncand + 1
                        End If
                    Next sg
                End If
            Next sb
        Next sa
    Else
        If fa(0) = 2 Then
            mx = fa(1): my = fa(2): dx = fa(3): dy = fa(4): c1x = fb(1): c1y = fb(2): rr = fb(3)
        Else
            mx = fb(1): my = fb(2): dx = fb(3): dy = fb(4): c1x = fa(1): c1y = fa(2): rr = fa(3)
        End If
        t0 = (c1x - mx) * dx + (c1y - my) * dy
        fxx = mx + dx * t0: fyy = my + dy * t0
        If Abs(Dist(fxx, fyy, c1x, c1y) - rr) > 0.1 * w Then
            Exit Function
        End If
        If fa(0) = 1 Then
            cand(0, 0) = fxx: cand(0, 1) = fyy
            cand(0, 2) = a0: cand(0, 3) = a1: FitProj fb, cand(0, 2), cand(0, 3)
        Else
            cand(0, 0) = b0: cand(0, 1) = b1: FitProj fa, cand(0, 0), cand(0, 1)
            cand(0, 2) = fxx: cand(0, 3) = fyy
        End If
        ncand = 1
    End If
    For ci = 0 To ncand - 1
        t1x = cand(ci, 0): t1y = cand(ci, 1): t2x = cand(ci, 2): t2y = cand(ci, 3)
        Lt = Dist(t1x, t1y, t2x, t2y)
        If Lt >= 0.3 * w Then
            dx = (t2x - t1x) / Lt: dy = (t2y - t1y) / Lt
            FitTangent fa, t1x, t1y, tax, tay, u1x, u1y
            FitTangent fb, t2x, t2y, tbx, tby, u2x, u2y
            If dx * u1x + dy * u1y >= 0.98 And dx * u2x + dy * u2y >= 0.98 Then
                If (t1x - a0) * tax + (t1y - a1) * tay >= -w And (b0 - t2x) * tbx + (b1 - t2y) * tby >= -w Then
                    If SegInside(t1x, t1y, t2x, t2y) Then
                        sc = Dist(a0, a1, t1x, t1y) + Dist(t2x, t2y, b0, b1)
                        If (Not hasBest) Or sc < bestSc Then
                            hasBest = True: bestSc = sc: B1x = t1x: B1y = t1y: B2x = t2x: B2y = t2y
                        End If
                    End If
                End If
            End If
        End If
    Next ci
    If Not hasBest Then
        Exit Function
    End If
    Lt = Dist(B1x, B1y, B2x, B2y): dx = (B2x - B1x) / Lt: dy = (B2y - B1y) / Lt
    ' хвост «до» — на кусочную кривую (дуга fa до T1, дальше прямая)
    reach = -((a0 - B1x) * dx + (a1 - B1y) * dy) + 0.3 * w
    If reach < 0.8 * w Then
        reach = 0.8 * w
    End If
    i = np_ - 1: s = 0
    Do While i >= 0
        tq = (xs(i) - B1x) * dx + (ys(i) - B1y) * dy
        If tq < 0 Then
            FitProj fa, xs(i), ys(i)
        Else
            If tq > Lt Then
                tq = Lt
            End If
            xs(i) = B1x + dx * tq: ys(i) = B1y + dy * tq
        End If
        If i > 0 Then
            s = s + Dist(xs(i), ys(i), xs(i - 1), ys(i - 1))
        End If
        If s > reach Then
            Exit Do
        End If
        i = i - 1
    Loop
    a0 = xs(np_ - 1): a1 = ys(np_ - 1)
    ' начало «после» — на кусочную кривую (прямая до T2, дальше дуга fb)
    bpx = b0: bpy = b1
    tq = (bpx - B1x) * dx + (bpy - B1y) * dy
    If tq > Lt Then
        FitProj fb, bpx, bpy
    Else
        If tq < 0 Then
            tq = 0
        End If
        bpx = B1x + dx * tq: bpy = B1y + dy * tq
    End If
    ta_ = (a0 - B1x) * dx + (a1 - B1y) * dy
    tb_ = (bpx - B1x) * dx + (bpy - B1y) * dy
    nf = 0
    If ta_ < 0 Then
        FitBetween fa, a0, a1, B1x, B1y, stp, fx, fy, nf
        PushPt fx, fy, nf, B1x, B1y
    End If
    tA0 = ta_
    If tA0 < 0 Then
        tA0 = 0
    End If
    tA1 = tb_
    If tA1 > Lt Then
        tA1 = Lt
    End If
    m = Int((tA1 - tA0) / stp)
    If m < 1 Then
        m = 1
    End If
    For k = 1 To m - 1
        PushPt fx, fy, nf, B1x + dx * (tA0 + (tA1 - tA0) * k / m), B1y + dy * (tA0 + (tA1 - tA0) * k / m)
    Next k
    If tb_ > Lt Then
        PushPt fx, fy, nf, B2x, B2y
        FitBetween fb, B2x, B2y, bpx, bpy, stp, fx, fy, nf
    End If
    If fb(0) = 1 Then
        pf(0) = 3: pf(1) = B2x: pf(2) = B2y: pf(3) = dx: pf(4) = dy: pf(5) = 0
        BRCX = fb(1): BRCY = fb(2): BRCR = fb(3)
    Else
        For k = 0 To 5: pf(k) = fb(k): Next k
    End If
    TangentBridge = True
End Function


Private Function ParallelAt(ax() As Double, ay() As Double, ByVal na As Long, ByVal i As Long, rx() As Double, ry() As Double, _
                            ByVal nr_ As Long, ByVal cx As Double, ByVal cy As Double) As Boolean
    Dim tcx As Double, tcy As Double, tdx As Double, tdy As Double, it As Long, k As Long, d As Double, bd As Double
    Dim i0 As Long, i1 As Long
    If i <= 0 Or i >= na - 1 Or nr_ < 2 Then
        Exit Function
    End If
    tcx = ax(i + 1) - ax(i - 1): tcy = ay(i + 1) - ay(i - 1): Norm tcx, tcy
    bd = BIG
    For k = 0 To nr_ - 1
        d = Dist(rx(k), ry(k), cx, cy)
        If d < bd Then
            bd = d: it = k
        End If
    Next k
    i1 = it + 1
    If i1 > nr_ - 1 Then
        i1 = nr_ - 1
    End If
    i0 = it - 1
    If i0 < 0 Then
        i0 = 0
    End If
    tdx = rx(i1) - rx(i0): tdy = ry(i1) - ry(i0): Norm tdx, tdy
    ParallelAt = (Abs(tcx * tdx + tcy * tdy) >= Cos(15 * PI_ / 180))
End Function


Private Function BlendHerm(xs() As Double, ys() As Double, np_ As Long, ByVal kf As Long, ByVal ex As Double, ByVal ey As Double, _
                           ByVal dx As Double, ByVal dy As Double, ByVal L As Double, ByVal h As Double) As Boolean
    Dim ax0 As Double, ay0 As Double, tkx As Double, tky As Double, fx() As Double, fy() As Double, nf As Long
    Dim nl As Long, k As Long, sg As Long, sd As Long, nkx As Double, nky As Double, sn As Double, s0 As Double
    Dim den As Double, r As Double, cx As Double, cy As Double, tc As Double, tx As Double, ty As Double
    Dim a0 As Double, a1 As Double, da As Double, at As Double, ttx As Double, tty As Double, na As Long, Ls As Double, dirS As Double
    ax0 = ex - dx * L: ay0 = ey - dy * L
    If kf < 1 Then
        Exit Function
    End If
    If (ax0 - xs(kf)) * dx + (ay0 - ys(kf)) * dy <= 0 Then
        Exit Function
    End If
    tkx = xs(kf) - xs(kf - 1): tky = ys(kf) - ys(kf - 1): Norm tkx, tky
    If tkx * dx + tky * dy < 0.5 Then
        Exit Function
    End If
    ReDim fx(0 To 31): ReDim fy(0 To 31): nf = 0
    s0 = (xs(kf) - ex) * dy - (ys(kf) - ey) * dx
    For sg = 1 To -1 Step -2
        nkx = -tky * sg: nky = tkx * sg
        sn = nkx * dy - nky * dx
        For sd = 1 To -1 Step -2
            den = sd - sn
            If Abs(den) >= 0.000000001 Then
                r = s0 / den
                If r > 0 Then
                    cx = xs(kf) + nkx * r: cy = ys(kf) + nky * r
                    tc = (cx - ex) * dx + (cy - ey) * dy
                    tx = ex + dx * tc: ty = ey + dy * tc
                    If Abs(Dist(cx, cy, tx, ty) - r) <= 0.000001 * (1 + r) And -tc >= -0.05 * L And -tc <= 3 * L Then
                        a0 = Atan2(ys(kf) - cy, xs(kf) - cx): a1 = Atan2(ty - cy, tx - cx)
                        da = a1 - a0
                        Do While da < 0
                            da = da + 2 * PI_
                        Loop
                        Do While da >= 2 * PI_
                            da = da - 2 * PI_
                        Loop
                        If sg < 0 Then
                            da = da - 2 * PI_
                        End If
                        at = a0 + da
                        If da > 0 Then
                            dirS = 1
                        Else
                            dirS = -1
                        End If
                        ttx = -Sin(at) * dirS: tty = Cos(at) * dirS
                        If ttx * dx + tty * dy >= 0.98 And Abs(da) <= PI_ Then
                            na = Int(Abs(da) * r / h)
                            If na < 2 Then
                                na = 2
                            End If
                            For k = 1 To na
                                PushPt fx, fy, nf, cx + r * Cos(a0 + da * k / na), cy + r * Sin(a0 + da * k / na)
                            Next k
                            Ls = -tc
                            If Ls > 0.000001 Then
                                nl = Int(Ls / h)
                                If nl < 1 Then
                                    nl = 1
                                End If
                                For k = 1 To nl
                                    PushPt fx, fy, nf, tx + dx * Ls * k / nl, ty + dy * Ls * k / nl
                                Next k
                            End If
                            np_ = kf + 1
                            For k = 0 To nf - 1
                                PushPt xs, ys, np_, fx(k), fy(k)
                            Next k
                            BlendHerm = True
                            Exit Function
                        End If
                    End If
                End If
            End If
        Next sd
    Next sg
    nf = 0
    HermSeg fx, fy, nf, xs(kf), ys(kf), tkx, tky, ax0, ay0, dx, dy, h
    PushPt fx, fy, nf, ax0, ay0
    nl = Int(L / h)
    If nl < 1 Then
        nl = 1
    End If
    For k = 1 To nl
        PushPt fx, fy, nf, ax0 + dx * L * k / nl, ay0 + dy * L * k / nl
    Next k
    np_ = kf + 1
    For k = 0 To nf - 1
        PushPt xs, ys, np_, fx(k), fy(k)
    Next k
    BlendHerm = True
End Function

Private Sub MergeCollinear(xs() As Double, ys() As Double, fixedF() As Boolean, ByVal n As Long, ByVal closed As Boolean, ByVal w As Double)
    Dim pa() As Long, pb() As Long, pm() As Boolean, np As Long, i As Long, j As Long, best As Long
    Dim k As Long, q As Long, a0 As Long, a1 As Long, b0 As Long, b1 As Long, changed As Boolean, bad As Boolean
    Dim px() As Double, py() As Double, m As Long, mxa As Double, mya As Double, dxa As Double, dya As Double, ra As Double
    Dim mxb As Double, myb As Double, dxb As Double, dyb As Double, rb As Double, off As Double, mid_ As Long
    Dim t As Double, okA As Boolean, okB As Boolean
    If n < 6 Or closed Then
        Exit Sub
    End If
    ReDim pa(0 To n): ReDim pb(0 To n): ReDim pm(0 To n): np = 0
    i = 0
    Do While i < n - 2
        j = i + 2: best = -1
        Do While j < n
            If fixedF(j - 1) And j - 1 > i Then
                Exit Do
            End If
            If LineDev(xs, ys, i, j) > 0.02 * w Then
                Exit Do
            End If
            best = j: j = j + 1
        Loop
        If best >= 0 Then
            If Dist(xs(i), ys(i), xs(best), ys(best)) >= w Then
                pa(np) = i: pb(np) = best: pm(np) = False: np = np + 1
                i = best + 1
            Else
                i = i + 1
            End If
        Else
            i = i + 1
        End If
    Loop
    ReDim px(0 To n): ReDim py(0 To n)
    Do
        changed = False
        For k = 0 To np - 2
            a0 = pa(k): a1 = pb(k): b0 = pa(k + 1): b1 = pb(k + 1)
            bad = (b0 - a1 > 3)
            If Not bad Then
                For q = a1 To b0
                    If fixedF(q) Then
                        bad = True
                    End If
                Next q
            End If
            If Not bad Then
                m = 0
                For q = a0 To a1: px(m) = xs(q): py(m) = ys(q): m = m + 1: Next q
                FitLineLS px, py, m, mxa, mya, dxa, dya, ra
                m = 0
                For q = b0 To b1: px(m) = xs(q): py(m) = ys(q): m = m + 1: Next q
                FitLineLS px, py, m, mxb, myb, dxb, dyb, rb
                If Abs(dxa * dxb + dya * dyb) >= Cos(3 * PI_ / 180) Then
                    mid_ = (b0 + b1) \ 2
                    off = Abs((xs(mid_) - mxa) * dya - (ys(mid_) - mya) * dxa)
                    If off <= 0.15 * w Then
                        pb(k) = b1: pm(k) = True
                        For q = k + 1 To np - 2
                            pa(q) = pa(q + 1): pb(q) = pb(q + 1): pm(q) = pm(q + 1)
                        Next q
                        np = np - 1: changed = True
                        Exit For
                    End If
                End If
            End If
        Next k
    Loop While changed
    For k = 0 To np - 1
        If pm(k) Then
            a0 = pa(k): a1 = pb(k)
            m = 0
            For q = a0 To a1: px(m) = xs(q): py(m) = ys(q): m = m + 1: Next q
            FitLineLS px, py, m, mxa, mya, dxa, dya, ra
            Do While a0 > 0
                If fixedF(a0 - 1) Then
                    Exit Do
                End If
                If Abs((xs(a0 - 1) - mxa) * dya - (ys(a0 - 1) - mya) * dxa) >= 0.2 * w Then
                    Exit Do
                End If
                a0 = a0 - 1
            Loop
            Do While a1 < n - 1
                If fixedF(a1 + 1) Then
                    Exit Do
                End If
                If Abs((xs(a1 + 1) - mxa) * dya - (ys(a1 + 1) - mya) * dxa) >= 0.2 * w Then
                    Exit Do
                End If
                a1 = a1 + 1
            Loop
            okA = (a0 = 0)
            If Not okA Then
                okA = fixedF(a0 - 1) Or fixedF(a0)
            End If
            okB = (a1 = n - 1)
            If Not okB Then
                okB = fixedF(a1 + 1) Or fixedF(a1)
            End If
            If okA And okB Then
                For q = a0 To a1
                    t = (xs(q) - mxa) * dxa + (ys(q) - mya) * dya
                    xs(q) = mxa + dxa * t: ys(q) = mya + dya * t
                Next q
            End If
        End If
    Next k
End Sub


Private Sub Deglitch(xs() As Double, ys() As Double, ByVal n As Long, ByVal closed As Boolean, ByVal w As Double)
    Dim prot() As Boolean, i As Long, it As Long, nx() As Double, ny() As Double, i0 As Long, i1 As Long
    Dim ia As Long, ib As Long, d1x As Double, d1y As Double, d2x As Double, d2y As Double, lim As Double
    If n < 5 Then
        Exit Sub
    End If
    ReDim prot(0 To n - 1)
    For i = 0 To n - 1
        prot(i) = (TurnAtW(xs, ys, n, i, closed, w) > 30)
    Next i
    lim = Cos(8 * PI_ / 180)
    If closed Then
        i0 = 0: i1 = n - 1
    Else
        i0 = 1: i1 = n - 2
    End If
    ReDim nx(0 To n - 1): ReDim ny(0 To n - 1)
    For it = 1 To 4
        For i = 0 To n - 1: nx(i) = xs(i): ny(i) = ys(i): Next i
        For i = i0 To i1
            If Not prot(i) Then
                ia = (i - 1 + n) Mod n: ib = (i + 1) Mod n
                d1x = xs(i) - xs(ia): d1y = ys(i) - ys(ia): Norm d1x, d1y
                d2x = xs(ib) - xs(i): d2y = ys(ib) - ys(i): Norm d2x, d2y
                If d1x * d2x + d1y * d2y < lim Then
                    nx(i) = 0.5 * xs(i) + 0.25 * (xs(ia) + xs(ib)): ny(i) = 0.5 * ys(i) + 0.25 * (ys(ia) + ys(ib))
                End If
            End If
        Next i
        For i = 0 To n - 1: xs(i) = nx(i): ys(i) = ny(i): Next i
    Next it
End Sub


Private Function TurnAtW(x() As Double, y() As Double, ByVal n As Long, ByVal i As Long, ByVal closed As Boolean, ByVal w As Double) As Double
    Dim a As Long, b As Long, L As Double, a2 As Long, d1x As Double, d1y As Double, d2x As Double, d2y As Double
    a = i: L = 0
    Do While L < w * 0.5
        If Not closed And a = 0 Then
            Exit Do
        End If
        a2 = (a - 1 + n) Mod n: L = L + Dist(x(a), y(a), x(a2), y(a2)): a = a2
        If a = i Then
            Exit Do
        End If
    Loop
    b = i: L = 0
    Do While L < w * 0.5
        If Not closed And b = n - 1 Then
            Exit Do
        End If
        a2 = (b + 1) Mod n: L = L + Dist(x(b), y(b), x(a2), y(a2)): b = a2
        If b = i Then
            Exit Do
        End If
    Loop
    If a = i Or b = i Then
        TurnAtW = 180
        Exit Function
    End If
    d1x = x(i) - x(a): d1y = y(i) - y(a): Norm d1x, d1y
    d2x = x(b) - x(i): d2y = y(b) - y(i): Norm d2x, d2y
    TurnAtW = ACosD(d1x * d2x + d1y * d2y)
End Function


'@@CORE_END

' =====================================================================
'  CorelDRAW: точка входа и обмен с ядром
' =====================================================================
Public Sub NeonCenterline()
    Dim sr As ShapeRange, s As String, msg As String
    Dim oldUnit As cdrUnit, lay As Layer, col As Collection
    Dim i As Long, nIn As Long, nOut As Long, t0 As Single, tCore As Single

    If ActiveDocument Is Nothing Then
        Exit Sub
    End If
    Set sr = ActiveSelectionRange
    If sr.Count = 0 Then
        MsgBox UStr("412 44B 434 435 43B 438 442 435 20 43A 43E 43D 442 443 440 44B 20 431 443 43A 432 2E"), vbExclamation, UStr("41D 435 43E 43D 3A 20 446 435 43D 442 440 430 43B 44C 43D 430 44F 20 43B 438 43D 438 44F")
        Exit Sub
    End If

    s = InputBox(UStr("414 438 430 43C 435 442 440 20 444 440 435 437 44B 20 2F 20 440 430 441 441 435 438 432 430 442 435 43B 44F 2C 20 43C 43C 3A") & vbCr & _
                 UStr("28 43A 43E 43D 435 446 20 43B 438 43D 438 438 20 441 442 430 432 438 442 441 44F 20 43D 430 20 44 2F 32 20 43E 442 20 43A 440 430 44F 20 43A 43E 43D 442 443 440 430 29"), _
                 UStr("41D 435 43E 43D 3A 20 446 435 43D 442 440 430 43B 44C 43D 430 44F 20 43B 438 43D 438 44F"), GetSetting(REG_APP, "Params", "Tool", "8"))
    If Len(Trim$(s)) = 0 Then
        Exit Sub
    End If
    s = Replace(Trim$(s), ",", ".")
    If Val(s) <= 0 Then
        MsgBox UStr("41D 435 432 435 440 43D 44B 439 20 434 438 430 43C 435 442 440 2E"), vbExclamation, UStr("41D 435 43E 43D 3A 20 446 435 43D 442 440 430 43B 44C 43D 430 44F 20 43B 438 43D 438 44F")
        Exit Sub
    End If
    SaveSetting REG_APP, "Params", "Tool", s
    CoreSetParams Val(s), P_TOL

    oldUnit = ActiveDocument.Unit
    ActiveDocument.Unit = cdrMillimeter
    ActiveDocument.BeginCommandGroup UStr("41D 435 43E 43D 3A 20 446 435 43D 442 440 430 43B 44C 43D 430 44F 20 43B 438 43D 438 44F")
    Optimization = True
    EventsEnabled = False
    t0 = Timer

    On Error GoTo Fail
    Set col = New Collection
    CollectShapes sr, col
    ' контуры всех выделенных объектов — в одно ядро: буквы, которые перекрываются
    ' даже будучи разными объектами, считаются одной заливкой
    CoreInitRings
    For i = 1 To col.Count
        If ReadShape(col(i)) Then
            nIn = nIn + 1
        End If
    Next i
    If nIn = 0 Then
        msg = UStr("412 20 432 44B 434 435 43B 435 43D 438 438 20 43D 435 442 20 437 430 43C 43A 43D 443 442 44B 445 20 43A 43E 43D 442 443 440 43E 432 2E")
        GoTo Fin
    End If
    DumpInput Val(s)
    tCore = Timer
    CoreRun
    tCore = Timer - tCore
    Set lay = GetOutLayer()
    On Error Resume Next
    lay.Visible = True: lay.Editable = True
    On Error GoTo Fail
    nOut = EmitShapes(lay)
    msg = UStr("413 43E 442 43E 432 43E 3A 20 43E 431 44A 435 43A 442 43E 432 20") & nIn & UStr("2C 20 43B 438 43D 438 439 20 441 43E 437 434 430 43D 43E 20") & nOut
    msg = msg & UStr("2E 20 412 440 435 43C 44F 20") & Format$(Timer - t0, "0.0") & UStr("20 441 20 28 440 430 441 447 451 442 20") & Format$(tCore, "0.0") & UStr("20 441 29 2E")

Fin:
    On Error Resume Next
    EventsEnabled = True
    Optimization = False
    ActiveDocument.EndCommandGroup
    ActiveDocument.Unit = oldUnit
    ActiveWindow.Refresh
    Application.Refresh
    If Len(msg) > 0 Then
        MsgBox msg, vbInformation, UStr("41D 435 43E 43D 3A 20 446 435 43D 442 440 430 43B 44C 43D 430 44F 20 43B 438 43D 438 44F")
    End If
    Exit Sub
Fail:
    msg = UStr("41E 448 438 431 43A 430 3A 20") & Err.Description
    Resume Fin
End Sub

' входные контуры (как их прочитал макрос) — в файл NeonCenterline_in.txt во временной папке
' Windows: для разбора проблемных случаев на стенде. Ошибки записи не мешают работе.
Private Sub DumpInput(ByVal tool As Double)
    Dim f As Integer, r As Long, i As Long, path As String
    On Error GoTo Done
    path = Environ$("TEMP") & "\NeonCenterline_in.txt"
    f = FreeFile
    Open path For Output As #f
    Print #f, "T " & Str$(tool) & " " & Str$(P_TOL)
    For r = 0 To NAR - 1
        Print #f, "R"
        For i = AS_(r) To AS_(r) + AN_(r) - 1
            Print #f, Str$(AX(i)) & " " & Str$(AY(i))
        Next i
    Next r
    Close #f
    Exit Sub
Done:
    On Error Resume Next
    Close #f
End Sub

' строка из кодов Unicode (чтобы русский текст не зависел от кодировки Windows)
Private Function UStr(ByVal hx As String) As String
    Dim a As Variant, i As Long, r As String
    a = Split(hx, " ")
    For i = 0 To UBound(a)
        r = r & ChrW(CLng("&H" & a(i)))
    Next i
    UStr = r
End Function

Private Function OutPrefix() As String
    OutPrefix = UStr("41D 435 43E 43D 3A 20 446 435 43D 442 440 20")   ' «Неон: центр »
End Function

Private Function LayerName() As String
    LayerName = UStr("41D 435 43E 43D 20 2D 20 446 435 43D 442 440")
End Function

Private Sub CollectShapes(sr As ShapeRange, col As Collection)
    Dim sh As Shape
    For Each sh In sr
        If sh.Type = cdrGroupShape Then
            CollectShapes sh.Shapes.All, col
        ElseIf Left$(sh.Name, Len(OutPrefix())) <> OutPrefix() Then
            ' пропускаются только линии, созданные самим макросом (по имени), а не весь слой
            col.Add sh
        End If
    Next sh
End Sub

Private Function GetOutLayer() As Layer
    Dim l As Layer
    For Each l In ActivePage.Layers
        If l.Name = LayerName() Then
            Set GetOutLayer = l
            Exit Function
        End If
    Next l
    Set GetOutLayer = ActivePage.CreateLayer(LayerName())
End Function

' прочитать замкнутые контуры объекта в ядро (текст и фигуры — через временную копию в кривых)
Private Function ReadShape(sh As Shape) As Boolean
    Dim tmp As Shape, ch As Shape, n0 As Long
    n0 = NAR
    If sh.Type = cdrCurveShape Then
        AppendRings sh.Curve
    Else
        On Error Resume Next
        Set tmp = sh.Duplicate
        If Not tmp Is Nothing Then
            tmp.ConvertToCurves
        End If
        On Error GoTo 0
        If tmp Is Nothing Then
            Exit Function
        End If
        If tmp.Type = cdrCurveShape Then
            AppendRings tmp.Curve
        ElseIf tmp.Type = cdrGroupShape Then
            For Each ch In tmp.Shapes
                If ch.Type = cdrCurveShape Then
                    AppendRings ch.Curve
                End If
            Next ch
        End If
        tmp.Delete
    End If
    ReadShape = (NAR > n0)
End Function

' разбиение кривых Безье на отрезки с допуском P_FLAT_TOL (мм), независимо от размера букв
Private Sub AppendRings(crv As Curve)
    Dim sp As SubPath, sg As Segment, st As Long, n As Long, k As Long
    Dim x0 As Double, y0 As Double, x1 As Double, y1 As Double
    Dim x2 As Double, y2 As Double, x3 As Double, y3 As Double, dd As Double, d2 As Double, t As Double, mt As Double
    Dim isClosed As Boolean
    For Each sp In crv.SubPaths
        isClosed = sp.Closed
        If Not isClosed Then
            ' формально открытый, но конец совпадает с началом (после ручной правки) — тоже контур
            On Error Resume Next
            isClosed = (Abs(sp.StartNode.PositionX - sp.EndNode.PositionX) < 0.05 And _
                        Abs(sp.StartNode.PositionY - sp.EndNode.PositionY) < 0.05 And sp.Segments.Count >= 2)
            On Error GoTo 0
        End If
        If isClosed Then
            st = CoreRingStart()
            For Each sg In sp.Segments
                x0 = sg.StartNode.PositionX: y0 = sg.StartNode.PositionY
                CoreAddPoint x0, y0
                If sg.Type = cdrCurveSegment Then
                    sg.GetStartingControlPointPosition x1, y1
                    sg.GetEndingControlPointPosition x2, y2
                    x3 = sg.EndNode.PositionX: y3 = sg.EndNode.PositionY
                    ' вторая разность контрольного многоугольника -> число сегментов по допуску
                    dd = Sqr((x0 - 2 * x1 + x2) ^ 2 + (y0 - 2 * y1 + y2) ^ 2)
                    d2 = Sqr((x1 - 2 * x2 + x3) ^ 2 + (y1 - 2 * y2 + y3) ^ 2)
                    If d2 > dd Then
                        dd = d2
                    End If
                    n = -Int(-Sqr(0.75 * dd / P_FLAT_TOL))
                    If n < 2 Then
                        n = 2
                    End If
                    If n > 256 Then
                        n = 256
                    End If
                    For k = 1 To n - 1
                        t = k / n: mt = 1 - t
                        CoreAddPoint mt * mt * mt * x0 + 3 * mt * mt * t * x1 + 3 * mt * t * t * x2 + t * t * t * x3, _
                                     mt * mt * mt * y0 + 3 * mt * mt * t * y1 + 3 * mt * t * t * y2 + t * t * t * y3
                    Next k
                End If
            Next sg
            CoreEndRing st
        End If
    Next sp
End Sub

' одна кривая на группу (букву / связку букв); возвращает число созданных кривых
Private Function EmitShapes(lay As Layer) As Long
    Dim crv As Curve, sp As SubPath, p As Long, s0 As Long, s1 As Long, i As Long, shp As Shape
    Dim g As Long, nSub As Long, nMade As Long
    For g = 0 To CoreGroupCount() - 1
        Set crv = Application.CreateCurve(ActiveDocument)
        nSub = 0
        For p = 0 To CorePathCount() - 1
            If CorePathGroup(p) = g Then
                s0 = CorePathStart(p)
                If p < CorePathCount() - 1 Then
                    s1 = CorePathStart(p + 1) - 1
                Else
                    s1 = CoreSegCount() - 1
                End If
                If s1 >= s0 Then
                    Set sp = crv.CreateSubPath(CoreSeg(1, s0), CoreSeg(2, s0))
                    For i = s0 To s1
                        If CoreSeg(0, i) = 0 Then
                            sp.AppendLineSegment CoreSeg(7, i), CoreSeg(8, i)
                        Else
                            sp.AppendCurveSegment2 CoreSeg(7, i), CoreSeg(8, i), CoreSeg(3, i), CoreSeg(4, i), CoreSeg(5, i), CoreSeg(6, i)
                        End If
                    Next i
                    If CorePathClosed(p) Then
                        On Error Resume Next
                        sp.Closed = True
                        On Error GoTo 0
                    End If
                    nSub = nSub + 1
                End If
            End If
        Next p
        If nSub > 0 Then
            Set shp = lay.CreateCurve(crv)
            shp.Fill.ApplyNoFill
            shp.Outline.SetProperties 0.1, Color:=CreateRGBColor(255, 0, 0)
            shp.Name = OutPrefix() & (nMade + 1)
            nMade = nMade + 1
        End If
    Next g
    EmitShapes = nMade
End Function
