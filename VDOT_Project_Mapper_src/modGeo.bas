Attribute VB_Name = "modGeo"
'Geometry helpers: polylines, ring split, even-odd clipping, densify, tie-point solve, scale bar sizing.
Option Explicit

Public Type TPoly
    N As Long
    X() As Double
    Y() As Double
    Cat As String
    Nm As String
End Type

Public Sub PolyInit(p As TPoly, ByVal cap As Long)
    If cap < 2 Then cap = 2
    p.N = 0
    ReDim p.X(0 To cap - 1)
    ReDim p.Y(0 To cap - 1)
End Sub

Public Sub PolyAdd(p As TPoly, ByVal x As Double, ByVal y As Double)
    If p.N > UBound(p.X) Then
        ReDim Preserve p.X(0 To 2 * p.N + 1)
        ReDim Preserve p.Y(0 To 2 * p.N + 1)
    End If
    p.X(p.N) = x
    p.Y(p.N) = y
    p.N = p.N + 1
End Sub

Public Function PolyArea(p As TPoly) As Double
    Dim i As Long, j As Long, s As Double
    For i = 0 To p.N - 1
        j = i + 1
        If j = p.N Then j = 0
        s = s + p.X(i) * p.Y(j) - p.X(j) * p.Y(i)
    Next
    PolyArea = s / 2#
End Function

Public Sub PolysBounds(polys() As TPoly, ByVal cnt As Long, x0 As Double, y0 As Double, x1 As Double, y1 As Double)
    Dim i As Long, k As Long
    x0 = 1E+300: y0 = 1E+300: x1 = -1E+300: y1 = -1E+300
    For k = 0 To cnt - 1
        For i = 0 To polys(k).N - 1
            If polys(k).X(i) < x0 Then x0 = polys(k).X(i)
            If polys(k).X(i) > x1 Then x1 = polys(k).X(i)
            If polys(k).Y(i) < y0 Then y0 = polys(k).Y(i)
            If polys(k).Y(i) > y1 Then y1 = polys(k).Y(i)
        Next
    Next
End Sub

'Esri rings: outer rings are clockwise (negative signed area).
Public Sub SplitRings(rings() As TPoly, ByVal cnt As Long, outers() As TPoly, nO As Long, holes() As TPoly, nH As Long)
    Dim i As Long
    nO = 0: nH = 0
    ReDim outers(0 To cnt): ReDim holes(0 To cnt)
    For i = 0 To cnt - 1
        If PolyArea(rings(i)) < 0 Then
            outers(nO) = rings(i): nO = nO + 1
        Else
            holes(nH) = rings(i): nH = nH + 1
        End If
    Next
    If nO = 0 Then
        For i = 0 To nH - 1
            outers(i) = holes(i)
        Next
        nO = nH: nH = 0
    End If
End Sub

Public Function PointInRings(ByVal x As Double, ByVal y As Double, rings() As TPoly, ByVal cnt As Long) As Boolean
    Dim k As Long, i As Long, j As Long, inside As Boolean
    For k = 0 To cnt - 1
        With rings(k)
            j = .N - 1
            For i = 0 To .N - 1
                If (.Y(i) > y) <> (.Y(j) > y) Then
                    If x < (.X(j) - .X(i)) * (y - .Y(i)) / (.Y(j) - .Y(i)) + .X(i) Then inside = Not inside
                End If
                j = i
            Next
        End With
    Next
    PointInRings = inside
End Function

'Clip open polylines to the even-odd area of rings (keeps the parts inside).
Public Sub ClipRoutes(routes() As TPoly, ByVal nR As Long, rings() As TPoly, ByVal nRings As Long, _
                      outP() As TPoly, nOut As Long)
    Dim ex0() As Double, ey0() As Double, ex1() As Double, ey1() As Double, nE As Long
    Dim k As Long, i As Long, j As Long
    For k = 0 To nRings - 1
        nE = nE + rings(k).N
    Next
    nOut = 0
    ReDim outP(0 To 15)
    If nE = 0 Then Exit Sub
    ReDim ex0(0 To nE - 1): ReDim ey0(0 To nE - 1): ReDim ex1(0 To nE - 1): ReDim ey1(0 To nE - 1)
    nE = 0
    For k = 0 To nRings - 1
        With rings(k)
            j = .N - 1
            For i = 0 To .N - 1
                ex0(nE) = .X(j): ey0(nE) = .Y(j): ex1(nE) = .X(i): ey1(nE) = .Y(i)
                nE = nE + 1
                j = i
            Next
        End With
    Next
    ' uniform grid index over ring edges
    Dim gx0 As Double, gy0 As Double, gx1 As Double, gy1 As Double, GN As Long, cw As Double, ch As Double
    PolysBounds rings, nRings, gx0, gy0, gx1, gy1
    GN = 96
    cw = (gx1 - gx0) / GN: ch = (gy1 - gy0) / GN
    If cw <= 0 Then cw = 1E-09
    If ch <= 0 Then ch = 1E-09
    Dim cellCnt() As Long, cellStart() As Long, cellItems() As Long, e As Long
    Dim cxa As Long, cxb As Long, cya As Long, cyb As Long, cx As Long, cy As Long, pass As Long, tot As Long
    ReDim cellCnt(0 To GN * GN - 1): ReDim cellStart(0 To GN * GN)
    For pass = 0 To 1
        For e = 0 To nE - 1
            cxa = Cell(IIf(ex0(e) < ex1(e), ex0(e), ex1(e)), gx0, cw, GN)
            cxb = Cell(IIf(ex0(e) > ex1(e), ex0(e), ex1(e)), gx0, cw, GN)
            cya = Cell(IIf(ey0(e) < ey1(e), ey0(e), ey1(e)), gy0, ch, GN)
            cyb = Cell(IIf(ey0(e) > ey1(e), ey0(e), ey1(e)), gy0, ch, GN)
            For cy = cya To cyb
                For cx = cxa To cxb
                    If pass = 0 Then
                        cellCnt(cy * GN + cx) = cellCnt(cy * GN + cx) + 1
                    Else
                        cellItems(cellStart(cy * GN + cx) + cellCnt(cy * GN + cx)) = e
                        cellCnt(cy * GN + cx) = cellCnt(cy * GN + cx) + 1
                    End If
                Next
            Next
        Next
        If pass = 0 Then
            tot = 0
            For i = 0 To GN * GN - 1
                cellStart(i) = tot
                tot = tot + cellCnt(i)
                cellCnt(i) = 0
            Next
            cellStart(GN * GN) = tot
            ReDim cellItems(0 To IIf(tot > 0, tot - 1, 0))
        End If
    Next
    Dim stamp() As Long, sid As Long
    ReDim stamp(0 To nE - 1)
    Dim ts() As Double, nT As Long, inside As Boolean, part As TPoly
    Dim ax As Double, ay As Double, bx As Double, by As Double, dx As Double, dy As Double
    Dim den As Double, t As Double, u As Double, a As Long, c As Long, tmp As Double
    ReDim ts(0 To 63)
    For k = 0 To nR - 1
        If k Mod 200 = 0 Then DoEvents
        With routes(k)
            If .N >= 2 Then
                inside = PointInRings(.X(0), .Y(0), rings, nRings)
                PolyInit part, 16
                If inside Then PolyAdd part, .X(0), .Y(0)
                For i = 0 To .N - 2
                    ax = .X(i): ay = .Y(i): bx = .X(i + 1): by = .Y(i + 1)
                    dx = bx - ax: dy = by - ay
                    sid = sid + 1
                    nT = 0
                    cxa = Cell(IIf(ax < bx, ax, bx), gx0, cw, GN): cxb = Cell(IIf(ax > bx, ax, bx), gx0, cw, GN)
                    cya = Cell(IIf(ay < by, ay, by), gy0, ch, GN): cyb = Cell(IIf(ay > by, ay, by), gy0, ch, GN)
                    If Not ((IIf(ax > bx, ax, bx) < gx0) Or (IIf(ax < bx, ax, bx) > gx1) Or _
                            (IIf(ay > by, ay, by) < gy0) Or (IIf(ay < by, ay, by) > gy1)) Then
                        For cy = cya To cyb
                            For cx = cxa To cxb
                                For a = cellStart(cy * GN + cx) To cellStart(cy * GN + cx + 1) - 1
                                    e = cellItems(a)
                                    If stamp(e) <> sid Then
                                        stamp(e) = sid
                                        den = dx * (ey1(e) - ey0(e)) - dy * (ex1(e) - ex0(e))
                                        If den <> 0 Then
                                            t = ((ex0(e) - ax) * (ey1(e) - ey0(e)) - (ey0(e) - ay) * (ex1(e) - ex0(e))) / den
                                            u = ((ex0(e) - ax) * dy - (ey0(e) - ay) * dx) / den
                                            If t > 0 And t <= 1 And u >= 0 And u < 1 Then
                                                If nT > UBound(ts) Then ReDim Preserve ts(0 To 2 * nT)
                                                ts(nT) = t: nT = nT + 1
                                            End If
                                        End If
                                    End If
                                Next
                            Next
                        Next
                    End If
                    For a = 1 To nT - 1
                        tmp = ts(a): c = a - 1
                        Do While c >= 0
                            If ts(c) <= tmp Then Exit Do
                            ts(c + 1) = ts(c): c = c - 1
                        Loop
                        ts(c + 1) = tmp
                    Next
                    For a = 0 To nT - 1
                        If inside Then
                            PolyAdd part, ax + ts(a) * dx, ay + ts(a) * dy
                            If part.N >= 2 Then
                                part.Cat = .Cat: part.Nm = .Nm
                                AppendPoly outP, nOut, part
                            End If
                            PolyInit part, 16
                        Else
                            PolyInit part, 16
                            PolyAdd part, ax + ts(a) * dx, ay + ts(a) * dy
                        End If
                        inside = Not inside
                    Next
                    If inside Then
                        If part.N > 0 Then
                            If part.X(part.N - 1) <> bx Or part.Y(part.N - 1) <> by Then PolyAdd part, bx, by
                        Else
                            PolyAdd part, bx, by
                        End If
                    End If
                Next
                If inside And part.N >= 2 Then
                    part.Cat = .Cat: part.Nm = .Nm
                    AppendPoly outP, nOut, part
                End If
            End If
        End With
    Next
End Sub

Private Function Cell(ByVal v As Double, ByVal v0 As Double, ByVal sz As Double, ByVal GN As Long) As Long
    Dim c As Long
    c = Int((v - v0) / sz)
    If c < 0 Then c = 0
    If c > GN - 1 Then c = GN - 1
    Cell = c
End Function

Public Sub AppendPoly(arr() As TPoly, cnt As Long, p As TPoly)
    If cnt > UBound(arr) Then ReDim Preserve arr(0 To 2 * cnt + 1)
    arr(cnt) = p
    ReDim Preserve arr(cnt).X(0 To p.N - 1)
    ReDim Preserve arr(cnt).Y(0 To p.N - 1)
    cnt = cnt + 1
End Sub

'Resample polylines with roughly `spacing` between points; result capped to maxPts (evenly thinned).
Public Function Densify(polys() As TPoly, ByVal cnt As Long, ByVal spacing As Double, ByVal maxPts As Long, _
                        ox() As Double, oy() As Double) As Long
    Dim k As Long, i As Long, j As Long, m As Long, nOut As Long, L As Double, dx As Double, dy As Double
    Dim tx() As Double, ty() As Double
    ReDim tx(0 To 1023): ReDim ty(0 To 1023)
    If spacing <= 0 Then spacing = 1E-09
    For k = 0 To cnt - 1
        With polys(k)
            If .N >= 2 Then
                For i = 0 To .N - 2
                    dx = .X(i + 1) - .X(i): dy = .Y(i + 1) - .Y(i)
                    L = Sqr(dx * dx + dy * dy)
                    m = Int(L / spacing)
                    If m < 1 Then m = 1
                    If m > 100000 Then m = 100000
                    For j = 0 To m - 1
                        If nOut > UBound(tx) Then
                            ReDim Preserve tx(0 To 2 * nOut + 1): ReDim Preserve ty(0 To 2 * nOut + 1)
                        End If
                        tx(nOut) = .X(i) + dx * j / m
                        ty(nOut) = .Y(i) + dy * j / m
                        nOut = nOut + 1
                    Next
                Next
                If nOut > UBound(tx) Then
                    ReDim Preserve tx(0 To 2 * nOut + 1): ReDim Preserve ty(0 To 2 * nOut + 1)
                End If
                tx(nOut) = .X(.N - 1): ty(nOut) = .Y(.N - 1)
                nOut = nOut + 1
            End If
        End With
    Next
    If nOut = 0 Then
        Densify = 0
        Exit Function
    End If
    If nOut > maxPts Then
        ReDim ox(0 To maxPts - 1): ReDim oy(0 To maxPts - 1)
        For i = 0 To maxPts - 1
            j = Int(i * (nOut - 1) / (maxPts - 1))
            ox(i) = tx(j): oy(i) = ty(j)
        Next
        Densify = maxPts
    Else
        ReDim Preserve tx(0 To nOut - 1): ReDim Preserve ty(0 To nOut - 1)
        ox = tx: oy = ty
        Densify = nOut
    End If
End Function

'---------------------------------------------------------------- linear algebra
'G(0..5) = [[g0, g1, g2], [g3, g4, g5]]  (normalized world -> preview pixel)
Public Sub ApplyG(G() As Double, ByVal x As Double, ByVal y As Double, px As Double, py As Double)
    px = G(0) * x + G(1) * y + G(2)
    py = G(3) * x + G(4) * y + G(5)
End Sub

Public Function InvertG(G() As Double, inv() As Double) As Boolean
    Dim det As Double
    ReDim inv(0 To 5)
    det = G(0) * G(4) - G(1) * G(3)
    If Abs(det) < 1E-300 Then Exit Function
    inv(0) = G(4) / det: inv(1) = -G(1) / det
    inv(3) = -G(3) / det: inv(4) = G(0) / det
    inv(2) = -(inv(0) * G(2) + inv(1) * G(5))
    inv(5) = -(inv(3) * G(2) + inv(4) * G(5))
    InvertG = True
End Function

'Solve A x = b (n x n, row-major in a(0..n*n-1)) by Gaussian elimination with partial pivoting.
Public Function SolveLinear(a() As Double, b() As Double, ByVal n As Long, x() As Double) As Boolean
    Dim i As Long, j As Long, k As Long, piv As Long, mx As Double, f As Double, t As Double
    Dim m() As Double, r() As Double
    m = a: r = b
    For k = 0 To n - 1
        piv = k: mx = Abs(m(k * n + k))
        For i = k + 1 To n - 1
            If Abs(m(i * n + k)) > mx Then mx = Abs(m(i * n + k)): piv = i
        Next
        If mx < 1E-300 Then Exit Function
        If piv <> k Then
            For j = 0 To n - 1
                t = m(k * n + j): m(k * n + j) = m(piv * n + j): m(piv * n + j) = t
            Next
            t = r(k): r(k) = r(piv): r(piv) = t
        End If
        For i = k + 1 To n - 1
            f = m(i * n + k) / m(k * n + k)
            For j = k To n - 1
                m(i * n + j) = m(i * n + j) - f * m(k * n + j)
            Next
            r(i) = r(i) - f * r(k)
        Next
    Next
    ReDim x(0 To n - 1)
    For i = n - 1 To 0 Step -1
        t = r(i)
        For j = i + 1 To n - 1
            t = t - m(i * n + j) * x(j)
        Next
        x(i) = t / m(i * n + i)
    Next
    SolveLinear = True
End Function

'Least-squares G from tie-point pairs (normalized world -> pixel). Returns RMS (px) or -1.
Public Function SolveFromPairs(wx() As Double, wy() As Double, px() As Double, py() As Double, ByVal n As Long, _
                               ByVal affine As Boolean, G() As Double) As Double
    Dim a() As Double, b() As Double, x() As Double, i As Long, r As Long, c As Long
    Dim row(0 To 3) As Double, rhs As Double, k As Long
    ReDim G(0 To 5)
    SolveFromPairs = -1
    If affine And n >= 3 Then
        Dim bx(0 To 2) As Double, by(0 To 2) As Double, sx() As Double, sy() As Double
        ReDim a(0 To 8)
        For i = 0 To n - 1
            row(0) = wx(i): row(1) = wy(i): row(2) = 1
            For r = 0 To 2
                For c = 0 To 2
                    a(r * 3 + c) = a(r * 3 + c) + row(r) * row(c)
                Next
                bx(r) = bx(r) + row(r) * px(i)
                by(r) = by(r) + row(r) * py(i)
            Next
        Next
        ReDim b(0 To 2)
        For r = 0 To 2: b(r) = bx(r): Next
        If Not SolveLinear(a, b, 3, sx) Then Exit Function
        For r = 0 To 2: b(r) = by(r): Next
        If Not SolveLinear(a, b, 3, sy) Then Exit Function
        G(0) = sx(0): G(1) = sx(1): G(2) = sx(2)
        G(3) = sy(0): G(4) = sy(1): G(5) = sy(2)
    ElseIf n >= 2 Then
        ' G = [[a, b, tx], [b, -a, ty]]; rows [wx, wy, 1, 0] -> px and [-wy, wx, 0, 1] -> py
        ReDim a(0 To 15): ReDim b(0 To 3)
        For i = 0 To n - 1
            For k = 0 To 1
                If k = 0 Then
                    row(0) = wx(i): row(1) = wy(i): row(2) = 1: row(3) = 0: rhs = px(i)
                Else
                    row(0) = -wy(i): row(1) = wx(i): row(2) = 0: row(3) = 1: rhs = py(i)
                End If
                For r = 0 To 3
                    For c = 0 To 3
                        a(r * 4 + c) = a(r * 4 + c) + row(r) * row(c)
                    Next
                    b(r) = b(r) + row(r) * rhs
                Next
            Next
        Next
        If Not SolveLinear(a, b, 4, x) Then Exit Function
        G(0) = x(0): G(1) = x(1): G(2) = x(2)
        G(3) = x(1): G(4) = -x(0): G(5) = x(3)
    Else
        Exit Function
    End If
    Dim s As Double, ex As Double, ey As Double
    For i = 0 To n - 1
        ApplyG G, wx(i), wy(i), ex, ey
        s = s + (ex - px(i)) ^ 2 + (ey - py(i)) ^ 2
    Next
    SolveFromPairs = Sqr(s / n)
End Function

'---------------------------------------------------------------- scale bar sizing
Public Function NiceFmt(ByVal v As Double) As String
    Dim s As String
    s = Format$(v, "0.000")
    s = Replace(s, ",", ".")
    Do While Right$(s, 1) = "0"
        s = Left$(s, Len(s) - 1)
    Loop
    If Right$(s, 1) = "." Then s = Left$(s, Len(s) - 1)
    NiceFmt = s
End Function

'Round scale-bar length near target (master units). Returns length (master units); labels(0..2) and unit name.
Public Function ScaleBarSpec(ByVal target As Double, ByVal masterUnit As String, labels() As String, _
                             unitName As String) As Double
    Dim u As String, factor As Double, t As Double, mag As Double, nice As Double, s As Variant
    u = LCase$(Trim$(masterUnit))
    Select Case u
    Case "ft", "'", "feet", "foot", "us survey feet", "sft", "survey feet", "us survey foot"
        If target >= 2640 Then factor = 5280: unitName = "Miles" Else factor = 1: unitName = "Feet"
    Case "m", "meters", "metres", "meter"
        If target >= 1000 Then factor = 1000: unitName = "Kilometers" Else factor = 1: unitName = "Meters"
    Case Else
        factor = 1: unitName = IIf(Len(masterUnit) > 0, masterUnit, "Units")
    End Select
    t = target / factor
    If t < 1E-09 Then t = 1E-09
    mag = 10 ^ Int(Log(t) / Log(10#) + 1E-12)
    nice = mag
    For Each s In Array(1#, 2#, 2.5, 5#)
        If s * mag <= t * 1.0000001 Then nice = s * mag
    Next
    ReDim labels(0 To 2)
    labels(0) = "0": labels(1) = NiceFmt(nice / 2#): labels(2) = NiceFmt(nice)
    ScaleBarSpec = nice * factor
End Function
