Attribute VB_Name = "modAlign"
'Preview image analysis and alignment (port of the Python Aligner).
'G maps normalized world (x,y) to preview pixels; preview arrays are 0-based, index = y * PW + x.
Option Explicit

Public PW As Long, PH As Long
Public Gray() As Byte
Public Region() As Byte
Public HasRegion As Boolean
Private EdgeD() As Single, HasEdge As Boolean
Private InkD() As Single, HasInk As Boolean

'objective state
Private oX() As Double, oY() As Double, oW() As Double, oCnt As Long, oWSum As Double
Private oMap As Long          ' 1 = edge map, 2 = ink map
Private oCap As Single
Private oMode As Long         ' 1 = similarity params, 2 = affine params
Private nmTick As Long

Public Sub ResetAnalysis()
    HasRegion = False: HasEdge = False: HasInk = False
    Erase Region: Erase EdgeD: Erase InkD
End Sub

'---------------------------------------------------------------- morphology helpers
Private Sub LongLineMask(m() As Byte, ByVal W As Long, ByVal H As Long, ByVal L As Long, outM() As Byte)
    Dim x As Long, y As Long, run As Long, k As Long, base As Long
    ReDim outM(0 To W * H - 1)
    For y = 0 To H - 1
        base = y * W
        run = 0
        For x = 0 To W
            If x < W Then
                If m(base + x) Then run = run + 1: GoTo nextX
            End If
            If run >= L Then
                For k = x - run To x - 1
                    outM(base + k) = 1
                Next
            End If
            run = 0
nextX:
        Next
    Next
    For x = 0 To W - 1
        run = 0
        For y = 0 To H
            If y < H Then
                If m(y * W + x) Then run = run + 1: GoTo nextY
            End If
            If run >= L Then
                For k = y - run To y - 1
                    outM(k * W + x) = 1
                Next
            End If
            run = 0
nextY:
        Next
    Next
End Sub

Private Function Dilate(a() As Byte, ByVal W As Long, ByVal H As Long, ByVal conn8 As Boolean) As Byte()
    Dim m() As Byte, o() As Byte, x As Long, y As Long, i As Long
    ReDim m(0 To W * H - 1)
    For y = 0 To H - 1
        For x = 0 To W - 1
            i = y * W + x
            If a(i) Then
                m(i) = 1
            ElseIf y > 0 Then
                If a(i - W) Then m(i) = 1
            End If
            If m(i) = 0 And y < H - 1 Then
                If a(i + W) Then m(i) = 1
            End If
        Next
    Next
    ReDim o(0 To W * H - 1)
    Dim src As Variant
    For y = 0 To H - 1
        For x = 0 To W - 1
            i = y * W + x
            If conn8 Then
                If m(i) Then
                    o(i) = 1
                ElseIf x > 0 Then
                    If m(i - 1) Then o(i) = 1
                End If
                If o(i) = 0 And x < W - 1 Then
                    If m(i + 1) Then o(i) = 1
                End If
            Else
                If m(i) Then
                    o(i) = 1
                ElseIf x > 0 Then
                    If a(i - 1) Then o(i) = 1
                End If
                If o(i) = 0 And x < W - 1 Then
                    If a(i + 1) Then o(i) = 1
                End If
            End If
        Next
    Next
    Dilate = o
End Function

Private Function Invert(a() As Byte) As Byte()
    Dim o() As Byte, i As Long
    o = a
    For i = 0 To UBound(o)
        o(i) = 1 - o(i)
    Next
    Invert = o
End Function

Private Function Erode(a() As Byte, ByVal W As Long, ByVal H As Long, ByVal conn8 As Boolean) As Byte()
    Erode = Invert(Dilate(Invert(a), W, H, conn8))
End Function

'Two-pass chamfer distance (1 / 1.4142) to the nearest set pixel, truncated at cap.
Private Sub DistanceMap(mask() As Byte, ByVal W As Long, ByVal H As Long, ByVal cap As Single, d() As Single)
    Dim x As Long, y As Long, i As Long, v As Single, t As Single
    Const DG As Single = 1.4142
    ReDim d(0 To W * H - 1)
    For i = 0 To W * H - 1
        If mask(i) Then d(i) = 0 Else d(i) = cap
    Next
    For y = 0 To H - 1
        For x = 0 To W - 1
            i = y * W + x
            v = d(i)
            If v > 0 Then
                If x > 0 Then t = d(i - 1) + 1: If t < v Then v = t
                If y > 0 Then
                    t = d(i - W) + 1: If t < v Then v = t
                    If x > 0 Then t = d(i - W - 1) + DG: If t < v Then v = t
                    If x < W - 1 Then t = d(i - W + 1) + DG: If t < v Then v = t
                End If
                d(i) = v
            End If
        Next
    Next
    For y = H - 1 To 0 Step -1
        For x = W - 1 To 0 Step -1
            i = y * W + x
            v = d(i)
            If v > 0 Then
                If x < W - 1 Then t = d(i + 1) + 1: If t < v Then v = t
                If y < H - 1 Then
                    t = d(i + W) + 1: If t < v Then v = t
                    If x < W - 1 Then t = d(i + W + 1) + DG: If t < v Then v = t
                    If x > 0 Then t = d(i + W - 1) + DG: If t < v Then v = t
                End If
                If v > cap Then v = cap
                d(i) = v
            End If
        Next
    Next
End Sub

'---------------------------------------------------------------- map region detection
Public Sub EnsureRegion()
    If HasRegion Then Exit Sub
    Dim content() As Byte, lines() As Byte, i As Long, L As Long
    ReDim content(0 To PW * PH - 1)
    For i = 0 To PW * PH - 1
        If Gray(i) < 235 Then content(i) = 1
    Next
    L = Int(PREVIEW_DPI * 1.1): If L < 20 Then L = 20
    LongLineMask content, PW, PH, L, lines
    For i = 0 To PW * PH - 1
        If lines(i) Then content(i) = 0
    Next
    Erase lines
    Dim f As Long, W2 As Long, H2 As Long, x As Long, y As Long, cx As Long, cy As Long, s As Long
    f = CLng(PREVIEW_DPI / 12#): If f < 1 Then f = 1
    W2 = PW \ f: H2 = PH \ f
    Dim dens() As Double, cells() As Byte
    ReDim dens(0 To W2 * H2 - 1): ReDim cells(0 To W2 * H2 - 1)
    For cy = 0 To H2 - 1
        For cx = 0 To W2 - 1
            s = 0
            For y = cy * f To cy * f + f - 1
                For x = cx * f To cx * f + f - 1
                    s = s + content(y * PW + x)
                Next
            Next
            dens(cy * W2 + cx) = s / (f * f)
            If dens(cy * W2 + cx) > 0.06 Then cells(cy * W2 + cx) = 1
        Next
    Next
    Erase content
    Dim dil() As Byte
    dil = Dilate(cells, W2, H2, True)
    dil = Dilate(dil, W2, H2, False)
    ' largest 8-connected component scored by density
    Dim lab() As Long, stack() As Long, sp As Long, cur As Long, best As Long, bestScore As Double, sc As Double
    Dim pIdx As Long, r As Long, c As Long, dr As Long, dc As Long, q As Long
    ReDim lab(0 To W2 * H2 - 1): ReDim stack(0 To W2 * H2)
    bestScore = -1
    For i = 0 To W2 * H2 - 1
        If dil(i) = 1 And lab(i) = 0 Then
            cur = cur + 1
            lab(i) = cur: sp = 0: stack(0) = i: sp = 1: sc = 0
            Do While sp > 0
                sp = sp - 1: pIdx = stack(sp)
                sc = sc + dens(pIdx)
                r = pIdx \ W2: c = pIdx - r * W2
                For dr = -1 To 1
                    If r + dr >= 0 And r + dr < H2 Then
                        For dc = -1 To 1
                            If c + dc >= 0 And c + dc < W2 Then
                                q = (r + dr) * W2 + c + dc
                                If dil(q) = 1 And lab(q) = 0 Then
                                    lab(q) = cur: stack(sp) = q: sp = sp + 1
                                End If
                            End If
                        Next
                    End If
                Next
            Loop
            If sc > bestScore Then bestScore = sc: best = cur
        End If
    Next
    Dim comp() As Byte, closed() As Byte
    ReDim comp(0 To W2 * H2 - 1)
    If best > 0 Then
        For i = 0 To W2 * H2 - 1
            If lab(i) = best Then comp(i) = 1
        Next
    End If
    Erase lab
    ' erode_n(comp, 2): conn8 then conn4, border counts as inside
    closed = ErodeBorderIn(comp, W2, H2, True)
    closed = ErodeBorderIn(closed, W2, H2, False)
    For i = 0 To W2 * H2 - 1
        closed(i) = closed(i) And comp(i)
    Next
    ' fill holes: background reachable from the border (4-connected)
    Dim reach() As Byte
    ReDim reach(0 To W2 * H2 - 1)
    sp = 0
    For i = 0 To W2 * H2 - 1
        r = i \ W2: c = i - r * W2
        If r = 0 Or r = H2 - 1 Or c = 0 Or c = W2 - 1 Then
            If closed(i) = 0 Then reach(i) = 1: stack(sp) = i: sp = sp + 1
        End If
    Next
    Do While sp > 0
        sp = sp - 1: pIdx = stack(sp)
        r = pIdx \ W2: c = pIdx - r * W2
        If r > 0 Then q = pIdx - W2: GoSub Visit
        If r < H2 - 1 Then q = pIdx + W2: GoSub Visit
        If c > 0 Then q = pIdx - 1: GoSub Visit
        If c < W2 - 1 Then q = pIdx + 1: GoSub Visit
    Loop
    ReDim Region(0 To PW * PH - 1)
    For cy = 0 To H2 - 1
        For cx = 0 To W2 - 1
            If reach(cy * W2 + cx) = 0 Then
                For y = cy * f To cy * f + f - 1
                    For x = cx * f To cx * f + f - 1
                        Region(y * PW + x) = 1
                    Next
                Next
            End If
        Next
    Next
    HasRegion = True
    Exit Sub
Visit:
    If closed(q) = 0 And reach(q) = 0 Then reach(q) = 1: stack(sp) = q: sp = sp + 1
    Return
End Sub

Private Function ErodeBorderIn(a() As Byte, ByVal W As Long, ByVal H As Long, ByVal conn8 As Boolean) As Byte()
    ' ~dilate(~a) where out-of-image counts as "a" (so the border does not erode)
    Dim inv() As Byte, x As Long, y As Long, i As Long, o() As Byte, m() As Byte
    inv = Invert(a)
    o = Dilate(inv, W, H, conn8)
    ErodeBorderIn = Invert(o)
End Function

Private Sub EnsureEdge()
    If HasEdge Then Exit Sub
    EnsureRegion
    Dim er() As Byte, edge() As Byte, i As Long
    er = Erode(Region, PW, PH, False)
    ReDim edge(0 To PW * PH - 1)
    For i = 0 To PW * PH - 1
        If Region(i) = 1 And er(i) = 0 Then edge(i) = 1
    Next
    DistanceMap edge, PW, PH, 40, EdgeD
    HasEdge = True
End Sub

Private Sub EnsureInk()
    If HasInk Then Exit Sub
    Dim dark() As Byte, lines() As Byte, i As Long, L As Long
    ReDim dark(0 To PW * PH - 1)
    For i = 0 To PW * PH - 1
        If Gray(i) < 120 Then dark(i) = 1
    Next
    L = Int(PREVIEW_DPI * 1.1): If L < 20 Then L = 20
    LongLineMask dark, PW, PH, L, lines
    For i = 0 To PW * PH - 1
        If lines(i) Then dark(i) = 0
    Next
    DistanceMap dark, PW, PH, 10, InkD
    HasInk = True
End Sub

'---------------------------------------------------------------- fitting
Public Function InitialFit(rings() As TPoly, ByVal nRings As Long, G() As Double) As Boolean
    EnsureRegion
    Dim x As Long, y As Long, px0 As Long, px1 As Long, py0 As Long, py1 As Long, cnt As Long
    px0 = PW: py0 = PH: px1 = -1: py1 = -1
    For y = 0 To PH - 1
        For x = 0 To PW - 1
            If Region(y * PW + x) Then
                cnt = cnt + 1
                If x < px0 Then px0 = x
                If x > px1 Then px1 = x
                If y < py0 Then py0 = y
                If y > py1 Then py1 = y
            End If
        Next
    Next
    If cnt < 50 Then Err.Raise vbObjectError + 700, "Align", "Could not detect the map area in the PDF."
    px1 = px1 + 1: py1 = py1 + 1
    Dim wx0 As Double, wy0 As Double, wx1 As Double, wy1 As Double, s As Double
    PolysBounds rings, nRings, wx0, wy0, wx1, wy1
    s = Sqr(((px1 - px0) / (wx1 - wx0)) * ((py1 - py0) / (wy1 - wy0)))
    ReDim G(0 To 5)
    G(0) = s: G(1) = 0: G(3) = 0: G(4) = -s
    G(2) = (px0 + px1) / 2# - s * (wx0 + wx1) / 2#
    G(5) = (py0 + py1) / 2# + s * (wy0 + wy1) / 2#
    InitialFit = True
End Function

Private Sub SimToG(p() As Double, G() As Double)
    Dim s As Double, c As Double, sn As Double
    s = Exp(p(2)): c = Cos(p(3)): sn = Sin(p(3))
    ReDim G(0 To 5)
    G(0) = s * c: G(1) = s * sn: G(2) = p(0)
    G(3) = s * sn: G(4) = -s * c: G(5) = p(1)
End Sub

Private Sub GToSim(G() As Double, p() As Double)
    Dim h As Double
    ReDim p(0 To 3)
    h = Sqr(G(0) * G(0) + G(1) * G(1))
    If h < 1E-12 Then h = 1E-12
    p(0) = G(2): p(1) = G(5): p(2) = Log(h): p(3) = Atn2(G(1), G(0))
End Sub

Public Function Atn2(ByVal y As Double, ByVal x As Double) As Double
    Const PI As Double = 3.14159265358979
    If x > 0 Then
        Atn2 = Atn(y / x)
    ElseIf x < 0 Then
        If y >= 0 Then Atn2 = Atn(y / x) + PI Else Atn2 = Atn(y / x) - PI
    Else
        Atn2 = PI / 2 * Sgn(y)
    End If
End Function

Private Function EvalG(G() As Double) As Double
    Dim i As Long, x As Double, y As Double, x0 As Long, y0 As Long, fx As Double, fy As Double
    Dim v As Double, s As Double, k As Long, useW As Boolean
    useW = (oWSum > 0)
    For i = 0 To oCnt - 1
        x = G(0) * oX(i) + G(1) * oY(i) + G(2) - 0.5
        y = G(3) * oX(i) + G(4) * oY(i) + G(5) - 0.5
        If x >= 0 And y >= 0 And x < PW - 1 And y < PH - 1 Then
            x0 = Int(x): y0 = Int(y): fx = x - x0: fy = y - y0
            k = y0 * PW + x0
            If oMap = 1 Then
                v = EdgeD(k) * (1 - fx) * (1 - fy) + EdgeD(k + 1) * fx * (1 - fy) + _
                    EdgeD(k + PW) * (1 - fx) * fy + EdgeD(k + PW + 1) * fx * fy
            Else
                v = InkD(k) * (1 - fx) * (1 - fy) + InkD(k + 1) * fx * (1 - fy) + _
                    InkD(k + PW) * (1 - fx) * fy + InkD(k + PW + 1) * fx * fy
            End If
        Else
            v = oCap
        End If
        If useW Then s = s + v * oW(i) Else s = s + v
    Next
    If useW Then EvalG = s / oWSum Else EvalG = s / IIf(oCnt > 0, oCnt, 1)
End Function

Private Function Objective(p() As Double) As Double
    Dim G() As Double
    If oMode = 1 Then
        SimToG p, G
    Else
        ReDim G(0 To 5)
        G(0) = p(0): G(1) = p(1): G(2) = p(4): G(3) = p(2): G(4) = p(3): G(5) = p(5)
    End If
    Objective = EvalG(G)
End Function

Private Sub NelderMead(x0() As Double, steps As Variant, ByVal iters As Long, xBest() As Double)
    Dim n As Long, i As Long, j As Long, it As Long, pts() As Double, vals() As Double
    Dim c() As Double, xr() As Double, xe() As Double, xc() As Double, tmp() As Double
    Dim fr As Double, fe As Double, fc As Double, ord As Long, t As Double
    n = UBound(x0) + 1
    ReDim pts(0 To n, 0 To n - 1): ReDim vals(0 To n)
    ReDim c(0 To n - 1): ReDim xr(0 To n - 1): ReDim xe(0 To n - 1): ReDim xc(0 To n - 1): ReDim tmp(0 To n - 1)
    For i = 0 To n
        For j = 0 To n - 1
            pts(i, j) = x0(j)
        Next
        If i > 0 Then pts(i, i - 1) = pts(i, i - 1) + steps(i - 1)
        For j = 0 To n - 1: tmp(j) = pts(i, j): Next
        vals(i) = Objective(tmp)
    Next
    For it = 1 To iters
        ' insertion sort by value
        For i = 1 To n
            t = vals(i)
            For j = 0 To n - 1: tmp(j) = pts(i, j): Next
            ord = i - 1
            Do While ord >= 0
                If vals(ord) <= t Then Exit Do
                vals(ord + 1) = vals(ord)
                For j = 0 To n - 1: pts(ord + 1, j) = pts(ord, j): Next
                ord = ord - 1
            Loop
            vals(ord + 1) = t
            For j = 0 To n - 1: pts(ord + 1, j) = tmp(j): Next
        Next
        If Abs(vals(n) - vals(0)) <= 0.0000001 * (Abs(vals(0)) + 0.000000001) Then Exit For
        nmTick = nmTick + 1
        If nmTick Mod 25 = 0 Then DoEvents
        For j = 0 To n - 1
            c(j) = 0
            For i = 0 To n - 1
                c(j) = c(j) + pts(i, j)
            Next
            c(j) = c(j) / n
            xr(j) = c(j) + (c(j) - pts(n, j))
        Next
        fr = Objective(xr)
        If fr < vals(0) Then
            For j = 0 To n - 1: xe(j) = c(j) + 2# * (c(j) - pts(n, j)): Next
            fe = Objective(xe)
            If fe < fr Then
                For j = 0 To n - 1: pts(n, j) = xe(j): Next
                vals(n) = fe
            Else
                For j = 0 To n - 1: pts(n, j) = xr(j): Next
                vals(n) = fr
            End If
        ElseIf fr < vals(n - 1) Then
            For j = 0 To n - 1: pts(n, j) = xr(j): Next
            vals(n) = fr
        Else
            If fr >= vals(n) Then
                For j = 0 To n - 1: xc(j) = c(j) + 0.5 * (pts(n, j) - c(j)): Next
            Else
                For j = 0 To n - 1: xc(j) = c(j) + 0.5 * (xr(j) - c(j)): Next
            End If
            fc = Objective(xc)
            If fc < IIf(fr < vals(n), fr, vals(n)) Then
                For j = 0 To n - 1: pts(n, j) = xc(j): Next
                vals(n) = fc
            Else
                For i = 1 To n
                    For j = 0 To n - 1
                        pts(i, j) = pts(0, j) + 0.5 * (pts(i, j) - pts(0, j))
                        tmp(j) = pts(i, j)
                    Next
                    vals(i) = Objective(tmp)
                Next
            End If
        End If
    Next
    ord = 0
    For i = 1 To n
        If vals(i) < vals(ord) Then ord = i
    Next
    ReDim xBest(0 To n - 1)
    For j = 0 To n - 1: xBest(j) = pts(ord, j): Next
End Sub

Private Sub SetPoints(rings() As TPoly, ByVal nRings As Long, ByVal spacing As Double, ByVal maxPts As Long)
    oCnt = Densify(rings, nRings, spacing, maxPts, oX, oY)
    oWSum = 0
    Erase oW
End Sub

'Refine G against the map: stage 1 boundary vs. detected region edge, stage 2 boundary (+routes) vs. ink.
Public Sub Refine(G() As Double, rings() As TPoly, ByVal nRings As Long, routesN() As TPoly, ByVal nRoutes As Long, _
                  ByVal affine As Boolean, ByVal stage1 As Boolean)
    Dim s As Double, p() As Double, x() As Double
    s = Sqr(G(0) * G(0) + G(1) * G(1))
    oMode = 1
    If stage1 Then
        SetStatus "Auto-align: matching the county outline..."
        EnsureEdge
        SetPoints rings, nRings, 2# / s, 2500
        oMap = 1: oCap = 40
        GToSim G, p
        NelderMead p, Array(12#, 12#, 0.03, 0.02), 400, x
        NelderMead x, Array(4#, 4#, 0.01, 0.006), 300, p
        SimToG p, G
        s = Sqr(G(0) * G(0) + G(1) * G(1))
    End If
    SetStatus "Auto-align: matching the map linework..."
    EnsureInk
    SetPoints rings, nRings, 1# / s, 4000
    If nRoutes > 0 Then
        Dim rx() As Double, ry() As Double, nr As Long, i As Long, nb As Long, wr As Double
        nr = Densify(routesN, nRoutes, 2# / s, 4000, rx, ry)
        If nr > 0 Then
            nb = oCnt
            ReDim Preserve oX(0 To nb + nr - 1): ReDim Preserve oY(0 To nb + nr - 1)
            ReDim oW(0 To nb + nr - 1)
            wr = 0.5 * nb / nr
            For i = 0 To nb - 1: oW(i) = 1: Next
            For i = 0 To nr - 1
                oX(nb + i) = rx(i): oY(nb + i) = ry(i): oW(nb + i) = wr
            Next
            oCnt = nb + nr
            oWSum = nb + nr * wr
        End If
    End If
    oMap = 2: oCap = 10
    GToSim G, p
    NelderMead p, Array(3#, 3#, 0.006, 0.004), 400, x
    NelderMead x, Array(1#, 1#, 0.002, 0.0015), 300, p
    SimToG p, G
    If affine Then
        SetStatus "Auto-align: affine refinement..."
        oMode = 2
        s = Sqr(G(0) * G(0) + G(1) * G(1))
        ReDim p(0 To 5)
        p(0) = G(0): p(1) = G(1): p(2) = G(3): p(3) = G(4): p(4) = G(2): p(5) = G(5)
        NelderMead p, Array(0.004 * s, 0.004 * s, 0.004 * s, 0.004 * s, 1.5, 1.5), 900, x
        G(0) = x(0): G(1) = x(1): G(3) = x(2): G(4) = x(3): G(2) = x(4): G(5) = x(5)
    End If
End Sub

'Mean distance (px) of the boundary to map ink and the fraction within 2 px.
Public Sub ScoreFit(G() As Double, rings() As TPoly, ByVal nRings As Long, meanD As Double, frac As Double)
    Dim s As Double, i As Long, x As Double, y As Double, x0 As Long, y0 As Long, fx As Double, fy As Double
    Dim v As Double, k As Long, sum As Double, good As Long, px() As Double, py() As Double, n As Long
    EnsureInk
    s = Sqr(G(0) * G(0) + G(1) * G(1))
    n = Densify(rings, nRings, 1# / s, 8000, px, py)
    If n = 0 Then meanD = 10: frac = 0: Exit Sub
    For i = 0 To n - 1
        x = G(0) * px(i) + G(1) * py(i) + G(2) - 0.5
        y = G(3) * px(i) + G(4) * py(i) + G(5) - 0.5
        If x >= 0 And y >= 0 And x < PW - 1 And y < PH - 1 Then
            x0 = Int(x): y0 = Int(y): fx = x - x0: fy = y - y0
            k = y0 * PW + x0
            v = InkD(k) * (1 - fx) * (1 - fy) + InkD(k + 1) * fx * (1 - fy) + _
                InkD(k + PW) * (1 - fx) * fy + InkD(k + PW + 1) * fx * fy
        Else
            v = 10
        End If
        sum = sum + v
        If v <= 2 Then good = good + 1
    Next
    meanD = sum / n
    frac = good / n
End Sub
