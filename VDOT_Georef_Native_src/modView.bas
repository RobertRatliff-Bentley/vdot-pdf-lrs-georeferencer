Attribute VB_Name = "modView"
'Map preview: PDF raster with mip levels, alignment overlay, pins, block fence and mouse tools.
'The viewport is rendered into a BMP that is shown by the "imgView" Image control.
Option Explicit

#If VBA7 Then
Private Declare PtrSafe Function GetDC Lib "user32" (ByVal hwnd As LongPtr) As LongPtr
Private Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hwnd As LongPtr, ByVal hdc As LongPtr) As Long
Private Declare PtrSafe Function GetDeviceCaps Lib "gdi32" (ByVal hdc As LongPtr, ByVal nIndex As Long) As Long
#Else
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
#End If

Public Const TOOL_PAN As Long = 0
Public Const TOOL_ADJUST As Long = 1
Public Const TOOL_TIE As Long = 2
Public Const TOOL_FENCE As Long = 3
Public gTool As Long

Private Type TMip
    W As Long
    H As Long
    K As Double          ' mip pixels per preview pixel
    B() As Byte
End Type
Private mips(1 To 6) As TMip, nMip As Long

Private VW As Long, VH As Long, vz As Double, vox As Double, voy As Double, viewInit As Boolean
Private baseBuf() As Byte, baseOK As Boolean, buf() As Byte, flip As Boolean
Private pendOn As Boolean, pendX As Double, pendY As Double
Private dragMode As Long, lastSX As Double, lastSY As Double, startSX As Double, startSY As Double
Private startVox As Double, startVoy As Double, lastDraw As Long
Private fOn As Boolean, fX0 As Double, fY0 As Double, fX1 As Double, fY1 As Double
Private mPtPx As Double

'---------------------------------------------------------------- PDF preview
Public Sub RenderPreview(ByVal path As String)
    Dim wpt As Double, hpt As Double, i As Long
    SetStatus "Rendering preview" & UEll()
    Erase gDisp
    For i = 1 To 6: Erase mips(i).B: Next
    nMip = 0
    RenderPdfPage path, 96, gDisp, DW, DH, wpt, hpt, 16000000#
    gPdfWpt = wpt: gPdfHpt = hpt
    PW = CLng(wpt * PREVIEW_DPI / 72#): PH = CLng(hpt * PREVIEW_DPI / 72#)
    If PW < 8 Or PH < 8 Then Err.Raise vbObjectError + 800, , "The PDF page is too small."
    gDispK = DW / (wpt * PREVIEW_DPI / 72#)
    Debug.Print "preview: page " & wpt & " x " & hpt & " pt, display " & DW & " x " & DH & " (k=" & gDispK & ")"
    BuildMips
    BuildGray
    ResetAnalysis
    baseOK = False
End Sub

Private Sub BuildMips()
    Dim sw As Long, sh As Long, k As Double
    sw = DW: sh = DH: k = gDispK
    nMip = 0
    Do While (sw > 700 Or sh > 700) And nMip < 6 And sw >= 4 And sh >= 4
        nMip = nMip + 1
        If nMip = 1 Then
            Half gDisp, sw, sh, mips(1).B, mips(1).W, mips(1).H
        Else
            Half mips(nMip - 1).B, sw, sh, mips(nMip).B, mips(nMip).W, mips(nMip).H
        End If
        k = k / 2#
        mips(nMip).K = k
        sw = mips(nMip).W: sh = mips(nMip).H
        If nMip Mod 1 = 0 Then DoEvents
    Loop
End Sub

Private Sub Half(src() As Byte, ByVal sw As Long, ByVal sh As Long, dst() As Byte, dw2 As Long, dh2 As Long)
    Dim x As Long, y As Long, r0 As Long, r1 As Long, o As Long, i As Long, j As Long
    dw2 = sw \ 2: dh2 = sh \ 2
    ReDim dst(0 To dw2 * dh2 * 3 - 1)
    For y = 0 To dh2 - 1
        r0 = 2 * y * sw * 3: r1 = r0 + sw * 3: o = y * dw2 * 3
        For x = 0 To dw2 - 1
            i = r0 + x * 6: j = r1 + x * 6
            dst(o) = (CLng(src(i)) + src(i + 3) + src(j) + src(j + 3) + 2) \ 4
            dst(o + 1) = (CLng(src(i + 1)) + src(i + 4) + src(j + 1) + src(j + 4) + 2) \ 4
            dst(o + 2) = (CLng(src(i + 2)) + src(i + 5) + src(j + 2) + src(j + 5) + 2) \ 4
            o = o + 3
        Next
    Next
End Sub

'Gray preview at PREVIEW_DPI (PW x PH), box-filtered from the closest level with >= 1 px per preview px.
Private Sub BuildGray()
    Dim L As Long, k As Double, x As Long, y As Long, cs() As Long, ce() As Long, rs As Long, re As Long
    Dim sx As Long, sy As Long, s As Long, n As Long, o As Long, sw As Long, sh As Long
    L = 0: k = gDispK
    Do While L < nMip
        If mips(L + 1).K < 1 Then Exit Do
        L = L + 1: k = mips(L).K
    Loop
    If L = 0 Then sw = DW: sh = DH Else sw = mips(L).W: sh = mips(L).H
    ReDim Gray(0 To PW * PH - 1)
    ReDim cs(0 To PW - 1): ReDim ce(0 To PW - 1)
    For x = 0 To PW - 1
        cs(x) = Int(x * k): ce(x) = Int((x + 1) * k) - 1
        If ce(x) < cs(x) Then ce(x) = cs(x)
        If ce(x) > sw - 1 Then ce(x) = sw - 1
        If cs(x) > sw - 1 Then cs(x) = sw - 1
    Next
    For y = 0 To PH - 1
        rs = Int(y * k): re = Int((y + 1) * k) - 1
        If re < rs Then re = rs
        If re > sh - 1 Then re = sh - 1
        If rs > sh - 1 Then rs = sh - 1
        For x = 0 To PW - 1
            s = 0: n = 0
            For sy = rs To re
                o = (sy * sw + cs(x)) * 3
                For sx = cs(x) To ce(x)
                    If L = 0 Then
                        s = s + 114& * gDisp(o) + 587& * gDisp(o + 1) + 299& * gDisp(o + 2)
                    Else
                        s = s + 114& * mips(L).B(o) + 587& * mips(L).B(o + 1) + 299& * mips(L).B(o + 2)
                    End If
                    n = n + 1: o = o + 3
                Next
            Next
            Gray(y * PW + x) = (s \ n + 500) \ 1000
        Next
        If y Mod 200 = 0 Then DoEvents
    Next
End Sub

'---------------------------------------------------------------- viewport geometry
Private Function PtPx() As Double
    Dim hdc As LongPtr
    If mPtPx = 0 Then
        hdc = GetDC(0)
        mPtPx = GetDeviceCaps(hdc, 88) / 72#
        ReleaseDC 0, hdc
        If mPtPx <= 0 Then mPtPx = 96# / 72#
    End If
    PtPx = mPtPx
End Function

Private Function SyncSize() As Boolean
    Dim c As Object, w As Long, h As Long
    Set c = ViewCtl()
    If c Is Nothing Then Exit Function
    w = CLng(c.Width * PtPx()): h = CLng(c.Height * PtPx())
    If w < 16 Then w = 16
    If h < 16 Then h = 16
    If w <> VW Or h <> VH Then
        If viewInit And VW > 0 Then
            vox = vox + (w - VW) / 2#: voy = voy + (h - VH) / 2#
        End If
        VW = w: VH = h
        baseOK = False
    End If
    SyncSize = True
End Function

Public Sub ViewFit()
    Dim pw2 As Double, ph2 As Double, a As Double, b As Double
    If Not SyncSize() Then Exit Sub
    If gHasPdf Then
        pw2 = gPdfWpt * PREVIEW_DPI / 72#: ph2 = gPdfHpt * PREVIEW_DPI / 72#
        a = VW / pw2: b = VH / ph2
        vz = IIf(a < b, a, b) * 0.97
        vox = (VW - pw2 * vz) / 2#: voy = (VH - ph2 * vz) / 2#
        viewInit = True
    End If
    ViewRedraw True
End Sub

Public Sub ViewZoom(ByVal f As Double)
    Dim cx As Double, cy As Double
    If Not gHasPdf Then Exit Sub
    SyncSize
    cx = VW / 2#: cy = VH / 2#
    If vz * f > 40 Or vz * f < 0.01 Then Exit Sub
    vz = vz * f
    vox = cx - (cx - vox) * f: voy = cy - (cy - voy) * f
    ViewRedraw True
End Sub

'---------------------------------------------------------------- drawing
Public Sub ViewRedraw(ByVal full As Boolean)
    If Not FormIsOpen() Then Exit Sub
    If Not SyncSize() Then Exit Sub
    If gHasPdf And Not viewInit Then ViewFit: Exit Sub
    If full Or Not baseOK Then BuildBase
    buf = baseBuf
    DrawOverlay
    Present
    lastDraw = GetTickCount
End Sub

Private Sub SplitRGB(ByVal c As Long, r As Long, g As Long, b As Long)
    r = c And &HFF&: g = (c \ &H100&) And &HFF&: b = (c \ &H10000) And &HFF&
End Sub

Private Sub BuildBase()
    Dim bgR As Long, bgG As Long, bgB As Long, row() As Byte, x As Long, y As Long, i As Long
    SplitRGB PalC("canvas"), bgR, bgG, bgB
    ReDim baseBuf(0 To VW * VH * 3 - 1)
    ReDim row(0 To VW * 3 - 1)
    For x = 0 To VW - 1
        row(x * 3) = bgB: row(x * 3 + 1) = bgG: row(x * 3 + 2) = bgR
    Next
    For y = 0 To VH - 1
        CopyMemory baseBuf(y * VW * 3), row(0), VW * 3
    Next
    baseOK = True
    If Not gHasPdf Then Exit Sub

    Dim L As Long, k As Double, sw As Long, sh As Long, a As Double, useRgn As Boolean
    L = 0: k = gDispK
    Do While L < nMip
        If mips(L + 1).K < vz Then Exit Do
        L = L + 1: k = mips(L).K
    Loop
    If L = 0 Then sw = DW: sh = DH Else sw = mips(L).W: sh = mips(L).H
    a = GetVal("scrPdf") / 100#
    If a <= 0 Or a > 1 Then a = 1
    useRgn = Chk("chkRegion")
    If useRgn And Not HasRegion Then
        SetStatus "Detecting the map area" & UEll()
        EnsureRegion
        SetStatus "Ready"
    End If

    Dim cI() As Long, rI() As Long, pc() As Long, pr() As Long, p As Double
    ReDim cI(0 To VW - 1): ReDim pc(0 To VW - 1)
    ReDim rI(0 To VH - 1): ReDim pr(0 To VH - 1)
    For x = 0 To VW - 1
        p = (x + 0.5 - vox) / vz
        cI(x) = Int(p * k): pc(x) = Int(p)
        If cI(x) < 0 Or cI(x) >= sw Then cI(x) = -1
        If pc(x) < 0 Or pc(x) >= PW Then pc(x) = -1
    Next
    For y = 0 To VH - 1
        p = (y + 0.5 - voy) / vz
        rI(y) = Int(p * k): pr(y) = Int(p)
        If rI(y) < 0 Or rI(y) >= sh Then rI(y) = -1
        If pr(y) < 0 Or pr(y) >= PH Then pr(y) = -1
    Next
    ' look-up tables: opacity blend with the canvas, and region tint (34,211,238 @ 70/255) on top
    Dim lB(0 To 255) As Byte, lG(0 To 255) As Byte, lR(0 To 255) As Byte
    Dim tB(0 To 255) As Byte, tG(0 To 255) As Byte, tR(0 To 255) As Byte, ta As Double
    ta = 70# / 255#
    For i = 0 To 255
        lB(i) = CLng(bgB + (i - bgB) * a): lG(i) = CLng(bgG + (i - bgG) * a): lR(i) = CLng(bgR + (i - bgR) * a)
        tB(i) = CLng(lB(i) * (1 - ta) + 238 * ta): tG(i) = CLng(lG(i) * (1 - ta) + 211 * ta): tR(i) = CLng(lR(i) * (1 - ta) + 34 * ta)
    Next
    Dim o As Long, s As Long, srow As Long, prow As Long, c As Long
    For y = 0 To VH - 1
        If rI(y) >= 0 Then
            srow = rI(y) * sw * 3
            If pr(y) >= 0 Then prow = pr(y) * PW Else prow = -1
            o = y * VW * 3
            For x = 0 To VW - 1
                c = cI(x)
                If c >= 0 Then
                    s = srow + c * 3
                    If useRgn And prow >= 0 And pc(x) >= 0 Then
                        If modAlign.Region(prow + pc(x)) Then
                            If L = 0 Then
                                baseBuf(o) = tB(gDisp(s)): baseBuf(o + 1) = tG(gDisp(s + 1)): baseBuf(o + 2) = tR(gDisp(s + 2))
                            Else
                                baseBuf(o) = tB(mips(L).B(s)): baseBuf(o + 1) = tG(mips(L).B(s + 1)): baseBuf(o + 2) = tR(mips(L).B(s + 2))
                            End If
                            GoTo NextX
                        End If
                    End If
                    If L = 0 Then
                        baseBuf(o) = lB(gDisp(s)): baseBuf(o + 1) = lG(gDisp(s + 1)): baseBuf(o + 2) = lR(gDisp(s + 2))
                    Else
                        baseBuf(o) = lB(mips(L).B(s)): baseBuf(o + 1) = lG(mips(L).B(s + 1)): baseBuf(o + 2) = lR(mips(L).B(s + 2))
                    End If
                End If
NextX:
                o = o + 3
            Next
        End If
    Next
End Sub

Private Sub DrawOverlay()
    Dim k As Long, g As Long, i As Long, r As Long, gg As Long, b As Long
    If gHasG And gHasPdf And Chk("chkOverlay") Then
        If Chk("chkShowRoutes") And gNRt > 0 Then
            For k = 0 To gNRt - 1
                g = gRtGrp(k)
                SplitRGB GrpPrev(g), r, gg, b
                DrawPolyN gRoutesN(k), False, IIf(g <= 2, 2, 1), r, gg, b
            Next
        End If
        For k = 0 To gNR - 1
            DrawPolyN gRingsN(k), True, 5, &HB, &H12, &H20
        Next
        For k = 0 To gNR - 1
            DrawPolyN gRingsN(k), True, 3, &H22, &HD3, &HEE
        Next
    End If
    If gHasPdf Then
        Dim sx As Double, sy As Double, qx As Double, qy As Double
        For i = 0 To gNPairs - 1
            If gHasG Then
                NormToScreen gPairWX(i), gPairWY(i), sx, sy
            Else
                sx = gPairPX(i) * vz + vox: sy = gPairPY(i) * vz + voy
            End If
            qx = gPairPX(i) * vz + vox: qy = gPairPY(i) * vz + voy
            ThickLine sx, sy, qx, qy, 1, &HF4, &H72, &HB6
            Marker sx, sy, &H22, &HD3, &HEE
            Marker qx, qy, &HF4, &H72, &HB6
            DrawNumber i + 1, qx + 5, qy + 4
        Next
        If pendOn And gHasG Then
            NormToScreen pendX, pendY, sx, sy
            Marker sx, sy, &H22, &HD3, &HEE
        End If
        If fOn Then
            DrawFence fX0, fY0, fX1, fY1
        ElseIf gHasFence Then
            DrawFence gFx0, gFy0, gFx1, gFy1
        End If
    End If
End Sub

Private Sub NormToScreen(ByVal x As Double, ByVal y As Double, sx As Double, sy As Double)
    sx = (gG(0) * x + gG(1) * y + gG(2)) * vz + vox
    sy = (gG(3) * x + gG(4) * y + gG(5)) * vz + voy
End Sub

Private Sub DrawPolyN(p As TPoly, ByVal closed As Boolean, ByVal w As Long, ByVal r As Long, ByVal g As Long, ByVal b As Long)
    Dim i As Long, x0 As Double, y0 As Double, x1 As Double, y1 As Double, fx As Double, fy As Double
    If p.N < 2 Then Exit Sub
    NormToScreen p.X(0), p.Y(0), x0, y0
    fx = x0: fy = y0
    For i = 1 To p.N - 1
        NormToScreen p.X(i), p.Y(i), x1, y1
        ThickLine x0, y0, x1, y1, w, r, g, b
        x0 = x1: y0 = y1
    Next
    If closed Then ThickLine x0, y0, fx, fy, w, r, g, b
End Sub

Private Sub Stamp(ByVal cx As Long, ByVal cy As Long, ByVal w As Long, ByVal r As Long, ByVal g As Long, ByVal b As Long)
    Dim x As Long, y As Long, h As Long, o As Long, x0 As Long, x1 As Long, cut As Long
    h = w \ 2
    For y = cy - h To cy - h + w - 1
        If y >= 0 And y < VH Then
            cut = 0
            If w >= 5 Then
                If y = cy - h Or y = cy - h + w - 1 Then cut = 1
            End If
            x0 = cx - h + cut: x1 = cx - h + w - 1 - cut
            If x0 < 0 Then x0 = 0
            If x1 > VW - 1 Then x1 = VW - 1
            o = (y * VW + x0) * 3
            For x = x0 To x1
                buf(o) = b: buf(o + 1) = g: buf(o + 2) = r: o = o + 3
            Next
        End If
    Next
End Sub

Private Sub ThickLine(ByVal x0 As Double, ByVal y0 As Double, ByVal x1 As Double, ByVal y1 As Double, _
                      ByVal w As Long, ByVal r As Long, ByVal g As Long, ByVal b As Long)
    Dim t0 As Double, t1 As Double, dx As Double, dy As Double, n As Long, i As Long, m As Double
    Dim px As Long, py As Long, qx As Long, qy As Long
    ' Liang-Barsky clip to the (padded) viewport
    t0 = 0: t1 = 1: dx = x1 - x0: dy = y1 - y0
    If Not ClipT(-dx, x0 - (-w), t0, t1) Then Exit Sub
    If Not ClipT(dx, (VW + w) - x0, t0, t1) Then Exit Sub
    If Not ClipT(-dy, y0 - (-w), t0, t1) Then Exit Sub
    If Not ClipT(dy, (VH + w) - y0, t0, t1) Then Exit Sub
    x1 = x0 + t1 * dx: y1 = y0 + t1 * dy
    x0 = x0 + t0 * dx: y0 = y0 + t0 * dy
    dx = x1 - x0: dy = y1 - y0
    m = IIf(Abs(dx) > Abs(dy), Abs(dx), Abs(dy))
    n = CLng(m)
    If n < 1 Then n = 1
    qx = -999999
    For i = 0 To n
        px = Int(x0 + dx * i / n): py = Int(y0 + dy * i / n)
        If px <> qx Or py <> qy Then
            If w <= 1 Then
                If px >= 0 And px < VW And py >= 0 And py < VH Then
                    buf((py * VW + px) * 3) = b: buf((py * VW + px) * 3 + 1) = g: buf((py * VW + px) * 3 + 2) = r
                End If
            Else
                Stamp px, py, w, r, g, b
            End If
            qx = px: qy = py
        End If
    Next
End Sub

Private Function ClipT(ByVal p As Double, ByVal q As Double, t0 As Double, t1 As Double) As Boolean
    Dim t As Double
    If p = 0 Then
        ClipT = (q >= 0)
        Exit Function
    End If
    t = q / p
    If p < 0 Then
        If t > t1 Then Exit Function
        If t > t0 Then t0 = t
    Else
        If t < t0 Then Exit Function
        If t < t1 Then t1 = t
    End If
    ClipT = True
End Function

Private Sub Disc(ByVal cx As Double, ByVal cy As Double, ByVal rad As Double, ByVal r As Long, ByVal g As Long, ByVal b As Long)
    Dim x As Long, y As Long, o As Long
    For y = Int(cy - rad) To Int(cy + rad) + 1
        If y >= 0 And y < VH Then
            For x = Int(cx - rad) To Int(cx + rad) + 1
                If x >= 0 And x < VW Then
                    If (x + 0.5 - cx) ^ 2 + (y + 0.5 - cy) ^ 2 <= rad * rad Then
                        o = (y * VW + x) * 3
                        buf(o) = b: buf(o + 1) = g: buf(o + 2) = r
                    End If
                End If
            Next
        End If
    Next
End Sub

Private Sub Marker(ByVal x As Double, ByVal y As Double, ByVal r As Long, ByVal g As Long, ByVal b As Long)
    Disc x, y, 6.5, &H11, &H18, &H27
    Disc x, y, 5, r, g, b
End Sub

'3x5 digit font, drawn at 2x
Private Sub DrawNumber(ByVal n As Long, ByVal x As Double, ByVal y As Double)
    Dim s As String, i As Long, d As Long, pat As Variant, row As Long, col As Long, bits As String
    Dim px As Long, py As Long, o As Long, sx As Long, sy As Long
    pat = Array("111101101101111", "010110010010111", "111001111100111", "111001111001111", "101101111001001", _
                "111100111001111", "111100111101111", "111001001001001", "111101111101111", "111101111001111")
    s = CStr(n)
    For i = 1 To Len(s)
        d = Asc(Mid$(s, i, 1)) - 48
        bits = pat(d)
        For row = 0 To 4
            For col = 0 To 2
                If Mid$(bits, row * 3 + col + 1, 1) = "1" Then
                    For sy = 0 To 1
                        For sx = 0 To 1
                            px = Int(x) + (i - 1) * 8 + col * 2 + sx: py = Int(y) + row * 2 + sy
                            If px >= 0 And px < VW And py >= 0 And py < VH Then
                                o = (py * VW + px) * 3
                                buf(o) = &H8A: buf(o + 1) = &HE6: buf(o + 2) = &HFD
                            End If
                        Next
                    Next
                End If
            Next
        Next
    Next
End Sub

Private Sub DrawFence(ByVal a0 As Double, ByVal b0 As Double, ByVal a1 As Double, ByVal b1 As Double)
    Dim x0 As Double, y0 As Double, x1 As Double, y1 As Double, t As Double
    Dim xi As Long, yi As Long, o As Long, xa As Long, xb As Long, ya As Long, yb As Long
    x0 = a0 * vz + vox: x1 = a1 * vz + vox: y0 = b0 * vz + voy: y1 = b1 * vz + voy
    If x1 < x0 Then t = x0: x0 = x1: x1 = t
    If y1 < y0 Then t = y0: y0 = y1: y1 = t
    xa = Int(x0): xb = Int(x1): ya = Int(y0): yb = Int(y1)
    If xa < 0 Then xa = 0
    If ya < 0 Then ya = 0
    If xb > VW - 1 Then xb = VW - 1
    If yb > VH - 1 Then yb = VH - 1
    For yi = ya To yb
        o = (yi * VW + xa) * 3
        For xi = xa To xb
            buf(o) = buf(o) + (CLng(&HB) - buf(o)) * 40 \ 255
            buf(o + 1) = buf(o + 1) + (CLng(&H9E) - buf(o + 1)) * 40 \ 255
            buf(o + 2) = buf(o + 2) + (CLng(&HF5) - buf(o + 2)) * 40 \ 255
            o = o + 3
        Next
    Next
    Dashed x0, y0, x1, y0: Dashed x1, y0, x1, y1: Dashed x1, y1, x0, y1: Dashed x0, y1, x0, y0
End Sub

Private Sub Dashed(ByVal x0 As Double, ByVal y0 As Double, ByVal x1 As Double, ByVal y1 As Double)
    Dim L As Double, s As Double, e As Double, dx As Double, dy As Double
    L = Sqr((x1 - x0) ^ 2 + (y1 - y0) ^ 2)
    If L < 0.5 Then Exit Sub
    dx = (x1 - x0) / L: dy = (y1 - y0) / L
    s = 0
    Do While s < L
        e = s + 9: If e > L Then e = L
        ThickLine x0 + dx * s, y0 + dy * s, x0 + dx * e, y0 + dy * e, 2, &HF5, &H9E, &HB
        s = s + 15
    Loop
End Sub

Private Sub Present()
    Dim p As String, c As Object
    Set c = ViewCtl()
    If c Is Nothing Then Exit Sub
    flip = Not flip
    p = gTmpDir & IIf(flip, "\view_a.bmp", "\view_b.bmp")
    On Error Resume Next
    WriteBmp p, buf, VW, 0, 0, VW, VH
    Set c.Picture = LoadPicture(p)
    On Error GoTo 0
    ShowDropHint Not gHasPdf
End Sub

'---------------------------------------------------------------- tools
Public Sub CancelPending()
    pendOn = False
End Sub

Public Sub FenceChanged()
    Dim f As Boolean
    f = gHasFence
    SetCap "chkClip", IIf(f, "Mask outside fence", "Mask outside county")
    SetTip "chkClip", "Make everything outside the " & IIf(f, "block fence", "county boundary") & " transparent in the image"
    SetEnabled "chkHoles", Not f
    SetEnabled "btnFenceClear", f
End Sub

Private Sub ScreenToPreview(ByVal sx As Double, ByVal sy As Double, px As Double, py As Double)
    px = (sx - vox) / vz: py = (sy - voy) / vz
End Sub

Private Sub ClampPreview(px As Double, py As Double)
    Dim mx As Double, my As Double
    mx = gPdfWpt * PREVIEW_DPI / 72#: my = gPdfHpt * PREVIEW_DPI / 72#
    If px < 0 Then px = 0
    If py < 0 Then py = 0
    If px > mx Then px = mx
    If py > my Then py = my
End Sub

Private Function NearestVertex(ByVal sx As Double, ByVal sy As Double, ByVal tol As Double, nx As Double, ny As Double) As Boolean
    Dim k As Long, i As Long, x As Double, y As Double, d As Double, bd As Double
    bd = tol * tol
    For k = 0 To gNR - 1
        For i = 0 To gRingsN(k).N - 1
            NormToScreen gRingsN(k).X(i), gRingsN(k).Y(i), x, y
            d = (x - sx) ^ 2 + (y - sy) ^ 2
            If d < bd Then bd = d: nx = gRingsN(k).X(i): ny = gRingsN(k).Y(i): NearestVertex = True
        Next
    Next
    If Chk("chkShowRoutes") Then
        For k = 0 To gNRt - 1
            For i = 0 To gRoutesN(k).N - 1
                NormToScreen gRoutesN(k).X(i), gRoutesN(k).Y(i), x, y
                d = (x - sx) ^ 2 + (y - sy) ^ 2
                If d < bd Then bd = d: nx = gRoutesN(k).X(i): ny = gRoutesN(k).Y(i): NearestVertex = True
            Next
        Next
    End If
End Function

Private Sub ToPx(ByVal X As Single, ByVal Y As Single, sx As Double, sy As Double)
    sx = X * PtPx(): sy = Y * PtPx()
End Sub

Public Sub ViewMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim sx As Double, sy As Double, px As Double, py As Double, nx As Double, ny As Double
    If Not gHasPdf Then Exit Sub
    ToPx X, Y, sx, sy
    lastSX = sx: lastSY = sy: startSX = sx: startSY = sy
    Select Case gTool
    Case TOOL_FENCE
        If Button <> 1 Then Exit Sub
        ScreenToPreview sx, sy, px, py
        ClampPreview px, py
        fOn = True: fX0 = px: fY0 = py: fX1 = px: fY1 = py
        dragMode = 5
        ViewRedraw False
    Case TOOL_ADJUST
        If Button <> 1 Or Not gHasG Then Exit Sub
        If (Shift And 1) <> 0 Then
            dragMode = 3
        ElseIf (Shift And 2) <> 0 Then
            dragMode = 4
        Else
            dragMode = 2
        End If
    Case TOOL_TIE
        If Not gHasG Then Exit Sub
        If Button = 2 Then
            CancelPending
            ViewRedraw False
            Exit Sub
        End If
        If Button <> 1 Then Exit Sub
        If Not pendOn Then
            If Not NearestVertex(sx, sy, 25, nx, ny) Then
                Toast "Click closer to the cyan overlay to pick a source point."
                Exit Sub
            End If
            pendOn = True: pendX = nx: pendY = ny
            ViewRedraw False
        Else
            ScreenToPreview sx, sy, px, py
            pendOn = False
            AddPair pendX, pendY, px, py
            ViewRedraw False
        End If
    Case Else
        If Button <> 1 Then Exit Sub
        dragMode = 1
        startVox = vox: startVoy = voy
    End Select
End Sub

Public Sub ViewMouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim sx As Double, sy As Double, px As Double, py As Double, cx As Double, cy As Double
    Dim a As Double, co As Double, si As Double, f As Double
    If dragMode = 0 Then Exit Sub
    If (Button And 1) = 0 Then ViewMouseUp 1, Shift, X, Y: Exit Sub
    ToPx X, Y, sx, sy
    Select Case dragMode
    Case 1
        vox = startVox + (sx - startSX): voy = startVoy + (sy - startSY)
        If GetTickCount - lastDraw > 150 Then ViewRedraw True
    Case 2
        ComposeG 1, 0, (sx - lastSX) / vz, 0, 1, (sy - lastSY) / vz
    Case 3
        ScreenToPreview VW / 2#, VH / 2#, cx, cy
        a = Atn2(sy - VH / 2#, sx - VW / 2#) - Atn2(lastSY - VH / 2#, lastSX - VW / 2#)
        co = Cos(a): si = Sin(a)
        ComposeG co, -si, cx - co * cx + si * cy, si, co, cy - si * cx - co * cy
    Case 4
        ScreenToPreview VW / 2#, VH / 2#, cx, cy
        f = Exp(-(sy - lastSY) * 0.004)
        ComposeG f, 0, cx * (1 - f), 0, f, cy * (1 - f)
    Case 5
        ScreenToPreview sx, sy, px, py
        ClampPreview px, py
        fX1 = px: fY1 = py
    End Select
    lastSX = sx: lastSY = sy
    If dragMode >= 2 Then
        If GetTickCount - lastDraw > 90 Then ViewRedraw False
    End If
End Sub

Public Sub ViewMouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim m As Long, x0 As Double, y0 As Double, x1 As Double, y1 As Double
    m = dragMode
    dragMode = 0
    Select Case m
    Case 1
        ViewRedraw True
    Case 2, 3, 4
        UpdateFitLabel
        ViewRedraw False
    Case 5
        fOn = False
        If Abs(fX1 - fX0) * vz < 6 Or Abs(fY1 - fY0) * vz < 6 Then
            ViewRedraw False
            Exit Sub
        End If
        x0 = IIf(fX0 < fX1, fX0, fX1): x1 = IIf(fX0 < fX1, fX1, fX0)
        y0 = IIf(fY0 < fY1, fY0, fY1): y1 = IIf(fY0 < fY1, fY1, fY0)
        OnFenceDrawn x0, y0, x1, y1
        ViewRedraw False
    End Select
End Sub

Public Sub ViewInvalidate()
    baseOK = False
End Sub
