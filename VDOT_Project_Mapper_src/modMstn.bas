Attribute VB_Name = "modMstn"
'MicroStation side: PDF rendering through Raster Manager, BMP/PNG files, masked image, georeferenced
'attachments, boundary/route elements, scale bar and north arrow, fence and view fit.
Option Explicit

Private mTmp As Raster          ' temporary (invisible) PDF attachment used for rendering
Private mImgT() As Double       ' last masked image: pixel -> world
Private mImgW As Long, mImgH As Long
' high-resolution tile state (BeginMaskedTiles / MakeMaskedTile / EndMaskedTiles)
Private mBig As String, mBigF As Integer, mBigOff As Long, mBigStride As Long, mBigBpp As Long
Private mBigW As Long, mBigH As Long, mBigTop As Boolean
Private mKx As Double, mKy As Double, mRx() As TPoly, mNM As Long, mWhite As Boolean
Private mCx0 As Long, mCy0 As Long, mCw As Long, mCh As Long, mNTX As Long, mNTY As Long
Private mOutBase As String

'---------------------------------------------------------------- little-endian helpers
Private Function GetL(b() As Byte, ByVal o As Long) As Long
    Dim v As Long
    CopyMemory v, b(o), 4
    GetL = v
End Function

Private Function GetI(b() As Byte, ByVal o As Long) As Long
    GetI = b(o) + 256& * b(o + 1)
End Function

Private Sub PutL(b() As Byte, ByVal o As Long, ByVal v As Long)
    CopyMemory b(o), v, 4
End Sub

'---------------------------------------------------------------- BMP files
'Reads a BMP into a top-down BGR buffer (stride W*3). Supports 8/24/32 bit.
Public Function ReadBmp(ByVal path As String, px() As Byte, W As Long, H As Long) As Boolean
    Dim f As Integer, hdr(0 To 53) As Byte, off As Long, bpp As Long, topDown As Boolean
    Dim stride As Long, y As Long, x As Long, row() As Byte, srcY As Long, hs As Long, pal() As Byte, nPal As Long
    Dim i As Long, o As Long, c As Long
    f = FreeFile
    Open path For Binary Access Read As #f
    Get #f, 1, hdr
    If hdr(0) <> 66 Or hdr(1) <> 77 Then Close #f: Err.Raise vbObjectError + 700, , "Not a BMP file: " & path
    off = GetL(hdr, 10): hs = GetL(hdr, 14)
    W = GetL(hdr, 18): H = GetL(hdr, 22): bpp = GetI(hdr, 28)
    If H < 0 Then H = -H: topDown = True
    If W <= 0 Or H <= 0 Then Close #f: Err.Raise vbObjectError + 701, , "Empty bitmap."
    stride = ((W * bpp + 31) \ 32) * 4
    ReDim px(0 To W * H * 3 - 1)
    ReDim row(0 To stride - 1)
    If bpp <= 8 Then
        nPal = GetL(hdr, 46): If nPal = 0 Then nPal = 2 ^ bpp
        ReDim pal(0 To nPal * 4 - 1)
        Get #f, 15 + hs, pal
    End If
    For y = 0 To H - 1
        If topDown Then srcY = y Else srcY = H - 1 - y
        Get #f, off + 1 + CLng(srcY) * stride, row
        o = y * W * 3
        Select Case bpp
        Case 24
            CopyMemory px(o), row(0), W * 3
        Case 32
            For x = 0 To W - 1
                i = x * 4
                px(o) = row(i): px(o + 1) = row(i + 1): px(o + 2) = row(i + 2): o = o + 3
            Next
        Case 8
            For x = 0 To W - 1
                c = row(x) * 4
                px(o) = pal(c): px(o + 1) = pal(c + 1): px(o + 2) = pal(c + 2): o = o + 3
            Next
        Case 4
            For x = 0 To W - 1
                If x Mod 2 = 0 Then c = (row(x \ 2) \ 16) * 4 Else c = (row(x \ 2) And 15) * 4
                px(o) = pal(c): px(o + 1) = pal(c + 1): px(o + 2) = pal(c + 2): o = o + 3
            Next
        Case 1
            For x = 0 To W - 1
                If (row(x \ 8) And (2 ^ (7 - (x Mod 8)))) <> 0 Then c = 4 Else c = 0
                px(o) = pal(c): px(o + 1) = pal(c + 1): px(o + 2) = pal(c + 2): o = o + 3
            Next
        Case Else
            Close #f
            Err.Raise vbObjectError + 702, , "Unsupported BMP depth " & bpp
        End Select
    Next
    Close #f
    ReadBmp = True
End Function

Private Sub BmpHeader(hdr() As Byte, ByVal W As Long, ByVal H As Long, stride As Long)
    stride = ((W * 3 + 3) \ 4) * 4
    ReDim hdr(0 To 53)
    hdr(0) = 66: hdr(1) = 77
    PutL hdr, 2, 54 + stride * H: PutL hdr, 10, 54: PutL hdr, 14, 40
    PutL hdr, 18, W: PutL hdr, 22, H
    hdr(26) = 1: hdr(28) = 24
    PutL hdr, 34, stride * H
    PutL hdr, 38, 3780: PutL hdr, 42, 3780
End Sub

'Writes a top-down BGR buffer (stride srcW*3) or a sub-rectangle of it as a 24-bit BMP.
Public Sub WriteBmp(ByVal path As String, px() As Byte, ByVal srcW As Long, ByVal x0 As Long, ByVal y0 As Long, _
                    ByVal W As Long, ByVal H As Long)
    Dim hdr() As Byte, stride As Long, f As Integer, y As Long, row() As Byte
    BmpHeader hdr, W, H, stride
    If Dir$(path) <> "" Then Kill path
    f = FreeFile
    Open path For Binary Access Write As #f
    Put #f, 1, hdr
    ReDim row(0 To stride - 1)
    For y = H - 1 To 0 Step -1
        CopyMemory row(0), px(((y0 + y) * srcW + x0) * 3), W * 3
        Put #f, , row
    Next
    Close #f
End Sub

'Converts a BMP to PNG with Windows Image Acquisition. Returns False when WIA is unavailable.
Public Function BmpToPng(ByVal bmp As String, ByVal png As String) As Boolean
    Dim img As Object, ip As Object
    On Error GoTo EH
    Set img = CreateObject("WIA.ImageFile")
    img.LoadFile bmp
    Set ip = CreateObject("WIA.ImageProcess")
    ip.Filters.Add ip.FilterInfos("Convert").FilterID
    ip.Filters(1).Properties("FormatID").Value = "{B96B3CAF-0728-11D3-9D7B-0000F81EF32E}"
    Set img = ip.Apply(img)
    If Dir$(png) <> "" Then Kill png
    img.SaveFile png
    BmpToPng = True
    Exit Function
EH:
    BmpToPng = False
End Function

'---------------------------------------------------------------- PDF rendering
Public Sub PdfOpen(ByVal pdf As String, wpt As Double, hpt As Double, extX As Double, extY As Double)
    Dim ri As RasterInformation, gi As GeoReferenceInformation, res As Point2d
    PdfClose
    If Dir$(pdf) = "" Then Err.Raise vbObjectError + 710, , "File not found: " & pdf
    Set mTmp = RasterManager.Rasters.Attach(pdf, 0)
    If mTmp Is Nothing Then Err.Raise vbObjectError + 711, , "MicroStation could not open " & FileTitle(pdf)
    Set ri = mTmp.RasterInformation
    extX = ri.Extent.X: extY = ri.Extent.Y
    Set gi = mTmp.GeoReferenceInformation
    res = gi.ScanningResolution
    If res.X <= 0 Then res.X = 72
    If res.Y <= 0 Then res.Y = res.X
    wpt = extX * 72# / res.X
    hpt = extY * 72# / res.Y
End Sub

'Renders the open PDF page to a W x H BGR buffer.
Public Sub PdfRender(ByVal W As Long, ByVal H As Long, px() As Byte, outW As Long, outH As Long)
    Dim bs As Point2d, so As Point2d, se As Point2d, bm As RasterBitmap, tmp As String
    If mTmp Is Nothing Then Err.Raise vbObjectError + 712, , "No PDF open."
    bs.X = W: bs.Y = H
    se = mTmp.RasterInformation.Extent
    Set bm = mTmp.GetBitmap(bs, so, se)
    If bm Is Nothing Then Err.Raise vbObjectError + 713, , "MicroStation could not render the PDF page."
    tmp = gTmpDir & "\render.bmp"
    If Dir$(tmp) <> "" Then Kill tmp
    bm.ExportRaster tmp, 9, 1, 0, 4
    Set bm = Nothing
    ReadBmp tmp, px, outW, outH
    Kill tmp
End Sub

Public Sub PdfClose()
    Dim el As Element
    If mTmp Is Nothing Then Exit Sub
    On Error Resume Next
    Set el = RasterManager.Rasters.GetElementFromRaster(mTmp)
    Set mTmp = Nothing
    If Not el Is Nothing Then ActiveModelReference.RemoveElement el
    Err.Clear
End Sub

'Render a PDF page; scales down (up to 3 times) if MicroStation cannot produce the requested size.
Public Sub RenderPdfPage(ByVal pdf As String, ByVal dpi As Double, px() As Byte, W As Long, H As Long, _
                         wpt As Double, hpt As Double, Optional ByVal maxPx As Double = 0)
    Dim ex As Double, ey As Double, tw As Long, th As Long, tries As Long, msg As String
    On Error GoTo EH
    PdfOpen pdf, wpt, hpt, ex, ey
    If maxPx > 0 Then
        If (wpt / 72# * dpi) * (hpt / 72# * dpi) > maxPx Then dpi = Sqr(maxPx / ((wpt / 72#) * (hpt / 72#)))
    End If
Retry:
    tw = CLng(wpt / 72# * dpi): th = CLng(hpt / 72# * dpi)
    If tw < 1 Then tw = 1
    If th < 1 Then th = 1
    PdfRender tw, th, px, W, H
    PdfClose
    Exit Sub
EH:
    tries = tries + 1
    msg = Err.Description
    If tries <= 3 And Not mTmp Is Nothing Then
        dpi = dpi * 0.7
        Resume Retry
    End If
    PdfClose
    Err.Raise vbObjectError + 714, , "Rendering the PDF failed: " & msg
End Sub

'---------------------------------------------------------------- masked map image
'image pixel (top-left origin, this render) -> world:  wx = t(0)*i + t(1)*j + t(2);  wy = t(3)*i + t(4)*j + t(5)
Private Sub ImageToWorldT(ByVal kx As Double, ByVal ky As Double, ByVal ox As Double, ByVal oy As Double, t() As Double)
    Dim x0 As Double, y0 As Double, x1 As Double, y1 As Double, x2 As Double, y2 As Double
    ReDim t(0 To 5)
    PreviewToWorld (ox) / kx, (oy) / ky, x0, y0
    PreviewToWorld (ox + 1) / kx, (oy) / ky, x1, y1
    PreviewToWorld (ox) / kx, (oy + 1) / ky, x2, y2
    t(0) = x1 - x0: t(1) = x2 - x0: t(2) = x0
    t(3) = y1 - y0: t(4) = y2 - y0: t(5) = y0
End Sub

Private Sub WorldToImage(ByVal wx As Double, ByVal wy As Double, ByVal kx As Double, ByVal ky As Double, ix As Double, iy As Double)
    Dim px As Double, py As Double
    ApplyG gG, (wx - gNormCx) / gNormS, (wy - gNormCy) / gNormS, px, py
    ix = px * kx: iy = py * ky
End Sub

'---------------------------------------------------------------- high-resolution masked tiles
'The page is rendered once at the chosen DPI to a temporary BMP; each tile (<= TILE_PX square) of the mask's
'bounding box is then streamed from that file, masked (KEY_COLOR outside the even-odd rings) and saved as
'<pdf>_county_tNN.png. Tiles that lie completely outside the mask are skipped.

'Renders the page and prepares the tile grid. Returns the number of tile slots; dpi returns the DPI used.
Public Function BeginMaskedTiles(mask() As TPoly, ByVal nM As Long, ByVal whiteOut As Boolean, dpi As Double) As Long
    Dim wpt As Double, hpt As Double, ex As Double, ey As Double, bs As Point2d, so As Point2d, se As Point2d
    Dim bm As RasterBitmap, tries As Long, msg As String, hdr(0 To 53) As Byte, k As Long, i As Long
    Dim ix As Double, iy As Double, bx0 As Double, by0 As Double, bx1 As Double, by1 As Double
    EndMaskedTiles
    If (gPdfWpt / 72# * dpi) * (gPdfHpt / 72# * dpi) > IMAGE_MAX_TOTAL_PX Then
        dpi = Sqr(IMAGE_MAX_TOTAL_PX / ((gPdfWpt / 72#) * (gPdfHpt / 72#)))
    End If
    mBig = gTmpDir & "\render_hires.bmp"
    If Dir$(mBig) <> "" Then Kill mBig
    PdfOpen gPdfPath, wpt, hpt, ex, ey
    On Error GoTo EH
Retry:
    SetStatus "Rendering the PDF at " & CLng(dpi) & " DPI (large maps can take a few minutes)" & UEll()
    DoEvents
    bs.X = CLng(gPdfWpt / 72# * dpi): bs.Y = CLng(gPdfHpt / 72# * dpi)
    se = mTmp.RasterInformation.Extent
    Set bm = mTmp.GetBitmap(bs, so, se)
    If bm Is Nothing Then Err.Raise vbObjectError + 713, , "MicroStation could not render the PDF page."
    bm.ExportRaster mBig, 9, 1, 0, 4
    Set bm = Nothing
    On Error GoTo 0
    PdfClose

    mBigF = FreeFile
    Open mBig For Binary Access Read As #mBigF
    Get #mBigF, 1, hdr
    mBigOff = GetL(hdr, 10): mBigW = GetL(hdr, 18): mBigH = GetL(hdr, 22): mBigBpp = GetI(hdr, 28)
    mBigTop = mBigH < 0: If mBigTop Then mBigH = -mBigH
    If mBigBpp <> 24 And mBigBpp <> 32 Then
        EndMaskedTiles
        Err.Raise vbObjectError + 702, , "Unsupported render depth " & mBigBpp
    End If
    mBigStride = ((mBigW * mBigBpp + 31) \ 32) * 4
    mKx = mBigW / (gPdfWpt * PREVIEW_DPI / 72#)
    mKy = mBigH / (gPdfHpt * PREVIEW_DPI / 72#)
    mNM = nM: mWhite = whiteOut
    If nM > 0 Then
        ReDim mRx(0 To nM - 1)
        bx0 = 1E+300: by0 = 1E+300: bx1 = -1E+300: by1 = -1E+300
        For k = 0 To nM - 1
            PolyInit mRx(k), mask(k).N
            For i = 0 To mask(k).N - 1
                WorldToImage mask(k).X(i), mask(k).Y(i), mKx, mKy, ix, iy
                PolyAdd mRx(k), ix, iy
                If ix < bx0 Then bx0 = ix
                If ix > bx1 Then bx1 = ix
                If iy < by0 Then by0 = iy
                If iy > by1 Then by1 = iy
            Next
        Next
        If bx0 < 0 Then bx0 = 0
        If by0 < 0 Then by0 = 0
        If bx1 > mBigW Then bx1 = mBigW
        If by1 > mBigH Then by1 = mBigH
        mCx0 = Int(bx0): mCy0 = Int(by0)
        mCw = -Int(-bx1) - mCx0: mCh = -Int(-by1) - mCy0
        If mCw <= 0 Or mCh <= 0 Then
            EndMaskedTiles
            Err.Raise vbObjectError + 720, , "The county boundary does not overlap the PDF page; check the alignment."
        End If
    Else
        mCx0 = 0: mCy0 = 0: mCw = mBigW: mCh = mBigH
    End If
    mNTX = -Int(-mCw / TILE_PX): mNTY = -Int(-mCh / TILE_PX)

    ' output location: next to the PDF, else LOCALAPPDATA; remove stale tiles from earlier placements
    Dim baseName As String, cands(0 To 1) As String, c As Variant, old As String, dirp As String
    baseName = FileTitle(gPdfPath)
    baseName = Left$(baseName, InStrRev(baseName, ".") - 1) & IIf(nM > 0, "_county", "_map")
    cands(0) = Left$(gPdfPath, InStrRev(gPdfPath, "\")) & baseName
    cands(1) = Environ$("LOCALAPPDATA") & "\VDOT_Georef\" & baseName
    mOutBase = ""
    For Each c In cands
        If CanWrite(CStr(c) & ".png") Then mOutBase = CStr(c): Exit For
    Next
    If Len(mOutBase) = 0 Then EndMaskedTiles: Err.Raise vbObjectError + 721, , "Could not write the map image."
    dirp = Left$(mOutBase, InStrRev(mOutBase, "\"))
    On Error Resume Next
    old = Dir$(mOutBase & "_t??.*")
    Do While Len(old) > 0
        DetachExisting dirp & old
        Kill dirp & old
        old = Dir$()
    Loop
    Err.Clear
    On Error GoTo 0
    BeginMaskedTiles = mNTX * mNTY
    Exit Function
EH:
    tries = tries + 1
    msg = Err.Description
    Set bm = Nothing
    If tries <= 3 And Not mTmp Is Nothing Then
        dpi = dpi * 0.7
        Resume Retry
    End If
    PdfClose
    Err.Raise vbObjectError + 714, , "Rendering the PDF failed: " & msg
End Function

'Builds tile t (0-based). Returns "" when the tile is entirely outside the mask; otherwise the file path,
'with mImgT / mImgW / mImgH set for AttachImage.
Public Function MakeMaskedTile(ByVal t As Long, ByVal seq As Long, iw As Long, ih As Long) As String
    Dim tx0 As Long, ty0 As Long, cw As Long, ch As Long, k As Long, i As Long, j As Long, jA As Long, jB As Long
    Dim cnt() As Long, start() As Long, xs() As Double, fillp() As Long, nX As Long
    Dim xa As Double, ya As Double, xb As Double, yb As Double
    tx0 = mCx0 + (t Mod mNTX) * TILE_PX: ty0 = mCy0 + (t \ mNTX) * TILE_PX
    cw = mCx0 + mCw - tx0: If cw > TILE_PX Then cw = TILE_PX
    ch = mCy0 + mCh - ty0: If ch > TILE_PX Then ch = TILE_PX

    If mNM > 0 Then
        ReDim cnt(0 To ch): ReDim start(0 To ch)
        For k = 0 To mNM - 1
            For i = 0 To mRx(k).N - 1
                ya = mRx(k).Y(i)
                If i < mRx(k).N - 1 Then yb = mRx(k).Y(i + 1) Else yb = mRx(k).Y(0)
                If ya <> yb Then
                    RowRange ya, yb, ty0, ch, jA, jB
                    For j = jA To jB: cnt(j) = cnt(j) + 1: Next
                End If
            Next
        Next
        For j = 0 To ch - 1: start(j + 1) = start(j) + cnt(j): Next
        nX = start(ch)
        If nX = 0 Then Exit Function
        ReDim xs(0 To nX - 1)
        ReDim fillp(0 To ch)
        For j = 0 To ch: fillp(j) = start(j): Next
        For k = 0 To mNM - 1
            For i = 0 To mRx(k).N - 1
                xa = mRx(k).X(i): ya = mRx(k).Y(i)
                If i < mRx(k).N - 1 Then xb = mRx(k).X(i + 1): yb = mRx(k).Y(i + 1) Else xb = mRx(k).X(0): yb = mRx(k).Y(0)
                If ya <> yb Then
                    RowRange ya, yb, ty0, ch, jA, jB
                    For j = jA To jB
                        xs(fillp(j)) = xa + (ty0 + j + 0.5 - ya) * (xb - xa) / (yb - ya) - tx0
                        fillp(j) = fillp(j) + 1
                    Next
                End If
            Next
        Next
    End If

    Dim hdr() As Byte, stride As Long, f As Integer, row() As Byte, keyRow() As Byte, src() As Byte
    Dim a As Long, b As Long, s As Long, e As Long, p As Long, q As Long, tmpBmp As String, tv As Double
    Dim keyB As Byte, keyG As Byte, keyR As Byte, anyIn As Boolean, srcY As Long, nb As Long, x As Long
    keyR = KEY_COLOR And &HFF: keyG = (KEY_COLOR \ &H100) And &HFF: keyB = (KEY_COLOR \ &H10000) And &HFF
    nb = mBigBpp \ 8
    BmpHeader hdr, cw, ch, stride
    ReDim row(0 To stride - 1): ReDim keyRow(0 To stride - 1): ReDim src(0 To cw * nb - 1)
    For i = 0 To cw - 1
        keyRow(i * 3) = keyB: keyRow(i * 3 + 1) = keyG: keyRow(i * 3 + 2) = keyR
    Next
    tmpBmp = gTmpDir & "\masked_tile.bmp"
    If Dir$(tmpBmp) <> "" Then Kill tmpBmp
    f = FreeFile
    Open tmpBmp For Binary Access Write As #f
    Put #f, 1, hdr
    For j = 0 To ch - 1
        If mBigTop Then srcY = ty0 + j Else srcY = mBigH - 1 - (ty0 + j)
        Get #mBigF, mBigOff + 1 + CLng(srcY) * mBigStride + CLng(tx0) * nb, src
        If nb = 4 Then
            For x = 0 To cw - 1
                src(x * 3) = src(x * 4): src(x * 3 + 1) = src(x * 4 + 1): src(x * 3 + 2) = src(x * 4 + 2)
            Next
        End If
        If mNM > 0 Then
            CopyMemory row(0), keyRow(0), stride
            a = start(j): b = start(j + 1) - 1
            For p = a + 1 To b
                tv = xs(p): q = p - 1
                Do While q >= a
                    If xs(q) <= tv Then Exit Do
                    xs(q + 1) = xs(q): q = q - 1
                Loop
                xs(q + 1) = tv
            Next
            For p = a To b - 1 Step 2
                s = -Int(-(xs(p) - 0.5)): e = -Int(-(xs(p + 1) - 0.5)) - 1
                If s < 0 Then s = 0
                If e > cw - 1 Then e = cw - 1
                If e >= s Then CopyMemory row(s * 3), src(s * 3), (e - s + 1) * 3: anyIn = True
            Next
        Else
            CopyMemory row(0), src(0), cw * 3
            anyIn = True
        End If
        If mWhite Then
            For i = 0 To cw * 3 - 1 Step 3
                If row(i) >= 235 Then
                    If row(i + 1) >= 235 Then
                        If row(i + 2) >= 235 Then row(i) = keyB: row(i + 1) = keyG: row(i + 2) = keyR
                    End If
                End If
            Next
        End If
        Put #f, 55 + CLng(ch - 1 - j) * stride, row
        If j Mod 500 = 0 Then DoEvents
    Next
    Close #f
    If Not anyIn Then Kill tmpBmp: Exit Function

    ImageToWorldT mKx, mKy, tx0, ty0, mImgT
    mImgW = cw: mImgH = ch
    iw = cw: ih = ch
    Dim outp As String
    outp = mOutBase & "_t" & Format$(seq, "00") & ".png"
    DetachExisting outp
    If Not BmpToPng(tmpBmp, outp) Then
        outp = Left$(outp, Len(outp) - 4) & ".bmp"
        DetachExisting outp
        If Dir$(outp) <> "" Then Kill outp
        FileCopy tmpBmp, outp
    End If
    Kill tmpBmp
    MakeMaskedTile = outp
End Function

Public Sub EndMaskedTiles()
    On Error Resume Next
    If mBigF <> 0 Then Close #mBigF
    mBigF = 0
    If Len(mBig) > 0 Then If Dir$(mBig) <> "" Then Kill mBig
    mBig = ""
    PdfClose
    Err.Clear
End Sub

Private Sub RowRange(ByVal ya As Double, ByVal yb As Double, ByVal cy0 As Long, ByVal ch As Long, jA As Long, jB As Long)
    Dim lo As Double, hi As Double
    If ya < yb Then lo = ya: hi = yb Else lo = yb: hi = ya
    ' rows whose centre (cy0 + j + 0.5) lies in [lo, hi)
    jA = -Int(-(lo - 0.5 - cy0))
    jB = -Int(-(hi - 0.5 - cy0)) - 1
    If jA < 0 Then jA = 0
    If jB > ch - 1 Then jB = ch - 1
End Sub

Private Function CanWrite(ByVal path As String) As Boolean
    Dim d As String, f As Integer, t As String
    On Error GoTo No
    d = Left$(path, InStrRev(path, "\") - 1)
    If Dir$(d, vbDirectory) = "" Then MkDir d
    t = d & "\~vdot_write_test.tmp"
    f = FreeFile
    Open t For Output As #f
    Close #f
    Kill t
    CanWrite = True
    Exit Function
No:
    CanWrite = False
End Function

'Remove an earlier attachment of the same file so it can be overwritten.
Private Sub DetachExisting(ByVal path As String)
    Dim i As Long, els As New Collection, r As Raster, e As Variant
    On Error Resume Next
    For i = 1 To RasterManager.Rasters.Count
        Set r = RasterManager.Rasters(i)
        If LCase$(r.RasterInformation.FullName) = LCase$(path) Then els.Add RasterManager.Rasters.GetElementFromRaster(r)
    Next
    Set r = Nothing
    For Each e In els
        ActiveModelReference.RemoveElement e
    Next
    Err.Clear
End Sub

'---------------------------------------------------------------- raster attachments
Private Sub Georeference(r As Raster, t() As Double, ByVal W As Long, ByVal H As Long)
    Dim gi As GeoReferenceInformation, o As Point3d, e As Point2d
    Dim ax As Double, ay As Double, bx As Double, by As Double, th As Double, al As Double
    ax = t(0) * W: ay = t(3) * W
    bx = -t(1) * H: by = -t(4) * H
    o.X = t(1) * H + t(2): o.Y = t(4) * H + t(5)
    th = Atn2(ay, ax)
    al = Atn2(by, bx) - th - 2 * Atn(1)
    Do While al > 4 * Atn(1): al = al - 8 * Atn(1): Loop
    Do While al < -4 * Atn(1): al = al + 8 * Atn(1): Loop
    e.X = Sqr(ax * ax + ay * ay): e.Y = Sqr(bx * bx + by * by)
    Set gi = r.GeoReferenceInformation
    gi.Origin = o
    gi.Extent = e
    gi.RotationAboutZ = th
    gi.Affinity = al
End Sub

Private Sub SetRasterLevel(r As Raster, ByVal lvName As String)
    Dim el As Element
    On Error Resume Next
    Set el = RasterManager.Rasters.GetElementFromRaster(r)
    el.Level = GetLevel(lvName)
    el.Rewrite
    Err.Clear
End Sub

Public Sub AttachImage(ByVal path As String, ByVal lvName As String, ByVal transpPct As Long, ByVal back As Boolean)
    Dim r As Raster, ri As RenderingInformation
    Set r = RasterManager.Rasters.Attach(path, 255)
    If r Is Nothing Then Err.Raise vbObjectError + 730, , "Attaching " & FileTitle(path) & " failed."
    Georeference r, mImgT, mImgW, mImgH
    Set ri = r.RenderingInformation
    On Error Resume Next
    ri.Transparency = True
    ri.TransparencyColor = KEY_COLOR
    ri.TransparencyColorLevel = 100
    ri.TransparencyAllColorsLevel = transpPct
    If back Then ri.DisplayPriorityPlane = msdRasterDisplayPriorityPlaneBack
    On Error GoTo 0
    SetRasterLevel r, lvName
End Sub

Public Sub AttachPdfHidden(ByVal lvName As String, Optional iw As Long, Optional ih As Long)
    Dim r As Raster, ri As RenderingInformation, t() As Double, ex As Double, ey As Double, kx As Double, ky As Double
    DetachHiddenPdf
    Set r = RasterManager.Rasters.Attach(gPdfPath, 0)
    If r Is Nothing Then Err.Raise vbObjectError + 731, , "Attaching the PDF failed."
    ex = r.RasterInformation.Extent.X: ey = r.RasterInformation.Extent.Y
    kx = ex / (gPdfWpt * PREVIEW_DPI / 72#): ky = ey / (gPdfHpt * PREVIEW_DPI / 72#)
    ImageToWorldT kx, ky, 0, 0, t
    Georeference r, t, CLng(ex), CLng(ey)
    Set ri = r.RenderingInformation
    On Error Resume Next
    ri.View1 = False: ri.View2 = False: ri.View3 = False: ri.View4 = False
    ri.View5 = False: ri.View6 = False: ri.View7 = False: ri.View8 = False
    ri.DisplayPriorityPlane = msdRasterDisplayPriorityPlaneBack
    On Error GoTo 0
    SetRasterLevel r, lvName
    iw = CLng(ex): ih = CLng(ey)
End Sub

'Earlier hidden copies of this PDF (same file, off in every view) are replaced on re-placement.
Private Sub DetachHiddenPdf()
    Dim i As Long, els As New Collection, r As Raster, e As Variant, ri As RenderingInformation
    On Error Resume Next
    For i = 1 To RasterManager.Rasters.Count
        Set r = RasterManager.Rasters(i)
        If LCase$(r.RasterInformation.FullName) = LCase$(gPdfPath) Then
            Set ri = r.RenderingInformation
            If Not (ri.View1 Or ri.View2 Or ri.View3 Or ri.View4 Or ri.View5 Or ri.View6 Or ri.View7 Or ri.View8) Then
                If r.GeoReferenceInformation.Extent.X > 1 Then els.Add RasterManager.Rasters.GetElementFromRaster(r)
            End If
        End If
    Next
    Set r = Nothing: Set ri = Nothing
    For Each e In els
        ActiveModelReference.RemoveElement e
    Next
    Err.Clear
End Sub

'---------------------------------------------------------------- delete placed drawing
'Collects the active model's elements on the given levels (raster attachments via Raster Manager).
Public Sub CollectOnLevels(names() As String, ByVal n As Long, els As Collection, nRast As Long)
    Dim seen As Object, i As Long, sc As ElementScanCriteria, ee As ElementEnumerator, el As Element
    Dim r As Raster, nm As String, key As String, anyLv As Boolean, lv As Level
    Set seen = CreateObject("Scripting.Dictionary"): seen.CompareMode = 1
    For i = 0 To n - 1: seen("lv:" & names(i)) = True: Next
    On Error Resume Next
    For i = 1 To RasterManager.Rasters.Count
        Set r = RasterManager.Rasters(i)
        Set el = Nothing
        Set el = RasterManager.Rasters.GetElementFromRaster(r)
        If Not el Is Nothing Then
            nm = "": nm = el.Level.Name
            key = DLongToString(el.ID)
            If seen.Exists("lv:" & nm) And Not seen.Exists(key) Then
                seen(key) = True: els.Add el: nRast = nRast + 1
            End If
        End If
    Next
    Err.Clear
    Set sc = New ElementScanCriteria
    sc.ExcludeAllLevels
    For i = 0 To n - 1
        Set lv = Nothing
        Set lv = ActiveDesignFile.Levels(names(i))
        If Not lv Is Nothing Then sc.IncludeLevel lv: anyLv = True
    Next
    Err.Clear
    On Error GoTo 0
    If Not anyLv Then Exit Sub
    Set ee = ActiveModelReference.Scan(sc)
    Do While ee.MoveNext
        Set el = ee.Current
        key = DLongToString(el.ID)
        If el.Type <> 94 And Not seen.Exists(key) Then seen(key) = True: els.Add el
    Loop
End Sub

Public Function RemoveElements(els As Collection) As Long
    Dim e As Variant
    On Error Resume Next
    For Each e In els
        Err.Clear
        ActiveModelReference.RemoveElement e
        If Err.Number = 0 Then RemoveElements = RemoveElements + 1
    Next
    Err.Clear
End Function

'---------------------------------------------------------------- vector elements
Public Function GetLevel(ByVal nm As String) As Level
    Dim lv As Level
    On Error Resume Next
    Set lv = ActiveDesignFile.Levels(nm)
    On Error GoTo 0
    If lv Is Nothing Then
        Set lv = ActiveDesignFile.AddNewLevel(nm)
        ActiveDesignFile.Levels.Rewrite
    End If
    Set GetLevel = lv
End Function

'Creates the level (if missing) with by-level colour/weight and a description, so it mirrors the LRS class.
Public Sub EnsureLevel(ByVal nm As String, ByVal color As Long, ByVal weight As Long, ByVal desc As String)
    Dim lv As Level
    On Error Resume Next
    Set lv = ActiveDesignFile.Levels(nm)
    If Not lv Is Nothing Then Exit Sub
    Err.Clear
    Set lv = ActiveDesignFile.AddNewLevel(nm)
    If lv Is Nothing Then Exit Sub
    lv.ElementColor = color
    lv.ElementLineWeight = weight
    lv.Description = desc
    ActiveDesignFile.Levels.Rewrite
End Sub

Private Sub Props(el As Element, ByVal lvName As String, ByVal color As Long, ByVal weight As Long)
    el.Level = GetLevel(lvName)
    If color >= 0 Then el.Color = color
    If weight >= 0 Then el.LineWeight = weight
End Sub

Public Function AddCurve(p As TPoly, ByVal closed As Boolean, ByVal lvName As String, ByVal color As Long, _
                         ByVal weight As Long, Optional ByVal filled As Boolean = False) As Boolean
    Dim v() As Point3d, n As Long, i As Long, el As Element
    Const MAXV As Long = 4900
    If p.N < 2 Then Exit Function
    ReDim v(0 To p.N)
    For i = 0 To p.N - 1
        If n > 0 Then
            If Abs(p.X(i) - v(n - 1).X) < 1E-09 And Abs(p.Y(i) - v(n - 1).Y) < 1E-09 Then GoTo NextPt
        End If
        v(n).X = p.X(i): v(n).Y = p.Y(i): n = n + 1
NextPt:
    Next
    If closed Then
        If Abs(v(0).X - v(n - 1).X) > 1E-09 Or Abs(v(0).Y - v(n - 1).Y) > 1E-09 Then v(n) = v(0): n = n + 1
        If n < 4 Then Exit Function
    ElseIf n < 2 Then
        Exit Function
    End If
    ReDim Preserve v(0 To n - 1)
    On Error GoTo EH
    If n <= MAXV Then
        If closed Then
            Set el = CreateShapeElement1(Nothing, v, IIf(filled, msdFillModeFilled, msdFillModeNotFilled))
        Else
            Set el = CreateLineElement1(Nothing, v)
        End If
    Else
        Dim parts() As ChainableElement, np As Long, s As Long, cnt As Long, seg() As Point3d, k As Long
        np = (n - 2) \ (MAXV - 1) + 1
        ReDim parts(0 To np - 1)
        s = 0
        For k = 0 To np - 1
            cnt = MAXV
            If s + cnt > n Then cnt = n - s
            ReDim seg(0 To cnt - 1)
            For i = 0 To cnt - 1: seg(i) = v(s + i): Next
            Set parts(k) = CreateLineElement1(Nothing, seg)
            s = s + cnt - 1
        Next
        If closed Then
            Set el = CreateComplexShapeElement1(parts, IIf(filled, msdFillModeFilled, msdFillModeNotFilled))
        Else
            Set el = CreateComplexStringElement1(parts)
        End If
    End If
    Props el, lvName, color, weight
    ActiveModelReference.AddElement el
    AddCurve = True
    Exit Function
EH:
    AddCurve = False
End Function

Private Function AddShapeXY(xs As Variant, ys As Variant, ByVal lvName As String, ByVal color As Long, ByVal filled As Boolean) As Boolean
    Dim p As TPoly, i As Long
    PolyInit p, UBound(xs) + 2
    For i = 0 To UBound(xs): PolyAdd p, xs(i), ys(i): Next
    AddShapeXY = AddCurve(p, True, lvName, color, 1, filled)
End Function

'Centre-middle text: created with the active text style at height h, then moved so its range centre is (x, y).
Private Function AddText(ByVal s As String, ByVal x As Double, ByVal y As Double, ByVal h As Double, _
                         ByVal lvName As String, ByVal color As Long) As Boolean
    Dim ts As TextStyle, oh As Double, ow As Double, oj As Long, el As TextElement, rg As Range3d, d As Point3d
    On Error GoTo EH
    Set ts = ActiveSettings.TextStyle
    oh = ts.Height: ow = ts.Width: oj = ts.Justification
    ts.Height = h: ts.Width = h: ts.Justification = msdTextJustificationCenterCenter
    Set el = CreateTextElement1(Nothing, s, Point3dFromXYZ(x, y, 0), Matrix3dIdentity)
    ts.Height = oh: ts.Width = ow: ts.Justification = oj
    'TextStyle.Height units vary with the file's UORs, so size the text by its measured range.
    rg = el.Range
    If rg.High.Y - rg.Low.Y > 0 Then
        el.Transform Transform3dFromMatrix3dAndFixedPoint3d(Matrix3dFromScale((h * 0.72) / (rg.High.Y - rg.Low.Y)), Point3dFromXYZ(x, y, 0))
        rg = el.Range
    End If
    d.X = x - (rg.Low.X + rg.High.X) / 2#: d.Y = y - (rg.Low.Y + rg.High.Y) / 2#
    el.Move d
    Props el, lvName, color, -1
    ActiveModelReference.AddElement el
    AddText = True
    Exit Function
EH:
    On Error Resume Next
    If oh > 0 Then ts.Height = oh: ts.Width = ow: ts.Justification = oj
    AddText = False
End Function

'Scale bar (4 alternating segments) and true-north arrow below the bottom-right of the rings.
Public Function AddScaleAndNorth(rings() As TPoly, ByVal n As Long, ByVal lvName As String, desc As String) As Long
    Dim x0 As Double, y0 As Double, x1 As Double, y1 As Double, wdt As Double, L As Double, labels() As String
    Dim unitTxt As String, hb As Double, th As Double, margin As Double, bx As Double, by As Double, seg As Double
    Dim i As Long, cnt As Long, fr As Variant
    Const col As Long = 0
    PolysBounds rings, n, x0, y0, x1, y1
    wdt = x1 - x0
    L = ScaleBarSpec(wdt / 4#, gUnit, labels, unitTxt)
    hb = L / 22#
    th = hb * 1.3
    margin = IIf(wdt * 0.03 > hb * 3, wdt * 0.03, hb * 3)
    bx = x1 - L
    by = y0 - margin - hb
    seg = L / 4#
    For i = 0 To 3
        If AddShapeXY(Array(bx + i * seg, bx + (i + 1) * seg, bx + (i + 1) * seg, bx + i * seg), _
                      Array(by, by, by + hb, by + hb), lvName, col, (i Mod 2 = 0)) Then cnt = cnt + 1
    Next
    fr = Array(0#, 0.5, 1#)
    For i = 0 To 2
        If AddText(labels(i), bx + fr(i) * L, by + hb + th * 0.9, th, lvName, col) Then cnt = cnt + 1
    Next
    If AddText(unitTxt, bx + L / 2#, by - th * 0.9, th, lvName, col) Then cnt = cnt + 1

    Dim A As Double, cx As Double, cy As Double, dx As Double, dy As Double, lon As Double, lat As Double
    Dim wx As Double, wy As Double, vl As Double
    A = IIf(L * 0.22 > hb * 7, L * 0.22, hb * 7)
    cx = bx - A * 0.9: cy = by + hb / 2# + A * 0.25
    dx = 0: dy = 1
    On Error Resume Next
    WorldToLL cx, cy, lon, lat
    LLToWorld lon, lat + 0.01, wx, wy
    If Err.Number = 0 Then
        vl = Sqr((wx - cx) ^ 2 + (wy - cy) ^ 2)
        If vl > 0 Then dx = (wx - cx) / vl: dy = (wy - cy) / vl
    End If
    Err.Clear
    On Error GoTo 0
    Dim tipX As Double, tipY As Double, nX As Double, nY As Double, lX As Double, lY As Double, rX As Double, rY As Double
    NP cx, cy, dx, dy, 0, A / 2#, tipX, tipY
    NP cx, cy, dx, dy, 0, -A / 5#, nX, nY
    NP cx, cy, dx, dy, -A * 0.28, -A / 2#, lX, lY
    NP cx, cy, dx, dy, A * 0.28, -A / 2#, rX, rY
    If AddShapeXY(Array(tipX, nX, lX), Array(tipY, nY, lY), lvName, col, True) Then cnt = cnt + 1
    If AddShapeXY(Array(tipX, rX, nX), Array(tipY, rY, nY), lvName, col, False) Then cnt = cnt + 1
    Dim tx As Double, ty As Double
    NP cx, cy, dx, dy, 0, A / 2# + th * 1.4, tx, ty
    If AddText("N", tx, ty, th * 1.6, lvName, col) Then cnt = cnt + 1
    desc = labels(2) & " " & unitTxt
    AddScaleAndNorth = cnt
End Function

Private Sub NP(ByVal cx As Double, ByVal cy As Double, ByVal dx As Double, ByVal dy As Double, _
               ByVal u As Double, ByVal v As Double, x As Double, y As Double)
    x = cx + u * dy + v * dx
    y = cy - u * dx + v * dy
End Sub

'---------------------------------------------------------------- fence and view
Public Sub DefineMstnFence(ring As TPoly)
    Dim pts() As Point3d, i As Long, n As Long
    n = ring.N
    If n > 1 Then
        If ring.X(0) = ring.X(n - 1) And ring.Y(0) = ring.Y(n - 1) Then n = n - 1
    End If
    ReDim pts(0 To n - 1)
    For i = 0 To n - 1: pts(i).X = ring.X(i): pts(i).Y = ring.Y(i): Next
    ActiveDesignFile.Fence.DefineFromModelPoints 1, pts
End Sub

Public Sub FitMstnView()
    On Error Resume Next
    CadInputQueue.SendKeyin "FIT VIEW EXTENDED"
End Sub
