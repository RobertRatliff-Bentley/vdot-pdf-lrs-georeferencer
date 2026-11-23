Attribute VB_Name = "modMain"
'VDOT Project Mapper - native MicroStation VBA edition (no Python required).
'Run with:  vba run [VDOT_Project_Mapper]modMain.Main
Option Explicit

#If VBA7 Then
Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare PtrSafe Function GetTickCount Lib "kernel32" () As Long
Public Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As LongPtr)
#Else
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function GetTickCount Lib "kernel32" () As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
#End If

Public Const APP_TITLE As String = "VDOT Project Mapper"
Public Const APP_VERSION As String = "1.0"
Public Const REG_APP As String = "VDOT_Georef_Native"
Public Const PREVIEW_DPI As Long = 48
Public Const MASK_IMAGE_DPI As Double = 200
Public Const MASK_IMAGE_MAX_PX As Double = 10000
Public Const KEY_COLOR As Long = &H1FF01          ' colour-key used for transparent pixels (R=1, G=255, B=1)

'---------------------------------------------------------------- route groups
Public Const NGRP As Long = 7
Public Const VDOT_MAPS_URL As String = "https://www.vdot.virginia.gov/travel-traffic/maps/"
Public GrpName(0 To 6) As String, GrpCats(0 To 6) As Variant, GrpOn(0 To 6) As Boolean
Public GrpColor(0 To 6) As Long, GrpWeight(0 To 6) As Long, GrpPrev(0 To 6) As Long, GrpPx(0 To 6) As Double

'---------------------------------------------------------------- state
Public gPdfPath As String, gHasPdf As Boolean
Public gPdfWpt As Double, gPdfHpt As Double
Public gDisp() As Byte, DW As Long, DH As Long, gDispK As Double  ' display image (BGR, top-down, stride DW*3)

Public gJuris As Collection
Public gCurJuris As Object
Public gRingsLL() As TPoly, gRingsW() As TPoly, gRingsN() As TPoly, gNR As Long
Public gNormCx As Double, gNormCy As Double, gNormS As Double
Public gRoutesLL() As TPoly, gRoutesW() As TPoly, gRoutesN() As TPoly, gNRt As Long, gRtGrp() As Long

Public gG() As Double, gHasG As Boolean
Public gGAuto() As Double, gHasGAuto As Boolean
Public gPairWX() As Double, gPairWY() As Double, gPairPX() As Double, gPairPY() As Double, gNPairs As Long

Public gHasFence As Boolean, gFx0 As Double, gFy0 As Double, gFx1 As Double, gFy1 As Double

Public gGcs As GeographicCoordinateSystem
Public gHasGcs As Boolean, gGcsLabel As String, gOutSR As Long, gUnit As String

Public gTmpDir As String

'---------------------------------------------------------------- unicode helpers (source stays ASCII)
Public Function UDot() As String: UDot = ChrW(183): End Function
Public Function UEll() As String: UEll = ChrW(8230): End Function
Public Function UTimes() As String: UTimes = ChrW(215): End Function
Public Function UArrow() As String: UArrow = ChrW(8594): End Function
Public Function UDash() As String: UDash = ChrW(8211): End Function
Public Function UApprox() As String: UApprox = ChrW(8776): End Function
Public Function AppCredit() As String
    AppCredit = "Created by Robert Ratliff " & UDot() & " Build 2026-11-19"
End Function

Public Sub InitGroups()
    Dim i As Long
    GrpName(0) = "Interstate": GrpCats(0) = Array("Interstate"): GrpOn(0) = True
    GrpColor(0) = 3: GrpWeight(0) = 4: GrpPrev(0) = HexRGB("#ef4444")
    GrpName(1) = "US Primary": GrpCats(1) = Array("US Highway Primary"): GrpOn(1) = True
    GrpColor(1) = 1: GrpWeight(1) = 3: GrpPrev(1) = HexRGB("#3b82f6")
    GrpName(2) = "State Primary": GrpCats(2) = Array("State Highway Primary"): GrpOn(2) = True
    GrpColor(2) = 4: GrpWeight(2) = 3: GrpPrev(2) = HexRGB("#eab308")
    GrpName(3) = "Secondary": GrpCats(3) = Array("Secondary"): GrpOn(3) = True
    GrpColor(3) = 2: GrpWeight(3) = 1: GrpPrev(3) = HexRGB("#22c55e")
    GrpName(4) = "Ramps & Frontage"
    GrpCats(4) = Array("Interstate Ramp", "Interstate Frontage Road", "Non-Interstate Ramp", _
                       "Non-Interstate Frontage Road", "Street IS Ramp", "Street Non-IS Ramp")
    GrpOn(4) = False: GrpColor(4) = 6: GrpWeight(4) = 1: GrpPrev(4) = HexRGB("#f97316")
    GrpName(5) = "Urban & Streets": GrpCats(5) = Array("Urban", "Urban Streets", "Streets", "County Roads")
    GrpOn(5) = False: GrpColor(5) = 5: GrpWeight(5) = 1: GrpPrev(5) = HexRGB("#a855f7")
    GrpName(6) = "Other"
    GrpCats(6) = Array("Federal Lands Roads", "Institutional Property", "Private Property", "School Roads", "Shared Use Path")
    GrpOn(6) = False: GrpColor(6) = 8: GrpWeight(6) = 0: GrpPrev(6) = HexRGB("#94a3b8")
    For i = 0 To 6
        If i <= 2 Then GrpPx(i) = 2# Else GrpPx(i) = 1.2
    Next
End Sub

Public Function GroupOf(ByVal cat As String) As Long
    Dim i As Long, c As Variant
    For i = 0 To NGRP - 1
        For Each c In GrpCats(i)
            If c = cat Then GroupOf = i: Exit Function
        Next
    Next
    GroupOf = NGRP - 1
End Function

'Per-class route level: <routes level>_<class>, e.g. VDOT_LRS_Routes_US_Primary.
Public Function RouteLevelName(ByVal prefix As String, ByVal g As Long) As String
    Dim s As String
    s = Replace(Replace(GrpName(g), " & ", "_"), " ", "_")
    RouteLevelName = prefix & "_" & s
End Function

'#rrggbb -> VBA colour (&HBBGGRR)
Public Function HexRGB(ByVal h As String) As Long
    h = Replace(h, "#", "")
    HexRGB = RGB(CLng("&H" & Mid$(h, 1, 2)), CLng("&H" & Mid$(h, 3, 2)), CLng("&H" & Mid$(h, 5, 2)))
End Function

'---------------------------------------------------------------- entry point
Public Sub Main()
    On Error GoTo EH
    InitGroups
    gTmpDir = Environ$("TEMP") & "\vdot_georef_native"
    If Dir$(gTmpDir, vbDirectory) = "" Then MkDir gTmpDir
    InitGcs
    If FormIsOpen() Then
        frmGeoref.Show vbModeless
        Exit Sub
    End If
    frmGeoref.Show vbModeless
    DoLoadJurisdictions
    Exit Sub
EH:
    MsgBox "Could not start " & APP_TITLE & ":" & vbCrLf & Err.Description, vbExclamation, APP_TITLE
End Sub

Public Sub InitGcs()
    Dim nm As String, dn As String
    gHasGcs = False: gGcsLabel = "No GCS": gOutSR = 4269: gUnit = "units"
    On Error Resume Next
    Set gGcs = Nothing
    Set gGcs = ActiveModelReference.GetGCS(True)
    If Not gGcs Is Nothing Then
        nm = gGcs.Name
        dn = UCase$(Replace(gGcs.DatumName, "-", ""))
        If InStr(dn, "WGS84") > 0 Then gOutSR = 4326
        gHasGcs = True
        gGcsLabel = IIf(Len(nm) > 0, nm, "GCS")
    End If
    gUnit = ActiveModelReference.MasterUnit.Label
    If gUnit = "'" Then gUnit = "ft"
    If gUnit = """" Then gUnit = "in"
    If Len(gUnit) = 0 Then gUnit = "units"
    Err.Clear
End Sub

'---------------------------------------------------------------- coordinate conversion
Public Sub LLToWorld(ByVal lon As Double, ByVal lat As Double, x As Double, y As Double)
    Dim gp As GeoPoint2D, p As Point2d
    gp.Longitude = lon: gp.Latitude = lat
    p = gGcs.MasterUnitsFromLatLong2d(gp)
    x = p.X: y = p.Y
End Sub

Public Sub WorldToLL(ByVal x As Double, ByVal y As Double, lon As Double, lat As Double)
    Dim gp As GeoPoint2D, p As Point2d
    p.X = x: p.Y = y
    gp = gGcs.LatLongFromMasterUnits2d(p)
    lon = gp.Longitude: lat = gp.Latitude
End Sub

Public Sub PolysToWorld(src() As TPoly, ByVal n As Long, dst() As TPoly)
    Dim k As Long, i As Long, x As Double, y As Double
    If n = 0 Then ReDim dst(0 To 0): Exit Sub
    ReDim dst(0 To n - 1)
    For k = 0 To n - 1
        PolyInit dst(k), src(k).N
        For i = 0 To src(k).N - 1
            LLToWorld src(k).X(i), src(k).Y(i), x, y
            PolyAdd dst(k), x, y
        Next
        dst(k).Cat = src(k).Cat: dst(k).Nm = src(k).Nm
        If k Mod 200 = 0 Then DoEvents
    Next
End Sub

Public Sub PolysToNorm(src() As TPoly, ByVal n As Long, dst() As TPoly)
    Dim k As Long, i As Long
    If n = 0 Then ReDim dst(0 To 0): Exit Sub
    ReDim dst(0 To n - 1)
    For k = 0 To n - 1
        PolyInit dst(k), src(k).N
        For i = 0 To src(k).N - 1
            PolyAdd dst(k), (src(k).X(i) - gNormCx) / gNormS, (src(k).Y(i) - gNormCy) / gNormS
        Next
        dst(k).Cat = src(k).Cat: dst(k).Nm = src(k).Nm
    Next
End Sub

'preview pixel -> world (master units); needs gG
Public Sub PreviewToWorld(ByVal px As Double, ByVal py As Double, x As Double, y As Double)
    Dim inv() As Double, nx As Double, ny As Double
    InvertG gG, inv
    ApplyG inv, px, py, nx, ny
    x = nx * gNormS + gNormCx
    y = ny * gNormS + gNormCy
End Sub

'---------------------------------------------------------------- jurisdictions
Public Sub DoLoadJurisdictions()
    On Error GoTo EH
    BusyOn "Loading Virginia jurisdictions from VDOT" & UEll()
    Set gJuris = ListJurisdictions()
    Dim nJ As Long, dist As Collection, dj As Variant
    nJ = gJuris.Count
    On Error Resume Next
    Set dist = ListDistricts()
    If Err.Number <> 0 Then LogMsg "Could not load VDOT districts: " & Err.Description, "warn": Set dist = New Collection
    Err.Clear
    On Error GoTo EH
    For Each dj In dist
        gJuris.Add dj
    Next
    SetCap "lblJuris", nJ & " jurisdictions, " & dist.Count & " districts available"
    LogMsg "Loaded " & nJ & " jurisdictions and " & dist.Count & " VDOT districts.", "ok"
    FilterJuris
    BusyOff
    If gHasPdf Then AutodetectJuris
    Exit Sub
EH:
    BusyOff
    SetCap "lblJuris", "Could not load jurisdictions " & UDash() & " check the network connection."
    LogMsg Err.Description, "err"
End Sub

Public Function GuessJurisdiction(ByVal pdfPath As String) As Object
    Dim stem As String, i As Long, ch As String, digits As String, words As Collection, w As String
    Dim text As String, j As Variant, hits As Collection, named As Collection, code As String
    stem = FileTitle(pdfPath)
    If InStrRev(stem, ".") > 0 Then stem = Left$(stem, InStrRev(stem, ".") - 1)
    Set words = New Collection
    For i = 1 To Len(stem)
        ch = Mid$(stem, i, 1)
        If ch Like "[A-Za-z]" Then
            w = w & LCase$(ch)
        ElseIf Len(w) Then
            words.Add w: w = ""
        End If
    Next
    If Len(w) Then words.Add w
    text = ""
    For i = 1 To words.Count
        text = text & IIf(i > 1, " ", "") & words(i)
    Next
    i = 1
    Do While i <= Len(stem)
        If Mid$(stem, i, 1) <> " " Then Exit Do
        i = i + 1
    Loop
    Do While i <= Len(stem)
        ch = Mid$(stem, i, 1)
        If ch Like "#" Then digits = digits & ch Else Exit Do
        i = i + 1
    Loop
    If Len(digits) >= 1 And Len(digits) <= 3 Then
        code = Right$("000" & digits, 3)
        Set hits = New Collection
        For Each j In gJuris
            If Right$("000" & CStr(DGet(j, "VDOT_CODE", "")), 3) = code And Len(CStr(DGet(j, "VDOT_CODE", ""))) > 0 Then hits.Add j
        Next
        If hits.Count > 1 And words.Count > 0 Then
            Set named = New Collection
            For Each j In hits
                If Len(DGet(j, "NAME", "")) > 0 Then
                    If InStr(1, text, LCase$(CStr(DGet(j, "NAME", ""))), vbBinaryCompare) > 0 Then named.Add j
                End If
            Next
            If named.Count > 0 Then Set hits = named
        End If
        If hits.Count > 0 Then
            For Each j In hits
                If DGet(j, "JURISTYPE", "") = "CO" Then Set GuessJurisdiction = j: Exit Function
            Next
            Set GuessJurisdiction = hits(1)
            Exit Function
        End If
    End If
    If words.Count > 0 Then
        Dim isCity As Boolean, isDist As Boolean, best As Object, bestScore As Long, sc As Long, nm As String
        For i = 1 To words.Count
            If words(i) = "city" Then isCity = True
            If words(i) = "district" Then isDist = True
        Next
        bestScore = -1
        For Each j In gJuris
            nm = LCase$(CStr(DGet(j, "NAME", "")))
            If CStr(DGet(j, "JURISTYPE", "")) = "DI" And Not isDist Then nm = ""
            If Len(nm) > 0 Then
                If WholeWord(" " & text & " ", nm) Then
                    sc = Len(nm)
                    If (DGet(j, "JURISTYPE", "") = "CI") = isCity Then sc = sc + 5
                    If DGet(j, "JURISTYPE", "") = "CO" Then sc = sc + 2
                    If DGet(j, "JURISTYPE", "") = "DI" Then sc = sc + 20
                    If sc > bestScore Then bestScore = sc: Set best = j
                End If
            End If
        Next
        If Not best Is Nothing Then Set GuessJurisdiction = best
    End If
End Function

Private Function WholeWord(ByVal padded As String, ByVal nm As String) As Boolean
    Dim p As Long, b As String, a As String
    p = InStr(1, padded, nm, vbBinaryCompare)
    Do While p > 0
        b = Mid$(padded, p - 1, 1): a = Mid$(padded, p + Len(nm), 1)
        If Not (b Like "[a-z0-9]") And Not (a Like "[a-z0-9]") Then WholeWord = True: Exit Function
        p = InStr(p + 1, padded, nm, vbBinaryCompare)
    Loop
End Function

Public Function AutodetectJuris() As Boolean
    Dim j As Object
    If gJuris Is Nothing Then Exit Function
    Set j = GuessJurisdiction(gPdfPath)
    If j Is Nothing Then
        SetCap "lblJuris", "Could not detect jurisdiction from the file name " & UDash() & " pick it from the list."
        Exit Function
    End If
    SelectJuris j
    SetCap "lblJuris", "Auto-detected from file name: " & DGet(j, "NAMELSAD", "")
    LogMsg "Detected jurisdiction: " & DGet(j, "NAMELSAD", "") & IIf(Len(CStr(DGet(j, "VDOT_CODE", ""))) > 0, " (VDOT code " & DGet(j, "VDOT_CODE", "") & ")", ""), "ok"
    Dim same As Boolean
    If Not gCurJuris Is Nothing Then same = (CLng(DGet(gCurJuris, "OBJECTID", -1)) = CLng(DGet(j, "OBJECTID", -2)))
    If Not same Then DoLoadBoundary
    AutodetectJuris = True
End Function

'---------------------------------------------------------------- boundary
Public Sub DoLoadBoundary()
    Dim j As Object, rings() As TPoly, n As Long, i As Long, nv As Long
    Dim x0 As Double, y0 As Double, x1 As Double, y1 As Double
    On Error GoTo EH
    Set j = SelectedJuris()
    If j Is Nothing Then Exit Sub
    If Not gHasGcs Then
        LogMsg "The active model has no geographic coordinate system. Assign one (e.g. VA83-SF) first.", "err"
        Exit Sub
    End If
    BusyOn "Downloading boundary for " & DGet(j, "NAMELSAD", "") & UEll()
    n = GetBoundary(CLng(DGet(j, "OBJECTID", 0)), gOutSR, rings)
    If n = 0 Then Err.Raise vbObjectError + 610, , "The boundary has no usable rings."
    Set gCurJuris = j
    gRingsLL = rings: gNR = n
    PolysToWorld gRingsLL, gNR, gRingsW
    PolysBounds gRingsW, gNR, x0, y0, x1, y1
    gNormCx = (x0 + x1) / 2#: gNormCy = (y0 + y1) / 2#
    gNormS = IIf(x1 - x0 > y1 - y0, x1 - x0, y1 - y0)
    If gNormS <= 0 Then gNormS = 1
    PolysToNorm gRingsW, gNR, gRingsN
    gNRt = 0
    SetCap "lblRoutes", "Not loaded"
    CardClear 3
    gHasG = False: gHasGAuto = False: gNPairs = 0
    CancelPending
    For i = 0 To n - 1: nv = nv + gRingsLL(i).N: Next
    CardDone 2, CStr(DGet(j, "NAME", "")), True
    LogMsg "Boundary loaded: " & n & " ring(s), " & nv & " vertices.", "ok"
    Dim o() As TPoly, nO As Long, h() As TPoly, nH As Long
    SplitRings gRingsW, gNR, o, nO, h, nH
    SetTip "chkHoles", "Exclude interior holes of the boundary (e.g. independent cities). " & _
        IIf(nH > 0, "This boundary has " & nH & " hole(s).", "This boundary has no holes.")
    BusyOff
    UpdateFitLabel
    ViewRedraw True
    If gHasPdf And gNR > 0 Then
        If Not TryRestoreAlignment() Then DoAutoAlign True
    End If
    Exit Sub
EH:
    BusyOff
    LogMsg Err.Description, "err"
End Sub

'---------------------------------------------------------------- routes
Public Function DoFetchRoutes() As Boolean
    Dim cats As New Collection, i As Long, c As Variant, x0 As Double, y0 As Double, x1 As Double, y1 As Double
    Dim raw() As TPoly, nRaw As Long, fetched As Long, clipped() As TPoly, nC As Long, nv As Long
    On Error GoTo EH
    If gNR = 0 Then Exit Function
    For i = 0 To NGRP - 1
        If Chk("chkGrp" & i) Then
            For Each c In GrpCats(i): cats.Add c: Next
        End If
    Next
    If cats.Count = 0 Then LogMsg "Select at least one route category.", "warn": Exit Function
    PolysBounds gRingsLL, gNR, x0, y0, x1, y1
    BusyOn "Downloading LRS routes" & UEll()
    nRaw = GetRoutes(x0 - 0.002, y0 - 0.002, x1 + 0.002, y1 + 0.002, cats, gOutSR, raw, fetched)
    If Chk("chkClipRoutes") And nRaw > 0 Then
        SetStatus "Clipping routes to the boundary" & UEll()
        ClipRoutes raw, nRaw, gRingsLL, gNR, clipped, nC
    Else
        clipped = raw: nC = nRaw
    End If
    LogMsg "Projecting " & nC & " route parts" & UEll()
    gRoutesLL = clipped: gNRt = nC
    PolysToWorld gRoutesLL, gNRt, gRoutesW
    PolysToNorm gRoutesW, gNRt, gRoutesN
    If gNRt > 0 Then ReDim gRtGrp(0 To gNRt - 1)
    For i = 0 To gNRt - 1
        gRtGrp(i) = GroupOf(gRoutesLL(i).Cat)
        nv = nv + gRoutesLL(i).N
    Next
    SetCap "lblRoutes", gNRt & " route parts " & UDot() & " " & nv & " vertices (" & fetched & " features fetched)"
    CardDone 3, CStr(gNRt), True
    LogMsg "LRS routes ready: " & gNRt & " parts from " & nRaw & " paths.", "ok"
    BusyOff
    ViewRedraw True
    DoFetchRoutes = True
    Exit Function
EH:
    BusyOff
    LogMsg Err.Description, "err"
End Function

'---------------------------------------------------------------- PDF
Public Sub DoBrowsePdf()
    Dim p As String, start As String
    If Len(gPdfPath) Then start = Left$(gPdfPath, InStrRev(gPdfPath, "\") - 1) Else start = GetSetting(REG_APP, "ui", "lastdir", "")
    p = BrowseForPdf(start)
    If Len(p) Then
        SaveSetting REG_APP, "ui", "lastdir", Left$(p, InStrRev(p, "\") - 1)
        DoLoadPdf p
    End If
End Sub

Public Sub DoLoadPdf(ByVal path As String)
    Dim newFile As Boolean
    On Error GoTo EH
    BusyOn "Rendering preview of " & FileTitle(path) & UEll()
    RenderPreview path
    newFile = (LCase$(path) <> LCase$(gPdfPath))
    gPdfPath = path: gHasPdf = True
    SetCap "txtPdf", path
    SetTip "txtPdf", path
    SetCap "lblPdf", "Page 1 " & UDot() & " " & Format$(gPdfWpt / 72#, "0.0") & " " & UTimes() & " " & Format$(gPdfHpt / 72#, "0.0") & " in"
    SetChk "chkRegion", False
    CardDone 1, FileTitle(path), True
    LogMsg "PDF loaded: " & FileTitle(path), "ok"
    If newFile Then
        gHasG = False: gHasGAuto = False: gNPairs = 0
        SetCap "lblOut", ""
        CancelPending
        ClearFence
    End If
    BusyOff
    ViewFit
    UpdateFitLabel
    If Not gJuris Is Nothing And newFile Then
        If AutodetectJuris() Then Exit Sub
    End If
    If gNR > 0 And Not gHasG Then
        If Not TryRestoreAlignment() Then DoAutoAlign True
    End If
    Exit Sub
EH:
    BusyOff
    LogMsg Err.Description, "err"
End Sub

'---------------------------------------------------------------- alignment
Public Sub SetG(G() As Double, Optional ByVal label As String = "")
    gG = G: gHasG = True
    UpdateState
    UpdateFitLabel label
    ViewRedraw False
End Sub

Public Sub UpdateFitLabel(Optional ByVal prefix As String = "")
    Dim meanD As Double, frac As Double, det As Double, ground As Double, t As String
    If Not gHasG Or Not gHasPdf Or gNR = 0 Then
        SetCap "lblFit", "No alignment yet"
        Exit Sub
    End If
    ScoreFit gG, gRingsN, gNR, meanD, frac
    det = Abs(gG(0) * gG(4) - gG(1) * gG(3))
    If det < 1E-300 Then det = 1E-300
    ground = gNormS / Sqr(det) * meanD
    t = IIf(Len(prefix), prefix & " " & UDot() & " ", "") & "Fit: " & Format$(meanD, "0.0") & " px mean " & UDot() & " " & _
        CLng(frac * 100) & "% on linework " & UDot() & " " & UApprox() & " " & Format$(ground, "0") & " " & gUnit
    SetCap "lblFit", t
    CardDone 4, IIf(frac > 0.6, "aligned", "check fit"), frac > 0.6
End Sub

Public Sub DoAutoAlign(ByVal full As Boolean)
    Dim G() As Double, affine As Boolean, useR As Boolean, t0 As Long, dummy() As TPoly
    On Error GoTo EH
    If Not gHasPdf Or gNR = 0 Then Exit Sub
    affine = Opt("optAff")
    useR = Chk("chkUseRoutes") And gNRt > 0
    BusyOn IIf(full, "Auto-aligning PDF to boundary" & UEll(), "Refining alignment" & UEll())
    t0 = GetTickCount
    If full Or Not gHasG Then
        InitialFit gRingsN, gNR, G
        SetG G
        DoEvents
        If useR Then Refine G, gRingsN, gNR, gRoutesN, gNRt, affine, True Else Refine G, gRingsN, gNR, dummy, 0, affine, True
    Else
        G = gG
        If useR Then Refine G, gRingsN, gNR, gRoutesN, gNRt, affine, False Else Refine G, gRingsN, gNR, dummy, 0, affine, False
    End If
    gGAuto = G: gHasGAuto = True
    SetG G, IIf(full, "Auto", "Refined")
    LogMsg "Alignment complete. " & GetCap("lblFit"), "ok"
    Debug.Print "align ms " & (GetTickCount - t0)
    BusyOff
    ViewFit
    Exit Sub
EH:
    BusyOff
    LogMsg Err.Description, "err"
End Sub

Public Sub DoResetAlignment()
    If gHasGAuto Then
        SetG gGAuto, "Auto"
    Else
        DoAutoAlign True
    End If
End Sub

'compose a preview-pixel transform M (a,b,c,d,e,f: x' = a x + b y + c ; y' = d x + e y + f) onto G
Public Sub ComposeG(ByVal a As Double, ByVal b As Double, ByVal c As Double, ByVal d As Double, ByVal e As Double, ByVal f As Double)
    Dim n(0 To 5) As Double, i As Long
    If Not gHasG Then Exit Sub
    n(0) = a * gG(0) + b * gG(3): n(1) = a * gG(1) + b * gG(4): n(2) = a * gG(2) + b * gG(5) + c
    n(3) = d * gG(0) + e * gG(3): n(4) = d * gG(1) + e * gG(4): n(5) = d * gG(2) + e * gG(5) + f
    For i = 0 To 5: gG(i) = n(i): Next
End Sub

'---------------------------------------------------------------- tie points
Public Sub AddPair(ByVal wx As Double, ByVal wy As Double, ByVal px As Double, ByVal py As Double)
    Dim affine As Boolean, need As Long, G() As Double, rms As Double
    ReDim Preserve gPairWX(0 To gNPairs): ReDim Preserve gPairWY(0 To gNPairs)
    ReDim Preserve gPairPX(0 To gNPairs): ReDim Preserve gPairPY(0 To gNPairs)
    gPairWX(gNPairs) = wx: gPairWY(gNPairs) = wy: gPairPX(gNPairs) = px: gPairPY(gNPairs) = py
    gNPairs = gNPairs + 1
    affine = Opt("optAff")
    need = IIf(affine, 3, 2)
    If gNPairs >= need Then
        rms = SolveFromPairs(gPairWX, gPairWY, gPairPX, gPairPY, gNPairs, affine, G)
        If rms >= 0 Then
            SetG G, "Tie points (RMS " & Format$(rms, "0.0") & " px)"
            LogMsg "Solved from " & gNPairs & " tie points, RMS " & Format$(rms, "0.00") & " px.", "ok"
        End If
    Else
        Toast "Tie point " & gNPairs & " added " & UDash() & " " & (need - gNPairs) & " more needed."
        ViewRedraw False
    End If
    UpdateState
End Sub

Public Sub DoClearPins()
    gNPairs = 0
    CancelPending
    ViewRedraw False
    UpdateState
End Sub

'---------------------------------------------------------------- block fence
Public Sub OnFenceDrawn(ByVal x0 As Double, ByVal y0 As Double, ByVal x1 As Double, ByVal y1 As Double)
    gHasFence = True: gFx0 = x0: gFy0 = y0: gFx1 = x1: gFy1 = y1
    FenceChanged
    LogMsg "Block fence set: " & Format$((x1 - x0) / PREVIEW_DPI, "0.0") & " " & UTimes() & " " & _
        Format$((y1 - y0) / PREVIEW_DPI, "0.0") & " in on the PDF. The image, routes and scale bar will be clipped to it.", "ok"
    UpdateState
End Sub

Public Sub ClearFence()
    Dim had As Boolean
    had = gHasFence
    gHasFence = False
    FenceChanged
    ViewRedraw False
    If had Then LogMsg "Block fence cleared " & UDash() & " clipping to the county boundary."
    UpdateState
End Sub

'world ring of the fence (closed, 5 points); returns False when not available
Public Function FenceWorld(ring As TPoly) As Boolean
    Dim xs As Variant, ys As Variant, i As Long, x As Double, y As Double
    If Not gHasFence Or Not gHasG Then Exit Function
    xs = Array(gFx0, gFx1, gFx1, gFx0, gFx0)
    ys = Array(gFy1, gFy1, gFy0, gFy0, gFy1)
    PolyInit ring, 5
    For i = 0 To 4
        PreviewToWorld xs(i), ys(i), x, y
        PolyAdd ring, x, y
    Next
    FenceWorld = True
End Function

'---------------------------------------------------------------- sidecar
Public Function SidecarPath() As String
    If Len(gPdfPath) = 0 Then Exit Function
    SidecarPath = Left$(gPdfPath, InStrRev(gPdfPath, ".") - 1) & ".georef.json"
End Function

Public Sub SaveAlignment()
    Dim k As Double, i As Long, j As Long, u As Double, v As Double, x As Double, y As Double
    Dim lon As Double, lat As Double, s As String, f As Integer, arr As String
    On Error GoTo EH
    k = PREVIEW_DPI / 72#
    For i = 0 To 3
        For j = 0 To 3
            u = gPdfWpt * i / 3#: v = gPdfHpt * j / 3#
            PreviewToWorld u * k, v * k, x, y
            WorldToLL x, y, lon, lat
            arr = arr & IIf(Len(arr), "," & vbLf & "  ", "") & "[" & JNum(u) & ", " & JNum(v) & ", " & JNum(lon) & ", " & JNum(lat) & "]"
        Next
    Next
    s = "{" & vbLf & " ""pdf"": " & JStr(FileTitle(gPdfPath)) & "," & vbLf & " ""page"": 0," & vbLf & _
        " ""jurisdiction"": " & JDictFlat(gCurJuris) & "," & vbLf & _
        " ""affine"": " & IIf(Opt("optAff"), "true", "false") & "," & vbLf & _
        " ""page_pt_to_lonlat"": [" & vbLf & "  " & arr & vbLf & " ]" & vbLf & "}"
    f = FreeFile
    Open SidecarPath() For Output As #f
    Print #f, s;
    Close #f
    Exit Sub
EH:
    LogMsg "Could not save alignment sidecar: " & Err.Description, "warn"
End Sub

Public Function TryRestoreAlignment() As Boolean
    Dim sp As String, by() As Byte, f As Integer, js As Object, arr As Variant, n As Long, i As Long
    Dim wx() As Double, wy() As Double, px() As Double, py() As Double, x As Double, y As Double
    Dim G() As Double, rms As Double, jj As Object
    sp = SidecarPath()
    If Len(sp) = 0 Or gCurJuris Is Nothing Then Exit Function
    If Dir$(sp) = "" Then Exit Function
    On Error GoTo EH
    f = FreeFile
    Open sp For Binary Access Read As #f
    If LOF(f) = 0 Then Close #f: Exit Function
    ReDim by(0 To LOF(f) - 1)
    Get #f, , by
    Close #f
    Set js = JsonParseBytes(by)
    Set jj = js("jurisdiction")
    If CLng(DGet(jj, "OBJECTID", -1)) <> CLng(DGet(gCurJuris, "OBJECTID", -2)) Then Exit Function
    If CLng(DGet(js, "page", 0)) <> 0 Then Exit Function
    arr = js("page_pt_to_lonlat")
    n = UBound(arr) + 1
    If n < 3 Then Exit Function
    ReDim wx(0 To n - 1): ReDim wy(0 To n - 1): ReDim px(0 To n - 1): ReDim py(0 To n - 1)
    For i = 0 To n - 1
        px(i) = arr(i)(0) * PREVIEW_DPI / 72#: py(i) = arr(i)(1) * PREVIEW_DPI / 72#
        LLToWorld arr(i)(2), arr(i)(3), x, y
        wx(i) = (x - gNormCx) / gNormS: wy(i) = (y - gNormCy) / gNormS
    Next
    rms = SolveFromPairs(wx, wy, px, py, n, True, G)
    If rms < 0 Then Exit Function
    gGAuto = G: gHasGAuto = True
    SetG G, "Restored"
    LogMsg "Restored saved alignment from " & FileTitle(sp) & " (residual " & Format$(rms, "0.00") & " px).", "ok"
    ViewFit
    TryRestoreAlignment = True
    Exit Function
EH:
    LogMsg "Ignoring alignment sidecar: " & Err.Description, "warn"
End Function

'---------------------------------------------------------------- placement
Public Sub DoPlace()
    Dim fence As TPoly, hasF As Boolean, doRaster As Boolean, lv(0 To 3) As String, i As Long, nEl As Long
    Dim outers() As TPoly, nO As Long, holes() As TPoly, nH As Long, cnt As Long
    Dim defs As Variant
    On Error GoTo EH
    If gHasFence And Not gHasG Then LogMsg "Align the PDF before placing with a block fence.", "warn": Exit Sub
    hasF = FenceWorld(fence)
    doRaster = Chk("chkPlaceRaster")
    If doRaster And (Not gHasG Or Not gHasPdf) Then LogMsg "Load a PDF and align it first.", "warn": Exit Sub
    If Chk("chkPlaceRoutes") And gNRt = 0 Then
        If Not DoFetchRoutes() Then Exit Sub
    End If
    BusyOn "Placing in MicroStation" & UEll(), 4
    defs = Array("VDOT_LRS_Boundary", "VDOT_LRS_Routes", "VDOT_PDF_Map", "VDOT_Map_Annotation")
    For i = 0 To 3
        lv(i) = Trim$(GetCap("txtLv" & i))
        If Len(lv(i)) = 0 Then lv(i) = defs(i)
    Next
    SplitRings gRingsW, gNR, outers, nO, holes, nH
    If Chk("chkPlaceBnd") Then
        For i = 0 To gNR - 1
            If AddCurve(gRingsW(i), True, lv(0), 3, 4) Then nEl = nEl + 1
        Next
        LogMsg "Placed " & nEl & " boundary shape(s) on level " & lv(0) & ".", "ok"
    End If
    SetProgress 1
    If Chk("chkPlaceRoutes") And gNRt > 0 Then
        Dim rts() As TPoly, nR As Long, g As Long
        If hasF Then
            Dim fr(0 To 0) As TPoly, frr() As TPoly
            ReDim frr(0 To 0): frr(0) = fence
            ClipRoutes gRoutesW, gNRt, frr, 1, rts, nR
        Else
            rts = gRoutesW: nR = gNRt
        End If
        Dim gl(0 To NGRP - 1) As String, gc(0 To NGRP - 1) As Long, summ As String
        For g = 0 To NGRP - 1
            gl(g) = RouteLevelName(lv(1), g)
        Next
        For i = 0 To nR - 1
            g = GroupOf(rts(i).Cat)
            If gc(g) = 0 Then EnsureLevel gl(g), GrpColor(g), GrpWeight(g), "VDOT LRS " & GrpName(g) & " routes"
            If AddCurve(rts(i), False, gl(g), GrpColor(g), GrpWeight(g)) Then cnt = cnt + 1: gc(g) = gc(g) + 1
            If i Mod 250 = 0 Then DoEvents
        Next
        For g = 0 To NGRP - 1
            If gc(g) > 0 Then summ = summ & vbLf & "    " & gl(g) & ": " & gc(g)
        Next
        LogMsg "Placed " & cnt & " LRS route element(s) on per-class levels:" & summ, "ok"
    End If
    SetProgress 2
    If doRaster Then
        Dim mask() As TPoly, nM As Long, img As String, iw As Long, ih As Long
        nM = 0
        If Chk("chkClip") And hasF Then
            ReDim mask(0 To 0): mask(0) = fence: nM = 1
        ElseIf Chk("chkClip") Then
            ReDim mask(0 To nO + nH)
            For i = 0 To nO - 1: mask(nM) = outers(i): nM = nM + 1: Next
            If Chk("chkHoles") Then
                For i = 0 To nH - 1: mask(nM) = holes(i): nM = nM + 1: Next
                If nH = 0 Then LogMsg "Mask holes: this boundary has no interior holes, so nothing was masked.", "warn"
            End If
        End If
        SetProgress 3
        img = MakeMaskedImage(mask, nM, Chk("chkWhite"), iw, ih)
        AttachImage img, lv(2), CLng(GetVal("scrTransp")), True
        LogMsg "Map image attached: " & FileTitle(img) & " (" & iw & " " & UTimes() & " " & ih & " px)" & _
            IIf(nM > 0, ", everything outside the " & IIf(hasF, "fence", "boundary") & " is transparent", ""), "ok"
        If Chk("chkPdfOff") Then
            AttachPdfHidden lv(2)
            LogMsg "Original PDF attached, turned off in all views " & UDash() & " turn it on per view in Raster Manager.", "ok"
        End If
        SetCap "lblOut", "Image: " & img
        SaveAlignment
    End If
    If Chk("chkScaleNorth") And (hasF Or gNR > 0) Then
        Dim desc As String, ne As Long
        On Error Resume Next
        If hasF Then
            Dim fa(0 To 0) As TPoly, fb() As TPoly
            ReDim fb(0 To 0): fb(0) = fence
            ne = AddScaleAndNorth(fb, 1, lv(3), desc)
        Else
            ne = AddScaleAndNorth(gRingsW, gNR, lv(3), desc)
        End If
        If Err.Number <> 0 Then
            LogMsg "Scale bar / north arrow failed: " & Err.Description, "err"
        Else
            LogMsg "Placed scale bar (" & desc & ") and north arrow (" & ne & " elements) on level " & lv(3) & ".", "ok"
        End If
        Err.Clear
        On Error GoTo EH
    End If
    SetProgress 4
    If hasF Then
        On Error Resume Next
        DefineMstnFence fence
        If Err.Number = 0 Then LogMsg "Block fence also set as the active MicroStation fence.", "ok" Else LogMsg "Could not set the MicroStation fence: " & Err.Description, "warn"
        Err.Clear
        On Error GoTo EH
    End If
    FitMstnView
    CardDone 5, "placed", True
    LogMsg "Done " & UDash() & " elements placed in " & ActiveModelReference.Name & ".", "ok"
    BusyOff
    Exit Sub
EH:
    BusyOff
    LogMsg Err.Description, "err"
End Sub

Public Sub OpenVdotMaps()
    On Error Resume Next
    CreateObject("Shell.Application").ShellExecute VDOT_MAPS_URL
    If Err.Number <> 0 Then
        LogMsg "Could not open " & VDOT_MAPS_URL & ": " & Err.Description, "err"
    Else
        LogMsg "Opened the VDOT maps website in your browser.", "ok"
    End If
End Sub

'---------------------------------------------------------------- delete placed drawing
Public Sub DoDeletePlaced()
    Dim defs As Variant, names() As String, n As Long, i As Long, g As Long, s As String
    Dim els As New Collection, nRast As Long, nDel As Long, lst As String
    On Error GoTo EH
    defs = Array("VDOT_LRS_Boundary", "VDOT_LRS_Routes", "VDOT_PDF_Map", "VDOT_Map_Annotation")
    ReDim names(0 To 3 + NGRP)
    For i = 0 To 3
        s = Trim$(GetCap("txtLv" & i))
        If Len(s) = 0 Then s = defs(i)
        names(i) = s
    Next
    n = 4
    For g = 0 To NGRP - 1
        names(n) = RouteLevelName(names(1), g): n = n + 1
    Next
    CollectOnLevels names, n, els, nRast
    If els.Count = 0 Then
        LogMsg "Nothing to delete " & UDash() & " no elements on this tool's levels in the active model.", "warn"
        Exit Sub
    End If
    For i = 0 To n - 1: lst = lst & vbLf & "  " & names(i): Next
    If MsgBox("Delete " & (els.Count - nRast) & " element(s) from the active model and detach " & nRast & _
              " raster attachment(s)?" & vbLf & vbLf & "Levels:" & lst & vbLf & vbLf & _
              "Anything else you drew on these levels is deleted too. Use MicroStation Undo to reverse.", _
              vbYesNo + vbExclamation + vbDefaultButton2, APP_TITLE) <> vbYes Then Exit Sub
    BusyOn "Deleting placed drawing" & UEll(), 0
    nDel = RemoveElements(els)
    BusyOff
    If nDel = els.Count Then
        LogMsg "Deleted " & (nDel - nRast) & " element(s); " & nRast & " raster attachment(s) detached.", "ok"
    Else
        LogMsg "Deleted " & nDel & " of " & els.Count & " element(s); some could not be deleted.", "warn"
    End If
    CardDone 5, "deleted", False
    Exit Sub
EH:
    BusyOff
    LogMsg Err.Description, "err"
End Sub

'---------------------------------------------------------------- misc
Public Function FileTitle(ByVal p As String) As String
    FileTitle = Mid$(p, InStrRev(p, "\") + 1)
End Function
