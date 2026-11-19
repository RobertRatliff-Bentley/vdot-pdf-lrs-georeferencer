Attribute VB_Name = "modUI"
'Runtime-built modern UI for frmGeoref: header, step cards, preview viewbar, status bar and log.
'Mirrors the layout, texts and themes of the Python tool (vdot_pdf_lrs_clipper.py).
Option Explicit

Private Type OFNW
    lStructSize As Long
    hwndOwner As LongPtr
    hInstance As LongPtr
    lpstrFilter As LongPtr
    lpstrCustomFilter As LongPtr
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As LongPtr
    nMaxFile As Long
    lpstrFileTitle As LongPtr
    nMaxFileTitle As Long
    lpstrInitialDir As LongPtr
    lpstrTitle As LongPtr
    Flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As LongPtr
    lCustData As LongPtr
    lpfnHook As LongPtr
    lpTemplateName As LongPtr
    pvReserved As LongPtr
    dwReserved As Long
    FlagsEx As Long
End Type

Private Declare PtrSafe Function FindWindowW Lib "user32" (ByVal cls As LongPtr, ByVal title As LongPtr) As LongPtr
Private Declare PtrSafe Function GetWindowLongPtrW Lib "user32" (ByVal hwnd As LongPtr, ByVal idx As Long) As LongPtr
Private Declare PtrSafe Function SetWindowLongPtrW Lib "user32" (ByVal hwnd As LongPtr, ByVal idx As Long, ByVal v As LongPtr) As LongPtr
Private Declare PtrSafe Function DrawMenuBar Lib "user32" (ByVal hwnd As LongPtr) As Long
Private Declare PtrSafe Function SetWindowTextW Lib "user32" (ByVal hwnd As LongPtr, ByVal s As LongPtr) As Long
Private Declare PtrSafe Function GetSystemMetrics Lib "user32" (ByVal i As Long) As Long
Private Declare PtrSafe Function GetDC Lib "user32" (ByVal hwnd As LongPtr) As LongPtr
Private Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hwnd As LongPtr, ByVal hdc As LongPtr) As Long
Private Declare PtrSafe Function GetDeviceCaps Lib "gdi32" (ByVal hdc As LongPtr, ByVal nIndex As Long) As Long
Private Declare PtrSafe Function GetOpenFileNameW Lib "comdlg32" (ofn As OFNW) As Long

'layout (points)
Private Const HDR_H As Single = 62
Private Const SBW As Single = 336
Private Const CL As Single = 10
Private Const CWD As Single = 300
Private Const PAD As Single = 10
Private Const IW As Single = 280
Private Const VBAR_H As Single = 30
Private Const STAT_H As Single = 30
Private Const LOG_H As Single = 96
Private Const FONT_UI As String = "Segoe UI"
Private Const FONT_SYM As String = "Segoe UI Symbol"

Private mF As Object                 ' the UserForm
Private mHwnd As LongPtr
Private mCtl As Object               ' name -> control (caption label for buttons)
Private mRole As Object              ' name -> role
Private mBtnEn As Object             ' button name -> enabled
Private mBtnChk As Object            ' button name -> checked
Private mPh As Object                ' textbox name -> placeholder label name
Private mEvt As Collection
Private mPal As Object
Private mTheme As String
Private mHover As String
Private mBusy As Long, mSteps As Long, mBusyErr As Boolean
Private mSilent As Boolean
Private mLogOn As Boolean
Private mListJ As Collection
Private mCardOk(1 To 5) As Variant
Private mBuilding As Boolean

'================================================================ palette
Private Sub LoadPalette(ByVal theme As String)
    Dim s As String, kv As Variant, i As Long, p As Long
    If theme = "light" Then
        s = "text=#1f2937;bg=#f3f4f6;hdr1=#ffffff;border=#d6dae1;title=#111827;subtitle=#6b7280;" & _
            "chip_bg=#dcfce7;chip_fg=#166534;chip_bd=#86efac;warn_bg=#fef3c7;warn_fg=#92400e;warn_bd=#fcd34d;" & _
            "card=#ffffff;card_title=#111827;accent=#f28c28;on_accent=#111827;hint=#6b7280;metric=#1d4ed8;" & _
            "input_bg=#ffffff;input_bd=#cbd2dc;sel_bg=#fde7cf;sel_fg=#9a3412;btn=#f8fafc;btn_bd=#cbd2dc;" & _
            "btn_hover=#eef1f5;btn_hover_bd=#9ca3af;dis_fg=#a0a7b4;dis_bg=#f3f4f6;dis_bd=#e2e5ea;" & _
            "primary_hover=#f59e0b;primary_dis_bg=#e5e7eb;primary_dis_fg=#9ca3af;accent2_bd=#0e7490;" & _
            "accent2_fg=#0e7490;accent2_hover=#e0f7fb;groove=#d1d5db;statusbar=#ffffff;viewbar=#f8fafc;" & _
            "canvas=#e5e7eb;ok=#15803d;warn=#b45309;err=#dc2626;info=#374151;drop_text=#6b7280"
    Else
        s = "text=#e5e7eb;bg=#15171c;hdr1=#1f2937;border=#2a2f3a;title=#f9fafb;subtitle=#9ca3af;" & _
            "chip_bg=#1e3a2f;chip_fg=#6ee7b7;chip_bd=#2f6b54;warn_bg=#3b2a12;warn_fg=#fbbf24;warn_bd=#6b4e1f;" & _
            "card=#1d2027;card_title=#f3f4f6;accent=#f28c28;on_accent=#111827;hint=#8b93a3;metric=#93c5fd;" & _
            "input_bg=#12141a;input_bd=#323846;sel_bg=#3a2a1a;sel_fg=#fdba74;btn=#262b36;btn_bd=#343b4a;" & _
            "btn_hover=#2f3542;btn_hover_bd=#4b5563;dis_fg=#5b6170;dis_bg=#1c1f26;dis_bd=#262a33;" & _
            "primary_hover=#f59e0b;primary_dis_bg=#3a3f4b;primary_dis_fg=#6b7280;accent2_bd=#0e7490;" & _
            "accent2_fg=#67e8f9;accent2_hover=#10343d;groove=#323846;statusbar=#111318;viewbar=#1a1d24;" & _
            "canvas=#0f1115;ok=#6ee7b7;warn=#fbbf24;err=#f87171;info=#cbd5e1;drop_text=#9ca3af"
        theme = "dark"
    End If
    mTheme = theme
    Set mPal = CreateObject("Scripting.Dictionary")
    kv = Split(s, ";")
    For i = 0 To UBound(kv)
        p = InStr(kv(i), "=")
        If p > 0 Then mPal(Left$(kv(i), p - 1)) = HexRGB(Mid$(kv(i), p + 1))
    Next
End Sub

Public Function PalC(ByVal key As String) As Long
    If mPal Is Nothing Then LoadPalette GetSetting(REG_APP, "ui", "theme", "dark")
    If mPal.Exists(key) Then PalC = mPal(key) Else PalC = 0
End Function

'================================================================ building blocks
Private Function Has(ByVal nm As String) As Boolean
    If mCtl Is Nothing Then Exit Function
    Has = mCtl.Exists(nm)
End Function

Private Function C(ByVal nm As String) As Object
    If Has(nm) Then Set C = mCtl(nm)
End Function

Private Sub SetFont(ByVal o As Object, ByVal size As Single, Optional ByVal bold As Boolean, Optional ByVal sym As Boolean)
    On Error Resume Next
    o.Font.Name = IIf(sym, FONT_SYM, FONT_UI)
    o.Font.size = size
    o.Font.bold = bold
End Sub

Private Sub Hook(ByVal o As Object, ByVal nm As String)
    Dim e As CEvt
    Set e = New CEvt
    e.Name = nm
    Select Case TypeName(o)
        Case "Label": Set e.eLbl = o
        Case "CheckBox": Set e.eChk = o
        Case "OptionButton": Set e.eOpt = o
        Case "ComboBox": Set e.eCbo = o
        Case "TextBox": Set e.eTxt = o
        Case "ListBox": Set e.eLst = o
        Case "ScrollBar": Set e.eScr = o
        Case "Image": Set e.eImg = o
        Case "Frame": Set e.eFrm = o
        Case Else: Exit Sub
    End Select
    mEvt.Add e
End Sub

Private Function Add(ByVal par As Object, ByVal progId As String, ByVal nm As String, ByVal role As String, _
                     ByVal L As Single, ByVal T As Single, ByVal W As Single, ByVal H As Single, _
                     Optional ByVal cap As String = "", Optional ByVal tip As String = "", _
                     Optional ByVal hookIt As Boolean = True) As Object
    Dim o As Object
    Set o = par.Controls.Add(progId, nm, True)
    o.Left = L: o.Top = T
    If W > 0 Then o.Width = W
    If H > 0 Then o.Height = H
    SetFont o, 9
    Select Case TypeName(o)
        Case "Label"
            o.BackStyle = 0: o.Caption = cap: o.WordWrap = False
        Case "CheckBox", "OptionButton"
            o.BackStyle = 0: o.Caption = cap: o.SpecialEffect = 0
        Case "TextBox", "ComboBox", "ListBox"
            o.SpecialEffect = 0: o.BorderStyle = 1
        Case "Frame"
            o.Caption = "": o.SpecialEffect = 0: o.BorderStyle = 0
    End Select
    If Len(tip) Then o.ControlTipText = tip
    mCtl(nm) = Empty
    Set mCtl(nm) = o
    mRole(nm) = role
    If hookIt Then Hook o, nm
    Set Add = o
End Function

Private Function LineH(ByVal size As Single, ByVal bold As Boolean, ByVal sym As Boolean) As Single
    Dim t As Object
    Set t = mF.Controls.Add("Forms.Label.1", "zzMeasure", False)
    SetFont t, size, bold, sym
    t.WordWrap = False: t.Caption = "Ag": t.AutoSize = True
    LineH = t.Height
    mF.Controls.Remove "zzMeasure"
End Function

'flat button = background label (border/colour) + centred caption label
Private Sub AddBtn(ByVal par As Object, ByVal nm As String, ByVal cap As String, ByVal role As String, _
                   ByVal L As Single, ByVal T As Single, ByVal W As Single, ByVal H As Single, _
                   Optional ByVal tip As String = "", Optional ByVal sym As Boolean = True)
    Dim bg As Object, tx As Object, sz As Single, bold As Boolean, lh As Single
    sz = IIf(role = "primary", 10.5, 9): bold = (role = "primary")
    Set bg = Add(par, "Forms.Label.1", nm & "_bg", "btnbg", L, T, W, H, "", tip)
    bg.BackStyle = 1: bg.BorderStyle = 1
    Set tx = Add(par, "Forms.Label.1", nm, role, L + 1, T, W - 2, 0, cap, tip)
    SetFont tx, sz, bold, sym
    tx.TextAlign = 2
    lh = LineH(sz, bold, sym)
    tx.Height = lh
    tx.Top = T + (H - lh) / 2
    mBtnEn(nm) = True
    mBtnChk(nm) = False
End Sub

Private Sub MoveBtn(ByVal nm As String, ByVal L As Single, ByVal T As Single, Optional ByVal W As Single = -1)
    Dim bg As Object, tx As Object, dy As Single
    Set bg = C(nm & "_bg"): Set tx = C(nm)
    If bg Is Nothing Then Exit Sub
    dy = tx.Top - bg.Top
    bg.Left = L: bg.Top = T
    If W > 0 Then bg.Width = W: tx.Width = W - 2
    tx.Left = L + 1: tx.Top = T + dy
End Sub

Private Function IsBtn(ByVal nm As String) As Boolean
    If mBtnEn Is Nothing Then Exit Function
    IsBtn = mBtnEn.Exists(nm)
End Function

Private Sub StyleBtn(ByVal nm As String)
    Dim bg As Object, tx As Object, role As String, en As Boolean, ck As Boolean, hv As Boolean
    Dim b As Long, bd As Long, fg As Long
    Set bg = C(nm & "_bg"): Set tx = C(nm)
    If bg Is Nothing Then Exit Sub
    role = mRole(nm): en = mBtnEn(nm): ck = mBtnChk(nm): hv = (mHover = nm) And en
    Select Case role
        Case "primary"
            If Not en Then
                b = PalC("primary_dis_bg"): bd = b: fg = PalC("primary_dis_fg")
            Else
                b = IIf(hv, PalC("primary_hover"), PalC("accent")): bd = b: fg = PalC("on_accent")
            End If
        Case Else
            If Not en Then
                b = PalC("dis_bg"): bd = PalC("dis_bd"): fg = PalC("dis_fg")
            ElseIf ck Then
                b = PalC("sel_bg"): bd = PalC("accent"): fg = PalC("sel_fg")
            ElseIf role = "accent" Then
                b = IIf(hv, PalC("accent2_hover"), PalC("btn")): bd = PalC("accent2_bd"): fg = PalC("accent2_fg")
            Else
                b = IIf(hv, PalC("btn_hover"), PalC("btn")): bd = IIf(hv, PalC("btn_hover_bd"), PalC("btn_bd"))
                fg = PalC("text")
            End If
    End Select
    bg.BackColor = b: bg.BorderColor = bd: tx.ForeColor = fg
    bg.MousePointer = IIf(en, 99, 0): tx.MousePointer = IIf(en, 99, 0)
End Sub

Private Function CardFrame(ByVal n As Long, ByVal title As String, ByVal T As Single, ByVal H As Single) As Object
    Dim f As Object, b As Object
    Set f = Add(C("fraSide"), "Forms.Frame.1", "fraCard" & n, "card", CL, T, CWD, H)
    f.BorderStyle = 1
    Set b = Add(f, "Forms.Label.1", "lblBadge" & n, "badge", PAD, 9, 17, 17, CStr(n), "", False)
    SetFont b, 9, True
    b.BackStyle = 1: b.TextAlign = 2
    Add(f, "Forms.Label.1", "lblTitle" & n, "cardTitle", PAD + 25, 9, 200, 18, title, "", False).Font.bold = True
    C("lblTitle" & n).Font.size = 10.5
    With Add(f, "Forms.Label.1", "lblCard" & n, "hint", PAD + 120, 11, IW - 120, 14, "", "", False)
        .TextAlign = 3
        SetFont C("lblCard" & n), 9, False, True
    End With
    mCardOk(n) = Null
    Set CardFrame = f
End Function

Private Sub Placeholder(ByVal par As Object, ByVal txtName As String, ByVal text As String)
    Dim t As Object, p As Object
    Set t = C(txtName)
    Set p = Add(par, "Forms.Label.1", txtName & "_ph", "hint", t.Left + 4, t.Top + 3, t.Width - 8, t.Height - 5, text)
    SetFont p, 9, False, True
    p.MousePointer = 3
    mPh(txtName) = txtName & "_ph"
End Sub

Private Sub SyncPlaceholder(ByVal txtName As String)
    If mPh Is Nothing Then Exit Sub
    If Not mPh.Exists(txtName) Then Exit Sub
    C(mPh(txtName)).Visible = (Len(C(txtName).text) = 0)
End Sub

'================================================================ build
Public Sub BuildUI(ByVal frm As Object)
    Dim f As Object, o As Object, i As Long, T As Single, cw2 As Single, lv As Variant, lvl As Variant
    mBuilding = True
    Set mF = frm
    Set mCtl = CreateObject("Scripting.Dictionary")
    Set mRole = CreateObject("Scripting.Dictionary")
    Set mBtnEn = CreateObject("Scripting.Dictionary")
    Set mBtnChk = CreateObject("Scripting.Dictionary")
    Set mPh = CreateObject("Scripting.Dictionary")
    Set mEvt = New Collection
    Set mListJ = New Collection
    mHover = "": mBusy = 0: mLogOn = False: gTool = TOOL_PAN
    LoadPalette GetSetting(REG_APP, "ui", "theme", "dark")

    SetupWindow
    SetFont mF, 9

    '---- header
    Set f = Add(mF, "Forms.Frame.1", "fraHdr", "hdr", 0, 0, 800, HDR_H)
    Set o = Add(f, "Forms.Label.1", "lblLogoBox", "logobox", 16, 11, 133, 40, "", "Virginia Department of Transportation", False)
    o.BackStyle = 1
    Set o = Add(f, "Forms.Image.1", "imgLogo", "logo", 22, 14.5, 121, 33, "", "Virginia Department of Transportation", False)
    o.BorderStyle = 0: o.BackStyle = 0: o.PictureSizeMode = 1
    On Error Resume Next
    o.Picture = LoadPicture(LogoFile(gTmpDir))
    On Error GoTo 0
    Set o = Add(f, "Forms.Label.1", "lblAppTitle", "title", 162, 6, 500, 22, APP_TITLE)
    SetFont o, 14, True
    Add f, "Forms.Label.1", "lblSubtitle", "subtitle", 162, 29, 560, 14, _
        "PDF road map  " & UArrow() & "  VDOT LRS boundary  " & UArrow() & "  georeferenced, clipped raster in MicroStation"
    Add f, "Forms.Label.1", "lblCredit", "hint", 162, 43, 560, 14, AppCredit()
    AddBtn f, "btnTheme", "", "tool", 500, 19, 104, 24
    Set o = Add(f, "Forms.Label.1", "lblChip", IIf(gHasGcs, "chip", "chipWarn"), 620, 21, 100, 20, _
        IIf(gHasGcs, ChrW(&H2316) & " " & gGcsLabel, ChrW(&H26A0) & " Active model has no GCS"), _
        "Coordinates from the VDOT services are reprojected with the active model's GCS.")
    SetFont o, 9, False, True
    o.BackStyle = 1: o.BorderStyle = 1: o.TextAlign = 2: o.WordWrap = False: o.AutoSize = True
    T = o.Height: o.AutoSize = False: o.Width = o.Width + 20: o.Height = T + 6
    Set o = Add(mF, "Forms.Label.1", "lnHdr", "line", 0, HDR_H, 800, 1)
    o.BackStyle = 1

    '---- sidebar
    Set f = Add(mF, "Forms.Frame.1", "fraSide", "bg", 0, HDR_H + 1, SBW, 400)
    f.ScrollBars = 2: f.KeepScrollBarsVisible = 0
    T = 12

    ' 1 Map PDF
    Set f = CardFrame(1, "Map PDF", T, 82)
    Set o = Add(f, "Forms.TextBox.1", "txtPdf", "input", PAD, 34, IW - 72, 20)
    o.Locked = True
    Placeholder f, "txtPdf", "Select a VDOT county road map PDF" & UEll()
    AddBtn f, "btnBrowse", "Browse" & UEll(), "btn", PAD + IW - 66, 34, 66, 20
    Add f, "Forms.Label.1", "lblPage", "text", PAD, 62, 60, 14, "Page  1"
    Set o = Add(f, "Forms.Label.1", "lblPdf", "hint", PAD + 60, 62, IW - 60, 14, "")
    o.TextAlign = 3
    T = T + 82 + 12

    ' 2 Jurisdiction
    Set f = CardFrame(2, "Jurisdiction boundary", T, 214)
    Add f, "Forms.TextBox.1", "txtSearch", "input", PAD, 34, IW - 84, 20
    Placeholder f, "txtSearch", ChrW(&H2315) & "  Search counties, cities, towns" & UEll()
    Set o = Add(f, "Forms.ComboBox.1", "cboType", "input", PAD + IW - 78, 34, 78, 20)
    o.Style = 2
    o.AddItem "All": o.AddItem "Counties": o.AddItem "Cities": o.AddItem "Towns"
    mSilent = True: o.ListIndex = 1: mSilent = False
    Set o = Add(f, "Forms.ListBox.1", "lstJuris", "input", PAD, 60, IW, 114)
    o.IntegralHeight = False
    Set o = Add(f, "Forms.Label.1", "lblJuris", "hint", PAD, 182, IW - 100, 26, "Loading jurisdictions" & UEll())
    o.WordWrap = True
    AddBtn f, "btnBoundary", "Load boundary", "btn", PAD + IW - 94, 182, 94, 22
    T = T + 214 + 12

    ' 3 LRS routes
    Set f = CardFrame(3, "LRS routes", T, 160)
    cw2 = IW / 2
    For i = 0 To NGRP - 1
        Set o = Add(f, "Forms.CheckBox.1", "chkGrp" & i, "grp", PAD + (i Mod 2) * cw2, 34 + (i \ 2) * 18, cw2 - 4, 17, GrpName(i))
        o.Value = GrpOn(i)
    Next
    Add(f, "Forms.CheckBox.1", "chkClipRoutes", "chk", PAD, 108, cw2 - 4, 17, "Clip to boundary").Value = True
    Add(f, "Forms.CheckBox.1", "chkShowRoutes", "chk", PAD + cw2, 108, cw2 - 4, 17, "Show in preview").Value = True
    Add f, "Forms.Label.1", "lblRoutes", "hint", PAD, 134, IW - 100, 14, "Not loaded"
    AddBtn f, "btnRoutes", "Fetch routes", "btn", PAD + IW - 94, 130, 94, 22
    T = T + 160 + 12

    ' 4 Align
    Set f = CardFrame(4, "Align PDF to boundary", T, 152)
    Add f, "Forms.Label.1", "lblModel", "text", PAD, 36, 34, 14, "Model"
    Set o = Add(f, "Forms.OptionButton.1", "optSim", "chk", PAD + 36, 34, 74, 17, "Similarity", _
        "Shift + uniform scale + rotation (recommended)")
    o.GroupName = "model": o.Value = True
    Set o = Add(f, "Forms.OptionButton.1", "optAff", "chk", PAD + 112, 34, 60, 17, "Affine", _
        "Adds independent X/Y scale and skew (projection differences)")
    o.GroupName = "model"
    Add f, "Forms.CheckBox.1", "chkUseRoutes", "chk", PAD + IW - 80, 34, 80, 17, "Use routes", _
        "Also fit LRS routes to the road linework (after fetching routes)"
    AddBtn f, "btnAuto", ChrW(&H2728) & " Auto-align", "accent", PAD, 58, 136, 24
    AddBtn f, "btnRefine", "Refine", "btn", PAD + 142, 58, 66, 24, "Fine-tune the current alignment to the map linework"
    AddBtn f, "btnReset", "Reset", "btn", PAD + 214, 58, 66, 24
    AddBtn f, "btnPan", ChrW(&H270B) & " Pan", "tool", PAD, 88, 89, 22, "Pan / zoom the preview (drag; use + / " & ChrW(8722) & " to zoom)"
    AddBtn f, "btnAdjust", ChrW(&H2927) & " Adjust", "tool", PAD + 95, 88, 90, 22, _
        "Drag to move overlay " & UDot() & " Shift+drag rotate " & UDot() & " Ctrl+drag scale"
    AddBtn f, "btnTie", ChrW(&H2316) & " Tie points", "tool", PAD + 191, 88, 89, 22, _
        "Click the cyan boundary, then the matching spot on the PDF. 2+ pairs (3+ for affine) solve the fit. Right-click cancels."
    Set o = Add(f, "Forms.Label.1", "lblFit", "metric", PAD, 118, IW - 90, 28, "No alignment yet")
    o.WordWrap = True: o.Font.bold = True
    AddBtn f, "btnClearPins", "Clear pins", "btn", PAD + IW - 82, 118, 82, 22
    T = T + 152 + 12

    ' 5 Output
    Set f = CardFrame(5, "Place in MicroStation", T, 300)
    lv = Array("chkPlaceBnd", "Boundary shapes", True, "", _
               "chkPlaceRoutes", "LRS routes", True, "", _
               "chkPlaceRaster", "Map image", True, "Render the PDF page to a PNG (saved next to the PDF) and attach it georeferenced", _
               "chkClip", "Mask outside county", True, "Make everything outside the county boundary transparent in the image", _
               "chkWhite", "Transparent white", False, "Make white paper transparent so imagery/linework shows through", _
               "chkHoles", "Mask holes", False, "Exclude interior holes of the boundary (e.g. independent cities)", _
               "chkPdfOff", "Also attach PDF (off)", True, "Also attach the full original PDF page, georeferenced but turned off in all views", _
               "chkScaleNorth", "Scale bar & north arrow", True, "Place a scale bar and a true-north arrow below the county's bottom-right corner")
    For i = 0 To 7
        Set o = Add(f, "Forms.CheckBox.1", CStr(lv(i * 4)), "chk", PAD + (i Mod 2) * cw2, 34 + (i \ 2) * 18, cw2 - 4, 17, _
                    CStr(lv(i * 4 + 1)), CStr(lv(i * 4 + 3)))
        o.Value = lv(i * 4 + 2)
    Next
    Add f, "Forms.Label.1", "lblTranspCap", "text", PAD, 112, 104, 14, "Raster transparency"
    Set o = Add(f, "Forms.ScrollBar.1", "scrTransp", "scroll", PAD + 108, 112, IW - 108 - 42, 13)
    o.Orientation = 1: o.Min = 0: o.Max = 90: o.SmallChange = 1: o.LargeChange = 10: o.Value = 0
    Set o = Add(f, "Forms.Label.1", "lblTransp", "text", PAD + IW - 36, 112, 36, 14, "0 %")
    o.TextAlign = 3
    lvl = Array("Boundary level", "VDOT_LRS_Boundary", "Routes level", "VDOT_LRS_Routes", _
                "Raster level", "VDOT_PDF_Map", "Annotation level", "VDOT_Map_Annotation")
    For i = 0 To 3
        Add f, "Forms.Label.1", "lblLv" & i, "text", PAD, 136 + i * 23, 104, 14, CStr(lvl(i * 2))
        Add(f, "Forms.TextBox.1", "txtLv" & i, "input", PAD + 108, 133 + i * 23, IW - 108, 19).text = CStr(lvl(i * 2 + 1))
        If i = 1 Then SetTip "lblLv1", "Prefix for the per-class route levels": SetTip "txtLv1", "Routes go on one level per LRS class: <prefix>_Interstate, _US_Primary, _State_Primary, _Secondary, _Ramps_Frontage, _Urban_Streets, _Other."
    Next
    AddBtn f, "btnPlace", "Place in MicroStation  " & ChrW(&H279C), "primary", PAD, 230, IW, 30
    Set o = Add(f, "Forms.Label.1", "lblOut", "hint", PAD, 266, IW, 30, "")
    o.WordWrap = True
    T = T + 300 + 12
    C("fraSide").ScrollHeight = T

    '---- viewbar
    Set f = Add(mF, "Forms.Frame.1", "fraVbar", "viewbar", SBW + 1, HDR_H + 1, 600, VBAR_H)
    AddBtn f, "btnFit", ChrW(&H2922) & " Fit", "tool", 10, 5, 48, 20
    AddBtn f, "btnZoomIn", "+", "tool", 62, 5, 26, 20
    AddBtn f, "btnZoomOut", ChrW(8722), "tool", 92, 5, 26, 20
    Add f, "Forms.CheckBox.1", "chkOverlay", "chk", 132, 7, 60, 17, "Overlay"
    C("chkOverlay").Value = True
    Add f, "Forms.CheckBox.1", "chkRegion", "chk", 194, 7, 112, 17, "Detected map area"
    AddBtn f, "btnFence", ChrW(&H25AD) & " Fence", "tool", 314, 5, 60, 20, _
        "Drag a block fence on the PDF. When a fence is drawn, the map image, LRS routes" & vbLf & _
        "and scale bar are clipped to it instead of the county boundary."
    AddBtn f, "btnFenceClear", ChrW(&H2715) & " Clear fence", "tool", 378, 5, 82, 20, _
        "Remove the fence and clip to the county boundary again"
    Add f, "Forms.Label.1", "lblPdfOp", "text", 474, 8, 22, 14, "PDF"
    Set o = Add(f, "Forms.ScrollBar.1", "scrPdf", "scroll", 498, 8, 90, 13)
    o.Orientation = 1: o.Min = 10: o.Max = 100: o.SmallChange = 5: o.LargeChange = 10
    mSilent = True: o.Value = 100: mSilent = False
    Add f, "Forms.Label.1", "lblHint", "hint", 600, 8, 300, 14, ""
    Set o = Add(mF, "Forms.Label.1", "lnVbar", "line", SBW + 1, HDR_H + 1 + VBAR_H, 600, 1)
    o.BackStyle = 1
    Set o = Add(mF, "Forms.Label.1", "lnSide", "line", SBW, HDR_H + 1, 1, 400)
    o.BackStyle = 1

    '---- view
    Set o = Add(mF, "Forms.Image.1", "imgView", "canvas", SBW + 1, HDR_H + VBAR_H + 2, 600, 400)
    o.BorderStyle = 0: o.PictureSizeMode = 1: o.PictureAlignment = 0
    Set o = Add(mF, "Forms.Label.1", "lblDrop", "drop", SBW + 40, 300, 400, 20, _
        "Open a VDOT map PDF with Browse" & UEll() & " in step 1")
    o.TextAlign = 2: SetFont o, 11
    o.MousePointer = 99

    '---- status bar
    Set f = Add(mF, "Forms.Frame.1", "fraStatus", "statusbar", 0, 600, 800, STAT_H)
    Set o = Add(mF, "Forms.Label.1", "lnStatus", "line", 0, 600, 800, 1)
    o.BackStyle = 1
    Add f, "Forms.Label.1", "lblStatus", "status", 14, 9, 400, 14, "Ready"
    Set o = Add(f, "Forms.Label.1", "lblProgTrack", "progtrack", 420, 11, 195, 9, "", "", False)
    o.BackStyle = 1: o.BorderStyle = 1
    Set o = Add(f, "Forms.Label.1", "lblProgFill", "progfill", 421, 12, 0, 7, "", "", False)
    o.BackStyle = 1: o.Visible = False
    AddBtn f, "btnLog", "Log  " & ChrW(&H25BE), "tool", 640, 5, 60, 20
    Set o = Add(f, "Forms.TextBox.1", "txtLog", "log", 14, 30, 700, LOG_H - 6)
    o.MultiLine = True: o.WordWrap = True: o.ScrollBars = 2: o.Locked = True
    SetFont o, 8.5
    o.Visible = False

    ApplyTheme mTheme
    FenceChanged
    SetToolButtons
    mBuilding = False
    SizeWindow
    Layout
    UpdateState
End Sub

Private Sub SetupWindow()
    Dim tmp As String, st As LongPtr, ttl As String
    tmp = "VDOTGEOREF_" & Hex$(GetTickCount)
    mF.Caption = tmp
    mHwnd = FindWindowW(StrPtr("ThunderDFrame"), StrPtr(tmp))
    If mHwnd = 0 Then mHwnd = FindWindowW(0, StrPtr(tmp))
    ttl = APP_TITLE & " " & APP_VERSION & " " & ChrW(8212) & " " & AppCredit()
    mF.Caption = APP_TITLE & " " & APP_VERSION
    If mHwnd <> 0 Then
        st = GetWindowLongPtrW(mHwnd, -16)
        st = st Or &H40000 Or &H10000 Or &H20000     ' WS_THICKFRAME | WS_MAXIMIZEBOX | WS_MINIMIZEBOX
        SetWindowLongPtrW mHwnd, -16, st
        DrawMenuBar mHwnd
        SetWindowTextW mHwnd, StrPtr(ttl)
    End If
End Sub

Private Sub SizeWindow()
    Dim hdc As LongPtr, dpi As Long, sw As Single, sh As Single, W As Single, H As Single
    hdc = GetDC(0): dpi = GetDeviceCaps(hdc, 88): ReleaseDC 0, hdc
    If dpi <= 0 Then dpi = 96
    sw = GetSystemMetrics(16) * 72# / dpi: sh = GetSystemMetrics(17) * 72# / dpi   ' full-screen client area
    W = 1110: H = 760
    If W > sw - 20 Then W = sw - 20
    If H > sh - 10 Then H = sh - 10
    If W < 760 Then W = 760
    If H < 520 Then H = 520
    mF.StartUpPosition = 0
    mF.Width = W: mF.Height = H
    mF.Left = (sw - W) / 2: mF.Top = (sh - H) / 2
    If mF.Top < 0 Then mF.Top = 0
End Sub

Public Sub Layout()
    Dim W As Single, H As Single, sy As Single, vy As Single, vh As Single, x As Single, o As Object
    If mF Is Nothing Or mBuilding Then Exit Sub
    On Error Resume Next
    W = mF.InsideWidth: H = mF.InsideHeight
    If W < 600 Then W = 600
    If H < 400 Then H = 400
    sy = H - STAT_H - IIf(mLogOn, LOG_H, 0)
    ' header
    C("fraHdr").Width = W: C("lnHdr").Width = W
    Set o = C("lblChip")
    o.Left = W - 18 - o.Width
    MoveBtn "btnTheme", o.Left - 10 - 104, 19
    C("lblSubtitle").Width = o.Left - 10 - 104 - 172
    C("lblCredit").Width = C("lblSubtitle").Width
    ' sidebar
    C("fraSide").Height = sy - HDR_H - 1
    C("lnSide").Height = sy - HDR_H - 1
    ' viewbar
    C("fraVbar").Width = W - SBW - 1
    C("lnVbar").Width = W - SBW - 1
    C("lblHint").Width = W - SBW - 1 - 600 - 8
    ' view
    vy = HDR_H + VBAR_H + 2: vh = sy - vy
    If vh < 40 Then vh = 40
    Set o = C("imgView")
    o.Left = SBW + 1: o.Top = vy: o.Width = W - SBW - 1: o.Height = vh
    Set o = C("lblDrop")
    o.Left = SBW + 21: o.Width = W - SBW - 42: o.Top = vy + vh / 2 - 10
    ' status bar
    Set o = C("fraStatus")
    o.Top = sy: o.Width = W: o.Height = H - sy
    C("lnStatus").Top = sy: C("lnStatus").Width = W
    x = W - 14 - 60
    MoveBtn "btnLog", x, 5
    C("lblProgTrack").Left = x - 12 - 195
    C("lblProgFill").Left = x - 12 - 194
    C("lblStatus").Width = x - 12 - 195 - 14 - 10
    C("txtLog").Width = W - 28
    C("txtLog").Height = LOG_H - 6
    ViewRedraw True
End Sub

'================================================================ theme
Public Sub ApplyTheme(ByVal theme As String)
    Dim k As Variant, o As Object, role As String, n As Long
    LoadPalette theme
    SaveSetting REG_APP, "ui", "theme", mTheme
    mF.BackColor = PalC("bg")
    For Each k In mCtl.Keys
        Set o = mCtl(k): role = mRole(k)
        On Error Resume Next
        Select Case role
            Case "hdr": o.BackColor = PalC("hdr1")
            Case "bg": o.BackColor = PalC("bg")
            Case "card": o.BackColor = PalC("card"): o.BorderColor = PalC("border")
            Case "viewbar": o.BackColor = PalC("viewbar")
            Case "statusbar": o.BackColor = PalC("statusbar")
            Case "line": o.BackColor = PalC("border")
            Case "logobox": o.BackColor = RGB(255, 255, 255)
            Case "title": o.ForeColor = PalC("title")
            Case "subtitle": o.ForeColor = PalC("subtitle")
            Case "hint": o.ForeColor = PalC("hint")
            Case "metric": o.ForeColor = PalC("metric")
            Case "text", "chk": o.ForeColor = PalC("text")
            Case "grp": o.ForeColor = GrpPrev(CLng(Mid$(k, 7)))
            Case "cardTitle": o.ForeColor = PalC("card_title")
            Case "badge": o.BackColor = PalC("accent"): o.ForeColor = PalC("on_accent")
            Case "chip": o.BackColor = PalC("chip_bg"): o.ForeColor = PalC("chip_fg"): o.BorderColor = PalC("chip_bd")
            Case "chipWarn": o.BackColor = PalC("warn_bg"): o.ForeColor = PalC("warn_fg"): o.BorderColor = PalC("warn_bd")
            Case "input", "log"
                o.BackColor = PalC("input_bg"): o.ForeColor = PalC("text"): o.BorderColor = PalC("input_bd")
            Case "scroll": o.BackColor = PalC("groove"): o.ForeColor = PalC("accent")
            Case "canvas": o.BackColor = PalC("canvas")
            Case "drop": o.ForeColor = PalC("drop_text")
            Case "status": o.ForeColor = PalC("info")
            Case "progtrack": o.BackColor = PalC("input_bg"): o.BorderColor = PalC("border")
            Case "progfill": o.BackColor = PalC("accent")
        End Select
        On Error GoTo 0
    Next
    For Each k In mBtnEn.Keys
        StyleBtn CStr(k)
    Next
    For n = 1 To 5
        If Not IsNull(mCardOk(n)) And Not IsEmpty(mCardOk(n)) Then
            C("lblCard" & n).ForeColor = IIf(mCardOk(n), PalC("ok"), PalC("warn"))
        End If
    Next
    If mTheme = "dark" Then
        SetCap "btnTheme", ChrW(&H2600) & "  Light mode": SetTip "btnTheme", "Switch to light mode"
    Else
        SetCap "btnTheme", ChrW(&H263E) & "  Dark mode": SetTip "btnTheme", "Switch to dark mode"
    End If
    If Not mBuilding Then ViewRedraw True
End Sub

'================================================================ events
Public Sub OnEvent(ByVal nm As String, ByVal kind As String)
    If mSilent Or mBuilding Or mF Is Nothing Then Exit Sub
    If Right$(nm, 3) = "_bg" Then nm = Left$(nm, Len(nm) - 3)
    If Right$(nm, 3) = "_ph" Then
        On Error Resume Next
        C(Left$(nm, Len(nm) - 3)).SetFocus
        If nm = "txtPdf_ph" And mBusy = 0 Then DoBrowsePdf
        Exit Sub
    End If
    If IsBtn(nm) Then
        If Not mBtnEn(nm) Then Exit Sub
    End If
    ' controls that only update their own display stay live while busy
    Select Case nm
        Case "scrTransp": SetCap "lblTransp", C("scrTransp").Value & " %": Exit Sub
        Case "txtSearch": SyncPlaceholder nm
        Case "btnLog": ToggleLog: Exit Sub
    End Select
    If mBusy > 0 Then Exit Sub
    Select Case nm
        Case "btnBrowse", "lblDrop": DoBrowsePdf
        Case "txtSearch", "cboType": FilterJuris
        Case "lstJuris"
            If kind = "dblclick" Then DoLoadBoundary Else UpdateState
        Case "btnBoundary": DoLoadBoundary
        Case "btnRoutes": DoFetchRoutes
        Case "chkShowRoutes", "chkOverlay": ViewRedraw False
        Case "btnAuto": DoAutoAlign True
        Case "btnRefine": DoAutoAlign False
        Case "btnReset": DoResetAlignment
        Case "btnPan": SetTool TOOL_PAN
        Case "btnAdjust": SetTool TOOL_ADJUST
        Case "btnTie": SetTool TOOL_TIE
        Case "btnFence": SetTool TOOL_FENCE
        Case "btnClearPins": DoClearPins
        Case "chkPlaceRaster": UpdateState
        Case "btnPlace": DoPlace
        Case "btnFit": ViewFit
        Case "btnZoomIn": ViewZoom 1.25
        Case "btnZoomOut": ViewZoom 0.8
        Case "chkRegion": ToggleRegion
        Case "btnFenceClear": ClearFence
        Case "scrPdf": ViewInvalidate: ViewRedraw True
        Case "btnTheme": ApplyTheme IIf(mTheme = "dark", "light", "dark")
    End Select
End Sub

Public Sub OnHover(ByVal nm As String)
    Dim old As String
    If mF Is Nothing Then Exit Sub
    If Right$(nm, 3) = "_bg" Then nm = Left$(nm, Len(nm) - 3)
    If Not IsBtn(nm) Then nm = ""
    If nm = mHover Then Exit Sub
    old = mHover: mHover = nm
    If Len(old) Then StyleBtn old
    If Len(nm) Then StyleBtn nm
End Sub

Private Sub ToggleRegion()
    If Chk("chkRegion") And gHasPdf Then
        BusyOn "Detecting the map area" & UEll()
        EnsureRegion
        BusyOff
    End If
    ViewInvalidate
    ViewRedraw True
End Sub

Private Sub ToggleLog()
    mLogOn = Not mLogOn
    C("txtLog").Visible = mLogOn
    SetCap "btnLog", "Log  " & IIf(mLogOn, ChrW(&H25B4), ChrW(&H25BE))
    mBtnChk("btnLog") = mLogOn: StyleBtn "btnLog"
    Layout
End Sub

Public Sub SetTool(ByVal i As Long)
    gTool = i
    CancelPending
    SetToolButtons
    ViewRedraw False
End Sub

Private Sub SetToolButtons()
    Dim h As Variant
    h = Array("Drag = pan " & UDot() & " + / " & ChrW(8722) & " = zoom", _
              "Drag = move " & UDot() & " Shift+drag = rotate " & UDot() & " Ctrl+drag = scale", _
              "Click cyan boundary " & UArrow() & " click matching PDF spot " & UDot() & " right-click = cancel", _
              "Drag a rectangle on the PDF to set the block fence")
    SetCap "lblHint", CStr(h(gTool))
    mBtnChk("btnPan") = (gTool = TOOL_PAN): StyleBtn "btnPan"
    mBtnChk("btnAdjust") = (gTool = TOOL_ADJUST): StyleBtn "btnAdjust"
    mBtnChk("btnTie") = (gTool = TOOL_TIE): StyleBtn "btnTie"
    mBtnChk("btnFence") = (gTool = TOOL_FENCE): StyleBtn "btnFence"
End Sub

Public Sub OnClose()
    On Error Resume Next
    PdfClose
    Set mEvt = Nothing
    Set mCtl = Nothing
    Set mBtnEn = Nothing
    Set mF = Nothing
End Sub

'================================================================ control access
Public Function FormIsOpen() As Boolean
    Dim f As Object
    If mF Is Nothing Or mCtl Is Nothing Then Exit Function
    For Each f In VBA.UserForms
        If TypeName(f) = "frmGeoref" Then FormIsOpen = True: Exit Function
    Next
End Function

Public Function ViewCtl() As Object
    If Not FormIsOpen() Then Exit Function
    Set ViewCtl = C("imgView")
End Function

Public Sub ShowDropHint(ByVal show As Boolean)
    If Has("lblDrop") Then C("lblDrop").Visible = show
End Sub

Public Sub SetCap(ByVal nm As String, ByVal text As String)
    Dim o As Object
    Set o = C(nm)
    If o Is Nothing Then Exit Sub
    mSilent = True
    Select Case TypeName(o)
        Case "TextBox", "ComboBox": o.text = text: SyncPlaceholder nm
        Case Else: o.Caption = text
    End Select
    mSilent = False
End Sub

Public Function GetCap(ByVal nm As String) As String
    Dim o As Object
    Set o = C(nm)
    If o Is Nothing Then Exit Function
    Select Case TypeName(o)
        Case "TextBox", "ComboBox": GetCap = o.text
        Case Else: GetCap = o.Caption
    End Select
End Function

Public Sub SetTip(ByVal nm As String, ByVal tip As String)
    If Not Has(nm) Then Exit Sub
    C(nm).ControlTipText = tip
    If Has(nm & "_bg") Then C(nm & "_bg").ControlTipText = tip
End Sub

Public Sub SetEnabled(ByVal nm As String, ByVal v As Boolean)
    If Not Has(nm) Then Exit Sub
    If IsBtn(nm) Then
        If mBtnEn(nm) <> v Then mBtnEn(nm) = v: StyleBtn nm
    Else
        If C(nm).Enabled <> v Then C(nm).Enabled = v
    End If
End Sub

Public Function Chk(ByVal nm As String) As Boolean
    Dim v As Variant
    If Not Has(nm) Then Exit Function
    v = C(nm).Value
    If Not IsNull(v) Then Chk = CBool(v)
End Function

Public Sub SetChk(ByVal nm As String, ByVal v As Boolean)
    If Not Has(nm) Then Exit Sub
    mSilent = True
    C(nm).Value = v
    mSilent = False
End Sub

Public Function Opt(ByVal nm As String) As Boolean
    Opt = Chk(nm)
End Function

Public Function GetVal(ByVal nm As String) As Variant
    If Not Has(nm) Then GetVal = 0: Exit Function
    GetVal = C(nm).Value
End Function

Public Sub CardDone(ByVal n As Long, ByVal text As String, ByVal ok As Boolean)
    If Not Has("lblCard" & n) Then Exit Sub
    mCardOk(n) = ok
    C("lblCard" & n).Caption = IIf(ok, ChrW(&H2714) & " ", ChrW(8226) & " ") & text
    C("lblCard" & n).ForeColor = IIf(ok, PalC("ok"), PalC("warn"))
End Sub

Public Sub CardClear(ByVal n As Long)
    If Not Has("lblCard" & n) Then Exit Sub
    mCardOk(n) = Null
    C("lblCard" & n).Caption = ""
End Sub

'================================================================ busy / status / log
Private Sub Pump()
    On Error Resume Next
    If Not mF Is Nothing Then mF.Repaint
    DoEvents
End Sub

Public Sub BusyOn(ByVal msg As String, Optional ByVal steps As Long = 0)
    mBusy = mBusy + 1
    If mBusy = 1 Then mBusyErr = False
    mSteps = steps
    LogMsg msg
    If FormIsOpen() Then
        C("fraSide").Enabled = False
        C("imgView").Enabled = False
        mF.MousePointer = 11
        ShowProgress IIf(steps > 0, 0, -1)
    End If
    Pump
End Sub

Public Sub BusyOff()
    If mBusy > 0 Then mBusy = mBusy - 1
    If mBusy = 0 And FormIsOpen() Then
        C("fraSide").Enabled = True
        C("imgView").Enabled = True
        mF.MousePointer = 0
        ShowProgress IIf(mBusyErr, 0, 1)
    End If
    If FormIsOpen() Then UpdateState
End Sub

Private Sub ShowProgress(ByVal frac As Double)
    Dim f As Object, tw As Single
    If Not Has("lblProgFill") Then Exit Sub
    Set f = C("lblProgFill")
    tw = C("lblProgTrack").Width - 2
    If frac < 0 Then
        f.Visible = True: f.Width = tw * 0.3      ' indeterminate
    ElseIf frac <= 0 Then
        f.Visible = False
    Else
        If frac > 1 Then frac = 1
        f.Width = tw * frac: f.Visible = True
    End If
End Sub

Public Sub SetProgress(ByVal n As Long)
    If mSteps > 0 Then ShowProgress n / mSteps
    Pump
End Sub

Public Sub SetStatus(ByVal msg As String)
    If Not Has("lblStatus") Then Exit Sub
    C("lblStatus").Caption = msg
    C("lblStatus").ForeColor = PalC("info")
    Pump
End Sub

Public Sub LogMsg(ByVal msg As String, Optional ByVal kind As String = "info")
    Dim line As String, t As Object, f As Integer
    line = Format$(Now, "hh:nn:ss") & "  " & msg
    Debug.Print line
    On Error Resume Next
    f = FreeFile
    Open Environ$("TEMP") & "\vdot_georef_native.log" For Append As #f
    Print #f, line
    Close #f
    On Error GoTo 0
    If kind = "err" Then mBusyErr = True
    If Not Has("lblStatus") Then Exit Sub
    Set t = C("txtLog")
    If Len(t.text) > 60000 Then t.text = Mid$(t.text, 20000)
    t.text = t.text & IIf(Len(t.text), vbCrLf, "") & line
    On Error Resume Next
    t.SelStart = Len(t.text)
    On Error GoTo 0
    C("lblStatus").Caption = msg
    Select Case kind
        Case "ok", "warn", "err": C("lblStatus").ForeColor = PalC(kind)
        Case Else: C("lblStatus").ForeColor = PalC("info")
    End Select
    If kind = "err" And Not mLogOn Then ToggleLog
    Pump
End Sub

Public Sub Toast(ByVal msg As String)
    If Not Has("lblStatus") Then Exit Sub
    C("lblStatus").Caption = msg
    C("lblStatus").ForeColor = PalC("warn")
End Sub

Public Sub UpdateState()
    Dim hasPdf As Boolean, hasBnd As Boolean, hasG As Boolean
    If Not FormIsOpen() Then Exit Sub
    hasPdf = gHasPdf: hasBnd = gNR > 0: hasG = gHasG
    SetEnabled "btnBoundary", C("lstJuris").ListIndex >= 0
    SetEnabled "btnRoutes", hasBnd
    SetEnabled "btnAuto", hasPdf And hasBnd
    SetEnabled "btnRefine", hasPdf And hasBnd And hasG
    SetEnabled "btnReset", hasG
    SetEnabled "btnAdjust", hasG
    SetEnabled "btnTie", hasG
    SetEnabled "btnFence", hasPdf
    SetEnabled "btnClearPins", gNPairs > 0
    SetEnabled "btnPlace", hasBnd And gHasGcs And (hasG Or Not Chk("chkPlaceRaster"))
    SetEnabled "chkUseRoutes", gNRt > 0
    If Not gHasGcs Then
        SetTip "btnPlace", "The active model has no geographic coordinate system " & UDash() & " assign one first."
    Else
        SetTip "btnPlace", ""
    End If
End Sub

'================================================================ jurisdictions
Private Function TypeLabel(ByVal t As String) As String
    Select Case t
        Case "CO": TypeLabel = "County"
        Case "CI": TypeLabel = "City"
        Case "TO": TypeLabel = "Town"
    End Select
End Function

Public Sub FilterJuris()
    Dim txt As String, typ As String, cur As Variant, j As Variant, nm As String, code As String
    Dim lst As Object, sel As Long
    If Not FormIsOpen() Or gJuris Is Nothing Then Exit Sub
    txt = LCase$(Trim$(C("txtSearch").text))
    Select Case C("cboType").ListIndex
        Case 1: typ = "CO"
        Case 2: typ = "CI"
        Case 3: typ = "TO"
    End Select
    Set lst = C("lstJuris")
    cur = Empty
    If lst.ListIndex >= 0 Then cur = DGet(mListJ(lst.ListIndex + 1), "OBJECTID", Empty)
    mSilent = True
    lst.Clear
    Set mListJ = New Collection
    sel = -1
    For Each j In gJuris
        If Len(typ) = 0 Or CStr(DGet(j, "JURISTYPE", "")) = typ Then
            nm = CStr(DGet(j, "NAMELSAD", ""))
            If Len(nm) = 0 Then nm = CStr(DGet(j, "NAME", "?"))
            code = CStr(DGet(j, "VDOT_CODE", ""))
            If Len(txt) = 0 Or InStr(LCase$(nm), txt) > 0 Or InStr(code, txt) > 0 Then
                lst.AddItem nm & "   " & UDot() & "  " & TypeLabel(CStr(DGet(j, "JURISTYPE", ""))) & _
                            IIf(Len(code), " #" & code, "")
                mListJ.Add j
                If Not IsEmpty(cur) Then
                    If DGet(j, "OBJECTID", Empty) = cur Then sel = lst.ListCount - 1
                End If
            End If
        End If
    Next
    If sel >= 0 Then lst.ListIndex = sel
    mSilent = False
    UpdateState
End Sub

Public Sub SelectJuris(ByVal j As Object)
    Dim typ As String, i As Long, lst As Object
    If Not FormIsOpen() Then Exit Sub
    Select Case C("cboType").ListIndex
        Case 1: typ = "CO"
        Case 2: typ = "CI"
        Case 3: typ = "TO"
    End Select
    mSilent = True
    If Len(typ) Then
        If CStr(DGet(j, "JURISTYPE", "")) <> typ Then C("cboType").ListIndex = 0
    End If
    C("txtSearch").text = ""
    SyncPlaceholder "txtSearch"
    mSilent = False
    FilterJuris
    Set lst = C("lstJuris")
    For i = 1 To mListJ.Count
        If DGet(mListJ(i), "OBJECTID", -1) = DGet(j, "OBJECTID", -2) Then
            mSilent = True
            lst.ListIndex = i - 1
            lst.TopIndex = IIf(i > 3, i - 3, 0)
            mSilent = False
            Exit For
        End If
    Next
    UpdateState
End Sub

Public Function SelectedJuris() As Object
    Dim lst As Object
    If Not FormIsOpen() Then Exit Function
    Set lst = C("lstJuris")
    If lst.ListIndex >= 0 And lst.ListIndex < mListJ.Count Then Set SelectedJuris = mListJ(lst.ListIndex + 1)
End Function

'================================================================ file dialog
Public Function BrowseForPdf(ByVal start As String) As String
    Dim o As OFNW, buf As String, flt As String, ttl As String, p As Long
    buf = String$(1024, vbNullChar)
    flt = "PDF maps (*.pdf)" & vbNullChar & "*.pdf" & vbNullChar & "All files (*.*)" & vbNullChar & "*.*" & vbNullChar & vbNullChar
    ttl = "Open a VDOT map PDF"
    o.lStructSize = LenB(o)
    o.hwndOwner = mHwnd
    o.lpstrFilter = StrPtr(flt)
    o.nFilterIndex = 1
    o.lpstrFile = StrPtr(buf)
    o.nMaxFile = 1024
    If Len(start) Then o.lpstrInitialDir = StrPtr(start)
    o.lpstrTitle = StrPtr(ttl)
    o.Flags = &H1000 Or &H800 Or &H4 Or &H80000    ' FILEMUSTEXIST | PATHMUSTEXIST | HIDEREADONLY | EXPLORER
    If GetOpenFileNameW(o) <> 0 Then
        p = InStr(buf, vbNullChar)
        If p > 1 Then BrowseForPdf = Left$(buf, p - 1)
    End If
End Function
