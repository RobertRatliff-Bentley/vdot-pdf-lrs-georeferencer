Attribute VB_Name = "modRest"
'VDOT ArcGIS REST client (jurisdictions, boundary, LRS routes).
Option Explicit

Public Const JURIS_URL As String = "https://services.arcgis.com/p5v98VHDX9Atv3l7/arcgis/rest/services/VA_Jurisdictions/FeatureServer/0"
Public Const DISTRICT_URL As String = "https://services.arcgis.com/p5v98VHDX9Atv3l7/arcgis/rest/services/VDOTAdministrativeBoundaries/FeatureServer/2"
Public Const LRS_URL As String = "https://services.arcgis.com/p5v98VHDX9Atv3l7/arcgis/rest/services/LRS_Route_Master/FeatureServer/0"
Public Const ROUTE_ITEM_FIELDS As String = "OBJECTID,RTE_NM,RTE_COMMON_NM,RTE_CATEGORY_NM,RTE_NBR,RTE_TYPE_NM,RTE_DIRECTION_CD,RTE_JURIS_PROPER_NM,ROUTESYSIDSHORT"
Public gRouteAttrs As Object    ' RTE_NM -> LRS attribute dictionary of the last route download

Public Function UrlEnc(ByVal s As String) As String
    Dim by() As Byte, i As Long, o As String, c As Long
    If Len(s) = 0 Then Exit Function
    by = Utf8Bytes(s)
    For i = 0 To UBound(by)
        c = by(i)
        Select Case c
        Case 48 To 57, 65 To 90, 97 To 122, 45, 46, 95, 126
            o = o & Chr$(c)
        Case 32
            o = o & "+"
        Case Else
            o = o & "%" & Right$("0" & Hex$(c), 2)
        End Select
    Next
    UrlEnc = o
End Function

Private Function Num8(ByVal v As Double, ByVal decimals As Long) As String
    Num8 = Replace(Format$(v, "0." & String$(decimals, "0")), ",", ".")
End Function

'POST layer/query with retries; returns the parsed JSON dictionary.
Public Function ArcQuery(ByVal layerUrl As String, params As Object) As Object
    Dim body As String, k As Variant, attempt As Long, lastErr As String, js As Object
    For Each k In params.Keys
        body = body & IIf(Len(body), "&", "") & UrlEnc(CStr(k)) & "=" & UrlEnc(CStr(params(k)))
    Next
    body = body & "&f=json"
    For attempt = 0 To 2
        lastErr = ""
        Set js = HttpPostJson(layerUrl & "/query", body, lastErr)
        If Not js Is Nothing Then
            If js.Exists("error") Then
                Dim em As String
                If IsObject(js("error")) Then em = CStr(DGet(js("error"), "message", "unknown error")) Else em = CStr(js("error"))
                lastErr = "Service error: " & em
                Set js = Nothing
            Else
                Set ArcQuery = js
                Exit Function
            End If
        End If
        WaitSeconds 1 + attempt
    Next
    Err.Raise vbObjectError + 600, "REST", "Request failed: " & lastErr
End Function

Private Function HttpPostJson(ByVal url As String, ByVal body As String, errMsg As String) As Object
    Dim progs As Variant, i As Long, x As Object, t0 As Single, by() As Byte, v As Variant
    progs = Array("MSXML2.XMLHTTP.6.0", "MSXML2.ServerXMLHTTP.6.0")
    For i = 0 To 1
        On Error Resume Next
        Set x = Nothing
        Set x = CreateObject(progs(i))
        If Not x Is Nothing Then
            x.Open "POST", url, True
            x.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
            x.send body
            If Err.Number = 0 Then
                t0 = Timer
                Do While x.readyState <> 4
                    DoEvents
                    Sleep 15
                    If Abs(Timer - t0) > 120 Then Exit Do
                Loop
                If x.readyState = 4 Then
                    If x.Status = 200 Then
                        by = x.responseBody
                        If Err.Number = 0 Then
                            On Error GoTo 0
                            v = Empty
                            Dim o As Object
                            Set o = JsonParseBytes(by)
                            Set HttpPostJson = o
                            Exit Function
                        End If
                    Else
                        errMsg = "HTTP " & x.Status & " " & x.statusText
                    End If
                Else
                    errMsg = "timeout"
                    x.abort
                End If
            End If
            If Err.Number <> 0 Then errMsg = Err.Description
        End If
        Err.Clear
        On Error GoTo 0
    Next
End Function

Public Sub WaitSeconds(ByVal s As Double)
    Dim t0 As Single
    t0 = Timer
    Do While Abs(Timer - t0) < s
        DoEvents
        Sleep 20
    Loop
End Sub

Public Function ListJurisdictions() As Collection
    Dim out As New Collection, offset As Long, js As Object, feats As Variant, i As Long, prm As Object
    Do
        Set prm = CreateObject("Scripting.Dictionary")
        prm("where") = "1=1"
        prm("outFields") = "OBJECTID,NAME,NAMELSAD,JURISTYPE,VDOT_CODE,STCOFIPS,DISTRICT"
        prm("returnGeometry") = "false"
        prm("orderByFields") = "NAMELSAD"
        prm("resultOffset") = CStr(offset)
        prm("resultRecordCount") = "1000"
        Set js = ArcQuery(JURIS_URL, prm)
        feats = DGet(js, "features", Array())
        For i = 0 To UBound(feats)
            out.Add feats(i)("attributes")
        Next
        If UBound(feats) < 0 Then Exit Do
        If Not CBool(DGet(js, "exceededTransferLimit", False)) Then Exit Do
        offset = offset + UBound(feats) + 1
    Loop
    Set ListJurisdictions = out
End Function

'VDOT district boundaries as jurisdiction-like records (negative OBJECTIDs).
Public Function ListDistricts() As Collection
    Dim out As New Collection, js As Object, feats As Variant, i As Long, prm As Object, a As Object, d As Object, nm As String
    Set prm = CreateObject("Scripting.Dictionary")
    prm("where") = "1=1"
    prm("outFields") = "OBJECTID,DISTRICT_NAME,DISTRICT_CODE"
    prm("returnGeometry") = "false"
    prm("orderByFields") = "DISTRICT_CODE"
    Set js = ArcQuery(DISTRICT_URL, prm)
    feats = DGet(js, "features", Array())
    For i = 0 To UBound(feats)
        Set a = feats(i)("attributes")
        nm = Trim$(Replace(CStr(DGet(a, "DISTRICT_NAME", "")), "  ", " "))
        If Len(nm) Then
            Set d = CreateObject("Scripting.Dictionary")
            d("OBJECTID") = -CLng(DGet(a, "OBJECTID", 0))
            d("NAME") = nm
            d("NAMELSAD") = nm & " District"
            d("JURISTYPE") = "DI"
            d("VDOT_CODE") = ""
            d("STCOFIPS") = ""
            d("DISTRICT") = nm
            d("DISTRICT_CODE") = DGet(a, "DISTRICT_CODE", "")
            out.Add d
        End If
    Next
    Set ListDistricts = out
End Function

'Rings (lon/lat) of the jurisdiction polygon (negative id = VDOT district). Returns ring count.
Public Function GetBoundary(ByVal objectId As Long, ByVal outSR As Long, rings() As TPoly) As Long
    Dim prm As Object, js As Object, feats As Variant, rr As Variant, i As Long, j As Long, n As Long
    Set prm = CreateObject("Scripting.Dictionary")
    prm("where") = "OBJECTID=" & Abs(objectId)
    prm("outFields") = "OBJECTID"
    prm("returnGeometry") = "true"
    prm("outSR") = CStr(outSR)
    prm("geometryPrecision") = "8"
    Set js = ArcQuery(IIf(objectId < 0, DISTRICT_URL, JURIS_URL), prm)
    feats = DGet(js, "features", Array())
    If UBound(feats) < 0 Then Err.Raise vbObjectError + 601, "REST", "Boundary not found (OBJECTID=" & objectId & ")"
    rr = feats(0)("geometry")("rings")
    ReDim rings(0 To UBound(rr) + 1)
    For i = 0 To UBound(rr)
        If UBound(rr(i)) + 1 >= 4 Then
            PolyInit rings(n), UBound(rr(i)) + 1
            For j = 0 To UBound(rr(i))
                PolyAdd rings(n), rr(i)(j)(0), rr(i)(j)(1)
            Next
            n = n + 1
        End If
    Next
    GetBoundary = n
End Function

Public Function GetRoutes(ByVal bx0 As Double, ByVal by0 As Double, ByVal bx1 As Double, ByVal by1 As Double, _
                          cats As Collection, ByVal outSR As Long, routes() As TPoly, fetched As Long) As Long
    Dim where As String, c As Variant, prm As Object, js As Object, feats As Variant, f As Long, offset As Long
    Dim g As Object, a As Object, paths As Variant, pi As Long, j As Long, n As Long, pth As Variant, p As TPoly
    For Each c In cats
        where = where & IIf(Len(where), ",", "") & "'" & Replace(CStr(c), "'", "''") & "'"
    Next
    where = "RTE_CATEGORY_NM IN (" & where & ")"
    Set gRouteAttrs = CreateObject("Scripting.Dictionary")
    ReDim routes(0 To 255)
    Do
        Set prm = CreateObject("Scripting.Dictionary")
        prm("where") = where
        prm("geometry") = Num8(bx0, 8) & "," & Num8(by0, 8) & "," & Num8(bx1, 8) & "," & Num8(by1, 8)
        prm("geometryType") = "esriGeometryEnvelope"
        prm("inSR") = CStr(outSR)
        prm("spatialRel") = "esriSpatialRelIntersects"
        prm("outFields") = ROUTE_ITEM_FIELDS
        prm("returnGeometry") = "true"
        prm("outSR") = CStr(outSR)
        prm("maxAllowableOffset") = "0.00001"
        prm("geometryPrecision") = "7"
        prm("orderByFields") = "OBJECTID"
        prm("resultOffset") = CStr(offset)
        prm("resultRecordCount") = "2000"
        Set js = ArcQuery(LRS_URL, prm)
        feats = DGet(js, "features", Array())
        For f = 0 To UBound(feats)
            Set a = feats(f)("attributes")
            If Len(CStr(DGet(a, "RTE_NM", ""))) Then
                If Not gRouteAttrs.Exists(CStr(DGet(a, "RTE_NM", ""))) Then Set gRouteAttrs(CStr(DGet(a, "RTE_NM", ""))) = a
            End If
            If feats(f).Exists("geometry") Then
                If IsObject(feats(f)("geometry")) Then
                    Set g = feats(f)("geometry")
                    paths = DGet(g, "paths", Array())
                    For pi = 0 To UBound(paths)
                        pth = paths(pi)
                        If UBound(pth) >= 1 Then
                            PolyInit p, UBound(pth) + 1
                            For j = 0 To UBound(pth)
                                PolyAdd p, pth(j)(0), pth(j)(1)
                            Next
                            p.Cat = CStr(DGet(a, "RTE_CATEGORY_NM", ""))
                            p.Nm = CStr(DGet(a, "RTE_NM", ""))
                            AppendPoly routes, n, p
                        End If
                    Next
                End If
            End If
        Next
        offset = offset + UBound(feats) + 1
        fetched = offset
        SetStatus "Downloading LRS routes... " & offset & " features"
        If UBound(feats) < 0 Then Exit Do
        If Not CBool(DGet(js, "exceededTransferLimit", False)) Then Exit Do
    Loop
    GetRoutes = n
End Function
