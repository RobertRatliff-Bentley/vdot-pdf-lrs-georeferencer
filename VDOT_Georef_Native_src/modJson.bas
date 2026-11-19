Attribute VB_Name = "modJson"
'Minimal, fast JSON parser for UTF-8 byte arrays.
'Objects -> Scripting.Dictionary, arrays -> Variant(), numbers -> Double, true/false -> Boolean, null -> Null.
Option Explicit

Private b() As Byte
Private p As Long
Private n As Long

Public Function JsonParseBytes(bytes() As Byte) As Variant
    b = bytes
    n = UBound(b) - LBound(b) + 1
    p = 0
    If n >= 3 Then
        If b(0) = &HEF And b(1) = &HBB And b(2) = &HBF Then p = 3
    End If
    Dim v As Variant
    ParseValue v
    If IsObject(v) Then Set JsonParseBytes = v Else JsonParseBytes = v
    Erase b
End Function

Public Function JsonParseString(ByVal s As String) As Variant
    Dim by() As Byte
    by = Utf8Bytes(s)
    b = by
    n = UBound(b) + 1
    p = 0
    Dim v As Variant
    ParseValue v
    Erase b
    If IsObject(v) Then Set JsonParseString = v Else JsonParseString = v
End Function

Public Function Utf8Bytes(ByVal s As String) As Byte()
    Dim out() As Byte, i As Long, c As Long, k As Long
    If Len(s) = 0 Then
        ReDim out(0 To 0)
        Utf8Bytes = out
        Exit Function
    End If
    ReDim out(0 To Len(s) * 3)
    For i = 1 To Len(s)
        c = AscW(Mid$(s, i, 1)) And &HFFFF&
        If c < &H80 Then
            out(k) = c: k = k + 1
        ElseIf c < &H800 Then
            out(k) = &HC0 Or (c \ &H40): out(k + 1) = &H80 Or (c And &H3F): k = k + 2
        Else
            out(k) = &HE0 Or (c \ &H1000): out(k + 1) = &H80 Or ((c \ &H40) And &H3F)
            out(k + 2) = &H80 Or (c And &H3F): k = k + 3
        End If
    Next
    ReDim Preserve out(0 To k - 1)
    Utf8Bytes = out
End Function

Private Sub SkipWs()
    Do While p < n
        Select Case b(p)
        Case 32, 9, 10, 13
            p = p + 1
        Case Else
            Exit Sub
        End Select
    Loop
End Sub

Private Sub ParseValue(ByRef v As Variant)
    SkipWs
    If p >= n Then Err.Raise vbObjectError + 513, "JSON", "Unexpected end of JSON"
    Select Case b(p)
    Case 123: Set v = ParseObject()
    Case 91: v = ParseArray()
    Case 34: v = ParseString()
    Case 116: p = p + 4: v = True
    Case 102: p = p + 5: v = False
    Case 110: p = p + 4: v = Null
    Case Else: v = ParseNumber()
    End Select
End Sub

Private Function ParseObject() As Object
    Dim d As Object, k As String, v As Variant
    Set d = CreateObject("Scripting.Dictionary")
    p = p + 1
    SkipWs
    If b(p) = 125 Then
        p = p + 1
        Set ParseObject = d
        Exit Function
    End If
    Do
        SkipWs
        k = ParseString()
        SkipWs
        p = p + 1
        v = Empty
        ParseValue v
        If IsObject(v) Then Set d(k) = v Else d(k) = v
        SkipWs
        If p >= n Then Exit Do
        If b(p) = 44 Then
            p = p + 1
        Else
            p = p + 1
            Exit Do
        End If
    Loop
    Set ParseObject = d
End Function

Private Function ParseArray() As Variant
    Dim a() As Variant, cnt As Long
    ReDim a(0 To 7)
    p = p + 1
    SkipWs
    If b(p) = 93 Then
        p = p + 1
        ParseArray = Array()
        Exit Function
    End If
    Do
        If cnt > UBound(a) Then ReDim Preserve a(0 To 2 * cnt + 1)
        ParseValue a(cnt)
        cnt = cnt + 1
        SkipWs
        If p >= n Then Exit Do
        If b(p) = 44 Then
            p = p + 1
        Else
            p = p + 1
            Exit Do
        End If
    Loop
    ReDim Preserve a(0 To cnt - 1)
    ParseArray = a
End Function

Private Function ParseString() As String
    Dim s As String, c As Long, c2 As Long, c3 As Long, st As Long
    p = p + 1
    st = p
    Do While p < n
        c = b(p)
        If c = 34 Or c = 92 Or c >= 128 Then Exit Do
        p = p + 1
    Loop
    If p > st Then s = AsciiSlice(st, p - st)
    If p < n Then
        If b(p) = 34 Then
            p = p + 1
            ParseString = s
            Exit Function
        End If
    End If
    Do While p < n
        c = b(p)
        If c = 34 Then
            p = p + 1
            Exit Do
        ElseIf c = 92 Then
            p = p + 1
            c = b(p)
            Select Case c
            Case 110: s = s & vbLf
            Case 114: s = s & vbCr
            Case 116: s = s & vbTab
            Case 98: s = s & Chr$(8)
            Case 102: s = s & Chr$(12)
            Case 117
                s = s & ChrW(CLng("&H" & Chr$(b(p + 1)) & Chr$(b(p + 2)) & Chr$(b(p + 3)) & Chr$(b(p + 4))))
                p = p + 4
            Case Else: s = s & Chr$(c)
            End Select
            p = p + 1
        ElseIf c < 128 Then
            s = s & Chr$(c)
            p = p + 1
        ElseIf c < &HE0 Then
            c2 = b(p + 1)
            s = s & ChrW(((c And &H1F) * 64) Or (c2 And &H3F))
            p = p + 2
        ElseIf c < &HF0 Then
            c2 = b(p + 1): c3 = b(p + 2)
            s = s & ChrW(((c And &HF) * 4096) Or ((c2 And &H3F) * 64) Or (c3 And &H3F))
            p = p + 3
        Else
            s = s & "?"
            p = p + 4
        End If
    Loop
    ParseString = s
End Function

Private Function AsciiSlice(ByVal st As Long, ByVal ln As Long) As String
    Dim t() As Byte, i As Long
    ReDim t(0 To ln - 1)
    For i = 0 To ln - 1
        t(i) = b(st + i)
    Next
    AsciiSlice = StrConv(t, vbUnicode)
End Function

Private Function ParseNumber() As Double
    Dim neg As Boolean, m As Double, fd As Long, ex As Long, exNeg As Boolean, c As Long
    If b(p) = 45 Then neg = True: p = p + 1
    If b(p) = 43 Then p = p + 1
    Do While p < n
        c = b(p)
        If c < 48 Or c > 57 Then Exit Do
        m = m * 10# + (c - 48)
        p = p + 1
    Loop
    If p < n Then
        If b(p) = 46 Then
            p = p + 1
            Do While p < n
                c = b(p)
                If c < 48 Or c > 57 Then Exit Do
                m = m * 10# + (c - 48)
                fd = fd + 1
                p = p + 1
            Loop
        End If
    End If
    If p < n Then
        If b(p) = 101 Or b(p) = 69 Then
            p = p + 1
            If b(p) = 45 Then
                exNeg = True: p = p + 1
            ElseIf b(p) = 43 Then
                p = p + 1
            End If
            Do While p < n
                c = b(p)
                If c < 48 Or c > 57 Then Exit Do
                ex = ex * 10 + (c - 48)
                p = p + 1
            Loop
        End If
    End If
    If exNeg Then ex = -ex
    ex = ex - fd
    If ex <> 0 Then m = m * (10# ^ ex)
    If neg Then m = -m
    ParseNumber = m
End Function

'---------------------------------------------------------------- writing helpers
Public Function JStr(ByVal s As String) As String
    Dim i As Long, c As Long, o As String
    For i = 1 To Len(s)
        c = AscW(Mid$(s, i, 1)) And &HFFFF&
        Select Case c
        Case 34: o = o & "\"""
        Case 92: o = o & "\\"
        Case 10: o = o & "\n"
        Case 13: o = o & "\r"
        Case 9: o = o & "\t"
        Case Is < 32, Is > 126: o = o & "\u" & Right$("000" & LCase$(Hex$(c)), 4)
        Case Else: o = o & ChrW(c)
        End Select
    Next
    JStr = """" & o & """"
End Function

Public Function JNum(ByVal v As Double) As String
    Dim s As String
    s = Trim$(Str$(v))
    If Left$(s, 1) = "." Then s = "0" & s
    If Left$(s, 2) = "-." Then s = "-0" & Mid$(s, 2)
    JNum = s
End Function

Public Function JVal(v As Variant) As String
    If IsNull(v) Or IsEmpty(v) Then
        JVal = "null"
    ElseIf VarType(v) = vbBoolean Then
        JVal = IIf(v, "true", "false")
    ElseIf IsNumeric(v) And VarType(v) <> vbString Then
        JVal = JNum(CDbl(v))
    Else
        JVal = JStr(CStr(v))
    End If
End Function

Public Function JDictFlat(ByVal d As Object) As String
    Dim k As Variant, o As String
    For Each k In d.Keys
        If Len(o) Then o = o & ", "
        o = o & JStr(CStr(k)) & ": " & JVal(d(k))
    Next
    JDictFlat = "{" & o & "}"
End Function

Public Function DGet(ByVal d As Object, ByVal key As String, Optional ByVal dflt As Variant = "") As Variant
    If d Is Nothing Then
        DGet = dflt
    ElseIf d.Exists(key) Then
        If IsNull(d(key)) Then DGet = dflt Else DGet = d(key)
    Else
        DGet = dflt
    End If
End Function
