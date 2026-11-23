Attribute VB_Name = "VdotProjectMapper"
'======================================================================
' VDOT Project Mapper - MicroStation VBA launcher
' Created by Robert Ratliff  -  Build 2026-11-19
'
' Starts the VDOT Project Mapper (vdot_project_mapper.py) inside
' MicroStation / OpenRoads Designer, so the UI and behaviour are exactly
' the Python tool's (block fence in the PDF viewer, scale bar & north
' arrow, county mask, hidden PDF, light/dark themes).
' Older copies of the script (without the block fence) are skipped.
'
' Run:    vba run [VDOT_Project_Mapper_Launcher]VdotProjectMapper.Main
'   or    Utilities > Macro > Macros... > VDOT_Project_Mapper_Launcher > Main
'======================================================================
Option Explicit

Private Const SCRIPT_NAME As String = "vdot_project_mapper.py"
Private Const APP_TITLE As String = "VDOT Project Mapper"
Private Const REG_APP As String = "VDOT_Project_Mapper_Launcher"
' Text only the current script version contains (block fence tool).
Private Const REQUIRED_MARK As String = "MODE_FENCE"

Private mLog As String
Private mOldPath As String

Public Sub Main()
    Dim path As String
    mLog = "VDOT_Georef launcher " & Now & vbCrLf
    path = FindScript()
    If Len(path) = 0 Then path = BrowseForScript()
    If Len(path) = 0 And Len(mOldPath) > 0 Then
        If MsgBox("Only an older copy of " & SCRIPT_NAME & " was found (no block fence tool):" & vbCrLf & vbCrLf & _
                  mOldPath & vbCrLf & vbCrLf & _
                  "Run Install.bat from the latest installer to update it." & vbCrLf & _
                  "Start the older version anyway?", vbExclamation + vbYesNo, APP_TITLE) = vbYes Then path = mOldPath
    End If
    WriteLog
    If Len(path) = 0 Then
        MsgBox SCRIPT_NAME & " was not found." & vbCrLf & vbCrLf & _
               "Run Install.bat from the VDOT_Project_Mapper_Installer package, " & _
               "or copy the script into the WorkSet or Organization Macros folder." & vbCrLf & vbCrLf & _
               "Folders searched are listed in " & LogPath(), vbExclamation, APP_TITLE
        Exit Sub
    End If
    SaveSetting REG_APP, "Launcher", "ScriptPath", path
    CadInputQueue.SendKeyin "python load " & QuoteIfNeeded(path)
End Sub

' Looks in the remembered location, the Python/VBA/macro search paths,
' the Organization and WorkSet Macros folders, then every Bentley product.
Private Function FindScript() As String
    Dim vars As Variant, v As Variant, p As String

    p = GetSetting(REG_APP, "Launcher", "ScriptPath", "")
    If Len(p) > 0 Then If TryFile(p) Then FindScript = p: Exit Function

    vars = Array("MS_PYTHONSEARCHDIRECTORIES", "MS_PYTHONSCRIPTS", "MS_VBASEARCHDIRECTORIES", _
                 "MS_MACRO", "_USTN_WORKSETSTANDARDS", "_USTN_WORKSPACESTANDARDS", "_USTN_ORGANIZATION")
    For Each v In vars
        p = SearchList(CfgValue(CStr(v)))
        If Len(p) > 0 Then FindScript = p: Exit Function
    Next v

    p = SearchBentleyFolders()
    If Len(p) > 0 Then FindScript = p: Exit Function
    FindScript = ""
End Function

Private Function CfgValue(ByVal name As String) As String
    On Error Resume Next
    If ActiveWorkspace.IsConfigurationVariableDefined(name) Then
        CfgValue = ActiveWorkspace.ExpandConfigurationVariable("$(" & name & ")")
        If Len(CfgValue) = 0 Then CfgValue = ActiveWorkspace.ConfigurationVariableValue(name, True)
    End If
    On Error GoTo 0
End Function

' Checks each folder in a ";"-separated list, and its Macros subfolder.
Private Function SearchList(ByVal list As String) As String
    Dim d As Variant, dirPath As String
    For Each d In Split(list, ";")
        dirPath = AddSlash(Trim$(Replace(CStr(d), "/", "\")))
        If Len(dirPath) > 1 Then
            If TryFile(dirPath & SCRIPT_NAME) Then SearchList = dirPath & SCRIPT_NAME: Exit Function
            If TryFile(dirPath & "Macros\" & SCRIPT_NAME) Then SearchList = dirPath & "Macros\" & SCRIPT_NAME: Exit Function
        End If
    Next d
End Function

' C:\ProgramData\Bentley\<product>\Configuration\Organization\Macros and
' ...\WorkSpaces\<ws>\WorkSets\<set>\Standards\Macros for every product.
Private Function SearchBentleyFolders() As String
    Dim fso As Object, prod As Object, ws As Object, wset As Object, cfg As String
    On Error GoTo done
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists("C:\ProgramData\Bentley") Then GoTo done
    For Each prod In fso.GetFolder("C:\ProgramData\Bentley").SubFolders
        cfg = prod.path & "\Configuration\"
        If TryFile(cfg & "Organization\Macros\" & SCRIPT_NAME) Then SearchBentleyFolders = cfg & "Organization\Macros\" & SCRIPT_NAME: Exit Function
        If fso.FolderExists(cfg & "WorkSpaces") Then
            For Each ws In fso.GetFolder(cfg & "WorkSpaces").SubFolders
                If TryFile(ws.path & "\Standards\Macros\" & SCRIPT_NAME) Then SearchBentleyFolders = ws.path & "\Standards\Macros\" & SCRIPT_NAME: Exit Function
                If fso.FolderExists(ws.path & "\WorkSets") Then
                    For Each wset In fso.GetFolder(ws.path & "\WorkSets").SubFolders
                        If TryFile(wset.path & "\Standards\Macros\" & SCRIPT_NAME) Then SearchBentleyFolders = wset.path & "\Standards\Macros\" & SCRIPT_NAME: Exit Function
                    Next wset
                End If
            Next ws
        End If
    Next prod
done:
End Function

' Last resort: let the user pick the folder that holds the script.
Private Function BrowseForScript() As String
    Dim sh As Object, f As Object, p As String
    On Error GoTo done
    If MsgBox(SCRIPT_NAME & " was not found automatically." & vbCrLf & vbCrLf & _
              "Browse to the folder that contains it?", vbQuestion + vbYesNo, APP_TITLE) <> vbYes Then GoTo done
    Set sh = CreateObject("Shell.Application")
    Set f = sh.BrowseForFolder(0, "Select the folder containing " & SCRIPT_NAME, &H10 Or &H40, 17)
    If f Is Nothing Then GoTo done
    p = AddSlash(f.Self.path) & SCRIPT_NAME
    If TryFile(p) Then BrowseForScript = p
done:
End Function

Private Function TryFile(ByVal p As String) As Boolean
    Dim fso As Object
    On Error Resume Next
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim exists As Boolean, current As Boolean
    exists = fso.FileExists(p)
    If Err.Number <> 0 Then
        Err.Clear
        exists = (Len(Dir$(p, vbNormal)) > 0)
    End If
    On Error GoTo 0
    If exists Then current = IsCurrent(p)
    If exists And Not current And Len(mOldPath) = 0 Then mOldPath = p
    TryFile = current
    mLog = mLog & IIf(current, "FOUND    ", IIf(exists, "OUTDATED ", "missing  ")) & p & vbCrLf
End Function

' True when the script contains the current-version marker.
Private Function IsCurrent(ByVal p As String) As Boolean
    Dim fh As Integer, txt As String
    On Error GoTo bad
    fh = FreeFile
    Open p For Binary Access Read As #fh
    txt = Space$(LOF(fh))
    Get #fh, , txt
    Close #fh
    IsCurrent = (InStr(1, txt, REQUIRED_MARK, vbBinaryCompare) > 0)
    Exit Function
bad:
    On Error Resume Next
    Close #fh
End Function

Private Function AddSlash(ByVal p As String) As String
    If Len(p) > 0 And Right$(p, 1) <> "\" Then p = p & "\"
    AddSlash = p
End Function

Private Function QuoteIfNeeded(ByVal p As String) As String
    If InStr(p, " ") > 0 Then QuoteIfNeeded = """" & p & """" Else QuoteIfNeeded = p
End Function

Private Function LogPath() As String
    LogPath = Environ$("TEMP") & "\vdot_georef_vba.log"
End Function

Private Sub WriteLog()
    Dim fh As Integer
    On Error Resume Next
    fh = FreeFile
    Open LogPath() For Output As #fh
    Print #fh, mLog
    Close #fh
End Sub