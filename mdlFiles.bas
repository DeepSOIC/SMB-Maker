Attribute VB_Name = "mdlFiles"
Option Explicit


Function ValFolder(ByVal strFolder As String) As String
If Right$(strFolder, 1) <> "\" Then
    ValFolder = strFolder + "\"
Else
    ValFolder = strFolder
End If
End Function

Function AppPath() As String
Static Foo As Boolean
Static AP As String
If Not Foo Then
    AP = ValFolder(App.Path)
    Foo = True
End If
AppPath = AP
End Function

'Example: D:\tmp.bmp -> tmp.bmp
Function GetFileTitle(ByRef strPath As String) As String
Dim Pos As Long
Pos = InStrRev(strPath, "\")
GetFileTitle = Mid$(strPath, Pos + 1)
End Function

Function CropExt(ByRef strFileName As String) As String
Dim tmp As String
Const Slash As String = "."
Dim i As Long, i1 As Long
tmp = GetFileTitle(strFileName)
i = 0
i1 = 0
Do
    i1 = InStr(i + 1, tmp, Slash)
    If i1 > 0 Then
        i = i1
    End If
Loop Until i1 = 0
If i = 0 Then i = Len(tmp) + 1
CropExt = Left$(strFileName, (Len(strFileName) - Len(tmp)) + i - 1)
End Function

Function GetExt(ByRef strPath As String) As String
Dim tmp As String
Dim Pos As Long
tmp = GetFileTitle(strPath)
Pos = InStrRev(tmp, ".")
If Pos > 0 Then
    GetExt = Right$(tmp, Len(tmp) - Pos)
End If
End Function

Public Function dbChDir(ByVal strPath As String)
Dim Drv As String
strPath = Trim(strPath)
If Mid$(strPath, 2, 1) = ":" Then
    Drv = Mid$(strPath, 1, 3)
    ChDrive Drv
End If
ChDir ValFolder(strPath)
End Function

Public Function GetDirName(ByVal strPath As String) As String
Dim i As Long, m As String
i = InStrRev(strPath, "\")
GetDirName = Left$(strPath, i)
End Function

Public Function FolderExists(ByRef strFolder As String) As Boolean
FolderExists = FileFolderExists(ValFolder(strFolder))
End Function

Public Function TempPath() As String
Dim strBuffer As String
Dim BufLen As Long
BufLen = 1024
strBuffer = Space$(BufLen)
BufLen = APIGetTempPath(BufLen, ByVal strBuffer)
TempPath = ValFolder(Mid$(strBuffer, 1, BufLen)) + App.ProductName + "\"
End Function

Public Sub CreateFolder(ByRef strPath As String)
Dim Spl() As String, i As Long, tmp As String
Dim Fld As String
If Len(strPath) = 0 Then Exit Sub
If Right$(strPath, 1) = "\" Then
    Fld = Left$(strPath, Len(strPath) - 1)
Else
    Fld = strPath
End If
Spl = Split(Fld, "\")
tmp = Spl(0)
If InStr(tmp, ":") = 0 Then
    tmp = ValFolder(CurDir) + tmp
    If Not FileFolderExists(tmp + "\") Then
        MkDir tmp
    End If
End If
For i = 1 To UBound(Spl)
    tmp = tmp + "\" + Spl(i)
    If Not FileFolderExists(tmp + "\") Then
        MkDir tmp
    End If
Next i
End Sub

Public Function FileFolderExists(ByRef Folder As String) As Boolean
Dim Fld As String
Fld = GetDirName(Folder)
If Len(Fld) = 0 Then
    FileFolderExists = True
Else
    On Error Resume Next
    Err.Clear
    FileFolderExists = Len(Dir(Fld, vbDirectory Or vbHidden Or vbReadOnly Or vbSystem)) > 0
    If Err.Number <> 0 Then FileFolderExists = False
    'tests for .\ directory
End If
End Function

Public Sub WinRun(Path As String, Optional ShowWindow As Integer = 1)
    ShellExecute 0, "open", Path, "", "", ShowWindow
End Sub

Public Function FileExists(ByRef File As String) As Boolean
Dim nmb As Long
On Error GoTo eh
FileExists = True
nmb = FreeFile
Open File For Input As nmb
Close nmb
Exit Function
eh:
FileExists = False
End Function

Public Function ShowSaveDlg(ByVal Filter As DlgFilter, _
                            ByVal hWndOwner As Long, _
                            Optional ByVal OpenFlags As dhFileOpenConstants, _
                            Optional ByRef InitFileName As String = vbNullString, _
                            Optional ByVal RaiseErrors As Boolean = True, _
                            Optional ByVal NoSaveInitDir As Boolean = False, _
                            Optional ByRef Purpose As String)
Dim UB As Long
Dim i As Long
UB = -1
If AryDims(AryPtr(InitDirs)) = 1 Then
    UB = UBound(InitDirs)
End If
If UB = -1 Then
    ReDim InitDirs(0 To 0)
End If
On Error GoTo eh
With CDl
    .CancelError = True
    .DialogTitle = ""
    .hWndOwner = hWndOwner
    .FileName = InitFileName
    If Purpose = "" Then
        If UBound(InitDirs) < Filter Then
            ReDim Preserve InitDirs(0 To Filter)
            For i = 0 To UBound(InitDirs)
                If Len(InitDirs(i)) = 0 Then
                    InitDirs(i) = dbGetSetting("InitDirs", "Dir" + CStr(i), AppPath)
                End If
            Next i
        End If
        .InitDir = InitDirs(Filter)
    Else
        .InitDir = GetSMBCurDir(Purpose)
    End If
    .Filter = GetDlgFilter(Filter)
    .Flags = 0&
    .OpenFlags = OpenFlags Xor (cdlOFNHideReadOnly Or cdlOFNNoReadOnlyReturn Or cdlOFNOverwritePrompt)
    .ShowSave
    ShowSaveDlg = .FileName
    .InitDir = GetDirName(.FileName)
    If Not NoSaveInitDir Then
        If Purpose = "" Then
            InitDirs(Filter) = .InitDir
        Else
            SaveSMBCurDir ValFolder(.InitDir), Purpose
        End If
    End If
End With

If Purpose = "" Then
    For i = 0 To UBound(InitDirs)
        dbSaveSetting "InitDirs", "Dir" + CStr(i), InitDirs(i)
    Next i
End If

Exit Function
eh:
If Err.Number = dbCWS Then
    If RaiseErrors Then
        Err.Raise Err.Number, Err.Source, Err.Description
    Else
        ShowSaveDlg = vbNullString
    End If
Else
    Err.Raise Err.Number, Err.Source, Err.Description
End If
End Function

'if allowmultiselect flag is on, it returns a list
'of files, separated by Chr$(1)
Public Function ShowOpenDlg(ByVal Filter As DlgFilter, _
                            ByVal hWndOwner As Long, _
                            Optional ByVal OpenFlags As dhFileOpenConstants, _
                            Optional ByRef InitFileName As String = vbNullString, _
                            Optional ByVal RaiseErrors As Boolean = True, _
                            Optional ByVal NoSaveInitDir As Boolean = False, _
                            Optional ByVal Purpose As String) As String
Dim UB As Long
Dim i As Long
Dim Files() As String
Dim Out As String
UB = -1
If AryDims(AryPtr(InitDirs)) = 1 Then
    UB = UBound(InitDirs)
End If
If UB = -1 Then
    ReDim InitDirs(0 To 0)
End If
On Error GoTo eh
With CDl
    .CancelError = True
    .DialogTitle = ""
    .hWndOwner = hWndOwner
    .FileName = InitFileName
    If Purpose = "" Then
      Debug.Assert False
        If UBound(InitDirs) < Filter Then
            ReDim Preserve InitDirs(0 To Filter)
            For i = 0 To UBound(InitDirs)
                If Len(InitDirs(i)) = 0 Then
                    InitDirs(i) = dbGetSetting("InitDirs", "Dir" + CStr(i), AppPath)
                End If
            Next i
        End If
    End If
    If Purpose = "" Then
        .InitDir = InitDirs(Filter)
    Else
        .InitDir = GetSMBCurDir(Purpose)
    End If
    .Filter = GetDlgFilter(Filter)
    .Flags = 0&
    .OpenFlags = OpenFlags Xor (cdlOFNFileMustExist Or cdlOFNHideReadOnly)
    
    .ShowOpen
    
    If (OpenFlags And cdlOFNAllowMultiselect) = cdlOFNAllowMultiselect Then
        Files = .FileList
        Files(0) = ValFolder(Files(0))
        Out = Files(0) + Files(1)
        For i = 2 To UBound(Files)
            Out = Out + Chr$(1) + Files(0) + Files(i)
        Next i
        ShowOpenDlg = Out
        .InitDir = Files(0)
        Erase Files
    Else
        ShowOpenDlg = .FileName
        .InitDir = GetDirName(.FileName)
    End If
    If Not NoSaveInitDir Then
        If Purpose = "" Then
            InitDirs(Filter) = .InitDir
        Else
            SaveSMBCurDir ValFolder(.InitDir), Purpose
        End If
    End If
End With

If Purpose = "" Then
    For i = 0 To UBound(InitDirs)
        dbSaveSetting "InitDirs", "Dir" + CStr(i), InitDirs(i)
    Next i
End If

Exit Function
eh:
If Err.Number = dbCWS Then
    If RaiseErrors Then
        Err.Raise Err.Number, Err.Source, Err.Description
    Else
        ShowOpenDlg = vbNullString
    End If
Else
    Err.Raise Err.Number, Err.Source, Err.Description
End If
End Function

Function GetDlgFilter(ind As DlgFilter) As String
GetDlgFilter = GRSF(1000 + ind)
End Function


'always ends with \
Public Function GetSMBCurDir(Optional ByRef Purpose As String) As String
Dim Result As String
Dim DefPath As String
If Len(Purpose) > 0 Then
    Result = dbGetSetting("Formats", _
                          "Last directory for " + Purpose, _
                          DefValue:=ExePath)
Else
    Result = dbGetSetting("Formats", "Last directory")
End If
If Len(Result) = 0 Then Result = ValFolder(CurDir)
GetSMBCurDir = Result
End Function

'if NewDir contains a path to a file, only the directory is used.
'The directory must end with a backslash
Public Sub SaveSMBCurDir(ByRef NewDir As String, _
                         Optional ByRef Purpose As String)
Dim Pos As Long
Pos = InStrRev(NewDir, "\")
If Pos > 0 Then
    If Len(Purpose) > 0 Then
        dbSaveSetting "Formats", "Last directory for " + Purpose, Left$(NewDir, Pos)
    Else
        dbSaveSetting "Formats", "Last directory", Left$(NewDir, Pos)
    End If
End If
End Sub

