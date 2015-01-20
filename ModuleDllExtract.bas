Attribute VB_Name = "ModuleDllExtract"
Option Explicit

Private Declare Function FindResource Lib "kernel32" Alias "FindResourceA" _
        (ByVal hInstance As Long, ByVal lpName As String, ByVal lpType As String) As Long
Private Declare Function LoadResource Lib "kernel32" _
        (ByVal hInstance As Long, ByVal hResInfo As Long) As Long
Private Declare Function LockResource Lib "kernel32" _
        (ByVal hResData As Long) As Long
Private Declare Function SizeofResource Lib "kernel32" _
        (ByVal hInstance As Long, ByVal hResInfo As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
        (Destination As Any, Source As Any, ByVal Length As Long)
        
Public Sub ExtractDll(ByRef File As String, _
                      ByRef ResType As String, _
                      ByRef ResID As String)
Dim hResource As Long
Dim hData As Long
Dim ptrData As Long
Dim ResSize As Long
Dim ResData() As Byte
Dim nmb As Long
Dim Answ As VbMsgBoxResult
On Error GoTo eh
If FileExists(File) Then
    Exit Sub
End If
hResource = FindResource(App.hInstance, ResID, ResType)
If hResource = 0 Then
    Err.Raise 119, "ExtractDll", "Cannot open the resource. (FindResource failed.)"
    Exit Sub
End If
ResSize = SizeofResource(App.hInstance, hResource)
If ResSize = 0 Then
    Err.Raise 119, "ExtractDll", "Cannot determine the size of the resource. (SizeOfResource failed.)"
    Exit Sub
End If
hData = LoadResource(App.hInstance, hResource)
If hData = 0 Then
    Err.Raise 119, "ExtractDll", "Cannot load the resource. (LoadResource failed.)"
    Exit Sub
End If
ptrData = LockResource(hData)
If ptrData = 0 Then
    Err.Raise 119, "ExtractDll", "Cannot lock the resource. (LockResource failed.)"
    Exit Sub
End If
ReDim ResData(0 To ResSize - 1)
CopyMemory ResData(0), ByVal ptrData, ResSize

If Not StartWrite(File) Then
    Err.Raise 119, "ExtractDll", "Cannot open the file for writing."
    Exit Sub
End If

nmb = FreeFile
Open File For Binary Access Write As nmb
    Put nmb, 1, ResData
Close nmb

Erase ResData
Exit Sub
Resume
eh:
ErrRaise "ExtractDll"
End Sub


Public Sub CheckDll(ByRef DLLName As String)
Dim Found As Boolean
Static PathTo As String
'If Not Found Then
    PathTo = dbGetSettingEx("Setup", DLLName + " path", vbString, "", AllUsers:=True)
    If Len(PathTo) = 0 Or Not FolderExists(PathTo) Then
        PathTo = AppPath
        If Not FileExists(AppPath + DLLName) Then
            If Not FileExists(AppPath + "dll.no") Then
                On Error Resume Next
                ExtractDll AppPath + DLLName, "DLL", DLLName
                On Error GoTo 0
            End If
        End If
    End If
    dbChDir PathTo
    Do While Not FileExists(DLLName)
        On Error GoTo eh
        CurDll = DLLName 'for inputbox browse button
        '2384=The dynamic link library named $dll is not found or does not work. Please specify the path, where it can be found.\n\nNote that you cannot change it's name. It is fixed.`DLL not found
        PathTo = dbInputBox(grs(2384, "$dll", DLLName), PathTo, CancelError:=True, BrowseButton:=True)
        dbChDir PathTo
    Loop
    Found = True
    dbSaveSettingEx "Setup", DLLName + " path", PathTo, AllUsers:=True
    On Error GoTo 0

Exit Sub

eh:
If Err.Number = dbCWS Then
    Err.Raise dbCWS, "CheckDll", "Cancel was selected."
Else
    Resume Next
End If
End Sub

