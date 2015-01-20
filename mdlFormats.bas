Attribute VB_Name = "mdlFormats"
Option Explicit
Public FormatList() As FormatTemplate
Public nFormats As Long

Public Const errNewFile = 1597

Public Sub ConnectFormat(ByRef Format As FormatTemplate)
Dim ID As String
Format.GetInfo ID, "", False
On Error Resume Next
Err.Clear
FormatFromID ID
If Err.Number <> 1212 Then
    On Error GoTo 0
    Err.Raise 1212, "ConnectFormat", "Format already connected: " + ID
End If
On Error GoTo 0
If nFormats = 0 Then
    ReDim FormatList(0 To 0)
Else
    ReDim Preserve FormatList(0 To nFormats)
End If
Set FormatList(nFormats) = Format
nFormats = nFormats + 1
End Sub

Public Function FormatFromID(ByRef ID As String) As FormatTemplate
Set FormatFromID = FormatList(FormatIndexFromID(ID))
End Function

Public Function FormatIndexFromID(ByRef ID As String) As Long
Dim i As Long
Dim tID As String
ID = UCase$(ID)
ConnectFormats
For i = 0 To nFormats - 1
    FormatList(i).GetInfo tID, "", False
    If UCase$(tID) = ID Then
        Exit For
    End If
Next i
If i >= nFormats Then Err.Raise 1212, "FormatFromID", "Format not found: " + ID
FormatIndexFromID = i
End Function

Public Sub WriteAlphaToData(ByRef Data() As Long, _
                            ByRef Alpha() As Long, _
                            ByRef AlphaPresent As Boolean)
Dim i As Long
Dim w As Long, h As Long
Dim wa As Long, ha As Long
Dim RGBData() As RGBQUAD
Dim RGBAlpha() As RGBQUAD
Dim Length As Long
TestDims Data
AlphaPresent = AryDims(AryPtr(Alpha)) = 2
If AlphaPresent Then
    AryWH AryPtr(Data), w, h
    AryWH AryPtr(Alpha), wa, ha
    
    dbStretch Alpha, w, h, RaiseErrors:=True
    
    Length = w * h
    On Error GoTo eh
    ConstructAry AryPtr(RGBData), VarPtr(Data(0, 0)), 4, Length
    ConstructAry AryPtr(RGBAlpha), VarPtr(Alpha(0, 0)), 4, Length
    For i = 0 To Length - 1
        RGBData(i).rgbReserved = (CLng(RGBAlpha(i).rgbBlue) + _
                                  CLng(RGBAlpha(i).rgbGreen) + _
                                  CLng(RGBAlpha(i).rgbRed)) \ 3
    Next i
    UnReferAry AryPtr(RGBData)
    UnReferAry AryPtr(RGBAlpha)
End If
Exit Sub
eh:
UnReferAry AryPtr(RGBData)
UnReferAry AryPtr(RGBAlpha)
ErrRaise "WriteAlphaToData"
End Sub

Public Sub CleanData(ByRef Data() As Long)
vtRepair Data
End Sub

Public Sub vtSavePicture(ByRef Data() As Long, _
                         ByRef Alpha() As Long, _
                         ByRef FileName As String, _
                         Optional ByRef FormatID As String, _
                         Optional ByVal ShowDialog = True, _
                         Optional ByRef Purpose As String)
Dim frm As frmFileSave
Dim Format As FormatTemplate
Dim FormatFailed As Boolean

On Error Resume Next
If Not ShowDialog Then
    Set Format = FormatFromID(FormatID)
    If Not Format.CanSave(AryDims(AryPtr(Alpha)), 0) Then
        FormatFailed = True
    End If
End If
On Error GoTo 0
If ShowDialog Or FormatFailed Then
rsmDialog:
    Set frm = New frmFileSave
    Load frm
    On Error GoTo eh1
    With frm
        .SetPtrData AryPtr(Data)
        .AlphaExchange Alpha
        .SetFormatID FormatID
        .SetPurpose Purpose
        .SetFileName FileName
        .Show vbModal 'and save!!!
        If Len(.Tag) > 0 Then
            Err.Raise dbCWS, "SavePicture"
        End If
        FileName = .SavedFileName
        FormatID = .SavedFormatID
        .AlphaExchange Alpha
    End With
Else
    dbSave FileName, AryPtr(Data), Alpha, Format
End If
On Error Resume Next
ShowStatus grs(1201, "%FS", Format_Size(FileLen(FileName), 1024, 3)), HoldTime:=2
Exit Sub
eh1:
PushError
frm.RemovePtrData
frm.AlphaExchange Alpha
Unload frm
PopError
ErrRaise "SavePicture"
End Sub

Public Sub ConnectFormats()
Static Foo As Boolean
Dim bmp As New clsFormatBmp
Dim png As New clsFormatPNG
Dim jpg As New clsFormatJPEG
Dim smb As New clsFormatSMB
Dim pal As New clsFormatPal
Dim vb As New clsFormatVB
If Not Foo Then
    Foo = True
    ConnectFormat bmp
    ConnectFormat png
    ConnectFormat jpg
    ConnectFormat smb
    ConnectFormat pal
    
    ConnectFormat vb 'must be the last
End If
End Sub


Public Sub dbSave(ByRef File As String, _
                  ByVal ptrData As Long, _
                  ByRef Alpha() As Long, _
                  ByRef Format As FormatTemplate)
Dim AlphaPresent As Boolean
Dim Data() As Long
If Format.AlphaSupproted Then
    ReferAry AryPtr(Data), ptrData
    On Error GoTo eh
    WriteAlphaToData Data, Alpha, AlphaPresent
End If
Format.LoadSettings
Format.SetPtrData ptrData
Format.SaveFile File, AlphaPresent
Format.RemovePtrData
If AlphaPresent Then
    CleanData Data
End If
UnReferAry AryPtr(Data)
Exit Sub
eh:
If AlphaPresent Then
    PushError
    CleanData Data
    PopError
End If
UnReferAry AryPtr(Data)
ErrRaise "dbSave"
End Sub

'to pass extension, use filename:="."+ext
Public Function FileFormatFromExt(ByRef FileName As String) As String
Dim ExtList As String
Dim ext As String
Dim ID As String
Dim i As Long
ext = GetExt(FileName)
If Len(ext) = 0 Then Exit Function
ConnectFormats
For i = 0 To nFormats - 1
    ExtractExtsFromFilter FormatList(i).GetFilter(ftForLoading), ExtList
    If InStr(1, ExtList, ext, vbTextCompare) > 0 Then
        FormatList(i).GetInfo ID, "", False
        FileFormatFromExt = ID
        Exit For
    End If
Next i
End Function

Public Sub ExtractExtsFromFilter(ByRef Filter As String, ByRef ExtList As String)
Dim Ary() As String
Dim AryExts() As String
Dim Accu As String, Delimiter As String
Dim i As Long, j As Long
Ary = Split(Replace(Filter, "*", ""), "|")
For i = 0 To (UBound(Ary) + 1) \ 2 - 1
    Accu = Accu + Delimiter + Replace(Replace(Replace(Ary(2 * i + 1), ";", "|"), " ", ""), ".", "")
    If Len(Delimiter) = 0 Then Delimiter = "|"
Next i
ExtList = Accu
End Sub


Public Sub vtLoadPicture(ByRef Data() As Long, _
                         ByRef Alpha() As Long, _
                         ByRef FileName As String, _
                         Optional ByRef FormatID As String, _
                         Optional ByVal UpdateSettings As Boolean = False, _
                         Optional ByVal ShowDialog As Boolean = False, _
                         Optional ByRef Purpose As String)
    Dim nmb As Long
    Dim i As Long
    On Error GoTo eh
    If Len(FileName) = 0 Then ShowDialog = True
    If ShowDialog Then
        FileName = ShowPictureOpenDialog(FileName, Purpose:=Purpose)
    End If
    nmb = FreeFile
    If FileLen(FileName) = 0 Then
        FormatID = FileFormatFromExt(FileName)
        On Error GoTo 0
        Err.Raise errNewFile, "vtLoadPicture", "The file is a new file!"
    End If
    
    ConnectFormats
    
    Open FileName For Binary Access Read As nmb
        For i = 0 To nFormats - 1
            If FormatList(i).IsFormat(nmb) Then
                Exit For
            End If
        Next i
    Close nmb
    If i >= nFormats Then
        Err.Raise 159, "vtLoadPicture", "Unsupported format."
    End If
    vtLoadPictureUsingFormat FileName, AryPtr(Data), Alpha, FormatList(i), UpdateSettings
    FormatList(i).GetInfo FormatID, "", False
    CheckExt FileName, FormatID
Exit Sub

tryvb:
    i = FormatIndexFromID("VB")
    vtLoadPictureUsingFormat FileName, AryPtr(Data), Alpha, FormatList(i), UpdateSettings
    FormatList(i).GetInfo FormatID, "", False
Exit Sub
Resume
eh:
    Dim Asked As Boolean
    PushError
    If Asked Or Err.Number = errNewFile Or GetFormatID(FormatList(i)) = "VB" Then
        PopError
        ErrRaise "vtLoadPicture"
    Else
        Asked = True
        Select Case MsgError("File loading failed." + vbNewLine + vbNewLine + "Err.Description" + vbNewLine + vbNewLine + "Would you like to try loading it using visual basic's method?", vbYesNo)
            Case vbYes
                PopError
                Resume tryvb
            Case vbNo
                PopError
                ErrRaise "vtLoadPicture"
            Case Else
                PopError
                ErrRaise "vtLoadPicture"
        End Select
    End If
End Sub

Private Function GetFormatID(ByRef Format As FormatTemplate) As String
Dim ID As String
Format.GetInfo ID, "", False
GetFormatID = ID
End Function

Public Sub vtLoadPictureUsingFormat(ByRef File As String, _
                                    ByVal ptrData As Long, _
                                    ByRef Alpha() As Long, _
                                    ByRef Format As FormatTemplate, _
                                    ByVal UpdateSettings As Boolean)
    
    Dim Data() As RGBQUAD
    Dim i As Long
    Dim w As Long, h As Long
    Dim AlphaPresent As Boolean
    Dim vData() As RGBQUAD
    Dim VAlpha() As Long
    AlphaPresent = Format.LoadFile(File, UpdateSettings)
    Format.ExtractData ptrData
    
    If AlphaPresent Then
        AryWH ptrData, w, h
        If w = 0 Or h = 0 Then
            Erase Alpha
            Exit Sub
        End If
        
        ReDim Alpha(0 To w - 1, 0 To h - 1)
        
        ReferAry AryPtr(Data), ptrData
        ConstructAry AryPtr(vData), VarPtr(Data(0, 0)), 4, w * h
        ConstructAry AryPtr(VAlpha), VarPtr(Alpha(0, 0)), 4, w * h
        On Error GoTo eh
            For i = 0 To w * h - 1
                VAlpha(i) = vData(i).rgbReserved * &H10101
                vData(i).rgbReserved = 0
            Next i
            
        UnReferAry AryPtr(VAlpha)
        UnReferAry AryPtr(vData)
        UnReferAry AryPtr(Data)
    Else
        Erase Alpha
    End If
Exit Sub

eh:
    UnReferAry AryPtr(VAlpha)
    UnReferAry AryPtr(vData)
    UnReferAry AryPtr(Data)
ErrRaise "dbLoad"

End Sub


Public Function GetLoadFilter() As String
Dim i As Long
Dim AllList As String, AllFilter As String
Dim ExtLists() As String
Dim iEL As Long
Dim ID As String
Dim CanLoad As Boolean
Dim Filters() As String

ConnectFormats

ReDim ExtLists(0 To nFormats - 1)
ReDim Filters(0 To nFormats - 1)
iEL = 0

For i = 0 To nFormats - 1
    With FormatList(i)
        .GetInfo ID, "", False, CanLoad:=CanLoad
        If CanLoad Then
            Filters(iEL) = .GetFilter(FilterType:=ftForLoading)
            ExtractExtsFromFilter Filters(iEL), ExtLists(iEL)
            iEL = iEL + 1
        End If
    End With
Next i

If iEL > 0 Then
    ReDim Preserve ExtLists(0 To iEL - 1)
    ReDim Preserve Filters(0 To iEL - 1)
Else
    Err.Raise 1212, "GetLoadFilter", "No load compatible formats!"
End If
AllList = Join(ExtLists, "|")

AllFilter = grs(1001, "$al$", Replace(AllList, "|", " "), _
                      "$alf$", "*." + Replace(AllList, "|", ";*."))
GetLoadFilter = AllFilter + "|" + Join(Filters, "|")
End Function

Public Function ShowPictureOpenDialog(ByRef InitFileName As String, _
                                      Optional ByVal OnlyForDir As Boolean = True, _
                                      Optional ByRef Purpose As String) As String
With CDl
    .CancelError = True
    .DialogTitle = "Open picture"
    .Filter = GetLoadFilter
    .OpenFlags = cdlOFNFileMustExist Or cdlOFNHideReadOnly
    If Len(InitFileName) > 0 Then
        .InitDir = GetDirName(InitFileName)
    Else
        .FileName = ""
        .InitDir = GetSMBCurDir(Purpose:=Purpose)
    End If
    If Not OnlyForDir Then
        .FileName = InitFileName
    End If
    .hWndOwner = Screen.ActiveForm.hWnd
    
    .ShowOpen
    
    ShowPictureOpenDialog = .FileName
    SaveSMBCurDir .FileName, Purpose:=Purpose
End With

End Function

Private Sub CheckExt(ByRef FileName As String, ByRef FormatID As String)
Dim ExtFID As String
Dim FormatDesc As String
Dim GoodExt As String
Dim Filter As String
Dim Pos As Long
ExtFID = FileFormatFromExt(FileName)
If FormatID <> ExtFID Then
    If FormatID <> "" And GetExt(FileName) <> "" Then
        
        With FormatFromID(FormatID)
            .GetInfo "", FormatDesc, False
            ExtractExtsFromFilter .GetFilter(ftForLoading), GoodExt
            If Len(GoodExt) = 0 Then
                GoodExt = GRSF(2618) '<extension failed>
            Else
                Pos = InStr(GoodExt, "|")
                If Pos > 0 Then GoodExt = Mid$(GoodExt, 1, Pos - 1)
            End If
        End With
        
        dbMsgBox grs(2451, "$fn", FileName, _
                           "$ext", GetExt(FileName), _
                           "$typ", FormatDesc, _
                           "$goodext", GoodExt), vbExclamation
    
    ElseIf FormatID = "" And ExtFID <> "" Then
        
        With FormatFromID(ExtFID)
            .GetInfo "", FormatDesc, False
        End With
        dbMsgBox grs(2452, "$fn", FileName, _
                           "$ext", GetExt(FileName), _
                           "$typ", FormatDesc), vbExclamation
    
    ElseIf FormatID <> "" And GetExt(FileName) = "" Then
    
        With FormatFromID(FormatID)
            .GetInfo "", FormatDesc, False
            ExtractExtsFromFilter .GetFilter(ftForLoading), GoodExt
            If Len(GoodExt) = 0 Then
                GoodExt = GRSF(2618) '<extension failed>
            Else
                Pos = InStr(GoodExt, "|")
                If Pos > 0 Then GoodExt = Mid$(GoodExt, 1, Pos - 1)
            End If
        End With
        dbMsgBox grs(2453, "$fn", FileName, _
                           "$typ", FormatDesc, _
                           "$goodext", GoodExt), vbExclamation
    
    End If
End If
End Sub
