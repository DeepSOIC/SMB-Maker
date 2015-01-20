Attribute VB_Name = "mdlJPEG"
Option Explicit

Private Const DLLName As String = "JpegSave.dll"
Private Declare Sub psSaveJPEG Lib "JpegSave.dll" Alias "SaveToJpeg" _
            (ByVal hBitmap As Long, _
            ByVal FileName As String, _
            ByRef jpInfo As vtJPEGInfo)

Public Type vtJPEGInfo
    Quality As Long
    Progressive As Boolean
End Type

Public Function IsJPEG(ByVal FileNumber As Long) As Boolean
Const FileID As String = "ÿØÿ"
Dim Bytes() As Byte
Dim i As Long
ReDim Bytes(0 To Len(FileID) - 1)

Get FileNumber, 1, Bytes

IsJPEG = True
For i = 0 To Len(FileID) - 1
    If Bytes(i) <> Asc(Mid$(FileID, i + 1, 1)) Then
        IsJPEG = False
        Exit For
    End If
Next i
End Function

Public Sub SaveToJPEG(ByRef Data() As Long, _
                      ByRef jpInfo As vtJPEGInfo, _
                      ByRef FileName As String)
Dim hBmp As Long

CheckDll DLLName


hBmp = MakeBitmap(Data)
On Error GoTo eh
StartWrite FileName
psSaveJPEG hBmp, FileName, jpInfo
DeleteObject hBmp

Exit Sub
eh:
DeleteObject hBmp
ErrRaise "SaveToJPEG"
End Sub

Public Function MakeBitmap(ByRef Data() As Long, _
                           Optional ByRef ptrData As Long) As Long
Dim bmi As BITMAPINFO
Dim hBmp As Long
Dim w As Long, h As Long
Dim hDC As Long
'hDC = MainForm.hDC
AryWH AryPtr(Data), w, h

With bmi.bmiHeader
    .biSize = Len(bmi.bmiHeader)
    .biBitCount = 32
    .biPlanes = 1
    .biWidth = w
    .biHeight = -h
    .biSizeImage = w * h * 4&
End With
hBmp = CreateDIBSection(hDC, bmi, DIB_RGB_COLORS, VarPtr(ptrData), 0&, 0&)
If hBmp = 0 Then Err.Raise 1212, "mdlJPEG:MakeBitmap", "Failed to create DIB section."
On Error GoTo eh
If w > 0 Then
    CopyMemory ByVal ptrData, Data(0, 0), bmi.bmiHeader.biSizeImage
End If
MakeBitmap = hBmp
Exit Function
eh:
DeleteObject hBmp
ErrRaise "MakeBitmap"
End Function
