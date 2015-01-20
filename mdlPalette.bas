Attribute VB_Name = "mdlPalettes"
Option Explicit
Private Const RIFF_FILE_SIGNATURE = &H46464952 'RIFF
Private Const RIFF_FILE_TYPE = &H204C4150 '"PAL "
Private Const RIFF_CHUNK_SIGNATURE = &H61746164 '"data"

Type vtColorTip
    nColor As Integer 'the palette index of the color with the tip
    strTip As String
End Type

Public Type vtPaletteWTips
    Colors() As Long
    Tips() As vtColorTip 'the list of tips with their numbers
End Type


Type RIFFFile
'   Name                    Offset
    Signature_RIFF As Long  '0
    FileLen_8 As Long       '4
    RIFFType As Long        '8
    'chunk
    Signature_Chunk As Long '12
    ChunkSize As Long       '16
    PalVer As Integer       '20
    NOfColors As Integer    '22
                            '24
End Type


Type Pal_Entry
    Left As Long
    Top As Long
    Width As Long
    Height As Long
    BackColor As Long
    Tip As String
End Type

Sub LoadDefPal(ByRef pal() As Long)
Dim Bytes() As Byte, i As Long
    If UBound(pal) = 1 Then
        pal(0) = 0
        pal(1) = RGB(255, 255, 255)
    Else
        Bytes = LoadResData("D" + CStr(UBound(pal) + 1), "PALETTE")
        For i = 0 To UBound(pal)
            pal(i) = RGB(Bytes(24 + i * 4), Bytes(24 + i * 4 + 1), Bytes(24 + i * 4 + 2))
        Next i
    End If
End Sub


Public Sub SavePaletteEx(ByRef Palette As vtPaletteWTips, ByVal File As String)
Dim nmb As Long
Dim PalInfo As RIFFFile
Dim nTips As Integer
Dim nColors As Long
Dim TipsPresent As Boolean
Dim tmp() As Long
Dim i As Long
    
    On Error GoTo eh
    
    StartWrite File
    If AryDims(AryPtr(Palette.Tips)) <> 1 Then
        TipsPresent = False
    Else
        TipsPresent = True
        nTips = AryLen(AryPtr(Palette.Tips))
    End If
    
    nColors = AryLen(AryPtr(Palette.Colors))
    If nColors = 0 Then Err.Raise 1212, , "Empty palette cannot be saved."
    ReDim tmp(0 To nColors - 1)
    For i = 0 To nColors - 1
        tmp(i) = ConvertColorLng(Palette.Colors(i))
    Next i
    
    With PalInfo
        .Signature_RIFF = RIFF_FILE_SIGNATURE
        .FileLen_8 = 16 + (nColors) * 4
        .RIFFType = RIFF_FILE_TYPE
        .Signature_Chunk = RIFF_CHUNK_SIGNATURE
        .ChunkSize = 4 * (nColors) + 4
        .PalVer = &H300
        .NOfColors = nColors
    End With
    
    nmb = FreeFile
    Open File For Binary Access Write As nmb
        Put nmb, 1, PalInfo
        Put nmb, 25, tmp
        If TipsPresent Then
            Put nmb, , nTips
            Put nmb, , Palette.Tips
        End If
    Close nmb
    
Exit Sub
eh:
    PushError
    Close nmb
    PopError
ErrRaise "SavePaletteEx"
End Sub

Public Sub SavePalette(ByRef Colors() As Long, _
                       ByVal File As String, _
                       Optional ByVal ptrAryTips As Long)
Dim Tips() As vtColorTip
Dim pal As vtPaletteWTips
On Error GoTo eh
ReferAry AryPtr(pal.Colors), AryPtr(Colors)
    If ptrAryTips <> 0 Then
        SwapArys AryPtr(pal.Tips), ptrAryTips
    End If
    SavePaletteEx pal, File
    If ptrAryTips <> 0 Then
        SwapArys AryPtr(pal.Tips), ptrAryTips
    End If
UnReferAry AryPtr(pal.Colors)

Exit Sub
eh:
    UnReferAry AryPtr(pal.Colors)
    If ptrAryTips <> 0 Then
        SwapArys AryPtr(pal.Tips), ptrAryTips
    End If
ErrRaise "SavePalette"
End Sub

Public Sub LoadPaletteEx(ByVal File As String, Palette As vtPaletteWTips)
Dim i As Long
Dim nmb As Long
Dim r As Byte, g As Byte, b As Byte
Dim PalInfo As RIFFFile
Dim nOfTips As Integer

Dim nColors As Long

nmb = FreeFile
Open File For Binary Access Read As nmb
    Get nmb, 1, PalInfo
    nColors = PalInfo.NOfColors
    If nColors = 0 Or _
       PalInfo.ChunkSize <> nColors * 4& + 4& Or _
       PalInfo.PalVer <> &H300 Or _
       PalInfo.RIFFType <> RIFF_FILE_TYPE Or _
       PalInfo.Signature_Chunk <> RIFF_CHUNK_SIGNATURE Or _
       PalInfo.Signature_RIFF <> RIFF_FILE_SIGNATURE Then
        Err.Raise 1001, , GRSF(1165) ' The palette file is bad. It might be corrupted.
    End If
    If nColors <= 0 Then Err.Raise 1212, , "Palettes with no colors are not supported in SMB Maker!"
    
    ReDim Palette.Colors(0 To nColors - 1)
    Get nmb, 25, Palette.Colors
    
    For i = 0 To nColors - 1
        Palette.Colors(i) = ConvertColorLng(Palette.Colors(i))
    Next i
    
    Get nmb, , nOfTips
    If nOfTips > 0 Then
        ReDim Tips(0 To nOfTips - 1)
        Get nmb, , Tips
    Else
        Erase Tips
    End If
Close nmb
End Sub

Public Sub LoadPalette(ByVal File As String, _
                       ByRef Palette() As Long, _
                       Optional ByVal ptrTips As Long)
Dim pal As vtPaletteWTips
LoadPaletteEx File, pal
SwapArys AryPtr(Palette), AryPtr(pal.Colors)
If ptrTips <> 0 Then
    SwapArys ptrTips, AryPtr(pal.Tips)
End If
End Sub

Public Sub ExtractResPal(ByRef Palette() As Long, ByVal ResID As String)
Dim Bytes() As Byte, nCols As Long, i As Long
    'On Error Resume Next
    On Error GoTo 0
    Bytes = LoadResData(ResID, "PALETTE")
    nCols = Bytes(22) + CLng(Bytes(23)) * 256&
    ReDim Palette(0 To nCols - 1)
    For i = 0 To nCols - 1
        Palette(i) = Bytes(24 + i * 4) + Bytes(24 + i * 4 + 1&) * &H100& + Bytes(24 + i * 4 + 2&) * &H10000
    Next i
End Sub

Public Sub Pal2Image(ByRef Image() As Long, ByRef Palette() As Long)
Dim nColors As Long
nColors = AryLen(AryPtr(Palette))
If nColors = 0 Then Err.Raise 1212, , "Empty palettes are not supported!"
ReDim Image(0 To nColors - 1, 0 To 0)
CopyMemory Image(0, 0), Palette(0), nColors * 4
End Sub

Public Sub Image2Pal(ByRef Image() As Long, ByRef Palette() As Long)
Dim nColors As Long
Dim w As Long, h As Long
AryWH AryPtr(Image), w, h
If w * h = 0 Then Err.Raise 1212, , "Empty palettes are not supported!"
nColors = w * h
If nColors > &H7FFF& Then
    Err.Raise 1212, , "The image is too large to be converted into a palette."
End If
ReDim Palette(0 To nColors - 1)
CopyMemory Palette(0), Image(0, 0), nColors * 4
End Sub

Public Function IsPal(ByVal FileNumber As Long) As Boolean
Dim PalInfo As RIFFFile
Get FileNumber, 1, PalInfo
IsPal = PalInfo.NOfColors <> 0 And _
    PalInfo.ChunkSize = PalInfo.NOfColors * 4& + 4& And _
    PalInfo.PalVer = &H300 And _
    PalInfo.RIFFType = RIFF_FILE_TYPE And _
    PalInfo.Signature_Chunk = RIFF_CHUNK_SIGNATURE And _
    PalInfo.Signature_RIFF = RIFF_FILE_SIGNATURE

End Function
