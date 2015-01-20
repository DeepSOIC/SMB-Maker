VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFormatBmp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'common declarations
Implements FormatTemplate
Dim pPtrData As Long
Dim LoadedData() As Long
Dim SettingsLoaded As Boolean

'-----------------------------------
'format specific declarations
Dim frmSettings As New frmFormatBMP
Dim WriteRasterSize As Boolean
Dim DpiX As Long, DpiY As Long
Dim TopDownDIB As Boolean

Private Type BITMAPFILEHEADER
    bfType As Integer
    bfSize As Long
    bfReserved1 As Integer
    bfReserved2 As Integer
    bfOffBits As Long
End Type


Private Sub FormatTemplate_GetInfo(ID As String, ListName As String, HasSettings As Boolean, Optional CanSave As Boolean, Optional CanLoad As Boolean)
ID = "BMP"
ListName = "Windows Bitmap (BMP)"
HasSettings = True
CanSave = True
CanLoad = True
End Sub

Private Sub FormatTemplate_GetFileTypeInfo(Optional ByRef ExtList As String, _
                                           Optional ByRef TypeDescription As String, _
                                           Optional ByRef IconString As String, _
                                           Optional ByRef DefEditor As Boolean)
ExtList = "bmp|dib"
TypeDescription = GRSF(2605)
DefEditor = False
IconString = ""
End Sub

Private Function FormatTemplate_GetFilter(Optional ByVal FilterType As eFilterType = 0&) As String
If FilterType = ftForLoading Then
    FormatTemplate_GetFilter = "Windows bitmap (.bmp .dib)|*.bmp;*.dib"
ElseIf FilterType = ftForSaving Then
    FormatTemplate_GetFilter = "Windows bitmap (.bmp .dib)|*.bmp;*.dib"
End If
End Function

Private Function FormatTemplate_CanSave(ByVal AlphaPresent As Boolean, Reason As eBadSettings) As Boolean
FormatTemplate_CanSave = True
End Function

Private Function FormatTemplate_AlphaSupproted() As Boolean
FormatTemplate_AlphaSupproted = FormatTemplate_CanSave(False, 0) Or FormatTemplate_CanSave(True, 0)
End Function

Private Function FormatTemplate_GetDescriptionResID() As Long
FormatTemplate_GetDescriptionResID = 2611
End Function

Private Sub FormatTemplate_Customize()
SettingsToForm
On Error GoTo eh
With frmSettings
    .Show vbModal
    If Len(.Tag) > 0 Then
        Err.Raise dbCWS
    End If
End With
FormToSettings
Exit Sub
eh:
PushError
Unload frmSettings
PopError
ErrRaise
End Sub

Private Sub FormatTemplate_ExtractData(ByVal ptrData As Long)
SwapArys ptrData, AryPtr(LoadedData)
Erase LoadedData
End Sub

Private Function FormatTemplate_IsFormat(ByVal FileNumber As Long) As Boolean
FormatTemplate_IsFormat = IsFormat(FileNumber)
End Function

Private Function FormatTemplate_LoadFile(FileName As String, ByVal UpdateSettings As Boolean) As Boolean
FormatTemplate_LoadFile = LoadFile(FileName, UpdateSettings)
End Function

Private Sub FormatTemplate_RemovePtrData()
pPtrData = 0
End Sub

Private Sub FormatTemplate_SaveFile(FileName As String, ByVal HasAlpha As Boolean)
SaveFile FileName, HasAlpha
End Sub

Private Sub FormatTemplate_LoadSettings()

    If SettingsLoaded Then Exit Sub
    
    On Error Resume Next
    
    WriteRasterSize = dbGetSettingEx("Formats\BMP", "WriteRasterSize", vbBoolean, True)
    DpiX = dbGetSettingEx("Formats\BMP", "PixelsPerMeterX", vbLong, 0&)
    DpiY = dbGetSettingEx("Formats\BMP", "PixelsPerMeterY", vbLong, 0&)
    TopDownDIB = dbGetSettingEx("Formats\BMP", "TopDownDIB", vbBoolean, False)
 
    SettingsLoaded = True

End Sub

Private Sub FormatTemplate_SaveSettings()

    If Not SettingsLoaded Then Exit Sub
    
    dbSaveSettingEx "Formats\BMP", "WriteRasterSize", WriteRasterSize
    dbSaveSettingEx "Formats\BMP", "PixelsPerMeterX", DpiX
    dbSaveSettingEx "Formats\BMP", "PixelsPerMeterY", DpiY
    dbSaveSettingEx "Formats\BMP", "TopDownDIB", TopDownDIB
    

End Sub

Private Sub FormatTemplate_SetPtrData(ByVal ptrData As Long)
pPtrData = ptrData
End Sub





'-------------------------------------------------------------------
Private Sub SettingsToForm()
FormatTemplate_LoadSettings

Load frmSettings
With frmSettings

    .WriteRS(0).Value = Not WriteRasterSize
    .WriteRS(1).Value = WriteRasterSize
    
    .DpiX = IIf(DpiX >= 0, CStr(DpiX), "&H" + Hex$(DpiX))
    .DpiY = IIf(DpiY >= 0, CStr(DpiY), "&H" + Hex$(DpiY))
    
    .optBottomUp.Value = Not TopDownDIB
    .optTopDown.Value = TopDownDIB

End With
End Sub

Private Sub FormToSettings()

    With frmSettings
        
        WriteRasterSize = .WriteRS(1).Value
        
        DpiX = CLng(.DpiX)
        DpiY = CLng(.DpiY)
        
        TopDownDIB = .optTopDown.Value
        
    End With
    
    Unload frmSettings
    
End Sub
'/dialogs

'-------------------------------------------------------------------
'File reading

Private Function IsFormat(ByVal FileNumber As Long)
Const FILE_ID = &H4D42 'BM
Dim nInt As Integer
Get FileNumber, 1, nInt
IsFormat = nInt = FILE_ID
End Function

Private Function LoadFile(ByRef FileName As String, _
                          ByVal SetSettings As Boolean) As Boolean
    
    Dim nmb As Long
    
    Dim BFH As BITMAPFILEHEADER
    Dim bmi As BITMAPINFOHEADER
    Dim FL As Long
    Dim OfcData As Long
    Dim pal() As Long
    Dim PalRGB() As RGBQUAD
    Dim nColors As Long
    Dim w As Long, h As Long
    Dim TopDown As Boolean
    Dim lScanLine8 As Long
    Dim nColorsPresent As Long
    Dim RawLine() As Byte
    Dim i As Long
    Dim x As Long, y As Long
    Dim y1 As Long
    Dim BPP As Long
    Dim RGBData() As RGBQUAD
    
    Dim ErrorIgnorable As Boolean
    nmb = FreeFile
    Open FileName For Binary Access Read As nmb
    On Error GoTo eh
    
        Get nmb, 1, BFH
        
'        If BFH.bfType <> Asc("B") + 256& * Asc("M") Then
'            Err.Raise 481, , "Invalid bitmap file" 'invalid picture
'        End If
        OfcData = BFH.bfOffBits
        
        Get nmb, , bmi
        
        If bmi.biSize <> Len(bmi) Then
            ErrorIgnorable = True
            Err.Raise 111, , "Incorrect structure size. It should be " + CStr(Len(bmi)) + ", but is " + CStr(bmi.biSize) + "."
            ErrorIgnorable = False
        End If
        
        If bmi.biCompression <> 0 Then
            ErrorIgnorable = True
            Err.Raise 111, , "Compressed bitmaps are not supported. Compression: " + CStr(bmi.biCompression) + "."
            ErrorIgnorable = False
        End If
        
        BPP = bmi.biBitCount
        
        Select Case BPP
            Case 1, 4, 8
                nColors = Stepen2(BPP)
            Case 24, 32
                nColors = 0&
            Case Else
                Err.Raise 111, , "Bit count not supported: " + CStr(BPP) + "."
        End Select
        
        w = bmi.biWidth
        If w <= 0 Then
            Err.Raise 111, , "Width must be greater than zero. It is: " + CStr(w) + "."
        End If
        lScanLine8 = -Int(-w * BPP / 8 / 4) * 4
        
        ReDim RawLine(0 To lScanLine8 - 1)
        
        TopDown = h < 0
        h = Abs(bmi.biHeight)
        
        If h = 0 Then
            Err.Raise 111, , "Height must be nonzero but it is zero."
        End If
        
        If bmi.biSizeImage <> 0 And bmi.biSizeImage <> lScanLine8 * h Then
            ErrorIgnorable = True
            Err.Raise 1212, , "The size field of this bitmap contains invalid value. Resave the file to fix the problem."
            ErrorIgnorable = False
        End If
        
        If SetSettings Then
            DpiX = bmi.biXPelsPerMeter
            DpiY = bmi.biYPelsPerMeter
            WriteRasterSize = bmi.biSizeImage
            TopDownDIB = TopDown
        End If
        
        ReDim LoadedData(0 To w - 1, 0 To h - 1)
    
        
        Select Case bmi.biBitCount
            Case 1, 4, 8
                
                ReDim pal(0 To nColors - 1)
                LoadDefPal pal
                
                nColorsPresent = Min(OfcData - Len(BFH) - Len(bmi), 4& * nColors) \ 4
                
                If nColorsPresent > 0 Then
                    ReDim PalRGB(0 To nColorsPresent - 1)
                    Get nmb, , PalRGB
                    For i = 0 To nColorsPresent - 1
                        pal(i) = RGB(PalRGB(i).rgbBlue, PalRGB(i).rgbGreen, PalRGB(i).rgbRed)
                    Next i
                End If
                
                For y = 0 To h - 1
                    If TopDown Then y1 = y Else y1 = h - 1 - y
                    Get nmb, 1 + OfcData + lScanLine8 * y, RawLine
                    For x = 0 To w - 1
                        LoadedData(x, y1) = pal(GetBits(RawLine(x * BPP \ 8&), x * BPP Mod 8&, BPP))
                    Next x
                Next y
                
            Case 24
                ReferAry AryPtr(RGBData), AryPtr(LoadedData)
                For y = 0 To h - 1
                    If TopDown Then y1 = y Else y1 = h - 1 - y
                    Get nmb, 1 + OfcData + lScanLine8 * y, RawLine
                    For x = 0 To w - 1
                        RGBData(x, y1).rgbBlue = RawLine(3& * x)
                        RGBData(x, y1).rgbGreen = RawLine(3& * x + 1&)
                        RGBData(x, y1).rgbRed = RawLine(3& * x + 2&)
                    Next x
                Next y
                UnReferAry AryPtr(RGBData)
                
            Case 32
                For y = 0 To h - 1
                    If TopDown Then y1 = y Else y1 = h - 1 - y
                    Get nmb, 1 + OfcData + lScanLine8 * y, RawLine
                    CopyMemory LoadedData(0, y1), RawLine(0), lScanLine8
                    'For x = 0 To w - 1
                    '    RGBData(x, y1).rgbBlue = RawLine(3& * x)
                    '    RGBData(x, y1).rgbGreen = RawLine(3& * x + 1&)
                    '    RGBData(x, y1).rgbRed = RawLine(3& * x + 2&)
                    'Next x
                Next y
                LoadFile = True 'alpha channel present
            
        End Select
    
    Close nmb

Exit Function
eh:
Dim Answ As VbMsgBoxResult
PushError
    If ErrorIgnorable Then
        Answ = MsgError(Message:="Bitmap loading error." + vbNewLine + "Err.Description" + vbNewLine + "You can ignore this error.", Style:=vbAbortRetryIgnore)
        Select Case Answ
            Case vbRetry
                PopError
                Resume
            Case vbAbort
                'continue
            Case vbIgnore
                PopError
                Resume Next
        End Select
    End If
    
    UnReferAry AryPtr(RGBData)
    Close nmb
PopError
ErrRaise "BMP:LoadFile"
End Function

'gets number represented by bits from FirstBit (0 is most significant)
Public Function GetBits(ByVal b As Byte, _
                        ByVal FirstBit As Long, _
                        ByVal nBits As Long) As Integer
GetBits = (CInt(b) \ Stepen2(8& - (FirstBit + nBits))) _
          And _
          (Stepen2(nBits) - 1)
End Function

Public Sub SaveFile(ByRef FileName As String, _
                    ByVal WriteAlphaChannel As Boolean)
Dim RGBData() As RGBQUAD
If pPtrData = 0 Then
    Err.Raise 1212, , "Nothing set to data. Cannot save. (internal error)"
End If
Const LinesInPack = 16 'the number of lines to collect before write
Dim CurLinesInPack As Long

Dim bmiH As BITMAPINFOHEADER
Dim BFH As BITMAPFILEHEADER

Dim x As Long, y As Long
Dim YRead As Long

Dim OutBits() As Byte 'bits pack to write
Dim ty As Long 'write y position in OutBits
Dim OfcYWrite As Long
Dim xx3 As Long

Dim LineLen As Long 'Line length in bytes
Dim w As Long, h As Long
Dim TopDown As Boolean

Dim nmb As Long

On Error GoTo eh
ReferAry AryPtr(RGBData), pPtrData

If AryDims(AryPtr(RGBData)) <> 2 Then
    Err.Raise 1111, "SaveBitmap", "Incorrect number of dimensions in passed data!"
End If

AryWH AryPtr(RGBData), w, h
If WriteAlphaChannel Then
    LineLen = w * 4&
Else
    LineLen = PadDWORD(w * 3)
End If

'prepare headers
With bmiH
    If WriteAlphaChannel Then
        .biBitCount = 32
    Else
        .biBitCount = 24
    End If
    .biWidth = w
    If TopDown Then
        .biHeight = -h
    Else
        .biHeight = h
    End If
    .biSize = Len(bmiH)
    .biSizeImage = LineLen * h
    .biXPelsPerMeter = DpiX
    .biYPelsPerMeter = DpiY
    .biPlanes = 1
End With

With BFH
    .bfType = &H4D42 'BM
    .bfOffBits = Len(BFH) + Len(bmiH)
    .bfSize = Len(BFH) + Len(bmiH) + bmiH.biSizeImage
End With

If Not WriteRasterSize Then
    bmiH.biSizeImage = 0&
End If

'open the file
StartWrite FileName
nmb = FreeFile
Open FileName For Binary Access Write As nmb
    'write headers
    Put nmb, 1, BFH
    Put nmb, , bmiH
    'Map arrays, prepare
    CurLinesInPack = Min(LinesInPack, h)
    ReDim OutBits(0 To CurLinesInPack * LineLen - 1)
    'write the data
    For y = 0 To h - 1
        'write the line
        If TopDown Then
            YRead = y
        Else
            YRead = (h - 1 - y)  'Upside-down image
        End If
        OfcYWrite = ty * LineLen
        xx3 = OfcYWrite
        If WriteAlphaChannel Then
            CopyMemory OutBits(xx3), RGBData(0, YRead), LineLen
        Else
            For x = 0& To w - 1&
                OutBits(xx3) = RGBData(x, YRead).rgbBlue
                OutBits(xx3 + 1&) = RGBData(x, YRead).rgbGreen
                OutBits(xx3 + 2&) = RGBData(x, YRead).rgbRed
                xx3 = xx3 + 3&
            Next x
        End If
        ty = ty + 1&
        'test for
        If ty = CurLinesInPack Then
            'currently, y+1 lines were processed
            Put nmb, , OutBits
            If y + 1& <> h Then 'if not all the lines processed
                CurLinesInPack = Min(h - (y + 1&), LinesInPack)
                If CurLinesInPack < LinesInPack Then
                    ReDim Preserve OutBits(0 To CurLinesInPack * LineLen - 1)
                End If
                ty = 0&
            End If
        End If
    Next y
    Debug.Assert ty = CurLinesInPack
Close nmb
UnReferAry AryPtr(RGBData)
Exit Sub
eh:
    UnReferAry AryPtr(RGBData)
    Close nmb
    ErrRaise "BMP:SaveFile"

End Sub

Function PadDWORD(ByVal Length As Long) As Long
PadDWORD = -Int(-Length / 4#) * 4#
End Function


