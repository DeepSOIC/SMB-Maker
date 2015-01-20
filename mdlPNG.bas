Attribute VB_Name = "mdlPNG"
Option Explicit
'contains code for writing PNG's
'Used file: mpng.dll
Private Declare Function mLoadPNG _
      Lib "mpng.dll" (ByVal strFileName As String, _
                      hPNG As Long, _
                      bminfo As BITMAPINFO, _
                      Length As Long) As Long
'retrieves information about PNG file.
'Returns a handle to opened png.
'Length is the length of decoded picture bits.
'BPP can be found in returned BITMAPINFO
'Must return 1 on success

Private Declare Function mEndPNG _
      Lib "mpng.dll" (hPNG As Long) As Long
'Close the png and free all the memory
      
Private Declare Function mGetPNGData _
      Lib "mpng.dll" (hPNG As Long, _
                      buf As Any) As Long
'read png data into buf.
'must return 1 on success

Private Declare Function mWritePNG _
      Lib "mpng.dll" (ByVal strFileName As String, _
                      lpdat As Any, _
                      bminfo As BITMAPINFO, _
                      ByVal Interlace As Long) As Long
'Writes a png file. Interlace can be 0 and 1.
'Bit count and the palette are specified in bminfo.
'returns nonzero if succeed amd zero if failes.

'types redefined for encapsulation
Private Type BITMAPINFOHEADER
    biSize As Long
    biWidth As Long
    biHeight As Long
    biPlanes As Integer
    biBitCount As Integer
    biCompression As Long
    biSizeImage As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed As Long
    biClrImportant As Long
End Type
Private Type RGBQUAD
    rgbBlue As Byte
    rgbGreen As Byte
    rgbRed As Byte
    rgbReserved As Byte
End Type
Private Type BITMAPINFO
    bmiHeader As BITMAPINFOHEADER
    bmiColors(0 To 255) As RGBQUAD
End Type

'PicData. Long colors in BGRA format.

'Doesn't check the file name thinking that it is ok.
'(non-existing or not write-protected). You should
'check it before calling this function.
'(The file name should be ok, the file should be
'creatable/writable)
Public Sub SavePNG32(ByRef PicData() As Long, _
                     ByRef File As String, _
                     Optional ByVal Interlace As Boolean = False)
Dim bmi As BITMAPINFO
Dim w As Long, h As Long
Dim Bits() As RGBQUAD
Dim Ret As Long
CheckPNGDll
w = UBound(PicData, 1) + 1
h = UBound(PicData, 2) + 1
With bmi.bmiHeader
    .biBitCount = 32
    .biHeight = h
    .biWidth = w
    .biSize = Len(bmi.bmiHeader)
    .biSizeImage = w * h * 4 'bytes
End With
Erase bmi.bmiColors 'fill bmicolors with zeros to be sure

StartWrite File
Ret = mWritePNG(File, PicData(0, 0), bmi, IIf(Interlace, 1, 0))


If Ret = 0 Then
    Err.Raise 2121, "SavePNG32", "Cannot write a PNG. The function returned zero."
End If
End Sub

'PicData. Long colors in BGR0 format.

'Assumes that Alpha channel size matches the PicData size!

'Doesn't check the file name thinking that it is ok.
'(non-existing or not write-protected). You should
'check it before calling this function.
'(The file name should be ok, the file should be
'creatable/writable)
Public Sub SavePNG24(ByRef PicData() As Long, _
                     ByRef File As String, _
                     Optional ByVal Interlace As Boolean = False)
Dim bmi As BITMAPINFO
Dim w As Long, h As Long
Dim Bits() As Byte, Bits1() As Byte
Dim AlphaRGB() As RGBQUAD
Dim Ret As Long
CheckPNGDll
w = UBound(PicData, 1) + 1
h = UBound(PicData, 2) + 1
With bmi.bmiHeader
    .biBitCount = 24
    .biHeight = h
    .biWidth = w
    .biSize = Len(bmi.bmiHeader)
    '.biSizeImage = w * h * 3& not correct because of padding
End With
Erase bmi.bmiColors 'fill bmicolors with zeros to be sure

'ReDim Bits(0 To w * h * 4& - 1)

'CopyMemory Bits(0), PicData(0, 0), w * h * 4&

bmi.bmiHeader.biSizeImage = BitsPreProcessing24(PicData, w, h, Bits1)
'Erase Bits
'ChDir App.Path

Ret = mWritePNG(File, Bits1(0), bmi, IIf(Interlace, 1, 0))

Erase Bits
If Ret = 0 Then
    Err.Raise 2121, "SavePNG24", "Cannot write a PNG. The function returned zero."
End If
End Sub
'
''The palette must contain 256 colors
'Public Sub SavePNG8(ByRef PicData() As Byte, _
'                    ByRef Pal() As Long, _
'                    ByRef File As String, _
'                    Optional ByVal Interlace As Boolean = False)
'Dim BMI As BITMAPINFO
'Dim w As Long, h As Long
'Dim Bits() As Byte
'Dim Ret As Long
'CheckPNGDll
'w = UBound(PicData, 1) + 1
'h = UBound(PicData, 2) + 1
'With BMI.bmiHeader
'    .biBitCount = 8
'    .biHeight = h
'    .biWidth = w
'    .biSize = Len(BMI.bmiHeader)
'    .biSizeImage = PadDWORD(w * 1&) * h
'    .biClrUsed = 255
'    .biClrImportant = 255
'End With
'CopyMemory BMI.bmiColors(0), Pal(0), 255
'
'BitsPreProcessing8 Bits, PicData
'
'Ret = mWritePNG(File, Bits(0), BMI, IIf(Interlace, 1, 0))
'
'Erase Bits
'If Ret = 0 Then
'    Err.Raise 2121, "SavePNG8", "Cannot write a PNG. The function returned zero."
'End If
'End Sub



'Adds alpha if neccessary
'Alpha is not changed.
'Swaps r/b in Bits
Private Sub BitsPreProcessing32(ByRef aBits() As Long, _
                                ByRef aAlpha() As Long, _
                                ByVal AlphaSeparate As Boolean, _
                                ByRef Output() As RGBQUAD)
Dim X As Long, Y As Long, y1 As Long
Dim UBX As Long, UBY As Long
'Dim tmp As RGBQUAD
Dim Inv3 As Double
Dim tmpb As Byte
Dim Bits() As RGBQUAD
Dim Alpha() As RGBQUAD
Dim nPix As Long
On Error GoTo ExitHere
Err.Clear
Inv3 = 1# / 3#
UBX = UBound(aBits, 1)
UBY = UBound(aBits, 2)
nPix = (UBX + 1) * (UBY + 1)
ReDim Output(0 To nPix - 1)
ConstructAry AryPtr(Bits), VarPtr(aBits(0, 0)), 4, nPix
If AlphaSeparate Then
    ConstructAry AryPtr(Alpha), VarPtr(aAlpha(0, 0)), 4, nPix
End If

If AlphaSeparate Then
    For X = 0 To nPix - 1
        Output(X).rgbRed = Bits(X).rgbRed
        Output(X).rgbGreen = Bits(X).rgbGreen
        Output(X).rgbBlue = Bits(X).rgbBlue
        Output(X).rgbReserved = (CDbl(Alpha(X).rgbBlue) + _
                                      Alpha(X).rgbGreen + _
                                      Alpha(X).rgbRed) * Inv3
    Next X
Else
'    For x = 0 To nPix - 1
'        Output(x).rgbRed = Bits(x).rgbRed
'        Output(x).rgbGreen = Bits(x).rgbGreen
'        Output(x).rgbBlue = Bits(x).rgbBlue
'        Output(x).rgbReserved = Bits(x).rgbReserved
'    Next x
End If
ExitHere:
UnReferAry AryPtr(Bits), False
UnReferAry AryPtr(Alpha), False
If Err.Number <> 0 Then Err.Raise Err.Number, Err.Source, Err.Description
End Sub

'Returns the new size of Bits in bytes
Private Function BitsPreProcessing24(ByRef InData() As Long, _
                                     ByVal w As Long, _
                                     ByVal h As Long, _
                                     ByRef Bits1() As Byte) As Long
Dim X As Long
Dim ix3 As Long, ix4 As Long
Dim tmp As Byte
Dim ImagePixCount As Long
Dim Y As Long, y1 As Long
Dim tmpArr() As Byte
Dim LineLen As Long
Dim PaddedLineLen As Long
Dim Base As Long, SrcBase As Long
Dim Bits() As Byte
ImagePixCount = w * h
PaddedLineLen = -Int(-w * 3# / 4#) * 4
On Error GoTo ExitHere
ConstructAry AryPtr(Bits), VarPtr(InData(0, 0)), 1, ImagePixCount * 4&
'0123456789
'rgb0rgb0rgb0rgb0
'bgr bgr bgr bgr bgr bgr
'012 345 678 9
ReDim Bits1(0 To PaddedLineLen * h - 1)
For Y = 0 To h - 1&
    Base = Y * PaddedLineLen
    SrcBase = Y * w * 4&
    For X = 0 To w - 1&
        ix3 = X * 3& + Base
        ix4 = X * 4& + SrcBase
        Bits1(ix3) = Bits(ix4)
        Bits1(ix3 + 1&) = Bits(ix4 + 1&)
        Bits1(ix3 + 2&) = Bits(ix4 + 2&)
    Next X
Next Y
BitsPreProcessing24 = PaddedLineLen * h
ExitHere:
UnReferAry AryPtr(Bits), False
If Err.Number <> 0 Then Err.Raise Err.Number, Err.Source, Err.Description
End Function
'
'
'Private Sub BitsPreProcessing8(ByRef InData() As Byte, _
'                               ByRef Bits() As Byte)
'Dim i As Long
'Dim ix3 As Long, ix4 As Long
'Dim tmp As Byte
'Dim ImagePixCount As Long
'Dim y As Long, y1 As Long
'Dim tmpArr() As Byte
'Dim LineLen As Long, LineLenPadded As Long
'Dim w As Long, h As Long
'w = UBound(InData, 1) + 1
'h = UBound(InData, 2) + 1
'
'LineLen = w
'LineLenPadded = -Int(-w / 4#) * 4&
'
'ReDim Bits(0 To LineLenPadded * h - 1)
'
'For y = 0 To h - 1
'    CopyMemory Bits(LineLenPadded * y), InData(0, y), LineLen
'Next y
'End Sub




Public Sub CheckPNGDll()
CheckDll "mpng.dll"
End Sub

'Private Function DLLWorks() As Boolean
'Dim bmi As BITMAPINFO
'Dim File As String
'On Error Resume Next
'bmi.bmiHeader.biBitCount = 0
'bmi.bmiHeader.biHeight = 0
'File = TempPath
'CreateFolder File
'File = File + "test.png"
'Err.Clear
'mWritePNG File, ByVal 0, bmi, 0
'DLLWorks = Err.Number = 0
'End Function

'returns the bpp
Public Function LoadPNG(ByRef File As String, _
                        ByRef Data() As Long, _
                        ByRef AlphaPresent As Boolean) As Long
Dim bmi As BITMAPINFO
Dim hPNG As Long
Dim Length As Long
Dim Ret As Long
Dim w As Long, h As Long
Dim BPP As Long
Dim nColors As Long
Dim i As Long
Dim Bits() As Byte
Dim PaddedLineLen As Long
Dim X As Long, Y As Long
Dim xx3 As Long
Dim BitsRGB() As RGBQUAD
Dim OfcY As Long

Dim pal() As Long

'If Not IsPNGFN(File) Then
'    Err.Raise 32123, "LoadPNG", "Not a png file!"
'End If
CheckPNGDll

Ret = mLoadPNG(File, hPNG, bmi, Length)
If Ret <> 1 Then
    Err.Raise 32123, "LoadPNG", "PNG loading failure. Error code: " + CStr(Ret) + "."
End If
w = bmi.bmiHeader.biWidth
h = Abs(bmi.bmiHeader.biHeight)
BPP = bmi.bmiHeader.biBitCount
With bmi.bmiHeader
    Select Case BPP
        Case 8
            nColors = bmi.bmiHeader.biClrUsed
            If nColors = 0 Then nColors = 256
                
            ReDim pal(0 To 256 - 1)
            
            CopyMemory pal(0), bmi.bmiColors(0), 4& * 256
            'For i = 0 To 256 - 1
            '    Pal(i) = bgr(bmi.bmiColors(i).rgbRed, _
            '                 bmi.bmiColors(i).rgbGreen, _
            '                 bmi.bmiColors(i).rgbBlue)
            'Next i
            
            PaddedLineLen = PadDWORD(w * 1) '-Int(-w / 4#) * 4&
            
            ReDim Bits(0 To PaddedLineLen - 1, 0 To h - 1)
            
            Ret = mGetPNGData(hPNG, Bits(0, 0))
            If Ret <> 1 Then
                mEndPNG hPNG
                Err.Raise 32123, "LoadPNG", "Cannot read the data in PNG. Error code: " + CStr(Ret) + "."
            End If
            
            ReDim Data(0 To w - 1, 0 To h - 1)
            
            For Y = 0 To h - 1
                For X = 0 To w - 1
                    Data(X, Y) = pal(Bits(X, Y))
                Next X
            Next Y
            Erase Bits
            AlphaPresent = False
            
        Case 24
            PaddedLineLen = PadDWORD(w * 3) '-Int(-w * 3# / 4#) * 4&
            ShowStatus "Allocating memory..."
            ReDim Bits(0 To PaddedLineLen * h - 1)
            ShowStatus "Reading png data..."
            Ret = mGetPNGData(hPNG, Bits(0))
            If Ret <> 1 Then
                mEndPNG hPNG
                Err.Raise 32123, "LoadPNG", "Cannot read the data in PNG. Error code: " + CStr(Ret) + "."
            End If
            ShowStatus "Allocating memory..."
            ReDim Data(0 To w - 1, 0 To h - 1)
            ShowStatus "Converting bit depth..."
            ConstructAry AryPtr(BitsRGB), VarPtr(Data(0, 0)), 4, w, h
            On Error GoTo eh
            For Y = 0 To h - 1
                OfcY = Y * PaddedLineLen
                For X = 0 To w - 1
                    xx3 = X * 3& + OfcY
                    BitsRGB(X, Y).rgbBlue = Bits(xx3)
                    BitsRGB(X, Y).rgbGreen = Bits(xx3 + 1&)
                    BitsRGB(X, Y).rgbRed = Bits(xx3 + 2&)
                Next X
            Next Y
            UnReferAry AryPtr(BitsRGB)
            Erase Bits
            ShowStatus "PNG loading done!"
            AlphaPresent = False
        
        Case 32
            ReDim Data(0 To w - 1, 0 To h - 1)
            On Error GoTo eh
            ConstructAry AryPtr(BitsRGB), VarPtr(Data(0, 0)), 4, w, h
            Ret = mGetPNGData(hPNG, BitsRGB(0, 0))
            If Ret <> 1 Then
                mEndPNG hPNG
                Err.Raise 32123, "LoadPNG", "Cannot read the data in PNG. Error code: " + CStr(Ret) + "."
            End If
            ReDim Alpha(0 To w - 1, 0 To h - 1)
'            For y = 0 To h - 1
'                For x = 0 To w - 1
'                    Alpha(x, y) = BitsRGB(x, y).rgbReserved * &H10101
'                    BitsRGB(x, y).rgbReserved = 0&
'                Next x
'            Next y
            UnReferAry AryPtr(BitsRGB)
            AlphaPresent = True
        Case Else
            Err.Raise 32123, "LoadPNG", "Bit count not supported: " + CStr(BPP) + "."
    End Select
    mEndPNG hPNG
End With
LoadPNG = BPP
Exit Function
eh:
UnReferAry AryPtr(BitsRGB)
RaiseError Err
End Function

Private Sub RaiseError(ByRef aErr As ErrObject)
Err.Raise aErr.Number, aErr.Source, aErr.Description
End Sub

Public Function IsPNGFN(ByRef FileName As String) As Boolean
Dim nmb As Long
If Not FileExists(FileName) Then
    Err.Raise 32132, "IsPNG", "File does not exist!"
End If
nmb = FreeFile
Open FileName For Binary Access Read As nmb
    IsPNGFN = IsPNG(nmb)
Close nmb
End Function

Public Function IsPNG(ByVal FileNum As Long)
Dim Bytes() As Byte
'png signature is $89 $50 $4E $47 $0D $0A
'                 ??  P   N   G   cr  lf
ReDim Bytes(0 To 5)
Get FileNum, 1, Bytes
IsPNG = (Bytes(0) = &H89) And _
        (Bytes(1) = &H50) And _
        (Bytes(2) = &H4E) And _
        (Bytes(3) = &H47) And _
        (Bytes(4) = &HD) And _
        (Bytes(5) = &HA)
End Function


Public Function BGR(ByVal red As Long, _
                    ByVal Green As Long, _
                    ByVal Blue As Long)
BGR = &H10000 * red + &H100& * Green + Blue
End Function



Private Function PadDWORD(ByVal Length As Long) As Long
PadDWORD = -Int(-Length / 4#) * 4#
End Function

