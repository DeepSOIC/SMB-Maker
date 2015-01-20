Attribute VB_Name = "mdlIcons"
Option Explicit

'this number of times more important than other vals
Const BrightnessK As Double = 3


Public Enum ePaletteCreationMode
    pcmGenerateBest = 0
    pcmUseStandard = 1
    pcmUseProvided = 2
End Enum

Public Enum eIconType
    itIcon = 1
    itCursor = 2
End Enum

Public Type vtIconImage
    BitDepth As Long
    Data() As Long '1-8 bit - each pixels holds the index to palette
                   '24 bit - the color, 32 - with alpha
    Mask() As Byte 'the monochrome mask. 0 or 1
    Palette() As RGBQUAD
    HotSpot As POINTAPI
End Type

Public Type vtIcon
    Type As eIconType
    Images() As vtIconImage
End Type


Public Type vtIFDImage
    Data() As Long
End Type

Public Type vtIconForDrawing
    Images() As vtIFDImage
End Type

Public Type IconFileHeader
    Reserved As Integer  '=0
    ResType As Integer
    NImages As Integer
End Type

Public Type IconFileImageDesc
    Width As Byte
    Height As Byte
    ColorCount As Byte
    Reserved As Byte
    Planes As Integer 'Holds HotX if cursor
    BitsPerPixel As Integer 'Holds HotY if cursor
    SizeBitmap As Long 'Size of (InfoHeader + XORbitmap + ANDbitmap)
    Offset As Long 'FilePos, where InfoHeader starts
End Type

Dim CurBytes() As Byte 'for memory-icon reading


Public Sub GeneratePal(ByRef Data() As Long, _
                       ByRef Palette() As Long, _
                       ByVal NColorsNeeded As Long)
Dim w As Long, h As Long
Dim PtrElem As Long
Dim Weights() As Long
Dim tmpPal() As Long
w = UBound(Data, 1) + 1
If AryDims(AryPtr(Data)) = 2 Then
    h = UBound(Data, 2) + 1
    PtrElem = VarPtr(Data(0, 0))
Else
    h = 1
    PtrElem = VarPtr(Data(0))
End If
ReDim tmpPal(0 To w * h - 1)
CopyMemory tmpPal(0), ByVal PtrElem, w * h * 4

CalcFreqs tmpPal, Palette, Weights

RemoveFrequent Palette, Weights, w * h, NColorsNeeded
'if arydims(aryptr(data))
End Sub

Public Sub CalcFreqs(ByRef Palette() As Long, _
                     ByRef OutPal() As Long, _
                     ByRef Weights() As Long)
Dim i As Long, j As Long
Dim Last As Long
Dim cnt As Long
OutPal = Palette
SortLongArray OutPal, 0, UBound(OutPal)
Last = OutPal(0) - 1
ReDim Weights(0 To UBound(OutPal))
For i = 0 To UBound(OutPal)
    If Last = OutPal(i) Then
        cnt = cnt + 1&
    Else
        Weights(j) = cnt + 1
        Last = OutPal(i)
        OutPal(j) = Last
        j = j + 1&
    End If
Next i
ReDim Preserve OutPal(0 To j - 1)
ReDim Preserve Weights(0 To j - 1)
End Sub

Public Sub RemoveFrequent(ByRef Palette() As Long, _
                          ByRef Weights() As Long, _
                          ByRef nPixels As Long, _
                          ByVal DesiredNColors As Long)
Const LeastSignificantK As Double = 0.01
Dim LUT() As Long
Dim nColorsToRemove1 As Long, nColorsToRemove2
Dim nColors As Long
'Dim nFailures As Long, nRemoved As Long
Dim FrToFind As Long
Dim cnt As Long
Dim i As Long
Dim j As Long
Dim ThePos As Long
Dim OutPal() As Long, OutWeights() As Long
FrToFind = LeastSignificantK * nPixels
GenerateLUT Weights, LUT
nColors = AryLen(AryPtr(Palette))
    
ThePos = BinarySearchLng(Weights, FrToFind, 0, nColors - 1, AryPtr(LUT))
i = ThePos
j = ThePos
For cnt = 1 To nColors - DesiredNColors
    If (i - ThePos) * ThePos > (ThePos - j) * (nColors - ThePos) Then
        j = j - 1&
        Weights(LUT(j)) = 0
    Else
        Weights(LUT(i)) = 0
        i = i + 1&
    End If
Next cnt
    
j = 0
ReDim OutPal(0 To nColors - 1)
ReDim OutWeights(0 To nColors - 1)
For i = 0 To nColors - 1
    If Weights(LUT(i)) Then
        OutPal(j) = Palette(LUT(i))
        OutWeights(j) = Weights(LUT(i))
        j = j + 1&
    End If
Next i
Debug.Assert j = DesiredNColors
SwapArys AryPtr(OutPal), AryPtr(Palette)
SwapArys AryPtr(OutWeights), AryPtr(Weights)
End Sub

Public Sub SortByBrightness(Palette() As RGBQUAD, _
                            Brightness() As Long, _
                            LUT() As Long)

Dim brt As Long
Dim i As Long
Dim nColors As Long
'Dim LUT() As Long
'Dim OutPal() As Long, OutBrt() As Long
nColors = AryLen(AryPtr(Palette))
ReDim Brightness(0 To nColors - 1)
For i = 0 To nColors - 1
    brt = CLng(Palette(i).rgbBlue) + Palette(i).rgbGreen + Palette(i).rgbRed
    Brightness(i) = Round(brt / 3 / BrightnessK)
Next i

GenerateLUT Brightness, LUT

'ReDim OutPal(0 To nColors - 1)
'ReDim OutBrt(0 To nColors - 1)
'For i = 0 To nColors - 1
'    OutPal(i) = LUT(Palette(i))
'    OutBrt(i) = LUT(Brightness(i))
'Next i
'SwapArys AryPtr(Palette), AryPtr(OutPal)
'SwapArys AryPtr(Brightness), AryPtr(OutBrt)
End Sub

Public Sub ToIndexedColors(ByRef Data() As RGBQUAD, _
                           ByRef Palette() As RGBQUAD, _
                           ByRef iData() As Long)
Dim w As Long, h As Long
Dim X As Long, Y As Long
Dim i As Long, j As Long
Dim brt As Long
Dim LUT() As Long
Dim Brightness() As Long
Dim LBr() As Long, RBr() As Long 'ranges of colors with this brightness
Dim MaxBr As Long
Dim nColors As Long
Dim LastBr As Long
Dim TheIndex As Long, HowCloseIsTheIndex As Long
Dim Cmp As Long
Dim MidBr As Long, NewBr As Long
Dim br As Long

SortByBrightness Palette, Brightness, LUT
nColors = AryLen(AryPtr(Palette))

MaxBr = Round(255 / BrightnessK)
ReDim IByBr(0 To MaxBr)

For br = 0 To MaxBr
    i = BinarySearchLng(Brightness, br, 0, nColors - 1, AryPtr(LUT))
    If i < nColors Then
        j = i - 1
        If Brightness(LUT(i)) = br Then
            Cmp = br
        Else
            Cmp = Brightness(LUT(Max(0, j)))
        End If
        Do While j >= 0
            If Brightness(j) <> Cmp Then Exit Do
            j = j - 1
        Loop
        LBr(br) = j + 1
        
        j = i + 1
        If Brightness(LUT(j)) = br Then
        Else
            Cmp = Brightness(LUT(Min(nColors - 1, j)))
        End If
        Do While j < nColors
            If Brightness(j) <> Cmp Then Exit Do
            j = j + 1
        Loop
        RBr(br) = j - 1
    Else
        If br > 0 Then
            RBr(br) = RBr(br - 1)
            LBr(br) = LBr(br - 1)
        End If
    End If
Next br
'LastBr = Brightness(LUT(0)) - 1
'For i = 0 To nColors - 1
'    If Brightness(LUT(i)) <> LastBr Then
'        NewBr = Brightness(LUT(i))
'        MidBr = (LastBr + NewBr) \ 2
'        For j = LastBr + 1 To MidBr
'            LBr(j) = LBr(LastBr)
'            RBr(j) = RBr(LastBr)
'        Next j
'        If (LastBr + Brightness(LUT(i))) Mod 2 = 0 Then
'            j = MidBr
'            LBr(j) = LBr(LastBr)
'            RBr(j) = RBr(NewBr)
'        End If
'        For j = MidBr + 1 To NewBr
'
'        Next j
'        LastBr = Brightness(LUT(i))
'        LBr(LastBr) = i
'    Else
'        RBr(LastBr) = i
'    End If
'Next i

Cmp = 0 ' reset cmp because of it's new meaning

AryWH AryPtr(Data), w, h
ReDim iData(0 To w - 1, 0 To h - 1)
For Y = 0 To h - 1
    For X = 0 To w - 1
        brt = CLng(Data(X, Y).rgbBlue) + _
              Data(X, Y).rgbBlue + _
              Data(X, Y).rgbBlue
        brt = Round(brt / 3 / BrightnessK)
        
        HowCloseIsTheIndex = 100000
        For i = LBr(brt) To RBr(brt)
            Cmp = Abs(CLng(Data(X, Y).rgbBlue) - Palette(LUT(i)).rgbBlue) + _
                  Abs(CLng(Data(X, Y).rgbGreen) - Palette(LUT(i)).rgbGreen) + _
                  Abs(CLng(Data(X, Y).rgbRed) - Palette(LUT(i)).rgbRed)
            If Cmp < HowCloseIsTheIndex Then
                HowCloseIsTheIndex = Cmp
                TheIndex = i
            End If
        Next i
        iData(X, Y) = TheIndex
    Next X
Next Y
End Sub



Public Sub WriteIcon(ByRef FileName As Long, _
                     ByRef Icon As vtIcon)
Dim nmb As Long
'Structures
Dim FileHeader As IconFileHeader
Dim ImageDesc As IconFileImageDesc
Dim bmiH As BITMAPINFOHEADER
'/Structures

'Lengths
Dim nIcons As Long
Dim w As Long, h As Long
Dim XORWidthBytes As Long, ANDWidthBytes As Long
Dim XORLen As Long, ANDLen As Long, HeaderLen As Long
'/Lengths

'Arrays
Dim DataBytes() As Byte
Dim RGBData() As RGBQUAD
'/arrays

Dim ErrorIgnorable As Boolean

'loop vars
Dim iIcon As Long
Dim X As Long, Y As Long
Dim y1 As Long
'/loop vars

'positions
Dim IconDescsOffset As Long
Dim BitPos As Long, BytePos As Long
Dim CurPos As Long 'the write position in the bitmap heap
'/positions

Dim BitsKick As Long
nIcons = AryLen(AryPtr(Icon.Images))

StartWrite FileName

nmb = FreeFile
Open FileName For Binary Access Write As nmb
    
    On Error GoTo eh
    
    With FileHeader
        .Reserved = 0
        .ResType = Icon.Type
        .NImages = nIcons
    End With
    
    Put nmb, 1, FileHeader
    CurPos = CurPos + Len(FileHeader)
    
    IconDescsOffset = CurPos
    CurPos = IconDescsOffset + Len(ImageDesc) * nIcons
    
    For iIcon = 0 To nIcons - 1
        With ImageDesc
            AryWH AryPtr(Icon.Images(iIcon).Data), w, h
            If w > 255 Or h > 255 Then
                ErrorIgnorable = True
                Err.Raise 1212, "WriteIcon", "Image width and height have to be less then 255. If you ignore this error, the icon may be unreadable."
                ErrorIgnorable = False
            End If
            .Width = Min(w, 255)
            .Height = Min(h, 255)
            
            .Planes = 1
            .BitsPerPixel = Icon.Images(iIcon).BitDepth
            .ColorCount = IIf(.BitsPerPixel < 8, 2& ^ .BitsPerPixel, 0&)
            .Offset = CurPos
            
            XORWidthBytes = ((w * .BitsPerPixel + 31&) And -32&) \ 8&
            ANDWidthBytes = ((w * 1& + 31&) And -32&) \ 8&
            
            HeaderLen = Len(bmiH)
            XORLen = XORWidthBytes * h
            ANDLen = ANDWidthBytes * h
            
            .SizeBitmap = HeaderLen + XORLen + ANDLen
            
            If Icon.Type = itCursor Then
                .Planes = Icon.Images(iIcon).HotSpot.X
                .BitsPerPixel = Icon.Images(iIcon).HotSpot.X
            End If
        End With
        Put nmb, IconDescsOffset + Len(ImageDesc) * iIcon + 1, ImageDesc
        
        With bmiH
            .biSize = Len(bmiH)
            
            .biBitCount = Icon.Images(iIcon).BitDepth
            .biWidth = w
            .biHeight = h * 2
            .biPlanes = 1
            
            .biSizeImage = 0&
        End With
        Put nmb, CurPos + 1, bmiH
        
        ReDim DataBytes(0 To XORWidthBytes - 1, 0 To h - 1)
        
        With Icon.Images(iIcon)
            Select Case .BitDepth
                Case 1, 4, 8
                    
                    BitsKick = 2& ^ .BitDepth
                    For Y = 0 To h - 1
                        y1 = h - 1 - Y
                        For X = 0 To w - 1
                            BitPos = X * .BitDepth
                            BytePos = BitPos \ 8&
                            DataBytes(BytePos, y1) = DataBytes(BytePos, y1) * BitsKick Or .Data(X, Y)
                        Next X
                    Next Y
                
                Case 24
                    ReferAry AryPtr(RGBData), AryPtr(.Data)
                        For Y = 0 To h - 1
                            y1 = h - 1 - Y
                            For X = 0 To w - 1
                                DataBytes(X * 3, y1) = RGBData(X, Y).rgbBlue
                                DataBytes(X * 3 + 1, y1) = RGBData(X, Y).rgbGreen
                                DataBytes(X * 3 + 2, y1) = RGBData(X, Y).rgbRed
                            Next X
                        Next Y
                    UnReferAry AryPtr(RGBData)
                Case 32
                        For Y = 0 To h - 1
                            y1 = h - 1 - Y
                            CopyMemory DataBytes(0, y1), .Data(0, Y), w * 4
                        Next Y
            End Select
            
            Put nmb, CurPos, DataBytes
            CurPos = CurPos + XORWidthBytes * h
        
            If IsAryEmpty(AryPtr(.Mask)) And .BitDepth = 32 Then
                ReferAry AryPtr(RGBData), AryPtr(.Data)
                    ReDim .Mask(0 To w - 1, 0 To h - 1)
                    For Y = 0 To h - 1
                        For X = 0 To w - 1
                            .Mask(X, Y) = 1& - RGBData(X, Y).rgbReserved \ 128&
                        Next X
                    Next Y
                UnReferAry AryPtr(RGBData)
            End If
            
            
            ReDim DataBytes(0 To ANDWidthBytes - 1, 0 To h - 1)
            
            BitsKick = 2&
            For Y = 0 To h - 1
                y1 = h - 1 - Y
                For X = 0 To w - 1
                    BitPos = X
                    BytePos = BitPos \ 8&
                    DataBytes(BytePos, y1) = DataBytes(BytePos, y1) * BitsKick Or .Mask(X, Y)
                Next X
            Next Y
            
            Put nmb, CurPos + 1, DataBytes
            CurPos = CurPos + ANDWidthBytes * h
            
        End With
        
        
    Next iIcon
Close nmb

Exit Sub
eh:
If ErrorIgnorable Then
    'PushError
    Select Case MsgError(, vbAbortRetryIgnore)
        Case vbAbort
            Err.Raise dbCWS, "WriteIcon"
        Case vbRetry
            Resume
        Case vbIgnore
            Resume Next
    End Select
Else
    PushError
        UnReferAry AryPtr(RGBData)
        Close nmb
    PopError
    ErrRaise "WriteIcon"
End If
End Sub


Public Sub CreateIconImage(ByRef II As vtIconImage, _
                           ByRef Data() As Long, _
                           ByVal BitDepth As Long, _
                           ByRef Palette() As Long, _
                           ByVal PaletteCreationMode As ePaletteCreationMode)
Dim Brightness() As Long
Dim ColorsCount As Long
Dim l As Long
Dim DefPal() As Long
Dim RGBPal() As RGBQUAD
Dim i As Long
Dim RGBData() As RGBQUAD

Select Case BitDepth
    Case 1, 4, 8
        'indexed image
        ColorsCount = 2 ^ BitDepth
        Select Case PaletteCreationMode
            Case ePaletteCreationMode.pcmGenerateBest
                GeneratePal Data, Palette, ColorsCount
            Case ePaletteCreationMode.pcmUseProvided
                'do nothing
            Case ePaletteCreationMode.pcmUseStandard
                ReDim Palette(0 To ColorsCount - 1)
                LoadDefPal Palette
        End Select
        
        l = AryLen(AryPtr(Palette))
        If l <> ColorsCount Then
            If l < ColorsCount Then
            
                ReDim DefPal(0 To ColorsCount - 1)
                LoadDefPal DefPal
                
                If l > 0 Then
                    ReDim Preserve Palette(0 To ColorsCount - 1)
                Else
                    ReDim Palette(0 To ColorsCount - 1)
                End If
                
                For i = l To ColorsCount - 1
                    Palette(i) = DefPal(i)
                Next i
                
                Erase DefPal
                
            ElseIf l > ColorsCount Then
                
                ReDim Preserve Palette(0 To ColorsCount - 1)
            
            End If
        End If
        
        II.BitDepth = BitDepth
        On Error GoTo eh
        If AryLen(AryPtr(Palette)) <= 0 Then Err.Raise 1212, , "The palette cannot be empty at this stage! (internal error)"
        ReDim RGBPal(0 To UBound(Palette))
        CopyMemory RGBPal(0), Palette(0), (UBound(Palette) + 1) * 4
        ReferAry AryPtr(RGBData), AryPtr(Data)
        ToIndexedColors RGBData, RGBPal, II.Data
        UnReferAry AryPtr(RGBData)
        
        II.Palette = RGBPal
        
        
    Case 24, 32
        II.BitDepth = BitDepth
        II.Data = Data
        Erase II.Palette
        Erase II.Mask
End Select

Exit Sub
eh:
UnReferAry AryPtr(RGBData)
ErrRaise "CreateIconImage"
End Sub


Public Sub ReadIconFile(ByRef File As String, ByRef Icon As vtIcon)
Dim Bytes() As Byte
Dim Length As Long
Dim nmb As Long
On Error GoTo eh
Open File For Binary Access Read As nmb
    Length = LOF(nmb)
    If Length > 64& * 1024 * 1024 Then
        If MsgBox("The file is very large. It may be something other than an icon. Are you sure you want to continue loadint it? Note: SMB Maker will load the file into the memory, and only after than parse it.", vbYesNo) = vbNo Then
            Err.Raise dbCWS
        End If
    End If
    ReDim Bytes(0 To Length - 1)
    Get nmb, 1, Bytes
    ReadIconMem Bytes, Icon
Close nmb

Exit Sub
eh:
PushError
Close nmb
PopError
ErrRaise "ReadIconFile"
End Sub


Public Sub ReadIconMem(ByRef File() As Byte, _
                       ByRef Icon As vtIcon)
Dim nmb As Long
'Loop counters
Dim iIcon As Long
Dim X As Long, Y As Long
Dim y1 As Long
'/Loop counters

'Offsets
Dim OfcData As Long
Dim ofcDescs As Long
'/Offsets

'Lengths
Dim nIcons As Long
Dim w As Long, h As Long
Dim XORWidthBytes As Long, ANDWidthBytes As Long
Dim nColors As Long
'/Lengths

'Bit operations
Dim BPP As Long 'bits per pixel
Dim BitPos As Long
Dim BitMask As Long
Dim BitsKick As Long
Dim BytePos As Long
'/Bit operations

'Structures
Dim FileHeader As IconFileHeader
Dim ImageDesc As IconFileImageDesc
Dim bmiH As BITMAPINFOHEADER
'/Structures

'Arrays
Dim DataBytes() As Byte
Dim RGBData() As RGBQUAD 'for mappint to Icon.Image().Data
'/Arrays

Dim ErrorIgnorable As Boolean

'nmb = FreeFile
'Open File For Binary Access Read As nmb
SwapArys AryPtr(File), AryPtr(CurBytes)
    On Error GoTo eh
    MemGet 0, VarPtr(FileHeader), Len(FileHeader)
    
    If FileHeader.Reserved <> 0 Then
        Err.Raise 1212, , "Not an Icon file. Error in header."
    End If
    Select Case FileHeader.ResType
        Case eIconType.itCursor, eIconType.itIcon
            Icon.Type = FileHeader.ResType
        Case Else
            Err.Raise 1212, , "Unknown resource type in the icon."
    End Select
    
    nIcons = FileHeader.NImages
    
    If nIcons < 0 Then Err.Raise 1212, , "Number of icons in the file is negative :("
    
    If nIcons > 0 Then
        ReDim Icon.Images(0 To nIcons - 1)
    Else
        Erase Icon.Images
    End If
    ofcDescs = Len(FileHeader)
    For iIcon = 0 To nIcons - 1
        MemGet ofcDescs + iIcon * Len(ImageDesc), VarPtr(ImageDesc), Len(ImageDesc)
        
        With ImageDesc
            If Icon.Type = itIcon Then
                BPP = .BitsPerPixel
                If .Planes <> 1 Then
                    ErrorIgnorable = True
                    Err.Raise 1212, , "Warning: The number of color planes should be 1, but is " + CStr(.Planes) + " :("
                    ErrorIgnorable = False
                End If
            Else
                With Icon.Images(iIcon).HotSpot
                    .X = ImageDesc.Planes
                    .Y = ImageDesc.BitsPerPixel
                End With
            End If
            
            w = .Width
            h = .Height
            OfcData = .Offset
        End With
        
        MemGet OfcData, VarPtr(bmiH), Len(bmiH)
        With bmiH
            If .biSize <> Len(bmiH) Then
                ErrorIgnorable = True
                Err.Raise 1212, , "Some other version than 3 of bitmap info header is used. The bitmap may load incorrectly."
            End If
            
            If Icon.Type = itCursor Then BPP = .biBitCount
            If .biBitCount <> ImageDesc.BitsPerPixel Then
                ErrorIgnorable = True
                Err.Raise 1212, , "Bit depths indicated in image descripor and in bitmap info header do not equal! If ignored, the one from info header will be used."
                ErrorIgnorable = False
            End If

            If .biWidth <> w Or .biHeight <> h * 2 Then
                ErrorIgnorable = True
                Err.Raise 1212, , "Widths indicated in image descripor and in bitmap info header do not equal! If ignored, the one from info header will be used."
                ErrorIgnorable = False
            End If
            If .biCompression <> 0 Then
                ErrorIgnorable = True
                Err.Raise 1212, , "This icon might have compressed images. Compressed icons are not supported by SMB Maker. If ignored, the data will be treated as uncompressed."
                ErrorIgnorable = False
            End If
            
            BPP = .biBitCount
            If BPP <= 8 Then nColors = 2& ^ BPP Else nColors = 0
            w = .biWidth
            h = .biHeight \ 2
            
            XORWidthBytes = ((w * BPP + 31) And -32) \ 8
            ANDWidthBytes = ((w * 1 + 31) And -32) \ 8
            
        End With
        
        With Icon.Images(iIcon)
            
            If BPP <= 8 Then
                ReDim .Palette(0 To nColors - 1)
                MemGet OfcData + Len(bmiH), VarPtr(.Palette(0)), nColors * 4
            End If
        
            ReDim DataBytes(0 To XORWidthBytes - 1, 0 To h - 1)
            MemGet OfcData + Len(bmiH) + nColors * 4, VarPtr(DataBytes(0, 0)), XORWidthBytes * h
        
            ReDim .Data(0 To w - 1, 0 To h - 1)
            
            Select Case BPP
                Case 1, 4, 8
                    BitsKick = 2 ^ BPP
                    BitMask = BitsKick - 1
                    For Y = 0 To h - 1
                        y1 = h - 1 - Y
                        For X = 0 To w - 1
                            BitPos = X * BPP
                            BytePos = BitPos \ 8
                            .Data(X, y1) = DataBytes(BytePos, Y) And BitMask
                            DataBytes(BytePos, Y) = DataBytes(BytePos, Y) \ BitsKick
                        Next X
                    Next Y
                Case 24
                    ReferAry AryPtr(RGBData), AryPtr(.Data)
                        For Y = 0 To h - 1
                            y1 = h - 1 - Y
                            For X = 0 To w - 1
                                RGBData(X, y1).rgbBlue = DataBytes(X * 3, Y)
                                RGBData(X, y1).rgbGreen = DataBytes(X * 3 + 1, Y)
                                RGBData(X, y1).rgbRed = DataBytes(X * 3 + 2, Y)
                            Next X
                        Next Y
                    UnReferAry AryPtr(RGBData)
                Case 32
                    For Y = 0 To h - 1
                        y1 = h - 1 - Y
                        CopyMemory .Data(0, y1), DataBytes(0, Y), w * 4
                    Next Y
                Case Else
                    Err.Raise 1212, , "BitsPerPixel = " + CStr(BPP) + " is not supported."
            End Select
        
            ReDim DataBytes(0 To ANDWidthBytes - 1, 0 To h - 1)
            MemGet OfcData + Len(bmiH) + nColors * 4 + XORWidthBytes * h, VarPtr(DataBytes(0, 0)), ANDWidthBytes * h
            
            BitsKick = 2
            BitMask = BitsKick - 1
            For Y = 0 To h - 1
                y1 = h - 1 - Y
                For X = 0 To w - 1
                    BitPos = X * BPP
                    BytePos = BitPos \ 8
                    .Mask(X, y1) = DataBytes(BytePos, Y) And BitMask
                    DataBytes(BytePos, Y) = DataBytes(BytePos, Y) \ BitsKick
                Next X
            Next Y
            
        End With
    Next iIcon
    
    
    
SwapArys AryPtr(CurBytes), AryPtr(File)
Erase CurBytes

Exit Sub
eh:
If ErrorIgnorable Then
    Select Case MsgError(, vbAbortRetryIgnore)
        Case vbAbort
            Err.Raise dbCWS, "ReadIcon"
        Case vbRetry
            Resume
        Case vbIgnore
            Resume Next
    End Select
Else
    PushError
        UnReferAry AryPtr(RGBData)
        SwapArys AryPtr(CurBytes), AryPtr(File)
        Erase CurBytes
    PopError
    ErrRaise
End If
End Sub

Public Sub MemGet(ByVal Offset As Long, _
                  ByVal ptrReadTo As Long, _
                  ByVal Length As Long)
Dim l As Long
If Offset < 0 Then Exit Sub
l = AryLen(AryPtr(CurBytes))
Length = Min(Length, l - Offset)
If Length > 0 Then
    CopyMemory ByVal ptrReadTo, CurBytes(Offset), Length
End If
End Sub

Public Sub MakeIconForDrawing(ByRef IFD As vtIconForDrawing, _
                              ByRef Icon As vtIcon)
Dim nIcons As Long
Dim w As Long, h As Long
Dim X As Long, Y As Long
Dim iIcon As Long
Dim RGBData() As RGBQUAD
Dim lngData() As Long
Dim Palette() As Long

nIcons = AryLen(AryPtr(Icon.Images))
If nIcons > 0 Then
    ReDim IFD.Images(0 To nIcons - 1)
Else
    Erase IFD.Images
End If
On Error GoTo eh
For iIcon = 0 To nIcons - 1
    AryWH AryPtr(Icon.Images(iIcon).Data), w, h
    
    ReDim IFD.Images(iIcon).Data(0 To w - 1, 0 To h - 1)
    
    ReferAry AryPtr(lngData), AryPtr(IFD.Images(iIcon).Data)
    ReferAry AryPtr(RGBData), AryPtr(IFD.Images(iIcon).Data)
'    If BPP <= 8 Then
'        ReferAry AryPtr(Palette), AryPtr(Icon.Images(iIcon).Palette)
'    End If
    With Icon.Images(iIcon)
        Select Case .BitDepth
            Case 1, 4, 8
                For Y = 0 To h - 1
                    For X = 0 To w - 1
                        RGBData(X, Y) = .Palette(.Data(X, Y))
                    Next X
                Next Y
            
            Case 24, 32
                CopyMemory lngData(0, 0), .Data(0, 0), w * h * 4
        End Select
        
        If .BitDepth = 1 Or _
           .BitDepth = 4 Or _
           .BitDepth = 8 Or _
           .BitDepth = 24 Then
            For Y = 0 To h - 1
                For X = 0 To w - 1
                    RGBData(X, Y).rgbReserved = Not .Mask(X, Y) * CByte(&HFF)
                    'LngData(x, y) = (LngData(x, y) And &HFFFFFF) Or (Not (.Mask(x, y) * &HFF000000) And &HFF000000)
                Next X
            Next Y
        End If
    End With
    UnReferAry AryPtr(RGBData)
    UnReferAry AryPtr(lngData)
'    UnReferAry AryPtr(Palette)
Next iIcon

Exit Sub
eh:
PushError
    UnReferAry AryPtr(RGBData)
    UnReferAry AryPtr(lngData)
'    UnReferAry AryPtr(Palette)
PopError
ErrRaise "MakeIconForDrawing"
End Sub


Public Sub vtAlphaBlend(ByRef DestData() As Long, _
                        ByRef SourceData() As Long, _
                        ByVal Left As Long, _
                        ByVal Top As Long)
Dim X As Long, Y As Long
Dim SrcW As Long, SrcH As Long
Dim DstW As Long, DstH As Long
Dim DstRct As RECT
Dim SrcRct As RECT 'all rects with respect to the source
Dim RGBSrc() As RGBQUAD
Dim RGBDst() As RGBQUAD
Dim a As Long
Dim OfcSrc As Long, OfcDst As Long
AryWH AryPtr(DestData), DstW, DstH
AryWH AryPtr(SourceData), SrcW, SrcH
With DstRct
    .Left = -Left
    .Top = -Top
    .Right = .Left + DstW
    .Bottom = .Top + DstH
End With
With SrcRct
    .Left = 0
    .Top = 0
    .Right = .Left + SrcW
    .Bottom = .Top + SrcH
End With
DstRct = IntersectRects(DstRct, SrcRct)

If IsRectEmpty(DstRct) Then Exit Sub

On Error GoTo eh
ConstructAry AryPtr(RGBSrc), VarPtr(SourceData(0, 0)), 4, SrcW * SrcH
ConstructAry AryPtr(RGBDst), VarPtr(DestData(0, 0)), 4, DstW * DstH
For Y = DstRct.Top To DstRct.Bottom - 1
    OfcSrc = Y * SrcW
    OfcDst = (Y + Top) * DstW + Left
    For X = DstRct.Left To DstRct.Right - 1
        a = RGBSrc(OfcSrc + X).rgbReserved
        RGBDst(OfcDst + X).rgbBlue = _
           RGBDst(OfcDst + X).rgbBlue + _
           ((CLng(RGBSrc(OfcSrc + X).rgbBlue) - RGBDst(OfcDst + X).rgbBlue) * a + 127&) \ 255
        RGBDst(OfcDst + X).rgbGreen = _
           RGBDst(OfcDst + X).rgbGreen + _
           ((CLng(RGBSrc(OfcSrc + X).rgbGreen) - RGBDst(OfcDst + X).rgbGreen) * a + 127&) \ 255
        RGBDst(OfcDst + X).rgbRed = _
           RGBDst(OfcDst + X).rgbRed + _
           ((CLng(RGBSrc(OfcSrc + X).rgbRed) - RGBDst(OfcDst + X).rgbRed) * a + 127&) \ 255
    Next X
Next Y
UnReferAry AryPtr(RGBSrc)
UnReferAry AryPtr(RGBDst)
Exit Sub
eh:
UnReferAry AryPtr(RGBSrc)
UnReferAry AryPtr(RGBDst)
ErrRaise "vtAlphaBlend"
End Sub
