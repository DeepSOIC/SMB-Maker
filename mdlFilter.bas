Attribute VB_Name = "mdlFilter"
Option Explicit

Private Declare Sub ZeroMemory Lib "kernel32" Alias "RtlZeroMemory" _
            (Dest As Any, ByVal numBytes As Long)

'Dim ArString() As RGBTriLong
Dim ArStringR() As Long
Dim ArStringG() As Long
Dim ArStringB() As Long
'Dim ArStringCounts() As RGBTriLong
Dim ASCR() As Long
Dim ASCG() As Long
Dim ASCB() As Long
Dim TextureMode As Boolean

Dim Pic() As RGBQUAD
Dim PicW As Long, PicH As Long
Dim MaskW As Long, MaskH As Long
Dim xFrom As Long, yFrom As Long
Dim xTo As Long, yTo As Long
Dim DiffMode As Boolean


Public Type typFilterOpts
    FilterMode As eFilterMode
    Brightness As Double
    Absolute As Boolean
    TextureMode As Boolean
End Type

Public Enum eFilterMode
    fmNormal = 0
    fmDifference = 1
    fmAnti = 2
End Enum

Public Type PointDBL
    x As Double
    y As Double
End Type

Public Const MassCentreError = 11111

Private Sub ShowProgress(ByVal Progress As Double)
MainForm.ShowProgress Progress * 100, DoDoEvents:=True
End Sub

Public Sub vtFilter(ByRef InData() As Long, _
                    ByRef Mask As FilterMask, _
                    ByRef FilterOpts As typFilterOpts, _
                    ByRef Region As RECT, _
                    ByRef LOutData() As Long)
Const n As Long = 0
Dim xf As Long, xt As Long
Dim yf As Long, yt As Long
Dim MC As PointDBL
Dim pMask() As RGBTriLong
Dim OutData() As RGBQUAD
Dim OutDatav() As RGBQUAD 'mapped to OutData
xf = Region.Left
xt = Region.Right - 1
yf = Region.Top
yt = Region.Bottom - 1
Dim y As Long, ty As Long
Dim x As Long
Dim MskX As Long, MskY As Long
Dim OfcX As Long, OfcY As Long
Dim MaskMass As RGBTriLong

Dim Brightness As Double
Dim CX As Long, CY As Long
Dim Absolute As Boolean

Dim t As Long
Dim cb As Long
Dim CountWholeString As Boolean
Dim UseCountsWholeString As Boolean

Dim xMinus As Long, xPlus As Long
Dim yMinus As Long, yPlus As Long

Dim OutW As Long, OutH As Long
Dim OutBaseX As Long, OutBaseY As Long
Dim OutX As Long, OutY As Long
OutBaseX = Region.Left
OutBaseY = Region.Top
OutW = Region.Right - Region.Left
OutH = Region.Bottom - Region.Top
Dim OfcOutY As Long

'Test input
If AryDims(AryPtr(InData)) <> 2 Or AryDims(AryPtr(Mask.Mask)) <> 2 Then
    Err.Raise 12321, "vtFilter", "A 2-dimensional array is required."
End If

'process mask
If Not Mask.CenterFilled Then
    MC = MaskMassCentre(Mask.Mask)
    Mask.Center.x = MC.x
    Mask.Center.y = MC.y
    Mask.CenterFilled = True
End If

MaskMass = MaskWeight(Mask.Mask, FixZeroProblem:=True)

GenerateModedMask pMask, Mask, MaskMass, FilterOpts
DiffMode = FilterOpts.FilterMode = fmDifference

'fill in mode
TextureMode = FilterOpts.TextureMode
Brightness = FilterOpts.Brightness
Absolute = FilterOpts.Absolute

'Fill in vars
PicW = UBound(InData, 1) + 1
PicH = UBound(InData, 2) + 1
MaskW = UBound(pMask, 1) + 1
MaskH = UBound(pMask, 2) + 1
'test mask centre
If Mask.Center.x < 0 Or _
   Mask.Center.x >= MaskW Or _
   Mask.Center.y < 0 Or _
   Mask.Center.y >= MaskH Then
   Err.Raise 12321, , "Mask center point is outside of it."
End If
'end test mask centre

xf = Max(0, xf)
yf = Max(0, yf)
xt = Min(xt, PicW - 1)
yt = Min(yt, PicH - 1)
'test empty region
If xt < xf Or yt < yf Then Exit Sub 'empty region - nothing to do
'end test empty region
xFrom = xf
xTo = xt

CX = Mask.Center.x
CY = Mask.Center.y
xMinus = CX
xPlus = MaskW - 1 - CX
yMinus = CY
yPlus = MaskH - 1 - CY

'Allocate memory and map arrays
ReDim OutData(0 To OutW - 1, 0 To OutH - 1)
On Error GoTo eh
ConstructAry AryPtr(Pic), VarPtr(InData(0, 0)), 4, PicW * PicH
ConstructAry AryPtr(OutDatav), VarPtr(OutData(0, 0)), 4, OutW * OutH

'initialize temporary storage
ReDim ArStringR(0 To PicW - 1)
ReDim ArStringG(0 To PicW - 1)
ReDim ArStringB(0 To PicW - 1)
If Not TextureMode Then
    ReDim ASCR(0 To PicW - 1)
    ReDim ASCG(0 To PicW - 1)
    ReDim ASCB(0 To PicW - 1)
End If

'Processing
For y = yf To yt
    OutY = y - OutBaseY
    CountWholeString = y < 2 * yMinus Or PicH - 1 - y < 2 * yPlus
    UseCountsWholeString = y < yMinus Or PicH - 1 - y < yPlus
    'clean up the accumulator string
    ZeroMemory ArStringR(xf), (xt - xf + 1) * 4&
    ZeroMemory ArStringG(xf), (xt - xf + 1) * 4&
    ZeroMemory ArStringB(xf), (xt - xf + 1) * 4&
    If Not TextureMode Then
        'clean only areas near the border
        If CountWholeString Or True Then
            ZeroMemory ASCR(xf), (xt - xf + 1) * 4&
            ZeroMemory ASCG(xf), (xt - xf + 1) * 4&
            ZeroMemory ASCB(xf), (xt - xf + 1) * 4&
        Else
            xFrom = Max(0, xf)
            xTo = Min(2 * xMinus - 1 + n, xt)
            cb = xTo - xFrom + 1
            If cb > 0 Then
                ZeroMemory ASCR(xFrom), cb * 4&
                ZeroMemory ASCG(xFrom), cb * 4&
                ZeroMemory ASCB(xFrom), cb * 4&
            End If
            
            xFrom = Max(Max(PicW - 1 - 2 * xPlus + 1 - n, xf), 2 * xMinus + n)
            xTo = Min(PicW - 1, xt)
            cb = xTo - xFrom + 1
            If cb > 0 Then
                ZeroMemory ASCR(xFrom), cb * 4&
                ZeroMemory ASCG(xFrom), cb * 4&
                ZeroMemory ASCB(xFrom), cb * 4&
            End If
        End If
    End If
    
    'add multiplied strings to accumulator
    For MskY = 0 To MaskH - 1
        For MskX = 0 To MaskW - 1
            ty = y + MskY - CY
            If TextureMode Then
                ty = ty Mod PicH
                If ty < 0 Then ty = ty + PicH
                AddToString XOffset:=MskX - CX, _
                            y:=ty, _
                            Multiplier:=pMask(MskX, MskY), _
                            CountSumMul:=False
            Else
                If ty >= 0 And ty < PicH Then
                    OfcY = ty * PicW
                    If CountWholeString Then
                        xFrom = xf
                        xTo = xt
                        AddToString XOffset:=MskX - CX, _
                                    y:=ty, _
                                    Multiplier:=pMask(MskX, MskY), _
                                    CountSumMul:=True
                    Else
                        xFrom = Max(0, xf)
                        xTo = Min(2 * xMinus - 1 + n, xt)
                        AddToString XOffset:=MskX - CX, _
                                    y:=ty, _
                                    Multiplier:=pMask(MskX, MskY), _
                                    CountSumMul:=True
                        xFrom = Max(2 * xMinus + n, xf)
                        xTo = Min(PicW - 1 - 2 * xPlus - n, xt)
                        AddToString XOffset:=MskX - CX, _
                                    y:=ty, _
                                    Multiplier:=pMask(MskX, MskY), _
                                    CountSumMul:=False
                        xFrom = Max(Max(PicW - 1 - 2 * xPlus + 1 - n, xf), 2 * xMinus + n)
                        xTo = Min(PicW - 1, xt)
                        AddToString XOffset:=MskX - CX, _
                                    y:=ty, _
                                    Multiplier:=pMask(MskX, MskY), _
                                    CountSumMul:=True
                    End If
                End If
            End If
        Next MskX
    Next MskY
    
    'store accumulator
    OfcY = y * PicW
    OfcOutY = OutY * OutW
    If TextureMode Then
        GoSub StoreString2
    Else
        If UseCountsWholeString Then
            xFrom = xf
            xTo = xt
            GoSub StoreString1
        Else
            
'            xFrom = xf
'            xTo = xt
'            GoSub StoreString2
            
            xFrom = Max(0, xf)
            xTo = Min(xMinus - 1, xt)
            GoSub StoreString1
            xFrom = Max(xMinus, xf)
            xTo = Min(PicW - 1 - xPlus, xt)
            GoSub StoreString2
            xFrom = Max(Max(PicW - 1 - xPlus + 1, xf), xMinus)
            xTo = Min(PicW - 1, xt)
            GoSub StoreString1
        End If
    End If
    ShowProgress (OutY + 1) / OutH
Next y

'finish - deallocate memory and free arrays
UnReferAry AryPtr(Pic)
UnReferAry AryPtr(OutDatav)
Erase ArStringR, ArStringG, ArStringB
Erase ASCR, ASCG, ASCB
'SwapArys AryPtr(InData), AryPtr(OutData)
SwapArys AryPtr(LOutData), AryPtr(OutData)
Erase OutData
ShowProgress 1.01
Exit Sub

StoreString1: 'store the string using the InData in counts string
    For x = xFrom To xTo
        OutX = x - OutBaseX
        t = ArStringB(x) * Brightness / ASCB(x)
        If t < 0 Then
            If Absolute Then t = -t Else t = 0
        End If
        If t > 255 Then t = 255
        OutDatav(OfcOutY + OutX).rgbBlue = t
        
        t = ArStringG(x) * Brightness / ASCG(x)
        If t < 0 Then
            If Absolute Then t = -t Else t = 0
        End If
        If t > 255 Then t = 255
        OutDatav(OfcOutY + OutX).rgbGreen = t
        
        t = ArStringR(x) * Brightness / ASCR(x)
        If t < 0 Then
            If Absolute Then t = -t Else t = 0
        End If
        If t > 255 Then t = 255
        OutDatav(OfcOutY + OutX).rgbRed = t
    Next x
Return

StoreString2: 'store string using only mass
    For x = xFrom To xTo
        OutX = x - OutBaseX
        t = ArStringB(x) * Brightness / MaskMass.rgbBlue
        If t < 0 Then
            If Absolute Then t = -t Else t = 0
        End If
        If t > 255 Then t = 255
        OutDatav(OfcOutY + OutX).rgbBlue = t
        
        t = ArStringG(x) * Brightness / MaskMass.rgbGreen
        If t < 0 Then
            If Absolute Then t = -t Else t = 0
        End If
        If t > 255 Then t = 255
        OutDatav(OfcOutY + OutX).rgbGreen = t
        
        t = ArStringR(x) * Brightness / MaskMass.rgbRed
        If t < 0 Then
            If Absolute Then t = -t Else t = 0
        End If
        If t > 255 Then t = 255
        OutDatav(OfcOutY + OutX).rgbRed = t
    Next x
Return


eh:
    UnReferAry AryPtr(Pic)
    UnReferAry AryPtr(OutDatav)
    Erase ArStringR, ArStringG, ArStringB
    Erase ASCR, ASCG, ASCB
    ErrRaise "vtFilter"
End Sub

Private Sub AddToString(ByVal XOffset As Long, _
                       ByVal y As Long, _
                       ByRef Multiplier As RGBTriLong, _
                       ByVal CountSumMul As Boolean)
Dim MulR As Long, MulG As Long, MulB As Long
Dim AMulR As Long, AMulG As Long, AMulB As Long
Dim xf As Long, xt As Long
MulR = Multiplier.rgbRed
MulG = Multiplier.rgbGreen
MulB = Multiplier.rgbBlue
If DiffMode Then
    AMulR = Abs(MulR)
    AMulG = Abs(MulG)
    AMulB = Abs(MulB)
Else
    AMulR = MulR
    AMulG = MulG
    AMulB = MulB
End If
Dim x As Long
Dim OfcY As Long
Dim xPlus As Long
Dim tmp As Long
If MulR = 0 And MulG = 0 And MulB = 0 Then Exit Sub
OfcY = y * PicW
If TextureMode Then
    XOffset = XOffset Mod PicW
    If XOffset < 0 Then XOffset = XOffset + PicW
    xPlus = OfcY - XOffset + PicW
    xf = xFrom
    xt = Min(XOffset - 1, xTo)
    GoSub DoIt2
'    For x = xFrom To XOffset - 1
'        ArStringB(x) = ArStringB(x) + Pic(xPlus + x).rgbBlue * MulB
'        ArStringG(x) = ArStringG(x) + Pic(xPlus + x).rgbGreen * MulG
'        ArStringR(x) = ArStringR(x) + Pic(xPlus + x).rgbRed * MulR
'    Next x
    xPlus = OfcY - XOffset
    xf = Max(XOffset, xFrom)
    xt = xTo
    GoSub DoIt2
'    For x = XOffset To xTo
'        ArStringB(x) = ArStringB(x) + Pic(xPlus + x).rgbBlue * MulB
'        ArStringG(x).rgbGreen = ArStringG(x) + Pic(xPlus + x).rgbGreen * MulG
'        ArStringR(x).rgbRed = ArStringR(x) + Pic(xPlus + x).rgbRed * MulR
'    Next x
Else
    xPlus = OfcY - XOffset
    xf = Max(XOffset, xFrom)
    xt = Min(XOffset + PicW - 1, xTo)
    If CountSumMul Then
        'Also fill multiplier accumulators
        GoSub DoIt1
'        For x = Max(XOffset, xFrom) To Min(XOffset + PicW - 1, xTo)
'            ArStringB(x) = ArStringB(x) + Pic(xPlus + x).rgbBlue * MulB
'            ArStringG(x) = ArStringG(x) + Pic(xPlus + x).rgbGreen * MulG
'            ArStringR(x) = ArStringR(x) + Pic(xPlus + x).rgbRed * MulR
'
'            ArStringCounts(x).rgbBlue = ArStringCounts(x).rgbBlue + AMulR
'            ArStringCounts(x).rgbGreen = ArStringCounts(x).rgbGreen + AMulG
'            ArStringCounts(x).rgbRed = ArStringCounts(x).rgbRed + AMulB
'        Next x
    Else
        GoSub DoIt2
'        For x = Max(XOffset, xFrom) To Min(XOffset + PicW - 1, xTo)
'            ArStringB(x) = ArStringB(x) + Pic(xPlus + x).rgbBlue * MulB
'            ArStringG(x) = ArStringG(x) + Pic(xPlus + x).rgbGreen * MulG
'            ArStringR(x) = ArStringR(x) + Pic(xPlus + x).rgbRed * MulR
'        Next x
    End If
End If
Exit Sub
DoIt1:
    GoSub DoIt2
    For x = xf To xt
        ASCB(x) = ASCB(x) + AMulR
        ASCG(x) = ASCG(x) + AMulG
        ASCR(x) = ASCR(x) + AMulB
    Next x
'    For X = xf To xt
'        ASCB(X) = ASCB(X) + AMulR
'    Next X
'    For X = xf To xt
'        ASCG(X) = ASCG(X) + AMulG
'    Next X
'    For X = xf To xt
'        ASCR(X) = ASCR(X) + AMulB
'    Next X
Return

DoIt2:
    tmp = xf + xPlus
    For x = xf To xt
        ArStringB(x) = ArStringB(x) + Pic(tmp).rgbBlue * MulB
        ArStringG(x) = ArStringG(x) + Pic(tmp).rgbGreen * MulG
        ArStringR(x) = ArStringR(x) + Pic(tmp).rgbRed * MulR
        tmp = tmp + 1&
    Next x

'    For X = xf To xt
'        ArStringB(X) = ArStringB(X) + Pic(xPlus + X).rgbBlue * MulB
'    Next X
'    For X = xf To xt
'        ArStringG(X) = ArStringG(X) + Pic(xPlus + X).rgbGreen * MulG
'    Next X
'    For X = xf To xt
'        ArStringR(X) = ArStringR(X) + Pic(xPlus + X).rgbRed * MulR
'    Next X
Return
End Sub

Public Function MaskWeight(ByRef Mask() As RGBTriLong, _
                           Optional ByVal FixZeroProblem As Boolean) As RGBTriLong
Dim UBX As Long, UBY As Long
Dim x As Long, y As Long
Dim Res As RGBTriLong
UBX = UBound(Mask, 1)
UBY = UBound(Mask, 2)
If FixZeroProblem Then
    For y = 0 To UBY
        For x = 0 To UBX
            Res.rgbBlue = Res.rgbBlue + Abs(Mask(x, y).rgbBlue)
            Res.rgbGreen = Res.rgbGreen + Abs(Mask(x, y).rgbGreen)
            Res.rgbRed = Res.rgbRed + Abs(Mask(x, y).rgbRed)
        Next x
    Next y
Else
    For y = 0 To UBY
        For x = 0 To UBX
            Res.rgbBlue = Res.rgbBlue + Mask(x, y).rgbBlue
            Res.rgbGreen = Res.rgbGreen + Mask(x, y).rgbGreen
            Res.rgbRed = Res.rgbRed + Mask(x, y).rgbRed
        Next x
    Next y
End If
MaskWeight = Res
End Function

Public Function MaskMassCentre(ByRef Mask() As RGBTriLong) As PointDBL
Dim UBX As Long, UBY As Long
Dim x As Long, y As Long
Dim Mass As Double
Dim dm As Currency
Dim XM As Currency, YM As Currency
UBX = UBound(Mask, 1)
UBY = UBound(Mask, 2)
For y = 0 To UBY
    For x = 0 To UBX
        dm = Abs(Mask(x, y).rgbBlue) + Abs(Mask(x, y).rgbGreen) + Abs(Mask(x, y).rgbRed)
        Mass = Mass + dm
        XM = XM + dm * x
        YM = YM + dm * y
    Next x
Next y
If Mass = 0 Then
    Err.Raise MassCentreError, "MaskMassCentre", "Mask's mass is zero. Cannot calculate centre of mass."
End If
MaskMassCentre.x = CDbl(XM) / Mass
MaskMassCentre.y = CDbl(YM) / Mass
End Function

Private Sub GenerateModedMask(ByRef OutMask() As RGBTriLong, _
                              ByRef Mask As FilterMask, _
                              ByRef MaskMass As RGBTriLong, _
                              ByRef Opts As typFilterOpts)
Dim CX As Long, CY As Long
Dim x As Long, y As Long
Dim UBX As Long, UBY As Long
Dim m As Long
CX = Mask.Center.x
CY = Mask.Center.y
OutMask = Mask.Mask
UBX = UBound(OutMask, 1)
UBY = UBound(OutMask, 2)

'mask mass modification
If Opts.FilterMode = fmNormal Then
    m = Max(MaskMass.rgbBlue, Max(MaskMass.rgbGreen, MaskMass.rgbRed))
    MaskMass.rgbBlue = m
    MaskMass.rgbGreen = m
    MaskMass.rgbRed = m
Else
  If MaskMass.rgbBlue = 0 Then
    MaskMass.rgbBlue = 255& - OutMask(CX, CY).rgbBlue
    OutMask(CX, CY).rgbBlue = 255
  End If
  If MaskMass.rgbGreen = 0 Then
    MaskMass.rgbGreen = 255& - OutMask(CX, CY).rgbGreen
    OutMask(CX, CY).rgbGreen = 255
  End If
  If MaskMass.rgbRed = 0 Then
    MaskMass.rgbRed = 255& - OutMask(CX, CY).rgbRed
    OutMask(CX, CY).rgbRed = 255
  End If
End If

'mask modification
Select Case Opts.FilterMode
    Case eFilterMode.fmAnti
        For y = 0 To UBY
            For x = 0 To UBX
                OutMask(x, y).rgbBlue = -OutMask(x, y).rgbBlue * Opts.Brightness
                OutMask(x, y).rgbGreen = -OutMask(x, y).rgbGreen * Opts.Brightness
                OutMask(x, y).rgbRed = -OutMask(x, y).rgbRed * Opts.Brightness
            Next x
        Next y
        OutMask(CX, CY).rgbBlue = OutMask(CX, CY).rgbBlue + MaskMass.rgbBlue * (1 + Opts.Brightness)
        OutMask(CX, CY).rgbGreen = OutMask(CX, CY).rgbGreen + MaskMass.rgbGreen * (1 + Opts.Brightness)
        OutMask(CX, CY).rgbRed = OutMask(CX, CY).rgbRed + MaskMass.rgbRed * (1 + Opts.Brightness)
        Opts.Brightness = 1
    Case eFilterMode.fmDifference
        OutMask(CX, CY).rgbBlue = OutMask(CX, CY).rgbBlue - MaskMass.rgbBlue
        OutMask(CX, CY).rgbGreen = OutMask(CX, CY).rgbGreen - MaskMass.rgbGreen
        OutMask(CX, CY).rgbRed = OutMask(CX, CY).rgbRed - MaskMass.rgbRed
        Opts.Absolute = True
    Case eFilterMode.fmNormal
End Select
End Sub

Private Function Min(ByVal a As Long, ByVal b As Long) As Long
If a > b Then Min = b Else Min = a
End Function

Private Function Max(ByVal a As Long, ByVal b As Long) As Long
If a > b Then Max = a Else Max = b
End Function


