Attribute VB_Name = "mdlDifferentiate"
Option Explicit

Dim w As Long, h As Long
Dim pTextureMode As Boolean

Private Sub ShowProgressFD(ByVal Value As Double)
MainForm.ShowProgress Value * 100, DoDoEvents:=True
End Sub

Sub vtDifferentiate(ByRef InData() As Long, _
                    ByRef OutData() As Long, _
                    Range As RECT, _
                    ByRef KOut As Double, _
                    ByRef kData As Double, _
                    ByRef Offset As POINTAPI, _
                    ByRef TextureMode As Boolean)
Static tbl() As Double
Static Foo As Boolean
Dim OutW As Long, OutH As Long
Dim bx As Long, by As Long 'base x,y
Dim xf As Long, yf As Long 'range
Dim xt As Long, yt As Long 'range
Dim InDataV() As RGBQUAD
Dim OutDatav() As RGBQUAD
'Dim InDataRGB() As RGBQUAD
'Dim OutDataRGB() As RGBQUAD
Dim InOfc As Long, InModOfc As Long
Dim OutOfc As Long
Dim x As Long, y As Long 'output pos
Dim x1 As Long, y1 As Long 'input pos
Dim r As Long, g As Long, b As Long

Dim dx As Long, dy As Long 'mixing data offset
'Dim ReDx As Long, ReDy As Long

pTextureMode = TextureMode

bx = Range.Left
by = Range.Top

OutW = Range.Right - Range.Left
OutH = Range.Bottom - Range.Top
If OutW * OutH = 0 Then Exit Sub
RedimIfNeeded OutData, OutW, OutH

AryWH AryPtr(InData), w, h
If w * h = 0 Then Exit Sub

dx = Offset.x 'Mod w
dy = Offset.y 'Mod h
If kData = 0 Then
    dx = 0
    dy = 0
End If
'If dx < 0 Then dx = dx + w
'If dy < 0 Then dy = dy + h

If Not Foo Then
    BuildDistMapDbl tbl
    Foo = True
End If

'for non-border areas
xf = MaxMany(Range.Left, Abs(dx), 1) - bx
yf = MaxMany(Range.Top, Abs(dy), 1) - by
xt = MinMany(Range.Right, w - 1, w - Abs(dx)) - 1 - bx
yt = MinMany(Range.Bottom, h - 1, h - Abs(dy)) - 1 - by

On Error GoTo eh
ConstructAry AryPtr(InDataV), VarPtr(InData(0, 0)), 4, w * h
ConstructAry AryPtr(OutDatav), VarPtr(OutData(0, 0)), 4, OutW * OutH

For y = yf To yt
    y1 = y + by 'input y
    InOfc = y1 * w
    InModOfc = (y1 + dy) * w
    OutOfc = y * OutW
    For x = xf To xt
        x1 = x + bx
        b = _
            InDataV(x1 + dx + InModOfc).rgbBlue * kData + _
            tbl(Abs(0& + InDataV(x1 - 1 + InOfc).rgbBlue - InDataV(x1 + 1 + InOfc).rgbBlue), _
                Abs(0& + InDataV(x1 - w + InOfc).rgbBlue - InDataV(x1 + w + InOfc).rgbBlue)) _
                * KOut
        g = _
            InDataV(x1 + dx + InModOfc).rgbGreen * kData + _
            tbl(Abs(0& + InDataV(x1 - 1 + InOfc).rgbGreen - InDataV(x1 + 1 + InOfc).rgbGreen), _
                Abs(0& + InDataV(x1 - w + InOfc).rgbGreen - InDataV(x1 + w + InOfc).rgbGreen)) _
                * KOut
        r = _
            InDataV(x1 + dx + InModOfc).rgbRed * kData + _
            tbl(Abs(0& + InDataV(x1 - 1 + InOfc).rgbRed - InDataV(x1 + 1 + InOfc).rgbRed), _
                Abs(0& + InDataV(x1 - w + InOfc).rgbRed - InDataV(x1 + w + InOfc).rgbRed)) _
                * KOut
        GoSub StoreRGB
        
    Next x
    ShowProgressFD (y - yf) / (Range.Bottom - Range.Top)
Next y

'left border
xf = Max(Range.Left, 0) - bx
yf = Max(Range.Top, 0) - by
xt = Min(Range.Right, Max(Abs(dx), 1)) - 1 - bx
yt = Min(Range.Bottom, h) - 1 - by
GoSub ProcessBorder

'right border
xf = Max(Range.Left, w - Max(1, Abs(dx))) - bx
yf = Max(Range.Top, 0) - by
xt = Min(Range.Right, w) - 1 - bx
yt = Min(Range.Bottom, h) - 1 - by
GoSub ProcessBorder

'top border
xf = Max(Range.Left, 0) - bx
yf = Max(Range.Top, 0) - by
xt = Min(Range.Right, w) - 1 - bx
yt = Min(Range.Bottom, Max(Abs(dy), 1)) - 1 - by
GoSub ProcessBorder

'bottom border
xf = Max(Range.Left, 0) - bx
yf = Max(Range.Top, h - Max(1, Abs(dy))) - by
xt = Min(Range.Right, w) - 1 - bx
yt = Min(Range.Bottom, h) - 1 - by
GoSub ProcessBorder


UnReferAry AryPtr(InDataV)
UnReferAry AryPtr(OutDatav)

ShowProgress 1.01

Exit Sub
eh:
UnReferAry AryPtr(InDataV)
UnReferAry AryPtr(OutDatav)
ErrRaise "vtDifferentiate"

StoreRGB:
    If b And &HFFFFFF00 Then
        If b And &H70000000 Then
            b = 0
        Else
            b = &HFF
        End If
    End If
    OutDatav(x + OutOfc).rgbBlue = b
    
    If g And &HFFFFFF00 Then
        If g And &H70000000 Then
            g = 0
        Else
            g = &HFF
        End If
    End If
    OutDatav(x + OutOfc).rgbGreen = g
    
    If r And &HFFFFFF00 Then
        If r And &H70000000 Then
            r = 0
        Else
            r = &HFF
        End If
    End If
    OutDatav(x + OutOfc).rgbRed = r
    
Return

ProcessBorder:
    Dim xd As Long, yd As Long
    Dim xLeft As Long, XRight As Long
    Dim YAbove As Long, YBelow As Long
    Dim PosData As Long
    Dim PosAbove As Long, PosBelow As Long
    Dim PosLeft As Long, PosRight As Long
    For y = yf To yt
        For x = xf To xt
            OutOfc = y * OutW
            
            y1 = y + by 'input y
            x1 = x + bx
            
            xd = x1 + dx
            yd = y1 + dy
            Clip xd, yd
            PosData = xd + yd * w
            
            
            xLeft = x1 - 1
            YAbove = y1 - 1
            Clip xLeft, YAbove
            
            PosLeft = xLeft + y1 * w
            PosAbove = x1 + YAbove * w
            
            XRight = x1 + 1
            YBelow = y1 + 1
            Clip XRight, YBelow
            
            PosRight = XRight + y1 * w
            PosBelow = x1 + YBelow * w
            
            
            
            b = _
                InDataV(PosData).rgbBlue * kData + _
                tbl(Abs(0& + InDataV(PosAbove).rgbBlue - InDataV(PosBelow).rgbBlue), _
                    Abs(0& + InDataV(PosLeft).rgbBlue - InDataV(PosRight).rgbBlue)) _
                    * KOut
            g = _
                InDataV(PosData).rgbGreen * kData + _
                tbl(Abs(0& + InDataV(PosAbove).rgbGreen - InDataV(PosBelow).rgbGreen), _
                    Abs(0& + InDataV(PosLeft).rgbGreen - InDataV(PosRight).rgbGreen)) _
                    * KOut
            r = _
                InDataV(PosData).rgbRed * kData + _
                tbl(Abs(0& + InDataV(PosAbove).rgbRed - InDataV(PosBelow).rgbRed), _
                    Abs(0& + InDataV(PosLeft).rgbRed - InDataV(PosRight).rgbRed)) _
                    * KOut
            GoSub StoreRGB
            
        Next x
    Next y
Return

End Sub

Private Function Clip(ByRef x As Long, _
                      ByRef y As Long)
If pTextureMode Then
    x = x Mod w
    If x < 0 Then x = x + w
    y = y Mod h
    If y < 0 Then y = y + h
Else
    If x < 0 Then x = 0
    If y < 0 Then y = 0
    If x > w - 1 Then x = w - 1
    If y > h - 1 Then y = h - 1
End If
End Function


Private Sub BuildDistMapDbl(ByRef Map() As Double)
Dim v As Long, h As Long
Dim o As Double
ReDim Map(0 To 255, 0 To 255)
For v = 0 To 255
    For h = 0 To 255
        o = Sqr(v * v + h * h)
        'If o > 255 Then o = 255
        Map(h, v) = o
    Next h
Next v
End Sub


