Attribute VB_Name = "DrawingEngine"
Option Explicit


Public AntiAliasingSharpness As Double
'1 for 1-pixel antialiasing,
'large for no antialiasing,
'small for smoothing.
'(zero and negative not allowed)

Type vtVertex
    x As Double
    y As Double
    Color As Long
    Weight As Double
End Type
Type vtVertex3opaq
    x As Double
    y As Double
    Color As Long
    Weight As Double
    opaq As Long
End Type


Type FadeDescCalculatableVals
    FCount As Double
    Power As Double
    Offset As Double
    CntxPi As Double
    EqPower As Double
End Type

Type FadeDesc
    FCount As Double
    Power As Double
    Offset As Double
    Mode As FMode
    AutoColor1 As Boolean
    AutoColor2 As Boolean
    'Width As Single
    CalculatedVals As FadeDescCalculatableVals
    'Prog As SMP
End Type

Public Enum FMode
    dbFSine = 0
    dbFLinear = 1
    dbFProg = -1
End Enum

Type AlphaPixel
    x As Integer
    y As Integer
    rgbRed As Byte
    rgbGreen As Byte
    rgbBlue As Byte
    rgbOpacity As Byte '0 does not paint, 255 paints
    drawOpacity As Long
End Type

Public Type PixelsStack
    Pixels() As AlphaPixel
    nPixels As Long
End Type

Public Type ComplexPixels
    Elements() As PixelsStack
End Type

Public Type UndoPixel
    x As Integer
    y As Integer
    Color As RGBQUAD
End Type

Public Type UndoElement
    Pixels() As UndoPixel
    nPixels As Long
End Type

'Calculates pixels of a line.
'Starting from Point1 with corresponding weight,
'  ending in point 2 with the other weight.
'

'nMem sets the init number of pixels and returns the end number of pixels
Public Sub pntGradientLineHQ(Point1 As vtVertex, _
                             Point2 As vtVertex, _
                             FadeDsc As FadeDesc, _
                             ByRef Pixels() As AlphaPixel, _
                             ByRef nMem As Long)
'Begin geometry declarations
    Dim dx As Double, dy As Double 'line vector
    Dim Pos As Double 'zero at start, 1 at end; normalized x
    Dim LineLength As Double
    Dim DirX As Double, DirY As Double 'normalized dx,dy
    Dim tmpX As Double, tmpY As Double 'overnormalized dx,dy
    Dim X1 As Double, Y1 As Double 'start point
    Dim X2 As Double, Y2 As Double 'end point
    Dim vx As Double, vy As Double 'current vector (from start point)
    Dim Dist As Double 'distance to baseline (y)
    Dim d As Double 'distance to the figure
    Dim Wgt As Double 'current half-weight
    Dim hw1 As Double, hw2 As Double 'start/end half-weights
    
    'Dim lPos1 As Double, lPos2 As Double
    Dim lhw1 As Double, lhw2 As Double 'trapezoid height at start/end
    Dim dlhw As Double '= lhw2-lhw1
    'Dim Delta1 As Double, Delta2 As Double
    Dim si As Double, co As Double 'sin and cos of border angle
    Dim ta As Double '=si/co
    
    Dim AAR As Double 'anti-aliasing radius
'end geometry declarations

'begin color declaration
    Dim r As Long, g As Long, b As Long, o As Long, a As Long 'output color
    Dim ddr As Double, ddg As Double, ddb As Double, ddo As Double 'deltas
    Dim r1 As Long, g1 As Long, b1 As Long, o1 As Long 'of first pixel
    Dim rgb1 As RGBQUAD, rgb2 As RGBQUAD 'for conversion
    Dim dj As Double
'end color declaration

'begin walker declarations
    Dim x As Long, y As Long 'position to process
    Dim SeedX As Long, SeedY As Long
    Dim yTo As Long
    Dim nPixels As Long
    Dim cnt As Long 'pixel counter
    Dim lcnt As Long 'cnt before the line processed
    Dim xFrom As Long, xTo As Long
    Dim OnlyStart As Boolean, OnlyEnd As Boolean  'whether to draw only start/end
    Dim PixelPainted As Boolean
'end walker declarations

'begin color initialization
    CopyMemory rgb1, Point1.Color, 4
    CopyMemory rgb2, Point2.Color, 4
    
    r1 = rgb1.rgbRed
    g1 = rgb1.rgbGreen
    b1 = rgb1.rgbBlue
    o1 = rgb1.rgbReserved
    
    ddr = rgb2.rgbRed - r1
    ddg = rgb2.rgbGreen - g1
    ddb = rgb2.rgbBlue - b1
    ddo = rgb2.rgbReserved - o1
'end color initialization

'begin geometry initialization
    hw1 = Point1.Weight * 0.5
    hw2 = Point2.Weight * 0.5
    AAR = 1 / AntiAliasingSharpness
    
    X1 = Point1.x
    Y1 = Point1.y
    X2 = Point2.x
    Y2 = Point2.y
    
    dx = X2 - X1
    dy = Y2 - Y1
    
    LineLength = Sqr(dx * dx + dy * dy)
    If LineLength > 0.0000000001 Then
        DirX = dx / LineLength
        DirY = dy / LineLength
        tmpX = DirX / LineLength
        tmpY = DirY / LineLength
        dj = 1 / LineLength
    End If
    
    If LineLength > 0.0000000001 Then
        si = (hw1 - hw2) / LineLength
    Else
        If hw1 > hw2 Then
            si = 10000000000#
        Else
            si = -10000000000#
        End If
    End If
    
'    lx1 = x1
'    ly1 = y1
'    lx2 = x2
'    ly2 = y2
    If si > (1 - 0.0000000001) Then
        lhw1 = hw1
        lhw2 = hw1
        OnlyStart = True
    '    Delta1 = hw1 + 3 + AAR
    '    Delta2 = hw1 + 3 + AAR
    ElseIf si < -(1 - 0.0000000001) Then
        'lx1 = x2
        'ly1 = y2
        'lx2 = x2
        'ly2 = y2
        lhw1 = hw2
        lhw2 = hw2
        OnlyEnd = True
    '    Delta1 = -(hw2 + 3 + AAR)
    '    Delta2 = -(hw2 + 3 + AAR)
    Else
        co = Sqr(1 - si * si)
        ta = si / co
        lhw1 = hw1 / co
        lhw2 = hw2 / co
    '    Delta1 = hw1 * si
    '    Delta2 = hw2 * si
    End If
    
    dlhw = lhw2 - lhw1
    
    If LineLength > 0.0000000001 Then
        'ltmpX = DirX / lLineLength
        'ltmpY = DirY / lLineLength
        'lPos1 = Delta1 / LineLength
        'lPos2 = 1 + Delta2 / LineLength
        'lPosK = 1 / Co ^ 2
    End If
'end geometry initialization

'begin memory allocation
    'nPixels = -Int(-((hw1 + hw2) / 2 * LineLength + 2 * hw1 ^ 2 + 2 * hw2 ^ 2 + 8) * 1.7)
    'ReDim Pixels(0 To nPixels - 1)
    cnt = 0
'end memory allocation

'begin walker
    'begin walker initialization
        If Y1 - hw1 < Y2 - hw2 Then
            SeedY = Y1 - (hw1 + 1 + 1 / AntiAliasingSharpness)
            SeedX = X1
        Else
            SeedY = Y2 - (hw2 + 1 + 1 / AntiAliasingSharpness)
            SeedX = X2
        End If
        xFrom = SeedX - 1
        xTo = SeedX + 1
        y = SeedY
    'end walker initialization
    
    'begin walker loops
        Do 'y
            lcnt = cnt
            x = SeedX
            For x = xFrom To xTo
                GoSub PutPixel
                If PixelPainted Then
                    SeedX = x
                    Exit For
                End If
            Next x
            If PixelPainted Then
                Do
                    x = x - 1
                    GoSub PutPixel
                Loop While PixelPainted
                xFrom = x
                x = SeedX
                Do
                    x = x + 1
                    GoSub PutPixel
                Loop While PixelPainted
                xTo = x
            End If
            y = y + 1&
        Loop Until cnt = lcnt And y > SeedY + 2 + AAR
    'end walker loops
    
'end walker

'begin memory cleanup
    If cnt = 0 Then
        Erase Pixels
    Else
        ReDim Preserve Pixels(0 To cnt - 1)
    End If
    nMem = cnt
'end memory cleanup

Exit Sub

PutPixel: '(x,y)
    vx = x - X1
    vy = y - Y1
    Pos = vx * tmpX + vy * tmpY
    Dist = Abs(vx * DirY - vy * DirX)
    'lPos = (Pos - lPosMinus) * lPosK
    If (Pos * LineLength <= Dist * ta Or OnlyStart) And Not OnlyEnd Then
        Wgt = hw1
        d = (hw1 - Sqr(vx * vx + vy * vy))
    ElseIf (LineLength - Pos * LineLength <= Dist * -ta Or OnlyEnd) And Not OnlyStart Then
        Wgt = hw2
        d = (hw2 - Sqr((X2 - x) * (X2 - x) + (Y2 - y) * (Y2 - y)))
    Else
        Wgt = Pos * dlhw + lhw1
        d = (Wgt - Dist) * co
        'If d > 0.5 Then d = 0.5
    End If
        'If Wgt < 1 / 256 Then Wgt = 1 / 256
    PixelPainted = d > -1#
    d = d * AntiAliasingSharpness + 0.5
    'If d > 2# * Wgt Then d = 2# * Wgt
    'If d > 1# Then d = 1#
    
    a = 255# * d
    PixelPainted = PixelPainted Or a > 0
    If a And &HFFFFFF00 Then 'byte overflow
        If a And &H80000000 Then 'a<0
            a = 0&
        Else
            a = 255&
        End If
    End If
    If PixelPainted Then
        Pos = CountJEx(Pos, FadeDsc, dj)
        r = r1 + CLng(ddr * Pos)
        g = g1 + CLng(ddg * Pos)
        b = b1 + CLng(ddb * Pos)
        o = o1 + CLng(ddo * Pos)
        If r And &HFFFFFF00 Then
            If r And &H80000000 Then
                r = 0&
            Else
                r = 255&
            End If
        End If
        
        If g And &HFFFFFF00 Then
            If g And &H80000000 Then
                g = 0&
            Else
                g = 255&
            End If
        End If
        
        If b And &HFFFFFF00 Then
            If b And &H80000000 Then
                b = 0&
            Else
                b = 255&
            End If
        End If
        
        If o And &HFFFFFF00 Then 'byte overflow
            If o And &H80000000 Then 'a<0
                o = 0&
            Else
                o = 255&
            End If
        End If
        
        If cnt >= nMem Then
            If nMem <= 0 Then
                nMem = 1000
                ReDim Pixels(0 To 0)
            Else
                nMem = nMem * 2
            End If
            ReDim Preserve Pixels(0 To nMem - 1)
        End If
        
        Pixels(cnt).x = x
        Pixels(cnt).y = y
        Pixels(cnt).rgbBlue = b
        Pixels(cnt).rgbGreen = g
        Pixels(cnt).rgbRed = r
        Pixels(cnt).rgbOpacity = o
        Pixels(cnt).drawOpacity = a
        
        cnt = cnt + 1
        If cnt Mod 10000 = 0 Then
          If BreakKeyPressed Then
              Err.Raise dbCWS
          End If
        End If
    End If
Return 'from PutPixel:
End Sub
'
'Private Sub SwapLongs(ByRef a As Long, ByRef b As Long)
'Dim tmp As Long
'tmp = a
'a = b
'b = tmp
'End Sub


Public Function CountJ(ByVal j As Double, _
                       ByVal CountPereliv As Single, _
                       ByVal Stepen As Double, _
                       ByVal Offset As Single, _
                       ByVal Mode As Integer)
Dim tmp As Double
Static lSt As Double, lStCnt As Double
If lSt <> Stepen Then
    If Stepen = 0 Then
        'lStCnt = 0
        CountJ = 1
        Exit Function
    ElseIf 1# - Stepen < 0.00001 Then
        'lStCnt = Log(0.5) / Log(0.99999)
        CountJ = 0
        Exit Function
    Else
        lStCnt = Log(0.5) / Log(Stepen)
    End If
    lSt = Stepen
End If

Select Case Mode
    Case 0
        CountJ = (-0.5 * Cos((j - Offset) * Pi * CountPereliv) + 0.5) ^ lStCnt
    Case 1
        tmp = (j - Offset) * CountPereliv
        tmp = tmp - Int(tmp * 0.5) * 2
        If tmp > 1 Then tmp = 2 - tmp
        CountJ = tmp ^ lStCnt
End Select
End Function


Public Function CountJEx(ByVal j As Double, _
                         ByRef FadeDsc As FadeDesc, _
                         Optional ByVal dj As Double = 0) As Double
Dim tmp As Double
'Static lSt As Double, lStCnt As Double
Dim Stepen As Double
Dim kfc As Double
Dim Result As Double, MD As Double

Stepen = FadeDsc.Power
If FadeDsc.CalculatedVals.Power <> Stepen Then
    If Stepen = 0 Then
        'lStCnt = 0
        CountJEx = 1
        Exit Function
    ElseIf 1 - Stepen < 0.00001 Then
        'lStCnt = Log(0.5) / Log(0.99999)
        CountJEx = 0
        Exit Function
    Else
        FadeDsc.CalculatedVals.EqPower = Log(0.5) / Log(Stepen)
    End If
End If
If dj < 0.00001 Then dj = 0.00001
j = j - FadeDsc.Offset
Select Case FadeDsc.Mode
    Case FMode.dbFLinear
'        dj = (dj * FadeDsc.FCount) ^ 5
        If FadeDsc.FCount = 1 And FadeDsc.Offset = 0 Then
          CountJEx = j
        Else
          tmp = (j - FadeDsc.Offset) * FadeDsc.FCount
          tmp = tmp - Int(tmp * 0.5) * 2
          If tmp > 1 Then tmp = 2 - tmp
          CountJEx = tmp ^ FadeDsc.CalculatedVals.EqPower
        End If
'        Md = 1 / (1 + FadeDsc.Power)
'        CountJEx = Result + (Md - Result) * dj / (1 + dj)
'        j1 = j * FadeDsc.FCount / 2
'        j2 = (j + dj) * FadeDsc.FCount / 2
'        Intj1 = Int(j1)
'        Intj2 = Int(j2)
'        jmod1 = j1 - Intj1
'        jmod2 = j2 - Intj2
'        y1 = Abs(jmod1 - 0.5) * 2
'        y2 = Abs(jmod2 - 0.5) * 2
'        If jmod1 > 0.5 Then
'            i1 = Intj1 + 0.5 + (1 - jmod1) * y1
'        Else
'            i1 = Intj1 + jmod1 * y1
'        End If
'        If jmod2 > 0.5 Then
'            i2 = Intj2 + 0.5 + (1 - jmod2) * y2
'        Else
'            i2 = Intj2 + jmod2 * y2
'        End If
'        CountJEx = ((i2 - i1) / dj) ^ FadeDsc.CalculatedVals.EqPower
        
    
    Case FMode.dbFSine
        
        'CountJEx = (-0.5 * Cos((j - FadeDsc.Offset) * Pi * FadeDsc.FCount) + 0.5) ^ lStCnt
        kfc = FadeDsc.FCount * Pi * 0.5
        CountJEx = (Sin(j * kfc) ^ 2) ^ FadeDsc.CalculatedVals.EqPower
'        If kfc < 1E-20 Then
'            CountJEx = 0
'        Else
'            CountJEx = (0.5 - Sin(kfc * dj * 0.5) * Cos((j + dj * 0.5) * kfc) / kfc / dj)
'        End If
'    Case FMode.dbFProg
'        Static EV As New clsEVal
'        With FadeDsc.Prog
'            .Vars(0).Value = (j - FadeDsc.Offset) * FadeDsc.FCount 'vPos
'            .Vars(1).Value = j 'rPos
'
'            .Vars(2).Value = FadeDsc.Offset 'offset
'
'            .Vars(3).Value = FadeDsc.FCount 'count
'
'            .Vars(4).Value = FadeDsc.Power 'Degree
'            .Vars(5).Value = lStCnt 'Power
'        End With
'        On Error GoTo eh
'        tmp = EV.ExecuteSMP(FadeDsc.Prog)
'        If tmp < 0 Then tmp = 0
'        If tmp > 1 Then tmp = 1
'        CountJEx = tmp ^ lStCnt
        
End Select
Exit Function
eh:
PushError
ShowStatus Err.Description, HoldTime:=5
PopError
ErrRaise "CountJ"
End Function

'no range-checking
Public Sub DrawPixels(ByRef Data() As RGBQUAD, _
                       ByRef Pixels() As AlphaPixel, _
                       ByVal nPixels As Long, _
                       Optional ByVal AryPtrUndo As Long = 0, _
                       Optional ByRef nUndoPixels As Long)
Dim i As Long
Dim x As Long, y As Long
Dim UndoAry() As UndoPixel
Dim iUndo As Long
If nPixels <= 0 Then Exit Sub
If AryPtrUndo <> 0 Then
    SwapArys AryPtr(UndoAry), AryPtrUndo
    If nUndoPixels = 0 Then
        ReDim UndoAry(0 To nUndoPixels + nPixels - 1)
    Else
        ReDim Preserve UndoAry(0 To nUndoPixels + nPixels - 1)
    End If
    On Error GoTo eh
    For i = 0 To nPixels - 1
        x = Pixels(i).x
        y = Pixels(i).y
        UndoAry(i + nUndoPixels).Color = Data(x, y)
        Data(x, y).rgbBlue = Data(x, y).rgbBlue + (CLng(Pixels(i).rgbBlue) - Data(x, y).rgbBlue) * Pixels(i).drawOpacity \ 255
        Data(x, y).rgbGreen = Data(x, y).rgbGreen + (CLng(Pixels(i).rgbGreen) - Data(x, y).rgbGreen) * Pixels(i).drawOpacity \ 255
        Data(x, y).rgbRed = Data(x, y).rgbRed + (CLng(Pixels(i).rgbRed) - Data(x, y).rgbRed) * Pixels(i).drawOpacity \ 255
        Data(x, y).rgbReserved = Data(x, y).rgbReserved + (CLng(Pixels(i).rgbOpacity) - Data(x, y).rgbReserved) * Pixels(i).drawOpacity \ 255
    Next i
    nUndoPixels = nUndoPixels + 1
    SwapArys AryPtr(UndoAry), AryPtrUndo
Else
    For i = 0 To nPixels - 1
        x = Pixels(i).x
        y = Pixels(i).y
        Data(x, y).rgbBlue = Data(x, y).rgbBlue + (CLng(Pixels(i).rgbBlue) - Data(x, y).rgbBlue) * Pixels(i).drawOpacity \ 255
        Data(x, y).rgbGreen = Data(x, y).rgbGreen + (CLng(Pixels(i).rgbGreen) - Data(x, y).rgbGreen) * Pixels(i).drawOpacity \ 255
        Data(x, y).rgbRed = Data(x, y).rgbRed + (CLng(Pixels(i).rgbRed) - Data(x, y).rgbRed) * Pixels(i).drawOpacity \ 255
        Data(x, y).rgbReserved = Data(x, y).rgbReserved + (CLng(Pixels(i).rgbOpacity) - Data(x, y).rgbReserved) * Pixels(i).drawOpacity \ 255
    Next i
End If
Exit Sub
eh:
If AryPtrUndo <> 0 Then
    PushError
    SwapArys AryPtr(UndoAry), AryPtrUndo
    nUndoPixels = i + 1
    PopError
End If
ErrRaise "DrawPixels"
End Sub

'Private Sub ApplyUndo(ByRef Data() As RGBQUAD, _
'                      ByRef UndoPixels() As UndoPixel, _
'                      ByRef nPixels As Long, _
'                      Optional ByVal ptrAryRedo As Long, _
'                      Optional ByRef nPixelsInRedo As Long)
'Dim i As Long
'Dim iRedo As Long
'Dim RedoAry() As UndoPixel
'If nPixels = 0 Then Exit Sub
'If ptrAryRedo <> 0 Then
'    SwapArys
'Else
'End If
'End Sub

Public Sub RangeCheckComplexPixels(ByVal w As Long, _
                                   ByVal h As Long, _
                                   ByRef InCPixels As ComplexPixels)
Dim iElement As Long
Dim nElems As Long
Dim n As Long
Dim i As Long
Dim cnt As Long
Dim AryWriting() As AlphaPixel
Dim x As Long, y As Long
If AryDims(AryPtr(InCPixels.Elements)) = 0 Then Exit Sub
nElems = UBound(InCPixels.Elements) + 1
For iElement = 0 To nElems - 1
    With InCPixels.Elements(iElement)
        n = .nPixels
        If n > 0 Then
            ReDim AryWriting(0 To n - 1)
            cnt = 0
            For i = 0 To n - 1
                x = .Pixels(i).x
                y = .Pixels(i).y
                If x >= 0 And y >= 0 And x < w And y < h Then
                    AryWriting(cnt) = .Pixels(i)
                    cnt = cnt + 1&
                End If
            Next i
            If cnt < n Then
                If cnt = 0 Then
                    Erase AryWriting
                Else
                    ReDim Preserve AryWriting(0 To cnt - 1)
                End If
                SwapArys AryPtr(AryWriting), AryPtr(.Pixels)
                .nPixels = cnt
            End If
        End If
    End With
Next iElement
End Sub

Public Sub DrawComplexPixelsWU(ByRef Data() As Long, _
                               ByRef Pixels As ComplexPixels, _
                               ByRef UndoCP As ComplexPixels, _
                               Optional ByVal RangeCheck As Boolean = False)
Dim RGBData() As RGBQUAD
Dim w As Long, h As Long
Dim iElement As Long
Dim nElems As Long

AryWH AryPtr(Data), w, h

If RangeCheck Then
    RangeCheckComplexPixels w, h, Pixels
End If

nElems = AryLen(AryPtr(Pixels.Elements))
If nElems > 0 Then
    On Error GoTo eh
    ReferAry AryPtr(RGBData), AryPtr(Data)
    For iElement = 0 To nElems - 1
        DrawPixels RGBData, Pixels.Elements(iElement).Pixels, Pixels.Elements(iElement).nPixels
    Next iElement
    UnReferAry AryPtr(RGBData)
End If
Exit Sub
Resume
eh:
UnReferAry AryPtr(RGBData)
ErrRaise "DrawComplexPixels"
End Sub

'Calculates pixels of a rotated rectangle.
'Rect is centered at pntCenter and rotated around it by Angle degrees CCW.
'

'nMem sets the init number of pixels and returns the end number of pixels
Public Sub pntRectangle(pntCenter As vtVertex, _
                             ByVal rW As Double, ByVal rH As Double, _
                             ByVal AngleDeg As Double, _
                             ByRef Pixels() As AlphaPixel, _
                             ByRef nMem As Long)

'begin color initialization
    Dim r As Long, g As Long, b As Long, o As Long, a As Long 'output color
    Dim rgb1 As RGBQUAD, rgb2 As RGBQUAD 'for conversion
    
    CopyMemory rgb1, pntCenter.Color, 4
    
    r = rgb1.rgbRed
    g = rgb1.rgbGreen
    b = rgb1.rgbBlue
    o = rgb1.rgbReserved
    
'end color initialization

'begin walker declarations
    Dim x As Long, y As Long 'position to process
    Dim SeedX As Long, SeedY As Long
    Dim yTo As Long
    Dim nPixels As Long
    Dim cnt As Long 'pixel counter
    Dim lcnt As Long 'cnt before the line processed
    Dim xFrom As Long, xTo As Long
    Dim OnlyStart As Boolean, OnlyEnd As Boolean  'whether to draw only start/end
    Dim PixelPainted As Boolean
'end walker declarations

'begin geometry initialization
    Dim AAR As Double 'anti-aliasing radius (aa'ed border is twice as wide)
    AAR = 0.5 / AntiAliasingSharpness
    
    Dim cx As Double, cy As Double
    cx = pntCenter.x: cy = pntCenter.y
    
    Dim dx As Double, dy As Double 'unity orientation vector (along w)
    Dim AngleRad As Double
    AngleRad = AngleDeg / 180 * Pi 'converted to radians
    dx = Cos(AngleRad): dy = Sin(AngleRad)
    
    Dim wdx As Double, wdy As Double 'unity vector along w edges
    wdx = dx
    wdy = dy
    
    Dim hdx As Double, hdy As Double 'unity vector along h edges
    hdx = -dy
    hdy = dx
    
    Dim hw As Double, hh As Double 'half-width and half-height
    hw = rW / 2: hh = rH / 2
    If rW <= 0 Or rH <= 0 Then Exit Sub 'nothing to paint
    
    'corner numbering: 1-2-3-4 stand for lefttop, righttop, leftbot, rightbot
    '1--2
    '|  |
    '3--4
    'these are coordinates of outermost pixels to be painted
    '(including blurring due to antialiasing) with respect to
    'rect center
    Dim X1 As Double, Y1 As Double
    Dim X2 As Double, Y2 As Double
    Dim X3 As Double, Y3 As Double
    Dim X4 As Double, Y4 As Double
    X1 = wdx * -(hw + AAR) + hdx * -(hh + AAR)
    Y1 = wdy * -(hw + AAR) + hdy * -(hh + AAR)
    X2 = wdx * (hw + AAR) + hdx * -(hh + AAR)
    Y2 = wdy * (hw + AAR) + hdy * -(hh + AAR)
    X3 = wdx * -(hw + AAR) + hdx * (hh + AAR)
    Y3 = wdy * -(hw + AAR) + hdy * (hh + AAR)
    X4 = wdx * (hw + AAR) + hdx * (hh + AAR)
    Y4 = wdy * (hw + AAR) + hdy * (hh + AAR)
    
    'calculate where to seed
    Dim ymin As Double, xofmin As Double
    ymin = Y1
    xofmin = X1
    If Y2 < ymin Then
      ymin = Y2
      xofmin = X2
    End If
    If Y3 < ymin Then
      ymin = Y3
      xofmin = X3
    End If
    If Y4 < ymin Then
      ymin = Y4
      xofmin = X4
    End If
    SeedX = cx + xofmin
    SeedY = Int(cy + ymin)
    
    Dim vx As Double, vy As Double 'current pixel position relative to the center
    Dim DistW As Double 'distance from center along W
    Dim DistH As Double 'distance from center along H
    Dim d1 As Double, d2 As Double  'a variable to hold opacity (valid range 0..1)
'end geometry initialization


'begin walker
    'begin walker initialization
        'cnt = 0
        cnt = nMem
        xFrom = SeedX - 1
        xTo = SeedX + 1
        y = SeedY
    'end walker initialization
    
    'begin walker loops
        Do 'instead of for-next loop by y
            lcnt = cnt 'save current pixel number to later compare it
                       'with cnt and find out if any pixels were painted
            'find the first painted pixel in the scan line.
            'that position is saved in seedx
            For x = xFrom To xTo
                GoSub PutPixel
                If PixelPainted Then
                    SeedX = x
                    Exit For
                End If
            Next x
            If PixelPainted Then
                'paint all pixels to the left
                Do
                    x = x - 1
                    GoSub PutPixel
                Loop While PixelPainted
                xFrom = x 'and save the leftmost painted pixel as search bound for next scan line
                'paint all pixels to the right
                x = SeedX
                Do
                    x = x + 1
                    GoSub PutPixel
                Loop While PixelPainted
                xTo = x 'and save the rightmost painted pixel as the other search bound for next scan line
            End If
            y = y + 1&
        Loop Until cnt = lcnt And y > SeedY + 2 + AAR
        'terminate loop if no pixels were painted in last scan line.
        'The second condition prevents walker termination
        'on first scan lines because of no painted pixels
    'end walker loops
    
'end walker

'begin memory cleanup
    If cnt = 0 Then
        Erase Pixels
    Else
        ReDim Preserve Pixels(0 To cnt - 1)
    End If
    nMem = cnt
'end memory cleanup

Exit Sub

PutPixel: '(x,y)
    vx = x - cx
    vy = y - cy
    DistW = Abs(vx * wdx + vy * wdy) 'distance from center along W
    DistH = Abs(vx * hdx + vy * hdy) 'distance from center along H
    
    'the following is accurate only if both w and h are >1
    d1 = (hw - DistW) * AntiAliasingSharpness + 0.5 'opacity dictated by H-edges
    If d1 <= 0# Then
      PixelPainted = False
      Return 'from PutPixel:
    End If
    If d1 > 1# Then d1 = 1#
    d2 = (hh - DistH) * AntiAliasingSharpness + 0.5 'opacity dictated by H-edges
    If d2 <= 0# Then
      PixelPainted = False
      Return 'from PutPixel:
    End If
    If d2 > 1# Then d2 = 1#
    PixelPainted = True 'though (long)a can be zero, paint the pixel
    
    'If d > 1 Then d = 1 '(done with longs - faster! will work with distances up to 8 million pix)
    
    a = 255# * d1 * d2
    If a And &HFFFFFF00 Then 'd>1, i.e. a>255
      a = 255&
    End If
    'allocate extra mem if running out of space
    If cnt >= nMem Then
        If nMem <= 0 Then
            nMem = 1000
            ReDim Pixels(0 To 0)
        Else
            nMem = nMem * 2
        End If
        ReDim Preserve Pixels(0 To nMem - 1)
    End If
    
    Pixels(cnt).x = x
    Pixels(cnt).y = y
    Pixels(cnt).rgbBlue = b
    Pixels(cnt).rgbGreen = g
    Pixels(cnt).rgbRed = r
    Pixels(cnt).rgbOpacity = o
    Pixels(cnt).drawOpacity = a
    
    cnt = cnt + 1
    If cnt Mod 10000 = 0 Then
      If BreakKeyPressed Then
          Err.Raise dbCWS
      End If
    End If
Return 'from PutPixel:
End Sub


