Attribute VB_Name = "mdlWaves"
Option Explicit


Type WavePos 'identical to PointAPI
    X As Long
    Y As Long
End Type

Type typWaveSource
    Pos As WavePos
    WaveLength As Long
    Strength As Double 'the charge
    Color As RGBQUAD 'the color. Defines r/g/b charges or line color
    Selected As Boolean  'Used in customize dialog
    Padding As Integer
End Type

Public Sub ShowProgress(ByVal Part As Double, Optional ByVal DoDoEvents As Boolean)
MainForm.ShowProgress Part * 100, DoDoEvents
End Sub

Public Function GetPi() As Double
GetPi = 4# * Atn(1#)
End Function

Public Sub DrawWaves(ByRef Data() As Long, _
                     ByRef WS() As typWaveSource, _
                     FallDownFactor As Double, _
                     ByVal Absolute As Boolean)
Dim rgbData() As RGBQUAD
Dim w As Long, h As Long
Dim X As Long, Y As Long
Dim Reds() As Double, Greens() As Double, Blues() As Double
Dim xs() As Double, ys() As Double
Dim WLs() As Double ' pi*2/wl
Dim WSUB As Long
Dim OfcY As Long
Dim Dist2 As Double, Dist As Double
Dim i As Long
Dim r As Double, g As Double, B As Double
Dim rl As Long, gl As Long, bl As Long
Dim Mul As Double
Dim Pi As Double
Pi = GetPi
On Error GoTo eh

If AryDims(AryPtr(WS)) <> 1 Then
    If AryDims(AryPtr(WS)) = 0 Then
        Err.Raise dbCWS
    Else
        Err.Raise 12321, , "A one-dimensional array is required (internal error)."
    End If
End If
WSUB = UBound(WS)
ShowProgress 0
ReDim Reds(0 To WSUB), Greens(0 To WSUB), Blues(0 To WSUB)
ReDim xs(0 To WSUB), ys(0 To WSUB)
ReDim WLs(0 To WSUB)
For i = 0 To WSUB
    Reds(i) = WS(i).Color.rgbRed * WS(i).Strength / 16#
    Greens(i) = WS(i).Color.rgbGreen * WS(i).Strength / 16#
    Blues(i) = WS(i).Color.rgbBlue * WS(i).Strength / 16#
    xs(i) = WS(i).Pos.X
    ys(i) = WS(i).Pos.Y
    WLs(i) = 2# * Pi / WS(i).WaveLength
Next i

TestDims Data, 2
w = UBound(Data, 1) + 1
h = UBound(Data, 2) + 1

ConstructAry AryPtr(rgbData), VarPtr(Data(0, 0)), 4, w * h

For Y = 0 To h - 1
    OfcY = Y * w
    For X = 0 To w - 1
        r = 0#
        g = 0#
        B = 0#
        For i = 0 To WSUB
            Dist2 = (X - xs(i)) * (X - xs(i)) + (Y - ys(i)) * (Y - ys(i))
            Dist = Sqr(Dist2)
            Mul = Cos(Dist * WLs(i)) * Exp(-Dist * FallDownFactor)
            r = r + Mul * Reds(i)
            g = g + Mul * Greens(i)
            B = B + Mul * Blues(i)
        Next i
        
        If Absolute Then
            rl = Abs(r)
            gl = Abs(g)
            bl = Abs(B)
        Else
            rl = CLng(r) + 128&
            gl = CLng(g) + 128&
            bl = CLng(B) + 128&
        End If
        
        If rl < 0& Then rl = 0& Else If rl > 255& Then rl = 255&
        rgbData(OfcY + X).rgbRed = rl
        If gl < 0& Then gl = 0& Else If gl > 255& Then gl = 255&
        rgbData(OfcY + X).rgbGreen = gl
        If bl < 0& Then bl = 0& Else If bl > 255& Then bl = 255&
        rgbData(OfcY + X).rgbBlue = bl
    Next X
    ShowProgress (Y + 1) / h, DoDoEvents:=True
Next Y

UnReferAry AryPtr(rgbData)

ShowProgress 1.01

Exit Sub
eh:
UnReferAry AryPtr(rgbData)
ErrRaise "DrawWaves"
End Sub

Public Sub DrawELines(ByRef Data() As Long, _
                      ByRef WS() As typWaveSource, _
                      Power As Double)
Const nq As Double = 1 'lines per 1 in strength

Const dStep As Double = 0.7
Const kWH As Double = 20

Dim rgbData() As RGBQUAD

Dim w As Long, h As Long
Dim UBX As Long, UBY As Long 'w-1,h-1
Dim X As Double, Y As Double 'current position
Dim dx As Double, dy As Double 'different needs
Dim ds As Double

Dim i As Long
'Dim Reds() As Double, Greens() As Double, Blues() As Double
Dim Colors() As RGBTriLong
Dim xs() As Double, ys() As Double
Dim Ses() As Double '=strength
Dim WSUB As Long
Dim Ofc As Long

Dim Dist2 As Double, Dist As Double
Dim Ex As Double, Ey As Double
Dim InCharge As Long
Dim Foo As Double

Dim nLines As Long 'total number of lines tobe drawn
Dim LinesCounter As Long

Dim CurColor As RGBTriLong 'the color of current line

Dim Fi0 As Double 'Angle for the first line
Dim Fi As Double

Dim Pi As Double
Pi = GetPi

Dim iFrom As Long 'loop var for the source, from which the line goes
Dim iLine As Long 'loop var for lines around the current source
Dim PixCounter As Long 'for the limit for a line - no infinite loops
Dim MPC As Long

Dim SumQ As Double 'sum charge. Should always be positive.
                   'If not, invert all the system.
Dim n As Long 'the number of lines from current source
Power = Power + 1

Dim AbsE As Double, InvAbsE As Double

Dim tx As Long, ty As Long
Dim ctx As Long, cty As Long

On Error GoTo eh

If AryDims(AryPtr(WS)) <> 1 Then
    If AryDims(AryPtr(WS)) = 0 Then
        Err.Raise dbCWS
    Else
        Err.Raise 12321, , "A one-dimensional array is required (internal error)."
    End If
End If
WSUB = UBound(WS)
ShowProgress 0

ReDim Colors(0 To WSUB)
ReDim xs(0 To WSUB), ys(0 To WSUB)
ReDim Ses(0 To WSUB)
For i = 0 To WSUB
    xs(i) = WS(i).Pos.X
    ys(i) = WS(i).Pos.Y
    Ses(i) = WS(i).Strength
    Colors(i).rgbRed = WS(i).Color.rgbRed
    Colors(i).rgbGreen = WS(i).Color.rgbGreen
    Colors(i).rgbBlue = WS(i).Color.rgbBlue
    If Ses(i) > 0 Then
        nLines = nLines + Round(Ses(i) * nq)
    End If
    SumQ = SumQ + Ses(i)
Next i
If SumQ < 0 Then
    nLines = 0
    For i = 0 To WSUB
        Ses(i) = -Ses(i)
        If Ses(i) > 0 Then
            nLines = nLines + Round(Ses(i) * nq)
        End If
    Next i
End If

TestDims Data, 2
w = UBound(Data, 1) + 1
h = UBound(Data, 2) + 1
UBX = w - 1
UBY = h - 1
ConstructAry AryPtr(rgbData), VarPtr(Data(0, 0)), 4&, w * h

MPC = -Int(-(w + h) * kWH)

For iFrom = 0 To WSUB
    n = Round(Ses(iFrom) * nq)
    If n > 0 Then
        CurColor = Colors(iFrom)
        'first determine field direction
        X = xs(iFrom)
        Y = ys(iFrom)
        GoSub CalcE
        Fi0 = Arg(-Ex, -Ey)
        'second - calc first angle
        Fi0 = Fi0 + 2 * Pi / n * 0.5
        For iLine = 0 To n - 1
            X = xs(iFrom) + Cos(Fi0 + iLine * 2 * Pi / n) * dStep * 1.001
            Y = ys(iFrom) + Sin(Fi0 + iLine * 2 * Pi / n) * dStep * 1.001
            
            PixCounter = 0
            InCharge = 0
            Do
                GoSub CalcE
                AbsE = Ex * Ex + Ey * Ey
                If AbsE = 0# Then Exit Do
                InvAbsE = 1# / Sqr(AbsE)
                Ex = Ex * InvAbsE
                Ey = Ey * InvAbsE
                X = X + Ex
                Y = Y + Ey
                GoSub PutPixel
                PixCounter = PixCounter + 1&
            Loop Until PixCounter > MPC Or CBool(InCharge)
            LinesCounter = LinesCounter + 1
            ShowProgress LinesCounter / nLines
        Next iLine
    End If
Next iFrom

UnReferAry AryPtr(rgbData)
ShowProgress 1.01

Exit Sub

CalcE: '(x,y) excludes the charge if is too close and sets InCharge
    Ex = 0
    Ey = 0
    For i = 0 To WSUB
        dx = X - xs(i)
        dy = Y - ys(i)
        Dist = Sqr(dx * dx + dy * dy)
        If Dist < dStep Then
            'ignore this charge
            InCharge = i
        Else
            Foo = Ses(i) / Dist ^ Power
            Ex = Ex + dx * Foo
            Ey = Ey + dy * Foo
        End If
    Next i
Return


PutPixel:
    tx = Int(X)
    ty = Int(Y)
    If tx >= 0 And ty >= 0 And tx < UBX And ty < UBY Then
        ctx = tx
        cty = ty
        GoSub Add
        
        ctx = tx + 1
        GoSub Add
        
        cty = ty + 1
        GoSub Add
        
        ctx = tx
        GoSub Add
    ElseIf tx >= -1 And ty >= -1 And tx <= UBX And ty <= UBY Then
        ctx = tx
        cty = ty
        GoSub AddWCheck
        
        ctx = tx + 1
        GoSub AddWCheck
        
        cty = ty + 1
        GoSub AddWCheck
        
        ctx = tx
        GoSub AddWCheck
    End If
    
Return

ProcessY:
    
Return

Add:
    Ofc = ctx + cty * w
    dx = 1 - Abs(X - ctx)
    dy = 1 - Abs(Y - cty)
    ds = dx * dy
    rgbData(Ofc).rgbBlue = rgbData(Ofc).rgbBlue * (1 - ds) + CurColor.rgbBlue * ds
    rgbData(Ofc).rgbGreen = rgbData(Ofc).rgbGreen * (1 - ds) + CurColor.rgbGreen * ds
    rgbData(Ofc).rgbRed = rgbData(Ofc).rgbRed * (1 - ds) + CurColor.rgbRed * ds
Return

AddWCheck:
    If ctx >= 0 And cty >= 0 And ctx <= UBX And cty <= UBY Then
        Ofc = ctx + cty * w
        dx = 1 - Abs(X - ctx)
        dy = 1 - Abs(Y - cty)
        ds = dx * dy
        rgbData(Ofc).rgbBlue = rgbData(Ofc).rgbBlue * (1 - ds) + CurColor.rgbBlue * ds
        rgbData(Ofc).rgbGreen = rgbData(Ofc).rgbGreen * (1 - ds) + CurColor.rgbGreen * ds
        rgbData(Ofc).rgbRed = rgbData(Ofc).rgbRed * (1 - ds) + CurColor.rgbRed * ds
    End If
Return

Exit Sub
eh:
UnReferAry AryPtr(rgbData)
ErrRaise "DrawELines"
End Sub

Public Function Arg(ByVal X As Double, ByVal Y As Double) As Double
Dim Rslt As Double
If X = 0# And Y = 0# Then
    Rslt = 0#
ElseIf Abs(X) > Abs(Y) Then
    Rslt = Atan2(X, Y)
Else
    Rslt = Pi * 0.5 - Atan2(Y, X)
End If
Arg = Rslt + Int((-Rslt + Pi) * 0.5 / Pi) * 2# * Pi
End Function

Private Function Atan2(ByVal X As Double, ByVal Y As Double) As Double
If X > 0# Then
    Atan2 = Atn(Y / X)
Else
    Atan2 = Pi + Atn(Y / X)
End If
End Function


