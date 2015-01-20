Attribute VB_Name = "mdlReRGB"
Option Explicit

Private Sub ShowProgress(ByVal Part As Double)
MainForm.ShowProgress Part * 100, DoDoEvents:=True
End Sub

Public Sub dbMatrix(ByRef IOData() As Long, ByRef Matrix() As Double)
Dim Range As RECT
Dim tmpData() As Long
AryWH AryPtr(IOData), Range.Right, Range.Bottom
dbMatrixEx IOData, Matrix, tmpData, Range
SwapArys AryPtr(IOData), AryPtr(tmpData)
'Dim X As Long, Y As Long
''Dim tmpData() As RGBQUAD
'Dim rr As Double, rg As Double, rb As Double
'Dim gr As Double, gg As Double, gb As Double
'Dim br As Double, BG As Double, bb As Double
'Dim r As Long, g As Long, B As Long
'Dim w As Long, h As Long
'Dim pDataRGB() As RGBQUAD
'
'On Error GoTo eh
'
'MainForm.DisableMe
'MainForm.ShowProgress 0, Not DontDoEvents
'
'w = UBound(pData, 1) + 1
'h = UBound(pData, 2) + 1
'
'ConstructAry AryPtr(pDataRGB), VarPtr(pData(0, 0)), 4, w, h
'
''ReDim tmpData(0 To w - 1, 0 To h - 1)
''CopyMemory tmpData(0, 0), pData(0, 0), w * h * 4
'
'rr = Matrix(0, 0)
'rg = Matrix(0, 1)
'rb = Matrix(0, 2)
'gr = Matrix(1, 0)
'gg = Matrix(1, 1)
'gb = Matrix(1, 2)
'br = Matrix(2, 0)
'BG = Matrix(2, 1)
'bb = Matrix(2, 2)
'
'For Y = 0 To h - 1
'    For X = 0 To w - 1
'        r = pDataRGB(X, Y).rgbBlue * rb + pDataRGB(X, Y).rgbGreen * rg + pDataRGB(X, Y).rgbRed * rr
'        g = pDataRGB(X, Y).rgbBlue * gb + pDataRGB(X, Y).rgbGreen * gg + pDataRGB(X, Y).rgbRed * gr
'        B = pDataRGB(X, Y).rgbBlue * bb + pDataRGB(X, Y).rgbGreen * BG + pDataRGB(X, Y).rgbRed * br
'        If r < 0 Then r = 0
'        If r > 255 Then r = 255
'        If g < 0 Then g = 0
'        If g > 255 Then g = 255
'        If B < 0 Then B = 0
'        If B > 255 Then B = 255
'        pDataRGB(X, Y).rgbRed = r
'        pDataRGB(X, Y).rgbGreen = g
'        pDataRGB(X, Y).rgbBlue = B
'    Next X
'    MainForm.ShowProgress Y * 100 / (h - 1), Not DontDoEvents
'Next Y
'ExitHere:
'On Error GoTo 0
'UnReferAry AryPtr(pDataRGB)
'MainForm.ShowProgress 101
'MainForm.RestoreMeEnabled
'Exit Sub
'eh:
'UnReferAry AryPtr(pDataRGB)
'MainForm.ClearMeEnabledStack
'ErrRaise "dbMatrix"
End Sub


Public Sub dbMatrixEx(ByRef InData() As Long, _
                      ByRef Matrix() As Double, _
                      ByRef OutData() As Long, _
                      ByRef Range As RECT)
Dim X As Long, Y As Long
Dim xOut As Long, yOut As Long
Dim OfcIn As Long, OfcOut As Long
Dim PosIn As Long, PosOut As Long
Dim AddToPosIn As Long
Dim MInData() As RGBQUAD, MOutData() As RGBQUAD
Dim InW As Long, InH As Long
Dim OutW As Long, OutH As Long
Dim BaseX As Long, BaseY As Long
Dim xf As Long, yf As Long
Dim xt As Long, yt As Long

Dim rr As Double, rg As Double, rb As Double
Dim gr As Double, gg As Double, gb As Double
Dim br As Double, BG As Double, bb As Double
Dim r As Long, g As Long, B As Long

If IsRectEmpty(Range) Then Exit Sub

TestDims InData
AryWH AryPtr(InData), InW, InH

OutW = Range.Right - Range.Left
OutH = Range.Bottom - Range.Top
BaseX = Range.Left
BaseY = Range.Bottom

RedimIfNeeded OutData, OutW, OutH

BaseX = -Range.Left
BaseY = -Range.Top
xf = Max(Range.Left, 0)
yf = Max(Range.Top, 0)
xt = Min(InW - 1, Range.Right - 1)
yt = Min(InH - 1, Range.Bottom - 1)

If yf > yt Or yf > yt Then Exit Sub

rr = Matrix(0, 0)
rg = Matrix(0, 1)
rb = Matrix(0, 2)
gr = Matrix(1, 0)
gg = Matrix(1, 1)
gb = Matrix(1, 2)
br = Matrix(2, 0)
BG = Matrix(2, 1)
bb = Matrix(2, 2)

On Error GoTo eh
ConstructAry AryPtr(MInData), VarPtr(InData(0, 0)), 4, InW * InH
ConstructAry AryPtr(MOutData), VarPtr(OutData(0, 0)), 4, OutW * OutH

For Y = yf To yt
    ShowProgress (Y - yf) / (yt - yf + 1)
    OfcIn = Y * InW
    yOut = Y + BaseY
    OfcOut = yOut * OutW
    
    AddToPosIn = OfcOut - OfcIn + BaseX
    '+PosOut = +OfcOut -OfcIn  + BaseX+ PosIn
    For PosIn = xf + OfcIn To xt + OfcIn
'        X = -OfcIn + PosIn
'        xOut = X + BaseX
'        PosOut = OfcOut + xOut
        PosOut = PosIn + AddToPosIn
        
        B = MInData(PosIn).rgbBlue * bb + MInData(PosIn).rgbGreen * BG + MInData(PosIn).rgbRed * br
        If B < 0 Then B = 0
        If B > 255 Then B = 255
        MOutData(PosOut).rgbBlue = B
        
        g = MInData(PosIn).rgbBlue * gb + MInData(PosIn).rgbGreen * gg + MInData(PosIn).rgbRed * gr
        If g < 0 Then g = 0
        If g > 255 Then g = 255
        MOutData(PosOut).rgbGreen = g
        
        r = MInData(PosIn).rgbBlue * rb + MInData(PosIn).rgbGreen * rg + MInData(PosIn).rgbRed * rr
        If r < 0 Then r = 0
        If r > 255 Then r = 255
        MOutData(PosOut).rgbRed = r
    Next PosIn
Next Y
ShowProgress 1.01
UnReferAry AryPtr(MInData)
UnReferAry AryPtr(MOutData)
Exit Sub
eh:
UnReferAry AryPtr(MInData)
UnReferAry AryPtr(MOutData)
ErrRaise "dbMatrixEx"
End Sub
