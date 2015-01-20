Attribute VB_Name = "mdlFill"
Option Explicit
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Const MemInit As Long = &H10000
Const MemIncFactor As Double = 1.5

'Flags constants (some may be unused)
Const PixPainted = &H1000000
Const PixQueued = &H2000000 'unused
Const PixBorder = &H4000000 'unused
'Constants for flags tests
Const IsProcessed = PixPainted Or PixQueued Or PixBorder


Dim ImageW As Long, ImageH As Long
Dim rgbData() As RGBQUAD
Dim lngData() As Long
Dim gSrcColor As Long ', gToColor As Long
Dim gSrcColorRGB As RGBQUAD ', gToColorRGB As RGBQUAD
Dim gTreshold As Long
Dim gTresholdRGB As RGBQUAD
Dim gSrcX As Long, gSrcY As Long
Dim gFillMode As dbFillBorderMode
Dim gPixels As typPixelList
Dim gxMin As Long, gyMin As Long
Dim gxMax As Long, gyMax As Long
Dim tmpRGB1 As RGBQUAD
Dim gCurPointX As Long, gCurPointY As Long

Public Enum dbFillBorderMode
    dbFillSingleColor = 0
    'Fills the single-coloured area
    dbFillBorderColor = 1
    'Fills the area whiches border has the color specified by Treshold
    dbFillGradientBorder = 2
    'The border is where gradient is more than Treshold
    dbFillColorDelta = 3
    'Fills where the color differs the one given by source pixel
    'not more than by Treshold
    dbFillColorUp = 4
    'fills while the color grows
    dbFillColorDown = 5
    'opposite to UP
End Enum

Public Enum dbFillMode
    FMSingleColor = 0
    FMColorAlphaBlended = 1
    FMTextured = 2
End Enum

Type FillSettings
    BorderMode As dbFillBorderMode
    Treshold As Long
    FillMode As dbFillMode
    FillAlpha As Long
    Texture() As Long
    TexOrigin As New clsAligner
End Type

Public Type vtPointInt
    X As Integer
    Y As Integer
End Type

Public Type typPixelList
    Points() As vtPointInt
    nPoints As Long 'Number of legal points in stack
    nMem As Long 'Number of elements in the array
End Type


'currently, returns nothing
'Pixels contain all the pixels needed to be painted
'Data's Reserved channel is used. It must contain zeros.
'It will contain zeros after execution if no error occurs.
'Rect is not a rect but bounds. Right and bottom are including.
Public Function vtCalcFillPixels(ByVal X As Long, ByVal Y As Long, _
                                 ByVal FillMode As dbFillBorderMode, _
                                 ByVal Treshold As Long, _
                                 ByRef Data() As Long, _
                                 ByRef Pixels As typPixelList, _
                                 ByRef Bounds As RECT) As Long
Dim n As Long 'n of pixels output
Dim i As Long 'pointer
If AryDims(AryPtr(Data)) <> 2 Then
    Err.Raise 1111, "vtCalcFillPixels", "Bad data passed. Incorrect number of dimensions."
End If
BreakKeyPressed
ImageW = UBound(Data, 1) + 1
ImageH = UBound(Data, 2) + 1
If X < 0 Or Y < 0 Or X >= ImageW Or Y >= ImageH Then
    Err.Raise 1111, "vtCalcFillPixels", "Bad source point. It lies outside of the picture."
End If
On Error GoTo eh
ConstructAry AryPtr(lngData), VarPtr(Data(0, 0)), 4, ImageW, ImageH
ConstructAry AryPtr(rgbData), VarPtr(Data(0, 0)), 4, ImageW, ImageH

'Fill some global vars
gSrcX = X
gSrcY = Y
gFillMode = FillMode
gTreshold = Treshold
CopyMemory gTresholdRGB, gTreshold, 4
gSrcColor = lngData(X, Y)
gSrcColorRGB = rgbData(X, Y)
gxMin = X
gyMin = Y
gxMax = X
gyMax = Y

'Put a single pixel into the stack and mark it
ReDim gPixels.Points(0 To 0)
With gPixels.Points(0)
    .X = X
    .Y = Y
End With
gPixels.nPoints = 1
gPixels.nMem = 1
lngData(X, Y) = lngData(X, Y) Or PixPainted

i = 0 'pointer to in-pixels
n = 1 'number of in-pixels

Do
    'i points to in-pixels
    i = i + n
    'now i points to out-pixels
    n = ProcessStack(i - n, n, i)
    'and now out-pixels become in-pixels. thus, i points to in-pixels
    'and n is now the new size of inpixels
Loop While n > 0 And Not BreakKeyPressed

gPixels.nMem = gPixels.nPoints
Debug.Assert gPixels.nPoints = i + n 'Please, dont press stop!!!
If gPixels.nMem = 0 Then
    Erase gPixels.Points
Else
    ReDim Preserve gPixels.Points(0 To gPixels.nMem - 1)
End If

'The worst stage finished at last.
'Now it's time to clear all these marks in data.
RemoveFlags
'Don't forget to send this data away
Erase Pixels.Points
SwapArys AryPtr(gPixels.Points), AryPtr(Pixels.Points)
Pixels.nPoints = gPixels.nPoints
Pixels.nMem = gPixels.nMem
gPixels.nPoints = 0
gPixels.nMem = 0
Bounds.Left = gxMin
Bounds.Top = gyMin
Bounds.Right = gxMax
Bounds.Bottom = gyMax


UnReferAry AryPtr(lngData)
UnReferAry AryPtr(rgbData)
Exit Function
eh:
UnReferAry AryPtr(lngData)
UnReferAry AryPtr(rgbData)
mdlErrors.RaiseError "vtCalcFillPixels"
End Function

Private Function FillIsBorder(ByVal X As Long, ByVal Y As Long) As Boolean
'If x < 0 Or y < 0 Or x >= ImageW Or y >= ImageH Then 'range check - always border
'    FillIsBorder = True
'    Exit Function
'End If
Select Case gFillMode
    Case dbFillBorderMode.dbFillBorderColor
        FillIsBorder = (lngData(X, Y) And &HFFFFFF) = gTreshold
    Case dbFillBorderMode.dbFillSingleColor
        FillIsBorder = (lngData(X, Y) And &HFFFFFF) <> gSrcColor
    Case dbFillBorderMode.dbFillGradientBorder
        CalcGradient X, Y
        If tmpRGB1.rgbRed > gTresholdRGB.rgbRed Then
            FillIsBorder = True
        ElseIf tmpRGB1.rgbGreen > gTresholdRGB.rgbGreen Then
            FillIsBorder = True
        ElseIf tmpRGB1.rgbBlue > gTresholdRGB.rgbBlue Then
            FillIsBorder = True
        End If
    Case dbFillBorderMode.dbFillColorDelta
        If Abs(CLng(rgbData(X, Y).rgbRed) - gSrcColorRGB.rgbRed) > gTresholdRGB.rgbRed Then
            FillIsBorder = True
        ElseIf Abs(CLng(rgbData(X, Y).rgbGreen) - gSrcColorRGB.rgbGreen) > gTresholdRGB.rgbGreen Then
            FillIsBorder = True
        ElseIf Abs(CLng(rgbData(X, Y).rgbBlue) - gSrcColorRGB.rgbBlue) > gTresholdRGB.rgbBlue Then
            FillIsBorder = True
        End If
    Case dbFillBorderMode.dbFillColorDown
        If CLng(rgbData(X, Y).rgbRed) + rgbData(X, Y).rgbGreen + rgbData(X, Y).rgbBlue _
           > _
           CLng(rgbData(gCurPointX, gCurPointY).rgbRed) + rgbData(gCurPointX, gCurPointY).rgbGreen + rgbData(gCurPointX, gCurPointY).rgbBlue Then
            FillIsBorder = True
'        ElseIf rgbData(x, y).rgbGreen > rgbData(gCurPointX, gCurPointY).rgbGreen Then
'            FillIsBorder = True
'        ElseIf rgbData(x, y).rgbBlue > rgbData(gCurPointX, gCurPointY).rgbBlue Then
'            FillIsBorder = True
        End If
    Case dbFillBorderMode.dbFillColorUp
        If CLng(rgbData(X, Y).rgbRed) + rgbData(X, Y).rgbGreen + rgbData(X, Y).rgbBlue _
           < _
           CLng(rgbData(gCurPointX, gCurPointY).rgbRed) + rgbData(gCurPointX, gCurPointY).rgbGreen + rgbData(gCurPointX, gCurPointY).rgbBlue Then
            FillIsBorder = True
'        ElseIf rgbData(x, y).rgbGreen < rgbData(gCurPointX, gCurPointY).rgbGreen Then
'            FillIsBorder = True
'        ElseIf rgbData(x, y).rgbBlue < rgbData(gCurPointX, gCurPointY).rgbBlue Then
'            FillIsBorder = True
        End If
End Select
End Function

'Places the gradient into tmpRGB1
'x,y must be range-checked (but not x+1,y and so on)
Private Sub CalcGradient(ByVal X As Long, _
                         ByVal Y As Long)
Static DistLUT(-255 To 255, -255 To 255) As Byte
Static Foo As Boolean 'For filling the DistLUT
Static dx As Long, dy As Long 'static for performance
Static tx As Long, ty As Long
Static tColor1 As RGBQUAD, tColor2 As RGBQUAD
Static tColor3 As RGBQUAD, tColor4 As RGBQUAD
Static InvSqr2 As Double
If Not Foo Then
    Foo = True
    InvSqr2 = 1 / Sqr(2)
    For dy = -255 To 255
        For dx = -255 To 255
            DistLUT(dx, dy) = Sqr(dx * dx + dy * dy) * InvSqr2
        Next dx
    Next dy
End If

tx = X + 1
If tx >= ImageW Then tx = ImageW - 1
ty = Y
tColor1 = rgbData(tx, ty)

tx = X - 1
If tx < 0 Then tx = 0
tColor2 = rgbData(tx, ty)

tx = X
ty = Y + 1
If ty >= ImageH Then ty = ImageH - 1
tColor3 = rgbData(tx, ty)

ty = Y - 1
If ty < 0 Then ty = 0
tColor4 = rgbData(tx, ty)

tmpRGB1.rgbRed = DistLUT(CLng(tColor1.rgbRed) - tColor2.rgbRed, _
                         CLng(tColor3.rgbRed) - tColor4.rgbRed)
tmpRGB1.rgbGreen = DistLUT(CLng(tColor1.rgbGreen) - tColor2.rgbGreen, _
                           CLng(tColor3.rgbGreen) - tColor4.rgbGreen)
tmpRGB1.rgbBlue = DistLUT(CLng(tColor1.rgbBlue) - tColor2.rgbBlue, _
                          CLng(tColor3.rgbBlue) - tColor4.rgbBlue)
End Sub

'Takes the pixels-stack and fills the OutStack with new points
'Returns the number of points added to outstack.
Private Function ProcessStack(ByVal InStart As Long, _
                              ByVal InLen As Long, _
                              ByVal OutStart As Long) As Long
Static i As Long 'static for speed. This function should not
                 'be called recursively
Static X As Long, Y As Long
Static Counter As Long
Static OutputPointer As Long
OutputPointer = OutStart
    Counter = 0&
    'process all the points in old stack (neighbour ones)
    With gPixels
    For i = InStart To InStart + InLen - 1&
        gCurPointX = .Points(i).X
        gCurPointY = .Points(i).Y
        X = gCurPointX + 1&
        Y = gCurPointY
        GoSub AddOutPoint
    
        X = gCurPointX - 1&
        Y = gCurPointY
        GoSub AddOutPoint
    
        X = gCurPointX
        Y = gCurPointY + 1&
        GoSub AddOutPoint
    
        X = gCurPointX
        Y = gCurPointY - 1&
        GoSub AddOutPoint
    Next i
    End With
    
    'Validate nPoints
    gPixels.nPoints = OutputPointer '-1+1
    
    'output a value
    ProcessStack = Counter

Exit Function

AddOutPoint: '(x,y)
    If X < 0& Or Y < 0& Or X >= ImageW Or Y >= ImageH Then 'range check - always border
        Return
    End If
    If lngData(X, Y) And IsProcessed Then
        Return
    End If
    If FillIsBorder(X, Y) Then
        Return
    End If
    lngData(X, Y) = lngData(X, Y) Or PixPainted 'mark this pixel
    If OutputPointer >= gPixels.nMem Then
        gPixels.nMem = -Int(-gPixels.nMem * MemIncFactor)
        If gPixels.nMem = 0 Then
            ReDim gPixels.Points(0 To 0)
            gPixels.nMem = MemInit
        End If
        ReDim Preserve gPixels.Points(0 To gPixels.nMem - 1)
    End If
    gPixels.Points(OutputPointer).X = X
    gPixels.Points(OutputPointer).Y = Y
    OutputPointer = OutputPointer + 1&
    Counter = Counter + 1&
    'increase region
    If X < gxMin Then gxMin = X
    If Y < gyMin Then gyMin = Y
    If X > gxMax Then gxMax = X
    If Y > gyMax Then gyMax = Y
    
Return

End Function

'Removes all marks produced by vtCalcFillPixels
Private Function RemoveFlags()
Dim i As Long
Dim X As Integer, Y As Integer
For i = 0 To gPixels.nPoints - 1
    X = gPixels.Points(i).X
    Y = gPixels.Points(i).Y
    lngData(X, Y) = lngData(X, Y) And &HFFFFFF
Next i
End Function

'Fills a set of pixels with a single color.
Public Function vtFillPixels(ByRef Data() As Long, _
                             ByRef Pixels As typPixelList, _
                             ByVal lngColor As Long)
Dim i As Long
Dim X As Long, Y As Long
lngColor = lngColor And &HFFFFFF
For i = 0 To Pixels.nPoints - 1
    Data(Pixels.Points(i).X, Pixels.Points(i).Y) = lngColor
Next i
End Function

'Fills the set of pixels with the specified texture
Public Sub vtFillTexturize(ByRef Data() As Long, _
                           ByRef Pixels As typPixelList, _
                           ByRef Texture() As Long, _
                           ByVal OrgX As Long, _
                           ByVal OrgY As Long)
Dim X As Long, Y As Long
Dim TexW As Long, TexH As Long
Dim i As Long
If AryDims(AryPtr(Texture)) <> 2 Then
    Err.Raise 1111, "vtFillTexturize", "Texture must be a bidimensional array."
End If
TexW = UBound(Texture, 1) + 1
TexH = UBound(Texture, 2) + 1
OrgX = OrgX Mod TexW
If OrgX > 0 Then OrgX = OrgX - TexW
OrgY = OrgY Mod TexH
If OrgY > 0 Then OrgY = OrgY - TexH
For i = 0 To Pixels.nPoints - 1
    X = Pixels.Points(i).X
    Y = Pixels.Points(i).Y
    Data(X, Y) = Texture((X - OrgX) Mod TexW, (Y - OrgY) Mod TexH)
Next i
End Sub

'Paints the set of pixels with given Alpha value
'Alpha: 0 - fully replaces, 255 - leave source
Public Sub vtFillAlphaBlend(ByRef Data() As Long, _
                            ByRef Pixels As typPixelList, _
                            ByVal lngColor As Long, _
                            ByVal Alpha As Long)
       '(old-color)
Dim RLUT(0 To 255) As Byte
Dim GLUT(0 To 255) As Byte
Dim BLUT(0 To 255) As Byte
Dim i As Long
Dim rgbColor As RGBQUAD
Dim rgbData() As RGBQUAD
Dim AlphaDbl As Double
Dim w As Long, h As Long
Dim X As Long, Y As Long
CopyMemory rgbColor, lngColor, 4
AlphaDbl = Alpha / 255#
For i = 0 To 255
    RLUT(i) = (i - rgbColor.rgbRed) * AlphaDbl + rgbColor.rgbRed
    GLUT(i) = (i - rgbColor.rgbGreen) * AlphaDbl + rgbColor.rgbGreen
    BLUT(i) = (i - rgbColor.rgbBlue) * AlphaDbl + rgbColor.rgbBlue
Next i
w = UBound(Data, 1) + 1
h = UBound(Data, 2) + 1
On Error GoTo eh
ConstructAry AryPtr(rgbData), VarPtr(Data(0, 0)), 4, w, h
For i = 0 To Pixels.nPoints - 1
    X = Pixels.Points(i).X
    Y = Pixels.Points(i).Y
    rgbData(X, Y).rgbRed = RLUT(rgbData(X, Y).rgbRed)
    rgbData(X, Y).rgbGreen = GLUT(rgbData(X, Y).rgbGreen)
    rgbData(X, Y).rgbBlue = BLUT(rgbData(X, Y).rgbBlue)
Next i
UnReferAry AryPtr(rgbData)
Exit Sub
eh:
UnReferAry AryPtr(rgbData)
Debug.Assert False
ErrRaise "vtFillAlphaBlend"
End Sub
