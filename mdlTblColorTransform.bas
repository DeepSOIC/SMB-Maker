Attribute VB_Name = "mdlTblColorTransform"
Option Explicit

Public Sub dbApplyGamma(ByRef IOData() As Long, _
                        ByVal mnoj As Double)
Dim tbl() As Byte
Dim i As Long
Dim tmp As Long
ReDim tbl(0 To 255)
For i = 0 To 255
    tmp = i * mnoj
    If tmp > 255& Then tmp = 255&
    tbl(i) = tmp
Next i
dbMapColors IOData, tbl, tbl, tbl
End Sub


Public Sub dbMapColors(ByRef IOData() As Long, _
                       ByRef rTbl() As Byte, _
                       ByRef gTbl() As Byte, _
                       ByRef bTbl() As Byte)
Dim Rct As RECT
TestDims IOData
Rct.Left = 0
Rct.Top = 0
Rct.Right = UBound(IOData, 1) + 1
Rct.Bottom = UBound(IOData, 2) + 1
dbMapColorsEx IOData, _
              rTbl, gTbl, bTbl, _
              Rct, _
              IOData
End Sub


Public Sub dbMapColorsEx(ByRef InData() As Long, _
                         ByRef rTbl() As Byte, _
                         ByRef gTbl() As Byte, _
                         ByRef bTbl() As Byte, _
                         ByRef Range As RECT, _
                         ByRef OutData() As Long)
Dim X As Long, Y As Long, tmp As RGBQUAD
Dim x1 As Long, y1 As Long
Dim xFrom As Long, xTo As Long
Dim w As Long, h As Long
Dim Wtodo As Long, Htodo As Long
Dim BaseX As Long, BaseY As Long
Dim AryRGBQ() As RGBQUAD
Dim OutRGBQ() As RGBQUAD
Dim OfcIn As Long
Dim OfcOut As Long

If IsRectEmpty(Range) Then Exit Sub

'DisableMe

BaseX = Range.Left
BaseY = Range.Top
Wtodo = Range.Right - Range.Left
Htodo = Range.Bottom - Range.Top

w = UBound(InData, 1) + 1
h = UBound(InData, 2) + 1

ConstructAry AryPtr(AryRGBQ), VarPtr(InData(0, 0)), 4, w * h
On Error GoTo eh

RedimIfNeeded OutData, Wtodo, Htodo
ConstructAry AryPtr(OutRGBQ), VarPtr(OutData(0, 0)), 4, Wtodo * Htodo

xFrom = Max(0, -BaseX)
xTo = Min(Wtodo - 1, w - 1 - BaseX)

For Y = Max(0, -BaseY) To Min(Htodo - 1, h - 1 - BaseY)
    y1 = BaseY + Y
    OfcOut = xFrom + Y * Wtodo
    OfcIn = BaseX + xFrom + y1 * w
    For X = xFrom To xTo
        OutRGBQ(OfcOut).rgbBlue = bTbl(AryRGBQ(OfcIn).rgbBlue)
        OutRGBQ(OfcOut).rgbGreen = gTbl(AryRGBQ(OfcIn).rgbGreen)
        OutRGBQ(OfcOut).rgbRed = rTbl(AryRGBQ(OfcIn).rgbRed)
        OfcIn = OfcIn + 1&
        OfcOut = OfcOut + 1&
    Next X
Next Y

UnReferAry AryPtr(OutRGBQ)
UnReferAry AryPtr(AryRGBQ)
'RestoreMeEnabled
Exit Sub
eh:
UnReferAry AryPtr(OutRGBQ)
UnReferAry AryPtr(AryRGBQ)
'ClearMeEnabledStack
ErrRaise "dbMapColors"
End Sub



