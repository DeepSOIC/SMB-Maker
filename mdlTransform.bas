Attribute VB_Name = "mdlTransform"
Option Explicit

Public Type FloatPoint
  x As Double
  y As Double
End Type
Public Type vtQGon
  v(1 To 4) As FloatPoint
  '1--2
  '----
  '3--4
End Type


Public Type FloatRect
  X1 As Double
  Y1 As Double
  X2 As Double 'not including
  Y2 As Double 'not including
End Type

'output rect will include source rect
Public Function RectFromFloatRect(rct As FloatRect) As RECT
Dim o As RECT
o.Left = Int(rct.X1)
o.Top = Int(rct.Y1)
o.Right = -Int(-rct.X2)
o.Bottom = -Int(-rct.Y2)
RectFromFloatRect = o
End Function

Public Function FloatRectFromRect(rct As RECT) As FloatRect
Dim o As FloatRect
o.X1 = rct.Left
o.Y1 = rct.Top
o.X2 = rct.Right
o.Y2 = rct.Bottom
FloatRectFromRect = o
End Function


'srcgon from srcdata is thansferred to
'dstrect on dstdata,
'processrect defines the area to outputon dstdata,
'not necessarily inside of dstrect
'Background is bg-color. if =-1 then texture mode
Sub TransformBlock(ByRef DstData() As Long, _
                   ByRef dstRect As FloatRect, _
                   ByRef srcData() As Long, _
                   ByRef srcGon As vtQGon, _
                   ByRef ProcessRect As RECT, _
                   Optional ByVal Background As Long)
Dim x As Long, y As Long 'position in dstData
Dim ex As Double, ey As Double 'x,y normalized to DstRect
Dim dst() As RGBQUAD
Dim tx As Double, ty As Double 'position in srcdata
Dim itx As Long, ity As Long 'integer part of tx,ty
Dim fx As Double, fy As Double 'fractional part of tx,ty
Dim src() As RGBQUAD
Dim c1 As RGBQUAD, c2 As RGBQUAD, c3 As RGBQUAD, c4 As RGBQUAD
Dim BackgroundRGB As RGBQUAD
Dim Tex As Boolean
Tex = Background = -1
If Not Tex Then CopyMemory BackgroundRGB, Background, 4
'initialize vectors
  Dim topDX As Double, topDY As Double
  Dim botDX As Double, botDY As Double
  Dim lefDX As Double, lefDY As Double
  Dim rigDX As Double, rigDY As Double
  topDX = srcGon.v(2).x - srcGon.v(1).x
  topDY = srcGon.v(2).y - srcGon.v(1).y
  botDX = srcGon.v(4).x - srcGon.v(3).x
  botDY = srcGon.v(4).y - srcGon.v(3).y
  lefDX = srcGon.v(3).x - srcGon.v(1).x
  lefDY = srcGon.v(3).y - srcGon.v(1).y
  rigDX = srcGon.v(4).x - srcGon.v(2).x
  rigDY = srcGon.v(4).y - srcGon.v(2).y
  Dim sx0 As Double, sy0 As Double 'src starting point (=vertex1)
  sx0 = srcGon.v(1).x
  sy0 = srcGon.v(1).y
  Dim drw As Double, drh As Double
  Dim x0 As Double, y0 As Double 'vars useful to calculate ex,ey
  With dstRect
    drw = .X2 - .X1
    drh = .Y2 - .Y1
    x0 = .X1
    y0 = .Y1
    If drw <= 0 Or drh <= 0 Then Err.Raise 12345, "TransformBlock", "Destination block defined improperly (it must have positive width, height)."
  End With
'initialize arrays
Dim dstW As Long, dstH As Long
Dim srcW As Long, srcH As Long
AryWH AryPtr(DstData), dstW, dstH
If dstW * dstH <= 0 Then Exit Sub
AryWH AryPtr(srcData), srcW, srcH
If srcW * srcH <= 0 Then
  If Tex Then
    For y = Max(0, ProcessRect.Top) To Min(ProcessRect.Bottom - 1, dstH - 1)
      For x = Max(0, ProcessRect.Left) To Min(ProcessRect.Right - 1, dstW - 1)
        DstData(x, y) = Background
      Next x
    Next y
  End If
  Exit Sub
End If
SwapArys AryPtr(dst), AryPtr(DstData)
SwapArys AryPtr(src), AryPtr(srcData)
On Error GoTo eh
If Not Tex Then
  For y = Max(0, ProcessRect.Top) To Min(ProcessRect.Bottom - 1, dstH - 1)
    ey = (y - y0) / drh
    For x = Max(0, ProcessRect.Left) To Min(ProcessRect.Right - 1, dstW - 1)
      ex = (x - x0) / drw
      'working SMBScript tx,ty formulas
      'sx0+ex*(topdx'*(1-ey)+botdx*(ey)')+ey*(lefdx*(1-ex)+rigdx*(ex)),
      'sy0+ex*(topdy'*(1-ey)+botdy*(ey)')+ey*(lefdy*(1-ex)+rigdy*(ex))
      tx = sx0 + ex * topDX + ey * lefDX + (rigDX - lefDX) * ex * ey
      itx = Int(tx)
      fx = tx - itx
      
      ty = sy0 + ex * topDY + ey * lefDY + (rigDY - lefDY) * ex * ey
      ity = Int(ty)
      fy = ty - ity
      If itx >= 0 And ity >= 0 And itx <= srcW - 2 And ity <= srcH - 2 Then
        'most of pixels - need fastest response for them
        dst(x, y).rgbBlue = src(itx, ity).rgbBlue * (1 - fx) * (1 - fy) + _
                    src(itx + 1, ity).rgbBlue * (fx) * (1 - fy) + _
                    src(itx, ity + 1).rgbBlue * (1 - fx) * (fy) + _
                    src(itx + 1, ity + 1).rgbBlue * (fx) * (fy)
        dst(x, y).rgbGreen = src(itx, ity).rgbGreen * (1 - fx) * (1 - fy) + _
                    src(itx + 1, ity).rgbGreen * (fx) * (1 - fy) + _
                    src(itx, ity + 1).rgbGreen * (1 - fx) * (fy) + _
                    src(itx + 1, ity + 1).rgbGreen * (fx) * (fy)
        dst(x, y).rgbRed = src(itx, ity).rgbRed * (1 - fx) * (1 - fy) + _
                    src(itx + 1, ity).rgbRed * (fx) * (1 - fy) + _
                    src(itx, ity + 1).rgbRed * (1 - fx) * (fy) + _
                    src(itx + 1, ity + 1).rgbRed * (fx) * (fy)
      ElseIf itx < -1 Or itx >= srcW Or ity < -1 Or ity >= srcH Then
        'surely will be background
        dst(x, y) = BackgroundRGB
      Else
        If InData(srcW, srcH, itx, ity) Then
          c1 = src(itx, ity)
        Else
          c1 = BackgroundRGB
        End If
        If InData(srcW, srcH, itx + 1, ity) Then
          c2 = src(itx + 1, ity)
        Else
          c2 = BackgroundRGB
        End If
        If InData(srcW, srcH, itx, ity + 1) Then
          c3 = src(itx, ity + 1)
        Else
          c3 = BackgroundRGB
        End If
        If InData(srcW, srcH, itx + 1, ity + 1) Then
          c4 = src(itx + 1, ity + 1)
        Else
          c4 = BackgroundRGB
        End If
        dst(x, y).rgbBlue = c1.rgbBlue * (1 - fx) * (1 - fy) + _
                    c2.rgbBlue * (fx) * (1 - fy) + _
                    c3.rgbBlue * (1 - fx) * (fy) + _
                    c4.rgbBlue * (fx) * (fy)
        dst(x, y).rgbGreen = c1.rgbGreen * (1 - fx) * (1 - fy) + _
                    c2.rgbGreen * (fx) * (1 - fy) + _
                    c3.rgbGreen * (1 - fx) * (fy) + _
                    c4.rgbGreen * (fx) * (fy)
        dst(x, y).rgbRed = c1.rgbRed * (1 - fx) * (1 - fy) + _
                    c2.rgbRed * (fx) * (1 - fy) + _
                    c3.rgbRed * (1 - fx) * (fy) + _
                    c4.rgbRed * (fx) * (fy)
      End If
    Next x
  Next y
Else
  MsgBox "аффтар идиот забыл реализовать текстурный режим!"

End If
SwapArys AryPtr(dst), AryPtr(DstData)
SwapArys AryPtr(src), AryPtr(srcData)
Exit Sub
eh:
SwapArys AryPtr(dst), AryPtr(DstData)
SwapArys AryPtr(src), AryPtr(srcData)
ErrRaise "TransformBlock"
End Sub

Private Function InData(w As Long, h As Long, x As Long, y As Long) As Boolean
InData = x >= 0 And y >= 0 And x < w And y < h
End Function
