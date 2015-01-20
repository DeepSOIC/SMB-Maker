VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPolygon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Renders a set of polygons. Uses odd-even method of painting.
' When one polygon is inside the other, it makes a void.

Private Type ScanLinePoints
  Points() As Long 'each entry holds an intersection
                   'of an edge with the scan line
  nPoints As Long 'number of valid points in the array
  nAlloc As Long 'number of items allocated
End Type

Private Type TwoVector
  x As Double
  y As Double
End Type

Private Type SubPolygon
  Vs() As TwoVector
  nV As Long 'number of vertices
End Type

Dim SPlgs() As SubPolygon
Dim nSP As Long ' number of sub-polygons

Dim ymin As Long, ymax As Long 'scan-line range of all polygons. All values are inclusive.

Dim pColor As RGBQUAD

'nPix - number of valid pixels in Pix
'nAlloc - length of an array (may be greater than nPix)
'if there is insufficient space in an array, extra space
'is automatically allocated and reflected in nAlloc
Friend Sub GetPoints(ByRef Pix() As AlphaPixel, _
                     ByRef nPix As Long, _
                     ByRef nAlloc As Long)
On Error GoTo ExitHere
CalcMinMax
On Error GoTo 0
Dim ScanLines() As ScanLinePoints
ReDim ScanLines(0 To ymax - ymin + 1)

Dim x As Long, y As Long
Dim x1 As Double, y1 As Double 'endpoints of an edge
Dim x2 As Double, y2 As Double
Dim ip As Long 'index or subpolygon
Dim iV As Long, ivn As Long  'index of vertex
Dim PrevDir As Long, CurDir As Long
Dim y1sh As Long
Dim xf As Long, xt As Long

For ip = 0 To nSP - 1
  With SPlgs(ip)
    'this first loop is needed to get PrevDir
    For iV = 0 To .nV - 1
      y1 = .Vs(iV).y
      ivn = iV + 1
      If ivn = .nV Then ivn = 0
      y2 = .Vs(ivn).y
      If y2 > y1 Then PrevDir = 1
      If y2 < y1 Then PrevDir = -1
    Next iV
    'this is the actual job
    For iV = 0 To .nV - 1
      x1 = .Vs(iV).x
      y1 = .Vs(iV).y
      ivn = iV + 1
      If ivn = .nV Then ivn = 0
      x2 = .Vs(ivn).x
      y2 = .Vs(ivn).y
      If y1 <> y2 Then 'otherwise, the edge is exactly horizontal, and should not be considered
        CurDir = Sgn(y2 - y1)
        If CurDir = PrevDir Then
          y1sh = CurDir
        Else
          y1sh = 0
        End If
        For y = CLng(y1) + y1sh To CLng(y2) Step CurDir
          x = x1 + (x2 - x1) * (CDbl(y) - y1) / (y2 - y1)
          xf = MinL(CLng(x1), CLng(x2))
          xt = MaxL(CLng(x1), CLng(x2))
          If x < xf Then x = xf
          If x > xt Then x = xt
          GoSub AddToScanLine
        Next y
        PrevDir = CurDir
      End If
    Next iV
  End With
Next ip

For y = ymin To ymax
  With ScanLines(y - ymin)
    SortLongArray .Points, 0, .nPoints - 1
    'fill the scan line
    Debug.Assert .nPoints Mod 2 = 0
    For iV = 0 To .nPoints - 2 Step 2
      For x = .Points(iV) To .Points(iV + 1)
        'add pixel
        If nPix + 1 > nAlloc Then
          If nAlloc = 0 Then
            nAlloc = 300
            ReDim Pix(0 To nAlloc - 1)
          Else
            nAlloc = nAlloc * 1.5
            ReDim Preserve Pix(0 To nAlloc - 1)
          End If
        End If
        Pix(nPix).x = x
        Pix(nPix).y = y
        Pix(nPix).drawOpacity = 255
        Pix(nPix).rgbBlue = pColor.rgbBlue
        Pix(nPix).rgbGreen = pColor.rgbGreen
        Pix(nPix).rgbRed = pColor.rgbRed
        Pix(nPix).rgbOpacity = pColor.rgbReserved
        nPix = nPix + 1
      Next x
    Next iV
  End With
Next y

ExitHere:
Exit Sub

AddToScanLine: '(x,y)
  With ScanLines(y - ymin)
    If .nPoints + 1 > .nAlloc Then
      If .nAlloc = 0 Then
        .nAlloc = 5
        ReDim .Points(0 To .nAlloc - 1)
      Else
        .nAlloc = .nAlloc * 2
        ReDim Preserve .Points(0 To .nAlloc - 1)
      End If
    End If
    .Points(.nPoints) = x
    .nPoints = .nPoints + 1
  End With
Return

End Sub

Private Sub CalcMinMax()
Dim ip As Long ' index of subpolygon
Dim iV As Long 'index of vertex in subpolygon
Dim x As Long, y As Long
Dim FirstPoint As Boolean
FirstPoint = True
For ip = 0 To nSP - 1
  With SPlgs(ip)
    For iV = 0 To .nV - 1
      x = .Vs(iV).x 'use implicit rounding
      y = .Vs(iV).y
      If y > ymax Or FirstPoint Then
        ymax = y
      End If
      If y < ymin Or FirstPoint Then
        ymin = y
      End If
      FirstPoint = False
    Next iV
  End With
Next ip
If FirstPoint Then Err.Raise 12345, "clsPolygon", "Epmty polygon!!"
End Sub


Public Sub NewSubpolygon()
If nSP = 0 Then
  nSP = 1
  ReDim SPlgs(0 To nSP - 1)
Else
  nSP = nSP + 1
  ReDim Preserve SPlgs(0 To nSP - 1)
End If
End Sub

Public Sub AddVertex(ByVal x As Double, ByVal y As Double)
If nSP = 0 Then NewSubpolygon
With SPlgs(nSP - 1)
  If .nV = 0 Then
    .nV = 1
    ReDim .Vs(0 To .nV - 1)
  Else
    .nV = .nV + 1
    ReDim Preserve .Vs(0 To .nV - 1)
  End If
  With .Vs(.nV - 1)
    .x = x
    .y = y
  End With
End With
End Sub

'removes all vertices and all subpolygons
Public Sub Clear()
Erase SPlgs
nSP = 0
End Sub

'convention - BGR0
Public Property Let Color(ByVal NewColor As Long)
CopyMemory pColor, NewColor, 4
End Property

Public Property Get Color() As Long
Dim ret As Long
CopyMemory ret, pColor, 4
Color = ret
End Property

Private Function MinL(ByVal a As Long, ByVal b As Long) As Long
If a > b Then MinL = b Else MinL = a
End Function

Private Function MaxL(ByVal a As Long, ByVal b As Long) As Long
If a > b Then MaxL = a Else MaxL = b
End Function

Public Function IsEmpty() As Boolean
Dim i As Long
IsEmpty = True
For i = 0 To nSP - 1
  If SPlgs(i).nV >= 3 Then IsEmpty = False: Exit Function
Next i
End Function
