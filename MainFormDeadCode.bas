Attribute VB_Name = "MainFormDeadCode"
''''''''''polygon testing
'Dim plg As New clsPolygon
'plg.AddVertex 100, 90
'plg.AddVertex 200, 400
'plg.AddVertex 300, 100
'plg.NewSubpolygon
''plg.AddVertex 150, 130
''plg.AddVertex 180, 350
''plg.AddVertex 250, 120
'Dim i As Long
'For i = 0 To 200 - 1
'  plg.AddVertex 200 + Sin(i / 200 * Pi * 2) * 100, 200 + Cos(i / 200 * Pi * 2) * 100
'Next i
'Dim Pix() As AlphaPixel, nPix As Long, nAlloc As Long
'plg.GetPoints Pix, nPix, nAlloc
'StartPixelAction
'dbDrawPixels Pix, nPix, DrawTemp:=False, ForceDraw:=True, DrawMode:=dmNormal
'
'Exit Sub

''''''''''''''''rotrect testing
'Dim Pix() As AlphaPixel, nPix As Long
'Dim Pnt As vtVertex
'DrawingEngine.AntiAliasingSharpness = 1
'Dim i As Long
'Dim j As Long
'For j = -3 To 3
'  For i = -20 To 20
'    Pnt.x = intW / 2 + 140 * j
'    Pnt.y = intH / 2 + 120 * i
'    DrawingEngine.pntRectangle Pnt, 100 * Rnd(1), 50 * Rnd(1), Rnd(1) * 360, Pix, nPix
'    'DrawingEngine.pntRectangle Pnt, 100, 50, -i * 2, Pix, nPix
'  Next i
'Next j
'Debug.Print nPix
'StartPixelAction
'dbDrawPixels Pix, nPix, DrawTemp:=False, ForceDraw:=False, DrawMode:=dmNormal
'Refr

