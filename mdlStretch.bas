Attribute VB_Name = "mdlStretch"
Option Explicit
Public Enum eStretchMode
    SMClear = 0
    SMPreserve = 1
    SMCenter = 2
    SMSimple = 3
    SMSquares = 4
    SMLinInterpol = 5
    SMWin = 6
    SMTile = 7
End Enum


Private Function ACol(ByVal Index As Long) As Long
ACol = MainForm.GetACol(Index)
End Function

Private Function TexMode() As Boolean
TexMode = MainForm.TexMode
End Function

Private Sub ShowProgress(ByVal PerCents As Double, _
                         Optional ByVal DoDoEvents As Boolean = False)
MainForm.ShowProgress PerCents, DoDoEvents
End Sub

'---------------------------Engine--------------------------------
Sub dbStretch(ByRef Data() As Long, _
              ByVal NewWidth As Long, _
              ByVal NewHeight As Long, _
              Optional ByVal StretchMode As eStretchMode = SMSquares, _
              Optional ByVal RaiseErrors As Boolean = False)
Dim OldHeight As Long, OldWidth As Long
On Error GoTo eh
MainForm.DisableMe
If NewWidth < 0 Then
    MainForm.dbTurn Data, dbFlipHor
    NewWidth = Abs(NewWidth)
End If
If NewHeight < 0 Then
    MainForm.dbTurn Data, dbFlipVer
    NewHeight = Abs(NewHeight)
End If

OldWidth = UBound(Data, 1) + 1
OldHeight = UBound(Data, 2) + 1
If OldWidth = NewWidth And OldHeight = NewHeight Then
    'do nothing
Else
    Select Case StretchMode
        Case eStretchMode.SMCenter
            vtStretchCenter Data, NewWidth, NewHeight, ACol(2), TexMode
        Case eStretchMode.SMClear
            vtStretchClear Data, NewWidth, NewHeight, ACol(2)
        Case eStretchMode.SMLinInterpol
            vtStretchLinInterpol Data, NewWidth, NewHeight, TexMode
        Case eStretchMode.SMPreserve
            vtStretchPreserve Data, NewWidth, NewHeight, ACol(2)
        Case eStretchMode.SMSimple
            vtStretchSimple Data, NewWidth, NewHeight
        Case eStretchMode.SMSquares
            vtStretchSquares Data, NewWidth, NewHeight
        Case eStretchMode.SMWin
            vtStretchWin Data, NewWidth, NewHeight
        Case eStretchMode.SMTile
            vtStretchTile Data, NewWidth, NewHeight, 0, 0
    End Select
End If
MainForm.RestoreMeEnabled
Exit Sub
eh:
MainForm.RestoreMeEnabled
If RaiseErrors Then
    Err.Raise Err.Number, "dbStretch", Err.Description
Else
    MsgBox Err.Description
End If
Exit Sub
Resume
Exit Sub
End Sub





Sub vtStretchSquares(ByRef Data() As Long, _
                     ByVal NewW As Long, _
                     ByVal NewH As Long)
Const Float_Delta As Double = 0.00001
Dim OutX As Long, OutY As Long 'Main loop vars
Dim InData() As RGBQUAD 'referred to Data
Dim OutData() As RGBQUAD 'Created. At the end swapped with Data
Dim OldW As Long, OldH As Long 'Size of the Data

Dim InX As Double, InY As Double 'Coordinates of New pixel in old-pixels
Dim PixW As Double, PixH As Double 'New pixel dimensions in old-pixels
Dim PixS As Double, InvPixS As Double 'pixel square and 1/pixel square
Dim dx As Double, dy As Double 'size of area that is inside of newpixel
Dim ds As Double 'dx*dy/pixS
Dim r As Double, g As Double, b As Double 'accumulators
Dim x As Long, y As Long 'loop variables in GetColor
Dim fx As Long, tx As Long 'bounds of old-pixels, covered by new-pixel
Dim fy As Long, ty As Long
Dim InX2 As Double, InY2 As Double
Dim MeEn As Boolean
Dim OldWtoNewW As Double

ReDim OutData(0 To NewW - 1, 0 To NewH - 1)
If AryDims(AryPtr(Data)) <> 2 Then
    'just fill OutData with back color
    SwapArys AryPtr(Data), AryPtr(OutData)
    For OutY = 0 To NewH - 1
        For OutX = 0 To NewW - 1
            Data(OutX, OutY) = ACol(2)
        Next OutX
    Next OutY
    'and exit
    Exit Sub
End If


OldW = UBound(Data, 1) + 1
OldH = UBound(Data, 2) + 1
PixW = OldW / NewW
PixH = OldH / NewH
PixS = PixW * PixH
InvPixS = 1 / PixS

On Error GoTo eh
ConstructAry AryPtr(InData), VarPtr(Data(0, 0)), 4, OldW, OldH

MeEn = MainForm.MeEnabled
For OutY = 0 To NewH - 1
    ShowProgress OutY / NewH * 100, Not MeEn
    
    InY = OutY * PixH
    InY2 = InY + PixH
    
    fy = Int(InY + Float_Delta)
    ty = -Int(-InY2 + Float_Delta) - 1
    
    For OutX = 0 To NewW - 1
        
        InX = OutX * PixW
        InX2 = InX + PixW
        
        fx = Int(InX + Float_Delta)
        tx = -Int(-InX2 + Float_Delta) - 1
                
                
        r = 0#
        g = 0#
        b = 0#
        For y = fy To ty
            
            dy = 1#
            If y < InY Then
                dy = dy - InY + y
            End If
            If y + 1& > InY2 Then
                dy = dy - (y + 1&) + InY2
            End If
            dy = dy * InvPixS
            
            For x = fx To tx
                dx = 1#
    
                If x < InX Then
                    dx = dx - InX + x
                End If
                If x + 1& > InX2 Then
                    dx = dx - (x + 1&) + InX2
                End If
        
                ds = dx * dy
                r = r + ds * InData(x, y).rgbRed
                g = g + ds * InData(x, y).rgbGreen
                b = b + ds * InData(x, y).rgbBlue
            Next x
        Next y
        
        
        OutData(OutX, OutY).rgbRed = r
        OutData(OutX, OutY).rgbGreen = g
        OutData(OutX, OutY).rgbBlue = b
        
    Next OutX

Next OutY

UnReferAry AryPtr(InData)

SwapArys AryPtr(Data), AryPtr(OutData)
ShowProgress 101
Exit Sub
'GetColor: 'gets color of point inx,iny with width of PixWidth and height of PixHeight
'    r = 0#
'    g = 0#
'    B = 0#
'    For y = fy To ty
'        For x = fx To tx
'            dx = 1#
'
'            If x < InX Then
'                dx = dx - InX + x
'            End If
'            If x + 1& > InX2 Then
'                dx = dx - (x + 1&) + InX2
'            End If
'
'            dy = InvPixS
'            If y < InY Then
'                dy = (dy - InY + y) * InvPixS
'            End If
'            If y + 1& > InY2 Then
'                dy = (dy - (y + 1&) + InY2) * InvPixS
'            End If
'
'            ds = dx * dy
'            r = r + ds * InData(x, y).rgbRed
'            g = g + ds * InData(x, y).rgbGreen
'            B = B + ds * InData(x, y).rgbBlue
'        Next x
'    Next y
'Return
Resume
eh:
Debug.Assert False
UnReferAry AryPtr(InData)
ErrRaise "vtStretchSquares"
End Sub

Sub vtStretchLinInterpol(ByRef Data() As Long, _
                         ByVal NewW As Long, _
                         ByVal NewH As Long, _
                         Optional ByVal TexMode As Boolean = True)
Dim OutX As Long, OutY As Long 'Main loop vars
Dim InData() As RGBQUAD 'referred to Data
Dim OutData() As RGBQUAD 'Created. At the end swapped with Data
Dim OldW As Long, OldH As Long 'Size of the Data

Dim InX As Double, InY As Double 'Coordinates of New pixel in old-pixels
Dim dx As Double, dy As Double 'offsets from x,y to inx,iny
Dim r As Byte, g As Byte, b As Byte 'accumulators
Dim x As Long, y As Long 'loop variables in GetColor
Dim x1 As Long, y1 As Long 'coords of surrounding pixels
Dim x2 As Long, y2 As Long
Dim x3 As Long, y3 As Long
Dim x4 As Long, y4 As Long
Dim MaxX As Long, MaxY As Long 'old-data related
Dim MeEn As Boolean
Dim rgb1 As RGBQUAD, rgb2 As RGBQUAD
Dim rgb3 As RGBQUAD, rgb4 As RGBQUAD

If AryDims(AryPtr(Data)) <> 2 Then
    'just fill OutData with back color
    vtStretchClear Data, NewW, NewH
    'and exit
    Exit Sub
End If

ReDim OutData(0 To NewW - 1, 0 To NewH - 1)

OldW = UBound(Data, 1) + 1
OldH = UBound(Data, 2) + 1
MaxX = OldW - 1
MaxY = OldH - 1
On Error GoTo eh
ConstructAry AryPtr(InData), VarPtr(Data(0, 0)), 4, OldW, OldH

MeEn = MainForm.MeEnabled
For OutY = 0 To NewH - 1
    ShowProgress OutY / NewH * 100, Not MeEn
    
    InY = (OutY + 0.5) * OldH / NewH
    
    y1 = Int(InY - 0.5 + 0.00001)
    dy = InY - y1 - 0.5
    If y1 < 0 Then
        If TexMode Then
            y1 = MaxY
        Else
            y1 = 0
        End If
    End If
    
    y2 = y1
    
    y3 = y1 + 1
    If y3 > MaxY Then
        If TexMode Then
            y3 = 0
        Else
            y3 = MaxY
        End If
    End If
    
    y4 = y3
    
    For OutX = 0 To NewW - 1
        InX = (OutX + 0.5) * OldW / NewW
        

        x1 = Int(InX - 0.5 + 0.00001)
        dx = InX - x1 - 0.5
        If x1 < 0 Then
            If TexMode Then
                x1 = MaxX
            Else
                x1 = 0
            End If
        End If
    
        x2 = x1 + 1
        If x2 > MaxX Then
            If TexMode Then
                x2 = 0
            Else
                x2 = MaxX
            End If
        End If
    
        x3 = x1
    
        x4 = x2
    
    
        rgb1 = InData(x1, y1)
        rgb2 = InData(x2, y2)
        rgb3 = InData(x3, y3)
        rgb4 = InData(x4, y4)
        OutData(OutX, OutY).rgbRed = ((CDbl(rgb2.rgbRed) - rgb1.rgbRed) * dx + rgb1.rgbRed) * (1# - dy) + ((CDbl(rgb4.rgbRed) - rgb3.rgbRed) * dx + rgb3.rgbRed) * dy
        OutData(OutX, OutY).rgbGreen = ((CDbl(rgb2.rgbGreen) - rgb1.rgbGreen) * dx + rgb1.rgbGreen) * (1# - dy) + ((CDbl(rgb4.rgbGreen) - rgb3.rgbGreen) * dx + rgb3.rgbGreen) * dy
        OutData(OutX, OutY).rgbBlue = ((CDbl(rgb2.rgbBlue) - rgb1.rgbBlue) * dx + rgb1.rgbBlue) * (1# - dy) + ((CDbl(rgb4.rgbBlue) - rgb3.rgbBlue) * dx + rgb3.rgbBlue) * dy
        
        
        'OutData(OutX, OutY).rgbRed = r
        'OutData(OutX, OutY).rgbGreen = g
        'OutData(OutX, OutY).rgbBlue = B
    Next OutX
Next OutY

UnReferAry AryPtr(InData)

SwapArys AryPtr(Data), AryPtr(OutData)

ShowProgress 101
Exit Sub
'GetColor: 'gets color of point inx,iny
'
'    x1 = Int(InX - 0.5 + 0.00001)
'    y1 = Int(InY - 0.5 + 0.00001)
'    dx = InX - x1 - 0.5
'    dy = InY - y1 - 0.5
'    If x1 < 0 Then
'        If TexMode Then
'            x1 = MaxX
'        Else
'            x1 = 0
'        End If
'    End If
'    If y1 < 0 Then
'        If TexMode Then
'            y1 = MaxY
'        Else
'            y1 = 0
'        End If
'    End If
'
'    x2 = x1 + 1
'    y2 = y1
'    If x2 > MaxX Then
'        If TexMode Then
'            x2 = 0
'        Else
'            x2 = MaxX
'        End If
'    End If
'
'    x3 = x1
'    y3 = y1 + 1
'    If y3 > MaxY Then
'        If TexMode Then
'            y3 = 0
'        Else
'            y3 = MaxY
'        End If
'    End If
'
'    x4 = x2
'    y4 = y3
'
'
'    rgb1 = InData(x1, y1)
'    rgb2 = InData(x2, y2)
'    rgb3 = InData(x3, y3)
'    rgb4 = InData(x4, y4)
'    r = ((CDbl(rgb2.rgbRed) - rgb1.rgbRed) * dx + rgb1.rgbRed) * (1# - dy) + ((CDbl(rgb4.rgbRed) - rgb3.rgbRed) * dx + rgb3.rgbRed) * dy
'    g = ((CDbl(rgb2.rgbGreen) - rgb1.rgbGreen) * dx + rgb1.rgbGreen) * (1# - dy) + ((CDbl(rgb4.rgbGreen) - rgb3.rgbGreen) * dx + rgb3.rgbGreen) * dy
'    B = ((CDbl(rgb2.rgbBlue) - rgb1.rgbBlue) * dx + rgb1.rgbBlue) * (1# - dy) + ((CDbl(rgb4.rgbBlue) - rgb3.rgbBlue) * dx + rgb3.rgbBlue) * dy
'Return
Resume
eh:
Debug.Assert False
UnReferAry AryPtr(InData)
ErrRaise "vtStretchLinInterpol"
End Sub

Sub vtStretchCenter(ByRef Data() As Long, _
                    ByVal NewW As Long, _
                    ByVal NewH As Long, _
                    ByVal BackColor As Long, _
                    Optional ByVal TexMode As Boolean = False)
Dim x1 As Long, y1 As Long
Dim x2 As Long, y2 As Long
Dim fx As Long, fy As Long
Dim tx As Long, ty As Long
Dim yt As Long
Dim x As Long, y As Long
Dim OldW As Long, OldH As Long
Dim OutData() As Long

If AryDims(AryPtr(Data)) <> 2 Then
    'just fill OutData with back color
    vtStretchClear Data, NewW, NewH, BackColor
    'and exit
    Exit Sub
End If

OldW = UBound(Data, 1) + 1
OldH = UBound(Data, 2) + 1

If TexMode Then
    vtStretchTile Data, NewW, NewH, (NewW - OldW) \ 2, (NewH - OldH) \ 2
    Exit Sub
End If

x1 = (NewW - OldW) \ 2
y1 = (NewH - OldH) \ 2
x2 = x1 + OldW - 1
y2 = y1 + OldH - 1

fx = Max(x1, 0)
fy = Max(y1, 0)
tx = Min(x2, NewW - 1)
ty = Min(y2, NewH - 1)

ReDim OutData(0 To NewW - 1, 0 To NewH - 1)

For y = 0 To y1 - 1
    For x = 0 To NewW - 1
        OutData(x, y) = BackColor
    Next x
Next y

For y = fy To ty
    yt = y - y1
    For x = 0 To x1 - 1
        OutData(x, y) = BackColor
    Next x
    'For x = fx To tx
    '    OutData(x, y) = Data(x - x1, yt)
    'Next x
    CopyMemory OutData(fx, y), Data(fx - x1, yt), (tx - fx + 1) * 4
    For x = tx + 1 To NewW - 1
        OutData(x, y) = BackColor
    Next x
Next y

For y = ty + 1 To NewH - 1
    For x = 0 To NewW - 1
        OutData(x, y) = BackColor
    Next x
Next y

SwapArys AryPtr(Data), AryPtr(OutData)

End Sub

Sub vtStretchSimple(ByRef Data() As Long, _
                    ByVal NewW As Long, _
                    ByVal NewH As Long)
Dim x As Long, y As Long
Dim ty As Long
Dim OutData() As Long
Dim OldW As Long, OldH As Long
If AryDims(AryPtr(Data)) <> 2 Then
    'just fill OutData with back color
    vtStretchClear Data, NewW, NewH
    'and exit
    Exit Sub
End If

ReDim OutData(0 To NewW - 1, 0 To NewH - 1)

OldW = UBound(Data, 1) + 1
OldH = UBound(Data, 2) + 1

For y = 0 To NewH - 1
    ty = y * OldW \ NewW
    For x = 0 To NewW - 1
        OutData(x, y) = Data(x * OldW \ NewW, ty)
    Next x
Next y

SwapArys AryPtr(OutData), AryPtr(Data)
End Sub

Sub vtStretchPreserve(ByRef Data() As Long, _
                      ByVal NewW As Long, _
                      ByVal NewH As Long, _
                      Optional ByVal BackColor As Long)
Dim x As Long, y As Long
Dim OutData() As Long
Dim OldW As Long, OldH As Long
Dim CopyW As Long, CopyH As Long
If BackColor = -1 Then BackColor = ACol(2)
If AryDims(AryPtr(Data)) <> 2 Then
    'just fill OutData with back color
    vtStretchClear Data, NewW, NewH, BackColor
    'and exit
    Exit Sub
End If

OldW = UBound(Data, 1) + 1
OldH = UBound(Data, 2) + 1

ReDim OutData(0 To NewW - 1, 0 To NewH - 1)

CopyW = Min(OldW, NewW)
CopyH = Min(OldH, NewH)
For y = 0 To CopyH - 1
    CopyMemory OutData(0, y), Data(0, y), CopyW * 4
    For x = CopyW To NewW - 1
        OutData(x, y) = BackColor
    Next x
Next y
For y = CopyH To NewH - 1
    For x = 0 To NewW - 1
        OutData(x, y) = BackColor
    Next x
Next y
SwapArys AryPtr(Data), AryPtr(OutData)
End Sub

Sub vtStretchWin(ByRef Data() As Long, _
                 ByVal NewW As Long, _
                 ByVal NewH As Long, _
                 Optional ByVal StretchMode As APIStretchMode = HALFTONE)
Dim hDC As Long
Dim bmiIn As BITMAPINFO
Dim bmiOut As BITMAPINFO
Dim hDef As Long
Dim hDIB As Long
Dim ptrOutData As Long
Dim tmpData() As Long
Dim x As Long, y As Long
Dim OfcY As Long
Dim OldW As Long, OldH As Long
Dim Ret As Long
If AryDims(AryPtr(Data)) <> 2 Then
    'just fill OutData with back color
    vtStretchClear Data, NewW, NewH
    'and exit
    Exit Sub
End If

With bmiOut.bmiHeader
    .biSize = Len(bmiOut.bmiHeader)
    .biBitCount = 32
    .biWidth = NewW
    .biHeight = -NewH
    .biSizeImage = NewW * NewH * 4
    .biPlanes = 1
End With

On Error GoTo eh

hDC = CreateCompatibleDC(MainForm.MP.hDC)
If hDC = 0 Then
    Err.Raise 1111, , "Cannot create device context!"
End If

Ret = SetStretchBltMode(hDC, StretchMode)
If Ret = 0 Then
    Err.Raise 1111, , "SetStretchBltMode failed."
End If

hDIB = CreateDIBSection(hDC, bmiOut, DIB_RGB_COLORS, VarPtr(ptrOutData), 0, 0)
If hDIB = 0 Or ptrOutData = 0 Then
    Err.Raise 1111, , "Cannot create dib section."
End If
hDef = SelectObject(hDC, hDIB)
If hDef = 0 Then
    Err.Raise 1111, , "SelectObject failed!"
End If

OldW = UBound(Data, 1) + 1
OldH = UBound(Data, 2) + 1

With bmiIn.bmiHeader
    .biSize = Len(bmiIn.bmiHeader)
    .biBitCount = 32
    .biWidth = OldW
    .biHeight = -OldH
    .biSizeImage = OldW * OldH * 4
    .biPlanes = 1
End With

Ret = StretchDIBits(hDC, _
                    0, 0, _
                    NewW, NewH, _
                    0, 0, _
                    OldW, OldH, _
                    Data(0, 0), bmiIn, _
                    DIB_RGB_COLORS, _
                    SRCCOPY)
If Ret = &HFFFF& Then
    Err.Raise 1111, , "StretchDIBits failed!"
End If

ReDim Data(0 To NewW - 1, 0 To NewH - 1)

ConstructAry AryPtr(tmpData), ptrOutData, 4, NewW * NewH
For y = 0 To NewH - 1
    OfcY = y * NewW
    For x = 0 To NewW - 1
        Data(x, y) = tmpData(OfcY + x) And &HFFFFFF
    Next x
Next y
UnReferAry AryPtr(tmpData)

SelectObject hDC, hDef
DeleteObject hDIB
DeleteDC hDC


UnReferAry AryPtr(tmpData)
Exit Sub
eh:
UnReferAry AryPtr(tmpData)
If hDC <> 0 Then
    If hDef <> 0 Then
        SelectObject hDC, hDef
    End If
    If hDIB <> 0 Then
        DeleteObject hDIB
    End If
    DeleteDC hDC
End If
ErrRaise "vtStretchWin"
End Sub

Sub vtStretchClear(ByRef Data() As Long, _
                   ByRef NewW As Long, _
                   ByRef NewH As Long, _
                   Optional ByVal BackColor As Long = -1)
Dim OutX As Long, OutY As Long
ReDim Data(0 To NewW - 1, 0 To NewH - 1)
If BackColor = -1 Then BackColor = ACol(2)
If BackColor <> 0 Then
    For OutY = 0 To NewH - 1
        For OutX = 0 To NewW - 1
            Data(OutX, OutY) = BackColor
        Next OutX
    Next OutY
End If
End Sub

Sub vtStretchTile(ByRef Data() As Long, _
                  ByVal NewW As Long, _
                  ByVal NewH As Long, _
                  Optional ByVal OfcX As Long, _
                  Optional ByVal OfcY As Long)
Dim OldW As Long, OldH As Long
Dim OutData() As Long
Dim OutX As Long, OutY As Long
Dim y1 As Long
If AryDims(AryPtr(Data)) <> 2 Then
    'just fill OutData with back color
    vtStretchClear Data, NewW, NewH
    'and exit
    Exit Sub
End If

ReDim OutData(0 To NewW - 1, 0 To NewH - 1)
AryWH AryPtr(Data), OldW, OldH

OfcX = OfcX Mod OldW
If OfcX < 0 Then OfcX = OfcX + OldW
OfcY = OfcY Mod OldH
If OfcY < 0 Then OfcY = OfcY + OldH

For OutY = 0 To NewH - 1
    y1 = (OutY + OfcY) Mod OldH
    For OutX = 0 To NewW - 1
        OutData(OutX, OutY) = Data((OutX + OfcX) Mod OldW, y1)
    Next OutX
Next OutY

SwapArys AryPtr(Data), AryPtr(OutData)

End Sub

















'-------------------------User interface---------------------------
Sub LoadStretches(ByRef Combo As ComboBox, _
                  ByRef Names() As String, _
                  ByRef Descs() As String)
Const BaseResID As Long = 2462
Dim sArr() As String
Dim ID As Long, i As Long
Dim tmp As String
Dim NamesCnt As Long
ReDim Names(0 To 0)
ReDim Descs(0 To 0)
NamesCnt = 1
On Error GoTo eh1
ID = BaseResID
Do
    tmp = LoadResString(ID)
    If Len(tmp) > 0 Then
        sArr = Split(tmp, "|", 3)
        If UBound(sArr) <> 2 Then Exit Do
        i = Val(sArr(0))
        If i > NamesCnt - 1 Then
            NamesCnt = i + 1
            ReDim Preserve Names(0 To NamesCnt - 1)
            ReDim Preserve Descs(0 To NamesCnt - 1)
        End If
        Names(i) = Trim(sArr(1))
        Descs(i) = sArr(2)
    End If
    ID = ID + 1
Loop While tmp <> "<EOL>"
Combo.Clear
For i = 0 To NamesCnt - 1
    If Len(Names(i)) > 0 Then
        Combo.AddItem Names(i)
        Combo.ItemData(Combo.NewIndex) = i
    End If
Next i

Exit Sub
eh1:
Debug.Assert False
End Sub


