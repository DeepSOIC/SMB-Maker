Attribute VB_Name = "mdlGraphicsOutput"
Option Explicit

Public Sub vtSetDIBitsToDevice(ByVal hDC As Long, _
                               ByRef Data() As Long, _
                               ByVal SrcX As Long, ByVal SrcY As Long, _
                               ByVal X As Long, ByVal Y As Long, _
                               ByVal w As Long, ByVal h As Long)
Dim bmi As BITMAPINFO
Dim DataW As Long, DataH As Long
Dim Range As RECT
Dim Ret As Long
AryWH AryPtr(Data), DataW, DataH
Range.Left = SrcX
Range.Top = SrcY
Range.Right = Range.Left + w
Range.Bottom = Range.Top + h
If Range.Left < 0 Then
    X = X + -Range.Left
    Range.Right = Range.Right + -Range.Left
    Range.Left = Range.Left + -Range.Left
End If
If Range.Top < 0 Then
    Y = Y + -Range.Top
    Range.Bottom = Range.Bottom + -Range.Top
    Range.Top = Range.Top + -Range.Top
End If
If Range.Right > DataW Then
    Range.Right = DataW
End If
If Range.Bottom > DataH Then
    Range.Bottom = DataH
End If

With bmi.bmiHeader
    .biSize = Len(bmi.bmiHeader)
    .biPlanes = 1
    .biBitCount = 32
    .biWidth = DataW
    .biHeight = -DataH
    .biSizeImage = DataW * DataH * 4
End With

Ret = SetDIBitsToDevice(hDC, _
                  X, Y, _
                  Range.Right - Range.Left, _
                  Range.Bottom - Range.Top, _
                  Range.Left, DataH - Range.Bottom, _
                  0, DataH, _
                  Data(0, 0), bmi, DIB_RGB_COLORS)
Debug.Assert Ret = DataH
End Sub

'this sub is obsolete in this module
'and is used only by other modules (e.g. frmSoft)
'Grid is not supported any more
'ImageHandle argument is provided for compatibility purpose
'   and is not used at all
Public Sub RefrEx(ByVal ImageHandle As Long, _
                  ByVal hDC As Long, _
                  ByRef pData() As Long, _
                  Optional ByVal Zm As Integer = 1, _
                  Optional ByVal Grid As GREnum = dbAsmnuGrid)
Dim bmi As BITMAPINFO
Dim w As Long, h As Long
Dim hBitmap As Long, hDefBitmap As Long
Dim tmpDC As Long
Dim ptrBits As Long
Dim Ret As Long
If AryDims(AryPtr(pData)) <> 2 Then
    Err.Raise 1111, "RefrEx", "Incorrect number of dimensions in the array!"
End If
'comon assignments
If Zm < 0 Then Zm = -Zm
If Zm = 0 Then Err.Raise 1111, "RefrEx", "Zoom ratio cannot be zero."
w = UBound(pData, 1) + 1
h = UBound(pData, 2) + 1
With bmi.bmiHeader
    .biSize = Len(bmi.bmiHeader)
    .biBitCount = 32
    .biPlanes = 1
    .biWidth = w
    .biHeight = -h
    .biSizeImage = .biWidth * Abs(.biHeight) * 4
End With

'zm=1 is simlper than zm>1
If Zm = 1 Then
    SetDIBitsToDevice hDC, 0, 0, w, h, 0, 0, 0, h, pData(0, 0), bmi, DIB_RGB_COLORS
Else
    hBitmap = CreateDIBSection(hDC, bmi, 0, VarPtr(ptrBits), 0, 0)
    If hBitmap = 0 Then
        Err.Raise 1111, "RefrEx", "Failed to create the dib section."
    End If
    On Error GoTo eh
    'create temporary DC
    tmpDC = CreateCompatibleDC(hDC)
    If tmpDC = 0 Then
        Err.Raise 1111, "RefrEx", "CreateCompatibleDC failed!"
    End If
    'select DIB into the tmpDC
    hDefBitmap = SelectObject(tmpDC, hBitmap)
    If hDefBitmap = 0 Then
        Err.Raise 1111, "RefrEx", "SelectObject failed!"
    End If
    'Copy bits
    CopyMemory ByVal ptrBits, pData(0, 0), w * h * 4
    'draw them on the destination
    Ret = StretchBlt(hDC, _
                     0, 0, w * Zm, h * Zm, _
                     tmpDC, _
                     0, 0, w, h, _
                     VBRUN.RasterOpConstants.vbSrcCopy)
    If Ret = 0 Then
        Err.Raise 1111, "RefrEx", "StretchBlt Failed!"
    End If
    'free resources
    SelectObject tmpDC, hDefBitmap
    DeleteObject hBitmap
    DeleteDC tmpDC
End If

Exit Sub
eh:
If hDefBitmap <> 0 Then
    SelectObject tmpDC, hDefBitmap
End If
If hBitmap <> 0 Then
    DeleteObject hBitmap
End If
If tmpDC <> 0 Then
    DeleteDC tmpDC
End If
Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Sub vtStretchDIBits(ByVal hDC As Long, _
                           ByVal ptrData As Long, _
                           ByVal X As Long, ByVal Y As Long, _
                           ByVal Width As Long, ByVal Height As Long, _
                           Optional ByVal StretchMode As APIStretchMode = APIStretchMode.COLORONCOLOR)
Dim Data() As Long
Dim w As Long, h As Long
Dim ptr As Long
'ReferAry AryPtr(Data), ptrData
On Error GoTo eh
If AryDims(ptrData) = 0 Then
  Exit Sub
End If
If AryDims(ptrData) <> 2 Then
  Err.Raise 1212, , "Incorrect array dimension!"
End If
AryWH ptrData, w, h

On Error Resume Next
SwapArys AryPtr(Data), ptrData
ptr = VarPtr(Data(0, 0))
SwapArys AryPtr(Data), ptrData
On Error GoTo eh
If ptr = 0 Then Err.Raise 1212, , "Failed to get pointer to array data!"

Dim bmi As BITMAPINFO
Dim PrevSM As Long
Dim Ret As Long
With bmi.bmiHeader
    .biSize = Len(bmi.bmiHeader)
    .biBitCount = 32
    .biPlanes = 1
    .biWidth = w
    .biHeight = -h
    .biSizeImage = .biWidth * Abs(.biHeight) * 4
End With

PrevSM = SetStretchBltMode(hDC, StretchMode)

Ret = StretchDIBits(hDC, X, Y, Width, Height, 0, 0, w, h, ByVal ptr, bmi, DIB_RGB_COLORS, SRCCOPY)
If Ret = &HFFFF& Then Err.Raise 1213, , "StretchDIBits failed. Dll error = " + CStr(Err.LastDllError)

If PrevSM <> 0 Then
  SetStretchBltMode hDC, PrevSM
End If

Exit Sub
eh:
'PushError
'UnReferAry AryPtr(Data)
'PopError
ErrRaise "vtStretchDIBits"
End Sub


Public Sub vtGetDIBitsFromDevice(ByRef Data() As Long, _
                                 ByVal hDC As Long, _
                                 Optional ByVal ClearReserved As Boolean = True, _
                                 Optional ByVal StartY As Long = -1)
Dim tmpBitmap As Long
Dim bmp As Long

tmpBitmap = CreateCompatibleBitmap(hDC, 1, 1)
bmp = SelectObject(hDC, tmpBitmap)
On Error GoTo eh
If bmp = 0 Then
    Err.Raise 1212, , "The HDC has no image. Or SelectObject failed."
End If

vtGetDIBitsFromImage Data, bmp, hDC, ClearReserved, StartY

SelectObject hDC, bmp
DeleteObject tmpBitmap
Exit Sub
eh:
PushError
If bmp <> 0 Then
    SelectObject hDC, bmp
End If
If tmpBitmap <> 0 Then
    DeleteObject tmpBitmap
End If
PopError
ErrRaise "vtGetDIBitsFromDevice"
End Sub

Public Sub vtGetDIBitsFromImage(ByRef Data() As Long, _
                                ByVal hImage As Long, _
                                ByVal hDC As Long, _
                                Optional ByVal ClearReserved As Boolean = True, _
                                Optional ByVal StartY As Long = -1)
Dim bmi As BITMAPINFO
Dim w As Long, h As Long
Dim dw As Long, dh As Long
Dim Ret As Long
Dim ScanLines As Long

With bmi.bmiHeader
    .biSize = Len(bmi.bmiHeader)
End With
Ret = APIGetDiBits(hDC, hImage, 0, 0, ByVal 0, bmi, DIB_RGB_COLORS)
If Ret = 0 Then
    Err.Raise 1212, "vtGetDIBitsFromImage", "GetDIBIts for examining the dimensions failed!"
End If

With bmi.bmiHeader
    w = .biWidth
    h = Abs(.biHeight)
    .biBitCount = 32
    .biSizeImage = w * h * 4&
    .biPlanes = 1
    .biHeight = -h
    .biCompression = 0
End With
ScanLines = h

AryWH AryPtr(Data), dw, dh
If dw = w And StartY <> -1 Then
  If StartY < 0 Then Err.Raise "StartY cannot be <0 (-1 is exception)."
  If StartY + h > dh Then Err.Raise 2342, , "Not enough space in data!"  ' ScanLines = dh - StartY
Else
  If StartY <> -1 Then Err.Raise 2342, "vtGetDIBitsFromImage", "If you use StartY, you must allocate data and it's width must match."
  ReDim Data(0 To w - 1, 0 To h - 1)
  StartY = 0
End If

Ret = APIGetDiBits(hDC, hImage, 0, ScanLines, Data(0, StartY), bmi, DIB_RGB_COLORS)
If Ret = 0 Then
    Err.Raise 1212, , "Getting bits failed!"
End If

If ClearReserved Then
    vtRepair Data, StartY, h
End If

End Sub

Public Sub dbGetDIBits(ByVal hImage As Long, _
                       ByVal hDC As Long, _
                       ByRef Data() As Long, _
                       Optional ByVal ClearReserved As Boolean = True, _
                       Optional ByVal StartY As Long = -1)
If GetCurrentObject(hDC, OBJ_BITMAP) <> hImage Then
    vtGetDIBitsFromImage Data, hImage, hDC, ClearReserved, StartY
Else
    vtGetDIBitsFromDevice Data, hDC, ClearReserved, StartY
End If
'Dim bmi As BITMAPINFO, tmpData() As RGBQUAD
'Dim i As Long, j As Long
'Dim w As Long, h As Long
'Dim ind As Long
'Dim x As Long, y As Long
'Dim tmpBitmap As Long
'
'    With bmi.bmiHeader
'        .biBitCount = 0
'        .biClrImportant = 0
'        .biClrUsed = 0
'        .biCompression = 0
'        .biPlanes = 1
'        .biSize = Len(bmi.bmiHeader)
'        .biXPelsPerMeter = 0
'        .biYPelsPerMeter = 0
'    End With
'    If APIGetDiBits(hDC, pichandle, 0, 0, ByVal 0, bmi, DIB_RGB_COLORS) = 0 Then
'        Err.Raise 1001, "dbGetDIBits", "APIGetDIBits failed. Dll error: " + CStr(Err.LastDllError) + "."
'    End If
'    w = bmi.bmiHeader.biWidth
'    h = bmi.bmiHeader.biHeight
'    If w = 0 Or h = 0 Then Err.Raise 1001, "GetDIBits", "No picture"
'    ReDim tmpData(0 To w * h - 1)
'    With bmi.bmiHeader
'        .biBitCount = 32
'        .biSizeImage = 4& * w * h
'        .biHeight = -Abs(h)
'    End With
'    ReDim gData(0 To w - 1, 0 To h - 1)
'    If APIGetDiBits(hDC, pichandle, 0, h, gData(0, 0), bmi, DIB_RGB_COLORS) = 0 Then
'        Err.Raise 1001, "dbGetDIBits", "APIGetDIBits failed. Dll error: " + CStr(Err.LastDllError) + "."
'    End If
'
'    vtRepair gData
'
'    ind = 0
'    Exit Sub
'VBHandler:
'    For i = 0 To h - 1
'        For j = 0 To w - 1
'            gData(j, h - 1 - i) = RGB(tmpData(ind).rgbRed, tmpData(ind).rgbGreen, tmpData(ind).rgbBlue)
'            ind = ind + 1&
'        Next j
'    Next i
'    Erase tmpData
'Exit Sub
'FailDll:
'    If DllPresent Then
'        MsgBox Err.Description
'    End If
'    Resume VBHandler
End Sub

'transforms iPictureDisp to Data and Alpha.
'if aryptralpha not specified, alpha channel will be _
  embedded into data, otherwise written into specified array.
'if alpha is full opaque, Alpha is erased.
'returns true if alpha is present.
Public Function GetPicData(ByRef Pic As IPictureDisp, _
                           ByRef Data() As Long, _
                           Optional ByVal CalcAlpha As Boolean = False, _
                           Optional ByVal aryptrAlpha As Long = 0) As Boolean
Dim AlphaPresent As Boolean
Dim BlackData() As Long
Dim WhiteData() As Long
Dim w As Long, h As Long
Dim i As Long, iShift As Long '=startY*w
Dim hDIB As Long
Dim frm As New frmFormatVB
Dim AlphaAccu As Long
Dim a As Long
Dim RGBBlackData() As RGBQUAD
Dim RGBWhiteData() As RGBQUAD
Dim r As Long, g As Long, b As Long
'Dim CalcAlpha As Boolean
Dim th As Long 'height of processing block, w*h1<2Mpixel
Dim h1 As Long 'height of current processing block (can be <th for last block)
Dim StartY As Long 'first scanline of processing block
Dim AlphaVector() As Long
Dim Alpha() As Long

Load frm
On Error GoTo eh
Dim Sz As Size
'GetBitmapDimensionEx Pic.Handle, Sz
w = frm.Pic.ScaleX(Pic.Width, vbHimetric, vbPixels)
h = frm.Pic.ScaleY(Pic.Height, vbHimetric, vbPixels)
'CalcAlpha = (w * h <= 1280& * 1024&) And (aryptrAlpha <> 0)
If w * h = 0 Then
  Erase Data
  If aryptrAlpha <> 0 Then
    SwapArys aryptrAlpha, AryPtr(Alpha)
    Erase Alpha
  End If
  Unload frm
  Exit Function
End If

If w * h > 2000000 Then
  th = 2000000 \ w
  If th < 1 Then th = 1
Else
  th = h
End If

With frm.Pic
    .Cls
    .Move 0, 0, w, th
    .BackColor = vbBlack
    .Cls
End With

ReDim BlackData(0 To w - 1, 0 To h - 1)
ConstructAry AryPtr(RGBBlackData), VarPtr(BlackData(0, 0)), 4, w * h
If CalcAlpha Then
  ReDim WhiteData(0 To w - 1, 0 To th - 1)
  ConstructAry AryPtr(RGBWhiteData), VarPtr(WhiteData(0, 0)), 4, w * h
  If aryptrAlpha <> 0 Then
    SwapArys aryptrAlpha, AryPtr(Alpha)
    ReDim Alpha(0 To w - 1, 0 To h - 1)
    ConstructAry AryPtr(AlphaVector), VarPtr(Alpha(0, 0)), 4, w * h
  End If
  AlphaAccu = &HFF&
End If
For StartY = 0 To h - 1 Step th
  iShift = StartY * w
  h1 = Min(th, h - StartY)
  With frm.Pic
    If h1 < h Then
      .Cls
      .Move 0, 0, w, h1
      .Cls
    End If
    .BackColor = vbBlack
    frm.Pic.Line (0, 0)-(w, h1), .BackColor, BF
    .PaintPicture Pic, 0, -StartY
    vtGetDIBitsFromDevice BlackData, .hDC, ClearReserved:=Not CalcAlpha, StartY:=StartY
    If CalcAlpha Then
      .BackColor = vbWhite
      frm.Pic.Line (0, 0)-(w, th), .BackColor, BF
      .PaintPicture Pic, 0, -StartY
      vtGetDIBitsFromDevice WhiteData, .hDC, ClearReserved:=False, StartY:=0
    End If
  End With
  If CalcAlpha Then
      If aryptrAlpha <> 0 Then
        Debug.Assert iShift + h1 * w <= h * w
        For i = iShift To Min(iShift + h1 * w, h * w) - 1
            a = 255& - (CLng(RGBWhiteData(i - iShift).rgbBlue) + RGBWhiteData(i - iShift).rgbGreen + RGBWhiteData(i - iShift).rgbRed _
                        - RGBBlackData(i).rgbBlue - RGBBlackData(i).rgbGreen - RGBBlackData(i).rgbRed) \ 3&
            If a > 255 Then a = 255
            'a is the opacity
            AlphaAccu = AlphaAccu And a
            If a > 0& Then
                b = RGBBlackData(i).rgbBlue * 255& / a
                If b And &HFFFFFF00 Then b = &HFF
                RGBBlackData(i).rgbBlue = b
                
                g = RGBBlackData(i).rgbGreen * 255& / a
                If g And &HFFFFFF00 Then g = &HFF
                RGBBlackData(i).rgbGreen = g
                
                r = RGBBlackData(i).rgbRed * 255& / a
                If r And &HFFFFFF00 Then r = &HFF
                RGBBlackData(i).rgbRed = r
            Else
                a = 0&
            End If
            RGBBlackData(i).rgbReserved = 0
            AlphaVector(i) = a * &H10101
        Next i
      Else
        For i = iShift To Min(iShift + h1 * w, h * w) - 1
            a = 255& - (CLng(RGBWhiteData(i - iShift).rgbBlue) + RGBWhiteData(i - iShift).rgbGreen + RGBWhiteData(i - iShift).rgbRed _
                        - RGBBlackData(i).rgbBlue - RGBBlackData(i).rgbGreen - RGBBlackData(i).rgbRed) \ 3&
            If a > 255 Then a = 255
            'a is the opacity
            AlphaAccu = AlphaAccu And a
            If a > 0& Then
                b = RGBBlackData(i).rgbBlue * 255& / a
                If b And &HFFFFFF00 Then b = &HFF
                RGBBlackData(i).rgbBlue = b
                
                g = RGBBlackData(i).rgbGreen * 255& / a
                If g And &HFFFFFF00 Then g = &HFF
                RGBBlackData(i).rgbGreen = g
                
                r = RGBBlackData(i).rgbRed * 255& / a
                If r And &HFFFFFF00 Then r = &HFF
                RGBBlackData(i).rgbRed = r
            Else
                a = 0&
            End If
            RGBBlackData(i).rgbReserved = a
        Next i
      End If
  End If
Next StartY
If CalcAlpha Then
  GetPicData = Not (AlphaAccu = &HFF)
  If AlphaAccu = &HFF Then
    If aryptrAlpha <> 0 Then
      Erase Alpha
    Else
      For i = 0 To w * h - 1
        RGBBlackData(i).rgbReserved = 0
      Next i
    End If
  End If
  If aryptrAlpha <> 0 Then
    SwapArys aryptrAlpha, AryPtr(Alpha)
  End If
End If
UnReferAry AryPtr(AlphaVector)
UnReferAry AryPtr(RGBBlackData)
UnReferAry AryPtr(RGBWhiteData)
Erase WhiteData
SwapArys AryPtr(BlackData), AryPtr(Data)
Unload frm
Exit Function
eh:
    PushError
    UnReferAry AryPtr(AlphaVector)
    UnReferAry AryPtr(RGBBlackData)
    UnReferAry AryPtr(RGBWhiteData)
    Unload frm
    PopError
ErrRaise "GetPicData"

End Function

