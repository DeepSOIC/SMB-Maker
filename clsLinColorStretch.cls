VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLinColorStretch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements clsEffect
Public WithEvents Frm As frmGain
Attribute Frm.VB_VarHelpID = -1
Public frmPreview As frmEffectPreview

Dim MulFactor As Double

Private Sub clsEffect_SetPreviewWindow(Frm As frmEffectPreview)
Set frmPreview = Frm
End Sub

Friend Sub clsEffect_UnloadForm()
Set frmPreview = Nothing
Unload Frm
End Sub

Friend Sub clsEffect_LoadForm(ByRef hwndPreviewWindow As Long, _
                              ByRef Left As Long, _
                              ByRef Top As Long, _
                              ByRef Right As Long, _
                              ByRef Bottom As Long)
Set Frm = New frmGain
Load Frm
On Error Resume Next
Err.Clear
Left = Frm.ScaleX(Frm.iPreview.Left, Frm.ScaleMode, vbPixels)
If Err.Number = 0 Then
    On Error GoTo 0
    Right = Frm.ScaleX(Frm.iPreview.Left + Frm.iPreview.Width, Frm.ScaleMode, vbPixels)
    Top = Frm.ScaleX(Frm.iPreview.Top, Frm.ScaleMode, vbPixels)
    Bottom = Frm.ScaleX(Frm.iPreview.Top + Frm.iPreview.Height, Frm.ScaleMode, vbPixels)
    hwndPreviewWindow = Frm.hWnd
Else
    hwndPreviewWindow = 0
End If
End Sub

Friend Sub clsEffect_ShowDialog()
Frm.Show vbModal
If Len(Frm.Tag) > 0 Then
    'clsEffect_UnloadForm
    Err.Raise dbCWS, "clsLinColorStretch:ShowDialog"
End If
End Sub

Friend Sub clsEffect_FormToSettings()
With Frm
    MulFactor = .GetFactor
End With
End Sub

Friend Sub clsEffect_SettingsToFrom()
With Frm
    .SetFactor MulFactor
End With
End Sub

Friend Sub clsEffect_LoadSettings()
MulFactor = dbGetSettingEx("Effects\ColorStretch", "Factor", vbDouble, 0.75)
End Sub

Friend Sub clsEffect_SaveSettings()
dbSaveSettingEx "Effects\ColorStretch", "Factor", MulFactor
End Sub

Private Sub clsEffect_GetEffectDesc(ByRef IDName As String, _
                                    ByRef MenuName As String)
MenuName = "Multiplication"
IDName = "Gamma"
End Sub


Friend Sub clsEffect_PerformEffect(ByRef InData() As Long, _
                                   ByRef OutData() As Long, _
                                   ByRef RangeL As Long, _
                                   ByRef RangeT As Long, _
                                   ByRef RangeR As Long, _
                                   ByRef RangeB As Long)
Dim Range As RECT
Range.Left = RangeL
Range.Top = RangeT
Range.Right = RangeR
Range.Bottom = RangeB
dbApplyGammaEx InData, MulFactor, Range, OutData
End Sub


Friend Sub dbApplyGammaEx(ByRef InData() As Long, _
                          ByVal mnoj As Double, _
                          ByRef Range As RECT, _
                          ByRef OutData() As Long)
Dim tbl() As Byte
Dim i As Long
Dim tmp As Long
ReDim tbl(0 To 255)
For i = 0 To 255
    tmp = i * mnoj
    If tmp > 255& Then tmp = 255&
    tbl(i) = tmp
Next i
dbMapColorsEx InData, tbl, tbl, tbl, Range, OutData
End Sub

Function vtCountMaxMultiplier(ByRef InData() As Long) As Double
Dim x As Long, y As Long
Dim Res As Byte 'Stores the maximum
Dim tmp As Long
Dim w As Long, h As Long
Dim RGBData() As RGBQUAD
Dim OfcY As Long
Dim x1 As Long, x2 As Long
Dim y1 As Long, y2 As Long

If AryDims(AryPtr(InData)) <> 2 Then
    Err.Raise 1111, "vtCountMaxMultiplier", "A bidimensional array is required!"
End If
w = UBound(InData, 1) + 1
h = UBound(InData, 2) + 1
    
x1 = 0
y1 = 0
x2 = w - 1
y2 = h - 1

On Error GoTo eh
ConstructAry AryPtr(RGBData), VarPtr(InData(0, 0)), 4, w * h

Res = 0
For y = y1 To y2
    OfcY = y * w
    For x = x1 To x2
        If Res < RGBData(OfcY + x).rgbRed Then Res = RGBData(OfcY + x).rgbRed
        If Res < RGBData(OfcY + x).rgbGreen Then Res = RGBData(OfcY + x).rgbGreen
        If Res < RGBData(OfcY + x).rgbBlue Then Res = RGBData(OfcY + x).rgbBlue
        If Res = 255 Then Exit For
    Next x
    If Res = 255 Then Exit For
Next y

If Res = 0 Then
    vtCountMaxMultiplier = -1
Else
    vtCountMaxMultiplier = 255# / CDbl(Res)
End If

UnReferAry AryPtr(RGBData)

Exit Function
eh:
UnReferAry AryPtr(RGBData)
ErrRaise "vtCountMaxMultiplier"
End Function


Private Sub Class_Initialize()
clsEffect_LoadSettings
End Sub

Private Sub Class_Terminate()
Unload Frm
clsEffect_SaveSettings
End Sub

Private Sub Frm_Change()
frmPreview.Update True
End Sub
