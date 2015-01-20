VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsReRGB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements clsEffect
Public WithEvents Frm As frmReRGB
Attribute Frm.VB_VarHelpID = -1
Public frmPreview As frmEffectPreview
Dim SettingsLoaded As Boolean

Dim Matrix() As Double

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
Set Frm = New frmReRGB
Load Frm
On Error Resume Next
Err.Clear
Left = Frm.ScaleX(Frm.iPreview.Left, Frm.ScaleMode, vbPixels)
If Err.Number = 0 Then
    On Error GoTo 0
    Right = Frm.ScaleX(Frm.iPreview.Left + Frm.iPreview.Width, Frm.ScaleMode, vbPixels)
    Top = Frm.ScaleX(Frm.iPreview.Top, Frm.ScaleMode, vbPixels)
    Bottom = Frm.ScaleX(Frm.iPreview.Top + Frm.iPreview.Height, Frm.ScaleMode, vbPixels)
    hwndPreviewWindow = Frm.hwnd
Else
    hwndPreviewWindow = 0
End If
End Sub

Friend Sub clsEffect_ShowDialog()
Frm.Show vbModal
If Len(Frm.Tag) > 0 Then
    'clsEffect_UnloadForm
    Err.Raise dbCWS, "clsReRGB:ShowDialog"
End If
End Sub

Friend Sub clsEffect_FormToSettings()
With Frm
    .GetMatrix Matrix
    .tmrChanger.Enabled = False
End With
End Sub

Friend Sub clsEffect_SettingsToFrom()
If Not SettingsLoaded Then clsEffect_LoadSettings
With Frm
    .SetMatrix Matrix
End With
End Sub

Friend Sub clsEffect_LoadSettings()
Dim i As Long
LoadMatrix Matrix, "Effects\ReRGB", "Matrix"
SettingsLoaded = True
If AryDims(AryPtr(Matrix)) = 2 Then
    If LBound(Matrix, 1) = 0 And UBound(Matrix, 1) = 2 And _
       LBound(Matrix, 2) = 0 And UBound(Matrix, 2) = 2 Then
        Exit Sub
    End If
End If
ReDim Matrix(0 To 2, 0 To 2)
For i = 0 To 2
    Matrix(i, i) = 1
Next i
End Sub

Friend Sub clsEffect_SaveSettings()
Dim MaskData() As Byte
Dim MaskW As Long, MaskH As Long
If Not SettingsLoaded Then Exit Sub
SaveMatrix Matrix, "Effects\ReRGB", "Matrix"
End Sub

Private Sub clsEffect_GetEffectDesc(ByRef IDName As String, _
                                    ByRef MenuName As String)
MenuName = "Color Matrix"
IDName = "ReRGB"
End Sub


Friend Sub clsEffect_PerformEffect(ByRef InData() As Long, _
                                   ByRef OutData() As Long, _
                                   ByRef RangeL As Long, _
                                   ByRef RangeT As Long, _
                                   ByRef RangeR As Long, _
                                   ByRef RangeB As Long)
Dim Range As RECT
If Not SettingsLoaded Then clsEffect_LoadSettings
Range.Left = RangeL
Range.Top = RangeT
Range.Right = RangeR
Range.Bottom = RangeB
dbMatrixEx InData, Matrix, OutData, Range
End Sub


Private Sub Class_Initialize()
'clsEffect_LoadSettings
End Sub

Private Sub Class_Terminate()
Unload Frm
clsEffect_SaveSettings
End Sub

Private Sub Frm_Change()
On Error GoTo eh
frmPreview.Update True
Exit Sub
eh:
ShowStatus Err.Source + ":" + Err.Description
vtBeep
End Sub



