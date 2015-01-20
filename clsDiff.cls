VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDiff"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements clsEffect
Public WithEvents frm As frmDiff
Attribute frm.VB_VarHelpID = -1
Public frmPreview As frmEffectPreview
Dim BlockChanges As Boolean

Dim dx As Long, dy As Long
Dim DiffAmp As Variant, DataAmp As Variant 'Decimals
Dim TextureMode As Boolean
Dim Relief As Boolean


Private Sub clsEffect_SetPreviewWindow(frm As frmEffectPreview)
Set frmPreview = frm
End Sub

Friend Sub clsEffect_UnloadForm()
Set frmPreview = Nothing
Unload frm
End Sub

Friend Sub clsEffect_LoadForm(ByRef hwndPreviewWindow As Long, _
                              ByRef Left As Long, _
                              ByRef Top As Long, _
                              ByRef Right As Long, _
                              ByRef Bottom As Long)
Set frm = New frmDiff
Load frm
On Error Resume Next
Err.Clear
Left = frm.ScaleX(frm.iPreview.Left, frm.ScaleMode, vbPixels)
If Err.Number = 0 Then
    On Error GoTo 0
    Right = frm.ScaleX(frm.iPreview.Left + frm.iPreview.Width, frm.ScaleMode, vbPixels)
    Top = frm.ScaleX(frm.iPreview.Top, frm.ScaleMode, vbPixels)
    Bottom = frm.ScaleX(frm.iPreview.Top + frm.iPreview.Height, frm.ScaleMode, vbPixels)
    hwndPreviewWindow = frm.hWnd
Else
    hwndPreviewWindow = 0
End If
End Sub

Friend Sub clsEffect_ShowDialog()
frm.Show vbModal
If Len(frm.Tag) > 0 Then
    'clsEffect_UnloadForm
    Err.Raise dbCWS, "clsLinColorStretch:ShowDialog"
End If
End Sub

Friend Sub clsEffect_FormToSettings()
With frm
    Relief = .optMode(1).Value
    DiffAmp = .nmbDiffAmp.Value
    DataAmp = .nmbDataAmp.Value
    TextureMode = .chkTex.Value
    dx = .nmbX.Value
    dy = .nmbY.Value
End With
End Sub

Friend Sub clsEffect_SettingsToFrom()
With frm
    BlockChanges = True
    If Relief Then
        .optMode(1).Value = True
    Else
        .optMode(0).Value = True
    End If
    .nmbDiffAmp.Value = DiffAmp
    .nmbDataAmp.Value = DataAmp
    
    .chkTex.Value = IIf(TextureMode, vbChecked, vbUnchecked)
    .nmbX.Value = dx
    .nmbY.Value = dy
    BlockChanges = False
End With
End Sub

Friend Sub clsEffect_LoadSettings()
Relief = dbGetSettingEx("Effects\Differentiation", "ReliefMode", vbBoolean, False)
DiffAmp = dbGetSettingEx("Effects\Differentiation", "DiffAmp", vbDecimal, 1)
DataAmp = dbGetSettingEx("Effects\Differentiation", "DataAmp", vbDecimal, 1)
TextureMode = dbGetSettingEx("Effects\Differentiation", "TextureMode", vbBoolean, True)
dx = dbGetSettingEx("Effects\Differentiation", "dx", vbLong, -3)
dy = dbGetSettingEx("Effects\Differentiation", "dy", vbLong, -2)
End Sub

Friend Sub clsEffect_SaveSettings()
 dbSaveSettingEx "Effects\Differentiation", "ReliefMode", Relief
 dbSaveSettingEx "Effects\Differentiation", "DiffAmp", DiffAmp
 dbSaveSettingEx "Effects\Differentiation", "DataAmp", DataAmp
 dbSaveSettingEx "Effects\Differentiation", "TextureMode", TextureMode
 dbSaveSettingEx "Effects\Differentiation", "dx", dx
 dbSaveSettingEx "Effects\Differentiation", "dy", dy
End Sub

Private Sub clsEffect_GetEffectDesc(ByRef IDName As String, _
                                    ByRef MenuName As String)
MenuName = "Differentiation"
IDName = "Diff"
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
Dim Offset As POINTAPI
If Relief Then
    Offset.x = dx
    Offset.y = dy
    vtDifferentiate InData, OutData, Range, (DiffAmp), (DataAmp), Offset, TextureMode
Else
    Offset.x = 0
    Offset.y = 0
    vtDifferentiate InData, OutData, Range, (DiffAmp), kData:=0, Offset:=Offset, TextureMode:=TextureMode
End If
End Sub

Private Sub Class_Initialize()
clsEffect_LoadSettings
End Sub

Private Sub Class_Terminate()
Unload frm
clsEffect_SaveSettings
End Sub

Private Sub Frm_Change()
If BlockChanges Then Exit Sub
frmPreview.Update True
End Sub


