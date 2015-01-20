VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements clsEffect
Public WithEvents frm As frmSoft
Attribute frm.VB_VarHelpID = -1
Public frmPreview As frmEffectPreview
Dim SettingsLoaded As Boolean

Dim Settings As typFilterOpts
Dim Mask As FilterMask

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
Set frm = New frmSoft
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
    Err.Raise dbCWS, "clsFilter:ShowDialog"
End If
End Sub

Friend Sub clsEffect_FormToSettings()
With frm
    Settings.FilterMode = .GetFilterMode
    Settings.Brightness = dBtoFactor(.nmbGain.Value)
    Settings.TextureMode = CBool(.chkTexMode.Value)
    .ExtractData2 Mask
    .tmrUpdater.Enabled = False
End With
End Sub

Friend Sub clsEffect_SettingsToFrom()
If Not SettingsLoaded Then clsEffect_LoadSettings
With frm
    .SetFilterMode Settings.FilterMode
    .nmbGain.Value = FactorToDB(Settings.Brightness)
    .chkTexMode.Value = Abs(Settings.TextureMode)
    .SetData2 Mask
    Settings.Absolute = Settings.FilterMode = fmDifference
End With
End Sub

Friend Sub clsEffect_LoadSettings()
Dim MaskData() As Byte
Dim MaskW As Long, MaskH As Long
Dim l As Long
Settings.Brightness = dbGetSettingEx("Effects\Filtering", "Brightness", vbDouble, 1#)
Settings.FilterMode = dbGetSettingEx("Effects\Filtering", "FilterMode", vbLong, eFilterMode.fmNormal)
Settings.TextureMode = dbGetSettingEx("Effects\Filtering", "TextureMode", vbBoolean, False)
Settings.Absolute = Settings.FilterMode = fmDifference

dbGetSettingBin "Effects\Filtering", "LastUsedMask", MaskData
If AryDims(AryPtr(MaskData)) <> 1 Then
    ReDim MaskData(0 To 4 + 3& * 3 * 8 * 3 - 1)
    MaskData(0) = 3
    MaskData(2) = 3
    MaskData(4 + 4 * 3 * 0) = 0

    MaskData(4 + 4 * 3 * 1) = 255
    MaskData(4 + 4 * 3 * 1 + 4) = 255
    MaskData(4 + 4 * 3 * 1 + 8) = 255

    CopyMemory MaskData(4 + 4 * 3 * 3), MaskData(4 + 4 * 3 * 1), 4 * 3
    CopyMemory MaskData(4 + 4 * 3 * 4), MaskData(4 + 4 * 3 * 1), 4 * 3
    CopyMemory MaskData(4 + 4 * 3 * 5), MaskData(4 + 4 * 3 * 1), 4 * 3
    CopyMemory MaskData(4 + 4 * 3 * 7), MaskData(4 + 4 * 3 * 1), 4 * 3
End If
MaskW = MaskData(0)
MaskH = MaskData(2)
If MaskW * MaskH <= 0 Then
    Erase Mask.Mask
    Mask.CenterFilled = False
Else
    ReDim Mask.Mask(0 To MaskW - 1, 0 To MaskH - 1)
    l = Min(MaskW * MaskH * 4 * 3, UBound(MaskData) + 1 - 4)
    If l > 0 Then
        CopyMemory Mask.Mask(0, 0), MaskData(4), l
    End If
    Mask.CenterFilled = False
    Mask.Center.x = -1
    Mask.Center.y = -1
End If
SettingsLoaded = True
End Sub

Friend Sub clsEffect_SaveSettings()
Dim MaskData() As Byte
Dim MaskW As Long, MaskH As Long
If Not SettingsLoaded Then Exit Sub
dbSaveSettingEx "Effects\Filtering", "Brightness", Settings.Brightness
dbSaveSettingEx "Effects\Filtering", "FilterMode", Settings.FilterMode
dbSaveSettingEx "Effects\Filtering", "TextureMode", Settings.TextureMode

AryWH AryPtr(Mask.Mask), MaskW, MaskH
ReDim MaskData(0 To (4 + MaskW * MaskH * 4 * 3) - 1)
MaskData(0) = MaskW
MaskData(2) = MaskH
If MaskW * MaskH > 0 Then
    CopyMemory MaskData(4), Mask.Mask(0, 0), MaskW * MaskH * 4 * 3
End If
dbSaveSettingBin "Effects\Filtering", "LastUsedMask", MaskData
End Sub

Private Sub clsEffect_GetEffectDesc(ByRef IDName As String, _
                                    ByRef MenuName As String)
MenuName = "Filtering"
IDName = "Filtering"
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
vtFilter InData, Mask, Settings, Range, OutData
End Sub


Private Sub Class_Initialize()
'clsEffect_LoadSettings
End Sub

Private Sub Class_Terminate()
Unload frm
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


