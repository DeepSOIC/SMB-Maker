VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsColorMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements clsEffect
Public WithEvents frm As frmContrast
Attribute frm.VB_VarHelpID = -1
Public frmPreview As frmEffectPreview

Dim Graphs() As dbGraph

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
Set frm = New frmContrast
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
    Err.Raise dbCWS, "clsColorMap:ShowDialog"
End If
End Sub

Friend Sub clsEffect_FormToSettings()
With frm
    .GetGraphs Graphs
'    .tmrChanger.Enabled = False
End With
End Sub

Friend Sub clsEffect_SettingsToFrom()
With frm
    .SetGraphs Graphs
    .gRefr
End With
End Sub

Friend Sub clsEffect_LoadSettings()
Dim MsgText As String, Answ As VbMsgBoxResult
On Error GoTo eh
    ReDim Graphs(1 To 3)
    Answ = 0
    ExtractGraphString dbGetSetting("Effects\Graph", "LastRedGraph", "00 00 FF FF0"), Graphs(1)
    If Answ = vbIgnore Then InitGraph Graphs(1).Points
    Answ = 0
    ExtractGraphString dbGetSetting("Effects\Graph", "LastGreenGraph", "00 00 FF FF0"), Graphs(2)
    If Answ = vbIgnore Then InitGraph Graphs(2).Points
    Answ = 0
    ExtractGraphString dbGetSetting("Effects\Graph", "LastBlueGraph", "00 00 FF FF0"), Graphs(3)
    If Answ = vbIgnore Then InitGraph Graphs(3).Points
Exit Sub
eh:
Answ = dbMsgBox(MsgText + vbCrLf + Err.Description + "`Error", vbCritical Or vbAbortRetryIgnore)
Select Case Answ
    Case vbAbort
        End
    Case vbIgnore
'        ShowAllButtons
        Resume Next
    Case vbRetry
'        ShowAllButtons
        Resume
End Select
End Sub

Friend Sub clsEffect_SaveSettings()
dbSaveSetting "Effects\Graph", "LastRedGraph", BuildGraphString(Graphs(1))
dbSaveSetting "Effects\Graph", "LastGreenGraph", BuildGraphString(Graphs(2))
dbSaveSetting "Effects\Graph", "LastBlueGraph", BuildGraphString(Graphs(3))
End Sub

Private Sub clsEffect_GetEffectDesc(ByRef IDName As String, _
                                    ByRef MenuName As String)
MenuName = "Color graph"
IDName = "Graph"
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
dbMapColorsEx InData, _
              Graphs(1).Table, _
              Graphs(2).Table, _
              Graphs(3).Table, _
              Range, _
              OutData
End Sub

Private Sub Class_Initialize()
clsEffect_LoadSettings
End Sub

Private Sub Class_Terminate()
Unload frm
clsEffect_SaveSettings
End Sub

Private Sub Frm_Change()
frmPreview.Update True
End Sub




Private Function BuildGraphString(ByRef Gph As dbGraph) As String
Dim tmp As String, i As Integer
With Gph
    tmp = Space$((UBound(.Points) + 1) * 6 - 1)
    For i = 0 To UBound(.Points)
        Mid(tmp, 1 + i * 6, 2) = VedNullStr(Hex$(.Points(i).x), 2)
        Mid(tmp, 1 + 3 + i * 6, 2) = VedNullStr(Hex$(.Points(i).y), 2)
    Next i
End With
BuildGraphString = tmp + CStr(Gph.InterpolationMode)
End Function

Private Sub ExtractGraphString(ByRef St As String, ByRef Gph As dbGraph)
Dim Splt() As String, i As Integer, n As Integer

Gph.InterpolationMode = CLng(Right$(St, 1))

If (Gph.InterpolationMode < 0) Or (Gph.InterpolationMode > 1) Then
    Gph.InterpolationMode = dbIMLinear
End If
Gph.NeedsInterpolation = True

St = Left$(St, Len(St) - 1)

Splt = Split(St, " ")
n = (UBound(Splt) + 1) \ 2 - 1
If n < 1 Then
    Err.Raise 5, "ExtractGraphString", "Too Short String"
End If
With Gph
    ReDim .Points(0 To n)
    For i = 0 To n
        .Points(i).x = CByte("&H" + Splt(i * 2))
        .Points(i).y = CByte("&H" + Splt(i * 2 + 1))
    Next i
    If .Points(0).x <> 0 Then .Points(0).x = 0
    If .Points(n).x <> 255 Then .Points(n).x = 255
End With
End Sub

Friend Sub InitGraph(ByRef Grph() As PointByte)
ReDim Grph(0 To 1)
Grph(0).x = 0
Grph(0).y = 0
Grph(1).x = 255
Grph(1).y = 255
End Sub

