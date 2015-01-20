Attribute VB_Name = "mdlEffects"
Option Explicit
Dim Effects() As clsEffect
Dim nEffects As Long

Public Sub ConnectEffect(ByRef Effect As clsEffect)
If nEffects = 0 Then
    ReDim Effects(0 To 0)
End If
ReDim Preserve Effects(0 To nEffects)
Set Effects(nEffects) = Effect
nEffects = nEffects + 1
End Sub

Public Function GetEffectByID(ByRef stID As String) As clsEffect
Dim EfID As String, EfMnu As String
Dim StIDU As String
Dim i As Long
StIDU = UCase$(stID)
For i = 0 To nEffects - 1
    Effects(i).GetEffectDesc EfID, EfMnu
    If UCase$(EfID) = UCase$(StIDU) Then
        Set GetEffectByID = Effects(i)
        Exit For
    End If
Next i
End Function

Public Sub PerformEffectEx(ByRef Effect As clsEffect, _
                           ByRef InData() As Long, _
                           ByRef OutData() As Long, _
                           Optional ByVal ShowDialog As Boolean = True, _
                           Optional ByVal xFrom As Long = &H80000000, _
                           Optional ByVal yFrom As Long, _
                           Optional ByVal xTo As Long, _
                           Optional ByVal yTo As Long)
Dim Range As RECT
On Error GoTo eh
If xFrom = &H80000000 Then
    TestDims InData
    xFrom = 0
    yFrom = 0
    xTo = UBound(InData, 1)
    yTo = UBound(InData, 2)
End If
With Range
    .Left = xFrom
    .Top = yFrom
    .Right = xTo + 1
    .Bottom = yTo + 1
End With
If ShowDialog Then
    CustomizeEffect Effect, InData
End If
MainForm.DisableMe
Effect.PerformEffect InData, OutData, Range.Left, Range.Top, Range.Right, Range.Bottom
MainForm.RestoreMeEnabled
Exit Sub
eh:
MainForm.ClearMeEnabledStack
ErrRaise Err
End Sub

Public Sub CustomizeEffect(ByRef Effect As clsEffect, _
                           ByRef InData() As Long)
Dim hWnd As Long, Pos As RECT
Dim frm As frmEffectPreview
Effect.LoadForm hWnd, Pos.Left, Pos.Top, Pos.Right, Pos.Bottom
Effect.SettingsToFrom
ShowPreviewWindow hWnd, Pos, True, Effect, InData, frm
Effect.SetPreviewWindow frm
On Error GoTo eh
Effect.ShowDialog
On Error GoTo 0
Effect.FormToSettings
Effect.SaveSettings
ShowPreviewWindow hWnd, Pos, False, Effect, InData, frm
Effect.UnloadForm
Exit Sub
eh:
PushError
ShowPreviewWindow hWnd, Pos, False, Effect, InData, frm
Effect.UnloadForm
PopError
ErrRaise Err
End Sub

Public Sub ShowPreviewWindow(ByVal hWndParent As Long, _
                             ByRef Pos As RECT, _
                             ByVal Show As Boolean, _
                             ByRef Effect As clsEffect, _
                             ByRef InData() As Long, _
                             ByRef frmPreview As frmEffectPreview)
Static frm As New frmEffectPreview
If hWndParent = 0 Then Exit Sub
If Show Then
    Load frm
    With frm
        vtSetParent frm.hWnd, hWndParent
        .SetEffect Effect
        .SetData AryPtr(InData)
        SetWindowPos .hWnd, HWND_TOP, _
                     Pos.Left, Pos.Top, _
                     Pos.Right - Pos.Left, Pos.Bottom - Pos.Top, SWP_NOOWNERZORDER
        ShowWindow .hWnd, SW_SHOWNA
    End With
    Set frmPreview = frm
Else
    With frm
        ShowWindow .hWnd, SW_HIDE
        .UnreferEffect
        .UnSetData
    End With
    Unload frm
    Set frmPreview = Nothing
End If
End Sub


Public Sub ConnectEffects()
Dim Gamma As New clsLinColorStretch
'Set Gamma = New clsLinColorStretch
ConnectEffect Gamma

Dim Graph As New clsColorMap
ConnectEffect Graph

Dim Filter As New clsFilter
ConnectEffect Filter

Dim ReRGB As New clsReRGB
ConnectEffect ReRGB

Dim Diff As New clsDiff
ConnectEffect Diff

End Sub

