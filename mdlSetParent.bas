Attribute VB_Name = "mdlSetParent"
Option Explicit

Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Const GWL_STYLE As Long = (-16)
Private Const WS_CHILD As Long = &H40000000
Private Const WS_POPUP As Long = &H80000000

Private Sub SetWindowStyle(ByVal hWnd As Long, ByVal NewWS As Long)
SetWindowLong hWnd, GWL_STYLE, NewWS
End Sub

Private Function GetWindowStyle(ByVal hWnd As Long) As Long
GetWindowStyle = GetWindowLong(hWnd, GWL_STYLE)
End Function

'Flags specified by RemoveFlags will be removed.
'Setting has larger priority than removal
Private Sub ChangeWindowStyle(ByVal hWnd As Long, _
                              ByVal SetFlags As Long, _
                              ByVal RemoveFlags As Long)
SetWindowStyle hWnd, GetWindowStyle(hWnd) And Not RemoveFlags Or SetFlags
End Sub

Public Function vtSetParent(ByVal hWndChild As Long, ByVal hWndParent As Long) As Long
If hWndParent <> 0 Then
    ChangeWindowStyle hWndChild, SetFlags:=WS_CHILD, RemoveFlags:=WS_POPUP
End If
vtSetParent = SetParent(hWndChild, hWndParent)
If hWndParent = 0 Then
    ChangeWindowStyle hWndChild, SetFlags:=WS_POPUP, RemoveFlags:=WS_CHILD
End If
End Function

