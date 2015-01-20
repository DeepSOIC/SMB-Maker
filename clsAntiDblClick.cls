VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAntiDblClick"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public CancelDblClick As Boolean 'set to True to disable doubleclicks

Dim BS(1 To 4) As Boolean 'stores if mousebutton currently pressed
Public LastMDB As Integer 'last seen MouseDown button
Public mmx As Single 'last seen position in mousemove
Public mmy As Single
Public mdx As Single 'position of last mousedown
Public mdy As Single
Public mmShift As Integer

Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event DblClick(Button As Integer, Shift As Integer, X As Single, Y As Single)


Public Sub MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'  If .ButtonState(1) Or .ButtonState(2) Or .ButtonState(4) Then
'  End If
TestBI Button
BS(Button) = True
LastMDB = Button
mdx = X
mdy = Y
mmx = X
mmy = Y
mmShift = Shift
RaiseEvent MouseDown(Button, Shift, X, Y)

End Sub

Public Sub MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
mmx = X
mmy = Y
mmShift = Shift
RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Public Sub MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
TestBI Button
BS(Button) = False
RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Public Sub DblClick()
Dim Button As Integer, Shift As Integer, X As Single, Y As Single
'With Picture1_MS
  'BS(LastMDB) = True
  TestBI LastMDB
  Button = LastMDB
  Shift = mmShift
  X = mmx
  Y = mmy
  If CancelDblClick Then
    MouseDown Button, mmShift, mmx, mmy
  Else
    BS(Button) = True
    'LastMDB = Button
    mdx = X
    mdy = Y
    'mmx = x
    'mmy = y
    'mmShift = Shift
    RaiseEvent DblClick(Button, Shift, X, Y)
  End If

End Sub

Private Sub TestBI(Button As Integer)
Select Case Button
  Case 1, 2, 4
  Case Else
    Err.Raise 5674, "clsAntiDblClick", "Bad button index value. It must be 1 or 2 or 4."
End Select
End Sub

Public Property Get ButtonState(ByVal Index As Integer) As Boolean
TestBI Index
ButtonState = BS(Index)
End Property

Public Property Let ButtonState(ByVal Index As Integer, ByVal NewVal As Boolean)
TestBI Index
BS(Index) = NewVal
End Property

Private Sub Class_Initialize()
CancelDblClick = True
End Sub
