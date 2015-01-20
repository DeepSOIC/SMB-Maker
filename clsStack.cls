VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsStack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Const IncDelta = 10
Public DefValue As Variant 'value returned if empty stack
Dim Storage() As Variant
Dim NumStored As Long 'number of items in the array
Dim StackPointer As Long 'points to the first empty position

Public Function Pop() As Variant
If StackPointer = 0 Then
    Pop = DefValue
Else
    StackPointer = StackPointer - 1&
    Pop = Storage(StackPointer)
    Storage(StackPointer) = Empty
End If
End Function

Public Sub Push(ByRef Value As Variant)
If NumStored = 0 Then
    NumStored = IncDelta
    ReDim Storage(0 To NumStored - 1)
    StackPointer = 0
End If
If StackPointer >= NumStored Then
    NumStored = NumStored + 10
    ReDim Preserve Storage(0 To NumStored - 1)
End If
Storage(StackPointer) = Value
StackPointer = StackPointer + 1&
End Sub

Public Sub Clear()
StackPointer = 0
NumStored = 0
Erase Storage
End Sub
