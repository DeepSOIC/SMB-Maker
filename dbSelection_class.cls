VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "dbSelection_class"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Dim SelData() As gg
Public XM As Long, YM As Long
Dim pSelected As Boolean
Public Moving As Boolean
Public TransColor As Long
Public TransRatio As Single
Public StretchMode As eStretchMode
'Dim TransOrigData() As Long, TransData() As Long
Dim pSelMode As dbSelMode
Public x1 As Long
Public y1 As Long
Public x2 As Long
Public y2 As Long
Public mnuSelection As Menu
Public pIsText As Boolean
Dim TransDataSaved() As Long


Public Property Get SelMode() As dbSelMode
If IsText Then
  SelMode = dbSuperTransparent
Else
  SelMode = pSelMode
End If
End Property

Public Property Let SelMode(ByVal nSelMode As dbSelMode)
If Not IsText Then
  If nSelMode = dbUseCurSelMode Then
      nSelMode = MainForm.CurSelMode
  End If
  If pSelMode = dbSuperTransparent And nSelMode <> dbSuperTransparent Then
      Erase TransData
      pSelMode = nSelMode
  ElseIf pSelMode <> dbSuperTransparent And nSelMode = dbSuperTransparent Then
      pSelMode = nSelMode
  Else
      pSelMode = nSelMode
  End If
End If
End Property

Public Function TransDataPresent() As Boolean
TransDataPresent = (AryDims(AryPtr(TransData)) = 2) ',2 and exactly 2
End Function

Public Function TransOrigDataPresent() As Boolean
TransOrigDataPresent = AryDims(AryPtr(TransOrigData)) = 2
End Function

'Public Property Get x1() As Long
'x1 = pX1
'End Property
'
'Public Property Let x1(ByVal nC As Long)
''ResizeSel nC, pY1, pX2, pY2
'pX1 = nC
'End Property
'
'Public Property Get y1() As Long
'y1 = pY1
'End Property
'
'Public Property Let y1(ByVal nC As Long)
''ResizeSel pX1, nC, pX2, pY2
'pY1 = nC
'End Property
'
'Public Property Get X2() As Long
'X2 = pX2
'End Property
'
'Public Property Let X2(ByVal nC As Long)
''ResizeSel pX1, pY1, nC, pY2
'pX2 = nC
'End Property
'
'Public Property Get Y2() As Long
'Y2 = pY2
'End Property
'
'Public Property Let Y2(ByVal nC As Long)
'ResizeSel pX1, pY1, pX2, nC
'End Property

Public Sub ResizeSel(ByVal lngX1 As Long, _
                     ByVal lngY1 As Long, _
                     ByVal lngX2 As Long, _
                     ByVal lngY2 As Long)
Dim w As Long, h As Long
Dim ppX1 As Long, ppY1 As Long, ppX2 As Long, ppY2 As Long
Dim DDE As Boolean
x1 = Min(lngX1, lngX2)
x2 = Max(lngX1, lngX2)
y1 = Min(lngY1, lngY2)
y2 = Max(lngY1, lngY2)
'w = Abs(lngX2 - lngX1) + 1
'h = Abs(lngY2 - lngY1) + 1
'If pSelMode = dbSuperTransparent And Not (Moving) And Selected Then
'    If TransOrigDataPresent Then
'        TransData = TransOrigData
'        On Error GoTo eh
'        DDE = DontDoEvents
'        DontDoEvents = True
'        MainForm.dbStretch TransData, w, h, False
'        DontDoEvents = DDE
'        On Error GoTo 0
'    Else
'        ReDim TransData(0 To w - 1, 0 To h - 1)
'    End If
'End If
'pX1 = ppX1
'pY1 = ppY1
'pX2 = ppX2
'pY2 = ppY2
Exit Sub
eh:
DontDoEvents = DDE
Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Function InSel(ByVal X As Long, ByVal Y As Long) As Boolean
InSel = (X >= pX1 And Y >= pY1 And X <= pX2 And Y <= pY2)
End Function


Public Property Get Selected() As Boolean
Selected = pSelected
End Property

Public Property Let Selected(ByVal vNewValue As Boolean)
Dim Changed As Boolean
Changed = vNewValue Xor pSelected
pSelected = vNewValue
If Changed Then
    If Not mnuSelection Is Nothing Then
        mnuSelection.Visible = pSelected
    End If
    If pIsText And Not vNewValue Then
      Erase TransOrigData
      Erase TransData
      SwapArys AryPtr(TransOrigData), AryPtr(TransDataSaved)
      pIsText = False
    End If
End If
End Property

Public Property Get IsText() As Boolean
IsText = pIsText
End Property

'IsText can be set. But cannot be reset - only by deselecting
Public Sub SetIsText()
If pIsText Then Exit Sub
Erase TransDataSaved
SwapArys AryPtr(TransOrigData), AryPtr(TransDataSaved)
Erase TransData
pIsText = True
End Sub
