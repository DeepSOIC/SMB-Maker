VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAligner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public dx As Long, dy As Long
Public AnchorX As OrgAnchor
Public AnchorY As OrgAnchor
Public BaseAnchorX As OrgAnchor
Public BaseAnchorY As OrgAnchor
Public BasePointSupported As Boolean
Public DestPointSupported As Boolean

Public Enum OrgAnchor
    ATop = 0
    ACenter = 1
    ABottom = 2
    APoint = 3
End Enum

Friend Function GetOffset(ByVal BaseW As Long, _
                          ByVal BaseH As Long, _
                          ByVal w As Long, _
                          ByVal h As Long, _
                          Optional ByVal BasePointX As Long, _
                          Optional ByVal BasePointY As Long, _
                          Optional ByVal DestPointX As Long, _
                          Optional ByVal DestPointY As Long) As POINTAPI
GetOffset.X = GetOffsetSingle(BaseW, w, BasePointX, DestPointX, BaseAnchorX, AnchorX, dx)
GetOffset.Y = GetOffsetSingle(BaseH, h, BasePointY, DestPointY, BaseAnchorY, AnchorY, dy)
End Function

Private Function GetOffsetSingle(ByVal BaseH As Long, _
                                 ByVal h As Long, _
                                 ByVal BasePoint As Long, _
                                 ByVal DestPoint As Long, _
                                 ByVal BaseAnc As OrgAnchor, _
                                 ByVal DestAnc As OrgAnchor, _
                                 ByVal dy As Long)
GetAncPoint BaseH, BaseAnc, BasePoint
GetAncPoint h, DestAnc, DestPoint
GetOffsetSingle = BasePoint + dy - DestPoint
End Function

Private Sub GetAncPoint(ByVal h As Long, _
                        ByVal Anc As OrgAnchor, _
                        ByRef Point As Long)
Select Case Anc
    Case OrgAnchor.ATop
        Point = 0
    Case OrgAnchor.ACenter
        Point = h \ 2
    Case OrgAnchor.ABottom
        Point = h
    Case OrgAnchor.APoint
        'do nothing
End Select
End Sub

Public Sub SaveToReg(ByRef Section As String, _
                     ByRef Parameter As String)
Dim sArr() As String
ReDim sArr(0 To 5)
sArr(0) = dbCStr(BaseAnchorX)
sArr(1) = dbCStr(BaseAnchorY)
sArr(2) = dbCStr(AnchorX)
sArr(3) = dbCStr(AnchorY)
sArr(4) = dbCStr(dx)
sArr(5) = dbCStr(dy)
dbSaveSetting Section, Parameter, Join(sArr, ", ")
End Sub

Public Sub LoadFromReg(ByRef Section As String, _
                       ByRef Parameter As String)
Dim tmp As String
Dim sArr() As String
On Error GoTo eh
tmp = dbGetSetting(Section, Parameter, "0, 0, 0, 0, 0, 0")
sArr = Split(tmp, ",")
Select Case UBound(sArr)
    Case Is > 5
        Err.Raise 1111, , "Aligner: number of options too large!"
    Case 4, 2 To 0
        Err.Raise 1111, , "Aligner: incorrect number of options!"
End Select
BaseAnchorX = dbVal(sArr(0))
BaseAnchorY = dbVal(sArr(1))
AnchorX = dbVal(sArr(2))
AnchorY = dbVal(sArr(3))
If UBound(sArr) = 5 Then
    dx = dbVal(sArr(4))
    dy = dbVal(sArr(5))
End If
Exit Sub
eh:
ErrRaise "Aligner::LoadFromReg"
End Sub

Public Sub Customize(Optional ByVal RaiseErrors As Boolean = False)
Dim OK As Boolean
On Error GoTo eh
Load frmAligner
With frmAligner
    .SetSupport BasePointSupported, DestPointSupported
    .SetProps BaseAnchorX, BaseAnchorY, AnchorX, AnchorY, dx, dy
    .Show vbModal
    If Len(.Tag) = 0 Then
        OK = True
        .GetProps BaseAnchorX, BaseAnchorY, AnchorX, AnchorY, dx, dy
    End If
End With
Unload frmAligner
If Not OK Then Err.Raise dbCWS
Exit Sub
eh:
If RaiseErrors Then
    ErrRaise "Aligner::Customize"
Else
    If Err.Number = dbCWS Then Exit Sub
    MsgError
End If
End Sub

