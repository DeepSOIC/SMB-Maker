VERSION 5.00
Begin VB.UserControl dbFrame 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1860
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2430
   ControlContainer=   -1  'True
   ScaleHeight     =   124
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   162
   ToolboxBitmap   =   "dbFrame.ctx":0000
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   285
      TabIndex        =   0
      Top             =   0
      Width           =   465
   End
End
Attribute VB_Name = "dbFrame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim pNo3D As Boolean
Dim lngBC As Long
Dim En As Boolean
Dim pResID As Long
Dim strCaption As String
Dim EAC As Boolean
Const p_RES_BackPicture = 108

Public Event Resize()
Public Event MouseUp(ByVal x As Long, ByVal y As Long, ByVal Button As Integer, ByVal Shift As Integer)
Public Event Paint()


Private Sub UserControl_InitProperties()
lngBC = vbButtonFace
Label1.Caption = UserControl.Name
En = True
UserControl.Enabled = True
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
RaiseEvent MouseUp(x, y, Button, Shift)
End Sub

Private Sub UserControl_Paint()
If pNo3D Then
    'UserControl.Cls
Else
    UserControl.PaintPicture LoadResPicture(p_RES_BackPicture, vbResBitmap), 0, 0, ScaleWidth, ScaleHeight
End If
RaiseEvent Paint
End Sub

Public Property Get Caption() As String
Caption = Label1.Caption
End Property

Public Property Let Caption(ByVal vNewValue As String)
strCaption = vNewValue
FreshCaption
End Property

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
With PropBag
    strCaption = .ReadProperty("Caption", "")
    pNo3D = .ReadProperty("No3D", False)
    lngBC = .ReadProperty("BackColor", vbButtonFace)
    Enabled = .ReadProperty("Enabled", True)
    pResID = .ReadProperty("ResID", 0)
    EAC = .ReadProperty("EAC", False)
    FreshCaption
End With
End Sub

Private Sub UserControl_Resize()
RaiseEvent Resize
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
With PropBag
    .WriteProperty "Caption", strCaption, ""
    .WriteProperty "No3D", pNo3D, False
    .WriteProperty "BackColor", lngBC, vbButtonFace
    .WriteProperty "Enabled", En, True
    .WriteProperty "ResID", pResID, 0
    .WriteProperty "EAC", EAC
End With
End Sub

Public Property Get No3D() As Boolean
No3D = pNo3D
End Property

Public Property Let No3D(ByVal vNewValue As Boolean)
pNo3D = vNewValue
UserControl.Refresh
End Property

Public Property Get BackColor() As OLE_COLOR
BackColor = lngBC
End Property

Public Property Let BackColor(ByVal vNewValue As OLE_COLOR)
lngBC = vNewValue
UserControl.BackColor = lngBC
Refresh
End Property

Public Property Get Enabled() As Boolean
Enabled = En
End Property

Public Property Let Enabled(ByVal bNew As Boolean)
En = bNew
UserControl.Enabled = En
EnableAllControls En
End Property

Public Sub EnableAllControls(ByVal bEn As Boolean)
Dim Ctl As Control
On Error Resume Next
For Each Ctl In UserControl.ContainedControls
    Ctl.Enabled = bEn
Next
End Sub

Public Property Get SyncEnableControls() As Boolean
SyncEnableControls = EAC
End Property

Public Property Let SyncEnableControls(ByVal bNew As Boolean)
EAC = bNew
End Property

Public Property Get ResID() As Long
ResID = pResID
End Property

Public Property Let ResID(ByVal vNewValue As Long)
pResID = vNewValue
FreshCaption
End Property

Private Sub FreshCaption()
Dim tmpCaption As String
If ResID = 0 Then
    tmpCaption = strCaption
Else
    On Error GoTo eh
    tmpCaption = LoadResString(pResID)
    On Error GoTo 0
End If
Label1.Caption = tmpCaption
Exit Sub
eh:
tmpCaption = strCaption
Resume Next
End Sub

'Public Function DrawText(ByRef St As String, _
'                         ByVal X As Long, ByVal Y As Long, _
'                         ByVal w As Long, _
'                         ByVal lngColor As Long, _
'                         FontSize As Long, _
'                         ByVal CenterText As Boolean) As Long
'Dim DTP As DRAWTEXTPARAMS
'Dim Rct As RECT
'Rct.Left = X
'Rct.Top = Y
'Rct.Right = X + w
'Rct.Bottom = Y + 2
'UserControl.ForeColor = lngColor
'UserControl.FontSize = FontSize
'DTP.cbSize = LenB(DTP)
'DTP.iLeftMargin = 1
'DTP.iRightMargin = 1
'DTP.iTabLength = 4
''DrawTextEx UserControl.hDC, St, Len(St), rct, DT_EXPANDTABS Or DT_CALCRECT Or DT_NOCLIP Or DT_NOPREFIX Or DT_TABSTOP Or DT_WORDBREAK, DTP
'DrawTextEx UserControl.hDC, St, Len(St), _
'           Rct, _
'           DT_CALCRECT Or _
'           DT_EXPANDTABS Or _
'           DT_NOPREFIX Or _
'           DT_TABSTOP Or _
'           DT_WORDBREAK Or _
'           (DT_CENTER And CenterText), _
'           DTP
'Rct.Left = X
'Rct.Top = Y
'Rct.Right = X + w
'DrawTextEx UserControl.hDC, St, Len(St), _
'           Rct, _
'           DT_EXPANDTABS Or _
'           DT_NOPREFIX Or _
'           DT_TABSTOP Or _
'           DT_WORDBREAK Or _
'           (DT_CENTER And CenterText), _
'           DTP
'DrawText = Rct.Bottom
'End Function

Public Function ScaleWidth()
ScaleWidth = UserControl.ScaleWidth
End Function

Public Function ScaleHeight()
ScaleHeight = UserControl.ScaleHeight
End Function

Public Sub Refresh()
UserControl.Refresh
End Sub
