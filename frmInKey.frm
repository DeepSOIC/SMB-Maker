VERSION 5.00
Begin VB.Form frmInKey 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2430
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   2160
   ControlBox      =   0   'False
   Icon            =   "frmInKey.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2430
   ScaleWidth      =   2160
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkNoShift 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Ignore Alt/Ctrl/Shift"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1740
      Width           =   1350
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   885
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "frmInKey.frx":000C
      Top             =   0
      Width           =   2040
   End
   Begin SMBMaker.dbButton OkButton 
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   960
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      MouseIcon       =   "frmInKey.frx":001B
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Others          =   $"frmInKey.frx":0037
      OthersPresent   =   -1  'True
   End
   Begin SMBMaker.dbButton CancelButton 
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   1320
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      MouseIcon       =   "frmInKey.frx":0089
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Others          =   $"frmInKey.frx":00A5
      OthersPresent   =   -1  'True
   End
End
Attribute VB_Name = "frmInKey"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CancelButton_Click()
Me.Tag = ""
Me.Hide
End Sub

Private Sub chkNoShift_Click()
If Val(Me.Tag) > 0 Then
    Text1_KeyDown CLng(Val(Me.Tag)) And &H7FF, 0
End If
Text1.SetFocus
End Sub

Private Sub Form_Activate()
Text1.SetFocus
End Sub

Private Sub Form_Load()
dbLoadCaptions
OkButton.Enabled = False
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
If Me.Tag = CStr(0) Then
    Text1.Text = GRSF(2245)
End If
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
Dim ShiftKeyCode As Long
'Me.Tag = CStr(KeyCode)
'Me.Hide
ShiftKeyCode = chkNoShift.Value * &H800 Or Shift * &H100 Or KeyCode

If KeyCode = 16 Or KeyCode = 17 Or KeyCode = 18 Then
    KeyCode = 0
    ShiftKeyCode = chkNoShift.Value * &H800 + Shift * &H100 + KeyCode
    Me.Tag = CStr(0)
    Text1.Text = GRSF(2338) + vbCrLf + GetKeyName(ShiftKeyCode)
    OkButton.Enabled = False
Else
    Me.Tag = CStr(ShiftKeyCode)
    Text1.Text = GRSF(2338) + vbCrLf + GetKeyName(ShiftKeyCode)
    OkButton.Enabled = True
End If
End Sub

Private Sub Form_Resize()
Dim XBorder As Long, YBorder As Long
XBorder = Me.Width - Me.ScaleWidth
YBorder = Me.Height - Me.ScaleHeight
Me.Move Me.Left, Me.Top, CancelButton.Width + OkButton.Width + XBorder, chkNoShift.Height + CancelButton.Height + Text1.Height + YBorder
Text1.Move 0, 0, Me.ScaleWidth
OkButton.Move 0, Text1.Height
CancelButton.Move OkButton.Width, Text1.Height
chkNoShift.Move 0, OkButton.Top + OkButton.Height, Me.ScaleWidth
End Sub

Public Function dbLoadCaptions()
Text1.Text = GRSF(2245)
chkNoShift.Caption = GRSF(2349)
End Function

Private Sub OkButton_Click()
Me.Hide
End Sub
