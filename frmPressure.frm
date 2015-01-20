VERSION 5.00
Begin VB.Form frmPressure 
   Caption         =   "Tab"
   ClientHeight    =   2925
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7680
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   204
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   195
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   512
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox Check1 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Disable click-and-hold (Vista only)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1920
      Width           =   2820
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   5115
      Top             =   2295
   End
   Begin VB.PictureBox Picture1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   780
      Left            =   0
      ScaleHeight     =   48
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   512
      TabIndex        =   0
      Top             =   285
      Width           =   7740
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmPressure.frx":0000
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   105
      TabIndex        =   4
      Top             =   1095
      Width           =   7485
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pressure gauge"
      Height          =   195
      Left            =   210
      TabIndex        =   2
      Top             =   45
      Width           =   1125
   End
   Begin SMBMaker.dbButton OkButton 
      Default         =   -1  'True
      Height          =   360
      Left            =   2775
      TabIndex        =   1
      Top             =   2295
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   635
      MouseIcon       =   "frmPressure.frx":017A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Others          =   $"frmPressure.frx":0196
      OthersPresent   =   -1  'True
   End
End
Attribute VB_Name = "frmPressure"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Pressure As Long
Public UnlNow As Boolean
Public MaxPr As Long
Public OldMaxPr As Long

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case 112 'f1
        dbMsgBox 2352, vbInformation
End Select
End Sub

Private Sub Form_Load()
dbLoadCaptions
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If UnloadMode = 0 Then
    Cancel = 1
    Me.Tag = "C"
    UnlNow = True
    'Me.Hide
End If
End Sub

Private Sub Form_Resize()
Dim bw As Long, bh As Long
On Error Resume Next
bw = Me.Width \ Screen.TwipsPerPixelX - Me.ScaleWidth
bh = Me.Height \ Screen.TwipsPerPixelY - Me.ScaleHeight
Me.Move Me.Left, Me.Top, _
        Me.Width, _
        (Picture1.Height + OkButton.Height + bh) * Screen.TwipsPerPixelY
Picture1.Move 0, 0, Me.ScaleWidth
OkButton.Move (Me.ScaleWidth - OkButton.Width) \ 2, Picture1.Height
End Sub

Private Sub Form_Paint()
On Error Resume Next
Me.PaintPicture gBackPicture, 0, 0, Me.ScaleWidth, Me.ScaleHeight
End Sub

Private Sub HoldMessages()
Do
    If Not dbProcessMessages Then DoEvents
    If UnlNow Then
        Exit Do
    End If
Loop
Me.Hide
End Sub

Public Function dbProcessMessages(Optional ByVal WaitMes As Boolean = False) As Boolean
Dim Msgs As Msg
Dim ExInfo As Long
Static lPr As Long
WaitMessage
dbProcessMessages = False
If PeekMessage(Msgs, Me.hWnd, WM_MOUSEFIRST, WM_MOUSELAST, 0) Then
    ExInfo = GetMessageExtraInfo
    Pressure = ExInfo And &H1FF
    If Pressure <> lPr Then
        lPr = Pressure
        PrChange Pressure
    End If
    dbProcessMessages = False
End If
End Function

Public Sub PrChange(ByVal NewPressure As Long)
    If Pressure > MaxPr Then MaxPr = Pressure
    If MaxPr > Picture1.ScaleWidth - 20 Then
        Me.Move Me.Left, Me.Top, Me.Width + (MaxPr - Picture1.ScaleWidth + 20) * Screen.TwipsPerPixelX
        Exit Sub
    End If
    Picture1.Line (0, 0)-(Pressure, Picture1.ScaleHeight), vbYellow, BF
    Picture1.Line (Pressure, 0)-(Picture1.ScaleWidth, Picture1.ScaleHeight), vbBlack, BF
    Picture1.Line (MaxPr + 1, 0)-(MaxPr + 1, Picture1.ScaleHeight), vbRed
    If OldMaxPr > 0 Then
        Picture1.Line (OldMaxPr + 1, 0)-(OldMaxPr + 1, Picture1.ScaleHeight), vbBlue
    End If
End Sub

Private Sub OkButton_Click()
Me.Tag = ""
MainForm.MaxPenPressure = MaxPr
UnlNow = True
End Sub

Private Sub Picture1_Paint()
PrChange Pressure
End Sub

Private Sub Timer1_Timer()
Timer1.Enabled = False
HoldMessages
End Sub

Public Sub dbLoadCaptions()
Me.Caption = GRSF(2351)
End Sub
