VERSION 5.00
Begin VB.Form frmMono 
   BackColor       =   &H00E3DFE0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Monochromize"
   ClientHeight    =   1455
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   1815
   Icon            =   "frmMono.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1455
   ScaleWidth      =   1815
   StartUpPosition =   1  'CenterOwner
   Begin SMBMaker.dbButton dbButton1 
      Default         =   -1  'True
      Height          =   375
      Index           =   1
      Left            =   0
      TabIndex        =   1
      Top             =   375
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      MouseIcon       =   "frmMono.frx":8C02
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Others          =   $"frmMono.frx":8C1E
      OthersPresent   =   -1  'True
   End
   Begin SMBMaker.dbButton CancelButton 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   750
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      MouseIcon       =   "frmMono.frx":8C6F
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Others          =   $"frmMono.frx":8C8B
      OthersPresent   =   -1  'True
   End
   Begin SMBMaker.dbButton dbButton1 
      Height          =   375
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      MouseIcon       =   "frmMono.frx":8CDB
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Others          =   $"frmMono.frx":8CF7
      OthersPresent   =   -1  'True
   End
End
Attribute VB_Name = "frmMono"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CancelButton_Click()
Me.Tag = "" 'non-standard behaivour
Me.Hide
End Sub

Private Sub dbButton1_Click(Index As Integer)
Me.Tag = CStr(Index) 'non-standard behaivour
Me.Hide
End Sub

Private Sub Form_Initialize()
dbLoadCaptions
End Sub

Private Sub Form_Load()
Dim tHgt As Long, tWdt As Long
tHgt = Me.Height - Me.ScaleHeight
tWdt = Me.Width - Me.ScaleWidth
Me.Width = CancelButton.Width + tWdt
Me.Height = CancelButton.Height * 3 + tHgt
End Sub

Private Sub dbLoadCaptions()
dbButton1(0).Caption = GRSF(2241)
dbButton1(1).Caption = GRSF(2242)
CancelButton.Caption = GRSF(2243)
Me.Caption = GRSF(2244)
End Sub
