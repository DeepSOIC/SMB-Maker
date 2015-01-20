VERSION 5.00
Begin VB.Form frmCircle 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Circle preferences"
   ClientHeight    =   2475
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7665
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2475
   ScaleWidth      =   7665
   StartUpPosition =   1  'CenterOwner
   Begin SMBMaker.ctlResourcizator Resr1 
      Left            =   5580
      Top             =   1485
      _ExtentX        =   1164
      _ExtentY        =   741
      ResID           =   9915
   End
   Begin VB.CheckBox Chk 
      BackColor       =   &H00FFC0FF&
      Caption         =   "High Quality"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Index           =   3
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Enable antialiasing."
      Top             =   1620
      Width           =   1995
   End
   Begin VB.CheckBox Chk 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Dot lines"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Index           =   2
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Use non-solid line"
      Top             =   1125
      Width           =   1995
   End
   Begin VB.CheckBox Chk 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Draw focuses of ellipse"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Index           =   1
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Draw focuses in the final ellipse."
      Top             =   765
      Width           =   1995
   End
   Begin VB.CheckBox Chk 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Draw center pixel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Index           =   0
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Draw center dot in final ellipse."
      Top             =   405
      Width           =   1995
   End
   Begin SMBMaker.dbFrame dbFrame1 
      Height          =   1860
      Left            =   180
      TabIndex        =   0
      Top             =   120
      Width           =   2565
      _ExtentX        =   4524
      _ExtentY        =   3281
      Caption         =   "Draw method"
      EAC             =   0   'False
      Begin VB.OptionButton ModeOpt 
         BackColor       =   &H0080FFFF&
         Caption         =   "Outside rectangle"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         Index           =   3
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Ellipse escribed around the rectangle given by start and end points."
         Top             =   1335
         Width           =   2265
      End
      Begin VB.OptionButton ModeOpt 
         BackColor       =   &H0080FFFF&
         Caption         =   "Inside rectangle"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         Index           =   2
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Ellipse inscribed into the rectangle given by start and end points."
         Top             =   975
         Width           =   2265
      End
      Begin VB.OptionButton ModeOpt 
         BackColor       =   &H0080FFFF&
         Caption         =   "On Radius"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         Index           =   1
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Start point is the center, end point defines the radius."
         Top             =   615
         Width           =   2265
      End
      Begin VB.OptionButton ModeOpt 
         BackColor       =   &H0080FFFF&
         Caption         =   "On diameter"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         Index           =   0
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "The smallest circle containign both start and end points."
         Top             =   255
         Width           =   2265
      End
   End
   Begin SMBMaker.dbButton dbButton1 
      Height          =   360
      Left            =   165
      TabIndex        =   8
      Top             =   2055
      Width           =   7410
      _ExtentX        =   13070
      _ExtentY        =   635
      MouseIcon       =   "frmCircle.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Others          =   $"frmCircle.frx":001C
      OthersPresent   =   -1  'True
   End
   Begin SMBMaker.dbButton OkButton 
      Default         =   -1  'True
      Height          =   375
      Left            =   5550
      TabIndex        =   6
      Top             =   255
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      MouseIcon       =   "frmCircle.frx":0072
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Others          =   $"frmCircle.frx":008E
      OthersPresent   =   -1  'True
   End
   Begin SMBMaker.dbButton CancelButton 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   5535
      TabIndex        =   7
      Top             =   705
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      MouseIcon       =   "frmCircle.frx":00DA
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Others          =   $"frmCircle.frx":00F6
      OthersPresent   =   -1  'True
   End
End
Attribute VB_Name = "frmCircle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim prvFDSC As FadeDesc

Private Sub CancelButton_Click()
Me.Tag = "C"
Me.Hide
End Sub

Private Sub dbButton1_Click()
Load Pereliv
With Pereliv
    MainForm.SendFadeDesc prvFDSC
    Pereliv.Show vbModal
    If .Tag = "" Then
        MainForm.ExtractFadeDesc prvFDSC
    End If
End With
Unload Pereliv
End Sub

Private Sub Form_Load()
Resr1.LoadCaptions
End Sub

Private Sub Form_Paint()
On Error Resume Next
Me.PaintPicture gBackPicture, 0, 0, Me.Width, Me.Height
End Sub

Friend Function GetProps(ByRef FadeDsc As FadeDesc) As Long
Dim i As Integer
Dim ModeIndex As Long
ModeIndex = 0
For i = 0 To 15
    If ModeOpt(i).Value Then
        ModeIndex = i
        Exit For
    End If
Next i
If Chk(0).Value = 1 Then
    ModeIndex = ModeIndex Or &H10&
End If
If Chk(1).Value = 1 Then
    ModeIndex = ModeIndex Or &H20&
End If
If Chk(2).Value = 1 Then
    ModeIndex = ModeIndex Or &H40&
End If
If Chk(3).Value = 1 Then
    ModeIndex = ModeIndex Or &H80&
End If
GetProps = (ModeIndex)
FadeDsc = prvFDSC
End Function

Friend Sub SetProps(ByVal Flags As Long, ByRef FadeDsc As FadeDesc)
Dim i As Long
i = Flags And &HF 'Mode mask
ModeOpt(i).Value = True
Chk(0).Value = Abs(CBool(Flags And &H10&))
Chk(1).Value = Abs(CBool(Flags And &H20&))
Chk(2).Value = Abs(CBool(Flags And &H40&))
Chk(3).Value = Abs(CBool(Flags And &H80&))
prvFDSC = FadeDsc
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If UnloadMode = 0 Then
    Cancel = True
    CancelButton_Click
End If
End Sub

Private Sub OkButton_Click()
Me.Tag = ""
Me.Hide
End Sub
