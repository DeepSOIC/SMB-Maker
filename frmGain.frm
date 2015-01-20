VERSION 5.00
Begin VB.Form frmGain 
   BackColor       =   &H00E3DFE0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Scale colors"
   ClientHeight    =   3855
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   3480
   Icon            =   "frmGain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   257
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   232
   StartUpPosition =   1  'CenterOwner
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin SMBMaker.ctlResourcizator Resr1 
      Left            =   1410
      Top             =   330
      _ExtentX        =   1164
      _ExtentY        =   741
      ResID           =   9906
   End
   Begin SMBMaker.ctlNumBox nmbMul 
      Height          =   540
      Left            =   75
      TabIndex        =   4
      Top             =   270
      Width           =   1305
      _ExtentX        =   1667
      _ExtentY        =   953
      Min             =   -30
      Max             =   30
      NumType         =   5
      HorzMode        =   0   'False
      EditName        =   "Gain, in dB. To input the multiplier, type '=dB(multiplier)'."
      NLn             =   0
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Preview:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   150
      TabIndex        =   6
      Top             =   1290
      Width           =   630
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Gain:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   90
      TabIndex        =   5
      Top             =   30
      Width           =   1275
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00E3DFE0&
      BackStyle       =   0  'Transparent
      Caption         =   "The color multiplier is "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   105
      TabIndex        =   3
      Top             =   990
      Width           =   1530
   End
   Begin SMBMaker.dbButton Autodetect 
      Height          =   345
      Left            =   1395
      TabIndex        =   2
      ToolTipText     =   "Detects the gain that reaches color saturation."
      Top             =   1650
      Width           =   1395
      _ExtentX        =   0
      _ExtentY        =   0
      MouseIcon       =   "frmGain.frx":1CFA
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Others          =   $"frmGain.frx":1D16
      OthersPresent   =   -1  'True
   End
   Begin VB.Image iPreview 
      Height          =   2160
      Left            =   150
      Top             =   1545
      Width           =   3180
   End
   Begin SMBMaker.dbButton CancelButton 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   2085
      TabIndex        =   1
      Top             =   495
      Width           =   1215
      _ExtentX        =   0
      _ExtentY        =   0
      MouseIcon       =   "frmGain.frx":1D67
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Others          =   $"frmGain.frx":1D83
      OthersPresent   =   -1  'True
   End
   Begin SMBMaker.dbButton OKButton 
      Default         =   -1  'True
      Height          =   375
      Left            =   2085
      TabIndex        =   0
      Top             =   60
      Width           =   1215
      _ExtentX        =   0
      _ExtentY        =   0
      MouseIcon       =   "frmGain.frx":1DD3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Others          =   $"frmGain.frx":1DEF
      OthersPresent   =   -1  'True
   End
End
Attribute VB_Name = "frmGain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Public Event Change()

Private Sub Form_Paint()
Me.PaintPicture LoadResPicture(RES_BackPicture, vbResBitmap), 0, 0, Me.ScaleWidth, Me.ScaleHeight
End Sub

Private Sub Autodetect_Click()
Err.Raise 111, "Autodetect_Click", "Implement it!!! (internal error)"
End Sub

Private Sub CancelButton_Click()
Me.Tag = "c"
Me.Hide
End Sub

Private Sub Form_Load()
dbLoadCaptions
'nmbMul.Left = Label1.Left + Label1.Width + 60 / 15
'Label2.Left = nmbMul.Left + nmbMul.Width + 60 / 15
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If UnloadMode = 0 Then
    Cancel = True
    CancelButton_Click
End If
End Sub


Private Sub nmbMul_Change()
RaiseEvent Change
nmbMul_InputChange
End Sub

Private Sub nmbMul_InputChange()
Label1.Caption = grs(284, "%k", CStr(Round(GetFactor, 3)))
End Sub

Private Sub OkButton_Click()
'If Val(Text.Text) < 0 Or Val(Text.Text) > 256 Then dbMsgBox GRSF(1132), vbOKOnly Or vbInformation: Exit Sub
Me.Tag = ""
Me.Hide
End Sub



Sub dbLoadCaptions()
Resr1.LoadCaptions
'Me.Caption = GRSF(2133)
'Autodetect.Caption = GRSF(280)
'Autodetect.ToolTipText = GRSF(281)
'Label2.Caption = GRSF(283)
'Label1.Caption = GRSF(284)
'CancelButton.Caption = GRSF(285)
'OKButton.Caption = GRSF(286)
'Me.Icon = LoadResPicture(Me.Name, vbResIcon)
End Sub

Public Sub SetFactor(ByVal NewFactor As Double)
NewFactor = Round(NewFactor, 3)
On Error Resume Next
nmbMul.Value = FactorToDB(NewFactor)
End Sub

Public Function GetFactor() As Double
GetFactor = dBtoFactor(nmbMul.Value)
End Function
