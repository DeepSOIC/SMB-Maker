VERSION 5.00
Begin VB.Form frmZoom 
   BackColor       =   &H00E3DFE0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Увеличение"
   ClientHeight    =   3000
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   4170
   ControlBox      =   0   'False
   Icon            =   "frmZoom.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3000
   ScaleWidth      =   4170
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin SMBMaker.ctlResourcizator Resr1 
      Left            =   3555
      Top             =   0
      _ExtentX        =   1164
      _ExtentY        =   741
      ResID           =   9919
   End
   Begin SMBMaker.dbFrame Frame1 
      Height          =   2565
      Left            =   0
      TabIndex        =   7
      Top             =   15
      Width           =   4170
      _ExtentX        =   7355
      _ExtentY        =   4524
      Caption         =   "Выберите увеличение"
      BackColor       =   14933984
      EAC             =   0   'False
      Begin VB.TextBox Text1 
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
         Left            =   1590
         MaxLength       =   4
         TabIndex        =   8
         ToolTipText     =   "Только положительные целые числа!"
         Top             =   2130
         Width           =   900
      End
      Begin VB.OptionButton Opt 
         BackColor       =   &H0080FFFF&
         Caption         =   "16x"
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
         Index           =   5
         Left            =   150
         Style           =   1  'Graphical
         TabIndex        =   4
         Tag             =   "16"
         Top             =   1500
         Width           =   3840
      End
      Begin VB.OptionButton Opt 
         BackColor       =   &H0080FFFF&
         Caption         =   "8x"
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
         Index           =   4
         Left            =   150
         Style           =   1  'Graphical
         TabIndex        =   3
         Tag             =   "8"
         Top             =   1185
         Width           =   3840
      End
      Begin VB.OptionButton Opt 
         BackColor       =   &H0080FFFF&
         Caption         =   "4x"
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
         Index           =   3
         Left            =   150
         Style           =   1  'Graphical
         TabIndex        =   2
         Tag             =   "4"
         Top             =   870
         Width           =   3840
      End
      Begin VB.OptionButton Opt 
         BackColor       =   &H0080FFFF&
         Caption         =   "2x"
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
         Index           =   2
         Left            =   150
         Style           =   1  'Graphical
         TabIndex        =   1
         Tag             =   "2"
         Top             =   555
         Width           =   3840
      End
      Begin VB.OptionButton Opt 
         BackColor       =   &H0080FFFF&
         Caption         =   "32x"
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
         Index           =   6
         Left            =   150
         Style           =   1  'Graphical
         TabIndex        =   5
         Tag             =   "32"
         Top             =   1815
         Width           =   3840
      End
      Begin VB.OptionButton Opt 
         BackColor       =   &H0080FFFF&
         Caption         =   "Реальный размер (1:1)"
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
         Index           =   1
         Left            =   150
         Style           =   1  'Graphical
         TabIndex        =   0
         Tag             =   "1"
         Top             =   240
         Value           =   -1  'True
         Width           =   3840
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "другое:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   135
         TabIndex        =   10
         Top             =   2190
         Width           =   1425
      End
      Begin SMBMaker.dbButton OKButton 
         Height          =   315
         Left            =   2490
         TabIndex        =   9
         Top             =   2130
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   556
         MouseIcon       =   "frmZoom.frx":18BA
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Others          =   $"frmZoom.frx":18D6
         OthersPresent   =   -1  'True
      End
   End
   Begin SMBMaker.dbButton CancelButton 
      Cancel          =   -1  'True
      Default         =   -1  'True
      Height          =   390
      Left            =   0
      TabIndex        =   6
      Top             =   2610
      Width           =   4170
      _ExtentX        =   7355
      _ExtentY        =   688
      MouseIcon       =   "frmZoom.frx":1922
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Others          =   $"frmZoom.frx":193E
      OthersPresent   =   -1  'True
   End
End
Attribute VB_Name = "frmZoom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Frozen As Boolean

Private Sub CancelButton_Click()
Me.Tag = "C"
Me.Hide
End Sub

Private Sub Form_Activate()
Text1.SetFocus
Text1.SelStart = 0
Text1.SelLength = Len(Text1.Text)
End Sub

Private Sub Form_Load()
Resr1.LoadCaptions
End Sub

Private Sub Form_Paint()
On Error Resume Next
Me.PaintPicture gBackPicture, 0, 0, Me.ScaleWidth, Me.ScaleHeight
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
'Select Case Chr$(KeyAscii)
'Case "1", "2", "3", "4", "5", "6"
'    Opt(Val(Chr$(KeyAscii))).Value = True
'End Select
End Sub

Private Sub OkButton_Click()
Me.Tag = ""
Me.Hide
End Sub

Friend Sub SetZoom(ByVal r As Long)
Text1.Text = CStr(r)
End Sub

Friend Function GetZoom() As Long
Dim bCorrect As Boolean
Dim n As Long
n = Int(Val(Text1.Text))
bCorrect = (n > 0 And n < 32)
If bCorrect Then
    GetZoom = n
End If
End Function


Private Sub Opt_Click(Index As Integer)
If Frozen Then Exit Sub
On Error Resume Next
Frozen = True
Text1.Text = Opt(Index).Tag
ValidateOptions
Frozen = False
OkButton_Click
End Sub

Private Sub Text1_Change()
If Frozen Then Exit Sub
ValidateOptions
End Sub

Sub ValidateOptions()
Dim bCorrect As Boolean
Dim i As Long
Dim n As Long
On Error GoTo eh
n = Int(dbVal(Text1.Text))
bCorrect = (n > 0 And n < 32)
OkButton.Enabled = bCorrect
If bCorrect And n <= 6 Then
    If Not Frozen Then
        Frozen = True
        'On Error Resume Next
        For i = Opt.lBound To Opt.UBound
            Opt(i).Value = n = CLng(Opt(i).Tag)
        Next i
        Frozen = False
    End If
Else
    If Not Frozen Then
        Frozen = True
        On Error Resume Next
        For i = Opt.lBound To Opt.UBound
            Opt(i).Value = False
        Next i
        Frozen = False
    End If
End If
Exit Sub
eh:
bCorrect = False
Resume Next
End Sub
