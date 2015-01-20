VERSION 5.00
Begin VB.Form frmKeyb 
   BackColor       =   &H00E3DFE0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Mouse control steps"
   ClientHeight    =   1980
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5265
   Icon            =   "frmKeyb.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1980
   ScaleWidth      =   5265
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin SMBMaker.dbFrame Frame1 
      Height          =   1950
      Left            =   15
      TabIndex        =   10
      Top             =   15
      Width           =   3960
      _ExtentX        =   0
      _ExtentY        =   0
      Caption         =   "Steps"
      BackColor       =   14933984
      EAC             =   0   'False
      Begin VB.TextBox Stp 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   3
         Left            =   1920
         TabIndex        =   6
         Text            =   "1"
         Top             =   1455
         Width           =   525
      End
      Begin VB.TextBox Stp 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   2
         Left            =   1905
         TabIndex        =   4
         Text            =   "1"
         Top             =   1080
         Width           =   525
      End
      Begin VB.TextBox Stp 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   1
         Left            =   1920
         TabIndex        =   2
         Text            =   "1"
         Top             =   690
         Width           =   525
      End
      Begin VB.TextBox Stp 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   0
         Left            =   1920
         TabIndex        =   0
         Text            =   "1"
         Top             =   300
         Width           =   525
      End
      Begin SMBMaker.dbButton Edi 
         Height          =   300
         Index           =   3
         Left            =   2460
         TabIndex        =   7
         Tag             =   "True"
         Top             =   1455
         Width           =   1350
         _ExtentX        =   0
         _ExtentY        =   0
         MouseIcon       =   "frmKeyb.frx":0442
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Others          =   $"frmKeyb.frx":045E
         OthersPresent   =   -1  'True
      End
      Begin SMBMaker.dbButton Edi 
         Height          =   300
         Index           =   2
         Left            =   2460
         TabIndex        =   5
         Tag             =   "True"
         Top             =   1080
         Width           =   1350
         _ExtentX        =   0
         _ExtentY        =   0
         MouseIcon       =   "frmKeyb.frx":04AE
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Others          =   $"frmKeyb.frx":04CA
         OthersPresent   =   -1  'True
      End
      Begin SMBMaker.dbButton Edi 
         Height          =   300
         Index           =   1
         Left            =   2460
         TabIndex        =   3
         Tag             =   "True"
         Top             =   690
         Width           =   1350
         _ExtentX        =   0
         _ExtentY        =   0
         MouseIcon       =   "frmKeyb.frx":051A
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Others          =   $"frmKeyb.frx":0536
         OthersPresent   =   -1  'True
      End
      Begin SMBMaker.dbButton Edi 
         Height          =   300
         Index           =   0
         Left            =   2460
         TabIndex        =   1
         Tag             =   "True"
         Top             =   300
         Width           =   1350
         _ExtentX        =   0
         _ExtentY        =   0
         MouseIcon       =   "frmKeyb.frx":0586
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Others          =   $"frmKeyb.frx":05A2
         OthersPresent   =   -1  'True
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00E3DFE0&
         BackStyle       =   0  'Transparent
         Caption         =   "With Alt"
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
         Left            =   135
         TabIndex        =   14
         Top             =   1470
         Width           =   570
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00E3DFE0&
         BackStyle       =   0  'Transparent
         Caption         =   "With Shift"
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
         Left            =   135
         TabIndex        =   13
         Top             =   735
         Width           =   705
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00E3DFE0&
         BackStyle       =   0  'Transparent
         Caption         =   "With Ctrl"
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
         Left            =   135
         TabIndex        =   12
         Top             =   1110
         Width           =   630
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00E3DFE0&
         BackStyle       =   0  'Transparent
         Caption         =   "Default"
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
         Left            =   135
         TabIndex        =   11
         Top             =   360
         Width           =   525
      End
   End
   Begin SMBMaker.dbButton CancelButton 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   4020
      TabIndex        =   9
      Top             =   525
      Width           =   1215
      _ExtentX        =   0
      _ExtentY        =   0
      MouseIcon       =   "frmKeyb.frx":05F2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Others          =   $"frmKeyb.frx":060E
      OthersPresent   =   -1  'True
   End
   Begin SMBMaker.dbButton OKButton 
      Default         =   -1  'True
      Height          =   375
      Left            =   4020
      TabIndex        =   8
      Top             =   45
      Width           =   1215
      _ExtentX        =   0
      _ExtentY        =   0
      MouseIcon       =   "frmKeyb.frx":065E
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Others          =   $"frmKeyb.frx":067A
      OthersPresent   =   -1  'True
   End
End
Attribute VB_Name = "frmKeyb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim KeysList() As Integer
Option Explicit

Private Sub Form_Paint()
On Error Resume Next
Me.PaintPicture gBackPicture, 0, 0, Me.ScaleWidth, Me.ScaleHeight
End Sub

Private Sub Combo1_Change()

End Sub

Private Sub CancelButton_Click()
Me.Tag = "C"
Me.Hide
End Sub

Sub Edi_Click(Index As Integer)
With Edi(Index)
    .Tag = CStr(Not (CBool(.Tag)))
    If CBool(.Tag) Then .Caption = GRSF(1213) Else .Caption = GRSF(1214)
End With
End Sub

Private Sub Form_Load()
dbLoadCaptions
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If UnloadMode = 0 Then
    Cancel = True
    Me.Tag = "C"
    Me.Hide
End If
End Sub

Private Sub OkButton_Click()
Dim i As Long, j As Long
For i = 0 To Stp.UBound
    With Stp(i)
        If Val(.Text) > 1024 Or Val(.Text) < 1 Then
            dbMsgBox GRSF(1141), vbInformation
            .SetFocus
            Exit Sub
        End If
    End With
Next i
Me.Tag = ""
Me.Hide
End Sub

Sub dbLoadCaptions()
Me.Caption = GRSF(2135)
Frame1.Caption = GRSF(2028)
Edi(3).Caption = GRSF(2029)
Edi(2).Caption = GRSF(2029)
Edi(1).Caption = GRSF(2029)
Edi(0).Caption = GRSF(2029)
Label5.Caption = GRSF(2030)
Label4.Caption = GRSF(2031)
Label3.Caption = GRSF(2032)
Label2.Caption = GRSF(2033)
CancelButton.Caption = GRSF(2034)
OkButton.Caption = GRSF(2035)
End Sub

