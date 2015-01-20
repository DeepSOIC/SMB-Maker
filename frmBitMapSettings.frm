VERSION 5.00
Begin VB.Form frmFormatBMP 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bitmap Saving Options"
   ClientHeight    =   5310
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4785
   Icon            =   "frmBitMapSettings.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5310
   ScaleWidth      =   4785
   StartUpPosition =   1  'CenterOwner
   Begin SMBMaker.dbFrame dbFrame2 
      Height          =   1080
      Left            =   120
      TabIndex        =   1
      Top             =   2385
      Width           =   4560
      _ExtentX        =   8043
      _ExtentY        =   1905
      Caption         =   "Pixels Per Meter"
      EAC             =   0   'False
      Begin VB.TextBox DpiY 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1350
         TabIndex        =   3
         Text            =   "0"
         Top             =   570
         Width           =   1680
      End
      Begin VB.TextBox DpiX 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1350
         TabIndex        =   2
         Text            =   "0"
         Top             =   240
         Width           =   1680
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Y"
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
         Left            =   1050
         TabIndex        =   5
         Top             =   630
         Width           =   90
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "X"
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
         Left            =   1050
         TabIndex        =   4
         Top             =   285
         Width           =   90
      End
   End
   Begin SMBMaker.dbFrame Frame1 
      Height          =   1335
      Left            =   120
      TabIndex        =   0
      Top             =   1035
      Width           =   4560
      _ExtentX        =   8043
      _ExtentY        =   2355
      Caption         =   "Raster Size"
      EAC             =   0   'False
      Begin VB.OptionButton WriteRS 
         BackColor       =   &H0080FFFF&
         Caption         =   "Write fact size"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Index           =   1
         Left            =   255
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   750
         Value           =   -1  'True
         Width           =   3915
      End
      Begin VB.OptionButton WriteRS 
         BackColor       =   &H0080FFFF&
         Caption         =   "Write zero"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Index           =   0
         Left            =   255
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   225
         Width           =   3915
      End
   End
   Begin SMBMaker.dbFrame dbFrame1 
      Height          =   1335
      Left            =   120
      TabIndex        =   10
      Top             =   3480
      Width           =   4560
      _ExtentX        =   8043
      _ExtentY        =   2355
      Caption         =   "Raster writing"
      EAC             =   0   'False
      Begin VB.OptionButton optBottomUp 
         BackColor       =   &H0080FF80&
         Caption         =   "Bottom-up (normal)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   255
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   225
         Value           =   -1  'True
         Width           =   3915
      End
      Begin VB.OptionButton optTopDown 
         BackColor       =   &H0080FF80&
         Caption         =   "Top-down (not normal)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   255
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   750
         Width           =   3915
      End
   End
   Begin SMBMaker.dbButton CancelButton 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   2505
      TabIndex        =   13
      Top             =   4875
      Width           =   1830
      _ExtentX        =   3228
      _ExtentY        =   661
      MouseIcon       =   "frmBitMapSettings.frx":0442
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Others          =   $"frmBitMapSettings.frx":045E
      OthersPresent   =   -1  'True
   End
   Begin SMBMaker.dbButton OkButton 
      Default         =   -1  'True
      Height          =   375
      Left            =   435
      TabIndex        =   9
      Top             =   4875
      Width           =   1830
      _ExtentX        =   3228
      _ExtentY        =   661
      MouseIcon       =   "frmBitMapSettings.frx":04AE
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Others          =   $"frmBitMapSettings.frx":04CA
      OthersPresent   =   -1  'True
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Use these settings only when you really need it. Some of them may cause the saved file to be opened incorrectly by other editors."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   870
      Left            =   120
      TabIndex        =   8
      Top             =   45
      Width           =   4515
   End
End
Attribute VB_Name = "frmFormatBMP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CancelButton_Click()
Me.Tag = "C"
Me.Hide
End Sub

Private Sub Form_Paint()
On Error Resume Next
Me.PaintPicture gBackPicture, 0, 0, Me.ScaleWidth, Me.ScaleHeight
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If UnloadMode = 0 Then
    Cancel = True
    CancelButton_Click
End If
End Sub

Public Function dbValidateControls() As Boolean
Dim tLng As Long
On Error Resume Next
With DpiX
    tLng = CLng(.Text)
    If Err.Number <> 0 Then
        vtBeep
        .SetFocus
        dbValidateControls = False
        Exit Function
    End If
    .Text = IIf(tLng >= 0, CStr(tLng), "&H" + Hex$(tLng))
End With
With DpiY
    tLng = CLng(.Text)
    If Err.Number <> 0 Then
        vtBeep
        .SetFocus
        dbValidateControls = False
        Exit Function
    End If
    .Text = IIf(tLng >= 0, CStr(tLng), "&H" + Hex$(tLng))
End With
dbValidateControls = True
End Function

Private Sub OkButton_Click()
If Not (dbValidateControls) Then Exit Sub
Me.Tag = ""
Me.Hide
End Sub
