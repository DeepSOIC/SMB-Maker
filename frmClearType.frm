VERSION 5.00
Begin VB.Form frmClearType 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ClearType"
   ClientHeight    =   1125
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   2535
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1125
   ScaleWidth      =   2535
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   345
      Top             =   360
   End
   Begin SMBMaker.dbButton btnCancel 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   750
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   661
      MouseIcon       =   "frmClearType.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Others          =   $"frmClearType.frx":001C
      OthersPresent   =   -1  'True
   End
   Begin SMBMaker.dbButton btnAnti 
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   375
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   661
      MouseIcon       =   "frmClearType.frx":006C
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Others          =   $"frmClearType.frx":0088
      OthersPresent   =   -1  'True
   End
   Begin SMBMaker.dbButton btnNormal 
      Default         =   -1  'True
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   661
      MouseIcon       =   "frmClearType.frx":00D7
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Others          =   $"frmClearType.frx":00F3
      OthersPresent   =   -1  'True
   End
End
Attribute VB_Name = "frmClearType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Anti As Boolean

Private Sub btnAnti_Click()
Anti = True
OK
End Sub

Private Sub btnCancel_Click()
Me.Tag = "C"
Me.Hide
End Sub

Private Sub btnNormal_Click()
Anti = False
OK
End Sub

Public Sub OK()
SaveSettings
Me.Tag = ""
Me.Hide
End Sub

Public Sub SaveSettings()
dbSaveSettingEx "Effects\ClearType", "AntiMode", Anti
End Sub

Public Sub LoadSettings()
Anti = dbGetSettingEx("Effects\ClearType", "AntiMode", vbBoolean, False)
End Sub

Private Sub Form_Load()
LoadSettings
Timer1.Enabled = True
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If UnloadMode = 0 Then
    Cancel = True
    btnCancel_Click
End If
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
If Anti Then btnAnti.SetFocus Else btnNormal.SetFocus
Timer1.Enabled = Err.Number <> 0
End Sub
