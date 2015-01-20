VERSION 5.00
Begin VB.Form frmConfirmExit 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Save"
   ClientHeight    =   2445
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6915
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2445
   ScaleWidth      =   6915
   StartUpPosition =   1  'CenterOwner
   Begin SMBMaker.dbButton btnCancel 
      Cancel          =   -1  'True
      Height          =   315
      Left            =   2062
      TabIndex        =   4
      Top             =   2115
      Width           =   2790
      _ExtentX        =   4921
      _ExtentY        =   556
      MouseIcon       =   "frmConfirmExit.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Others          =   $"frmConfirmExit.frx":001C
      OthersPresent   =   -1  'True
   End
   Begin SMBMaker.dbButton btnSaveToBackUp 
      Height          =   360
      Left            =   4852
      TabIndex        =   3
      Top             =   1740
      Width           =   2040
      _ExtentX        =   3598
      _ExtentY        =   635
      MouseIcon       =   "frmConfirmExit.frx":0069
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Others          =   $"frmConfirmExit.frx":0085
      OthersPresent   =   -1  'True
   End
   Begin SMBMaker.dbButton btnDontSave 
      Height          =   360
      Left            =   3045
      TabIndex        =   2
      Top             =   1740
      Width           =   1800
      _ExtentX        =   3175
      _ExtentY        =   635
      MouseIcon       =   "frmConfirmExit.frx":00D3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Others          =   $"frmConfirmExit.frx":00EF
      OthersPresent   =   -1  'True
   End
   Begin SMBMaker.dbButton btnSaveAs 
      Height          =   360
      Left            =   1485
      TabIndex        =   1
      Top             =   1740
      Width           =   1560
      _ExtentX        =   2752
      _ExtentY        =   635
      MouseIcon       =   "frmConfirmExit.frx":013D
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Others          =   $"frmConfirmExit.frx":0159
      OthersPresent   =   -1  'True
   End
   Begin SMBMaker.dbButton btnSave 
      Default         =   -1  'True
      Height          =   360
      Left            =   22
      TabIndex        =   0
      Top             =   1740
      Width           =   1470
      _ExtentX        =   2593
      _ExtentY        =   635
      MouseIcon       =   "frmConfirmExit.frx":01A7
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Others          =   $"frmConfirmExit.frx":01C3
      OthersPresent   =   -1  'True
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Save?"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1530
      Left            =   330
      TabIndex        =   5
      Top             =   75
      Width           =   6255
   End
End
Attribute VB_Name = "frmConfirmExit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnCancel_Click()
Tag = "C"
Me.Hide
End Sub

Private Sub btnDontSave_Click()
HideW
End Sub

Private Sub btnSave_Click()
MainForm.mnuSaveFile_Click
If Not Len(MainForm.OpenedFileName) Then HideW
End Sub

Private Sub btnSaveAs_Click()
Dim b As Boolean
On Error GoTo eh
MainForm.SaveAuto ShowDialog:=True
'Load frmSave
'frmSave.Show vbModal
'b = Len(frmSave.Tag) > 0
'Unload frmSave
'If b And Not Len(MainForm.OpenedFileName) Then HideW
Exit Sub
eh:
MsgError
End Sub

Private Sub btnSaveToBackUp_Click()
On Error GoTo eh
MainForm.BuildBackup , False
HideW
Exit Sub
eh:
If Err.Number = dbCWS Then Exit Sub
MsgBox Err.Description, vbCritical, Err.Source
End Sub

Public Sub HideW()
Me.Tag = ""
Me.Hide
End Sub

Private Sub Form_Load()
LoadCaptions
End Sub

Private Sub Form_Paint()
Me.PaintPicture gBackPicture, 0, 0, ScaleWidth, ScaleHeight
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If UnloadMode = 0 Then
    Cancel = True
    btnDontSave_Click
End If
End Sub

Public Sub LoadCaptions()
Me.Caption = GRSF(2430)
Label1.Caption = GRSF(1152)
End Sub
