VERSION 5.00
Begin VB.Form frmCapture 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Capture Window"
   ClientHeight    =   1935
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   3345
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1935
   ScaleWidth      =   3345
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin SMBMaker.dbFrame dbFrame1 
      Height          =   1530
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   3345
      _ExtentX        =   5900
      _ExtentY        =   2699
      Caption         =   "Capture what"
      ResID           =   2355
      EAC             =   0   'False
      Begin VB.OptionButton OptMode 
         BackColor       =   &H0080FFFF&
         Caption         =   "Entire screen (5 secs.)"
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
         Height          =   345
         Index           =   2
         Left            =   105
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   990
         Width           =   3060
      End
      Begin VB.OptionButton OptMode 
         BackColor       =   &H0080FFFF&
         Caption         =   "Active Window (After 5 seconds)"
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
         Height          =   345
         Index           =   1
         Left            =   105
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   645
         Width           =   3060
      End
      Begin VB.OptionButton OptMode 
         BackColor       =   &H0080FFFF&
         Caption         =   "Point by mouse"
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
         Height          =   345
         Index           =   0
         Left            =   105
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   300
         Value           =   -1  'True
         Width           =   3060
      End
   End
   Begin SMBMaker.dbButton OkButton 
      Height          =   396
      Left            =   0
      TabIndex        =   0
      Top             =   1536
      Width           =   1512
      _ExtentX        =   2672
      _ExtentY        =   688
      MouseIcon       =   "frmCapture.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Others          =   $"frmCapture.frx":001C
      OthersPresent   =   -1  'True
   End
   Begin SMBMaker.dbButton CancelButton 
      Cancel          =   -1  'True
      Height          =   390
      Left            =   1815
      TabIndex        =   1
      Top             =   1545
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   688
      MouseIcon       =   "frmCapture.frx":006B
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Others          =   $"frmCapture.frx":0087
      OthersPresent   =   -1  'True
   End
End
Attribute VB_Name = "frmCapture"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub CancelButton_Click()
Me.Tag = "C"
Me.Hide
End Sub

Private Sub Form_Load()
dbLoadCaptions
LoadSettings
End Sub

Private Sub Form_Paint()
On Error Resume Next
Me.PaintPicture gBackPicture, 0, 0, Me.ScaleWidth, Me.ScaleHeight
End Sub

Private Sub OkButton_Click()
If optMode(0).Value Then
    dbMsgBox 1198, vbInformation
End If
SaveSettings
Me.Tag = vbNullString
Me.Hide
End Sub

Sub dbLoadCaptions()
Me.Caption = GRSF(2354)
optMode(0).Caption = GRSF(2356)
optMode(1).Caption = GRSF(2357)
optMode(2).Caption = GRSF(2358)
optMode(0).ToolTipText = GRSF(2361)
optMode(1).ToolTipText = GRSF(2362)
optMode(2).ToolTipText = GRSF(2363)
End Sub

Function GetMode() As Long
Dim i As Long
For i = optMode.lBound To optMode.UBound
    If optMode(i).Value Then Exit For
Next i
If i = optMode.UBound + 1 Then
    GetMode = -1
Else
    GetMode = i
End If
End Function

Sub SetMode(ByVal NewMode As Long)
On Error Resume Next
optMode(NewMode).Value = True
End Sub

Public Sub LoadSettings()
SetMode dbGetSetting("Options", "CaptureMode", CStr(0))
End Sub

Public Sub SaveSettings()
dbSaveSetting "Options", "CaptureMode", CStr(GetMode)
End Sub
