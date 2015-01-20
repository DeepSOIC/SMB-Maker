VERSION 5.00
Begin VB.Form frmFormatJPEG 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "JPEG format settings"
   ClientHeight    =   3105
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   2475
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   204
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   207
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   165
   StartUpPosition =   1  'CenterOwner
   Begin SMBMaker.dbFrame dbFrame1 
      Height          =   1320
      Left            =   112
      TabIndex        =   2
      Top             =   1200
      Width           =   2250
      _ExtentX        =   3969
      _ExtentY        =   2328
      Caption         =   "Write mode"
      EAC             =   0   'False
      Begin VB.OptionButton optProgressiveOn 
         BackColor       =   &H0080FFFF&
         Caption         =   "Progressive"
         Height          =   330
         Left            =   285
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Progressive mode is better for Web. Compression is a bit lower, but image can be displayed while downloading."
         Top             =   735
         Value           =   -1  'True
         Width           =   1710
      End
      Begin VB.OptionButton optProgressiveOff 
         BackColor       =   &H0080FFFF&
         Caption         =   "Normal"
         Height          =   330
         Left            =   285
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Normal mode is best for local storage. It provides faster loading and better compression."
         Top             =   390
         Width           =   1710
      End
   End
   Begin SMBMaker.ctlNumBox nmbQuality 
      Height          =   555
      Left            =   457
      TabIndex        =   1
      Top             =   420
      Width           =   1560
      _ExtentX        =   2752
      _ExtentY        =   979
      Value           =   75
      Min             =   1
      Max             =   100
      NumType         =   2
      HorzMode        =   0   'False
      EditName        =   $"frmJPEG.frx":0000
      NLn             =   0
   End
   Begin SMBMaker.dbButton OkButton 
      Default         =   -1  'True
      Height          =   375
      Left            =   187
      TabIndex        =   6
      Top             =   2625
      Width           =   1050
      _ExtentX        =   1852
      _ExtentY        =   661
      MouseIcon       =   "frmJPEG.frx":00F7
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Others          =   $"frmJPEG.frx":0113
      OthersPresent   =   -1  'True
   End
   Begin SMBMaker.dbButton CancelButton 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   1237
      TabIndex        =   5
      Top             =   2625
      Width           =   1050
      _ExtentX        =   1852
      _ExtentY        =   661
      MouseIcon       =   "frmJPEG.frx":015F
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Others          =   $"frmJPEG.frx":017B
      OthersPresent   =   -1  'True
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Quality"
      Height          =   225
      Left            =   457
      TabIndex        =   0
      Top             =   180
      Width           =   1560
   End
End
Attribute VB_Name = "frmFormatJPEG"
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
LoadSettings
End Sub

Private Sub Form_Paint()
On Error Resume Next
Me.PaintPicture gBackPicture, 0, 0, ScaleWidth, ScaleHeight
End Sub

Private Sub OkButton_Click()
SaveSettings
Me.Tag = ""
Me.Hide
End Sub

Friend Sub GetInfo(ByRef jpInfo As vtJPEGInfo)
jpInfo.Quality = nmbQuality.Value
jpInfo.Progressive = optProgressiveOn.Value
End Sub

Public Sub SaveSettings()
dbSaveSettingEx "Formats\JPEG", "Quality", nmbQuality.Value
dbSaveSettingEx "Formats\JPEG", "Progressive", optProgressiveOn.Value
End Sub

Public Sub LoadSettings()
Dim bln As Boolean
On Error Resume Next
nmbQuality = dbGetSettingEx("Formats\JPEG", "Quality", vbInteger, 75)
bln = dbGetSettingEx("Formats\JPEG", "Progressive", vbBoolean, False)
optProgressiveOn.Value = bln
optProgressiveOff.Value = Not bln
End Sub
