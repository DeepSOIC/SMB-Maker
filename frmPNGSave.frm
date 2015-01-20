VERSION 5.00
Begin VB.Form frmFormatPNG 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PNG format settings"
   ClientHeight    =   1575
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4050
   Icon            =   "frmPNGSave.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   105
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   270
   StartUpPosition =   1  'CenterOwner
   Begin SMBMaker.dbFrame dbFrame2 
      Height          =   1050
      Left            =   30
      TabIndex        =   1
      Top             =   15
      Width           =   3990
      _ExtentX        =   7038
      _ExtentY        =   1852
      Caption         =   "Options"
      EAC             =   0   'False
      Begin VB.CheckBox chkInterlace 
         BackColor       =   &H00FFC0FF&
         Caption         =   "Interlaced"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   765
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   330
         Width           =   2565
      End
   End
   Begin SMBMaker.dbButton CancelButton 
      Cancel          =   -1  'True
      Height          =   420
      Left            =   2064
      TabIndex        =   3
      Top             =   1116
      Width           =   1932
      _ExtentX        =   3413
      _ExtentY        =   741
      MouseIcon       =   "frmPNGSave.frx":058A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Others          =   $"frmPNGSave.frx":05A6
      OthersPresent   =   -1  'True
   End
   Begin SMBMaker.dbButton OkButton 
      Default         =   -1  'True
      Height          =   420
      Left            =   48
      TabIndex        =   0
      Top             =   1116
      Width           =   1932
      _ExtentX        =   3413
      _ExtentY        =   741
      MouseIcon       =   "frmPNGSave.frx":05F6
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Others          =   $"frmPNGSave.frx":0612
      OthersPresent   =   -1  'True
   End
End
Attribute VB_Name = "frmFormatPNG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CancelButton_Click()
Me.Tag = "C"
Me.Hide
End Sub

Private Sub Form_Paint()
On Error Resume Next
Me.PaintPicture gBackPicture, 0, 0, ScaleWidth, ScaleHeight
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

Public Property Get Interlaced() As Boolean
Interlaced = chkInterlace.Value = vbChecked
End Property

Public Property Let Interlaced(ByVal vNewValue As Boolean)
chkInterlace.Value = IIf(vNewValue, vbChecked, vbUnchecked)
End Property
