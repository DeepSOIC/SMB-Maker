VERSION 5.00
Begin VB.Form frmInput 
   BackColor       =   &H00E3DFE0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Input"
   ClientHeight    =   2025
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6000
   Icon            =   "frmInput.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2025
   ScaleWidth      =   6000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtInp 
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
      Left            =   180
      TabIndex        =   0
      Top             =   1560
      Width           =   5640
   End
   Begin SMBMaker.dbButton btnBrowse 
      Height          =   360
      Left            =   4515
      TabIndex        =   4
      Top             =   1080
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   635
      MouseIcon       =   "frmInput.frx":000C
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Others          =   $"frmInput.frx":0028
      OthersPresent   =   -1  'True
   End
   Begin SMBMaker.dbButton CancelButton 
      Cancel          =   -1  'True
      Height          =   360
      Left            =   4515
      TabIndex        =   2
      Top             =   592
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   635
      MouseIcon       =   "frmInput.frx":007F
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Others          =   $"frmInput.frx":009B
      OthersPresent   =   -1  'True
   End
   Begin SMBMaker.dbButton OKButton 
      Default         =   -1  'True
      Height          =   360
      Left            =   4515
      TabIndex        =   1
      Top             =   105
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   635
      MouseIcon       =   "frmInput.frx":00EE
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Others          =   $"frmInput.frx":010A
      OthersPresent   =   -1  'True
   End
   Begin VB.Label lblMsg 
      BackColor       =   &H00E3DFE0&
      BackStyle       =   0  'Transparent
      Caption         =   "Message text"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1125
      Left            =   180
      TabIndex        =   3
      Top             =   180
      Width           =   3930
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmInput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Const Res_Input = 2230
Option Explicit

Private Sub btnBrowse_Click()
Dim tDir As String
On Error GoTo eh
With CDl
    tDir = .InitDir
    If FolderExists(txtInp.Text) Then
        .InitDir = txtInp.Text
    End If
    .Filter = CurDll + "|" + CurDll
    .OpenFlags = cdlOFNFileMustExist
    .hWndOwner = hWnd
    .ShowOpen
    txtInp = GetDirName(.FileName)
    .InitDir = tDir
End With
Exit Sub
eh:
If Err.Number = dbCWS Then
    Exit Sub
Else
    MsgBox Err.Description, vbCritical, Err.Source
End If
End Sub

Private Sub Form_Activate()
txtInp.SetFocus
End Sub

Private Sub Form_Paint()
Me.PaintPicture LoadResPicture(RES_BackPicture, vbResBitmap), 0, 0, Me.ScaleWidth, Me.ScaleHeight
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If UnloadMode = vbFormControlMenu Then
    Cancel = True
    CancelButton_Click
End If
End Sub

Private Sub OkButton_Click()
Me.Tag = ""
Me.Hide
End Sub

Private Sub CancelButton_Click()
Me.Tag = "C"
Me.Hide
End Sub

Public Sub SetProps(ByVal strMsg As String, ByVal strCaption As String, ByVal strInit As String, ByVal BrowseButton As Boolean, MaxLen As Long)
If strCaption = "" Then
    strCaption = GRSF(Res_Input)
End If
Me.Caption = strCaption
lblMsg.Caption = strMsg
txtInp.MaxLength = MaxLen
txtInp.Text = strInit
btnBrowse.Visible = BrowseButton
End Sub

Public Property Get Text() As String
Text = txtInp.Text
End Property

Public Property Let Text(ByVal Txt As String)
txtInp.Text = Txt
End Property

Private Sub Form_Initialize()
LoadCaptions
End Sub

Private Sub LoadCaptions()
'OKButton.Caption = grsf(2231)
'CancelButton.Caption = grsf(2232)
End Sub

Private Sub txtInp_GotFocus()
With txtInp
    .SelStart = 0
    .SelLength = Len(.Text)
End With
End Sub
