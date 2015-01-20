VERSION 5.00
Begin VB.Form frmColorReplace 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Replace Colors"
   ClientHeight    =   1620
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   2580
   Icon            =   "frmColorReplace.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1620
   ScaleWidth      =   2580
   StartUpPosition =   1  'CenterOwner
   Begin SMBMaker.ctlColor clrReplace 
      Height          =   360
      Left            =   1395
      TabIndex        =   6
      Top             =   630
      Width           =   360
      _ExtentX        =   635
      _ExtentY        =   635
   End
   Begin SMBMaker.ctlResourcizator Resr1 
      Left            =   1830
      Top             =   585
      _ExtentX        =   1164
      _ExtentY        =   741
      ResID           =   9918
   End
   Begin VB.TextBox txtSens 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1935
      MaxLength       =   5
      TabIndex        =   2
      Text            =   "0"
      Top             =   0
      Width           =   645
   End
   Begin SMBMaker.ctlColor clrFind 
      Height          =   360
      Left            =   1410
      TabIndex        =   7
      Top             =   30
      Width           =   360
      _ExtentX        =   635
      _ExtentY        =   635
   End
   Begin SMBMaker.dbButton OkButton 
      Default         =   -1  'True
      Height          =   420
      Left            =   0
      TabIndex        =   4
      Top             =   1200
      Width           =   1290
      _ExtentX        =   2275
      _ExtentY        =   741
      MouseIcon       =   "frmColorReplace.frx":0442
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Others          =   $"frmColorReplace.frx":045E
      OthersPresent   =   -1  'True
   End
   Begin SMBMaker.dbButton CancelButton 
      Cancel          =   -1  'True
      Height          =   420
      Left            =   1290
      TabIndex        =   5
      Top             =   1200
      Width           =   1290
      _ExtentX        =   2275
      _ExtentY        =   741
      MouseIcon       =   "frmColorReplace.frx":04AA
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Others          =   $"frmColorReplace.frx":04C6
      OthersPresent   =   -1  'True
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Replace with:"
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
      Left            =   60
      TabIndex        =   3
      Top             =   705
      Width           =   1350
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "±"
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
      Left            =   1815
      TabIndex        =   1
      Top             =   60
      Width           =   120
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Find Color:"
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
      Left            =   60
      TabIndex        =   0
      Top             =   75
      Width           =   1365
   End
   Begin VB.Menu mnuPP 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu mnuFromACol 
         Caption         =   "Fill From Fore Color"
         Index           =   1
      End
      Begin VB.Menu mnuFromACol 
         Caption         =   "Fill From Back Color"
         Index           =   2
      End
   End
End
Attribute VB_Name = "frmColorReplace"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Clicked As Integer

Private Sub CancelButton_Click()
Me.Tag = "C"
Me.Hide
End Sub

Private Sub Form_Load()
LoadCaptions
LoadSettings
End Sub

Public Sub LoadCaptions()
Resr1.LoadCaptions
'Me.Caption = GRSF(2279)
'Label1.Caption = GRSF(2280)
'Label3.Caption = GRSF(2281)
'mnuFromACol(1).Caption = GRSF(2225)
'mnuFromACol(2).Caption = GRSF(2226)
'txtSens.ToolTipText = GRSF(2282)
End Sub

Private Sub Form_Paint()
Me.PaintPicture gBackPicture, 0, 0, ScaleWidth, ScaleHeight
End Sub

Public Function dbValidateControls() As Boolean
Dim tInt As Integer
On Error Resume Next
With txtSens
    Err.Clear
    tInt = CInt(txtSens)
    If tInt < 0 Or tInt > 255 * 3 Or Err.Number <> 0 Then
        .SetFocus
        vtBeep
        dbValidateControls = False
        Exit Function
    End If
    txtSens = CStr(tInt)
End With
dbValidateControls = True
End Function


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If UnloadMode = 0 Then
    Cancel = True
    CancelButton_Click
End If
End Sub

Public Sub SaveSettings()
dbSaveSetting "Effects\Replace", "ColorFind", "&H" + Hex$(clrFind.Color)
dbSaveSetting "Effects\Replace", "ColorReplace", "&H" + Hex$(clrReplace.Color)
dbSaveSetting "Effects\Replace", "Sensitivity", txtSens.Text
End Sub

Public Sub LoadSettings()
Dim Answ As VbMsgBoxResult, MsgText As String
On Error GoTo eh
    MsgText = "Bad ColorFind"
    clrFind.Color = &HFFFFFF And CLng(dbGetSetting("Effects\Replace", "ColorFind", "&H0"))
    MsgText = "Bad ColorReplace"
    clrReplace.Color = &HFFFFFF And CLng(dbGetSetting("Effects\Replace", "ColorReplace", "&H0"))
    MsgText = "Bad Sensivity"
    SetSens CInt(dbGetSetting("Effects\Replace", "Sensitivity", "0"))
    
Exit Sub
eh:
If Err.Number = dbCWS Then
    Resume Next
Else
    Answ = MsgBox(MsgText, vbCritical Or vbAbortRetryIgnore, "Error")
    Select Case Answ
    Case vbAbort
        End
    Case vbRetry
        Resume
    Case vbIgnore
        Resume Next
    End Select
End If
End Sub

Public Sub SetSens(ByVal nSens As Integer)
If nSens < 0 Or nSens > 255 * 3 Then
Err.Raise 5, "SetSens", "Invalid argument"
End If
txtSens.Text = CStr(nSens)
End Sub

Private Sub mnuFromACol_Click(Index As Integer)
Dim Col As Long
Col = MainForm.ActiveColor(Index).BackColor
If Clicked = 0 Then
    clrFind.Color = Col
Else
    clrReplace.Color = Col
End If
End Sub

Private Sub OkButton_Click()
If Not (dbValidateControls) Then Exit Sub
Me.Tag = ""
SaveSettings
Me.Hide
End Sub
