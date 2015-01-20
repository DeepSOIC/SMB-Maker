VERSION 5.00
Begin VB.Form frmSpin 
   BackColor       =   &H00E3DFE0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Helix"
   ClientHeight    =   3645
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   7080
   Icon            =   "frmSpin.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3645
   ScaleWidth      =   7080
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   1500
      Left            =   2265
      Picture         =   "frmSpin.frx":058A
      ScaleHeight     =   1440
      ScaleWidth      =   1440
      TabIndex        =   12
      ToolTipText     =   $"frmSpin.frx":3254
      Top             =   165
      Width           =   1500
   End
   Begin SMBMaker.dbButton FadeEditor 
      Height          =   540
      Left            =   4185
      TabIndex        =   5
      Top             =   1095
      Width           =   2640
      _ExtentX        =   4657
      _ExtentY        =   953
      MouseIcon       =   "frmSpin.frx":32E9
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Others          =   $"frmSpin.frx":3305
      OthersPresent   =   -1  'True
   End
   Begin SMBMaker.dbFrame Frame2 
      Height          =   1455
      Left            =   195
      TabIndex        =   9
      Top             =   1995
      Width           =   6645
      _ExtentX        =   0
      _ExtentY        =   0
      Caption         =   "Start radius"
      BackColor       =   14933984
      EAC             =   0   'False
      Begin VB.TextBox txtRV 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5055
         TabIndex        =   4
         Text            =   "2"
         ToolTipText     =   "Ratio between the ending radius to the beginning radius. Otherwise, how much is the ending radius greater than the beginning one."
         Top             =   840
         Width           =   1380
      End
      Begin VB.TextBox txtRF 
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
         Left            =   3825
         TabIndex        =   3
         Text            =   "0"
         ToolTipText     =   "The start radius. It can be not integer (delimiter - ""."")."
         Top             =   398
         Width           =   1350
      End
      Begin VB.OptionButton Opt 
         BackColor       =   &H0080FFFF&
         Caption         =   "Depending on ending radius"
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
         Index           =   1
         Left            =   270
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   795
         Width           =   3180
      End
      Begin VB.OptionButton Opt 
         BackColor       =   &H0080FFFF&
         Caption         =   "Fixed"
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
         Index           =   0
         Left            =   270
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   390
         Value           =   -1  'True
         Width           =   3165
      End
      Begin VB.Label LV 
         AutoSize        =   -1  'True
         BackColor       =   &H00E3DFE0&
         BackStyle       =   0  'Transparent
         Caption         =   "Ratio (R/R0):"
         Enabled         =   0   'False
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
         Left            =   4080
         TabIndex        =   11
         Top             =   870
         Width           =   960
      End
      Begin VB.Label LF 
         AutoSize        =   -1  'True
         BackColor       =   &H00E3DFE0&
         BackStyle       =   0  'Transparent
         Caption         =   "pixels."
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
         Left            =   5205
         TabIndex        =   10
         Top             =   465
         Width           =   465
      End
   End
   Begin SMBMaker.dbFrame Frame1 
      Height          =   960
      Left            =   210
      TabIndex        =   8
      Top             =   195
      Width           =   1680
      _ExtentX        =   0
      _ExtentY        =   0
      Caption         =   "Spin count"
      BackColor       =   14933984
      EAC             =   0   'False
      Begin VB.TextBox txtN 
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
         Left            =   330
         TabIndex        =   0
         Text            =   "5"
         ToolTipText     =   "Full-turns count. The value can be not integer (delimiter - ""."")."
         Top             =   345
         Width           =   795
      End
   End
   Begin SMBMaker.dbButton CancelButton 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   5580
      TabIndex        =   7
      Top             =   630
      Width           =   1215
      _ExtentX        =   0
      _ExtentY        =   0
      MouseIcon       =   "frmSpin.frx":3355
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Others          =   $"frmSpin.frx":3371
      OthersPresent   =   -1  'True
   End
   Begin SMBMaker.dbButton OKButton 
      Default         =   -1  'True
      Height          =   375
      Left            =   5580
      TabIndex        =   6
      Top             =   150
      Width           =   1215
      _ExtentX        =   0
      _ExtentY        =   0
      MouseIcon       =   "frmSpin.frx":33BB
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Others          =   $"frmSpin.frx":33D7
      OthersPresent   =   -1  'True
   End
End
Attribute VB_Name = "frmSpin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim prvFDSC As FadeDesc
Option Explicit

Private Sub CancelButton_Click()
Me.Tag = "C"
Me.Hide
End Sub

Private Sub FadeEditor_Click()
Load Pereliv
With Pereliv
    MainForm.SendFadeDesc prvFDSC
    .Show vbModal
    If .Tag = "" Then
        MainForm.ExtractFadeDesc prvFDSC
    End If
End With
Unload Pereliv
End Sub

Private Sub Form_Paint()
Me.PaintPicture LoadResPicture(RES_BackPicture, vbResBitmap), 0, 0, Me.ScaleWidth, Me.ScaleHeight
End Sub

Friend Sub SetProps(ByRef pHSet As HelixSettings, ByRef FadeDsc As FadeDesc)
Me.txtN = Trim$(Str$(pHSet.Numb))
Me.txtRF = Trim$(Str$(pHSet.RFixed))
Me.txtRV = Trim$(Str$(pHSet.RK))
Me.Opt(pHSet.RMode).Value = True
prvFDSC = FadeDsc
End Sub

Friend Sub GetProps(ByRef pHSet As HelixSettings, ByRef FadeDsc As FadeDesc)
pHSet.Numb = Val(txtN.Text)
pHSet.RFixed = Val(txtRF.Text)
pHSet.RK = Val(txtRV.Text)
pHSet.RMode = IIf(Opt(0).Value, 0, 1)
FadeDsc = prvFDSC
End Sub

Private Sub Form_Load()
'Const Action = 1, F = 2059, T = 2069
'Dim obj As Object, i As Integer, index As Integer, tmp As String
'On Error Resume Next
'For Each obj In Me
'    index = -1
'    index = obj.index
'    Err.Clear
'    tmp = obj.Caption
'    If Err = 0 Then
'        If Action = 0 Then
'        Debug.Print obj.Name + IIf((index = -1), "", "(" + CStr(index) + ")") + ".Caption = ", """" + tmp + """"
'        Else
'        For i = F To T
'            If grsf(i) = tmp Then
'            Debug.Print obj.Name + IIf((index = -1), "", "(" + CStr(index) + ")") + ".Caption = ", "grsf(" + CStr(i) + ")"
'            Exit For
'            End If
'        Next i
'        End If
'    Else
'        Err.Clear
'        'Debug.Print "-"
'    End If
'
'    index = -1
'    index = obj.index
'    Err.Clear
'    tmp = obj.ToolTipText
'    If Err = 0 Then
'        If Action = 0 Then
'        Debug.Print obj.Name + IIf((index = -1), "", "(" + CStr(index) + ")") + ".ToolTipText = ", """" + tmp + """"
'        Else
'        For i = F To T
'            If grsf(i) = tmp Then
'            Debug.Print obj.Name + IIf((index = -1), "", "(" + CStr(index) + ")") + ".tooltiptext = ", "grsf(" + CStr(i) + ")"
'            Exit For
'            End If
'        Next i
'        End If
'    Else
'        Err.Clear
'        'Debug.Print "-"
'    End If
'
'Next
'End
dbLoadCaptions
LV.Left = txtRV.Left - LV.Width - 60
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If UnloadMode = vbFormControlMenu Then
    Cancel = True
    Me.Hide
End If
End Sub

Private Sub OkButton_Click()
Me.ValidateControls
Me.Tag = ""
SaveSettings
Me.Hide
End Sub

Private Sub Opt_Click(Index As Integer)
LF.Enabled = Opt(0).Value
txtRF.Enabled = Opt(0).Value
LV.Enabled = Opt(1).Value
txtRV.Enabled = Opt(1).Value
End Sub

Private Sub txtN_Validate(Cancel As Boolean)
With txtN
    If Val(.Text) < 0 Then .Text = "0": Cancel = True
    If Val(.Text) > 100 Then .Text = "100": Cancel = True
    .Text = Str(Val(.Text))
End With
End Sub

Private Sub txtRF_Validate(Cancel As Boolean)
With txtRF
    If Val(.Text) < 0 Then .Text = "0": Cancel = True
    If Val(.Text) > 1000 Then .Text = "1000": Cancel = True
    .Text = Str(Val(.Text))
End With

End Sub

Private Sub txtRV_Validate(Cancel As Boolean)
With txtRV
    If Val(.Text) <= 1 Then .Text = "2": Cancel = True
    If Val(.Text) > 100 Then .Text = "100": Cancel = True
    .Text = Str(Val(.Text))
End With

End Sub

Sub dbLoadCaptions()
Me.Caption = GRSF(2140)
Picture1.ToolTipText = GRSF(2059)
FadeEditor.Caption = GRSF(2060)
Frame2.Caption = GRSF(2061)
txtRV.ToolTipText = GRSF(2062)
txtRF.ToolTipText = GRSF(2063)
Opt(1).Caption = GRSF(2064)
Opt(0).Caption = GRSF(2065)
Frame1.Caption = GRSF(2066)
txtN.ToolTipText = GRSF(2067)
CancelButton.Caption = GRSF(2068)
OkButton.Caption = GRSF(2069)
LF.Caption = GRSF(2147)
LV.Caption = GRSF(2148)
'Me.Icon = LoadResPicture(Me.Name, vbResIcon)
End Sub

Public Sub LoadSettings()
Dim MsgText As String, i As Long, Answ As VbMsgBoxResult
On Error GoTo eh
    MsgText = "Invalid Spin.Count"
    txtN.Text = dbGetSetting("Tool", "SpinCount", "5")
    MsgText = "Invalid Spin.Mode"
    Opt(0).Value = CBool(dbGetSetting("Tool", "SpinR2Mode", "True")): frmSpin.Opt(1).Value = Not (frmSpin.Opt(0).Value)
    MsgText = "Invalid Spin.RFixed"
    txtRF.Text = dbGetSetting("Tool", "SpinR2Fixed", "0")
    MsgText = "Invalid Spin.RVariable"
    txtRV.Text = dbGetSetting("Tool", "SpinR2Variable", "2")
Exit Sub
eh:
    Answ = MsgBox(MsgText, vbCritical Or vbAbortRetryIgnore, "Loading")
    If Answ = vbRetry Then
    Resume
    ElseIf Answ = vbIgnore Then
    Resume Next
    ElseIf Answ = vbAbort Then
    End
    Exit Sub
    End If
End Sub

Public Sub SaveSettings()
Dim i As Long
dbSaveSetting "Tool", "SpinCount", txtN.Text
dbSaveSetting "Tool", "SpinR2Mode", CStr(Opt(0).Value)
dbSaveSetting "Tool", "SpinR2Fixed", txtRF.Text
dbSaveSetting "Tool", "SpinR2Variable", txtRV.Text
End Sub



