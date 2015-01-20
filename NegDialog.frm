VERSION 5.00
Begin VB.Form NegDialog 
   BackColor       =   &H00E3DFE0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Negative method"
   ClientHeight    =   3195
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6030
   Icon            =   "NegDialog.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.OptionButton Opt 
      BackColor       =   &H0080FFFF&
      Caption         =   "Advanced"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Index           =   1
      Left            =   240
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1905
      Width           =   1980
   End
   Begin SMBMaker.dbFrame Frame 
      Height          =   855
      Index           =   1
      Left            =   210
      TabIndex        =   9
      Top             =   1860
      Width           =   4020
      _ExtentX        =   0
      _ExtentY        =   0
      Caption         =   "dbFrame"
      BackColor       =   14933984
      EAC             =   0   'False
      Begin VB.TextBox Text 
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
         Left            =   1965
         MaxLength       =   6
         TabIndex        =   5
         Text            =   "FFFFFF"
         ToolTipText     =   "RGB-pattern for negative. RRGGBB"
         Top             =   480
         Width           =   900
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackColor       =   &H00E3DFE0&
         BackStyle       =   0  'Transparent
         Caption         =   "Use pattern"
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
         Left            =   1005
         TabIndex        =   10
         Top             =   510
         Width           =   855
      End
   End
   Begin VB.OptionButton Opt 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0080FFFF&
      Caption         =   "Standard"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Index           =   0
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   270
      Value           =   -1  'True
      Width           =   1980
   End
   Begin SMBMaker.dbFrame Frame 
      Height          =   1290
      Index           =   0
      Left            =   210
      TabIndex        =   8
      Top             =   240
      Width           =   4020
      _ExtentX        =   0
      _ExtentY        =   0
      Caption         =   "dbFrame"
      BackColor       =   14933984
      EAC             =   0   'False
      Begin VB.CheckBox StChk 
         BackColor       =   &H00FFC0FF&
         Caption         =   "Blue"
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
         Index           =   3
         Left            =   195
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   870
         Value           =   1  'Checked
         Width           =   2580
      End
      Begin VB.CheckBox StChk 
         BackColor       =   &H00FFC0FF&
         Caption         =   "Green"
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
         Index           =   2
         Left            =   195
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   585
         Value           =   1  'Checked
         Width           =   2580
      End
      Begin VB.CheckBox StChk 
         BackColor       =   &H00FFC0FF&
         Caption         =   "Red"
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
         Index           =   1
         Left            =   195
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   300
         Value           =   1  'Checked
         Width           =   2580
      End
   End
   Begin SMBMaker.dbButton CancelButton 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   4665
      TabIndex        =   7
      Top             =   1080
      Width           =   1215
      _ExtentX        =   0
      _ExtentY        =   0
      MouseIcon       =   "NegDialog.frx":030A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Others          =   $"NegDialog.frx":0326
      OthersPresent   =   -1  'True
   End
   Begin SMBMaker.dbButton OKButton 
      Default         =   -1  'True
      Height          =   375
      Left            =   4680
      TabIndex        =   6
      Top             =   525
      Width           =   1215
      _ExtentX        =   0
      _ExtentY        =   0
      MouseIcon       =   "NegDialog.frx":0376
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Others          =   $"NegDialog.frx":0392
      OthersPresent   =   -1  'True
   End
End
Attribute VB_Name = "NegDialog"
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
Me.PaintPicture LoadResPicture(RES_BackPicture, vbResBitmap), 0, 0, Me.ScaleWidth, Me.ScaleHeight
End Sub

Public Sub LoadSettings()
Dim i As Integer
Opt(Val(dbGetSetting("Effects\Negative", "Mode", "0"))).Value = True
Opt_Click (0)
For i = 1 To 3
StChk(i).Value = Val(dbGetSetting("Effects\Negative", "Chk" + CStr(i), "1"))
Next i
Text = dbGetSetting("Effects\Negative", "Advanced")
End Sub

Private Sub Form_Initialize()
LoadSettings
'Const Action = 1, F = 2088, T = 2096
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
Text.Left = Label.Left + Label.Width
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim i As Integer
If UnloadMode = VBRUN.QueryUnloadConstants.vbFormControlMenu Then
    Cancel = True
    CancelButton_Click
End If
End Sub

Public Sub SaveSettings()
Dim i As Integer
    dbSaveSetting "Effects\Negative", "Mode", Abs(Opt(1).Value)
    For i = 1 To 3
        dbSaveSetting "Effects\Negative", "Chk" + CStr(i), StChk(i).Value
    Next i
    dbSaveSetting "Effects\Negative", "Advanced", Text.Text
End Sub

'Private Sub Form_Terminate()
'Form_QueryUnload 0, vbFormCode
'End Sub
'
Private Sub OkButton_Click()
Me.Tag = ""
SaveSettings
Me.Hide
End Sub

Public Function GetMask() As Long
Dim tmp As String, i As Integer
If Opt(0).Value Then
    tmp = ""
    For i = 3 To 1 Step -1
        If CBool(StChk(i).Value) Then tmp = tmp + "FF" Else tmp = tmp + "00"
    Next i
    GetMask = CLng("&H" + tmp)
Else
    Err.Clear
    On Error Resume Next
    GetMask = CLng("&H" + Text.Text)
    If Err.Number <> 0 Then
        dbMsgBox GRSF(1129), vbInformation
        GetMask = &HFFFFFF
        Exit Function
    End If
    
End If
End Function

Private Sub Opt_Click(Index As Integer)
Dim i As Integer
For i = 0 To 1
Frame(i).Enabled = Opt(i).Value
Next i
Text.Enabled = Opt(1).Value
Label.Enabled = Opt(1).Value
For i = StChk.lBound To StChk.UBound
    StChk(i).Enabled = Opt(0).Value
Next i
End Sub

Sub dbLoadCaptions()
Me.Caption = GRSF(2144)
Opt(1).Caption = GRSF(2088)
Text.ToolTipText = GRSF(2089)
Label.Caption = GRSF(2090)
Opt(0).Caption = GRSF(2091)
StChk(3).Caption = GRSF(2092)
StChk(2).Caption = GRSF(2093)
StChk(1).Caption = GRSF(2094)
CancelButton.Caption = GRSF(2095)
OkButton.Caption = GRSF(2096)
'Me.Icon = LoadResPicture(Me.Name, vbResIcon)
End Sub
