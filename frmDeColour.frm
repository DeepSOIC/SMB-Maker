VERSION 5.00
Begin VB.Form frmDeColour 
   BackColor       =   &H00E3DFE0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Decolour"
   ClientHeight    =   750
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6360
   Icon            =   "frmDeColour.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   750
   ScaleWidth      =   6360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.TextBox Text 
      Alignment       =   2  'Center
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
      Left            =   855
      TabIndex        =   1
      Text            =   "100"
      Top             =   300
      Width           =   3060
   End
   Begin VB.HScrollBar Scroll 
      Height          =   300
      LargeChange     =   10
      Left            =   855
      Max             =   100
      TabIndex        =   0
      Top             =   0
      Value           =   100
      Width           =   3060
   End
   Begin SMBMaker.dbButton CancelButton 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   5145
      TabIndex        =   3
      Top             =   375
      Width           =   1215
      _ExtentX        =   0
      _ExtentY        =   0
      MouseIcon       =   "frmDeColour.frx":1CFA
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Others          =   $"frmDeColour.frx":1D16
      OthersPresent   =   -1  'True
   End
   Begin SMBMaker.dbButton OKButton 
      Default         =   -1  'True
      Height          =   375
      Left            =   5145
      TabIndex        =   2
      Top             =   0
      Width           =   1215
      _ExtentX        =   0
      _ExtentY        =   0
      MouseIcon       =   "frmDeColour.frx":1D66
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Others          =   $"frmDeColour.frx":1D82
      OthersPresent   =   -1  'True
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00E3DFE0&
      BackStyle       =   0  'Transparent
      Caption         =   "%"
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
      Left            =   4005
      TabIndex        =   6
      Top             =   345
      Width           =   165
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00E3DFE0&
      BackStyle       =   0  'Transparent
      Caption         =   "Less colour"
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
      Left            =   3990
      TabIndex        =   5
      Top             =   60
      Width           =   795
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00E3DFE0&
      BackStyle       =   0  'Transparent
      Caption         =   "More colour"
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
      Left            =   0
      TabIndex        =   4
      Top             =   60
      Width           =   840
   End
End
Attribute VB_Name = "frmDeColour"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub CancelButton_Click()
Me.Tag = "c"
Me.Hide
End Sub

Private Sub Form_Initialize()
'Dim obj As Object, i As Integer, index As Integer, tmp As String
'On Error Resume Next
'For Each obj In Me
'    index = -1
'    index = obj.index
'    Err.Clear
'    tmp = obj.Caption
'    If Err = 0 Then
''        Debug.Print obj.Name + IIf((index = -1), "", "(" + CStr(index) + ")") + ".Caption = ", """" + tmp + """"
'        For i = 273 To 276
'            If grsf(i) = tmp Then
'            Debug.Print obj.Name + IIf((index = -1), "", "(" + CStr(index) + ")") + ".Caption = ", "grsf(" + CStr(i) + ")"
'            Exit For
'            End If
'        Next i
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
''        Debug.Print obj.Name + IIf((index = -1), "", "(" + CStr(index) + ")") + ".ToolTipText = ", """" + tmp + """"
'        For i = 273 To 276
'            If grsf(i) = tmp Then
'            Debug.Print obj.Name + IIf((index = -1), "", "(" + CStr(index) + ")") + ".tooltiptext = ", "grsf(" + CStr(i) + ")"
'            Exit For
'            End If
'        Next i
'    Else
'        Err.Clear
'        'Debug.Print "-"
'    End If
'
'Next
'End
dbLoadCaptions
Label1.Left = Scroll.Left - Label1.Width - 60
LoadSettings
End Sub

Public Sub LoadSettings()
Text.Text = dbGetSetting("Effects\DeColour", "Percents", "100")
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If UnloadMode = 0 Then
    Cancel = True
    CancelButton_Click
End If
End Sub

Public Sub SaveSettings()
dbSaveSetting "Effects\DeColour", "Percents", Text.Text
End Sub

Private Sub OkButton_Click()
Me.Tag = ""
SaveSettings
Me.Hide
End Sub

Private Sub Scroll_Change()
Text.Tag = "C"
Text.Text = CStr(Scroll.Value)
Text.Tag = ""
End Sub

Private Sub Scroll_Scroll()
Scroll_Change
End Sub

Private Sub Text_Change()
On Error GoTo eh
If Not Text.Tag = "" Then Exit Sub
Scroll.Tag = "C"
Text.Text = CStr(Val(Text))
Scroll.Value = Int(Val(Text))
Scroll.Tag = ""
Exit Sub
eh:
dbMsgBox GRSF(1133), vbInformation
Text.Text = "100"
End Sub

Sub dbLoadCaptions()
Me.Caption = GRSF(2132)
CancelButton.Caption = GRSF(273)
OkButton.Caption = GRSF(274)
Label2.Caption = GRSF(275)
Label1.Caption = GRSF(276)
'Me.Icon = LoadResPicture(Me.Name, vbResIcon)
End Sub

Private Sub Form_Paint()
Me.PaintPicture LoadResPicture(RES_BackPicture, vbResBitmap), 0, 0, Me.ScaleWidth, Me.ScaleHeight
End Sub

