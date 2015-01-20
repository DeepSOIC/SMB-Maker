VERSION 5.00
Begin VB.Form frmAero 
   BackColor       =   &H00E3DFE0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Airbrush"
   ClientHeight    =   3930
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6765
   Icon            =   "frmAero.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3930
   ScaleWidth      =   6765
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin SMBMaker.dbFrame Frame3 
      Height          =   1080
      Left            =   3708
      TabIndex        =   14
      Top             =   1548
      Width           =   2928
      _ExtentX        =   5159
      _ExtentY        =   1905
      Caption         =   "Intensity"
      EAC             =   0   'False
      Begin VB.TextBox aIntens 
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
         Left            =   555
         MaxLength       =   5
         TabIndex        =   15
         Text            =   "5"
         ToolTipText     =   "Points count to be set by one event."
         Top             =   360
         Width           =   600
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00E3DFE0&
         BackStyle       =   0  'Transparent
         Caption         =   "points per event"
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
         Left            =   1215
         TabIndex        =   16
         Top             =   420
         Width           =   1185
      End
   End
   Begin SMBMaker.dbFrame Frame4 
      Height          =   1080
      Left            =   195
      TabIndex        =   11
      Top             =   2715
      Width           =   3300
      _ExtentX        =   5821
      _ExtentY        =   1905
      Caption         =   "Color"
      No3D            =   -1  'True
      EAC             =   0   'False
      Begin VB.OptionButton cOpt 
         BackColor       =   &H0080FFFF&
         Caption         =   "Manual"
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
         Height          =   300
         Index           =   0
         Left            =   420
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   210
         Value           =   -1  'True
         Width           =   2460
      End
      Begin VB.OptionButton cOpt 
         BackColor       =   &H0080FFFF&
         Caption         =   "Random"
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
         Height          =   300
         Index           =   1
         Left            =   420
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   510
         Width           =   2460
      End
   End
   Begin SMBMaker.dbFrame Frame2 
      Height          =   1080
      Left            =   195
      TabIndex        =   8
      Top             =   1545
      Width           =   3300
      _ExtentX        =   5821
      _ExtentY        =   1905
      Caption         =   "Airbrush size"
      EAC             =   0   'False
      Begin VB.CheckBox chkSizePressure 
         BackColor       =   &H00FFC0FF&
         Caption         =   "Pressure sensitive (for tablet)"
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
         Height          =   270
         Left            =   615
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   615
         Width           =   2430
      End
      Begin VB.TextBox aSize 
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
         Left            =   600
         MaxLength       =   3
         TabIndex        =   9
         Text            =   "10"
         ToolTipText     =   "Area radius. Random pixels will be put onto this area."
         Top             =   210
         Width           =   435
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00E3DFE0&
         BackStyle       =   0  'Transparent
         Caption         =   "pixels"
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
         Left            =   1035
         TabIndex        =   10
         Top             =   285
         Width           =   405
      End
   End
   Begin SMBMaker.dbFrame Frame1 
      Height          =   960
      Left            =   144
      TabIndex        =   4
      Top             =   156
      Width           =   3300
      _ExtentX        =   5821
      _ExtentY        =   1693
      Caption         =   "Call airbrush on"
      EAC             =   0   'False
      Begin VB.CheckBox Chk 
         BackColor       =   &H00FFC0FF&
         Caption         =   "Mouse down"
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
         Height          =   300
         Index           =   0
         Left            =   96
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   252
         Value           =   1  'Checked
         Width           =   1524
      End
      Begin VB.CheckBox Chk 
         BackColor       =   &H00FFC0FF&
         Caption         =   "Mouse move"
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
         Height          =   300
         Index           =   1
         Left            =   396
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   552
         Value           =   1  'Checked
         Width           =   2460
      End
      Begin VB.CheckBox Chk 
         BackColor       =   &H00FFC0FF&
         Caption         =   "Mouse up"
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
         Height          =   300
         Index           =   2
         Left            =   1620
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   252
         Value           =   1  'Checked
         Width           =   1524
      End
   End
   Begin VB.TextBox TxtHlp 
      BackColor       =   &H80000018&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000017&
      Height          =   1110
      Left            =   3705
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      TabStop         =   0   'False
      Text            =   "frmAero.frx":014A
      Top             =   2700
      Width           =   2925
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Height          =   1500
      Left            =   3705
      ScaleHeight     =   96
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   96
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   45
      Width           =   1500
   End
   Begin SMBMaker.dbButton CancelButton 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   5340
      TabIndex        =   1
      Top             =   720
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      MouseIcon       =   "frmAero.frx":0150
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Others          =   $"frmAero.frx":016C
      OthersPresent   =   -1  'True
   End
   Begin SMBMaker.dbButton OKButton 
      Default         =   -1  'True
      Height          =   375
      Left            =   5340
      TabIndex        =   0
      Top             =   210
      Width           =   1215
      _ExtentX        =   0
      _ExtentY        =   0
      MouseIcon       =   "frmAero.frx":01BC
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Others          =   $"frmAero.frx":01D8
      OthersPresent   =   -1  'True
   End
End
Attribute VB_Name = "frmAero"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub aIntens_Change()
If Val(aIntens.Text) <= 0 Then aIntens.Text = CStr(1)
If Val(aIntens.Text) > 15400 Then aIntens.Text = CStr(15400)
aIntens.Text = CStr(Int(Val(aIntens.Text)))
UpdatePreview
End Sub

Private Sub aSize_Change()
If Val(aSize.Text) <= 0 Then aSize.Text = CStr(1)
'If Val(aSize.Text) > 100 Then aSize.Text = "100"
aSize.Text = CStr(Int(Val(aSize.Text)))
UpdatePreview
End Sub

Private Sub UpdatePreview()
Picture1.Cls
dbAero Picture1.ScaleWidth \ 2, Picture1.ScaleHeight \ 2, _
       Val(aSize.Text), _
       Val(aIntens.Text), _
       IIf(cOpt(0).Value, vbWhite, -1)
End Sub

Public Sub dbAero(ByVal X As Long, ByVal Y As Long, _
                  ByVal cSize As Long, _
                  ByVal Intens As Long, _
                  ByVal lngColor As Long)
Dim i As Integer, j As Integer, h As Long, rc As Boolean
Dim cSize2 As Long
cSize2 = cSize * cSize
If cSize <= 0 Then Exit Sub
If lngColor = -1 Then rc = True
For h = 1 To Intens
    Do
        i = Int(Rnd(1) * cSize * 2 - cSize + Y)
        j = Int(Rnd(1) * cSize * 2 - cSize + X)
    Loop Until (i - Y) * (i - Y) + (j - X) * (j - X) <= cSize2
    If rc Then lngColor = Rnd * vbWhite
    SetPixel Picture1.hDC, j, i, lngColor
Next h
End Sub


Private Sub cOpt_Click(Index As Integer)
UpdatePreview
End Sub

Private Sub Form_Load()
dbLoadCaptions
End Sub

Private Sub Form_Paint()
Me.PaintPicture LoadResPicture(RES_BackPicture, vbResBitmap), 0, 0, Me.ScaleWidth, Me.ScaleHeight
End Sub

Private Sub Form_Initialize()
'Dim obj As Object, tmp As String, i As Integer, index As Integer
'On Error Resume Next
'For Each obj In Me
'    index = -1
'    index = obj.index
'    Err.Clear
'    tmp = obj.Caption
'    If Err = 0 Then
'        'Debug.Print obj.Name + IIf((index = -1), "", "(" + CStr(index) + ")") + ".Caption = ", """" + tmp + """"
'        For i = 254 To 266
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
'        For i = 254 To 266
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
LoadSettings
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If UnloadMode = vbFormControlMenu Then
    Cancel = True
    Me.Hide
End If
End Sub

Private Sub OkButton_Click()
Dim i As Integer, tmp As Boolean
tmp = False
For i = 0 To Chk.UBound
    If Chk(i).Value = 1 Then tmp = True
Next i
If Not (tmp) Then
    dbMsgBox GRSF(1128), vbInformation 'At least one event must be checked.`Events
    Chk(1).SetFocus
    Exit Sub
End If
Me.Tag = ""
SaveSettings
Me.Hide
End Sub

Function dbValidateControls() As Boolean
Dim tmp As Boolean, i As Integer
tmp = False
For i = 0 To Chk.UBound
    If Chk(i).Value = 1 Then tmp = True
Next i
If tmp Then
    dbMsgBox GRSF(1128), vbInformation 'At least one event must be checked.`Events
    Chk(1).SetFocus
    dbValidateControls = False
    Exit Function
End If
With aIntens
    .Text = CLng(Val(.Text))
    If Val(.Text) > 15000 Then
        vtBeep
        .Text = "15000"
        .SetFocus
        dbValidateControls = False
        Exit Function
    End If
    If Val(.Text) < 1 Then
        vtBeep
        .Text = "1"
        .SetFocus
        dbValidateControls = False
        Exit Function
    End If
End With
With aSize
    .Text = CLng(Val(.Text))
    If Val(.Text) < 1 Then
        .Text = "1"
        vtBeep
        .SetFocus
        dbValidateControls = False
        Exit Function
    End If
    If Val(.Text) > 2000 Then
        .Text = 2000
        .SetFocus
        vtBeep
        dbValidateControls = False
        Exit Function
    End If
End With
End Function

Sub dbLoadCaptions()
Me.Caption = GRSF(2130)
Frame3.Caption = GRSF(254)
aIntens.ToolTipText = GRSF(255)
Label2.Caption = GRSF(256)
Frame2.Caption = GRSF(257)
aSize.ToolTipText = GRSF(258)
Label1.Caption = GRSF(259)
Frame1.Caption = GRSF(260)
Frame1.ToolTipText = GRSF(261)
Chk(2).Caption = GRSF(262)
Chk(1).Caption = GRSF(263)
Chk(0).Caption = GRSF(264)
CancelButton.Caption = GRSF(265)
OkButton.Caption = GRSF(266)
cOpt(0).Caption = GRSF(2151)
cOpt(1).Caption = GRSF(2152)
Frame4.Caption = GRSF(2150)
TxtHlp.Text = GRSF(10073)
'Me.Icon = LoadResPicture(Me.Name, vbResIcon)
End Sub

Public Property Get AColorMode() As Integer
Dim i As Integer
For i = 0 To cOpt.UBound
    If cOpt(i).Value Then AColorMode = i: Exit Property
Next i
AColorMode = -1
End Property

Public Property Let AColorMode(ByVal vvv As Integer)
cOpt(vvv).Value = True
End Property

Public Sub LoadSettings()
Dim MsgText As String, i As Long, Answ As VbMsgBoxResult
On Error GoTo eh
    MsgText = "Invalid aero.intens"
    aIntens.Text = dbGetSetting("Tool", "AeroIntens", 10)
    MsgText = "Invalid aero.size"
    aSize.Text = dbGetSetting("Tool", "AeroSize", 10)
    MsgText = "Invalid aero.useEventDown"
    Chk(0).Value = dbGetSetting("Tool", "AeroEventDown", "1")
    MsgText = "Invalid aero.UseEventMove"
    Chk(1).Value = dbGetSetting("Tool", "AeroEventMove", "1")
    MsgText = "Invalid aero.UseEventUp"
    Chk(2).Value = dbGetSetting("Tool", "AeroEventUp", "1")
    MsgText = "Invalid Aero.ColorMode"
    AColorMode = Val(dbGetSetting("Tool", "AeroColorMode", "0"))
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
dbSaveSetting "Tool", "AeroEventDown", CStr(Chk(0))
dbSaveSetting "Tool", "AeroEventMove", CStr(Chk(1))
dbSaveSetting "Tool", "AeroEventUp", CStr(Chk(2))
dbSaveSetting "Tool", "AeroSize", CStr(Val(aSize.Text))
dbSaveSetting "Tool", "AeroIntens", CStr(Val(aIntens.Text))
dbSaveSetting "Tool", "AeroColorMode", CStr(AColorMode)
End Sub



