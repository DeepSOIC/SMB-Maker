VERSION 5.00
Begin VB.Form SelOpts 
   BackColor       =   &H00E3DFE0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Selection preferences"
   ClientHeight    =   5460
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5400
   Icon            =   "SelOpts.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5460
   ScaleWidth      =   5400
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox Combo1 
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
      ItemData        =   "SelOpts.frx":0442
      Left            =   1493
      List            =   "SelOpts.frx":0450
      Style           =   2  'Dropdown List
      TabIndex        =   15
      Top             =   4335
      Width           =   2415
   End
   Begin SMBMaker.dbFrame Frame1 
      Height          =   3975
      Left            =   90
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   90
      Width           =   5220
      _ExtentX        =   9208
      _ExtentY        =   7011
      Caption         =   "Mode"
      BackColor       =   14933984
      EAC             =   0   'False
      Begin VB.CommandButton btnTransColor 
         BackColor       =   &H00000000&
         Height          =   525
         Left            =   3495
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   2715
         Width           =   600
      End
      Begin VB.OptionButton Option 
         BackColor       =   &H0080FFFF&
         Caption         =   "Overlay"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   12
         Left            =   90
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   2985
         Width           =   3405
      End
      Begin VB.OptionButton Option 
         BackColor       =   &H0080FFFF&
         Caption         =   "Mix channels"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   11
         Left            =   90
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   3525
         Width           =   3405
      End
      Begin VB.TextBox TransR 
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
         Left            =   3495
         TabIndex        =   11
         Text            =   "50"
         Top             =   555
         Width           =   564
      End
      Begin VB.OptionButton Option 
         BackColor       =   &H0080FFFF&
         Caption         =   "With AlphaRGB-channel"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   10
         Left            =   90
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   3255
         Width           =   3405
      End
      Begin VB.OptionButton Option 
         BackColor       =   &H0080FFFF&
         Caption         =   "Transparent"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   9
         Left            =   90
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   2715
         Width           =   3405
      End
      Begin VB.OptionButton Option 
         BackColor       =   &H0080FFFF&
         Caption         =   "Semi-transparent:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   276
         Index           =   1
         Left            =   90
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "<color> = (<image color> + <selection color>) / 2"
         Top             =   555
         Width           =   3405
      End
      Begin VB.OptionButton Option 
         BackColor       =   &H0080FFFF&
         Caption         =   "EQV"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   8
         Left            =   90
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "<color> = <image color> EQV <color in selection>"
         Top             =   2445
         Width           =   3405
      End
      Begin VB.OptionButton Option 
         BackColor       =   &H0080FFFF&
         Caption         =   "NOT"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   7
         Left            =   90
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "<color> = NOT <image color> "
         Top             =   2175
         Width           =   3405
      End
      Begin VB.OptionButton Option 
         BackColor       =   &H0080FFFF&
         Caption         =   "IMP"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   6
         Left            =   90
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "<color> = <image color> IMP <color in selection>"
         Top             =   1905
         Width           =   3405
      End
      Begin VB.OptionButton Option 
         BackColor       =   &H0080FFFF&
         Caption         =   "XOR"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   5
         Left            =   90
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "<color> = <image color> XOR <color in selection>"
         Top             =   1635
         Width           =   3405
      End
      Begin VB.OptionButton Option 
         BackColor       =   &H0080FFFF&
         Caption         =   "AND"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   4
         Left            =   90
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "<color> = <image color> AND <color in selection>"
         Top             =   1365
         Width           =   3405
      End
      Begin VB.OptionButton Option 
         BackColor       =   &H0080FFFF&
         Caption         =   "OR"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   3
         Left            =   90
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "<color> = <image color> OR <color in selection>"
         Top             =   1095
         Width           =   3405
      End
      Begin VB.OptionButton Option 
         BackColor       =   &H0080FFFF&
         Caption         =   "Add"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   2
         Left            =   90
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "<color> = <image color> + <color in selection>"
         Top             =   825
         Width           =   3405
      End
      Begin VB.OptionButton Option 
         BackColor       =   &H0080FFFF&
         Caption         =   "Replace"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   0
         Left            =   90
         Style           =   1  'Graphical
         TabIndex        =   0
         ToolTipText     =   "<color> = <color in selection>"
         Top             =   285
         Width           =   3405
      End
      Begin SMBMaker.dbButton btnSuperView 
         Height          =   270
         Left            =   3495
         TabIndex        =   22
         Top             =   3255
         Width           =   570
         _ExtentX        =   1005
         _ExtentY        =   476
         MouseIcon       =   "SelOpts.frx":046C
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Others          =   $"SelOpts.frx":0488
         OthersPresent   =   -1  'True
      End
      Begin SMBMaker.dbButton btnEditRGB 
         Height          =   270
         Left            =   3495
         TabIndex        =   20
         Top             =   3525
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   476
         MouseIcon       =   "SelOpts.frx":04D7
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Others          =   $"SelOpts.frx":04F3
         OthersPresent   =   -1  'True
      End
      Begin SMBMaker.dbButton btnEdit 
         Height          =   270
         Left            =   4065
         TabIndex        =   13
         Top             =   3255
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   476
         MouseIcon       =   "SelOpts.frx":0545
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Others          =   $"SelOpts.frx":0561
         OthersPresent   =   -1  'True
      End
      Begin SMBMaker.dbButton btnChooseTR 
         Height          =   270
         Left            =   4680
         TabIndex        =   14
         Top             =   3255
         Width           =   390
         _ExtentX        =   688
         _ExtentY        =   476
         MouseIcon       =   "SelOpts.frx":05B3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Others          =   $"SelOpts.frx":05CF
         OthersPresent   =   -1  'True
      End
   End
   Begin SMBMaker.dbButton Command1 
      Cancel          =   -1  'True
      Default         =   -1  'True
      Height          =   495
      Left            =   593
      TabIndex        =   17
      Top             =   4830
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   873
      MouseIcon       =   "SelOpts.frx":061D
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Others          =   $"SelOpts.frx":0639
      OthersPresent   =   -1  'True
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00E3DFE0&
      BackStyle       =   0  'Transparent
      Caption         =   "Stretch mode:"
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
      Left            =   593
      TabIndex        =   18
      Top             =   4110
      Width           =   4215
   End
   Begin VB.Menu ppMnu 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu mnuFromAC 
         Caption         =   "Fill from ForeColor"
         Index           =   1
      End
      Begin VB.Menu mnuFromAC 
         Caption         =   "Fill from BackColor"
         Index           =   2
      End
   End
End
Attribute VB_Name = "SelOpts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim TransColor As Long
Dim MixMatrix() As Double
Dim pIsText As Boolean


Function CheckInfo() As Boolean
Dim tmp As Boolean
tmp = True
With TransR
    .Text = Val(.Text)
    If Val(.Text) > 100 Then .Text = Trim(Str(100))
    If Val(.Text) < 0 Then .Text = Trim(Str(0))
End With
CheckInfo = tmp
If Not tmp Then vtBeep
End Function

Private Sub btnChooseTR_Click()
Dim File As String, tData() As Long
Dim Alpha() As Long ' not used
On Error GoTo eh
vtLoadPicture tData, Alpha, "", ShowDialog:=True, Purpose:="SelTrans"
MainForm.SetTransData tData
Exit Sub
eh:
If Err.Number = dbCWS Then
    Exit Sub
Else
    MsgBox Err.Description, vbCritical
End If
End Sub

Private Sub btnEdit_Click()
'If Not IsDataFull(TransData) Then
'    ReDim TransData(0 To 31, 0 To 31)
'End If
Dim tData() As Long
On Error GoTo eh
tData = TransOrigData
EditPicture tData
MainForm.SetTransData tData
Exit Sub
eh:
MsgError
End Sub

Private Sub btnEditRGB_Click()
Load frmMatrixMix
With frmMatrixMix
    .SetMatrix MixMatrix
    .Show vbModal
    If .Tag = "" Then
        .GetMatrix MixMatrix
    End If
End With
Unload frmMatrixMix
End Sub

Private Sub btnSuperView_Click()
ViewImage TransOrigData, "SelTrans"
End Sub

Private Sub Form_Load()
Dim Names() As String
Dim Descs() As String
LoadStretches Combo1, Names, Descs
End Sub

Private Sub Form_Paint()
Me.PaintPicture LoadResPicture(RES_BackPicture, vbResBitmap), 0, 0, Me.ScaleWidth, Me.ScaleHeight
End Sub

Private Sub btnTransColor_Click()
On Error GoTo CWS
TransColor = CDl.PickColor(TransColor, True)
btnTransColor.BackColor = TransColor
Exit Sub
CWS:
If Err.Number <> dbCWS Then
    MsgBox Err.Description
End If
End Sub

Private Sub btnTransColor_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
    PopupMenu ppMnu
End If
End Sub

Private Sub Command1_Click()
If Not CheckInfo Then Exit Sub
SaveSettings
Me.Hide
End Sub

Private Sub Form_Initialize()
dbLoadCaptions
On Error Resume Next
Me.Option(Val(dbGetSetting("Tool", "SelMode", 0))).Value = True
Me.TransR.Text = dbGetSetting("Tool", "SelTranspRatio", "50")
Me.Combo1.ListIndex = Val(dbGetSetting("Tool", "SelStretchMode", "0"))
End Sub

Friend Sub SetProps(ByVal Mode As dbSelMode, _
                    ByVal pStretchMode As eStretchMode, _
                    ByVal pTransColor As Long, _
                    ByVal TransRatio As Single, _
                    ByRef SelMatrix() As Double, _
                    ByVal IsText As Boolean)
Dim i As Long
Me.Option(Mode).Value = True
Me.TransR.Text = dbCStr(TransRatio * 100#)
TransColor = pTransColor
btnTransColor.BackColor = TransColor
For i = 0 To Combo1.ListCount - 1
    If Combo1.ItemData(i) = pStretchMode Then
        Combo1.ListIndex = i
        Exit For
    End If
Next i
MixMatrix = SelMatrix
pIsText = IsText
Dim Ctr As Control
On Error Resume Next
For Each Ctr In Me
  Ctr.Enabled = Not IsText
Next
If IsText Then
  
  Frame1.Enabled = True
  Frame1.EnableAllControls False
  btnSuperView.Enabled = True
  btnEdit.Enabled = True
  btnChooseTR.Enabled = True
  Command1.Enabled = True
End If
End Sub

Friend Sub GetProps(ByRef Mode As dbSelMode, _
                    ByRef pStretchMode As eStretchMode, _
                    ByRef pTransColor As Long, _
                    ByRef TransRatio As Single, _
                    ByRef SelMatrix() As Double)
Mode = GetMode
pStretchMode = Combo1.ItemData(Combo1.ListIndex)
pTransColor = TransColor
TransRatio = dbVal(TransR.Text, vbDouble) / 100#
SelMatrix = MixMatrix
End Sub

Friend Function GetMode() As dbSelMode
Dim i As Integer
For i = Me.Option.lBound To Me.Option.UBound
    If Me.Option(i).Value Then GetMode = i: Exit For
Next i
If i = Me.Option.UBound + 1 Then GetMode = -1
End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim i As Integer
If UnloadMode = VBRUN.QueryUnloadConstants.vbFormControlMenu Then
    Cancel = True
    Command1_Click
End If
End Sub

Public Sub SaveSettings()
Dim i As Long
    For i = Me.Option.lBound To Me.Option.UBound
    If Me.Option(i).Value Then dbSaveSetting "Tool", "SelMode", CStr(i)
    Next i
    dbSaveSetting "Tool", "SelTranspRatio", Me.TransR.Text
    dbSaveSetting "Tool", "SelTransColor", "&H" + Hex$(btnTransColor.BackColor)
    dbSaveSetting "Tool", "SelStretchMode", CStr(Combo1.ListIndex)
End Sub

Sub dbLoadCaptions()
Me.Caption = GRSF(2146)
Command1.Caption = GRSF(2109)
Frame1.Caption = GRSF(2110)
Me.Option(8).Caption = GRSF(2111)
Me.Option(8).ToolTipText = GRSF(2112)
Me.Option(7).Caption = GRSF(2113)
Me.Option(7).ToolTipText = GRSF(2114)
Me.Option(6).Caption = GRSF(2115)
Me.Option(6).ToolTipText = GRSF(2116)
Me.Option(5).Caption = GRSF(2117)
Me.Option(5).ToolTipText = GRSF(2118)
Me.Option(4).Caption = GRSF(2119)
Me.Option(4).ToolTipText = GRSF(2120)
Me.Option(3).Caption = GRSF(2121)
Me.Option(3).ToolTipText = GRSF(2122)
Me.Option(2).Caption = GRSF(2123)
Me.Option(2).ToolTipText = GRSF(2124)
Me.Option(1).Caption = GRSF(2125)
Me.Option(1).ToolTipText = GRSF(2126)
Me.Option(0).Caption = GRSF(2127)
Me.Option(0).ToolTipText = GRSF(2128)
Me.Option(9).Caption = GRSF(2223)
Me.Option(9).ToolTipText = GRSF(2224)
Me.Option(10).Caption = GRSF(2257)
Me.Option(10).ToolTipText = GRSF(2258)
mnuFromAC(1).Caption = GRSF(2225)
mnuFromAC(2).Caption = GRSF(2226)
Label1.Caption = GRSF(2240)
'Me.Icon = LoadResPicture(Me.Name, vbResIcon)
End Sub

Private Sub mnuFromAC_Click(Index As Integer)
btnTransColor.BackColor = MainForm.ActiveColor(Index).BackColor
TransColor = MainForm.GetACol(Index)
End Sub

Private Sub Option_Click(Index As Integer)
TransR.Enabled = (Me.Option(1).Value)
btnTransColor.Enabled = (Me.Option(9).Value Or Me.Option(12).Value)
btnSuperView.Enabled = Me.Option(10).Value
btnChooseTR.Enabled = (Me.Option(10).Value)
btnEdit.Enabled = (Me.Option(10).Value)
btnEditRGB.Enabled = (Me.Option(11).Value)
End Sub

Private Sub Option_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.Option(Index).Value = True
End Sub
