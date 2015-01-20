VERSION 5.00
Begin VB.Form Pereliv 
   BackColor       =   &H00E3DFE0&
   Caption         =   "Fade options"
   ClientHeight    =   2445
   ClientLeft      =   1410
   ClientTop       =   1200
   ClientWidth     =   5880
   Icon            =   "Pereliv.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   163
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   392
   StartUpPosition =   1  'CenterOwner
   Begin VB.OptionButton Opts 
      BackColor       =   &H0080FFFF&
      Caption         =   "Program"
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
      Index           =   2
      Left            =   3750
      Style           =   1  'Graphical
      TabIndex        =   11
      Tag             =   "-1"
      ToolTipText     =   "Algorithm - linear"
      Top             =   1710
      Visible         =   0   'False
      Width           =   1215
   End
   Begin SMBMaker.ctlNumBox nmbOffset 
      Height          =   675
      Left            =   705
      TabIndex        =   10
      Top             =   1005
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1191
      Max             =   1
      HorzMode        =   0   'False
      EditName        =   "$2423"
      NLn             =   0
   End
   Begin SMBMaker.ctlNumBox nmbPower 
      Height          =   630
      Left            =   2850
      TabIndex        =   9
      Top             =   1005
      Width           =   1230
      _ExtentX        =   2170
      _ExtentY        =   1111
      Value           =   0,5
      Max             =   1
      HorzMode        =   0   'False
      EditName        =   "$2104"
      NLn             =   0
   End
   Begin SMBMaker.ctlNumBox nmbCount 
      Height          =   645
      Left            =   1710
      TabIndex        =   8
      Top             =   1005
      Width           =   1110
      _ExtentX        =   1958
      _ExtentY        =   1138
      Value           =   1
      Max             =   2000
      HorzMode        =   0   'False
      EditName        =   "$2101"
      NLn             =   2
   End
   Begin VB.OptionButton Opts 
      BackColor       =   &H0080FFFF&
      Caption         =   "Linear"
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
      Left            =   2430
      Style           =   1  'Graphical
      TabIndex        =   1
      Tag             =   "1"
      ToolTipText     =   "Algorithm - linear"
      Top             =   1680
      Value           =   -1  'True
      Width           =   1215
   End
   Begin VB.OptionButton Opts 
      BackColor       =   &H0080FFFF&
      Caption         =   "Sine"
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
      Left            =   750
      Style           =   1  'Graphical
      TabIndex        =   0
      Tag             =   "0"
      ToolTipText     =   "Algorithm - sine."
      Top             =   1665
      Width           =   1590
   End
   Begin VB.TextBox Counter 
      Height          =   375
      Left            =   210
      TabIndex        =   6
      Text            =   "100"
      Top             =   360
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E3DFE0&
      Height          =   660
      Left            =   105
      ScaleHeight     =   40
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   289
      TabIndex        =   5
      ToolTipText     =   "Preview."
      Top             =   60
      Width           =   4395
   End
   Begin SMBMaker.dbButton OkButton 
      Cancel          =   -1  'True
      Default         =   -1  'True
      Height          =   285
      Left            =   1050
      TabIndex        =   2
      ToolTipText     =   "Закрыть окно"
      Top             =   2160
      Width           =   2070
      _ExtentX        =   0
      _ExtentY        =   0
      MouseIcon       =   "Pereliv.frx":18BA
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Others          =   $"Pereliv.frx":18D6
      OthersPresent   =   -1  'True
   End
   Begin SMBMaker.dbButton Command 
      Height          =   435
      Left            =   1545
      TabIndex        =   7
      Top             =   300
      Visible         =   0   'False
      Width           =   765
      _ExtentX        =   0
      _ExtentY        =   0
      MouseIcon       =   "Pereliv.frx":1922
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Others          =   $"Pereliv.frx":193E
      OthersPresent   =   -1  'True
   End
   Begin SMBMaker.dbButton bColor 
      Height          =   735
      Index           =   1
      Left            =   4065
      TabIndex        =   4
      Tag             =   "True"
      ToolTipText     =   "Color #2 (right-click me to toggle mode)"
      Top             =   960
      Width           =   495
      _ExtentX        =   0
      _ExtentY        =   0
      MouseIcon       =   "Pereliv.frx":198A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Others          =   $"Pereliv.frx":19A6
      OthersPresent   =   -1  'True
   End
   Begin SMBMaker.dbButton bColor 
      Height          =   675
      Index           =   0
      Left            =   165
      TabIndex        =   3
      Tag             =   "True"
      ToolTipText     =   "Color #1 (right-click me to toggle mode)"
      Top             =   960
      Width           =   495
      _ExtentX        =   0
      _ExtentY        =   0
      MouseIcon       =   "Pereliv.frx":19EA
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Others          =   $"Pereliv.frx":1A06
      OthersPresent   =   -1  'True
   End
End
Attribute VB_Name = "Pereliv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Colors() As Long
Dim CurFDsc As FadeDesc

Private Sub CancelButton_Click()
Me.Tag = "C"
Me.Hide
End Sub

Private Sub Form_Load()
dbLoadCaptions
End Sub

Private Sub Form_Paint()
Me.PaintPicture LoadResPicture(RES_BackPicture, vbResBitmap), 0, 0, Me.ScaleWidth, Me.ScaleHeight
End Sub

'Private Sub Color1_Click()
'With Color
'.Color = Color1.BackColor
'.ShowColor
'Color1.BackColor = .Color
'End With
'UpdateFade
'End Sub

Public Property Get GetColor(Index As Long) As Long
GetColor = Colors(Index)
End Property

'Private Sub Color2_Click()
'With Color
'.Color = Color2.BackColor
'.ShowColor
'Color2.BackColor = .Color
'End With
'UpdateFade
'End Sub

Sub UpdateFade(Optional NoRefr As Boolean = False)
Dim i As Long, j As Double, h As Integer, tmp As Long
Dim Wdt As Long
Dim cnt As Single, Pow As Single, Mode As Integer
Dim rgb1 As RGBTriCurr, rgb2 As RGBTriCurr
Dim Ofc As Single
Dim dj As Double
If Not (MainForm.Tag = "") Then Exit Sub

On Error GoTo eh

'Pow = nmbPower.Value
'Cnt = nmbCount.Value
'Mode = FadeMode
'Ofc = nmbOffset.Value
UpdateFDsc

GetRGBQuadFloatEx bColor(0).BackColor, rgb1
GetRGBQuadFloatEx bColor(1).BackColor, rgb2

Wdt = Picture1.ScaleWidth
If Wdt > 0 Then
    dj = 1 / Wdt
    If Not Me.Visible Then Exit Sub
    ReDim Colors(0 To Wdt - 1)
    For i = 0 To Wdt - 1
        j = i / CDbl(Wdt - 1)
        j = CountJEx(j, CurFDsc, dj)
        tmp = RGB(Int(j * (rgb2.rgbRed - rgb1.rgbRed) + rgb1.rgbRed), _
                  Int(j * (rgb2.rgbGreen - rgb1.rgbGreen) + rgb1.rgbGreen), _
                  Int(j * (rgb2.rgbBlue - rgb1.rgbBlue) + rgb1.rgbBlue))
        If Not NoRefr Then
            Picture1.Line (i, 0)-Step(0, Picture1.ScaleHeight), tmp
        End If
        Colors(i) = tmp
    Next i
End If
Picture1.Refresh
Exit Sub
eh:
vtBeep
End Sub

Private Function GetAtr(ByVal lngColor As Long, ByVal n As Integer) As Long
Dim t As Long
Select Case n
Case 1
t = 1
Case 2
t = 256
Case 3
t = 65536
End Select
GetAtr = ((lngColor) And (CLng(255) * t)) / t
End Function

Private Sub Form_Unoad(Cancel As Integer)
Cancel = True
CancelButton_Click
End Sub

Private Sub bColor_Click(Index As Integer)
    With CDl
        On Error GoTo eh
            .Flags = 0
            .ColorFlags = cdlCCRGBInit
            .Color = bColor(Index).BackColor
            .ShowColor
        On Error GoTo 0
        bColor(Index).BackColor = .Color
        UpdateFade
    End With
eh:
End Sub

Sub bColor_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
    bColor(Index).Tag = CStr(Not (CBool(bColor(Index).Tag)))
    If CBool(bColor(Index).Tag) Then
        bColor(Index).Caption = "*"
    Else
        bColor(Index).Caption = ""
    End If
    'UpdateFade
End If
End Sub

Friend Sub UpdateCaptions()
Dim i As Integer
bColor(0).Caption = IIf(CBool(bColor(0).Tag), "*", "")
bColor(1).Caption = IIf(CBool(bColor(1).Tag), "*", "")
End Sub

Private Sub CPer_Change()
UpdateFade
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If UnloadMode = VBRUN.QueryUnloadConstants.vbFormControlMenu Then
    Cancel = 1
    Me.Tag = "C"
    Me.Hide
End If
End Sub

Private Sub Form_Resize()
Dim w1 As Long
Dim p1 As Long, p2 As Long
On Error Resume Next
Picture1.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight - bColor(0).Height - Opts(0).Height - OkButton.Height
bColor(0).Move 0, Picture1.Height
bColor(1).Move Me.ScaleWidth - bColor(0).Width, bColor(0).Top, bColor(0).Width, bColor(0).Height
w1 = Me.ScaleWidth - 2 * bColor(0).Width
p1 = w1 \ 3 + bColor(0).Width
p2 = (w1 * 2) \ 3 + bColor(0).Width
nmbOffset.Move bColor(0).Width, bColor(0).Top, p1 - bColor(0).Width, bColor(0).Height
nmbCount.Move p1, bColor(0).Top, p2 - p1, bColor(0).Height
nmbPower.Move p2, bColor(0).Top, Me.ScaleWidth - bColor(1).Width - p2, bColor(0).Height
Opts(0).Move 0, bColor(0).Top + bColor(0).Height, Me.ScaleWidth \ 2
Opts(1).Move Opts(0).Width, Opts(0).Top, Me.ScaleWidth - Opts(0).Width, Opts(0).Height
OkButton.Move (Me.ScaleWidth - OkButton.Width) \ 2, Me.ScaleHeight - OkButton.Height
UpdateFade
End Sub

Private Sub nmbCount_Change()
UpdateFade
End Sub

Private Sub nmbCount_InputChange()
nmbCount_Change
End Sub

Private Sub nmbOffset_Change()
UpdateFade
End Sub

Private Sub nmbOffset_InputChange()
nmbOffset_Change
End Sub

Private Sub nmbPower_Change()
UpdateFade
End Sub

Private Sub nmbPower_InputChange()
nmbPower_Change
End Sub

Private Sub OkButton_Click()
    Me.Tag = ""
    Me.Hide
End Sub

Private Sub Opts_Click(Index As Integer)

UpdateFade
End Sub

Public Property Get FadeMode() As Integer
Dim i As Integer
For i = 0 To Opts.UBound
    If Opts(i).Value Then
        'FadeMode = CLng(Opts(i).Tag)
        Exit For
    End If
Next i
If i = Opts.UBound + 1 Then
    FadeMode = 0
Else
    FadeMode = CLng(Opts(i).Tag)
End If
End Property

Friend Sub ExtractFadeDesc(ByRef FadeDsc As FadeDesc)
FadeDsc = CurFDsc
End Sub

Friend Sub UpdateFDsc()
On Error Resume Next
With CurFDsc
    .Mode = FadeMode
    .FCount = nmbCount.Value
    .Offset = nmbOffset.Value
    .Power = nmbPower.Value
End With
End Sub

Sub dbLoadCaptions()
Me.Caption = GRSF(2145)
Opts(1).Caption = GRSF(2097)
Opts(1).ToolTipText = GRSF(2098)
Opts(0).Caption = GRSF(2099)
Opts(0).ToolTipText = GRSF(2100)
'CPer.ToolTipText = grsf(2101)
OkButton.Caption = GRSF(2102)
OkButton.ToolTipText = GRSF(2103)
'Stepen.ToolTipText = grsf(2104)
Command.Caption = GRSF(2105)
bColor(1).ToolTipText = GRSF(2106)
bColor(0).ToolTipText = GRSF(2107)
Picture1.ToolTipText = GRSF(2108)
End Sub
