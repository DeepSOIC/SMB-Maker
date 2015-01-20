VERSION 5.00
Begin VB.Form frmDynSettings 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Настройки прокрутки"
   ClientHeight    =   5984
   ClientLeft      =   44
   ClientTop       =   341
   ClientWidth     =   7557
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.07
      Charset         =   204
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5984
   ScaleWidth      =   7557
   Begin SMBMaker.ctlResourcizator Resr1 
      Left            =   165
      Top             =   5390
      _ExtentX        =   1097
      _ExtentY        =   671
      ResID           =   9920
   End
   Begin SMBMaker.dbFrame dbFrame3 
      Height          =   2670
      Left            =   4740
      TabIndex        =   26
      Top             =   2640
      Width           =   2730
      _ExtentX        =   4816
      _ExtentY        =   4714
      Caption         =   "Навигационный режим"
      EAC             =   0   'False
      Begin VB.OptionButton optNaviAbs 
         BackColor       =   &H008080FF&
         Caption         =   "Абсолютный"
         Height          =   345
         Left            =   510
         Style           =   1  'Graphical
         TabIndex        =   29
         ToolTipText     =   "Это планшет или сенсорный экран. Если он есть - лучще выбрать этот вариант (с мышью совместим)."
         Top             =   2145
         Value           =   -1  'True
         Width           =   1560
      End
      Begin VB.OptionButton optNaviRel 
         BackColor       =   &H008080FF&
         Caption         =   "Неабсолютный"
         Height          =   345
         Left            =   495
         Style           =   1  'Graphical
         TabIndex        =   28
         ToolTipText     =   "Это обычная мышь. Или тачпад. Если у вас нет планшета, лучше выбрать этот вариант."
         Top             =   1785
         Width           =   1560
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmDynSettings.frx":0000
         Height          =   1470
         Left            =   105
         TabIndex        =   27
         Top             =   225
         Width           =   2460
         WordWrap        =   -1  'True
      End
   End
   Begin SMBMaker.dbFrame dbFrame2 
      Height          =   2640
      Left            =   90
      TabIndex        =   15
      Top             =   2640
      Width           =   4575
      _ExtentX        =   8067
      _ExtentY        =   4653
      Caption         =   "Автопрокрутка"
      EAC             =   0   'False
      Begin VB.OptionButton optPtrPen 
         BackColor       =   &H0080FFFF&
         Caption         =   "планшета"
         Height          =   345
         Left            =   2565
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   240
         Width           =   1560
      End
      Begin VB.OptionButton optPtrNormal 
         BackColor       =   &H0080FFFF&
         Caption         =   "мыши"
         Height          =   345
         Left            =   990
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   225
         Value           =   -1  'True
         Width           =   1560
      End
      Begin SMBMaker.ctlNumBox nmbLef 
         Height          =   315
         Left            =   180
         TabIndex        =   16
         Top             =   1500
         Width           =   900
         _ExtentX        =   1585
         _ExtentY        =   549
         Min             =   -10000
         Max             =   10000
         HorzMode        =   0   'False
         EditName        =   $"frmDynSettings.frx":00C4
         SliderVisible   =   0   'False
      End
      Begin SMBMaker.ctlNumBox nmbTop 
         Height          =   315
         Left            =   1815
         TabIndex        =   17
         Top             =   900
         Width           =   900
         _ExtentX        =   1585
         _ExtentY        =   549
         Min             =   -10000
         Max             =   10000
         HorzMode        =   0   'False
         EditName        =   $"frmDynSettings.frx":015F
         SliderVisible   =   0   'False
      End
      Begin SMBMaker.ctlNumBox nmbRig 
         Height          =   315
         Left            =   3450
         TabIndex        =   18
         Top             =   1515
         Width           =   900
         _ExtentX        =   1585
         _ExtentY        =   549
         Min             =   -10000
         Max             =   10000
         HorzMode        =   0   'False
         EditName        =   $"frmDynSettings.frx":01FC
         SliderVisible   =   0   'False
      End
      Begin SMBMaker.ctlNumBox nmbBot 
         Height          =   315
         Left            =   1815
         TabIndex        =   19
         Top             =   2130
         Width           =   900
         _ExtentX        =   1585
         _ExtentY        =   549
         Min             =   -10000
         Max             =   10000
         HorzMode        =   0   'False
         EditName        =   $"frmDynSettings.frx":0298
         SliderVisible   =   0   'False
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ширины областей автопрокрутки:"
         ForeColor       =   &H00000000&
         Height          =   187
         Left            =   154
         TabIndex        =   25
         Top             =   627
         Width           =   2508
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "зона рисования"
         ForeColor       =   &H00808080&
         Height          =   405
         Left            =   1155
         TabIndex        =   24
         Top             =   1275
         Width           =   1500
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "зона автопрокрутки"
         ForeColor       =   &H00008000&
         Height          =   405
         Left            =   255
         TabIndex        =   23
         ToolTipText     =   "Когда вы рисуете в зоне автопрокрутки, автоматически происходит смещение поля зрения (окна)."
         Top             =   960
         Width           =   1500
      End
      Begin VB.Shape Shape2 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   900
         Left            =   1095
         Top             =   1230
         Width           =   2355
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Для:"
         Height          =   405
         Left            =   135
         TabIndex        =   20
         Top             =   315
         Width           =   1125
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H0080FF80&
         FillStyle       =   0  'Solid
         Height          =   1575
         Left            =   165
         Top             =   885
         Width           =   4200
      End
   End
   Begin SMBMaker.dbFrame dbFrame1 
      Height          =   2519
      Left            =   88
      TabIndex        =   2
      Top             =   66
      Width           =   7403
      _ExtentX        =   13066
      _ExtentY        =   4450
      Caption         =   "Плавность"
      EAC             =   0   'False
      Begin VB.PictureBox pctTest 
         BackColor       =   &H0080FF80&
         ClipControls    =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.49
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1925
         Left            =   3201
         ScaleHeight     =   171
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   345
         TabIndex        =   12
         Top             =   330
         Width           =   3839
         Begin VB.PictureBox MP 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.49
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   352
            Left            =   1215
            Picture         =   "frmDynSettings.frx":0334
            ScaleHeight     =   352
            ScaleWidth      =   352
            TabIndex        =   13
            Top             =   825
            Width           =   352
         End
      End
      Begin VB.CheckBox chkEnabled 
         BackColor       =   &H00FFC0FF&
         Caption         =   "Включить плавность"
         Height          =   390
         Left            =   198
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   297
         Width           =   2715
      End
      Begin VB.TextBox txtJestkost 
         Height          =   300
         Left            =   1298
         TabIndex        =   5
         Top             =   781
         Width           =   1335
      End
      Begin VB.TextBox txtEnL 
         Height          =   300
         Left            =   1298
         TabIndex        =   4
         Top             =   1089
         Width           =   1335
      End
      Begin VB.Timer Timer1 
         Interval        =   1
         Left            =   2574
         Top             =   1342
      End
      Begin VB.HScrollBar scrTR 
         Height          =   225
         LargeChange     =   10
         Left            =   198
         Max             =   100
         Min             =   4
         TabIndex        =   3
         Top             =   1749
         Value           =   4
         Width           =   2760
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Область проверки"
         Height          =   187
         Left            =   3212
         TabIndex        =   14
         Top             =   110
         Width           =   1320
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Жёсткость"
         Height          =   195
         Left            =   270
         TabIndex        =   11
         Top             =   840
         Width           =   1020
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Замедление"
         Height          =   195
         Left            =   270
         TabIndex        =   10
         Top             =   1140
         Width           =   1020
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Разрешение таймеров"
         Height          =   220
         Left            =   220
         TabIndex        =   9
         ToolTipText     =   $"frmDynSettings.frx":0F76
         Top             =   1540
         Width           =   2750
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "высокое"
         Height          =   198
         Left            =   198
         TabIndex        =   8
         Top             =   1980
         Width           =   858
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "низкое"
         Height          =   220
         Left            =   2068
         TabIndex        =   7
         Top             =   1958
         Width           =   869
      End
   End
   Begin SMBMaker.dbButton OkButton 
      Default         =   -1  'True
      Height          =   540
      Left            =   315
      TabIndex        =   0
      Top             =   5340
      Width           =   3300
      _ExtentX        =   5812
      _ExtentY        =   955
      MouseIcon       =   "frmDynSettings.frx":1012
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.064
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Others          =   $"frmDynSettings.frx":102E
      OthersPresent   =   -1  'True
   End
   Begin SMBMaker.dbButton CancelButton 
      Cancel          =   -1  'True
      Height          =   540
      Left            =   3960
      TabIndex        =   1
      Top             =   5340
      Width           =   3300
      _ExtentX        =   5812
      _ExtentY        =   955
      MouseIcon       =   "frmDynSettings.frx":107A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.064
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Others          =   $"frmDynSettings.frx":1096
      OthersPresent   =   -1  'True
   End
End
Attribute VB_Name = "frmDynSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim x0 As Long, y0 As Long
'Dim Jestkost As Single, EnL As Single
Dim UnlNow As Boolean
Dim MoveTimerRes As Long
Dim pScrollSettings As typScrollSettings
Dim bLock As Boolean
Dim bPen As Boolean 'follower of optPtrPen

Private Sub CancelButton_Click()
UnlNow = True
Me.Tag = "C"
Me.Hide
End Sub

Private Sub dbButton1_Click()
dbMsgBox " 1. Use input fields to set the rigidity and the slowdown of the scrolling." + vbCrLf + _
         " 2. Scrollbar sets timer resolution. Left is great resolution, but slow computer may just hang. Right is a poor resolution. Scrolling is not so smooth. Short stops in movement mean that the resolution is to high.", vbInformation
End Sub

Private Sub chkEnabled_Click()
pScrollSettings.DS_Enabled = (chkEnabled.Value = vbChecked)
End Sub

Public Sub Form_Load()
LoadCaptions
pScrollSettings.DS_Jestkost = 0.67
pScrollSettings.DS_EnL = 0.91
MoveTimerRes = 8
UpdateTexts
UnlNow = False
End Sub

Public Sub LoadCaptions()
Resr1.LoadCaptions
Dim kArr() As String, nKeys As Long
Dim KeyStr As String
nKeys = Keyb.ListKeys(cmdNaviMode, kArr)
If nKeys = 0 Then
    KeyStr = GRSF(2427)
Else
    KeyStr = kArr(0)
End If
Label11.Caption = Replace(Label11.Caption, "%key", KeyStr)
'Me.Caption = GRSF(2409)
'chkEnabled.Caption = GRSF(2415)
'Label1.Caption = GRSF(2410)
'Label2.Caption = GRSF(2411)
'Label3.Caption = GRSF(2412)
'Label4.Caption = GRSF(2416)
'Label5.Caption = GRSF(2417)
End Sub

Public Sub MoveMP(Optional ByVal x As Long = &H7FFFFFFF, Optional ByVal y As Long = &H7FFFFFFF, Optional ByVal Flags As Long, Optional ByVal ForceUpdate As Boolean = False)
On Error GoTo eh
Static mx As Long, my As Long
Static tx As Single, ty As Single
Static vx As Single, vy As Single
Static LastUpdateTime As Long
Dim t As Long
Dim ax As Single, ay As Single
Dim i As Long
Dim PrC As Long
Static vk As Single
Static PowEnL As Single
Static JxVKd5 As Single

If x <> &H7FFFFFFF Then
    If Flags And &H1 Then
        mx = mx + (x And &HFFFF)
    Else
        mx = x
    End If
End If
If y <> &H7FFFFFFF Then
    If Flags And &H2 Then
        my = my + (y And &HFFFF)
    Else
        my = y
    End If
End If
If Flags And &H4 Or Not pScrollSettings.DS_Enabled Then
    tx = mx
    ty = my
End If

If CBool(Flags And &H8) Or vk = 0 Then
    vk = MoveTimerRes / 16
    PowEnL = pScrollSettings.DS_EnL ^ vk
    JxVKd5 = pScrollSettings.DS_Jestkost * vk / 5
End If


If ForceUpdate Then
    Do
        t = mGetTickCount
        If LastUpdateTime = 0 Then LastUpdateTime = t
    Loop Until (t - LastUpdateTime >= MoveTimerRes)
    ax = -((tx - mx) * JxVKd5)
    ay = -((ty - my) * JxVKd5)
    vx = vx + ax
    vy = vy + ay
    vx = vx * PowEnL
    vy = vy * PowEnL
    
    tx = tx + vx * vk
    ty = ty + vy * vk
    LastUpdateTime = t 'LastUpdateTime + MoveTimerRes

Else
    t = mGetTickCount
    If LastUpdateTime = 0 Then LastUpdateTime = t
    If t - LastUpdateTime < MoveTimerRes Then Exit Sub
    i = 0
    Do
        ax = -((tx - mx) * JxVKd5)
        ay = -((ty - my) * JxVKd5)
        vx = vx + ax
        vy = vy + ay
        vx = vx * PowEnL
        vy = vy * PowEnL
        
        tx = tx + vx * vk
        ty = ty + vy * vk
        t = mGetTickCount
        LastUpdateTime = LastUpdateTime + MoveTimerRes
        i = i + 1
        If i >= 4000 Then
            dbMsgBox "Timer resolution is lowered!!!", vbInformation
            MoveTimerRes = MoveTimerRes * 2
            UpdateTexts
            MoveMP , , &H8, True
            Exit Sub
        End If
    Loop Until t - LastUpdateTime < MoveTimerRes
End If
    
MP.Move Round(tx), Round(ty)
Exit Sub
eh:
tx = 0
ty = 0
mx = 0
my = 0
vx = 0
vy = 0
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

Private Sub MP_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
UnlNow = True
If Button = 1 Then
    x0 = x
    y0 = y
End If
End Sub

Private Sub MP_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim xn As Long, yn As Long
If Button = 1 Then
    xn = MP.Left + x - x0
    yn = MP.Top + y - y0
    MoveMP xn, yn, , False
End If
End Sub

Private Sub nmbBot_Change()
RetreiveGapNumbers
End Sub

Private Sub nmbLef_Change()
RetreiveGapNumbers
End Sub

Private Sub nmbRig_Change()
RetreiveGapNumbers
End Sub

Private Sub nmbTop_Change()
RetreiveGapNumbers
End Sub

Private Sub OkButton_Click()
UnlNow = True
Me.Tag = ""
Me.Hide
End Sub

Private Sub optPtrNormal_Click()
bPen = optPtrPen.Value
FillGapNumbers
End Sub

Private Sub optPtrPen_Click()
bPen = optPtrPen.Value
FillGapNumbers
End Sub

Private Sub pctTest_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
pctTest_MouseMove Button, Shift, x, y
End Sub

Private Sub pctTest_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
'If Button = 1 Then
    pctTest.SetFocus
    MoveMP x - MP.Width \ 2, y - MP.Height \ 2, , False
'End If
End Sub

Private Sub scrTR_Change()
scrTR_Scroll
End Sub

Private Sub scrTR_Scroll()
MoveTimerRes = scrTR.Value
MoveMP , , &H8, True
End Sub

Private Sub Timer1_Timer()
'MoveMP
Timer1.Enabled = False
TakeMessagesControl
End Sub

Private Sub txtEnL_LostFocus()
With txtEnL
    If Val(.Text) >= 1 Or Val(.Text) <= 0 Then
        dbMsgBox "Slowdown must be between 0 and 1", vbInformation
    Else
        pScrollSettings.DS_EnL = Val(.Text)
        .Text = Trim(Str(pScrollSettings.DS_EnL))
    End If
End With
MoveMP , , &H8
End Sub

Private Sub txtJestkost_LostFocus()
With txtJestkost
    If Val(.Text) <= 0 Then
        dbMsgBox "Rigidity must be greater than 0", vbInformation
    Else
        pScrollSettings.DS_Jestkost = Val(.Text)
        .Text = Trim(Str(pScrollSettings.DS_Jestkost))
    End If
End With
MoveMP , , &H8
End Sub

Private Sub UpdateTexts()
txtJestkost.Text = Trim(Str(pScrollSettings.DS_Jestkost))
txtEnL.Text = Trim(Str(pScrollSettings.DS_EnL))
scrTR.Value = MoveTimerRes
End Sub

Friend Sub SetProps(ScrollSettings As typScrollSettings, ByVal pMoveTimerRes As Long)
bLock = True
pScrollSettings = ScrollSettings
chkEnabled.Value = IIf(pScrollSettings.DS_Enabled, vbChecked, vbUnchecked)
'pScrollSettings.DS_Jestkost = pJestkost
'pScrollSettings.DS_EnL = pEnL
MoveTimerRes = pMoveTimerRes
If MoveTimerRes > scrTR.Max Then MoveTimerRes = scrTR.Max
If MoveTimerRes < scrTR.Min Then MoveTimerRes = scrTR.Min
UpdateTexts
MoveMP , , &H8
FillGapNumbers
optNaviRel.Value = Not pScrollSettings.NaviAbsoluteMode
optNaviAbs.Value = pScrollSettings.NaviAbsoluteMode
End Sub

Friend Sub GetProps(ScrollSettings As typScrollSettings, ByRef pMoveTimerRes As Long)
'pEn = CBool(chkEnabled.Value)
'pJestkost = pScrollSettings.DS_Jestkost
'pEnL = pScrollSettings.DS_EnL
pScrollSettings.NaviAbsoluteMode = optNaviAbs.Value
ScrollSettings = pScrollSettings
pMoveTimerRes = MoveTimerRes
End Sub

Public Sub TakeMessagesControl()
Do
    MoveMP
    DoEvents
    If UnlNow Then Exit Sub
Loop
End Sub

Private Sub FillGapNumbers()
Dim ASS As typAutoscrollSgs
If optPtrPen.Value Then
  ASS = pScrollSettings.ASS_pen
Else
  ASS = pScrollSettings.ASS
End If
bLock = True
On Error Resume Next
nmbLef.Value = ASS.GapLef
nmbTop.Value = ASS.GapTop
nmbRig.Value = ASS.GapRig
nmbBot.Value = ASS.GapBot
bLock = False
End Sub

Private Sub RetreiveGapNumbers()
Dim ASS As typAutoscrollSgs
If bLock Then Exit Sub
ASS.GapLef = nmbLef.Value
ASS.GapTop = nmbTop.Value
ASS.GapRig = nmbRig.Value
ASS.GapBot = nmbBot.Value
If bPen Then
  pScrollSettings.ASS_pen = ASS
Else
  pScrollSettings.ASS = ASS
End If
End Sub
