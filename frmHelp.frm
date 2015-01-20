VERSION 5.00
Begin VB.Form frmHelp 
   Caption         =   "SMB Maker floating help"
   ClientHeight    =   3315
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5385
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   221
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   359
   Begin SMBMaker.dbFrame fTopics 
      Height          =   435
      Left            =   60
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   60
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   767
      EAC             =   0   'False
      Begin VB.CheckBox chkLockTopic 
         BackColor       =   &H008080FF&
         Height          =   225
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   0
         Width           =   225
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   4515
      End
   End
   Begin SMBMaker.ctlTaggedText fHelp 
      Height          =   1785
      Left            =   510
      TabIndex        =   2
      Top             =   1080
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   3149
   End
   Begin SMBMaker.dbButton btnTopic 
      Height          =   315
      Index           =   0
      Left            =   60
      TabIndex        =   3
      Top             =   495
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   556
      MouseIcon       =   "frmHelp.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Others          =   $"frmHelp.frx":001C
      OthersPresent   =   -1  'True
   End
End
Attribute VB_Name = "frmHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ClHeight As Long
Dim VScrollEnabled As Boolean
Dim HText As String
Public bLockTopic As Boolean
Option Explicit

Public Sub btnTopic_Click(Index As Integer)
On Error GoTo eh
fHelp.SetResID CInt(btnTopic(Index).Tag)
fHelp.Scroll -2000
Exit Sub
eh:
fHelp.SetResID 2396
End Sub

Private Sub chkLockTopic_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
    bLockTopic = Not bLockTopic
    chkLockTopic.Value = IIf(bLockTopic, vbChecked, vbUnchecked)
End If
End Sub

Private Sub chkLockTopic_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
    chkLockTopic.Value = IIf(bLockTopic, vbChecked, vbUnchecked)
End If
End Sub

'Private Sub fHelp_Resize()
'On Error Resume Next
'VScrollEnabled = (fHelp.Height > ClHeight)
'VScroll.Enabled = VScrollEnabled
'If VScrollEnabled Then
'    VScroll.Max = fHelp.Height - ClHeight
'    VScroll.LargeChange = 0.8 * ClHeight + 1
'End If
''fHelp.Refresh
'ApplyScroll
'End Sub
'
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Dim ShiftKeyCode As Long
ShiftKeyCode = GetShiftKeyCode(KeyCode, Shift)
Select Case ShiftKeyCode
    Case 112 'f1
        KeyCode = 0
    Case 27 'Esc
        HideHelpWindow
End Select
End Sub

Private Sub Form_Load()
'HText = "Test text<Center=1>Centered text<center=0>Left-aligned text 17-sized" + vbCrLf
'HText = HText + "jsahl dfjkhl wakjeyhui oryqoiuye otuiyqowu ieyoriuyqwo iueyt oiuqw yoeiur7 2369 578169785619 876 987459 6ewuyr oiuqwyeor uiqyowe iuy fajsdhf jasdfh lakjshdlkjfhlkj qhwlekuyroui fysdkjh flkae jyouiryfl djhcvkjzxcnv,mnzx. cmvn.hjsdhf lkjhqweiuyro iquwydhf lkjhasl "
'HText = HText + "jsahl dfjkhl wakjeyhui oryqoiuye otuiyqowu ieyoriuyqwo iueyt oiuqw yoeiur7 2369 578169785619 876 987459 6ewuyr oiuqwyeor uiqyowe iuy fajsdhf jasdfh lakjshdlkjfhlkj qhwlekuyroui fysdkjh flkae jyouiryfl djhcvkjzxcnv,mnzx. cmvn.hjsdhf lkjhqweiuyro iquwydhf lkjhasl "
'HText = HText + "jsahl dfjkhl wakjeyhui oryqoiuye otuiyqowu ieyoriuyqwo iueyt oiuqw yoeiur7 2369 578169785619 876 987459 6ewuyr oiuqwyeor uiqyowe iuy fajsdhf jasdfh lakjshdlkjfhlkj qhwlekuyroui fysdkjh flkae jyouiryfl djhcvkjzxcnv,mnzx. cmvn.hjsdhf lkjhqweiuyro iquwydhf lkjhasl "
'HText = HText + "jsahl dfjkhl wakjeyhui oryqoiuye otuiyqowu ieyoriuyqwo iueyt oiuqw yoeiur7 2369 578169785619 876 987459 6ewuyr oiuqwyeor uiqyowe iuy fajsdhf jasdfh lakjshdlkjfhlkj qhwlekuyroui fysdkjh flkae jyouiryfl djhcvkjzxcnv,mnzx. cmvn.hjsdhf lkjhqweiuyro iquwydhf lkjhasl "
LoadHelp 2396
'LoadWindowPos frmHelp
End Sub

Public Sub LoadHelp(ByVal ResID As Integer)
Dim sArr() As String
Dim i As Long
Dim nOfButtons As Long

If bLockTopic Then
    Exit Sub
End If

sArr = Split(GRSF(ResID), "|")
If (UBound(sArr)) Mod 2 <> 0 Or UBound(sArr) < 2 Then
    fHelp.SetResID 2395
End If
nOfButtons = (UBound(sArr)) \ 2
If nOfButtons > 0 Then
For i = btnTopic.UBound To nOfButtons Step -1
    Unload btnTopic(i)
Next i
For i = btnTopic.UBound + 1 To nOfButtons - 1
    Load btnTopic(i)
    btnTopic(i).ZOrder vbBringToFront
Next i
Label1.Caption = sArr(0)
For i = 0 To nOfButtons - 1
    With btnTopic(i)
        .Caption = sArr(2 * i + 1)
        .Tag = sArr(2 * i + 2)
        '.Visible = True
    End With
Next i
End If
MoveTopics
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
HelpWindowVisible = False
'SaveWindowPos frmHelp
End Sub

Private Sub Form_Resize()
On Error Resume Next
fTopics.Move 0, 0, Me.ScaleWidth
MoveTopics
fHelp.Move 0, btnTopic(0).Top + btnTopic(0).Height, Me.ScaleWidth, Me.ScaleHeight - (btnTopic(0).Top + btnTopic(0).Height)
'ClHeight = Me.ScaleHeight - btnTopic(0).Top - btnTopic(0).Height
'VScroll.Move Me.ScaleWidth - VScroll.Width, btnTopic(0).Top + btnTopic(0).Height, VScroll.Width, ClHeight
'fHelp_Resize
'ApplyScroll
End Sub

'Private Sub ApplyScroll()
'If VScrollEnabled Then
'    fHelp.Move 0, btnTopic(0).Top + btnTopic(0).Height - VScroll.Value, Me.ScaleWidth - VScroll.Width
'Else
'    fHelp.Move 0, btnTopic(0).Top + btnTopic(0).Height, Me.ScaleWidth - VScroll.Width
'End If
'End Sub
'
Private Sub MoveTopics()
Dim BUB As Long
Dim i As Long
Dim X As Long
Dim tppX As Long, tppY As Long
tppX = Screen.TwipsPerPixelX
tppY = Screen.TwipsPerPixelY
BUB = btnTopic.UBound
btnTopic(0).Move 0, fTopics.Height, 1 * ScaleWidth \ (BUB + 1)
btnTopic(0).Visible = True
For i = 1 To BUB
    X = btnTopic(i - 1).Left + btnTopic(i - 1).Width
    btnTopic(i).Move X, btnTopic(0).Top, (i + 1) * ScaleWidth \ (BUB + 1) - X + 1, btnTopic(0).Height
    btnTopic(i).Visible = True
Next i
End Sub

Private Sub fTopics_Resize()
Label1.Move 0, 0, fTopics.ScaleWidth * Screen.TwipsPerPixelX
End Sub

Friend Sub Scroll(ByVal Movement As Long)
fHelp.Scroll Movement
End Sub

'Private Sub VScroll_Change()
'ApplyScroll
'End Sub
'
'Private Sub VScroll_Scroll()
'ApplyScroll
'End Sub
