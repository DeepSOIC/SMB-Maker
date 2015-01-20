VERSION 5.00
Begin VB.Form frmSoft 
   BackColor       =   &H00E3DFE0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Фильтрация"
   ClientHeight    =   5368
   ClientLeft      =   2761
   ClientTop       =   3751
   ClientWidth     =   7623
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.07
      Charset         =   204
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   11006
   Icon            =   "frmSoft.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   488
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   693
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   WhatsThisHelp   =   -1  'True
   Begin SMBMaker.ctlResourcizator Resr1 
      Left            =   3927
      Top             =   4862
      _ExtentX        =   1097
      _ExtentY        =   671
      ResID           =   9904
   End
   Begin VB.Timer tmrUpdater 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   6435
      Top             =   3840
   End
   Begin VB.CheckBox chkTexMode 
      BackColor       =   &H00FF80FF&
      Caption         =   "Режим текстуры"
      Height          =   780
      Left            =   6135
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Включите, если создаете или редактируете текстуру."
      Top             =   330
      Width           =   1245
   End
   Begin SMBMaker.ctlNumBox nmbGain 
      Height          =   540
      Left            =   6225
      TabIndex        =   8
      Top             =   2670
      Width           =   1020
      _ExtentX        =   1808
      _ExtentY        =   955
      Min             =   -30
      Max             =   30
      NumType         =   5
      HorzMode        =   0   'False
      EditName        =   $"frmSoft.frx":0ABA
      NLn             =   0
   End
   Begin VB.OptionButton optMode 
      BackColor       =   &H0080FFFF&
      Caption         =   "Резкость"
      Height          =   405
      Index           =   2
      Left            =   5955
      Style           =   1  'Graphical
      TabIndex        =   6
      Tag             =   "2"
      ToolTipText     =   "Вместо смягчения усилить резкость рисунка."
      Top             =   1920
      Width           =   1590
   End
   Begin VB.OptionButton optMode 
      BackColor       =   &H0080FFFF&
      Caption         =   "Разность"
      Height          =   405
      Index           =   1
      Left            =   5955
      Style           =   1  'Graphical
      TabIndex        =   5
      Tag             =   "1"
      ToolTipText     =   "Отфильтровать рисунок, вывести отличие результата от исходного рисунка."
      Top             =   1515
      Width           =   1590
   End
   Begin VB.OptionButton optMode 
      BackColor       =   &H0080FFFF&
      Caption         =   "Фильтрация"
      Height          =   405
      Index           =   0
      Left            =   5955
      Style           =   1  'Graphical
      TabIndex        =   4
      Tag             =   "0"
      ToolTipText     =   "Отфильтровать рисунок и вывести результат."
      Top             =   1110
      Value           =   -1  'True
      Width           =   1590
   End
   Begin SMBMaker.dbFrame Shablon 
      Height          =   4587
      Left            =   66
      TabIndex        =   2
      Top             =   121
      Width           =   5808
      _ExtentX        =   10241
      _ExtentY        =   8087
      Caption         =   "Маска фильтрации"
      BackColor       =   14933984
      EAC             =   0   'False
      Begin SMBMaker.dbFrame frmSavedMasks 
         Height          =   4026
         Left            =   44
         TabIndex        =   11
         Top             =   275
         Width           =   1969
         _ExtentX        =   3475
         _ExtentY        =   7092
         Caption         =   "Сохранённые маски"
         EAC             =   0   'False
         Begin VB.ListBox List1 
            Height          =   2662
            Left            =   451
            TabIndex        =   12
            ToolTipText     =   "Двойной щелчок для загрузки маски."
            Top             =   275
            Width           =   1335
         End
         Begin SMBMaker.dbButton btnMoveDown 
            Height          =   506
            Left            =   33
            TabIndex        =   15
            ToolTipText     =   "Переместить выбранную маску вниз."
            Top             =   1694
            Width           =   385
            _ExtentX        =   691
            _ExtentY        =   894
            MouseIcon       =   "frmSoft.frx":0B9F
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Wingdings 3"
               Size            =   14.4
               Charset         =   2
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Others          =   $"frmSoft.frx":0BBB
            OthersPresent   =   -1  'True
         End
         Begin SMBMaker.dbButton Savebtn 
            Height          =   341
            Left            =   440
            TabIndex        =   14
            ToolTipText     =   "Добавить текущую маску в список."
            Top             =   2981
            Width           =   1331
            _ExtentX        =   2357
            _ExtentY        =   610
            MouseIcon       =   "frmSoft.frx":0C06
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.064
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Others          =   $"frmSoft.frx":0C22
            OthersPresent   =   -1  'True
         End
         Begin SMBMaker.dbButton btnMoveUp 
            Height          =   506
            Left            =   33
            TabIndex        =   13
            ToolTipText     =   "Переместить выбранную маску вверх."
            Top             =   1177
            Width           =   385
            _ExtentX        =   691
            _ExtentY        =   894
            MouseIcon       =   "frmSoft.frx":0C78
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Wingdings 3"
               Size            =   14.4
               Charset         =   2
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Others          =   $"frmSoft.frx":0C94
            OthersPresent   =   -1  'True
         End
      End
      Begin VB.PictureBox View 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.49
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   990
         Left            =   2255
         ScaleHeight     =   86
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   101
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   528
         Width           =   1155
      End
      Begin SMBMaker.dbButton btnOper 
         Height          =   374
         Left            =   3465
         TabIndex        =   17
         Top             =   0
         Width           =   1430
         _ExtentX        =   2520
         _ExtentY        =   650
         MouseIcon       =   "frmSoft.frx":0CDF
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.064
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Others          =   $"frmSoft.frx":0CFB
         OthersPresent   =   -1  'True
      End
      Begin SMBMaker.dbButton btnFileMenu 
         Height          =   374
         Left            =   2101
         TabIndex        =   16
         Top             =   0
         Width           =   1353
         _ExtentX        =   2377
         _ExtentY        =   650
         MouseIcon       =   "frmSoft.frx":0D53
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.064
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Others          =   $"frmSoft.frx":0D6F
         OthersPresent   =   -1  'True
      End
      Begin VB.Image Sizer 
         Height          =   99
         Left            =   3421
         MousePointer    =   8  'Size NW SE
         Picture         =   "frmSoft.frx":0DC8
         ToolTipText     =   "Используйте для изменения размеров маски."
         Top             =   1518
         Width           =   99
      End
   End
   Begin VB.Label lblPreview 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Просмотр:"
      Height          =   255
      Left            =   6045
      TabIndex        =   10
      Top             =   3345
      Width           =   1365
   End
   Begin VB.Image iPreview 
      Height          =   1680
      Left            =   6045
      Top             =   3570
      Width           =   1380
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Усиление (dB)"
      Height          =   240
      Left            =   5940
      TabIndex        =   7
      Top             =   2430
      Width           =   1590
   End
   Begin SMBMaker.dbButton CancelButton 
      Cancel          =   -1  'True
      Height          =   374
      Left            =   4719
      TabIndex        =   1
      Top             =   4873
      Width           =   1210
      _ExtentX        =   2134
      _ExtentY        =   671
      MouseIcon       =   "frmSoft.frx":0F06
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.064
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Others          =   $"frmSoft.frx":0F22
      OthersPresent   =   -1  'True
   End
   Begin SMBMaker.dbButton OKButton 
      Default         =   -1  'True
      Height          =   374
      Left            =   110
      TabIndex        =   0
      ToolTipText     =   "Начать выполнение эффекта. Может занять много времени. Если надо будет прервать, жмите Esc или Break."
      Top             =   4873
      Width           =   1210
      _ExtentX        =   2134
      _ExtentY        =   671
      MouseIcon       =   "frmSoft.frx":0F72
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.064
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Others          =   $"frmSoft.frx":0F8E
      OthersPresent   =   -1  'True
   End
   Begin VB.Menu mnuListPopup 
      Caption         =   "mnuListPopup"
      Visible         =   0   'False
      Begin VB.Menu mnuLPPRename 
         Caption         =   "Переименовать..."
      End
      Begin VB.Menu mnuLPPDelete 
         Caption         =   "Удалить..."
      End
      Begin VB.Menu mnuLPPOverwrite 
         Caption         =   "Перезаписать..."
      End
      Begin VB.Menu mnuLPPsep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLPPAddStd 
         Caption         =   "Добавить стандартные маски"
      End
   End
   Begin VB.Menu mnuFilePopup 
      Caption         =   "mnuFilePopup"
      Visible         =   0   'False
      Begin VB.Menu mnuFPPSave 
         Caption         =   "Сохранить как..."
      End
      Begin VB.Menu mnuFPPLoad 
         Caption         =   "Загрузить..."
      End
   End
   Begin VB.Menu mnuOperPopup 
      Caption         =   "mnuOperPopup"
      Visible         =   0   'False
      Begin VB.Menu mnuOPPClear 
         Caption         =   "Очистить"
      End
      Begin VB.Menu mnuOPPFlipH 
         Caption         =   "Отразить горизонтально"
      End
      Begin VB.Menu mnuOPPFilpV 
         Caption         =   "Отразить вертикально"
      End
      Begin VB.Menu mnuOPPRotate 
         Caption         =   "Повернуть на 90°"
      End
      Begin VB.Menu mnuOPPsep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOPPGenCircle 
         Caption         =   "Сделать кружок..."
      End
      Begin VB.Menu mnuOPPsep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOPPEdit 
         Caption         =   "Редактировать"
      End
      Begin VB.Menu mnuOPPUseSel 
         Caption         =   "Взять из выделения"
      End
   End
End
Attribute VB_Name = "frmSoft"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Dot() As Long
Dim mon As Byte
Const Max_Width = 25, Max_Height = 25
Option Explicit

Public Event Change()

Private Sub pChange()
tmrUpdater.Enabled = True
End Sub

Private Sub btnClear_Click()
End Sub

Private Sub btnDefPresets_Click()
End Sub

Private Sub btnEdit_Click()
End Sub

Private Sub btnFileMenu_Click()
PopupMenu mnuFilePopup
End Sub

Private Sub btnFromSel_Click()
End Sub

Public Sub TruncDot(ByRef Dot() As Long)
Dim x As Long, y As Long
Dim r As Long, b As Long
Dim tmpDot() As Long
b = Min(UBound(Dot, 2), Max_Height - 1)
r = Min(UBound(Dot, 1), Max_Width - 1)
ReDim tmpDot(0 To r, 0 To b)
For y = 0 To b
    For x = 0 To r
        tmpDot(x, y) = Dot(x, y)
    Next x
Next y
Dot = tmpDot
End Sub

Private Sub btnOper_Click()
PopupMenu mnuOperPopup
End Sub

Private Sub chkTexMode_Click()
pChange
End Sub

Private Sub Form_Activate()
'ShowHelp Me.HelpContextID
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Dim ShiftKeyCode As Long
ShiftKeyCode = GetShiftKeyCode(KeyCode, Shift)
If ShiftKeyCode = 112 Then 'F1
    ToggleHelpWindow
    KeyCode = 0
End If

End Sub

Private Sub Form_Paint()
Me.PaintPicture LoadResPicture(RES_BackPicture, vbResBitmap), 0, 0, Me.ScaleWidth, Me.ScaleHeight
End Sub

Private Sub btnDel_Click()
End Sub

Public Sub SaveMask(ByRef pDot() As Long, ByVal Index As Variant, ByVal strName As String)
Dim li As String, tmp As String, i As Long, j As Long
Dim w As Long, h As Long
Dim rgb1 As RGBQUAD
On Error GoTo eh
If IsNumeric(Index) Then
    li = VedNull(Index + 1, 2)
Else
    li = Index
End If
On Error GoTo 0
w = UBound(pDot, 1) + 1
h = UBound(pDot, 2) + 1
dbSaveSettingEx "Effects\SoftEx\" + li, "Height", h, True
dbSaveSettingEx "Effects\SoftEx\" + li, "Width", w, True
dbSaveSettingEx "Effects\SoftEx\" + li, "Name", strName, True
tmp = Space$(w * h * 3& * 2&)
For i = 0 To UBound(pDot, 2)
    For j = 0 To UBound(pDot, 1)
        Mid$(tmp, (i * w + j) * 6& + 1&, 6&) = _
              VedNullStr(Hex$(pDot(j, i) And &HFFFFFF), 6&)
    Next j
Next i
dbSaveSettingEx "Effects\SoftEx\" + li, "Data", tmp, True
Exit Sub
eh:
vtBeep
End Sub


Public Sub LoadMask(ByRef pDot() As Long, ByVal Index As Variant, ByRef strName As String)
Dim tmp As String, i As Long, j As Long, n As Long, m As Long
Dim li As String, Counter As Long
If IsNumeric(Index) Then
    li = VedNull(Index + 1, 2)
Else
    li = Index
End If

n = Val(dbGetSettingEx("Effects\SoftEx\" + li, "Height", vbByte, "3", True))
m = Val(dbGetSettingEx("Effects\SoftEx\" + li, "Width", vbByte, "3", True))

If m > Max_Width Then m = Max_Width
If n > Max_Height Then n = Max_Height

ReDim pDot(0 To m - 1, 0 To n - 1)

tmp = dbGetSettingEx("Effects\SoftEx\" + li, "Data", vbString, "000000CCCCCC000000CCCCCCFFFFFFCCCCCC000000CCCCCC000000", True)

Counter = 0
On Error Resume Next
For i = 0 To n - 1
    For j = 0 To m - 1
        Counter = Counter + 1
        pDot(j, i) = CLng("&H" + Mid$(tmp, (i * m + j) * 6 + 1, 6))
    Next j
Next i
strName = dbGetSettingEx("Effects\SoftEx\" + li, "Name", vbString, "<No Name>", True)
End Sub
Private Sub btnLoadFile_Click()
End Sub

Public Function GetPresetCount() As Integer
GetPresetCount = Val(dbGetSettingEx("Effects\SoftEx", "Count", vbLong, 0, True))
End Function

Public Sub SetPresetCount(ByVal Count As Integer)
dbSaveSettingEx "Effects\SoftEx", "Count", Count, True
End Sub

Public Sub MovePreset(ByVal Number As Integer, ByVal Dest As Integer)
Dim i As Integer, i2 As Integer, tmp As String
Dim tDot1() As Long, tDot2() As Long
Dim tName1 As String, tName2 As String
i = Number
i2 = Dest
If (i < 0) Or (i > GetPresetCount - 1) Or (i2 < 0) Or (i2 > GetPresetCount - 1) Then
    vtBeep
    Exit Sub
End If
LoadMask tDot1, i, tName1
LoadMask tDot2, i2, tName2
SaveMask tDot1, i2, tName1
SaveMask tDot2, i, tName2
'tmp = GetSetting("Common", App.Title + "\Effects\Soft\" + VedNull(i2, 2), "Name", "")
'SaveSetting "Common", App.Title + "\Effects\Soft\" + VedNull(i2, 2), "Name", _
'            GetSetting("Common", App.Title + "\Effects\Soft\" + VedNull(i, 2), "Name", "")
'SaveSetting "Common", App.Title + "\Effects\Soft\" + VedNull(i, 2), "Name", tmp
'
'
'tmp = GetSetting("Common", App.Title + "\Effects\Soft\" + VedNull(i2, 2), "Width", "")
'SaveSetting "Common", App.Title + "\Effects\Soft\" + VedNull(i2, 2), "Width", _
'            GetSetting("Common", App.Title + "\Effects\Soft\" + VedNull(i, 2), "Width", "")
'SaveSetting "Common", App.Title + "\Effects\Soft\" + VedNull(i, 2), "Width", tmp
'
'
'tmp = GetSetting("Common", App.Title + "\Effects\Soft\" + VedNull(i2, 2), "Height", "")
'SaveSetting "Common", App.Title + "\Effects\Soft\" + VedNull(i2, 2), "Height", _
'            GetSetting("Common", App.Title + "\Effects\Soft\" + VedNull(i, 2), "Height", "")
'SaveSetting "Common", App.Title + "\Effects\Soft\" + VedNull(i, 2), "Height", tmp
'
'
'tmp = GetSetting("Common", App.Title + "\Effects\Soft\" + VedNull(i2, 2), "Data", "")
'SaveSetting "Common", App.Title + "\Effects\Soft\" + VedNull(i2, 2), "Data", _
'            GetSetting("Common", App.Title + "\Effects\Soft\" + VedNull(i, 2), "Data", "")
'SaveSetting "Common", App.Title + "\Effects\Soft\" + VedNull(i, 2), "Data", tmp
dbFreshList
List1.ListIndex = i2
End Sub

Private Sub btnMoveDown_Click()
MovePreset List1.ListIndex, List1.ListIndex + 1
End Sub

Private Sub btnMoveUp_Click()
MovePreset List1.ListIndex, List1.ListIndex - 1
End Sub

Private Sub btnReName_Click()

End Sub

Private Sub btnSaveFile_Click()
End Sub

Private Sub CancelButton_Click()
Me.Tag = "c"
Me.Hide
End Sub

Private Sub btnTurn_Click()
End Sub

Private Sub Form_Load()
dbLoadCaptions
ValidateSize
If (GetPresetCount = 0) And (FirstLoad) Then
    InitializeRegFilters
End If
dbFreshList
'LoadSettings
End Sub

'Private Sub LoadSettings()
'Dim tmp As String, i As Long, j As Long, n As Long, m As Long
'Dim Counter As Long
'On Error GoTo eh
'LoadMask Dot, "LastUsed", ""
'dRefr
'dbFreshList
'SetFilterMode dbGetSettingEx("Effects\Soft", "FilterMode", vbLong, 0)
'nmbGain.Value = dbGetSettingEx("Effects\Soft", "Gain", vbDouble, 0#)
'chkTexMode.Value = Abs(dbGetSettingEx("Effects\Soft", "TextureMode", vbBoolean, False))
'Exit Sub
'eh:
'MsgBox "Initilaizing frmSoft failed.", vbCritical
'dbDeleteSetting "Effects\SoftEx\LastUsed", True
'End Sub

Sub dbFreshList()
Dim n As Long, i As Long
On Error GoTo eh
n = GetPresetCount
List1.Clear
For i = 1 To n
    List1.AddItem dbGetSettingEx("Effects\SoftEx\" + VedNull(i, 2), "Name", vbString, "<no name>", True)
Next i
eh:
End Sub

'Public Function SaveSettings()
'Dim li As String, tmp As String, i As Long, j As Long
'SaveMask Dot, "LastUsed", ""
'dbSaveSettingEx "Effects\Soft", "FilterMode", GetFilterMode
'dbSaveSettingEx "Effects\Soft", "Gain", nmbGain.Value
'dbSaveSettingEx "Effects\Soft", "TextureMode", CBool(chkTexMode.Value)
'End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If UnloadMode = 0 Then
    Cancel = True
    CancelButton_Click
End If
End Sub

Private Sub HFlip_Click()
End Sub

Private Sub List1_DblClick()
Dim tName As String
LoadMask Dot, List1.ListIndex, tName
dRefr
pChange
End Sub

Private Sub dRefr()
Dim i As Long, j As Long, tmp As Long
View.Cls
View.Move View.Left, View.Top, _
          ((UBound(Dot, 1) + 1) * 8 + 4) * Screen.TwipsPerPixelX, _
          ((UBound(Dot, 2) + 1) * 8 + 4) * Screen.TwipsPerPixelY
'For i = 0 To UBound(Dot, 2)
'    For j = 0 To UBound(Dot, 1)
'        View.Line (j * 8, i * 8)-((j + 1) * 8 - 1, (i + 1) * 8 - 1), Dot(j, i), BF
'    Next j
'Next i

DontDoEvents = True
RefrEx View.Image.Handle, View.hDC, Dot, 8, dbNoGrid
DontDoEvents = False
End Sub

Private Sub List1_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
Case 32
    List1_DblClick
End Select
End Sub

Private Sub List1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then PopupMenu mnuListPopup
End Sub

Private Sub mnuFPPLoad_Click()
Dim File As String, tData() As Long
Dim Alpha() As Long
On Error GoTo eh
'    With pCDl
'        .Flags = 0
'        .OpenFlags = cdlOFNFileMustExist
'        .Filter = GetDlgFilter(dbBLoad)
'        .FileName = ""
'        .ShowOpen
'        File = .FileName
'        .InitDir = GetDirName(File)
'    End With
    vtLoadPicture tData, Alpha, "", ShowDialog:=True, Purpose:="FilterMask"
    If UBound(tData, 1) + 1 > Max_Width Or UBound(tData, 2) + 1 > Max_Height Then
        dbMsgBox 2418, vbInformation 'Too large
        Err.Raise dbCWS
    End If
    Resize UBound(tData, 1) + 1, UBound(tData, 2) + 1
    Dot = tData
    Erase tData
    dRefr
    pChange
Exit Sub
eh:
If Err.Number = dbCWS Then
    Exit Sub
Else
    MsgBox Err.Description, vbCritical
End If
Exit Sub

End Sub

Private Sub mnuFPPSave_Click()
Dim File As String ', tData() As Byte
Dim Alpha() As Long
On Error GoTo eh
'    With pCDl
'        .Flags = 0
'        .OpenFlags = cdlOFNOverwritePrompt Or cdlOFNNoReadOnlyReturn
'        .Filter = GetDlgFilter(dbBSave)
'        .FileName = ""
'        .ShowSave
'        File = .FileName
'        .InitDir = GetDirName(File)
'    End With
    vtSavePicture Dot, Alpha, FileName:="", ShowDialog:=True, Purpose:="FilterMask"
    'Resize UBound(tData, 1) + 1, UBound(tData, 2) + 1
    'Dot = tData
    'Erase tData
    'dRefr
Exit Sub
eh:
If Err.Number = dbCWS Then
    Exit Sub
Else
    MsgBox Err.Description, vbCritical
End If
Exit Sub
End Sub

Private Sub mnuLPPAddStd_Click()
InitializeRegFilters
dbFreshList
mnuLPPAddStd.Enabled = False
End Sub

Private Sub mnuLPPDelete_Click()
Dim i As Integer, j As Integer
Dim n As Integer, ID As Integer, t As String

Dim tName As String
Dim tDot() As Long

ID = List1.ListIndex
If ID = -1 Then
    dbMsgBox 1147, vbInformation '"Select an item to delete`Mask Manager"
    Exit Sub
End If
'ID = ID + 1
If dbMsgBox(GRSF(1148), vbYesNo Or vbDefaultButton2) = vbNo Then Exit Sub
            '"Are you sure you want to delete this mask`Mask Manager"
n = GetPresetCount
For i = ID + 1 To n - 1
'    t = GetSetting("Common", "SMB Maker\Effects\Soft\" + VedNull(i, 2), "Name", "<No Name>")
'    SaveSetting "Common", "SMB Maker\Effects\Soft\" + VedNull(i - 1, 2), "Name", t
'    t = GetSetting("Common", "SMB Maker\Effects\Soft\" + VedNull(i, 2), "Width", "1")
'    SaveSetting "Common", "SMB Maker\Effects\Soft\" + VedNull(i - 1, 2), "Width", t
'    t = GetSetting("Common", "SMB Maker\Effects\Soft\" + VedNull(i, 2), "Height", "1")
'    SaveSetting "Common", "SMB Maker\Effects\Soft\" + VedNull(i - 1, 2), "Height", t
'    t = GetSetting("Common", "SMB Maker\Effects\Soft\" + VedNull(i, 2), "Data", "3")
'    SaveSetting "Common", "SMB Maker\Effects\Soft\" + VedNull(i - 1, 2), "Data", t
    LoadMask tDot, i, tName
    SaveMask tDot, i - 1, tName
Next i
dbDeleteSetting "Effects\Soft\" + VedNull(i - 1, 2), True
SetPresetCount n - 1
dbFreshList
'
End Sub

Private Sub mnuLPPOverwrite_Click()
If List1.ListIndex = -1 Then
    dbMsgBox GRSF(1151), vbInformation '"Select an item to rename`Mask Manager"
    Exit Sub
End If
SaveMask Dot, List1.ListIndex, dbGetSettingEx("Effects\SoftEx\" + VedNull(List1.ListIndex + 1, 2), "Name", vbString, "No Name Found", True)
dbFreshList
End Sub

Private Sub mnuLPPRename_Click()
Dim strName As String
If List1.ListIndex = -1 Then
    dbMsgBox GRSF(1149), vbInformation '"Select an item to rename`Mask Manager"
    Exit Sub
End If
strName = dbGetSettingEx("Effects\SoftEx\" + VedNull(List1.ListIndex + 1, 2), "Name", vbString, "No Name Found", True)
strName = dbInputBox(GRSF(1150), strName) '"Please input the name"
If Not (strName = "") Then
    dbSaveSettingEx "Effects\SoftEx\" + VedNull(List1.ListIndex + 1, 2), "Name", strName, True
    dbFreshList
End If
'
End Sub

Private Sub mnuOPPClear_Click()
ReDim Dot(0 To UBound(Dot, 1), 0 To UBound(Dot, 2))
View.Cls
pChange
'
End Sub

Private Sub mnuOPPEdit_Click()
On Error GoTo eh
EditPicture Dot
TruncDot Dot
dRefr
pChange
Exit Sub
eh:
If Err.Number = dbCWS Then
    Exit Sub
End If
End Sub

Private Sub mnuOPPFilpV_Click()
Dim w As Long, h As Long
Dim i As Long, j As Long
Dim tmpDot() As Long
tmpDot = Dot
w = UBound(Dot, 1)
h = UBound(Dot, 2)

For i = 0 To h
    For j = 0 To w
        Dot(j, i) = tmpDot(j, h - i)
    Next j
Next i
Erase tmpDot
dRefr
pChange
End Sub

Private Sub mnuOPPFlipH_Click()
Dim w As Long, h As Long
Dim i As Long, j As Long
Dim tmpDot() As Long
tmpDot = Dot
w = UBound(Dot, 1)
h = UBound(Dot, 2)

For i = 0 To h
    For j = 0 To w
        Dot(j, i) = tmpDot(w - j, i)
    Next j
Next i
Erase tmpDot
dRefr
pChange
End Sub

Private Sub mnuOPPGenCircle_Click()
Dim Diam As Double
On Error GoTo eh
Diam = dbGetSettingEx("SoftEx", "GenCircle diameter", vbDouble, 8)
EditNumber Diam, 1195, 1, Max_Width '"Diameter of the circle to create, in pixels:"
dbSaveSettingEx "SoftEx", "GenCircle diameter", Diam
Dim w As Long
w = -Int(-(Diam - 1) / 2) * 2 + 1
Resize w, w, PreserveValues:=False
DrawingEngine.AntiAliasingSharpness = 1
Dim vtx As vtVertex
vtx.Color = vbWhite
vtx.Weight = Diam
vtx.x = (w - 1) / 2
vtx.y = (w - 1) / 2
Dim fdsc As FadeDesc
Dim Pixels() As AlphaPixel
Dim nMem As Long
DrawingEngine.pntGradientLineHQ vtx, vtx, fdsc, Pixels, nMem
Dim i As Long
Dim x As Long, y As Long
For i = 0 To nMem - 1
  x = Pixels(i).x
  y = Pixels(i).y
  If x >= 0 And y >= 0 And x <= w - 1 And y <= w - 1 Then
    Dot(x, y) = Pixels(i).drawOpacity * &H10101
  End If
Next i
dRefr
pChange
Exit Sub
eh:
MsgError
End Sub

Private Sub mnuOPPRotate_Click()
Dim i As Long, j As Long, tmpDot() As Long, w As Long, h As Long
tmpDot = Dot
w = UBound(Dot, 1)
h = UBound(Dot, 2)
Resize h + 1, w + 1
For i = 0 To UBound(Dot, 2)
    For j = 0 To UBound(Dot, 1)
        Dot(j, i) = tmpDot(w - i, j)
    Next j
Next i
dRefr
pChange
End Sub

Private Sub mnuOPPUseSel_Click()
If MainForm.SelectionPresent Then
    Dot = CurSel_SelData
    TruncDot Dot
    MainForm.dbDeselect False
    dRefr
    pChange
Else
    dbMsgBox 2419, vbInformation 'no selection
End If
End Sub

Private Sub nmbGain_Change()
'vChange
End Sub

Private Sub nmbGain_InputChange()
pChange
End Sub

Private Sub OkButton_Click()
Me.Tag = ""
'SaveSettings
Me.Hide
End Sub

'Public Sub ExtractData(ByRef tDot() As Byte)
'Erase tDot
'tDot = Dot
'End Sub

Friend Sub ExtractData2(ByRef Mask As FilterMask)
Dim x As Long, y As Long
Dim w As Long, h As Long
Dim rgb1 As RGBQUAD
w = UBound(Dot, 1)
h = UBound(Dot, 2)
ReDim Mask.Mask(0 To w, 0 To h)
For y = 0 To h
    For x = 0 To w
        GetRgbQuadEx Dot(x, y), rgb1
        Mask.Mask(x, y).rgbRed = rgb1.rgbBlue
        Mask.Mask(x, y).rgbGreen = rgb1.rgbGreen
        Mask.Mask(x, y).rgbBlue = rgb1.rgbRed
    Next x
Next y
Mask.Center.x = -1
Mask.Center.y = -1
Mask.CenterFilled = False
End Sub

Friend Function SetData2(ByRef Mask As FilterMask)
Dim MaskW As Long, MaskH As Long
Dim x As Long, y As Long
Erase Dot
AryWH AryPtr(Mask.Mask), MaskW, MaskH
Resize MaskW, MaskH
For y = 0 To MaskH - 1
    For x = 0 To MaskW - 1
        Dot(x, y) = BGR(Mask.Mask(x, y).rgbRed, _
                        Mask.Mask(x, y).rgbGreen, _
                        Mask.Mask(x, y).rgbBlue)
    Next x
Next y
dRefr
End Function

Private Sub optMode_Click(Index As Integer)
pChange
End Sub

Private Sub Savebtn_Click()
Dim tmp As String
tmp = dbInputBox(GRSF(1143), "")
If tmp = "" Then Exit Sub
SetPresetCount GetPresetCount + 1
SaveMask Dot, GetPresetCount - 1, tmp
dbFreshList
End Sub

Private Sub Sizer_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button > 0 Then
x = x - Sizer.Width \ 2 + Sizer.Left
y = y - Sizer.Height \ 2 + Sizer.Top
Resize Round(((x - View.Left) / Screen.TwipsPerPixelX - 4) / 8), Round(((y - View.Top) / Screen.TwipsPerPixelY - 4) / 8)
pChange
End If
End Sub

Private Sub tmrUpdater_Timer()
tmrUpdater.Enabled = False
RaiseEvent Change
End Sub

Private Sub VFlip_Click()
End Sub

Private Sub View_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim i As Long, j As Long, ic As Long, jc As Long
i = y \ 8
j = x \ 8
If i < 0 Or j < 0 Or i > UBound(Dot, 2) Or j > UBound(Dot, 1) Then Exit Sub
If Button = 1 Then
    If mon = 1 Then
        Dot(j, i) = Dot(j, i) Or vbWhite
    ElseIf mon = 2 Then
        Dot(j, i) = 0
    ElseIf mon = 0 Then
        Dot(j, i) = Not (Dot(j, i)) And &HFFFFFF 'Abs(Not (CBool(Dot(i, j))))
        If Dot(j, i) = &HFFFFFF Then mon = 1 Else mon = 2
    End If
    
    View.Line (j * 8, i * 8)-((j + 1) * 8 - 1, (i + 1) * 8 - 1), Dot(j, i), BF
ElseIf Button = 2 Then
End If
End Sub

Sub Resize(ByVal w As Long, ByVal h As Long, Optional ByVal PreserveValues As Boolean = True)
Dim tmpData() As Long
Dim i As Long, j As Long
If h <= 0 Then h = 1
If w <= 0 Then w = 1
If h > Max_Height Then h = Max_Height
If w > Max_Width Then w = Max_Width
On Error GoTo eh
If PreserveValues Then
    tmpData = Dot
End If
red:
ReDim Dot(0 To w - 1, 0 To h - 1)
If PreserveValues And AryDims(AryPtr(tmpData)) = 2 Then
    For i = 0 To Min(h - 1, UBound(tmpData, 2))
        For j = 0 To Min(w - 1, UBound(tmpData, 1))
            Dot(j, i) = tmpData(j, i)
        Next j
    Next i
End If
On Error GoTo 0
View.Move View.Left, View.Top, (w * 8 + 4) * Screen.TwipsPerPixelX, (h * 8 + 4) * Screen.TwipsPerPixelY
'View.Width = (w * 8 + 4) * Screen.TwipsPerPixelX
'View.Height = (h * 8 + 4) * Screen.TwipsPerPixelY
Sizer.Move View.Width + View.Left, View.Height + View.Top
'Sizer.Top = View.Height + View.Top
'Sizer.Left = View.Width + View.Left
If PreserveValues Then
    dRefr
Else
    View.Cls
End If
Exit Sub
eh:
    PreserveValues = False
    Resume red
End Sub

Function cTmp(t As Byte) As Long
Select Case t
    Case 0
        cTmp = vbWhite
    Case 1
        cTmp = vbBlue
    Case 2
        cTmp = vbRed
    Case 3
        cTmp = vbMagenta
End Select
End Function

Private Sub View_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 1 Then View_MouseDown Button, Shift, x, y
End Sub

'Private Sub View_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'
'End Sub
Private Sub View_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
mon = 0
pChange
End Sub

Sub dbLoadCaptions()
btnMoveUp.Caption = Chr$(&H87)
btnMoveDown.Caption = Chr$(&H88)
Resr1.LoadCaptions
End Sub

Private Sub ValidateSize()
Dim YBorder As Long, Spacing As Long
Spacing = Shablon.Top
YBorder = Me.Height - Me.ScaleHeight
'Me.Move Me.Left, Me.Top, Me.Width, Shablon.Top + Shablon.Height + YBorder + Spacing
End Sub

Private Sub View_Resize()
Sizer.Move View.Left + View.Width, _
           View.Top + View.Height
End Sub

Private Sub InitializeRegFilters()
Dim Index As Long
Dim i As Long, j As Long, tData() As Long
Dim tmp As String, nLen As Long, w As Long, h As Long, cI As Long
Dim Begindex As Long
Dim rgb1 As RGBQUAD
Dim pName As String
Begindex = GetPresetCount
dbFreshList
For Index = 0 To 99
    If Begindex + Index >= 99 Then Exit For
    On Error GoTo eh
    LoadResSMB Index + 101, "FILTER", tData
    On Error GoTo 0
    cI = 0
    
    pName = GRSF(1600 + Index)
    For i = 0 To List1.ListCount - 1
        If List1.List(i) = pName Then Exit For
    Next i
    If i = List1.ListCount Then
        i = Begindex
        Begindex = Begindex + 1
    End If
    SaveMask tData, i, pName
            
Next Index

ExitHere:
'    j = GetPresetCount
'    If Index > j Then
'        SetPresetCount Index
'    End If
SetPresetCount Begindex
Exit Sub
eh:
Resume ExitHere
End Sub

'Public Function GetFlags() As Long
'Dim i As Long, Rslt As Long
'For i = 0 To optMode.UBound
'    Rslt = Rslt Or (CLng(CBool(optMode(i).Value)) And CLng(optMode(i).Tag))
'Next i
'GetFlags = Rslt
'End Function


'Friend Function GetBrightness() As Double
'GetBrightness = 10 ^ (nmbGain.Value / 10)
'End Function

Friend Function GetFilterMode() As Long
Dim Opt As OptionButton
For Each Opt In optMode
    If Opt.Value Then
        GetFilterMode = CLng(Opt.Tag)
        Exit For
    End If
Next
End Function

Friend Sub SetFilterMode(ByVal NewFM As eFilterMode)
Dim Opt As OptionButton
For Each Opt In optMode
    If CLng(Opt.Tag) = NewFM Then
        Opt.Value = True
        Exit For
    End If
Next
End Sub
