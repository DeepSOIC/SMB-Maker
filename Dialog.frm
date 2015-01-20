VERSION 5.00
Begin VB.Form Dialog 
   BackColor       =   &H00E3DFE0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Picture Size"
   ClientHeight    =   3990
   ClientLeft      =   21600
   ClientTop       =   8475
   ClientWidth     =   10395
   Icon            =   "Dialog.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3990
   ScaleWidth      =   10395
   StartUpPosition =   1  'CenterOwner
   Begin SMBMaker.dbFrame dbFrame3 
      Height          =   1815
      Left            =   75
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   1680
      Width           =   4440
      _ExtentX        =   7832
      _ExtentY        =   3201
      Caption         =   "Information"
      ResID           =   2459
      EAC             =   0   'False
      Begin VB.TextBox txtInfo 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   1290
         Left            =   255
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   5
         Text            =   "Dialog.frx":0ABA
         Top             =   330
         Width           =   3855
      End
   End
   Begin SMBMaker.dbFrame dbFrame2 
      Height          =   3450
      Left            =   4545
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   45
      Width           =   5805
      _ExtentX        =   10239
      _ExtentY        =   6085
      Caption         =   "Resize options"
      ResID           =   2457
      EAC             =   0   'False
      Begin SMBMaker.ctlTaggedText txtStretch 
         Height          =   2205
         Left            =   150
         TabIndex        =   6
         Top             =   1110
         Width           =   5490
         _ExtentX        =   9684
         _ExtentY        =   3889
         DisableHist     =   -1  'True
         No3D            =   -1  'True
         ForeColor       =   128
         BackColor       =   14737632
      End
      Begin VB.CheckBox chk 
         BackColor       =   &H00FFC0FF&
         Caption         =   "Stretch the picture"
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
         Height          =   525
         Left            =   90
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "select it to stretch the picture to new size."
         Top             =   300
         Value           =   1  'Checked
         Width           =   2805
      End
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
         ItemData        =   "Dialog.frx":0AC6
         Left            =   3210
         List            =   "Dialog.frx":0AC8
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   510
         Width           =   2415
      End
      Begin VB.Shape Shape1 
         Height          =   2235
         Left            =   135
         Top             =   1095
         Width           =   5520
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Selected method description:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   135
         TabIndex        =   15
         Top             =   885
         Width           =   5475
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Stretching method:"
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
         Left            =   3210
         TabIndex        =   13
         Top             =   255
         Width           =   2400
      End
   End
   Begin SMBMaker.dbFrame dbFrame1 
      Height          =   1605
      Left            =   75
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   45
      Width           =   4440
      _ExtentX        =   7832
      _ExtentY        =   2831
      Caption         =   "Size"
      ResID           =   2455
      EAC             =   0   'False
      Begin VB.CheckBox chkLockRatio 
         BackColor       =   &H00FF80FF&
         Caption         =   "Lock ratio"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   450
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   1020
         Width           =   3270
      End
      Begin SMBMaker.ctlNumBox nmbH 
         Height          =   330
         Left            =   1320
         TabIndex        =   1
         Top             =   600
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   582
         Value           =   2
         Min             =   2
         Max             =   9000
         NumType         =   2
         HorzMode        =   -1  'True
         EditName        =   "$244"
         NativeValues    =   "8|8|16|16|32|32|64|64|128|128|256|256|480|480|600|600|768|768|1200|1200|1536|1536"
         NativeValuesResID=   2446
      End
      Begin SMBMaker.ctlNumBox nmbW 
         Height          =   330
         Left            =   1320
         TabIndex        =   0
         Top             =   225
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   582
         Value           =   2
         Min             =   2
         Max             =   9000
         NumType         =   2
         HorzMode        =   -1  'True
         EditName        =   "$245"
         NativeValues    =   "8|8|16|16|32|32|64|64|128|128|256|256|640|640|800|800|1024|1024|1600|1600|2048|2048"
         NativeValuesResID=   2445
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Height"
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
         Left            =   113
         TabIndex        =   11
         Top             =   690
         Width           =   465
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Width"
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
         Left            =   135
         TabIndex        =   10
         Top             =   300
         Width           =   420
      End
   End
   Begin SMBMaker.dbButton CancelButton 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   9120
      TabIndex        =   8
      Top             =   3570
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      MouseIcon       =   "Dialog.frx":0ACA
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Others          =   $"Dialog.frx":0AE6
      OthersPresent   =   -1  'True
   End
   Begin SMBMaker.dbButton OKButton 
      Default         =   -1  'True
      Height          =   375
      Left            =   7830
      TabIndex        =   7
      Top             =   3570
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      MouseIcon       =   "Dialog.frx":0B2F
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Others          =   $"Dialog.frx":0B4B
      OthersPresent   =   -1  'True
   End
   Begin VB.Menu mnuPP 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu mnuMult2 
         Caption         =   "x2"
      End
      Begin VB.Menu mnuDiv2 
         Caption         =   "/ 2"
      End
      Begin VB.Menu mnuMult 
         Caption         =   "x #..."
      End
      Begin VB.Menu mnuDiv 
         Caption         =   "/ #..."
      End
      Begin VB.Menu mnuAdd 
         Caption         =   "+ #..."
      End
   End
End
Attribute VB_Name = "Dialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RClick As Integer
Dim AspW As Long, AspH As Long
Dim bLock As Boolean
Dim Names() As String, Descs() As String
Dim LockState As Integer '0 - nothing. 1-nmbw changed. 2-nmbh changing rightr after editing nmbw (not used)
Option Explicit

Private Sub CancelButton_Click()
Me.Tag = "c"
Me.Hide
End Sub

Private Sub chk_Click()
LockState = 0
Combo1.Enabled = CBool(Chk.Value)
End Sub

Private Sub chkLockRatio_Click()
LockState = 0
If chkLockRatio.Value Then
    AspW = nmbW.Value
    AspH = nmbH.Value
End If
UpdateInfo
End Sub

Private Sub Combo1_Change()
LockState = 0
If Combo1.ListIndex = -1 Then
    txtStretch.SetText ""
Else
    txtStretch.SetText Descs(Combo1.ItemData(Combo1.ListIndex))
End If
End Sub

Private Sub Combo1_Click()
Combo1_Change
End Sub

Private Sub Form_Paint()
Me.PaintPicture LoadResPicture(RES_BackPicture, vbResBitmap), 0, 0, Me.ScaleWidth, Me.ScaleHeight
End Sub

Private Sub Form_Load()
dbLoadCaptions
LoadStretches Combo1, Names, Descs
Combo1_Change
LoadSettings
End Sub

Function dbValidateControls() As Boolean
Dim Res As Boolean
Dim t As Long
On Error Resume Next
t = nmbW.Value
t = nmbH.Value
dbValidateControls = Err.Number = 0
End Function

Private Sub Form_QueryUnload(uCancel As Integer, UnloadMode As Integer)
If UnloadMode = VBRUN.QueryUnloadConstants.vbFormControlMenu Then
    uCancel = 1
    CancelButton_Click
Else
    'SaveSettings
End If
End Sub

Private Sub nmbH_Change()
If bLock Then Exit Sub
If chkLockRatio.Value Then
    If LockState <> 1 Then
        bLock = True
        On Error Resume Next
        nmbW.Value = CLng(AspW / AspH * nmbH.Value)
        On Error GoTo 0
        bLock = False
    End If
End If
UpdateInfo
End Sub

Private Sub nmbH_InputChange()
nmbH_Change
End Sub

Private Sub nmbW_Change()
If bLock Then Exit Sub
LockState = 1
If chkLockRatio.Value Then
    bLock = True
    On Error Resume Next
    nmbH.Value = CLng(AspH / AspW * nmbW.Value)
    On Error GoTo 0
    bLock = False
End If
UpdateInfo
End Sub

Private Sub nmbW_InputChange()
nmbW_Change
End Sub

Private Sub OkButton_Click()
If Not (dbValidateControls) Then Exit Sub
SaveSettings
Me.Tag = ""
Me.Hide
End Sub

Sub dbLoadCaptions()
Me.Caption = GRSF(2129)
Chk.Caption = GRSF(242)
Chk.ToolTipText = GRSF(243)
'txtV.ToolTipText = grsf(244)
'txtH.ToolTipText = grsf(245)
Combo1.List(0) = GRSF(2227)
Combo1.List(1) = GRSF(2228)
Combo1.ToolTipText = GRSF(2229)
chkLockRatio.Caption = GRSF(2456)
Label1.Caption = GRSF(2460)
Label2.Caption = GRSF(2461)
Label3.Caption = GRSF(2458)
Label4.Caption = GRSF(2517)

'Me.Icon = LoadResPicture(Me.Name, vbResIcon)
End Sub

Public Sub LoadSettings()
Dim MsgText As String, i As Long, Answ As VbMsgBoxResult
On Error GoTo eh
    MsgText = "Invalid Boolean value (StretchPicture)"
    Chk.Value = Abs(CBool(dbGetSetting("Imagesize", "StretchPicture", "True")))
    MsgText = "Bad StretchMethod"
    Combo1.ListIndex = CInt(dbGetSetting("ImageSize", "StretchMethod", CStr(eStretchMode.SMSquares)))
    chkLockRatio.Value = Abs(dbGetSettingEx("ImageSize", "LockAspectRatio", vbBoolean, True))
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
dbSaveSetting "ImageSize", "StretchPicture", CStr(CBool(Chk.Value))
dbSaveSetting "ImageSize", "StretchMethod", CStr(Combo1.ListIndex)
dbSaveSettingEx "ImageSize", "LockAspectRatio", CBool(chkLockRatio.Value)
End Sub

Friend Sub ExtractSz(ByRef Sz As Dims)
Sz.w = nmbW.Value
Sz.h = nmbH.Value
End Sub

Friend Sub SetSz(ByRef Sz As Dims)
bLock = True
nmbW.Value = Sz.w
nmbH.Value = Sz.h
AspW = Sz.w
AspH = Sz.h
bLock = False
UpdateInfo
End Sub

Friend Sub UpdateInfo()
Dim ImageSize As Currency
Dim w As Long, h As Long
Dim RatioSt As String
w = nmbW.Value
h = nmbH.Value
If Not CBool(chkLockRatio.Value) Then
    AspW = w
    AspH = h
    RatioSt = AspectRatioToStr(w, h)
Else
    RatioSt = AspectRatioToStr(w, h) + GRSF(2502) + AspectRatioToStr(AspW, AspH)
End If
On Error Resume Next
ImageSize = CCur(h) * CCur(-Int(-w * 3@ / 4@) * 4@)

txtInfo.Text = grs(2454, _
           "%pix", Format_Size(CCur(w) * CCur(h), 1000@, 3), _
           "%bmp", Format_Size(ImageSize + 54@, 1024@, 4), _
           "%mem", Format_Size(CCur(w) * CCur(h) * 4@, 1024@, 4), _
           "$asp", RatioSt)
End Sub

Sub LoadStretches(ByRef Combo As ComboBox, _
                  ByRef Names() As String, _
                  ByRef Descs() As String)
Const BaseResID As Long = 2462
Dim sArr() As String
Dim ID As Long, i As Long
Dim tmp As String
Dim NamesCnt As Long
ReDim Names(0 To 0)
ReDim Descs(0 To 0)
NamesCnt = 1
On Error GoTo eh1
ID = BaseResID
Do
    tmp = LoadResString(ID)
    If Len(tmp) > 0 Then
        sArr = Split(tmp, "|", 3)
        If UBound(sArr) <> 2 Then Exit Do
        i = Val(sArr(0))
        If i > NamesCnt - 1 Then
            NamesCnt = i + 1
            ReDim Preserve Names(0 To NamesCnt - 1)
            ReDim Preserve Descs(0 To NamesCnt - 1)
        End If
        Names(i) = Trim(sArr(1))
        Descs(i) = sArr(2)
    End If
    ID = ID + 1
Loop While tmp <> "<EOL>"
Combo1.Clear
For i = 0 To NamesCnt - 1
    If Len(Names(i)) > 0 Then
        Combo.AddItem Names(i)
        Combo.ItemData(Combo.NewIndex) = i
    End If
Next i

Exit Sub
eh1:
Debug.Assert False
End Sub

Function GetStretchMode() As eStretchMode
If Chk.Value Then
    If Combo1.ListIndex = -1 Then
        GetStretchMode = SMSquares
    Else
        GetStretchMode = Combo1.ItemData(Combo1.ListIndex)
    End If
Else
    GetStretchMode = SMPreserve
End If
End Function

Sub SetStretchMode(ByVal StretchMode As eStretchMode)
Dim i As Long
    For i = 0 To Combo1.ListCount - 1
        If Combo1.ItemData(i) = StretchMode Then
            Combo1.ListIndex = i
            Exit For
        End If
    Next i
End Sub

Private Sub txtInfo_GotFocus()
LockState = 0
End Sub

Private Sub txtStretch_GotFocus()
LockState = 0
End Sub

