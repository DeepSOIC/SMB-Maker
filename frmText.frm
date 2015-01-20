VERSION 5.00
Begin VB.Form frmText 
   BackColor       =   &H00E3DFE0&
   Caption         =   "Вставка текста"
   ClientHeight    =   4037
   ClientLeft      =   2772
   ClientTop       =   3762
   ClientWidth     =   6545
   Icon            =   "frmText.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4037
   ScaleWidth      =   6545
   StartUpPosition =   1  'CenterOwner
   Begin SMBMaker.ctlResourcizator Resr1 
      Left            =   242
      Top             =   3531
      _ExtentX        =   1097
      _ExtentY        =   671
      ResID           =   9921
   End
   Begin SMBMaker.ctlColor ColorSel 
      Height          =   375
      Left            =   5820
      TabIndex        =   6
      ToolTipText     =   "Цвет текста"
      Top             =   1185
      Width           =   375
      _ExtentX        =   671
      _ExtentY        =   671
   End
   Begin SMBMaker.ctlNumBox nmbSize 
      Height          =   525
      Left            =   690
      TabIndex        =   5
      Top             =   3360
      Width           =   5205
      _ExtentX        =   9185
      _ExtentY        =   935
      Value           =   20
      Max             =   2160
      HorzMode        =   0   'False
      EditName        =   "Размер шрифта."
   End
   Begin VB.CheckBox chkTrans 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Прозрачный"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.07
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Текст будет отрисован в маску прозрачности."
      Top             =   2325
      Width           =   1530
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.07
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3195
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Text            =   "frmText.frx":014A
      Top             =   0
      Width           =   4590
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Цвет: "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.07
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   187
      Left            =   4664
      TabIndex        =   8
      Top             =   1210
      Width           =   451
   End
   Begin SMBMaker.dbButton btnQuality 
      Height          =   555
      Left            =   4680
      TabIndex        =   7
      ToolTipText     =   "Результат зависит от настройки сглаживания шрифтов в Windows."
      Top             =   1650
      Width           =   1530
      _ExtentX        =   2703
      _ExtentY        =   975
      MouseIcon       =   "frmText.frx":0154
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.064
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Others          =   $"frmText.frx":0170
      OthersPresent   =   -1  'True
   End
   Begin SMBMaker.dbButton FntButton 
      Height          =   375
      Left            =   4680
      TabIndex        =   1
      Top             =   2760
      Width           =   1530
      _ExtentX        =   2703
      _ExtentY        =   671
      MouseIcon       =   "frmText.frx":01C4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.064
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Others          =   $"frmText.frx":01E0
      OthersPresent   =   -1  'True
   End
   Begin SMBMaker.dbButton OkButton 
      Height          =   390
      Left            =   4725
      TabIndex        =   2
      ToolTipText     =   "(Shift + Enter)"
      Top             =   90
      Width           =   1530
      _ExtentX        =   2703
      _ExtentY        =   691
      MouseIcon       =   "frmText.frx":0233
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.064
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Others          =   $"frmText.frx":024F
      OthersPresent   =   -1  'True
   End
   Begin SMBMaker.dbButton CancelButton 
      Cancel          =   -1  'True
      Height          =   390
      Left            =   4725
      TabIndex        =   3
      Top             =   570
      Width           =   1530
      _ExtentX        =   2703
      _ExtentY        =   691
      MouseIcon       =   "frmText.frx":029B
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.064
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Others          =   $"frmText.frx":02B7
      OthersPresent   =   -1  'True
   End
   Begin VB.Menu mnuPPQuality 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu mnuDummy1 
         Caption         =   "-=[ Качество отрисовки текста ]=-"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuPPQItem 
         Caption         =   "полное"
         Index           =   0
      End
      Begin VB.Menu mnuPPQItem 
         Caption         =   "без ClearType"
         Index           =   1
      End
      Begin VB.Menu mnuPPQItem 
         Caption         =   "без сглаживания"
         Index           =   2
      End
   End
End
Attribute VB_Name = "frmText"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
#Const Testing = True
'Dim lngColor As Long
Option Explicit

Dim pQuality As eTextQuality


Public Enum eTextQuality
  etqFull = 0
  etqGray = 1
  etqMono = 2
End Enum


Private Sub btnQuality_Click()
PopupMenu mnuPPQuality, vbPopupMenuLeftAlign, btnQuality.Left + btnQuality.Width, btnQuality.Top
End Sub

Private Sub CancelButton_Click()
Me.Tag = "C"
Me.Hide
End Sub

Private Sub Form_Activate()
On Error Resume Next
Text1.SetFocus
Text1.SelStart = 0
Text1.SelLength = Len(Text1.Text)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Dim ShiftKeyCode As Long
ShiftKeyCode = GetShiftKeyCode(KeyCode, Shift)
Select Case ShiftKeyCode
  Case 269
    KeyCode = 0
    OkButton_Click
End Select
End Sub

Private Sub Form_Paint()
On Error Resume Next
Me.PaintPicture LoadResPicture(RES_BackPicture, vbResBitmap), 0, 0, Me.ScaleWidth, Me.ScaleHeight
End Sub
'
'Public Sub UpdateColorSel()
'ColorSel.Color = ConvertColorLng(lngColor)
'End Sub
'
'Private Sub ColorSel_Click()
'On Error GoTo CWS
'lngColor = CDl.PickColor(lngColor)
'UpdateColorSel
'Exit Sub
'CWS:
'If Err.Number = dbCWS Then
'    Exit Sub
'Else
'    MsgBox Err.Description
'End If
'End Sub

Private Sub FntButton_Click()
#If Testing Then
    Dim CDl As New CommonDlg
    CDl.CancelError = True
#End If
Dim fnt As New StdFont
On Error GoTo eh
Set fnt = Text1.Font
With CDl
    .Flags = 0
    .FontFlags = cdlCFEffects
    .FontName = fnt.Name
    .FontBold = fnt.Bold
    .FontItalic = fnt.Italic
    .FontSize = fnt.Size
    .FontStrikeThrough = fnt.Strikethrough
    .FontUnderline = fnt.Underline
    .FontScript = fnt.Charset
'    .FontWeight = fnt.Weight
    .ShowFont
    fnt.Name = .FontName
    fnt.Bold = .FontBold
    fnt.Italic = .FontItalic
    fnt.Size = .FontSize
    fnt.Strikethrough = .FontStrikeThrough
    fnt.Underline = .FontUnderline
    fnt.Charset = .FontScript
End With

Set Text1.Font = fnt
Text1.Font.Size = fnt.Size
nmbSize.Value = fnt.Size
Exit Sub
eh:
If Err.Number <> dbCWS Then
    MsgBox Err.Description, vbExclamation
Else
    Exit Sub
End If
End Sub

Private Sub Form_Load()
LoadCaptions
LoadSettings
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If UnloadMode = vbFormControlMenu Then
    Cancel = True
    CancelButton_Click
End If
End Sub

Private Sub Form_Resize()
On Error Resume Next
Text1.Move 0, 0, ScaleWidth - OkButton.Width, ScaleHeight - nmbSize.Height
OkButton.Move ScaleWidth - OkButton.Width, 0
CancelButton.Move OkButton.Left, OkButton.Top + OkButton.Height, OkButton.Width, OkButton.Height
nmbSize.Move 0, ScaleHeight - nmbSize.Height, ScaleWidth
FntButton.Move OkButton.Left, nmbSize.Top - OkButton.Height, OkButton.Width, OkButton.Height
chkTrans.Move OkButton.Left, FntButton.Top - chkTrans.Height, OkButton.Width
btnQuality.Move OkButton.Left, chkTrans.Top - btnQuality.Height, OkButton.Width
ColorSel.Move ScaleWidth - ColorSel.Width, btnQuality.Top - ColorSel.Height
Label1.Move ColorSel.Left - Label1.Width, ColorSel.Top + (ColorSel.Height - Label1.Height) \ 2
nmbSize.ZOrder vbBringToFront
End Sub

Private Sub mnuPPQItem_Click(Index As Integer)
pQuality = Index
RefreshQuality
End Sub

Private Sub nmbSize_Change()
Dim s As Single
On Error Resume Next
Text1.Font.Size = nmbSize.Value
Text1.Refresh
End Sub

Private Sub nmbSize_InputChange()
nmbSize_Change
End Sub

Private Sub OkButton_Click()
If Not (TestControls) Then
    Exit Sub
End If
SaveSettings
Me.Tag = ""
Me.Hide

End Sub

Private Function TestControls() As Boolean
Dim v As Variant
TestControls = True
Exit Function
Fail:
    TestControls = False
    vtBeep
Exit Function
End Function

Private Sub Text1_GotFocus()
OkButton.Default = False
OkButton.Default1 = False
End Sub

Private Sub Text1_LostFocus()
OkButton.Default = True
OkButton.Default1 = True
End Sub

'Private Sub txtSize_Change()
'End Sub

'Private Sub txtSize_GotFocus()
'txtSize.SelStart = 0
'txtSize.SelLength = Len(txtSize.Text)
'End Sub

Private Sub LoadCaptions()
Resr1.LoadCaptions
'OkButton.Caption = GRSF(2233)
'CancelButton.Caption = GRSF(2234)
'FntButton.Caption = GRSF(2235)
'nmbSize.EditName = GRSF(2236)
'Me.Caption = GRSF(2237)
''Text1.Text = grsf(2238)
'ColorSel.ToolTipText = GRSF(2239)
'chkTrans.Caption = GRSF(2382)
'chkTrans.ToolTipText = GRSF(2383)
End Sub

Public Sub SaveSettings()
Dim fnt As New StdFont
Set fnt = Text1.Font
dbSaveSetting "Tool", "FontBold", CStr(fnt.Bold)
dbSaveSetting "Tool", "FontCharset", CStr(fnt.Charset)
dbSaveSetting "Tool", "FontItalic", CStr(fnt.Italic)
dbSaveSetting "Tool", "FontName", CStr(fnt.Name)
dbSaveSetting "Tool", "FontStrikeThrough", CStr(fnt.Strikethrough)
dbSaveSetting "Tool", "FontUnderline", CStr(fnt.Underline)
dbSaveSetting "Tool", "FontSize", Trim(Str(Text1.FontSize))
dbSaveSetting "Tool", "TextToTrans", CStr(CBool(chkTrans.Value))
dbSaveSetting "Tool", "LastText", Text1.Text
dbSaveSettingEx "Tool", "RenderQuality", pQuality
End Sub

Public Sub LoadSettings()
Dim fnt As New StdFont
On Error GoTo eh
'Set fnt = Text1.Font
fnt.Bold = CBool(dbGetSetting("Tool", "FontBold", "False"))
fnt.Charset = CLng(dbGetSetting("Tool", "FontCharset", CStr(RUSSIAN_CHARSET)))
fnt.Italic = CBool(dbGetSetting("Tool", "FontItalic", "False"))
fnt.Name = dbGetSetting("Tool", "FontName", "Comic Sans MS")
fnt.Size = Val(dbGetSetting("Tool", "FontSize", "20"))
fnt.Strikethrough = CBool(dbGetSetting("Tool", "FontStrikeThrough", "False"))
fnt.Underline = CBool(dbGetSetting("Tool", "FontUnderline", "False"))
Set Text1.Font = fnt
nmbSize.Value = fnt.Size
chkTrans.Value = Abs(CBool(dbGetSetting("Tool", "TextToTrans", CStr(True))))
Text1.Text = dbGetSetting("Tool", "LastText", GRSF(2238))

pQuality = dbGetSettingEx("Tool", "RenderQuality", vbLong)
RefreshQuality
Exit Sub
eh:

End Sub

Public Sub SetColor(ByVal clr As Long)
'lngColor = clr And &HFFFFFF
ColorSel.Color = clr And &HFFFFFF
'chkTrans.Value = Abs(CBool(Clr And &H1000000))
End Sub

Public Function GetColor() As Long
GetColor = ColorSel.Color Or &H1000000 * Abs(chkTrans.Value)
End Function

Public Function GetQuality() As eTextQuality
GetQuality = pQuality
End Function

Public Function RefreshQuality()
Dim mnu As Menu
For Each mnu In mnuPPQItem
  mnu.Checked = mnu.Index = pQuality
  If mnu.Checked Then btnQuality.Caption = GRSF(2385) + mnu.Caption
Next
End Function
