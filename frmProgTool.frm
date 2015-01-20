VERSION 5.00
Begin VB.Form frmProgTool 
   Caption         =   "Tool Program"
   ClientHeight    =   4994
   ClientLeft      =   165
   ClientTop       =   484
   ClientWidth     =   7139
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   454
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   649
   StartUpPosition =   1  'CenterOwner
   Begin VB.ListBox pList 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.07
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1770
      IntegralHeight  =   0   'False
      ItemData        =   "frmProgTool.frx":0000
      Left            =   60
      List            =   "frmProgTool.frx":0007
      TabIndex        =   3
      Top             =   435
      Width           =   1590
   End
   Begin VB.TextBox txtPrg 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.79
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1350
      Left            =   2783
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   2222
      Width           =   3840
   End
   Begin SMBMaker.dbButton btnMoveDown 
      Height          =   390
      Left            =   75
      TabIndex        =   8
      Top             =   2220
      Width           =   1230
      _ExtentX        =   2174
      _ExtentY        =   691
      MouseIcon       =   "frmProgTool.frx":0014
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.064
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Others          =   $"frmProgTool.frx":0030
      OthersPresent   =   -1  'True
   End
   Begin SMBMaker.dbButton btnMoveUp 
      Height          =   390
      Left            =   75
      TabIndex        =   7
      Top             =   30
      Width           =   1230
      _ExtentX        =   2174
      _ExtentY        =   691
      MouseIcon       =   "frmProgTool.frx":007F
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.064
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Others          =   $"frmProgTool.frx":009B
      OthersPresent   =   -1  'True
   End
   Begin SMBMaker.dbButton btnRem 
      Height          =   495
      Left            =   210
      TabIndex        =   5
      Top             =   3915
      Width           =   1080
      _ExtentX        =   1910
      _ExtentY        =   874
      MouseIcon       =   "frmProgTool.frx":00EB
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.064
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Others          =   $"frmProgTool.frx":0107
      OthersPresent   =   -1  'True
   End
   Begin SMBMaker.dbButton btnAdd 
      Height          =   510
      Left            =   210
      TabIndex        =   4
      Top             =   2895
      Width           =   1080
      _ExtentX        =   1910
      _ExtentY        =   894
      MouseIcon       =   "frmProgTool.frx":015B
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.064
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Others          =   $"frmProgTool.frx":0177
      OthersPresent   =   -1  'True
   End
   Begin SMBMaker.dbButton OkButton 
      Height          =   420
      Left            =   2880
      TabIndex        =   1
      ToolTipText     =   "(Shift + Enter)"
      Top             =   4335
      Width           =   2010
      _ExtentX        =   3536
      _ExtentY        =   732
      MouseIcon       =   "frmProgTool.frx":01C8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.064
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Others          =   $"frmProgTool.frx":01E4
      OthersPresent   =   -1  'True
   End
   Begin SMBMaker.dbButton CancelButton 
      Cancel          =   -1  'True
      Height          =   420
      Left            =   4950
      TabIndex        =   2
      Top             =   4320
      Width           =   2010
      _ExtentX        =   3536
      _ExtentY        =   732
      MouseIcon       =   "frmProgTool.frx":0233
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.064
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Others          =   $"frmProgTool.frx":024F
      OthersPresent   =   -1  'True
   End
   Begin SMBMaker.dbButton btnRename 
      Height          =   495
      Left            =   210
      TabIndex        =   6
      Top             =   3420
      Width           =   1080
      _ExtentX        =   1910
      _ExtentY        =   874
      MouseIcon       =   "frmProgTool.frx":02A2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.064
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Others          =   $"frmProgTool.frx":02BE
      OthersPresent   =   -1  'True
   End
   Begin VB.Menu nmuFile 
      Caption         =   "File"
      Begin VB.Menu mnuLoad 
         Caption         =   "Create from file..."
      End
      Begin VB.Menu mnuSave 
         Caption         =   "Write file..."
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "Edit"
      Begin VB.Menu mnuEditSelAll 
         Caption         =   "Select all"
         Shortcut        =   ^A
      End
   End
   Begin VB.Menu mnuSpecial 
      Caption         =   "Special"
      Begin VB.Menu mnuUpgradeStorage 
         Caption         =   "Get programs from rev<229"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
   End
End
Attribute VB_Name = "frmProgTool"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Prg As SMP
Dim Vars() As Variable
Dim CurProgIndex As Long
Dim eLock As Boolean
Public DontCompile As Boolean
Public Section As String
Public Filt As DlgFilter
Dim pMode As eProgMode

Public Enum eProgMode
  epmTool = 0
  epmDraw = 1
End Enum

Dim Sgs As clsIniFile

Private Sub btnAdd_Click()
Dim tmp As String
Dim i As Long
On Error GoTo eh
tmp = dbInputBox(2394, "", True)
If Len(tmp) = 0 Then Err.Raise dbCWS
i = pList.ListIndex + 1
CreateProg i, tmp
Exit Sub
eh:
If Err.Number = dbCWS Then Exit Sub
MsgBox Err.Description
End Sub

Public Sub CreateProg(ByVal i As Long, ByRef pName As String)
Dim j As Long
If i = -1 Then i = pList.ListCount
SaveProg CurProgIndex
pList.AddItem pName, i
For j = pList.ListCount - 1 To i Step -1
    LoadProg j
    SaveProg j + 1
Next j
FlushList
eLock = True
pList.ListIndex = i
CurProgIndex = i
LoadProg 0, True
'txtPrg.Text = tmp
eLock = False
End Sub

Private Sub btnMoveDown_Click()
SwapProgs pList.ListIndex, pList.ListIndex + 1
End Sub

Private Sub btnMoveUp_Click()
SwapProgs pList.ListIndex, pList.ListIndex - 1
End Sub

Private Sub btnRem_Click()
Dim i As Long, j As Long
If pList.ListCount <= 1 Or pList.ListIndex = -1 Then
    vtBeep
    Exit Sub
End If
If dbMsgBox(2422, vbYesNo Or vbQuestion Or vbDefaultButton2) = vbNo Then
    Exit Sub
End If
i = pList.ListIndex
For j = i + 1 To pList.ListCount - 1
    LoadProg j
    SaveProg j - 1
Next j
DelProg pList.ListCount - 1
pList.RemoveItem i
FlushList
eLock = True
pList.ListIndex = IIf(i > pList.ListCount - 1, pList.ListCount - 1, i)
LoadProg pList.ListIndex
CurProgIndex = pList.ListIndex
eLock = False
End Sub

Private Sub SwapProgs(ByVal i1 As Long, ByVal i2 As Long)
Dim p1 As String, p2 As String
If i1 < 0 Or i1 > pList.ListCount - 1 Then Exit Sub
If i2 < 0 Or i2 > pList.ListCount - 1 Then Exit Sub
If i1 = i2 Then Exit Sub
ReadProg i1, p1
ReadProg i2, p2
WriteProg i1, p2
WriteProg i2, p1
p1 = pList.List(i1)
p2 = pList.List(i2)
pList.List(i1) = p2
pList.List(i2) = p1
FlushList
eLock = True
On Error GoTo eh
Dim i As Long
i = pList.ListIndex
Select Case i
  Case i1
    pList.ListIndex = i2
  Case i2
    pList.ListIndex = i1
End Select
eLock = False
Exit Sub
eh:
eLock = False
MsgError
End Sub

Friend Sub Restore()
DontCompile = False
Section = "Tool"
Filt = dbPToolLoad
pMode = epmTool
End Sub

Private Sub btnReName_Click()
Dim i As Long
Dim tmp As String
i = pList.ListIndex
If i = -1 Then
    vtBeep
    Exit Sub
End If
On Error GoTo eh
tmp = dbInputBox("Input the new name:", pList.List(i), True)
If tmp = "" Then Err.Raise dbCWS
On Error GoTo 0
pList.List(i) = tmp
FlushList
Exit Sub
eh:
If Err.Number = dbCWS Then
    Exit Sub
Else
    MsgBox Err.Description, vbCritical, "Error"
End If

End Sub

Private Sub CancelButton_Click()
SaveSettings
Restore
Me.Tag = "C"
Me.Hide
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Dim ShiftKeyCode As Long
ShiftKeyCode = GetShiftKeyCode(KeyCode, Shift)
If ShiftKeyCode = 112 Then
    KeyCode = 0
    ShowHelpWindow
End If
Select Case ShiftKeyCode
    Case 116 'F5
        OkButton_Click
    Case 269 'shift+enter
        KeyCode = 0
        OkButton_Click
End Select
End Sub

Private Sub Form_Load()
Restore
LoadSettings
LoadWindowPos frmProgTool
LoadCaptions
End Sub

Public Sub LoadCaptions()
Me.Caption = GRSF(Switch(pMode = epmTool, 2390, pMode = epmDraw, 2393))
mnuLoad.Caption = GRSF(2515)
mnuSave.Caption = GRSF(2516)
'txtHlp.Text = GRSF(10083)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
SaveWindowPos frmProgTool
If UnloadMode = 0 Then
    Cancel = True
    CancelButton_Click
End If
End Sub

Private Sub Form_Resize()
On Error Resume Next
OkButton.Move 0, Me.ScaleHeight - OkButton.Height, Me.ScaleWidth \ 2

CancelButton.Move OkButton.Width, OkButton.Top, Me.ScaleWidth - Me.ScaleWidth \ 2, OkButton.Height
'txtHlp.Move pList.Width, OkButton.Top - txtHlp.Height, Me.ScaleWidth - pList.Width
txtPrg.Move pList.Width, 0, Me.ScaleWidth - pList.Width, OkButton.Top

btnRem.Move 0, OkButton.Top - btnRem.Height, pList.Width
btnRename.Move 0, btnRem.Top - btnRename.Height, pList.Width
btnAdd.Move 0, btnRename.Top - btnAdd.Height, pList.Width
btnMoveDown.Move 0, btnAdd.Top - btnMoveDown.Height, pList.Width

btnMoveUp.Move 0, 0, pList.Width

pList.Move 0, btnMoveUp.Top + btnMoveUp.Height, pList.Width, btnMoveDown.Top - (btnMoveUp.Top + btnMoveUp.Height)
End Sub

Private Sub mnuEditSelAll_Click()
With txtPrg
  .SelStart = 0
  .SelLength = Len(.Text)
End With
End Sub

Private Sub mnuHelp_Click()
ShowHelpWindow
End Sub

Private Sub mnuLoad_Click()
Dim nmb As Long
Dim Files() As String
Dim tmp1 As String, tmp2 As String
Dim i As Long
On Error GoTo eh
Files = Split(ShowOpenDlg(Filt, Me.hWnd, cdlOFNAllowMultiselect, Purpose:=Section + "PROG"), Chr$(1))
For i = 0 To UBound(Files)
    nmb = FreeFile
    Open Files(i) For Input As nmb
        tmp1 = Input(LOF(nmb), nmb)
    Close nmb
    tmp2 = dbInputBox(2394, GetFileTitle(CropExt(Files(i))), True, , , , 30)
    CreateProg pList.ListIndex + 1, tmp2
    txtPrg.Text = tmp1
    SaveProg CurProgIndex
Next i
Exit Sub
eh:
If Err.Number = dbCWS Then Exit Sub
MsgBox Err.Description, vbCritical, "Error"
End Sub

Private Sub mnuSave_Click()
Dim File As String
Dim nmb As Long
On Error GoTo eh
File = ShowSaveDlg(Filt + 1, Me.hWnd, Purpose:=Section + "PROG")
nmb = FreeFile
Open File For Output As nmb
    Print #nmb, txtPrg.Text;
Close nmb
Exit Sub
eh:
If Err.Number = dbCWS Then Exit Sub
MsgBox Err.Description, vbCritical, "Error"
End Sub

Private Sub mnuUpgradeStorage_Click()
If dbMsgBox(1100, vbYesNo Or vbDefaultButton2) = vbYes Then 'This will erase all the programs you see. Continue anyway?`Warning
  ConvertFrom1To2
  WriteSettingsFile
  LoadSettings
End If
End Sub

Private Sub OkButton_Click()
If Not DontCompile Then
    If Not Compile Then Exit Sub
End If
SaveSettings
Restore
Me.Tag = ""
Me.Hide
End Sub

Public Sub SaveSettings()
dbSaveSetting Section, "ToolProgIndex", CStr(CurProgIndex)
SaveProg CurProgIndex
WriteSettingsFile
End Sub

Public Sub SaveProg(ByVal ProgIndex As Long)
WriteProg ProgIndex, txtPrg.Text
End Sub

Public Sub WriteProg(ByVal ProgIndex As Long, ByRef ProgText As String)
dbSaveSetting Section, "PrgToolProg" + CStr(ProgIndex), ProgText
End Sub

Public Sub LoadSettings()
ReadSettingsFile
CurProgIndex = Val(dbGetSetting(Section, "ToolProgIndex", CStr(0)))
LoadProg CurProgIndex
FreshList
End Sub

Public Sub LoadProg(ByVal ProgIndex As Long, _
                    Optional ByVal Template As Boolean = False)
Dim tmp As String
ReadProg ProgIndex, tmp, Template:=Template
txtPrg.Text = tmp
End Sub

'template: reads resid ProgIndex, reads default if =0
Public Sub ReadProg(ByVal ProgIndex As Long, _
                    ByRef ProgText As String, _
                    Optional ByVal Template As Boolean = False)
Dim TemplateResID As Long
Dim DefResID As Long
Select Case pMode
  Case epmDraw
    TemplateResID = 0
    DefResID = 2447
  Case epmTool
    TemplateResID = 2437
    DefResID = 2389
End Select
If Not Template Then
  If DefResID = 0 Then
    ProgText = dbGetSetting(Section, "PrgToolProg" + CStr(ProgIndex), "")
  Else
    ProgText = dbGetSetting(Section, "PrgToolProg" + CStr(ProgIndex), GRSF(DefResID))
  End If
Else
  If TemplateResID = 0 Then
    If ProgIndex = 0 Then
      ProgText = ""
    Else
      ProgText = GRSF(ProgIndex)
    End If
  Else
    ProgText = GRSF(IIf(ProgIndex = 0, TemplateResID, ProgIndex))
  End If
End If
End Sub

Public Sub DelProg(ByVal ProgIndex As Long)
dbDeleteSetting Section, "PrgToolProg" + CStr(ProgIndex)
End Sub

'Friend Sub GetVars(ByRef pVars() As Variable)
'pVars = Vars
'End Sub

Friend Sub GetPrg(ByRef pPrg As SMP)
pPrg = Prg
End Sub

Public Function Compile() As Boolean
Dim EV As clsEVal
Dim ErrNum As Long, ErrSrc As String, ErrDsc As String
'Dim Statements() As String
Dim i As Long, j As Long
Dim tCode As Variant
Compile = False
Set EV = New clsEVal

With Prg
  Select Case pMode
  Case epmTool
    ReDim .Vars(0 To 15)
    
    .Vars(0).Name = "X0"
    .Vars(1).Name = "Y0"
    .Vars(2).Name = "X1"
    .Vars(3).Name = "Y1"
    .Vars(4).Name = "BUTTON"
    .Vars(5).Name = "SHIFT"
    .Vars(6).Name = "CX"
    .Vars(7).Name = "CY"
    .Vars(8).Name = "W"
    .Vars(9).Name = "H"
    .Vars(10).Name = "FC"
    .Vars(11).Name = "BC"
    .Vars(12).Name = "EVENT"
    .Vars(13).Name = "PENPRESSURE"
    .Vars(14).Name = "WHEELPOS"
    .Vars(15).Name = "CANCELSCROLL"
    
  Case epmDraw
    ReDim .Vars(0 To 5)
    
    .Vars(0).Name = "CX"
    .Vars(1).Name = "CY"
    .Vars(2).Name = "W"
    .Vars(3).Name = "H"
    .Vars(4).Name = "FC"
    .Vars(5).Name = "BC"
    
  End Select
EV.MatVars .Vars
End With

j = 0
On Error GoTo eh
EV.CompileSMP txtPrg.Text, Prg
On Error GoTo 0

Compile = True
Exit Function
Resume
eh:
    ErrNum = Err.Number
    ErrSrc = Err.Source
    ErrDsc = Err.Description
    ShowStatus 1244
    dbMsgBox ErrDsc + "`Compile error", vbCritical '"Compile error in statement: index = " + CStr(lin) + "." + vbCrLf + Err.Description, vbCritical

End Function

Private Sub pList_Click()
If (pList.ListIndex <> -1) Then
  If Not eLock Then
    SaveProg CurProgIndex
  End If
  CurProgIndex = pList.ListIndex
  If Not eLock Then
    LoadProg CurProgIndex
  End If
End If
End Sub

Public Sub FreshList()
Dim sArr() As String
Dim i As Long
sArr = Split(dbGetSetting(Section, "ProgsNames", IIf(pMode = epmTool, "Soft pen", "Graph plotter")), Chr$(1))
pList.Clear
For i = 0 To UBound(sArr)
    pList.AddItem sArr(i), i
Next i
eLock = True
pList.ListIndex = CurProgIndex
eLock = False
End Sub

Public Sub FlushList()
Dim sArr() As String
Dim i As Long
ReDim sArr(0 To pList.ListCount - 1)
For i = 0 To UBound(sArr)
    sArr(i) = pList.List(i)
Next i
dbSaveSetting Section, "ProgsNames", Join(sArr, Chr$(1))
End Sub

Private Sub pList_DblClick()
OkButton_Click
End Sub

Public Sub WriteSettingsFile()
If Sgs Is Nothing Then
  Set Sgs = New clsIniFile
End If
If Len(Sgs.FileName) = 0 Then
  Sgs.DefFile "Progs"
End If
Sgs.SaveFile
End Sub

Public Sub ReadSettingsFile()
If Sgs Is Nothing Then
  Set Sgs = New clsIniFile
End If
If Len(Sgs.FileName) = 0 Then
  Sgs.DefFile "Progs"
End If
Sgs.LoadFile
End Sub

Public Function dbGetSetting(ByRef Key As String, ByRef Parameter As String, _
                             Optional ByRef DefValue As String = vbNullString) As String
Dim tmp As String
  If Sgs.QuerySetting(Key, Parameter, tmp) Then
      dbGetSetting = tmp
  Else
      dbGetSetting = DefValue
  End If
End Function

Public Sub dbDeleteSetting(ByVal Key As String, _
                           Optional ByRef Parameter As String = "")
Dim tmp As String
Dim Ret As Boolean
    If Len(Key) = 0 Then
        Set Sgs = Nothing
        Set Sgs = New clsIniFile
        Sgs.DefFile
    Else
        If Len(Parameter) = 0 Then
            Ret = Sgs.DeleteSection(Key)
        Else
            Ret = Sgs.DeleteSetting(Key, Parameter)
        End If
    End If

End Sub


Public Sub dbSaveSetting(ByRef Key As String, ByRef Parameter As String, _
                                              ByRef Value As String)
    Sgs.SetSetting Key, Parameter, Value
End Sub

Public Sub ConvertFrom1To2()
Dim sgsVals() As String
Dim sgsNames() As String
Dim n As Long, i As Long
n = mdlSettings.dbGetAllSettings(Section, sgsNames, sgsVals)
For i = 0 To n - 1
  dbSaveSetting Section, sgsNames(i), sgsVals(i)
Next i
End Sub

Private Sub pList_KeyDown(KeyCode As Integer, Shift As Integer)
Dim ShiftKeyCode As Long
ShiftKeyCode = GetShiftKeyCode(KeyCode, Shift)
Select Case ShiftKeyCode
  Case 550 'ctrl+up
    btnMoveUp_Click
    KeyCode = 0
  Case 552
    btnMoveDown_Click
    KeyCode = 0
End Select
End Sub

Public Sub SetMode(ByVal Mode As eProgMode)
Select Case Mode
  Case epmTool
    Restore
    pMode = Mode
    LoadCaptions
    LoadSettings
  Case epmDraw
    Section = "Draw"
    Filt = dbPLoad
    pMode = Mode
    LoadCaptions
    LoadSettings
End Select
End Sub
