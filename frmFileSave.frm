VERSION 5.00
Begin VB.Form frmFileSave 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Save picture"
   ClientHeight    =   4917
   ClientLeft      =   44
   ClientTop       =   363
   ClientWidth     =   7183
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.07
      Charset         =   204
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmFileSave.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   447
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   653
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin SMBMaker.ctlResourcizator Resr1 
      Left            =   4935
      Top             =   3300
      _extentx        =   1097
      _extenty        =   671
      resid           =   9917
   End
   Begin VB.ComboBox cmbFile 
      Height          =   315
      Left            =   195
      TabIndex        =   13
      Top             =   285
      Width           =   6795
   End
   Begin SMBMaker.dbFrame dbFrame1 
      Height          =   1245
      Left            =   5160
      TabIndex        =   10
      Top             =   1110
      Width           =   1905
      _ExtentX        =   3353
      _ExtentY        =   2195
      EAC             =   0   'False
      Begin SMBMaker.dbButton btnViewAlpha 
         Height          =   285
         Left            =   292
         TabIndex        =   2
         Top             =   285
         Width           =   1260
         _ExtentX        =   2215
         _ExtentY        =   508
         MouseIcon       =   "frmFileSave.frx":000C
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.064
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Others          =   $"frmFileSave.frx":0028
         OthersPresent   =   -1  'True
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Alpha channel"
         Height          =   285
         Left            =   0
         TabIndex        =   12
         Top             =   0
         Width           =   1845
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "(alpha status)"
         Height          =   495
         Left            =   195
         TabIndex        =   11
         Top             =   600
         Width           =   1455
         WordWrap        =   -1  'True
      End
   End
   Begin VB.ListBox List1 
      Height          =   605
      Left            =   135
      TabIndex        =   0
      Top             =   1365
      Width           =   4725
   End
   Begin SMBMaker.ctlTaggedText hlpFormat 
      Height          =   1530
      Left            =   150
      TabIndex        =   7
      Top             =   2655
      Width           =   3810
      _extentx        =   6726
      _extenty        =   2703
      forecolor       =   -2147483630
      backcolor       =   12632256
   End
   Begin SMBMaker.dbButton CancelButton 
      Cancel          =   -1  'True
      Height          =   270
      Left            =   5625
      TabIndex        =   6
      Top             =   4470
      Width           =   1305
      _ExtentX        =   2296
      _ExtentY        =   467
      MouseIcon       =   "frmFileSave.frx":0077
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.064
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Others          =   $"frmFileSave.frx":0093
      OthersPresent   =   -1  'True
   End
   Begin SMBMaker.dbButton OkButton 
      Default         =   -1  'True
      Height          =   360
      Left            =   105
      TabIndex        =   5
      Top             =   4455
      Width           =   2070
      _ExtentX        =   3658
      _ExtentY        =   630
      MouseIcon       =   "frmFileSave.frx":00E3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   11.52
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Others          =   $"frmFileSave.frx":00FF
      OthersPresent   =   -1  'True
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Selected format description"
      Height          =   210
      Left            =   150
      TabIndex        =   9
      Top             =   2430
      Width           =   3810
   End
   Begin SMBMaker.dbButton btnCustomize 
      Height          =   330
      Left            =   3630
      TabIndex        =   1
      Top             =   2025
      Width           =   1215
      _ExtentX        =   2134
      _ExtentY        =   589
      MouseIcon       =   "frmFileSave.frx":014D
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.064
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Others          =   $"frmFileSave.frx":0169
      OthersPresent   =   -1  'True
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Format"
      Height          =   225
      Left            =   210
      TabIndex        =   8
      Top             =   1140
      Width           =   4560
   End
   Begin SMBMaker.dbButton btnBrowse 
      Height          =   285
      Left            =   6008
      TabIndex        =   3
      Top             =   645
      Width           =   1080
      _ExtentX        =   1910
      _ExtentY        =   508
      MouseIcon       =   "frmFileSave.frx":01C0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.064
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Others          =   $"frmFileSave.frx":01DC
      OthersPresent   =   -1  'True
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "File name"
      Height          =   225
      Left            =   173
      TabIndex        =   4
      Top             =   60
      Width           =   6915
   End
End
Attribute VB_Name = "frmFileSave"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const RestrictedChars As String = "/\:*?""<>|" + vbNewLine + vbTab
Const Number_Of_Paths = 10

Dim IDs() As String
Dim Exts() As String
Dim Sgs() As Boolean

Dim Alpha() As Long
Dim ptrData As Long

Dim tmpFN As String

Public SavedFileName As String
Public SavedFormatID As String
Dim Purpose As String

Public Sub SetFormatID(ByRef ID As String)
Dim iFmt As Long
Dim tID As String
Dim i As Long
For i = 0 To List1.ListCount - 1
    iFmt = List1.ItemData(i)
    FormatList(iFmt).GetInfo tID, "", False
    If UCase$(ID) = UCase$(tID) Then
        List1.ListIndex = i
        Exit For
    End If
Next i
End Sub

Public Sub AlphaExchange(ByRef DataAlpha() As Long)
SwapArys AryPtr(DataAlpha), AryPtr(Alpha)
Update
End Sub

Public Sub SetPtrData(ByVal aptrData As Long)
ptrData = aptrData
End Sub

Public Sub RemovePtrData()
ptrData = 0
End Sub

Private Sub btnBrowse_Click()
Dim Pos As String
Dim Fld As String
Dim Pieces() As String
Dim i As Long
On Error Resume Next
MakeFullPath
'CreateFolder GetDirName(txtFile.Text)
On Error GoTo eh
Fld = GetDirName(txtFile)
If Len(Fld) > 0 Then
    If Not FolderExists(Fld) Then
        Pieces = Split(Fld, "\")
        Fld = ""
        For i = 0 To UBound(Pieces) - 1
            If FolderExists(Fld + Pieces(i) + "\") Then
                Fld = Fld + Pieces(i) + "\"
            Else
                Exit For
            End If
        Next i
    End If
End If
With CDl
    .InitDir = Fld
'    If Not FolderExists(.InitDir) Then
'        .InitDir = ""
'    End If
    .OpenFlags = cdlOFNHideReadOnly Or cdlOFNNoReadOnlyReturn
    .DialogTitle = ""
    .Filter = GRSF(2603)
    .FileName = GetFileTitle(txtFile.Text)
    On Error Resume Next
    .Filter = FormatList(SelectedFormatIndex).GetFilter + "|" + .Filter
    On Error GoTo eh
    .hWndOwner = Me.hWnd
    
    .ShowSave
    txtFile.Text = .FileName
    SaveSMBCurDir txtFile.Text, Purpose:=Purpose
End With
Exit Sub

eh:
MsgError
End Sub



Private Sub btnViewAlpha_Click()
On Error GoTo eh
ViewImage Alpha, "MP"
Update
Exit Sub
eh:
MsgError
Update
End Sub

Private Sub btnCustomize_Click()
Dim Selected As Boolean
On Error GoTo eh
With FormatList(SelectedFormatIndex)
    Selected = True
    .SetPtrData ptrData
    .Customize
    .RemovePtrData
End With
Update

Exit Sub
eh:
PushError
    If Selected Then
        FormatList(SelectedFormatIndex).RemovePtrData
    End If
PopError
MsgError
Update
End Sub

Private Sub CancelButton_Click()
Me.Tag = "C"
Me.Hide
End Sub

Public Sub FillFileCombo()
Dim n As Long
Dim Pth As String
Dim i As Long
With cmbFile
    .Clear
    n = 0
    For i = 0 To Number_Of_Paths - 1
        Pth = dbGetSetting("SaveDialog", "Path#" + VedNull(i, 2))
        If Len(Pth) > 0 Then
            .AddItem Pth
            .ItemData(.NewIndex) = 1
        End If
    Next i
    .AddItem "-", .ListCount
    .AddItem "<Add item...>", .ListCount
    .ItemData(.NewIndex) = -1
End With
End Sub

Private Sub cmbFile_Change()
tmpFN = txtFile.Text
End Sub

Private Sub cmbFile_Click()
Dim i As Long
Dim tmp As String
i = cmbFile.ListIndex
If i < 0 Then
    tmpFN = txtFile.Text
    Exit Sub
End If
Select Case cmbFile.ItemData(i)
    Case 0
        tmp = tmpFN
        'cmbFile.ListIndex = -1
        cmbFile.Refresh
        cmbFile.Text = tmp
    Case 1
        cmbFile_Change
    Case -1
        txtFile.Text = tmpFN
        On Error GoTo eh
        txtFile.Text = dbInputBox(2630, tmpFN, CancelError:=True) '"The path please!`Add item to the list"
        AddPathToList
        
End Select
Exit Sub
eh:
MsgError
End Sub

Public Sub AddPathToList()
Dim FN As String
Dim i As Long
FN = txtFile.Text
If Len(FN) <= 0 Then Exit Sub
For i = Number_Of_Paths - 1 To 1 Step -1
    dbSaveSetting "SaveDialog", "Path#" + VedNull(i, 2), dbGetSetting("SaveDialog", "Path#" + VedNull(i - 1, 2))
Next i
dbSaveSetting "SaveDialog", "Path#" + VedNull(0, 2), FN
FillFileCombo
txtFile.Text = FN
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Dim ShiftKeyCode As Long
ShiftKeyCode = GetShiftKeyCode(KeyCode, Shift)
Select Case ShiftKeyCode
    Case GetShiftKeyCode(vbKeyB, vbAltMask)
        btnBrowse.RaiseClick
End Select
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
    Case vbKeyEscape, vbKeyReturn, Asc(vbTab), vbKeySpace
    Case Else
        If Me.ActiveControl.Name <> txtFile.Name Then
            On Error Resume Next
            txtFile.SetFocus
            'here, a key redirect to txtFile must stand
'            Dim tSS As Long
'            tSS = txtFile.SelStart
'            txtFile.SelLength = 0
'            txtFile.Text = Mid$(txtFile.Text, 1, tSS) + _
'                           Chr$(KeyAscii) + _
'                           Mid$(txtFile.Text, tSS + 1)
'            txtFile.SetFocus
'            txtFile.SelStart = Len(txtFile.Text)
'            txtFile.SelLength = 0
'            txtFile.SelText = Chr$(KeyAscii)
'            txtFile.SelLength = 0
'            txtFile.SelStart = txtFile.SelStart + 1
'            'SendKeys Chr$(KeyAscii)
            On Error GoTo 0
        End If
End Select
End Sub

Private Sub Form_Load()
On Error Resume Next
ConnectFormats
Resr1.LoadCaptions
hlpFormat.DisableHist = True
FillList
FillFileCombo
List1.ListIndex = -1
Update
txtFile.Text = GetSMBCurDir
End Sub

Private Sub Form_Paint()
Me.PaintPicture gBackPicture, 0, 0, ScaleWidth, ScaleHeight
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If UnloadMode = 0 Then
    Cancel = True
    CancelButton_Click
End If
End Sub

Private Sub List1_Click()
On Error Resume Next
UpdateExt ReplaceExisting:=True
Update
End Sub

'returns True if extension is bad
Private Function UpdateExt(Optional ByVal ReplaceExisting As Boolean = False, _
                           Optional ByVal AddIfMissing As Boolean = True) As Boolean
Dim File As String
Dim OnlyFile As String
Dim Pos As Long
Dim EnableMessage As Boolean
Dim aExts() As String
Dim i As Long
Dim ext As String
On Error GoTo eh
MakeFullPath
EnableMessage = True
File = txtFile.Text
Pos = InStrRev(File, "\")
If Pos < 1 Then
    Err.Raise 1212, , "MakeFullPath malfunctioned (internal error)!"
End If
OnlyFile = Mid$(File, Pos + 1)
Pos = InStrRev(OnlyFile, ".")
If Len(Exts(SelectedFormatIndex)) = 0 Then Exit Function
aExts = Split(Exts(SelectedFormatIndex), "|")
UpdateExt = True
If Pos < 1 And AddIfMissing Then
    File = File + "." + aExts(0)
    UpdateExt = False
Else
    ext = Mid$(OnlyFile, Pos + 1)
    For i = 0 To UBound(aExts)
        If UCase$(ext) = UCase$(aExts(i)) Then Exit For
    Next i
    If i = UBound(aExts) + 1 Then  'ext not found
        If ReplaceExisting Then
            Pos = InStrRev(File, ".")
            File = Left$(File, Pos) + aExts(0)
            UpdateExt = False
        End If
    Else
        UpdateExt = False
    End If
End If
txtFile.Text = File
Exit Function
eh:
'If EnableMessage Then MsgError
ErrRaise
End Function

Private Function SelectedFormatIndex() As Long
Dim i As Long
i = List1.ListIndex
If i < 0 Then
    Err.Raise dbCWS
Else
    SelectedFormatIndex = List1.ItemData(i)
End If
End Function

Private Sub OkButton_Click()
On Error GoTo eh
If Not OkButton.Enabled Then Exit Sub
If UpdateExt Then
    Select Case dbMsgBox(2629, vbYesNoCancel) '"The extension is not usual for this format. It may cause the file to be opened incorrectly in some programs. Are you sure you want to keep the extension?"
        Case vbYes
            'do nothing
        Case vbNo
            UpdateExt ReplaceExisting:=True
        Case vbCancel
            Err.Raise dbCWS
    End Select
End If

If Not FileFolderExists(txtFile.Text) Then
    Select Case dbMsgBox(2628, vbQuestion Or vbYesNo) '"The directory does not exist. Do you want it to be created?"
        Case vbNo
            Err.Raise dbCWS
        Case vbYes
            CreateFolder GetDirName(txtFile.Text)
    End Select
Else
  'check if exists
  If FileExists(txtFile.Text) Then
    If dbMsgBox(grs(1105, "%FN%", txtFile.Text), vbExclamation Or vbYesNo) = vbNo Then 'overwrite?
      Err.Raise dbCWS
    End If
  End If
End If
'StartWrite txtFile.Text
dbSave txtFile.Text, ptrData, Alpha, FormatList(SelectedFormatIndex)
SavedFileName = txtFile.Text
FormatList(SelectedFormatIndex).GetInfo SavedFormatID, "", False
FormatList(SelectedFormatIndex).SaveSettings

SaveSMBCurDir txtFile.Text, Purpose:=Purpose
Me.Tag = ""
Me.Hide
Exit Sub
eh:
MsgError Assertion:=Err.Number = 1212
End Sub

Public Sub MakeFullPath()
txtFile.Text = SimplifyFileName(txtFile)
End Sub

Public Function SimplifyFileName(ByRef FileName As String, _
                                 Optional ByVal FileNameRequired As Boolean = True) As String
Dim Pth As String
Dim Pieces() As String
Dim i As Long
Dim smbCurDir As String

smbCurDir = GetSMBCurDir(Purpose)

Pieces = Split(ValFolder(smbCurDir) + FileName, "\")

For i = 0 To UBound(Pieces)
    AddToPath Pth, Pieces(i)
Next i

If FileNameRequired Then
    If InStr(Pth, "\") <= 0 Then
        Err.Raise 1212, "SimplifyFileName", GRSF(2627) ' "Bad file name!"
    End If
    If Pieces(UBound(Pieces)) = "." Or Pieces(UBound(Pieces)) = ".." Then
        Err.Raise 1212, "SimplifyFileName", GRSF(2627) '"Bad file name!"
    End If
End If

SimplifyFileName = Pth
End Function

Private Sub AddToPath(ByRef Pth As String, ByRef DirToAdd As String)
Dim Pos As Long
If Len(DirToAdd) > 0 Then
    If Right$(DirToAdd, 1) = ":" Then
        Pth = DirToAdd
    Else
        If DirToAdd = "." Then
            'do nothing
        ElseIf DirToAdd = ".." Then
            Pos = InStrRev(Pth, "\")
            If Pos > 0 Then
                Pth = Left$(Pth, Pos - 1)
            End If
        Else
            If Trim$(DirToAdd) <> DirToAdd Then
                Err.Raise 1212, "AddToPath", GRSF(2625) '"The file or directory name must not begin or end with space character."
            End If
            If HasRestrictedChars(DirToAdd) Then
                Err.Raise 1212, "AddToPath", grs(2626, "$rc$", Replace(Replace(Replace(RestrictedChars, vbCr, ""), vbLf, ""), vbTab, "")) '"Directory name cannot contain the following characters: <" + Replace(Replace(Replace(RestrictedChars, vbCr, ""), vbLf, ""), vbTab, "") + ">, tabs and newlines."
            End If
            Pth = Pth + "\" + DirToAdd
        End If
    End If
Else
    Pth = Left$(Pth, InStr(Pth, ":"))
End If
End Sub

Public Function HasRestrictedChars(ByRef St As String) As Boolean
Dim Chs(0 To 255) As Boolean
Dim i As Long
Dim Res As Boolean
Dim Bytes() As Byte
Chs(0) = True
Chs(10) = True
Chs(13) = True
For i = 1 To Len(RestrictedChars)
    Chs(Asc(Mid$(RestrictedChars, i, 1))) = True
Next i
Bytes = StrConv(St, vbFromUnicode)
For i = 0 To Len(St) - 1
    Res = Res Or Chs(Bytes(i))
Next i
HasRestrictedChars = Res
End Function

Private Sub cmbFile_DblClick()
'SelTextInTextBox txtFile
cmbFile.SelStart = 0
cmbFile.SelLength = Len(cmbFile.Text)
End Sub

Private Sub cmbFile_GotFocus()
Dim Pos As Long
Pos = InStrRev(txtFile.Text, "\")
txtFile.SelStart = Pos
txtFile.SelLength = Len(txtFile.Text) - Pos
End Sub

Private Sub cmbFile_LostFocus()
On Error Resume Next
MakeFullPath
End Sub

Public Sub FillList()
Dim Txt As String
Dim i As Long
Dim CanSave As Boolean

With List1
    .Clear
    If nFormats > 0 Then
        ReDim IDs(0 To nFormats - 1)
        ReDim Exts(0 To nFormats - 1)
        ReDim Sgs(0 To nFormats - 1)
        For i = 0 To nFormats - 1
            FormatList(i).GetInfo IDs(i), Txt, Sgs(i), CanSave
            If CanSave Then
                ExtractExtsFromFilter FormatList(i).GetFilter, Exts(i)
                List1.AddItem Txt
                List1.ItemData(List1.NewIndex) = i
            End If
        Next i
    End If
End With
End Sub

Public Sub Update()
Dim i As Long
Dim ResID As Long
Dim Reason As eBadSettings
Dim FormatHasSettings As Boolean
i = List1.ListIndex
If i = -1 Then
    hlpFormat.SetResID 2601, PutHistory:=False 'Please select the format for saving. It's description will be displayed here.
    OkButton.Enabled = False
    dbFrame1.EnableAllControls True
    btnCustomize.Enabled = False
    btnCustomize.ToolTipText = GRSF(2623) '"First select the format."
Else
    ResID = 2602 'Sorry. The description is missing or the format module does not work.
    On Error Resume Next
    i = List1.ItemData(i)
    ResID = FormatList(i).GetDescriptionResID
    If ResID = 0 Then ResID = 2602
    hlpFormat.SetResID ResID, PutHistory:=False
    
    FormatList(i).GetInfo "", "", FormatHasSettings
    btnCustomize.Enabled = FormatHasSettings
    If Not FormatHasSettings Then btnCustomize.ToolTipText = GRSF(2624) '"The format is not customizable."
    
    dbFrame1.EnableAllControls FormatList(i).AlphaSupproted
    
    OkButton.Enabled = FormatList(i).CanSave(AryDims(AryPtr(Alpha)) = 2, Reason)
End If

If AryDims(AryPtr(Alpha)) = 2 Then
    Label5.Caption = GRSF(2622) '"(present)"
Else
    If Reason = bsAlphaRequired Then Label5.ForeColor = vbRed Else Label5.ForeColor = vbBlack
    Label5.Caption = GRSF(2621) '"(absent)"
End If

If Reason = bsSettingsInvalid Then
    btnCustomize.ForeColor = vbRed
Else
    btnCustomize.ForeColor = vbBlack
End If
End Sub

Public Sub SetFileName(ByRef FileName As String)
txtFile.Text = FileName
If Len(txtFile.Text) = 0 Then txtFile.Text = GetSMBCurDir(Purpose:=Purpose)
End Sub

Public Property Get txtFile() As Object
Set txtFile = cmbFile
End Property

Public Property Set txtFile(ByRef Obj As Object)
'Set cmbFile = Obj
End Property

Public Sub SetPurpose(aPurpose As String)
Purpose = aPurpose
End Sub
