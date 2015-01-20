VERSION 5.00
Begin VB.Form frmKeys 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Keys & Actions"
   ClientHeight    =   5235
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7785
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   349
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   519
   StartUpPosition =   1  'CenterOwner
   Begin SMBMaker.ctlResourcizator Resr1 
      Left            =   3210
      Top             =   3045
      _ExtentX        =   1164
      _ExtentY        =   741
      ResID           =   9910
   End
   Begin SMBMaker.dbFrame dbFrame1 
      Height          =   1335
      Left            =   0
      TabIndex        =   4
      Top             =   3420
      Width           =   7755
      _ExtentX        =   13679
      _ExtentY        =   2355
      Caption         =   "Selected item:"
      EAC             =   0   'False
      Begin VB.ComboBox cmbAction 
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
         ItemData        =   "frmKeys.frx":0000
         Left            =   0
         List            =   "frmKeys.frx":0002
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   6
         ToolTipText     =   "Select the action for the selected key."
         Top             =   360
         Width           =   7665
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Key"
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
         Left            =   30
         TabIndex        =   14
         Top             =   765
         Width           =   270
      End
      Begin SMBMaker.dbButton btnChangeKey 
         Height          =   270
         Left            =   15
         TabIndex        =   13
         ToolTipText     =   "Change the key for selected shortcut."
         Top             =   990
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   476
         MouseIcon       =   "frmKeys.frx":0004
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Others          =   $"frmKeys.frx":0020
         OthersPresent   =   -1  'True
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Action"
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
         Left            =   0
         TabIndex        =   5
         Top             =   165
         Width           =   450
      End
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2790
      Left            =   0
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   240
      Width           =   7755
   End
   Begin SMBMaker.dbButton btnUp 
      Height          =   360
      Left            =   5580
      TabIndex        =   12
      ToolTipText     =   "Move selected shortcut up in the list."
      Top             =   3045
      Width           =   1080
      _ExtentX        =   1905
      _ExtentY        =   635
      MouseIcon       =   "frmKeys.frx":0075
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Others          =   $"frmKeys.frx":0091
      OthersPresent   =   -1  'True
   End
   Begin SMBMaker.dbButton btnSave 
      Height          =   450
      Left            =   1755
      TabIndex        =   11
      ToolTipText     =   "Shows Save context menu."
      Top             =   4770
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   794
      MouseIcon       =   "frmKeys.frx":00DE
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Others          =   $"frmKeys.frx":00FA
      OthersPresent   =   -1  'True
   End
   Begin SMBMaker.dbButton btnDown 
      Height          =   360
      Left            =   6660
      TabIndex        =   10
      ToolTipText     =   "Move selected shortcut down in the list."
      Top             =   3045
      Width           =   1080
      _ExtentX        =   1905
      _ExtentY        =   635
      MouseIcon       =   "frmKeys.frx":014C
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Others          =   $"frmKeys.frx":0168
      OthersPresent   =   -1  'True
   End
   Begin SMBMaker.dbButton dbButton1 
      Height          =   450
      Left            =   0
      TabIndex        =   9
      ToolTipText     =   "Shows Load context menu."
      Top             =   4770
      Width           =   1755
      _ExtentX        =   3096
      _ExtentY        =   794
      MouseIcon       =   "frmKeys.frx":01B7
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Others          =   $"frmKeys.frx":01D3
      OthersPresent   =   -1  'True
   End
   Begin SMBMaker.dbButton CancelButton 
      Cancel          =   -1  'True
      Height          =   390
      Left            =   6615
      TabIndex        =   8
      Top             =   4815
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   688
      MouseIcon       =   "frmKeys.frx":0225
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Others          =   $"frmKeys.frx":0241
      OthersPresent   =   -1  'True
   End
   Begin SMBMaker.dbButton OkButton 
      Height          =   390
      Left            =   5490
      TabIndex        =   7
      ToolTipText     =   "Shift + Enter"
      Top             =   4815
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   688
      MouseIcon       =   "frmKeys.frx":0291
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Others          =   $"frmKeys.frx":02AD
      OthersPresent   =   -1  'True
   End
   Begin SMBMaker.dbButton btnRemove 
      Height          =   360
      Left            =   1080
      TabIndex        =   3
      ToolTipText     =   "Delete selected shortcut."
      Top             =   3045
      Width           =   1080
      _ExtentX        =   1905
      _ExtentY        =   635
      MouseIcon       =   "frmKeys.frx":02F9
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Others          =   $"frmKeys.frx":0315
      OthersPresent   =   -1  'True
   End
   Begin SMBMaker.dbButton btnAdd 
      Height          =   360
      Left            =   0
      TabIndex        =   2
      ToolTipText     =   "Add a new shortcut."
      Top             =   3045
      Width           =   1080
      _ExtentX        =   1905
      _ExtentY        =   635
      MouseIcon       =   "frmKeys.frx":0366
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Others          =   $"frmKeys.frx":0382
      OthersPresent   =   -1  'True
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Key Shortcuts:"
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
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   1065
   End
   Begin VB.Menu ppMnu 
      Caption         =   "ppMnu"
      Visible         =   0   'False
      Begin VB.Menu mnuDefSMB 
         Caption         =   "SMB Maker's default"
      End
      Begin VB.Menu sep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLoadFile 
         Caption         =   "Load file..."
      End
      Begin VB.Menu Sep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPreset 
         Caption         =   "..."
         Index           =   0
      End
   End
   Begin VB.Menu ppMnuSave 
      Caption         =   "ppMnuSave"
      Visible         =   0   'False
      Begin VB.Menu mnuSavePreset 
         Caption         =   "Save to preset..."
      End
      Begin VB.Menu mnuSaveFile 
         Caption         =   "Save to file..."
      End
   End
End
Attribute VB_Name = "frmKeys"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim pkArr() As kShortcut
Const KeyDescSeparator As String = vbTab

Friend Sub SetKeys(ByRef kArr() As kShortcut)
pkArr = kArr
FreshList
End Sub

Public Sub FreshList()
Dim i As Long
Dim j As Long
j = List1.ListIndex
List1.Clear
If AryDims(AryPtr(pkArr)) = 1 Then
    For i = 0 To UBound(pkArr)
        List1.AddItem GetKeyName(pkArr(i).Key) + KeyDescSeparator + GetActionDescription(pkArr(i).Act), i
    Next i
End If
If j < List1.ListCount Then
    List1.ListIndex = j
End If
SetEn List1.ListIndex > -1
End Sub

Private Sub btnAdd_Click()
Dim nK As Long
Dim i As Long
Dim UB As Long
Dim j As Long
On Error GoTo eh1
nK = dbInKey(RaiseErrors:=True)
If AryDims(AryPtr(pkArr)) = 1 Then
    ReDim Preserve pkArr(0 To UBound(pkArr) + 1)
    UB = UBound(pkArr)
    i = List1.ListIndex
    If i = -1 Then i = UB
    For j = UB - 1 To i Step -1
        pkArr(j + 1) = pkArr(j)
    Next j
Else
    ReDim pkArr(0 To 0)
    UB = 0
    i = 0
End If
pkArr(i).Key = nK
List1.AddItem GetKeyName(pkArr(i).Key) + KeyDescSeparator + GetActionDescription(pkArr(i).Act), i
List1.ListIndex = i
Exit Sub
'eh:
If Err.Number = dbCWS Then Exit Sub
ReDim pkArr(0 To 0)
Resume Next
eh1:
MsgError
End Sub

Private Sub btnChangeKey_Click()
Dim i As Long
Dim nK As Long
On Error GoTo eh
i = List1.ListIndex
If i = -1 Or AryDims(AryPtr(pkArr)) <> 1 Then Exit Sub
nK = dbInKey
pkArr(i).Key = nK
FreshItem i
Exit Sub
eh:
MsgError
End Sub

Private Sub btnDown_Click()
Dim i1 As Long, i2 As Long
Dim t As kShortcut
i1 = List1.ListIndex
If i1 = -1 Or i1 >= List1.ListCount - 1 Or AryDims(AryPtr(pkArr)) <> 1 Then
    vtBeep
    Exit Sub
End If
i2 = i1 + 1
SwapListElemsList1 i1, i2
t = pkArr(i1)
pkArr(i1) = pkArr(i2)
pkArr(i2) = t
List1.ListIndex = i2
End Sub

Private Sub btnRemove_Click()
Dim i As Long, j As Long
Dim h As Long
j = List1.ListIndex
If j = -1 Then vtBeep: Exit Sub
If AryDims(AryPtr(pkArr)) <> 1 Then Exit Sub
For i = j To UBound(pkArr) - 1
    pkArr(i) = pkArr(i + 1)
Next i
If UBound(pkArr) = 0 Then
    Erase pkArr
Else
    ReDim Preserve pkArr(0 To UBound(pkArr) - 1)
End If
List1.RemoveItem j
SetEn List1.ListIndex > -1
'FreshList
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Dim ShiftKeyCode As Long
ShiftKeyCode = GetShiftKeyCode(KeyCode, Shift)
Select Case ShiftKeyCode
  Case 269 'Shift+Enter
    OkButton.RaiseClick
End Select
End Sub

Private Sub List1_KeyDown(KeyCode As Integer, Shift As Integer)
Dim ShiftKeyCode As Long
ShiftKeyCode = GetShiftKeyCode(KeyCode, Shift)
Select Case ShiftKeyCode
  Case 550 'Ctrl+UpArrow
    btnUp.RaiseClick
    KeyCode = 0 'to disable list item change
  Case 552 'Ctrl+DownArrow
    btnDown.RaiseClick
    KeyCode = 0 'to disable list item change
  Case 13 'Enter
    btnChangeKey.RaiseClick
    'KeyCode = 0 'to disable ok button click
  Case 46 'delete
    btnRemove.RaiseClick
End Select
End Sub

Private Sub mnuLoadFile_Click()
Dim File As String
On Error GoTo eh
File = ShowOpenDlg(dbKeysLoad, Me.hWnd, Purpose:="KEYS")
LoadKeysFromFile File, pkArr
FreshList
Exit Sub
eh:
If Err.Number = dbCWS Then
    Exit Sub
End If
MsgError
End Sub

Private Sub mnuSaveFile_Click()
Dim File As String
On Error GoTo eh
File = ShowSaveDlg(dbKeysSave, Me.hWnd, Purpose:="KEYS")
SaveKeysToFile File, pkArr
Exit Sub
eh:
If Err.Number = dbCWS Then
    Exit Sub
End If
MsgError
End Sub

Private Sub mnuSavePreset_Click()
Dim tmp As String
On Error GoTo eh
tmp = dbInputBox(10087, Default:="", CancelError:=True, MaxLen:=64)  'Name
CheckName tmp, True
If dbSettingPresent("Keyboard", tmp) Then
    If dbMsgBox(10086, vbYesNo) = vbNo Then 'Owerwrite?
        Exit Sub
    End If
End If
SaveKeysToReg pkArr, "Keyboard", tmp
FreshPresetList
Exit Sub
eh:
If Err.Number = 114 Then
    MsgError , False
ElseIf Err.Number = dbCWS Then
    Exit Sub
Else
    MsgError
End If
End Sub


Private Sub btnSave_Click()
PopupMenu ppMnuSave
End Sub

Private Function CheckName(ByRef St As String, Optional ByVal RaiseError As Boolean = False) As Boolean
If Len(St) = 0 Then
    CheckName = False
    If RaiseError Then Err.Raise 114, "CheckName", GRSF(10088)
ElseIf InStr(1, St, "\") Then
    CheckName = False
    If RaiseError Then Err.Raise 114, "CheckName", GRSF(10089) '"Restricted character present (\)"
Else
    CheckName = True
End If
End Function

Public Sub LoadPreset(ByRef PresetName As String)
LoadKeysFromReg pkArr, "Keyboard", PresetName
FreshList
End Sub

'Private Sub btnSort_Click()
'SortCombo cmbAction, 0, cmbAction.ListCount - 1
'List1_Click
'End Sub
'
Private Sub SwapListElemsList1(ByVal i1 As Long, ByVal i2 As Long)
Dim tTxt As String, tID As Long
With List1
tTxt = .List(i1)
tID = .ItemData(i1)
.List(i1) = .List(i2)
.ItemData(i1) = .ItemData(i2)
.List(i2) = tTxt
.ItemData(i2) = tID
End With
End Sub

Private Sub SwapListElems(ByVal i1 As Long, ByVal i2 As Long)
Dim tTxt As String, tID As Long
With cmbAction
tTxt = .List(i1)
tID = .ItemData(i1)
.List(i1) = .List(i2)
.ItemData(i1) = .ItemData(i2)
.List(i2) = tTxt
.ItemData(i2) = tID
End With
End Sub

Private Sub btnUp_Click()
Dim i1 As Long, i2 As Long
Dim t As kShortcut
i1 = List1.ListIndex
If i1 < 1 Or AryDims(AryPtr(pkArr)) <> 1 Then
    vtBeep
    Exit Sub
End If
i2 = i1 - 1
SwapListElemsList1 i1, i2
t = pkArr(i1)
pkArr(i1) = pkArr(i2)
pkArr(i2) = t
List1.ListIndex = i2
End Sub

Private Sub CancelButton_Click()
Me.Tag = "C"
Me.Hide
End Sub

Private Sub cmbAction_Click()
Dim i As Long
i = List1.ListIndex
If i = -1 Then Exit Sub
pkArr(i).Act = cmbAction.ItemData(cmbAction.ListIndex)
FreshItem i
End Sub

Private Sub dbButton1_Click()
PopupMenu ppMnu, 2
End Sub

Private Sub Form_Load()
Dim i As Long
Dim j As Long
Dim tmp As String
With cmbAction
    .Clear
    j = 0
    For i = 0 To LastActionNumber
        tmp = GetActionDescription(i)
        If UCase$(tmp) <> "<OBSOLETE>" Then
            .AddItem GetActionDescription(i)
            .ItemData(.NewIndex) = i
            j = j + 1
        End If
    Next i
End With
dbLoadCaptions
FreshPresetList
SetEn False
Erase pkArr
'btnSort_Click
End Sub

Public Sub FreshPresetList()
Dim sArr() As String
Dim vArr() As String
Dim i As Long, j As Long
Dim n As Long, m As Long
n = dbGetAllSettings("Keyboard", sArr, vArr) - 1
m = mnuPreset.UBound
j = 0
For i = 0 To n
    If Len(sArr(i)) = 0 Then
        n = n - 1
    Else
        vArr(j) = sArr(i)
        j = j + 1
    End If
Next i
If n >= 0 Then
    ReDim Preserve vArr(0 To n)
Else
    Erase vArr
End If
If n > m Then
    For i = m + 1 To n
        Load mnuPreset(i)
    Next i
ElseIf n < m Then
    If n > -1 Then
        For i = m To n + 1 Step -1
            Unload mnuPreset(i)
        Next i
    Else
        For i = m To n + 2 Step -1
            Unload mnuPreset(i)
        Next i
        mnuPreset(0).Visible = False
    End If
End If
For i = 0 To n
    With mnuPreset(i)
        .Caption = vArr(i)
        .Visible = True
    End With
Next i

End Sub

Private Sub Form_Paint()
On Error Resume Next
Me.PaintPicture gBackPicture, 0, 0, Me.ScaleWidth, Me.ScaleHeight
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim Answ As VbMsgBoxResult
If UnloadMode = 0 Then
    Cancel = True
    Answ = dbMsgBox(2350, vbYesNoCancel Or vbInformation) '"Apply changes?`Confirm"
    Select Case Answ
        Case vbYes
            OkButton_Click
        Case vbNo
            CancelButton_Click
        Case vbCancel
            Exit Sub
    End Select
End If
End Sub

Private Sub FreshItem(ByVal ItemIndex)
List1.List(ItemIndex) = GetKeyName(pkArr(ItemIndex).Key) + KeyDescSeparator + GetActionDescription(pkArr(ItemIndex).Act)
End Sub

Private Sub List1_Click()
Dim i As Long
i = List1.ListIndex
If i > -1 Then
    SetCmbAction pkArr(i).Act
End If
SetEn (i > -1)
End Sub

Private Sub SetEn(ByVal En As Boolean)
Label2.Enabled = En
Label3.Enabled = En
cmbAction.Enabled = En
End Sub

Private Sub SetCmbAction(ByVal nAct As dbCommands)
Dim i As Long
With cmbAction
    For i = 0 To .ListCount - 1
        If .ItemData(i) = nAct Then
            .ListIndex = i
            Exit For
        End If
    Next i
End With
End Sub

Private Sub mnuDefSMB_Click()
StringToKeys DefKeysString, pkArr
FreshList
End Sub

Private Sub mnuPreset_Click(Index As Integer)
LoadPreset mnuPreset(Index).Caption
End Sub

Private Sub OkButton_Click()
Dim i As Long
i = -1
If AryDims(AryPtr(pkArr)) = 1 Then
    i = UBound(pkArr)
End If
If i = -1 Then
    dbMsgBox 10085, vbExclamation 'Cannot quit
    Exit Sub
Else
    Me.Tag = ""
    Me.Hide
End If
End Sub

Friend Sub ExtractKeys(ByRef kArr() As kShortcut)
kArr = pkArr
End Sub

Public Sub dbLoadCaptions()
Resr1.LoadCaptions
'Me.Caption = GRSF(2339)
'Label1.Caption = GRSF(2340)
'Label2.Caption = GRSF(2343)
'Label3.Caption = GRSF(2344)
'mnuDefSMB.Caption = GRSF(2368)
End Sub

Public Sub DeletePreset(ByRef PresetName As String, Optional ByVal Prompt As Boolean = True)
If Prompt Then
    '2052="Are you sure you want to delete """
    If dbMsgBox(grs(2052, "$p", PresetName), vbYesNo Or vbQuestion Or vbDefaultButton2) = vbNo Then
        Exit Sub
    End If
End If
dbDeleteSetting "Keyboard", PresetName
End Sub
