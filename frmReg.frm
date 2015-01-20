VERSION 5.00
Begin VB.Form frmReg 
   BackColor       =   &H00E3DFE0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Extension associations"
   ClientHeight    =   5730
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6225
   HelpContextID   =   11130
   Icon            =   "frmReg.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5730
   ScaleWidth      =   6225
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin SMBMaker.ctlResourcizator Resr1 
      Left            =   4905
      Top             =   2550
      _ExtentX        =   1164
      _ExtentY        =   741
      ResID           =   9911
   End
   Begin VB.CheckBox Report 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Show report"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1005
      Left            =   4890
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "If checked, the report will be generated."
      Top             =   1350
      Width           =   960
   End
   Begin SMBMaker.dbFrame Frame1 
      Height          =   3420
      Left            =   1500
      TabIndex        =   9
      Top             =   1155
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   6033
      Caption         =   "Selected type options"
      BackColor       =   14933984
      EAC             =   0   'False
      Begin VB.TextBox txtDesc 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Text            =   "Open with SMB Maker"
         ToolTipText     =   "Type description for current extension."
         Top             =   2940
         Width           =   3015
      End
      Begin VB.TextBox CommandCaption 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Text            =   "Open with SMB Maker"
         ToolTipText     =   "The caption of the item in file right-click menu."
         Top             =   2280
         Width           =   3030
      End
      Begin VB.CheckBox ReplaceIcon 
         BackColor       =   &H00FFC0FF&
         Caption         =   "Replace icon"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Force the file icon be the selected one."
         Top             =   840
         Width           =   3015
      End
      Begin VB.CheckBox OpenDefault 
         BackColor       =   &H00FFC0FF&
         Caption         =   "Open with SMB Maker by default"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "If checked, the file with selected extension will be opened with SMB Maker on doubleclick."
         Top             =   240
         Width           =   3030
      End
      Begin VB.Label Label3 
         BackColor       =   &H00E3DFE0&
         BackStyle       =   0  'Transparent
         Caption         =   "File type description:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   2700
         Width           =   3000
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E3DFE0&
         BackStyle       =   0  'Transparent
         Caption         =   "Menu caption:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   2040
         Width           =   3030
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
      Height          =   2430
      IntegralHeight  =   0   'False
      Left            =   60
      TabIndex        =   0
      ToolTipText     =   "The list of extensions to associate with SMB Maker. "
      Top             =   1155
      Width           =   1335
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "NOTE: everything above has no relation to current file associations state."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   75
      TabIndex        =   15
      Top             =   5340
      Width           =   6015
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmReg.frx":0ABA
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   945
      Left            =   90
      TabIndex        =   12
      Top             =   90
      Width           =   6015
   End
   Begin SMBMaker.dbButton btnRemAll 
      Height          =   495
      Left            =   2445
      TabIndex        =   11
      ToolTipText     =   "Remove all extension associations."
      Top             =   4680
      Width           =   2325
      _ExtentX        =   4101
      _ExtentY        =   873
      MouseIcon       =   "frmReg.frx":0B6E
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Others          =   $"frmReg.frx":0B8A
      OthersPresent   =   -1  'True
   End
   Begin SMBMaker.dbButton btnDel 
      Height          =   375
      Left            =   75
      TabIndex        =   2
      ToolTipText     =   "Removes the selected extension from the list."
      Top             =   4200
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      MouseIcon       =   "frmReg.frx":0BDF
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Others          =   $"frmReg.frx":0BFB
      OthersPresent   =   -1  'True
   End
   Begin SMBMaker.dbButton btnAdd 
      Height          =   375
      Left            =   75
      TabIndex        =   1
      ToolTipText     =   "Adds the extension to the list."
      Top             =   3720
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      MouseIcon       =   "frmReg.frx":0C4C
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Others          =   $"frmReg.frx":0C68
      OthersPresent   =   -1  'True
   End
   Begin SMBMaker.dbButton CancelButton 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   5085
      TabIndex        =   8
      Top             =   4740
      Width           =   1005
      _ExtentX        =   1773
      _ExtentY        =   661
      MouseIcon       =   "frmReg.frx":0CB6
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Others          =   $"frmReg.frx":0CD2
      OthersPresent   =   -1  'True
   End
   Begin SMBMaker.dbButton OkButton 
      Default         =   -1  'True
      Height          =   510
      Left            =   165
      TabIndex        =   7
      Top             =   4665
      Width           =   2280
      _ExtentX        =   4022
      _ExtentY        =   900
      MouseIcon       =   "frmReg.frx":0D22
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Others          =   $"frmReg.frx":0D3E
      OthersPresent   =   -1  'True
   End
End
Attribute VB_Name = "frmReg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Assocs() As vtRegFileType
Option Explicit

Private Sub btnRemAll_Click()
On Error GoTo eh
UninstallFileTypes
dbMsgBox 2620, vbInformation 'What has been done:...
CancelButton_Click
Exit Sub
eh:
    Select Case MsgError("Failed." + vbNewLine + "Err.Description", vbExclamation Or vbAbortRetryIgnore)
        Case vbAbort
            Exit Sub
        Case vbIgnore
            Resume Next
        Case vbRetry
            Resume
    End Select
End Sub

Private Sub CommandCaption_Change()
Dim i As Long
i = List1.ListIndex
If i = -1 Then Exit Sub
Assocs(i).MenuCaption = CommandCaption.Text
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Dim ShiftKeyCode As Long
ShiftKeyCode = GetShiftKeyCode(KeyCode, Shift)
If ShiftKeyCode = 112 Then
    KeyCode = 0
    ShowHelpWindow
End If
End Sub

Private Sub Form_Paint()
Me.PaintPicture LoadResPicture(RES_BackPicture, vbResBitmap), 0, 0, Me.ScaleWidth, Me.ScaleHeight
End Sub

Private Sub RemoveItem(ByVal Index As Integer)
Dim Temp() As vtRegFileType
Dim h As Integer
Dim i As Integer
If Index < 0 Or Index > List1.ListCount - 1 Or List1.ListCount = 1 Then
    vtBeep
    Exit Sub
End If
List1.RemoveItem Index
Temp = Assocs
h = 0
ReDim Assocs(0 To UBound(Assocs) - 1)
For i = 0 To UBound(Temp)
    If Not (i = Index) Then
        Assocs(h) = Temp(i)
        h = h + 1
    End If
Next i
UpdateList
End Sub

Private Sub btnAdd_Click()
Dim tmp As String
tmp = Trim(dbInputBox(1161, ""))
If tmp = "" Then Exit Sub
If Not (Mid(tmp, 1, 1) = ".") Then
    If Mid(tmp, 1, 2) = "*." Then
        tmp = Mid(tmp, 2, Len(tmp) - 1)
    ElseIf Not CBool(InStr(1, tmp, ".")) Then
        tmp = "." + tmp
    Else
        vtBeep
        Exit Sub
    End If
End If
If Len(tmp) > 6 Or CBool(InStr(1, tmp, "\")) _
                Or CBool(InStr(1, tmp, "/")) _
                Or CBool(InStr(1, tmp, ":")) _
                Or CBool(InStr(1, tmp, "*")) _
                Or CBool(InStr(1, tmp, "?")) _
                Or CBool(InStr(1, tmp, """")) _
                Or CBool(InStr(1, tmp, ">")) _
                Or CBool(InStr(1, tmp, "<")) _
                Or CBool(InStr(1, tmp, "|")) Or CBool(InStr(2, tmp, ".")) Then
    vtBeep
    Exit Sub
End If
AddItem tmp
End Sub

Private Sub btnDel_Click()
Dim i As Integer
i = List1.ListIndex
If i < 0 Then
    vtBeep
    Exit Sub
End If
RemoveItem i
End Sub

Private Sub CancelButton_Click()
Me.Hide
End Sub

Private Sub LoadFormatExts(ByRef Format As FormatTemplate, _
                           ByRef List() As vtRegFileType)
Dim sExtList As String
Dim sDescList As String
Dim ExtList() As String
Dim DescList() As String
Dim Item As vtRegFileType
Dim i As Long
Dim n As Long
Dim BaseIndex As Long
Dim CurDesc As String

Format.GetFileTypeInfo sExtList, sDescList, Item.IconName, Item.DefaultEditor
If Len(sExtList) = 0 Then Exit Sub
Item.MenuCaption = grs(2149, "|1", App.Title)

ExtList = Split(sExtList, "|")
DescList = Split(sDescList, "|")
n = UBound(ExtList) + 1
If Len(sDescList) > 0 Then
    ReDim Preserve DescList(0 To n - 1)
Else
    ReDim DescList(0 To n - 1)
End If
    
BaseIndex = AryLen(AryPtr(List))
If BaseIndex = 0 Then
    ReDim List(0 To BaseIndex + n - 1)
Else
    ReDim Preserve List(0 To BaseIndex + n - 1)
End If
For i = 0 To n - 1
    If Len(DescList(i)) > 0 Then CurDesc = DescList(i)
    List(BaseIndex + i) = Item
    With List(BaseIndex + i)
        .Extension = "." + ExtList(i)
        .FileDescription = CurDesc
        If .IconName = "" Then .IconName = "-8"
        .IconName = ExePath + "," + .IconName
    End With
Next i

End Sub

Private Sub Form_Load()
Dim i As Long
Const Icon_Standard = 23
Const Icon_Pal = 10
Const Icon_Brh = 3

Resr1.LoadCaptions

Erase Assocs
ConnectFormats
For i = 0 To nFormats - 1
    LoadFormatExts FormatList(i), Assocs
Next i

'ReDim Assocs(0 To 8)
'With Assocs(0)
'    .Extension = ".smb"
'    .MenuCaption = grs(2149, "|1", App.Title)
'    .DefaultEditor = True
'    .ReplaceIcon = True
'    .IconName = ExePath + "," + CStr(Icon_Standard)
'End With
'With Assocs(1)
'    .Extension = ".bmp"
'    .MenuCaption = grs(2149, "|1", App.Title)
'    .DefaultEditor = False
'    .ReplaceIcon = False
'    .IconName = ExePath + "," + CStr(Icon_Standard)
'End With
'With Assocs(2)
'    .Extension = ".ico"
'    .MenuCaption = grs(2149, "|1", App.Title)
'    .DefaultEditor = True
'    .ReplaceIcon = True
'    .IconName = "%1"
'End With
'With Assocs(3)
'    .Extension = ".cur"
'    .MenuCaption = grs(2149, "|1", App.Title)
'    .DefaultEditor = True
'    .ReplaceIcon = True
'    .IconName = "%1"
'End With
'With Assocs(4)
'    .Extension = ".jpg"
'    .MenuCaption = grs(2149, "|1", App.Title)
'    .DefaultEditor = False
'    .ReplaceIcon = False
'    .IconName = ExePath + "," + CStr(Icon_Standard)
'End With
'With Assocs(5)
'    .Extension = ".gif"
'    .MenuCaption = grs(2149, "|1", App.Title)
'    .DefaultEditor = False
'    .ReplaceIcon = False
'    .IconName = ExePath + "," + CStr(Icon_Standard)
'End With
'With Assocs(6)
'    .Extension = ".pal"
'    .MenuCaption = grs(2149, "|1", App.Title)
'    .DefaultEditor = True
'    .ReplaceIcon = True
'    .IconName = ExePath + "," + CStr(Icon_Pal)
'End With
'With Assocs(7)
'    .Extension = ".brh"
'    .MenuCaption = grs(2149, "|1", App.Title)
'    .DefaultEditor = True
'    .ReplaceIcon = True
'    .IconName = ExePath + "," + CStr(Icon_Brh)
'End With
'With Assocs(8)
'    .Extension = ".png"
'    .MenuCaption = grs(2149, "|1", App.Title)
'    .DefaultEditor = False
'    .ReplaceIcon = False
'    .IconName = ExePath + "," + CStr(Icon_Standard)
'End With
UpdateList
End Sub

Private Sub UpdateList()
Dim i As Integer
With List1
    .Clear
    '.ListCount = UBound(Assocs)
    For i = 0 To UBound(Assocs)
        .AddItem Assocs(i).Extension, i
    Next i
    .ListIndex = 0
    'List1_Click
End With
End Sub

Private Sub AddItem(ByVal ext As String)
Dim i As Integer
For i = 0 To List1.ListCount - 1
    If ext = List1.List(i) Then
        vtBeep
        Exit Sub
    End If
Next i
List1.AddItem ext
ReDim Preserve Assocs(0 To UBound(Assocs) + 1)
With Assocs(UBound(Assocs))
    .Extension = ext
    .MenuCaption = grs(2149, "|1", App.Title)
    .DefaultEditor = False
    .ReplaceIcon = False
    .IconName = ExePath + ",-8"
End With
UpdateList
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If UnloadMode = 0 Then
    Cancel = True
    Me.Hide
End If
End Sub

Private Sub List1_Click()
Dim i As Integer
i = List1.ListIndex
If i = -1 Then Exit Sub
With Assocs(i)
    CommandCaption.Text = .MenuCaption
    ReplaceIcon.Value = Abs(.ReplaceIcon)
    OpenDefault.Value = Abs(.DefaultEditor)
    txtDesc.Text = .FileDescription
End With
End Sub

Private Sub OkButton_Click()
Dim i As Integer, l As String
l = ""
For i = 0 To UBound(Assocs)
    SetFileType Assocs(i), l
Next i
On Error GoTo eh
RegisterEXE
dbMsgBox 2619, vbInformation 'What has been done:...
rsm:
If Report.Value = 1 Then
    dbLongMsgBox l, "Log"
End If

Me.Hide
Exit Sub
eh:
If MsgError(GRSF(1103, RaiseErrors:=False), vbYesNo) = vbYes Then
  Resume rsm2
Else
  Resume rsm
End If
Exit Sub
rsm2:
MainForm.ShowFile ExePath
MainForm.dbEnd
End Sub

Private Sub DefaultEditor_Click()
Dim i As Integer
i = List1.ListIndex
If i = -1 Then Exit Sub
Assocs(i).DefaultEditor = (OpenDefault.Value = 1)
End Sub

Private Sub OpenDefault_Click()
Dim i As Integer
i = List1.ListIndex
If i = -1 Then Exit Sub
Assocs(i).DefaultEditor = (OpenDefault.Value = vbChecked)
End Sub

Private Sub ReplaceIcon_Click()
Dim i As Integer
i = List1.ListIndex
If i = -1 Then Exit Sub
Assocs(i).ReplaceIcon = (ReplaceIcon.Value = vbChecked)
'btnSelect.Enabled = Assocs(i).ReplaceIcon
End Sub

Private Sub txtDesc_Change()
Dim i As Long
i = List1.ListIndex
If i = -1 Then Exit Sub
Assocs(i).FileDescription = txtDesc.Text
End Sub
