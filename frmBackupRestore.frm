VERSION 5.00
Begin VB.Form frmBackupRestore 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Backed up"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3630
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   3630
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Refresher 
      Interval        =   1000
      Left            =   2745
      Top             =   2085
   End
   Begin VB.Timer MainFormEnabler 
      Interval        =   100
      Left            =   2265
      Top             =   2040
   End
   Begin VB.ListBox FList 
      Height          =   2790
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   225
      Width           =   1755
   End
   Begin SMBMaker.dbButton btnView 
      Height          =   510
      Left            =   1935
      TabIndex        =   5
      Top             =   645
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   900
      MouseIcon       =   "frmBackupRestore.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Others          =   $"frmBackupRestore.frx":001C
      OthersPresent   =   -1  'True
   End
   Begin SMBMaker.dbButton btnClose 
      Cancel          =   -1  'True
      Height          =   390
      Left            =   2115
      TabIndex        =   4
      Top             =   2700
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   688
      MouseIcon       =   "frmBackupRestore.frx":006B
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Others          =   $"frmBackupRestore.frx":0087
      OthersPresent   =   -1  'True
   End
   Begin SMBMaker.dbButton btnDel 
      Height          =   510
      Left            =   1920
      TabIndex        =   3
      Top             =   1155
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   900
      MouseIcon       =   "frmBackupRestore.frx":00D6
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Others          =   $"frmBackupRestore.frx":00F2
      OthersPresent   =   -1  'True
   End
   Begin SMBMaker.dbButton btnLoad 
      Height          =   510
      Left            =   1935
      TabIndex        =   2
      Top             =   135
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   900
      MouseIcon       =   "frmBackupRestore.frx":0151
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Others          =   $"frmBackupRestore.frx":016D
      OthersPresent   =   -1  'True
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Files found"
      Height          =   225
      Left            =   135
      TabIndex        =   1
      Top             =   15
      Width           =   1770
   End
End
Attribute VB_Name = "frmBackupRestore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim FileList() As String
Dim nFiles As Long

Private Sub btnClose_Click()
Unload Me
End Sub

Private Sub btnDel_Click()
Dim i As Long
On Error GoTo eh
i = FList.ListIndex
If i = -1 Then Exit Sub
If dbMsgBox(2429, vbQuestion Or vbYesNo Or vbDefaultButton2) = vbNo Then Exit Sub
Kill FileList(FList.ItemData(i))
FreshList
Exit Sub
eh:
MsgBox Err.Description, vbCritical, Err.Source
End Sub

Private Sub btnLoad_Click()
Dim i As Long
i = FList.ListIndex
If i <> -1 Then
    On Error Resume Next
    MainForm.LoadFile FileList(FList.ItemData(i))
    Me.Hide
End If
End Sub

Private Sub btnView_Click()
Dim i As Long
i = FList.ListIndex
If i <> -1 Then
    On Error GoTo eh
    Dim Data() As Long, Alpha() As Long
    vtLoadPicture Data, Alpha, FileList(FList.ItemData(i))
    ViewImage Data
End If
Exit Sub
eh:
MsgError
End Sub

Private Sub Form_Load()
FreshList
Refresher.Enabled = True
End Sub

Public Sub FreshList()
Dim i As Long
nFiles = MainForm.EnumBackUps(FileList)
If nFiles = 0 Then
    btnClose_Click
    Exit Sub
End If
FList.Clear
For i = 0 To nFiles - 1
    FList.AddItem GetFileTitle(FileList(i))
    FList.ItemData(FList.NewIndex) = i
Next i
End Sub

Private Sub Form_Paint()
On Error Resume Next
Me.PaintPicture gBackPicture, 0, 0, ScaleWidth, ScaleHeight
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Refresher.Enabled = False
End Sub

Private Sub MainFormEnabler_Timer()
'EnableWindow MainForm.hWnd, CLng(True)
End Sub

Private Sub Refresher_Timer()
Dim i As Long
On Error Resume Next
i = FList.ListIndex
On Error Resume Next
FreshList
FList.ListIndex = i
End Sub
