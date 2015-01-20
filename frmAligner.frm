VERSION 5.00
Begin VB.Form frmAligner 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Offset"
   ClientHeight    =   2775
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8190
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   204
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2775
   ScaleWidth      =   8190
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox cmbDestY 
      Height          =   315
      ItemData        =   "frmAligner.frx":0000
      Left            =   5580
      List            =   "frmAligner.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   1665
      Width           =   2310
   End
   Begin VB.ComboBox cmbDestX 
      Height          =   315
      ItemData        =   "frmAligner.frx":0004
      Left            =   5580
      List            =   "frmAligner.frx":0006
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   900
      Width           =   2310
   End
   Begin VB.ComboBox cmbBaseY 
      Height          =   315
      ItemData        =   "frmAligner.frx":0008
      Left            =   2745
      List            =   "frmAligner.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1665
      Width           =   2310
   End
   Begin VB.ComboBox cmbBaseX 
      Height          =   315
      ItemData        =   "frmAligner.frx":000C
      Left            =   2745
      List            =   "frmAligner.frx":000E
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   900
      Width           =   2310
   End
   Begin SMBMaker.ctlNumBox nmbDX 
      Height          =   555
      Left            =   570
      TabIndex        =   0
      Top             =   780
      Width           =   1560
      _ExtentX        =   2752
      _ExtentY        =   979
      Min             =   -10000
      Max             =   10000
      NumType         =   5
      HorzMode        =   0   'False
      EditName        =   "$2581"
      NLn             =   0
   End
   Begin SMBMaker.ctlNumBox nmbDY 
      Height          =   555
      Left            =   570
      TabIndex        =   1
      Top             =   1545
      Width           =   1560
      _ExtentX        =   2752
      _ExtentY        =   979
      Min             =   -10000
      Max             =   10000
      NumType         =   5
      HorzMode        =   0   'False
      EditName        =   "$2582"
      NLn             =   0
   End
   Begin SMBMaker.dbButton CancelButton 
      Cancel          =   -1  'True
      Height          =   420
      Left            =   4118
      TabIndex        =   7
      Top             =   2310
      Width           =   2235
      _ExtentX        =   3942
      _ExtentY        =   741
      MouseIcon       =   "frmAligner.frx":0010
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Others          =   $"frmAligner.frx":002C
      OthersPresent   =   -1  'True
   End
   Begin SMBMaker.dbButton OkButton 
      Default         =   -1  'True
      Height          =   420
      Left            =   1838
      TabIndex        =   6
      Top             =   2310
      Width           =   2235
      _ExtentX        =   3942
      _ExtentY        =   741
      MouseIcon       =   "frmAligner.frx":007C
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Others          =   $"frmAligner.frx":0098
      OthersPresent   =   -1  'True
   End
   Begin VB.Line Line4 
      X1              =   135
      X2              =   8115
      Y1              =   660
      Y2              =   660
   End
   Begin VB.Line Line3 
      X1              =   135
      X2              =   8115
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Line Line2 
      X1              =   5280
      X2              =   5280
      Y1              =   105
      Y2              =   2040
   End
   Begin VB.Line Line1 
      X1              =   2400
      X2              =   2400
      Y1              =   105
      Y2              =   2040
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Image anchor (hot-spot)"
      Height          =   480
      Left            =   5580
      TabIndex        =   12
      Top             =   75
      Width           =   2310
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Base anchor (coordinate zero)"
      Height          =   480
      Left            =   2745
      TabIndex        =   11
      Top             =   75
      Width           =   2310
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Offset"
      Height          =   405
      Left            =   570
      TabIndex        =   10
      Top             =   75
      Width           =   1560
   End
   Begin VB.Label lblY 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Y:"
      Height          =   195
      Left            =   120
      TabIndex        =   9
      Top             =   1725
      Width           =   150
   End
   Begin VB.Label lblX 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "X:"
      Height          =   195
      Left            =   120
      TabIndex        =   8
      Top             =   960
      Width           =   150
   End
End
Attribute VB_Name = "frmAligner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim pDestPointSupported As Boolean
Dim pBasePointSupported As Boolean

Public Sub SetSupport(ByVal Base As Boolean, _
                      ByVal Dest As Boolean)
pDestPointSupported = Dest
pBasePointSupported = Base
LoadCaptions
End Sub

Public Sub SetProps(ByVal BaseAncX As OrgAnchor, _
                    ByVal BaseAncY As OrgAnchor, _
                    ByVal DestAncX As OrgAnchor, _
                    ByVal DestAncY As OrgAnchor, _
                    ByVal dx As Long, _
                    ByVal dy As Long)
SetComboItemByData cmbBaseX, BaseAncX
SetComboItemByData cmbBaseY, BaseAncY
SetComboItemByData cmbDestX, DestAncX
SetComboItemByData cmbDestY, DestAncY
On Error Resume Next
nmbDX.Value = dx
nmbDY.Value = dy
On Error GoTo 0
End Sub

Public Sub GetProps(ByRef BaseAncX As OrgAnchor, _
                    ByRef BaseAncY As OrgAnchor, _
                    ByRef DestAncX As OrgAnchor, _
                    ByRef DestAncY As OrgAnchor, _
                    ByRef dx As Long, _
                    ByRef dy As Long)
BaseAncX = ComboSelItemData(cmbBaseX)
BaseAncY = ComboSelItemData(cmbBaseY)
DestAncX = ComboSelItemData(cmbDestX)
DestAncY = ComboSelItemData(cmbDestY)
dx = nmbDX.Value
dy = nmbDY.Value
End Sub

Friend Function ComboSelItemData(ByRef Combo As ComboBox, _
                                 Optional ByVal Def As Long) As Long
If Combo.ListIndex = -1 Then
    ComboSelItemData = Def
Else
    ComboSelItemData = Combo.ItemData(Combo.ListIndex)
End If
End Function

Friend Sub SetComboItemByData(ByRef Combo As ComboBox, _
                              ByVal ItemData As Long, _
                              Optional ByVal NotFoundIndex As Long = 0)
Dim i As Long
If Combo.ListCount = 0 Then Exit Sub
For i = 0 To Combo.ListCount - 1
    If Combo.ItemData(i) = ItemData Then
        Exit For
    End If
Next i
If i = Combo.ListCount Then
    If NotFoundIndex >= 0 And NotFoundIndex < Combo.ListCount Then
        Combo.ListIndex = NotFoundIndex
    Else
        Combo.ListIndex = 0
    End If
Else
    Combo.ListIndex = i
End If
End Sub

Private Sub CancelButton_Click()
Me.Tag = "C"
Me.Hide
End Sub

Private Sub Form_Load()
LoadCaptions
End Sub

Public Sub LoadCaptions()
FillCombo cmbBaseX, 2482, pBasePointSupported
FillCombo cmbBaseY, 2492, pBasePointSupported
FillCombo cmbDestX, 2482, pDestPointSupported
FillCombo cmbDestY, 2492, pDestPointSupported
Label3.Caption = GRSF(2575)
Label4.Caption = GRSF(2576)
Label5.Caption = GRSF(2577)
Me.Caption = GRSF(2580)
End Sub

Private Sub FillCombo(ByRef Combo As ComboBox, _
                      ByVal BaseResID As Long, _
                      ByVal PointSupported As Boolean)
Dim tmp As String
Dim ID As Long
Dim Pos As Long
Dim ItemData As Long
On Error GoTo eh
Combo.Clear
ID = BaseResID
Do
    On Error Resume Next
    tmp = ""
    tmp = LoadResString(ID)
    On Error GoTo 0
    If Len(tmp) = 0 Then
        Exit Do
    End If
    Pos = InStr(2, tmp, "|")
    If Pos = 0 Then Exit Do
    ItemData = Val(Mid$(tmp, 1, Pos - 1))
    If ItemData <> 3 Or PointSupported Then
        Combo.AddItem Mid$(tmp, Pos + 1)
        Combo.ItemData(Combo.NewIndex) = ItemData
    End If
    ID = ID + 1
Loop Until tmp = "<EOL>"
Exit Sub
eh:
MsgError
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

Private Sub OkButton_Click()
Me.Tag = ""
Me.Hide
End Sub
