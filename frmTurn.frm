VERSION 5.00
Begin VB.Form frmTurn 
   BackColor       =   &H00E3DFE0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Flip/Rotate"
   ClientHeight    =   3045
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   3885
   ControlBox      =   0   'False
   Icon            =   "frmTurn.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3045
   ScaleWidth      =   3885
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture3 
      AutoSize        =   -1  'True
      Height          =   1020
      Left            =   2730
      Picture         =   "frmTurn.frx":0ABA
      ScaleHeight     =   960
      ScaleWidth      =   960
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   1470
      Width           =   1020
   End
   Begin VB.PictureBox Picture2 
      AutoSize        =   -1  'True
      Height          =   540
      Left            =   2730
      Picture         =   "frmTurn.frx":0F87
      ScaleHeight     =   480
      ScaleWidth      =   960
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   690
      Width           =   1020
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   540
      Left            =   2730
      Picture         =   "frmTurn.frx":27C9
      ScaleHeight     =   480
      ScaleWidth      =   960
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   150
      Width           =   1020
   End
   Begin VB.OptionButton Opt 
      BackColor       =   &H0080FFFF&
      Caption         =   "270°"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   4
      Left            =   300
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2190
      Width           =   2235
   End
   Begin VB.OptionButton Opt 
      BackColor       =   &H0080FFFF&
      Caption         =   "180°"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   3
      Left            =   300
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1920
      Width           =   2235
   End
   Begin VB.OptionButton Opt 
      BackColor       =   &H0080FFFF&
      Caption         =   "90°"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   2
      Left            =   300
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1650
      Width           =   2235
   End
   Begin VB.OptionButton Opt 
      BackColor       =   &H0080FFFF&
      Caption         =   "Vertical flip"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   1
      Left            =   210
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   795
      Width           =   2445
   End
   Begin VB.OptionButton Opt 
      BackColor       =   &H0080FFFF&
      Caption         =   "Horizontal flip"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   0
      Left            =   210
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   465
      Value           =   -1  'True
      Width           =   2445
   End
   Begin SMBMaker.dbFrame dbFrame1 
      Height          =   1080
      Left            =   135
      TabIndex        =   10
      Top             =   150
      Width           =   2595
      _ExtentX        =   4577
      _ExtentY        =   1905
      Caption         =   "Flip"
      ResID           =   2251
      EAC             =   0   'False
   End
   Begin SMBMaker.dbFrame Label1 
      Height          =   1140
      Left            =   135
      TabIndex        =   11
      Top             =   1410
      Width           =   2595
      _ExtentX        =   4577
      _ExtentY        =   2011
      Caption         =   "Rotate by angle:"
      EAC             =   0   'False
   End
   Begin SMBMaker.dbButton CancelButton 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   2347
      TabIndex        =   6
      Top             =   2610
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      MouseIcon       =   "frmTurn.frx":400B
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Others          =   $"frmTurn.frx":4027
      OthersPresent   =   -1  'True
   End
   Begin SMBMaker.dbButton OKButton 
      Default         =   -1  'True
      Height          =   375
      Left            =   322
      TabIndex        =   5
      Top             =   2610
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      MouseIcon       =   "frmTurn.frx":4077
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Others          =   $"frmTurn.frx":4093
      OthersPresent   =   -1  'True
   End
End
Attribute VB_Name = "frmTurn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub CancelButton_Click()
Me.Tag = "c"
Me.Hide
End Sub

Private Sub Form_Paint()
Me.PaintPicture LoadResPicture(RES_BackPicture, vbResBitmap), 0, 0, Me.ScaleWidth, Me.ScaleHeight
End Sub

Private Sub Form_Initialize()
'Const Action = 1, F = 2070, T = 2078
'Dim obj As Object, i As Integer, index As Integer, tmp As String
'On Error Resume Next
'For Each obj In Me
'    index = -1
'    index = obj.index
'    Err.Clear
'    tmp = obj.Caption
'    If Err = 0 Then
'        If Action = 0 Then
'        Debug.Print obj.Name + IIf((index = -1), "", "(" + CStr(index) + ")") + ".Caption = ", """" + tmp + """"
'        Else
'        For i = F To T
'            If grsf(i) = tmp Then
'            Debug.Print obj.Name + IIf((index = -1), "", "(" + CStr(index) + ")") + ".Caption = ", "grsf(" + CStr(i) + ")"
'            Exit For
'            End If
'        Next i
'        End If
'    Else
'        Err.Clear
'        'Debug.Print "-"
'    End If
'
'    index = -1
'    index = obj.index
'    Err.Clear
'    tmp = obj.ToolTipText
'    If Err = 0 Then
'        If Action = 0 Then
'        Debug.Print obj.Name + IIf((index = -1), "", "(" + CStr(index) + ")") + ".ToolTipText = ", """" + tmp + """"
'        Else
'        For i = F To T
'            If grsf(i) = tmp Then
'            Debug.Print obj.Name + IIf((index = -1), "", "(" + CStr(index) + ")") + ".tooltiptext = ", "grsf(" + CStr(i) + ")"
'            Exit For
'            End If
'        Next i
'        End If
'    Else
'        Err.Clear
'        'Debug.Print "-"
'    End If
'
'Next
'End
dbLoadCaptions
LoadSettings
End Sub

Public Sub LoadSettings()
Dim i As Integer
i = Val(dbGetSetting("Effects\Turn", "Item", "0"))
Opt(i).Value = True
End Sub

Public Sub SaveSettings()
Dim i As Integer
For i = Opt.lBound To Opt.UBound
    If Opt(i).Value Then dbSaveSetting "Effects\Turn", "Item", CStr(i)
Next i
End Sub

Private Sub OkButton_Click()
Me.Tag = ""
SaveSettings
Me.Hide
End Sub

Public Property Get Method() As Integer
Dim i As Integer
For i = Opt.lBound To Opt.UBound
If Opt(i).Value Then Method = i
Next i
End Property

Public Property Let Method(ByVal iNew As Integer)
Opt(iNew).Value = True
End Property

Sub dbLoadCaptions()
Me.Caption = GRSF(2141)
Opt(4).Caption = GRSF(2070)
Opt(3).Caption = GRSF(2071)
Opt(2).Caption = GRSF(2072)
Opt(1).Caption = GRSF(2073)
Opt(0).Caption = GRSF(2074)
CancelButton.Caption = GRSF(2075)
OkButton.Caption = GRSF(2076)
Label1.Caption = GRSF(2077)
Label1.ToolTipText = GRSF(2078)
'Me.Icon = LoadResPicture(Me.Name, vbResIcon)
End Sub

