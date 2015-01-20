VERSION 5.00
Begin VB.Form frmMatrixMix 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mixing matrix"
   ClientHeight    =   2040
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6855
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   136
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   457
   StartUpPosition =   1  'CenterOwner
   Begin SMBMaker.ctlResourcizator Resr1 
      Left            =   390
      Top             =   1380
      _ExtentX        =   1164
      _ExtentY        =   741
      ResID           =   9914
   End
   Begin VB.TextBox txtItem 
      BackColor       =   &H00808000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   11
      Left            =   6075
      TabIndex        =   28
      Top             =   870
      Width           =   735
   End
   Begin VB.TextBox txtItem 
      BackColor       =   &H000000FF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   3
      Left            =   4605
      TabIndex        =   27
      Top             =   585
      Width           =   735
   End
   Begin VB.TextBox txtItem 
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   17
      Left            =   6075
      TabIndex        =   26
      Top             =   1155
      Width           =   735
   End
   Begin VB.TextBox txtItem 
      BackColor       =   &H00808000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   16
      Left            =   5340
      TabIndex        =   25
      Top             =   1155
      Width           =   735
   End
   Begin VB.TextBox txtItem 
      BackColor       =   &H00800080&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   15
      Left            =   4605
      TabIndex        =   24
      Top             =   1155
      Width           =   735
   End
   Begin VB.TextBox txtItem 
      BackColor       =   &H0000FF00&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   10
      Left            =   5340
      TabIndex        =   23
      Top             =   885
      Width           =   735
   End
   Begin VB.TextBox txtItem 
      BackColor       =   &H00008080&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   9
      Left            =   4605
      TabIndex        =   22
      Top             =   870
      Width           =   735
   End
   Begin VB.TextBox txtItem 
      BackColor       =   &H00800080&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   5
      Left            =   6075
      TabIndex        =   21
      Top             =   585
      Width           =   735
   End
   Begin VB.TextBox txtItem 
      BackColor       =   &H00008080&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   4
      Left            =   5340
      TabIndex        =   20
      Top             =   585
      Width           =   735
   End
   Begin VB.TextBox txtItem 
      BackColor       =   &H00008080&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   1
      Left            =   3105
      TabIndex        =   8
      Top             =   585
      Width           =   735
   End
   Begin VB.TextBox txtItem 
      BackColor       =   &H00800080&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   2
      Left            =   3840
      TabIndex        =   7
      Top             =   585
      Width           =   735
   End
   Begin VB.TextBox txtItem 
      BackColor       =   &H00008080&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   6
      Left            =   2370
      TabIndex        =   6
      Top             =   870
      Width           =   735
   End
   Begin VB.TextBox txtItem 
      BackColor       =   &H0000FF00&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   7
      Left            =   3105
      TabIndex        =   5
      Top             =   870
      Width           =   735
   End
   Begin VB.TextBox txtItem 
      BackColor       =   &H00800080&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   12
      Left            =   2370
      TabIndex        =   4
      Top             =   1155
      Width           =   735
   End
   Begin VB.TextBox txtItem 
      BackColor       =   &H00808000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   13
      Left            =   3105
      TabIndex        =   3
      Top             =   1155
      Width           =   735
   End
   Begin VB.TextBox txtItem 
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   14
      Left            =   3840
      TabIndex        =   2
      Top             =   1155
      Width           =   735
   End
   Begin VB.TextBox txtItem 
      BackColor       =   &H000000FF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   0
      Left            =   2370
      TabIndex        =   1
      Top             =   585
      Width           =   735
   End
   Begin VB.TextBox txtItem 
      BackColor       =   &H00808000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   8
      Left            =   3840
      TabIndex        =   0
      Top             =   870
      Width           =   735
   End
   Begin SMBMaker.dbButton OkButton 
      Default         =   -1  'True
      Height          =   420
      Left            =   2025
      TabIndex        =   9
      Top             =   1530
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   741
      MouseIcon       =   "frmMatrixMix.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Others          =   $"frmMatrixMix.frx":001C
      OthersPresent   =   -1  'True
   End
   Begin SMBMaker.dbButton CancelButton 
      Cancel          =   -1  'True
      Height          =   420
      Left            =   3435
      TabIndex        =   10
      Top             =   1530
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   741
      MouseIcon       =   "frmMatrixMix.frx":0068
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Others          =   $"frmMatrixMix.frx":0084
      OthersPresent   =   -1  'True
   End
   Begin SMBMaker.dbButton btnNorm 
      Height          =   555
      Left            =   5430
      TabIndex        =   11
      Top             =   1455
      Width           =   1380
      _ExtentX        =   2434
      _ExtentY        =   979
      MouseIcon       =   "frmMatrixMix.frx":00D4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Others          =   $"frmMatrixMix.frx":00F0
      OthersPresent   =   -1  'True
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Picture"
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
      Index           =   1
      Left            =   4590
      TabIndex        =   32
      Top             =   90
      Width           =   2220
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Red:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   2
      Left            =   4605
      TabIndex        =   31
      Top             =   345
      Width           =   735
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Green:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   195
      Index           =   2
      Left            =   5340
      TabIndex        =   30
      Top             =   345
      Width           =   735
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Blue:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   2
      Left            =   6075
      TabIndex        =   29
      Top             =   345
      Width           =   735
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Green:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   285
      Index           =   0
      Left            =   1515
      TabIndex        =   19
      Top             =   870
      Width           =   840
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Blue:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Index           =   0
      Left            =   1515
      TabIndex        =   18
      Top             =   1155
      Width           =   840
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Red:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   285
      Index           =   0
      Left            =   1515
      TabIndex        =   17
      Top             =   585
      Width           =   840
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Blue:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   1
      Left            =   3840
      TabIndex        =   16
      Top             =   345
      Width           =   735
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Green:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   195
      Index           =   1
      Left            =   3105
      TabIndex        =   15
      Top             =   345
      Width           =   735
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Red:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   1
      Left            =   2370
      TabIndex        =   14
      Top             =   345
      Width           =   735
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Fragment"
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
      Index           =   0
      Left            =   2355
      TabIndex        =   13
      Top             =   90
      Width           =   2220
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Result"
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
      Left            =   -195
      TabIndex        =   12
      Top             =   885
      Width           =   1710
   End
   Begin VB.Menu mnuPopNorm 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu mnuNorm 
         Caption         =   "Red"
         Index           =   0
      End
      Begin VB.Menu mnuNorm 
         Caption         =   "Green"
         Index           =   1
      End
      Begin VB.Menu mnuNorm 
         Caption         =   "Blue"
         Index           =   2
      End
   End
End
Attribute VB_Name = "frmMatrixMix"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Text
Option Explicit
Dim Matrix() As Double
Dim bLocked As Boolean

Public Sub GetMatrix(ByRef pMatrix() As Double)
pMatrix = Matrix
End Sub

Private Sub btnNorm_Click()
PopupMenu mnuPopNorm
End Sub

Private Sub Form_Load()
ReDim Matrix(0 To 2, 0 To 5)
'LoadSettings
Resr1.LoadCaptions
End Sub

Private Sub Form_Paint()
Me.PaintPicture LoadResPicture(RES_BackPicture, vbResBitmap), 0, 0, Me.ScaleWidth, Me.ScaleHeight
End Sub

Private Sub CancelButton_Click()
Me.Tag = "C"
Me.Hide
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If UnloadMode = vbFormControlMenu Then
    Cancel = True
    CancelButton_Click
End If
End Sub

Private Sub mnuNorm_Click(Index As Integer)
Dim s As Double
Dim i As Long
For i = 0 To UBound(Matrix, 2)
    s = s + Matrix(Index, i)
Next i
If s <> 0 Then
    For i = 0 To UBound(Matrix, 2)
        Matrix(Index, i) = Matrix(Index, i) / s
    Next i
    UpdateTexts
End If
End Sub

Private Sub OkButton_Click()
txtItem_LostFocus 0
'SaveSettings
Me.Tag = ""
Me.Hide
End Sub
'
'Private Sub dbLoadCaptions()
'Dim i As Long
'For i = Label1.lBound To Label1.UBound
'Label1(i).Caption = GRSF(2173)
'Label2(i).Caption = GRSF(2176)
'Label3(i).Caption = GRSF(2175)
'Next i
'
'Label5(0).Caption = GRSF(2404)
'Label5(1).Caption = GRSF(2405)
'Label6.Caption = GRSF(2397)
'
'Me.Caption = GRSF(2406)
'
'mnuNorm(0).Caption = GRSF(2401)
'mnuNorm(1).Caption = GRSF(2402)
'mnuNorm(2).Caption = GRSF(2403)
'End Sub

Private Sub txtItem_GotFocus(Index As Integer)
SelTextInTextBox txtItem(Index)
End Sub

Private Sub txtItem_LostFocus(Index As Integer)
If bLocked Then Exit Sub
UpdateMatrix
UpdateTexts
End Sub

Public Sub UpdateMatrix()
Dim i As Long, j As Long
For i = 0 To 2
    For j = 0 To 5
        Matrix(i, j) = Val(txtItem(j + i * 6).Text)
    Next j
Next i
End Sub

Public Sub UpdateTexts()
Dim i As Long, j As Long
bLocked = True
For i = 0 To 2
    For j = 0 To 5
        txtItem(j + i * 6).Text = Trim(Str(Round(Matrix(i, j), 8)))
    Next j
Next i
bLocked = False
End Sub

Public Sub SetMatrix(ByRef pMatrix() As Double)
Matrix = pMatrix
UpdateTexts
End Sub

