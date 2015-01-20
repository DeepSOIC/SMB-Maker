VERSION 5.00
Begin VB.Form frmDiff 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Differentiate"
   ClientHeight    =   4500
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7380
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   204
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   300
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   492
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkTex 
      BackColor       =   &H00FF80FF&
      Caption         =   "Texture Mode"
      Height          =   405
      Left            =   225
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   3975
      Width           =   2490
   End
   Begin SMBMaker.ctlResourcizator Resr1 
      Left            =   3090
      Top             =   3975
      _ExtentX        =   1164
      _ExtentY        =   741
      ResID           =   9909
   End
   Begin VB.OptionButton optMode 
      BackColor       =   &H0000FFFF&
      Caption         =   "Relief"
      Height          =   285
      Index           =   1
      Left            =   3795
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   75
      Width           =   3465
   End
   Begin VB.OptionButton optMode 
      BackColor       =   &H0000FFFF&
      Caption         =   "Differentiate"
      Height          =   285
      Index           =   0
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   75
      Value           =   -1  'True
      Width           =   3465
   End
   Begin SMBMaker.dbFrame fRelief 
      Height          =   1095
      Left            =   3765
      TabIndex        =   2
      Top             =   45
      Width           =   3525
      _ExtentX        =   6218
      _ExtentY        =   1931
      EAC             =   -1  'True
      Begin SMBMaker.ctlNumBox nmbX 
         Height          =   285
         Left            =   1155
         TabIndex        =   14
         Top             =   390
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   503
         Value           =   -3
         Min             =   -12
         Max             =   12
         NumType         =   3
         HorzMode        =   -1  'True
         EditName        =   "Horizontal offset of source data added. In pixels."
         NLn             =   0
         NativeValues    =   "-3|-3|+3|3"
         Enabled         =   0   'False
      End
      Begin SMBMaker.ctlNumBox nmbY 
         Height          =   285
         Left            =   1155
         TabIndex        =   15
         Top             =   675
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   503
         Value           =   -2
         Min             =   -12
         Max             =   12
         NumType         =   3
         HorzMode        =   -1  'True
         EditName        =   "Vertical offset of source data added. In pixels."
         NLn             =   0
         NativeValues    =   "-2|-2|+2|2"
         Enabled         =   0   'False
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Y offset"
         Height          =   195
         Left            =   45
         TabIndex        =   6
         Top             =   720
         Width           =   570
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "X offset"
         Height          =   195
         Left            =   45
         TabIndex        =   5
         Top             =   435
         Width           =   570
      End
   End
   Begin SMBMaker.dbFrame dbFrame1 
      Height          =   1605
      Left            =   165
      TabIndex        =   9
      Top             =   2145
      Width           =   3090
      _ExtentX        =   5450
      _ExtentY        =   2831
      Caption         =   "Advanced"
      EAC             =   0   'False
      Begin SMBMaker.ctlNumBox nmbDiffAmp 
         Height          =   270
         Left            =   225
         TabIndex        =   16
         Top             =   525
         Width           =   2670
         _ExtentX        =   4710
         _ExtentY        =   476
         Value           =   1
         Min             =   -255
         Max             =   255
         NumType         =   14
         HorzMode        =   -1  'True
         NLn             =   0
         SliderVisible   =   0   'False
      End
      Begin SMBMaker.ctlNumBox nmbDataAmp 
         Height          =   270
         Left            =   210
         TabIndex        =   17
         Top             =   1155
         Width           =   2670
         _ExtentX        =   4710
         _ExtentY        =   476
         Value           =   1
         Min             =   -255
         Max             =   255
         NumType         =   14
         HorzMode        =   -1  'True
         NLn             =   0
         Enabled         =   0   'False
         SliderVisible   =   0   'False
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Relief source data amplification"
         Height          =   240
         Left            =   60
         TabIndex        =   11
         Top             =   900
         Width           =   2925
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Result Amplification"
         Height          =   240
         Left            =   60
         TabIndex        =   10
         Top             =   255
         Width           =   2925
      End
   End
   Begin VB.Line Line1 
      X1              =   10
      X2              =   478
      Y1              =   135
      Y2              =   135
   End
   Begin VB.Image iPreview 
      Height          =   1515
      Left            =   3750
      Top             =   2385
      Width           =   3240
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Preview:"
      Height          =   240
      Left            =   3780
      TabIndex        =   12
      Top             =   2130
      Width           =   2235
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "The relief effect is simply differentiated image added to the source image with custom offset. Here you can choose the offset."
      Height          =   975
      Left            =   3780
      TabIndex        =   8
      Top             =   1155
      Width           =   3480
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmDiff.frx":0000
      Height          =   1905
      Left            =   120
      TabIndex        =   7
      Top             =   345
      Width           =   3480
   End
   Begin SMBMaker.dbButton CancelButton 
      Cancel          =   -1  'True
      Height          =   330
      Left            =   5940
      TabIndex        =   1
      Top             =   4050
      Width           =   1290
      _ExtentX        =   2275
      _ExtentY        =   582
      MouseIcon       =   "frmDiff.frx":00B4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Others          =   $"frmDiff.frx":00D0
      OthersPresent   =   -1  'True
   End
   Begin SMBMaker.dbButton OkButton 
      Default         =   -1  'True
      Height          =   330
      Left            =   4470
      TabIndex        =   0
      Top             =   4050
      Width           =   1290
      _ExtentX        =   2275
      _ExtentY        =   582
      MouseIcon       =   "frmDiff.frx":0120
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Others          =   $"frmDiff.frx":013C
      OthersPresent   =   -1  'True
   End
End
Attribute VB_Name = "frmDiff"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Dim DirX As Single, DirY As Single
Dim bLocked As Boolean
'Dim DiffAmp As Variant
'Dim DataAmp As Variant

Public Event Change()


Private Sub CancelButton_Click()
Me.Tag = "C"
Me.Hide
End Sub

Private Sub chkTex_Click()
vChange
End Sub

Private Sub Form_Load()
Resr1.LoadCaptions
'DirX = -3
'DirY = -2
'LoadSettings
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

Private Sub nmbDataAmp_Change()
vChange
End Sub

Private Sub nmbDiffAmp_Change()
vChange
End Sub

Private Sub nmbX_Change()
vChange
End Sub

Private Sub nmbY_Change()
vChange
End Sub

Private Sub OkButton_Click()
'SaveSettings
Me.Tag = ""
Me.Hide
End Sub

Private Sub optMode_Click(Index As Integer)
Dim b As Boolean
b = optMode(1).Value
fRelief.Enabled = b
Label6.Enabled = b
nmbDataAmp.Enabled = b
vChange
End Sub

Sub vChange()
RaiseEvent Change
End Sub
'
'Private Sub pctDir_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
'Dim CX As Long, CY As Long
'CX = pctDir.ScaleWidth \ 2
'CY = pctDir.ScaleHeight \ 2
'If Button > 0 Then
'    DirX = x - CX
'    DirY = y - CY
'    'DrawArrow cx, cy, X, Y
'    TruncateDir
'    pctDir_Paint
'    UpdateTexts
'End If
'End Sub
'
'Private Sub DrawArrow(ByVal x0 As Long, ByVal y0 As Long, ByVal x1 As Long, ByVal y1 As Long)
'Const ArL As Single = 7
'Const ArH As Single = 4
'Dim dx As Long, dy As Long
'Dim dl As Single
'dx = x1 - x0
'dy = y1 - y0
'dl = Sqr(dx * dx + dy * dy)
'pctDir.Line (x0, y0)-(x1, y1)
'If dl > 0 Then
'    pctDir.Line (x1, y1)-Step(-dx / dl * ArL + dy / dl * ArH, -dy / dl * ArL - dx / dl * ArH)
'    pctDir.Line (x1, y1)-Step(-dx / dl * ArL - dy / dl * ArH, -dy / dl * ArL + dx / dl * ArH)
'End If
'End Sub
'
'Private Sub pctDir_Paint()
'Dim CX As Long, CY As Long
'CX = pctDir.ScaleWidth \ 2
'CY = pctDir.ScaleHeight \ 2
'
'On Error GoTo eh
'pctDir.PaintPicture gBackPicture, 0, 0, pctDir.ScaleWidth, pctDir.ScaleHeight
'On Error GoTo 0
'DrawArrow CX, CY, CX + DirX, CY + DirY
'Exit Sub
'eh:
'pctDir.Cls
'Resume Next
'End Sub
'
'Private Sub UpdateTexts()
'bLocked = True
'txtX.Text = CStr(CInt(DirX))
'txtY.Text = CStr(CInt(DirY))
'bLocked = False
'End Sub
'
'Private Sub pctDir_Resize()
'pctDir.Scale (0, 0)-(18, 18)
'End Sub
'
'Sub ReadData()
'DataAmp = dbVal(txtDataAmp.Text, vbDecimal)
'DiffAmp = dbVal()
'End Sub
'
'Private Sub txtX_Change()
'If bLocked Then Exit Sub
'DirX = Val(txtX.Text)
'pctDir_Paint
'End Sub
'
'Private Sub txtX_LostFocus()
'TruncateDir
'pctDir_Paint
'UpdateTexts
'End Sub
'
'Private Sub txtY_Change()
'If bLocked Then Exit Sub
'DirY = Val(txtY.Text)
'pctDir_Paint
'End Sub
'
'Private Sub TruncateDir()
'    If Abs(DirX) > 15 Then
'        DirY = DirY * 15 / Abs(DirX)
'        DirX = 15 * Sgn(DirX)
'    End If
'    If Abs(DirY) > 15 Then
'        DirX = DirX * 15 / Abs(DirY)
'        DirY = 15 * Sgn(DirY)
'    End If
'
'End Sub
'
'Private Sub txtY_LostFocus()
'txtX_LostFocus
'End Sub
'
'Public Sub LoadSettings()
'On Error Resume Next
'optMode(CInt(dbGetSetting("Effects\Diff", "Mode", "0"))).Value = True
'
'DirX = CInt(dbGetSetting("Effects\Diff", "DirX", "-3"))
'DirY = CInt(dbGetSetting("Effects\Diff", "DirY", "-2"))
'
'TruncateDir
'UpdateTexts
'End Sub
'
'Public Sub SaveSettings()
'Dim i As Long
'dbSaveSetting "Effects\Diff", "DirX", CStr(CInt(DirX))
'dbSaveSetting "Effects\Diff", "DirY", CStr(CInt(DirY))
'For i = 0 To optMode.UBound
'    If optMode(i).Value Then
'        dbSaveSetting "Effects\Diff", "Mode", CStr(i)
'        Exit For
'    End If
'Next i
'End Sub
'
'Public Function GetMode() As Long
'Dim i As Long
'For i = 0 To optMode.UBound
'    If optMode(i).Value Then
'        GetMode = i
'        Exit Function
'    End If
'Next i
'End Function
'
'Friend Function GetDir(ByRef xyDir As POINTAPI)
'xyDir.x = DirX
'xyDir.y = DirY
'End Function
'
