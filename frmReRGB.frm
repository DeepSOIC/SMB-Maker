VERSION 5.00
Begin VB.Form frmReRGB 
   BackColor       =   &H00E3DFE0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "RGB Преобразование"
   ClientHeight    =   3990
   ClientLeft      =   2760
   ClientTop       =   4050
   ClientWidth     =   4185
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   204
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmReRGB.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3990
   ScaleWidth      =   4185
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin SMBMaker.ctlResourcizator Resr1 
      Left            =   180
      Top             =   165
      _ExtentX        =   1164
      _ExtentY        =   741
      ResID           =   9913
   End
   Begin VB.Timer tmrChanger 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   990
      Top             =   2025
   End
   Begin VB.TextBox txtItem 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   5
      Left            =   3330
      TabIndex        =   5
      Top             =   870
      Width           =   735
   End
   Begin VB.TextBox txtItem 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   1860
      TabIndex        =   0
      Top             =   585
      Width           =   735
   End
   Begin VB.TextBox txtItem 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   8
      Left            =   3330
      TabIndex        =   8
      Top             =   1155
      Width           =   735
   End
   Begin VB.TextBox txtItem 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   7
      Left            =   2595
      TabIndex        =   7
      Top             =   1155
      Width           =   735
   End
   Begin VB.TextBox txtItem 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   6
      Left            =   1860
      TabIndex        =   6
      Top             =   1155
      Width           =   735
   End
   Begin VB.TextBox txtItem 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   4
      Left            =   2595
      TabIndex        =   4
      Top             =   870
      Width           =   735
   End
   Begin VB.TextBox txtItem 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   3
      Left            =   1860
      TabIndex        =   3
      Top             =   870
      Width           =   735
   End
   Begin VB.TextBox txtItem 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   2
      Left            =   3330
      TabIndex        =   2
      Top             =   585
      Width           =   735
   End
   Begin VB.TextBox txtItem 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   2595
      TabIndex        =   1
      Top             =   585
      Width           =   735
   End
   Begin SMBMaker.dbButton btnTemplates 
      Height          =   510
      Left            =   2685
      TabIndex        =   11
      Top             =   2535
      Width           =   1380
      _ExtentX        =   2434
      _ExtentY        =   900
      MouseIcon       =   "frmReRGB.frx":4852
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Others          =   $"frmReRGB.frx":486E
      OthersPresent   =   -1  'True
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Просмотр:"
      Height          =   195
      Left            =   75
      TabIndex        =   22
      Top             =   1455
      Width           =   780
   End
   Begin VB.Image iPreview 
      Height          =   1800
      Left            =   75
      Top             =   1665
      Width           =   2400
   End
   Begin SMBMaker.dbButton btnInv 
      Height          =   555
      Left            =   2685
      TabIndex        =   9
      Top             =   1425
      Width           =   1380
      _ExtentX        =   2434
      _ExtentY        =   979
      MouseIcon       =   "frmReRGB.frx":48C5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Others          =   $"frmReRGB.frx":48E1
      OthersPresent   =   -1  'True
   End
   Begin SMBMaker.dbButton OkButton 
      Default         =   -1  'True
      Height          =   420
      Left            =   120
      TabIndex        =   12
      Top             =   3495
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   741
      MouseIcon       =   "frmReRGB.frx":493C
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Others          =   $"frmReRGB.frx":4958
      OthersPresent   =   -1  'True
   End
   Begin SMBMaker.dbButton CancelButton 
      Cancel          =   -1  'True
      Height          =   420
      Left            =   2670
      TabIndex        =   13
      Top             =   3495
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   741
      MouseIcon       =   "frmReRGB.frx":49A4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Others          =   $"frmReRGB.frx":49C0
      OthersPresent   =   -1  'True
   End
   Begin SMBMaker.dbButton btnNorm 
      Height          =   555
      Left            =   2685
      TabIndex        =   10
      Top             =   1980
      Width           =   1380
      _ExtentX        =   2434
      _ExtentY        =   979
      MouseIcon       =   "frmReRGB.frx":4A10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Others          =   $"frmReRGB.frx":4A2C
      OthersPresent   =   -1  'True
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Выход"
      Height          =   195
      Left            =   -885
      TabIndex        =   21
      Top             =   885
      Width           =   1710
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Вход"
      Height          =   195
      Left            =   1845
      TabIndex        =   20
      Top             =   90
      Width           =   2220
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "красный"
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   1
      Left            =   1860
      TabIndex        =   19
      Top             =   345
      Width           =   735
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "зеленый"
      ForeColor       =   &H0000FF00&
      Height          =   195
      Index           =   1
      Left            =   2595
      TabIndex        =   18
      Top             =   345
      Width           =   735
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "синий"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   1
      Left            =   3330
      TabIndex        =   17
      Top             =   345
      Width           =   735
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "красный"
      ForeColor       =   &H000000FF&
      Height          =   285
      Index           =   0
      Left            =   1005
      TabIndex        =   16
      Top             =   585
      Width           =   840
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "синий"
      ForeColor       =   &H00FF0000&
      Height          =   285
      Index           =   0
      Left            =   1005
      TabIndex        =   15
      Top             =   1155
      Width           =   840
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "зеленый"
      ForeColor       =   &H0000FF00&
      Height          =   285
      Index           =   0
      Left            =   1005
      TabIndex        =   14
      Top             =   870
      Width           =   840
   End
   Begin VB.Menu mnuLoad 
      Caption         =   "Загрузить"
   End
   Begin VB.Menu mnuSave 
      Caption         =   "Сохранить"
   End
   Begin VB.Menu mnuPopNorm 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu mnuNorm 
         Caption         =   "красный"
         Index           =   0
      End
      Begin VB.Menu mnuNorm 
         Caption         =   "зеленый"
         Index           =   1
      End
      Begin VB.Menu mnuNorm 
         Caption         =   "синий"
         Index           =   2
      End
      Begin VB.Menu mnuNormRGB 
         Caption         =   "все"
      End
   End
   Begin VB.Menu mnuPopTempl 
      Caption         =   ""
      Enabled         =   0   'False
      Visible         =   0   'False
      Begin VB.Menu mnuLoadIdent 
         Caption         =   "Загрузить: единичная"
      End
      Begin VB.Menu mnuLoadZero 
         Caption         =   "Загрузить: нулевая"
      End
      Begin VB.Menu Sep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMulGray 
         Caption         =   "Умножить: обесцветить на 25%"
      End
      Begin VB.Menu mnuMulGrayPercent 
         Caption         =   "Умножить: обесцветить на ...%"
      End
      Begin VB.Menu mnuMulSwapRB 
         Caption         =   "Умножить: обменять синий - красный"
      End
      Begin VB.Menu mnuMulColorize 
         Caption         =   "Умножить: окрасить"
      End
      Begin VB.Menu mnuMulRotColor 
         Caption         =   "Умножить: поворот оттенка"
      End
      Begin VB.Menu mnuMulLoad 
         Caption         =   "Умножить: Загр. из файла..."
      End
      Begin VB.Menu Sep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSqr 
         Caption         =   "Возвести матр. в квадрат"
      End
   End
End
Attribute VB_Name = "frmReRGB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Text
Option Explicit

Const Matrix_File_Signature As String = "VT Matrix v1 3x3"

Private Type TriVector
    X As Double
    Y As Double
    z As Double
End Type

Dim Matrix() As Double
Dim bLocked As Boolean
Public Event Change()
'
'Public Sub SaveSettings()
'Dim i As Long, j As Long
'For i = 0 To 2
'    For j = 0 To 2
'        dbSaveSetting "Effects\Matrix", CStr(i) + CStr(j), Trim(Str(Matrix(i, j)))
'    Next j
'Next i
'End Sub
''
''Public Sub LoadSettings()
''Dim i As Long, j As Long
''For i = 0 To 2
''    For j = 0 To 2
''        Matrix(i, j) = Val(dbGetSetting("Effects\Matrix", CStr(i) + CStr(j), Trim(Str(IIf(i = j, 1, 0)))))
''    Next j
''Next i
''UpdateTexts
''End Sub

Public Sub GetMatrix(ByRef pMatrix() As Double)
UpdateMatrix
pMatrix = Matrix
End Sub

Public Sub SetMatrix(ByRef pMatrix() As Double)
If AryDims(AryPtr(pMatrix)) <> 2 Then Exit Sub
If LBound(pMatrix, 1) <> 0 Or UBound(pMatrix, 1) <> 2 Or _
   LBound(pMatrix, 2) <> 0 Or UBound(pMatrix, 2) <> 2 Then
    Exit Sub
End If
Matrix = pMatrix
UpdateTexts
End Sub

Private Sub btnInv_Click()
On Error GoTo eh
MakeUnMatrix Matrix
UpdateTexts
vChange
Exit Sub
eh:
If Err.Number = dbCWS Then Exit Sub
If Err.Number = 116 Then
    dbMsgBox 2399, vbInformation '"The inverse matrix cannot be generated."
Else
    MsgBox Err.Description
End If
End Sub

Private Sub btnNorm_Click()
PopupMenu mnuPopNorm
End Sub

Private Sub btnTemplates_Click()
PopupMenu mnuPopTempl
End Sub

Private Sub Form_Load()
ReDim Matrix(0 To 2, 0 To 2)
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

Private Sub mnuLoad_Click()
Dim File As String
On Error GoTo eh
File = ShowOpenDlg(dbMatrixLoad, Me.hWnd, Purpose:="MATR")
LoadMatrixFromFile File, Matrix
UpdateTexts
vChange
Exit Sub
eh:
MsgError 2617
End Sub

Private Sub mnuLoadIdent_Click()
Matrix = Matrix3(1, 0, 0, _
                 0, 1, 0, _
                 0, 0, 1)
UpdateTexts
vChange
End Sub

Private Sub mnuLoadZero_Click()
Matrix = Matrix3(0, 0, 0, _
                 0, 0, 0, _
                 0, 0, 0)
UpdateTexts
vChange
End Sub

Private Sub mnuMulColorize_Click()
Dim tMatrix() As Double
Dim lngColor As Long
Dim rgbColor As RGBQUAD
On Error GoTo eh
lngColor = CDl.PickColor(dbGetSettingEx("Matrix", "ColorizeColor", vbLong, &HFF0000, HexNumber:=True))
dbSaveSettingEx "Matrix", "ColorizeColor", lngColor, HexNumber:=True
CopyMemory rgbColor, lngColor, 3
tMatrix = Matrix3(rgbColor.rgbRed / 255, 0, 0, _
                  0, rgbColor.rgbGreen / 255, 0, _
                  0, 0, rgbColor.rgbBlue / 255)
Matrix = MulMatrices(tMatrix, Matrix)
UpdateTexts
vChange
Exit Sub
eh:
MsgError
End Sub

Private Function RotationMatrix(ByVal X As Double, ByVal Y As Double, ByVal z As Double, _
                                ByVal Angle As Double)
Dim AroundVector As TriVector
Dim Vct1 As TriVector
Dim Vct2 As TriVector
Dim Vct3 As TriVector

With AroundVector
    .X = X: .Y = Y: .z = z
End With

With Vct1
    .X = 1: .Y = 0: .z = 0
End With
RotateVector Vct1, AroundVector, Angle

With Vct2
    .X = 0: .Y = 1: .z = 0
End With
RotateVector Vct2, AroundVector, Angle

With Vct3
    .X = 0: .Y = 0: .z = 1
End With
RotateVector Vct3, AroundVector, Angle

RotationMatrix = Matrix3(Vct1.X, Vct2.X, Vct3.X, _
                         Vct1.Y, Vct2.Y, Vct3.Y, _
                         Vct1.z, Vct2.z, Vct3.z)
End Function


Private Sub RotateVector(ByRef v As TriVector, _
                        ByRef VectorAround As TriVector, _
                        ByVal Angle As Double)
Dim vy As TriVector
Dim vx As TriVector
Dim l As Double
Dim X As Double, Y As Double
Dim xn As Double, yn As Double
Dim tV As TriVector
vy = VMul(VectorAround, v)
l = Normalize(vy)
If l = 0 Then Exit Sub
vx = VMul(vy, VectorAround)
Normalize vx
X = ProjectVector(v, vx, tV)
Y = ProjectVector(v, vy, tV)
xn = X * Cos(Angle) - Y * Sin(Angle)
yn = Y * Cos(Angle) + X * Sin(Angle)
ProjectVector v, VectorAround, tV
v.X = vx.X * xn + vy.X * yn + tV.X
v.Y = vx.Y * xn + vy.Y * yn + tV.Y
v.z = vx.z * xn + vy.z * yn + tV.z
End Sub

Private Function VMul(ByRef v1 As TriVector, _
                ByRef v2 As TriVector) As TriVector
VMul.X = v1.Y * v2.z - v2.Y * v1.z
VMul.Y = v1.z * v2.X - v2.z * v1.X
VMul.z = v1.X * v2.Y - v2.X * v1.Y
End Function

Private Function Normalize(ByRef Vct As TriVector) As Double
Dim l As Double
l = Sqr(Vct.X * Vct.X + Vct.Y * Vct.Y + Vct.z * Vct.z)
Normalize = l
If l = 1 Then Exit Function
If l > 0 Then
    Vct.X = Vct.X / l
    Vct.Y = Vct.Y / l
    Vct.z = Vct.z / l
End If
End Function

Private Function ProjectVector(ByRef VectorToProject As TriVector, _
                              ByRef VectorProjectTo As TriVector, _
                              ByRef Result As TriVector) As Double
Dim l As Double
Dim tmpVec As TriVector
Dim t As Double
tmpVec = VectorProjectTo
Normalize tmpVec
t = tmpVec.X * VectorToProject.X + _
    tmpVec.Y * VectorToProject.Y + _
    tmpVec.z * VectorToProject.z
Result.X = tmpVec.X * t
Result.Y = tmpVec.Y * t
Result.z = tmpVec.z * t
ProjectVector = t
End Function


Private Sub MulGray(ByVal Dec As Double)
Dim MatrDeColour() As Double
Dim MatrixIdent() As Double
Dim tMatrix() As Double
MatrDeColour = Matrix3(1 / 3, 1 / 3, 1 / 3, _
                       1 / 3, 1 / 3, 1 / 3, _
                       1 / 3, 1 / 3, 1 / 3)
MatrixIdent = Matrix3(1, 0, 0, _
                      0, 1, 0, _
                      0, 0, 1)
tMatrix = SumMatrices(MulMatrices(MatrDeColour, Diag(Dec, Dec, Dec)), _
                      MulMatrices(MatrixIdent, Diag(1 - Dec, 1 - Dec, 1 - Dec)))
Matrix = MulMatrices(tMatrix, Matrix)
UpdateTexts
vChange
End Sub

Private Sub mnuMulGray_Click()
Const Dec = 0.25
MulGray Dec
End Sub

Private Sub mnuMulGrayPercent_Click()
Dim Dec As Double
On Error GoTo eh
Dec = dbGetSettingEx("RGBMatrix", "DecolourPercentage", vbDouble, 50#)
EditNumber Dec, Message:=2616, MinValue:=-1000, MaxValue:=1000
MulGray Dec / 100
dbSaveSettingEx "RGBMatrix", "DecolourPercentage", Dec
Exit Sub
eh:
MsgError
End Sub

Private Sub mnuMulLoad_Click()
Dim File As String
Dim tMatrix() As Double
On Error GoTo eh
File = ShowOpenDlg(dbMatrixLoad, Me.hWnd, Purpose:="MATR")
LoadMatrixFromFile File, tMatrix

Matrix = MulMatrices(tMatrix, Matrix)
UpdateTexts
vChange
Exit Sub
eh:
MsgError
End Sub

Private Sub mnuMulRotColor_Click()
Dim tMatrix() As Double
Dim Angle As Double
On Error GoTo eh
'Angle = Angle / Pi * 180
Angle = dbGetSettingEx("RGBMatrix", "RotationAngle", vbDouble, 45#)
EditNumber Angle, 2615, -360, 360 'Please input the hue rotation angle, in degrees. From -360 to +360.`Hue rotation
dbSaveSettingEx "RGBMatrix", "RotationAngle", Angle
'Angle = Angle * Pi / 180
tMatrix = RotationMatrix(1, 1, 1, Angle * Pi / 180)
Matrix = MulMatrices(tMatrix, Matrix)
UpdateTexts
vChange
Exit Sub
eh:
MsgError
End Sub

Private Sub mnuMulSwapRB_Click()
Matrix = MulMatrices(Matrix3(0, 0, 1, 0, 1, 0, 1, 0, 0), Matrix)
UpdateTexts
vChange
End Sub

Private Sub mnuNorm_Click(Index As Integer)
Dim s As Double
Dim i As Long
For i = 0 To UBound(Matrix)
    s = s + Matrix(Index, i)
Next i
If s <> 0 Then
    For i = 0 To UBound(Matrix)
        Matrix(Index, i) = Matrix(Index, i) / s
    Next i
    UpdateTexts
    vChange
End If
End Sub

Private Sub mnuNormRGB_Click()
mnuNorm_Click 0
mnuNorm_Click 1
mnuNorm_Click 2
End Sub

Private Sub mnuSave_Click()
Dim File As String
On Error GoTo eh
File = ShowSaveDlg(dbMatrixSave, Me.hWnd, Purpose:="MATR")
SaveMatrixToFile Matrix, File
Exit Sub
eh:
MsgError
End Sub

Private Sub mnuSqr_Click()
Matrix = MulMatrices(Matrix, Matrix)
UpdateTexts
vChange
End Sub

Private Sub OkButton_Click()
txtItem_LostFocus 0
'SaveSettings
Me.Tag = ""
Me.Hide
End Sub
'
'Private Sub dbLoadCaptions()
'Label1(0).Caption = GRSF(2173)
'Label1(1).Caption = GRSF(2173)
'
'Label2(0).Caption = GRSF(2176)
'Label2(1).Caption = GRSF(2176)
'
'Label3(0).Caption = GRSF(2175)
'Label3(1).Caption = GRSF(2175)
'
'Label5.Caption = GRSF(2174)
'Label6.Caption = GRSF(2397)
'
'Me.Caption = GRSF(2177)
'OkButton.Caption = GRSF(2178)
'CancelButton.Caption = GRSF(2179)
'
'mnuNorm(0).Caption = GRSF(2401)
'mnuNorm(1).Caption = GRSF(2402)
'mnuNorm(2).Caption = GRSF(2403)
'End Sub

Private Sub tmrChanger_Timer()
tmrChanger.Enabled = False
RaiseEvent Change
End Sub

Private Sub vChange()
tmrChanger.Enabled = False
tmrChanger.Enabled = True
End Sub

Private Sub txtItem_Change(Index As Integer)
vChange
End Sub

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
    For j = 0 To 2
        Matrix(i, j) = Val(txtItem(j + i * 3).Text)
    Next j
Next i
End Sub

Public Sub UpdateTexts()
Dim i As Long, j As Long
bLocked = True
For i = 0 To 2
    For j = 0 To 2
        txtItem(j + i * 3).Text = Trim(Str(Round(Matrix(i, j), 8)))
    Next j
Next i
bLocked = False
End Sub

Public Sub MakeUnMatrix(ByRef pMatrix() As Double)
Dim i As Long, j As Long
Dim UnMatrix() As Double, Matrix() As Double
Dim k As Long
Dim n As Long
Matrix = pMatrix
n = UBound(Matrix, 1) + 1
If UBound(Matrix, 2) + 1 <> n Then
    Err.Raise 116, "MakeUnMatrix", "No inverse matrix."
End If
ReDim UnMatrix(0 To n - 1, 0 To n - 1)
For i = 0 To n - 1
    UnMatrix(i, i) = 1
Next i
'Stage 1. make the low-daig elems 0
For i = 0 To n - 1
    'search for non-zero line
    For j = i To n - 1
        If Matrix(j, i) <> 0 Then
            Exit For
        End If
    Next j
    If j = n Then
        Err.Raise 116, "MakeUnMatrix", "No inverse matrix."
    End If
    AddRow Matrix, UnMatrix, i, j, (1 - Matrix(i, i)) / Matrix(j, i)
    Matrix(i, i) = 1
    
    'make zeros under i,i
    For j = i + 1 To n - 1
        AddRow Matrix, UnMatrix, j, i, -Matrix(j, i)
        Matrix(j, i) = 0
    Next j
Next i

'Stage 2. Make upper diag elems 0's
For j = n - 1 To 1 Step -1
    For i = 0 To j - 1
        AddRow Matrix, UnMatrix, i, j, -Matrix(i, j)
        Matrix(i, j) = 0
    Next i
Next j

pMatrix = UnMatrix
End Sub

Private Sub AddRow(ByRef Matrix1() As Double, ByRef Matrix2() As Double, ByVal IDest As Long, ByVal ISrc As Long, Multiplier As Double)
Dim i As Long
If ISrc = IDest And Multiplier = -1 Then Err.Raise 118, "AddRow", "Programmer's error"
For i = 0 To UBound(Matrix1, 2)
    Matrix1(IDest, i) = Matrix1(IDest, i) + Matrix1(ISrc, i) * Multiplier
    Matrix2(IDest, i) = Matrix2(IDest, i) + Matrix2(ISrc, i) * Multiplier
Next i
End Sub

Private Function MulMatrices(ByRef Matrix1() As Double, ByRef Matrix2() As Double) As Double()
Dim w1 As Long, h1 As Long, w2 As Long, h2 As Long
Dim tmp As Double
Dim i As Long, j As Long, k As Long
Dim Result() As Double
w1 = UBound(Matrix1, 2) + 1
h1 = UBound(Matrix1, 1) + 1
w2 = UBound(Matrix2, 2) + 1
h2 = UBound(Matrix2, 1) + 1
If w1 <> h2 Then
    Err.Raise 32123, "MulMatrices", "Matrices cannot be multiplied. Dimension problem."
End If
ReDim Result(0 To h1 - 1, 0 To w2 - 1)
For j = 0 To w2 - 1
    For i = 0 To h1 - 1
        tmp = 0
        For k = 0 To w1 - 1
            tmp = tmp + Matrix2(k, j) * Matrix1(i, k)
        Next k
        Result(i, j) = tmp
    Next i
Next j
MulMatrices = Result
End Function

Private Function SumMatrices(ByRef Matrix1() As Double, ByRef Matrix2() As Double) As Double()
Dim w1 As Long, h1 As Long, w2 As Long, h2 As Long
Dim tmp As Double
Dim i As Long, j As Long, k As Long
Dim Result() As Double
w1 = UBound(Matrix1, 2) + 1
h1 = UBound(Matrix1, 1) + 1
w2 = UBound(Matrix2, 2) + 1
h2 = UBound(Matrix2, 1) + 1
If w1 <> w2 Or h1 <> h2 Then
    Err.Raise 32123, "MulMatrices", "Matrices cannot be multiplied. Dimension problem."
End If
ReDim Result(0 To h1 - 1, 0 To w1 - 1)
For j = 0 To w2 - 1
    For i = 0 To h1 - 1
        Result(i, j) = Matrix1(i, j) + Matrix2(i, j)
    Next i
Next j
SumMatrices = Result
End Function

Public Function Matrix3(ParamArray Vals() As Variant) As Double()
Dim i As Long, j As Long
Dim t As Long
Dim UB As Long
Dim Result() As Double
ReDim Result(0 To 2, 0 To 2)
UB = UBound(Vals)
t = 0
For i = 0 To 2
    For j = 0 To 2
        If t > UB Then Exit For
        Result(i, j) = Vals(t)
        t = t + 1
    Next j
    If t > UB Then Exit For
Next i
Matrix3 = Result
End Function

Public Function Diag(ParamArray Vals() As Variant) As Double()
Dim i As Long
Dim UB As Long
Dim Result() As Double
UB = UBound(Vals)
ReDim Result(0 To UB, 0 To UB)
For i = 0 To UB
    Result(i, i) = Vals(i)
Next i
Diag = Result
End Function

Private Sub LoadMatrixFromFile(ByRef File As String, ByRef Matrix() As Double)
Dim MatrixFile As Long
Dim tmp As String
Dim tMatrix() As Double
Dim i As Long, j As Long
Dim Arr() As String
Dim nArr As Long, ArrPos As Long
MatrixFile = FreeFile
Open File For Input As MatrixFile
    On Error GoTo eh
    
    Line Input #MatrixFile, tmp
    tmp = Trim(tmp)
    If tmp <> Matrix_File_Signature Then
        Err.Raise 1212, , "Incorrect matrix file signature!"
    End If
    
    ReDim tMatrix(0 To 2, 0 To 2)
    
    For i = 0 To 2
        For j = 0 To 2
            GoSub ReadVal
            
            tMatrix(i, j) = dbVal(tmp, vbDouble, -100000#, 100000#)
        Next j
    Next i
    
Close MatrixFile
Matrix = tMatrix
Exit Sub
eh:
PushError
Close MatrixFile
PopError
ErrRaise "LoadMatrixFromFile"

ReadVal:
    Do
        GoSub ReadVal1
    Loop Until Len(tmp) > 0
Return

ReadVal1:
    If ArrPos >= nArr Then
        Line Input #MatrixFile, tmp
        If Len(tmp) = 0 Then
            GoSub ReadVal1
        Else
            Arr = Split(tmp, " ")
            nArr = UBound(Arr) + 1
            ArrPos = 0
        End If
        tmp = ""
    Else
        tmp = Arr(ArrPos)
        ArrPos = ArrPos + 1
    End If
Return

End Sub

Private Sub SaveMatrixToFile(ByRef Matrix() As Double, ByRef File As String)
    Dim MatrixFile As String
    Dim i As Long, j As Long
    
    MatrixFile = FreeFile
    Open File For Output As MatrixFile
        On Error GoTo eh
        Print #MatrixFile, Matrix_File_Signature
        For i = 0 To UBound(Matrix, 1)
            For j = 0 To UBound(Matrix, 2)
                Print #MatrixFile, Str(Matrix(i, j)),
            Next j
            Print #MatrixFile, ""
        Next i
    Close MatrixFile

Exit Sub
eh:
    PushError
    Close MatrixFile
    PopError
    ErrRaise "SaveMatrixToFile"
End Sub
