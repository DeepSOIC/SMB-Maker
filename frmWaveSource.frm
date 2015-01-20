VERSION 5.00
Begin VB.Form frmWaveSource 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Wave source"
   ClientHeight    =   2550
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   3495
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
   ScaleHeight     =   2550
   ScaleWidth      =   3495
   StartUpPosition =   1  'CenterOwner
   Begin SMBMaker.ctlResourcizator Resr1 
      Left            =   960
      Top             =   1470
      _ExtentX        =   1164
      _ExtentY        =   741
      ResID           =   9902
   End
   Begin VB.TextBox txtY 
      Height          =   300
      Left            =   1794
      TabIndex        =   11
      ToolTipText     =   "Y-coordinate of a source. In pixels. Integer value."
      Top             =   1968
      Width           =   780
   End
   Begin VB.TextBox txtX 
      Height          =   300
      Left            =   918
      TabIndex        =   9
      ToolTipText     =   "X-coordinate of a source. In pixels. Integer value."
      Top             =   1968
      Width           =   780
   End
   Begin SMBMaker.ctlColor clr 
      Height          =   252
      Left            =   1272
      TabIndex        =   3
      Top             =   1116
      Width           =   948
      _ExtentX        =   1667
      _ExtentY        =   450
      Color           =   16761024
   End
   Begin SMBMaker.ctlNumBox nmbStrength 
      Height          =   528
      Left            =   588
      TabIndex        =   2
      Top             =   252
      Width           =   1020
      _ExtentX        =   1799
      _ExtentY        =   926
      Min             =   -16
      Max             =   16
      HorzMode        =   0   'False
      EditName        =   "The charge or the intensity of wave source."
      NLn             =   0
   End
   Begin SMBMaker.ctlNumBox nmbWaveLen 
      Height          =   528
      Left            =   1884
      TabIndex        =   6
      Top             =   252
      Width           =   1020
      _ExtentX        =   1799
      _ExtentY        =   926
      Value           =   1
      Min             =   1
      Max             =   255
      NumType         =   3
      HorzMode        =   0   'False
      EditName        =   "Wavelength of wave source. Affects only waves."
   End
   Begin SMBMaker.dbButton CancelButton 
      Cancel          =   -1  'True
      Height          =   276
      Left            =   192
      TabIndex        =   1
      Top             =   1680
      Width           =   732
      _ExtentX        =   1296
      _ExtentY        =   476
      MouseIcon       =   "frmWaveSource.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Others          =   $"frmWaveSource.frx":001C
      OthersPresent   =   -1  'True
   End
   Begin SMBMaker.dbButton OkButton 
      Default         =   -1  'True
      Height          =   276
      Left            =   2568
      TabIndex        =   0
      Top             =   1680
      Width           =   732
      _ExtentX        =   1296
      _ExtentY        =   476
      MouseIcon       =   "frmWaveSource.frx":006C
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Others          =   $"frmWaveSource.frx":0088
      OthersPresent   =   -1  'True
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Y"
      Height          =   270
      Left            =   1800
      TabIndex        =   10
      Top             =   2280
      Width           =   255
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      Height          =   270
      Left            =   1410
      TabIndex        =   8
      Top             =   2280
      Width           =   285
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "WaveLength"
      Height          =   192
      Left            =   1896
      TabIndex        =   7
      Top             =   60
      Width           =   1008
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Color"
      Height          =   192
      Left            =   1272
      TabIndex        =   5
      Top             =   924
      Width           =   948
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Strength"
      Height          =   192
      Left            =   600
      TabIndex        =   4
      Top             =   60
      Width           =   1008
   End
End
Attribute VB_Name = "frmWaveSource"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public AskMany As Boolean
Public Answer_Many As Boolean

Friend Sub EditWS(ByRef WS() As typWaveSource, _
                  ByVal nWS As Long, _
                  ByVal OnlySelected As Boolean, _
                  Optional ByVal FillFields As Boolean = True)
Dim WS1 As typWaveSource
Dim i As Long
Dim FirstFound As Boolean
Dim ColorDiffers As Boolean
Dim StrengthDiffers As Boolean
Dim WLDiffers As Boolean
Dim XDiffers As Boolean
Dim YDiffers As Boolean
Dim MoreThanOne As Boolean

For i = 0 To nWS - 1
    If WS(i).Selected Or Not OnlySelected Then
        If Not FirstFound Then
            WS1 = WS(i)
            FirstFound = True
        Else
'            If WS(i).Color <> WS1.Color Then
'                ColorDiffers = True
'            End If
'            If WS(i).Strength <> WS1.Strength Then
'                StrengthDiffers = True
'            End If
'            If WS(i).WaveLength <> WS1.WaveLength Then
'                WLDiffers = True
'            End If
'            If WS(i).Pos.x <> WS1.Pos.x Then
'                XDiffers = True
'            End If
'            If WS(i).Pos.y <> WS1.Pos.y Then
'                YDiffers = True
'            End If
            MoreThanOne = True
        End If
    End If
Next i

If Not FirstFound Then Err.Raise dbCWS, "EditWS"

Load Me

AskMany = MoreThanOne
'nmbStrength.Enabled = Not StrengthDiffers
'nmbWaveLen.Enabled = Not WLDiffers
'clr.Enabled = Not ColorDiffers
'txtX.Enabled = Not XDiffers
'txtY.Enabled = Not YDiffers

On Error Resume Next
If FillFields Then
    nmbStrength.Value = WS1.Strength
    nmbWaveLen.Value = WS1.WaveLength
    clr.Color = RGB(WS1.Color.rgbBlue, WS1.Color.rgbGreen, WS1.Color.rgbRed)
    txtX.Text = dbCStr(WS1.Pos.X)
    txtY.Text = dbCStr(WS1.Pos.Y)
End If
On Error GoTo 0

Me.Show vbModal

If Me.Tag = "C" Then
    Unload Me
    Err.Raise dbCWS, "EditWS"
End If

On Error Resume Next
If FillFields Then
    StrengthDiffers = Not FPEqual(WS1.Strength, nmbStrength.Value)
    WLDiffers = WS1.WaveLength <> nmbWaveLen.Value
    ColorDiffers = RGB(WS1.Color.rgbBlue, WS1.Color.rgbGreen, WS1.Color.rgbRed) <> clr.Color
    XDiffers = WS1.Pos.X <> dbVal(txtX.Text, vbLong)
    YDiffers = WS1.Pos.Y <> dbVal(txtY.Text, vbLong)
Else
    StrengthDiffers = True
    WLDiffers = True
    ColorDiffers = True
    XDiffers = True
    YDiffers = True
End If

If Not (XDiffers Or _
       YDiffers Or _
       StrengthDiffers Or _
       WLDiffers Or _
       ColorDiffers) Then
    Unload Me
    Err.Raise dbCWS, "EditWS"
End If

CopyMemory WS1.Color, clr.Color, 4&
WS1.WaveLength = nmbWaveLen.Value
WS1.Strength = nmbStrength.Value
WS1.Pos.X = dbVal(txtX.Text, vbLong)
WS1.Pos.Y = dbVal(txtY.Text, vbLong)

For i = 0 To nWS - 1
    If WS(i).Selected Or Not OnlySelected Then
        If ColorDiffers Then
            WS(i).Color = WS1.Color
        End If
        If StrengthDiffers Then
            WS(i).Strength = WS1.Strength
        End If
        If WLDiffers Then
            WS(i).WaveLength = WS1.WaveLength
        End If
        If XDiffers Then
            WS(i).Pos.X = WS1.Pos.X
        End If
        If YDiffers Then
            WS(i).Pos.Y = WS1.Pos.Y
        End If
    End If
Next i

On Error GoTo 0
Unload Me
End Sub

Public Function FPEqual(ByVal Val1 As Double, ByVal Val2 As Double) As Boolean
FPEqual = Abs(Val1 - Val2) < 0.0001
End Function

Private Sub CancelButton_Click()
Me.Tag = "C"
Me.Hide
End Sub

Private Sub Form_Load()
Resr1.LoadCaptions
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

Private Sub txtX_Change()
UpdateOK
End Sub

Private Sub txtY_Change()
UpdateOK
End Sub

Public Sub UpdateOK()
Dim tmp As Double
On Error GoTo eh
tmp = dbVal(txtX.Text)
tmp = dbVal(txtY.Text)
OkButton.Enabled = True
Exit Sub
eh:
OkButton.Enabled = False
End Sub
