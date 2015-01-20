VERSION 5.00
Begin VB.Form frmReallyGamma 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Color gamma detection"
   ClientHeight    =   4200
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11655
   Icon            =   "frmReallyGamma.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   280
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   777
   StartUpPosition =   2  'CenterScreen
   Begin SMBMaker.ctlResourcizator Resr1 
      Left            =   3450
      Top             =   1785
      _ExtentX        =   1164
      _ExtentY        =   741
      ResID           =   9908
   End
   Begin VB.CheckBox chkAll 
      BackColor       =   &H00FFC0FF&
      Caption         =   "All together"
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
      Left            =   4297
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Makes all the components change simultaneously. Use it when you dont want to set each component independently."
      Top             =   3600
      Width           =   3060
   End
   Begin VB.CheckBox chkUn 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Uncorrection"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   465
      Left            =   450
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Makes inverse dependence, which removes gamma correction."
      Top             =   3630
      Width           =   2925
   End
   Begin SMBMaker.ctlNumBox nmbComp 
      Height          =   540
      Index           =   0
      Left            =   0
      TabIndex        =   3
      Top             =   3045
      Width           =   3885
      _ExtentX        =   6853
      _ExtentY        =   953
      Value           =   128
      Min             =   1
      Max             =   254
      NumType         =   3
      HorzMode        =   0   'False
      EditName        =   "Red"
      NLn             =   0
   End
   Begin VB.PictureBox pctPreview 
      AutoRedraw      =   -1  'True
      FillStyle       =   0  'Solid
      Height          =   2010
      Index           =   2
      Left            =   8595
      ScaleHeight     =   130
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   145
      TabIndex        =   2
      Top             =   1035
      Width           =   2235
   End
   Begin VB.PictureBox pctPreview 
      AutoRedraw      =   -1  'True
      FillStyle       =   0  'Solid
      Height          =   2010
      Index           =   1
      Left            =   4695
      ScaleHeight     =   130
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   146
      TabIndex        =   1
      Top             =   1035
      Width           =   2250
   End
   Begin VB.PictureBox pctPreview 
      AutoRedraw      =   -1  'True
      FillStyle       =   0  'Solid
      Height          =   2010
      Index           =   0
      Left            =   810
      ScaleHeight     =   130
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   146
      TabIndex        =   0
      Top             =   1035
      Width           =   2250
   End
   Begin SMBMaker.ctlNumBox nmbComp 
      Height          =   540
      Index           =   1
      Left            =   3885
      TabIndex        =   4
      Top             =   3045
      Width           =   3885
      _ExtentX        =   6853
      _ExtentY        =   953
      Value           =   128
      Min             =   1
      Max             =   254
      NumType         =   3
      HorzMode        =   0   'False
      EditName        =   "Red"
      NLn             =   0
   End
   Begin SMBMaker.ctlNumBox nmbComp 
      Height          =   540
      Index           =   2
      Left            =   7770
      TabIndex        =   5
      Top             =   3045
      Width           =   3885
      _ExtentX        =   6853
      _ExtentY        =   953
      Value           =   128
      Min             =   1
      Max             =   254
      NumType         =   3
      HorzMode        =   0   'False
      EditName        =   "Red"
      NLn             =   0
   End
   Begin SMBMaker.dbButton OkButton 
      Height          =   495
      Left            =   8370
      TabIndex        =   6
      Top             =   3675
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   873
      MouseIcon       =   "frmReallyGamma.frx":000C
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Others          =   $"frmReallyGamma.frx":0028
      OthersPresent   =   -1  'True
   End
   Begin SMBMaker.dbButton CancelButton 
      Cancel          =   -1  'True
      Height          =   495
      Left            =   9780
      TabIndex        =   7
      Top             =   3675
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   873
      MouseIcon       =   "frmReallyGamma.frx":0074
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Others          =   $"frmReallyGamma.frx":0090
      OthersPresent   =   -1  'True
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmReallyGamma.frx":00E0
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1020
      Left            =   187
      TabIndex        =   9
      Top             =   0
      Width           =   11280
   End
End
Attribute VB_Name = "frmReallyGamma"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Comps(0 To 2) As Long

Private Sub pctPreview_Init(ByVal Index As Long, Optional ByVal DrawBG As Boolean = False)
Dim Data() As Long
Dim w As Long, h As Long
Dim bl As Long
Dim X As Long, Y As Long
Dim RGBMask As Long
'Static Initd(0 To 2) As Boolean
RGBMask = 255 * 256 ^ (2 - Index)
w = pctPreview(Index).ScaleWidth
h = pctPreview(Index).ScaleHeight

If DrawBG Then
    ReDim Data(0 To w - 1, 0 To h - 1)
    
    For Y = 0 To h - 1
        bl = ((Y Mod 2) = 0)
        For X = 0 To w - 1
            Data(X, Y) = bl And RGBMask
            'bl = Not bl
        Next X
    Next Y
End If

With pctPreview(Index)
    If DrawBG Then
        DontDoEvents = True
        .Cls
        RefrEx .Image.Handle, .hDC, Data, 1, dbNoGrid
        DontDoEvents = False
    End If
    pctPreview(Index).Line ((w - w / 4) / 2, 0)-((w + w / 4) / 2, h), Comps(Index) * 256& ^ (Index), BF
    pctPreview(Index).Refresh
End With
End Sub

Private Sub CancelButton_Click()
Me.Tag = "C"
Me.Hide
End Sub

Private Sub Form_Load()
LoadSettings
Resr1.LoadCaptions
UpdateValues True
End Sub

Public Sub LoadSettings()
On Error Resume Next
Comps(0) = 128
Comps(1) = 128
Comps(2) = 128

Comps(0) = dbGetSettingEx("Effects\Graph\Gamma", "GammaR", vbByte, 128)
Comps(1) = dbGetSettingEx("Effects\Graph\Gamma", "GammaG", vbByte, 128)
Comps(2) = dbGetSettingEx("Effects\Graph\Gamma", "GammaB", vbByte, 128)
chkUn.Value = IIf(dbGetSettingEx("Effects\Graph\Gamma", "Uncorrection", vbBoolean, False), vbChecked, vbUnchecked)
chkAll.Value = IIf(dbGetSettingEx("Effects\Graph\Gamma", "AllTogether", vbBoolean, True), vbChecked, vbUnchecked)
End Sub

Public Sub UpdateValues(Optional ByVal DrawBG As Boolean = False)
Dim i As Long
For i = 0 To 2
    If Comps(i) > 254 Then Comps(i) = 254
    If Comps(i) < 1 Then Comps(i) = 1
    nmbComp(i).Value = Comps(i)
    pctPreview_Init i, DrawBG
Next i
End Sub

Private Sub Form_Paint()
On Error Resume Next
Me.PaintPicture gBackPicture, 0, 0, ScaleWidth, ScaleHeight
End Sub

Private Sub nmbComp_Change(Index As Integer)
nmbComp_InputChange Index
End Sub

Private Sub nmbComp_InputChange(Index As Integer)
Dim i As Long
Static bLock As Boolean
If bLock Then Exit Sub
On Error Resume Next
If chkAll.Value = 1 Then
    bLock = True
    For i = 0 To 2
        Comps(i) = nmbComp(Index).Value
        If i <> Index Then
            nmbComp(i).Value = Comps(i)
        End If
        pctPreview_Init i
    Next i
    bLock = False
Else
    Comps(Index) = nmbComp(Index).Value
    pctPreview_Init Index
End If
End Sub

Friend Sub SetComps(ByRef rgbQ As RGBQUAD)
Comps(0) = rgbQ.rgbRed
Comps(1) = rgbQ.rgbGreen
Comps(2) = rgbQ.rgbBlue
End Sub

Friend Sub GetComps(ByRef rgbQ As RGBQUAD)
rgbQ.rgbRed = Comps(0)
rgbQ.rgbGreen = Comps(1)
rgbQ.rgbBlue = Comps(2)
End Sub

Public Sub SaveSettings()
dbSaveSettingEx "Effects\Graph\Gamma", "GammaR", Comps(0)
dbSaveSettingEx "Effects\Graph\Gamma", "GammaG", Comps(1)
dbSaveSettingEx "Effects\Graph\Gamma", "GammaB", Comps(2)
dbSaveSettingEx "Effects\Graph\Gamma", "UnCorrection", CBool(chkUn.Value)
dbSaveSettingEx "Effects\Graph\Gamma", "AllTogether", CBool(chkAll.Value)
End Sub

Private Sub OkButton_Click()
Dim i As Long
On Error GoTo eh
For i = 0 To 2
    Comps(i) = nmbComp(i).Value
Next i
On Error GoTo 0
SaveSettings
Me.Tag = ""
Me.Hide
Exit Sub
eh:
MsgBox Err.Description, vbCritical
End Sub
