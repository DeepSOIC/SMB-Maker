VERSION 5.00
Begin VB.Form frmLine 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Настройки линии"
   ClientHeight    =   7110
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8325
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7110
   ScaleWidth      =   8325
   StartUpPosition =   1  'CenterOwner
   Begin SMBMaker.ctlResourcizator Resr1 
      Left            =   1320
      Top             =   6465
      _ExtentX        =   1164
      _ExtentY        =   741
      ResID           =   9916
   End
   Begin SMBMaker.dbFrame dbFrame3 
      Height          =   1095
      Left            =   225
      TabIndex        =   17
      Top             =   5130
      Width           =   7860
      _ExtentX        =   13864
      _ExtentY        =   1931
      Caption         =   "Заливка"
      EAC             =   0   'False
      Begin SMBMaker.dbButton dbButton1 
         Height          =   555
         Left            =   1185
         TabIndex        =   18
         Top             =   270
         Width           =   5355
         _ExtentX        =   9446
         _ExtentY        =   979
         MouseIcon       =   "frmLine.frx":0000
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Others          =   $"frmLine.frx":001C
         OthersPresent   =   -1  'True
      End
   End
   Begin SMBMaker.dbFrame dbFrame2 
      Height          =   3015
      Left            =   225
      TabIndex        =   7
      Top             =   2040
      Width           =   7860
      _ExtentX        =   13864
      _ExtentY        =   5318
      Caption         =   "Форма"
      EAC             =   0   'False
      Begin VB.CheckBox chkColorful 
         BackColor       =   &H00FF80FF&
         Caption         =   "В цвете"
         Height          =   330
         Left            =   5415
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   2535
         Width           =   1590
      End
      Begin VB.PictureBox picPreview 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         Height          =   2160
         Left            =   4860
         ScaleHeight     =   140
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   167
         TabIndex        =   14
         ToolTipText     =   "Здесь отображается только форма."
         Top             =   405
         Width           =   2565
      End
      Begin SMBMaker.ctlNumBox nmbAA 
         Height          =   495
         Left            =   150
         TabIndex        =   8
         Top             =   510
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   873
         Min             =   -20
         Max             =   20
         NumType         =   5
         HorzMode        =   0   'False
         NLn             =   0
      End
      Begin SMBMaker.ctlNumBox nmbW 
         Height          =   495
         Left            =   2325
         TabIndex        =   9
         Top             =   510
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   873
         Value           =   0,004
         Min             =   0.004
         Max             =   60
         HorzMode        =   0   'False
         EditName        =   "Толщина линии, в пикселях."
         NLn             =   0,75
         NativeValues    =   "1 пиксель|1|3 пикселя|3"
         NativeValuesResID=   2442
      End
      Begin SMBMaker.ctlNumBox nmbWeightsRelation 
         Height          =   495
         Left            =   2325
         TabIndex        =   12
         Top             =   1455
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   873
         Min             =   -1
         Max             =   1
         HorzMode        =   0   'False
         EditName        =   $"frmLine.frx":007B
         NLn             =   0
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Просмотр:"
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
         Left            =   5025
         TabIndex        =   15
         Top             =   165
         Width           =   1725
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Распределение толщины:"
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
         Left            =   2355
         TabIndex        =   13
         Top             =   1200
         Width           =   1725
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Сглаживание:"
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
         Left            =   180
         TabIndex        =   11
         Top             =   255
         Width           =   1725
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Толщина:"
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
         Left            =   2355
         TabIndex        =   10
         Top             =   255
         Width           =   1725
      End
   End
   Begin SMBMaker.dbFrame dbFrame1 
      Height          =   1800
      Left            =   232
      TabIndex        =   1
      Top             =   135
      Width           =   7860
      _ExtentX        =   13864
      _ExtentY        =   3175
      Caption         =   "Режим построения"
      EAC             =   0   'False
      Begin VB.OptionButton ModeOpt 
         BackColor       =   &H0080FFFF&
         Caption         =   "Обычная"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   0
         Left            =   165
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   225
         Width           =   2190
      End
      Begin VB.OptionButton ModeOpt 
         BackColor       =   &H0080FFFF&
         Caption         =   "Центрально симметричная"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   1
         Left            =   165
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   570
         Width           =   2190
      End
      Begin VB.OptionButton ModeOpt 
         BackColor       =   &H0080FFFF&
         Caption         =   "Перпендикулярная"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   2
         Left            =   165
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   915
         Width           =   2190
      End
      Begin VB.OptionButton ModeOpt 
         BackColor       =   &H0080FFFF&
         Caption         =   "Параллельная"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   3
         Left            =   165
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   1260
         Width           =   2190
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmLine.frx":0146
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1395
         Left            =   2700
         TabIndex        =   16
         Top             =   180
         Width           =   4815
      End
   End
   Begin SMBMaker.dbButton OkButton 
      Default         =   -1  'True
      Height          =   405
      Left            =   4725
      TabIndex        =   0
      Top             =   6345
      Width           =   1590
      _ExtentX        =   2805
      _ExtentY        =   714
      MouseIcon       =   "frmLine.frx":01FF
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Others          =   $"frmLine.frx":021B
      OthersPresent   =   -1  'True
   End
   Begin SMBMaker.dbButton CancelButton 
      Cancel          =   -1  'True
      Height          =   405
      Left            =   6480
      TabIndex        =   6
      Top             =   6345
      Width           =   1590
      _ExtentX        =   2805
      _ExtentY        =   714
      MouseIcon       =   "frmLine.frx":0267
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Others          =   $"frmLine.frx":0283
      OthersPresent   =   -1  'True
   End
End
Attribute VB_Name = "frmLine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim prvFDSC As FadeDesc
Dim pColorful As Boolean
Dim bLock As Boolean

Private Sub CancelButton_Click()
Me.Tag = "C"
Me.Hide
End Sub

Private Sub chkColorful_Click()
pColorful = chkColorful.Value = vbChecked
Update
End Sub

Private Sub dbButton1_Click()
Load Pereliv
With Pereliv
    MainForm.SendFadeDesc prvFDSC
    .Show vbModal
    If .Tag = "" Then
        MainForm.ExtractFadeDesc prvFDSC
    End If
End With
Unload Pereliv
Update
End Sub

Private Sub Form_Load()
Resr1.LoadCaptions
End Sub

Private Sub Form_Paint()
Me.PaintPicture gBackPicture, 0, 0, Me.ScaleWidth, Me.ScaleHeight
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If UnloadMode = 0 Then
    Cancel = True
    CancelButton_Click
End If
End Sub

Private Sub nmbAA_InputChange()
Update
End Sub

Private Sub nmbWeightsRelation_InputChange()
Update
End Sub

Private Sub OkButton_Click()
Me.Tag = ""
Me.Hide
End Sub

Friend Sub SetProps(ByRef Opts As LineSettings, ByRef FadeDsc As FadeDesc)
bLock = True
On Error GoTo eh
ModeOpt(Opts.GeoMode).Value = True
prvFDSC = FadeDsc
nmbW.Value = Opts.Weight
On Error Resume Next
nmbWeightsRelation.Value = (Opts.RelWeight2 - Opts.RelWeight1) / (Opts.RelWeight2 + Opts.RelWeight1)
On Error GoTo eh
nmbAA.Value = -FactorToDB(Opts.AntiAliasing)
chkColorful_Click
bLock = False
Update
Exit Sub
eh:
bLock = False
ErrRaise
End Sub

Friend Function GetProps(ByRef Opts As LineSettings, ByRef FadeDsc As FadeDesc) As Long
Dim i As Long, Res As Long
Res = 0
For i = 0 To ModeOpt.UBound
    If ModeOpt(i).Value Then
        Res = i
        Exit For
    End If
Next i
Opts.GeoMode = Res

FadeDsc = prvFDSC
Opts.Weight = nmbW.Value
Opts.RelWeight1 = 1 - nmbWeightsRelation.Value
Opts.RelWeight2 = 1 + nmbWeightsRelation.Value
Opts.AntiAliasing = dBtoFactor(-nmbAA.Value)
End Function

Private Sub nmbW_InputChange()
Dim t As Single
On Error Resume Next
t = nmbW.Value
OkButton.Enabled = Err.Number = 0
Update
End Sub

Public Sub Update()
Dim Pic() As Long
Dim RGBPic() As RGBQUAD
Dim w As Long, h As Long
Dim CP As ComplexPixels
Dim Vtx1 As vtVertex, Vtx2 As vtVertex
Dim Fade As FadeDesc
If bLock Then
  Cls
  Exit Sub
End If
w = picPreview.ScaleWidth
h = picPreview.ScaleHeight
ReDim Pic(0 To w - 1, 0 To h - 1)

With Vtx1
    .Weight = nmbW.Value * (1 - nmbWeightsRelation.Value)
    .Color = IIf(pColorful, MainForm.GetACol(1), &HFFFFFF)
    .X = Int(w * 0.3)
    .Y = h \ 2
End With
CopyMemory Vtx2, Vtx1, Len(Vtx1)
With Vtx2
    .Y = .Y + 3
    .Weight = nmbW.Value * (1 + nmbWeightsRelation.Value)
    .Color = IIf(pColorful, MainForm.GetACol(2), &HFFFFFF)
    .X = w - 1 - Int(w * 0.3)
End With

AntiAliasingSharpness = dBtoFactor(-nmbAA.Value)

If pColorful Then
  Fade = prvFDSC
Else
  With Fade
      .FCount = 0
      .Power = 0.5
      .Mode = dbFLinear
      .Offset = 0
  End With
End If

With CP
    ReDim .Elements(0 To 0)
End With
DrawingEngine.pntGradientLineHQ Vtx1, Vtx2, Fade, CP.Elements(0).Pixels, CP.Elements(0).nPixels

DrawingEngine.RangeCheckComplexPixels w, h, CP
SwapArys AryPtr(Pic), AryPtr(RGBPic)
DrawingEngine.DrawPixels RGBPic, CP.Elements(0).Pixels, CP.Elements(0).nPixels
SwapArys AryPtr(Pic), AryPtr(RGBPic)

RefrEx picPreview.Image.Handle, picPreview.hDC, Pic

picPreview.Refresh
End Sub
