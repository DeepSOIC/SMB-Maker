VERSION 5.00
Begin VB.Form frmPrint 
   BackColor       =   &H00E3DFE0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Print"
   ClientHeight    =   5819
   ClientLeft      =   2761
   ClientTop       =   3751
   ClientWidth     =   7029
   Icon            =   "frmPrint.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   529
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   639
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.CheckBox FixPr 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Fix proportion"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.07
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   4155
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2775
      Width           =   2610
   End
   Begin SMBMaker.dbFrame dbFrame1 
      Height          =   2145
      Left            =   4050
      TabIndex        =   13
      Top             =   1035
      Width           =   2805
      _ExtentX        =   4958
      _ExtentY        =   3780
      Caption         =   "Position"
      ResID           =   2510
      EAC             =   0   'False
      Begin SMBMaker.ctlNumBox nmbY 
         Height          =   315
         Left            =   300
         TabIndex        =   22
         Top             =   540
         Width           =   1365
         _ExtentX        =   2195
         _ExtentY        =   792
         Min             =   -10000
         Max             =   10000
         HorzMode        =   -1  'True
         NLn             =   0
      End
      Begin SMBMaker.ctlNumBox nmbX 
         Height          =   300
         Left            =   300
         TabIndex        =   23
         Top             =   225
         Width           =   1365
         _ExtentX        =   2418
         _ExtentY        =   528
         Min             =   -10000
         Max             =   10000
         HorzMode        =   -1  'True
         NLn             =   0
      End
      Begin SMBMaker.ctlNumBox nmbW 
         Height          =   300
         Left            =   300
         TabIndex        =   24
         Top             =   1050
         Width           =   1365
         _ExtentX        =   2418
         _ExtentY        =   528
         Value           =   100
         Max             =   10000
         HorzMode        =   -1  'True
         NLn             =   0
      End
      Begin SMBMaker.ctlNumBox nmbH 
         Height          =   315
         Left            =   300
         TabIndex        =   25
         Top             =   1365
         Width           =   1365
         _ExtentX        =   2418
         _ExtentY        =   549
         Value           =   100
         Max             =   10000
         HorzMode        =   -1  'True
         NLn             =   0
      End
      Begin VB.ComboBox cmbH 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.07
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IntegralHeight  =   0   'False
         ItemData        =   "frmPrint.frx":1272
         Left            =   1680
         List            =   "frmPrint.frx":1288
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   1365
         Width           =   1065
      End
      Begin VB.ComboBox cmbW 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.07
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IntegralHeight  =   0   'False
         ItemData        =   "frmPrint.frx":12C3
         Left            =   1680
         List            =   "frmPrint.frx":12D9
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   1035
         Width           =   1065
      End
      Begin VB.ComboBox cmbY 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.07
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IntegralHeight  =   0   'False
         ItemData        =   "frmPrint.frx":1314
         Left            =   1680
         List            =   "frmPrint.frx":132A
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   540
         Width           =   1065
      End
      Begin VB.ComboBox cmbX 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.07
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IntegralHeight  =   0   'False
         ItemData        =   "frmPrint.frx":1365
         Left            =   1680
         List            =   "frmPrint.frx":137B
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   210
         Width           =   1065
      End
      Begin VB.Label lblH 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "H:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.07
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   60
         TabIndex        =   20
         Top             =   1425
         Width           =   165
      End
      Begin VB.Label lblW 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "W:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.07
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   30
         TabIndex        =   18
         Top             =   1110
         Width           =   210
      End
      Begin VB.Label lblY 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Y:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.07
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   60
         TabIndex        =   16
         Top             =   585
         Width           =   150
      End
      Begin VB.Label lblX 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "X:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.07
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   60
         TabIndex        =   14
         Top             =   270
         Width           =   150
      End
   End
   Begin SMBMaker.dbFrame Frame1 
      Height          =   1035
      Left            =   4440
      TabIndex        =   12
      Top             =   3945
      Width           =   1920
      _ExtentX        =   0
      _ExtentY        =   0
      Caption         =   "Orientation"
      BackColor       =   14933984
      EAC             =   0   'False
      Begin VB.OptionButton Orient 
         BackColor       =   &H0080FFFF&
         Caption         =   "Landscape"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.07
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   150
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   570
         Width           =   1575
      End
      Begin VB.OptionButton Orient 
         BackColor       =   &H0080FFFF&
         Caption         =   "Book"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.07
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   135
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   255
         Value           =   -1  'True
         Width           =   1575
      End
   End
   Begin VB.TextBox Copies 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.07
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   5130
      TabIndex        =   3
      Text            =   "1"
      Top             =   5325
      Width           =   855
   End
   Begin VB.PictureBox Page 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   5550
      Left            =   60
      ScaleHeight     =   501
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   353
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   75
      Width           =   3924
      Begin VB.PictureBox Heighter 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   88
         Left            =   945
         Picture         =   "frmPrint.frx":13B6
         ScaleHeight     =   8
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   8
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   2115
         Width           =   88
      End
      Begin VB.PictureBox Widther 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   88
         Left            =   1605
         Picture         =   "frmPrint.frx":15F0
         ScaleHeight     =   8
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   8
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   1245
         Width           =   88
      End
      Begin VB.PictureBox Sizer 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   88
         Left            =   1575
         Picture         =   "frmPrint.frx":182A
         ScaleHeight     =   8
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   8
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   2085
         Width           =   88
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ClipControls    =   0   'False
         Height          =   1440
         Left            =   465
         ScaleHeight     =   131
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   101
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   645
         Width           =   1110
      End
   End
   Begin SMBMaker.dbButton btnResolution 
      Height          =   585
      Left            =   4605
      TabIndex        =   26
      Top             =   3210
      Width           =   1665
      _ExtentX        =   2946
      _ExtentY        =   1036
      MouseIcon       =   "frmPrint.frx":1A64
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.064
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Others          =   $"frmPrint.frx":1A80
      OthersPresent   =   -1  'True
   End
   Begin SMBMaker.dbButton CancelButton 
      Cancel          =   -1  'True
      Height          =   360
      Left            =   5565
      TabIndex        =   5
      Top             =   585
      Width           =   1335
      _ExtentX        =   2357
      _ExtentY        =   630
      MouseIcon       =   "frmPrint.frx":1ADB
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.064
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Others          =   $"frmPrint.frx":1AF7
      OthersPresent   =   -1  'True
   End
   Begin SMBMaker.dbButton OKButton 
      Default         =   -1  'True
      Height          =   360
      Left            =   5565
      TabIndex        =   4
      Top             =   105
      Width           =   1335
      _ExtentX        =   2357
      _ExtentY        =   630
      MouseIcon       =   "frmPrint.frx":1B3D
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.064
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Others          =   $"frmPrint.frx":1B59
      OthersPresent   =   -1  'True
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00E3DFE0&
      BackStyle       =   0  'Transparent
      Caption         =   "Copies:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.07
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4065
      TabIndex        =   8
      Top             =   5025
      Width           =   2985
   End
End
Attribute VB_Name = "frmPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mx As Single, my As Single, PrW As Single, PrH As Single, NoResize As Boolean

Dim OutX As Long, OutY As Long
Dim OutW As Long, OutH As Long
Dim OrigW As Long, OrigH As Long
Dim pSWidth As Long, pSHeight As Long
Dim tLock As Boolean
Dim ptrData As Long
Dim DataW As Long, DataH As Long
Option Explicit

Private Sub btnResolution_Click()
Dim rsl As Long
On Error GoTo eh
rsl = dbGetSettingEx("Printing", "ImageResolution", vbLong, 300&)
EditNumber rsl, 2508, 50, 4800 '"Please input the resolution (in dpi) for printing this picture. Use this when you have scanned the image, and you know it's actual resolution. This will make the image match it's original size."
dbSaveSettingEx "Printing", "ImageResolution", rsl
nmbW.Value = Printer.ScaleX(DataW / rsl, vbInches, WSM)
nmbH.Value = Printer.ScaleY(DataH / rsl, vbInches, HSM)
Exit Sub
eh:
MsgError
End Sub

Private Sub cmbH_Click()
If tLock Then Exit Sub
UpdateTexts
End Sub

Private Sub cmbW_Click()
If tLock Then Exit Sub
UpdateTexts
End Sub

Private Sub cmbX_Click()
If tLock Then Exit Sub
UpdateTexts
End Sub

Private Sub cmbY_Click()
If tLock Then Exit Sub
UpdateTexts
End Sub

Private Sub Form_Paint()
Me.PaintPicture LoadResPicture(RES_BackPicture, vbResBitmap), 0, 0, Me.ScaleWidth, Me.ScaleHeight
End Sub

Private Sub CancelButton_Click()
Me.Hide
End Sub

Private Sub Copies_Change()
With Copies
    If Not (.Text = "") Then
        .Text = CStr(Int(Val(.Text)))
        If Val(.Text) < 0 Then
            .Text = ""
            vtBeep
        End If
        If Val(.Text) > 200 Then
            .Text = "200"
            vtBeep
        End If
    End If
End With
End Sub

Function dbValidateControls() As Boolean
dbValidateControls = True
With Copies
    If Val(.Text) <= 0 Or Val(.Text) > 256 Then
        .SetFocus
        vtBeep
        dbValidateControls = False
        Exit Function
    End If
End With
If Not (Orient(0).Value) And Not (Orient(1).Value) Then
    Orient(0).SetFocus
    vtBeep
    dbValidateControls = False
End If
End Function

Private Sub FixPr_Click()
If FixPr.Value = 1 Then
    If OutW <> 0 And OutH <> 0 Then
        PrW = OutW
        PrH = OutH
    End If
End If
End Sub

Private Sub UpdateTxtPos()
Dim px As Single, py As Single

'txtX.Text = Trim(Str$(vbscalex))
End Sub

Private Sub Form_Load()
On Error GoTo eh
LoadCaptions
MovePic Picture1.Left, Picture1.Top
Picture1_Resize
LoadSettings

Printer.ScaleMode = vbPixels
pSWidth = Printer.ScaleWidth
pSHeight = Printer.ScaleHeight
OutW = DataW '!!!
OutH = DataH
OutX = 0
OutY = 0
tLock = True
cmbX.ListIndex = 0
cmbY.ListIndex = 0
cmbW.ListIndex = 0
cmbH.ListIndex = 0
tLock = False
UpdatePos
UpdateTexts
Picture1_Resize
Exit Sub
eh:
dbMsgBox "Error while accessing the printer.", vbCritical
End Sub

Private Function RandomColor() As Long
Const Intensity = 64
RandomColor = RGB(255 - Rnd * Intensity, 255 - Rnd * Intensity, 255 - Rnd * Intensity)
End Function

Private Sub DrawPage()
Const Density = 0.2 'dots/pixel
Dim i As Long
Dim w As Long, h As Long
Page.Cls
w = Page.ScaleWidth
h = Page.ScaleHeight
For i = 0 To w * h * Density
    SetPixelV Page.hDC, Rnd * w, Rnd * h, RandomColor
Next i
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If UnloadMode = vbFormControlMenu Then
    Cancel = True
    Me.Hide
End If
End Sub

Private Sub nmbH_InputChange()
nmbH_Change
End Sub

Private Sub nmbW_InputChange()
nmbW_Change
End Sub

Private Sub nmbX_InputChange()
nmbX_Change
End Sub

Private Sub nmbY_InputChange()
nmbY_Change
End Sub

Private Sub OkButton_Click()
Dim i As Integer
If Not (dbValidateControls) Then Exit Sub
On Error GoTo eh

Printer.Copies = Val(Copies.Text)
Printer.PSet (0, 0), vbWhite
'Printer.PaintPicture PP.Image, _
    OutX, _
    OutY, _
    OutW, _
    OutH
vtStretchDIBits Printer.hDC, ptrData, OutX, OutY, OutW, OutH
                
Printer.EndDoc
SaveSettings
Me.Hide
Exit Sub
eh:
If dbMsgBox(grs(1146, "|1", Err.Description), vbCritical Or vbRetryCancel Or vbSystemModal) = vbRetry Then Resume
           '"An error occured during sending to printer. Maybe there's no default printer. Error description:" + Err.Description + "`Error"
End Sub

Private Sub Orient_Click(Index As Integer)
Dim t As Single
If Orient(1).Value Then 'landscape
    'Page.Width = 2775 * 2
    Printer.Orientation = VBRUN.PrinterObjectConstants.vbPRORLandscape
ElseIf Orient(0).Value Then 'portrait
    'Page.Width = 1962 * 2
    Printer.Orientation = VBRUN.PrinterObjectConstants.vbPRORPortrait
End If
Printer.ScaleMode = vbPixels
pSWidth = Printer.ScaleWidth
pSHeight = Printer.ScaleHeight
Page.Height = (pSHeight / pSWidth * Page.ScaleWidth) + 4
UpdatePos
End Sub

Private Sub Page_Resize()
DrawPage
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then mx = X: my = Y
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim t As Single, l As Single
If Button = 1 Then
    l = Picture1.Left + X - mx
    t = Picture1.Top + Y - my
    If l + Picture1.Width < 2 Then l = -Picture1.Width + 2
    If l + 2 > Page.ScaleWidth Then l = Page.ScaleWidth - 2
    If t + Picture1.Height < 2 Then t = 2 - Picture1.Height
    If t + 2 > Page.ScaleHeight Then t = Page.ScaleHeight - 2
    MovePic l, t
    CoordsFromPos OutX, OutY, OutW, OutH
    UpdateTexts
    Page.Refresh
    'Picture1.Refresh
End If
End Sub

Sub TestPos()
Dim l As Single, t As Single
    l = Picture1.Left
    t = Picture1.Top
    If l + Picture1.Width < 2 Then l = -Picture1.Width + 2
    If l + 2 > Page.ScaleWidth Then l = Page.ScaleWidth - 2
    If t + Picture1.Height < 2 Then t = 2 - Picture1.Height
    If t + 2 > Page.ScaleHeight Then t = Page.ScaleHeight - 2
    If Int(l) <> Int(Picture1.Left) Or Int(t) <> Int(Picture1.Top) Then MovePic l, t
End Sub

Sub MovePic(ByVal X As Single, ByVal Y As Single)
'Picture1.Top = Y
Picture1.Move X, Y
'Widther.Top = (Picture1.Height - Widther.Height) / 2 + Y
'Widther.Left = Picture1.Width + X
Widther.Move Picture1.Width + X, (Picture1.Height - Widther.Height) / 2 + Y
'Heighter.Left = (Picture1.Width - Widther.Width) / 2 + X
'Heighter.Top = Picture1.Height + Y
Heighter.Move (Picture1.Width - Widther.Width) / 2 + X, Picture1.Height + Y
'Sizer.Left = Picture1.Width + X
'Sizer.Top = Picture1.Height + Y
Sizer.Move Picture1.Width + X, Picture1.Height + Y
'TestPos
End Sub

Private Sub Heighter_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim px As Single, py As Single
If Button = 1 Then
    With Heighter
        'px = .Left + X - .Width / 2
        py = .Top + Y - .Height / 2
    End With
    'If px - Picture1.Left <= 0 Then px = Picture1.Left + 1
    If py - Picture1.Top <= 0 Then py = Picture1.Top + 1
    'Picture1.Width = px - Picture1.Left
    NoResize = True
    Picture1.Height = py - Picture1.Top
    'If FixPr.Value = 1 Then Picture1.Width = Picture1.Height * PrW / PrH
    NoResize = False
    Picture1_Resize
    MovePic Picture1.Left, Picture1.Top
    CoordsFromPos OutX, OutY, OutW, OutH
    If FixPr.Value Then
        Proportionize OutW, OutH, 2
        UpdatePos
    End If
    UpdateTexts
End If
End Sub

Private Sub Picture1_Paint()
'Picture1.PaintPicture PP.Image, 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight
vtStretchDIBits Picture1.hDC, ptrData, 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight
End Sub

Private Sub Picture1_Resize()
Dim h As Single, w As Single
If NoResize Then Exit Sub
If FixPr Then
    'h = Picture1.Width * PrH / PrW
    'w = Picture1.Height * PrW / PrH
'    If Picture1.Width > w Then
'        NoResize = True
'        Picture1.Width = w
'        NoResize = False
'    Else
'        NoResize = True
'        Picture1.Height = h
'        NoResize = False
'    End If
    MovePic Picture1.Left, Picture1.Top
End If
'Picture1.PaintPicture PP.Image, 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight
Page.Refresh
Picture1.Refresh
End Sub

Private Sub Sizer_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim px As Single, py As Single
Dim pw As Long, ph As Long
If Button = 1 Then
    px = Sizer.Left + X - Sizer.Width / 2
    py = Sizer.Top + Y - Sizer.Height / 2
    If px - Picture1.Left <= 0 Then px = Picture1.Left + 1
    If py - Picture1.Top <= 0 Then py = Picture1.Top + 1
    NoResize = True
    pw = px - Picture1.Left
    NoResize = True
    ph = py - Picture1.Top
    NoResize = False
    'Picture1_Resize
    'MovePic Picture1.Left, Picture1.Top
    CoordsFromPosEx OutX, OutY, OutW, OutH, _
                    Picture1.Left, Picture1.Top, pw, ph
    If FixPr.Value Then
        Proportionize OutW, OutH, 3
    End If
    UpdatePos
    UpdateTexts
End If
End Sub

Private Sub nmbH_Change()
If tLock Then Exit Sub
TextsToCoords OutX, OutY, OutW, OutH
If FixPr.Value Then
    Proportionize OutW, OutH, 2
    UpdateTexts 2
End If
UpdatePos
End Sub

Private Sub nmbW_Change()
If tLock Then Exit Sub
TextsToCoords OutX, OutY, OutW, OutH
If FixPr.Value Then
    Proportionize OutW, OutH, 1
    UpdateTexts 1
End If
UpdatePos
End Sub

Private Sub PosChange()
If tLock Then Exit Sub
TextsToCoords OutX, OutY, OutW, OutH
UpdatePos
End Sub

Private Sub nmbX_Change()
PosChange
End Sub

Private Sub nmbY_Change()
PosChange
End Sub

Private Sub Widther_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim px As Single, py As Single
If Button = 1 Then
    With Widther
        px = .Left + X - .Width / 2
        'py = .Top + Y - .Height / 2
    End With
    If px - Picture1.Left <= 0 Then px = Picture1.Left + 1
    'If py - Picture1.Top <= 0 Then py = Picture1.Top + 1
    NoResize = True
    Picture1.Width = px - Picture1.Left
    'If FixPr.Value = 1 Then Picture1.Height = Picture1.Width * PrH / PrW
    NoResize = False
    Picture1_Resize
'    Picture1.Height = py - Picture1.Top
    MovePic Picture1.Left, Picture1.Top
    CoordsFromPos OutX, OutY, OutW, OutH
    If FixPr.Value Then
        Proportionize OutW, OutH, 1
        UpdatePos
    End If
    UpdateTexts
End If
End Sub

Public Sub LoadCaptions()
Frame1.Caption = GRSF(2153)
Orient(1).Caption = GRSF(2154)
Orient(0).Caption = GRSF(2155)
FixPr.Caption = GRSF(2156)
CancelButton.Caption = GRSF(2157)
OkButton.Caption = GRSF(2158)
Label1.Caption = GRSF(2159)
FillCombo cmbX, GRSF(2509)
FillCombo cmbY, GRSF(2509)
FillCombo cmbW, GRSF(2509)
FillCombo cmbH, GRSF(2509)
'Me.Icon = LoadResPicture(Me.Name, vbResIcon)
End Sub

Private Sub FillCombo(ByRef Cmb As ComboBox, ByRef List As String)
Dim sArr() As String
Dim i As Long
Dim Pos As Long
Cmb.Clear
If Len(List) = 0 Then Exit Sub
sArr = Split(List, vbCrLf)
For i = 0 To UBound(sArr)
    Pos = InStrRev(sArr(i), "=")
    Cmb.AddItem Left$(sArr(i), Pos - 1)
    If Pos > 0 Then
        Cmb.ItemData(Cmb.NewIndex) = Val(Mid$(sArr(i), Pos + 1))
    End If
Next i
End Sub

Public Sub SetData(ByRef Data() As Long)
Dim h As Long, w As Long
Dim k As Double
Dim PrnSW As Long, PrnSH As Long
ptrData = AryPtr(Data)
AryWH ptrData, w, h
If w * h = 0 Then Err.Raise errNewFile, "frmPrint.SetData", "Empty picture - cannot be printed."
DataW = w
DataH = h
On Error GoTo eh
Printer.ScaleMode = 3 'pixel
pSWidth = Printer.ScaleWidth
pSHeight = Printer.ScaleHeight
rsm:
'tppX = Screen.TwipsPerPixelX
'tppY = Screen.TwipsPerPixelY
'    PP.AutoSize = False
'    PP.AutoRedraw = True
'    PP.ScaleMode = 3
'    DataW = (UBound(Data, 1) + 1)
'    DataH = (UBound(Data, 2) + 1)
'    PP.Cls
'    For i = 0 To UBound(Data, 1)
'        For j = 0 To UBound(Data, 2)
'            PP.PSet (j, i), Data(i, j)
'        Next j
'    Next i
'    RefrEx PP.Image.Handle, PP.hDC, Data, 1, dbNoGrid
'    PP.Refresh
    NoResize = True
    'Picture1.Width = w * CLng(Page.ScaleWidth) \ PrnSW
    'Picture1.Height = h * CLng(Page.ScaleHeight) \ PrnSH
    'OutX = 0
    'OutY = 0
    If OrigW = 0 Then
        OutW = w
        OutH = h
    Else
        k = OutW / OrigW
        OutW = k * (w)
        OutH = k * (h)
    End If
    OrigW = w
    OrigH = h
    
    NoResize = False
    UpdatePos
    UpdateTexts
    'MovePic 0 + (Page.ScaleWidth - Picture1.Width) \ 2, (Page.ScaleHeight - Picture1.Height) \ 2
    PrW = OutW
    PrH = OutH
    Picture1_Resize
    Orient_Click IIf(Orient(0).Value, 0, 1)
Exit Sub
Resume
eh:
'MsgBox "Cannot open the printer!"
Err.Raise 12421, "SetData", "Cannot open the printer (" + Err.Description + ")."
'pSHeight = 2000
'pSWidth = 1000
End Sub

Public Sub LoadSettings()
Dim MsgText As String, i As Long, Answ As VbMsgBoxResult
On Error GoTo eh
    MsgText = "Bad printer orientation"
    Orient(dbGetSetting("Options", "PrinterOrient", "0")).Value = True
    MsgText = "Bad boolean value(PrintFixProportions)"
    FixPr.Value = Abs(CBool(dbGetSetting("Options", "PrinterFixProportions", "True")))

Exit Sub
eh:
    Answ = MsgBox(MsgText, vbCritical Or vbAbortRetryIgnore, "Loading")
    If Answ = vbRetry Then
    Resume
    ElseIf Answ = vbIgnore Then
    Resume Next
    ElseIf Answ = vbAbort Then
    End
    Exit Sub
    End If
End Sub

Public Sub SaveSettings()
Dim i As Long
dbSaveSetting "Options", "PrinterOrient", CStr(Abs(Orient(1).Value))
dbSaveSetting "Options", "PrinterFixProportions", CStr(CBool(FixPr.Value = 1))
End Sub



Public Function XSM() As ScaleModeConstants
XSM = cmbX.ItemData(cmbX.ListIndex)
End Function

Public Function YSM() As ScaleModeConstants
YSM = cmbY.ItemData(cmbY.ListIndex)
End Function

Public Function WSM() As ScaleModeConstants
WSM = cmbW.ItemData(cmbW.ListIndex)
End Function

Public Function HSM() As ScaleModeConstants
HSM = cmbH.ItemData(cmbH.ListIndex)
End Function

Public Sub UpdateTexts(Optional ByVal ExcludeIndex As Long = -1)
tLock = True
'first - set min/max of numboxes
nmbX.Max = Abs(Printer.ScaleX(32767, vbPixels, XSM))
nmbX.Min = -nmbX.Max
nmbY.Max = Abs(Printer.ScaleY(32767, vbPixels, YSM))
nmbY.Min = -nmbY.Max
nmbW.Max = Abs(Printer.ScaleX(32767, vbPixels, WSM))
nmbH.Max = Abs(Printer.ScaleY(32767, vbPixels, HSM))
If ExcludeIndex <> 0 Then
    nmbX.Value = Printer.ScaleX(OutX, vbPixels, XSM)
    nmbX.NativeValues = grs(2443, "%1", dbCStr(Printer.ScaleX(Printer.ScaleWidth, Printer.ScaleMode, XSM)))
End If
If ExcludeIndex <> 3 Then
    nmbY.Value = Printer.ScaleY(OutY, vbPixels, YSM)
    nmbY.NativeValues = grs(2444, "%1", dbCStr(Printer.ScaleY(Printer.ScaleHeight, Printer.ScaleMode, YSM)))
End If
If ExcludeIndex <> 1 Then
    nmbW.Value = Printer.ScaleX(OutW, vbPixels, WSM)
    nmbW.NativeValues = grs(2443, "%1", dbCStr(Printer.ScaleX(Printer.ScaleWidth, Printer.ScaleMode, WSM)))
End If
If ExcludeIndex <> 2 Then
    nmbH.Value = Printer.ScaleY(OutH, vbPixels, HSM)
    nmbH.NativeValues = grs(2444, "%1", dbCStr(Printer.ScaleY(Printer.ScaleHeight, Printer.ScaleMode, HSM)))
End If
tLock = False
End Sub

Public Sub TextsToCoords(ByRef px As Long, ByRef py As Long, ByRef pw As Long, ByRef ph As Long)
px = Printer.ScaleX(nmbX.Value, XSM, vbPixels)
py = Printer.ScaleY(nmbY.Value, YSM, vbPixels)
pw = Printer.ScaleX(nmbW.Value, WSM, vbPixels)
ph = Printer.ScaleY(nmbH.Value, HSM, vbPixels)
If pw <= 0 Then pw = 1
If ph <= 0 Then ph = 1
End Sub

Public Sub UpdatePos()
Picture1.Move Picture1.Left, Picture1.Top, _
              CDbl(OutW) * CDbl(Page.ScaleWidth) / CDbl(pSWidth), _
              CDbl(OutH) * CDbl(Page.ScaleHeight) / CDbl(pSHeight)
MovePic CDbl(OutX) * CDbl(Page.ScaleWidth) / CDbl(pSWidth), _
        CDbl(OutY) * CDbl(Page.ScaleHeight) / CDbl(pSHeight)
End Sub

Public Sub CoordsFromPos(ByRef px As Long, ByRef py As Long, ByRef pw As Long, ByRef ph As Long)
px = CDbl(Picture1.Left) * CDbl(pSWidth) / CDbl(Page.ScaleWidth)
py = CDbl(Picture1.Top) * CDbl(pSHeight) / CDbl(Page.ScaleHeight)
pw = CDbl(Picture1.Width) * CDbl(pSWidth) / CDbl(Page.ScaleWidth)
ph = CDbl(Picture1.Height) * CDbl(pSHeight) / CDbl(Page.ScaleHeight)
End Sub

Public Sub CoordsFromPosEx(ByRef px As Long, ByRef py As Long, _
                           ByRef pw As Long, ByRef ph As Long, _
                           ByVal posX As Long, ByVal posY As Long, _
                           ByVal posW As Long, ByVal posH As Long)
px = CDbl(posX) * CDbl(pSWidth) / CDbl(Page.ScaleWidth)
py = CDbl(posY) * CDbl(pSHeight) / CDbl(Page.ScaleHeight)
pw = CDbl(posW) * CDbl(pSWidth) / CDbl(Page.ScaleWidth)
ph = CDbl(posH) * CDbl(pSHeight) / CDbl(Page.ScaleHeight)
End Sub

Public Sub Proportionize(ByRef w As Long, ByRef h As Long, Optional ByVal ByWhat As Long)
Dim nW As Long, nH As Long
Select Case ByWhat
    Case 1 'by width
        h = CDbl(w) * PrH / PrW
    Case 2 'by height
        w = CDbl(h) * PrW / PrH
    Case 3 'by the least
        nH = CDbl(w) * PrH / PrW
        nW = CDbl(h) * PrW / PrH
        If CDbl(h) > w * CDbl(PrH) / PrW Then
            h = nH
        Else
            w = nW
        End If
    Case 4 'by the most
        nH = CDbl(w) * PrH / PrW
        nW = CDbl(h) * PrW / PrH
        If CDbl(h) / w < CDbl(PrH) / PrW Then
            h = nH
        Else
            w = nW
        End If
End Select
End Sub

