VERSION 5.00
Begin VB.Form frmBrush 
   BackColor       =   &H00E3DFE0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Brush"
   ClientHeight    =   4380
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4890
   Icon            =   "frmBrush.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   292
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   326
   StartUpPosition =   1  'CenterOwner
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin SMBMaker.ctlResourcizator Resr1 
      Left            =   3315
      Top             =   1425
      _ExtentX        =   1164
      _ExtentY        =   741
      ResID           =   9912
   End
   Begin VB.PictureBox pPal 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   30
      ScaleHeight     =   17
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   256
      TabIndex        =   9
      Top             =   4110
      Width           =   3840
      Begin VB.Image ImgCursor 
         Enabled         =   0   'False
         Height          =   240
         Left            =   30
         Picture         =   "frmBrush.frx":0442
         Top             =   0
         Width           =   240
      End
   End
   Begin VB.PictureBox View 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   990
      Left            =   30
      ScaleHeight     =   62
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   73
      TabIndex        =   4
      TabStop         =   0   'False
      ToolTipText     =   "The brush pattern. Left-click to switch pixel on and off."
      Top             =   45
      Width           =   1155
   End
   Begin SMBMaker.dbButton btnLoad 
      Height          =   555
      Left            =   3195
      TabIndex        =   3
      Top             =   2700
      Width           =   1605
      _ExtentX        =   2831
      _ExtentY        =   979
      MouseIcon       =   "frmBrush.frx":058C
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Others          =   $"frmBrush.frx":05A8
      OthersPresent   =   -1  'True
   End
   Begin SMBMaker.dbButton btnSave 
      Height          =   555
      Left            =   3195
      TabIndex        =   2
      Top             =   2070
      Width           =   1605
      _ExtentX        =   2831
      _ExtentY        =   979
      MouseIcon       =   "frmBrush.frx":0604
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Others          =   $"frmBrush.frx":0620
      OthersPresent   =   -1  'True
   End
   Begin SMBMaker.dbButton OkButton 
      Height          =   390
      Left            =   3285
      TabIndex        =   0
      Top             =   0
      Width           =   1560
      _ExtentX        =   2752
      _ExtentY        =   688
      MouseIcon       =   "frmBrush.frx":067A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Others          =   $"frmBrush.frx":0696
      OthersPresent   =   -1  'True
   End
   Begin SMBMaker.dbButton CancelButton 
      Cancel          =   -1  'True
      Height          =   390
      Left            =   3285
      TabIndex        =   1
      Top             =   390
      Width           =   1560
      _ExtentX        =   2752
      _ExtentY        =   688
      MouseIcon       =   "frmBrush.frx":06E2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Others          =   $"frmBrush.frx":06FE
      OthersPresent   =   -1  'True
   End
   Begin SMBMaker.dbButton btnClear 
      Height          =   285
      Left            =   285
      TabIndex        =   5
      Top             =   3435
      Width           =   1230
      _ExtentX        =   2170
      _ExtentY        =   503
      MouseIcon       =   "frmBrush.frx":074E
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Others          =   $"frmBrush.frx":076A
      OthersPresent   =   -1  'True
   End
   Begin SMBMaker.dbButton HFlip 
      Height          =   345
      Left            =   30
      TabIndex        =   6
      Top             =   3750
      Width           =   1755
      _ExtentX        =   3096
      _ExtentY        =   609
      MouseIcon       =   "frmBrush.frx":07BA
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Others          =   $"frmBrush.frx":07D6
      OthersPresent   =   -1  'True
   End
   Begin SMBMaker.dbButton VFlip 
      Height          =   345
      Left            =   1785
      TabIndex        =   7
      Top             =   3405
      Width           =   1755
      _ExtentX        =   3096
      _ExtentY        =   609
      MouseIcon       =   "frmBrush.frx":082A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Others          =   $"frmBrush.frx":0846
      OthersPresent   =   -1  'True
   End
   Begin SMBMaker.dbButton btnTurn 
      Height          =   345
      Left            =   1785
      TabIndex        =   8
      Top             =   3750
      Width           =   1755
      _ExtentX        =   3096
      _ExtentY        =   609
      MouseIcon       =   "frmBrush.frx":089A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Others          =   $"frmBrush.frx":08B6
      OthersPresent   =   -1  'True
   End
   Begin VB.Image Sizer 
      Height          =   120
      Left            =   1230
      MousePointer    =   8  'Size NW SE
      Picture         =   "frmBrush.frx":090C
      ToolTipText     =   "Pull me to resize and clear the box below."
      Top             =   1035
      Width           =   120
   End
End
Attribute VB_Name = "frmBrush"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Dot() As Byte
Dim mon As Byte
Dim pCDl As CommonDlg
Dim ACol As Byte

Private Sub btnLoad_Click()
Dim File As String
On Error GoTo eh
'With pCDl
'    .Flags = 0 'FileMustExist
'    .OpenFlags = cdlOFNFileMustExist
'    .Filter = GetDlgFilter(dbBLoad)
'    .FileName = ""
'    .ShowOpen
'    File = .FileName
'    .InitDir = GetDirName(File)
'End With
File = ShowOpenDlg(dbBLoad, Me.hWnd, Purpose:="BRUSH")
On Error GoTo eh2
LoadBrushFile File, Dot
dRefr
eh:
Exit Sub
Resume
eh2:
MsgBox Err.Description, vbCritical, "Error"
End Sub

Private Sub btnSave_Click()
Dim File As String
On Error GoTo eh
'With pCDl
'    .Flags = 0
'    .OpenFlags = cdlOFNFileMustExist
'    .Filter = GetDlgFilter(dbBSave)
'    .FileName = ""
'    .ShowSave
'    File = .FileName
'    .InitDir = GetDirName(File)
'End With
File = ShowSaveDlg(dbBSave, Me.hWnd, Purpose:="BRUSH")
On Error GoTo eh2
SaveBrushToFile Dot, File

eh:
Exit Sub
eh2:
MsgBox Err.Description, vbCritical, "Error"
End Sub

Private Sub CancelButton_Click()
Me.Tag = "C"
Me.Hide
End Sub

Private Sub Form_Initialize()
Set pCDl = New CommonDlg
pCDl.CancelError = True
Resr1.LoadCaptions
End Sub

Private Sub Form_Load()
ACol = 255
UpdateCursorPos ACol
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If UnloadMode = vbFormControlMenu Then
    Cancel = True
    Me.Tag = "C"
    Me.Hide
End If
End Sub

Private Sub OkButton_Click()
Me.Tag = ""
Me.Hide
End Sub

Private Sub Form_Paint()
Me.PaintPicture LoadResPicture(RES_BackPicture, vbResBitmap), 0, 0, Me.ScaleWidth, Me.ScaleHeight
End Sub

Private Sub Picture1_Click()

End Sub

Private Sub pPal_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Y < 0 Or Y > pPal.ScaleHeight - 1 Then
    UpdateCursorPos ACol
    Exit Sub
End If
If X < 0 Then X = 0
If X > 255 Then X = 255
If Button > 0 Then
    UpdateCursorPos X
End If
pPal.ToolTipText = CStr(X)
End Sub

Private Sub pPal_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Not (Y < 0 Or Y > pPal.ScaleHeight - 1) Then
    If X < 0 Then X = 0
    If X > 255 Then X = 255

    ACol = X
End If
UpdateCursorPos ACol
End Sub

Private Sub pPal_Paint()
Dim i As Long, h As Long
h = pPal.ScaleHeight
For i = 0 To 255
    pPal.Line (i, 0)-Step(0, h), i * &H10101
Next i

End Sub

Private Sub UpdateCursorPos(ByVal X As Long)
ImgCursor.Move X - 3, 0
End Sub

Private Sub Sizer_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button > 0 Then
X = X \ Screen.TwipsPerPixelX
Y = Y \ Screen.TwipsPerPixelY
X = X - Sizer.Width \ 2 + Sizer.Left
Y = Y - Sizer.Height \ 2 + Sizer.Top
Resize Round(((X - View.Left) - 4) / 8), Round(((Y - View.Top) - 4) / 8)
End If
End Sub

Sub Resize(w As Long, h As Long, Optional ByVal PreserveValues As Boolean = True)
'If h <= 0 Or w <= 0 Then Exit Sub
'If h > 25 Or w > 25 Then Exit Sub
Dim tmpData() As Byte
Dim i As Long, j As Long
If h <= 0 Then h = 1
If w <= 0 Then w = 1
If h > 25 Then h = 25
If w > 25 Then w = 25
On Error GoTo eh
If PreserveValues Then
    tmpData = Dot
End If
red:
ReDim Dot(0 To w - 1, 0 To h - 1)
If PreserveValues Then
    For i = 0 To Min(h - 1, UBound(tmpData, 2))
        For j = 0 To Min(w - 1, UBound(tmpData, 1))
            Dot(j, i) = tmpData(j, i)
        Next j
    Next i
End If
On Error GoTo 0
View.Move View.Left, View.Top, (w * 8 + 4), (h * 8 + 4)
'View.Width = (w * 8 + 4) * Screen.TwipsPerPixelX
'View.Height = (h * 8 + 4) * Screen.TwipsPerPixelY
Sizer.Move View.Width + View.Left, View.Height + View.Top
'Sizer.Top = View.Height + View.Top
'Sizer.Left = View.Width + View.Left
If PreserveValues Then
    dRefr
Else
    View.Cls
End If
Exit Sub
eh:
    PreserveValues = False
    Resume red

'If h <= 0 Then h = 1
'If h > 25 Then h = 25
'If w <= 0 Then w = 1
'If w > 25 Then w = 25
'ReDim Dot(w - 1, h - 1)
'View.Width = (w * 8 + 4) * Screen.TwipsPerPixelX
'View.Height = (h * 8 + 4) * Screen.TwipsPerPixelY
'Sizer.Top = View.Height + View.Top
'Sizer.Left = View.Width + View.Left
'View.Cls
End Sub

Private Sub View_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Long, j As Long, ic As Long, jc As Long
Dim InRgn As Boolean
i = Y \ 8
j = X \ 8
InRgn = Not (i < 0 Or j < 0 Or i > UBound(Dot, 2) Or j > UBound(Dot, 1))

If Button = 1 And InRgn Then
    If mon = 1 Then
        Dot(j, i) = ACol
    ElseIf mon = 2 Then
        Dot(j, i) = 0
    ElseIf mon = 0 Then
        If Dot(j, i) = 0 Then
            Dot(j, i) = ACol
            mon = 1
        Else
            Dot(j, i) = 0
            mon = 2
        End If
    End If
    
    View.Line (j * 8, i * 8)-((j + 1) * 8 - 1, (i + 1) * 8 - 1), Dot(j, i) * &H10101, BF
ElseIf Button = 4 Then
    If InRgn Then
        UpdateCursorPos Dot(i, j)
    Else
        UpdateCursorPos ACol
    End If
End If
End Sub

Function cTmp(t As Boolean) As Long
Select Case t
    Case False
        cTmp = RGB(255, 255, 255)
    Case True
        cTmp = RGB(0, 0, 255)
End Select
End Function

Private Sub View_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    View_MouseDown Button, Shift, X, Y
End Sub

Private Sub View_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Long, j As Long
Dim InRgn As Boolean
i = Y \ 8
j = X \ 8
InRgn = Not i < 0 Or j < 0 Or i > UBound(Dot, 2) Or j > UBound(Dot, 1)
mon = 0
If Button = 4 And InRgn Then
    ACol = Dot(j, i)
    UpdateCursorPos ACol
End If
End Sub

Sub GetBrush(ByRef dbBrush() As Byte)
dbBrush = Dot
End Sub

Private Sub dRefr()
Dim i As Long, j As Long, tmp As Long, w As Long, h As Long
w = UBound(Dot, 1) + 1
h = UBound(Dot, 2) + 1
If h > 25 Or w > 25 Then
    Resize w, h
    Exit Sub
End If
View.Move View.Left, View.Top, (w * 8 + 4), (h * 8 + 4)
Sizer.Move View.Left + View.Width, View.Top + View.Height
View.Cls
For i = 0 To UBound(Dot, 2)
    For j = 0 To UBound(Dot, 1)
        View.Line (j * 8, i * 8)-((j + 1) * 8 - 1, (i + 1) * 8 - 1), Dot(j, i) * &H10101, BF
    Next j
Next i
End Sub

Sub SetBrush(ByRef dbBrush() As Byte)
Erase Dot
Resize UBound(dbBrush, 1) + 1, UBound(dbBrush, 2) + 1
Dot = dbBrush
dRefr
End Sub
'
'Sub LoadCaptions()
''OkButton.Caption = GRSF(2178)
''CancelButton.Caption = GRSF(2179)
''View.ToolTipText = GRSF(2180)
''Sizer.ToolTipText = GRSF(2181)
''Me.Caption = GRSF(2182)
''btnSave.Caption = GRSF(2199)
''btnLoad.Caption = GRSF(2198)
''Me.Icon = LoadResPicture(Me.Name, vbResIcon)
'End Sub

Private Sub btnClear_Click()
ReDim Dot(0 To UBound(Dot, 1), 0 To UBound(Dot, 2))
View.Cls
End Sub

Private Sub VFlip_Click()
Dim w As Long, h As Long
Dim i As Long, j As Long
Dim tmpDot() As Byte
tmpDot = Dot
w = UBound(Dot, 1)
h = UBound(Dot, 2)

For i = 0 To h
    For j = 0 To w
        Dot(j, i) = tmpDot(j, h - i)
    Next j
Next i
Erase tmpDot
dRefr
End Sub

Private Sub HFlip_Click()
Dim w As Long, h As Long
Dim i As Long, j As Long
Dim tmpDot() As Byte
tmpDot = Dot
w = UBound(Dot, 1)
h = UBound(Dot, 2)

For i = 0 To h
    For j = 0 To w
        Dot(j, i) = tmpDot(w - j, i)
    Next j
Next i
Erase tmpDot
dRefr
End Sub

Private Sub btnTurn_Click()
Dim i As Long, j As Long, tmpDot() As Byte, w As Long, h As Long
tmpDot = Dot
w = UBound(Dot, 1)
h = UBound(Dot, 2)
Resize h + 1, w + 1
For i = 0 To UBound(Dot, 2)
    For j = 0 To UBound(Dot, 1)
        Dot(j, i) = tmpDot(w - i, j)
    Next j
Next i
dRefr
End Sub
