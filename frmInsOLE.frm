VERSION 5.00
Begin VB.Form frmInsOLE 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Formula (OLE object)"
   ClientHeight    =   3870
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   7095
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   258
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   473
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin SMBMaker.ctlResourcizator Resr1 
      Left            =   1065
      Top             =   1980
      _ExtentX        =   1164
      _ExtentY        =   741
      ResID           =   9903
   End
   Begin VB.Timer tmrShow 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   570
      Top             =   1980
   End
   Begin SMBMaker.ctlNumBox nmbFS 
      Height          =   525
      Left            =   5985
      TabIndex        =   6
      Top             =   1245
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   926
      Value           =   2
      Min             =   0.1
      Max             =   20
      NumType         =   5
      HorzMode        =   0   'False
      EditName        =   "Zoom factor of the formula."
   End
   Begin VB.Timer Timer1 
      Interval        =   2000
      Left            =   120
      Top             =   1995
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      Height          =   1320
      Left            =   60
      ScaleHeight     =   1260
      ScaleWidth      =   6900
      TabIndex        =   1
      ToolTipText     =   "Click to refresh"
      Top             =   2505
      Width           =   6960
      Begin SMBMaker.dbButton dbButton1 
         Height          =   465
         Left            =   5700
         TabIndex        =   7
         Top             =   15
         Visible         =   0   'False
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   820
         MouseIcon       =   "frmInsOLE.frx":0000
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Others          =   $"frmInsOLE.frx":001C
         OthersPresent   =   -1  'True
      End
   End
   Begin SMBMaker.dbButton CancelButton 
      Cancel          =   -1  'True
      Height          =   360
      Left            =   5865
      TabIndex        =   5
      Top             =   510
      Width           =   1230
      _ExtentX        =   2170
      _ExtentY        =   635
      MouseIcon       =   "frmInsOLE.frx":006E
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Others          =   $"frmInsOLE.frx":008A
      OthersPresent   =   -1  'True
   End
   Begin SMBMaker.dbButton OkButton 
      Default         =   -1  'True
      Height          =   360
      Left            =   5865
      TabIndex        =   4
      Top             =   120
      Width           =   1230
      _ExtentX        =   2170
      _ExtentY        =   635
      MouseIcon       =   "frmInsOLE.frx":00DA
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Others          =   $"frmInsOLE.frx":00F6
      OthersPresent   =   -1  'True
   End
   Begin SMBMaker.dbButton btnSelObj 
      Height          =   570
      Left            =   5955
      TabIndex        =   3
      Top             =   1845
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   1005
      MouseIcon       =   "frmInsOLE.frx":0142
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Others          =   $"frmInsOLE.frx":015E
      OthersPresent   =   -1  'True
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Zoom:"
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
      Left            =   5985
      TabIndex        =   2
      Top             =   1035
      Width           =   1065
   End
   Begin VB.OLE OLE1 
      AutoActivate    =   0  'Manual
      Class           =   "Equation.DSMT4"
      Height          =   600
      HostName        =   "SMB Maker"
      Left            =   45
      OleObjectBlob   =   "frmInsOLE.frx":01B9
      OLETypeAllowed  =   1  'Embedded
      SizeMode        =   2  'AutoSize
      TabIndex        =   0
      Top             =   45
      UpdateOptions   =   2  'Manual
      Width           =   1500
   End
End
Attribute VB_Name = "frmInsOLE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim bLock As Boolean

Private Sub btnSelObj_Click()
'OLE1.CreateEmbed ""
On Error GoTo eh
OLE1.InsertObjDlg
Exit Sub
eh:
MsgError
End Sub

Private Sub CancelButton_Click()
Me.Tag = "C"
Me.Hide
End Sub

Private Sub dbButton1_Click()
Dim w As Long, h As Long
Dim k As Double
Dim Pct As IPictureDisp
'On Error Resume Next
'OLE1.DoVerb -8
'OLE1.ObjectVerbsCount
On Error GoTo eh
If bLock Then Exit Sub
k = nmbFS.Value
OLE1.Update
OLE1.DoVerb 0
Picture1.AutoRedraw = True
Set Pct = OLE1.Picture
'Pct.Handle = OLE1.Picture.Handle
Picture1.Cls
w = ScaleX(Pct.Width, vbHimetric, Picture1.ScaleMode)
h = ScaleY(Pct.Height, vbHimetric, Picture1.ScaleMode)
Picture1.PaintPicture Pct, 0, 0, w * k, h * k
OLE1.Close
Exit Sub
eh:
MsgError
End Sub

Private Sub Form_Load()
'Me.FontSize = 40
'OLE1.Format
Resr1.LoadCaptions
LoadSettings
Timer1.Enabled = True
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

Private Sub nmbFS_Change()
If bLock Then Exit Sub
UpdatePreview_Fast
tmrShow.Enabled = False
tmrShow.Enabled = True
End Sub

Public Sub UpdatePreview_Fast()
Dim hdcOLE As Long
Dim Rct As RECT
Dim w As Long, h As Long
Dim wf As Long, hf As Long
Dim hWnd As Long
Dim Ret As Long
'Ret = GetClientRect(hWnd, Rct)
'If Ret = 0 Then Err.Raise 2345, , "GetClientRect failed. GetLastError=" + CStr(Err.LastDllError)
Rct.Left = OLE1.Left + 2
Rct.Top = OLE1.Top + 2
w = OLE1.Width - 4
h = OLE1.Height - 4
'w = Rct.Right - Rct.Left
'h = Rct.Bottom - Rct.Top
hdcOLE = Me.hDC
If hdcOLE = 0 Then Err.Raise 2345, , "GetDC failed!"
On Error GoTo eh
wf = w * nmbFS
hf = h * nmbFS
Picture1.Line (0, 0)-(Picture1.ScaleWidth, Picture1.ScaleHeight), Picture1.BackColor, BF
If w > 0 And h > 0 Then
  If StretchBlt(Picture1.hDC, 0, 0, wf, hf, hdcOLE, Rct.Left, Rct.Top, w, h, vbSrcCopy) = 0 Then
    Err.Raise 2345, , "StretchBlt failed. GetLastError=" + CStr(Err.LastDllError)
  End If
End If
'ReleaseDC hWnd, hdcOLE
Picture1.Refresh

Exit Sub
eh:
PushError
ReleaseDC hWnd, hdcOLE
PopError
ErrRaise
End Sub

Private Sub nmbFS_InputChange()
nmbFS_Change
End Sub

Private Sub OkButton_Click()
Dim w As Long, h As Long
Dim Pct As IPictureDisp
Dim k As Double
On Error GoTo eh
MainForm.TempBox.ForeColor = vbBlack
MainForm.TempBox.BackColor = vbWhite
'MainForm.TempBox.FontSize = dbVal(txtFS.Text, vbSingle)
MainForm.TempBox.Cls
'MainForm.TempBox.AutoSize = True
OLE1.SizeMode = vbOLESizeAutoSize
OLE1.Update
OLE1.DoVerb
Set Pct = OLE1.Picture
k = nmbFS.Value
w = ScaleX(Pct.Width, vbHimetric, vbPixels)
h = ScaleY(Pct.Height, vbHimetric, vbPixels)
'Debug.Print OLE1.DataText
With MainForm.TempBox
    .Width = w * k
    .Height = h * k
    .Cls
    .PaintPicture Pct, 0, 0, w * k, h * k
End With
SaveSettings
OLE1.Close
Me.Tag = ""
Me.Hide
Exit Sub
eh:
MsgError
End Sub

Private Sub OLE1_DblClick()
On Error GoTo eh
OLE1.Verb = 0
OLE1.Action = 7
OLE1.Update
Exit Sub
eh:
MsgError
End Sub

Private Sub OLE1_Updated(Code As Integer)
'OLE1.Refresh
Static Rec As Boolean
If Not Rec Then
    Rec = True
'    OLE1.Update
    Rec = False
End If
'dbButton1_Click
End Sub

Private Sub Picture1_Click()
tmrShow_Timer
End Sub

Private Sub Timer1_Timer()
'OLE1.Refresh
On Error GoTo eh
OLE1.Update
Exit Sub
eh:
MsgError
End Sub

Public Sub LoadSettings()
On Error Resume Next
bLock = True
nmbFS.Value = dbGetSettingEx("Options", "EquationZoom", vbDouble, 2&)
bLock = False
End Sub

Public Sub SaveSettings()
dbSaveSettingEx "Options", "EquationZoom", nmbFS.Value
End Sub

Private Sub tmrShow_Timer()
If Not bLock Then
    tmrShow.Enabled = False
    dbButton1_Click
End If
End Sub
