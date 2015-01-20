VERSION 5.00
Begin VB.Form Bytes 
   Caption         =   "Output"
   ClientHeight    =   3157
   ClientLeft      =   165
   ClientTop       =   550
   ClientWidth     =   4675
   Icon            =   "Bytes.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   3157
   ScaleWidth      =   4675
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Updater 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   3894
      Top             =   407
   End
   Begin VB.TextBox Text 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.07
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2280
      HideSelection   =   0   'False
      Left            =   110
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   143
      Width           =   3540
   End
   Begin SMBMaker.dbButton btnClear 
      Height          =   418
      Left            =   2409
      TabIndex        =   2
      Top             =   2574
      Width           =   1991
      _extentx        =   3515
      _extenty        =   732
      otherspresent   =   -1
      others          =   $"Bytes.frx":0ABA
      mouseicon       =   "Bytes.frx":0B0A
      font            =   "Bytes.frx":0B28
   End
   Begin SMBMaker.dbButton btnOK 
      Height          =   506
      Left            =   220
      TabIndex        =   1
      Top             =   2497
      Width           =   1804
      _extentx        =   3170
      _extenty        =   894
      otherspresent   =   -1
      others          =   $"Bytes.frx":0B50
      mouseicon       =   "Bytes.frx":0B9C
      font            =   "Bytes.frx":0BBA
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuSave 
         Caption         =   "Save..."
      End
      Begin VB.Menu Sep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuClose 
         Caption         =   "Close"
      End
   End
End
Attribute VB_Name = "Bytes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim AccuS As StringAccumulator

Private Sub btnClear_Click()
Text.Text = ""
End Sub

Private Sub btnOK_Click()
Me.Hide
End Sub

Private Sub Form_Initialize()
Set AccuS = New StringAccumulator
dbLoadCaptions
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 And Shift = 1 Then 'shift+enter
btnOk.RaiseClick
End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If UnloadMode = 0 Then
    Cancel = -1
    Me.Hide
End If
End Sub

Private Sub Form_Resize()
On Error Resume Next
Text.Move 0, 0, ScaleWidth, ScaleHeight - btnOk.Height
'Text.Width = ScaleWidth
'Text.Height = ScaleHeight - btnOK.Height
btnOk.Move 0, ScaleHeight - btnOk.Height, ScaleWidth \ 2
btnClear.Move btnOk.Left + btnOk.Width, btnOk.Top, ScaleWidth - (btnOk.Left + btnOk.Width), btnOk.Height
'btnCancel.Move btnOK.Width, btnOK.Top, btnOK.Height, ScaleWidth - ScaleWidth \ 2
End Sub

Private Sub mnuClose_Click()
Me.Hide
End Sub

Private Sub mnuSave_Click()
Dim File As String, nmb As Long
On Error GoTo eh
'With CDl
'    .Filter = "Text file (*.txt)|*.txt|File (*.*)|*.*"
'    .DialogTitle = "Save output"
'    .FileName = ""
'    .ShowSave
'    File = .FileName
'    .DialogTitle = ""
'End With
File = ShowSaveDlg(dbTextSave, Me.hWnd, Purpose:="LOG")
nmb = FreeFile
Open File For Output As nmb
    Print #(nmb), Text.Text;
Close nmb
Exit Sub
eh:
Reset
End Sub

Sub dbLoadCaptions()
mnuFile.Caption = GRSF(183)
mnuSave.Caption = GRSF(174)
mnuClose.Caption = GRSF(175)
'Me.Icon = LoadResPicture(Me.Name, vbResIcon)
End Sub

Sub Output(ByVal St As String)
  AccuS.Append St
  Updater.Enabled = True
End Sub

Private Sub Text_KeyPress(KeyAscii As Integer)
If KeyAscii = Asc(vbTab) Then
End If
End Sub

Private Sub Updater_Timer()
If AccuS.Length > 0 Then
  Dim t As Long
  t = Text.SelStart
  Text.SelText = ""
  Text.Text = Left$(Text.Text, t) + AccuS.Content + Mid$(Text.Text, t + 1)
  Text.SelStart = t + AccuS.Length
  AccuS.Clear
End If
Updater.Enabled = False
End Sub
