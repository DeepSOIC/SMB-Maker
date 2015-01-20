VERSION 5.00
Begin VB.Form frmLens 
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   4680
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox pctCapture 
      AutoRedraw      =   -1  'True
      Height          =   2715
      Left            =   15
      ScaleHeight     =   177
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   235
      TabIndex        =   0
      Top             =   0
      Width           =   3585
   End
   Begin VB.Menu mnuP 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu mnuZoomIn 
         Caption         =   "Zoom in	Shift+LMB"
      End
      Begin VB.Menu mnuZoomOut 
         Caption         =   "Zoom Out	Shift+RMB"
      End
      Begin VB.Menu sep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDock 
         Caption         =   "Dock into SMB Maker	Alt+LMB"
      End
      Begin VB.Menu Sep2 
         Caption         =   "-"
      End
      Begin VB.Menu HintItem1 
         Caption         =   "LMB = Left"
         Enabled         =   0   'False
      End
      Begin VB.Menu HintItem2 
         Caption         =   "  mouse button"
         Enabled         =   0   'False
      End
      Begin VB.Menu HintItem3 
         Caption         =   "RMB = right "
         Enabled         =   0   'False
      End
      Begin VB.Menu HintItem4 
         Caption         =   "  mouse button"
         Enabled         =   0   'False
      End
   End
End
Attribute VB_Name = "frmLens"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mx As Long, my As Long

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
MainForm.SetFocus
End Sub

Private Sub Form_Resize()
Static Rec As Boolean
Dim l As Long, t As Long
Dim w As Long, h As Long
If Rec Then Exit Sub
Rec = True
Me.WindowState = vbNormal
On Error GoTo eh
w = Me.Width
h = Me.Height
If w > Screen.Width Then w = Screen.Width
If h > Screen.Height Then h = Screen.Height
If w < 32 * Screen.TwipsPerPixelX Then w = 32 * Screen.TwipsPerPixelX
If h < 32 * Screen.TwipsPerPixelY Then h = 32 * Screen.TwipsPerPixelY
t = Me.Top
l = Me.Left
If t < 0 Then t = 0
If l < 0 Then l = 0
If t + h > Screen.Height Then t = Screen.Height - h
If l + w > Screen.Width Then l = Screen.Width - w
If t <> Me.Top Or l <> Me.Left Or w <> Me.Width Or h <> Me.Height Then
    Me.Move l, t, w, h
End If
pctCapture.Move 0, 0, ScaleWidth, ScaleHeight
eh:
Rec = False
End Sub

Private Sub PctCapture_DblClick()
MainForm.PctCapture_DblClick
End Sub

Private Sub pctCapture_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 1 And Shift = 0 Then
    mx = x
    my = y
Else
    mx = -1
    MainForm.pctCapture_MouseDown Button, Shift, x, y
End If

End Sub

Private Sub pctCapture_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim dx As Long, dy As Long
Dim l As Long, t As Long
If Button = 1 And mx <> -1 Then
    dx = x - mx
    dy = y - my
    l = Me.Left + dx * Screen.TwipsPerPixelX
    t = Me.Top + dy * Screen.TwipsPerPixelY
    If l < 0 Then l = 0
    If t < 0 Then t = 0
    
    If l + Me.Width > Screen.Width Then
        l = Screen.Width - Me.Width
    End If
    
    If t + Me.Height > Screen.Height Then
        t = Screen.Height - Me.Height
    End If
    
    Me.Move l, t
Else
    'MainForm.pctCapture_MouseMove Button, Shift, X, Y
End If
End Sub
