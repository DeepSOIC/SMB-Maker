VERSION 5.00
Begin VB.Form frmToolTip 
   BackColor       =   &H80000018&
   BorderStyle     =   0  'None
   ClientHeight    =   1650
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4410
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   204
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000017&
   LinkTopic       =   "Form1"
   ScaleHeight     =   110
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   294
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Left            =   945
      Top             =   240
   End
End
Attribute VB_Name = "frmToolTip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strTip As String
Public w As Long, h As Long, l As Long, t As Long
Public tW As Long, tH As Long
Dim sW As Long, sH As Long


Private Sub Form_Load()
sW = Screen.Width \ Screen.TwipsPerPixelX
sH = Screen.Height \ Screen.TwipsPerPixelY
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
HideToolTipWindow
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Static cnt
cnt = cnt + 1
If cnt > 10 Then
    cnt = 0
    HideToolTipWindow
End If
End Sub

Friend Sub Form_Paint()
Const ST_Width = 500
Dim Rct As RECT
Dim DTP As DRAWTEXTPARAMS

On Error Resume Next
    Me.PaintPicture LoadResPicture("TOOLTIPBG", vbResBitmap), 0, 0, ScaleWidth, ScaleHeight
On Error GoTo 0
DTP.cbSize = LenB(DTP)
DTP.iLeftMargin = 4
DTP.iRightMargin = 4
DTP.iTabLength = 4
DTP.uiLengthDrawn = 0
If w = 0 Then w = ST_Width
Rct.Top = 1
Rct.Right = w
DrawTextEx Me.hDC, strTip, Len(strTip), Rct, DT_CALCRECT Or DT_WORDBREAK Or DT_NOCLIP, DTP
tW = Rct.Right - Rct.Left
tH = Rct.Bottom - Rct.Top
w = tW + 2
h = tH + 4

DrawTextEx Me.hDC, strTip, Len(strTip), Rct, DT_WORDBREAK Or DT_NOCLIP, DTP
'Me.Line (0, 0)-(ScaleWidth - 1, ScaleHeight - 1), Me.ForeColor, B
End Sub

Friend Sub SetText(ByRef Txt As String)
strTip = Txt
Form_Paint
End Sub

Friend Sub CalcPos(ByRef AroundRect As RECT)
If w > sW Then w = sW
l = (AroundRect.Left + AroundRect.Right) \ 2 - w \ 2
If l + w > sW Then l = sW - w
If l < 0 Then l = 0

t = AroundRect.Bottom
If t + h > sH Then
    t = AroundRect.Top - h
    If t < 0 Then
        If AroundRect.Top > (sH - AroundRect.Bottom) Then
            'do nothing
        Else
            t = AroundRect.Bottom
        End If
    End If
End If
End Sub

Friend Sub UpdatePos()
If w = 0 Or h = 0 Then Form_Paint
Me.Move l * Screen.TwipsPerPixelX, t * Screen.TwipsPerPixelY, w * Screen.TwipsPerPixelX, h * Screen.TwipsPerPixelY
End Sub

Friend Function Fits(ByVal X As Long, ByVal Y As Long) As Boolean
Fits = (X >= 0) And (Y >= 0) And (X + w <= sW) And (Y + h <= sH)
End Function

Private Sub Form_Resize()
Form_Paint
End Sub

Private Sub Timer1_Timer()
HideToolTipWindow
End Sub
