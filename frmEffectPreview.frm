VERSION 5.00
Begin VB.Form frmEffectPreview 
   BorderStyle     =   0  'None
   ClientHeight    =   3750
   ClientLeft      =   105
   ClientTop       =   105
   ClientWidth     =   4755
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   250
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   317
   ShowInTaskbar   =   0   'False
   Begin SMBMaker.ctlResourcizator Resr1 
      Left            =   855
      Top             =   3210
      _ExtentX        =   1164
      _ExtentY        =   741
      ResID           =   9905
   End
   Begin VB.PictureBox Pic 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3090
      Left            =   45
      ScaleHeight     =   202
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   301
      TabIndex        =   0
      ToolTipText     =   "Right-click for context menu."
      Top             =   15
      Width           =   4575
      Begin VB.Timer tmrEnabler 
         Interval        =   100
         Left            =   405
         Top             =   2400
      End
   End
   Begin VB.Menu mnuPP 
      Caption         =   ""
      Enabled         =   0   'False
      Visible         =   0   'False
      Begin VB.Menu mnuAutoUpdate 
         Caption         =   "Auto Update"
      End
   End
End
Attribute VB_Name = "frmEffectPreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim pEffect As clsEffect
Dim AryPtrData As Long
Dim ViewPosX As Long, ViewPosY As Long
Dim CX As Long, CY As Long
Dim intW As Long, intH As Long
Dim mdx As Long, mdy As Long
Dim mmx As Long, mmy As Long

Private Sub UpdateCXCY()
CX = Pic.ScaleWidth \ 2
CY = Pic.ScaleHeight \ 2
AryWH AryPtrData, intW, intH
End Sub

Public Sub SetEffect(ByRef Effect As clsEffect)
Dim sMenuName As String, sID As String
Set pEffect = Effect
pEffect.GetEffectDesc sID, sMenuName
mnuAutoUpdate.Checked = dbGetSettingEx("Effects\" + sID, "PreviewAutoUpdate", vbBoolean, True)
End Sub

Public Sub UnreferEffect()
Set pEffect = Nothing
End Sub

Public Sub SetData(ByVal ptrArray As Long)
AryPtrData = ptrArray
On Error Resume Next
UpdateCXCY
ViewPosX = intW \ 2
ViewPosY = intH \ 2
End Sub

Public Sub UnSetData()
AryPtrData = 0
End Sub

Public Sub Update(ByVal WithEffect As Boolean)
Dim tw As Long, th As Long
Dim St As String
Dim Rct As RECT
Dim DTP As DRAWTEXTPARAMS
If mnuAutoUpdate.Checked Then
    prvUpdate WithEffect:=WithEffect
Else
    Pic.Cls
    St = GRSF(2600)
    th = Pic.TextHeight(St)
    Rct.Left = 0
    Rct.Top = (Pic.ScaleHeight - th) \ 2
    Rct.Right = Pic.ScaleWidth
    Rct.Bottom = Pic.Height
    DTP.cbSize = Len(DTP)
    DrawTextEx Pic.hDC, St, Len(St), Rct, DT_CENTER, DTP
'    tw = Pic.TextWidth(St)
'    Pic.CurrentX = (Pic.ScaleWidth - tw) \ 2
'    Pic.CurrentY = (Pic.ScaleHeight - th) \ 2
'    Pic.Print St;
End If
End Sub

Private Sub prvUpdate(ByVal WithEffect As Boolean)
Dim Rct As RECT
Dim DrawPosX As Long, DrawPosY As Long
Dim Src() As Long
Static Data() As Long
On Error GoTo eh

If WithEffect Then
    pEffect.FormToSettings
End If


UpdateCXCY
Rct.Left = ViewPosX - CX
Rct.Top = ViewPosY - CY
Rct.Right = Rct.Left + Pic.ScaleWidth
Rct.Bottom = Rct.Top + Pic.ScaleHeight
DrawPosX = Max(0, -Rct.Left)
DrawPosY = Max(0, -Rct.Top)
If Rct.Top < 0 Then Rct.Top = 0
If Rct.Left < 0 Then Rct.Left = 0
If Rct.Right > intW Then Rct.Right = intW
If Rct.Bottom > intH Then Rct.Bottom = intH
If IsRectEmpty(Rct) Then Exit Sub

On Error GoTo eh
ReferAry AryPtr(Src), AryPtrData
If WithEffect Then
    CancelDoEvents
    pEffect.PerformEffect Src, Data, Rct.Left, Rct.Top, Rct.Right, Rct.Bottom
    RestoreDoEvents
    Pic.Line (0, 0)-(Pic.ScaleWidth - 1, Pic.ScaleHeight - 1), Pic.BackColor, BF
    vtSetDIBitsToDevice Pic.hDC, Data, _
                        0, 0, _
                        DrawPosX, DrawPosY, _
                        Rct.Right - Rct.Left, Rct.Bottom - Rct.Top
Else
    Pic.Line (0, 0)-(Pic.ScaleWidth - 1, Pic.ScaleHeight - 1), Pic.BackColor, BF
    vtSetDIBitsToDevice Pic.hDC, Src, _
                        Rct.Left, Rct.Top, _
                        DrawPosX, DrawPosY, _
                        Rct.Right - Rct.Left, Rct.Bottom - Rct.Top
End If

UnReferAry AryPtr(Src)
Pic.Refresh
Exit Sub
Resume
eh:
PushError
UnReferAry AryPtr(Src)
RestoreDoEvents
PopError
ErrRaise "frmEffectPreview:Update"
End Sub

Private Sub Form_Load()
Resr1.LoadCaptions
End Sub

Private Sub Form_Resize()
Pic.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
End Sub

Private Sub mnuAutoUpdate_Click()
Dim sID As String, sMenu As String
mnuAutoUpdate.Checked = Not mnuAutoUpdate.Checked
pEffect.GetEffectDesc sID, sMenu
dbSaveSettingEx "Effects\" + sID, "PreviewAutoUpdate", mnuAutoUpdate.Checked
End Sub

Private Sub Pic_Click()
'Static Foo As Boolean
'Foo = Not Foo
'Update Foo
End Sub

Private Sub Pic_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
mdx = x
mdy = y
mmx = x
mmy = y
If Button = 1 Then
    prvUpdate False
ElseIf Button = 2 Then
    PopupMenu mnuPP, vbPopupMenuRightButton
End If
End Sub

Private Sub Pic_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 1 Then
    ViewPosX = ViewPosX - (x - mmx)
    ViewPosY = ViewPosY - (y - mmy)
    ValidateViewPos
    prvUpdate False
End If
mmx = x
mmy = y
End Sub

Private Sub Pic_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error GoTo eh
If Button = 1 Then
    prvUpdate True
End If
Exit Sub
eh:
vtBeep
ShowStatus Err.Source + ":" + Err.Description
End Sub

Private Sub Pic_Resize()
On Error GoTo eh
Update True
Exit Sub
eh:
MsgError
End Sub

Private Sub ValidateViewPos()
UpdateCXCY
If ViewPosX < 0 Then ViewPosX = 0
If ViewPosY < 0 Then ViewPosY = 0
If ViewPosX > intW - 1 Then ViewPosX = intW - 1
If ViewPosY > intH - 1 Then ViewPosY = intH - 1
End Sub

Private Sub tmrEnabler_Timer()
EnableWindow Me.hWnd, True
End Sub
