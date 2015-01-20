VERSION 5.00
Begin VB.Form frmViewImage 
   Caption         =   "View image"
   ClientHeight    =   3510
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8760
   Icon            =   "frmViewImage.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   234
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   584
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox picContainer 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2565
      Left            =   0
      ScaleHeight     =   167
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   351
      TabIndex        =   2
      Top             =   0
      Width           =   5325
      Begin VB.PictureBox picPicture 
         BackColor       =   &H00FF00FF&
         Height          =   1095
         Left            =   150
         ScaleHeight     =   1035
         ScaleWidth      =   900
         TabIndex        =   3
         Top             =   60
         Visible         =   0   'False
         Width           =   960
      End
   End
   Begin VB.VScrollBar VScroll 
      Height          =   2580
      Left            =   5355
      SmallChange     =   8
      TabIndex        =   1
      Top             =   0
      Width           =   240
   End
   Begin VB.HScrollBar HScroll 
      Height          =   240
      Left            =   0
      SmallChange     =   8
      TabIndex        =   0
      Top             =   2580
      Width           =   5310
   End
   Begin VB.PictureBox ButsBar 
      Height          =   450
      Left            =   30
      ScaleHeight     =   26
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   575
      TabIndex        =   4
      Top             =   2910
      Width           =   8685
      Begin SMBMaker.dbButton btnInvert 
         Height          =   390
         Left            =   4335
         TabIndex        =   9
         Top             =   15
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   688
         MouseIcon       =   "frmViewImage.frx":000C
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Others          =   $"frmViewImage.frx":0028
         OthersPresent   =   -1  'True
      End
      Begin SMBMaker.dbButton btnUnload 
         Height          =   405
         Left            =   7410
         TabIndex        =   5
         Top             =   -15
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   714
         MouseIcon       =   "frmViewImage.frx":007C
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Others          =   $"frmViewImage.frx":0098
         OthersPresent   =   -1  'True
      End
      Begin SMBMaker.dbButton btnEdit 
         Height          =   405
         Left            =   2940
         TabIndex        =   6
         Top             =   0
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   714
         MouseIcon       =   "frmViewImage.frx":00EC
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Others          =   $"frmViewImage.frx":0108
         OthersPresent   =   -1  'True
      End
      Begin SMBMaker.dbButton btnClose 
         Cancel          =   -1  'True
         Default         =   -1  'True
         Height          =   420
         Left            =   0
         TabIndex        =   7
         Top             =   45
         Width           =   1710
         _ExtentX        =   3016
         _ExtentY        =   741
         MouseIcon       =   "frmViewImage.frx":015A
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Others          =   $"frmViewImage.frx":0176
         OthersPresent   =   -1  'True
      End
      Begin SMBMaker.dbButton btnLoad 
         Height          =   405
         Left            =   1740
         TabIndex        =   8
         Top             =   15
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   714
         MouseIcon       =   "frmViewImage.frx":01C7
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Others          =   $"frmViewImage.frx":01E3
         OthersPresent   =   -1  'True
      End
      Begin SMBMaker.dbButton btnRotate 
         Height          =   390
         Left            =   5895
         TabIndex        =   10
         Top             =   15
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   688
         MouseIcon       =   "frmViewImage.frx":0235
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Others          =   $"frmViewImage.frx":0251
         OthersPresent   =   -1  'True
      End
   End
   Begin VB.Menu mnuPP 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu mnuRotClockwise 
         Caption         =   "Rotate clockwise"
         Tag             =   "2"
      End
      Begin VB.Menu mnuRotAntiClockwise 
         Caption         =   "Rotate anti-clockwise"
         Tag             =   "4"
      End
      Begin VB.Menu mnuRot180 
         Caption         =   "Rotate 180 degrees"
         Tag             =   "3"
      End
      Begin VB.Menu ppSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFlipHorz 
         Caption         =   "Flip horizontally"
         Tag             =   "0"
      End
      Begin VB.Menu mnuFlipVert 
         Caption         =   "Flip vertically"
         Tag             =   "1"
      End
   End
End
Attribute VB_Name = "frmViewImage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim gData() As Long
Dim Purpose As String

'Note that this sub sets data to nothing. Use Getimage
'to extract it back
'The image might be different or even absent
Friend Sub SetImage(ByRef Data() As Long)
SwapArys AryPtr(Data), AryPtr(gData)
VScroll.Value = 0
HScroll.Value = 0
Refr
End Sub

Friend Sub SetPurpose(ByRef aPurpose As String)
Purpose = aPurpose
End Sub

Friend Sub GetImage(ByRef Data() As Long)
SetImage Data
End Sub

Private Sub btnClose_Click()
Me.Tag = ""
Me.Hide
End Sub

Private Sub btnEdit_Click()
On Error GoTo eh
EditPicture gData
Refr
Exit Sub
eh:
If Err.Number = dbCWS Then Exit Sub
MsgError
End Sub

Private Sub btnInvert_Click()
If AryDims(AryPtr(gData)) = 2 Then
    MainForm.dbNegative gData, -1, 0, 0, 0, &HFFFFFF
    Refr
End If
End Sub

Private Sub btnLoad_Click()
Dim File As String
Dim Alpha() As Long
On Error GoTo eh
vtLoadPicture gData, Alpha, FileName:="", ShowDialog:=True, Purpose:=Purpose
Refr
Exit Sub
eh:
If Err.Number = dbCWS Then Exit Sub
MsgError
End Sub

Private Sub btnRotate_Click()
PopupMenu mnuPP, vbPopupMenuRightButton
End Sub

Private Sub btnUnload_Click()
Erase gData
Refr
End Sub

Private Sub ButsBar_Resize()
Const nButtons As Long = 6&
Dim i As Long
On Error Resume Next
i = 1
btnClose.Move 0, 0, ButsBar.ScaleWidth * i \ nButtons, ButsBar.ScaleHeight
i = i + 1
With btnClose
    btnLoad.Move .Left + .Width, 0, _
                 ButsBar.ScaleWidth * i \ nButtons - (.Left + .Width), _
                 ButsBar.ScaleHeight
End With
i = i + 1
With btnLoad
    btnEdit.Move .Left + .Width, 0, _
                 ButsBar.ScaleWidth * i \ nButtons - (.Left + .Width), _
                 ButsBar.ScaleHeight
End With
i = i + 1
With btnEdit
    btnInvert.Move .Left + .Width, 0, _
                 ButsBar.ScaleWidth * i \ nButtons - (.Left + .Width), _
                 ButsBar.ScaleHeight
End With
i = i + 1
With btnInvert
    btnRotate.Move .Left + .Width, 0, _
                 ButsBar.ScaleWidth * i \ nButtons - (.Left + .Width), _
                 ButsBar.ScaleHeight
End With
i = i + 1
With btnRotate
    btnUnload.Move .Left + .Width, 0, _
                 ButsBar.ScaleWidth * i \ nButtons - (.Left + .Width), _
                 ButsBar.ScaleHeight
End With
Debug.Assert i = nButtons
End Sub


Private Sub Refr()
Dim ImageW As Long, ImageH As Long
Dim bmi As BITMAPINFO
If AryDims(AryPtr(gData)) = 2 Then
    picContainer.Tag = ""
    picContainer.Cls
    ImageW = UBound(gData, 1) + 1
    ImageH = UBound(gData, 2) + 1
    picPicture.AutoRedraw = True
    picPicture.Cls
    picPicture.Move -HScroll.Value, -VScroll.Value, _
                    ImageW + 4, ImageH + 4
    With bmi.bmiHeader
        .biSize = Len(bmi.bmiHeader)
        .biWidth = ImageW
        .biHeight = -ImageH
        .biBitCount = 32
        .biSizeImage = ImageW * ImageH * 4
        .biPlanes = 1
    End With
    SetDIBitsToDevice picPicture.hDC, _
                      0, 0, ImageW, ImageH, _
                      0, 0, _
                      0, ImageH, _
                      gData(0, 0), _
                      bmi, DIB_RGB_COLORS
    picPicture.Visible = True
Else
    picPicture.Cls
    picPicture.Visible = False
    picPicture.Move 0, 0, 1, 1
    picContainer.Tag = GRSF(2539) '"<NO IMAGE>"
    picContainer_Paint
End If
picContainer_Resize
End Sub

Private Sub Form_Load()
LoadCaptions
End Sub

Friend Sub LoadCaptions()
Me.Caption = GRSF(2538)
mnuFlipHorz.Caption = GRSF(2549)
mnuFlipVert.Caption = GRSF(2550)
mnuRot180.Caption = GRSF(2548)
mnuRotClockwise.Caption = GRSF(2546)
mnuRotAntiClockwise.Caption = GRSF(2547)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If UnloadMode = 0 Then
    Cancel = True
    Me.Hide
End If
End Sub

Private Sub Form_Resize()
On Error Resume Next
ButsBar.Move 0, Me.ScaleHeight - ButsBar.Height, Me.ScaleWidth
HScroll.Move 0, ButsBar.Top - HScroll.Height, Me.ScaleWidth - HScroll.Height
VScroll.Move Me.ScaleWidth - HScroll.Height, 0, HScroll.Height, ButsBar.Top - HScroll.Height
picContainer.Move 0, 0, VScroll.Left, HScroll.Top
End Sub

Private Sub HScroll_Change()
UpdatePos
End Sub

Private Sub HScroll_Scroll()
UpdatePos
End Sub

Private Sub mnuFlipHorz_Click()
Rot Val(mnuFlipHorz.Tag)
End Sub

Private Sub mnuFlipVert_Click()
Rot Val(mnuFlipVert.Tag)
End Sub

Private Sub mnuRot180_Click()
Rot Val(mnuRot180.Tag)
End Sub

Private Sub mnuRotAntiClockwise_Click()
Rot Val(mnuRotAntiClockwise.Tag)
End Sub

Private Sub mnuRotClockwise_Click()
Rot Val(mnuRotClockwise.Tag)
End Sub

Private Sub Rot(ByVal Method As dbTurnMethod)
If AryDims(AryPtr(gData)) = 2 Then
    On Error GoTo eh
    MainForm.dbTurn gData, Method
    Refr
End If
Exit Sub
eh:
If Err.Number = dbCWS Then Exit Sub
MsgError
End Sub

Private Sub picContainer_Paint()
Dim w As Long, h As Long
On Error Resume Next
picContainer.PaintPicture gBackPicture, 0, 0, ScaleWidth, ScaleHeight
If Len(picContainer.Tag) > 0 Then
    w = picContainer.TextWidth(picContainer.Tag)
    h = picContainer.TextHeight(picContainer.Tag)
    picContainer.CurrentX = (picContainer.ScaleWidth - w) \ 2
    picContainer.CurrentY = (picContainer.ScaleHeight - h) \ 2
    picContainer.Print picContainer.Tag
End If
End Sub

Private Sub picContainer_Resize()
Dim hMax As Long, vMax As Long
hMax = picPicture.Width - picContainer.ScaleWidth
vMax = picPicture.Height - picContainer.ScaleHeight
If hMax > 0 Then
    HScroll.Enabled = True
    HScroll.Max = hMax
    HScroll.LargeChange = picContainer.ScaleWidth * 0.875
Else
    HScroll.Enabled = False
    HScroll.Max = 0
End If
If vMax > 0 Then
    VScroll.Enabled = True
    VScroll.Max = vMax
    VScroll.LargeChange = picContainer.ScaleHeight * 0.875
Else
    VScroll.Enabled = False
    VScroll.Max = 0
End If
UpdatePos
picContainer_Paint
End Sub

Private Sub UpdatePos()
Dim l As Long, t As Long
If HScroll.Enabled Then
    l = -HScroll.Value
Else
    l = (picContainer.ScaleWidth - picPicture.Width) \ 2
End If
If VScroll.Enabled Then
    t = -VScroll.Value
Else
    t = (picContainer.ScaleHeight - picPicture.Height) \ 2
End If
picPicture.Move l, t
End Sub

Private Sub VScroll_Change()
UpdatePos
End Sub

Private Sub VScroll_Scroll()
UpdatePos
End Sub
