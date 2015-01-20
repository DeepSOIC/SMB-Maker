VERSION 5.00
Begin VB.Form frmSplash 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5040
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   7380
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmSplash.frx":000C
   ScaleHeight     =   336
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   492
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label lblVersion 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Version"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   450
      Left            =   2280
      TabIndex        =   0
      Top             =   2790
      Visible         =   0   'False
      Width           =   1110
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim VersionString As String

Private Sub Form_Load()
    Dim Wdt As Long, hgt As Long
    Dim XBorder As Long, YBorder As Long
    XBorder = GetSystemMetrics(SM_CXSIZEFRAME) * Screen.TwipsPerPixelX
    YBorder = GetSystemMetrics(SM_CYSIZEFRAME) * Screen.TwipsPerPixelY
    VersionString = CStr(App.Major) + "." + CStr(App.Minor) + "." + CStr(App.Revision)
    lblVersion.Caption = VersionString
    Wdt = ScaleX(Me.Picture.Width, vbHimetric, vbTwips) + 2 * XBorder
    hgt = ScaleY(Me.Picture.Height, vbHimetric, vbTwips) + 2 * YBorder
    Me.Move (Screen.Width - Wdt) \ 2, (Screen.Height - hgt) \ 2, Wdt, hgt
End Sub

Private Sub Form_Paint()
    Me.Font = lblVersion.Font
    Me.FontSize = lblVersion.FontSize
    
    Me.ForeColor = vbBlack
    Me.CurrentX = lblVersion.Left + 1
    Me.CurrentY = lblVersion.Top + 1
    Me.Print VersionString
    Me.CurrentX = lblVersion.Left - 1
    Me.CurrentY = lblVersion.Top - 1
    Me.Print VersionString
    Me.CurrentX = lblVersion.Left + 1
    Me.CurrentY = lblVersion.Top - 1
    Me.Print VersionString
    Me.CurrentX = lblVersion.Left - 1
    Me.CurrentY = lblVersion.Top + 1
    Me.Print VersionString

    
    Me.ForeColor = vbYellow
    Me.CurrentX = lblVersion.Left
    Me.CurrentY = lblVersion.Top
    Me.Print VersionString

End Sub
