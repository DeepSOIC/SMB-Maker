VERSION 5.00
Begin VB.Form frmFill 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Filling options"
   ClientHeight    =   4905
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4620
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4905
   ScaleWidth      =   4620
   StartUpPosition =   1  'CenterOwner
   Begin SMBMaker.dbFrame dbFrame2 
      Height          =   1605
      Left            =   188
      TabIndex        =   7
      Top             =   2880
      Width           =   4245
      _ExtentX        =   7488
      _ExtentY        =   2831
      Caption         =   "Filling"
      ResID           =   2565
      EAC             =   0   'False
      Begin VB.OptionButton optFill 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Textured"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   2
         Left            =   60
         Style           =   1  'Graphical
         TabIndex        =   11
         Tag             =   "2"
         Top             =   1035
         Width           =   2670
      End
      Begin VB.TextBox nmbAlpha 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2775
         TabIndex        =   10
         Tag             =   "1"
         Text            =   "128"
         Top             =   645
         Width           =   885
      End
      Begin VB.OptionButton optFill 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Single-color alpha-blended"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   1
         Left            =   60
         Style           =   1  'Graphical
         TabIndex        =   9
         Tag             =   "1"
         Top             =   645
         Width           =   2670
      End
      Begin VB.OptionButton optFill 
         BackColor       =   &H00C0FFC0&
         Caption         =   "With single color"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   0
         Left            =   60
         Style           =   1  'Graphical
         TabIndex        =   8
         Tag             =   "0"
         Top             =   255
         Width           =   2670
      End
      Begin SMBMaker.dbButton btnViewTex 
         Height          =   255
         Left            =   2745
         TabIndex        =   12
         Tag             =   "2"
         Top             =   1035
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   450
         MouseIcon       =   "frmFill.frx":0000
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Others          =   $"frmFill.frx":001C
         OthersPresent   =   -1  'True
      End
      Begin SMBMaker.dbButton btnOrgTex 
         Height          =   210
         Left            =   2730
         TabIndex        =   14
         Tag             =   "2"
         Top             =   1290
         Width           =   1110
         _ExtentX        =   1958
         _ExtentY        =   370
         MouseIcon       =   "frmFill.frx":006E
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Others          =   $"frmFill.frx":008A
         OthersPresent   =   -1  'True
      End
   End
   Begin SMBMaker.dbFrame dbFrame1 
      Height          =   2760
      Left            =   188
      TabIndex        =   0
      Top             =   90
      Width           =   4245
      _ExtentX        =   7488
      _ExtentY        =   4868
      Caption         =   "Border detection"
      ResID           =   2552
      EAC             =   0   'False
      Begin SMBMaker.ctlColor clrTreshold 
         Height          =   390
         Index           =   1
         Left            =   3240
         TabIndex        =   15
         Top             =   645
         Width           =   570
         _ExtentX        =   1005
         _ExtentY        =   688
      End
      Begin VB.OptionButton optBorder 
         BackColor       =   &H0080FFFF&
         Caption         =   "Border by gradient"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   5
         Left            =   15
         Style           =   1  'Graphical
         TabIndex        =   6
         Tag             =   "2"
         Top             =   2205
         Width           =   3210
      End
      Begin VB.OptionButton optBorder 
         BackColor       =   &H0080FFFF&
         Caption         =   "Paint while the color falls"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   4
         Left            =   15
         Style           =   1  'Graphical
         TabIndex        =   5
         Tag             =   "5"
         Top             =   1815
         Width           =   3210
      End
      Begin VB.OptionButton optBorder 
         BackColor       =   &H0080FFFF&
         Caption         =   "Paint while the color grows"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   3
         Left            =   15
         Style           =   1  'Graphical
         TabIndex        =   4
         Tag             =   "4"
         Top             =   1425
         Width           =   3210
      End
      Begin VB.OptionButton optBorder 
         BackColor       =   &H0080FFFF&
         Caption         =   "Single-coloured area (non-exact):"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   2
         Left            =   15
         Style           =   1  'Graphical
         TabIndex        =   3
         Tag             =   "3"
         Top             =   1035
         Width           =   3210
      End
      Begin VB.OptionButton optBorder 
         BackColor       =   &H0080FFFF&
         Caption         =   "Border color:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   1
         Left            =   15
         Style           =   1  'Graphical
         TabIndex        =   2
         Tag             =   "1"
         Top             =   645
         Width           =   3210
      End
      Begin VB.OptionButton optBorder 
         BackColor       =   &H0080FFFF&
         Caption         =   "Paint single-coloured area"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   0
         Left            =   15
         Style           =   1  'Graphical
         TabIndex        =   1
         Tag             =   "0"
         Top             =   255
         Value           =   -1  'True
         Width           =   3210
      End
      Begin SMBMaker.ctlColor clrTreshold 
         Height          =   390
         Index           =   2
         Left            =   3240
         TabIndex        =   16
         Top             =   1035
         Width           =   570
         _ExtentX        =   1005
         _ExtentY        =   688
      End
      Begin SMBMaker.ctlColor clrTreshold 
         Height          =   390
         Index           =   5
         Left            =   3240
         TabIndex        =   17
         Top             =   2205
         Width           =   570
         _ExtentX        =   1005
         _ExtentY        =   688
      End
   End
   Begin SMBMaker.dbButton OkButton 
      Cancel          =   -1  'True
      Default         =   -1  'True
      Height          =   330
      Left            =   495
      TabIndex        =   13
      Top             =   4545
      Width           =   3630
      _ExtentX        =   6403
      _ExtentY        =   582
      MouseIcon       =   "frmFill.frx":00DE
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Others          =   $"frmFill.frx":00FA
      OthersPresent   =   -1  'True
   End
End
Attribute VB_Name = "frmFill"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim FillTex() As Long
Dim TexOrg As New clsAligner

Private Sub btnOrgTex_Click()
TexOrg.Customize RaiseErrors:=False
End Sub

Private Sub btnViewTex_Click()
ViewImage FillTex, "FillTex"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If UnloadMode = 0 Then
    Cancel = True
    OkButton_Click
End If
End Sub

Private Sub OkButton_Click()
Me.Tag = ""
Me.Hide
End Sub

Private Sub Form_Load()
optBorder_Click 0
optFill_Click 0
TexOrg.BasePointSupported = True
TexOrg.DestPointSupported = False
LoadCaptions
End Sub

Friend Sub LoadCaptions()
Me.Caption = GRSF(2551)
optBorder(0).Caption = GRSF(2553)
optBorder(0).ToolTipText = GRSF(2554)
optBorder(1).Caption = GRSF(2555)
optBorder(1).ToolTipText = GRSF(2556)
optBorder(2).Caption = GRSF(2557)
optBorder(2).ToolTipText = GRSF(2558)
optBorder(3).Caption = GRSF(2559)
optBorder(3).ToolTipText = GRSF(2560)
optBorder(4).Caption = GRSF(2561)
optBorder(4).ToolTipText = GRSF(2562)
optBorder(5).Caption = GRSF(2563)
optBorder(5).ToolTipText = GRSF(2564)

optFill(0).Caption = GRSF(2566)
optFill(1).Caption = GRSF(2567)
optFill(1).ToolTipText = GRSF(2568)
optFill(2).Caption = GRSF(2569)
optFill(2).ToolTipText = GRSF(2570)

btnViewTex.ToolTipText = GRSF(2572)
btnOrgTex.ToolTipText = GRSF(2574)
End Sub

Private Sub optBorder_Click(Index As Integer)
Dim Ctl As Control
For Each Ctl In clrTreshold
    Ctl.Enabled = optBorder(Ctl.Index).Value
Next
End Sub

Private Sub optFill_Click(Index As Integer)
With nmbAlpha
    .Enabled = optFill(Val(.Tag)).Value
End With
With btnViewTex
    .Enabled = optFill(Val(.Tag)).Value
End With
With btnOrgTex
    .Enabled = optFill(Val(.Tag)).Value
End With
End Sub

Private Function GetBorderMode()
Dim Obj As OptionButton
For Each Obj In optBorder
    If Obj.Value Then
        GetBorderMode = Obj.Index
    End If
Next
End Function

Private Function GetFillMode()
Dim Obj As OptionButton
For Each Obj In optFill
    If Obj.Value Then
        GetFillMode = Obj.Index
    End If
Next
End Function

Friend Sub GetMode(ByRef FillOpts As FillSettings)
FillOpts.BorderMode = Val(optBorder(GetBorderMode).Tag)
On Error Resume Next
FillOpts.Treshold = clrTreshold(GetBorderMode).Color
On Error GoTo 0
FillOpts.FillMode = Val(optFill(GetFillMode).Tag)
FillOpts.FillAlpha = nmbAlpha
SwapArys AryPtr(FillOpts.Texture), AryPtr(FillTex)
Set FillOpts.TexOrigin = TexOrg
End Sub

Private Sub SetBorderMode(ByVal BM As dbFillBorderMode)
Dim Obj As OptionButton
For Each Obj In optBorder
    If CLng(Obj.Tag) = BM Then
        Obj.Value = True
    End If
Next
End Sub

Private Sub SetFillMode(ByVal FM As dbFillMode)
Dim Obj As OptionButton
For Each Obj In optFill
    If CLng(Obj.Tag) = FM Then
        Obj.Value = True
    End If
Next
End Sub

Friend Sub SetMode(ByRef FillOpts As FillSettings)
SetBorderMode FillOpts.BorderMode
On Error Resume Next
clrTreshold(GetBorderMode).Color = FillOpts.Treshold
On Error GoTo 0
SetFillMode FillOpts.FillMode
nmbAlpha = FillOpts.FillAlpha
SwapArys AryPtr(FillOpts.Texture), AryPtr(FillTex)
Set TexOrg = New clsAligner
With FillOpts.TexOrigin
    TexOrg.AnchorX = .AnchorX
    TexOrg.AnchorY = .AnchorY
    TexOrg.BaseAnchorX = .BaseAnchorX
    TexOrg.BaseAnchorY = .BaseAnchorY
    TexOrg.BasePointSupported = True
    TexOrg.DestPointSupported = False
    TexOrg.dx = .dx
    TexOrg.dy = .dy
End With
End Sub

