VERSION 5.00
Begin VB.Form frmRect 
   BackColor       =   &H00E3DFE0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Rectangle"
   ClientHeight    =   2460
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   3855
   Icon            =   "frmRect.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2460
   ScaleWidth      =   3855
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.OptionButton Opt 
      BackColor       =   &H0080FFFF&
      Height          =   270
      Index           =   1
      Left            =   2085
      Style           =   1  'Graphical
      TabIndex        =   1
      Tag             =   "1"
      ToolTipText     =   "Bar (filled rectangle)"
      Top             =   1365
      Width           =   255
   End
   Begin VB.OptionButton Opt 
      BackColor       =   &H0080FFFF&
      Height          =   270
      Index           =   2
      Left            =   15
      Style           =   1  'Graphical
      TabIndex        =   2
      Tag             =   "3"
      ToolTipText     =   "Background-filled rectangle"
      Top             =   1350
      Width           =   255
   End
   Begin VB.OptionButton Opt 
      BackColor       =   &H0080FFFF&
      Height          =   270
      Index           =   0
      Left            =   15
      Style           =   1  'Graphical
      TabIndex        =   0
      Tag             =   "0"
      ToolTipText     =   "Border (transperent rectangle)"
      Top             =   15
      Value           =   -1  'True
      Width           =   255
   End
   Begin SMBMaker.dbButton CancelButton 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   2505
      TabIndex        =   4
      Top             =   525
      Width           =   1215
      _ExtentX        =   0
      _ExtentY        =   0
      MouseIcon       =   "frmRect.frx":0ABA
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Others          =   $"frmRect.frx":0AD6
      OthersPresent   =   -1  'True
   End
   Begin SMBMaker.dbButton OKButton 
      Height          =   375
      Left            =   2505
      TabIndex        =   3
      Top             =   45
      Width           =   1215
      _ExtentX        =   0
      _ExtentY        =   0
      MouseIcon       =   "frmRect.frx":0B26
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Others          =   $"frmRect.frx":0B42
      OthersPresent   =   -1  'True
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H0000FF00&
      BorderWidth     =   3
      FillColor       =   &H0000FF00&
      FillStyle       =   0  'Solid
      Height          =   960
      Left            =   2385
      Top             =   1380
      Width           =   1380
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H0000FF00&
      BorderWidth     =   3
      FillStyle       =   0  'Solid
      Height          =   960
      Left            =   285
      Top             =   1350
      Width           =   1380
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0000FF00&
      BorderWidth     =   3
      Height          =   915
      Left            =   300
      Top             =   60
      Width           =   1335
   End
End
Attribute VB_Name = "frmRect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Form_Paint()
Me.PaintPicture LoadResPicture(RES_BackPicture, vbResBitmap), 0, 0, Me.ScaleWidth, Me.ScaleHeight
End Sub

Private Sub CancelButton_Click()
Me.Tag = "C"
Me.Hide
End Sub

Private Sub Form_Initialize()
'Const Action = 1, F = 2044, T = 2048
'Dim obj As Object, i As Integer, index As Integer, tmp As String
'On Error Resume Next
'For Each obj In Me
'    index = -1
'    index = obj.index
'    Err.Clear
'    tmp = obj.Caption
'    If Err = 0 Then
'        If Action = 0 Then
'        Debug.Print obj.Name + IIf((index = -1), "", "(" + CStr(index) + ")") + ".Caption = ", """" + tmp + """"
'        Else
'        For i = F To T
'            If grsf(i) = tmp Then
'            Debug.Print obj.Name + IIf((index = -1), "", "(" + CStr(index) + ")") + ".Caption = ", "grsf(" + CStr(i) + ")"
'            Exit For
'            End If
'        Next i
'        End If
'    Else
'        Err.Clear
'        'Debug.Print "-"
'    End If
'
'    index = -1
'    index = obj.index
'    Err.Clear
'    tmp = obj.ToolTipText
'    If Err = 0 Then
'        If Action = 0 Then
'        Debug.Print obj.Name + IIf((index = -1), "", "(" + CStr(index) + ")") + ".ToolTipText = ", """" + tmp + """"
'        Else
'        For i = F To T
'            If grsf(i) = tmp Then
'            Debug.Print obj.Name + IIf((index = -1), "", "(" + CStr(index) + ")") + ".tooltiptext = ", "grsf(" + CStr(i) + ")"
'            Exit For
'            End If
'        Next i
'        End If
'    Else
'        Err.Clear
'        'Debug.Print "-"
'    End If
'
'Next
'End
dbLoadCaptions
LoadSettings
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If UnloadMode = 0 Then
    Cancel = True
    CancelButton_Click
End If
End Sub

Private Sub OkButton_Click()
Me.Tag = ""
Me.Hide
End Sub


Public Property Get rs() As Long
Dim i As Long
On Error Resume Next

For i = Opt.lBound To Opt.UBound
If Opt(i).Value Then rs = Opt(i).Tag: Exit Property
Next i
rs = 0
End Property

Public Property Let rs(ByVal lNew As Long)
Dim i As Integer
For i = Opt.lBound To Opt.UBound
If CLng(Opt(i).Tag) = lNew Then Opt(i).Value = True
Next i
End Property

Sub dbLoadCaptions()
Me.Caption = GRSF(2137)
Opt(1).ToolTipText = GRSF(2044)
Opt(2).ToolTipText = GRSF(2045)
Opt(0).ToolTipText = GRSF(2046)
CancelButton.Caption = GRSF(2047)
OkButton.Caption = GRSF(2048)
'Me.Icon = LoadResPicture(Me.Name, vbResIcon)
End Sub

Public Sub LoadSettings()

End Sub
