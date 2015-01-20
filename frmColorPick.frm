VERSION 5.00
Begin VB.Form frmColorPick 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Color"
   ClientHeight    =   3150
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3150
   ScaleWidth      =   6000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin SMBMaker.dbFrame dbFrame2 
      Height          =   2370
      Left            =   4500
      TabIndex        =   3
      Top             =   780
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   4180
      Caption         =   "RGB Color"
      ResID           =   2374
      EAC             =   0   'False
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   150
         TabIndex        =   14
         Text            =   "#FFFFFF"
         ToolTipText     =   "Web color. Use this field to copy/paste colors."
         Top             =   1770
         Width           =   1200
      End
      Begin VB.PictureBox pctAll 
         BackColor       =   &H00FFFFFF&
         Height          =   660
         Left            =   150
         ScaleHeight     =   600
         ScaleWidth      =   1140
         TabIndex        =   10
         ToolTipText     =   "Take this with RMB and pull to get color from screen. Also try clicking with LMB."
         Top             =   1095
         Width           =   1200
      End
      Begin VB.PictureBox pctC 
         BackColor       =   &H00FF0000&
         Height          =   285
         Index           =   2
         Left            =   150
         ScaleHeight     =   15
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   21
         TabIndex        =   9
         Top             =   810
         Width           =   375
      End
      Begin VB.TextBox txtC 
         Height          =   285
         Index           =   2
         Left            =   525
         MaxLength       =   4
         TabIndex        =   8
         ToolTipText     =   "Blue channel. From 0 to 255."
         Top             =   810
         Width           =   825
      End
      Begin VB.PictureBox pctC 
         BackColor       =   &H0000FF00&
         Height          =   285
         Index           =   1
         Left            =   150
         ScaleHeight     =   15
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   21
         TabIndex        =   7
         Top             =   525
         Width           =   375
      End
      Begin VB.TextBox txtC 
         Height          =   285
         Index           =   1
         Left            =   525
         MaxLength       =   4
         TabIndex        =   6
         ToolTipText     =   "Green channel. From 0 to 255."
         Top             =   525
         Width           =   825
      End
      Begin VB.TextBox txtC 
         Height          =   285
         Index           =   0
         Left            =   525
         MaxLength       =   4
         TabIndex        =   5
         ToolTipText     =   "Red channel. From 0 to 255."
         Top             =   240
         Width           =   825
      End
      Begin VB.PictureBox pctC 
         BackColor       =   &H000000FF&
         Height          =   285
         Index           =   0
         Left            =   150
         ScaleHeight     =   15
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   21
         TabIndex        =   4
         Top             =   240
         Width           =   375
      End
   End
   Begin SMBMaker.dbFrame dbFrame1 
      Height          =   3150
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4500
      _ExtentX        =   7938
      _ExtentY        =   5556
      Caption         =   "Standard colors"
      ResID           =   2371
      EAC             =   0   'False
      Begin VB.OptionButton OptPal 
         BackColor       =   &H0080FFFF&
         Caption         =   "System Colors"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   390
         Index           =   1
         Left            =   2115
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Windows color palette. Changes with settings in Control panel/screen/appearance."
         Top             =   2610
         Width           =   1890
      End
      Begin VB.OptionButton OptPal 
         BackColor       =   &H0080FFFF&
         Caption         =   "Palette"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   390
         Index           =   0
         Left            =   225
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Standard palette of commonly used colors."
         Top             =   2610
         Value           =   -1  'True
         Width           =   1890
      End
      Begin VB.PictureBox pctPal 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   2325
         Left            =   225
         ScaleHeight     =   155
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   251
         TabIndex        =   11
         Top             =   270
         Width           =   3765
      End
   End
   Begin SMBMaker.dbButton OkButton 
      Height          =   390
      Left            =   4500
      TabIndex        =   1
      ToolTipText     =   "Update the color and close the dialog."
      Top             =   0
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   688
      MouseIcon       =   "frmColorPick.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Others          =   $"frmColorPick.frx":001C
      OthersPresent   =   -1  'True
   End
   Begin SMBMaker.dbButton CancelButton 
      Cancel          =   -1  'True
      Height          =   390
      Left            =   4500
      TabIndex        =   2
      ToolTipText     =   "Discard color change and close the dialog."
      Top             =   390
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   688
      MouseIcon       =   "frmColorPick.frx":006B
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Others          =   $"frmColorPick.frx":0087
      OthersPresent   =   -1  'True
   End
End
Attribute VB_Name = "frmColorPick"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Pnts() As POINTAPI
Dim Comps() As Long
Dim tmpComps() As Long
Dim Plt() As Long 'here, in rgb0 format
Dim DisUpdate As Boolean
Dim Tips() As String

Dim MD As Integer

Private Sub CancelButton_Click()
Me.Tag = "C"
Me.Hide
End Sub

Private Sub Form_Load()
ReDim Pnts(0 To 2)
ReDim Comps(0 To 2)
ReDim tmpComps(0 To 2)
ExtractResPal Plt, "COLORPICK"
UpdateTips
dbLoadCaptions
End Sub

Private Sub UpdateTips()
Dim i As Long
ReDim Tips(0 To UBound(Plt))
For i = 0 To UBound(Plt)
    Tips(i) = GenerateColorTip(ConvertColorLng(Plt(i)))
Next i
End Sub

Private Sub ExtractSysColors()
Dim Vals() As String, Names() As String
Dim tmpArr() As String
Dim i As Long
gReg.GetAllValues HKEY_CURRENT_USER, "Control Panel\Colors", Names, Vals
ReDim Plt(0 To UBound(Names))
ReDim Tips(0 To UBound(Names))
For i = 0 To UBound(Names)
    tmpArr = Split(Vals(i), " ")
    Plt(i) = RGB(Val(tmpArr(0)), Val(tmpArr(1)), Val(tmpArr(2)))
    Tips(i) = Names(i)
Next i
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If UnloadMode = vbFormControlMenu Then
    Cancel = True
    CancelButton_Click
    Exit Sub
End If
End Sub

Private Sub OkButton_Click()
Me.Tag = ""
Me.Hide
End Sub

Private Sub OptPal_Click(Index As Integer)
Select Case Index
    Case 0
        ExtractResPal Plt, "COLORPICK"
        UpdateTips
    Case 1
        ExtractSysColors
End Select
pctPal.Refresh
End Sub

Private Sub pctAll_Click()
On Error GoTo eh
With CDl
    .CancelError = True
    .ColorFlags = cdlCCRGBInit
    .hWndOwner = Me.hWnd
    .Flags = 0
    .Color = GetColor
    .ShowColor UseWindowsDialog:=True
    SetColor .Color
End With
Exit Sub
eh:
If Err.Number = dbCWS Then
    Exit Sub
End If
MsgBox Err.Description
End Sub


Private Sub pctAll_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
MD = MD Or Button
End Sub

Private Sub pctAll_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim a As POINTAPI
GetCursorPos a
If Button = 2 Then
    SetColor ConvertColorLng(CapturePixel(a.X, a.Y)), True
End If
End Sub

Private Sub pctAll_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim a As POINTAPI
GetCursorPos a
If Button = 2 And MD = 2 Then
    MD = MD And (Not Button)
    SetColor ConvertColorLng(CapturePixel(a.X, a.Y))
Else
    MD = MD And (Not Button)
    SetColor GetColor, True
End If
End Sub

Private Sub pctAll_Paint()
pctC_Paint (-1)
End Sub

Private Sub pctC_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
    Pnts(Index).X = X
    Pnts(Index).Y = Y
    'pctC_Paint Index
ElseIf Button = 2 Then
    Pnts(Index).X = X - MaxArr(Comps)
    Pnts(Index).Y = Y
End If
End Sub

Private Sub pctC_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Long
If Button = 1 Then
    tmpComps(Index) = Comps(Index) + (X - Pnts(Index).X)
    ValidateCompRef tmpComps(Index)
    pctC_Paint Index
    UpdateTexts tmpComps
ElseIf Button = 2 Then
    For i = 0 To 2
        tmpComps(i) = Comps(i) * (X - Pnts(Index).X) \ MaxArr(Comps)
        ValidateCompRef tmpComps(i)
    Next i
    pctC_Paint 3
    UpdateTexts tmpComps
End If
End Sub

Private Sub pctC_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Long
If Button = 1 Then
    tmpComps(Index) = Comps(Index) + (X - Pnts(Index).X)
    ValidateCompRef tmpComps(Index)
    Comps(Index) = tmpComps(Index)
    pctC_Paint Index
    UpdateTexts Comps
ElseIf Button = 2 Then
    For i = 0 To 2
        tmpComps(i) = Comps(i) * (X - Pnts(Index).X) \ MaxArr(Comps)
        ValidateCompRef tmpComps(i)
    Next i
    Comps = tmpComps
    pctC_Paint 3
    UpdateTexts Comps
End If
End Sub

Private Sub ValidateCompRef(ByRef Comp As Long)
If Comp < 0 Then Comp = 0
If Comp > 255 Then Comp = 255
End Sub

Private Sub pctC_Paint(Index As Integer)
Dim CX As Long
Dim i As Long
Static RGBMultipliers(0 To 2) As Long
If RGBMultipliers(0) = 0 Then
    RGBMultipliers(0) = &H1&
    RGBMultipliers(1) = &H100&
    RGBMultipliers(2) = &H10000
End If
If Index >= 0 Then
    For i = 0 To 2
        If Index = 3 Or i = Index Then
            With pctC(i)
                CX = .ScaleWidth \ 2
                pctC(i).Line (0, 0)-(CX, .ScaleHeight), tmpComps(i) * RGBMultipliers(i), BF
                pctC(i).Line (CX, 0)-(.ScaleWidth, .ScaleHeight), Comps(i) * RGBMultipliers(i), BF
            End With
        End If
    Next i
End If
With pctAll
    CX = .ScaleWidth \ 2
    pctAll.Line (0, 0)-(CX, .ScaleHeight), RGB(tmpComps(0), tmpComps(1), tmpComps(2)), BF
    pctAll.Line (CX, 0)-(.ScaleWidth, .ScaleHeight), RGB(Comps(0), Comps(1), Comps(2)), BF
End With

End Sub

Private Function MaxArr(ByRef Arr() As Long)
Dim i As Long
Dim m As Long
For i = LBound(Arr) To UBound(Arr)
    If Arr(i) > m Then m = Arr(i)
Next i
If m = 0 Then m = 1
MaxArr = m
End Function

Private Sub UpdateTexts(ByRef cmpArr() As Long, _
                        Optional ByVal UpdateTxtHex As Boolean = True)
Dim i As Long
DisUpdate = True
For i = 0 To 2
    txtC(i).Text = CStr(cmpArr(i))
    txtC(i).Refresh
Next i
If UpdateTxtHex Then
txtHex.Text = RGBToWebColor(cmpArr(0), cmpArr(1), cmpArr(2))
End If
DisUpdate = False
End Sub

Public Function RGBToWebColor(ByVal r As Long, _
                               ByVal g As Long, _
                               ByVal b As Long) As String
RGBToWebColor = "#" + VedNullStr(Hex$(r), 2) + _
                       VedNullStr(Hex$(g), 2) + _
                       VedNullStr(Hex$(b), 2)
End Function

Public Function VedNullStr(ByRef St As String, _
                           ByVal NumberOfDigits As Long) As String
VedNullStr = String$(NumberOfDigits - Len(St), "0") + St
End Function

Private Sub UpDateComps(ByRef cmpArr() As Long)
Dim i As Long
For i = 0 To 2
    cmpArr(i) = Val(txtC(i).Text)
    ValidateCompRef cmpArr(i)
Next i
End Sub

Private Sub pctPal_DblClick()
If OkButton.Enabled Then OkButton_Click
End Sub

Private Sub pctPal_Paint()
Dim X As Long, Y As Long
Dim nY As Long, nX As Long
Dim n As Long
nX = 8
nY = -Int(-(UBound(Plt) + 1) / nX)
For Y = 0 To nY - 1
    For X = 0 To nX - 1
        n = Y * nX + X
        If n > UBound(Plt) Then
            pctPal.Line (PalScaleX(X, nX) + 1, PalScaleY(Y, nY) + 1)- _
                        (PalScaleX(X + 1, nX) - 2, PalScaleY(Y + 1, nY) - 2), vbWhite, B
        Else
            pctPal.Line (PalScaleX(X, nX), PalScaleY(Y, nY))- _
                        (PalScaleX(X + 1, nX) - 1, PalScaleY(Y + 1, nY) - 1), Plt(n), BF
        End If
    Next X
Next Y
End Sub

Public Function PalScaleX(ByVal px As Long, ByVal nX As Long) As Long
PalScaleX = -Int(-px * pctPal.ScaleWidth / nX)
End Function

Public Function PalScaleY(ByVal py As Long, ByVal nY As Long) As Long
PalScaleY = -Int(-py * pctPal.ScaleHeight / nY)
End Function

Public Function PctToPal(ByVal sx As Long, ByVal sy As Long) As Long
Dim n As Long
Dim nX As Long, nY As Long
If sx < 0 Or sy < 0 Or sx >= pctPal.ScaleWidth Or sy >= pctPal.ScaleHeight Then
    PctToPal = -1
Else
    nX = 8
    nY = -Int(-(UBound(Plt) + 1) / nX)
    n = (sx * nX \ pctPal.ScaleWidth) + (sy * nY \ pctPal.ScaleHeight) * nX
    If n > UBound(Plt) Then
        n = -1
    End If
    PctToPal = n
End If
End Function

Private Sub pctPal_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim n As Long
If Button = 1 Then
    n = PctToPal(X, Y)
    If n = -1 Then
        tmpComps = Comps
        UpdateTexts Comps
        pctC_Paint 3
    Else
        SetColor Plt(n), True
    End If
End If
End Sub

Private Sub pctPal_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim n As Long
n = PctToPal(X, Y)
If Button = 1 Then
    If n = -1 Then
        tmpComps = Comps
        UpdateTexts Comps
        pctC_Paint 3
    Else
        SetColor Plt(n), True
    End If
End If
If n = -1 Then
    pctPal.ToolTipText = GRSF(2369)
Else
    pctPal.ToolTipText = Tips(n)
End If
End Sub

Private Sub pctPal_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim n As Long
If Button = 1 Then
    n = PctToPal(X, Y)
    If n = -1 Then
        tmpComps = Comps
        UpdateTexts Comps
        pctC_Paint 3
    Else
        SetColor Plt(n), False
    End If
End If
End Sub

Private Sub txtC_Change(Index As Integer)
If DisUpdate Then Exit Sub
UpDateComps tmpComps
pctC_Paint Index
End Sub

Private Sub txtC_Validate(Index As Integer, Cancel As Boolean)
If DisUpdate Then Exit Sub
UpDateComps Comps
UpdateTexts Comps
pctC_Paint Index
End Sub

Friend Sub SetColor(ByVal lngColor As Long, _
                    Optional ByVal ToTemp As Boolean = False, _
                    Optional ByVal UpdateTxtHex As Boolean = True)
Dim rgb1 As RGBQUAD
GetRgbQuadEx lngColor, rgb1
tmpComps(0) = rgb1.rgbRed
tmpComps(1) = rgb1.rgbGreen
tmpComps(2) = rgb1.rgbBlue
UpdateTexts tmpComps, UpdateTxtHex
If Not ToTemp Then
    Comps = tmpComps
End If
pctC_Paint 3
End Sub

Friend Function GetColor(Optional ByVal FromTmp As Boolean = False) As Long
If FromTmp Then
    GetColor = RGB(tmpComps(0), tmpComps(1), tmpComps(2))
Else
    GetColor = RGB(Comps(0), Comps(1), Comps(2))
End If
End Function

Friend Sub dbLoadCaptions()
Dim i As Long
Me.Caption = GRSF(2370)
OptPal(0).Caption = GRSF(2372)
OptPal(1).Caption = GRSF(2373)

For i = 0 To 2
    pctC(i).ToolTipText = GRSF(2375 + i) + " " + GRSF(2378)
Next i

pctAll.ToolTipText = GRSF(2379)
End Sub

Private Sub txtHex_Change()
If DisUpdate Then Exit Sub
SetColor WebColorToLong(txtHex.Text), True, False
End Sub

Private Sub txtHex_LostFocus()
If DisUpdate Then Exit Sub
SetColor WebColorToLong(txtHex.Text)
End Sub

Public Function WebColorToLong(ByRef St As String) As Long 'rgb color
Dim tmp As String
Dim clr As Long
Dim r As Long, g As Long, b As Long
Dim i As Long
If DisUpdate Then Exit Function
tmp = Replace(St, " ", "")
If Len(tmp) > 0 Then
    i = InStr(1, tmp, "#")
    If i > 0 Then
        tmp = Mid$(tmp, i + 1)
    End If
    tmp = UCase$(tmp)
    For i = 1 To Len(tmp)
        Select Case Mid$(tmp, i, 1)
            Case "0" To "9"
            Case "A" To "F"
            Case Else
                Exit For
        End Select
    Next i
    tmp = Mid$(tmp, 1, i - 1)
    If Len(tmp) >= 6 Then
        r = Val("&H" + Mid$(tmp, 1, 2))
        g = Val("&H" + Mid$(tmp, 3, 2))
        b = Val("&H" + Mid$(tmp, 5, 2))
    Else
        If Len(tmp) < 3 Then
            tmp = VedNullStr(tmp, 3)
        End If
        r = Val("&H" + Mid$(tmp, 1, 1) + Mid$(tmp, 1, 1))
        g = Val("&H" + Mid$(tmp, 2, 1) + Mid$(tmp, 2, 1))
        b = Val("&H" + Mid$(tmp, 3, 1) + Mid$(tmp, 3, 1))
    End If
End If
clr = RGB(r, g, b)
WebColorToLong = clr
End Function
