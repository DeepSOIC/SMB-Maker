VERSION 5.00
Begin VB.Form frmWavesCreate 
   Caption         =   "Field Images. Place the sources."
   ClientHeight    =   7290
   ClientLeft      =   165
   ClientTop       =   480
   ClientWidth     =   10290
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   486
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   686
   StartUpPosition =   1  'CenterOwner
   Begin SMBMaker.ctlResourcizator Resr1 
      Left            =   7770
      Top             =   1410
      _ExtentX        =   1164
      _ExtentY        =   741
      ResID           =   9901
   End
   Begin VB.PictureBox StatusBar 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   900
      ScaleHeight     =   31
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   443
      TabIndex        =   2
      Top             =   4470
      Width           =   6705
      Begin SMBMaker.ctlColor clrCurColor 
         Height          =   285
         Left            =   0
         TabIndex        =   4
         Top             =   0
         Width           =   330
         _ExtentX        =   582
         _ExtentY        =   503
         ColorFormat     =   0
         Color           =   33023
      End
      Begin SMBMaker.dbButton btnToggleView 
         Height          =   360
         Left            =   5370
         TabIndex        =   5
         Top             =   30
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   635
         MouseIcon       =   "frmWavesCreate.frx":0000
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Others          =   $"frmWavesCreate.frx":001C
         OthersPresent   =   -1  'True
      End
      Begin VB.Label StatusText 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   495
         TabIndex        =   3
         Top             =   15
         Width           =   3915
      End
   End
   Begin SMBMaker.dbFrame Toolbar 
      Height          =   30
      Left            =   150
      TabIndex        =   1
      Top             =   15
      Width           =   7140
      _ExtentX        =   12594
      _ExtentY        =   53
      EAC             =   0   'False
   End
   Begin VB.PictureBox pctSpace 
      AutoRedraw      =   -1  'True
      Height          =   3930
      Left            =   150
      ScaleHeight     =   258
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   472
      TabIndex        =   0
      Top             =   390
      Width           =   7140
   End
   Begin VB.Menu mnuSources 
      Caption         =   "Sources"
      Begin VB.Menu mnuClear 
         Caption         =   "Clear field	Ctrl+N"
      End
      Begin VB.Menu mnuAddWS 
         Caption         =   "Add source..."
      End
      Begin VB.Menu srcSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "Save to file"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuLoad 
         Caption         =   "Load from file"
         Visible         =   0   'False
      End
      Begin VB.Menu srcSep2 
         Caption         =   "-"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "Edit"
      Begin VB.Menu mnuUndo 
         Caption         =   "Undo	Ctrl+Z"
      End
   End
   Begin VB.Menu mnuSelection 
      Caption         =   "Selection"
      Begin VB.Menu mnuSelSelectAll 
         Caption         =   "Select all"
      End
      Begin VB.Menu mnuSelDeselectAll 
         Caption         =   "Deselect all"
      End
      Begin VB.Menu mnuSelInvert 
         Caption         =   "Invert selection"
      End
      Begin VB.Menu selSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuselDelete 
         Caption         =   "Delete selected"
      End
      Begin VB.Menu mnuSelClone 
         Caption         =   "Clone selected"
      End
      Begin VB.Menu selSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSelEdit 
         Caption         =   "Edit selected"
      End
      Begin VB.Menu mnuSelColorize 
         Caption         =   "Colorize selection"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "View"
      Begin VB.Menu mnuVReset 
         Caption         =   "Reset scale"
      End
      Begin VB.Menu mnuVActualZoom 
         Caption         =   "Zoom 1x"
      End
      Begin VB.Menu mnuVShowSelected 
         Caption         =   "Show selected"
      End
      Begin VB.Menu ViewSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuVColors 
         Caption         =   "Show colors"
      End
      Begin VB.Menu mnuVWaveLen 
         Caption         =   "Show wavelength"
      End
   End
   Begin VB.Menu mnuMode 
      Caption         =   "Mode"
      Begin VB.Menu infMode1 
         Caption         =   "(Cursor mode)"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuModes 
         Caption         =   "Creating	C"
         Index           =   0
         Tag             =   "0"
      End
      Begin VB.Menu mnuModes 
         Caption         =   "Selecting	S, Shift"
         Index           =   1
         Tag             =   "1"
      End
      Begin VB.Menu mnuModes 
         Caption         =   "Editing	E"
         Index           =   2
         Tag             =   "2"
      End
      Begin VB.Menu mnuModes 
         Caption         =   "Moving	M, Ctrl"
         Index           =   3
         Tag             =   "3"
      End
      Begin VB.Menu mnuModes 
         Caption         =   "Navigating	N, Space"
         Index           =   4
         Tag             =   "4"
      End
      Begin VB.Menu mnuModes 
         Caption         =   "Colorizing"
         Index           =   5
         Tag             =   "5"
      End
   End
   Begin VB.Menu mnuTransform 
      Caption         =   "Transform"
      Visible         =   0   'False
      Begin VB.Menu mnuTMove 
         Caption         =   "Move by..."
      End
      Begin VB.Menu mnuTRotate 
         Caption         =   "Rotate..."
      End
      Begin VB.Menu mnuTScale 
         Caption         =   "Scale..."
      End
      Begin VB.Menu mnuTFlipVert 
         Caption         =   "Flip vertically"
      End
      Begin VB.Menu mnuTFlipHorz 
         Caption         =   "Flip horizontally"
      End
      Begin VB.Menu tranSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTRandomize 
         Caption         =   "Randomize"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuHlpOrient 
         Caption         =   "Orientation..."
      End
      Begin VB.Menu mnuHlpModes 
         Caption         =   "Modes..."
      End
   End
   Begin VB.Menu mnuDone 
      Caption         =   "Done"
      Begin VB.Menu mnuDoneWaves 
         Caption         =   "Draw Waves"
      End
      Begin VB.Menu mnuDoneWavesOpts 
         Caption         =   "Options"
         Begin VB.Menu mnuDoneWavesOptsFalldown 
            Caption         =   "Falldown factor..."
         End
         Begin VB.Menu mnuDoneWavesOptsAbsolute 
            Caption         =   "Absolute"
         End
      End
      Begin VB.Menu DoneSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDoneLines 
         Caption         =   "Draw Field Lines"
      End
      Begin VB.Menu mnuDoneLinesOpts 
         Caption         =   "Options"
         Begin VB.Menu mnuDoneLinesOptsPower 
            Caption         =   "Field power"
         End
      End
      Begin VB.Menu DoneSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDoneClose 
         Caption         =   "Close"
      End
   End
End
Attribute VB_Name = "frmWavesCreate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const SpotR As Long = 3
Const SpotPosClr As Long = vbRed
Const SpotNegClr As Long = vbBlue
Const SpotNeuClr As Long = vbGreen
Const SpotSelClr As Long = vbYellow
Const SpotBorderClr As Long = vbBlack
Const errTMS = 10020

Const WaveLenClr As Long = vbMagenta

Dim DisplayMode As dbWSDM
Dim SL As Long 'ScaleLeft
Dim St As Long 'ScaleTop
Dim SR As Long 'ScaleRight
Dim SB As Long 'ScaleBottom
Dim sW As Long, sH As Long 'ScaleWidth, ScaleHeight
Dim ScaleFactor As Double 'pctX=imgX*scalefactor-SL
Dim WS() As typWaveSource
Dim nWS As Long
Dim SelIndex As Long 'Index of the selected WS
Dim CurColor As RGBQUAD 'the color for the new source
'Dim CurSign As Long
Dim CurWL As Long, CurStrength As Double
Public ImageW As Long, ImageH As Long
Dim CurEditMode As eCursorMode

Dim Undos() As typAryWS
Dim nUndos As Long

Dim mdx As Long, mdy As Long 'MouseDown pos
Dim iMDx As Long, iMDY As Long 'MouseDown pos on image
Dim mmx As Long, mmy As Long  'Pos of previous MouseMove
Dim MDs As Long 'source index under mousedown
Dim MDSF As Double 'MouseDownScaleFactor
Dim MDShift As Integer
Dim MDAryWS() As typWaveSource 'wave sources when mousedown occured
Dim MDEditMode As eCursorMode
Dim MDWasSpace As Boolean
Dim RotAccum As Double
Dim MBState(1 To 4) As Boolean
Dim ClickN As Long
Dim WithEvents pctSpaceMS As clsAntiDblClick
Attribute pctSpaceMS.VB_VarHelpID = -1

Dim LastR As Long

Public Mode As eWPMode
Public mFallDownFactor As Double
Public mFieldPower As Double
Public Enum eWPMode
    wpmWavesPicture
    wpmFieldLines
End Enum


Private Enum eCursorMode
    cmCreate = 0
    cmSel = 1
    cmEdit = 2
    cmMove = 3
    cmNavi = 4
    cmColorize = 5
End Enum

Public Enum dbWSDM 'Wave Sources Display Mode
    wsdmWaveLength = 8 'if not set, use rgb. If set, display wavelen
    wsdmStrength = 1 'Display strength and color
'    wsdmColor = 2 'Display Green channel
'    wsdmB = 4 'Display Blue channel
End Enum

Private Enum WSRadiusIndex
    wsriMax = 0 'used for calculating optimal zoom
    wsriRed = 1
    wsriGreen = 2
    wsriBlue = 3
End Enum

Private Type typAryWS
    WS() As typWaveSource
    nWS As Long
End Type


Private Function GetRadius(ByVal iWS As Long, _
                           ByVal RadiusIndex As WSRadiusIndex)
If iWS > nWS - 1 Or iWS < 0 Then Err.Raise 9, "GetRadius", "Subscript out of range"
With WS(iWS)
    Select Case RadiusIndex
        Case WSRadiusIndex.wsriMax
            If DisplayMode And wsdmStrength Then
                GetRadius = Abs(.Strength)
            ElseIf DisplayMode And wsdmWaveLength Then
                GetRadius = .WaveLength
            End If
'        Case WSRadiusIndex.wsriRed
'            GetRadius = .r
'        Case WSRadiusIndex.wsriGreen
'            GetRadius = .g
'        Case WSRadiusIndex.wsriBlue
'            GetRadius = .b
    End Select
End With
End Function

Private Sub SetRadius(ByVal iWS As Long, _
                      ByVal NewRadius As Double)
Dim rgb1 As RGBQUAD
If iWS > nWS - 1 Or iWS < 0 Then Err.Raise 9, "SetRadius", "Subscript out of range"
CopyMemory rgb1, CurColor, 4&
With WS(iWS)
'    If .r = 0 Then .r = rgb1.rgbRed
'    If .g = 0 Then .g = rgb1.rgbGreen
'    If .b = 0 Then .b = rgb1.rgbBlue
    .Color = CurColor
    If (DisplayMode And wsdmWaveLength) Then
        If .WaveLength = 0 Then .WaveLength = 10
'        If (DisplayMode And wsdmR) <> 0 Then
'            .r = NewRadius
'        End If
'        If (DisplayMode And wsdmG) <> 0 Then
'            .g = NewRadius
'        End If
'        If (DisplayMode And wsdmB) <> 0 Then
'            .b = NewRadius
'        End If
        .WaveLength = NewRadius
    ElseIf DisplayMode And wsdmStrength Then
        SetStrength .Strength, NewRadius, True
    End If
End With
End Sub

Private Sub SetStrength(ByRef Strength As Double, _
                        ByVal vNew As Double, _
                        Optional ByVal PreserveSign As Boolean = True)
Dim Sign As Double
Sign = IIf(PreserveSign, Sgn(Strength), Sgn(vNew))
If Sign = 0# Then
    Sign = Sgn(CurStrength)
    If Sign = 0# Then Sign = 1#
End If
Strength = Sign * MinD(Abs(vNew), 16#)
End Sub


'Uses values stored in SL,SR,ST,SB and creates the scale for them
Public Sub RecalcScale(Optional ByVal CenterView As Boolean = False)
Dim ScaleFactorX As Double, ScaleFactorY As Double
Dim CX As Double, CY As Double
sW = SR - SL
sH = SB - St
If sW = 0 Or sH = 0 Then
    SL = 0
    St = 0
    SR = ImageW
    SB = ImageH
    sW = ImageW
    sH = ImageH
End If
Debug.Assert sW > 0 And sH > 0
ScaleFactorX = pctSpace.ScaleWidth / sW
ScaleFactorY = pctSpace.ScaleHeight / sH
If ScaleFactorX < ScaleFactorY Then
    ScaleFactor = ScaleFactorX
    If CenterView Then
        CY = (SB + St) / 2
        St = CY - (pctSpace.ScaleHeight / 2) / ScaleFactor
        SB = St + pctSpace.ScaleHeight / ScaleFactor
    End If
Else
    ScaleFactor = ScaleFactorY
    If CenterView Then
        CX = (SR + SL) / 2
        SL = CX - (pctSpace.ScaleWidth / 2) / ScaleFactor
        SR = SL + pctSpace.ScaleWidth / ScaleFactor
    End If
End If
End Sub

Private Function MinD(a As Double, b As Double) As Double
If a < b Then MinD = a Else MinD = b
End Function

Private Function ImageToSpaceX(ByVal X As Double) As Double
ImageToSpaceX = (X - SL) * ScaleFactor
End Function

Private Function ImageToSpaceY(ByVal Y As Double) As Double
ImageToSpaceY = (Y - St) * ScaleFactor
End Function

Private Function SpaceToImageX(ByVal X As Double) As Double
SpaceToImageX = X / ScaleFactor + SL
End Function

Private Function SpaceToImageY(ByVal Y As Double) As Double
SpaceToImageY = Y / ScaleFactor + St
End Function

Private Sub DrawStar(ByVal X As Long, ByVal Y As Long, _
                     Optional ByVal Selected As Boolean = False, _
                     Optional ByVal Sign As Long = 0)
Dim clr As Long
pctSpace.DrawMode = vbCopyPen
'border
pctSpace.Line (X - SpotR, Y - SpotR)- _
              (X + SpotR, Y + SpotR), IIf(Selected, SpotSelClr, SpotBorderClr), B
'filling
Select Case Sign
    Case 0
        clr = SpotNeuClr
    Case 1
        clr = SpotPosClr
    Case -1
        clr = SpotNegClr
End Select
pctSpace.Line (X - SpotR + 1, Y - SpotR + 1)- _
              (X + SpotR - 1, Y + SpotR - 1), clr, BF
End Sub

Private Sub DrawWS(ByRef WS As typWaveSource, _
                   Optional ByVal bDrawStar As Boolean = True, _
                   Optional ByVal bDrawCircles As Boolean = True)
Dim X As Long, Y As Long
X = ImageToSpaceX(WS.Pos.X)
Y = ImageToSpaceY(WS.Pos.Y)
If bDrawCircles Then
    If CBool(DisplayMode And wsdmWaveLength) Then
        pctSpace.DrawMode = vbXorPen
        pctSpace.Circle (X, Y), WS.WaveLength * ScaleFactor, WaveLenClr
    ElseIf CBool(DisplayMode And wsdmStrength) Then
        pctSpace.DrawMode = vbXorPen
        pctSpace.FillStyle = vbFSSolid
        pctSpace.FillColor = RGB(WS.Color.rgbRed, WS.Color.rgbGreen, WS.Color.rgbBlue)
        pctSpace.Circle (X, Y), Abs(WS.Strength * ScaleFactor), vbYellow
        pctSpace.FillStyle = FillStyleConstants.vbFSTransparent
        
        'pctSpace.DrawMode = vbXorPen
        'pctSpace.Circle (x, y), Abs(WS.Strength * ScaleFactor), vbRed
    End If
End If

If bDrawStar Then
    DrawStar X, Y, WS.Selected, Sgn(WS.Strength)
End If
End Sub

Private Sub DrawPicture()
Dim x1 As Long, y1 As Long
Dim x2 As Long, y2 As Long
x1 = ImageToSpaceX(0)
y1 = ImageToSpaceY(0)
x2 = ImageToSpaceX(ImageW) - 1
y2 = ImageToSpaceY(ImageH) - 1
If x2 < x1 Then x2 = x1
If y2 < y1 Then y2 = y1
    
pctSpace.DrawMode = vbCopyPen
pctSpace.Line (x1, y1)- _
              (x2, y2), vbBlack, B

If x2 > x1 + 1 And y2 > y1 + 1 Then
    pctSpace.Line (x1 + 1, y1 + 1)- _
                  (x2 - 1, y2 - 1), &H606060, BF
End If
End Sub

Private Function IsWSVisible(ByVal Index As Long) As Boolean
Dim X As Long, Y As Long
Dim r As Long
X = WS(Index).Pos.X
Y = WS(Index).Pos.Y
r = GetRadius(Index, wsriMax) * ScaleFactor
X = ImageToSpaceX(X)
Y = ImageToSpaceY(Y)
IsWSVisible = (X + r >= 0) And _
              (Y + r >= 0) And _
              (X - r < pctSpace.ScaleWidth) And _
              (Y - r < pctSpace.ScaleHeight)
End Function

Public Sub UpdateView()
Dim i As Long
pctSpace.Cls
DrawPicture
For i = 0 To nWS - 1
    If IsWSVisible(i) Then
        DrawWS WS(i)
    End If
Next i
End Sub

Public Sub SetnWS(ByVal NewN As Long)
If NewN < 0 Then Err.Raise 5
If NewN > 256 Then
    Err.Raise errTMS, "SetnWS", "Too many sources!"
End If
If NewN = 0 Then
    Erase WS
ElseIf nWS = 0 Then
    ReDim WS(0 To NewN - 1)
Else
    ReDim Preserve WS(0 To NewN - 1)
End If
nWS = NewN
End Sub

Friend Sub MakeWS(ByVal x1 As Long, ByVal y1 As Long, _
                  ByVal x2 As Long, ByVal y2 As Long, _
                  ByRef WS As typWaveSource, _
                  ByVal SetOthers As Boolean, _
                  ByVal Shift As Long)
Dim r As Double
WS.Selected = False
If SetOthers Then
    With WS.Pos
        .X = SpaceToImageX(x1)
        .Y = SpaceToImageY(y1)
    End With
End If
With WS.Pos
    x1 = ImageToSpaceX(.X)
    y1 = ImageToSpaceY(.Y)
End With
r = Sqr((x2 - x1) ^ 2 + (y2 - y1) ^ 2)
With WS
    If SetOthers Then
        .Strength = IIf(CBool(Shift And 4), -CurStrength, CurStrength)
        .WaveLength = LastR
        .Color = CurColor
    End If
    r = r / ScaleFactor
    If (DisplayMode And wsdmWaveLength) <> 0 Then
        .WaveLength = r
        If .WaveLength = 0 Then .WaveLength = LastR
    Else
        SetStrength .Strength, IIf(Shift And 4, -r, r), False
'        If r > 255 Then r = 255
'        If (DisplayMode And wsdmR) <> 0 Then
'            If CurSign <> 0 Then
'                .r = r * CurSign
'            Else
'                .r = r
'            End If
'        End If
'        If (DisplayMode And wsdmG) <> 0 Then
'            .g = r
'        End If
'        If (DisplayMode And wsdmB) <> 0 Then
'            .b = r
'        End If
    End If
End With
End Sub

'in space coords
Public Function NearestWS(ByVal X As Long, ByVal Y As Long, _
                          Optional ByVal OnlySelected As Boolean = False) As Long
Dim sx As Long, sy As Long
Dim i As Long
Dim MinDist As Currency, ind As Long
Dim Dist As Currency
If nWS = 0 Then NearestWS = -1
MinDist = 922337203685477@
ind = -1
For i = 0 To nWS - 1
    If Not OnlySelected And IsWSVisible(i) Or OnlySelected And WS(i).Selected Then
        sx = ImageToSpaceX(WS(i).Pos.X)
        sy = ImageToSpaceY(WS(i).Pos.Y)
        Dist = CCur(sx - X) * CCur(sx - X) + _
               CCur(sy - Y) * CCur(sy - Y)
        If Dist < MinDist Then
            MinDist = Dist
            ind = i
        End If
    End If
Next i
NearestWS = ind
End Function

Private Function ChangeWS(ByRef WS As typWaveSource, _
                          ByVal x1 As Long, ByVal y1 As Long, _
                          ByVal x2 As Long, ByVal y2 As Long, _
                          ByVal Shift As Long)
MakeWS x1, y1, x2, y2, WS, False, Shift
End Function

'moves selected wavesources by dx and dy
Private Sub MoveSources(ByVal dx As Long, ByVal dy As Long, _
                        Optional ByVal Index As Long = -1)
Dim i As Long
Dim idx As Long, idy As Long
Dim OneSelected As Boolean
idx = dx / ScaleFactor
idy = dy / ScaleFactor
OneSelected = False
For i = 0 To nWS - 1
    If WS(i).Selected Then
        OneSelected = True
        With WS(i).Pos
            .X = .X + idx
            .Y = .Y + idy
        End With
    End If
Next i
If Not OneSelected And Index >= 0 And Index < nWS Then
    With WS(Index).Pos
        .X = .X + idx
        .Y = .Y + idy
    End With
End If
End Sub

Private Function Min(ByVal a As Long, ByVal b As Long) As Long
If a < b Then Min = a Else Min = b
End Function

Private Function Max(ByVal a As Long, ByVal b As Long) As Long
If a > b Then Max = a Else Max = b
End Function

Public Sub AutoScale(Optional ByVal UseOnlySelected As Boolean = False)
Dim i As Long
Dim fx As Long, fy As Long
Dim tx As Long, ty As Long
Dim r As Long
Dim nSelected As Long
For i = 0 To nWS - 1
    If WS(i).Selected Or Not UseOnlySelected Then
        With WS(i).Pos
            r = GetRadius(i, wsriMax)
            fx = Min(fx, .X - r)
            fy = Min(fy, .Y - r)
            tx = Max(tx, .X + r)
            ty = Max(ty, .Y + r)
        End With
        nSelected = nSelected + 1&
    End If
Next i
If UseOnlySelected And nSelected > 0 Then
    'do nothing - image should not be added to them
Else
    If fx > 0 Then fx = 0
    If fy > 0 Then fy = 0
    If tx < ImageW - 1 Then tx = ImageW - 1
    If ty < ImageH - 1 Then ty = ImageH - 1
End If
SL = fx
SR = tx
St = fy
SB = ty
RecalcScale True
Refr
End Sub

Public Sub Refr(Optional ByVal Cls As Boolean = False)
Dim i As Long
If Cls Then
    pctSpace.BackColor = 0
    pctSpace.Cls
Else
    pctSpace.DrawMode = vbCopyPen
    pctSpace.Line (0, 0)-(pctSpace.ScaleWidth - 1, pctSpace.ScaleHeight - 1), 0, BF
End If
DrawPicture
For i = 0 To nWS - 1
    If IsWSVisible(i) Then
        DrawWS WS(i), False, True
    End If
Next i
For i = 0 To nWS - 1
    If IsWSVisible(i) Then
        DrawWS WS(i), True, False
    End If
Next i
End Sub

Private Sub btnToggleView_Click()
If DisplayMode And wsdmStrength Then
    mnuVWaveLen_Click
Else
    mnuVColors_Click
End If
pctSpace.SetFocus
End Sub

Private Sub clrCurColor_Change()
GetRgbQuadEx clrCurColor, CurColor
End Sub

Private Sub ctlResourcizator1_GotFocus()

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Dim ShiftKeyCode As Long
If KeyCode = 0 Then Exit Sub
ShiftKeyCode = GetShiftKeyCode(KeyCode, Shift)
MoveMouse
If KeyCode < 16 Or KeyCode > 18 And Not KeyCode = 32 Then
    If IsKey(1) Or IsKey(2) Or IsKey(4) Then
             vtBeep
        Exit Sub
    End If
End If
Select Case ShiftKeyCode
    Case 46 'del
        DeleteSelected
        Refr
        pctSpace.Refresh
    Case 602 'Ctrl+Z
        mnuUnDo_Click
    Case 590 'Ctrl-N
        mnuClear_Click
    Case 83 'S
        SetCursorMode cmSel
    Case 67 'C
        SetCursorMode cmCreate
    Case 77 'M
        SetCursorMode cmMove
    Case 69 'E
        SetCursorMode cmEdit
    Case 80 'P
        SetCursorMode cmColorize
    Case 577 'Ctrl-A
        mnuSelSelectAll_Click
End Select
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 0 Then Exit Sub
MoveMouse
End Sub

Private Sub Form_Load()
Set pctSpaceMS = New clsAntiDblClick
'This is temporary
LoadCaptions
CurColor.rgbBlue = 255
CurColor.rgbGreen = 255
CurColor.rgbRed = 255
UpdateCurColor

CurStrength = 16

'SetnWS 1
'WS(0).Color.rgbRed = 255
'WS(0).Color.rgbGreen = 128
'WS(0).Color.rgbBlue = 64
'WS(0).Strength = 16
'WS(0).WaveLength = 10
DisplayMode = wsdmWaveLength
LoadFromReg
LoadSettings


SetCursorMode cmCreate

ImageW = 100
ImageH = 100
AutoScale
Refr True
pctSpace.Refresh

UpdateIndiDM
End Sub

Public Sub LoadCaptions()
Resr1.LoadCaptions
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim Answ As VbMsgBoxResult
On Error GoTo eh
If UnloadMode = 0 Then
    UserCancel
    Cancel = True
End If
If UnloadMode = QueryUnloadConstants.vbAppWindows Then
    SaveToReg
    SaveSettings
End If
Exit Sub
eh:
Cancel = MsgError(Err, vbCritical Or vbOKCancel) = vbCancel
End Sub

Friend Sub Form_Resize()
On Error Resume Next
Toolbar.Move 0, 0, ScaleWidth
StatusBar.Move 0, ScaleHeight - StatusBar.Height, ScaleWidth
pctSpace.Move 0, Toolbar.Top + Toolbar.Height, Me.ScaleWidth, StatusBar.Top - (Toolbar.Top + Toolbar.Height)
End Sub

Private Sub mnuAddWS_Click()
Dim tWS() As typWaveSource
Dim i As Long
On Error GoTo eh
ReDim tWS(0 To 0)
tWS(0).WaveLength = LastR
tWS(0).Strength = CurStrength
tWS(0).Color = CurColor
tWS(0).Selected = True
frmWaveSource.EditWS tWS, 1, False, True
BUD
SetnWS nWS + 1
WS(nWS - 1) = tWS(0)
For i = 0 To nWS - 2
    WS(i).Selected = False
Next i
Refr
Exit Sub
eh:
MsgError
End Sub

Private Sub mnuClear_Click()
BUD
nWS = 0
Erase WS
Refr
pctSpace.Refresh
End Sub

'Selects a source. If not deselectthers then toggles selection.
'If ws already selected and none of others is selected and
'   deselectothers then it deselects ws.
Private Sub SelectWS(ByVal Index As Long, _
                     Optional ByVal DeselectOthers As Boolean = True)
Dim i As Long
Dim OthersSelected As Boolean
If Index < 0 Or Index >= nWS Then Exit Sub
If DeselectOthers Then
    OthersSelected = False
    For i = 0 To nWS - 1
        If i <> Index Then
            If WS(i).Selected Then OthersSelected = True
            WS(i).Selected = False
        End If
    Next i
    WS(Index).Selected = Not WS(Index).Selected Or OthersSelected
Else
    WS(Index).Selected = Not WS(Index).Selected
End If
End Sub

Private Sub mnuDoneClose_Click()
CancelButton_Click
End Sub

Private Sub mnuDoneLines_Click()
Mode = wpmFieldLines
OK
End Sub

Private Sub mnuDoneLinesOptsPower_Click()
On Error GoTo eh
'"Please input the force decreasing factor. The law will be E=q/r^power. 2 is usual electrical forcefield, but the 2-dimensional picture of it cannot be correct. It may look odd. 1 is the power for 2-dimensional space, and it produces nice picture."
EditNumber mFieldPower, Message:=2583, _
           MinValue:=-1, MaxValue:=9
Exit Sub
eh:
MsgError
End Sub

Private Sub mnuDoneWaves_Click()
Mode = wpmWavesPicture
OK
End Sub

Private Sub mnuDoneWavesOptsAbsolute_Click()
mnuDoneWavesOptsAbsolute.Checked = Not mnuDoneWavesOptsAbsolute.Checked
End Sub

Private Sub mnuDoneWavesOptsFalldown_Click()
On Error GoTo eh
'2584= "Please input the falldown factor for waves. Falldown is defined as exp(-r/wl*f), where wl is wavelength, f is falldown factor. If 0, there's no falldown at all."
EditNumber mFallDownFactor, Message:=2584, _
           MinValue:=-1, MaxValue:=3
Exit Sub
eh:
MsgError
End Sub

Private Sub mnuHlpModes_Click()
'2597=Creating\nUsed to create new sources. To create the negative char...
dbMsgBox 2597, vbInformation
End Sub

Private Sub mnuHlpOrient_Click()
'2585=You can navigate around by zooming and moving the point of view....
dbMsgBox 2585, vbInformation
End Sub

Private Sub mnuModes_Click(Index As Integer)
SetCursorMode Val(mnuModes(Index).Tag)
UpdateStatus
End Sub
'
'Private Sub mnuModesNav_Click()
'dbMsgBox "Use the mouse holding Space to navigate (move and zoom in/out)." + vbNewLine + _
'         "Use left mouse button to move." + vbNewLine + _
'         "Rotate pointer clockwise/counterclockwise holding right mouse button to zoom in/out." + vbNewLine + _
'         "All the actions above should be performed holding Space key. Navigate mode cannot be set constantly, the menu item is provided only for help.", vbInformation
'End Sub

Private Sub mnuSelClone_Click()
Dim i As Long
Dim bb As Boolean
For i = 0 To nWS - 1
    If WS(i).Selected Then
        If Not bb Then BUD
        bb = True
        SetnWS nWS + 1
        WS(nWS - 1) = WS(i)
        WS(nWS - 1).Pos.X = WS(i).Pos.X + 10
        WS(nWS - 1).Pos.Y = WS(i).Pos.Y + 10
        WS(i).Selected = False
    End If
Next i
Refr
pctSpace.Refresh
End Sub

Private Sub mnuSelColorize_Click()
Dim i As Long
For i = 0 To nWS - 1
    If WS(i).Selected Then
        WS(i).Color = CurColor
    End If
Next i
mnuVColors_Click
End Sub

Private Sub mnuselDelete_Click()
BUD
DeleteSelected
Refr
pctSpace.Refresh
End Sub

Private Sub mnuSelDeselectAll_Click()
Dim i As Long
For i = 0 To nWS - 1
    WS(i).Selected = False
Next i
Refr
pctSpace.Refresh
End Sub

Private Sub mnuSelEdit_Click()
Dim tWS() As typWaveSource
On Error GoTo eh
tWS = WS
frmWaveSource.EditWS tWS, nWS, True
BUD
WS = tWS
Refr
pctSpace.Refresh
Exit Sub
eh:
MsgError
End Sub

Private Sub mnuSelInvert_Click()
Dim i As Long
For i = 0 To nWS - 1
    WS(i).Selected = Not WS(i).Selected
Next i
Refr
pctSpace.Refresh
End Sub

Private Sub mnuSelSelectAll_Click()
Dim i As Long
For i = 0 To nWS - 1
    WS(i).Selected = True
Next i
Refr
pctSpace.Refresh
End Sub

Private Sub mnuUnDo_Click()
Undo
Refr
pctSpace.Refresh
End Sub

Private Sub mnuVActualZoom_Click()
Dim CX As Long, CY As Long
CX = pctSpace.ScaleWidth \ 2
CY = pctSpace.ScaleHeight \ 2
ReZoom ScaleFactor, 1#, _
       SpaceToImageX(CX), SpaceToImageY(CY), _
       CX, CY, _
       True
Refr
pctSpace.Refresh
End Sub

Private Sub mnuVColors_Click()
DisplayMode = DisplayMode And Not wsdmWaveLength Or wsdmStrength
UpdateIndiDM
Refr
pctSpace.Refresh
End Sub

Private Sub UpdateIndiDM()
mnuVColors.Checked = CBool(DisplayMode And wsdmStrength)
mnuVWaveLen.Checked = CBool(DisplayMode And wsdmWaveLength)
If mnuVColors.Checked Then
    btnToggleView.Caption = GRSF(2599) 'mnuvcolors.caption
Else
    btnToggleView.Caption = GRSF(2598) 'mnuVWaveLen.Caption
End If
'mnuModCompFlags(0).Checked = DisplayMode And wsdmR
'mnuModCompFlags(1).Checked = DisplayMode And wsdmG
'mnuModCompFlags(2).Checked = DisplayMode And wsdmB
End Sub

Private Sub mnuVReset_Click()
AutoScale
End Sub

Private Function GetPEditMode(Shift As Integer) As eCursorMode
GetPEditMode = CurEditMode
If IsKey(32) Then
    GetPEditMode = cmNavi
ElseIf Shift = 2 Then
    GetPEditMode = cmMove
ElseIf CBool(Shift And 1) Then
    GetPEditMode = cmSel
    Shift = 1 And CBool(Shift And 2) Or Shift And 4
End If

End Function

Private Sub mnuVShowSelected_Click()
AutoScale True
Refr
pctSpace.Refresh
End Sub

Private Sub mnuVWaveLen_Click()
DisplayMode = DisplayMode And Not wsdmStrength Or wsdmWaveLength
UpdateIndiDM
Refr
pctSpace.Refresh
End Sub

Private Sub pctSpace_DblClick()
pctSpaceMS.DblClick
End Sub

Private Sub pctSpace_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
pctSpaceMS.MouseDown Button, Shift, X, Y
End Sub

Private Sub pctSpace_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
pctSpaceMS.MouseMove Button, Shift, X, Y

End Sub

Private Sub pctSpace_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
pctSpaceMS.MouseUp Button, Shift, X, Y

End Sub

Private Sub pctSpaceMS_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim pEditMode As eCursorMode

On Error GoTo eh

If LastR = 0 Then LastR = 10

If Button > 0 And Button <= 4 Then
    MBState(Button) = True
End If

mdx = X
mdy = Y
iMDx = SpaceToImageX(X)
iMDY = SpaceToImageY(Y)
mmx = mdx
mmy = mdy
MDShift = Shift
MDAryWS = WS
pEditMode = GetPEditMode(Shift)
MDEditMode = pEditMode
RotAccum = 0#

MDs = NearestWS(X, Y)

Select Case pEditMode
    Case eCursorMode.cmNavi
        MDSF = ScaleFactor
    Case eCursorMode.cmSel
        SelectWaveSources X, Y, X, Y, Shift = 1, MDs ', Not (Shift = 1)
        Refr
        DrawSelRect X, Y, X, Y
        pctSpace.Refresh
    Case eCursorMode.cmEdit
        If Button = 1 Then
            BUD
            ChangeSelectedWS X, Y, Shift
            Refr
            pctSpace.Refresh
        ElseIf Button = 2 Then
            BUD
            ChangeSelectedSign
            Refr
            pctSpace.Refresh
        End If
'        If MDs >= 0 And MDs < nWS Then
'            BUD
'            If (Shift And 4) Then
'                WS(MDs).Strength = -WS(MDs).Strength
'            Else
'                ChangeWS WS(MDs), 0, 0, x, y, Shift
'            End If
'            Refr
'            pctSpace.Refresh
'        End If
    Case eCursorMode.cmCreate
        If Button = 1 Or Button = 2 Then
            BUD
            SetnWS nWS + 1
            MakeWS X, Y, X, Y, WS(nWS - 1), True, Shift
            If Button = 2 Then
                WS(nWS - 1).WaveLength = LastR
                SetStrength WS(nWS - 1).Strength, IIf((Shift And 4) = 4, -Abs(CurStrength), Abs(CurStrength)), False
            End If
            Refr
            pctSpace.Refresh
        End If
    Case eCursorMode.cmMove
        If Button = 1 Or Button = 2 Then
            BUD
        End If
        
    Case eCursorMode.cmColorize
        If Button = 1 Then
            If MDs >= 0 And MDs < nWS Then
                BUD
                WS(MDs).Color = CurColor
                Refr
                pctSpace.Refresh
            End If
        End If
End Select

UpdateStatus

Exit Sub
eh:
MsgError , Err.Number = errTMS
End Sub

Private Sub SelectWaveSources(ByVal x1 As Long, ByVal y1 As Long, _
                              ByVal x2 As Long, ByVal y2 As Long, _
                              ByVal MergeSelection As Boolean, _
                              ByVal NearestWS As Long)
Dim i As Long
Dim nInRect As Long
Dim OthersSelected As Boolean

x1 = SpaceToImageX(x1)
x2 = SpaceToImageX(x2)
y1 = SpaceToImageY(y1)
y2 = SpaceToImageY(y2)

If x1 > x2 Then
    SwapLng x1, x2
End If
If y1 > y2 Then
    SwapLng y1, y2
End If

For i = 0 To nWS - 1
    With WS(i).Pos
        If .X >= x1 And .X <= x2 And .Y >= y1 And .Y <= y2 Then
            nInRect = nInRect + 1
            WS(i).Selected = MergeSelection And Not MDAryWS(i).Selected _
                             Or Not MergeSelection And True
        Else
            WS(i).Selected = MergeSelection And MDAryWS(i).Selected _
                          Or Not MergeSelection And False
        End If
        If i <> NearestWS And MDAryWS(i).Selected Then
            OthersSelected = True
        End If
    End With
Next i

If nInRect = 0 And NearestWS >= 0 And NearestWS < nWS Then
    i = NearestWS
    If MergeSelection Then
        WS(i).Selected = Not MDAryWS(i).Selected
    Else
        WS(i).Selected = OthersSelected And True Or _
                         Not OthersSelected And Not MDAryWS(i).Selected
    End If
End If

End Sub

Private Sub SwapLng(ByRef l1 As Long, ByRef l2 As Long)
Dim l3 As Long
l3 = l1
l1 = l2
l2 = l3
End Sub

Private Sub pctSpaceMS_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim pEditMode As eCursorMode

UpdateStatus

On Error GoTo eh
If Button > 0 And Button <= 4 Then
    If Not MBState(Button) Then Exit Sub
End If

pEditMode = GetPEditMode(Shift)

Select Case MDEditMode
    Case eCursorMode.cmNavi
        If Button = 1 Then
            MoveView -(mdx - mmx), -(mdy - mmy)
            MoveView -(X - mdx), -(Y - mdy)
            Refr
            pctSpace.Refresh
        ElseIf Button = 2 Then
            RotateByMouse X, Y
            ReZoom MDSF, -RotAccum * 20, iMDx, iMDY, mdx, mdy
            Refr
            DrawStar mdx, mdy
            pctSpace.Refresh
        End If
    Case eCursorMode.cmSel
        If Button = 1 Then
            SelectWaveSources mdx, mdy, X, Y, Shift = 1, MDs
            Refr
            DrawSelRect mdx, mdy, X, Y
            pctSpace.Refresh
        End If
    Case eCursorMode.cmCreate
        If Button = 1 Then
            MakeWS mdx, mdy, X, Y, WS(nWS - 1), True, Shift
            LastR = WS(nWS - 1).WaveLength
            Refr
            pctSpace.Refresh
        End If
    Case eCursorMode.cmEdit
        If Button = 1 Then
            ChangeSelectedWS X, Y, Shift
            Refr
            pctSpace.Refresh
'            If MDs >= 0 And MDs < nWS Then
'                ChangeWS WS(MDs), _
'                         WS(MDs).Pos.x, WS(MDs).Pos.y, _
'                         x, y, 0
'                If (MDShift And 4) <> (Shift And 4) Then
'                    MDShift = MDShift And Not 4 Or Shift And 4
'                    WS(MDs).Strength = -WS(MDs).Strength
'                End If
'                Refr
'                pctSpace.Refresh
'            End If
        End If
    Case eCursorMode.cmMove
        If Button = 1 Then
            MoveSources mdx - mmx, mdy - mmy, MDs
            MoveSources X - mdx, Y - mdy, MDs
            Refr
            pctSpace.Refresh
        ElseIf Button = 2 Then
            RotateByMouse X, Y
            RotateSources MDAryWS, WS, nWS, RotAccum, _
                          iMDx, iMDY, MDs
            Refr
            DrawStar ImageToSpaceX(iMDx), ImageToSpaceY(iMDY)
            pctSpace.Refresh
        End If
End Select
mmx = X
mmy = Y
Exit Sub
eh:
MsgError , Err.Number = errTMS
End Sub

'uses mdx,mdy,rotaccum,mdaryws(),mmx,mmy,imdx,imdy
Private Sub RotateByMouse(ByVal X As Long, ByVal Y As Long)
Dim k As Double
Dim dAng As Double
k = ((mdx - X) * (mdx - X) + (mdy - Y) * (mdy - Y)) / (100 * 100)
dAng = -(Arg(X - mdx, Y - mdy) - Arg(mmx - mdx, mmy - mdy))
dAng = dAng - Int(dAng / Pi / 2) * Pi * 2
If dAng > Pi Then dAng = dAng - 2 * Pi
RotAccum = RotAccum + dAng * k
End Sub

Friend Sub MoveView(ByVal dx As Long, ByVal dy As Long)
dx = dx / ScaleFactor
dy = dy / ScaleFactor
SL = SL + dx
St = St + dy
SR = SR + dx
SB = SB + dy
End Sub

Friend Sub SetCursorMode(NewMode As eCursorMode)
'If CurEditMode <> NewMode Then
    CurEditMode = NewMode
    UpdateIndiMode
    If NewMode = cmColorize Then
        mnuVColors_Click
    End If
'End If
End Sub

Private Sub UpdateIndiMode()
Dim clr As Long
Dim TextClr As Long
Dim mnu As Menu
Select Case CurEditMode
    Case eCursorMode.cmCreate
        clr = vbRed
    Case eCursorMode.cmEdit
        clr = vbMagenta
    Case eCursorMode.cmMove
        clr = vbBlue
        TextClr = vbWhite
    Case eCursorMode.cmSel
        clr = vbYellow
    Case eCursorMode.cmNavi
        clr = vbWhite
    Case eCursorMode.cmColorize
        clr = RGB(255, 128, 64)
End Select
StatusText.ForeColor = TextClr
StatusBar.BackColor = clr

For Each mnu In mnuModes
    mnu.Checked = Val(mnu.Tag) = CurEditMode
Next
UpdateStatus
End Sub

'dzoom is space coords
'cx,cy are image coords
Friend Sub ReZoom(ByVal OrigZM As Double, _
                  ByVal dZoom As Double, _
                  ByVal icx As Long, ByVal icy As Long, _
                  ByVal CX As Long, ByVal CY As Long, _
                  Optional ByVal AbsoluteNewZoom As Boolean = False)
If AbsoluteNewZoom Then
    If dZoom <= 0 Then Err.Raise 111, "ReZoom", "ScaleFactor must be positive!"
    ScaleFactor = dZoom
Else
    ScaleFactor = Exp(dZoom / 200 * Log(4)) * OrigZM
End If
If ScaleFactor > 32 Then ScaleFactor = 32
If ScaleFactor < 1 / 100 Then ScaleFactor = 1 / 100
SL = icx - CX / ScaleFactor
St = icy - CY / ScaleFactor
SR = SL + pctSpace.ScaleWidth / ScaleFactor
SB = St + pctSpace.ScaleHeight / ScaleFactor
RecalcScale 'for sure
End Sub

Private Sub DeleteSelected(Optional ByVal StoreUndo As Boolean = False)
Dim NewWS() As typWaveSource
Dim i As Long, j As Long
Dim bb As Boolean
If nWS = 0 Then Exit Sub
ReDim NewWS(0 To nWS)
For i = 0 To nWS - 1
    If Not WS(i).Selected Then
        NewWS(j) = WS(i)
        j = j + 1&
    Else
        If Not bb Then If StoreUndo Then BUD
        bb = True
    End If
Next i
WS = NewWS
SetnWS j
End Sub

Private Sub pctSpaceMS_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button > 0 And Button <= 4 Then
    MBState(Button) = False
End If
Select Case MDEditMode
    Case eCursorMode.cmCreate
        If nWS > 0 Then
            CurStrength = WS(nWS - 1).Strength
        End If
    Case eCursorMode.cmEdit
        If MDs >= 0 And MDs < nWS Then
            CurStrength = WS(MDs).Strength
        End If
End Select
Refr
pctSpace.Refresh
UpdateStatus
End Sub

Private Sub pctSpace_Resize()
If ScaleFactor = 0 Then
    AutoScale
End If
SR = SL + pctSpace.Width / ScaleFactor
SB = St + pctSpace.Height / ScaleFactor
Refr
pctSpace.Refresh
End Sub

Private Sub DrawSelRect(ByVal x1 As Long, ByVal y1 As Long, _
                        ByVal x2 As Long, ByVal y2 As Long)
x1 = SpaceToImageX(x1)
x2 = SpaceToImageX(x2)
y1 = SpaceToImageY(y1)
y2 = SpaceToImageY(y2)

If x1 > x2 Then
    SwapLng x1, x2
End If
If y1 > y2 Then
    SwapLng y1, y2
End If

'draw selection rectangle
pctSpace.DrawMode = vbXorPen
pctSpace.DrawStyle = DrawStyleConstants.vbDot
pctSpace.Line (ImageToSpaceX(x1), ImageToSpaceY(y1))- _
              (ImageToSpaceX(x2), ImageToSpaceY(y2)), vbYellow, B
pctSpace.DrawStyle = DrawStyleConstants.vbSolid
End Sub

Friend Sub BUD()
If nUndos = 0 Then
    ReDim Undos(0 To 0)
End If
nUndos = nUndos + 1
ReDim Preserve Undos(0 To nUndos - 1)
Undos(nUndos - 1).WS = WS
Undos(nUndos - 1).nWS = nWS
End Sub

Friend Sub Undo()
If nUndos = 0 Then
    vtBeep
    Exit Sub
End If
WS = Undos(nUndos - 1).WS
nWS = Undos(nUndos - 1).nWS
nUndos = nUndos - 1
If nUndos = 0 Then
    Erase Undos
Else
    ReDim Preserve Undos(0 To nUndos - 1)
End If
End Sub

Friend Sub RotateSources(InWS() As typWaveSource, _
                         OutWS() As typWaveSource, _
                         ByVal nWS As Long, _
                         ByVal ByAngle As Double, _
                         ByVal CX As Double, _
                         ByVal CY As Double, _
                         Optional ByVal Nearest As Long = -1, _
                         Optional ByVal OnlySelected As Boolean = True, _
                         Optional ByVal RotateSingle As Boolean = True)
Dim i As Long
Dim si As Double, co As Double
Dim SelectedFound As Boolean
Dim tx As Long, ty As Long
si = Sin(ByAngle)
co = Cos(ByAngle)
OutWS = InWS
For i = 0 To nWS - 1
    If InWS(i).Selected Or Not OnlySelected Then
        SelectedFound = True
        GoSub RotateIt
    End If
Next i
If Not SelectedFound Then
    If RotateSingle Then
        If Nearest >= 0 And Nearest < nWS Then
            i = Nearest
            GoSub RotateIt
        End If
    Else
        For i = 0 To nWS - 1
            GoSub RotateIt
        Next i
    End If
End If

Exit Sub

RotateIt:
    With InWS(i).Pos
        tx = .X - CX
        ty = .Y - CY
        OutWS(i).Pos.X = tx * co + ty * si + CX
        OutWS(i).Pos.Y = -tx * si + ty * co + CY
    End With
Return

End Sub

Private Function Pi() As Double
Static pPi As Double
Static Foo As Boolean
If Foo Then
    Pi = pPi
Else
    pPi = Atn(1) * 4
    Foo = True
    Pi = pPi
End If
End Function


Private Function Arg(ByVal X As Double, ByVal Y As Double) As Double
Dim Rslt As Double
If X = 0# And Y = 0# Then
    Rslt = 0#
ElseIf Abs(X) > Abs(Y) Then
    Rslt = Atan2(X, Y)
Else
    Rslt = Pi * 0.5 - Atan2(Y, X)
End If
Arg = Rslt + Int((-Rslt + Pi) * 0.5 / Pi) * 2# * Pi
End Function

Private Function Atan2(ByVal X As Double, ByVal Y As Double) As Double
If X > 0# Then
    Atan2 = Atn(Y / X)
Else
    Atan2 = Pi + Atn(Y / X)
End If
End Function

Private Function MBPressed() As Boolean
MBPressed = MBState(1) Or MBState(2) Or MBState(4)
End Function

Private Sub UpdateStatus()
Dim Shift As Integer
Dim tmp As String
Dim pEditMode As eCursorMode
Shift = GetShiftState \ &H100
pEditMode = GetPEditMode(Shift)
If MBPressed Then
    pEditMode = MDEditMode
End If
Select Case pEditMode
    Case eCursorMode.cmCreate
        tmp = GRSF(2586) '"Creating. Use LMB to create. Use RMB instead to preserve radius."
    Case eCursorMode.cmEdit
        tmp = GRSF(2587) '"Editing. Use LMB to change radius. Use Alt+RMB to change sign."
    Case eCursorMode.cmMove
        If MBState(1) Then
            tmp = GRSF(2588) '"Drag pointer to drag selected sources or the nearest one."
        ElseIf MBState(2) Then
            tmp = GRSF(2589) '"Rotate the pointer around the green square."
        Else
            tmp = GRSF(2590) '"Moving. Use LMB to drag. Use RMB to zoom."
        End If
    Case eCursorMode.cmSel
        tmp = GRSF(2591) '"Selecting. Click to select the nearest. Drag to select multiple."
    Case eCursorMode.cmNavi
        If MBState(1) Then
            tmp = GRSF(2592) '"Drag pointer do drag the view."
        ElseIf MBState(2) Then
            tmp = GRSF(2593) '"Rotate pointer around green square to zoom in/out."
        Else
            tmp = GRSF(2594) '"Navigating. LMB to move. RMB to zoom."
        End If
    Case eCursorMode.cmColorize
        tmp = GRSF(2595) '"Colorizing. Click the source to colorize it with current color."
        
End Select
StatusText.Caption = tmp
End Sub

Public Sub UpdateCurColor()
clrCurColor.DisableNextChange = True
clrCurColor.Color = RGB(CurColor.rgbRed, _
                        CurColor.rgbGreen, _
                        CurColor.rgbBlue)
End Sub

'uses mdx,mdy,mds,mdaryws
Public Sub ChangeSelectedWS(ByVal X As Long, _
                            ByVal Y As Long, _
                            ByVal Shift As Integer)
Dim nSel As Long
Dim SelectedFound As Boolean
Dim i As Long
Dim OldR As Double
Dim UseStrength As Boolean
Dim WSi As Long
Dim Sign As Long
Dim r As Double
Dim sx As Long, sy As Long
If nWS = 0 Then Exit Sub
WSi = NearestWS(mdx, mdy, True)
If WSi < 0 Then
    WSi = MDs
    If WSi < 0 Then
        Exit Sub
    End If
Else
    SelectedFound = True
End If
UseStrength = CBool(DisplayMode And wsdmStrength)
If SelectedFound Then
    If UseStrength Then
        OldR = Abs(MDAryWS(WSi).Strength)
    Else
        OldR = MDAryWS(WSi).WaveLength
    End If
End If
Sign = IIf(CBool(Shift And 4), -1, 1)
sx = ImageToSpaceX(MDAryWS(WSi).Pos.X)
sy = ImageToSpaceY(MDAryWS(WSi).Pos.Y)
r = Sqr((X - sx) * (X - sx) + (Y - sy) * (Y - sy)) / ScaleFactor
If SelectedFound Then
    For i = 0 To nWS - 1
        If MDAryWS(i).Selected Then
            GoSub ChangeIt
        End If
    Next i
Else
    i = WSi
    GoSub ChangeIt
End If
Exit Sub
ChangeIt:
    WS(i).Strength = MDAryWS(i).Strength
    WS(i).WaveLength = MDAryWS(i).WaveLength
    If UseStrength Then
        If OldR <> 0 Then
            WS(i).Strength = r / OldR * MDAryWS(i).Strength
        Else
            WS(i).Strength = r
        End If
        If Abs(WS(i).Strength) > 16 Then
            WS(i).Strength = 16 * Sgn(WS(i).Strength)
        End If
    Else
        If OldR <> 0 Then
            WS(i).WaveLength = r / OldR * MDAryWS(i).WaveLength
        Else
            WS(i).WaveLength = r
        End If
        If WS(i).WaveLength > 1000000 Then
            WS(i).WaveLength = 1000000
        End If
    End If
    WS(i).Strength = WS(i).Strength * Sign
Return
End Sub

Private Sub ChangeSelectedSign()
Dim i As Long
Dim SelectedFound As Boolean
For i = 0 To nWS - 1
    If WS(i).Selected Then
        SelectedFound = True
        WS(i).Strength = -WS(i).Strength
    End If
Next i
If Not SelectedFound Then
    If MDs >= 0 And MDs < nWS Then
        WS(MDs).Strength = -WS(MDs).Strength
    End If
End If
End Sub

Private Sub StatusBar_ClickVT()
Dim i As Long
For i = 0 To mnuModes.UBound
    If mnuModes(i).Tag = CurEditMode Then
        Exit For
    End If
Next i
If i = mnuModes.UBound + 1 Then
    i = 0
Else
    i = i + 1
End If
i = i Mod (mnuModes.UBound + 1)
mnuModes_Click (i)
End Sub

Private Sub StatusBar_DblClick()
StatusBar_MouseDown 1, 0, 0, 0
End Sub

Private Sub StatusBar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
StatusText_MouseDown Button, Shift, X, Y
End Sub

Private Sub StatusBar_Resize()
clrCurColor.Move 0, 0, StatusBar.ScaleHeight, StatusBar.ScaleHeight
btnToggleView.Move StatusBar.ScaleWidth - btnToggleView.Width, _
                   0, _
                   btnToggleView.Width, _
                   StatusBar.ScaleHeight
StatusText.Move clrCurColor.Width, _
                (StatusBar.ScaleHeight - StatusText.Height) \ 2, _
                StatusBar.ScaleWidth - clrCurColor.Width - btnToggleView.Width
End Sub

Private Sub StatusText_ClickVT()
StatusBar_ClickVT
End Sub

Private Sub OK()
On Error GoTo eh
SaveToReg
SaveSettings
Me.Tag = ""
Me.Hide
Exit Sub
eh:
MsgError
End Sub

Friend Sub SaveSettings()
dbSaveSettingEx "Draw\Waves", "FieldPower", mFieldPower
dbSaveSettingEx "Draw\Waves", "FalldownFactor", mFallDownFactor
dbSaveSettingEx "Draw\Waves", "WavesDrawAbsolute", mnuDoneWavesOptsAbsolute.Checked
dbSaveSettingEx "Draw\Waves", "DialogDisplayMode", DisplayMode
End Sub

Friend Sub LoadSettings()
mFieldPower = dbGetSettingEx("Draw\Waves", "FieldPower", vbDouble, 2#)
mFallDownFactor = dbGetSettingEx("Draw\Waves", "FalldownFactor", vbDouble, 0#)
mnuDoneWavesOptsAbsolute.Checked = dbGetSettingEx("Draw\Waves", "WavesDrawAbsolute", vbBoolean, True)
DisplayMode = dbGetSettingEx("Draw\Waves", "DialogDisplayMode", vbLong, dbWSDM.wsdmWaveLength)
End Sub


Private Sub CancelButton_Click()
On Error GoTo eh
UserCancel
Exit Sub
eh:
MsgError
End Sub

Private Sub UserCancel()
Dim Answ As VbMsgBoxResult
'2596= "Remember the configuration?"
Answ = dbMsgBox(2596, vbQuestion Or vbYesNoCancel)
If Answ = vbYes Then
    SaveToReg
    SaveSettings
End If
If Answ = vbCancel Then
    Err.Raise dbCWS
End If
Me.Tag = "C"
Me.Hide
End Sub

Friend Sub SaveWSToReg(ByRef WS() As typWaveSource, _
                       ByVal nWS As Long, _
                       ByRef Section As String, _
                       ByRef Parameter As String)
Dim Bytes() As Byte
Dim cb As Long
If nWS <= 0 Then
    dbDeleteSetting Section, Parameter
    Exit Sub
End If
cb = nWS * LenB(WS(0))
ReDim Bytes(0 To cb - 1)
CopyMemory Bytes(0), WS(0), cb
dbSaveSettingBin Section, Parameter, Bytes
End Sub

Friend Sub LoadWSFromReg(ByRef WS() As typWaveSource, _
                         ByRef nWS As Long, _
                         ByRef Section As String, _
                         ByRef Parameter As String)
Dim Bytes() As Byte
Dim cb As Long
Dim tst As typWaveSource
cb = dbGetSettingBin(Section, Parameter, Bytes)
If cb Mod LenB(tst) <> 0 Then Exit Sub
nWS = cb \ LenB(tst)
If nWS > 0 Then
    ReDim WS(0 To nWS - 1)
    CopyMemory WS(0), Bytes(0), cb
Else
    Erase WS
End If
End Sub

Public Sub SaveToReg()
SaveWSToReg WS, nWS, "Waves", "LastConfig"
End Sub

Public Sub LoadFromReg()
LoadWSFromReg WS, nWS, "Waves", "LastConfig"
End Sub

Friend Function GetSources(ByRef aWS() As typWaveSource) As Long
aWS = WS
GetSources = nWS
End Function

Private Sub StatusText_DblClick()
StatusText_MouseDown 1, 0, 0, 0
End Sub

Private Sub StatusText_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
    PopupMenu mnuMode, vbPopupMenuRightButton
ElseIf Button = 1 Then
    StatusText_ClickVT
End If
End Sub
