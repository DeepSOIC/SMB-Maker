VERSION 5.00
Begin VB.UserControl ctlTaggedText 
   Appearance      =   0  'Flat
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ClipControls    =   0   'False
   ControlContainer=   -1  'True
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   204
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HasDC           =   0   'False
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   Begin VB.PictureBox pctBox 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      Height          =   3195
      Left            =   195
      ScaleHeight     =   213
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   186
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   240
      Width           =   2790
   End
   Begin VB.VScrollBar VScroll 
      Height          =   1950
      Left            =   4275
      TabIndex        =   0
      Top             =   255
      Width           =   255
   End
   Begin VB.Menu mnuPP 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu mnuBack 
         Caption         =   "Back	RMB+LMB"
      End
      Begin VB.Menu mnuFW 
         Caption         =   "Forward	LMB+RMB"
      End
   End
End
Attribute VB_Name = "ctlTaggedText"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Declare Function DrawTextEx Lib "user32" Alias "DrawTextExA" (ByVal hDC As Long, ByVal lpsz As String, ByVal n As Long, lpRect As RECT, ByVal un As TextDrawMode, lpDrawTextParams As DRAWTEXTPARAMS) As Long
Public DisableHist As Boolean
Dim pNo3D As Boolean

Dim WithEvents pctBoxMS As clsAntiDblClick
Attribute pctBoxMS.VB_VarHelpID = -1

Private Enum TextDrawMode
    DT_BOTTOM = &H8
    DT_CALCRECT = &H400
    DT_CENTER = &H1
    DT_CHARSTREAM = 4
    DT_DISPFILE = 6
    DT_EXPANDTABS = &H40
    DT_EXTERNALLEADING = &H200
    DT_INTERNAL = &H1000
    DT_LEFT = &H0
    DT_METAFILE = 5
    DT_NOCLIP = &H100
    DT_NOPREFIX = &H800
    DT_RIGHT = &H2
    DT_SINGLELINE = &H20
    DT_TABSTOP = &H80
    DT_TOP = &H0
    DT_VCENTER = &H4
    DT_WORDBREAK = &H10
End Enum

Private Type DRAWTEXTPARAMS
    cbSize As Long
    iTabLength As Long
    iLeftMargin As Long
    iRightMargin As Long
    uiLengthDrawn As Long
End Type

Public Enum TextBlockType
    dbTBTText = 0
    dbTBTLink = 1
End Enum

Private Type TextBlock
    Type As TextBlockType
    Text As String
    LinkTo As Long
    Centered As Boolean
    Size As Integer
    yStart As Long
    yEnd As Long
End Type

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Dim HText As String
Dim CurResID As Long
Dim Blocks() As TextBlock

Dim ClHeight As Long
Dim VScrollEnabled As Boolean

Dim NeedPosRebuild As Boolean
Dim History() As Long
Dim HistSize As Long
Dim HistIndex As Long

Dim MBPressed(1 To 4) As Boolean
Dim CancelPopUp As Boolean

Public Event Resize()

Private Sub LinkClick(ByVal ToResID As Integer)
If ToResID <> 0 Then
VScroll.Value = 0
SetResID ToResID
End If
End Sub

Private Sub mnuBack_Click()
HistBackward
End Sub

Private Sub mnuFW_Click()
HistForward
End Sub

Private Sub pctBox_DblClick()
pctBoxMS.DblClick
End Sub

Private Sub pctBox_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
pctBoxMS.MouseDown Button, Shift, X, Y
End Sub

Private Sub pctBox_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
pctBoxMS.MouseMove Button, Shift, X, Y
End Sub

Private Sub pctBox_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
pctBoxMS.MouseUp Button, Shift, X, Y
End Sub

Private Sub pctBoxMS_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim BN As Long
On Error Resume Next
CancelPopUp = False
MBPressed(Button) = True
If (Button = 1 And MBPressed(2)) Then
    HistBackward
    CancelPopUp = True
    Exit Sub
ElseIf Button = 2 And MBPressed(1) Then
    HistForward
    CancelPopUp = True
    Exit Sub
End If
On Error GoTo 0
If Button = 1 Then
    BN = GetBlockByY(Blocks, Y)
    If BN >= 0 Then
        If Blocks(BN).Type = dbTBTLink Then
            LinkClick Blocks(BN).LinkTo
        End If
    End If
End If
End Sub

Private Function GetBlockByY(ByRef Blocks() As TextBlock, ByVal Y As Long) As Integer
Dim UB As Long
Dim LB As Long
Dim MidB As Long
Dim InBl As Integer
    
UB = -1
On Error Resume Next
If AryDims(AryPtr(Blocks)) = 1 Then
    UB = UBound(Blocks)
End If
If UB = -1 Then
    GetBlockByY = -1
    Exit Function
End If
LB = 0
Do
    MidB = (UB + LB) \ 2&
    InBl = InBlock(Blocks, MidB, Y)
    If InBl = 0 Then
        GetBlockByY = MidB
        Exit Function
    ElseIf InBl = -1 Then
        UB = MidB - 1
    ElseIf InBl = 1 Then
        LB = MidB + 1
    End If
    If UB = LB Then
        If InBlock(Blocks, UB, Y) = 0 Then
            GetBlockByY = UB
        Else
            GetBlockByY = -1
        End If
        Exit Function
    End If
Loop
End Function

Private Function InBlock(ByRef Blocks() As TextBlock, _
                        ByVal Index As Long, _
                        ByVal Y As Long) As Long
If Y < Blocks(Index).yStart Then
    InBlock = -1
ElseIf Y >= Blocks(Index).yEnd Then
    InBlock = 1
Else
    InBlock = 0
End If
End Function

Private Sub pctBoxMS_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
MBPressed(Button) = False
If Button = 2 Then
    If Not CancelPopUp Then
        ValidateFB
        If Not DisableHist Then
            PopupMenu mnuPP, vbPopupMenuRightButton
        End If
    End If
End If
End Sub

Private Sub pctBox_Resize()
Static OldW As Long
UpdateScroll
ApplyScroll
If OldW <> pctBox.Width Then
    NeedPosRebuild = True
    pctBox.Refresh
End If
End Sub

Private Sub pctBox_Paint()
On Error Resume Next
Dim yf As Long, yh As Long
Dim Pct As IPictureDisp
Dim PctW As Long, PctH As Long
Set Pct = LoadResPicture(107, vbResBitmap)
PctH = pctBox.ScaleY(Pct.Height, vbHimetric, vbPixels)
PctW = pctBox.ScaleX(Pct.Width, vbHimetric, vbPixels)
yf = -pctBox.Top * PctH \ pctBox.ScaleHeight
yh = -(-UserControl.ScaleHeight * PctH / pctBox.ScaleHeight) + 2
If yh > PctH Then yh = PctH
If Not pNo3D Then
    pctBox.PaintPicture Pct, 0, -Int(-(yf) * pctBox.ScaleHeight / PctH), pctBox.ScaleWidth, -Int(-yh * pctBox.ScaleHeight / PctH), 0, yf, PctW, yh
'Else
'    pctBox.Line (0, 0)-(pctBox.ScaleWidth, pctBox.ScaleHeight), pctBox.backcolor, BF
End If
On Error GoTo 0
PaintText
End Sub

Public Function DrawText(ByRef St As String, _
                         ByVal X As Long, ByVal Y As Long, _
                         ByVal w As Long, _
                         ByVal lngColor As Long, _
                         ByVal FontSize As Integer, _
                         ByVal CenterText As Boolean, _
                         Optional ByVal ReturnValue As Boolean = True, _
                         Optional ByVal LinkStyle As Boolean = False) As Long
Dim DTP As DRAWTEXTPARAMS
Dim Rct As RECT
Dim FU As Boolean
Static WasLinkStyle As Boolean
Static LastFontSize As Integer
Static LastColor As Long
Static LastFU As Boolean

If Len(St) = 0 Then
    If ReturnValue Then DrawText = Y
    Exit Function
End If

If St = vbNewLine Then
    If ReturnValue Then DrawText = Y + pctBox.ScaleY(FontSize, vbPoints, vbPixels)
    Exit Function
End If

FU = LinkStyle
If LinkStyle Then
    lngColor = IIf(lngColor = 0, vbBlue, lngColor)
End If
If LastColor <> lngColor Then
    pctBox.ForeColor = lngColor
    LastColor = lngColor
End If
If LastFontSize <> FontSize Then
    pctBox.FontSize = FontSize
    LastFontSize = FontSize
End If
If LastFU <> FU Then
    pctBox.FontUnderline = FU
    LastFU = FU
End If
Rct.Left = X
Rct.Top = Y
Rct.Right = X + w
Rct.Bottom = Y + 2
DTP.cbSize = LenB(DTP)
DTP.iLeftMargin = 1
DTP.iRightMargin = 1
DTP.iTabLength = 4
'DrawTextEx UserControl.hDC, St, Len(St), rct, DT_EXPANDTABS Or DT_CALCRECT Or DT_NOCLIP Or DT_NOPREFIX Or DT_TABSTOP Or DT_WORDBREAK, DTP
If ReturnValue Then
    DrawTextEx pctBox.hDC, St, Len(St), _
               Rct, _
               DT_CALCRECT Or _
               DT_EXPANDTABS Or _
               DT_NOPREFIX Or _
               DT_TABSTOP Or _
               DT_WORDBREAK Or _
               (DT_CENTER And CenterText), _
               DTP
    Rct.Left = X
    Rct.Top = Y
    Rct.Right = X + w
End If

DrawTextEx pctBox.hDC, St, Len(St), _
           Rct, _
           DT_EXPANDTABS Or _
           DT_NOPREFIX Or _
           DT_NOCLIP Or _
           DT_TABSTOP Or _
           DT_WORDBREAK Or _
           (DT_CENTER And CenterText), _
           DTP
           
If ReturnValue Then
    DrawText = Rct.Bottom
End If
End Function

Private Sub ProcessText()
Dim i As Long, j As Long, k As Long, BN As Long
Dim PosOpened As Long
Dim bOpened As Boolean
Dim TagText As String
Dim tTagText As String, Prefix As String, Suffix As String
Dim TTP As String
Dim m As String * 1
Dim CenterAlign As Boolean, nCa As Boolean
Dim FontSize As Long, nFS As Long
Dim PrTag As Boolean
Dim RmLf As Boolean
Dim PreserveText As Boolean
Dim Y As Long
Dim sAry() As String
Dim f_ As Long, t_ As Long
Dim nLF As Long

HText = Replace(HText, Chr$(13) + Chr$(13) + Chr$(10), "")

j = 1
bOpened = False
TagText = ""
FontSize = 10 ': nFS = 10
ReDim Blocks(0 To 0)
BN = 0
Blocks(0).Text = Space$(Len(HText))
Blocks(0).Size = FontSize
Blocks(0).Centered = CenterAlign

For i = 1 To Len(HText)
    m = Mid$(HText, i, 1)
    If m = "<" Then
        If bOpened Then
            'Err.Raise 115, "PaintHelp", "Somwhere > character is needed."
            TagText = "<" + TagText
            GoSub OutputTagText
            TagText = ""
        End If
        bOpened = True
        PosOpened = i
    ElseIf m = ">" And bOpened Then
        'If Not bOpened Then
        '    Err.Raise 115, "PaintHelp", "Somwhere < character is needed."
        'End If
        bOpened = False
        
        'process the tag
        'tTagText = Trim$(TagText)
        k = InStr(1, TagText, "=")
        PrTag = False
        RmLf = False
        If k = 0 Then
            TagText = "<" + TagText + ">"
            PrTag = True 'Print its contents
        Else
            Prefix = Trim$(Mid$(TagText, 1, k - 1))
            Suffix = Trim$(Right$(TagText, Len(TagText) - k))
            Select Case UCase$(Prefix)
                Case "CENTER"
                    CenterAlign = CBool(Val(Suffix))
                    RmLf = True
                Case "FSIZE"
                    FontSize = Val(Suffix)
                    RmLf = True
                Case "CHAR"
                    TagText = Chr$(Val(Suffix))
                    PrTag = True
                Case "RESID"
                    TagText = GRSF(Val(Suffix))
                    PrTag = True
                Case "COMMENT"
                    TagText = ""
                    PrTag = True
                Case "FRAGMENT"
                    HText = Left$(HText, PosOpened - 1) + GRSF(Val(Suffix)) + Mid$(HText, i + 1, Len(HText) - i)
                    ProcessText
                    Exit Sub
                Case "FRAGMENTS" '<fragments=Start_ID=Stop_ID[,NumberOfLineFeeds=2]>
                    k = InStr(1, Suffix, "-")
                    TagText = "<" + TagText + ">"
                    PrTag = True
                    If k > 0 Then
                        nLF = 2
                        nFS = InStr(k, Suffix, ",")
                        If nFS <= 0 Then
                            nFS = Len(Suffix) + 1
                        Else
                            nLF = Val(Mid$(Suffix, nFS + 1, Len(Suffix) - nFS))
                        End If
                        f_ = Val(Left$(Suffix, k - 1))
                        t_ = Val(Mid$(Suffix, k + 1, nFS - 1 - k))
                        If f_ <= t_ Then
                            ReDim sAry(0 To t_ - f_)
                            nFS = 0
                            For k = f_ To t_
                                On Error Resume Next
                                    Suffix = vbNullString
                                    Suffix = GRSF(k, RaiseErrors:=True)
                                On Error GoTo 0
                                If Len(Suffix) > 0 Then
                                    sAry(nFS) = GRSF(k, RaiseErrors:=True)
                                    nFS = nFS + 1
                                End If
                            Next k
                            If nFS > 0 Then
                                ReDim Preserve sAry(0 To nFS - 1)
                                Suffix = ""
                                For k = 1 To nLF
                                    Suffix = Suffix + vbCrLf
                                Next k
                            End If
                            HText = Left$(HText, PosOpened - 1) + Join(sAry, Suffix) + Mid$(HText, i + 1, Len(HText) - i)
                            ProcessText
                            Exit Sub
                        End If
                    End If
                    
                Case "LINK"
                    On Error Resume Next
                    k = InStr(2, Suffix, ",")
                    Err.Clear
                    t_ = CLng(Mid$(Suffix, k + 1))
                    If Err.Number <> 0 Or k < 0 Then
                        On Error GoTo 0
                        PrTag = True
                        TagText = "<" + TagText + ">"
                    Else
                        On Error GoTo 0
                        RmLf = True
                        GoSub BeginNewBlock
                        Blocks(BN).Type = dbTBTLink
                        Blocks(BN).Text = Left$(Suffix, k - 1)
                        Blocks(BN).LinkTo = t_
                        PreserveText = True
                    End If
                
                Case Else
                    TagText = "<" + TagText + ">"
                    PrTag = True 'Print its contents
            End Select
        End If
        If PrTag Then
            GoSub OutputTagText
        Else
            'Y = DrawText(Mid$(TTP, 1, j - 1), 0, Y, fHelp.ScaleWidth, 0, FontSize, CenterAlign)
            GoSub BeginNewBlock
'            FontSize = nFS
'            CenterAlign = nCa
        End If
        TagText = ""
    Else
        If bOpened Then
            TagText = TagText + m
        Else
            Mid$(Blocks(BN).Text, j, 1) = m
            j = j + 1
        End If
    End If
Next i
If j > 1 Then
    'Y = DrawText(Mid$(TTP, 1, j - 1), 0, Y, fHelp.ScaleWidth, 0, FontSize, CenterAlign)
    Blocks(BN).Text = Mid$(Blocks(BN).Text, 1, j - 1)
End If

NeedPosRebuild = True
'If Y <> 0 Then
'    Height = Y + 4 + 2
'End If
Exit Sub

OutputTagText:
    Mid$(Blocks(BN).Text, j, Len(TagText)) = TagText
    j = j + Len(TagText)
Return

BeginNewBlock:
    If Not PreserveText Then
        Blocks(BN).Text = Left$(Blocks(BN).Text, j - 1)
        If RmLf And Len(Blocks(BN).Text) >= 2 Then
            If Right$(Blocks(BN).Text, 2) = vbCrLf Then
                Blocks(BN).Text = Left$(Blocks(BN).Text, j - 3)
            End If
        End If
        If RmLf And Len(Blocks(BN).Text) >= 2 Then
            If Left$(Blocks(BN).Text, 2) = vbCrLf Then
                Blocks(BN).Text = Mid$(Blocks(BN).Text, 3)
            End If
        End If
    End If
    PreserveText = False
    BN = BN + 1
    ReDim Preserve Blocks(0 To BN)
    Blocks(BN).Centered = CenterAlign
    Blocks(BN).Size = FontSize
    Blocks(BN).Text = Space$(Len(HText) - i)
    j = 1
Return

End Sub
'
'Public Sub LoadLinkButton(ByVal Index As Long)
'Dim i As Long
'Dim n As Long
'For i = 0 To Index
'    On Error Resume Next
'    Load btnLink(i)
'Next i
'End Sub

'Private Sub AddBlock(ByRef Blocks() As TextBlock, ByRef Text As String, ByVal CenterAligned As Boolean, ByVal FontSize As Integer)
'Dim i As Long
'i = UBound(Blocks)
'ReDim Preserve Blocks(0 To i + 1)
'Blocks(i + 1).Text = Text
'Blocks(i + 1).Centered = CenterAligned
'End Sub

Private Sub PaintText()
Dim i As Long, UB As Long
Dim Y As Long
Dim iLink As Long
UB = -1
On Error Resume Next
If AryDims(AryPtr(Blocks)) = 1 Then
    UB = UBound(Blocks)
End If
On Error GoTo 0
If UB > -1 Then
    Y = 0
    iLink = 0
    For i = 0 To UB
        If NeedPosRebuild Then
            Blocks(i).yStart = Y
'            Select Case Blocks(i).Type
'                Case TextBlockType.dbTBTText
                    Y = DrawText(Blocks(i).Text, 0, Y, pctBox.ScaleWidth, 0, Blocks(i).Size, Blocks(i).Centered, , Blocks(i).Type = dbTBTLink)
'                Case TextBlockType.dbTBTLink
'                    LoadLinkButton iLink
'                    btnLink(iLink).Move 4, Y, pctBox.ScaleWidth - 8
'                    btnLink(iLink).Caption = Blocks(i).Text
'                    btnLink(iLink).Tag = CStr(Blocks(i).LinkTo)
'                    btnLink(iLink).Visible = True
'                    Y = Y + btnLink(iLink).Height
'                    iLink = iLink + 1&
'            End Select
            Blocks(i).yEnd = Y
        Else
            If Blocks(i).yEnd > -pctBox.Top And Blocks(i).yStart < -pctBox.Top + ScaleHeight Then
'                If Blocks(i).Type = dbTBTText Then
                    DrawText Blocks(i).Text, 0, Blocks(i).yStart, pctBox.ScaleWidth, 0, Blocks(i).Size, Blocks(i).Centered, False, Blocks(i).Type = dbTBTLink
'                End If
            End If
        End If
    Next i
    If NeedPosRebuild Then
        'HideLinks iLink
        pctBox.Height = (Y + 4 + 2)
        NeedPosRebuild = False
    End If
End If
End Sub

'Public Sub HideLinks(ByVal cnt As Long)
'Dim i As Long
'For i = cnt To btnLink.UBound
'    btnLink(i).Visible = False
'Next i
'End Sub
'
Public Sub SetResID(ByVal ResID As Integer, _
                    Optional ByVal PutHistory As Boolean = True)
If PutHistory Then
    bh ResID
End If
HText = GRSF(ResID)
ProcessText
pctBox.Refresh
CurResID = ResID
End Sub

Public Sub SetText(ByRef Text As String)
DisableHist = True
HText = Text
ProcessText
pctBox.Refresh
CurResID = 0
End Sub

Private Sub ApplyScroll()
If VScrollEnabled Then
    pctBox.Move 0, -VScroll.Value, ScaleWidth - VScroll.Width
Else
    pctBox.Move 0, 0, ScaleWidth - VScroll.Width
    VScroll.Value = 0
End If
End Sub

Private Sub UpdateScroll()
VScrollEnabled = UserControl.ScaleHeight < pctBox.Height
VScroll.Enabled = VScrollEnabled
If VScrollEnabled Then
    VScroll.Max = pctBox.Height - UserControl.ScaleHeight
    VScroll.LargeChange = Round(ScaleHeight * 0.8 + 0.5)
Else
    VScroll.Max = 0
End If
End Sub

Private Sub UserControl_Initialize()
Set pctBoxMS = New clsAntiDblClick
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
With PropBag
    DisableHist = .ReadProperty("DisableHist", False)
    No3D = .ReadProperty("No3D", False)
    ForeColor = .ReadProperty("ForeColor", 0&)
    BackColor = .ReadProperty("BackColor", RGB(200, 200, 200))
End With
End Sub

Private Sub UserControl_Resize()
UpdateScroll
VScroll.Move ScaleWidth - VScroll.Width, 0, VScroll.Width, ScaleHeight
ApplyScroll
RaiseEvent Resize
Refresh
End Sub

Public Sub Refresh()
UserControl.Refresh
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
With PropBag
    .WriteProperty "DisableHist", DisableHist, False
    .WriteProperty "No3D", No3D, False
    .WriteProperty "ForeColor", ForeColor, 0&
    .WriteProperty "BackColor", BackColor, RGB(200, 200, 200)
End With
End Sub

Private Sub VScroll_Change()
VScroll_Scroll
End Sub

Private Sub VScroll_Scroll()
ApplyScroll
End Sub

Public Sub Scroll(ByVal Value As Long)
Dim NV As Long
If Not VScroll.Enabled Then Exit Sub
NV = VScroll.Value + Value
If NV > VScroll.Max Then NV = VScroll.Max
If NV < VScroll.Min Then NV = VScroll.Min
VScroll.Value = NV
End Sub

Public Sub bh(ByVal ResID As Long)
If DisableHist Then Exit Sub
If ResID = 2395 Or ResID = 0 Then Exit Sub
If HistSize = 0 Then
    ReDim History(0 To 0)
    HistIndex = -1
End If
HistIndex = HistIndex + 1
HistSize = HistIndex + 1
ReDim Preserve History(0 To HistSize - 1)
History(HistIndex) = ResID
End Sub

Public Sub HistForward()
If HistSize = 0 Then Exit Sub
If DisableHist Then Exit Sub
HistIndex = HistIndex + 1
If HistIndex > HistSize - 1 Then
    HistIndex = HistSize - 1
    Exit Sub
End If
If History(HistIndex) > 0 Then
    SetResID History(HistIndex), False
End If
End Sub

Public Sub HistBackward()
If HistSize = 0 Then Exit Sub
If DisableHist Then Exit Sub
HistIndex = HistIndex - 1
If HistIndex < 0 Then
    HistIndex = 0
    Exit Sub
End If
If History(HistIndex) > 0 Then
    SetResID History(HistIndex), False
End If
End Sub

Public Sub ValidateFB()
mnuFW.Enabled = HistIndex < HistSize - 1
mnuBack.Enabled = HistIndex > 0
End Sub

Public Property Get BackColor() As OLE_COLOR
BackColor = pctBox.BackColor
End Property

Public Property Let BackColor(ByVal vNewValue As OLE_COLOR)
pctBox.BackColor = vNewValue
pctBox.Refresh
UserControl.BackColor = vNewValue
End Property

Public Property Get ForeColor() As OLE_COLOR
ForeColor = pctBox.ForeColor
End Property

Public Property Let ForeColor(ByVal vNewValue As OLE_COLOR)
pctBox.ForeColor = vNewValue
pctBox.Refresh
End Property


Public Property Get No3D() As Boolean
No3D = pNo3D
End Property

Public Property Let No3D(ByVal vNewValue As Boolean)
Dim NeedRefr As Boolean
NeedRefr = pNo3D <> vNewValue
pNo3D = vNewValue
If NeedRefr Then pctBox.Refresh
End Property

