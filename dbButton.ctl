VERSION 5.00
Begin VB.UserControl dbButton 
   Alignable       =   -1  'True
   ClientHeight    =   465
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   960
   ClipControls    =   0   'False
   DefaultCancel   =   -1  'True
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
   KeyPreview      =   -1  'True
   ScaleHeight     =   31
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   64
   ToolboxBitmap   =   "dbButton.ctx":0000
   Windowless      =   -1  'True
End
Attribute VB_Name = "dbButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer

Private Declare Function pDrawTextEx Lib "user32" Alias "DrawTextExA" (ByVal hDC As Long, ByVal lpsz As String, ByVal n As Long, lpRect As pRECT, ByVal un As pTextDrawMode, lppDRAWTEXTPARAMS As pDRAWTEXTPARAMS) As Long
Private Enum pTextDrawMode
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
Private Type pDRAWTEXTPARAMS
    cbSize As Long
    iTabLength As Long
    iLeftMargin As Long
    iRightMargin As Long
    uiLengthDrawn As Long
End Type
Private Type pRECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type



Dim BV As Boolean, strCaption As String, tmpMM As Boolean, En As Boolean, NPP As Boolean
Dim HasFocus As Boolean
Public ResID As Long
Attribute ResID.VB_VarDescription = "Returns/sets the resource ID containing it's caption. If =0, the caption property will be used."
Public dbTag1 As Variant, dbTag2 As Variant, dbTag3 As Variant
Dim pPictureResID As Variant, pPictureResType As LoadResConstants
Dim Def As Boolean, Cnc As Boolean
Dim MyOrigCaption As String
Dim pNoTags As Boolean
Dim pAutoRedraw As Boolean
Private MouseState(1 To 4) As Boolean, CancelPress As Boolean
Public pVisible As Boolean
Public Event Click()
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event KeyPress(KeyAscii As Integer)
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event KeyUp(KeyCode As Integer, Shift As Integer)

Sub FreshButton(Value As Boolean)
Dim px As Integer, py As Integer
Dim tmpPicture As IPictureDisp
Dim tmpCaption As String
If Not pVisible Then Exit Sub
If ResID <> 0 Then
    On Error GoTo eh
    strCaption = LoadResString(ResID)
    On Error GoTo 0
Else
    strCaption = MyOrigCaption
End If

If pNoTags Then
    tmpCaption = strCaption
Else
    tmpCaption = strCaption
End If

On Error GoTo eh2

'UserControl.Cls
UserControl.Line (0, 0)- _
                 (UserControl.ScaleWidth - 1, UserControl.ScaleHeight - 1), _
                 UserControl.BackColor, BF
On Error Resume Next
If Not (NPP) Then
    If En Then
        If Value Then
            UserControl.PaintPicture LoadResPicture("BTN_DOWN", vbResBitmap), 2, 2, UserControl.ScaleWidth - 4, UserControl.ScaleHeight - 4
            'down
        Else
            If HasFocus Then
                'has focus
                UserControl.PaintPicture LoadResPicture("BTN_FOCUS", vbResBitmap), 2 + Abs(Value), 2 + Abs(Value), UserControl.ScaleWidth - 4, UserControl.ScaleHeight - 4
            ElseIf Def Then
                'default button
                UserControl.PaintPicture LoadResPicture("BTN_DEFAULT", vbResBitmap), 2 + Abs(Value), 2 + Abs(Value), UserControl.ScaleWidth - 4, UserControl.ScaleHeight - 4
            ElseIf Cnc Then
                'cancel button
                UserControl.PaintPicture LoadResPicture("BTN_CANCEL", vbResBitmap), 2 + Abs(Value), 2 + Abs(Value), UserControl.ScaleWidth - 4, UserControl.ScaleHeight - 4
            Else
                'doesn't have focus
                UserControl.PaintPicture LoadResPicture("BTN_NOFOCUS", vbResBitmap), 2 + Abs(Value), 2 + Abs(Value), UserControl.ScaleWidth - 4, UserControl.ScaleHeight - 4
            End If
        End If
        On Error GoTo 0
    Else
        'disabled
        UserControl.PaintPicture LoadResPicture("BTN_DISABLED", vbResBitmap), 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight
    End If
End If
If Len(pPictureResID) > 0 Then
    On Error GoTo PicNoResID
    If IsNumeric(pPictureResID) Then
        Set tmpPicture = LoadResPicture(CInt(pPictureResID), pPictureResType)
    Else
        Set tmpPicture = LoadResPicture(CStr(pPictureResID), pPictureResType)
    End If
    px = Int(UserControl.ScaleWidth - ScaleX(tmpPicture.Width, vbHimetric, vbPixels)) \ 2 + Abs(Value)
    py = Int(UserControl.ScaleHeight - ScaleY(tmpPicture.Height, vbHimetric, vbPixels)) \ 2 + Abs(Value)
    UserControl.PaintPicture tmpPicture, px, py
    If HasFocus Then
        If UserControl.ScaleWidth >= 9 And UserControl.ScaleHeight >= 9 Then
            UserControl.DrawMode = VBRUN.DrawModeConstants.vbXorPen
            UserControl.Line (3, 3)- _
                             (UserControl.ScaleWidth - 4, UserControl.ScaleHeight - 4), &H808080, B
            UserControl.DrawMode = VBRUN.DrawModeConstants.vbCopyPen
        End If
    End If
    Set tmpPicture = Nothing
End If
PicResIDFail:
On Error GoTo eh2
If En Then
    If Value Then
        UserControl.Line (0, 0)-(UserControl.ScaleWidth - 1, 0), RGB(0, 0, 0)
        UserControl.Line (0, 0)-(0, UserControl.ScaleHeight - 1), RGB(0, 0, 0)
        UserControl.Line (0, UserControl.ScaleHeight - 1)-(UserControl.ScaleWidth, UserControl.ScaleHeight - 1), RGB(255, 255, 255)
        UserControl.Line (UserControl.ScaleWidth - 1, 0)-(UserControl.ScaleWidth - 1, UserControl.ScaleHeight), RGB(255, 255, 255)
        UserControl.Line (1, 1)-(UserControl.ScaleWidth - 2, 1), RGB(157, 157, 161)
        UserControl.Line (1, 1)-(1, UserControl.ScaleHeight - 2), RGB(157, 157, 161)
        UserControl.Line (UserControl.ScaleWidth - 2, UserControl.ScaleHeight - 2)-(0, UserControl.ScaleHeight - 2), RGB(241, 239, 226)
        UserControl.Line (UserControl.ScaleWidth - 2, UserControl.ScaleHeight - 2)-(UserControl.ScaleWidth - 2, 0), RGB(241, 239, 226)
    Else
        UserControl.Line (0, 0)-(UserControl.ScaleWidth - 1, 0), RGB(255, 255, 255)
        UserControl.Line (0, 0)-(0, UserControl.ScaleHeight - 1), RGB(255, 255, 255)
        UserControl.Line (0, UserControl.ScaleHeight - 1)-(UserControl.ScaleWidth, UserControl.ScaleHeight - 1), RGB(0, 0, 0)
        UserControl.Line (UserControl.ScaleWidth - 1, 0)-(UserControl.ScaleWidth - 1, UserControl.ScaleHeight), RGB(0, 0, 0)
        UserControl.Line (1, 1)-(UserControl.ScaleWidth - 2, 1), RGB(241, 239, 226)
        UserControl.Line (1, 1)-(1, UserControl.ScaleHeight - 2), RGB(241, 239, 226)
        UserControl.Line (UserControl.ScaleWidth - 2, UserControl.ScaleHeight - 2)-(0, UserControl.ScaleHeight - 2), RGB(157, 157, 161)
        UserControl.Line (UserControl.ScaleWidth - 2, UserControl.ScaleHeight - 2)-(UserControl.ScaleWidth - 2, 0), RGB(157, 157, 161)
    End If
End If
UserControl.CurrentX = (UserControl.ScaleWidth - UserControl.TextWidth(tmpCaption)) \ 2 + Abs(Value)
UserControl.CurrentY = (UserControl.ScaleHeight - UserControl.TextHeight(tmpCaption)) \ 2 + Abs(Value)
dbPrint tmpCaption
Exit Sub
PicNoResID:
Resume PicResIDFail
Exit Sub
eh:
strCaption = MyOrigCaption
Resume Next
eh2:
If Err.Number = 398 Then
    'client site not availiable. Exit here.
Else
    Resume Next
End If
End Sub


Private Sub dbPrint(St As String)
Dim i As Integer, m As String, s As Boolean
Dim Rct As pRECT
Dim DTP As pDRAWTEXTPARAMS
Dim h As Long
If Len(St) = 0 Then Exit Sub
DTP.cbSize = LenB(DTP)
DTP.iLeftMargin = 1
DTP.iRightMargin = 1
DTP.iTabLength = 4

Rct.Left = 2
Rct.Top = 2
Rct.Right = UserControl.ScaleWidth - 2
Rct.Bottom = UserControl.ScaleHeight - 2

pDrawTextEx UserControl.hDC, St, Len(St), Rct, _
           DT_CENTER Or _
           DT_WORDBREAK Or DT_CALCRECT, DTP

h = Rct.Bottom - Rct.Top
Rct.Left = 2
Rct.Top = (UserControl.ScaleHeight - h) \ 2
Rct.Bottom = Rct.Top + h
Rct.Right = UserControl.ScaleWidth - 2

pDrawTextEx UserControl.hDC, St, Len(St), Rct, _
           DT_CENTER Or _
           DT_WORDBREAK, DTP

's = False
'UserControl.Font.Underline = False
'    For i = 1 To Len(St)
'        m = Mid$(St, i, 1)
'        If m = "&" Then
'            UserControl.Font.Underline = True
'            s = True
'        Else
'            UserControl.Print m;
'            If s Then
'                UserControl.Font.Underline = False
'                s = False
'            End If
'        End If
'    Next i
End Sub

Private Sub UserControl_AmbientChanged(PropertyName As String)
Select Case UCase$(PropertyName)
    Case "DEFAULT"
        Default1 = Extender.Default
    Case "CANCEL"
        Cancel1 = Extender.Cancel
End Select
End Sub

Private Function IsKey(KeyCode As Long) As Boolean
IsKey = (Abs(GetKeyState(KeyCode)) > 1)
End Function

Private Sub UserControl_DblClick()
If IsKey(1) Then
    UserControl_MouseDown 1, 0, 0, 0
End If
End Sub

Private Sub UserControl_GotFocus()
HasFocus = True
Refresh
End Sub

Private Sub UserControl_Hide()
pVisible = False
End Sub

Private Sub UserControl_LostFocus()
HasFocus = False
Refresh
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Debug.Print "down"
If Not En Then Exit Sub
pVisible = True
If Button > 0 Then
    If MouseState(1) Then
        Exit Sub
    End If
    MouseState(Button) = True
End If
If Button = 1 Then
    BV = True
    Refresh
    tmpMM = True
End If
RaiseEvent MouseDown(Button, Shift, X, Y)
'DoEvents
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim ir As Boolean
If Not En Then
    Exit Sub
End If
If Button > 0 Then

    If Not (MouseState(Button)) Then
        Exit Sub
    End If
End If
    
If Button = 1 Then
    ir = IsRgn(X, Y)
    If ir <> tmpMM Then
        BV = ir
        Refresh
        tmpMM = ir
    End If
End If
RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Debug.Print "up"
If Not En Then Exit Sub
If Button > 0 Then
    If Not (MouseState(Button)) Then
        Exit Sub
    End If
End If
MouseState(Button) = False
If Button = 1 And BV = True Then
'    DoEvents
    If IsRgn(X, Y) Then RaiseEvent Click
    BV = False
    Refresh
End If
RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Function IsRgn(ByVal X As Single, Y As Single) As Boolean
IsRgn = ((X >= 0) And (Y >= 0) And (X <= UserControl.ScaleWidth - 1) And (Y <= UserControl.ScaleHeight - 1))
End Function

Private Sub UserControl_AccessKeyPress(KeyAscii As Integer)
If Not En Then Exit Sub
Select Case KeyAscii
Case 13, 27
    RaiseEvent Click
Case Asc(UserControl.AccessKeys)
    UserControl.SetFocus
End Select
End Sub

Private Sub UserControl_Initialize()
If ResID <> 0 Then
    strCaption = LoadResString(ResID)
End If
End Sub

Private Sub UserControl_InitProperties()
En = True
UserControl.BackColor = &H8000000F
UserControl.ForeColor = &H80000012
UserControl.MousePointer = vbDefault
MyOrigCaption = UserControl.Name
NPP = False
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
If Not En Then Exit Sub
RaiseEvent KeyDown(KeyCode, Shift)
If KeyCode = 32 Then
    UserControl_MouseDown 1, Shift, UserControl.ScaleWidth \ 2, UserControl.ScaleHeight \ 2
End If
End Sub


Private Sub btp_KeyPress(KeyAscii As Integer)
If Not En Then Exit Sub
If KeyAscii = 13 Then
    RaiseEvent Click
End If
RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
If Not En Then Exit Sub
If KeyCode = 32 Then
If BV Then
    UserControl_MouseUp 1, Shift, UserControl.ScaleWidth \ 2, UserControl.ScaleHeight \ 2
End If
End If
RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub UserControl_Paint()
If pAutoRedraw Then
  If Not UserControl.AutoRedraw Then
    UserControl.AutoRedraw = True
    Refresh
    Exit Sub
  End If
    FreshButton BV
Else
'    Debug.Print "paint " + CStr(Rnd(1))
    FreshButton BV
End If
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
Dim i As Integer, m As String
Dim pArr As Variant
On Error Resume Next
'Debug.Print Timer, " ReadProperties"
With PropBag
    If .ReadProperty("OthersPresent", False) Then
    
        pArr = Split(.ReadProperty("Others"), Chr$(1))
    
        MyOrigCaption = pArr(0) '.ReadProperty("Caption", "")
        i = InStr(strCaption, "&")
        If i = 0 Or i = Len(strCaption) Then
            m = ""
        Else
            m = Mid(strCaption, i + 1, 1)
        End If
        UserControl.AccessKeys = m
        
        UserControl.BackColor = CLng(pArr(1)) '.ReadProperty("BackColor")
        UserControl.ForeColor = CLng(pArr(2)) '.ReadProperty("ForeColor")
        'Set bPict.Picture = .ReadProperty("Picture")
        En = CBool(pArr(4)) '.ReadProperty("Enabled")
        UserControl.Enabled = En
        UserControl.MousePointer = CInt(pArr(5)) '.ReadProperty("MousePointer")
        Set UserControl.MouseIcon = .ReadProperty("MouseIcon")
        NPP = CBool(pArr(7)) '.ReadProperty("No3D", False)
        ResID = CLng(pArr(8)) ' .ReadProperty("ResID", 0)
        UserControl_Initialize
        pPictureResID = CVar(pArr(9)) '.ReadProperty("PictureResID", "")
        pPictureResType = CInt(pArr(10))  '.ReadProperty("PictureResType", vbResBitmap)
        
        Def = CBool(pArr(11)) ' .ReadProperty("Default1", False)
        Cnc = CBool(pArr(12))  '.ReadProperty("Cancel1", False)
        Set UserControl.Font = .ReadProperty("Font", UserControl.Font)
        UserControl.FontSize = UserControl.Font.Size
        pNoTags = CBool(pArr(13))  '.ReadProperty("NoTags", False)
        
        dbTag1 = pArr(14)  '.ReadProperty("dbTag1", "")
        dbTag2 = pArr(15)  '.ReadProperty("dbTag2", "")
        dbTag3 = pArr(16)  '.ReadProperty("dbTag3", "")
        
        If UBound(pArr) >= 17 Then
            pAutoRedraw = CBool(pArr(17))
        End If
    Else
        MyOrigCaption = .ReadProperty("Caption", "")
        i = InStr(strCaption, "&")
        If i = 0 Or i = Len(strCaption) Then
            m = ""
        Else
            m = Mid(strCaption, i + 1, 1)
        End If
        UserControl.AccessKeys = m
        
        UserControl.BackColor = .ReadProperty("BackColor")
        UserControl.ForeColor = .ReadProperty("ForeColor")
        'Set bPict.Picture = .ReadProperty("Picture")
        En = .ReadProperty("Enabled")
        UserControl.Enabled = En
        UserControl.MousePointer = .ReadProperty("MousePointer")
        Set UserControl.MouseIcon = .ReadProperty("MouseIcon")
        NPP = .ReadProperty("No3D", False)
        ResID = .ReadProperty("ResID", 0)
        UserControl_Initialize
        pPictureResID = .ReadProperty("PictureResID", "")
        pPictureResType = .ReadProperty("PictureResType", vbResBitmap)
        
        Def = .ReadProperty("Default1", False)
        Cnc = .ReadProperty("Cancel1", False)
        Set UserControl.Font = .ReadProperty("Font", UserControl.Font)
        UserControl.FontSize = UserControl.Font.Size
        pNoTags = .ReadProperty("NoTags", False)
        
        dbTag1 = .ReadProperty("dbTag1", "")
        dbTag2 = .ReadProperty("dbTag2", "")
        dbTag3 = .ReadProperty("dbTag3", "")
        
        pAutoRedraw = .ReadProperty("AutoRedraw", False)
    End If
    If Ambient.UserMode = False Then
        'pAutoRedraw = True
    End If
End With
End Sub

Private Sub UserControl_Resize()
If pAutoRedraw Then
  Dim tmp As Boolean
  Cls
'  UserControl.AutoRedraw = False
'  UserControl.AutoRedraw = True
End If
Refresh
End Sub

Private Sub UserControl_Show()
pVisible = True
UserControl.AutoRedraw = pAutoRedraw
Refresh
End Sub

'Properties Sequence:
'0  Caption
'1  BackColor
'2  ForeColor
'3  Picture
'4  Enabled
'5  MousePointer
'6  MouseIcon
'7  No3D
'8  ResID
'9  PictureResID
'10 PictureResType
'11 Default1
'12 Cancel1
'13 NoTags
'
'14 dbTag1
'15 dbTag2
'16 dbTag3

'17 AutoReddraw

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
Dim pArr() As String
'On Error Resume Next
'Debug.Print Timer; " WriteProperties"
With PropBag
    ReDim pArr(0 To 17)
'    .WriteProperty "Caption", MyOrigCaption, ""
    pArr(0) = MyOrigCaption
'    .WriteProperty "BackColor", , ""
    pArr(1) = CStr(UserControl.BackColor)
'    .WriteProperty "ForeColor", , ""
    pArr(2) = CStr(UserControl.ForeColor)
'    .WriteProperty "Picture", bPict.Picture, ""
'    pArr(3) = m
'    .WriteProperty "Enabled", , ""
    pArr(4) = CStr(En)
'    .WriteProperty "MousePointer",
    pArr(5) = CStr(UserControl.MousePointer)
    .WriteProperty "MouseIcon", UserControl.MouseIcon
'    pArr(6) =
'    .WriteProperty "No3D",, False
    pArr(7) = CStr(NPP)
'    .WriteProperty "ResID", , 0
    pArr(8) = CStr(ResID)
'    .WriteProperty "PictureResID", , ""
    pArr(9) = CStr(pPictureResID)
'    .WriteProperty "PictureResType", , ""
    pArr(10) = CStr(pPictureResType)
'    .WriteProperty "Default1", , False
    pArr(11) = CStr(Def)
'    .WriteProperty "Cancel1", , False
    pArr(12) = CStr(Cnc)
    .WriteProperty "Font", UserControl.Font
'    pArr(14) = m
'    .WriteProperty "NoTags",
    pArr(13) = CStr(pNoTags)
'
'    .WriteProperty "dbTag1", , ""
    pArr(14) = dbTag1
'    .WriteProperty "dbTag2", , ""
    pArr(15) = dbTag2
'    .WriteProperty "dbTag3", , ""
    pArr(16) = dbTag3
    
    pArr(17) = pAutoRedraw
    
    .WriteProperty "Others", Join(pArr, Chr$(1))
    .WriteProperty "OthersPresent", True
End With
End Sub

Public Property Get Caption() As String
Attribute Caption.VB_Description = "Returns or sets the title displayed in center of the object"
Attribute Caption.VB_ProcData.VB_Invoke_Property = ";Appearance"
Attribute Caption.VB_UserMemId = -518
Caption = MyOrigCaption
End Property

Public Property Let Caption(ByVal nC As String)
Dim i As Integer, m As String
strCaption = nC
MyOrigCaption = nC
i = InStr(strCaption, "&")
If i = 0 Or i = Len(strCaption) Then
    m = ""
Else
    m = Mid(strCaption, i + 1, 1)
End If
UserControl.AccessKeys = m
Refresh
End Property

Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_UserMemId = 0
BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal nC As OLE_COLOR)
On Error Resume Next
UserControl.BackColor = nC
Refresh
End Property

Public Property Get ForeColor() As OLE_COLOR
ForeColor = UserControl.ForeColor
End Property

Public Property Let ForeColor(ByVal vNewValue As OLE_COLOR)
UserControl.ForeColor = vNewValue
Refresh
End Property

Public Property Get Enabled() As Boolean
Enabled = En
End Property

Public Property Let Enabled(ByVal vNewValue As Boolean)
Dim Adj As Boolean
Adj = Not En = vNewValue
If Adj Then
    En = vNewValue
    UserControl.Enabled = vNewValue
    Refresh
End If
End Property

Public Property Get MousePointer() As MousePointerConstants
MousePointer = UserControl.MousePointer
End Property

Public Property Let MousePointer(ByVal vNewValue As MousePointerConstants)
UserControl.MousePointer = vNewValue
End Property

Public Property Get MouseIcon() As stdole.StdPicture
Attribute MouseIcon.VB_ProcData.VB_Invoke_Property = ";Behavior"
Set MouseIcon = UserControl.MouseIcon
End Property

Public Property Set MouseIcon(ByVal NI As StdPicture)
If Not NI.Type = vbPicTypeIcon Then Error 5
Set UserControl.MouseIcon = NI
End Property

Public Property Get No3D() As Boolean
No3D = NPP
End Property

Public Property Let No3D(ByVal vNewValue As Boolean)
NPP = vNewValue
Refresh
End Property

Public Property Get Default1() As Boolean
Default1 = Def
End Property

Public Property Let Default1(ByVal vNewValue As Boolean)
Def = vNewValue
Refresh
End Property

Public Property Get Cancel1() As Boolean
Cancel1 = Cnc
End Property

Public Property Let Cancel1(ByVal vNewValue As Boolean)
Cnc = vNewValue
Refresh
End Property

Public Property Get PictureResID() As Variant
PictureResID = pPictureResID
End Property

Public Property Let PictureResID(ByVal ppPictureResID As Variant)
pPictureResID = ppPictureResID
Refresh
End Property

Public Property Get Font() As StdFont
Set Font = UserControl.Font
End Property

Public Property Set Font(ByVal vNewValue As StdFont)
Set UserControl.Font = vNewValue
UserControl.FontSize = UserControl.Font.Size
Refresh
End Property

Public Property Get NoTags() As Boolean
NoTags = pNoTags
End Property

Public Property Let NoTags(ByVal bNew As Boolean)
pNoTags = bNew
If BV Then
    Refresh
End If
End Property

Public Property Get PictureResType() As LoadResConstants
PictureResType = pPictureResType
End Property

Public Property Let PictureResType(ByVal lNew As LoadResConstants)
pPictureResType = lNew
Refresh
End Property

Public Property Get AutoRedraw() As Boolean
AutoRedraw = pAutoRedraw
End Property

Public Property Let AutoRedraw(ByVal vNewValue As Boolean)
Dim Changed As Boolean
Changed = vNewValue <> pAutoRedraw
pAutoRedraw = vNewValue
If Changed Then
    UserControl.AutoRedraw = pAutoRedraw
    UserControl.Cls
    Refresh
End If
End Property

Public Sub Refresh()
If pAutoRedraw Then
    FreshButton BV
    UserControl.Refresh
Else
    UserControl.Refresh
End If
End Sub


Public Sub RaiseClick()
UserControl_KeyDown vbKeySpace, 0
UserControl_KeyUp vbKeySpace, 0
End Sub
