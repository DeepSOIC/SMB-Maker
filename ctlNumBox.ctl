VERSION 5.00
Begin VB.UserControl ctlNumBox 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   KeyPreview      =   -1  'True
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   ToolboxBitmap   =   "ctlNumBox.ctx":0000
   Begin VB.Timer tmrApply 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   1410
      Top             =   2385
   End
   Begin VB.PictureBox Picture1 
      Height          =   225
      Left            =   585
      ScaleHeight     =   11
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   133
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   405
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   585
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   90
      Width           =   2055
   End
   Begin VB.Menu mnuPP 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu mnuNativeValue 
         Caption         =   "<native value>"
         Index           =   0
      End
   End
End
Attribute VB_Name = "ctlNumBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim Cv As Variant
Dim Typ As VbVarType
Dim tmpText As String
Dim bLock As Boolean
Dim NoDraw As Boolean
Dim bErr As Boolean

Dim pMin As Double, pMax As Double
Dim lVal As Double

Dim pHorzMode As Boolean

Dim pEditName As String

Dim Pow As Single
Dim NLn As Single

Dim txtToolTip As String
Dim Reselect As Boolean

Dim pSliderVisible As Boolean


Private Type NativeValue
    Name As String
    Value As String
End Type

Dim lstNativeValues() As NativeValue
Dim nNativeValues As Long
Dim pNativeValuesResID As Integer

Dim pEn As Boolean

Public Event InputChange()
Public Event Change()
Attribute Change.VB_MemberFlags = "200"


Private Sub mnuNativeValue_Click(Index As Integer)
If nNativeValues < 1 Then Exit Sub
On Error Resume Next
Text1.Text = lstNativeValues(Index).Value
End Sub

Private Sub Picture1_GotFocus()
On Error Resume Next
'Text1.SetFocus
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 1 Then
    tmpText = Text1.Text
    Picture1_MouseMove Button, Shift, x, y
ElseIf Button = 2 Then
    UpdateMenu
    If nNativeValues > 0 Then
        PopupMenu mnuPP, vbPopupMenuRightButton
    End If
End If
End Sub

Friend Sub UpdateMenu()
Dim i As Long
If nNativeValues < 1 Then Exit Sub
On Error Resume Next
For i = mnuNativeValue.UBound To nNativeValues Step -1
    Unload mnuNativeValue(i)
Next i
For i = mnuNativeValue.lBound + 1 To nNativeValues - 1
    Load mnuNativeValue(i)
    mnuNativeValue(i).Visible = True
Next i
For i = 0 To nNativeValues - 1
    mnuNativeValue(i).Caption = lstNativeValues(i).Name
Next i
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 1 Then
    If y < -100 Or y > Picture1.ScaleHeight + 100 - 1 Then
        Text1.Text = tmpText
    Else
        If x < 0 Then x = 0
        If x > Picture1.ScaleWidth - 1 Then x = Picture1.ScaleWidth - 1
        Cv = (x / (Picture1.ScaleWidth - 1)) ^ (1 / Pow) * (pMax - pMin) + pMin
        bLock = True
            Text1.Text = dbCStr(CVtyp(Cv, Typ))
        bLock = False
        RaiseEvent InputChange
    End If
    Picture1_Paint
End If
End Sub

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
RaiseEvent InputChange
RaiseEvent Change
End Sub

Private Sub Picture1_Paint()
On Error Resume Next
Dim i As Long
If NoDraw Then Exit Sub
If Not pSliderVisible Then Exit Sub
If pEn Then
    Picture1.PaintPicture LoadResPicture(107, vbResBitmap), 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight
    For i = 0 To 10
        Picture1.DrawMode = vbMergePen
        Picture1.Line ((i / 10) ^ Pow * (Picture1.ScaleWidth - 1), 0)-Step(0, Picture1.ScaleHeight), RGB(255, 0, 255)
    Next i
    
    If Not bErr Then
        Picture1.DrawMode = DrawModeConstants.vbMaskNotPen
        Picture1.Line (((Cv - pMin) / (pMax - pMin)) ^ Pow * (Picture1.ScaleWidth - 1), 0)-Step(0, Picture1.ScaleHeight), RGB(0, 255, 0)
    Else
        Picture1.DrawMode = vbCopyPen
        Picture1.CurrentX = (Picture1.ScaleWidth - Picture1.TextWidth("Error")) / 2
        Picture1.CurrentY = (Picture1.ScaleHeight - Picture1.TextHeight("Error")) / 2
        Picture1.Print "Error!"
    End If
Else
    Picture1.PaintPicture LoadResPicture("BTN_DISABLED", vbResBitmap), 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight
End If
End Sub


Private Sub Picture1_Resize()
Picture1.Refresh
End Sub

Private Sub Text1_Change()
On Error Resume Next
If bLock Then Exit Sub
Err.Clear
ValidateData False
bErr = Err.Number <> 0
If Not bErr Then RaiseEvent InputChange
Picture1_Paint
tmrApply.Enabled = False
tmrApply.Enabled = True
End Sub

Private Sub Text1_GotFocus()
Text1.SelStart = 0
Text1.SelLength = Len(Text1.Text)
UserControl_GotFocus
ShowToolTip UserControl.hWnd, txtToolTip
Reselect = True
End Sub

Private Sub Text1_LostFocus()
On Error Resume Next
HideToolTipWindow
Err.Clear
ValidateData True
bErr = Err.Number <> 0
If bErr Then
    UserControl_KeyDown 27, 0
End If
RaiseEvent Change

Exit Sub
rsm:
On Error Resume Next
'Text1.SetFocus
Exit Sub
eh:
'MsgBox Err.Description
Resume rsm
End Sub

Private Sub Text1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Reselect = False
End Sub

Private Sub Text1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If Reselect Then
    Reselect = False
    Text1.SelStart = 0
    Text1.SelLength = Len(Text1.Text)
End If
End Sub

Private Sub tmrApply_Timer()
tmrApply.Enabled = False
On Error Resume Next
If bLock Then Exit Sub
Err.Clear
ValidateData False
bErr = Err.Number <> 0
If Not bErr Then
    RaiseEvent Change
End If
End Sub

Private Sub UserControl_GotFocus()
On Error Resume Next
lVal = pMin
lVal = Value
End Sub

Private Sub UserControl_Initialize()
pMin = 0
pMax = 10000000000#
Pow = 0.5
NLn = 1
pEn = True
pSliderVisible = True
End Sub

Public Sub ValidateData(Optional ByVal ReplaceWithValue As Boolean = False)
Dim EV As New clsEVal
Dim Vars() As Variable
Dim n As Double
Text1.Text = Trim(Text1.Text)
If Len(Text1.Text) = 0 Then
    Err.Raise 120, "ValidateData", "Empty string is not allowed."
End If

If Left$(Text1.Text, 1) = "=" Then
    n = EV.EVal(Right$(Text1.Text, Len(Text1.Text) - 1), "Max", pMax, "Min", pMin)
Else
    n = dbVal(Text1.Text, Typ)
End If
If n > pMax Or n < pMin Then
    Err.Raise 120, "ValidateData", "Min/Max limit exceeded."
End If
If ReplaceWithValue Then
    Text1.Text = dbCStr(CVtyp(n, Typ))
End If
Cv = n
End Sub


Public Property Get Min() As Double
Min = pMin
End Property

Public Property Let Min(ByVal aMin As Double)
If aMin >= pMax Then
    Err.Raise 120, "Min", "Min is above or equal to max. Set Max first."
End If
pMin = aMin
If Cv < pMin Then Value = pMin
Picture1.Refresh
UpdateToolTip
End Property



Public Property Get Max() As Double
Max = pMax
End Property

Public Property Let Max(ByVal aMax As Double)
If pMin >= aMax Then
    Err.Raise 120, "Min", "Max is less or equal to Min. Set Min first."
End If
pMax = aMax
If Cv > pMax Then Value = pMax
Picture1.Refresh
UpdateToolTip
End Property

Public Function CVtyp(ByRef Value As Variant, ByVal Typ As VbVarType)
Select Case Typ
    Case VbVarType.vbBoolean
        CVtyp = CBool(Value)
    Case VbVarType.vbByte
        CVtyp = CByte(Value)
    Case VbVarType.vbCurrency
        CVtyp = CCur(Value)
    Case VbVarType.vbDecimal
        CVtyp = CDec(Value)
    Case VbVarType.vbDouble
        CVtyp = CDbl(Value)
    Case VbVarType.vbInteger
        CVtyp = CInt(Value)
    Case VbVarType.vbLong
        CVtyp = CLng(Value)
    Case VbVarType.vbSingle
        CVtyp = CSng(Value)
    Case VbVarType.vbVariant
        CVtyp = Value
End Select
End Function

Public Property Get NumType() As VbVarType
NumType = Typ
End Property

Public Property Let NumType(ByVal nTyp As VbVarType)
Select Case nTyp
    Case VbVarType.vbBoolean, _
         VbVarType.vbByte, _
         VbVarType.vbCurrency, _
         VbVarType.vbDecimal, _
         VbVarType.vbDouble, _
         VbVarType.vbInteger, _
         VbVarType.vbLong, _
         VbVarType.vbSingle
        Typ = nTyp
    Case Else
        Err.Raise 120, "NumType[Property Get]", "Type " + TypeName(nTyp) + " is not supported."
End Select
If Not NoDraw Then
    Picture1.Refresh
End If
End Property



Public Property Get Value() As Variant
Attribute Value.VB_UserMemId = 0
On Error Resume Next
Err.Clear
ValidateData ReplaceWithValue:=False
Value = CVtyp(Cv, Typ)
If Err.Number <> 0 Then
    Value = lVal
End If
End Property

Public Property Let Value(ByVal vNew As Variant)
On Error GoTo eh
If vNew < pMin Then vNew = pMin
If vNew > pMax Then vNew = pMax
Cv = CVtyp(vNew, Typ)
'bLock = True
Text1.Text = dbCStr(vNew)
tmrApply.Enabled = False
'bLock = False
Picture1.Refresh
Exit Property
eh:
bLock = False
Err.Raise Err.Number, "Value", Err.Description
End Property

Private Sub UserControl_InitProperties()
pMin = 0
pMax = 10000000000#
Typ = vbSingle
NonLinearity = 1
Value = 0
pSliderVisible = True
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
    If lVal >= pMin And lVal <= pMax Then
        Value = lVal
    End If
End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
Select Case Chr$(KeyAscii)
    Case "+", "-", "*", "/"
        If Not Left$(Text1.Text, 1) = "=" And Not (pMin < 0 And Chr$(KeyAscii) = "-" And Text1.SelLength > 0 And Text1.SelStart = 0) Then
            Text1.Text = "=" + Text1.Text
            Text1.SelLength = 0
            Text1.SelStart = Len(Text1.Text)
        End If
    Case "="
        If Not Left$(Text1.Text, 1) = "=" Then
            Text1.Text = "=" + Text1.Text
            KeyAscii = 0
            Text1.SelLength = 0
            Text1.SelStart = Len(Text1.Text)
        End If
        
End Select
End Sub

Private Sub UserControl_LostFocus()
Text1_LostFocus
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

With PropBag
    NoDraw = True
    pMin = .ReadProperty("Min", 0)
    pMax = .ReadProperty("Max", 10000000000#)
    If pMin >= pMax Then pMax = pMin + 1
    NumType = .ReadProperty("NumType", VbVarType.vbSingle)
    Value = .ReadProperty("Value", 0)
    pHorzMode = .ReadProperty("HorzMode", False)
    UserControl_Resize
    EditName = .ReadProperty("EditName", "")
    NonLinearity = .ReadProperty("NLn", 1)
    NativeValuesResID = .ReadProperty("NativeValuesResID", 0)
    NativeValues = .ReadProperty("NativeValues", "")
    Enabled = .ReadProperty("Enabled", True)
    SliderVisible = .ReadProperty("SliderVisible", True)
    NoDraw = False
    Picture1.Refresh
End With
End Sub

Private Sub UserControl_Resize()
On Error Resume Next
If pSliderVisible Then
    If pHorzMode Then
        Text1.Move 0, 0, UserControl.ScaleWidth \ 2, UserControl.ScaleHeight
        Picture1.Move Text1.Width, 0, UserControl.ScaleWidth - Text1.Width, UserControl.ScaleHeight
    Else
        Picture1.Move 0, UserControl.ScaleHeight - 15, UserControl.ScaleWidth, 15
        Text1.Move 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight - Picture1.Height
    End If
Else
    Text1.Move 0, 0, ScaleWidth, ScaleHeight
End If
End Sub

Private Sub UserControl_Show()
Picture1_Paint
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
With PropBag
    .WriteProperty "Value", Value, 0
    .WriteProperty "Min", Min, 0
    .WriteProperty "Max", Max, 10000000000#
    .WriteProperty "NumType", NumType, VbVarType.vbSingle
    .WriteProperty "HorzMode", pHorzMode
    .WriteProperty "EditName", EditName, ""
    .WriteProperty "NLn", NonLinearity, 1
    .WriteProperty "NativeValues", NativeValues, ""
    .WriteProperty "NativeValuesResID", NativeValuesResID, 0
    .WriteProperty "Enabled", Enabled, True
    .WriteProperty "SliderVisible", SliderVisible, True
End With
End Sub




Public Function dbCStr(ByRef Value As Variant, _
                       Optional ByVal HexNumber As Boolean = False) As String
Select Case VarType(Value)
    Case VbVarType.vbBoolean
        dbCStr = CStr(Value)
    Case VbVarType.vbByte, VbVarType.vbInteger, VbVarType.vbLong 'Integers
        If HexNumber Then
            dbCStr = "&H" + Hex$(CLng(Value))
        Else
            dbCStr = Trim(Str(Value))
        End If
    Case VbVarType.vbCurrency, VbVarType.vbDecimal, VbVarType.vbDouble, VbVarType.vbSingle 'other numbers
        dbCStr = Trim(Str(Value))
    Case VbVarType.vbString
        dbCStr = Value
'    Case VbVarType.vbUserDefinedType
'        MsgBox "Direct saving of " + TypeName(Value) + "s is not currently supported. " + vbCrLf + _
'               "Attempt: " + Key + "::" + Parameter, vbCritical, "Type mismatch"
'        Exit Sub
    Case Else
        Err.Raise 118, "dbCStr", "This expression cannot be converted to a string"
End Select
End Function

Public Function dbVal(ByVal Value As String, _
                      Optional ByVal TypeID As VbVarType = vbDouble, _
                      Optional nMin As Variant, _
                      Optional nMax As Variant) As Variant
Dim Delimiter As String
Dim Result As Variant
If InStr(1, CStr(0.1), ",") > 0 Then
    Delimiter = ","
    Value = Replace(Value, ".", Delimiter)
Else
    Delimiter = "."
    Value = Replace(Value, ",", Delimiter)
End If
    Select Case TypeID
        Case VbVarType.vbBoolean
            Result = CBool(Value)
        Case VbVarType.vbByte
            Result = CByte(Value)
        Case VbVarType.vbInteger
            Result = CInt(Value)
        Case VbVarType.vbLong 'Integers
            Result = CLng(Value)
        Case VbVarType.vbCurrency
            Result = CCur(Value)
        Case VbVarType.vbDecimal
            Result = CDec(Value)
        Case VbVarType.vbDouble
            Result = CSng(Value)
        Case VbVarType.vbSingle 'other numbers
            Result = CSng(Value)
        Case VbVarType.vbString
            Result = Value
        Case Else
            Err.Raise 119, "dbVal", "Unsupported type (" + TypeID + ")."
    End Select
    
    If Not IsMissing(nMin) Then
        If Result < nMin Then
            Err.Raise 119, "dbVal", "Limit exceeded. Minimum is " + CStr(nMin) + "."
        End If
    End If
    If Not IsMissing(nMax) Then
        If Result > nMin Then
            Err.Raise 119, "dbVal", "Limit exceeded. Maximum is " + CStr(nMax) + "."
        End If
    End If
    dbVal = Result
End Function





Public Property Get HorzMode() As Boolean
HorzMode = pHorzMode
End Property

Public Property Let HorzMode(ByVal bNew As Boolean)
pHorzMode = bNew
UserControl_Resize
End Property

Public Property Get EditName() As String
EditName = pEditName
End Property

Public Property Let EditName(ByVal stNew As String)
pEditName = stNew
UpdateToolTip
End Property

Private Sub UpdateToolTip()
Dim tmp As String
If Len(pEditName) > 0 Then
    If Left$(pEditName, 1) = "$" Then
        tmp = LoadResString(Val(Mid$(pEditName, 2)))
    Else
        tmp = pEditName
    End If
End If
If Len(tmp) > 0 Then
    tmp = grs(2420, "%t", tmp + vbNewLine + vbNewLine, "%min", CStr(Min), "%max", CStr(Max))
Else
    tmp = grs(2420, "%t", "", "%min", CStr(Min), "%max", CStr(Max))
End If
txtToolTip = tmp
End Sub



Public Property Get NonLinearity() As Single
NonLinearity = NLn
End Property

Public Property Let NonLinearity(ByVal vNew As Single)
If vNew < 0 Then vNew = 0
If vNew > 5 Then vNew = 5
Pow = 2 ^ -vNew
NLn = vNew
Picture1_Paint
End Property

Public Property Get NativeValues() As String
Dim i As Long
Dim sArr() As String
If nNativeValues > 0 Then
    ReDim sArr(0 To nNativeValues * 2 - 1)
    For i = 0 To nNativeValues - 1
        sArr(i * 2) = lstNativeValues(i).Name
        sArr(i * 2 + 1) = lstNativeValues(i).Value
    Next i
    NativeValues = Join(sArr, "|")
Else
    NativeValues = ""
End If
End Property

Public Property Let NativeValues(ByVal strNew As String)
Dim sArr() As String
Dim nVals As Long
Dim i As Long
If pNativeValuesResID > 0 Then
    On Error Resume Next
    strNew = LoadResString(pNativeValuesResID)
    On Error GoTo 0
End If
sArr = Split(strNew, "|")
nVals = (UBound(sArr) + 1) \ 2
If nVals = 0 Then
    nNativeValues = 0
    Erase lstNativeValues
Else
    nNativeValues = nVals
    ReDim lstNativeValues(0 To nVals - 1)
    For i = 0 To nVals - 1
        lstNativeValues(i).Name = sArr(i * 2)
        lstNativeValues(i).Value = sArr(i * 2 + 1)
    Next i
End If
End Property

Public Property Get NativeValuesResID() As Integer
NativeValuesResID = pNativeValuesResID
End Property

Public Property Let NativeValuesResID(ByVal iNew As Integer)
pNativeValuesResID = iNew
NativeValues = NativeValues
End Property

Public Property Get Enabled() As Boolean
Enabled = pEn
End Property

Public Property Let Enabled(ByVal bNew As Boolean)
Dim bRefr As Boolean
bRefr = pEn <> bNew
pEn = bNew
If bRefr Then
    Text1.Enabled = pEn
    Picture1.Enabled = pEn
    Picture1.Refresh
End If
End Property

Public Property Get SliderVisible() As Boolean
SliderVisible = pSliderVisible
End Property

Public Property Let SliderVisible(ByVal vNew As Boolean)
If pSliderVisible <> vNew Then
    pSliderVisible = vNew
    Picture1.Visible = pSliderVisible
    UserControl_Resize
End If
End Property
