Attribute VB_Name = "MainModule"
Option Explicit



'////////////////   G L O B A L   V A R I A B L E S   \\\\\\\\\\\\\\\\\\
Public Stepen2(0 To 8) As Integer, Stepen256(0 To 3) As Long, RGBMask(1 To 3) As Long
Public Stepen2Long(0 To 31) As Long
Public ExePath As String
Global CDl As CommonDlg
Global FirstLoad As Boolean
Global ExeCount As Long
Global TransData() As Long, TransOrigData() As Long, CurSel_SelData() As Long
Global gBackPicture As IPictureDisp
Global SMBDefPal() As Long
Global gReg As Reg
Global InitDirs() As String
Global DontDoEvents As Boolean
Dim TimerInitialized As Boolean
Global MoveTimerRes As Long
Global PrgDrawMode As PDM
Global PrgBuffer() As Double, PrgBufferLen As Long
Global HelpWindowVisible As Boolean
Global LensWindowVisible As Boolean
Global ToolTipWindowVisible As Boolean
Global CurHelpResId As Integer
Dim HookHandle As Long, HookEnabled As Boolean
Global Keyb As New clsShortcuts
Private DontDoEventsStack(0 To 100) As Boolean
Private DontDoEventsStackPointer As Long
Public CurDll As String 'the name of a dll to find (for dbInputBox)
Public MoveMouseDx As Long, MoveMouseDy As Long
Global gProjectCompiled As Boolean
Public VistaSetPixelBugDetected As Boolean

Private OriginalMPWndProc As Long
Private OriginalMWWndProc As Long


'Global MPUpdateRect As RECT
'***************///G L O B A L   V A R I A B L E S///**************************************



'***************   G L O B A L   C O N S T A N T S   **************************************
Public Const DIB_RGB_COLORS = 0
Public Const RedMask As Long = &HFF&
Public Const GreenMask As Long = &HFF00&
Public Const BlueMask As Long = &HFF0000
Public Const dbCWS As Long = 32755
Public Const dbULE As Long = 12321
Public Const dbCantCreatePic = 480
Public Const dbErr_FileNew = 1001
Global Const RES_BackPicture = 108
Global Pi As Double
Global Const AppTitle = "SMBMaker"
Public Const FileMustExist = &H1000 Or &H4
Public Const OverWritePrompt = &H2
Public Const Max_Pixels_to_Draw = 96 * 96
Global Const MaxPrgBufferLength = 1024& * 1024&
Global Const LastActionNumber = cmdLast
Public Const SRCCOPY = &HCC0020
'Public Const SMBDefKeys As String = "003F0200F0202D0202A02E100433823064231652326623367234682366A2376B2386C2396D24E0307204071060785D2710525309153083530A4530707B024082E2592F2430C22D0C2560D12D0D25212056350742B0732900D29272053720B8263A8253B8283C8273D8683E8643F862408664106B3306D341BB330BD3424D502356985A588535A8585901B0F241702468012E0C12E1025102054162427F2547C65B4B"
#Const DllPath = "D:\dbnz\DelphiPrograms\DLL\SMB Maker\SMBMakerDll.dll"
Public Const STT_READY = 1204
Public Const STT_Copying = 1200
Public Const STT_BUD = 1202
Public Const STT_Processing = 1203
Public Const STT_Loading = 1209
Public Const STT_Working = 1212
Public Const STT_Cancelled = 1217
Public Const STT_Resizing = 1236
Public Const STT_Displaying = 1218
Public Const STT_Error = 1227
Global Const ToolPen = 0
Global Const ToolLine = 1
Global Const ToolFade = 2
Global Const ToolStar = 3
Global Const ToolFStar = 4
Global Const ToolColorSel = 5
Global Const ToolVFade = 6
Global Const ToolPaint = 7
Global Const ToolCircle = 8
Global Const ToolHFade = 9
Global Const ToolSel = 10
Global Const ToolRect = 11
Global Const ToolPoly = 12
Global Const ToolAir = 13
Global Const ToolHelix = 14
Global Const ToolBrush = 16
Global Const ToolPal = 17
Global Const ToolText = 18
Global Const ToolWav = 19
Global Const ToolOrg = 20
Global Const ToolProg = 21
'***************///G L O B A L   C O N S T A N T S///**************************************




'*******************************   C O D E   ******************************************************

Public Function grs(ByVal ResID As Integer, ParamArray vReplacements() As Variant) As String
    Dim intMacro As Integer
    Dim strResString As String
    Dim ErrNum As Boolean
    Dim ErrDesc As Boolean
    Dim DllErr As Boolean
    Dim i As Integer
    Dim varReplacements() As String
    If UBound(vReplacements) = -1 Then
        ReDim varReplacements(0 To 0)
    Else
        ReDim varReplacements(LBound(vReplacements) To UBound(vReplacements))
    End If
    
    For i = 0 To UBound(vReplacements)
        varReplacements(i) = vReplacements(i)
    Next i
    
    For i = LBound(varReplacements) To UBound(varReplacements) Step 2
        Select Case UCase$(varReplacements(i))
        Case "Err.Number"
            ErrNum = True
        Case "Err.Description"
            ErrDesc = True
        Case "Err.LastDllError"
            DllErr = True
        End Select
    Next i
    If Not ErrNum Then
        If UBound(varReplacements) = 0 Then
            ReDim varReplacements(0 To 1)
        Else
            ReDim Preserve varReplacements(LBound(varReplacements) To UBound(varReplacements) + 2)
        End If
        varReplacements(UBound(varReplacements) - 1) = "Err.Number"
        varReplacements(UBound(varReplacements)) = Err.Number
    End If
    If Not ErrDesc Then
        If UBound(varReplacements) = 0 Then
            ReDim varReplacements(0 To 1)
        Else
            ReDim Preserve varReplacements(LBound(varReplacements) To UBound(varReplacements) + 2)
        End If
        varReplacements(UBound(varReplacements) - 1) = "Err.Description"
        varReplacements(UBound(varReplacements)) = Err.Description
    End If
    If Not DllErr Then
        If UBound(varReplacements) = 0 Then
            ReDim varReplacements(0 To 1)
        Else
            ReDim Preserve varReplacements(LBound(varReplacements) To UBound(varReplacements) + 2)
        End If
        varReplacements(UBound(varReplacements) - 1) = "Err.LastDllError"
        varReplacements(UBound(varReplacements)) = Err.LastDllError
    End If
    
    
    strResString = LoadResString(ResID)
    
    ' For each macro/value pair passed in...
    For intMacro = LBound(varReplacements) To UBound(varReplacements) Step 2
        Dim strMacro As String
        Dim strValue As String
        
        strMacro = varReplacements(intMacro)
        On Error GoTo MismatchedPairs
        strValue = varReplacements(intMacro + 1)
        On Error GoTo 0
        
        ' Replace all occurrences of strMacro with strValue
        Dim intPos As Integer
        Do
            intPos = InStr(strResString, strMacro)
            If intPos > 0 Then
                strResString = Left$(strResString, intPos - 1) & strValue & Right$(strResString, Len(strResString) - Len(strMacro) - intPos + 1)
            End If
        Loop Until intPos = 0
    Next intMacro
    
    grs = strResString
    
    Exit Function
    
MismatchedPairs:
    Resume Next
End Function

Function dbMsgBox(strMessage As Variant, intStyle As VbMsgBoxStyle) As VbMsgBoxResult
Dim strText As String, strCaption As String, intPos As Long
If IsNumeric(strMessage) Then
    strMessage = grs(strMessage)
ElseIf Mid(strMessage, 1, 1) = "$" Then
    strMessage = grs(CInt(Mid(strMessage, 2, Len(strMessage) - 1)))
End If
intPos = InStr(strMessage, "`")
If intPos = 0 Then
    strText = strMessage
    strCaption = ""
Else
    strText = Mid$(strMessage, 1, intPos - 1)
    If Not intPos = Len(strMessage) Then
        strCaption = Mid$(strMessage, intPos + 1, Len(strMessage) - intPos)
    Else
        strCaption = ""
    End If
End If
dbMsgBox = MsgBox(strText, intStyle, strCaption)
End Function
'uses CurDll
Function dbInputBox(ByVal strMessage As Variant, _
                    Optional ByVal Default As String = "", _
                    Optional ByVal CancelError As Boolean = False, _
                    Optional ByVal ShowInTaskBar As Boolean = False, _
                    Optional ByVal OwnerForm As Boolean, _
                    Optional ByVal BrowseButton As Boolean = False, _
                    Optional ByVal MaxLen As Long = 0, _
                    Optional ByVal MinLength As Long = 1) As String
Dim strText As String, strCaption As String, intPos As Long
If IsNumeric(strMessage) Then
    strMessage = grs(strMessage)
ElseIf Mid(strMessage, 1, 1) = "$" Then
    strMessage = grs(CInt(Mid(strMessage, 2, Len(strMessage) - 1)))
End If
intPos = InStr(strMessage, "`")
If intPos = 0 Then
    strText = strMessage
    strCaption = ""
Else
    strText = Mid$(strMessage, 1, intPos - 1)
    If Not intPos = Len(strMessage) Then
        strCaption = Mid$(strMessage, intPos + 1, Len(strMessage) - intPos)
    Else
        strCaption = ""
    End If
End If
With frmInput
    .SetProps strText, strCaption, Default, BrowseButton, MaxLen
    If ShowInTaskBar Then ShowTask strCaption
    If OwnerForm Then
        .Show vbModal, MainForm
    Else
        .Show vbModal
    End If
    If ShowInTaskBar Then HideTask
    If .Tag = "" And Len(.Text) >= MinLength Then
        dbInputBox = .Text
    Else
        dbInputBox = ""
        If CancelError Then Err.Raise dbCWS, "dbInputBox", "Cancel was selected"
    End If
End With
End Function

Public Sub LoadBrushFileByte(ByVal File As String, ByRef bData() As Byte)
Dim w As Byte, h As Byte, i As Long, j As Long, l As Long, Bytes() As Byte, nmb As Long, II As Long
nmb = FreeFile
Open File For Binary As nmb
    l = LOF(nmb)
    ReDim Bytes(0 To l - 1)
    Get nmb, 1, Bytes
Close nmb
w = Bytes(0)
h = Bytes(1)
II = 2
ReDim bData(0 To w - 1, 0 To h - 1)
For i = 0 To h - 1
    For j = 0 To w - 1
        bData(j, i) = (Bytes(II))
        II = II + 1
    Next j
Next i
End Sub

Public Sub LoadBrushFile(ByVal File As String, ByRef bData() As Byte)
Dim w As Byte, h As Byte, i As Long, j As Long, l As Long, Bytes() As Byte, nmb As Long, II As Long
nmb = FreeFile
Open File For Binary As nmb
    l = LOF(nmb)
    ReDim Bytes(0 To l - 1)
    Get nmb, 1, Bytes
Close nmb
w = Bytes(0)
h = Bytes(1)
II = 2
ReDim bData(0 To w - 1, 0 To h - 1)
For i = 0 To h - 1
    For j = 0 To w - 1
        bData(j, i) = (Bytes(II))
        II = II + 1
    Next j
Next i
End Sub

Sub SaveBrushToFile(ByRef Brush() As Byte, ByVal File As String)
Dim nmb As Long, i As Long, j As Long, h As Long, Bytes() As Byte, Wdt As Byte, hgt As Long
    Wdt = UBound(Brush, 1) + 1
    hgt = UBound(Brush, 2) + 1
    ReDim Bytes(0 To CLng(Wdt) * CLng(hgt) + 2& - 1&)
    Bytes(0) = Wdt
    Bytes(1) = hgt
    h = 2
    For i = 0 To UBound(Brush, 2)
        For j = 0 To UBound(Brush, 1)
            Bytes(h) = CByte(Abs(Brush(j, i)))
            h = h + 1
        Next j
    Next i
    If Not (StartWrite(File)) Then
        dbMsgBox 1173, vbCritical
        Err.Raise dbCWS, "Cancel was selected"
    End If
    nmb = FreeFile
    Open File For Binary As nmb
        Put nmb, 1, Bytes
    Close nmb
End Sub

Sub Main()
Dim tmp As String, tmp1 As String
Dim Tried As Boolean
Dim tErr As New ErrObject

'Dim cl As New clsIniFile
'cl.TestMe: End
TestForSetPixelBug

HookEnabled = True

Pi = Atn(1) * 4
ReDim PrgBuffer(0 To 1024)
DrawingEngine.AntiAliasingSharpness = 1#

MoveTimerRes = 8

Set CDl = New CommonDlg
Set gReg = New Reg

CDl.CancelError = True
ExePath = AppPath + App.EXEName + ".exe"
FirstLoad = Not CBool(dbGetSetting("Setup", "WasRun", "False")) Or (UCase$(Trim$(Command)) = "/INSTALL")
Set gBackPicture = LoadResPicture(RES_BackPicture, vbResBitmap)
On Error GoTo 0
MainForm.Show
ExeCount = Val(dbGetSetting("Special", "ExecCount", "0", True, True))
Exit Sub
End Sub

Sub dbLongMsgBox(ByRef strText As String, ByVal strCaption As String)
Dim OldCaption As String, OldLocked As Boolean
With Bytes
    OldCaption = .Caption
    OldLocked = .Text.Locked
    .Text.Locked = True
    .Caption = strCaption
    .Text.Text = strText
    .Show vbModal
    .Text.Locked = OldLocked
    .Caption = OldCaption
End With
End Sub

'Takes a color in RGB format. Fills the RGBQuad correctly.
Public Sub GetRgbQuadEx(ByVal clr As Long, ByRef Res As RGBQUAD)
Res.rgbRed = clr And &HFF&
Res.rgbGreen = (clr And &HFF00&) \ &H100&
Res.rgbBlue = clr \ &H10000
'Res.rgbReserved = &H0
End Sub

Public Sub GetRgbQuadLongEx(ByVal clr As Long, ByRef Res As RGBQuadLong)
Res.rgbRed = clr And &HFF&
Res.rgbGreen = (clr And &HFF00&) \ &H100&
Res.rgbBlue = (clr And &HFF0000) \ &H10000
'Res.rgbReserved = &H0
End Sub

Public Sub GetRgbQuadLongEx2(ByVal clr As Long, ByRef r As Long, ByRef g As Long, ByRef b As Long)
r = clr And &HFF&
g = (clr And &HFF00&) \ &H100&
b = (clr And &HFF0000) \ &H10000
'Res.rgbReserved = &H0
End Sub

Public Sub GetRGBQuadFloatEx(ByRef clr As Long, ByRef Res As RGBTriCurr)
Res.rgbRed = CSng(clr And &HFF&)
Res.rgbGreen = CSng((clr And &HFF00&) \ &H100&)
Res.rgbBlue = CSng(clr \ &H10000)
End Sub


Public Function CompareColorsRGB(ByRef rgb1 As RGBQUAD, ByRef rgb2 As RGBQUAD) As Long
CompareColorsRGB = Abs(rgb1.rgbBlue - rgb2.rgbBlue) + Abs(rgb1.rgbGreen - rgb2.rgbGreen) + Abs(rgb1.rgbBlue - rgb2.rgbBlue)
End Function


'returns an error if fails
Public Function StartWrite(ByVal File As String) As Boolean 'truncates the file.
Dim nmb As Long
nmb = FreeFile
Open File For Output As nmb
Close nmb
StartWrite = True
End Function

Function ContinueWrite(ByVal File As String) As Boolean
Dim nmb As Long
nmb = FreeFile
Open File For Binary Access Read Write As nmb 'check file is good
Close nmb
ContinueWrite = True
End Function

Sub ShowStatus(ByVal strMessage As Variant, Optional ByVal Obsolete As Integer = 1, Optional ByVal HoldTime As Integer = 0)
'Static Levels(1 To 4) As String
Static TimeExpiration As Long
Dim i As Integer, tmp As String
If Len(strMessage) = 0 Then Exit Sub
If IsNumeric(strMessage) Then
    strMessage = GRSF(CInt(strMessage))
End If
If Mid(strMessage, 1, 1) = "$" Then
    strMessage = GRSF(Val(Mid(strMessage, 2, Len(strMessage) - 1)))
End If
tmp = strMessage

If Abs(CCur(GetTickCount) - TimeExpiration) > 100000 Then
    TimeExpiration = GetTickCount
End If

If HoldTime = 0 Then
    If GetTickCount > TimeExpiration Then
        MainForm.Status.Caption = tmp
        MainForm.Status.Refresh
    End If
Else
    MainForm.Status.Caption = tmp
    MainForm.Status.Refresh
    TimeExpiration = IIf(CCur(GetTickCount) + HoldTime * 1000 > 2147483647, 2147483647, CCur(GetTickCount) + HoldTime * 1000)
End If

End Sub

Function BoolToStr_OnOff(ByVal b As Boolean) As String
If b Then
    BoolToStr_OnOff = GRSF(1241)
Else
    BoolToStr_OnOff = GRSF(1242)
End If
End Function

Public Function IsKey(KeyCode As KeyCodes) As Boolean
IsKey = (Abs(GetKeyState(KeyCode)) > 1)
End Function

Public Function GetShiftState() As dbShiftConstants
GetShiftState = IIf(IsKey(dbShift), dbStateShift, 0) Or _
                IIf(IsKey(dbCtrl), dbStateCtrl, 0) Or _
                IIf(IsKey(dbAlt), dbStateAlt, 0)
End Function

Public Function IsKeyToggled(KeyCode As KeyCodes)
IsKeyToggled = CBool(GetKeyState(KeyCode) And &H1)
End Function

Public Function VedNull(ByVal Number As Long, ByVal DigCount As Long) As String
VedNull = String$(DigCount - Len(CStr(Number)), "0") + CStr(Number)
End Function

Public Function VedNullStr(ByVal Number As String, ByVal DigCount As Long) As String
VedNullStr = String$(DigCount - Len(Number), "0") + Number
End Function

Function Max(ByVal a As Variant, ByVal b As Variant) As Variant
If a > b Then Max = a Else Max = b
End Function

Function Min(ByVal a As Variant, ByVal b As Variant) As Variant
If a < b Then Min = a Else Min = b
End Function

Function MinMany(ParamArray Vals() As Variant) As Variant
Dim Var As Variant
Dim Result As Variant
Result = Vals(0)
For Each Var In Vals
    If Var < Result Then
        Result = Var
    End If
Next
MinMany = Result
End Function

Function MaxMany(ParamArray Vals() As Variant) As Variant
Dim Var As Variant
Dim Result As Variant
Result = Vals(0)
For Each Var In Vals
    If Var > Result Then
        Result = Var
    End If
Next
MaxMany = Result
End Function

Function MaxD(ByVal a As Double, ByVal b As Double) As Double
If a > b Then MaxD = a Else MaxD = b
End Function

Function MinD(ByVal a As Double, ByVal b As Double) As Double
If a < b Then MinD = a Else MinD = b
End Function


Public Sub HideTask()
frmTB.Hide
End Sub

Public Sub ShowTask(ByVal strCaption As String)
With frmTB
    .Caption = strCaption
    .Width = 1
    .Height = 1
    .Visible = True
End With
End Sub

Public Sub SaveGraphToFile(ByRef Gph As dbGraph, File As String)
Const ID_String = "IOGRAPH"
Dim n As Integer
Dim St() As Byte, nmb As Long
Dim i As Long
With Gph
    n = UBound(.Points)
    ReDim St(0 To Len(ID_String) - 1)
    For i = 0 To Len(ID_String) - 1
        St(i) = Asc(Mid$(ID_String, 1 + i, 1))
    Next i
    If Not (StartWrite(File)) Then
        dbMsgBox 1173, vbCritical
        Err.Raise dbCWS, "SaveGraphToFile"
    End If
    nmb = FreeFile
    Open File For Binary As nmb
        Put nmb, 1, St
        Put nmb, , n
        Put nmb, , .Points
        Put nmb, , .InterpolationMode
    Close nmb
End With
End Sub

Public Sub LoadGraphFile(ByVal File As String, ByRef Gph As dbGraph)
Const ID_String = "IOGRAPH"
Dim nmb As Long, n As Integer, St() As Byte, tmp As String
Dim i As Long
nmb = FreeFile
ReDim St(0 To Len(ID_String) - 1)
Open File For Binary As nmb
    Get nmb, 1, St
    tmp = Space$(Len(ID_String))
    For i = 0 To Len(ID_String) - 1
        Mid(tmp, 1 + i, 1) = Chr$(St(i))
    Next i
    If tmp <> ID_String Then
        dbMsgBox 1189, vbCritical
        Reset
        Err.Raise dbCWS, "LoadGraphFile"
    End If
    With Gph
        Get nmb, , n
        ReDim .Points(0 To n)
        Get nmb, , .Points
        Get nmb, , .InterpolationMode
        If .InterpolationMode < 0 Or .InterpolationMode > 1 Then
            .InterpolationMode = 0
        End If
        .NeedsInterpolation = True
    End With
Close nmb
End Sub

Public Function CompareColorsLng(ByVal LngColor1 As Long, ByVal LngColor2 As Long) As Long
Dim tmp1 As RGBQuadLong, tmp2 As RGBQuadLong
If LngColor1 = LngColor2 Then
    CompareColorsLng = 0
    Exit Function
End If
GetRgbQuadLongEx LngColor1, tmp1
GetRgbQuadLongEx LngColor2, tmp2
CompareColorsLng = Abs(tmp1.rgbRed - tmp2.rgbRed) + _
                   Abs(tmp1.rgbGreen - tmp2.rgbGreen) + _
                   Abs(tmp1.rgbBlue - tmp2.rgbBlue)
End Function

Public Sub EditPicture(ByRef bData() As Long)
Dim strFile As String, i As Long
On Error GoTo eh
CreateFolder TempPath
Do
    strFile = ValFolder(TempPath) + "SelTrans" + VedNull(i, 4) + ".smb"
    i = i + 1
Loop Until Not FileExists(strFile)
On Error GoTo eh2
SaveSMB bData, strFile
ExecuteAppModally ExePath + " """ + strFile + """"
LoadSMB bData, strFile
On Error GoTo eh
Kill strFile
Exit Sub
kf:
On Error Resume Next
Kill strFile
On Error GoTo 0
PopError
ErrRaise
Exit Sub
eh:
ErrRaise "EditPicture"

eh2:
If Err.Number = errNewFile Then
    Erase bData
    Resume Next
End If
PushError
Resume kf
End Sub

Public Sub ExecuteAppModally(ByRef Comm As String)
Dim hInst As Long
hInst = Shell(Comm, vbNormalFocus)
On Error GoTo eh
Do
    WaitMessage
    DoEvents
    If GetActiveWindow <> 0 Then
        AppActivate hInst
    End If
Loop
eh:
End Sub

'decimal separator is always "."
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
    Case Else
        Err.Raise 118, "dbCStr", "This expression cannot be converted to a string"
End Select
End Function

Public Function GRSF(ByVal ResID As Long, Optional ByVal RaiseErrors As Boolean = False) As String
If Not RaiseErrors Then
    On Error GoTo eh
End If
GRSF = LoadResString(ResID)
Exit Function
Resume
eh:
Debug.Assert False
MsgBox "Cannot load resid " + CStr(ResID) + ". Function GRSF. Please mail to VT-Dbnz@yandex.ru about this error. Maybe you can fix it manually by loading SMBMaker.exe into an resource editor and creating a string with this ID."
GRSF = "<caption not found: resid=" + CStr(ResID) + ">"
End Function

Public Function GRSM(ByVal ResID As Long, ByVal Act As dbCommands, ByRef pKeys As clsShortcuts)
On Error GoTo eh
GRSM = ASTC(LoadResString(ResID), Act, pKeys)
Exit Function
Resume
eh:
Debug.Assert False
MsgBox "Cannot load resid " + CStr(ResID) + ". Function GRSM. Please mail to VT-Dbnz@yandex.ru about this error. Maybe you can fix it manually by loading SMBMaker.exe into an resource editor and creating a string with this ID."
GRSM = "<caption not found: resid=" + CStr(ResID) + ">"
End Function

Public Function WasKey(KeyCode As KeyCodes) As Boolean
Dim n As Long
n = KeyCode
WasKey = CBool(GetAsyncKeyState(KeyCode) And 1)
End Function

Public Sub ExtractUniqueColors(pData() As Long, pal() As Long)
Dim w As Long, h As Long
Dim i As Long, j As Long, LastColor As Long
Dim P As Long
On Error Resume Next
If AryDims(AryPtr(pData)) = 1 Then
    Err.Raise 1111
ElseIf AryDims(AryPtr(pData)) = 2 Then
    h = UBound(pData, 2)
Else
    Exit Sub
End If
On Error GoTo 0
w = UBound(pData, 1)
If Err.Number = 0 Then
    ReDim pal(0 To (w + 1) * (h + 1) - 1)
    For i = 0 To h
        P = (w + 1) * i
        For j = 0 To w
            pal(P + j) = pData(j, i)
        Next j
    Next i
Else
    pal = pData
End If

SortLongArray pal, 0, UBound(pal)

LastColor = pal(0)
P = 1

For i = 1 To UBound(pal)
    If pal(i) <> LastColor Then
        P = P + 1
        LastColor = pal(i)
        pal(P) = LastColor
    End If
Next i

ReDim Preserve pal(0 To P)
End Sub

Public Function dbAlphaBlend(ByVal LngColor1 As Long, ByVal LngColor2 As Long, ByVal Alpha As Long) As Long
Const r As Long = &HFF&
Const g As Long = &HFF00&
Const b As Long = &HFF0000
Dim r1 As Long, g1 As Long, b1 As Long
Dim r2 As Long, g2 As Long, b2 As Long
If Alpha = 0 Then
    dbAlphaBlend = LngColor1
    Exit Function
ElseIf Alpha = 255 Then
    dbAlphaBlend = LngColor2
    Exit Function
End If
r1 = LngColor1 And r
g1 = LngColor1 And g
b1 = LngColor1 And b
r2 = LngColor2 And r
g2 = LngColor2 And g
b2 = LngColor2 And b
dbAlphaBlend = ((r2 - r1) * Alpha \ 255 + r1) Or _
               ((g2 - g1) * Alpha \ 255 + g1) And g Or _
               (((b2 - b1) \ 255) * Alpha + b1) And b
End Function

Public Function dbAlphaBlendRGB(ByVal LngColor1 As Long, ByVal LngColor2 As Long, ByVal AlphaRGB As Long) As Long
Const r As Long = &HFF&
Const g As Long = &HFF00&
Const b As Long = &HFF0000
Dim r1 As Long, g1 As Long, b1 As Long
Dim r2 As Long, g2 As Long, b2 As Long
Dim r3 As Long, g3 As Long, b3 As Long
If AlphaRGB = 0 Then
    dbAlphaBlendRGB = LngColor1
    Exit Function
ElseIf AlphaRGB = &HFFFFFF Then
    dbAlphaBlendRGB = LngColor2
    Exit Function
End If
r1 = LngColor1 And r
g1 = LngColor1 And g
b1 = LngColor1 And b
r2 = LngColor2 And r
g2 = LngColor2 And g
b2 = LngColor2 And b
r3 = AlphaRGB And r
g3 = (AlphaRGB And g) \ 256
b3 = (AlphaRGB And b) \ 65536
dbAlphaBlendRGB = ((r2 - r1) * r3 \ 255 + r1) Or _
               ((g2 - g1) * g3 \ 255 + g1) And g Or _
               (((b2 - b1) \ 255) * b3 + b1) And b
End Function

Public Function GetKeyName(ByVal KeyCode As Long) As String
Dim KK As Long
Dim Shift As Integer
Dim tmp As String
Dim i As Long, j As Long, h As Long
Dim II As Long, jj As Long
Dim NoShift As Boolean
Static Names As Boolean
Static NamesArr() As String
Dim tmpArr() As String
Const ShiftShift = &H100
Const ShiftControl = &H200
Const ShiftAlt = &H400

KK = KeyCode And &HFF
Shift = KeyCode And &H700
NoShift = (KeyCode And &H800) <> 0

If Not NoShift Then
    If KeyCode And ShiftControl Then
        tmp = tmp + "Ctrl+"
    End If
    If KeyCode And ShiftAlt Then
        tmp = tmp + "Alt+"
    End If
    If KeyCode And ShiftShift Then
        tmp = tmp + "Shift+"
    End If
End If

If Not Names Then
    tmpArr = Split(LoadResString(1798), vbCrLf)
    ReDim NamesArr(0 To 0)
    
    For i = 0 To UBound(tmpArr)
        j = InStr(1, tmpArr(i), "=")
        II = Val(Left$(tmpArr(i), j - 1))
        If UBound(NamesArr) < II Then
            ReDim Preserve NamesArr(0 To II)
        End If
        NamesArr(II) = Trim(Mid$(tmpArr(i), j + 1))
    Next i
    Names = True
End If

If KK > UBound(NamesArr) Then
    tmp = tmp + "Key " + CStr(KK)
ElseIf Len(NamesArr(KK)) = 0 And KK <> 0 Then
    tmp = tmp + "Key " + CStr(KK)
Else
    tmp = tmp + NamesArr(KK)
End If

GetKeyName = IIf(NoShift, "[" + tmp + "]", tmp)

End Function

Public Function dbInKey(Optional ByVal RaiseErrors As Boolean = True) As Long
With frmInKey
    .Show vbModal
    If .Tag = "" Then
        Unload frmInKey
        If RaiseErrors Then
            Err.Raise dbCWS, "dbInKey", "Cancel Was Selected"
        Else
            dbInKey = -1
        End If
        Exit Function
    End If
    dbInKey = Val(.Tag)
End With
Unload frmInKey
End Function

Public Function GetActionDescription(ByVal cAct As dbCommands)
GetActionDescription = GRSF(1800 - 1 + cAct)
End Function

Public Sub SaveKeysToFile(ByRef File As String, ByRef sKeys() As kShortcut)
Dim tmp As String
Dim nmb As Long
Dim Bytes1() As Byte
Dim i As Long

'lSet sKeys1 = sKeys0
Const FILE_ID = "SMB_KEYS1"
'KeysToString sKeys, tmp
If FileExists(File) Then
    If CBool(GetAttr(File) And vbReadOnly) Then
        'dbMsgBox 1167, vbCritical '"Access denied"
        Err.Raise 1111, "SaveKeysToFile", "Access denied. The file is read-only, not writable."
    End If
End If

If Not (StartWrite(File)) Then
    'dbMsgBox 1173, vbCritical
    Err.Raise dbCWS, "SaveKeysToFile", "Cannot open the file for writing."
End If

nmb = FreeFile
Open File For Binary As nmb
    ReDim Bytes1(0 To Len(FILE_ID) - 1)
    For i = 0 To UBound(Bytes1)
        Bytes1(i) = Asc(Mid$(FILE_ID, i + 1, 1))
    Next i
    
    Put nmb, 1, Bytes1
    If AryDims(AryPtr(sKeys)) = 1 Then
        Put nmb, , CLng(UBound(sKeys) + 1)
        Put nmb, , sKeys
    Else
        Put nmb, , CLng(0&)
    End If
Close nmb
End Sub

Public Sub LoadKeysFromFile(ByRef File As String, ByRef sKeys() As kShortcut)
Dim nmb As Long
Dim Bytes1() As Byte
Dim n As Long
Dim i As Long
Const FILE_ID = "SMB_KEYS1"
ReDim Bytes1(0 To Len(FILE_ID) - 1)
nmb = FreeFile
If Not FileExists(File) Then
    'dbMsgBox "No File to Open", vbCritical
    Err.Raise 1111, "LoadKeysFromFile", "File not found: """ + File + """!"
End If
Open File For Binary As nmb
    Get nmb, 1, Bytes1
    For i = 0 To UBound(Bytes1)
        If Bytes1(i) <> Asc(Mid$(FILE_ID, i + 1, 1)) Then
            Reset
            'dbMsgBox "Not a Keys file", vbCritical
            Err.Raise 1111, "LoadKeysFromFile", "The file's format is incorrect."
        End If
    Next i
    Get nmb, , n
    If n < 1 Then
        Erase sKeys
    Else
        ReDim sKeys(0 To n - 1)
        Get nmb, , sKeys
    End If
Close nmb
End Sub

Public Sub SaveKeysToReg(ByRef sKeys() As kShortcut, Key As String, Parameter As String)
Dim tmp As String
KeysToString sKeys, tmp
dbSaveSetting Key, Parameter, tmp
End Sub

Public Sub LoadKeysFromReg(ByRef sKeys() As kShortcut, Key As String, Parameter As String)
'Dim i As Long, j As Long
Dim tmp As String
'Dim kData As String
If Len(Parameter) = 0 Then
    tmp = DefKeysString
Else
    tmp = dbGetSetting(Key, Parameter, DefKeysString)
End If
StringToKeys tmp, sKeys
End Sub

Public Sub KeysToString(ByRef sKeys() As kShortcut, ByRef kData As String)
Dim i As Long, j As Integer
Dim tmp As String
'Dim kData As String
If AryDims(AryPtr(sKeys)) <> 1 Then
    j = -1
Else
    j = UBound(sKeys)
End If
kData = Space$((j + 1) * 5 + 4)
Mid$(kData, 1, 4) = VedNullStr(Hex$(j), 4)
For i = 0 To j
    Mid$(kData, 1& + i * 5& + 4&, 3) = VedNullStr(Hex$(sKeys(i).Key), 3)
    Mid$(kData, 4& + i * 5& + 4&, 2) = VedNullStr(Hex$(sKeys(i).Act), 2)
Next i
End Sub

Public Sub StringToKeys(ByRef kStr As String, ByRef sKeys() As kShortcut)
Dim i As Long, j As Long
Dim tmp As String
Dim kData As String
kData = Mid$(kStr, 5)
tmp = Mid$(kStr, 1, 4)
j = CInt("&H" + tmp)
If Len(kData) < (j + 1) * 5 Then
    j = Len(kData) \ 5 - 1
    'Err.Raise 111, "LoadKeysFromReg", "Incorrect length of a string"
End If
If j < 0 Then
    Erase sKeys
Else
    ReDim sKeys(0 To j)
    For i = 0 To j
        sKeys(i).Key = CLng("&H" + Mid$(kData, 1 + i * 5, 3))
        sKeys(i).Act = CLng("&H" + Mid$(kData, 4 + i * 5, 2))
    Next i
End If
End Sub

Public Function KeyCodesEqual(ByVal KK1 As Long, ByVal KK2 As Long)
If ((KK1 Or KK2) And &H800) <> 0 Then
    KeyCodesEqual = (KK1 And &HFF) = (KK2 And &HFF)
Else
    KeyCodesEqual = KK1 = KK2
End If
End Function

'adds shotcuts to menu caption
Public Function ASTC(ByRef strCaption As String, _
                     ByVal Act As dbCommands, _
                     ByRef pKeys As clsShortcuts) As String
Dim j As Long, i As Long
Dim kArr() As String

j = pKeys.ListKeys(Act, kArr)
If j = 0 Then
    j = InStr(1, strCaption, Chr$(9))
    If j > 0 Then
        ASTC = Left$(strCaption, j - 1)
    Else
        ASTC = strCaption
    End If
Else
    'i = pKeys(i).Key
    j = InStr(1, strCaption, Chr$(9))
    If j > 0 Then
        ASTC = Left$(strCaption, j) + Join(kArr, ", ")
    Else
        ASTC = strCaption + Chr$(9) + Join(kArr, ", ")
    End If
End If
End Function

'returns the number of keys found
'sArr() will contain their captions
Public Function ListKeys(ByVal Act As dbCommands, _
                         ByRef pKeys() As kShortcut, _
                         ByRef sArr() As String) As Long
Dim j As Long, i As Long
If AryDims(AryPtr(pKeys)) <> 1 Then Exit Function
j = 0
ReDim sArr(0 To 0)
For i = 0 To UBound(pKeys)
    If pKeys(i).Act = Act Then
        ReDim Preserve sArr(0 To j)
        sArr(j) = GetKeyName(pKeys(i).Key)
        j = j + 1
    End If
Next i
If j = 0 Then Erase sArr
ListKeys = j
End Function

Public Function dbWindowFromPoint(Pnt As POINTAPI) As Long
dbWindowFromPoint = WindowFromPoint(Pnt.x, Pnt.y)
End Function

Public Sub GetDesktopRect(ByRef rct As RECT)
Dim hWndDesktop As Long
hWndDesktop = GetDesktopWindow
GetWindowRect hWndDesktop, rct
End Sub

Public Sub CaptureWindow(ByRef Data() As Long, _
                         ByVal hWnd As Long)
Dim htmpDC As Long
Dim hDefObj As Long
Dim hBmp As Long
Dim hDC As Long
Dim rct As RECT
Dim Rct1 As RECT
Dim w As Long, h As Long
On Error GoTo eh
If hWnd <> 0 Then
    GetWindowRect hWnd, rct 'Get dimensions of a window
    GetDesktopRect Rct1
    rct = IntersectRects(rct, Rct1)
End If
h = rct.Bottom - rct.Top '+ 1
w = rct.Right - rct.Left '+ 1
hDC = GetWindowDC(hWnd) '
If hDC = 0 Then Err.Raise 111, "CaptureWindow", "Cannot get device context"
If hWnd = 0 Then
    w = GetDeviceCaps(hDC, HORZRES)
    h = GetDeviceCaps(hDC, VERTRES)
End If
hBmp = CreateCompatibleBitmap(hDC, w, h)
htmpDC = CreateCompatibleDC(hDC) 'create temporary DC
hDefObj = SelectObject(htmpDC, hBmp) 'Select bitmap to the tmp DC
BitBlt htmpDC, 0, 0, w, h, hDC, 0, 0, SRCCOPY 'copy image from window dc to tmp dc
hBmp = SelectObject(htmpDC, hDefObj) 'restore def obj for tmp DC
On Error Resume Next
dbGetDIBits hBmp, htmpDC, Data 'Get picture data
On Error GoTo 0
DeleteDC htmpDC 'Delete tmp DC
DeleteObject hBmp 'Delete Bitmap
ReleaseDC hWnd, hDC 'Release window DC
Exit Sub
eh:
DeleteDC htmpDC 'Delete tmp DC
DeleteObject hBmp 'Delete Bitmap
ReleaseDC hWnd, hDC 'Release window DC
ErrRaise "CaptureWindow"
End Sub

Public Function CaptureZoomIn2(ByRef AroundPoint As POINTAPI, _
                               ByVal DestHalfW As Long, _
                               ByVal DestHalfH As Long, _
                               ByVal Zm As Long, _
                               ByVal hDCOut As Long, _
                               ByVal OutCX As Long, _
                               ByVal OutCY As Long)
Dim DestX As Long, DestY As Long
Dim SrcX As Long, SrcY As Long
Dim DestW As Long, DestH As Long
Dim srcW As Long, srcH As Long
Dim ScreenDC As Long
srcW = -Int(-DestHalfW / Zm - 0.5)
srcH = -Int(-DestHalfH / Zm - 0.5)
DestHalfW = srcW * Zm
DestHalfH = srcH * Zm
SrcX = AroundPoint.x - srcW
SrcY = AroundPoint.y - srcH
DestX = OutCX - DestHalfW - Zm \ 2&
DestY = OutCY - DestHalfH - Zm \ 2&
On Error GoTo eh
ScreenDC = GetWindowDC(0&)
StretchBlt hDCOut, _
           DestX, DestY, _
           (srcW * 2 + 1) * Zm, (srcH * 2 + 1) * Zm, _
           ScreenDC, _
           SrcX, SrcY, _
           srcW * 2 + 1, srcH * 2 + 1, _
           SRCCOPY
ReleaseDC 0&, ScreenDC
Exit Function
eh:
ReleaseDC 0&, ScreenDC
ErrRaise "CaptureZoomIn2"
End Function

Public Function CapturePixel(ByVal px As Long, ByVal py As Long) As Long
Dim WDC As Long
WDC = GetWindowDC(0&)
CapturePixel = ConvertColorLng(GetPixel(WDC, px, py))
ReleaseDC 0&, WDC
End Function

Public Sub dbCapture(ByVal cMode As Long, ByRef cData() As Long)
Dim cHWnd As Long
Dim Pnt As POINTAPI
Dim PrevWS As FormWindowStateConstants
Dim kTgl As Boolean
Select Case cMode
    Case 0 'Window from point
        MainForm.MeEnabled = False
        PrevWS = MainForm.WindowState
        MainForm.WindowState = vbMinimized
        
        'WasKey 145 'Pause
        WasKey 27 'Esc
        DoEvents
        kTgl = IsKeyToggled(145)
        Do
            DoEvents
            If WasKey(27) Then
                MainForm.WindowState = PrevWS
                MainForm.MeEnabled = True
                Err.Raise dbCWS
            End If
        Loop While IsKeyToggled(145) = kTgl 'Pause
        GetCursorPos Pnt
        cHWnd = dbWindowFromPoint(Pnt)
        CaptureWindow cData, cHWnd
        MainForm.WindowState = PrevWS
        MainForm.MeEnabled = True
    Case 1 'Active window
        'WasKey 27
        ShowStatus 2424, , 5
        BreakKeyPressed
        MainForm.WaitW 0, True, True
        MainForm.WaitW 5000, False, True, True
        'If WasKey(27) Then Err.Raise dbCWS
        cHWnd = GetForegroundWindow
        CaptureWindow cData, cHWnd
    Case 2 'Entire screen
        ShowStatus 2425, , 5
        'WasKey 27
        BreakKeyPressed
        MainForm.WaitW 0, True, True
        MainForm.WaitW 5000, False, True, True
        'If WasKey(27) Then Err.Raise dbCWS
        cHWnd = 0
        CaptureWindow cData, cHWnd
End Select
End Sub

Public Function GetColorName(ByVal bgrColor As Long) As String
Const StartResID = 1700
Dim tmp As String
Dim i As Long
Dim Res As String
Dim pos As Long
Dim clr As Long
Dim rgb1 As RGBQUAD

i = StartResID
On Error Resume Next
Err.Clear
Do
    tmp = GRSF(i, RaiseErrors:=True)
    If Err.Number = 0 Then
        pos = InStrRev(tmp, "|")
        If pos > 0 Then
            clr = WebColorToLong(Mid$(tmp, pos + 1))
            If clr = bgrColor Then
                Res = Left$(tmp, pos - 1)
                Exit Do
            End If
        End If
    End If
    i = i + 1&
Loop Until tmp = "EOL" Or Err.Number <> 0
On Error GoTo 0
If Len(Res) = 0 Then
    CopyMemory rgb1, bgrColor, 4
    If rgb1.rgbRed = rgb1.rgbGreen And rgb1.rgbGreen = rgb1.rgbBlue Then
        Res = grs(1797, "<n>", CStr(rgb1.rgbRed), _
                        "<d>", CStr((255 - rgb1.rgbRed) * 100 \ 254)) _
             '"Gray " + CStr(rgb1.rgbRed) + " (" + CStr((255 - rgb1.rgbRed) * 100 \ 254) + "%)"
    End If
End If
GetColorName = Res
End Function

'bgrColor should be in BGR format.
Public Function GenerateColorTip(ByVal bgrColor As Long) As String
Dim tmp As String, rgb1 As RGBQUAD
'Now the color is in RGB format
CopyMemory rgb1, bgrColor, 4
tmp = GetColorName(bgrColor)
If Len(tmp) > 0 Then
    tmp = tmp + ". "
End If
tmp = tmp + grs(2284, "%r", CStr(rgb1.rgbRed), _
                      "%g", CStr(rgb1.rgbGreen), _
                      "%b", CStr(rgb1.rgbBlue), _
                      "%h", LongToWebColor(bgrColor)) _
           '"RGB: R = " + CStr(rgb1.rgbRed) + ", " + _
                 "G = " + CStr(rgb1.rgbGreen) + ", " + _
                 "B = " + CStr(rgb1.rgbBlue) + ". " + _
                 "Hex code: " + VedNullStr(Hex$(bgrColor), 6) + "."
                 
GenerateColorTip = tmp

End Function

Public Function LongToWebColor(ByVal bgrColor As Long) As String
Dim rgb1 As RGBQUAD
CopyMemory rgb1, bgrColor, 4
LongToWebColor = RGBToWebColor(rgb1.rgbRed, rgb1.rgbGreen, rgb1.rgbBlue)
End Function

Public Function RGBToWebColor(ByVal r As Long, _
                              ByVal g As Long, _
                              ByVal b As Long) As String
RGBToWebColor = "#" + VedNullStr(Hex$(r), 2) + _
                      VedNullStr(Hex$(g), 2) + _
                      VedNullStr(Hex$(b), 2)
End Function

Public Function WebColorToLong(ByRef St As String) As Long 'bgr color
Dim tmp As String
Dim clr As Long
Dim r As Long, g As Long, b As Long
Dim i As Long
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
clr = BGR(r, g, b)
WebColorToLong = clr
End Function

Public Sub InitHiResTimer()
If TimerInitialized Then Exit Sub
If timeBeginPeriod(1&) <> 0 Then
    Err.Raise 111, "InitHiResTimer", "Cannot initialize multimedia timer."
End If
TimerInitialized = True
End Sub

Public Sub DestroyHiResTimer()
If Not TimerInitialized Then Exit Sub
timeEndPeriod 1&
TimerInitialized = False
End Sub

Public Function mGetTickCount() As Long
mGetTickCount = timeGetTime
End Function

Public Sub SaveWindowPos(ByRef frm As Form)
Dim sArr() As String
Dim LWS As Long
ReDim sArr(0 To 4)
sArr(0) = CStr(frm.WindowState)
LWS = frm.WindowState
frm.WindowState = 0
sArr(1) = CStr(frm.Left)
sArr(2) = CStr(frm.Top)
sArr(3) = CStr(frm.Width)
sArr(4) = CStr(frm.Height)
frm.WindowState = LWS
dbSaveSetting "Windows", frm.Name, Join(sArr, ",")
End Sub

Public Sub LoadWindowPos(ByRef frm As Form, _
                         Optional ByVal OnlyWH As Boolean = True)
Dim sArr() As String
sArr = Split(dbGetSetting("Windows", frm.Name), ",")
If UBound(sArr) <> 4 Then Exit Sub
Dim l As Long, t As Long
Dim w As Long, h As Long
frm.WindowState = vbNormal
l = IIf(OnlyWH, frm.Left, Val(sArr(1)))
t = IIf(OnlyWH, frm.Top, Val(sArr(2)))
w = Val(sArr(3))
h = Val(sArr(4))
If l < 0 Then l = 0
If t < 0 Then t = 0
If l + w > Screen.Width Or t + h > Screen.Height Then
  l = 0
  t = 0
End If
frm.Move l, t, w, h
frm.WindowState = Val(sArr(0))
End Sub

Public Function GetShiftKeyCode(ByVal KeyCode As Long, ByVal Shift As Long) As Long
GetShiftKeyCode = (&H100& * Shift) Or KeyCode
End Function

Public Sub ShowHelp(ByVal TopicsResID As Integer)
CurHelpResId = TopicsResID
If HelpWindowVisible Then
    frmHelp.LoadHelp CurHelpResId
    If Not frmHelp.bLockTopic Then frmHelp.btnTopic_Click (0)
    EnableWindow frmHelp.hWnd, True
End If
End Sub


Public Sub ShowHelpWindow()
Const HWPortion As Single = 0.3333
Dim HWW As Long
Dim MFL As Long
HelpWindowVisible = True
Load frmHelp
frmHelp.bLockTopic = False
HWW = Screen.Width * HWPortion
MainForm.WindowState = vbNormal
If MainForm.Left + MainForm.Width + frmHelp.Width > Screen.Width Then
    MFL = Screen.Width - HWW - MainForm.Width
    If MFL < 0 Then
        MFL = 0
        MainForm.Move 0, MainForm.Top, Screen.Width - frmHelp.Width
    Else
        MainForm.Move MFL
    End If
    
End If
With frmHelp
    .Move MainForm.Left + MainForm.Width, MainForm.Top, HWW, MainForm.Height
    If CurHelpResId <> 0 Then
        .LoadHelp CurHelpResId
        .btnTopic_Click 0
    End If
    '.Show
    ShowWindow frmHelp.hWnd, SW_SHOWNA
End With
End Sub

Public Sub HideHelpWindow()
If HelpWindowVisible Then
    'frmHelp.Hide
    frmHelp.bLockTopic = False
    ShowWindow frmHelp.hWnd, SW_HIDE
    Unload frmHelp
End If
HelpWindowVisible = False
End Sub

Public Sub ToggleHelpWindow()
If HelpWindowVisible Then
    HideHelpWindow
Else
    ShowHelpWindow
End If
End Sub

Public Sub ShowFormModal(ByRef frm As Form)
Dim CurHID As Long
If frm.HelpContextID <> 0 Then
    CurHID = CurHelpResId
    ShowHelp frm.HelpContextID
End If
frm.Show vbModal
If frm.HelpContextID <> 0 Then
    ShowHelp CurHID
End If
End Sub

Public Sub MoveMouse(Optional ByVal dx As Long, _
                     Optional ByVal dy As Long, _
                     Optional ByVal Immediate As Boolean = False)
Dim a As POINTAPI
If Immediate Then
    'MoveMouseDx = 0
    'MoveMouseDy = 0
    GetCursorPos a
    SetCursorPos a.x + dx, a.y + dy
Else
    MoveMouseDx = dx
    MoveMouseDy = dy
    MainForm.tmrMoveMouser.Enabled = True
End If
End Sub

Public Sub SelTextInTextBox(ByRef TB As TextBox)
TB.SelStart = 0
TB.SelLength = Len(TB.Text)
End Sub


Public Sub SaveMatrix(ByRef Matrix() As Double, ByRef Key As String, ByRef Parameter As String)
Dim i As Long, j As Long
Dim strAry() As String
ReDim strAry(0 To (UBound(Matrix, 1) + 1) * (UBound(Matrix, 2) + 1) - 1)
For i = 0 To UBound(Matrix, 1)
    For j = 0 To UBound(Matrix, 2)
        strAry(i * (UBound(Matrix, 2) + 1) + j) = dbCStr(Matrix(i, j))
    Next j
Next i
dbSaveSetting Key, Parameter, Join(strAry, ";")
End Sub

'dimension of the matrix are the same as of the input one.
' it is not obtained from the registry.
Public Sub LoadMatrix(ByRef Matrix() As Double, ByRef Key As String, ByRef Parameter As String, Optional ByRef DefString As String = "1;0;0;0;1;0;0;0;1")
Dim i As Long, j As Long
Dim tmp As String, strAry() As String
On Error GoTo eh
tmp = dbGetSetting(Key, Parameter, "")
strAry = Split(tmp, ";")
If UBound(strAry) + 1 <> (UBound(Matrix, 1) + 1) * (UBound(Matrix, 2) + 1) Then
    Err.Raise 1111, "LoadMatrix", "Not enough numbers to fill the matrix."
End If
For i = 0 To UBound(Matrix, 1)
    For j = 0 To UBound(Matrix, 2)
        Matrix(i, j) = dbVal(strAry(i * (UBound(Matrix, 2) + 1) + j))
    Next j
Next i
'UpdateTexts
eh:
End Sub

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
            Result = CDbl(Value)
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
        If Result > nMax Then
            Err.Raise 119, "dbVal", "Limit exceeded. Maximum is " + CStr(nMax) + "."
        End If
    End If
    dbVal = Result
End Function


Public Sub ShowLensWindow()
Dim HWW As Long
Dim MFL As Long
LensWindowVisible = True
Load frmLens
With frmLens
    LoadWindowPos frmLens
    ShowWindow .hWnd, SW_SHOWNA
    ReposFrmLens
End With
MainForm.PctCapture.Cls
End Sub

Public Sub ReposFrmLens(Optional ByVal DisableEnable As Long)
Static Enabled As Boolean
Static Initd As Boolean
If Not Initd Then
    Initd = True
    Enabled = True
End If
If DisableEnable = -1 Then
    Enabled = False
ElseIf DisableEnable = 1 Then
    Enabled = True
End If
If Not Enabled Then Exit Sub
If LensWindowVisible Then
    SetWindowPos frmLens.hWnd, HWND_TOPMOST, 0, 0, 0, 0, _
                 SWP_NOMOVE Or _
                 SWP_NOSIZE Or _
                 SWP_NOACTIVATE Or _
                 SWP_NOOWNERZORDER
    EnableWindow frmLens.hWnd, True
End If
End Sub

Public Sub HideLensWindow()
If LensWindowVisible Then
    'frmHelp.Hide
    ShowWindow frmLens.hWnd, SW_HIDE
    SaveWindowPos frmLens
    Unload frmLens
End If
LensWindowVisible = False
End Sub

Public Sub ToggleLensWindow()
If LensWindowVisible Then
    HideLensWindow
Else
    ShowLensWindow
End If
End Sub



Public Sub ShowToolTipWindow()
Dim HWW As Long
Dim MFL As Long
ToolTipWindowVisible = True
Load frmToolTip
With frmToolTip
    ShowWindow .hWnd, SW_SHOWNOACTIVATE
    SetWindowPos .hWnd, HWND_TOPMOST, 0, 0, 0, 0, _
                 SWP_NOMOVE Or _
                 SWP_NOSIZE Or _
                 SWP_NOOWNERZORDER Or _
                 SWP_NOACTIVATE
    
End With
End Sub

Public Sub HideToolTipWindow()
If ToolTipWindowVisible Then
    'frmHelp.Hide
    ShowWindow frmToolTip.hWnd, SW_HIDE
    Unload frmToolTip
End If
ToolTipWindowVisible = False
End Sub

Public Sub ToggleToolTipWindow()
If ToolTipWindowVisible Then
    HideToolTipWindow
Else
    ShowToolTipWindow
End If
End Sub

Public Sub ShowToolTip(ByVal hWnd As Long, ByRef Message As Variant, Optional ByVal msHoldTime As Long)
Dim rct As RECT
If IsNumeric(Message) Then
    Message = GRSF(CInt(Message))
End If
If Len(Message) > 1 Then
    If Left$(Message, 1) = "$" Then
        Message = GRSF(CInt(Mid$(Message, 2)))
    End If
End If
If Not ToolTipWindowVisible Then
    Load frmToolTip
End If
With frmToolTip
    .SetText CStr(Message)
    GetWindowRect hWnd, rct
    .CalcPos rct
    .UpdatePos
End With
If Not ToolTipWindowVisible Then
    ShowToolTipWindow
End If
frmToolTip.Timer1.Interval = msHoldTime
frmToolTip.Timer1.Enabled = msHoldTime > 0

End Sub


Public Function MainFormHook(ByVal HookCode As PM, ByVal wParam As PM, ByRef Message As Msg) As Long
Dim Answ As VbMsgBoxResult
Dim SS As dbShiftConstants
Dim GM As Long
Dim ShiftKeyCode As Long
Dim i As Long
Dim nActs As Long
Dim ActsAry() As Long
Static Rec As Boolean
Static Counter As Long
If Not HookEnabled Then
    'this is made for handling visual basic's stop button
    'or unexpected program termination
    UninstallHook ReadLastHookHandle:=True
    Exit Function
End If
On Error GoTo eh
If Rec Then Exit Function
Rec = True
SS = GetShiftState
Select Case Message.Message
    Case WM.WM_LBUTTONDOWN, _
         WM.WM_MBUTTONDOWN, _
         WM.WM_RBUTTONDOWN
         
        HideToolTipWindow
End Select

If wParam = PM_REMOVE Then
    Select Case Message.Message
        Case WM.WM_Wheel
            If IsChild(MainForm.hWnd, Message.hWnd) Then
                MainForm.Form_Wheel ((Message.wParam And &HFFFF0000) \ &H10000) / CDbl(WHEEL_DELTA), SS
                'message processed - zero out the movement, otherwise subclassed proc will also detect it
                Message.wParam = Message.wParam And Not &HFFFF0000
            ElseIf HelpWindowVisible Then
                If IsChild(frmHelp.hWnd, Message.hWnd) Then
                    frmHelp.Scroll -((Message.wParam And &HFFFF0000) \ &H10000) \ 2
                    Message.wParam = Message.wParam And Not &HFFFF0000
                End If
            End If
        Case WM.WM_MOUSEFIRST To WM.WM_MOUSELAST
            GM = GetMessageExtraInfo
            If (GM And &HFF0000) = &H430000 Then
                With MainForm
                    GM = GM And &H1FF
                        .PenPressure = GM
                    If .MaxPenPressure = -1 And .PenPressure > 0 Then
                        Answ = _
                        dbMsgBox(1197, vbYesNoCancel Or vbQuestion)
                                '"You have just used a tablet or other pressure sensitive device. Do you want to set it up now?" + vbCrLf + _
                                 "Yes = Set up the device;" + vbCrLf + _
                                 "No = Forget about such devices;" + vbCrLf + _
                                 "Cancel = Do not set up the device and remind if you use it again.", vbYesNoCancel)
                        Select Case Answ
                            Case vbYes
                                Unload frmPressure
                                Load frmPressure
                                With frmPressure
                                    .MaxPr = 0
                                    .Timer1.Enabled = True
                                    .UnlNow = False
                                    .Show vbModal
                                End With
                                
                                Unload frmPressure
                            Case vbNo
                                .MaxPenPressure = -2
                            Case vbCancel
                                .MaxPenPressure = -1
                            
                        End Select
                    End If
                    If .MaxPenPressure < 0 Then .PenPressure = 0
                End With
            End If
                        
            'Message.Message = 0
        Case WM.WM_KEYDOWN
            ShiftKeyCode = Message.wParam And &HFFFF&
            ShiftKeyCode = ShiftKeyCode Or GetShiftState
            nActs = Keyb.GetActs(ShiftKeyCode, ActsAry)
            For i = 0 To nActs - 1
                If ActsAry(i) = dbCommands.cmdExtremeSave Then
                    MainForm.BuildBackup
                End If
            Next i
        Case WM.WM_TABLET_QUERYSYSTEMGESTURESTATUS
            If Message.hWnd = MainForm.MP.hWnd Then
              MainFormHook = eTabletWMResponse.TABLET_DISABLE_PRESSANDHOLD
              Exit Function
            End If
    End Select
End If
Select Case Message.Message
        Case WM.WM_ACTIVATEAPP, _
             WM.WM_ACTIVATE, _
             WM.WM_NCACTIVATE, _
             WM.WM_NCLBUTTONDOWN, _
             WM.WM_NCMBUTTONDOWN, _
             WM.WM_NCRBUTTONDOWN
            HideToolTipWindow
End Select
MainFormHook = CallNextHookEx(HookHandle, HookCode, wParam, Message)
eh:
Rec = False
End Function

Public Sub InstallHook()
If HookHandle <> 0 Then Exit Sub
HookHandle = SetWindowsHookEx(WH_GETMESSAGE, AddressOf MainFormHook, App.hInstance, App.ThreadID)
If HookHandle = 0 Then
    Err.Raise 111, "InstallHook", "Failed to install the hook (used for wheel movement detection)."
End If
dbSaveSettingEx "Special", "LastHookHandle", HookHandle, ForceRegistry:=True
End Sub

Public Sub UninstallHook(Optional ByVal ReadLastHookHandle As Boolean = False)
If HookHandle = 0 Then
    If ReadLastHookHandle Then
        'this is made for handling visual basic's stop button
        'or unexpected program termination
        HookHandle = dbGetSettingEx("Special", "LastHookHandle", vbLong, 0, ForceRegistry:=True)
        If HookHandle = 0 Then
            Err.Raise 111, , "Failed to uninstall the hook!!!"
        End If
    Else
        Exit Sub
    End If
End If
UnhookWindowsHookEx HookHandle
End Sub

Public Function IntersectRects(ByRef Rect1 As RECT, ByRef Rect2 As RECT) As RECT
Dim Rslt As RECT
With Rslt
    .Left = Max(Rect1.Left, Rect2.Left)
    .Top = Max(Rect1.Top, Rect2.Top)
    .Right = Min(Rect1.Right, Rect2.Right)
    .Bottom = Min(Rect1.Bottom, Rect2.Bottom)
    If .Left >= .Right Or .Top >= .Bottom Then
        .Left = Rect1.Left
        .Top = Rect1.Top
        .Right = Rect1.Left
        .Bottom = Rect1.Top
    End If
End With
IntersectRects = Rslt
End Function

Public Function BreakKeyPressed() As Boolean
Dim ret As Boolean
ret = WasKey(27) Or WasKey(19)
'            Esc           Break
If ret Then
    MainForm.MeEnabled = True
    MainForm.ClearMeEnabledStack
End If
BreakKeyPressed = ret
End Function

Public Sub EditProg(ByRef Prg As SMP, ByRef Section As String, ByVal Filt As DlgFilter)
Dim OrigPrg As SMP
Dim EV As New clsEVal
OrigPrg = Prg
On Error GoTo eh
rsm:
Load frmProgTool
With frmProgTool
    .Filt = Filt
    .Section = .Section
    .DontCompile = True
    .LoadSettings
    ShowFormModal frmProgTool
    If .Tag <> "" Then Exit Sub
    Prg.Source = .txtPrg.Text
    EV.CompileExpression_Ex Prg.Source, Prg.Code, Prg.Vars
End With
Unload frmProgTool
Exit Sub
eh:
If Err.Number = dbCWS Then
Else
    If dbMsgBox(Err.Description + "`Compile Error", vbCritical Or vbRetryCancel) = vbRetry Then
        Resume rsm
    End If
End If
Unload frmProgTool
Prg = OrigPrg
Exit Sub
End Sub

Public Sub InterpolateInt(ByRef Points() As PointByte, ByRef Output() As Byte, ByVal Mode As vtInterpolMode)
Const MaxN = 15
Dim i As Long, j As Long, h As Long
Dim y As Double, s As Double
'Dim P() As Double
Dim n As Long
ReDim Output(0 To 255)
n = UBound(Points)

If n < MaxN And Mode = dbIMPolynomial Then
    ReDim P(0 To n, 0 To n)
    For j = 0 To n
        For i = 0 To n
            If i <> j Then
                P(j, i) = 1 / (CDbl(Points(j).x) - Points(i).x)
            End If
        Next i
    Next j
    For h = 0 To 255
        s = 0
        For i = 0 To n
            y = Points(i).y
            For j = 0 To n
                If i <> j Then
                    y = y * (Points(j).x - h) * P(j, i)
                End If
            Next j
            s = s + y
        Next i
        If s > 255 Then s = 255
        If s < 0 Then s = 0
        Output(h) = s
    Next h
Else
    For i = 0 To 255
        If i = Points(j + 1).x Then
            j = j + 1
            Output(i) = Points(j).y
        Else
            Output(i) = Points(j).y + (CLng(Points(j + 1).y) - CLng(Points(j).y)) * (i - Points(j).x) / (0& + Points(j + 1).x - Points(j).x)
        End If
    Next i
End If


End Sub

Public Sub CancelDoEvents(Optional ByVal aDontDoEvents As Boolean = True)
DontDoEventsStack(0) = False
DontDoEventsStackPointer = DontDoEventsStackPointer + 1&
DontDoEventsStack(DontDoEventsStackPointer) = DontDoEvents
DontDoEvents = aDontDoEvents
End Sub

Public Sub RestoreDoEvents()
DontDoEvents = DontDoEventsStack(DontDoEventsStackPointer)
DontDoEventsStackPointer = DontDoEventsStackPointer - 1&
If DontDoEventsStackPointer < 0 Then DontDoEventsStackPointer = 0
End Sub

Public Function ConvertColorLng(ByVal lngColor As Long) As Long
ConvertColorLng = (lngColor And &HFF00&) Or _
                  (lngColor And &HFF&) * &H10000 Or _
                  (lngColor And &HFF0000) \ &H10000
End Function

Public Sub ViewImage(ByRef Data() As Long, _
                     Optional ByRef Purpose As String)
Load frmViewImage
With frmViewImage
    .SetImage Data
    .SetPurpose Purpose
    .Show vbModal
    .GetImage Data
End With
Unload frmViewImage
End Sub

Public Function Format_Size(ByVal Size As Currency, _
                            ByVal n1000 As Currency, _
                            ByVal nDigits As Long)
Dim Divider As Currency
Dim PowDivider As Long
Dim Sz As Currency
Dim tmp As String
Dim i As Long
Dim cnt As Long
Dim WasComma As Boolean
Dim tmp1 As String
If nDigits <= 0 Then
    Err.Raise 1111, "Format_Size", "Cannot output a zero- or negative-digit-count number."
End If
If Size < 0 Then
    Err.Raise 1111, "Format_Size", "Negative values are not supported."
End If
Divider = 1
PowDivider = 0
Do While Size / Divider > n1000
    Divider = Divider * n1000
    PowDivider = PowDivider + 1
Loop
Sz = Size / Divider
tmp = dbCStr(Sz)
cnt = 0
For i = 1 To Len(tmp)
    Select Case Mid$(tmp, i, 1)
        Case ",", "."
            WasComma = True
        Case "0" To "9"
            cnt = cnt + 1
    End Select
    If cnt <= nDigits Then
        tmp1 = tmp1 + Mid$(tmp, i, 1)
    Else
        If Not WasComma Then
            tmp1 = tmp1 + Mid$(tmp, i, 1)
        End If
    End If
Next i
Debug.Assert Len(tmp1) > 0
If Mid$(tmp1, Len(tmp1), 1) = "," Or Mid$(tmp1, Len(tmp1), 1) = "." Then
    tmp1 = Left$(tmp1, Len(tmp1) - 1)
End If
Select Case PowDivider
    Case 0
        'nothing
    Case 1
        tmp1 = tmp1 + "K"
    Case 2
        tmp1 = tmp1 + "M"
    Case 3
        tmp1 = tmp1 + "G"
    Case 4
        tmp1 = tmp1 + "T"
    Case Else
        Debug.Assert False
End Select
Format_Size = tmp1
End Function

Public Sub SimplifyFraction(ByRef Val1 As Long, ByRef Val2 As Long)
Dim Sq As Long
Dim i As Long
Sq = Min(Abs(Val1), Abs(Val2))
For i = 2 To Sq
    Do While (Val1 Mod i = 0) And (Val2 Mod i) = 0
        Val1 = Val1 \ i
        Val2 = Val2 \ i
    Loop
Next i
End Sub

Public Function AspectRatioToStr(ByVal Val1 As Long, ByVal Val2 As Long) As String
SimplifyFraction Val1, Val2
AspectRatioToStr = dbCStr(Val1) + ":" + dbCStr(Val2)
End Function

Public Sub vtRepair(ByRef Data() As Long, Optional ByVal StartY As Long = 0, Optional ByVal NScanLines As Long = -1)
Dim w As Long, h As Long
Dim i As Long
Dim lngRaw() As Long
Dim iStart As Long, iEnd As Long

On Error GoTo eh
TestDims Data
w = UBound(Data, 1) + 1
h = UBound(Data, 2) + 1
ConstructAry AryPtr(lngRaw), VarPtr(Data(0, 0)), 4, w * h
If NScanLines < 0 Or StartY < 0 Then
  iStart = 0
  iEnd = w * h - 1
Else
  If StartY >= h Then GoTo ExitHere
  iStart = StartY * w
  iEnd = Min(iStart + NScanLines * w, h * w) - 1
End If
For i = iStart To iEnd
    lngRaw(i) = lngRaw(i) And &HFFFFFF
Next i
ExitHere:
UnReferAry AryPtr(lngRaw)
Exit Sub
eh:
UnReferAry AryPtr(lngRaw)
ErrRaise "vtRepair"
End Sub

Public Sub TestDims(ByRef Ary() As Long, Optional ByVal nDims As Long = 2)
If AryDims(AryPtr(Ary)) <> nDims Then
    Err.Raise 1111, , "A " + CStr(nDims) + "-dimensional array is required."
End If
End Sub

Public Sub EditNumber(ByRef Number As Variant, _
                      ByRef Message As Variant, _
                      ByVal MinValue As Variant, _
                      ByVal MaxValue As Variant)
Dim OrigNumber As Variant
Dim VT As VbVarType
On Error GoTo eh
OrigNumber = Number
VT = VarType(Number)
If VT = vbEmpty Then
    VT = vbDouble
End If
rsm:
Number = dbVal(dbInputBox(Message, Number, CancelError:=True), VT)
If Not IsMissing(MinValue) And Not IsEmpty(MinValue) And _
   Not IsMissing(MaxValue) And Not IsEmpty(MaxValue) Then
    If Number < MinValue Then
        Err.Raise 101
    End If
    If Number > MaxValue Then
        Err.Raise 102
    End If
End If
Exit Sub
eh:
If Err.Number = dbCWS Then
    Number = OrigNumber
    ErrRaise "EditNumber"
ElseIf Err.Number = 101 Then
    dbMsgBox grs(2504, "%min", dbCStr(MinValue), _
                       "%max", dbCStr(MaxValue)), vbInformation
    Number = MinValue
    Resume rsm
ElseIf Err.Number = 102 Then
    dbMsgBox grs(2505, "%min", dbCStr(MinValue), _
                       "%max", dbCStr(MaxValue)), vbInformation
    Number = MaxValue
    Resume rsm
ElseIf Err.Number = 13 Then
    dbMsgBox 2506, vbInformation
    Resume rsm
ElseIf Err.Number = 6 Then
    dbMsgBox 2507, vbInformation
    Resume rsm
Else
    ErrRaise "EditNumber"
End If
End Sub

Public Function dBtoFactor(ByVal dB As Double) As Double
dBtoFactor = 10 ^ (dB / 10)
End Function

Public Function FactorToDB(ByVal Factor As Double) As Double
FactorToDB = Log(Factor) / Log(10) * 10
End Function

Public Sub vtBeep()
    
    'Debug.Assert False
       Beep
    
End Sub

Public Function GetExecCount()
GetExecCount = Val(dbGetSetting("Special", "ExecCount", "0", _
                                CommonSetting:=True, _
                                AllUsers:=True, _
                                ForceRegistry:=True _
                                ) _
                   )
End Function

Public Sub IncExecCount()
dbSaveSetting "Special", "ExecCount", _
              CStr(GetExecCount + 1), _
              CommonSetting:=True, _
              AllUsers:=True, _
              ForceRegistry:=True
End Sub

Public Function ProjectCompiled() As Boolean
gProjectCompiled = True
'if not compiled, assertion will execute. Use it to determine.
Debug.Assert AssignProjectNotCompiled
ProjectCompiled = gProjectCompiled
End Function

Private Function AssignProjectNotCompiled() As Boolean
gProjectCompiled = False
AssignProjectNotCompiled = True 'for assertion not to occur
End Function

Public Function DefKeysString() As String
Dim St As String
Dim ByteArray() As Byte
ByteArray = LoadResData("DEFKEYS", "SHORTCUTS")
St = ByteArray
DefKeysString = StrConv(St, vbUnicode)
End Function

Public Sub SetWndDblClick(ByVal hWnd As Long, _
                          ByVal EnableDblclicks As Boolean)
WinAPI.ModifyWindowStyle hWnd, CS_DBLCLKS And CLng(EnableDblclicks), CS_DBLCLKS And CLng(Not EnableDblclicks)
End Sub


Public Sub SubClassMP()
If OriginalMPWndProc <> 0 Then Exit Sub
OriginalMPWndProc = SetWindowLong(MainForm.MP.hWnd, GWL_WNDPROC, AddressOf MPWndProc)
If OriginalMPWndProc = 0 Then Err.Raise 12345, "SubClassMP", "Window procedure replacement failed"
End Sub

Public Sub UnSubClassMP()
If OriginalMPWndProc = 0 Then Exit Sub
SetWindowLong MainForm.MP.hWnd, GWL_WNDPROC, OriginalMPWndProc
OriginalMPWndProc = 0
End Sub

Public Function MPWndProc(ByVal hWnd As Long, ByVal Message As WM, ByVal wParam As Long, ByVal lParam As Long) As Long
'If Message >= &H2C0& And Message <= &H2C0& + &H20& Then
'  Debug.Print Message - &H2C0&
'End If
Dim SS As dbShiftConstants
Dim SBevent As SB
Select Case Message
  Case WM.WM_TABLET_QUERYSYSTEMGESTURESTATUS
    MPWndProc = (eTabletWMResponse.TABLET_DISABLE_PRESSANDHOLD And 0& Or eTabletWMResponse.TABLET_DISABLE_FLICKS Or eTabletWMResponse.TABLET_DISABLE_TOUCHUIFORCEOFF)
    MainForm.MDPen = True
  Case WM.WM_Wheel
    SS = GetShiftState
    MainForm.Form_Wheel ((wParam And &HFFFF0000) \ &H10000) / CDbl(WHEEL_DELTA), SS
  Case WM.WM_VSCROLL
    SS = GetShiftState
    SBevent = wParam And &HFFFF&
    If SBevent = SB_LINEUP Then
      MainForm.MoveByWheel 1, SS
    ElseIf SBevent = SB_LINEDOWN Then
      MainForm.MoveByWheel -1, SS
    End If
  Case WM.WM_HSCROLL
    SS = GetShiftState Xor dbStateShift
    SBevent = wParam And &HFFFF&
    If SBevent = SB_LINELEFT Then
      MainForm.MoveByWheel 1, SS
    ElseIf SBevent = SB_LINERIGHT Then
      MainForm.MoveByWheel -1, SS
    End If
  Case Else
    MPWndProc = WinAPI.CallWindowProc(OriginalMPWndProc, hWnd, Message, wParam, lParam)
End Select
End Function



'vista setpixel bug: in non-autoredraw windows, setpixel does nothing
'if bit 8 of x-coordinate is set. Appears only when Aero is OFF.
Public Sub TestForSetPixelBug()
Dim frm As New frmFormatVB
Load frm
frm.Show
SetWindowPos frm.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
DoEvents
Const TestColor = &HABCDEF
Const TestX1 = 255
Const TestX2 = 256
Dim origcolor1 As Long, origcolor2 As Long
Dim Bug1 As Boolean, Bug2 As Boolean
origcolor1 = GetPixel(frm.hDC, TestX1, 0)
SetPixel frm.hDC, TestX1, 0, TestColor
Bug1 = GetPixel(frm.hDC, TestX1, 0) = origcolor1
origcolor2 = GetPixel(frm.hDC, TestX2, 0)
SetPixel frm.hDC, TestX2, 0, TestColor
Bug2 = GetPixel(frm.hDC, TestX2, 0) = origcolor2
If Bug1 Then
  VistaSetPixelBugDetected = False
  Debug.Print "Vista bug: failed to determine"
  'MsgBox "Vista bug: failed to determine"
Else
  VistaSetPixelBugDetected = Bug2
  If VistaSetPixelBugDetected Then
    Debug.Print "Vista bug: PRESENT!"
    'MsgBox "Vista bug: PRESENT!"
  Else
    Debug.Print "Vista bug: absent"
    'MsgBox "Vista bug: absent"
  End If
End If
Unload frm
End Sub
