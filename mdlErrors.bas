Attribute VB_Name = "mdlErrors"
Option Explicit

Const DisableAssertions As Boolean = True

Public Type vtError
    Number As Long
    Source As String
    Description As String
End Type

Private ProjectName As String
Private ErrorStack() As vtError
Private nInStack As Long

Public Sub ReadError_Arg(ByRef vErr As vtError)
vErr.Number = Err.Number
vErr.Source = Err.Source
vErr.Description = Err.Description
End Sub

Public Function ReadError() As vtError
Dim vErr As vtError
ReadError_Arg vErr
ReadError = vErr
End Function

Public Sub vtRaiseError(ByRef aErr As vtError, _
                      Optional ByVal ProcedureName As String)
If Len(ProjectName) = 0 Then
    GetProjectName
End If
Debug.Assert aErr.Number = dbCWS Or DisableAssertions
If Len(ProcedureName) > 0 Then
    If aErr.Source = ProjectName Then
        Err.Raise aErr.Number, ProcedureName, aErr.Description
    Else
        Err.Raise aErr.Number, aErr.Source, aErr.Description
    End If
Else
    Err.Raise aErr.Number, aErr.Source, aErr.Description
End If
End Sub

Private Function GetProjectName()
Dim vErr As vtError
PushError
On Error Resume Next
Err.Raise 5
ProjectName = Err.Source
PopError
End Function

Public Sub RaiseError(Optional ByVal ProcedureName As String)
Dim vErr As vtError
ReadError_Arg vErr
vtRaiseError vErr, ProcedureName
'If Len(ProjectName) = 0 Then
'    GetProjectName
'End If
'If Len(ProcedureName) > 0 Then
'    If aErr.Source = ProjectName Then
'        Err.Raise aErr.Number, ProcedureName, aErr.Description
'    Else
'        Err.Raise aErr.Number, aErr.Source, aErr.Description
'    End If
'Else
'    Err.Raise aErr.Number, aErr.Source, aErr.Description
'End If
End Sub


Public Sub ErrRaise(Optional ByVal ProcedureName As String)
RaiseError ProcedureName
End Sub

Public Sub vtErrRaise(ByRef aErr As vtError, _
                    Optional ByVal ProcedureName As String)
vtRaiseError aErr, ProcedureName
End Sub

Function MsgError(Optional ByVal Message As Variant = "", _
                  Optional ByVal Style As VbMsgBoxStyle = vbCritical, _
                  Optional ByVal Assertion As Boolean = False) As VbMsgBoxResult
Dim strMessage As String
Dim bAddErrDesc As Boolean
PushError
If Err.Number = dbCWS Then
    MsgError = vbCancel
    Exit Function
End If
On Error Resume Next
    If IsNumeric(Message) Then
        strMessage = GRSF(Message, RaiseErrors:=True)
    Else
        strMessage = CStr(Message)
    End If
On Error GoTo 0
PopError
If Len(strMessage) = 0 Then
    strMessage = Err.Description
ElseIf InStr(1&, strMessage, "Err.Description", vbTextCompare) Then
    strMessage = Replace(strMessage, "Err.Description", Err.Description)
Else
    strMessage = strMessage + vbNewLine + "(" + Err.Description + ")"
End If
MsgError = dbMsgBox(strMessage, Style)
Debug.Assert Assertion Or DisableAssertions
End Function

Function vtMsgError(ByRef pErr As vtError, _
                    Optional ByVal Style As VbMsgBoxStyle = vbCritical, _
                    Optional ByVal Assertion As Boolean = False) As VbMsgBoxResult
If pErr.Number = dbCWS Then
    vtMsgError = vbCancel
    Exit Function
End If
vtMsgError = MsgBox(pErr.Description, Style, pErr.Source)
Debug.Assert Assertion Or DisableAssertions
End Function


'Sub SetWindowCursor(ByVal hWnd As Long, _
'                    ByVal ResID As Integer)
'Dim hCur As Long
'If ResID = 0 Then
'    hCur = LoadCursor(0, IDC_ARROW)
'Else
'    hCur = LoadImage(App.hInstance, ByVal ResID, vbResCursor, 0, 0, LR_DEFAULTSIZE Or LR_SHARED)
'End If
'SetClassLong hWnd, GCL_HCURSOR, hCur
'End Sub

Public Sub PushError()
Dim vErr As vtError
ReadError_Arg vErr
If nInStack = 0 Then
    ReDim ErrorStack(0 To 0)
Else
    ReDim Preserve ErrorStack(0 To nInStack)
End If
ErrorStack(nInStack) = vErr
nInStack = nInStack + 1
End Sub

Public Sub PopError(Optional ByVal RaiseIt As Boolean = False)
Dim vErr As vtError
nInStack = nInStack - 1
If nInStack >= 0 Then
    vErr = ErrorStack(nInStack)
Else
    nInStack = 0
End If
If nInStack > 0 Then
    ReDim Preserve ErrorStack(0 To nInStack - 1)
Else
    Erase ErrorStack
End If
If Not RaiseIt Then On Error Resume Next
Err.Raise vErr.Number, vErr.Source, vErr.Description
End Sub

Public Sub WriteError(ByVal ErrNumber As Long, _
                      ByVal ErrSource As String, _
                      ByVal ErrDescription As String)
On Error Resume Next
Err.Raise ErrNumber, ErrSource, ErrDescription
End Sub

