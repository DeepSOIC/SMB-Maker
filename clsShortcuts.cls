VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsShortcuts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Public Enum dbCommands
    cmdNewInstance = 1
    cmdCloseInstance = 2

    cmdNew = 3
    cmdOpen = 4
    cmdLoadPAL = 5
    cmdSave = 6
    cmdoSaveICO = 7
    cmdoSaveSMB = 8
    cmdSaveAs = 9
    cmdoSavePNG = 10
    cmdSavePAL = 11
    
    cmdCopy = 12
    cmdPaste = 13
    cmdLoadSel = 14
    
    cmdDeselect = 15
    cmdDeleteSel = 16
    cmdDeselect_no_delete = 17
    
    cmdResize = 18
    cmdNextTool = 19
    cmdPrevTool = 20
    cmdToolSel = 21
    cmdToolTxt = 22
    cmdToolPen = 23
    cmdToolLin = 24
    cmdToolRec = 25
    cmdToolCir = 26
    cmdToolPol = 27
    cmdToolSta = 28
    cmdToolAir = 29
    cmdToolBrh = 30
    cmdToolPnt = 31
    cmdToolGet = 32
    cmdToolFad = 33
    cmdToolVFd = 34
    cmdToolHFd = 35
    cmdToolFSt = 36
    cmdToolHel = 37
    cmdToolPal = 38
    cmdoToolHot = 39
    cmdoToolWav = 40
    cmdToolProps = 41
    
    cmdPaintBrush = 42
    cmdRefresh = 43
    cmdEndPoly = 44
    cmdCyclePoly = 45
    
    cmdUndo = 46
    cmdRedo = 47
    cmdClearUndo = 48
    cmdSetUndoCount = 49
    cmdDisableUndo = 50
    
    cmdZoomIn = 51
    cmdZoomOut = 52
    cmdZoom = 53
    cmdoToggleAutoRedraw = 54
    cmdToggleGrid = 55
    cmdPalVisible = 56
    cmdToolBarVisible = 57
    
    cmdCursorUp = 58
    cmdCursorLeft = 59
    cmdCursorDown = 60
    cmdCursorRight = 61
    cmdScrollUp = 62
    cmdScrollLeft = 63
    cmdScrollDown = 64
    cmdScrollRight = 65
    
    cmdMakePalDef = 66
    cmdStretchPal = 67
    cmdNPalEntries = 68
    cmdPalFillTips = 69
    cmdPalClearTips = 70
    cmdPalReset = 71
    cmdoSyncACol = 72
    cmdChangeFCol = 73
    cmdChangeBCol = 74
    cmdDefPalSysColors = 75
    cmdDefPalBRH = 76
    cmdDefPal16 = 77
    cmdDefPal256 = 78
    
    cmdKeyboard = 79
    cmdToggleMiddleButtonUse = 80
    cmdoBMPSettings = 81
    cmdoBGToolBar = 82
    cmdoBGWorkspace = 83
    cmdExts = 84
    cmdoDllPath = 85
    cmdFullReset = 86
    
    cmdWhatsThis = 87
    
    cmdLMB = 88
    cmdRMB = 89
    cmdMMB = 90
    
    cmdoClearCharges = 91
    cmdoDelLastCharge = 92
    
    cmdClearPic = 93
    cmdMinimize = 94
    cmdDrawPlain = 95
    cmdDrawBg = 96
    cmdDrawBBg = 97
    
    cmdPrgDrawings = 98
    cmdoPrgPhys = 99 'Obsolete
    
    cmdEffect0 = 100
    cmdEffect1 = 101
    cmdEffect2 = 102
    cmdEffect3 = 103
    cmdEffect4 = 104
    cmdEffect5 = 105
    cmdEffect6 = 106
    cmdEffect7 = 107
    cmdEffect8 = 108
    cmdEffect9 = 109
    
    cmdRepLastEffect = 110
    
    cmdSaveSel = 111
    cmdSelectAll = 112
    cmdCropSel = 113
    
    cmdCapture = 114
    cmdCaptureActiveWindow = 115
    cmdCapturePointedWindow = 116
    cmdCaptureScreen = 117
    cmdCapturePoint = 118
    
    cmdToolOrg = 119
    cmdToolPrg = 120
    
    cmdEffect10 = 121
    
    cmdDynamicDialog = 122
    
    cmdIdleMessage = 123
    
    cmdTexMode = 124
    cmdResetOrg = 125
    cmdRestoreOrg = 126
    
    cmdExtremeSave = 127
    
    cmdFormula = 128
    
    cmdEffect11 = 129 'ClearType
    
    cmdOAutoScrolling = 130
    cmdGlueMouse = 131
    
    cmdSelMoveTo = 132
    cmdSelHCenter = 133
    cmdSelVCenter = 134
    cmdSelClear = 135
    cmdSelResize = 136
    cmdSelEdit = 137
    cmdSelToggleRedraw = 138
    cmdSelShow = 139
    
    cmdDrawWaves = 140
    
    cmdNaviMode = 141
    
    cmdSelShowSize = 142
    
    cmdWheelUp = 143
    cmdWheelDown = 144
    
    cmdZoom1 = 145
        
    cmdLast = 145
    
    cmdNone = 0
    cmdUnknown = -1
End Enum


Dim pkArr() As kShortcut
Dim ActLists(0 To 255) As typKeyList

Private Type typKeyList
    List() As Long 'key
    nItems As Long
End Type

Private Type typShortcutsPack
    Keys() As kShortcut
    Actkeys(0 To 255) As typKeyList
End Type


Function GetCount()
If AryDims(AryPtr(pkArr)) = 1 Then
    GetCount = UBound(pkArr) + 1
Else
    GetCount = 0
End If
End Function

Function GetKey(ByVal Index As Long) As Long
If Index < 0 Or Index > GetCount - 1 Then
    Err.Raise 9, "GetKey", "Range-checking failed!"
End If
GetKey = pkArr(Index).Key
End Function

Function GetAction(ByVal Index As Long) As Long
If Index < 0 Or Index > GetCount - 1 Then
    Err.Raise 9, "GetAction", "Range-checking failed!"
End If
GetAction = pkArr(Index).Act
End Function

Public Function ListKeys(ByVal Act As dbCommands, _
                         ByRef sArr() As String) As Long
Dim i As Long
Dim n As Long
If Act = cmdNone Or Act = cmdUnknown Then Exit Function
n = ActLists(Act).nItems
If n > 0 Then
    ReDim sArr(0 To n - 1)
Else
    Erase sArr
End If
For i = 0 To n - 1
    sArr(i) = GetKeyName(ActLists(Act).List(i))
Next i
ListKeys = n
End Function

Private Sub UpdateKeyLists()
Dim Act As Long
Dim i As Long
For Act = 0 To 255
    ActLists(Act).nItems = 0
    Erase ActLists(Act).List
Next Act
For i = 0 To GetCount - 1
    AddKeyToList ActLists(pkArr(i).Act), pkArr(i).Key
Next i
End Sub

Private Sub AddKeyToList(ByRef List As typKeyList, _
                         ByVal KeyCode As Long, _
                         Optional ByVal ScanForSimilar As Boolean = True)
Dim i As Long
If List.nItems = 0 Then
    ReDim List.List(0 To 0)
End If

If ScanForSimilar Then
    For i = 0 To List.nItems - 1
        If List.List(i) = KeyCode Then
            Exit Sub
        End If
    Next i
End If

List.nItems = List.nItems + 1
ReDim Preserve List.List(0 To List.nItems - 1)
List.List(List.nItems - 1) = KeyCode
End Sub

Public Function GetActs(ByVal KeyCode As Long, _
                        ByRef AryActs() As Long) As Long
Dim i As Long
Dim ActsCnt As Long
For i = 0 To GetCount - 1
    If KeyCodesEqual(pkArr(i).Key, KeyCode) Then
        GoSub AddItem
    End If
Next i
GetActs = ActsCnt
Exit Function
AddItem:
    If ActsCnt = 0 Then
        ReDim AryActs(0 To 0)
    End If
    ActsCnt = ActsCnt + 1
    ReDim Preserve AryActs(0 To ActsCnt - 1)
    AryActs(ActsCnt - 1) = pkArr(i).Act
Return
End Function

Sub SaveToReg(ByRef Section As String, ByRef Parameter As String)
SaveKeysToReg pkArr, Section, Parameter
End Sub

Sub LoadFromReg(ByRef Section As String, ByRef Parameter As String)
LoadKeysFromReg pkArr, Section, Parameter
UpdateKeyLists
End Sub

Sub ShowEditDialog(Optional ByVal RaiseCancelError As Boolean = True)
Dim Cancelled As Boolean
Load frmKeys
With frmKeys
    .SetKeys pkArr
    .Show vbModal
    If .Tag = "" Then
        .ExtractKeys pkArr
        UpdateKeyLists
    Else
        Cancelled = True
    End If
End With
Unload frmKeys
If Cancelled And RaiseCancelError Then Err.Raise dbCWS
End Sub
