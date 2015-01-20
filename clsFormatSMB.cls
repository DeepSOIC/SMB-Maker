VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFormatSMB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'common declarations
Implements FormatTemplate
Dim pPtrData As Long
Dim LoadedData() As Long
Dim SettingsLoaded As Boolean

'-----------------------------------
'format specific declarations

Private Sub FormatTemplate_GetInfo(ID As String, ListName As String, HasSettings As Boolean, Optional CanSave As Boolean, Optional CanLoad As Boolean)
ID = "SMB"
ListName = "SMB bitmap (SMB)"
HasSettings = False
CanSave = True
CanLoad = True
End Sub

Private Sub FormatTemplate_GetFileTypeInfo(Optional ByRef ExtList As String, _
                                           Optional ByRef TypeDescription As String, _
                                           Optional ByRef IconString As String, _
                                           Optional ByRef DefEditor As Boolean)
ExtList = "smb"
TypeDescription = GRSF(2608)
DefEditor = False
IconString = ""
End Sub

Private Function FormatTemplate_GetFilter(Optional ByVal FilterType As eFilterType = 0&) As String
If FilterType = ftForLoading Then
    FormatTemplate_GetFilter = "SMB's (.smb)|*.smb"
ElseIf FilterType = ftForSaving Then
    FormatTemplate_GetFilter = "SMB (.smb)|*.smb"
End If
End Function

Private Function FormatTemplate_CanSave(ByVal AlphaPresent As Boolean, Reason As eBadSettings) As Boolean
FormatTemplate_CanSave = True
End Function

Private Function FormatTemplate_AlphaSupproted() As Boolean
FormatTemplate_AlphaSupproted = True
End Function

Private Function FormatTemplate_GetDescriptionResID() As Long
FormatTemplate_GetDescriptionResID = 2613
End Function

Private Sub FormatTemplate_Customize()
'SettingsToForm
'On Error GoTo eh
'With frmSettings
'    .Show vbModal
'    If Len(.Tag) > 0 Then
'        Err.Raise dbCWS
'    End If
'End With
'FormToSettings
'Exit Sub
'eh:
'PushError
'Unload frmSettings
'PopError
'ErrRaise
End Sub

Private Sub FormatTemplate_ExtractData(ByVal ptrData As Long)
SwapArys ptrData, AryPtr(LoadedData)
Erase LoadedData
End Sub

Private Function FormatTemplate_IsFormat(ByVal FileNumber As Long) As Boolean
FormatTemplate_IsFormat = IsFormat(FileNumber)
End Function

Private Function FormatTemplate_LoadFile(FileName As String, ByVal UpdateSettings As Boolean) As Boolean
FormatTemplate_LoadFile = LoadFile(FileName, UpdateSettings)
End Function

Private Sub FormatTemplate_RemovePtrData()
pPtrData = 0
End Sub

Private Sub FormatTemplate_SaveFile(FileName As String, ByVal HasAlpha As Boolean)
SaveFile FileName, HasAlpha
End Sub

Private Sub FormatTemplate_LoadSettings()

    If SettingsLoaded Then Exit Sub
    
    On Error Resume Next
 
    SettingsLoaded = True

End Sub

Private Sub FormatTemplate_SaveSettings()

    If Not SettingsLoaded Then Exit Sub
    

End Sub

Private Sub FormatTemplate_SetPtrData(ByVal ptrData As Long)
pPtrData = ptrData
End Sub





'-------------------------------------------------------------------
Private Sub SettingsToForm()
End Sub

Private Sub FormToSettings()

    
End Sub
'/dialogs

'-------------------------------------------------------------------
'File reading

Private Function IsFormat(ByVal FileNumber As Long)
IsFormat = IsSMB(FileNumber)
End Function

Private Function LoadFile(ByRef FileName As String, _
                          ByVal SetSettings As Boolean) As Boolean
LoadFile = LoadSMB(LoadedData, FileName)
End Function


Public Sub SaveFile(ByRef FileName As String, _
                    ByVal WriteAlphaChannel As Boolean)
    Dim Data() As Long
    On Error GoTo eh
    ReferAry AryPtr(Data), pPtrData
    SaveSMB Data, FileName
    UnReferAry AryPtr(Data)
Exit Sub

eh:
    UnReferAry AryPtr(Data)
ErrRaise
End Sub


