VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFormatVB"
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
ID = "VB"
ListName = "VB's supported formats"
HasSettings = False
CanSave = False
CanLoad = True
End Sub

Private Sub FormatTemplate_GetFileTypeInfo(Optional ByRef ExtList As String, _
                                           Optional ByRef TypeDescription As String, _
                                           Optional ByRef IconString As String, _
                                           Optional ByRef DefEditor As Boolean)
ExtList = "rle|gif|jpg|jpeg|wmf|emf"
TypeDescription = GRSF(2607)
DefEditor = False
IconString = ""
End Sub

Private Function FormatTemplate_GetFilter(Optional ByVal FilterType As eFilterType = 0&) As String
If FilterType = ftForLoading Then
    FormatTemplate_GetFilter = "Windows Bitmaps (.bmp)|*.bmp;*.dib|JPEG's (.jpg)|*.jpg|GIF's (.gif)|*.gif|Metafiles (.emf .wmf)|*.emf;*.wmf|Icons, cursors (.ico .cur)|*.ico;*.cur|RLE-compressed Bitmaps (.rle)|*.rle"
ElseIf FilterType = ftForSaving Then
    FormatTemplate_GetFilter = ""
End If
End Function

Private Function FormatTemplate_CanSave(ByVal AlphaPresent As Boolean, Reason As eBadSettings) As Boolean
FormatTemplate_CanSave = False
End Function

Private Function FormatTemplate_AlphaSupproted() As Boolean
FormatTemplate_AlphaSupproted = True
End Function

Private Function FormatTemplate_GetDescriptionResID() As Long
FormatTemplate_GetDescriptionResID = 0
End Function

Private Sub FormatTemplate_Customize()
End Sub

Private Sub FormatTemplate_ExtractData(ByVal ptrData As Long)
SwapArys ptrData, AryPtr(LoadedData)
Erase LoadedData
End Sub

Private Function FormatTemplate_IsFormat(ByVal FileNumber As Long) As Boolean
FormatTemplate_IsFormat = True 'IsFormat(FileNumber)
End Function

Private Function FormatTemplate_LoadFile(FileName As String, ByVal UpdateSettings As Boolean) As Boolean
FormatTemplate_LoadFile = LoadFile(FileName, UpdateSettings)
End Function

Private Sub FormatTemplate_RemovePtrData()
pPtrData = 0
End Sub

Private Sub FormatTemplate_SaveFile(FileName As String, ByVal HasAlpha As Boolean)
End Sub

Private Sub FormatTemplate_LoadSettings()

'    If SettingsLoaded Then Exit Sub
'
'    On Error Resume Next
'
'    Interlace = dbGetSettingEx("Formats\PNG", "Interlace", vbBoolean, False)
'
'    SettingsLoaded = True
'
End Sub

Private Sub FormatTemplate_SaveSettings()

'    If Not SettingsLoaded Then Exit Sub
'
'    dbSaveSettingEx "Formats\PNG", "Interlace", Interlace
    

End Sub

Private Sub FormatTemplate_SetPtrData(ByVal ptrData As Long)
'pPtrData = ptrData
End Sub





'-------------------------------------------------------------------
Private Sub SettingsToForm()
'FormatTemplate_LoadSettings
'
'Load frmSettings
'With frmSettings
'
'    .Interlaced = Interlace
'
'End With
End Sub

Private Sub FormToSettings()
'
'    With frmSettings
'
'        Interlace = .Interlaced
'
'    End With
'
'    Unload frmSettings
'
End Sub
'/dialogs

'-------------------------------------------------------------------
'File reading

Private Function LoadFile(ByRef FileName As String, _
                          ByVal SetSettings As Boolean) As Boolean
Dim Image As IPictureDisp

Set Image = LoadPicture(FileName)
LoadFile = GetPicData(Image, LoadedData, CalcAlpha:=UCase$(GetExt(FileName)) <> "JPG")

Exit Function
eh:
ErrRaise "clsFormatVB:LoadFile"
End Function
