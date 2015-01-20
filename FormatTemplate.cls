VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FormatTemplate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Enum eBadSettings
    bsAlphaRequired = 1
    bsSettingsInvalid = 2
End Enum

Public Enum eFilterType
    ftForSaving = 0
    ftForLoading = 1
End Enum

Public Sub SetPtrData(ByVal ptrData As Long)

End Sub

Public Sub RemovePtrData()

End Sub

Public Sub ExtractData(ByVal ptrData As Long)

End Sub

Public Sub Customize()

End Sub

Public Function IsFormat(ByVal FileNumber As Long) As Boolean

End Function

Public Function LoadFile(ByRef FileName As String, _
                         ByVal UpdateSettings As Boolean) As Boolean
'returns True if alpha channel is present
End Function

Public Sub SaveFile(ByRef FileName As String, ByVal HasAlpha As Boolean)

End Sub

Public Sub SaveSettings()

End Sub

Public Sub LoadSettings()

End Sub

Public Sub GetInfo(ByRef ID As String, _
                   ByRef ListName As String, _
                   ByRef HasSettings As Boolean, _
                   Optional ByRef CanSave As Boolean, _
                   Optional ByRef CanLoad As Boolean)

End Sub

Public Sub GetFileTypeInfo(Optional ByRef ExtList As String, _
                           Optional ByRef TypeDescription As String, _
                           Optional ByRef IconString As String, _
                           Optional ByRef DefEditor As Boolean)

End Sub

Public Function GetFilter(Optional ByVal FilterType As eFilterType = eFilterType.ftForSaving) As String

End Function

Public Function GetDescriptionResID() As Long

End Function

Public Function AlphaSupproted() As Boolean

End Function

Public Function CanSave(ByVal AlphaPresent As Boolean, ByRef Reason As eBadSettings) As Boolean

End Function

