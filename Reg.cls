VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Reg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Module for working with the windows registry.
'Uses: mdlVBArrays.

Option Explicit
Private Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As samDesired_Types, ByVal lpSecurityAttributes As Long, phkResult As Long, lpdwDisposition As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As eParameterType, lpData As Any, ByVal cbData As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As eParameterType, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As samDesired_Types, phkResult As Long) As Long
Private Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Private Declare Function RegEnumKeyEx Lib "advapi32.dll" Alias "RegEnumKeyExA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, lpcbName As Long, ByVal lpReserved As Long, ByVal lpClass As Any, lpcbClass As Long, lpftLastWriteTime As FILETIME) As Long
Private Type SECURITY_ATTRIBUTES
        nLength As Long
        lpSecurityDescriptor As Long
        bInheritHandle As Long
End Type

Private Type FILETIME
        dwLowDateTime As Long
        dwHighDateTime As Long
End Type

Public Enum samDesired_Types
 KEY_CREATE_LINK = &H20
 KEY_CREATE_SUB_KEY = &H4
 KEY_ENUMERATE_SUB_KEYS = &H8
 KEY_QUERY_VALUE = &H1
 KEY_SET_VALUE = &H2
 STANDARD_RIGHTS_ALL = &H1F0000
 KEY_NOTIFY = &H10
 SYNCHRONIZE = &H100000
 KEY_ALL_ACCESS = ((STANDARD_RIGHTS_ALL Or KEY_QUERY_VALUE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY Or KEY_CREATE_LINK) And (Not SYNCHRONIZE))
End Enum
 
Public Enum HKEYS
HKEY_CLASSES_ROOT = &H80000000
HKEY_CURRENT_CONFIG = &H80000005
HKEY_CURRENT_USER = &H80000001
HKEY_DYN_DATA = &H80000006
HKEY_LOCAL_MACHINE = &H80000002
HKEY_PERFORMANCE_DATA = &H80000004
HKEY_USERS = &H80000003
End Enum

Public Enum eParameterType
    REG_BINARY = 3                     ' Free form binary
    REG_DWORD = 4                      ' 32-bit number
    REG_DWORD_BIG_ENDIAN = 5           ' 32-bit number
    REG_DWORD_LITTLE_ENDIAN = 4        ' 32-bit number (same as REG_DWORD)
    REG_EXPAND_SZ = 2                  ' Unicode nul terminated string
    REG_MULTI_SZ = 7                   ' Multiple Unicode strings
    REG_LINK = 6                       ' Symbolic Link (unicode)
    REG_SZ = 1                         ' Unicode nul terminated string
End Enum

Private StrBuffer1 As String
Const pDebug As Boolean = True

Public Sub SetValue(ByVal hKey As HKEYS, _
                    ByRef SubKey As String, _
                    ByRef ValueName As String, _
                    ByRef NewValue As String, _
                    Optional ByVal RaiseErrors As Boolean = True)
Dim h As Long, c As Long, sss As SECURITY_ATTRIBUTES

If RegCreateKeyEx(hKey, SubKey, 0, "", 0, 131135, VarPtr(sss), h, c) = 0 Then
    If RegSetValueEx(h, ValueName, 0, 1, ByVal NewValue, Len(NewValue)) <> 0 Then NewValue = ""
    RegCloseKey h
Else
    If RaiseErrors Then
        Err.Raise 112, "Reg.SetValue", "Cannot open a key: '" + BuildPath(hKey, SubKey, ValueName) + "'."
    End If
End If
End Sub

Public Function GetHkeyNameByHandle(ByVal hK As HKEYS) As String
Dim Rslt As String
Select Case hK
    Case HKEYS.HKEY_CLASSES_ROOT
        Rslt = "HKCR"
    Case HKEYS.HKEY_CURRENT_CONFIG
        Rslt = "HKCC"
    Case HKEYS.HKEY_CURRENT_USER
        Rslt = "HKCU"
    Case HKEYS.HKEY_DYN_DATA
        Rslt = "HKDD"
    Case HKEYS.HKEY_LOCAL_MACHINE
        Rslt = "HKLM"
    Case HKEYS.HKEY_PERFORMANCE_DATA
        Rslt = "HKPD"
    Case HKEYS.HKEY_USERS
        Rslt = "HKU"
    Case Else
        Rslt = "<incorrect hkey handle:" + CStr(hK) + ">"
End Select
GetHkeyNameByHandle = Rslt
End Function

Public Function BuildPath(ByVal hK As HKEYS, ByVal SubKey As String, ByVal ValueName As String) As String
BuildPath = GetHkeyNameByHandle(hK) + "\" + SubKey + IIf(Len(ValueName) > 0, "::" + ValueName, "")
End Function

Public Function GetValue(ByRef hKey As HKEYS, _
                         ByRef SubKey As String, _
                         ByRef ValueName As String, _
                         Optional ByVal RaiseErrors As Boolean = True) As String
Dim h As Long
Dim lngLen As Long
Dim sss As SECURITY_ATTRIBUTES
Dim c As Long
Dim ts As Long
Dim OLen As Long
If RegCreateKeyEx(hKey, SubKey, 0, vbNullString, 0, 131135, VarPtr(sss), h, c) = 0 Then
    'strBuffer = Space$(65000)
    lngLen = Len(StrBuffer1)
    OLen = lngLen
    
    ts = RegQueryValueEx(h, ValueName, 0, 1, ByVal StrBuffer1, lngLen)
    If ts = 234 Then
        Do While ts = 234
            StrBuffer1 = Space$(Len(StrBuffer1) * 2)
            lngLen = Len(StrBuffer1)
            ts = RegQueryValueEx(h, ValueName, 0, 1, ByVal StrBuffer1, lngLen)
        Loop
    End If
    If ts <> 0 Then
        'strBuffer = ""
        If Len(StrBuffer1) <> OLen Then Class_Initialize
        If RaiseErrors Then
            RegCloseKey h
            Err.Raise 112, "Reg.GetValue", "Cannot Read the Key: '" + BuildPath(hKey, SubKey, ValueName) + "'. Registry error: " + CStr(ts) + "."
        End If
    Else
        If lngLen > 0 Then
        GetValue = Left$(StrBuffer1, lngLen - 1)
        End If
        If Len(StrBuffer1) <> OLen Then Class_Initialize
    End If
    RegCloseKey h
End If
End Function

Public Function GetAllValues(ByVal hKey As HKEYS, ByVal SubKey As String, ByRef Names() As String, ByRef Values() As String) As Long
Dim Rslt As Long
Dim nLen1 As Long
Dim strBuffer2 As String, nLen2 As Long
Dim lngType As Long
Dim i As Long
Dim ts As Long
If RegOpenKeyEx(hKey, ByVal SubKey, 0, KEY_ENUMERATE_SUB_KEYS Or KEY_QUERY_VALUE, Rslt) <> 0 Then
    Err.Raise 5, "GetAllValues", "Cannot open a key: '" + BuildPath(hKey, SubKey, "") + "'."
Else
    On Error GoTo eh
    ReDim Names(0 To 0), Values(0 To 0)
    i = 0
    
        StrBuffer1 = Space$(1024)
        nLen1 = 1024
        strBuffer2 = Space$(1024)
        nLen2 = 1024
    Do
        ts = RegEnumValue(Rslt, i, ByVal StrBuffer1, nLen1, 0&, lngType, ByVal strBuffer2, nLen2) <> 0
        If ts = 234 Then
            Do While ts = 234
                strBuffer2 = Space$(Len(strBuffer2) * 2)
                ts = RegEnumValue(Rslt, i, ByVal StrBuffer1, nLen1, 0&, lngType, ByVal strBuffer2, nLen2) <> 0
            Loop
            ts = 0
        End If
        If ts = 0 Then
        ReDim Preserve Names(0 To i), Values(0 To i)
            Values(i) = Left$(strBuffer2, nLen2)
            Names(i) = Left$(StrBuffer1, nLen1)
            
            i = i + 1
            StrBuffer1 = Space$(1024)
            nLen1 = 1024
            strBuffer2 = Space$(1024)
            nLen2 = 1024
        End If
    Loop Until ts <> 0
    
    RegCloseKey Rslt
End If
GetAllValues = i
Exit Function
eh:
RegCloseKey Rslt
Err.Raise Err.Number, "GetAllValues", Err.Description
End Function

Public Sub DeleteKey(ByVal hK As HKEYS, ByVal strKey As String, Optional ByVal RaiseErrors As Boolean = True)
Dim Rslt As Long, h As Long
Dim SubKeys() As String
Dim i As Long
Dim FT As FILETIME
Dim l As Long
Dim Buff As String
Dim Buff2 As String, l2 As Long

On Error GoTo eh
Buff = strKey + vbNullChar
Rslt = RegOpenKeyEx(hK, ByVal Buff, 0&, 131135, h)
If Rslt <> 0 Then
    If RaiseErrors Then
        Err.Raise 112, "DeleteKey", "Cannot open a key: '" + BuildPath(hK, strKey, "") + "'. Registry error: " + CStr(Rslt) + "."
    Else
        RegCloseKey h
        Exit Sub
    End If
Else
    i = -1
    ReDim SubKeys(0 To 0)
    Do
        i = i + 1
        ReDim Preserve SubKeys(0 To i)
        l = 1024
        Buff = Space$(l)
        l2 = 256
        Buff2 = Space$(l2)
        Rslt = RegEnumKeyEx(h, i, ByVal Buff, l, 0&, Buff2, l2, FT)
        SubKeys(i) = Mid(Buff, 1, l)
    Loop Until Rslt <> 0
    i = i - 1
    RegCloseKey h
    On Error GoTo 0
    If i > -1 Then
        ReDim Preserve SubKeys(0 To i)
        For i = 0 To UBound(SubKeys)
            DeleteKey hK, strKey + "\" + SubKeys(i), RaiseErrors:=True
        Next i
    End If
End If
Rslt = RegDeleteKey(hK, strKey)
If Rslt <> 0 Then
  Err.Raise 112, "DeleteKey", "RegDeleteKey failed"
End If
If pDebug Then
    Debug.Print "Key '" + BuildPath(hK, strKey, "") + "' was deleted"
End If
Exit Sub
eh:
PushError
RegCloseKey h
PopError
ErrRaise
End Sub

Private Sub Class_Initialize()
StrBuffer1 = Space$(256)
End Sub

Private Sub Class_Terminate()
StrBuffer1 = vbNullString
End Sub

Public Sub GetAllKeys(ByVal hKey As HKEYS, ByRef SubKey As String, ByRef Keys() As String)
Dim hK As Long
Dim i As Long
Dim FT As FILETIME
Dim l As Long
Dim Class As String, lClass As Long
Dim tmp As String
If RegOpenKeyEx(hKey, SubKey, 0&, KEY_ALL_ACCESS, hK) <> 0 Then
    Err.Raise 113, "GetAllKeys", "Cannot open a key: " + BuildPath(hKey, SubKey, "")
Else
    i = 0
    ReDim Keys(0 To 0)
    tmp = String$(1000, 32)
    l = Len(tmp)
    Class = String$(1000, 32)
    lClass = Len(Class)
    Do While RegEnumKeyEx(hK, i, ByVal tmp, l, 0&, Class, lClass, FT) = 0
        Keys(i) = Left$(tmp, l)
        i = i + 1
        ReDim Preserve Keys(0 To i)
        tmp = String$(1000, 0)
        l = Len(tmp)
        Class = String$(1000, 32)
        lClass = Len(Class)
    Loop
    ReDim Preserve Keys(0 To i - 1)
    
    RegCloseKey hK
End If
End Sub

Public Function KeyExists(ByVal hKey As HKEYS, ByRef SubKey As String) As Boolean
Dim hK As Long
If RegOpenKeyEx(hKey, SubKey, 0&, KEY_ALL_ACCESS, hK) <> 0 Then
    KeyExists = False
Else
    RegCloseKey hK
    KeyExists = True
End If
End Function

Public Function ValueExists(ByVal hKey As HKEYS, ByRef SubKey As String, ByRef Value As String) As Boolean
Dim hK As Long
If RegOpenKeyEx(hKey, SubKey, 0&, KEY_ALL_ACCESS, hK) <> 0 Then
    ValueExists = False
Else
    ValueExists = (RegQueryValueEx(hK, Value, 0&, ByVal 0&, ByVal 0&, ByVal 0&) = 0)
    RegCloseKey hK
End If
End Function

Public Sub DeleteValue(ByVal hKey As HKEYS, ByRef SubKey As String, ByRef Value As String)
Dim hK As Long
Dim Rslt As Long
If RegOpenKeyEx(hKey, SubKey, 0&, KEY_ALL_ACCESS, hK) <> 0 Then
    Err.Raise 113, "DeleteValue", "Cannot open a key: " + BuildPath(hKey, SubKey, "")
Else
    Rslt = RegDeleteValue(hK, Value)
    RegCloseKey hK
    If Rslt <> 0 Then
      Err.Raise 1212, "RegDeleteValue failed!"
    End If
End If

End Sub


'file storage is not currently supported
Public Sub WriteBits(ByVal hKey As HKEYS, _
                     ByRef SubKey As String, _
                     ByRef ValueName As String, _
                     ByRef Data() As Byte, _
                     Optional ByVal UseFilesForLong As Boolean = True, _
                     Optional ByVal RaiseErrors As Boolean = True)
Dim Handle As Long
Dim Disposition As Long
Dim Ret As Long
On Error GoTo eh
If AryDims(AryPtr(Data)) <> 1 Then
    Err.Raise 1111, ""
End If
Ret = RegCreateKeyEx(hKey, _
                    SubKey, _
                    0, _
                    "", _
                    0, _
                    KEY_SET_VALUE, _
                    ByVal 0, _
                    Handle, _
                    Disposition)
If Ret <> 0 Then
    Err.Raise 112, "Reg.WriteBits", "Cannot open a key: '" + BuildPath(hKey, SubKey, ValueName) + "'."
End If

Ret = RegSetValueEx(Handle, ValueName, _
                    0, _
                    REG_BINARY, _
                    ByVal VarPtr(Data(0)), _
                    UBound(Data) + 1)

If Ret <> 0 Then
    Err.Raise 113, "Reg.WriteBits", "Failed to write the value: '" + BuildPath(hKey, SubKey, ValueName) + "'."
End If

RegCloseKey Handle

Exit Sub
eh:
If Handle <> 0 Then
    RegCloseKey Handle
End If
If RaiseErrors Then
    Err.Raise Err.Number, Err.Source, Err.Description
End If
End Sub

Public Sub ReadBits(ByVal hKey As HKEYS, _
                    ByRef SubKey As String, _
                    ByRef ValueName As String, _
                    ByRef Data() As Byte, _
                    Optional ByRef Length As Long, _
                    Optional ByVal UseFilesForLong As Boolean = True, _
                    Optional ByVal RaiseErrors As Boolean = True)
Dim Sz As Long
Dim Typ As eParameterType
Dim Ret As Long
Dim Handle As Long
Length = 0
Ret = RegOpenKeyEx(hKey, SubKey, 0&, KEY_QUERY_VALUE, Handle)
If Ret <> 0 Then
    Err.Raise 112, "Reg.ReadBits", "Cannot open a key: '" + BuildPath(hKey, SubKey, ValueName) + "'."
End If

Ret = RegQueryValueEx(Handle, ValueName, 0&, Typ, ByVal 0&, Sz)
If Ret <> 0 Then
    Err.Raise 112, "Reg.ReadBits", "Cannot query the value: '" + BuildPath(hKey, SubKey, ValueName) + "'."
End If
If Typ <> REG_BINARY Then
    Err.Raise 112, "Reg.ReadBits", "'" + BuildPath(hKey, SubKey, ValueName) + "'. The type of the value is not REG_BINARY!"
End If

If Sz > 0 Then
    ReDim Data(0 To Sz - 1)
    Ret = RegQueryValueEx(Handle, ValueName, 0&, ByVal 0&, Data(0), Sz)
    If Ret <> 0 Then
        Err.Raise 112, "Reg.ReadBits", "Cannot query value data: '" + BuildPath(hKey, SubKey, ValueName) + "'."
    End If
    Length = Sz
Else
    Erase Data
    Length = 0
End If
RegCloseKey Handle
Exit Sub
eh:
If Handle <> 0 Then
    RegCloseKey Handle
End If
If RaiseErrors Then
    Err.Raise Err.Number, Err.Source, Err.Description
End If
End Sub


