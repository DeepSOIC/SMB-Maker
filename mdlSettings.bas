Attribute VB_Name = "mdlSettings"
Option Explicit

Public PrePathCommon As String
Public PrePathLocal As String
Public UseRegistryForSettings As Boolean
Public SettingsStorage As clsIniFile
Dim Initd As Boolean

Public Sub InitializeSettings()
If Not Initd Then
    Set SettingsStorage = New clsIniFile
    Initd = True
    On Error GoTo eh
    SettingsStorage.DefFile
    SettingsStorage.LoadFile
End If
Exit Sub
eh:
MsgError
End Sub

Public Sub FlushSettings()
If Not Initd Then Exit Sub
On Error Resume Next
SettingsStorage.SaveFile
End Sub

Public Sub dbSaveSetting(ByRef Key As String, ByRef Parameter As String, _
                                              ByRef Value As String, _
                         Optional ByVal CommonSetting As Boolean = False, _
                         Optional ByVal AllUsers As Boolean = False, _
                         Optional ByVal ForceRegistry As Boolean = False)
Dim hKey As HKEYS
Dim PrePath As String
If UseRegistryForSettings Or ForceRegistry Then
    If AllUsers Then
        hKey = HKEY_LOCAL_MACHINE
    Else
        hKey = HKEY_CURRENT_USER
    End If
    If Len(PrePathCommon) = 0 Then
        LoadRegPaths
    End If
    If CommonSetting Then
        PrePath = PrePathCommon + Key '"Software\Dbnz\SMBMaker\Common\"
    Else
        PrePath = PrePathLocal + Key '"Software\Dbnz\SMBMaker\" + CStr(App.Revision) + "\"
    End If
    On Error GoTo eh
    gReg.SetValue hKey, PrePath, Parameter, Value
Else
    InitializeSettings
    SettingsStorage.SetSetting Key, Parameter, Value
End If
Exit Sub
eh:
If Err.Number = 91 Then
    Set gReg = New Reg
    Resume
End If
End Sub

Public Sub dbSaveSettingEx(ByRef Key As String, ByRef Parameter As String, _
                           ByRef Value As Variant, _
                           Optional ByVal VersionsCommon As Boolean = False, _
                           Optional ByVal AllUsers As Boolean = False, _
                           Optional ByVal HexNumber As Boolean = False, _
                           Optional ByVal ForceRegistry As Boolean = False)
Dim tmp As String
If (VarType(Value) And vbArray) <> 0 Then
    MsgBox "Direct saving of arrays is not currently supported. " + vbCrLf + _
           "Attempt: " + Key + "::" + Parameter, vbCritical, "Type mismatch"
    Exit Sub
End If

Select Case VarType(Value)
    Case VbVarType.vbBoolean
    Case VbVarType.vbByte, VbVarType.vbInteger, VbVarType.vbLong 'Integers
    Case VbVarType.vbCurrency, VbVarType.vbDecimal, VbVarType.vbDouble, VbVarType.vbSingle 'other numbers
    Case VbVarType.vbString
    Case Else
        MsgBox "Direct saving of '" + TypeName(Value) + "'s is not currently supported. " + vbCrLf + _
               "Attempt: " + Key + "::" + Parameter, vbCritical, "Type mismatch"
        Exit Sub
End Select
dbSaveSetting Key, Parameter, dbCStr(Value), VersionsCommon, AllUsers, ForceRegistry
End Sub

Public Sub dbSaveSettingBin(ByRef Key As String, _
                            ByRef Parameter As String, _
                            ByRef Data() As Byte, _
                            Optional ByVal CommonSetting As Boolean = False, _
                            Optional ByVal AllUsers As Boolean = False, _
                            Optional ByVal ForceRegistry As Boolean = False)
Dim hKey As HKEYS
Dim PrePath As String
Dim St As String
If UseRegistryForSettings Or ForceRegistry Then
    If AllUsers Then
        hKey = HKEY_LOCAL_MACHINE
    Else
        hKey = HKEY_CURRENT_USER
    End If
    If Len(PrePathCommon) = 0 Then
        LoadRegPaths
    End If
    If CommonSetting Then
        PrePath = PrePathCommon + Key '"Software\Dbnz\SMBMaker\Common\"
    Else
        PrePath = PrePathLocal + Key '"Software\Dbnz\SMBMaker\" + CStr(App.Revision) + "\"
    End If
    On Error GoTo eh
    gReg.WriteBits hKey, PrePath, Parameter, Data
Else
    InitializeSettings
    St = Data
    dbSaveSetting Key, Parameter, StrConv(St, vbUnicode)
End If
Exit Sub
eh:
If Err.Number = 91 Then
    Set gReg = New Reg
    Resume
End If
End Sub

Public Sub LoadRegPaths()
    
    PrePathCommon = "Software\Dbnz\SMBMaker\Common\"
    PrePathLocal = "Software\Dbnz\SMBMaker\" + CStr(App.Revision) + "\"

End Sub

Public Function dbGetSetting(ByRef Key As String, ByRef Parameter As String, _
                             Optional ByRef DefValue As String = vbNullString, _
                             Optional ByVal CommonSetting As Boolean = False, _
                             Optional ByVal AllUsers As Boolean = False, _
                             Optional ByVal ForceRegistry As Boolean = False) As String
Dim hKey As HKEYS
Dim PrePath As String
Dim tmp As String
If UseRegistryForSettings Or ForceRegistry Then
    If AllUsers Then
        hKey = HKEY_LOCAL_MACHINE
    Else
        hKey = HKEY_CURRENT_USER
    End If
    If Len(PrePathCommon) = 0 Then
        LoadRegPaths
    End If
    If CommonSetting Then
        PrePath = PrePathCommon + Key '"Software\" + App.CompanyName + "\" + AppTitle + "\Common\"
    Else
        PrePath = PrePathLocal + Key '"Software\" + App.CompanyName + "\" + AppTitle + "\" + CStr(App.Revision) + "\"
    End If
    On Error GoTo eh
    tmp = gReg.GetValue(hKey, PrePath, Parameter, True)
    On Error GoTo 0
    dbGetSetting = tmp
Else
    InitializeSettings
    If SettingsStorage.QuerySetting(Key, Parameter, tmp) Then
        dbGetSetting = tmp
    Else
        dbGetSetting = DefValue
    End If
End If
Exit Function
eh:
If Err.Number = 91 Then
    Set gReg = New Reg
    Resume
ElseIf Err.Number = 112 Then
    tmp = DefValue
    'Debug.Print Err.Source, Err.Description
    Resume Next
End If
End Function

Public Function dbGetSettingEx(ByRef Key As String, ByRef Parameter As String, _
                          Optional ByVal TypeID As VbVarType = 0, _
                          Optional ByRef DefValue As Variant, _
                          Optional ByVal VersionsCommon As Boolean = False, _
                          Optional ByVal AllUsers As Boolean = False, _
                          Optional ByVal HexNumber As Boolean = False, _
                          Optional ByVal ForceRegistry As Boolean = False) As Variant
Dim tmp As String
If TypeID = 0 Then TypeID = IIf(IsMissing(DefValue), vbString, VarType(DefValue))
If (TypeID And vbArray) <> 0 Then
    MsgBox "Direct loading of arrays is not currently supported. " + vbCrLf + _
           "Attempt: " + Key + "::" + Parameter, vbCritical, "Type mismatch"
    Exit Function
End If
tmp = dbGetSetting(Key, Parameter, vbNullChar, VersionsCommon, AllUsers, ForceRegistry)
If tmp = vbNullChar Then
    dbGetSettingEx = DefValue
Else
    Select Case TypeID
        Case VbVarType.vbBoolean
        Case VbVarType.vbByte
        Case VbVarType.vbInteger
        Case VbVarType.vbLong 'Integers
        Case VbVarType.vbCurrency, VbVarType.vbDecimal, VbVarType.vbDouble, VbVarType.vbSingle 'other numbers
        Case VbVarType.vbString
            dbGetSettingEx = tmp
            Exit Function
        Case VbVarType.vbUserDefinedType
            MsgBox "Direct loading of '" + dbCStr(TypeID) + "'s is not currently supported. " + vbCrLf + _
                   "Attempt: " + Key + "::" + Parameter, vbCritical, "Type mismatch"
            Exit Function
    End Select
    dbGetSettingEx = dbVal(tmp, TypeID)
End If
End Function

Public Function dbGetSettingBin(ByRef Key As String, _
                                ByRef Parameter As String, _
                                ByRef Data() As Byte, _
                                Optional ByVal CommonSetting As Boolean = False, _
                                Optional ByVal AllUsers As Boolean = False, _
                                Optional ByVal ForceRegistry As Boolean = False) As Long
Dim hKey As HKEYS
Dim PrePath As String
Dim tmp As String
Dim cb As Long
If UseRegistryForSettings Or ForceRegistry Then
    If AllUsers Then
        hKey = HKEY_LOCAL_MACHINE
    Else
        hKey = HKEY_CURRENT_USER
    End If
    If Len(PrePathCommon) = 0 Then
        LoadRegPaths
    End If
    If CommonSetting Then
        PrePath = PrePathCommon + Key '"Software\" + App.CompanyName + "\" + AppTitle + "\Common\"
    Else
        PrePath = PrePathLocal + Key '"Software\" + App.CompanyName + "\" + AppTitle + "\" + CStr(App.Revision) + "\"
    End If
    On Error GoTo eh
    gReg.ReadBits hKey, PrePath, Parameter, Data, cb
    On Error GoTo 0
    dbGetSettingBin = cb
Else
    InitializeSettings
    If SettingsStorage.QuerySetting(Key, Parameter, tmp) Then
        dbGetSettingBin = Len(tmp)
        Data = StrConv(tmp, vbFromUnicode)
    Else
        dbGetSettingBin = 0
        Erase Data
    End If
End If
Exit Function
eh:
If Err.Number = 91 Then
    Set gReg = New Reg
    Resume
ElseIf Err.Number = 112 Then
    'Debug.Print Err.Source, Err.Description
    Resume Next
End If
End Function


Public Sub dbDeleteSetting(ByVal Key As String, _
                           Optional ByRef Parameter As String = "", _
                           Optional ByVal CommonSetting As Boolean = False, _
                           Optional ByVal AllUsers As Boolean = False, _
                           Optional ByVal ForceRegistry As Boolean = False)
Dim hKey As HKEYS
Dim PrePath As String
Dim tmp As String
Dim Ret As Boolean
If UseRegistryForSettings Or ForceRegistry Then
    If AllUsers Then
        hKey = HKEY_LOCAL_MACHINE
    Else
        hKey = HKEY_CURRENT_USER
    End If
    If CommonSetting Then
        PrePath = PrePathCommon + Key '"Software\" + App.CompanyName + "\" + AppTitle + "\Common\"
    Else
        PrePath = PrePathLocal + Key '"Software\" + App.CompanyName + "\" + AppTitle + "\" + CStr(App.Revision) + "\"
    End If
    If Key = "" Then
        PrePath = Left$(PrePath, Len(PrePath) - 1)
    End If
    On Error GoTo eh
    If Len(Parameter) = 0 Then
        gReg.DeleteKey hKey, PrePath, True
    Else
        gReg.DeleteValue hKey, PrePath, Parameter
    End If
    On Error GoTo 0
Else
    InitializeSettings
    If Len(Key) = 0 Then
        Set SettingsStorage = Nothing
        Set SettingsStorage = New clsIniFile
        SettingsStorage.DefFile
    Else
        If Len(Parameter) = 0 Then
            Ret = SettingsStorage.DeleteSection(Key)
        Else
            Ret = SettingsStorage.DeleteSetting(Key, Parameter)
        End If
    End If
End If
Exit Sub
eh:
If Err.Number = 91 Then
    Set gReg = New Reg
    Resume
ElseIf Err.Number = 112 Then
    Debug.Print Err.Source, Err.Description
    Resume Next
End If
End Sub

Public Function dbSettingPresent(ByVal Key As String, _
                                 Optional ByVal Parameter As String = "", _
                                 Optional ByVal CommonSetting As Boolean = False, _
                                 Optional ByVal AllUsers As Boolean = False, _
                                 Optional ByVal ForceRegistry As Boolean = False)
Dim hKey As HKEYS
Dim PrePath As String
Dim tmp As String
If UseRegistryForSettings Or ForceRegistry Then
    If AllUsers Then
        hKey = HKEY_LOCAL_MACHINE
    Else
        hKey = HKEY_CURRENT_USER
    End If
    If CommonSetting Then
        PrePath = PrePathCommon + Key '"Software\" + App.CompanyName + "\" + AppTitle + "\Common\"
    Else
        PrePath = PrePathLocal + Key '"Software\" + App.CompanyName + "\" + AppTitle + "\" + CStr(App.Revision) + "\"
    End If
    On Error GoTo eh
    If Len(Parameter) = 0 Then
        dbSettingPresent = gReg.KeyExists(hKey, PrePath)
    Else
        dbSettingPresent = gReg.ValueExists(hKey, PrePath, Parameter)
    End If
    On Error GoTo 0
Else
    If Len(Parameter) = 0 Then
        dbSettingPresent = SettingsStorage.SectionPresent(Key)
    Else
        dbSettingPresent = SettingsStorage.SettingPresent(Key, Parameter)
    End If
End If
Exit Function
eh:
If Err.Number = 91 Then
    Set gReg = New Reg
    Resume
ElseIf Err.Number = 112 Then
    'tmp = DefValue
    'Debug.Print Err.Source, Err.Description
    Resume Next
End If
End Function


'returns the number of strings returned.
Public Function dbGetAllSettings(ByRef Key As String, _
                                 ByRef Params() As String, _
                                 ByRef Vals() As String, _
                                 Optional ByVal CommonSetting As Boolean = False, _
                                 Optional ByVal AllUsers As Boolean = False, _
                                 Optional ByVal ForceRegistry As Boolean = False) As Long
Dim hKey As HKEYS
Dim PrePath As String
Dim tmp As String
If UseRegistryForSettings Or ForceRegistry Then
    If AllUsers Then
        hKey = HKEY_LOCAL_MACHINE
    Else
        hKey = HKEY_CURRENT_USER
    End If
    If CommonSetting Then
        PrePath = PrePathCommon + Key '"Software\" + App.CompanyName + "\" + AppTitle + "\Common\"
    Else
        PrePath = PrePathLocal + Key '"Software\" + App.CompanyName + "\" + AppTitle + "\" + CStr(App.Revision) + "\"
    End If
    On Error GoTo eh
    dbGetAllSettings = gReg.GetAllValues(hKey, PrePath, Params, Vals)
Else
    InitializeSettings
    On Error GoTo eh2
    dbGetAllSettings = SettingsStorage.GetAllSettings(Key, Params, Vals)
    On Error GoTo 0
End If
Exit Function
eh:
If Err.Number = 91 Then
    Set gReg = New Reg
    Resume
ElseIf Err.Number = 112 Then
    'tmp = DefValue
    'Debug.Print Err.Source, Err.Description
    Resume Next
End If

eh2:
    Erase Params, Vals
    dbGetAllSettings = 0
    Resume Next
End Function


