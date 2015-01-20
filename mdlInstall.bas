Attribute VB_Name = "mdlInstall"
Option Explicit

Type vtRegFileType
    Extension As String
    DefaultEditor As Boolean
    ReplaceIcon As Boolean
    IconName As String
    MenuCaption As String
    FileDescription As String
End Type


Public Sub RegisterEXE()
With gReg
    On Error Resume Next
        .DeleteKey HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\" + App.EXEName + ".exe"
        .DeleteKey HKEY_CLASSES_ROOT, "Applications\" + App.EXEName + ".exe"
    On Error GoTo 0
    .SetValue HKEY_CLASSES_ROOT, "Applications\" + App.EXEName + ".exe", "", ""
    .SetValue HKEY_CLASSES_ROOT, "Applications\" + App.EXEName + ".exe" + "\shell\open\command", "", """" + ExePath + """" + " %1"
    .SetValue HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\" + App.EXEName + ".exe", "", ExePath
End With
End Sub

Public Sub UnregisterEXE()
With gReg
    On Error Resume Next
        .DeleteKey HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\" + App.EXEName + ".exe"
        .DeleteKey HKEY_CLASSES_ROOT, "Applications\" + App.EXEName + ".exe"
    On Error GoTo 0
End With
End Sub


Public Sub UninstallFileTypes()
Dim Keys() As String
Dim i As Long
gReg.GetAllKeys HKEY_CLASSES_ROOT, "", Keys
'gReg.DeleteKey HKEY_CLASSES_ROOT, "*"
For i = 0 To UBound(Keys)
    If gReg.ValueExists(HKEY_CLASSES_ROOT, Keys(i), "SMBMaker.rem") Then
        gReg.DeleteKey HKEY_CLASSES_ROOT, Keys(i)
    Else
        If gReg.KeyExists(HKEY_CLASSES_ROOT, Keys(i) + "\Shell\" + AppTitle) Then
            gReg.DeleteKey HKEY_CLASSES_ROOT, Keys(i) + "\Shell\" + AppTitle
            If gReg.ValueExists(HKEY_CLASSES_ROOT, Keys(i) + "\Shell", "SMBMaker.bak") Then
                gReg.SetValue HKEY_CLASSES_ROOT, Keys(i) + "\Shell", "", gReg.GetValue(HKEY_CLASSES_ROOT, Keys(i) + "\Shell", "SMBMaker.bak")
                gReg.DeleteValue HKEY_CLASSES_ROOT, Keys(i) + "\Shell", "SMBMaker.bak"
            End If
            If gReg.ValueExists(HKEY_CLASSES_ROOT, Keys(i) + "\DefaultIcon", "SMBMaker.bak") Then
                gReg.SetValue HKEY_CLASSES_ROOT, Keys(i) + "\DefaultIcon", "", gReg.GetValue(HKEY_CLASSES_ROOT, Keys(i) + "\DefaultIcon", "SMBMaker.bak")
                gReg.DeleteValue HKEY_CLASSES_ROOT, Keys(i) + "\DefaultIcon", "SMBMaker.bak"
            End If
        End If
    End If
Next i

UnregisterEXE

End Sub

Sub SetFileType(ByRef ext As vtRegFileType, ByRef strLog As String)
Dim tmp As String, tKeyName As String, bln As Boolean
Dim l As String, i As Integer, tmpI As String ', AppTitle As String
l = strLog + vbCrLf + GRSF(1222) '"Initializing..."
On Error GoTo eh
bln = True
    l = l + vbCrLf + grs(1223, "ext", ext.Extension) + vbCrLf + _
            GRSF(1224) '"Begin " + Chr(34) + Ext.Extension + Chr(34) + vbCrLf + _
            "Trying to get info... "
    Err.Clear
    tmp = gReg.GetValue(HKEY_CLASSES_ROOT, ext.Extension, "", False)
    If Not tmp = "" Then
        If Not gReg.KeyExists(HKEY_CLASSES_ROOT, tmp) Then
            gReg.SetValue HKEY_CLASSES_ROOT, tmp, "SMBMaker.rem", "1"
        End If
        l = l + grs(1225, "|1", tmp) + vbCrLf + _
            GRSF(1226) '"  Success. Result = " + Chr(34) + tmp + Chr(34) + vbCrLf + _
            "Creating item..."
        gReg.SetValue HKEY_CLASSES_ROOT, tmp + "\shell\" + AppTitle, "", ext.MenuCaption
        If Err.Number <> 0 Then l = l + GRSF(1227): Err.Clear Else l = l + GRSF(1234) '"   An error occured" | "   Success"
        
        l = l + vbCrLf + GRSF(1228) '"Creating command..."
        gReg.SetValue HKEY_CLASSES_ROOT, tmp + "\shell\" + AppTitle + "\Command", "", Chr$(34) + AppPath + App.EXEName + ".exe" + Chr$(34) + " %1"
        If Err.Number <> 0 Then l = l + GRSF(1227): Err.Clear Else l = l + GRSF(1234) '"   An error occured" | "   Success"
        
        If ext.DefaultEditor Then
            l = l + vbCrLf + GRSF(1229) '"Applying Default action..."
            gReg.SetValue HKEY_CLASSES_ROOT, tmp + "\Shell", "SMBMaker.bak", gReg.GetValue(HKEY_CLASSES_ROOT, tmp + "\Shell", "", False)
            gReg.SetValue HKEY_CLASSES_ROOT, tmp + "\Shell", "", AppTitle
            If Err.Number <> 0 Then l = l + GRSF(1227): Err.Clear Else l = l + GRSF(1234) '"   An error occured" | "   Success"
        End If
        
        If ext.ReplaceIcon Then
            l = l + vbCrLf + GRSF(1230) '"Updating icon..."
            gReg.SetValue HKEY_CLASSES_ROOT, tmp + "\DefaultIcon", "SMBMaker.bak", gReg.GetValue(HKEY_CLASSES_ROOT, tmp + "\DefaultIcon", "", False)
            gReg.SetValue HKEY_CLASSES_ROOT, tmp + "\DefaultIcon", "", ext.IconName
        Else
            l = l + vbCrLf + GRSF(1231) '"Searching for icon..."
            On Error Resume Next
            tmpI = gReg.GetValue(HKEY_CLASSES_ROOT, tmp + "\DefaultIcon", "")
            On Error GoTo eh
            If tmpI = "" Then
                l = l + GRSF(1232) '"   No icon found. Creating..."
                gReg.SetValue HKEY_CLASSES_ROOT, tmp + "\DefaultIcon", "", ext.IconName
            End If
        End If
        
        If Err.Number <> 0 Then l = l + GRSF(1227): Err.Clear Else l = l + GRSF(1234) '"   An error occured" | "   Success"
    ElseIf Err.Number = 0 Then
        l = l + vbCrLf + GRSF(1233) '"No info found. Creating..."
        tKeyName = Mid(ext.Extension, 2, Len(ext.Extension) - 1) & "_File"
        gReg.SetValue HKEY_CLASSES_ROOT, ext.Extension, "", tKeyName
        gReg.SetValue HKEY_CLASSES_ROOT, tKeyName, "", GetFileDescription(ext.FileDescription, ext.Extension)
        gReg.SetValue HKEY_CLASSES_ROOT, tKeyName, "SMBMaker.rem", "1"
        If Err.Number <> 0 Then l = l + GRSF(1227): Err.Clear Else l = l + GRSF(1234) '"   An error occured" | "   Success"
       
        l = l + vbCrLf + GRSF(1230) '"Creating icon reference..."
        gReg.SetValue HKEY_CLASSES_ROOT, tKeyName + "\DefaultIcon", "", ext.IconName
        If Err.Number <> 0 Then l = l + GRSF(1227): Err.Clear Else l = l + GRSF(1234) '"   An error occured" | "   Success"
        
        l = l + vbCrLf + GRSF(1229) '"Creating item..."
        gReg.SetValue HKEY_CLASSES_ROOT, tKeyName & "\Shell", "", AppTitle
        If Err.Number <> 0 Then l = l + GRSF(1227): Err.Clear Else l = l + GRSF(1234) '"   An error occured" | "   Success"
        
        l = l + vbCrLf + GRSF(1226) '"Creating item..."
        gReg.SetValue HKEY_CLASSES_ROOT, tKeyName & "\Shell\" + AppTitle, "", ext.MenuCaption
        If Err.Number <> 0 Then l = l + GRSF(1227): Err.Clear Else l = l + GRSF(1234) '"   An error occured" | "   Success"
        
        l = l + vbCrLf + GRSF(1228) '"Creating command..."
        gReg.SetValue HKEY_CLASSES_ROOT, tKeyName & "\Shell\" + AppTitle + "\Command", "", Chr(34) + AppPath + App.EXEName + ".exe" + Chr(34) + " %1"
        If Err.Number <> 0 Then l = l + GRSF(1227): Err.Clear Else l = l + GRSF(1234) '"   An error occured" | "   Success"
        
    End If
strLog = l
Exit Sub
eh:
bln = False
If MsgBox("An error occured. View info?", vbYesNo, "Error") = vbYes Then
dbLongMsgBox "tmp=" + tmp + "; ext=" + ext.Extension + "; log:" + vbCrLf + l, "Log"
End If
Resume Next

End Sub


Function GetFileDescription(ByVal Desc As String, ByVal ext As String) As String
'Dim i As Long, strRes As String
If Len(Desc) > 0 Then
    GetFileDescription = Desc
Else
    GetFileDescription = grs(1500, "$ext$", ext)
End If
'ext = LCase(ext)
'If Mid(ext, 1, 2) = "*." Then
'    ext = Mid(ext, 3, Len(ext) - 2)
'ElseIf Mid(ext, 1, 1) = "." Then
'    ext = Mid(ext, 2, Len(ext) - 1)
'End If
'If (InStr(1, ext, ".") > 0) Or (ext = "") Then
'    Err.Raise 5, "Resource", "Illegal extension: " + ext + "."
'    Exit Function
'End If
'Do
'    strRes = GRSF(1501 + i)
'    i = i + 1
'    If Not (strRes = "<EOL>") Then
'        If Mid(strRes, 1, InStr(1, strRes, "=") - 1) = ext Then
'            strRes = Mid(strRes, Len(ext) + 2, Len(strRes) - Len(ext) - 1)
'            GetFileDescription = strRes
'            Exit Function
'        End If
'    End If
'Loop Until strRes = "<EOL>"
'GetFileDescription = GRSF(1500) '"Picture file"
End Function

