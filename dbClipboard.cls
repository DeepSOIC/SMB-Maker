VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "dbClipboard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


Public Property Get cData() As IPictureDisp
Set cData = Clipboard.GetData(CGFormat)
End Property

Public Property Let cData(ByVal vNewValue As IPictureDisp)
Clipboard.SetData vNewValue
End Property

Function CFormat() As VBRUN.ClipBoardConstants
On Error GoTo eh
For tmp = 1 To 15
    If Clipboard.GetFormat(tmp) Then Exit For
Resm:
Next tmp
If tmp = 16 Then CFormat = True Else CFormat = tmp
Exit Function
eh:
Resume Resm
End Function

Function CGFormat() As VBRUN.ClipBoardConstants
Dim tmp As ClipBoardConstants
On Error GoTo eh
tmp = vbCFBitmap
If Clipboard.GetFormat(tmp) Then CGFormat = tmp: Exit Function
tmp = vbCFDIB
If Clipboard.GetFormat(tmp) Then CGFormat = tmp: Exit Function
tmp = vbCFEMetafile
If Clipboard.GetFormat(tmp) Then CGFormat = tmp: Exit Function
tmp = vbCFMetafile
If Clipboard.GetFormat(tmp) Then CGFormat = tmp: Exit Function
Resm:
Exit Function
eh:
Resume Resm
End Function

