Attribute VB_Name = "ModuleIcoExtract"
Public Declare Function ExtractIconEx Lib "shell32.dll" Alias "ExtractIconExA" (ByVal lpszFile As String, ByVal nIconIndex As Long, phiconLarge As Long, phiconSmall As Long, ByVal nIcons As Long) As Long
Public Declare Function DestroyIcon Lib "user32" (ByVal hIcon As Long) As Long
Public Declare Function DrawIconEx Lib "user32" (ByVal hDC As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long
'Public Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As SM) As Long


'Public Enum SM
'    SM_CXICONSPACING = 38
'    SM_CYICONSPACING = 39
'    SM_CXICON = 11
'    SM_CYICON = 12
'End Enum

Function PaintIcon(ByVal IconHandle As Long, ByVal hDC As Long, ByVal X As Long, ByVal Y As Long)
DrawIconEx hDC, X, Y, IconHandle, 0, 0, 0, 0, 3
End Function
