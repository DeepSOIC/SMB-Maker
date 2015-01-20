VERSION 5.00
Begin VB.UserControl ctlColor 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   Begin VB.CommandButton cmdSel 
      BackColor       =   &H00FFFFFF&
      Height          =   1740
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   120
      Width           =   1545
   End
End
Attribute VB_Name = "ctlColor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim prvColor As Long 'this color is always in RGB0 format
Dim En As Boolean
Private Declare Function GetWindowDC Lib "user32" _
            (ByVal hWnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" _
            (ByVal hWnd As Long, _
            ByVal hDC As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

Private Type POINTAPI
        X As Long
        Y As Long
End Type


Public ColorFormat As eColorFormat
Public DisableNextChange As Boolean

Public Enum eColorFormat
    [RGB0 Model] = 0
    [BGR0 Model] = 1
End Enum

Public Event Change()
Public Event Click()
Public Event MouseDown(ByVal Button As Long, ByVal Shift As Long, ByVal X As Long, ByVal Y As Long)
Public Event MouseMove(ByVal Button As Long, ByVal Shift As Long, ByVal X As Long, ByVal Y As Long)
Public Event MouseUp(ByVal Button As Long, ByVal Shift As Long, ByVal X As Long, ByVal Y As Long)

Public Function ConvertColorLng(ByVal lngColor As Long) As Long
ConvertColorLng = (lngColor And &HFF00&) Or _
                  (lngColor And &HFF&) * &H10000 Or _
                  (lngColor And &HFF0000) \ &H10000
End Function

Private Sub cmdSel_Click()
Dim CDl As New CommonDlg
On Error GoTo eh
RaiseEvent Click
prvColor = CDl.PickColor(prvColor, True, False)
RefreshColor
pChange
Exit Sub
eh:
If Err.Number = dbCWS Then Exit Sub
MsgError
End Sub

Private Function MsgError(Optional ByVal Style As VbMsgBoxStyle = vbCritical)
MsgBox Err.Description, Style, Err.Source
Debug.Assert False
End Function


Public Property Get Color() As OLE_COLOR
Attribute Color.VB_ProcData.VB_Invoke_Property = ";Data"
Attribute Color.VB_UserMemId = 0
If ColorFormat = [BGR0 Model] Then
    Color = ConvertColorLng(prvColor)
Else
    Color = prvColor
End If
End Property

Public Property Let Color(ByVal NewColor As OLE_COLOR)
Dim tColor As Long
If ColorFormat = [BGR0 Model] Then
    tColor = ConvertColorLng(NewColor)
Else
    tColor = NewColor
End If
If tColor <> prvColor Then
    prvColor = tColor
    pChange
End If
DisableNextChange = False
RefreshColor
End Property

Private Sub pChange()
If Not DisableNextChange Then RaiseEvent Change
DisableNextChange = False
End Sub

Public Sub RefreshColor()
cmdSel.BackColor = prvColor
End Sub

Private Sub cmdSel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub cmdSel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim Pnt As POINTAPI
If Button = 2 Then
    GetCursorPos Pnt
    cmdSel.BackColor = CapturePixel(Pnt.X, Pnt.Y)
End If
RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub cmdSel_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim Pnt As POINTAPI
If Button = 2 Then
    GetCursorPos Pnt
    prvColor = CapturePixel(Pnt.X, Pnt.Y)
    RefreshColor
    pChange
End If
RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Public Function CapturePixel(ByVal px As Long, ByVal py As Long) As Long
Dim WDC As Long
WDC = GetWindowDC(0&)
CapturePixel = GetPixel(WDC, px, py)
ReleaseDC 0&, WDC
End Function


Private Sub UserControl_InitProperties()
ColorFormat = [BGR0 Model]
Color = &H0&
Enabled = True
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
With PropBag
    ColorFormat = .ReadProperty("ColorFormat", [BGR0 Model])
    DisableNextChange = True
    Color = .ReadProperty("Color", &H0&)
    Enabled = .ReadProperty("Enabled", True)
End With
End Sub

Private Sub UserControl_Resize()
cmdSel.Move 0, 0, ScaleWidth, ScaleHeight
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
With PropBag
    .WriteProperty "ColorFormat", ColorFormat, [BGR0 Model]
    .WriteProperty "Color", Color, &H0&
    .WriteProperty "Enabled", Enabled, True
End With
End Sub

Public Property Get Enabled() As Boolean
Enabled = En
End Property

Public Property Let Enabled(ByVal nEn As Boolean)
En = nEn
cmdSel.Enabled = En
End Property
