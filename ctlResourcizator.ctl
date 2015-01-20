VERSION 5.00
Begin VB.UserControl ctlResourcizator 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   300
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2550
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   204
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   20
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   170
End
Attribute VB_Name = "ctlResourcizator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim pFormResID As Integer

Public Sub LoadCaptions()
Attribute LoadCaptions.VB_Description = "Call this method to fill the form using the string in the resource pointed by ResID."
If pFormResID <> 0 Then
    FillFormFromRes UserControl.Parent, pFormResID
Else
    Err.Raise 1212, UserControl.Name + ":LoadCaptions", "ResID property not set."
End If
End Sub

Public Property Get MakeStringNow() As Boolean
Attribute MakeStringNow.VB_Description = "Set to true to create a string describing your form. Will be copied to clipboard."
MakeStringNow = False
End Property

Public Property Let MakeStringNow(ByVal vNewValue As Boolean)
If vNewValue Then
    Clipboard.SetText MakeStringFromForm(UserControl.Parent)
    MsgBox "The string has been copied to clipboard. Please past it into the resource editor to the string number: " + vbNewLine + CStr(ResID)
End If
End Property



Public Property Get FillForm() As Boolean
Attribute FillForm.VB_Description = "Set to True to fill the form using the string currently in the clipboard."
FillForm = False
End Property

Private Function Min(ByVal a As Long, ByVal B As Long) As Long
If a > B Then Min = B Else Min = a
End Function

Public Property Let FillForm(ByVal vNewValue As Boolean)
Dim Txt As String
If vNewValue Then
    If Clipboard.GetFormat(ClipBoardConstants.vbCFText) Then
        Txt = Clipboard.GetText
        If MsgBox("The string contains:" + vbNewLine + _
                  Left(Txt, Min(100, Len(Txt))) + vbNewLine + _
                  "Are you sure you want to load it?", vbYesNo) = vbNo Then
            Exit Property
        End If
        If Len(Txt) > 0 Then
            FillFormUsingString Txt, UserControl.Parent
        End If
    End If
End If
End Property

Public Property Get ResID() As Integer
Attribute ResID.VB_Description = "The resource identifier for the string describing the form."
Attribute ResID.VB_UserMemId = 0
ResID = pFormResID
End Property

Public Property Let ResID(ByVal vNewValue As Integer)
pFormResID = vNewValue
End Property

Private Sub UserControl_AmbientChanged(PropertyName As String)
If UCase$(PropertyName) = UCase$("DisplayName") Then
    On Error Resume Next
    UserControl_Resize
End If
End Sub

Private Sub UserControl_InitProperties()
pFormResID = 0
If Not TypeOf Parent Is Form Then
    Err.Raise 12321, UserControl.Name, "This control can be placed on form object only!"
End If
End Sub

Private Sub UserControl_Paint()
Dim Txt As String
Dim tw As Long, th As Long
Txt = Ambient.DisplayName
tw = UserControl.TextWidth(Txt)
th = UserControl.TextHeight(Txt)
CurrentX = (ScaleWidth - tw) \ 2
CurrentY = (ScaleHeight - th) \ 2
UserControl.Print Txt
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
With PropBag
    pFormResID = .ReadProperty("ResID", 0)
End With
End Sub

Private Sub UserControl_Resize()
Dim Txt As String
Static bLock As Boolean
If Ambient.UserMode Then Exit Sub
If bLock Then Exit Sub
On Error Resume Next
If TypeOf Parent Is Form Then
Else
    Err.Raise 12321, UserControl.Name, "This control can be placed on form object only!"
End If
Txt = Ambient.DisplayName
'Extender.Move Extender.Left, Extender.Top, _
'                 ScaleX(TextWidth(Txt), ScaleMode, Parent.ScaleMode), _
'                 ScaleY(TextHeight(Txt), ScaleMode, Parent.ScaleMode)
bLock = True
Width = ScaleX(ScaleX(TextWidth(Txt), ScaleMode, vbPixels) + 4 + 10, vbPixels, vbTwips)
Height = ScaleY(ScaleY(TextHeight(Txt), ScaleMode, vbPixels) + 4 + 10, vbPixels, vbTwips)
bLock = False
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
With PropBag
    .WriteProperty "ResID", pFormResID, 0
End With
End Sub
