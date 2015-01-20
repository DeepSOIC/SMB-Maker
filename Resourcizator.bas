Attribute VB_Name = "Resourcizator"
Option Explicit

'Список свойств, которые надо запоминать
Const PropsToSaveSt As String = "Caption|ToolTipText|EditName"

Dim PropsToSave() As String
Dim nProps As Long
Dim DisableMe As Boolean

Private Sub MakePropsArray()
PropsToSave = Split(PropsToSaveSt, "|")
nProps = UBound(PropsToSave) + 1
End Sub

'Берет форму и создает строку, описывающую ее
Public Function MakeStringFromForm(ByRef frm As Form) As String
Dim Obj As Control
Dim s As String
Dim Index As Long
Dim PropName As String
Dim Accu As String
Dim iProp As Long
Dim Objs() As String, Props() As String, Vals() As String
Dim n As Long
Dim i As Long
Dim LeadingCrLf As String
MakePropsArray
DisableMe = True
On Error Resume Next
For Each Obj In frm
    For iProp = 0 To nProps - 1
        s = ""
        s = CallByName(Obj, PropsToSave(iProp), VbGet)
        If Len(s) > 0 Then
            Index = -1
            Index = Obj.Index
            If n = 0 Then
                ReDim Objs(0 To n), Props(0 To n), Vals(0 To n)
            Else
                ReDim Preserve Objs(0 To n), Props(0 To n), Vals(0 To n)
            End If
            If Index <> -1 Then
                Objs(n) = Obj.Name + "(" + CStr(Index) + ")"
            Else
                Objs(n) = Obj.Name
            End If
            Props(n) = PropsToSave(iProp)
            Vals(n) = s
            n = n + 1
        End If
    Next iProp
Next Obj

Accu = frm.Caption
For i = 0 To n
    Accu = Accu + vbCrLf + "PropertY " + Objs(i) + "." + Props(i) + vbCrLf + Vals(i)
Next i

MakeStringFromForm = Accu
DisableMe = False
End Function

'Берет строку и заполняет названия элементов на форме
Public Sub FillFormUsingString(ByRef St As String, ByRef frm As Form)
Dim i As Long
Dim PPs() As String
Dim DotPos As Long
Dim CrLfPos As Long
Dim NameAndProp As String
Dim Name As String, Prop As String
Dim Obj As Control
Dim OBPos As Long
Dim Index As Long
If DisableMe Then Exit Sub

PPs = Split(St, vbCrLf + "PropertY ")

frm.Caption = PPs(0)
On Error GoTo eh
'Debug.Print ".Caption = " + PPs(0)
For i = 1 To UBound(PPs)
    CrLfPos = InStr(2, PPs(i), vbCrLf)
    NameAndProp = Mid$(PPs(i), 1, CrLfPos - 1)
    DotPos = InStr(1, NameAndProp, ".")
    Debug.Assert DotPos > 1 '"Bad form object's titles string. Please check it!"
    Name = Mid$(NameAndProp, 1, DotPos - 1)
    Prop = Mid$(NameAndProp, DotPos + 1)
    If Right$(Name, 1) = ")" Then
        OBPos = InStr(Name, "(")
        Index = CInt(Mid$(Name, OBPos + 1, Len(Name) - 1 - OBPos))
        On Error GoTo ehpa
        CallByName CallByName(frm, Left$(Name, OBPos - 1), VbGet).Item(Index), Prop, VbLet, Mid$(PPs(i), CrLfPos + 2)
        If False Then
ehpaResume:
            On Error GoTo eh
            CallByName CallByName(frm, Left$(Name, OBPos - 1), VbGet), Prop, VbLet, Mid$(PPs(i), CrLfPos + 2)
        End If
        
    Else
        CallByName CallByName(frm, Name, VbGet), Prop, VbLet, Mid$(PPs(i), CrLfPos + 2)
    End If
    'Debug.Print "." + Name + "." + Prop + " = " + Mid$(PPs(i), CrLfPos + 2)
Next i
Exit Sub
Resume
eh:
MsgBox ("A problem in loading form. Please check the objects." + vbNewLine + vbNewLine + _
        "Form.Name = '" + frm.Name + vbNewLine + "'" + _
        "Name.Prop = '" + NameAndProp + "'")
Debug.Assert False
If Not ProjectCompiled Then Resume Next
Exit Sub
ehpa:
'sometimes, if there is only one item in control array,
'it becomes a control instead of a collection.
Resume ehpaResume
End Sub

'для ресурсов
Public Sub FillFormFromRes(ByRef frm As Form, ByVal ResID As Long)
FillFormUsingString LoadResString(ResID), frm
End Sub
