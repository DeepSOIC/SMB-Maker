Attribute VB_Name = "mdlRects"
Option Explicit

Public Function IsRectEmpty(ByRef Rct As RECT) As Boolean
IsRectEmpty = Not ((Rct.Right > Rct.Left) And (Rct.Bottom > Rct.Top))
End Function

Public Function UnionRects(ByRef Rect1 As RECT, ByRef Rect2 As RECT) As RECT
If IsRectEmpty(Rect1) Then
    If Not IsRectEmpty(Rect2) Then
        UnionRects = Rect2
    End If
Else
    If IsRectEmpty(Rect2) Then
        UnionRects = Rect1
    Else
        UnionRects.Left = Min(Rect1.Left, Rect2.Left)
        UnionRects.Top = Min(Rect1.Top, Rect2.Top)
        UnionRects.Right = Max(Rect1.Right, Rect2.Right)
        UnionRects.Bottom = Max(Rect1.Bottom, Rect2.Bottom)
    End If
End If
End Function

