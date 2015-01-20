Attribute VB_Name = "mdlSorting"
Option Explicit
Dim CurAry() As Long, CurLUT() As Long


Public Sub GenerateLUT(ByRef Ary() As Long, ByRef LUT() As Long)
Dim l As Long
Dim i As Long
Dim LB As Long, UB As Long
l = AryLen(AryPtr(Ary))
If l = 0 Then Exit Sub
LB = LBound(Ary): UB = UBound(Ary)
SwapArys AryPtr(Ary), AryPtr(CurAry)
SwapArys AryPtr(LUT), AryPtr(CurLUT)
On Error GoTo eh

ReDim CurLUT(LB To UB)
For i = LB To UB
    CurLUT(i) = i
Next i

recGenerateLUT LB, UB

SwapArys AryPtr(Ary), AryPtr(CurAry)
SwapArys AryPtr(LUT), AryPtr(CurLUT)
Erase CurAry
Erase CurLUT
Exit Sub
eh:
    SwapArys AryPtr(Ary), AryPtr(CurAry)
    SwapArys AryPtr(LUT), AryPtr(CurLUT)
ErrRaise "GenerateLUT"
End Sub

Private Sub recGenerateLUT(ByVal lngLeft As Long, _
                          ByVal lngRight As Long)
    Dim i As Long
    Dim j As Long
    Dim lngTestVal As Long
    Dim lngMid As Long
    Dim tmp As Long
   
    If lngLeft < lngRight Then
        lngMid = (lngLeft + lngRight) \ 2
        lngTestVal = CurAry(CurLUT(lngMid))
        i = lngLeft
        j = lngRight
        Do
            Do While CurAry(CurLUT(i)) < lngTestVal
                i = i + 1
            Loop
            Do While CurAry(CurLUT(j)) > lngTestVal
                j = j - 1
            Loop
            If i <= j Then
                tmp = CurLUT(i)
                CurLUT(i) = CurLUT(j)
                CurLUT(j) = tmp
                i = i + 1
                j = j - 1
            End If
        Loop Until i > j
        ' To optimize the sort, always sort the
        ' smallest segment first.
        If j <= lngMid Then
            Call recGenerateLUT(lngLeft, j)
            Call recGenerateLUT(i, lngRight)
        Else
            Call recGenerateLUT(i, lngRight)
            Call recGenerateLUT(lngLeft, j)
        End If
    End If

End Sub




Public Sub SortLongArray(ByRef vArr() As Long, ByVal lngLeft As Long, ByVal lngRight As Long)
    Dim i As Long
    Dim j As Long
    Dim lngTestVal As Long
    Dim lngMid As Long

    'If lngLeft = dhcMissing Then lngLeft = LBound(varr)
    'If lngRight = dhcMissing Then lngRight = UBound(varr)
   
    If lngLeft < lngRight Then
        lngMid = (lngLeft + lngRight) \ 2
        lngTestVal = vArr(lngMid)
        i = lngLeft
        j = lngRight
        Do
            Do While (vArr(i) < lngTestVal)
                i = i + 1
            Loop
            Do While (vArr(j) > lngTestVal)
                j = j - 1
            Loop
            If i <= j Then
                SwapLongs vArr(i), vArr(j)
                i = i + 1
                j = j - 1
            End If
        Loop Until i > j
        ' To optimize the sort, always sort the
        ' smallest segment first.
        If j <= lngMid Then
            Call SortLongArray(vArr, lngLeft, j)
            Call SortLongArray(vArr, i, lngRight)
        Else
            Call SortLongArray(vArr, i, lngRight)
            Call SortLongArray(vArr, lngLeft, j)
        End If
    End If

End Sub

Private Sub SwapLongs(ByRef a As Long, ByRef b As Long)
Dim c As Long
c = a
a = b
b = c
End Sub

'Public Sub SortCombo(ByRef cmb As ComboBox, ByVal lngLeft As Long, ByVal lngRight As Long)
'    Dim i As Long
'    Dim j As Long
'    Dim strTestVal As String
'    Dim lngMid As Long
'
'    'If lngLeft = dhcMissing Then lngLeft = LBound(varr)
'    'If lngRight = dhcMissing Then lngRight = UBound(varr)
'
'    If lngLeft < lngRight Then
'        lngMid = (lngLeft + lngRight) \ 2
'        strTestVal = cmb.List(lngMid)
'        i = lngLeft
'        j = lngRight
'        Do
'            Do While (cmb.List(i) < strTestVal)
'                i = i + 1
'            Loop
'            Do While (cmb.List(j) > strTestVal)
'                j = j - 1
'            Loop
'            If i <= j Then
'                SwapComboElems cmb, i, j
'                i = i + 1
'                j = j - 1
'            End If
'        Loop Until i > j
'        ' To optimize the sort, always sort the
'        ' smallest segment first.
'        If j <= lngMid Then
'            Call SortCombo(cmb, lngLeft, j)
'            Call SortCombo(cmb, i, lngRight)
'        Else
'            Call SortCombo(cmb, i, lngRight)
'            Call SortCombo(cmb, lngLeft, j)
'        End If
'    End If
'
'End Sub
'
'Sub SwapComboElems(cmb As ComboBox, ByVal i1 As Long, ByVal i2 As Long)
'Dim tTxt As String, tID As Long
'With cmb
'tTxt = .List(i1)
'tID = .ItemData(i1)
'.List(i1) = .List(i2)
'.ItemData(i1) = .ItemData(i2)
'.List(i2) = tTxt
'.ItemData(i2) = tID
'End With
'End Sub
'




'returns an index in array varr< where the value is equal to varsought
'Warning: array must be sorted
Public Function BinarySearchLng(ByRef vArr() As Long, _
                                ByVal lngSought As Long, _
                                ByVal LB As Long, _
                                ByVal UB As Long, _
                                Optional ByVal ptrLUT As Long) As Long
    'Dim lngLower As Long
    Dim lngMiddle As Long
    'Dim lngUpper As Long
    
    'lngLower = LBound(vArr)
    'lngUpper = UBound(vArr)
    If ptrLUT <> 0 Then
        ReferAry AryPtr(CurLUT), ptrLUT
        Do While LB < UB
            ' Increase lower and decrease upper boundary,
            ' keeping varSought in range, if it's there at all.
            lngMiddle = (LB + UB) \ 2
            If lngSought > vArr(CurLUT(lngMiddle)) Then
                LB = lngMiddle + 1
            Else
                UB = lngMiddle
            End If
        Loop
        If vArr(CurLUT(LB)) = lngSought Then
            BinarySearchLng = LB
        Else
            BinarySearchLng = LB
        End If
        
        UnReferAry AryPtr(CurLUT)
    Else
        Do While LB < UB
            ' Increase lower and decrease upper boundary,
            ' keeping varSought in range, if it's there at all.
            lngMiddle = (LB + UB) \ 2
            If lngSought > vArr(lngMiddle) Then
                LB = lngMiddle + 1
            Else
                UB = lngMiddle
            End If
        Loop
        If vArr(LB) = lngSought Then
            BinarySearchLng = LB
        Else
            BinarySearchLng = LB
        End If
    
    End If
    

End Function

