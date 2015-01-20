Attribute VB_Name = "mdlVBArrays"
Option Explicit
'WARNING! Before using this module, turn on Autosave in
'Tools/Options, Environment. (Save Changes or Prompt To Save Changes)
'Then restart VB to force it to remember this setting.
'Comment this out or delete when ready.
'This is because if Access Violation occures, you may loose
' unsaved changes. Inaccurate usage of this module can
' cause it very quickly.

'The module for pointed-arrays access using simple VB array.
'It is very useful for bitmaps operation for byte-by-byte access
'and for effective use of the CreateDIBSection API function.
'
'By VT.

'Include this module into your project and use! To find out what is
' this module for, scroll to the end and read comments

'
'Functions:

'AryPtr(<array>) - returns the pointer to the array, not to it's data
'This function must be used to supply pointers to functions:
'  AryDims(ptr)
'  IsAryEmpty(ptr)
'  ConstructAry(ptr,...)
'  ReferAry(ptr1, ptr2)
'  UnreferAry(ptr)
'  SwapArys(ptr1, ptr2)


'ConstructAry(PtrAry, ptrData, ElemSize, Length1[, Length2])
'
'This function should be used to create an array mapped
' to the data pointed by ptrData. Usage:
'ConstructAry AryPtr(MyArray1), VarPtr(MySourceArray(0,0)), 4, ub1+1, ub2+1
'Maximum number of dimensions is 2, minimum is 1. You must supply
' dimension information. You can create an unlimited array by
' setting the last dimension to a very large number. But this is
' not recommended, because it does not protect you from access
' violation.
'Lengths are the numbers of entries. The array base is zero.
' It is not affected by Option Base statement.
'ElemSize is the size of one entry, in bytes. This should match
' the type of the array being referred to.
'Dim ByteAry () as...               ElemSize
'Byte                               1
'Integer                            2
'Boolean                            2
'Long                               2
'Single                             4
'Double                             8
'Currency                           8
'User Defined Type                  LenB()
'String                             DOES NOT WORK AT ALL
'Other data types should be used with accuracy because they have
' not been tested. The exception is user-defined ones, the length of
' which can be calculated using LenB.
'ONCE BEING REFERRED TO, THE ARRAY MUST BE UNREFERRED USING
' UNREFERARY FUNCTION!!! Failure to do this can cause future (!)
' access violation because VB does deallocate the memory automatically.
' (Maybe it does not, I really don't know how it manages memory)
'You must pass an empty array. If the array is not empty,
' the message box is displayed and the error is raised.
'Note. To obtain the pointer to array data, use VarPtr(DataAry(0))
' or something like. Do not forget to pass the pointer, but not the
' element value!
'Warning: Do not use Redim, Erase, Redim Preserve on the created
' array. And call the UnreferAry to this array to avoid the memory
' being deallocated.

'ReferAry(DestAryPtr, SourceAryPtr)
'
'Makes one array(DestAryPtr) coherent with other(SourceAryPtr).
'However, they cannot be redimensioned.
'Only one- and two-dimensional arrays are supported.
'The only requirements are:
'  -the element size of both arrays must match
'  -you must call UnReferAry(DestAryPtr) before the array pointed
'    by DestAryPtr goes out of scope.
'   If you forget to do so, you will have a run-time error
'    (The array is fixed or temporarily locked)
'    when attempting to redimension the source array.

'UnreferAry(PtrAry)
'Be sure to call this function after every call to ConstructAry or
' ReferAry!
' Makes the array empty. Does not deallocate the memory
' unlike Erase does.
'Usage:
'UnReferAry AryPtr(MyArray1)

'AryDims(PtrAry)
'Gets the number of dimensions of an array. Returns 0 if the array
' is empty. Why does not VB provide such function... very useful for
' testing array for emptyness

'IsAryEmpty(PtrAry)
'Returnes the boolean value indicating is the array empty or not.
'The same result gives the following:
'(AryDims(PtrAry)=0)

'SwapArys(PtrAry1, PtrAry2)
'Swaps two arrays. Does not copy the data, only swaps pointers.
' The ElemSize must be the same! If not, the Access Violation can
' occur. Or memory leak.

'TestMdlArys()
'This function tests if there were any arrays unreferred.
' Call it on application termination when debugging.
'No other functions depend on this one, so you can freely
' comment it out when distributing your application.


'Note:
'If you accidently forgot to unrefer an array, restart VB to
'avoid memory leak.


'//////////////the code starts here\\\\\\\\\\\\\\\\\\

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Public Declare Sub ZeroMemory Lib "kernel32" Alias "RtlZeroMemory" (Dest As Any, ByVal numBytes As Long)
Public Declare Function AryPtr Lib "msvbvm60.dll" Alias "VarPtr" (Ary() As Any) As Long

'C++ declaration for SafeArray structure.
'typedef struct FARSTRUCT tagSAFEARRAY {
'    unsigned short cDims;         // Count of dimensions in this array.
'    unsigned short fFeatures;    // Flags used by the SafeArray
'                                // routines documented below.
'#if defined(WIN32)
'    unsigned long cbElements;    // Size of an element of the array.
'                                // Does not include size of
'                                // pointed-to data.
'    unsigned long cLocks;        // Number of times the array has been
'                                // locked without corresponding unlock.
'#Else
'    unsigned short cbElements;
'    unsigned short cLocks;
'    unsigned long handle;        // Unused but kept for compatibility.
'#End If
'    void HUGEP* pvData;                 // Pointer to the data.
'    SAFEARRAYBOUND rgsabound[1];        // One bound for each dimension.
'} SAFEARRAY;

'VB declaration for SafeArray structure
Public Type SafeArrayBound '8 bytes
    cElems As Long '4
    lBound As Long '4
End Type

Public Type SafeArray
    '                length    offset
    cDims As Integer '2        0
    fFeatures As Integer '2    2
    ElemSize As Long '4        4
    cLocks As Long '4          8
    ptrData As Long '4         12
    Bounds(0 To 1) As SafeArrayBound '16 bytes         16
    Mapped As Boolean '2       32      (this and the following is used by this module internally)
    Padding As Boolean '2              Unaligned access works badly
    PtrStruc As Long '4        36      Not used...
    ptrLocksToDecrease As Long '40     Pointer to cLocks member of the array to unlock when releasing
End Type

Private MappedSAs() As SafeArray
Private nMapped As Long

'This sub is here because the dimension data should not be
'deallocated.
'If there are non-unreferred arrays, some structures will not
'be empty (their Mapped field will be True). You can effectively
'use this fact to check is there are forgotten arrays.
'Use TestMdlArys()!
Private Function CreateMappedStruc() As Long
Dim i As Long
Dim Found As Long
If nMapped = 0 Then
    nMapped = 1000 'Increase this number if neccessary - I have _
                   'never tested the module for n of arys >1000
    'If you can refer an ary more than 1000 times, the module may
    ' crash - I am not sure if Redim Preserve does not reallocate
    ' the array or does not.
    ReDim MappedSAs(0 To nMapped - 1)
End If
'find unmapped structure
Found = -1
For i = 0 To nMapped - 1
    If Not MappedSAs(i).Mapped Then
        Found = i
        Exit For
    End If
Next i
If Found = -1 Then
    Found = nMapped
    nMapped = nMapped + 1
    ReDim Preserve MappedSAs(0 To nMapped - 1)
End If
'mark it as mapped
MappedSAs(Found).Mapped = True
CreateMappedStruc = Found
End Function

Public Function ConstructAry(ByVal ptrArray As Long, _
                             ByVal RefToPtr As Long, _
                             ByVal ElemSize As Long, _
                             ByVal Dim1Len As Long, _
                             Optional ByVal Dim2Len As Long = 0) As Long
Dim nDims As Long
Dim PtrStruc As Long
Dim nStruc As Long
If ptrArray = 0 Then
    Err.Raise 1111, "ConstructAry", "No array!"
End If
If RefToPtr = 0 Then
    MsgBox "Null pointer passed. Cannot refer to nothing. Check if VarPtr is put! Or use UnreferAry!"
    Err.Raise 1111, "ConstructAry", "Null pointer passed. Cannot refer to nothing. Check if VarPtr is put! Or use UnreferAry!"
End If
CopyMemory PtrStruc, ByVal ptrArray, 4
If PtrStruc <> 0 Then
    MsgBox "Array is being mapped to is not empty! It must be empty!"
    Debug.Assert False
    Err.Raise 1111, "ConstructAry", "Array is being mapped to is not empty! It must be empty!"
End If
nDims = IIf(Dim2Len = 0, 1, 2)

nStruc = CreateMappedStruc

MappedSAs(nStruc).cDims = nDims
MappedSAs(nStruc).Bounds(0).lBound = 0
If nDims >= 1 Then
    MappedSAs(nStruc).Bounds(0).cElems = Dim1Len
    If nDims = 2 Then
        MappedSAs(nStruc).Bounds(0).cElems = Dim2Len
        MappedSAs(nStruc).Bounds(1).lBound = 0
        MappedSAs(nStruc).Bounds(1).cElems = Dim1Len
    End If
End If
MappedSAs(nStruc).ptrData = RefToPtr
MappedSAs(nStruc).cLocks = 1
MappedSAs(nStruc).ElemSize = ElemSize
MappedSAs(nStruc).fFeatures = &H10& 'no reallocate or resize
MappedSAs(nStruc).ptrLocksToDecrease = 0
CopyMemory ByVal ptrArray, VarPtr(MappedSAs(nStruc)), 4
End Function

Public Sub ReferAry(ByVal ptrAryDest As Long, _
                    ByVal PtrArySrc As Long)
Dim SASrc As SafeArray
Dim ptrStrucDest As Long
Dim ptrStrucSrc As Long
Dim ptrLocks As Long
Dim nStruc As Long
If ptrAryDest = 0 Then
    Err.Raise 1111, "ReferAry", "No array!"
End If
CopyMemory ptrStrucDest, ByVal ptrAryDest, 4
If ptrStrucDest <> 0 Then
    MsgBox "Array is being mapped to is not empty! It must be empty!"
    Debug.Assert False
    Err.Raise 1111, "ReferAry", "Array is being mapped to is not empty! It must be empty!"
End If

CopyMemory ptrStrucSrc, ByVal PtrArySrc, 4
If ptrStrucSrc = 0 Then Exit Sub

CopyMemory SASrc, ByVal ptrStrucSrc, 16
If SASrc.cDims <= 0 Or SASrc.cDims > 2 Then
    Err.Raise 1111, "ReferAry", "Only 1- and 2-dimensional arrays are supported."
End If
CopyMemory SASrc.Bounds(0), ByVal (ptrStrucSrc + 16), 8 * SASrc.cDims

SASrc.cLocks = SASrc.cLocks + 1
ptrLocks = ptrStrucSrc + 8
CopyMemory ByVal ptrLocks, SASrc.cLocks, 4

nStruc = CreateMappedStruc

CopyMemory MappedSAs(nStruc), SASrc, 32
With MappedSAs(nStruc)
    .fFeatures = &H10& 'no reallocate or resize
    .cLocks = 1
    .ptrLocksToDecrease = ptrLocks
End With
CopyMemory ByVal ptrAryDest, VarPtr(MappedSAs(nStruc)), 4
End Sub

Public Function AryDims(ByVal ptrAry As Long) As Long
Dim PtrStruc As Long
Dim SA As SafeArray
If ptrAry = 0 Then
    Debug.Assert False
    Err.Raise 1111, "AryDims", "Illegal pointer passed. Use AryPtr(array)! And do not rely on this error in your code!!!"
End If
CopyMemory PtrStruc, ByVal ptrAry, 4
If PtrStruc = 0 Then
    AryDims = 0
Else
    CopyMemory SA, ByVal PtrStruc, 2
    AryDims = SA.cDims
End If
End Function

Public Function AryWH(ByVal ptrAry As Long, _
                      ByRef w As Long, _
                      ByRef h As Long) As Long
Dim PtrStruc As Long
Dim SA As SafeArray
If ptrAry = 0 Then
    Err.Raise 1111, "AryWH", "No array!"
End If
CopyMemory PtrStruc, ByVal ptrAry, 4
If PtrStruc = 0 Then
    w = 0
    h = 0
Else
    CopyMemory SA, ByVal PtrStruc, 16
    If SA.cDims <> 2 Then
        Err.Raise 1111, "AryWH", "A bidimensional array is required!"
    End If
    CopyMemory SA.Bounds(0), ByVal (PtrStruc + 16), 8 * 2
    w = SA.Bounds(1).cElems
    h = SA.Bounds(0).cElems
End If
End Function

Public Function AryLen(ByVal ptrAry As Long) As Long
Dim PtrStruc As Long
Dim SA As SafeArray
If ptrAry = 0 Then
    Err.Raise 1111, "AryLen", "No array!"
End If
CopyMemory PtrStruc, ByVal ptrAry, 4
If PtrStruc = 0 Then
    AryLen = 0
Else
    CopyMemory SA, ByVal PtrStruc, 16
    If SA.cDims <> 1 Then
        Err.Raise 1111, "AryWH", "A one-dimensional array is required!"
    End If
    CopyMemory SA.Bounds(0), ByVal (PtrStruc + 16), 8 * 1
    AryLen = SA.Bounds(0).cElems
End If
End Function

Public Function IsAryEmpty(ByVal ptrAry As Long) As Boolean
Dim PtrStruc As Long
If ptrAry = 0 Then
    Err.Raise 1111, "IsAryEmpty", "Illegal pointer passed. Use AryPtr(array)! And do not rely on this error in your code!!!"
End If
CopyMemory PtrStruc, ByVal ptrAry, 4
IsAryEmpty = PtrStruc = 0&
End Function

Public Sub UnReferAry(ByVal ptrAry As Long, _
                      Optional ByVal RaiseErrors As Boolean = False)
Dim OldPointer As Long
Dim SA As SafeArray 'only to determine it's length
Dim ptrLocks As Long
Dim cLocks As Long
If ptrAry = 0 Then
    Err.Raise 1111, "UnReferAry", "No array!"
End If
CopyMemory OldPointer, ByVal ptrAry, 4
If OldPointer = 0 Then
    If RaiseErrors Then
        Err.Raise 1111, "UnReferAry", "No array mapped. The array is already empty."
    End If
Else
    'unmap the structure, make the array empty
    CopyMemory ByVal ptrAry, 0&, 4 'copy the number 0 to array structure pointer
    
    CopyMemory SA, ByVal OldPointer, Len(SA)
    If SA.ptrLocksToDecrease <> 0 Then
        ptrLocks = SA.ptrLocksToDecrease
        CopyMemory cLocks, ByVal ptrLocks, 4
        cLocks = cLocks - 1
        If cLocks < 0 Then cLocks = 0
        CopyMemory ByVal ptrLocks, cLocks, 4
    End If
    
    'mark used structure that it is unused (simply fill it with zeros)
    ZeroMemory ByVal (OldPointer), Len(SA)
End If
End Sub

'Takes two identical (by type) arrays. Swaps them. Very fast!
Public Sub SwapArys(ByVal ptrAry1 As Long, ByVal ptrAry2 As Long)
Dim ptrStruc1 As Long, ptrStruc2 As Long
If ptrAry1 = 0 Or ptrAry2 = 0 Then
    Err.Raise 1111, "SwapArys", "No array!"
End If
CopyMemory ptrStruc1, ByVal ptrAry1, 4
CopyMemory ptrStruc2, ByVal ptrAry2, 4

CopyMemory ByVal ptrAry1, ptrStruc2, 4
CopyMemory ByVal ptrAry2, ptrStruc1, 4
End Sub

Public Sub RedimIfNeeded(ByRef Ary() As Long, _
                         ByVal w As Long, _
                         ByVal h As Long)
Dim Needed As Boolean
Needed = True
If AryDims(AryPtr(Ary)) = 2 Then
    If UBound(Ary, 1) = w - 1 And UBound(Ary, 2) = h - 1 Then
        Needed = False
    End If
End If
If Needed Then ReDim Ary(0 To w - 1, 0 To h - 1)
End Sub

'Tests if there are not-unreferred arrays
Public Sub TestMdlArys()
Dim i As Long
For i = 0 To nMapped - 1
    If MappedSAs(i).Mapped Then
        MsgBox "Warning: there was a forgotten array. It's number was " + CStr(i) + "."
        Debug.Assert False
        'Please save changes and restart VB!
        'Or you can find out anything about the forgotten one
        ' by watching MappedSAs(i) (Select it, right-click, Add watch)
        'Press F5 to continue searching
    End If
Next i
End Sub

'******************************************************

'When you might need this module:
'example:
'You have a function which has to do something with the array.
' For example - to reverse it.
' The usual way is this:
'    Function Reverse(ByRef Ary() As Long)
'
'        Dim i As Long 'loop counter
'        Dim AryOut() As Long 'The array to write output to
'        Dim UB As Long 'Ubound of Ary
'        UB = UBound(Ary)
'
'        ReDim AryOut(0 To UB) 'consider the base always 0
'
'        For i = 0 To UB
'            AryOut(i) = Ary(UB - i)
'        Next i
'
'        Ary = AryOut 'Copying of the array - lots of processing
'
'    End Function

'The last stage contains copying of array data (from Ary to
' AryOut). It may take some time if the array is large.
'Let's avoid copying using this module:

'    Function Reverse(ByRef Ary() As Long)
'
'        Dim i As Long 'loop counter
'        Dim AryOut() As Long 'The array to write output to
'        Dim UB As Long 'Ubound of Ary
'        UB = UBound(Ary)
'
'        ReDim AryOut(0 To UB) 'consider the base always 0
'
'        For i = 0 To UB
'            AryOut(i) = Ary(UB - i)
'        Next i
'
'        SwapArys AryPtr(Ary), AryPtr(AryOut) 'swapping pointers - fast
'
'    End Function

'Of course, there are other ways to do it efficiently,
' but this one is simple and fast. It is a simple transform
' which does not need to have a copy of the source array
' in memory, but there can be other transforms not so simple
' as this, to which creating a code without having a full
' copy im memory is a very difficult task.
'Note that SwapArys is safe and will not result in
' access violations. The only thing is to be sure that
' the lengths of data types match.

'Examples of using this module you can find in SMB Maker.
'It is the graphics editor, with open source.
'Please look vt-dbnz.narod.ru for version 3,
'which was not yet published (20.09.2005) because not
'finished. Note that Version 2.x did not use this module,
'and that is one of the main reasons for it to be very slow.

