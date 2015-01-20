VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIniFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const FileSignature = "SS SMBMaker settings file v1.0"
Private Const InitialSizes = 100

Private Type Setting
    HashCode As Long
    SettingName As String
    SettingValue As String
End Type

Private Type Section
    HashCode As Long
    SectionName As String
    Heap() As Setting
    HeapMemSize As Long
    HeapLength As Long
    LUT() As Long
    LUTMemSize As Long
    nLUT As Long
End Type

Private Type Base
    Heap() As Section
    HeapMemSize As Long
    HeapLength As Long
    LUT() As Long
    LUTMemSize As Long
    nLUT As Long
End Type

Dim Settings As Base

Public FileName As String

Public Function CalcHashCode(ByRef St As String)
Const ZeroHash = &H499388B3
Dim ByteSeq(0 To 3) As Byte
Dim LongsSeq(0 To 3) As Long
Dim BytesSt() As Byte
Dim i As Long
Dim Ret As Long
If Len(St) = 0 Then
    CalcHashCode = ZeroHash
    Exit Function
End If
BytesSt = StrConv(St, vbFromUnicode)
ReDim Preserve BytesSt(0 To -Int(-Len(St) / 4) * 4 - 1)
For i = 0 To UBound(BytesSt) Step 4
    LongsSeq(0) = LongsSeq(0) + BytesSt(i)
    LongsSeq(1) = LongsSeq(1) + BytesSt(i + 1)
    LongsSeq(2) = LongsSeq(2) + BytesSt(i + 2)
    LongsSeq(3) = LongsSeq(3) + BytesSt(i + 3)
Next i
ByteSeq(0) = LongsSeq(0) Mod 256
ByteSeq(1) = LongsSeq(1) Mod 256
ByteSeq(2) = LongsSeq(2) Mod 256
ByteSeq(3) = LongsSeq(3) Mod 256
CopyMemory Ret, ByteSeq(0), 4
If Ret = 0 Then Ret = ZeroHash
CalcHashCode = Ret

End Function

'////////////////////Section functions\\\\\\\\\\\\\\\\\\\\\\\\
Private Function FindSetting(ByRef Section As Section, _
                             ByRef SettingName As String, _
                             ByRef StHash As Long, _
                             ByRef IndexToLUT As Long) As Boolean
If StHash = 0 Then StHash = CalcHashCode(SettingName)
Dim PosFrom As Long, PosTo As Long
Dim PosMid As Long
Dim i As Long
Dim Ret As Boolean
With Section
    If .nLUT = 0 Then
        IndexToLUT = 0
        FindSetting = False
        Exit Function
    End If
    PosFrom = 0
    PosTo = .nLUT
    
    Do
        PosMid = (PosFrom + PosTo) \ 2
        If .Heap(.LUT(PosMid)).HashCode > StHash Then
            PosTo = PosMid
        ElseIf .Heap(.LUT(PosMid)).HashCode < StHash Then
            PosFrom = PosMid + 1
        Else
            PosFrom = PosMid
            PosTo = PosMid
        End If
    Loop Until PosFrom >= PosTo
    
    IndexToLUT = PosFrom
    
    If PosFrom < .nLUT Then
        If .Heap(.LUT(PosFrom)).HashCode = StHash Then
            'Hash code found
            Ret = .Heap(.LUT(PosFrom)).SettingName = SettingName
            FindSetting = Ret
            If Not Ret Then
                'Some other setting with such hash. Look Forwards/Backwards.
                i = PosFrom - 1
                Do While i >= 0
                    If .Heap(.LUT(i)).HashCode <> StHash Then Exit Do
                    If .Heap(.LUT(i)).SettingName = SettingName Then
                        FindSetting = True
                        IndexToLUT = i
                        Exit Function
                    End If
                    i = i - 1
                Loop
                
                i = PosFrom + 1
                Do While i < .nLUT
                    If .Heap(.LUT(i)).HashCode <> StHash Then Exit Do
                    If .Heap(.LUT(i)).SettingName = SettingName Then
                        FindSetting = True
                        IndexToLUT = i
                        Exit Function
                    End If
                    i = i + 1
                Loop
                
                'Setting not found
                FindSetting = False
            End If
        Else
            'hash not found
            FindSetting = False
        End If
    Else
        'hash not found (out of range)
        FindSetting = False
    End If
End With
End Function

Private Sub IncSectionHeap(ByRef Section As Section)
With Section
    If .HeapMemSize = 0 Then
        .HeapMemSize = InitialSizes
        ReDim .Heap(0 To .HeapMemSize - 1)
    Else
        .HeapMemSize = .HeapMemSize * 1.5 + 1
        ReDim Preserve .Heap(0 To .HeapMemSize - 1)
    End If
End With
End Sub

Private Sub IncSection(ByRef Section As Section)
With Section
    If .LUTMemSize = 0 Then
        .LUTMemSize = InitialSizes
        ReDim .LUT(0 To .LUTMemSize - 1)
    Else
        .LUTMemSize = .LUTMemSize * 1.5 + 1
        ReDim Preserve .LUT(0 To .LUTMemSize - 1)
    End If
End With
End Sub

Private Function AddSetting(ByRef Section As Section, _
                            ByRef Index As Long) As Long
Dim IndexToHeap As Long
Dim i As Long
With Section
    If .HeapMemSize <= .HeapLength Then
        IncSectionHeap Section
    End If
    IndexToHeap = .HeapLength
    AddSetting = IndexToHeap
    .HeapLength = .HeapLength + 1
    
    If .nLUT >= .LUTMemSize Then
        IncSection Section
    End If
    For i = .nLUT - 1 To Index Step -1
        .LUT(i + 1) = .LUT(i)
    Next i
    .LUT(Index) = IndexToHeap
    .nLUT = .nLUT + 1
End With
End Function

Private Sub RemoveSetting(ByRef Section As Section, _
                          ByVal LUTIndex As Long)
Dim i As Long
With Section.Heap(Section.LUT(LUTIndex))
    .HashCode = 0
    .SettingName = ""
    .SettingValue = ""
End With
With Section
    For i = LUTIndex To .nLUT - 1
        .LUT(i) = .LUT(i + 1)
    Next i
    .nLUT = .nLUT - 1
    If .nLUT < 0.8 * .HeapLength Then OptimiseSectionHeap Section
End With
End Sub

Private Sub SaveSectionSetting(ByRef Section As Section, _
                               ByRef SettingName As String, _
                               ByRef SettingValue As String)
Dim HashCode As Long
Dim IndexToLUT As Long, IndexToHeap As Long
Dim uSettingName As String
uSettingName = UCase$(SettingName)
With Section
    If FindSetting(Section, uSettingName, HashCode, IndexToLUT) Then
        IndexToHeap = .LUT(IndexToLUT)
    Else
        IndexToHeap = AddSetting(Section, IndexToLUT)
        .Heap(IndexToHeap).HashCode = HashCode
        .Heap(IndexToHeap).SettingName = uSettingName
    End If
    .Heap(IndexToHeap).SettingValue = SettingValue
End With
End Sub

'returns True if setting present, False otherwise
Private Function GetSectionSetting(ByRef Section As Section, _
                                   ByRef SettingName As String, _
                                   ByRef SettingValue As String) As Boolean
Dim IndexToLUT As Long, IndexToHeap As Long
Dim Ret As Boolean
Dim HashCode As Long
Dim uSettingName As String
uSettingName = UCase$(SettingName)
Ret = FindSetting(Section, uSettingName, HashCode, IndexToLUT)
With Section
    If Ret Then
        SettingValue = .Heap(.LUT(IndexToLUT)).SettingValue
    Else
        SettingValue = ""
    End If
End With
GetSectionSetting = Ret
End Function

Private Function DeleteSectionSetting(ByRef Section As Section, _
                                      ByVal SettingName As String) _
                                      As Boolean
Dim LUTIndex As Long
Dim uSettingName As String
Dim HashCode As Long
uSettingName = UCase$(SettingName)
If FindSetting(Section, uSettingName, HashCode, LUTIndex) Then
    RemoveSetting Section, LUTIndex
    DeleteSectionSetting = True
Else
    DeleteSectionSetting = False
End If
End Function

Private Sub OptimiseSectionHeap(ByRef Section As Section)
Dim NewHeap() As Setting
Dim i As Long
With Section
    .HeapLength = .nLUT
    .HeapMemSize = .nLUT * 1.25 + 1
    ReDim NewHeap(0 To .HeapMemSize - 1)
    For i = 0 To .nLUT - 1
        NewHeap(i) = .Heap(.LUT(i))
        .LUT(i) = i
    Next i
    SwapArys AryPtr(.Heap), AryPtr(NewHeap)
    Erase NewHeap
End With
End Sub

Private Sub TestSectionSorting(ByRef Section As Section)
Dim i As Long
With Section
    For i = 0 To .nLUT - 2
        If .Heap(.LUT(i)).HashCode > .Heap(.LUT(i + 1)).HashCode Then
            Debug.Assert False
            Err.Raise 456, "clsIniFile:TestSorting", "No sorting!"
        End If
    Next i
End With
End Sub

Private Function SectionHasSetting(ByRef Section As Section, _
                                   ByRef SettingName As String) _
                                   As Boolean
Dim LUTIndex As Long
Dim HashCode As Long
SectionHasSetting = FindSetting(Section, SettingName, HashCode, LUTIndex)
End Function
'\\\\\\\\\\\\\\\\\\\Section functions/////////////////////////////


'/////////////////////Base functions\\\\\\\\\\\\\\\\\\\\\\\\\
Private Function FindSection(ByRef Base As Base, _
                             ByRef SectionName As String, _
                             ByRef StHash As Long, _
                             ByRef IndexToLUT As Long) As Boolean
If StHash = 0 Then StHash = CalcHashCode(SectionName)
Dim PosFrom As Long, PosTo As Long
Dim PosMid As Long
Dim i As Long
Dim Ret As Boolean
With Base
    If .nLUT = 0 Then
        IndexToLUT = 0
        FindSection = False
        Exit Function
    End If
    PosFrom = 0
    PosTo = .nLUT
    
    Do
        PosMid = (PosFrom + PosTo) \ 2
        If .Heap(.LUT(PosMid)).HashCode > StHash Then
            PosTo = PosMid
        ElseIf .Heap(.LUT(PosMid)).HashCode < StHash Then
            PosFrom = PosMid + 1
        Else
            PosFrom = PosMid
            PosTo = PosMid
        End If
    Loop Until PosFrom >= PosTo
    
    IndexToLUT = PosFrom
    
    If PosFrom < .nLUT Then
        If .Heap(.LUT(PosFrom)).HashCode = StHash Then
            'Hash code found
            Ret = .Heap(.LUT(PosFrom)).SectionName = SectionName
            FindSection = Ret
            If Not Ret Then
                'Some other Section with such hash. Look Forwards/Backwards.
                i = PosFrom - 1
                Do While i >= 0
                    If .Heap(.LUT(i)).HashCode <> StHash Then Exit Do
                    If .Heap(.LUT(i)).SectionName = SectionName Then
                        FindSection = True
                        IndexToLUT = i
                        Exit Function
                    End If
                    i = i - 1
                Loop
                
                i = PosFrom + 1
                Do While i < .nLUT
                    If .Heap(.LUT(i)).HashCode <> StHash Then Exit Do
                    If .Heap(.LUT(i)).SectionName = SectionName Then
                        FindSection = True
                        IndexToLUT = i
                        Exit Function
                    End If
                    i = i + 1
                Loop
                
                'Section not found
                FindSection = False
            End If
        Else
            'hash not found
            FindSection = False
        End If
    Else
        'hash not found (out of range)
        FindSection = False
    End If
End With
End Function

Private Sub IncBaseHeap(ByRef Base As Base)
With Base
    If .HeapMemSize = 0 Then
        .HeapMemSize = InitialSizes
        ReDim .Heap(0 To .HeapMemSize - 1)
    Else
        .HeapMemSize = .HeapMemSize * 1.5 + 1
        ReDim Preserve .Heap(0 To .HeapMemSize - 1)
    End If
End With
End Sub

Private Sub IncBase(ByRef Base As Base)
With Base
    If .LUTMemSize = 0 Then
        .LUTMemSize = InitialSizes
        ReDim .LUT(0 To .LUTMemSize - 1)
    Else
        .LUTMemSize = .LUTMemSize * 1.5 + 1
        ReDim Preserve .LUT(0 To .LUTMemSize - 1)
    End If
End With
End Sub

Private Function AddSection(ByRef Base As Base, _
                            ByRef Index As Long) As Long
Dim IndexToHeap As Long
Dim i As Long
With Base
    If .HeapMemSize <= .HeapLength Then
        IncBaseHeap Base
    End If
    IndexToHeap = .HeapLength
    AddSection = IndexToHeap
    .HeapLength = .HeapLength + 1
    
    If .nLUT >= .LUTMemSize Then
        IncBase Base
    End If
    For i = .nLUT - 1 To Index Step -1
        .LUT(i + 1) = .LUT(i)
    Next i
    .LUT(Index) = IndexToHeap
    .nLUT = .nLUT + 1
End With
End Function

Private Sub RemoveSection(ByRef Base As Base, _
                          ByVal LUTIndex As Long)
Dim i As Long
With Base.Heap(Base.LUT(LUTIndex))
    .HashCode = 0
    Erase .Heap
    .HeapLength = 0
    .HeapMemSize = 0
    Erase .LUT
    .LUTMemSize = 0
    .nLUT = 0
    .SectionName = ""
End With
With Base
    For i = LUTIndex To .nLUT - 1
        .LUT(i) = .LUT(i + 1)
    Next i
    .nLUT = .nLUT - 1
    If .nLUT < 0.8 * .HeapLength Then OptimiseBaseHeap Base
End With
End Sub

Private Sub MakeBaseSection(ByRef Base As Base, _
                            ByRef SectionName As String, _
                            ByRef LUTIndex As Long)
Dim HashCode As Long
Dim IndexToLUT As Long, IndexToHeap As Long
Dim uSectionName As String
uSectionName = UCase$(SectionName)
With Base
    If FindSection(Base, uSectionName, HashCode, IndexToLUT) Then
        LUTIndex = IndexToLUT
        IndexToHeap = .LUT(IndexToLUT)
    Else
        IndexToHeap = AddSection(Base, IndexToLUT)
        .Heap(IndexToHeap).HashCode = HashCode
        .Heap(IndexToHeap).SectionName = uSectionName
        LUTIndex = IndexToLUT
    End If
End With
End Sub

''returns True if Section present, False otherwise
'Private Function GetBaseSection(ByRef Base As Base, _
'                                   ByRef SectionName As String, _
'                                   ByRef SectionValue As String) As Boolean
'Dim IndexToLUT As Long, IndexToHeap As Long
'Dim Ret As Boolean
'Dim HashCode As Long
'Dim uSectionName As String
'uSectionName = UCase$(SectionName)
'Ret = FindSection(Base, uSectionName, HashCode, IndexToLUT)
'With Base
'    If Ret Then
'        SectionValue = .Heap(.LUT(IndexToLUT)).SectionValue
'    Else
'        SectionValue = ""
'    End If
'End With
'GetSection = Ret
'End Function

Private Function DeleteBaseSection(ByRef Base As Base, _
                                   ByVal SectionName As String) _
                                   As Boolean
Dim LUTIndex As Long
Dim uSectionName As String
Dim HashCode As Long
uSectionName = UCase$(SectionName)
If FindSection(Base, uSectionName, HashCode, LUTIndex) Then
    RemoveSection Base, LUTIndex
    DeleteBaseSection = True
Else
    DeleteBaseSection = False
End If
End Function

Private Sub OptimiseBaseHeap(ByRef Base As Base)
Dim NewHeap() As Section
Dim i As Long
With Base
    .HeapLength = .nLUT
    .HeapMemSize = .nLUT * 1.25 + 1
    ReDim NewHeap(0 To .HeapMemSize - 1)
    For i = 0 To .nLUT - 1
        NewHeap(i) = .Heap(.LUT(i))
        .LUT(i) = i
    Next i
    SwapArys AryPtr(.Heap), AryPtr(NewHeap)
    Erase NewHeap
End With
End Sub

Private Sub TestBaseSorting(ByRef Base As Base)
Dim i As Long
With Base
    For i = 0 To .nLUT - 2
        If .Heap(.LUT(i)).HashCode > .Heap(.LUT(i + 1)).HashCode Then
            Debug.Assert False
            Err.Raise 10050, "clsIniFile:TestSorting", "No sorting!"
        End If
    Next i
    For i = 0 To .nLUT - 1
        TestSectionSorting .Heap(.LUT(i))
    Next i
End With
End Sub

Private Function BaseHasSection(ByRef Base As Base, _
                                   ByRef SectionName As String) _
                                   As Boolean
Dim LUTIndex As Long
Dim HashCode As Long
BaseHasSection = FindSection(Base, SectionName, HashCode, LUTIndex)
End Function
'\\\\\\\\\\\\\\\\\\\\\\\\Base functions/////////////////////////////



'//////////////////////General functions\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Sub LoadFile()
Dim nmb As Long
Dim ID As String
Dim tHeapMemSize As Long
Dim tHeapLength As Long
Dim tnLUT As Long
Dim tLUTMemSize As Long
If Len(FileName) = 0 Then
    Err.Raise 10021, "clsIniFile:LoadFile", "File name is not set."
End If
ID = Space$(Len(FileSignature))
nmb = FreeFile
Open FileName For Binary Access Read As nmb
    If LOF(nmb) = 0 Then Exit Sub 'emty file
    
    Get nmb, 1, ID
    If ID <> FileSignature Then
        Close nmb
        Err.Raise 10020, "clsIniFile:LoadFile", "Bad file header."
    End If
    Get nmb, , Settings
'    Get nmb, , tHeapMemSize
'    Get nmb, , tHeapLength
'    Get nmb, , tLUTMemSize
'    Get nmb, , tnLUT
'
'    If tnLUT < 0 Or _
'       tHeapLength < 0 Or _
'       tHeapLength < 0 Or _
'       tHeapLength < 0 Then
'       Close nmb
'        Err.Raise 123123, "clsIniFile:LoadFile", "Incorrect file."
'    End If
'
'    Erase Heap
'    Erase SortedSettingsLUT
'    HeapMemSize = tHeapMemSize
'    HeapLength = tHeapLength
'    LUTMemSize = tLUTMemSize
'    nLUT = tnLUT
'
'    If HeapMemSize > 0 Then
'        ReDim Heap(0 To HeapMemSize - 1)
'        Get nmb, , Heap
'    End If
'    If LUTMemSize > 0 Then
'        ReDim SortedSettingsLUT(0 To LUTMemSize - 1)
'        Get nmb, , SortedSettingsLUT
'    End If
    
Close nmb
TestBaseSorting Settings
End Sub

Public Sub SaveFile()
Dim nmb As Long
Dim ID As String
If Len(FileName) = 0 Then
    Err.Raise 10021, "clsIniFile:SaveFile", "File name is not set."
End If
ID = FileSignature
nmb = FreeFile
Open FileName For Binary Access Write As nmb
    Put nmb, 1, ID
    Put nmb, , Settings
'    Get nmb, , HeapMemSize
'    Get nmb, , HeapLength
'    Get nmb, , LUTMemSize
'    Get nmb, , nLUT
'
'
'    If HeapMemSize > 0 Then
'        Put nmb, , Heap
'    End If
'    If LUTMemSize > 0 Then
'        Put nmb, , SortedSettingsLUT
'    End If
    
Close nmb
End Sub

'warning: if filetitle is empty, it will be changed!!!
Public Sub DefFile(Optional ByRef FileTitle As String)
Dim AP As String
Dim EXEName As String
Dim nmb As Long
Dim TempTry As Boolean
If Len(App.Path) = 0 Then
    AP = Environ$("TEMP")
    If AP = "" Then AP = CurDir
Else
    AP = App.Path
End If

If Len(FileTitle) = 0 Then
  If Len(EXEName) = 0 Then
      EXEName = Replace(App.ProductName, " ", "")
  Else
      EXEName = App.EXEName
  End If
  FileTitle = EXEName
End If
rsm:
FileName = ValFolder(AP) + FileTitle + ".settings"

On Error GoTo eh
If Not FileExists(FileName) Then
    nmb = FreeFile
    Open FileName For Binary Access Write As nmb
    Close nmb
End If

Exit Sub
eh:
If Not TempTry Then
    'first, try using Temp folder
    TempTry = True
    AP = Environ$("TEMP")
    Resume rsm
Else
    'then, raise an error
    Err.Raise 1212, "clsIniFile:DefFile", "Cannot open or create settings file!"
End If
End Sub

'The same as GetSetting, but uses ByRef Argument to return a value.
Public Function QuerySetting(ByRef SectionName As String, _
                             ByRef SettingName As String, _
                             ByRef SettingValue As String) As Boolean
Dim HashSection As Long
Dim IndexToLUT As Long
With Settings
    If FindSection(Settings, UCase$(SectionName), HashSection, IndexToLUT) Then
        QuerySetting = GetSectionSetting(.Heap(.LUT(IndexToLUT)), SettingName, SettingValue)
    Else
        QuerySetting = False
    End If
End With
End Function

Public Function GetSetting(ByRef SectionName As String, _
                           ByRef SettingName As String) As String
Dim SettingValue As String
Dim uSectionName As String
uSectionName = UCase$(SectionName)
If QuerySetting(uSectionName, SettingName, SettingValue) Then
    GetSetting = SettingValue
Else
    Err.Raise 10020, "clsIniFile:GetSetting", "The setting not found."
End If
End Function

'returns the number of settings found
Public Function GetAllSettings(ByRef SectionName As String, _
                          ByRef AryNames() As String, _
                          ByRef AryValues() As String, _
                          Optional ByVal SortThem As Boolean = True) _
                          As Long
Dim n As Long
Dim uSectionName As String
Dim HashSection As Long
Dim IndexToLUT As Long
Dim i As Long
Dim SortLUT() As Long
uSectionName = UCase$(SectionName)

If FindSection(Settings, uSectionName, HashSection, IndexToLUT) Then
    With Settings.Heap(Settings.LUT(IndexToLUT))
        n = .nLUT
        If n > 0 Then
            ReDim AryNames(0 To n - 1)
            ReDim AryValues(0 To n - 1)
            For i = 0 To n - 1
                AryNames(i) = .Heap(.LUT(i)).SettingName
                AryValues(i) = .Heap(.LUT(i)).SettingValue
            Next i
            If SortThem Then
                strQSortGenerateLUT AryNames, SortLUT
                For i = 0 To n - 1
                    AryNames(i) = .Heap(.LUT(SortLUT(i))).SettingName
                    AryValues(i) = .Heap(.LUT(SortLUT(i))).SettingValue
                Next i
            End If
        Else
            Erase AryNames
            Erase AryValues
        End If
    End With
    GetAllSettings = n
Else
    Err.Raise 10020, "clsIniFile:GetSetting", "The section not found."
End If
End Function

Public Function EnumSections(ByRef Sections() As String) As Long
Dim n As Long
Dim i As Long
With Settings
    n = .nLUT
    If n > 0 Then
        ReDim Sections(0 To n - 1)
        For i = 0 To n - 1
            Sections(i) = .Heap(.LUT(i)).SectionName
        Next i
    Else
        Erase Sections
    End If
End With
End Function

Public Sub SetSetting(ByRef SectionName As String, _
                      ByRef SettingName As String, _
                      ByRef SettingValue As String)
Dim HashSection As Long
Dim IndexToLUT As Long
MakeBaseSection Settings, SectionName, IndexToLUT
With Settings
    SaveSectionSetting .Heap(.LUT(IndexToLUT)), SettingName, SettingValue
End With
End Sub

Private Sub strQSortGenerateLUT(ByRef StArray() As String, _
                                ByRef LUT() As Long, _
                                Optional ByVal lngLeft As Long = -1, _
                                Optional ByVal lngRight As Long = -1)
Dim i As Long
Dim NeedReLUT As Boolean

    If lngLeft > lngRight Then Exit Sub
    
    If lngLeft = -1 Then lngLeft = 0
    If lngRight = -1 Then lngRight = UBound(StArray)
    
    If AryDims(AryPtr(LUT)) = 0 Then
        NeedReLUT = True
    Else
        NeedReLUT = UBound(LUT) <> UBound(StArray)
    End If
    If NeedReLUT Then
        ReDim LUT(0 To UBound(StArray))
        For i = 0 To UBound(LUT)
            LUT(i) = i
        Next i
    End If

    strQSortGenerateLUTRec StArray, LUT, lngLeft, lngRight
                
End Sub

Private Sub strQSortGenerateLUTRec(ByRef StArray() As String, _
                                ByRef LUT() As Long, _
                                ByVal lngLeft As Long, _
                                ByVal lngRight As Long)
    Dim i As Long
    Dim j As Long
    Dim strTestVal As String
    Dim lngMid As Long
    Dim tmp As Long
    
   
    If lngLeft < lngRight Then
        lngMid = (lngLeft + lngRight) \ 2
        strTestVal = StArray(LUT(lngMid))
        i = lngLeft
        j = lngRight
        Do
            Do While (StArray(LUT(i)) < strTestVal)
                i = i + 1
            Loop
            Do While (StArray(LUT(j)) > strTestVal)
                j = j - 1
            Loop
            If i <= j Then
                'swap lut(i) and lut(j)
                tmp = LUT(i)
                LUT(i) = LUT(j)
                LUT(j) = tmp
                
                i = i + 1
                j = j - 1
            End If
        Loop Until i > j
        ' To optimize the sort, always sort the
        ' smallest segment first.
        If j <= lngMid Then
            Call strQSortGenerateLUTRec(StArray, LUT, lngLeft, j)
            Call strQSortGenerateLUTRec(StArray, LUT, i, lngRight)
        Else
            Call strQSortGenerateLUTRec(StArray, LUT, i, lngRight)
            Call strQSortGenerateLUTRec(StArray, LUT, lngLeft, j)
        End If
    End If
End Sub

Public Function DeleteSetting(ByRef SectionName As String, _
                              ByRef SettingName As String) As Boolean
Dim LUTIndex As Long
Dim uSectionName As String
Dim HashCode As Long
uSectionName = UCase$(SectionName)
If FindSection(Settings, uSectionName, HashCode, LUTIndex) Then
    DeleteSetting = DeleteSectionSetting(Settings.Heap(Settings.LUT(LUTIndex)), SettingName)
Else
    DeleteSetting = False
End If
End Function

Public Function DeleteSection(ByRef SectionName As String) As Boolean
Dim uSectionName As String
uSectionName = UCase$(SectionName)
DeleteSection = DeleteBaseSection(Settings, SectionName)
End Function

Public Function SettingPresent(ByRef SectionName As String, _
                               ByRef SettingName As String) _
                               As Boolean
Dim HashCode As Long
Dim LUTIndex As Long
If FindSection(Settings, UCase$(SectionName), HashCode, LUTIndex) Then
    SettingPresent = SectionHasSetting(Settings.Heap(Settings.LUT(LUTIndex)), UCase$(SettingName))
Else
    SettingPresent = False
End If
End Function

Public Function SectionPresent(ByRef SectionName As String) As Boolean
SectionPresent = BaseHasSection(Settings, UCase$(SectionName))
End Function

Public Sub TestMe()
Dim sSetting As String
Dim i As Long
Dim Vl As String
Dim Sec As String, Par As String
Me.SetSetting "Section1", "Test1", "Tst 1 setting"
Me.SetSetting "Section1", "Test2", "Test2 setting"
Me.SetSetting "Section1", "Test3", "Test3 setting"
Me.SetSetting "Section2", "puk", "puk setting"
Me.SetSetting "Section2", "gaka", "gaka setting"
Me.SetSetting "Section2", "pukpukpukpuk", "pukpukpukpuk setting"
Me.SetSetting "Section2", "mud", "mud setting"
Me.SetSetting "Section2", "T1", "T1 setting"

DefFile
For i = 0 To 3000
    Vl = CStr(Int(Rnd(1) * 2000000))
    Sec = CStr(Int(Rnd(1) * 200)) + "m"
    Par = CStr(Int(Rnd(1) * 200)) + "k"
    Me.SetSetting Sec, Par, Vl
    On Error Resume Next
    Debug.Assert Me.GetSetting(UCase(Sec), UCase(Par)) = Vl
Next i
SaveFile
LoadFile
OptimiseBaseHeap Settings
Debug.Print Me.GetSetting("Section2", "t1")
End Sub

Private Function FileExists(ByRef File As String) As Boolean
Dim nmb As Long
On Error GoTo eh
FileExists = True
nmb = FreeFile
Open File For Input As nmb
Close nmb
Exit Function
eh:
FileExists = False
End Function

