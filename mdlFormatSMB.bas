Attribute VB_Name = "mdlFormatSMB"
Option Explicit
Private Const SMBID = &H31304D53
Private Const SMBID2 = &H32304D53

Public Sub SaveSMB(ByRef Data() As Long, ByVal File As String)
Dim nmb As Long, i As Long, j As Long 'h As Byte, n As Long
Dim w As Long, h As Long
Dim BitCount As Integer

On Error GoTo eh
StartWrite File
nmb = FreeFile
If AryDims(AryPtr(Data)) = 2 Then
    AryWH AryPtr(Data), w, h
    BitCount = 32
    Open File For Binary Access Write As nmb
        Put nmb, 1, SMBID2
        Put nmb, , BitCount
        Put nmb, , w
        Put nmb, , h
        Put nmb, , Data
    Close nmb
End If
Exit Sub
eh:
Reset
Err.Raise Err.Number
End Sub


Public Function LoadSMB(ByRef Data() As Long, ByVal File As String) As Boolean
    Dim tID As Long, BitCount As Integer
    On Error GoTo eh
    Dim nmb As Long, w As Long, h As Long
    Dim i As Long
    Dim tmp As Long
    Dim RGBData() As RGBQUAD
    Dim vData() As Long
    Dim AlphaAccu As Long
    nmb = FreeFile
    Open File For Binary Access Read As nmb
        If LOF(nmb) = 0 Then
            Err.Raise errNewFile, , "New File"
        End If
        
        Get nmb, 1, tID
        
        If tID = SMBID Then
        
            Get nmb, , BitCount
            If BitCount <> 32 Then
                Err.Raise 1212, , "Only 32-bit SMBs are currently supported."
            End If
            
            Get nmb, , w
            Get nmb, , h
            
            GoSub LoadRGBA
            
        ElseIf tID = SMBID2 Then
        
            Get nmb, , BitCount
            If BitCount <> 32 Then
                MsgBox "Incompatible format"
                LoadSMB = False
                Reset
                Exit Function
            End If
            
            Get nmb, , w
            Get nmb, , h
            
            GoSub LoadBGRA
            
        Else
        
            Get nmb, 1, w
            Get nmb, , h
            If LOF(nmb) = 8 + w * h * 4 Then
                GoSub LoadRGBA
            Else
                Err.Raise 1212, , "Not an SMB file."
            End If
            
        End If
    Close nmb
    
    LoadSMB = CBool(AlphaAccu)
    
Exit Function
Resume
eh:
    PushError
    Close nmb
    UnReferAry AryPtr(RGBData)
    UnReferAry AryPtr(vData)
    PopError
ErrRaise "LoadSMB"

Exit Function

LoadRGBA:
    ReDim Data(0 To w - 1, 0 To h - 1)
    Get nmb, , Data
    
    ConstructAry AryPtr(RGBData), VarPtr(Data(0, 0)), 4, w * h
    ConstructAry AryPtr(vData), VarPtr(Data(0, 0)), 4, w * h
    For i = 0 To w * h - 1
        tmp = RGBData(i).rgbBlue
        RGBData(i).rgbBlue = RGBData(i).rgbRed
        RGBData(i).rgbRed = tmp
        AlphaAccu = AlphaAccu Or vData(i)
    Next i
    UnReferAry AryPtr(RGBData)
    UnReferAry AryPtr(vData)
Return

LoadBGRA:
    ReDim Data(0 To w - 1, 0 To h - 1)
    Get nmb, , Data
    
    ConstructAry AryPtr(vData), VarPtr(Data(0, 0)), 4, w * h
    For i = 0 To w * h - 1
        AlphaAccu = AlphaAccu Or vData(i)
    Next i
    UnReferAry AryPtr(vData)
Return

End Function


Public Function IsSMB(ByVal FileNumber As Long) As Boolean
    Dim tID As Long
    Dim w As Long, h As Long
        
    Get FileNumber, 1, tID
    
    If tID = SMBID Then
        IsSMB = True
    ElseIf tID = SMBID2 Then
        IsSMB = True
    Else
        Get FileNumber, 1, w
        Get FileNumber, , h
        On Error GoTo eh
        If CCur(LOF(FileNumber)) = 8@ + CCur(w) * h * 4@ Then
            IsSMB = True
        End If
    End If
Exit Function
eh:
IsSMB = False
End Function



Public Sub LoadResSMB(ByVal ID As String, ByVal ResType As String, ByRef OutData() As Long)
Dim Ary() As Byte
Dim Sz As Dims
Dim t As Long
If IsNumeric(ID) Then
    Ary = LoadResData(Val(ID), ResType)
Else
    Ary = LoadResData(ID, ResType)
End If
CopyMemory t, Ary(0), 4
If t <> SMBID Then
    Err.Raise 111, "LoadResSMB", "Invalid SMB file. Not an SMB1 format."
End If
CopyMemory Sz, Ary(4 + 2), LenB(Sz)
ReDim OutData(0 To Sz.w - 1, 0 To Sz.h - 1)
CopyMemory OutData(0, 0), Ary(4 + 2 + LenB(Sz)), Sz.w * Sz.h * 4
Dim x As Long, y As Long
For y = 0 To Sz.h - 1
  For x = 0 To Sz.w - 1
    'fill every component using green channel -
    'registry masks are erroneously not purely grey.
    OutData(x, y) = ((OutData(x, y) And &HFF00&) \ &H100&) * &H10101
  Next x
Next y
End Sub

