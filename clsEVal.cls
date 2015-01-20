VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEVal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim WaveSources() As typWaveSource
Dim UndoBegan As Boolean
Dim Priorities() As Integer
Public DrawTemp As Boolean

'Private Type Variable
'    Name As String
'    Value As Double
'End Type
'
Private Enum Operation
    dbOpMult = 0
    dbOpDiv = 1
    dbOpPlus = 2
    dbOpMinus = 3
    dbOpEqv = 4
    dbOpAbove = 5
    dbOpBelow = 6
    dbOpSemicolon = 99
End Enum

Private Type Oper
    OpCode As Operation
    opExpr As String
    opV As Double
End Type


Friend Function EVal(ByRef St As String, ParamArray Vars() As Variant) As Double
Dim i As Long
Dim pVars() As Variable
Dim Var As Variant
If UBound(Vars) > 0 Then
    ReDim pVars(0 To UBound(Vars) \ 2)
    For i = 0 To UBound(pVars)
        pVars(i).Name = UCase$(Vars(i * 2))
        pVars(i).Value = CDbl(Vars(i * 2 + 1))
    Next i
End If
MatVars pVars
CompileExpression_Ex St, Var, pVars
EVal = Execute_Ex(Var, pVars)
End Function

'Friend Function pEVal(ByRef St As String, ByRef Vars() As Variable) As Double
'Dim i As Long, j As Long, h As Long
'Dim m As String
'Dim Expr As String
'Dim Ops() As Oper
'Dim ParCount As Long
'Dim V As Double
'Dim sParams() As String, dParams() As Double
'Dim FuncName As String
'm = Mid$(St, 1, 1)
'If m <> "-" And m <> "+" Then
'    St = "+" + St
'    m = "+"
'End If
'ReDim Ops(0 To 0)
'Ops(0).OpCode = GetOpCode(m)
'For i = 2 To Len(St)
'    m = Mid$(St, i, 1)
'    If m = "(" Then ParCount = ParCount + 1
'    If m = ")" Then ParCount = ParCount - 1
'    If ParCount < 0 Then
'        Err.Raise 111, "EVal", "Parenthesis error"
'    End If
'    If IsOperator(m) And ParCount = 0 Then
'        If i = Len(St) Then
'            Err.Raise 111, "pEVal", "Incorrect Expression"
'        End If
'        Ops(j).opExpr = Expr
'        If Len(Expr) = 0 Then
'            Err.Raise 111, "EVal", "Incorrect operators combination"
'        End If
'        Expr = vbNullString
'        j = j + 1
'        ReDim Preserve Ops(0 To j)
'        Ops(j).OpCode = GetOpCode(m)
'    Else
'        Expr = Expr + m
'    End If
'Next i
'If ParCount <> 0 Then
'    Err.Raise 111, "EVal", "Parenthesis error"
'End If
'Ops(j).opExpr = Expr
'
'For i = 0 To j
'    If IsDigital(Ops(i).opExpr) Then
'        Ops(i).opV = CDbl(Ops(i).opExpr)
'    Else
'        h = InStr(1, Ops(i).opExpr, "(")
'        If h = 0 Then
'            Ops(i).opV = GetVar(Ops(i).opExpr, Vars)
'        ElseIf h = 1 Then
'            Ops(i).opV = pEVal(Mid$(Ops(i).opExpr, 2, Len(Ops(i).opExpr) - 2), Vars)
'        ElseIf h > 1 Then
'            FuncName = Mid$(Ops(i).opExpr, 1, h - 1)
'            sParams = Split(Mid$(Ops(i).opExpr, h + 1, Len(Ops(i).opExpr) - h - 1), ",")
'
'            If UCase(FuncName) <> "SUM" Then
'                ReDim dParams(0 To UBound(sParams))
'
'                For h = LBound(sParams) To UBound(sParams)
'                    dParams(h) = pEVal(sParams(h), Vars)
'                Next h
'                Ops(i).opV = CountFunction(FuncName, dParams)
'            Else
'                Ops(i).opV = CountSum(sParams, Vars)
'            End If
'            Erase dParams, sParams
'        End If
'    End If
'Next i
'
'For i = UBound(Ops) To 1 Step -1
'    If Ops(i).OpCode = dbOpDiv Or Ops(i).OpCode = dbOpMult Then
'        ExecuteOp Ops, i
'    End If
'Next i
'
'For i = 0 To UBound(Ops)
'    If Ops(i).OpCode = dbOpMinus Then
'        Ops(i).opV = -Ops(i).opV
'        Ops(i).OpCode = dbOpPlus
'    End If
'Next i
'
'For i = UBound(Ops) To 1 Step -1
'    If Ops(i).OpCode = dbOpMinus Or Ops(i).OpCode = dbOpPlus Then
'        ExecuteOp Ops, i
'    End If
'Next i
'
'For i = UBound(Ops) To 1 Step -1
'    If Ops(i).OpCode = dbOpEqv Then
'        ExecuteOp Ops, i
'    End If
'Next i
'
''If Ops(0).OpCode = dbOpPlus Then
'    pEVal = Ops(0).opV
''Else
''    pEVal = -Ops(0).opV
''End If
'End Function
'
'Private Sub ExecuteOp(ByRef Ops() As Oper, ByVal OpIndex As Long)
'Dim i As Long, i2 As Long
'i = OpIndex
'Select Case Ops(i).OpCode
'    Case dbOpMult
'        Ops(i - 1).opV = Ops(i - 1).opV * Ops(i).opV
'    Case dbOpDiv
'        Ops(i - 1).opV = Ops(i - 1).opV / Ops(i).opV
'    Case dbOpPlus
'        Ops(i - 1).opV = Ops(i - 1).opV + Ops(i).opV
'    Case dbOpMinus
'        Ops(i - 1).opV = Ops(i - 1).opV - Ops(i).opV
'    Case dbOpEqv
'        Ops(i - 1).opV = Abs(Ops(i - 1).opV = Ops(i).opV)
'End Select
'For i = OpIndex To UBound(Ops) - 1
'    Ops(i) = Ops(i + 1)
'Next i
'ReDim Preserve Ops(0 To UBound(Ops) - 1)
'End Sub
'
'Private Function CountFunction(ByRef FuncName As String, ByRef Params() As Double) As Double
'Dim i As Long, j As Long
'On Error GoTo eh
'Select Case UCase(FuncName)
'    Case "PUT"
'        'todo:write execution of the function Put
'    Case "SIN"
'        If UBound(Params) <> 0 Then GoTo ParamError
'        CountFunction = Sin(Params(0))
'    Case "COS"
'        If UBound(Params) <> 0 Then GoTo ParamError
'        CountFunction = Cos(Params(0))
'    Case "TAN"
'        If UBound(Params) <> 0 Then GoTo ParamError
'        CountFunction = Tan(Params(0))
'    Case "ATN"
'        If UBound(Params) <> 0 Then GoTo ParamError
'        CountFunction = Atn(Params(0))
'    Case "RND"
'        CountFunction = Rnd(Params(0))
'    Case "SQR"
'        CountFunction = Sqr(Params(0))
'    Case "EXP"
'        If UBound(Params) <> 0 Then GoTo ParamError
'        CountFunction = Exp(Params(0))
'    Case "INT"
'        If UBound(Params) <> 0 Then GoTo ParamError
'        CountFunction = Int(Params(0))
'    Case "IIF"
'        If UBound(Params) <> 2 Then GoTo ParamError
'        CountFunction = IIf(Params(0), Params(1), Params(2))
'
'    Case Else
'        On Error GoTo 0
'        Err.Raise 111, "EVal", "Unknown function: " + FuncName
'End Select
'Exit Function
'eh:
'Err.Raise 111, "EVal", "Illegal function call"
'ParamError:
'Err.Raise 111, "EVal", "Argument count mismatch"
'End Function


Private Function GetFunctionIndex(ByRef FuncName As String, ByVal pC As Long) As Double
'pc is paramcount-1
Dim ParamCount As Long
ParamCount = pC + 1
On Error GoTo eh
Select Case UCase(FuncName)
    Case "PUT" '-1
        If pC > 5 Or pC < 4 Then GoTo ParamError
        GetFunctionIndex = -1
    Case "SIN" '0
        If pC <> 0 Then GoTo ParamError
        GetFunctionIndex = 0
    Case "COS" '1
        If pC <> 0 Then GoTo ParamError
        GetFunctionIndex = 1
    Case "TAN" '2
        If pC <> 0 Then GoTo ParamError
        GetFunctionIndex = 2
    Case "ATN" '3
        If pC <> 0 Then GoTo ParamError
        GetFunctionIndex = 3
    Case "SQR" '4
        If pC <> 0 Then GoTo ParamError
        GetFunctionIndex = 4
    Case "RND" '5
        If pC <> 0 Then GoTo ParamError
        GetFunctionIndex = 5
    Case "EXP" '6
        If pC <> 0 Then GoTo ParamError
        GetFunctionIndex = 6
    Case "INT" '7
        If pC <> 0 Then GoTo ParamError
        GetFunctionIndex = 7
    Case "IIF", "IF" '8
        If pC <> 1 And pC <> 2 Then GoTo ParamError
        GetFunctionIndex = 8
    Case "SUM" '9
        If pC < 3 Then GoTo ParamError
        GetFunctionIndex = 9
    Case "LN" '10
        If pC <> 0 Then GoTo ParamError
        GetFunctionIndex = 10
    Case "LOG" '11
        If pC <> 1 Then GoTo ParamError
        GetFunctionIndex = 11
    Case "LG" '12
        If pC <> 0 Then GoTo ParamError
        GetFunctionIndex = 12
    Case "SGN" '13
        If pC <> 0 Then GoTo ParamError
        GetFunctionIndex = 13
    Case "MAX" '14
        If pC <> 1 Then GoTo ParamError
        GetFunctionIndex = 14
    Case "MIN" '15
        If pC <> 1 Then GoTo ParamError
        GetFunctionIndex = 15
    Case "EABOVE" '16
        If pC <> 1 Then GoTo ParamError
        GetFunctionIndex = 16
    Case "EBELOW" '17
        If pC <> 1 Then GoTo ParamError
        GetFunctionIndex = 17
    Case "ABS" '18
        If pC <> 0 Then GoTo ParamError
        GetFunctionIndex = 18
    Case "POW" '19
        If pC <> 1 Then GoTo ParamError
        GetFunctionIndex = 19
        

    Case "ASSIGN" '20
        If pC <> 1 Then GoTo ParamError
        GetFunctionIndex = 20
        
    Case "RGB" '29
        If pC <> 2 Then GoTo ParamError
        GetFunctionIndex = 29
    
    Case "DATAR" '30
        If pC <> 1 Then GoTo ParamError
        GetFunctionIndex = 30
    Case "DATAG" '31
        If pC <> 1 Then GoTo ParamError
        GetFunctionIndex = 31
    Case "DATAB" '32
        If pC <> 1 Then GoTo ParamError
        GetFunctionIndex = 32
    Case "DATARGB" '33
        If pC <> 1 Then GoTo ParamError
        GetFunctionIndex = 33
        
    Case "PSET" '34
        If pC <> 2 Then GoTo ParamError
        GetFunctionIndex = 34
    Case "PSETEX" '35
        If pC <> 2 Then GoTo ParamError
        GetFunctionIndex = 35
        '            3  4  5  6      7       8      9    10   11
    Case "LINE" '36'x1,y1,x2,y2,color1,[color2,[count,power,mode,[width]]]
        If pC <> 4 And pC <> 5 And pC <> 8 And pC <> 9 Then GoTo ParamError
        GetFunctionIndex = 36
    Case "CIRCLE" '37
        If pC <> 4 And pC <> 5 And pC <> 8 Then GoTo ParamError
        GetFunctionIndex = 37
    Case "RECT" '38
        If pC <> 4 And pC <> 3 Then GoTo ParamError
        GetFunctionIndex = 38
    Case "FRECT" '39
        If pC <> 4 And pC <> 3 Then GoTo ParamError
        GetFunctionIndex = 39
    Case "INRECT" '40
        If pC <> 4 And pC <> 3 Then GoTo ParamError
        GetFunctionIndex = 40
    
    
    Case "PROGRESS" '50
        If pC <> 0 Then GoTo ParamError
        GetFunctionIndex = 50
    Case "GETARY" '51
        If pC <> 0 Then GoTo ParamError
        GetFunctionIndex = 51
    Case "SETARY" '52
        If pC <> 1 Then GoTo ParamError
        GetFunctionIndex = 52
    Case "REFRESH" '53
        If pC <> 0 Then GoTo ParamError
        GetFunctionIndex = 53
    Case "INIT" '54 'a special case
        If pC <> 1 Then GoTo ParamError
        GetFunctionIndex = 54
    Case "MOVEMOUSE" '55
        If pC <> 0 Then GoTo ParamError
        GetFunctionIndex = 55
        
    Case "SETDRAWMODE" '56
        If pC <> 0 Then GoTo ParamError
        GetFunctionIndex = 56
    Case "REMOVETEMP" '57
        If pC <> 0 Then GoTo ParamError
        GetFunctionIndex = 57
    Case "LOOP" '58
        If pC <> 3 Then GoTo ParamError
        GetFunctionIndex = 58
    Case "ATAN2", "ARG" '59
        If pC <> 1 Then GoTo ParamError
        GetFunctionIndex = 59
    Case "XYLEN" '60
        If pC <> 1 Then GoTo ParamError
        GetFunctionIndex = 60
    Case "VECTORLEN" '61
        GetFunctionIndex = 61
    Case "NORMALIZE" '62 'A special case
        GetFunctionIndex = 62
    Case "DB" '63
        If pC <> 0 Then GoTo ParamError
        GetFunctionIndex = 63
    Case "OUTPUT" '64
        GetFunctionIndex = 64
    Case "LINEEX" '65
        GetFunctionIndex = 65
        Select Case ParamCount
          Case 3, 4, 6, 7, 8, 9, 13
          Case Else
            GoTo ParamError
        End Select
    Case "LINE3OPAQ"
        GetFunctionIndex = 66
        Select Case ParamCount
          Case 4, 5, 7, 8, 9, 10, 14
          Case Else
            GoTo ParamError
        End Select
    Case "BEGINTRANSFORM"
      GetFunctionIndex = 67
      Select Case ParamCount
        Case 0, 1, 2
        Case Else
          GoTo ParamError
      End Select
    Case "ENDTRANSFORM"
      GetFunctionIndex = 68
      Select Case ParamCount
        Case 0, 1
        Case Else
          GoTo ParamError
      End Select
    Case "TRANSFORMBLOCK"
      GetFunctionIndex = 69
      Select Case ParamCount
        Case 12, 16
        Case Else
          GoTo ParamError
      End Select
    Case Else
        On Error GoTo 0
        Err.Raise 111, "EVal", "Unknown function: " + FuncName + "."
End Select
Exit Function
Resume
eh:
Err.Raise 111, "EVal", "Unexpected error."
ParamError:
On Error GoTo 0
Err.Raise 111, "EVal", "Argument count mismatch in function " + FuncName + "."
End Function

Private Function CountFunctionCompiled(ByRef Var As Variant, ByRef Vars() As Variable) As Double
Const Log_10 As Double = 2.30258509299405
Dim LB As Double, UB As Double
Dim t As Integer
Dim ErrorOccured As Boolean
Dim i As Long, j As Long
Dim s As Double
Dim lngUB As Long
Dim ParamCount As Long
Dim r As Double, g As Double, b  As Double
Dim x As Double, y As Double
Dim FDSC As FadeDesc
Dim X1 As Single, Y1 As Single
Dim X2 As Single, Y2 As Single
Dim Color1 As Long, Color2 As Long
Dim nFrom As Double, nTo As Double
Dim w1 As Double, w2 As Double
Dim LO As LineSettings, LOt As LineSettings
Dim aa As Double
Select Case Var(2)
    Case 20 'Assign
        Vars(Var(3)).Value = Execute_Ex(Var(4), Vars)
    Case 8 'IIf
        If Execute_Ex(Var(3), Vars) Then
            CountFunctionCompiled = Execute_Ex(Var(4), Vars)
        ElseIf UBound(Var) >= 5 Then
            CountFunctionCompiled = Execute_Ex(Var(5), Vars)
        End If

    Case -1 'put
        ParamCount = UBound(Var)
        On Error Resume Next
        If AryDims(AryPtr(WaveSources)) = 1 Then
            lngUB = UBound(WaveSources) + 1
        End If
        On Error GoTo 0
        If lngUB = 0 Then
            ReDim WaveSources(0 To lngUB)
        Else
            ReDim Preserve WaveSources(0 To lngUB)
        End If
        With WaveSources(lngUB)
            .Pos.x = Execute_Ex(Var(3), Vars)
            .Pos.y = Execute_Ex(Var(4), Vars)
            .Strength = Execute_Ex(Var(5), Vars)
            '.r = Execute_Ex(Var(5), Vars)
            '.g = Execute_Ex(Var(6), Vars)
            '.b = Execute_Ex(Var(7), Vars)
            If ParamCount >= 8 Then
                .WaveLength = Execute_Ex(Var(8), Vars)
                If .WaveLength = 0 Then Err.Raise 114, "CountFunctionCompiled", "WaveLength cannot be 0."
            Else
                .WaveLength = 10
            End If
        End With
        
    Case 0
        CountFunctionCompiled = Sin(Execute_Ex(Var(3), Vars))
    Case 1
        CountFunctionCompiled = Cos(Execute_Ex(Var(3), Vars))
    Case 2
        CountFunctionCompiled = Tan(Execute_Ex(Var(3), Vars))
    Case 3
        CountFunctionCompiled = Atn(Execute_Ex(Var(3), Vars))
    Case 4
        CountFunctionCompiled = Sqr(Execute_Ex(Var(3), Vars))
    Case 5
        CountFunctionCompiled = Rnd(Execute_Ex(Var(3), Vars))
    Case 6
        CountFunctionCompiled = Exp(Execute_Ex(Var(3), Vars))
    Case 7
        CountFunctionCompiled = Int(Execute_Ex(Var(3), Vars))
    '8 was if, moved to the beginning
    Case 9 'sum
        t = Var(3)
        nFrom = Execute_Ex(Var(4), Vars)
        nTo = Execute_Ex(Var(5), Vars)
        UB = UBound(Var)
        s = 0
        For i = nFrom To nTo
            Vars(t).Value = i
            For j = 6 To UB
                s = s + Execute_Ex(Var(j), Vars)
            Next j
            i = Vars(t).Value
        Next i
        CountFunctionCompiled = s
    Case 10 'Ln
        CountFunctionCompiled = Log(Execute_Ex(Var(3), Vars))
    Case 11 'Log
        CountFunctionCompiled = Log(Execute_Ex(Var(4), Vars)) / Log(Execute_Ex(Var(3), Vars))
    Case 12 'Lg
        CountFunctionCompiled = Log(Execute_Ex(Var(3), Vars)) / Log_10
    Case 13 'Sgn
        CountFunctionCompiled = Sgn(Execute_Ex(Var(3), Vars))
    Case 14 'Max
        x = Execute_Ex(Var(3), Vars)
        y = Execute_Ex(Var(4), Vars)
        If x > y Then
            CountFunctionCompiled = x
        Else
            CountFunctionCompiled = y
        End If
    Case 15 'Min
        x = Execute_Ex(Var(3), Vars)
        y = Execute_Ex(Var(4), Vars)
        If x < y Then
            CountFunctionCompiled = x
        Else
            CountFunctionCompiled = y
        End If
    Case 16 'eAbove
        x = Execute_Ex(Var(3), Vars)
        y = Execute_Ex(Var(4), Vars)
        CountFunctionCompiled = (x >= y)
    Case 17 'eBelow
        x = Execute_Ex(Var(3), Vars)
        y = Execute_Ex(Var(4), Vars)
        CountFunctionCompiled = (x <= y)
    Case 18 'ABS
        CountFunctionCompiled = Abs(Execute_Ex(Var(3), Vars))
    Case 19 'Pow
        x = Execute_Ex(Var(3), Vars)
        y = Execute_Ex(Var(4), Vars)
        CountFunctionCompiled = x ^ y
    'here was assign
    
    Case 29 'RGB
        r = Int(Execute_Ex(Var(3), Vars))
        If r > 255 Then r = 255
        If r < 0 Then r = 0
        g = Int(Execute_Ex(Var(4), Vars))
        If g > 255 Then g = 255
        If g < 0 Then g = 0
        b = Int(Execute_Ex(Var(5), Vars))
        If b > 255 Then b = 255
        If b < 0 Then b = 0
        CountFunctionCompiled = RGB(b, g, r)
    Case 30 'DataR
        x = Round(Execute_Ex(Var(3), Vars))
        y = Round(Execute_Ex(Var(4), Vars))
        CountFunctionCompiled = MainForm.GetAttr(MainForm.Pixel(x, y), 3)
    Case 31 'DataG
        x = Round(Execute_Ex(Var(3), Vars))
        y = Round(Execute_Ex(Var(4), Vars))
        CountFunctionCompiled = MainForm.GetAttr(MainForm.Pixel(x, y), 2)
    Case 32 'DataB
        x = Round(Execute_Ex(Var(3), Vars))
        y = Round(Execute_Ex(Var(4), Vars))
        CountFunctionCompiled = MainForm.GetAttr(MainForm.Pixel(x, y), 1)
    Case 33 'DataRGB
        x = Round(Execute_Ex(Var(3), Vars))
        y = Round(Execute_Ex(Var(4), Vars))
        CountFunctionCompiled = MainForm.Pixel(x, y)
    
    Case 34 'PSet
        x = Round(Execute_Ex(Var(3), Vars))
        y = Round(Execute_Ex(Var(4), Vars))
        r = CLng(Execute_Ex(Var(5), Vars)) And &HFFFFFF
'        If r > 255 Then r = 255
'        If r < 0 Then r = 0
'        g = Int(Execute_Ex(Var(6), Vars))
'        If g > 255 Then g = 255
'        If g < 0 Then g = 0
'        b = Int(Execute_Ex(Var(7), Vars))
'        If b > 255 Then b = 255
'        If b < 0 Then b = 0
        CheckUndo
        On Error GoTo ehDraw
        If DrawTemp Then
            MainForm.dbPSet x, y, CLng(r), , False, , PrgDrawMode = dbDrawDirect
        Else
            MainForm.Data_SetPixel x, y, CLng(r)
        End If
    Case 35 'PSetEx
        x = Execute_Ex(Var(3), Vars)
        y = Execute_Ex(Var(4), Vars)
        r = CLng(Execute_Ex(Var(5), Vars)) And &HFFFFFF
'        If r > 255 Then r = 255
'        If r < 0 Then r = 0
'        g = Int(Execute_Ex(Var(6), Vars))
'        If g > 255 Then g = 255
'        If g < 0 Then g = 0
'        b = Int(Execute_Ex(Var(7), Vars))
'        If b > 255 Then b = 255
'        If b < 0 Then b = 0
        CheckUndo
        On Error GoTo ehDraw
        MainForm.dbPutPoint x, y, CLng(r), DrawTemp, PrgDrawMode = dbDrawDirect
    Case 36 'Line(x1,y1,x2,y2,color1,color2,c,p,m,width)
        Dim LineWeight As Double
        X1 = Execute_Ex(Var(3), Vars)
        Y1 = Execute_Ex(Var(4), Vars)
        X2 = Execute_Ex(Var(5), Vars)
        Y2 = Execute_Ex(Var(6), Vars)
        Color1 = Int(Execute_Ex(Var(7), Vars))
        Color2 = 0
        FDSC.FCount = 0
        FDSC.Power = 0.5
        LineWeight = 1
        If UBound(Var) = 8 Then
            Color2 = Execute_Ex(Var(8), Vars)
            FDSC.FCount = 1
        ElseIf UBound(Var) >= 11 Then
            Color2 = Execute_Ex(Var(8), Vars)
            FDSC.FCount = Execute_Ex(Var(9), Vars)
            FDSC.Power = Execute_Ex(Var(10), Vars)
            FDSC.Mode = Execute_Ex(Var(11), Vars)
            If UBound(Var) = 12 Then
                LineWeight = Execute_Ex(Var(12), Vars)
            End If
        End If
        CheckUndo
        On Error GoTo ehDraw
        MainForm.dbFade X1, Y1, X2, Y2, Color1, Color2, FDSC, DrawTemp, 0, , , PrgDrawMode = dbDrawDirect, Weight1:=LineWeight
    Case 37 'Circle (cx,cy,ry,rx,color1,[color2,[c,p,m]])
        X1 = Execute_Ex(Var(3), Vars)
        Y1 = Execute_Ex(Var(4), Vars)
        X2 = Execute_Ex(Var(5), Vars)
        Y2 = Execute_Ex(Var(6), Vars)
        Color1 = Int(Execute_Ex(Var(7), Vars))
        Color2 = 0
        FDSC.FCount = 0
        FDSC.Power = 1
        If UBound(Var) = 8 Then
            Color2 = Execute_Ex(Var(8), Vars)
        ElseIf UBound(Var) = 11 Then
            Color2 = Execute_Ex(Var(8), Vars)
            FDSC.FCount = Execute_Ex(Var(9), Vars)
            FDSC.Power = Execute_Ex(Var(10), Vars)
            FDSC.Mode = Execute_Ex(Var(11), Vars)
        End If
        CheckUndo
        On Error GoTo ehDraw
        MainForm.dbSimpleCircle_Ellipse X1, Y1, X2, Y2, Color1, Color2, FDSC, DrawTemp, 0, , PrgDrawMode = dbDrawDirect
    Case 38 'Rect(x1,y1,x2,y2,[Color])
        X1 = Execute_Ex(Var(3), Vars)
        Y1 = Execute_Ex(Var(4), Vars)
        X2 = Execute_Ex(Var(5), Vars)
        Y2 = Execute_Ex(Var(6), Vars)
        If UBound(Var) >= 7 Then
            Color1 = Int(Execute_Ex(Var(7), Vars))
        Else
            Color1 = MainForm.GetACol(1)
        End If
        CheckUndo
        On Error GoTo ehDraw
        MainForm.dbRect X1, Y1, X2, Y2, Color1, , DrawTemp, , , PrgDrawMode = dbDrawDirect
    Case 39 'FRect 'RectFilled
        X1 = Execute_Ex(Var(3), Vars)
        Y1 = Execute_Ex(Var(4), Vars)
        X2 = Execute_Ex(Var(5), Vars)
        Y2 = Execute_Ex(Var(6), Vars)
        If UBound(Var) >= 7 Then
            Color1 = Int(Execute_Ex(Var(7), Vars))
        Else
            Color1 = MainForm.GetACol(1)
        End If
        CheckUndo
        On Error GoTo ehDraw
        MainForm.dbRect X1, Y1, X2, Y2, Color1, , DrawTemp, , dbFilled, PrgDrawMode = dbDrawDirect
    Case 40 'InRect 'RectInside
        X1 = Execute_Ex(Var(3), Vars)
        Y1 = Execute_Ex(Var(4), Vars)
        X2 = Execute_Ex(Var(5), Vars)
        Y2 = Execute_Ex(Var(6), Vars)
        If UBound(Var) >= 7 Then
            Color1 = Int(Execute_Ex(Var(7), Vars))
        Else
            Color1 = MainForm.GetACol(1)
        End If
        CheckUndo
        On Error GoTo ehDraw
        MainForm.dbRect X1, Y1, X2, Y2, Color1, , DrawTemp, MainForm.GetACol(2), dbFilledBG, PrgDrawMode = dbDrawDirect
        
        
    Case 50 'Progress
        X1 = Execute_Ex(Var(3), Vars)
        MainForm.MeEnabled = False
        MainForm.ShowProgress X1, True
        
    Case 51 'GetAry
        X1 = Round(Execute_Ex(Var(3), Vars))
        X1 = dbModDL(X1, MaxPrgBufferLength)
        If X1 >= PrgBufferLen Then
            ReDim Preserve PrgBuffer(0 To X1)
            PrgBufferLen = X1 + 1
        End If
        CountFunctionCompiled = PrgBuffer(X1)
    Case 52 'SetAry
        X1 = Round(Execute_Ex(Var(3), Vars))
        X1 = dbModDL(X1, MaxPrgBufferLength)
        Y1 = Execute_Ex(Var(4), Vars)
        If X1 >= PrgBufferLen Then
            ReDim Preserve PrgBuffer(0 To X1)
            PrgBufferLen = X1 + 1
        End If
        PrgBuffer(X1) = Y1
        CountFunctionCompiled = Y1
    Case 53 'Refresh
        If Execute_Ex(Var(3), Vars) = 0 Then
            If PrgDrawMode = dbDrawDirect Then
                If MainForm.MP.AutoRedraw Then MainForm.MP.Refresh
            Else
                MainForm.Refr True
            End If
        End If
    Case 54 'Init
        'do nothing
    Case 55 'MoveMouse
        MoveMouse
    
    Case 56 'SetDrawMode
        DrawTemp = Execute_Ex(Var(3), Vars) <> 0
    Case 57 'RemoveTemp
        MainForm.RemoveTempPixels
    Case 58 'Loop
        Execute_Ex Var(3), Vars
        Do While Execute_Ex(Var(5), Vars) = 0
            Execute_Ex Var(6), Vars
            Execute_Ex Var(4), Vars
        Loop
        
    Case 59 'ARG, ATAN2
        x = Execute_Ex(Var(3), Vars)
        y = Execute_Ex(Var(4), Vars)
        CountFunctionCompiled = Arg(x, y)
    Case 60 'XYLen
        CountFunctionCompiled = XYLen(Execute_Ex(Var(3), Vars), _
                                      Execute_Ex(Var(4), Vars))
    Case 61 'VectorLen
        CountFunctionCompiled = CountVLen(Var, Vars)
    Case 62 'Normalize
        CountFunctionCompiled = Normalize(Var, Vars)
    Case 63 'dB
        CountFunctionCompiled = Log(Execute_Ex(Var(3), Vars)) / Log(10) * 20
    Case 64 'Output
        CountFunctionCompiled = CountOutput(Var, Vars)
    Case 65
      'LineEx( x,y,width[,color[,x,y[,width[,color[,aa[,fm,fc,fp,fo]]]]]
      '        1 2   3      4    5 6    7      8     9  10 11 12 13
'LineEx(x1, y1, width1[, color1[, x2, y2[, width2[, color[, aa[, fm, fc, fp, fo]]]]]])
'Функция аналоги
      X1 = ExecuteArg(Var, 1, Vars)
      Y1 = ExecuteArg(Var, 2, Vars)
      w1 = ExecuteArg(Var, 3, Vars, 1)
      Color1 = ExecuteArg(Var, 4, Vars, MainForm.GetACol(1))
      X2 = ExecuteArg(Var, 5, Vars, X1)
      Y2 = ExecuteArg(Var, 6, Vars, Y1)
      w2 = ExecuteArg(Var, 7, Vars, w1)
      Color2 = ExecuteArg(Var, 8, Vars, Color1)
      aa = ExecuteArg(Var, 9, Vars, 1)
      If aa <= 0 Then Err.Raise 1234, "LineEx", "Antialiasing must be greater than zero!"
      
      FDSC.Mode = ExecuteArg(Var, 10, Vars, dbFLinear)
      FDSC.FCount = ExecuteArg(Var, 11, Vars, 1)
      FDSC.Offset = ExecuteArg(Var, 13, Vars, 0)
      MainForm.GetLineOpts LO
      LOt = LO
      LO.AntiAliasing = 1 / aa
      MainForm.SetLineOpts LO
      On Error GoTo ehLineEx
              CheckUndo
'        On Error GoTo ehDraw
'        MainForm.dbFade x1, y1, x2, y2, Color1, Color2, FDSC, DrawTemp, 0, , , PrgDrawMode = dbDrawDirect, Weight1:=LineWeight
'
      MainForm.dbFade X1, Y1, X2, Y2, Color1, Color2, FDSC, DrawTemp, 0, , , PrgDrawMode = dbDrawDirect, Weight1:=w1, Weight2:=w2
      MainForm.SetLineOpts LOt
      On Error GoTo 0
    Case 66
      'Line3opaq(opaq, x,y,width[,color[,x,y[,width[,color[,aa[,fm,fc,fp,fo]]]]]
      '          1     2 3   4      5    6 7    8      9     10 11 12 13 14
      Dim Vtx1 As vtVertex3opaq, Vtx2 As vtVertex3opaq
      Vtx1.opaq = ExecuteArg(Var, 1, Vars)
      Vtx1.x = ExecuteArg(Var, 2, Vars)
      Vtx1.y = ExecuteArg(Var, 3, Vars)
      Vtx1.Weight = ExecuteArg(Var, 4, Vars, 1)
      Vtx1.Color = ExecuteArg(Var, 5, Vars, MainForm.GetACol(1))
      Vtx2.x = ExecuteArg(Var, 6, Vars, Vtx1.x)
      Vtx2.y = ExecuteArg(Var, 7, Vars, Vtx1.y)
      Vtx2.Weight = ExecuteArg(Var, 8, Vars, Vtx1.Weight)
      Vtx2.Color = ExecuteArg(Var, 9, Vars, Vtx1.Color)
      aa = ExecuteArg(Var, 10, Vars, 1)
      If aa <= 0 Then Err.Raise 1234, "LineEx", "Antialiasing must be greater than zero!"
      
      FDSC.Mode = ExecuteArg(Var, 11, Vars, dbFLinear)
      FDSC.FCount = ExecuteArg(Var, 12, Vars, 1)
      FDSC.Offset = ExecuteArg(Var, 14, Vars, 0)
'      MainForm.GetLineOpts LO
'      LOt = LO
'      LO.AntiAliasing = 1 / aa
'      MainForm.SetLineOpts LO
      On Error GoTo ehDraw
              CheckUndo
'        On Error GoTo ehDraw
'        MainForm.dbFade x1, y1, x2, y2, Color1, Color2, FDSC, DrawTemp, 0, , , PrgDrawMode = dbDrawDirect, Weight1:=LineWeight
'
      MainForm.dbLine3opaq Vtx1, Vtx2, aa, FDSC, ForceDraw:=PrgDrawMode = dbDrawDirect, DrawTemp:=DrawTemp
      'MainForm.SetLineOpts LOt
      On Error GoTo 0
    Case 67
      'BeginTransform([w,[h]])
      MainForm.BeginTransform ExecuteArg(Var, 1, Vars, 0), ExecuteArg(Var, 2, Vars, 0)
    Case 68
      'EndTransform()
      MainForm.EndTransform
      NewUndo 'EndTransform does not call StartPixelAction, so it should be called on next drawing
    Case 69
      'TransformBlock(tx1, ty1, tx2, ty2,  sx1, sy1, sx2, sy2, sx3, sy3, sx4, sy4, [px1,py1,px2,py2])
      '                1    2    3    4     5    6    7    8    9    10   11   12    13  14  15  16
      Dim dstRect As FloatRect
      Dim srcGon As vtQGon
      Dim ProcessRect As RECT
      With dstRect
        .X1 = ExecuteArg(Var, 1, Vars)
        .Y1 = ExecuteArg(Var, 2, Vars)
        .X2 = ExecuteArg(Var, 3, Vars)
        .Y2 = ExecuteArg(Var, 4, Vars)
      End With
      For i = 1 To 4
        srcGon.v(i).x = ExecuteArg(Var, 5 + (i - 1) * 2, Vars)
        srcGon.v(i).y = ExecuteArg(Var, 6 + (i - 1) * 2, Vars)
      Next i
      If UBound(Var) - 2 = 16 Then
        With ProcessRect
          .Left = Int(ExecuteArg(Var, 13, Vars) + 0.5)
          .Top = Int(ExecuteArg(Var, 14, Vars) + 0.5)
          .Right = Int(ExecuteArg(Var, 15, Vars) + 0.5 + 1)
          .Bottom = Int(ExecuteArg(Var, 16, Vars) + 0.5 + 1)
        End With
      Else
        ProcessRect = RectFromFloatRect(dstRect)
      End If
      MainForm.pTransformBlock dstRect, srcGon, ProcessRect
      If Not MainForm.TransformDataPresent Then NewUndo
End Select
Exit Function
ehDraw:
If Not ErrorOccured Then
    ErrorOccured = True
    MainForm.StartPixelAction
    Resume
Else
    Err.Raise Err.Number, Err.Source, Err.Description
End If
ehLineEx:
PushError
If Not ErrorOccured Then
    ErrorOccured = True
    MainForm.StartPixelAction
    PopError
    Resume
Else
    MainForm.SetLineOpts LOt
    PopError
    ErrRaise
End If
End Function

'argnumber is 1-based
Private Function ExecuteArg(ByRef Args As Variant, _
                            ByVal ArgNumber As Long, _
                            ByRef Vars() As Variable, _
                            Optional ByVal DefValue As Double _
                            ) As Double
Dim i As Long
i = ArgNumber + 2
If i > UBound(Args) Then
  ExecuteArg = DefValue
Else
  ExecuteArg = Execute_Ex(Args(i), Vars())
End If
End Function

Private Function CountVLen(ByRef Args As Variant, ByRef Vars() As Variable)
Dim i As Long
Dim n As Long 'number of args
Dim xArr() As Double
n = UBound(Args) - 2 'because ub=3 is one argument
ReDim xArr(0 To n - 1)
For i = 0 To n - 1
    xArr(i) = Execute_Ex(Args(3 + i), Vars)
Next i
CountVLen = VectorLen(xArr)
End Function

Private Function CountOutput(ByRef Args As Variant, ByRef Vars() As Variable)
Dim i As Long
Dim n As Long 'number of args
Dim xArr() As Double
n = UBound(Args) - 2 'because ub=3 is one argument
ReDim xArr(0 To n - 1)
For i = 0 To n - 1
    xArr(i) = Execute_Ex(Args(3 + i), Vars)
Next i
CountOutput = Output(xArr)
End Function

Private Function Normalize(ByRef Args As Variant, ByRef Vars() As Variable)
Dim i As Long
Dim n As Long 'number of args
Dim xArr() As Double
Dim xLen As Double, invXLen As Double
n = UBound(Args) - 3 + 1 'number of args to process
ReDim xArr(0 To n - 1)

For i = 0 To n - 1
    xArr(i) = Vars(Args(3& + i)).Value
Next i

xLen = VectorLen(xArr)
If xLen = 0 Then
    invXLen = 0
Else
    invXLen = 1 / xLen
End If

For i = 0 To n - 1
    Vars(Args(3& + i)).Value = xArr(i) * invXLen
Next i

Normalize = xLen

End Function

Private Function dbModDL(ByVal Value As Double, ByVal Base As Long) As Double
dbModDL = Int(Value) - Int(Value / Base) * Base
End Function

Private Function dbModDD(ByVal Value As Double, ByVal Base As Double) As Double
dbModDD = Value - Int(Value / Base) * Base
End Function
'Private Function CountSum(ByRef sParams() As String, ByRef Vars() As Variable) As Double
'Static UsedLetter As String
'Dim i As Long, j As Long
'Dim LB As Long, UB As Long
'Dim Rslt As Double
'If UBound(sParams) <> 3 Then
'    Err.Raise 111, "EVal", "Argument count mismatch"
'End If
'On Error Resume Next
'j = UBound(Vars) + 1
'If Err.Number = 0 Then
'    ReDim Preserve Vars(0 To j)
'Else
'    j = 0
'    ReDim Vars(0 To j)
'End If
'On Error GoTo 0
'Vars(j).Name = UCase(sParams(0))
'
'LB = Int(pEVal(sParams(1), Vars))
'UB = Int(pEVal(sParams(2), Vars))
'For i = LB To UB
'    Vars(j).Value = i
'    Rslt = Rslt + pEVal(sParams(3), Vars)
'Next i
'CountSum = Rslt
'End Function

'Private Function GetVar(ByVal VarName As String, Vars() As Variable) As Double
'Dim i As Long
'VarName = UCase$(VarName)
'If VarName = "PI" Then
'    GetVar = 3.14159265358979
'    '        3.1415926535897932
'    Exit Function
'End If
'For i = 0 To UBound(Vars)
'    If VarName = Vars(i).Name Then
'        GetVar = Vars(i).Value
'        Exit Function
'    End If
'Next i
'Err.Raise 111, "EVal", "Unknown Variable: " + VarName
'End Function
'
Private Function GetVarIndex(ByRef VarName As String, ByRef Vars() As Variable) As Integer
Dim i As Long
VarName = UCase$(VarName)
'If VarName = "PI" Then
'    GetVarIndex = -1
'    '        3.1415926535897932
'    Exit Function
'End If
For i = 0 To UBound(Vars)
    If VarName = Vars(i).Name Then
        GetVarIndex = i
        Exit Function
    End If
Next i
Err.Raise 111, "EVal", "Unknown Variable: " + VarName
End Function

Private Function IsDigital(ByVal St As String) As Boolean
Dim i As Long
Dim ts As Double
Dim Rslt As Boolean
Dim mk As String
mk = Mid$(St, 1, 1)
On Error Resume Next
IsDigital = (mk >= "0" And mk <= "9") '(Mid$(St, 1, 1) >= "0" And Mid(St, 1, 1) <= "9")
End Function

Private Function IsOperator(ByVal St As String) As Boolean
IsOperator = (St = "+" _
           Or St = "-" _
           Or St = "*" _
           Or St = "/" _
           Or St = "=" _
           Or St = "<" _
           Or St = ">" _
           Or St = ";" _
              )
End Function

Private Function GetOpCode(ByVal St As String) As Operation
Select Case St
    Case "+"
        GetOpCode = dbOpPlus
    Case "-"
        GetOpCode = dbOpMinus
    Case "/"
        GetOpCode = dbOpDiv
    Case "*"
        GetOpCode = dbOpMult
    Case "="
        GetOpCode = dbOpEqv
    Case ">"
        GetOpCode = dbOpAbove
    Case "<"
        GetOpCode = dbOpBelow
    Case ";"
        GetOpCode = dbOpSemicolon
End Select
End Function

Private Sub PrepareExpr(ByRef St As String)
Dim i As Long
St = Replace(St, " ", "")
St = Replace(St, "[", "(")
St = Replace(St, "]", ")")
St = Replace(St, "{", "(")
St = Replace(St, "}", ")")
St = Replace(St, Chr$(9), "") 'tabs
St = Replace(St, vbCrLf, "") 'line feeds
i = InStr(1, St, ")(")
If i > 0 Then
    Err.Raise 111, "CompileExpression", "Missing operator." + vbNewLine + "..." + MidFT(St, i - 10, i + 10) + "..."
End If
End Sub


'Record structure
'(-1) - Operation
'(0) - sign
'   true=-; false=+
'(1) - type of expr
'   0 - value
'   1 - expression
'   2 - function
'   3 - variable
'if type= 0 then
'   (2) = value
'elseif type=1 then
'   (2) = array containing the expression
'elseif type=2 then
'   (2) = index of a function
'   (3..) = expressions for parameters
'elseif type=3 then
'   (2) = index of a variable
'end if

Friend Function CompileExpression_Ex(ByRef St As String, ByRef Data As Variant, ByRef Vars() As Variable)
Dim i As Long, j As Long, h As Long
Dim m As String
Dim Expr As String
Dim Ops() As Oper
Dim Ops2() As Oper
Dim ParCount As Long
Dim v As Double
Dim sParams() As String, dParams() As Double
Dim FuncName As String, FuncArgs As String
Dim k As Long
Dim mLast As String, mLast2 As String
Dim bComment As Boolean
'Dim tData As Variant
PrepareExpr St
If Len(St) = 0 Then
    ReDim Data(0 To 0)
    RedimVar Data(0), -1, 2
    Data(0)(-1) = CInt(dbOpPlus)
    Data(0)(0) = False
    Data(0)(1) = CInt(0)
    Data(0)(2) = CDbl(0)
    Exit Function
End If
m = Mid$(St, 1, 1)
If m <> "-" And m <> "+" Then
    St = "+" + St
    m = "+"
End If
ReDim Data(0 To 0)
ReDim Ops(0 To 0)
Ops(0).OpCode = GetOpCode(m)
For i = 2 To Len(St)
    If Mid$(St, i, 1) = "'" Then
        bComment = Not (bComment)
    ElseIf Not bComment Then
        mLast2 = mLast
        mLast = m
        m = Mid$(St, i, 1)
        If m = "(" Then ParCount = ParCount + 1
        If m = ")" Then ParCount = ParCount - 1
        If ParCount < 0 Then
            Err.Raise 111, "CompileExpression_Ex", "Parenthesis error. Too many closed." + vbNewLine + "..." + MidFT(St, i - 10, i + 10) + "..."
        End If
        If IsOperator(m) And ParCount = 0 And Not ((m = "+" Or m = "-") And (UCase$(mLast) = "E" And IsDigital(mLast2))) Then
'            If i = Len(St) Then
'                Err.Raise 111, "CompileExpression_Ex", "Incorrect Expression. Operator must be followed with an expression." + vbNewLine + "..." + MidFT(St, i - 10, i + 10) + "..."
'            End If
            Ops(j).opExpr = Expr
            If Len(Expr) = 0 Then
                If Not Ops(j).OpCode = Operation.dbOpSemicolon Then
                    Err.Raise 111, "CompileExpression_Ex", "Incorrect operators combination." + vbNewLine + "..." + MidFT(St, i - 10, i + 10) + "..."
                End If
            End If
            Expr = vbNullString
            j = j + 1
            ReDim Preserve Ops(0 To j)
            Ops(j).OpCode = GetOpCode(m)
        Else
            Expr = Expr + m
        End If
    End If
Next i
If ParCount <> 0 Then
    Err.Raise 111, "CompileExpression_Ex", "Parenthesis error. Not enough closing ones." + vbNewLine + "..." + MidFT(St, i - 10, i + 10) + "..."
End If
Ops(j).opExpr = Expr
ReDim Data(0 To UBound(Ops))

For i = 0 To j
    If Len(Ops(i).opExpr) = 0 Then
        Data(i) = Empty
    ElseIf IsDigital(Ops(i).opExpr) Then
        RedimVar Data(i), -1, 2
        Data(i)(-1) = CInt(Ops(i).OpCode)
        Data(i)(0) = CBool(False)
        Data(i)(1) = CInt(0)
        Data(i)(2) = CDbl(Val(Ops(i).opExpr))
    Else
        h = InStr(1, Ops(i).opExpr, "(")
        If h = 0 Then
            RedimVar Data(i), -1, 2
            Data(i)(-1) = CInt(Ops(i).OpCode)
            Data(i)(0) = CBool(False)
            Data(i)(1) = CInt(3)
            Data(i)(2) = CInt(GetVarIndex(Ops(i).opExpr, Vars))
        ElseIf h = 1 Then
            'Ops(i).opV = pEVal(Mid$(Ops(i).opExpr, 2, Len(Ops(i).opExpr) - 2), Vars)
            RedimVar Data(i), -1, 2
            Data(i)(-1) = CInt(Ops(i).OpCode)
            Data(i)(0) = CBool(False)
            Data(i)(1) = CInt(1)
            CompileExpression_Ex Mid$(Ops(i).opExpr, 2, Len(Ops(i).opExpr) - 2), Data(i)(2), Vars
        ElseIf h > 1 Then
            FuncName = Mid$(Ops(i).opExpr, 1, h - 1)
            FuncArgs = Mid$(Ops(i).opExpr, h + 1, Len(Ops(i).opExpr) - h - 1)
            'sParams = Split(FuncArgs, ",")
            
            ParCount = 0
            k = 0
            ReDim sParams(0 To k)
            For h = 1 To Len(FuncArgs)
                m = Mid$(FuncArgs, h, 1)
                If m = "(" Then ParCount = ParCount + 1
                If m = ")" Then ParCount = ParCount - 1
                If m = "," And ParCount = 0 Then
                    k = k + 1
                    ReDim Preserve sParams(0 To k)
                Else
                    sParams(k) = sParams(k) + m
                End If
            Next h
            
            'If UCase(FuncName) <> "SUM" Then
            '    ReDim dParams(0 To UBound(sParams))
            '
            '    For h = 0 To UBound(sParams)
            '        dParams(h) = pEVal(sParams(h), Vars)
            '    Next h
            '    Ops(i).opV = CountFunction(FuncName, dParams)
            'Else
            '    Ops(i).opV = CountSum(sParams, Vars)
            'End If
            
            'unusual functions are listed here
            'unusual functions are those which need a byref variable, _
             e.g. "sum", or which are executed at compile time
            If UCase$(FuncName) = "SUM" Then
                RedimVar Data(i), -1, 3 + k
                Data(i)(-1) = CInt(Ops(i).OpCode)
                Data(i)(0) = CBool(False)
                Data(i)(1) = CInt(2)
                Data(i)(2) = CInt(GetFunctionIndex(FuncName, k))
                Data(i)(3) = CInt(CreateVar(sParams(0), Vars, False, , FuncName))
                For h = 1 To UBound(sParams)
                    CompileExpression_Ex sParams(h), Data(i)(3 + h), Vars
                Next h
            ElseIf UCase$(FuncName) = "ASSIGN" Then
                RedimVar Data(i), -1, 3 + k
                Data(i)(-1) = CInt(Ops(i).OpCode)
                Data(i)(0) = CBool(False)
                Data(i)(1) = CInt(2)
                Data(i)(2) = CInt(GetFunctionIndex(FuncName, k))
                Data(i)(3) = CInt(CreateVar(sParams(0), Vars, False, , FuncName))
                For h = 1 To UBound(sParams)
                    CompileExpression_Ex sParams(h), Data(i)(3 + h), Vars
                Next h
            
            ElseIf UCase$(FuncName) = "NORMALIZE" Then
                RedimVar Data(i), -1, 3 + k
                Data(i)(-1) = CInt(Ops(i).OpCode)
                Data(i)(0) = CBool(False)
                Data(i)(1) = CInt(2)
                Data(i)(2) = CInt(GetFunctionIndex(FuncName, k))
                For h = 0 To UBound(sParams)
                    Data(i)(3 + h) = CInt(CreateVar(sParams(h), Vars, False, True, FuncName))
                Next h
           
            ElseIf UCase$(FuncName) = "INIT" Then
                RedimVar Data(i), -1, 4
                Data(i)(-1) = CInt(Ops(i).OpCode)
                Data(i)(0) = CBool(False)
                Data(i)(1) = CInt(2)
                Data(i)(2) = CInt(GetFunctionIndex(FuncName, k))
                Data(i)(3) = CInt(CreateVar(sParams(0), Vars, False, , FuncName))
                CompileExpression_Ex sParams(1), Data(i)(4), Vars
                Vars(Data(i)(3)).Value = Execute_Ex(Data(i)(4), Vars)
            Else
                RedimVar Data(i), -1, 3 + k
                Data(i)(-1) = CInt(Ops(i).OpCode)
                Data(i)(0) = CBool(False)
                Data(i)(1) = CInt(2)
                Data(i)(2) = CInt(GetFunctionIndex(FuncName, k))
                For h = 0 To UBound(sParams)
                    CompileExpression_Ex sParams(h), Data(i)(3 + h), Vars
                Next h
            End If
            Erase dParams, sParams
        End If
    End If
Next i

For i = 0 To UBound(Data)
    If Not IsEmpty(Data(i)) Then
        If Data(i)(-1) = dbOpMinus Then
            Data(i)(-1) = dbOpPlus
            Data(i)(0) = Not CBool(Data(i)(0))
        End If
    End If
Next i

End Function

Private Function CreateVar(ByRef VarName As String, _
                           ByRef Vars() As Variable, _
                           Optional ByVal ForceCreate As Boolean = True, _
                           Optional ByVal ForceExist As Boolean = False, _
                           Optional ByRef FuncName As String = "<missing>") As Integer
Dim i As Long
Dim UB As Long
Dim tmp As String
tmp = UCase$(VarName)
UB = -1
If AryDims(AryPtr(Vars)) = 1 Then
    UB = UBound(Vars)
End If

If IsDigital(VarName) Or InStr(VarName, "(") Or InStr(VarName, ")") Or _
        InStr(VarName, "-") Or InStr(VarName, "+") Or _
        InStr(VarName, "*") Or InStr(VarName, "/") Or _
        InStr(VarName, "=") Then
    Err.Raise 111, "CreateVar", "Incorrect variable name (" + VarName + ")."
End If

For i = 0 To UB
    If Vars(i).Name = tmp Then
        If ForceCreate Then
            Err.Raise 111, "CreateVar", "Variable " + VarName + " already exists. You must use a non-existing name in " + FuncName + " function."
        Else
            Exit For
        End If
    End If
Next i
If i = UB + 1 And ForceExist Then
    Err.Raise 111, "CreateVar", "Variable " + VarName + " does not exist. You must specify the existing one for " + FuncName + " function."
End If
If i = UB + 1 Then
    ReDim Preserve Vars(0 To UB + 1)
    Vars(UB + 1).Name = tmp
    Vars(UB + 1).Value = 0
    CreateVar = UB + 1
Else
    CreateVar = i
End If
End Function


Private Sub SwapOps(ByRef Op1 As Oper, ByRef Op2 As Oper)
Dim Op3 As Oper
Op3 = Op1
Op1 = Op2
Op2 = Op3
End Sub

Private Sub RedimVar(ByRef Var As Variant, ByVal pLBound As Long, ByVal pUBound As Long)
ReDim Var(pLBound To pUBound) As Variant
End Sub

Friend Sub CompileExpression(ByVal St As String, Data As Variant, ParamArray Vars() As Variant)
Dim i As Long
Dim pVars() As Variable
If UBound(Vars) >= 0 Then
    ReDim pVars(0 To (UBound(Vars) + 1) \ 2 - 1)
    For i = 0 To UBound(pVars)
        pVars(i).Name = UCase$(CStr(Vars(i * 2)))
        pVars(i).Value = CDbl(Vars(i * 2 + 1))
    Next i
End If
CompileExpression_Ex St, Data, pVars
End Sub


Friend Function Execute_Ex(ByRef Data As Variant, ByRef Vars() As Variable) As Double
'Const MaxOpCount = 128
Dim i As Long
Dim UB As Long
'Dim d As Double
'Dim v1 As Variant, v2 As Variant
'Dim h As Long
Dim Dbls() As Double
Dim Ops() As Operation
'Dim Dbls(0 To MaxOpCount) As Double
'Dim Dbls1(0 To MaxOpCount) As Double
'Dim Ops(0 To MaxOpCount) As Operation
Dim Res As Double
'Dim t As Integer
Dim j As Long
Static CallCount As Long
Static MMME As Boolean
Static Rec As Boolean
Dim RmRec As Boolean
'ReDim v1(0 To 0) As Variant

'On Error GoTo eh

If Not Rec Then
    MMME = MainForm.MeEnabled
    Rec = True
    RmRec = True
End If

If IsEmpty(Data) Then
    Res = 0
Else
    UB = UBound(Data)
    ReDim Dbls(0 To UB)
    ReDim Ops(0 To UB)
    For i = 0 To UB
        If Not IsEmpty(Data(i)) Then
            Ops(i) = Data(i)(-1)
            Select Case Data(i)(1)
                Case 0
                    Dbls(i) = Data(i)(2)
                Case 1
                    Dbls(i) = Execute_Ex(Data(i)(2), Vars)
                Case 2
                    Dbls(i) = CountFunctionCompiled(Data(i), Vars)
                Case 3
                    Dbls(i) = Vars(Data(i)(2)).Value
            End Select
            If Data(i)(0) Then
                Dbls(i) = -Dbls(i)
            End If
        End If
    Next i
    Res = ProcessPriorityBlock(Dbls, Ops, j, 0, UB)
End If

Execute_Ex = Res

If RmRec Then
    Rec = False
    MainForm.MeEnabled = MMME
End If

CallCount = CallCount + 1
If CallCount > 100 Then
    If BreakKeyPressed Then
        Err.Raise dbCWS
    End If
    CallCount = 0
End If
Exit Function
eh:
    Rec = False
    MainForm.MeEnabled = MMME
    Err.Raise Err.Number, Err.Source, Err.Description
Exit Function
End Function

Friend Function Execute(ByRef Data As Variant, ParamArray pVars() As Variant) As Double
Dim i As Long
Dim Vars() As Variable
If UBound(pVars) >= 0 Then
    ReDim Vars(0 To (UBound(pVars) + 1) \ 2 - 1)
    For i = 0 To UBound(Vars)
        Vars(i).Name = UCase$(CStr(pVars(i * 2)))
        Vars(i).Value = CDbl(pVars(i * 2 + 1))
    Next i
End If
Execute = Execute_Ex(Data, Vars)
End Function


Friend Function BuildVars(ByVal AddStandVars As Boolean, ParamArray pVars() As Variant) As Variable()
Dim i As Long
Dim Vars() As Variable
If UBound(pVars) >= 0 Then
    ReDim Vars(0 To (UBound(pVars) + 1) \ 2 - 1)
    For i = 0 To UBound(Vars)
        Vars(i).Name = UCase$(CStr(pVars(i * 2)))
        Vars(i).Value = CDbl(pVars(i * 2 + 1))
    Next i
End If
If AddStandVars Then
    MatVars Vars()
End If
BuildVars = Vars
End Function

Friend Sub ExtractWaveSources(ByRef pWaveSources() As typWaveSource, Optional ByVal EraseOld As Boolean = True)
    pWaveSources = WaveSources
End Sub

Friend Sub EraseWaveSources()
    Erase WaveSources
End Sub

Friend Sub MatVars(ByRef Vars() As Variable)
Dim i As Long

i = CreateVar("PI", Vars, False)
Vars(i).Value = 3.14159265358979

i = CreateVar("E", Vars, False)
Vars(i).Value = 2.71828182845905
End Sub

Private Sub Class_Initialize()
ReDim Priorities(0 To 99)
Priorities(Operation.dbOpSemicolon) = 0

Priorities(Operation.dbOpAbove) = 1
Priorities(Operation.dbOpBelow) = 1
Priorities(Operation.dbOpEqv) = 0

Priorities(Operation.dbOpPlus) = 2
Priorities(Operation.dbOpMinus) = 2

Priorities(Operation.dbOpMult) = 3
Priorities(Operation.dbOpDiv) = 3

End Sub

Private Function ProcessPriorityBlock(ByRef Dbls() As Double, _
                                      ByRef Ops() As Operation, _
                                      ByRef i As Long, _
                                      Prior As Integer, _
                                      ByVal UB As Long) As Double
'Dim UB As Long
Dim Result As Double
Dim NextVal As Double
Dim curOp As Operation
'UB = UBound(Dbls)

curOp = Ops(i)
GoSub GetNextVal
Result = NextVal

Do
    If i + 1 > UB Then
        ProcessPriorityBlock = Result
        Exit Function
    ElseIf Priorities(Ops(i + 1)) < Prior Then
        ProcessPriorityBlock = Result
        Exit Function
    End If
    
    curOp = Ops(i + 1)
    i = i + 1
    GoSub GetNextVal
    
    Select Case curOp
        Case Operation.dbOpSemicolon
            Result = NextVal
            
        Case Operation.dbOpAbove
            Result = Result > NextVal
        Case Operation.dbOpBelow
            Result = Result < NextVal
        Case Operation.dbOpEqv
            Result = Result = NextVal
        
        Case Operation.dbOpPlus
            Result = Result + NextVal
        Case Operation.dbOpMinus
            Result = Result - NextVal
            
        Case Operation.dbOpMult
            Result = Result * NextVal
        Case Operation.dbOpDiv
            Result = Result / NextVal
            
    End Select
    
Loop

Exit Function
GetNextVal:
    If i + 1 > UB Then
        NextVal = Dbls(i)
    ElseIf Priorities(Ops(i + 1)) > Prior Then
        NextVal = ProcessPriorityBlock(Dbls, Ops, i, Prior + 1, UB)
    ElseIf Priorities(Ops(i + 1)) <= Prior Then
        NextVal = Dbls(i)
    End If
Return


End Function

Friend Function ExecuteSMP(ByRef aProg As SMP) As Double
ExecuteSMP = Execute_Ex(aProg.Code, aProg.Vars)
End Function

Friend Sub CompileSMP(ByRef CodeText As String, ByRef aProg As SMP)
CompileExpression_Ex CodeText, aProg.Code, aProg.Vars
End Sub

Public Function MidFT(ByRef St As String, ByVal nStart As Long, Optional ByVal nEnd As Long = -1) As String
If nEnd = -1 Then nEnd = Len(St)
If nStart < 1 Then nStart = 1
If nEnd > Len(St) Then nEnd = Len(St)
If nStart > nEnd Then
    MidFT = ""
    Exit Function
End If
MidFT = Mid$(St, nStart, nEnd - nStart + 1)
End Function

Public Sub NewUndo()
UndoBegan = False
End Sub

Public Sub CheckUndo()
If (Not UndoBegan And Not DrawTemp) Then 'And MainForm.CurUndoType <> dbUndoPixels Then
    MainForm.StartPixelAction
    UndoBegan = True
End If
End Sub

Private Function Arg(ByVal x As Double, ByVal y As Double) As Double
Dim Rslt As Double
If x = 0# And y = 0# Then
    Rslt = 0#
ElseIf Abs(x) > Abs(y) Then
    Rslt = Atan2(x, y)
Else
    Rslt = Pi * 0.5 - Atan2(y, x)
End If
Arg = Rslt + Int((-Rslt + Pi) * 0.5 / Pi) * 2# * Pi
End Function

Private Function Atan2(ByVal x As Double, ByVal y As Double) As Double
If x > 0# Then
    Atan2 = Atn(y / x)
Else
    Atan2 = Pi + Atn(y / x)
End If
End Function

Private Function XYLen(ByVal x As Double, ByVal y As Double) As Double
XYLen = Sqr(x * x + y * y)
End Function

Private Function VectorLen(ByRef xArr() As Double) As Double
Dim SquaredSum As Double
Dim i As Long
SquaredSum = 0
For i = LBound(xArr) To UBound(xArr)
    SquaredSum = SquaredSum + xArr(i) * xArr(i)
Next i
VectorLen = Sqr(SquaredSum)
End Function


Private Function Output(ByRef ArgArray() As Double) As Double
Dim St() As String
Dim i As Long
If Not Bytes.Visible Then Bytes.Show , MainForm
ReDim St(LBound(ArgArray) To UBound(ArgArray))
For i = LBound(ArgArray) To UBound(ArgArray)
  St(i) = CStr(ArgArray(i))
Next i
Bytes.Output Join(St, vbTab) + vbNewLine

End Function
