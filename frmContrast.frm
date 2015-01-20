VERSION 5.00
Begin VB.Form frmContrast 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Color Graph"
   ClientHeight    =   7950
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6450
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   204
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmContrast.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   530
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   430
   StartUpPosition =   1  'CenterOwner
   Begin SMBMaker.ctlResourcizator Resr1 
      Left            =   5325
      Top             =   1695
      _ExtentX        =   1164
      _ExtentY        =   741
      ResID           =   9907
   End
   Begin SMBMaker.dbFrame dbFrame1 
      Height          =   1725
      Left            =   3975
      TabIndex        =   12
      Top             =   2370
      Visible         =   0   'False
      Width           =   2400
      _ExtentX        =   4233
      _ExtentY        =   3043
      EAC             =   0   'False
      Begin VB.HScrollBar HScroll1 
         Height          =   240
         LargeChange     =   10
         Left            =   0
         Max             =   256
         Min             =   4
         TabIndex        =   18
         Top             =   255
         Value           =   4
         Width           =   2340
      End
      Begin VB.TextBox txtRec 
         Height          =   315
         Left            =   1125
         TabIndex        =   14
         Text            =   "1"
         ToolTipText     =   "Number of times to apply sinusoid graph to itself."
         Top             =   930
         Width           =   1140
      End
      Begin VB.TextBox txtPow 
         Height          =   315
         Left            =   1125
         TabIndex        =   13
         Text            =   "1"
         ToolTipText     =   "The power of the sinusoida."
         Top             =   525
         Width           =   1140
      End
      Begin SMBMaker.dbButton btnOK2 
         Height          =   330
         Left            =   390
         TabIndex        =   17
         ToolTipText     =   "Hide this small window."
         Top             =   1305
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   582
         MouseIcon       =   "frmContrast.frx":0442
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Others          =   $"frmContrast.frx":045E
         OthersPresent   =   -1  'True
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Accuracy"
         Height          =   195
         Left            =   840
         TabIndex        =   19
         Top             =   45
         Width           =   675
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Recurse"
         Height          =   195
         Left            =   30
         TabIndex        =   16
         Top             =   960
         Width           =   585
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Power"
         Height          =   195
         Left            =   30
         TabIndex        =   15
         Top             =   570
         Width           =   450
      End
   End
   Begin SMBMaker.dbFrame dbFrame2 
      Height          =   690
      Left            =   105
      TabIndex        =   23
      Top             =   5520
      Width           =   4740
      _ExtentX        =   8361
      _ExtentY        =   1217
      Caption         =   "Interpolation"
      EAC             =   0   'False
      Begin VB.OptionButton optInter 
         BackColor       =   &H00FFC0FF&
         Caption         =   "Polynomial"
         Height          =   315
         Index           =   1
         Left            =   2355
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "The interpolation polynom is used as a function. Switches to linear on large number of nodes."
         Top             =   285
         Width           =   2205
      End
      Begin VB.OptionButton optInter 
         BackColor       =   &H00FFC0FF&
         Caption         =   "Linear"
         Height          =   315
         Index           =   0
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "The color changes linearly from node to node"
         Top             =   285
         Value           =   -1  'True
         Width           =   2205
      End
   End
   Begin SMBMaker.dbFrame dbFrame3 
      Height          =   1290
      Left            =   105
      TabIndex        =   3
      Top             =   60
      Width           =   4200
      _ExtentX        =   7408
      _ExtentY        =   2275
      Caption         =   "Component"
      EAC             =   0   'False
      Begin VB.CheckBox chkTogether 
         BackColor       =   &H00FFFFC0&
         Caption         =   "All together"
         Height          =   825
         Left            =   2460
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   240
         Width           =   1125
      End
      Begin VB.OptionButton Cmp 
         BackColor       =   &H00FF0000&
         Caption         =   "Blue"
         Height          =   285
         Index           =   2
         Left            =   495
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Edit graph for blue component."
         Top             =   840
         Width           =   1725
      End
      Begin VB.OptionButton Cmp 
         BackColor       =   &H0000FF00&
         Caption         =   "Green"
         Height          =   285
         Index           =   1
         Left            =   495
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Edit graph for green component."
         Top             =   555
         Width           =   1725
      End
      Begin VB.OptionButton Cmp 
         BackColor       =   &H000000FF&
         Caption         =   "Red"
         Height          =   285
         Index           =   0
         Left            =   495
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Edit graph for red component."
         Top             =   270
         Value           =   -1  'True
         Width           =   1725
      End
      Begin SMBMaker.dbButton btnCopyCmp 
         Height          =   285
         Index           =   0
         Left            =   90
         TabIndex        =   9
         ToolTipText     =   "Copy current Graph to Red"
         Top             =   270
         Width           =   360
         _ExtentX        =   635
         _ExtentY        =   503
         MouseIcon       =   "frmContrast.frx":04AD
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Wingdings 3"
            Size            =   12
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Others          =   $"frmContrast.frx":04C9
         OthersPresent   =   -1  'True
      End
      Begin SMBMaker.dbButton btnCopyCmp 
         Height          =   285
         Index           =   1
         Left            =   90
         TabIndex        =   10
         ToolTipText     =   "Copy current Graph to Green"
         Top             =   555
         Width           =   360
         _ExtentX        =   635
         _ExtentY        =   503
         MouseIcon       =   "frmContrast.frx":0514
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Wingdings 3"
            Size            =   12
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Others          =   $"frmContrast.frx":0530
         OthersPresent   =   -1  'True
      End
      Begin SMBMaker.dbButton btnCopyCmp 
         Height          =   285
         Index           =   2
         Left            =   90
         TabIndex        =   11
         ToolTipText     =   "Copy current Graph to Blue"
         Top             =   840
         Width           =   360
         _ExtentX        =   635
         _ExtentY        =   503
         MouseIcon       =   "frmContrast.frx":057B
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Wingdings 3"
            Size            =   12
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Others          =   $"frmContrast.frx":0597
         OthersPresent   =   -1  'True
      End
   End
   Begin VB.PictureBox GraphPic 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3840
      Left            =   105
      ScaleHeight     =   256
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   256
      TabIndex        =   1
      Top             =   1365
      Width           =   3840
      Begin VB.Image Sizer 
         Height          =   120
         Index           =   0
         Left            =   465
         Picture         =   "frmContrast.frx":05E2
         ToolTipText     =   "Узел графика."
         Top             =   1470
         Width           =   120
      End
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Preview:"
      Height          =   195
      Left            =   225
      TabIndex        =   27
      Top             =   6285
      Width           =   780
   End
   Begin SMBMaker.dbButton btnOk 
      Default         =   -1  'True
      Height          =   420
      Left            =   4710
      TabIndex        =   26
      Top             =   120
      Width           =   1605
      _ExtentX        =   2831
      _ExtentY        =   741
      MouseIcon       =   "frmContrast.frx":06E4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Others          =   $"frmContrast.frx":0700
      OthersPresent   =   -1  'True
   End
   Begin VB.Image iPreview 
      Height          =   1335
      Left            =   195
      Top             =   6540
      Width           =   6045
   End
   Begin SMBMaker.dbButton btnGamma 
      Height          =   330
      Left            =   4005
      TabIndex        =   22
      ToolTipText     =   "Click to show the gamma detection dialog."
      Top             =   2745
      Width           =   2310
      _ExtentX        =   4075
      _ExtentY        =   582
      MouseIcon       =   "frmContrast.frx":074C
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Others          =   $"frmContrast.frx":0768
      OthersPresent   =   -1  'True
   End
   Begin SMBMaker.dbButton btnLoad 
      Height          =   360
      Left            =   3975
      TabIndex        =   21
      ToolTipText     =   "Load a graph from file."
      Top             =   4830
      Width           =   2385
      _ExtentX        =   4207
      _ExtentY        =   635
      MouseIcon       =   "frmContrast.frx":07C3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Others          =   $"frmContrast.frx":07DF
      OthersPresent   =   -1  'True
   End
   Begin SMBMaker.dbButton btnSave 
      Height          =   360
      Left            =   3975
      TabIndex        =   20
      ToolTipText     =   "Save current graph into a file."
      Top             =   4470
      Width           =   2385
      _ExtentX        =   4207
      _ExtentY        =   635
      MouseIcon       =   "frmContrast.frx":083B
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Others          =   $"frmContrast.frx":0857
      OthersPresent   =   -1  'True
   End
   Begin SMBMaker.dbButton btnCreateSine 
      Height          =   345
      Left            =   4005
      TabIndex        =   8
      ToolTipText     =   "Click to make scontrast curve."
      Top             =   3195
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   609
      MouseIcon       =   "frmContrast.frx":08B1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Others          =   $"frmContrast.frx":08CD
      OthersPresent   =   -1  'True
   End
   Begin SMBMaker.dbButton btnReset 
      Height          =   345
      Left            =   3960
      TabIndex        =   7
      ToolTipText     =   "Load identity graph."
      Top             =   1365
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   609
      MouseIcon       =   "frmContrast.frx":091C
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Others          =   $"frmContrast.frx":0938
      OthersPresent   =   -1  'True
   End
   Begin SMBMaker.dbButton btnCancel 
      Cancel          =   -1  'True
      Height          =   420
      Left            =   4710
      TabIndex        =   0
      Top             =   570
      Width           =   1605
      _ExtentX        =   2831
      _ExtentY        =   741
      MouseIcon       =   "frmContrast.frx":0988
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Others          =   $"frmContrast.frx":09A4
      OthersPresent   =   -1  'True
   End
   Begin VB.Label lblCoord 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "I, O"
      Height          =   255
      Left            =   75
      TabIndex        =   2
      ToolTipText     =   "I=Component value in current picture. O=component value to put instead."
      Top             =   5220
      Width           =   3840
   End
End
Attribute VB_Name = "frmContrast"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Dim Graph() As PointByte
'Dim rGraph() As PointByte
'Dim gGraph() As PointByte
'Dim bGraph() As PointByte
Public Event Change()

Dim Graphs() As dbGraph
Dim Colors() As Long


Dim CurIndex As Integer
Dim CGE As Long
Dim bLock As Boolean

Dim pCDl As CommonDlg
'Const Pi As Double = 3.14159265358979

Private Sub btnCopyCmp_Click(Index As Integer)
Graphs(Index + 1) = Graphs(CurIndex)
Cmp(Index).Value = True
RaiseEvent Change
End Sub

Private Sub btnCreateSine_Click()
dbFrame1.Visible = True
txtPow_Change
'Dim i As Long
'ReDim Graph(0 To 16)
'For i = 0 To 16
'Graph(i).X = i * 255& \ 16
'Graph(i).Y = 255 * (1 - Cos(Pi * (i / 16))) / 2
'Next i
'gRefr
End Sub

Private Sub btnGamma_Click()
Dim rgbQ As RGBQUAD
Dim xR As Single, xG As Single, xb As Single
Dim Log05 As Single
Dim t As Long
Dim i As Long
Dim bUn As Boolean
Load frmReallyGamma
With frmReallyGamma
    ShowFormModal frmReallyGamma
    If .Tag <> "" Then
        Unload frmReallyGamma
        Exit Sub
    End If
    .GetComps rgbQ
    bUn = CBool(.chkUn.Value)
End With
Unload frmReallyGamma

Log05 = Log(128 / 255)
xR = Log(rgbQ.rgbRed / 255) / Log05
xG = Log(rgbQ.rgbGreen / 255) / Log05
xb = Log(rgbQ.rgbBlue / 255) / Log05
If bUn Then
    xR = 1 / xR
    xG = 1 / xG
    xb = 1 / xb
End If
ReDim Graphs(1).Points(0 To 255)
ReDim Graphs(2).Points(0 To 255)
ReDim Graphs(3).Points(0 To 255)
For i = 0 To 255
    Graphs(1).Points(i).X = i
    Graphs(2).Points(i).X = i
    Graphs(3).Points(i).X = i
    
    t = Round((i / 255) ^ xR * 255)
    If t > 255 Then t = 255
    Graphs(1).Points(i).Y = t
    
    t = Round((i / 255) ^ xG * 255)
    If t > 255 Then t = 255
    Graphs(2).Points(i).Y = t
    
    t = Round((i / 255) ^ xb * 255)
    If t > 255 Then t = 255
    Graphs(3).Points(i).Y = t
Next i

Graphs(1).InterpolationMode = dbIMLinear
Graphs(2).InterpolationMode = dbIMLinear
Graphs(3).InterpolationMode = dbIMLinear
Graphs(1).NeedsInterpolation = True
Graphs(2).NeedsInterpolation = True
Graphs(3).NeedsInterpolation = True
gRefr
RaiseEvent Change
End Sub

Private Sub btnLoad_Click()
Dim File As String
On Error GoTo eh
'With pCDl
'    .Filter = GetDlgFilter(dbGLoad)
'    .FileName = ""
'    .DialogTitle = ""
'    .hWndOwner = hWnd
'    .OpenFlags = cdlOFNFileMustExist
'    .ShowOpen
'    File = .FileName
'    .InitDir = GetDirName(File)
'End With
File = ShowOpenDlg(dbGLoad, Me.hWnd, Purpose:="GRAPH")
LoadGraphFile File, Graphs(CurIndex)
Graphs(CurIndex).NeedsInterpolation = True
'SaveSetting_LastDir
gRefr
RaiseEvent Change
Exit Sub
eh:
If Err.Number = dbCWS Then
    Exit Sub
Else
    MsgBox Err.Description, vbCritical, Err.Source
End If
End Sub

Private Sub btnOK2_Click()
dbFrame1.Visible = False
End Sub

Private Sub btnReset_Click()
'ReDim Graph(0 To 1)
'Graph(0).X = 0
'Graph(0).Y = 0
'Graph(1).X = 255
'Graph(1).Y = 255
InitGraph Graphs(CurIndex).Points
Graphs(CurIndex).NeedsInterpolation = True
gRefr
RaiseEvent Change
End Sub

Public Sub gRefr(Optional ByVal NoCls As Boolean = False)
Dim i As Long
With Graphs(CurIndex)
    If .NeedsInterpolation Then
        InterpolateInt .Points, .Table, .InterpolationMode
        .NeedsInterpolation = False
    End If
End With

If chkTogether.Value = vbChecked Then
    For i = 1 To 3
        If i <> CurIndex Then
            Graphs(i) = Graphs(CurIndex)
        End If
    Next i
End If

If Not NoCls Then
    UpdateBG
End If

For i = 1 To 3
    If i <> CurIndex Then
        DrawGraph Graphs(i), Colors(i)
    End If
Next i

With Graphs(CurIndex)
    SortPointArr .Points
    LoadImages UBound(.Points) + 1
    
'    DrawGraph Graphs(CurIndex), Colors(CurIndex)
    UpdateSizers
    GraphPic.Refresh
End With
End Sub

Public Sub UpdateSizers()
Dim i As Long
With Graphs(CurIndex)
    For i = 0 To UBound(.Points)
        Sizer(i).Move .Points(i).X - Sizer(i).Width \ 2, 255 - .Points(i).Y - Sizer(i).Height \ 2
    Next i
End With
End Sub

Private Sub DrawGraph(ByRef Gph As dbGraph, ByVal lngColor As Long)
Dim i As Integer
Dim tColor As Long
With Gph
'If .InterpolationMode = dbIMLinear Then
'    For i = 0 To UBound(.Points) - 1
'        GraphPic.Line (.Points(i).X, 255 - .Points(i).Y)- _
'                      (.Points(i + 1).X, 255 - .Points(i + 1).Y), lngColor
'    Next i
'Else
    If .NeedsInterpolation Then
        InterpolateInt .Points, .Table, .InterpolationMode
        .NeedsInterpolation = False
    End If
    GraphPic.CurrentX = 0
    GraphPic.CurrentY = 255& - .Table(0)
    tColor = GraphPic.ForeColor
    GraphPic.ForeColor = lngColor
    For i = 0 To 255
        GraphPic.Line -(i, 255& - .Table(i))
    Next i
    GraphPic.ForeColor = tColor
'End If
End With
End Sub

Private Sub SortPointArr(ByRef Arr() As PointByte)
Dim i As Long, j As Long, ts As PointByte, blnSorted As Boolean
For i = UBound(Arr) To 1 Step -1
    blnSorted = True
    For j = 0 To i - 1
        If Arr(j).X > Arr(j + 1).X Then
            blnSorted = False
            ts = Arr(j + 1)
            Arr(j + 1) = Arr(j)
            Arr(j) = ts
        End If
    Next j
    If blnSorted Then Exit For
Next i
End Sub

Private Sub LoadImages(ByVal nCount As Integer)
Dim i As Integer
Static VisCount As Integer
If VisCount = 0 Then VisCount = Sizer.UBound
'If nCount > VisCount Then
    For i = VisCount + 1 To nCount - 1
        On Error Resume Next
        Load Sizer(i)
        On Error GoTo 0
        Sizer(i).Visible = True
    Next i
'ElseIf nCount < VisCount Then
    For i = VisCount To nCount Step -1
        Sizer(i).Visible = False
    Next i
'End If
VisCount = nCount - 1
End Sub

Private Sub btnSave_Click()
Dim File As String
On Error GoTo eh
'With pCDl
'    .Filter = GetDlgFilter(dbGSave)
'    .FileName = ""
'    .DialogTitle = ""
'    .hWndOwner = hWnd
'    .OpenFlags = cdlOFNNoReadOnlyReturn Or cdlOFNOverwritePrompt
'    .ShowSave
'    File = .FileName
'    .InitDir = GetDirName(File)
'End With
File = ShowSaveDlg(dbGSave, Me.hWnd, Purpose:="GRAPH")
SaveGraphToFile Graphs(CurIndex), File

gRefr
Exit Sub
eh:
If Err.Number = dbCWS Then
    Exit Sub
Else
    MsgBox Err.Description, vbCritical, Err.Source
End If
End Sub

Private Sub chkTogether_Click()
Dim Opt As OptionButton
gRefr
For Each Opt In Cmp
    Opt.Enabled = Not chkTogether.Value
Next
RaiseEvent Change
End Sub

Private Sub Cmp_Click(Index As Integer)
If bLock Then Exit Sub

CurIndex = Index + 1
bLock = True
optInter(Graphs(CurIndex).InterpolationMode).Value = True
bLock = False
gRefr

End Sub

Private Sub GraphPic_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
CGE = CreateGraphEntry(Int(X), 255 - Int(Y))
End Sub

Public Function CreateGraphEntry(ByVal X As Byte, ByVal Y As Byte) As Integer
Dim i As Long, j As Long
Dim UB As Long
With Graphs(CurIndex)
    UB = UBound(.Points)
    For i = 0 To UB
        If .Points(i).X > X Then
            Exit For
        ElseIf .Points(i).X = X Then
            CreateGraphEntry = i
            vtBeep
            Exit Function
        End If
    Next i
    Debug.Assert i <= UB
    
    ReDim Preserve .Points(0 To UB + 1)
    For j = UB To i Step -1
        .Points(j + 1) = .Points(j)
    Next j
    .Points(i).X = X
    .Points(i).Y = Y
    .NeedsInterpolation = True
    gRefr
    CreateGraphEntry = i
End With
End Function

Public Sub RemoveGraphEntry(ByVal Index As Integer)
Dim i As Integer, j As Integer
Dim UB As Long

With Graphs(CurIndex)
    UB = UBound(.Points)
    If Index = 0 Or Index = UB Then
        vtBeep
        Exit Sub
    End If
    
    For i = Index To UB - 1
        .Points(i) = .Points(i + 1)
    Next i
    
    ReDim Preserve .Points(0 To UB - 1)
    
    .NeedsInterpolation = True
End With
gRefr
End Sub

Public Sub ChangeGraphEntry(ByVal Index As Integer, ByVal X As Integer, ByVal Y As Integer)
With Graphs(CurIndex)
    If Index = 0 Or Index = UBound(.Points) Then
        .Points(Index).Y = Y
    Else
        If X <= .Points(Index - 1).X Then X = .Points(Index - 1).X + 1
        If X >= .Points(Index + 1).X Then X = .Points(Index + 1).X - 1
        .Points(Index).X = X
        .Points(Index).Y = Y
    End If
    .NeedsInterpolation = True
End With
gRefr
End Sub

Private Sub GraphPic_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 0 Then
    ShowCoord X, 255 - Y
End If
If Button = 1 Then
    If X < 0 Then X = 0
    If X > 255 Then X = 255
    If Y < 0 Then Y = 0
    If Y > 255 Then Y = 255
    ChangeGraphEntry CGE, X, 255 - Y
End If
End Sub

Private Sub GraphPic_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent Change
End Sub

Private Sub HScroll1_Change()
txtPow_Change
End Sub

Private Sub optInter_Click(Index As Integer)
If bLock Then Exit Sub
Graphs(CurIndex).InterpolationMode = Index
Graphs(CurIndex).NeedsInterpolation = True
gRefr
RaiseEvent Change
End Sub

Private Sub Sizer_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
    If Index = 0 Or Index = Sizer.UBound Then Exit Sub
    RemoveGraphEntry Index
End If
If Button = 1 Then
    If Index = 0 Or Index = Sizer.UBound Then
        Sizer(Index).Move Sizer(Index).Left, _
                          Sizer(Index).Top + Y \ Screen.TwipsPerPixelY - Sizer(Index).Height \ 2
    Else
        Sizer(Index).Move Sizer(Index).Left + X \ Screen.TwipsPerPixelX - Sizer(Index).Width \ 2, _
                          Sizer(Index).Top + Y \ Screen.TwipsPerPixelY - Sizer(Index).Height \ 2
    End If
End If
End Sub

Private Sub Sizer_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim nX As Integer, nY As Integer
If Button = 1 Then
    If Index = 0 Or Index = Sizer.UBound Then
        nX = IIf(Index = 0, 0, 255)
    Else
        nX = Sizer(Index).Left + X \ Screen.TwipsPerPixelX
    End If
    nY = Sizer(Index).Top + Y \ Screen.TwipsPerPixelY
    If nX < 0 Then nX = 0
    If nX > 255 Then nX = 255
    If nY < 0 Then nY = 0
    If nY > 255 Then nY = 255
    'Sizer(Index).Move nX - Sizer(Index).Width \ 2, nY - Sizer(Index).Height \ 2
    ChangeGraphEntry Index, nX, 255 - nY
    ShowCoord Graphs(CurIndex).Points(Index).X, Graphs(CurIndex).Points(Index).Y
End If
If Button = 0 Then
    ShowCoord Graphs(CurIndex).Points(Index).X, Graphs(CurIndex).Points(Index).Y
End If
End Sub

Private Sub Sizer_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim nX As Integer, nY As Integer
If Button = 1 Then
    If Index = 0 Or Index = Sizer.UBound Then
        nX = IIf(Index = 0, 0, 255)
    Else
        nX = Sizer(Index).Left + X \ Screen.TwipsPerPixelX
    End If
    nY = Sizer(Index).Top + Y \ Screen.TwipsPerPixelY
    If nX < 0 Then nX = 0
    If nX > 255 Then nX = 255
    If nY < 0 Then nY = 0
    If nY > 255 Then nY = 255
    ChangeGraphEntry Index, nX, 255 - nY
    'Sizer(Index).Move nX - Sizer(Index).Width \ 2, nY - Sizer(Index).Height \ 2
    
End If
RaiseEvent Change
End Sub

Private Sub ShowCoord(ByVal X As Integer, ByVal Y As Integer)
lblCoord.Caption = "I = " + CStr(X) + ",   O = " + CStr(Y)
End Sub

'Private Function ExtractTable(ByRef Grph() As PointByte, ByRef Tbl() As Byte)
'Dim i As Integer, j As Integer
'ReDim Tbl(0 To 255)
'For i = 0 To 255
'    If i = Grph(j + 1).x Then
'        j = j + 1
'        Tbl(i) = Grph(j).y
'    Else
'        Tbl(i) = Grph(j).y + CLng(CLng(Grph(j + 1).y) - CLng(Grph(j).y)) * CLng(i - Grph(j).x) / (Grph(j + 1).x - Grph(j).x)
'    End If
'Next i
'End Function
'



Private Sub btnCancel_Click()
Me.Tag = "C"
Me.Hide
End Sub

Private Sub btnOK_Click()
Me.Tag = ""
Me.Hide
End Sub

Private Sub Form_Load()
Dim i As Long
ReDim Colors(1 To 3)
Colors(1) = &HFF&
Colors(2) = &HFF00&
Colors(3) = &HFF0000

ReDim Graphs(1 To 3)
For i = 1 To 3
    InitGraph Graphs(i).Points
    Graphs(i).NeedsInterpolation = True
Next i

LoadCaptions
Set pCDl = New CommonDlg
pCDl.CancelError = True


'LoadSettings
CurIndex = 1
Cmp(CurIndex + 1).Value = True
'gRefr
GraphPic.Move GraphPic.Left, GraphPic.Top, 256, 256
lblCoord.Move GraphPic.Left, GraphPic.Top + GraphPic.Height, GraphPic.Width
End Sub

Sub UpdateBG()
Dim i As Long, n As Long, n1 As Long
On Error Resume Next
GraphPic.Cls
If AryDims(AryPtr(Graphs(CurIndex).Table)) <> 1 Then Exit Sub
For i = 0 To 255
    n = Graphs(CurIndex).Table(i)
    GraphPic.Line (i, 255 - n)-(i, GraphPic.ScaleHeight), n * &H10101 And Colors(CurIndex)
    n1 = ((n \ 32) * 32 + 128) Mod 256
    GraphPic.Line (i, 255 - n - 1)-(i, -1), n1 * &H10101 And Colors(CurIndex)
Next i
'gRefr True
End Sub

Friend Sub InitGraph(ByRef Grph() As PointByte)
ReDim Grph(0 To 1)
Grph(0).X = 0
Grph(0).Y = 0
Grph(1).X = 255
Grph(1).Y = 255
End Sub

Private Sub Form_Paint()
Me.PaintPicture LoadResPicture(RES_BackPicture, vbResBitmap), 0, 0, Me.ScaleWidth, Me.ScaleHeight
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If UnloadMode = 0 Then
    Cancel = True
    btnCancel_Click
Else
    Set pCDl = Nothing
End If
End Sub

Public Sub LoadCaptions()
Resr1.LoadCaptions
End Sub

Friend Sub GetGraphs(ByRef aGraphs() As dbGraph)
Dim i As Long
For i = 1 To 3
    If Graphs(i).NeedsInterpolation Then
        InterpolateInt Graphs(i).Points, Graphs(i).Table, Graphs(i).InterpolationMode
        Graphs(i).NeedsInterpolation = False
    End If
Next i
aGraphs = Graphs
End Sub

Friend Sub SetGraphs(ByRef aGraphs() As dbGraph)
Dim i As Long
For i = 1 To 3
    Graphs(i) = aGraphs(i)
Next i
End Sub

Private Sub txtPow_Change()
If Not (dbFrame1.Visible) Then Exit Sub
Dim i As Long, Pow As Single, k As Single, j As Integer, Rec As Integer
Dim n As Integer
n = HScroll1.Value
On Error GoTo eh
Pow = CSng(txtPow.Text)
Rec = CInt(txtRec.Text)
If Rec <= 0 Then
    Err.Raise 5, "txtPow_Change", "Rec cannot be less than 1"
End If
With Graphs(CurIndex)
    ReDim .Points(0 To n)
    For i = 0 To n
        .Points(i).X = i * 255& \ n
        k = i / n
        For j = 1 To Val(txtRec.Text)
            k = ((1 - Cos(Pi * (k))) / 2) ^ Pow
        Next j
        .Points(i).Y = 255 * k
    Next i
    .NeedsInterpolation = True
End With
gRefr
RaiseEvent Change
Exit Sub
eh:
vtBeep
End Sub

Private Sub txtRec_Change()
txtPow_Change
End Sub
