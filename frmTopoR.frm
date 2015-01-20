VERSION 5.00
Begin VB.Form frmTopoR 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Render TopoR PCB"
   ClientHeight    =   4961
   ClientLeft      =   2761
   ClientTop       =   3751
   ClientWidth     =   6446
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.79
      Charset         =   204
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4961
   ScaleWidth      =   6446
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox chkNoSpokes 
      Caption         =   "No spokes"
      Height          =   319
      Left            =   715
      TabIndex        =   31
      Top             =   3234
      Width           =   1188
   End
   Begin VB.CheckBox chkRemember 
      Caption         =   "remember settings"
      Height          =   264
      Left            =   3025
      TabIndex        =   30
      Top             =   4609
      Width           =   3190
   End
   Begin VB.Frame Frame2 
      Caption         =   "What to render"
      Height          =   3179
      Left            =   110
      TabIndex        =   4
      Top             =   737
      Width           =   3729
      Begin SMBMaker.ctlColor clrMet 
         Height          =   264
         Left            =   1485
         TabIndex        =   29
         Top             =   2178
         Width           =   264
         _ExtentX        =   467
         _ExtentY        =   467
      End
      Begin SMBMaker.ctlColor clrHoles 
         Height          =   264
         Left            =   3311
         TabIndex        =   28
         Top             =   2706
         Width           =   264
         _ExtentX        =   467
         _ExtentY        =   467
         Color           =   16777215
      End
      Begin SMBMaker.ctlColor clrContour 
         Height          =   264
         Left            =   3311
         TabIndex        =   27
         Top             =   2195
         Width           =   264
         _ExtentX        =   467
         _ExtentY        =   467
      End
      Begin SMBMaker.ctlColor clrVias 
         Height          =   264
         Left            =   3311
         TabIndex        =   26
         Top             =   1687
         Width           =   264
         _ExtentX        =   467
         _ExtentY        =   467
      End
      Begin SMBMaker.ctlColor clrWires 
         Height          =   264
         Left            =   3311
         TabIndex        =   24
         Top             =   671
         Width           =   264
         _ExtentX        =   467
         _ExtentY        =   467
      End
      Begin VB.CheckBox chkElHoles 
         Caption         =   "Holes"
         Height          =   451
         Left            =   1881
         TabIndex        =   23
         Top             =   2629
         Value           =   1  'Checked
         Width           =   1738
      End
      Begin VB.CheckBox chkElContour 
         Caption         =   "Board contour"
         Height          =   451
         Left            =   1881
         TabIndex        =   22
         Top             =   2079
         Value           =   1  'Checked
         Width           =   1738
      End
      Begin VB.CheckBox chkElTexts 
         Caption         =   "Texts"
         Enabled         =   0   'False
         Height          =   451
         Left            =   165
         TabIndex        =   12
         Top             =   2629
         Width           =   1738
      End
      Begin VB.CheckBox chkElVias 
         Caption         =   "Vias"
         Height          =   451
         Left            =   1881
         TabIndex        =   11
         Top             =   1579
         Value           =   1  'Checked
         Width           =   1738
      End
      Begin VB.CheckBox chkElMet 
         Caption         =   "Metallization areas"
         Height          =   451
         Left            =   165
         TabIndex        =   10
         Top             =   2123
         Value           =   1  'Checked
         Width           =   1738
      End
      Begin VB.CheckBox chkElWires 
         Caption         =   "Wires"
         Height          =   451
         Left            =   1881
         TabIndex        =   8
         Top             =   572
         Value           =   1  'Checked
         Width           =   1738
      End
      Begin VB.ListBox lstLayers 
         Height          =   1364
         IntegralHeight  =   0   'False
         Left            =   209
         TabIndex        =   5
         Top             =   539
         Width           =   1595
      End
      Begin SMBMaker.ctlColor clrPads 
         Height          =   264
         Left            =   3311
         TabIndex        =   25
         Top             =   1179
         Width           =   264
         _ExtentX        =   467
         _ExtentY        =   467
      End
      Begin VB.CheckBox chkElPads 
         Caption         =   "Pads"
         Height          =   451
         Left            =   1881
         TabIndex        =   9
         Top             =   1089
         Value           =   1  'Checked
         Width           =   1738
      End
      Begin VB.Label Label3 
         Caption         =   "Elements"
         Height          =   231
         Left            =   1947
         TabIndex        =   7
         Top             =   297
         Width           =   1342
      End
      Begin VB.Label Label2 
         Caption         =   "Layer"
         Height          =   341
         Left            =   220
         TabIndex        =   6
         Top             =   275
         Width           =   1573
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Renderer"
      Height          =   3729
      Left            =   3938
      TabIndex        =   3
      Top             =   825
      Width           =   2244
      Begin VB.OptionButton optSzNeeded 
         Caption         =   "to fit only what's being rendered"
         Height          =   462
         Left            =   165
         TabIndex        =   21
         Top             =   2992
         Value           =   -1  'True
         Width           =   1848
      End
      Begin VB.OptionButton optSzWhole 
         Caption         =   "to fit whole board"
         Height          =   275
         Left            =   165
         TabIndex        =   20
         Top             =   2585
         Width           =   1804
      End
      Begin SMBMaker.ctlNumBox nmbSegments 
         Height          =   275
         Left            =   231
         TabIndex        =   18
         ToolTipText     =   $"frmTopoR.frx":0000
         Top             =   1749
         Width           =   1782
         _ExtentX        =   3150
         _ExtentY        =   488
         Value           =   120
         Min             =   8
         Max             =   1000
         NumType         =   3
         HorzMode        =   0   'False
         EditName        =   "Segments per circle"
         SliderVisible   =   0   'False
      End
      Begin VB.CheckBox chkAntialiasing 
         Caption         =   "Antialiasing"
         Height          =   319
         Left            =   198
         TabIndex        =   15
         ToolTipText     =   "Recommended when for screen viewing. Not recommended when for printing."
         Top             =   1067
         Width           =   1573
      End
      Begin SMBMaker.ctlNumBox nmbResolution 
         Height          =   286
         Left            =   187
         TabIndex        =   14
         ToolTipText     =   "Fractional values are also accepted. For maximum quality, set equal to printer resolution."
         Top             =   638
         Width           =   1034
         _ExtentX        =   1829
         _ExtentY        =   508
         Value           =   600
         Min             =   10
         Max             =   4000
         NumType         =   5
         HorzMode        =   0   'False
         EditName        =   "Resolution in pixels per inch"
         SliderVisible   =   0   'False
      End
      Begin VB.Label Label6 
         Caption         =   "Size picture:"
         Height          =   396
         Left            =   264
         TabIndex        =   19
         Top             =   2277
         Width           =   1551
      End
      Begin VB.Label Label5 
         Caption         =   "Circle N of segments:"
         Height          =   264
         Left            =   220
         TabIndex        =   17
         Top             =   1518
         Width           =   1925
      End
      Begin VB.Label Label4 
         Caption         =   "resolution (DPI):"
         Height          =   231
         Left            =   187
         TabIndex        =   13
         Top             =   330
         Width           =   1441
      End
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "..."
      Height          =   231
      Left            =   5830
      TabIndex        =   2
      Top             =   385
      Visible         =   0   'False
      Width           =   319
   End
   Begin VB.TextBox txtFile 
      BackColor       =   &H8000000F&
      Height          =   352
      Left            =   121
      Locked          =   -1  'True
      TabIndex        =   1
      Text            =   "txtFile"
      Top             =   330
      Width           =   5742
   End
   Begin SMBMaker.dbButton btnRender 
      Height          =   638
      Left            =   132
      TabIndex        =   16
      Top             =   4004
      Width           =   2255
      _ExtentX        =   3983
      _ExtentY        =   1118
      MouseIcon       =   "frmTopoR.frx":009F
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12.096
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Others          =   $"frmTopoR.frx":00BB
      OthersPresent   =   -1  'True
   End
   Begin VB.Label Label1 
      Caption         =   "File to be rendered:"
      Height          =   220
      Left            =   99
      TabIndex        =   0
      Top             =   88
      Width           =   1650
   End
End
Attribute VB_Name = "frmTopoR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim bSuppress As Boolean 'used to suppress saving settings during load

Dim PCB As DOMDocument

Const xqRoot = "TopoR_PCB_File"
Const xqLayers = xqRoot + "/Layers/StackUpLayers/Layer" 'will return layer elements. attributes of interest: name, type
Const xqPadStacks = xqRoot + "/LocalLibrary/Padstacks/Padstack"
Const xqViaStacks = xqRoot + "/LocalLibrary/Viastacks/Viastack"
Const xqWires = xqRoot + "/Connectivity/Wires/Wire"
Const xqComponents = xqRoot + "/ComponentsOnBoard/Components/CompInstance"
Const xqContour = xqRoot + "/Constructive/BoardOutline/Contour/Shape"
Const xqContourVoids = xqRoot + "/Constructive/BoardOutline/Voids/Shape"
Const xqVias = xqRoot + "/Connectivity/Vias/Via"
Const xqCoppers = xqRoot + "/Connectivity/Coppers/Copper"

Dim pDrawMode As eDrawMode
Dim pCurColor As Long
Dim bDraw As Boolean 'when true, DrawElement draws the element,
                     'when false - drawelement fills range vars
'the range vars - to get the picture size needed for the board
'to fit
Dim xminmm As Double, yminmm As Double
Dim xmaxmm As Double, ymaxmm As Double
Dim RangeFoo As Boolean 'false indicates range vars are
                        'uninitialized and have no meaning. _
                        False means some points already have _
                        been considered
Private Type vtTopScale
  PixPerMM As Double
  OrgX As Double 'in pixels
  OrgY As Double 'in pixels
  'orgx and orgy is the pixels position of origin of fst file (where x and y =0)
End Type

Dim CurScale As vtTopScale

Dim pSelectedLayer As String

Dim gPlg As New clsPolygon

Private Sub btnRender_Click()
On Error GoTo eh
Static Rec As Boolean
If Rec Then Exit Sub
Rec = True
'calculate range
RangeFoo = False
bDraw = False
ShowStatus "Calculating size..."
DrawEverything Sizing:=True

'resize pic and set the scale
ShowStatus "Resizing picture..."
PrepareForDraw DoResize:=True

'and, finally, RENDERRR!
bDraw = True
DrawingEngine.AntiAliasingSharpness = IIf(chkAntialiasing.Value = vbChecked, 1, 1000)
DrawEverything
'TODO: if aa, convert to srgb
MainForm.Refr
ShowStatus "TopoR RULEZZZZ!", HoldTime:=5
Rec = False
Exit Sub
eh:
Rec = False
MsgError
End Sub

Private Sub chkRemember_Click()
If chkRemember.Value = vbChecked And Not bSuppress Then SaveSettings
dbSaveSettingEx "TopoRr", "SaveSettings", chkRemember.Value = vbChecked
End Sub

Private Sub cmdBrowse_Click()
On Error GoTo eh
Browse
Exit Sub
eh:
MsgError
End Sub

Public Sub Browse()
Dim FileName As String
FileName = ShowOpenDlg(dbFSTLoad, Me.hWnd, Purpose:="TOPOR")
txtFile.Text = FileName
LoadFST FileName
End Sub

Public Sub LoadFST(ByRef FileName As String)
Set PCB = ParseFST(FileName)
'fill layers list
With lstLayers
  .Clear
  Dim Layers As IXMLDOMNodeList
  Set Layers = PCB.selectNodes(xqLayers)
  Dim i As Long
  For i = 0 To Layers.Length - 1
    If GetAttribute(Layers(i), "type") = "Signal" Then
      .AddItem GetAttribute(Layers(i), "name")
      If .List(.NewIndex) = "Top" Then lstLayers.ListIndex = .NewIndex
    End If
  Next i
End With
'calculate width/height of board. Use board outline for this.

End Sub

Private Function ParseFST(ByRef FileName As String) As DOMDocument
Dim tmp As String
Dim fst As Long
fst = FreeFile
On Error GoTo eh
'read the whole file into a string
Open FileName For Input As fst
tmp = input(LOF(fst), fst)
Close fst

'find
Dim Pos As Long
Pos = InStr(1, tmp, "<TopoR_PCB_File>", vbTextCompare)
If Pos = 0 Then Err.Raise 12345, , "Not TopoR PCB file (<TopoR_PCB_File> element not found)"
Dim xml As New DOMDocument
'load just the interesting part (otherwise xml paser jams at <?xml version=....?>)
xml.loadXML Mid$(tmp, Pos)
Set ParseFST = xml
Exit Function
eh:
Close fst
Err.Raise 12345, , "Not TopoR PCB file (file is binary or I/O error)"
End Function

Private Function GetAttribute(Element As IXMLDOMNode, _
                              AttName As String, _
                              Optional ByVal DefString As Variant) As String
Dim attn As IXMLDOMNode
Set attn = Element.Attributes.getNamedItem(AttName)
If attn Is Nothing Then
  If IsMissing(DefString) Then
    Err.Raise 13411, "GetAttribute", "Attribute " + AttName + " of element <" + Element.nodeName + "> does not exist."
  Else
    GetAttribute = DefString
  End If
Else
  GetAttribute = attn.nodeValue
End If
End Function

Private Sub DrawElement(Element As IXMLDOMNode, _
                             ByVal LineWidthMM As Double)
Dim x1 As Double, y1 As Double
Dim x2 As Double, y2 As Double
Select Case Element.nodeName
  Case "Line", "ThermalSpoke"
    GetXY Element.childNodes(0), x1, y1
    GetXY Element.childNodes(1), x2, y2
    LineWidthMM = Val(GetAttribute(Element, "lineWidth", DefString:=Str(LineWidthMM)))
    DrawLine x1, y1, x2, y2, LineWidthMM
  Case "Rect"
    x1 = Val(GetAttribute(Element.childNodes(0), "x"))
    y1 = Val(GetAttribute(Element.childNodes(0), "y"))
    x2 = Val(GetAttribute(Element.childNodes(1), "x"))
    y2 = Val(GetAttribute(Element.childNodes(1), "y"))
    DrawLine x1, y1, x1, y2, LineWidthMM
    DrawLine x1, y2, x2, y2, LineWidthMM
    DrawLine x2, y2, x2, y1, LineWidthMM
    DrawLine x2, y1, x1, y1, LineWidthMM
  Case "Arc"
    Dim xc As Double, yc As Double
    Dim CNode As IXMLDOMNode
    Dim Node1 As IXMLDOMNode, Node2 As IXMLDOMNode
    Set CNode = Element.selectSingleNode("Center")
    Set Node1 = Element.selectSingleNode("Start")
    Set Node2 = Element.selectSingleNode("End")
    
    xc = Val(GetAttribute(CNode, "x"))
    yc = Val(GetAttribute(CNode, "y"))
    x1 = Val(GetAttribute(Node1, "x"))
    y1 = Val(GetAttribute(Node1, "y"))
    x2 = Val(GetAttribute(Node2, "x"))
    y2 = Val(GetAttribute(Node2, "y"))
    
    DrawArc xc, yc, x1, y1, x2, y2, LineWidthMM
  Case "Circle"
    Dim Diam As Double
    Diam = Val(GetAttribute(Element, "diameter"))
    Set CNode = Element.selectSingleNode("Center")
    GetXY CNode, xc, yc
    DrawArc xc, yc, xc, yc + Diam * 0.5, xc, yc + Diam * 0.5, LineWidthMM
  'Case "CompInstance"
    
  Case "Subwire" 'width is taken from the element, the one supplied as an argument is ignored
    Dim prevx As Double, prevy As Double
    Dim x As Double, y As Double
    Dim n1 As IXMLDOMNode
    Dim Wdt As Double
    Wdt = Val(GetAttribute(Element, "width"))
    
    Dim Tears As IXMLDOMNodeList
    Set Tears = Element.selectNodes("Teardrops/Teardrop")
    Dim it As Long
    For it = 0 To Tears.Length - 1
      DrawFilledShape Tears(it), Wdt
    Next it
    
    Set CNode = Element.selectSingleNode("Start")
    GetXY CNode, prevx, prevy
    Do Until CNode.nextSibling Is Nothing
      Set CNode = CNode.nextSibling
      Select Case CNode.nodeName
        Case "TrackLine"
          GetXY CNode.selectSingleNode("End"), x, y
          DrawLine prevx, prevy, x, y, Wdt
          prevx = x: prevy = y
        Case "TrackArc"
          GetXY CNode.selectSingleNode("End"), x, y
          GetXY CNode.selectSingleNode("Center"), xc, yc
          DrawArc xc, yc, prevx, prevy, x, y, Wdt
          prevx = x: prevy = y
        Case "TrackArcCW"
          GetXY CNode.selectSingleNode("End"), x, y
          GetXY CNode.selectSingleNode("Center"), xc, yc
          DrawArc xc, yc, x, y, prevx, prevy, Wdt
          prevx = x: prevy = y
        Case Else
          Err.Raise 12345, "DrawElement", "Unrecognized track segment: <" + CNode.nodeName + ">"
      End Select
    Loop
  Case "Copper"
    Err.Raise 11211, "DrawElement", "Metallization areas are not supported by now =("
  Case Else
    Err.Raise 11211, "DrawElement", "<" + Element.nodeName + "> not supported!"
End Select
End Sub

'drLayer - the layer that is being drawn
Private Sub DrawComponent(ByRef Comp As IXMLDOMNode)
Dim cx As Double, cy As Double
Dim Bot As Boolean 'true if bomponent is on the other (bottom) side
Dim Side As String
Side = GetAttribute(Comp, "side")
Select Case Side
  Case "Top"
  Case "Bottom"
    Bot = True
  Case Else
    Err.Raise 12121, "DrawComponent", "Side not supported: " + Side
End Select
GetXY Comp.selectSingleNode("Org"), cx, cy
Dim AngDeg As Double, AngRad As Double
AngDeg = Val(GetAttribute(Comp, "angle", "0"))
AngRad = AngDeg / 180 * Pi
Dim Pins As IXMLDOMNodeList
Set Pins = Comp.selectNodes("Pins/Pin")
Dim PinX As Double, PinY As Double
Dim rpx As Double, rpy As Double 'rotated pin coordinates
Dim PSRef As String
Dim i As Long
For i = 0 To Pins.Length - 1
  GetXY Pins(i).selectSingleNode("Org"), PinX, PinY
  If Bot Then PinX = -PinX
  rpx = PinX * Cos(AngRad) - PinY * Sin(AngRad)
  rpy = PinX * Sin(AngRad) + PinY * Cos(AngRad)
  PSRef = GetAttribute(Pins(i).selectSingleNode("PadstackRef"), "name")
  DrawPad PSRef, cx + rpx, cy + rpy, AngDeg, Bot
Next i
End Sub

Private Sub DrawPad(ByRef PSRef As String, _
                    ByVal x As Double, ByVal y As Double, _
                    ByVal AngDeg As Double, _
                    ByVal BotSide As Boolean, _
                    Optional ByVal bVia As Boolean = False)
Dim PS As IXMLDOMNode
Set PS = PadStack(PSRef, bVia)
Dim th As Boolean
th = GetAttribute(PS, "type", "Through") <> "SMD" 'will work for vias, too
Dim HoleDia As Double
HoleDia = Val(GetAttribute(PS, "holeDiameter", "0"))
Dim Pad As IXMLDOMNode
Dim i As Long
'first, search through all strictly defined layers
Dim ltf As String 'Layer to find
ltf = IIf(BotSide Xor pSelectedLayer = "Top", "Top", "Bottom")
Dim xqPads As String
xqPads = IIf(bVia, "ViaPads", "Pads")
Set Pad = PS.selectSingleNode(xqPads + "/*/LayerRef[@name=""" + ltf + """]")
If Pad Is Nothing Then
  'there is no defined pad for the layer. Let's find the one with layer type specified
  'if via, check if current layer is within range
  If bVia Then
    If Not PS.selectSingleNode("LayerRange") Is Nothing Then
      If PS.selectSingleNode("LayerRange/AllLayers") Is Nothing Then
        'test if layer we want is in range
        Dim LayerFrom As String
        Dim LayerTo As String
        Dim TwoLayerRefs As IXMLDOMNodeList
        Set TwoLayerRefs = PS.selectNodes("LayerRange/LayerRef")
        LayerFrom = GetAttribute(TwoLayerRefs(0), "name")
        LayerTo = GetAttribute(TwoLayerRefs(1), "name")
        Dim Layers As IXMLDOMNodeList
        Set Layers = PCB.selectNodes(xqLayers)
        Dim ifrom As Long, icur As Long, ito As Long
        Dim LNm As String
        For i = 0 To Layers.Length - 1
          LNm = GetAttribute(Layers(i), "name")
          If LNm = LayerFrom Then ifrom = i
          If LNm = pSelectedLayer Then icur = i
          If LNm = LayerTo Then ito = i
        Next i
        'if not, nothing to draw...
        If icur < ifrom Or icur > ito Then Exit Sub
      End If
    End If
  End If
  Set Pad = PS.selectSingleNode(xqPads + "/*/LayerTypeRef[@type=""Signal""]")
  Dim PadOnTheWrongSide As Boolean
  PadOnTheWrongSide = BotSide Xor pSelectedLayer <> "Top"
  'do not draw smd pad if component is on the other side
  If PadOnTheWrongSide And Not th Then Set Pad = Nothing
End If
If Pad Is Nothing Then Exit Sub
Set Pad = Pad.parentNode
Dim si As Double, co As Double
si = Sin(AngDeg / 180 * Pi): co = Cos(AngDeg / 180 * Pi)
Dim shiftx As Double, shifty As Double
Dim x1 As Double, y1 As Double
Dim x2 As Double, y2 As Double
Select Case Pad.nodeName
  Case "PadCircle"
    Dim Dia As Double
    Dia = Val(GetAttribute(Pad, "diameter"))
    DrawLine x, y, x * 1.0000000000001, y, Dia
  Case "PadOval"
    Dim vx As Double, vy As Double
    Dia = Val(GetAttribute(Pad, "diameter"))
    GetXY Pad.selectSingleNode("Stretch"), vx, vy
    GetXY Pad.selectSingleNode("Shift"), shiftx, shifty, IgnoreNothing:=True
    If BotSide Then
      shiftx = -shiftx
      vx = -vx
    End If
    x1 = (shiftx - 0.5 * vx) * co + (shifty - 0.5 * vy) * -si
    y1 = (shiftx - 0.5 * vx) * si + (shifty - 0.5 * vy) * co
    x2 = (shiftx + 0.5 * vx) * co + (shifty + 0.5 * vy) * -si
    y2 = (shiftx + 0.5 * vx) * si + (shifty + 0.5 * vy) * co
    
    DrawLine x + x1, y + y1, x + x2, y + y2, Dia
  Case "PadRect"
    Dim w As Double, h As Double
    w = Val(GetAttribute(Pad, "width"))
    h = Val(GetAttribute(Pad, "height"))
    GetXY Pad.selectSingleNode("Shift"), vx, vy, IgnoreNothing:=True
    If BotSide Then
      vx = -vx
    End If
    shiftx = vx * co + vy * -si
    shifty = vx * si + vy * co
    DrawRect x + shiftx, y + shifty, w, h, AngDeg
  Case Else
    Err.Raise 12345, "DrawPad", "Pad of type " + Pad.nodeName + " not supported"
End Select
If HoleDia > 0 And chkElHoles.Value = vbChecked Then
  Dim tmp As Long
  tmp = pCurColor
  pCurColor = clrHoles.Color
  DrawLine x, y, x * 1.0000000000001, y, HoleDia
  pCurColor = tmp
End If
End Sub

Private Sub GetXY(Node As IXMLDOMNode, ByRef x As Double, ByRef y As Double, Optional ByVal IgnoreNothing As Boolean = False)
If Node Is Nothing And IgnoreNothing Then
  x = 0
  y = 0
Else
  x = Val(GetAttribute(Node, "x"))
  y = Val(GetAttribute(Node, "y"))
End If
End Sub

Private Function PadStack(ByRef PSName As String, _
                          Optional ByVal bVia As Boolean = False _
                          ) As IXMLDOMNode
Dim ret As IXMLDOMNode
If bVia Then
  Set ret = PCB.selectSingleNode(xqViaStacks + "[@name=""" + PSName + """]")
  If ret Is Nothing Then Err.Raise 12431, "PadStack", "Padstack " + PSName + " not found"
  Set PadStack = ret
Else
  Set ret = PCB.selectSingleNode(xqPadStacks + "[@name=""" + PSName + """]")
  If ret Is Nothing Then Err.Raise 12431, "PadStack", "Padstack " + PSName + " not found"
  Set PadStack = ret
End If
End Function

Private Function PixFromMMX(ByVal xmm As Double) As Double
PixFromMMX = xmm * CurScale.PixPerMM + CurScale.OrgX
End Function

Private Function PixFromMMY(ByVal ymm As Double) As Double
PixFromMMY = -ymm * CurScale.PixPerMM + CurScale.OrgY
End Function

Private Sub DrawLine(ByVal x1mm As Double, ByVal y1mm As Double, _
                          ByVal x2mm As Double, ByVal y2mm As Double, _
                          ByVal LineWidthMM As Double)
If bDraw Then
  Dim Pixels() As AlphaPixel
  Dim nPix As Long
  Dim p1 As vtVertex, p2 As vtVertex
  p1.x = PixFromMMX(x1mm)
  p1.y = PixFromMMY(y1mm)
  p1.Color = pCurColor
  p1.Weight = LineWidthMM * CurScale.PixPerMM
  
  p2.x = PixFromMMX(x2mm)
  p2.y = PixFromMMY(y2mm)
  p2.Color = pCurColor
  p2.Weight = LineWidthMM * CurScale.PixPerMM
  Dim FDSC As FadeDesc
  FDSC.Power = 1
  FDSC.Mode = dbFLinear
  DrawingEngine.pntGradientLineHQ p1, p2, FDSC, Pixels, nPix
  MainForm.dbDrawPixels Pixels, nPix, DrawTemp:=False, _
                        ForceDraw:=False, DrawMode:=pDrawMode, _
                        KillAA:=chkAntialiasing.Value <> vbChecked
Else
  RangePoint Min(x1mm, x2mm) - LineWidthMM * 0.5, Min(y1mm, y2mm) - LineWidthMM * 0.5
  RangePoint Max(x1mm, x2mm) + LineWidthMM * 0.5, Max(y1mm, y2mm) + LineWidthMM * 0.5
End If
End Sub

'if start and end match, draws a full circle
Private Sub DrawArc(ByVal cxmm As Double, ByVal cymm As Double, _
                         ByVal xstart As Double, ByVal yStart As Double, _
                         ByVal xend As Double, ByVal yEnd As Double, _
                         ByVal LineWidthMM As Double)
Dim Angle1 As Double, Angle2 As Double 'in radians
'1 - start, 2 - end

Angle1 = Arg(xstart - cxmm, yStart - cymm)
Angle2 = Arg(xend - cxmm, yEnd - cymm)
If Angle2 - Angle1 <= 0 Then Angle2 = Angle2 + 2 * Pi
If Angle2 - Angle1 <= 0 Then Angle2 = Angle2 + 2 * Pi
If Angle2 - Angle1 > 2 * Pi Then Angle2 = Angle2 - 2 * Pi
If Angle2 - Angle1 > 2 * Pi Then Angle2 = Angle2 - 2 * Pi

Dim Radius As Double
Radius = 0.5 * (XYLen(xstart - cxmm, yStart - cymm) + _
                XYLen(xend - cxmm, yEnd - cymm))

Dim n As Long, i As Long
n = -Int(-nmbSegments.Value * (Angle2 - Angle1) / 2 / Pi)
If n = 0 Then n = 1
For i = 0 To n - 1
  Dim a1 As Double, a2 As Double
  a1 = Angle1 + CDbl(i) / n * (Angle2 - Angle1)
  a2 = Angle1 + CDbl(i + 1) / n * (Angle2 - Angle1)
  DrawLine cxmm + Cos(a1) * Radius, cymm + Sin(a1) * Radius, _
           cxmm + Cos(a2) * Radius, cymm + Sin(a2) * Radius, _
           LineWidthMM
Next i
'athough calculation of range can be made more efficient _
 than calling DrawLine for every segment, I'm too lazy to _
 implement it
End Sub

Private Sub DrawRect(ByVal x As Double, ByVal y As Double, _
                      ByVal w As Double, ByVal h As Double, _
                      ByVal AngDeg As Double)
Dim si As Double, co As Double
si = Sin(AngDeg / 180 * Pi): co = Cos(AngDeg / 180 * Pi)
If bDraw Then
  Dim Pixels() As AlphaPixel
  Dim nPix As Long
  Dim cnt As vtVertex
  cnt.x = x * CurScale.PixPerMM + CurScale.OrgX
  cnt.y = y * -CurScale.PixPerMM + CurScale.OrgY
  cnt.Color = pCurColor
  
  DrawingEngine.pntRectangle cnt, _
                             w * CurScale.PixPerMM, _
                             h * CurScale.PixPerMM, _
                             -AngDeg, _
                             Pixels, nPix
  MainForm.dbDrawPixels Pixels, nPix, DrawTemp:=False, _
                        ForceDraw:=False, DrawMode:=pDrawMode, _
                        KillAA:=chkAntialiasing.Value <> vbChecked
Else
  RangePoint x - w * co + h * -si, y - w * si + h * co
  RangePoint x + w * co + h * -si, y + w * si + h * co
  RangePoint x - w * co - h * -si, y - w * si - h * co
  RangePoint x + w * co - h * -si, y + w * si - h * co
End If
End Sub

Private Sub AddPlgPoint(ByVal xmm As Double, ByVal ymm As Double)
If bDraw Then
  gPlg.AddVertex PixFromMMX(xmm), PixFromMMY(ymm)
Else
  RangePoint xmm, ymm
End If
End Sub

Private Sub RenderGPlg()
Dim Pixels() As AlphaPixel
Dim nPix As Long, nAlloc As Long
If gPlg.IsEmpty Then Exit Sub
gPlg.GetPoints Pixels, nPix, nAlloc
MainForm.dbDrawPixels Pixels, nPix, DrawTemp:=False, _
                      ForceDraw:=False, DrawMode:=pDrawMode
gPlg.Clear
End Sub

Private Function XYLen(ByVal x As Double, ByVal y As Double) As Double
XYLen = Sqr(x * x + y * y)
End Function

'extends rangevars to fit the specified point
Private Sub RangePoint(xmm As Double, ymm As Double)
If xmm < xminmm Or Not RangeFoo Then
  xminmm = xmm
End If
If ymm < yminmm Or Not RangeFoo Then
  yminmm = ymm
End If
If xmm > xmaxmm Or Not RangeFoo Then
  xmaxmm = xmm
End If
If ymm > ymaxmm Or Not RangeFoo Then
  ymaxmm = ymm
End If
RangeFoo = True
End Sub

'a redefined Val function
Private Function Val(ByRef St As String) As Double
'international hazard! This is introduced for debug only _
 and should be erased later
Static decsep As String
Static Foo As Boolean
If Not Foo Then decsep = Mid$(CStr(1.1), 2, 1): Foo = True
Val = CDbl(Replace(Replace(St, decsep, "kaka"), ".", decsep))
End Function

Public Sub DrawTracks()
Dim Wires As IXMLDOMNodeList
Set Wires = PCB.selectNodes(xqWires)
Dim iw As Long, isw As Long
For iw = 0 To Wires.Length - 1
  Dim Lr  As String
  Lr = GetAttribute(Wires(iw).selectSingleNode("LayerRef"), "name")
  If Lr = pSelectedLayer Then
    Dim SubWires As IXMLDOMNodeList
    Set SubWires = Wires(iw).selectNodes("Subwire")
    For isw = 0 To SubWires.Length - 1
      DrawElement SubWires(isw), 0
    Next isw
  End If
  ShowProgress iw / Wires.Length, DoDoEvents:=True
Next iw
ShowProgress 1.01
End Sub

Private Sub DrawComponents()
Dim Comps As IXMLDOMNodeList
Set Comps = PCB.selectNodes(xqComponents)
Dim i As Long
For i = 0 To Comps.Length - 1
  DrawComponent Comps(i)
Next i
End Sub

Private Sub DrawVias()
Dim Vias As IXMLDOMNodeList
Set Vias = PCB.selectNodes(xqVias)
Dim i As Long
For i = 0 To Vias.Length - 1
  Dim PSRef As String
  Dim x As Double, y As Double
  PSRef = GetAttribute(Vias(i).selectSingleNode("ViastackRef"), "name")
  GetXY Vias(i).selectSingleNode("Org"), x, y
  DrawPad PSRef, x, y, AngDeg:=0, BotSide:=False, bVia:=True
Next i
End Sub

Private Sub DrawContour()
Dim Shapes As IXMLDOMNodeList
Dim LineWidth As Double
Set Shapes = PCB.selectNodes(xqContour)
Dim i As Long
For i = 0 To Shapes.Length - 1
  LineWidth = Val(GetAttribute(Shapes(i), "lineWidth", "0.1"))
  DrawElement Shapes(i).childNodes(0), LineWidth
Next i

Set Shapes = PCB.selectNodes(xqContourVoids)
For i = 0 To Shapes.Length - 1
  LineWidth = Val(GetAttribute(Shapes(i), "lineWidth", "0.1"))
  DrawElement Shapes(i).childNodes(0), LineWidth
Next i
End Sub


Private Sub DrawMetallization()

If bDraw Then
  gPlg.Color = pCurColor
End If

Dim Cprs As IXMLDOMNodeList
Set Cprs = PCB.selectNodes(xqCoppers)
Dim i As Long
For i = 0 To Cprs.Length - 1
  DrawCopper Cprs(i)
Next i

End Sub

Private Sub DrawCopper(Cpr As IXMLDOMNode)
Dim LayerName As String
LayerName = GetAttribute(Cpr.selectSingleNode("LayerRef"), "name")
If LayerName <> pSelectedLayer Then Exit Sub
Dim LineWidth As Double
LineWidth = Val(GetAttribute(Cpr, "lineWidth"))
Dim FillType As String
FillType = GetAttribute(Cpr, "fillType", "Solid")
Dim bNonSolid As Boolean
bNonSolid = FillType <> "Solid"
Dim State As String
State = GetAttribute(Cpr, "state", "Unpoured")
Select Case State
  Case "Unpoured"
    MsgBox "Warning: unfilled polygon found."
    DrawFilledShape Cpr.selectSingleNode("Shape/*"), LineWidth, DontFill:=True
  Case "Poured", "Locked"
    Dim Islands As IXMLDOMNodeList
    Set Islands = Cpr.selectNodes("Islands/Island")
    If Islands.Length = 0 Then
      If State = "Locked" Then
        'there are no islands. Draw shape instead.
        DrawFilledShape Cpr.selectNodes("Shape/*"), LineWidth, DontFill:=bNonSolid
      Else
        'there are no islands... nothing to draw
      End If
    Else
      Dim II As Long
      For II = 0 To Islands.Length - 1
        
        Dim Plg As IXMLDOMNode
        Set Plg = Islands(II).selectSingleNode("Polygon")
        DrawFilledShape Plg, LineWidth, DontFill:=bNonSolid
        
        Dim Voids As IXMLDOMNodeList
        Set Voids = Islands(II).selectNodes("Voids/Polygon")
        Dim iV As Long
        For iV = 0 To Voids.Length - 1
          DrawFilledShape Voids(iV), LineWidth, DontFill:=bNonSolid
        Next iV
        
        If chkNoSpokes.Value = vbUnchecked Then
          Dim Spokes As IXMLDOMNodeList
          Set Spokes = Islands(II).selectNodes("ThermalSpoke")
          Dim iSP As Long
          For iSP = 0 To Spokes.Length - 1
            DrawElement Spokes(iSP), LineWidthMM:=0 '(linewidth is defined in spoke, drawelement will recognize it)
          Next iSP
        End If
      Next II
      If bNonSolid Then
        Dim Lines As IXMLDOMNodeList
        Set Lines = Cpr.selectNodes("Fill/Line")
        If Lines.Length = 0 Then MsgBox "Error. FillType=" + FillType + ", but <Fill> element not found..."
        Dim i As Long
        For i = 0 To Lines.Length - 1
          DrawElement Lines(i), LineWidth
        Next i
      End If
    End If
End Select
If bDraw Then RenderGPlg
End Sub

Private Sub DrawFilledShape(ByVal Shp As IXMLDOMNode, _
                            ByVal LineWidth As Double, _
                            Optional ByVal DontFill As Boolean)
Dim x As Double, y As Double
Dim x1 As Double, y1 As Double
Dim x2 As Double, y2 As Double
'!!!!!!!!
'DontFill = True
Select Case Shp.nodeName
  Case "FilledCircle"
    Dim Dia As Double
    Dia = Val(GetAttribute(Shp, "diameter"))
    GetXY Shp.selectSingleNode("Center"), x, y
    If DontFill Then
      DrawArc x, y, x + Dia / 2, y, x + Dia / 2, y, LineWidth
    Else
      DrawLine x, y, x * 1.0000000000001, y, Dia + LineWidth
    End If
  Case "FilledRect"
    GetXY Shp.childNodes(0), x1, y1
    GetXY Shp.childNodes(1), x2, y2
    If Not DontFill Then
      DrawRect (x1 + x2) * 0.5, (y1 + y2) * 0.5, Abs(x2 - x1), Abs(y2 - y1), AngDeg:=0
    End If
    DrawLine x1, y1, x1, y2, LineWidth
    DrawLine x1, y2, x2, y2, LineWidth
    DrawLine x2, y2, x2, y1, LineWidth
    DrawLine x2, y1, x1, y1, LineWidth
  Case "Polygon", "Teardrop"
    Dim Verts As IXMLDOMNodeList
    If Not DontFill Then gPlg.NewSubpolygon
    Set Verts = Shp.childNodes
    Dim i As Long
    For i = 0 To Verts.Length - 1
      GetXY Verts(i), x1, y1
      Dim inv As Long
      inv = i + 1
      If inv = Verts.Length Then inv = 0
      GetXY Verts(inv), x2, y2
      DrawLine x1, y1, x2, y2, LineWidth
      If Not DontFill Then AddPlgPoint x1, y1
    Next i
End Select
End Sub

'considers range already has been calculated
Private Sub PrepareForDraw(ByVal DoResize As Boolean)
Const MARGIN_PX = 2
CurScale.PixPerMM = nmbResolution.Value / 25.4
Dim PicW As Long, PicH As Long
If Not RangeFoo Then Err.Raise 12345, "PrepareForDraw", "Range variables are uninitialized, picture size cannot be determined!"
PicW = (xmaxmm - xminmm) * CurScale.PixPerMM + MARGIN_PX * 2
PicH = (ymaxmm - yminmm) * CurScale.PixPerMM + MARGIN_PX * 2
CurScale.OrgX = MARGIN_PX - xminmm * CurScale.PixPerMM
CurScale.OrgY = MARGIN_PX + ymaxmm * CurScale.PixPerMM
If DoResize Then
  MainForm.ChangeActiveColor 2, vbWhite
  MainForm.Resize PicW, PicH, StoreUndo:=True, StretchMode:=SMClear
  MainForm.StartPixelAction
End If
End Sub

Private Sub Form_Load()
loadSettings
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If chkRemember.Value = vbChecked Then SaveSettings
End Sub

Private Sub lstLayers_Click()
If lstLayers.ListIndex < 0 Then Exit Sub
pSelectedLayer = lstLayers.List(lstLayers.ListIndex)
End Sub

Private Sub DrawEverything(Optional ByVal Sizing As Boolean = False)
Dim All As Boolean
All = Sizing And optSzWhole.Value

pDrawMode = dmMinimum

If Not Sizing Then ShowStatus "Drawing Wires..."
pCurColor = clrWires.Color
If chkElWires.Value = vbChecked Or All Then DrawTracks

If Not Sizing Then ShowStatus "Drawing Metallization..."
pCurColor = clrMet.Color
If chkElMet.Value = vbChecked Or All Then DrawMetallization

If Not Sizing Then ShowStatus "Drawing Contour..."
pCurColor = clrContour.Color
If chkElContour.Value = vbChecked Or All Then DrawContour

pDrawMode = dmNormal

If Not Sizing Then ShowStatus "Drawing Vias..."
pCurColor = clrVias.Color
If chkElVias.Value = vbChecked Or All Then DrawVias

If Not Sizing Then ShowStatus "Drawing Pads..."
pCurColor = clrPads.Color
If chkElPads.Value = vbChecked Or All Then DrawComponents

End Sub


Public Sub SaveSettings()
If Not chkRemember.Value = vbChecked Then Exit Sub
dbSaveSettingEx "TopoRr", "bDraw_wires", chkElWires.Value = vbChecked
dbSaveSettingEx "TopoRr", "bDraw_pads", chkElPads.Value = vbChecked
dbSaveSettingEx "TopoRr", "bDraw_vias", chkElVias.Value = vbChecked
dbSaveSettingEx "TopoRr", "bDraw_met", chkElMet.Value = vbChecked
dbSaveSettingEx "TopoRr", "bDraw_contour", chkElContour.Value = vbChecked
dbSaveSettingEx "TopoRr", "bDraw_holes", chkElHoles.Value = vbChecked

dbSaveSettingEx "TopoRr", "Color_wires", clrWires.Color, HexNumber:=True
dbSaveSettingEx "TopoRr", "Color_pads", clrPads.Color, HexNumber:=True
dbSaveSettingEx "TopoRr", "Color_vias", clrVias.Color, HexNumber:=True
dbSaveSettingEx "TopoRr", "Color_met", clrMet.Color, HexNumber:=True
dbSaveSettingEx "TopoRr", "Color_contour", clrContour.Color, HexNumber:=True
dbSaveSettingEx "TopoRr", "Color_holes", clrHoles.Color, HexNumber:=True

dbSaveSettingEx "TopoRr", "resolution", nmbResolution.Value
dbSaveSettingEx "TopoRr", "antialias", chkAntialiasing.Value = vbChecked
dbSaveSettingEx "TopoRr", "nsegments", nmbSegments.Value

dbSaveSettingEx "TopoRr", "size_fitall", optSzNeeded.Value

End Sub

Public Sub loadSettings()
bSuppress = True
chkRemember.Value = Abs(dbGetSettingEx("TopoRr", "SaveSettings", vbBoolean, False))
bSuppress = False
  
chkElWires.Value = Abs(dbGetSettingEx("TopoRr", "bDraw_wires", vbBoolean, True))
chkElPads.Value = Abs(dbGetSettingEx("TopoRr", "bDraw_pads", vbBoolean, True))
chkElVias.Value = Abs(dbGetSettingEx("TopoRr", "bDraw_vias", vbBoolean, True))
chkElMet.Value = Abs(dbGetSettingEx("TopoRr", "bDraw_met", vbBoolean, True))
chkElContour.Value = Abs(dbGetSettingEx("TopoRr", "bDraw_contour", vbBoolean, True))
chkElHoles.Value = Abs(dbGetSettingEx("TopoRr", "bDraw_holes", vbBoolean, True))

clrWires.Color = dbGetSettingEx("TopoRr", "Color_wires", vbLong, 0)
clrPads.Color = dbGetSettingEx("TopoRr", "Color_pads", vbLong, 0)
clrVias.Color = dbGetSettingEx("TopoRr", "Color_vias", vbLong, 0)
clrMet.Color = dbGetSettingEx("TopoRr", "Color_met", vbLong, 0)
clrContour.Color = dbGetSettingEx("TopoRr", "Color_contour", vbLong, 0)
clrHoles.Color = dbGetSettingEx("TopoRr", "Color_holes", vbLong, vbWhite)

nmbResolution.Value = dbGetSettingEx("TopoRr", "resolution", vbDouble, 600)
chkAntialiasing.Value = Abs(dbGetSettingEx("TopoRr", "antialias", vbBoolean, False))
nmbSegments.Value = dbGetSettingEx("TopoRr", "nsegments", vbInteger, 120)

optSzWhole.Value = Not dbGetSettingEx("TopoRr", "size_fitall", vbBoolean, False)

End Sub
