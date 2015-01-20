Attribute VB_Name = "TypesDefinitions"
Option Explicit

'**************************   T Y P E S   **********************************************


Type Dims
    w As Long
    h As Long
End Type

'Type MaskInfo
'    w As Long
'    h As Long
'    cx As Long
'    cy As Long
'End Type
'
Type dbData
    d() As Long
End Type

Type dbPixel
    x As Long
    y As Long
    c As Long
End Type

Type Rectangle
    x1 As Long
    y1 As Long
    x2 As Long
    y2 As Long
End Type

Type POINTAPI
        x As Long
        y As Long
End Type

Type dbUndoData
    EntryType As dbUndoTypes
    Pixels() As dbPixel
    PixelsUB As Long
    PixelsVUB As Long
    d() As Long
    Region As Rectangle
    Org As POINTAPI
End Type

Type dbUndoStorage
    Index As Long
    FirstIndex As Long
    uData() As dbUndoData
End Type

Type RGBQuadLong
        rgbBlue As Long
        rgbGreen As Long
        rgbRed As Long
        rgbReserved As Long
End Type

Type RGBTriCurr
        rgbBlue As Currency
        rgbGreen As Currency
        rgbRed As Currency
End Type

Type RGBTriSng
    rgbRed As Single
    rgbGreen As Single
    rgbBlue As Single
End Type

Type RGBTriInt
    rgbRed As Integer
    rgbGreen As Integer
    rgbBlue As Integer
End Type

Type RGBTriLong
    rgbRed As Long
    rgbGreen As Long
    rgbBlue As Long
End Type

Type FilterMask
    Center As POINTAPI
    CenterFilled As Boolean
    Mask() As RGBTriLong
End Type

Type IconMask
    d() As Boolean
End Type

Type PointByte
    x As Byte
    y As Byte
End Type

Type Variable
    Name As String
    Value As Double
End Type

Type SMP
    Vars() As Variable
    Code As Variant
    Source As String
End Type

Type HelixSettings
    Numb As Single
    RMode As Integer
    RFixed As Single
    RK As Single
End Type

Type kShortcut
    Key As Long
    Act As Long 'dbCommands
End Type

Type ArrShortcuts
    Keys() As kShortcut
End Type

Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Type DRAWTEXTPARAMS
    cbSize As Long
    iTabLength As Long
    iLeftMargin As Long
    iRightMargin As Long
    uiLengthDrawn As Long
End Type

Type dbGraph
    InterpolationMode As vtInterpolMode
    NeedsInterpolation As Boolean
    Points() As PointByte
    Table() As Byte
End Type

'Public Type ErrInfo
'    Number As Long
'    Source As String
'    Description As String
'End Type
'

Public Type LineSettings
    GeoMode As Long
    Weight As Double
    RelWeight1 As Double
    RelWeight2 As Double
    AntiAliasing As Double
End Type



Public Type typAutoscrollSgs
  ' these vals specify how close to border of visible area
  '  the pointer should be to start autoscrolling.
  '  Separately for each side.
  GapLef As Long
  GapTop As Long
  GapRig As Long
  GapBot As Long
End Type

Public Type typScrollSettings
  'Dim MouseGlued As Boolean
  MouseGlued As Boolean
  'Dim AutoScroll_Field_Size As Long '= 100
  ASS As typAutoscrollSgs
  ASS_tmp As typAutoscrollSgs 'temporary autoscroll settings, filled on mousedown to reduce gaps if their sum is greater than visible area
  ASS_tmp2 As typAutoscrollSgs 'temporary autoscroll settings, used for smart autoscrolling
  ASS_pen As typAutoscrollSgs
  'Dim Jestkost As Single, EnL As Single
  DS_Jestkost As Single
  DS_EnL As Single
  'Dim DynamicScrolling As Boolean
  DS_Enabled As Boolean 'Do not change directly - use ChangeDynScrolling()
  'Dim DynScrollingIfAR As Long
  'DS_EnableIfAR As Long 'seems to be actually unused
  'Dim CancelWheelScroll As Boolean
  CancelWheelScroll As Boolean 'not to be saved
  'Dim DontScroll As Boolean
  DontScroll As Boolean 'not to be saved
  
  NaviAbsoluteMode As Boolean 'True tells navi mode not to use SetCursorPos

End Type

Public Type BrushPoint
  x1 As Double
  y1 As Double
  x2 As Double
  y2 As Double
  color As Long
End Type
