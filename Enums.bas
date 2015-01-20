Attribute VB_Name = "Enums"
Option Explicit


Public Enum AWF
    AW_HOR_POSITIVE = &H1&
    AW_HOR_NEGATIVE = &H2&
    AW_VER_POSITIVE = &H4&
    AW_VER_NEGATIVE = &H8&
    AW_CENTER = &H10&
    AW_HIDE = &H10000
    AW_ACTIVATE = &H20000
    AW_SLIDE = &H40000
    AW_BLEND = &H80000
End Enum

Public Enum WH
    WH_GETMESSAGE = 3
    WH_HARDWARE = 8
    WH_JOURNALPLAYBACK = 1
    WH_JOURNALRECORD = 0
    WH_KEYBOARD = 2
    WH_MAX = 11
    WH_MIN = (-1)
    WH_MOUSE = 7
    WH_MSGFILTER = (-1)
    WH_SHELL = 10
    WH_SYSMSGFILTER = 6
End Enum

Public Enum PM
    PM_NOREMOVE = &H0
    PM_REMOVE = &H1
End Enum

Public Enum WM
    'App messages
    WM_ACTIVATEAPP = &H1C
    
    'Mouse messages
    WM_Wheel = 522
    WM_MOUSEFIRST = &H200
    WM_MOUSELAST = &H209
    WM_MOUSEMOVE = &H200
    WM_LBUTTONDBLCLK = &H203
    WM_LBUTTONDOWN = &H201
    WM_LBUTTONUP = &H202
    WM_MBUTTONDBLCLK = &H209
    WM_MBUTTONDOWN = &H207
    WM_MBUTTONUP = &H208
    WM_RBUTTONDBLCLK = &H206
    WM_RBUTTONDOWN = &H204
    WM_RBUTTONUP = &H205
    
    'Tablet messages
    WM_TABLET_QUERYSYSTEMGESTURESTATUS = 716
    
    'Active accessibility
    WM_GETOBJECT = &H3D&
    
    'not sorted yet
' Window Messages
    WM_NULL = &H0
    WM_CREATE = &H1
    WM_DESTROY = &H2
    WM_MOVE = &H3
    WM_SIZE = &H5

    WM_ACTIVATE = &H6
'
'  WM_ACTIVATE state values

    WM_SETFOCUS = &H7
    WM_KILLFOCUS = &H8
    WM_ENABLE = &HA
    WM_SETREDRAW = &HB
    WM_SETTEXT = &HC
    WM_GETTEXT = &HD
    WM_GETTEXTLENGTH = &HE
    WM_PAINT = &HF
    WM_CLOSE = &H10
    WM_QUERYENDSESSION = &H11
    WM_QUIT = &H12
    WM_QUERYOPEN = &H13
    WM_ERASEBKGND = &H14
    WM_SYSCOLORCHANGE = &H15
    WM_ENDSESSION = &H16
    WM_SHOWWINDOW = &H18
    WM_WININICHANGE = &H1A
    WM_DEVMODECHANGE = &H1B
    WM_FONTCHANGE = &H1D
    WM_TIMECHANGE = &H1E
    WM_CANCELMODE = &H1F
    WM_SETCURSOR = &H20
    WM_MOUSEACTIVATE = &H21
    WM_CHILDACTIVATE = &H22
    WM_QUEUESYNC = &H23

    WM_GETMINMAXINFO = &H24

    WM_PAINTICON = &H26
    WM_ICONERASEBKGND = &H27
    WM_NEXTDLGCTL = &H28
    WM_SPOOLERSTATUS = &H2A
    WM_DRAWITEM = &H2B
    WM_MEASUREITEM = &H2C
    WM_DELETEITEM = &H2D
    WM_VKEYTOITEM = &H2E
    WM_CHARTOITEM = &H2F
    WM_SETFONT = &H30
    WM_GETFONT = &H31
    WM_SETHOTKEY = &H32
    WM_GETHOTKEY = &H33
    WM_QUERYDRAGICON = &H37
    WM_COMPAREITEM = &H39
    WM_COMPACTING = &H41


    WM_WINDOWPOSCHANGING = &H46
    WM_WINDOWPOSCHANGED = &H47

    WM_POWER = &H48

    WM_COPYDATA = &H4A
    WM_CANCELJOURNAL = &H4B


    WM_NCCREATE = &H81
    WM_NCDESTROY = &H82
    WM_NCCALCSIZE = &H83
    WM_NCHITTEST = &H84
    WM_NCPAINT = &H85
    WM_NCACTIVATE = &H86
    WM_GETDLGCODE = &H87
    WM_NCMOUSEMOVE = &HA0
    WM_NCLBUTTONDOWN = &HA1
    WM_NCLBUTTONUP = &HA2
    WM_NCLBUTTONDBLCLK = &HA3
    WM_NCRBUTTONDOWN = &HA4
    WM_NCRBUTTONUP = &HA5
    WM_NCRBUTTONDBLCLK = &HA6
    WM_NCMBUTTONDOWN = &HA7
    WM_NCMBUTTONUP = &HA8
    WM_NCMBUTTONDBLCLK = &HA9

    WM_KEYFIRST = &H100
    WM_KEYDOWN = &H100
    WM_KEYUP = &H101
    WM_CHAR = &H102
    WM_DEADCHAR = &H103
    WM_SYSKEYDOWN = &H104
    WM_SYSKEYUP = &H105
    WM_SYSCHAR = &H106
    WM_SYSDEADCHAR = &H107
    WM_KEYLAST = &H108
    WM_INITDIALOG = &H110
    WM_COMMAND = &H111
    WM_SYSCOMMAND = &H112
    WM_TIMER = &H113
    WM_HSCROLL = &H114
    WM_VSCROLL = &H115
    WM_INITMENU = &H116
    WM_INITMENUPOPUP = &H117
    WM_MENUSELECT = &H11F
    WM_MENUCHAR = &H120
    WM_ENTERIDLE = &H121

    WM_CTLCOLORMSGBOX = &H132
    WM_CTLCOLOREDIT = &H133
    WM_CTLCOLORLISTBOX = &H134
    WM_CTLCOLORBTN = &H135
    WM_CTLCOLORDLG = &H136
    WM_CTLCOLORSCROLLBAR = &H137
    WM_CTLCOLORSTATIC = &H138

    WM_PARENTNOTIFY = &H210
    WM_ENTERMENULOOP = &H211
    WM_EXITMENULOOP = &H212
    WM_MDICREATE = &H220
    WM_MDIDESTROY = &H221
    WM_MDIACTIVATE = &H222
    WM_MDIRESTORE = &H223
    WM_MDINEXT = &H224
    WM_MDIMAXIMIZE = &H225
    WM_MDITILE = &H226
    WM_MDICASCADE = &H227
    WM_MDIICONARRANGE = &H228
    WM_MDIGETACTIVE = &H229
    WM_MDISETMENU = &H230
    WM_DROPFILES = &H233
    WM_MDIREFRESHMENU = &H234


    WM_CUT = &H300
    WM_COPY = &H301
    WM_PASTE = &H302
    WM_CLEAR = &H303
    WM_UNDO = &H304
    WM_RENDERFORMAT = &H305
    WM_RENDERALLFORMATS = &H306
    WM_DESTROYCLIPBOARD = &H307
    WM_DRAWCLIPBOARD = &H308
    WM_PAINTCLIPBOARD = &H309
    WM_VSCROLLCLIPBOARD = &H30A
    WM_SIZECLIPBOARD = &H30B
    WM_ASKCBFORMATNAME = &H30C
    WM_CHANGECBCHAIN = &H30D
    WM_HSCROLLCLIPBOARD = &H30E
    WM_QUERYNEWPALETTE = &H30F
    WM_PALETTEISCHANGING = &H310
    WM_PALETTECHANGED = &H311
    WM_HOTKEY = &H312

    WM_PENWINFIRST = &H380
    WM_PENWINLAST = &H38F

' NOTE: All Message Numbers below 0x0400 are RESERVED.

' Private Window Messages Start Here:
    WM_USER = &H400
End Enum

Public Enum eTabletWMResponse
'  #define TABLET_DISABLE_PRESSANDHOLD        0x00000001
'  #define TABLET_DISABLE_PENTAPFEEDBACK      0x00000008
'  #define TABLET_DISABLE_PENBARRELFEEDBACK   0x00000010
'  #define TABLET_DISABLE_TOUCHUIFORCEON      0x00000100
'  #define TABLET_DISABLE_TOUCHUIFORCEOFF     0x00000200
'  #define TABLET_DISABLE_TOUCHSWITCH         0x00008000
'  #define TABLET_DISABLE_FLICKS              0x00010000
'  #define TABLET_ENABLE_FLICKSONCONTEXT      0x00020000
'  #define TABLET_ENABLE_FLICKLEARNINGMODE    0x00040000
'  #define TABLET_DISABLE_SMOOTHSCROLLING     0x00080000
'  #define TABLET_DISABLE_FLICKFALLBACKKEYS   0x00100000
'  #define TABLET_ENABLE_MULTITOUCHDATA       0x01000000
  TABLET_DISABLE_PRESSANDHOLD = &H1&
  TABLET_DISABLE_PENTAPFEEDBACK = &H8&
  TABLET_DISABLE_PENBARRELFEEDBACK = &H10&
  TABLET_DISABLE_TOUCHUIFORCEON = &H100&
  TABLET_DISABLE_TOUCHUIFORCEOFF = &H200&
  TABLET_DISABLE_TOUCHSWITCH = &H8000&
  TABLET_DISABLE_FLICKS = &H10000
  TABLET_ENABLE_FLICKSONCONTEXT = &H20000
  TABLET_ENABLE_FLICKLEARNINGMODE = &H40000
  TABLET_DISABLE_SMOOTHSCROLLING = &H80000
  TABLET_DISABLE_FLICKFALLBACKKEYS = &H100000
  TABLET_ENABLE_MULTITOUCHDATA = &H1000000
End Enum


Public Enum SWP_Flags
    SWP_DRAWFRAME = &H20
    SWP_FRAMECHANGED = &H20        '  The frame changed: send WM_NCCALCSIZE
    SWP_HIDEWINDOW = &H80
    SWP_NOACTIVATE = &H10
    SWP_NOCOPYBITS = &H100
    SWP_NOMOVE = &H2
    SWP_NOOWNERZORDER = &H200      '  Don't do owner Z ordering
    SWP_NOREDRAW = &H8
    SWP_NOREPOSITION = &H200
    SWP_NOSIZE = &H1
    SWP_NOZORDER = &H4
    SWP_SHOWWINDOW = &H40
End Enum

Public Enum SWP_InsAfter
    HWND_BOTTOM = 1
    HWND_NOTOPMOST = -2
    HWND_TOP = 0
    HWND_TOPMOST = -1
End Enum

Public Enum Metric
    SM_CMETRICS = 44
    SM_CMOUSEBUTTONS = 43
    SM_CXBORDER = 5
    SM_CXCURSOR = 13
    SM_CXDLGFRAME = 7
    SM_CXDOUBLECLK = 36
    SM_CXFIXEDFRAME = SM_CXDLGFRAME
    SM_CXFRAME = 32
    SM_CXFULLSCREEN = 16
    SM_CXHSCROLL = 21
    SM_CXHTHUMB = 10
    SM_CXICON = 11
    SM_CXICONSPACING = 38
    SM_CXMIN = 28
    SM_CXMINTRACK = 34
    SM_CXSCREEN = 0
    SM_CXSIZE = 30
    SM_CXSIZEFRAME = SM_CXFRAME
    SM_CXVSCROLL = 2
    SM_CYBORDER = 6
    SM_CYCAPTION = 4
    SM_CYCURSOR = 14
    SM_CYDLGFRAME = 8
    SM_CYDOUBLECLK = 37
    SM_CYFIXEDFRAME = SM_CYDLGFRAME
    SM_CYFRAME = 33
    SM_CYFULLSCREEN = 17
    SM_CYHSCROLL = 3
    SM_CYICON = 12
    SM_CYICONSPACING = 39
    SM_CYKANJIWINDOW = 18
    SM_CYMENU = 15
    SM_CYMIN = 29
    SM_CYMINTRACK = 35
    SM_CYSCREEN = 1
    SM_CYSIZE = 31
    SM_CYSIZEFRAME = SM_CYFRAME
    SM_CYVSCROLL = 20
    SM_CYVTHUMB = 9
    SM_DBCSENABLED = 42
    SM_DEBUG = 22
    SM_MENUDROPALIGNMENT = 40
    SM_MOUSEPRESENT = 19
    SM_PENWINDOWS = 41
    SM_RESERVED1 = 24
    SM_RESERVED2 = 25
    SM_RESERVED3 = 26
    SM_RESERVED4 = 27
    SM_SWAPBUTTON = 23
End Enum

Public Enum MEFlags
    dbMOUSEEVENT_LEFTDOWN = &H2
    dbMOUSEEVENT_LEFTUP = &H4
    dbMOUSEEVENTF_RIGHTDOWN = &H8
    dbMOUSEEVENTF_RIGHTUP = &H10
    dbMOUSEEVENTF_MIDDLEDOWN = &H20
    dbMOUSEEVENTF_MIDDLEUP = &H40
    dbMOUSE_MOVED = &H1
    dbmouse_eventC = &H2
End Enum

'Public Enum PenEnum
'    PS_INSIDEFRAME = 6
'End Enum
'

Public Enum CapIndex
    HORZSIZE = 4           '  Horizontal size in millimeters
    VERTSIZE = 6           '  Vertical size in millimeters
    VERTRES = 10           '  Vertical width in pixels
    HORZRES = 8            '  Horizontal width in pixels
    LOGPIXELSX = 88        '  Logical pixels/inch in X
    LOGPIXELSY = 90        '  Logical pixels/inch in Y
    BITSPIXEL = 12         '  Number of bits per pixel
End Enum

Public Enum TextDrawMode
    DT_BOTTOM = &H8
    DT_CALCRECT = &H400
    DT_CENTER = &H1
    DT_CHARSTREAM = 4
    DT_DISPFILE = 6
    DT_EXPANDTABS = &H40
    DT_EXTERNALLEADING = &H200
    DT_INTERNAL = &H1000
    DT_LEFT = &H0
    DT_METAFILE = 5
    DT_NOCLIP = &H100
    DT_NOPREFIX = &H800
    DT_RIGHT = &H2
    DT_SINGLELINE = &H20
    DT_TABSTOP = &H80
    DT_TOP = &H0
    DT_VCENTER = &H4
    DT_WORDBREAK = &H10
End Enum

Public Enum ShowWindowCommand
    SW_SHOWNA = 8
    SW_HIDE = 0
    SW_SHOWNOACTIVATE = 4
End Enum


Public Enum KeyCodes
    dbDown = 98
    dbUP = 104
    dbLeft = 100
    dbRight = 102
    dbShift = 16
    dbCtrl = 17
    dbAlt = 18
End Enum

Public Enum dbShiftConstants
    dbStateShift = &H100
    dbStateCtrl = &H200
    dbStateAlt = &H400
End Enum

Public Enum dbCircleFlags
    'draw mode
    dbCircleDiameter = 0
    dbCircleRadius = 1
    dbCircleInside = 2
    dbCircleOutSide = 3
    'style
    dbCirclePutCenter = &H10&
    dbCirclePutFocuses = &H20&
    dbCirclePunktir = &H40&
    dbCircleHQ = &H80&
End Enum

Public Enum eLineGeoMode
    'draw mode
    dbLineSimple = 0
    dbLineDouble = 1
    dbLinePerp = 2
    dbLineParallel = 3
End Enum

Public Enum dbUndoTypes
    dbUndoInvalid = -1
    dbUndoFull = 0
    dbUndoPixels = 1
    dbUndoFragment = 2
End Enum

Public Enum STT_Messages
    eSTT_READY = 1204
    eSTT_Copying = 1200
    eSTT_BUD = 1202
    eSTT_Processing = 1203
    eSTT_Loading = 1209
    eSTT_Working = 1212
    eSTT_Cancelled = 1217
    eSTT_Resizing = 1236
    eSTT_Displaying = 1218
    eSTT_Error = 1227
End Enum

Public Enum DlgFilter
'dbSaveSMB = 1
'dbSaveBMP = 2
'dbLoad = 3
'dbLoadOld = 4
'dbWinPicture = 5
'dbSaveIco = 6
'dbSaveCur = 7
dbFLoadPal = 8
dbFSavePal = 9
dbBSave = 10
dbBLoad = 11
dbGLoad = 12
dbGSave = 13
dbPLoad = 14
dbPSave = 15
dbKeysLoad = 16
dbKeysSave = 17
dbTextSave = 18
dbPToolLoad = 19
dbPToolSave = 20
'dbPNGSave = 21
'dbJPSave = 22
dbMatrixLoad = 23
dbMatrixSave = 24
dbAsciiLoad = 25
dbFSTLoad = 26
End Enum

Public Enum PDM
    dbDrawToBuffer = 0
    dbDrawDirect = 1
End Enum

Public Enum dbMouseEvent
    dbEvMouseDown = 0
    dbEvMouseMove = 1
    dbEvMouseUp = 2
End Enum

Public Enum dbKeyEvent
    KeyDown = 0
    KeyPress = 1
    KeyUp = 2
End Enum

Public Enum vtInterpolMode
    dbIMLinear = 0
    dbIMPolynomial = 1
End Enum
'**************************///E N U M S///******************************************

