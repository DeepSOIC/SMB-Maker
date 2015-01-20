Attribute VB_Name = "WinAPI"
Option Explicit

Public Declare Function APIGetTempPath Lib "kernel32" _
            Alias "GetTempPathA" (ByVal nBufferLength As Long, _
                                  ByVal lpBuffer As String) As Long

Public Declare Function GetTickCount Lib "kernel32" () As Long

Public Declare Function SetDIBitsToDevice Lib "gdi32" _
            (ByVal hDC As Long, _
            ByVal x As Long, ByVal y As Long, _
            ByVal dx As Long, ByVal dy As Long, _
            ByVal SrcX As Long, ByVal SrcY As Long, _
            ByVal Scan As Long, ByVal NumScans As Long, _
            ByRef Bits As Any, BitsInfo As BITMAPINFO, _
            ByVal wUsage As Long) As Long

Public Declare Function StretchDIBits Lib "gdi32" _
            (ByVal hDC As Long, _
            ByVal x As Long, ByVal y As Long, _
            ByVal dx As Long, ByVal dy As Long, _
            ByVal SrcX As Long, ByVal SrcY As Long, _
            ByVal wSrcWidth As Long, ByVal wSrcHeight As Long, _
            lpBits As Any, lpBitsInfo As BITMAPINFO, _
            ByVal wUsage As Long, _
            ByVal dwRop As Long) As Long
Type RGBQUAD
    rgbBlue As Byte
    rgbGreen As Byte
    rgbRed As Byte
    rgbReserved As Byte
End Type
Type BITMAPINFOHEADER
    biSize As Long
    biWidth As Long
    biHeight As Long
    biPlanes As Integer
    biBitCount As Integer
    biCompression As Long
    biSizeImage As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed As Long
    biClrImportant As Long
End Type
Type BITMAPINFO
        bmiHeader As BITMAPINFOHEADER
        bmiColors As RGBQUAD
End Type

Public Declare Function CreateDIBSection _
       Lib "gdi32" (ByVal hDC As Long, _
                    pBitmapInfo As BITMAPINFO, _
                    ByVal iUsage As Long, _
                    ByVal lplpData As Long, _
                    ByVal hSection As Long, _
                    ByVal dwOffset As Long) As Long

Public Declare Function GdiFlush Lib "gdi32" () As Long

Public Declare Function APIGetDiBits Lib "gdi32" Alias "GetDIBits" _
            (ByVal aHDC As Long, _
            ByVal hBitmap As Long, _
            ByVal nStartScan As Long, ByVal nNumScans As Long, _
            ByRef lpBits As Any, lpBI As BITMAPINFO, _
            ByVal wUsage As Long) As Long

Public Declare Function GetSystemMetrics Lib "user32" _
            (ByVal nIndex As Metric) As Long

Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer

Public Declare Function GetMessage Lib "user32" _
            Alias "GetMessageA" (lpMsg As Msg, _
                                ByVal hWnd As Long, _
                                ByVal wMsgFilterMin As WM, _
                                ByVal wMsgFilterMax As WM) As Long
Type Msg
    hWnd As Long
    Message As WM
    wParam As Long
    lParam As Long
    pTime As Long
    pt As POINTAPI
End Type

Public Declare Function PeekMessage Lib "user32" _
            Alias "PeekMessageA" (lpMsg As Msg, _
                                  ByVal hWnd As Long, _
                                  ByVal wMsgFilterMin As WM, _
                                  ByVal wMsgFilterMax As WM, _
                                  ByVal wRemoveMsg As Long) As Long

Public Declare Function GetMessageExtraInfo Lib "user32" () As Long

Public Declare Function WaitMessage Lib "user32" () As Long

Public Const WHEEL_DELTA = 120&

Public Enum SB
  SB_HORZ = 0
  SB_LINELEFT = 0
  SB_LINEUP = 0
  SB_LINEDOWN = 1
  SB_LINERIGHT = 1
  SB_VERT = 1
  SB_CTL = 2
  SB_PAGELEFT = 2
  SB_PAGEUP = 2
  SB_BOTH = 3
  SB_PAGEDOWN = 3
  SB_PAGERIGHT = 3
  SB_THUMBPOSITION = 4
  SB_THUMBTRACK = 5
  SB_LEFT = 6
  SB_TOP = 6
  SB_BOTTOM = 7
  SB_RIGHT = 7
  SB_ENDSCROLL = 8
End Enum

Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

Public Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long

Public Declare Sub Mouse_Event Lib "user32" Alias "mouse_event" _
            (ByVal dwFlags As MEFlags, _
            ByVal dx As Long, ByVal dy As Long, _
            ByVal cButtons As Long, _
            Optional ByVal dwExtraInfo As Long = 1)
            
Public Declare Function SetCapture Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function GetCapture Lib "user32" () As Long



Public Declare Function SetDIBits Lib "gdi32" _
            (ByVal hDC As Long, _
            ByVal hBitmap As Long, _
            ByVal nStartScan As Long, ByVal nNumScans As Long, _
            lpBits As Any, _
            lpBI As BITMAPINFO, _
            ByVal wUsage As Long) As Long

Public Declare Function GetAsyncKeyState Lib "user32" _
            (ByVal vKey As Long) As Integer

'Public Declare Function Rectangle Lib "gdi32" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
'Public Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
'Public Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As PenEnum, ByVal nWidth As Long, ByVal crColor As Long) As Long
'Public Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
'Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
'DC functions

Public Declare Function ReleaseDC Lib "user32" _
            (ByVal hWnd As Long, _
            ByVal hDC As Long) As Long

Public Declare Function GetWindowDC Lib "user32" _
            (ByVal hWnd As Long) As Long

Public Declare Function GetDC _
            Lib "user32" _
            (ByVal hWnd As Long) As Long

Public Declare Function CreateCompatibleDC Lib "gdi32" _
            (ByVal hDC As Long) As Long

Public Declare Function DeleteDC Lib "gdi32" _
            (ByVal hDC As Long) As Long

Public Declare Function GetDeviceCaps Lib "gdi32" _
            (ByVal hDC As Long, _
            ByVal nIndex As CapIndex) As Long


'Window functions

Public Declare Function GetDesktopWindow Lib "user32" () As Long

Public Declare Function GetWindowRect Lib "user32" _
            (ByVal hWnd As Long, _
            lpRect As RECT) As Long
Public Declare Function GetClientRect Lib "user32" _
            (ByVal hWnd As Long, _
            lpRect As RECT) As Long


Public Declare Function WindowFromPoint Lib "user32" _
            (ByVal xPoint As Long, ByVal yPoint As Long) As Long

Public Declare Function GetForegroundWindow Lib "user32" () As Long

Public Declare Function EnableWindow Lib "user32" _
            (ByVal hWnd As Long, _
            ByVal fEnable As Long) As Long

Public Declare Function ShowWindow Lib "user32" _
            (ByVal hWnd As Long, _
            ByVal nCmdShow As ShowWindowCommand) As Long

Public Declare Function SetWindowPos Lib "user32" ( _
            ByVal hWnd As Long, _
            ByVal hWndInsertAfter As SWP_InsAfter, _
            ByVal x As Long, ByVal y As Long, _
            ByVal CX As Long, ByVal CY As Long, _
            ByVal wFlags As SWP_Flags) As Long

Public Declare Function IsChild Lib "user32" _
            (ByVal hWndParent As Long, _
            ByVal hWnd As Long) As Long

Public Declare Function AnimateWindow Lib "user32" _
            (ByVal nWnd As Long, _
            ByVal msTime As Long, _
            ByVal dwFlags As AWF) As Long

Public Declare Function GetActiveWindow Lib "user32" () As Long

Public Declare Function SetWindowLong Lib "user32" _
            Alias "SetWindowLongA" _
            (ByVal hWnd As Long, _
            ByVal nIndex As GWL_constants, _
            ByVal dwNewLong As Long) As Long
Public Declare Function GetWindowLong Lib "user32" _
            Alias "GetWindowLongA" _
            (ByVal hWnd As Long, ByVal nIndex As GWL_constants) As Long
Public Enum GWL_constants
  GWL_EXSTYLE = (-20)
  GWL_HINSTANCE = (-6)
  GWL_HWNDPARENT = (-8)
  GWL_ID = (-12)
  GWL_STYLE = (-16)
  GWL_USERDATA = (-21)
  GWL_WNDPROC = (-4)
End Enum
Public Enum GWL_STYLE_constants
  CS_DBLCLKS = &H8
End Enum

Public Declare Function ClientToScreen Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long

'Public Declare Function DefWindowProc Lib "user32" Alias "DefWindowProcA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long


'Hook functions

Public Declare Function SetWindowsHookEx Lib "user32" _
            Alias "SetWindowsHookExA" (ByVal idHook As WH, _
                                       ByVal lpfn As Long, _
                                       ByVal hmod As Long, _
                                       ByVal dwThreadId As Long) As Long

Public Declare Function UnhookWindowsHookEx Lib "user32" _
            (ByVal hHook As Long) As Long

Public Declare Function CallNextHookEx Lib "user32" _
            (ByVal hHook As Long, _
            ByVal ncode As Long, _
            ByVal wParam As Long, _
            lParam As Any) As Long


'GDI Object Functions

Public Declare Function SelectObject Lib "gdi32" _
            (ByVal hDC As Long, _
            ByVal hObject As Long) As Long

Public Declare Function DeleteObject Lib "gdi32" _
            (ByVal hObject As Long) As Long

Public Declare Function CreateCompatibleBitmap Lib "gdi32" _
            (ByVal hDC As Long, _
            ByVal nWidth As Long, ByVal nHeight As Long) As Long

Public Declare Function GetCurrentObject Lib "gdi32" _
            (ByVal hDC As Long, _
            ByVal uObjectType As eGDIObjectType) As Long
Public Enum eGDIObjectType
    OBJ_BITMAP = 7
    OBJ_BRUSH = 2
    OBJ_ENHMETADC = 12
    OBJ_ENHMETAFILE = 13
    OBJ_EXTPEN = 11
    OBJ_FONT = 6
    OBJ_MEMDC = 10
    OBJ_METADC = 4
    OBJ_METAFILE = 9
    OBJ_PAL = 5
    OBJ_PEN = 1
    OBJ_REGION = 8
End Enum
'Bitmap functions

'Public Declare Function APIGetDiBits Lib "gdi32" Alias "GetDIBits" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long

Public Declare Function BitBlt Lib "gdi32" _
            (ByVal hDestDC As Long, _
            ByVal x As Long, ByVal y As Long, _
            ByVal nWidth As Long, ByVal nHeight As Long, _
            ByVal hSrcDC As Long, _
            ByVal xSrc As Long, ByVal ySrc As Long, _
            ByVal dwRop As Long) As Long

Public Declare Function StretchBlt Lib "gdi32" _
            (ByVal hDestDD As Long, _
            ByVal x As Long, ByVal y As Long, _
            ByVal nWidth As Long, ByVal nHeight As Long, _
            ByVal hSrcDC As Long, _
            ByVal xSrc As Long, ByVal ySrc As Long, _
            ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, _
            ByVal dwRop As Long) As Long

Public Declare Function GetStretchBltMode Lib "gdi32" _
            (ByVal hDC As Long) As APIStretchMode

Public Declare Function SetStretchBltMode Lib "gdi32" _
            (ByVal hDC As Long, _
            ByVal nStretchMode As APIStretchMode) As Long
Public Enum APIStretchMode
    BLACKONWHITE = 1
    COLORONCOLOR = 3
    HALFTONE = 4
    WHITEONBLACK = 2
End Enum

Public Declare Function GetBitmapDimensionEx Lib "gdi32" _
            (ByVal hBitmap As Long, lpDimension As Size) As Long


'Memory functions

Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)


'Timers functions

Public Declare Function timeBeginPeriod Lib "winmm.dll" (ByVal uPeriod As Long) As Long

Public Declare Function timeEndPeriod Lib "winmm.dll" (ByVal uPeriod As Long) As Long

Public Declare Function timeGetTime Lib "winmm.dll" () As Long


'Drawing functions

Public Declare Function SetPixel Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long

Public Declare Function SetPixelV Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long

Public Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long) As Long


Public Declare Function GetUpdateRect Lib "user32" _
            (ByVal hWnd As Long, _
            lpRect As RECT, _
            ByVal bErase As Long) As Long


'Text drawing functions

Public Declare Function DrawTextEx Lib "user32" _
            Alias "DrawTextExA" (ByVal hDC As Long, _
                                 ByVal lpszString As String, _
                                 ByVal nCharsInString As Long, _
                                 lpRect As RECT, _
                                 ByVal DrawMode As TextDrawMode, _
                                 lpDrawTextParams As DRAWTEXTPARAMS) As Long

Public Declare Function GetTextExtentPoint32 Lib "gdi32" _
            Alias "GetTextExtentPoint32A" (ByVal hDC As Long, _
                                           ByVal lpsz As String, _
                                           ByVal cbString As Long, _
                                           lpSize As Size) As Long
Public Type Size
  w As Long
  h As Long
End Type

Public Declare Function GetCharABCWidths Lib "gdi32" _
            Alias "GetCharABCWidthsA" (ByVal hDC As Long, _
                                      ByVal uFirstChar As Long, _
                                      ByVal uLastChar As Long, _
                                      lpabc As ABC) As Long
Public Type ABC
        abcA As Long
        abcB As Long
        abcC As Long
End Type


'Shell functions

Public Declare Function ShellExecute Lib "shell32" _
       Alias "ShellExecuteA" (ByVal hWnd As Long, _
                              ByVal lpOperation As String, _
                              ByVal lpFile As String, _
                              ByVal lpParameters As String, _
                              ByVal lpDirectory As String, _
                              ByVal nShowCmd As Long) As Long

'Cursor functions

'Public Declare Function SetWindowLong Lib "user32" _
          Alias "SetWindowLongA" (ByVal hwnd As Long, _
                                  ByVal nIndex As Long, _
                                  ByVal dwNewLong As Long) As Long

Public Declare Function GetClassLong Lib "user32" _
          Alias "GetClassLongA" (ByVal hWnd As Long, _
                                 ByVal nIndex As eGCL) As Long
Public Enum eGCL
    GCL_CBCLSEXTRA = (-20)
    GCL_CBWNDEXTRA = (-18)
    GCL_CONVERSION = &H1
    GCL_HBRBACKGROUND = (-10)
    GCL_HCURSOR = (-12)
    GCL_HICON = (-14)
    GCL_HMODULE = (-16)
    GCL_MENUNAME = (-8)
    GCL_REVERSE_LENGTH = &H3
    GCL_REVERSECONVERSION = &H2
    GCL_STYLE = (-26)
    GCL_WNDPROC = (-24)
End Enum

Public Declare Function SetClassLong Lib "user32" _
            Alias "SetClassLongA" (ByVal hWnd As Long, _
                                   ByVal nIndex As eGCL, _
                                   ByVal dwNewLong As Long) As Long


Public Declare Function CallWindowProc Lib "user32" _
            Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, _
                                     ByVal hWnd As Long, _
                                     ByVal Msg As Long, _
                                     ByVal wParam As Long, _
                                     ByVal lParam As Long) As Long


Public Declare Function LoadImage Lib "user32" _
            Alias "LoadImageA" (ByVal hInst As Long, _
                                ByVal Name As String, _
                                ByVal ImgType As VBRUN.LoadResConstants, _
                                ByVal w As Long, _
                                ByVal h As Long, _
                                ByVal Flags As eLR) As Long
Public Enum eLR
  LR_DEFAULTCOLOR = &H0&
  LR_MONOCHROME = &H1&
  LR_COLOR = &H2&
  LR_COPYRETURNORG = &H4&
  LR_COPYDELETEORG = &H8&
  LR_LOADFROMFILE = &H10&
  LR_LOADTRANSPARENT = &H20&
  LR_DEFAULTSIZE = &H40&
  LR_VGACOLOR = &H80&
  LR_LOADMAP3DCOLORS = &H1000&
  LR_CREATEDIBSECTION = &H2000&
  LR_COPYFROMRESOURCE = &H4000&
  LR_SHARED = &H8000&
End Enum

Public Declare Function LoadCursor Lib "user32" _
            Alias "LoadCursorA" (ByVal hInstance As Long, _
                                 ByVal lpCursorName As eIDC) As Long
Public Enum eIDC
    IDC_APPSTARTING = 32650&
    IDC_ARROW = 32512&
    IDC_CROSS = 32515&
    IDC_IBEAM = 32513&
    IDC_ICON = 32641&
    IDC_NO = 32648&
    IDC_SIZE = 32640&
    IDC_SIZEALL = 32646&
    IDC_SIZENESW = 32643&
    IDC_SIZENS = 32645&
    IDC_SIZENWSE = 32642&
    IDC_SIZEWE = 32644&
    IDC_UPARROW = 32516&
    IDC_WAIT = 32514&
End Enum


'Message box functions

Public Declare Function MessageBeep Lib "user32" (ByVal wType As MBStyle) As Long
Public Enum MBStyle
    MB_ICONASTERISK = &H40&
    MB_ICONEXCLAMATION = &H30&
    MB_ICONHAND = &H10&
    MB_ICONINFORMATION = MB_ICONASTERISK
    MB_ICONQUESTION = &H20&
    MB_ICONSTOP = MB_ICONHAND
    MB_OK = &H0&
    MB_PCSPEAKER = &HFFFFFFFF
End Enum



'Known issue - it fails with an error if windowstyle is zero before calling
Public Sub ModifyWindowStyle(ByVal hWnd As Long, _
                             ByVal StyleSet As GWL_STYLE_constants, _
                             ByVal StyleUnset As GWL_STYLE_constants)
Dim WS As Long, WS_old As Long
Dim Ret As Long
WS = GetWindowLong(hWnd, GWL_STYLE)
If WS = 0 Then Err.Raise 4737, "ModifyWindowStyle", "Window style reading failed"
WS_old = WS
WS = WS Or StyleSet
WS = WS And Not StyleUnset
If WS <> WS_old Then
  Ret = SetWindowLong(hWnd, GWL_STYLE, WS)
  If Ret = 0 Then Err.Raise 4737, "ModifyWindowStyle", "Window style writing failed"
End If
End Sub
