Attribute VB_Name = "SMBMakerDll"
Option Explicit

'**************   SMBMaker.dll  D E C L A R E S   ****************************************
'#Const pDebug = True 'set to false when distributing
'#Const DEMO = True

'#Const DebuggingDll = False 'Set to false when compiling an EXE file
'    Public Declare Sub dllLongToRGBQuad Lib "SMBMaker.dll" (ByRef pData As Any, ByRef pRGB As Any, ByVal nLen As Long)
'    Public Declare Sub dllZoomIn Lib "SMBMaker.dll" Alias "dllFullRebuild" (ByRef pData As Any, ByRef pData2 As Any, ByVal Wdt As Long, ByVal hgt As Long, ByVal Zm As Byte, ByVal bGrid As Boolean, Optional ByVal Flags As Long = 0)
'    Public Declare Sub dllFilterPicture Lib "SMBMaker.dll" _
            (ByRef pData As Any, ByRef tmpData As Any, _
            ByVal pWdt As Long, ByVal pHgt As Long, _
            ByVal x1 As Long, ByVal y1 As Long, _
            ByVal X2 As Long, ByVal Y2 As Long, _
            ByRef fMask As Any, ByVal mWdt As Long, ByVal mHgt As Long, _
            ByVal Flags As Long)
'    Public Declare Sub dllGetRgbQuad Lib "SMBMaker.dll" (ByRef lngColor As Long, ByRef Res As RGBQUAD)
'    Public Declare Sub dllSetProgressWndPos Lib "SMBMaker.dll" (ByVal intLeft As Long, ByVal intTop As Long, ByVal intWidth As Long, ByVal intHeight As Long)
'    Public Declare Sub dllSetProgressProcedure Lib "SMBMaker.dll" (ByVal pProc As Long)
'    Public Declare Sub dllTurnUpsideDown Lib "SMBMaker.dll" (ByRef pData As Any, ByVal Width As Long, ByVal Height As Long)
'    Public Declare Sub dllStretch_Increase Lib "SMBMaker.dll" (ByRef pData As Any, ByRef pRetData As Any, ByVal OldW As Long, ByVal OldH As Long, ByVal NewW As Long, ByVal NewH As Long, ByVal Flags As Long)
'    Public Declare Sub dllStretch_Decrease Lib "SMBMaker.dll" (ByRef pData As Any, ByRef pRetData As Any, ByVal OldW As Long, ByVal OldH As Long, ByVal NewW As Long, ByVal NewH As Long, ByVal Flags As Long)
'    Public Declare Function dllCompareColorsLng Lib "SMBMaker.dll" (ByVal c1 As Long, ByVal c2 As Long) As Long
'    Public Declare Sub dllDrawWaves Lib "SMBMaker.dll" (ByRef pData As Any, ByVal Wdt As Long, ByVal hgt As Long, ByRef WS As Any, ByVal nCount As Long, ByVal Flags As Long)
'    Public Declare Sub dllVtFilter Lib "SMBMaker.dll" _
                  (ByRef Sz As Dims, _
                   ByRef InData As Any, _
                   ByRef OutData As RGBTriLong, _
                   ByRef MaskSz As MaskInfo, _
                   ByRef Mask As RGBTriLong, _
                   ByRef Load As RGBTriLong, _
                   ByVal TexMode As Boolean)
                   
 '   Public Declare Sub dllAddWOfc Lib "SMBMaker.dll" _
                  (ByRef InData As Any, _
                   ByRef OutData As Any, _
                   ByRef Sz As Dims, _
                   ByVal dx As Long, ByVal dy As Long, _
                   ByRef Mul As RGBTriLong, _
                   ByRef Loads As RGBTriLong, _
                   ByVal DetectMax As Long)
'**************///SMBMaker.dll  D E C L A R E S///****************************************


