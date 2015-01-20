VERSION 5.00
Begin VB.Form frmDrawAFM 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   2596
   ClientLeft      =   33
   ClientTop       =   341
   ClientWidth     =   5764
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2596
   ScaleWidth      =   5764
   StartUpPosition =   3  'Windows Default
   Begin SMBMaker.dbFrame dbFrame2 
      Height          =   1529
      Left            =   2310
      TabIndex        =   7
      Top             =   440
      Width           =   2849
      _ExtentX        =   5019
      _ExtentY        =   2703
      Caption         =   "padding (px)"
      EAC             =   0   'False
      Begin SMBMaker.ctlColor clrBG 
         Height          =   385
         Left            =   1243
         TabIndex        =   12
         Top             =   495
         Width           =   451
         _ExtentX        =   792
         _ExtentY        =   671
         Color           =   16777215
      End
      Begin VB.TextBox txtPadBot 
         Height          =   286
         Left            =   1111
         TabIndex        =   11
         Text            =   "0"
         Top             =   1177
         Width           =   671
      End
      Begin VB.TextBox txtPadRight 
         Height          =   319
         Left            =   2189
         TabIndex        =   10
         Text            =   "0"
         Top             =   561
         Width           =   616
      End
      Begin VB.TextBox txtPadTop 
         Height          =   275
         Left            =   1166
         TabIndex        =   9
         Text            =   "0"
         Top             =   -11
         Width           =   682
      End
      Begin VB.TextBox txtPadLeft 
         Height          =   352
         Left            =   22
         TabIndex        =   8
         Text            =   "0"
         Top             =   528
         Width           =   748
      End
   End
   Begin SMBMaker.dbFrame dbFrame1 
      Height          =   2046
      Left            =   44
      TabIndex        =   2
      Top             =   462
      Width           =   2002
      _ExtentX        =   3536
      _ExtentY        =   3617
      Caption         =   "coloration"
      EAC             =   0   'False
      Begin VB.OptionButton optBriApp 
         Caption         =   "apparent linear"
         Height          =   275
         Left            =   913
         TabIndex        =   15
         Top             =   594
         Width           =   913
      End
      Begin VB.OptionButton optBriLin 
         Caption         =   "true linear"
         Height          =   264
         Left            =   880
         TabIndex        =   14
         Top             =   209
         Value           =   -1  'True
         Width           =   891
      End
      Begin VB.TextBox txtZHI 
         Height          =   330
         Left            =   1034
         TabIndex        =   5
         Text            =   "4.38"
         Top             =   1155
         Width           =   770
      End
      Begin VB.TextBox txtZLO 
         Height          =   319
         Left            =   209
         TabIndex        =   4
         Text            =   "0"
         Top             =   1188
         Width           =   704
      End
      Begin SMBMaker.ctlColor clrHue 
         Height          =   374
         Left            =   165
         TabIndex        =   3
         Top             =   264
         Width           =   429
         _ExtentX        =   752
         _ExtentY        =   650
         Color           =   16757760
      End
      Begin SMBMaker.dbButton btnDrawScale 
         Height          =   385
         Left            =   165
         TabIndex        =   16
         Top             =   1529
         Width           =   1606
         _ExtentX        =   2824
         _ExtentY        =   671
         MouseIcon       =   "frmDrawAFM.frx":0000
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.064
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Others          =   $"frmDrawAFM.frx":001C
         OthersPresent   =   -1  'True
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Z-range"
         Height          =   231
         Left            =   231
         TabIndex        =   6
         Top             =   902
         Width           =   1595
      End
   End
   Begin VB.TextBox txtFile 
      Height          =   319
      Left            =   1188
      TabIndex        =   0
      Top             =   55
      Width           =   4400
   End
   Begin SMBMaker.dbButton btnDraw 
      Height          =   407
      Left            =   3839
      TabIndex        =   13
      Top             =   2046
      Width           =   1320
      _ExtentX        =   2337
      _ExtentY        =   711
      MouseIcon       =   "frmDrawAFM.frx":0071
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.064
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Others          =   $"frmDrawAFM.frx":008D
      OthersPresent   =   -1  'True
   End
   Begin VB.Label Label1 
      Caption         =   "file name"
      Height          =   275
      Left            =   33
      TabIndex        =   1
      Top             =   88
      Width           =   1166
   End
End
Attribute VB_Name = "frmDrawAFM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ColorMode As eAFMColorMode
Enum eAFMColorMode
  afmcmApparentLinear = 0
  afmcmTrueLinear = 1
End Enum

Dim rRef As Double, gRef As Double, bRef As Double

Dim lastH As Long 'height of last image

Private Sub btnDraw_Click()
Dim Topo() As Double
On Error GoTo eh
ReadFile txtFile.Text, Topo
Dim x As Double, y As Double 'in AFM picture
Dim wt As Long, ht As Long 'width-height of AFM picture
Dim orgX As Long, orgY As Long
Dim wi As Long, hi As Long 'width-height of output
Dim picData() As Long
Dim rgbData() As RGBQUAD
PrecalcColors
'calculate and collect widths/paddings
AryWH AryPtr(Topo), wt, ht
lastH = ht
orgX = dbVal(txtPadLeft.Text, vbLong)
orgY = dbVal(txtPadTop.Text, vbLong)
wi = wt + orgX + dbVal(txtPadRight.Text, vbLong)
hi = ht + orgY + dbVal(txtPadBot.Text, vbLong)

'fill data with bg color
ReDim picData(0 To wi - 1, 0 To hi - 1)
Dim BGColor As Long
BGColor = clrBG
For y = 0 To hi - 1
  For x = 0 To wi - 1
    picData(x, y) = BGColor
  Next x
Next y

SwapArys AryPtr(picData), AryPtr(rgbData) 'we'll now use RGBquads
'prepare Zlow/Zhigh
Dim ZLO As Double, ZHI As Double
Dim ZD As Double
Dim relH As Double
ZLO = dbVal(txtZLO.Text)
ZHI = dbVal(txtZHI.Text)
ZD = ZHI - ZLO

'draw
For y = Max(orgY, 0) To Min(orgY + ht, hi) - 1
  For x = Max(orgX, 0) To Min(orgX + wt, wi) - 1
    relH = (Topo(x - orgX, y - orgY) - ZLO) / ZD
    If relH < 0 Then relH = 0
    If relH > 1 Then relH = 1
    Height2Color relH, rgbData(x, y)
  Next x
Next y
SwapArys AryPtr(picData), AryPtr(rgbData) 'back to longs
MainForm.dbMakeSelData 0, 0, picData
Exit Sub
eh:
MsgError
End Sub

Public Sub DrawScale()
Dim wi As Long, hi As Long
wi = 20: hi = lastH
If hi = 0 Then hi = 256
Dim x As Long, y As Long
Dim rgbData() As RGBQUAD, Data() As Long
ReDim rgbData(0 To wi - 1, 0 To hi - 1)
Dim clr As RGBQUAD
PrecalcColors
For y = 0 To hi - 1
  Height2Color 1 - (CDbl(y) / (hi - 1)), clr
  For x = 0 To wi - 1
    rgbData(x, y) = clr
  Next x
Next y
SwapArys AryPtr(rgbData), AryPtr(Data)
MainForm.dbDeselect True
MainForm.dbMakeSelData 0, 0, Data
End Sub

Public Sub ReadFile(ByRef FileName As String, ByRef Topo() As Double)
Dim nmb As Long
nmb = FreeFile
Open FileName For Input As nmb
On Error GoTo eh
  Dim Collector As New StringAccumulator 'the image data is gathered into this before interpretation
  Dim bCollect As Boolean 'indicates if the data area is already reached
  Do Until EOF(nmb)
    Dim l As String
    Line Input #(nmb), l
    If bCollect Then
      Collector.Append l + vbNewLine
    Else
      Dim pos As Long
      pos = InStr(l, "=")
      Dim strValue As String
      Dim strParam As String
      If pos = 0 Then
        strValue = ""
        strParam = Trim$(l)
      Else
        strValue = Trim$(Mid$(l, pos + 1))
        strParam = Trim$(Left$(l, pos - 1))
      End If
      
      Dim ScData As Double
      Dim w As Long, h As Long
      Select Case UCase$(strParam)
        Case "SCALE DATA"
          ScData = dbVal(strValue)
        Case "NX"
          w = dbVal(strValue, vbLong, nMin:=1&)
        Case "NY"
          h = dbVal(strValue, vbLong, nMin:=1&)
        Case "START OF DATA :"
          bCollect = True
      End Select
    End If
  Loop
On Error GoTo 0

Close nmb

'interpret data
Dim tmp As String
tmp = Collector.Content
'convert common delimiters to spaces
tmp = Replace$(tmp, vbNewLine, " ")
tmp = Replace$(tmp, vbTab, " ")
tmp = Replace$(tmp, Chr(10), " ")
tmp = Replace$(tmp, Chr(13), " ")
'get rid of redundant spaces
tmp = Replace$(tmp, ";", " ")
tmp = Replace$(tmp, "  ", " ")
tmp = Replace$(tmp, "  ", " ")
'convert decimal separator
tmp = Replace(tmp, ",", ".")
'split into an array
Dim sArr() As String
sArr = Split(tmp, " ")
ReDim Topo(0 To w - 1, 0 To h - 1)
Dim x As Long, y As Long
For y = 0 To h - 1 'y is top-to-bottom
  For x = 0 To w - 1
    Topo(x, y) = ScData * Val(sArr(x + y * w))
  Next x
Next y
Exit Sub
eh:
PushError
Close nmb
PopError
ErrRaise
End Sub

Friend Sub Height2Color(ByVal relH As Double, ByRef clr As RGBQUAD)
If ColorMode = afmcmApparentLinear Then relH = LinFromSRGB(relH)
If relH < 0 Then relH = 0
If relH > 1 Then relH = 1
Dim r As Double, g As Double, b As Double
'multiply ref color (linear, prenormalized) by desired brightness
r = rRef * relH
g = gRef * relH
b = bRef * relH

'if components clip, spread them into others propotionally
Dim fr As Double, fg As Double, fb As Double
If r > 1 Then
  fg = g / (g + b)
  fb = b / (g + b)
  g = g + fg * (r - 1)
  b = b + fb * (r - 1)
End If
If g > 1 Then
  If r > 1 Then
    fb = 1
    fr = 0
  Else
    fr = r / (r + b)
    fb = b / (r + b)
  End If
  r = r + fr * (g - 1)
  b = b + fb * (g - 1)
End If
If b > 1 Then
  If r > 1 Then
    fr = 0
    fg = 1
  ElseIf g > 1 Then
    fr = 1
    fg = 0
  Else
    fr = r / (r + g)
    fg = g / (r + g)
  End If
  r = r + fr * (b - 1)
  g = g + fg * (b - 1)
End If
'and finally - clip values and output the color
If r > 1 Then r = 1
If g > 1 Then g = 1
If b > 1 Then b = 1
clr.rgbRed = SRGBFromLin(r) * 255
clr.rgbGreen = SRGBFromLin(g) * 255
clr.rgbBlue = SRGBFromLin(b) * 255
End Sub

'input 0..1, output 0..1
Public Function LinFromSRGB(ByVal code As Double) As Double
Dim bri As Double
If code <= 0.03928 Then
  bri = code / 12.92
Else
  bri = ((code + 0.055) / 1.055) ^ 2.4
End If
LinFromSRGB = bri
End Function

'input 0..1, output 0..1
Public Function SRGBFromLin(ByVal bri As Double) As Double
Dim code As Double
If bri <= 0.00304 Then
  code = 12.92 * bri
Else
  code = 1.055 * bri ^ (1# / 2.4) - 0.055
End If
SRGBFromLin = code
End Function


Public Sub PrecalcColors()
Dim briC As Double '(total brightness, 0..1)
Dim rgbHue As RGBQUAD

'get color and convert it to linear
CopyMemory rgbHue, CLng(clrHue), 4
rRef = LinFromSRGB(rgbHue.rgbRed / 255) + 1E-16
gRef = LinFromSRGB(rgbHue.rgbGreen / 255) + 1E-16
bRef = LinFromSRGB(rgbHue.rgbBlue / 255) + 1E-16

'normalize it (may clip as a result)
briC = (rRef + gRef + bRef) / 3
rRef = rRef / briC
gRef = gRef / briC
bRef = bRef / briC
If optBriLin.Value Then ColorMode = afmcmTrueLinear
If optBriApp.Value Then ColorMode = afmcmApparentLinear
End Sub

Private Sub btnDrawScale_Click()
DrawScale
End Sub
