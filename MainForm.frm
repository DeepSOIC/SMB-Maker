VERSION 5.00
Begin VB.Form MainForm 
   BackColor       =   &H8000000C&
   Caption         =   "SMB Maker"
   ClientHeight    =   7348
   ClientLeft      =   176
   ClientTop       =   770
   ClientWidth     =   11748
   HelpContextID   =   11000
   Icon            =   "MainForm.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MouseIcon       =   "MainForm.frx":8C02
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   668
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1068
   StartUpPosition =   3  'Windows Default
   WhatsThisHelp   =   -1  'True
   Begin VB.Timer tmrStatusFlasher 
      Interval        =   250
      Left            =   5325
      Top             =   1095
   End
   Begin VB.Timer tmrMoveMouser 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   4905
      Top             =   1095
   End
   Begin VB.PictureBox MPHolder 
      AutoRedraw      =   -1  'True
      Height          =   2775
      Left            =   75
      OLEDropMode     =   1  'Manual
      ScaleHeight     =   248
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   250
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   495
      Width           =   2790
      Begin VB.PictureBox MP 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ClipControls    =   0   'False
         Height          =   2280
         Left            =   195
         MouseIcon       =   "MainForm.frx":8F0C
         MousePointer    =   99  'Custom
         OLEDropMode     =   1  'Manual
         ScaleHeight     =   207
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   198
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   165
         Width           =   2175
         Begin VB.PictureBox SelPicture 
            AutoRedraw      =   -1  'True
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            Height          =   1095
            Left            =   1095
            ScaleHeight     =   100
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   44
            TabIndex        =   47
            TabStop         =   0   'False
            Top             =   510
            Visible         =   0   'False
            Width           =   480
         End
         Begin VB.Shape SelRect 
            BorderColor     =   &H00FFFFFF&
            BorderStyle     =   4  'Dash-Dot
            FillColor       =   &H0000FFFF&
            Height          =   1275
            Left            =   345
            Top             =   360
            Visible         =   0   'False
            Width           =   1215
         End
      End
   End
   Begin VB.Timer tmrSelBorderAnimator 
      Interval        =   500
      Left            =   4485
      Top             =   1095
   End
   Begin VB.Timer tmrCheckBackUp 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   4065
      Top             =   1095
   End
   Begin VB.Timer tmrInstall 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   3630
      Top             =   1095
   End
   Begin VB.Timer MPMover 
      Interval        =   1
      Left            =   5295
      Top             =   600
   End
   Begin VB.PictureBox ToolBar2 
      Align           =   1  'Align Top
      ClipControls    =   0   'False
      Height          =   420
      Left            =   0
      OLEDropMode     =   1  'Manual
      ScaleHeight     =   34
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   1064
      TabIndex        =   45
      TabStop         =   0   'False
      ToolTipText     =   "The toolbar"
      Top             =   0
      Width           =   11748
      Begin SMBMaker.dbButton ToolBarButton 
         Height          =   360
         Index           =   3
         Left            =   1320
         TabIndex        =   48
         TabStop         =   0   'False
         Tag             =   "0"
         ToolTipText     =   "Save as..."
         Top             =   0
         Width           =   360
         _ExtentX        =   630
         _ExtentY        =   630
         MouseIcon       =   "MainForm.frx":9216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.064
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Others          =   $"MainForm.frx":9232
         OthersPresent   =   -1  'True
      End
      Begin SMBMaker.dbButton ToolBarButton 
         Height          =   360
         Index           =   0
         Left            =   0
         TabIndex        =   24
         TabStop         =   0   'False
         ToolTipText     =   "New picture"
         Top             =   0
         Width           =   360
         _ExtentX        =   630
         _ExtentY        =   630
         MouseIcon       =   "MainForm.frx":928E
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.064
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Others          =   $"MainForm.frx":92AA
         OthersPresent   =   -1  'True
      End
      Begin SMBMaker.dbButton ToolBarButton 
         Height          =   360
         Index           =   2
         Left            =   945
         TabIndex        =   26
         TabStop         =   0   'False
         Tag             =   "2"
         ToolTipText     =   "Save"
         Top             =   0
         Width           =   360
         _ExtentX        =   630
         _ExtentY        =   630
         MouseIcon       =   "MainForm.frx":9302
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.064
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Others          =   $"MainForm.frx":931E
         OthersPresent   =   -1  'True
      End
      Begin SMBMaker.dbButton ToolBarButton 
         Height          =   360
         Index           =   5
         Left            =   3360
         TabIndex        =   28
         TabStop         =   0   'False
         Tag             =   "4"
         ToolTipText     =   "Undo"
         Top             =   0
         Width           =   360
         _ExtentX        =   630
         _ExtentY        =   630
         MouseIcon       =   "MainForm.frx":9378
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.064
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Others          =   $"MainForm.frx":9394
         OthersPresent   =   -1  'True
      End
      Begin SMBMaker.dbButton ToolBarButton 
         Height          =   360
         Index           =   6
         Left            =   3720
         TabIndex        =   29
         TabStop         =   0   'False
         ToolTipText     =   "Redo"
         Top             =   0
         Width           =   360
         _ExtentX        =   630
         _ExtentY        =   630
         MouseIcon       =   "MainForm.frx":93EF
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.064
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Others          =   $"MainForm.frx":940B
         OthersPresent   =   -1  'True
      End
      Begin SMBMaker.dbButton ToolBarButton 
         Height          =   360
         Index           =   4
         Left            =   2880
         TabIndex        =   27
         TabStop         =   0   'False
         Tag             =   "1"
         ToolTipText     =   "Save palette as..."
         Top             =   0
         Width           =   360
         _ExtentX        =   630
         _ExtentY        =   630
         MouseIcon       =   "MainForm.frx":9466
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.064
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Others          =   $"MainForm.frx":9482
         OthersPresent   =   -1  'True
      End
      Begin SMBMaker.dbButton ToolBarButton 
         Height          =   360
         Index           =   7
         Left            =   4200
         TabIndex        =   30
         TabStop         =   0   'False
         Tag             =   "4"
         ToolTipText     =   "Copy to clipboard"
         Top             =   0
         Width           =   360
         _ExtentX        =   630
         _ExtentY        =   630
         MouseIcon       =   "MainForm.frx":94E3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.064
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Others          =   $"MainForm.frx":94FF
         OthersPresent   =   -1  'True
      End
      Begin SMBMaker.dbButton ToolBarButton 
         Height          =   360
         Index           =   8
         Left            =   4560
         TabIndex        =   31
         TabStop         =   0   'False
         ToolTipText     =   "Paste from clipboard"
         Top             =   0
         Width           =   360
         _ExtentX        =   630
         _ExtentY        =   630
         MouseIcon       =   "MainForm.frx":955A
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.064
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Others          =   $"MainForm.frx":9576
         OthersPresent   =   -1  'True
      End
      Begin SMBMaker.dbButton ToolBarButton 
         Height          =   360
         Index           =   11
         Left            =   6300
         TabIndex        =   35
         TabStop         =   0   'False
         Tag             =   "2"
         ToolTipText     =   "Zoom to"
         Top             =   0
         Width           =   360
         _ExtentX        =   630
         _ExtentY        =   630
         MouseIcon       =   "MainForm.frx":95D3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.064
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Others          =   $"MainForm.frx":95EF
         OthersPresent   =   -1  'True
      End
      Begin SMBMaker.dbButton ToolBarButton 
         Height          =   360
         Index           =   13
         Left            =   7080
         TabIndex        =   37
         TabStop         =   0   'False
         ToolTipText     =   "Zoom In"
         Top             =   0
         Width           =   360
         _ExtentX        =   630
         _ExtentY        =   630
         MouseIcon       =   "MainForm.frx":964A
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.064
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Others          =   $"MainForm.frx":9666
         OthersPresent   =   -1  'True
      End
      Begin SMBMaker.dbButton ToolBarButton 
         Height          =   360
         Index           =   12
         Left            =   6720
         TabIndex        =   36
         TabStop         =   0   'False
         Tag             =   "1"
         ToolTipText     =   "Zoom Out"
         Top             =   0
         Width           =   360
         _ExtentX        =   630
         _ExtentY        =   630
         MouseIcon       =   "MainForm.frx":96C5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.064
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Others          =   $"MainForm.frx":96E1
         OthersPresent   =   -1  'True
      End
      Begin SMBMaker.dbButton ToolBarButton 
         Height          =   360
         Index           =   9
         Left            =   4980
         TabIndex        =   32
         TabStop         =   0   'False
         Tag             =   "1"
         ToolTipText     =   "Load a file and put it into selection"
         Top             =   0
         Width           =   360
         _ExtentX        =   630
         _ExtentY        =   630
         MouseIcon       =   "MainForm.frx":9742
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.064
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Others          =   $"MainForm.frx":975E
         OthersPresent   =   -1  'True
      End
      Begin SMBMaker.dbButton ToolBarButton 
         Height          =   360
         Index           =   10
         Left            =   5460
         TabIndex        =   34
         TabStop         =   0   'False
         Tag             =   "2"
         ToolTipText     =   "Resize image"
         Top             =   0
         Width           =   360
         _ExtentX        =   630
         _ExtentY        =   630
         MouseIcon       =   "MainForm.frx":97B7
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.064
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Others          =   $"MainForm.frx":97D3
         OthersPresent   =   -1  'True
      End
      Begin SMBMaker.dbButton ToolBarButton 
         Height          =   360
         Index           =   1
         Left            =   495
         TabIndex        =   25
         TabStop         =   0   'False
         Tag             =   "2"
         ToolTipText     =   "Open file"
         Top             =   0
         Width           =   360
         _ExtentX        =   630
         _ExtentY        =   630
         MouseIcon       =   "MainForm.frx":9832
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.064
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Others          =   $"MainForm.frx":984E
         OthersPresent   =   -1  'True
      End
   End
   Begin VB.Timer MessageLoopStarter 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   4875
      Top             =   600
   End
   Begin VB.Timer FormResizer 
      Interval        =   5000
      Left            =   4035
      Top             =   600
   End
   Begin VB.PictureBox ToolBar 
      Align           =   4  'Align Right
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E3DFE0&
      Height          =   6075
      Left            =   7172
      Picture         =   "MainForm.frx":98A8
      ScaleHeight     =   548
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   412
      TabIndex        =   44
      TabStop         =   0   'False
      Top             =   420
      Width           =   4575
      Begin VB.CheckBox btnTool 
         BackColor       =   &H00E3DFE0&
         ForeColor       =   &H00808080&
         Height          =   525
         Index           =   20
         Left            =   1080
         Picture         =   "MainForm.frx":A10A
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Tool options..."
         Top             =   3495
         Width           =   540
      End
      Begin VB.CheckBox btnTool 
         BackColor       =   &H00E3DFE0&
         ForeColor       =   &H00808080&
         Height          =   525
         Index           =   19
         Left            =   555
         Picture         =   "MainForm.frx":AD4C
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Tool options..."
         Top             =   3510
         Width           =   540
      End
      Begin VB.PictureBox PctCapture 
         AutoRedraw      =   -1  'True
         ClipControls    =   0   'False
         Height          =   2010
         Left            =   1125
         ScaleHeight     =   179
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   160
         TabIndex        =   46
         TabStop         =   0   'False
         ToolTipText     =   "Lens window. (Right-click for options)"
         Top             =   4005
         Width           =   1800
         Begin VB.Image imgCursor 
            Height          =   352
            Left            =   462
            Picture         =   "MainForm.frx":B98E
            Top             =   770
            Visible         =   0   'False
            Width           =   352
         End
      End
      Begin VB.CheckBox btnTool 
         BackColor       =   &H00E3DFE0&
         ForeColor       =   &H00808080&
         Height          =   525
         Index           =   18
         Left            =   0
         Picture         =   "MainForm.frx":BC98
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   3510
         Width           =   540
      End
      Begin VB.CheckBox btnTool 
         BackColor       =   &H00E3DFE0&
         ForeColor       =   &H00808080&
         Height          =   525
         Index           =   1
         Left            =   540
         Picture         =   "MainForm.frx":C8DA
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Tool options..."
         Top             =   0
         Width           =   540
      End
      Begin VB.CheckBox btnTool 
         BackColor       =   &H00E3DFE0&
         Height          =   525
         Index           =   17
         Left            =   0
         MouseIcon       =   "MainForm.frx":D51C
         Picture         =   "MainForm.frx":D826
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   2730
         Width           =   540
      End
      Begin VB.CheckBox btnTool 
         BackColor       =   &H00E3DFE0&
         Height          =   525
         Index           =   9
         Left            =   3780
         MouseIcon       =   "MainForm.frx":E468
         Picture         =   "MainForm.frx":E772
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Tool options..."
         Top             =   675
         Width           =   540
      End
      Begin VB.CheckBox btnTool 
         BackColor       =   &H00E3DFE0&
         Height          =   525
         Index           =   8
         Left            =   3240
         MouseIcon       =   "MainForm.frx":F3B4
         Picture         =   "MainForm.frx":F6BE
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Свойства инструмента"
         Top             =   675
         Width           =   540
      End
      Begin VB.CheckBox btnTool 
         BackColor       =   &H00E3DFE0&
         Height          =   525
         Index           =   16
         Left            =   2175
         MouseIcon       =   "MainForm.frx":10300
         Picture         =   "MainForm.frx":1060A
         Style           =   1  'Graphical
         TabIndex        =   17
         Tag             =   "14"
         ToolTipText     =   "Свойства инструмента"
         Top             =   2055
         Width           =   540
      End
      Begin VB.CheckBox btnTool 
         BackColor       =   &H00E3DFE0&
         Height          =   525
         Index           =   6
         Left            =   2160
         MouseIcon       =   "MainForm.frx":1124C
         Picture         =   "MainForm.frx":11556
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   675
         Width           =   540
      End
      Begin VB.CheckBox btnTool 
         BackColor       =   &H00E3DFE0&
         Height          =   525
         Index           =   4
         Left            =   1080
         MouseIcon       =   "MainForm.frx":12198
         Picture         =   "MainForm.frx":124A2
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   675
         Width           =   540
      End
      Begin VB.CheckBox btnTool 
         BackColor       =   &H00E3DFE0&
         Height          =   525
         Index           =   0
         Left            =   0
         MouseIcon       =   "MainForm.frx":130E4
         Picture         =   "MainForm.frx":133EE
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   0
         Width           =   540
      End
      Begin VB.CheckBox btnTool 
         BackColor       =   &H00E3DFE0&
         Height          =   525
         Index           =   11
         Left            =   555
         MouseIcon       =   "MainForm.frx":14030
         Picture         =   "MainForm.frx":1433A
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   1335
         Width           =   540
      End
      Begin VB.CheckBox btnTool 
         BackColor       =   &H00E3DFE0&
         Height          =   525
         Index           =   10
         Left            =   0
         MouseIcon       =   "MainForm.frx":14F7C
         Picture         =   "MainForm.frx":15286
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   1335
         Width           =   555
      End
      Begin VB.CheckBox btnTool 
         BackColor       =   &H00E3DFE0&
         Height          =   525
         Index           =   14
         Left            =   1095
         MouseIcon       =   "MainForm.frx":15EC8
         Picture         =   "MainForm.frx":161D2
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   2055
         Width           =   540
      End
      Begin VB.CheckBox btnTool 
         BackColor       =   &H00E3DFE0&
         Height          =   525
         Index           =   13
         Left            =   555
         MouseIcon       =   "MainForm.frx":16E14
         Picture         =   "MainForm.frx":1711E
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   2055
         Width           =   540
      End
      Begin VB.CheckBox btnTool 
         BackColor       =   &H00E3DFE0&
         Height          =   525
         Index           =   5
         Left            =   1620
         MouseIcon       =   "MainForm.frx":17D60
         Picture         =   "MainForm.frx":1806A
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   675
         Width           =   540
      End
      Begin VB.CheckBox btnTool 
         BackColor       =   &H00E3DFE0&
         Height          =   525
         Index           =   15
         Left            =   1635
         MouseIcon       =   "MainForm.frx":18CAC
         Picture         =   "MainForm.frx":18FB6
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   2055
         Width           =   540
      End
      Begin VB.CheckBox btnTool 
         BackColor       =   &H00E3DFE0&
         Height          =   525
         Index           =   7
         Left            =   2700
         MouseIcon       =   "MainForm.frx":19BF8
         Picture         =   "MainForm.frx":19F02
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   675
         Width           =   540
      End
      Begin VB.CheckBox btnTool 
         BackColor       =   &H00E3DFE0&
         Height          =   525
         Index           =   12
         Left            =   15
         MouseIcon       =   "MainForm.frx":1AB44
         Picture         =   "MainForm.frx":1AE4E
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   2055
         Width           =   540
      End
      Begin VB.CheckBox btnTool 
         BackColor       =   &H00E3DFE0&
         Height          =   525
         Index           =   3
         Left            =   540
         MouseIcon       =   "MainForm.frx":1BA90
         Picture         =   "MainForm.frx":1BD9A
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   675
         Value           =   1  'Checked
         Width           =   540
      End
      Begin VB.CheckBox btnTool 
         BackColor       =   &H00FFFFFF&
         Height          =   525
         Index           =   2
         Left            =   0
         MouseIcon       =   "MainForm.frx":1C9DC
         Picture         =   "MainForm.frx":1CCE6
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   675
         Width           =   540
      End
   End
   Begin VB.Timer Mover 
      Interval        =   16
      Left            =   3615
      Top             =   600
   End
   Begin VB.PictureBox frmColors 
      Align           =   2  'Align Bottom
      BackColor       =   &H00E3DFE0&
      Height          =   570
      Left            =   0
      ScaleHeight     =   48
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   1064
      TabIndex        =   43
      TabStop         =   0   'False
      Top             =   6495
      Width           =   11748
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      AutoRedraw      =   -1  'True
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.07
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   0
      ScaleHeight     =   22
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   1064
      TabIndex        =   41
      TabStop         =   0   'False
      ToolTipText     =   "Строка состояния."
      Top             =   7065
      Width           =   11748
      Begin VB.Label Status 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Design..."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.07
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   60
         TabIndex        =   42
         ToolTipText     =   "Status bar"
         Top             =   -15
         Width           =   10215
      End
   End
   Begin VB.VScrollBar VScroll 
      Enabled         =   0   'False
      Height          =   2925
      LargeChange     =   3
      Left            =   6165
      SmallChange     =   16
      TabIndex        =   40
      TabStop         =   0   'False
      Top             =   390
      Width           =   345
   End
   Begin VB.HScrollBar HScroll 
      Enabled         =   0   'False
      Height          =   270
      LargeChange     =   3
      Left            =   0
      SmallChange     =   16
      TabIndex        =   39
      TabStop         =   0   'False
      Top             =   3375
      Width           =   5655
   End
   Begin VB.PictureBox TempBox 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   750
      Left            =   3045
      ScaleHeight     =   68
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   76
      TabIndex        =   38
      TabStop         =   0   'False
      Top             =   2265
      Visible         =   0   'False
      Width           =   840
   End
   Begin SMBMaker.dbButton ActiveColor 
      Height          =   435
      Index           =   1
      Left            =   5730
      TabIndex        =   22
      Top             =   3330
      Width           =   435
      _ExtentX        =   772
      _ExtentY        =   772
      MouseIcon       =   "MainForm.frx":1D928
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.064
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Others          =   $"MainForm.frx":1E202
      OthersPresent   =   -1  'True
   End
   Begin SMBMaker.dbButton ActiveColor 
      Height          =   420
      Index           =   2
      Left            =   5895
      TabIndex        =   23
      Top             =   3495
      Width           =   420
      _ExtentX        =   732
      _ExtentY        =   732
      MouseIcon       =   "MainForm.frx":1E241
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.064
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Others          =   $"MainForm.frx":1EB1B
      OthersPresent   =   -1  'True
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuNew 
         Caption         =   "&New image..."
      End
      Begin VB.Menu mnuClear 
         Caption         =   "&Clear picture"
      End
      Begin VB.Menu FileSep0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLoad 
         Caption         =   "&Load picture..."
      End
      Begin VB.Menu mnuFileImportAscii 
         Caption         =   "&Import ascii..."
      End
      Begin VB.Menu FileSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSaveFile 
         Caption         =   "&Save"
      End
      Begin VB.Menu mnuFolder 
         Caption         =   "Open file's &folder"
      End
      Begin VB.Menu FileSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSaveSel 
         Caption         =   "Save se&lection"
      End
      Begin VB.Menu FileSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSaveAs 
         Caption         =   "Save as..."
      End
      Begin VB.Menu FileSep4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBuildBackUp 
         Caption         =   "Build &Backup"
      End
      Begin VB.Menu FileSep5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPrint 
         Caption         =   "&Print"
      End
      Begin VB.Menu FileSep6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Quit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuUndo 
         Caption         =   "&Undo"
      End
      Begin VB.Menu mnuRedo 
         Caption         =   "&Redo"
         Visible         =   0   'False
      End
      Begin VB.Menu EditSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuClearUndo 
         Caption         =   "Clear Undo History"
      End
      Begin VB.Menu EditSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCopy 
         Caption         =   "&Copy to clipboard"
      End
      Begin VB.Menu mnuPaste 
         Caption         =   "&Paste"
      End
      Begin VB.Menu EditSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSelectAll 
         Caption         =   "&Select all"
      End
      Begin VB.Menu mnuClear2 
         Caption         =   "Clear picture"
      End
      Begin VB.Menu EditSep4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuResize 
         Caption         =   "Image si&ze..."
      End
      Begin VB.Menu EditSep5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMix 
         Caption         =   "Insert from &file..."
      End
      Begin VB.Menu mnuEditSel 
         Caption         =   "Edit Selection..."
      End
      Begin VB.Menu mnuCapture 
         Caption         =   "Capture..."
      End
      Begin VB.Menu EditSep6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCrop 
         Caption         =   "Crop to selection"
      End
      Begin VB.Menu EditSep8 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuDeadEnds 
         Caption         =   "Fill dead ends"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuPal 
         Caption         =   "&Palette"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuToolBarVis 
         Caption         =   "&Toolbar"
      End
      Begin VB.Menu mnuDynamicScr 
         Caption         =   "Dynamic scrolling"
      End
      Begin VB.Menu ViewSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuZoom 
         Caption         =   "Zoom..."
      End
      Begin VB.Menu ViewSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRefresh 
         Caption         =   "&Refresh image"
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tool"
      Begin VB.Menu mnuTool 
         Caption         =   "Selection"
         Index           =   0
         Tag             =   "10"
      End
      Begin VB.Menu mnuTool 
         Caption         =   "Insert text"
         Index           =   1
         Tag             =   "18"
      End
      Begin VB.Menu mnuTool 
         Caption         =   "&Pencil"
         Index           =   2
         Tag             =   "0"
      End
      Begin VB.Menu mnuTool 
         Caption         =   "&Line"
         Checked         =   -1  'True
         Index           =   3
         Tag             =   "1"
      End
      Begin VB.Menu mnuTool 
         Caption         =   "&Rectangle"
         Index           =   4
         Tag             =   "11"
      End
      Begin VB.Menu mnuTool 
         Caption         =   "&Circle"
         Index           =   5
         Tag             =   "8"
      End
      Begin VB.Menu mnuTool 
         Caption         =   "Pol&ygon"
         Index           =   6
         Tag             =   "12"
      End
      Begin VB.Menu mnuTool 
         Caption         =   "&Star"
         Index           =   7
         Tag             =   "3"
      End
      Begin VB.Menu mnuTool 
         Caption         =   "&Airbrush"
         Index           =   8
         Tag             =   "13"
      End
      Begin VB.Menu mnuTool 
         Caption         =   "Brush"
         Index           =   9
         Tag             =   "16"
      End
      Begin VB.Menu mnuTool 
         Caption         =   "&Paint"
         Index           =   10
         Tag             =   "7"
      End
      Begin VB.Menu mnuTool 
         Caption         =   "&Color select"
         Index           =   11
         Tag             =   "5"
      End
      Begin VB.Menu mnuTool 
         Caption         =   "&Fade line"
         Index           =   12
         Tag             =   "2"
      End
      Begin VB.Menu mnuTool 
         Caption         =   "Fade vertical"
         Index           =   13
         Tag             =   "6"
      End
      Begin VB.Menu mnuTool 
         Caption         =   "Fade horizontal"
         Index           =   14
         Tag             =   "9"
      End
      Begin VB.Menu mnuTool 
         Caption         =   "Fade star"
         Index           =   15
         Tag             =   "4"
      End
      Begin VB.Menu mnuTool 
         Caption         =   "Heli&x"
         Index           =   16
         Tag             =   "14"
      End
      Begin VB.Menu mnuTool 
         Caption         =   "Palette"
         Index           =   17
         Tag             =   "17"
      End
      Begin VB.Menu mnuTool 
         Caption         =   "Change origin"
         Index           =   18
         Tag             =   "20"
      End
      Begin VB.Menu mnuTool 
         Caption         =   "Programmable"
         Index           =   19
         Tag             =   "21"
      End
      Begin VB.Menu ToolSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuToolOpts 
         Caption         =   "Tool properties"
      End
      Begin VB.Menu mnuSelBrush 
         Caption         =   "Use Selection As Brush"
      End
   End
   Begin VB.Menu mnuSelAll 
      Caption         =   "Selection"
      Visible         =   0   'False
      Begin VB.Menu mnuSelShow 
         Caption         =   "Show"
      End
      Begin VB.Menu mnuSelMoveTo 
         Caption         =   "Move to..."
      End
      Begin VB.Menu SelSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSelCenterHorz 
         Caption         =   "Center horizontally"
      End
      Begin VB.Menu mnuSelCenterVert 
         Caption         =   "Center vertically"
      End
      Begin VB.Menu SelSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSelClear 
         Caption         =   "Clear (with back-color)"
      End
      Begin VB.Menu mnuSelResize 
         Caption         =   "Resize..."
      End
      Begin VB.Menu mnuSelEditSel 
         Caption         =   "Edit..."
      End
      Begin VB.Menu SelSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSelStamp 
         Caption         =   "Stamp"
      End
      Begin VB.Menu mnuSelDelete 
         Caption         =   "Delete"
      End
      Begin VB.Menu mnuSelAutoRepaint 
         Caption         =   "Toggle redraw on move"
      End
      Begin VB.Menu SelSep4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSelShowSize 
         Caption         =   "Show sel size"
      End
   End
   Begin VB.Menu mnuDraw 
      Caption         =   "Draw"
      Begin VB.Menu mnuGrid 
         Caption         =   "Grid"
      End
      Begin VB.Menu DrawSep0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDrawPlain 
         Caption         =   "Plain"
      End
      Begin VB.Menu mnuDrawBBg 
         Caption         =   "BBG"
      End
      Begin VB.Menu mnuDrawBG 
         Caption         =   "Background"
      End
      Begin VB.Menu DrawSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDrawWaves 
         Caption         =   "Waves..."
      End
      Begin VB.Menu DrawSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFormula 
         Caption         =   "Formula (OLE)..."
      End
      Begin VB.Menu DrawSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDrawTopoR 
         Caption         =   "Render TopoR PCB..."
      End
      Begin VB.Menu mnuDrawAFM 
         Caption         =   "Nova AFM image (ascii)..."
      End
   End
   Begin VB.Menu mnuTexture 
      Caption         =   "Texture"
      Begin VB.Menu mnuTexMode 
         Caption         =   "Texture Mode"
      End
      Begin VB.Menu TexSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuResetOrg 
         Caption         =   "Reset Origin"
      End
      Begin VB.Menu mnuRestoreOrg 
         Caption         =   "Restore Origin"
      End
   End
   Begin VB.Menu mnuPrgs 
      Caption         =   "Programmable"
      Begin VB.Menu mnuPrgDraw 
         Caption         =   "Drawings"
      End
   End
   Begin VB.Menu mnuEffects 
      Caption         =   "&Effects"
      Begin VB.Menu mnuEffect 
         Caption         =   "Soften"
         Index           =   0
      End
      Begin VB.Menu mnuEffect 
         Caption         =   "Сделать что-то непонятное"
         Index           =   1
         Tag             =   "1"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuEffect 
         Caption         =   "Color scale correction"
         Index           =   2
         Tag             =   "3"
      End
      Begin VB.Menu mnuEffect 
         Caption         =   "Negative"
         Index           =   3
         Tag             =   "2"
      End
      Begin VB.Menu mnuEffect 
         Caption         =   "Flip/Rotate"
         Index           =   4
         Tag             =   "4"
      End
      Begin VB.Menu mnuEffect 
         Caption         =   "Decolour"
         Index           =   5
         Tag             =   "5"
      End
      Begin VB.Menu mnuEffect 
         Caption         =   "RGB Matrix"
         Index           =   6
         Tag             =   "6"
      End
      Begin VB.Menu mnuEffect 
         Caption         =   "Make monochrome"
         Index           =   7
         Tag             =   "7"
      End
      Begin VB.Menu mnuEffect 
         Caption         =   "Contrast"
         Index           =   8
         Tag             =   "8"
      End
      Begin VB.Menu mnuEffect 
         Caption         =   "Replace Colors"
         Index           =   9
         Tag             =   "9"
      End
      Begin VB.Menu mnuEffect 
         Caption         =   "Differentiate"
         Index           =   10
         Tag             =   "10"
      End
      Begin VB.Menu mnuEffect 
         Caption         =   "RB offset (ClearType)"
         Index           =   11
         Tag             =   "11"
      End
      Begin VB.Menu EffectsSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLastEffect 
         Caption         =   "Repeat last effect"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu mnuPalette 
      Caption         =   "Palette"
      Begin VB.Menu mnuLoadPal 
         Caption         =   "Load Palette..."
      End
      Begin VB.Menu mnuSavePal 
         Caption         =   "Save palette..."
      End
      Begin VB.Menu PalSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuResetPal 
         Caption         =   "Reset"
      End
      Begin VB.Menu mnuPalDef 
         Caption         =   "Save palette as default"
      End
      Begin VB.Menu PalSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPalCount 
         Caption         =   "Palette items count..."
      End
      Begin VB.Menu mnuStretchPal 
         Caption         =   "Stretch palette..."
      End
      Begin VB.Menu PalSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFillTips 
         Caption         =   "Fill tips automatically"
      End
      Begin VB.Menu mnuEmptyTips 
         Caption         =   "Empty All Tips"
      End
      Begin VB.Menu PalSep4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAutoPal 
         Caption         =   "Predefined Palettes"
         Begin VB.Menu mnuSysPal 
            Caption         =   "System colors palette"
         End
         Begin VB.Menu mnuQBColors 
            Caption         =   "QB Colors"
         End
         Begin VB.Menu ppSep1 
            Caption         =   "-"
         End
         Begin VB.Menu mnuDefPal16 
            Caption         =   "Default 16-color"
         End
         Begin VB.Menu mnuDefPal256 
            Caption         =   "Default 256-color"
         End
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
      Begin VB.Menu mnuWheelUse 
         Caption         =   "Use middle mouse button for"
         Begin VB.Menu mnuUseWheel 
            Caption         =   "color detection and scrolling"
            Checked         =   -1  'True
            Index           =   0
         End
         Begin VB.Menu mnuUseWheel 
            Caption         =   "scrolling only"
            Index           =   1
         End
      End
      Begin VB.Menu mnuKeyb 
         Caption         =   "Keys..."
      End
      Begin VB.Menu mnuKeyb2 
         Caption         =   "Shortcuts..."
      End
      Begin VB.Menu mnuPSens 
         Caption         =   "Pressure sensitive device"
      End
      Begin VB.Menu OptionsSep0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSetAutoScrolling 
         Caption         =   "Scrolling setup..."
      End
      Begin VB.Menu mnuMouseAttr 
         Caption         =   "Mouse attracted to scrolling"
      End
      Begin VB.Menu OptionsSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuUndoLim 
         Caption         =   "Undo count..."
      End
      Begin VB.Menu mnuNoUndoRedo 
         Caption         =   "Disable Undo/Redo"
      End
      Begin VB.Menu OptionsSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuShowSplash 
         Caption         =   "Show Splash on Load"
      End
      Begin VB.Menu mnuOptRememberWndPos 
         Caption         =   "Remember window position"
      End
      Begin VB.Menu OptionsSep4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuReg 
         Caption         =   "Register types"
      End
      Begin VB.Menu OptionsSep5 
         Caption         =   "-"
      End
      Begin VB.Menu OptionsSep7 
         Caption         =   "-"
      End
      Begin VB.Menu mnuResetAll 
         Caption         =   "Reset all settings"
      End
      Begin VB.Menu mnuUnInstall 
         Caption         =   "Uninstall SMB Maker"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHowToUse 
         Caption         =   "How to use?"
      End
      Begin VB.Menu mnuIdleMessage 
         Caption         =   "Show Tip in status bar"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "About SMB Maker..."
         Shortcut        =   ^{F1}
      End
      Begin VB.Menu HelpSep1 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuRegister 
         Caption         =   "Register..."
         Shortcut        =   +{F1}
      End
      Begin VB.Menu HelpSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuWeb 
         Caption         =   "VT web page"
      End
      Begin VB.Menu mnuWebUpdates 
         Caption         =   "Check for updates"
      End
      Begin VB.Menu mnuWebMail 
         Caption         =   "Mail VT (bug, question, etc)"
      End
      Begin VB.Menu mnuWebForum 
         Caption         =   "Visit forum (ideas, bugs, questions, etc)"
      End
   End
   Begin VB.Menu mnuWhatsThis 
      Caption         =   "?"
      Visible         =   0   'False
   End
   Begin VB.Menu mnuPopLns 
      Caption         =   "<lens pop-up menu>"
      Visible         =   0   'False
      Begin VB.Menu mnuLnsZoomIn 
         Caption         =   "Zoom in	Shift+LMB"
      End
      Begin VB.Menu mnuLnsZoomOut 
         Caption         =   "Zoom Out	Shift+RMB"
      End
      Begin VB.Menu lnsSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLnsToggle 
         Caption         =   "Enable/Disable	Alt+LMB"
      End
      Begin VB.Menu lnsSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLnsDock 
         Caption         =   "Dock into SMB Maker	Ctrl+LMB"
      End
      Begin VB.Menu lnsSep3 
         Caption         =   "-"
      End
      Begin VB.Menu lnsHint1 
         Caption         =   "LMB = left"
         Enabled         =   0   'False
      End
      Begin VB.Menu lnsHint2 
         Caption         =   "  mouse button"
         Enabled         =   0   'False
      End
      Begin VB.Menu lnsHint3 
         Caption         =   "RMB = right"
         Enabled         =   0   'False
      End
      Begin VB.Menu lnsHint4 
         Caption         =   "  mouse button"
         Enabled         =   0   'False
      End
   End
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Data() As Long
Dim DataAlpha() As Long
Dim intH As Long, intW As Long
Dim SelW As Long, SelH As Long
Dim pngBPP As Long

Dim ACol(1 To 2) As Long
Dim x0 As Long, y0 As Long
Dim XO As Long, YO As Long
Dim XOP As Long, YOP As Long
Dim fx0 As Double, fy0 As Double
Dim fXO As Double, fYO As Double
Dim dblX0 As Double, dblY0 As Double

Dim ScrollSettings As typScrollSettings

Dim FormW As Long, FormH As Long
'Dim FormL As Long, FormT As Long

Dim UndoData As dbUndoStorage
Dim RedoData As dbUndoStorage
Dim UndoSize As Long
Dim DisUndoRedo As Boolean
'Dim UndoIndex As Long, RedoIndex As Long

Dim TUData() As dbPixel
Dim TUSz As Long, TUPointer As Long
Public TempDrawingMode As Boolean

Dim CurSel As dbSelection_class
Dim SelMatrix() As Double
Dim TransDataChanged As Boolean
Dim CurPol As dbPolygon
'Dim Scroll_Speed As Long

'Dim ToolBarPicture As IPictureDisp
Dim FastLoad As Boolean
Dim Steps(0 To 3) As Long, Edins(0 To 3) As Boolean
Dim CurKeyPreset As String
Dim SelPictureAutoRepaint As Boolean

Dim CircleFlags As dbCircleFlags
Dim DrawingCircle As Boolean
Dim gFDSC As FadeDesc
Dim HSet As HelixSettings
Dim FillOpts As FillSettings

Dim LineOpts As LineSettings
Dim DrawingLine As Boolean, LineK As Single, newLineK As Single
Dim LineStyle As Long, NewLineStyle As Long
Dim MWM As Long
Dim CancelActiveColorModeChange As Boolean

Dim prgToolProg As SMP
'Dim PrgToolProg() As Variant, ToolVars() As Variable

Dim Zm As Integer
Dim dbRS As dbRectStyle
Dim dbMS(1 To 3, 1 To 2) As dbMouseState 'first dimension is button index; the second - 1 for events and 2 for KB
Dim WSMS(1 To 4) As Boolean
Public MDPen As Boolean 'if last mousedown was done with the pen, public to be accessible by subclass procedure
Dim WithEvents MPMS As clsAntiDblClick
Attribute MPMS.VB_VarHelpID = -1
Dim WithEvents FormMS As clsAntiDblClick
Attribute FormMS.VB_VarHelpID = -1
Dim MouseErr As vtError
Dim LUT As Integer
Dim PrevSS As dbShiftConstants
Public OpenedFileName As String ', OpenedFileType As dbOpenedFileTypes
Public OpenedFileFormatID As String
Dim FileChanged As Boolean
Dim LastEffectIndex As Integer
Dim CurBrush() As Byte
Dim ToolBar2Visible As Boolean

Dim gTimer As Long
'Dim SmthScrl As Boolean
'Dim SmthHlast As Long, SmthVlast As Long
Dim NeedRefr As Boolean
Dim LastPalIndex As Integer
Dim prvMeEnabled As Boolean
Dim MeEnabledStack As New clsStack
Dim FreezeRefresh As Boolean

Dim ChCol() As Pal_Entry

Dim dbTag As String

Dim tdN As Long
Dim tdRGB As RGBQuadLong

Public PenPressure As Long
Public MaxPenPressure As Long

Dim LastUserActionTime As Long

Dim LastPickedColor As Long

Public TexMode As Boolean
Dim TexOrg As POINTAPI
Dim OrgUndoBuilt As Boolean

Dim PctCaptureDisabled As Boolean
Dim PctCaptureZm As Long

Dim BackUpBuildDelay As Long


Dim MPBitsRGB() As RGBQUAD
Dim MPBitsLNG() As Long
Dim MPhDefBitmap As Long
Dim MPBitsWidth As Long
Dim MPBitsHeight As Long

Dim VScrollEnabled As Boolean
Dim HScrollEnabled As Boolean

Dim NaviMousePos As POINTAPI
Dim NaviEnabled As Boolean

Dim StatusFlashesLeft As Long

#Const GetDIBitsErrors = True

Dim Downed As Boolean

'transform is currently only accessible from prog tool
'when BeginTransfrom is called, transformdata is created and it becomes target data
'clear of transfromdata caused by:  begintransfrom(-1,....), tool change, undo
'to apply transformdata, call EndTransform
'if begintransfrom was not called, transformblock will transform pic onto itself, with possible overlapping-related problems
Dim TransformData() As Long
'undo's:
'bud is called on EndTransform only
'if no transformdata then storefragment,startpixelaction
'also, empty pixel undoes should be autodeleted

Private Type dbPolygon
    Button As Integer
    Active As Boolean
    bx As Long
    by As Long
End Type

Private Type dbColor
    Comp(1 To 3) As Long
End Type

Private Type dbForStretch
    Color As RGBQUAD
    Rasst As Single
End Type

Public Enum GREnum
dbNoGrid = 0
dbGrid = 1
dbAsmnuGrid = 2
End Enum

Public Enum dbDialogAction
dbShowOpen = 1
dbShowSave = 2
dbShowColor = 3
dbShowFont = 4
'dbShowPrinter = 5
'dbShowHelp = 6
End Enum

Public Enum dbSelMode
    dbReplace = 0
    dbMerge = 1
    dbAdd = 2
    dbOR = 3
    dbAND = 4
    dbXOR = 5
    dbIMP = 6
    dbNOT = 7
    dbEQV = 8
    dbTransparent = 9
    dbSuperTransparent = 10
    dbMatrixMixed = 11
    dbOverlayed = 12
    dbUseCurSelMode = 255
End Enum

Public Enum dbRectStyle
    dbEmpty = 0
    dbFilled = &H1
    dbFilledBG = &H3
End Enum

Public Enum dbTurnMethod
    dbFlipHor = 0
    dbFlipVer = 1
    dbTurn90 = 2
    dbTurn180 = 3
    dbTurn270 = 4
End Enum

Public Enum dbMouseState
    dbButtonDown = 1
    dbButtonUp = 0
End Enum

Public Enum eDrawMode
  dmNormal = 0
  dm3opaq = 1
  dmMinimum = 2
  dmMaximum = 3
End Enum

Sub LoadFileEx(ByRef pData() As Long, _
               ByVal File As String, _
               Optional ByVal ReadMaskToSelTransData As Boolean = True)
Dim Alpha() As Long
vtLoadPicture pData, Alpha, File
If ReadMaskToSelTransData Then
    SwapArys AryPtr(TransOrigData), AryPtr(Alpha)
    Erase TransData
    Erase Alpha
End If
End Sub

Public Sub LoadFile(File As String)
    Dim fID As String
    On Error GoTo eh
    vtLoadPicture Data, DataAlpha, File, fID, UpdateSettings:=True, Purpose:="MP"
    FileChanged = False
    NeedRefr = True
    ReZoom 1
    If NeedRefr Then Refr
rsm:
    'Refr
    OrgUndoBuilt = False
    mnuResetOrg_Click
    
    OpenedFileName = File
    OpenedFileFormatID = fID
    FreshCaption
    ClearUndo

'End If
Exit Sub
eh:
If Err.Number = errNewFile Then
    NewPicture
    FileChanged = True
    Resume rsm
Else
    ErrRaise "LoadFile"
End If
End Sub

Public Sub ValidateZoom()
ToolBarButton(GetTLBIndex(cmdZoomOut)).Enabled = Not (Zm <= 1)
ToolBarButton(GetTLBIndex(cmdZoomIn)).Enabled = Not (Zm >= 32)
End Sub

Public Sub UpdateWH()
If AryDims(AryPtr(Data)) <> 2 Then
    intW = 0
    intH = 0
Else
    intW = UBound(Data, 1) + 1
    intH = UBound(Data, 2) + 1
End If
If SelectionPresent Then
    SelW = UBound(CurSel_SelData, 1) + 1
    SelH = UBound(CurSel_SelData, 2) + 1
    CurSel.x2 = CurSel.x1 + SelW - 1
    CurSel.y2 = CurSel.y1 + SelH - 1
Else
    SelW = 0
    SelH = 0
End If
End Sub

Public Sub Resize(ByVal NewW As Long, ByVal NewH As Long, _
                  Optional ByVal Stretch As Boolean = True, _
                  Optional ByVal PreserveValues As Boolean = True, _
                  Optional ByVal StoreUndo As Boolean = True, _
                  Optional ByVal StretchMode As eStretchMode = SMSquares)
Dim tmp As Integer, y As Long, x As Long
Dim Temp() As Long
Dim tmpZm As Integer
UpdateWH
If intW = NewW And intH = NewH Then
    MPHolder_Resize
    Exit Sub
End If
If StoreUndo Then
    BUD
End If
If PreserveValues Then
    On Error GoTo eh
    If Not Stretch Then
        vtStretchPreserve Data, NewW, NewH, ACol(2)
    Else
        dbStretch Data, NewW, NewH, StretchMode, True
    End If
lRes:
    On Error GoTo 0
Else
    ReDim Data(NewW - 1, NewH - 1)
End If
'RestoreMP
intW = NewW
intH = NewH

Refr
NeedRefr = False
HScroll_Change
VScroll_Change
Exit Sub
Resume
eh:
ErrRaise "Resize"
    
End Sub

Private Sub ActiveColor_Click(Index As Integer)
Dim tColor As Long, tmp As RGBQUAD
ShowStatus 10070

tColor = ConvertColorLng(ACol(Index))

On Error GoTo eh
tColor = CDl.PickColor(tColor, ConvertColors:=False)
GetRgbQuadEx tColor, tmp
ShowStatus grs(10072, "%r", CStr(tmp.rgbRed), _
                      "%g", CStr(tmp.rgbGreen), _
                      "%b", CStr(tmp.rgbBlue), _
                      "$h", Hex$(tColor), _
                      "#a", BoolToStr_OnOff(IIf(Index = 1, gFDSC.AutoColor1, gFDSC.AutoColor2))) _
          , , 3

On Error GoTo 0
ChangeActiveColor Index, ConvertColorLng(tColor)

ExitHere:
On Error Resume Next
MP.SetFocus

Exit Sub
eh:
If Err.Number = dbCWS Then
    ShowStatus STT_Cancelled, , 1
    Resume ExitHere
    Exit Sub
Else
    MsgError
End If
End Sub

'indexes: 1 for foreground, 2 for backgroung
Public Sub ChangeActiveColor(ByVal Index As Integer, ByVal lngColor As Long, Optional ByVal setPer As Boolean = True)
Attribute ChangeActiveColor.VB_Description = "Changes active color"
On Error GoTo eh
ACol(Index) = lngColor
If (IIf(Index = 1, gFDSC.AutoColor1, gFDSC.AutoColor2)) And setPer And IsAutoColorTool(ActiveTool) Then
    If dbMsgBox(GRSF(1101), vbQuestion Or vbYesNo) = vbYes Then
        If Index = 1 Then
            gFDSC.AutoColor1 = False
        Else
            gFDSC.AutoColor2 = False
        End If
    End If
End If
FreshActiveColors
Exit Sub
eh:
If dbMsgBox(GRSF(1102), vbYesNo Or vbExclamation) = vbYes Then
    dbRepair
End If

End Sub

Function IsAutoColorTool(ToolIndex As Integer) As Boolean
Select Case ToolIndex
Case ToolFade, ToolFStar, ToolVFade, ToolHFade
IsAutoColorTool = True
Case Else
IsAutoColorTool = False
End Select
End Function

Private Sub ActiveColor_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then CancelActiveColorModeChange = False
End Sub

Private Sub ActiveColor_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
Dim PP  As POINTAPI
If Button = 1 And Not (InButton(ActiveColor(Index), x, y, False, True)) Then
    ActiveColor(Index).Drag
ElseIf Button = 2 Then
    If InButton(ActiveColor(Index), x, y, False, True) Then
        FreshActiveColors
    Else
        CancelActiveColorModeChange = True
        GetCursorPos PP
        ActiveColor(Index).BackColor = ConvertColorLng(CapturePixel(PP.x, PP.y))
    End If
End If
End Sub

Private Sub ActiveColor_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
Dim tmp As String, rgb1 As RGBQUAD
Dim ModeSet As Boolean
Dim PP As POINTAPI
If (Button = 2) Then
    
    If InButton(ActiveColor(Index), x, y, False, True) Then
        If Not CancelActiveColorModeChange Then
            If Index = ActiveColor.lBound Then
                gFDSC.AutoColor1 = Not gFDSC.AutoColor1
                ModeSet = gFDSC.AutoColor1
            Else
                gFDSC.AutoColor2 = Not gFDSC.AutoColor2
                ModeSet = gFDSC.AutoColor2
            End If
        End If
    Else
        GetCursorPos PP
        ACol(Index) = CapturePixel(PP.x, PP.y)
    End If
    FreshActiveColors
    GetRgbQuadEx ActiveColor(Index).BackColor, rgb1
    ShowStatus grs(10072, "%r", CStr(rgb1.rgbRed), _
                      "%g", CStr(rgb1.rgbGreen), _
                      "%b", CStr(rgb1.rgbBlue), _
                      "$h", Hex$(ActiveColor(Index).BackColor), _
                      "#a", BoolToStr_OnOff(gFDSC.AutoColor1)), , 2
End If
End Sub

Private Sub btnTool_Click(Index As Integer)
Dim i As Integer
If btnTool(Index).Tag <> "" Then Exit Sub
If Not Index = btnTool.UBound Then
    mnuTool_Click Index
Else
    mnuToolOpts_Click
    
    btnTool(Index).Tag = "Using"
    btnTool(Index).Value = False
    btnTool(Index).Tag = ""
End If
On Error Resume Next
MP.SetFocus
On Error GoTo 0
End Sub

Private Sub btnTool_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case 37 To 40
        MP.SetFocus
        Form_KeyDown KeyCode, Shift
End Select
End Sub

Private Sub btnTool_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then
    If Not (Index = btnTool.UBound) Then
        If Val(mnuTool(Index).Tag) <> ActiveTool Then
            btnTool_Click Index
        End If
        mnuToolOpts_Click
    Else
        mnuToolOpts_Click
    End If
End If
End Sub

Private Sub ChCol_DragDrop(Index As Integer, Source As Control, x As Single, y As Single)
If UCase(Source.Name) = UCase("ActiveColor") Then
    ChColBackColor Index, ACol(Source.Index)
End If
End Sub

Private Sub ChCol_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 1 Or Button = 2 Then
    ChangeActiveColor Button, ChCol(Index).BackColor, setPer:=False
    ShowStatus grs(1238, "|1", CStr(Index + 1), _
            "%r", CStr(GetAttr(ChCol(Index).BackColor, 1)), _
            "%g", CStr(GetAttr(ChCol(Index).BackColor, 2)), _
            "%b", CStr(GetAttr(ChCol(Index).BackColor, 3)), _
            "$h", Hex$(ChCol(Index).BackColor))             '"Palette entry # " + CStr(Index)
End If
End Sub

Private Sub ChCol_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 1 Or Button = 2 Then
    ShowStatus grs(1238, "|1", CStr(Index + 1), _
            "%r", CStr(GetAttr(ChCol(Index).BackColor, 1)), _
            "%g", CStr(GetAttr(ChCol(Index).BackColor, 2)), _
            "%b", CStr(GetAttr(ChCol(Index).BackColor, 3)), _
            "$h", Hex$(ChCol(Index).BackColor))             '"Palette entry # " + CStr(Index)
End If
If Len(ChCol(Index).Tip) > 0 Then
    frmColors.ToolTipText = ChCol(Index).Tip
Else
    frmColors.ToolTipText = GRSF(241)
End If
End Sub

Private Sub ChCol_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
If (Button = 1 Or Button = 2) And (x >= 0 And y >= 0 And x < ChCol(Index).Width * Screen.TwipsPerPixelX And y < ChCol(Index).Height * Screen.TwipsPerPixelY) Then
    If (Shift And &H1) = &H1 Then ChChCol Index
    If (Shift And &H2) = &H2 Then
        On Error Resume Next
        Err.Clear
        ChCol(Index).Tip = dbInputBox(grs(1190, "%d", GenerateColorTip(ChCol(Index).BackColor)), ChCol(Index).Tip, True)
        If Err.Number > 0 Then
            ChCol(Index).Tip = ""
        End If
        On Error GoTo 0
    Else
        If (Shift And &H1) = &H1 Then ChCol(Index).Tip = ""
    End If
    ChangeActiveColor Button, ChCol(Index).BackColor
    On Error Resume Next
    MP.SetFocus
    On Error GoTo 0
    LastPalIndex = Index
End If
End Sub

Sub ChChCol(ByVal Index As Integer)
Attribute ChChCol.VB_Description = "Chenges palette item color"
    On Error GoTo eh
    With CDl
        ChCol(Index).BackColor = .PickColor(ChCol(Index).BackColor)
        DrawPalEntry Index
    End With
    On Error GoTo 0
Exit Sub
eh:
If Err.Number = dbCWS Then Exit Sub
MsgError
End Sub

Private Sub Form_Activate()
On Error Resume Next
'MP.SetFocus
'Debug.Print "mai form_activate" + Rnd(1)
End Sub

Private Sub Form_DblClick()
FormMS.DblClick
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
FormMS.MouseDown Button, Shift, x, y
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
FormMS.MouseMove Button, Shift, x, y
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
FormMS.MouseUp Button, Shift, x, y
End Sub

Private Sub FormMS_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 1 And _
   x >= VScroll.Left And x < VScroll.Left + VScroll.Width And _
   y >= HScroll.Top And y < HScroll.Top + HScroll.Height Then
    
  Dim FC As Long, BC As Long
  FC = ACol(1)
  BC = ACol(2)
  ChangeActiveColor 1, BC, False
  ChangeActiveColor 2, FC, False
End If
End Sub

Private Sub mnuDrawAFM_Click()
frmDrawAFM.Show , Me
End Sub

Private Sub mnuDrawTopoR_Click()
On Error GoTo eh
Load frmTopoR
frmTopoR.Browse
frmTopoR.Show , Me
Exit Sub
eh:
MsgError
End Sub

Private Sub mnuDrawWaves_Click()
Dim WS() As typWaveSource
Dim Mode As eWPMode
Dim FDFactor As Double, Pow As Double
Dim bAbs As Boolean
On Error GoTo eh
Load frmWavesCreate
With frmWavesCreate
    UpdateWH
    .ImageW = intW
    .ImageH = intH
    .Form_Resize
    .AutoScale
    .Show vbModal
    If .Tag <> "" Then
        Unload frmWavesCreate
        Err.Raise dbCWS
    End If
    .GetSources WS
    Mode = .Mode
    FDFactor = .mFallDownFactor
    Pow = .mFieldPower
    bAbs = .mnuDoneWavesOptsAbsolute.Checked
End With
Unload frmWavesCreate
BUD
Select Case Mode
    Case eWPMode.wpmFieldLines
        mdlWaves.DrawELines Data, WS, Pow
    Case eWPMode.wpmWavesPicture
        mdlWaves.DrawWaves Data, WS, FDFactor, bAbs
End Select
Refr

Exit Sub
Resume
eh:
MsgError
End Sub

Private Sub SaveSelTest()
If SelectionPresent Then
    Select Case dbMsgBox(2631, vbYesNoCancel)
        Case vbYes
            dbDeselect True
        Case vbNo
            'do nothing
        Case vbCancel
            Err.Raise dbCWS
    End Select
End If
End Sub

Private Sub mnuFileImportAscii_Click()
Dim FN As String
On Error GoTo eh
  MsgBox ("The ascii file must be a single column of data. The data will be put to the picture left to right, filling the rows from top to bottom.")
  FN = ShowOpenDlg(dbAsciiLoad, Me.hWnd, Purpose:="LOG")
  UpdateWH
  Dim tf As Long
  tf = FreeFile
  Open FN For Input As tf
    Dim curX As Long, curY As Long
    
    Do Until EOF(tf)
      Dim tmp As String
      Line Input #(tf), tmp
      If Len(tmp) > 0 Then
        On Error GoTo BadVal
        Dim Bri As Byte
        Bri = dbVal(tmp, vbByte)
        On Error GoTo eh
        If curX = 0 And curY = 0 Then StartPixelAction
        dbPSet curX, curY, CLng(Bri) * &H10101, StoreToData:=True, ForceDraw:=True
        curX = curX + 1
        If curX = intW Then
          curX = 0
          curY = curY + 1
          If curY = intW Then Err.Raise 12345, "ImportASCII", "Not enough space in the image!"
        End If
      End If
rsm:
      On Error GoTo eh
    Loop
  Close tf
  Refr
Exit Sub
eh:
  PushError
  If tf <> 0 Then Close tf
  PopError
  MsgError
Exit Sub
BadVal:
Resume rsm

End Sub

Private Sub mnuOptRememberWndPos_Click()
mnuOptRememberWndPos.Checked = Not mnuOptRememberWndPos.Checked
dbSaveSettingEx Me.Name, "RememberPos", mnuOptRememberWndPos.Checked
End Sub

Private Sub mnuSaveAs_Click()
On Error GoTo eh
'vtSavePicture Data, DataAlpha, OpenedFileName, OpenedFileFormatID, ShowDialog:=True, Purpose:="MP"
SaveAuto ShowDialog:=True
FreshCaption
Exit Sub
eh:
MsgError
End Sub

Private Sub mnuSelAutoRepaint_Click()
SelPictureAutoRepaint = Not (SelPictureAutoRepaint)
mnuSelAutoRepaint.Checked = SelPictureAutoRepaint
End Sub

Private Sub mnuSelCenterHorz_Click()
If Not SelectionPresent Then Exit Sub
UpdateWH
MoveSel (intW - GetSelW) \ 2, CurSel.y1
End Sub

Private Sub mnuSelCenterVert_Click()
If Not SelectionPresent Then Exit Sub
UpdateWH
MoveSel CurSel.x1, (intH - GetSelH) \ 2
End Sub

Private Sub mnuSelClear_Click()
If Not SelectionPresent Then Exit Sub
ClearPic CurSel_SelData, ACol(2)
dbPutSel
End Sub

Private Sub mnuSelEditSel_Click()
If mnuEditSel.Enabled Then mnuEditSel_Click
End Sub

Private Sub mnuSelMoveTo_Click()
Dim Al As New clsAligner
Dim Point As POINTAPI
On Error GoTo eh
Al.DestPointSupported = False
Al.BasePointSupported = False
Al.LoadFromReg "Options", "SelMoveToOpts"
Al.Customize RaiseErrors:=True
UpdateWH
Point = Al.GetOffset(intW, intH, SelW, SelH)
MoveSel Point.x, Point.y
Al.SaveToReg "Options", "SelMoveToOpts"
Exit Sub
eh:
If Err.Number = dbCWS Then Exit Sub
MsgError
End Sub

Private Sub mnuselDelete_Click()
dbDeselect False
End Sub

Private Sub mnuSelResize_Click()
Dim Sz As Dims
If Not SelectionPresent Then Exit Sub
Load Dialog
With Dialog
    Sz.w = GetSelW
    Sz.h = GetSelH
    .SetSz Sz
    .Show vbModal
    If Len(.Tag) = 0 Then
        .ExtractSz Sz
        dbStretch CurSel_SelData, Sz.w, Sz.h, .GetStretchMode
        dbPutSel
    End If
End With
Unload Dialog
End Sub

Private Sub mnuSelShow_Click()
Dim x As Long, y As Long
Dim w As Long, h As Long
UpdateWH
w = SelW
h = SelH
x = ((MPHolder.ScaleWidth - GetSelW * Zm) \ 2 - MP.Left) \ Zm
y = ((MPHolder.ScaleHeight - GetSelH * Zm) \ 2 - MP.Top) \ Zm
x = Min(x, intW - w)
y = Min(y, intH - h)
x = Max(x, 0)
y = Max(y, 0)
MoveSel x, y
End Sub

Public Function GetSelW() As Long
UpdateWH
GetSelW = SelW
End Function

Public Function GetSelH() As Long
UpdateWH
GetSelH = SelH
End Function

Public Sub MoveSel(ByVal x As Long, ByVal y As Long)
Dim SelW As Long, SelH As Long
If Not CurSel.Selected Then Exit Sub
SelW = GetSelW
SelH = GetSelH
CurSel.x1 = x
CurSel.y1 = y
CurSel.x2 = CurSel.x1 + SelW - 1
CurSel.y2 = CurSel.y1 + SelH - 1
UpdateSelPic RedrawSelPic:=True, SetToData:=False, StoreUndo:=False
End Sub

Private Sub mnuSelShowSize_Click()
UpdateWH
ShowStatus grs(1253, "%w%", CStr(SelW), _
                     "%h%", CStr(SelH), _
                     "%x%", CStr(CurSel.x1), _
                     "%y%", CStr(CurSel.y1)), HoldTime:=5
FlashStatusBar
End Sub

Private Sub mnuSelStamp_Click()
If Not SelectionPresent Then Exit Sub
UpdateSelPic RedrawSelPic:=False, SetToData:=True
End Sub

Private Sub mnuShowSplash_Click()
mnuShowSplash.Checked = Not mnuShowSplash.Checked
dbSaveSettingEx "Options", "ShowSplash", mnuShowSplash.Checked
End Sub

Private Sub mnuUnInstall_Click()
Uninstall
End Sub

Private Sub MP_DblClick()
MPMS.DblClick
End Sub

Private Sub MP_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub MP_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
MPMS.MouseDown Button, Shift, x, y
End Sub

Private Sub MP_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
MPMS.MouseMove Button, Shift, x, y
End Sub

Private Sub MP_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
MPMS.MouseUp Button, Shift, x, y
End Sub

Private Sub MPHolder_DblClick()
If WSMS(1) Then
    MPHolder_MouseDown 2, 0, 0, 0
ElseIf WSMS(2) Then
    MPHolder_MouseDown 1, 0, 0, 0
Else
    mnuResize_Click
    Erase WSMS
End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Long, j As Long, a As POINTAPI, t As Long, Ed As Integer
Dim ShiftKeyCode As Integer
Dim ActsAry() As Long
Dim nActs As Long
Dim DontBeep As Boolean
Const ShiftMask = &H100
Const CtrlMask = &H200
Const AltMask = &H400

If KeyCode = 0 Then Exit Sub

UserMadeAction

If Not (MeEnabled) Then Exit Sub
On Error GoTo eh
ShiftKeyCode = &H100 * Shift + KeyCode
If ShiftKeyCode = 112 Then
    ToggleHelpWindow
    KeyCode = 0
End If
nActs = Keyb.GetActs(ShiftKeyCode, ActsAry)
For i = 0 To nActs - 1
    Select Case ActsAry(i)
        Case dbCommands.cmdLMB, dbCommands.cmdMMB, dbCommands.cmdRMB
            ExecuteCmd ActsAry(i), Shift, KeyDown
        Case dbCommands.cmdCapturePoint, _
             dbCommands.cmdoClearCharges, _
             dbCommands.cmdCursorDown, _
             dbCommands.cmdCursorLeft, _
             dbCommands.cmdCursorRight, _
             dbCommands.cmdCursorUp, _
             dbCommands.cmdCyclePoly, _
             dbCommands.cmdEndPoly, _
             dbCommands.cmdExtremeSave, _
             dbCommands.cmdIdleMessage, _
             dbCommands.cmdNone, _
             dbCommands.cmdPaintBrush, _
             dbCommands.cmdPalVisible, _
             dbCommands.cmdScrollDown, _
             dbCommands.cmdScrollLeft, _
             dbCommands.cmdScrollRight, _
             dbCommands.cmdScrollRight, _
             dbCommands.cmdToolBarVisible, _
             dbCommands.cmdToggleGrid, _
             dbCommands.cmdNaviMode, _
             dbCommands.cmdWheelUp, _
             dbCommands.cmdWheelDown
            ExecuteCmd ActsAry(i), Shift, KeyDown
            DontBeep = True
                
        Case Else
            If dbMS(1, 1) Or dbMS(2, 1) Or dbMS(3, 1) Then
                ShowStatus 2441, , 3
                If Not DontBeep Then vtBeep
            Else
                ExecuteCmd ActsAry(i), Shift, KeyDown
            End If
    End Select
Next i
Select Case KeyCode
  Case 16, 17, 18 'shift,ctrl,alt
    If ActiveTool = ToolProg Then
      MoveMouse
    End If
End Select

Exit Sub
Resume
eh:
If Err.Number = dbCWS Then
    ShowStatus STT_Cancelled, , 1
    Exit Sub
Else
    MsgError
End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
Dim i As Integer, a As POINTAPI, ShiftKeyCode As Integer
Dim ActsAry() As Long
Dim nActs As Long
Const ShiftMask = &H100
Const CtrlMask = &H200
Const AltMask = &H400
ShiftKeyCode = Shift * &H100 + KeyCode
nActs = Keyb.GetActs(ShiftKeyCode, ActsAry)
For i = 0 To nActs - 1
    Select Case ActsAry(i)
        Case dbCommands.cmdLMB, _
             dbCommands.cmdMMB, _
             dbCommands.cmdRMB, _
             dbCommands.cmdNaviMode
            ExecuteCmd ActsAry(i), Shift, KeyUp
    End Select
Next i
Select Case KeyCode
  Case 16, 17, 18 'shift,ctrl,alt
    If ActiveTool = ToolProg Then
      MoveMouse
    End If
End Select
End Sub

Private Sub Form_Load()
Set MPMS = New clsAntiDblClick
Set FormMS = New clsAntiDblClick

ShowHelp Me.HelpContextID
SetClassLong MP.hWnd, GCL_HBRBACKGROUND, 0
WinAPI.ModifyWindowStyle MP.hWnd, 0, StyleUnset:=CS_DBLCLKS
MainModule.SubClassMP
ConnectEffects
End Sub

Private Sub MPHolder_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 And WSMS(1) Then
    If mnuRedo.Enabled Then
        mnuRedo_Click
    Else
        vtBeep
    End If
ElseIf Button = 1 And WSMS(2) Then
    If mnuUndo.Enabled Then
        mnuUnDo_Click
    Else
        vtBeep
    End If
End If
WSMS(Button) = True
End Sub

Private Sub MPHolder_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
WSMS(Button) = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub frmColors_DragDrop(Source As Control, x As Single, y As Single)
Dim Index As Integer
If x >= 0 And y >= 0 And x < frmColors.ScaleWidth And y < frmColors.ScaleHeight Then
Index = GetChColIndex(x, y)
If Not (Index > UBound(ChCol)) Then
    ChCol_DragDrop Index, Source, x Mod ChCol(0).Width, y Mod ChCol(0).Height
End If
End If
End Sub

Private Sub frmColors_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim Index As Integer
Dim tx As Integer, ty As Integer
On Error Resume Next
If x >= 0 And y >= 0 And x < frmColors.ScaleWidth And y < frmColors.ScaleHeight Then
Index = GetChColIndex(x, y)
If Not (Index > UBound(ChCol)) Then
    tx = x Mod ChCol(0).Width
    ty = y Mod ChCol(0).Height
    ChCol_MouseDown Index, Button, Shift, (tx), (ty)
End If
End If

End Sub

Private Sub frmColors_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim Index As Integer
Dim tx As Integer, ty As Integer
On Error Resume Next
If x >= 0 And y >= 0 And x < frmColors.ScaleWidth And y < frmColors.ScaleHeight Then
Index = GetChColIndex(x, y)
If Not (Index > UBound(ChCol)) Then
    tx = x Mod ChCol(0).Width
    ty = y Mod ChCol(0).Height
    ChCol_MouseMove Index, Button, Shift, (tx), (ty)
End If
End If
End Sub

Private Sub frmColors_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim Index As Integer
Dim tx As Integer, ty As Integer
On Error Resume Next
If x >= 0 And y >= 0 And x < frmColors.ScaleWidth And y < frmColors.ScaleHeight Then
Index = GetChColIndex(x, y)
If Not (Index > UBound(ChCol)) Then
    tx = x Mod ChCol(0).Width
    ty = y Mod ChCol(0).Height
    ChCol_MouseUp Index, Button, Shift, (tx), (ty)
End If
End If
End Sub

Private Sub frmColors_Paint()
Dim i As Integer
For i = 0 To UBound(ChCol)
    DrawPalEntry i
Next i
End Sub

Private Sub mnuAbout_Click()
frmAbout.Show vbModal
End Sub

Friend Sub mnuBuildBackUp_Click()
On Error GoTo eh
BuildBackup
Exit Sub
eh:
MsgBox "Backup building failed." + vbNewLine + Err.Description, vbCritical, Err.Source
End Sub

Private Sub mnuCapture_Click()
Dim tData() As Long
On Error GoTo eh
Load frmCapture
With frmCapture
    .Show vbModal
    If Len(.Tag) = 0 Then
        dbCapture .GetMode, tData
        dbMakeSel 0, 0, UBound(tData, 1) + 1, UBound(tData, 2) + 1
        CurSel_SelData = tData
        Erase tData
        dbPutSel
    End If
End With
Unload frmCapture
Exit Sub
eh:
If Err.Number = dbCWS Then
    ShowStatus STT_Cancelled, , 3
    Exit Sub
End If
MsgError
End Sub

Private Sub mnuClear2_Click()
mnuClear_Click
End Sub

Private Sub mnuClearUndo_Click()
ShowStatus GRSF(STT_Working)
ClearUndo
ShowStatus GRSF(10006)
dbMsgBox GRSF(1145), vbInformation ' "Undo/Redo history was cleared successfully"
ShowStatus GRSF(STT_READY)
End Sub

Private Sub mnuCopy_Click()
On Error GoTo eh
If CurSel.Selected Then
    With TempBox
    Status.Caption = GRSF(1200)
    Status.Refresh
    .ZOrder
    .Visible = True
    .AutoSize = False
    .Picture = LoadPicture("")
    .Cls
    .Width = Abs(CurSel.x1 - CurSel.x2) + 1
    .Height = Abs(CurSel.y1 - CurSel.y2) + 1
    .Refresh
    RefrEx .Image.Handle, .hDC, CurSel_SelData, 1, dbNoGrid
    .Refresh
    Clipboard.Clear
    Clipboard.SetData .Image, VBRUN.ClipBoardConstants.vbCFBitmap
    .Visible = False
    Status.Caption = GRSF(1204)
    End With
Else
  Dim ret As VbMsgBoxResult
  ret = dbMsgBox(GRSF(1104), vbInformation Or vbOKCancel)
  If ret = vbCancel Then
    Err.Raise dbCWS
  End If
  With TempBox
    Status.Caption = GRSF(1200)
    Status.Refresh
    .ZOrder
    .Visible = True
    .AutoSize = False
    .Picture = LoadPicture("")
    UpdateWH
    .Width = intW 'Abs(CurSel.x1 - CurSel.x2) + 1
    .Height = intH 'Abs(CurSel.y1 - CurSel.y2) + 1
    .Cls
    .Refresh
    RefrEx .Image.Handle, .hDC, Data, 1, dbNoGrid
    .Refresh
    Clipboard.Clear
    Clipboard.SetData .Image, VBRUN.ClipBoardConstants.vbCFBitmap
    .Height = 32
    .Width = 32
    .Cls
    .Visible = False
    Status.Caption = GRSF(1204)
  End With
End If
Exit Sub
eh:
MsgError
End Sub

Private Sub mnuCrop_Click()
If Not CurSel.Selected Then Exit Sub
BUD
Resize UBound(CurSel_SelData, 1) + 1, UBound(CurSel_SelData, 2) + 1, False, False, False
Data = CurSel_SelData
dbDeselect False
Refr
End Sub

Private Sub mnuDrawBBg_Click()
Dim x As Single, y As Single, Fi As Double
Dim lx As Single, ly As Single
Dim k1 As Double, k2 As Double
Dim k3 As Double, k4 As Double
Dim k5 As Double, k6 As Double
Dim k7 As Double, k8 As Double
Dim k9 As Double, kA As Double
ShowStatus STT_BUD
StartPixelAction
ShowStatus STT_Working
Randomize Timer
k1 = Int(Rnd(1) * 10)
k2 = Int(Rnd(1) * 10)
k3 = Int(Rnd(1) * 10)
k4 = Int(Rnd(1) * 10)
k5 = Int(Rnd(1) * 10)
k6 = Int(Rnd(1) * 10)
k7 = Int(Rnd(1) * 10)
k8 = Int(Rnd(1) * 10)
k9 = Int(Rnd(1) * 10)
kA = Int(Rnd(1) * 10)
On Error Resume Next
MeEnabled = False
For Fi = 0 To 2 * Pi * 10 Step 0.0005
    x = intW * 0.5 + Cos(Fi * 0.1) * ((Cos((Fi) * k1) ^ k3 + Sin(Fi * k2) ^ k4) * 100 + 200)
    y = intH * 0.5 + Sin(Fi * 0.1) * ((Cos((Fi) * k5) ^ k7 + Sin(Fi * k6) ^ k8) * 100 + 200)
    If (x - lx) ^ 2 + Abs(y - ly) ^ 2 > 0.25 Then
        dbPutPoint x, y, RGB(Sin(Fi * k5) * 128 + 128, (Cos(Fi * k6) + Sin(Fi * k5)) * 128 * 0.5 + 128, Cos(2 * Fi) ^ 2 * 255)
        lx = x
        ly = y
    End If
    ShowProgress Fi / (2 * Pi * 10) * 100, True
Next Fi
MeEnabled = True
Refr
FileChanged = True
End Sub

Private Sub mnuDefPal16_Click()
Dim DefPal() As Long, i As Long
Const pC As Long = 16
ReDim DefPal(0 To pC - 1)
LoadDefPal DefPal
ChangePalCount pC
For i = 0 To pC - 1
    ChCol(i).BackColor = DefPal(i)
    ChCol(i).Tip = CStr(i)
Next i
frmColors.Refresh
End Sub

Private Sub mnuDefPal256_Click()
Dim DefPal() As Long, i As Long
Const pC As Long = 256
ReDim DefPal(0 To pC - 1)
LoadDefPal DefPal
ChangePalCount pC
For i = 0 To pC - 1
    ChCol(i).BackColor = DefPal(i)
    ChCol(i).Tip = CStr(i)
Next i
frmColors.Refresh
End Sub

Private Sub mnuDrawBG_Click()
Dim i As Long, j As Long
BUD
For i = 0 To UBound(Data, 2)
    For j = 0 To UBound(Data, 1)
        Data(j, i) = ACol(2)
    Next j
Next i
Refr
FileChanged = True
End Sub

Private Sub mnuDrawPlain_Click()
Dim i As Single, j As Single
BUD
For i = 1 To intH Step 0.125
    For j = 0 To intW - 1 Step CSng(i) / 2
        Data(Round(j), i - 1) = ACol(1)
        Data(intW - 1 - Round(j), i - 1) = ACol(1)
    Next j
Next i
Refr
FileChanged = True
End Sub

Private Sub mnuDynamicScr_Click()
Dim b As Boolean
Load frmDynSettings
With frmDynSettings
    .SetProps ScrollSettings, MoveTimerRes
    .Show vbModal
    If .Tag = "" Then
        .GetProps ScrollSettings, MoveTimerRes
        ChangeDynScrolling ScrollSettings.DS_Enabled
    End If
End With
Unload frmDynSettings
MoveMP , , &H8, True
End Sub

Private Sub mnuEditSel_Click()
If Not CurSel.Selected Then
    vtBeep
    Exit Sub
End If
ViewImage CurSel_SelData, "Sel"
If AryDims(AryPtr(CurSel_SelData)) <> 2 Then
    CurSel.Selected = False
    SelPicture.Visible = False
Else
    CurSel.x2 = CurSel.x1 + UBound(CurSel_SelData)
    CurSel.y2 = CurSel.y1 + UBound(CurSel_SelData)
End If
dbPutSel
End Sub

Private Sub mnuEffect_Click(Index As Integer)
On Error GoTo eh
dbEffect Index, NoDialog:=False
Exit Sub
eh:
MsgError
End Sub

Public Sub dbEffect(ByVal Index As Integer, Optional ByVal NoDialog As Boolean)
Static bDot() As Byte
Static Mask As FilterMask
Dim i As Integer, j As Integer
Dim mnoj As Single
Dim Answ As VbMsgBoxResult
Dim intR As Integer, intG As Integer, intB As Integer
Dim Matrix() As Double
Dim xyDir As POINTAPI
Static rTbl() As Byte, gTbl() As Byte, bTbl() As Byte
Dim NegMask As Long
Dim DataOrig() As Long
Static LastMCH As Integer
Dim EfID As String


Dim tDataRGB() As RGBQUAD, tDataRGBL() As RGBTriLong
Dim Loads As RGBTriLong
Dim Anti As Boolean
Dim NeedEnable As Boolean

On Error GoTo eh

NeedEnable = DontDoEvents
If NeedEnable Then
    DisableMe
End If

Select Case Val(mnuEffect(Index).Tag)
    Case 0 'soft
        ShowStatus "$10019", , 3
        EfID = "Filtering"
    Case 1 'chzch
        If CurSel.Selected Then
            MonochromizeSimple CurSel_SelData, 0, 0, UBound(CurSel_SelData, 1), UBound(CurSel_SelData, 2)
            dbPutSel
        Else
            Status.Caption = GRSF(1202)
            Status.Refresh
            BUD
            Status.Caption = GRSF(1203)
            Status.Refresh
            MonochromizeSimple Data, 0, 0, intW - 1, intH - 1
            FileChanged = True
            Refr
        End If
    Case 2 'negative
        ShowStatus "$10020"
        
        If Not NoDialog Then
            NegDialog.Show vbModal
            If NegDialog.Tag <> "" Then
                ShowStatus STT_Cancelled
                Exit Sub
            End If
        End If
        NegMask = NegDialog.GetMask
        If CurSel.Selected Then
            ShowStatus GRSF(STT_Processing)
            dbNegative CurSel_SelData, 0, 0, UBound(CurSel_SelData, 1), UBound(CurSel_SelData, 2), NegMask
            ShowStatus GRSF(STT_Displaying)
            dbPutSel
            ShowStatus GRSF(STT_READY)
        Else
            ShowStatus GRSF(STT_BUD)
            BUD
            ShowStatus GRSF(STT_Processing)
            dbNegative Data, 0, 0, intW - 1, intH - 1, NegMask
            FileChanged = True
            Refr
        End If
        ShowStatus GRSF(STT_Cancelled)
    Case 3 'gamma
        EfID = "Gamma"
    Case 4 'turn
        ShowStatus "$10022"
        If CurSel.Selected Then
            If Not NoDialog Then
                frmTurn.Show vbModal
                If Not (frmTurn.Tag = "") Then
                    ShowStatus GRSF(STT_Cancelled)
                    Exit Sub
                End If
            End If
            i = frmTurn.Method
            ShowStatus GRSF(STT_Processing)
            dbTurn CurSel_SelData, i
            If AryDims(AryPtr(TransOrigData)) = 2 Then
              dbTurn TransOrigData, i
            End If
            If AryDims(AryPtr(TransData)) = 2 Then
              dbTurn TransData, i
            End If
            
            ShowStatus GRSF(STT_Resizing)
            CurSel.x2 = CurSel.x1 + UBound(CurSel_SelData, 1)
            CurSel.y2 = CurSel.y1 + UBound(CurSel_SelData, 2)
            With SelRect
                .Width = (UBound(CurSel_SelData, 1) + 1) * Zm
                .Height = (UBound(CurSel_SelData, 2) + 1) * Zm
            End With
            With SelPicture
                .Width = (UBound(CurSel_SelData, 1) + 1) * Zm
                .Height = (UBound(CurSel_SelData, 2) + 1) * Zm
            End With
            ShowStatus GRSF(STT_Displaying)
            dbPutSel
            ShowStatus GRSF(STT_READY)
        Else
            If Not NoDialog Then
                frmTurn.Show vbModal
                If Not (frmTurn.Tag = "") Then
                    ShowStatus GRSF(STT_Cancelled)
                    Exit Sub
                End If
            End If
            i = frmTurn.Method
            ShowStatus GRSF(STT_BUD)
            BUD
            ShowStatus GRSF(STT_Processing)
            dbTurn Data, i
            Refr
        End If
    Case 5 'DeColour
        ShowStatus "$10023"
        If CurSel.Selected Then
            If Not NoDialog Then
                frmDeColour.Show vbModal
                If Not frmDeColour.Tag = "" Then
                    ShowStatus GRSF(STT_Cancelled)
                    Exit Sub
                End If
            End If
            mnoj = Val(frmDeColour.Text.Text) / 100
            ShowStatus GRSF(STT_Processing)
            dbDeColour CurSel_SelData, 0, 0, UBound(CurSel_SelData, 1), UBound(CurSel_SelData, 2), mnoj
            ShowStatus GRSF(STT_Displaying)
            dbPutSel
            ShowStatus GRSF(STT_READY)
        Else
            If Not NoDialog Then
                frmDeColour.Show vbModal
                If Not frmDeColour.Tag = "" Then
                    ShowStatus GRSF(STT_Cancelled)
                    Exit Sub
                End If
            End If
            mnoj = Val(frmDeColour.Text.Text) / 100
            ShowStatus GRSF(STT_BUD)
            BUD
            ShowStatus GRSF(STT_Processing)
            dbDeColour Data, 0, 0, intW - 1, intH - 1, mnoj
            FileChanged = True
            Refr
        End If
    Case 6 'ReRGB
        ShowStatus 10024
        EfID = "ReRGB"
        
    Case 7 'make monochrome
        ShowStatus "$10025"
        If CurSel.Selected Then
            If Not NoDialog Then
                frmMono.Show vbModal, MainForm
                If frmMono.Tag = "" Then
                    ShowStatus GRSF(STT_Cancelled)
                    Exit Sub
                End If
            End If
            ShowStatus GRSF(STT_Processing)
            If Not NoDialog Then
                If Val(frmMono.Tag) = 1 Then
                    dbMakeMono CurSel_SelData
                    LastMCH = 1
                ElseIf Val(frmMono.Tag) = 0 Then
                    dbMakeMonoRND CurSel_SelData
                    LastMCH = 0
                End If
            Else
                If LastMCH = 1 Then
                    dbMakeMono CurSel_SelData
                    LastMCH = 1
                ElseIf LastMCH = 0 Then
                    dbMakeMonoRND CurSel_SelData
                    LastMCH = 0
                End If
            End If
            ShowStatus GRSF(STT_Displaying)
            dbPutSel
            ShowStatus GRSF(STT_READY)
        Else
            If Not NoDialog Then
                frmMono.Show vbModal, MainForm
                If frmMono.Tag = "" Then
                    ShowStatus GRSF(STT_Cancelled)
                    Exit Sub
                End If
            End If
            ShowStatus GRSF(STT_BUD)
            BUD
            ShowStatus GRSF(STT_Processing)
            If Not NoDialog Then
                If Val(frmMono.Tag) = 1 Then
                    dbMakeMono Data
                    LastMCH = 1
                ElseIf Val(frmMono.Tag) = 0 Then
                    dbMakeMonoRND Data
                    LastMCH = 0
                End If
            Else
                If LastMCH = 1 Then
                    dbMakeMono Data
                    LastMCH = 1
                ElseIf LastMCH = 0 Then
                    dbMakeMonoRND Data
                    LastMCH = 0
                End If
            End If
            FileChanged = True
            Refr
        End If
    Case 8 'IO Graph
        EfID = "Graph"
        
        
    Case 9 'Replace
        Dim ColorFind As Long, Sens As Integer, ColorReplace As Long
        With frmColorReplace
            If Not NoDialog Then
                .Show vbModal
                If .Tag <> "" Then
                    ShowStatus STT_Cancelled
                    Exit Sub
                End If
            End If
            ColorFind = .clrFind.Color
            ColorReplace = .clrReplace.Color
            Sens = CInt(.txtSens.Text)
        End With
        If CurSel.Selected Then
            dbReplaceColors CurSel_SelData, ColorFind, Sens, ColorReplace
            dbPutSel
        Else
            BUD
            dbReplaceColors Data, ColorFind, Sens, ColorReplace
            Refr
        End If
    Case 10 'Differenciate
        EfID = "Diff"
        
    Case 11 'ClearType
        On Error GoTo eh
        
        Load frmClearType
        With frmClearType
            .LoadSettings
            .Tag = ""
            If Not NoDialog Then
                .Show vbModal
            End If
            If Len(.Tag) > 0 Then Err.Raise dbCWS, "dbEffect", "Cancel Was Selected"
            Anti = .Anti
        End With
        Unload frmClearType
        
        MeEnabled = False
        SetMousePtr True
        If CurSel.Selected Then
            vtClearType CurSel_SelData, TexMode, Anti
            dbPutSel
        Else
            BUD
            vtClearType Data, TexMode, Anti
            Refr
        End If
        

End Select
If Len(EfID) > 0 Then
    If CurSel.Selected Then
        PerformEffectEx GetEffectByID(EfID), CurSel_SelData, DataOrig, ShowDialog:=Not NoDialog
        SwapArys AryPtr(CurSel_SelData), AryPtr(DataOrig)
        dbPutSel
    Else
        PerformEffectEx GetEffectByID(EfID), Data, DataOrig, ShowDialog:=Not NoDialog
        SwapArys AryPtr(Data), AryPtr(DataOrig)
        BUD AryPtr(DataOrig)
        Refr
    End If
End If
ExitHere:
LastEffectIndex = Index 'Val(mnuEffect(Index).Tag)
mnuLastEffect.Enabled = True
FileChanged = True
If NeedEnable Then RestoreMeEnabled
Exit Sub
Resume
eh:
If NeedEnable Then
    PushError
    ClearMeEnabledStack
    PopError
End If
ErrRaise "dbEffect"
End Sub

Private Sub mnuEmptyTips_Click()
Dim i As Integer
For i = 0 To UBound(ChCol)
    ChCol(i).Tip = vbNullString
Next i
End Sub

Private Sub mnuFile_Click()
UserMadeAction
End Sub

Private Sub mnuFillTips_Click()
Dim i As Integer, clr As Long
For i = 0 To UBound(ChCol)
    ChCol(i).Tip = GenerateColorTip(ChCol(i).BackColor)
Next i
End Sub

Private Sub mnuFolder_Click()
If OpenedFileName <> "" Then
    ShowFile OpenedFileName
End If
End Sub

Public Sub ExploreFolder(ByRef strFolder As String)
Shell "Explorer.exe """ + strFolder + """", vbNormalFocus
End Sub

Public Sub ShowFile(ByRef strFile As String)
Shell "Explorer.exe /select, """ + strFile + """", vbNormalFocus
End Sub

Private Sub mnuFormula_Click()
Dim x As Long, y As Long
Dim AC As Long
On Error GoTo eh
Load frmInsOLE
With frmInsOLE
    .Timer1.Enabled = True
    .Show vbModal
    If .Tag = "" Then
        'mnuPaste_Click
        If CurSel.Selected Then dbDeselect True
        dbMakeSel 0, 0, TempBox.Width, TempBox.Height, Draw:=False
        CurSel.SetIsText
        #If GetDIBitsErrors Then
        On Error Resume Next
        #End If
        dbGetDIBits TempBox.Image.Handle, TempBox.hDC, TransOrigData
        Erase TransData
        #If GetDIBitsErrors Then
        On Error GoTo 0
        #End If
        
        CurSel.SelMode = dbSuperTransparent
        'CurSel.ResizeSel CurSel.x1, CurSel.y1, CurSel.x2, CurSel.y2
        
        AC = ACol(1)
        ClearPic CurSel_SelData, AC

        dbPutSel
        mnuSelShow_Click
        ShowSelPicture
        
    End If
End With
unl:
'Unload frmInsOLE
frmInsOLE.Timer1.Enabled = False
Exit Sub
eh:
If Err.Number = dbCWS Then Resume unl
End Sub

Public Sub ShowSelPicture()
SelPicture.Visible = True
End Sub

Private Sub mnuHowToUse_Click()
ToggleHelpWindow
End Sub

Private Sub mnuIdleMessage_Click()
ShowIdleMessage 5
On Error Resume Next
LastUserActionTime = GetTickCount - 20000
FlashStatusBar 2
End Sub

Private Sub mnuKeyb_Click()
Dim i As Integer
ShowStatus 10031, , 3
Load frmKeyb
For i = 0 To UBound(Steps)
    frmKeyb.Stp(i).Text = CStr(Steps(i))
    If frmKeyb.Edi(i).Tag <> CStr(Edins(i)) Then frmKeyb.Edi_Click i
Next i

frmKeyb.Show vbModal
If Not frmKeyb.Tag = "" Then GoTo ExitHere

For i = 0 To UBound(Steps)
    Steps(i) = Val(frmKeyb.Stp(i).Text)
Next i
For i = 0 To UBound(Edins)
    Edins(i) = CBool(frmKeyb.Edi(i).Tag)
Next i

ExitHere:
Unload frmKeyb
End Sub

Private Sub mnuKeyb2_Click()
On Error GoTo eh
Keyb.ShowEditDialog
dbLoadCaptions
Keyb.SaveToReg "Keyboard", "(Current)"
Exit Sub
eh:
If Err.Number = dbCWS Then Exit Sub
MsgError
End Sub

Private Sub mnuLastEffect_Click()
On Error GoTo eh
dbEffect LastEffectIndex, True
Exit Sub
eh:
MsgError
End Sub

Private Sub mnuLnsDock_Click()
pctCapture_MouseDown 1, 2, 0, 0
End Sub

Private Sub mnuLnsToggle_Click()
pctCapture_MouseDown 1, 4, 0, 0
End Sub

Private Sub mnuLnsZoomIn_Click()
pctCapture_MouseDown 1, 1, 0, 0
End Sub

Private Sub mnuLnsZoomOut_Click()
pctCapture_MouseDown 2, 1, 0, 0
End Sub

Private Sub mnuLoadPal_Click()
Dim File As String
On Error GoTo eh
ShowStatus "$10002", , 3
File = ShowOpenDlg(dbFLoadPal, Me.hWnd, Purpose:="PAL")
On Error GoTo eh2
ShowStatus STT_Loading
dbLoadPal File
ShowStatus STT_READY
eh:
Exit Sub
Resume
eh2:
If Err.Number = dbCWS Then
    ShowStatus STT_Cancelled
    Exit Sub
End If
dbMsgBox grs(1135, "|1", Err.Description), vbCritical
End Sub

Sub dbLoadPal(File As String)
Dim dbPal() As Long, i As Integer, intCount As Integer
Dim Tips() As vtColorTip
LoadPalette File, dbPal, AryPtr(Tips)

intCount = UBound(dbPal)
If intCount < UBound(ChCol) Then
    ReDim ChCol(0 To intCount)
ElseIf intCount > UBound(ChCol) Then
    ReDim ChCol(0 To intCount)
End If

For i = 0 To UBound(dbPal)
    ChCol(i).BackColor = dbPal(i)
Next i
mnuEmptyTips_Click
On Error GoTo eh
For i = 0 To UBound(Tips)
    ChCol(Tips(i).nColor).Tip = Tips(i).strTip
Next i
On Error GoTo 0
Erase Tips
frmColors_Resize
Exit Sub
eh:
ReDim Tips(0 To 255)
Resume
End Sub

Private Sub mnuMix_Click()
Dim i As Long, j As Long, File As String
Dim iTop As Long, iLeft As Long, IHeight As Long, IWidth As Long
Dim pColor As Long, tColor As Long, UseBlack As Boolean, pData() As Long
On Error GoTo eh
ShowStatus "$10002", , 3
File = ShowPictureOpenDialog("", Purpose:="Sel")
LoadIntoSel File, Int(HScroll.Value / Zm), Int(VScroll.Value / Zm)
Exit Sub
eh:
If Err.Number = dbCWS Then
    ShowStatus GRSF(STT_Cancelled), , 1
    Exit Sub
Else
    'Err.Raise Err.Number
    dbMsgBox GRSF(1156), vbCritical 'The image might be corrupt or the format is not supported.
End If
Exit Sub
Resume
End Sub

Public Sub LoadIntoSel(ByRef File As String, _
                       ByVal x As Long, ByVal y As Long, _
                       Optional ByVal CenterCoords As Boolean = False)
Dim pData() As Long
Dim w As Long, h As Long
On Error GoTo eh
If CurSel.Selected Then dbDeselect blnApply:=True
vtLoadPicture pData, TransOrigData, File
TransDataChanged = True
Erase TransData
If AryDims(AryPtr(pData)) <> 2 Then Exit Sub
AryWH AryPtr(pData), w, h
If AryDims(AryPtr(TransOrigData)) = 2 Then
  dbNegative TransOrigData, 0, 0, w - 1, h - 1, &HFFFFFF
End If

ChTool ToolSel
x0 = 0
y0 = 0
dbMakeSel x, y, w, h, Draw:=False
SwapArys AryPtr(CurSel_SelData), AryPtr(pData)
Erase pData
If CenterCoords Then
    mnuSelShow_Click
End If
dbPutSel
ShowSelPicture
Exit Sub
eh:
If Err.Number = dbCWS Then
    ShowStatus GRSF(STT_Cancelled), , 1
    Exit Sub
Else
    'Err.Raise Err.Number
    dbMsgBox GRSF(1156), vbCritical 'The image might be corrupt or the format is not supported.
End If
Exit Sub
End Sub

Private Sub mnuMouseAttr_Click()
ScrollSettings.MouseGlued = Not ScrollSettings.MouseGlued
mnuMouseAttr.Checked = ScrollSettings.MouseGlued
ShowStatus IIf(ScrollSettings.MouseGlued, 2448, 2449), , 3
End Sub

Private Sub mnuNew_Click()
On Error GoTo eh
NewPicture
Exit Sub
eh:
MsgError
End Sub

'if neww,newh are set - dialog will appear
'otherwise,
Public Sub NewPicture(Optional ByVal NewW As Long, Optional ByVal NewH As Long)
    Dim Sz As Dims
    On Error Resume Next
    Load Dialog
    On Error GoTo eh
    
    If AryDims(AryPtr(Data)) = 2 Then
        Sz.w = UBound(Data, 1) + 1
        Sz.h = UBound(Data, 2) + 1
    Else
        Sz.w = 800
        Sz.h = 600
    End If
    
    ShowStatus GRSF(10001), , 3
    With Dialog
        .dbFrame2.Enabled = False
        .txtStretch.Visible = False
        
        .SetSz Sz
        .Show vbModal
        ShowStatus GRSF(STT_READY)
        
        .dbFrame2.Enabled = True
        .txtStretch.Visible = True
        
        .ExtractSz Sz
        
        If Len(.Tag) > 0 Then
            Err.Raise dbCWS
        End If
        
        'Resize Sz.w, Sz.h, Stretch:=False
        ReDim Data(0 To Sz.w - 1, 0 To Sz.h - 1)
        ClearPic Data, ACol(2)
        Refr
        OpenedFileName = ""
        OpenedFileFormatID = ""
        FileChanged = False
        FreshCaption
    End With
    
ExitHere:
    Unload Dialog
Exit Sub

eh:
    PushError
        Unload Dialog
    PopError
ErrRaise "NewPicture"
End Sub

Private Sub mnuNoUndoRedo_Click()
Dim i As Long
mnuNoUndoRedo.Checked = Not (mnuNoUndoRedo.Checked)
DisUndoRedo = mnuNoUndoRedo.Checked
If DisUndoRedo Then
    ShowStatus 10036, , 3
    ClearUndo
Else
    ShowStatus 10035, , 3
    ValidateUndoRedo
End If
End Sub

Private Sub mnuPalCount_Click()
Dim tmp As String
ShowStatus 10027, , 3
tmp = dbInputBox(LoadResString(1107), CStr(UBound(ChCol) + 1))
If tmp = "" Then
    ShowStatus STT_Cancelled
    Exit Sub
End If
Do While Val(tmp) > 512 Or Val(tmp) < 2
    tmp = dbInputBox("$1108", CStr(UBound(ChCol) + 1))
    If tmp = "" Then Exit Sub
Loop
ChangePalCount Val(tmp)
ShowStatus STT_READY
End Sub

Private Sub mnuPalDef_Click()
Dim i As Long
Dim tmp1 As String, tmp2 As String
Dim StrArr() As String
ShowStatus 10026, , 3
tmp1 = String$((UBound(ChCol) + 1) * 6, "0")
ReDim StrArr(0 To UBound(ChCol))
For i = 0 To UBound(ChCol)
    StrArr(i) = ChCol(i).Tip
    Mid(tmp1, i * 6 + 1, 6) = VedNullStr(Hex$(ChCol(i).BackColor And &HFFFFFF), 6)
Next i
tmp2 = Join(StrArr, Chr$(1))
dbSaveSetting "DefaultPalette", "Colors", tmp1, True
dbSaveSetting "DefaultPalette", "Tips", tmp2, True
dbSaveSetting "DefaultPalette", "ColorsCount", CStr((UBound(ChCol) + 1)), True
End Sub

Private Sub mnuPaste_Click()
Dim cb As New dbClipboard, tmp As VBRUN.ClipBoardConstants
Dim i As Long, j As Long
Dim w As Long, h As Long
Dim Img As IPictureDisp
Dim Data() As Long, Alpha() As Long
On Error GoTo eh
tmp = cb.CGFormat
If tmp = vbCFBitmap Or tmp = vbCFDIB Or tmp = vbCFEMetafile Or tmp = vbCFMetafile Then
    If CurSel.Selected Then
        StoreFragment CurSel.x1, CurSel.y1, CurSel.x2, CurSel.y2
        dbDeselect True
    End If
    Set Img = cb.cData
    GetPicData Img, Data, CalcAlpha:=True, aryptrAlpha:=AryPtr(Alpha)
    AryWH AryPtr(Data), w, h
    If w * h = 0 Then Err.Raise dbCWS
    If AryDims(AryPtr(Alpha)) = 2 Then
      dbNegative Alpha, 0, 0, w - 1, h - 1, &HFFFFFF
    End If
    ChTool ToolSel
    CurSel.Selected = True
    SwapArys AryPtr(CurSel_SelData), AryPtr(Data)
    SwapArys AryPtr(TransOrigData), AryPtr(Alpha)
    CurSel.SelMode = dbSuperTransparent
    mnuSelShow_Click
    dbPutSel
    ShowSelPicture
'    With TempBox
'        .ZOrder
'        .Visible = True
'        .AutoSize = True
'        .ForeColor = IIf(tmp = vbCFMetafile Or tmp = vbCFEMetafile, ACol(1), 0&)
'        .BackColor = IIf(tmp = vbCFMetafile Or tmp = vbCFEMetafile, ACol(2), 0&)
'        .Cls
'        .Picture = cb.cData
'        .Refresh
'        dbMakeSel 0, 0, .Width, .Height
'
'        #If GetDIBitsErrors Then
'        On Error Resume Next
'        #End If
'        dbGetDIBits .Image.Handle, MP.hDC, CurSel_SelData
'        #If GetDIBitsErrors Then
'        On Error GoTo 0
'        #End If
        If h > intH Or w > intW Then
          If dbMsgBox(GRSF(1110), vbYesNo Or vbQuestion) = vbYes Then
            Resize w, h, False
          End If
        End If
'        dbPutSel
'        .Visible = False
'    End With
ElseIf Clipboard.GetFormat(vbCFText) Then
  OutText Clipboard.GetText
Else
    dbMsgBox GRSF(1111), vbInformation
End If

Exit Sub
eh:
MsgError

End Sub

Private Sub mnuPrgDraw_Click()
Dim Prg As SMP
Dim EV As New clsEVal
On Error GoTo eh
UpdateWH
Load frmProgTool
With frmProgTool
    ShowHelp 11012
    .SetMode epmDraw
    .Show vbModal
    ShowHelp Me.HelpContextID
    If .Tag = "" Then
        .GetPrg Prg
        With Prg
          .Vars(0).Value = intW \ 2
          .Vars(1).Value = intH \ 2
          .Vars(2).Value = intW
          .Vars(3).Value = intH
          .Vars(4).Value = ACol(1)
          .Vars(5).Value = ACol(2)
        End With
        EV.ExecuteSMP Prg
        Refr
        '.GetVars ToolVars
    End If
End With
Unload frmProgTool
Exit Sub
eh:
MsgError
End Sub

Private Sub mnuPrint_Click()
Dim f As Boolean
On Error GoTo eh
ShowStatus GRSF(1235) 'Preparing the dialog
frmPrint.SetData Data
ShowStatus GRSF(10005), , 3
frmPrint.Show vbModal
ShowStatus GRSF(STT_READY)
Exit Sub
eh:
MsgError
End Sub

Private Sub mnuPSens_Click()
Load frmPressure
With frmPressure
    .Timer1.Enabled = True
    .UnlNow = False
    .OldMaxPr = MaxPenPressure
    .MaxPr = 0
    .Show vbModal
End With
Unload frmPressure
End Sub

Private Sub mnuQBColors_Click()
Dim i As Long
ChangePalCount 16
For i = 0 To 15
    ChCol(i).BackColor = QBColor(i)
    ChCol(i).Tip = CStr(i)
Next i
frmColors.Refresh
mnuDeadEnds.Enabled = True
End Sub

Private Sub mnuRedo_Click()
If MeEnabled And mnuRedo.Enabled Then
    Redo
End If
End Sub

Private Sub mnuResetAll_Click()
'Dim nRegName As String, nRegCode As String
On Error Resume Next
If dbMsgBox(1185, vbYesNo Or vbDefaultButton2) = vbNo Then 'Are you sure?
    Exit Sub
End If
SetMousePtr True

dbDeleteSetting ""
FlushSettings
dbSaveSetting "Special", "ExecCount", CStr(ExeCount), True, True
If dbMsgBox(1187, vbYesNo Or vbDefaultButton2) = vbYes Then 'Reset filter presets?
    dbDeleteSetting "Effects\Soft", , True
End If
If dbMsgBox(1188, vbYesNo Or vbDefaultButton2) = vbYes Then 'Reset default palette?
    dbDeleteSetting "DefaultPalette", , True
End If
BuildBackup ShowMessage:=False
UninstallHook
DestroyHiResTimer
Shell ExePath
End
End Sub

Private Sub mnuResetOrg_Click()
TexOrg.x = 0
TexOrg.y = 0
End Sub

Private Sub mnuResetPal_Click()
Dim i As Integer, MsgText As String, Answ As VbMsgBoxResult
Dim tmp1 As String, tmp2 As String
Dim StrArr() As String
Dim UB As Long
Dim ResUB As Long
'ChangePalCount Val(dbGetSetting("DefaultPalette", "ColorsCount", "66", True))
UB = Val(dbGetSetting("DefaultPalette", "ColorsCount", "66", True)) - 1
ReDim ChCol(0 To UB)
On Error GoTo eh
MsgText = "Unable to load colors"

tmp1 = dbGetSetting("DefaultPalette", "Colors", "", True)
tmp2 = dbGetSetting("DefaultPalette", "Tips", "", True)

ResUB = Len(tmp1) \ 6 - 1

If tmp2 <> "" Then
    StrArr = Split(tmp2, Chr$(1))
    tmp2 = vbNullString
Else
    ReDim StrArr(0 To 0)
End If

For i = 0 To UBound(ChCol)
    If i <= ResUB Then
        ChCol(i).BackColor = CLng("&H" + Mid$(tmp1, i * 6 + 1, 6))  'CLng(dbGetSetting("DefaultPalette", "Color#" + String$(3 - Len(CStr(i)), "0") + CStr(i), CStr(SMBDefPal(i)), True))
    ElseIf i <= UBound(SMBDefPal) Then
        ChCol(i).BackColor = SMBDefPal(i)
    Else
        ChCol(i).BackColor = Int(Rnd(1) * &H1000000)
    End If
    
    ChCol(i).BackColor = ConvertColorLng(ChCol(i).BackColor)
    
    If i <= UBound(StrArr) Then
        ChCol(i).Tip = StrArr(i)
    End If
Next i
frmColors_Resize
ShowStatus 10028, , 3
Exit Sub
eh:
    ShowStatus Err.Description
    Answ = MsgBox(MsgText, vbCritical Or vbAbortRetryIgnore, "Reset")
    If Answ = vbRetry Then
        Resume
    ElseIf Answ = vbIgnore Then
        Resume Next
    ElseIf Answ = vbAbort Then
        Exit Sub
    End If
End Sub

Private Sub mnuRestoreOrg_Click()
If Not MeEnabled Then Exit Sub
If TexOrg.x = 0 And TexOrg.y = 0 Then Exit Sub
BUD
MoveDataOrg Data, TexOrg.x, TexOrg.y
TexOrg.x = 0
TexOrg.y = 0
Refr
End Sub

Friend Sub mnuSaveFile_Click()
On Error GoTo eh
SaveAuto
ShowStatus STT_READY
Exit Sub
eh:
MsgError
End Sub

Public Sub SavePal_Auto(ByVal File As String)
Dim dbPal() As Long, i As Integer, UB As Integer
Dim Tips() As vtColorTip
Dim j As Integer
On Error GoTo eh2
UB = UBound(ChCol)
ReDim dbPal(0 To UB)
For i = 0 To UB
    dbPal(i) = ChCol(i).BackColor
Next i
ReDim Tips(0 To UB)
For i = 0 To UB
    If Len(ChCol(i).Tip) > 0 Then
        Tips(j).nColor = i
        Tips(j).strTip = ChCol(i).Tip
        j = j + 1
    End If
Next i
If j > 0 Then
    ReDim Preserve Tips(0 To j - 1)
Else
    Erase Tips
End If
SavePalette dbPal, File, AryPtr(Tips)
eh:
Exit Sub
Resume
eh2:
If Err.Number = dbCWS Then
    ShowStatus STT_Cancelled
    Exit Sub
End If
dbMsgBox grs(1157, "|1", Err.Description), vbCritical
End Sub

Private Sub mnuSavePal_Click()
Dim File As String, dbPal() As Long, i As Integer, UB As Integer
Dim Tips() As vtColorTip
Dim j As Integer
On Error GoTo eh
ShowStatus 10003, , 3
File = ShowSaveDlg(dbFSavePal, MP.hWnd, Purpose:="PAL")
On Error GoTo eh2
SavePal_Auto File
eh:
Exit Sub
Resume
eh2:
If Err.Number = dbCWS Then
    ShowStatus STT_Cancelled
    Exit Sub
End If
dbMsgBox grs(1157, "|1", Err.Description), vbCritical
End Sub

Private Sub mnuSaveSel_Click()
Dim Alpha() As Long
If SelectionPresent Then
    On Error GoTo eh
    vtSavePicture CurSel_SelData, TransData, FileName:="", ShowDialog:=True, Purpose:="Sel"
End If
Exit Sub
eh:
MsgError
End Sub

Private Sub mnuSelBrush_Click()
Dim x As Long, y As Long
Dim UBX As Long, UBY As Long
Dim r As Long, g As Long, b As Long
If Not CurSel.Selected Then
    Exit Sub
End If
With CurSel
    If Abs(.x2 - .x1) + 1 > 25 Or Abs(.y2 - .y1) > 25 Then
        dbMsgBox 1196, vbCritical
        Exit Sub
    End If
    UBX = UBound(CurSel_SelData, 1)
    UBY = UBound(CurSel_SelData, 2)
    ReDim CurBrush(0 To UBX, 0 To UBY)
    For y = 0 To UBY
        For x = 0 To UBX
            GetRgbQuadLongEx2 CurSel_SelData(x, y), r, g, b
            CurBrush(x, y) = (r + g + b) \ 3
        Next x
    Next y
End With
End Sub

Private Sub mnuSelectAll_Click()
Dim i As Long, j As Long
If CurSel.Selected Then
    'StoreFragment CurSel.x1, CurSel.y1, CurSel.x2, CurSel.y2
    dbDeselect True
End If
'BUD
UpdateWH
BUD
dbMakeSel 0, 0, intW, intH, Draw:=False
SwapArys AryPtr(CurSel_SelData), AryPtr(Data)
ClearPic Data, ACol(2)
Refr 'dbPutSel dbUseCurSelMode
SelPicture.Visible = True
End Sub

Private Sub mnuSetAutoScrolling_Click()
'Dim tmp As Long
'On Error GoTo eh
'With ScrollSettings.ASS
'  tmp = (.GapLef + .GapTop + .GapRig + .GapBot) / 4
'  EditNumber tmp, 2436, MinValue:=-10000, MaxValue:=10000
'  'tmp = dbVal(dbInputBox(2436, CStr(AutoScroll_Field_Size),  True), vbLong, 0, 2000)
'  'AutoScroll_Field_Size = tmp
'  .GapLef = tmp: .GapTop = tmp: .GapRig = tmp: .GapBot = tmp
'End With
'Exit Sub
'eh:
'    MsgError
mnuDynamicScr_Click
End Sub


Private Sub mnuStretchPal_Click()
Dim tmp As String
Dim nCount As Long
Dim UB As Long, NUB As Long
Dim i As Long
Dim k As Single
Dim j As Long
Dim rgb1 As RGBQuadLong, rgb2 As RGBQuadLong
Dim frgb1 As RGBTriCurr
Dim RGBPal() As RGBTriCurr
Dim w As Single
Dim dx As Single
On Error GoTo eh
UB = UBound(ChCol)
tmp = dbInputBox("$1191", CStr(UBound(ChCol) + 1), True)
If tmp = "" Then
    ShowStatus STT_Cancelled
End If
nCount = Val(tmp)
Do While nCount < 2 Or nCount > 512
    tmp = dbInputBox("$1192", CStr(UB + 1))
    If tmp = "" Then
        ShowStatus STT_Cancelled
    End If
    nCount = Val(tmp)
Loop
NUB = nCount - 1
If NUB > UB Then
    ReDim Preserve ChCol(0 To nCount - 1)
    For i = NUB To 0 Step -1
        j = i * UB \ NUB
        k = i * UB / NUB - j
        GetRgbQuadLongEx ChCol(j).BackColor, rgb1
        If k > 0 Then
            GetRgbQuadLongEx ChCol(j + 1).BackColor, rgb2
        End If
        ChCol(i).BackColor = RGB( _
                (rgb2.rgbRed - rgb1.rgbRed) * k + rgb1.rgbRed, _
                (rgb2.rgbGreen - rgb1.rgbGreen) * k + rgb1.rgbGreen, _
                (rgb2.rgbBlue - rgb1.rgbBlue) * k + rgb1.rgbBlue)
    Next i
    frmColors_Resize
ElseIf NUB < UB Then
    ReDim RGBPal(0 To NUB)
    For i = 0 To UB
        GetRgbQuadLongEx ChCol(i).BackColor, rgb1
        j = i * (NUB + 1) \ (UB + 1)
        w = NUB / UB
        dx = (j + 1) * (UB + 1) / (NUB + 1) - i
        If dx > 1 Then dx = 1
        k = dx * w
        
        RGBPal(j).rgbRed = RGBPal(j).rgbRed + rgb1.rgbRed * k
        RGBPal(j).rgbGreen = RGBPal(j).rgbGreen + rgb1.rgbGreen * k
        RGBPal(j).rgbBlue = RGBPal(j).rgbBlue + rgb1.rgbBlue * k
        
        If dx < 1 Then
            k = (1 - dx) * w
            RGBPal(j + 1).rgbRed = RGBPal(j + 1).rgbRed + rgb1.rgbRed * k
            RGBPal(j + 1).rgbGreen = RGBPal(j + 1).rgbGreen + rgb1.rgbGreen * k
            RGBPal(j + 1).rgbBlue = RGBPal(j + 1).rgbBlue + rgb1.rgbBlue * k
        End If
    Next i
    ReDim ChCol(0 To NUB)
    For i = 0 To NUB
        ChCol(i).BackColor = RGB(Round(RGBPal(i).rgbRed), Round(RGBPal(i).rgbGreen), Round(RGBPal(i).rgbBlue))
    Next i
    frmColors_Resize
End If
Exit Sub
eh:
If Err.Number = dbCWS Then
    ShowStatus STT_Cancelled
    Exit Sub
End If
ShowStatus Err.Description, , 3
vtBeep
End Sub

Private Sub mnuSysPal_Click()
Dim Vals() As String, Names() As String
Dim tmpArr() As String
Dim rrr As Reg
Dim i As Long
Set rrr = New Reg
    rrr.GetAllValues HKEY_CURRENT_USER, "Control Panel\Colors", Names, Vals
Set rrr = Nothing
ChangePalCount UBound(Names) + 1
For i = 0 To UBound(Names)
    tmpArr = Split(Vals(i), " ")
    ChCol(i).BackColor = RGB(Val(tmpArr(0)), Val(tmpArr(1)), Val(tmpArr(2)))
    ChCol(i).Tip = Names(i)
Next i
frmColors.Refresh
End Sub

Private Sub mnuTexMode_Click()
SetTexMode Not TexMode
End Sub

Friend Sub SetTexMode(ByVal NewMode As Boolean)
If TexMode = NewMode Then Exit Sub
TexMode = NewMode
mnuTexMode.Checked = TexMode
End Sub

Private Sub mnuToolBarVis_Click()
SetToolBar2Visible Not (ToolBar2Visible)
End Sub

Private Sub mnuUndoLim_Click()
Dim tmp As String
ShowStatus 10034, , 3
tmp = dbInputBox(GRSF(1113), CStr(UndoSize \ (1024& * 1024&)))
If tmp = "" Then Exit Sub
Do While Val(tmp) < 1 Or Val(tmp) > 256
    tmp = dbInputBox(GRSF(1114), CStr(UndoSize))
    If tmp = "" Then
        ShowStatus STT_Cancelled
        Exit Sub
    End If
Loop
ChUndoSize Val(tmp) * 1024& * 1024&
ShowStatus STT_READY
End Sub

Private Sub mnuUseWheel_Click(Index As Integer)
Dim i As Integer
    For i = mnuUseWheel.lBound To mnuUseWheel.UBound
        mnuUseWheel(i).Checked = (i = Index)
    Next i
ShowStatus 10029 + Index, , 3
End Sub

Public Sub dbSetWheelUse(Index As Integer)
Dim i As Integer
    For i = mnuUseWheel.lBound To mnuUseWheel.UBound
        mnuUseWheel(i).Checked = (i = Index)
    Next i
End Sub

Public Function dbGetWheelUse() As Integer
Dim i As Integer
    For i = mnuUseWheel.lBound To mnuUseWheel.UBound
        If mnuUseWheel(i).Checked Then dbGetWheelUse = i: Exit Function
    Next i
    dbGetWheelUse = -1
End Function

Private Sub mnuWeb_Click()
WinRun "http://vt-dbnz.narod.ru"
End Sub

Private Sub mnuWebForum_Click()
WinRun "http://forum.biophysicist.net/viewforum.php?f=10"
End Sub

Private Sub mnuWebMail_Click()
If dbMsgBox(2450, vbInformation Or vbOKCancel) = vbCancel Then Exit Sub
WinRun "mailto:vt-dbnz@yandex.ru"
End Sub

Private Sub mnuWebUpdates_Click()
WinRun "http://vt-dbnz.narod.ru/vb/smb/smb.html"
mnuAbout_Click
End Sub

Private Sub mnuZoom_Click()
Load frmZoom
ShowStatus "$10008", , 3
frmZoom.SetZoom Zm
frmZoom.Show vbModal
If frmZoom.Tag = "" And frmZoom.GetZoom <> Zm Then
    ReZoom frmZoom.GetZoom
End If
Unload frmZoom
ShowStatus GRSF(STT_READY)
End Sub

Public Sub ShowZoomStatus()
ShowStatus grs(1252, "%zm%", Zm), HoldTime:=1
End Sub

Public Sub ReZoom(ByVal NewZm As Integer, _
                  Optional ByVal CenterPointX As Long = -1, _
                  Optional ByVal CenterPointY As Long = -1)
Dim NAR As Boolean
Dim LZM As Long
Dim k As Single
Dim dh As Long
Dim OFR As Boolean
Dim nV As Long
Dim dx As Long, dy As Long
Dim OldMouseGlued As Boolean
Const Max_Size_For_AutoRedraw As Long = 2000& * 2000&
On Error GoTo eh

UpdateWH

If NewZm < 1 Then NewZm = 1: vtBeep
If NewZm > 32 Then NewZm = 32: vtBeep
If intH * NewZm > 16383 Or intW * NewZm > 16383 Then
    NewZm = Min(16383 \ intH, 16383 \ intW)
    vtBeep
End If
If Zm = NewZm Then Exit Sub
LZM = Zm
Zm = NewZm

OldMouseGlued = ScrollSettings.MouseGlued
ScrollSettings.MouseGlued = False
Refr

ShowZoomStatus
'Form_Resize

If CenterPointX = -1 Then CenterPointX = -MP.Left + MPHolder.ScaleWidth \ 2
If CenterPointY = -1 Then CenterPointY = -MP.Top + MPHolder.ScaleHeight \ 2

dx = CenterPointX - CenterPointX * Zm / LZM
dy = CenterPointY - CenterPointY * Zm / LZM
ScrollSettings.DontScroll = True
ChangeScrollBarValue HScroll, -dx
ChangeScrollBarValue VScroll, -dy
ScrollSettings.DontScroll = False
ApplyScrollBarsValues True

ScrollSettings.MouseGlued = OldMouseGlued
ValidateZoom
'FreezeRefresh = OFR

Exit Sub
eh:
'FreezeRefresh = False
Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Sub ReZoomPtr(ByVal NewZm As Integer)
Dim pos As POINTAPI
Dim Wnd As RECT
GetCursorPos pos
GetWindowRect MPHolder.hWnd, Wnd

'remove frame - like GetClientRect
Wnd.Left = Wnd.Left + 2
Wnd.Top = Wnd.Top + 2
Wnd.Right = Wnd.Right - 2
Wnd.Bottom = Wnd.Bottom - 2

If InRectAPI(pos, Wnd) Then
    GetWindowRect MP.hWnd, Wnd
    pos.x = pos.x - (Wnd.Left)
    pos.y = pos.y - (Wnd.Top)
    ReZoom NewZm, pos.x, pos.y
Else
    ReZoom NewZm
End If


End Sub

Private Function InRectAPI(ByRef TestPoint As POINTAPI, ByRef Rectangle As RECT) As Boolean
InRectAPI = TestPoint.x >= Rectangle.Left And _
         TestPoint.x < Rectangle.Right And _
         TestPoint.y >= Rectangle.Top And _
         TestPoint.y < Rectangle.Bottom
End Function

Private Sub Mover_Timer()
Dim papi As POINTAPI
Static PrtScr As Boolean
Static Counter1 As Long
If Not PrtScr Then
    PrtScr = True
    WasKey 44
End If
GetCursorPos papi
If ActiveTool = ToolAir Then
    SetCursorPos papi.x, papi.y
End If
PctCapture_Paint
'check whether PrtScr was pressed
If WasKey(44) Then
    'prtscr WAS pressed.
End If
Counter1 = Counter1 + 1&
BreakKeyPressed
If Counter1 = 40& Then
    Counter1 = 0
    On Error Resume Next
    If Abs(GetTickCount - LastUserActionTime) >= 30000 Then
        ShowIdleMessage
        LastUserActionTime = GetTickCount - 20000
    End If
    On Error GoTo 0
    If HelpWindowVisible Then
        EnableWindow frmHelp.hWnd, True
    End If
    If LensWindowVisible Then
        ReposFrmLens
    End If
End If
'dbProcessMessages
End Sub

Private Sub Form_Initialize()
Dim h As Long, v As Long, Answ As VbMsgBoxResult, MsgText As String
Dim i As Long, Obj As Object, tmp As String, j As Long
Dim nmb As Long, p1 As Long, p2 As Long, tmpPal() As Long
Dim ShowSplash As Boolean
    Zm = 1
       dbTag = "Loading..."
ReDim ChCol(0 To 1)
    ShowSplash = dbGetSettingEx("Options", "ShowSplash", vbBoolean, False) _
                 Or Len(Command$) > 0
    If ShowSplash Then
        frmSplash.Show
        frmSplash.Refresh
    End If
    ChangeDynScrolling True
    ScrollSettings.DS_EnL = 0.5
    ScrollSettings.DS_Jestkost = 0.5
    ReDim SelMatrix(0 To 2, 0 To 5)
    MeEnabled = False
    InitToolsSizes
    ToolBar2Visible = True
    For i = 0 To 8
        Stepen2(i) = 2 ^ i
    Next i
    For i = 0 To 3
        Stepen256(i) = 256 ^ i
    Next i
    Stepen2Long(0) = &H1&
    Stepen2Long(1) = &H2&
    Stepen2Long(2) = &H4&
    Stepen2Long(3) = &H8&
    Stepen2Long(4) = &H10&
    Stepen2Long(5) = &H20&
    Stepen2Long(6) = &H40&
    Stepen2Long(7) = &H80&
    Stepen2Long(8) = &H100&
    Stepen2Long(9) = &H200&
   Stepen2Long(10) = &H400&
   Stepen2Long(11) = &H800&
   Stepen2Long(12) = &H1000&
   Stepen2Long(13) = &H2000&
   Stepen2Long(14) = &H4000&
   Stepen2Long(15) = &H8000&
   Stepen2Long(16) = &H10000
   Stepen2Long(17) = &H20000
   Stepen2Long(18) = &H40000
   Stepen2Long(19) = &H80000
   Stepen2Long(20) = &H100000
   Stepen2Long(21) = &H200000
   Stepen2Long(22) = &H400000
   Stepen2Long(23) = &H800000
   Stepen2Long(24) = &H1000000
   Stepen2Long(25) = &H2000000
   Stepen2Long(26) = &H4000000
   Stepen2Long(27) = &H8000000
   Stepen2Long(28) = &H10000000
   Stepen2Long(29) = &H20000000
   Stepen2Long(30) = &H40000000
   Stepen2Long(31) = &H80000000
    
    RGBMask(1) = &HFF&
    RGBMask(2) = &HFF00&
    RGBMask(3) = &HFF0000
    Set CurSel = New dbSelection_class
    Set CurSel.mnuSelection = mnuSelAll
'    Scroll_Speed = -1
    'ScrollSettings.DS_Jestkost = 10
    'ScrollSettings.DS_EnL = 1.21
    
    InitHiResTimer
    InstallHook
    
    On Error GoTo eh
    FastLoad = True 'CBool(dbGetSetting("Options", "FastLoad", True))
    On Error GoTo eh
    '----------------------window---------------------------
    'LoadWindowPos Me, OnlyWH:=
    If dbGetSettingEx(Me.Name, "RememberPos", vbBoolean, DefValue:=False) Then
      MsgText = "Invalid Form.Top"
      Me.Top = dbGetSettingEx("MainForm", "Top", vbLong, CLng(Me.Top))
      MsgText = "Invalid Form.Left"
      Me.Left = dbGetSettingEx("MainForm", "Left", vbLong, CLng(Me.Left))
      mnuOptRememberWndPos.Checked = True
    Else
      mnuOptRememberWndPos.Checked = False
    End If
    MsgText = "Invalid Form.Height"
    FormH = dbGetSettingEx("MainForm", "Height", vbLong, 5235&)
    MsgText = "Invalid Form.Width"
    FormW = dbGetSettingEx("MainForm", "Width", vbLong, 6210&)
    MsgText = "Invalid Form.WindowState"
    Me.WindowState = dbGetSettingEx("MainForm", "WindowState", vbLong, Me.WindowState)
    UpdateFormSize
    
    

    MsgText = "Invalid boolean value (Palette.Visible)"
    If Not dbGetSettingEx("View", "Palette", vbBoolean, True) Then mnuPal_Click
    
    MsgText = "Invalid Boolean value (Toolbar)"
    SetToolBar2Visible dbGetSettingEx("View", "Toolbar", vbBoolean, True)
    '-------------------------image appearence---------------------
    MsgText = "Invalid zoom ratio"
    Zm = dbGetSettingEx("View", "Zoom", vbInteger, 6)
    If Zm <= 0 Then Zm = 1
    ValidateZoom
    
    
    LoadScrollSettings
    '-------------------------tool------------------------
    MsgText = "Invalid tool index"
    ChTool dbGetSettingEx("Tool", "LastUsedToolIndex", vbInteger, 0)
    
    LoadFadeDesc gFDSC, MsgText
    
    LoadHSet HSet, MsgText
    
    MsgText = "Invalid rectangle style"
    dbRS = dbGetSetting("Tool", "RectStyle", vbLong, 0)

    MsgText = "Error in last brush"
    LoadLastBrush CurBrush
    
    MsgText = "Invalid Circle.Flags"
    LoadCircleFlags CircleFlags
    
    MsgText = "Invalid Line.Flags"
    LoadLineFlags LineOpts
    
    CurSel.TransRatio = dbGetSettingEx("Tool", "SelTransR", vbSingle, 50) / 100#
    CurSel.StretchMode = dbGetSettingEx("Tool", "SelStretchMode", vbInteger, eStretchMode.SMSquares)
    CurSel.SelMode = dbGetSettingEx("Tool", "SelMode", vbInteger, dbSelMode.dbReplace)
    LoadMatrix SelMatrix, "Tool", "SelMatrix", "0;0;0;1;0;0;0;0;0;0;1;0;0;0;0;0;0;1"
    
    FillOpts.TexOrigin.LoadFromReg "Tool", "FillTexAlign"
    
    
    LoadLensSettings
    
    '----------------------options----------------------
    MsgText = "Bad undo limit"
    ChUndoSize dbGetSettingEx("Options", "UndoSizeLimit", vbLong, 16& * 1024& * 1024&)
    
    MsgText = "Bad Wheeluse"
    dbSetWheelUse dbGetSettingEx("Options", "UseWheelButton", vbInteger, 0)
    
    MaxPenPressure = dbGetSettingEx("Options", "MaxPenPressure", vbLong, -1)
    
    Keyb.LoadFromReg "Keyboard", "(Current)"
    dbLoadCaptions
    
    For i = 0 To UBound(Steps)
        MsgText = "Bad mouse step"
        Steps(i) = dbGetSettingEx("Options", "MouseStep" + CStr(i), vbInteger, Choose(i + 1, 1, 1, 16, 32))
        MsgText = "Bad boolaen value (options.mouse step unit)"
        Edins(i) = dbGetSettingEx("Options", "MouseStepZm" + CStr(i), vbBoolean, Choose(i + 1, True, False, True, True))
    Next i
    
    MsgText = "Bad boolean value (MouseAttached)"
    'ScrollSettings.MouseGlued = dbGetSettingEx("Options", "MouseAttached", vbBoolean, False)
    'mnuMouseAttr.Checked = ScrollSettings.MouseGlued
    ''moved to LoadScrollSettings
    
    'MsgText = "Bad AutoScroll field size"
    'LoadAutoScrollFieldSize
    
    MsgText = "Bad boolean value (DisableUndo)"
    mnuNoUndoRedo.Checked = dbGetSettingEx("Options", "DisableUndo", vbBoolean, False)
    DisUndoRedo = mnuNoUndoRedo.Checked
    
    mnuShowSplash.Checked = ShowSplash
    
    
    '-------------------------palette--------------------------
    ExtractResPal SMBDefPal, "DEFAULT"
    ReDim Preserve SMBDefPal(0 To 512)
    
    MsgText = "Bad number of colors in palette"
    ChangePalCount dbGetSettingEx("DefaultPalette", "ColorsCount", vbInteger, 66, True)
    mnuResetPal_Click
    
    MsgText = "Invalid foreground color"
    ACol(1) = dbGetSettingEx("Colors", "FColor", vbLong, &HFFFFFF, , , True) And &HFFFFFF
    MsgText = "Invalid background color"
    ACol(2) = dbGetSettingEx("Colors", "BColor", vbLong, &H0&, , , True) And &HFFFFFF
           
           
    FreshActiveColors
        
    '-------------------------------PNGs----------------------------
    MsgText = "Invalid png BPP"
    pngBPP = dbGetSettingEx("Options", "PNG bits-per-pixel", vbInteger, 24)
    ClearUndo
    '/frmIcLoad
    '------------------------------kernel--------------------------
    'Debug.Print GetTimer
    IncExecCount
    MsgText = ""
    '-------------------------end of settings-------------------------
    On Error GoTo 0
    For i = mnuTool.lBound To mnuTool.UBound
        btnTool(i).ToolTipText = dbRemoveAmpersand(mnuTool(i).Caption)
    Next i
    btnTool(btnTool.UBound).ToolTipText = dbRemoveAmpersand(mnuToolOpts.Caption)
    On Error GoTo eh
    If Not Command = "" And Not UCase$(Command) = "/INSTALL" Then
        tmp = Command
        p1 = InStr(1, tmp, """")
        If p1 > 0 Then
            p2 = InStr(p1 + 1, tmp, """")
            tmp = Mid(tmp, p1 + 1, p2 - p1 - 1)
        Else
            tmp = Trim(tmp)
        End If
        On Error GoTo LoadFileErr
        LoadFile (tmp)
        Refr
        CDl.FileName = Command
Cnc:
        On Error GoTo eh
    ElseIf UCase(Trim(Command$)) = "/INSTALL" Then
        tmrInstall.Enabled = True
        LoadWH
        Refr
    Else
ClearCreate:
        On Error GoTo eh
        LoadWH
        Refr
    End If
    dbTag = ""
    frmColors_Resize
    ToolBar_Resize
    If ShowSplash Then Unload frmSplash
    MeEnabled = False
    ShowGreeting
    If NeedRefr And MP.AutoRedraw Then Refr
    MessageLoopStarter.Enabled = True
    MPHolder_Resize
    'Debug.Print GetTimer
Exit Sub
eh:
    If Len(MsgText) = 0 Then MsgText = Err.Description
    Debug.Assert False
    Answ = MsgBox(MsgText, vbCritical Or vbAbortRetryIgnore, "Loading")
    If Answ = vbRetry Then
        Resume
    ElseIf Answ = vbIgnore Then
        Resume Next
    ElseIf Answ = vbAbort Then
        DestroyHiResTimer
        UninstallHook
        End
    End If
    Exit Sub
LoadFileErr:
    If Err.Number = dbCWS Then
        dbEnd
    End If
    MsgError
    Resume ClearCreate
End Sub

Public Sub dbEnd()
DestroyHiResTimer
UninstallHook
End
End Sub

Public Sub LoadWH()
Dim w As Long, h As Long
w = dbGetSettingEx("ImageSize", "Width", vbInteger, 800)
If w < 2 Or w > 9000 Then
    w = 800
End If
h = dbGetSettingEx("ImageSize", "Height", vbInteger, 600)
If h < 2 Or h > 9000 Then
    h = 800
End If
ReDim Data(0 To w - 1, 0 To h - 1)
ClearPic Data, ACol(2)
End Sub

Public Sub UpdateFormSize()
If Me.WindowState = vbNormal Then
    Me.Move Me.Left, Me.Top, FormW, FormH
End If
End Sub

Public Sub LoadCircleFlags(ByRef Flags As dbCircleFlags)
Dim tmp As dbCircleFlags
tmp = dbGetSettingEx("Tool", "CircleFlags", vbLong, &H2& Or &H80&)
If ((tmp And Not (&HF& Or &H10& Or &H20& Or &H40& Or &H80&)) <> 0) Or ((tmp And &HF) > 3) Then
    Flags = 1
    Err.Raise 101, "LoadCircleFlags", "Bad Value"
End If
Flags = tmp
End Sub

Private Sub LoadLineFlags(ByRef Flags As LineSettings)
Flags.GeoMode = dbGetSettingEx("Tool", "LineGeoMode", vbLong, eLineGeoMode.dbLineSimple)
Flags.AntiAliasing = dBtoFactor(-dbGetSettingEx("Tool", "LineAntiAliasing", vbDouble, 0))
Flags.Weight = dbGetSettingEx("Tool", "LineWeight", vbDouble, 1)
Flags.RelWeight1 = dbGetSettingEx("Tool", "LineRelWeight1", vbDouble, 1)
Flags.RelWeight2 = dbGetSettingEx("Tool", "LineRelWeight2", vbDouble, 1)
End Sub

Private Sub SaveLineFlags(ByRef Flags As LineSettings)
dbSaveSettingEx "Tool", "LineGeoMode", Flags.GeoMode
dbSaveSettingEx "Tool", "LineAntiAliasing", -FactorToDB(Flags.AntiAliasing)
dbSaveSettingEx "Tool", "LineWeight", Flags.Weight
dbSaveSettingEx "Tool", "LineRelWeight1", Flags.RelWeight1
dbSaveSettingEx "Tool", "LineRelWeight2", Flags.RelWeight2
End Sub

Sub Refr(Optional ByVal DoDoEvents As Boolean = True)
Dim OldIcon As Integer
Dim NAR As Boolean
If FreezeRefresh Then Exit Sub
If Len(dbTag) > 0 Then
    NeedRefr = True
    Exit Sub
End If
OldIcon = Screen.MousePointer
Screen.MousePointer = vbHourglass
NeedRefr = False
ShowStatus GRSF(1212) 'Wroking

ClearTempStorage
UpdateWH
If intW = 0 Then Exit Sub
If intH * Zm > 16383 Or intW * Zm > 16383 Then
    Zm = Min(16383 \ intH, 16383 \ intW)
    ShowZoomStatus
    'Exit Sub
End If

NAR = intH * Zm * intW * Zm <= 2000& * 2000&
If NAR Then
    If MP.AutoRedraw And MPhDefBitmap <> 0 Then
        If intH * Zm <> MPBitsHeight Or intW * Zm <> MPBitsWidth Then
            'remake only if resized
            RestoreMP
            MP.Cls
            MakeARMP
        End If
    ElseIf MP.AutoRedraw And MPhDefBitmap = 0 Then
        MP.Cls
        MakeARMP
    ElseIf Not MP.AutoRedraw Then
        If MPhDefBitmap <> 0 Then 'for sure
            RestoreMP
        End If
        MP.AutoRedraw = True
        MP.Cls
        MakeARMP
    End If
Else 'if not NAR
    If MPhDefBitmap <> 0 Then
        RestoreMP
        MP.Cls
    End If
    MP.AutoRedraw = False
End If
MP.Move MP.Left, MP.Top, intW * Zm, intH * Zm

If MP.AutoRedraw Then
    'On Error GoTo eh2
    CancelDoEvents Not DoDoEvents
    SendDataToMP
    RestoreDoEvents
'    MP.Refresh
    SelPicture_Paint
Else
    MP.Refresh
    SelPicture_Paint
End If
ShowStatus GRSF(1204) 'Ready!
ShowStatus "", 2
Screen.MousePointer = OldIcon
Exit Sub
eh:
MsgError
SetMousePtr False
Exit Sub
eh2:
If Err.Number = dbCantCreatePic Then
    dbMsgBox 1171, vbCritical
Else
    MsgError
End If
SetMousePtr False
Exit Sub
Resume
End Sub

Private Function SHeight()
Dim dh As Long
dh = 0
If ToolBar2Visible Then
    dh = dh + ToolBar2.Height
End If
If frmColors.Visible Then
    dh = dh + frmColors.Height
End If
If Picture1.Visible Then
    dh = dh + Picture1.Height
End If
SHeight = Me.ScaleHeight - dh
End Function


Private Sub frmColors_Resize()
Dim i As Long, tmp As Long
If Not dbTag = "" Then Exit Sub
tmp = ((UBound(ChCol) + 1) + 1) \ 2
For i = 0 To tmp - 1
    With ChCol(i)
        '.Move frmColors.ScaleWidth * i \ tmp, 0, (frmColors.ScaleWidth * (i + 1&) \ tmp) - (frmColors.ScaleWidth * i \ tmp), frmColors.ScaleHeight \ 2&
        .Left = frmColors.ScaleWidth * i \ tmp
        .Top = 0
        .Width = (frmColors.ScaleWidth * (i + 1&) \ tmp) - (frmColors.ScaleWidth * i \ tmp)
        .Height = frmColors.ScaleHeight \ 2&
    End With
Next i
For i = tmp To UBound(ChCol)
    With ChCol(i)
        '.Move frmColors.ScaleWidth * (i - tmp) \ tmp, frmColors.ScaleHeight \ 2, frmColors.ScaleWidth * (i - tmp + 1&) \ tmp - frmColors.ScaleWidth * (i - tmp) \ tmp, (frmColors.ScaleHeight + 1&) \ 2&
        .Left = frmColors.ScaleWidth * (i - tmp) \ tmp
        .Top = frmColors.ScaleHeight \ 2
        .Width = frmColors.ScaleWidth * (i - tmp + 1&) \ tmp - frmColors.ScaleWidth * (i - tmp) \ tmp
        .Height = (frmColors.ScaleHeight + 1&) \ 2&
    End With
Next i
If frmColors.Visible Then frmColors.Refresh
End Sub

Private Sub mnuGrid_Click()
If Zm = 1 Then
    dbMsgBox 2537, vbExclamation
End If
DrawGrid
End Sub

Private Sub mnuReg_Click()
ShowStatus 10038, , 3
Load frmReg
ShowFormModal frmReg
Unload frmReg
ShowStatus STT_READY
End Sub

Private Sub mnuUnDo_Click()
If mnuUndo.Enabled And MeEnabled Then
    Undo
End If
End Sub

Private Sub MPMS_DblClick(Button As Integer, Shift As Integer, ix As Single, iy As Single)
CurPol.Active = False
End Sub

Private Sub MPMS_MouseDown(Button As Integer, Shift As Integer, ix As Single, iy As Single)
Dim x As Long, y As Long
Dim IsInRgn As Boolean
Dim Btn As Integer
Dim h As Integer
On Error GoTo eh
x = ix
y = iy
XOP = x
YOP = y
If ActiveTool = ToolPoly Then
  MPMS.CancelDblClick = False
Else
  MPMS.CancelDblClick = True
End If
SetWndDblClick MP.hWnd, Not MPMS.CancelDblClick
    UserMadeAction
    If Not MeEnabled Then Exit Sub
    'If NaviEnabled Then Exit Sub
    AutoScroll x, y, MD:=True
    If Button = 1 Or Button = 2 Or Button = 4 Then
        If Button = 4 Then Btn = 3 Else Btn = Button
        For h = 1 To 3
            If dbMS(h, 1) = dbButtonDown Then
                If (h = 1 And Button = 2) Or (h = 2 And Button = 1) Then
                    'MP_SpecialClick Button
                End If
                Exit Sub
            End If
        Next h
        dbMS(Btn, 1) = dbButtonDown
        dbMS(Btn, 2) = dbButtonDown
        
    End If
    
    
    IsInRgn = ((x >= 0) And (Int(x / Zm) <= intW - 1) And (y >= 0) And (Int(y / Zm) <= intH - 1)) Or (x0 = -1 Or y0 = -1)
        
    If Button = 4 And mnuUseWheel(0).Checked Then
        MainForm.ChangeActiveColor IIf(Shift = 0, 1, 2), _
                                   Data(x \ Zm, y \ Zm), _
                                   False
    End If
    
    If Button = 1 Or Button = 2 Then FileChanged = True
    
    fx0 = x / Zm
    fy0 = y / Zm
    
    Select Case ActiveTool()
    Case ToolLine, ToolFade 'lines
        If IsInRgn And (Button = 1 Or Button = 2) Then
            x0 = Int(x / Zm): y0 = Int(y / Zm)
            StartPixelAction
            DrawingLine = ActiveTool = ToolFade
            LineK = 1
            newLineK = 1
            LineStyle = 0
            'NewLineStyle = 0
        End If
    Case ToolStar, ToolFStar
        If IsInRgn And (Button = 1 Or Button = 2) Then
            If (x0 = -1 Or y0 = -1) Then
                x0 = x \ Zm
                y0 = y \ Zm
            End If
            StartPixelAction
        End If
    Case ToolColorSel
        Static SelTipShown As Boolean
        If IsInRgn And (Button = 1 Or Button = 2) Then
            ChangeActiveColor Button, Data(x \ Zm, y \ Zm), False
            tdN = 0
            tdRGB.rgbRed = 0
            tdRGB.rgbGreen = 0
            tdRGB.rgbBlue = 0
            
            x0 = 1
            y0 = 1
            
            If Not SelTipShown Then
                ShowStatus 10090, , 6 'For middle colors, click them with free button while holding another. The last is over which mouse is raised.
                FlashStatusBar 5
                SelTipShown = True
            End If
        End If
    Case ToolVFade 'v fade
        If IsInRgn And Button = 1 Then
            x0 = Int(x / Zm)
            y0 = Int(y / Zm)
            XO = x0
            YO = y0
            StartPixelAction
        ElseIf Button = 2 Then
            XO = Int(x / Zm) + IIf(x \ Zm > x0, -1, 1) 'for first event to work
            YO = Int(y / Zm)
            StartPixelAction
            MPMS_MouseMove Button, Shift, (x), (y)
        End If
    Case ToolPaint 'paint
        If IsInRgn And Button > 0 And Button <= 2 Then
            'StartPixelAction
            Fill Int(x / Zm), Int(y / Zm), ACol(Button) 'Acol(1) * Abs(Button = 1) + Acol(2) * Abs(Button = 2)
            XO = x \ Zm
            YO = y \ Zm
            'If MP.AutoRedraw Then MP.Refresh
        End If
    Case ToolCircle 'circle
        If IsInRgn And (Button = 1 Or Button = 2) Then
            x0 = x \ Zm
            y0 = y \ Zm
            StartPixelAction
            DrawingCircle = True
        End If
    Case ToolHFade 'h fade
        If IsInRgn And (Button = 1) Then
            x0 = Int(x / Zm)
            y0 = Int(y / Zm)
            XO = x0
            YO = y0
            StartPixelAction
        ElseIf IsInRgn And Button = 2 Then
            XO = x \ Zm
            YO = y \ Zm + IIf(y \ Zm > y0, -1, 1) ' for first event to work
            StartPixelAction
            MPMS_MouseMove Button, Shift, (x), (y)
        End If
    Case ToolPen
        If IsInRgn And (Button = 1 Or Button = 2) Then
            x0 = x \ Zm
            y0 = y \ Zm
            StartPixelAction
        End If
    Case ToolSel
        x = x \ Zm
        y = y \ Zm
        If Button = 1 Or Button = 2 Then
            If CurSel.Selected Then
                UpdateSelPic False, True
            End If
            CurSel.Selected = False
            SelPicture.Visible = False
            CurSel.x1 = x
            CurSel.y1 = y
            CurSel.x2 = x
            CurSel.y2 = y
            x0 = x
            y0 = y
            CurSel.Moving = True
            UpdateSelRect True
        End If
    Case ToolRect 'rectangle
        If IsInRgn And (Button = 1 Or Button = 2) Then
            x = x \ Zm
            y = y \ Zm
            x0 = x
            y0 = y
            XO = x
            YO = y
            'StartPixelAction
        End If
    Case ToolPoly 'polygon
        If IsInRgn And (Button = 1 Or Button = 2) Then
            MPMS.CancelDblClick = False
            x = x \ Zm
            y = y \ Zm
            If Not CurPol.Active Then
                x0 = x
                y0 = y
                XO = x
                YO = y
                CurPol.Active = True
                CurPol.Button = Button
                CurPol.bx = x
                CurPol.by = y
                StartPixelAction
            Else
                dbLine x0, y0, x, y, ACol(CurPol.Button), False
                x0 = x
                y0 = y
            End If
        End If
        
    Case ToolAir 'airbrush
        x0 = x \ Zm
        y0 = y \ Zm
        If (Button = 1 Or Button = 2) Then
            StartPixelAction
            If (frmAero.Chk(0).Value = 1) Then
                If frmAero.cOpt(0).Value Then
                    dbAero x0, y0, frmAero.aSize.Text, frmAero.aIntens.Text, ACol(Button)
                Else
                    dbAero x0, y0, frmAero.aSize.Text, frmAero.aIntens.Text, -1
                End If
                'If MP.AutoRedraw Then MP.Refresh
            End If
        End If
    Case ToolHelix
        If IsRgn(x, y, False) And (Button = 1 Or Button = 2) Then
            x0 = x \ Zm
            y0 = y \ Zm
            
            StartPixelAction
            
            XO = x0
            YO = y0
        End If
        
    Case ToolBrush
        RemoveTempPixels
        If IsRgn(x, y, False) And (Button = 1 Or Button = 2) Then
            x0 = x \ Zm
            y0 = y \ Zm
            StartPixelAction
            'dbBrushLine CurBrush, x0, y0, x0, y0, ACol(Button), False
        End If
    Case ToolPal 'pal
        x0 = 1
        y0 = 1
        If IsRgn(x, y, False) And (Button = 1) Then
            ChColBackColor LastPalIndex, Data(x \ Zm, y \ Zm)
            LastPalIndex = LastPalIndex + 1
            If LastPalIndex > UBound(ChCol) Then LastPalIndex = 0
            LastPickedColor = Data(x \ Zm, y \ Zm)
        End If
        XO = x \ Zm
        YO = y \ Zm
    
    Case ToolOrg 'Texture Origin
        x0 = x \ Zm
        y0 = y \ Zm
    
    Case ToolProg
        dblX0 = x / Zm
        dblY0 = y / Zm
        x0 = x \ Zm
        y0 = y \ Zm
        DrawingLine = True
        ScrollSettings.CancelWheelScroll = True
        TempDrawingMode = True
        ToolPrg_MouseEvent x, y, Button, Shift, dbEvMouseDown, GetPressureLevel
    End Select
    
Exit Sub
eh:
If Err.Number = dbCWS Then Exit Sub
MouseErr.Number = Err.Number
MouseErr.Source = Err.Source
MouseErr.Description = Err.Description
MessageBeep MB_ICONHAND
ShowStatus "ERROR!    " + MouseErr.Description, , 3
End Sub

Public Function InSel(ByVal x As Long, ByVal y As Long) As Boolean
InSel = CurSel.InSel(x, y) '(Not (x < CurSel.X1 Or y < CurSel.Y1 Or x > CurSel.X2 Or y > CurSel.Y2)) And CurSel.Selected
End Function

Friend Sub GetVisRect(ByRef rct As RECT, _
                      Optional ByVal UseMPPos As Boolean = False)
Dim MainRect As RECT
If UseMPPos Then
    rct.Top = -MP.Top  'VScroll.Value - 8&
    rct.Left = -MP.Left
Else
    rct.Top = VScroll.Value - 8
    rct.Left = HScroll.Value - 8
End If
rct.Right = rct.Left + MPHolder.ScaleWidth
rct.Bottom = rct.Top + MPHolder.ScaleHeight

If UseMPPos Then
    MainRect.Right = intW * Zm
    MainRect.Bottom = intH * Zm
    rct = IntersectRects(rct, MainRect)
End If
End Sub

Public Sub AutoScroll(ByVal x As Long, ByVal y As Long, Optional ByVal MD As Boolean = False)
Dim xf As Long, xt As Long
Dim yf As Long, yt As Long
Dim hPos As Long, vPos As Long
Dim xMoved As Boolean, yMoved As Boolean
Dim VisRct As RECT
    GetVisRect VisRct
    
    If MD Then
      'ScrollSettings.ASS_tmp = ScrollSettings.ASS
      With ScrollSettings
        'copy from original
        If MDPen Then
          .ASS_tmp = .ASS_pen
        Else
          .ASS_tmp = .ASS
        End If
        'shrink gaps if needed
        Dim CP As Long
        If .ASS_tmp.GapLef + .ASS_tmp.GapRig + 1 > Max(VisRct.Right - VisRct.Left, 0) Then
          CP = VisRct.Left + CDbl(VisRct.Right - VisRct.Left) * .ASS_tmp.GapLef / (.ASS_tmp.GapLef + .ASS_tmp.GapRig)
          .ASS_tmp.GapLef = CP - VisRct.Left
          .ASS_tmp.GapRig = VisRct.Right - 1 - CP
        End If
        If .ASS_tmp.GapTop + .ASS_tmp.GapBot + 1 > Max(VisRct.Bottom - VisRct.Top, 0) Then
          CP = VisRct.Top + CDbl(VisRct.Bottom - VisRct.Top) * .ASS_tmp.GapTop / (.ASS_tmp.GapTop + .ASS_tmp.GapBot)
          .ASS_tmp.GapTop = CP - VisRct.Top
          .ASS_tmp.GapBot = VisRct.Bottom - 1 - CP
        End If
        'fill tmp2 with gaps reduced to avoid scrolling on mousedown
        .ASS_tmp2.GapLef = Min(.ASS_tmp.GapLef, x - VisRct.Left)
        .ASS_tmp2.GapTop = Min(.ASS_tmp.GapTop, y - VisRct.Top)
        .ASS_tmp2.GapRig = Min(.ASS_tmp.GapRig, VisRct.Right - x - 1)
        .ASS_tmp2.GapBot = Min(.ASS_tmp.GapBot, VisRct.Bottom - y - 1)
      End With
    Else
      'enlarge gaps of tmp2 to whatever possible, but not larger than tmp.
      With ScrollSettings
        '.ASS_tmp.GapLef = Max(.ASS_tmp.GapLef, Min(.ASS.GapLef, x - VisRct.Left))
                          'enlarge only        not exceedeing either tmp or mouse ptr
        .ASS_tmp2.GapLef = Max(.ASS_tmp2.GapLef, Min(.ASS_tmp.GapLef, x - VisRct.Left))
        .ASS_tmp2.GapTop = Max(.ASS_tmp2.GapTop, Min(.ASS_tmp.GapTop, y - VisRct.Top))
        .ASS_tmp2.GapRig = Max(.ASS_tmp2.GapRig, Min(.ASS_tmp.GapRig, VisRct.Right - x - 1))
        .ASS_tmp2.GapBot = Max(.ASS_tmp2.GapBot, Min(.ASS_tmp.GapBot, VisRct.Bottom - y - 1))
      End With
    End If
    
    xf = VisRct.Left + ScrollSettings.ASS_tmp2.GapLef
    yf = VisRct.Top + ScrollSettings.ASS_tmp2.GapTop
    xt = VisRct.Right - ScrollSettings.ASS_tmp2.GapRig - 1
    yt = VisRct.Bottom - ScrollSettings.ASS_tmp2.GapBot - 1
    
    If x < xf Then
        hPos = x - xf
        xMoved = True
    End If
    If y < yf Then
        vPos = y - yf
        yMoved = True
    End If
    If x > xt Then
        'If xMoved Then
        '    hPos = x - (xf + xt) \ 2
        'Else
            hPos = x - xt
        'End If
        xMoved = True
    End If
    If y > yt Then
        'If yMoved Then
        '    vPos = y - (yf + yt) \ 2
        'Else
            vPos = y - yt
        'End If
        yMoved = True
    End If
    
    ScrollSettings.DontScroll = True
    If xMoved And HScrollEnabled Then
        ChangeScrollBarValue HScroll, hPos
    Else
        xMoved = False
    End If
    If yMoved And VScrollEnabled Then
        ChangeScrollBarValue VScroll, vPos
    Else
        yMoved = False
    End If
    ScrollSettings.DontScroll = False
    If xMoved Or yMoved Then
        ApplyScrollBarsValues
        'Refresh
    End If
End Sub

Private Sub MPMS_MouseMove(Button As Integer, Shift As Integer, ix As Single, iy As Single)
Dim Btn As Integer
Dim h As Integer
Dim Ded As Long, blniy As Boolean, xx As Currency, yy As Currency, i As Variant, j As Long
Dim xx2 As Long, yy2 As Long, Rtmp As Long, r1 As Single, r2 As Single
Dim Stt As String, SS As dbShiftConstants
'Dim xMoved As Boolean, yMoved As Boolean
Dim GTT As Long, cnt As Long
Dim x As Long, y As Long
Dim fx As Double, fy As Double
Dim FDSC As FadeDesc
Dim pColor1 As Long, pColor2 As Long
Dim rgb1 As RGBQuadLong
Dim ReallyMoved As Boolean
If MouseErr.Number <> 0 Then Exit Sub
If NaviEnabled Then
    NaviMove
    Exit Sub
End If
On Error GoTo eh
x = ix
y = iy

For h = 1 To 3
    If CBool(Button And Stepen2(h - 1)) Then
        If Not dbMS(h, 1) = dbButtonDown Then
            Exit Sub
        End If
    Else
        If dbMS(h, 1) = dbButtonDown Then
            Exit Sub
        End If
    End If
Next h

ReallyMoved = XOP <> x Or YOP <> y

SS = Shift * &H100
If (Button = 1 Or Button = 2 Or Button = 4) Then
    AutoScroll x, y
End If

If Button = 4 And IsRgn(x, y, False) And mnuUseWheel(0).Checked Then MainForm.ChangeActiveColor 1, Data(x \ Zm, y \ Zm), False

If Not (ActiveTool = 10 And (Button = 1 Or Button = 2)) Then
    Stt = grs(1205, "|1", CStr(Int(x / Zm)) + "," + CStr(Int(y / Zm)))
    If CurSel.Selected Then Stt = Stt + grs(1206, "|1", CStr(Abs(CurSel.x1 - CurSel.x2) + 1) + "x" + CStr(Abs(CurSel.y2 - CurSel.y1) + 1))
    If InSel(x \ Zm, y \ Zm) Then Stt = Stt + grs(1207, "|1", CStr(x \ Zm - CurSel.x1) + "," + CStr(y \ Zm - CurSel.y1))
    'ShowStatus Stt
End If
If Button = 3 Then Exit Sub
fx = x / Zm
fy = y / Zm
Select Case ActiveTool
Case ToolFStar 'Fade Star
    If (x0 = -1) Or (y0 = -1) Then Exit Sub
    If Button = 1 Or Button = 2 Then
        'MPMS_MouseUp Button, 100, iX, iY
        'dbMS(Button, 1) = dbButtonDown
        'dbMS(Button, 2) = dbButtonDown
        x = Int(x / Zm)
        y = Int(y / Zm)
        
        pColor1 = ACol(1)
        pColor2 = ACol(2)
        
        ValidatePerelivColors x0, y0, x, y, pColor1, pColor2
        
        dbFade x0, y0, x, y, pColor1, pColor2, gFDSC, , 0, 0, HighQ:=True, ForceDraw:=True
        
        If MP.AutoRedraw Then MP.Refresh
    End If
Case ToolStar 'Star
    If (x0 = -1) Or (y0 = -1) Then Exit Sub
    If Button = 1 Or Button = 2 Then
        'MPMS_MouseUp Button, 100, iX, iY
        'dbMS(Button, 1) = dbButtonDown
        'dbMS(Button, 2) = dbButtonDown
         dbLine x0, y0, x \ Zm, y \ Zm, ACol(Button)
        If MP.AutoRedraw Then MP.Refresh
    End If
    
Case ToolLine 'line
    If (x0 = -1) Or (y0 = -1) Then Exit Sub
    x = Int(x / Zm): y = Int(y / Zm)
    If (x <> XO Or y <> YO Or SS <> PrevSS) And (Button = 1 Or Button = 2) Then
        'removing temp line
        'dbLine x0, y0, XO, YO, -1, True, , , PrevSS
        RemoveTempPixels
        'draw new temp line
        dbLine x0, y0, x, y, ACol(Button), True, , , SS
        If MP.AutoRedraw Then MP.Refresh
    End If
       XO = x: YO = y
Case ToolFade 'fade
    'validating x and y
    If (x0 = -1) Or (y0 = -1) Then Exit Sub
    x = Int(x / Zm): y = Int(y / Zm)
    Stt = "(" + CStr(x) + ", " + CStr(y) + ")    " + GRSF(10091)
    If (Button = 1 Or Button = 2) And (fx <> fXO Or fy <> fYO Or SS <> PrevSS Or LineK <> newLineK Or LineStyle <> NewLineStyle) Then
    
        RemoveTempPixels
        
        LineK = newLineK
        LineStyle = NewLineStyle
        'drawing new temp fade
        pColor1 = ACol(1)
        pColor2 = ACol(2)
        dbFadeEx fx0, fy0, fx, fy, pColor1, pColor2, gFDSC, True, SS
        If MP.AutoRedraw Then MP.Refresh
        
    End If
    XO = x: YO = y
    fXO = fx
    fYO = fy

Case ToolPaint 'fill
        x = x \ Zm
        y = y \ Zm
        If IsRgn(x, y, False) And (x <> XO Or y <> YO) And Button > 0 And Button <= 2 Then
            XO = x
            YO = y
        End If
Case ToolCircle 'circle
    x = Int(x / Zm)
    y = Int(y / Zm)
    If Button = 1 Or Button = 2 Then
        'removing old circle
        RemoveTempPixels
        LineStyle = NewLineStyle
        If IsRgn(x, y, True) Then
            dbCircle x0, y0, x, y, ACol(1), ACol(2), gFDSC, True, CircleFlags
        End If
        If MP.AutoRedraw Then MP.Refresh
        
    End If
    XO = x
    YO = y
    
Case ToolVFade 'v fade
    Dim tmp As Long
    If (YO <> Int(y / Zm) Or XO <> Int(x / Zm)) And (Button = 1 Or Button = 2) Then
        x = Int(x / Zm)
        y = Int(y / Zm)
        
        pColor1 = ACol(1)
        pColor2 = ACol(2)
        
        If IsRgn(XO, YO, True) And IsRgn(x, y, True) Then
            Ded = Abs(y - YO)
            'tmp = Timer
            For i = 0 To Ded
                If Abs(GetTickCount - tmp) >= 100 Then
                    If MP.AutoRedraw Then
                        MP.Refresh
                        tmp = GetTickCount
                    End If
                End If
                If Not Ded = 0 Then xx = (i * (x - XO) / Ded + XO): yy = (i * (y - YO) / Ded + YO) Else xx = XO: yy = YO
                'y0 = yy
                ValidatePerelivColors x0, yy, xx, yy, pColor1, pColor2
                dbFade x0, yy, xx, yy, pColor1, pColor2, gFDSC, , 0, 0, False, True
                'dbMS(Button, 1) = 1
            Next i
            'dbMS(Button, 1) = dbButtonDown
            'dbMS(Button, 2) = dbButtonDown
        ElseIf IsRgn(x, y, True) Then
            'y0 = y
            ValidatePerelivColors x0, y, x, y, pColor1, pColor2
            dbFade x0, y, x, y, pColor1, pColor2, gFDSC, , 0, 0, False, True
            'dbMS(Button, 1) = dbButtonDown
            'dbMS(Button, 2) = dbButtonDown
        End If
        y0 = y
        XO = x
        YO = y
    End If

Case ToolHFade 'h fade
    
    If (YO <> Int(y / Zm) Or XO <> Int(x / Zm)) And (Button = 1 Or Button = 2) Then
        x = Int(x / Zm)
        y = Int(y / Zm)
        
        pColor1 = ACol(1)
        pColor2 = ACol(2)
        
        If IsRgn(XO, YO, True) And IsRgn(x, y, True) Then
            Ded = Abs(x - XO)
            For i = 0 To Ded
                If Abs(GetTickCount - tmp) >= 100 Then
                    If MP.AutoRedraw Then
                        MP.Refresh
                        tmp = GetTickCount
                    End If
                End If
                If Not Ded = 0 Then xx = Int(i * (x - XO) / Ded + XO): yy = Int(i * (y - YO) / Ded + YO) Else xx = XO: yy = YO
                ValidatePerelivColors xx, y0, xx, yy, pColor1, pColor2
                dbFade xx, y0, xx, yy, pColor1, pColor2, gFDSC, , 0, 0, False, True

            Next i
        ElseIf IsRgn(x, y, True) Then
            ValidatePerelivColors x, y0, x, y, pColor1, pColor2
            dbFade x, y0, x, y, pColor1, pColor2, gFDSC, , 0, 0, False, True
        End If
        x0 = x
        XO = x
        YO = y
    End If

Case ToolPen 'pen
    If Button = 1 Or Button = 2 Then
        MPMS_MouseUp Button, Shift, CSng(x), CSng(y)
        dbMS(Button, 1) = dbButtonDown
        dbMS(Button, 2) = dbButtonDown
    End If
    x0 = Int(x / Zm)
    y0 = Int(y / Zm)
Case ToolColorSel 'color detect
    Stt = "(" + CStr(x \ Zm) + ", " + CStr(y \ Zm) + ")    " + GRSF(10091)
    If InPicture(x, y) And (Button = 1 Or Button = 2) Then
        GetRgbQuadLongEx Data(x \ Zm, y \ Zm), rgb1
        rgb1.rgbRed = (rgb1.rgbRed + tdRGB.rgbRed) / (tdN + 1)
        rgb1.rgbGreen = (rgb1.rgbGreen + tdRGB.rgbGreen) / (tdN + 1)
        rgb1.rgbBlue = (rgb1.rgbBlue + tdRGB.rgbBlue) / (tdN + 1)
        ChangeActiveColor Button, RGB(rgb1.rgbRed, rgb1.rgbGreen, rgb1.rgbBlue), False
    End If
Case ToolSel 'selection
    If (Button = 1 Or Button = 2) Then
        x = Int(x / Zm)
        y = Int(y / Zm)
        CurSel.x1 = x0
        CurSel.y1 = y0
        CurSel.x2 = x
        CurSel.y2 = y
        UpdateSelRect True
         
    End If
Case ToolRect 'rectangle
    y = Int(y / Zm)
    x = Int(x / Zm)
    If (x <> XO Or y <> YO) And (Button = 1 Or Button = 2) Then
        'remove last rectangle
        RemoveTempPixels
        'draw a new rectangle
            dbRect x0, y0, x, y, ACol(Button), dbGrid, True, ACol(3 - Button), dbRS
        If MP.AutoRedraw Then MP.Refresh
    End If
    XO = x
    YO = y
Case ToolPoly 'polygon
    x = Int(x / Zm)
    y = Int(y / Zm)
    If IsRgn(x, y, True) And CurPol.Active Then
        If Button = 1 Or Button = 2 Then
            'draw new segment (not temp)
            dbLine x0, y0, x, y, ACol(CurPol.Button), False
            x0 = x
            y0 = y
        Else
            'remove temp
            RemoveTempPixels
            'draw new temp segment
            If IsRgn(x, y, True) Then
                dbLine x0, y0, x, y, ACol(CurPol.Button), True
            End If
            If MP.AutoRedraw Then MP.Refresh
        End If
    End If
    XO = x
    YO = y
Case ToolAir 'airbrush
        If (Button = 1 Or Button = 2) And (frmAero.Chk(1).Value = 1) Then
            If frmAero.cOpt(0).Value Then
                dbAero x \ Zm, y \ Zm, GetAirBrushSize, frmAero.aIntens.Text, ACol(Button)
            Else
                dbAero x \ Zm, y \ Zm, GetAirBrushSize, frmAero.aIntens.Text, -1
            End If
            If MP.AutoRedraw Then MP.Refresh
            MoveMouse
        End If
        
Case ToolHelix 'helix
    x = x \ Zm
    y = y \ Zm
    If (Button = 1 Or Button = 2) Then
        'remove old spin
        RemoveTempPixels
        
        'draw new
        If IsRgn(x, y, True) Then
            r2 = Sqr((x - x0) ^ 2 + (y - y0) ^ 2) / 2
            r1 = CountR1(r2)
            dbSpin CSng(x + x0) / 2, (y + y0) / 2, r1, r2, Val(frmSpin.txtN.Text), ACol(1), ACol(2), gFDSC, True
        End If
        If MP.AutoRedraw Then MP.Refresh
    End If
    XO = x
    YO = y
Case ToolBrush 'brush
    x = x \ Zm
    y = y \ Zm
    If Not MeEnabled Then Exit Sub
    If XO <> x Or YO <> y Then
        If (Button = 0) Then
            RemoveTempPixels 'dbBrushLine CurBrush, XO, YO, XO, YO, -1, True
        End If
        If IsRgn(x, y, True) Then
            If Button = 0 Then
                dbBrushLine CurBrush, x, y, x, y, ACol(1), DrawTemp:=True
                'If MP.AutoRedraw Then MP.Refresh'refresh is done by dbBrushLine
            ElseIf Button = 1 Or Button = 2 Then
                'If (Abs(XO - x) > 5 Or Abs(y - y0) > 5) And Not NeedRefr Then
                '    ShowStatus 10081, 2, 2
                '    NeedRefr = True
                'End If
                'if function returns true - means it called
                'DoEvents from within, so do not need to update
                xx = x0
                yy = y0
                'this function is called later!

                x0 = x
                y0 = y
                If dbBrushLine( _
                  CurBrush, _
                  xx, yy, x, y, _
                  ACol(Button), DrawTemp:=False, EnableTimeout:=True) Then MoveMouse
                
                'If MP.AutoRedraw Then MP.Refresh'refreshment is done by dbBrushLine
            End If
        End If
    Else 'if mouse pos unchanged
      If Button = 1 Or Button = 2 Then
        If dbBrushLine( _
                        CurBrush, _
                        xx, yy, x, y, _
                        ACol(Button), _
                        DrawTemp:=False, _
                        EnableTimeout:=True, _
                        DontAddPoint:=True) Then
          MoveMouse
        End If
      End If
    End If
    XO = x
    YO = y
    dbProcessMessages
    If Not PctCaptureDisabled Then DrawPctCapture x * Zm + Zm \ 2, y * Zm + Zm \ 2, True
Case ToolPal 'palette
    x = x \ Zm
    y = y \ Zm
    If IsRgn(x, y, True) And (Button = 1) And (x <> XO Or y <> YO) Then
        If LastPickedColor <> Data(x, y) Then
            ChColBackColor LastPalIndex, Data(x, y)
            LastPalIndex = LastPalIndex + 1
            If LastPalIndex > UBound(ChCol) Then LastPalIndex = 0
            LastPickedColor = Data(x, y)
        End If
    End If
    XO = x
    YO = y
    
Case ToolOrg 'Texture Origin
If Button = 1 Then
    If Not OrgUndoBuilt Then
        BUD
        OrgUndoBuilt = True
    End If
    MoveDataOrg Data, x0 - x \ Zm, y0 - y \ Zm
    TexOrg.x = TexOrg.x - (x0 - x \ Zm)
    TexOrg.y = TexOrg.y - (y0 - y \ Zm)
    x0 = x \ Zm
    y0 = y \ Zm
    
    DontDoEvents = True
    Refr
    DontDoEvents = False
    If Not TexMode Then
        mnuTexMode_Click
        ShowStatus 2408, , 3
        FlashStatusBar 6
    End If
End If
Case ToolProg
    ToolPrg_MouseEvent x, y, Button, Shift, dbEvMouseMove, GetPressureLevel
End Select
ShowStatus Stt
PrevSS = SS
If ScrollSettings.DS_Enabled And MP.AutoRedraw Then
    MoveMP
End If
If ReallyMoved Then
    PctCapture_Paint
End If
XOP = x
YOP = y
dbProcessMessages False
Exit Sub
eh:
If Err.Number = dbCWS Then Exit Sub
MouseErr.Number = Err.Number
MouseErr.Source = Err.Source
MouseErr.Description = Err.Description
MessageBeep MB_ICONHAND
ShowStatus "ERROR!    " + MouseErr.Description, , 2
End Sub

Public Sub dbPutSel(Optional ByVal SelMode As dbSelMode = dbSelMode.dbUseCurSelMode, _
                    Optional ByVal blnStore As Boolean = False, _
                    Optional ByVal UseSelPict As Boolean = True, _
                    Optional ByVal StoreUndo As Boolean = False)
UpdateSelPic UseSelPict, blnStore, SelMode, StoreUndo
End Sub

Friend Function MainRect() As RECT
Dim Rslt As RECT
With Rslt
    .Left = 0
    .Top = 0
    .Right = intW
    .Bottom = intH
End With
MainRect = Rslt
End Function

Friend Function RctSelRect() As RECT
Dim Rslt As RECT
With Rslt
    .Left = CurSel.x1
    .Top = CurSel.y1
    If CurSel.Selected Then
        .Right = CurSel.x2 + 1
        .Bottom = CurSel.y2 + 1
    Else
        .Right = .Left
        .Bottom = .Top
    End If
End With
RctSelRect = Rslt

End Function

Private Sub MPMS_MouseUp(Button As Integer, Shift As Integer, ix As Single, iy As Single)
Dim h As Integer
Dim Btn As Integer
Dim Ded As Long, blniy As Boolean, xx As Long, yy As Long, i As Variant, j As Long
Dim pColor As Long, g As Integer, InRgn As Boolean, xx2 As Long, yy2 As Long
Dim Radius As Single, sH As Long, sW As Long, r1 As Single
Dim SS As dbShiftConstants
Dim x As Long, y As Long
Dim FDSC As FadeDesc
Dim pColor1 As Long, pColor2 As Long
Dim rgb1 As RGBQuadLong
Dim x1 As Long, y1 As Long
Dim x2 As Long, y2 As Long
Dim fx As Double, fy As Double
Static wLastR As Long

'If NaviEnabled Then Exit Sub
UserMadeAction

x = ix
y = iy
If Button = 1 Or Button = 2 Or Button = 4 Then
    Btn = IIf(Button = 4, 3, Button)
    If (Button = 1 And dbMS(2, 1) = dbButtonDown) Or _
       (Button = 2 And dbMS(1, 1) = dbButtonDown) Then
        MP_SpecialClick Button, x \ Zm, y \ Zm
    End If
    For h = 1 To 3
        If Button = Stepen2(h - 1) And dbMS(h, 1) <> dbButtonDown Then
            Exit Sub
        End If
    Next h
    dbMS(Btn, 1) = dbButtonUp
    dbMS(Btn, 2) = dbButtonUp
    MDPen = False
End If

If MouseErr.Number <> 0 Then
    MsgBox MouseErr.Description, vbCritical, MouseErr.Source
    MouseErr.Number = 0
    Exit Sub
End If
On Error GoTo eh

SS = Shift * &H100&
g = Abs(mnuGrid.Checked)

InRgn = ((x >= 0) And (x \ Zm <= intW - 1) And (y >= 0) And (y \ Zm <= intH - 1))
    If Button = 4 And InRgn And mnuUseWheel(0).Checked Then MainForm.ChangeActiveColor 1, Data(x \ Zm, y \ Zm), False

If x0 = -1 Or y0 = -1 Then Exit Sub

DrawingLine = False

fx = x / Zm
fy = y / Zm

Select Case ActiveTool()
Case ToolLine, ToolStar, ToolPen
    x = x \ Zm
    y = y \ Zm
    If Button = 1 Or Button = 2 Then
        If ActiveTool = ToolLine Then  'line only
            RemoveTempPixels
            'dbLine x0, y0, XO, YO, -1, True, , , PrevSS
        End If
        If ActiveTool = ToolLine Then
            dbLine x0, y0, x, y, ACol(Button), , , , SS
        Else
            dbLine x0, y0, x, y, ACol(Button)
        End If
    End If
    
'---------------------------------------------------------------
Case ToolFade 'fade line
    DrawingLine = False
    RemoveTempPixels
    If (Button = 1 Or Button = 2) Then
        'validate x,y
        
        
        LineK = newLineK
        LineStyle = NewLineStyle
        pColor1 = ACol(1)
        pColor2 = ACol(2)
        dbFadeEx fx0, fy0, fx, fy, pColor1, pColor2, gFDSC, , SS, True
    End If
    
'---------------------------------------------------------------
Case ToolFStar, ToolVFade, ToolHFade  'fadestar,vfade,hfade
    If (Button = 1 Or Button = 2) And InRgn And IsRgn(x0, y0, True) Then
        'validate x,y
        x = x \ Zm
        y = y \ Zm
        If ActiveTool = 9 And Not x0 = x Then x0 = x
        
        If ActiveTool = 6 And Not (y0 = y) Then y0 = y
        
        pColor1 = ACol(1)
        pColor2 = ACol(2)
        
        ValidatePerelivColors x0, y0, x, y, pColor1, pColor2
        
        dbFade x0, y0, x, y, pColor1, pColor2, gFDSC, , 0, 0, HighQ:=ActiveTool = ToolFStar, ForceDraw:=True
    End If
    
'---------------------------------------------------------------
Case ToolPaint 'fill
    x = x \ Zm
    y = y \ Zm
    If InRgn And (x <> XO Or y <> YO) And Button > 0 And Button <= 2 Then
        'Fill x, y, ACol(Button), PSens    'Acol(1) * Abs(Button = 1) + Acol(2) * Abs(Button = 2)
        'If MP.AutoRedraw Then MP.Refresh
        XO = -1
        YO = -1
    End If
    
'---------------------------------------------------------------
Case ToolCircle 'circle
    RemoveTempPixels
    If InRgn And (Button = 1 Or Button = 2) Then
        x = x \ Zm
        y = y \ Zm
        'If IsRgn(XO, YO, True) Then
        '    dbCircle x0, y0, XO, YO, -1, -1, gFDSC, True, CircleFlags
        'End If

        LineStyle = NewLineStyle
        dbCircle x0, y0, x, y, ACol(1), ACol(2), gFDSC, False, CircleFlags
    End If
    
'---------------------------------------------------------------
Case ToolColorSel
    If InRgn And (Button = 1 Or Button = 2) Then
        GetRgbQuadLongEx Data(x \ Zm, y \ Zm), rgb1
        rgb1.rgbRed = (rgb1.rgbRed + tdRGB.rgbRed) / (tdN + 1)
        rgb1.rgbGreen = (rgb1.rgbGreen + tdRGB.rgbGreen) / (tdN + 1)
        rgb1.rgbBlue = (rgb1.rgbBlue + tdRGB.rgbBlue) / (tdN + 1)
        ChangeActiveColor Button, RGB(rgb1.rgbRed, rgb1.rgbGreen, rgb1.rgbBlue), False
        ChTool LUT
    End If
    
'---------------------------------------------------------------
Case ToolSel 'selection
    If Button = 1 Or Button = 2 Then
        x = x \ Zm
        y = y \ Zm
        CurSel.x1 = x0
        CurSel.y1 = y0
        CurSel.x2 = x
        CurSel.y2 = y
        CheckCurSelCoords
        GetSel
        CurSel.Moving = False
        CurSel.Selected = True
        UpdateSelRect False
        UpdateSelPic True, False
        SelPicture.Visible = True
    End If
    
'---------------------------------------------------------------
Case ToolRect 'rect
    x = x \ Zm
    y = y \ Zm
    RemoveTempPixels
    'If IsRgn(x, y, True) And (Button = 1 Or Button = 2) Then
    If dbRS <> dbEmpty Then
        StoreFragment Min(x, x0), Min(y, y0), Max(x, x0), Max(y, y0)
    Else
        StartPixelAction
    End If
    dbRect x0, y0, x, y, ACol(Button), , False, ACol(3 - Button), dbRS, , False
    'End If
    If MP.AutoRedraw Then MP.Refresh

    
'---------------------------------------------------------------
Case ToolAir 'airbrush
        If (Button = 1 Or Button = 2) And (frmAero.Chk(2).Value = 1) Then
            If frmAero.cOpt(0).Value Then
                dbAero x \ Zm, y \ Zm, frmAero.aSize.Text, frmAero.aIntens.Text, ACol(Button)
            Else
                dbAero x \ Zm, y \ Zm, frmAero.aSize.Text, frmAero.aIntens.Text, -1
            End If
            If MP.AutoRedraw Then MP.Refresh
        End If
            
'---------------------------------------------------------------
Case ToolHelix 'helix
    RemoveTempPixels
    If InRgn And (Button = 1 Or Button = 2) Then
        x = x \ Zm: y = y \ Zm
        Radius = Sqr((x - x0) ^ 2 + (y - y0) ^ 2) / 2
        r1 = CountR1(Radius)
        dbSpin (x0 + x) / 2, (y0 + y) / 2, r1, Radius, HSet.Numb, ACol(1), ACol(2), gFDSC, False
    End If
    
    
'---------------------------------------------------------------
Case ToolBrush 'brush

    x = x \ Zm
    y = y \ Zm
    If Not MeEnabled Then Exit Sub
    'If XO <> x Or YO <> y Then'now need to draw up everything left, whatever.
        If (Button = 0) Then
            RemoveTempPixels 'dbBrushLine CurBrush, XO, YO, XO, YO, -1, True
        End If
        'If IsRgn(x, y, True) Then'---//---
            If Button = 1 Or Button = 2 Then
                dbBrushLine CurBrush, _
                  x0, y0, x, y, _
                  ACol(Button), _
                  DrawTemp:=False, DontAddPoint:=XO = x And YO = y
                
                x0 = x
                y0 = y
                If MP.AutoRedraw Then MP.Refresh
            End If
        'End If
    'End If
    XO = x
    YO = y
    XO = XO + 1 'to force brush preview redraw
    MoveMouse

    If NeedRefr Then
        Refr
    End If
    
'---------------------------------------------------------------
Case ToolPal 'Palette
    x = x \ Zm
    y = y \ Zm
    If InPicture(x, y) And (Button = 1) And (x <> XO Or y <> YO) Then
        If LastPickedColor <> Data(x, y) Then
            ChColBackColor LastPalIndex, Data(x, y)
            LastPalIndex = LastPalIndex + 1
            If LastPalIndex > UBound(ChCol) Then LastPalIndex = 0
            LastPickedColor = Data(x, y)
        End If
    End If
    x0 = -1
    y0 = -1
    
'------------------------------------------------------------------
Case ToolOrg 'Texture Origin
    If Not OrgUndoBuilt Then
        BUD
        OrgUndoBuilt = True
    End If
    MoveDataOrg Data, x0 - x \ Zm, y0 - y \ Zm
    TexOrg.x = TexOrg.x - (x0 - x \ Zm)
    TexOrg.y = TexOrg.y - (y0 - y \ Zm)
    
    DontDoEvents = True
    Refr
    DontDoEvents = False
    If Not TexMode Then
        mnuTexMode_Click
        ShowStatus 2408, , 3
    End If
Case ToolProg
    DrawingLine = False
    ToolPrg_MouseEvent x, y, Button, Shift, dbEvMouseUp, GetPressureLevel
End Select
PrevSS = SS
ScrollSettings.CancelWheelScroll = False
If MP.AutoRedraw Then MP.Refresh
Exit Sub
eh:
If Err.Number = dbCWS Then Exit Sub
MsgError
End Sub

Private Sub MP_SpecialClick(ByVal Button As Integer, ByVal x As Long, ByVal y As Long)
Dim rgb1 As RGBQuadLong
If DrawingLine Or DrawingCircle Then
    NewLineStyle = LineStyle + 1
    MoveMouse
End If
If ActiveTool = 5 Then
    If InPicture(x, y) Then
        GetRgbQuadLongEx Data(x, y), rgb1
        tdN = tdN + 1
        tdRGB.rgbRed = tdRGB.rgbRed + rgb1.rgbRed
        tdRGB.rgbGreen = tdRGB.rgbGreen + rgb1.rgbGreen
        tdRGB.rgbBlue = tdRGB.rgbBlue + rgb1.rgbBlue
    End If
End If
End Sub

Public Sub SelPicture_SpecialClick(ByVal Button As Integer, ByVal x As Long, ByVal y As Long)
SelPicture_DblClick
End Sub

Friend Function getpoint(ByVal x As Double, ByVal y As Double) As Long
Dim dx As Double, dy As Double
Dim bx As Long, by As Long
Dim c1 As Long, c2 As Long, c3 As Long, c4 As Long
'c1 c2
'c3 c4
bx = Int(x)
by = Int(y)
dx = x - bx
dy = y - by
If InPicture(bx, by) Then
    c1 = Data(bx, by)
Else
    c1 = ACol(2)
End If
If dx > 0 Then
    If InPicture(bx + 1, by) Then
        c2 = Data(bx + 1, by)
    Else
        c2 = ACol(2)
    End If
End If
If dy > 0 Then
    If InPicture(bx, by + 1) Then
        c3 = Data(bx, by + 1)
    Else
        c3 = ACol(2)
    End If
End If
If dy > 0 And dx > 0 Then
    If InPicture(bx + 1, by + 1) Then
        c4 = Data(bx + 1, by + 1)
    Else
        c4 = ACol(2)
    End If
End If
c1 = dbAlphaBlend(c1, c2, dx * 255)
c3 = dbAlphaBlend(c3, c4, dx * 255)
getpoint = dbAlphaBlend(c1, c3, dy * 255)
End Function


Public Sub ValidatePerelivColors(ByVal x1 As Double, ByVal y1 As Double, _
                                 ByVal x2 As Double, ByVal y2 As Double, _
                                 ByRef LngColor1 As Long, ByRef LngColor2 As Long)
Dim i As Integer

If gFDSC.AutoColor1 Then
    LngColor1 = getpoint(x1, y1)
End If
If gFDSC.AutoColor2 Then
    LngColor2 = getpoint(x2, y2)
End If

End Sub

'procedure dbDeselect
'**************
'Removes current selection

Public Sub dbDeselect(blnApply As Boolean)
Dim i As Long, j As Long
Dim x1 As Long, y1 As Long, x2 As Long, y2 As Long
    If Not (CurSel.Selected) Then Exit Sub
    On Error GoTo eh
    SelRect.Visible = False
    
    If blnApply Then
        UpdateSelPic False, True
    End If
    With CurSel
        .Selected = False
        .Moving = False
    End With
    On Error GoTo 0
    With SelPicture
        .Visible = False
        .Move 0, 0, 1, 1
        .Cls
    End With
    
    CurSel.XM = -1
    CurSel.YM = -1
    
    Erase CurSel_SelData
    If NeedRefr Then Refr
Exit Sub
eh:
MsgError
End Sub

'procedure dbApply
'*****************
'Applies current selection to Data
'(does output it to screen)
'
Public Sub dbApply(ByVal SelMode As dbSelMode)
UpdateSelPic RedrawSelPic:=False, SetToData:=True, SelMode:=SelMode, StoreUndo:=True
End Sub

Public Function SWidth() As Long
If Toolbar.Visible Then
    SWidth = Me.ScaleWidth - (Toolbar.Width)
Else
    SWidth = Me.ScaleWidth
End If
End Function

Private Function ToAbsY(ByVal y As Long) As Long
Dim dh As Long
dh = 0
If ToolBar2Visible Then
    dh = dh + ToolBar2.Height
End If
ToAbsY = y + dh
End Function


Private Sub Form_Resize()
On Error GoTo eh
Dim lngTB As Long, En As Boolean
Static LastWS As FormWindowStateConstants
If Not dbTag = "" Then Exit Sub
If WindowState = vbMinimized Then Exit Sub
If WindowState = vbNormal And LastWS <> vbNormal Then
    LastWS = WindowState
    Me.Move Me.Left, Me.Top, FormW, FormH
    Exit Sub
ElseIf WindowState = vbNormal Then
    FormW = Me.Width
    FormH = Me.Height
End If
HScroll.Move (0), ToAbsY(SHeight - VScroll.Width), Me.SWidth - VScroll.Width, VScroll.Width
VScroll.Move (HScroll.Width), ToAbsY(0), VScroll.Width, SHeight - VScroll.Width
MPHolder.Move 0, ToAbsY(0), SWidth - VScroll.Width, SHeight - HScroll.Height
With ActiveColor(1)
.Move Me.SWidth - VScroll.Width, ToAbsY(SHeight - VScroll.Width), 2 * VScroll.Width \ 3, 2 * VScroll.Width \ 3
End With
With ActiveColor(2)
.Move Me.SWidth - 2 * VScroll.Width \ 3, ToAbsY(SHeight - 2 * VScroll.Width \ 3), 2 * VScroll.Width \ 3, 2 * VScroll.Width \ 3
End With
Refresh
Exit Sub
eh:
Refresh
Exit Sub
Resume
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim i As Integer, Answ As VbMsgBoxResult
If FileChanged Then
    If UnloadMode = VBRUN.QueryUnloadConstants.vbAppWindows Then
        BuildBackup , False
    Else
        If Me.WindowState = vbMinimized Then Me.WindowState = vbNormal
        Load frmConfirmExit
        With frmConfirmExit
            .Show vbModal
            If .Tag = "C" Then
                Cancel = True
                Unload frmConfirmExit
                Exit Sub
            End If
        End With
    End If
End If
dbTag = "Unloading..."
Me.Hide
DestroyHiResTimer
UnSubClassMP
UninstallHook
RestoreMP
#If pDebug Then
TestMdlArys
#End If
UpdateWH
dbSaveSetting "ImageSize", "Width", CStr(intW)
dbSaveSetting "ImageSize", "Height", CStr(intH)

dbSaveSetting "Colors", "FColor", "&H" + Hex$(ACol(1))
dbSaveSetting "Colors", "BColor", "&H" + Hex$(ACol(2))

dbSaveSetting "View", "Palette", CStr(mnuPal.Checked)
dbSaveSetting "View", "Grid", CStr(mnuGrid.Checked)
dbSaveSetting "View", "Zoom", CStr(Zm)
dbSaveSetting "View", "Toolbar", CStr(ToolBar2Visible)

'dbSaveSetting "View", "DynamicScrolling", CStr(IIf(MP.AutoRedraw, ScrollSettings.DS_Enabled, CBool(ScrollSettings.DS_EnableIfAR And &H1)))
'dbSaveSetting "View", "DynamicScrollingNoAR", CStr(IIf(MP.AutoRedraw, CBool(ScrollSettings.DS_EnableIfAR And &H2), ScrollSettings.DS_Enabled))
'dbSaveSetting "View", "ScrollHumidity", CStr(ScrollSettings.DS_Jestkost)
'dbSaveSetting "View", "ScrollEnL", CStr(ScrollSettings.DS_EnL)
'dbSaveSetting "View", "ScrollTimerRes", CStr(MoveTimerRes)
SaveScrollSettings

'dbsavesetting  "Options", "ColorsCount", CStr((ubound(chcol)+1))
dbSaveSetting "Options", "UndoSizeLimit", CStr(UndoSize)
dbSaveSetting "Options", "UseWheelButton", CStr(dbGetWheelUse)
dbSaveSetting "Options", "MaxPenPressure", CStr(MaxPenPressure)
For i = 0 To UBound(Steps)
    dbSaveSettingEx "Options", "MouseStep" + CStr(i), Steps(i)
    dbSaveSettingEx "Options", "MouseStepZm" + CStr(i), Edins(i)
Next i
'dbSaveSettingEx "Options", "MouseAttached", ScrollSettings.MouseGlued
'dbSaveSettingEx "Options", "AutoScrollFieldSize", CLng((ScrollSettings.ASS.GapLef + ScrollSettings.ASS.GapTop + ScrollSettings.ASS.GapRig + ScrollSettings.ASS.GapBot) / 4) 'AutoScroll_Field_Size
dbSaveSetting "Options", "DisableUndo", CStr(mnuNoUndoRedo.Checked)


dbSaveSetting "Tool", "LastUsedToolIndex", CStr(ActiveTool)
dbSaveSetting "Tool", "RectStyle", CStr(dbRS)
SaveFadeDesc gFDSC
SaveCurBrush CurBrush
SaveHSet HSet
dbSaveSetting "Tool", "CircleFlags", "&H" + Hex$(CircleFlags)
'dbSaveSetting "Tool", "LineFlags", "&H" + Hex$(LineFlags)
SaveLineFlags LineOpts
dbSaveSetting "Tool", "SelTransR", Trim(Str(CurSel.TransRatio))
dbSaveSetting "Tool", "SelStretchMode", Trim(Str(CurSel.StretchMode))
dbSaveSetting "Tool", "SelMode", Trim(Str(CurSel.SelMode))
SaveMatrix SelMatrix, "Tool", "SelMatrix"
FillOpts.TexOrigin.SaveToReg "Tool", "FillTexAlign"

SaveLensSettings

dbSaveSetting "MainForm", "WindowState", CStr(Me.WindowState)
dbSaveSetting "MainForm", "Height", CStr(FormH)
dbSaveSetting "MainForm", "Width", CStr(FormW)
If Me.WindowState = vbNormal Then
  dbSaveSetting "MainForm", "Left", CStr(Me.Left)
  dbSaveSetting "MainForm", "Top", CStr(Me.Top)
End If

dbSaveSetting "Dialog", "LastFolder", CStr(CDl.InitDir)

dbSaveSettingEx "Options", "PNG bits-per-pixel", CInt(pngBPP)

dbSaveSetting "Setup", "WasRun", "True"

FlushSettings
End Sub

Private Sub HScroll_Change()
If ScrollSettings.DontScroll Then Exit Sub
If ScrollSettings.DS_Enabled Then
    ApplyScrollBarsValues
Else
    HScroll_Scroll
End If
End Sub

Private Sub HScroll_Scroll()
    If ScrollSettings.DontScroll Then Exit Sub
    ApplyScrollBarsValues
End Sub

Private Sub mnuClear_Click()
ShowStatus "Clearing..."
BUD
'MP.Cls
ClearPic Data, ACol(2)
Refr

If CurSel.Selected Then
    dbPutSel
End If
ShowStatus GRSF(STT_READY)
End Sub

Private Sub mnuExit_Click()
Unload Me
End Sub

Private Sub mnuLoad_Click()
Dim File As String
ShowStatus GRSF(10002), , 3

On Error GoTo eh1
ShowStatus "$1209" 'loading
LoadAuto
On Error GoTo 0

ShowStatus GRSF(STT_READY)

Exit Sub

eh1:
    If Err.Number = dbCWS Then
        ShowStatus GRSF(STT_Cancelled)
        Exit Sub
    Else
        MsgError
        ShowStatus GRSF(STT_Error), , 3
    End If
Exit Sub
Resume
End Sub

Private Sub mnuPal_Click()
    mnuPal.Checked = Not (mnuPal.Checked)
    frmColors.Visible = mnuPal.Checked
    Form_Resize
End Sub

Private Sub mnuRefresh_Click()
If Not MeEnabled Then Exit Sub
Refr
End Sub

Private Sub mnuResize_Click()
Dim w As Long, h As Long
Dim Sz As Dims
ShowStatus 10007, , 3
On Error GoTo eh
Load Dialog
With Dialog
    Sz.w = intW
    Sz.h = intH
    .SetSz Sz
    .Show vbModal
    If .Tag <> "" Then GoTo ExitHere
    .ExtractSz Sz
    w = Sz.w
    h = Sz.h
    
    If (h <> intH) Or (w <> intW) Then
        If TexMode Then
'            If dbMsgBox(2407, vbYesNo) = vbNo Then 'origin information will be lost
'                ShowStatus STT_Cancelled
'                GoTo ExitHere
'            End If
        End If
        Resize w, h, CBool(.Chk.Value), , , .GetStretchMode
        FileChanged = True
    End If
End With
ExitHere:
Unload Dialog

Exit Sub
eh:
PushError
Unload Dialog
PopError
MsgError
End Sub

Function GetAttr(ByRef Color As Long, ByVal Index As Byte) As Byte
GetAttr = (Color And RGBMask(Index)) \ Stepen256(Index - 1)
End Function

Private Sub mnuTool_Click(Index As Integer)
Dim i As Integer

UserMadeAction

DrawingLine = False

ScrollSettings.CancelWheelScroll = False

Erase TransformData
'ShowHelp

ShowStatus 10041 + Val(mnuTool(Index).Tag), , 3
LUT = ActiveTool

If CurSel.Selected Then 'Если имеется выделение
    'StoreFragment CurSel.x1, CurSel.y1, CurSel.x2, CurSel.y2
    dbDeselect True 'уже делает undo
End If
If (LUT = ToolPoly) And (CurPol.Active) Then
    RemoveTempPixels
    
    If MP.AutoRedraw Then MP.Refresh
    CurPol.Active = False
End If
If LUT = ToolBrush Then
    If IsRgn(XO, YO, True) Then
        RemoveTempPixels 'dbBrushLine CurBrush, XO, YO, XO, YO, -1, True
    End If
End If
If LUT = ToolOrg Then
    OrgUndoBuilt = False
End If
If Index = ToolProg And LUT <> ToolProg Then
    MWM = 0
End If


For i = 0 To mnuTool.UBound
    'DebugMsg "Begin mnuTool(" + CStr(i) + ") Value:=" + CStr(Abs(i = Index))
    mnuTool(i).Checked = (i = Index)
    btnTool(i).Tag = "Changing"
    btnTool(i).Value = IIf(i = Index, vbChecked, vbUnchecked)
    btnTool(i).Tag = ""
    
Next i

If ActiveTool = ToolText Then '18
    ShowStatus "$10017", , 3
    On Error GoTo ehText
    OutText
    Exit Sub
End If
If ActiveTool = ToolProg Then
  On Error GoTo ehtp
    Load frmProgTool
    With frmProgTool
        'ShowHelp 11012
        '.Show vbModal
        .Tag = "c"
        .OkButton.RaiseClick
        'ShowHelp Me.HelpContextID
        If .Tag = "" Then
            .GetPrg prgToolProg
            MWM = prgToolProg.Vars(14).Value
            '.GetVars ToolVars
        Else
          mnuToolOpts_Click
        End If
    End With
    Unload frmProgTool
ehptrs:
End If
If Not (btnTool(Index).MouseIcon.Handle = 0) Then
    Set MP.MouseIcon = btnTool(Index).MouseIcon
    MP.MousePointer = 99
Else
    MP.MousePointer = 0
End If
x0 = -1: y0 = -1
mnuCopy.Enabled = (ActiveTool = 10)
If ActiveTool = 17 Then
    LastPalIndex = 0
End If
Exit Sub
ehtp:
Unload frmProgTool
mnuToolOpts_Click
Resume ehptrs
ehText:
  PushError
  ChTool ToolSel
  PopError
MsgError
End Sub

Function ActiveTool() As Integer
Attribute ActiveTool.VB_Description = "Returns active tool index"
Static at As Integer
Dim i As Integer
    If Not (mnuTool(at).Checked) Then
        For i = mnuTool.lBound To mnuTool.UBound
            If mnuTool(i).Checked Then
                ActiveTool = Val(mnuTool(i).Tag)
                at = i
                Exit For
            End If
        Next i
        If i = mnuTool.UBound + 1 Then ActiveTool = True
    Else
        ActiveTool = Val(mnuTool(at).Tag)
    End If
End Function

Private Sub mnuToolOpts_Click()
Dim tmp As Integer, ltmp As Long, i As Integer, tmp1 As Single, tmp2 As Long
Select Case ActiveTool
    Case 2 'Fade line
        Load frmLine
        With frmLine
            .SetProps LineOpts, gFDSC
            .Show vbModal
            If .Tag = "" Then
                .GetProps LineOpts, gFDSC
                LineStyle = 0
                NewLineStyle = 0
                FreshActiveColors
            End If
        End With
        Unload frmLine
    Case 4, 6, 9  'Fade
        ShowStatus "$10010", , 3
        'Pereliv.Counter.Text = "300"
        Load Pereliv
        With Pereliv
            SendFadeDesc gFDSC
            .Show vbModal
            If .Tag = "" Then
'                For i = 1 To 2
'                    If CBool(Pereliv.bColor(i - 1).Tag) Then ActiveColor(i).Caption = "" Else ActiveColor(i).Caption = "*"
'                Next i
                ExtractFadeDesc gFDSC
                FreshActiveColors
            End If
        End With
        Unload Pereliv
    Case 8
        Load frmCircle
        With frmCircle
            .SetProps CircleFlags, gFDSC
            .Show vbModal
            If .Tag = "" Then
                CircleFlags = .GetProps(gFDSC)
                FreshActiveColors
            End If
        End With
        Unload frmCircle
    Case ToolSel 'select
        Dim SelM As Long, StrM As Long, TransR As Single, TransC As Long
        ShowStatus "$10011", , 3
        SelM = CurSelMode
        TransR = dbCurSelTrR
        TransC = CurSel.TransColor
        StrM = CurSel.StretchMode
        
        SelOpts.SetProps SelM, StrM, TransC, TransR, SelMatrix, CurSel.IsText
        SelOpts.Show vbModal
        SelOpts.GetProps SelM, StrM, TransC, TransR, SelMatrix
        
        If (CurSelMode <> SelM) Or (TransR <> dbCurSelTrR) Or (TransC <> CurSel.TransColor) Or SelM = dbSelMode.dbMatrixMixed Or SelM = dbSelMode.dbSuperTransparent Then
            CurSel.SelMode = SelM
            CurSel.TransRatio = TransR
            CurSel.TransColor = TransC
            If CurSel.SelMode = dbSuperTransparent Then
                TransDataChanged = True
            End If
            dbPutSel
        End If
        CurSel.StretchMode = StrM
        
    Case 11 'rectangle
        ShowStatus "$10012", , 3
        ltmp = dbRS
        frmRect.rs = dbRS
        frmRect.Show vbModal
        If frmRect.Tag = "" Then dbRS = frmRect.rs
    Case 13 'Aero
        ShowStatus "$10013", , 3
        frmAero.Show vbModal
    Case 14 'spin
        ShowStatus "$10014", , 3
        Load frmSpin
        With frmSpin
            .SetProps HSet, gFDSC
            .Show vbModal
            If .Tag = "" Then
                .GetProps HSet, gFDSC
                FreshActiveColors
            End If
        End With
        Unload frmSpin
    Case 16 'Brush
        ShowStatus "$10015", , 3
        frmBrush.SetBrush CurBrush
        frmBrush.Show vbModal
        If frmBrush.Tag <> "" Then
            ShowStatus GRSF(STT_Cancelled)
            Exit Sub
        End If
        RemoveTempPixels 'dbBrushLine CurBrush, XO, YO, XO, YO, -1, True
        frmBrush.GetBrush CurBrush
    
    Case ToolPaint
        Load frmFill
        With frmFill
            .SetMode FillOpts
            .Show vbModal
            .GetMode FillOpts
        End With
        Unload frmFill
    Case ToolProg
        Load frmProgTool
        With frmProgTool
            ShowHelp 11012
            .Show vbModal
            ShowHelp Me.HelpContextID
            If .Tag = "" Then
                .GetPrg prgToolProg
                MWM = prgToolProg.Vars(14).Value
                ScrollSettings.CancelWheelScroll = CBool(prgToolProg.Vars(15).Value)
                '.GetVars ToolVars
            End If
        End With
        Unload frmProgTool
    Case Else
        ShowStatus "$10016", , 3
        dbMsgBox GRSF(1124), vbInformation
End Select
ShowStatus GRSF(STT_READY)
End Sub

Private Sub MP_OLEDragDrop(oData As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
Dim tmp As String

On Error GoTo eh
tmp = oData.Files(1)
If Len(tmp) = 0 Then Err.Raise dbCWS
On Error GoTo HE

LoadIntoSel tmp, x, y, True
eh:
Exit Sub
HE:
If Err.Number = dbCWS Then Exit Sub
dbMsgBox grs(1126, "|1", CStr(Err.Number), "|2", Err.Description), vbOKOnly Or vbCritical
End Sub

Public Sub SaveAuto(Optional ByVal ShowDialog As Boolean = False)
SaveSelTest
vtSavePicture Data, DataAlpha, OpenedFileName, OpenedFileFormatID, ShowDialog:=ShowDialog Or Len(OpenedFileName) = 0, Purpose:="MP"
FileChanged = False
End Sub

Public Sub LoadAuto()
Dim FileName As String
FileName = ShowPictureOpenDialog(FileName, Purpose:="MP")
LoadFile FileName
End Sub

Private Sub Form_OLEDragDrop(oData As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
Dim tmp As String
On Error GoTo HE
Select Case dbMsgBox(GRSF(1125), vbYesNoCancel Or vbQuestion Or vbMsgBoxSetForeground)
Case vbYes
    SaveAuto
Case vbCancel
    Exit Sub
End Select
On Error GoTo eh
tmp = oData.Files(1)
On Error GoTo HE
LoadFile tmp
Refr
eh:
Exit Sub
HE:
MsgError
End Sub

Private Sub MP_Paint()
If FreezeRefresh Then Exit Sub
If Len(dbTag) > 0 Then Exit Sub
Dim DrawRect As RECT
Dim dh As Long
On Error GoTo eh
FreezeRefresh = True
SetMousePtr False
UpdateWH
If intW = 0 Then GoTo ExitHere
GetVisRect DrawRect, True
DivideRect DrawRect, Zm
DrawData Data, Zm, MP.hDC, DrawRect
RedrawTempPixels
NeedRefr = False
ExitHere:
FreezeRefresh = False
Exit Sub
eh:
MsgError
FreezeRefresh = False
End Sub

Private Sub MP_Resize()
MPHolder_Resize
UpdateWH
ClearTempStorage
MP.Cls
End Sub

Private Sub MPHolder_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
Form_OLEDragDrop Data, Effect, Button, Shift, x, y
End Sub

Private Sub MPHolder_Resize()
Dim vMax As Long, hMax As Long
If Len(Me.Tag) > 0 Then Exit Sub
On Error Resume Next
MPHolder.PaintPicture gBackPicture, 0, 0, MPHolder.ScaleWidth, MPHolder.ScaleHeight
hMax = MP.Width + 16 - MPHolder.ScaleWidth
vMax = MP.Height + 16 - MPHolder.ScaleHeight
If hMax > 0 Then
    HScroll.Max = hMax
    HScroll.LargeChange = -Int(-MPHolder.ScaleWidth * 0.75)
    If Not HScrollEnabled Then
        HScrollEnabled = True
        HScroll.Enabled = HScrollEnabled
    End If
Else
    If HScrollEnabled Then
        HScrollEnabled = False
        HScroll.Enabled = HScrollEnabled
        HScroll.Max = 0
    End If
End If
If vMax > 0 Then
    VScroll.Max = vMax
    VScroll.LargeChange = -Int(-MPHolder.ScaleHeight * 0.75)
    If Not VScrollEnabled Then
        VScrollEnabled = True
        VScroll.Enabled = VScrollEnabled
    End If
Else
    If VScrollEnabled Then
        VScrollEnabled = False
        VScroll.Enabled = VScrollEnabled
        VScroll.Max = 0
    End If
End If
ApplyScrollBarsValues True
End Sub

Private Sub LoadLensSettings()
PctCaptureDisabled = Not dbGetSettingEx("Lens", "Enabled", vbBoolean, True)
PctCaptureZm = dbGetSettingEx("Lens", "Magnif", vbByte, 6)
If Not dbGetSettingEx("Lens", "Docked", vbBoolean, True) Then ShowLensWindow
End Sub

Private Sub SaveLensSettings()
dbSaveSettingEx "Lens", "Enabled", Not PctCaptureDisabled
dbSaveSettingEx "Lens", "Magnif", PctCaptureZm
dbSaveSettingEx "Lens", "Docked", Not LensWindowVisible
End Sub

Friend Sub PctCapture_DblClick()
pctCapture_MouseDown 1, 2, 0, 0
End Sub

Friend Sub pctCapture_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 1 And Shift = 4 Then
    PctCaptureDisabled = Not PctCaptureDisabled
    If PctCaptureDisabled Then pctCapture.Cls
ElseIf Button = 1 And Shift = 1 And Not PctCaptureDisabled Then
    PctCaptureZm = PctCaptureZm + 1
    If PctCaptureZm > 15 Then
        PctCaptureZm = 15
        vtBeep
    End If
    PctCapture_Paint
ElseIf Button = 2 And Shift = 1 And Not PctCaptureDisabled Then
    PctCaptureZm = PctCaptureZm - 1
    If PctCaptureZm < 1 Then
        PctCaptureZm = 1
        vtBeep
    End If
    PctCapture_Paint
    
ElseIf Button = 1 And Shift = 2 Then
    ToggleLensWindow
ElseIf Button = 2 And Shift = 0 Then
    On Error Resume Next
    ReposFrmLens -1
    mnuLnsDock.Caption = GRSF(IIf(LensWindowVisible, 326, 327))
    PopupMenu mnuPopLns, vbPopupMenuRightButton
    ReposFrmLens 1
End If
End Sub

Private Sub PctCapture_Paint()
Dim papi As POINTAPI
If PctCaptureDisabled Then Exit Sub
GetCursorPos papi
DrawPctCapture papi.x, papi.y
If pctCapture.AutoRedraw Then
    If LensWindowVisible Then
        frmLens.pctCapture.Cls
        DrawPctCapture papi.x, papi.y
        frmLens.pctCapture.Refresh
    Else
        pctCapture.Cls
        DrawPctCapture papi.x, papi.y
        pctCapture.Refresh
    End If
Else
    DrawPctCapture papi.x, papi.y
End If
End Sub

Private Sub DrawPctCapture(ByVal x As Long, ByVal y As Long, Optional ByVal RelToMP As Boolean = False)
Dim papi As POINTAPI
Dim DestHdc As Long
If x = -1 Then
    GetCursorPos papi
Else
    papi.x = x
    papi.y = y
End If
If RelToMP Then
    ToScreenCoords papi, MP.hWnd
End If
If PctCaptureZm = 0 Then PctCaptureZm = 6
If LensWindowVisible Then
    With frmLens.pctCapture
'        CaptureZoomIn papi, _
                      -Int(-.ScaleWidth / PctCaptureZm) + 1, _
                      -Int(-.ScaleHeight / PctCaptureZm) + 1, _
                      PctCaptureZm, _
                      .hDC, _
                       .ScaleWidth \ 2, .ScaleHeight \ 2
        CaptureZoomIn2 papi, _
                       .ScaleWidth \ 2, .ScaleHeight \ 2, _
                       PctCaptureZm, _
                       .hDC, _
                       .ScaleWidth \ 2, .ScaleHeight \ 2
        .PaintPicture ImgCursor.Picture, .ScaleWidth \ 2 - 15, .ScaleHeight \ 2 - 15
    End With
Else
    With pctCapture
'        CaptureZoomIn papi, _
                      -Int(-.ScaleWidth / PctCaptureZm) + 1, _
                      -Int(-.ScaleHeight / PctCaptureZm) + 1, _
                      PctCaptureZm, _
                      .hDC, _
                      .ScaleWidth \ 2, .ScaleHeight \ 2
        CaptureZoomIn2 papi, _
                       .ScaleWidth \ 2, .ScaleHeight \ 2, _
                       PctCaptureZm, _
                       .hDC, _
                       .ScaleWidth \ 2, .ScaleHeight \ 2
        
        .PaintPicture ImgCursor.Picture, .ScaleWidth \ 2 - 15, .ScaleHeight \ 2 - 15
    End With
End If
End Sub

Private Sub Picture1_Click()
ShowStatus 2440
End Sub

Private Sub Picture1_Paint()
Picture1.PaintPicture LoadResPicture(RES_BackPicture, vbResBitmap), 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight
End Sub

Private Sub Picture1_Resize()
Picture1_Paint
Status.Move Status.Left, 0, Picture1.ScaleWidth - 2 * Status.Left, Picture1.ScaleHeight
End Sub

Private Sub SelPicture_DblClick()
mnuSelAutoRepaint_Click
End Sub

Private Sub SelPicture_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim Btn As Long
Dim h As Long
If Not MeEnabled Then Exit Sub
UserMadeAction
If Not ActiveTool = 10 Then Exit Sub
    If Button = 1 Or Button = 2 Or Button = 4 Then
        If Button = 4 Then Btn = 3 Else Btn = Button
        For h = 1 To 3
            If dbMS(h, 1) = dbButtonDown Then
                If (h = 1 And Button = 2) Or (h = 2 And Button = 1) Then
                    'MP_SpecialClick Button
                    Exit Sub
                End If
            End If
        Next h
        dbMS(Btn, 1) = dbButtonDown
        dbMS(Btn, 2) = dbButtonDown
        
    End If

Select Case Button
Case 1, 4
    If CurSel.Selected Then
        CurSel.XM = x \ Zm
        CurSel.YM = y \ Zm
        CurSel.Moving = True
        If Shift = 2 Then
            dbPutSel blnStore:=True, StoreUndo:=True
        End If
    End If
Case 2
    If CurSel.Selected Then
        CurSel.XM = x \ Zm
        CurSel.YM = y \ Zm
        If Not (CurSel.XM = 0 Or CurSel.YM = 0) Then
            CurSel.Moving = True
            SelRect.Move SelPicture.Left, _
                         SelPicture.Top, _
                         SelPicture.Width, _
                         SelPicture.Height
            SelRect.Visible = True
            SelPicture.Visible = False
        Else
            CurSel.XM = -1
            CurSel.YM = -1
        End If
    End If
End Select

x = SelPicture.Left + x
y = SelPicture.Top + y
If (Button = 1 Or Button = 2 Or Button = 4) And Not (x = -1 Or y = -1) Then
    AutoScroll x, y, MD:=True
End If

End Sub

Private Sub SelPicture_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim Rtmp As Long, NeedRefresh As Boolean, Stt As String
Dim h As Long
If Not MeEnabled Then Exit Sub
For h = 1 To 3
    If CBool(Button And Stepen2(h - 1)) Then
        If Not dbMS(h, 1) = dbButtonDown Then
            Exit Sub
        End If
    Else
        If dbMS(h, 1) = dbButtonDown Then
            Exit Sub
        End If
    End If
Next h
If Not ActiveTool = 10 Then Exit Sub
    Stt = grs(1205, "|1", CStr(Int((x + SelPicture.Left) / Zm)) + "," + CStr(Int((y + SelPicture.Top) / Zm)))
    Stt = Stt + grs(1206, "|1", CStr(Abs(CurSel.x1 - CurSel.x2) + 1) + "x" + CStr(Abs(CurSel.y2 - CurSel.y1) + 1))
    Stt = Stt + grs(1207, "|1", CStr(x \ Zm) + "," + CStr(y \ Zm))
    ShowStatus Stt

If Button = 1 Then

    CurSel.x1 = (SelPicture.Left + x) \ Zm - CurSel.XM 'SelPicture.Left \ Zm
    CurSel.y1 = (SelPicture.Top + y) \ Zm - CurSel.YM 'SelPicture.Top \ Zm
    CurSel.x2 = CurSel.x1 + (SelPicture.Width) \ Zm - 1
    CurSel.y2 = CurSel.y1 + (SelPicture.Height) \ Zm - 1
    If SelPictureAutoRepaint Then
        CancelDoEvents
        dbPutSel
        RestoreDoEvents
        If MP.AutoRedraw Then MP.Refresh
    Else
        SelPicture.Move CurSel.x1 * Zm, CurSel.y1 * Zm
    End If
ElseIf Button = 2 Then
    If Not (CurSel.XM <= 0 Or CurSel.YM <= 0) Then
        MoveSelRect SelPicture.Left, _
                    SelPicture.Top, _
                    (1 + (Abs(CurSel.x2 - CurSel.x1)) * (x \ Zm) \ CurSel.XM) * Zm, _
                    (1 + (Abs(CurSel.y2 - CurSel.y1)) * (y \ Zm) \ CurSel.YM) * Zm
        AnimateSelRect
    End If
End If

x = SelPicture.Left + x
y = SelPicture.Top + y
If (Button = 1 Or Button = 2 Or Button = 4) And Not (x = -1 Or y = -1) Then
    AutoScroll x, y
End If


End Sub

Private Sub SelPicture_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim hgt As Long, Wdt As Long
Dim Btn As Long
Dim h As Long
If Not MeEnabled Then Exit Sub
If Button = 1 Or Button = 2 Or Button = 4 Then
    Btn = IIf(Button = 4, 3, Button)
    If (Button = 1 And dbMS(2, 1) = dbButtonDown) Or _
       (Button = 2 And dbMS(1, 1) = dbButtonDown) Then
        SelPicture_SpecialClick Button, x \ Zm, y \ Zm
    End If
    For h = 1 To 3
        If Button = Stepen2(h - 1) And dbMS(h, 1) <> dbButtonDown Then
            Exit Sub
        End If
    Next h
    dbMS(Btn, 1) = dbButtonUp
    dbMS(Btn, 2) = dbButtonUp
End If
If Not ActiveTool = 10 Then Exit Sub
If Button = 1 Or Button = 4 Then
    With SelPicture
        .Move .Left + (x \ Zm - CurSel.XM) * Zm, .Top + (y \ Zm - CurSel.YM) * Zm
    End With
    CurSel.x1 = SelPicture.Left \ Zm
    CurSel.y1 = SelPicture.Top \ Zm
    CurSel.x2 = (SelPicture.Left + SelPicture.Width) \ Zm - 1
    CurSel.y2 = (SelPicture.Top + SelPicture.Height) \ Zm - 1
    CurSel.Moving = False
    If Not CurSelMode = dbReplace Then dbPutSel
ElseIf Button = 2 Then
    If Not (CurSel.XM <= 0 Or CurSel.YM <= 0) Then
        SelRect.Visible = False
        SelPicture.Visible = True
        Wdt = (1 + (Abs(CurSel.x2 - CurSel.x1)) * (x \ Zm) \ CurSel.XM)
        hgt = (1 + (Abs(CurSel.y2 - CurSel.y1)) * (y \ Zm) \ CurSel.YM)
    
        dbStretch CurSel_SelData, _
                  Wdt, _
                  hgt, CurSel.StretchMode
        With CurSel
            .Moving = False
            .x1 = IIf(Wdt < 0, .x1 + Wdt, .x1)
            .y1 = IIf(hgt < 0, .y1 + hgt, .y1)
            .x2 = .x1 + Abs(Wdt) - 1
            .y2 = .y1 + Abs(hgt) - 1
        End With
        dbPutSel
    End If
End If
End Sub

Private Sub SelPicture_Paint()
If FreezeRefresh Then Exit Sub
On Error GoTo eh
FreezeRefresh = True
UpdateSelPic RedrawSelPic:=True, SetToData:=False
FreezeRefresh = False
Exit Sub
eh:
MsgError
FreezeRefresh = False
End Sub

Private Sub Status_Click()
ShowStatus 2440, , 3
End Sub

Private Sub TempBox_Click()
ShowStatus 10071, , 3
End Sub

Private Sub FormResizer_Timer()
Form_Resize
End Sub


Private Sub MessageLoopStarter_Timer()
'Registrator.Enabled = True
tmrCheckBackUp.Enabled = True
MessageLoopStarter.Enabled = False
MeEnabled = True
TakeMessagesControl
End Sub

Private Sub MPMover_Timer()
MoveMP
dbProcessMessages False
End Sub

Private Sub tmrCheckBackUp_Timer()
tmrCheckBackUp.Enabled = False
If GetLastBackupNumber > 0 Then
    Load frmBackupRestore
    With frmBackupRestore
        'Debug.Print "show"
        .Show vbModal
    End With
End If
End Sub

Private Sub tmrInstall_Timer()
tmrInstall.Enabled = False
mnuReg_Click
End Sub

Private Sub tmrMoveMouser_Timer()
MoveMouse MoveMouseDx, MoveMouseDy, Immediate:=True
tmrMoveMouser.Enabled = False
End Sub

Private Sub tmrStatusFlasher_Timer()
If StatusFlashesLeft = 0 Then
    tmrStatusFlasher.Enabled = False
    Exit Sub
End If
StatusFlashesLeft = StatusFlashesLeft - 1
If StatusFlashesLeft = 0 Then
    Picture1_Resize
    Picture1.Refresh
    tmrStatusFlasher.Enabled = False
Else
    Picture1.DrawMode = vbXorPen
    Picture1.Line (0, 0)-(Picture1.ScaleWidth, Picture1.ScaleHeight), &HFFFFFF, BF
    Picture1.DrawMode = vbCopyPen
End If
End Sub

Private Sub ToolBar_DblClick()
dbMsgBox 2438, vbInformation
End Sub

Private Sub ToolBar_Resize()
Dim CountX As Integer, CountY As Integer, y As Integer, x As Integer
Dim blnResized As Boolean, i As Long
On Error GoTo eh
If (dbTag <> "") Then Exit Sub
CountY = Toolbar.ScaleHeight \ btnTool(0).Height
CountX = Int(btnTool.Count / CountY + 1 - 0.0001)
blnResized = False

If Not Toolbar.Width = CountX * btnTool(0).Width + 4 Then
    blnResized = True
    Toolbar.Width = CountX * btnTool(0).Width + 4
    Form_Resize
    Exit Sub
End If

For i = 0 To btnTool.Count - 1
    With btnTool(btnTool.lBound + i)
        y = i \ CountX
        x = i Mod CountX
        If Not (.Top = y * .Height) Or Not (.Left = x * .Width) Then
            blnResized = True
            .Move x * .Width, y * .Height
        End If
    End With
Next i
y = ((btnTool.Count - 1) \ CountX + 1) * btnTool(0).Height
pctCapture.Move 0, y, Toolbar.ScaleWidth, Toolbar.ScaleHeight - y
On Error Resume Next
Toolbar.PaintPicture gBackPicture, 0, 0, Toolbar.ScaleWidth, Toolbar.ScaleHeight
On Error GoTo eh
eh:

End Sub

Private Sub ToolBar2_DblClick()
dbMsgBox 2439, vbInformation
End Sub

Private Sub ToolBar2_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then
    dbMsgBox 2439, vbInformation
End If
End Sub

Private Sub ToolBar2_Paint()
ToolBar2.PaintPicture gBackPicture, 0, 0, ToolBar2.ScaleWidth, ToolBar2.ScaleHeight
End Sub

Private Sub ToolBar2_Resize()
Static Ordered As Boolean
If Not Ordered Then
    OrderToolBarControls
    Ordered = True
End If
End Sub

Public Sub OrderToolBarControls()
Const Button_Width = 24
Const Button_Height = 24

Dim i As Integer, curX As Long
ToolBar2.Height = Button_Height + 4

ToolBarButton(0).Move 0, 0, Button_Width, Button_Height
curX = 0
For i = 1 To ToolBarButton.UBound
    On Error GoTo Skip
    
    If Len(ToolBarButton(i).dbTag2) = 0 Then
        curX = curX + Button_Width + Val(ToolBarButton(i).Tag) * 4
        ToolBarButton(i).Move curX, 0, Button_Width, Button_Height
    Else
        ToolBarButton(i).Visible = False
        ToolBarButton(i).ZOrder vbSendToBack
    End If
Skip:
Next i
End Sub

Private Sub ToolBarButton_Click(Index As Integer)
Dim strName As String
On Error GoTo eh
ExecuteCmd Val(ToolBarButton(Index).dbTag1), 0
On Error Resume Next
MP.SetFocus
Exit Sub
eh:
If Err.Number = dbCWS Then Exit Sub
MsgError
End Sub

Public Function GetTLBIndex(ByVal Act As dbCommands) As Long
Dim i As Long
Dim tmp As Long
Static tbl(0 To 500) As Long 'lookup table
'On Error Resume Next
GetTLBIndex = -1
If tbl(Act) = 0 Then
    For i = ToolBarButton.lBound To ToolBarButton.UBound
        tmp = Val(ToolBarButton(i).dbTag1)
        If tmp = Act Then
            tbl(Act) = i
            Exit For
        End If
    Next i
End If
GetTLBIndex = tbl(Act)
End Function

Private Sub ToolBarButton_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case 37 To 40
        Form_KeyDown KeyCode, Shift
        MP.SetFocus
End Select
End Sub

Private Sub tmrSelBorderAnimator_Timer()
If CurSel.Selected Then
    dbFreshSelBorder
    If SelPicture.AutoRedraw Then SelPicture.Refresh
End If
End Sub

Private Sub VScroll_Change()
'VScroll_Scroll
If ScrollSettings.DontScroll Then Exit Sub
If ScrollSettings.DS_Enabled Then
    ApplyScrollBarsValues
Else
    VScroll_Scroll
End If
End Sub

Private Sub VScroll_Scroll()
'If ScrollSettings.DontScroll Then Exit Sub
    If ScrollSettings.DontScroll Then Exit Sub
    ApplyScrollBarsValues
End Sub

Friend Sub ClearStorage(ByRef Storage As dbUndoStorage)
Erase Storage.uData
Storage.Index = -1
Storage.FirstIndex = 0
End Sub

Friend Sub StorePixelToStorage(ByRef puData As dbUndoData, ByVal x As Long, ByVal y As Long, Optional ByVal NextColor As Long = -1)
Dim i As Long, h As Long
Const DiscreteSize = 1000
On Error GoTo eh
If DisUndoRedo Then Exit Sub

i = puData.PixelsVUB
If i = 0 Then
    Err.Raise 1
    ReDim puData.Pixels(0 To 0)
    puData.PixelsVUB = 0
    puData.PixelsUB = 0
End If
If i > puData.PixelsUB Then
    puData.PixelsUB = puData.PixelsVUB + DiscreteSize
    ReDim Preserve puData.Pixels(0 To puData.PixelsVUB)
End If
puData.Pixels(i).x = x
puData.Pixels(i).y = y
puData.Pixels(i).c = Data(x, y)
puData.PixelsVUB = i + 1

Exit Sub


eh:
    ReDim puData.Pixels(0 To 0)
    puData.PixelsVUB = 0
    i = 0
    
Resume Next
End Sub

Friend Function CreateUndoEntry(ByRef Storage As dbUndoStorage, _
                                ByVal EntrySize As Long) As Long
Storage.Index = Storage.Index + 1
CreateUndoEntry = Storage.Index
If Storage.Index = 0 Then
    ReDim Storage.uData(0 To 0)
Else
    ReDim Preserve Storage.uData(0 To Storage.Index)
End If
TruncStorage Storage, EntrySize
End Function

'this function returns an error if AddSpace is larger than the limit.
Friend Sub TruncStorage(ByRef Storage As dbUndoStorage, _
                        ByVal AddSize As Long)
Dim i As Long
Dim LastToBeDeleted As Long
Dim SizeAccum As Long
Dim LeaveCurEntry As Boolean
LastToBeDeleted = -1
SizeAccum = AddSize
For i = Storage.Index To Storage.FirstIndex Step -1
    SizeAccum = SizeAccum + CalcEntrySize(Storage.uData(i))
    If SizeAccum > UndoSize Then
        Exit For
    End If
Next i
LastToBeDeleted = i
For i = Storage.FirstIndex To LastToBeDeleted
    With Storage.uData(i)
        If i = Storage.Index And .EntryType = dbUndoPixels Then
          'leave entrytype unchanged, discard all previous pixels
          LeaveCurEntry = True
        Else
          .EntryType = dbUndoInvalid
        End If
        Erase .d
        Erase .Pixels
        .PixelsUB = 0
        .PixelsVUB = 0
    End With
Next i
If LeaveCurEntry Then
  Storage.FirstIndex = Storage.Index
  ShowStatus 1210, HoldTime:=4 'Undo will be impossible: size limit exceeded.
  FlashStatusBar
Else
  Storage.FirstIndex = LastToBeDeleted + 1
End If
If AddSize > UndoSize Then
    Err.Raise 1111, "CreateUndoEntry", "Entry size exceeds limit. Please increase the limit!"
End If
End Sub
                        

Private Function CalcEntrySize(ByRef Entry As dbUndoData)
Dim Res As Long
Select Case Entry.EntryType
    Case dbUndoTypes.dbUndoFragment
        With Entry.Region
            Res = (.x2 - .x1 + 1) * (.y2 - .y1 + 1) * 4
            If Res < 0 Then Res = 0
        End With
    Case dbUndoTypes.dbUndoFull
        With Entry
            If AryDims(AryPtr(.d)) <> 2 Then
                Res = 0
            Else
                Res = (UBound(.d, 1) + 1) * (UBound(.d, 2) + 1) * 4
            End If
        End With
    Case dbUndoTypes.dbUndoInvalid
        Res = 0
    Case dbUndoTypes.dbUndoPixels
        Res = Entry.PixelsUB * 4& * 3&
    Case Else
        'may be forgotten
        Debug.Assert False
End Select
CalcEntrySize = Res
End Function

Friend Sub RemoveUndoEntry(ByRef Storage As dbUndoStorage)
Dim h As Long
'h = Storage.Index
'Erase Storage.uData(h).d
'Erase Storage.uData(h).Pixels
If Storage.Index < Storage.FirstIndex Or (Storage.uData(Storage.Index).EntryType = dbUndoInvalid) Then
    Err.Raise 115, "RemoveUndoEntry", "Cannot remove undo enty. No items left."
End If
Storage.Index = Storage.Index - 1
If Storage.Index = -1 Then
    ClearStorage Storage
Else
    ReDim Preserve Storage.uData(0 To Storage.Index)
End If

End Sub

Public Function CurUndoType() As dbUndoTypes
If UndoData.Index = -1 Then
    CurUndoType = dbUndoInvalid
Else
    CurUndoType = UndoData.uData(UndoData.Index).EntryType
End If
End Function

'Range and size checking is not needed
Friend Sub StoreFragmentTo(ByRef puData As dbUndoData, ByRef Frag As Rectangle)
Dim x As Long, y As Long
Dim x0 As Long, y0 As Long
Dim fx As Long, fy As Long
Dim tx As Long, ty As Long
Dim w As Long, h As Long
Dim i As Long
fx = Max(Frag.x1, 0)
fy = Max(Frag.y1, 0)
tx = Min(Frag.x2, intW - 1)
ty = Min(Frag.y2, intH - 1)
w = tx - fx + 1
h = ty - fy + 1
x0 = fx
y0 = fy

If w > 0 And h > 0 Then
    ReDim puData.d(0 To w - 1, 0 To h - 1)
    For y = 0 To h - 1
        For x = 0 To w - 1
            puData.d(x, y) = Data(x + x0, y + y0)
        Next x
    Next y
Else
    Erase puData.d
End If
puData.EntryType = dbUndoFragment
Erase puData.Pixels
puData.Region.x1 = fx
puData.Region.y1 = fy
puData.Region.x2 = tx
puData.Region.y2 = ty
End Sub

Friend Sub UnStoreFragmentFrom(ByRef puData As dbUndoData)
Dim x As Long, y As Long
Dim w As Long, h As Long
Dim t As Long, l As Long
Dim Frag As Rectangle
Dim i As Long
Frag = puData.Region
w = Frag.x2 - Frag.x1
h = Frag.y2 - Frag.y1
t = Frag.y1
l = Frag.x1
If w >= 0 And h >= 0 Then
    For y = 0 To h
        For x = 0 To w
'            If InPicture(x + l, y + t) Then
                Data(x + l, y + t) = puData.d(x, y)
'            End If
        Next x
    Next y
Else
    'Erase puData.d
End If
puData.EntryType = dbUndoFragment
Erase puData.Pixels
puData.Region = Frag
UpdateRegion Frag.x1, Frag.y1, Frag.x2, Frag.y2
If MP.AutoRedraw Then MP.Refresh
End Sub

Friend Sub PackPixels(ByRef rPixels() As dbPixel, ByRef MapPixels() As dbPixel)
Dim UB As Long
Dim i As Long
UB = -1
If AryDims(AryPtr(MapPixels)) = 1 Then
    UB = UBound(MapPixels)
End If
If UB = -1 Then
    Erase rPixels
Else
    rPixels = MapPixels
    For i = 0 To UB
        rPixels(i).c = Data(rPixels(i).x, rPixels(i).y)
    Next i
End If
End Sub

Friend Sub UnPackPixels(ByRef rPixels() As dbPixel)
Dim UB As Long
Dim i As Long
UB = -1
If AryDims(AryPtr(rPixels)) = 1 Then
    UB = UBound(rPixels)
End If
If UB > Max_Pixels_to_Draw Then
    For i = UB To 0 Step -1
        Data(rPixels(i).x, rPixels(i).y) = rPixels(i).c
    Next i
    NeedRefr = True
Else
    For i = UB To 0 Step -1
        Data(rPixels(i).x, rPixels(i).y) = rPixels(i).c
        dbPSet rPixels(i).x, rPixels(i).y, rPixels(i).c, dbAsmnuGrid, False, Not MP.AutoRedraw, True
    Next i
End If
End Sub


Friend Sub UndoEx(ByRef PUS As dbUndoStorage, ByRef PRS As dbUndoStorage)
Dim h As Long
Dim hr As Long
h = PUS.Index
Select Case PUS.uData(h).EntryType
    Case dbUndoTypes.dbUndoFull
        UpdateWH
        hr = CreateUndoEntry(PRS, intW * intH * 4&)
        With PRS.uData(hr)
            .EntryType = dbUndoFull
            .d = Data
            Erase .Pixels
            .Org = TexOrg
        End With
        
        Resize UBound(PUS.uData(h).d, 1) + 1, UBound(PUS.uData(h).d, 2) + 1, False, False, False
        Data = PUS.uData(h).d
        TexOrg = PUS.uData(h).Org
        Refr
        RemoveUndoEntry PUS
        
    Case dbUndoTypes.dbUndoFragment
        hr = CreateUndoEntry(PRS, CalcEntrySize(PUS.uData(h)))
        StoreFragmentTo PRS.uData(hr), PUS.uData(h).Region
        
        UnStoreFragmentFrom PUS.uData(h)
        RemoveUndoEntry PUS
    
    Case dbUndoTypes.dbUndoPixels
        hr = CreateUndoEntry(PRS, CalcEntrySize(PUS.uData(h)))
        With PRS.uData(hr)
            .EntryType = dbUndoPixels
            'remove unused entries
            If PUS.uData(h).PixelsVUB = 0 Then
                Erase PUS.uData(h).Pixels
                .PixelsUB = -1
                .PixelsVUB = 0
            Else
                ReDim Preserve PUS.uData(h).Pixels(0 To PUS.uData(h).PixelsVUB - 1)
                PUS.uData(h).PixelsUB = PUS.uData(h).PixelsVUB - 1
                'store pixels for restore
                PackPixels .Pixels, PUS.uData(h).Pixels
                .PixelsUB = UBound(.Pixels)
                .PixelsVUB = UBound(.Pixels) + 1
            End If
        End With
        UnPackPixels PUS.uData(h).Pixels
        RemoveUndoEntry PUS
        If NeedRefr Then
            Refr
        ElseIf MP.AutoRedraw Then
            MP.Refresh
        End If
        
    Case Else
        vtBeep
        Debug.Assert False
End Select
End Sub

Public Sub Undo()
If DisUndoRedo Then Exit Sub
Erase TransformData
UndoEx UndoData, RedoData
ValidateUndoRedo
OrgUndoBuilt = False
If CurSel.Selected Then
    dbPutSel
End If
End Sub

Public Sub Redo()
If DisUndoRedo Then Exit Sub
UndoEx RedoData, UndoData
ValidateUndoRedo
If CurSel.Selected Then
    dbPutSel
End If
End Sub

Public Sub BUD(Optional ByVal aryptrOldData As Long)
If DisUndoRedo Then Exit Sub
On Error Resume Next
StartUndo dbUndoFull, aryptrOldData
End Sub

Public Sub StartUndo(ByRef uType As dbUndoTypes, _
                     Optional ByVal aryptrOldData As Long = 0)
Dim h As Long
FileChanged = True
If Not UndoData.Index < UndoData.FirstIndex Then
  With UndoData.uData(UndoData.Index)
    If .EntryType = dbUndoPixels And .PixelsVUB = 0 Then
        .EntryType = dbUndoInvalid
        UndoData.Index = UndoData.Index - 1
    End If
  End With
End If
Select Case uType
    Case dbUndoTypes.dbUndoFull
        ClearStorage RedoData
        UpdateWH
        h = CreateUndoEntry(UndoData, intW * intH * 4&)
        With UndoData.uData(h)
            .EntryType = uType
            If aryptrOldData = 0 Then
                .d = Data
            Else
                SwapArys AryPtr(.d), aryptrOldData
            End If
            .Org = TexOrg
        End With
        OrgUndoBuilt = False
    Case dbUndoTypes.dbUndoFragment
        ClearStorage RedoData
        h = CreateUndoEntry(UndoData, 0)
        With UndoData.uData(h)
            .EntryType = uType
        End With
    Case dbUndoTypes.dbUndoPixels
        ClearStorage RedoData
        h = CreateUndoEntry(UndoData, 0)
        With UndoData.uData(h)
            .EntryType = uType
        End With
    Case Else
        Debug.Assert False
End Select
ValidateUndoRedo
End Sub

Public Sub TruncUndo(Optional ByVal AddSize As Long)
TruncStorage UndoData, AddSize
End Sub

Public Sub StartPixelAction()
If DisUndoRedo Then Exit Sub
StartUndo dbUndoPixels
End Sub

'Check for range before calling!!!
Public Sub StorePixel(ByVal x As Long, ByVal y As Long, Optional ByVal NextColor As Long = -1)
Const DiscreteSize = &H4000&
If DisUndoRedo Then Exit Sub

Dim oColor As Long
Dim UB As Long
#Const CheckUndoType = True
On Error GoTo eh
oColor = Data(x, y)
If NextColor = oColor Then Exit Sub
#If CheckUndoType Then
    If UndoData.uData(UndoData.Index).EntryType <> dbUndoPixels Then
        MsgBox "StorePixel without StartPixelAction", vbCritical
        Debug.Assert False
        StartPixelAction
        Exit Sub
    End If
#End If

With UndoData.uData(UndoData.Index)
UB = .PixelsVUB
If UB = 0 Then
    ReDim .Pixels(0 To 0)
    .PixelsUB = 0
End If
If UB > .PixelsUB Then
    .PixelsUB = UB + DiscreteSize
    ReDim Preserve .Pixels(0 To .PixelsUB)
    TruncUndo
End If
rsm:
.Pixels(UB).x = x
.Pixels(UB).y = y
.Pixels(UB).c = oColor
.PixelsVUB = UB + 1

End With

Exit Sub
Resume
eh:
UB = 0
ReDim UndoData.uData(UndoData.Index).Pixels(0 To 0)
Resume rsm
End Sub

'including x2,y2
Public Sub StoreFragment(ByVal x1 As Long, ByVal y1 As Long, _
                        ByVal x2 As Long, ByVal y2 As Long, _
                        Optional ByVal RaiseErrors As Boolean = False)
Dim Frag As Rectangle
On Error GoTo eh
If x2 < x1 Or y2 < y1 Then Exit Sub
FileChanged = True
If DisUndoRedo Then Exit Sub
With Frag
    .x1 = x1
    .y1 = y1
    .x2 = x2
    .y2 = y2
    TruncUndo (.x2 - .x1 + 1) * (.y2 - .y1 + 1) * 4
End With
StartUndo dbUndoFragment
StoreFragmentTo UndoData.uData(UndoData.Index), Frag
Exit Sub
eh:
If Err.Number = dbULE Then Exit Sub

If RaiseErrors Then ErrRaise "StoreFragment"
End Sub

Public Sub StoreSel(ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long)
StoreFragment x1, y1, x2, y2
End Sub

Friend Sub ClearUndo()
ClearStorage UndoData
ClearStorage RedoData
ValidateUndoRedo
End Sub

'**************Temp drawings functions*********************

'Range checking not needed
Public Sub StoreTempPixel(ByVal x As Long, ByVal y As Long, ByVal NextColor As Long)
Const DiscreteSize = 1000
'Dim UB As Long

On Error GoTo eh

If Not InPicture(x, y) Then Exit Sub
If NextColor = Data(x, y) Then Exit Sub

If TUPointer > TUSz - 1& Then
    If TUSz = 0& Then
        TUSz = DiscreteSize
        ReDim TUData(0 To TUSz - 1)
    Else
        TUSz = TUPointer + DiscreteSize + 1
        ReDim Preserve TUData(0 To TUSz - 1)
    End If
End If
rsm:
TUData(TUPointer).x = x
TUData(TUPointer).y = y
TUData(TUPointer).c = NextColor
TUPointer = TUPointer + 1&


Exit Sub
eh:
TUPointer = 0
ReDim TUData(0 To 0)
TUSz = 1
Resume rsm
End Sub

Public Sub RemoveTempPixels()
Dim i As Long
Dim x As Long, y As Long
Dim AR As Boolean
AR = MP.AutoRedraw
For i = TUPointer - 1 To 0 Step -1
    x = TUData(i).x
    y = TUData(i).y
    dbPsetFast x, y, Data(x, y), False, True, AR
Next i
TUPointer = 0
TUSz = 0
Erase TUData
End Sub

Public Sub RedrawTempPixels()
Dim i As Long
Dim x As Long, y As Long
Dim AR As Boolean
Dim fx As Long, tx As Long
Dim fy As Long, ty As Long
Dim wrect As RECT
GetVisRect wrect, True
DivideRect wrect, Zm
fx = wrect.Left
fy = wrect.Top
tx = wrect.Right - 1
ty = wrect.Bottom - 1
AR = MP.AutoRedraw
For i = 0 To TUPointer - 1
    x = TUData(i).x
    y = TUData(i).y
    If x >= fx And x <= tx And y >= fy And y <= ty Then
      dbPsetFast x, y, TUData(i).c, False, True, AR
    End If
Next i
End Sub


Public Sub ClearTempStorage()
TUPointer = 0
TUSz = 0
Erase TUData
End Sub



Public Function PerelivMode() As Integer
Dim i As Integer
For i = Pereliv.Opts.lBound To Pereliv.Opts.UBound
If Pereliv.Opts(i).Value Then PerelivMode = i: Exit Function
Next i
PerelivMode = -1
End Function

Public Sub Fill(ByVal x As Long, ByVal y As Long, _
                ByVal lngColor As Long)
Dim nR As Boolean
Dim OldIcon As Integer
Dim rct As RECT
Dim Pixels As typPixelList
Dim Pnt As POINTAPI
On Error GoTo eh
DisableMe
ShowStatus "Building area..."
mdlFill.vtCalcFillPixels x, y, FillOpts.BorderMode, FillOpts.Treshold, Data, Pixels, rct
ShowStatus "Storing undo..."
StoreFragment rct.Left, rct.Top, rct.Right, rct.Bottom
ShowStatus "Painting pixels..."
Select Case FillOpts.FillMode
    Case dbFillMode.FMSingleColor
        mdlFill.vtFillPixels Data, Pixels, lngColor
    Case dbFillMode.FMColorAlphaBlended
        mdlFill.vtFillAlphaBlend Data, Pixels, lngColor, FillOpts.FillAlpha
    Case dbFillMode.FMTextured
        If AryDims(AryPtr(FillOpts.Texture)) = 2 Then
            UpdateWH
            Pnt = FillOpts.TexOrigin.GetOffset(intW, intH, _
                       UBound(FillOpts.Texture, 1) + 1, _
                       UBound(FillOpts.Texture, 2) + 1, _
                       x, y)
            mdlFill.vtFillTexturize Data, Pixels, FillOpts.Texture, Pnt.x, Pnt.y
        Else
            mdlFill.vtFillPixels Data, Pixels, lngColor
        End If
End Select
ShowStatus "Updating view..."
UpdateRegion rct.Left, rct.Top, rct.Right, rct.Bottom
ShowStatus STT_READY
RestoreMeEnabled
Exit Sub
Resume
eh:
ClearMeEnabledStack
ErrRaise "Fill"
End Sub

Public Function IsRgn(ByVal x As Long, ByVal y As Long, ByVal Divided As Boolean) As Boolean
If Divided Then
IsRgn = ((x >= 0) And (y >= 0) And (x < intW) And (y < intH))
Else
IsRgn = ((x >= 0) And (y >= 0) And (x \ Zm < intW) And (y \ Zm < intH))
End If
End Function

Public Sub ChangePalCount(ByVal intCount As Integer)
Attribute ChangePalCount.VB_Description = "Changes item count in palette"
Dim i As Integer, tmpPal() As Pal_Entry
Dim OldLen As Integer
Dim tmp1 As String
Dim ResUB As Long
If intCount < 2 Or intCount > 512 Then Error 5
Randomize Timer
intCount = (intCount) - 1
If intCount < UBound(ChCol) Then
    tmpPal = ChCol
    ReDim ChCol(0 To intCount)
    For i = 0 To UBound(ChCol)
        ChCol(i).BackColor = tmpPal(i).BackColor
    Next i
    frmColors_Resize
ElseIf intCount > UBound(ChCol) Then
    OldLen = UBound(ChCol)
    ReDim Preserve ChCol(0 To intCount)
    
    tmp1 = dbGetSetting("DefaultPalette", "Colors", "")
    ResUB = Len(tmp1) \ 6 - 1
    
    For i = OldLen + 1 To intCount
        If i <= ResUB Then
            ChCol(i).BackColor = CLng("&H" + Mid$(tmp1, i * 6 + 1, 6))
        ElseIf i <= UBound(SMBDefPal) Then
            ChCol(i).BackColor = SMBDefPal(i)
        Else
            ChCol(i).BackColor = Int(Rnd(1) * &H1000000)
        End If
    Next i
    frmColors_Resize
End If
End Sub

Public Sub ChUndoSize(ByVal SizeInBytes As Long)
If SizeInBytes \ (1024& * 1024&) > 256 Then Err.Raise 5, "ChUndoSize", "Bad undo size"
On Error GoTo eh
UndoSize = SizeInBytes
ValidateUndoRedo
Exit Sub
eh:
MsgError "Failed to change undo count!"
End Sub

Friend Sub dbSimpleCircle_Ellipse( _
            ByVal CenterX As Single, ByVal CenterY As Single, _
            ByVal RadiusV As Single, ByVal RadiusH As Single, _
            ByVal lngColor As Long, _
            ByVal LngColor2 As Long, _
            ByRef FadeDsc As FadeDesc, _
            Optional ByVal DrawAsTemp As Boolean, _
            Optional ByVal Punktir As Long = &H0, _
            Optional ByVal HighQ As Boolean = True, _
            Optional ByVal ForceDraw As Boolean = True)
Dim i As Double, Ded As Long, xx As Long, yy As Long, g As GREnum
Dim xt As Long, yt As Long, j As Long, pColor As Long, h As Byte
Dim Radius As Single, PixCount As Long, AR As Boolean
Dim sxx As Single, syy As Single
Dim rgb1 As RGBQuadLong, rgb2 As RGBQuadLong
Dim CPer As Single, Stepen As Single, Ofc As Single
Dim PM As Single
Dim Div As Integer
'If Not DrawAsTemp Then g = Abs(mnuGrid.Checked) Else g = Abs(Not (Zm = 1))
If DrawAsTemp Then
    If Zm > 1 Then
        g = dbGrid
    Else
        g = dbNoGrid
    End If
Else
    If Zm > 1 Then
        g = dbAsmnuGrid
    Else
        g = dbNoGrid
    End If
End If
If lngColor = -1 Then
    DrawAsTemp = True
    LngColor2 = -1
End If
PM = FadeDsc.Mode
CPer = FadeDsc.FCount
Stepen = FadeDsc.Power
Ofc = FadeDsc.Offset
AR = MP.AutoRedraw
Radius = Max(Abs(RadiusV), Abs(RadiusH))
Ded = Int(2 * Pi * Radius) * (IIf(HighQ, 0, 2) + Sqr(2))
Debug.Assert LngColor2 <> -1 And lngColor <> -1
If Not (LngColor2 = -1) Then
    If Not lngColor = -1 Then
        GetRgbQuadLongEx lngColor, rgb1
    End If
    GetRgbQuadLongEx LngColor2, rgb2
    If HighQ And Not DrawAsTemp Then SetMousePtr True
    For j = 0 To Ded - 1
        i = 2 * Pi * j / (Ded - 1)
        If HighQ And Not DrawAsTemp Then
            sxx = Cos(i) * RadiusH + CenterX
            syy = Sin(i) * RadiusV + CenterY
            If IsRgn(sxx, syy, True) Then
                PixCount = PixCount + 1
                If Not CBool(Stepen2Long((PixCount) Mod 32) And Punktir) Then
                    i = CountJ(j / Ded, CPer, Stepen, Ofc, PM)
                    
                    pColor = RGB(i * (rgb2.rgbRed - rgb1.rgbRed) + rgb1.rgbRed, _
                                 i * (rgb2.rgbGreen - rgb1.rgbGreen) + rgb1.rgbGreen, _
                                 i * (rgb2.rgbBlue - rgb1.rgbBlue) + rgb1.rgbBlue)
                    dbPutPoint sxx, syy, pColor, DrawAsTemp, ForceDraw
                End If
            End If
        Else
            xx = Round(Cos(i) * RadiusH + CenterX)
            yy = Round(Sin(i) * RadiusV + CenterY)
            If xt <> xx Or yt <> yy Then
            If InPicture(xx, yy) Then
                PixCount = PixCount + 1
                If Not CBool(Stepen2Long((PixCount) Mod 32) And Punktir) Then
                    i = CountJ(j / Ded, CPer, Stepen, Ofc, PM)
                    
                    pColor = RGB(i * (rgb2.rgbRed - rgb1.rgbRed) + rgb1.rgbRed, _
                                 i * (rgb2.rgbGreen - rgb1.rgbGreen) + rgb1.rgbGreen, _
                                 i * (rgb2.rgbBlue - rgb1.rgbBlue) + rgb1.rgbBlue)
                    
                    dbPSet xx, yy, pColor, g, Not DrawAsTemp, Not AR
                End If
            End If
            xt = xx: yt = yy
            End If
        End If
    Next j
    If HighQ And Not DrawAsTemp Then SetMousePtr False
Else
End If
End Sub


Friend Sub dbCircle(ByVal x1 As Long, ByVal y1 As Long, _
                    ByVal x2 As Long, ByVal y2 As Long, _
                    ByVal lngColor As Long, _
                    ByVal LngColor2 As Long, _
                    ByRef FadeDsc As FadeDesc, _
                    Optional ByVal DrawAsTemp As Boolean, _
                    Optional ByVal Flags As dbCircleFlags = -1)
Dim cx As Single, cy As Single
Dim Rv As Single, rH As Single
Dim dx As Single, dy As Single
Select Case (Flags And &HF)
    Case 0 'on the diameter
        Select Case LineStyle Mod 4
            Case 0
                Rv = Sqr((x1 - x2) * (x1 - x2) + (y1 - y2) * (y1 - y2)) / 2
                rH = Rv
                cx = (x1 + x2) / 2
                cy = (y1 + y2) / 2
            Case 1
                Rv = Sqr((x1 - x2) * (x1 - x2) + (y1 - y2) * (y1 - y2)) / 2
                rH = -Rv
                cx = (x1 + x2) / 2
                cy = (y1 + y2) / 2
            Case 2
                Rv = -Sqr((x1 - x2) * (x1 - x2) + (y1 - y2) * (y1 - y2)) / 2
                rH = Rv
                cx = (x1 + x2) / 2
                cy = (y1 + y2) / 2
            Case 3
                Rv = -Sqr((x1 - x2) * (x1 - x2) + (y1 - y2) * (y1 - y2)) / 2
                rH = -Rv
                cx = (x1 + x2) / 2
                cy = (y1 + y2) / 2
        End Select
    Case 1 'on the radius
        Select Case LineStyle Mod 8
            Case 0
                Rv = Sqr((x1 - x2) * (x1 - x2) + (y1 - y2) * (y1 - y2))
                rH = Rv
                cx = x1
                cy = y1
            Case 1
                Rv = Sqr((x1 - x2) * (x1 - x2) + (y1 - y2) * (y1 - y2))
                rH = -Rv
                cx = x1
                cy = y1
            Case 2
                Rv = -Sqr((x1 - x2) * (x1 - x2) + (y1 - y2) * (y1 - y2))
                rH = -Rv
                cx = x1
                cy = y1
            Case 3
                Rv = -Sqr((x1 - x2) * (x1 - x2) + (y1 - y2) * (y1 - y2))
                rH = Rv
                cx = x1
                cy = y1
            Case 4
                Rv = Sqr((x1 - x2) * (x1 - x2) + (y1 - y2) * (y1 - y2))
                rH = Rv
                cx = x2
                cy = y2
            Case 5
                Rv = Sqr((x1 - x2) * (x1 - x2) + (y1 - y2) * (y1 - y2))
                rH = -Rv
                cx = x2
                cy = y2
            Case 6
                Rv = -Sqr((x1 - x2) * (x1 - x2) + (y1 - y2) * (y1 - y2))
                rH = -Rv
                cx = x2
                cy = y2
            Case 7
                Rv = -Sqr((x1 - x2) * (x1 - x2) + (y1 - y2) * (y1 - y2))
                rH = Rv
                cx = x2
                cy = y2
            End Select
    Case 2 'inside rectangle
        Select Case LineStyle Mod 2
            Case 0
                Rv = Abs(y1 - y2) / 2
                rH = Abs(x1 - x2) / 2
                cx = (x1 + x2) / 2
                cy = (y1 + y2) / 2
            Case 1
                Rv = Abs(y1 - y2)
                rH = Abs(x1 - x2)
                cx = x1
                cy = y1
        End Select
    Case 3 'outside rectangle
        Select Case LineStyle Mod 2
            Case 0
                cx = (x1 + x2) / 2
                cy = (y1 + y2) / 2
                rH = (x2 - x1) / Sqr(2)
                Rv = (y2 - y1) / Sqr(2)
            Case 1
                cx = x1
                cy = y1
                rH = (x2 - x1) * 2 / Sqr(2)
                Rv = (y2 - y1) * 2 / Sqr(2)
        End Select
        'Rv = (y2 - cY) * Sqr(((x2 - cX) ^ 2 + (y2 - cY) ^ 2) / (1 + (y2 - cY) ^ 4))
        'Rh = (x2 - cX) * Sqr(((x2 - cX) ^ 2 + (y2 - cY) ^ 2) / (1 + (x2 - cX) ^ 4))
End Select
dbSimpleCircle_Ellipse cx, cy, Rv, rH, lngColor, LngColor2, FadeDsc, DrawAsTemp, IIf(CBool(CircleFlags And dbCirclePunktir), &HFF00FF00, &H0&), CircleFlags And dbCircleHQ
If CBool(Flags And dbCirclePutCenter) Or DrawAsTemp Then
    If DrawAsTemp Then
        dbPSet cx, cy, lngColor, dbAsmnuGrid, Not DrawAsTemp, Not MP.AutoRedraw
    Else
        dbPutPoint cx, cy, lngColor, , True
    End If
End If
If CBool(Flags And dbCirclePutFocuses) Or DrawAsTemp Then
    If Abs(Rv) > Abs(rH) Then
        dy = Sqr(Rv * Rv - rH * rH)
        If DrawAsTemp Then
            dbPSet cx, cy + dy, lngColor, dbAsmnuGrid, Not DrawAsTemp, Not MP.AutoRedraw
            dbPSet cx, cy - dy, lngColor, dbAsmnuGrid, Not DrawAsTemp, Not MP.AutoRedraw
        Else
            dbPutPoint cx, cy + dy, lngColor, , True
            dbPutPoint cx, cy - dy, lngColor, , True
        End If
    ElseIf Abs(rH) > Abs(Rv) Then
        dx = Sqr(-Rv * Rv + rH * rH)
        If DrawAsTemp Then
            dbPSet cx - dx, cy, lngColor, dbAsmnuGrid, Not DrawAsTemp, Not MP.AutoRedraw
            dbPSet cx + dx, cy, lngColor, dbAsmnuGrid, Not DrawAsTemp, Not MP.AutoRedraw
        Else
            dbPutPoint cx - dx, cy, lngColor, , True
            dbPutPoint cx + dx, cy, lngColor, , True
        End If
    End If
End If
End Sub

Friend Sub dbSpin(ByVal CenterX As Single, ByVal CenterY As Single, _
                  ByVal r1 As Single, ByVal r2 As Single, _
                  ByVal n As Single, _
                  ByVal LngColor1 As Long, ByVal LngColor2 As Long, _
                  ByRef FadeDsc As FadeDesc, _
                  Optional ByVal DrawAsTemp As Boolean = False)
Dim i As Double, Ded As Long, xx As Long, yy As Long, g As GREnum
Dim xt As Long, yt As Long, j As Long, pColor As Long, h As Byte
Dim CPer As Long, Stepen As Single, Ofc As Single, sPerelivMode As Integer
Dim Radius As Double
Dim OldIcon As Integer, tmp1 As dbColor, tmp2 As dbColor
Dim XXxZM As Long, YYxZM As Long
Dim AR As Boolean
Dim FDSC As FadeDesc
If DrawAsTemp Then
    FDSC.AutoColor1 = False
    FDSC.AutoColor2 = False
    FDSC.FCount = 16
    FDSC.Mode = dbFSine
    FDSC.Power = 0.5
    If LngColor1 = -1 Then
        dbSimpleCircle_Ellipse CenterX, CenterY, Max(r1, r2), Max(r1, r2), -1, -1, FadeDsc, True
    Else
        dbSimpleCircle_Ellipse CenterX, CenterY, Max(r1, r2), Max(r1, r2), 0, vbWhite, FDSC, True
    End If
    Exit Sub
End If
If DrawAsTemp Then
    If Zm > 1 Then
        g = dbGrid
    Else
        g = dbNoGrid
    End If
Else
    If Zm > 1 Then
        g = dbAsmnuGrid
    Else
        g = dbNoGrid
    End If
End If
AR = MP.AutoRedraw
For h = 1 To 3
    tmp1.Comp(h) = GetAttr(LngColor1, h)
    tmp2.Comp(h) = GetAttr(LngColor2, h)
Next h
OldIcon = Screen.MousePointer
Screen.MousePointer = vbHourglass

If r1 > r2 Then dbSwap r1, r2
CPer = FadeDsc.FCount
Stepen = FadeDsc.Power
sPerelivMode = FadeDsc.Mode
Ofc = FadeDsc.Offset
Ded = Int(2 * Pi * r2) * (2 + Sqr(2)) * 2 * n
'g = Abs(mnuGrid.Checked)
If LngColor2 = -1 Then LngColor2 = LngColor1
If Not (LngColor1 = -1) Then
        For j = 0 To Ded - 1
            i = 2 * Pi * (j) / (Ded - 1) * n
            Radius = CDbl(j * (r2 - r1)) / (Ded - 1) + r1
            xx = Round(Cos(i) * Radius + CenterX)
            yy = Round(Sin(i) * Radius + CenterY)
            If IsRgn(xx, yy, True) Then
                If (xt <> xx Or yt <> yy) Then
                    i = CountJ(j / (Ded - 1), CPer, Stepen, Ofc, sPerelivMode)
                    XXxZM = xx * Zm
                    YYxZM = yy * Zm
                    
                    pColor = 0
                    For h = 1 To 3
                        pColor = pColor + Int(i * (CLng(tmp2.Comp(h)) - CLng(tmp1.Comp(h))) + tmp1.Comp(h)) * Stepen256(h - 1)
                    Next h
                    
'                    MP.Line (XXxZM + g, YYxZM + g)-(XXxZM + Zm - 1, YYxZM + Zm - 1), _
                            pColor, BF
                    dbPSet xx, yy, pColor, g, Not DrawAsTemp, Not AR
                    'If Not DrawAsTemp Then Data(xx, yy) = pColor
                    xt = xx: yt = yy
                End If
            End If
        Next j
Else
        For j = 0 To Ded - 1
            i = 2 * Pi * (j) / (Ded - 1) * n
            Radius = CDbl(j * (r2 - r1)) / (Ded - 1) + r1
            xx = Round(Cos(i) * Radius + CenterX)
            yy = Round(Sin(i) * Radius + CenterY)
            If IsRgn(xx, yy, True) Then
            If (xt <> xx Or yt <> yy) Then
                
                If IsRgn(xx, yy, True) Then
'                    MP.Line (xx * Zm + g, yy * Zm + g)-(xx * Zm + Zm - 1, yy * Zm + Zm - 1), _
                            Data(xx, yy), BF
                    dbPSet xx, yy, pColor, g, False, Not AR
                    xt = xx: yt = yy
                End If
            End If
            End If
        Next j

End If
Screen.MousePointer = OldIcon
End Sub


'Flags:
'&HF - all bits reserved for grid modes
'&H10 - do not store temp pixels (for the procedure that removes temp)
Public Sub dbPSet(ByVal x As Long, ByVal y As Long, _
                  ByVal lngColor As Long, _
                  Optional ByVal Flags As Long, _
                  Optional ByVal StoreToData As Boolean = False, _
                  Optional ByVal NoPsetOutOfScreen As Boolean = False, _
                  Optional ByVal ForceDraw As Boolean = True)
Dim g As Integer
Static Gr_Init As Boolean, hPen As Long, hBrush As Long, hPenDef As Long, hBrushDef As Long
Dim Grid As Long
    Grid = Flags And &HF&
    
    g = 0
    
    'If UseSelPict Then
    '    If Zm = 1& Then
    '        SetPixel SelPicture.hDC, X, Y, ConvertColorLng(lngColor)
    '    Else
    '        SelPicture.Line (X * Zm - SelPicture.Left + g, Y * Zm - SelPicture.Top + g)-(X * Zm + Zm - SelPicture.Left - 1, Y * Zm + Zm - SelPicture.Top - 1), lngColor, BF
    '    End If
    '    'Rectangle MP.hDC, X * Zm - SelPicture.Left + g, Y * Zm - SelPicture.Top + g, X * Zm + Zm - SelPicture.Left - 1, Y * Zm + Zm - SelPicture.Top - 1
    'Else
        If InPicture(x, y) Then
            If lngColor = -1 Then lngColor = Data(x, y)
            If ForceDraw Then
                If MP.AutoRedraw Then
                    If ForceDraw Then
                        FastPSet x, y, Zm, lngColor
                    End If
                Else
                    If Zm = 1 Then
                        SetPixel MP.hDC, x, y, ConvertColorLng(lngColor)
                        If x And &H100& And MainModule.VistaSetPixelBugDetected Then
                          If Not MP.AutoRedraw Then MP.Line (x, y)-(x, y), ConvertColorLng(lngColor), BF
                        End If
                    Else
                        MP.Line (x * Zm + g, y * Zm + g)-(x * Zm + Zm - 1, y * Zm + Zm - 1), ConvertColorLng(lngColor), BF
                    End If
                End If
            End If
            
            If StoreToData Then
                StorePixel x, y, lngColor
                Data(x, y) = lngColor
            ElseIf (Flags And &H10&) = 0& Then
                StoreTempPixel x, y, lngColor
            End If
        End If
        
    'End If 'if useselpict
End Sub

Public Sub dbLine(ByVal x1 As Long, ByVal y1 As Long, _
                  ByVal x2 As Long, ByVal y2 As Long, _
                  ByVal lngColor As Long, _
                  Optional ByVal TempLine As Boolean = False, _
                  Optional ByVal ForBrush As Boolean = False, _
                  Optional ByVal Draw As Boolean = True, _
                  Optional ByVal SS As dbShiftConstants = 0)
Dim i As Long, Ded As Long, g As GREnum, xx As Long, yy As Long
Dim X2_X1 As Long, Y2_Y1 As Long, XXxZM As Long, YYxZM As Long, IdivDED As Double
Dim AR As Boolean
If SS = -1 Then SS = GetShiftState
X2_X1 = x2 - x1
Y2_Y1 = y2 - y1
If SS = dbStateShift Then
    If Abs(x2 - x1) > Abs(y2 - y1) / 2 And Abs(y2 - y1) > Abs(x2 - x1) / 2 Then
        X2_X1 = Min(Abs(X2_X1), Abs(Y2_Y1)) * Sgn(X2_X1)
        Y2_Y1 = Abs(X2_X1) * Sgn(Y2_Y1)
    ElseIf Abs(x2 - x1) <= Abs(y2 - y1) / 2 Then
        X2_X1 = 0
    ElseIf Abs(y2 - y1) <= Abs(x2 - x1) / 2 Then
        Y2_Y1 = 0
    End If
End If
If Not TempLine Then g = Abs(mnuGrid.Checked) Else g = Abs(Not (Zm = 1))
If Not ((lngColor = -1) And (TempLine)) Then
    Ded = Max(Abs(X2_X1), Abs(Y2_Y1))
    For i = 0 To Ded
        If Not Ded = 0 Then
            If ForBrush Then
                xx = (i) * (X2_X1) \ Ded + x1
                yy = (i) * (Y2_Y1) \ Ded + y1
            Else
                xx = (i) * (X2_X1) / Ded + x1
                yy = (i) * (Y2_Y1) / Ded + y1
            End If
        Else
            xx = x1
            yy = y1
        End If
        If IsRgn(xx, yy, True) Then
            If Not (TempLine) Then
                If (Data(xx, yy) <> lngColor) And Draw Then
                    'XXxZM = xx * Zm
                    'YYxZM = yy * Zm
                    'MP.Line (XXxZM + g, YYxZM + g)-(XXxZM + Zm - 1, YYxZM + Zm - 1), lngColor, BF
                    dbPSet xx, yy, lngColor, dbAsmnuGrid, False, Not MP.AutoRedraw, True
                End If
                StorePixel xx, yy, lngColor
                Data(xx, yy) = lngColor
            Else
'                XXxZM = xx * Zm
'                YYxZM = yy * Zm
'                MP.Line (XXxZM + g, YYxZM + g)-(XXxZM + Zm - 1, YYxZM + Zm - 1), lngColor, BF
                dbPSet xx, yy, lngColor, dbGrid, False, True, True
            End If
        End If
    Next i
Else
    Ded = Max(Abs(X2_X1), Abs(Y2_Y1))
    For i = 0 To Ded
        If Not Ded = 0 Then
            If ForBrush Then
                xx = (i) * (X2_X1) \ Ded + x1
                yy = (i) * (Y2_Y1) \ Ded + y1
            Else
                xx = (i) * (X2_X1) / Ded + x1
                yy = (i) * (Y2_Y1) / Ded + y1
            End If
        Else
            xx = x1
            yy = y1
        End If
'            XXxZM = xx * Zm
'            YYxZM = yy * Zm
'            MP.Line (XXxZM + g, YYxZM + g)-(XXxZM + Zm - 1, YYxZM + Zm - 1), Data(xx, yy), BF
            dbPSet xx, yy, lngColor, dbGrid, False, Not MP.AutoRedraw, True
    Next i
End If
End Sub

Friend Sub dbFade(ByVal x1 As Double, ByVal y1 As Double, _
                  ByVal x2 As Double, ByVal y2 As Double, _
                  ByVal LngColor1 As Long, ByVal LngColor2 As Long, _
                  ByRef FadeDsc As FadeDesc, _
                  Optional ByVal TempFade As Boolean = False, _
                  Optional ByVal SS As dbShiftConstants = 0, _
                  Optional ByVal Punktir As Long = &H0&, _
                  Optional ByVal HighQ As Boolean = True, _
                  Optional ByVal ForceDraw As Boolean = True, _
                  Optional ByVal Weight1 As Double = -1, _
                  Optional ByVal Weight2 As Double = -1)
Dim Ded As Long, i As Long, j As Double, xx As Long, yy As Long
Dim g As GREnum, pColor As Long, h As Byte
Dim IdivDED As Double
Dim CPer As Single, Stepen As Double, Ofc As Single, PM As Integer
Dim X2_X1 As Double, Y2_Y1 As Double, XXxZM As Long, YYxZM As Long
Dim tmp1 As dbColor, tmp2 As dbColor
Dim AR As Boolean
Dim sxx As Double, syy As Double
On Error GoTo eh
'If Not TempFade Then g = Abs(mnuGrid.Checked) Else g = Abs(Not (Zm = 1))
If TempFade Then
    If Zm > 1 Then
        g = dbGrid
    Else
        g = dbNoGrid
    End If
Else
    If Zm > 1 Then
        g = dbAsmnuGrid
    Else
        g = dbNoGrid
    End If
End If
AR = MP.AutoRedraw

For h = 1 To 3
    tmp1.Comp(h) = GetAttr(LngColor1, h)
    tmp2.Comp(h) = GetAttr(LngColor2, h)
Next h
X2_X1 = x2 - x1
Y2_Y1 = y2 - y1
If SS = dbStateShift Then
    If Abs(x2 - x1) > Abs(y2 - y1) / 2 And Abs(y2 - y1) > Abs(x2 - x1) / 2 Then
        X2_X1 = Min(Abs(X2_X1), Abs(Y2_Y1)) * Sgn(X2_X1)
        Y2_Y1 = Abs(X2_X1) * Sgn(Y2_Y1)
    ElseIf Abs(x2 - x1) <= Abs(y2 - y1) / 2 Then
        X2_X1 = 0
    ElseIf Abs(y2 - y1) <= Abs(x2 - x1) / 2 Then
        Y2_Y1 = 0
    End If
End If
Ded = Round(Max(Abs(X2_X1), Abs(Y2_Y1)))
If HighQ And Not TempFade Then
    Ded = Sqr(X2_X1 ^ 2 + Y2_Y1 ^ 2)
End If

CPer = FadeDsc.FCount
Stepen = FadeDsc.Power
PM = FadeDsc.Mode
Ofc = FadeDsc.Offset
If HighQ Then
    dbFadeHQ x1, y1, x1 + X2_X1, y1 + Y2_Y1, LngColor1, LngColor2, FadeDsc, ForceDraw, TempFade, Weight1:=Weight1, Weight2:=Weight2
    Exit Sub
End If
If Not Ded = 0 Or HighQ Then
    For i = 0 To Ded
        IdivDED = CDbl(i) / IIf(Ded = 0, 1, Ded)
        If HighQ And Not (TempFade) Then
        Else
            xx = Round((i) * (X2_X1) / Ded + x1)
            yy = Round((i) * (Y2_Y1) / Ded + y1)
            If Not CBool(Punktir And Stepen2Long(i Mod 32)) Then
            If IsRgn(xx, yy, True) Then
'                XXxZM = xx * Zm
'                YYxZM = yy * Zm
                j = CountJ(IdivDED, CPer, Stepen, Ofc, PM)
                pColor = 0
                For h = 1 To 3
                    pColor = pColor + Int(j * (CLng(tmp2.Comp(h)) - CLng(tmp1.Comp(h))) + tmp1.Comp(h)) * Stepen256(h - 1)
                Next h
    '            MP.Line (XXxZM + g, YYxZM + g)-(XXxZM + Zm - 1, YYxZM + Zm - 1), pColor, BF
                dbPSet xx, yy, pColor, g, Not TempFade, Not AR
                'If Not TempFade Then Data(xx, yy) = pColor
            End If
            End If
        End If
    Next i
End If
eh:
End Sub

'TODO: use dbDrawPixels
Friend Sub dbFadeHQ(ByVal x0 As Double, ByVal y0 As Double, _
                    ByVal x1 As Double, ByVal y1 As Double, _
                    ByVal LngColor1 As Long, ByVal LngColor2 As Long, _
                    FadeDsc As FadeDesc, _
                    Optional ByVal ForceDraw As Boolean = True, _
                    Optional ByVal DrawTemp As Boolean = False, _
                    Optional ByVal Weight1 As Double, _
                    Optional ByVal Weight2 As Double)
Dim Pixels() As AlphaPixel
Dim nPixels As Long
Dim RGBData() As RGBQUAD

Dim Pnt1 As vtVertex, Pnt2 As vtVertex

Dim w As Long, h As Long
Dim x As Long, y As Long
Dim r As Long, g As Long, b As Long
Dim i As Long

Dim t1 As Long, t2 As Long
Dim t As Long

AryWH AryPtr(Data), w, h

If Weight2 = -1 Then
  With Pnt1
      .x = x0
      .y = y0
      .Weight = LineOpts.Weight * LineOpts.RelWeight1
      .Color = LngColor1
  End With
  With Pnt2
      .x = x1
      .y = y1
      .Weight = LineOpts.Weight * LineOpts.RelWeight2
      .Color = LngColor2
  End With
  If Weight1 >= 0 Then
      Pnt1.Weight = Weight1
      Pnt2.Weight = Weight1
  End If
Else
    With Pnt1
      .x = x0
      .y = y0
      .Weight = Weight1
      .Color = LngColor1
  End With
  With Pnt2
      .x = x1
      .y = y1
      .Weight = Weight2
      .Color = LngColor2
  End With
'  If Weight >= 0 Then
'      Pnt1.Weight = Weight
'      Pnt2.Weight = Weight
'  End If
End If
DrawingEngine.AntiAliasingSharpness = LineOpts.AntiAliasing

DrawingEngine.pntGradientLineHQ Pnt1, Pnt2, FadeDsc, Pixels, nPixels

On Error GoTo eh
ReferAry AryPtr(RGBData), AryPtr(Data)
    If TexMode Then
        For i = 0 To nPixels - 1
            x = Pixels(i).x Mod w
            If x < 0 Then x = x + w
            y = Pixels(i).y Mod h
            If y < 0 Then y = y + h
            
            b = RGBData(x, y).rgbBlue
            b = b + (Pixels(i).rgbBlue - b) * Pixels(i).drawOpacity \ 255
            g = RGBData(x, y).rgbGreen
            g = (g + (Pixels(i).rgbGreen - g) * Pixels(i).drawOpacity \ 255) * &H100
            r = RGBData(x, y).rgbRed
            r = (r + (Pixels(i).rgbRed - r) * Pixels(i).drawOpacity \ 255) * &H10000
            dbPSet x, y, r + g + b, StoreToData:=Not DrawTemp, ForceDraw:=ForceDraw
        Next i
    Else
        For i = 0 To nPixels - 1
            x = Pixels(i).x
            If x >= 0 And x < w Then
                y = Pixels(i).y
                If y >= 0 And y < h Then
                    b = RGBData(x, y).rgbBlue
                    b = b + (Pixels(i).rgbBlue - b) * Pixels(i).drawOpacity \ 255
                    g = RGBData(x, y).rgbGreen
                    g = (g + (Pixels(i).rgbGreen - g) * Pixels(i).drawOpacity \ 255) * &H100
                    r = RGBData(x, y).rgbRed
                    r = (r + (Pixels(i).rgbRed - r) * Pixels(i).drawOpacity \ 255) * &H10000
                    dbPSet x, y, r + g + b, StoreToData:=Not DrawTemp, ForceDraw:=ForceDraw
                End If
            End If
        Next i
    End If
UnReferAry AryPtr(RGBData)


Exit Sub
Resume
eh:
UnReferAry AryPtr(RGBData)
ErrRaise "dbFadeLineHQ"
End Sub

'only 1st vertex opacity is treated
'antialiasing is blur-width
'TODO: use dbDrawPixels
Friend Sub dbLine3opaq(Vtx1 As vtVertex3opaq, Vtx2 As vtVertex3opaq, _
                    AntiAliasing As Double, _
                    FadeDsc As FadeDesc, _
                    Optional ByVal ForceDraw As Boolean = True, _
                    Optional ByVal DrawTemp As Boolean = False)
Dim Pixels() As AlphaPixel
Dim nPixels As Long
Dim RGBData() As RGBQUAD

Dim Pnt1 As vtVertex, Pnt2 As vtVertex

Dim w As Long, h As Long
Dim x As Long, y As Long
Dim r As Long, g As Long, b As Long
Dim opR As Long, opG As Long, opB As Long
Dim i As Long

'Dim t1 As Long, t2 As Long
'Dim t As Long

AryWH AryPtr(Data), w, h

  LSet Pnt1 = Vtx1
  LSet Pnt2 = Vtx2
Dim tmpC As RGBQUAD
GetRgbQuadEx Vtx1.opaq, tmpC
opR = tmpC.rgbRed
opG = tmpC.rgbGreen
opB = tmpC.rgbBlue

DrawingEngine.AntiAliasingSharpness = 1 / AntiAliasing
DrawingEngine.pntGradientLineHQ Pnt1, Pnt2, FadeDsc, Pixels, nPixels

On Error GoTo eh
ReferAry AryPtr(RGBData), AryPtr(Data)
    If TexMode Then
        For i = 0 To nPixels - 1
            x = Pixels(i).x Mod w
            If x < 0 Then x = x + w
            y = Pixels(i).y Mod h
            If y < 0 Then y = y + h
            
            b = RGBData(x, y).rgbBlue
            b = b + (Pixels(i).rgbBlue - b) * Pixels(i).drawOpacity * opB \ 65025
            g = RGBData(x, y).rgbGreen
            g = (g + (Pixels(i).rgbGreen - g) * Pixels(i).drawOpacity * opG \ 65025) * &H100
            r = RGBData(x, y).rgbRed
            r = (r + (Pixels(i).rgbRed - r) * Pixels(i).drawOpacity * opR \ 65025) * &H10000
            dbPSet x, y, r + g + b, StoreToData:=Not DrawTemp, ForceDraw:=ForceDraw
        Next i
    Else
        For i = 0 To nPixels - 1
            x = Pixels(i).x
            If x >= 0 And x < w Then
                y = Pixels(i).y
                If y >= 0 And y < h Then
                    b = RGBData(x, y).rgbBlue
                    b = b + (Pixels(i).rgbBlue - b) * Pixels(i).drawOpacity * opB \ 65025
                    g = RGBData(x, y).rgbGreen
                    g = (g + (Pixels(i).rgbGreen - g) * Pixels(i).drawOpacity * opG \ 65025) * &H100
                    r = RGBData(x, y).rgbRed
                    r = (r + (Pixels(i).rgbRed - r) * Pixels(i).drawOpacity * opR \ 65025) * &H10000
                    dbPSet x, y, r + g + b, StoreToData:=Not DrawTemp, ForceDraw:=ForceDraw
                End If
            End If
        Next i
    End If
UnReferAry AryPtr(RGBData)


Exit Sub
Resume
eh:
UnReferAry AryPtr(RGBData)
ErrRaise "dbLine3opaq"
End Sub

Friend Sub dbDrawPixels(ByRef Pixels() As AlphaPixel, _
                      ByVal nPixels As Long, _
                      ByVal DrawTemp As Boolean, _
                      ByVal ForceDraw As Boolean, _
                      ByVal DrawMode As eDrawMode, _
                      Optional ByVal Opaq3ch As Long = &HFFFFFF, _
                      Optional ByVal KillAA As Boolean = False)
Dim RGBData() As RGBQUAD


Dim w As Long, h As Long
Dim x As Long, y As Long
Dim r As Long, g As Long, b As Long
Dim i As Long

Dim tmpC As RGBQUAD
GetRgbQuadEx Opaq3ch, tmpC
Dim opR As Long, opG As Long, opB As Long
opR = tmpC.rgbRed
opG = tmpC.rgbGreen
opB = tmpC.rgbBlue

AryWH AryPtr(Data), w, h

If KillAA Then
  For i = 0 To nPixels - 1
    Pixels(i).drawOpacity = 255 * Abs(Pixels(i).drawOpacity > 128)
  Next i
End If

On Error GoTo eh
ReferAry AryPtr(RGBData), AryPtr(Data)
' a huge list of drawing mode variants
Select Case DrawMode
  Case eDrawMode.dmNormal
    If TexMode Then
        For i = 0 To nPixels - 1
            x = Pixels(i).x Mod w
            If x < 0 Then x = x + w
            y = Pixels(i).y Mod h
            If y < 0 Then y = y + h
            
            b = RGBData(x, y).rgbBlue
            b = b + (Pixels(i).rgbBlue - b) * Pixels(i).drawOpacity \ 255
            g = RGBData(x, y).rgbGreen
            g = (g + (Pixels(i).rgbGreen - g) * Pixels(i).drawOpacity \ 255) * &H100
            r = RGBData(x, y).rgbRed
            r = (r + (Pixels(i).rgbRed - r) * Pixels(i).drawOpacity \ 255) * &H10000
            dbPSet x, y, r + g + b, StoreToData:=Not DrawTemp, ForceDraw:=ForceDraw
        Next i
    Else
        For i = 0 To nPixels - 1
            x = Pixels(i).x
            If x >= 0 And x < w Then
                y = Pixels(i).y
                If y >= 0 And y < h Then
                    b = RGBData(x, y).rgbBlue
                    b = b + (Pixels(i).rgbBlue - b) * Pixels(i).drawOpacity \ 255
                    g = RGBData(x, y).rgbGreen
                    g = (g + (Pixels(i).rgbGreen - g) * Pixels(i).drawOpacity \ 255) * &H100
                    r = RGBData(x, y).rgbRed
                    r = (r + (Pixels(i).rgbRed - r) * Pixels(i).drawOpacity \ 255) * &H10000
                    dbPSet x, y, r + g + b, StoreToData:=Not DrawTemp, ForceDraw:=ForceDraw
                End If
            End If
        Next i
    End If
  Case eDrawMode.dm3opaq
    If TexMode Then
        For i = 0 To nPixels - 1
            x = Pixels(i).x Mod w
            If x < 0 Then x = x + w
            y = Pixels(i).y Mod h
            If y < 0 Then y = y + h
            
            b = RGBData(x, y).rgbBlue
            b = b + (Pixels(i).rgbBlue - b) * Pixels(i).drawOpacity * opB \ 65025
            g = RGBData(x, y).rgbGreen
            g = (g + (Pixels(i).rgbGreen - g) * Pixels(i).drawOpacity * opG \ 65025) * &H100
            r = RGBData(x, y).rgbRed
            r = (r + (Pixels(i).rgbRed - r) * Pixels(i).drawOpacity * opR \ 65025) * &H10000
            dbPSet x, y, r + g + b, StoreToData:=Not DrawTemp, ForceDraw:=ForceDraw
        Next i
    Else
        For i = 0 To nPixels - 1
            x = Pixels(i).x
            If x >= 0 And x < w Then
                y = Pixels(i).y
                If y >= 0 And y < h Then
                    b = RGBData(x, y).rgbBlue
                    b = b + (Pixels(i).rgbBlue - b) * Pixels(i).drawOpacity \ 65025
                    g = RGBData(x, y).rgbGreen
                    g = (g + (Pixels(i).rgbGreen - g) * Pixels(i).drawOpacity \ 65025) * &H100
                    r = RGBData(x, y).rgbRed
                    r = (r + (Pixels(i).rgbRed - r) * Pixels(i).drawOpacity \ 65025) * &H10000
                    dbPSet x, y, r + g + b, StoreToData:=Not DrawTemp, ForceDraw:=ForceDraw
                End If
            End If
        Next i
    End If
  Case eDrawMode.dmMinimum
    If TexMode Then
        For i = 0 To nPixels - 1
            x = Pixels(i).x Mod w
            If x < 0 Then x = x + w
            y = Pixels(i).y Mod h
            If y < 0 Then y = y + h
            
            b = 255& - (255& - Pixels(i).rgbBlue) * Pixels(i).drawOpacity \ 255
            If b > RGBData(x, y).rgbBlue Then b = RGBData(x, y).rgbBlue
            g = 255& - (255& - Pixels(i).rgbGreen) * Pixels(i).drawOpacity \ 255
            If g > RGBData(x, y).rgbGreen Then g = RGBData(x, y).rgbGreen
            g = g * &H100
            r = 255& - (255& - Pixels(i).rgbRed) * Pixels(i).drawOpacity \ 255
            If r > RGBData(x, y).rgbRed Then r = RGBData(x, y).rgbRed
            r = r * &H10000
            dbPSet x, y, r + g + b, StoreToData:=Not DrawTemp, ForceDraw:=ForceDraw
        Next i
    Else
        For i = 0 To nPixels - 1
            x = Pixels(i).x
            If x >= 0 And x < w Then
                y = Pixels(i).y
                If y >= 0 And y < h Then
                    b = 255& - (255& - Pixels(i).rgbBlue) * Pixels(i).drawOpacity \ 255
                    If b > RGBData(x, y).rgbBlue Then b = RGBData(x, y).rgbBlue
                    g = 255& - (255& - Pixels(i).rgbGreen) * Pixels(i).drawOpacity \ 255
                    If g > RGBData(x, y).rgbGreen Then g = RGBData(x, y).rgbGreen
                    g = g * &H100
                    r = 255& - (255& - Pixels(i).rgbRed) * Pixels(i).drawOpacity \ 255
                    If r > RGBData(x, y).rgbRed Then r = RGBData(x, y).rgbRed
                    r = r * &H10000
                    dbPSet x, y, r + g + b, StoreToData:=Not DrawTemp, ForceDraw:=ForceDraw
                End If
            End If
        Next i
    End If
  Case eDrawMode.dmMaximum
    If TexMode Then
        For i = 0 To nPixels - 1
            x = Pixels(i).x Mod w
            If x < 0 Then x = x + w
            y = Pixels(i).y Mod h
            If y < 0 Then y = y + h
            
            b = Pixels(i).rgbBlue * Pixels(i).drawOpacity \ 255
            If b < RGBData(x, y).rgbBlue Then b = RGBData(x, y).rgbBlue
            g = Pixels(i).rgbGreen * Pixels(i).drawOpacity \ 255
            If g < RGBData(x, y).rgbGreen Then g = RGBData(x, y).rgbGreen
            g = g * &H100
            r = Pixels(i).rgbRed * Pixels(i).drawOpacity \ 255
            If r < RGBData(x, y).rgbRed Then r = RGBData(x, y).rgbRed
            r = r * &H10000
            dbPSet x, y, r + g + b, StoreToData:=Not DrawTemp, ForceDraw:=ForceDraw
        Next i
    Else
        For i = 0 To nPixels - 1
            x = Pixels(i).x
            If x >= 0 And x < w Then
                y = Pixels(i).y
                If y >= 0 And y < h Then
                    b = Pixels(i).rgbBlue * Pixels(i).drawOpacity \ 255
                    If b < RGBData(x, y).rgbBlue Then b = RGBData(x, y).rgbBlue
                    g = Pixels(i).rgbGreen * Pixels(i).drawOpacity \ 255
                    If g < RGBData(x, y).rgbGreen Then g = RGBData(x, y).rgbGreen
                    g = g * &H100
                    r = Pixels(i).rgbRed * Pixels(i).drawOpacity \ 255
                    If r < RGBData(x, y).rgbRed Then r = RGBData(x, y).rgbRed
                    r = r * &H10000
                    dbPSet x, y, r + g + b, StoreToData:=Not DrawTemp, ForceDraw:=ForceDraw
                End If
            End If
        Next i
    End If
  
End Select
UnReferAry AryPtr(RGBData)


Exit Sub
Resume
eh:
UnReferAry AryPtr(RGBData)
ErrRaise "dbLine3opaq"
End Sub

Public Sub dbPsetEx(ByVal x As Long, ByVal y As Long, _
                    ByVal lngColor As Long, ByVal Alpha As Long, _
                    Optional ByVal ForceDraw As Boolean = False, _
                    Optional ByVal TempPSet As Boolean = False)
'If Alpha = 0 Then Exit Sub
If Not InPicture(x, y) Then Exit Sub
    If Data(x, y) = lngColor And Not TempPSet Or Alpha = 0 Then Exit Sub
    dbPSet x, y, dbAlphaBlend(Data(x, y), lngColor, Alpha), dbNoGrid, Not TempPSet, Not MP.AutoRedraw, ForceDraw
    
End Sub

Friend Sub dbFadeEx(ByVal x1 As Double, ByVal y1 As Double, _
                    ByVal x2 As Double, ByVal y2 As Double, _
                    ByVal LngColor1 As Long, ByVal LngColor2 As Long, _
                    ByRef FadeDsc As FadeDesc, _
                    Optional ByVal TempFade As Boolean = False, _
                    Optional ByVal SS As dbShiftConstants = 0, _
                    Optional ByVal Obsolete As Long, _
                    Optional ByVal AutoDetectColors As Boolean = True)
Dim tx1 As Double, tx2 As Double, ty1 As Double, ty2 As Double
Dim dx As Double, dy As Double
Dim l As Double
Dim Punktir As Long
Dim HighQ As Boolean
Dim tFDSC As FadeDesc
Dim oFDSC As FadeDesc
'If LineFlags And dbLinePunktir Then
'    Punktir = &HFF00FF00
'End If
HighQ = True '((LineFlags And dbLineHQ) = dbLineHQ)
oFDSC = FadeDsc
If TempFade Then
'    oFDSC.Width = 1
End If
Select Case LineOpts.GeoMode
    Case eLineGeoMode.dbLineSimple
        dx = LineK * (x1 - x2)
        dy = LineK * (y1 - y2)
        Select Case LineStyle Mod 4
            Case 0
                tx1 = x1
                ty1 = y1
                tx2 = x2
                ty2 = y2
                ScrollSettings.CancelWheelScroll = False
            Case 1
                tx1 = (x1 + x2) / 2
                ty1 = (y1 + y2) / 2
                tx2 = tx1 + dy / 2
                ty2 = ty1 - dx / 2
                ScrollSettings.CancelWheelScroll = True
                SS = 0
            Case 2
                tx1 = (x1 + x2) / 2
                ty1 = (y1 + y2) / 2
                tx2 = tx1 - dy / 2
                ty2 = ty1 + dx / 2
                ScrollSettings.CancelWheelScroll = True
                SS = 0
            Case 3
                tx1 = (x1 + x2) / 2 + dy / 2
                ty1 = (y1 + y2) / 2 - dx / 2
                tx2 = tx1 - dy
                ty2 = ty1 + dx
                ScrollSettings.CancelWheelScroll = True
                SS = 0
        End Select
        If AutoDetectColors And LngColor1 <> -1 Then
            ValidatePerelivColors tx1, ty1, _
                                  tx2, ty2, _
                                  LngColor1, LngColor2
        End If
        dbFade tx1, ty1, tx2, ty2, LngColor1, LngColor2, oFDSC, TempFade, SS, Punktir, HighQ
        ScrollSettings.CancelWheelScroll = True
    
    
    Case eLineGeoMode.dbLineDouble
        dx = (x2 - x1) * LineK
        dy = (y2 - y1) * LineK
        If AutoDetectColors And LngColor1 <> -1 Then
            ValidatePerelivColors x1 - dx, y1 - dy, _
                                  x1 + dx, y1 + dy, _
                                  LngColor1, LngColor2
        End If
        SS = 0
        dbFade x1 - dx, y1 - dy, x1 + dx, y1 + dy, LngColor1, LngColor2, oFDSC, TempFade, SS, Punktir, HighQ
        ScrollSettings.CancelWheelScroll = True
        
    
    Case eLineGeoMode.dbLinePerp
        dx = (y1 - y2) * LineK
        dy = (x1 - x2) * LineK
        'tx1 = X1 - (Y1 - Y2) * LineK
        'ty1 = Y1 + (X1 - X2) * LineK
        If TempFade Then
            dbLine x1, y1, x2, y2, LngColor1, True, , , 0
        End If
        ScrollSettings.CancelWheelScroll = True
        SS = 0
        Select Case LineStyle Mod 6
            Case 0
                tx1 = x1
                ty1 = y1
                tx2 = x1 - dx
                ty2 = y1 + dy
            Case 1
                tx1 = x1
                ty1 = y1
                tx2 = x1 + dx
                ty2 = y1 - dy
            Case 2
                tx1 = x1 + dx
                ty1 = y1 - dy
                tx2 = x1 - dx
                ty2 = y1 + dy
            Case 3
                tx1 = x1
                ty1 = y1
                tx2 = x1 + dy
                ty2 = y1 - dx
            Case 4
                tx1 = x1
                ty1 = y1
                tx2 = x1 - dy
                ty2 = y1 + dx
            Case 5
                tx1 = x1 + dy
                ty1 = y1 - dx
                tx2 = x1 - dy
                ty2 = y1 + dx
        End Select
        If AutoDetectColors And LngColor1 <> -1 Then
            ValidatePerelivColors tx1, ty1, tx2, ty2, _
                                  LngColor1, LngColor2
        End If
        dbFade tx1, ty1, tx2, ty2, LngColor1, LngColor2, oFDSC, TempFade, 0, Punktir, HighQ
    
    
    Case eLineGeoMode.dbLineParallel
        l = Sqr((x2 - x1) * (x2 - x1) + (y2 - y1) * (y2 - y1))
        If l = 0 Then Exit Sub
        dx = 64# * (y1 - y2) / l
        dy = 64# * (x1 - x2) / l
        SS = 0
        If TempFade Then
            tFDSC.FCount = 20
            tFDSC.Mode = dbFSine
            tFDSC.Offset = 0
            tFDSC.Power = 0.5
            'tFDSC.Width = 1
            dbFadeHQ x2 - dx, y2 + dy, x2 + dx, y2 - dy, vbBlack, vbWhite, tFDSC, True, True
        End If
        dx = LineK * (y1 - y2)
        dy = LineK * (x1 - x2)
        'tx1 = x1 - y1 + Y
        'ty1 = y1 + x1 - X
        ScrollSettings.CancelWheelScroll = True
        Select Case LineStyle Mod 9
            Case 0
                tx1 = x1
                ty1 = y1
                tx2 = x1 - dx
                ty2 = y1 + dy
            Case 1
                tx1 = x1
                ty1 = y1
                tx2 = x1 + dx
                ty2 = y1 - dy
            Case 2
                tx1 = x1 + dx
                ty1 = y1 - dy
                tx2 = x1 - dx
                ty2 = y1 + dy
            Case 3
                tx1 = x1
                ty1 = y1
                tx2 = x1 + dy
                ty2 = y1 + dx
            Case 4
                tx1 = x1
                ty1 = y1
                tx2 = x1 - dy
                ty2 = y1 - dx
            Case 5
                tx1 = x1 + dy
                ty1 = y1 + dx
                tx2 = x1 - dy
                ty2 = y1 - dx
            Case 6
                tx1 = x1
                ty1 = y1
                tx2 = x1 + dy
                ty2 = y1 - dx
            Case 7
                tx1 = x1
                ty1 = y1
                tx2 = x1 - dy
                ty2 = y1 + dx
            Case 8
                tx1 = x1 + dy
                ty1 = y1 - dx
                tx2 = x1 - dy
                ty2 = y1 + dx
        End Select
        If AutoDetectColors And LngColor1 <> -1 Then
            ValidatePerelivColors tx1, ty1, tx2, ty2, _
                                  LngColor1, LngColor2
        End If
        dbFade tx1, ty1, tx2, ty2, LngColor1, LngColor2, oFDSC, TempFade, 0, Punktir, HighQ


End Select
End Sub

Friend Sub CompileFadeDscProg(FadeDsc As FadeDesc)
'Dim EV As New clsEVal
'With FadeDsc.Prog
'    ReDim .Vars(0 To 5)
'    .Vars(0).Name = "VPOS"
'    .Vars(1).Name = "RPOS"
'    .Vars(2).Name = "OFFSET"
'    .Vars(3).Name = "COUNT"
'    .Vars(4).Name = "DEGREE"
'    .Vars(5).Name = "POWER"
'    EV.MatVars .Vars
'    EV.CompileExpression_Ex .Source, .Code, .Vars
'End With
End Sub

Public Sub ChTool(ByVal ToolIndex As Integer)
Attribute ChTool.VB_Description = "Changes tool index"
Dim i As Integer
For i = mnuTool.lBound To mnuTool.UBound
If Val(mnuTool(i).Tag) = ToolIndex Then mnuTool_Click i: Exit For
Next i
End Sub

Public Sub dbRect(ByVal x1 As Long, ByVal y1 As Long, _
                  ByVal x2 As Long, ByVal y2 As Long, _
                  ByVal lngColor As Long, _
                  Optional ByVal Grid As GREnum = dbAsmnuGrid, _
                  Optional ByVal DrawAsTemp As Boolean = False, _
                  Optional ByVal bColor As Long, _
                  Optional ByVal dbStyle As dbRectStyle = dbEmpty, _
                  Optional ByVal ForceDraw As Boolean = True, _
                  Optional ByVal StoreUndo As Boolean = True)
Dim tmp As Long
Dim x As Long, y As Long
Dim fx As Long, tx As Long
Dim fy As Long, ty As Long

Dim OldIcon As Integer
OldIcon = Screen.MousePointer

If (DrawAsTemp And ForceDraw) Or dbStyle = dbEmpty Then
    dbLine x1, y1, x2, y1, lngColor, DrawAsTemp, False, ForceDraw
    dbLine x1, y1, x1, y2, lngColor, DrawAsTemp, False, ForceDraw
    dbLine x1, y2, x2, y2, lngColor, DrawAsTemp, False, ForceDraw
    dbLine x2, y1, x2, y2, lngColor, DrawAsTemp, False, ForceDraw
    Exit Sub
End If

fx = Min(x1, x2)
tx = Max(x1, x2)
fy = Min(y1, y2)
ty = Max(y1, y2)

x1 = fx
y1 = fy
x2 = tx
y2 = ty
fx = Max(0, x1 + 1)
fy = Max(0, y1 + 1)
tx = Min(intW - 1, x2 - 1)
ty = Min(intH - 1, y2 - 1)

'draw wire-frame
y = y1
If y >= 0 And y <= intH - 1 Then
    If StoreUndo Then
        For x = Max(0, x1) To Min(intW - 1, x2)
            StorePixel x, y, lngColor
            Data(x, y) = lngColor
        Next x
    Else
        For x = Max(0, x1) To Min(intW - 1, x2)
            Data(x, y) = lngColor
        Next x
    End If
End If
y = y2
If y >= 0 And y <= intH - 1 Then
    If StoreUndo Then
        For x = Max(0, x1) To Min(intW - 1, x2)
            StorePixel x, y, lngColor
            Data(x, y) = lngColor
        Next x
    Else
        For x = Max(0, x1) To Min(intW - 1, x2)
            Data(x, y) = lngColor
        Next x
    End If
End If
x = x1
If x >= 0 And x <= intW - 1 Then
    If StoreUndo Then
        For y = Max(0, y1) To Min(intH - 1, y2)
            StorePixel x, y, lngColor
            Data(x, y) = lngColor
        Next y
    Else
        For y = Max(0, y1) To Min(intH - 1, y2)
            Data(x, y) = lngColor
        Next y
    End If
End If
x = x2
If x >= 0 And x <= intW - 1 Then
    If StoreUndo Then
        For y = Max(0, y1) To Min(intH - 1, y2)
            StorePixel x, y, lngColor
            Data(x, y) = lngColor
        Next y
    Else
        For y = Max(0, y1) To Min(intH - 1, y2)
            Data(x, y) = lngColor
        Next y
    End If
End If

'fill
If dbStyle <> dbEmpty Then
    tmp = IIf(dbStyle = dbFilled, lngColor, bColor)
    If StoreUndo Then
        For y = fy To ty
            For x = fx To tx
                StorePixel x, y, tmp
                Data(x, y) = tmp
            Next x
        Next y
    Else
        For y = fy To ty
            For x = fx To tx
                Data(x, y) = tmp
            Next x
        Next y
    End If
End If

If ForceDraw Then
    UpdateRegion x1, y1, x2, y2
End If

End Sub

Public Function fffCurSelMode() As dbSelMode
Dim i As Integer
For i = SelOpts.Option.lBound To SelOpts.Option.UBound
    If SelOpts.Option(i).Value Then fffCurSelMode = i: Exit For
Next i
If i = SelOpts.Option.UBound + 1 Then fffCurSelMode = -1
End Function

Public Function CurSelMode() As dbSelMode
CurSelMode = CurSel.SelMode
End Function

Public Sub dbMakeSelData(ByVal x As Long, ByVal y As Long, ByRef sData() As Long)
Dim tBln As Boolean
If AryDims(AryPtr(sData)) <> 2 Then
    Err.Raise 1111, "dbMakeSelData", "A bidimensional array is required!"
End If
If Not ActiveTool = 10 Then ChTool 10
'If CurSel.Selected Then dbDeselect blnApply:=True
With CurSel
    .Selected = True
    .x1 = x
    .y1 = y
    .x2 = .x1 + UBound(sData, 1)
    .y2 = .y1 + UBound(sData, 2)
    'ReDim CurSel_SelData(0 To Width - 1, 0 To Height - 1)
    If AryDims(AryPtr(CurSel_SelData)) <> 2 Then
        CurSel_SelData = sData
    Else
        If VarPtr(CurSel_SelData(0, 0)) <> VarPtr(sData(0, 0)) Then
            CurSel_SelData = sData
        End If
    End If
    x0 = 0
    y0 = 0
    .XM = 0
    .YM = 0
End With
UpdateSelPic True, False
SelPicture.Visible = True
End Sub

'Makes the new selection. If neccessary, changes the tool.
'If Draw is Flase, does not draw and DOES NOT set the selpicture
' visible.
Public Sub dbMakeSel(ByVal x As Long, ByVal y As Long, _
                     ByVal Width As Long, ByVal Height As Long, _
                     Optional ByVal Draw As Boolean = True)
Dim tBln As Boolean
If Not ActiveTool = 10 Then ChTool 10
If CurSel.Selected Then dbDeselect True
With CurSel
    .Selected = True
    .x1 = x
    .y1 = y
    .x2 = .x1 + Width - 1
    .y2 = .y1 + Height - 1
    ReDim CurSel_SelData(0 To Width - 1, 0 To Height - 1)
    CurSel.Moving = tBln
    x0 = 0
    y0 = 0
    .XM = 0
    .YM = 0
End With
If Draw Then
    UpdateSelPic True, False
    SelPicture.Visible = True
End If
End Sub

Private Function dbRemoveAmpersand(strText As String) As String
Dim tmp As String, i As Long
tmp = Replace(strText, "&", vbNullString)
i = InStr(tmp, vbTab)
If i = 0 Then i = Len(tmp) + 1
dbRemoveAmpersand = Left$(tmp, i - 1)
End Function

Public Sub MonochromizeSimple(ByRef mData() As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long)
Dim x As Long, y As Long
'Dim rgb1 As RGBQUAD, t As Long
Dim w As Long, h As Long
AryWH AryPtr(mData), w, h
Dim RGBData() As RGBQUAD

ReferAry AryPtr(RGBData), AryPtr(mData)
On Error GoTo eh
For y = Max(y1, 0) To Min(y2, h - 1)
    For x = Max(x1, 0) To Min(x2, w - 1)
        'GetRgbQuadEx mData(x, y), rgb1
        't = CLng(rgb1.rgbRed) + CLng(rgb1.rgbGreen) + CLng(rgb1.rgbBlue)
        'mData(j, i) = vbWhite And CLng(t > 381)
        mData(x, y) = vbWhite And CLng(CLng(RGBData(x, y).rgbBlue) + RGBData(x, y).rgbGreen + RGBData(x, y).rgbRed > 381)
    Next x
Next y
UnReferAry AryPtr(RGBData)
Exit Sub
eh:
PushError
UnReferAry AryPtr(RGBData)
PopError
ErrRaise
End Sub

Friend Sub vtLongToRgbQuad(ByRef lData() As Long, ByRef RGBData() As RGBQUAD)
Dim w As Long, h As Long
w = UBound(lData, 1) + 1
h = UBound(lData, 2) + 1
ReDim RGBData(0 To w - 1, 0 To h - 1)
CopyMemory RGBData(0, 0), lData(0, 0), w * h * 4&
End Sub






'Show progress is not supported
Public Sub dbRepair(Optional ByVal bShowProgress As Boolean)
vtRepair Data
End Sub

Public Sub dbNegative(ByRef mData() As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long, Colors As Long)
Dim i As Long, j As Long
Dim w As Long, h As Long
If AryDims(AryPtr(mData)) <> 2 Then
    Err.Raise 1111, "dbNegative", "A bidimensional array is required!"
End If
w = UBound(mData, 1) + 1
h = UBound(mData, 2) + 1
If x1 = -1 Then
    x1 = 0
    y1 = 0
    x2 = w - 1
    y2 = h - 1
End If
If x1 < 0 Then x1 = 0
If y1 < 0 Then y1 = 0
If x2 > w - 1 Then x2 = w - 1
If y2 > h - 1 Then y2 = h - 1
For i = y1 To y2
    For j = x1 To x2
        mData(j, i) = (mData(j, i) Xor Colors)
    Next j
Next i
End Sub


Private Function IsDisplay(x As Long, y As Long) As Boolean
IsDisplay = (x >= (HScroll.Value - 8) \ Zm And y >= (VScroll.Value - 8) \ Zm And x <= (HScroll.Value + VScroll.Left) \ Zm + 1 And y <= (VScroll.Value + HScroll.Top) \ Zm + 1)
End Function

Public Sub dbTurn(ByRef mData() As Long, ByVal Method As dbTurnMethod)
Dim i As Long, j As Long, rData() As Long, m As Long, n As Long
If AryDims(AryPtr(mData)) <> 2 Then
    Err.Raise 1111, "dbTurn", "A bidimensional array is required!"
End If
m = UBound(mData, 2)
n = UBound(mData, 1)
Select Case Method
    Case dbTurn90
        ReDim rData(m, n)
        For i = 0 To m
            For j = 0 To n
                rData(m - i, j) = mData(j, i)
            Next j
        Next i
    Case dbTurn180
        ReDim rData(n, m)
        For i = 0 To m
            For j = 0 To n
                rData(n - j, m - i) = mData(j, i)
            Next j
        Next i
    Case dbTurn270
        ReDim rData(m, n)
        For i = 0 To m
            For j = 0 To n
                rData(i, n - j) = mData(j, i)
            Next j
        Next i
    Case dbFlipHor
        ReDim rData(n, m)
        For i = 0 To m
            For j = 0 To n
                rData(n - j, i) = mData(j, i)
            Next j
        Next i
    Case dbFlipVer
        ReDim rData(n, m)
        For i = 0 To m
            For j = 0 To n
                rData(j, m - i) = mData(j, i)
            Next j
        Next i
End Select
'Erase mData
'mData = rData
SwapArys AryPtr(mData), AryPtr(rData)
End Sub


Public Sub dbDeColour(ByRef mData() As Long, _
                      ByVal x1 As Long, _
                      ByVal y1 As Long, _
                      ByVal x2 As Long, _
                      ByVal y2 As Long, _
                      ByVal Amount As Double, _
                      Optional ByVal NoProgress As Boolean = False)
Dim y As Long, x As Long
'Dim h As Byte
'Dim cc(1 To 3) As Byte
Dim rData() As RGBQUAD
Dim w As Long, h As Long
Dim Mid As Double
Dim r As Long, g As Long, b As Long
Dim AAm As Single
Dim OfcY As Long
Dim Inv3 As Double
Inv3 = 1# / 3#

If Amount > 1# Or Amount < 0# Then
    Err.Raise 1111, "dbDeColour", "Decolour amount must be from 0 to 1. But it is " + CStr(Amount) + "!"
End If

On Error GoTo eh
w = UBound(mData, 1) + 1
h = UBound(mData, 2) + 1

If x1 < 0 Then x1 = 0
If y1 < 0 Then y1 = 0
If x2 > w - 1 Then x2 = w - 1
If y2 > h - 1 Then y2 = h - 1

ConstructAry AryPtr(rData), VarPtr(mData(0, 0)), 4, w * h

If Not NoProgress Then
  DisableMe
  BreakKeyPressed
End If

For y = y1 To y2
    OfcY = w * y
    For x = x1 To x2
        Mid = (CDbl(rData(OfcY + x).rgbRed) + rData(OfcY + x).rgbGreen + rData(OfcY + x).rgbBlue) * Inv3
        rData(x + OfcY).rgbRed = (Mid - rData(OfcY + x).rgbRed) * Amount + rData(OfcY + x).rgbRed
        rData(x + OfcY).rgbGreen = (Mid - rData(OfcY + x).rgbGreen) * Amount + rData(OfcY + x).rgbGreen
        rData(x + OfcY).rgbBlue = (Mid - rData(OfcY + x).rgbBlue) * Amount + rData(OfcY + x).rgbBlue
    Next x
    If Not NoProgress Then ShowProgress y / (y2 - y1 + 1) * 100, True
Next y
UnReferAry AryPtr(rData)
If Not NoProgress Then
  ShowProgress 101
  RestoreMeEnabled
End If
Exit Sub
eh:
ClearMeEnabledStack
UnReferAry AryPtr(rData)
ErrRaise "dbDeColour"
End Sub

Public Sub dbFreshSelBorder()
Dim i As Long, j As Long, sH As Long, sW As Long, tmp As Long
Dim x As Long, y As Long
Dim rct As RECT 'this rect is smaller by 1px
Static Flasher As Boolean
sH = SelPicture.ScaleHeight - 1
sW = SelPicture.ScaleWidth - 1
If SelPicture.AutoRedraw Then
    rct.Right = sW
    rct.Bottom = sH
Else
    GetSelVisibilityRect rct
End If
'On Error Resume Next
Flasher = Not Flasher
tmp = CLng(Flasher) And &HFFFFFF

'left vertical
x = 0
If rct.Left <= x And rct.Right >= x Then
    For y = rct.Top To rct.Bottom Step 2&
        tmp = Not (tmp) And &HFFFFFF
        SetPixel SelPicture.hDC, x, y, tmp
        If x And &H100& And MainModule.VistaSetPixelBugDetected Then
          If Not SelPicture.AutoRedraw Then SelPicture.Line (x, y)-(x, y), tmp, BF
        End If
        
    Next y
End If

'right vertical
x = sW
If rct.Left <= x And rct.Right >= x Then
    For y = rct.Top To rct.Bottom Step 2&
        tmp = Not (tmp) And &HFFFFFF
        SetPixel SelPicture.hDC, x, y, tmp
        If x And &H100& And MainModule.VistaSetPixelBugDetected Then
          If Not SelPicture.AutoRedraw Then SelPicture.Line (x, y)-(x, y), tmp, BF
        End If
    Next y
End If

'top horizontal
y = 0
If rct.Top <= y And rct.Bottom >= y Then
    For x = rct.Left To rct.Right Step 2&
        tmp = Not (tmp) And &HFFFFFF
        SetPixel SelPicture.hDC, x, y, tmp
        If x And &H100& And MainModule.VistaSetPixelBugDetected Then
          If Not SelPicture.AutoRedraw Then SelPicture.Line (x, y)-(x, y), tmp, BF
        End If
    Next x
End If

'bottom horizontal
y = sH
If rct.Top <= y And rct.Bottom >= y Then
    For x = rct.Left To rct.Right Step 2&
        tmp = Not (tmp) And &HFFFFFF
        SetPixel SelPicture.hDC, x, y, tmp
        If x And &H100& And MainModule.VistaSetPixelBugDetected Then
          If Not SelPicture.AutoRedraw Then SelPicture.Line (x, y)-(x, y), tmp, BF
        End If
    Next x
End If


End Sub

Public Sub dbAero(ByVal x As Long, ByVal y As Long, _
                  ByVal cSize As Long, _
                  ByVal Intens As Long, _
                  ByVal lngColor As Long)
Dim dy As Long, px As Long, h As Long, rc As Boolean
Dim cSize2 As Long
cSize2 = cSize * cSize
If lngColor = -1 Then rc = True
For h = 1 To Intens
    Do
        dy = Int(Rnd(1) * cSize * 2 - cSize + y)
        px = Int(Rnd(1) * cSize * 2 - cSize + x)
    Loop Until (dy - y) * (dy - y) + (px - x) * (px - x) <= cSize2
    If rc Then lngColor = Rnd * vbWhite
    If TexMode Then
        px = px Mod intW
        If px < 0 Then px = px + intW
        dy = dy Mod intH
        If dy < 0 Then dy = dy + intH
        dbPSet px, dy, lngColor, dbAsmnuGrid, True, Not MP.AutoRedraw
    Else
        If InPicture(px, dy) Then dbPSet px, dy, lngColor, dbAsmnuGrid, True, Not MP.AutoRedraw
    End If
Next h
End Sub

Private Function CountR1(ByVal r2 As Single) As Single
If HSet.RMode = 0 Then
    CountR1 = HSet.RFixed
Else
    CountR1 = r2 / HSet.RK
End If
End Function


Public Sub dbSwap(ByRef a As Variant, ByRef b As Variant)
Dim c As Variant
c = a
a = b
b = c
End Sub

Public Function GetWidth() As Long
GetWidth = intW
End Function

Public Function GetHeight()
GetHeight = intH
End Function

Public Sub dbLoadCaptions()
Me.Caption = GRSF(2143)

'************MENU**************
mnuFile.Caption = GRSF(101)
    mnuNew.Caption = GRSM(102, cmdNew, Keyb)
    
    mnuLoad.Caption = GRSM(103, cmdOpen, Keyb)
    mnuClear.Caption = GRSM(104, cmdClearPic, Keyb)
    
    mnuSaveFile.Caption = GRSM(177, cmdSave, Keyb)
    mnuFolder.Caption = GRSM(310, cmdUnknown, Keyb)
    mnuSaveSel.Caption = GRSM(182, cmdSaveSel, Keyb)
    
    mnuSaveAs.Caption = GRSM(105, cmdSaveAs, Keyb)
'    mnuSaveBmp.Caption = GRSM(105, cmdSaveBMP, Keyb)
'    mnuSavePNG.Caption = GRSM(191, cmdSavePNG, Keyb)
    
    mnuBuildBackUp.Caption = GRSM(338, cmdExtremeSave, Keyb)
    
    mnuPrint.Caption = GRSF(109)
    
    mnuExit.Caption = GRSM(108, cmdCloseInstance, Keyb)


mnuEdit.Caption = GRSF(110)
    mnuUndo.Caption = GRSM(111, cmdUndo, Keyb)
    mnuRedo.Caption = GRSM(112, cmdRedo, Keyb)
    
    mnuClearUndo.Caption = GRSM(118, cmdClearUndo, Keyb)
    
    mnuCopy.Caption = GRSM(113, cmdCopy, Keyb)
    mnuPaste.Caption = GRSM(114, cmdPaste, Keyb)
    
    mnuSelectAll.Caption = GRSM(115, cmdSelectAll, Keyb)
    mnuClear2.Caption = mnuClear.Caption
    
    mnuResize.Caption = GRSM(116, cmdResize, Keyb)
    
    mnuMix.Caption = GRSM(117, cmdLoadSel, Keyb)
    mnuEditSel.Caption = GRSM(304, cmdUnknown, Keyb)
    mnuCrop.Caption = GRSM(318, cmdCropSel, Keyb)
    
    mnuCapture.Caption = GRSM(316, cmdCapture, Keyb)
    
    mnuFormula.Caption = GRSM(317, cmdFormula, Keyb)

mnuView.Caption = GRSF(120)
    mnuPal.Caption = GRSM(121, cmdPalVisible, Keyb)
    'MsgBox mnuPal.Checked
    mnuGrid.Caption = GRSM(122, cmdToggleGrid, Keyb)
    mnuToolBarVis.Caption = GRSM(311, cmdToolBarVisible, Keyb)
    mnuDynamicScr.Caption = GRSM(332, cmdDynamicDialog, Keyb)
    
    mnuZoom.Caption = GRSM(123, cmdZoom, Keyb)
    
    mnuRefresh.Caption = GRSM(124, cmdRefresh, Keyb)
    
    
    
mnuTools.Caption = GRSF(129)
    mnuTool(GetToolIndexByID(0)).Caption = GRSM(130, cmdToolPen, Keyb)
    mnuTool(GetToolIndexByID(1)).Caption = GRSM(131, cmdToolLin, Keyb)
    mnuTool(GetToolIndexByID(2)).Caption = GRSM(132, cmdToolFad, Keyb)
    mnuTool(GetToolIndexByID(3)).Caption = GRSM(133, cmdToolSta, Keyb)
    mnuTool(GetToolIndexByID(4)).Caption = GRSM(134, cmdToolFSt, Keyb)
    mnuTool(GetToolIndexByID(8)).Caption = GRSM(135, cmdToolCir, Keyb)
    mnuTool(GetToolIndexByID(6)).Caption = GRSM(136, cmdToolVFd, Keyb)
    mnuTool(GetToolIndexByID(9)).Caption = GRSM(137, cmdToolHFd, Keyb)
    mnuTool(GetToolIndexByID(7)).Caption = GRSM(138, cmdToolPnt, Keyb)
    mnuTool(GetToolIndexByID(5)).Caption = GRSM(139, cmdToolGet, Keyb)
    mnuTool(GetToolIndexByID(10)).Caption = GRSM(140, cmdToolSel, Keyb)
    mnuTool(GetToolIndexByID(11)).Caption = GRSM(141, cmdToolRec, Keyb)
    mnuTool(GetToolIndexByID(12)).Caption = GRSM(142, cmdToolPol, Keyb)
    mnuTool(GetToolIndexByID(13)).Caption = GRSM(143, cmdToolAir, Keyb)
    mnuTool(GetToolIndexByID(14)).Caption = GRSM(144, cmdToolHel, Keyb)
    mnuTool(GetToolIndexByID(16)).Caption = GRSM(146, cmdToolBrh, Keyb)
    mnuTool(GetToolIndexByID(17)).Caption = GRSM(147, cmdToolPal, Keyb)
    mnuTool(GetToolIndexByID(18)).Caption = GRSM(148, cmdToolTxt, Keyb)
    mnuTool(GetToolIndexByID(ToolOrg)).Caption = GRSM(321, cmdToolOrg, Keyb)
    mnuTool(GetToolIndexByID(ToolProg)).Caption = GRSM(320, cmdToolPrg, Keyb)
    
    
    mnuToolOpts.Caption = GRSM(149, cmdToolProps, Keyb)
    mnuSelBrush.Caption = GRSM(312, cmdUnknown, Keyb)

mnuSelAll.Caption = GRSF(346)
    
    mnuSelShow.Caption = GRSM(356, cmdSelMoveTo, Keyb)
    mnuSelMoveTo.Caption = GRSM(347, cmdSelMoveTo, Keyb)
    
    mnuSelCenterHorz.Caption = GRSM(348, cmdSelHCenter, Keyb)
    mnuSelCenterVert.Caption = GRSM(349, cmdSelVCenter, Keyb)
    
    mnuSelClear.Caption = GRSM(350, cmdSelClear, Keyb)
    mnuSelResize.Caption = GRSM(351, cmdSelResize, Keyb)
    mnuSelEditSel.Caption = GRSM(352, cmdSelEdit, Keyb)
    
    mnuSelStamp.Caption = GRSM(353, cmdDeselect_no_delete, Keyb)
    mnuselDelete.Caption = GRSM(354, cmdDeleteSel, Keyb)
    mnuSelAutoRepaint.Caption = GRSM(355, cmdSelToggleRedraw, Keyb)
    
    mnuSelShowSize.Caption = GRSM(358, cmdSelShowSize, Keyb)

mnuDraw.Caption = GRSF(185)
    mnuDrawPlain.Caption = GRSM(186, cmdDrawPlain, Keyb)
    mnuDrawBG.Caption = GRSM(187, cmdDrawBg, Keyb)
    mnuDrawBBg.Caption = GRSM(313, cmdDrawBBg, Keyb)
    mnuDrawWaves.Caption = GRSM(357, cmdDrawWaves, Keyb)
    
mnuTexture.Caption = GRSF(334)
    mnuTexMode.Caption = GRSM(335, cmdTexMode, Keyb)
    
    mnuResetOrg.Caption = GRSM(336, cmdResetOrg, Keyb)
    mnuRestoreOrg.Caption = GRSM(337, cmdRestoreOrg, Keyb)
    

mnuPrgs.Caption = GRSF(302)
    mnuPrgDraw.Caption = GRSM(307, cmdPrgDrawings, Keyb)

mnuEffects.Caption = GRSF(150)
    mnuEffect(0).Caption = GRSM(151, cmdEffect0, Keyb)
    mnuEffect(2).Caption = GRSM(152, cmdEffect2, Keyb)
    mnuEffect(3).Caption = GRSM(153, cmdEffect3, Keyb)
    mnuEffect(4).Caption = GRSM(154, cmdEffect4, Keyb)
    mnuEffect(5).Caption = GRSM(155, cmdEffect5, Keyb)
    mnuEffect(6).Caption = GRSM(156, cmdEffect6, Keyb)
    mnuEffect(7).Caption = GRSM(157, cmdEffect7, Keyb)
    mnuEffect(8).Caption = GRSM(158, cmdEffect8, Keyb)
    mnuEffect(9).Caption = GRSM(194, cmdEffect9, Keyb)
    mnuEffect(10).Caption = GRSM(322, cmdEffect10, Keyb)
    mnuEffect(11).Caption = GRSM(339, cmdEffect11, Keyb)
    
    mnuLastEffect.Caption = GRSM(309, cmdRepLastEffect, Keyb)


mnuPalette.Caption = GRSF(178)
    mnuLoadPal.Caption = GRSM(179, cmdLoadPAL, Keyb)
    mnuSavePal.Caption = GRSM(180, cmdSavePAL, Keyb)
    
    mnuResetPal.Caption = GRSM(181, cmdPalReset, Keyb)
    mnuPalDef.Caption = GRSM(166, cmdMakePalDef, Keyb)
    
    mnuPalCount.Caption = GRSM(165, cmdNPalEntries, Keyb)
    mnuStretchPal.Caption = GRSM(306, cmdStretchPal, Keyb)
    
    mnuFillTips.Caption = GRSM(195, cmdPalFillTips, Keyb)
    mnuEmptyTips.Caption = GRSM(196, cmdPalClearTips, Keyb)
    
    mnuAutoPal.Caption = GRSF(199)
        mnuSysPal.Caption = GRSM(197, cmdDefPalSysColors, Keyb)
        mnuQBColors.Caption = GRSM(198, cmdDefPalBRH, Keyb)
        
        mnuDefPal16.Caption = GRSM(300, cmdDefPal16, Keyb)
        mnuDefPal256.Caption = GRSM(301, cmdDefPal256, Keyb)
    
mnuOptions.Caption = GRSF(160)
    mnuWheelUse.Caption = GRSF(161)
        mnuUseWheel(0).Caption = GRSM(162, cmdToggleMiddleButtonUse, Keyb)
        mnuUseWheel(1).Caption = GRSM(163, cmdToggleMiddleButtonUse, Keyb)
    
    mnuKeyb.Caption = GRSM(164, cmdUnknown, Keyb)
    mnuKeyb2.Caption = GRSM(314, cmdKeyboard, Keyb)
    mnuPSens.Caption = GRSM(315, cmdUnknown, Keyb)
    
    mnuMouseAttr.Caption = GRSM(341, cmdGlueMouse, Keyb)
    mnuSetAutoScrolling.Caption = GRSM(340, cmdDynamicDialog, Keyb)
    
    mnuUndoLim.Caption = GRSM(168, cmdUnknown, Keyb)
    mnuNoUndoRedo.Caption = GRSM(176, cmdDisableUndo, Keyb)
    
    mnuReg.Caption = GRSM(170, cmdExts, Keyb)
    
    
    mnuShowSplash.Caption = GRSM(359, cmdUnknown, Keyb)
    mnuOptRememberWndPos.Caption = GRSM(167, cmdUnknown, Keyb)
    
    mnuResetAll.Caption = GRSM(192, cmdFullReset, Keyb)
    mnuUnInstall.Caption = GRSM(360, cmdUnknown, Keyb)


mnuHelp.Caption = GRSF(171)
    mnuHowToUse.Caption = GRSF(172)
    mnuIdleMessage.Caption = GRSM(333, cmdIdleMessage, Keyb)
    mnuAbout.Caption = GRSF(173)
    mnuWeb.Caption = GRSF(342)
    mnuWebMail.Caption = GRSF(344)
    mnuWebUpdates.Caption = GRSF(343)
    mnuWebForum.Caption = GRSF(345)


mnuLnsZoomIn.Caption = GRSF(323)
mnuLnsZoomOut.Caption = GRSF(324)

mnuLnsToggle.Caption = GRSF(325)

mnuLnsDock.Caption = GRSF(326)

lnsHint1.Caption = GRSF(328)
lnsHint2.Caption = GRSF(329)
lnsHint3.Caption = GRSF(330)
lnsHint4.Caption = GRSF(331)


#If NOREG Then
    mnuRegister.Visible = False
#Else
    mnuRegister.Caption = GRSF(190)
#End If
'*************/MENU/**************

ActiveColor(1).ToolTipText = GRSF(231)
frmColors.ToolTipText = GRSF(241)
Picture1.ToolTipText = GRSF(201)
Status.Caption = GRSF(210) 'Ready
ActiveColor(2).ToolTipText = GRSF(221)
End Sub

Public Function dbCurSelTrR() As Single
    dbCurSelTrR = CurSel.TransRatio
End Function

Public Sub FreshCaption()
If Len(OpenedFileName) > 0 Then
Me.Caption = GRSF(2143) + " - [" + OpenedFileName + "]"
Else
Me.Caption = GRSF(2143)
End If
End Sub

Friend Sub DrawData(ByRef dData() As Long, _
                    ByVal ZoomRatio As Long, _
                    ByVal DestHdc As Long, _
                    ByRef DrawRect As RECT, _
                    Optional ByVal mx As Long = &H7FFFFFFF, Optional ByVal my As Long = &H7FFFFFFF)
Dim RGBData() As RGBQUAD
Dim i As Long
Dim DestSize32 As Long
Dim DestSize8
Dim LineSize8 As Long
Dim LineOfc As Long
Dim StartLine As Long
Dim w As Long, h As Long
Dim DrawW As Long, DrawH As Long
Dim tmpR As Byte
Dim j As Long
Dim Ofc32 As Long
Dim x1 As Long, y1 As Long, x2 As Long, y2 As Long
Dim x As Long, y As Long
w = UBound(dData, 1) + 1
h = UBound(dData, 2) + 1

x1 = DrawRect.Left
y1 = DrawRect.Top
x2 = DrawRect.Right - 1
y2 = DrawRect.Bottom - 1

DrawW = DrawRect.Right - DrawRect.Left
DrawH = DrawRect.Bottom - DrawRect.Top
If DrawW <= 0 Or DrawH <= 0 Then Exit Sub
DestSize32 = DrawW * DrawH - 1
DestSize8 = DrawW * DrawH * 4

LineSize8 = DrawW * 4

StartLine = DrawRect.Top
LineOfc = DrawRect.Left


Dim bmpI As BITMAPINFO
With bmpI.bmiHeader
    .biWidth = DrawW
    .biHeight = -DrawH
    
    .biBitCount = 32
    .biPlanes = 1
    
    .biSize = Len(bmpI.bmiHeader)
    .biSizeImage = DestSize8
End With

Dim tmpHDC As Long
Dim hDefBitmap As Long
Dim hNewBitmap As Long
Dim MappedData() As RGBQUAD
Dim Bits() As RGBQUAD
Dim ptrBits As Long
hDefBitmap = -1
'hNewBitmap = CreateCompatibleBitmap(DestHdc, DrawW, DrawH)
tmpHDC = CreateCompatibleDC(DestHdc)
hNewBitmap = CreateDIBSection(tmpHDC, bmpI, DIB_RGB_COLORS, VarPtr(ptrBits), 0, 0)

ConstructAry AryPtr(Bits), ptrBits, 4, DrawW, DrawH
On Error GoTo eh

'sending bits
GdiFlush
For y = 0 To DrawH - 1
    y2 = y1 + y
    CopyMemory Bits(0, y), dData(x1, y2), DrawW * 4&
Next y
UnReferAry AryPtr(Bits), False

hDefBitmap = SelectObject(tmpHDC, hNewBitmap)
SetStretchBltMode DestHdc, COLORONCOLOR

If mx = &H7FFFFFFF Then
    mx = LineOfc
End If
If my = &H7FFFFFFF Then
    my = StartLine
End If

Call StretchBlt(DestHdc, mx * ZoomRatio, my * ZoomRatio, DrawW * ZoomRatio, DrawH * ZoomRatio, tmpHDC, 0, 0, DrawW, DrawH, SRCCOPY)
rsm:
On Error GoTo 0
If hDefBitmap <> -1 Then
    SelectObject tmpHDC, hDefBitmap
End If
DeleteObject hNewBitmap
DeleteDC tmpHDC


Exit Sub
eh:
UnReferAry AryPtr(Bits), False
UnReferAry AryPtr(MappedData), False
Resume rsm
End Sub

Function GetImageSize(ByVal w As Long, ByVal h As Long, ByVal BytesPerPixel As Long) As Long
Dim Aw As Long
Aw = -Int(-w * BytesPerPixel / 4) * 4
GetImageSize = Aw * h
End Function

'returns True if exited due to timeout
Function dbBrushLine(ByRef dbBrush() As Byte, _
                ByVal x1 As Long, ByVal y1 As Long, _
                ByVal x2 As Long, ByVal y2 As Long, _
                ByVal lngColor As Long, _
                Optional ByVal DrawTemp As Boolean = False, _
                Optional ByVal EnableTimeout As Boolean = False, _
                Optional ByVal DontAddPoint As Boolean = False) As Boolean
Const Max_Processing_Time As Long = 10 'ms
Const Refresh_Period As Long = 50 'ms
Static i As Long
Static stX1 As Long, stY1 As Long, stX2 As Long, stY2 As Long 'pos of current processing' point
Static Ded As Long, invDed As Double
Static xx As Long, yy As Long
Static ToDraw() As BrushPoint 'array of points left to draw
Static iLast As Long 'array index where to put the new point
Static iFirst As Long 'index of the item to be/being drawn right now
Static l As Long
Static Stopped As Boolean ' is set to true upon exiting sub due to timeout
Static LastRefreshTime As Long
Dim EnterTime As Long
Dim t As Long 'timegettime
If Not DontAddPoint Then GoSub addpoint
On Error GoTo eh
t = timeGetTime
EnterTime = t
If Not Stopped Then LastRefreshTime = t
Do While iFirst < iLast
  If Stopped Then
    Stopped = False
    GoTo ResumeLastLine
  End If
  Stopped = False
  With ToDraw(iFirst)
    stX1 = .x1
    stY1 = .y1
    stX2 = .x2
    stY2 = .y2
    lngColor = .Color
  End With
  iFirst = iFirst + 1
  Ded = Max(Abs(stX1 - stX2), Abs(stY1 - stY2))
  If Ded = 0 Then
      dbBSet dbBrush, stX1, stY1, lngColor, DrawTemp, Draw:=True
  Else
      invDed = 1 / Ded
      i = 0
      Do Until i >= Ded
          xx = Round((stX2 - stX1) * i * invDed) + stX1
          yy = Round((stY2 - stY1) * i * invDed) + stY1
          dbBSet dbBrush, xx, yy, lngColor, DrawTemp, Draw:=True
          t = timeGetTime
          If t - LastRefreshTime > Refresh_Period Then
            LastRefreshTime = t
            GoSub Refresh
          End If
          If EnableTimeout And t - EnterTime > Max_Processing_Time Then
            dbBrushLine = True
            Stopped = True
            Exit Function
          End If
ResumeLastLine:
          i = i + 1
      Loop
  End If
Loop
iFirst = 0
iLast = 0
GoSub Refresh
Exit Function
eh:
Stopped = False
Debug.Assert False
Exit Function
addpoint:
l = AryLen(AryPtr(ToDraw))
If l = 0 Then
  l = 256
  ReDim ToDraw(0 To l - 1)
ElseIf l - 1 < iLast Then
  l = l + 256
  ReDim Preserve ToDraw(0 To l - 1)
End If
ToDraw(iLast).x1 = x1
ToDraw(iLast).y1 = y1
ToDraw(iLast).x2 = x2
ToDraw(iLast).y2 = y2
ToDraw(iLast).Color = lngColor
iLast = iLast + 1
Return

Refresh:
If MP.AutoRedraw Then MP.Refresh
Return
End Function

Public Sub dbBSet(ByRef Brh() As Byte, _
                  ByVal x As Long, ByVal y As Long, _
                  ByVal lngColor As Long, _
                  Optional ByVal DrawTemp As Boolean = False, _
                  Optional ByVal Draw As Boolean = True)
Dim i As Long, j As Long
Dim tx As Long, ty As Long
Dim cx As Long, cy As Long
Dim UBX As Long, UBY As Long
Dim rgb1 As RGBQuadLong, rgb2 As RGBQuadLong
Dim r1 As Long, g1 As Long, b1 As Long
Dim r2 As Long, g2 As Long, b2 As Long
Dim Blend As Long
Dim c As Currency

UBX = UBound(Brh, 1)
UBY = UBound(Brh, 2)
cx = UBX \ 2
cy = UBY \ 2
If lngColor = -1 Then
    For i = 0 To UBY
        For j = 0 To UBX
            If Brh(j, i) Then
                tx = x + j - cx
                ty = y + i - cy
                If InPicture(tx, ty) Then
                    dbPSet tx, ty, Data(tx, ty), dbAsmnuGrid, Not DrawTemp, Not MP.AutoRedraw, Draw
                End If
            End If
        Next j
    Next i
Else
    r1 = lngColor And &HFF&
    g1 = lngColor And &HFF00&
    b1 = lngColor And &HFF0000
    For i = 0 To UBY
        For j = 0 To UBX
            tx = x + j - cx
            ty = y + i - cy
            If InPicture(tx, ty) Then
                Blend = 255 - Brh(j, i)
                If Blend < 255 Then
                    r2 = Data(tx, ty)
                    g2 = (((r2 And &HFF&) - r1) * Blend \ 255 + r1) Or _
                         ((((r2 And &HFF00&) - g1) * Blend \ 255 + g1) And &HFF00&) Or _
                         (((((r2 And &HFF0000) - b1) \ 255) * Blend + b1) And &HFF0000)
                    If Draw Then
                        dbPSet tx, ty, g2, dbAsmnuGrid, Not DrawTemp, Not MP.AutoRedraw, Draw
                    Else
                        StorePixel tx, ty
                        Data(tx, ty) = g2
                    End If
                End If
            End If
        Next j
    Next i
End If
End Sub

Sub LoadLastBrush(ByRef dbBrush() As Byte)
Dim i As Long, j As Long, w As Long, h As Long, n As Long, tmp As String
Dim ptr As Long
On Error GoTo eh
w = dbGetSettingEx("Tool\LastBrush", "Width", vbLong, 3)
h = dbGetSettingEx("Tool\LastBrush", "Height", vbLong, 3)
If h <= 0 Or w <= 0 Then
    w = 3
    h = 3
    tmp = "00FF00FFFFFF00FF00"
Else
    tmp = dbGetSetting("Tool\LastBrush", "Data", "00FF00FFFFFF00FF00")
End If
ReDim dbBrush(0 To w - 1, 0 To h - 1)
n = 0
ptr = 1
For i = 0 To h - 1
    For j = 0 To w - 1
        dbBrush(j, i) = CByte("&H" + Mid$(tmp, ptr, 2))
        ptr = ptr + 2
    Next j
Next i
Exit Sub
eh:
ReDim dbBrush(0 To 2, 0 To 2)
Err.Raise 1001, "Registry", "Error loading last used brush"
Exit Sub
End Sub

Sub SaveCurBrush(ByRef dbBrush() As Byte)
Dim i As Long, j As Long, w_1 As Long, h_1 As Long, tmp As String, n As Long
On Error GoTo eh
If AryDims(AryPtr(dbBrush)) <> 2 Then Exit Sub
w_1 = UBound(dbBrush, 1)
h_1 = UBound(dbBrush, 2)
On Error GoTo 0
tmp = Space$((w_1 + 1) * (h_1 + 1) * 2)
n = 1
For i = 0 To h_1
    For j = 0 To w_1
        Mid(tmp, n, 2) = Hex$(dbBrush(j, i))
        n = n + 2&
    Next j
Next i
dbSaveSettingEx "Tool\LastBrush", "Width", (w_1 + 1)
dbSaveSettingEx "Tool\LastBrush", "Height", (h_1 + 1)
dbSaveSettingEx "Tool\LastBrush", "Data", tmp
eh:
End Sub

Friend Function GetLngColor(ByRef rgbColor As RGBQUAD)
GetLngColor = RGB(rgbColor.rgbRed, rgbColor.rgbGreen, rgbColor.rgbBlue)
End Function

Friend Function GetLngColorBGR(ByRef rgbColor As RGBQUAD)
GetLngColorBGR = RGB(rgbColor.rgbBlue, rgbColor.rgbGreen, rgbColor.rgbRed)
End Function

'Function InData(ByRef pData() As Long, ByVal i As Long, ByVal j As Long) As Boolean
'InData = (i >= LBound(pData, 2) And j >= LBound(pData, 1) And i <= UBound(pData, 2) And j <= UBound(pData, 1))
'End Function
'
Function ShowProgress(ByVal PerCents As Single, Optional ByVal DoDoEvents As Boolean = False) '0 to 100
Dim Prots As Single, gtc As Long, i As Long, SL As Single
Static t As Long, s As String, Pic As IPictureDisp
gtc = GetTickCount
If PerCents <= 100! Then
    If gtc - t > 200& Then
        t = GetTickCount
        Prots = PerCents * 0.01!
        SL = TextWidth(s)
        s = CStr(Int(PerCents)) + "%"
        If Pic Is Nothing Then
            Set Pic = LoadResPicture("BTN_FOCUS", vbResBitmap)
        End If
        Picture1_Paint
        If Prots > 0! Then
            On Error Resume Next
            Picture1.PaintPicture Pic, 0&, 0&, Picture1.ScaleWidth * Prots, Picture1.ScaleHeight
            On Error GoTo 0
        End If
        Picture1.CurrentX = Picture1.ScaleWidth - Picture1.TextWidth(s)
        Picture1.CurrentY = (Picture1.ScaleHeight - Picture1.TextHeight(s)) / 2
        Picture1.Print s
        Picture1.Refresh
        If BreakKeyPressed Then
            Err.Raise dbCWS
        End If
        If DoDoEvents And Not DontDoEvents Then DoEvents
    End If
Else
    Picture1_Paint
    Status.Refresh
End If

End Function

Public Sub dbMakeMonoRND(ByRef eData() As Long)
Dim i As Long, j As Long, tmp As RGBQUAD, Intens As Long
Randomize 0
Rnd -1
On Error GoTo eh
DisableMe
For i = 0 To UBound(eData, 2)
    For j = 0 To UBound(eData, 1)
        GetRgbQuadEx eData(j, i), tmp
        Intens = CLng(tmp.rgbBlue) + CLng(tmp.rgbRed) + CLng(tmp.rgbGreen)
        If Intens = 0 Then
                eData(j, i) = 0
        Else
            If Int(Rnd(1) * (765 / CSng(Intens))) = 0 Then
                eData(j, i) = vbWhite
            Else
                eData(j, i) = 0
            End If
        End If
    Next j
    ShowProgress i / (UBound(eData, 2) + 1) * 100
Next i
ShowProgress 101
RestoreMeEnabled
Exit Sub
eh:
ClearMeEnabledStack
Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Sub dbMakeMono(ByRef eData() As Long)
Const PatternCount = 16
Dim i As Long, j As Long, tmp As RGBQUAD, Intens As Long, n As Integer, w As Long, h As Long
Dim Patterns(0 To PatternCount) As IconMask
On Error GoTo eh
DisableMe
ShowStatus "Loading patterns"
For i = 0 To PatternCount
    LoadResBrush 200 + i, "PATTERN", Patterns(i)
Next i
ShowStatus "Processing"
For i = 0 To UBound(eData, 2)
    For j = 0 To UBound(eData, 1)
        GetRgbQuadEx eData(j, i), tmp
        Intens = CLng(tmp.rgbBlue) + CLng(tmp.rgbRed) + CLng(tmp.rgbGreen)
        n = Round(Intens * PatternCount / (255 * 3))
        w = UBound(Patterns(n).d, 1) + 1
        h = UBound(Patterns(n).d, 2) + 1
        If Patterns(n).d(i Mod h, j Mod w) Then
            eData(j, i) = vbWhite
        Else
            eData(j, i) = 0
        End If
    Next j
    ShowProgress i / (UBound(eData, 2) + 1) * 100
Next i
ShowProgress 101
RestoreMeEnabled
Exit Sub
eh:
ClearMeEnabledStack
Err.Raise Err.Number, Err.Source, Err.Description
End Sub


Sub StartTimer()
gTimer = GetTickCount
End Sub

Function GetTimer() As Long
GetTimer = GetTickCount - gTimer
End Function

Sub SetMousePtr(Busy As Boolean)
If Busy Then
    Screen.MousePointer = vbHourglass
Else
    Screen.MousePointer = 0
End If
End Sub

Private Sub ApplyScrollBarsValues(Optional ByVal Immediately As Boolean = False)
Dim x As Long, y As Long
If HScrollEnabled Then
    x = 8 - HScroll.Value
Else
    x = (MPHolder.ScaleWidth - MP.Width) \ 2
End If
If VScrollEnabled Then
    y = 8 - VScroll.Value
Else
    y = (MPHolder.ScaleHeight - MP.Height) \ 2
End If
If ScrollSettings.DS_Enabled Then
    MoveMP x, y, IIf(Immediately, &H4, &H0)
Else
    MoveMP x, y, &H4
    If Not MP.AutoRedraw Then
        MP_Paint
    End If
End If
'SmthHlast = HScroll.Value
'SmthVlast = VScroll.Value
End Sub


Public Property Get MeEnabled() As Boolean
MeEnabled = prvMeEnabled
End Property

Public Property Let MeEnabled(ByVal bNew As Boolean)
Me.Enabled = bNew
prvMeEnabled = bNew
End Property



Private Sub MoveSelRect(ByVal x As Long, ByVal y As Long, ByVal Width As Long, ByVal Height As Long)
Dim tmpX As Long, tmpY As Long, tmpW As Long, tmpH As Long
If Height < 0 And Width < 0 Then
    tmpX = x + Width
    tmpY = y + Height
    tmpW = Abs(Width)
    tmpH = Abs(Height)
ElseIf Height < 0 Then
    tmpX = x
    tmpY = y + Height '+ because height is already <0
    tmpW = Width
    tmpH = Abs(Height)
ElseIf Width < 0 Then
    tmpX = x + Width
    tmpY = y
    tmpW = Abs(Width)
    tmpH = Height
Else
    tmpX = x
    tmpY = y
    tmpW = Width
    tmpH = Height
End If
SelRect.Move tmpX, tmpY, tmpW, tmpH
End Sub


Public Sub OutText(Optional ByRef Text As String)
Dim Txt As String, fnt As New StdFont ', tmpData() As Long
Dim Transp As Boolean
Dim AC As Long
Dim x As Long, y As Long
Dim w As Long, h As Long
Dim Q As eTextQuality
Dim TextPic() As Long
Load frmText
With frmText
    .SetColor ACol(1)
    If Len(Text) = 0 Then
      .Show vbModal
      If .Tag <> "" Then
          ChTool LUT
          Exit Sub
      End If
      Txt = .Text1.Text
    Else
      Txt = Text
    End If
    If Len(Txt) = 0 Then
      ChTool LUT
      Exit Sub
    End If
    ChTool ToolSel
    Set fnt = .Text1.Font
    fnt.Size = .nmbSize.Value
    ChangeActiveColor 1, .GetColor And &HFFFFFF, False
    Transp = CBool(.GetColor And &H1000000)
    Q = .GetQuality
End With
Unload frmText
With TempBox
    .Visible = True
    'If Transp Then
        .ForeColor = vbBlack
        .BackColor = vbWhite
    'Else
    '    .ForeColor = ConvertColorLng(ACol(1))
    '    .BackColor = ConvertColorLng(ACol(2))
    'End If
    Set .Font = fnt
    .FontSize = fnt.Size
    Dim Sz As Size
    GetTextExtentPoint32 .hDC, Txt, Len(Txt), Sz
    w = .TextWidth(Txt)
    h = .TextHeight(Txt)
    Dim ABC1 As ABC, ABC2 As ABC
    GetCharABCWidths .hDC, Asc(Left$(Txt, 1)), Asc(Left$(Txt, 1)), ABC1
    GetCharABCWidths .hDC, Asc(Right$(Txt, 1)), Asc(Right$(Txt, 1)), ABC2
    w = w + Max(0, -ABC1.abcA) + Max(0, -ABC2.abcC) + 1
    If w * h > 4000000 Then
      'Rendering such a large amount of text will take a lot of resources (%M megapixels). Are you sure you want to continue?`Text
      If dbMsgBox(grs(1112, "%M", CStr(Round(w * h / 1000000#, 1))), vbYesNo) = vbNo Then Err.Raise dbCWS
    End If
    .Width = w
    .Height = h
    .AutoRedraw = True
    .Picture = LoadPicture("")
    .Cls
    .CurrentX = Max(0, -ABC1.abcA)
    .CurrentY = 0
    TempBox.Print Txt
    .Refresh
    #If GetDIBitsErrors Then
        On Error Resume Next
    #End If
    Call dbGetDIBits(TempBox.Image.Handle, TempBox.hDC, TextPic)
    Select Case Q
      Case etqGray
        dbDeColour TextPic, 0, 0, w - 1, h - 1, Amount:=1, NoProgress:=True
      Case etqMono
        MonochromizeSimple TextPic, 0, 0, w - 1, h - 1
      Case Else
        'use full quality - do nothing
    End Select
    
    dbMakeSel (HScroll.Value) \ Zm, VScroll.Value \ Zm, .Width, .Height, Draw:=False
    If Transp Then
        'CurSel.SelMode = dbSuperTransparent
        CurSel.SetIsText
        SwapArys AryPtr(TransOrigData), AryPtr(TextPic) 'Call dbGetDIBits(TempBox.Image.Handle, TempBox.hDC, TransOrigData)
        Erase TransData
        AC = ACol(1)
        ClearPic CurSel_SelData, AC
'        For Y = 0 To UBound(CurSel_SelData, 2)
'            For X = 0 To UBound(CurSel_SelData, 1)
'                CurSel_SelData(X, Y) = AC
'            Next X
'        Next Y
    Else
        'text is rendered b/w. Render it in color.
        Dim tbl1() As Byte, tbl2() As Byte, tbl3() As Byte
        GenerateOutTextColorMaps ACol(1), ACol(2), tbl1, tbl2, tbl3
        Dim Range As RECT
        Range.Right = w
        Range.Bottom = h
        dbMapColorsEx TextPic, tbl1, tbl2, tbl3, Range, TextPic
        'and use it as selection picture
        SwapArys AryPtr(CurSel_SelData), AryPtr(TextPic) 'Call dbGetDIBits(TempBox.Image.Handle, TempBox.hDC, CurSel_SelData)
        TransDataChanged = True
        'Erase TransData
    End If
    #If GetDIBitsErrors Then
        On Error GoTo 0
    #End If
    mnuSelShow_Click
    SelPicture.Visible = True
    'dbPutSel
    .BackColor = vbBlack
    .ForeColor = vbWhite
    .Width = 32
    .Height = 32
    .Cls
    .Visible = False
End With
    ShowStatus "$10018"
End Sub

Private Sub GenerateOutTextColorMaps(ByVal FC As Long, _
                                     ByVal BC As Long, _
                                     rTbl() As Byte, _
                                     gTbl() As Byte, _
                                     bTbl() As Byte)
Dim fcrgb As RGBQUAD, bcrgb As RGBQUAD
CopyMemory fcrgb, FC, 4
CopyMemory bcrgb, BC, 4
ReDim rTbl(0 To 255)
ReDim gTbl(0 To 255)
ReDim bTbl(0 To 255)
Dim i As Long
For i = 0 To 255
  rTbl(i) = Int((255& - i) * (CLng(fcrgb.rgbRed) - bcrgb.rgbRed) / 255 + 0.5) + bcrgb.rgbRed
  gTbl(i) = Int((255& - i) * (CLng(fcrgb.rgbGreen) - bcrgb.rgbGreen) / 255 + 0.5) + bcrgb.rgbGreen
  bTbl(i) = Int((255& - i) * (CLng(fcrgb.rgbBlue) - bcrgb.rgbBlue) / 255 + 0.5) + bcrgb.rgbBlue
Next i
End Sub

Public Sub InitToolsSizes()
Dim i As Integer
Const Tool_Height = 32 + 3 'I don't know why 3
Const Tool_Width = 32 + 4
For i = btnTool.lBound To btnTool.UBound
    btnTool(i).Move btnTool(i).Left, btnTool(i).Top, Tool_Width, Tool_Height
Next i
End Sub

Friend Function LoadResBrush(ResID As Integer, ByVal ResType As String, ByRef Brh As IconMask)
Dim Bytes() As Byte, i As Long, j As Long, w As Long, h As Long, P As Long
Bytes = LoadResData(ResID, UCase$(ResType))
w = Bytes(0)
h = Bytes(1)
P = 2
ReDim Brh.d(0 To w - 1, 0 To h - 1)
For i = 0 To h - 1
    For j = 0 To w - 1
        Brh.d(j, i) = CBool(Bytes(P))
        P = P + 1&
    Next j
Next i
End Function

Public Function InButton(ByRef cont As Control, _
                         ByVal x As Long, ByVal y As Long, _
                         Optional ByVal ParentTwips As Boolean = False, _
                         Optional ByVal IsdbButton As Boolean = True) As Boolean
If IsdbButton Then
    If (ParentTwips) Then
        x = x * Screen.TwipsPerPixelX
        y = y * Screen.TwipsPerPixelY
    End If
Else
    If Not (ParentTwips) Then
        x = x \ Screen.TwipsPerPixelX
        y = y \ Screen.TwipsPerPixelY
    End If
End If
InButton = (x >= 0 And y >= 0 And x < cont.Width And y < cont.Height)
End Function

Private Sub DrawPalEntry(Index As Integer)
    frmColors.Line (ChCol(Index).Left, ChCol(Index).Top)- _
               Step(ChCol(Index).Width - 1, ChCol(Index).Height - 1), _
               ConvertColorLng(ChCol(Index).BackColor), BF
End Sub

Private Function ChCol_Count() As Integer
ChCol_Count = UBound(ChCol) + 1
End Function

Private Function GetChColIndex(ByVal x As Long, ByVal y As Long) As Integer
Dim n As Integer
n = (ChCol_Count + 1) \ 2
GetChColIndex = n * (y \ ChCol(0).Height) + (x + 1) * n \ frmColors.ScaleWidth
End Function

Sub ChColBackColor(ByVal Index As Integer, ByVal lngColor As Long)
If Index > UBound(ChCol) Then Exit Sub
ChCol(Index).BackColor = lngColor
DrawPalEntry Index
End Sub

Public Sub SetTransData(ByRef tData() As Long)
TransOrigData = tData
TransDataChanged = True
End Sub

Public Sub MoveByWheel(ByVal Movement As Double, ByVal Shift As dbShiftConstants)
Dim dcr As Long
Const MovementK As Double = 60
Static LeftToMove As Double
  Movement = Movement + LeftToMove
  dcr = -MovementK * Movement
  LeftToMove = Movement - Int(MovementK * Movement) / MovementK
  If Shift = 0 Then
      ChangeScrollBarValue VScroll, dcr
  ElseIf Shift = dbStateShift Then
      ChangeScrollBarValue HScroll, dcr
  End If
  If Not MP.AutoRedraw Then ApplyScrollBarsValues
End Sub

Friend Sub Form_Wheel(ByVal Movement As Double, ByVal Shift As dbShiftConstants)
Dim dcr As Long, i As Long
Dim IntMovement As Long
IntMovement = Sgn(Movement) * Int(Max(Abs(Movement), 1) + 0.5)
If Not DrawingLine Then
        If Shift = dbStateCtrl Then
            ReZoomPtr Zm + IntMovement
        ElseIf Shift = dbStateAlt Then
            For i = 0 To mnuTool.UBound
                If mnuTool(i).Checked Then
                    Exit For
                End If
            Next i
            If i <> mnuTool.UBound + 1 Then
                Debug.Print "Tool changed to " + CStr(i)
                i = i - Sgn(Movement)
                If i > mnuTool.UBound Then i = mnuTool.lBound
                If i < mnuTool.lBound Then i = mnuTool.UBound
                mnuTool_Click CInt(i)
            End If
        End If
    If Movement <> 0 Then
        If Not ScrollSettings.CancelWheelScroll Then
            MoveByWheel Movement, Shift
        Else
            If IntMovement <> 0 Then
              MWM = MWM + IntMovement
              MoveMouse
            End If
        End If
    End If

Else
    If ActiveTool <> ToolProg Then
        If IntMovement > 0 Then
            newLineK = LineK * 1.5 ^ IntMovement
            If newLineK > 1.5 ^ 8 Then newLineK = 1.5 ^ 8
        ElseIf IntMovement < 0 Then
            newLineK = LineK * 1.5 ^ IntMovement
            If newLineK < 1.5 ^ (-8) Then newLineK = 1.5 ^ (-8)
        End If
        If Not ScrollSettings.CancelWheelScroll Then
            MoveByWheel Movement, Shift
        End If
        MoveMouse
    Else
        'ToolPrg_MouseEvent xo,yo,
        If Not ScrollSettings.CancelWheelScroll Then
            MoveByWheel Movement, Shift
        End If
        If IntMovement <> 0 Then
          MWM = MWM + IntMovement
          MoveMouse
        End If
    End If
End If
End Sub

Public Function dbProcessMessages(Optional ByVal Wait As Boolean = True) As Boolean
If Not ScrollSettings.DS_Enabled Then WaitMessage
End Function

Private Sub ChangeScrollBarValue(ByRef ScrB As Control, _
                                 ByVal Increment As Long)
Dim nVal As Long, oVal As Long
    
    With ScrB
        oVal = .Value
        nVal = .Value + Increment
        If nVal > .Max Then
            nVal = .Max
        ElseIf nVal < .Min Then
            nVal = .Min
        End If
        
        If Not (.Enabled) Then
            If nVal = 0 Then .Value = nVal
        Else
            .Value = nVal
        End If
    End With

End Sub

Public Sub TakeMessagesControl()
Do
    If MoveMP Then WaitMessage
    If Not dbProcessMessages(Not ScrollSettings.DS_Enabled) Then DoEvents
Loop
End Sub

Public Sub dbReplaceColors(ByRef bData() As Long, _
                           ByVal lngColorFind As Long, ByVal Sens As Integer, _
                           ByVal lngColorReplace)
Dim x As Long, y As Long
If Sens = 0 Then
    For y = 0 To UBound(bData, 2)
        For x = 0 To UBound(bData, 1)
            If bData(x, y) = lngColorFind Then
                bData(x, y) = lngColorReplace
            End If
        Next x
    Next y
Else
    For y = 0 To UBound(bData, 2)
        For x = 0 To UBound(bData, 1)
            If CompareColorsLng(bData(x, y), lngColorFind) <= Sens Then
                bData(x, y) = lngColorReplace
            End If
        Next x
    Next y
End If
End Sub

Public Function GetToolIndexByID(ByVal iToolID As Integer) As Integer
Dim i As Integer
For i = 0 To mnuTool.UBound
    If CInt(mnuTool(i).Tag) = iToolID Then Exit For
Next i
GetToolIndexByID = i
End Function

Public Function GetToolIDByIndex(ByVal iToolIndex As Integer) As Integer
GetToolIDByIndex = CInt(mnuTool(iToolIndex).Tag)
End Function

Public Sub ValidateUndoRedo()
Dim CanUndo As Boolean, CanRedo As Boolean
With UndoData
    CanUndo = ((.Index - .FirstIndex) >= 0)
End With
With RedoData
    CanRedo = ((.Index - .FirstIndex) >= 0)
End With
If Not (mnuNoUndoRedo.Checked) Then
    ToolBarButton(GetTLBIndex(cmdUndo)).Enabled = CanUndo
    mnuUndo.Enabled = CanUndo
    mnuUndo.Visible = True
    ToolBarButton(GetTLBIndex(cmdRedo)).Enabled = CanRedo
    mnuRedo.Enabled = CanRedo
    mnuRedo.Visible = True
Else
    ToolBarButton(GetTLBIndex(cmdUndo)).Enabled = False
    ToolBarButton(GetTLBIndex(cmdRedo)).Enabled = False
    mnuUndo.Visible = False
    mnuRedo.Visible = False
End If
End Sub

Public Sub SetToolBar2Visible(ByVal nVal As Boolean)
ToolBar2Visible = nVal
ToolBar2.Visible = nVal
mnuToolBarVis.Checked = nVal
Form_Resize
ApplyScrollBarsValues
End Sub

Public Sub WaitW(ByVal Interval As Long, ByVal Init As Boolean, Optional ByVal DoDoEvents As Boolean = True, Optional ByVal Breakable As Boolean = False)
Static t As Long
If Init Then
    t = GetTickCount
Else
    Do
        If DoDoEvents Then dbProcessMessages (False): DoEvents
        If BreakKeyPressed Then Err.Raise dbCWS, "WaitW", "Operation cancelled."
    Loop Until Abs(GetTickCount - t) >= Interval
End If
End Sub

Public Function MoveMP(Optional ByVal x As Long = &H7FFFFFFF, _
                       Optional ByVal y As Long = &H7FFFFFFF, _
                       Optional ByVal Flags As Long, _
                       Optional ByVal ForceUpdate As Boolean = False) As Boolean
Static mx As Long, my As Long
Static tx As Single, ty As Single
Static vx As Single, vy As Single
Static LastUpdateTime As Long
Dim t As Long
Dim ax As Single, ay As Single
Dim i As Long
Dim PrC As Long
Static vk As Single
Static PowEnL As Single
Static JxVKd5 As Single
Dim MPl As Long, MPt As Long

If x <> &H7FFFFFFF Then
    If Flags And &H1 Then
        mx = mx + (x And &HFFFF)
    Else
        mx = x
    End If
End If
If y <> &H7FFFFFFF Then
    If Flags And &H2 Then
        my = my + (y And &HFFFF)
    Else
        my = y
    End If
End If
If Flags And &H4 Then
    tx = mx
    ty = my
End If

If CBool(Flags And &H8) Or vk = 0 Then
    vk = MoveTimerRes / 16
    PowEnL = ScrollSettings.DS_EnL ^ vk
    JxVKd5 = ScrollSettings.DS_Jestkost * vk / 5
End If

If Not MeEnabled Then Exit Function


If ForceUpdate Then
    Do
        t = mGetTickCount
        If LastUpdateTime = 0 Then LastUpdateTime = t
    Loop Until (t - LastUpdateTime >= MoveTimerRes) Or Not ScrollSettings.DS_Enabled
    ax = -((tx - mx) * JxVKd5)
    ay = -((ty - my) * JxVKd5)
    vx = vx + ax
    vy = vy + ay
    vx = vx * PowEnL
    vy = vy * PowEnL
    
    tx = tx + vx * vk
    ty = ty + vy * vk
    MPt = Round(ty)
    MPl = Round(tx)
    If MP.Left <> MPl Or MP.Top <> MPt Then
        MoveMPWithMouse MPl, MPt
        If MP.AutoRedraw Then
            'Refresh
        End If
    End If
    LastUpdateTime = t 'LastUpdateTime + MoveTimerRes

ElseIf ScrollSettings.DS_Enabled Then
    t = mGetTickCount
    If LastUpdateTime = 0 Then LastUpdateTime = t
    If t - LastUpdateTime < MoveTimerRes Then Exit Function
    i = 0
    Do
        ax = -(tx - mx) * JxVKd5
        ay = -(ty - my) * JxVKd5
        vx = vx + ax
        vy = vy + ay
        vx = vx * PowEnL
        vy = vy * PowEnL
        
        tx = tx + vx * vk
        ty = ty + vy * vk
        t = mGetTickCount
        LastUpdateTime = LastUpdateTime + MoveTimerRes
        i = i + 1
    Loop Until t - LastUpdateTime < MoveTimerRes Or i = 100
    MPt = Round(ty)
    MPl = Round(tx)
    If MP.Left <> MPl Or MP.Top <> MPt Then
        MoveMPWithMouse MPl, MPt
        If MP.AutoRedraw Then
            'Refresh
        End If
    End If
    If i = 100 Then LastUpdateTime = t
    MoveMP = Abs(tx - mx) < 0.5 And Abs(ty - my) < 0.5 And Abs(vx) < 0.05 And Abs(vy) < 0.05
ElseIf Flags And &H4 Then
    MoveMPWithMouse tx, ty
End If
End Function

Public Sub MoveMPWithMouse(ByVal Left As Long, ByVal Top As Long)
Dim Pnt As POINTAPI
If CBool(dbMS(1, 1)) Or CBool(dbMS(2, 1)) Or CBool(dbMS(3, 1)) Then
    GoSub MM
ElseIf ScrollSettings.MouseGlued Then
    If GetCapture = MP.hWnd Then
        GoSub MM
    ElseIf GetCapture = 0 Then
        GetCursorPos Pnt
        If WindowFromPoint(Pnt.x, Pnt.y) = MP.hWnd Then
            GoSub MM
        End If
    End If
End If

MP.Move Left, Top
Exit Sub
MM:
    If Not NaviEnabled Then
        MoveMouse Left - MP.Left, Top - MP.Top, Immediate:=True
    End If
Return
End Sub


Public Sub dbPutPoint(ByVal x As Single, ByVal y As Single, _
            ByVal lngColor As Long, _
            Optional ByVal TempPSet As Boolean = False, _
            Optional ByVal ForceDraw As Boolean = False)
Dim ix As Long, iy As Long
Dim dx As Single, dy As Single
Dim rgb0 As RGBQUAD, rgb00 As RGBQUAD
Dim rgb1 As RGBQUAD
Dim k As Single
Dim tmp As Long
If TempPSet And Not ForceDraw Then Exit Sub

ix = Int(x)
iy = Int(y)
dx = x - ix
dy = y - iy
If lngColor <> -1 Then
    k = (1 - dx) * (1 - dy)
    
    'X.
    '..
    If InPicture(ix, iy) Then
        tmp = dbAlphaBlend(lngColor, Data(ix, iy), Not CByte(k * 255))
        If ForceDraw Then
            dbPSet ix, iy, tmp, dbAsmnuGrid, Not TempPSet, Not MP.AutoRedraw
        Else
            StorePixel ix, iy
            Data(ix, iy) = tmp
        End If
    End If
    '.X
    '..
    If dx > 0 Then
        k = dx * (1 - dy)
        If InPicture(ix + 1, iy) Then
            tmp = dbAlphaBlend(lngColor, Data(ix + 1, iy), Not CByte(k * 255))
            If ForceDraw Then
                dbPSet ix + 1, iy, tmp, dbAsmnuGrid, Not TempPSet, Not MP.AutoRedraw
            Else
                StorePixel ix + 1, iy
                Data(ix + 1, iy) = tmp
            End If
        End If
    End If
    
    '..
    'X.
    If dy > 0 Then
        If InPicture(ix, iy + 1) Then
            k = dy * (1 - dx)
            tmp = dbAlphaBlend(lngColor, Data(ix, iy + 1), Not CByte(k * 255))
            If ForceDraw Then
                dbPSet ix, iy + 1, tmp, dbAsmnuGrid, Not TempPSet, Not MP.AutoRedraw
            Else
                StorePixel ix, iy + 1
                Data(ix, iy + 1) = tmp
            End If
        End If
    End If
    
    '..
    '.X
    If dx > 0 And dy > 0 Then
        If InPicture(ix + 1, iy + 1) Then
            k = dx * dy
            tmp = dbAlphaBlend(lngColor, Data(ix + 1, iy + 1), Not CByte(k * 255))
            If ForceDraw Then
                dbPSet ix + 1, iy + 1, tmp, dbAsmnuGrid, Not TempPSet, Not MP.AutoRedraw
            Else
                StorePixel ix + 1, iy + 1
                Data(ix + 1, iy + 1) = tmp
            End If
        End If
    End If
ElseIf TempPSet Then
    dbPSet ix, iy, Data(ix, iy), dbAsmnuGrid, False, Not MP.AutoRedraw
    If dx > 0 Then
        dbPSet ix + 1, iy, Data(ix + 1, iy), dbAsmnuGrid, False, Not MP.AutoRedraw
    End If
    If dy > 0 Then
        dbPSet ix, iy, Data(ix, iy + 1), dbAsmnuGrid, False, Not MP.AutoRedraw
    End If
    If dx > 0 And dy > 0 Then
        dbPSet ix, iy, Data(ix + 1, iy + 1), dbAsmnuGrid, False, Not MP.AutoRedraw
    End If
End If
End Sub

Public Function InPicture(ByVal x As Long, ByVal y As Long)
InPicture = x >= 0& And y >= 0& And x < intW And y < intH
End Function

Public Function Pixel(ByVal x As Long, ByVal y As Long) As Long
If InPicture(x, y) Then
    Pixel = Data(x, y)
End If
End Function

Public Sub Data_SetPixel(ByVal x As Long, ByVal y As Long, ByVal lngColor As Long)
If PrgDrawMode = dbDrawToBuffer Then
If InPicture(x, y) Then
    StorePixel x, y, lngColor
    Data(x, y) = lngColor
    NeedRefr = True
End If
ElseIf PrgDrawMode = dbDrawDirect Then
    dbPSet x, y, lngColor, dbNoGrid, True, Not MP.AutoRedraw, True
End If
End Sub

Friend Sub ExtractFadeDesc(ByRef FadeDsc As FadeDesc)
On Error Resume Next
FadeDsc.FCount = Pereliv.nmbCount.Value
FadeDsc.Power = Pereliv.nmbPower.Value
FadeDsc.Offset = Pereliv.nmbOffset.Value
FadeDsc.Mode = PerelivMode
FadeDsc.AutoColor1 = CBool(Pereliv.bColor(0).Tag)
FadeDsc.AutoColor2 = CBool(Pereliv.bColor(1).Tag)
End Sub

Friend Sub SendFadeDesc(ByRef FadeDsc As FadeDesc)
With Pereliv
    .nmbCount.Value = FadeDsc.FCount
    .nmbPower.Value = FadeDsc.Power
    .nmbOffset.Value = FadeDsc.Offset
    .Opts(FadeDsc.Mode).Value = True
    .bColor(0).BackColor = ACol(1)
    .bColor(1).BackColor = ACol(2)
    .bColor(0).Tag = CStr(FadeDsc.AutoColor1)
    .bColor(1).Tag = CStr(FadeDsc.AutoColor2)
    .UpdateCaptions
    .UpdateFade
End With
End Sub

Friend Sub LoadFadeDesc(ByRef FadeDsc As FadeDesc, ByRef MsgText As String)
MsgText = "Invalid Fade Count"
FadeDsc.FCount = dbGetSettingEx("Tool", "FadeCount", vbSingle, 1)

MsgText = "Invalid Fade Power"
FadeDsc.Power = dbGetSettingEx("Tool", "FadeDegree", vbSingle, 0.5)

MsgText = "Invalid Fade Offset"
FadeDsc.Offset = dbGetSettingEx("Tool", "FadeOffset", vbSingle, 0)

MsgText = "Invalid Fade Mode"
FadeDsc.Mode = dbGetSettingEx("Tool", "FadeMode", vbInteger, 0)
If FadeDsc.Mode <> 0 And FadeDsc.Mode <> 1 Then Err.Raise 111, "LoadFadeDesc", "Fade Mode"

MsgText = "Invalid boolean value (AutoColor1)"
FadeDsc.AutoColor1 = dbGetSettingEx("Tool", "AutoColor1", vbBoolean, False)
MsgText = "Invalid boolean value (AutoColor2)"
FadeDsc.AutoColor2 = dbGetSettingEx("Tool", "AutoColor2", vbBoolean, False)
End Sub

Friend Sub SaveFadeDesc(ByRef FadeDsc As FadeDesc)
dbSaveSettingEx "Tool", "FadeCount", FadeDsc.FCount
dbSaveSettingEx "Tool", "FadeDegree", FadeDsc.Power
dbSaveSettingEx "Tool", "FadeOffset", FadeDsc.Offset
dbSaveSetting "Tool", "FadeMode", CStr((FadeDsc.Mode))
dbSaveSetting "Tool", "AutoColor1", CStr(FadeDsc.AutoColor1)
dbSaveSetting "Tool", "AutoColor2", CStr(FadeDsc.AutoColor2)
End Sub

Friend Sub LoadHSet(ByRef pHSet As HelixSettings, ByRef MsgText As String)
MsgText = "Invalid Helix.Count"
pHSet.Numb = dbGetSettingEx("Tool", "HelixCount", vbSingle, 5!)

MsgText = "Invalid Helix.RFixed"
pHSet.RFixed = dbGetSettingEx("Tool", "HelixRFixed", vbSingle, 0!)

MsgText = "Invalid Helix.RVar"
pHSet.RK = dbGetSettingEx("Tool", "HelixRVar", vbSingle, 2)

MsgText = "Invalid Helix.RMode"
pHSet.RMode = dbGetSettingEx("Tool", "HelixRMode", vbInteger, 0)
If pHSet.RMode < 0 Or pHSet.RMode > 1 Then Err.Raise 111, "LoadHSet", "Helix.Mode"
End Sub

Friend Sub SaveHSet(ByRef pHSet As HelixSettings)
dbSaveSetting "Tool", "HelixCount", CStr(pHSet.Numb)
dbSaveSetting "Tool", "HelixRFixed", CStr(pHSet.RFixed)
dbSaveSetting "Tool", "HelixRVar", CStr(pHSet.RK)
dbSaveSetting "Tool", "HelixRMode", CStr(pHSet.RMode)
End Sub

Friend Sub FreshActiveColors()
With ActiveColor(ActiveColor.lBound)
    .BackColor = ConvertColorLng(ACol(1))
    .Caption = IIf(gFDSC.AutoColor1, "*", "")
    .MousePointer = IIf(gFDSC.AutoColor1, 99, 0)
End With
With ActiveColor(ActiveColor.lBound + 1)
    .BackColor = ConvertColorLng(ACol(2))
    .Caption = IIf(gFDSC.AutoColor2, "*", "")
    .MousePointer = IIf(gFDSC.AutoColor2, 99, 0)
End With
End Sub

'1-based indexes: 1 is forecolor, 2 is backcolor
Friend Function GetACol(ByVal Index As Integer) As Long
GetACol = ACol(Index)
End Function

Private Function GetAirBrushSize() As Long
Dim s As Long
With frmAero
    s = Val(.aSize.Text)
    If .chkSizePressure.Value Then
        If PenPressure > 0 Then
            s = Round(s * GetPressureLevel)
        End If
    End If
    GetAirBrushSize = s
End With
End Function

Public Function GetPressureLevel() As Single
If MaxPenPressure <= 0 Then
    GetPressureLevel = 1
Else
    GetPressureLevel = PenPressure / MaxPenPressure
End If
End Function


Public Sub ExecuteCmd(ByVal cmd As dbCommands, _
                      Optional ByVal Shift As Integer, _
                      Optional ByVal KeyEvent As dbKeyEvent, _
                      Optional ByRef KeyCode As Integer)
Dim a As POINTAPI, t As Long
Dim oa As POINTAPI
Dim tData() As Long
Dim UB As Long
On Error GoTo eh
GetCursorPos a
oa = a
Select Case Shift
    Case 0
        If Edins(0) Then t = Steps(0) * Zm Else t = Steps(0)
    Case 1
        If Edins(1) Then t = Steps(1) * Zm Else t = Steps(1)
    Case 2
        If Edins(2) Then t = Steps(2) * Zm Else t = Steps(2)
    Case 4
        If Edins(3) Then t = Steps(3) * Zm Else t = Steps(3)
End Select
Select Case cmd
    Case dbCommands.cmdNone
        ShowStatus 1208
    Case dbCommands.cmdoBGToolBar
        GoTo Obsolete 'Obsolete
    Case dbCommands.cmdoBGWorkspace
        GoTo Obsolete
    Case dbCommands.cmdoBMPSettings
        GoTo Obsolete
        'mnuBMPSettings_Click
    Case dbCommands.cmdChangeBCol
        ActiveColor_Click 2
    Case dbCommands.cmdChangeFCol
        ActiveColor_Click 1
    Case dbCommands.cmdClearUndo
        mnuClearUndo_Click
    Case dbCommands.cmdCloseInstance
        Unload Me
    Case dbCommands.cmdCopy
        mnuCopy_Click
        
    Case dbCommands.cmdCursorLeft
        a.x = a.x - t
        KeyCode = 0
    Case dbCommands.cmdCursorUp
        a.y = a.y - t
        KeyCode = 0
    Case dbCommands.cmdCursorRight
        a.x = a.x + t
        KeyCode = 0
    Case dbCommands.cmdCursorDown
        a.y = a.y + t
        KeyCode = 0
    Case dbCommands.cmdScrollRight
        If HScroll.Enabled Then
            If HScroll.Value + t > HScroll.Max Then HScroll.Value = HScroll.Max Else HScroll.Value = HScroll.Value + t
        End If
    Case dbCommands.cmdScrollDown
        If VScroll.Enabled Then
            If VScroll.Value + t > VScroll.Max Then VScroll.Value = VScroll.Max Else VScroll.Value = VScroll.Value + t
        End If
    Case dbCommands.cmdScrollLeft
        If HScroll.Enabled Then
            If HScroll.Value - t < HScroll.Min Then HScroll.Value = HScroll.Min Else HScroll.Value = HScroll.Value - t
        End If
    Case dbCommands.cmdScrollUp
        If VScroll.Enabled Then
            If VScroll.Value - t < VScroll.Min Then VScroll.Value = VScroll.Min Else VScroll.Value = VScroll.Value - t
        End If
        
    Case dbCommands.cmdCyclePoly
        If CurPol.Active And ActiveTool = ToolPoly Then
            CurPol.Active = False
            dbLine x0, y0, XO, YO, -1, True
            dbLine x0, y0, CurPol.bx, CurPol.by, ACol(CurPol.Button)
            If MP.AutoRedraw Then MP.Refresh
        Else
            'vtBeep
        End If

    Case dbCommands.cmdDefPal16
        mnuDefPal16_Click
    Case dbCommands.cmdDefPal256
        mnuDefPal256_Click
    Case dbCommands.cmdDefPalBRH
        mnuQBColors_Click
    Case dbCommands.cmdDefPalSysColors
        mnuSysPal_Click
    Case dbCommands.cmdDeleteSel
        If CurSel.Selected Then dbDeselect False
    Case dbCommands.cmdDeselect
        If CurSel.Selected Then
            StoreFragment CurSel.x1, CurSel.y1, CurSel.x2, CurSel.y2
            dbDeselect True
        End If
    Case dbCommands.cmdDeselect_no_delete
        StoreFragment CurSel.x1, CurSel.y1, CurSel.x2, CurSel.y2
        dbApply dbUseCurSelMode
    Case dbCommands.cmdDisableUndo
        mnuNoUndoRedo_Click
    Case dbCommands.cmdoDllPath
        GoTo Obsolete
        'mnuSetDllPath_Click
    Case dbCommands.cmdEndPoly
        If ActiveTool = ToolPoly Then
            CurPol.Active = False
        End If
    Case dbCommands.cmdExts
        mnuReg_Click
    Case dbCommands.cmdFullReset
        mnuResetAll_Click
    Case dbCommands.cmdKeyboard
        mnuKeyb2_Click
    Case dbCommands.cmdLoadPAL
        mnuLoadPal_Click
    Case dbCommands.cmdLoadSel
        mnuMix_Click
    Case dbCommands.cmdMakePalDef
        mnuPalDef_Click
    Case dbCommands.cmdNew
        mnuNew_Click
    Case dbCommands.cmdNewInstance
        Shell ExePath, vbNormalFocus
    Case dbCommands.cmdNextTool
        Form_Wheel 1, dbStateAlt
    Case dbCommands.cmdNPalEntries
        mnuPalCount_Click
    Case dbCommands.cmdOpen
        mnuLoad_Click
    Case dbCommands.cmdPaintBrush
        If ActiveTool = ToolBrush Then
            If NeedRefr Then
                Refr
                NeedRefr = False
            End If
        End If
    Case dbCommands.cmdPalClearTips
        mnuEmptyTips_Click
    Case dbCommands.cmdPalFillTips
        mnuFillTips_Click
    Case dbCommands.cmdPalReset
        mnuResetPal_Click
    Case dbCommands.cmdPalVisible
        mnuPal_Click
    Case dbCommands.cmdPaste
        mnuPaste_Click
    Case dbCommands.cmdPrevTool
        Form_Wheel -1, dbStateAlt
    Case dbCommands.cmdRedo
        mnuRedo_Click
    Case dbCommands.cmdRefresh
        mnuRefresh_Click
    Case dbCommands.cmdResize
        mnuResize_Click
    Case dbCommands.cmdSave
        mnuSaveFile_Click
    Case dbCommands.cmdSaveAs
        mnuSaveAs_Click
    Case dbCommands.cmdoSavePNG
        GoTo Obsolete
'        mnuSavePNG_Click
    Case dbCommands.cmdoSaveICO
        GoTo Obsolete
        'mnuSaveIcon_Click
    Case dbCommands.cmdSavePAL
        mnuSavePal_Click
    Case dbCommands.cmdoSaveSMB
        GoTo Obsolete
        'mnuSave_Click
    Case dbCommands.cmdSetUndoCount
        mnuUndoLim_Click
    Case dbCommands.cmdStretchPal
        mnuStretchPal_Click
    Case dbCommands.cmdoSyncACol
        GoTo Obsolete
    Case dbCommands.cmdoToggleAutoRedraw
        'mnuAutoRedraw_Click
        GoTo Obsolete
    Case dbCommands.cmdToggleGrid
        'mnuGrid_Click
        'GoTo Obsolete
        DrawGrid
    Case dbCommands.cmdToggleMiddleButtonUse
        If mnuUseWheel(0).Checked Then
            mnuUseWheel_Click 1
        Else
            mnuUseWheel_Click 0
        End If
    Case dbCommands.cmdToolAir
        ChTool ToolAir
    Case dbCommands.cmdToolBarVisible
        mnuToolBarVis_Click
    Case dbCommands.cmdToolBrh
        ChTool ToolBrush
    Case dbCommands.cmdToolCir
        ChTool ToolCircle
    Case dbCommands.cmdToolFad
        ChTool ToolFade
    Case dbCommands.cmdToolFSt
        ChTool ToolFStar
    Case dbCommands.cmdToolGet
        ChTool ToolColorSel
    Case dbCommands.cmdToolHel
        ChTool ToolHelix
    Case dbCommands.cmdToolHFd
        ChTool ToolHFade
    Case dbCommands.cmdoToolHot
        GoTo Obsolete
        'ChTool ToolHot
    Case dbCommands.cmdToolLin
        ChTool ToolLine
    Case dbCommands.cmdToolPal
        ChTool ToolPal
    Case dbCommands.cmdToolPen
        ChTool ToolPen
    Case dbCommands.cmdToolPnt
        ChTool ToolPaint
    Case dbCommands.cmdToolPol
        ChTool ToolPoly
    Case dbCommands.cmdToolProps
        mnuToolOpts_Click
    Case dbCommands.cmdToolRec
        ChTool ToolRect
    Case dbCommands.cmdToolSel
        ChTool ToolSel
    Case dbCommands.cmdToolSta
        ChTool ToolStar
    Case dbCommands.cmdToolTxt
        ChTool ToolText
    Case dbCommands.cmdToolVFd
        ChTool ToolVFade
    Case dbCommands.cmdoToolWav
        'ChTool ToolWav
        GoTo Obsolete
    Case dbCommands.cmdToolOrg
        ChTool ToolOrg
    Case dbCommands.cmdToolPrg
        ChTool ToolProg
    Case dbCommands.cmdUndo
        mnuUnDo_Click
    Case dbCommands.cmdWhatsThis
        'mnuWhatsThis_Click
        GoTo Obsolete
    Case dbCommands.cmdZoom
        mnuZoom_Click
    Case dbCommands.cmdZoomIn
        ReZoomPtr Zm + 1
    Case dbCommands.cmdZoomOut
        ReZoomPtr Zm - 1
    Case dbCommands.cmdLMB
        If (KeyEvent = KeyUp) And (dbMS(1, 2) = dbButtonDown) Then
            Mouse_Event dbMOUSEEVENT_LEFTUP, a.x, a.y, 2, 0
            dbMS(1, 2) = dbButtonUp
        ElseIf (KeyEvent = KeyDown) And (dbMS(1, 2) = dbButtonUp) Then
            'If dbMS(2, 2) = dbButtonDown Or dbMS(2, 1) = dbButtonDown Then
            '    vtbeep
            '    Exit Sub
            'End If
            Mouse_Event dbMOUSEEVENT_LEFTDOWN, a.x, a.y, 2, 0
            dbMS(1, 2) = dbButtonDown
        End If
    
    Case dbCommands.cmdMMB
        If (KeyEvent = KeyUp) And (dbMS(3, 2) = dbButtonDown) Then
            Mouse_Event dbMOUSEEVENTF_MIDDLEUP, a.x, a.y, 2, 0
            dbMS(3, 2) = dbButtonUp
        ElseIf (KeyEvent = KeyDown) And (dbMS(3, 2) = dbButtonUp) Then
            'If dbMS(1, 2) = dbButtonDown Or dbMS(1, 1) = dbButtonDown Then
            '    vtbeep
            '    Exit Sub
            'End If
            Mouse_Event dbMOUSEEVENTF_MIDDLEDOWN, a.x, a.y, 2, 0
            dbMS(3, 2) = dbButtonDown
        End If
    
    Case dbCommands.cmdRMB
        If (KeyEvent = KeyUp) And (dbMS(2, 2) = dbButtonDown) Then
            Mouse_Event dbMOUSEEVENTF_RIGHTUP, a.x, a.y, 2, 0
            dbMS(2, 2) = dbButtonUp
        ElseIf (KeyEvent = KeyDown) And (dbMS(2, 2) = dbButtonUp) Then
            'If dbMS(1, 2) = dbButtonDown Or dbMS(1, 1) = dbButtonDown Then
            '    vtbeep
            '    Exit Sub
            'End If
            Mouse_Event dbMOUSEEVENTF_RIGHTDOWN, a.x, a.y, 2, 0
            dbMS(2, 2) = dbButtonDown
        End If
        
    Case dbCommands.cmdoDelLastCharge
        GoTo Obsolete
    Case dbCommands.cmdoClearCharges
        GoTo Obsolete
    Case dbCommands.cmdClearPic
        mnuClear_Click
    Case dbCommands.cmdMinimize
        Me.WindowState = vbMinimized
    Case dbCommands.cmdDrawBBg
        mnuDrawBBg_Click
    Case dbCommands.cmdDrawBg
        mnuDrawBG_Click
    Case dbCommands.cmdPrgDrawings
        mnuPrgDraw_Click
    Case dbCommands.cmdoPrgPhys
        GoTo Obsolete
        'mnuRhysPrg_Click
    Case dbCommands.cmdEffect0
        mnuEffect_Click 0
    Case dbCommands.cmdEffect1
        mnuEffect_Click 1
    Case dbCommands.cmdEffect2
        mnuEffect_Click 2
    Case dbCommands.cmdEffect3
        mnuEffect_Click 3
    Case dbCommands.cmdEffect4
        mnuEffect_Click 4
    Case dbCommands.cmdEffect5
        mnuEffect_Click 5
    Case dbCommands.cmdEffect6
        mnuEffect_Click 6
    Case dbCommands.cmdEffect7
        mnuEffect_Click 7
    Case dbCommands.cmdEffect8
        mnuEffect_Click 8
    Case dbCommands.cmdEffect9
        mnuEffect_Click 9
    Case dbCommands.cmdEffect10
        mnuEffect_Click 10
    Case dbCommands.cmdRepLastEffect
        If mnuLastEffect.Enabled Then mnuLastEffect_Click
    Case dbCommands.cmdSaveSel
        If CurSel.Selected Then
            mnuSaveSel_Click
        End If
    Case dbCommands.cmdSelectAll
        mnuSelectAll_Click
    Case dbCommands.cmdCropSel
        mnuCrop_Click
    Case dbCommands.cmdCapture
        mnuCapture_Click
    Case dbCommands.cmdCapturePointedWindow
        dbCapture 0, tData
        dbMakeSelData 0, 0, tData
        dbPutSel
    Case dbCommands.cmdCaptureActiveWindow
        dbCapture 1, tData
        dbMakeSelData 0, 0, tData
        dbPutSel
    Case dbCommands.cmdCaptureScreen
        dbCapture 2, tData
        dbMakeSelData 0, 0, tData
        dbPutSel
    Case dbCommands.cmdCapturePoint
        ChangeActiveColor 1, CapturePixel(a.x, a.y), False
    Case dbCommands.cmdDynamicDialog
        mnuDynamicScr_Click
    Case dbCommands.cmdIdleMessage
        mnuIdleMessage_Click
    Case dbCommands.cmdTexMode
        mnuTexMode_Click
    Case dbCommands.cmdResetOrg
        mnuResetOrg_Click
    Case dbCommands.cmdRestoreOrg
        mnuRestoreOrg_Click
    Case dbCommands.cmdExtremeSave
        'Implemented in hook procedure
        'because it is extreme
    Case dbCommands.cmdFormula
        If MeEnabled Then
            mnuFormula_Click
        End If
    Case dbCommands.cmdOAutoScrolling
        'mnuSetAutoScrolling_Click
        GoTo Obsolete
    Case dbCommands.cmdEffect11
        dbEffect 11
    Case dbCommands.cmdGlueMouse
        mnuMouseAttr_Click
    Case dbCommands.cmdSelMoveTo
        If mnuSelAll.Visible Then mnuSelMoveTo_Click
    Case dbCommands.cmdSelHCenter
        If mnuSelAll.Visible Then mnuSelCenterHorz_Click
    Case dbCommands.cmdSelVCenter
        If mnuSelAll.Visible Then mnuSelCenterVert_Click
    Case dbCommands.cmdSelClear
        If mnuSelAll.Visible Then mnuSelClear_Click
    Case dbCommands.cmdSelResize
        If mnuSelAll.Visible Then mnuSelResize_Click
    Case dbCommands.cmdSelEdit
        If mnuSelAll.Visible Then mnuSelEditSel_Click
    Case dbCommands.cmdSelToggleRedraw
        If mnuSelAll.Visible Then mnuSelAutoRepaint_Click
    Case dbCommands.cmdSelShow
        If mnuSelAll.Visible Then mnuSelShow_Click
        
    Case dbCommands.cmdDrawWaves
        mnuDrawWaves_Click
    
    Case dbCommands.cmdNaviMode
        If KeyEvent = KeyDown Then
            NaviStart
        Else
            NaviStop
        End If
    
    Case dbCommands.cmdSelShowSize
        If mnuSelAll.Visible Then mnuSelShowSize_Click
        
    Case dbCommands.cmdWheelUp
        Form_Wheel Movement:=1, Shift:=Shift
    Case dbCommands.cmdWheelDown
        Form_Wheel Movement:=-1, Shift:=Shift
    
    Case dbCommands.cmdZoom1
        ReZoom 1
    
    Case Else
        MsgBox "This action is forgotten to implement (" + GetActionDescription(cmd) + "). Please mail to the author to VT-Dbnz@yandex.ru about this.", vbExclamation, "Forgotten!! :("
End Select

If oa.x <> a.x Or oa.y <> a.y Then
    SetCursorPos a.x, a.y
End If

Exit Sub
Obsolete:
dbMsgBox 1199, vbInformation
Exit Sub
eh:
MsgError
End Sub

Private Sub ToScreenCoords(Pnt As POINTAPI, hWnd As Long)
'Dim rct As RECT
'GetWindowRect hwnd, rct
'Pnt.x = Pnt.x + rct.Left
'Pnt.y = Pnt.y + rct.Top
ClientToScreen hWnd, Pnt
End Sub

Public Sub ChangeDynScrolling(ByVal bNew As Boolean)
ScrollSettings.DS_Enabled = bNew
MPMover.Enabled = bNew
'mnuDynamicScr.Checked = bNew
If Not bNew Then ApplyScrollBarsValues
End Sub

Public Sub ShowGreeting()
Dim h As Long
h = Hour(Time)
Select Case h
    Case 0 To 4
        ShowStatus 10092, , 4
    Case 5 To 7
        ShowStatus 10093, , 4
    Case 8 To 10
        ShowStatus 10094, , 4
    Case 11 To 17
        ShowStatus 10095, , 4
    Case 17 To 23
        ShowStatus 10096, , 4
End Select
End Sub

Public Sub UserMadeAction()
LastUserActionTime = GetTickCount
End Sub

Public Sub ShowIdleMessage(Optional ByVal HoldTime As Long = 0)
Dim MsgText As String
Dim ExtremeSaveKeys() As String
Dim sExtremeSaveKeys As String
Dim nExtremeSaveKeys As Long
Const First_Message = 10097
Static Last_Message As Long

If Last_Message = 0 Then
    On Error Resume Next
    Last_Message = First_Message
    Err.Clear
    Do While Len(GRSF(Last_Message + 1, True)) > 1
        If Err.Number <> 0 Then Exit Do
        Last_Message = Last_Message + 1
    Loop
'    Debug.Print "The last idle message res id is " + CStr(Last_Message) + "."
End If

'On Error GoTo eh
nExtremeSaveKeys = Keyb.ListKeys(cmdExtremeSave, ExtremeSaveKeys)
If nExtremeSaveKeys = 0 Then
    sExtremeSaveKeys = GRSF(2427)
Else
    sExtremeSaveKeys = Join(ExtremeSaveKeys, ", ")
End If
MsgText = grs(Int(Rnd * (Last_Message - First_Message + 1) + First_Message), _
              "%ExtremeSaveKey%", sExtremeSaveKeys)
ShowStatus MsgText, , HoldTime
Exit Sub
eh:
MsgError
End Sub

Public Sub MoveDataOrg(ByRef Data() As Long, ByVal XMovement As Long, ByVal YMovement As Long)
Dim w As Long, h As Long
Dim tmpData() As Long
Dim XM As Long, YM As Long
Dim Xp As Long, Yp As Long
Dim x As Long, y As Long
Dim y0 As Long, y1 As Long
Dim tmpDataV() As Long 'for mapping
Dim DataV() As Long 'for mapping

TestDims Data
w = UBound(Data, 1) + 1
h = UBound(Data, 2) + 1
XM = XMovement - Int(XMovement / w) * w
YM = YMovement - Int(YMovement / h) * h
If XM = 0 And YM = 0 Then Exit Sub

On Error GoTo eh
ReDim tmpData(0 To w - 1, 0 To h - 1)
SwapArys AryPtr(tmpData), AryPtr(Data)
ConstructAry AryPtr(tmpDataV), VarPtr(tmpData(0, 0)), 4, w * h
ConstructAry AryPtr(DataV), VarPtr(Data(0, 0)), 4, w * h


Xp = w - XM
Yp = h - YM
For y = 0 To YM - 1
    y0 = y * w
    y1 = (Yp + y) * w
    For x = 0 To XM - 1
        DataV(y1 + Xp + x) = tmpDataV(x + y0)
    Next x
    For x = XM To w - 1
        DataV(y1 + x - XM) = tmpDataV(x + y0)
    Next x
Next y
For y = YM To h - 1
    y0 = y * w
    y1 = (y - YM) * w
    For x = 0 To XM - 1
        DataV(Xp + x + y1) = tmpDataV(x + y0)
    Next x
    For x = XM To w - 1
        DataV(x - XM + y1) = tmpDataV(x + y0)
    Next x
Next y
UnReferAry AryPtr(DataV)
UnReferAry AryPtr(tmpDataV)
Exit Sub

eh:
UnReferAry AryPtr(DataV)
UnReferAry AryPtr(tmpDataV)
ErrRaise "MoveDataOrg"
End Sub

Public Sub ToolPrg_MouseEvent(ByVal x As Long, ByVal y As Long, _
                              ByVal Button As Long, _
                              ByVal Shift As Long, _
                              Optional ByVal nEvent As dbMouseEvent = dbEvMouseUp, _
                              Optional ByVal PenPressure As Single)
Dim UB As Long, UBV As Long
Static EV As New clsEVal
Dim i As Long
UB = -1
UBV = -1
If AryDims(AryPtr(prgToolProg.Vars)) = 1 Then
    UBV = UBound(prgToolProg.Vars)
End If
On Error GoTo eh
If IsEmpty(prgToolProg.Code) And nEvent = dbEvMouseUp Then
    mnuToolOpts_Click
    GoTo ExitHere
End If

PrgDrawMode = dbDrawDirect
If UBV >= 15 Then
    With prgToolProg
        .Vars(0).Value = dblX0
        .Vars(1).Value = dblY0
        .Vars(2).Value = x / Zm
        .Vars(3).Value = y / Zm
        .Vars(4).Value = Button
        .Vars(5).Value = Shift
        .Vars(6).Value = intW \ 2 'cx
        .Vars(7).Value = intH \ 2 'cy
        .Vars(8).Value = intW 'w
        .Vars(9).Value = intH 'h
        .Vars(10).Value = ACol(1) 'FC
        .Vars(11).Value = ACol(2) 'BC
        .Vars(12).Value = nEvent
        .Vars(13).Value = PenPressure
        .Vars(14).Value = MWM
        If (.Vars(15).Value <> 0#) <> ScrollSettings.CancelWheelScroll Then
            .Vars(15).Value = CDbl(ScrollSettings.CancelWheelScroll)
        End If
    End With
Else
    GoTo ExitHere
End If
If nEvent = dbEvMouseDown Then
    EV.DrawTemp = False
End If

EV.ExecuteSMP prgToolProg

MWM = prgToolProg.Vars(14).Value
ScrollSettings.CancelWheelScroll = CBool(prgToolProg.Vars(15).Value)
MeEnabled = True
If nEvent = dbEvMouseUp Then
    EV.NewUndo
    EV.DrawTemp = False
    ScrollSettings.CancelWheelScroll = False
End If
ExitHere:
PrgDrawMode = dbDrawToBuffer
Exit Sub
Resume
eh:
If Err.Number = dbCWS Then
    ShowStatus 1249, , 5 'Program execution has been cancelled.
    PrgDrawMode = dbDrawToBuffer
    Exit Sub
End If
ErrRaise "ToolPrg_MouseEvent"
PrgDrawMode = dbDrawToBuffer
End Sub

Public Sub dbMatrix2(ByRef pDataPic() As Long, ByRef pDataSel() As Long, _
                     ByVal SelOrgX As Long, ByVal SelOrgY As Long, _
                     ByRef Matrix() As Double, _
                     ByRef tSelData() As Long, _
                     ByVal x1 As Long, ByVal y1 As Long, _
                     ByVal x2 As Long, ByVal y2 As Long)
Dim x As Long, y As Long
Dim tx As Long, ty As Long
Dim rgb1 As RGBQUAD
Dim tmpSelData() As RGBQUAD
Dim xf As Long, yf As Long
Dim xt As Long, yt As Long
Dim wp As Long, hp As Long
Dim WS As Long, Hs As Long

Dim rrs As Double, rgs As Double, rbs As Double
Dim grs As Double, ggs As Double, gbs As Double
Dim brs As Double, bgs As Double, bbs As Double

Dim rrp As Double, rgp As Double, rbp As Double
Dim grp As Double, ggp As Double, gbp As Double
Dim brp As Double, bgp As Double, bbp As Double

Dim r As Long, g As Long, b As Long

Dim Pd() As RGBQUAD 'picture
Dim Sd() As RGBQUAD 'selection
Dim Od() As RGBQUAD 'output

On Error GoTo eh
ConstructAry AryPtr(Pd), VarPtr(pDataPic(0, 0)), 4&, UBound(pDataPic, 1) + 1, UBound(pDataPic, 2) + 1
ConstructAry AryPtr(Sd), VarPtr(pDataSel(0, 0)), 4&, UBound(pDataSel, 1) + 1, UBound(pDataSel, 2) + 1
ConstructAry AryPtr(Od), VarPtr(tSelData(0, 0)), 4&, UBound(pDataSel, 1) + 1, UBound(pDataSel, 2) + 1


rrs = Matrix(0, 0)
rgs = Matrix(0, 1)
rbs = Matrix(0, 2)
grs = Matrix(1, 0)
ggs = Matrix(1, 1)
gbs = Matrix(1, 2)
brs = Matrix(2, 0)
bgs = Matrix(2, 1)
bbs = Matrix(2, 2)

rrp = Matrix(0, 3)
rgp = Matrix(0, 4)
rbp = Matrix(0, 5)
grp = Matrix(1, 3)
ggp = Matrix(1, 4)
gbp = Matrix(1, 5)
brp = Matrix(2, 3)
bgp = Matrix(2, 4)
bbp = Matrix(2, 5)


wp = UBound(pDataPic, 1) + 1
hp = UBound(pDataPic, 2) + 1
WS = UBound(pDataSel, 1) + 1
Hs = UBound(pDataSel, 2) + 1


'ReDim tmpSelData(0 To WS - 1, 0 To Hs - 1)
'CopyMemory tmpSelData(0, 0), pDataSel(0, 0), WS * Hs * 4&

For y = y1 To y2
    ty = y + SelOrgY
    For x = x1 To x2
        tx = x + SelOrgX
        r = Sd(x, y).rgbRed * rrs + Sd(x, y).rgbGreen * rgs + Sd(x, y).rgbBlue * rbs + _
            Pd(tx, ty).rgbRed * rrp + Pd(tx, ty).rgbGreen * rgp + Pd(tx, ty).rgbBlue * rbp
        g = Sd(x, y).rgbRed * grs + Sd(x, y).rgbGreen * ggs + Sd(x, y).rgbBlue * gbs + _
            Pd(tx, ty).rgbRed * grp + Pd(tx, ty).rgbGreen * ggp + Pd(tx, ty).rgbBlue * gbp
        b = Sd(x, y).rgbRed * brs + Sd(x, y).rgbGreen * bgs + Sd(x, y).rgbBlue * bbs + _
            Pd(tx, ty).rgbRed * brp + Pd(tx, ty).rgbGreen * bgp + Pd(tx, ty).rgbBlue * bbp
        If r < 0& Then r = 0&
        If r > 255& Then r = 255&
        If g < 0& Then g = 0&
        If g > 255& Then g = 255&
        If b < 0& Then b = 0&
        If b > 255& Then b = 255&
        Od(x, y).rgbRed = r
        Od(x, y).rgbGreen = g
        Od(x, y).rgbBlue = b
    Next x
Next y
UnReferAry AryPtr(Pd), False
UnReferAry AryPtr(Sd), False
UnReferAry AryPtr(Od), False

Exit Sub
eh:
UnReferAry AryPtr(Pd), False
UnReferAry AryPtr(Sd), False
UnReferAry AryPtr(Od), False
ErrRaise "dbMatrix2"
End Sub

Public Sub AnimateSelRect()
Const nIter As Long = 16
Static ToColor As RGBTriInt
Static CurColor As RGBTriInt
Static Initd As Boolean
Dim i As Long
If Not Initd Then
    ToColor.rgbBlue = Int(Rnd * 255)
    ToColor.rgbGreen = Int(Rnd * 255)
    ToColor.rgbRed = Int(Rnd * 255)
    CurColor.rgbBlue = 255
    CurColor.rgbGreen = 255
    CurColor.rgbRed = 255
    Initd = True
End If

For i = 1 To nIter
    CurColor.rgbRed = CurColor.rgbRed + Sgn(ToColor.rgbRed - CurColor.rgbRed)
Next i
If CurColor.rgbRed = ToColor.rgbRed Then
    ToColor.rgbRed = Int(Rnd * 255)
End If

For i = 1 To nIter
    CurColor.rgbGreen = CurColor.rgbGreen + Sgn(ToColor.rgbGreen - CurColor.rgbGreen)
Next i
If CurColor.rgbGreen = ToColor.rgbGreen Then
    ToColor.rgbGreen = Int(Rnd * 255)
End If

For i = 1 To nIter
    CurColor.rgbBlue = CurColor.rgbBlue + Sgn(ToColor.rgbBlue - CurColor.rgbBlue)
Next i
If CurColor.rgbBlue = ToColor.rgbBlue Then
    ToColor.rgbBlue = Int(Rnd * 255)
End If

SelRect.BorderColor = RGB(CurColor.rgbRed, CurColor.rgbGreen, CurColor.rgbBlue)
End Sub


Public Function SelectionPresent() As Boolean
SelectionPresent = CurSel.Selected And (AryDims(AryPtr(CurSel_SelData)) = 2)
End Function

Public Sub vtClearType(ByRef IOData() As Long, ByVal TexMode As Boolean, Optional ByVal Anti As Boolean = False)
Dim RGBData() As RGBQUAD
Dim w As Long, h As Long
Dim x As Long, y As Long

w = UBound(IOData, 1) + 1
h = UBound(IOData, 2) + 1
ReDim RGBData(0 To w - 1, 0 To h - 1)
CopyMemory RGBData(0, 0), IOData(0, 0), w * h * 4&
If Anti Then
    For y = 0 To h - 1
        If TexMode Then
            IOData(0, y) = RGB((RGBData(0, y).rgbBlue * 2& + RGBData(1&, y).rgbBlue) \ 3&, _
                               RGBData(0, y).rgbGreen, _
                               (RGBData(0, y).rgbRed * 2& + RGBData(w - 1&, y).rgbRed) \ 3&)
            IOData(w - 1&, y) = RGB((RGBData(w - 1&, y).rgbBlue * 2& + RGBData(0, y).rgbBlue) \ 3&, _
                                    RGBData(w - 1&, y).rgbGreen, _
                                    (RGBData(w - 1&, y).rgbRed * 2& + RGBData(w - 2&, y).rgbRed) \ 3&)
        Else
            IOData(0, y) = RGB((RGBData(0, y).rgbBlue * 2 + RGBData(1, y).rgbBlue) \ 3&, _
                               RGBData(0, y).rgbGreen, _
                               RGBData(0, y).rgbRed)
            IOData(w - 1&, y) = RGB(RGBData(w - 1&, y).rgbBlue, _
                                    RGBData(w - 1&, y).rgbGreen, _
                                    (RGBData(w - 1&, y).rgbRed * 2& + RGBData(w - 2&, y).rgbRed) \ 3&)
        End If
        For x = 1& To w - 2&
            IOData(x, y) = RGB((RGBData(x, y).rgbBlue * 2& + RGBData(x + 1&, y).rgbBlue) \ 3&, _
                               RGBData(x, y).rgbGreen, _
                               (RGBData(x, y).rgbRed * 2& + RGBData(x - 1&, y).rgbRed) \ 3&)
        Next x
        ShowProgress y * 100& \ (h - 1&), Not MeEnabled
    Next y
Else
    For y = 0 To h - 1
        If TexMode Then
            IOData(0, y) = RGB((RGBData(0, y).rgbBlue * 2& + RGBData(w - 1&, y).rgbBlue) \ 3&, _
                               RGBData(0, y).rgbGreen, _
                               (RGBData(0, y).rgbRed * 2& + RGBData(1, y).rgbRed) \ 3&)
            IOData(w - 1&, y) = RGB((RGBData(w - 1&, y).rgbBlue * 2& + RGBData(w - 2&, y).rgbBlue) \ 3&, _
                                    RGBData(w - 1&, y).rgbGreen, _
                                    (RGBData(w - 1&, y).rgbRed * 2& + RGBData(0, y).rgbRed) \ 3&)
        Else
            IOData(0, y) = RGB(RGBData(0, y).rgbBlue, _
                               RGBData(0, y).rgbGreen, _
                               (RGBData(0, y).rgbRed * 2& + RGBData(1, y).rgbRed) \ 3&)
            IOData(w - 1&, y) = RGB((RGBData(w - 1&, y).rgbBlue * 2& + RGBData(w - 2&, y).rgbBlue) \ 3&, _
                                    RGBData(w - 1&, y).rgbGreen, _
                                    RGBData(w - 1&, y).rgbRed)
        End If
        For x = 1& To w - 2&
            IOData(x, y) = RGB((RGBData(x, y).rgbBlue * 2& + RGBData(x - 1&, y).rgbBlue) \ 3&, _
                               RGBData(x, y).rgbGreen, _
                               (RGBData(x, y).rgbRed * 2& + RGBData(x + 1&, y).rgbRed) \ 3&)
        Next x
        ShowProgress y * 100& \ (h - 1&), Not MeEnabled
    Next y
End If
End Sub

'Filename is the relative path to the backup directory
Public Sub BuildBackup(Optional ByVal File As String, _
                       Optional ByVal ShowMessage As Boolean = True)
If Len(File) = 0 Then
    File = BackUpFileName(GetLastBackupNumber)
Else
    File = BackUpPath + File
End If
SaveSMB Data, File
If ShowMessage Then dbMsgBox 2428, vbInformation
End Sub

Public Function RestoreBackUp(ByVal i As Long, Optional ByVal File As String) As Boolean
If Len(File) = 0 Then
    File = BackUpFileName(i)
Else
    File = BackUpPath + File
End If
If Not FileExists(File) Then Err.Raise 1111, "RestoreBackUp", "File" + vbNewLine + File + vbNewLine + " not found."
LoadFile File
End Function

Public Function GetLastBackupNumber() As Long
Dim i As Long
Dim tmp As String
Dim MaxI As Long
On Error GoTo eh
'i = 1000
tmp = Dir(BackUpPath + "BKP????.smb")
MaxI = -1
Do While Len(tmp) > 0
    On Error Resume Next
    i = Val(Left$(Mid$(tmp, Len(tmp) - 7), 4))
    If i > MaxI Then MaxI = i
    tmp = Dir
Loop
i = MaxI + 1
If i > 1000& Then Err.Raise 1111, "GetLastBackUpNumber", "Too many files!!!"
GetLastBackupNumber = i
Exit Function
eh:
MsgError
GetLastBackupNumber = 0
End Function

'returns the number of files found
'FileList will contain the list of found files
Public Function EnumBackUps(ByRef FileList() As String) As Long
Dim l As Long
Dim BP As String
BP = BackUpPath
l = 0
ReDim FileList(0 To 0)
FileList(l) = Dir(BP + "BKP????.smb")
Do While Len(FileList(l)) > 0
    FileList(l) = BP + FileList(l)
    l = l + 1
    ReDim Preserve FileList(0 To l)
    FileList(l) = Dir
Loop
If l = 0 Then Erase FileList
EnumBackUps = l
End Function

Public Function BackUpFileName(ByVal i As Long) As String
BackUpFileName = BackUpPath + "BKP" + VedNull(i, 4) + ".smb"
End Function

Public Function BackUpPath() As String
Static BP As String
If Len(BP) = 0 Then
    BP = ValFolder(TempPath) + "BackUps\"
End If
If Not FolderExists(BP) Then
    CreateFolder BP
End If
BackUpPath = BP
End Function

Public Function KillBackUp(ByVal i As Long, Optional ByVal File As String) As Boolean
If Len(File) = 0 Then
    File = BackUpFileName(i)
Else
    File = BackUpPath + File
End If
'If Not FileExists(File) Then Err.Raise 1111, "RestoreBackUp", "File" + vbNewLine + File + vbNewLine + " not found."
Kill File
End Function

Public Sub LoadAutoScrollFieldSize()
'AutoScroll_Field_Size =
ScrollSettings.ASS.GapLef = dbGetSettingEx("Options", "AutoScrollFieldSize", vbLong, 200&)
ScrollSettings.ASS.GapRig = ScrollSettings.ASS.GapLef
ScrollSettings.ASS.GapTop = ScrollSettings.ASS.GapLef
ScrollSettings.ASS.GapBot = ScrollSettings.ASS.GapLef
If ScrollSettings.ASS.GapLef < -10000 Or ScrollSettings.ASS.GapLef > 10000 Then
    ScrollSettings.ASS.GapLef = 100
    ScrollSettings.ASS.GapRig = ScrollSettings.ASS.GapLef
    ScrollSettings.ASS.GapTop = ScrollSettings.ASS.GapLef
    ScrollSettings.ASS.GapBot = ScrollSettings.ASS.GapLef
    Err.Raise 111, "LoadAutoScrollFieldSize", "Incorrect value for AutoScroll_Field_Size"
End If
End Sub

Public Sub ExtractData(ByRef DataTo() As Long, _
                       Optional ByVal DecIntens As Boolean = False)
Dim clsEf As New clsLinColorStretch
DataTo = Data
If DecIntens Then
    CancelDoEvents True
    dbApplyGamma DataTo, 0.5
    RestoreDoEvents
End If
End Sub


Public Sub RestoreMP()
If MPhDefBitmap = 0 Then Exit Sub
Dim hCurBitmap As Long
hCurBitmap = SelectObject(MP.hDC, MPhDefBitmap)
DeleteObject hCurBitmap
MPhDefBitmap = 0
UnReferAry AryPtr(MPBitsRGB), False
UnReferAry AryPtr(MPBitsLNG), False
End Sub

'Creates the MP bits. If cancels autoredraw then returns False
Public Function MakeARMP() As Boolean
Dim hBitmap As Long
Dim bmi As BITMAPINFO
Dim ptrBits As Long
Dim Width As Long, Height As Long
If MPhDefBitmap <> 0 Then RestoreMP
If Not MP.AutoRedraw Then Err.Raise 1111, "MakeARMP", "MP is not autoredraw!"
MakeARMP = True
UpdateWH
Width = intW * Zm
Height = intH * Zm
MP.Move MP.Left, MP.Top, Width, Height
MPBitsWidth = Width
MPBitsHeight = Height
With bmi.bmiHeader
    .biSize = Len(bmi.bmiHeader)
    .biWidth = Width
    .biHeight = -Height
    .biBitCount = 32
    .biPlanes = 1
    .biSizeImage = Width * Height * 4
End With
hBitmap = CreateDIBSection(MP.hDC, bmi, DIB_RGB_COLORS, VarPtr(ptrBits), 0, 0)
If hBitmap = 0 Then
    Err.Raise 1111, "MakeARMP", "Bitmap creation failed!"
End If
MPhDefBitmap = SelectObject(MP.hDC, hBitmap)
If MPhDefBitmap = 0 Then
    Err.Raise 1111, "MakeARMP", "SelectObject failed!"
End If
ConstructAry AryPtr(MPBitsRGB), ptrBits, 4, Width * Height
ConstructAry AryPtr(MPBitsLNG), ptrBits, 4, Width * Height
End Function

'DOES NOT check the range
'Call GDIFlush before use
'Call mp.refresh after use
Public Sub FastPSet(ByVal x As Long, _
                    ByVal y As Long, _
                    ByVal Zm As Long, _
                    ByVal Color As Long)
Static xc As Long
Static yc As Long
Dim tColor As Long
Dim Ofc As Long
If MPhDefBitmap = 0& Then Exit Sub
x = x * Zm
y = y * Zm
MPBitsLNG(x + y * MPBitsWidth) = Color
If Zm > 1& Then
    For yc = 0& To Zm - 1
        Ofc = (y + yc) * MPBitsWidth
        For xc = 0& To Zm - 1&
            MPBitsLNG(x + xc + Ofc) = Color
        Next xc
    Next yc
End If
End Sub

Public Sub SendDataToMP()
UpdateRegion 0, 0, intW - 1, intH - 1
End Sub

Public Sub ClearPic(ByRef Data() As Long, ByVal BGColor As Long)
Dim x As Long
Dim y As Long
Dim i As Long
Dim w As Long, h As Long
TestDims Data
w = UBound(Data, 1) + 1
h = UBound(Data, 2) + 1
For x = 0 To w - 1
    Data(x, 0) = BGColor
Next x
For y = 1 To h - 1
    CopyMemory Data(0, y), Data(0, 0), w * 4&
Next y
End Sub

Public Sub ClearMP()
If MP.AutoRedraw Then
    ZeroMemory MPBitsLNG(0), intH * intW * Zm * Zm
Else
    MP.Cls
End If
End Sub

'does not check for bounds!
'Does not support lngcolor=-1
Public Sub dbPsetFast(ByVal x As Long, _
                      ByVal y As Long, _
                      ByVal lngColor As Long, _
                      ByVal Store As Boolean, _
                      ByVal Draw As Boolean, _
                      ByVal MPAutoRedraw As Boolean)
'
If Store Then
    Data(x, y) = lngColor
End If
If Draw Then
    If MPAutoRedraw Then
        FastPSet x, y, Zm, lngColor
    Else
        If Zm = 1 Then
            SetPixel MP.hDC, x, y, ConvertColorLng(lngColor)
            If x And &H100& And MainModule.VistaSetPixelBugDetected Then
              If Not MP.AutoRedraw Then MP.Line (x, y)-(x, y), ConvertColorLng(lngColor), BF
            End If

        Else
            MP.Line (x * Zm, y * Zm)-(x * Zm + Zm - 1, y * Zm + Zm - 1), ConvertColorLng(lngColor), BF
        End If
    End If
End If
End Sub

'Allocates memory for seloutput only if neccessary
'Use one-dimensional array for SelOutput!
Public Sub ProcessSelData(ByVal SelMode As dbSelMode, _
                          ByVal x1 As Long, ByVal y1 As Long, _
                          ByVal x2 As Long, ByVal y2 As Long, _
                          ByRef SelOutPut() As Long)
Dim x As Long, y As Long 'loop variables
Dim xd As Long, yd As Long 'Data-position loop variables
Dim x0 As Long, y0 As Long 'selection origin
Dim SelW As Long, SelH As Long 'selection width and height
Dim ProcessW As Long, ProcessH As Long 'Width and Height to process
Dim OfcYSel As Long 'Offset given by Y in the selection output data
Dim OfcXPic As Long, OfcYPic As Long '
Dim SelOutputRGB() As RGBQUAD 'mapped to SelOutput
Dim SelDataRGB() As RGBQUAD 'Mapped to CurSel_SelData
Dim DataRGB() As RGBQUAD 'Mapped to Data
Dim TransDataRGB() As RGBQUAD 'mapped to TransData
Dim r As Long, g As Long, b As Long 'different uses
Dim TrC As Long 'Transparent color
x0 = CurSel.x1
y0 = CurSel.y1
SelW = CurSel.x2 - CurSel.x1 + 1
SelH = CurSel.y2 - CurSel.y1 + 1
x1 = Max(0, x1)
y1 = Max(0, y1)
x2 = Min(x2, SelW - 1)
y2 = Min(y2, SelH - 1)
ProcessW = x2 - x1 + 1
ProcessH = y2 - y1 + 1
If ProcessW <= 0 Or ProcessH <= 0 Then Exit Sub
If AryDims(AryPtr(SelOutPut)) <> 1 Then
    ReDim SelOutPut(0 To SelW * SelH - 1)
End If
On Error GoTo eh

ConstructAry AryPtr(SelOutputRGB), VarPtr(SelOutPut(0)), 4, SelW * SelH
ConstructAry AryPtr(SelDataRGB), VarPtr(CurSel_SelData(0, 0)), 4, SelW * SelH
ConstructAry AryPtr(DataRGB), VarPtr(Data(0, 0)), 4, intW, intH

If SelMode = dbUseCurSelMode Then
    SelMode = CurSel.SelMode
End If

If SelMode = dbSuperTransparent Then
    UpdateCurselTransData
    If AryDims(AryPtr(TransData)) = 2 Then
        ConstructAry AryPtr(TransDataRGB), VarPtr(TransData(0, 0)), 4, SelW * SelH
    Else
        SelMode = dbReplace
    End If
End If

Select Case SelMode
    Case dbSelMode.dbReplace
        CopyMemory SelOutPut(0), CurSel_SelData(0, 0), SelW * SelH * 4&
    Case dbSelMode.dbAdd
        For y = y1 To y2
            OfcYSel = y * SelW
            yd = y0 + y
            For x = x1 To x2
                xd = x0 + x
                r = CLng(DataRGB(xd, yd).rgbRed) + SelDataRGB(OfcYSel + x).rgbRed
                If r > 255& Then r = 255&
                SelOutputRGB(OfcYSel + x).rgbRed = r
                
                g = CLng(DataRGB(xd, yd).rgbGreen) + SelDataRGB(OfcYSel + x).rgbGreen
                If g > 255& Then g = 255&
                SelOutputRGB(OfcYSel + x).rgbGreen = g
                
                b = CLng(DataRGB(xd, yd).rgbBlue) + SelDataRGB(OfcYSel + x).rgbBlue
                If b > 255& Then b = 255&
                SelOutputRGB(OfcYSel + x).rgbBlue = b
            Next x
        Next y
    Case dbSelMode.dbMerge
        Dim AlphaLUT() As Byte
        BuildAlphaBlendingLUT dbCurSelTrR, AlphaLUT
        
        For y = y1 To y2
            OfcYSel = y * SelW
            yd = y0 + y
            For x = x1 To x2
                xd = x0 + x
                SelOutputRGB(OfcYSel + x).rgbRed = AlphaLUT(DataRGB(xd, yd).rgbRed, SelDataRGB(OfcYSel + x).rgbRed)
                SelOutputRGB(OfcYSel + x).rgbGreen = AlphaLUT(DataRGB(xd, yd).rgbGreen, SelDataRGB(OfcYSel + x).rgbGreen)
                SelOutputRGB(OfcYSel + x).rgbBlue = AlphaLUT(DataRGB(xd, yd).rgbBlue, SelDataRGB(OfcYSel + x).rgbBlue)
            Next x
        Next y

    Case dbSelMode.dbAND
        For y = y1 To y2
            OfcYSel = y * SelW
            yd = y0 + y
            For x = x1 To x2
                xd = x0 + x
                SelOutPut(OfcYSel + x) = Data(xd, yd) And CurSel_SelData(x, y)
            Next x
        Next y
    Case dbSelMode.dbOR
        For y = y1 To y2
            OfcYSel = y * SelW
            yd = y0 + y
            For x = x1 To x2
                xd = x0 + x
                SelOutPut(OfcYSel + x) = Data(xd, yd) Or CurSel_SelData(x, y)
            Next x
        Next y
    Case dbSelMode.dbXOR
        For y = y1 To y2
            OfcYSel = y * SelW
            yd = y0 + y
            For x = x1 To x2
                xd = x0 + x
                SelOutPut(OfcYSel + x) = (Data(xd, yd) Xor CurSel_SelData(x, y)) And &HFFFFFF
            Next x
        Next y
    Case dbSelMode.dbEQV
        For y = y1 To y2
            OfcYSel = y * SelW
            yd = y0 + y
            For x = x1 To x2
                xd = x0 + x
                SelOutPut(OfcYSel + x) = (Data(xd, yd) Eqv CurSel_SelData(x, y)) And &HFFFFFF
            Next x
        Next y
    Case dbSelMode.dbIMP
        For y = y1 To y2
            OfcYSel = y * SelW
            yd = y0 + y
            For x = x1 To x2
                xd = x0 + x
                SelOutPut(OfcYSel + x) = (CurSel_SelData(x, y) Imp Data(xd, yd)) And &HFFFFFF
            Next x
        Next y
    Case dbSelMode.dbNOT
        For y = y1 To y2
            OfcYSel = y * SelW
            yd = y0 + y
            For x = x1 To x2
                xd = x0 + x
                SelOutPut(OfcYSel + x) = &HFFFFFF And Not CurSel_SelData(x, y)
            Next x
        Next y
    Case dbSelMode.dbTransparent
        TrC = CurSel.TransColor
        For y = y1 To y2
            OfcYSel = y * SelW
            yd = y0 + y
            For x = x1 To x2
                xd = x0 + x
                If CurSel_SelData(x, y) = TrC Then
                    SelOutPut(OfcYSel + x) = Data(xd, yd)
                Else
                    SelOutPut(OfcYSel + x) = CurSel_SelData(x, y)
                End If
            Next x
        Next y
    Case dbSelMode.dbOverlayed
        TrC = CurSel.TransColor
        For y = y1 To y2
            OfcYSel = y * SelW
            yd = y0 + y
            For x = x1 To x2
                xd = x0 + x
                If Data(xd, yd) = TrC Then
                    SelOutPut(OfcYSel + x) = CurSel_SelData(x, y)
                Else
                    SelOutPut(OfcYSel + x) = Data(xd, yd)
                End If
            Next x
        Next y
    Case dbSelMode.dbSuperTransparent
        For y = y1 To y2
            OfcYSel = y * SelW
            yd = y0 + y
            For x = x1 To x2
                xd = x0 + x
                SelOutputRGB(OfcYSel + x).rgbRed = ((CLng(DataRGB(xd, yd).rgbRed) - SelDataRGB(OfcYSel + x).rgbRed) * TransDataRGB(OfcYSel + x).rgbRed) \ 255& + SelDataRGB(OfcYSel + x).rgbRed
                SelOutputRGB(OfcYSel + x).rgbGreen = ((CLng(DataRGB(xd, yd).rgbGreen) - SelDataRGB(OfcYSel + x).rgbGreen) * TransDataRGB(OfcYSel + x).rgbGreen) \ 255& + SelDataRGB(OfcYSel + x).rgbGreen
                SelOutputRGB(OfcYSel + x).rgbBlue = ((CLng(DataRGB(xd, yd).rgbBlue) - SelDataRGB(OfcYSel + x).rgbBlue) * TransDataRGB(OfcYSel + x).rgbBlue) \ 255& + SelDataRGB(OfcYSel + x).rgbBlue
            Next x
        Next y
    Case dbSelMode.dbMatrixMixed
        Dim tSelData() As Long
        ConstructAry AryPtr(tSelData), VarPtr(SelOutPut(0)), 4, SelW, SelH
        dbMatrix2 Data, CurSel_SelData, _
                  x0, y0, _
                  SelMatrix, _
                  tSelData, _
                  x1, y1, x2, y2
        UnReferAry AryPtr(tSelData), False
    Case Else
        Err.Raise 1111, "ProcessSelData", "Not implemented!"
End Select
UnReferAry AryPtr(SelOutputRGB), False
UnReferAry AryPtr(SelDataRGB), False
UnReferAry AryPtr(DataRGB), False
UnReferAry AryPtr(TransDataRGB), False
UnReferAry AryPtr(tSelData), False
Exit Sub
eh:
UnReferAry AryPtr(SelOutputRGB), False
UnReferAry AryPtr(SelDataRGB), False
UnReferAry AryPtr(DataRGB), False
UnReferAry AryPtr(TransDataRGB), False
UnReferAry AryPtr(tSelData), False
Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Sub UpdateCurselTransData()
Dim SelW As Long, SelH As Long
Dim b As Boolean
b = False
If Not CurSel.Selected Then Exit Sub
SelW = CurSel.x2 - CurSel.x1 + 1
SelH = CurSel.y2 - CurSel.y1 + 1
If SelW <= 0 Or SelH <= 0 Then Err.Raise 1111, "UpdateCurSelTransData", "Zero width/height of the selection!"
If AryDims(AryPtr(TransOrigData)) = 2 Then
    If AryDims(AryPtr(TransData)) <> 2 Then
        TransData = TransOrigData
        b = True
    End If
    If UBound(TransData, 1) + 1 <> SelW Or UBound(TransData, 2) + 1 <> SelH _
       Or TransDataChanged Then
        If Not b Then
            TransData = TransOrigData
        End If
        CancelDoEvents
        dbStretch TransData, SelW, SelH, SMSquares, RaiseErrors:=True
        RestoreDoEvents
    End If
    If UBound(TransData, 1) + 1 <> SelW Or UBound(TransData, 2) + 1 <> SelH Then
        Err.Raise 1111, "UpdateCurselTransData", "Failed to stretch the transparency channel. dbStretch failed!"
    End If
Else
    Erase TransData
End If
TransDataChanged = False
End Sub

'builds alpha-blending LUT.
'out=LUT(BG,FG)
'out is BG when AlphaDbl is 1 (fully transparent)
Public Sub BuildAlphaBlendingLUT(ByVal AlphaDbl As Double, _
                                 ByRef LUT() As Byte)
Dim c1 As Long, c2 As Long
Dim InvAlpha As Double
Static LastAlpha As Double
Static Initd As Boolean
Static LastLUT() As Byte
If AlphaDbl > 1# Then AlphaDbl = 1#
If AlphaDbl < 0# Then AlphaDbl = 0#
If Not Initd Or (Abs(AlphaDbl - LastAlpha) < 0.001) Then
    InvAlpha = 1# - AlphaDbl
    ReDim LastLUT(0 To 255, 0 To 255)
    For c1 = 0 To 255
        For c2 = 0 To 255
            LastLUT(c1, c2) = c1 * AlphaDbl + c2 * InvAlpha
        Next c2
    Next c1
    LastAlpha = AlphaDbl
    Initd = True
End If
LUT = LastLUT
End Sub

Public Sub UpdateSelPic(ByVal RedrawSelPic As Boolean, _
                        ByVal SetToData As Boolean, _
                        Optional ByVal SelMode As dbSelMode = dbSelMode.dbUseCurSelMode, _
                        Optional ByVal StoreUndo As Boolean = True) ', _
                        Optional ByVal ClipToView As Boolean = False)
Dim bmi As BITMAPINFO
Dim hBitmap As Long

Dim tSelData() As Long, ptrData As Long
Dim SelW As Long, SelH As Long
Dim rct As RECT
Dim x As Long, y As Long, ty As Long
Dim yOfc As Long
Dim x0 As Long, y0 As Long
Dim ClipToView As Boolean
If Not CurSel.Selected Then Exit Sub
If Not RedrawSelPic And Not SetToData Then Exit Sub

If AryDims(AryPtr(CurSel_SelData)) <> 2 Then Exit Sub

CurSel.x2 = CurSel.x1 + UBound(CurSel_SelData, 1)
CurSel.y2 = CurSel.y1 + UBound(CurSel_SelData, 2)

x0 = CurSel.x1
y0 = CurSel.y1

SelW = CurSel.x2 - CurSel.x1 + 1
SelH = CurSel.y2 - CurSel.y1 + 1
If SelW * SelH * Zm * Zm > 2000& * 2000& Then
    If SelPicture.AutoRedraw Then SelPicture.Cls
    SelPicture.AutoRedraw = False
    ClipToView = Not SetToData
End If
rct.Left = Max(0, -x0)
rct.Top = Max(0, -y0)
rct.Right = Min(SelW, intW - CurSel.x1)
rct.Bottom = Min(SelH, intH - CurSel.y1)
If ClipToView Then
    GetSelVisibilityRect rct
    DivideRect rct, Zm
End If

With bmi.bmiHeader
    .biSize = Len(bmi.bmiHeader)
    .biBitCount = 32
    .biHeight = -SelH
    .biWidth = SelW
    .biSizeImage = SelW * SelH * 4&
    .biPlanes = 1
End With
hBitmap = CreateDIBSection(SelPicture.hDC, bmi, DIB_RGB_COLORS, VarPtr(ptrData), 0, 0)
If hBitmap = 0 Or ptrData = 0 Then
    If hBitmap <> 0 Then DeleteObject hBitmap
    Err.Raise 1111, "UpdateSelPic", "Cannot create the bitmap!"
End If
ConstructAry AryPtr(tSelData), ptrData, 4, SelW * SelH

On Error GoTo eh

ProcessSelData SelMode, rct.Left, rct.Top, rct.Right - 1, rct.Bottom - 1, tSelData

If SetToData Then
    'set to data
    If StoreUndo Then
        StoreFragment x0 + rct.Left, y0 + rct.Top, _
                      x0 + rct.Right - 1, y0 + rct.Bottom - 1
    End If
    For y = rct.Top To rct.Bottom - 1
        yOfc = SelW * y
        ty = y0 + y
        For x = rct.Left To rct.Right - 1
            Data(x0 + x, ty) = tSelData(yOfc + x) And &HFFFFFF
        Next x
    Next y
    
    'draw it to the MP
    PaintBitmap hBitmap, MP.hDC, _
                rct.Left, rct.Top, rct.Right - rct.Left, rct.Bottom - rct.Top, _
                (x0 + rct.Left) * Zm, (y0 + rct.Top) * Zm, (rct.Right - rct.Left) * Zm, (rct.Bottom - rct.Top) * Zm
End If

If RedrawSelPic Then
    'draw it to the SelPicture
    If SelW * SelH * Zm * Zm <= 2000& * 2000& Then
        SelPicture.AutoRedraw = True
        If SelPicture.Width <> SelW * Zm Or SelPicture.Height <> SelH * Zm Then
            'clear only if needed - if resized
            SelPicture.Move SelPicture.Left, SelPicture.Top, SelW * Zm, SelH * Zm
            SelPicture.Cls
        End If
    End If
    'move before if not autoredraw (to avoid repainting)
    If Not SelPicture.AutoRedraw Then
        SelPicture.Move x0 * Zm, y0 * Zm, SelW * Zm, SelH * Zm
    End If
    PaintBitmap hBitmap, SelPicture.hDC, _
                rct.Left, rct.Top, rct.Right - rct.Left, rct.Bottom - rct.Top, _
                rct.Left * Zm, rct.Top * Zm, (rct.Right - rct.Left) * Zm, (rct.Bottom - rct.Top) * Zm
    dbFreshSelBorder
    'or if autoredraw, move after (for smoothness)
    If SelPicture.AutoRedraw Then
        SelPicture.Move x0 * Zm, y0 * Zm
        SelPicture.Refresh
    End If
    
End If

If ptrData <> 0 Then
    UnReferAry AryPtr(tSelData)
    DeleteObject hBitmap
End If
Exit Sub
Resume
eh:
If ptrData <> 0 Then
    UnReferAry AryPtr(tSelData)
    DeleteObject hBitmap
End If
Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Sub PaintBitmap(ByVal hBitmap As Long, _
                       ByVal DestHdc As Long, _
                       ByVal SrcX As Long, ByVal SrcY As Long, _
                       ByVal srcW As Long, ByVal srcH As Long, _
                       ByVal dstX As Long, ByVal dstY As Long, _
                       ByVal dstW As Long, ByVal dstH As Long)
Dim tmpDC As Long
Dim hDef As Long
Dim ret As Long
tmpDC = CreateCompatibleDC(DestHdc)
If tmpDC = 0 Then Err.Raise 1111, "PaintBitmap", "Cannot create compatible DC."
On Error GoTo eh
hDef = SelectObject(tmpDC, hBitmap)
If hDef = 0 Then Err.Raise 1111, "PaintBitmap", "SelectObject failed!"

ret = StretchBlt(DestHdc, _
                 dstX, dstY, dstW, dstH, _
                 tmpDC, _
                 SrcX, SrcY, srcW, srcH, _
                 VBRUN.RasterOpConstants.vbSrcCopy)
If ret = 0 Then Err.Raise 1111, "PaintBitmap", "StretchBlt Failed!"
SelectObject tmpDC, hDef
DeleteDC tmpDC
Exit Sub
eh:
If hDef <> 0 Then
    SelectObject tmpDC, hDef
End If
If tmpDC <> 0 Then
    DeleteDC tmpDC
End If


End Sub

'depends on GetVisRect
Friend Sub GetSelVisibilityRect(ByRef rct As RECT)
Dim CurSelRect As RECT
Dim DataRect As RECT
If CurSel.Selected Then
    GetVisRect rct, UseMPPos:=True
    
    CurSelRect.Left = CurSel.x1 * Zm
    CurSelRect.Top = CurSel.y1 * Zm
    CurSelRect.Right = (CurSel.x2 + 1) * Zm
    CurSelRect.Bottom = (CurSel.y2 + 1) * Zm
    rct = IntersectRects(rct, CurSelRect)
    
    
    DataRect.Left = 0
    DataRect.Top = 0
    DataRect.Right = (intW) * Zm
    DataRect.Bottom = (intH) * Zm
    rct = IntersectRects(rct, DataRect)
    
    
    rct.Left = rct.Left - CurSelRect.Left
    rct.Top = rct.Top - CurSelRect.Top
    rct.Right = rct.Right - CurSelRect.Left
    rct.Bottom = rct.Bottom - CurSelRect.Top
    
    
    
Else
    rct.Left = 0
    rct.Top = 0
    rct.Right = 0
    rct.Bottom = 0
End If
End Sub

Friend Sub DivideRect(ByRef rct As RECT, ByVal Zm As Integer)
rct.Left = Int(rct.Left / Zm)
rct.Top = Int(rct.Top / Zm)
rct.Right = -Int(-rct.Right / Zm)
rct.Bottom = -Int(-rct.Bottom / Zm)
End Sub

Friend Sub UpdateSelRect(ByVal Show As Boolean)
CheckCurSelCoords
With SelRect
    .Visible = False
    .Move CurSel.x1 * Zm, _
          CurSel.y1 * Zm, _
          Zm * (CurSel.x2 - CurSel.x1 + 1), _
          Zm * (CurSel.y2 - CurSel.y1 + 1)
    AnimateSelRect
    .Visible = Show
End With
End Sub

Friend Sub CheckCurSelCoords()
Dim x1 As Long, y1 As Long
Dim x2 As Long, y2 As Long
x1 = Min(CurSel.x1, CurSel.x2)
y1 = Min(CurSel.y1, CurSel.y2)
x2 = Max(CurSel.x1, CurSel.x2)
y2 = Max(CurSel.y1, CurSel.y2)

If x1 < 0 Then x1 = 0
If x1 > intW - 1 Then x1 = intW - 1

If x2 < 0 Then x2 = 0
If x2 > intW - 1 Then x2 = intW - 1

If y1 < 0 Then y1 = 0
If y1 > intH - 1 Then y1 = intH - 1

If y2 < 0 Then y2 = 0
If y2 > intH - 1 Then y2 = intH - 1

CurSel.x1 = x1
CurSel.y1 = y1
CurSel.x2 = x2
CurSel.y2 = y2
End Sub

'Gets the selection from data. Optionally sets the area to bgcolor.
'if bgcolor is missing, the current one is assumed.
'If cuts, is also draws.
Friend Sub GetSel(Optional ByVal Cut As Boolean = True, _
                  Optional ByVal BGColor As Long = -1)
Dim x As Long, y As Long
Dim x0 As Long, y0 As Long
Dim SelW As Long, SelH As Long
StoreFragment CurSel.x1, CurSel.y1, CurSel.x2, CurSel.y2
x0 = CurSel.x1
y0 = CurSel.y1
SelW = CurSel.x2 - CurSel.x1 + 1
SelH = CurSel.y2 - CurSel.y1 + 1
ReDim CurSel_SelData(0 To SelW - 1, 0 To SelH - 1)
If Cut Then
    If BGColor = -1 Then
        BGColor = ACol(2)
    End If
    For y = 0 To SelH - 1
        For x = 0 To SelW - 1
            CurSel_SelData(x, y) = Data(x + x0, y + y0)
            Data(x + x0, y + y0) = BGColor
        Next x
    Next y
    MP.Line (CurSel.x1 * Zm, CurSel.y1 * Zm)-(CurSel.x2 * Zm + Zm - 1, CurSel.y2 * Zm + Zm - 1), BGColor, BF
    GdiFlush
Else
    For y = 0 To SelH - 1
        For x = 0 To SelW - 1
            CurSel_SelData(x, y) = Data(x + x0, y + y0)
        Next x
    Next y
End If
End Sub

Public Sub DrawGrid()
Dim x As Long, y As Long
Dim rct As RECT
Dim Ofc As Long
If Zm <= 1 Then Exit Sub
Dim tColor As Long
If MP.AutoRedraw Then
    GdiFlush
    For y = 0 To intH * Zm - 1 Step Zm
        Ofc = y * MPBitsWidth
        For x = 0 To intW * Zm - 1 Step Zm
            MPBitsLNG(x + Ofc) = (MPBitsLNG(x + 1 + Ofc) + &H808080) And &HFFFFFF
        Next x
    Next y
    MP.Refresh
Else
    GetVisRect rct, True
    DivideRect rct, Zm
    If rct.Left < 0 Then rct.Left = 0
    If rct.Top < 0 Then rct.Top = 0
    If rct.Right > intW Then rct.Right = intW
    If rct.Bottom > intH Then rct.Bottom = intH
    For y = rct.Top To rct.Bottom - 1
        For x = rct.Left To rct.Right - 1
            SetPixel MP.hDC, x * Zm, y * Zm, (Data(x, y) + &H808080) And &HFFFFFF
            If x * Zm And &H100& And MainModule.VistaSetPixelBugDetected Then
              If Not MP.AutoRedraw Then MP.Line (x * Zm, y * Zm)-(x * Zm, y * Zm), (Data(x, y) + &H808080) And &HFFFFFF, BF
            End If
        Next x
    Next y
End If

End Sub

'Also sets mouse pointer to hourglass if the form disabled
Friend Sub DisableMe(Optional ByVal aEnabled As Boolean = False)
If IsEmpty(MeEnabledStack.DefValue) Then
    MeEnabledStack.DefValue = True
End If
MeEnabledStack.Push MeEnabled
MeEnabled = aEnabled
SetMousePtr Not MeEnabled
End Sub

Friend Sub RestoreMeEnabled()
If IsEmpty(MeEnabledStack.DefValue) Then
    MeEnabledStack.DefValue = True
End If
MeEnabled = MeEnabledStack.Pop
SetMousePtr Not MeEnabled
End Sub

Friend Sub ClearMeEnabledStack(Optional ByVal EnableMe As Boolean = True)
MeEnabledStack.Clear
If EnableMe Then
    MeEnabled = True
    SetMousePtr Not MeEnabled
End If
End Sub


'x2 and y2 are including
Friend Sub UpdateRegion(ByVal fx As Long, ByVal fy As Long, _
                        ByVal tx As Long, ByVal ty As Long)
If MP.AutoRedraw Then
    Dim x As Long, y As Long
    Dim x1 As Long, y1 As Long
    Dim ox As Long, oy As Long
    Dim OfcX As Long, OfcY As Long
    Dim tmp As Long
    Dim Rct1 As RECT
    'If Not MakeARMP Then Exit Sub
    If MPhDefBitmap = 0 Then
        Refr
        Exit Sub
    End If
    
    TestDims Data
    
    intW = UBound(Data, 1) + 1
    intH = UBound(Data, 2) + 1
    
    Rct1.Left = Max(fx, 0)
    Rct1.Top = Max(fy, 0)
    Rct1.Right = Min(tx + 1, intW)
    Rct1.Bottom = Min(ty + 1, intH)
    If Rct1.Right <= Rct1.Left Or Rct1.Bottom <= Rct1.Top Then
        'empty region
        Exit Sub
    End If
    
    If Zm = 1 Then
        If Rct1.Right - Rct1.Left = intW Then
            CopyMemory MPBitsLNG(Rct1.Top * intW), Data(0, Rct1.Top), (Rct1.Bottom - Rct1.Top) * intW * 4&
        Else
            For y = Rct1.Top To Rct1.Bottom - 1
                CopyMemory MPBitsLNG(y * intW + Rct1.Left), _
                           Data(Rct1.Left, y), _
                           4& * (Rct1.Right - Rct1.Left)
            Next y
        End If
    Else
        For y = Rct1.Top To Rct1.Bottom - 1
            For y1 = 0 To Zm - 1
                oy = y * Zm + y1
                OfcY = oy * MPBitsWidth
                For x = Rct1.Left To Rct1.Right - 1
                    OfcX = OfcY + x * Zm
                    tmp = Data(x, y)
                    For x1 = 0& To Zm - 1&
                        MPBitsLNG(x1 + OfcX) = tmp
                    Next x1
                Next x
            Next y1
        Next y
    End If
    MP.Refresh

Else
    If FreezeRefresh Then
        NeedRefr = True
        Exit Sub
    End If
    Dim DrawRect As RECT
    Dim rct As RECT
    Dim dh As Long
    
    GetVisRect DrawRect, True
    DivideRect DrawRect, Zm
    
    rct.Left = Max(fx, 0)
    rct.Top = Max(fy, 0)
    rct.Right = Min(tx + 1, intW)
    rct.Bottom = Min(ty + 1, intH)
    
    DrawRect = IntersectRects(DrawRect, rct)
    
    DrawData Data, Zm, MP.hDC, DrawRect
End If
End Sub


Public Sub Uninstall()
Dim sArr() As String
Dim n As Long
If dbMsgBox(2503, vbQuestion Or vbOKCancel) = vbCancel Then Exit Sub
ShowStatus 1250 '"Restoring file associations..."
'Load frmReg
UninstallFileTypes
'Unload frmReg
ShowStatus 1251 '"Deleting registry settings..."
With gReg
    .DeleteKey HKEY_CURRENT_USER, "Software\Dbnz\SMBMaker"
    .DeleteKey HKEY_LOCAL_MACHINE, "Software\Dbnz\SMBMaker\" + CStr(App.Revision)
End With
If Len(AppPath) > 0 Then
    dbMsgBox 2536, vbInformation
    ExploreFolder AppPath
Else
    dbMsgBox 2535, vbCritical
    'dbMsgBox "Cannot detect application's path. You have to find it manually. Search computer for ""SMBMaker.exe"" to delete it.", vbCritical
End If
dbEnd
End Sub

Sub NaviStart()
If NaviEnabled Then Exit Sub
GetCursorPos NaviMousePos
NaviEnabled = True
SetCapture MP.hWnd
End Sub

Sub NaviStop()
If Not (MPMS.ButtonState(1) Or MPMS.ButtonState(2) Or MPMS.ButtonState(4)) Then
  ReleaseCapture
End If
NaviEnabled = False
End Sub

Sub NaviMove()
Dim pos As POINTAPI
Dim Movement As POINTAPI
GetCursorPos pos
Movement.x = pos.x - NaviMousePos.x
Movement.y = pos.y - NaviMousePos.y
If Movement.x <> 0 Or Movement.y <> 0 Then
  If ScrollSettings.NaviAbsoluteMode Then
    NaviMousePos = pos
  Else
    SetCursorPos NaviMousePos.x, NaviMousePos.y
  End If
  If Movement.x ^ 2 + Movement.y ^ 2 < (Screen.Height / Screen.TwipsPerPixelY / 8) ^ 2 Then
    ScrollSettings.DontScroll = True
    ChangeScrollBarValue HScroll, Movement.x * 3
    ChangeScrollBarValue VScroll, Movement.y * 3
    ScrollSettings.DontScroll = False
    ApplyScrollBarsValues
  End If
End If
End Sub




Public Sub FlashStatusBar(Optional ByVal NumberOfFlashes = 4)
StatusFlashesLeft = NumberOfFlashes * 2
tmrStatusFlasher.Enabled = True
End Sub

Friend Sub GetLineOpts(ByRef LO As LineSettings)
LO = LineOpts
End Sub

Friend Sub SetLineOpts(ByRef LO As LineSettings)
LineOpts = LO
End Sub

Public Sub BeginTransform(ByVal NewW As Long, ByVal NewH As Long)
If NewW = -1 Then
  Erase TransformData
  Exit Sub
End If
If NewW = 0 Then
  UpdateWH
  NewW = intW
  NewH = intH
End If
If NewW <= 0 Or NewH <= 0 Then Err.Raise 12211, "BeginTransform", "Illegal Width or Height!"
ReDim TransformData(0 To NewW - 1, 0 To NewH - 1)
ClearPic TransformData, ACol(2)
End Sub

Public Sub EndTransform()
If AryDims(AryPtr(TransformData)) = 2 Then
  BUD AryPtr(Data)
  SwapArys AryPtr(TransformData), AryPtr(Data)
  UpdateWH
  Refr
  'StartPixelAction
  SetTexMode NewMode:=False
Else
  Err.Raise 121212, "EndTransform", "EndTransform without BeginTransform"
End If
End Sub

Friend Sub pTransformBlock(ByRef dstRect As FloatRect, _
                           ByRef srcGon As vtQGon, _
                           ByRef ProcessRect As RECT)
If AryDims(AryPtr(TransformData)) = 2 Then
  TransformBlock DstData:=TransformData, _
                 dstRect:=dstRect, _
                 srcData:=Data, _
                 srcGon:=srcGon, _
                 ProcessRect:=ProcessRect, _
                 Background:=IIf(TexMode, -1, ACol(2))
Else
  StoreFragment ProcessRect.Left, ProcessRect.Top, ProcessRect.Right - 1, ProcessRect.Bottom - 1
  TransformBlock DstData:=Data, _
                 dstRect:=dstRect, _
                 srcData:=Data, _
                 srcGon:=srcGon, _
                 ProcessRect:=ProcessRect, _
                 Background:=IIf(TexMode, -1, ACol(2))
  'StartPixelAction
  UpdateRegion ProcessRect.Left, ProcessRect.Top, ProcessRect.Right - 1, ProcessRect.Bottom - 1
End If
End Sub

'used from clsEVal
Friend Function TransformDataPresent() As Boolean
TransformDataPresent = AryDims(AryPtr(TransformData)) = 2
End Function

    
    
    'MsgText = "Bad Boolean Value (Dynamic scrolling)"
    'ScrollSettings.DS_EnableIfAR = &H1 And dbGetSettingEx("View", "DynamicScrolling", , True)
    'ScrollSettings.DS_EnableIfAR = ScrollSettings.DS_EnableIfAR Or (&H2 And dbGetSettingEx("View", "DynamicScrollingNoAR", , False))
    'ChangeDynScrolling CBool(ScrollSettings.DS_EnableIfAR And &H1)
    'ScrollSettings.DS_EnL = dbGetSettingEx("View", "ScrollEnL", vbSingle, ScrollSettings.DS_EnL)
    'ScrollSettings.DS_Jestkost = dbGetSettingEx("View", "ScrollHumidity", vbSingle, ScrollSettings.DS_Jestkost)
    'MoveTimerRes = dbGetSettingEx("View", "ScrollTimerRes", vbInteger, 8)
    'If MoveTimerRes <= 0 Then MoveTimerRes = 8
    'MoveMP Flags:=&H8&
Private Sub LoadScrollSettings()
On Error GoTo eh
With ScrollSettings
  .DS_EnL = dbGetSettingEx("ScrollSettings\Dynamic Scrolling", "EnL", vbSingle, DefValue:=0.5)
  .DS_Jestkost = dbGetSettingEx("ScrollSettings\Dynamic Scrolling", "Rig", vbSingle, DefValue:=0.5)
  ChangeDynScrolling dbGetSettingEx("ScrollSettings\Dynamic Scrolling", "Enabled", vbBoolean, DefValue:=True)
  MoveTimerRes = dbGetSettingEx("ScrollSettings\Dynamic Scrolling", "TimerRes", vbInteger, DefValue:=8)
  If MoveTimerRes <= 0 Then MoveTimerRes = 8
  MoveMP Flags:=&H8&
  .MouseGlued = dbGetSettingEx("ScrollSettings", "MouseGlued", DefValue:=False)
  mnuMouseAttr.Checked = ScrollSettings.MouseGlued
  LoadASS "ScrollSettings\ASS_normal", .ASS, DefGap:=100
  LoadASS "ScrollSettings\ASS_pen", .ASS_pen, DefGap:=20
  .NaviAbsoluteMode = dbGetSettingEx("ScrollSettings", "NaviModeAbs", vbBoolean, DefValue:=False)
End With
Exit Sub
eh:
Select Case MsgError("Scroll settings loading error.", vbAbortRetryIgnore Or vbCritical)
  Case vbAbort
    ErrRaise
  Case vbRetry
    Resume
  Case vbIgnore
    Resume Next
End Select
End Sub

Private Sub LoadASS(ByRef Section As String, ByRef ASS As typAutoscrollSgs, ByVal DefGap As Long)
With ASS
  .GapLef = dbGetSettingEx(Section, "GapLef", vbLong, DefValue:=DefGap)
  .GapTop = dbGetSettingEx(Section, "GapTop", vbLong, DefValue:=DefGap)
  .GapRig = dbGetSettingEx(Section, "GapRig", vbLong, DefValue:=DefGap)
  .GapBot = dbGetSettingEx(Section, "GapBot", vbLong, DefValue:=DefGap)
End With
End Sub

Private Sub SaveScrollSettings()
With ScrollSettings
  dbSaveSettingEx "ScrollSettings\Dynamic Scrolling", "EnL", .DS_EnL
  dbSaveSettingEx "ScrollSettings\Dynamic Scrolling", "Rig", .DS_Jestkost
  dbSaveSettingEx "ScrollSettings\Dynamic Scrolling", "Enabled", .DS_Enabled
  dbSaveSettingEx "ScrollSettings\Dynamic Scrolling", "TimerRes", MoveTimerRes
  dbSaveSettingEx "ScrollSettings", "MouseGlued", .MouseGlued
  SaveASS "ScrollSettings\ASS_normal", .ASS
  SaveASS "ScrollSettings\ASS_pen", .ASS_pen
  dbSaveSettingEx "ScrollSettings", "NaviModeAbs", .NaviAbsoluteMode
  
End With

End Sub

Private Sub SaveASS(ByRef Section As String, ByRef ASS As typAutoscrollSgs)
With ASS
  dbSaveSettingEx Section, "GapLef", .GapLef
  dbSaveSettingEx Section, "GapTop", .GapTop
  dbSaveSettingEx Section, "GapRig", .GapRig
  dbSaveSettingEx Section, "GapBot", .GapBot
End With
End Sub

