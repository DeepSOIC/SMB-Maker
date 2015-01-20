VERSION 5.00
Begin VB.Form frmIconManager 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Icon manager"
   ClientHeight    =   3190
   ClientLeft      =   2761
   ClientTop       =   3751
   ClientWidth     =   6028
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.07
      Charset         =   204
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   290
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   548
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      Height          =   1050
      Left            =   2085
      ScaleHeight     =   1001
      ScaleWidth      =   1122
      TabIndex        =   2
      Top             =   705
      Width           =   1170
   End
   Begin VB.ListBox List1 
      Height          =   2101
      Left            =   180
      TabIndex        =   0
      Top             =   615
      Width           =   1530
   End
   Begin VB.Label Label2 
      Caption         =   "Image view"
      Height          =   225
      Left            =   2100
      TabIndex        =   3
      Top             =   465
      Width           =   1170
   End
   Begin VB.Label Label1 
      Caption         =   "Images in the icon:"
      Height          =   240
      Left            =   120
      TabIndex        =   1
      Top             =   255
      Width           =   1665
   End
End
Attribute VB_Name = "frmIconManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim pIcon As vtIcon
Dim pDIcon As vtIconForDrawing
Dim NewFlags() As Boolean 'holds the information about the new image position

Public Sub SetIcon(ByVal ptrIcon As Long, ByVal NewImageIndex As Long)
Dim prvIcon As vtIcon
CopyMemory prvIcon, ByVal ptrIcon, Len(pIcon)
pIcon = prvIcon
ZeroMemory prvIcon, Len(prvIcon)
End Sub

Public Sub FreshList()

End Sub

