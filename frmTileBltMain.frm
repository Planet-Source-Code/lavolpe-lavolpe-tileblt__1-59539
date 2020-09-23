VERSION 5.00
Begin VB.Form frmTileBltMain 
   Caption         =   "Tile Blting - Select Image, Options, & Click ""Show Me"""
   ClientHeight    =   6975
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7845
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6975
   ScaleWidth      =   7845
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkStagger 
      Caption         =   "Stagger alternating rows"
      Height          =   285
      Left            =   5250
      TabIndex        =   6
      Top             =   5580
      Width           =   2565
   End
   Begin VB.ComboBox cboBorder 
      Height          =   315
      Left            =   5235
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   5925
      Width           =   2505
   End
   Begin VB.CheckBox chkLayered 
      Caption         =   "Layered (maintain background)"
      Height          =   285
      Left            =   5250
      TabIndex        =   4
      Top             =   5205
      Width           =   2565
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Show Me"
      Height          =   495
      Left            =   5205
      TabIndex        =   1
      Top             =   6330
      Width           =   2565
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00E0E0E0&
      Height          =   6390
      Left            =   210
      ScaleHeight     =   6330
      ScaleWidth      =   4830
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   420
      Width           =   4890
   End
   Begin VB.Label Label1 
      Caption         =   "Transparent GIF ^^"
      Height          =   255
      Index           =   2
      Left            =   5610
      TabIndex        =   7
      Top             =   4920
      Width           =   2025
   End
   Begin VB.Image Image1 
      Height          =   705
      Index           =   5
      Left            =   5280
      Picture         =   "frmTileBltMain.frx":0000
      Top             =   4080
      Width           =   2430
   End
   Begin VB.Shape shpSelect 
      Height          =   585
      Left            =   6150
      Top             =   1200
      Width           =   750
   End
   Begin VB.Label Label1 
      Caption         =   "Icons Cursors"
      Height          =   405
      Index           =   1
      Left            =   6840
      TabIndex        =   3
      Top             =   3540
      Width           =   615
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   4
      Left            =   6150
      Picture         =   "frmTileBltMain.frx":0496
      Top             =   3525
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   3
      Left            =   5490
      Top             =   3525
      Width           =   480
   End
   Begin VB.Label Label1 
      Caption         =   "Bitmaps can be Jpgs && Gifs too"
      Height          =   810
      Index           =   0
      Left            =   6795
      TabIndex        =   2
      Top             =   2625
      Width           =   615
   End
   Begin VB.Image Image1 
      Height          =   720
      Index           =   1
      Left            =   5850
      Picture         =   "frmTileBltMain.frx":07A0
      Top             =   2610
      Width           =   720
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   2
      Left            =   5280
      Picture         =   "frmTileBltMain.frx":0CA4
      Top             =   2625
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   1920
      Index           =   0
      Left            =   5295
      Picture         =   "frmTileBltMain.frx":0EF2
      Top             =   420
      Width           =   1920
   End
End
Attribute VB_Name = "frmTileBltMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
UpdateBltSample
Picture1.Refresh
End Sub

Private Sub Form_Load()
Set Image1(3) = Me.Icon
cboBorder.AddItem "No Offset"
cboBorder.AddItem "8 pixel Border"
cboBorder.AddItem "16 pixel Border"
cboBorder.AddItem "32 pixel Border"
Picture1.AutoRedraw = True
cboBorder.ListIndex = 0
Call Image1_Click(5)
chkStagger = 1
Call Command1_Click
End Sub

Private Sub UpdateBltSample()

Dim PicRect As RECT
Dim xyOffset As Long

' to show that TileBlt only does the area it is told to
xyOffset = Choose(cboBorder.ListIndex + 1, 0, 8, 16, 32)

' like most functions that use APIs, measurements are in pixels
PicRect.Left = xyOffset
PicRect.Top = xyOffset
PicRect.Right = Picture1.ScaleWidth / Screen.TwipsPerPixelX - xyOffset
PicRect.Bottom = Picture1.ScaleHeight / Screen.TwipsPerPixelY - xyOffset

If chkLayered.Value = 0 Or chkLayered.Enabled = False Then Picture1.Cls

' the function has many options, including using handles vs picture objects
' See that function's remarks for more info
TileBltRectEx Picture1.hdc, PicRect, Image1(Val(shpSelect.Tag)).Picture, _
    (chkStagger.Value = 1 And chkStagger.Enabled = True), _
    (chkLayered.Value = 1 And chkLayered.Enabled = True)
End Sub

Private Sub Image1_Click(Index As Integer)

' move the selector over the clicked image
With shpSelect
    .Move Image1(Index).Left - 60, Image1(Index).Top - 60, Image1(Index).Width + 120, Image1(Index).Height + 120
End With
shpSelect.Tag = Index
chkLayered.Enabled = (Image1(Index).Picture.Type = vbPicTypeIcon Or Index = 5)

End Sub
