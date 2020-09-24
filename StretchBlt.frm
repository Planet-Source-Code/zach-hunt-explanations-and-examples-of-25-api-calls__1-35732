VERSION 5.00
Begin VB.Form StretchBlt 
   Caption         =   "StretchBlt"
   ClientHeight    =   4320
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   288
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "Shrink"
      Height          =   750
      Left            =   2550
      TabIndex        =   4
      Top             =   2400
      Width           =   1500
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Stretch"
      Height          =   750
      Left            =   525
      TabIndex        =   3
      Top             =   2400
      Width           =   1500
   End
   Begin VB.PictureBox Picture2 
      Height          =   1800
      Left            =   2175
      ScaleHeight     =   1740
      ScaleWidth      =   1740
      TabIndex        =   2
      Top             =   225
      Width           =   1800
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   975
      Left            =   600
      Picture         =   "StretchBlt.frx":0000
      ScaleHeight     =   915
      ScaleWidth      =   915
      TabIndex        =   1
      Top             =   600
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Back"
      Height          =   600
      Left            =   1575
      TabIndex        =   0
      Top             =   3450
      Width           =   1500
   End
End
Attribute VB_Name = "StretchBlt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'const needed for StretchBlt or you can use the vb const vbSrcCopy
Private Const SRCCOPY = &HCC0020
'look for RasterOpConstants in the object browser for other StretchBlt options
'or you can use these which are the API viewer constants
Private Const SRCAND = &H8800C6 'combines the destination and the source image
Private Const SRCERASE = &H440328 'Inverts the destination bitmap; combines the result with the source bitmap
Private Const SRCINVERT = &H660046 'Combines the inverted source bitmap/pattern with a destination bitmap
Private Const SRCPAINT = &HEE0086 'Combines the destination bitmap with the pattern
'StretchBlt declaration
Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long

Private Sub Command1_Click()
Unload Me 'Unload the form
End Sub

Private Sub Command2_Click()
Picture2.Cls 'clear the picture

'Stretch picture1's picture and
'transfer it to picture2's picture
StretchBlt Picture2.hdc, 0, 0, Picture2.Width, _
    Picture2.Height, Picture1.hdc, 0, 0, _
    Picture1.Width - 3, Picture1.Height - 3, SRCCOPY
    
End Sub

Private Sub Command3_Click()
Picture2.Cls 'clear the picture

'Shrink picture1's picture and
'transfer it to picture2's picture
StretchBlt Picture2.hdc, 0, 0, Picture1.Width / 2, _
    Picture1.Height / 2, Picture1.hdc, 0, 0, _
    Picture1.Width - 3, Picture1.Height - 3, SRCCOPY
End Sub

