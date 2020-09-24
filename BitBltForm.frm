VERSION 5.00
Begin VB.Form BitBltForm 
   Caption         =   "BitBlt"
   ClientHeight    =   3600
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   975
      Left            =   750
      Picture         =   "BitBltForm.frx":0000
      ScaleHeight     =   915
      ScaleWidth      =   915
      TabIndex        =   2
      Top             =   300
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Back"
      Height          =   600
      Left            =   1680
      TabIndex        =   1
      Top             =   2700
      Width           =   1350
   End
   Begin VB.CommandButton Command1 
      Caption         =   "BitBlt"
      Height          =   600
      Left            =   1425
      TabIndex        =   0
      Top             =   1800
      Width           =   1800
   End
End
Attribute VB_Name = "BitBltForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'const needed for BitBlt or you can use the vb const vbSrcCopy
Private Const SRCCOPY = &HCC0020
'look for RasterOpConstants in the object browser for other BitBlt options
'or you can use these which are the API viewer constants
Private Const SRCAND = &H8800C6 'combines the destination and the source image
Private Const SRCERASE = &H440328 'Inverts the destination bitmap; combines the result with the source bitmap
Private Const SRCINVERT = &H660046 'Combines the inverted source bitmap/pattern with a destination bitmap
Private Const SRCPAINT = &HEE0086 'Combines the destination bitmap with the pattern


'BitBlt declaration
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Private Sub Command1_Click()
    'paint the contents of picture1 onto the form
    BitBlt BitBltForm.hdc, 180, 22, Picture1.Width - 3, Picture1.Height - 3, _
        Picture1.hdc, 0, 0, SRCCOPY
End Sub

Private Sub Command2_Click()
    Unload Me 'unload the form
End Sub

