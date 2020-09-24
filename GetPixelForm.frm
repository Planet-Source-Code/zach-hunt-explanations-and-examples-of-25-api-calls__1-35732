VERSION 5.00
Begin VB.Form GetPixelForm 
   Caption         =   "GetPixel"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5475
   LinkTopic       =   "Form1"
   ScaleHeight     =   213
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   365
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   375
      MultiLine       =   -1  'True
      TabIndex        =   3
      Text            =   "GetPixelForm.frx":0000
      Top             =   75
      Width           =   3750
   End
   Begin VB.PictureBox Picture2 
      Height          =   1200
      Left            =   2250
      ScaleHeight     =   1140
      ScaleWidth      =   1140
      TabIndex        =   2
      Top             =   870
      Width           =   1200
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   975
      Left            =   675
      Picture         =   "GetPixelForm.frx":0052
      ScaleHeight     =   61
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   61
      TabIndex        =   1
      Top             =   975
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Back"
      Height          =   600
      Left            =   1275
      TabIndex        =   0
      Top             =   2475
      Width           =   1800
   End
   Begin VB.Label bl 
      AutoSize        =   -1  'True
      Caption         =   "B="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   4050
      TabIndex        =   6
      Top             =   1800
      Width           =   285
   End
   Begin VB.Label gl 
      AutoSize        =   -1  'True
      Caption         =   "G="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   4050
      TabIndex        =   5
      Top             =   1350
      Width           =   300
   End
   Begin VB.Label rl 
      AutoSize        =   -1  'True
      Caption         =   "R="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   4050
      TabIndex        =   4
      Top             =   900
      Width           =   300
   End
End
Attribute VB_Name = "GetPixelForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'integers used to store the red, green,
'and bluecolor values of a color
Dim r As Integer
Dim g As Integer
Dim b As Integer

'GetPixel declaration
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long

Private Sub Command1_Click()
    Unload Me 'Unload the form
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    'store the color of the pixel that is clicked
    Dim color As Long
    
    'get the color of the selected pixel and
    'set it to picture2's background color
    color = GetPixel(Picture1.hdc, x, y)
    Picture2.BackColor = color
    r = color Mod 256
    b = Int(color / 65536)
    g = (color - (b * 65536) - r) / 256
    
    'this is another set of formulas that also works great
    'r = color And &H10000FF
    'g = (color And &H100FF00) / (2 ^ 8)
    'b = (color And &H1FF0000) / (2 ^ 16)
    'set the r, g, and b values to the labels
    
    rl.Caption = "R = " & r
    bl.Caption = "G = " & g
    gl.Caption = "B = " & b
End Sub
