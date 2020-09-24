VERSION 5.00
Begin VB.Form SetPixelForm 
   Caption         =   "SetPixel"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5475
   LinkTopic       =   "Form1"
   ScaleHeight     =   213
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   365
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Back"
      Height          =   600
      Left            =   2025
      TabIndex        =   0
      Top             =   2400
      Width           =   1350
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Click anywhere on the form to set a pixel to a random color"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   180
      TabIndex        =   1
      Top             =   75
      Width           =   5025
   End
End
Attribute VB_Name = "SetPixelForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'SetPixel Declaration
Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long


Private Sub Command1_Click()
    Unload Me 'Unload the form
End Sub

Private Sub Form_Load()
    'so the colors are different every time
    Randomize
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    'set the pixel where the mouse clicked to a random color
    SetPixel SetPixelForm.hdc, x, y, RGB(Rnd * 256, Rnd * 256, Rnd * 256)
End Sub
