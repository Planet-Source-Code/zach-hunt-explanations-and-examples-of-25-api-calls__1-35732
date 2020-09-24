VERSION 5.00
Begin VB.Form GetWindowRect_ClipCursor 
   Caption         =   "GetWindowRect_ClipCursor"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   213
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "Release mouse movement restriction"
      Height          =   900
      Left            =   2550
      TabIndex        =   2
      Top             =   750
      Width           =   1500
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Restrict Mouse Movement to within this form"
      Height          =   900
      Left            =   375
      TabIndex        =   1
      Top             =   750
      Width           =   1500
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Back"
      Height          =   600
      Left            =   1500
      TabIndex        =   0
      Top             =   2250
      Width           =   1500
   End
End
Attribute VB_Name = "GetWindowRect_ClipCursor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'a variable type that stores the top,left,
'right and bottom coordinates of a rectangle
Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

'a RECT type variable that will
'hold the dimensions of an object
Dim mouse As RECT

'GetWindowRect declaration
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long

'ClipCursor declaration
Private Declare Function ClipCursor Lib "user32" (lpRect As Any) As Long


Private Sub Command1_Click()
    Unload Me 'Unload the form
End Sub

Private Sub Command2_Click()
    GetWindowRect GetWindowRect_ClipCursor.hwnd, mouse
    ClipCursor mouse
End Sub

Private Sub Command3_Click()
    'Don't forget to release the mouse from its rectangle
    'otherwise it will stay that way throughout windows
    mouse.Left = 0
    mouse.Top = 0
    mouse.Right = Screen.Width
    mouse.Bottom = Screen.Height
    ClipCursor mouse
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Don't forget to release the mouse from its rectangle
    'otherwise it will stay that way throughout windows
    mouse.Left = 0
    mouse.Top = 0
    mouse.Right = Screen.Width
    mouse.Bottom = Screen.Height
    ClipCursor mouse
End Sub
