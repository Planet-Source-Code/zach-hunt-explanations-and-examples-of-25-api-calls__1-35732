VERSION 5.00
Begin VB.Form ShowCursor 
   Caption         =   "ShowCursor"
   ClientHeight    =   2100
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3885
   LinkTopic       =   "Form1"
   ScaleHeight     =   140
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   259
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   3120
      Top             =   1440
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Back"
      Height          =   525
      Left            =   1200
      TabIndex        =   2
      Top             =   1350
      Width           =   1500
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Hide Cursor"
      Height          =   525
      Left            =   2100
      TabIndex        =   1
      Top             =   525
      Width           =   1500
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Show Cursor"
      Height          =   525
      Left            =   300
      TabIndex        =   0
      Top             =   525
      Width           =   1500
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Press S if you can't find the mouse cursor"
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
      Left            =   150
      TabIndex        =   3
      Top             =   120
      Width           =   3555
   End
End
Attribute VB_Name = "ShowCursor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'ShowCursor Declaration
Private Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long

'Used if the user can't find the mouse cursor
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

Private Sub Command1_Click()
    ShowCursor True 'show the cursor
End Sub

Private Sub Command2_Click()
    ShowCursor False 'hide the cursor
End Sub

Private Sub Command3_Click()
    Unload Me 'Unload the form
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Make sure the mouse cursor is visible before the form unloads
    ShowCursor True
End Sub

Private Sub Timer1_Timer()
    'Show where the mouse cursor is
    If GetAsyncKeyState(vbKeyS) Then ShowCursor True
End Sub
