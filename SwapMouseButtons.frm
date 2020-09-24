VERSION 5.00
Begin VB.Form SwapMouseButtons 
   Caption         =   "SwapMouseButtons"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2775
   LinkTopic       =   "Form1"
   ScaleHeight     =   213
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   185
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Back"
      Height          =   600
      Left            =   675
      TabIndex        =   1
      Top             =   2100
      Width           =   1500
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Swap Mouse Buttons"
      Height          =   750
      Left            =   525
      TabIndex        =   0
      Top             =   750
      Width           =   1800
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Left"
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
      Left            =   1125
      TabIndex        =   2
      Top             =   300
      Width           =   390
   End
End
Attribute VB_Name = "SwapMouseButtons"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'SwapMouseButton declaration
Private Declare Function SwapMouseButton Lib "user32" (ByVal bSwap As Long) As Long


Private Sub Command1_Click()
'Switch the left and right mouse button clicks
    If Label1.Caption = "Left" Then
    'the right button is the main button
        SwapMouseButton True
        Label1.Caption = "Right"
    ElseIf Label1.Caption = "Right" Then
    'the left button is the main button
        SwapMouseButton False
        Label1.Caption = "Left"
    End If
End Sub

Private Sub Command2_Click()
    Unload Me 'Unload the form
End Sub



