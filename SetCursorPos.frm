VERSION 5.00
Begin VB.Form SetCursorPos 
   Caption         =   "SetCursorPos"
   ClientHeight    =   2595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3885
   LinkTopic       =   "Form2"
   ScaleHeight     =   173
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   259
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox y 
      Height          =   300
      Left            =   1800
      TabIndex        =   5
      Top             =   900
      Width           =   1200
   End
   Begin VB.TextBox x 
      Height          =   300
      Left            =   1800
      TabIndex        =   4
      Top             =   300
      Width           =   1200
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Back"
      Height          =   375
      Left            =   1350
      TabIndex        =   1
      Top             =   2040
      Width           =   1200
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Set Cursor Position"
      Height          =   375
      Left            =   1050
      TabIndex        =   0
      Top             =   1500
      Width           =   1800
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Y Position"
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
      Left            =   525
      TabIndex        =   3
      Top             =   900
      Width           =   1065
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "X Position"
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
      Left            =   525
      TabIndex        =   2
      Top             =   300
      Width           =   1050
   End
End
Attribute VB_Name = "SetCursorPos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'SetCursorPos Declaration
Private Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long

Private Sub Command1_Click()
    'Call the setcursorpos API
    SetCursorPos Val(x.Text), Val(y.Text)
End Sub

Private Sub Command2_Click()
    Unload Me 'Unload the form
End Sub

