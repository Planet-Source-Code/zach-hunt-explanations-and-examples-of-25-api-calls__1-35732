VERSION 5.00
Begin VB.Form GetCursorPos 
   Caption         =   "GetCursorPos"
   ClientHeight    =   2400
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3885
   LinkTopic       =   "Form1"
   ScaleHeight     =   160
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   259
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox y 
      Height          =   300
      Left            =   1800
      TabIndex        =   4
      Top             =   870
      Width           =   1500
   End
   Begin VB.TextBox x 
      Height          =   300
      Left            =   1800
      TabIndex        =   3
      Top             =   240
      Width           =   1500
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Back"
      Height          =   600
      Left            =   1200
      TabIndex        =   0
      Top             =   1500
      Width           =   1575
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   3240
      Top             =   1440
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
      Left            =   450
      TabIndex        =   2
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
      Left            =   450
      TabIndex        =   1
      Top             =   300
      Width           =   1050
   End
End
Attribute VB_Name = "GetCursorPos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Declaration of POINTAPI type
Private Type POINTAPI
        x As Long
        y As Long
End Type

'Variable to store x and y coordinates
Dim MouseCoordinates As POINTAPI

'GetCursorPos Declaration
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

Private Sub Command1_Click()
    Unload Me 'Unload the form
End Sub

Private Sub Timer1_Timer()
    'Get the mouse coordinates
    GetCursorPos MouseCoordinates
    'put the coordinates into the text boxes
    x.Text = MouseCoordinates.x
    y.Text = MouseCoordinates.y
End Sub
