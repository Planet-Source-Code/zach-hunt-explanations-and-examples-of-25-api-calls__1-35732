VERSION 5.00
Begin VB.Form GetKeyState 
   Caption         =   "GetKeyState"
   ClientHeight    =   2400
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3375
   LinkTopic       =   "Form1"
   ScaleHeight     =   160
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   225
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   120
      Top             =   720
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Back"
      Height          =   600
      Left            =   750
      TabIndex        =   0
      Top             =   1500
      Width           =   1800
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Press the Caps Lock key on and off"
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
      TabIndex        =   2
      Top             =   75
      Width           =   3060
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "No Caption"
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
      Left            =   1200
      TabIndex        =   1
      Top             =   750
      Width           =   960
   End
End
Attribute VB_Name = "GetKeyState"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'variable that keeps track of if caps lock is on or off
Dim capson

'GetKeyState declaration
Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
'You can also use any other vbkeycode constants

Private Sub Command1_Click()
    Unload Me 'Unload the form
End Sub


Private Sub Timer1_Timer()
    'check to see if capslock is on or off
    capson = GetKeyState(vbKeyCapital)
    capson = (capson = 1 Or capson = -127)
    'set the label to the proper caption
    If capson = True Then
        Label1.Caption = "Caps Lock is on"
    ElseIf capson = False Then
        Label1.Caption = "Caps Lock is off"
    End If
End Sub
