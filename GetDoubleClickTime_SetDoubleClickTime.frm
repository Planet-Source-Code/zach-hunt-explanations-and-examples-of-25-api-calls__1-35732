VERSION 5.00
Begin VB.Form GetDoubleClickTime_SetDoubleClickTime 
   Caption         =   "GetDoubleClickTime_SetDoubleClickTime"
   ClientHeight    =   3390
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4845
   LinkTopic       =   "Form1"
   ScaleHeight     =   226
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   323
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text2 
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
      Height          =   450
      Left            =   150
      MultiLine       =   -1  'True
      TabIndex        =   4
      Text            =   "GetDoubleClickTime_SetDoubleClickTime.frx":0000
      Top             =   225
      Width           =   2550
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   825
      TabIndex        =   3
      Top             =   750
      Width           =   1200
   End
   Begin VB.CommandButton Command3 
      Caption         =   "GetDoubleClickTime"
      Height          =   750
      Left            =   2850
      TabIndex        =   2
      Top             =   1200
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "SetDoubleClickTime "
      Height          =   750
      Left            =   525
      TabIndex        =   1
      Top             =   1200
      Width           =   1800
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Back"
      Height          =   600
      Left            =   1800
      TabIndex        =   0
      Top             =   2400
      Width           =   1455
   End
End
Attribute VB_Name = "GetDoubleClickTime_SetDoubleClickTime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'GetDoubleClickTime declaration
Private Declare Function GetDoubleClickTime Lib "user32" () As Long

'SetDoubleClickTime declaration
Private Declare Function SetDoubleClickTime Lib "user32" (ByVal wCount As Long) As Long


Private Sub Command1_Click()
    Unload Me 'Unload the form
End Sub


Private Sub Command2_Click()
    'set the double click time with the value in the textbox
    SetDoubleClickTime Val(Text1.Text)
End Sub

Private Sub Command3_Click()
    'get the double click time and display it in a message box
    MsgBox "Time in milliseconds allowed between mouse clicks" & vbNewLine & "          Current Setting:  " & GetDoubleClickTime
End Sub

