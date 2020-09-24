VERSION 5.00
Begin VB.Form Sleep 
   Caption         =   "Sleep"
   ClientHeight    =   3300
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2925
   LinkTopic       =   "Form1"
   ScaleHeight     =   220
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   195
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Sleep"
      Height          =   750
      Left            =   525
      TabIndex        =   3
      Top             =   1200
      Width           =   1800
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1500
      TabIndex        =   1
      Top             =   525
      Width           =   1200
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Back"
      Height          =   600
      Left            =   675
      TabIndex        =   0
      Top             =   2400
      Width           =   1500
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "1000 Milliseconds = 1 Second"
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
      TabIndex        =   4
      Top             =   150
      Width           =   2565
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Milliseconds"
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
      Left            =   225
      TabIndex        =   2
      Top             =   555
      Width           =   1050
   End
End
Attribute VB_Name = "Sleep"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Sleep declaration
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)


Private Sub Command1_Click()
    Unload Me 'Unload the form
End Sub

Private Sub Command2_Click()
    'put the sleep call before what you want paused
    Sleep Val(Text1.Text)
    'Msgbox Done after the sleep is done
    MsgBox "Done", , "Sleep"
End Sub

