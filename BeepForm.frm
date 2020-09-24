VERSION 5.00
Begin VB.Form BeepForm 
   Caption         =   "Beep"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3135
   LinkTopic       =   "Form1"
   ScaleHeight     =   213
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   209
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Beep"
      Height          =   600
      Left            =   675
      TabIndex        =   5
      Top             =   1575
      Width           =   1800
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1320
      TabIndex        =   2
      Top             =   975
      Width           =   1500
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1320
      TabIndex        =   1
      Top             =   300
      Width           =   1500
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Back"
      Height          =   600
      Left            =   900
      TabIndex        =   0
      Top             =   2400
      Width           =   1350
   End
   Begin VB.Label Label2 
      Caption         =   "Duration   (1-5000)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   225
      TabIndex        =   4
      Top             =   930
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Frequency (1-10000)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   225
      TabIndex        =   3
      Top             =   255
      Width           =   1095
   End
End
Attribute VB_Name = "BeepForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Beep declaration
Private Declare Function Beep Lib "kernel32" (ByVal dwFreq As Long, ByVal dwDuration As Long) As Long

Private Sub Command1_Click()
    Unload Me 'Unload the form
End Sub

Private Sub Command2_Click()
    'Beep at the specified frequency and
    'for the specified duration of time
    Beep Val(Text1.Text), Val(Text2.Text)
End Sub

