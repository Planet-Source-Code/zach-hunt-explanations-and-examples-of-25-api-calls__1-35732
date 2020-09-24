VERSION 5.00
Begin VB.Form MoveToEx_LineTo 
   Caption         =   "MoveToEx_LineTo"
   ClientHeight    =   5595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6885
   LinkTopic       =   "Form1"
   ScaleHeight     =   373
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   459
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "Draw Line To"
      Height          =   600
      Left            =   4200
      TabIndex        =   10
      Top             =   4050
      Width           =   1500
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Move Current Position"
      Height          =   600
      Left            =   900
      TabIndex        =   9
      Top             =   4050
      Width           =   1500
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   4650
      TabIndex        =   8
      Top             =   3600
      Width           =   1200
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   4650
      TabIndex        =   7
      Top             =   3000
      Width           =   1200
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1350
      TabIndex        =   6
      Top             =   3600
      Width           =   1200
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1350
      TabIndex        =   5
      Top             =   3000
      Width           =   1200
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Back"
      Height          =   450
      Left            =   4050
      TabIndex        =   0
      Top             =   4920
      Width           =   1500
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "0,0"
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
      Left            =   1395
      TabIndex        =   12
      Top             =   5175
      Width           =   330
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Old Current Position Coordinates"
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
      Left            =   300
      TabIndex        =   11
      Top             =   4875
      Width           =   2775
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Y Position"
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
      Left            =   3675
      TabIndex        =   4
      Top             =   3630
      Width           =   870
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "X Position"
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
      Left            =   3675
      TabIndex        =   3
      Top             =   3030
      Width           =   870
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Y Position"
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
      Left            =   375
      TabIndex        =   2
      Top             =   3630
      Width           =   870
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "X Position"
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
      Left            =   375
      TabIndex        =   1
      Top             =   3030
      Width           =   870
   End
End
Attribute VB_Name = "MoveToEx_LineTo"
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

'variable that stores the x and y
'coordinates of the old position
Dim oldcoordinates As POINTAPI

'MoveToEx declaration
Private Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, lpPoint As POINTAPI) As Long

'LineTo declaration
Private Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long

Private Sub Command1_Click()
    Unload Me 'Unload the form
End Sub

Private Sub Command2_Click()
    'move the current position with MoveToEx
    MoveToEx MoveToEx_LineTo.hdc, Val(Text1.Text), _
        Val(Text2.Text), oldcoordinates
    'set the old current position coordinates to
    'the label in the bottom right of the screen
    Label6.Caption = oldcoordinates.x & " , " & _
        oldcoordinates.y
End Sub

Private Sub Command3_Click()
    'draw a line from the current position to
    'the specified coordinates with LineTo
    LineTo MoveToEx_LineTo.hdc, Val(Text3.Text), _
        Val(Text4.Text)
End Sub

