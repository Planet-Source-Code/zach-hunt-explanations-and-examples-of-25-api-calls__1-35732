VERSION 5.00
Begin VB.Form Rectangleform 
   Caption         =   "Rectangle"
   ClientHeight    =   5595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6885
   LinkTopic       =   "Form1"
   ScaleHeight     =   373
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   459
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   4800
      TabIndex        =   8
      Top             =   3825
      Width           =   1200
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   4800
      TabIndex        =   7
      Top             =   3150
      Width           =   1200
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Draw Rectangle"
      Height          =   525
      Left            =   2400
      TabIndex        =   9
      Top             =   4350
      Width           =   2100
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1950
      TabIndex        =   4
      Top             =   3825
      Width           =   1200
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1950
      TabIndex        =   3
      Top             =   3150
      Width           =   1200
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Back"
      Height          =   450
      Left            =   2775
      TabIndex        =   0
      Top             =   5025
      Width           =   1500
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Y2 Position"
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
      Left            =   3600
      TabIndex        =   6
      Top             =   3855
      Width           =   975
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "X2 Position"
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
      Left            =   3600
      TabIndex        =   5
      Top             =   3180
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Y1 Position"
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
      Left            =   750
      TabIndex        =   2
      Top             =   3855
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "X1 Position"
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
      Left            =   750
      TabIndex        =   1
      Top             =   3180
      Width           =   1005
   End
End
Attribute VB_Name = "Rectangleform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Rectangle declaration
Private Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long

Private Sub Command1_Click()
    Unload Me 'Unload the form
End Sub

Private Sub Command2_Click()
    'make the rectangle easier to see
    Rectangleform.DrawWidth = 3
    'draw the rectangle using the specified coordinates
    Rectangle Rectangleform.hdc, Val(Text1.Text), _
        Val(Text2.Text), Val(Text3.Text), Val(Text4.Text)
End Sub

