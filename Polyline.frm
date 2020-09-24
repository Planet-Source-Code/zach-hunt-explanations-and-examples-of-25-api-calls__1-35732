VERSION 5.00
Begin VB.Form Polylineform 
   AutoRedraw      =   -1  'True
   Caption         =   "Polyline"
   ClientHeight    =   5595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7875
   LinkTopic       =   "Form1"
   ScaleHeight     =   373
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   525
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Polyline"
      Height          =   675
      Left            =   6120
      TabIndex        =   1
      Top             =   4125
      Width           =   1650
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Back"
      Height          =   525
      Left            =   6375
      TabIndex        =   0
      Top             =   4950
      Width           =   1200
   End
End
Attribute VB_Name = "Polylineform"
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

'array used by polyline function to draw multiple lines
Dim points(11) As POINTAPI
Const size = 11 'the size of the points array

'Polyline declaration
Private Declare Function Polyline Lib "gdi32" (ByVal hdc As Long, lpPoint As POINTAPI, ByVal nCount As Long) As Long

Private Sub Command1_Click()
    Unload Me 'Unload the form
End Sub

Private Sub Command2_Click()
    Dim i As Integer 'used for a for loop
    For i = 0 To size Step 2
        points(i).x = (i - 1) * 56 + 10
        points(i).y = 5
        points(i + 1).x = (i - 1) * 36 + 10
        points(i + 1).y = 365
    Next i
    'draws the lines starting with the first
    'point in the points array points(1)
    Polyline Polylineform.hdc, points(0), size
    Me.Refresh 'needed to put the lines onto the screen
End Sub


