VERSION 5.00
Begin VB.Form GetAsyncKeyStateForm 
   Caption         =   "GetAsyncKeyState"
   ClientHeight    =   2100
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2805
   LinkTopic       =   "Form1"
   ScaleHeight     =   140
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   187
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Back"
      Height          =   450
      Left            =   750
      TabIndex        =   1
      Top             =   1500
      Width           =   1350
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   2280
      Top             =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Press any letter on the keyboard"
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
      Left            =   0
      TabIndex        =   0
      Top             =   75
      Width           =   2775
   End
End
Attribute VB_Name = "GetAsyncKeyStateForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'GetAsyncKeyState Declaration
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

Private Sub Command1_Click()
Unload Me 'unload the form
End Sub

Private Sub Form_Load()
    'Make the form font bold
    GetAsyncKeyStateForm.Font.Bold = True
End Sub

Private Sub Timer1_Timer()
    Me.Cls 'clear the form of previous text
    CurrentY = 25 'set the current y position for text
    'Determine what keys were pressed
    If GetAsyncKeyState(vbKeyA) Then Print "A was pressed"
    If GetAsyncKeyState(vbKeyB) Then Print "B was pressed"
    If GetAsyncKeyState(vbKeyC) Then Print "C was pressed"
    If GetAsyncKeyState(vbKeyD) Then Print "D was pressed"
    If GetAsyncKeyState(vbKeyE) Then Print "E was pressed"
    If GetAsyncKeyState(vbKeyF) Then Print "F was pressed"
    If GetAsyncKeyState(vbKeyG) Then Print "G was pressed"
    If GetAsyncKeyState(vbKeyH) Then Print "H was pressed"
    If GetAsyncKeyState(vbKeyI) Then Print "I was pressed"
    If GetAsyncKeyState(vbKeyJ) Then Print "J was pressed"
    If GetAsyncKeyState(vbKeyK) Then Print "K was pressed"
    If GetAsyncKeyState(vbKeyL) Then Print "L was pressed"
    If GetAsyncKeyState(vbKeyM) Then Print "M was pressed"
    If GetAsyncKeyState(vbKeyN) Then Print "N was pressed"
    If GetAsyncKeyState(vbKeyO) Then Print "O was pressed"
    If GetAsyncKeyState(vbKeyP) Then Print "P was pressed"
    If GetAsyncKeyState(vbKeyQ) Then Print "Q was pressed"
    If GetAsyncKeyState(vbKeyR) Then Print "R was pressed"
    If GetAsyncKeyState(vbKeyS) Then Print "S was pressed"
    If GetAsyncKeyState(vbKeyT) Then Print "T was pressed"
    If GetAsyncKeyState(vbKeyU) Then Print "U was pressed"
    If GetAsyncKeyState(vbKeyV) Then Print "V was pressed"
    If GetAsyncKeyState(vbKeyW) Then Print "W was pressed"
    If GetAsyncKeyState(vbKeyX) Then Print "X was pressed"
    If GetAsyncKeyState(vbKeyY) Then Print "Y was pressed"
    If GetAsyncKeyState(vbKeyZ) Then Print "Z was pressed"
End Sub
