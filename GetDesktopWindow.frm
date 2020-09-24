VERSION 5.00
Begin VB.Form GetDesktopWindow_GetDC 
   AutoRedraw      =   -1  'True
   Caption         =   "GetDesktopWindow & GetDC"
   ClientHeight    =   5595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7875
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   373
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   525
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   3600
      Top             =   600
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Paint Desktop"
      Height          =   750
      Left            =   375
      TabIndex        =   0
      Top             =   375
      Width           =   1500
   End
End
Attribute VB_Name = "GetDesktopWindow_GetDC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim desktophwnd 'variable for storing the desktop hwnd
Dim desktophdc 'variable for storing the desktop hdc

'const needed for BitBlt
Private Const SRCCOPY = &HCC0020

'GetDesktopWindow Declaration
Private Declare Function GetDesktopWindow Lib "user32" () As Long

'GetDC Declaration
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long

'BitBlt declaration
'Used to paint the desktop onto the form background
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Private Sub Command1_Click()
'minimize the form so it is not in the way of the desktop
    GetDesktopWindow_GetDC.WindowState = vbMinimized
    Command1.Visible = False 'hide the button
    Timer1.Enabled = True 'start the timer
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
'So the user can get back to the main screen
    If KeyAscii = vbKeyEscape Then
        Unload GetDesktopWindow_GetDC 'unloads the form
        Main.WindowState = vbNormal 'returns the main screen to normal
    End If
End Sub

Private Sub Form_Load()
    MsgBox "Press the button in the top left of the form to paint" & vbNewLine & "   the desktop onto the background of the form", , "GetDesktopWindow & GetDC"
End Sub

Private Sub Timer1_Timer()
    desktophwnd = GetDesktopWindow 'get the hwnd of the desktop
    
    desktophdc = GetDC(desktophwnd) 'get the hdc of the desktop
    
    GetDesktopWindow_GetDC.WindowState = vbMaximized 'maximize the form
    
'paint the desktop window onto the background of the form with BitBlt
    BitBlt GetDesktopWindow_GetDC.hdc, 0, 0, GetDesktopWindow_GetDC.ScaleWidth _
        , GetDesktopWindow_GetDC.ScaleHeight, desktophdc, 0, 0, SRCCOPY
    Timer1.Enabled = False 'turn off the timer
    MsgBox "Press Esc to go back", , "GetDesktopWindow & GetDC"
End Sub
