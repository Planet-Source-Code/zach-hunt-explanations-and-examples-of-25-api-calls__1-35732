VERSION 5.00
Begin VB.Form SndPlaySound 
   Caption         =   "SndPlaySound"
   ClientHeight    =   2790
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3885
   LinkTopic       =   "Form1"
   ScaleHeight     =   186
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   259
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "Stop Sound"
      Height          =   900
      Left            =   2100
      TabIndex        =   2
      Top             =   375
      Width           =   1500
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Back"
      Height          =   600
      Left            =   1200
      TabIndex        =   1
      Top             =   1800
      Width           =   1500
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Play Sound"
      Height          =   900
      Left            =   300
      TabIndex        =   0
      Top             =   375
      Width           =   1500
   End
End
Attribute VB_Name = "SndPlaySound"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'look in the API viewer for other SND constants
Const SND_SYNC = &H0         'play synchronously (default)

'other constants that can be used for the sndPlaySound API
Private Const SND_ASYNC = &H1 'play asynchronously
Private Const SND_LOOP = &H8 'loop the sound until next sndPlaySound must be combine with SND_ASYNC
Private Const SND_MEMORY = &H4 'pszSoundName points to a memory file
Private Const SND_NODEFAULT = &H2 'silence not default, if sound not found
Private Const SND_NOSTOP = &H10 'don't stop any currently playing sound
Private Const SND_NOWAIT = &H2000 'don't wait if the driver is busy
Private Const SND_PURGE = &H40 'purge non-static events for task


'sndPlaySound declaration
Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

Private Sub Command1_Click()
    Dim flags
    'determine if the user wants to loop the sound
    If MsgBox("Loop the sound", vbYesNo, "     Loop Sound") = vbYes Then
        'SND_LOOP & SND_SYNC are needed to loop a sound
        flags = SND_LOOP + SND_ASYNC
    Else
        'the default or most common is SND_SYNC
        flags = SND_SYNC
    End If
    'play a WAV file using the sndPlaySound function
    sndPlaySound App.Path & "/ringin.wav", flags
End Sub


Private Sub Command2_Click()
    Unload Me 'unload the form
End Sub

Private Sub Command3_Click()
    'to stop a looping sound you have to
    'call the sndPlaySound function again
    sndPlaySound App.Path & "/ringin.wav", SND_SYNC
End Sub
