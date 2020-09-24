VERSION 5.00
Begin VB.Form GetDiskFreeSpaceform 
   Caption         =   "GetDiskFreeSpace"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3885
   LinkTopic       =   "Form1"
   ScaleHeight     =   213
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   259
   StartUpPosition =   2  'CenterScreen
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   825
      TabIndex        =   2
      Top             =   675
      Width           =   2250
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Back"
      Height          =   600
      Left            =   1200
      TabIndex        =   1
      Top             =   2400
      Width           =   1500
   End
   Begin VB.CommandButton Command1 
      Caption         =   "GetDiskFreeSpace"
      Height          =   750
      Left            =   1050
      TabIndex        =   0
      Top             =   1275
      Width           =   1815
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Only works for A and C drive"
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
      TabIndex        =   3
      Top             =   225
      Width           =   2445
   End
End
Attribute VB_Name = "GetDiskFreeSpaceform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'GetDiskFreeSpace declaration
Private Declare Function GetDiskFreeSpace Lib "kernel32" Alias "GetDiskFreeSpaceA" (ByVal lpRootPathName As String, lpSectorsPerCluster As Long, lpBytesPerSector As Long, lpNumberOfFreeClusters As Long, lpTotalNumberOfClusters As Long) As Long


Private Sub Command1_Click()
    'used to store how many sectors are in each cluster
    Dim SectorsPerCluster As Long
    'used to store how many bytes are in each sector
    Dim BytesPerSector As Long
    'used to store how many free clusters there are
    Dim NumberOfFreeClusters As Long
    'used to store the total number of cluster
    Dim TotalNumberOfClusters As Long
    'used to store how many free bytes their are
    Dim FreeKiloBytes As Variant
    
    'use GetDiskFreeSpace to find out how much
    'free space in bytes is on the selected drive
    GetDiskFreeSpace Drive1.Drive & "\", SectorsPerCluster, BytesPerSector, _
        NumberOfFreeClusters, TotalNumberOfClusters
    'times the number of free clusters times how many
    'sectors are in each cluster times how many bytes
    'are in each sector to get the amount of free bytes
    FreeKiloBytes = (NumberOfFreeClusters / 1000) * _
        SectorsPerCluster * BytesPerSector
    'have to divide the number by 1000 because a variant
    'can't hold that big of numbers
    
    'display the results of the GetDiskFreeSpace
    MsgBox "      " & UCase(Left(Drive1.Drive, 1)) & "  drive free space" & vbNewLine _
        & vbNewLine & "        Bytes:  " & (FreeKiloBytes * 1000) & _
        vbNewLine & "   KiloBytes:  " & FreeKiloBytes & vbNewLine _
        & "MegaBytes:  " & (FreeKiloBytes / 1000) & vbNewLine & _
        "  GigaBytes:  " & (FreeKiloBytes / 1000000), , "GetDiskFreeSpace"

    'if there is no error then exit the sub
    Exit Sub

End Sub

Private Sub Command2_Click()
    Unload Me 'Unload the form
End Sub

   ' The Long FreeBytes contains the number of free bytes on the drive
Private Sub Form_Load()

End Sub
