VERSION 5.00
Begin VB.Form Main 
   Caption         =   "API Functions"
   ClientHeight    =   5190
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8925
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   346
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   595
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   975
      Width           =   3000
   End
   Begin VB.TextBox Variables 
      Height          =   3675
      Left            =   5775
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   8
      Top             =   1350
      Width           =   3000
   End
   Begin VB.TextBox ED 
      Height          =   1620
      Left            =   3300
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   6
      Top             =   1350
      Width           =   2325
   End
   Begin VB.TextBox WID 
      Height          =   510
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   5
      Top             =   4500
      Width           =   5535
   End
   Begin VB.TextBox Declaration 
      Height          =   870
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   3300
      Width           =   5535
   End
   Begin VB.ListBox List1 
      Height          =   1620
      ItemData        =   "Main.frx":0000
      Left            =   120
      List            =   "Main.frx":004F
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   1350
      Width           =   3000
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Double click on the API function to see an example of it"
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
      Left            =   1500
      TabIndex        =   12
      Top             =   375
      Width           =   5670
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Single click on the API function to see more information about it"
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
      Left            =   1650
      TabIndex        =   11
      Top             =   120
      Width           =   5460
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "API Functions - Type the first few letters of the function that you are looking for "
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
      Left            =   120
      TabIndex        =   10
      Top             =   675
      Width           =   6855
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "API Function Variables"
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
      Left            =   6075
      TabIndex        =   9
      Top             =   1050
      Width           =   2370
   End
   Begin VB.Label Label3 
      Caption         =   "Extra Declarations"
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
      Left            =   3525
      TabIndex        =   7
      Top             =   1050
      Width           =   1695
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "What does it do"
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
      TabIndex        =   4
      Top             =   4275
      Width           =   1365
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Declaration"
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
      Top             =   3075
      Width           =   990
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim x As Integer 'Used for the list1 search textbox

Private Sub List1_Click()
 Select Case List1.List(List1.ListIndex)
    Case "Beep"
        Declaration.Text = "Private Declare Function Beep Lib ""kernel32"" (ByVal dwFreq As Long, ByVal dwDuration As Long) As Long"
        WID.Text = "Makes the system beep at a specified frequency and for a specified duration of time"
        ED.Text = "None"
        Variables.Text = "dwFreq - the frequency of the beep (1-10000)" & vbNewLine & vbNewLine & "dwDuration - the duration of the beep in milliseconds"
    Case "BitBlt"
        Declaration.Text = "Private Declare Function BitBlt Lib ""gdi32"" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long"
        WID.Text = "Transfers a rectangular area of pixels from a source device(hdc) to a destination device(hdc)"
        ED.Text = "Private Const SRCCOPY = &HCC0020" & vbNewLine & vbNewLine & " Or you can use the" & vbNewLine & "   VB constant vbSrcCopy"
        Variables.Text = "hDestDC - the hdc of the destination object" & vbNewLine & vbNewLine & "x - the x coordinate at which the destination picture starts at in the object" & vbNewLine & vbNewLine & "y - the y coordinate at which the destination picture starts at in the object" & vbNewLine & vbNewLine & "nWidth - the width of the source picture that will be copied" & vbNewLine & vbNewLine & "nHeight - the height of the source picture that will be copied" & vbNewLine & vbNewLine & "hSrcDC - the hdc of the source object" & vbNewLine & vbNewLine & "xSrc - the x coodrinate at which the picture will start at in the source object" & vbNewLine & vbNewLine & "ySrc - the y coodrinate at which the picture will start at in the source object" & vbNewLine & vbNewLine & "dwRop - a constant defining what type of copy will take place, look for RasterOpConstants in the VB Object Browser for all of the possible constants"
    Case "ClipCursor"
        Declaration.Text = "Private Declare Function ClipCursor Lib ""user32"" (lpRect As Any) As Long"
        WID.Text = "Restricts mouse movement to within a specified rectangle using the RECT type"
        ED.Text = "Public Type RECT" & vbNewLine & "        Left As Long" & vbNewLine & "        Top As Long" & vbNewLine & "        Right As Long" & vbNewLine & "        Bottom As Long" & vbNewLine & "End Type"
        Variables.Text = "lpRect - a RECT type variable that holds the dimensions for a rectangle"
    Case "Ellipse"
        Declaration.Text = "Private Declare Function Ellipse Lib ""gdi32"" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long"
        WID.Text = "Draws an ellipse that fits inside of a specified rectangle, this is much faster than the circle function in VB"
        ED.Text = "None"
        Variables.Text = "hdc - the hdc of the object that you want to draw the ellipse on" & vbNewLine & vbNewLine & "X1 - the x coordinate of the top left corner of the rectangle that the ellipse will go in" & vbNewLine & vbNewLine & "Y1 - the y coordinate of the top left corner of the rectangle that the ellipse will go in" & vbNewLine & vbNewLine & "X2 - the x coordinate of the bottom right corner of the rectangle that the ellipse will go in" & vbNewLine & vbNewLine & "Y2 - the y coordinate of the bottom right corner of the rectangle that the ellipse will go in"
    Case "GetAsyncKeyState"
        Declaration.Text = "Private Declare Function GetAsyncKeyState Lib ""user32"" (ByVal vKey As Long) As Integer"
        WID.Text = "Checks if a certain key is pressed even if the form is not active"
        ED.Text = "None"
        Variables.Text = "vKey - the VB keycode constant for what key you want to detect"
    Case "GetCursorPos"
        Declaration.Text = "Private Declare Function GetCursorPos Lib ""user32"" (lpPoint As POINTAPI) As Long"
        WID.Text = "Gets the x and y coordinates of where the mouse cursor is and stores them in a POINTAPI variable"
        ED.Text = "Public Type POINTAPI" & vbNewLine & "   x as long" & vbNewLine & "   y as long" & vbNewLine & "End Type"
        Variables.Text = "lpPoint - a POINTAPI variable that you want to store the x and y coordinates into"
    Case "GetDC"
        Declaration.Text = "Private Declare Function GetDC Lib ""user32"" (ByVal hwnd As Long) As Long"
        WID.Text = "Gets the hdc of an object when given its hwnd"
        ED.Text = "None"
        Variables.Text = "hwnd - the hwnd of the object that you want to get the hdc from"
    Case "GetDesktopWindow"
        Declaration.Text = "Private Declare Function GetDesktopWindow Lib ""user32"" () As Long"
        WID.Text = "Gets the hwnd of the desktop window"
        ED.Text = "None"
        Variables.Text = "None"
    Case "GetDiskFreeSpace"
        Declaration.Text = "Private Declare Function GetDiskFreeSpace Lib ""kernel32"" Alias ""GetDiskFreeSpaceA"" (ByVal lpRootPathName As String, lpSectorsPerCluster As Long, lpBytesPerSector As Long, lpNumberOfFreeClusters As Long, lpTotalNumberOfClusters As Long) As Long"
        WID.Text = "Gets the amount of free space in a specified drive"
        ED.Text = "None"
        Variables.Text = "lpRootPathName - the name of a drive that you want to check the free space of followed by a semicolon and a \" & vbNewLine & vbNewLine & "lpSectorsPerCluster - the name of a variable that you want to store the amount of sectors per cluster in" & vbNewLine & vbNewLine & "lpBytesPerCluster - the name of a variable that you want to store the amount of bytes per cluster in" & vbNewLine & vbNewLine & "lpNumberOfFreeClusters - the name of a variable that you want to store the amount of free clusters" & vbNewLine & vbNewLine & "lpTotalNumberOfClusters - the name of a variable that you want to store the total number of clusters in"
    Case "GetDoubleClickTime"
        Declaration.Text = "Private Declare Function GetDoubleClickTime Lib ""user32"" () As Long"
        WID.Text = "Gets the current time allowed in milliseconds between two mouse clicks for a double click"
        ED.Text = "None"
        Variables.Text = "None"
    Case "GetKeyState"
        Declaration.Text = "Private Declare Function GetKeyState Lib ""user32"" (ByVal nVirtKey As Long) As Integer"
        WID.Text = "Checks the current state of any key on the keyboard"
        ED.Text = "None"
        Variables.Text = "nVirtKey - the VB keycode constant for what key you want to check"
    Case "GetPixel"
        Declaration.Text = "Private Declare Function GetPixel Lib ""gdi32"" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long"
        WID.Text = "Gets the color of a single pixel when given the x and y coordinates of the pixel"
        ED.Text = "None"
        Variables.Text = "hdc - the hdc of the object on which you want to get the pixel from" & vbNewLine & vbNewLine & "x - the x coordinate of the pixel that you want" & vbNewLine & vbNewLine & "y - the y coordinate of the pixel that you want"
    Case "GetWindowRect"
        Declaration.Text = "Private Declare Function GetWindowRect Lib ""user32"" (ByVal hwnd As Long, lpRect As RECT) As Long"
        WID.Text = "Gets the dimensions of an object specified by the hwnd and stores them in a RECT type variable"
        ED.Text = "Public Type RECT" & vbNewLine & "        Left As Long" & vbNewLine & "        Top As Long" & vbNewLine & "        Right As Long" & vbNewLine & "        Bottom As Long" & vbNewLine & "End Type"
        Variables.Text = "hwnd - the hwnd of the object that you want to get the dimensions of" & vbNewLine & vbNewLine & "lpRect - a Rect type variable that has the dimensions of a rectangle to restrict the mouse movement to"
    Case "LineTo"
        Declaration.Text = "Private Declare Function LineTo Lib ""gdi32"" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long"
        WID.Text = "Draws a line from the current position to a specified position"
        ED.Text = "None"
        Variables.Text = "hdc - the hdc of the object that you want to draw the line on" & vbNewLine & vbNewLine & "x - x the coordinate that the line will be drawn to" & vbNewLine & vbNewLine & "y - the y coordinate that the line will be drawn to"
    Case "MoveToEx"
        Declaration.Text = "Private Declare Function MoveToEx Lib ""gdi32"" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, lpPoint As POINTAPI) As Long"
        WID.Text = "Moves the current position to the specified coordinates"
        ED.Text = "Public Type POINTAPI" & vbNewLine & "   x as long" & vbNewLine & "   y as long" & vbNewLine & "End Type"
        Variables.Text = "hdc - the hdc of the object that you want to move the current position in" & vbNewLine & vbNewLine & "x - the x coordinate of the new current position" & vbNewLine & vbNewLine & "y - the y coordinate of the new current position" & vbNewLine & vbNewLine & "lpPoint - a POINTAPI variable that will store the old x and y coordinates of the current position"
    Case "Polyline"
        Declaration.Text = "Private Declare Function Polyline Lib ""gdi32"" (ByVal hdc As Long, lpPoint As POINTAPI, ByVal nCount As Long) As Long"
        WID.Text = "Draws a set of lines using an array POINTAPI type variables"
        ED.Text = "Public Type POINTAPI" & vbNewLine & "   x as long" & vbNewLine & "   y as long" & vbNewLine & "End Type"
        Variables.Text = "hdc - the hdc of the object that you want to draw the line on" & vbNewLine & vbNewLine & "lpPoint - a POINTAPI variable that defines the starts and ends of the lines" & vbNewLine & vbNewLine & "nCount - the number of points that are in the POINTAPI array"
    Case "Rectangle"
        Declaration.Text = "Private Declare Function Rectangle Lib ""gdi32"" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long"
        WID.Text = "Draws a rectangle when the x and y coordinates of the top left and bottom right corners are specified, this is much faster than the Line BF function in VB"
        ED.Text = "None"
        Variables.Text = "hdc - the hdc of the object that you want to draw the rectangle on" & vbNewLine & vbNewLine & "X1 - the x coordinate of the top left corner of the rectangle" & vbNewLine & vbNewLine & "Y1 - the y coordinate of the top left corner of the rectangle" & vbNewLine & vbNewLine & "X2 - the x coordinate of the bottom right corner of the rectangle" & vbNewLine & vbNewLine & "Y2 - the y coordinate of the bottom right corner of the rectangle"
    Case "SetCursorPos"
        Declaration.Text = "Private Declare Function SetCursorPos Lib ""user32"" (ByVal x As Long, ByVal y As Long) As Long"
        WID.Text = "Sets the mouse cursor at a specified location"
        ED.Text = "None"
        Variables.Text = "x - the x coordinate of where you want the mouse cursor to be on the screen" & vbNewLine & vbNewLine & "y - the y coordinate of where you want the mouse cursor to be on the screen"
    Case "SetDoubleClickTime"
        Declaration.Text = "Private Declare Function SetDoubleClickTime Lib ""user32"" (ByVal wCount As Long) As Long"
        WID.Text = "Sets the time allowed in milliseconds between two mouse clicks for a double click"
        ED.Text = "None"
        Variables.Text = "wCount - the time in milliseconds for the new time allowed between mouse clicks for a double click"
    Case "SetPixel"
        Declaration.Text = "Private Declare Function SetPixel Lib ""gdi32"" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long"
        WID.Text = "Sets the color of a single pixel when given the x and y coordinates of the pixel"
        ED.Text = "None"
        Variables.Text = "hdc - the hdc of the object that you want to set the pixel on" & vbNewLine & vbNewLine & "x - the x coordinate of the pixel that you want to set" & vbNewLine & vbNewLine & "y - the y coordinate of the pixel that you want to set" & vbNewLine & vbNewLine & "crColor - the color of the pixel that you want to set"
    Case "ShowCursor"
        Declaration.Text = "Private Declare Function ShowCursor Lib ""user32"" (ByVal bShow As Long) As Long"
        WID.Text = "Hides and shows the mouse cursor"
        ED.Text = "None"
        Variables.Text = "bShow - a statement that evaluates to true or false"
    Case "Sleep"
        Declaration.Text = "Private Declare Sub Sleep Lib ""kernel32"" (ByVal dwMilliseconds As Long)"
        WID.Text = "Causes the system to sleep/pause for a specified amount of milliseconds"
        ED.Text = "None"
        Variables.Text = "dwMilliseconds - the duration of the sleep in milliseconds"
    Case "sndPlaySound"
        Declaration.Text = "Private Declare Function sndPlaySound Lib ""winmm.dll"" Alias ""sndPlaySoundA"" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long"
        WID.Text = "Plays a WAV file"
        ED.Text = "Private Const SND_SYNC = &H0"
        Variables.Text = "lpszSoundName - the file path of the WAV file that you want to play" & vbNewLine & vbNewLine & "uFlags - a flag that will determine how the WAV file will play, type in SND in the API text viewer under the constants category to see all possible constants"
    Case "StretchBlt"
        Declaration.Text = "Private Declare Function StretchBlt Lib ""gdi32"" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long"
        WID.Text = "Copies a rectangular area of a bitmap to another area and stretches or shrinks the original bitmap"
        ED.Text = "Private Const SRCCOPY = &HCC0020" & vbNewLine & vbNewLine & " Or you can use the" & vbNewLine & "   VB constant vbSrcCopy"
        Variables.Text = "hdc - the hdc of the destination object" & vbNewLine & vbNewLine & "x - the x coordinate at which the destination picture starts at in the object" & vbNewLine & vbNewLine & "y - the y coordinate at which the destination picture starts at in the object" & vbNewLine & vbNewLine & "nWidth - the width of the new picture to be stretched or shrunk to" & vbNewLine & vbNewLine & "nHeight - the height of the new picture to be stretched or shrunk to" & vbNewLine & vbNewLine & "hSrcDC - the hdc of the source object" & vbNewLine & vbNewLine & "xSrc - the x coodrinate at which the picture will start at in the source object" & vbNewLine & vbNewLine & "ySrc - the y coodrinate at which the picture will start at in the source object" & vbNewLine & vbNewLine & _
                            "nSrcWidth - the width of the source picture that will be copied" & vbNewLine & vbNewLine & "nSrcHeight - the height of the source picture that will be copied" & vbNewLine & vbNewLine & "dwRop - a constant defining what type of copy will take place, look for RasterOpConstants in the VB Object Browser for all of the possible constants"
    Case "SwapMouseButtons"
        Declaration.Text = "Private Declare Function SwapMouseButton Lib ""user32"" (ByVal bSwap As Long) As Long"
        WID.Text = "Switches the right and left mouse buton clicks"
        ED.Text = "None"
       Variables.Text = "bSwap - a statement that evaluates to true or false"
    End Select
End Sub

Private Sub List1_DblClick()
    Select Case List1.List(List1.ListIndex)
        Case "Beep"
            BeepForm.Show
        Case "BitBlt"
            BitBltForm.Show
        Case "ClipCursor"
            GetWindowRect_ClipCursor.Show
        Case "Ellipse"
            Ellipseform.Show
        Case "GetAsyncKeyState"
            GetAsyncKeyStateForm.Show
        Case "GetCursorPos"
            GetCursorPos.Show
        Case "GetDC"
            GetDesktopWindow_GetDC.Show
        Case "GetDesktopWindow"
            Main.WindowState = vbMinimized
            GetDesktopWindow_GetDC.Show
        Case "GetDiskFreeSpace"
            GetDiskFreeSpaceform.Show
        Case "GetDoubleClickTime"
            GetDoubleClickTime_SetDoubleClickTime.Show
        Case "GetKeyState"
            GetKeyState.Show
        Case "GetPixel"
            GetPixelForm.Show
        Case "GetWindowRect"
            GetWindowRect_ClipCursor.Show
        Case "LineTo"
            MoveToEx_LineTo.Show
        Case "MoveToEx"
            MoveToEx_LineTo.Show
        Case "Polyline"
            Polylineform.Show
        Case "Rectangle"
            Rectangleform.Show
        Case "SetCursorPos"
            SetCursorPos.Show
        Case "SetDoubleClickTime"
            GetDoubleClickTime_SetDoubleClickTime.Show
        Case "SetPixel"
            SetPixelForm.Show
        Case "ShowCursor"
            ShowCursor.Show
        Case "Sleep"
            Sleep.Show
        Case "sndPlaySound"
            sndPlaySound.Show
        Case "StretchBlt"
            StretchBlt.Show
        Case "SwapMouseButtons"
            SwapMouseButtons.Show
        End Select
End Sub

Private Sub Text2_Change()
    On Error GoTo Error
    For x = 0 To List1.ListCount
        If LCase(Left(List1.List(x), Len(Text2.Text))) = Text2.Text Then
            List1.TopIndex = x
            List1.ListIndex = x
            Exit Sub
        End If
    Next x
    Exit Sub
Error:
    List1.ListIndex = 0
End Sub
