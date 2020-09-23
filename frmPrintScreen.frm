VERSION 5.00
Begin VB.Form frmPrintScreen 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Print On Screen Example - By: Synthesize"
   ClientHeight    =   6975
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8475
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6975
   ScaleWidth      =   8475
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picSonicFlood 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Height          =   1650
      Left            =   -20000
      Picture         =   "frmPrintScreen.frx":0000
      ScaleHeight     =   1590
      ScaleWidth      =   9600
      TabIndex        =   7
      Top             =   0
      Width           =   9660
   End
   Begin VB.TextBox txtDeskOffset 
      Height          =   360
      Index           =   1
      Left            =   4240
      TabIndex        =   6
      Text            =   "Pixel ""Y"" Offset (Text/Picture On Desktop)"
      ToolTipText     =   "Pixel ""Y"" Offset - Text/Picture On Desktop (You are only allowed you input numbers [0-9])"
      Top             =   6300
      Width           =   4235
   End
   Begin VB.TextBox txtDeskOffset 
      Height          =   360
      Index           =   0
      Left            =   0
      TabIndex        =   5
      Text            =   "Pixel ""X"" Offset (Text/Picture On Desktop)"
      ToolTipText     =   "Pixel ""X"" Offset - Text/Picture On Desktop (You are only allowed you input numbers [0-9])"
      Top             =   6300
      Width           =   4235
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print It!"
      Height          =   315
      Left            =   0
      TabIndex        =   4
      ToolTipText     =   "Click here to print the picture/text onto the screen"
      Top             =   6660
      Width           =   8475
   End
   Begin VB.TextBox txtPicOffset 
      Height          =   360
      Index           =   1
      Left            =   4240
      TabIndex        =   3
      Text            =   "Pixel ""Y"" Offset (Text In PictureBox)"
      ToolTipText     =   "Pixel ""Y"" Offset - Text In PictureBox (You are only allowed you input numbers [0-9])"
      Top             =   5940
      Width           =   4235
   End
   Begin VB.TextBox txtPicOffset 
      Height          =   360
      Index           =   0
      Left            =   0
      TabIndex        =   2
      Text            =   "Pixel ""X"" Offset (Text In PictureBox)"
      ToolTipText     =   "Pixel ""X"" Offset - Text In PictureBox (You are only allowed you input numbers [0-9])"
      Top             =   5940
      Width           =   4235
   End
   Begin VB.TextBox txtText 
      Height          =   1575
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Text            =   "frmPrintScreen.frx":4C25
      ToolTipText     =   "Text to print onto the screen"
      Top             =   4320
      Width           =   8475
   End
   Begin VB.PictureBox picSource 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   4275
      Left            =   0
      ScaleHeight     =   281
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   561
      TabIndex        =   0
      ToolTipText     =   "Source PictureBox"
      Top             =   0
      Width           =   8475
   End
End
Attribute VB_Name = "frmPrintScreen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Declare our API calls.
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long

' Declare our constants.
Private Const SRCCOPY = &HCC0020

Private Sub cmdPrint_Click()
    ' Print the text/picture onto the screen.
    Call PrintScreen(picSource, Val(txtDeskOffset(0)), Val(txtDeskOffset(1)))
End Sub

Private Sub Form_Load()
    ' Call the txtText_Change function so that the PictureBox
    '   will contain the picture and the text.
    Call txtText_Change
End Sub

Private Sub txtDeskOffset_KeyPress(Index As Integer, KeyAscii As Integer)
    ' Declare our variables.
    Dim strChr As String
    
    ' Get the character that was pressed and
    '   set it to our variable strChr.
    strChr = Chr(KeyAscii)
    
    ' If strChr is NOT in the set [0-9] AND is NOT
    '   BackSpace, set the the KeyAscii to 0 so that
    '   the TextBox thinks that no key was pressed.
    If (Not strChr Like "[0-9]") And (KeyAscii <> 8) Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtPicOffset_Change(Index As Integer)
    ' Call the txtText_Change function so that we
    '   refresh the PictureBox.
    Call txtText_Change
End Sub

Private Sub txtText_Change()
    ' Set picSource's Picture to nothing.
    Set picSource.Picture = Nothing
    ' Copy and stretch the picture from picSonicFlood
    '   to picSource.
    Call StretchPicture(picSonicFlood, picSource)
    ' Clear extra painted stuff on picSource.
    picSource.Cls
    ' Set the CurrentX to the value in txtPicOffset(0).
    picSource.CurrentX = Val(txtPicOffset(0))
    ' Set the CurrentY to the value in txtPicOffset(1).
    picSource.CurrentY = Val(txtPicOffset(1))
    ' Print the text in txtText onto the picture.
    picSource.Print txtText.Text
    ' "Stain" the image into the picture (set the picture
    '   to what is visible in the PictureBox).
    Set picSource.Picture = picSource.Image
End Sub

Private Sub txtPicOffset_KeyPress(Index As Integer, KeyAscii As Integer)
    ' Declare our variables.
    Dim strChr As String
    
    ' Get the character that was pressed and
    '   set it to our variable strChr.
    strChr = Chr(KeyAscii)
    
    ' If strChr is NOT in the set [0-9] AND is NOT
    '   BackSpace, set the the KeyAscii to 0 so that
    '   the TextBox thinks that no key was pressed.
    If (Not strChr Like "[0-9]") And (KeyAscii <> 8) Then
        KeyAscii = 0
    End If
End Sub

Public Sub PrintScreen(ByVal PictureBoxSource As PictureBox, Optional ByVal XOffset As Single = 0, Optional ByVal YOffset As Single = 0)
    ' This sub will print some text and/or a picture
    '   to the screen where select.
    
    ' Declare our variables.
    Dim lngColor As Long
    Dim DeskTopWin As Long
    Dim DeskTopDC As Long
    Dim i, i2
    
    ' Get the handle of the desktop window (Desktop
    '   Window's Class Name: #32769).
    DeskTopWin = FindWindow("#32769", vbNullString)
    ' Get the DC of the desktop window from it's handle.
    DeskTopDC = GetDC(DeskTopWin)
    
    ' I couldn't find a good and fast way to refresh
    '   the desktop, so I used my own way!
    Call RefreshDesktop
    
    ' Set PictureBoxSource's ScaleMode to pixels (vbPixels).
    PictureBoxSource.ScaleMode = vbPixels
    
    ' Do a loop through the ScaleWidth of PictureBoxSource.
    For i = 0 To PictureBoxSource.ScaleWidth - 1
        ' Do a loop through the ScaleHeight of PictureBoxSource.
        For i2 = 0 To PictureBoxSource.ScaleHeight - 1
            ' Get the color of the pixel in PictureBoxSource
            '   from the current X and the current Y.
            lngColor = GetPixel(PictureBoxSource.hdc, i, i2)
            ' If lngColor is NOT the color of PictureBoxSource's
            '   BackColor, allow the pixel to be painted to the
            '   desktop.
            If lngColor <> PictureBoxSource.BackColor Then
                ' Paint the pixel of the current X and Y of the
                '   color that was in PictureBoxSource.
                Call SetPixel(DeskTopDC, XOffset + i, YOffset + i2, lngColor)
            End If
        ' Continue through our loop of the ScaleHeight.
        Next i2
    ' Continue through our loop of the ScaleWidth.
    Next i
End Sub

Public Sub StretchPicture(ByVal SourceBox As PictureBox, ByVal DestBox As PictureBox)
    ' This sub will stretch the picture of SourceBox
    '   to the ScaleWidth and ScaleHeight and set it
    '   to DestBox.
    
    ' Set DestBox's ScaleMode to pixels (vbPixels).
    DestBox.ScaleMode = vbPixels
    ' Set SourceBox's ScaleMode to pixels (vbPixels).
    SourceBox.ScaleMode = vbPixels
    ' Stretch the picture of SourceBox and paint it
    '   on DestBox.
    Call StretchBlt(DestBox.hdc, 0&, 0&, DestBox.ScaleWidth, DestBox.ScaleHeight, SourceBox.hdc, 0&, 0&, SourceBox.ScaleWidth, SourceBox.ScaleHeight, SRCCOPY)
    ' Set the Picture of DestBox to the image (what
    '   is visible in the PictureBox) of DestBox.
    Set DestBox.Picture = DestBox.Image
End Sub

Public Sub RefreshDesktop()
    ' This is my simple way of refreshing the desktop.
    '   This sub simply just copies an image of the
    '   desktop to frmRefresh, shows frmRefresh, then
    '   Unloads frmRefresh.
    
    ' Declare our variables.
    Dim DeskTopWin As Long
    Dim DeskTopDC As Long
    
    ' Get the handle of the desktop window (Desktop
    '   Window's Class Name: #32769).
    DeskTopWin = FindWindow("#32769", vbNullString)
    ' Get the DC of the desktop window from it's handle.
    DeskTopDC = GetDC(DeskTopWin)
    
    ' Load frmRefresh so that we can use it.
    Load frmRefresh
    ' Allow it to load.
    DoEvents
    ' Set frmRefresh's AutoRedraw property to True so
    '   that the image of the desktop will still be
    '   visible when we show frmRefresh.
    frmRefresh.AutoRedraw = True
    ' Set frmRefresh's ScaleMode to pixels (vbPixels).
    frmRefresh.ScaleMode = vbPixels
    ' Set frmRefresh's WindowState to vbMaximized so
    '   that when we show the form, it will be maximized.
    frmRefresh.WindowState = vbMaximized
    
    ' Allow the properties to change before we move on.
    DoEvents
    ' Copy an image of the desktop to frmRefresh's interface.
    Call StretchBlt(frmRefresh.hdc, 0&, 0&, Screen.Width / Screen.TwipsPerPixelX, Screen.Height / Screen.TwipsPerPixelY, DeskTopDC, 0&, 0&, Screen.Width / Screen.TwipsPerPixelX, Screen.Height / Screen.TwipsPerPixelY, SRCCOPY)
    
    ' Allow the image to copy.
    DoEvents
    ' Show frmRefresh.
    frmRefresh.Show
    ' Allow it to show.
    DoEvents
    ' Hide frmRefresh.
    frmRefresh.Hide
    ' Unload frmRefresh.
    Unload frmRefresh
    ' Allow it to unload and process other stuff
    '   happening in Windows.
    DoEvents
End Sub
