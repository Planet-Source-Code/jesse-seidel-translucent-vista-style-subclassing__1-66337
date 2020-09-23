VERSION 5.00
Begin VB.Form frmBorder 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   2865
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   2865
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Image Image11 
      Height          =   105
      Left            =   360
      Picture         =   "frmBorder.frx":0000
      Top             =   2520
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.Image Image10 
      Height          =   105
      Left            =   1560
      Picture         =   "frmBorder.frx":067E
      Stretch         =   -1  'True
      Top             =   2520
      Visible         =   0   'False
      Width           =   3465
   End
   Begin VB.Image Image9 
      Height          =   75
      Left            =   6360
      Picture         =   "frmBorder.frx":0768
      Stretch         =   -1  'True
      Top             =   3840
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Image Image8 
      Height          =   120
      Left            =   6360
      Picture         =   "frmBorder.frx":392A
      Stretch         =   -1  'True
      Top             =   3960
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Image Image7 
      Height          =   1470
      Left            =   6360
      Picture         =   "frmBorder.frx":3A2C
      Stretch         =   -1  'True
      Top             =   2280
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Image Image6 
      Height          =   1470
      Left            =   120
      Picture         =   "frmBorder.frx":439E
      Stretch         =   -1  'True
      Top             =   840
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Image Image5 
      Height          =   120
      Left            =   120
      Picture         =   "frmBorder.frx":4D10
      Stretch         =   -1  'True
      Top             =   2520
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Image Image4 
      Height          =   75
      Left            =   120
      Picture         =   "frmBorder.frx":4E12
      Stretch         =   -1  'True
      Top             =   2400
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Image Image1 
      Height          =   375
      Left            =   120
      Picture         =   "frmBorder.frx":7FD4
      Top             =   360
      Visible         =   0   'False
      Width           =   3435
   End
   Begin VB.Image Image2 
      Height          =   375
      Left            =   3600
      Picture         =   "frmBorder.frx":C346
      Stretch         =   -1  'True
      Top             =   360
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Image Image3 
      Height          =   375
      Left            =   3720
      Picture         =   "frmBorder.frx":1C7B0
      Top             =   360
      Visible         =   0   'False
      Width           =   1500
   End
End
Attribute VB_Name = "frmBorder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long

Private Const HTTOPRIGHT = 14
Private Const HTTOPLEFT = 13
Private Const HTTOP = 12
Private Const HTBOTTOM = 15
Private Const HTBOTTOMLEFT = 16
Private Const HTBOTTOMRIGHT = 17
Private Const HTLEFT = 10
Private Const HTRIGHT = 11

Private Const WM_NCLBUTTONDOWN = &HA1

Dim NotepadLoaded As Long

Private Sub Form_DblClick()

  RestoreBorder SkinWindow
  
  Unload Me
  Unload frmMain

End Sub

Private Sub Form_GotFocus()

  SetTopMost SkinWindow, True
  SetActiveWindow SkinWindow
  
End Sub

Private Sub Form_Load()
    
    NotepadLoaded = Shell("notepad", vbNormalFocus)
    
    Do Until NotepadLoaded > 0
      DoEvents
    Loop
    
    SkinWindow = FindWindow("notepad", vbNullString)
    
    Init
    
    RemoveBorder SkinWindow
    
    SetActiveWindow SkinWindow
    'SetTopMost hwnd, True
    SetTopMost SkinWindow, True
    
    GetWindowRect SkinWindow, r
    
    Top = (r.Top * Screen.TwipsPerPixelY) - Image1.Height
    Left = (r.Left * Screen.TwipsPerPixelX) - (Image6.Width)
    Height = (r.Bottom * Screen.TwipsPerPixelY) - (Top) + Image10.Height
    Width = (r.Right * Screen.TwipsPerPixelX) - (Left) + Image6.Width
    
    'frmMain.Show , Me
    
    SetTranslucent hwnd, 35
    
    pOldWindPoc = SetWindowLong(Me.hwnd, GWL_WNDPROC, AddressOf WndProc)
    
    BackColor = frmMain.BackColor
    
    RepaintWindow

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    SetTopMost SkinWindow, True
    
    If y <= Image1.Height Then
        ReleaseCapture
        SendMessage Me.hwnd, &H112, &HF012&, 0
    ElseIf y >= Height - Image8.Height And x >= Width - Image8.Width Then
        ReleaseCapture
        SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTBOTTOMRIGHT, ByVal 0&
    ElseIf y <= Image8.Height And x >= Width - Image8.Width Then
        ReleaseCapture
        SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTTOPRIGHT, ByVal 0&
    End If
    
    RepaintWindow
    
    SetTopMost SkinWindow, False

End Sub

Public Sub ResizeWin()

    RepaintWindow
    
    frmMain.Top = Top + Image1.Height
    frmMain.Left = Left + Image7.Width
    frmMain.Width = Width - (Image7.Width * 2)
    frmMain.Height = Height - Image1.Height - Image11.Height

End Sub

Private Sub Form_Unload(Cancel As Integer)

  SetWindowLong Me.hwnd, GWL_WNDPROC, pOldWindPoc
  SetTopMost SkinWindow, False
  
End Sub

Public Sub JustMoved()

  frmMain.Top = Top + Image1.Height
  frmMain.Left = Left + Image7.Width
  frmMain.Width = Width - (Image7.Width * 2)
  frmMain.Height = Height - Image1.Height - Image11.Height

End Sub

Public Sub RepaintWindow()

    On Error Resume Next
    
    PaintPicture Image1, 0, 0
    PaintPicture Image2, Image1.Width, 0, Width - Image1.Width - Image3.Width
    PaintPicture Image3, Width - Image3.Width, 0
    
    PaintPicture Image6, 0, Image1.Height
    PaintPicture Image4, 0, Image1.Height + Image6.Height, , Height - Image5.Height - Image1.Height - Image6.Height
    PaintPicture Image5, 0, Height - Image5.Height
    
    PaintPicture Image11, Image5.Width, Height - Image11.Height
    PaintPicture Image10, Image11.Width + Image5.Width, Height - Image10.Height, Width - Image11.Width - Image5.Width - Image8.Width
    
    PaintPicture Image9, Width - Image9.Width, Image1.Height + Image7.Height, , Height - Image1.Height - Image7.Height - Image8.Height
    PaintPicture Image7, Width - Image7.Width, Image1.Height
    PaintPicture Image8, Width - Image8.Width, Height - Image8.Height
    
    Refresh
    
End Sub
