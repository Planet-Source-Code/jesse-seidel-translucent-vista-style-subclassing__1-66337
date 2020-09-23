Attribute VB_Name = "modStick"
Option Explicit

Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" _
                          (ByVal lpPrevWndFunc As Long, _
                           ByVal hwnd As Long, _
                           ByVal msg As Long, _
                           ByVal wParam As Long, _
                           ByVal lParam As Long) As Long

Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32.dll" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Public Declare Function GetActiveWindow Lib "user32" () As Long
 
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long

Public Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type
  
Public Declare Function SetActiveWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function CloseWindow Lib "user32" (ByVal hwnd As Long) As Long

Private Const SWP_FRAMECHANGED = &H20
Private Const SW_SHOW = 5
  
Private Const GWL_EXSTYLE          As Long = (-20)
Private Const GWL_STYLE            As Long = (-16)
Private Const WS_EX_LAYERED        As Long = &H80000
Private Const LWA_ALPHA            As Long = &H2

Private Const WS_CLIPSIBLINGS      As Long = &H4000000
Private Const WS_EX_WINDOWEDGE     As Long = &H100&
Private Const WS_THICKFRAME        As Long = &H40000
Private Const WS_VISIBLE           As Long = &H10000000
Private Const WS_DLGFRAME          As Long = &H400000
Private Const WS_CAPTION           As Long = &HC00000
Private Const WS_BORDER            As Long = &H800000
Private Const SWP_NOACTIVATE = &H10
Private Const SWP_SHOWWINDOW = &H40
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1
Public Const WM_SETFOCUS = &H7
Public Const WM_KILLFOCUS = &H8

Private Const HWND_NOTOPMOST = -2
Private Const HWND_TOPMOST = -1

Private Const WM_ACTIVATE = &H6

Public pOldWindPoc As Long

Private Const WM_MOVE As Long = &H3
Private Const WM_SIZE = &H5
Private Const WM_NCLBUTTONDOWN = &HA1

Public Const GWL_WNDPROC& = (-4)

Public r As RECT, SkinWindow As Long, ActiveWindow As Boolean, sOldStyle As Long, sOldStyleEx As Long

Public Sub Init()

  ActiveWindow = False
  sOldStyle = GetWindowLong(SkinWindow, GWL_STYLE)
  sOldStyleEx = GetWindowLong(SkinWindow, GWL_EXSTYLE)

End Sub

Public Function SetTopMost(wHandle As Long, Optional bTopMost As Boolean = True)

    Dim wFlags As Long, Placement As Long
    
    wFlags = SWP_NOMOVE Or SWP_NOSIZE Or SWP_SHOWWINDOW Or SWP_NOACTIVATE

    Select Case bTopMost
    Case True
        Placement = HWND_TOPMOST
    Case False
        Placement = HWND_NOTOPMOST
    End Select
    
    SetWindowPos wHandle, Placement, 0, 0, 0, 0, wFlags

End Function

Public Function WndProc(ByVal hwnd As Long, _
     ByVal uMsg As Long, _
     ByVal wParam As Long, _
     ByVal lParam As Long) As Long

  Dim FormCaption As String
    
  If uMsg = WM_MOVE Then
    
    SetTopMost SkinWindow, True
    'SetActiveWindow SkinWindow
    
    GetWindowRect frmBorder.hwnd, r
    SetWindowPos SkinWindow, 0, (frmBorder.Left + frmBorder.Image6.Width) / Screen.TwipsPerPixelX, _
    (frmBorder.Top + frmBorder.Image1.Height) / Screen.TwipsPerPixelY, _
    r.Right - (frmBorder.Left / Screen.TwipsPerPixelX) - ((frmBorder.Image6.Width * 2) / Screen.TwipsPerPixelX), _
    r.Bottom - (frmBorder.Top / Screen.TwipsPerPixelY) - ((frmBorder.Image1.Height + frmBorder.Image10.Height) / Screen.TwipsPerPixelY), SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_FRAMECHANGED
    
    'frmBorder.RepaintWindow
    
    'SetTopMost SkinWindow, False
    'SetActiveWindow SkinWindow
  
  ElseIf uMsg = WM_KILLFOCUS Then
    'CloseWindow SkinWindow
    'GetWindowLong hwnd, GWL_EXSTYLE
    'SetWindowLong SkinWindow, GWL_STYLE, (WM_KILLFOCUS)
    SetTopMost SkinWindow, False
  
  ElseIf uMsg = WM_SIZE Then
    
    SetTopMost SkinWindow, True
    'SetActiveWindow SkinWindow
    
    GetWindowRect frmBorder.hwnd, r
    SetWindowPos SkinWindow, 0, (frmBorder.Left + frmBorder.Image6.Width) / Screen.TwipsPerPixelX, _
    (frmBorder.Top + frmBorder.Image1.Height) / Screen.TwipsPerPixelY, _
    r.Right - (frmBorder.Left / Screen.TwipsPerPixelX) - ((frmBorder.Image6.Width * 2) / Screen.TwipsPerPixelX), _
    r.Bottom - (frmBorder.Top / Screen.TwipsPerPixelY) - ((frmBorder.Image1.Height + frmBorder.Image10.Height) / Screen.TwipsPerPixelY), SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_FRAMECHANGED
    
    frmBorder.RepaintWindow
    
  Else
    WndProc = CallWindowProc(pOldWindPoc, hwnd, uMsg, wParam, lParam)
  End If

End Function

Public Sub SetTranslucent(hwnd As Long, Level As Long)

Dim NormalWindowStyle As Long

    On Error Resume Next
    DoEvents

        NormalWindowStyle = GetWindowLong(hwnd, GWL_EXSTYLE)
        SetWindowLong hwnd, GWL_EXSTYLE, WS_EX_LAYERED
        SetLayeredWindowAttributes hwnd, 0, 255 * (1 - (Val(Level) / 100)), LWA_ALPHA

    On Error GoTo 0
End Sub

Public Sub RemoveBorder(hwnd As Long)

  Dim NormalWindowStyle As Long
  
  'On Error Resume Next
    DoEvents

        NormalWindowStyle = GetWindowLong(hwnd, GWL_EXSTYLE)
        SetWindowLong hwnd, GWL_STYLE, (WS_VISIBLE Or WS_CLIPSIBLINGS Or SWP_FRAMECHANGED)

    On Error GoTo 0

End Sub

Public Sub RestoreBorder(hwnd As Long)

  Dim NormalWindowStyle As Long
  
  'On Error Resume Next
    DoEvents

        NormalWindowStyle = GetWindowLong(hwnd, GWL_EXSTYLE)
        SetWindowLong hwnd, GWL_STYLE, sOldStyle
        SetWindowLong hwnd, GWL_EXSTYLE, sOldStyleEx
        
        SetWindowPos hwnd, 0, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_FRAMECHANGED

    On Error GoTo 0
    
End Sub
