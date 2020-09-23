Attribute VB_Name = "basMinMaxSize"
Option Explicit

Private Const GWL_WNDPROC = -4
Private Const WM_GETMINMAXINFO = &H24

Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Type MINMAXINFO
    ptReserved As POINTAPI
    ptMaxSize As POINTAPI
    ptMaxPosition As POINTAPI
    ptMinTrackSize As POINTAPI
    ptMaxTrackSize As POINTAPI
End Type

Global lpPrevWndProc As Long
Global gHW As Long

Private Declare Function DefWindowProc Lib "user32" Alias _
   "DefWindowProcA" (ByVal hWnd As Long, ByVal wMsg As Long, _
    ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function CallWindowProc Lib "user32" Alias _
   "CallWindowProcA" (ByVal lpPrevWndFunc As Long, _
    ByVal hWnd As Long, ByVal Msg As Long, _
    ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias _
   "SetWindowLongA" (ByVal hWnd As Long, _
    ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Sub CopyMemoryToMinMaxInfo Lib "KERNEL32" Alias _
   "RtlMoveMemory" (hpvDest As MINMAXINFO, ByVal hpvSource As Long, _
    ByVal cbCopy As Long)
Private Declare Sub CopyMemoryFromMinMaxInfo Lib "KERNEL32" Alias _
   "RtlMoveMemory" (ByVal hpvDest As Long, hpvSource As MINMAXINFO, _
    ByVal cbCopy As Long)

Public Sub Hook()
    'Start subclassing.
    lpPrevWndProc = SetWindowLong(gHW, GWL_WNDPROC, _
       AddressOf WindowProc)
End Sub

Public Sub Unhook()
    Dim temp As Long

    'Cease subclassing.
    temp = SetWindowLong(gHW, GWL_WNDPROC, lpPrevWndProc)
End Sub

Function WindowProc(ByVal hw As Long, ByVal uMsg As Long, _
   ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim MinMax As MINMAXINFO

    'Check for request for min/max window sizes.
    If uMsg = WM_GETMINMAXINFO Then
        'Retrieve default MinMax settings
        CopyMemoryToMinMaxInfo MinMax, lParam, Len(MinMax)

        'Specify new minimum size for window.
        MinMax.ptMinTrackSize.x = 6840 / Screen.TwipsPerPixelX
        MinMax.ptMinTrackSize.y = 7890 / Screen.TwipsPerPixelY

        'Specify new maximum size for window.
        MinMax.ptMaxTrackSize.x = Screen.Width / Screen.TwipsPerPixelX
        MinMax.ptMaxTrackSize.y = Screen.Height / Screen.TwipsPerPixelY

        'Copy local structure back.
        CopyMemoryFromMinMaxInfo lParam, MinMax, Len(MinMax)

        WindowProc = DefWindowProc(hw, uMsg, wParam, lParam)
    Else
        WindowProc = CallWindowProc(lpPrevWndProc, hw, uMsg, _
           wParam, lParam)
    End If
End Function

