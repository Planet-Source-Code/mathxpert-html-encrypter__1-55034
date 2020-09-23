Attribute VB_Name = "basAPI"
Public Const TB = "        "

Public Const EM_GETSEL = &HB0
Public Const EM_SETSEL = &HB1
Public Const EM_GETLINECOUNT = &HBA
Public Const EM_LINEINDEX = &HBB
Public Const EM_SCROLLCARET = &HB7
Public Const EM_LINELENGTH = &HC1
Public Const EM_CANUNDO = &HC6
Public Const EM_LINEFROMCHAR = &HC9
Public Const EM_EMPTYUNDOBUFFER = &HCD

Public Const WM_CUT = &H300
Public Const WM_COPY = &H301
Public Const WM_PASTE = &H302
Public Const WM_CLEAR = &H303
Public Const WM_UNDO = &H304

Public Declare Sub ReleaseCapture Lib "user32" ()
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Public origF As Boolean, compF As Boolean, tmpPos As Long, tmpText As String
