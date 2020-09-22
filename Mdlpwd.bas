Attribute VB_Name = "Module1"
Public Declare Function GetDesktopWindow Lib "user32" () As Long
Public Declare Function EnumChildWindows Lib "user32" (ByVal hWndParent As Long, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Const EM_SETPASSWORDCHAR = &HCC
'pour les menus
Public Const MF_ENABLED = &H0&
Public Const MF_BYCOMMAND = &H0&
Public Const MF_DISABLED = &H2&
Public Const MF_GRAYED = &H1&
Public Const MF_BYPOSITION = &H400&
Public Const MF_CHANGE = &H80&
'pour les fenêtre
Public Const GW_CHILD = 5
Public Const GW_HWNDNEXT = 2

'types
Public Type POINTAPI
    X As Long
    Y As Long
End Type

'API
Public Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal LpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
'public Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Public Declare Function EnableWindow Lib "user32" (ByVal hWnd As Long, ByVal fEnable As Long) As Long

Public Declare Function GetMenu Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
'public Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long

Public Declare Function EnableMenuItem Lib "user32" (ByVal hMenu As Long, ByVal wIDEnableItem As Long, ByVal wEnable As Long) As Long

Public Declare Function GetCursorPos Lib "user32" (X As POINTAPI) As Long
Public Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long


Declare Function EnumWindows Lib "user32" ( _
        ByVal lpEnumFunc As Long, _
        ByVal lParam As Long) As Long



Private Const EM_GETPASSWORDCHAR = &HD2

Private Const EM_SETMODIFY = &HB9
Private Const SW_HIDE = 0
Private Const SW_SHOW = 5


Private Function EnumWindowsProc(ByVal hWnd As Long, ByVal lParam As Long) As Long
  EnumChildWindows hWnd, AddressOf EnumWindowsProc2, 1
  EnumWindowsProc = True
End Function

Private Function EnumWindowsProc2(ByVal hWnd As Long, ByVal lParam As Long) As Long
  If SendMessage(hWnd, EM_GETPASSWORDCHAR, 0, 1) Then
   UpdateWindow hWnd
  End If
  EnumWindowsProc2 = True
End Function

Private Sub UpdateWindow(hWnd As Long)
  SendMessage hWnd, EM_SETPASSWORDCHAR, 0, 1
  SendMessage hWnd, EM_SETMODIFY, True, 1
  ShowWindow hWnd, SW_HIDE
  ShowWindow hWnd, SW_SHOW
End Sub

Public Function findpass()
  EnumWindows AddressOf EnumWindowsProc, 1
End Function




