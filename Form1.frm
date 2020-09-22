VERSION 5.00
Begin VB.Form frmcle 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "HackDaWin"
   ClientHeight    =   450
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   1455
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   450
   ScaleWidth      =   1455
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkshow 
      Caption         =   "show"
      Height          =   195
      Left            =   600
      TabIndex        =   1
      Top             =   240
      Width           =   975
   End
   Begin VB.CheckBox chkenbl 
      Caption         =   "enable"
      Height          =   255
      Left            =   600
      TabIndex        =   0
      Top             =   0
      Value           =   1  'Checked
      Width           =   855
   End
End
Attribute VB_Name = "frmcle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function IsWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetSystemMenu Lib "user32" (ByVal hWnd As Long, ByVal bRevert As Long) As Long
Private Declare Function GetMenuItemInfo Lib "user32" Alias "GetMenuItemInfoA" (ByVal hMenu As Long, ByVal un As Long, ByVal b As Boolean, lpmii As MENUITEMINFO) As Long
Private Declare Function SetMenuItemInfo Lib "user32" Alias "SetMenuItemInfoA" (ByVal hMenu As Long, ByVal uItem As Long, ByVal fByPosition As Long, lpmii As MENUITEMINFO) As Long
Private Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Private Declare Function GetMenuString Lib "user32" Alias "GetMenuStringA" (ByVal hMenu As Long, ByVal wIDItem As Long, ByVal lpString As String, ByVal nMaxCount As Long, ByVal wFlag As Long) As Long
Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long
Private Type MENUITEMINFO
  cbSize As Long
  fMask As Long
  fType As Long
  fState As Long
  wID As Long
  hSubMenu As Long
  hbmpChecked As Long
  hbmpUnchecked As Long
  dwItemData As Long
  dwTypeData As String
  cch As Long
End Type

Private Const MIIM_STATE = &H1
Private Const SC_CLOSE = &HF060
Private Const MIIM_ID = &H2
Private Const MFS_GRAYED = &H3&
Private Const WM_NCACTIVATE = &H86

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.MousePointer = vbCrosshair
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim bWnd As Long
    ' Dim sSave As String * 250
    Dim i As Long
    Dim j As Long
    Dim ret As Long
    Dim ptapp As Toolbar
    Dim kaleman As Long
    Dim kaleman2 As Long
    Dim pt As POINTAPI, mWnd As Long ', nDC As Long
    Dim MyStr As String
    Dim hMenu As Long, hSubMenu As Long
    Dim cpt As Long
    Dim ave As Long
    Dim nrpos As Long
    Dim pos As Long
    Dim stete As String
    Dim Counter As Long, xxx As Long, Par As Long
    Const Clase_Name As String = "ThunderTextBox"
    Const Clase_Name2 As String = "Edit"
    MousePointer = vbHourglass
    GetCursorPos pt
    
    mWnd = WindowFromPoint(pt.X, pt.Y)
    If chkenbl.Value = vbChecked Then Call EnableWindow(mWnd, 1)

    For j = 0 To 50000
        If GetWindowTextLength(j) > 0 Then
            Call EnableWindow(j, 1)
        End If
    Next
    
    hMenu = GetMenu(mWnd)

    If hMenu <> 0 Then
'      menu trouvé
        For j = 0 To 1000
            ret = EnableMenuItem(hMenu, j, MF_BYPOSITION + MF_ENABLED)
            stete = Space$(50)
            ave = GetMenuString(hMenu, j, stete, 50, MF_BYPOSITION)
            If ave > 0 Then Debug.Print stete
            DoEvents
            hSubMenu = GetSubMenu(hMenu, j)
            If hSubMenu <> 0 Then
                For i = 0 To 1000 'GetMenuItemCount(hSubMenu) + 1
                    ret = EnableMenuItem(hSubMenu, i, MF_BYPOSITION + MF_ENABLED + MF_CHANGE)
                    stete = Space$(50)
                    ave = GetMenuString(hSubMenu, i, stete, 50, MF_BYPOSITION)
                    If ave > 0 Then Debug.Print stete
                    
                    DoEvents
                Next i
            End If
            
        Next j
    End If
    
    bWnd = GetWindow(mWnd, GW_CHILD)
    
    For i = 1 To 1000
'        GetClassName bWnd, sSave, 250
'            If InStr(1, sSave, "Toolbar") <> 0 Then
'               toolbar trouvé
'            End If

           If chkshow.Value = vbChecked Then Call ShowWindow(bWnd, 1)
           If chkenbl.Value = vbChecked Then Call EnableWindow(bWnd, 1)
          
            
        bWnd = GetWindow(bWnd, GW_HWNDNEXT)
    Next
    
MousePointer = vbHourglass
    findpass
Me.MousePointer = vbCrosshair
End Sub



