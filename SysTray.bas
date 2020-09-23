Attribute VB_Name = "basSysTray"
Option Explicit

Private Type PointAPI
    X   As Long
    Y   As Long
End Type

Private Type NotifyIconData
    cbSize          As Long
    hWnd            As Long
    uID             As Long
    uFlags          As Long
    uCallBackMsg    As Long
    hIcon           As Long
    szTip           As String * 64
End Type

Private Const STI_ADD = 0
Private Const STI_MODIFY = 1
Private Const STI_DELETE = 2
Private Const STI_MESSAGE = 1
Private Const STI_ICON = 2
Private Const STI_TIP = 4
Public Const WM_MOUSEMOVE = &H200
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_RBUTTONDOWN = &H204
Public Const WM_RBUTTONUP = &H205
Public Const WM_RBUTTONDBLCLK = &H206

Public Enum SysTrayAction
    AddIcon = STI_ADD
    ModifyIcon = STI_MODIFY
    DeleteIcon = STI_DELETE
End Enum

Private guIconData As NotifyIconData

Private Declare Function SysTrayIcon Lib "SHELL32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NotifyIconData) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As PointAPI) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, lpPoint As PointAPI) As Long

Public Sub SetSysTrayIcon(eAction As SysTrayAction, hMsgWnd As Long, hIcon As Long, sToolTip As String)

'Example:
'   Call SetSysTrayIcon(ModifyIcon, chkHiddenCheckBox.hWnd, Me.Icon, "My Program Name")
'   chkHiddenCheckBox_MouseMove() event will receive all messages from the System Tray Icon.
'   chkHiddenCheckBox could be any control with a MouseMove event.

Dim lRet As Long

    With guIconData
        .cbSize = Len(guIconData)
        .hWnd = hMsgWnd
        .uID = vbNull
        .uFlags = STI_MESSAGE Or STI_ICON Or STI_TIP
        .uCallBackMsg = WM_MOUSEMOVE
        .hIcon = hIcon
        .szTip = sToolTip & Chr$(0)
    End With
    
    lRet = SysTrayIcon(eAction, guIconData)
    
End Sub

Public Sub ShowPopupAtCursor(frmCalling As Form, mnuPopupMenu As Menu, Optional eAlignment As MenuControlConstants = 0, Optional mnuDefaultItem As Menu)

'Example:
'ShowPopupAtCursor Me, mnuEdit, _
'   vbPopupMenuRightButton + vbPopupMenuRightAlign _
'   mnuEdit_Copy

Dim lRet As Long
Dim uCursorPt As PointAPI

    'Get Cursor Position on Screen.
    lRet = GetCursorPos(uCursorPt)
    
    'Convert cursor position to coordinates relative to form.
    Call ScreenToClient(frmCalling.hWnd, uCursorPt)
    
    'Convert to Twips
    uCursorPt.X = uCursorPt.X * Screen.TwipsPerPixelX
    uCursorPt.Y = uCursorPt.Y * Screen.TwipsPerPixelY
    
    'Show the Menu wherever the cursor is on the Screen.
    If mnuDefaultItem Is Nothing Then
        frmCalling.PopupMenu mnuPopupMenu, eAlignment, uCursorPt.X, uCursorPt.Y
    Else
        frmCalling.PopupMenu mnuPopupMenu, eAlignment, uCursorPt.X, uCursorPt.Y, mnuDefaultItem
    End If
    
End Sub


