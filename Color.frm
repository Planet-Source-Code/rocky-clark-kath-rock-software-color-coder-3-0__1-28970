VERSION 5.00
Begin VB.Form frmColor 
   BorderStyle     =   0  'None
   ClientHeight    =   1575
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1425
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   15  'Size All
   ScaleHeight     =   1575
   ScaleWidth      =   1425
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmColor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type RECT
    Left    As Long
    Top     As Long
    Right   As Long
    Bottom  As Long
End Type

Private Const WM_NCLBUTTONDOWN = &HA1
Private Const HTCAPTION = 2

Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Sub Form_Activate()

    Call ReleaseCapture
    Call SendMessage(Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0)
    Unload Me
    
End Sub

Private Sub Form_Load()

Dim rcBox   As RECT
Dim lLeft   As Long
Dim lTop    As Long
Dim lWidth  As Long
Dim lHeight As Long

    Call GetWindowRect(frmMain.fraColor.hWnd, rcBox)
    lLeft = (rcBox.Left * Screen.TwipsPerPixelX) _
        + frmMain.lblColorBox.Left + 30
    lTop = (rcBox.Top * Screen.TwipsPerPixelY) _
        + frmMain.lblColorBox.Top + 30
    lWidth = frmMain.lblColorBox.Width - 60
    lHeight = frmMain.lblColorBox.Height - 60

    Me.Move lLeft, lTop, lWidth, lHeight
    Me.BackColor = frmMain.lblColorBox.BackColor
    
End Sub


