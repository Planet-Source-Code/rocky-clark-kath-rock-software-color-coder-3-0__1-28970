VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Color Coder"
   ClientHeight    =   5085
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5280
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MouseIcon       =   "Main.frx":0442
   ScaleHeight     =   5085
   ScaleWidth      =   5280
   Begin VB.CheckBox chkCapture 
      Caption         =   "&Capture Color"
      Height          =   375
      Left            =   2940
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   3285
      Width           =   1185
   End
   Begin VB.CommandButton cmdAbout 
      Caption         =   "&About"
      Height          =   375
      Left            =   4230
      TabIndex        =   24
      Top             =   3285
      Width           =   885
   End
   Begin VB.TextBox txtColor 
      Alignment       =   2  'Center
      BackColor       =   &H80000014&
      Height          =   285
      Index           =   5
      Left            =   1785
      Locked          =   -1  'True
      TabIndex        =   37
      Text            =   "&H00FFFF&"
      Top             =   4665
      Width           =   1605
   End
   Begin VB.TextBox txtColor 
      Alignment       =   2  'Center
      BackColor       =   &H80000014&
      Height          =   285
      Index           =   4
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   35
      Text            =   "#FFFF00"
      Top             =   4650
      Width           =   1605
   End
   Begin VB.TextBox txtColor 
      Alignment       =   2  'Center
      BackColor       =   &H80000014&
      Height          =   285
      Index           =   3
      Left            =   3450
      Locked          =   -1  'True
      TabIndex        =   33
      Text            =   "16777215"
      Top             =   4050
      Width           =   1665
   End
   Begin VB.TextBox txtColor 
      Alignment       =   2  'Center
      BackColor       =   &H80000014&
      Height          =   285
      Index           =   2
      Left            =   1785
      Locked          =   -1  'True
      TabIndex        =   31
      Text            =   "RGB(255, 255, 255)"
      Top             =   4050
      Width           =   1605
   End
   Begin VB.TextBox txtColor 
      Alignment       =   2  'Center
      BackColor       =   &H80000014&
      Height          =   285
      Index           =   1
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   29
      Text            =   "QBColor(15)"
      Top             =   4050
      Width           =   1605
   End
   Begin VB.CheckBox chkMsg 
      Height          =   195
      Left            =   1410
      TabIndex        =   39
      Top             =   3825
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton cmdHide 
      Caption         =   "&Hide"
      Height          =   375
      Left            =   3450
      TabIndex        =   38
      Top             =   4575
      Width           =   1665
   End
   Begin VB.Frame fraSysColors 
      Caption         =   "System Colors"
      Height          =   1275
      Left            =   2940
      TabIndex        =   20
      Top             =   1905
      Width           =   2175
      Begin VB.TextBox txtColor 
         Alignment       =   2  'Center
         BackColor       =   &H80000014&
         Height          =   285
         Index           =   7
         Left            =   60
         Locked          =   -1  'True
         TabIndex        =   27
         Text            =   "&H8000000F&"
         Top             =   915
         Width           =   2040
      End
      Begin VB.TextBox txtColor 
         Alignment       =   2  'Center
         BackColor       =   &H80000014&
         Height          =   285
         Index           =   6
         Left            =   60
         Locked          =   -1  'True
         TabIndex        =   26
         Text            =   "vbButtonFace"
         Top             =   585
         Width           =   2040
      End
      Begin VB.ComboBox cboSysColors 
         Height          =   315
         Left            =   60
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   225
         Width           =   2040
      End
   End
   Begin VB.Frame fraShade 
      Caption         =   "Darker"
      Height          =   495
      Left            =   1545
      TabIndex        =   18
      Top             =   1350
      Width           =   3570
      Begin VB.HScrollBar hsbShade 
         Height          =   210
         LargeChange     =   16
         Left            =   75
         Max             =   255
         TabIndex        =   17
         Top             =   210
         Width           =   3390
      End
      Begin VB.Label lblLighter 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Lighter"
         Height          =   195
         Left            =   2910
         TabIndex        =   16
         Top             =   -15
         Width           =   480
      End
   End
   Begin VB.Frame fraColor 
      Caption         =   "Selected Color"
      Height          =   1770
      Left            =   1545
      TabIndex        =   22
      Top             =   1905
      Width           =   1305
      Begin VB.Label lblColorBox 
         BorderStyle     =   1  'Fixed Single
         Height          =   1440
         Left            =   75
         MousePointer    =   15  'Size All
         TabIndex        =   21
         ToolTipText     =   "Drag to test color match"
         Top             =   240
         Width           =   1140
      End
   End
   Begin VB.Frame fraRGBColors 
      Caption         =   "RGB Colors"
      Height          =   1185
      Left            =   1545
      TabIndex        =   15
      Top             =   120
      Width           =   3570
      Begin VB.CheckBox chkHex 
         Caption         =   "Hex"
         Height          =   195
         Left            =   2850
         TabIndex        =   11
         Top             =   0
         Width           =   600
      End
      Begin VB.HScrollBar hsbRGB 
         Height          =   210
         Index           =   2
         LargeChange     =   16
         Left            =   615
         Max             =   255
         TabIndex        =   10
         Top             =   855
         Width           =   2505
      End
      Begin VB.HScrollBar hsbRGB 
         Height          =   210
         Index           =   1
         LargeChange     =   16
         Left            =   615
         Max             =   255
         TabIndex        =   9
         Top             =   555
         Width           =   2505
      End
      Begin VB.HScrollBar hsbRGB 
         Height          =   210
         Index           =   0
         LargeChange     =   16
         Left            =   615
         Max             =   255
         TabIndex        =   8
         Top             =   255
         Width           =   2505
      End
      Begin VB.TextBox txtRGB 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         Height          =   210
         Index           =   0
         Left            =   3120
         MaxLength       =   3
         TabIndex        =   12
         Top             =   255
         Width           =   315
      End
      Begin VB.TextBox txtRGB 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         Height          =   210
         Index           =   1
         Left            =   3120
         MaxLength       =   3
         TabIndex        =   13
         Top             =   555
         Width           =   315
      End
      Begin VB.TextBox txtRGB 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         Height          =   210
         Index           =   2
         Left            =   3120
         MaxLength       =   3
         TabIndex        =   14
         Top             =   855
         Width           =   315
      End
      Begin VB.Label lblRGB 
         Alignment       =   2  'Center
         BackColor       =   &H00C00000&
         Caption         =   "Blue"
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   2
         Left            =   120
         TabIndex        =   7
         Top             =   855
         Width           =   495
      End
      Begin VB.Label lblRGB 
         Alignment       =   2  'Center
         BackColor       =   &H00008000&
         Caption         =   "Green"
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   1
         Left            =   120
         TabIndex        =   6
         Top             =   555
         Width           =   495
      End
      Begin VB.Label lblRGB 
         Alignment       =   2  'Center
         BackColor       =   &H000000C0&
         Caption         =   "Red"
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   0
         Left            =   120
         TabIndex        =   5
         Top             =   255
         Width           =   495
      End
      Begin VB.Label lblBorder 
         BorderStyle     =   1  'Fixed Single
         Height          =   270
         Index           =   0
         Left            =   90
         TabIndex        =   2
         Top             =   225
         Width           =   3375
      End
      Begin VB.Label lblBorder 
         BorderStyle     =   1  'Fixed Single
         Height          =   270
         Index           =   1
         Left            =   90
         TabIndex        =   3
         Top             =   525
         Width           =   3375
      End
      Begin VB.Label lblBorder 
         BorderStyle     =   1  'Fixed Single
         Height          =   270
         Index           =   2
         Left            =   90
         TabIndex        =   4
         Top             =   825
         Width           =   3375
      End
   End
   Begin VB.Frame fraQBColors 
      Caption         =   "QB Colors"
      Height          =   3540
      Left            =   120
      TabIndex        =   1
      Top             =   135
      Width           =   1320
      Begin VB.TextBox txtColor 
         Alignment       =   2  'Center
         BackColor       =   &H80000014&
         Height          =   285
         Index           =   0
         Left            =   75
         Locked          =   -1  'True
         TabIndex        =   25
         Text            =   "vbConst"
         Top             =   3195
         Width           =   1155
      End
      Begin VB.PictureBox picQBColors 
         AutoRedraw      =   -1  'True
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2940
         Left            =   90
         ScaleHeight     =   192
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   72
         TabIndex        =   0
         Top             =   225
         Width           =   1140
      End
   End
   Begin VB.Label lblCap 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "VB Hex Color Code:"
      Height          =   195
      Index           =   4
      Left            =   1800
      TabIndex        =   36
      Top             =   4455
      Width           =   1410
   End
   Begin VB.Label lblCap 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "HTML Color Code:"
      Height          =   195
      Index           =   3
      Left            =   135
      TabIndex        =   34
      Top             =   4455
      Width           =   1320
   End
   Begin VB.Label lblCap 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Decimal Color Code:"
      Height          =   195
      Index           =   2
      Left            =   3465
      TabIndex        =   32
      Top             =   3855
      Width           =   1440
   End
   Begin VB.Label lblCap 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "RGB Color Code:"
      Height          =   195
      Index           =   1
      Left            =   1800
      TabIndex        =   30
      Top             =   3825
      Width           =   1215
   End
   Begin VB.Label lblCap 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "QBColor Code:"
      Height          =   195
      Index           =   0
      Left            =   135
      TabIndex        =   28
      Top             =   3840
      Width           =   1050
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Visible         =   0   'False
      Begin VB.Menu mnuFile_Show 
         Caption         =   "&Show"
      End
      Begin VB.Menu mnuFile_Line10 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFile_Exit 
         Caption         =   "E&xit"
      End
      Begin VB.Menu mnuFile_Line20 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFile_Cancel 
         Caption         =   "&Cancel"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type PointAPI
    X   As Long
    Y   As Long
End Type

Private mbNoChange      As Boolean
Private mbCapture       As Boolean
Private miCurQBIdx      As Integer
Private mlCurColor      As Long
Private msaSysColors()  As String

Private Declare Function GetCursorPos Lib "user32" (lpPoint As PointAPI) As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SetCapture Lib "user32" (ByVal hWnd As Long) As Long

Private Function ConvertToSysColor(ByVal lColor As Long) As Long

'Find a system color that matches lColor

Dim lIdx As Long
Dim sHex As String

    If lColor < 0 Then
        'Already a system color
        ConvertToSysColor = lColor
    Else
        For lIdx = 0 To 24
            If GetSysColor(lIdx) = lColor Then
                'Found a match
                sHex = Hex$(lIdx)
                If Len(sHex) < 2 Then
                    sHex = "0" & sHex
                End If
                ConvertToSysColor = Val("&H800000" & sHex)
                Exit For
            End If
        Next
        If lIdx > 24 Then
            'Didn't find a match
            ConvertToSysColor = -1
        End If
    End If
    
End Function

Private Sub DrawButtonUp(iIdx As Integer)

'Draw the button in the "unpressed" position

Dim iX          As Integer
Dim iY          As Integer
Dim iBtnWidth   As Integer
Dim iBtnHeight  As Integer

    If iIdx >= 0 Then
        iBtnWidth = picQBColors.ScaleWidth / 2
        iBtnHeight = picQBColors.ScaleHeight / 8
        iX = Int(iIdx / 8) * iBtnWidth
        iY = (iIdx Mod 8) * iBtnHeight
        picQBColors.Line (iX, iY)-Step(iBtnWidth - 1, iBtnHeight - 1), vbButtonFace, BF
        picQBColors.Line (iX, iY)-Step(iBtnWidth - 1, iBtnHeight - 1), vb3DDKShadow, B
        picQBColors.Line (iX, iY)-Step(iBtnWidth - 2, iBtnHeight - 2), vb3DHighlight, B
        picQBColors.Line (iX + 1, iY + 1)-Step(iBtnWidth - 3, iBtnHeight - 3), vbButtonShadow, B
        picQBColors.Line (iX + 1, iY + 1)-Step(iBtnWidth - 4, iBtnHeight - 4), vbButtonFace, BF
        picQBColors.Line (iX + 4, iY + 4)-Step(iBtnWidth - 10, iBtnHeight - 10), QBColor(iIdx), BF
        picQBColors.Line (iX + 4, iY + 4)-Step(iBtnWidth - 10, iBtnHeight - 10), &H0&, B
        picQBColors.CurrentX = Int(((iX + (iX + iBtnWidth - 1)) / 2) - (picQBColors.TextWidth(CStr(iIdx)) / 2))
        picQBColors.CurrentY = Int(((iY + (iY + iBtnHeight - 1)) / 2) - (picQBColors.TextHeight(CStr(iIdx)) / 2))
        picQBColors.ForeColor = Abs(CInt(iIdx < 10)) * &HFFFFFF
        picQBColors.Print CStr(iIdx)
        miCurQBIdx = -1
    End If
    
End Sub

Private Sub DrawButtonDown(iIdx As Integer)

'Draw the button in the "pressed" position

Dim iX          As Integer
Dim iY          As Integer
Dim iBtnWidth   As Integer
Dim iBtnHeight  As Integer
Dim sConst      As String

    If iIdx >= 0 Then
        iBtnWidth = picQBColors.ScaleWidth / 2
        iBtnHeight = picQBColors.ScaleHeight / 8
        iX = Int(iIdx / 8) * iBtnWidth
        iY = (iIdx Mod 8) * iBtnHeight
        picQBColors.Line (iX, iY)-Step(iBtnWidth - 1, iBtnHeight - 1), vbButtonFace, BF
        picQBColors.Line (iX, iY)-Step(iBtnWidth - 1, iBtnHeight - 1), vb3DDKShadow, B
        picQBColors.Line (iX + 1, iY + 1)-Step(iBtnWidth - 3, iBtnHeight - 3), vb3DHighlight, B
        picQBColors.Line (iX + 1, iY + 1)-Step(iBtnWidth - 4, iBtnHeight - 4), vbButtonShadow, B
        picQBColors.Line (iX + 2, iY + 2)-Step(iBtnWidth - 5, iBtnHeight - 5), vbButtonFace, B
        picQBColors.Line (iX + 5, iY + 5)-Step(iBtnWidth - 10, iBtnHeight - 10), QBColor(iIdx), BF
        picQBColors.Line (iX + 5, iY + 5)-Step(iBtnWidth - 10, iBtnHeight - 10), &H0&, B
        picQBColors.CurrentX = Int(((iX + (iX + iBtnWidth - 1)) / 2) - (picQBColors.TextWidth(CStr(iIdx)) / 2)) + 1
        picQBColors.CurrentY = Int(((iY + (iY + iBtnHeight - 1)) / 2) - (picQBColors.TextHeight(CStr(iIdx)) / 2)) + 1
        picQBColors.ForeColor = Abs(CInt(iIdx < 10)) * &HFFFFFF
        picQBColors.Print CStr(iIdx)
        miCurQBIdx = iIdx
        mlCurColor = QBColor(iIdx)
    End If
    
End Sub

Private Sub DrawQBButtons()

'Draw all the QBColor buttons

Dim iIdx As Integer

    For iIdx = 0 To 15
        miCurQBIdx = 0
        Call DrawButtonUp(iIdx)
    Next iIdx

End Sub

Private Sub GetSystemColors()

Dim iIdx    As Integer
Dim aColors As Variant

    'Setup the Text and VB constants for the system colors.
    aColors = Array(Array("Scroll Bars", "Desktop", "Active Title Bar", "Inactive Title Bar", "Menu Bar", "Window Background", _
        "Window Frame", "Menu Text", "Window Text", "Active Title Bar Text", "Active Border", "Inactive Border", _
        "Application Workspace", "Highlight", "Highlight Text", "Button Face", "Button Shadow", "Disabled Text", "Button Text", _
        "Inactive Title Bar Text", "Button Highlight", "Button Dark Shadow", "Button Light Shadow", "ToolTip Text", "ToolTip Background"), _
        Array("vbScrollBars", "vbDesktop", "vbActiveTitleBar", "vbInactiveTitleBar", "vbMenuBar", "vbWindowBackground", _
        "vbWindowFrame", "vbMenuText", "vbWindowText", "vbActiveTitleBarText", "vbActiveBorder", "vbInactiveBorder", _
        "vbApplicationWorkspace", "vbHighlight", "vbHighlightText", "vbButtonFace", "vbButtonShadow", "vbGrayText", "vbButtonText", _
        "vbInactiveTitleBarText", "vb3DHighlight", "vb3DDKShadow", "vb3DLight", "vbInfoText", "vbInfoBackground"))
        
    ReDim msaSysColors(UBound(aColors(0)))
    For iIdx = 0 To UBound(aColors(0))
        cboSysColors.AddItem aColors(0)(iIdx)
        msaSysColors(iIdx) = aColors(1)(iIdx)
    Next
    
    Erase aColors
    
End Sub

Private Sub ResetControls()

'Set all controls to their correct values
'based on the current color setting (mlCurColor).

Dim lColor  As Long
Dim iIdx    As Integer
Dim iRed    As Integer
Dim iGreen  As Integer
Dim iBlue   As Integer
Dim sRed    As String
Dim sGreen  As String
Dim sBlue   As String
Dim sHex    As String

    'Stop controls from triggering events that would call this
    'procedure again while changing control properties here.
    mbNoChange = True
    
    'Find out if it's a system color or it matches
    'a system color (Returns -1 if it's not).
    lColor = ConvertToSysColor(mlCurColor)
    
    'Convert system color, if needed.
    'A Long of &H80000000& or greater is a negative number.
    'System colors in VB are &H80000000& to &H80000018& (-2147483648 to -2147483624).
    'All other colors are &H00000000& to &H00FFFFFF& (0 to 16777215).
    If lColor < -1 Then  'If it's a System color...
        cboSysColors.ListIndex = (lColor And &HFF&)
        txtColor(6).Text = msaSysColors(cboSysColors.ListIndex)
        txtColor(7).Text = "&H" & Hex$(lColor) & "&"
        '(&H8000000F& And &HFF&) = &HF(15), which is the real system
        'color index for ButtonFace. GetSysColor() will return the
        'actual color setting for the system based on this index.
        lColor = GetSysColor(lColor And &HFF&)
    Else
        'Not a system color
        lColor = mlCurColor
        txtColor(6).Text = "N/A"
        txtColor(7).Text = "N/A"
        cboSysColors.ListIndex = -1
    End If
    
    'Show the color in the Color box.
    lblColorBox.BackColor = lColor
    
    'VB QBColor and Constants
    Select Case lColor
        Case vbBlack
            txtColor(0).Text = "vbBlack"
            txtColor(1).Text = "QBColor(0)"
        Case vbBlue
            txtColor(0).Text = "vbBlue"
            txtColor(1).Text = "QBColor(9)"
        Case vbGreen
            txtColor(0).Text = "vbGreen"
            txtColor(1).Text = "QBColor(10)"
        Case vbCyan
            txtColor(0).Text = "vbCyan"
            txtColor(1).Text = "QBColor(11)"
        Case vbRed
            txtColor(0).Text = "vbRed"
            txtColor(1).Text = "QBColor(12)"
        Case vbMagenta
            txtColor(0).Text = "vbMagenta"
            txtColor(1).Text = "QBColor(13)"
        Case vbYellow
            txtColor(0).Text = "vbYellow"
            txtColor(1).Text = "QBColor(14)"
        Case vbWhite
            txtColor(0).Text = "vbWhite"
            txtColor(1).Text = "QBColor(15)"
        Case Else
            txtColor(0).Text = "N/A"
            Select Case lColor
                Case QBColor(1)
                    txtColor(1).Text = "QBColor(1)"
                Case QBColor(2)
                    txtColor(1).Text = "QBColor(2)"
                Case QBColor(3)
                    txtColor(1).Text = "QBColor(3)"
                Case QBColor(4)
                    txtColor(1).Text = "QBColor(4)"
                Case QBColor(5)
                    txtColor(1).Text = "QBColor(5)"
                Case QBColor(6)
                    txtColor(1).Text = "QBColor(6)"
                Case QBColor(7)
                    txtColor(1).Text = "QBColor(7)"
                Case QBColor(8)
                    txtColor(1).Text = "QBColor(8)"
                Case Else
                    txtColor(1).Text = "QBColor(N/A)"
            End Select
    End Select
    
    'RGB Color - extract Red, Green and Blue values.
    iRed = (lColor And &HFF&)
    iGreen = (lColor And &HFF00&) / &H100
    iBlue = (lColor And &HFF0000) / &H10000
    txtColor(2).Text = "RGB(" & CStr(iRed) & ", " _
        & CStr(iGreen) & ", " & CStr(iBlue) & ")"
    
    'Decimal Color
    txtColor(3).Text = CStr(lColor)
    
    'Hex Color
    sHex = Hex(lColor)
    If Len(sHex) < 6 Then
        sHex = String(6 - Len(sHex), "0") & sHex
    End If
    txtColor(5).Text = "&H" & sHex & "&"
    
    'HTML Color - Reverse Hex
    txtColor(4).Text = "#" & Right$(sHex, 2) & Mid$(sHex, 3, 2) & Left$(sHex, 2)
    
    'Enable/disable the TextBoxes based on whether
    'or not they contain "N/A".
    For iIdx = 0 To txtColor.UBound
        '(InStr(1, txtColor(iIdx).Text, "N/A") = 0) = True if "N/A" is NOT present
        txtColor(iIdx).Enabled = (InStr(1, txtColor(iIdx).Text, "N/A") = 0)
    Next
    
    'Set the QBColor buttons
    If InStr(1, txtColor(1).Text, "N/A") = 0 Then
        iIdx = Val(Mid$(txtColor(1).Text, 9))
        If iIdx <> miCurQBIdx Then
            Call DrawButtonUp(miCurQBIdx)
            Call DrawButtonDown(iIdx)
        End If
    Else
        Call DrawButtonUp(miCurQBIdx)
    End If
    
    'Setup the color values for the RGB labels
    If chkHex.Value = vbChecked Then
        sRed = Hex(iRed)
        sGreen = Hex(iGreen)
        sBlue = Hex(iBlue)
    Else
        sRed = CStr(iRed)
        sGreen = CStr(iGreen)
        sBlue = CStr(iBlue)
    End If
    
    'Set the RGB and Shade scroll bars
    hsbRGB(0).Value = iRed
    hsbRGB(1).Value = iGreen
    hsbRGB(2).Value = iBlue
    hsbShade.Value = (iRed + iGreen + iBlue) / 3
    
    'Set the RGB TextBoxes if they changed
    If txtRGB(0).Text <> sRed Then
        txtRGB(0).Text = sRed
        txtRGB(0).SelStart = Len(txtRGB(0).Text)
    End If
    If txtRGB(1).Text <> sGreen Then
        txtRGB(1).Text = sGreen
        txtRGB(1).SelStart = Len(txtRGB(1).Text)
    End If
    If txtRGB(2).Text <> sBlue Then
        txtRGB(2).Text = sBlue
        txtRGB(2).SelStart = Len(txtRGB(2).Text)
    End If
    
    'Allow control events again
    mbNoChange = False
    
End Sub

Private Sub cboSysColors_Click()

Dim sHex    As String

    'If change not coming from ResetControls
    If Not mbNoChange Then
        If cboSysColors.ListIndex >= 0 Then
            sHex = Hex(cboSysColors.ListIndex)
            If Len(sHex) < 2 Then
                sHex = "0" & sHex
            End If
            mlCurColor = Val("&H800000" & sHex)
        End If
        Call ResetControls
    End If
    
End Sub

Private Sub chkHex_Click()

Dim iIdx As Integer
Dim iMax As Integer

    If chkHex.Value = vbChecked Then
        iMax = 2
    Else
        iMax = 3
    End If
    
    For iIdx = 0 To 2
        txtRGB(iIdx).MaxLength = iMax
    Next
    
    Call ResetControls

End Sub

Private Sub chkMsg_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Dim wParam As Long
    
    'X contains the wParam from the Windows messaging system
    'when a mouse event occurs over the System Tray Icon.
    'Since the checkbox coordinates are in twips, VB has
    'already converted the message by multiplying the X by
    'Screen.TwipsPerPixelX, so convert it back.
    wParam = X / Screen.TwipsPerPixelX
    
    'This is only using Double-Click and Right-Click,
    'but all of the following events are returned.
    Select Case wParam
        Case WM_MOUSEMOVE
        Case WM_LBUTTONDOWN
        Case WM_LBUTTONUP
        Case WM_LBUTTONDBLCLK
            If Forms.Count > 1 Then
                Call AppActivate(Me.Caption)
            Else
                Me.Show
            End If
        Case WM_RBUTTONDOWN
            ShowPopupAtCursor Me, mnuFile, vbPopupMenuRightButton + vbPopupMenuRightAlign, mnuFile_Show
        Case WM_RBUTTONUP
        Case WM_RBUTTONDBLCLK
    End Select

End Sub


Private Sub cmdAbout_Click()

    frmAbout.Show vbModal, Me
    
End Sub

Private Sub chkCapture_Click()

    If chkCapture.Value = vbChecked Then
        mbCapture = True
        Me.MousePointer = vbCustom
        Call ReleaseCapture
        Call SetCapture(Me.hWnd)
    End If
    
End Sub

Private Sub cmdHide_Click()
    
    Me.WindowState = vbMinimized

End Sub

Private Sub Form_Load()

Dim iIdx As Integer

    If App.PrevInstance Then
        MsgBox "The Color Selector Program is Already Running" _
            & vbCrLf & vbCrLf & "Please check the System Tray or Your Taskbar", _
            vbInformation, "Previous Instance"
        End
    End If
    Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
    DrawQBButtons
    miCurQBIdx = -1
    GetSystemColors
    
    For iIdx = 0 To txtColor.UBound
        txtColor(iIdx).ToolTipText = "Right-click to copy this color code"
    Next
    
    picQBColors_MouseDown 1, 0, picQBColors.ScaleWidth - 5, picQBColors.ScaleHeight - 5
    SetSysTrayIcon AddIcon, chkMsg.hWnd, Me.Icon, Me.Caption
    
    If UCase$(Command$) = "/H" Then
        Me.Hide
    End If
    
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

Dim lColor  As Long
Dim hDeskDC As Long
Dim ptMouse As PointAPI

    If mbCapture Then
        Call ReleaseCapture
        Call GetCursorPos(ptMouse)
        hDeskDC = GetDC(0)
        lColor = GetPixel(hDeskDC, ptMouse.X, ptMouse.Y)
        If lColor <> -1 Then
            mlCurColor = lColor
            Call ResetControls
        End If
        Me.MousePointer = vbDefault
        chkCapture.Value = vbUnchecked
    End If
    
End Sub


Private Sub Form_Paint()

Dim iIdx    As Integer
Dim iSavIdx As Integer
Dim lLeft   As Long
Dim lTop    As Long
Dim lRight  As Long
Static bBusy As Boolean

    'Don't allow re-entry caused by our own painting
    If Not bBusy Then
        bBusy = True
        lLeft = fraQBColors.Left - Screen.TwipsPerPixelX
        lRight = txtColor(3).Left + txtColor(3).Width + (2 * Screen.TwipsPerPixelX)
        
        lTop = lblCap(0).Top - ((lblCap(0).Top - (fraQBColors.Top + fraQBColors.Height)) / 2)
        lTop = Int(lTop / Screen.TwipsPerPixelY) * Screen.TwipsPerPixelY
        Me.Line (lLeft, lTop)-(lRight, lTop), vb3DShadow
        Me.Line (lLeft, lTop + Screen.TwipsPerPixelY)-(lRight, lTop + Screen.TwipsPerPixelY), vb3DHighlight
        
'        lTop = Me.ScaleHeight - ((Me.ScaleHeight - (txtColor(4).Top + txtColor(4).Height)) / 2)
'        lTop = Int(lTop / Screen.TwipsPerPixelY) * Screen.TwipsPerPixelY
'        Me.Line (lLeft, lTop)-(lRight, lTop), vb3DShadow
'        Me.Line (lLeft, lTop + Screen.TwipsPerPixelY)-(lRight, lTop + Screen.TwipsPerPixelY), vb3DHighlight
        
        iSavIdx = miCurQBIdx
        For iIdx = 0 To 15
            Call DrawButtonUp(iIdx)
        Next iIdx
        miCurQBIdx = iSavIdx
        If miCurQBIdx >= 0 Then
            Call DrawButtonDown(miCurQBIdx)
        End If
        
        bBusy = False
    End If
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'
'    If UnloadMode = vbFormControlMenu Then
'        Cancel = True
'        cmdHide.Value = True
'    End If
    
End Sub

Private Sub Form_Resize()

    If Me.WindowState = vbMinimized Then
        Me.Hide
        Me.WindowState = vbNormal
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    SetSysTrayIcon DeleteIcon, chkMsg.hWnd, Me.Icon, Me.Caption

End Sub

Private Sub hsbRGB_Change(Index As Integer)

    'If change not coming from ResetControls
    If Not mbNoChange Then
        mlCurColor = RGB(hsbRGB(0).Value, hsbRGB(1).Value, hsbRGB(2).Value)
        Call ResetControls
    End If
    
End Sub

Private Sub hsbRGB_Scroll(Index As Integer)
    
    Call hsbRGB_Change(Index)

End Sub

Private Sub hsbShade_Change()

Dim iRed        As Integer
Dim iGreen      As Integer
Dim iBlue       As Integer
Dim iChange     As Integer
Dim lColor      As Long
Static iOldVal  As Integer

    'If change not coming from ResetControls
    If Not mbNoChange Then
        iChange = hsbShade.Value - iOldVal
        lColor = RGB(hsbRGB(0).Value, hsbRGB(1).Value, hsbRGB(2).Value)
        iRed = (lColor And &HFF&)
        iGreen = (lColor And &HFF00&) / &H100
        iBlue = (lColor And &HFF0000) / &H10000
        iRed = iRed + iChange
        iGreen = iGreen + iChange
        iBlue = iBlue + iChange
        If iRed > 255 Then
            iRed = 255
        ElseIf iRed < 0 Then
            iRed = 0
        End If
        If iGreen > 255 Then
            iGreen = 255
        ElseIf iGreen < 0 Then
            iGreen = 0
        End If
        If iBlue > 255 Then
            iBlue = 255
        ElseIf iBlue < 0 Then
            iBlue = 0
        End If
        mlCurColor = RGB(iRed, iGreen, iBlue)
        Call ResetControls
    End If

    iOldVal = hsbShade.Value
    
End Sub

Private Sub hsbShade_Scroll()
    Call hsbShade_Change
End Sub

Private Sub lblColorBox_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    frmColor.Show
    
End Sub


Private Sub mnuFile_Exit_Click()
    
    On Error Resume Next
    Unload frmAbout
    Unload Me

End Sub

Private Sub mnuFile_Show_Click()
    
    If Forms.Count > 1 Then
        Call AppActivate(Me.Caption)
    Else
        Me.Show
    End If

End Sub


Private Sub picQBColors_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

Dim iIdx    As Integer
Dim sColor  As String

    If Button = 1 Then
        iIdx = (Int(X / (picQBColors.ScaleWidth / 2)) * 8) + _
            Int(Y / (picQBColors.ScaleHeight / 8))
        If mlCurColor <> QBColor(iIdx) Then
            mlCurColor = QBColor(iIdx)
            Call ResetControls
        End If
    End If
    
End Sub

Private Sub txtColor_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    With txtColor(Index)
        .SelStart = 0
        .SelLength = Len(.Text)
    End With

End Sub


Private Sub txtRGB_Change(Index As Integer)

Dim iIdx    As Integer
Dim iVal    As Integer
Dim sChar   As String
Dim sText   As String
Static bBusy As Boolean
Static sOldText(2) As String

    If Not (bBusy Or mbNoChange) Then
        bBusy = True
        With txtRGB(Index)
            For iIdx = 1 To Len(.Text)
                sChar = UCase$(Mid$(.Text, iIdx, 1))
                If sChar >= "0" And sChar <= "9" Then
                    sText = sText & sChar
                ElseIf chkHex.Value = vbChecked Then
                    If sChar >= "A" And sChar <= "F" Then
                        sText = sText & sChar
                    End If
                End If
            Next
            If chkHex.Value = vbChecked Then
                iVal = Val("&H" & sText)
            Else
                iVal = Val(sText)
            End If
            If Len(sText) <> Len(.Text) Or iVal > 255 Then
                Beep
                .Text = sOldText(Index)
                .SelStart = Len(.Text)
            End If
            sOldText(Index) = .Text
            hsbRGB(Index).Value = iVal
        End With
        bBusy = False
    End If
    
End Sub

Private Sub txtRGB_GotFocus(Index As Integer)

    With txtRGB(Index)
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
    
End Sub


Private Sub txtRGB_KeyPress(Index As Integer, KeyAscii As Integer)

    KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
    
    If KeyAscii > 8 And (KeyAscii < vbKey0 Or KeyAscii > vbKey9) Then
        If chkHex.Value <> vbChecked Then
            Beep
            KeyAscii = 0
        ElseIf KeyAscii < vbKeyA Or KeyAscii > vbKeyF Then
            Beep
            KeyAscii = 0
        End If
    End If
        
End Sub


