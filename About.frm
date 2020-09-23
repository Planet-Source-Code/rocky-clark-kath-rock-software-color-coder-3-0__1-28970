VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " About Color Coder..."
   ClientHeight    =   2775
   ClientLeft      =   45
   ClientTop       =   255
   ClientWidth     =   4515
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2775
   ScaleWidth      =   4515
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   360
      Left            =   1830
      TabIndex        =   3
      Top             =   2025
      Width           =   1065
   End
   Begin VB.PictureBox picLogo 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   525
      Left            =   165
      MouseIcon       =   "About.frx":0000
      MousePointer    =   99  'Custom
      Picture         =   "About.frx":0152
      ScaleHeight     =   465
      ScaleWidth      =   1635
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
   Begin VB.Frame fraAbout 
      Height          =   1110
      Left            =   135
      MouseIcon       =   "About.frx":0DB1
      TabIndex        =   1
      Top             =   720
      Width           =   4215
      Begin VB.Label lblAbout 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Original design and coding by Rocky Clark kathrock@cfl.rr.com"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   165
         TabIndex        =   2
         Top             =   345
         Width           =   3840
      End
   End
   Begin VB.Label lblStatus 
      BorderStyle     =   1  'Fixed Single
      Height          =   270
      Left            =   0
      TabIndex        =   5
      Top             =   2505
      Width           =   4515
   End
   Begin VB.Label lblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Color Coder"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   465
      Left            =   2040
      TabIndex        =   4
      Top             =   165
      Width           =   2220
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const SW_RESTORE = 9
Private Const msEmailAddr = "mailto: kathrock@cfl.rr.com?subject=RE: Color Coder"
Private Const msWebAddr = "http://www.Kath-Rock.com"

Private Declare Function GetCapture Lib "user32" () As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function SetCapture Lib "user32" (ByVal hWnd As Long) As Long

Private Sub cmdOK_Click()

    Unload Me
    
End Sub

Private Sub fraAbout_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If GetCapture() <> fraAbout.hWnd Then
        Call SetCapture(fraAbout.hWnd)
    End If
    
    If (X < 0) Or (X > fraAbout.Width) Or _
      (Y < 0) Or (Y > fraAbout.Height) Then
        Call ReleaseCapture
    Else
        If (X >= lblAbout.Left) And _
          (X < lblAbout.Left + lblAbout.Width) And _
          (Y >= lblAbout.Top) And _
          (Y < lblAbout.Top + lblAbout.Height) Then
            If lblAbout.ForeColor <> &H800000 Then
                lblAbout.ForeColor = &H800000
                lblAbout.Font.Underline = True
                fraAbout.MousePointer = vbCustom
                lblStatus.Caption = " " & msEmailAddr
            End If
        Else
            If lblAbout.ForeColor <> vbButtonText Then
                lblAbout.ForeColor = vbButtonText
                lblAbout.Font.Underline = False
                fraAbout.MousePointer = vbDefault
                lblStatus.Caption = ""
            End If
        End If
    End If
        
End Sub


Private Sub fraAbout_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If fraAbout.MousePointer = vbCustom Then
        If lblAbout.ForeColor = &H800000 Then
            Call lblAbout_Click
        End If
    End If

End Sub


Private Sub lblAbout_Click()

    Call ShellExecute(Me.hWnd, "Open", msEmailAddr, _
        vbNullString, vbNullString, SW_RESTORE)
    lblStatus.Caption = ""
    
End Sub


Private Sub picLogo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If GetCapture <> picLogo.hWnd Then
        Call SetCapture(picLogo.hWnd)
        lblStatus.Caption = " " & msWebAddr
    End If
    
    If X < 0 Or X > picLogo.ScaleWidth Or Y < 0 Or Y > picLogo.ScaleHeight Then
        Call ReleaseCapture
        lblStatus.Caption = ""
    End If
    
End Sub

Private Sub picLogo_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = vbLeftButton Then
        'This should open the default email client on any system.
        Call ShellExecute(Me.hWnd, "Open", msWebAddr, _
            vbNullString, vbNullString, SW_RESTORE)
        lblStatus.Caption = ""
    End If

End Sub



