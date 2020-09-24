VERSION 5.00
Begin VB.Form frmSplash 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4425
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6465
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmSplash.frx":0000
   ScaleHeight     =   295
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   431
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   2640
      Top             =   1980
   End
   Begin VB.Label lblSchoolName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Welcome"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   210
      Left            =   90
      TabIndex        =   0
      Top             =   3780
      Width           =   765
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Declare Function UpdateLayeredWindow Lib "user32" (ByVal hwnd As Long, ByVal hdcDst As Long, pptDst As Any, psize As Any, ByVal hdcSrc As Long, pptSrc As Any, crKey As Long, ByVal pblend As Long, ByVal dwFlags As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Const G = (-20)
Private Const LWA_COLORKEY = &H1
Private Const LWA_ALPHA = &H2
Private Const ULW_COLORKEY = &H1
Private Const ULW_ALPHA = &H2
Private Const ULW_OPAQUE = &H4
Private Const WS_EX_LAYERED = &H80000

Private Declare Function SetWindowPos Lib "user32" _
    (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, _
    ByVal X As Long, ByVal Y As Long, ByVal cx As Long, _
    ByVal cy As Long, ByVal wFlags As Long) As Long
    
    Private Const SWP_NOMOVE = 2
Private Const SWP_NOSIZE = 1
Private Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2



Dim isAbout As Boolean



Dim tl As Integer

Public Sub ShowSplash()
    isAbout = False
    lblSchoolName.Caption = CurrentSchool.SchoolName
    Me.Show
    DoEvents
End Sub


Public Sub UnloadSplash()
    Me.Enabled = False
    Timer1.Enabled = True
End Sub



Public Function ShowAbout()
    isAbout = True
    lblSchoolName.Caption = CurrentSchool.SchoolName
    Me.Show
    
End Function

 
Private Sub Form_Activate()
    SetWindowPos Me.hwnd, HWND_TOPMOST, _
    0, 0, 0, 0, FLAGS
End Sub

Private Sub Form_Deactivate()
    If isAbout = True Then
        UnloadSplash
    End If
End Sub

Private Sub Trans(Level As Integer)
        Dim Msg As Long

        Msg = GetWindowLong(Me.hwnd, G)
        Msg = Msg Or WS_EX_LAYERED
        SetWindowLong Me.hwnd, G, Msg
        SetLayeredWindowAttributes Me.hwnd, 0, Level, LWA_ALPHA
        MakeSemiTransparent = 0
End Sub

Private Sub Form_Load()
    tl = 100
End Sub

Private Sub Picture1_Click()

End Sub

Private Sub Timer1_Timer()
    Trans tl
    
    tl = tl - 10
    
    If tl < 10 Then
        Timer1.Enabled = False
        tl = 100
        Unload Me

    End If
End Sub
