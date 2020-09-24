VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form PickUser 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   0  'None
   Caption         =   "Select User"
   ClientHeight    =   2880
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4170
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   192
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   278
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList imgListUser 
      Left            =   2835
      Top             =   1575
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUserList.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUserList.frx":059A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView listUsers 
      Height          =   2535
      Left            =   30
      TabIndex        =   0
      Top             =   315
      Width           =   4110
      _ExtentX        =   7250
      _ExtentY        =   4471
      View            =   3
      LabelWrap       =   0   'False
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      FlatScrollBar   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HotTracking     =   -1  'True
      HoverSelection  =   -1  'True
      _Version        =   393217
      Icons           =   "imgListUser"
      SmallIcons      =   "imgListUser"
      ForeColor       =   -2147483640
      BackColor       =   16777215
      Appearance      =   0
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   4763
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Object.Width           =   2910
      EndProperty
   End
   Begin VB.Image Image1 
      Height          =   240
      Left            =   45
      Picture         =   "frmUserList.frx":0E74
      Top             =   45
      Width           =   240
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Select User"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   360
      TabIndex        =   2
      Top             =   45
      Width           =   960
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "[ENTER] Select  [ESC] Cancel"
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Left            =   1950
      TabIndex        =   1
      Top             =   45
      Width           =   2100
   End
End
Attribute VB_Name = "PickUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Dim sGetUserName As String



Public Function GetUserName(Optional TextObject) As String
    
    Dim R As RECT
    Dim P As POINTAPI
        


    'set position
    If Not IsMissing(TextObject) Then
    
        GetWindowRect TextObject.hwnd, R
        Me.Left = R.Left * Screen.TwipsPerPixelX
        Me.Top = R.Bottom * Screen.TwipsPerPixelY
    Else
        
        GetCursorPos P
        Me.Left = P.X * Screen.TwipsPerPixelX
        Me.Top = P.Y * Screen.TwipsPerPixelY
    End If
    
    
    'add all user to list
        FillList
    
    'clear temporary return variable
    sGetUserName = ""

    
    'show form
    Me.Show vbModal
    
    GetUserName = sGetUserName
End Function


Private Sub FillList()

End Sub




Private Sub Form_Activate()
    listUsers.SetFocus
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyEscape
            Call CancelGetUser
        Case vbKeyReturn
            Call ReturnGetUser
    End Select
End Sub

Private Sub ReturnGetUser()
    'return selected user name
    sGetUserName = listUsers.SelectedItem.Text
    Unload Me
End Sub
Private Sub CancelGetUser()
    'return null string
    sGetUserName = ""
    Unload Me
End Sub

Private Sub listUsers_DblClick()
    Call ReturnGetUser
End Sub
