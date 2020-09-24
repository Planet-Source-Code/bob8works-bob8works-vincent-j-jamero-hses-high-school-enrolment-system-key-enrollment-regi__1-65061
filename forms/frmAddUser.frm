VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmAddUser 
   BackColor       =   &H00D8E9EC&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add New User"
   ClientHeight    =   6075
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9015
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAddUser.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   405
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   601
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtUserName 
      Height          =   330
      Left            =   1230
      MaxLength       =   20
      TabIndex        =   4
      Top             =   810
      Width           =   3540
   End
   Begin VB.TextBox txtPassword 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      IMEMode         =   3  'DISABLE
      Left            =   1230
      MaxLength       =   20
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   1230
      Width           =   3540
   End
   Begin VB.TextBox txtFullName 
      Height          =   330
      Left            =   1230
      MaxLength       =   60
      TabIndex        =   2
      Top             =   1590
      Width           =   3540
   End
   Begin VB.ComboBox cmbType 
      Height          =   315
      Left            =   1230
      TabIndex        =   1
      Text            =   "Select Type"
      Top             =   1950
      Width           =   3585
   End
   Begin HSES.b8Line b8Line1 
      Height          =   60
      Left            =   -30
      TabIndex        =   0
      Top             =   2400
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   106
   End
   Begin HSES.b8SContainer b8SContainer1 
      Height          =   570
      Left            =   -60
      TabIndex        =   5
      Top             =   5520
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   1005
      BorderColor     =   14737632
      Begin lvButton.lvButtons_H cmdSave 
         Height          =   405
         Left            =   7500
         TabIndex        =   6
         Top             =   90
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   714
         Caption         =   "&Save"
         CapAlign        =   2
         BackStyle       =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cBhover         =   14215660
         cGradient       =   14215660
         Gradient        =   3
         Mode            =   0
         Value           =   0   'False
         cBack           =   16777215
      End
      Begin lvButton.lvButtons_H cmdCancel 
         Height          =   405
         Left            =   5940
         TabIndex        =   7
         Top             =   90
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   714
         Caption         =   "&Cancel"
         CapAlign        =   2
         BackStyle       =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cBhover         =   14215660
         cGradient       =   14215660
         Gradient        =   3
         Mode            =   0
         Value           =   0   'False
         cBack           =   16777215
      End
   End
   Begin MSComctlLib.ListView listAccess 
      Height          =   2745
      Left            =   60
      TabIndex        =   8
      Top             =   2730
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   4842
      View            =   3
      Arrange         =   1
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Title"
         Object.Width           =   5821
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "For User"
         Object.Width           =   3175
      EndProperty
   End
   Begin HSES.b8Line b8Line2 
      Height          =   60
      Left            =   0
      TabIndex        =   14
      Top             =   510
      Width           =   8985
      _ExtentX        =   15849
      _ExtentY        =   106
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Add User"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   345
      Left            =   120
      TabIndex        =   15
      Top             =   180
      Width           =   1290
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Restrictions"
      Height          =   195
      Left            =   90
      TabIndex        =   13
      Top             =   2490
      Width           =   840
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "User Name"
      Height          =   195
      Left            =   270
      TabIndex        =   12
      Top             =   840
      Width           =   780
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      Height          =   195
      Left            =   270
      TabIndex        =   11
      Top             =   1260
      Width           =   690
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Full Name"
      Height          =   195
      Left            =   270
      TabIndex        =   10
      Top             =   1590
      Width           =   690
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Type"
      Height          =   195
      Left            =   270
      TabIndex        =   9
      Top             =   1980
      Width           =   360
   End
   Begin VB.Image Image2 
      Height          =   525
      Left            =   0
      Picture         =   "frmAddUser.frx":058A
      Stretch         =   -1  'True
      Top             =   0
      Width           =   8985
   End
End
Attribute VB_Name = "frmAddUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Added As Boolean
Dim oldTypeListIndex As Integer
Public Function ShowForm(Optional CanAddAdmin As Boolean = False) As Boolean
    
    'default
    Added = False
    
    If CurrentUser.UserName <> "system" Then
        'check current user
        If CurrentUser.UserType <> sAdministratortitle Then
            MsgBox "Unable to show Manage Users window." & vbNewLine & _
                    "You are not permitted to aceess it. Please contact your Administrator.", vbExclamation
                    
            Unload Me
            Exit Function
        End If
        
        'set parameter
        If CanAddAdmin = True Then
            'bgAccess.Enabled = True
            ShowAccessTypes
        Else
            'bgAccess.Enabled = False
            ShowAccessTypes
        End If
       
    End If
    
    
    
    'show form
    Me.Show vbModal
    
    'return
    ShowForm = Added
End Function


Private Function ShowAsSystem()
    
    Dim i As Integer
    
    
    ShowAccessTypes

    For i = 1 To listAccess.ListItems.Count
        listAccess.ListItems(i).Checked = True
    Next
    
    txtUserName.Text = "Administrator"
    txtUserName.Enabled = False
    
    cmbType.ListIndex = 0
    cmbType.Enabled = False
    listAccess.Enabled = False
    
    
    lblTitle.Caption = "Create Administrator Account"
    
End Function

Private Sub ShowAccessTypes()
    Dim vRS As Recordset
    Dim sSQL As String
    
    
    sSQL = "SELECT tblAccessType.AccessTitle, tblAccessType.AccessTitle, tblAccessType.AccessibleFor" & _
        " FROM tblAccessType" & _
        " ORDER BY Rank"

    
    'connect Table Access Type
    If ConnectRS(HSESDB, vRS, sSQL) = True Then
        If AnyRecordExisted(vRS) = True Then
            
            FillRecordToList vRS, listAccess, keyUser
        End If
    Else
        CatchError "frmAddUser", "ShowAccessTypes", "ConnectRS(HSESDB, vRS, SELECT * FROM tblAccessType)"
    End If
        
    Set vRS = Nothing
End Sub



Private Sub cmbType_Click()
    
    Dim lvItem As ListItem
    
    
    
    If cmbType.Text = sAdministratortitle Then
    
        For Each lvItem In listAccess.ListItems
            If lvItem.SubItems(1) = sAdministratortitle Then
                lvItem.Checked = True
            End If
        Next
        
        
    ElseIf cmbType.Text = sEncoderTitle Then
    
        For Each lvItem In listAccess.ListItems
            If lvItem.SubItems(1) = sAdministratortitle Then
                lvItem.Checked = False
            End If
        Next
        
    End If
    
    
End Sub

Private Sub cmbType_GotFocus()
    oldTypeListIndex = cmbType.ListIndex
End Sub

Private Sub cmbType_LostFocus()
    If cmbType.ListIndex < 0 Then
        cmbType.ListIndex = oldTypeListIndex
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    Dim newUser As User
    Dim sAccessTitle() As String
    Dim lvItem As ListItem
    Dim i As Integer
    Dim ItemCount As Integer
    
    If CheckTextBox(txtUserName, "Please fill up User Name.") = False Then
        Exit Sub
    End If
    
    If CheckTextBox(txtPassword, "Please fill up Password.") = False Then
        Exit Sub
    End If
    
    If CheckTextBox(txtFullName, "Please fill up Full Name.") = False Then
        Exit Sub
    End If
    
    If CheckTextBox(cmbType, "Please fill up valid User Type.") = False Then
        Exit Sub
    End If
    
    
    'check access
    ItemCount = -1
    For Each lvItem In listAccess.ListItems
        If lvItem.Checked = True Then
            ItemCount = ItemCount + 1
        End If
    Next
    
    If ItemCount < 0 Then
        MsgBox "Unable to Save User Entry." & vbNewLine & _
            "Please check at least one of users access.", vbExclamation
        
        'listAccess.SetFocus
        Exit Sub
    End If
    
    'set new user
    newUser.UserName = txtUserName.Text
    newUser.Password = txtPassword.Text
    newUser.FullName = txtFullName.Text
    newUser.UserType = cmbType.Text
    newUser.CreationDate = Now
    newUser.CreatedBy = CurrentUser.UserName
    
    'save user
    Select Case AddUser(newUser)
        Case Success
        'SaveAccess
            
            If listAccess.ListItems.Count - 1 > -1 Then
                
                ItemCount = -1
                For Each lvItem In listAccess.ListItems
                    If lvItem.Checked = True Then
                        ItemCount = ItemCount + 1
                    End If
                Next
                
                
                ReDim sAccessTitle(ItemCount)
                
               i = 0
               For Each lvItem In listAccess.ListItems
                
                    If lvItem.Checked = True Then
                    
                        sAccessTitle(i) = lvItem.Text
                        i = i + 1
                    End If
                Next
                
                Select Case SaveAccess(CurrentUser.UserName, txtUserName.Text, sAccessTitle)
                    Case TranDBResult.Success
                        'USER ADDED Succesfully
                        
                        'set flag to ADDED
                        Added = True
                        'unload this form
                        
                        If Added = True And CurrentUser.UserName = "system" Then
                            Unload Me
                            frmLogin.ShowLogin
                            Exit Sub
                        End If
                        
                        
                        If MsgBox("New User entry succesfull created." & vbNewLine & vbNewLine & _
                                    "Dou want to Add another?", vbQuestion + vbYesNo) = vbYes Then
                            'close this form
                            Unload Me
                            'add another
                            frmAddUser.ShowForm
                        End If
                        
                        'close this form
                        Unload Me
                        
                    Case TranDBResult.UserNotExist
                        MsgBox "Saving Access Settings Error: User Not Found", vbCritical
                    Case Else
                        CatchError "frmAddUser", "AddUser", "Save Access - User does not exist"
                End Select
            End If
            
        Case TranDBResult.UserDuplicate
            MsgBox "The User named '" & txtUserName.Text & "' is already existed. Please enter another User Name.", vbExclamation
            HLTxt txtUserName
            
    End Select
End Sub

Private Sub Form_Activate()
    If CurrentUser.UserName = "system" Then
        ShowAsSystem
    End If
End Sub

Private Sub Form_Load()
    'set user type
    cmbType.Clear
    cmbType.AddItem sAdministratortitle
    cmbType.AddItem sEncoderTitle
    cmbType.ListIndex = 0
End Sub


Private Sub listAccess_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    Call cmbType_Click
End Sub
