VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmEditUser 
   BackColor       =   &H00D8E9EC&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Edit User"
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
   Icon            =   "frmEditUser.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   405
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   601
   StartUpPosition =   2  'CenterScreen
   Begin HSES.b8Line b8Line1 
      Height          =   60
      Left            =   -30
      TabIndex        =   9
      Top             =   2400
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   106
   End
   Begin VB.ComboBox cmbType 
      Height          =   315
      Left            =   1230
      TabIndex        =   7
      Text            =   "Select Type"
      Top             =   1950
      Width           =   3585
   End
   Begin VB.TextBox txtFullName 
      Height          =   330
      Left            =   1230
      MaxLength       =   60
      TabIndex        =   6
      Top             =   1590
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
      TabIndex        =   4
      Top             =   1230
      Width           =   3540
   End
   Begin VB.TextBox txtUserName 
      Height          =   330
      Left            =   1230
      MaxLength       =   20
      TabIndex        =   2
      Top             =   810
      Width           =   3540
   End
   Begin HSES.b8SContainer b8SContainer1 
      Height          =   570
      Left            =   -60
      TabIndex        =   0
      Top             =   5520
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   1005
      BorderColor     =   14737632
      Begin lvButton.lvButtons_H cmdSave 
         Height          =   405
         Left            =   7500
         TabIndex        =   10
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
         TabIndex        =   11
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
      TabIndex        =   12
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
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Edit User"
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
      Top             =   120
      Width           =   1305
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
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Type"
      Height          =   195
      Left            =   270
      TabIndex        =   8
      Top             =   1980
      Width           =   360
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Full Name"
      Height          =   195
      Left            =   270
      TabIndex        =   5
      Top             =   1590
      Width           =   690
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      Height          =   195
      Left            =   270
      TabIndex        =   3
      Top             =   1260
      Width           =   690
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "User Name"
      Height          =   195
      Left            =   270
      TabIndex        =   1
      Top             =   840
      Width           =   780
   End
   Begin VB.Image Image2 
      Height          =   525
      Left            =   0
      Picture         =   "frmEditUser.frx":058A
      Stretch         =   -1  'True
      Top             =   0
      Width           =   8985
   End
End
Attribute VB_Name = "frmEditUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Edited As Boolean
Dim oldTypeListIndex As Integer

Dim curUser As User

Public Function ShowForm(sUserName As String) As Boolean
    
    'default
    Edited = False
    
    'check current user
    If CurrentUser.UserType <> sAdministratortitle Then
        MsgBox "Unable to show Manage Users window." & vbNewLine & _
                "You are not permitted to aceess it. Please contact your Administrator.", vbExclamation
                
        Unload Me
        Exit Function
    End If
    
    If GetUserByName(curUser, sUserName) = Success Then
        txtUserName.Text = curUser.UserName
        txtPassword.Text = curUser.Password
        txtFullName.Text = curUser.FullName
        cmbType.Text = curUser.UserType
    Else
        MsgBox "Unable to continue Editing User." & vbNewLine & _
            "The selected User account does not exist.", vbExclamation
        
        Unload Me
        Exit Function
    End If
    
    'show access types
    ShowAccessTypes
    
    'get old access values
    GetAccessTypes
    
    'refresh access list
    cmbType_Click
    
    'show form
    Me.Show vbModal
    
    'return
    ShowForm = Edited
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

Private Sub GetAccessTypes()
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    sSQL = "SELECT tblUser.UserName, tblUserAccess.AllowedTo" & _
            " FROM tblUser INNER JOIN tblUserAccess ON tblUser.UserName = tblUserAccess.UserName" & _
            " WHERE (((tblUser.UserName)='" & curUser.UserName & "'));"

    If ConnectRS(HSESDB, vRS, sSQL) = True Then
        If AnyRecordExisted(vRS) = True Then
        
            vRS.MoveFirst
            
            While Not vRS.EOF
            
                SetCheckList ReadField(vRS.Fields("AllowedTo"))
                
                vRS.MoveNext
            Wend
            
        End If
    End If
    
    Set vRS = Nothing
End Sub


Private Sub SetCheckList(sAccessTitle As String)
    Dim lvItem As ListItem
    
    For Each lvItem In listAccess.ListItems
        If lvItem.Text = sAccessTitle Then
            lvItem.Checked = True
        End If
    Next
End Sub

Private Sub bgAccess_Click()

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
        
        listAccess.SetFocus
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
    Select Case EditUser(newUser)
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
                        'USER Edited Succesfully
                        
                        MsgBox "User account successfully edited.", vbInformation
                        'set flag to Edited
                        Edited = True
                        
                        'unload this form
                        Unload Me

                    Case TranDBResult.UserNotExist
                        MsgBox "Editing Access Settings Error: User Not Found", vbCritical
                    Case Else
                        CatchError "frmAddUser", "AddUser", "Save Access - User does not exist"
                End Select
            End If

        Case Else
        
        CatchError "frmEditUser", "cmdSave_Click", "Edit User - unknown error"
            
    End Select
End Sub

Private Sub Form_Load()
    'set user type
    cmbType.Clear
    cmbType.AddItem sAdministratortitle
    cmbType.AddItem sEncoderTitle
    cmbType.ListIndex = 1
End Sub

Private Sub listAccess_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    SortLV listAccess, ColumnHeader.Index - 1

End Sub

Private Sub listAccess_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    Call cmbType_Click
End Sub

