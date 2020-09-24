VERSION 5.00
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmAddDepartment 
   BackColor       =   &H00D8E9EC&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Department"
   ClientHeight    =   2205
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "frmAddDepartment.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   147
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtDepartmentID 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1290
      Locked          =   -1  'True
      MaxLength       =   4
      TabIndex        =   1
      Top             =   810
      Width           =   3135
   End
   Begin VB.TextBox txtDepartmentTitle 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1290
      MaxLength       =   20
      TabIndex        =   0
      Top             =   1200
      Width           =   3135
   End
   Begin HSES.b8Line b8Line1 
      Height          =   60
      Left            =   30
      TabIndex        =   2
      Top             =   510
      Width           =   5235
      _ExtentX        =   9234
      _ExtentY        =   106
   End
   Begin lvButton.lvButtons_H cmdSave 
      Default         =   -1  'True
      Height          =   360
      Left            =   3180
      TabIndex        =   3
      Top             =   1800
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   635
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
      Focus           =   0   'False
      cGradient       =   14215660
      Gradient        =   3
      Mode            =   0
      Value           =   0   'False
      cBack           =   16185592
   End
   Begin lvButton.lvButtons_H cmdCancel 
      Cancel          =   -1  'True
      Height          =   360
      Left            =   1650
      TabIndex        =   4
      Top             =   1800
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   635
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
      Focus           =   0   'False
      cGradient       =   14215660
      Gradient        =   3
      Mode            =   0
      Value           =   0   'False
      cBack           =   16185592
   End
   Begin HSES.b8Line b8Line2 
      Height          =   60
      Left            =   0
      TabIndex        =   5
      Top             =   1710
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   106
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Title"
      Height          =   195
      Left            =   240
      TabIndex        =   8
      Top             =   1260
      Width           =   300
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ID"
      Height          =   195
      Left            =   225
      TabIndex        =   7
      Top             =   840
      Width           =   165
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "New Department"
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
      Left            =   60
      TabIndex        =   6
      Top             =   180
      Width           =   2430
   End
   Begin VB.Image Image2 
      Height          =   525
      Left            =   0
      Picture         =   "frmAddDepartment.frx":058A
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4695
   End
End
Attribute VB_Name = "frmAddDepartment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RecordAdded As Boolean

Public Function ShowForm() As Boolean
    
    Dim sNewID As String
    '--------------------------------------------------
    
    
    '-------------------------------------------------------
    'check user access
    '-------------------------------------------------------
    If UserAllowedTo(CurrentUser.UserName, sCanAddDepartment) = False Then
        MsgBox "Unable to continue adding Department entry." & vbNewLine & vbNewLine & _
                "You are not permitted to do this. Please contact your administrator for more information.", vbExclamation
        Exit Function
    End If
    '-------------------------------------------------------

    

    
    'generate new id
    If GetNewDepartmentID(sNewID) = Failed Then
        CatchError "frmAddDepartment", "ShowForm()", "GetNewDepartmentID(sNewID) = Failed"
        Exit Function
    End If
    
    'set id
    txtDepartmentID.Text = sNewID
    
    
    '--------------------------------------------------
    'show form
    Me.Show vbModal
    'return
    ShowForm = RecordAdded
    
End Function


Private Sub cmdCancel_Click()
    Unload Me
End Sub


Private Function SaveNewDepartment()

    If Not CheckTextBox(txtDepartmentID, "Please enter Department ID") Then
        Exit Function
    End If
    
    If Not CheckTextBox(txtDepartmentTitle, "Please enter Department Title") Then
        Exit Function
    End If
    
    
    'save
    Dim newDepartment As tDepartment
    
    newDepartment.DepartmentID = txtDepartmentID.Text
    newDepartment.DepartmentTitle = txtDepartmentTitle.Text
    
    Select Case AddDepartment(newDepartment)
        Case TranDBResult.Success  'success
            MsgBox "New Department succesfully added", vbInformation
            RecordAdded = True
            Unload Me
            
        Case TranDBResult.DuplicateID
            MsgBox "Invalid Department ID!" & vbNewLine & "The Department ID that you have entered is already existed. Enter another Department ID.", vbExclamation
            HLTxt txtDepartmentID
            
        Case TranDBResult.DuplicateTitle
        
            MsgBox "Invalid Department Title!" & vbNewLine & "The Department Title that you have entered is already existed. Enter another Department Title.", vbExclamation
            HLTxt txtDepartmentTitle
            
        Case Else
            MsgBox "Unknown Error", vbExclamation
            CatchError "frmAddDepartment", "SaveNewDepartment", "Unknown result in Add New Department"
    End Select
    
    
End Function

Private Sub cmdSave_Click()
    SaveNewDepartment
End Sub

