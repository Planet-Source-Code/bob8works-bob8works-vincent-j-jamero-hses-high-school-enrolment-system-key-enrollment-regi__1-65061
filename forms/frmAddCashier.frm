VERSION 5.00
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmAddCashier 
   BackColor       =   &H00D8E9EC&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cashier"
   ClientHeight    =   4920
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   5115
   Icon            =   "frmAddCashier.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4920
   ScaleWidth      =   5115
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtAddress 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   1455
      MaxLength       =   70
      MultiLine       =   -1  'True
      TabIndex        =   17
      Top             =   3465
      Width           =   3135
   End
   Begin VB.TextBox txtLastName 
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
      Left            =   1455
      MaxLength       =   50
      TabIndex        =   14
      Top             =   2745
      Width           =   3135
   End
   Begin VB.TextBox txtContactNumber 
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
      Left            =   1455
      MaxLength       =   20
      TabIndex        =   13
      Top             =   3105
      Width           =   3135
   End
   Begin VB.TextBox txtFirstName 
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
      Left            =   1455
      MaxLength       =   50
      TabIndex        =   10
      Top             =   2025
      Width           =   3135
   End
   Begin VB.TextBox txtMiddleName 
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
      Left            =   1455
      MaxLength       =   50
      TabIndex        =   9
      Top             =   2385
      Width           =   3135
   End
   Begin VB.TextBox txtPassword 
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
      IMEMode         =   3  'DISABLE
      Left            =   1455
      MaxLength       =   20
      PasswordChar    =   "*"
      TabIndex        =   7
      Top             =   1665
      Width           =   3135
   End
   Begin VB.TextBox txtCashierID 
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
      Left            =   1440
      Locked          =   -1  'True
      MaxLength       =   10
      TabIndex        =   2
      Top             =   900
      Width           =   3135
   End
   Begin VB.TextBox txtLoginName 
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
      Left            =   1455
      MaxLength       =   20
      TabIndex        =   1
      Top             =   1290
      Width           =   3135
   End
   Begin HSES.b8Line b8Line1 
      Height          =   60
      Left            =   0
      TabIndex        =   0
      Top             =   510
      Width           =   5385
      _ExtentX        =   9499
      _ExtentY        =   106
   End
   Begin HSES.b8Line b8Line2 
      Height          =   60
      Left            =   -15
      TabIndex        =   3
      Top             =   4440
      Width           =   5385
      _ExtentX        =   9499
      _ExtentY        =   106
   End
   Begin lvButton.lvButtons_H cmdSave 
      Default         =   -1  'True
      Height          =   360
      Left            =   3630
      TabIndex        =   19
      Top             =   4530
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
      Left            =   2100
      TabIndex        =   20
      Top             =   4530
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
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Address"
      Height          =   195
      Left            =   360
      TabIndex        =   18
      Top             =   3555
      Width           =   570
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Last Name"
      Height          =   195
      Left            =   360
      TabIndex        =   16
      Top             =   2790
      Width           =   765
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Contact No."
      Height          =   195
      Left            =   360
      TabIndex        =   15
      Top             =   3195
      Width           =   855
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "First Name"
      Height          =   195
      Left            =   360
      TabIndex        =   12
      Top             =   2070
      Width           =   750
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Middle Name"
      Height          =   195
      Left            =   360
      TabIndex        =   11
      Top             =   2430
      Width           =   930
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      Height          =   195
      Left            =   360
      TabIndex        =   8
      Top             =   1665
      Width           =   690
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Log-In Name"
      Height          =   195
      Left            =   360
      TabIndex        =   6
      Top             =   1305
      Width           =   915
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ID"
      Height          =   195
      Left            =   345
      TabIndex        =   5
      Top             =   930
      Width           =   165
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "New Cashier"
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
      TabIndex        =   4
      Top             =   180
      Width           =   1800
   End
   Begin VB.Image Image2 
      Height          =   525
      Left            =   0
      Picture         =   "frmAddCashier.frx":000C
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5295
   End
End
Attribute VB_Name = "frmAddCashier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim RecordSaved As Boolean

Public Function ShowForm() As Boolean
    
    '-------------------------------------------------------
    'check user access
    '-------------------------------------------------------
    If UserAllowedTo(CurrentUser.UserName, sCanAddCashier) = False Then
        MsgBox "Unable to continue adding Cashier entry." & vbNewLine & vbNewLine & _
                "You are not permitted to do this. Please contact your administrator for more information.", vbExclamation
        Exit Function
    End If
    '-------------------------------------------------------
    
    'generate code
    
    
    'show form
    Me.Show vbModal
    
    'return
    ShowForm = RecordSaved
End Function

Private Sub GenerateID()
    
    Dim lNewID As Long
    
    lNewID = GetNewCashierID()
    
    If lNewID < 1 Then
        'fatal error
        'temp
        MsgBox "error"
        Exit Sub
    End If
    
    txtCashierID.Text = String$(10 - Len(Trim(lNewID)), "0") & Trim(lNewID)
    
End Sub


Private Sub cmdCancel_Click()
    Unload Me
End Sub
Private Sub Form_Activate()
    
    GenerateID
    
End Sub

Private Sub cmdSave_Click()
    
    If ValidateData = False Then
        Exit Sub
    End If
    
    SaveData
End Sub

Private Function ValidateData() As Boolean
    
    'default
    ValidateData = False
    
    If Len(txtLoginName.Text) < 1 Then
        MsgBox "Please enter Login Name. This must must not be empty", vbExclamation
        HLTxt txtLoginName
        Exit Function
    End If
    
    If Len(txtPassword.Text) < 1 Then
        MsgBox "Please enter Password. This must must not be empty", vbExclamation
        HLTxt txtPassword
        Exit Function
    End If
    
    If Len(txtFirstName) < 1 Then
        MsgBox "Please enter First Name. This must must not be empty", vbExclamation
        HLTxt txtFirstName
        Exit Function
    End If
    
    If Len(txtMiddleName) < 1 Then
        MsgBox "Please enter Middle Name. This must must not be empty", vbExclamation
        HLTxt txtMiddleName
        Exit Function
    End If
    
    If Len(txtLastName) < 1 Then
        MsgBox "Please enter Last Name. This must must not be empty", vbExclamation
        HLTxt txtLastName
        Exit Function
    End If
    
    If Len(txtContactNumber) < 1 Then
        MsgBox "Please enter Contact Number. This must must not be empty", vbExclamation
        HLTxt txtContactNumber
        Exit Function
    End If
    
    If Len(txtAddress) < 1 Then
        MsgBox "Please enter Address. This must must not be empty", vbExclamation
        HLTxt txtAddress
        Exit Function
    End If
    
    
    'return success
    ValidateData = True
End Function


Private Sub SaveData()
    
    Select Case AddCashier(CLng(txtCashierID.Text), _
                            txtLoginName.Text, _
                            txtPassword.Text, _
                            txtFirstName.Text, _
                            txtMiddleName.Text, _
                            txtLastName.Text, _
                            txtAddress.Text, _
                            txtContactNumber.Text, _
                            Now, _
                            CurrentUser.UserName)
        Case TranDBResult.Success
            MsgBox "New Cashier entry successfully created.", vbInformation
            
            'set flag
            RecordSaved = True
            
            'close this form
            Unload Me
        Case TranDBResult.DuplicateID
            MsgBox "Unable to save Cashier entry. Duplicate Login Name found.", vbExclamation
            GenerateID
        Case TranDBResult.DuplicateLoginName
            MsgBox "Unable to save Cashier entry. Duplicate Login Name found.", vbExclamation
            HLTxt txtLoginName
        Case Else
            'fatal error
            'temp
            MsgBox "error"
                        
    End Select

End Sub
