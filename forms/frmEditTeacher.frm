VERSION 5.00
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmEditTeacher 
   BackColor       =   &H00D8E9EC&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Edit Teacher Account"
   ClientHeight    =   5250
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5205
   Icon            =   "frmEditTeacher.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   350
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   347
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtTeacherID 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1650
      Locked          =   -1  'True
      MaxLength       =   10
      TabIndex        =   7
      Top             =   840
      Width           =   3225
   End
   Begin VB.TextBox txtTeacherTitle 
      Height          =   345
      Left            =   1650
      MaxLength       =   20
      TabIndex        =   6
      Top             =   1335
      Width           =   3225
   End
   Begin VB.TextBox txtPassword 
      Height          =   345
      Left            =   1650
      MaxLength       =   20
      TabIndex        =   5
      Top             =   1755
      Width           =   3225
   End
   Begin VB.TextBox txtFirstName 
      Height          =   345
      Left            =   1650
      MaxLength       =   50
      TabIndex        =   4
      Top             =   2190
      Width           =   3225
   End
   Begin VB.TextBox txtMiddleName 
      Height          =   345
      Left            =   1650
      MaxLength       =   50
      TabIndex        =   3
      Top             =   2610
      Width           =   3225
   End
   Begin VB.TextBox txtLastName 
      Height          =   345
      Left            =   1650
      MaxLength       =   50
      TabIndex        =   2
      Top             =   3030
      Width           =   3225
   End
   Begin VB.TextBox txtAddress 
      Height          =   675
      Left            =   1650
      MaxLength       =   70
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   3450
      Width           =   3225
   End
   Begin VB.TextBox txtContactNumber 
      Height          =   345
      Left            =   1650
      MaxLength       =   20
      TabIndex        =   0
      Top             =   4215
      Width           =   3225
   End
   Begin HSES.b8Line b8Line1 
      Height          =   60
      Left            =   60
      TabIndex        =   16
      Top             =   510
      Width           =   5955
      _ExtentX        =   10504
      _ExtentY        =   106
   End
   Begin HSES.b8Line b8Line2 
      Height          =   60
      Left            =   0
      TabIndex        =   18
      Top             =   4770
      Width           =   5385
      _ExtentX        =   9499
      _ExtentY        =   106
   End
   Begin lvButton.lvButtons_H cmdUpdate 
      Default         =   -1  'True
      Height          =   360
      Left            =   3735
      TabIndex        =   19
      Top             =   4860
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   635
      Caption         =   "&Update"
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
      Left            =   2205
      TabIndex        =   20
      Top             =   4860
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
      Caption         =   "Add NewTeacher"
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
      Left            =   150
      TabIndex        =   17
      Top             =   180
      Width           =   2415
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ID"
      Height          =   195
      Left            =   300
      TabIndex        =   15
      Top             =   900
      Width           =   165
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Log-In Name"
      Height          =   195
      Left            =   300
      TabIndex        =   14
      Top             =   1380
      Width           =   915
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      Height          =   195
      Left            =   300
      TabIndex        =   13
      Top             =   1830
      Width           =   690
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "First Name"
      Height          =   195
      Left            =   300
      TabIndex        =   12
      Top             =   2250
      Width           =   765
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Middle Name"
      Height          =   195
      Left            =   300
      TabIndex        =   11
      Top             =   2700
      Width           =   900
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Last Name"
      Height          =   195
      Left            =   300
      TabIndex        =   10
      Top             =   3060
      Width           =   750
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Address"
      Height          =   195
      Left            =   300
      TabIndex        =   9
      Top             =   3510
      Width           =   585
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Contact Number"
      Height          =   195
      Left            =   270
      TabIndex        =   8
      Top             =   4290
      Width           =   1170
   End
   Begin VB.Image Image2 
      Height          =   525
      Left            =   15
      Picture         =   "frmEditTeacher.frx":08CA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5445
   End
End
Attribute VB_Name = "frmEditTeacher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Dim CurrentTeacher As tTeacher

Public Sub ShowEdit(sTeacherID As String)
    
    If GetTeacherByID(sTeacherID, CurrentTeacher) <> Success Then
        MsgBox "Unable to continue proccess." & vbNewLine & "The selected TEACHER entry was not found in record.", vbExclamation
        Unload Me
        Exit Sub
    End If
    
    Call SetTextField
    
   
    Me.Show vbModal
        
End Sub




Private Function SetTextField()
    txtTeacherID = CurrentTeacher.TeacherID
    txtTeacherTitle = CurrentTeacher.TeacherTitle
    txtPassword = CurrentTeacher.Password
    txtFirstName = CurrentTeacher.FirstName
    txtMiddleName = CurrentTeacher.MiddleName
    txtLastName = CurrentTeacher.LastName
    txtAddress = CurrentTeacher.Address
    txtContactNumber = CurrentTeacher.ContactNumber
    'txtCreationDate = CurrentTeacher.CreationDate
End Function




Private Function ValidateData() As Boolean
    
    'set default fail
    ValidateData = False
    
    If Not CheckTextBox(txtTeacherID, "Teacher ID") Then Exit Function
    If Not CheckTextBox(txtTeacherTitle, "Teacher Name") Then Exit Function
    If Not CheckTextBox(txtPassword, "Password") Then Exit Function
    If Not CheckTextBox(txtFirstName, "First Name") Then Exit Function
    If Not CheckTextBox(txtMiddleName, "MiddleName") Then Exit Function
    If Not CheckTextBox(txtLastName, "LastName") Then Exit Function
    If Not CheckTextBox(txtAddress, "Address") Then Exit Function
    If Not CheckTextBox(txtContactNumber, "ContactNumber") Then Exit Function
        
    ValidateData = True
End Function


Private Function SaveData() As Boolean
    
    'default
    SaveData = False
        
    If Not SetRSField Then
        SaveData = False
        Exit Function
    End If
    

    'save

    
        Select Case EditTeacher(CurrentTeacher)
            
            Case TranDBResult.Success
                'success
                '----------------------------
                MsgBox "TEACHER entry successfull edited.", vbInformation
                Unload Me
                '----------------------------
                
            Case TranDBResult.DuplicateID
                MsgBox "The TEACHER ID you have entered is already existed." & vbNewLine & "Please enter a different value.", vbExclamation
                txtTeacherID.SetFocus
                
            Case TranDBResult.DuplicateTitle
                MsgBox "The TEACHER TITLE you have entered is already existed." & vbNewLine & "Please enter a different value.", vbExclamation
                txtTeacherTitle.SetFocus
                
                
            Case TranDBResult.InvalidTeacherTitle
                MsgBox "Invalid TEACHER TITLE.", vbExclamation
                txtTeacherTitle.SetFocus
        

            Case TranDBResult.InvalidTeacherPassword
                MsgBox "Invalid PASSWORD.", vbExclamation
                txtPassword.SetFocus
        

            Case TranDBResult.InvalidTeacherFirstName
                MsgBox "Invalid FIRST NAME.", vbExclamation
                txtFirstName.SetFocus
        

            Case TranDBResult.InvalidTeacherMiddleName
                MsgBox "Invalid MIDDLE NAME.", vbExclamation
                txtMiddleName.SetFocus
        

            Case TranDBResult.InvalidTeacherLastName
                MsgBox "Invalid LAST NAME.", vbExclamation
                txtLastName.SetFocus
        

            Case TranDBResult.InvalidTeacherContactNumber
                MsgBox "Invalid CONTACT NUMBER.", vbExclamation
                txtContactNumber.SetFocus
                
                
            Case TranDBResult.InvalidTeacherAddress
                MsgBox "Invalid ADDRESS.", vbExclamation
                txtAddress.SetFocus

        End Select
    
    
    
End Function 'SaveData---------------------------

Private Function SetRSField() As Boolean
    
    SetRSField = False
    
    If Not ValidateData Then Exit Function
    
    'set data
    CurrentTeacher.TeacherID = Trim(txtTeacherID)
    CurrentTeacher.TeacherTitle = Trim(txtTeacherTitle)
    CurrentTeacher.Password = Trim(txtPassword)
    CurrentTeacher.FirstName = Trim(txtFirstName)
    CurrentTeacher.MiddleName = Trim(txtMiddleName)
    CurrentTeacher.LastName = Trim(txtLastName)
    CurrentTeacher.Address = Trim(txtAddress)
    CurrentTeacher.ContactNumber = Trim(txtContactNumber)
    
    
    SetRSField = True
End Function


Private Sub ExecEditTeacher()
    If SaveData Then
        'added
        Unload Me
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub


Private Sub cmdUpdate_Click()
    ExecEditTeacher
End Sub



