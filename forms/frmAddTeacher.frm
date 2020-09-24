VERSION 5.00
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmAddTeacher 
   Appearance      =   0  'Flat
   BackColor       =   &H00D8E9EC&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Teacher"
   ClientHeight    =   5265
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   5055
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAddTeacher.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   351
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   337
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtContactNumber 
      Height          =   345
      Left            =   1560
      MaxLength       =   20
      TabIndex        =   11
      Top             =   4215
      Width           =   3225
   End
   Begin VB.TextBox txtAddress 
      Height          =   675
      Left            =   1560
      MaxLength       =   70
      MultiLine       =   -1  'True
      TabIndex        =   10
      Top             =   3480
      Width           =   3225
   End
   Begin VB.TextBox txtLastName 
      Height          =   345
      Left            =   1560
      MaxLength       =   50
      TabIndex        =   9
      Top             =   3060
      Width           =   3225
   End
   Begin VB.TextBox txtMiddleName 
      Height          =   345
      Left            =   1560
      MaxLength       =   50
      TabIndex        =   8
      Top             =   2640
      Width           =   3225
   End
   Begin VB.TextBox txtFirstName 
      Height          =   345
      Left            =   1560
      MaxLength       =   50
      TabIndex        =   7
      Top             =   2220
      Width           =   3225
   End
   Begin VB.TextBox txtPassword 
      Height          =   345
      Left            =   1560
      MaxLength       =   20
      TabIndex        =   6
      Top             =   1785
      Width           =   3225
   End
   Begin VB.TextBox txtTeacherTitle 
      Height          =   345
      Left            =   1560
      MaxLength       =   20
      TabIndex        =   5
      Top             =   1365
      Width           =   3225
   End
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
      Left            =   1560
      Locked          =   -1  'True
      MaxLength       =   10
      TabIndex        =   4
      Top             =   810
      Width           =   3225
   End
   Begin VB.CommandButton cmdGetYearLevelTitle 
      BackColor       =   &H00D8E9EC&
      Height          =   330
      Left            =   6960
      Picture         =   "frmAddTeacher.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2865
      Width           =   345
   End
   Begin VB.CommandButton cmdGetDepartmentTitle 
      BackColor       =   &H00D8E9EC&
      Height          =   330
      Left            =   6960
      Picture         =   "frmAddTeacher.frx":0E54
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2415
      Width           =   345
   End
   Begin HSES.b8Line b8Line1 
      Height          =   60
      Left            =   60
      TabIndex        =   0
      Top             =   510
      Width           =   5925
      _ExtentX        =   10451
      _ExtentY        =   106
   End
   Begin HSES.b8Line b8Line2 
      Height          =   60
      Left            =   0
      TabIndex        =   20
      Top             =   4770
      Width           =   5385
      _ExtentX        =   9499
      _ExtentY        =   106
   End
   Begin lvButton.lvButtons_H cmdSave 
      Default         =   -1  'True
      Height          =   360
      Left            =   3645
      TabIndex        =   21
      Top             =   4860
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
      Left            =   2115
      TabIndex        =   22
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
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Contact Number"
      Height          =   195
      Left            =   180
      TabIndex        =   19
      Top             =   4290
      Width           =   1170
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Address"
      Height          =   195
      Left            =   210
      TabIndex        =   18
      Top             =   3540
      Width           =   585
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Last Name"
      Height          =   195
      Left            =   210
      TabIndex        =   17
      Top             =   3090
      Width           =   750
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Middle Name"
      Height          =   195
      Left            =   210
      TabIndex        =   16
      Top             =   2730
      Width           =   900
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "First Name"
      Height          =   195
      Left            =   210
      TabIndex        =   15
      Top             =   2280
      Width           =   765
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      Height          =   195
      Left            =   210
      TabIndex        =   14
      Top             =   1860
      Width           =   690
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Log-In Name"
      Height          =   195
      Left            =   210
      TabIndex        =   13
      Top             =   1410
      Width           =   915
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ID"
      Height          =   195
      Left            =   210
      TabIndex        =   12
      Top             =   870
      Width           =   165
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
      Left            =   60
      TabIndex        =   1
      Top             =   180
      Width           =   2415
   End
   Begin VB.Image Image2 
      Height          =   525
      Left            =   0
      Picture         =   "frmAddTeacher.frx":13DE
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5445
   End
End
Attribute VB_Name = "frmAddTeacher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim RecordAdded As Boolean

'START ---------------------------------------
Public Function ShowForm() As Boolean
    
    '-------------------------------------------------------
    'check user access
    '-------------------------------------------------------
    If UserAllowedTo(CurrentUser.UserName, sCanAddTeacher) = False Then
        MsgBox "Unable to continue Adding Teacher entry." & vbNewLine & vbNewLine & _
                "You are not permitted to do this. Please contact your administrator for more information.", vbExclamation
        Exit Function
    End If
    '-------------------------------------------------------

    
    Me.Show vbModal
    
    
End Function
'ShowForm -----------------------------



Private Function ValidateData() As Boolean
    
    'set default failed
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
    
    Dim newTeacher As tTeacher

    
    'set failed
    SaveData = False

    If Not ValidateData Then Exit Function
    
    
    'set data
    newTeacher.TeacherID = Trim(txtTeacherID)
    newTeacher.TeacherTitle = Trim(txtTeacherTitle)
    newTeacher.Password = Trim(txtPassword)
    newTeacher.FirstName = Trim(txtFirstName)
    newTeacher.MiddleName = Trim(txtMiddleName)
    newTeacher.LastName = Trim(txtLastName)
    newTeacher.Address = Trim(txtAddress)
    newTeacher.ContactNumber = Trim(txtContactNumber)
    newTeacher.CreationDate = Now
    
    
    Select Case AddTeacher(newTeacher)
            
            Case TranDBResult.Success
                'success
                '----------------------------
                MsgBox "TEACHER entry successfull Added.", vbInformation
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
            Case Else
                'temp
                'fatal
                MsgBox "Unknown result: Adding Teacher entry", vbCritical
        End Select
    
    
End Function


Private Sub cmdCancel_Click()
    Unload Me
End Sub


Private Sub cmdSave_Click()
     If SaveData Then
        'added
        RecordAdded = True
        Unload Me
    End If
End Sub



Private Sub Form_Activate()
    Dim sNewID As String
    
    
    'generate id
    GetNewTeacherID sNewID
    
    txtTeacherID.Text = sNewID
End Sub

