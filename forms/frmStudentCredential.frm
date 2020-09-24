VERSION 5.00
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmStudentCredential 
   BackColor       =   &H00D8E9EC&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Student Credential"
   ClientHeight    =   3000
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   5100
   Icon            =   "frmStudentCredential.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   200
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   340
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdGetStudent 
      BackColor       =   &H00D8E9EC&
      Height          =   270
      Left            =   4230
      Picture         =   "frmStudentCredential.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   1320
      Width           =   345
   End
   Begin VB.CommandButton cmdGetCredential 
      BackColor       =   &H00D8E9EC&
      Height          =   270
      Left            =   4230
      Picture         =   "frmStudentCredential.frx":0E54
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   930
      Width           =   345
   End
   Begin VB.TextBox txtCredential 
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
      Left            =   1470
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   11
      Top             =   900
      Width           =   3135
   End
   Begin lvButton.lvButtons_H cmdHistory 
      Height          =   345
      Left            =   30
      TabIndex        =   10
      Top             =   2610
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   609
      Caption         =   "Entry History"
      CapAlign        =   2
      BackStyle       =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFore           =   8421504
      cFHover         =   8421504
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      Enabled         =   0   'False
      cBack           =   14215660
   End
   Begin VB.TextBox txtRemarks 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      IMEMode         =   3  'DISABLE
      Left            =   1455
      MaxLength       =   100
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   1665
      Width           =   3135
   End
   Begin VB.TextBox txtStudent 
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
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   1
      Top             =   1290
      Width           =   3135
   End
   Begin HSES.b8Line b8Line1 
      Height          =   60
      Left            =   0
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   510
      Width           =   5385
      _ExtentX        =   9499
      _ExtentY        =   106
   End
   Begin HSES.b8Line b8Line2 
      Height          =   60
      Left            =   -45
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   2520
      Width           =   5385
      _ExtentX        =   9499
      _ExtentY        =   106
   End
   Begin lvButton.lvButtons_H cmdSave 
      Default         =   -1  'True
      Height          =   360
      Left            =   3600
      TabIndex        =   4
      Top             =   2610
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
      Left            =   2070
      TabIndex        =   5
      Top             =   2610
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
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Remarks"
      Height          =   195
      Left            =   360
      TabIndex        =   2
      Top             =   1665
      Width           =   630
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Student"
      Height          =   195
      Left            =   360
      TabIndex        =   0
      Top             =   1305
      Width           =   555
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Credential"
      Height          =   195
      Left            =   345
      TabIndex        =   9
      Top             =   930
      Width           =   705
   End
   Begin VB.Label lblFormTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "New Student Credential"
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
      Left            =   90
      TabIndex        =   8
      Top             =   180
      Width           =   3420
   End
   Begin VB.Image Image2 
      Height          =   525
      Left            =   0
      Picture         =   "frmStudentCredential.frx":13DE
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5295
   End
End
Attribute VB_Name = "frmStudentCredential"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim RecordSaved As Boolean
Dim sFormState As String

Dim curCredentialID As String
Dim curStudentID As String

Dim pStudentID As String

Dim curCreationDate As Date
Dim curCreatedBy As String
Dim curModifiedDate As Date
Dim curModifiedBy As String


Public Function ShowAdd(Optional sStudentID As String = "") As Boolean
    
    '-------------------------------------------------------
    'check user access
    '-------------------------------------------------------
    If UserAllowedTo(CurrentUser.UserName, sCanAddCredential) = False Then
        MsgBox "Unable to continue adding Credential entry." & vbNewLine & vbNewLine & _
                "You are not permitted to do this. Please contact your administrator for more information.", vbExclamation
        Exit Function
    End If
    '-------------------------------------------------------
    
    'set ui & var
    pStudentID = sStudentID
    
    sFormState = "add"
    
    lblFormTitle = "New Credential"
    cmdSave.Caption = "&Save"
    
    
    
    
    'show form
    Me.Show vbModal
    
    'return
    ShowAdd = RecordSaved
End Function




Private Sub cmdCancel_Click()
    Unload Me
End Sub


Private Sub cmdGetCredential_Click()
    
    Dim sCredentialID As String
    Dim sCredentialTitle As String
    
    sCredentialID = PickCredential.GetCredentialID(txtCredential, , sCredentialTitle)
    
    If Len(sCredentialID) > 0 Then
        curCredentialID = sCredentialID
        txtCredential.Text = sCredentialTitle
    Else
        curCredentialID = ""
        txtCredential.Text = ""
    End If
End Sub

Private Sub cmdGetStudent_Click()
    
    Dim sStudentID As String
    Dim sStudentname As String
    
    sStudentID = PickStudent.GetStudentID(txtStudent, sStudentname)
    
    If Len(sStudentID) > 0 Then
        curStudentID = sStudentID
        txtStudent.Text = sStudentname
    Else
        curStudentID = ""
        txtStudent.Text = ""
    End If
End Sub

Private Sub cmdSave_Click()
    'check form state
    Select Case sFormState
    
        Case "add"
            SaveAdd
            

    End Select
End Sub

Private Sub Form_Activate()
    Dim vStudent As tStudent

    'check form state
    Select Case sFormState
    
        Case "add"
            SetUIAdd

    End Select
    
    If Len(pStudentID) > 0 Then
        If GetStudentByID(pStudentID, vStudent) = Success Then
            curStudentID = pStudentID
            txtStudent.Text = vStudent.LastName & ", " & vStudent.FirstName & " " & vStudent.MiddleName
        Else
            txtStudent.Text = ""
            curStudentID = ""
        End If
    
        pStudentID = ""
    End If
    
End Sub

Private Sub SetUIAdd()

    
    
    
End Sub






Private Sub SaveAdd()
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    
    'check
    If Len(txtStudent.Text) < 1 Then
        MsgBox "Please fill up Student.", vbExclamation
        HLTxt txtStudent
        GoTo ReleaseAndExit
    End If

    If Len(txtCredential.Text) < 1 Then
        MsgBox "Please fill up Credential.", vbExclamation
        HLTxt txtCredential
        GoTo ReleaseAndExit
    End If
    
    sSQL = "SELECT * FROM tblStudentCredential WHERE tblStudentCredential.CredentialID='" & curCredentialID & "' AND tblStudentCredential.StudentID='" & curStudentID & "'"
    
    If ConnectRS(HSESDB, vRS, sSQL) = False Then
        CatchError "frmStudentCredential", "SaveAdd", "Unable to connect Recordset with Sql Expression '" & sSQL & "'"
        MsgBox "Unable to connect Recordset.", vbExclamation
        GoTo ReleaseAndExit
    End If
    
    If AnyRecordExisted(vRS) = True Then
        MsgBox "The Student with the Credential that you have entered was already added records." & vbNewLine & _
            "Please enter different Student or Credential.", vbExclamation
        HLTxt txtCredential
        GoTo ReleaseAndExit
    End If
    
    
    
    'save record
    On Error GoTo ErrSaving
    
    vRS.AddNew
        
    vRS.Fields("CredentialID").Value = curCredentialID
    vRS.Fields("StudentID").Value = curStudentID
    vRS.Fields("Remarks").Value = txtRemarks.Text
    
    vRS.Fields("CreatedBy").Value = CurrentUser.UserName
    vRS.Fields("CreationDate").Value = Now
        
    vRS.Update
    
    'show success message
    MsgBox "New Student Credential entry successfully created.", vbInformation
    'set flag
    RecordSaved = True
    'close this form
    Unload Me
    
    
ReleaseAndExit:
    Set vRS = Nothing
    Exit Sub

ErrSaving:
    CatchError "frmStudentCredential", "SaveAdd", "Unable to update Record with Sql Expression '" & sSQL & "'"
    GoTo ReleaseAndExit
End Sub

