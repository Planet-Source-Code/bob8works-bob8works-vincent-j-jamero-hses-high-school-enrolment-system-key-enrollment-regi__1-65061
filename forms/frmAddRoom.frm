VERSION 5.00
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmAddRoom 
   BackColor       =   &H00D8E9EC&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Subject"
   ClientHeight    =   3915
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4995
   Icon            =   "frmAddRoom.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   261
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   333
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdGetYearLevelTitle 
      BackColor       =   &H00D8E9EC&
      Height          =   285
      Left            =   4320
      Picture         =   "frmAddRoom.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   2880
      Width           =   345
   End
   Begin VB.CommandButton cmdGetDepartmentTitle 
      BackColor       =   &H00D8E9EC&
      Height          =   300
      Left            =   4320
      Picture         =   "frmAddRoom.frx":0B14
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   2460
      Width           =   345
   End
   Begin VB.TextBox txtYearLevelTitle 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1470
      TabIndex        =   9
      Top             =   2850
      Width           =   3225
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
      Height          =   345
      Left            =   1470
      TabIndex        =   8
      Top             =   2430
      Width           =   3225
   End
   Begin VB.TextBox txtSubjectTitle 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1470
      MaxLength       =   50
      TabIndex        =   7
      Top             =   1260
      Width           =   3225
   End
   Begin VB.TextBox txtSubjectID 
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
      Height          =   345
      Left            =   1470
      MaxLength       =   20
      TabIndex        =   6
      Top             =   780
      Width           =   3225
   End
   Begin VB.TextBox txtDescription 
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
      Left            =   1470
      MaxLength       =   100
      MultiLine       =   -1  'True
      TabIndex        =   5
      Top             =   1680
      Width           =   3225
   End
   Begin HSES.b8Line b8Line1 
      Height          =   60
      Left            =   30
      TabIndex        =   0
      Top             =   510
      Width           =   5925
      _extentx        =   10451
      _extenty        =   106
   End
   Begin lvButton.lvButtons_H cmdSave 
      Default         =   -1  'True
      Height          =   360
      Left            =   3510
      TabIndex        =   1
      Top             =   3510
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
      cGradient       =   14215660
      Gradient        =   3
      Mode            =   0
      Value           =   0   'False
      cBack           =   16777215
   End
   Begin lvButton.lvButtons_H cmdCancel 
      Cancel          =   -1  'True
      Height          =   360
      Left            =   2040
      TabIndex        =   2
      Top             =   3510
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
   Begin HSES.b8Line b8Line2 
      Height          =   60
      Left            =   0
      TabIndex        =   3
      Top             =   3420
      Width           =   5955
      _extentx        =   10504
      _extenty        =   106
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Subject ID"
      Height          =   195
      Left            =   240
      TabIndex        =   14
      Top             =   810
      Width           =   750
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Year Level"
      Height          =   195
      Left            =   240
      TabIndex        =   13
      Top             =   2850
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Departent"
      Height          =   195
      Left            =   240
      TabIndex        =   12
      Top             =   2430
      Width           =   735
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Subject Title"
      Height          =   195
      Left            =   240
      TabIndex        =   11
      Top             =   1305
      Width           =   885
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      Height          =   195
      Left            =   240
      TabIndex        =   10
      Top             =   1725
      Width           =   795
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Add New Subject"
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
      TabIndex        =   4
      Top             =   150
      Width           =   2460
   End
   Begin VB.Image Image2 
      Height          =   525
      Left            =   30
      Picture         =   "frmAddRoom.frx":109E
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5925
   End
End
Attribute VB_Name = "frmAddRoom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Dim RecordAdded As Boolean

Public Function ShowForm(Optional sDepartmentTitle As String = "", Optional sYearLevelTitle As String = "", Optional sTeacherTitle As String = "") As Boolean
    
    Dim sNewSubjectID As String
    
    'set defaults
    ShowForm = False
    RecordAdded = False
    
    
    'check if other related recordset rentry exist
    If DepartmentRecordExist <> Success Then
        MsgBox "Unable to continue Adding Subject." & vbNewLine & "Department entries not exist", vbExclamation
        Unload Me
        Exit Function
    End If
    
    If YearLevelRecordExist <> Success Then
        MsgBox "Unable to continue Adding Subject." & vbNewLine & "Year Level entries not exist", vbExclamation
        Unload Me
        Exit Function
    End If
    
    
    'set text fields
    txtDepartmentTitle.Text = sDepartmentTitle
    txtYearLevelTitle.Text = sYearLevelTitle
    If GetNewSubjectID(sNewSubjectID) = Success Then
        txtSubjectID.Text = sNewSubjectID
    End If

    
    'show form
    Me.Show vbModal
    
    'return
    ShowForm = RecordAdded
End Function






Private Function SaveData() As Boolean
    
    Dim newSubject As tSubject
    Dim vDepartment As tDepartment
    Dim vYearlevel As tYearLevel
    
    'set default
    SaveData = False
    
    'validate date
    If Not ValidateData Then Exit Function
    
    
    
    'set/check departmentid
    If GetDepartmentByTitle(txtDepartmentTitle.Text, vDepartment) <> Success Then
        MsgBox "Invalid Department Title", vbExclamation
        HLTxt txtDepartmentTitle
        Exit Function
    End If
    
    'set/check year level
    If GetYearLevelbyTitle(txtYearLevelTitle, vYearlevel) <> Success Then
            MsgBox "Year Level Title not found", vbExclamation
            HLTxt txtYearLevelTitle
            Exit Function
    End If
    
    
    'set rs field
    newSubject.SubjectID = txtSubjectID
    newSubject.SubjectTitle = txtSubjectTitle
    newSubject.DepartmentID = vDepartment.DepartmentID
    newSubject.YearLevelID = vYearlevel.YearLevelID
    newSubject.Description = txtDescription
    'try
    


    Select Case AddSubject(newSubject)
        Case TranDBResult.Success
            'success
            '-------------------------------------
            'Subject successfully saved
            'return success
            SaveData = True

        
        
        Case TranDBResult.DuplicateID
            MsgBox "ID already existed.", vbExclamation
            HLTxt txtSubjectID
            SaveData = False
        
        Case TranDBResult.DuplicateTitle
            MsgBox "Title already existed.", vbExclamation
            HLTxt txtSubjectTitle
            SaveData = False
            
        Case TranDBResult.InvalidSubjectDepartmentID
            MsgBox "Invalid Department.", vbExclamation
            HLTxt txtDepartmentTitle
            SaveData = False
            
        Case TranDBResult.InvalidSubjectYearLevelID
            MsgBox "Invalid Year Level.", vbExclamation
            HLTxt txtYearLevelTitle
            SaveData = False
        
        Case TranDBResult.InvalidSubjectDescription
            MsgBox "Invalid Description.", vbExclamation
            HLTxt txtDescription
            SaveData = False
            
        Case Else
            'fatal
            'temp
            MsgBox "Unknown Error.", vbExclamation
            SaveData = False
    End Select
End Function



Private Function ValidateData() As Boolean
    
    'default
    ValidateData = False
    
    'check id
    If Not CheckTextBox(txtSubjectID, "Please Enter Subject ID") Then
        Exit Function
    End If
    
    'check title
    If Not CheckTextBox(txtSubjectTitle, "Please Enter Subject Title") Then
        Exit Function
    End If
    
    'check title
    If Not CheckTextBox(txtDescription, "Please Enter Description") Then
        Exit Function
    End If
    
    'check departmentid
    If DepartmentExistByTitle(txtDepartmentTitle.Text) <> Success Then
        MsgBox "Invalid Department Title", vbExclamation
        HLTxt txtDepartmentTitle
        Exit Function
    End If
    
    
    'check year level

    If YearLevelExistByTitle(txtYearLevelTitle) <> Success Then
            MsgBox "Year Level Title not found", vbExclamation
            HLTxt txtYearLevelTitle
            Exit Function
    End If
    
    
    'check Description
    If Not CheckTextBox(txtDescription, "Please Enter Description.") Then
            Exit Function
    End If
    

    
    
    'return success
    ValidateData = True
End Function




Private Sub cmdCancel_Click()
    Unload Me
End Sub




Private Sub cmdGetItem_Click()
    
End Sub


Private Sub cmdGetDepartmentTitle_Click()
    Dim sDepartmentTitle As String

     PickDepartment.GetItem txtDepartmentTitle, sDepartmentTitle
    If sDepartmentTitle <> "" Then
        txtDepartmentTitle = sDepartmentTitle
    End If
End Sub

Private Sub cmdGetYearLevelTitle_Click()
    Dim sYearLevelTitle As String
    
    sYearLevelTitle = PickYearLevel.GetYearLevelTitle
    If sYearLevelTitle <> "" Then
        txtYearLevelTitle.Text = sYearLevelTitle
    End If
End Sub




Private Sub cmdSave_Click()
    If SaveData Then 'added
        MsgBox "Subject Entry successfully added.", vbInformation
        RecordAdded = True
        Unload Me
    End If
End Sub
