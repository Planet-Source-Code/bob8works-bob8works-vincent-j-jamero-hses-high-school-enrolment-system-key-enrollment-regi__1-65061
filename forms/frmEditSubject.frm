VERSION 5.00
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmEditSubject 
   BackColor       =   &H00D8E9EC&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Subject"
   ClientHeight    =   4620
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5310
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmEditSubject.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4620
   ScaleWidth      =   5310
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtDescription 
      Height          =   675
      Left            =   1530
      MaxLength       =   100
      MultiLine       =   -1  'True
      TabIndex        =   4
      Top             =   1890
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
      Left            =   1530
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   3
      Top             =   900
      Width           =   3225
   End
   Begin VB.TextBox txtSubjectTitle 
      Height          =   345
      Left            =   1530
      MaxLength       =   50
      TabIndex        =   2
      Top             =   1380
      Width           =   3225
   End
   Begin VB.TextBox txtDepartmentTitle 
      BackColor       =   &H00D8E9EC&
      Height          =   345
      Left            =   1530
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   2730
      Width           =   3225
   End
   Begin VB.TextBox txtYearLevelTitle 
      BackColor       =   &H00D8E9EC&
      Height          =   345
      Left            =   1530
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   3240
      Width           =   3225
   End
   Begin HSES.b8Line b8Line1 
      Height          =   60
      Left            =   30
      TabIndex        =   5
      Top             =   510
      Width           =   5925
      _ExtentX        =   10451
      _ExtentY        =   106
   End
   Begin lvButton.lvButtons_H cmdSave 
      Default         =   -1  'True
      Height          =   360
      Left            =   3720
      TabIndex        =   6
      Top             =   4140
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   635
      Caption         =   "&Save"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdCancel 
      Cancel          =   -1  'True
      Height          =   360
      Left            =   2070
      TabIndex        =   7
      Top             =   4140
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   635
      Caption         =   "&Cancel"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin HSES.b8Line b8Line2 
      Height          =   60
      Left            =   0
      TabIndex        =   8
      Top             =   3990
      Width           =   5955
      _ExtentX        =   10504
      _ExtentY        =   106
   End
   Begin lvButton.lvButtons_H cmdGetYearLevelTitle 
      Height          =   375
      Left            =   4770
      TabIndex        =   9
      Top             =   3240
      Width           =   405
      _ExtentX        =   714
      _ExtentY        =   661
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      Image           =   "frmEditSubject.frx":058A
      Enabled         =   0   'False
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdGetDepartmentTitle 
      Height          =   375
      Left            =   4770
      TabIndex        =   10
      Top             =   2730
      Width           =   405
      _ExtentX        =   714
      _ExtentY        =   661
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      Image           =   "frmEditSubject.frx":0B24
      Enabled         =   0   'False
      cBack           =   -2147483633
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Edit Subject Entry"
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
      TabIndex        =   16
      Top             =   150
      Width           =   2595
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      Height          =   195
      Left            =   300
      TabIndex        =   15
      Top             =   1935
      Width           =   795
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Subject Title"
      Height          =   195
      Left            =   300
      TabIndex        =   14
      Top             =   1425
      Width           =   885
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Departent"
      Height          =   195
      Left            =   300
      TabIndex        =   13
      Top             =   2730
      Width           =   735
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Year Level"
      Height          =   195
      Left            =   300
      TabIndex        =   12
      Top             =   3240
      Width           =   900
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Subject ID"
      Height          =   195
      Left            =   300
      TabIndex        =   11
      Top             =   930
      Width           =   750
   End
   Begin VB.Image Image2 
      Height          =   525
      Left            =   30
      Picture         =   "frmEditSubject.frx":10BE
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5925
   End
End
Attribute VB_Name = "frmEditSubject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Dim RecordEdited As Boolean

Dim CurrentSubject As tSubject


Public Function ShowEdit(sSubjectID As String) As Boolean
    Dim vDepartment As tDepartment
    Dim vYearlevel As tYearLevel
    
    'set defaults
    ShowEdit = False
    
    'get subject
    If GetSubjectByID(sSubjectID, CurrentSubject) = Success Then
        
        'set text fields
        txtSubjectID = CurrentSubject.SubjectID
        txtSubjectTitle = CurrentSubject.SubjectTitle
        txtDescription = CurrentSubject.Description
        
        If GetDepartmentByID(CurrentSubject.DepartmentID, vDepartment) = Success Then
            txtDepartmentTitle = vDepartment.DepartmentTitle
        End If
        
        If GetYearLevelByID(CurrentSubject.YearLevelID, vYearlevel) = Success Then
            txtYearLevelTitle = vYearlevel.YearLevelTitle
        End If
        
    Else
        
        Unload Me
        Exit Function
    
    End If
    'show form
    Me.Show vbModal
    
    'return
    ShowEdit = RecordEdited
End Function






Private Function SaveData() As Boolean
    
    Dim newSubject As tSubject
    Dim vDepartment As tDepartment
    Dim vYearlevel As tYearLevel
    
    'set default
    SaveData = False
    
    'validate date
    If Not ValidateData Then Exit Function
    
    'check id, it must be existed
    If SubjectExistByID(LCase(Trim(txtSubjectID))) <> Success Then
        MsgBox "ID not found.", vbExclamation
        HLTxt txtSubjectID
        Exit Function
    End If
    
    'check dulicate title
    If LCase(Trim(txtSubjectTitle)) <> CurrentSubject.SubjectTitle Then
        If SubjectExistByTitle(LCase(Trim(txtSubjectTitle))) = Success Then
            MsgBox "Title already existed.", vbExclamation
            HLTxt txtSubjectTitle
            Exit Function
        End If
    End If
    
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
    


    Select Case EditSubject(newSubject)
        Case TranDBResult.Success
            'success
            '-------------------------------------
            'Subject successfully saved
            'return success
            SaveData = True
        
    
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
    Dim sDepartmentTitle As String

    sDepartmentTitle = PickDepartment.GetItem
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
    If SaveData Then 'Edited
        MsgBox "Subject Entry successfully Edited.", vbInformation
        RecordEdited = True
        Unload Me
    End If
End Sub
