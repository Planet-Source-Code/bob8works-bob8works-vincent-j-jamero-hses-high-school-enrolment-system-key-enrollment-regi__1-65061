VERSION 5.00
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmAddSection 
   BackColor       =   &H00D8E9EC&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Section"
   ClientHeight    =   3060
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4965
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAddSection.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   204
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   331
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdGetItem 
      BackColor       =   &H00D8E9EC&
      Height          =   300
      Left            =   4350
      Picture         =   "frmAddSection.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   1650
      Width           =   345
   End
   Begin VB.CommandButton cmdGetYearLevelTitle 
      BackColor       =   &H00D8E9EC&
      Height          =   285
      Left            =   4350
      Picture         =   "frmAddSection.frx":0B14
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   2070
      Width           =   345
   End
   Begin VB.TextBox txtSectionID 
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
      Left            =   1500
      Locked          =   -1  'True
      MaxLength       =   10
      TabIndex        =   6
      Top             =   750
      Width           =   3225
   End
   Begin VB.TextBox txtSectionTitle 
      Height          =   315
      Left            =   1500
      MaxLength       =   50
      TabIndex        =   5
      Top             =   1230
      Width           =   3225
   End
   Begin VB.TextBox txtDepartmentTitle 
      Height          =   345
      Left            =   1500
      MaxLength       =   50
      TabIndex        =   4
      Top             =   1620
      Width           =   3225
   End
   Begin VB.TextBox txtYearLevelTitle 
      Height          =   345
      Left            =   1500
      MaxLength       =   20
      TabIndex        =   3
      Top             =   2040
      Width           =   3225
   End
   Begin HSES.b8Line b8Line1 
      Height          =   60
      Left            =   0
      TabIndex        =   0
      Top             =   510
      Width           =   5955
      _ExtentX        =   10504
      _ExtentY        =   106
   End
   Begin HSES.b8Line b8Line2 
      Height          =   60
      Left            =   0
      TabIndex        =   1
      Top             =   2580
      Width           =   5385
      _ExtentX        =   9499
      _ExtentY        =   106
   End
   Begin lvButton.lvButtons_H cmdSave 
      Default         =   -1  'True
      Height          =   360
      Left            =   3480
      TabIndex        =   13
      Top             =   2670
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
      Left            =   1950
      TabIndex        =   14
      Top             =   2670
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
      Caption         =   "Section Title"
      Height          =   195
      Left            =   240
      TabIndex        =   10
      Top             =   1275
      Width           =   870
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Departent"
      Height          =   195
      Left            =   270
      TabIndex        =   9
      Top             =   1680
      Width           =   735
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Year Level"
      Height          =   195
      Left            =   270
      TabIndex        =   8
      Top             =   2070
      Width           =   750
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Section ID"
      Height          =   195
      Left            =   240
      TabIndex        =   7
      Top             =   780
      Width           =   735
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Add New Section"
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
      TabIndex        =   2
      Top             =   180
      Width           =   2445
   End
   Begin VB.Image Image2 
      Height          =   525
      Left            =   0
      Picture         =   "frmAddSection.frx":109E
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5445
   End
End
Attribute VB_Name = "frmAddSection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Dim RecordAdded As Boolean

Public Function ShowForm(Optional sDepartmentTitle As String = "", Optional sYearLevelTitle As String = "", Optional sTeacherTitle As String = "") As Boolean
    
    Dim sNewSectionID As String
    
    
    '-------------------------------------------------------
    'check user access
    '-------------------------------------------------------
    If UserAllowedTo(CurrentUser.UserName, sCanAddSection) = False Then
        MsgBox "Unable to continue adding Section entry." & vbNewLine & vbNewLine & _
                "You are not permitted to do this. Please contact your administrator for more information.", vbExclamation
        Exit Function
    End If
    '-------------------------------------------------------



    'set defaults
    ShowForm = False
    
    If UserAllowedTo(CurrentUser.UserName, sCanAddSection) = False Then
        MsgBox "Unable to continue adding Section entry." & vbNewLine & vbNewLine & _
                "You are not permitted to do this. Please contact your administrator for more information.", vbExclamation
        Exit Function
    End If

    
    'check if other related recordset rentry exist
    If DepartmentRecordExist <> Success Then
        MsgBox "Unable to continue Adding Section." & vbNewLine & "Department entries not exist", vbExclamation
        Unload Me
        Exit Function
    End If
    
    If YearLevelRecordExist <> Success Then
        MsgBox "Unable to continue Adding Section." & vbNewLine & "Year Level entries not exist", vbExclamation
        Unload Me
        Exit Function
    End If
    
    
    'set text fields
    txtDepartmentTitle.Text = sDepartmentTitle
    txtYearLevelTitle.Text = sYearLevelTitle
    
    'generate new id
    If GetNewSectionID(sNewSectionID) = Success Then
        txtSectionID.Text = sNewSectionID
    Else
    End If
    
    'show form
    Me.Show vbModal
    
    'return
    ShowForm = RecordAdded
End Function






Private Function SaveData() As Boolean
    
    Dim newSection As tSection
    Dim vDepartment As tDepartment
    Dim vYearlevel As tYearLevel
    Dim vTeacher As tTeacher
    
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
    newSection.SectionID = txtSectionID
    newSection.SectionTitle = txtSectionTitle
    newSection.DepartmentID = vDepartment.DepartmentID
    newSection.YearLevelID = vYearlevel.YearLevelID
    newSection.CreationDate = Now
    newSection.CreatedBy = CurrentUser.UserName
    
    'try


    Dim ir As Integer
    ir = AddSection(newSection)
    Select Case ir
        Case TranDBResult.Success
            'success
            '-------------------------------------
            'section successfully saved
            'return success
            SaveData = True
        
        
        Case TranDBResult.DuplicateID
            MsgBox "ID already existed.", vbExclamation
            HLTxt txtSectionID
            SaveData = False
        
        Case TranDBResult.DuplicateTitle
            MsgBox "Title already existed.", vbExclamation
            HLTxt txtSectionTitle
            SaveData = False

        Case Else
            CatchError "frmAddSetion", "SaveData", "Saving Section"
            SaveData = False
    End Select
End Function



Private Function ValidateData() As Boolean
    
    'default
    ValidateData = False
    
    'check id
    If Not CheckTextBox(txtSectionID, "Please Enter Section ID") Then
        Exit Function
    End If
    
    'check title
    If Not CheckTextBox(txtSectionTitle, "Please Enter Section Title") Then
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

    
    'return success
    ValidateData = True
End Function




Private Sub cmdCancel_Click()
    Unload Me
End Sub




Private Sub cmdGetItem_Click()
    Dim sDepartmentID As String
    Dim sDepartmentTitle As String

    sDepartmentID = PickDepartment.GetItem(txtDepartmentTitle, sDepartmentTitle)
    If sDepartmentID <> "" Then
        txtDepartmentTitle = sDepartmentTitle
    End If
End Sub





Private Sub cmdGetDepartmentTitle_Click()

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
    
        MsgBox "SECTION Entry successfully added.", vbInformation
        RecordAdded = True
        
        
        If MsgBox("Do you want to offer this section for enrolment?", vbQuestion + vbYesNo) = vbYes Then
            frmAddSectionOffering.ShowForm txtYearLevelTitle.Text & " - " & txtSectionTitle.Text
        End If
        
        Unload Me
    End If
End Sub





