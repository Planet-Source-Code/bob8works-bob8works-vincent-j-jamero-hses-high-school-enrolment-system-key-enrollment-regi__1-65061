VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmStudentGrade 
   BackColor       =   &H00D8E9EC&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Grades"
   ClientHeight    =   6585
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8910
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   439
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   594
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin lvButton.lvButtons_H cmdEditGrade 
      Height          =   345
      Left            =   90
      TabIndex        =   11
      Top             =   6060
      Width           =   1830
      _ExtentX        =   3228
      _ExtentY        =   609
      Caption         =   "&Edit Selected Grade"
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
      cBack           =   -2147483633
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00D8E9EC&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   465
      Left            =   0
      ScaleHeight     =   465
      ScaleWidth      =   8955
      TabIndex        =   2
      Top             =   510
      Width           =   8955
      Begin VB.ComboBox cmbYearLevelTitle 
         Height          =   315
         Left            =   4035
         TabIndex        =   7
         Top             =   90
         Width           =   1995
      End
      Begin VB.CommandButton cmbGetStudentID 
         BackColor       =   &H00D8E9EC&
         Height          =   315
         Left            =   2370
         Picture         =   "frmStudentGrade.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   90
         Width           =   345
      End
      Begin VB.TextBox txtStudentID 
         Height          =   315
         Left            =   960
         TabIndex        =   3
         Text            =   "Enter Student ID"
         Top             =   90
         Width           =   1395
      End
      Begin HSES.b8Line b8Line2 
         Height          =   60
         Left            =   0
         TabIndex        =   24
         Top             =   0
         Width           =   8865
         _ExtentX        =   15637
         _ExtentY        =   106
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Year Level"
         Height          =   195
         Left            =   3135
         TabIndex        =   8
         Top             =   120
         Width           =   750
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Student ID"
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   120
         Width           =   780
      End
      Begin VB.Image Image2 
         Height          =   135
         Left            =   0
         Picture         =   "frmStudentGrade.frx":058A
         Stretch         =   -1  'True
         Top             =   0
         Width           =   8925
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00F6F8F8&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   525
      Left            =   0
      ScaleHeight     =   525
      ScaleWidth      =   8955
      TabIndex        =   1
      Top             =   0
      Width           =   8955
      Begin VB.Image Image1 
         Height          =   405
         Left            =   30
         Picture         =   "frmStudentGrade.frx":0627
         Stretch         =   -1  'True
         Top             =   120
         Width           =   8925
      End
   End
   Begin MSComctlLib.ListView listRecord 
      Height          =   4185
      Left            =   120
      TabIndex        =   6
      Top             =   1770
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   7382
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   8388608
      BackColor       =   -2147483643
      Appearance      =   0
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Grade"
         Object.Width           =   1323
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Subject"
         Object.Width           =   3598
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "School Year"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Department"
         Object.Width           =   3969
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "YL"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Section"
         Object.Width           =   3175
      EndProperty
   End
   Begin HSES.b8Container b8Container1 
      Height          =   4335
      Left            =   60
      TabIndex        =   0
      Top             =   1710
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   7646
      BackColor       =   16185592
   End
   Begin HSES.b8Line b8Line1 
      Height          =   60
      Left            =   0
      TabIndex        =   9
      Top             =   960
      Width           =   8865
      _ExtentX        =   15637
      _ExtentY        =   106
   End
   Begin lvButton.lvButtons_H cmdCLose 
      Height          =   405
      Left            =   7470
      TabIndex        =   10
      Top             =   6120
      Width           =   1350
      _ExtentX        =   2381
      _ExtentY        =   714
      Caption         =   "&Close"
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
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H00F6F8F8&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   660
      Left            =   0
      ScaleHeight     =   660
      ScaleWidth      =   8910
      TabIndex        =   12
      Top             =   1050
      Width           =   8910
      Begin VB.TextBox txtGrade 
         Appearance      =   0  'Flat
         BackColor       =   &H00F6F8F8&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   285
         Index           =   3
         Left            =   8025
         TabIndex        =   17
         Text            =   "--"
         Top             =   270
         Width           =   585
      End
      Begin VB.TextBox txtGrade 
         Appearance      =   0  'Flat
         BackColor       =   &H00F6F8F8&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   285
         Index           =   2
         Left            =   7425
         TabIndex        =   16
         Text            =   "--"
         Top             =   270
         Width           =   585
      End
      Begin VB.TextBox txtGrade 
         Appearance      =   0  'Flat
         BackColor       =   &H00F6F8F8&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   285
         Index           =   1
         Left            =   6825
         TabIndex        =   15
         Text            =   "--"
         Top             =   270
         Width           =   585
      End
      Begin VB.TextBox txtGrade 
         Appearance      =   0  'Flat
         BackColor       =   &H00F6F8F8&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   285
         Index           =   0
         Left            =   6240
         TabIndex        =   14
         Text            =   "--"
         Top             =   270
         Width           =   585
      End
      Begin VB.TextBox txtStudentFullName 
         Appearance      =   0  'Flat
         BackColor       =   &H00F6F8F8&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   300
         Left            =   105
         MultiLine       =   -1  'True
         TabIndex        =   13
         Text            =   "frmStudentGrade.frx":06C4
         Top             =   255
         Width           =   4845
      End
      Begin HSES.b8Line b8Line4 
         Height          =   60
         Left            =   -105
         TabIndex        =   25
         Top             =   555
         Width           =   9030
         _ExtentX        =   15928
         _ExtentY        =   106
      End
      Begin VB.Image Image7 
         Height          =   30
         Left            =   0
         Picture         =   "frmStudentGrade.frx":06DB
         Stretch         =   -1  'True
         Top             =   0
         Width           =   8925
      End
      Begin VB.Image Image6 
         Height          =   105
         Left            =   -15
         Picture         =   "frmStudentGrade.frx":0778
         Stretch         =   -1  'True
         Top             =   570
         Width           =   8925
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "IV"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   8055
         TabIndex        =   23
         Top             =   60
         Width           =   150
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "III"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   7455
         TabIndex        =   22
         Top             =   60
         Width           =   180
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "II"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   6825
         TabIndex        =   21
         Top             =   60
         Width           =   120
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "I"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   6285
         TabIndex        =   20
         Top             =   60
         Width           =   60
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Grades:"
         ForeColor       =   &H00808080&
         Height          =   195
         Left            =   5325
         TabIndex        =   19
         Top             =   60
         Width           =   690
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Name:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   18
         Top             =   60
         Width           =   465
      End
   End
   Begin VB.Image Image5 
      Height          =   105
      Left            =   -90
      Picture         =   "frmStudentGrade.frx":0815
      Stretch         =   -1  'True
      Top             =   6480
      Width           =   9555
   End
   Begin VB.Image Image4 
      Height          =   3195
      Left            =   -30
      Picture         =   "frmStudentGrade.frx":08B2
      Stretch         =   -1  'True
      Top             =   2910
      Width           =   9315
   End
   Begin VB.Image Image3 
      Height          =   345
      Left            =   -60
      Picture         =   "frmStudentGrade.frx":094F
      Stretch         =   -1  'True
      Top             =   6030
      Width           =   9675
   End
End
Attribute VB_Name = "frmStudentGrade"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const AllYLTitle = "All Year Level"

Private Form_State As Integer

Dim iYLOldListIndex As Integer

Dim Form_Visible As Boolean

Dim vRS As New ADODB.Recordset



Public Function ShowForm(sStudentID As String, Optional sYearLevelTitle As String = AllYLTitle, Optional NewForm_State As Integer = 0)
    
    'set cur busy
    mdiMain.MousePointer = vbHourglass
    
    txtStudentID.Text = sStudentID
    
    Select Case sYearLevelTitle
        Case AllYLTitle
            cmbYearLevelTitle.ListIndex = 0
        Case "I"
            cmbYearLevelTitle.ListIndex = 1
        Case "II"
            cmbYearLevelTitle.ListIndex = 2
        Case "III"
            cmbYearLevelTitle.ListIndex = 3
        Case "IV"
            cmbYearLevelTitle.ListIndex = 4
    End Select
    
    Select Case NewForm_State
        Case 0 'normal view
        
        Case 1 'modify grade for single student only
        
            txtStudentID.Enabled = False
            cmbGetStudentID.Enabled = False
            
    End Select
    
    If GenerateList = False Then
        Unload Me
        mdiMain.MousePointer = vbDefault
        Exit Function
    End If
    
    'restore mousepointer
    mdiMain.MousePointer = vbDefault
    
    
    'show form
    '-----------------------------------------------
    Me.Show vbModal
    '-----------------------------------------------
End Function

Private Function Fill_List()
    vRS.Requery
    FillRecordToList vRS, listRecord, KeyGrade
End Function
Private Function GenerateList() As Boolean
    Dim sSQL As String
    Dim vEnrolment As tEnrolment
    Dim i As Integer
    Dim vYearlevel As tYearLevel
    Dim dAveGrade As Double
    

    If cmbYearLevelTitle.Text = AllYLTitle Then
        'show all year level
        sSQL = "SELECT tblGrade.GradeID, tblGrade.GradeValue, tblSubject.SubjectTitle, tblEnrolment.SchoolYear, tblDepartment.DepartmentTitle, tblYearLevel.YearLevelTitle, tblSection.SectionTitle" & _
                " FROM (tblDepartment INNER JOIN (tblYearLevel INNER JOIN (tblSubject INNER JOIN (tblStudent INNER JOIN (tblEnrolment INNER JOIN tblGrade ON tblEnrolment.EnrolmentID = tblGrade.EnrolmentID) ON tblStudent.StudentID = tblEnrolment.StudentID) ON tblSubject.SubjectID = tblGrade.SubjectOfferingID) ON tblYearLevel.YearLevelID = tblSubject.YearLevelID) ON tblDepartment.DepartmentID = tblSubject.DepartmentID) INNER JOIN tblSection ON (tblYearLevel.YearLevelID = tblSection.YearLevelID) AND (tblSection.SectionID = tblEnrolment.SectionID) AND (tblDepartment.DepartmentID = tblSection.DepartmentID)" & _
                " Where (((tblStudent.StudentID) = '" & txtStudentID.Text & "'))" & _
                " GROUP BY tblGrade.GradeID, tblGrade.GradeValue, tblSubject.SubjectTitle, tblEnrolment.SchoolYear, tblDepartment.DepartmentTitle, tblYearLevel.YearLevelTitle, tblSection.SectionTitle" & _
                " ORDER BY tblYearLevel.YearLevelTitle;"
        
        'update grade in all year level
        'i starts from 1 to skip "all category"
      
        For i = 1 To cmbYearLevelTitle.ListCount - 1
            If GetEnrolmentByStudentIDByYearLevelTitle(txtStudentID.Text, cmbYearLevelTitle.List(i), vEnrolment) = Success Then
                                
                'get average grade
                dAveGrade = 0
                If GetAveGradeByStudentIDByYLTitle(vEnrolment.StudentID, cmbYearLevelTitle.List(i), dAveGrade) = Success Then
                    
                    txtGrade(i - 1).Text = FormatNumber(dAveGrade, 2)
                    
                Else
                    MsgBox "Fatal: not enroled", vbCritical
                    GenerateList = False
                    Exit Function
                End If
                
            Else
            
                txtGrade(i - 1).Text = "--"
            End If
        Next
        
    Else
            
    
        'show by yearlevel title
        sSQL = "SELECT tblGrade.GradeID, tblGrade.GradeValue, tblSubject.SubjectTitle, tblEnrolment.SchoolYear, tblDepartment.DepartmentTitle, tblYearLevel.YearLevelTitle, tblSection.SectionTitle" & _
                " FROM tblYearLevel INNER JOIN (tblStudent INNER JOIN (((tblDepartment INNER JOIN tblSection ON tblDepartment.DepartmentID = tblSection.DepartmentID) INNER JOIN tblSubject ON tblDepartment.DepartmentID = tblSubject.DepartmentID) INNER JOIN (tblEnrolment INNER JOIN tblGrade ON tblEnrolment.EnrolmentID = tblGrade.EnrolmentID) ON (tblSubject.SubjectID = tblGrade.SubjectOfferingID) AND (tblSection.SectionID = tblEnrolment.SectionID)) ON tblStudent.StudentID = tblEnrolment.StudentID) ON (tblYearLevel.YearLevelID = tblSubject.YearLevelID) AND (tblYearLevel.YearLevelID = tblSection.YearLevelID)" & _
                " WHERE (((tblStudent.StudentID)='" & txtStudentID.Text & "'))" & _
                " GROUP BY tblGrade.GradeID, tblGrade.GradeValue, tblSubject.SubjectTitle, tblEnrolment.SchoolYear, tblDepartment.DepartmentTitle, tblYearLevel.YearLevelTitle, tblSection.SectionTitle" & _
                " Having (((tblYearLevel.YearLevelTitle) = '" & cmbYearLevelTitle & "'))" & _
                " ORDER BY tblYearLevel.YearLevelTitle;"
        
                 'get average grade
                dAveGrade = 0
                If GetAveGradeByStudentIDByYLTitle(txtStudentID.Text, cmbYearLevelTitle.Text, dAveGrade) = Success Then
                    If dAveGrade >= 60 Then
                        txtGrade(cmbYearLevelTitle.ListIndex - 1).Text = FormatNumber(dAveGrade, 2)
                    Else
                        txtGrade(cmbYearLevelTitle.ListIndex - 1).Text = "--"
                    End If
                    
                Else
                    MsgBox "Fatal: not enroled", vbCritical
                    txtGrade(cmbYearLevelTitle.ListIndex - 1).Text = "--"
                    GenerateList = False
                    Exit Function
                End If
    End If
    
    
    


    If ConnectRS(DB, vRS, sSQL) = True Then
        If AnyRecordExisted(vRS) = True Then
            Call Fill_List
        Else
            
            MsgBox "Student is not enroled in selected Year Level.", vbExclamation
            listRecord.ListItems.Clear
            
            GenerateList = False
            Exit Function
            
        End If
    Else
        'fatal
        'temp
        MsgBox "RS Connection Error!", vbExclamation
        GenerateList = False
        Exit Function
    End If
    
    GenerateList = True
End Function

Private Sub cmbGetStudentID_Click()
    Dim sStudentID As String

    sStudentID = PickStudent.GetStudentID
    
    If sStudentID <> "" Then
        txtStudentID.Text = sStudentID
    End If
End Sub


Private Sub cmbYearLevelTitle_Click()
    If txtStudentID = "Enter Student ID" Then
        listRecord.ListItems.Clear: txtStudentFullName.Text = "No Student Selected..."
    Else
        GenerateList
    End If
End Sub

Private Sub cmbYearLevelTitle_GotFocus()
    iYLOldListIndex = cmbYearLevelTitle.ListIndex
End Sub

Private Sub cmbYearLevelTitle_LostFocus()
    If cmbYearLevelTitle.ListIndex < 0 Then
        cmbYearLevelTitle.ListIndex = iYLOldListIndex
    End If
End Sub


Private Sub cmdCLose_Click()
    Unload Me
End Sub

Private Sub cmdEditGrade_Click()
    EditSelectedGrade
End Sub

Private Sub Form_Activate()
    Form_Visible = True
End Sub

Private Sub Form_Load()

    Dim vRS As New ADODB.Recordset
    Dim vYearlevel As tYearLevel
    
    
    cmbYearLevelTitle.Clear
    cmbYearLevelTitle.AddItem AllYLTitle
    
    If CreateDefaultRSYearLevel(vRS) <> Success Then
        'fatal
        MsgBox "Fatal:frmstudentgrade.Form_load - CreateDefaultRSYearLevel", vbExclamation
        Exit Sub
    End If
    
    If RSMoveFirst(vRS) <> True Then
        'fatal
        MsgBox "Fatal:frmstudentgrade.Form_load - YearLevelMoveFirst", vbExclamation
        Exit Sub
    End If
    
    
        While GetYearLevelMoveNext(vRS, vYearlevel) = True
            cmbYearLevelTitle.AddItem vYearlevel.YearLevelTitle
        Wend
        
        cmbYearLevelTitle.ListIndex = 0
    
    Set vRS = Nothing
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Set vRS = Nothing
    Form_Visible = False
End Sub


Private Sub listRecord_DblClick()
    EditSelectedGrade
End Sub

Private Sub txtStudentID_Change()
    Dim vStudent As tStudent
    
    If GetStudentByID(txtStudentID, vStudent) = Success Then
        txtStudentFullName.Text = vStudent.LastName & ", " & vStudent.FirstName & " " & vStudent.MiddleName
        cmbYearLevelTitle.ListIndex = 0
        If Form_Visible = True Then
            GenerateList
        End If
    Else
        listRecord.ListItems.Clear: txtStudentFullName.Text = "No Student Selected..."
        txtGrade(0) = "--"
        txtGrade(1) = "--"
        txtGrade(2) = "--"
        txtGrade(3) = "--"
    End If
    
End Sub

Private Sub txtStudentID_LostFocus()
    If StudentExistByID(txtStudentID.Text) <> Success Then
        
        txtStudentID = "Enter Student ID"
        listRecord.ListItems.Clear: txtStudentFullName.Text = "No Student Selected..."
        
    End If
End Sub















Private Sub EditSelectedGrade()
    Dim lvKey As String
    Dim isEditable As Boolean
    
    lvKey = GetLVKey(listRecord.SelectedItem)
    
    If Len(lvKey) > 0 Then
        If IsGradeEditable(lvKey, isEditable) = Success Then
        
            If isEditable = True Then
                'edit student grade
                If frmNewGrade.ShowForm(lvKey) = True Then
                    GenerateList
                End If
            Else
            
                MsgBox "This Grade cannot be edited." & vbNewLine & "Selected Student is already enroled at the next Year Level.", vbExclamation
                
            End If
        Else
            MsgBox "FATAL ERROR: frmStudentGrade.ListRecord_DblClick", vbCritical
        End If
    Else
        MsgBox "Please select Subject", vbExclamation
    End If
End Sub
