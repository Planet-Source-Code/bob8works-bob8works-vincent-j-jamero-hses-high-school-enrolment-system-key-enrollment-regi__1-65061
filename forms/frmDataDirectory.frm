VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmRecordExplorer 
   BackColor       =   &H00F8D0B7&
   Caption         =   "Record Explorer"
   ClientHeight    =   6540
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8055
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmDataDirectory.frx":0000
   LinkTopic       =   "frmRecordExplorer"
   MDIChild        =   -1  'True
   ScaleHeight     =   436
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   537
   Begin MSComctlLib.ImageList imgListEnrolment 
      Left            =   3555
      Top             =   2835
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDataDirectory.frx":08CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDataDirectory.frx":0E64
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDataDirectory.frx":13FE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDataDirectory.frx":1998
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDataDirectory.frx":1F32
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDataDirectory.frx":24CC
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvFiles 
      Height          =   5550
      Left            =   3090
      TabIndex        =   1
      Top             =   570
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   9790
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      _Version        =   393217
      Icons           =   "imgListEnrolment"
      SmallIcons      =   "imgListEnrolment"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin MSComctlLib.TreeView tvFolder 
      Height          =   5550
      Left            =   0
      TabIndex        =   2
      Top             =   570
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   9790
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   423
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      ImageList       =   "imgListEnrolment"
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSForms.ComboBox cmbFolderAddress 
      Height          =   330
      Left            =   540
      TabIndex        =   0
      Top             =   120
      Width           =   4380
      VariousPropertyBits=   748701723
      BorderStyle     =   1
      DisplayStyle    =   3
      Size            =   "7726;582"
      MatchEntry      =   1
      ListStyle       =   1
      ShowDropButtonWhen=   1
      BorderColor     =   12164479
      SpecialEffect   =   0
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Image Image2 
      Height          =   585
      Left            =   0
      Top             =   0
      Width           =   15360
      BorderColor     =   13874590
      BackColor       =   16306359
      BorderStyle     =   0
      SizeMode        =   1
      Size            =   "27093;1032"
      Picture         =   "frmDataDirectory.frx":2626
   End
End
Attribute VB_Name = "frmRecordExplorer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim IsSideBarVisible As Boolean

Const sRootEnrolment = "Enrolment"
Const sRootStudentAccount = "StudentAccount"



Const sKeyAddSchoolYear = "addschoolyear"
Const sKeyviewSchoolYear = "viewschoolyear"

Const sKeyAddDepartment = "adddepartment"
Const sKeyViewDepartment = "viewdepartment"

Const sKeyAddEnrolment = "addenrolment"

Const sKeyAddSection = "addsection"
Const sKeyAddSectionEnrolment = "addSectionenrollment"
Private Started As Boolean

'starts here
Public Sub ShowMe(Optional sRecordPath As String)
    
    'initialize
    If Started And sRecordPath = "" Then
        Me.Show
        Exit Sub
    End If
    
    Started = True
    
    If sRecordPath = "" Then
        'normal call
        FillRoots
    Else
        'call with record address
    
    End If
    'show form
    Me.Show
End Sub



Private Sub ArrangeControls()
    
    'arrange vertical
    If Me.ScaleWidth > 204 Then
        
        tvFolder.Left = 0
        tvFolder.Width = 200
        
        lvFiles.Left = 204
        lvFiles.Width = Me.ScaleWidth - lvFiles.Left
    
    End If
    
    'Arrange horizontal
    If Me.ScaleHeight > tvFolder.Top Then
    
        tvFolder.Height = Me.ScaleHeight - tvFolder.Top
        lvFiles.Height = Me.ScaleHeight - lvFiles.Top
    
    End If
End Sub

Private Sub FillRoots()

    
    tvFolder.Nodes.Clear
    lvFiles.ListItems.Clear
    

    tvFolder.Nodes.Add , , "HSES", "HSES", 1
    
    tvFolder.Nodes.Add "HSES", tvwChild, sRootEnrolment, "Enrolment Entries", 2
    tvFolder.Nodes.Add "HSES", tvwChild, sRootStudentAccount, "Student Accounts", 2
    
    FillRootFolders
    
End Sub

Private Sub FillRootFolders()
    lvFiles.ListItems.Clear
    
    
    
    lvFiles.ListItems.Add , sRootEnrolment, "Enrolment Entries", 2
    lvFiles.ListItems.Add , sRootStudentAccount, "Student Accounts", 2

End Sub


Private Sub FillSchoolYears()
    
    Dim vRS As New ADODB.Recordset
    Dim vSchoolYear As tSchoolYear
    

    lvFiles.ListItems.Clear
    
    lvFiles.ListItems.Add , sKeyAddSchoolYear, "Add New School Year", 6
    lvFiles.ListItems.Add , sKeyviewSchoolYear, "View School Year List"
    
    
    If CreateDefaultRSSchoolYear(vRS) = 1 Then
        If RSMoveFirst(vRS) Then
            While GetSchoolYearMoveNext(vRS, vSchoolYear) = success
                If NotInList(sRootEnrolment & "/" & vSchoolYear.SchoolYearTitle) Then
                    
                    tvFolder.Nodes.Add sRootEnrolment, tvwChild, sRootEnrolment & "/" & vSchoolYear.SchoolYearTitle, vSchoolYear.SchoolYearTitle, 3
                    
                End If
                
                lvFiles.ListItems.Add , sRootEnrolment & "/" & vSchoolYear.SchoolYearTitle, vSchoolYear.SchoolYearTitle, 3
            Wend
        End If
    End If
    
    Set vRS = Nothing
End Sub

Private Sub FillDepartments(sParentKey As String)
    Dim vDepartment As tDepartment
    Dim vRS As New ADODB.Recordset
    
    lvFiles.ListItems.Clear
    
    lvFiles.ListItems.Add , sKeyAddDepartment, "Add New Department", 6
    lvFiles.ListItems.Add , sKeyViewDepartment, "View Department List"
    
    
    If CreateDefaultRSDepartment(vRS) = success Then
        If RSMoveFirst(vRS) = True Then
        
            While GetDepartmentMoveNext(vRS, vDepartment) = success
                
                If NotInList(sParentKey & "/" & vDepartment.DepartmentTitle) Then
                
                    tvFolder.Nodes.Add sParentKey, tvwChild, sParentKey & "/" & vDepartment.DepartmentTitle, vDepartment.DepartmentTitle, 3
                    
                End If
                
                
                lvFiles.ListItems.Add , sParentKey & "/" & vDepartment.DepartmentTitle, vDepartment.DepartmentTitle, 3
                
            Wend
            
        End If
    End If
End Sub


Private Sub FillYearLevels(sParentKey As String)
    Dim vYearLevel As tYearLevel
    Dim vRS As New ADODB.Recordset
    lvFiles.ListItems.Clear
    
    If CreateDefaultRSYearLevel(vRS) = success Then
        If RSMoveFirst(vRS) = True Then
            While GetYearLevelMoveNext(vRS, vYearLevel) = True
                If NotInList(sParentKey & "/" & vYearLevel.YearLevelTitle) Then
                    
                    tvFolder.Nodes.Add sParentKey, tvwChild, sParentKey & "/" & vYearLevel.YearLevelTitle, vYearLevel.YearLevelTitle, 3
                    
                End If
                
                lvFiles.ListItems.Add , sParentKey & "/" & vYearLevel.YearLevelTitle, vYearLevel.YearLevelTitle, 3
            Wend
        End If
    End If
    Set vRS = Nothing
End Sub

Private Sub FillSections(sParentKey As String)
    Dim QRYSection As New ADODB.Recordset
    
    Dim vSection As tSection
    
    Dim vDepartment As tDepartment
    Dim vYearLevel As tYearLevel
    
    Dim splitPath() As String
    
    splitPath = Split(sParentKey, "/")
    
    GetDepartmentByTitle splitPath(2), vDepartment
    GetYearLevelbyTitle splitPath(3), vYearLevel
    
    
    'clear files
    lvFiles.ListItems.Clear
    
    If ConnectRS(DB, QRYSection, "select * from tblsection where departmentid = '" & vDepartment.DepartmentID & "' and yearlevelid = " & vYearLevel.YearLevelID) Then
        
        
        If AnyRecordExisted(QRYSection) Then
            
            QRYSection.MoveFirst
            
            
            
            While Not QRYSection.EOF
                
                If NotInList(sParentKey & "/" & QRYSection.Fields("sectiontitle").Value) Then
                    
                    tvFolder.Nodes.Add sParentKey, tvwChild, sParentKey & "/" & QRYSection.Fields("sectiontitle").Value, QRYSection.Fields("sectiontitle").Value, 3

                End If
                
                lvFiles.ListItems.Add , sParentKey & "/" & QRYSection.Fields("sectiontitle").Value, QRYSection.Fields("sectiontitle").Value, 3
                
                QRYSection.MoveNext
                
            Wend
        End If
    
    End If
    
    
    
    Set QRYSection = Nothing
End Sub


Private Sub FillEnromentEntries(sParentKey As String)

    Dim splitPath() As String
    
    Dim vSchooYear As tSchoolYear
    Dim vSection As tSection
    
    Dim QRYEnrolment As New ADODB.Recordset
    Dim sSQL As String
    
    'clear file list
    lvFiles.ListItems.Clear
    
    'split key
    splitPath = Split(sParentKey, "/")

      sSQL = "SELECT tblEnrolment.EnrolmentID, tblStudent.LastName, tblStudent.FirstName, tblStudent.MiddleName, tblEnrolment.SchoolYear, tblSection.SectionTitle, tblDepartment.DepartmentTitle, tblTeacher.TeacherTitle FROM tblTeacher INNER JOIN (tblStudent INNER JOIN ((tblDepartment INNER JOIN tblSection ON tblDepartment.DepartmentID = tblSection.DepartmentID) INNER JOIN tblEnrolment ON tblSection.SectionID = tblEnrolment.SectionID) ON tblStudent.StudentID = tblEnrolment.StudentID) ON tblTeacher.TeacherID = tblSection.TeacherID WHERE (((tblEnrolment.SchoolYear)='" & splitPath(1) & "') AND ((tblSection.SectionTitle)='" & splitPath(4) & "'));"

    If ConnectRS(DB, QRYEnrolment, sSQL) Then
        FillRecordToList QRYEnrolment, lvFiles, KeyEnrolment
    End If
        
        
    tvFolder.SelectedItem.Text = splitPath(4) & " (" & getRecordCount(QRYEnrolment) & ")"
        
    Set QRYEnrolment = Nothing
End Sub




Private Function NotInList(sKey As String) As Boolean
    Dim Node As MSComctlLib.Node
    
    For Each Node In tvFolder.Nodes
        If sKey = Node.Key Then
            NotInList = False
            
            Exit Function
        End If
    Next

    NotInList = True
End Function




Private Sub Form_Activate()
    'get sidebar visibility
    IsSideBarVisible = mdiMain.SideBar.Visible
    
    'hide sidebar
    mdiMain.SideBar.Visible = False
End Sub

Private Sub Form_Deactivate()
    'restore sidebar visibility
    mdiMain.SideBar.Visible = IsSideBarVisible
End Sub

Private Sub Form_Resize()
    ArrangeControls
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'restore sidebar visibility
    mdiMain.SideBar.Visible = IsSideBarVisible
    
    Started = False
End Sub







Private Sub lvFiles_DblClick()
    If lvFiles.ListItems.Count > 0 Then
        tvFolderSelectItem lvFiles.SelectedItem.Key
    End If
End Sub





Private Sub tvFolder_Expand(ByVal Node As MSComctlLib.Node)
    Select Case sGetFirstKey(Node.Key)
        Case "HSES"
            FillRootFolders
        Case sRootEnrolment
            EnrolmentFolderExpand Node
    End Select
    
End Sub
Private Sub tvFolder_NodeClick(ByVal Node As MSComctlLib.Node)


    Select Case sGetFirstKey(Node.Key)
        Case "HSES"
            FillRootFolders
        Case sRootEnrolment
            EnrolmentFolderExpand tvFolder.SelectedItem
    End Select
    
    Node.Expanded = True

End Sub

Private Function tvFolderSelectItem(sKey As String)
    Dim Node As MSComctlLib.Node
    
    For Each Node In tvFolder.Nodes
        If Node.Key = sKey Then
            Node.Selected = True
            EnrolmentFolderExpand Node
            Exit For
        End If
    Next
End Function


Private Sub EnrolmentFolderExpand(Node As MSComctlLib.Node)

    Dim splitPath() As String
    splitPath = Split(Node.Key, "/")

If sGetFirstKey(tvFolder.SelectedItem.Key) <> sRootEnrolment Then Exit Sub



    Select Case UBound(splitPath)
        Case 0 'enrolment enries folder expanded
            FillSchoolYears
            lvFiles.View = lvwIcon
            
        Case 1 'school year entries folder expanded
            FillDepartments Node.Key
            lvFiles.View = lvwIcon
            
        Case 2 'department entries folder expanded
            FillYearLevels Node.Key
            lvFiles.View = lvwIcon
            
        Case 3 'YearLevels entries folder expanded
            FillSections Node.Key
            lvFiles.View = lvwIcon
            
        Case 4 'section selected
            FillEnromentEntries Node.Key
            lvFiles.View = lvwReport
                        
    End Select
End Sub



Private Function sGetFirstKey(sKey As String) As String
    Dim splitPath() As String
    splitPath = Split(sKey, "/")
    
    sGetFirstKey = splitPath(0)

End Function
