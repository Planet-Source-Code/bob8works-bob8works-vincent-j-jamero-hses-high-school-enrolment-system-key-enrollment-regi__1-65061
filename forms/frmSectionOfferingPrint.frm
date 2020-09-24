VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPrintSectionOffering 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Print Sec. Offr."
   ClientHeight    =   6300
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   6555
   Icon            =   "frmSectionOfferingPrint.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   420
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   437
   WindowState     =   2  'Maximized
   Begin VB.PictureBox bgMain 
      Appearance      =   0  'Flat
      BackColor       =   &H00D8E9EC&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4995
      Left            =   0
      ScaleHeight     =   333
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   471
      TabIndex        =   0
      Top             =   0
      Width           =   7065
      Begin VB.PictureBox pbBGButton 
         Appearance      =   0  'Flat
         BackColor       =   &H00D8E9EC&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   525
         Left            =   90
         ScaleHeight     =   525
         ScaleWidth      =   6675
         TabIndex        =   1
         Top             =   360
         Width           =   6675
      End
      Begin MSComctlLib.ImageList icoHeader 
         Left            =   6240
         Top             =   3270
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   2
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmSectionOfferingPrint.frx":08CA
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmSectionOfferingPrint.frx":0E64
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin HSES.b8ChildTitleBar b8Title 
         Height          =   345
         Left            =   60
         TabIndex        =   2
         Top             =   0
         Width           =   6615
         _extentx        =   11668
         _extenty        =   609
         backcolor       =   12307149
         caption         =   "PrintSectionOffering Information"
         font            =   "frmSectionOfferingPrint.frx":13FE
         fontbold        =   -1  'True
         fontname        =   "Tahoma"
         fontsize        =   9.75
         forecolor       =   16512
         gradtheme       =   2
      End
      Begin MSComctlLib.ImageList ilRecordIco 
         Left            =   6360
         Top             =   1830
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   1
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmSectionOfferingPrint.frx":1426
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ListView listRecord 
         Height          =   4275
         Left            =   0
         TabIndex        =   3
         Top             =   930
         Width           =   6105
         _ExtentX        =   10769
         _ExtentY        =   7541
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         HotTracking     =   -1  'True
         _Version        =   393217
         Icons           =   "imgListIco"
         SmallIcons      =   "imgListIco"
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
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Select.."
            Object.Width           =   7911
         EndProperty
      End
   End
   Begin MSComctlLib.ImageList imgListIco 
      Left            =   6270
      Top             =   2520
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSectionOfferingPrint.frx":19C0
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Shape b8cMain 
      BorderColor     =   &H00C0C0C0&
      Height          =   2055
      Left            =   2130
      Top             =   3510
      Width           =   3315
   End
End
Attribute VB_Name = "frmPrintSectionOffering"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim curSectionOfferingID As String
Dim curSchoolYear As String


Public Function ShowForm(Optional sSectionOfferingID As String = "", Optional sSchoolYear As String = "")
    'temp
    'On Error Resume Next
    curSectionOfferingID = sSectionOfferingID
    
    curSchoolYear = sSchoolYear
    
    b8Title.Caption = "Print Department Entries - " & _
    "SY: " & curSchoolYear & " / Sect. Offr. ID: " & curSectionOfferingID
    
    
    Me.Show
    Me.SetFocus
    
End Function


Public Sub Form_Activate()
    RefreshReportList
    
    mdiMain.RegMDIChild Me
End Sub


Private Sub Form_Resize()
    On Error Resume Next
    Me.ScaleMode = vbPixels
    
    
    b8cMain.Move Form_LeftMargin - 1, Form_TopMargin - 1, Me.ScaleWidth - (Form_LeftMargin - 1) * 2, Me.ScaleHeight - (Form_TopMargin - 1) * 2
    bgMain.Move Form_LeftMargin, Form_TopMargin, Me.ScaleWidth - Form_LeftMargin * 2, Me.ScaleHeight - Form_TopMargin * 2
    b8Title.Move 0, 0, bgMain.Width
    listRecord.Move listRecord.Left, pbBGButton.Top + pbBGButton.Height, bgMain.Width - (listRecord.Left * 2), Me.ScaleHeight - (pbBGButton.Top + pbBGButton.Height)
    listRecord.ColumnHeaders(1).Width = listRecord.Width - 6

End Sub
Private Sub listRecord_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call listRecord_DblClick
    End If
End Sub


Public Sub RefreshReportList()
    
    listRecord.ListItems.Clear
    
    listRecord.ListItems.Add , "SectionList", "Section List", 1, 1
    listRecord.ListItems.Add , "SectionWithStudentList", "Section With Student List", 1, 1
    listRecord.ListItems.Add , "SectionWithSubjectList", "Section With Subject List", 1, 1
    listRecord.ListItems.Add , "SectionWithStudentListWithGrade", "Section With Student List And Grade", 1, 1

       
End Sub

Public Sub listRecord_DblClick()
    
    Select Case listRecord.SelectedItem.Key
            
        Case "SectionWithStudentList"
            Call ShowSectionWithStudentList
            
        Case "SectionWithSubjectList"
            Call ShowSectionWithSubjectList
            
        Case "SectionList"
            Call ShowSectionList
        
        Case "SectionWithStudentListWithGrade"
            Call ShowSectionWithStudentListWithGrade
            
    End Select
        
    
End Sub



Private Sub ShowSectionList()
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    Dim sSchoolYear As String
    
    sSchoolYear = curSchoolYear
    
    If sSchoolYear = "" Then
        sSchoolYear = PickSchoolYear.GetItem()
    End If
    
    If sSchoolYear = "" Then
        GoTo ReleaseAndExit
    End If
    
    sSQL = "SELECT tblSectionOffering.SchoolYear, tblDepartment.DepartmentTitle, [tblYearLevel]![YearLevelTitle] & ' - ' & [tblSection]![SectionTitle] AS SectionFullTitle, tblRoom.Room, Count(tblEnrolment.EnrolmentID) AS CountOfEnrolmentID, tblSectionOffering.MaxStudentCount, [MinGrade] & ' - ' & [MaxGrade] AS GradeAllowed, [tblTeacher_1]![LastName] & ', ' & [tblTeacher_1]![FirstName] & ' ' & [tblTeacher_1]![MiddleName] AS SectionTeacherFullName" & _
            " FROM (tblRoom INNER JOIN (tblDepartment INNER JOIN (tblTeacher AS tblTeacher_1 INNER JOIN (tblYearLevel INNER JOIN (tblSection INNER JOIN tblSectionOffering ON tblSection.SectionID = tblSectionOffering.SectionID) ON tblYearLevel.YearLevelID = tblSection.YearLevelID) ON tblTeacher_1.TeacherID = tblSectionOffering.TeacherID) ON tblDepartment.DepartmentID = tblSection.DepartmentID) ON tblRoom.RoomID = tblSectionOffering.RoomID) LEFT JOIN tblEnrolment ON tblSectionOffering.SectionOfferingID = tblEnrolment.SectionOfferingID" & _
            " WHERE (((tblSectionOffering.SchoolYear)='" & sSchoolYear & "'))" & _
            " GROUP BY tblSectionOffering.SchoolYear, tblDepartment.DepartmentTitle, [tblYearLevel]![YearLevelTitle] & ' - ' & [tblSection]![SectionTitle], tblRoom.Room, tblSectionOffering.MaxStudentCount, [MinGrade] & ' - ' & [MaxGrade], [tblTeacher_1]![LastName] & ', ' & [tblTeacher_1]![FirstName] & ' ' & [tblTeacher_1]![MiddleName]" & _
            " ORDER BY tblSectionOffering.SchoolYear;"

    If ConnectRS(HSESDB, vRS, sSQL) = False Then
        MsgBox "error"
        GoTo ReleaseAndExit
    End If
    
    Set drSectionOfferingList.DataSource = vRS
    drSectionOfferingList.Show vbModal

ReleaseAndExit:

    sSchoolYear = ""

    Set vRS = Nothing
    
End Sub


Private Sub ShowSectionWithSubjectList()
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    Dim bNotEmp As Boolean
    
    If curSectionOfferingID = "" Then
        curSectionOfferingID = PickSectionOffering.GetSectionOfferingID()
        bNotEmp = True
    End If
    
    If curSectionOfferingID = "" Then
        GoTo ReleaseAndExit
    End If
    
        sSQL = "SELECT tblSectionOffering.SectionOfferingID, [tblYearLevel]![YearLevelTitle] & ' - ' & [tblSection]![SectionTitle] AS SectionFullTitle, tblSectionOffering.SchoolYear, [tblTeacher_1]![LastName] & ', ' & [tblTeacher_1]![FirstName] & ' ' & [tblTeacher_1]![MiddleName] AS SectionTeacherFullName, tblRoom.Room, tblSectionOffering.MaxStudentCount, tblSectionOffering.MinGrade, tblSectionOffering.MaxGrade, tblSubject.SubjectTitle, tblSubjectOffering.Days, tblSubjectOffering.SchedTimeStart, tblSubjectOffering.SchedTimeEnd, [tblTeacher]![LastName] & ', ' & [tblTeacher]![FirstName] & ' ' & [tblTeacher]![MiddleName] AS TeacherFullName" & _
            " FROM tblRoom INNER JOIN (tblTeacher AS tblTeacher_1 INNER JOIN ((tblYearLevel INNER JOIN (tblSection INNER JOIN tblSectionOffering ON tblSection.SectionID = tblSectionOffering.SectionID) ON tblYearLevel.YearLevelID = tblSection.YearLevelID) INNER JOIN (tblSubject INNER JOIN (tblTeacher INNER JOIN tblSubjectOffering ON tblTeacher.TeacherID = tblSubjectOffering.TeacherID) ON tblSubject.SubjectID = tblSubjectOffering.SubjectID) ON tblSectionOffering.SectionOfferingID = tblSubjectOffering.SectionOfferingID) ON tblTeacher_1.TeacherID = tblSectionOffering.TeacherID) ON tblRoom.RoomID = tblSectionOffering.RoomID" & _
            " WHERE (((tblSectionOffering.SectionOfferingID)='" & curSectionOfferingID & "'));"

    
    If ConnectRS(HSESDB, vRS, sSQL) = False Then
        'fatal error
        MsgBox "error"
        GoTo ReleaseAndExit
    End If
    
    drSectionOfferingDetail.Sections("secDetail").Controls("lbl1").Caption = ReadField(vRS.Fields("SectionFullTitle"))
    drSectionOfferingDetail.Sections("secDetail").Controls("lbl2").Caption = ReadField(vRS.Fields("SchoolYear"))
    drSectionOfferingDetail.Sections("secDetail").Controls("lbl3").Caption = ReadField(vRS.Fields("SectionTeacherFullName"))
    drSectionOfferingDetail.Sections("secDetail").Controls("lbl4").Caption = ReadField(vRS.Fields("MaxStudentCount"))
    drSectionOfferingDetail.Sections("secDetail").Controls("lbl5").Caption = ReadField(vRS.Fields("MinGrade")) & " - " & ReadField(vRS.Fields("maxGrade"))
    drSectionOfferingDetail.Sections("secDetail").Controls("lblRoom").Caption = ReadField(vRS.Fields("Room"))

    Set drSectionOfferingDetail.DataSource = vRS
    
    drSectionOfferingDetail.Show vbModal
    
    
ReleaseAndExit:
    If bNotEmp = True Then
        curSectionOfferingID = ""
    End If
    Set vRS = Nothing

End Sub


Private Sub ShowSectionWithStudentList()
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    Dim bNotEmp As Boolean
    
    Dim lEnrolessCount As Long
    
    
    If curSectionOfferingID = "" Then
        curSectionOfferingID = PickSectionOffering.GetSectionOfferingID
        bNotEmp = True
    End If
    
    If curSectionOfferingID = "" Then
        GoTo ReleaseAndExit
    End If
    
    sSQL = "SELECT tblSectionOffering.SectionOfferingID, [tblYearLevel]![YearLevelTitle] & ' - ' & [tblSection]![SectionTitle] AS SectionFullTitle, tblSectionOffering.SchoolYear, [tblTeacher_1]![LastName] & ', ' & [tblTeacher_1]![FirstName] & ' ' & [tblTeacher_1]![MiddleName] AS SectionTeacherFullName, tblSectionOffering.MaxStudentCount, tblSectionOffering.MinGrade, tblSectionOffering.MaxGrade, [tblStudent]![LastName] & ', ' & [tblStudent]![FirstName] & ' ' & [tblStudent]![MiddleName] AS StudentFullName, tblEnrolment.DateEnroled, tblStudent.Gender, tblRoom.Room" & _
            " FROM tblRoom INNER JOIN (tblStudent INNER JOIN ((tblTeacher AS tblTeacher_1 INNER JOIN (tblYearLevel INNER JOIN (tblSection INNER JOIN tblSectionOffering ON tblSection.SectionID = tblSectionOffering.SectionID) ON tblYearLevel.YearLevelID = tblSection.YearLevelID) ON tblTeacher_1.TeacherID = tblSectionOffering.TeacherID) INNER JOIN tblEnrolment ON tblSectionOffering.SectionOfferingID = tblEnrolment.SectionOfferingID) ON tblStudent.StudentID = tblEnrolment.StudentID) ON tblRoom.RoomID = tblSectionOffering.RoomID" & _
            " GROUP BY tblSectionOffering.SectionOfferingID, [tblYearLevel]![YearLevelTitle] & ' - ' & [tblSection]![SectionTitle], tblSectionOffering.SchoolYear, [tblTeacher_1]![LastName] & ', ' & [tblTeacher_1]![FirstName] & ' ' & [tblTeacher_1]![MiddleName], tblSectionOffering.MaxStudentCount, tblSectionOffering.MinGrade, tblSectionOffering.MaxGrade, [tblStudent]![LastName] & ', ' & [tblStudent]![FirstName] & ' ' & [tblStudent]![MiddleName], tblEnrolment.DateEnroled, tblStudent.Gender, tblRoom.Room" & _
            " HAVING (((tblSectionOffering.SectionOfferingID)='" & curSectionOfferingID & "'))" & _
            " ORDER BY [tblStudent]![LastName] & ', ' & [tblStudent]![FirstName] & ' ' & [tblStudent]![MiddleName]"

    
    If ConnectRS(HSESDB, vRS, sSQL) = False Then
        'fatal error
        MsgBox "error"
        GoTo ReleaseAndExit
    End If
    
    'get number of enrolles
    If GetEnrolmentCountBySectionOfferingID(curSectionOfferingID, lEnrolessCount) <> Success Then
        'fatal error
        CatchError "frmPrintSectionOffering", "ShowSectionWithStudentList", "GetEnrolmentCountBySectionOfferingID went failed"
        GoTo ReleaseAndExit
    End If
    drSectionOfferingWithEnrolment.Sections("secDetail").Controls("lbl1").Caption = ReadField(vRS.Fields("SectionFullTitle"))
    drSectionOfferingWithEnrolment.Sections("secDetail").Controls("lbl2").Caption = ReadField(vRS.Fields("SchoolYear"))
    drSectionOfferingWithEnrolment.Sections("secDetail").Controls("lbl3").Caption = ReadField(vRS.Fields("SectionTeacherFullName"))
    drSectionOfferingWithEnrolment.Sections("secDetail").Controls("lbl4").Caption = ReadField(vRS.Fields("MaxStudentCount"))
    drSectionOfferingWithEnrolment.Sections("secDetail").Controls("lbl5").Caption = ReadField(vRS.Fields("MinGrade")) & " - " & ReadField(vRS.Fields("maxGrade"))
    drSectionOfferingWithEnrolment.Sections("secDetail").Controls("lblRoom").Caption = ReadField(vRS.Fields("Room"))
    drSectionOfferingWithEnrolment.Sections("secDetail").Controls("lblEnrollesCount").Caption = lEnrolessCount

    'get number of enrolles
    If GetEnrolmentCountBySectionOfferingID(curSectionOfferingID, lEnrolessCount) <> Success Then
        'fatal error
        CatchError "frmAddEnrolment", "ShowSectionDetail", "GetEnrolmentCountBySectionOfferingID(txtSectionOfferingID.Text, selSectionStudentCount) went failed"
    End If
    Set drSectionOfferingWithEnrolment.DataSource = vRS
    
    drSectionOfferingWithEnrolment.Show vbModal
    
    
ReleaseAndExit:
    If bNotEmp = True Then
        curSectionOfferingID = ""
    End If
    Set vRS = Nothing

End Sub


Private Sub ShowSectionWithStudentListWithGrade()

    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    Dim bNotEmp As Boolean
    
    Dim lEnrolessCount As Long
    
    
    If curSectionOfferingID = "" Then
        curSectionOfferingID = PickSectionOffering.GetSectionOfferingID
        bNotEmp = True
    End If
    
    If curSectionOfferingID = "" Then
        GoTo ReleaseAndExit
    End If
    
    sSQL = "SELECT tblSectionOffering.SectionOfferingID, [tblYearLevel]![YearLevelTitle] & ' - ' & [tblSection]![SectionTitle] AS SectionFullTitle, tblSectionOffering.SchoolYear, [tblTeacher_1]![LastName] & ', ' & [tblTeacher_1]![FirstName] & ' ' & [tblTeacher_1]![MiddleName] AS SectionTeacherFullName, tblSectionOffering.MaxStudentCount, tblSectionOffering.MinGrade, tblSectionOffering.MaxGrade, [tblStudent]![LastName] & ', ' & [tblStudent]![FirstName] & ' ' & [tblStudent]![MiddleName] AS StudentFullName, tblEnrolment.DateEnroled, tblStudent.Gender, Avg(tblGrade.GradeValue) AS AvgOfGradeValue, tblRoom.Room" & _
            " FROM (tblRoom INNER JOIN (tblStudent INNER JOIN ((tblTeacher AS tblTeacher_1 INNER JOIN (tblYearLevel INNER JOIN (tblSection INNER JOIN tblSectionOffering ON tblSection.SectionID = tblSectionOffering.SectionID) ON tblYearLevel.YearLevelID = tblSection.YearLevelID) ON tblTeacher_1.TeacherID = tblSectionOffering.TeacherID) INNER JOIN tblEnrolment ON tblSectionOffering.SectionOfferingID = tblEnrolment.SectionOfferingID) ON tblStudent.StudentID = tblEnrolment.StudentID) ON tblRoom.RoomID = tblSectionOffering.RoomID) INNER JOIN tblGrade ON tblEnrolment.EnrolmentID = tblGrade.EnrolmentID" & _
            " GROUP BY tblSectionOffering.SectionOfferingID, [tblYearLevel]![YearLevelTitle] & ' - ' & [tblSection]![SectionTitle], tblSectionOffering.SchoolYear, [tblTeacher_1]![LastName] & ', ' & [tblTeacher_1]![FirstName] & ' ' & [tblTeacher_1]![MiddleName], tblSectionOffering.MaxStudentCount, tblSectionOffering.MinGrade, tblSectionOffering.MaxGrade, [tblStudent]![LastName] & ', ' & [tblStudent]![FirstName] & ' ' & [tblStudent]![MiddleName], tblEnrolment.DateEnroled, tblStudent.Gender, tblRoom.Room" & _
            " HAVING (((tblSectionOffering.SectionOfferingID)='" & curSectionOfferingID & "'))" & _
            " ORDER BY [tblStudent]![LastName] & ', ' & [tblStudent]![FirstName] & ' ' & [tblStudent]![MiddleName]"


    
    If ConnectRS(HSESDB, vRS, sSQL) = False Then
        'fatal error
        MsgBox "error"
        GoTo ReleaseAndExit
    End If
    
    'get number of enrolles
    If GetEnrolmentCountBySectionOfferingID(curSectionOfferingID, lEnrolessCount) <> Success Then
        'fatal error
        CatchError "frmPrintSectionOffering", "ShowSectionWithStudentListWithGrade", "GetEnrolmentCountBySectionOfferingID went failed"
        GoTo ReleaseAndExit
    End If
    drSectionOfferingWithEnrolmentWithGrade.Sections("secDetail").Controls("lbl1").Caption = ReadField(vRS.Fields("SectionFullTitle"))
    drSectionOfferingWithEnrolmentWithGrade.Sections("secDetail").Controls("lbl2").Caption = ReadField(vRS.Fields("SchoolYear"))
    drSectionOfferingWithEnrolmentWithGrade.Sections("secDetail").Controls("lbl3").Caption = ReadField(vRS.Fields("SectionTeacherFullName"))
    drSectionOfferingWithEnrolmentWithGrade.Sections("secDetail").Controls("lbl4").Caption = ReadField(vRS.Fields("MaxStudentCount"))
    drSectionOfferingWithEnrolmentWithGrade.Sections("secDetail").Controls("lbl5").Caption = ReadField(vRS.Fields("MinGrade")) & " - " & ReadField(vRS.Fields("maxGrade"))
    drSectionOfferingWithEnrolmentWithGrade.Sections("secDetail").Controls("lblRoom").Caption = ReadField(vRS.Fields("Room"))
    drSectionOfferingWithEnrolmentWithGrade.Sections("secDetail").Controls("lblEnrollesCount").Caption = lEnrolessCount

    'get number of enrolles
    If GetEnrolmentCountBySectionOfferingID(curSectionOfferingID, lEnrolessCount) <> Success Then
        'fatal error
        CatchError "frmAddEnrolment", "ShowSectionDetail", "GetEnrolmentCountBySectionOfferingID(txtSectionOfferingID.Text, selSectionStudentCount) went failed"
    End If
    Set drSectionOfferingWithEnrolmentWithGrade.DataSource = vRS
    
    drSectionOfferingWithEnrolmentWithGrade.Show vbModal
    
    
ReleaseAndExit:
    If bNotEmp = True Then
        curSectionOfferingID = ""
    End If
    Set vRS = Nothing

End Sub





