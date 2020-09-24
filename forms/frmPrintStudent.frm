VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPrintStudent 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Print Student Info"
   ClientHeight    =   6600
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7815
   ControlBox      =   0   'False
   Icon            =   "frmPrintStudent.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   440
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   521
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList imgListIco 
      Left            =   6360
      Top             =   3000
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
            Picture         =   "frmPrintStudent.frx":08CA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox bgMain 
      Appearance      =   0  'Flat
      BackColor       =   &H00D8E9EC&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4995
      Left            =   90
      ScaleHeight     =   333
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   471
      TabIndex        =   0
      Top             =   480
      Width           =   7065
      Begin HSES.b8SContainer pbBGButton 
         Height          =   690
         Left            =   45
         TabIndex        =   3
         Top             =   360
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   1217
         BorderColor     =   14215660
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
               Picture         =   "frmPrintStudent.frx":0E64
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPrintStudent.frx":13FE
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin HSES.b8ChildTitleBar b8Title 
         Height          =   345
         Left            =   60
         TabIndex        =   1
         Top             =   30
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   609
         BackColor       =   12307149
         Caption         =   "Print Student && Enrolment Information"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontBold        =   -1  'True
         FontName        =   "Tahoma"
         FontSize        =   9.75
         ForeColor       =   8421504
         GradTheme       =   2
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
               Picture         =   "frmPrintStudent.frx":1998
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ListView listRecord 
         Height          =   4275
         Left            =   0
         TabIndex        =   2
         Top             =   1140
         Width           =   6105
         _ExtentX        =   10769
         _ExtentY        =   7541
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         HideColumnHeaders=   -1  'True
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         HotTracking     =   -1  'True
         _Version        =   393217
         Icons           =   "imgListIco"
         SmallIcons      =   "imgListIco"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
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
            Object.Width           =   7911
         EndProperty
      End
   End
   Begin VB.Shape b8cMain 
      BorderColor     =   &H00C0C0C0&
      Height          =   2055
      Left            =   3840
      Top             =   3810
      Width           =   3315
   End
End
Attribute VB_Name = "frmPrintStudent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim curSchoolYear As String
Dim curDepartmentID As String
Dim curYearLevelID As Integer
Dim curSectionOfferingID As String
Dim curStudentID As String
Dim curEnrolmentID As String


Public Function ShowForm(Optional sSchoolYear As String = "", Optional sDepartmentID As String = "", Optional iYearLevelID As Integer = 0, Optional sSectionOfferingID As String = "", Optional sStudentID As String = "", Optional sEnrolmentID As String = "")

    
    curSchoolYear = sSchoolYear
    curDepartmentID = sDepartmentID
    curYearLevelID = iYearLevelID
    curSectionOfferingID = sSectionOfferingID
    curStudentID = sStudentID
    curEnrolmentID = sEnrolmentID
    
    Me.Show

End Function






Public Sub RefreshReportList()
    
    listRecord.ListItems.Clear
    
    

    listRecord.ListItems.Add , "StudentAccountDetail", "Student Account Detail", 1, 1
    listRecord.ListItems.Add , "StudentAccountDetailByStudent", "Student Account Detail - Individual", 1, 1
        
    listRecord.ListItems.Add , , ""
    
    listRecord.ListItems.Add , "StudentCopyByEnrolment", "Student/School Copy By Student/YearLevel", 1, 1
    listRecord.ListItems.Add , "StudentCopyBySchoolYear", "Student/School Copy By School Year", 1, 1
        
    listRecord.ListItems.Add , , ""
    
    listRecord.ListItems.Add , "AllStudentList", "All Student List", 1, 1
    listRecord.ListItems.Add , "AllStudentListSchoolYearAsc", "All Student List (arrange by S.Y.,ascending)", 1, 1
    listRecord.ListItems.Add , "AllStudentListBySY", "All Student List By School Year", 1, 1
    listRecord.ListItems.Add , "AllStudentListByDepartment", "All Student List By Department", 1, 1
    listRecord.ListItems.Add , "AllStudentListBySYByDepartment", "All Student List By S.Y.,Department", 1, 1
    listRecord.ListItems.Add , "AllStudentListByYearLevel", "All Student List By Year Level", 1, 1
    listRecord.ListItems.Add , "AllStudentListBySYByYearLevel", "All Student List By S.Y.,Year Level", 1, 1

    listRecord.ListItems.Add , "AllStudentListByDepartmentByYL", "All Student List By Department,Y.L.", 1, 1
    
    listRecord.ListItems.Add , "AllStudentListBySYByDepartmentByYearLevel", "All Student List By S.Y.,Department,Year Level", 1, 1
    
    listRecord.ListItems.Add , , ""
    
    listRecord.ListItems.Add , "AllStudentListBySYMale", "All Student List By School Year, Male", 1, 1
    listRecord.ListItems.Add , "AllStudentListBySYFemale", "All Student List By School Year, Female", 1, 1
    
    End Sub


Public Sub Form_Activate()
    RefreshReportList
    
    mdiMain.RegMDIChild Me
End Sub

Public Sub Form_Load()
    listRecord.ColumnHeaders(1).Width = listRecord.Width - 6
End Sub



Private Sub Form_Resize()
    On Error Resume Next
    Me.ScaleMode = vbPixels
    
    
    
    b8cMain.Move Form_LeftMargin - 1, Form_TopMargin - 1, Me.ScaleWidth - (Form_LeftMargin - 1) * 2, Me.ScaleHeight - (Form_TopMargin - 1) * 2

    bgMain.Move Form_LeftMargin, Form_TopMargin, Me.ScaleWidth - Form_LeftMargin * 2, Me.ScaleHeight - Form_TopMargin * 2

    b8Title.Move 0, 0, bgMain.Width
    
    pbBGButton.Move 0, b8Title.Height, bgMain.Width
    

    listRecord.Move listRecord.Left, pbBGButton.Top + pbBGButton.Height, bgMain.Width - (listRecord.Left * 2), Me.ScaleHeight - (pbBGButton.Top + pbBGButton.Height)
    
End Sub

Public Sub listRecord_DblClick()
    
    Select Case listRecord.SelectedItem.Key
        
        
        Case "StudentAccountDetail"
            Call ShowStudentAccountDetail
        
        Case "StudentAccountDetailByStudent"
            Call ShowStudentAccountDetailByStudent
        
        
        Case "StudentCopyBySchoolYear"
            Call ShowStudentCopyBySchoolYear
            
        Case "StudentCopyByEnrolment"
            Call ShowStudentCopyByStudentYearlevel
            
            
        Case "AllStudentList"
            Call ShowAllStudentList
        
        Case "AllStudentListSchoolYearAsc"
            Call ShowAllStudentListSchoolYearAsc
        
        Case "AllStudentListBySY"
            Call ShowAllStudentListBySY
        
        Case "AllStudentListByDepartment"
            Call ShowAllStudentListByDepartment
        
        Case "AllStudentListBySYByDepartment"
            Call ShowAllStudentListBySYByDepartment
        
        Case "AllStudentListByYearLevel"
            Call ShowAllStudentListByYearLevel
            
        Case "AllStudentListBySYByYearLevel"
            Call ShowAllStudentListBySYByYearLevel
            
        Case "AllStudentListByDepartmentByYL"
            Call ShowAllStudentListByDepartmentByYearLevel
        
        Case "AllStudentListBySYByDepartmentByYearLevel"
            Call ShowAllStudentListBySYByDepartmentByYearLevel
            
        Case "AllStudentListBySYMale"
            Call ShowAllStudentListBySYByGender("Male")
            
        Case "AllStudentListBySYFemale"
            Call ShowAllStudentListBySYByGender("Female")
 
    End Select
        
    
End Sub

Public Function ShowStudentCopyByStudentYearlevel()
    
    Dim sStudentID As String
    Dim iYearLevelID As Integer
    Dim sYearLevelTitle As String
    Dim sSchoolYear As String
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    sStudentID = PickStudent.GetStudentID
    
    If sStudentID = "" Then
        GoTo ReleaseAndExit
    End If
    
    GetLatestSchoolYearYearLevel sStudentID, sSchoolYear, iYearLevelID
    
    If iYearLevelID < 1 Then
        MsgBox "Student not yet enroled", vbExclamation
        GoTo ReleaseAndExit
    End If
    
    
    sYearLevelTitle = PickYearLevel.GetYearLevelTitle
    
    If sYearLevelTitle = "" Then
        GoTo ReleaseAndExit
    End If
    
    iYearLevelID = YLTitleToID(sYearLevelTitle)


    sSQL = "SELECT tblEnrolment.EnrolmentID, tblEnrolment.StudentID, tblYearLevel.YearLevelID" & _
            " FROM tblYearLevel INNER JOIN (tblSection INNER JOIN (tblSectionOffering INNER JOIN tblEnrolment ON tblSectionOffering.SectionOfferingID = tblEnrolment.SectionOfferingID) ON tblSection.SectionID = tblSectionOffering.SectionID) ON tblYearLevel.YearLevelID = tblSection.YearLevelID" & _
            " WHERE (((tblEnrolment.StudentID)='" & sStudentID & "') AND ((tblYearLevel.YearLevelID)=" & iYearLevelID & "));"

    If ConnectRS(HSESDB, vRS, sSQL) = False Then
        'fatal error
        'temp
        MsgBox "error"
        GoTo ReleaseAndExit
    End If

    If AnyRecordExisted(vRS) = False Then
        MsgBox "The selected Student is not enroled in the selected Year Level", vbExclamation
        GoTo ReleaseAndExit
    End If
    
    ShowStudentCopyByEnrolment ReadField(vRS.Fields("enrolmentid"))

ReleaseAndExit:
    Set vRS = Nothing
End Function

Public Function ShowStudentCopyBySchoolYear()
    Dim sSchoolYear As String
    
    
    sSchoolYear = PickSchoolYear.GetItem
    
    If sSchoolYear = "" Then
        Exit Function
    End If
    
    frmPrintEnrolment.ShowForm sSchoolYear
    
End Function

Public Sub ShowAllStudentList()
    
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    sSQL = "SELECT [LastName] & ', ' & [Firstname] & ' ' & [MiddleName] AS StudentFullName, tblEnrolment.SchoolYear, tblDepartment.DepartmentTitle, [tblYearLevel]![YearLevelTitle] & ' - ' & [tblSection]![SectionTitle] AS SectionFullTitle" & _
            " FROM tblYearLevel INNER JOIN (tblStudent INNER JOIN ((tblDepartment INNER JOIN tblSection ON tblDepartment.DepartmentID = tblSection.DepartmentID) INNER JOIN (tblSectionOffering INNER JOIN tblEnrolment ON tblSectionOffering.SectionOfferingID = tblEnrolment.SectionOfferingID) ON tblSection.SectionID = tblSectionOffering.SectionID) ON tblStudent.StudentID = tblEnrolment.StudentID) ON tblYearLevel.YearLevelID = tblSection.YearLevelID" & _
            " ORDER BY [LastName] & ', ' & [Firstname] & ' ' & [MiddleName];"

    'set mouse pointer
    Me.MousePointer = vbHourglass
    
    If ConnectRS(HSESDB, vRS, sSQL) = False Then
        MsgBox "error"
        GoTo ReleaseAndExit
    End If
    
    Set drAllStudentList.DataSource = vRS
    
    'set mouse pointer
    Me.MousePointer = vbDefault
    drAllStudentList.Show vbModal
    
ReleaseAndExit:
    Set vRS = Nothing
End Sub
Public Sub ShowAllStudentListSchoolYearAsc()
    
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    sSQL = "SELECT [LastName] & ', ' & [Firstname] & ' ' & [MiddleName] AS StudentFullName, tblEnrolment.SchoolYear, tblDepartment.DepartmentTitle, [tblYearLevel]![YearLevelTitle] & ' - ' & [tblSection]![SectionTitle] AS SectionFullTitle" & _
            " FROM tblYearLevel INNER JOIN (tblStudent INNER JOIN ((tblDepartment INNER JOIN tblSection ON tblDepartment.DepartmentID = tblSection.DepartmentID) INNER JOIN (tblSectionOffering INNER JOIN tblEnrolment ON tblSectionOffering.SectionOfferingID = tblEnrolment.SectionOfferingID) ON tblSection.SectionID = tblSectionOffering.SectionID) ON tblStudent.StudentID = tblEnrolment.StudentID) ON tblYearLevel.YearLevelID = tblSection.YearLevelID" & _
            " ORDER BY [LastName] & ', ' & [Firstname] & ' ' & [MiddleName];"

    'set mouse pointer
    Me.MousePointer = vbHourglass
    
    If ConnectRS(HSESDB, vRS, sSQL) = False Then
        MsgBox "error"
        GoTo ReleaseAndExit
    End If
    
    Set drAllStudentList.DataSource = vRS
    'set mouse pointer
    Me.MousePointer = vbDefault
    drAllStudentList.Show vbModal
    
ReleaseAndExit:
    Set vRS = Nothing
End Sub

Public Sub ShowAllStudentListBySY()
    
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    Dim sSchoolYear As String
    
    
    sSchoolYear = PickSchoolYear.GetItem
    
    If sSchoolYear = "" Then
        GoTo ReleaseAndExit
    End If
    
    sSQL = "SELECT [LastName] & ', ' & [Firstname] & ' ' & [MiddleName] AS StudentFullName, tblEnrolment.SchoolYear, tblDepartment.DepartmentTitle, [tblYearLevel]![YearLevelTitle] & ' - ' & [tblSection]![SectionTitle] AS SectionFullTitle" & _
            " FROM tblYearLevel INNER JOIN (tblStudent INNER JOIN ((tblDepartment INNER JOIN tblSection ON tblDepartment.DepartmentID = tblSection.DepartmentID) INNER JOIN (tblSectionOffering INNER JOIN tblEnrolment ON tblSectionOffering.SectionOfferingID = tblEnrolment.SectionOfferingID) ON tblSection.SectionID = tblSectionOffering.SectionID) ON tblStudent.StudentID = tblEnrolment.StudentID) ON tblYearLevel.YearLevelID = tblSection.YearLevelID" & _
            " WHERE tblSectionOffering.SchoolYear='" & sSchoolYear & "'" & _
            " ORDER BY [LastName] & ', ' & [Firstname] & ' ' & [MiddleName];"

    'set mouse pointer
    Me.MousePointer = vbHourglass
    
    If ConnectRS(HSESDB, vRS, sSQL) = False Then
        MsgBox "error"
        GoTo ReleaseAndExit
    End If
    
    drAllStudentListBySY.Sections("Section4").Controls("lblSY").Caption = sSchoolYear
    Set drAllStudentListBySY.DataSource = vRS
    'set mouse pointer
    Me.MousePointer = vbDefault
    drAllStudentListBySY.Show vbModal
    
ReleaseAndExit:
    Set vRS = Nothing
End Sub

Public Sub ShowAllStudentListByDepartment()
    
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    Dim sDepartmentID As String
    Dim sDepartmentTitle As String
    
    
    sDepartmentID = PickDepartment.GetItem(, sDepartmentTitle)

    If sDepartmentID = "" Then
        GoTo ReleaseAndExit
    End If
    
    sSQL = "SELECT [LastName] & ', ' & [Firstname] & ' ' & [MiddleName] AS StudentFullName, tblEnrolment.SchoolYear, tblDepartment.DepartmentTitle, [tblYearLevel]![YearLevelTitle] & ' - ' & [tblSection]![SectionTitle] AS SectionFullTitle" & _
            " FROM tblYearLevel INNER JOIN (tblStudent INNER JOIN ((tblDepartment INNER JOIN tblSection ON tblDepartment.DepartmentID = tblSection.DepartmentID) INNER JOIN (tblSectionOffering INNER JOIN tblEnrolment ON tblSectionOffering.SectionOfferingID = tblEnrolment.SectionOfferingID) ON tblSection.SectionID = tblSectionOffering.SectionID) ON tblStudent.StudentID = tblEnrolment.StudentID) ON tblYearLevel.YearLevelID = tblSection.YearLevelID" & _
            " WHERE tblDepartment.DepartmentID='" & sDepartmentID & "'" & _
            " ORDER BY [LastName] & ', ' & [Firstname] & ' ' & [MiddleName];"

    'set mouse pointer
    Me.MousePointer = vbHourglass
    
    If ConnectRS(HSESDB, vRS, sSQL) = False Then
        MsgBox "error"
        GoTo ReleaseAndExit
    End If
    
    drAllStudentListByDepartment.Sections("Section4").Controls("lblDepartmentTitle").Caption = sDepartmentTitle
    Set drAllStudentListByDepartment.DataSource = vRS
    'set mouse pointer
    Me.MousePointer = vbDefault
    drAllStudentListByDepartment.Show vbModal
    
ReleaseAndExit:
    Set vRS = Nothing
End Sub



Public Sub ShowAllStudentListByYearLevel()
    
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    Dim sYearLevelTitle As String
    
    sYearLevelTitle = PickYearLevel.GetYearLevelTitle
    
    If sYearLevelTitle = "" Then
        GoTo ReleaseAndExit
    End If
    
    sSQL = "SELECT [LastName] & ', ' & [Firstname] & ' ' & [MiddleName] AS StudentFullName, tblEnrolment.SchoolYear, tblDepartment.DepartmentTitle, [tblYearLevel]![YearLevelTitle] & ' - ' & [tblSection]![SectionTitle] AS SectionFullTitle" & _
            " FROM tblYearLevel INNER JOIN (tblStudent INNER JOIN ((tblDepartment INNER JOIN tblSection ON tblDepartment.DepartmentID = tblSection.DepartmentID) INNER JOIN (tblSectionOffering INNER JOIN tblEnrolment ON tblSectionOffering.SectionOfferingID = tblEnrolment.SectionOfferingID) ON tblSection.SectionID = tblSectionOffering.SectionID) ON tblStudent.StudentID = tblEnrolment.StudentID) ON tblYearLevel.YearLevelID = tblSection.YearLevelID" & _
            " WHERE tblYearLevel.YearLevelTitle='" & sYearLevelTitle & "'" & _
            " ORDER BY tblYearLevel.YearLevelTitle;"

    'set mouse pointer
    Me.MousePointer = vbHourglass
    
    If ConnectRS(HSESDB, vRS, sSQL) = False Then
        MsgBox "error"
        GoTo ReleaseAndExit
    End If
    
    Set drAllStudentList.DataSource = vRS
    
    'set mouse pointer
    Me.MousePointer = vbDefault
    drAllStudentList.Show vbModal
    
ReleaseAndExit:
    Set vRS = Nothing
End Sub


Public Sub ShowAllStudentListByDepartmentByYearLevel()
    
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    Dim sDepartmentTitle As String
    Dim sDepartmentID As String
    Dim sYearLevelTitle As String
    
    
    sDepartmentID = PickDepartment.GetItem(, sDepartmentTitle)
    
    If sDepartmentID = "" Then
        GoTo ReleaseAndExit
    End If
    
    sYearLevelTitle = PickYearLevel.GetYearLevelTitle
    
    If sYearLevelTitle = "" Then
        GoTo ReleaseAndExit
    End If
    
    sSQL = "SELECT [LastName] & ', ' & [Firstname] & ' ' & [MiddleName] AS StudentFullName, tblEnrolment.SchoolYear, tblDepartment.DepartmentTitle, [tblYearLevel]![YearLevelTitle] & ' - ' & [tblSection]![SectionTitle] AS SectionFullTitle" & _
            " FROM tblYearLevel INNER JOIN (tblStudent INNER JOIN ((tblDepartment INNER JOIN tblSection ON tblDepartment.DepartmentID = tblSection.DepartmentID) INNER JOIN (tblSectionOffering INNER JOIN tblEnrolment ON tblSectionOffering.SectionOfferingID = tblEnrolment.SectionOfferingID) ON tblSection.SectionID = tblSectionOffering.SectionID) ON tblStudent.StudentID = tblEnrolment.StudentID) ON tblYearLevel.YearLevelID = tblSection.YearLevelID" & _
            " WHERE tblDepartment.DepartmentID='" & sDepartmentID & "' AND tblYearLevel.YearLevelTitle='" & sYearLevelTitle & "'" & _
            " ORDER BY tblYearLevel.YearLevelTitle;"

    'set mouse pointer
    Me.MousePointer = vbHourglass
    
    If ConnectRS(HSESDB, vRS, sSQL) = False Then
        MsgBox "error"
        GoTo ReleaseAndExit
    End If
    
    
    drAllStudentListByDepartmentByYearLevel.Sections("Section4").Controls("lblDepartmentTitle").Caption = sDepartmentTitle
    drAllStudentListByDepartmentByYearLevel.Sections("Section4").Controls("lblYearLevelTitle").Caption = sYearLevelTitle
    Set drAllStudentListByDepartmentByYearLevel.DataSource = vRS
    
    'set mouse pointer
    Me.MousePointer = vbDefault
    drAllStudentListByDepartmentByYearLevel.Show vbModal
    
ReleaseAndExit:
    Set vRS = Nothing
End Sub


'AllStudentListBySYByDepartment
Public Sub ShowAllStudentListBySYByDepartment()
    
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    Dim sDepartmentID As String
    Dim sDepartmentTitle As String
    Dim sSchoolYear As String
    
    If curSchoolYear = "" Then
        sSchoolYear = PickSchoolYear.GetItem
    Else
        sSchoolYear = curSchoolYear
    End If
    
    If sSchoolYear = "" Then
        GoTo ReleaseAndExit
    End If
    
    sDepartmentID = PickDepartment.GetItem(, sDepartmentTitle)

    If sDepartmentID = "" Then
        GoTo ReleaseAndExit
    End If
    
    sSQL = "SELECT [LastName] & ', ' & [Firstname] & ' ' & [MiddleName] AS StudentFullName, tblEnrolment.SchoolYear, tblDepartment.DepartmentTitle, [tblYearLevel]![YearLevelTitle] & ' - ' & [tblSection]![SectionTitle] AS SectionFullTitle" & _
            " FROM tblYearLevel INNER JOIN (tblStudent INNER JOIN ((tblDepartment INNER JOIN tblSection ON tblDepartment.DepartmentID = tblSection.DepartmentID) INNER JOIN (tblSectionOffering INNER JOIN tblEnrolment ON tblSectionOffering.SectionOfferingID = tblEnrolment.SectionOfferingID) ON tblSection.SectionID = tblSectionOffering.SectionID) ON tblStudent.StudentID = tblEnrolment.StudentID) ON tblYearLevel.YearLevelID = tblSection.YearLevelID" & _
            " WHERE tblSectionOffering.SchoolYear='" & sSchoolYear & "' AND tblDepartment.DepartmentID='" & sDepartmentID & "'" & _
            " ORDER BY [LastName] & ', ' & [Firstname] & ' ' & [MiddleName];"

    'set mouse pointer
    Me.MousePointer = vbHourglass
    
    If ConnectRS(HSESDB, vRS, sSQL) = False Then
        MsgBox "error"
        GoTo ReleaseAndExit
    End If
    
    drAllStudentListBySYByDepartment.Sections("Section4").Controls("lblDepartmentTitle").Caption = sDepartmentTitle
    drAllStudentListBySYByDepartment.Sections("Section4").Controls("lblSchoolYear").Caption = sSchoolYear

    Set drAllStudentListBySYByDepartment.DataSource = vRS
    'set mouse pointer
    Me.MousePointer = vbDefault
    drAllStudentListBySYByDepartment.Show vbModal
    
    
    
ReleaseAndExit:
    Set vRS = Nothing
End Sub

Public Sub ShowAllStudentListBySYByYearLevel()
    
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    Dim sSchoolYear As String
    Dim sYearLevelTitle As String
    
    If curSchoolYear = "" Then
        sSchoolYear = PickSchoolYear.GetItem
    Else
        sSchoolYear = curSchoolYear
    End If
    
    If sSchoolYear = "" Then
        GoTo ReleaseAndExit
    End If
    
    
    
    sYearLevelTitle = PickYearLevel.GetYearLevelTitle
    
    If sYearLevelTitle = "" Then
        GoTo ReleaseAndExit
    End If
    
    
    sSQL = "SELECT [LastName] & ', ' & [Firstname] & ' ' & [MiddleName] AS StudentFullName, tblEnrolment.SchoolYear, tblDepartment.DepartmentTitle, [tblYearLevel]![YearLevelTitle] & ' - ' & [tblSection]![SectionTitle] AS SectionFullTitle" & _
            " FROM tblYearLevel INNER JOIN (tblStudent INNER JOIN ((tblDepartment INNER JOIN tblSection ON tblDepartment.DepartmentID = tblSection.DepartmentID) INNER JOIN (tblSectionOffering INNER JOIN tblEnrolment ON tblSectionOffering.SectionOfferingID = tblEnrolment.SectionOfferingID) ON tblSection.SectionID = tblSectionOffering.SectionID) ON tblStudent.StudentID = tblEnrolment.StudentID) ON tblYearLevel.YearLevelID = tblSection.YearLevelID" & _
            " WHERE tblSectionOffering.SchoolYear='" & sSchoolYear & "' AND tblYearLevel.YearLevelTitle='" & sYearLevelTitle & "'" & _
            " ORDER BY [LastName] & ', ' & [Firstname] & ' ' & [MiddleName];"


    'set mouse pointer
    Me.MousePointer = vbHourglass
    
    If ConnectRS(HSESDB, vRS, sSQL) = False Then
        MsgBox "error"
        GoTo ReleaseAndExit
    End If
    
    
    
    drAllStudentListBySYByYearLevel.Sections("Section4").Controls("lblYearLevelTitle").Caption = sYearLevelTitle
    drAllStudentListBySYByYearLevel.Sections("Section4").Controls("lblSchoolYear").Caption = sSchoolYear

    Set drAllStudentListBySYByYearLevel.DataSource = vRS
    'set mouse pointer
    Me.MousePointer = vbDefault
    drAllStudentListBySYByYearLevel.Show vbModal

    
ReleaseAndExit:
    Set vRS = Nothing
End Sub


Public Sub ShowAllStudentListBySYByDepartmentByYearLevel()
    
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    Dim sSchoolYear As String
    Dim sDepartmentTitle As String
    Dim sDepartmentID As String
    Dim sYearLevelTitle As String
    
    If curSchoolYear = "" Then
        sSchoolYear = PickSchoolYear.GetItem
    Else
        sSchoolYear = curSchoolYear
    End If
    
    If sSchoolYear = "" Then
        GoTo ReleaseAndExit
    End If
    
    sDepartmentID = PickDepartment.GetItem(, sDepartmentTitle)

    If sDepartmentID = "" Then
        GoTo ReleaseAndExit
    End If
    
    sYearLevelTitle = PickYearLevel.GetYearLevelTitle
    
    If sYearLevelTitle = "" Then
        GoTo ReleaseAndExit
    End If
    
    
    sSQL = "SELECT [LastName] & ', ' & [Firstname] & ' ' & [MiddleName] AS StudentFullName, tblEnrolment.SchoolYear, tblDepartment.DepartmentTitle, [tblYearLevel]![YearLevelTitle] & ' - ' & [tblSection]![SectionTitle] AS SectionFullTitle" & _
            " FROM tblYearLevel INNER JOIN (tblStudent INNER JOIN ((tblDepartment INNER JOIN tblSection ON tblDepartment.DepartmentID = tblSection.DepartmentID) INNER JOIN (tblSectionOffering INNER JOIN tblEnrolment ON tblSectionOffering.SectionOfferingID = tblEnrolment.SectionOfferingID) ON tblSection.SectionID = tblSectionOffering.SectionID) ON tblStudent.StudentID = tblEnrolment.StudentID) ON tblYearLevel.YearLevelID = tblSection.YearLevelID" & _
            " WHERE tblSectionOffering.SchoolYear='" & sSchoolYear & "' AND tblDepartment.DepartmentID='" & sDepartmentID & "'AND tblYearLevel.YearLevelTitle='" & sYearLevelTitle & "'" & _
            " ORDER BY [LastName] & ', ' & [Firstname] & ' ' & [MiddleName];"


    'set mouse pointer
    Me.MousePointer = vbHourglass
    
    If ConnectRS(HSESDB, vRS, sSQL) = False Then
        MsgBox "error"
        GoTo ReleaseAndExit
    End If
    
    
    
    drAllStudentListBySYByDepartmentByYearLevel.Sections("Section4").Controls("lblYearLevelTitle").Caption = sYearLevelTitle
    drAllStudentListBySYByDepartmentByYearLevel.Sections("Section4").Controls("lblSchoolYear").Caption = sSchoolYear
    drAllStudentListBySYByDepartmentByYearLevel.Sections("Section4").Controls("lblDepartmentTitle").Caption = sDepartmentTitle

    Set drAllStudentListBySYByDepartmentByYearLevel.DataSource = vRS
    'set mouse pointer
    Me.MousePointer = vbDefault
    drAllStudentListBySYByDepartmentByYearLevel.Show vbModal

    
ReleaseAndExit:
    Set vRS = Nothing
End Sub


Public Sub ShowAllStudentListBySYByGender(sGender As String)
    
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    Dim sSchoolYear As String
    
    
    sSchoolYear = PickSchoolYear.GetItem
    
    If sSchoolYear = "" Then
        GoTo ReleaseAndExit
    End If
    
    sSQL = "SELECT [LastName] & ', ' & [Firstname] & ' ' & [MiddleName] AS StudentFullName, tblEnrolment.SchoolYear, tblDepartment.DepartmentTitle, [tblYearLevel]![YearLevelTitle] & ' - ' & [tblSection]![SectionTitle] AS SectionFullTitle" & _
            " FROM tblYearLevel INNER JOIN (tblStudent INNER JOIN ((tblDepartment INNER JOIN tblSection ON tblDepartment.DepartmentID = tblSection.DepartmentID) INNER JOIN (tblSectionOffering INNER JOIN tblEnrolment ON tblSectionOffering.SectionOfferingID = tblEnrolment.SectionOfferingID) ON tblSection.SectionID = tblSectionOffering.SectionID) ON tblStudent.StudentID = tblEnrolment.StudentID) ON tblYearLevel.YearLevelID = tblSection.YearLevelID" & _
            " WHERE tblSectionOffering.SchoolYear='" & sSchoolYear & "' AND tblStudent.Gender='" & sGender & "'" & _
            " ORDER BY [LastName] & ', ' & [Firstname] & ' ' & [MiddleName];"

    'set mouse pointer
    Me.MousePointer = vbHourglass
    
    If ConnectRS(HSESDB, vRS, sSQL) = False Then
        MsgBox "error"
        GoTo ReleaseAndExit
    End If
    
    drAllStudentListBySY.Sections("Section4").Controls("lblSY").Caption = sSchoolYear
    Set drAllStudentListBySY.DataSource = vRS
    'set mouse pointer
    Me.MousePointer = vbDefault
    drAllStudentListBySY.Show vbModal
    
ReleaseAndExit:
    Set vRS = Nothing
End Sub

Public Sub ShowStudentAccountDetail()

    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    sSQL = "SELECT tblStudent.StudentID, [tblStudent]![LastName] & ', ' & [tblStudent]![FirstName] & ' ' & [tblStudent]![MiddleName] AS StudentFullName, tblStudent.Gender, tblStudent.Status, tblStudent.Citizenship, tblStudent.BirthDate, tblStudent.PlaceOfBirth, tblStudent.HomeAddress, tblStudent.CityAddress, tblStudent.BloodType, tblStudent.Religion, tblStudent.LastSchoolName, tblStudent.LastSchoolContactNumber, tblStudent.LastSchoolAddress, tblStudent.MotherOccupation, tblStudent.MotherName, tblStudent.MotherOccupation, tblStudent.FatherName, tblStudent.FatherOccupation, tblStudent.ParentsContactNumber, tblStudent.ParentsAddress, tblStudent.GuardianName, tblStudent.GuardianAddress, tblStudent.GuardianContactNumber, tblStudent.OldAveGrade, tblStudent.CreationDate, tblStudent.CreationDate, tblStudent.CreatedBy, tblStudent.ModifiedDate, tblStudent.ModifiedBy" & _
            " FROM tblStudent" & _
            " ORDER BY [tblStudent]![LastName] & ', ' & [tblStudent]![FirstName] & ' ' & [tblStudent]![MiddleName]"

    
    If ConnectRS(HSESDB, vRS, sSQL) = False Then
        MsgBox "error"
        GoTo ReleaseAndExit
    End If
    
    
    Set drStudentDetail.DataSource = vRS
    drStudentDetail.Show vbModal
    
ReleaseAndExit:
    Set vRS = Nothing
End Sub


Public Sub ShowStudentAccountDetailByStudent(Optional sStudentID As String = "")

    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
   
    If sStudentID <> "" Then
        curStudentID = sStudentID
    End If
    
    If curStudentID = "" Then
        sStudentID = PickStudent.GetStudentID
    Else
        sStudentID = curStudentID
    End If
    
    
    If sStudentID = "" Then
        GoTo ReleaseAndExit
    End If
    
    sSQL = "SELECT tblStudent.StudentID, [tblStudent]![LastName] & ', ' & [tblStudent]![FirstName] & ' ' & [tblStudent]![MiddleName] AS StudentFullName, tblStudent.Gender, tblStudent.Status, tblStudent.Citizenship, tblStudent.BirthDate, tblStudent.PlaceOfBirth, tblStudent.HomeAddress, tblStudent.CityAddress, tblStudent.BloodType, tblStudent.Religion, tblStudent.LastSchoolName, tblStudent.LastSchoolContactNumber, tblStudent.LastSchoolAddress, tblStudent.MotherOccupation, tblStudent.MotherName, tblStudent.MotherOccupation, tblStudent.FatherName, tblStudent.FatherOccupation, tblStudent.ParentsContactNumber, tblStudent.ParentsAddress, tblStudent.GuardianName, tblStudent.GuardianAddress, tblStudent.GuardianContactNumber, tblStudent.OldAveGrade, tblStudent.CreationDate, tblStudent.CreationDate, tblStudent.CreatedBy, tblStudent.ModifiedDate, tblStudent.ModifiedBy" & _
            " FROM tblStudent" & _
            " WHERE tblStudent.StudentID='" & sStudentID & "'" & _
            " ORDER BY [tblStudent]![LastName] & ', ' & [tblStudent]![FirstName] & ' ' & [tblStudent]![MiddleName]"

    
    
    If ConnectRS(HSESDB, vRS, sSQL) = False Then
        MsgBox "error"
        GoTo ReleaseAndExit
    End If
    
    
    Set drStudentDetail.DataSource = vRS
    drStudentDetail.Show vbModal
    
ReleaseAndExit:
    Set vRS = Nothing
End Sub



Public Sub ShowStudentCopyByEnrolment(Optional sEnrolmentID As String = "")

    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
   
    
    
    If sEnrolmentID = "" Then
        GoTo ReleaseAndExit
    End If
    
    sSQL = "SELECT tblEnrolment.SchoolYear, tblDepartment.DepartmentTitle, [tblYearLevel]![YearLevelTitle] & ' - ' & [tblSection]![SectionTitle] AS SectionFullTitle, Left([tblteacher_1]![FirstName],1) & '. ' & [tblteacher_1]![LastName] AS AdviserFullName, tblStudent.StudentID, [tblStudent]![LastName] & ', ' & [tblStudent]![FirstName] & ' ' & [tblStudent]![MiddleName] AS StudentFullName, tblStudent.Gender, tblStudent.CityAddress, tblStudent.HomeAddress, tblSubject.SubjectTitle, [tblSubjectOffering]![SchedTimeStart] & ' - ' & [tblSubjectOffering]![SchedTimeEnd] AS TimeSchedule, tblSubjectOffering.Days, Left([tblteacher]![FirstName],1) & '. ' & [tblteacher]![LastName] AS TeacherFullName, tblEnrolment.EnrolmentID" & _
            " FROM tblTeacher AS tblTeacher_1 INNER JOIN (tblYearLevel INNER JOIN ((tblDepartment INNER JOIN tblSection ON tblDepartment.DepartmentID = tblSection.DepartmentID) INNER JOIN (tblTeacher INNER JOIN (tblSubject INNER JOIN (tblStudent INNER JOIN ((tblSectionOffering INNER JOIN tblEnrolment ON tblSectionOffering.SectionOfferingID = tblEnrolment.SectionOfferingID) INNER JOIN tblSubjectOffering ON tblSectionOffering.SectionOfferingID = tblSubjectOffering.SectionOfferingID) ON tblStudent.StudentID = tblEnrolment.StudentID) ON tblSubject.SubjectID = tblSubjectOffering.SubjectID) ON tblTeacher.TeacherID = tblSubjectOffering.TeacherID) ON tblSection.SectionID = tblSectionOffering.SectionID) ON tblYearLevel.YearLevelID = tblSection.YearLevelID) ON tblTeacher_1.TeacherID = tblSectionOffering.TeacherID" & _
            " WHERE (((tblEnrolment.EnrolmentID)='" & sEnrolmentID & "'));"

    
    If ConnectRS(HSESDB, vRS, sSQL) = False Then
        MsgBox "error"
        GoTo ReleaseAndExit
    End If
    
    
    
    Set drEnrolmentDetail.DataSource = vRS
    
    drEnrolmentDetail.Sections("secDetail").Controls("lblStudentFullName").Caption = ReadField(vRS.Fields("StudentFullName"))
    drEnrolmentDetail.Sections("secDetail").Controls("lblGender").Caption = ReadField(vRS.Fields("Gender"))
    drEnrolmentDetail.Sections("secDetail").Controls("lblCityAddress").Caption = ReadField(vRS.Fields("CityAddress"))
    drEnrolmentDetail.Sections("secDetail").Controls("lblHomeAddress").Caption = ReadField(vRS.Fields("HomeAddress"))
    drEnrolmentDetail.Sections("secDetail").Controls("lblStudentID").Caption = ReadField(vRS.Fields("StudentID"))
    
    drEnrolmentDetail.Sections("secDetail").Controls("lblSectionFullTitle").Caption = ReadField(vRS.Fields("SectionFullTitle"))
    drEnrolmentDetail.Sections("secDetail").Controls("lblSchoolYear").Caption = ReadField(vRS.Fields("SchoolYear"))
    drEnrolmentDetail.Sections("secDetail").Controls("lblDepartmentTitle").Caption = ReadField(vRS.Fields("DepartmentTitle"))
    drEnrolmentDetail.Sections("secDetail").Controls("lblAdviserFullName").Caption = ReadField(vRS.Fields("AdviserFullName"))

    drEnrolmentDetail.Show vbModal
    
ReleaseAndExit:
    Set vRS = Nothing
End Sub




