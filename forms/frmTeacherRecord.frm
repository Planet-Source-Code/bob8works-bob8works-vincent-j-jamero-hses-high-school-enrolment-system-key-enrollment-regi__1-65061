VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmTeacherRecord 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Teacher's Record"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9810
   ControlBox      =   0   'False
   Icon            =   "frmTeacherRecord.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   573
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   654
   WindowState     =   2  'Maximized
   Begin VB.PictureBox bgMain 
      Appearance      =   0  'Flat
      BackColor       =   &H00D8E9EC&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   7575
      Left            =   225
      ScaleHeight     =   505
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   651
      TabIndex        =   0
      Top             =   105
      Width           =   9765
      Begin MSComctlLib.ImageCombo cmbSchoolYear 
         Height          =   330
         Left            =   5850
         TabIndex        =   18
         Top             =   420
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   582
         _Version        =   393216
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ImageList       =   "ilListIco"
      End
      Begin HSES.b8Line b8Line2 
         Height          =   60
         Left            =   -30
         TabIndex        =   6
         Top             =   810
         Width           =   15360
         _ExtentX        =   27093
         _ExtentY        =   106
      End
      Begin VB.CommandButton cmdGetTeacher 
         BackColor       =   &H00D8E9EC&
         Height          =   285
         Left            =   5220
         Picture         =   "frmTeacherRecord.frx":0ECA
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   450
         Width           =   345
      End
      Begin VB.TextBox txtTeacherFullName 
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
         Left            =   870
         Locked          =   -1  'True
         MaxLength       =   17
         TabIndex        =   3
         Top             =   420
         Width           =   4710
      End
      Begin HSES.b8ChildTitleBar b8Title 
         Height          =   345
         Left            =   15
         TabIndex        =   1
         Top             =   0
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   609
         BackColor       =   14215660
         Caption         =   "Teacher's Record"
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
      Begin HSES.b8Line b8Line1 
         Height          =   60
         Index           =   4
         Left            =   0
         TabIndex        =   7
         Top             =   1710
         Width           =   15360
         _ExtentX        =   27093
         _ExtentY        =   106
      End
      Begin HSES.b8Container bgDetail 
         Height          =   855
         Left            =   0
         TabIndex        =   8
         Top             =   870
         Width           =   10005
         _ExtentX        =   17648
         _ExtentY        =   1508
         BorderColor     =   12307149
         BackColor       =   16185592
         ShadowColor1    =   13427430
         ShadowColor2    =   14215660
         Begin VB.Label lblLoginName 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "--"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C25418&
            Height          =   240
            Left            =   1155
            TabIndex        =   20
            Top             =   75
            Width           =   180
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Log-in Name:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   19
            Top             =   135
            Width           =   945
         End
         Begin VB.Label lblAddress 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "--"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C25418&
            Height          =   195
            Left            =   4035
            TabIndex        =   14
            Top             =   480
            Width           =   150
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Address:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   3360
            TabIndex        =   13
            Top             =   465
            Width           =   645
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Name:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   3375
            TabIndex        =   12
            Top             =   135
            Width           =   465
         End
         Begin VB.Label lblTeacherName 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "--"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C25418&
            Height          =   345
            Left            =   3915
            TabIndex        =   11
            Top             =   75
            Width           =   240
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Contact Number:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   105
            TabIndex        =   10
            Top             =   480
            Width           =   1230
         End
         Begin VB.Label lblContactNumber 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "--"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C25418&
            Height          =   195
            Left            =   1380
            TabIndex        =   9
            Top             =   495
            Width           =   150
         End
      End
      Begin TabDlg.SSTab tabMAin 
         Height          =   5505
         Left            =   30
         TabIndex        =   15
         Top             =   1785
         Width           =   9735
         _ExtentX        =   17171
         _ExtentY        =   9710
         _Version        =   393216
         Style           =   1
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   529
         Enabled         =   0   'False
         BackColor       =   14215660
         TabCaption(0)   =   "Section Advisory"
         TabPicture(0)   =   "frmTeacherRecord.frx":1454
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "bgTabCon(0)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "ilSO"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).ControlCount=   2
         TabCaption(1)   =   "Subject Teached"
         TabPicture(1)   =   "frmTeacherRecord.frx":1470
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "bgTabCon(1)"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).ControlCount=   1
         Begin MSComctlLib.ImageList ilSO 
            Left            =   3495
            Top             =   60
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
                  Picture         =   "frmTeacherRecord.frx":148C
                  Key             =   ""
               EndProperty
            EndProperty
         End
         Begin HSES.b8SContainer bgTabCon 
            Height          =   5160
            Index           =   0
            Left            =   15
            TabIndex        =   16
            Top             =   300
            Width           =   9600
            _ExtentX        =   16933
            _ExtentY        =   9102
            BorderColor     =   12307149
            Begin MSComctlLib.ListView listAdvisory 
               Height          =   4365
               Left            =   30
               TabIndex        =   21
               Top             =   330
               Width           =   8970
               _ExtentX        =   15822
               _ExtentY        =   7699
               View            =   3
               LabelEdit       =   1
               LabelWrap       =   -1  'True
               HideSelection   =   0   'False
               FullRowSelect   =   -1  'True
               GridLines       =   -1  'True
               _Version        =   393217
               Icons           =   "ilSO"
               SmallIcons      =   "ilSO"
               ForeColor       =   -2147483640
               BackColor       =   -2147483643
               Appearance      =   1
               NumItems        =   4
               BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Text            =   "Section"
                  Object.Width           =   4762
               EndProperty
               BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   1
                  Text            =   "School Year"
                  Object.Width           =   2117
               EndProperty
               BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   2
                  Text            =   "Department"
                  Object.Width           =   3704
               EndProperty
               BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   3
                  Text            =   "Y.L."
                  Object.Width           =   1587
               EndProperty
            End
         End
         Begin HSES.b8SContainer bgTabCon 
            Height          =   4995
            Index           =   1
            Left            =   -75000
            TabIndex        =   17
            Top             =   300
            Width           =   9690
            _ExtentX        =   17092
            _ExtentY        =   8811
            BorderColor     =   12307149
            Begin MSComctlLib.ListView listSubjects 
               Height          =   4365
               Left            =   30
               TabIndex        =   22
               Top             =   330
               Width           =   8970
               _ExtentX        =   15822
               _ExtentY        =   7699
               View            =   3
               LabelEdit       =   1
               LabelWrap       =   -1  'True
               HideSelection   =   0   'False
               FullRowSelect   =   -1  'True
               GridLines       =   -1  'True
               _Version        =   393217
               Icons           =   "ilSubject"
               SmallIcons      =   "ilSubject"
               ForeColor       =   -2147483640
               BackColor       =   -2147483643
               Appearance      =   1
               NumItems        =   8
               BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Text            =   "School Year"
                  Object.Width           =   2117
               EndProperty
               BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   1
                  Text            =   "Subject"
                  Object.Width           =   3175
               EndProperty
               BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   2
                  Text            =   "Time In"
                  Object.Width           =   2117
               EndProperty
               BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   3
                  Text            =   "Time Out"
                  Object.Width           =   2117
               EndProperty
               BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   4
                  Text            =   "Days"
                  Object.Width           =   2117
               EndProperty
               BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   5
                  Text            =   "Department"
                  Object.Width           =   3175
               EndProperty
               BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   6
                  Text            =   "Y.L."
                  Object.Width           =   1587
               EndProperty
               BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   7
                  Text            =   "Section"
                  Object.Width           =   3175
               EndProperty
            End
            Begin MSComctlLib.ImageList ilSubject 
               Left            =   4905
               Top             =   -150
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
                     Picture         =   "frmTeacherRecord.frx":1A26
                     Key             =   ""
                  EndProperty
               EndProperty
            End
         End
      End
      Begin MSComctlLib.ImageList ilListIco 
         Left            =   5760
         Top             =   1785
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
               Picture         =   "frmTeacherRecord.frx":1FC0
               Key             =   "sy"
            EndProperty
         EndProperty
      End
      Begin VB.Image Image4 
         Height          =   105
         Left            =   -90
         Picture         =   "frmTeacherRecord.frx":255A
         Stretch         =   -1  'True
         Top             =   720
         Width           =   30000
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Teacher"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   180
         TabIndex        =   4
         Top             =   480
         Width           =   600
      End
      Begin VB.Image Image3 
         Height          =   345
         Left            =   0
         Picture         =   "frmTeacherRecord.frx":25F7
         Stretch         =   -1  'True
         Top             =   360
         Width           =   30000
      End
   End
   Begin HSES.b8Container b8cMain 
      Height          =   3420
      Left            =   390
      TabIndex        =   2
      Top             =   360
      Width           =   7035
      _ExtentX        =   12409
      _ExtentY        =   6033
      BorderColor     =   12632256
      BackColor       =   16777215
      InsideBorderColor=   14215660
      ShadowColor1    =   16777215
      ShadowColor2    =   16777215
   End
End
Attribute VB_Name = "frmTeacherRecord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim curTeacher As tTeacher
Dim ReadyToRefresh As Boolean


Public Function ShowForm(Optional sTeacherID As String = "")
    On Error Resume Next
    'show form
    mdiMain.MousePointer = vbHourglass
    ReadyToRefresh = False
    RefreshSYList
    
    Me.Show
    Me.SetFocus
    DoEvents
    
    curTeacher.TeacherID = sTeacherID
    ReadyToRefresh = True
    RefreshRecord
    
    'restore mouse pointer
    mdiMain.MousePointer = vbDefault
    

End Function

Private Sub RefreshRecord()
    'clear UI
    txtTeacherFullName.Text = ""
    lblTeacherName.Caption = ""
    lblLoginName.Caption = ""
    lblAddress.Caption = ""
    lblContactNumber.Caption = ""
    
    
    If GetTeacherByID(curTeacher.TeacherID, curTeacher) = Success Then
        
        'set info
        txtTeacherFullName.Text = curTeacher.FirstName & " " & curTeacher.MiddleName & " " & curTeacher.LastName
        lblTeacherName.Caption = curTeacher.FirstName & " " & curTeacher.MiddleName & " " & curTeacher.LastName
        lblLoginName.Caption = curTeacher.TeacherTitle
    
        lblAddress.Caption = curTeacher.Address
        lblContactNumber.Caption = curTeacher.ContactNumber
    
        RefreshAdvisoryList
        RefreshSubjects
        
        tabMAin.Enabled = True
    End If
End Sub


Private Sub RefreshSYList()
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    On Error GoTo ReleaseAndExit
    
    sSQL = "SELECT tblSchoolYear.SchoolYearTitle" & _
            " FROM tblSchoolYear" & _
            " ORDER BY tblSchoolYear.SchoolYearTitle"
     
    If ConnectRS(HSESDB, vRS, sSQL) = False Then
        'fatal
        'temp
        MsgBox "ERROR"
        GoTo ReleaseAndExit
    End If
    
    cmbSchoolYear.ComboItems.Clear
    cmbSchoolYear.ComboItems.Add , "all", "All", "sy"
    While vRS.EOF = False
        
        cmbSchoolYear.ComboItems.Add , ReadField(vRS.Fields("SchoolYearTitle")), ReadField(vRS.Fields("SchoolYearTitle")), "sy"
        vRS.MoveNext
    
    Wend
    
    cmbSchoolYear.ComboItems.Item(1).Selected = True
    
ReleaseAndExit:
    Set vRS = Nothing
End Sub


Private Function RefreshAdvisoryList()
        
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    'clear list
    listAdvisory.ListItems.Clear
    
    On Error GoTo ReleaseAndExit
    
    If cmbSchoolYear.SelectedItem.Key = "all" Then
    
        sSQL = "SELECT tblSectionOffering.SectionOfferingID, tblSection.SectionTitle, tblSectionOffering.SchoolYear, tblDepartment.DepartmentTitle, tblYearLevel.YearLevelTitle" & _
                " FROM tblYearLevel INNER JOIN (tblDepartment INNER JOIN (tblSection INNER JOIN (tblTeacher INNER JOIN tblSectionOffering ON tblTeacher.TeacherID = tblSectionOffering.TeacherID) ON tblSection.SectionID = tblSectionOffering.SectionID) ON tblDepartment.DepartmentID = tblSection.DepartmentID) ON tblYearLevel.YearLevelID = tblSection.YearLevelID" & _
                " GROUP BY tblSectionOffering.SectionOfferingID, tblSection.SectionTitle, tblSectionOffering.SchoolYear, tblDepartment.DepartmentTitle, tblYearLevel.YearLevelTitle, tblTeacher.TeacherID" & _
                " HAVING (((tblTeacher.TeacherID)='" & curTeacher.TeacherID & "'))" & _
                " ORDER BY tblSectionOffering.SchoolYear DESC"
    Else
        
        sSQL = "SELECT tblSectionOffering.SectionOfferingID, tblSection.SectionTitle, tblSectionOffering.SchoolYear, tblDepartment.DepartmentTitle, tblYearLevel.YearLevelTitle" & _
                " FROM tblYearLevel INNER JOIN (tblDepartment INNER JOIN (tblSection INNER JOIN (tblTeacher INNER JOIN tblSectionOffering ON tblTeacher.TeacherID = tblSectionOffering.TeacherID) ON tblSection.SectionID = tblSectionOffering.SectionID) ON tblDepartment.DepartmentID = tblSection.DepartmentID) ON tblYearLevel.YearLevelID = tblSection.YearLevelID" & _
                " GROUP BY tblSectionOffering.SectionOfferingID, tblSection.SectionTitle, tblSectionOffering.SchoolYear, tblDepartment.DepartmentTitle, tblYearLevel.YearLevelTitle, tblTeacher.TeacherID" & _
                " HAVING (((tblTeacher.TeacherID)='" & curTeacher.TeacherID & "') AND tblSectionOffering.SchoolYear='" & cmbSchoolYear.SelectedItem.Text & "')" & _
                " ORDER BY tblSectionOffering.SchoolYear DESC"

    End If
    
     
    If ConnectRS(HSESDB, vRS, sSQL) = False Then
        'fatal
        'temp
        MsgBox "Unable to conect Teacher's Advisory Recordset.", vbCritical
        GoTo ReleaseAndExit
    End If
    
    If AnyRecordExisted(vRS) = True Then
        FillRecordToList vRS, listAdvisory, KeySectionOffering, , 32767, , True
    End If

ReleaseAndExit:
    Set vRS = Nothing
End Function


Private Function RefreshSubjects()
        
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    'clear subject list
    listSubjects.ListItems.Clear
    
    On Error GoTo ReleaseAndExit
    
    If cmbSchoolYear.SelectedItem.Key = "all" Then
        
        sSQL = "SELECT tblSubjectOffering.SubjectOfferingID, tblSectionOffering.SchoolYear, tblSubject.SubjectTitle, tblSubjectOffering.SchedTimeStart, tblSubjectOffering.SchedTimeEnd, tblSubjectOffering.Days, tblDepartment.DepartmentTitle, tblYearLevel.YearLevelTitle, tblSection.SectionTitle" & _
                " FROM ((tblDepartment INNER JOIN (tblYearLevel INNER JOIN tblSection ON tblYearLevel.YearLevelID = tblSection.YearLevelID) ON tblDepartment.DepartmentID = tblSection.DepartmentID) INNER JOIN tblSectionOffering ON tblSection.SectionID = tblSectionOffering.SectionID) INNER JOIN (tblSubject INNER JOIN (tblTeacher INNER JOIN tblSubjectOffering ON tblTeacher.TeacherID = tblSubjectOffering.TeacherID) ON tblSubject.SubjectID = tblSubjectOffering.SubjectID) ON tblSectionOffering.SectionOfferingID = tblSubjectOffering.SectionOfferingID" & _
                " WHERE (((tblTeacher.TeacherID)='" & curTeacher.TeacherID & "'))" & _
                " ORDER BY tblSectionOffering.SchoolYear DESC , tblSubject.SubjectTitle DESC"
    Else
    
                sSQL = "SELECT tblSubjectOffering.SubjectOfferingID, tblSectionOffering.SchoolYear, tblSubject.SubjectTitle, tblSubjectOffering.SchedTimeStart, tblSubjectOffering.SchedTimeEnd, tblSubjectOffering.Days, tblDepartment.DepartmentTitle, tblYearLevel.YearLevelTitle, tblSection.SectionTitle" & _
                " FROM ((tblDepartment INNER JOIN (tblYearLevel INNER JOIN tblSection ON tblYearLevel.YearLevelID = tblSection.YearLevelID) ON tblDepartment.DepartmentID = tblSection.DepartmentID) INNER JOIN tblSectionOffering ON tblSection.SectionID = tblSectionOffering.SectionID) INNER JOIN (tblSubject INNER JOIN (tblTeacher INNER JOIN tblSubjectOffering ON tblTeacher.TeacherID = tblSubjectOffering.TeacherID) ON tblSubject.SubjectID = tblSubjectOffering.SubjectID) ON tblSectionOffering.SectionOfferingID = tblSubjectOffering.SectionOfferingID" & _
                " WHERE (((tblTeacher.TeacherID)='" & curTeacher.TeacherID & "') AND tblSectionOffering.SchoolYear='" & cmbSchoolYear.SelectedItem.Text & "')" & _
                " ORDER BY tblSectionOffering.SchoolYear DESC , tblSubject.SubjectTitle DESC"

    End If

     
    If ConnectRS(HSESDB, vRS, sSQL) = False Then
        'fatal
        'temp
        MsgBox "Unable to conect Teacher Recordset.", vbCritical
        GoTo ReleaseAndExit
    End If
    
    If AnyRecordExisted(vRS) = True Then
        FillRecordToList vRS, listSubjects, KeySubjectOffering, , 32767, , True
    End If

ReleaseAndExit:
    Set vRS = Nothing
End Function

Public Function Form_CanRefresh() As Boolean
    Form_CanRefresh = True
End Function

Public Function Form_Refresh()
    RefreshRecord
End Function


Public Function Form_Can_Reload() As Boolean
    Form_Can_Reload = True
End Function

Public Function Form_Reload()
   RefreshRecord
End Function

Public Function Form_CanPrint() As Boolean
    'Form_CanPrint = True
End Function

Private Sub cmbSchoolYear_Click()
   If ReadyToRefresh = True Then
        RefreshRecord
    End If
End Sub

Private Sub cmdGetTeacher_Click()

    Dim sTeacherID As String
    
    sTeacherID = PickTeacher.GetTeacherID()
    
    If Len(sTeacherID) > 0 Then
        curTeacher.TeacherID = sTeacherID
        RefreshRecord
    Else
    End If
   
End Sub

Private Sub Form_Activate()
    mdiMain.RegMDIChild Me
End Sub

Private Sub Form_Resize()
    ReArrangeControls
End Sub
Private Sub ReArrangeControls()
On Error Resume Next
    
    
    Me.ScaleMode = vbPixels
    b8cMain.Move Form_LeftMargin - 3, Form_TopMargin - 3, Me.ScaleWidth - (Form_LeftMargin - 3) * 2, Me.ScaleHeight - (Form_TopMargin - 3) * 2
    
    bgMain.Move Form_LeftMargin, Form_TopMargin, Me.ScaleWidth - Form_LeftMargin * 2, Me.ScaleHeight - Form_TopMargin * 2
    
    b8Title.Move 0, 0, bgMain.Width
    
    bgDetail.Move bgDetail.Left, bgDetail.Top, bgMain.Width - bgDetail.Left * 2
    
    tabMAin.Move 0, tabMAin.Top, bgMain.Width, bgMain.Height - tabMAin.Top
    ReArrangeTab
    
End Sub

Private Sub ReArrangeTab()

    On Error Resume Next
    
    bgTabCon(tabMAin.Tab).Move 0, bgTabCon(tabMAin.Tab).Top, Screen.TwipsPerPixelX * tabMAin.Width, Screen.TwipsPerPixelY * tabMAin.Height - bgTabCon(tabMAin.Tab).Top

    
    Select Case tabMAin.Tab
        Case 0 'advisory
            listAdvisory.Move 30, listAdvisory.Top, bgTabCon(tabMAin.Tab).Width - (60), bgTabCon(tabMAin.Tab).Height - listAdvisory.Top - 30
        Case 1 ' subjects
            listSubjects.Move 30, listSubjects.Top, bgTabCon(tabMAin.Tab).Width - (60), bgTabCon(tabMAin.Tab).Height - listSubjects.Top - 30

    End Select
    
End Sub

Private Sub tabMAin_Click(PreviousTab As Integer)
    bgTabCon(tabMAin.Tab).Visible = True
    ReArrangeControls
End Sub

