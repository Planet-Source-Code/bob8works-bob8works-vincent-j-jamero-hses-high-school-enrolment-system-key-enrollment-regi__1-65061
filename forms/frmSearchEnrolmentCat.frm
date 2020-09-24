VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "LVbuttons.ocx"
Begin VB.Form frmGetERS 
   BackColor       =   &H00D8E9EC&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Search By Category"
   ClientHeight    =   5070
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6360
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   338
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   424
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin HSES.b8Container b8Container1 
      Height          =   4245
      Left            =   90
      TabIndex        =   2
      Top             =   150
      Width           =   6165
      _ExtentX        =   10874
      _ExtentY        =   7488
      BackColor       =   16185592
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H00F6F8F8&
         ForeColor       =   &H80000008&
         Height          =   1635
         Left            =   210
         ScaleHeight     =   107
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   381
         TabIndex        =   22
         Top             =   2430
         Width           =   5745
         Begin VB.CheckBox chkByDate 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Check3"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   60
            TabIndex        =   26
            Top             =   30
            Width           =   195
         End
         Begin VB.OptionButton optToday 
            BackColor       =   &H00F6F8F8&
            Caption         =   "Today"
            Enabled         =   0   'False
            Height          =   315
            Left            =   270
            TabIndex        =   25
            Top             =   330
            Value           =   -1  'True
            Width           =   885
         End
         Begin VB.OptionButton optSpecifyDate 
            BackColor       =   &H00F6F8F8&
            Caption         =   "Specify Date"
            Enabled         =   0   'False
            Height          =   315
            Left            =   1500
            TabIndex        =   24
            Top             =   330
            Width           =   1335
         End
         Begin VB.OptionButton optDateWithRange 
            BackColor       =   &H00F6F8F8&
            Caption         =   "Date with range"
            Enabled         =   0   'False
            Height          =   315
            Left            =   3480
            TabIndex        =   23
            Top             =   330
            Width           =   1845
         End
         Begin MSComCtl2.DTPicker dtpSpecifyDate 
            Height          =   315
            Left            =   1560
            TabIndex        =   27
            Top             =   720
            Width           =   1425
            _ExtentX        =   2514
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Format          =   56098817
            CurrentDate     =   36892
         End
         Begin MSComCtl2.DTPicker dtpFrom 
            Height          =   315
            Left            =   4020
            TabIndex        =   28
            Top             =   720
            Width           =   1425
            _ExtentX        =   2514
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Format          =   56098817
            CurrentDate     =   36892
         End
         Begin MSComCtl2.DTPicker dtpTo 
            Height          =   315
            Left            =   4020
            TabIndex        =   29
            Top             =   1140
            Width           =   1425
            _ExtentX        =   2514
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Format          =   56098817
            CurrentDate     =   36892
         End
         Begin VB.Label Label11 
            BackColor       =   &H00D8E9EC&
            Caption         =   "       By Date"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   255
            Left            =   0
            TabIndex        =   32
            Top             =   0
            Width           =   5700
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "From"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   195
            Left            =   3570
            TabIndex        =   31
            Top             =   720
            Width           =   360
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "To"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   195
            Left            =   3570
            TabIndex        =   30
            Top             =   1140
            Width           =   180
         End
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H00F6F8F8&
         ForeColor       =   &H80000008&
         Height          =   975
         Left            =   210
         ScaleHeight     =   63
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   381
         TabIndex        =   12
         Top             =   180
         Width           =   5745
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Section"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   195
            Left            =   3990
            TabIndex        =   21
            Top             =   660
            Width           =   525
         End
         Begin MSForms.ComboBox cmbSectionTitle 
            Height          =   345
            Left            =   3990
            TabIndex        =   20
            Top             =   330
            Width           =   1515
            VariousPropertyBits=   746604571
            DisplayStyle    =   3
            Size            =   "2672;609"
            MatchEntry      =   1
            ShowDropButtonWhen=   1
            DropButtonStyle =   2
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Yr."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   195
            Left            =   3360
            TabIndex        =   19
            Top             =   660
            Width           =   210
         End
         Begin MSForms.ComboBox cmbYearLevelTitle 
            Height          =   345
            Left            =   3360
            TabIndex        =   18
            Top             =   330
            Width           =   615
            VariousPropertyBits=   746604571
            DisplayStyle    =   3
            Size            =   "1085;609"
            MatchEntry      =   1
            ShowDropButtonWhen=   1
            DropButtonStyle =   2
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Department"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   195
            Left            =   1680
            TabIndex        =   17
            Top             =   660
            Width           =   855
         End
         Begin MSForms.ComboBox cmbDepartmentTitle 
            Height          =   345
            Left            =   1680
            TabIndex        =   16
            Top             =   330
            Width           =   1665
            VariousPropertyBits=   746604571
            DisplayStyle    =   3
            Size            =   "2937;609"
            MatchEntry      =   1
            ShowDropButtonWhen=   1
            DropButtonStyle =   2
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "School Year"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   195
            Left            =   150
            TabIndex        =   15
            Top             =   660
            Width           =   840
         End
         Begin MSForms.ComboBox cmbSchoolYearTitle 
            Height          =   345
            Left            =   150
            TabIndex        =   14
            Top             =   330
            Width           =   1515
            VariousPropertyBits=   746604571
            DisplayStyle    =   3
            Size            =   "2672;609"
            MatchEntry      =   1
            ShowDropButtonWhen=   1
            DropButtonStyle =   2
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label Label6 
            BackColor       =   &H00D8E9EC&
            Caption         =   "   By Directory"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   225
            Left            =   0
            TabIndex        =   13
            Top             =   0
            Width           =   5730
         End
      End
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BackColor       =   &H00F6F8F8&
         Height          =   945
         Left            =   2190
         ScaleHeight     =   915
         ScaleWidth      =   3735
         TabIndex        =   9
         Top             =   1320
         Width           =   3765
         Begin lvButton.lvButtons_H cmdGetStudentID 
            Height          =   345
            Left            =   3120
            TabIndex        =   33
            Top             =   390
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   609
            Caption         =   "..."
            CapAlign        =   2
            BackStyle       =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
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
         Begin MSForms.ComboBox txtStudentID 
            Height          =   345
            Left            =   330
            TabIndex        =   11
            Top             =   390
            Width           =   2775
            VariousPropertyBits=   746604571
            DisplayStyle    =   3
            Size            =   "4895;609"
            MatchEntry      =   1
            ShowDropButtonWhen=   1
            DropButtonStyle =   2
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label Label5 
            BackColor       =   &H00D8E9EC&
            Caption         =   "   By Student"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   225
            Left            =   0
            TabIndex        =   10
            Top             =   0
            Width           =   5730
         End
      End
      Begin VB.PictureBox Picture4 
         Appearance      =   0  'Flat
         BackColor       =   &H00F6F8F8&
         ForeColor       =   &H80000008&
         Height          =   945
         Left            =   210
         ScaleHeight     =   915
         ScaleWidth      =   1755
         TabIndex        =   3
         Top             =   1320
         Width           =   1785
         Begin VB.CheckBox Check2 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   150
            TabIndex        =   5
            Top             =   540
            Value           =   1  'Checked
            Width           =   195
         End
         Begin VB.CheckBox Check1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   150
            TabIndex        =   4
            Top             =   300
            Value           =   1  'Checked
            Width           =   195
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Male"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   195
            Left            =   390
            TabIndex        =   8
            Top             =   300
            Width           =   330
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Female"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   195
            Left            =   390
            TabIndex        =   7
            Top             =   540
            Width           =   510
         End
         Begin VB.Label Label7 
            BackColor       =   &H00D8E9EC&
            Caption         =   "   By Gender"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   225
            Left            =   0
            TabIndex        =   6
            Top             =   0
            Width           =   1920
         End
      End
   End
   Begin lvButton.lvButtons_H cmdOK 
      Default         =   -1  'True
      Height          =   405
      Left            =   4800
      TabIndex        =   0
      Top             =   4530
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   714
      Caption         =   "&OK"
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
      Height          =   405
      Left            =   3210
      TabIndex        =   1
      Top             =   4530
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   714
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
   Begin VB.Image Image2 
      Height          =   1155
      Left            =   0
      Picture         =   "frmSearchEnrolmentCat.frx":0000
      Stretch         =   -1  'True
      Top             =   3960
      Width           =   6495
   End
   Begin VB.Image Image4 
      Height          =   135
      Left            =   0
      Picture         =   "frmSearchEnrolmentCat.frx":009D
      Stretch         =   -1  'True
      Top             =   0
      Width           =   6495
   End
End
Attribute VB_Name = "frmGetERS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const sALL = "ALL"

Dim vERS As New ADODB.Recordset
Dim tmpGetEnrolmentRS As Integer

Public Function GetEnrolmentRS(ByRef vRS As ADODB.Recordset) As Integer

    
    'show form
    Me.Show vbModal
    
    If tmpGetEnrolmentRS = 1 Then
        GetEnrolmentRS = tmpGetEnrolmentRS
        Set vRS = vERS
    End If
    
    Set vERS = Nothing
End Function

Public Function ReturnCancel()
    tmpGetEnrolmentRS = 0
    Set vERS = Nothing
    Unload Me
End Function

Private Sub chkByDate_Click()
    If chkByDate = vbChecked Then
        optToday.Enabled = True
        optSpecifyDate.Enabled = True
        optDateWithRange.Enabled = True
    Else
        optToday.Enabled = False
        optSpecifyDate.Enabled = False
        optDateWithRange.Enabled = False
    End If
End Sub

Private Sub cmbSectionTitle_GotFocus()
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    Dim vSection As tSection

    'clear Sections
    cmbSectionTitle.Clear
    cmbSectionTitle.AddItem sALL
    
    'yearlevel check cound, if 0 then exit
    If Len(Trim(cmbYearLevelTitle.Text)) < 1 Or cmbYearLevelTitle.ListCount < 1 Then Exit Sub
    

    If cmbDepartmentTitle.Text = sALL And cmbYearLevelTitle.Text = sALL Then
        'include all
        
        Dim tmpRSSection As New ADODB.Recordset
        
        If CreateDefaultRSSection(tmpRSSection) = 1 Then
            If RSMoveFirst(tmpRSSection) Then
                While GetSectionMoveNext(tmpRSSection, vSection) = Success
                    cmbSectionTitle.AddItem vSection.SectionTitle
                Wend
            End If
        End If
        Set tmpRSSection = Nothing
    Else
    
        If cmbDepartmentTitle.Text <> sALL And cmbYearLevelTitle.Text = sALL Then
            'search by  department only
            sSQL = "SELECT tblSection.SectionTitle, tblDepartment.DepartmentTitle, tblYearLevel.YearLevelTitle FROM tblYearLevel INNER JOIN (tblDepartment INNER JOIN tblSection ON tblDepartment.DepartmentID = tblSection.DepartmentID) ON tblYearLevel.YearLevelID = tblSection.YearLevelID WHERE (((tblDepartment.DepartmentTitle)='" & Trim(cmbDepartmentTitle.Text) & "'));"
            
        ElseIf cmbDepartmentTitle.Text = sALL And cmbYearLevelTitle.Text <> sALL Then
            'search by  yearlevel only
            sSQL = "SELECT tblSection.SectionTitle, tblDepartment.DepartmentTitle, tblYearLevel.YearLevelTitle FROM tblYearLevel INNER JOIN (tblDepartment INNER JOIN tblSection ON tblDepartment.DepartmentID = tblSection.DepartmentID) ON tblYearLevel.YearLevelID = tblSection.YearLevelID WHERE (((tblYearLevel.YearLevelTitle)='" & Trim(cmbYearLevelTitle.Text) & "'));"
    
        Else
            'search by department and yearlevel
            sSQL = "SELECT tblSection.SectionTitle, tblDepartment.DepartmentTitle, tblYearLevel.YearLevelTitle FROM tblYearLevel INNER JOIN (tblDepartment INNER JOIN tblSection ON tblDepartment.DepartmentID = tblSection.DepartmentID) ON tblYearLevel.YearLevelID = tblSection.YearLevelID WHERE (((tblDepartment.DepartmentTitle)='" & Trim(cmbDepartmentTitle.Text) & "') AND ((tblYearLevel.YearLevelTitle)='" & Trim(cmbYearLevelTitle.Text) & "'));"
        End If
        
        
        If ConnectRS(DB, vRS, sSQL) Then
            If AnyRecordExisted(vRS) Then
                vRS.MoveFirst
                
                While vRS.EOF = False
                    cmbSectionTitle.AddItem ReadField(vRS.Fields("SectionTitle"))
                    vRS.MoveNext
                Wend
                
            End If
        End If
        
    End If
    
    
    
    Set vRS = Nothing
End Sub

Private Sub cmdCancel_Click()
    ReturnCancel
End Sub

Private Sub cmdGetStudentID_Click()
    txtStudentID = PickStudent.GetStudentID
End Sub

Private Sub cmdOK_Click()
    Dim sSQL As String
    Dim WHERE_Clause_Added As Boolean
    
    
    
    
    sSQL = "SELECT tblEnrolment.EnrolmentID, [tblStudent]![LastName]+', '+[tblStudent]![FirstName]+' '+[tblStudent]![MiddleName] AS StudentFullName, tblStudent.StudentID, tblEnrolment.SchoolYear, tblDepartment.DepartmentTitle, tblSection.SectionTitle, tblYearLevel.YearLevelTitle, tblStudent.Gender, tblEnrolment.CreationDate, tblEnrolment.CreatedUserName, tblEnrolment.ModifiedUserName, tblEnrolment.LastModified FROM tblYearLevel INNER JOIN (tblStudent INNER JOIN ((tblDepartment INNER JOIN tblSection ON tblDepartment.DepartmentID = tblSection.DepartmentID) INNER JOIN tblEnrolment ON tblSection.SectionID = tblEnrolment.SectionID) ON tblStudent.StudentID = tblEnrolment.StudentID) ON tblYearLevel.YearLevelID = tblSection.YearLevelID"
    
    
    If Len(cmbSchoolYearTitle.Text) > 0 And cmbSchoolYearTitle.Text <> sALL Then
        
        WHERE_Clause_Added = True
        sSQL = sSQL & " WHERE (((tblEnrolment.SchoolYear)='" & cmbSchoolYearTitle.Text & "')"
        
    End If
    
    
    
    If Len(cmbDepartmentTitle.Text) > 0 And cmbDepartmentTitle.Text <> sALL Then
            
        If WHERE_Clause_Added <> True Then
            sSQL = sSQL & " WHERE ("
            WHERE_Clause_Added = True
        Else
            sSQL = sSQL & " AND "
        End If

        sSQL = sSQL & " ((tblDepartment.DepartmentTitle)='" & cmbDepartmentTitle.Text & "')"
        
    End If
    
    
    
    If Len(cmbYearLevelTitle.Text) > 0 And cmbYearLevelTitle.Text <> sALL Then
        
        If WHERE_Clause_Added <> True Then
            sSQL = sSQL & " WHERE ("
            WHERE_Clause_Added = True
        Else
            sSQL = sSQL & " AND "
        End If
        
        sSQL = sSQL & " ((tblYearLevel.YearLevelTitle)='" & cmbYearLevelTitle.Text & "')"
        
    End If
    
    
    
    
    If Len(cmbSectionTitle.Text) > 0 And cmbSectionTitle.Text <> sALL Then
            
        If WHERE_Clause_Added <> True Then
            sSQL = sSQL & " WHERE ("
            WHERE_Clause_Added = True
        Else
            sSQL = sSQL & " AND "
        End If
        
        sSQL = sSQL & " ((tblSection.SectionTitle)='" & cmbSectionTitle.Text & "')"
        
    End If
    
    
    
    If Len(txtStudentID.Text) > 0 Then
                        
        If WHERE_Clause_Added <> True Then
            sSQL = sSQL & " WHERE ("
            WHERE_Clause_Added = True
        Else
            sSQL = sSQL & " AND "
        End If
        
        sSQL = sSQL & " ((tblStudent.StudentID)='" & txtStudentID.Text & "')"
        
    End If
    

    'by date
    If chkByDate.Value = vbChecked Then
    
        If optToday.Value = True Then
            
            If WHERE_Clause_Added <> True Then
                sSQL = sSQL & " WHERE ("
                WHERE_Clause_Added = True
            Else
                sSQL = sSQL & " AND "
            End If

            sSQL = sSQL & " ((tblEnrolment.CreationDate)=#" & FormatDateTime(Now, vbShortDate) & "#)"
        
        
        
        
        ElseIf optSpecifyDate.Value = True Then
            
            If WHERE_Clause_Added <> True Then
                sSQL = sSQL & " WHERE ("
                WHERE_Clause_Added = True
            Else
                sSQL = sSQL & " AND "
            End If

            sSQL = sSQL & " ((tblEnrolment.CreationDate)=#" & FormatDateTime(dtpSpecifyDate.Value, vbShortDate) & "#)"
        
        
        
        ElseIf optDateWithRange.Value = True Then
            
            If WHERE_Clause_Added <> True Then
                sSQL = sSQL & " WHERE ("
                WHERE_Clause_Added = True
            Else
                sSQL = sSQL & " AND "
            End If

            sSQL = sSQL & " ((tblEnrolment.CreationDate) Between #" & FormatDateTime(dtpFrom.Value, vbShortDate) & "# And #" & FormatDateTime(dtpTo.Value, vbShortDate) & "#)"
        End If
    End If



    If WHERE_Clause_Added = True Then
        sSQL = sSQL & ");"
    End If
    
    MsgBox sSQL
    
    If ConnectRS(DB, vERS, sSQL) Then
        tmpGetEnrolmentRS = 1
    End If

    Unload Me


End Sub





Private Sub Form_Load()
    FillCategory
End Sub

Private Sub FillCategory()
    
    Dim vSchoolYear As tSchoolYear
    Dim vDepartment As tDepartment
    Dim vYearLevel As tYearLevel
    
    Dim vRSSchoolYear As New ADODB.Recordset
    
    'add school year entries
    cmbSchoolYearTitle.Clear
    cmbSchoolYearTitle.AddItem sALL
    If CreateDefaultRSSchoolYear(vRSSchoolYear) = 1 Then
        If RSMoveFirst(vRSSchoolYear) Then
            While GetSchoolYearMoveNext(vRSSchoolYear, vSchoolYear) = Success
                cmbSchoolYearTitle.AddItem vSchoolYear.SchoolYearTitle
            Wend
        End If
    End If
    
    'add departments
    'cmbDepartmentTitle.Clear
    'cmbDepartmentTitle.AddItem sALL
    'If RSMoveFirst Then
    '    While GetDepartmentMoveNext(vDepartment) = Success
    '        cmbDepartmentTitle.AddItem vDepartment.DepartmentTitle
    '    Wend
    'End If
    
    
    'add YearLevels
    cmbYearLevelTitle.Clear
    cmbYearLevelTitle.AddItem sALL
    If YearLevelMovefirst Then
        While GetYearLevelMoveNext(vYearLevel)
            cmbYearLevelTitle.AddItem vYearLevel.YearLevelTitle
        Wend
    End If
    

    'release
    
    Set vRSSchoolYear = Nothing
End Sub





Private Sub dtpFrom_Change()
    If dtpFrom.Value > dtpTo.Value Then
        dtpFrom.Value = dtpTo.Value
    End If
End Sub

Private Sub dtpTo_Change()
    If dtpFrom.Value > dtpTo.Value Then
        dtpTo.Value = dtpFrom.Value
    End If
End Sub


Private Sub optDateWithRange_Click()
    Call optByDate_Change
End Sub


Private Sub optSpecifyDate_Click()
    Call optByDate_Change
End Sub

Private Sub optToday_Click()
    Call optByDate_Change
End Sub

Private Sub optByDate_Change()
    dtpSpecifyDate.Enabled = optSpecifyDate.Value
    dtpFrom.Enabled = optDateWithRange.Value
    dtpTo.Enabled = optDateWithRange.Value
End Sub



