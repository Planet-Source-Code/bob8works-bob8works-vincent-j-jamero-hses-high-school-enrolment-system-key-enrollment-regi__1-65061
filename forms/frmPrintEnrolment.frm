VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmPrintEnrolment 
   BackColor       =   &H00D8E9EC&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Print Student/School Copy"
   ClientHeight    =   2025
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4680
   Icon            =   "frmPrintEnrolment.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   135
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar progPrint 
      Height          =   345
      Left            =   90
      TabIndex        =   4
      Top             =   750
      Width           =   4485
      _ExtentX        =   7911
      _ExtentY        =   609
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin lvButton.lvButtons_H cmdAbort 
      Cancel          =   -1  'True
      Height          =   360
      Left            =   3090
      TabIndex        =   0
      Top             =   1500
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   635
      Caption         =   "&Abort"
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
      TabIndex        =   1
      Top             =   1290
      Width           =   5565
      _ExtentX        =   9816
      _ExtentY        =   106
   End
   Begin HSES.b8Line b8Line1 
      Height          =   60
      Left            =   0
      TabIndex        =   2
      Top             =   510
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   106
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Printing.."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H002F556A&
      Height          =   240
      Left            =   630
      TabIndex        =   3
      Top             =   120
      Width           =   870
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   60
      Picture         =   "frmPrintEnrolment.frx":08CA
      Top             =   30
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   525
      Left            =   30
      Picture         =   "frmPrintEnrolment.frx":1194
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5535
   End
End
Attribute VB_Name = "frmPrintEnrolment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Function ShowForm(sSchoolYear As String)
    
    Me.Show
    DoEvents
    
    ShowStudentCopyBySchoolYear sSchoolYear
    
    Unload Me
End Function


Public Function ShowStudentCopyBySchoolYear(Optional sSchoolYear As String = "")
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
   
    
    If sSchoolYear = "" Then
        sSchoolYear = PickSchoolYear.GetItem
    End If
    
    If sSchoolYear = "" Then
        GoTo ReleaseAndExit
    End If
    
    sSQL = "SELECT tblEnrolment.EnrolmentID" & _
            " From tblEnrolment" & _
            " WHERE (((tblEnrolment.SchoolYear)='" & sSchoolYear & "'));"


    If ConnectRS(HSESDB, vRS, sSQL) = False Then
        MsgBox "error"
        GoTo ReleaseAndExit
    End If
    
    If AnyRecordExisted(vRS) = False Then
        GoTo ReleaseAndExit
    End If
    
    progPrint.Max = getRecordCount(vRS) + 2
    progPrint.Value = progPrint.Min
    
    vRS.MoveFirst
    
    While vRS.EOF = False
        
        progPrint.Value = progPrint.Value + 1
        
        
        If MsgBox("Printing Student Copy #" & progPrint.Value & " / " & progPrint.Max - 2 & vbNewLine & vbNewLine & _
                "Press Cancel to Abort.", vbExclamation + vbOKCancel) = vbCancel Then
                GoTo ReleaseAndExit
        End If
        
        ShowStudentCopyByEnrolment (ReadField(vRS.Fields("EnrolmentID")))

        vRS.MoveNext
        
    Wend
    
ReleaseAndExit:
    Set vRS = Nothing
    
End Function

Private Sub Form_Deactivate()
    On Error Resume Next
    
    Me.Show
    Me.SetFocus
    
End Sub

Public Function ShowStudentCopyByEnrolment(Optional sEnrolmentID As String = "")

    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    
    If sEnrolmentID = "" Then
        GoTo ReleaseAndExit
    End If
    
    sSQL = "SELECT tblStudent.StudentID, [tblStudent]![LastName] & ', ' & [tblStudent]![FirstName] & ' ' & [tblStudent]![MiddleName] AS StudentFullName, tblStudent.Gender, tblStudent.CityAddress, tblStudent.HomeAddress, tblSubject.SubjectTitle, [tblSubjectOffering]![SchedTimeStart] & ' - ' & [tblSubjectOffering]![SchedTimeEnd] AS TimeSchedule, tblSubjectOffering.Days, Left([tblteacher]![FirstName],1) & '. ' & [tblteacher]![LastName] AS TeacherFullName, tblEnrolment.EnrolmentID" & _
            " FROM tblStudent INNER JOIN (tblTeacher INNER JOIN (tblSubject INNER JOIN ((tblSectionOffering INNER JOIN tblEnrolment ON tblSectionOffering.SectionOfferingID = tblEnrolment.SectionOfferingID) INNER JOIN tblSubjectOffering ON tblSectionOffering.SectionOfferingID = tblSubjectOffering.SectionOfferingID) ON tblSubject.SubjectID = tblSubjectOffering.SubjectID) ON tblTeacher.TeacherID = tblSubjectOffering.TeacherID) ON tblStudent.StudentID = tblEnrolment.StudentID" & _
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
    
    drEnrolmentDetail.Show vbModal

ReleaseAndExit:
    Set vRS = Nothing
End Function


Public Function ShowPrint(ByRef eRS As ADODB.Recordset)
    
       
    If AnyRecordExisted(eRS) = False Then
        GoTo ReleaseAndExit
    End If
    
    progPrint.Max = getRecordCount(eRS) + 2
    progPrint.Value = progPrint.Min
    
    eRS.MoveFirst
    
    While eRS.EOF = False
        
        
        progPrint.Value = progPrint.Value + 1
        
        If MsgBox("Printing Student Copy #" & progPrint.Value & vbNewLine & vbNewLine & _
                "Press Cancel to Abort.", vbExclamation + vbOKCancel) = vbCancel Then
                GoTo ReleaseAndExit
        End If
        
        ShowStudentCopyByEnrolment (ReadField(eRS.Fields("EnrolmentID")))
        
        eRS.MoveNext
        
    Wend
    
ReleaseAndExit:
    
    Unload Me
End Function
