VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmReports 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Reports"
   ClientHeight    =   5625
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6210
   ControlBox      =   0   'False
   Icon            =   "frmReports.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5625
   ScaleWidth      =   6210
   Begin VB.PictureBox bgMain 
      Appearance      =   0  'Flat
      BackColor       =   &H00D8E9EC&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4995
      Left            =   540
      ScaleHeight     =   333
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   471
      TabIndex        =   1
      Top             =   750
      Width           =   7065
      Begin MSComctlLib.ImageList icoHeader 
         Left            =   5805
         Top             =   3210
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
               Picture         =   "frmReports.frx":08CA
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmReports.frx":0E64
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin HSES.b8SContainer pbBGButton 
         Height          =   525
         Left            =   0
         TabIndex        =   2
         Top             =   345
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   926
         BorderColor     =   14215660
      End
      Begin MSComctlLib.ListView listRecord 
         Height          =   3480
         Left            =   0
         TabIndex        =   3
         Top             =   1110
         Width           =   5505
         _ExtentX        =   9710
         _ExtentY        =   6138
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         HideColumnHeaders=   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         Icons           =   "imgListIco"
         SmallIcons      =   "imgListIco"
         ForeColor       =   12582912
         BackColor       =   16777215
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
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Report Name"
            Object.Width           =   6350
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Description"
            Object.Width           =   15875
         EndProperty
      End
      Begin HSES.b8ChildTitleBar b8Title 
         Height          =   345
         Left            =   0
         TabIndex        =   4
         Top             =   0
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   609
         BackColor       =   16777215
         Caption         =   "Select Task"
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
         ForeColor       =   12582912
         GradTheme       =   2
      End
   End
   Begin HSES.b8Container b8cMain 
      Height          =   5940
      Left            =   300
      TabIndex        =   0
      Top             =   780
      Width           =   7035
      _ExtentX        =   12409
      _ExtentY        =   10478
      BorderColor     =   12632256
      BackColor       =   16185592
      InsideBorderColor=   14215660
      ShadowColor1    =   16777215
      ShadowColor2    =   16777215
   End
   Begin MSComctlLib.ImageList imgListIco 
      Left            =   0
      Top             =   0
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
            Picture         =   "frmReports.frx":13FE
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmReports"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Public Function ShowForm()
    Me.Show
End Function


Private Sub Form_Activate()
    mdiMain.RegMDIChild Me
    
    RefreshReportList
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
    pbBGButton.Move 0, b8Title.Top + b8Title.Height, bgMain.Width
    listRecord.Move listRecord.Left, pbBGButton.Top + pbBGButton.Height, bgMain.Width - (listRecord.Left * 2)
    listRecord.Height = bgMain.Height - (listRecord.Top)
   
End Sub

Private Sub RefreshReportList()
    
    listRecord.ListItems.Clear
    
    listRecord.ListItems.Add , "rptSchoolYear", "School Year List.", 1, 1
    'listRecord.ListItems(listRecord.ListItems.Count).SubItems(1) = "Print School Year List with Related Record Count."
    listRecord.ListItems.Add , , ""
    
    listRecord.ListItems.Add , "rptDepartment", "Department List.", 1, 1
    'listRecord.ListItems(listRecord.ListItems.Count).SubItems(1) = "Print Department List."
    listRecord.ListItems.Add , , ""
    
    listRecord.ListItems.Add , "rptSection", "Section List.", 1, 1
    'listRecord.ListItems(listRecord.ListItems.Count).SubItems(1) = "Print Section List."
    listRecord.ListItems.Add , , ""
    
    listRecord.ListItems.Add , "rptSectionOffering", "Section Offering List.", 1, 1
    'listRecord.ListItems(listRecord.ListItems.Count).SubItems(1) = "Print Section Offering List."
    listRecord.ListItems.Add , , ""
    
    listRecord.ListItems.Add , "rptSubject", "Subject List.", 1, 1
    listRecord.ListItems.Add , , ""
    
    listRecord.ListItems.Add , "rptStudent", "Student And Enrolment.", 1, 1
    listRecord.ListItems.Add , , ""
    
    listRecord.ListItems.Add , "rptTeacher", "Teacher Accounts.", 1, 1
    listRecord.ListItems.Add , , ""

End Sub



Private Sub listRecord_DblClick()

    
    Call Form_Print
End Sub

Private Function SetSectionPrintRS()
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    sSQL = "SELECT tblSection.SectionID AS lvKey, [tblYearLevel]![YearLevelTitle] & ' - ' & [tblSection]![SectionTitle] AS SectionFullTitle, tblDepartment.DepartmentTitle, tblSection.CreationDate,tblSection.CreatedBy" & _
            " FROM tblDepartment INNER JOIN (tblYearLevel INNER JOIN tblSection ON tblYearLevel.YearLevelID = tblSection.YearLevelID) ON tblDepartment.DepartmentID = tblSection.DepartmentID" & _
            " ORDER BY [tblYearLevel]![YearLevelTitle] & ' - ' & [tblSection]![SectionTitle]"

    If ConnectRS(HSESDB, vRS, sSQL) = False Then
        MsgBox "error"
        GoTo ReleaseAndExit
    End If

    frmPrintSections.ShowForm vRS

ReleaseAndExit:
    Set vRS = Nothing
End Function


Public Function Form_CanPrint() As Boolean
    On Error Resume Next
    'default
    Form_CanPrint = False
    If Len(listRecord.SelectedItem.Key) > 0 Then
        Form_CanPrint = True
    Else
        Form_CanPrint = False
    End If
End Function

Public Function Form_Print()


    Select Case listRecord.SelectedItem.Key
    
        Case "rptSchoolYear"
            frmPrintSchoolYear.ShowForm
        
        Case "rptDepartment"
            frmPrintDepartment.ShowForm
        
        Case "rptSection"
            SetSectionPrintRS
            
        Case "rptSectionOffering"
            frmPrintSectionOffering.ShowForm
            
        Case "rptSubject"
            frmPrintSubject.ShowForm
            
        Case "rptStudent"
            frmPrintStudent.ShowForm
            
        Case "rptTeacher"
            frmPrintTeachers.ShowForm
    End Select
    
End Function

Private Sub listRecord_ItemClick(ByVal Item As MSComctlLib.ListItem)
    
    Static lastIndex As Integer
    
    On Error Resume Next
     
    If listRecord.SelectedItem.Key = "" Then
        If lastIndex > listRecord.SelectedItem.Index Then
             listRecord.ListItems(listRecord.SelectedItem.Index - 1).Selected = True
        ElseIf lastIndex < listRecord.SelectedItem.Index Then
            listRecord.ListItems(listRecord.SelectedItem.Index + 1).Selected = True
        End If
    End If
    
    lastIndex = listRecord.SelectedItem.Index
End Sub
