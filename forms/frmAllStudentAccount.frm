VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmAllStudent 
   BackColor       =   &H00F6F8F8&
   Caption         =   "Students"
   ClientHeight    =   5670
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   10410
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAllStudentAccount.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   378
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   694
   Begin VB.PictureBox bgMain 
      Appearance      =   0  'Flat
      BackColor       =   &H00D8E9EC&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4995
      Left            =   420
      ScaleHeight     =   333
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   471
      TabIndex        =   0
      Top             =   390
      Width           =   7065
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
               Picture         =   "frmAllStudentAccount.frx":0ECA
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAllStudentAccount.frx":1464
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin HSES.b8SContainer b8SConStatus 
         Height          =   375
         Left            =   0
         TabIndex        =   1
         Top             =   4620
         Width           =   6315
         _ExtentX        =   11139
         _ExtentY        =   661
         Begin HSES.b8Nav b8NavRecord 
            Height          =   375
            Left            =   4380
            TabIndex        =   2
            Top             =   0
            Width           =   1800
            _ExtentX        =   3175
            _ExtentY        =   661
         End
         Begin VB.Label lblListInfo 
            AutoSize        =   -1  'True
            BackColor       =   &H00D8E9EC&
            BackStyle       =   0  'Transparent
            Caption         =   "No Record"
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
            Left            =   60
            TabIndex        =   4
            Top             =   75
            Width           =   855
         End
         Begin VB.Label lblPage 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00D8E9EC&
            BackStyle       =   0  'Transparent
            Caption         =   "Page 0 of 0"
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
            Left            =   3420
            TabIndex        =   3
            Top             =   75
            Width           =   930
         End
      End
      Begin MSComctlLib.ImageList ilRecordIco 
         Left            =   4530
         Top             =   3270
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
               Picture         =   "frmAllStudentAccount.frx":19FE
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin HSES.b8SContainer pbBGButton 
         Height          =   540
         Left            =   0
         TabIndex        =   6
         Top             =   345
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   953
         BorderColor     =   14215660
         Begin lvButton.lvButtons_H cmdShowAcademicRecord 
            Height          =   510
            Left            =   60
            TabIndex        =   7
            Top             =   60
            Width           =   1245
            _ExtentX        =   2196
            _ExtentY        =   900
            Caption         =   "Show Academic Record"
            CapAlign        =   2
            BackStyle       =   5
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cFore           =   8421504
            cFHover         =   8421504
            cBhover         =   14215660
            Focus           =   0   'False
            cGradient       =   14215660
            Gradient        =   4
            Mode            =   0
            Value           =   0   'False
            cBack           =   16185592
         End
         Begin lvButton.lvButtons_H cmdEnrolStudent 
            Height          =   510
            Left            =   1350
            TabIndex        =   10
            Top             =   60
            Width           =   645
            _ExtentX        =   1138
            _ExtentY        =   900
            Caption         =   "Enrol"
            CapAlign        =   2
            BackStyle       =   5
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cFore           =   8421504
            cFHover         =   8421504
            cBhover         =   14215660
            Focus           =   0   'False
            cGradient       =   14215660
            Gradient        =   4
            Mode            =   0
            Value           =   0   'False
            cBack           =   16185592
         End
      End
      Begin HSES.b8ChildTitleBar b8Title 
         Height          =   345
         Left            =   0
         TabIndex        =   8
         Top             =   0
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   609
         BackColor       =   12735512
         Caption         =   "Manage Student Accounts"
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
      Begin MSComctlLib.ListView listRecord 
         Height          =   1050
         Left            =   0
         TabIndex        =   9
         Top             =   945
         Width           =   5505
         _ExtentX        =   9710
         _ExtentY        =   1852
         View            =   3
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         Icons           =   "ilRecordIco"
         SmallIcons      =   "ilRecordIco"
         ColHdrIcons     =   "icoHeader"
         ForeColor       =   0
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
         NumItems        =   9
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Student ID"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Last Name"
            Object.Width           =   3704
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "First Name"
            Object.Width           =   3704
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Middle Name"
            Object.Width           =   3704
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Gender"
            Object.Width           =   1852
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "YL"
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Graduated"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Dropped"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Transferee"
            Object.Width           =   2117
         EndProperty
      End
   End
   Begin HSES.b8Container b8cMain 
      Height          =   5940
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   7035
      _ExtentX        =   12409
      _ExtentY        =   10478
      BorderColor     =   12632256
      BackColor       =   16185592
      InsideBorderColor=   14215660
      ShadowColor1    =   16777215
      ShadowColor2    =   16777215
   End
End
Attribute VB_Name = "frmAllStudent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim vRS As New ADODB.Recordset

Dim sDefaultSQL As String

Private Const sAllFields = "ALL FIELDS"

Private Form_Fields() As String
Private Form_OrigFields() As String

Dim MaxEntryCount As Long
Dim CurRecPos As Long
Dim CurStudentCount As Long

Dim curFilter As String
Dim curFilterMSG As String

Public Sub ShowFormList(Optional iMaxEntryCount As Long = 19, Optional iCurRecPos As Long = 0, Optional sFilter As String = "", Optional sFilterMSG As String = "")
    
    'show form
    mdiMain.MousePointer = vbHourglass
    
    DoEvents
    
    
    'apply parameter
    MaxEntryCount = iMaxEntryCount
    CurRecPos = iCurRecPos
    
    curFilter = sFilter
    curFilterMSG = sFilterMSG
    
    If sFilter = "" Then
        sFilter = ""
    Else
        sFilter = " WHERE " & sFilter
    End If
    
    'set default SQL
    sDefaultSQL = "SELECT tblStudent.StudentID AS lvKEY, [tblStudent]![StudentID] AS [Student ID], tblStudent.LastName AS [Last Name], tblStudent.FirstName AS [First Name], tblStudent.MiddleName AS [Middle Name], tblStudent.Gender, tblYearLevel.YearLevelTitle, IIf(Len(Trim([tblGraduate].[StudentID]))=12,'Yes','') AS Graduated, IIf(Len(Trim([tblDropped].[StudentID]))=12,'Yes',' ') AS Dropped, iif(tblStudent.Transferee=True,'Yes',' ') AS IsTransferee" & _
                    " FROM tblDropped RIGHT JOIN (tblGraduate RIGHT JOIN (tblDepartment RIGHT JOIN (tblYearLevel RIGHT JOIN (tblSection RIGHT JOIN (tblSectionOffering RIGHT JOIN (tblStudent LEFT JOIN tblEnrolment ON tblStudent.StudentID = tblEnrolment.StudentID) ON tblSectionOffering.SectionOfferingID = tblEnrolment.SectionOfferingID) ON tblSection.SectionID = tblSectionOffering.SectionID) ON tblYearLevel.YearLevelID = tblSection.YearLevelID) ON tblDepartment.DepartmentID = tblSection.DepartmentID) ON tblGraduate.StudentID = tblStudent.StudentID) ON tblDropped.StudentID = tblStudent.StudentID" & _
                    " " & sFilter & _
                    " ORDER BY tblStudent.LastName, tblStudent.FirstName, tblStudent.MiddleName;"

    
    'connect rs
    If ConnectRS(HSESDB, vRS, sDefaultSQL) Then
        b8navRecordRefresh
        Form_SetFieldList
        FillList vRS
        
        Me.Show
        Me.SetFocus
        
        
    Else
        MsgBox "Unable to show Student List.", vbCritical
        Unload Me
    End If
    
    mdiMain.MousePointer = vbDefault
    
End Sub


Private Function FillList(ByRef vRS As ADODB.Recordset)
        
        mdiMain.MousePointer = vbHourglass


        UnSortLV listRecord
        
        FillRecordToList vRS, listRecord, KeyStudent, CurRecPos, MaxEntryCount, , True
        
        SortLV listRecord, listRecord.SortKey, listRecord.SortOrder, False
        
        'refresh list info
        listRecord_Click

        
        mdiMain.MousePointer = vbDefault
End Function


    
















Private Sub b8navRecordRefresh()
    CurStudentCount = getRecordCount(vRS)
    
    If CurRecPos >= CurStudentCount Then
        If CurStudentCount > 1 Then
            'click previous
            b8navRecord_Click 1
        End If
    End If
    
    If CurRecPos > 0 Then
        b8NavRecord.FirstEnable = True
        b8NavRecord.PreviousEnable = True
    Else
        b8NavRecord.FirstEnable = False
        b8NavRecord.PreviousEnable = False
    End If
    
    If CurRecPos < CurStudentCount - MaxEntryCount Then
        b8NavRecord.LastEnable = True
        b8NavRecord.NextEnable = True
    Else
        b8NavRecord.LastEnable = False
        b8NavRecord.NextEnable = False
    End If
    
    listRecord_Click
End Sub

Private Sub b8navRecord_Click(Index As Integer)
    Select Case Index
        Case 0
            CurRecPos = 0
            FillList vRS
            listRecord_Click
            
            
        Case 1
            If CurRecPos - MaxEntryCount >= 0 Then
        
                CurRecPos = CurRecPos - MaxEntryCount
                
                FillList vRS
                listRecord_Click
            End If
            
            
        Case 2
            If CurRecPos + MaxEntryCount < getRecordCount(vRS) Then
        
                CurRecPos = CurRecPos + MaxEntryCount
                        
                FillList vRS
                listRecord_Click
            End If
    
    
        Case 3
        
            Dim RC As Long
    
            RC = getRecordCount(vRS)
            
            If MaxEntryCount < RC Then
                'temp
                'pwede pa mapababa
                If (RC Mod MaxEntryCount) = 0 Then
                    CurRecPos = RC - MaxEntryCount
                Else
                    CurRecPos = RC - (RC Mod MaxEntryCount)
                End If
                    
                FillList vRS
                
                listRecord_Click
            End If
    End Select
    
    'refresh buttons
    b8navRecordRefresh
End Sub







Private Sub cmdEnrol_Click()
    Dim lvKey As String
    
    lvKey = GetLVKey(listRecord.SelectedItem)
    
    If Len(lvKey) > 0 Then
        frmAddEnrolment.ShowForm lvKey
    Else
        MsgBox "Please select Student", vbExclamation
    End If
End Sub


Private Sub cmdShowEnrolment_Click()
    Dim lvKey As String
    
    lvKey = GetLVKey(listRecord.SelectedItem)
    
    If Len(lvKey) > 0 Then
        frmAllEnrolment.ShowFormList , , , , lvKey
    Else
        MsgBox "Please select Department", vbExclamation
    End If
End Sub

Private Sub cmdEnrolStudent_Click()
    frmAddEnrolment.ShowForm GetLVKey(listRecord.SelectedItem)
End Sub

Private Sub cmdShowAcademicRecord_Click()
    Dim sKey As String
    
    sKey = GetLVKey(listRecord.SelectedItem)
    
    If sKey <> "" Then
        frmStudentRecord.ShowForm sKey
    End If
End Sub

Private Sub listRecord_Click()
    Dim totalPage As Long
    Dim curPage As Long
    
    If listRecord.ListItems.Count < 1 Then
        lblListInfo.Caption = "No Record"
        lblPage.Caption = "Page 0 of 0"
    Else
        lblListInfo.Caption = "Selected Entry: " & listRecord.SelectedItem.Index + CurRecPos & "/" & CurStudentCount
        
        If (CurStudentCount Mod MaxEntryCount) > 0 Then
            totalPage = (CurStudentCount \ MaxEntryCount) + 1
        Else
            totalPage = (CurStudentCount \ MaxEntryCount)
        End If
        
        lblPage.Caption = "Page " & ((CurRecPos \ MaxEntryCount) + 1) & " of " & totalPage
    End If
End Sub


Private Sub listRecord_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    SortLV listRecord, ColumnHeader.Index - 1
End Sub

Private Sub listRecord_DblClick()
    Dim sKey As String
    
    sKey = GetLVKey(listRecord.SelectedItem)
    
    If sKey <> "" Then
        frmStudentRecord.ShowForm sKey
    End If
End Sub

Private Sub listRecord_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim curPos As Long
    If KeyCode = vbKeyDown Then
        If listRecord.SelectedItem.Index = listRecord.ListItems.Count Then
            
            b8navRecord_Click 2
            
            KeyCode = 0
        End If
    End If
    
    If KeyCode = vbKeyUp Then
        If listRecord.SelectedItem.Index = 1 Then
            curPos = CurRecPos
            
            b8navRecord_Click 1
            
            If curPos <> CurRecPos Then
                listRecord.SelectedItem.Selected = False
                listRecord.ListItems(listRecord.ListItems.Count).Selected = True
            End If
            
            KeyCode = 0
        End If
    End If
    
    If KeyCode = vbKeyPageDown Then
        b8navRecord_Click 2
    End If
    
    If KeyCode = vbKeyPageUp Then
        b8navRecord_Click 1
    End If
End Sub

Private Sub listRecord_KeyUp(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = 38 Or KeyCode = 40 Or KeyCode = 33 Or KeyCode = 34 Then listRecord_Click
    
    If KeyCode = vbKeyDelete Then
        Form_Delete
    End If
End Sub


Private Sub Form_Activate()
    mdiMain.RegMDIChild Me
    Me.WindowState = vbMaximized
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
    listRecord.Height = bgMain.Height - (listRecord.Top + b8SConStatus.Height)
    b8SConStatus.Move -1, bgMain.Height + 1 - b8SConStatus.Height, bgMain.Width + 1
    b8NavRecord.Left = b8SConStatus.Width * Screen.TwipsPerPixelX - b8NavRecord.Width
    lblPage.Left = (b8SConStatus.Width * Screen.TwipsPerPixelX) - b8NavRecord.Width - lblPage.Width - 30

    
    
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Set vRS = Nothing
    
    Set frmAllStudent = Nothing
End Sub
















'---------------------------------------------------------------
'Form Operations
'---------------------------------------------------------------

Public Function Form_Find()
    frmFindListItem.ShowFind listRecord
End Function

Public Sub Form_Add()
    'show add
    If frmAddStudent.ShowForm Then
        Form_Refresh
        b8navRecordRefresh
    End If
End Sub
Public Sub Form_Edit()
    'check if there is a record in the list
    If listRecord.ListItems.Count < 1 Then Exit Sub

    If Len(GetLVKey(listRecord.SelectedItem)) < 1 Then Exit Sub
    
    'show edit
    frmEditStudent.ShowEdit GetLVKey(listRecord.SelectedItem)
End Sub
Public Sub Form_Delete()
    
    If frmDeleteStudent.ShowForm(GetLVKey(listRecord.SelectedItem)) = True Then
        Form_Reload
    End If
    
End Sub


Public Sub Form_Refresh()
    vRS.Requery
    FillList vRS
    b8navRecordRefresh
End Sub

Public Sub Form_Reload()
    If ConnectRS(HSESDB, vRS, sDefaultSQL) = False Then
            'temp
            'fatal
            MsgBox "FATAL ERROR: frmAllStudent.Form_filter - Connecting Default StudentRecordset.", vbExclamation
            'close this form
            Unload Me
            Exit Sub
    End If
    
    b8navRecordRefresh
    Form_SetFieldList
    Form_Refresh
End Sub







Public Function Form_CanFind() As Boolean
    If listRecord.ListItems.Count > 0 Then
        Form_CanFind = True
    Else
        Form_CanFind = False
    End If
End Function

Public Function Form_CanFilter(ByRef FieldList() As String) As Boolean
    
    
    If Len(curFilter) > 0 Then
    
        Form_CanFilter = False
        
    Else
    
        ReDim FieldList(UBound(Form_Fields))
        
    
        If UBound(Form_Fields) > 0 Then
            FieldList = Form_Fields
            Form_CanFilter = True
        Else
            Form_CanFilter = False
        End If
        
    End If
End Function

Public Function Form_Filter(sFindWhat As String, sFieldName As String)
    
    Dim i As Integer
    Dim FIndex As Integer
    Dim bFieldMatch As Boolean
    Dim newFindWhat As String
    
    If Len(sFindWhat) < 1 Then
        Form_Reload
        Exit Function
    End If
    If Len(sFieldName) < 1 Then Exit Function
    
    Dim sSQL As String
    
    'get index
    bFieldMatch = False
    For i = 0 To UBound(Form_Fields)
        If Form_Fields(i) = sFieldName Then
            bFieldMatch = True
            FIndex = i
            Exit For
        End If
    Next
    
    If bFieldMatch <> True Then
        
        MsgBox "Please Select correct Field.", vbExclamation
        Exit Function
    End If
    
    
    If Form_OrigFields(FIndex) = sAllFields Then
                               
        sSQL = "SELECT tblStudent.StudentID AS lvKEY, [tblStudent]![StudentID] AS [Student ID], tblStudent.LastName AS [Last Name], tblStudent.FirstName AS [First Name], tblStudent.MiddleName AS [Middle Name], tblStudent.Gender, tblYearLevel.YearLevelTitle, IIf(Len(Trim([tblGraduate].[StudentID]))=12,'Yes','') AS Graduated, IIf(Len(Trim([tblDropped].[StudentID]))=12,'Yes',' ') AS Dropped, iif(tblStudent.Transferee=True,'Yes',' ') AS IsTransferee" & _
                " FROM tblDropped RIGHT JOIN (tblGraduate RIGHT JOIN (tblDepartment RIGHT JOIN (tblYearLevel RIGHT JOIN (tblSection RIGHT JOIN (tblSectionOffering RIGHT JOIN (tblStudent LEFT JOIN tblEnrolment ON tblStudent.StudentID = tblEnrolment.StudentID) ON tblSectionOffering.SectionOfferingID = tblEnrolment.SectionOfferingID) ON tblSection.SectionID = tblSectionOffering.SectionID) ON tblYearLevel.YearLevelID = tblSection.YearLevelID) ON tblDepartment.DepartmentID = tblSection.DepartmentID) ON tblGraduate.StudentID = tblStudent.StudentID) ON tblDropped.StudentID = tblStudent.StudentID" & _
                " Where "
                
        For i = 1 To UBound(Form_Fields)
            If i <> 1 Then
                sSQL = sSQL & " OR "
            End If
            sSQL = sSQL & " (((" & Form_OrigFields(i) & ") like "
            
          
            Select Case vRS.Fields(i).Type
                Case adDate
                    If IsDate(sFindWhat) Then
                    
                        newFindWhat = "#" & FormatDateTime(sFindWhat, vbShortDate) & "#"
                    Else
                        newFindWhat = "#" & "1/1/999" & "#"
                    End If
                Case Else
                    'string
                    newFindWhat = "'%" & sFindWhat & "%'"
            End Select
        
            sSQL = sSQL & newFindWhat & "))"
        Next
                
        sSQL = sSQL & " ORDER BY tblStudent.LastName, tblStudent.FirstName, tblStudent.MiddleName;"

    
    Else
    
        Select Case vRS.Fields(FIndex).Type
            Case adDate
                If IsDate(sFindWhat) Then
                
                    sFindWhat = "#" & FormatDateTime(sFindWhat, vbShortDate) & "#"
                Else
                    sFindWhat = "#" & "1/1/999" & "#"
                End If
            Case Else
                'string
                sFindWhat = "'%" & sFindWhat & "%'"
        End Select
        
        sSQL = "SELECT tblStudent.StudentID AS lvKEY, [tblStudent]![StudentID] AS [Student ID], tblStudent.LastName AS [Last Name], tblStudent.FirstName AS [First Name], tblStudent.MiddleName AS [Middle Name], tblStudent.Gender, tblYearLevel.YearLevelTitle, IIf(Len(Trim([tblGraduate].[StudentID]))=12,'Yes','') AS Graduated, IIf(Len(Trim([tblDropped].[StudentID]))=12,'Yes',' ') AS Dropped, iif(tblStudent.Transferee=True,'Yes',' ') AS IsTransferee" & _
                " FROM tblDropped RIGHT JOIN (tblGraduate RIGHT JOIN (tblDepartment RIGHT JOIN (tblYearLevel RIGHT JOIN (tblSection RIGHT JOIN (tblSectionOffering RIGHT JOIN (tblStudent LEFT JOIN tblEnrolment ON tblStudent.StudentID = tblEnrolment.StudentID) ON tblSectionOffering.SectionOfferingID = tblEnrolment.SectionOfferingID) ON tblSection.SectionID = tblSectionOffering.SectionID) ON tblYearLevel.YearLevelID = tblSection.YearLevelID) ON tblDepartment.DepartmentID = tblSection.DepartmentID) ON tblGraduate.StudentID = tblStudent.StudentID) ON tblDropped.StudentID = tblStudent.StudentID" & _
                " Where (((" & Form_OrigFields(FIndex) & ") like " & sFindWhat & "))" & _
                " ORDER BY tblStudent.LastName, tblStudent.FirstName, tblStudent.MiddleName;"
    End If
    

    'connect srs
    If ConnectRS(HSESDB, vRS, sSQL) = True Then
        CurRecPos = 0
        b8navRecordRefresh
        Form_SetFieldList
        Call Form_Refresh
    Else
    
        listRecord.ListItems.Clear
        
        If ConnectRS(HSESDB, vRS, sDefaultSQL) = False Then
            'temp
            'fatal
            MsgBox "FATAL ERROR: frmAllStudent.Form_filter - Connecting Default StudentRecordset.", vbExclamation
            'close this form
            Unload Me
        Else
            CurRecPos = 0
            Form_SetFieldList
            b8navRecordRefresh
        End If
        
    End If
    
    
End Function

Public Function Form_CanAdvanceFilter() As Boolean
        Form_CanAdvanceFilter = True

End Function

Public Function Form_AdvanceFilter()
    frmASStudent.ShowForm
End Function


Private Function Form_SetFieldList()
    Dim rsF As Field
    Dim i As Integer

    ReDim Form_Fields(vRS.Fields.Count - 1) As String
    ReDim Form_OrigFields(vRS.Fields.Count - 1) As String
    i = 0
    
    For Each rsF In vRS.Fields

        Form_OrigFields(i) = "[" & rsF.Properties.Item(1).Value & "]![" & rsF.Properties.Item(0).Value & "]"
        
        Form_Fields(i) = rsF.Name
        
        i = i + 1
    Next
    
    'change lvkey to ALL FIelds
    Form_OrigFields(0) = sAllFields
    Form_Fields(0) = sAllFields
    
    ReDim Preserve Form_OrigFields(7)
    ReDim Preserve Form_Fields(7)
End Function
'form
Public Function Form_Description() As String
    Form_Description = "Display All Student Account"
End Function
Public Function Form_Tip() As String
    
    Form_Tip = "First, select an entry. CLick Grades to view all grades of select entry."
End Function


'Record Operations
Public Function Form_CanAddEntry() As Boolean
    Form_CanAddEntry = True
End Function

Public Function Form_CanEditEntry() As Boolean
    If listRecord.ListItems.Count > 0 Then
        Form_CanEditEntry = True
    Else
        Form_CanEditEntry = False
    End If
End Function

Public Function Form_CanDeleteEntry() As Boolean
    If listRecord.ListItems.Count > 0 Then
        Form_CanDeleteEntry = True
    Else
        Form_CanDeleteEntry = False
    End If
End Function

Public Function Form_Can_Reload() As Boolean
    Form_Can_Reload = True
End Function

Public Function Form_CanShowListOption() As Boolean
    If listRecord.ListItems.Count > 0 Then
        Form_CanShowListOption = True
    Else
        Form_CanShowListOption = False
    End If
End Function

Public Function Form_CanResizeListFont() As Boolean
    If listRecord.ListItems.Count > 0 Then
        Form_CanResizeListFont = True
    Else
        Form_CanResizeListFont = False
    End If
End Function
Public Function Form_CanChangeListFont() As Boolean
    If listRecord.ListItems.Count > 0 Then
        Form_CanChangeListFont = True
    Else
        Form_CanChangeListFont = False
    End If
End Function

Public Function Form_CanPrint() As Boolean

    If listRecord.ListItems.Count > 0 Then
        Form_CanPrint = True
    Else
        Form_CanPrint = False
    End If
End Function

Public Function Form_Print()
    frmPrintStudent.ShowForm , , , , GetLVKey(listRecord.SelectedItem)
End Function
'----------------------------------------------------------------
'END FORM OPERATIONS
'----------------------------------------------------------------





