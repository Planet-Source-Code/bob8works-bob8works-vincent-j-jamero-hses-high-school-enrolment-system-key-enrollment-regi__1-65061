VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmAllDepartment 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Departments"
   ClientHeight    =   6390
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8025
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
   Icon            =   "frmAllDepartment.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   426
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   535
   Begin VB.PictureBox bgMain 
      Appearance      =   0  'Flat
      BackColor       =   &H00D8E9EC&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4995
      Left            =   810
      ScaleHeight     =   333
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   471
      TabIndex        =   0
      Top             =   450
      Width           =   7065
      Begin MSComctlLib.ImageList icoHeader 
         Left            =   5610
         Top             =   3225
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
               Picture         =   "frmAllDepartment.frx":0ECA
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAllDepartment.frx":1464
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin HSES.b8SContainer b8SConStatus 
         Height          =   375
         Left            =   0
         TabIndex        =   2
         Top             =   4620
         Width           =   6315
         _ExtentX        =   11139
         _ExtentY        =   661
         Begin HSES.b8Nav b8NavRecord 
            Height          =   375
            Left            =   4380
            TabIndex        =   3
            Top             =   0
            Width           =   1800
            _ExtentX        =   3175
            _ExtentY        =   661
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
            TabIndex        =   5
            Top             =   75
            Width           =   930
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
      End
      Begin MSComctlLib.ImageList ilRecordIco 
         Left            =   5580
         Top             =   2490
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
               Picture         =   "frmAllDepartment.frx":19FE
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ListView listRecord 
         Height          =   3480
         Left            =   0
         TabIndex        =   1
         Top             =   1110
         Width           =   5505
         _ExtentX        =   9710
         _ExtentY        =   6138
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
         ForeColor       =   8399906
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
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Title"
            Object.Width           =   5821
         EndProperty
      End
      Begin HSES.b8SContainer pbBGButton 
         Height          =   585
         Left            =   0
         TabIndex        =   6
         Top             =   345
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   1032
         BorderColor     =   14215660
         Begin lvButton.lvButtons_H cmdShowSection 
            Height          =   510
            Left            =   60
            TabIndex        =   7
            Top             =   75
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   900
            Caption         =   "Show Sections"
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
         Begin lvButton.lvButtons_H cmdShowEnrolment 
            Height          =   510
            Left            =   990
            TabIndex        =   8
            Top             =   75
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   900
            Caption         =   "Show Enrolments"
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
         TabIndex        =   9
         Top             =   0
         Width           =   6645
         _ExtentX        =   11721
         _ExtentY        =   609
         BackColor       =   12735512
         Caption         =   "Manage Department"
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
   End
   Begin HSES.b8Container b8cMain 
      Height          =   5940
      Left            =   720
      TabIndex        =   10
      Top             =   90
      Width           =   7035
      _ExtentX        =   12409
      _ExtentY        =   10478
      BorderColor     =   12632256
      BackColor       =   16185592
      ShadowColor1    =   16777215
      ShadowColor2    =   16777215
   End
End
Attribute VB_Name = "frmAllDepartment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Const sAllFields = "ALL FIELDS"



Dim vRS As New ADODB.Recordset


Private Form_Fields() As String
Private Form_OrigFields() As String

Dim sDefaultSQL As String


Dim MaxEntryCount As Long
Dim CurRecPos As Long
Dim CurStudentCount As Long



Public Sub ShowFormList(Optional iMaxEntryCount As Long = 21, Optional iCurRecPos As Long = 0)
    
    Dim sSQL As String
  
    
    'apply parameter
    MaxEntryCount = iMaxEntryCount
    CurRecPos = iCurRecPos
    'CurStudentCount
    
    'set default SQL)
    sDefaultSQL = "SELECT tblDepartment.DepartmentID as lvKey, tblDepartment.DepartmentTitle From tblDepartment"
    
    sSQL = "SELECT tblDepartment.DepartmentID as lvKey, tblDepartment.DepartmentTitle From tblDepartment"

    'connect rs
    If ConnectRS(HSESDB, vRS, sSQL) Then
                
        b8navRecordRefresh
        Form_SetFieldList
        FillList vRS
        
        'show form
        Me.Show
        Me.SetFocus
        
        
        
    Else
        MsgBox "Unable to show Department List.", vbCritical
        Unload Me
    End If
    
End Sub


Private Function FillList(ByRef vRS As ADODB.Recordset)
        
        mdiMain.MousePointer = vbHourglass

        'fill list to listview
        Form_SetFieldList
        
        UnSortLV listRecord
        
        FillRecordToList vRS, listRecord, KeyStudent, CurRecPos, MaxEntryCount, , True
        
        SortLV listRecord, listRecord.SortKey, listRecord.SortOrder, False
        
        'refresh list info
        listRecord_Click
        
        'refresh parent form
        mdiMain.RegMDIChild Me
        
        mdiMain.MousePointer = vbDefault
End Function


Private Sub b8navRecordRefresh()
    CurStudentCount = getRecordCount(vRS)
    
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


Private Sub cmdShowEnrolment_Click()
    
    If Len(listRecord.SelectedItem.Text) > 0 Then
        frmAllEnrolment.ShowFormList , listRecord.SelectedItem.Text
    Else
        MsgBox "Please select Department in the list.", vbExclamation
    End If
End Sub

Private Sub cmdShowSection_Click()
    Dim lvKey As String
    
    lvKey = GetLVKey(listRecord.SelectedItem)
    
    If Len(lvKey) > 0 Then
        frmAllSection.ShowFormList lvKey
    Else
        MsgBox "Please select Department", vbExclamation
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
    
    Set frmAllDepartment = Nothing
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
Private Sub listRecord_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
        Me.PopupMenu mdiMain.mnuEdit
    End If
End Sub





























'---------------------------------------------------------------
'Form Operations
'---------------------------------------------------------------


Public Function Form_Explore()

End Function
Public Function Form_Find()
    frmFindListItem.ShowFind listRecord
End Function

Public Sub Form_Add()
    'show add
    If frmAddDepartment.ShowForm Then Form_Refresh
End Sub
Public Sub Form_Edit()
    'check if there is a record in the list
    If listRecord.ListItems.Count < 1 Then Exit Sub

    If Len(GetLVKey(listRecord.SelectedItem)) < 1 Then Exit Sub
    
    'show edit
    frmEditDepartment.ShowEdit GetLVKey(listRecord.SelectedItem)

    'refresh list
    Me.Form_Reload
End Sub
Public Sub Form_Delete()

    If Len(listRecord.SelectedItem.Text) > 0 Then
        frmDeleteDepartment.ShowForm listRecord.SelectedItem.Text
    End If
    
    Me.Form_Reload
    
End Sub



Public Sub Form_Refresh()
    vRS.Requery
    FillList vRS
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
    
    ReDim FieldList(UBound(Form_Fields))
    

    If UBound(Form_Fields) > 0 Then
        FieldList = Form_Fields
        Form_CanFilter = True
    Else
        Form_CanFilter = False
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
    
        sSQL = "SELECT tblDepartment.DepartmentID as lvKey, tblDepartment.DepartmentTitle" & _
                " From tblDepartment" & _
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
                
        sSQL = sSQL & " ORDER BY tblDepartment.DepartmentTitle;"
    
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
        
        sSQL = "SELECT tblDepartment.DepartmentID as lvKey, tblDepartment.DepartmentTitle" & _
                " From tblDepartment" & _
                " Where (((" & Form_OrigFields(FIndex) & ") like " & sFindWhat & "))" & _
                " ORDER BY tblDepartment.DepartmentTitle;"
    End If
    

    'connect srs
    If ConnectRS(HSESDB, vRS, sSQL) = True Then
        CurRecPos = 0
        b8navRecordRefresh
        Form_SetFieldList
        Call Form_Refresh
    Else
    
        listRecord.ListItems.Clear
        
        If ConnectRS(HSESDB, vRS, sDefaultSQL & " ORDER BY tblDepartment.DepartmentTitle") = False Then
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
    Form_CanAdvanceFilter = False
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
    
End Function
'form
Public Function Form_Description() As String
    Form_Description = "Displays Department Entries"
End Function
Public Function Form_Tip() As String
    
    Form_Tip = ""
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
    frmPrintDepartment.ShowForm
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

Public Function Form_GetMenu(ByRef sMenu() As String) As Boolean
    
    ReDim sMenu(1)
    
    sMenu(0) = "Show Sections"
    sMenu(1) = "Show Enrolments"

    
    Form_GetMenu = True
End Function

Public Function Form_MenuClick(sMenu As String)

    
    Select Case sMenu
    
        Case "Show Sections"
            Call cmdShowSection_Click
        Case "Show Enrolments"
            Call cmdShowEnrolment_Click

            
    End Select
End Function























Private Sub listRecord_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    'sort list according to header
    SortLV listRecord, ColumnHeader.Index - 1
End Sub

Private Sub listRecord_DblClick()
    Dim lvKey As String
    
    lvKey = GetLVKey(listRecord.SelectedItem)
    
    If Len(lvKey) > 0 Then
        frmAllSection.ShowFormList lvKey
    Else
        MsgBox "Please select Department", vbExclamation
    End If

End Sub

Private Sub listRecord_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyDelete
            If Me.Form_CanDeleteEntry = True Then
                Me.Form_Delete
            End If
    End Select
End Sub
