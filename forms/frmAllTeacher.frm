VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAllTeacher 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Teachers"
   ClientHeight    =   6375
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7020
   ControlBox      =   0   'False
   Icon            =   "frmAllTeacher.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   425
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   468
   Begin VB.PictureBox bgMain 
      Appearance      =   0  'Flat
      BackColor       =   &H00D8E9EC&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4995
      Left            =   360
      ScaleHeight     =   333
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   471
      TabIndex        =   0
      Top             =   480
      Width           =   7065
      Begin VB.PictureBox pbBGButton 
         Appearance      =   0  'Flat
         BackColor       =   &H00D8E9EC&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   525
         Left            =   30
         ScaleHeight     =   525
         ScaleWidth      =   6675
         TabIndex        =   8
         Top             =   390
         Width           =   6675
      End
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
               Picture         =   "frmAllTeacher.frx":0ECA
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAllTeacher.frx":1464
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
            TabIndex        =   4
            Top             =   90
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
            TabIndex        =   3
            Top             =   75
            Width           =   855
         End
      End
      Begin MSComctlLib.ListView listRecord 
         Height          =   3480
         Left            =   0
         TabIndex        =   5
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
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Log-In Name"
            Object.Width           =   5821
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Full Name"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Creation Date"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Contact Number"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Address"
            Object.Width           =   5292
         EndProperty
      End
      Begin HSES.b8ChildTitleBar b8Title 
         Height          =   345
         Left            =   0
         TabIndex        =   6
         Top             =   0
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   609
         BackColor       =   16777215
         Caption         =   "Manage Teachers"
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
               Picture         =   "frmAllTeacher.frx":19FE
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin HSES.b8Container b8cMain 
      Height          =   5940
      Left            =   30
      TabIndex        =   7
      Top             =   60
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
Attribute VB_Name = "frmAllTeacher"
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


Public Sub ShowFormList(Optional iMaxEntryCount As Long = 21, Optional iCurRecPos As Long = 0)
    
    Dim sSQL As String
    

    
    'apply parameter
    MaxEntryCount = iMaxEntryCount
    CurRecPos = iCurRecPos
    
    'default
    sDefaultSQL = "SELECT tblTeacher.TeacherID AS lvKey, tblTeacher.TeacherTitle AS LoginName, [TblTeacher]![LastName]+', '+[tblTeacher]![FirstName]+' '+[tblTeacher]![MiddleName] AS FullName, tblTeacher.CreationDate AS [Creation Date], tblTeacher.ContactNumber AS [Contact Number], tblTeacher.Address" & _
                    " FROM tblTeacher"

    sSQL = "SELECT tblTeacher.TeacherID AS lvKey, tblTeacher.TeacherTitle AS LoginName, [TblTeacher]![LastName]+', '+[tblTeacher]![FirstName]+' '+[tblTeacher]![MiddleName] AS FullName, tblTeacher.CreationDate AS [Creation Date], tblTeacher.ContactNumber AS [Contact Number], tblTeacher.Address" & _
                    " FROM tblTeacher;"
    
    If ConnectRS(HSESDB, vRS, sSQL) = True Then
        b8navRecordRefresh
        Form_SetFieldList
        FillList vRS
            
        'show form
        Me.Show
        Me.SetFocus
    Else
        MsgBox "Unable to show Teacher list.", vbCritical
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


Public Sub Form_Find()
    frmFindListItem.ShowFind listRecord
End Sub
Public Sub Form_Add()
    'show add
    frmAddTeacher.ShowForm
End Sub




Public Sub Form_Edit()
    
    'check if there is a record in the list
    If listRecord.ListItems.Count < 1 Then Exit Sub

    If Len(GetLVKey(listRecord.SelectedItem)) < 1 Then Exit Sub
    
    'show edit
    frmEditTeacher.ShowEdit GetLVKey(listRecord.SelectedItem)
    
    Form_Reload
End Sub


Public Sub Form_Delete()
    Dim i As Integer
    Dim iSelectedCount As Integer
    Dim sMSG As String
    Dim sLVItemKey As String
    Dim EntryDeletedCount As Integer
    
    
    'defaults
    EntryDeletedCount = 0
    
    'get record count
    iSelectedCount = GetLVSelectedCount(listRecord)
    
    'check if there is a record in the list
    If iSelectedCount < 1 Then
        MsgBox "No selected entry to delete." & _
            vbNewLine & "Please select it first in the list.", vbExclamation
        
        Exit Sub
    End If
    
    If iSelectedCount = 1 Then
        sMSG = "WARNING:" & vbNewLine & "You are about to delete 1 Teacher Entry and you cannot Undo this operation." & vbNewLine & _
        " Are you sure to delete it?"
    Else
        sMSG = "WARNING:" & vbNewLine & "You are about to delete " & iSelectedCount & " Teacher Entries and you cannot Undo this operation." & vbNewLine & _
        " Are you sure to delete it?"
    End If
    
    If MsgBox(sMSG, vbQuestion + vbYesNo) = vbYes Then
        
        For i = 1 To listRecord.ListItems.Count
            
            'get key
            sLVItemKey = GetLVKey(listRecord.ListItems(i))
            If listRecord.ListItems(i).Selected = True And Len(sLVItemKey) > 0 Then
                   
                If DeleteTeacher(sLVItemKey) = Success Then
                    EntryDeletedCount = EntryDeletedCount + 1
                Else
                    'temp
                    'fatal
                    MsgBox "Unable to delete Teacher Entry with ID: " & sLVItemKey
                End If
            End If
            
        Next

    
        If EntryDeletedCount > 0 Then
            'show info
            MsgBox EntryDeletedCount & " Entry/s deleted.", vbInformation
            'refresh list
            Me.Form_Refresh
        End If
    
    End If 'user is sure to delete entry
End Sub



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

Public Sub Form_Refresh()
    vRS.Requery
    FillList vRS
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
    
        sSQL = sDefaultSQL & " Where "
                
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
                
        sSQL = sSQL & " ORDER BY tblTeacher.TeacherTitle;"
    
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
        
        sSQL = sDefaultSQL & _
                " Where (((" & Form_OrigFields(FIndex) & ") like " & sFindWhat & "))" & _
                " ORDER BY tblTeacher.TeacherTitle;"
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
    
    Form_OrigFields(2) = "[TblTeacher]![LastName]+', '+[tblTeacher]![FirstName]+' '+[tblTeacher]![MiddleName]"
    Form_Fields(2) = "Full Name"
        
    'change lvkey to ALL FIelds
    Form_OrigFields(0) = sAllFields
    Form_Fields(0) = sAllFields
    
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

'----------------------------------------------------------------
'END FORM OPERATIONS
'----------------------------------------------------------------


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
    Dim lvKey As String
    
    lvKey = GetLVKey(listRecord.SelectedItem)
    
    If lvKey <> "" Then
        frmTeacherRecord.ShowForm lvKey
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
