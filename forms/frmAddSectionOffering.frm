VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmAddSectionOffering 
   BackColor       =   &H00D8E9EC&
   Caption         =   "Section"
   ClientHeight    =   5205
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9915
   Icon            =   "frmAddSectionOffering.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   347
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   661
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab tabMain 
      Height          =   4110
      Left            =   60
      TabIndex        =   4
      Top             =   600
      Width           =   9780
      _ExtentX        =   17251
      _ExtentY        =   7250
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   529
      BackColor       =   14215660
      TabCaption(0)   =   "General"
      TabPicture(0)   =   "frmAddSectionOffering.frx":058A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "b8cTabBg(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Subjects"
      TabPicture(1)   =   "frmAddSectionOffering.frx":05A6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "b8cTabBg(1)"
      Tab(1).ControlCount=   1
      Begin HSES.b8Container b8cTabBg 
         Height          =   3675
         Index           =   0
         Left            =   60
         TabIndex        =   5
         Top             =   360
         Width           =   9675
         _ExtentX        =   17066
         _ExtentY        =   6482
         BorderColor     =   12307149
         BackColor       =   16185592
         ShadowColor1    =   13427430
         ShadowColor2    =   14215660
         Begin VB.ComboBox cmbRoom 
            Height          =   315
            Left            =   1590
            TabIndex        =   36
            Top             =   3165
            Width           =   3210
         End
         Begin VB.CommandButton cmdGetTeacher 
            BackColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   4440
            Picture         =   "frmAddSectionOffering.frx":05C2
            Style           =   1  'Graphical
            TabIndex        =   26
            Top             =   1530
            Width           =   345
         End
         Begin VB.CommandButton cmdGetSectionTitle 
            BackColor       =   &H00FFFFFF&
            Height          =   270
            Left            =   4470
            Picture         =   "frmAddSectionOffering.frx":0B4C
            Style           =   1  'Graphical
            TabIndex        =   25
            Top             =   720
            Width           =   345
         End
         Begin VB.CommandButton cmdGetSchoolYear 
            BackColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   4440
            Picture         =   "frmAddSectionOffering.frx":10D6
            Style           =   1  'Graphical
            TabIndex        =   24
            Top             =   1110
            Width           =   345
         End
         Begin VB.TextBox txtSectionOfferingID 
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
            Left            =   1590
            Locked          =   -1  'True
            MaxLength       =   20
            TabIndex        =   14
            Top             =   240
            Width           =   3225
         End
         Begin VB.TextBox txtSectionFullTitle 
            Height          =   330
            Left            =   1590
            MaxLength       =   50
            TabIndex        =   13
            Top             =   690
            Width           =   3225
         End
         Begin VB.TextBox txtSchoolYearTitle 
            Height          =   345
            Left            =   1590
            MaxLength       =   50
            TabIndex        =   12
            Top             =   1080
            Width           =   3225
         End
         Begin VB.TextBox txtTeacherFullName 
            Height          =   345
            Left            =   1590
            Locked          =   -1  'True
            MaxLength       =   20
            TabIndex        =   11
            Top             =   1500
            Width           =   3225
         End
         Begin VB.TextBox txtMaxStudentCount 
            Height          =   345
            Left            =   1590
            MaxLength       =   20
            TabIndex        =   10
            Text            =   "40"
            Top             =   1920
            Width           =   3225
         End
         Begin VB.TextBox txtMinGrade 
            Height          =   345
            Left            =   1590
            MaxLength       =   20
            TabIndex        =   9
            Text            =   "75"
            Top             =   2340
            Width           =   3225
         End
         Begin VB.TextBox txtMaxGrade 
            Height          =   345
            Left            =   1590
            MaxLength       =   20
            TabIndex        =   8
            Text            =   "100"
            Top             =   2760
            Width           =   3225
         End
         Begin VB.TextBox txtNote 
            Height          =   855
            Left            =   4980
            MaxLength       =   100
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   7
            Top             =   2655
            Width           =   4515
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Room"
            Height          =   195
            Left            =   150
            TabIndex        =   33
            Top             =   3240
            Width           =   420
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Section Offering ID"
            Height          =   195
            Left            =   150
            TabIndex        =   22
            Top             =   255
            Width           =   1350
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Section Title"
            Height          =   195
            Left            =   150
            TabIndex        =   21
            Top             =   690
            Width           =   885
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "School Year"
            Height          =   195
            Left            =   180
            TabIndex        =   20
            Top             =   1140
            Width           =   870
         End
         Begin VB.Label TeacherName 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "TeacherName"
            Height          =   195
            Left            =   150
            TabIndex        =   19
            Top             =   1560
            Width           =   1020
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Max. Student #"
            Height          =   195
            Left            =   180
            TabIndex        =   18
            Top             =   1980
            Width           =   1095
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Min. Grade"
            Height          =   195
            Left            =   180
            TabIndex        =   17
            Top             =   2400
            Width           =   780
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Max. Grade"
            Height          =   195
            Left            =   150
            TabIndex        =   16
            Top             =   2820
            Width           =   825
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Note:"
            Height          =   195
            Left            =   4980
            TabIndex        =   15
            Top             =   2415
            Width           =   390
         End
      End
      Begin HSES.b8Container b8cTabBg 
         Height          =   3660
         Index           =   1
         Left            =   -74940
         TabIndex        =   6
         Top             =   360
         Width           =   9675
         _ExtentX        =   17066
         _ExtentY        =   6456
         BorderColor     =   12307149
         BackColor       =   16185592
         ShadowColor1    =   13427430
         ShadowColor2    =   14215660
         Begin VB.CommandButton cmdRemoveAll 
            Height          =   330
            Left            =   2970
            Picture         =   "frmAddSectionOffering.frx":1660
            Style           =   1  'Graphical
            TabIndex        =   32
            Top             =   2130
            Width           =   525
         End
         Begin VB.CommandButton cmdRemoveOne 
            Height          =   330
            Left            =   2970
            Picture         =   "frmAddSectionOffering.frx":1BEA
            Style           =   1  'Graphical
            TabIndex        =   31
            Top             =   1740
            Width           =   525
         End
         Begin VB.CommandButton cmdAddOne 
            Height          =   330
            Left            =   2970
            Picture         =   "frmAddSectionOffering.frx":2174
            Style           =   1  'Graphical
            TabIndex        =   30
            Top             =   1170
            Width           =   525
         End
         Begin MSComctlLib.ListView listSubject 
            Height          =   3225
            Left            =   3540
            TabIndex        =   27
            Top             =   330
            Width           =   6015
            _ExtentX        =   10610
            _ExtentY        =   5689
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            HotTracking     =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            Appearance      =   1
            NumItems        =   7
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Title"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Subject ID"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Days"
               Object.Width           =   706
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "In"
               Object.Width           =   1411
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "Out"
               Object.Width           =   1411
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Text            =   "Teacher ID"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   6
               Text            =   "Teacher"
               Object.Width           =   5292
            EndProperty
         End
         Begin MSComctlLib.ListView listAvailableSubjects 
            Height          =   3225
            Left            =   90
            TabIndex        =   23
            Top             =   330
            Width           =   2865
            _ExtentX        =   5054
            _ExtentY        =   5689
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            Appearance      =   1
            NumItems        =   1
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Title"
               Object.Width           =   5054
            EndProperty
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Subjects Offered"
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
            Left            =   3570
            TabIndex        =   29
            Top             =   120
            Width           =   1230
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Available Subjects"
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
            TabIndex        =   28
            Top             =   120
            Width           =   1305
         End
      End
   End
   Begin VB.CommandButton cmdAddSubject 
      Appearance      =   0  'Flat
      BackColor       =   &H00D8E9EC&
      Height          =   285
      Left            =   9315
      Picture         =   "frmAddSectionOffering.frx":26FE
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   870
      Width           =   255
   End
   Begin HSES.b8Line b8Line1 
      Height          =   60
      Left            =   0
      TabIndex        =   0
      Top             =   510
      Width           =   10065
      _ExtentX        =   17754
      _ExtentY        =   106
   End
   Begin HSES.b8Line b8Line2 
      Height          =   60
      Left            =   15
      TabIndex        =   1
      Top             =   4725
      Width           =   10065
      _ExtentX        =   17754
      _ExtentY        =   106
   End
   Begin lvButton.lvButtons_H cmdSave 
      Default         =   -1  'True
      Height          =   360
      Left            =   8445
      TabIndex        =   34
      Top             =   4815
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   635
      Caption         =   "&Save"
      CapAlign        =   2
      BackStyle       =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cBhover         =   14215660
      Focus           =   0   'False
      cGradient       =   14215660
      Gradient        =   3
      Mode            =   0
      Value           =   0   'False
      cBack           =   16185592
   End
   Begin lvButton.lvButtons_H cmdCancel 
      Cancel          =   -1  'True
      Height          =   360
      Left            =   6915
      TabIndex        =   35
      Top             =   4815
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   635
      Caption         =   "&Cancel"
      CapAlign        =   2
      BackStyle       =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cBhover         =   14215660
      Focus           =   0   'False
      cGradient       =   14215660
      Gradient        =   3
      Mode            =   0
      Value           =   0   'False
      cBack           =   16185592
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Add New Section Offering"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   345
      Left            =   60
      TabIndex        =   2
      Top             =   150
      Width           =   3705
   End
   Begin VB.Image Image2 
      Height          =   525
      Left            =   0
      Picture         =   "frmAddSectionOffering.frx":2C88
      Stretch         =   -1  'True
      Top             =   0
      Width           =   10065
   End
End
Attribute VB_Name = "frmAddSectionOffering"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim RecordAdded As Boolean

Dim curSectionID As String
Dim curTeacherID As String

Dim listRoomID() As String

Public Function ShowForm(Optional sSectionFullTitle As String = "", Optional sSchoolYearTitle As String = "") As Boolean
    

    '-------------------------------------------------------
    'check user access
    '-------------------------------------------------------
    If UserAllowedTo(CurrentUser.UserName, sCanAddSectionOffering) = False Then
        MsgBox "Unable to continue adding Section Offering entry." & vbNewLine & vbNewLine & _
                "You are not permitted to do this. Please contact your administrator for more information.", vbExclamation
        Exit Function
    End If
    '-------------------------------------------------------



    'set parameters
    txtSectionFullTitle.Text = sSectionFullTitle
    If sSchoolYearTitle = "" Then
        txtSchoolYearTitle.Text = CurrentSchoolYear.SchoolYearTitle
    Else
        txtSchoolYearTitle.Text = sSchoolYearTitle
    End If
    
    'check user access
    If UserAllowedTo(CurrentUser.UserName, "Can Add Section Offering") = False Then
        MsgBox "Unable to show Add Section Offering window." & vbNewLine & _
                "You are not permitted to aceess it. Please contact your Administrator.", vbExclamation
        
        Unload Me
        Exit Function
    End If
    
    
    'show form
    Me.Show vbModal
    
    'return
    ShowForm = RecordAdded
End Function

Private Sub cmbRoom_LostFocus()
    If cmbRoom.ListIndex < 0 Then
        cmbRoom.ListIndex = 0
    End If
End Sub

Private Sub cmdAddOne_Click()
    
    Dim sTime() As String
    Dim lvItem As ListItem
    Dim i As Integer
    
    Dim sSubjectID As String
    Dim sSubjectOfferingID As String
    Dim sSchedTimeStart As String
    Dim sSchedTimeEnd As String
    Dim sTeacherID As String
    Dim sTeacherName As String
    Dim sDays As String
        
    'generate acquired time sched
    If listSubject.ListItems.Count < 1 Then
        ReDim sTime(0)
        sTime(0) = "?-0-0"
    Else
    
        ReDim sTime(listSubject.ListItems.Count)
        
        sTime(0) = "?-0-0"
        i = 1
        For Each lvItem In listSubject.ListItems
            sTime(i) = lvItem.SubItems(2) & "-" & lvItem.SubItems(3) & "-" & lvItem.SubItems(4)
            i = i + 1
        Next
        
    End If
    
    If listAvailableSubjects.ListItems.Count < 1 Then
        Exit Sub
    End If
    
    
    If Len(listAvailableSubjects.SelectedItem.Text) < 1 Then
        Exit Sub
    End If
    
    
    
    If frmAddToSubjectOffered.ShowForm(curSectionID, GetLVKey(listAvailableSubjects.SelectedItem), _
                                                listAvailableSubjects.SelectedItem.Text, sTime, _
                                                 sSubjectOfferingID, _
                                                 sSchedTimeStart, _
                                                 sSchedTimeEnd, _
                                                 sTeacherID, _
                                                 sDays, _
                                                 sTeacherName _
                                                ) = True Then
        
        listSubject.ListItems.Add , GetLVKey(listAvailableSubjects.SelectedItem), listAvailableSubjects.SelectedItem.Text
        listSubject.ListItems(listSubject.ListItems.Count).SubItems(1) = GetLVKey(listAvailableSubjects.SelectedItem)
        listSubject.ListItems(listSubject.ListItems.Count).SubItems(2) = sDays
        listSubject.ListItems(listSubject.ListItems.Count).SubItems(3) = sSchedTimeStart
        listSubject.ListItems(listSubject.ListItems.Count).SubItems(4) = sSchedTimeEnd
        listSubject.ListItems(listSubject.ListItems.Count).SubItems(5) = sTeacherID
        listSubject.ListItems(listSubject.ListItems.Count).SubItems(6) = sTeacherName
    
        listAvailableSubjects.ListItems.Remove listAvailableSubjects.SelectedItem.Index
    End If
                                                
End Sub

Private Sub cmdAddSubject_Click()
    
    Dim vSection As tSection
    Dim vDepartment As tDepartment
    Dim vYearlevel As tYearLevel
    Dim sDepartmentTitle As String
    Dim sYearLevelTitle As String
    
    
    sDepartmentTitle = ""
    sYearLevelTitle = ""
    
    If Len(txtSectionFullTitle.Text) > 0 Then
        If GetSectionByFullTitle(txtSectionFullTitle.Text, vSection) = Success Then
            If GetDepartmentByID(vSection.DepartmentID, vDepartment) = Success Then
                sDepartmentTitle = vDepartment.DepartmentTitle
            End If
            sYearLevelTitle = YLIDtoTitle(vSection.YearLevelID)
        End If
    End If
    
    If frmAddSubject.ShowForm(sDepartmentTitle, sYearLevelTitle) = True Then
        GenerateSubjects
        listSubject.ListItems(listSubject.ListItems.Count).Selected = True
        listSubject.ListItems(listSubject.ListItems.Count).EnsureVisible
        listSubject.SetFocus
    End If
    
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdGetSchoolYear_Click()
    Dim sSchoolYearTitle As String
    
    sSchoolYearTitle = PickSchoolYear.GetItem(txtSchoolYearTitle, , , True)
    
    If sSchoolYearTitle <> "" Then
        txtSchoolYearTitle.Text = sSchoolYearTitle
    End If
End Sub



Private Sub cmdGetSectionTitle_Click()
    Dim sSectionTitle As String
    Dim iYL As Integer
    
    curSectionID = PickSection.GetSectionID(txtSectionFullTitle, , , , sSectionTitle, iYL)
    
    If curSectionID <> "" Then
        txtSectionFullTitle.Text = YLIDtoTitle(iYL) & " - " & sSectionTitle
    End If
End Sub

Private Sub cmdGetTeacher_Click()
    Dim sTeacherID As String
    Dim sTeacherFullName As String
    
    sTeacherID = PickTeacher.GetTeacherID(sTeacherFullName)
    
    If sTeacherID <> "" Then
        curTeacherID = sTeacherID
        txtTeacherFullName.Text = sTeacherFullName
    End If
End Sub

Private Sub cmdRemoveAll_Click()
    
    Dim lvItem As ListItem
    
    For Each lvItem In listSubject.ListItems
        If Len(lvItem.Text) > 0 Then
            listAvailableSubjects.ListItems.Add , KeySubjectOffering & lvItem.SubItems(1), lvItem.Text
        End If
    Next
    
    listSubject.ListItems.Clear
End Sub

Private Sub cmdRemoveOne_Click()
    If Len(listSubject.SelectedItem.Text) < 1 Then
        Exit Sub
    End If
    
    listAvailableSubjects.ListItems.Add , KeySubjectOffering & listSubject.SelectedItem.SubItems(1), listSubject.SelectedItem.Text
    
    listSubject.ListItems.Remove listSubject.SelectedItem.Index
End Sub

Private Sub cmdSave_Click()
    Form_SaveData
End Sub





Private Sub Form_Activate()
    If RefreshRoomList = False Then
        MsgBox "There are no available Room to create Section Offering." & vbNewLine & _
            "Please add Room entry first.", vbExclamation
        Unload Me
    End If
End Sub

Private Function RefreshRoomList() As Boolean
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    'default
    RefreshRoomList = False
    
    sSQL = "SELECT tblRoom.RoomID, tblRoom.Room" & _
            " From tblRoom" & _
            " ORDER BY tblRoom.Room"
            
    If ConnectRS(HSESDB, vRS, sSQL) = False Then
        'temp
        'error
        CatchError "AddSectionOffering", "RefreshRommList", "Unable to connect Recordset with SQL Expression : '" & sSQL & "'"
        GoTo ReleaseAndExit
    End If
    
    If AnyRecordExisted(vRS) = False Then
        GoTo ReleaseAndExit
    End If
    
    ReDim listRoomID(getRecordCount(vRS) - 1)
    cmbRoom.Clear
    vRS.MoveFirst
    While vRS.EOF = False
        cmbRoom.AddItem ReadField(vRS("Room"))
        listRoomID(cmbRoom.ListCount - 1) = ReadField(vRS("RoomID"))
        vRS.MoveNext
    Wend
    cmbRoom.ListIndex = 0
    
        
    RefreshRoomList = True
    
ReleaseAndExit:
    Set vRS = Nothing
End Function

Private Sub tabMAin_Click(PreviousTab As Integer)
    
    Dim i As Integer
    
    For i = 0 To b8cTabBg.UBound
        If i <> tabMAin.Tab Then
            b8cTabBg(i).Visible = False
        End If
    Next
    
    b8cTabBg(tabMAin.Tab).Visible = True
End Sub

Private Sub txtMaxGrade_KeyPress(KeyAscii As Integer)
    If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8 Or KeyAscii = 45) Then KeyAscii = 0

End Sub

Private Sub txtMaxGrade_LostFocus()
    If Len(txtMaxGrade.Text) > 0 Then
        If IsNumeric(txtMaxGrade.Text) Then
            If Val(txtMaxGrade.Text) < 60 Or Val(txtMaxGrade.Text) > 100 Then
                MsgBox "Invalid Entry!" & vbNewLine & "Max. Grade must be range 60-100", vbExclamation
                HLTxt txtMaxGrade
            End If
        Else
            MsgBox "Invalid Entry!" & vbNewLine & "Max. Grade must be range 60-100", vbExclamation
            HLTxt txtMaxGrade
        End If
    End If
    
    CheckMinMaxGrade
End Sub

Private Sub txtMaxStudentCount_KeyPress(KeyAscii As Integer)
    If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8 Or KeyAscii = 45) Then KeyAscii = 0

End Sub

Private Sub txtMaxStudentCount_LostFocus()
    If Len(txtMaxStudentCount.Text) > 0 Then
        If IsNumeric(txtMaxStudentCount.Text) Then
            If Val(txtMaxStudentCount.Text) < 1 Or Val(txtMaxStudentCount.Text) > 100 Then
                MsgBox "Invalid Entry!" & vbNewLine & "Max. Student # must be range 1-100", vbExclamation
                HLTxt txtMaxStudentCount
            End If
        Else
            MsgBox "Invalid Entry!" & vbNewLine & "Max. Student # must be range 1-100", vbExclamation
            HLTxt txtMaxStudentCount
        End If
    End If
    
    
End Sub

Private Sub txtMinGrade_KeyPress(KeyAscii As Integer)
    If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8 Or KeyAscii = 45) Then KeyAscii = 0

End Sub

Private Sub txtMinGrade_LostFocus()
    If Len(txtMinGrade.Text) > 0 Then
        If IsNumeric(txtMinGrade.Text) Then
            If Val(txtMinGrade.Text) < 60 Or Val(txtMinGrade.Text) > 100 Then
                MsgBox "Invalid Entry!" & vbNewLine & "Min. Grade must be range 60-100", vbExclamation
                HLTxt txtMinGrade
            End If
        Else
            MsgBox "Invalid Entry!" & vbNewLine & "Min. Grade must be range 60-100", vbExclamation
            HLTxt txtMinGrade
        End If
    End If
    
    CheckMinMaxGrade
End Sub

Private Sub txtSchoolYearTitle_Change()
    GenerateSectionOfferingID
End Sub

Private Sub txtSectionFullTitle_Change()
    GenerateSectionOfferingID
End Sub

Private Sub GenerateSectionOfferingID()
    Dim vSection As tSection
    
    txtSectionOfferingID.Text = ""
    
    If Len(txtSectionFullTitle.Text) < 1 Or Len(txtSchoolYearTitle.Text) < 1 Then
        curSectionID = ""
        
        Exit Sub
        
    End If
    
    'delay 0.3 second
    'code by: VIncent J. Jamero
    '------------------------------------------------
    Static DelayStart As Single
    Static notFirst As Boolean
    DelayStart = GetTickCount + 300
    If notFirst = True Then Exit Sub
    notFirst = True
    While GetTickCount < DelayStart
        DoEvents
    Wend
    notFirst = False
    '------------------------------------------------
    'the next line will be if executed if user pause typing in 0.3 second


    If SectionExistByFullTitle(txtSectionFullTitle.Text) = Failed Or SchoolYearExistByTitle(txtSchoolYearTitle.Text) = Failed Then
        Exit Sub
    End If
    
    If GetSectionByFullTitle(txtSectionFullTitle.Text, vSection) <> Success Then
        Exit Sub
    End If
    
    'set section id
    curSectionID = vSection.SectionID

    'all prerequisits OK
    'generate ID
        
    txtSectionOfferingID.Text = VBA.Trim(txtSchoolYearTitle.Text) & "-" & vSection.SectionID

    If SectionOfferingExistByID(txtSectionOfferingID.Text) = Success Then
        MsgBox "This Section Offering Entry is already exist in record." & vbNewLine & "Please change Section or School Year.", vbExclamation
        HLTxt txtSectionFullTitle
    Else
        'generate subjects
        GenerateSubjects
    End If
End Sub

 
Private Sub GenerateSubjects()
    
    Dim vRS As ADODB.Recordset
    Dim sSQL As String
    Dim lvItem As ListItem
    
    sSQL = "SELECT tblSubject.SubjectID, tblSubject.SubjectTitle" & _
            " FROM tblDepartment INNER JOIN ((tblYearLevel INNER JOIN tblSection ON tblYearLevel.YearLevelID = tblSection.YearLevelID) INNER JOIN tblSubject ON tblYearLevel.YearLevelID = tblSubject.YearLevelID) ON (tblDepartment.DepartmentID = tblSubject.DepartmentID) AND (tblDepartment.DepartmentID = tblSection.DepartmentID)" & _
            " WHERE ((([tblYearLevel]![YearLevelTitle] & ' - ' & [tblSection]![SectionTitle])='" & txtSectionFullTitle.Text & "'));"

    'clear subject list
    listSubject.ListItems.Clear
    
    
    If ConnectRS(HSESDB, vRS, sSQL) = True Then
        If AnyRecordExisted(vRS) = True Then
            'fill all subjects to list
            FillRecordToList vRS, listAvailableSubjects, KeySubjectOffering

            
        Else
            'section
            MsgBox "There are no subjects available for this section. Please add subjects to to this section.", vbExclamation
        End If
    Else
        'fatal error
        CatchError "AddSectionOffering", "Generate Subjects", "Error in connecting Recordset."
    End If
    
    Set vRS = Nothing
End Sub
 
 
Private Sub CheckMinMaxGrade()
    If (Not IsNumeric(txtMaxGrade.Text)) Or (Not IsNumeric(txtMinGrade.Text)) Then
        Exit Sub
    End If
    
    If Val(txtMaxGrade.Text) < Val(txtMinGrade.Text) Then
        MsgBox "Min. Grade mus be LESS THAN or EQUAL to Max. Grade.", vbExclamation
        HLTxt txtMaxGrade
    End If
    
End Sub












Private Sub Form_SaveData()

    Dim newSectionOffering As tSectionOffering
    Dim vSection As tSection
    Dim vTeacher As tTeacher
    
    Dim lvItem As ListItem
    
    Dim i As Integer
    Dim newSubjectOffering As tSubjectOffering
    Dim ErrMSG As String
    
    
    If ValidateData = False Then Exit Sub
    
    If GetSectionByFullTitle(txtSectionFullTitle.Text, vSection) <> Success Then
        MsgBox "Invalid Section entry!" & vbNewLine & "Please Enter valid Section Full TItle.", vbExclamation
        HLTxt txtSectionFullTitle
        Exit Sub
    End If
    
    If Len(curTeacherID) < 1 Then
        MsgBox "Invalid Teahcer entry!" & vbNewLine & "Please Enter valid Teacher Full Name.", vbExclamation
        cmdGetTeacher.SetFocus
        Exit Sub
    Else
        If TeacherAssignedBySchoolYear(curTeacherID, txtSchoolYearTitle.Text) = Success Then
            MsgBox "The seleted Teacher entry is already assigned in the selected School Year." & vbNewLine & "Please select other Teacher entry", vbExclamation
            cmdGetTeacher.SetFocus
            Exit Sub
        End If
    End If
    
    'subjects
    If listSubject.ListItems.Count < 1 Then
        MsgBox "There are subjects that are not added in this Section Offering.", vbExclamation
                
        Exit Sub
        
    End If
    
    newSectionOffering.SectionOfferingID = txtSectionOfferingID.Text
    newSectionOffering.SectionID = vSection.SectionID
    newSectionOffering.SchoolYear = txtSchoolYearTitle.Text
    newSectionOffering.TeacherID = curTeacherID
    newSectionOffering.MaxStudentCount = Val(txtMaxStudentCount.Text)
    newSectionOffering.MaxGrade = Val(txtMaxGrade.Text)
    newSectionOffering.MinGrade = Val(txtMinGrade.Text)
    newSectionOffering.Note = txtNote.Text
    newSectionOffering.RoomID = listRoomID(cmbRoom.ListIndex)
    
    
    newSectionOffering.CreationDate = Now
    newSectionOffering.CreatedBy = CurrentUser.UserName
    
    Select Case AddSectionOffering(newSectionOffering)
        Case TranDBResult.Success
            'success
            'add subjects offering
            
            ErrMSG = ""
            
            For Each lvItem In listSubject.ListItems

                newSubjectOffering.SubjectOfferingID = newSectionOffering.SectionOfferingID & "-" & lvItem.Key
                newSubjectOffering.SectionOfferingID = newSectionOffering.SectionOfferingID
                newSubjectOffering.SubjectID = lvItem.SubItems(1)
                
                newSubjectOffering.TeacherID = lvItem.SubItems(5)
                newSubjectOffering.Days = lvItem.SubItems(2)
                newSubjectOffering.SchedTimeStart = lvItem.SubItems(3)
                newSubjectOffering.SchedTimeEnd = lvItem.SubItems(4)
                
                newSubjectOffering.CreationDate = Now
                newSubjectOffering.CreatedBy = CurrentUser.UserName
                                
                If AddSubjectOffering(newSubjectOffering) <> TranDBResult.Success Then
                    ErrMSG = vbNewLine & ErrMSG & "Error Adding Subject [ID: " & newSubjectOffering.SubjectID & "]  To Section [ID : " & newSectionOffering.SectionOfferingID & "]"
                End If
                
                'fatal error
                
                
            Next
            
            If ErrMSG <> "" Then
                'error found
                'just ignore
                MsgBox "FATAL ERROR:" & ErrMSG, vbCritical
            End If
            
            MsgBox "New Section Offering entry successfully added.", vbInformation
            RecordAdded = True
            Unload Me
        Case TranDBResult.DuplicateID
            MsgBox "This Section Offering Entry is already exist in record." & vbNewLine & "Please change Section or School Year.", vbExclamation
            HLTxt txtSectionFullTitle
        Case Else
            CatchError "frmAddsection", "Form_savedata", "AddSection Unknown result"
    End Select
End Sub

Private Function ValidateData() As Boolean

    Dim sSubjects() As String

    'default
    ValidateData = False
    
    If Not CheckTextBox(txtSectionOfferingID, "Please Enter valid Section Title and School Year to generate Section Offering ID.") Then
        Exit Function
    End If
    
    If SectionExistByFullTitle(txtSectionFullTitle.Text) <> Success Then
        MsgBox "Please enter valid Section Title", vbExclamation
        HLTxt txtSectionFullTitle
        Exit Function
    End If
    
    If SchoolYearExistByTitle(txtSchoolYearTitle.Text) <> Success Then
        MsgBox "Please enter valid School Year Title", vbExclamation
        HLTxt txtSchoolYearTitle
        Exit Function
    End If
    
    If SectionOfferingExistByID(txtSectionOfferingID.Text) = Success Then
        MsgBox "This Section Offering Entry is already exist in record." & vbNewLine & "Please change Section or School Year.", vbExclamation
        HLTxt txtSectionFullTitle
        Exit Function
    End If
    
    If Len(curTeacherID) < 1 Then
        MsgBox "Please enter valid Teacher Full Name", vbExclamation
        cmdGetTeacher.SetFocus
        Exit Function
    End If
    
    'check room
    If RoomExistBySY(listRoomID(cmbRoom.ListIndex), txtSchoolYearTitle.Text) = True Then
        If MsgBox("The selected Room are already used by another Section with in the selected School Year." & vbNewLine & _
            "Do you want to ignore this?", vbQuestion + vbOKCancel) = vbCancel Then
            'cmbRoom.SetFocus
            Exit Function
        End If
    End If
    
    
    
    
    'Max student count
    If IsNumeric(txtMaxStudentCount.Text) Then
        If Val(txtMaxStudentCount.Text) < 1 Or Val(txtMaxStudentCount.Text) > 100 Then
            MsgBox "Invalid Entry!" & vbNewLine & "Max. Student # must be range 1-100", vbExclamation
            HLTxt txtMaxStudentCount
            Exit Function
        End If
    Else
        MsgBox "Invalid Entry!" & vbNewLine & "Max. Student # must be range 1-100", vbExclamation
        HLTxt txtMaxStudentCount
        Exit Function
    End If
    
    If IsNumeric(txtMinGrade.Text) Then
        If Val(txtMinGrade.Text) < 60 Or Val(txtMinGrade.Text) > 100 Then
            MsgBox "Invalid Entry!" & vbNewLine & "Min. Grade must be range 60-100", vbExclamation
            HLTxt txtMinGrade
            Exit Function
        End If
    Else
        MsgBox "Invalid Entry!" & vbNewLine & "Min. Grade must be range 60-100", vbExclamation
        HLTxt txtMinGrade
        Exit Function
    End If
    
    
    If IsNumeric(txtMaxGrade.Text) Then
        If Val(txtMaxGrade.Text) < 60 Or Val(txtMaxGrade.Text) > 100 Then
            MsgBox "Invalid Entry!" & vbNewLine & "Max. Grade must be range 60-100", vbExclamation
            HLTxt txtMaxGrade
            Exit Function
        End If
    Else
        MsgBox "Invalid Entry!" & vbNewLine & "Max. Grade must be range 60-100", vbExclamation
        HLTxt txtMaxGrade
        Exit Function
    End If
    
    If Val(txtMaxGrade.Text) < Val(txtMinGrade.Text) Then
        MsgBox "Min. Grade mus be LESS THAN or EQUAL to Max. Grade.", vbExclamation
        HLTxt txtMaxGrade
        Exit Function
    End If
    

    
    'return success
    ValidateData = True
End Function





Private Sub txtTeacherFullName_Change()

    If Len(txtTeacherFullName.Text) < 1 Then Exit Sub

    
    'delay 0.3 second
    'code by: VIncent J. Jamero
    '------------------------------------------------
    Static DelayStart As Single
    Static notFirst As Boolean
    DelayStart = GetTickCount + 300
    If notFirst = True Then Exit Sub
    notFirst = True
    While GetTickCount < DelayStart
        DoEvents
    Wend
    notFirst = False
    '------------------------------------------------
    'the next line will be if executed if user pause typing in 0.3 second

        
        If TeacherAssignedBySchoolYear(curTeacherID, txtSchoolYearTitle.Text) = Success Then
            MsgBox "The seleted Teacher entry is already assigned in the selected School Year." & vbNewLine & "Please select other Teacher entry", vbExclamation
            cmdGetTeacher.SetFocus
        End If

    
End Sub


Private Function RoomExistBySY(sRoomID As String, sSchoolYear As String) As Boolean
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    'default
    RoomExistBySY = False
    
    sSQL = "SELECT tblSectionOffering.SchoolYear, tblSectionOffering.RoomID" & _
            " FROM tblSectionOffering" & _
            " WHERE (((tblSectionOffering.SchoolYear)='" & sSchoolYear & "') AND ((tblSectionOffering.RoomID)='" & sRoomID & "'));"

            
    If ConnectRS(HSESDB, vRS, sSQL) = False Then
        'temp
        'error
        CatchError "AddSectionOffering", "RoomExistBySY", "Unable to connect Recordset with SQL Expression : '" & sSQL & "'"
        GoTo ReleaseAndExit
    End If
    
    If AnyRecordExisted(vRS) = False Then
        GoTo ReleaseAndExit
    End If
    
    RoomExistBySY = True
    
ReleaseAndExit:
    Set vRS = Nothing
End Function
