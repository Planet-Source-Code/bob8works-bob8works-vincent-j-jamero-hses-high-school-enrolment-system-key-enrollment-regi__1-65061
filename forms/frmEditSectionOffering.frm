VERSION 5.00
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmEditSectionOffering 
   BackColor       =   &H00D8E9EC&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Section"
   ClientHeight    =   4755
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8880
   Icon            =   "frmEditSectionOffering.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   317
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   592
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cmbRoom 
      Height          =   315
      Left            =   1590
      TabIndex        =   25
      Top             =   3720
      Width           =   3210
   End
   Begin VB.TextBox txtNote 
      Height          =   855
      Left            =   4920
      MaxLength       =   100
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   13
      Top             =   3225
      Width           =   3795
   End
   Begin VB.TextBox txtMaxGrade 
      Height          =   345
      Left            =   1590
      MaxLength       =   20
      TabIndex        =   12
      Top             =   3300
      Width           =   3225
   End
   Begin VB.TextBox txtMinGrade 
      Height          =   345
      Left            =   1590
      MaxLength       =   20
      TabIndex        =   11
      Top             =   2880
      Width           =   3225
   End
   Begin VB.TextBox txtMaxStudentCount 
      Height          =   345
      Left            =   1590
      MaxLength       =   20
      TabIndex        =   10
      Top             =   2460
      Width           =   3225
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
      TabIndex        =   6
      Top             =   780
      Width           =   3225
   End
   Begin VB.CommandButton cmdGetSchoolYear 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   4440
      Picture         =   "frmEditSectionOffering.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1650
      Width           =   345
   End
   Begin VB.CommandButton cmdGetSectionTitle 
      BackColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   4470
      Picture         =   "frmEditSectionOffering.frx":0B14
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1260
      Width           =   345
   End
   Begin VB.CommandButton cmdGetTeacher 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   4440
      Picture         =   "frmEditSectionOffering.frx":109E
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2070
      Width           =   345
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
      Left            =   0
      TabIndex        =   1
      Top             =   4260
      Width           =   10065
      _ExtentX        =   17754
      _ExtentY        =   106
   End
   Begin lvButton.lvButtons_H cmdSave 
      Default         =   -1  'True
      Height          =   360
      Left            =   7380
      TabIndex        =   22
      Top             =   4350
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
      Left            =   5850
      TabIndex        =   23
      Top             =   4350
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
   Begin VB.TextBox txtTeacherFullName 
      Height          =   345
      Left            =   1590
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   9
      Top             =   2040
      Width           =   3225
   End
   Begin VB.TextBox txtSchoolYearTitle 
      Height          =   345
      Left            =   1590
      MaxLength       =   50
      TabIndex        =   8
      Top             =   1620
      Width           =   3225
   End
   Begin VB.TextBox txtSectionFullTitle 
      Height          =   330
      Left            =   1590
      MaxLength       =   50
      TabIndex        =   7
      Top             =   1230
      Width           =   3225
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Room"
      Height          =   195
      Left            =   150
      TabIndex        =   24
      Top             =   3780
      Width           =   420
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Note:"
      Height          =   195
      Left            =   4920
      TabIndex        =   21
      Top             =   2985
      Width           =   390
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Max. Grade"
      Height          =   195
      Left            =   150
      TabIndex        =   20
      Top             =   3360
      Width           =   825
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Min. Grade"
      Height          =   195
      Left            =   180
      TabIndex        =   19
      Top             =   2940
      Width           =   780
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Max. Student #"
      Height          =   195
      Left            =   180
      TabIndex        =   18
      Top             =   2520
      Width           =   1095
   End
   Begin VB.Label TeacherName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TeacherName"
      Height          =   195
      Left            =   150
      TabIndex        =   17
      Top             =   2100
      Width           =   1020
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "School Year"
      Height          =   195
      Left            =   180
      TabIndex        =   16
      Top             =   1680
      Width           =   870
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Section Title"
      Height          =   195
      Left            =   150
      TabIndex        =   15
      Top             =   1230
      Width           =   885
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Section Offering ID"
      Height          =   195
      Left            =   150
      TabIndex        =   14
      Top             =   795
      Width           =   1350
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Edit Section Offering"
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
      Left            =   90
      TabIndex        =   2
      Top             =   150
      Width           =   2985
   End
   Begin VB.Image Image2 
      Height          =   525
      Left            =   0
      Picture         =   "frmEditSectionOffering.frx":1628
      Stretch         =   -1  'True
      Top             =   0
      Width           =   10065
   End
End
Attribute VB_Name = "frmEditSectionOffering"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim RecordEdited As Boolean

Dim curSectionOffering As tSectionOffering

Dim curSectionEnrolmentCount As Long

Dim curTeacher As tTeacher

Dim listRoomID() As String

Public Function ShowForm(Optional sSectionOfferingID As String = "") As Boolean

    '-------------------------------------------------------
    'check user access
    '-------------------------------------------------------
    If UserAllowedTo(CurrentUser.UserName, sCanEditSectionOffering) = False Then
        MsgBox "Unable to continue Editing Section Offering entry." & vbNewLine & vbNewLine & _
                "You are not permitted to do this. Please contact your administrator for more information.", vbExclamation
        Exit Function
    End If
    '-------------------------------------------------------


    txtSectionOfferingID.Text = sSectionOfferingID
    
    'show form
    Me.Show vbModal
    
    'return
    ShowForm = RecordEdited
End Function

Private Sub ShowSectionOfferingDetails()

    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    sSQL = " SELECT tblSectionOffering.SectionOfferingID, [tblYearLevel]![YearLevelTitle] & ' - ' & [tblSection]![SectionTitle] AS txtSectionFullTitle, [tblTeacher]![LastName] & ', ' & [tblTeacher]![FirstName] & ' ' & [tblTeacher]![MiddleName] AS TeacherFullName,tblSectionOffering.TeacherID" & _
            " FROM tblYearLevel INNER JOIN (tblTeacher INNER JOIN (tblSection INNER JOIN tblSectionOffering ON tblSection.SectionID = tblSectionOffering.SectionID) ON tblTeacher.TeacherID = tblSectionOffering.TeacherID) ON tblYearLevel.YearLevelID = tblSection.YearLevelID" & _
            " WHERE tblSectionOffering.SectionOfferingID = '" & txtSectionOfferingID.Text & "'" & _
            " GROUP BY tblSectionOffering.SectionOfferingID, [tblYearLevel]![YearLevelTitle] & ' - ' & [tblSection]![SectionTitle], [tblTeacher]![LastName] & ', ' & [tblTeacher]![FirstName] & ' ' & [tblTeacher]![MiddleName],tblSectionOffering.TeacherID;"

    If GetSectionOfferingByID(txtSectionOfferingID.Text, curSectionOffering) = Success Then
    
        If ConnectRS(HSESDB, vRS, sSQL) = True Then
            If AnyRecordExisted(vRS) = True Then
            
                
                txtSectionFullTitle.Text = ReadField(vRS.Fields("txtSectionFullTitle"))
                
                txtTeacherFullName.Text = ReadField(vRS.Fields("TeacherFullName"))
                curTeacher.TeacherID = ReadField(vRS.Fields("TeacherID"))
                
                
                txtSchoolYearTitle.Text = curSectionOffering.SchoolYear
                
                txtMaxStudentCount.Text = curSectionOffering.MaxStudentCount
                txtMinGrade.Text = curSectionOffering.MinGrade
                
                txtMaxGrade.Text = curSectionOffering.MaxGrade
                
                txtNote.Text = curSectionOffering.Note

                
                DoEvents
                
       
                
                'get student count
                curSectionEnrolmentCount = -1
                GetEnrolmentCountBySectionOfferingID txtSectionOfferingID.Text, curSectionEnrolmentCount
                
                
            Else
                'record not existed
            End If
        Else
            'error in sql string
        End If
        
    End If
    
    Set vRS = Nothing
End Sub



Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdGetSchoolYear_Click()
    Dim sSchoolYearTitle As String
    
    sSchoolYearTitle = PickSchoolYear.GetItem(txtSchoolYearTitle)
    
    If sSchoolYearTitle <> "" Then
        txtSchoolYearTitle.Text = sSchoolYearTitle
    End If
End Sub



Private Sub cmdGetSectionTitle_Click()
    Dim sSectionTitle As String

    PickSection.GetSectionID , , , , sSectionTitle
    
    If sSectionTitle <> "" Then
        txtSectionFullTitle.Text = sSectionTitle
    End If
End Sub

Private Sub cmdGetTeacher_Click()
    Dim sTeacherID As String
    Dim sTeacherFullName As String
    
    sTeacherID = PickTeacher.GetTeacherID(sTeacherFullName)
    
    If sTeacherID <> "" Then
        curTeacher.TeacherID = sTeacherID
        txtTeacherFullName.Text = sTeacherFullName
    End If
End Sub

Private Sub cmdSave_Click()
    Form_SaveData
End Sub



Private Sub Form_Activate()
    Dim i As Integer
    
    If RefreshRoomList = False Then
        MsgBox "There are no available Room to create Section Offering." & vbNewLine & _
            "Please add Room entry first.", vbExclamation
        Unload Me
    End If
    
    For i = 0 To UBound(listRoomID)
        If listRoomID(i) = curSectionOffering.RoomID Then
            cmbRoom.ListIndex = i
            
            Exit For
        End If
    Next
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
    
    
    
    
        
    RefreshRoomList = True
    
ReleaseAndExit:
    Set vRS = Nothing
End Function

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
    
    
    If IsNumeric(txtMaxStudentCount.Text) Then
        If Val(txtMaxStudentCount.Text) < curSectionEnrolmentCount Then
            MsgBox "Invalid Max. Student Allowed value. It must be greater or equal " & curSectionEnrolmentCount & ".", vbExclamation
        
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


    Dim vSection As tSection

    Dim i As Integer
    Dim ErrMSG As String
    
    
    If ValidateData = False Then Exit Sub
    

        If curSectionOffering.TeacherID <> curTeacher.TeacherID Then
            If TeacherAssignedBySchoolYear(curTeacher.TeacherID, txtSchoolYearTitle.Text) = Success Then
                MsgBox "The seleted Teacher entry is already assigned in the selected School Year." & vbNewLine & "Please select other Teacher entry", vbExclamation
                HLTxt txtTeacherFullName
                Exit Sub
            End If
        End If

    
    curSectionOffering.SectionOfferingID = txtSectionOfferingID.Text
    'curSectionOffering.SectionID = vSection.SectionID
    'curSectionOffering.SchoolYear = txtSchoolYearTitle.Text
    curSectionOffering.TeacherID = curTeacher.TeacherID
    curSectionOffering.MaxStudentCount = Val(txtMaxStudentCount.Text)
    curSectionOffering.MaxGrade = Val(txtMaxGrade.Text)
    curSectionOffering.MinGrade = Val(txtMinGrade.Text)
    curSectionOffering.Note = txtNote.Text
    curSectionOffering.RoomID = listRoomID(cmbRoom.ListIndex)
    
    curSectionOffering.ModifiedDate = Now
    curSectionOffering.ModifiedBy = CurrentUser.UserName

    Select Case EditSectionOffering(curSectionOffering)
        Case TranDBResult.Success
            '----------------------------------------------
            ' S U C C E S S
            '----------------------------------------------
                MsgBox "Section Offering entry successfully edited.", vbInformation
                'set flag
                RecordEdited = True
                'close this form
                Unload Me
        Case TranDBResult.InvalidID
            MsgBox "Unable to update this entry." & vbNewLine & "The selected Section Offering entry not found in record.", vbExclamation
        Case Else
            MsgBox "Unknown error.", vbCritical
            
            CatchError "frmEditSectionOffering", "Form_SaveData", "Edit Section Offering return unknown result."
        
    End Select
End Sub

Private Function ValidateData() As Boolean

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
    
    If SectionOfferingExistByID(txtSectionOfferingID.Text) <> Success Then
        MsgBox "This Section Offering Entry is not exist in record." & vbNewLine & "Please change Section or School Year.", vbExclamation
        HLTxt txtSectionFullTitle
        Exit Function
        Exit Function
    End If
    

    
    
    
    'Max student count
    If IsNumeric(txtMaxStudentCount.Text) Then
        If Val(txtMaxStudentCount.Text) < 1 Or Val(txtMaxStudentCount.Text) > 100 Then
            MsgBox "Invalid Entry!" & vbNewLine & "Max. Student # must be range 1-100", vbExclamation
            HLTxt txtMaxStudentCount
            Exit Function
        End If
        
        If Val(txtMaxStudentCount.Text) < curSectionEnrolmentCount Then
            MsgBox "Invalid Max. Student Allowed value. It must be greater or equal " & curSectionEnrolmentCount & ".", vbExclamation
        
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



Private Sub txtSectionOfferingID_Change()

    If Len(txtSectionOfferingID.Text) < 1 Then
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


    ShowSectionOfferingDetails
End Sub

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

    
        If curSectionOffering.TeacherID <> curTeacher.TeacherID Then
            If TeacherAssignedBySchoolYear(curTeacher.TeacherID, txtSchoolYearTitle.Text) = Success Then
                MsgBox "The seleted Teacher entry is already assigned in the selected School Year." & vbNewLine & "Please select other Teacher entry", vbExclamation
                HLTxt txtTeacherFullName
            End If
        End If

    
End Sub
