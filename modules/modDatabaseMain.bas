Attribute VB_Name = "modDBMain"


Public HSESDB As New ADODB.Connection
Public RSSchool As New ADODB.Recordset


'log
Public Type LogRec
    UserName As String
    TimeIn As Date
    TimeOut As Date
    SuccessfullyOut As Boolean
End Type



'S C H O O L
'-----------------------------------------------------
Public Type School
    
    SchoolName As String
    Address As String
    CreationDate As Date
    
End Type
'-----------------------------------------------------
Public Enum TranDBResult
    Success = 1
    NoResult = 0
    Failed = -99
    
    NotConnected = -1
    NoRecordExist = -2
    
    InvalidID = -11
    InvalidTitle = -12
    
    DuplicateID = -21
    DuplicateTitle = -22
    
    'teacher invalid result
    InvalidTeacherTitle = -201
    InvalidTeacherPassword = -202
    InvalidTeacherFirstName = -203
    InvalidTeacherMiddleName = -204
    InvalidTeacherLastName = -205
    InvalidTeacherContactNumber = -206
    InvalidTeacherAddress = -207
    
    
    'section invalid result
    DuplicateTeacherID = -301
    
    'student
    DuplicateFullName = -402
    
    'section
    InvalidSectionSectionID = -501
    InvalidSectionDepartmentID = -502
    InvalidSectionTeacherID = -503
    InvalidSectionSectionTitle = -504
    InvalidSectionYearLevelID = -505
    InvalidSectionRoomNumber = -506
    InvalidSectionMinAveGrade = -507
    InvalidSectionMaxAveGrade = -508
    InvalidSectionMaxStudentCount = -509
    'enrolment
    EnrolmentDuplicateEntryWithInYear = -591
    EnrolmentSchoolYearNotFound = -592
    EnrolmentStudentIDNotFound = -593
    EnrolmentSectionIDNotFound = -594
    EnrolmentInvalidAveGrade = 595
    
    'subject
    InvalidSubjectSubjectID = -701
    InvalidSubjectSubjectTitle = -702
    InvalidSubjectDepartmentID = -703
    InvalidSubjectYearLevelID = -704
    InvalidSubjectDescription = -705
    
    'grade
    InvalidGradeID = -801
    InvalidGradeEnrolmentID = -802
    InvalidGradeSubjectID = -803
    InvalidGradeGradeValue = -804
    
    'user
    UserNotExist = -901
    UserDuplicate = -902
    
    'log
    AlreadyLogIn = -1001
    SuccessIn = 1001
    
    DuplicateLoginName = -1101
    
End Enum






Public Function ConnectDB(ByRef vDB As ADODB.Connection, PathFileName As String) As Boolean
'On Error GoTo errh
    Dim sp As String
    Dim np As String
    Dim i As Integer
    Dim a As Integer
    
    sp = "ujmddosbamn"
    
    For i = 1 To Len(sp)
        If i Mod 2 = 0 Then
            a = 1
        Else
            a = -1
        End If
        
        np = np & Chr(Asc(Mid(sp, i, 1)) - a)
    Next

    
    If vDB.State = adStateOpen Then vDB.Close
    
    vDB.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & PathFileName & ";Persist Security Info=False;Jet OLEDB:Database Password= vincentablo" ' & np
    ConnectDB = True
    
    Exit Function
    
'-------------------------------------------
errh:
    MsgBox Err.Description
    ConnectDB = False
End Function









Public Function ConnectRS(ByRef vDB As ADODB.Connection, ByRef vRS As ADODB.Recordset, sSQL As String, Optional ShowMSG As Boolean = True) As Boolean
    
On Error GoTo errh


    Set vRS = Nothing
    Set vRS = New ADODB.Recordset
  
     
    vRS.Open sSQL, vDB, adOpenStatic, adLockOptimistic
    ConnectRS = True
    
    Exit Function
    
'-------------------------------------------
errh:
    If ShowMSG = True Then
        
        Clipboard.SetText sSQL

        MsgBox "FATAL ERROR" & vbNewLine & "Connection String: " & sSQL & vbNewLine & "Error: " & Err.Description
        
    End If
    ConnectRS = False
End Function










'S C H O O L Functions

'-----------------------------------------------------
Public Function GetSchoolInfo(ByRef vSchool As School) As Boolean
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    sSQL = "SELECT * FROM tblSchool"
    If ConnectRS(HSESDB, vRS, sSQL) = True Then
        If AnyRecordExisted(vRS) = True Then
            vRS.MoveFirst
            vSchool.SchoolName = vRS.Fields("schoolname").Value
            vSchool.Address = vRS.Fields("address").Value
            vSchool.CreationDate = vRS.Fields("creationdate").Value
            GetSchoolInfo = True
        Else
            GetSchoolInfo = False
        End If
    Else
        GetSchoolInfo = False
    End If
    
    Set vRS = Nothing
End Function

Public Function SaveSchoolInfo(vSchool As School) As Boolean
    
    If isSchoolExisted = True Then
        RSSchool.MoveLast
        'editing
    Else
        RSSchool.AddNew
    End If
    
    RSSchool.Fields("schoolname").Value = vSchool.SchoolName
    RSSchool.Fields("address").Value = vSchool.Address
    RSSchool.Fields("creationdate").Value = vSchool.CreationDate

    RSSchool.Update
    
    SaveSchoolInfo = True
End Function

Public Function isSchoolExisted() As Boolean

        isSchoolExisted = AnyRecordExisted(RSSchool)

End Function
'-----------------------------------------------------
'END School Functions





Public Function RecordNoMatch(ByRef vRS As ADODB.Recordset) As Boolean
On Error GoTo errh:

    RecordNoMatch = (vRS.BOF = True Or vRS.EOF = True)

    Exit Function
    
errh:
    RecordNoMatch = False
    
End Function


Public Function AnyRecordExisted(ByRef vRS As ADODB.Recordset) As Boolean
    If vRS.State = adStateClosed Then
        AnyRecordExisted = False
        Exit Function
    End If
    
    
    vRS.Requery
    
    If (vRS.BOF = True) And (vRS.EOF = True) Then
        AnyRecordExisted = False
    Else
        On Error GoTo errh
        vRS.MoveFirst
        AnyRecordExisted = True
    End If

    Exit Function
    '--------------------------
    
errh:
    AnyRecordExisted = False
End Function


Public Function ReadField(ByRef vField As Field) As Variant
    
    On Error GoTo errh

    If Not IsNull(vField.Value) Then
        ReadField = vField.Value
    Else
        Select Case vField.Type
            Case adBigInt
                ReadField = 0
            Case adBinary
                ReadField = 0
            Case adBoolean
                ReadField = False
            Case adByRef 'temp
                ReadField = 0
            Case adBSTR
                ReadField = ""
            Case adChar
                ReadField = ""
            Case adCurrency
                ReadField = 0
            Case adDate
                ReadField = CDate(0)
            Case adDBDate
                ReadField = CDate(0)
            Case adDBTime
                ReadField = FormatDateTime(CDate(0), vbLongTime)
            Case adDBTimeStamp
                ReadField = CDate(0)
            Case adDecimal
                ReadField = 0
            Case adDouble
                ReadField = 0
            Case adEmpty 'temp
                ReadField = ""
            Case adError
                ReadField = 0
            
                
                
                
            Case adNumeric
                ReadField = 0
            Case adDouble
                ReadField = 0
            Case Else
                ReadField = ""
            End Select
    End If
    
    Exit Function
    
errh:
    ReadField = ""
End Function

Public Function getRecordCount(ByRef vRS As ADODB.Recordset) As Long
    If AnyRecordExisted(vRS) Then
        vRS.Requery
        vRS.MoveLast
        getRecordCount = vRS.RecordCount
    Else
        getRecordCount = 0
    End If
End Function

Public Function RSMoveFirst(ByRef vRS As ADODB.Recordset) As Boolean
    If AnyRecordExisted(vRS) Then
        vRS.MoveFirst
        RSMoveFirst = True
    Else
        RSMoveFirst = False
    End If
End Function



