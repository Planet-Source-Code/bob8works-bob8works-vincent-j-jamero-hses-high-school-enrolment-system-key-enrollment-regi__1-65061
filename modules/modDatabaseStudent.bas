Attribute VB_Name = "modRSStudent"
Option Explicit

Public Const KeyStudent = "stud"

Public Type tStudent
    StudentID As String
    FirstName As String
    MiddleName As String
    LastName As String
    
    CityAddress As String
    HomeAddress As String
    BirthDate As Date
    PlaceOfBirth As String
    Gender As String
    Status As String
    Citizenship As String
    BloodType As String
    Religion As String
    
    
    LastSchoolName As String
    LastSchoolContactNumber As String
    LastSchoolAddress As String
       
    'parents
    MotherName As String
    MotherOccupation As String
    FatherName As String
    FatherOccupation As String
    ParentsContactNumber As String
    ParentsAddress As String
    
    GuardianName As String
    GuardianAddress As String
    GuardianContactNumber As String
    
    OldAveGrade As Double
    
    Transferee As Boolean
    TransfereeYL As Integer
    
    CreationDate As Date
    CreatedBy As String
    ModifiedDate As Date
    ModifiedBy As String

End Type




Public Function GetNewStudentID() As String
    Dim sYear As String
    Dim sStudentNumber As String
    Dim sLastStudentNumber As String
    Dim sStudentID As String
    
    Dim sUserName As String
    
    Dim QRYStudentNewID As New ADODB.Recordset

    
        If ConnectRS(HSESDB, QRYStudentNewID, "SELECT CStr(Year(Now()))+'-'+Left('00000000',7-Len(CStr(Max(Val(Right([tblStudent]![StudentID],7)))+1)))+CStr(Max(Val(Right([tblStudent]![StudentID],7)))+1) AS maxId FROM tblStudent;") = True Then
            If AnyRecordExisted(QRYStudentNewID) Then
                sStudentID = QRYStudentNewID.Fields(0).Value
                If Len(sStudentID) < 1 Then
                    sYear = CStr(Year(Now))
                    sStudentID = Left(sYear, 4) & "-0000001"
                End If
            Else
                'set year
                sYear = CStr(Year(Now))
            
                sStudentID = Left(sYear, 4) & "-0000001"
            End If
        Else
        
            'set year
            sYear = CStr(Year(Now))
            
            sStudentID = Left(sYear, 4) & "-0000001"
        End If
    
        While StudentIDNotExistFromOther(sStudentID, sUserName) = True
            'set year
            sYear = CStr(Year(Now))
            sStudentNumber = Trim(Val(Right(Trim(sStudentID), 7)) + 1)
            sStudentNumber = Left("0000000", 7 - Len(sStudentNumber)) & sStudentNumber
            sStudentID = Left(sYear, 4) & "-" & sStudentNumber
        Wend
        
        'save user id
        SaveUserStudentID CurrentUser.UserName, sStudentID
    
        GetNewStudentID = sStudentID
End Function

Public Function DeleteUserStudentID(sUserName As String) As Boolean
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    'default
    DeleteUserStudentID = False
    '
    sSQL = "DELETE tftStudentID.UserName, tftStudentID.StudentID" & _
            " From tftStudentID" & _
            " WHERE (((tftStudentID.UserName)='" & sUserName & "'))"
    
    If ConnectRS(HSESDB, vRS, sSQL) = False Then
        'fatal error
        CatchError "modRSStudent", "DeleteUserStudentID", "Unable to connect Recordset with SQL Expression '" & sSQL & "'"
        GoTo ReleaseAndExit
    End If

   DeleteUserStudentID = True
    
    
ReleaseAndExit:
    Set vRS = Nothing
End Function


Private Function SaveUserStudentID(sUserName As String, sStudentID As String) As Boolean
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    'default
    SaveUserStudentID = False
    '
    sSQL = "SELECT tftStudentID.UserName, tftStudentID.StudentID" & _
            " From tftStudentID" & _
            " WHERE (((tftStudentID.UserName)='" & sUserName & "'))"
    
    If ConnectRS(HSESDB, vRS, sSQL) = False Then
        'fatal error
        CatchError "modRSStudent", "SaveUserStudentID", "Unable to connect Recordset with SQL Expression '" & sSQL & "'"
        GoTo ReleaseAndExit
    End If
    
On Error Resume Next
    While AnyRecordExisted(vRS) = True
        vRS.Delete
        vRS.Requery
    Wend
    
On Error GoTo ReleaseAndExit

    vRS.AddNew
    
    vRS.Fields("UserName") = sUserName
    vRS.Fields("StudentID") = sStudentID
    vRS.Update
   
   'return true
   SaveUserStudentID = True
    
    
ReleaseAndExit:
    Set vRS = Nothing
End Function

Public Function StudentIDNotExistFromOther(sStudentID As String, Optional ByRef sUserName As String) As Boolean
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    'default
    StudentIDNotExistFromOther = False
    
    sSQL = "SELECT tftStudentID.UserName" & _
            " From tftStudentID" & _
            " WHERE (((tftStudentID.StudentID)='" & sStudentID & "'))"
    
    If ConnectRS(HSESDB, vRS, sSQL) = False Then
        'fatal error
        CatchError "modRSStudent", "StudentIDNotExistFromOther", "Unable to connect Recordset with SQL Expression '" & sSQL & "'"
        GoTo ReleaseAndExit
    End If
    
    On Error Resume Next
    If AnyRecordExisted(vRS) = True Then
        sUserName = ReadField(vRS.Fields("UserName"))
        'return true
        StudentIDNotExistFromOther = True
    Else
        'not found
        StudentIDNotExistFromOther = False
    End If
    
    
ReleaseAndExit:
    Set vRS = Nothing
End Function


Public Function AddStudent(newStudent As tStudent) As TranDBResult
    
    Dim vRS As New ADODB.Recordset
    Dim sFullname As String
    
    
    'default
    AddStudent = Failed
    
    'check Duplicate ID
    If StudentExistByID(newStudent.StudentID) = Success Then
        AddStudent = DuplicateID
        GoTo ReleaseAndExit
    End If
    
    'check duplicate full name
    sFullname = LCase(Trim(newStudent.FirstName) & Trim(newStudent.MiddleName) & Trim(newStudent.LastName))
    If FindDuplicateFullName(sFullname) = Success Then
        AddStudent = DuplicateFullName
        GoTo ReleaseAndExit
    End If
    
    'check each field
    
    
    
    
    'save
    If CreateDefaultvrs(vRS) = Success Then
        'add new befor writing
        vRS.AddNew
        
        If WriteToRecord(vRS, newStudent) = Success Then
            AddStudent = Success
        Else
            AddStudent = Failed
        End If
    
    Else
        AddStudent = NotConnected
        GoTo ReleaseAndExit
    End If
    
ReleaseAndExit:
    Set vRS = Nothing
End Function

Public Function EditStudent(vStudent As tStudent) As TranDBResult
    
    Dim vRS As New ADODB.Recordset
    Dim OldStudent As tStudent
    Dim NewFullName As String
    Dim OldFullName As String
    
    'check if user exist
    If GetStudentByID(vStudent.StudentID, OldStudent) = Success Then
            
            'get old full name
            OldFullName = LCase(Trim(OldStudent.FirstName) + Trim(OldStudent.MiddleName) + Trim(OldStudent.LastName))

            'get new full name
            NewFullName = LCase(Trim(vStudent.FirstName) + Trim(vStudent.MiddleName) + Trim(vStudent.LastName))
        
            'compare
            If OldFullName <> NewFullName Then
            
                NewFullName = LCase(Trim(vStudent.FirstName) & Trim(vStudent.MiddleName) & Trim(vStudent.LastName))

                If FindDuplicateFullName(NewFullName) = Success Then
                    'found duplicate
                    EditStudent = DuplicateFullName
                    GoTo ReleaseAndExit
                End If
            End If
    Else
        'id not found
        EditStudent = Failed
        GoTo ReleaseAndExit
    End If
    
    
    
    'validated
    'update
    If ConnectRS(HSESDB, vRS, "SELECT * From tblStudent WHERE (((tblStudent.StudentID)='" & vStudent.StudentID & "'));") = True Then
        If AnyRecordExisted(vRS) Then
            vRS.MoveFirst
            
            '---------------------------------------------
            ' S A V E
            '---------------------------------------------
            If WriteToRecord(vRS, vStudent) = Success Then
                'success
                'added
                EditStudent = Success
                
            Else
                'write to record failed
                EditStudent = Failed
                GoTo ReleaseAndExit
            End If
            
        Else
            'no record existed
            EditStudent = Failed
            GoTo ReleaseAndExit
        End If
    Else
        'rs not connected
        EditStudent = Failed
        GoTo ReleaseAndExit
    End If
        
ReleaseAndExit:
    Set vRS = Nothing
End Function




























Private Function ReadFromRecord(ByRef vRS As ADODB.Recordset, sStudentID As String, ByRef vStudent As tStudent) As TranDBResult
    
        ReadFromRecord = Failed

            With vStudent
                .StudentID = ReadField(vRS.Fields("studentid"))
                .FirstName = ReadField(vRS.Fields("firstname"))
                .MiddleName = ReadField(vRS.Fields("MiddleName"))
                .LastName = ReadField(vRS.Fields("LastName"))
                
                .CityAddress = ReadField(vRS.Fields("CityAddress"))
                .HomeAddress = ReadField(vRS.Fields("HomeAddress"))
                .BirthDate = ReadField(vRS.Fields("birthdate"))
                .PlaceOfBirth = ReadField(vRS.Fields("PlaceOfBirth"))
                .Gender = ReadField(vRS.Fields("Gender"))
                .Status = ReadField(vRS.Fields("Status"))
                .Citizenship = ReadField(vRS.Fields("Citizenship"))
                
                .BloodType = ReadField(vRS.Fields("BloodType"))
                .Religion = ReadField(vRS.Fields("Religion"))
                
                .LastSchoolName = ReadField(vRS.Fields("LastSchoolName"))
                .LastSchoolContactNumber = ReadField(vRS.Fields("LastSchoolContactNumber"))
                .LastSchoolAddress = ReadField(vRS.Fields("LastSchoolAddress"))
                   
                'parents
                .MotherName = ReadField(vRS.Fields("MotherName"))
                .MotherOccupation = ReadField(vRS.Fields("MotherOccupation"))
                .FatherName = ReadField(vRS.Fields("FatherName"))
                .FatherOccupation = ReadField(vRS.Fields("FatherOccupation"))
                .ParentsContactNumber = ReadField(vRS.Fields("ParentsContactNumber"))
                .ParentsAddress = ReadField(vRS.Fields("ParentsAddress"))
                
                .GuardianName = ReadField(vRS.Fields("GuardianName"))
                .GuardianAddress = ReadField(vRS.Fields("GuardianAddress"))
                .GuardianContactNumber = ReadField(vRS.Fields("GuardianContactNumber"))
                .OldAveGrade = ReadField(vRS.Fields("OldAveGrade"))
                
                .Transferee = ReadField(vRS.Fields("Transferee"))
                If IsNull(vRS.Fields("TransfereeYL")) Then
                    .TransfereeYL = 0
                Else
                    .TransfereeYL = ReadField(vRS.Fields("TransfereeYL"))
                End If

                
                .CreationDate = ReadField(vRS.Fields("creationdate"))
                .CreatedBy = ReadField(vRS.Fields("CreatedBy"))
                .ModifiedBy = ReadField(vRS.Fields("ModifiedBy"))
                .ModifiedDate = ReadField(vRS.Fields("ModifiedDate"))

            End With
                
        ReadFromRecord = Success
End Function


Private Function WriteToRecord(ByRef vRS As ADODB.Recordset, vStudent As tStudent) As TranDBResult
    On Error GoTo ReleaseAndExit
       'default
       WriteToRecord = Failed
       
        With vRS
        
            .Fields("studentid").Value = vStudent.StudentID
            .Fields("firstname").Value = vStudent.FirstName
            .Fields("middlename").Value = vStudent.MiddleName
            .Fields("lastname").Value = vStudent.LastName
            
            'gen info
            .Fields("gender").Value = vStudent.Gender

            .Fields("status").Value = Trim(vStudent.Status)
            .Fields("Citizenship").Value = vStudent.Citizenship
            .Fields("birthdate").Value = vStudent.BirthDate
            .Fields("placeofbirth").Value = vStudent.PlaceOfBirth
            .Fields("homeaddress").Value = vStudent.HomeAddress
            .Fields("cityaddress").Value = vStudent.CityAddress
            
            .Fields("Religion").Value = vStudent.Religion
            .Fields("BloodType").Value = vStudent.BloodType

            'last school
            .Fields("lastschoolname").Value = vStudent.LastSchoolName
            .Fields("lastschoolcontactnumber").Value = vStudent.LastSchoolContactNumber
            .Fields("lastschooladdress").Value = vStudent.LastSchoolAddress
            
            'parents
            .Fields("mothername").Value = vStudent.MotherName
            .Fields("motheroccupation").Value = vStudent.MotherOccupation
            .Fields("fathername").Value = vStudent.FatherName
            .Fields("fatheroccupation").Value = vStudent.FatherOccupation
            .Fields("parentscontactnumber").Value = vStudent.ParentsContactNumber
            .Fields("parentsaddress").Value = vStudent.ParentsAddress
            
            'guardian
            .Fields("guardianname").Value = vStudent.GuardianName
            .Fields("guardiancontactnumber").Value = vStudent.GuardianContactNumber
            .Fields("guardianaddress").Value = vStudent.GuardianAddress
            .Fields("OldAveGrade") = vStudent.OldAveGrade
            
            .Fields("Transferee").Value = vStudent.Transferee
            .Fields("TransfereeYL").Value = vStudent.TransfereeYL


            .Fields("creationdate").Value = vStudent.CreationDate
            .Fields("CreatedBy").Value = vStudent.CreatedBy
            
            If Len(vStudent.ModifiedBy) > 0 Then
                .Fields("ModifiedDate").Value = vStudent.ModifiedDate
                .Fields("ModifiedBy").Value = vStudent.ModifiedBy
            End If
            
            .Update
        End With
        
    'return success
    WriteToRecord = Success
    Exit Function

ReleaseAndExit:
    CatchError "modRSStudent", "WriteToRecord", "Unable to continue updating record with Error:" & Err.Description
End Function





Public Function FindDuplicateFullName(sFullname As String) As TranDBResult
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    
    'deault
    FindDuplicateFullName = Failed
    
    sFullname = LCase(Trim(sFullname))
    sSQL = "SELECT tblStudent.StudentID  From tblStudent " & _
            " WHERE (LCase$(trim( [tblStudent]![FirstName] & [tblStudent]![MiddleName]& [tblStudent]![LastName] ))='" & sFullname & "');"
    
    If ConnectRS(HSESDB, vRS, sSQL) = True Then
        If AnyRecordExisted(vRS) = True Then
            FindDuplicateFullName = Success
        Else
            FindDuplicateFullName = Failed
        End If
    Else
        FindDuplicateFullName = Failed
    End If
    
    Set vRS = Nothing
End Function







Public Function DeleteStudent(sStudentID As String, Optional ShowMessage As Boolean = True) As TranDBResult
    Dim vRS As New ADODB.Recordset
    
    Dim lEnrolmentCount As Long
    
    'default
    DeleteStudent = Failed
    
    'check
    If GetEnrolmentCountByStudent(sStudentID, lEnrolmentCount) = Success Then
        If lEnrolmentCount > 0 Then
            If ShowMessage = True Then
                'temp
                MsgBox "temp: show is already used", vbExclamation
            End If
            
            DeleteStudent = Failed
            Exit Function
        End If
    Else
        'Student entry not exist
        CatchError "frmAllStudent", "listRecord_DblClick", "GetEnrolmentCountByStudent(lvKey, lEnrolmentCount) = success"
    End If
    
    
    
    If ConnectRS(HSESDB, vRS, "DELETE * From tblStudent WHERE (((tblStudent.StudentID)='" & sStudentID & "'));") = True Then
        DeleteStudent = Success
    Else
        DeleteStudent = Failed
    End If
    
    Set vRS = Nothing
End Function













Public Function CompleteGender(sCode As String) As String
    Select Case LCase(sCode)
        Case "m"
            CompleteGender = "Male"
        Case "f"
            CompleteGender = "Female"
        Case Else
            CompleteGender = "!Invalid Entry"
    End Select
End Function

Public Function CompleteStatus(sCode As String) As String
    Select Case Left(LCase(sCode), 2)
        Case "si"
            CompleteStatus = "Single"
        Case "ma"
            CompleteStatus = "Maried"
        Case "wi"
            CompleteStatus = "Widowed"
        Case "se"
            CompleteStatus = "separated"
        Case Else
            CompleteStatus = "!Invalid Entry"
    End Select
End Function

Public Function CompleteName(vStudent As tStudent) As String
    CompleteName = cSentenceCase(vStudent.LastName & ", " & vStudent.FirstName & " " & Left(vStudent.MiddleName, 1))
End Function

















Public Function GetStudentByID(sStudentID As String, ByRef vStudent As tStudent) As TranDBResult
    Dim vRS As New ADODB.Recordset
    
    If ConnectRS(HSESDB, vRS, "SELECT * From tblStudent WHERE (((tblStudent.StudentID)='" & sStudentID & "'));") = True Then
        If AnyRecordExisted(vRS) = True Then
            ReadFromRecord vRS, sStudentID, vStudent
            GetStudentByID = Success
        Else
            GetStudentByID = Failed
        End If
    Else
        GetStudentByID = Failed
    End If
    
    Set vRS = Nothing
End Function








Public Function StudentExistByID(sStudentID As String) As TranDBResult
    
    Dim vRS As New ADODB.Recordset
        
    If ConnectRS(HSESDB, vRS, "SELECT * From tblStudent WHERE (((tblStudent.StudentID)='" & sStudentID & "'));") Then
        If vRS.RecordCount > 0 Then
            StudentExistByID = Success
        Else
            StudentExistByID = Failed
        End If
    Else
        StudentExistByID = Failed
       
    End If
    
    'release
    Set vRS = Nothing
End Function


Public Function CreateDefaultvrs(ByRef vRS As ADODB.Recordset) As TranDBResult
    'default
    CreateDefaultvrs = Failed
    
    If ConnectRS(HSESDB, vRS, "SELECT * FROM tblStudent") Then
        CreateDefaultvrs = Success
    End If
End Function

Public Function StudentRecordExist() As TranDBResult
    Dim vRS As New ADODB.Recordset
    
    If CreateDefaultvrs(vRS) = Success Then
        
        If AnyRecordExisted(vRS) = True Then
            StudentRecordExist = Success
        Else
            StudentRecordExist = Failed
        End If
        
    Else
        StudentRecordExist = Failed
    End If
    
    Set vRS = Nothing
End Function


