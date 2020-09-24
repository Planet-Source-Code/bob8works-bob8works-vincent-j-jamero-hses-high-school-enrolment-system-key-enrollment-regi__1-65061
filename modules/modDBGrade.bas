Attribute VB_Name = "modDBGrade"
Option Explicit

Public Const KeyGrade = "grad"

Public Type tGrade

    GradeID As String
    EnrolmentID As String
    SubjectOfferingID As String
    GradeValue As String
    
End Type


Public Function CreateDefaultRSGrade(ByRef vRS As ADODB.Recordset) As TranDBResult
    'default
    CreateDefaultRSGrade = Failed
    
    If ConnectRS(HSESDB, vRS, "SELECT * FROM tblGrade") Then
        CreateDefaultRSGrade = Success
    End If
End Function





Public Function AddGrade(newGrade As tGrade) As TranDBResult
    'possibe return values
        'Success
        'IDNotFound
        'DuplicateTitle
    
    Dim vRS As New ADODB.Recordset
    
    
    
    'check enrolment ID
    If EnrolmentExistByID(newGrade.EnrolmentID) <> Success Then
        AddGrade = InvalidGradeEnrolmentID
        GoTo ReleaseAndExit
    End If
        
    'check Subject ID
    If SubjectOfferingExistByID(newGrade.SubjectOfferingID) <> Success Then
        AddGrade = InvalidGradeSubjectID
        GoTo ReleaseAndExit
    End If
    
    If newGrade.GradeValue < 0 Or newGrade.GradeValue > 100 Then
        AddGrade = InvalidGradeGradeValue
        GoTo ReleaseAndExit
    End If
    
    If CreateDefaultRSGrade(vRS) = Success Then
    
        'add new record
        vRS.AddNew
        vRS.Fields("GradeID").Value = Trim(newGrade.EnrolmentID) & "-" & Trim(newGrade.SubjectOfferingID)
        vRS.Fields("EnrolmentID").Value = Trim(newGrade.EnrolmentID)
        vRS.Fields("SubjectOfferingID").Value = Trim(newGrade.SubjectOfferingID)
        vRS.Fields("GradeValue").Value = FormatNumber(newGrade.GradeValue, 2)
        
        
        vRS.Update
        
        AddGrade = Success
    Else
        AddGrade = Failed
    End If
    
    
    
ReleaseAndExit:
    'release
    Set vRS = Nothing
End Function


Public Function UpdateGrade(newGrade As tGrade) As TranDBResult
    'possibe return values
        'Success
        'IDNotFound
        'DuplicateTitle
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    
    'check enrolment ID
    If EnrolmentExistByID(newGrade.EnrolmentID) <> Success Then
        UpdateGrade = InvalidGradeEnrolmentID
        GoTo ReleaseAndExit
    End If
        
    'check Subject ID
    If SubjectExistByID(newGrade.SubjectOfferingID) <> Success Then
        UpdateGrade = InvalidGradeSubjectID
        GoTo ReleaseAndExit
    End If
    
    If newGrade.GradeValue < 0 Or newGrade.GradeValue > 100 Then
        UpdateGrade = InvalidGradeGradeValue
        GoTo ReleaseAndExit
    End If
    
    
    sSQL = "SELECT tblGrade.GradeID, tblGrade.EnrolmentID, tblGrade.SubjectOfferingID, tblGrade.GradeValue" & _
        " From tblGrade" & _
        " Where (((tblGrade.EnrolmentID) = '" & newGrade.EnrolmentID & "') And ((tblGrade.SubjectOfferingID) = '" & newGrade.SubjectOfferingID & "'))" & _
        " GROUP BY tblGrade.GradeID, tblGrade.EnrolmentID, tblGrade.SubjectOfferingID, tblGrade.GradeValue;"

    If CreateDefaultRSGrade(vRS) = Success Then
        'Trim(newGrade.EnrolmentID) & "-" & Trim(newGrade.SubjectOfferingID)
        
        vRS.Find "GradeID='" & Trim(newGrade.EnrolmentID) & "-" & Trim(newGrade.SubjectOfferingID) & "'"
        
        If RecordNoMatch(vRS) = True Then
            'add new record
            vRS.AddNew
            vRS.Fields("GradeID").Value = Trim(newGrade.EnrolmentID) & "-" & Trim(newGrade.SubjectOfferingID)
            vRS.Fields("EnrolmentID").Value = Trim(newGrade.EnrolmentID)
            vRS.Fields("SubjectID").Value = Trim(newGrade.SubjectOfferingID)
            vRS.Fields("GradeValue").Value = Trim(newGrade.GradeValue)
            vRS.Update
            
        End If
        
        
        
        
        
        
        UpdateGrade = Success
    Else
        UpdateGrade = Failed
    End If
    
    
    
ReleaseAndExit:
    'release
    Set vRS = Nothing
End Function

Public Function EditGrade(newGrade As tGrade) As TranDBResult
    
    Dim oldGrade As tGrade

    Dim vRS As New ADODB.Recordset
    


    'get old Grade
    If GetGradeByID(newGrade.GradeID, oldGrade) <> Success Then
        'department not found
        'exit function
        EditGrade = InvalidID
        GoTo ReleaseAndExit
    End If
    

    'find record to edit

    If ConnectRS(HSESDB, vRS, "SELECT * From tblGrade WHERE (((tblGrade.GradeID)='" & newGrade.GradeID & "'));") Then
        If AnyRecordExisted(vRS) = False Then
            EditGrade = InvalidID
            GoTo ReleaseAndExit
        End If
    End If
        
      
        'vrs'editing
        vRS.MoveFirst
        vRS.Fields("GradeValue").Value = FormatNumber(newGrade.GradeValue, 2)

        vRS.Update
            
        EditGrade = Success
        

ReleaseAndExit:
    'release
    Set vRS = Nothing
End Function


Public Function DeleteGrade(sEnrolmentID As String) As TranDBResult
    
    Dim vRS As New ADODB.Recordset
    
    If ConnectRS(HSESDB, vRS, "Delete * From tblGrade WHERE (((tblGrade.EnrolmentID)='" & sEnrolmentID & "'));") Then
        DeleteGrade = Success
    Else
        DeleteGrade = Success
    End If
    
    'release
    Set vRS = Nothing
End Function


Public Function GetGradeByID(sGradeID As String, ByRef vGrade As tGrade) As TranDBResult
    
    Dim vRS As New ADODB.Recordset
        
    If ConnectRS(HSESDB, vRS, "SELECT *  FROM tblGrade WHERE (((tblGrade.GradeID)='" & sGradeID & "'));") Then
        If vRS.RecordCount > 0 Then
            
            vGrade.GradeID = ReadField(vRS.Fields("GradeID"))
            vGrade.EnrolmentID = ReadField(vRS.Fields("EnrolmentID"))
            vGrade.SubjectOfferingID = ReadField(vRS.Fields("SubjectOfferingID"))
            vGrade.GradeValue = ReadField(vRS.Fields("GradeValue"))
        
            GetGradeByID = Success
        Else
            GetGradeByID = Failed
        End If
    Else
        GetGradeByID = Failed
    End If
    
    'release
    Set vRS = Nothing
End Function


Public Function GetGradeInfoByID(sGradeID As String, ByRef sSubjectTitle As String, ByRef dGradeValue As Double, ByRef sStudentName As String, ByRef sYearLevelTitle As String) As TranDBResult
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    sSQL = "SELECT tblGrade.GradeID, tblSubject.SubjectTitle, tblGrade.GradeValue, [tblStudent]![LastName]+', '+[tblStudent]![FirstName]+' '+[tblStudent]![MiddleName] AS StudentName, tblYearLevel.YearLevelTitle" & _
            " FROM tblYearLevel INNER JOIN (tblSubject INNER JOIN (tblStudent INNER JOIN (tblEnrolment INNER JOIN tblGrade ON tblEnrolment.EnrolmentID = tblGrade.EnrolmentID) ON tblStudent.StudentID = tblEnrolment.StudentID) ON tblSubject.SubjectID = tblGrade.SubjectOfferingID) ON tblYearLevel.YearLevelID = tblSubject.YearLevelID" & _
            " WHERE (((tblGrade.GradeID)='" & sGradeID & "'));"


    If ConnectRS(HSESDB, vRS, sSQL) Then
        If vRS.RecordCount > 0 Then
        
            sSubjectTitle = ReadField(vRS.Fields("SubjectTitle"))
            dGradeValue = ReadField(vRS.Fields("GradeValue"))
            sStudentName = ReadField(vRS.Fields("StudentName"))
            sYearLevelTitle = ReadField(vRS.Fields("YearLevelTitle"))
            
            GetGradeInfoByID = Success
        Else
            GetGradeInfoByID = Failed
        End If
    Else
        GetGradeInfoByID = Failed
    End If
    
    'release
    Set vRS = Nothing
End Function


Public Function GradeExistByID(sEnrolmentID As String) As TranDBResult
    
    Dim vRS As New ADODB.Recordset
        
    If ConnectRS(HSESDB, vRS, "SELECT * From tblGrade WHERE (((tblGrade.EnrolmentID)='" & sEnrolmentID & "'));") Then
        If vRS.RecordCount > 0 Then
            GradeExistByID = Success
        Else
            GradeExistByID = Failed
        End If
        
    Else
        
        GradeExistByID = Failed
    End If
    
    'release
    Set vRS = Nothing
End Function

Public Function TeacherExistInGradeByID(sGradeValue As String) As TranDBResult
    
    Dim vRS As New ADODB.Recordset
        
    If ConnectRS(HSESDB, vRS, "SELECT tblGrade.GradeValue From tblGrade WHERE (((tblGrade.GradeValue)='" & sGradeValue & "'));") Then
        If vRS.RecordCount > 0 Then
            TeacherExistInGradeByID = Success
        Else
            TeacherExistInGradeByID = Failed
        End If
        
    Else
        
        TeacherExistInGradeByID = Failed
    End If
    
    'release
    Set vRS = Nothing
End Function










Public Function CreateRSGrade(ByRef vRS As ADODB.Recordset, Optional sDepartmentTitle As String, Optional sYearLevelTitle As String, Optional sTeacherTitle As String) As TranDBResult
    Dim sSQL As String
    Dim WHERE_Clause_Added As Boolean
    
    
    'default
    CreateRSGrade = Failed
    'set starting querry
    sSQL = "SELECT tblGrade.EnrolmentID, tblYearLevel.YearLevelTitle, tblGrade.GradeTitle, tblDepartment.DepartmentTitle, tblTeacher.TeacherTitle, tblGrade.RoomNumber FROM tblTeacher INNER JOIN (tblYearLevel INNER JOIN (tblDepartment INNER JOIN tblGrade ON tblDepartment.SubjectID = tblGrade.SubjectOfferingID) ON tblYearLevel.YearLevelID = tblGrade.YearLevelID) ON tblTeacher.GradeValue = tblGrade.GradeValue"

    
    'first criteria
    If Len(sDepartmentTitle) > 0 Then
        WHERE_Clause_Added = True
        sSQL = sSQL & " WHERE (((tblDepartment.DepartmentTitle)='" & sDepartmentTitle & "')"
        
    End If
    
    
    
    If Len(sYearLevelTitle) > 1 Then
            
        If WHERE_Clause_Added <> True Then
            sSQL = sSQL & " WHERE ("
            WHERE_Clause_Added = True
        Else
            sSQL = sSQL & " AND "
        End If

        sSQL = sSQL & " ((tblYearLevel.YearLevelTitle)='" & sYearLevelTitle & "')"
        
    End If
    
    
      
    

    
    'close querry
    If WHERE_Clause_Added = True Then
        sSQL = sSQL & ");"
    End If
    
    MsgBox sSQL
    
    If ConnectRS(HSESDB, vRS, sSQL) Then
    
        CreateRSGrade = Success
    End If

End Function


'create rs by student
Public Function CreateRSGradeByStudent(sStudentID As String, ByRef RSGrade As ADODB.Recordset) As TranDBResult
    
    Dim sSQL As String
    
    
    sSQL = "SELECT tblGrade.GradeID AS lvKey, tblSubject.SubjectTitle, tblGrade.GradeValue, [LastName]+', '+[FIrstName]+' '+[MiddleName] AS [Full Name], tblEnrolment.StudentID, tblSchoolYear.SchoolYearTitle, tblYearLevel.YearLevelTitle, tblSection.SectionTitle" & _
            " FROM tblYearLevel INNER JOIN (tblSubject INNER JOIN (tblStudent INNER JOIN (tblSection INNER JOIN (tblSchoolYear INNER JOIN (tblEnrolment INNER JOIN tblGrade ON tblEnrolment.EnrolmentID = tblGrade.EnrolmentID) ON tblSchoolYear.SchoolYearTitle = tblEnrolment.SchoolYear) ON tblSection.SectionID = tblEnrolment.SectionID) ON tblStudent.StudentID = tblEnrolment.StudentID) ON tblSubject.SubjectID = tblGrade.SubjectOfferingID) ON (tblYearLevel.YearLevelID = tblSubject.YearLevelID) AND (tblYearLevel.YearLevelID = tblSection.YearLevelID)" & _
            " Where (((tblEnrolment.StudentID) = '" & sStudentID & "'))" & _
            " ORDER BY [LastName]+', '+[FIrstName]+' '+[MiddleName];"
    
    If ConnectRS(HSESDB, RSGrade, sSQL) Then
        If AnyRecordExisted(RSGrade) Then
            CreateRSGradeByStudent = Success
        Else
        'FAILED: record does not exist
            CreateRSGradeByStudent = Failed
        End If
    Else
        CreateRSGradeByStudent = Failed
    End If
End Function

Public Function GetLatesAveGradeByStudentByYearLevel(ByRef AveGrade As Double, sStudentID As String, iYearLevelID As Integer) As TranDBResult
    Dim RSGrade As ADODB.Recordset
    Dim sSQL As String
    
    
    sSQL = "SELECT Avg(tblGrade.GradeValue) AS AvgOfGradeValue, tblEnrolment.StudentID, tblYearLevel.YearLevelID" & _
            " FROM tblYearLevel INNER JOIN (tblSection INNER JOIN (tblEnrolment INNER JOIN tblGrade ON tblEnrolment.EnrolmentID = tblGrade.EnrolmentID) ON tblSection.SectionID = tblEnrolment.SectionID) ON tblYearLevel.YearLevelID = tblSection.YearLevelID" & _
            " GROUP BY tblEnrolment.StudentID, tblYearLevel.YearLevelID" & _
            " HAVING (((tblEnrolment.StudentID)='" & sStudentID & "') AND ((tblYearLevel.YearLevelID)=" & iYearLevelID & "));"

    
    If ConnectRS(HSESDB, RSGrade, sSQL) Then
        If AnyRecordExisted(RSGrade) Then
            AveGrade = RSGrade.Fields("AvgOfGradeValue").Value
            GetLatesAveGradeByStudentByYearLevel = Success
        Else
        'FAILED: record does not exist
            GetLatesAveGradeByStudentByYearLevel = Failed
        End If
    Else
        GetLatesAveGradeByStudentByYearLevel = Failed
    End If
    
    Set RSGrade = Nothing
End Function


Public Function GetAveGradeByStudentIDByYLTitle(sStudentID As String, sYearLevelTitle As String, ByRef AveGrade As Double) As TranDBResult
    Dim vRS As ADODB.Recordset
    Dim sSQL As String
    
    
    sSQL = "SELECT Avg(tblGrade.GradeValue) AS AvgOfGradeValue" & _
            " FROM tblYearLevel INNER JOIN (tblSubject INNER JOIN (tblEnrolment INNER JOIN tblGrade ON tblEnrolment.EnrolmentID = tblGrade.EnrolmentID) ON tblSubject.SubjectID = tblGrade.SubjectOfferingID) ON tblYearLevel.YearLevelID = tblSubject.YearLevelID" & _
            " WHERE (((tblEnrolment.StudentID)='" & sStudentID & "') AND ((tblYearLevel.YearLevelTitle)='" & sYearLevelTitle & "'));"



    If ConnectRS(HSESDB, vRS, sSQL) Then
        If AnyRecordExisted(vRS) = True Then

            AveGrade = ReadField(vRS.Fields("AvgOfGradeValue"))
            GetAveGradeByStudentIDByYLTitle = Success
        
        Else
        'FAILED: record does not exist
            GetAveGradeByStudentIDByYLTitle = Failed
        End If
    Else
    
        'not connected
        GetAveGradeByStudentIDByYLTitle = Failed

    End If
    
    Set vRS = Nothing
    
    
End Function


Public Function GradeRecordExist() As TranDBResult
    Dim vRS As New ADODB.Recordset
    
    If CreateDefaultRSGrade(vRS) = Success Then
        If AnyRecordExisted(vRS) = True Then
            GradeRecordExist = Success
        Else
            GradeRecordExist = Failed
        End If
    Else
        GradeRecordExist = Failed
    End If
    
    Set vRS = Nothing
End Function

Public Function IsGradeEditable(sGradeID As String, ByRef Editable As Boolean) As TranDBResult
    
    Dim RSGStudent As New ADODB.Recordset
    Dim sSQL As String
    Dim sStudentID As String
    Dim curLyearLevelID As Integer
    Dim sSchoolYearTitle As String
    Dim lYearLevelID As Integer
    
    'default
    Editable = False
    
    
    sSQL = "SELECT tblEnrolment.StudentID, tblSection.YearLevelID, tblGrade.GradeID" & _
            " FROM tblSection INNER JOIN (tblEnrolment INNER JOIN tblGrade ON tblEnrolment.EnrolmentID = tblGrade.EnrolmentID) ON tblSection.SectionID = tblEnrolment.SectionID" & _
            " WHERE (((tblGrade.GradeID)='" & sGradeID & "'));"

    If ConnectRS(HSESDB, RSGStudent, sSQL) = False Then
        IsGradeEditable = InvalidGradeID
        Exit Function
    End If
    
    sStudentID = ReadField(RSGStudent.Fields("StudentID"))
    curLyearLevelID = ReadField(RSGStudent.Fields("YearLevelID"))
    
    
    If GetLatestSchoolYearYearLevel(sStudentID, sSchoolYearTitle, lYearLevelID) <> Success Then
        IsGradeEditable = InvalidGradeID
        Exit Function
    End If
    
    If curLyearLevelID < lYearLevelID Then
        Editable = False
    Else
        Editable = True
    End If
    
    Set RSGStudent = Nothing
    
    IsGradeEditable = Success
End Function








Public Function CreateRSGradeBySectionOfferingID(ByRef vRS As ADODB.Recordset, sSectionOfferingID As String) As TranDBResult
    
    Dim sSQL As String
    
    sSQL = "SELECT tblSubjectOffering.SubjectOfferingID " & _
            " FROM tblSectionOffering INNER JOIN tblSubjectOffering ON tblSectionOffering.SectionOfferingID = tblSubjectOffering.SectionOfferingID" & _
            " Where (((tblSectionOffering.SectionOfferingID) = '" & sSectionOfferingID & "'))"
    
    If ConnectRS(HSESDB, vRS, sSQL) = True Then
        CreateRSGradeBySectionOfferingID = Success
    Else
        CreateRSGradeBySectionOfferingID = Failed
    End If
End Function





