Attribute VB_Name = "modDBEnrolment"
Option Explicit

'started: December 19, 2005


Public Const KeyEnrolment = "enro"

Public Type tEnrolment
    EnrolmentID As String
    StudentID As String
    SchoolYear As String
    SectionOfferingID As String
    DateEnroled As Date
    
    CreationDate As Date
    CreatedBy As String
    ModifiedDate As Date
    ModifiedBy As String
End Type



Public Function AddEnrolment(NewEnrolment As tEnrolment) As TranDBResult
    
    Dim vRS As New ADODB.Recordset

    
    
    
    '--------------------------------------------------------------
    'check
    '--------------------------------------------------------------

    'find duplicate ID
    If EnrolmentExistByID(NewEnrolment.EnrolmentID) = Success Then
        AddEnrolment = DuplicateID
        GoTo ReleaseAndExit
    End If
    
    'check school year
    If SchoolYearExistByTitle(NewEnrolment.SchoolYear) <> Success Then
        AddEnrolment = EnrolmentSchoolYearNotFound
        GoTo ReleaseAndExit
    End If
    
    If StudentExistByID(NewEnrolment.StudentID) <> Success Then
        AddEnrolment = EnrolmentStudentIDNotFound
        GoTo ReleaseAndExit
    End If
    
    If SectionOfferingExistByID(NewEnrolment.SectionOfferingID) <> Success Then
        AddEnrolment = EnrolmentSectionIDNotFound
        GoTo ReleaseAndExit
    End If
    
   
    
    'set cration date
    NewEnrolment.CreationDate = FormatDateTime(Now, vbShortDate)

    
    
    
    If CreateDefaultRSEnrolment(vRS) = Success Then
        'add new record
        vRS.AddNew
        
        vRS.Fields("EnrolmentID") = NewEnrolment.EnrolmentID
        vRS.Fields("SchoolYear") = NewEnrolment.SchoolYear
        vRS.Fields("StudentID") = NewEnrolment.StudentID
        vRS.Fields("SectionOfferingID") = NewEnrolment.SectionOfferingID
        
        vRS.Fields("CreationDate") = NewEnrolment.CreationDate
        vRS.Fields("CreatedBy") = NewEnrolment.CreatedBy
        vRS.Fields("DateEnroled") = NewEnrolment.DateEnroled
        
        vRS.Update
        
        
        'Create Blank Grade
        If CreateBlankGrade(NewEnrolment) <> Success Then
            CatchError "ModRSEnrolment", "AddEnrolment", "'CreateBlankGrade(NewEnrolment)' went failed"
        End If
        
        AddEnrolment = Success
    
    Else
        AddEnrolment = Failed
    End If
    
    
    
ReleaseAndExit:
    'release
    Set vRS = Nothing
End Function

Public Function EditEnrolment(vEnrolment As tEnrolment) As TranDBResult
    'possibe return values
        'Success
        'InvalidID

    Dim vRS As New ADODB.Recordset
    

    

    If ConnectRS(HSESDB, vRS, "SELECT * From tblEnrolment WHERE (((tblEnrolment.EnrolmentID)='" & vEnrolment.EnrolmentID & "'));") Then
        If vRS.RecordCount < 1 Then
            EditEnrolment = InvalidID
            GoTo ReleaseAndExit
        End If
    End If
    
        'vRS.Fields("EnrolmentID") = vEnrolment.EnrolmentID
        vRS.Fields("SchoolYear") = vEnrolment.SchoolYear
        vRS.Fields("StudentID") = vEnrolment.StudentID
        vRS.Fields("SectionOfferingID") = vEnrolment.SectionOfferingID
        vRS.Fields("DateEnroled") = vEnrolment.DateEnroled
        
        vRS.Fields("ModifiedDate") = vEnrolment.ModifiedDate
        vRS.Fields("ModifiedBy") = vEnrolment.ModifiedBy

        vRS.Update
            
        EditEnrolment = Success
        

ReleaseAndExit:
    'release
    Set vRS = Nothing
End Function








Public Function ExecuteDeleteEnrolment(sEnrolmentID As String) As TranDBResult

      'check if record exist and if it is edited by other user
    If MsgBox("WARNING:" & vbNewLine & _
        "Deleting this Enrolment entry will affect all other record" & vbNewLine & vbNewLine & _
        "Delete this record anyway?", vbQuestion + vbYesNo) = vbYes Then
            
        If DeleteEnrolment(sEnrolmentID) = Success Then
            MsgBox "Enrolment entry and other related record succesfully deleted.", vbInformation
            ExecuteDeleteEnrolment = Success
        Else
            MsgBox "Deleting Enrolment entry went failed.", vbExclamation
            ExecuteDeleteEnrolment = Failed
        End If
    Else
        ExecuteDeleteEnrolment = Failed
    End If
End Function



Public Function DeleteEnrolment(sEnrolmentID As String) As TranDBResult
    
    Dim vRS As New ADODB.Recordset
    
    If ConnectRS(HSESDB, vRS, "Delete * From tblEnrolment WHERE (((tblEnrolment.EnrolmentID)='" & sEnrolmentID & "'));") Then
        DeleteEnrolment = Success
    Else
        DeleteEnrolment = Failed
    End If
    
    'release
    Set vRS = Nothing
End Function




Public Function GetEnrolmentMoveNext(ByRef vRS As ADODB.Recordset, ByRef vEnrolment As tEnrolment) As TranDBResult

    'asuming that vRS is already connected
    If Not vRS.EOF And Not vRS.BOF Then
    
        'SUCCESS: Record exist
        'get values
        '----------------------------------------------------------------
            vEnrolment.EnrolmentID = ReadField(vRS.Fields("EnrolmentID"))
            vEnrolment.SchoolYear = ReadField(vRS.Fields("SchoolYear"))
            vEnrolment.StudentID = ReadField(vRS.Fields("StudentID"))
            vEnrolment.SectionOfferingID = ReadField(vRS.Fields("SectionOfferingID"))
            vEnrolment.DateEnroled = ReadField(vRS.Fields("DateEnroled"))

            vEnrolment.CreationDate = ReadField(vRS.Fields("CreationDate"))
            vEnrolment.CreatedBy = ReadField(vRS.Fields("CreatedBy"))
            vEnrolment.ModifiedDate = ReadField(vRS.Fields("ModifiedDate"))
            vEnrolment.ModifiedBy = ReadField(vRS.Fields("ModifiedBy"))
           
           
        'move to the next record
        vRS.MoveNext
        'return true
        GetEnrolmentMoveNext = Success
    Else
        GetEnrolmentMoveNext = Failed
    End If
    
End Function



Public Function GetEnrolmentByID(sEnrolmentID As String, ByRef vEnrolment As tEnrolment) As TranDBResult
    
    Dim vRS As New ADODB.Recordset

    If ConnectRS(HSESDB, vRS, "SELECT * From tblEnrolment WHERE (((tblEnrolment.EnrolmentID)='" & sEnrolmentID & "'));") Then
        If AnyRecordExisted(vRS) Then
            'SUCCESS: Record exist
            'get values
            '----------------------------------------------------------------
            vEnrolment.EnrolmentID = ReadField(vRS.Fields("EnrolmentID"))
            vEnrolment.SchoolYear = ReadField(vRS.Fields("SchoolYear"))
            vEnrolment.StudentID = ReadField(vRS.Fields("StudentID"))
            vEnrolment.SectionOfferingID = ReadField(vRS.Fields("SectionOfferingID"))
            vEnrolment.DateEnroled = ReadField(vRS.Fields("DateEnroled"))
            
            vEnrolment.CreationDate = ReadField(vRS.Fields("CreationDate"))
            vEnrolment.CreatedBy = ReadField(vRS.Fields("CreatedBy"))
            vEnrolment.ModifiedDate = ReadField(vRS.Fields("ModifiedDate"))
            vEnrolment.ModifiedBy = ReadField(vRS.Fields("ModifiedBy"))
            
                        
            GetEnrolmentByID = Success
        
        Else

            'FAILED: record does not exist
            GetEnrolmentByID = Failed
        End If
    Else
        GetEnrolmentByID = Failed
    End If
    
    'release
    Set vRS = Nothing
End Function



Public Function GetEnrolmentByStudentIDByYearLevelTitle(sStudentID As String, sYearLevelTitle As String, ByRef vEnrolment As tEnrolment) As TranDBResult
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    sSQL = "SELECT tblEnrolment.EnrolmentID, tblEnrolment.SchoolYear, tblEnrolment.StudentID, tblEnrolment.SectionOfferingID, tblEnrolment.CreationDate" & _
            " FROM tblYearLevel INNER JOIN (tblSection INNER JOIN tblEnrolment ON tblSection.SectionOfferingID = tblEnrolment.SectionOfferingID) ON tblYearLevel.YearLevelID = tblSection.YearLevelID" & _
            " Where (((tblEnrolment.StudentID) = '" & sStudentID & "') And ((tblYearLevel.YearLevelTitle) = '" & sYearLevelTitle & "'))" & _
            " GROUP BY tblEnrolment.EnrolmentID, tblEnrolment.SchoolYear, tblEnrolment.StudentID, tblEnrolment.SectionOfferingID, tblEnrolment.CreationDate;"

    
    If ConnectRS(HSESDB, vRS, sSQL) Then
        If AnyRecordExisted(vRS) Then
            'SUCCESS: Record exist
            'get values
            '----------------------------------------------------------------
            vEnrolment.EnrolmentID = ReadField(vRS.Fields("EnrolmentID"))
            vEnrolment.SchoolYear = ReadField(vRS.Fields("SchoolYear"))
            vEnrolment.StudentID = ReadField(vRS.Fields("StudentID"))
            vEnrolment.SectionOfferingID = ReadField(vRS.Fields("SectionOfferingID"))
            vEnrolment.DateEnroled = ReadField(vRS.Fields("DateEnroled"))
            
            vEnrolment.CreationDate = ReadField(vRS.Fields("CreationDate"))
            vEnrolment.CreatedBy = ReadField(vRS.Fields("CreatedBy"))
            vEnrolment.ModifiedDate = ReadField(vRS.Fields("ModifiedDate"))
            vEnrolment.ModifiedBy = ReadField(vRS.Fields("ModifiedBy"))
            
            GetEnrolmentByStudentIDByYearLevelTitle = Success
        
        Else

            'FAILED: record does not exist
            GetEnrolmentByStudentIDByYearLevelTitle = Failed
        End If
    Else
        GetEnrolmentByStudentIDByYearLevelTitle = Failed
    End If
    
    'release
    Set vRS = Nothing
End Function

Public Function GetEnrolmentCountBySYBySectionTitle(sSchoolYearTitle As String, sSectionTitle As String, ByRef EnrolmentCount As Long, MaxAllowed As Long) As TranDBResult
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    sSQL = "SELECT Count(*) AS StudentCount, tblEnrolment.SchoolYear, tblSection.MaxStudentCount, tblEnrolment.SchoolYear, tblSection.SectionTitle" & _
            " FROM tblYearLevel INNER JOIN (tblSection INNER JOIN tblEnrolment ON tblSection.SectionOfferingID = tblEnrolment.SectionOfferingID) ON tblYearLevel.YearLevelID = tblSection.YearLevelID" & _
            " GROUP BY tblEnrolment.SchoolYear, tblSection.MaxStudentCount, tblEnrolment.SchoolYear, tblSection.SectionTitle" & _
            " Having (((tblEnrolment.SchoolYear) = '" & sSchoolYearTitle & "') And ((tblSection.SectionTitle) = '" & sSectionTitle & "'))" & _
            " ORDER BY tblSection.SectionTitle;"
    
    If ConnectRS(HSESDB, vRS, sSQL) Then
        If AnyRecordExisted(vRS) Then
            'SUCCESS: Record exist
            'get values
            '----------------------------------------------------------------
            EnrolmentCount = vRS.Fields("StudentCount").Value
            MaxAllowed = vRS.Fields("MaxStudentCount").Value
            GetEnrolmentCountBySYBySectionTitle = Success
        
        Else
            
            sSQL = "SELECT tblSection.MaxStudentCount" & _
                    " From tblSection" & _
                    " Where (((tblSection.SectionTitle) = '" & sSectionTitle & "'))" & _
                    " GROUP BY tblSection.MaxStudentCount;"
            If ConnectRS(HSESDB, vRS, sSQL) Then
                If AnyRecordExisted(vRS) Then
                    EnrolmentCount = 0
                    MaxAllowed = vRS.Fields("MaxStudentCount").Value
                    GetEnrolmentCountBySYBySectionTitle = Success
                Else
                    GetEnrolmentCountBySYBySectionTitle = Failed
                End If
            Else
                GetEnrolmentCountBySYBySectionTitle = Failed
            End If
            
        End If
    Else
        GetEnrolmentCountBySYBySectionTitle = Failed
    End If
    
    'release
    Set vRS = Nothing
End Function


Private Sub ReadFromRecord(ByRef vRS As ADODB.Recordset, ByRef vEnrolment As tEnrolment)
    
            vEnrolment.EnrolmentID = ReadField(vRS.Fields("EnrolmentID"))
            vEnrolment.SchoolYear = ReadField(vRS.Fields("SchoolYear"))
            vEnrolment.StudentID = ReadField(vRS.Fields("StudentID"))
            vEnrolment.SectionOfferingID = ReadField(vRS.Fields("SectionOfferingID"))
            vEnrolment.DateEnroled = ReadField(vRS.Fields("DateEnroled"))
            
            vEnrolment.CreationDate = ReadField(vRS.Fields("CreationDate"))
            vEnrolment.CreatedBy = ReadField(vRS.Fields("CreatedBy"))
            vEnrolment.ModifiedDate = ReadField(vRS.Fields("ModifiedDate"))
            vEnrolment.ModifiedBy = ReadField(vRS.Fields("ModifiedBy"))
End Sub


Public Function EnrolmentExistByStudentID(sStudentID As String) As TranDBResult
    
    Dim vRS As New ADODB.Recordset
        
    If ConnectRS(HSESDB, vRS, "SELECT * From tblEnrolment WHERE (((tblEnrolment.StudentID)='" & sStudentID & "'));") Then
        If vRS.RecordCount > 0 Then
            EnrolmentExistByStudentID = Success
        Else
            EnrolmentExistByStudentID = Failed
        End If
    Else
        EnrolmentExistByStudentID = Failed
    End If
    
    'release
    Set vRS = Nothing
End Function

Public Function EnrolmentExistByStudentIDBySchoolYear(sStudentID As String, sSchoolYear As String) As TranDBResult
    
    Dim vRS As New ADODB.Recordset
        
    If ConnectRS(HSESDB, vRS, "SELECT tblEnrolment.StudentID, tblEnrolment.SchoolYear From tblEnrolment WHERE (((tblEnrolment.StudentID)='" & sStudentID & "') AND ((tblEnrolment.SchoolYear)='" & sSchoolYear & "'));") Then
        If vRS.RecordCount > 0 Then
            EnrolmentExistByStudentIDBySchoolYear = Success
        Else
            EnrolmentExistByStudentIDBySchoolYear = Failed
        End If
    Else
        EnrolmentExistByStudentIDBySchoolYear = Failed
    End If
    
    'release
    Set vRS = Nothing
End Function

Public Function EnrolmentExistBySectionID(sSectionID As String) As TranDBResult
    
    Dim vRS As New ADODB.Recordset
        
    If ConnectRS(HSESDB, vRS, "SELECT * From tblEnrolment WHERE (((tblEnrolment.SectionOfferingID)='" & sSectionID & "'));") Then
        If vRS.RecordCount > 0 Then
            EnrolmentExistBySectionID = Success
        Else
            EnrolmentExistBySectionID = Failed
        End If
    Else
        EnrolmentExistBySectionID = Failed
    End If
    
    'release
    Set vRS = Nothing
End Function

Public Function EnrolmentExistByID(sEnrolmentID As String) As TranDBResult
    
    Dim vRS As New ADODB.Recordset
        
    If ConnectRS(HSESDB, vRS, "SELECT * From tblEnrolment WHERE (((tblEnrolment.EnrolmentID)='" & sEnrolmentID & "'));") Then
        If vRS.RecordCount > 0 Then
            EnrolmentExistByID = Success
        Else
            EnrolmentExistByID = Failed
        End If
    Else
        EnrolmentExistByID = Failed
       
    End If
    
    'release
    Set vRS = Nothing
End Function


Public Function CreateDefaultRSEnrolment(ByRef vRS As ADODB.Recordset) As TranDBResult
    'default
    CreateDefaultRSEnrolment = Failed
    
    If ConnectRS(HSESDB, vRS, "SELECT * FROM tblEnrolment") Then
        CreateDefaultRSEnrolment = Success
    End If
End Function

Public Function EnrolmentRecordExist() As TranDBResult
    Dim vRS As New ADODB.Recordset
    
    If CreateDefaultRSEnrolment(vRS) = Success Then
        
        If AnyRecordExisted(vRS) = True Then
            EnrolmentRecordExist = Success
        Else
            EnrolmentRecordExist = Failed
        End If
        
    Else
        EnrolmentRecordExist = Failed
    End If
    
    Set vRS = Nothing
End Function

Public Function GetNewEnrolmentID(sSchoolYear As String, sStudentID As String, ByRef sNewEnrolmentID As String) As TranDBResult
    'set default
    GetNewEnrolmentID = Failed
    
    If SchoolYearExistByTitle(sSchoolYear) <> Success Then
         GetNewEnrolmentID = EnrolmentSchoolYearNotFound
         Exit Function
    End If
    
    If StudentExistByID(sStudentID) <> Success Then
        GetNewEnrolmentID = EnrolmentStudentIDNotFound
        Exit Function
    End If
    
    If EnrolmentExistByStudentIDBySchoolYear(sStudentID, sSchoolYear) = Success Then
        GetNewEnrolmentID = EnrolmentDuplicateEntryWithInYear
        Exit Function
    End If
    
    sNewEnrolmentID = Left(Trim(sSchoolYear), 4) & "-" & Trim(sStudentID)
    GetNewEnrolmentID = Success
End Function


Public Function GetLatestSchoolYearYearLevel(sStudentID As String, ByRef sSchoolYearTitle As String, ByRef lYearLevelID As Integer) As TranDBResult
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    sSQL = "SELECT tblSectionOffering.SchoolYear, tblSection.YearLevelID" & _
            " FROM tblSection INNER JOIN (tblSectionOffering INNER JOIN tblEnrolment ON tblSectionOffering.SectionOfferingID = tblEnrolment.SectionOfferingID) ON tblSection.SectionID = tblSectionOffering.SectionID" & _
            " Where (((tblEnrolment.StudentID) = '" & sStudentID & "'))" & _
            " GROUP BY tblSectionOffering.SchoolYear, tblSection.YearLevelID" & _
            " ORDER BY tblSection.YearLevelID DESC;"

        If ConnectRS(HSESDB, vRS, sSQL) = True Then
            If AnyRecordExisted(vRS) = True Then
                sSchoolYearTitle = vRS.Fields("SchoolYear").Value
                lYearLevelID = vRS.Fields("YearLevelID").Value
                
                GetLatestSchoolYearYearLevel = Success
            Else
                sSchoolYearTitle = "0000"
                lYearLevelID = 0
                
                GetLatestSchoolYearYearLevel = Success
            End If
        Else
            GetLatestSchoolYearYearLevel = Failed
        End If
    
    Set vRS = Nothing
End Function


Public Function CreateBlankGrade(vEnrolment As tEnrolment) As TranDBResult
    
    Dim vRS As New ADODB.Recordset
    Dim vGrade As tGrade
    
    'deault
    CreateBlankGrade = Failed
    
    
            If CreateRSGradeBySectionOfferingID(vRS, vEnrolment.SectionOfferingID) = Success Then
                If AnyRecordExisted(vRS) = True Then
                    
                    CreateBlankGrade = Success
                    vRS.MoveFirst
                    
                    While vRS.EOF = False
                    
                        vGrade.EnrolmentID = vEnrolment.EnrolmentID
                        'default grade
                        vGrade.GradeValue = 60
                        vGrade.SubjectOfferingID = ReadField(vRS.Fields("SubjectOfferingID"))
                        
                        If AddGrade(vGrade) <> Success Then
                            MsgBox "Fatal Error adding grade"
                            CreateBlankGrade = Failed
                        End If
                        
                        vRS.MoveNext
                    Wend
                    
                End If
            Else
                'default subject rs not created
                CatchError "modDBEnrolment", "Create Blank Grade", "default subject rs not created"
            End If

    
ReleaseAndExit:
    Set vRS = Nothing
End Function



Public Function UpdateStudentGrades(vEnrolment As tEnrolment) As TranDBResult
    
    Dim vRS As New ADODB.Recordset
    Dim vSection As tSection
    Dim vGrade As tGrade
    
    'deault
    UpdateStudentGrades = Failed
    
    
    If GetSectionByID(vEnrolment.SectionOfferingID, vSection) = Success Then
        If CreatetRSSubjectByDeptByYL(vRS, vSection.DepartmentID, vSection.YearLevelID) = Success Then
            If AnyRecordExisted(vRS) = True Then
                
                UpdateStudentGrades = Success
                vRS.MoveFirst
                
                While vRS.EOF = False
                
                    vGrade.EnrolmentID = vEnrolment.EnrolmentID
                    vGrade.GradeValue = 0
                    vGrade.SubjectOfferingID = ReadField(vRS.Fields("SubjectOfferingID"))
                    
                    Dim AddGradeResult As TranDBResult
                    
                    AddGradeResult = UpdateGrade(vGrade)
                    If AddGradeResult <> Success Then
                        MsgBox "Error adding grade. procedure:" & AddGradeResult
                        UpdateStudentGrades = Failed
                    End If
                    
                    vRS.MoveNext
                Wend
                
            End If
        End If
    End If
    
ReleaseAndExit:
    Set vRS = Nothing
End Function


Public Function GetEnrolmentCountBySectionOfferingID(sSectionOfferingID As String, ByRef lStudentCount As Long) As TranDBResult
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    sSQL = "SELECT Count(tblEnrolment.EnrolmentID) AS CountOfEnrolmentID" & _
            " FROM tblSectionOffering INNER JOIN tblEnrolment ON tblSectionOffering.SectionOfferingID = tblEnrolment.SectionOfferingID" & _
            " WHERE (((tblSectionOffering.SectionOfferingID)='" & sSectionOfferingID & "'));"

    If ConnectRS(HSESDB, vRS, sSQL) = True Then
        lStudentCount = vRS.Fields("CountOfEnrolmentID").Value
        GetEnrolmentCountBySectionOfferingID = Success
    Else
        lStudentCount = 0
        GetEnrolmentCountBySectionOfferingID = Failed
    End If
    
    Set vRS = Nothing
End Function


Public Function GetEnrolmentCountBySubject(sSubjectID As String, ByRef lEnrolmentCount) As TranDBResult

    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    '------------------------------------------
    
    'default
    GetEnrolmentCountBySubject = Failed
    
    sSQL = "SELECT Count(tblEnrolment.EnrolmentID) AS EnrolmentCount, tblSubject.SubjectID" & _
            " FROM ((tblYearLevel INNER JOIN (tblSection INNER JOIN tblEnrolment ON tblSection.SectionOfferingID = tblEnrolment.SectionOfferingID) ON tblYearLevel.YearLevelID = tblSection.YearLevelID) INNER JOIN tblSubject ON tblYearLevel.YearLevelID = tblSubject.YearLevelID) INNER JOIN tblGrade ON (tblSubject.SubjectID = tblGrade.SubjectOfferingID) AND (tblEnrolment.EnrolmentID = tblGrade.EnrolmentID)" & _
            " GROUP BY tblSubject.SubjectID" & _
            " HAVING (((tblSubject.SubjectID)='" & sSubjectID & "'));"

    If ConnectRS(HSESDB, vRS, sSQL) = True Then
        lEnrolmentCount = vRS.Fields("EnrolmentCount").Value
        GetEnrolmentCountBySubject = Success
    Else
        lEnrolmentCount = -1
        GetEnrolmentCountBySubject = Failed
    End If
    
    Set vRS = Nothing
    
End Function


Public Function GetEnrolmentCountByStudent(sStudentID As String, ByRef lEnrolmentCount) As TranDBResult

    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    '------------------------------------------
    
    'default
    GetEnrolmentCountByStudent = Failed
    
    sSQL = "SELECT Count(tblEnrolment.EnrolmentID) AS EnrolmentCount" & _
            " FROM ((tblYearLevel INNER JOIN tblSection ON tblYearLevel.YearLevelID = tblSection.YearLevelID) INNER JOIN tblSubject ON tblYearLevel.YearLevelID = tblSubject.YearLevelID) INNER JOIN (tblSectionOffering INNER JOIN (tblStudent INNER JOIN tblEnrolment ON tblStudent.StudentID = tblEnrolment.StudentID) ON tblSectionOffering.SectionOfferingID = tblEnrolment.SectionOfferingID) ON tblSection.SectionID = tblSectionOffering.SectionID" & _
            " WHERE (((tblStudent.StudentID)='" & sStudentID & "'));"


    
    If ConnectRS(HSESDB, vRS, sSQL) = True Then
        lEnrolmentCount = vRS.Fields("EnrolmentCount").Value
        GetEnrolmentCountByStudent = Success
    Else
        lEnrolmentCount = -1
        GetEnrolmentCountByStudent = Failed
    End If
    
    Set vRS = Nothing
    
End Function


Public Function StudentPassedByYearLevel(sStudentID As String, iLyearLevelID As Integer, ByRef Passed As Boolean) As TranDBResult
    
End Function



'Creation Date: February 23, 2006
Public Function StudentEnroledBySchoolYear(sStudentID As String, sSchoolYear As String, ByRef IsEnroled As Boolean) As TranDBResult
    
    Dim vRS As Recordset
    Dim sSQL As String
    
    'defaul
    IsEnroled = False
        
        sSQL = "SELECT tblEnrolment.SchoolYear, tblStudent.StudentID" & _
                " FROM tblStudent INNER JOIN tblEnrolment ON tblStudent.StudentID = tblEnrolment.StudentID" & _
                " GROUP BY tblEnrolment.SchoolYear, tblStudent.StudentID" & _
                " HAVING (((tblEnrolment.SchoolYear)='" & sSchoolYear & "') AND ((tblStudent.StudentID)='" & sStudentID & "'));"


    If ConnectRS(HSESDB, vRS, sSQL) = True Then
        If AnyRecordExisted(vRS) = True Then
            IsEnroled = True
        Else
            IsEnroled = False
        End If
        StudentEnroledBySchoolYear = Success
    Else
        StudentEnroledBySchoolYear = Failed
        IsEnroled = False
    End If

    
    Set vRS = Nothing
End Function

Public Function GetAcademicRecord(sStudentID As String, iYearLevelID As Integer, ByRef dAveGrade As Double, ByRef bPassed As Boolean, ByRef sDepartmentID As String) As TranDBResult
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    
    'sSQL = "SELECT Avg(tblGrade.GradeValue) AS AvgOfGradeValue, IIf(Min([tblGrade].[GradeValue])<75 Or Avg([tblGrade].[GradeValue])<75,False,True) AS Passed" & _
    '        " FROM tblSection INNER JOIN (tblSectionOffering INNER JOIN (tblEnrolment INNER JOIN tblGrade ON tblEnrolment.EnrolmentID = tblGrade.EnrolmentID) ON tblSectionOffering.SectionOfferingID = tblEnrolment.SectionOfferingID) ON tblSection.SectionID = tblSectionOffering.SectionID" & _
    '        " WHERE (((tblSection.YearLevelID)=" & iYearLevelID & ") AND ((tblEnrolment.StudentID)='" & sStudentID & "'));"
'WHERE (((tblSection.YearLevelID)=1) AND ((tblEnrolment.StudentID)='2006-0000068'));

    

    sSQL = "SELECT First(tblSection.DepartmentID) AS FirstOfDepartmentID, Avg(tblGrade.GradeValue) AS AvgOfGradeValue, IIf(Min([tblGrade].[GradeValue])<75 Or Avg([tblGrade].[GradeValue])<75,False,True) AS Passed" & _
            " FROM tblSection INNER JOIN (tblSectionOffering INNER JOIN (tblEnrolment INNER JOIN tblGrade ON tblEnrolment.EnrolmentID = tblGrade.EnrolmentID) ON tblSectionOffering.SectionOfferingID = tblEnrolment.SectionOfferingID) ON tblSection.SectionID = tblSectionOffering.SectionID" & _
            " WHERE (((tblSection.YearLevelID)=" & iYearLevelID & ") AND ((tblEnrolment.StudentID)='" & sStudentID & "'))" & _
            " GROUP BY tblSection.DepartmentID;"


    If ConnectRS(HSESDB, vRS, sSQL) = True Then
        If AnyRecordExisted(vRS) = True Then
            dAveGrade = ReadField(vRS.Fields("AvgOfGradeValue"))
            bPassed = ReadField(vRS.Fields("Passed"))
            
            sDepartmentID = ReadField(vRS.Fields("FirstOfDepartmentID"))
        Else
            dAveGrade = 0
        End If
    Else
        dAveGrade = -1
    End If
    
    Set vRS = Nothing

    sSQL = "SELECT tblGrade.GradeValue" & _
            " FROM tblSection INNER JOIN ((tblSectionOffering INNER JOIN tblEnrolment ON tblSectionOffering.SectionOfferingID = tblEnrolment.SectionOfferingID) INNER JOIN tblGrade ON tblEnrolment.EnrolmentID = tblGrade.EnrolmentID) ON tblSection.SectionID = tblSectionOffering.SectionID" & _
            " Where (((tblSection.YearLevelID) = " & iYearLevelID & ") And ((tblEnrolment.StudentID) = '" & sStudentID & "'))" & _
            " GROUP BY tblGrade.GradeValue;"
            
    If ConnectRS(HSESDB, vRS, sSQL) = True Then
    Else
    End If
    
    Set vRS = Nothing
End Function
