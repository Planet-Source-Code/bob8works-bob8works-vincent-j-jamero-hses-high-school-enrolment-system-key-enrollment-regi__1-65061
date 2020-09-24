Attribute VB_Name = "modDBSubject"
Option Explicit

Public Const KeySubject = "subj"

Public Type tSubject
    SubjectTitle As String
    SubjectID As String
    Description As String
    DepartmentID As String
    YearLevelID As Integer
End Type


Public Function CreateDefaultRSSubject(ByRef vRS As ADODB.Recordset) As TranDBResult
    'default
    CreateDefaultRSSubject = Failed
    
    If ConnectRS(HSESDB, vRS, "SELECT * FROM tblSubject") Then
        CreateDefaultRSSubject = Success
    End If
End Function

Public Function CreateRSSubjectBySectionID(sSectionID As String, ByRef vRS As ADODB.Recordset) As TranDBResult
    
    Dim sSQL As String
    
    'default
    CreateRSSubjectBySectionID = Failed
    
    sSQL = "SELECT tblSubject.SubjectID, tblSubject.SubjectTitle, tblDepartment.DepartmentTitle, tblYearLevel.YearLevelTitle, tblTeacher.TeacherTitle, tblSubject.Description" & _
            " FROM tblTeacher INNER JOIN (tblDepartment INNER JOIN ((tblYearLevel INNER JOIN tblSection ON tblYearLevel.YearLevelID = tblSection.YearLevelID) INNER JOIN tblSubject ON tblYearLevel.YearLevelID = tblSubject.YearLevelID) ON (tblDepartment.DepartmentID = tblSubject.DepartmentID) AND (tblDepartment.DepartmentID = tblSection.DepartmentID)) ON tblTeacher.TeacherID = tblSection.TeacherID" & _
            " Where (((tblSection.SectionID) = '" & sSectionID & "'))" & _
            " GROUP BY tblSubject.SubjectID, tblSubject.SubjectTitle, tblDepartment.DepartmentTitle, tblYearLevel.YearLevelTitle, tblTeacher.TeacherTitle, tblSubject.Description;"


    If ConnectRS(HSESDB, vRS, sSQL) Then
        CreateRSSubjectBySectionID = Success
    End If
End Function



Public Function AddSubject(newSubject As tSubject) As TranDBResult
    'possibe return values
        'Success
        'IDNotFound
        'DuplicateTitle
    
    Dim vRS As New ADODB.Recordset
    
    
    
    'find duplicate ID
    If SubjectExistByID(newSubject.SubjectID) = Success Then
        AddSubject = DuplicateID
        GoTo ReleaseAndExit
    End If
    
    'find duplicate TITLE
    If SubjectExistByTitle(newSubject.SubjectTitle) = Success Then
        AddSubject = DuplicateTitle
        GoTo ReleaseAndExit
    End If
    

    
    'check each fields
    If Len(Trim(newSubject.SubjectID)) < 1 Then
        AddSubject = InvalidSubjectSubjectID
        GoTo ReleaseAndExit
    End If
    
    If Len(Trim(newSubject.SubjectTitle)) < 1 Then
        AddSubject = InvalidSubjectSubjectTitle
        GoTo ReleaseAndExit
    End If
    
    If Len(Trim(newSubject.Description)) < 1 Then
        AddSubject = InvalidSubjectDescription
        GoTo ReleaseAndExit
    End If
    
    
    'check if department exist
    If DepartmentExistByID(newSubject.DepartmentID) <> Success Then
        AddSubject = InvalidSubjectDepartmentID
        GoTo ReleaseAndExit
    End If
    
    'check if department exist
    If YearLevelExistByID(newSubject.YearLevelID) <> Success Then
        AddSubject = InvalidSubjectYearLevelID
        GoTo ReleaseAndExit
    End If
    
 
    
    
    If CreateDefaultRSSubject(vRS) = Success Then
    
        'add new record
        vRS.AddNew
    
        vRS.Fields("Subjectid").Value = Trim(newSubject.SubjectID)
        vRS.Fields("Subjecttitle").Value = Trim(newSubject.SubjectTitle)
        vRS.Fields("departmentid").Value = Trim(newSubject.DepartmentID)
        vRS.Fields("yearlevelid").Value = Trim(newSubject.YearLevelID)
        vRS.Fields("Description").Value = Trim(newSubject.Description)
        
        
        vRS.Update
        
        AddSubject = Success
    Else
        AddSubject = Failed
    End If
    
    
    
ReleaseAndExit:
    'release
    Set vRS = Nothing
End Function



Public Function EditSubject(newSubject As tSubject) As TranDBResult
    'possibe return values
        'Success
        'InvalidID
        'DuplicateTitle
    
    Dim oldSubject As tSubject

    Dim vRS As New ADODB.Recordset
    


    'get old Subject
    If GetSubjectByID(newSubject.SubjectID, oldSubject) = Success Then
                
        If oldSubject.SubjectTitle <> newSubject.SubjectTitle Then
            'find duplicate title
            If SubjectExistByTitle(newSubject.SubjectTitle) = Success Then
                EditSubject = DuplicateTitle
                'exit function
                GoTo ReleaseAndExit
            End If
            
        End If
    Else
        'department not found
        'exit function
        EditSubject = InvalidID
        GoTo ReleaseAndExit
    End If
    

    'find record to edit

    If ConnectRS(HSESDB, vRS, "SELECT * From tblSubject WHERE (((tblSubject.SubjectID)='" & newSubject.SubjectID & "'));") Then
        If vRS.RecordCount < 1 Then
            EditSubject = InvalidID
            GoTo ReleaseAndExit
        End If
    End If
        
      
        'vrs'editing
        vRS.MoveFirst
        
        vRS.Fields("Subjecttitle").Value = Trim(newSubject.SubjectTitle)
        vRS.Fields("departmentid").Value = Trim(newSubject.DepartmentID)
        vRS.Fields("yearlevelid").Value = Trim(newSubject.YearLevelID)
        vRS.Fields("Description").Value = Trim(newSubject.Description)

        vRS.Update
            
        EditSubject = Success
        

ReleaseAndExit:
    'release
    Set vRS = Nothing
End Function


Public Function DeleteSubject(sSubjectID As String, Optional ShowMessage As Boolean = True) As TranDBResult
    
    Dim vRS As New ADODB.Recordset
    Dim lEnrolmentCount As Long
    
    'default
    DeleteSubject = Failed
    
    'check
    If GetEnrolmentCountBySubject(sSubjectID, lEnrolmentCount) = Success Then
        If lEnrolmentCount > 0 Then
            If ShowMessage = True Then
                'temp
                MsgBox "temp: show is already used", vbExclamation
            End If
            
            DeleteSubject = Failed
            Exit Function
        End If
    Else
        'subject entry not exist
        CatchError "frmAllSubject", "listRecord_DblClick", "GetEnrolmentCountBySubject(lvKey, lEnrolmentCount) = success"
    End If
    
    
    '----------------------------------------------------
    'delete
    '----------------------------------------------------
    If ConnectRS(HSESDB, vRS, "Delete * From tblSubject WHERE (((tblSubject.SubjectID)='" & sSubjectID & "'));") Then
        DeleteSubject = Success
    Else
        DeleteSubject = Success
    End If
    
    'release
    Set vRS = Nothing
End Function







Public Function GetSubjectByTitle(sSubjectTitle As String, ByRef vSubject As tSubject) As TranDBResult
    
    Dim vRS As New ADODB.Recordset
        
    If ConnectRS(HSESDB, vRS, "SELECT *  FROM tblSubject WHERE (((tblSubject.SubjectTitle)='" & sSubjectTitle & "'));") Then
        If vRS.RecordCount > 0 Then
            
            vSubject.SubjectID = ReadField(vRS.Fields("Subjectid"))
            vSubject.SubjectTitle = ReadField(vRS.Fields("Subjecttitle"))
            vSubject.DepartmentID = ReadField(vRS.Fields("departmentid"))
            vSubject.YearLevelID = ReadField(vRS.Fields("yearlevelid"))
            vSubject.Description = ReadField(vRS.Fields("Description"))
        
            GetSubjectByTitle = Success
        Else
            GetSubjectByTitle = Failed
        End If
    Else
        GetSubjectByTitle = Failed
    End If
    
    'release
    Set vRS = Nothing
End Function

Public Function GetSubjectByTitle_Opti(ByRef vRS As ADODB.Recordset, sSubjectTitle As String, ByRef vSubject As tSubject) As TranDBResult
    
    'set default
    GetSubjectByTitle_Opti = Failed

    
    'assumes that the recordset is already connected
    
    vRS.MoveFirst
    vRS.Find "SubjectTitle = '" & sSubjectTitle & "'"
    
    If RecordNoMatch(vRS) Then
        GetSubjectByTitle_Opti = InvalidTitle
    Else
        
        vSubject.SubjectID = ReadField(vRS.Fields("Subjectid"))
        vSubject.SubjectTitle = ReadField(vRS.Fields("Subjecttitle"))
        vSubject.DepartmentID = ReadField(vRS.Fields("departmentid"))
        vSubject.YearLevelID = ReadField(vRS.Fields("yearlevelid"))
        vSubject.Description = ReadField(vRS.Fields("Description"))
 
        GetSubjectByTitle_Opti = Success
    End If
End Function








Public Function SubjectExistByTitle(sSubjectTitle As String) As TranDBResult
    
    Dim vRS As New ADODB.Recordset
        
    If ConnectRS(HSESDB, vRS, "SELECT * From tblSubject WHERE (((tblSubject.SubjectTitle)='" & sSubjectTitle & "'));") Then
        If vRS.RecordCount > 0 Then
            SubjectExistByTitle = Success
        Else
            SubjectExistByTitle = Failed
        End If
    
    Else
    
        SubjectExistByTitle = Failed
    End If
    
    'release
    Set vRS = Nothing
End Function


Public Function SubjectExistByID(sSubjectID As String) As TranDBResult
    
    Dim vRS As New ADODB.Recordset
        
    If ConnectRS(HSESDB, vRS, "SELECT * From tblSubject WHERE (((tblSubject.SubjectID)='" & sSubjectID & "'));") Then
        If vRS.RecordCount > 0 Then
            SubjectExistByID = Success
        Else
            SubjectExistByID = Failed
        End If
        
    Else
        
        SubjectExistByID = Failed
    End If
    
    'release
    Set vRS = Nothing
End Function


Public Function GetSubjectMoveNext(ByRef vRS As ADODB.Recordset, ByRef vSubject As tSubject) As TranDBResult
    
    'asuming that vRS is already connected
    If Not vRS.EOF And Not vRS.BOF Then
    
        'SUCCESS: Record exist
        'get values
        '----------------------------------------------------------------
        
        vSubject.SubjectID = ReadField(vRS.Fields("Subjectid"))
        vSubject.SubjectTitle = ReadField(vRS.Fields("Subjecttitle"))
        vSubject.DepartmentID = ReadField(vRS.Fields("departmentid"))
        vSubject.YearLevelID = ReadField(vRS.Fields("yearlevelid"))
        vSubject.Description = ReadField(vRS.Fields("Description"))
        
        vRS.MoveNext
        'move to the next record
        vRS.MoveNext
        'return true
        GetSubjectMoveNext = Success
    Else
        GetSubjectMoveNext = Failed
    End If
    
End Function




Public Function CreateRSSubject(ByRef vRS As ADODB.Recordset, Optional sDepartmentTitle As String, Optional sYearLevelTitle As String, Optional sTeacherTitle As String) As TranDBResult
    Dim sSQL As String
    Dim WHERE_Clause_Added As Boolean
    
    
    'default
    CreateRSSubject = Failed
    'set starting querry
    sSQL = "SELECT tblSubject.SubjectID, tblYearLevel.YearLevelTitle, tblSubject.SubjectTitle, tblDepartment.DepartmentTitle, tblTeacher.TeacherTitle, tblSubject.RoomNumber FROM tblTeacher INNER JOIN (tblYearLevel INNER JOIN (tblDepartment INNER JOIN tblSubject ON tblDepartment.DepartmentID = tblSubject.DepartmentID) ON tblYearLevel.YearLevelID = tblSubject.YearLevelID) ON tblTeacher.Description = tblSubject.Description"

    
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
    
        CreateRSSubject = Success
    End If

End Function



Public Function GetSubjectByID(sSubjectID As String, ByRef vSubject As tSubject) As TranDBResult
    
    Dim vRS As New ADODB.Recordset

    If ConnectRS(HSESDB, vRS, "SELECT * From tblSubject WHERE (((tblSubject.SubjectID)='" & sSubjectID & "'));") Then
        If AnyRecordExisted(vRS) Then
            'SUCCESS: Record exist
            'get values
            '----------------------------------------------------------------

            vSubject.SubjectID = ReadField(vRS.Fields("Subjectid"))
            vSubject.SubjectTitle = ReadField(vRS.Fields("Subjecttitle"))
            vSubject.DepartmentID = ReadField(vRS.Fields("departmentid"))
            vSubject.YearLevelID = ReadField(vRS.Fields("yearlevelid"))
            vSubject.Description = ReadField(vRS.Fields("Description"))
            GetSubjectByID = Success
        
        Else

            'FAILED: record does not exist
            GetSubjectByID = Failed
        End If
    Else
        GetSubjectByID = Failed
    End If
    
    'release
    Set vRS = Nothing
End Function


Public Function GetSubjectByID_Opti(ByRef vRS As ADODB.Recordset, sSubjectID As String, ByRef vSubject As tSubject) As TranDBResult
    'set default
    GetSubjectByID_Opti = Failed


    'assumes that the recordset is already connected
    If RSMoveFirst(vRS) Then
        vRS.Find "Subjectid = '" & sSubjectID & "'"
        
        If RecordNoMatch(vRS) Then
            GetSubjectByID_Opti = InvalidID
        Else
        
        vSubject.SubjectID = ReadField(vRS.Fields("Subjectid"))
        vSubject.SubjectTitle = ReadField(vRS.Fields("Subjecttitle"))
        vSubject.DepartmentID = ReadField(vRS.Fields("departmentid"))
        vSubject.YearLevelID = ReadField(vRS.Fields("yearlevelid"))
        vSubject.Description = ReadField(vRS.Fields("Description"))

        GetSubjectByID_Opti = Success
        End If
    Else
        GetSubjectByID_Opti = InvalidID
    End If
End Function




Public Function SubjectRecordExist() As TranDBResult
    Dim vRS As New ADODB.Recordset
    
    If CreateDefaultRSSubject(vRS) = Success Then
        If AnyRecordExisted(vRS) = True Then
            SubjectRecordExist = Success
        Else
            SubjectRecordExist = Failed
        End If
    Else
        SubjectRecordExist = Failed
    End If
    
    Set vRS = Nothing
End Function


Public Function ExecuteDeleteSubject(sSubjectID As String) As TranDBResult
    
      'check if record exist and if it is edited by other user
    If MsgBox("WARNING:" & vbNewLine & _
        "Deleting this Subject entry will affect all other record" & vbNewLine & vbNewLine & _
        "Delete this record anyway?", vbQuestion + vbYesNo) = vbYes Then
            
        If DeleteSubject(sSubjectID) = Success Then
            MsgBox "Subject entry and other related record succesfully deleted.", vbInformation
            ExecuteDeleteSubject = Success
        Else
            MsgBox "Deleting Subject entry went failed.", vbExclamation
            ExecuteDeleteSubject = Failed
        End If
    Else
        ExecuteDeleteSubject = Failed
    End If
End Function


Public Function GetSubjectByDeptByYLCount(sDepartmentID As String, lYearLevelID As Integer, lSubjectCount As Long) As TranDBResult
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    'default
    GetSubjectByDeptByYLCount = Failed
    sSQL = "SELECT Count(*) AS SubjectCount, tblSubject.DepartmentID, tblSubject.YearLevelID" & _
            " From tblSubject" & _
            " GROUP BY tblSubject.DepartmentID, tblSubject.YearLevelID" & _
            " HAVING (((tblSubject.DepartmentID)='" & sDepartmentID & "') AND ((tblSubject.YearLevelID)=" & lYearLevelID & "));"

    If ConnectRS(HSESDB, vRS, sSQL) = True Then
        lSubjectCount = vRS.Fields("SubjectCount").Value
        GetSubjectByDeptByYLCount = Success
    Else
        lSubjectCount = 0
        GetSubjectByDeptByYLCount = Failed
    End If
    
    
    Set vRS = Nothing
End Function


Public Function CreatetRSSubjectByDeptByYL(ByRef vRS As ADODB.Recordset, sDepartmentID As String, lYearLevelID As Integer) As TranDBResult
    
    
    Dim sSQL As String
    
    'default
    CreatetRSSubjectByDeptByYL = Failed
    sSQL = "SELECT tblSubject.SubjectID" & _
            " From tblSubject" & _
            " Where (((tblSubject.DepartmentID) = '" & sDepartmentID & "') And ((tblSubject.YearLevelID) = " & lYearLevelID & "))" & _
            " GROUP BY tblSubject.SubjectID;"


    If ConnectRS(HSESDB, vRS, sSQL) = True Then
        CreatetRSSubjectByDeptByYL = Success
    Else
        CreatetRSSubjectByDeptByYL = Failed
    End If
   

End Function



Public Function GetNewSubjectID(ByRef sNewID As String) As TranDBResult
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    Dim NewDNumber As Long
    
    sSQL = "SELECT 'SUB-' & String$(6-Len(Count(*)+1),'0') & Count(*)+1 AS NewID" & _
            " FROM tblSubject;"


    
    If ConnectRS(HSESDB, vRS, sSQL) = True Then

        sNewID = vRS.Fields("NewID").Value
        
        While SubjectExistByID(sNewID) = Success
            NewDNumber = Val(Right(sNewID, 6)) + 1
            sNewID = "SUB-" & String(6 - Len(NewDNumber), "0") & NewDNumber
        Wend
        
        GetNewSubjectID = Success

        GetNewSubjectID = Success
        
    Else

        GetNewSubjectID = Failed
    End If
    
    
    Set vRS = Nothing
End Function




