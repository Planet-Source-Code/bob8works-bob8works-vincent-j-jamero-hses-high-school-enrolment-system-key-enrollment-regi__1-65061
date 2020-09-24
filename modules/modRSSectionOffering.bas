Attribute VB_Name = "modRSSectionOffering"
Option Explicit

Public Const KeySectionOffering = "seof"

Public Type tSectionOffering
    
    SectionOfferingID As String
    SectionID As String
    SchoolYear As String
    TeacherID As String
    MaxStudentCount As Integer
    MaxGrade As Double
    MinGrade As Double
    Note As String
    RoomID As String
    
    CreationDate As Date
    CreatedBy As String
    ModifiedDate As Date
    ModifiedBy As String
End Type


Public Function AddSectionOffering(vSectionOffering As tSectionOffering) As TranDBResult
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    If SectionOfferingExistByID(vSectionOffering.SectionOfferingID) = Success Then
        AddSectionOffering = DuplicateID
        GoTo ReleaseAndExit
    End If
    
    sSQL = "SELECT * FROM tblSectionOffering"
    
    If ConnectRS(HSESDB, vRS, sSQL) = True Then
        vRS.AddNew
        
        vRS.Fields("SectionOfferingID").Value = vSectionOffering.SectionOfferingID
        vRS.Fields("SectionID").Value = vSectionOffering.SectionID
        vRS.Fields("SchoolYear").Value = vSectionOffering.SchoolYear
        vRS.Fields("TeacherID").Value = vSectionOffering.TeacherID
        vRS.Fields("MaxStudentCount").Value = vSectionOffering.MaxStudentCount
        vRS.Fields("MaxGrade").Value = vSectionOffering.MaxGrade
        vRS.Fields("MinGrade").Value = vSectionOffering.MinGrade
        vRS.Fields("Note").Value = vSectionOffering.Note
        vRS.Fields("RoomID").Value = vSectionOffering.RoomID

        vRS.Fields("CreationDate").Value = vSectionOffering.CreationDate
        vRS.Fields("CreatedBy").Value = vSectionOffering.CreatedBy
        
        vRS.Update
        
        AddSectionOffering = Success
    Else
        AddSectionOffering = Failed
    End If
    
ReleaseAndExit:
    Set vRS = Nothing
    
End Function


Public Function EditSectionOffering(vSectionOffering As tSectionOffering) As TranDBResult
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    If SectionOfferingExistByID(vSectionOffering.SectionOfferingID) <> Success Then
        EditSectionOffering = InvalidID
        GoTo ReleaseAndExit
    End If
    
    sSQL = "SELECT * FROM tblSectionOffering WHERE SectionOfferingID='" & vSectionOffering.SectionOfferingID & "'"
    
    If ConnectRS(HSESDB, vRS, sSQL) = True Then
        If AnyRecordExisted(vRS) = True Then
            vRS.Fields("SectionOfferingID").Value = vSectionOffering.SectionOfferingID
            vRS.Fields("SectionID").Value = vSectionOffering.SectionID
            vRS.Fields("SchoolYear").Value = vSectionOffering.SchoolYear
            vRS.Fields("TeacherID").Value = vSectionOffering.TeacherID
            vRS.Fields("MaxStudentCount").Value = vSectionOffering.MaxStudentCount
            vRS.Fields("MaxGrade").Value = vSectionOffering.MaxGrade
            vRS.Fields("MinGrade").Value = vSectionOffering.MinGrade
            vRS.Fields("Note").Value = vSectionOffering.Note
            vRS.Fields("RoomID").Value = vSectionOffering.RoomID
            
            vRS.Fields("ModifiedDate").Value = vSectionOffering.ModifiedDate
            vRS.Fields("ModifiedBy").Value = vSectionOffering.ModifiedBy
        
            vRS.Update
            
            EditSectionOffering = Success
            
        Else
        
            EditSectionOffering = InvalidID
        End If
    Else
        EditSectionOffering = Failed
    End If
    
ReleaseAndExit:
    Set vRS = Nothing
    
End Function



Public Function DeleteSectionOffering(sSectionOfferingID As String) As TranDBResult
    
    Dim vRS As New ADODB.Recordset
    
    If ConnectRS(HSESDB, vRS, "Delete * From tblSectionOffering WHERE (((tblSectionOffering.SectionOfferingID)='" & sSectionOfferingID & "'));") Then
        DeleteSectionOffering = Success
    Else
        DeleteSectionOffering = Failed
    End If
    
    'release
    Set vRS = Nothing
End Function





Public Function GetSectionOfferingByID(sSectionOfferingID As String, ByRef vSectionOffering As tSectionOffering) As TranDBResult
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    
    
    'default
    GetSectionOfferingByID = Failed
    
    
    If Len(sSectionOfferingID) < 1 Then
        GetSectionOfferingByID = Failed
        GoTo ReleaseAndExit
    End If
    
    sSQL = "SELECT * FROM tblSectionOffering WHERE SectionOfferingID='" & sSectionOfferingID & "'"
    
    If ConnectRS(HSESDB, vRS, sSQL) = True Then
        If AnyRecordExisted(vRS) = True Then
        
            vSectionOffering.SectionOfferingID = ReadField(vRS.Fields("SectionOfferingID"))
            vSectionOffering.SectionID = ReadField(vRS.Fields("SectionID"))
            vSectionOffering.SchoolYear = ReadField(vRS.Fields("SchoolYear"))
            vSectionOffering.TeacherID = ReadField(vRS.Fields("TeacherID"))
            vSectionOffering.MaxStudentCount = ReadField(vRS.Fields("MaxStudentCount"))
            vSectionOffering.MaxGrade = ReadField(vRS.Fields("MaxGrade"))
            vSectionOffering.MinGrade = ReadField(vRS.Fields("MinGrade"))
            vSectionOffering.Note = ReadField(vRS.Fields("Note"))
            vSectionOffering.RoomID = ReadField(vRS.Fields("RoomID"))

            vSectionOffering.CreationDate = ReadField(vRS.Fields("CreationDate"))
            vSectionOffering.CreatedBy = ReadField(vRS.Fields("CreatedBy"))
            vSectionOffering.ModifiedDate = ReadField(vRS.Fields("ModifiedDate"))
            vSectionOffering.ModifiedBy = ReadField(vRS.Fields("ModifiedBy"))
            
            GetSectionOfferingByID = Success
        
        Else
            GetSectionOfferingByID = Failed
        End If
    Else
        GetSectionOfferingByID = Failed
    End If
    
ReleaseAndExit:
    Set vRS = Nothing
End Function

Public Function SectionOfferingExistByID(sSectionOfferingID As String) As TranDBResult
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    'default
    SectionOfferingExistByID = Failed
    
    If Len(sSectionOfferingID) < 1 Then Exit Function
    
    sSQL = " SELECT tblSectionOffering.SectionOfferingID" & _
            " From tblSectionOffering " & _
            " WHERE (((tblSectionOffering.SectionOfferingID)='" & sSectionOfferingID & "'));"
            
    If ConnectRS(HSESDB, vRS, sSQL) = True Then
        If AnyRecordExisted(vRS) = True Then
            SectionOfferingExistByID = Success
        Else
            SectionOfferingExistByID = Failed
        End If
    Else
        SectionOfferingExistByID = Failed
    End If
    
    Set vRS = Nothing
End Function

Public Function TeacherAssignedBySchoolYear(sTeacherID As String, sSchoolYearTitle As String) As TranDBResult
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    'default
    TeacherAssignedBySchoolYear = Failed
    
    If Len(sSchoolYearTitle) < 1 Or Len(sTeacherID) < 1 Then Exit Function
    
    sSQL = "SELECT tblSectionOffering.TeacherID, tblSectionOffering.SchoolYear" & _
            " From tblSectionOffering" & _
            " WHERE (((tblSectionOffering.TeacherID)='" & sTeacherID & "') AND ((tblSectionOffering.SchoolYear)='" & sSchoolYearTitle & "'));"
            

    If ConnectRS(HSESDB, vRS, sSQL) = True Then
        If AnyRecordExisted(vRS) = True Then
            TeacherAssignedBySchoolYear = Success
        Else
            TeacherAssignedBySchoolYear = Failed
        End If
    Else
        TeacherAssignedBySchoolYear = Failed
    End If
    
    Set vRS = Nothing
End Function


Public Function GetAutoSectionOffering(sSchoolYear As String, sDepartmentID As String, iYearLevelID As Integer, dStudentPrevAveGrade As Double, ByRef sReturnSectionOfferingID As String) As TranDBResult
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String

    sSQL = "SELECT tblSectionOffering.SectionOfferingID, Count(tblEnrolment.EnrolmentID) AS CountOfEnrolmentID, tblSectionOffering.MinGrade, tblSectionOffering.MaxGrade, tblSectionOffering.MaxStudentCount, ([tblSectionOffering]![MaxGrade]+[tblSectionOffering]![MinGrade]) AS GradeRank, tblSectionOffering.MaxGrade, tblSectionOffering.CreationDate, tblSectionOffering.SchoolYear, tblSection.DepartmentID, tblSection.YearLevelID" & _
            " FROM tblSection INNER JOIN (tblSectionOffering LEFT JOIN tblEnrolment ON tblSectionOffering.SectionOfferingID = tblEnrolment.SectionOfferingID) ON tblSection.SectionID = tblSectionOffering.SectionID" & _
            " GROUP BY tblSectionOffering.SectionOfferingID, tblSectionOffering.MinGrade, tblSectionOffering.MaxGrade, tblSectionOffering.MaxStudentCount, ([tblSectionOffering]![MaxGrade]+[tblSectionOffering]![MinGrade]), tblSectionOffering.MaxGrade, tblSectionOffering.CreationDate, tblSectionOffering.SchoolYear, tblSection.DepartmentID, tblSection.YearLevelID" & _
            " Having (((Count(tblEnrolment.EnrolmentID)) < [tblSectionOffering]![MaxStudentCount]) And ((tblSectionOffering.MinGrade) <= " & dStudentPrevAveGrade & ") And ((tblSectionOffering.MaxGrade) >= " & dStudentPrevAveGrade & ") And ((tblSectionOffering.SchoolYear) = '" & sSchoolYear & "') And ((tblSection.DepartmentID) = '" & sDepartmentID & "') And ((tblSection.YearLevelID) = " & iYearLevelID & "))" & _
            " ORDER BY ([tblSectionOffering]![MaxGrade]+[tblSectionOffering]![MinGrade]) DESC , tblSectionOffering.MaxGrade DESC , tblSectionOffering.CreationDate;"


    'Clipboard.SetText sSQL
    'defaults
    sReturnSectionOfferingID = ""
    GetAutoSectionOffering = Failed
    
    If ConnectRS(HSESDB, vRS, sSQL) = False Then
        'temp
        GetAutoSectionOffering = Failed
        MsgBox "error"
        GoTo ReleaseAndExit
    End If
    
    If AnyRecordExisted(vRS) = False Then
        GetAutoSectionOffering = Failed
        GoTo ReleaseAndExit
    End If
    
    'success
    sReturnSectionOfferingID = ReadField(vRS.Fields("SectionOfferingID"))
    GetAutoSectionOffering = Success
    

ReleaseAndExit:
    Set vRS = Nothing
End Function
