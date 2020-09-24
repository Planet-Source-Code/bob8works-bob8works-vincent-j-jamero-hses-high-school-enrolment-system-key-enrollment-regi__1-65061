Attribute VB_Name = "modRSDropped"
Option Explicit

Public Const KeyDropped = "grad"

Public Type tDropped
    StudentID As String
    SchoolYear As String
    DateDropped As Date
    Note As String
    CreationDate As String
    CreatedBy As String
End Type


Public Function AddDropped(vDropped As tDropped) As TranDBResult
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    'default
    AddDropped = False
    
    'check
    If Len(vDropped.StudentID) < 1 Then
        AddDropped = InvalidID
        
        GoTo ReleaseAndExit
    End If
    
    
    sSQL = "SELECT * FROM tblDropped WHERE tblDropped.StudentID='" & vDropped.StudentID & "'"
    
    'conect
    If ConnectRS(HSESDB, vRS, sSQL) = False Then
        'fatal error
        MsgBox "error"
        GoTo ReleaseAndExit
    End If
    
    'check duplication
    If getRecordCount(vRS) > 0 Then
        AddDropped = DuplicateID
        GoTo ReleaseAndExit
    End If
    
    vRS.AddNew
    
    vRS.Fields("StudentID").Value = vDropped.StudentID
    vRS.Fields("SchoolYear").Value = vDropped.SchoolYear
    vRS.Fields("DateDropped").Value = vDropped.DateDropped
    vRS.Fields("Note").Value = vDropped.Note
    vRS.Fields("CreationDate").Value = vDropped.CreationDate
    vRS.Fields("CreatedBy").Value = vDropped.CreatedBy
    
    vRS.Update
    
    'return success
    AddDropped = Success
    
ReleaseAndExit:
    Set vRS = Nothing
End Function


Public Function IsStudentDropped(sStudentID As String) As TranDBResult
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    'set default
    IsStudentDropped = Failed
    
    sSQL = "SELECT tblDropped.StudentID" & _
            " FROM tblDropped INNER JOIN tblStudent ON tblDropped.StudentID = tblStudent.StudentID" & _
            " Where (((tblStudent.StudentID) = '" & sStudentID & "'))" & _
            " GROUP BY tblDropped.StudentID;"
    
    If ConnectRS(HSESDB, vRS, sSQL) = False Then
        MsgBox "Unable to connect Recordset.", vbCritical
        CatchError "modRSDropped", "IsStudentDropped", "Unable to connect Recordset. Possible cause: Invalid Sql Statement."
        GoTo ReleaseAndExit
    End If
    
    If AnyRecordExisted(vRS) = True Then
        IsStudentDropped = Success
    Else
        IsStudentDropped = Failed
    End If
    
ReleaseAndExit:
    Set vRS = Nothing
End Function



Public Function DeleteDropped(sStudentID As String) As TranDBResult
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    On Error GoTo ReleaseAndExit
    
    sSQL = "Delete tblDropped.StudentID" & _
        " From tblDropped" & _
        " WHERE (((tblDropped.StudentID)='" & sStudentID & "'));"

    'default
    DeleteDropped = Failed
    
    If ConnectRS(HSESDB, vRS, sSQL) = False Then
        'temp
        MsgBox "error"
        GoTo ReleaseAndExit
    End If
    
    DeleteDropped = Success
    
    
ReleaseAndExit:
    Set vRS = Nothing
    
End Function


