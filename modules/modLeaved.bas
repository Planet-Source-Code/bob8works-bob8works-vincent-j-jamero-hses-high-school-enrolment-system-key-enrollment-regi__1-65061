Attribute VB_Name = "modLeaved"
Option Explicit

Public Const KeyLeaved = "grad"

Public Type tLeaved
    StudentID As String
    SchoolYear As String
    DateLeaved As Date
    Note As String
    CreationDate As String
    CreatedBy As String
End Type


Public Function AddLeaved(vLeaved As tLeaved) As TranDBResult
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    'default
    AddLeaved = False
    
    'check
    If Len(vLeaved.StudentID) < 1 Then
        AddLeaved = InvalidID
        
        GoTo ReleaseAndExit
    End If
    
    
    sSQL = "SELECT * FROM tblLeaved WHERE tblLeaved.StudentID='" & vLeaved.StudentID & "'"
    
    'conect
    If ConnectRS(HSESDB, vRS, sSQL) = False Then
        'fatal error
        MsgBox "error"
        GoTo ReleaseAndExit
    End If
    
    'check duplication
    If getRecordCount(vRS) > 0 Then
        AddLeaved = DuplicateID
        GoTo ReleaseAndExit
    End If
    
    vRS.AddNew
    
    vRS.Fields("StudentID").Value = vLeaved.StudentID
    vRS.Fields("SchoolYear").Value = vLeaved.SchoolYear
    vRS.Fields("DateLeaved").Value = vLeaved.DateLeaved
    vRS.Fields("Note").Value = vLeaved.Note
    vRS.Fields("CreationDate").Value = vLeaved.CreationDate
    vRS.Fields("CreatedBy").Value = vLeaved.CreatedBy
    
    vRS.Update
    
    'return success
    AddLeaved = Success
    
ReleaseAndExit:
    Set vRS = Nothing
End Function


Public Function DeleteLeaved(sStudentID As String) As TranDBResult
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    On Error GoTo ReleaseAndExit
    
    sSQL = "Delete tblLeaved.StudentID" & _
        " From tblLeaved" & _
        " WHERE (((tblLeaved.StudentID)='" & sStudentID & "'));"

    'default
    DeleteLeaved = Failed
    
    If ConnectRS(HSESDB, vRS, sSQL) = False Then
        'temp
        MsgBox "error"
        GoTo ReleaseAndExit
    End If
    
    DeleteLeaved = Success
    
    
ReleaseAndExit:
    Set vRS = Nothing
    
End Function


