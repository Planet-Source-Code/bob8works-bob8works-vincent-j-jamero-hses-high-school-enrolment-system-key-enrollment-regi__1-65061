Attribute VB_Name = "modRSGraduate"
Option Explicit

Public Const KeyGraduate = "grad"

Public Type tGraduate
    StudentID As String
    SchoolYear As String
    DateGraduated As Date
    Note As String
    CreationDate As String
    CreatedBy As String
End Type


Public Function AddGraduate(vGraduate As tGraduate) As TranDBResult
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    'default
    AddGraduate = False
    
    'check
    If Len(vGraduate.StudentID) < 1 Then
        AddGraduate = InvalidID
        
        GoTo ReleaseAndExit
    End If
    
    
    sSQL = "SELECT * FROM tblGraduate WHERE tblGraduate.StudentID='" & vGraduate.StudentID & "'"
    
    'conect
    If ConnectRS(HSESDB, vRS, sSQL) = False Then
        'fatal error
        MsgBox "error"
        GoTo ReleaseAndExit
    End If
    
    'check duplication
    If getRecordCount(vRS) > 0 Then
        AddGraduate = DuplicateID
        GoTo ReleaseAndExit
    End If
    
    vRS.AddNew
    
    vRS.Fields("StudentID").Value = vGraduate.StudentID
    vRS.Fields("SchoolYear").Value = vGraduate.SchoolYear
    vRS.Fields("DateGraduated").Value = vGraduate.DateGraduated
    vRS.Fields("Note").Value = vGraduate.Note
    vRS.Fields("CreationDate").Value = vGraduate.CreationDate
    vRS.Fields("CreatedBy").Value = vGraduate.CreatedBy
    
    vRS.Update
    
    'return success
    AddGraduate = Success
    
ReleaseAndExit:
    Set vRS = Nothing
End Function


Public Function DeleteGraduate(sStudentID As String) As TranDBResult
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    On Error GoTo ReleaseAndExit
    
    sSQL = "Delete tblGraduate.StudentID" & _
        " From tblGraduate" & _
        " WHERE (((tblGraduate.StudentID)='" & sStudentID & "'));"

    'default
    DeleteGraduate = Failed
    
    If ConnectRS(HSESDB, vRS, sSQL) = False Then
        'temp
        MsgBox "error"
        GoTo ReleaseAndExit
    End If
    
    DeleteGraduate = Success
    
    
ReleaseAndExit:
    Set vRS = Nothing
    
End Function
