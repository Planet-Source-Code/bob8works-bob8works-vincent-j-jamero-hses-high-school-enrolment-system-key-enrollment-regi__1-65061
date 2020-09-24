Attribute VB_Name = "modRSSubjectOffering"
Option Explicit


Public Const KeySubjectOffering = "seof"

Public Type tSubjectOffering
    
    SubjectOfferingID As String
    SubjectID As String
    SectionOfferingID As String
    SchedTimeStart As String
    SchedTimeEnd As String
    TeacherID As String
    Days As String
    
    CreationDate As Date
    CreatedBy As String
    ModifiedDate As Date
    ModifiedBy As String
End Type



Public Function AddSubjectOffering(vSubjectOffering As tSubjectOffering) As TranDBResult
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    sSQL = "SELECT * FROM tblSubjectOffering"
    
    If SubjectOfferingExistByID(vSubjectOffering.SectionOfferingID) = Success Then
        AddSubjectOffering = DuplicateID
        GoTo ReleaseAndExit
    End If
    
    If ConnectRS(HSESDB, vRS, sSQL) = True Then
        
        vRS.AddNew
        

        
        vRS.Fields("SubjectOfferingID").Value = vSubjectOffering.SubjectOfferingID
        vRS.Fields("SubjectID").Value = vSubjectOffering.SubjectID
        vRS.Fields("SectionOfferingID").Value = vSubjectOffering.SectionOfferingID
        vRS.Fields("SchedTimeStart").Value = vSubjectOffering.SchedTimeStart
        vRS.Fields("SchedTimeEnd").Value = vSubjectOffering.SchedTimeEnd

        vRS.Fields("TeacherID").Value = vSubjectOffering.TeacherID
        vRS.Fields("Days").Value = vSubjectOffering.Days
        
        vRS.Fields("CreationDate").Value = vSubjectOffering.CreationDate
        vRS.Fields("CreatedBy").Value = vSubjectOffering.CreatedBy
        
        vRS.Update
        
        AddSubjectOffering = Success
    Else
        'fatal error
        AddSubjectOffering = Failed
    End If
    
    
ReleaseAndExit:
    Set vRS = Nothing
End Function

Public Function GetSubjectOffering()

End Function

Public Function SubjectOfferingExistByID(sSubjectOfferingID As String) As TranDBResult
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    sSQL = "SELECT tblSubjectOffering.SubjectOfferingID " & _
            " From tblSubjectOffering" & _
            " WHERE (((tblSubjectOffering.SubjectOfferingID)='" & sSubjectOfferingID & "'));"

    
    'default
    SubjectOfferingExistByID = Failed
    
    If ConnectRS(HSESDB, vRS, sSQL) = True Then
        If AnyRecordExisted(vRS) = True Then
            SubjectOfferingExistByID = Success
        Else
            SubjectOfferingExistByID = Failed
        End If
    Else
        'fatal error
        SubjectOfferingExistByID = Failed
    End If
    
    Set vRS = Nothing
End Function
