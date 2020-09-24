Attribute VB_Name = "modRSCharge"
Option Explicit

Public Function AddCharge(sEnrolmentID As String, lFeeID As Long, sNote As String, dCreationDate As Date, sCreatedBy As String) As TranDBResult
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    Dim sChargeID As String
    
    'default
    AddCharge = Failed
    
    'generate ID
    sChargeID = sEnrolmentID & "-" & String$(10 - Len(Trim(lFeeID)), "0") & lFeeID

    If ChargeExistByID(sChargeID) = Success Then
        
        AddCharge = DuplicateID
        GoTo ReleaseAndExit
    End If
    
    sSQL = "SELECT * FROM tblCharge"
    
    If ConnectRS(HSESDB, vRS, sSQL) = False Then
        'fatal error here
        AddCharge = NotConnected
        GoTo ReleaseAndExit
    End If
    
    'add to record set
    vRS.AddNew
    
    vRS.Fields("ChargeID").Value = sChargeID
    vRS.Fields("FeeID").Value = lFeeID
    vRS.Fields("EnrolmentID").Value = sEnrolmentID
    vRS.Fields("Note").Value = sNote
    vRS.Fields("CreationDate").Value = dCreationDate
    vRS.Fields("CreatedBy").Value = sCreatedBy
    
    vRS.Update
    
    'set flag
    AddCharge = Success
    
    'release
ReleaseAndExit:
    Set vRS = Nothing
End Function

Public Function ChargeExistByID(sChargeID As String) As TranDBResult
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    'default
    ChargeExistByID = Failed
    
    
    sSQL = "SELECT * From tblCharge WHERE (((tblCharge.ChargeID)='" & sChargeID & "'));"
    
    If ConnectRS(HSESDB, vRS, sSQL) = True Then
        If AnyRecordExisted(vRS) = True Then
            ChargeExistByID = TranDBResult.Success
        Else
            ChargeExistByID = TranDBResult.Failed
        End If
            
    Else
        ChargeExistByID = Failed
    End If

End Function
