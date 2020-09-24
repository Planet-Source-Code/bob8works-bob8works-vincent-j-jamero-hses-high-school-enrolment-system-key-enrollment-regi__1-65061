Attribute VB_Name = "modRSFee"
Option Explicit


Public Const KeyFee = "fees"

Public Function GetNewFeeID() As Long
    
    Dim vRS As ADODB.Recordset
    Dim sSQL As String
    
    
    'default
    GetNewFeeID = -1
    
    
    sSQL = "SELECT Max([tblFee].[FeeID])+1 AS NewID" & _
            " FROM tblFee;"
        
    If ConnectRS(HSESDB, vRS, sSQL) = True Then
        GetNewFeeID = Val(ReadField(vRS.Fields("newid")))
            
        If GetNewFeeID < 1 Then
            GetNewFeeID = 1
        End If
        
    Else
        GetNewFeeID = -1
    End If


    Set vRS = Nothing
End Function


Public Function AddFee(FeeID As Long, Title As String, Description As String, Amount As Double, SchoolYear As String, DepartmentID As String, YearLevelID As Integer, CreationDate As Date, CreatedBy As String) As TranDBResult
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    sSQL = "SELECT tblFee.FeeID, tblFee.Title, tblFee.Description, tblFee.Amount, tblFee.SchoolYear, tblFee.DepartmentID, tblFee.YearLevelID, tblFee.CreationDate, tblFee.CreatedBy" & _
            " FROM tblFee;"
    
    If ConnectRS(HSESDB, vRS, sSQL) = True Then
        
        vRS.AddNew
        
        
        vRS.Fields("FeeID").Value = FeeID
        
        vRS.Fields("Title").Value = Title
        vRS.Fields("Description").Value = Description
        vRS.Fields("Amount").Value = Amount
        vRS.Fields("SchoolYear").Value = SchoolYear
        vRS.Fields("DepartmentID").Value = DepartmentID
        vRS.Fields("YearLevelID").Value = YearLevelID
        vRS.Fields("CreationDate").Value = CreationDate
        vRS.Fields("CreatedBy").Value = CreatedBy
                
        vRS.Update
        
        AddFee = Success
    Else
        AddFee = Failed
    End If
    
    Set vRS = Nothing
End Function


Public Function DeleteFee(sFeeID As String) As TranDBResult
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    sSQL = "DELETE *" & _
            " From tblFee" & _
            " WHERE (((tblFee.FeeID)=" & sFeeID & "));"
                
    If ConnectRS(HSESDB, vRS, sSQL) Then
        DeleteFee = Success
    Else
        DeleteFee = Failed
    End If
    
    'release
    Set vRS = Nothing

End Function
