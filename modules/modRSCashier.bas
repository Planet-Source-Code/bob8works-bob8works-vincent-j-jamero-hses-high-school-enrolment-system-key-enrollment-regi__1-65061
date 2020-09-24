Attribute VB_Name = "modRSCashier"
Option Explicit

Public Const KeyCashier = "cash"

Public Type tCashier
    CashierID As Long
    LoginName As String
    Password As String
    FirstName As String
    MiddleName As String
    LastName As String
    Address As String
    ContactNumber As String
    CreationDate As Date
    CreatedBy As String
    ModifiedDate As Date
    ModifiedBy As String
End Type

Public Function GetNewCashierID() As Long
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    Dim lNewID As Long
    
    sSQL = "SELECT Max([tblCashier]![CashierID])+1 AS NewID" & _
            " FROM tblCashier;"

    'set default
    lNewID = 0
    
    If ConnectRS(HSESDB, vRS, sSQL) = True Then
        If IsNumeric(ReadField(vRS.Fields("NewID"))) Then
            
            lNewID = CLng(ReadField(vRS.Fields("NewID")))
            If lNewID < 1 Then
                lNewID = 1
            End If
        Else
            lNewID = 1
        End If
    Else
        'fatal error
        'temp
        MsgBox "Error"
        lNewID = -1
    End If
        
        While CashierExistByID(lNewID) = Success
            lNewID = lNewID + 1
        Wend
        
        'return
        GetNewCashierID = lNewID
    
    Set vRS = Nothing
End Function

Public Function CashierExistByID(lCashierID As Long) As TranDBResult
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    sSQL = "SELECT tblCashier.CashierID" & _
            " From tblCashier" & _
            " GROUP BY tblCashier.CashierID" & _
            " HAVING (((tblCashier.CashierID)=" & lCashierID & "));"

    If ConnectRS(HSESDB, vRS, sSQL) = True Then
        If AnyRecordExisted(vRS) = True Then
            CashierExistByID = Success
        Else
            CashierExistByID = Failed
        End If
    Else
        'fatal error
        'temp
        MsgBox "error"
        CashierExistByID = Failed
        
    End If
    
    Set vRS = Nothing
End Function

Public Function CashierExistByLoginName(sLoginName As String) As TranDBResult
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    sSQL = "SELECT tblCashier.LoginName" & _
            " From tblCashier" & _
            " WHERE tblCashier.LoginName='" & sLoginName & "'" & _
            " GROUP BY tblCashier.LoginName"

    If ConnectRS(HSESDB, vRS, sSQL) = True Then
        If AnyRecordExisted(vRS) = True Then
            CashierExistByLoginName = Success
        Else
            CashierExistByLoginName = Failed
        End If
    Else
        'fatal error
        'temp
        MsgBox "error"
        CashierExistByLoginName = Failed
        
    End If
    
    Set vRS = Nothing
End Function
Public Function AddCashier(CashierID As Long, LoginName As String, Password As String, FirstName As String, MiddleName As String, LastName As String, Address As String, ContactNumber As String, CreationDate As Date, CreatedBy As String) As TranDBResult
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String

    'check
    If CashierExistByID(CashierID) = Success Then
        AddCashier = DuplicateID
        Exit Function
    End If
    
    If CashierExistByLoginName(LoginName) = Success Then
        AddCashier = DuplicateLoginName
        Exit Function
    End If
    
    sSQL = "SELECT * FROM tblCashier"
    
    If ConnectRS(HSESDB, vRS, sSQL) = True Then
        
        vRS.AddNew

        vRS.Fields("CashierID").Value = CashierID
        vRS.Fields("LoginName").Value = LoginName
        vRS.Fields("Password").Value = Password
        vRS.Fields("FirstName").Value = FirstName
        vRS.Fields("MiddleName").Value = MiddleName
        vRS.Fields("LastName").Value = LastName
        vRS.Fields("Address").Value = Address
        vRS.Fields("ContactNumber").Value = ContactNumber
        vRS.Fields("CreationDate").Value = Now
        vRS.Fields("CreatedBy").Value = CurrentUser.UserName
        
        vRS.Update
        
        AddCashier = Success
    
    Else
        'fatal error
        MsgBox "error"
        AddCashier = Failed
    End If
    
    Set vRS = Nothing
End Function


Public Function GetCashierByID(lCashierID As Long, ByRef vCashier As tCashier) As TranDBResult
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    sSQL = "SELECT *" & _
            " From tblCashier" & _
            " WHERE (((tblCashier.CashierID)=" & lCashierID & "));"

    If ConnectRS(HSESDB, vRS, sSQL) = True Then
        If AnyRecordExisted(vRS) = True Then
        
            vCashier.CashierID = ReadField(vRS.Fields("CashierID"))
            vCashier.LoginName = ReadField(vRS.Fields("LoginName"))
            vCashier.Password = ReadField(vRS.Fields("Password"))
            vCashier.FirstName = ReadField(vRS.Fields("FirstName"))
            vCashier.MiddleName = ReadField(vRS.Fields("MiddleName"))
            vCashier.LastName = ReadField(vRS.Fields("LastName"))
            vCashier.Address = ReadField(vRS.Fields("Address"))
            vCashier.ContactNumber = ReadField(vRS.Fields("ContactNumber"))
            vCashier.CreationDate = ReadField(vRS.Fields("CreationDate"))
            vCashier.CreatedBy = ReadField(vRS.Fields("CreatedBy"))
            vCashier.ModifiedDate = ReadField(vRS.Fields("ModifiedDate"))
            vCashier.ModifiedBy = ReadField(vRS.Fields("ModifiedBy"))
            
            GetCashierByID = Success
        Else
            GetCashierByID = Failed
        End If
    Else
        'fatal error
        'temp
        MsgBox "error"
        GetCashierByID = Failed
        
    End If
    
    Set vRS = Nothing
End Function

