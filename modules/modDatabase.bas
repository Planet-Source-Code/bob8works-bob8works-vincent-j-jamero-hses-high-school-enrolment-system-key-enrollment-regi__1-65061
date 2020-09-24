Attribute VB_Name = "modRSUser"
Option Explicit

Public Const sAdministratortitle = "Administrator"
Public Const sEncoderTitle = "Encoder"


Public Const sCanAddUser = "Can Add User"
Public Const sCanEditUser = "Can Edit User"
Public Const sCanDeleteUser = "Can Delete User"
Public Const sCanViewUser = "Can View User"

Public Const sCanAddSchoolYear = "Can Add School Year"
Public Const sCanDeleteSchoolYear = "Can Delete School Year"
Public Const sCanLockUnlockSchoolYear = "Can Lock/Unlock School Year"


Public Const sCanAddDepartment = "Can Add Department"
Public Const sCanEditDepartment = "Can Edit Department"
Public Const sCanDeleteDepartment = "Can Delete Department"

Public Const sCanAddSection = "Can Add Section"
Public Const sCanEditSection = "Can Edit Section"
Public Const sCanDeleteSection = "Can Delete Section"

Public Const sCanAddSectionOffering = "Can Add Section Offering"
Public Const sCanEditSectionOffering = "Can Edit Section Offering"
Public Const sCanDeleteSectionOffering = "Can Delete Section Offering"

Public Const sCanAddTeacher = "Can Add Teacher"
Public Const sCanEditTeacher = "Can Edit Teacher"
Public Const sCanDeleteTeacher = "Can Delete Teacher"

Public Const sCanAddFee = "Can Add Fee"
Public Const sCanEditFee = "Can Edit Fee"
Public Const sCanDeleteFee = "Can Delete Fee"

Public Const sCanAddCashier = "Can Add Cashier"
Public Const sCanEditCashier = "Can Edit Cashier"
Public Const sCanDeleteCashier = "Can Delete Cashier"

Public Const sCanModifyDropped = "Can Add/Remove Dropped Student"

Public Const sCanAddEnrolment = "Can Add Enrolment"
Public Const sCanDeleteEnrolment = "Can Delete Enrolment"
Public Const sCanModifyGraduate = "Can Add/Remove Graduate Student"
Public Const sCanModifyLeaved = "Can Add/Remove Leaving Student"

Public Const sCanAddStudent = "Can Add Student"
Public Const sCanEditStudent = "Can Edit Student"
Public Const sCanDeleteStudent = "Can Delete Student"

Public Const sCanAddCredential = "Can Add Credential"
Public Const sCanEditCredential = "Can Edit Credential"
Public Const sCanDeleteCredential = "Can Delete Credential"

Public Const sCanAddStudentCredential = "Can Add Student Credential"
Public Const sCanDeleteStudentCredential = "Can Delete Student Credential"





Public Const keyUser = "user"


'U S E R
'-----------------------------------------------------
Public Type User
    
    UserName As String
    Password As String
    FullName As String
    UserType As String
    CreationDate As Date
    DateModified As Date
    LastModifiedBy As String
    CreatedBy As String
    
    'misc
    OnLine As Boolean
    
End Type



Public Const CanAddUser = "Can Add User"
Public Const CanEditUser = "Can Edit User"
Public Const CanDeleteUser = "Can Delete User"
Public Const CanClearUserLog = "Can Clear User Log"
    
Public Const CanAddSchoolYear = "Can Add School Year"
Public Const CanEditSchoolYear = "Can Edit School Year"
Public Const CanDeleteSchoolYear = "Can Delete School Year"

'-----------------------------------------------------
















'U S E R Functions
'-----------------------------------------------------

'final
Public Function GetUserByName(ByRef vUser As User, sUserName As String) As TranDBResult
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    
    'default
    GetUserByName = Failed
    
    sSQL = "SELECT tblUser.UserName, tblUser.Password, tblUser.FullName, tblUser.UserType, tblUser.CreationDate, tblUser.DateModified, tblUser.LastModifiedBy, tblUser.CreatedBy" & _
            " From tblUser" & _
            " WHERE (((tblUser.UserName)='" & sUserName & "'));"


    If ConnectRS(HSESDB, vRS, sSQL) = True Then
        If AnyRecordExisted(vRS) = False Then
            GetUserByName = Failed
            Exit Function
        End If
    Else
        GetUserByName = Failed
        Exit Function
    End If
    
    
    'found, set vUser
    vUser.UserName = ReadField(vRS.Fields("username"))
    vUser.Password = ReadField(vRS.Fields("password"))
    vUser.FullName = ReadField(vRS.Fields("fullname"))
    vUser.UserType = ReadField(vRS.Fields("usertype"))
    vUser.CreatedBy = ReadField(vRS.Fields("createdby"))
    vUser.LastModifiedBy = ReadField(vRS.Fields("lastmodifiedby"))
    vUser.DateModified = ReadField(vRS.Fields("datemodified"))
    vUser.CreationDate = ReadField(vRS.Fields("creationdate"))

    
    
    'return success
    GetUserByName = Success
    
    Set vRS = Nothing
End Function




'final
Public Function AddUser(newUser As User) As TranDBResult
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    'default
    AddUser = Failed
    
    
    If IsUserExisted(newUser.UserName) = False Then
        
        sSQL = "SELECT * FROM tblUser"
        
        If ConnectRS(HSESDB, vRS, sSQL) = True Then


            vRS.AddNew
            
            'account info
            vRS.Fields("username").Value = newUser.UserName
            vRS.Fields("password").Value = newUser.Password
            vRS.Fields("fullname").Value = newUser.FullName
            vRS.Fields("UserType").Value = newUser.UserType
            vRS.Fields("creationdate").Value = newUser.CreationDate
            'vRS.Fields("DateModified").Value = newUser.DateModified
            'vRS.Fields("LastModifiedBy").Value = newUser.LastModifiedBy
            vRS.Fields("CreatedBy").Value = newUser.CreatedBy
            
    
            'save
            vRS.Update
            AddUser = Success
        
        Else
        
            AddUser = Failed
        End If
        
    Else
        
        AddUser = Failed
    End If
    
    Set vRS = Nothing
    
End Function



'final
Public Function EditUser(OldUser As User) As TranDBResult
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    'default
    EditUser = Failed
    
    
    If IsUserExisted(OldUser.UserName) = True Then
        
        sSQL = "SELECT * FROM tblUser WHERE UserName='" & OldUser.UserName & "'"
        
        If ConnectRS(HSESDB, vRS, sSQL) = True Then

           If AnyRecordExisted(vRS) = True Then
            
                vRS.MoveFirst
                'account info
                'vRS.Fields("username").Value = OldUser.UserName
                vRS.Fields("password").Value = OldUser.Password
                vRS.Fields("fullname").Value = OldUser.FullName
                vRS.Fields("UserType").Value = OldUser.UserType
                'vRS.Fields("creationdate").Value = OldUser.CreationDate
                vRS.Fields("DateModified").Value = OldUser.DateModified
                vRS.Fields("LastModifiedBy").Value = OldUser.LastModifiedBy
                'vRS.Fields("CreatedBy").Value = OldUser.CreatedBy
                
        
                'save
                vRS.Update
                EditUser = Success
            
            Else
            
                EditUser = Failed
            End If
        
        Else
        
            EditUser = Failed
        End If
        
    Else
        
        EditUser = Failed
    End If
    
    Set vRS = Nothing
    
End Function




'final
Public Function SaveAccess(sCreatedBy As String, sUserName As String, sAccessTitle() As String) As TranDBResult
    Dim i As Integer
    Dim vRS As Recordset
    Dim sSQL As String
    
    'default
    SaveAccess = Failed
    
    If IsUserExisted(sUserName) = False Then
        SaveAccess = UserNotExist
    End If
    
    sSQL = "SELECT tblUserAccess.UserName, tblUserAccess.AllowedTo, tblUserAccess.CreatedBy" & _
            " FROM tblUserAccess" & _
            " WHERE UserName='" & sUserName & "'"
            
    If ConnectRS(HSESDB, vRS, sSQL) = True Then
        For i = 0 To UBound(sAccessTitle)
            
            vRS.Find "AllowedTo='" & sAccessTitle(i) & "'"
            If vRS.EOF Or vRS.BOF Then
                vRS.AddNew
            End If
            vRS.Fields("UserName").Value = sUserName
            vRS.Fields("AllowedTo").Value = sAccessTitle(i)
            vRS.Fields("CreatedBy").Value = sCreatedBy
            vRS.Update
        Next
        
        SaveAccess = Success
    Else
        SaveAccess = Failed
    End If
    
    Set vRS = Nothing
End Function

'final
Private Function IsUserExisted(sUserName As String) As Boolean
    Dim vRS As Recordset
    Dim sSQL As String
    
    sSQL = "SELECT tblUser.UserName" & _
            " From tblUser" & _
            " WHERE (((tblUser.UserName)='" & sUserName & "'));"

    If ConnectRS(HSESDB, vRS, sSQL) = True Then
        If AnyRecordExisted(vRS) = True Then
            IsUserExisted = True
        Else
            IsUserExisted = False
        End If
    Else
        IsUserExisted = False
    End If
    
    Set vRS = Nothing
End Function


Public Function UserAllowedTo(sUserName As String, sAccessTitle As String) As Boolean
    
    Dim vRS As Recordset
    Dim sSQL As String
    
    
    
    'default
    UserAllowedTo = False
    
    sSQL = "SELECT tblUserAccess.UserName, tblUserAccess.AllowedTo" & _
        " From tblUserAccess" & _
        " WHERE (((tblUserAccess.UserName)='" & sUserName & "') AND ((tblUserAccess.AllowedTo)='" & sAccessTitle & "'));"

    If ConnectRS(HSESDB, vRS, sSQL) = True Then
        If AnyRecordExisted(vRS) = True Then
            UserAllowedTo = True
        Else
            UserAllowedTo = False
        End If
    Else
        UserAllowedTo = False
    End If
    
    
    
    Set vRS = Nothing
End Function

Public Function UserLogin(sUserName As String, Optional TimeIn As Date) As TranDBResult
    
    Dim vRS As Recordset
    Dim sSQL As String
    
    
    
    'default
    UserLogin = Failed
    
    sSQL = "SELECT *" & _
            " From tblLogRecord"

    If ConnectRS(HSESDB, vRS, sSQL) = True Then

            vRS.AddNew
            
            vRS.Fields("UserName").Value = sUserName
            vRS.Fields("Login").Value = TimeIn
            vRS.Fields("SuccessfullyOut").Value = False
            
            vRS.Update
            
            currentUserLog.TimeIn = TimeIn
            currentUserLog.SuccessfullyOut = False
            
            UserLogin = Success
            
            
    Else
        UserLogin = Failed
    End If
    
    
    
    Set vRS = Nothing
End Function

Public Function UserLogOut(sUserName As String, Optional TimeOut As Date, Optional bSuccessfullyOut As Boolean = True) As TranDBResult
    
    Dim vRS As Recordset
    Dim sSQL As String
    
    
    
    'default
    UserLogOut = Failed
    
    sSQL = "SELECT *" & _
            " From tblLogRecord" & _
            " WHERE (((tblLogRecord.UserName)='" & sUserName & "') AND ((tblLogRecord.SuccessfullyOut)=False))" & _
            " ORDER BY tblLogRecord.Login DESC;"
            
    If ConnectRS(HSESDB, vRS, sSQL) = True Then
        If AnyRecordExisted(vRS) = True Then
            'edit
            vRS.MoveFirst
            vRS.Fields("UserName").Value = sUserName
            vRS.Fields("Logout").Value = TimeOut
            vRS.Fields("SuccessfullyOut").Value = bSuccessfullyOut
            
            vRS.Update
            
            currentUserLog.TimeOut = TimeOut
            
            UserLogOut = Success
        Else
            UserLogOut = Failed
        End If
    Else
        UserLogOut = Failed
    End If
    
    
    
    Set vRS = Nothing
End Function

'final
Public Function UserOnline(sUserName As String) As Boolean
    
    Dim vRS As Recordset
    Dim sSQL As String
    
    
    
    'default
    UserOnline = False
    
    sSQL = "SELECT *" & _
            " From tblLogRecord" & _
            " WHERE (((tblLogRecord.UserName)='" & sUserName & "'))" & _
            " ORDER BY tblLogRecord.Login DESC;"
            
    If ConnectRS(HSESDB, vRS, sSQL) = True Then
        If AnyRecordExisted(vRS) = True Then
            
            UserOnline = IIf(vRS.Fields("SuccessfullyOut").Value = True, False, True)
            
        Else
            UserOnline = False
        End If
    Else
        UserOnline = False
    End If
    
    
    
    Set vRS = Nothing
End Function


Public Function UserRecordExist() As Boolean
    
    Dim vRS As Recordset
    Dim sSQL As String
    
    
    
    'default
    UserRecordExist = False
    
    sSQL = "SELECT * From tblUser"
            
    If ConnectRS(HSESDB, vRS, sSQL) = True Then
        If AnyRecordExisted(vRS) = True Then
            
            UserRecordExist = True
            
        Else
            UserRecordExist = False
        End If
    Else
        UserRecordExist = False
    End If
    
    
    
    Set vRS = Nothing
End Function



Public Function DeleteUser(sUserName As String) As TranDBResult
    
    Dim vRS As Recordset
    Dim sSQL As String
        
    'default
    DeleteUser = Failed
    
    sSQL = "DELETE * From tblUser WHERE UserName='" & sUserName & "'"
            
    If ConnectRS(HSESDB, vRS, sSQL) = True Then
        DeleteUser = Success
    Else
        DeleteUser = Failed
    End If
        
    Set vRS = Nothing
End Function
    
