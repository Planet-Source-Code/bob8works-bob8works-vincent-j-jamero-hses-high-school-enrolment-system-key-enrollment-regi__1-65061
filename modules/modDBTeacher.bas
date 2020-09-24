Attribute VB_Name = "modDBTeacher"
Option Explicit




Public Const KeyTeacher = "teac"

Public Type tTeacher
    TeacherID As String
    TeacherTitle As String
    Password As String
    FirstName As String
    MiddleName As String
    LastName As String
    Address As String
    ContactNumber As String
    CreationDate As Date
End Type



Public Function TeacherRecordExisted() As TranDBResult
    Dim vRS As New ADODB.Recordset
    
    If CreateDefaultRSTeacher(vRS) = Success Then
        If AnyRecordExisted(vRS) = True Then
            TeacherRecordExisted = Success
        Else
            TeacherRecordExisted = Failed
        End If
    Else
        TeacherRecordExisted = Failed
    End If
    
    Set vRS = Nothing
End Function

Public Function AddTeacher(newTeacher As tTeacher) As TranDBResult
    
    Dim vRS As New ADODB.Recordset
    Dim oldTeacher As tTeacher
    
    
    'check duplicate id
    If TeacherExistByID(newTeacher.TeacherID) = Success Then
        AddTeacher = DuplicateID
        GoTo ReleaseAndExit
    End If
    
    If TeacherExistByTitle(newTeacher.TeacherTitle) = Success Then
        AddTeacher = DuplicateTitle
        GoTo ReleaseAndExit
    End If
    
   
        'check each field
        If Len(Trim(newTeacher.TeacherID)) < 1 Then
            AddTeacher = InvalidID
            GoTo ReleaseAndExit
        End If
        
        If Len(Trim(newTeacher.TeacherTitle)) < 1 Then
            AddTeacher = InvalidTitle
            GoTo ReleaseAndExit
        End If
        
        If Len(Trim(newTeacher.TeacherTitle)) < 1 Then
            AddTeacher = InvalidTeacherTitle
            GoTo ReleaseAndExit
        End If
        If Len(Trim(newTeacher.Password)) < 1 Then
            AddTeacher = InvalidTeacherPassword
            GoTo ReleaseAndExit
        End If
        If Len(Trim(newTeacher.FirstName)) < 1 Then
            AddTeacher = InvalidTeacherFirstName
            GoTo ReleaseAndExit
        End If
        If Len(Trim(newTeacher.MiddleName)) < 1 Then
            AddTeacher = InvalidTeacherMiddleName
            GoTo ReleaseAndExit
        End If
        If Len(Trim(newTeacher.LastName)) < 1 Then
            AddTeacher = InvalidTeacherLastName
            GoTo ReleaseAndExit
        End If
        If Len(Trim(newTeacher.Address)) < 1 Then
            AddTeacher = InvalidTeacherAddress
            GoTo ReleaseAndExit
        End If
        
        If ConnectRS(HSESDB, vRS, "SELECT * From tblTeacher WHERE (((tblTeacher.TeacherID)='" & newTeacher.TeacherID & "'));") Then
            
            vRS.AddNew
            vRS.Fields("TeacherID").Value = newTeacher.TeacherID
            vRS.Fields("TeacherTitle").Value = newTeacher.TeacherTitle
            vRS.Fields("Password").Value = newTeacher.Password
            vRS.Fields("FirstName").Value = newTeacher.FirstName
            vRS.Fields("MiddleName").Value = newTeacher.MiddleName
            vRS.Fields("LastName").Value = newTeacher.LastName
            vRS.Fields("Address").Value = newTeacher.Address
            vRS.Fields("ContactNumber").Value = newTeacher.ContactNumber
            'ignore creation date
    
            vRS.Update
            
            AddTeacher = Success
        Else
            AddTeacher = Failed
        End If


ReleaseAndExit:
    Set vRS = Nothing
End Function

Public Function EditTeacher(newTeacher As tTeacher) As TranDBResult
    
    Dim vRS As New ADODB.Recordset
    Dim oldTeacher As tTeacher
    
    If GetTeacherByID(newTeacher.TeacherID, oldTeacher) = Success Then
        
        If Trim(LCase(newTeacher.TeacherTitle)) <> LCase(Trim(oldTeacher.TeacherTitle)) Then
            'a new title
            'find duplicate title
            If TeacherExistByTitle(newTeacher.TeacherTitle) = Success Then
                'duplicate found
                EditTeacher = DuplicateTitle
                GoTo ReleaseAndExit
            End If
        Else

        End If
    
        'check each field
        If Len(Trim(newTeacher.TeacherTitle)) < 1 Then
            EditTeacher = InvalidTeacherTitle
            GoTo ReleaseAndExit
        End If
        If Len(Trim(newTeacher.Password)) < 1 Then
            EditTeacher = InvalidTeacherPassword
            GoTo ReleaseAndExit
        End If
        If Len(Trim(newTeacher.FirstName)) < 1 Then
            EditTeacher = InvalidTeacherFirstName
            GoTo ReleaseAndExit
        End If
        If Len(Trim(newTeacher.MiddleName)) < 1 Then
            EditTeacher = InvalidTeacherMiddleName
            GoTo ReleaseAndExit
        End If
        If Len(Trim(newTeacher.LastName)) < 1 Then
            EditTeacher = InvalidTeacherLastName
            GoTo ReleaseAndExit
        End If
        If Len(Trim(newTeacher.ContactNumber)) < 1 Then
            EditTeacher = InvalidTeacherContactNumber
            GoTo ReleaseAndExit
        End If
        If Len(Trim(newTeacher.Address)) < 1 Then
            EditTeacher = InvalidTeacherAddress
            GoTo ReleaseAndExit
        End If
        
        If ConnectRS(HSESDB, vRS, "SELECT * From tblTeacher WHERE (((tblTeacher.TeacherID)='" & newTeacher.TeacherID & "'));") Then
            vRS.Fields("TeacherTitle").Value = newTeacher.TeacherTitle
            vRS.Fields("Password").Value = newTeacher.Password
            vRS.Fields("FirstName").Value = newTeacher.FirstName
            vRS.Fields("MiddleName").Value = newTeacher.MiddleName
            vRS.Fields("LastName").Value = newTeacher.LastName
            vRS.Fields("Address").Value = newTeacher.Address
            vRS.Fields("ContactNumber").Value = newTeacher.ContactNumber
            'ignore creation date
    
            vRS.Update
            
            EditTeacher = Success
        Else
            EditTeacher = Failed
        End If
    Else
        'teacher by id not found
        EditTeacher = InvalidID
    End If
    
ReleaseAndExit:
    Set vRS = Nothing
End Function


Public Function ExecDeleteTeacher(sTeacherID As String) As TranDBResult
        
            If MsgBox("You are about to delete this Teacher with ID :" & vbNewLine & sTeacherID & vbNewLine & "Are you sure to DELETE this Teahcer Account?", vbQuestion + vbOKCancel) = vbOK Then
                
                If DeleteTeacher(sTeacherID) = Success Then
                    MsgBox "TEACHER entry successfully deleted.", vbInformation
                    ExecDeleteTeacher = Success
                Else
                    MsgBox "Unable to delete Teacher Account. The current was edited by another user", vbExclamation
                    ExecDeleteTeacher = Failed
                End If
            Else
                ExecDeleteTeacher = Failed
            End If
        
End Function


Public Function DeleteTeacher(sTeacherID As String) As TranDBResult
    Dim vRS As New ADODB.Recordset
    
    If ConnectRS(HSESDB, vRS, "DELETE tblTeacher.TeacherID From tblTeacher WHERE (((tblTeacher.TeacherID)='" & sTeacherID & "'));") Then
        DeleteTeacher = Success
    Else
        DeleteTeacher = Failed
    End If

    Set vRS = Nothing
End Function



Public Function GetTeacherMoveNext(ByRef vRS As ADODB.Recordset, ByRef vTeacher As tTeacher) As TranDBResult
    

    
    If Not vRS.EOF And Not vRS.BOF Then
        
        
        vTeacher.TeacherID = ReadField(vRS.Fields("teacherid"))
        vTeacher.TeacherTitle = ReadField(vRS.Fields("TeacherTitle"))
        vTeacher.Password = ReadField(vRS.Fields("Password"))
        vTeacher.FirstName = ReadField(vRS.Fields("FirstName"))
        vTeacher.MiddleName = ReadField(vRS.Fields("MiddleName"))
        vTeacher.LastName = ReadField(vRS.Fields("LastName"))
        vTeacher.Address = ReadField(vRS.Fields("Address"))
        vTeacher.ContactNumber = ReadField(vRS.Fields("ContactNumber"))
        vTeacher.CreationDate = ReadField(vRS.Fields("CreationDate"))

        
        vRS.MoveNext
        
        GetTeacherMoveNext = Success
        
    Else
    
        GetTeacherMoveNext = Failed
        
    End If
    
End Function


Public Function GetTeacherByTitle(sTeacherTitle As String, ByRef vTeacher As tTeacher) As TranDBResult
        
    Dim vRS As New ADODB.Recordset
    
    If ConnectRS(HSESDB, vRS, "SELECT * From tblTeacher WHERE (((tblTeacher.TeacherTitle)='" & sTeacherTitle & "'));") Then
        If AnyRecordExisted(vRS) Then
            vTeacher.TeacherID = ReadField(vRS.Fields("teacherid"))
            vTeacher.TeacherTitle = ReadField(vRS.Fields("TeacherTitle"))
            vTeacher.Password = ReadField(vRS.Fields("Password"))
            vTeacher.FirstName = ReadField(vRS.Fields("FirstName"))
            vTeacher.MiddleName = ReadField(vRS.Fields("MiddleName"))
            vTeacher.LastName = ReadField(vRS.Fields("LastName"))
            vTeacher.Address = ReadField(vRS.Fields("Address"))
            vTeacher.ContactNumber = ReadField(vRS.Fields("ContactNumber"))
            vTeacher.CreationDate = ReadField(vRS.Fields("CreationDate"))
            
            GetTeacherByTitle = Success
        Else
            GetTeacherByTitle = Failed
        End If
    Else
        
        GetTeacherByTitle = Failed
    End If
    
    Set vRS = Nothing
        
End Function

Public Function GetTeacherByFullName(sTeacherFullName As String, ByRef vTeacher As tTeacher) As TranDBResult
        
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    
    
    sSQL = "SELECT *" & _
            " From tblTeacher" & _
            " WHERE ((([LastName] & ', ' & [FirstName] & ' ' & [MiddleName])='" & sTeacherFullName & "'));"

    If ConnectRS(HSESDB, vRS, sSQL) Then
        If AnyRecordExisted(vRS) Then
            vTeacher.TeacherID = ReadField(vRS.Fields("teacherid"))
            vTeacher.TeacherTitle = ReadField(vRS.Fields("TeacherTitle"))
            vTeacher.Password = ReadField(vRS.Fields("Password"))
            vTeacher.FirstName = ReadField(vRS.Fields("FirstName"))
            vTeacher.MiddleName = ReadField(vRS.Fields("MiddleName"))
            vTeacher.LastName = ReadField(vRS.Fields("LastName"))
            vTeacher.Address = ReadField(vRS.Fields("Address"))
            vTeacher.ContactNumber = ReadField(vRS.Fields("ContactNumber"))
            vTeacher.CreationDate = ReadField(vRS.Fields("CreationDate"))
            
            GetTeacherByFullName = Success
        Else
            GetTeacherByFullName = Failed
        End If
    Else
        
        GetTeacherByFullName = Failed
    End If
    
    Set vRS = Nothing
        
End Function

Public Function GetTeacherByID(sTeacherID As String, ByRef vTeacher As tTeacher) As TranDBResult
    
    Dim vRS As New ADODB.Recordset
    
    If ConnectRS(HSESDB, vRS, "SELECT * From tblTeacher WHERE (((tblTeacher.TeacherID)='" & sTeacherID & "'));") Then
        If AnyRecordExisted(vRS) Then
            vTeacher.TeacherID = ReadField(vRS.Fields("teacherid"))
            vTeacher.TeacherTitle = ReadField(vRS.Fields("TeacherTitle"))
            vTeacher.Password = ReadField(vRS.Fields("Password"))
            vTeacher.FirstName = ReadField(vRS.Fields("FirstName"))
            vTeacher.MiddleName = ReadField(vRS.Fields("MiddleName"))
            vTeacher.LastName = ReadField(vRS.Fields("LastName"))
            vTeacher.Address = ReadField(vRS.Fields("Address"))
            vTeacher.ContactNumber = ReadField(vRS.Fields("ContactNumber"))
            vTeacher.CreationDate = ReadField(vRS.Fields("CreationDate"))
            
            GetTeacherByID = Success
        Else
            GetTeacherByID = Failed
        End If
    Else
        
        GetTeacherByID = Failed
    End If
    
    Set vRS = Nothing
End Function















Public Function TeacherExistByTitle(sTeacherTitle As String) As TranDBResult
    
    Dim vRS As New ADODB.Recordset
        
    If ConnectRS(HSESDB, vRS, "SELECT * From tblTeacher WHERE (((tblTeacher.TeacherTitle)='" & sTeacherTitle & "'));") Then
        If vRS.RecordCount > 0 Then
            TeacherExistByTitle = Success
        Else
            TeacherExistByTitle = Failed
        End If
    End If
    
    'release
    Set vRS = Nothing
End Function

Public Function TeacherExistByFullName(sTeacherFullName As String) As TranDBResult
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    'deafult
    TeacherExistByFullName = Failed
    
    
    If Len(sTeacherFullName) < 1 Then Exit Function
    
    sSQL = "SELECT [LastName] & ', ' & [FirstName] & ' ' & [MiddleName] AS TeacherFullName" & _
            " From tblTeacher" & _
            " WHERE ((([LastName] & ', ' & [FirstName] & ' ' & [MiddleName])='" & sTeacherFullName & "'));"

    If ConnectRS(HSESDB, vRS, sSQL) Then
        If vRS.RecordCount > 0 Then
            TeacherExistByFullName = Success
        Else
            TeacherExistByFullName = Failed
        End If
    End If
    
    'release
    Set vRS = Nothing
End Function

Public Function TeacherExistByID(sTeacherID As String) As TranDBResult
    
    Dim vRS As New ADODB.Recordset
        
    If ConnectRS(HSESDB, vRS, "SELECT * From tblTeacher WHERE (((tblTeacher.TeacherID)='" & sTeacherID & "'));") Then
        If vRS.RecordCount > 0 Then
            TeacherExistByID = Success
        Else
            TeacherExistByID = Failed
        End If
    End If
    
    'release
    Set vRS = Nothing
End Function


Public Function CreateDefaultRSTeacher(ByRef vRS As ADODB.Recordset) As TranDBResult
    'default
    CreateDefaultRSTeacher = Failed
    
    If ConnectRS(HSESDB, vRS, "SELECT * FROM tblTeacher") Then
        CreateDefaultRSTeacher = Success
    End If
End Function

Public Function TeacherRecordExist() As TranDBResult
    Dim vRS As New ADODB.Recordset
    
    If CreateDefaultRSTeacher(vRS) = Success Then
        If AnyRecordExisted(vRS) Then
            TeacherRecordExist = Success
        Else
            TeacherRecordExist = Failed
        End If
    Else
        TeacherRecordExist = Failed
    End If
    
    Set vRS = Nothing
End Function




Public Function TeacherLogin(sTeacherTitle As String, ByRef vLogRec As LogRec) As TranDBResult
    
    'declaration
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    '----------------------------------------------------------
    
    sSQL = "SELECT tblTeacherLog.TimeIn, tblTeacherLog.TimeOut, tblTeacherLog.SuccessfullyOut, tblTeacher.TeacherTitle" & _
            " FROM tblTeacher INNER JOIN tblTeacherLog ON tblTeacher.TeacherTitle = tblTeacherLog.TeacherTitle" & _
            " Where (((tblTeacher.TeacherTitle) = '" & sTeacherTitle & "'))" & _
            " ORDER BY tblTeacherLog.SuccessfullyOut DESC"
            

    If ConnectRS(HSESDB, vRS, sSQL) = True Then
        If AnyRecordExisted(vRS) = True Then
            
            
            'check if already login
            vRS.MoveFirst
            
            
            
            If ReadField(vRS.Fields("SuccessfullyOut")) = True Then
                If AddTeacherLogin(vLogRec) = Success Then
                        
                    TeacherLogin = SuccessIn
                Else
                    TeacherLogin = Failed
                End If
            Else
                'all ready login
                TeacherLogin = AlreadyLogIn
                
                vLogRec.UserName = ReadField(vRS.Fields("TeacherTitle"))
                vLogRec.TimeIn = ReadField(vRS.Fields("TimeIn"))
                vLogRec.TimeOut = ReadField(vRS.Fields("TimeOut"))
                vLogRec.SuccessfullyOut = ReadField(vRS.Fields("SuccessfullyOut"))

            End If
            
            
        Else
            'no record found
            'add log
            If AddTeacherLogin(vLogRec) = Success Then
                    
                TeacherLogin = SuccessIn
            Else
                TeacherLogin = Failed
            End If
        End If
    Else
        TeacherLogin = Failed
    End If
    
    
    
    'release
    '----------------------------------------------------------
    Set vRS = Nothing
End Function


Private Function AddTeacherLogin(vTeacherLog As LogRec) As TranDBResult
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    '------------------------------------------------------
    
    
    sSQL = "SELECT * from tblTeacherLog"
    If ConnectRS(HSESDB, vRS, sSQL) = True Then
        vRS.AddNew
        
        vRS.Fields("TeacherTitle").Value = vTeacherLog.UserName
        vRS.Fields("TimeIn").Value = vTeacherLog.TimeIn
        vRS.Fields("TimeOut").Value = vTeacherLog.TimeOut
        vRS.Fields("SuccessfullyOut").Value = False
        
        vRS.Update
        AddTeacherLogin = Success
    Else
        AddTeacherLogin = Failed
    End If
    
    '------------------------------------------------------
    Set vRS = Nothing
End Function

Public Function UpdateTeacherLogin(vTeacherLog As LogRec) As TranDBResult
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    '------------------------------------------------------
    
    
    sSQL = "SELECT tblTeacherLog.TeacherTitle, tblTeacherLog.TimeOut, tblTeacherLog.SuccessfullyOut" & _
            " From tblTeacherLog" & _
            " WHERE (((tblTeacherLog.TeacherTitle)='" & vTeacherLog.UserName & "') AND ((tblTeacherLog.SuccessfullyOut)=False));"

    If ConnectRS(HSESDB, vRS, sSQL) = True Then
        
        vRS.MoveFirst
        
        'vRS.Fields("TeacherTitle").Value = vTeacherLog.UserName
        'vRS.Fields("Timer").Value = vTeacherLog
        vRS.Fields("TimeOut").Value = vTeacherLog.TimeOut
        vRS.Fields("SuccessfullyOut").Value = vTeacherLog.SuccessfullyOut
        
        vRS.Update
        UpdateTeacherLogin = Success
    Else
        UpdateTeacherLogin = Failed
    End If
    
    '------------------------------------------------------
    Set vRS = Nothing
End Function


Public Function TeacherLogOut(vTeacherLog As LogRec) As TranDBResult
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    '------------------------------------------------------
    
    
    sSQL = "SELECT * " & _
            " From tblTeacherLog" & _
            " WHERE (((tblTeacherLog.TeacherTitle) = '" & vTeacherLog.UserName & "') And ((tblTeacherLog.SuccessfullyOut) = False ))" & _
            " ORDER BY tblTeacherLog.TimeOut DESC;"
   
    If ConnectRS(HSESDB, vRS, sSQL) = True Then
        
        vRS.MoveFirst
        
        vRS.Fields("TimeOut").Value = vTeacherLog.TimeOut
        vRS.Fields("SuccessfullyOut").Value = True
        
        vRS.Update
        TeacherLogOut = Success
    Else
        TeacherLogOut = Failed
    End If
    
    '------------------------------------------------------
    Set vRS = Nothing
End Function

Public Function GetTeacherLogStatus(sTeacherTitle As String, ByRef dLastTimeOut As Date, ByRef bSuccessfullOut As Boolean) As TranDBResult
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    '------------------------------------------------------
    
    
    sSQL = "SELECT tblTeacherLog.TeacherTitle, tblTeacherLog.TimeOut, tblTeacherLog.SuccessfullyOut" & _
            " From tblTeacherLog" & _
            " Where (((tblTeacherLog.TeacherTitle) = '" & sTeacherTitle & "'))" & _
            " ORDER BY tblTeacherLog.SuccessfullyOut DESC;"
    
    If ConnectRS(HSESDB, vRS, sSQL) = True Then
        
        vRS.MoveFirst
        
        dLastTimeOut = vRS.Fields("TimeOut").Value
        bSuccessfullOut = vRS.Fields("SuccessfullyOut").Value

        GetTeacherLogStatus = Success
    Else
        GetTeacherLogStatus = Failed
    End If
    
    '------------------------------------------------------
    Set vRS = Nothing
End Function


Public Function GetNewTeacherID(ByRef sNewTeacherID As String) As TranDBResult
    
    Dim vRS As New ADODB.Recordset
    Dim lNewNumber As Integer
    Dim sSQL As String


    'default
    sNewTeacherID = ""
    GetNewTeacherID = Failed
    sSQL = "SELECT 'TN-' & String$(7-Len(Max(Val(Right([tblTeacher].[TeacherID],7)))+1),'0') & Max(Val(Right([tblTeacher].[TeacherID],7)))+1 AS sNewID" & _
            " FROM tblTeacher;"


    If ConnectRS(HSESDB, vRS, sSQL) = True Then
        If AnyRecordExisted(vRS) = False Then
            sNewTeacherID = "SN-0000001"
            GetNewTeacherID = Success
            GoTo ReleaseAndExit
        End If
    Else
        'fatal error
        GetNewTeacherID = Failed
        GoTo ReleaseAndExit
    End If
    
    sNewTeacherID = ReadField(vRS.Fields("snewid"))
    lNewNumber = 0
    While TeacherExistByID(sNewTeacherID) = Success
        If IsNumeric(Right(sNewTeacherID, 7)) = True Then
            lNewNumber = Val(Right(sNewTeacherID, 7)) + 1
        Else
            lNewNumber = 1
        End If
        sNewTeacherID = "TN-" & String$(7 - Len(lNewNumber), "0") & lNewNumber
    Wend
    
    GetNewTeacherID = Success


ReleaseAndExit:
    Set vRS = Nothing
End Function
