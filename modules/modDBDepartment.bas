Attribute VB_Name = "modDBDepartment"
Option Explicit

Public Const KeyDepartment = "dept"

Public Type tDepartment
    DepartmentID As String
    DepartmentTitle As String
End Type







Public Function AddDepartment(newDepartment As tDepartment) As TranDBResult
    'possibe return values
        'Success
        'IDNotFound
        'DuplicateTitle
    
    Dim vRS As New ADODB.Recordset
    
    'find duplicate ID
    If DepartmentExistByID(newDepartment.DepartmentID) = Success Then
        AddDepartment = DuplicateID
        GoTo ReleaseAndExit
    End If
    
    'find duplicate TITLE
    If DepartmentExistByTitle(newDepartment.DepartmentTitle) = Success Then
        AddDepartment = DuplicateTitle
        GoTo ReleaseAndExit
    End If
    
    If CreateDefaultRSDepartment(vRS) = Success Then
        'add new record
        vRS.AddNew
        vRS.Fields("DepartmentID").Value = newDepartment.DepartmentID
        vRS.Fields("DepartmentTitle").Value = newDepartment.DepartmentTitle
        vRS.Update
        AddDepartment = Success
    Else
        AddDepartment = NotConnected
    End If
    
    
    
ReleaseAndExit:
    'release
    Set vRS = Nothing
End Function

Public Function EditDepartment(newDepartment As tDepartment) As TranDBResult
    'possibe return values
        'Success
        'InvalidID
        'DuplicateTitle
    
    Dim oldDepartment As tDepartment

    Dim vRS As New ADODB.Recordset
    

    'check duplicate title
    If GetDepartmentByID(newDepartment.DepartmentID, oldDepartment) Then
        If oldDepartment.DepartmentTitle = newDepartment.DepartmentTitle Then
            'there is nothing to do, no fields changes in NEW DEPARTMENT
            'return success
            EditDepartment = Success
            'exit function
            GoTo ReleaseAndExit
        Else
            'find duplicate title
            If DepartmentExistByTitle(newDepartment.DepartmentTitle) = Success Then
                EditDepartment = DuplicateTitle
                'exit function
                GoTo ReleaseAndExit
            End If
        End If
    Else
        'department not found
        'exit function
        EditDepartment = InvalidID
        GoTo ReleaseAndExit
    End If
    

    'find record to edit

    If ConnectRS(HSESDB, vRS, "SELECT * From tblDepartment WHERE (((tblDepartment.DepartmentID)='" & newDepartment.DepartmentID & "'));") Then
        If vRS.RecordCount < 1 Then
            EditDepartment = InvalidID
            GoTo ReleaseAndExit
        End If
    End If
    
      
        'vrs'editing
        vRS.Fields("Departmenttitle").Value = newDepartment.DepartmentTitle
        vRS.Update
            
        EditDepartment = Success
        

ReleaseAndExit:
    'release
    Set vRS = Nothing
End Function








Public Function ExecuteDeleteDepartment(sDepartmentID As String) As TranDBResult

      'check if record exist and if it is edited by other user
    If MsgBox("WARNING:" & vbNewLine & _
        "Deleting this DEPARTMENT entry will affect all other record" & vbNewLine & vbNewLine & _
        "Delete this record anyway?", vbQuestion + vbYesNo) = vbYes Then
            
        If DeleteDepartment(sDepartmentID) = Success Then
            MsgBox "DEPARTMENT entry and other related record succesfully deleted.", vbInformation
            ExecuteDeleteDepartment = Success
        Else
            MsgBox "Deleting DEPARTMENT entry went failed.", vbExclamation
            ExecuteDeleteDepartment = Failed
        End If
    Else
        ExecuteDeleteDepartment = Failed
    End If
End Function

Public Function DeleteDepartment(sDepartmentID As String) As TranDBResult
    
    Dim vRS As New ADODB.Recordset
    
    
    'default
    DeleteDepartment = Failed
    
    If ConnectRS(HSESDB, vRS, "Delete * From tblDepartment WHERE (((tblDepartment.DepartmentID)='" & sDepartmentID & "'));") Then
        DeleteDepartment = Success
    Else
        DeleteDepartment = Failed
    End If
    
    'release
    Set vRS = Nothing
End Function




Public Function GetDepartmentMoveNext(ByRef vRS As ADODB.Recordset, ByRef vDepartment As tDepartment) As TranDBResult

    'asuming that vRS is already connected
    If Not vRS.EOF And Not vRS.BOF Then
    
        'SUCCESS: Record exist
        'get values
        '----------------------------------------------------------------
        vDepartment.DepartmentID = ReadField(vRS.Fields("DepartmentID"))
        vDepartment.DepartmentTitle = ReadField(vRS.Fields("DepartmentTitle"))
        'move to the next record
        vRS.MoveNext
        'return true
        GetDepartmentMoveNext = Success
    Else
        GetDepartmentMoveNext = Failed
    End If
    
End Function



Public Function GetDepartmentByID(sDepartmentID As String, ByRef vDepartment As tDepartment) As TranDBResult
    
    Dim vRS As New ADODB.Recordset

    If ConnectRS(HSESDB, vRS, "SELECT * From tblDepartment WHERE (((tblDepartment.DepartmentID)='" & sDepartmentID & "'));") Then
        If AnyRecordExisted(vRS) Then
            'SUCCESS: Record exist
            'get values
            '----------------------------------------------------------------
            vDepartment.DepartmentID = ReadField(vRS.Fields("DepartmentID"))
            vDepartment.DepartmentTitle = ReadField(vRS.Fields("DepartmentTitle"))
            
            GetDepartmentByID = Success
        
        Else

            'FAILED: record does not exist
            GetDepartmentByID = Failed
        End If
    Else
        GetDepartmentByID = Failed
    End If
    
    'release
    Set vRS = Nothing
End Function

Public Function GetDepartmentByTitle(sDepartmentTitle As String, ByRef vDepartment As tDepartment) As TranDBResult

    Dim vRS As New ADODB.Recordset
        
    If ConnectRS(HSESDB, vRS, "SELECT *  FROM tblDepartment WHERE (((tblDepartment.DepartmentTitle)='" & sDepartmentTitle & "'));") Then
        If vRS.RecordCount > 0 Then
            vDepartment.DepartmentID = ReadField(vRS.Fields("DepartmentID"))
            vDepartment.DepartmentTitle = ReadField(vRS.Fields("DepartmentTitle"))
            
            GetDepartmentByTitle = Success
        Else
            GetDepartmentByTitle = Failed
        End If
    Else
        GetDepartmentByTitle = Failed
    End If
    
    'release
    Set vRS = Nothing
End Function


Private Sub ReadFromRecord(ByRef vRS As ADODB.Recordset, ByRef vDepartment As tDepartment)
    
    vDepartment.DepartmentID = vRS.Fields("Departmentid").Value
    vDepartment.DepartmentTitle = vRS.Fields("Departmenttitle").Value

End Sub


Public Function DepartmentExistByTitle(sDepartmentTitle As String) As TranDBResult
    
    Dim vRS As New ADODB.Recordset
        
    If ConnectRS(HSESDB, vRS, "SELECT * From tblDepartment WHERE (((tblDepartment.DepartmentTitle)='" & sDepartmentTitle & "'));") Then
        If vRS.RecordCount > 0 Then
            DepartmentExistByTitle = Success
        Else
            DepartmentExistByTitle = Failed
        End If
    Else
        DepartmentExistByTitle = Failed
    End If
    
    'release
    Set vRS = Nothing
End Function


Public Function DepartmentExistByID(sDepartmentID As String) As TranDBResult
    
    Dim vRS As New ADODB.Recordset
        
    If ConnectRS(HSESDB, vRS, "SELECT * From tblDepartment WHERE (((tblDepartment.DepartmentID)='" & sDepartmentID & "'));") Then
        If vRS.RecordCount > 0 Then
            DepartmentExistByID = Success
        Else
            DepartmentExistByID = Failed
        End If
    Else
        DepartmentExistByID = Failed
       
    End If
    
    'release
    Set vRS = Nothing
End Function


Public Function CreateDefaultRSDepartment(ByRef vRS As ADODB.Recordset) As TranDBResult
    'default
    CreateDefaultRSDepartment = Failed
    
    If ConnectRS(HSESDB, vRS, "SELECT * FROM tblDepartment") Then
        CreateDefaultRSDepartment = Success
    End If
End Function

Public Function DepartmentRecordExist() As TranDBResult
    Dim vRS As New ADODB.Recordset
    
    If CreateDefaultRSDepartment(vRS) = Success Then
        
        If AnyRecordExisted(vRS) = True Then
            DepartmentRecordExist = Success
        Else
            DepartmentRecordExist = Failed
        End If
        
    Else
        DepartmentRecordExist = Failed
    End If
    
    Set vRS = Nothing
End Function

Public Function GetNewDepartmentID(ByRef sNewDepartmentID As String) As TranDBResult
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    Dim NewDNumber As Integer
    
    
    'default
    GetNewDepartmentID = Failed
    
    sSQL = "SELECT 'D-' & String$(2-Len(Count(*)+1),'0') & Count(*)+1 AS NewID" & _
            " FROM tblDepartment;"
            
    If ConnectRS(HSESDB, vRS, sSQL) = True Then
        
        sNewDepartmentID = vRS.Fields("NewID").Value
        
        While DepartmentExistByID(sNewDepartmentID) = Success
            NewDNumber = Val(Right(sNewDepartmentID, 2)) + 1
            sNewDepartmentID = "D-" & String(2 - Len(NewDNumber), "0") & NewDNumber
        Wend
        
        GetNewDepartmentID = Success
    
    Else
    
        GetNewDepartmentID = Failed
    End If
    
    
    
    Set vRS = Nothing

End Function


