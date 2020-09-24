Attribute VB_Name = "modMain"
Public stateCheckFiles As Boolean
Public stateHSESDataExisted As Boolean



'MAIN
'-------------------------------------------------
Sub Main()

    frmMSG.ShowForm

    'show splash
    frmSplash.ShowSplash
    
    'initialize variables
    Call InitVar
    
    
    
    'change system color
    'Call select_color_type
    
    'check files
    If Not CheckFiles Then
        InitAppFailed "Checking Files"
        Exit Sub
    End If
    
    
    'connect databse
    If SetDatabse < 0 Then
        InitAppFailed "Set Databse"
        Exit Sub
    End If
    
    'get settings
    GetDataSettings
    
    frmSplash.lblSchoolName.Caption = CurrentSchool.SchoolName
    DoEvents
    
    'check user
    If CheckUser < 0 Then
        Exit Sub
    End If
    
    
    
        
    'unload splash
    frmSplash.UnloadSplash
    
    'show login
    
    frmLogin.ShowLogin
    
    
    
    
End Sub 'MAIN ------------------------------------


Private Function GetDataSettings()
    CurrentSchoolYear.SchoolYearTitle = GetActiveSchoolYear
End Function

Public Sub InitAppFailed(sMSG As String)
    'write some logo here
    
    MsgBox "Writng to Log: " & sMSG, vbCritical
    
    'exit
    End
End Sub




Private Function CheckUser() As Variant
    
    
    
    'check user count
    If UserRecordExist = False Then
        
        'found no user
                
        'unload splash
        frmSplash.UnloadSplash
        'show welcome
        frmWelcome.Show
        CheckUser = -1
    Else
    
        'set flag
        CheckUser = 1
    End If
End Function






Public Sub AfterLogin()
    
    'check school
    Call CheckSchoolInfo
    
    'get application settings
    Call GetAppSettings
    
End Sub

Public Function CheckSchoolInfo()

    

    If Not isSchoolExisted Then
        frmSchoolAccount.AddNew
        Exit Function
    End If
    
    SchoolExisted = True
    Call AfterCheckSchoolInfo
End Function

Public Sub AfterCheckSchoolInfo()
    'get school info
    Call GetSchoolInfo(CurrentSchool)
    'show main form
    mdiMain.Show
End Sub












Private Function SetDatabse() As Variant
    
    If Not ConnectDB(HSESDB, DBPathFileName) Then
        SetDatabse = -1
        MsgBox "Error: Connecting Databse"
        Exit Function
    End If
    
    

    
    
    'connect school
    If Not ConnectRS(HSESDB, RSSchool, "select * from tblschool") Then
        
        SetDatabse = -3
        MsgBox "Error: Connecting Recordset - School"
        Exit Function

    End If
    
    
    

    

    
    

    
    
    


    
    'return success
    SetDatabse = 1
    
End Function






Private Function InitVar() As Variant
    
    'set HSESDB path file name
    DBPathFileName = App.Path & "\HSESData.mdb"
    
    'set original school year database file
    SYOriginalFilePath = App.Path & "\HSESE.mdb"
    
    'set list key
    kSelectListKey = Asc("`")
    
    
    'set minimum grades
     MinGradeForI = 75
     MinGradeForII = 75
     MinGradeForIII = 75
     MinGradeForIV = 75
     MinGradeForGraduate = 75
     
     
End Function

Public Function GetAppSettings()
    'get settings
     AppSet_LockTimeOut = AppSet_GetLockTimeOut
End Function

Private Function CheckFiles() As Boolean
    
    Dim fso As New FileSystemObject
    Dim cdGetDB As CommonDialog

    
    
    'set default to failed
    CheckFiles = False
    
    'check database file
    While fso.FileExists(DBPathFileName) = False
        
        'database does not existed
        'prompt user
        
        Select Case MsgBox("Database file does not exist." & vbNewLine & "Dow you wand to locate this file?", vbQuestion + vbYesNoCancel)
            
            Case vbYes 'user will try to locate file
                    
                    'show dialog open
                    Set cdGetDB = frmTemp.ComDlg
        
                    cdGetDB.DialogTitle = "Locate HSES Database File"
                    cdGetDB.FileName = "HSESData.mdb"
                    cdGetDB.Filter = "HSESData.mdb|HSESData.mdb|All Files|*.*"
                    cdGetDB.InitDir = App.Path
                    cdGetDB.ShowOpen
                    
                    DBPathFileName = cdGetDB.FileName
                    'unload temporary form
                    Unload frmTemp
            
            Case vbNo
                MsgBox "Database file not found. The application will now exit.", vbExclamation
                Exit Function
            
            Case vbCancel 'aborted
                MsgBox "Database file not found. The application will now exit.", vbExclamation
                Exit Function
        
        End Select
        
        
        
        
    Wend
    
    'success
    'return true
    CheckFiles = True
    
End Function



Public Function FoundNoStudent() As Variant
    MsgBox "No Student Record Yet", vbInformation
End Function










'record transactions







