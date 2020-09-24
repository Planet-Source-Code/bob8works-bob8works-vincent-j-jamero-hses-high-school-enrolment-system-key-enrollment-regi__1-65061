Attribute VB_Name = "modAppSettings"
Option Explicit


Public Function AppSet_GetLockTimeOut(Optional sUserName As String = "") As Integer
    Dim sValue As String
    
    If sUserName = "" Then
        sUserName = CurrentUser.UserName
    End If
    
    sValue = GetSetting(App.Title, "settings", sUserName & "-LockTimeOut", 60)
    
    If IsNumeric(sValue) Then
        AppSet_GetLockTimeOut = Val(sValue)
    Else
        AppSet_GetLockTimeOut = 60
    End If
    
End Function

Public Function AppSet_SetLockTimeOut(Optional iNewLockTime As Integer = 60, Optional sUserName As String = "")
    
    If sUserName = "" Then
        sUserName = CurrentUser.UserName
    End If
    
    If iNewLockTime < 60 Then iNewLockTime = 60
    
    SaveSetting App.Title, "settings", sUserName & "-LockTimeOut", Str(iNewLockTime)
    

End Function

Public Function AppGet_LoginUserName() As String
    
    AppGet_LoginUserName = GetSetting(App.Title, "settings", "LoginUserName", "")

End Function

Public Function AppSet_LoginUserName(sUserName As String)
    SaveSetting App.Title, "settings", "LoginUserName", sUserName
End Function
