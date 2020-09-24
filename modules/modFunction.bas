Attribute VB_Name = "modFunction"
Option Explicit


'functions
Public Enum FindOptions
    PartOfWord = 0
    MatchCase = 1
    WholeWordOnly = 3
End Enum


'API for opening a browser
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
                (ByVal hwnd As Long, _
                ByVal lpOperation As String, _
                ByVal lpFile As String, _
                ByVal lpParameters As String, _
                ByVal lpDirectory As String, _
                ByVal nShowCmd As Long) As Long

Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Public Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type


Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function GetTickCount Lib "kernel32" () As Long


Public Type POINTAPI
        X As Long
        Y As Long
End Type

Private Declare Sub ReleaseCapture Lib "user32" ()
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long


Public Sub FormDrag(frmName As Form) 'procedure to drag a no-titlebar form
    ReleaseCapture
    Call SendMessage(frmName.hwnd, &HA1, 2, 0&)
End Sub



Public Function MakeGradient(ByRef frm As Object, Scheme As Integer)
    Dim cR(255) As Integer
    Dim cG(255) As Integer
    Dim cB(255) As Integer
    Dim d As Double
    Dim i As Integer
    
    
    Select Case Scheme
        Case 1
            For i = 0 To 255
                cR(i) = 255 - (i * 0.2)
                cG(i) = 255 - (i * 0.2)
                cB(i) = 255 - (i * 0.2)
            Next
    End Select
    

    frm.ScaleMode = vbPixels
    d = frm.ScaleHeight / 255
    frm.DrawWidth = d + 1
    For i = 0 To 255
        frm.ForeColor = RGB(cR(i), cG(i), cB(i))
        frm.Line (0, i * d)-(frm.ScaleWidth, i * d)
    Next
    'Frm.AutoRedraw = True
End Function











Public Function CheckTextBox(ByRef txt As Object, Optional sMSG As String = "TextBox", Optional ShowMSG As Boolean = True, Optional MinimumChar As Integer = 1) As Boolean
On Error Resume Next
    If Len(Trim(txt.Text)) < MinimumChar Then
        
        If ShowMSG Then
            MsgBox sMSG, vbExclamation
        End If
        
        txt.Text = ""
        txt.SetFocus
        
        CheckTextBox = False
    Else
        CheckTextBox = True
    End If
End Function

Public Function HLTxt(ByRef txt As Object)
On Error Resume Next
    txt.selStart = 0
    txt.selLength = Len(txt)
    txt.SetFocus
End Function


Public Function AddListItem(ByRef vListItem As ListView, sText As Variant, Optional imgIndex As Integer = 1, Optional sSubItem1 As Variant = "", Optional sSubItem2 As Variant = "", Optional sSubItem3 As Variant = "")
Dim lastIndex As Integer
    
On Error Resume Next
    
    lastIndex = vListItem.ListItems.Count + 1
    vListItem.ListItems.Add lastIndex, , sText, imgIndex, imgIndex
    If sSubItem1 <> "" Then _
        vListItem.ListItems(lastIndex).SubItems(1) = sSubItem1
        If sSubItem2 <> "" Then _
        vListItem.ListItems(lastIndex).SubItems(2) = sSubItem2
        If sSubItem3 <> "" Then _
        vListItem.ListItems(lastIndex).SubItems(3) = sSubItem3
End Function






Public Function cSentenceCase(sText As String) As String
    
    Dim splitText() As String
    Dim newWord As String
    Dim i As Integer
    
    'check if null---------------
    If Len(sText) < 1 Then
        cSentenceCase = ""
        Exit Function
    End If
    'end Null --------------------
    
    'convert
    sText = Trim(sText)
    
    splitText = Split(sText, " ")
    
    For i = 0 To UBound(splitText)
        If Len(Trim(splitText(i))) > 0 Then
            newWord = UCase(Left(Trim(splitText(i)), 1)) & LCase(Right(Trim(splitText(i)), Len(Trim(splitText(i))) - 1))
            cSentenceCase = cSentenceCase & " " & newWord
        End If
    Next
    
    cSentenceCase = Trim(cSentenceCase)
End Function



Public Function FillRecordToList(ByRef vRS As ADODB.Recordset, ByRef lv As ListView, sTableKey As String, Optional RecStartPos As Long = 0, Optional LimitCount As Long = 100, Optional WithID As Boolean = True, Optional WithIcon As Boolean = False)

    Dim i As Long
    Dim newColumnWidth As Integer
    Dim LimitCounter As Long
    Dim sCell As String
    Dim oldScaleMode As ScaleModeConstants
    
On Error Resume Next
    
    'minum fields must be 2
    If vRS.Fields.Count < 2 Then Exit Function
    
    'get old scale mode
    oldScaleMode = lv.Container.ScaleMode

    lv.Container.ScaleMode = vbTwips
    lv.ListItems.Clear
    
    
    If AnyRecordExisted(vRS) Then
        
        'create column headers
        For i = lv.ColumnHeaders.Count To vRS.Fields.Count - 2
            lv.ColumnHeaders.Add
        Next
        
        'set items
        vRS.Requery
        vRS.Move RecStartPos
        
        LimitCounter = 0
        
        
        While Not vRS.EOF
        
            
        
            'add
            If WithIcon = True Then
                If WithID = True Then
                    lv.ListItems.Add , SetLVKey(vRS.Fields(0).Value, sTableKey), vRS.Fields(1).Value, 1, 1
                Else
                    lv.ListItems.Add , , vRS.Fields(0).Value, 1, 1
                End If
            Else
                If WithID = True Then
                    lv.ListItems.Add , SetLVKey(vRS.Fields(0).Value, sTableKey), vRS.Fields(1).Value
                Else
                    lv.ListItems.Add , , vRS.Fields(0).Value
                End If
            End If
            'add sub items
            For i = 2 To vRS.Fields.Count - 1
           
                lv.ListItems(lv.ListItems.Count).SubItems(i - 1) = ReadField(vRS.Fields(i))
            Next
            
            vRS.MoveNext
            
            LimitCounter = LimitCounter + 1
            
            If LimitCounter >= LimitCount Then
                GoTo tagExitSub
            End If
            
        Wend
       
    End If
    
tagExitSub:
    'restore scale mode
    lv.Container.ScaleMode = oldScaleMode

End Function


Public Function FillRecordToListWN(ByRef vRS As ADODB.Recordset, ByRef lv As ListView, sTableKey As String, Optional RecStartPos As Long = 0, Optional LimitCount As Long = 100, Optional WithID As Boolean = True, Optional WithIcon As Boolean = False, Optional ColIndex As Integer = 0)

    Dim i As Long
    Dim newColumnWidth As Integer
    Dim LimitCounter As Long
    Dim sCell As String
    Dim oldScaleMode As ScaleModeConstants
    
    Dim CPOS As Long
    
On Error Resume Next
    
    'minum fields must be 2
    If vRS.Fields.Count < 2 Then Exit Function
    
    'get old scale mode
    oldScaleMode = lv.Container.ScaleMode

    lv.Container.ScaleMode = vbTwips
    lv.ListItems.Clear
    
    
    If AnyRecordExisted(vRS) Then
        
        'create column headers
        For i = lv.ColumnHeaders.Count To vRS.Fields.Count - 1
            lv.ColumnHeaders.Add
        Next
        
        'set items
        vRS.Requery
        vRS.Move RecStartPos
        
        LimitCounter = 0
        CPOS = RecStartPos
        
        While Not vRS.EOF
        
            CPOS = CPOS + 1
            
        
            'add
            If ColIndex = 0 Then
                
                If WithIcon = True Then
                    If WithID = True Then
                        lv.ListItems.Add , SetLVKey(vRS.Fields(0).Value, sTableKey), CPOS
                    Else
                        lv.ListItems.Add , , CPOS, 1, 1
                    End If
                Else
                    If WithID = True Then
                        lv.ListItems.Add , SetLVKey(vRS.Fields(0).Value, sTableKey), CPOS
                    Else
                        lv.ListItems.Add , , CPOS
                    End If
                End If
                
            
            Else
                
                If WithIcon = True Then
                    If WithID = True Then
                        lv.ListItems.Add , SetLVKey(vRS.Fields(0).Value, sTableKey), vRS.Fields(1).Value, 1, 1
                    Else
                        lv.ListItems.Add , , vRS.Fields(0).Value, 1, 1
                    End If
                Else
                    If WithID = True Then
                        lv.ListItems.Add , SetLVKey(vRS.Fields(0).Value, sTableKey), vRS.Fields(1).Value
                    Else
                        lv.ListItems.Add , , vRS.Fields(0).Value
                    End If
                End If
            
            End If
            
            
            
            
            'add sub items
            
            If ColIndex = 0 Then
            
                 For i = 1 To vRS.Fields.Count '- 1
                
                     lv.ListItems(lv.ListItems.Count).SubItems(i - 1) = ReadField(vRS.Fields(i - 1))
                 Next
                 
                 vRS.MoveNext
                 
                 LimitCounter = LimitCounter + 1
                 
                 If LimitCounter >= LimitCount Then
                     GoTo tagExitSub
                 End If
            

            Else
                Dim sbi As Integer
                sbi = 1
                For i = 2 To vRS.Fields.Count '- 1
                
                    If ColIndex = i Then
                        lv.ListItems(lv.ListItems.Count).SubItems(sbi) = CPOS
                        sbi = sbi + 1
                        lv.ListItems(lv.ListItems.Count).SubItems(sbi) = ReadField(vRS.Fields(i))
                    Else
                        lv.ListItems(lv.ListItems.Count).SubItems(sbi) = ReadField(vRS.Fields(i))
                    End If
                    
                    sbi = sbi + 1
                Next
                 
                vRS.MoveNext
                 
                LimitCounter = LimitCounter + 1
                 
                If LimitCounter >= LimitCount Then
                    GoTo tagExitSub
                End If
            
            End If
             
             
             
            
        Wend 'individual record
       
    End If
    
tagExitSub:
    'restore scale mode
    lv.Container.ScaleMode = oldScaleMode

End Function

Public Function SortLV(ByRef lv As ListView, Optional HeaderIndex As Integer = 0, Optional newSortOrder As ListSortOrderConstants = lvwAscending, Optional AutoOrder As Boolean = True)
    
    Dim lvHeader As ColumnHeader
    
    If AutoOrder = True Then
        If lv.SortOrder = lvwAscending Then
           lv.SortOrder = lvwDescending
        Else
           lv.SortOrder = lvwAscending
        End If
    Else
        lv.SortOrder = newSortOrder
    End If
    
    If HeaderIndex > lv.ColumnHeaders.Count - 1 Then
        HeaderIndex = 0
    End If
    
    lv.SortKey = HeaderIndex
    lv.Sorted = True
    lv.Refresh
    
    For Each lvHeader In lv.ColumnHeaders
        lvHeader.Icon = 0
    Next
    
    On Error Resume Next
    lv.ColumnHeaders(HeaderIndex + 1).Icon = lv.SortOrder + 1
End Function

Public Function UnSortLV(ByRef lv As ListView)
    
    Dim lvHeader As ColumnHeader
    
    lv.Sorted = False
    
    For Each lvHeader In lv.ColumnHeaders
        lvHeader.Icon = 0
    Next
End Function
    
    
    
    
    

Public Function GetLVKey(lvListItem As ListItem) As String
On Error GoTo errh:
    GetLVKey = Right(lvListItem.Key, Len(lvListItem.Key) - 4)
    Exit Function
errh:
    GetLVKey = ""
End Function
Public Function SetLVKey(sID As String, sTableKey As String) As String
    SetLVKey = Left(sTableKey, 4) & sID
End Function

Public Function FindLVItem(ByRef vLV As ListView, sCriteria As String, Optional iOption As FindOptions = 0, Optional MultiSelect As Boolean = False, Optional InverseSelection As Boolean = False, Optional FindNext As Boolean = False)

    Dim i As Integer
    Dim isFound As Boolean
    Dim li As Integer
    Dim StartPos As Integer
    
'On Error GoTo eh
    
    If vLV.ListItems.Count < 1 Then Exit Function

    If FindNext = True And vLV.SelectedItem.Index < vLV.ListItems.Count Then
        For li = 1 To vLV.SelectedItem.Index
            vLV.ListItems(li).Selected = False
        Next
        StartPos = vLV.SelectedItem.Index + 1
    Else
        For li = 1 To vLV.ListItems.Count
            vLV.ListItems(li).Selected = False
        Next
        StartPos = 1
    End If
    
    'set flag to default
    isFound = False
    
    For li = StartPos To vLV.ListItems.Count
        
        Select Case iOption
            
            Case FindOptions.PartOfWord  'normal

                If InStr(1, LCase(vLV.ListItems(li).Text), LCase(sCriteria)) > 0 Then
                                        
                    isFound = True

                Else

                    'check subitems
                    For i = 1 To vLV.ListItems(li).ListSubItems.Count
                        If InStr(1, LCase(vLV.ListItems(li).ListSubItems(i)), LCase(sCriteria)) > 0 Then
                            
                            isFound = True
                            Exit For
                        
                        End If
                    Next
                                        
                End If
                
            Case FindOptions.MatchCase  'match case
            
            Case FindOptions.WholeWordOnly  ' whole word only
                
            
        End Select
        
        
        
        
        If isFound Then
            
            vLV.ListItems(li).Selected = CBool(True - InverseSelection)
            vLV.ListItems(li).EnsureVisible
            
            If Not MultiSelect Then Exit For
        
        Else
            vLV.ListItems(li).Selected = CBool(False - InverseSelection)
        End If
        
    Next
    
    If FindNext = True And isFound = False And StartPos > 1 Then
        
        For li = 1 To StartPos
            
            Select Case iOption
                
                Case FindOptions.PartOfWord  'normal
    
                    If InStr(1, LCase(vLV.ListItems(li).Text), LCase(sCriteria)) > 0 Then
                                            
                        isFound = True
    
                    Else
    
                        'check subitems
                        For i = 1 To vLV.ListItems(li).ListSubItems.Count
                            If InStr(1, LCase(vLV.ListItems(li).ListSubItems(i)), LCase(sCriteria)) > 0 Then
                                
                                isFound = True
                                Exit For
                            
                            End If
                        Next
                                            
                    End If
                    
                Case FindOptions.MatchCase  'match case
                
                Case FindOptions.WholeWordOnly  ' whole word only
                    
                
            End Select
            
            
            
            
            If isFound Then
                
                vLV.ListItems(li).Selected = CBool(True - InverseSelection)
                vLV.ListItems(li).EnsureVisible
                
                If Not MultiSelect Then Exit For
            
            Else
                vLV.ListItems(li).Selected = CBool(False - InverseSelection)
            End If
            
        Next
    End If
'On Error Resume Next
Exit Function
eh:
    MsgBox Err.Description
    Resume Next
End Function


Public Function GetLVSelectedCount(ByRef lv As ListView) As Integer
    Dim i As Integer
    Dim iSelectedCount As Integer
    
    'default
    GetLVSelectedCount = 0
    
    'check if there is a record in the list
    If lv.ListItems.Count < 1 Then Exit Function
    
    
    iSelectedCount = 0
    For i = 1 To lv.ListItems.Count
        If lv.ListItems(i).Selected = True And Len(GetLVKey(lv.ListItems(i))) > 0 Then
            iSelectedCount = iSelectedCount + 1
        End If
    Next
    
    'return
    GetLVSelectedCount = iSelectedCount
End Function

Public Function CatchError(sModuleName As String, sRoutineName As String, sDetail As String)
    MsgBox sModuleName & " - " & sRoutineName & " - " & sDetail
    
End Function


Public Function CenterForm(ByRef frm As Form)
    frm.Move (Screen.Width - frm.Width) / 2, (Screen.Height - frm.Height) / 2
End Function


'credit: philip naparan
Public Sub OpenURL(urlADD As String, sourceHWND As Long)
     Call ShellExecute(sourceHWND, vbNullString, urlADD, "", vbNullString, 1)
End Sub
