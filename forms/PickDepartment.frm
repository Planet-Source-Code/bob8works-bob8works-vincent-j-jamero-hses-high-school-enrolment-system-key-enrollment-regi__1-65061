VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form PickDepartment 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Select Department"
   ClientHeight    =   4530
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5490
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "PickDepartment.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   302
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   366
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00D8E9EC&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4500
      Left            =   15
      ScaleHeight     =   300
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   364
      TabIndex        =   0
      Top             =   15
      Width           =   5460
      Begin VB.TextBox txtFind 
         Height          =   315
         Left            =   435
         TabIndex        =   1
         Top             =   405
         Width           =   3180
      End
      Begin lvButton.lvButtons_H cmdCancel 
         Cancel          =   -1  'True
         Height          =   375
         Left            =   3675
         TabIndex        =   2
         Top             =   360
         Width           =   870
         _ExtentX        =   1535
         _ExtentY        =   661
         Caption         =   "&Cancel"
         CapAlign        =   2
         BackStyle       =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFore           =   0
         cFHover         =   0
         cBhover         =   16185592
         cGradient       =   16185592
         Gradient        =   4
         Mode            =   0
         Value           =   0   'False
         cBack           =   14215660
      End
      Begin lvButton.lvButtons_H cmdSelect 
         Height          =   375
         Left            =   4545
         TabIndex        =   3
         Top             =   360
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   661
         Caption         =   "&Select"
         CapAlign        =   2
         BackStyle       =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFore           =   0
         cFHover         =   0
         cBhover         =   16185592
         cGradient       =   16185592
         Gradient        =   4
         Mode            =   0
         Value           =   0   'False
         cBack           =   14215660
      End
      Begin HSES.b8TitleBar b8TitleBar1 
         Height          =   345
         Left            =   -15
         TabIndex        =   4
         Top             =   -15
         Width           =   5475
         _ExtentX        =   9657
         _ExtentY        =   609
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontBold        =   -1  'True
         FontName        =   "Tahoma"
         FontSize        =   8.25
         Icon            =   "PickDepartment.frx":058A
      End
      Begin HSES.b8Container b8Container1 
         Height          =   3690
         Left            =   15
         TabIndex        =   5
         Top             =   795
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   6509
         BackColor       =   16185592
         Begin HSES.b8SContainer b8SContainer1 
            Height          =   375
            Left            =   30
            TabIndex        =   6
            Top             =   3285
            Width           =   5325
            _ExtentX        =   9393
            _ExtentY        =   661
            Begin HSES.b8Nav b8navRecord 
               Height          =   375
               Left            =   3540
               TabIndex        =   7
               Top             =   0
               Width           =   1800
               _ExtentX        =   3175
               _ExtentY        =   661
            End
            Begin VB.Label lblPage 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Page 0 of 0"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H0030A0B8&
               Height          =   195
               Left            =   2580
               TabIndex        =   9
               Top             =   90
               Width           =   930
            End
            Begin VB.Label lblListInfo 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "No Selected"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H0030A0B8&
               Height          =   195
               Left            =   60
               TabIndex        =   8
               Top             =   90
               Width           =   990
            End
         End
         Begin MSComctlLib.ListView listRecord 
            Height          =   3240
            Left            =   45
            TabIndex        =   10
            Top             =   45
            Width           =   5325
            _ExtentX        =   9393
            _ExtentY        =   5715
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            PictureAlignment=   4
            _Version        =   393217
            Icons           =   "ilRecordIco"
            SmallIcons      =   "ilRecordIco"
            ColHdrIcons     =   "ilRecordIco"
            ForeColor       =   8399906
            BackColor       =   16777215
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MouseIcon       =   "PickDepartment.frx":0B24
            NumItems        =   1
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Title"
               Object.Width           =   9366
            EndProperty
         End
      End
      Begin HSES.b8Line b8Line1 
         Height          =   60
         Left            =   0
         TabIndex        =   11
         Top             =   735
         Width           =   5505
         _ExtentX        =   9710
         _ExtentY        =   106
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Find"
         ForeColor       =   &H0030A0B8&
         Height          =   195
         Left            =   75
         TabIndex        =   12
         Top             =   435
         Width           =   300
      End
      Begin VB.Image Image4 
         Height          =   75
         Left            =   0
         Picture         =   "PickDepartment.frx":13FE
         Stretch         =   -1  'True
         Top             =   0
         Width           =   5565
      End
      Begin VB.Image Image1 
         Height          =   435
         Left            =   0
         Picture         =   "PickDepartment.frx":149B
         Stretch         =   -1  'True
         Top             =   345
         Width           =   5505
      End
   End
   Begin MSComctlLib.ImageList ilRecordIco 
      Left            =   2310
      Top             =   2235
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PickDepartment.frx":1538
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PickDepartment.frx":1AD2
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PickDepartment.frx":206C
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PickDepartment.frx":2606
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "PickDepartment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type
Dim R As RECT
Dim Alignable As Boolean


Dim tmpDepartment As String
Dim vRS As New ADODB.Recordset

Dim MaxEntryCount As Long
Dim CurRecPos As Long
Dim CurRecordCount As Long

Dim sOldDepartment As String

Dim sGetDepartmentTitle As String
Dim sDepartmentID As String


Public Function GetItem(Optional TextObject As Variant, Optional ByRef sDepartmentTitle As String, Optional lMaxEntryCount As Long = 15, Optional OldDepartment As String = "0000", Optional ExcludeClosed As Boolean = False) As String
    
    Dim sSQL As String
    Dim vDepartment As tDepartment
    
    'set fail to default
    GetItem = ""
    tmpDepartment = ""
    
    
    MaxEntryCount = lMaxEntryCount
    CurRecPos = 0
    
    sDepartmentID = ""
    sGetDepartmentTitle = ""
    
    If DepartmentRecordExist <> Success Then
        MsgBox "There are no record yet in Department Entries", vbExclamation
        Exit Function
    End If
    
    
    sSQL = "SELECT tblDepartment.DepartmentID as lvKey,tblDepartment.DepartmentTitle " & _
            " FROM tblDepartment" & _
            " ORDER BY tblDepartment.DepartmentTitle"
            
    'add yr to list
    If ConnectRS(HSESDB, vRS, sSQL) = True Then
        
        b8navRecordRefresh
        
        If CurRecordCount > 0 Then
            FillList CurRecPos, MaxEntryCount

        Else
            'error
            MsgBox "No Department  to be selected." & vbNewLine & "Please Add New Department  first.", vbExclamation
            Unload Me
            Exit Function
        End If
    Else
        'error
    End If

    'get pos
    If Not IsMissing(TextObject) Then
        GetWindowRect TextObject.hwnd, R
        Alignable = True
        Form_Activate
    Else
        Alignable = False
    End If
    
    'show form
    Me.Show vbModal
    
    'this next line will not execute unload this for will be unloaded
    sDepartmentTitle = sGetDepartmentTitle
    GetItem = tmpDepartment
End Function


Private Sub ReturnGetStudentID()
    If Len(GetLVKey(listRecord.SelectedItem)) > 0 Then
        sGetDepartmentTitle = listRecord.SelectedItem.Text
        tmpDepartment = GetLVKey(listRecord.SelectedItem)
        Unload Me
    End If
End Sub

Private Sub CancelGetStudentID()
    tmpDepartment = ""
    Unload Me
End Sub




Private Sub b8navRecordRefresh()
    CurRecordCount = getRecordCount(vRS)
    
    If CurRecPos > 0 Then
        b8navRecord.FirstEnable = True
        b8navRecord.PreviousEnable = True
    Else
        b8navRecord.FirstEnable = False
        b8navRecord.PreviousEnable = False
    End If
    
    If CurRecPos < CurRecordCount - MaxEntryCount Then
        b8navRecord.LastEnable = True
        b8navRecord.NextEnable = True
    Else
        b8navRecord.LastEnable = False
        b8navRecord.NextEnable = False
    End If
End Sub


Private Sub b8navRecord_Click(Index As Integer)
    Select Case Index
        Case 0
            CurRecPos = 0
            FillList CurRecPos, MaxEntryCount
            listRecord_Click
            
            
        Case 1
            If CurRecPos - MaxEntryCount >= 0 Then
        
                CurRecPos = CurRecPos - MaxEntryCount
                
                FillList CurRecPos, MaxEntryCount
                listRecord_Click
            End If
            
            
        Case 2
            If CurRecPos + MaxEntryCount < getRecordCount(vRS) Then
        
                CurRecPos = CurRecPos + MaxEntryCount
                        
                FillList CurRecPos, MaxEntryCount
                listRecord_Click
            End If
    
    
        Case 3
        
            Dim RC As Long
    
            RC = getRecordCount(vRS)
            
            If MaxEntryCount < RC Then
                'temp
                'pwede pa mapababa
                If (RC Mod MaxEntryCount) = 0 Then
                    CurRecPos = RC - MaxEntryCount
                Else
                    CurRecPos = RC - (RC Mod MaxEntryCount)
                End If
                    
                FillList CurRecPos, MaxEntryCount
                
                listRecord_Click
            End If
    End Select
    
    'refresh buttons
    b8navRecordRefresh
End Sub

Private Sub cmdCancel_Click()
    CancelGetStudentID
End Sub

Private Sub cmdFind_Click()
    Dim sSQL As String
    
    sSQL = "SELECT tblDepartment.DepartmentTitle AS lvKey, tblDepartment.DepartmentTitle" & _
            " From tblDepartment" & _
            " WHERE  ((DepartmentTitle) like '%" & txtFind.Text & "%')"


    If ConnectRS(HSESDB, vRS, sSQL) = True Then
        
        CurRecPos = 0
        b8navRecordRefresh
        
        
        If CurRecordCount > 0 Then
            
            FillList CurRecPos, MaxEntryCount

        Else
            'no result
            listRecord.ListItems.Clear
            listRecord_Click

        End If
    Else
        MsgBox "FATAL ERROR: PickStudent.cmdFind_Click - Connectrs"
    End If
    
End Sub








        

Private Sub cmdSelect_Click()
    ReturnGetStudentID
End Sub



Private Sub Form_Activate()
    Dim NewLeft As Long
    Dim NewTop As Long
    
    If Alignable = True Then
        If (R.Left * Screen.TwipsPerPixelX + Me.Width) > Screen.Width Then
            NewLeft = (R.Right * Screen.TwipsPerPixelX) - Me.Width
        Else
            NewLeft = R.Left * Screen.TwipsPerPixelX
        End If
        
        If (R.Bottom * Screen.TwipsPerPixelY + Me.Height) > Screen.Height Then
            NewTop = (R.Top * Screen.TwipsPerPixelY) - Me.Height
            If NewTop < 0 Then NewTop = 0
        Else
            NewTop = R.Bottom * Screen.TwipsPerPixelY
        End If
        
        Me.Left = NewLeft
        Me.Top = NewTop
        
    Else
    
        CenterForm Me
        
    End If
End Sub



Private Sub listRecord_Click()
    Dim totalPage As Long
    Dim curPage As Long
    
    If listRecord.ListItems.Count < 1 Then
        lblListInfo.Caption = "No Record"
        lblPage.Caption = "Page 0 of 0"
    Else
        lblListInfo.Caption = "Selected Entry: " & listRecord.SelectedItem.Index + CurRecPos & "/" & CurRecordCount
        
        totalPage = CurRecordCount \ MaxEntryCount + 1
        
        lblPage.Caption = "Page " & ((CurRecPos \ MaxEntryCount) + 1) & " of " & totalPage
    End If
End Sub



Private Sub listRecord_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    SortLV listRecord, ColumnHeader.Index - 1
End Sub

Private Sub listRecord_DblClick()
    ReturnGetStudentID
End Sub

Private Function FillList(lStart As Long, dCount As Long) As Boolean

    FillRecordToList vRS, listRecord, KeyStudent, lStart, dCount, , True
    listRecord_Click
End Function



Private Sub listRecord_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim curPos As Long
    If KeyCode = vbKeyDown Then
        If listRecord.SelectedItem.Index = listRecord.ListItems.Count Then
            
            b8navRecord_Click 2
            
            KeyCode = 0
        End If
    End If
    
    If KeyCode = vbKeyUp Then
        If listRecord.SelectedItem.Index = 1 Then
            curPos = CurRecPos
            
            b8navRecord_Click 1
            
            If curPos <> CurRecPos Then
                listRecord.SelectedItem.Selected = False
                listRecord.ListItems(listRecord.ListItems.Count).Selected = True
            End If
            
            KeyCode = 0
        End If
    End If
    
    If KeyCode = vbKeyPageDown Then
        b8navRecord_Click 2
    End If
    
    If KeyCode = vbKeyPageUp Then
        b8navRecord_Click 1
    End If
End Sub

Private Sub listRecord_KeyUp(KeyCode As Integer, Shift As Integer)
   
    If KeyCode = 38 Or KeyCode = 40 Or KeyCode = 33 Or KeyCode = 34 Then listRecord_Click
    
    If KeyCode = vbKeyReturn Then ReturnGetStudentID
    
    
End Sub

Private Sub txtFind_Change()
    'delay 0.3 second
    'code by: VIncent J. Jamero
    '------------------------------------------------
    Static DelayStart As Single
    Static notFirst As Boolean
    DelayStart = GetTickCount + 300
    If notFirst = True Then Exit Sub
    notFirst = True
    While GetTickCount < DelayStart
        DoEvents
    Wend
    notFirst = False
    '------------------------------------------------
    'the next line will be if executed if user pause typing in 0.3 second
    
    'If Len(Trim(txtFind.Text)) > 0 Then
        cmdFind_Click
    'End If
End Sub

Private Sub txtFind_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then
        listRecord.SetFocus
    End If
End Sub


'end p---


