VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmASSectionOffering 
   BackColor       =   &H00D8E9EC&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Search"
   ClientHeight    =   6885
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   7905
   Icon            =   "frmASSectionOffering.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   459
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   527
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin HSES.b8SContainer b8SContainer1 
      Height          =   3135
      Left            =   60
      TabIndex        =   12
      Top             =   3180
      Width           =   7755
      _ExtentX        =   13679
      _ExtentY        =   5530
      BorderColor     =   12835550
      Begin VB.CommandButton cmdRemoveAll 
         BackColor       =   &H00D8E9EC&
         Height          =   330
         Left            =   7140
         Picture         =   "frmASSectionOffering.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   60
         Width           =   525
      End
      Begin VB.CommandButton cmdRemoveOne 
         BackColor       =   &H00D8E9EC&
         Height          =   330
         Left            =   6600
         Picture         =   "frmASSectionOffering.frx":0E54
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   60
         Width           =   525
      End
      Begin MSComctlLib.ListView lvSearch 
         Height          =   2625
         Left            =   60
         TabIndex        =   13
         Top             =   420
         Width           =   7605
         _ExtentX        =   13414
         _ExtentY        =   4630
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Description"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "SQL Expression"
            Object.Width           =   7937
         EndProperty
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Criteria List"
         Height          =   195
         Left            =   90
         TabIndex        =   18
         Top             =   90
         Width           =   765
      End
   End
   Begin VB.ComboBox cmbCriteria 
      Height          =   315
      Left            =   990
      TabIndex        =   9
      Top             =   2010
      Width           =   3195
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00D8E9EC&
      Caption         =   "Next Expression"
      Height          =   675
      Left            =   4530
      TabIndex        =   6
      Top             =   1260
      Width           =   1965
      Begin VB.OptionButton oNE 
         BackColor       =   &H00D8E9EC&
         Caption         =   "OR"
         Height          =   285
         Index           =   1
         Left            =   1140
         TabIndex        =   8
         Top             =   300
         Width           =   735
      End
      Begin VB.OptionButton oNE 
         BackColor       =   &H00D8E9EC&
         Caption         =   "AND"
         Height          =   285
         Index           =   0
         Left            =   180
         TabIndex        =   7
         Top             =   300
         Value           =   -1  'True
         Width           =   795
      End
   End
   Begin VB.ComboBox cmbFields 
      Height          =   315
      Left            =   990
      TabIndex        =   4
      Top             =   1290
      Width           =   3195
   End
   Begin VB.ComboBox cmbOperand 
      Height          =   315
      Left            =   990
      TabIndex        =   3
      Top             =   1650
      Width           =   1365
   End
   Begin VB.CommandButton cmdAddOne 
      BackColor       =   &H00D8E9EC&
      Height          =   420
      Left            =   6780
      Picture         =   "frmASSectionOffering.frx":13DE
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2640
      Width           =   1005
   End
   Begin HSES.b8ChildTitleBar b8ChildTitleBar2 
      Height          =   285
      Left            =   0
      TabIndex        =   1
      Top             =   630
      Width           =   9675
      _ExtentX        =   17066
      _ExtentY        =   503
      BackColor       =   12835550
      Caption         =   "   Find Entries Where"
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
      ForeColor       =   3102058
      CloseButton     =   0   'False
   End
   Begin lvButton.lvButtons_H cmdOK 
      Height          =   360
      Left            =   6390
      TabIndex        =   5
      Top             =   6450
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   635
      Caption         =   "&OK"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin HSES.b8Line b8Line1 
      Height          =   60
      Left            =   0
      TabIndex        =   16
      Top             =   510
      Width           =   8925
      _ExtentX        =   15743
      _ExtentY        =   106
   End
   Begin HSES.b8Line b8Line2 
      Height          =   60
      Left            =   0
      TabIndex        =   17
      Top             =   6330
      Width           =   8925
      _ExtentX        =   15743
      _ExtentY        =   106
   End
   Begin HSES.b8Line b8Line3 
      Height          =   60
      Left            =   150
      TabIndex        =   19
      Top             =   2520
      Width           =   7635
      _ExtentX        =   13467
      _ExtentY        =   106
      BorderColor1    =   8421504
      BorderColor2    =   14215660
      BorderColor3    =   14215660
      BorderStyle1    =   3
   End
   Begin lvButton.lvButtons_H cmdCancel 
      Cancel          =   -1  'True
      Height          =   360
      Left            =   4740
      TabIndex        =   21
      Top             =   6450
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   635
      Caption         =   "&Cancel"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Filter Section Offering Entries"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   345
      Left            =   120
      TabIndex        =   22
      Top             =   120
      Width           =   4260
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Click To Add -->"
      ForeColor       =   &H00008080&
      Height          =   285
      Left            =   5430
      TabIndex        =   20
      Top             =   2730
      Width           =   1215
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Criteria"
      Height          =   195
      Left            =   180
      TabIndex        =   11
      Top             =   2100
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Operand"
      Height          =   195
      Left            =   180
      TabIndex        =   10
      Top             =   1740
      Width           =   615
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Field"
      Height          =   195
      Left            =   150
      TabIndex        =   0
      Top             =   1320
      Width           =   330
   End
   Begin VB.Image Image2 
      Height          =   525
      Left            =   0
      Picture         =   "frmASSectionOffering.frx":1968
      Stretch         =   -1  'True
      Top             =   0
      Width           =   8925
   End
End
Attribute VB_Name = "frmASSectionOffering"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sFL() As String
Dim sCS() As String
Dim sFT() As String
Dim sD() As String
Dim sSearchStr() As String
Dim iSSUbound As Integer

Dim cmbFields_i As Integer

Public Function ShowForm()
    
    Me.Show vbModal
End Function


Private Sub cmbCriteria_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        cmdAddOne.SetFocus
    End If
End Sub

Private Sub cmbFields_Change()
    RefreshCriteriaByField
End Sub

Private Sub cmbFields_Click()
    RefreshCriteriaByField
End Sub

Private Sub cmbFields_GotFocus()
    cmbFields_i = cmbFields.ListIndex
End Sub

Private Sub cmbFields_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        cmbOperand.SetFocus
    End If
End Sub

Private Sub cmbFields_LostFocus()
    
    If cmbFields.ListIndex < 0 Or Left(cmbFields.Text, 1) = "-" Then
        cmbFields.ListIndex = cmbFields_i
    End If
    
    RefreshCriteriaByField
End Sub

Private Function RefreshCriteriaByField()

    cmbCriteria.Clear

    Select Case LCase(cmbFields.Text)
    
        Case "school year"
            AddList_SchoolYear
            
        Case "department"
            AddList_Department
            
        Case "year level"
            cmbCriteria.Clear
            cmbCriteria.AddItem "I"
            cmbCriteria.AddItem "II"
            cmbCriteria.AddItem "III"
            cmbCriteria.AddItem "IV"
            cmbCriteria.ListIndex = 0
            
        Case "section"
            AddList_Section
            
            
    End Select
End Function


Private Function AddList_SchoolYear()
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    On Error GoTo ReleaseAndExit
    
    sSQL = "SELECT tblSchoolYear.SchoolYearTitle" & _
            " FROM tblSchoolYear" & _
            " ORDER BY tblSchoolYear.SchoolYearTitle"
     
    If ConnectRS(HSESDB, vRS, sSQL) = False Then
        'fatal
        'temp
        MsgBox "ERROR"
        GoTo ReleaseAndExit
    End If
    
    cmbCriteria.Clear
    
    While vRS.EOF = False
        
        cmbCriteria.AddItem ReadField(vRS.Fields("SchoolYearTitle"))
        vRS.MoveNext
    
    Wend
    
    cmbCriteria.ListIndex = 0
    
ReleaseAndExit:
    Set vRS = Nothing
End Function

Private Function AddList_Department()
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    On Error GoTo ReleaseAndExit
    
    sSQL = "SELECT tblDepartment.DepartmentTitle" & _
            " FROM tblDepartment" & _
            " ORDER BY tblDepartment.DepartmentTitle"
     
    If ConnectRS(HSESDB, vRS, sSQL) = False Then
        'fatal
        'temp
        MsgBox "ERROR"
        GoTo ReleaseAndExit
    End If
    
    cmbCriteria.Clear
    
    While vRS.EOF = False
        
        cmbCriteria.AddItem ReadField(vRS.Fields("DepartmentTitle"))
        vRS.MoveNext
    
    Wend
    
    cmbCriteria.ListIndex = 0
    
ReleaseAndExit:
    Set vRS = Nothing
End Function

Private Function AddList_Section()
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    On Error GoTo ReleaseAndExit
    
    sSQL = "SELECT tblSection.SectionTitle" & _
            " FROM tblSection" & _
            " ORDER BY tblSection.SectionTitle"
     
    If ConnectRS(HSESDB, vRS, sSQL) = False Then
        'fatal
        'temp
        MsgBox "ERROR"
        GoTo ReleaseAndExit
    End If
    
    cmbCriteria.Clear
    
    While vRS.EOF = False
        
        cmbCriteria.AddItem ReadField(vRS.Fields("SectionTitle"))
        vRS.MoveNext
    
    Wend
    
    cmbCriteria.ListIndex = 0
    
ReleaseAndExit:
    Set vRS = Nothing
End Function

Private Sub cmbOperand_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        cmbCriteria.SetFocus
    End If
End Sub

Private Sub cmdAddOne_Click()
    
    Dim sOperand As String
    Dim sONE As String
    Dim sT As String
    
    If Len(cmbCriteria.Text) < 1 Then
        MsgBox "Please Enter Criteria.", vbExclamation
        HLTxt cmbCriteria
        Exit Sub
    End If
    
    
    'field type
    Select Case sFT(cmbFields.ListIndex)
        Case "s"
            sT = "'"
            
        Case "n"
            If IsNumeric(cmbCriteria.Text) = False Then
                MsgBox "Please enter valid Numeric value.", vbExclamation
                HLTxt cmbCriteria
                Exit Sub
            End If
            
            sT = ""
            cmbCriteria.Text = Val(cmbCriteria.Text)
        
        Case "d"
            If IsDate(cmbCriteria.Text) = False Then
                MsgBox "Please enter valid Date value.", vbExclamation
                HLTxt cmbCriteria
                Exit Sub
            End If
            sT = "#"
            cmbCriteria.Text = CDate(cmbCriteria.Text)
        
        
        Case Else
            sT = "'"
            
    End Select
    
    
    'operand for criteria
    Select Case cmbOperand.Text
    
        Case "Equal"
            sOperand = " = "
            
        Case "Not Equal"
            sOperand = " <> "
            
        Case "Contain"
            sOperand = " like "
            
        Case Else
            sOperand = " = "
        
    End Select
    
    'Operand for next expression
    If iSSUbound < 1 Then
        sONE = ""
    Else
        If oNE(0).Value = True Then
            sONE = " AND "
        Else
            sONE = " OR "
        End If
    End If
    
    
    
    
    
    sSearchStr(iSSUbound) = sONE & " ((" & sFL(cmbFields.ListIndex) & _
                            ")" & sOperand & sT & cmbCriteria.Text & sT & ")"
    sD(iSSUbound) = "[" & cmbFields.Text & "] " & cmbOperand.Text & " [" & cmbCriteria.Text & "]"
    
    RefreshSearchList
    
    iSSUbound = iSSUbound + 1
    
   
    
End Sub

Private Function RefreshSearchList()
    Dim i As Integer
    
    lvSearch.ListItems.Clear
    
    For i = 0 To iSSUbound
        lvSearch.ListItems.Add , , sD(i)
        lvSearch.ListItems(lvSearch.ListItems.Count).SubItems(1) = sSearchStr(i)
    Next
    
End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    
    Dim sSQL As String
    
    Dim i As Integer
    
    For i = 1 To lvSearch.ListItems.Count
        sSQL = sSQL & lvSearch.ListItems(i).SubItems(1)
    Next
    
    Unload Me
    
    frmAllSectionOffering.ShowFormByCriteria sSQL

End Sub

Private Sub cmdRemoveAll_Click()
    RefreshFieldList
End Sub

Private Sub cmdRemoveOne_Click()
    On Error Resume Next
    lvSearch.ListItems.Remove lvSearch.ListItems.Count
    If iSSUbound > 0 Then
        iSSUbound = iSSUbound - 1
    End If
End Sub

Private Sub Form_Load()
    RefreshFieldList
End Sub




Private Function RefreshFieldList()
    
    Dim li As Integer
    Dim i As Integer
    
    'clear
 
    ReDim sFL(200)
    ReDim sCS(200)
    ReDim sFT(200)
    ReDim sD(200)
    ReDim sSearchStr(100)
    iSSUbound = 0
    
    
    li = 0
    
    sCS(li) = "School Year"
    sFT(li) = "s"
    sFL(li) = "tblSchoolYear.SchoolYearTitle"
    li = li + 1
    
    sCS(li) = "Department"
    sFT(li) = "s"
    sFL(li) = "tblDepartment.DepartmentTitle"
    li = li + 1
    
    sCS(li) = "Department ID"
    sFT(li) = "s"
    sFL(li) = "tblDepartment.DepartmentID"
    li = li + 1
    
    sCS(li) = "Year Level"
    sFT(li) = "s"
    sFL(li) = "tblYearLevel.YearLevelTitle"
    li = li + 1
    
    sCS(li) = "Section"
    sFT(li) = "s"
    sFL(li) = "tblSection.SectionTitle"
    li = li + 1
    
    sCS(li) = "Section ID"
    sFT(li) = "s"
    sFL(li) = "tblSection.SectionID"
    li = li + 1
    
    sCS(li) = "--------------------------------------"
    sFT(li) = ""
    sFL(li) = ""
    li = li + 1
    
    
    cmbFields.Clear
    For i = 0 To li - 1
        cmbFields.AddItem sCS(i)
    Next
    
    
    
    
    
    
    lvSearch.ListItems.Clear
    
    cmbOperand.Clear
    cmbOperand.AddItem "Equal"
    cmbOperand.AddItem "Not Equal"
    cmbOperand.AddItem "Contain"
    
    
    
    cmbFields.ListIndex = 0
    cmbOperand.ListIndex = 0
End Function
