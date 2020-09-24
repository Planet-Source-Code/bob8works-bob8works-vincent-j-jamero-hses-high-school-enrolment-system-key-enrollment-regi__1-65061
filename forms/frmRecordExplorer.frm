VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmRecordExplorer 
   Caption         =   "Record Explorer"
   ClientHeight    =   6240
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7215
   LinkTopic       =   "Form1"
   ScaleHeight     =   6240
   ScaleWidth      =   7215
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ImageList imgListEnrolment 
      Left            =   3615
      Top             =   2460
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRecordExplorer.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRecordExplorer.frx":059A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRecordExplorer.frx":0B34
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRecordExplorer.frx":10CE
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRecordExplorer.frx":1668
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRecordExplorer.frx":1C02
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView tvFolder 
      Height          =   5550
      Left            =   60
      TabIndex        =   0
      Top             =   195
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   9790
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   423
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      ImageList       =   "imgListEnrolment"
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
   End
End
Attribute VB_Name = "frmRecordExplorer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const keySchoolYear = "scho"
Private Const keyDepartment = "dept"
Private Const keyYearLevel = "year"
Private Const keySection = "sect"


Dim slSchoolYearTitle() As String 'holds sy title
Dim slDepartmentTitle() As String 'holds sy title
Dim slYearLevelTitle() As String 'holds sy title
'----------------------------------------------------------
'START UP
'----------------------------------------------------------
Public Function ShowForm(Optional sSectionTitle As String = "", Optional sSchoolYearTitle As String = "")
    
    
    'refresh section tree
    Refresh_Tree
    
    'set parameter
    If sSectionTitle <> "" Then
        SetSelectedSection sSectionTitle, sSchoolYearTitle
    End If
    
    'show form
    Me.Show vbModal
End Function




Private Function Refresh_Tree()
    'add school year
    Refresh_SchoolYear
    'add Department
    Refresh_Department
    'add year level
    Refresh_YearLevel
    'add section
    Refresh_Section
End Function

Private Function Refresh_SchoolYear()
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    Dim i As Integer
    
    'clear tree
    tvFolder.Nodes.Clear
    
    sSQL = "SELECT tblSchoolYear.SchoolYearTitle" & _
            " FROM tblSchoolYear;"
    
    If ConnectRS(DB, vRS, sSQL) <> True Then
        GoTo RealeaseAndExit
    End If
    
    If AnyRecordExisted(vRS) = False Then
        GoTo RealeaseAndExit
    End If
    
    ReDim slSchoolYearTitle(getRecordCount(vRS) - 1)
    
    i = 0
    vRS.MoveFirst
    
    While vRS.EOF = False
        
        slSchoolYearTitle(i) = ReadField(vRS.Fields("SchoolYearTitle"))
        AddSchoolYearToTree slSchoolYearTitle(i)
        
        
        i = i + 1
        vRS.MoveNext
    Wend
    
    
    
    
RealeaseAndExit:
    Set vRS = Nothing
End Function


Private Function Refresh_Department()

    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    Dim i As Integer
    Dim ii As Integer
    
    
    
    
    sSQL = "SELECT tblDepartment.DepartmentTitle" & _
            " FROM tblDepartment"

    If ConnectRS(DB, vRS, sSQL) <> True Then
        GoTo RealeaseAndExit
    End If
    
    If AnyRecordExisted(vRS) = False Then
        GoTo RealeaseAndExit
    End If
        
    ReDim slDepartmentTitle(getRecordCount(vRS) - 1)
    
    i = 0
    vRS.MoveFirst
    
    While vRS.EOF = False
    
        slDepartmentTitle(i) = ReadField(vRS.Fields("DepartmentTitle"))
        
        For ii = 0 To UBound(slSchoolYearTitle)
            AddDepartmentToTree slSchoolYearTitle(ii), slDepartmentTitle(i)
        Next
        
        i = i + 1
        vRS.MoveNext
    Wend
        
    
RealeaseAndExit:
    Set vRS = Nothing
End Function


Private Function Refresh_YearLevel()

    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    Dim i As Integer
    Dim ii As Integer
    Dim iii As Integer
    
    
    
    sSQL = "SELECT tblYearLevel.YearLevelTitle" & _
            " FROM tblYearLevel"

    If ConnectRS(DB, vRS, sSQL) <> True Then
        GoTo RealeaseAndExit
    End If
    
    If AnyRecordExisted(vRS) = False Then
        GoTo RealeaseAndExit
    End If
        
    ReDim slYearLevelTitle(getRecordCount(vRS))
    
    i = 0
    vRS.MoveFirst
    
    While vRS.EOF = False
        
        slYearLevelTitle(i) = ReadField(vRS.Fields("YearLevelTitle"))
        
        For ii = 0 To UBound(slSchoolYearTitle)
            For iii = 0 To UBound(slDepartmentTitle)
                AddYearLevelToTree slSchoolYearTitle(ii), slDepartmentTitle(iii), slYearLevelTitle(i)
            Next
        Next
        
        i = i + 1
        
        vRS.MoveNext
    Wend
        
    
RealeaseAndExit:
    Set vRS = Nothing
End Function


Private Function Refresh_Section()

    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    Dim i As Integer
    
    
    
    
    sSQL = "SELECT tblDepartment.DepartmentTitle, tblYearLevel.YearLevelTitle, tblSection.SectionTitle" & _
            " FROM tblYearLevel INNER JOIN (tblDepartment INNER JOIN tblSection ON tblDepartment.DepartmentID = tblSection.DepartmentID) ON tblYearLevel.YearLevelID = tblSection.YearLevelID;"

    If ConnectRS(DB, vRS, sSQL) <> True Then
        GoTo RealeaseAndExit
    End If
    
    If AnyRecordExisted(vRS) = False Then
        GoTo RealeaseAndExit
    End If
        
    
    vRS.MoveFirst
    
    While vRS.EOF = False
        For i = 0 To UBound(slSchoolYearTitle)
            AddSectionToTree slSchoolYearTitle(i), ReadField(vRS.Fields("DepartmentTitle")), ReadField(vRS.Fields("YearLevelTitle")), ReadField(vRS.Fields("sectionTitle"))
        Next
        vRS.MoveNext
    Wend
        
    
RealeaseAndExit:
    Set vRS = Nothing
End Function


Private Function AddSchoolYearToTree(sSchoolYearTitle As String)
    Dim tNode As Node
    
    For Each tNode In tvFolder.Nodes
        If tNode.Key = keySchoolYear & ";" & sSchoolYearTitle Then
            Exit Function
        End If
    Next
    
    tvFolder.Nodes.Add , , keySchoolYear & ";" & sSchoolYearTitle, sSchoolYearTitle, 3
End Function



Private Function AddDepartmentToTree(sSchoolYearTitle As String, sDepartmentTitle As String)
    Dim tNode As Node
    
    For Each tNode In tvFolder.Nodes
        If tNode.Key = keyDepartment & ";" & sSchoolYearTitle & ";" & sDepartmentTitle Then
            Exit Function
        End If
    Next
    
    tvFolder.Nodes.Add keySchoolYear & ";" & sSchoolYearTitle, tvwChild, keyDepartment & ";" & sSchoolYearTitle & ";" & sDepartmentTitle, sDepartmentTitle, 3
End Function


Private Function AddYearLevelToTree(sSchoolYearTitle As String, sDepartmentTitle As String, sYearLevelTitle As String)
    Dim tNode As Node
    
    For Each tNode In tvFolder.Nodes
        If tNode.Key = keyYearLevel & ";" & sSchoolYearTitle & ";" & sDepartmentTitle & ";" & sYearLevelTitle Then
            Exit Function
        End If
    Next
    
    tvFolder.Nodes.Add keyDepartment & ";" & sSchoolYearTitle & ";" & sDepartmentTitle, tvwChild, keyYearLevel & ";" & sSchoolYearTitle & ";" & sDepartmentTitle & ";" & sYearLevelTitle, sYearLevelTitle, 3

End Function


Private Function AddSectionToTree(sSchoolYearTitle As String, sDepartmentTitle As String, sYearLevelTitle As String, sSectionTitle As String)
    Dim tNode As Node
    
    For Each tNode In tvFolder.Nodes
        If tNode.Key = keySection & ";" & sSchoolYearTitle & ";" & sDepartmentTitle & ";" & sYearLevelTitle & ";" & sSectionTitle Then
            Exit Function
        End If
    Next
    
    tvFolder.Nodes.Add keyYearLevel & ";" & sSchoolYearTitle & ";" & sDepartmentTitle & ";" & sYearLevelTitle, tvwChild, keySection & ";" & sSchoolYearTitle & ";" & sDepartmentTitle & ";" & sYearLevelTitle & ";" & sSectionTitle, sSectionTitle, 3

End Function

Private Function SetSelectedSection(sSectionTitle As String, Optional sSchoolYearTitle As String = "")
    Dim tNode As Node
    Dim splitKey() As String


    For Each tNode In tvFolder.Nodes
    
        If tNode.Text = sSectionTitle And Left(tNode.Key, 4) = keySection Then
            
            splitKey = Split(tNode.Key, ";")
            
            If sSchoolYearTitle = "" Then
                tNode.Selected = True
                tNode.EnsureVisible
                Exit For
            Else
            
                If splitKey(1) = sSchoolYearTitle Then
                    tNode.Selected = True
                    tNode.EnsureVisible
                    Exit For
                End If
            End If
            
        End If
        
    Next
End Function

