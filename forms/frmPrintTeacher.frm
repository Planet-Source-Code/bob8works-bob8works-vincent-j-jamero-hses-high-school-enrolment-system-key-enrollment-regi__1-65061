VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmPrintTeacher 
   BackColor       =   &H00D8E9EC&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Print Teacher Information"
   ClientHeight    =   5385
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6165
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   359
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   411
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList imgListIco 
      Left            =   5460
      Top             =   150
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrintTeacher.frx":0000
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView listReports 
      Height          =   4275
      Left            =   30
      TabIndex        =   4
      Top             =   600
      Width           =   6105
      _ExtentX        =   10769
      _ExtentY        =   7541
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HotTracking     =   -1  'True
      _Version        =   393217
      Icons           =   "imgListIco"
      SmallIcons      =   "imgListIco"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   7911
      EndProperty
   End
   Begin lvButton.lvButtons_H cmdCancel 
      Cancel          =   -1  'True
      Height          =   360
      Left            =   4680
      TabIndex        =   0
      Top             =   4980
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   635
      Caption         =   "&Close"
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
   Begin HSES.b8Line b8Line2 
      Height          =   60
      Left            =   0
      TabIndex        =   1
      Top             =   4890
      Width           =   6315
      _ExtentX        =   11139
      _ExtentY        =   106
   End
   Begin HSES.b8Line b8Line1 
      Height          =   60
      Left            =   0
      TabIndex        =   2
      Top             =   510
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   106
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   60
      Picture         =   "frmPrintTeacher.frx":059A
      Top             =   30
      Width           =   480
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Print What?"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H002F556A&
      Height          =   240
      Left            =   630
      TabIndex        =   3
      Top             =   120
      Width           =   1155
   End
   Begin VB.Image Image2 
      Height          =   525
      Left            =   30
      Picture         =   "frmPrintTeacher.frx":0E64
      Stretch         =   -1  'True
      Top             =   0
      Width           =   6465
   End
End
Attribute VB_Name = "frmprintTeacher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Dim curTeacherID As String


Public Function ShowForm(Optional sTeacherID As String)

    curTeacherID = sTeacherID

    Me.Show vbModal
    
End Function




Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    RefreshReportList
End Sub

Private Sub Form_Load()
    listReports.ColumnHeaders(1).Width = listReports.Width - 6
End Sub

Private Sub RefreshReportList()
    
    listReports.ListItems.Clear
        
    listReports.ListItems.Add , "TeacherAccountList", "Teacher Acount Information List", 1, 1
    
    listReports.ListItems.Add , "TeacherAccountListIndividual", "Teacher Acount Information List - Individual", 1, 1

    
    
End Sub


Private Sub listReports_DblClick()
    
    Select Case listReports.SelectedItem.Key
        
        Case "TeacherAccountList"
            ShowTeacherAccountList
            
        Case "TeacherAccountListIndividual"
            ShowTeacherAccountListIndividual
    End Select
        
    
End Sub

Private Sub ShowTeacherAccountList()
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    sSQL = " SELECT tblTeacher.TeacherID, tblTeacher.TeacherTitle, [tblTeacher].[LastName] & ', ' & [tblTeacher].[FirstName] & ' ' & [tblTeacher].[MiddleName] AS TeacherFullName, tblTeacher.Address, tblTeacher.ContactNumber, tblTeacher.CreationDate, tblTeacher.CreatedBy, tblTeacher.ModifiedDate, tblTeacher.ModifiedBy" & _
            " FROM tblTeacher;"

    If ConnectRS(HSESDB, vRS, sSQL) = False Then
        MsgBox "Error"
    
        GoTo ReleaseAndExit
    End If
    
    Set drTeacherList.DataSource = vRS
    
    drTeacherList.Show vbModal
    
    
ReleaseAndExit:
    Set vRS = Nothing
End Sub





Private Sub ShowTeacherAccountListIndividual()
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    Dim sTeacherID As String
    
    If curTeacherID = "" Then
        sTeacherID = PickTeacher.GetTeacherID
    Else
        sTeacherID = curTeacherID
    End If
    
    If sTeacherID = "" Then
        GoTo ReleaseAndExit
    End If
    
    sSQL = " SELECT tblTeacher.TeacherID, tblTeacher.TeacherTitle, [tblTeacher].[LastName] & ', ' & [tblTeacher].[FirstName] & ' ' & [tblTeacher].[MiddleName] AS TeacherFullName, tblTeacher.Address, tblTeacher.ContactNumber, tblTeacher.CreationDate, tblTeacher.CreatedBy, tblTeacher.ModifiedDate, tblTeacher.ModifiedBy" & _
            " FROM tblTeacher " & _
            " WHERE tblTeacher.TeacherID='" & sTeacherID & "'" & _
            " ORDER BY [tblTeacher].[LastName] & ', ' & [tblTeacher].[FirstName] & ' ' & [tblTeacher].[MiddleName]"

    If ConnectRS(HSESDB, vRS, sSQL) = False Then
        MsgBox "Error"
    
        GoTo ReleaseAndExit
    End If
    
    Set drTeacherList.DataSource = vRS
    
    drTeacherList.Show vbModal
    
    
ReleaseAndExit:
    Set vRS = Nothing
End Sub
