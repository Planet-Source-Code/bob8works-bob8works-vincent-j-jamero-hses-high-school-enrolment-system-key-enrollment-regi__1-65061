VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPrintTeachers 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Print Subject"
   ClientHeight    =   6600
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7815
   ControlBox      =   0   'False
   Icon            =   "frmPrintTeachers.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   440
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   521
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList imgListIco 
      Left            =   6360
      Top             =   3000
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
            Picture         =   "frmPrintTeachers.frx":08CA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox bgMain 
      Appearance      =   0  'Flat
      BackColor       =   &H00D8E9EC&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4995
      Left            =   90
      ScaleHeight     =   333
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   471
      TabIndex        =   0
      Top             =   480
      Width           =   7065
      Begin VB.PictureBox pbBGButton 
         Appearance      =   0  'Flat
         BackColor       =   &H00D8E9EC&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   525
         Left            =   30
         ScaleHeight     =   525
         ScaleWidth      =   6675
         TabIndex        =   3
         Top             =   360
         Width           =   6675
      End
      Begin MSComctlLib.ImageList icoHeader 
         Left            =   6240
         Top             =   3270
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   2
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPrintTeachers.frx":0E64
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPrintTeachers.frx":13FE
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin HSES.b8ChildTitleBar b8Title 
         Height          =   345
         Left            =   60
         TabIndex        =   1
         Top             =   30
         Width           =   6795
         _ExtentX        =   11668
         _ExtentY        =   609
         BackColor       =   12307149
         Caption         =   "Print Subjects"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontBold        =   -1  'True
         FontName        =   "Tahoma"
         FontSize        =   9.75
         ForeColor       =   12582912
         GradTheme       =   2
      End
      Begin MSComctlLib.ImageList ilRecordIco 
         Left            =   6360
         Top             =   1830
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
               Picture         =   "frmPrintTeachers.frx":1998
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ListView listRecord 
         Height          =   4275
         Left            =   0
         TabIndex        =   2
         Top             =   870
         Width           =   6105
         _ExtentX        =   10769
         _ExtentY        =   7541
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         HideColumnHeaders=   -1  'True
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         HotTracking     =   -1  'True
         _Version        =   393217
         Icons           =   "imgListIco"
         SmallIcons      =   "imgListIco"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   7911
         EndProperty
      End
   End
   Begin VB.Shape b8cMain 
      BorderColor     =   &H00C0C0C0&
      Height          =   2055
      Left            =   3840
      Top             =   3810
      Width           =   3315
   End
End
Attribute VB_Name = "frmPrintTeachers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Dim curTeacherID As String


Public Function ShowForm(Optional sTeacherID As String)
    On Error Resume Next
    
    curTeacherID = sTeacherID

    Me.Show

    
End Function








Public Sub RefreshReportList()
    
    listRecord.ListItems.Clear
        
    listRecord.ListItems.Add , "TeacherAccountList", "Teacher Acount Information List", 1, 1
    
    listRecord.ListItems.Add , "TeacherAccountListIndividual", "Teacher Acount Information List - Individual", 1, 1

End Sub


Public Sub Form_Activate()
    RefreshReportList
    
    mdiMain.RegMDIChild Me
End Sub

Public Sub Form_Load()
    listRecord.ColumnHeaders(1).Width = listRecord.Width - 6
End Sub



Private Sub Form_Resize()
    On Error Resume Next
    Me.ScaleMode = vbPixels
    
    
    
    b8cMain.Move Form_LeftMargin - 1, Form_TopMargin - 1, Me.ScaleWidth - (Form_LeftMargin - 1) * 2, Me.ScaleHeight - (Form_TopMargin - 1) * 2

    bgMain.Move Form_LeftMargin, Form_TopMargin, Me.ScaleWidth - Form_LeftMargin * 2, Me.ScaleHeight - Form_TopMargin * 2

    b8Title.Move 0, 0, bgMain.Width

    

    listRecord.Move listRecord.Left, pbBGButton.Top + pbBGButton.Height, bgMain.Width - (listRecord.Left * 2), Me.ScaleHeight - (pbBGButton.Top + pbBGButton.Height)
    
End Sub

Public Sub listRecord_DblClick()
    
    Select Case listRecord.SelectedItem.Key
        
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


