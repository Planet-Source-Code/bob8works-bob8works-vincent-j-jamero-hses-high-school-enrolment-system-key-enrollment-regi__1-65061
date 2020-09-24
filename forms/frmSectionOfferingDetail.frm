VERSION 5.00
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmSectionOfferingDetail 
   BackColor       =   &H00D8E9EC&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Section Offering Details"
   ClientHeight    =   6180
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10020
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   412
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   668
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox bgSectionOfferingDetail 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4965
      Left            =   60
      ScaleHeight     =   331
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   657
      TabIndex        =   5
      Top             =   600
      Width           =   9855
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CreationDate:"
         ForeColor       =   &H00808080&
         Height          =   195
         Left            =   4020
         TabIndex        =   29
         Top             =   240
         Width           =   1020
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Created By:"
         ForeColor       =   &H00808080&
         Height          =   195
         Left            =   4020
         TabIndex        =   28
         Top             =   540
         Width           =   870
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Modified Date:"
         ForeColor       =   &H00808080&
         Height          =   195
         Left            =   4020
         TabIndex        =   27
         Top             =   840
         Width           =   1050
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "--"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   240
         Index           =   11
         Left            =   5370
         TabIndex        =   26
         Top             =   1140
         Width           =   180
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "--"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   240
         Index           =   10
         Left            =   5370
         TabIndex        =   25
         Top             =   840
         Width           =   180
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "--"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   240
         Index           =   9
         Left            =   5370
         TabIndex        =   24
         Top             =   540
         Width           =   180
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Modified By:"
         ForeColor       =   &H00808080&
         Height          =   195
         Left            =   4020
         TabIndex        =   23
         Top             =   1170
         Width           =   885
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "--"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   240
         Index           =   8
         Left            =   5370
         TabIndex        =   22
         Top             =   240
         Width           =   180
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "--"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   240
         Index           =   7
         Left            =   1680
         TabIndex        =   21
         Top             =   2340
         Width           =   180
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Note"
         ForeColor       =   &H00808080&
         Height          =   195
         Left            =   180
         TabIndex        =   20
         Top             =   2340
         Width           =   345
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "--"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   240
         Index           =   6
         Left            =   1680
         TabIndex        =   19
         Top             =   2010
         Width           =   180
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "--"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   240
         Index           =   5
         Left            =   1680
         TabIndex        =   18
         Top             =   1680
         Width           =   180
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "--"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   240
         Index           =   4
         Left            =   1680
         TabIndex        =   17
         Top             =   1380
         Width           =   180
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "--"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   240
         Index           =   3
         Left            =   1680
         TabIndex        =   16
         Top             =   1080
         Width           =   180
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "--"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   240
         Index           =   2
         Left            =   1680
         TabIndex        =   15
         Top             =   810
         Width           =   180
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "--"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   240
         Index           =   1
         Left            =   1680
         TabIndex        =   14
         Top             =   510
         Width           =   180
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "--"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   240
         Index           =   0
         Left            =   1680
         TabIndex        =   13
         Top             =   210
         Width           =   180
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Max. Grade"
         ForeColor       =   &H00808080&
         Height          =   195
         Left            =   180
         TabIndex        =   12
         Top             =   2010
         Width           =   825
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Min. Grade"
         ForeColor       =   &H00808080&
         Height          =   195
         Left            =   180
         TabIndex        =   11
         Top             =   1710
         Width           =   780
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Max. Student #"
         ForeColor       =   &H00808080&
         Height          =   195
         Left            =   180
         TabIndex        =   10
         Top             =   1410
         Width           =   1095
      End
      Begin VB.Label TeacherName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TeacherName"
         ForeColor       =   &H00808080&
         Height          =   195
         Left            =   180
         TabIndex        =   9
         Top             =   1110
         Width           =   1020
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "School Year"
         ForeColor       =   &H00808080&
         Height          =   195
         Left            =   180
         TabIndex        =   8
         Top             =   840
         Width           =   870
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Section Title"
         ForeColor       =   &H00808080&
         Height          =   195
         Left            =   180
         TabIndex        =   7
         Top             =   540
         Width           =   885
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Section Offering ID:"
         ForeColor       =   &H00808080&
         Height          =   195
         Left            =   180
         TabIndex        =   6
         Top             =   240
         Width           =   1440
      End
   End
   Begin HSES.b8Line b8Line1 
      Height          =   60
      Left            =   15
      TabIndex        =   0
      Top             =   510
      Width           =   10035
      _ExtentX        =   17701
      _ExtentY        =   106
   End
   Begin lvButton.lvButtons_H cmdSave 
      Default         =   -1  'True
      Height          =   360
      Left            =   8520
      TabIndex        =   1
      Top             =   5730
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   635
      Caption         =   "&Save"
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
   Begin lvButton.lvButtons_H cmdCancel 
      Cancel          =   -1  'True
      Height          =   360
      Left            =   6870
      TabIndex        =   2
      Top             =   5730
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
   Begin HSES.b8Line b8Line2 
      Height          =   60
      Left            =   0
      TabIndex        =   3
      Top             =   5610
      Width           =   10035
      _ExtentX        =   17701
      _ExtentY        =   106
   End
   Begin HSES.b8Container b 
      Height          =   5085
      Left            =   0
      TabIndex        =   30
      Top             =   540
      Width           =   10005
      _ExtentX        =   17648
      _ExtentY        =   8969
      BorderColor     =   12307149
      BackColor       =   16185592
      ShadowColor1    =   13427430
      ShadowColor2    =   14215660
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Section Offering Details"
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
      Left            =   615
      TabIndex        =   4
      Top             =   120
      Width           =   2280
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   45
      Picture         =   "frmSectionOfferingDetail.frx":0000
      Top             =   30
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   525
      Left            =   15
      Picture         =   "frmSectionOfferingDetail.frx":08CA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   10035
   End
End
Attribute VB_Name = "frmSectionOfferingDetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim curSectionOfferingID As String

Public Function ShowForm(Optional sSectionOfferingID As String = "")
    
    curSectionOfferingID = sSectionOfferingID
    
    
    'show form
    Me.Show vbModal
End Function

Private Sub Form_Activate()
    ShowSectionOfferingDetail
    ShowSubjects

End Sub

Private Sub ShowSectionOfferingDetail()
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    sSQL = "SELECT tblSectionOffering.SectionOfferingID, [tblYearLevel]![YearLevelTitle] & ' - ' & [tblSection]![SectionTitle] AS SectionFullTitle, tblSectionOffering.SchoolYear, [tblTeacher]![LastName] & ', ' & [tblTeacher]![FirstName] & ' ' & [tblTeacher]![MiddleName] AS TeacherFullName, tblSectionOffering.MaxStudentCount, tblSectionOffering.MinGrade, tblSectionOffering.MaxGrade, tblSectionOffering.Note, tblSectionOffering.CreationDate, tblSectionOffering.CreatedBy, tblSectionOffering.ModifiedDate, tblSectionOffering.ModifiedBy" & _
            " FROM tblTeacher INNER JOIN (tblYearLevel INNER JOIN (tblSection INNER JOIN tblSectionOffering ON tblSection.SectionID = tblSectionOffering.SectionID) ON tblYearLevel.YearLevelID = tblSection.YearLevelID) ON tblTeacher.TeacherID = tblSectionOffering.TeacherID" & _
            " WHERE (((tblSectionOffering.SectionOfferingID)='" & curSectionOfferingID & "'));"

    If ConnectRS(DB, vRS, sSQL) = False Then
        'fatal error
        CatchError "frmSectionOfferingDetail", "ShowSectionOfferingDetail", "Unable to connect Recordset"
        GoTo ReleaseAndExit
    End If
    
        
    If AnyRecordExisted(vRS) = False Then
        Call SectionOffering_NotFound
        'not found
        GoTo ReleaseAndExit
    End If

    'set form fields
    lblInfo(0).Caption = ReadField(vRS.Fields("SectionOfferingID"))
    lblInfo(1).Caption = ReadField(vRS.Fields("SectionFullTitle"))
    lblInfo(2).Caption = ReadField(vRS.Fields("SchoolYear"))
    lblInfo(3).Caption = ReadField(vRS.Fields("TeacherFullName"))
    lblInfo(4).Caption = ReadField(vRS.Fields("MaxStudentCount"))
    lblInfo(5).Caption = ReadField(vRS.Fields("MinGrade"))
    lblInfo(6).Caption = ReadField(vRS.Fields("MaxGrade"))
    lblInfo(7).Caption = ReadField(vRS.Fields("Note"))
    lblInfo(8).Caption = ReadField(vRS.Fields("CreationDate"))
    lblInfo(9).Caption = ReadField(vRS.Fields("CreatedBy"))
    If Len(ReadField(vRS.Fields("ModifiedBy"))) > 0 Then
        lblInfo(10).Caption = ReadField(vRS.Fields("ModifiedDate"))
        lblInfo(11).Caption = ReadField(vRS.Fields("ModifiedBy"))
    End If
'------------------------------------------------------------
ReleaseAndExit:
    Set vRS = Nothing
End Sub

Private Sub SectionOffering_NotFound()
    
End Sub
Private Sub ShowSubjects()

    
    
End Sub

Private Sub lblStudentFullName_Click()

End Sub
