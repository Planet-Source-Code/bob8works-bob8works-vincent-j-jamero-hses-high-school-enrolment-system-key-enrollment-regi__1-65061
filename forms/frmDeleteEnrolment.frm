VERSION 5.00
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmDeleteEnrolment 
   BackColor       =   &H00D8E9EC&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Delete"
   ClientHeight    =   2820
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5115
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2820
   ScaleWidth      =   5115
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin HSES.b8Line b8Line1 
      Height          =   60
      Left            =   0
      TabIndex        =   0
      Top             =   510
      Width           =   5385
      _ExtentX        =   9499
      _ExtentY        =   106
   End
   Begin HSES.b8Line b8Line2 
      Height          =   60
      Left            =   0
      TabIndex        =   1
      Top             =   2340
      Width           =   5385
      _ExtentX        =   9499
      _ExtentY        =   106
   End
   Begin lvButton.lvButtons_H cmdDelete 
      Default         =   -1  'True
      Height          =   360
      Left            =   3600
      TabIndex        =   5
      Top             =   2430
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   635
      Caption         =   "&Delete"
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
      cBhover         =   14215660
      Focus           =   0   'False
      cGradient       =   14215660
      Gradient        =   3
      Mode            =   0
      Value           =   0   'False
      Enabled         =   0   'False
      cBack           =   16185592
   End
   Begin lvButton.lvButtons_H cmdCancel 
      Cancel          =   -1  'True
      Height          =   360
      Left            =   2070
      TabIndex        =   6
      Top             =   2430
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   635
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
      cBhover         =   14215660
      Focus           =   0   'False
      cGradient       =   14215660
      Gradient        =   3
      Mode            =   0
      Value           =   0   'False
      cBack           =   16185592
   End
   Begin VB.Label lblDetail 
      BackStyle       =   0  'Transparent
      Caption         =   "Fee not found."
      ForeColor       =   &H00000000&
      Height          =   1380
      Left            =   150
      TabIndex        =   4
      Top             =   660
      Width           =   4755
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Delete Student"
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
      Left            =   90
      TabIndex        =   3
      Top             =   180
      Width           =   2130
   End
   Begin VB.Label lblAskMsg 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Are You Sure Want To Delete This Record?"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Left            =   210
      TabIndex        =   2
      Top             =   2055
      Visible         =   0   'False
      Width           =   3705
   End
   Begin VB.Image Image2 
      Height          =   525
      Left            =   30
      Picture         =   "frmDeleteEnrolment.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5295
   End
End
Attribute VB_Name = "frmDeleteEnrolment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim curEnrolmentID As String

Private RecordDeleted As Boolean

Public Function ShowForm(sEnrolmentID As String) As Boolean
    
    '-------------------------------------------------------
    'check user access
    '-------------------------------------------------------
    If UserAllowedTo(CurrentUser.UserName, sCanDeleteEnrolment) = False Then
        MsgBox "Unable to continue deleting Enrolment entry." & vbNewLine & vbNewLine & _
                "You are not permitted to do this. Please contact your administrator for more information.", vbExclamation
        Exit Function
    End If
    '-------------------------------------------------------

    curEnrolmentID = sEnrolmentID
    
    'show form
    Me.Show vbModal
    
    'return
    ShowForm = RecordDeleted
End Function

Private Sub ShowDetail()
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    Dim sSY As String
    Dim iYLID As Integer
    
    If GetLatestSchoolYearYearLevel(Right(curEnrolmentID, 12), sSY, iYLID) <> Success Then
        'fatal error
        'temp
        MsgBox "error"
        GoTo ReleaseAndExit
    End If
    
    '[ 1 ]
    'connect Recordset and get Enrolment/s count
    sSQL = "SELECT tblEnrolment.EnrolmentID, tblSection.YearLevelID" & _
            " FROM tblSection INNER JOIN (tblSectionOffering INNER JOIN tblEnrolment ON tblSectionOffering.SectionOfferingID = tblEnrolment.SectionOfferingID) ON tblSection.SectionID = tblSectionOffering.SectionID" & _
            " WHERE (((tblEnrolment.EnrolmentID)='" & curEnrolmentID & "'));"
            

    If ConnectRS(HSESDB, vRS, sSQL) = False Then
        'fatal error
        CatchError "frmDeleteEnrolment", "ShowDetail", "Unable to connect RS Enrolments count."
        'close this form
        Unload Me
        GoTo ReleaseAndExit
    End If
    
    '[ 2 ]
    'set form detail
    If iYLID <= ReadField(vRS.Fields("YearLevelID")) Then
    
        'ready to delete
        lblDetail.Visible = False
        lblAskMsg.Visible = True
        
        cmdDelete.Enabled = True
        
    Else
    
        'cannot be deleted
        lblDetail.Caption = "This Enrolment entry cannot be deleted." & vbNewLine & _
                "Reason: Selected Student is already enroled at Next Level."


        lblDetail.Visible = True
        lblAskMsg.Visible = False
        
        cmdDelete.Enabled = True
        
    End If
    
ReleaseAndExit:
    Set vRS = Nothing
End Sub


Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdDelete_Click()

    If DeleteEnrolment(curEnrolmentID) = Success Then
        
        'set flag
        RecordDeleted = True
        'close this form
        Unload Me
    
    Else
        MsgBox "Unable to delete Enrolment entry.", vbCritical
        
        'fatal error
        CatchError "frmDeleteEnrolment", "cmdDelete_Click", "Unable to delete Enrolment entry"
    
        'close ths form
        Unload Me
    End If
    
End Sub

Private Sub Form_Activate()
    ShowDetail
End Sub

