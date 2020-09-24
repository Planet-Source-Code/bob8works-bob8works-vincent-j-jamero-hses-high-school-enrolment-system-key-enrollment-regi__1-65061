VERSION 5.00
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmDeleteSchoolYear 
   BackColor       =   &H00D8E9EC&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Delete Record"
   ClientHeight    =   2820
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5055
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
   ScaleHeight     =   188
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   337
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdGetSchoolYear 
      BackColor       =   &H00D8E9EC&
      Height          =   330
      Left            =   4470
      Picture         =   "frmDeleteSchoolYear.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   735
      Width           =   345
   End
   Begin VB.TextBox txtSchoolYear 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1110
      MaxLength       =   20
      TabIndex        =   0
      Top             =   735
      Width           =   3375
   End
   Begin HSES.b8Line b8Line1 
      Height          =   60
      Left            =   0
      TabIndex        =   3
      Top             =   510
      Width           =   5385
      _ExtentX        =   9499
      _ExtentY        =   106
   End
   Begin HSES.b8Line b8Line3 
      Height          =   60
      Left            =   0
      TabIndex        =   5
      Top             =   1140
      Width           =   5385
      _ExtentX        =   9499
      _ExtentY        =   106
      BorderColor1    =   12307149
      BorderColor2    =   14215660
      BorderColor3    =   14215660
      BorderStyle1    =   3
   End
   Begin HSES.b8Line b8Line2 
      Height          =   60
      Left            =   0
      TabIndex        =   8
      Top             =   2340
      Width           =   5385
      _ExtentX        =   9499
      _ExtentY        =   106
   End
   Begin lvButton.lvButtons_H cmdDelete 
      Default         =   -1  'True
      Height          =   360
      Left            =   3540
      TabIndex        =   9
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
      Left            =   2010
      TabIndex        =   10
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
      Caption         =   "School Year not found."
      ForeColor       =   &H00000000&
      Height          =   690
      Left            =   195
      TabIndex        =   7
      Top             =   1290
      Width           =   4755
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
      Left            =   195
      TabIndex        =   6
      Top             =   2055
      Visible         =   0   'False
      Width           =   3705
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Delete School Year"
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
      TabIndex        =   4
      Top             =   180
      Width           =   2700
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Title:"
      Height          =   195
      Left            =   210
      TabIndex        =   2
      Top             =   765
      Width           =   360
   End
   Begin VB.Image Image2 
      Height          =   525
      Left            =   30
      Picture         =   "frmDeleteSchoolYear.frx":058A
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5295
   End
End
Attribute VB_Name = "frmDeleteSchoolYear"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim RecordDeleted As Boolean

Dim curSchoolYear As String


Public Function ShowForm(Optional sSchoolYear As String = "") As Boolean
    
    '-------------------------------------------------------
    'check user access
    '-------------------------------------------------------
    If UserAllowedTo(CurrentUser.UserName, sCanDeleteSchoolYear) = False Then
        MsgBox "Unable to continue deleting School Year entry." & vbNewLine & vbNewLine & _
                "You are not permitted to do this. Please contact your administrator for more information.", vbExclamation
        Exit Function
    End If
    '-------------------------------------------------------



    curSchoolYear = sSchoolYear
    
    Me.Show vbModal
    
    ShowForm = RecordDeleted
End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdDelete_Click()
    If DeleteSchoolYear(txtSchoolYear.Text) = Success Then
        MsgBox "School Year record successfully deleted.", vbInformation
        
        RecordDeleted = True
        Unload Me
    Else
        MsgBox "Unable to delete School Year Record!", vbExclamation
    End If
        
End Sub

Private Sub cmdGetSchoolYear_Click()
    Dim sSchoolYear As String
    
    sSchoolYear = PickSchoolYear.GetItem(txtSchoolYear)
    
    If sSchoolYear <> "" Then
        curSchoolYear = sSchoolYear
        txtSchoolYear.Text = sSchoolYear
    End If
End Sub

Private Sub Form_Activate()
    txtSchoolYear.Text = curSchoolYear
End Sub

Private Sub txtSchoolYear_Change()
    
    cmdDelete.Enabled = False: lblAskMsg.Visible = False

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


    ShowSchoolYearDetail
End Sub

Private Sub ShowSchoolYearDetail()
    
    Dim sSQL As String
    Dim vRS As New ADODB.Recordset
    
    sSQL = "SELECT Count(tblSectionOffering.SectionOfferingID) AS CountOfSectionOfferingID" & _
            " FROM tblSchoolYear LEFT JOIN tblSectionOffering ON tblSchoolYear.SchoolYearTitle = tblSectionOffering.SchoolYear" & _
            " WHERE (((tblSchoolYear.SchoolYearTitle)='" & txtSchoolYear.Text & "'));"
                
    
    If ConnectRS(HSESDB, vRS, sSQL) = True Then
        If AnyRecordExisted(vRS) = True Then
            If ReadField(vRS.Fields("CountOfSectionOfferingID")) < 1 Then
                
                lblDetail.Caption = "Ready to delete this record."
                
                cmdDelete.Enabled = True: lblAskMsg.Visible = True
            Else
            
                lblDetail.Caption = "This School Year Record cannot be deleted." & vbNewLine & _
                "Reason: This record contain " & ReadField(vRS.Fields("CountOfSectionOfferingID")) & " Section Offering record/s."
                
                cmdDelete.Enabled = False: lblAskMsg.Visible = False
            End If
        Else
            lblDetail.Caption = "School Year not found."

            'school year not found
            cmdDelete.Enabled = False: lblAskMsg.Visible = False
        End If
    Else
        'fatal error
        CatchError "frmDeleteSchoolYear", "ShowSchoolYearDetail", "Error connecting School Year RS"
    End If
    
    Set vRS = Nothing
End Sub

