VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmAddGraduate 
   BackColor       =   &H00D8E9EC&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Graduates Students"
   ClientHeight    =   3420
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   6195
   Icon            =   "frmAddGraduate.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   228
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   413
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtSchoolYear 
      Height          =   330
      Left            =   1530
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   10
      Top             =   1230
      Width           =   2595
   End
   Begin MSComCtl2.DTPicker dtDateGraduated 
      Height          =   345
      Left            =   1530
      TabIndex        =   8
      Top             =   1620
      Width           =   2625
      _ExtentX        =   4630
      _ExtentY        =   609
      _Version        =   393216
      Format          =   20643841
      CurrentDate     =   38807
   End
   Begin VB.TextBox txtNote 
      Height          =   690
      Left            =   1530
      MaxLength       =   50
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Top             =   2010
      Width           =   3945
   End
   Begin VB.CommandButton cmdGetStudentID 
      BackColor       =   &H00D8E9EC&
      Height          =   270
      Left            =   5100
      Picture         =   "frmAddGraduate.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   870
      Width           =   345
   End
   Begin VB.TextBox txtStudentName 
      Height          =   330
      Left            =   1530
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   3
      Top             =   840
      Width           =   3945
   End
   Begin HSES.b8Line b8Line1 
      Height          =   60
      Left            =   0
      TabIndex        =   0
      Top             =   510
      Width           =   6195
      _ExtentX        =   10927
      _ExtentY        =   106
   End
   Begin HSES.b8Line b8Line2 
      Height          =   60
      Left            =   0
      TabIndex        =   1
      Top             =   2940
      Width           =   6345
      _ExtentX        =   11192
      _ExtentY        =   106
   End
   Begin lvButton.lvButtons_H cmdSave 
      Default         =   -1  'True
      Height          =   360
      Left            =   4710
      TabIndex        =   12
      Top             =   3030
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   635
      Caption         =   "&Save"
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
      cBack           =   16185592
   End
   Begin lvButton.lvButtons_H cmdCancel 
      Cancel          =   -1  'True
      Height          =   360
      Left            =   3180
      TabIndex        =   13
      Top             =   3030
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
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "School Year"
      Height          =   195
      Left            =   330
      TabIndex        =   11
      Top             =   1290
      Width           =   870
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Date Graduated"
      Height          =   195
      Left            =   300
      TabIndex        =   9
      Top             =   1680
      Width           =   1140
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Note:"
      Height          =   195
      Left            =   300
      TabIndex        =   7
      Top             =   2070
      Width           =   390
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Student"
      Height          =   195
      Left            =   300
      TabIndex        =   4
      Top             =   840
      Width           =   555
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Add New Graduate"
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
      Left            =   30
      TabIndex        =   2
      Top             =   150
      Width           =   2685
   End
   Begin VB.Image Image2 
      Height          =   525
      Left            =   -60
      Picture         =   "frmAddGraduate.frx":0B14
      Stretch         =   -1  'True
      Top             =   0
      Width           =   6315
   End
End
Attribute VB_Name = "frmAddGraduate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit




Dim RecordSaved As Boolean

Dim curStudentID As String

Public Function ShowForm(Optional sStudentID As String = "") As Boolean
    
    '-------------------------------------------------------
    'check user access
    '-------------------------------------------------------
    If UserAllowedTo(CurrentUser.UserName, sCanModifyGraduate) = False Then
        MsgBox "Unable to continue adding Graduate entry." & vbNewLine & vbNewLine & _
                "You are not permitted to do this. Please contact your administrator for more information.", vbExclamation
        Exit Function
    End If
    '-------------------------------------------------------

    
        curStudentID = sStudentID
        
        'show form
        Me.Show vbModal
        
        'return
        ShowForm = RecordSaved
End Function


Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdGetStudentID_Click()
    Dim sStudentID As String
    Dim sStudentFullName As String
    Dim latestSchoolYear As String
    Dim latestYearLevelID As Integer
    
    'set mouse pointer
    Me.MousePointer = vbHourglass
    
    sStudentID = PickGraduate.GetStudentID(txtStudentName, sStudentFullName)
    
    If sStudentID <> "" Then
        curStudentID = sStudentID
        GetLatestSchoolYearYearLevel sStudentID, latestSchoolYear, latestYearLevelID
        txtSchoolYear.Text = latestSchoolYear
    End If
        
        txtStudentName.Text = sStudentFullName
        
    'restore mouse pointer
    Me.MousePointer = vbDefault
End Sub

Private Sub cmdSave_Click()
    Dim vGraduate As tGraduate
    
    'check
    If Len(curStudentID) < 1 Then
        MsgBox "Please select student.", vbExclamation
        
        cmdGetStudentID.SetFocus
    
        Exit Sub
    End If
    
    If Val(Year(dtDateGraduated.Value)) < Val(Right(txtSchoolYear.Text, 4)) Then
        MsgBox "Invalid Date Graduated.", vbExclamation
        
        dtDateGraduated.SetFocus
        
        Exit Sub
    End If
        
    
    
    vGraduate.StudentID = curStudentID
    vGraduate.SchoolYear = txtSchoolYear.Text
    vGraduate.DateGraduated = dtDateGraduated.Value
    vGraduate.Note = txtNote.Text
    vGraduate.CreationDate = Now
    vGraduate.CreatedBy = CurrentUser.UserName
    
    Select Case AddGraduate(vGraduate)
        
        Case TranDBResult.Success
                
            'success
            MsgBox "The select Student entry successfully added to Graduates Record.", vbInformation
            
            'set flag
            RecordSaved = True
            
            'close and return
            Unload Me
            
            
        Case TranDBResult.DuplicateID
            MsgBox "The selected Student entry is already exist in Graduates Record." & vbNewLine & _
                    "Please select another entry.", vbExclamation
                    
            cmdGetStudentID.SetFocus
            
        Case Else
            'fatal error
            MsgBox "Unable to save New Graduate Entry.", vbCritical
            CatchError "frmAddGraduate", "cmdSave_click", "AddGraduate return unknown error."
    End Select
    
End Sub

Private Sub Form_Load()
    
    'set default for Date Graduated = cur date
    
    dtDateGraduated.Value = Now
End Sub

