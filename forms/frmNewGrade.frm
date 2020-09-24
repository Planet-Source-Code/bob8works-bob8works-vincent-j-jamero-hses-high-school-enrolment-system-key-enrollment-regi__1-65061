VERSION 5.00
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "LVbuttons.ocx"
Begin VB.Form frmNewGrade 
   BackColor       =   &H00D8E9EC&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Enter New Grade"
   ClientHeight    =   2280
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3555
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
   ScaleHeight     =   152
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   237
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin lvButton.lvButtons_H cmdSave 
      Default         =   -1  'True
      Height          =   405
      Left            =   2070
      TabIndex        =   4
      Top             =   1830
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   714
      Caption         =   "&Save"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Focus           =   0   'False
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      cBack           =   14215660
   End
   Begin HSES.b8Container b8Container1 
      Height          =   855
      Left            =   90
      TabIndex        =   3
      Top             =   900
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   1508
      BackColor       =   16185592
      Begin VB.TextBox txtGrade 
         Height          =   345
         Left            =   1410
         TabIndex        =   6
         Text            =   "00"
         Top             =   270
         Width           =   1185
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "New Grade:"
         ForeColor       =   &H0030A0B8&
         Height          =   195
         Left            =   420
         TabIndex        =   7
         Top             =   300
         Width           =   855
      End
   End
   Begin HSES.b8Line b8Line1 
      Height          =   60
      Left            =   30
      TabIndex        =   2
      Top             =   750
      Width           =   3465
      _ExtentX        =   6112
      _ExtentY        =   106
   End
   Begin lvButton.lvButtons_H cmdCancel 
      Cancel          =   -1  'True
      Height          =   405
      Left            =   540
      TabIndex        =   5
      Top             =   1800
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   714
      Caption         =   "&Cancel"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Focus           =   0   'False
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      cBack           =   14215660
   End
   Begin VB.Label lblSubjectTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "--"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   195
      Left            =   840
      TabIndex        =   9
      Top             =   480
      Width           =   150
   End
   Begin VB.Label lblStudentName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "--"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   195
      Left            =   840
      TabIndex        =   8
      Top             =   180
      Width           =   150
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Subject:"
      ForeColor       =   &H0030A0B8&
      Height          =   195
      Left            =   150
      TabIndex        =   1
      Top             =   480
      Width           =   600
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Student:"
      ForeColor       =   &H0030A0B8&
      Height          =   195
      Left            =   150
      TabIndex        =   0
      Top             =   180
      Width           =   630
   End
   Begin VB.Image Image4 
      Height          =   1530
      Left            =   0
      Picture         =   "frmNewGrade.frx":0000
      Stretch         =   -1  'True
      Top             =   225
      Width           =   3630
   End
   Begin VB.Image Image3 
      Height          =   345
      Left            =   -60
      Picture         =   "frmNewGrade.frx":009D
      Stretch         =   -1  'True
      Top             =   1755
      Width           =   3705
   End
   Begin VB.Image Image5 
      Height          =   105
      Left            =   -60
      Picture         =   "frmNewGrade.frx":013A
      Stretch         =   -1  'True
      Top             =   2205
      Width           =   3660
   End
End
Attribute VB_Name = "frmNewGrade"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RecordSaved As Boolean

Dim curGrade As tGrade
Dim curStudent As tStudent

Public Function ShowForm(sGradeID As String) As Boolean
    
    Dim sSubjectTitle As String
    Dim dGradeValue As Double
    Dim sStudentName As String
    Dim sYearLevelTitle As String
            
    Dim isEditable As Boolean
    'set defaults
    RecordSaved = False
    
    If GetGradeByID(sGradeID, curGrade) <> success Then
        MsgBox "Invalid Grade Entry. Grade ID not found.", vbCritical
        Unload Me
        Exit Function
    End If
    
    If GetGradeInfoByID(sGradeID, sSubjectTitle, dGradeValue, sStudentName, sYearLevelTitle) <> success Then
        MsgBox "Invalid Grade Entry. Grade Info not found.", vbCritical
        Unload Me
        Exit Function
    End If
    
    
    
    lblStudentName.Caption = sStudentName
    lblSubjectTitle.Caption = sSubjectTitle
    txtGrade.Text = dGradeValue
    
    
    'show form ------------------------------
    Me.Show vbModal
    
    'return ---------------------------------
    ShowForm = RecordSaved
End Function

Private Function SaveRecord() As Boolean
    'default
    SaveRecord = False
    
    'check values
    If IsNumeric(txtGrade.Text) = True Then
        If Val(txtGrade.Text) < 60 Or Val(txtGrade.Text) > 100 Then
            MsgBox "Invalid Grade Value!" & vbNewLine & "It must be numeric and range 60-100", vbCritical
            HLTxt txtGrade
            Exit Function
        End If
    Else
        MsgBox "Invalid Grade Value!" & vbNewLine & "It must be numeric and range 60-100", vbCritical
        HLTxt txtGrade
        Exit Function
    End If
    
    'set new values
    curGrade.GradeValue = Val(txtGrade.Text)
    
    If EditGrade(curGrade) = success Then
        MsgBox "New Grade Value saved.", vbInformation
    Else
        MsgBox "Unable to save new Grade Value.", vbExclamation
        Exit Function
    End If
    
    'return success
    SaveRecord = True
    Unload Me
End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    RecordSaved = SaveRecord
End Sub

Private Sub Form_Activate()
    HLTxt txtGrade
End Sub

