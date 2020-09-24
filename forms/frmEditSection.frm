VERSION 5.00
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmEditSection 
   BackColor       =   &H00D8E9EC&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Section"
   ClientHeight    =   4275
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5070
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmEditSection.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   285
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   338
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00D8E9EC&
      Caption         =   "History"
      Height          =   1125
      Left            =   240
      TabIndex        =   13
      Top             =   2490
      Width           =   4515
      Begin VB.Label lblHistory 
         BackColor       =   &H00D8E9EC&
         BackStyle       =   0  'Transparent
         Caption         =   "--"
         Height          =   825
         Left            =   195
         TabIndex        =   14
         Top             =   225
         Width           =   4125
      End
   End
   Begin VB.TextBox txtSectionTitle 
      Height          =   345
      Left            =   1530
      MaxLength       =   50
      TabIndex        =   3
      Top             =   1230
      Width           =   3225
   End
   Begin VB.TextBox txtSectionID 
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
      Height          =   345
      Left            =   1530
      Locked          =   -1  'True
      MaxLength       =   10
      TabIndex        =   2
      Top             =   780
      Width           =   3225
   End
   Begin VB.CommandButton cmdGetYearLevelTitle 
      BackColor       =   &H00D8E9EC&
      Enabled         =   0   'False
      Height          =   285
      Left            =   4380
      Picture         =   "frmEditSection.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2100
      Width           =   345
   End
   Begin VB.CommandButton cmdGetDepartmentTitle 
      BackColor       =   &H00D8E9EC&
      Enabled         =   0   'False
      Height          =   300
      Left            =   4380
      Picture         =   "frmEditSection.frx":0B14
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1680
      Width           =   345
   End
   Begin HSES.b8Line b8Line1 
      Height          =   60
      Left            =   0
      TabIndex        =   6
      Top             =   510
      Width           =   5535
      _ExtentX        =   9499
      _ExtentY        =   106
   End
   Begin HSES.b8Line b8Line2 
      Height          =   60
      Left            =   -30
      TabIndex        =   7
      Top             =   3795
      Width           =   5820
      _ExtentX        =   9499
      _ExtentY        =   106
   End
   Begin lvButton.lvButtons_H cmdSave 
      Default         =   -1  'True
      Height          =   360
      Left            =   3600
      TabIndex        =   15
      Top             =   3885
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
      Left            =   2070
      TabIndex        =   16
      Top             =   3885
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
   Begin VB.TextBox txtYearLevelTitle 
      Enabled         =   0   'False
      Height          =   345
      Left            =   1530
      MaxLength       =   20
      TabIndex        =   5
      Top             =   2070
      Width           =   3225
   End
   Begin VB.TextBox txtDepartmentTitle 
      Enabled         =   0   'False
      Height          =   345
      Left            =   1530
      MaxLength       =   50
      TabIndex        =   4
      Top             =   1650
      Width           =   3225
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Edit Section"
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
      Left            =   60
      TabIndex        =   12
      Top             =   150
      Width           =   1725
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Section ID"
      Height          =   195
      Left            =   270
      TabIndex        =   11
      Top             =   810
      Width           =   735
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Year Level"
      Height          =   195
      Left            =   300
      TabIndex        =   10
      Top             =   2070
      Width           =   750
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Departent"
      Height          =   195
      Left            =   300
      TabIndex        =   9
      Top             =   1650
      Width           =   735
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Section Title"
      Height          =   195
      Left            =   270
      TabIndex        =   8
      Top             =   1275
      Width           =   870
   End
   Begin VB.Image Image2 
      Height          =   525
      Left            =   0
      Picture         =   "frmEditSection.frx":109E
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5550
   End
End
Attribute VB_Name = "frmEditSection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Dim RecordEdited As Boolean

Dim curSection As tSection


Public Function ShowForm(sSectionID As String) As Boolean
    
    Dim vDepartment As tDepartment
    
    '-------------------------------------------------------
    'check user access
    '-------------------------------------------------------
    If UserAllowedTo(CurrentUser.UserName, sCanEditSection) = False Then
        MsgBox "Unable to continue Editing Section entry." & vbNewLine & vbNewLine & _
                "You are not permitted to do this. Please contact your administrator for more information.", vbExclamation
        Exit Function
    End If
    '-------------------------------------------------------


    
    If GetSectionByID(sSectionID, curSection) <> Success Then
    
        MsgBox "Unable to continue Editing Section entry." & vbNewLine & "Selected Sectiohn cannot be found in record.", vbCritical
        Unload Me
        Exit Function
    End If
    
    'set form's text fields
    txtSectionID.Text = curSection.SectionID
    txtSectionTitle.Text = curSection.SectionTitle
    
    If GetDepartmentByID(curSection.DepartmentID, vDepartment) <> Success Then
        'hide msg to user
        'record error
        CatchError "frmEditSection", "ShowForm", "Error: Department By ID not found. " & vDepartment.DepartmentID
    End If
    txtDepartmentTitle.Text = vDepartment.DepartmentTitle
    txtYearLevelTitle.Text = YLIDtoTitle(curSection.YearLevelID)
    lblHistory.Caption = "Created Date: " & curSection.CreationDate & vbNewLine & _
                        "Created By: " & curSection.CreatedBy & vbNewLine & _
                        "Last Modified Date: " & IIf(curSection.ModifiedDate = CDate(0), "", curSection.ModifiedDate) & vbNewLine & _
                        "Last Modified By: " & curSection.ModifiedBy
    
    'show form
    Me.Show vbModal
    
    'return
    ShowForm = RecordEdited
End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    
    If Len(txtSectionTitle.Text) < 1 Then
        MsgBox "Section Title must not be empty." & vbNewLine & _
                "Please enter some value.", vbExclamation
        Exit Sub
    End If
    
    'set section's fields
    curSection.SectionTitle = Trim(txtSectionTitle.Text)
    curSection.ModifiedBy = CurrentUser.UserName
    curSection.ModifiedDate = Now
    
    'save
    Select Case EditSection(curSection)
        Case Success
            MsgBox "Section entry successfully edited.", vbInformation
            
            'set flag
            RecordEdited = True
            'close this form
            Unload Me
            
        Case TranDBResult.DuplicateTitle
            MsgBox "Section Title already exist. Please try another value.", vbExclamation
            
            HLTxt txtSectionTitle
            
        Case Else
            MsgBox "Unable to update changes", vbCritical
            
            CatchError "frmEditSection", "cmdSave_Click", "EditSection Return unknown error."
    End Select
End Sub

Private Sub Form_Activate()
    HLTxt txtSectionTitle
End Sub

