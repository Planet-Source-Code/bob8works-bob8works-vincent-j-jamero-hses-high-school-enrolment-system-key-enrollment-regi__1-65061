VERSION 5.00
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmAddFee 
   BackColor       =   &H00D8E9EC&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Fee"
   ClientHeight    =   4560
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5235
   Icon            =   "frmAddFee.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   304
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   349
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdGetSchoolYear 
      BackColor       =   &H00D8E9EC&
      Height          =   255
      Left            =   4140
      Picture         =   "frmAddFee.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   2070
      Width           =   345
   End
   Begin VB.TextBox txtSchoolYear 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1380
      MaxLength       =   50
      TabIndex        =   18
      Text            =   "All School Year"
      Top             =   2040
      Width           =   3135
   End
   Begin VB.CommandButton cmdGetDepartmentTitle 
      BackColor       =   &H00D8E9EC&
      Height          =   270
      Left            =   4140
      Picture         =   "frmAddFee.frx":0596
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   2460
      Width           =   345
   End
   Begin VB.ComboBox cmbYearLevel 
      Height          =   315
      Left            =   1380
      TabIndex        =   16
      Top             =   2820
      Width           =   3135
   End
   Begin VB.TextBox txtDescription 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   1380
      MaxLength       =   100
      MultiLine       =   -1  'True
      TabIndex        =   13
      Top             =   3210
      Width           =   3135
   End
   Begin VB.TextBox txtAmount 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1380
      MaxLength       =   10
      TabIndex        =   10
      Text            =   "1.00"
      Top             =   1650
      Width           =   3135
   End
   Begin VB.TextBox txtDepartmentTitle 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1380
      MaxLength       =   50
      TabIndex        =   9
      Top             =   2430
      Width           =   3135
   End
   Begin VB.TextBox txtFeeID 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1380
      Locked          =   -1  'True
      MaxLength       =   10
      TabIndex        =   2
      Top             =   840
      Width           =   3135
   End
   Begin VB.TextBox txtTitle 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1380
      MaxLength       =   50
      TabIndex        =   1
      Top             =   1260
      Width           =   3135
   End
   Begin HSES.b8Line b8Line1 
      Height          =   60
      Left            =   0
      TabIndex        =   0
      Top             =   510
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   106
   End
   Begin lvButton.lvButtons_H cmdSave 
      Default         =   -1  'True
      Height          =   360
      Left            =   3765
      TabIndex        =   3
      Top             =   4155
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
      cGradient       =   14215660
      Gradient        =   3
      Mode            =   0
      Value           =   0   'False
      cBack           =   16777215
   End
   Begin lvButton.lvButtons_H cmdCancel 
      Cancel          =   -1  'True
      Height          =   360
      Left            =   2115
      TabIndex        =   4
      Top             =   4155
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
      cGradient       =   14215660
      Gradient        =   3
      Mode            =   0
      Value           =   0   'False
      cBack           =   16777215
   End
   Begin HSES.b8Line b8Line2 
      Height          =   60
      Left            =   0
      TabIndex        =   5
      Top             =   4065
      Width           =   5385
      _ExtentX        =   9499
      _ExtentY        =   106
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "School Year"
      Height          =   195
      Left            =   300
      TabIndex        =   19
      Top             =   2100
      Width           =   870
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      Height          =   195
      Left            =   300
      TabIndex        =   15
      Top             =   3270
      Width           =   795
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Year Level"
      Height          =   195
      Left            =   300
      TabIndex        =   14
      Top             =   2850
      Width           =   765
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Department"
      Height          =   195
      Left            =   300
      TabIndex        =   12
      Top             =   2490
      Width           =   825
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "*Amount"
      Height          =   195
      Left            =   300
      TabIndex        =   11
      Top             =   1680
      Width           =   600
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "*Title"
      Height          =   195
      Left            =   300
      TabIndex        =   8
      Top             =   1320
      Width           =   360
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ID"
      Height          =   195
      Left            =   300
      TabIndex        =   7
      Top             =   900
      Width           =   165
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "New Fee"
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
      TabIndex        =   6
      Top             =   150
      Width           =   1230
   End
   Begin VB.Image Image2 
      Height          =   525
      Left            =   0
      Picture         =   "frmAddFee.frx":0B20
      Stretch         =   -1  'True
      Top             =   0
      Width           =   6495
   End
End
Attribute VB_Name = "frmAddFee"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const sAllDepartment = "All Department"
            

Dim RecordSaved As Boolean

Public Function ShowForm() As Boolean
    
    
    '-------------------------------------------------------
    'check user access
    '-------------------------------------------------------
    If UserAllowedTo(CurrentUser.UserName, sCanAddFee) = False Then
        MsgBox "Unable to continue adding Fee entry." & vbNewLine & vbNewLine & _
                "You are not permitted to do this. Please contact your administrator for more information.", vbExclamation
        Exit Function
    End If
    '-------------------------------------------------------

    
    
    Me.Show vbModal

    'return
    ShowForm = RecordSaved
End Function

Private Sub cmbYearLevel_LostFocus()
    If cmbYearLevel.ListIndex < 0 Then cmbYearLevel.ListIndex = 0
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdGetItem_Click()
    Dim sDepartmentTitle As String

    sDepartmentTitle = PickDepartment.GetItem
    If sDepartmentTitle <> "" Then
        txtDepartmentTitle.Text = sDepartmentTitle
    End If
End Sub

Private Sub cmdGetSchoolYear_Click()
    Dim sSchoolYearTitle As String
    
    sSchoolYearTitle = PickSchoolYear.GetItem(txtSchoolYear)
    
    If sSchoolYearTitle <> "" Then
        txtSchoolYear.Text = sSchoolYearTitle
    End If

End Sub

Private Sub cmdSave_Click()
    'validate data
    If ValidateFormData = False Then
        Exit Sub
    End If
    
    SaveData
    
End Sub


Private Sub SaveData()
    
    Dim vDepartment As tDepartment
    Dim sDepartmentID As String
    Dim sSchoolYear As String
    
    'set cursor
    Me.MousePointer = vbHourglass
    
    If GetDepartmentByTitle(txtDepartmentTitle.Text, vDepartment) = Success Then
        sDepartmentID = vDepartment.DepartmentID
    Else
        sDepartmentID = ""
    End If
    
    If SchoolYearExistByTitle(txtSchoolYear.Text) = Success Then
        sSchoolYear = txtSchoolYear.Text
    Else
        sSchoolYear = ""
    End If
    
    Select Case AddFee(CLng(txtFeeID.Text), _
                txtTitle.Text, _
                txtDescription.Text, _
                CDbl(txtAmount.Text), _
                sSchoolYear, _
                sDepartmentID, _
                cmbYearLevel.ListIndex, _
                Now, _
                CurrentUser.UserName)
                
        Case TranDBResult.Success
            MsgBox "New Fee entry successfully added.", vbInformation
            'set flag
            RecordSaved = True
            'close form
            Unload Me
        Case Else
            'fatal or unknown error
            MsgBox "Unable to saved Fee entry. Please check all fields.", vbExclamation
            CatchError "frmAddFee", "SaveData", "Unknown AddFee Result"
            
    End Select
    
    'restore cursor
    Me.MousePointer = vbDefault
End Sub
Private Function ValidateFormData() As Boolean
    
    'default
    ValidateFormData = False
    
    '---------------------------------
    If Len(txtTitle.Text) < 1 Then
        MsgBox "Title must not be empty. Please enter some value.", vbExclamation
    
        HLTxt txtTitle
        Exit Function
    End If
    
    If IsNumeric(txtAmount.Text) = True Then
        If Val(txtAmount.Text) > 0 Then
            txtAmount.Text = FormatNumber(txtAmount.Text, 2)
        Else
            MsgBox "Amount must be greater than 0.00.", vbExclamation
            HLTxt txtAmount
            Exit Function
        End If
    Else
        MsgBox "Amount must be in numeric value (ex: 100.00).", vbExclamation
        HLTxt txtAmount
        Exit Function
    End If
    
    'ckeck department
    If txtDepartmentTitle.Text <> sAllDepartment Then
        If DepartmentExistByTitle(txtDepartmentTitle.Text) <> Success Then
            txtDepartmentTitle.Text = sAllDepartment
        End If
    End If
    
    
    '---------------------------------
    
    'return success
    ValidateFormData = True
End Function

Private Sub Form_Activate()
    'generate id
    Dim lNewFeeID As Long
    
    lNewFeeID = GetNewFeeID

    If lNewFeeID > 0 Then
        txtFeeID.Text = String$(10 - Len(Trim(lNewFeeID)), "0") & lNewFeeID
    Else
        'fatal error
    End If
    
    
End Sub

Private Sub Form_Load()
    cmbYearLevel.Clear
    cmbYearLevel.AddItem "All Year Level"
    cmbYearLevel.AddItem "I"
    cmbYearLevel.AddItem "II"
    cmbYearLevel.AddItem "III"
    cmbYearLevel.AddItem "IV"
    cmbYearLevel.ListIndex = 0
    
    txtDepartmentTitle.Text = sAllDepartment
End Sub

Private Sub txtDepartmentTitle_LostFocus()
    If DepartmentExistByTitle(txtDepartmentTitle.Text) <> Success Then
        txtDepartmentTitle.Text = sAllDepartment
    End If
End Sub

Private Sub txtAmount_LostFocus()
    If IsNumeric(txtAmount.Text) = True Then
        If Val(txtAmount.Text) > 0 Then
            txtAmount.Text = FormatNumber(txtAmount.Text, 2)
        Else
            MsgBox "Amount must be greater than 0.00.", vbExclamation
            HLTxt txtAmount
        End If
    Else
        MsgBox "Amount must be in numeric value (ex: 100.00).", vbExclamation
        HLTxt txtAmount

    End If
End Sub

Private Sub txtSchoolYear_LostFocus()
    If SchoolYearExistByTitle(txtSchoolYear.Text) = Failed Then
        txtSchoolYear.Text = "All School Year"
    End If
End Sub

Private Sub txtTitle_LostFocus()
    'If Len(txtTitle.Text) < 1 Then
    '    MsgBox "Title must not be empty. Please enter some value.", vbExclamation
    
    '    HLTxt txtTitle
    'End If
    
End Sub
