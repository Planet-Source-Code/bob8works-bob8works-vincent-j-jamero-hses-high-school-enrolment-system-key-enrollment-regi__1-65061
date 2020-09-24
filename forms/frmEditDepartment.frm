VERSION 5.00
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmEditDepartment 
   BackColor       =   &H00D8E9EC&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Department"
   ClientHeight    =   2235
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5100
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmEditDepartment.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   149
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   340
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtDepartmentID 
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
      Left            =   1650
      Locked          =   -1  'True
      MaxLength       =   4
      TabIndex        =   2
      Top             =   840
      Width           =   3135
   End
   Begin VB.TextBox txtDepartmentTitle 
      Height          =   315
      Left            =   1650
      MaxLength       =   20
      TabIndex        =   1
      Top             =   1230
      Width           =   3135
   End
   Begin HSES.b8Line b8Line1 
      Height          =   60
      Left            =   60
      TabIndex        =   0
      Top             =   510
      Width           =   5985
      _ExtentX        =   10557
      _ExtentY        =   106
   End
   Begin lvButton.lvButtons_H cmdUpdate 
      Default         =   -1  'True
      Height          =   360
      Left            =   3630
      TabIndex        =   6
      Top             =   1860
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   635
      Caption         =   "&Update"
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
      Left            =   2100
      TabIndex        =   7
      Top             =   1860
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
   Begin HSES.b8Line b8Line2 
      Height          =   60
      Left            =   -30
      TabIndex        =   8
      Top             =   1770
      Width           =   5685
      _ExtentX        =   10028
      _ExtentY        =   106
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Department Title"
      Height          =   195
      Left            =   240
      TabIndex        =   5
      Top             =   1320
      Width           =   1170
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Department ID"
      Height          =   195
      Left            =   240
      TabIndex        =   4
      Top             =   840
      Width           =   1035
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Edit Department"
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
      TabIndex        =   3
      Top             =   180
      Width           =   2355
   End
   Begin VB.Image Image2 
      Height          =   525
      Left            =   30
      Picture         =   "frmEditDepartment.frx":058A
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5295
   End
End
Attribute VB_Name = "frmEditDepartment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim currentDepartment As tDepartment

Dim RecordEdited As Boolean



Public Function ShowEdit(sDepartmentID As String) As Boolean
    
    '-------------------------------------------------------
    'check user access
    '-------------------------------------------------------
    If UserAllowedTo(CurrentUser.UserName, sCanEditDepartment) = False Then
        MsgBox "Unable to continue editing Department entry." & vbNewLine & vbNewLine & _
                "You are not permitted to do this. Please contact your administrator for more information.", vbExclamation
        Exit Function
    End If
    '-------------------------------------------------------
    
    
    
        If GetDepartmentByID(sDepartmentID, currentDepartment) <> Success Then
            MsgBox "Unable to continue editing Department Information: Department ID not found!", vbCritical
            Exit Function
        End If


    'ready for edit
    'set data
                
    txtDepartmentID.Text = currentDepartment.DepartmentID
    txtDepartmentTitle.Text = currentDepartment.DepartmentTitle
                
    'show form
    Me.Show vbModal



    'return
    ShowEdit = RecordEdited
    
    
End Function


Private Sub cmdCancel_Click()
    'just close
    Unload Me
End Sub


Private Function SaveData()

    If Not CheckTextBox(txtDepartmentID, "Enter Department ID." & vbNewLine & " This field is required") Then
        Exit Function
    End If
    
    If Not CheckTextBox(txtDepartmentTitle, "Enter Department Title." & vbNewLine & " This field is required") Then
        Exit Function
    End If
    
    
    'save
    Dim newDepartment As tDepartment
    Dim EditResult As Integer
    
    newDepartment.DepartmentID = txtDepartmentID.Text
    newDepartment.DepartmentTitle = txtDepartmentTitle.Text

    Select Case EditDepartment(newDepartment)
        Case Success
        
            MsgBox "Department Information was successfully edited", vbInformation
            
            'set flag
            RecordEdited = True
        
            'close this form
            Unload Me
            
        Case DuplicateTitle
            MsgBox "The Department Title that you have enetered was already existed." & vbNewLine & " Enter another Duplicate Title", vbExclamation
            HLTxt txtDepartmentTitle

        Case Else
            MsgBox "UNKNOWN: Editing Department", vbCritical
    End Select

    
End Function


Private Sub cmdUpdate_Click()
    SaveData
End Sub
