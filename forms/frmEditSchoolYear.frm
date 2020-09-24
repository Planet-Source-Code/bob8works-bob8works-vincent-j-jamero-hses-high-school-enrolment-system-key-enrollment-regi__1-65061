VERSION 5.00
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmEditSchoolYear 
   BackColor       =   &H00D8E9EC&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Edit School Year"
   ClientHeight    =   3315
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4365
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   221
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   291
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtTo 
      BackColor       =   &H00D8E9EC&
      Height          =   315
      Left            =   2955
      MaxLength       =   4
      TabIndex        =   2
      Top             =   1530
      Width           =   1005
   End
   Begin VB.TextBox txtFrom 
      Height          =   315
      Left            =   1755
      MaxLength       =   4
      TabIndex        =   1
      Top             =   1530
      Width           =   1005
   End
   Begin VB.TextBox txtSchoolYear 
      BackColor       =   &H00C0FFFF&
      Height          =   315
      Left            =   1755
      Locked          =   -1  'True
      MaxLength       =   9
      TabIndex        =   0
      Top             =   990
      Width           =   2205
   End
   Begin lvButton.lvButtons_H cmdSave 
      Default         =   -1  'True
      Height          =   360
      Left            =   2625
      TabIndex        =   3
      Top             =   2490
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
      Left            =   975
      TabIndex        =   4
      Top             =   2490
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
End
Attribute VB_Name = "frmEditSchoolYear"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CurrentSchoolYear As tSchoolYear
Dim RecordEdited As Boolean
Public Function ShowEdit(sSchoolYearTitle As String) As Boolean
    If SchoolYearExistByTitle(sSchoolYearTitle) <> Success Then
        MsgBox "School Year entitled '" & sSchoolYearTitle & "' not existed in record.", vbExclamation
        Exit Function
    End If
    
    CurrentSchoolYear.SchoolYearTitle = sSchoolYearTitle
    lblOldSchoolYearTitle = sSchoolYearTitle
    
    txtFrom.Text = Left(CurrentSchoolYear.SchoolYearTitle, 4)
    
    'Show form
    Me.Show vbModal
    
    'this next line will be executed after the form was unloaded
    ShowEdit = RecordEdited
    
End Function


Private Function SaveEdit()
    'check if filled
    If Len(txtSchoolYear.Text) < 1 Then
        'temp
        MsgBox "Fill 'From Year' Text Field First", vbInformation
        Exit Function
    End If
    
    'check for duplicate
    If SchoolYearExistByTitle(txtSchoolYear.Text) = Success Then
        'temp
        MsgBox "This School Year is already existed.", vbExclamation
        Exit Function
    End If
    
    'save edit
    Dim newSchoolYear As tSchoolYear
    
    'set object
    newSchoolYear.SchoolYearTitle = txtSchoolYear.Text
    
    If EditSchoolYear(CurrentSchoolYear.SchoolYearTitle, newSchoolYear) = Success Then
        
        
        'EDIT success
        '------------------------------------------------------------
        MsgBox "School Year successfully changed", vbInformation
        'close this form
        Unload Me
        'set flag
        RecordEdited = True
        
        '------------------------------------------------------------
    Else
        'temp
        MsgBox "Error: Editing School Year", vbCritical
    End If
End Function


Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdUpdate_Click()
    
    Call SaveEdit
End Sub

Private Sub Form_Activate()
    HLTxt txtFrom
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyEscape
            Unload Me
        Case vbKeyReturn
            Call SaveEdit
    End Select
End Sub

Private Sub txtFrom_Change()
    If Len(txtFrom) = 4 And Val(txtFrom) > 1000 Then
            'auto fill
            txtTo.Text = Val(txtFrom) + 1
            txtSchoolYear.Text = txtFrom.Text & "-" & txtTo.Text
    Else
        txtTo.Text = ""
        txtSchoolYear.Text = ""
    End If
End Sub



Private Sub txtFrom_KeyPress(KeyAscii As Integer)
    If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8 Or KeyAscii = 45) Then KeyAscii = 0
End Sub
