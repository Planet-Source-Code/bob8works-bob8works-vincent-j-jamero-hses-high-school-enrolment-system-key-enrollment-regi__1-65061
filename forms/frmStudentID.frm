VERSION 5.00
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmStudentID 
   BackColor       =   &H00D8E9EC&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "  Create New Student ID"
   ClientHeight    =   1560
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3150
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
   ScaleHeight     =   104
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   210
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtStudentID 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   390
      MaxLength       =   4
      TabIndex        =   1
      Text            =   "2006"
      Top             =   510
      Width           =   795
   End
   Begin VB.TextBox txtStudentID 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   1230
      MaxLength       =   7
      TabIndex        =   0
      Text            =   "0000001"
      Top             =   510
      Width           =   1485
   End
   Begin lvButton.lvButtons_H cmdSave 
      Default         =   -1  'True
      Height          =   360
      Left            =   1620
      TabIndex        =   2
      Top             =   1170
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
      Left            =   90
      TabIndex        =   3
      Top             =   1170
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
      Left            =   0
      TabIndex        =   4
      Top             =   1080
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   106
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Student ID"
      Height          =   195
      Left            =   390
      TabIndex        =   5
      Top             =   270
      Width           =   780
   End
End
Attribute VB_Name = "frmStudentID"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim tmpNewStudentID As String


Public Function CreateNewStudentID() As String
    
    txtStudentID(0) = Left(Year(Now), 4)
    
    
    Me.Show vbModal
    
    'return
    CreateNewStudentID = tmpNewStudentID
End Function



Private Sub cmdCancel_Click()
    tmpNewStudentID = ""
    Unload Me
End Sub

Private Sub cmdSave_Click()
    Dim newID As String
    
    If Len(txtStudentID(1)) <> 7 Then
        MsgBox "Invalid Student ID. The right side must be 7 numeric characters", vbExclamation
        HLTxt txtStudentID(1)
        Exit Sub
    End If
    
    If Len(txtStudentID(0)) <> 4 Then
        MsgBox "Invalid Student ID. The left side must be 4 numeric characters", vbExclamation
        HLTxt txtStudentID(0)
        Exit Sub
    End If
    
    newID = txtStudentID(0) & "-" & txtStudentID(1)
    
    If StudentExistByID(newID) = Success Then
        MsgBox "The Student ID that you have entered is already existed." & vbNewLine & "Please try another value.", vbExclamation
        Exit Sub
    End If
    
    
    'set new id
    tmpNewStudentID = newID
    'close this form and return
    Unload Me
End Sub

Private Sub Image1_Click()
End Sub

Private Sub Label4_Click()

End Sub

Private Sub txtStudentID_LostFocus(Index As Integer)
    If Index = 1 Then
        
        If IsNumeric(txtStudentID(Index)) Then
            txtStudentID(Index).Text = Trim(txtStudentID(Index).Text)
            txtStudentID(Index) = Left("0000000", 7 - Len(txtStudentID(Index).Text)) & txtStudentID(Index)
        
        Else
            MsgBox "Invalid Student ID. It must be 7 numeric characters", vbExclamation
            HLTxt txtStudentID(Index)
        End If
    End If
End Sub

Private Sub txtStudentID_KeyPress(Index As Integer, KeyAscii As Integer)
    If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8 Or KeyAscii = 45) Then KeyAscii = 0
End Sub

Private Sub Image5_Click()

End Sub
