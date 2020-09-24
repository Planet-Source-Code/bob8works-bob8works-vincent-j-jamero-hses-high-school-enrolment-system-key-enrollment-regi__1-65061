VERSION 5.00
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmAddSchoolYear 
   BackColor       =   &H00D8E9EC&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "School Year"
   ClientHeight    =   2265
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4290
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAddSchoolYear.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   151
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   286
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtSchoolYear 
      BackColor       =   &H00C0FFFF&
      Height          =   315
      Left            =   1710
      Locked          =   -1  'True
      MaxLength       =   9
      TabIndex        =   5
      Top             =   810
      Width           =   2205
   End
   Begin VB.TextBox txtFrom 
      Height          =   315
      Left            =   1710
      MaxLength       =   4
      TabIndex        =   4
      Top             =   1230
      Width           =   1095
   End
   Begin VB.TextBox txtTo 
      BackColor       =   &H00D8E9EC&
      Height          =   315
      Left            =   2820
      MaxLength       =   4
      TabIndex        =   3
      Top             =   1230
      Width           =   1095
   End
   Begin HSES.b8Line b8Line1 
      Height          =   60
      Left            =   -30
      TabIndex        =   0
      Top             =   510
      Width           =   6075
      _ExtentX        =   10716
      _ExtentY        =   106
   End
   Begin HSES.b8Line b8Line2 
      Height          =   60
      Left            =   0
      TabIndex        =   1
      Top             =   1770
      Width           =   5385
      _ExtentX        =   9499
      _ExtentY        =   106
   End
   Begin lvButton.lvButtons_H cmdSave 
      Default         =   -1  'True
      Height          =   360
      Left            =   2790
      TabIndex        =   8
      Top             =   1860
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
      Left            =   1260
      TabIndex        =   9
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
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "School Year Title"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   180
      TabIndex        =   7
      Top             =   840
      Width           =   1185
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "From"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   210
      TabIndex        =   6
      Top             =   1290
      Width           =   360
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Add New School Year"
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
      TabIndex        =   2
      Top             =   150
      Width           =   3075
   End
   Begin VB.Image Image2 
      Height          =   525
      Left            =   30
      Picture         =   "frmAddSchoolYear.frx":08CA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5295
   End
End
Attribute VB_Name = "frmAddSchoolYear"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RecordSaved As Boolean

Public Function ShowForm(Optional newSchoolYearTitleFrom As String = "") As Boolean
    
    '-------------------------------------------------------
    'check user access
    '-------------------------------------------------------
    If UserAllowedTo(CurrentUser.UserName, sCanAddSchoolYear) = False Then
        MsgBox "Unable to continue adding School Year entry." & vbNewLine & vbNewLine & _
                "You are not permitted to do this. Please contact your administrator for more information.", vbExclamation
        Exit Function
    End If
    '-------------------------------------------------------
    
    If newSchoolYearTitleFrom <> "" Then
        txtFrom = newSchoolYearTitleFrom
    End If
    
    'show form
    Me.Show vbModal
    
    ShowForm = RecordSaved
End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()

    'check if filled
    If Len(txtSchoolYear.Text) < 1 Then
        'temp
        MsgBox "Fill 'From Year' Text Field First", vbInformation
        Exit Sub
    End If
    
    'save
    Dim newSchoolYear As tSchoolYear
    
    'set object
    newSchoolYear.SchoolYearTitle = txtSchoolYear.Text
    
    Select Case AddSchoolYear(newSchoolYear)
        Case TranDBResult.Success
        
            'ADD success
            '------------------------------------------------------
                        
            'temp
            MsgBox "School Year created.", vbInformation
            
            'return true
            RecordSaved = True
            
            'close this form
            
            Unload Me
        
        Case TranDBResult.DuplicateTitle
            MsgBox "The Entry you have entered is already existed. Enter another entry.", vbExclamation
            HLTxt txtFrom

            
        Case Else
            'temp
            MsgBox "Error: Creating School Year", vbCritical
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

Private Sub txtSchoolYear_GotFocus()
    txtFrom.SetFocus
End Sub




Private Sub txtTo_GotFocus()
    txtFrom.SetFocus
End Sub

