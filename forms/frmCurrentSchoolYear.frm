VERSION 5.00
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmCurrentSchoolYear 
   BackColor       =   &H00D8E9EC&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "School Year"
   ClientHeight    =   1890
   ClientLeft      =   3225
   ClientTop       =   4185
   ClientWidth     =   4335
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCurrentSchoolYear.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   126
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   289
   Begin VB.TextBox txtSchoolYear 
      Enabled         =   0   'False
      Height          =   330
      Left            =   1575
      TabIndex        =   1
      Top             =   870
      Width           =   2085
   End
   Begin VB.CommandButton cmdGetSchoolYear 
      BackColor       =   &H00D8E9EC&
      Height          =   315
      Left            =   3675
      Picture         =   "frmCurrentSchoolYear.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   870
      Width           =   345
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
      Left            =   -90
      TabIndex        =   4
      Top             =   1410
      Width           =   5385
      _ExtentX        =   9499
      _ExtentY        =   106
   End
   Begin lvButton.lvButtons_H cmdSelect 
      Default         =   -1  'True
      Height          =   360
      Left            =   2850
      TabIndex        =   5
      Top             =   1500
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   635
      Caption         =   "&Select"
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
      Left            =   1320
      TabIndex        =   6
      Top             =   1500
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
      Caption         =   "Set Active School year"
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
      TabIndex        =   7
      Top             =   150
      Width           =   3195
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "School Year Title"
      Height          =   195
      Left            =   315
      TabIndex        =   2
      Top             =   915
      Width           =   1185
   End
   Begin VB.Image Image2 
      Height          =   525
      Left            =   45
      Picture         =   "frmCurrentSchoolYear.frx":0B14
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5295
   End
End
Attribute VB_Name = "frmCurrentSchoolYear"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Sub setSchoolYear()
    
    If SchoolYearRecordExisted <> Success Then
        MsgBox "There are no records yet in School Year.", vbInformation
        Unload Me
        Exit Sub
    End If

    If Len(CurrentSchoolYear.SchoolYearTitle) > 0 Then
        txtSchoolYear.Text = CurrentSchoolYear.SchoolYearTitle
    End If

    Me.Show vbModal
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdGetSchoolYear_Click()
    Dim sSchoolYearTitle As String
    
    sSchoolYearTitle = PickSchoolYear.GetItem(txtSchoolYear)
    
    If sSchoolYearTitle <> "" Then
        txtSchoolYear.Text = sSchoolYearTitle
    End If
End Sub


Private Sub cmdSelect_Click()
    If SchoolYearExistByTitle(txtSchoolYear.Text) = Success Then
        SaveActiveSchoolYear txtSchoolYear.Text
        CurrentSchoolYear.SchoolYearTitle = txtSchoolYear.Text
        Unload Me
    Else
        MsgBox "The selected School Year does not exist in record!" & vbNewLine & _
        "Please enter valid School Year.", vbExclamation
    End If
End Sub

Private Sub txtSchoolYear_Change()
    If Len(txtSchoolYear) < 1 Then
        cmdSelect.Enabled = False
    Else
        cmdSelect.Enabled = True
    End If
End Sub
