VERSION 5.00
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmSchoolYearLock 
   BackColor       =   &H00D8E9EC&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "School Year"
   ClientHeight    =   2280
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3630
   Icon            =   "frmSchoolYearLock.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   152
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   242
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkLocked 
      BackColor       =   &H00D8E9EC&
      Caption         =   "Locked"
      Enabled         =   0   'False
      Height          =   315
      Left            =   1170
      TabIndex        =   5
      Top             =   1350
      Width           =   1395
   End
   Begin VB.CommandButton cmdGetSchoolYearTitle 
      BackColor       =   &H00D8E9EC&
      Height          =   330
      Left            =   3060
      Picture         =   "frmSchoolYearLock.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   810
      Width           =   345
   End
   Begin HSES.b8Line b8Line1 
      Height          =   60
      Left            =   60
      TabIndex        =   0
      Top             =   480
      Width           =   6045
      _ExtentX        =   10663
      _ExtentY        =   106
   End
   Begin VB.TextBox txtSchoolYear 
      Height          =   375
      Left            =   1170
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   4
      Top             =   780
      Width           =   2265
   End
   Begin HSES.b8Line b8Line2 
      Height          =   60
      Left            =   0
      TabIndex        =   6
      Top             =   1800
      Width           =   5385
      _ExtentX        =   9499
      _ExtentY        =   106
   End
   Begin lvButton.lvButtons_H cmdSave 
      Default         =   -1  'True
      Height          =   360
      Left            =   2160
      TabIndex        =   7
      Top             =   1890
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
      Left            =   630
      TabIndex        =   8
      Top             =   1890
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
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "School Year"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   180
      TabIndex        =   2
      Top             =   810
      Width           =   870
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Lock / Unlock"
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
      TabIndex        =   1
      Top             =   120
      Width           =   1980
   End
   Begin VB.Image Image2 
      Height          =   525
      Left            =   45
      Picture         =   "frmSchoolYearLock.frx":0E54
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5295
   End
End
Attribute VB_Name = "frmSchoolYearLock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RecordSaved As Boolean



Public Function ShowForm(Optional sSchoolYear As String = "") As Boolean
    
    'default
    RecordSaved = False
    
    '-------------------------------------------------------
    'check user access
    '-------------------------------------------------------
    If UserAllowedTo(CurrentUser.UserName, sCanLockUnlockSchoolYear) = False Then
        MsgBox "Unable to continue Lock/Unlock School Year entry." & vbNewLine & vbNewLine & _
                "You are not permitted to do this. Please contact your administrator for more information.", vbExclamation
        Exit Function
    End If
    '-------------------------------------------------------


    
    txtSchoolYear.Text = sSchoolYear
    
    Me.Show vbModal

    'return
    ShowForm = RecordSaved
End Function


Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdGetSchoolYearTitle_Click()
    Dim sSchoolYear As String
    
    sSchoolYear = PickSchoolYear.GetItem(txtSchoolYear)
    
    If sSchoolYear <> "" Then
    
        txtSchoolYear.Text = sSchoolYear

    End If
End Sub

Private Sub cmdSave_Click()
    Dim vSchoolYear As tSchoolYear
    
    
    If Len(txtSchoolYear.Text) < 1 Then
        MsgBox "Please enter School Year.", vbExclamation
        cmdGetSchoolYearTitle.SetFocus
        Exit Sub
    End If
    
    If GetSchoolYearByTitle(txtSchoolYear.Text, vSchoolYear) <> Success Then
        MsgBox "Please enter Valid School Year.", vbExclamation
        cmdGetSchoolYearTitle.SetFocus
        Exit Sub
    End If
    
    
    ChangeLockSchoolYear
    
End Sub

Private Function ChangeLockSchoolYear()

    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    sSQL = "SELECT * FROM tblSchoolYear WHERE tblSchoolYear.SchoolYearTitle='" & txtSchoolYear.Text & "'"
    
    If ConnectRS(HSESDB, vRS, sSQL) = False Then
        'fatal error
        'temp
        MsgBox "error"
        GoTo ReleaseAndExit
    End If
    
    If AnyRecordExisted(vRS) = False Then
        MsgBox "School Year Not Found.", vbExclamation
        GoTo ReleaseAndExit
    End If
    
    vRS.MoveFirst
    If chkLocked.Value = Checked Then
        vRS.Fields("Locked").Value = True
    Else
        vRS.Fields("Locked").Value = False
    End If
    
    vRS.Update
    
    
    MsgBox "School Year Entry successfully updated.", vbInformation
    RecordSaved = True
    Unload Me
    
ReleaseAndExit:
    Set vRS = Nothing
End Function
Private Sub txtSchoolYear_Change()
    
    Dim vSchoolYear As tSchoolYear
    
    
    If Len(txtSchoolYear.Text) < 1 Then
        Exit Sub
    End If
    
    If GetSchoolYearByTitle(txtSchoolYear.Text, vSchoolYear) <> Success Then
        chkLocked.Enabled = False
        Exit Sub
    End If
    
    
    chkLocked.Enabled = True
    
    If vSchoolYear.Locked = True Then
        chkLocked.Value = vbChecked
    Else
        chkLocked.Value = vbUnchecked
    End If
    
End Sub
