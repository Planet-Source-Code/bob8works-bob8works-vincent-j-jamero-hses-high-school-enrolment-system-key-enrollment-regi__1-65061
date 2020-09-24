VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmQuickLaunch 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Welcome"
   ClientHeight    =   7185
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7260
   ControlBox      =   0   'False
   FillColor       =   &H00FAE5D3&
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmQuickLaunch.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   479
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   484
   WindowState     =   2  'Maximized
   Begin VB.Timer timerUT 
      Interval        =   1000
      Left            =   6450
      Top             =   3840
   End
   Begin VB.PictureBox bgMain 
      Appearance      =   0  'Flat
      BackColor       =   &H00CC8661&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6855
      Left            =   105
      ScaleHeight     =   457
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   538
      TabIndex        =   0
      Top             =   150
      Width           =   8070
      Begin lvButton.lvButtons_H cmdB 
         Height          =   345
         Index           =   1
         Left            =   2505
         TabIndex        =   7
         Top             =   2715
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   609
         Caption         =   "About HSES"
         CapAlign        =   2
         BackStyle       =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFore           =   0
         cFHover         =   0
         cBhover         =   16777215
         Focus           =   0   'False
         LockHover       =   3
         cGradient       =   13403745
         Gradient        =   3
         Mode            =   0
         Value           =   0   'False
         ImgSize         =   32
         cBack           =   13403745
      End
      Begin lvButton.lvButtons_H cmdB 
         Height          =   345
         Index           =   0
         Left            =   420
         TabIndex        =   6
         Top             =   2715
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   609
         Caption         =   "Today's Transactions"
         CapAlign        =   2
         BackStyle       =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFore           =   0
         cFHover         =   0
         cBhover         =   16777215
         Focus           =   0   'False
         LockHover       =   3
         cGradient       =   13403745
         Gradient        =   3
         Mode            =   0
         Value           =   0   'False
         ImgSize         =   32
         cBack           =   13403745
      End
      Begin VB.PictureBox bgB 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFEFE1&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   3765
         Index           =   1
         Left            =   420
         ScaleHeight     =   251
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   331
         TabIndex        =   8
         Top             =   3045
         Visible         =   0   'False
         Width           =   4965
         Begin VB.PictureBox bgMe 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   0
            Left            =   0
            ScaleHeight     =   0
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   130
            TabIndex        =   18
            Top             =   0
            Visible         =   0   'False
            Width           =   1950
            Begin VB.Timer timerAniIn 
               Enabled         =   0   'False
               Interval        =   10
               Left            =   1215
               Top             =   2475
            End
            Begin VB.Image me2 
               Height          =   3540
               Left            =   0
               Picture         =   "frmQuickLaunch.frx":000C
               Top             =   0
               Visible         =   0   'False
               Width           =   1950
            End
            Begin VB.Image me1 
               Height          =   3540
               Left            =   0
               Picture         =   "frmQuickLaunch.frx":62BB
               Top             =   0
               Width           =   1950
            End
         End
         Begin lvButton.lvButtons_H cmdPW 
            Height          =   345
            Left            =   150
            TabIndex        =   14
            Top             =   570
            Width           =   3255
            _ExtentX        =   5741
            _ExtentY        =   609
            Caption         =   "Personal Website: www.bob8works.cjb.net"
            CapAlign        =   2
            BackStyle       =   5
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cBhover         =   12648447
            LockHover       =   1
            cGradient       =   0
            Mode            =   0
            Value           =   0   'False
            cBack           =   16773089
         End
         Begin lvButton.lvButtons_H cmdVote 
            Height          =   345
            Left            =   150
            TabIndex        =   15
            Top             =   1800
            Width           =   2115
            _ExtentX        =   3731
            _ExtentY        =   609
            Caption         =   "Click here to Vote this code"
            CapAlign        =   2
            BackStyle       =   5
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cFore           =   16576
            cFHover         =   16576
            cBhover         =   12648447
            LockHover       =   1
            cGradient       =   0
            Mode            =   0
            Value           =   0   'False
            cBack           =   16773089
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Thank you for downloading this code and to every one at PSCODE.com"
            ForeColor       =   &H00808080&
            Height          =   435
            Left            =   180
            TabIndex        =   17
            Top             =   1200
            Width           =   3765
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Credit: LaVolpe fro LVButton"
            ForeColor       =   &H00808080&
            Height          =   255
            Left            =   180
            TabIndex        =   16
            Top             =   990
            Width           =   3255
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "bob8works@yahoo.com"
            Height          =   255
            Left            =   210
            TabIndex        =   13
            Top             =   360
            Width           =   3975
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Developed by: Vincent J. Jamero"
            Height          =   255
            Left            =   210
            TabIndex        =   12
            Top             =   150
            Width           =   3975
         End
      End
      Begin VB.PictureBox bgB 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFEFE1&
         BorderStyle     =   0  'None
         ForeColor       =   &H00FFFFFF&
         Height          =   3015
         Index           =   0
         Left            =   420
         ScaleHeight     =   201
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   331
         TabIndex        =   4
         Top             =   3060
         Visible         =   0   'False
         Width           =   4965
         Begin MSComctlLib.ImageList ilStudent 
            Left            =   1470
            Top             =   2430
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   -2147483643
            ImageWidth      =   32
            ImageHeight     =   32
            MaskColor       =   12632256
            _Version        =   393216
            BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
               NumListImages   =   1
               BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmQuickLaunch.frx":D7B1
                  Key             =   ""
               EndProperty
            EndProperty
         End
         Begin MSComctlLib.ListView listNewStudent 
            Height          =   2565
            Left            =   60
            TabIndex        =   5
            Top             =   480
            Width           =   6645
            _ExtentX        =   11721
            _ExtentY        =   4524
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            HotTracking     =   -1  'True
            _Version        =   393217
            Icons           =   "ilStudent"
            SmallIcons      =   "ilStudent"
            ForeColor       =   -2147483640
            BackColor       =   16773089
            Appearance      =   0
            NumItems        =   0
         End
         Begin VB.Label lblStudMsg 
            BackStyle       =   0  'Transparent
            Caption         =   "Displays student list which been added and modified or it's other related data."
            ForeColor       =   &H00C00000&
            Height          =   255
            Left            =   90
            TabIndex        =   9
            Top             =   105
            Width           =   6120
         End
      End
      Begin VB.Label lblSchoolAddress 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "--"
         ForeColor       =   &H00FFFAEA&
         Height          =   195
         Left            =   3240
         TabIndex        =   20
         Top             =   465
         Width           =   120
      End
      Begin VB.Label lblSchoolName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "--"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   345
         Left            =   3225
         TabIndex        =   19
         Top             =   120
         Width           =   240
      End
      Begin VB.Label lblPreOut 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "--"
         ForeColor       =   &H00400000&
         Height          =   195
         Left            =   3465
         TabIndex        =   11
         Top             =   1395
         Width           =   120
      End
      Begin VB.Label lblIn 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "--"
         ForeColor       =   &H00400000&
         Height          =   195
         Left            =   3465
         TabIndex        =   10
         Top             =   1185
         Width           =   120
      End
      Begin VB.Label lblCurrentTime 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   4770
         TabIndex        =   3
         Top             =   2745
         Width           =   150
      End
      Begin VB.Label lblUserName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "--"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   4230
         TabIndex        =   2
         Top             =   930
         Width           =   180
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Welcome,"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   3225
         TabIndex        =   1
         Top             =   930
         Width           =   915
      End
      Begin VB.Image Image1 
         Height          =   1125
         Left            =   3360
         Picture         =   "frmQuickLaunch.frx":E08B
         Stretch         =   -1  'True
         Top             =   15
         Width           =   27000
      End
      Begin VB.Image Image3 
         Height          =   15360
         Left            =   0
         Picture         =   "frmQuickLaunch.frx":E1C7
         Stretch         =   -1  'True
         Top             =   2370
         Width           =   420
      End
      Begin VB.Image Image2 
         Height          =   2370
         Left            =   0
         Picture         =   "frmQuickLaunch.frx":E270
         Top             =   0
         Width           =   3375
      End
   End
End
Attribute VB_Name = "frmQuickLaunch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit





Private Sub StartShowAbout()
    me2.Visible = False
    bgMe.Move bgB(1).Width - bgMe.Width, 0, bgMe.Width, 1
    bgMe.Visible = True
    timerAniIn.Enabled = True
End Sub

Private Sub cmdB_Click(Index As Integer)
    Dim i As Integer

    
    'If bgB(Index).Visible = True Then
    '    For i = 0 To cmdB.UBound
    '        cmdB(i).GradientColor = cmdB(i).BackColor
    '        bgB(i).Visible = False
    '    Next
    '
    '    Exit Sub
    'End If
    
    cmdB(Index).GradientColor = &HFFEFE1
    bgB(Index).Visible = True
    
    For i = 0 To cmdB.UBound
        If i <> Index Then
            cmdB(i).GradientColor = cmdB(i).BackColor
            bgB(i).Visible = False
        End If
    Next
    
    ReArrangeControls
    
    Select Case Index
        Case 0 'student
            RefreshRecentStudents
        Case 1 'about
            StartShowAbout
    End Select
    
End Sub

Private Sub cmdPW_Click()
    OpenURL "www.bob8works.cjb.net", mdiMain.hwnd
End Sub

Private Sub cmdVote_Click()
    OpenURL "http://www.pscode.com/vb/scripts/ShowCode.asp?lngWId=1&txtCodeId=65061", Me.hwnd

End Sub

Private Sub Form_Activate()

    Dim frm As Form
    For Each frm In Forms
        If frm.Name <> mdiMain.Name And frm.Name <> Me.Name Then
            Unload frm
        End If
    Next

    mdiMain.RegMDIChild Me
    mdiMain.b8tListOption(3).Expanded = True
    Me.WindowState = vbMaximized
    
    
    On Error Resume Next
    RefreshInfo
    
    If RefreshRecentStudents = True Then
    
        cmdB_Click 0
    
    End If
End Sub


Private Function RefreshInfo()
    lblUserName.Caption = CurrentUser.UserName
    lblIn.Caption = "Time In: " & currentUserLog.TimeIn
        
        
    lblSchoolName.Caption = CurrentSchool.SchoolName
    lblSchoolAddress.Caption = CurrentSchool.Address
End Function

Private Sub Form_Deactivate()
    
    Unload Me
    
    mdiMain.timerMonChild.Enabled = True
End Sub

Private Sub Form_Load()
    'Set cmdB(1).Picture = mdiMain.Icon
    Set Me.Icon = mdiMain.Icon
End Sub

Private Sub Form_Resize()
        
    ReArrangeControls
End Sub

Private Sub ReArrangeControls()
    Dim preLeft As Integer
    Dim i As Integer
    
On Error Resume Next

    bgMain.Move Form_LeftMargin, Form_TopMargin, Me.ScaleWidth - Form_LeftMargin * 2, Me.ScaleHeight - Form_TopMargin * 2

    lblCurrentTime.Left = bgMain.Width - lblCurrentTime.Width - 5

    For i = 0 To bgB.UBound
        If bgB(i).Visible = True Then
            bgB(i).Move bgB(i).Left, bgB(i).Top, bgMain.Width - bgB(i).Left - 4, bgMain.Height - bgB(i).Top - 4
        End If
    Next
    
    If bgB(0).Visible = True Then
        listNewStudent.Width = bgB(0).Width - listNewStudent.Left * 2
        listNewStudent.Height = bgB(0).Height - listNewStudent.Top - listNewStudent.Left * 2
    End If
    
    If bgB(1).Visible = True Then
        bgMe.Move bgB(1).Width - bgMe.Width, 0, bgMe.Width
    End If
End Sub


Private Function RefreshRecentStudents() As Boolean

    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    'default
    RefreshRecentStudents = False
    
    
    sSQL = "SELECT tblStudent.StudentID, Left([tblStudent].[FirstName],1) & '. ' & [tblStudent].[LastName] AS StudentFullName, tblStudent.ModifiedDate, tblStudent.CreationDate, tblGraduate.CreationDate, tblEnrolment.CreationDate, tblEnrolment.ModifiedDate, tblStudentCredential.CreationDate, tblStudentCredential.ModifiedDate" & _
            " FROM ((tblGraduate RIGHT JOIN tblStudent ON tblGraduate.StudentID = tblStudent.StudentID) LEFT JOIN tblStudentCredential ON tblStudent.StudentID = tblStudentCredential.StudentID) LEFT JOIN tblEnrolment ON tblStudent.StudentID = tblEnrolment.StudentID" & _
            " WHERE (((Day([tblStudent].[ModifiedDate]))=Day(Now())) AND ((Year([tblStudent].[ModifiedDate]))=Year(Now()))) OR (((Day([tblStudent].[CreationDate]))=Day(Now())) AND ((Year([tblStudent].[CreationDate]))=Year(Now()))) OR (((Day([tblGraduate].[CreationDate]))>Day(Now())) AND ((Year([tblGraduate].[CreationDate]))>Year(Now()))) OR (((Day([tblEnrolment].[CreationDate]))=Day(Now())) AND ((Year([tblEnrolment].[CreationDate]))=Year(Now()))) OR (((Day([tblEnrolment].[ModifiedDate]))=Day(Now())) AND ((Year([tblEnrolment].[ModifiedDate]))=Year(Now()))) OR (((Day([tblStudentCredential].[CreationDate]))=Day(Now())) AND ((Year([tblStudentCredential].[CreationDate]))=Year(Now()))) OR (((Day([tblStudentCredential].[ModifiedDate]))=Day(Now())) AND ((Year([tblStudentCredential].[ModifiedDate]))=Year(Now())))"
                
    listNewStudent.ListItems.Clear
    
    If ConnectRS(HSESDB, vRS, sSQL) = False Then
        'temp
        'fatal
        'error
        MsgBox "error"
        GoTo ReleaseAndExit
    End If
    
    If AnyRecordExisted(vRS) = False Then
        lblStudMsg.Caption = "There are no transactions."
        listNewStudent.Enabled = False
        GoTo ReleaseAndExit
    Else
        lblStudMsg.Caption = "Displays student list which been added and modified or it's other related data."
        listNewStudent.Enabled = True
    End If
    
    FillRecordToList vRS, listNewStudent, KeyStudent, , , , True
    
    RefreshRecentStudents = True
    
ReleaseAndExit:
    Set vRS = Nothing
End Function


Private Sub listNewStudent_Click()
    Dim lvKey As String
    
    On Error Resume Next
    
    lvKey = GetLVKey(listNewStudent.SelectedItem)
    
    frmStudentRecord.ShowForm lvKey
End Sub

Private Sub timerAniIn_Timer()
    If bgMe.Height < me1.Height Then
        bgMe.Height = bgMe.Height + 10
    Else
        bgMe.Height = me1.Height
        timerAniIn.Enabled = False
        me2.Visible = True
    End If
End Sub

Private Sub timerUT_Timer()
    lblCurrentTime.Caption = "Today is " & FormatDateTime(Now, vbLongDate)
    lblPreOut.Caption = "Pre Log-out time: " & currentUserLog.TimeOut

End Sub

Public Function FormRefresh()
    If bgB(0).Visible = True Then
        RefreshRecentStudents
    End If
End Function
