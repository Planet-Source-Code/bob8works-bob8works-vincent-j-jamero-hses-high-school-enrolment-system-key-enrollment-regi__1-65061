VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSYStat 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Statistics"
   ClientHeight    =   6270
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7935
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   418
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   529
   WindowState     =   2  'Maximized
   Begin VB.PictureBox bgMain 
      Appearance      =   0  'Flat
      BackColor       =   &H00D8E9EC&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5190
      Left            =   585
      ScaleHeight     =   346
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   471
      TabIndex        =   1
      Top             =   480
      Width           =   7065
      Begin VB.Timer timerContractHiddenFields 
         Interval        =   1
         Left            =   5955
         Top             =   1605
      End
      Begin VB.PictureBox bgGT 
         Appearance      =   0  'Flat
         BackColor       =   &H00D8E9EC&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   540
         Left            =   540
         ScaleHeight     =   540
         ScaleWidth      =   5460
         TabIndex        =   8
         Top             =   4590
         Width           =   5460
         Begin VB.TextBox txtGF 
            Alignment       =   1  'Right Justify
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
            Left            =   3420
            Locked          =   -1  'True
            TabIndex        =   13
            Text            =   "0"
            Top             =   180
            Width           =   900
         End
         Begin VB.TextBox txtGM 
            Alignment       =   1  'Right Justify
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
            Left            =   2475
            Locked          =   -1  'True
            TabIndex        =   11
            Text            =   "0"
            Top             =   180
            Width           =   840
         End
         Begin VB.TextBox txtGrandTotal 
            Alignment       =   1  'Right Justify
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
            Left            =   4440
            Locked          =   -1  'True
            TabIndex        =   9
            Text            =   "0"
            Top             =   195
            Width           =   1020
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "All"
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
            Left            =   4470
            TabIndex        =   15
            Top             =   -15
            Width           =   165
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Female:"
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
            Left            =   3420
            TabIndex        =   14
            Top             =   0
            Width           =   570
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Male:"
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
            Left            =   2475
            TabIndex        =   12
            Top             =   0
            Width           =   390
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Grand Total:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   75
            TabIndex        =   10
            Top             =   30
            Width           =   1035
         End
      End
      Begin HSES.b8ChildTitleBar b8Title 
         Height          =   345
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   6585
         _ExtentX        =   11615
         _ExtentY        =   609
         BackColor       =   12735512
         Caption         =   "Statistics By School Year"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontBold        =   -1  'True
         FontName        =   "Tahoma"
         FontSize        =   9.75
         ForeColor       =   8421504
         GradTheme       =   2
      End
      Begin MSComctlLib.ImageList icoHeader 
         Left            =   5805
         Top             =   3210
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   2
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmSYStat.frx":0000
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmSYStat.frx":059A
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin HSES.b8SContainer pbBGButton 
         Height          =   585
         Left            =   -15
         TabIndex        =   3
         Top             =   345
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   1032
         BorderColor     =   14215660
         Begin VB.CommandButton cmdGetSchoolYear 
            BackColor       =   &H00FFFFFF&
            Height          =   300
            Left            =   3930
            Picture         =   "frmSYStat.frx":0B34
            Style           =   1  'Graphical
            TabIndex        =   5
            Top             =   165
            Width           =   345
         End
         Begin VB.TextBox txtSchoolYearTitle 
            BackColor       =   &H00C0FFFF&
            Height          =   345
            Left            =   1080
            MaxLength       =   50
            TabIndex        =   6
            Top             =   135
            Width           =   3225
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "School Year"
            Height          =   195
            Left            =   165
            TabIndex        =   7
            Top             =   180
            Width           =   870
         End
      End
      Begin MSComctlLib.ListView listRecord 
         Height          =   3405
         Left            =   15
         TabIndex        =   4
         Top             =   1035
         Width           =   5505
         _ExtentX        =   9710
         _ExtentY        =   6006
         View            =   3
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         Icons           =   "ilRecordIco"
         SmallIcons      =   "ilRecordIco"
         ColHdrIcons     =   "icoHeader"
         ForeColor       =   8399906
         BackColor       =   16777215
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   7
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "S.Y."
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Department"
            Object.Width           =   4233
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Y.L."
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Section"
            Object.Width           =   3704
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Male Enrolles"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Female Enrolles"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Total Enrolles"
            Object.Width           =   4233
         EndProperty
      End
      Begin MSComctlLib.ImageList ilRecordIco 
         Left            =   30
         Top             =   0
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   1
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmSYStat.frx":10BE
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin HSES.b8Container b8cMain 
      Height          =   5940
      Left            =   180
      TabIndex        =   0
      Top             =   165
      Width           =   7035
      _ExtentX        =   12409
      _ExtentY        =   10478
      BorderColor     =   12632256
      BackColor       =   16777215
      InsideBorderColor=   14215660
      ShadowColor1    =   16777215
      ShadowColor2    =   16777215
   End
End
Attribute VB_Name = "frmSYstat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Dim curVRS As New ADODB.Recordset
Dim curSQL As String

    
Public Function ShowForm(Optional sSYTitle As String = "")
    
    Me.Show
    On Error Resume Next
    Me.SetFocus
    
    
    If sSYTitle = "" Then
        sSYTitle = CurrentSchoolYear.SchoolYearTitle
    Else
        txtSchoolYearTitle.Text = sSYTitle
    End If
    
End Function



Private Sub cmdGetSchoolYear_Click()
    Dim sSchoolYearTitle As String
    
    sSchoolYearTitle = PickSchoolYear.GetItem(txtSchoolYearTitle, , , True)
    
    If sSchoolYearTitle <> "" Then
        txtSchoolYearTitle.Text = sSchoolYearTitle
    End If

End Sub


Public Function Form_FillList()
    
    mdiMain.MousePointer = vbHourglass
    
    'clear
    '--------------------------------------
    listRecord.ListItems.Clear
    listRecord.Enabled = False
    txtGrandTotal.Text = ""
    txtGM.Text = ""
    txtGF.Text = ""
    '--------------------------------------
    
    curSQL = "SELECT * FROM refStatisticBySY WHERE SchoolYearTitle='" & txtSchoolYearTitle.Text & "'"
    
    If ConnectRS(HSESDB, curVRS, curSQL) = False Then
        'error
        CatchError "frmSYStat", "Form_FillList", "Unable to connect Recordset with SQL Expression '" & curSQL & "'"
        GoTo ReleaseAndExit
    End If
        
    If AnyRecordExisted(curVRS) = False Then
        GoTo ReleaseAndExit
    End If

    'set ui
    '------------------------------------------------------------------
   On Error Resume Next
    txtGM.Text = curVRS.Fields("GTM")
    txtGF.Text = curVRS.Fields("GTF")
    txtGrandTotal.Text = curVRS.Fields("GT")
    
    
    FillRecordToList curVRS, listRecord, KeySectionOffering, , 32767, , True
    
    listRecord.ColumnHeaders.Item(8).Width = 0
    listRecord.ColumnHeaders.Item(9).Width = 0
    listRecord.ColumnHeaders.Item(10).Width = 0
    listRecord.Enabled = True
    
ReleaseAndExit:
    mdiMain.RegMDIChild Me
    mdiMain.MousePointer = vbDefault
End Function

Public Function Form_Refresh()
    Form_FillList
End Function

Public Function Form_CanPrint() As Boolean
    
    If listRecord.ListItems.Count > 0 Then
        Form_CanPrint = True
    Else
        Form_CanPrint = False
    End If
    
End Function

Public Function Form_Print()

    Set drSYStat.DataSource = curVRS
    curVRS.MoveFirst
    drSYStat.Sections("Section5").Controls("lblGM").Caption = ReadField(curVRS.Fields("GTM"))
    drSYStat.Sections("Section5").Controls("lblGF").Caption = ReadField(curVRS.Fields("GTF"))
    drSYStat.Sections("Section5").Controls("lblGrandTotal").Caption = ReadField(curVRS.Fields("GT"))
    drSYStat.Sections("Section4").Controls("lblSY").Caption = txtSchoolYearTitle.Text
    drSYStat.Show vbModal
    
End Function




























Private Sub Form_Activate()
    mdiMain.RegMDIChild Me
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    Me.ScaleMode = vbPixels
    b8cMain.Move Form_LeftMargin - 3, Form_TopMargin - 3, Me.ScaleWidth - (Form_LeftMargin - 3) * 2, Me.ScaleHeight - (Form_TopMargin - 3) * 2
   
    bgMain.Move Form_LeftMargin, Form_TopMargin, Me.ScaleWidth - Form_LeftMargin * 2, Me.ScaleHeight - Form_TopMargin * 2
    
    b8Title.Move 0, 0, bgMain.Width
    pbBGButton.Move 0, b8Title.Top + b8Title.Height, bgMain.Width
    listRecord.Move listRecord.Left, pbBGButton.Top + pbBGButton.Height, bgMain.Width - (listRecord.Left * 2)
    listRecord.Height = bgMain.Height - (listRecord.Top + bgGT.Height)
    
    bgGT.Move bgMain.Width - bgGT.Width, listRecord.Top + listRecord.Height + 2
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set curVRS = Nothing
End Sub

Private Sub timerContractHiddenFields_Timer()
    On Error Resume Next
    If listRecord.ColumnHeaders.Item(8).Width <> 0 Then listRecord.ColumnHeaders.Item(8).Width = 0
    If listRecord.ColumnHeaders.Item(9).Width <> 0 Then listRecord.ColumnHeaders.Item(9).Width = 0
    If listRecord.ColumnHeaders.Item(10).Width <> 0 Then listRecord.ColumnHeaders.Item(10).Width = 0
End Sub

Private Sub txtSchoolYearTitle_Change()
    'delay 0.3 second
    'code by: VIncent J. Jamero
    '------------------------------------------------
    Static DelayStart As Single
    Static notFirst As Boolean
    DelayStart = GetTickCount + 300
    If notFirst = True Then Exit Sub
    notFirst = True
    While GetTickCount < DelayStart
        DoEvents
    Wend
    notFirst = False
    '------------------------------------------------
    'the next line will be if executed if user pause typing in 0.3 second
    
    If Len(Trim(txtSchoolYearTitle)) > 0 Then
        Form_FillList
    End If
    
End Sub
