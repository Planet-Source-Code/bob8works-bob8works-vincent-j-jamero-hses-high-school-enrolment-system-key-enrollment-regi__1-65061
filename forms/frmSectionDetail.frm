VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmSectionDetail 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Section Offering Details"
   ClientHeight    =   6945
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10110
   ControlBox      =   0   'False
   Icon            =   "frmSectionDetail.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   463
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   674
   WindowState     =   2  'Maximized
   Begin VB.PictureBox bgMain 
      Appearance      =   0  'Flat
      BackColor       =   &H00D8E9EC&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6300
      Left            =   300
      ScaleHeight     =   420
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   528
      TabIndex        =   1
      Top             =   315
      Width           =   7920
      Begin TabDlg.SSTab tabDetail 
         Height          =   4770
         Left            =   0
         TabIndex        =   7
         Top             =   2820
         Width           =   7875
         _ExtentX        =   13891
         _ExtentY        =   8414
         _Version        =   393216
         Style           =   1
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   529
         BackColor       =   14215660
         TabCaption(0)   =   "Students"
         TabPicture(0)   =   "frmSectionDetail.frx":0ECA
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "bgStudents"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Subjects"
         TabPicture(1)   =   "frmSectionDetail.frx":0EE6
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "bgSubjects"
         Tab(1).ControlCount=   1
         Begin HSES.b8SContainer bgStudents 
            Height          =   3645
            Left            =   0
            TabIndex        =   9
            Top             =   280
            Width           =   7860
            _ExtentX        =   13864
            _ExtentY        =   6429
            BorderColor     =   12307149
            Begin MSComctlLib.ImageList imgStudent 
               Left            =   3855
               Top             =   2055
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
                     Picture         =   "frmSectionDetail.frx":0F02
                     Key             =   ""
                  EndProperty
               EndProperty
            End
            Begin MSComctlLib.ListView listEnrolment 
               Height          =   3000
               Left            =   60
               TabIndex        =   11
               Top             =   330
               Width           =   7800
               _ExtentX        =   13758
               _ExtentY        =   5292
               View            =   3
               LabelEdit       =   1
               LabelWrap       =   -1  'True
               HideSelection   =   0   'False
               FullRowSelect   =   -1  'True
               GridLines       =   -1  'True
               _Version        =   393217
               Icons           =   "imgStudent"
               SmallIcons      =   "imgStudent"
               ForeColor       =   -2147483640
               BackColor       =   -2147483643
               Appearance      =   0
               NumItems        =   4
               BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Text            =   "Name"
                  Object.Width           =   6879
               EndProperty
               BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   1
                  Text            =   "Grade"
                  Object.Width           =   3175
               EndProperty
               BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   2
                  Text            =   "Remark"
                  Object.Width           =   2540
               EndProperty
               BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   3
                  Text            =   "Date Enrolled"
                  Object.Width           =   3175
               EndProperty
            End
            Begin VB.CommandButton cmdAddEnrolment 
               BackColor       =   &H00D8E9EC&
               Height          =   315
               Left            =   7440
               Picture         =   "frmSectionDetail.frx":149C
               Style           =   1  'Graphical
               TabIndex        =   12
               Top             =   45
               Width           =   375
            End
            Begin VB.CommandButton cmdDeleteEnrolment 
               BackColor       =   &H00D8E9EC&
               Height          =   315
               Left            =   7140
               Picture         =   "frmSectionDetail.frx":1A26
               Style           =   1  'Graphical
               TabIndex        =   13
               Top             =   45
               Width           =   315
            End
            Begin VB.CommandButton cmdReloadEnrolment 
               BackColor       =   &H00D8E9EC&
               Height          =   315
               Left            =   6750
               Picture         =   "frmSectionDetail.frx":1FB0
               Style           =   1  'Graphical
               TabIndex        =   16
               Top             =   45
               Width           =   405
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "List Of Student/s that are enroled in this section"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C25418&
               Height          =   195
               Left            =   120
               TabIndex        =   10
               Top             =   90
               Width           =   3450
            End
         End
         Begin HSES.b8SContainer bgSubjects 
            Height          =   1665
            Left            =   -75000
            TabIndex        =   8
            Top             =   280
            Width           =   8280
            _ExtentX        =   14605
            _ExtentY        =   2937
            BorderColor     =   12307149
            Begin MSComctlLib.ListView listSubject 
               Height          =   1530
               Left            =   0
               TabIndex        =   14
               Top             =   345
               Width           =   8205
               _ExtentX        =   14473
               _ExtentY        =   2699
               View            =   3
               LabelEdit       =   1
               LabelWrap       =   -1  'True
               HideSelection   =   0   'False
               FullRowSelect   =   -1  'True
               GridLines       =   -1  'True
               _Version        =   393217
               ForeColor       =   -2147483640
               BackColor       =   -2147483643
               Appearance      =   0
               NumItems        =   1
               BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Text            =   "Title"
                  Object.Width           =   4048
               EndProperty
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "List of Subject/s that are offered in this Section"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C25418&
               Height          =   195
               Left            =   120
               TabIndex        =   15
               Top             =   90
               Width           =   3405
            End
         End
      End
      Begin VB.CommandButton cmdGetSectionTitle 
         BackColor       =   &H00D8E9EC&
         Height          =   330
         Left            =   4410
         Picture         =   "frmSectionDetail.frx":253A
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   450
         Width           =   345
      End
      Begin HSES.b8Line b8Line2 
         Height          =   60
         Left            =   -30
         TabIndex        =   17
         Top             =   810
         Width           =   15360
         _ExtentX        =   27093
         _ExtentY        =   106
      End
      Begin HSES.b8Line b8Line1 
         Height          =   60
         Left            =   0
         TabIndex        =   18
         Top             =   2700
         Width           =   15360
         _ExtentX        =   27093
         _ExtentY        =   106
      End
      Begin HSES.b8ChildTitleBar b8Title 
         Height          =   345
         Left            =   -30
         TabIndex        =   5
         Top             =   0
         Width           =   8730
         _ExtentX        =   15399
         _ExtentY        =   609
         BackColor       =   12735512
         Caption         =   "Manage Section Offerings"
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
         ForeColor       =   4210816
         GradTheme       =   2
      End
      Begin HSES.b8Container bgDetail 
         Height          =   1845
         Left            =   0
         TabIndex        =   3
         Top             =   870
         Width           =   10005
         _ExtentX        =   17648
         _ExtentY        =   3254
         BorderColor     =   12307149
         BackColor       =   16185592
         ShadowColor1    =   13427430
         ShadowColor2    =   14215660
         Begin VB.Label lblInfo 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "--"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   195
            Index           =   11
            Left            =   6570
            TabIndex        =   41
            Top             =   1530
            Width           =   120
         End
         Begin VB.Label lblInfo 
            AutoSize        =   -1  'True
            BackColor       =   &H00F6F8F8&
            BackStyle       =   0  'Transparent
            Caption         =   "--"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   195
            Index           =   10
            Left            =   6570
            TabIndex        =   40
            Top             =   1260
            Width           =   120
         End
         Begin VB.Label lblInfo 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "--"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   195
            Index           =   9
            Left            =   6570
            TabIndex        =   39
            Top             =   990
            Width           =   120
         End
         Begin VB.Label lblInfo 
            AutoSize        =   -1  'True
            BackColor       =   &H00F6F8F8&
            BackStyle       =   0  'Transparent
            Caption         =   "--"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   195
            Index           =   8
            Left            =   6570
            TabIndex        =   38
            Top             =   720
            Width           =   120
         End
         Begin VB.Label lblInfo 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "--"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   195
            Index           =   7
            Left            =   6570
            TabIndex        =   37
            Top             =   450
            Width           =   120
         End
         Begin VB.Label Label11 
            BackColor       =   &H00F6F8F8&
            BackStyle       =   0  'Transparent
            Caption         =   "  CreationDate:"
            ForeColor       =   &H00808080&
            Height          =   270
            Left            =   5040
            TabIndex        =   36
            Top             =   690
            Width           =   1500
         End
         Begin VB.Label Label10 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "  Created By:"
            ForeColor       =   &H00808080&
            Height          =   270
            Left            =   5040
            TabIndex        =   35
            Top             =   960
            Width           =   1500
         End
         Begin VB.Label Label9 
            BackColor       =   &H00F6F8F8&
            BackStyle       =   0  'Transparent
            Caption         =   "  Modified Date:"
            ForeColor       =   &H00808080&
            Height          =   270
            Left            =   5040
            TabIndex        =   34
            Top             =   1230
            Width           =   1500
         End
         Begin VB.Label Label8 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "  Modified By:"
            ForeColor       =   &H00808080&
            Height          =   270
            Left            =   5040
            TabIndex        =   33
            Top             =   1500
            Width           =   1500
         End
         Begin VB.Label Label18 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "  Note"
            ForeColor       =   &H00808080&
            Height          =   270
            Left            =   5040
            TabIndex        =   32
            Top             =   420
            Width           =   1500
         End
         Begin VB.Label lblInfo 
            AutoSize        =   -1  'True
            BackColor       =   &H00F6F8F8&
            BackStyle       =   0  'Transparent
            Caption         =   "--"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   195
            Index           =   6
            Left            =   1590
            TabIndex        =   31
            Top             =   1500
            Width           =   120
         End
         Begin VB.Label lblInfo 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "--"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   195
            Index           =   5
            Left            =   1590
            TabIndex        =   30
            Top             =   1230
            Width           =   120
         End
         Begin VB.Label lblInfo 
            AutoSize        =   -1  'True
            BackColor       =   &H00F6F8F8&
            BackStyle       =   0  'Transparent
            Caption         =   "--"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   195
            Index           =   4
            Left            =   1590
            TabIndex        =   29
            Top             =   960
            Width           =   120
         End
         Begin VB.Label lblInfo 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "--"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   195
            Index           =   3
            Left            =   1590
            TabIndex        =   28
            Top             =   690
            Width           =   120
         End
         Begin VB.Label Label7 
            BackColor       =   &H00F6F8F8&
            BackStyle       =   0  'Transparent
            Caption         =   "  Max. Grade"
            ForeColor       =   &H00808080&
            Height          =   270
            Left            =   150
            TabIndex        =   27
            Top             =   1500
            Width           =   1500
         End
         Begin VB.Label Label5 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "  Min. Grade"
            ForeColor       =   &H00808080&
            Height          =   270
            Left            =   150
            TabIndex        =   26
            Top             =   1230
            Width           =   1500
         End
         Begin VB.Label Label12 
            BackColor       =   &H00F6F8F8&
            BackStyle       =   0  'Transparent
            Caption         =   "  Max. Student #"
            ForeColor       =   &H00808080&
            Height          =   270
            Left            =   150
            TabIndex        =   25
            Top             =   960
            Width           =   1500
         End
         Begin VB.Label TeacherName 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "  TeacherName"
            ForeColor       =   &H00808080&
            Height          =   270
            Left            =   150
            TabIndex        =   24
            Top             =   690
            Width           =   1500
         End
         Begin VB.Label lblInfo 
            AutoSize        =   -1  'True
            BackColor       =   &H00F6F8F8&
            Caption         =   "--"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   195
            Index           =   2
            Left            =   1590
            TabIndex        =   23
            Top             =   450
            Width           =   120
         End
         Begin VB.Label Label13 
            BackColor       =   &H00F6F8F8&
            Caption         =   "  School Year"
            ForeColor       =   &H00808080&
            Height          =   270
            Left            =   150
            TabIndex        =   22
            Top             =   420
            Width           =   1500
         End
         Begin VB.Label lblInfo 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "--"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   390
            Index           =   1
            Left            =   270
            TabIndex        =   21
            Top             =   60
            Width           =   3330
         End
         Begin VB.Label lblInfo 
            BackColor       =   &H00F6F8F8&
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
            ForeColor       =   &H00C00000&
            Height          =   270
            Index           =   0
            Left            =   2220
            TabIndex        =   20
            Top             =   0
            Visible         =   0   'False
            Width           =   480
         End
         Begin VB.Label Label15 
            BackColor       =   &H00F6F8F8&
            Caption         =   "  Section Offering ID"
            ForeColor       =   &H00808080&
            Height          =   270
            Left            =   1680
            TabIndex        =   19
            Top             =   0
            Visible         =   0   'False
            Width           =   1500
         End
      End
      Begin VB.TextBox txtSectionOfferingID 
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
         Height          =   375
         Left            =   1530
         MaxLength       =   20
         TabIndex        =   4
         Top             =   420
         Width           =   3255
      End
      Begin VB.Image Image4 
         Height          =   105
         Left            =   -90
         Picture         =   "frmSectionDetail.frx":2AC4
         Stretch         =   -1  'True
         Top             =   720
         Width           =   30000
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Section Offering ID"
         Height          =   195
         Left            =   75
         TabIndex        =   6
         Top             =   480
         Width           =   1350
      End
      Begin VB.Image Image3 
         Height          =   345
         Left            =   0
         Picture         =   "frmSectionDetail.frx":2B61
         Stretch         =   -1  'True
         Top             =   360
         Width           =   30000
      End
   End
   Begin HSES.b8Container b8cMain 
      Height          =   6585
      Left            =   -180
      TabIndex        =   0
      Top             =   300
      Width           =   8760
      _ExtentX        =   15452
      _ExtentY        =   11615
      BorderColor     =   12632256
      BackColor       =   16185592
      ShadowColor1    =   16185592
      ShadowColor2    =   16185592
   End
End
Attribute VB_Name = "frmSectionDetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Dim curSchoolYearTitle As String
Dim curSectionOfferingID As String

Dim curForm_CanPrint As Boolean

Public Function ShowForm(Optional sSectionOfferingID As String = "", Optional iTab As Integer = 0)
        
    'show form
    Me.Show
    DoEvents

    tabDetail.Tab = iTab
    'tabDetail_Click 0
    
    curSectionOfferingID = sSectionOfferingID
    
    txtSectionOfferingID.Text = sSectionOfferingID


End Function

Private Sub cmdAddEnrolment_Click()
    Me.Enabled = False
    mdiMain.MousePointer = vbHourglass
    
    If frmAddEnrolment.ShowForm(, , txtSectionOfferingID.Text) = True Then
        GenerateStudentList
    End If
    
    mdiMain.MousePointer = vbDefault
    Me.Enabled = True
End Sub


Private Sub cmdDeleteEnrolment_Click()
    frmDeleteEnrolment.ShowForm GetLVKey(listEnrolment.SelectedItem)
End Sub

Private Sub cmdGetSectionTitle_Click()
    Dim sSectionOfferingID As String
    
    sSectionOfferingID = PickSectionOffering.GetSectionOfferingID(txtSectionOfferingID)
    
    If sSectionOfferingID <> "" Then
        txtSectionOfferingID.Text = sSectionOfferingID
    End If
End Sub

Private Sub cmdReloadEnrolment_Click()
    GenerateStudentList
End Sub

Private Sub Form_Activate()
    mdiMain.RegMDIChild Me
    Me.WindowState = vbMaximized
End Sub

Private Sub Form_Resize()

    On Error Resume Next
    
    Me.ScaleMode = vbPixels
    b8cMain.Move Form_LeftMargin - 3, Form_TopMargin - 3, Me.ScaleWidth - (Form_LeftMargin - 3) * 2, Me.ScaleHeight - (Form_TopMargin - 3) * 2
    
    bgMain.Move Form_LeftMargin, Form_TopMargin, Me.ScaleWidth - Form_LeftMargin * 2, Me.ScaleHeight - Form_TopMargin * 2
    
    b8Title.Width = bgMain.Width
    bgDetail.Width = bgMain.Width
    
    
    'arrange tab
    tabDetail.Move tabDetail.Left, tabDetail.Top, bgMain.Width - tabDetail.Left * 2, bgMain.Height - tabDetail.Top
    
    
    ArrangeTabCtl

End Sub





Private Sub tabDetail_Click(PreviousTab As Integer)
     ArrangeTabCtl

End Sub

Private Sub ArrangeTabCtl()
    Select Case tabDetail.Tab
        Case 0
            bgStudents.Move 0, bgStudents.Top, tabDetail.Width * Screen.TwipsPerPixelX - 1, (tabDetail.Height * Screen.TwipsPerPixelY) - bgStudents.Top - 1
            listEnrolment.Move 1, listEnrolment.Top, bgStudents.Width - 2, bgStudents.Height - listEnrolment.Top - 2
        
            cmdAddEnrolment.Left = bgStudents.Width - cmdAddEnrolment.Width
            cmdDeleteEnrolment.Left = cmdAddEnrolment.Left - cmdDeleteEnrolment.Width - 1
            cmdReloadEnrolment.Left = cmdDeleteEnrolment.Left - cmdDeleteEnrolment.Width - 1
            
        Case 1
            bgSubjects.Move 0, bgSubjects.Top, tabDetail.Width * Screen.TwipsPerPixelX - 1, (tabDetail.Height * Screen.TwipsPerPixelY) - bgSubjects.Top - 1
            listSubject.Move 1, listSubject.Top, bgSubjects.Width - 2, bgSubjects.Height - listSubject.Top - 2
        Case 2
            
            'bgSectionOfferingDetail.Move 1, 'bgSectionOfferingDetail.Top, tabDetail.Width * Screen.TwipsPerPixelX - 1, (tabDetail.Height * Screen.TwipsPerPixelY) - 'bgSectionOfferingDetail.Top - 1
            
        End Select

End Sub

Private Function ClearSectionOfferingDetail()
    'clear
    'lblInfo(0).Caption = "--"
    lblInfo(1).Caption = "--"
    lblInfo(2).Caption = "--"
    lblInfo(3).Caption = "--"
    lblInfo(4).Caption = "--"
    lblInfo(5).Caption = "--"
    lblInfo(6).Caption = "--"
    lblInfo(7).Caption = "--"
    lblInfo(8).Caption = "--"
    lblInfo(9).Caption = "--"
    lblInfo(10).Caption = "--"
    lblInfo(11).Caption = "--"
End Function
Private Sub ShowSectionOfferingDetail()
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    
    'clear
    ClearSectionOfferingDetail
    
    
    sSQL = "SELECT tblSectionOffering.SectionOfferingID, [tblYearLevel]![YearLevelTitle] & ' - ' & [tblSection]![SectionTitle] AS SectionFullTitle, tblSectionOffering.SchoolYear, [tblTeacher]![LastName] & ', ' & [tblTeacher]![FirstName] & ' ' & [tblTeacher]![MiddleName] AS TeacherFullName, tblSectionOffering.MaxStudentCount, tblSectionOffering.MinGrade, tblSectionOffering.MaxGrade, tblSectionOffering.Note, tblSectionOffering.CreationDate, tblSectionOffering.CreatedBy, tblSectionOffering.ModifiedDate, tblSectionOffering.ModifiedBy" & _
            " FROM tblTeacher INNER JOIN (tblYearLevel INNER JOIN (tblSection INNER JOIN tblSectionOffering ON tblSection.SectionID = tblSectionOffering.SectionID) ON tblYearLevel.YearLevelID = tblSection.YearLevelID) ON tblTeacher.TeacherID = tblSectionOffering.TeacherID" & _
            " WHERE (((tblSectionOffering.SectionOfferingID)='" & txtSectionOfferingID.Text & "'));"

    mdiMain.MousePointer = vbHourglass
    
    If ConnectRS(HSESDB, vRS, sSQL) = False Then
        'fatal error
        CatchError "frmSectionOfferingDetail", "ShowSectionOfferingDetail", "Unable to connect Recordset"
        GoTo ReleaseAndExit
    End If
    
        
    If AnyRecordExisted(vRS) = False Then
        'not found
        GoTo ReleaseAndExit
    End If

    'set form fields
    'lblInfo(0).Caption = ReadField(vRS.Fields("SectionOfferingID"))
    lblInfo(1).Caption = ReadField(vRS.Fields("SectionFullTitle"))
    lblInfo(2).Caption = ReadField(vRS.Fields("SchoolYear"))
    lblInfo(3).Caption = ReadField(vRS.Fields("TeacherFullName"))
    lblInfo(4).Caption = ReadField(vRS.Fields("MaxStudentCount"))
    lblInfo(5).Caption = ReadField(vRS.Fields("MinGrade"))
    lblInfo(6).Caption = ReadField(vRS.Fields("MaxGrade"))
    lblInfo(7).Caption = ReadField(vRS.Fields("Note"))
    lblInfo(8).Caption = ReadField(vRS.Fields("CreationDate"))
    lblInfo(9).Caption = ReadField(vRS.Fields("CreatedBy"))
    If Len(ReadField(vRS.Fields("ModifiedBy"))) > 0 Then
        lblInfo(10).Caption = ReadField(vRS.Fields("ModifiedDate"))
        lblInfo(11).Caption = ReadField(vRS.Fields("ModifiedBy"))
    End If
'------------------------------------------------------------
ReleaseAndExit:
    Set vRS = Nothing
    
    mdiMain.MousePointer = vbDefault
End Sub

Private Sub txtSectionOfferingID_Change()

    Dim vSectionOffering As tSectionOffering
    Dim sSectionFullTitle As String
    Dim vDepartment As tDepartment
    Dim vSection As tSection
    
    'default
    ClearSectionOfferingDetail
    
    curForm_CanPrint = False
    mdiMain.RegMDIChild Me
    
    
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
    
    
    If Len(txtSectionOfferingID.Text) < 1 Then Exit Sub
    
        If GetSectionOfferingByID(txtSectionOfferingID.Text, vSectionOffering) = Success Then
            'found
     
            
                        
            listEnrolment.Enabled = True
            listSubject.Enabled = True
            
            GenerateSubjectList
            
            GenerateStudentList
            
            ShowSectionOfferingDetail
            
            curForm_CanPrint = True
            
            mdiMain.RegMDIChild Me
            
            
            
            If mdiMain.b8tListOption(4).Expanded = True Then
                GetSectionByID vSectionOffering.SectionID, vSection
                GetDepartmentByID vSection.DepartmentID, vDepartment
                
                mdiMain.RecordTree.SelectNode KeySectionOffering & ";" & vSectionOffering.SchoolYear & ";" & vDepartment.DepartmentTitle & ";" & YLIDtoTitle(vSection.YearLevelID) & ";" & txtSectionOfferingID.Text
            End If
        Else
            'id not found
            'lblSectionFullTitle.Caption = ""
            'lblSchoolYearTitle.Caption = ""
            listEnrolment.ListItems.Clear
            listSubject.ListItems.Clear
            listEnrolment.Enabled = False
            listSubject.Enabled = False
            listSubject.ListItems.Clear
            listEnrolment.ListItems.Clear

        End If
        
        Form_Resize

End Sub

Private Function GenerateSubjectList()
    
   Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    
    'set mouse pointer
    mdiMain.MousePointer = vbHourglass
    
    listSubject.Enabled = False
    
    sSQL = "SELECT tblSubjectOffering.SubjectOfferingID, tblSubject.SubjectTitle, tblSubjectOffering.Days, tblSubjectOffering.SchedTimeStart,tblSubjectOffering.SchedTimeEnd, [LastName] & ', ' & [FirstName] & ' ' & [MiddleName] AS TeacherFullName" & _
                " FROM tblTeacher INNER JOIN (tblSubject INNER JOIN (tblSectionOffering INNER JOIN tblSubjectOffering ON tblSectionOffering.SectionOfferingID = tblSubjectOffering.SectionOfferingID) ON tblSubject.SubjectID = tblSubjectOffering.SubjectID) ON tblTeacher.TeacherID = tblSubjectOffering.TeacherID" & _
                " WHERE (((tblSectionOffering.SectionOfferingID)='" & txtSectionOfferingID.Text & "'))" & _
                " ORDER BY tblSubjectOffering.SchedTimeStart;"

    If ConnectRS(HSESDB, vRS, sSQL) = True Then
        FillRecordToList vRS, listSubject, KeySubject
        listSubject.Enabled = True
    End If
    
    Set vRS = Nothing
    
    'restore cursor
    mdiMain.MousePointer = vbDefault

End Function

Private Function GenerateStudentList()
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String


    sSQL = "SELECT tblEnrolment.EnrolmentID, [tblStudent]![LastName] & ', ' & [tblStudent]![FirstName] & ' ' & [tblStudent]![MiddleName] AS FullName, Avg(tblGrade.GradeValue) AS AvgOfGradeValue, IIf(Avg([tblGrade]![GradeValue])<75 Or Min([tblGrade]![GradeValue])<75,'Failed','Passed') AS Remark, tblEnrolment.DateEnroled" & _
            " FROM tblStudent INNER JOIN ((tblSection INNER JOIN tblSectionOffering ON tblSection.SectionID = tblSectionOffering.SectionID) INNER JOIN (tblEnrolment INNER JOIN tblGrade ON tblEnrolment.EnrolmentID = tblGrade.EnrolmentID) ON tblSectionOffering.SectionOfferingID = tblEnrolment.SectionOfferingID) ON tblStudent.StudentID = tblEnrolment.StudentID" & _
            " WHERE (((tblSectionOffering.SectionOfferingID)='" & txtSectionOfferingID.Text & "'))" & _
            " GROUP BY tblEnrolment.EnrolmentID, [tblStudent]![LastName] & ', ' & [tblStudent]![FirstName] & ' ' & [tblStudent]![MiddleName], tblEnrolment.DateEnroled" & _
            " ORDER BY [tblStudent]![LastName] & ', ' & [tblStudent]![FirstName] & ' ' & [tblStudent]![MiddleName]"

    listEnrolment.ListItems.Clear

    
    If ConnectRS(HSESDB, vRS, sSQL) = True Then
        If AnyRecordExisted(vRS) = True Then
            FillRecordToList vRS, listEnrolment, KeyEnrolment, , , , True
        Else
            'no records
        End If
    Else
        'fatal error
        CatchError "frmSectionDetail", "GenerateEnrolmentList", "Error connecting Enrolments"
    End If
    
    
    Set vRS = Nothing
End Function







'------------------------------------------------------
Public Function Form_CanExplore() As Boolean
    Form_CanExplore = True
End Function

Public Function Form_CanPrint() As Boolean
    Form_CanPrint = curForm_CanPrint
End Function

Public Function Form_Print()
    frmPrintSectionOffering.ShowForm txtSectionOfferingID.Text
End Function
