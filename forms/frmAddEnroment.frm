VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmAddEnrolment 
   BackColor       =   &H00D8E9EC&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Enrolment"
   ClientHeight    =   7605
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9795
   Icon            =   "frmAddEnroment.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   507
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   653
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComCtl2.DTPicker dtDateEnroled 
      Height          =   375
      Left            =   8235
      TabIndex        =   57
      Top             =   690
      Width           =   1470
      _ExtentX        =   2593
      _ExtentY        =   661
      _Version        =   393216
      CalendarBackColor=   16777215
      Format          =   104857601
      CurrentDate     =   38840
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00D8E9EC&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2715
      Left            =   60
      ScaleHeight     =   2715
      ScaleWidth      =   9705
      TabIndex        =   32
      Top             =   1320
      Width           =   9705
      Begin VB.CommandButton cmdGetStudentID 
         BackColor       =   &H00D8E9EC&
         Height          =   330
         Left            =   2310
         Picture         =   "frmAddEnroment.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   46
         Top             =   630
         Width           =   345
      End
      Begin VB.TextBox txtStudentID 
         Height          =   375
         Left            =   60
         MaxLength       =   20
         TabIndex        =   45
         Top             =   600
         Width           =   2625
      End
      Begin TabDlg.SSTab stabStudentDetail 
         Height          =   2415
         Left            =   2910
         TabIndex        =   33
         Top             =   300
         Visible         =   0   'False
         Width           =   6675
         _ExtentX        =   11774
         _ExtentY        =   4260
         _Version        =   393216
         TabOrientation  =   1
         Style           =   1
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         BackColor       =   14215660
         TabCaption(0)   =   "Student Detail"
         TabPicture(0)   =   "frmAddEnroment.frx":0E54
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "frameStudentDetail(0)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Previous Record"
         TabPicture(1)   =   "frmAddEnroment.frx":0E70
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "frameStudentDetail(1)"
         Tab(1).ControlCount=   1
         Begin VB.Frame frameStudentDetail 
            BackColor       =   &H00D8E9EC&
            Caption         =   "Details"
            Height          =   2115
            Index           =   0
            Left            =   0
            TabIndex        =   37
            Top             =   0
            Visible         =   0   'False
            Width           =   6675
            Begin HSES.b8Container b 
               Height          =   1905
               Left            =   60
               TabIndex        =   38
               Top             =   180
               Width           =   6525
               _ExtentX        =   11509
               _ExtentY        =   3360
               BorderColor     =   12307149
               BackColor       =   16185592
               ShadowColor1    =   13427430
               ShadowColor2    =   14215660
               Begin VB.Label lblStudentFullName 
                  Appearance      =   0  'Flat
                  AutoSize        =   -1  'True
                  BackColor       =   &H80000005&
                  BackStyle       =   0  'Transparent
                  Caption         =   "Student Name"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   14.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00C25418&
                  Height          =   345
                  Left            =   675
                  TabIndex        =   44
                  Top             =   105
                  Width           =   2040
               End
               Begin VB.Label lblStudentAge 
                  Appearance      =   0  'Flat
                  AutoSize        =   -1  'True
                  BackColor       =   &H80000005&
                  BackStyle       =   0  'Transparent
                  Caption         =   "Age"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Left            =   705
                  TabIndex        =   43
                  Top             =   570
                  Width           =   390
               End
               Begin VB.Label Label7 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Name:"
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
                  Left            =   165
                  TabIndex        =   42
                  Top             =   180
                  Width           =   465
               End
               Begin VB.Label Label6 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Age:"
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
                  Left            =   165
                  TabIndex        =   41
                  Top             =   585
                  Width           =   345
               End
               Begin VB.Label Label9 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Entrance Grade:"
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
                  Left            =   165
                  TabIndex        =   40
                  Top             =   900
                  Width           =   1185
               End
               Begin VB.Label lblOldAveGrade 
                  Appearance      =   0  'Flat
                  AutoSize        =   -1  'True
                  BackColor       =   &H80000005&
                  BackStyle       =   0  'Transparent
                  Caption         =   "00.00"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Left            =   1440
                  TabIndex        =   39
                  Top             =   885
                  Width           =   540
               End
            End
         End
         Begin VB.Frame frameStudentDetail 
            BackColor       =   &H00D8E9EC&
            Caption         =   "Previous Record"
            Height          =   2115
            Index           =   1
            Left            =   -75000
            TabIndex        =   34
            Top             =   0
            Visible         =   0   'False
            Width           =   6675
            Begin HSES.b8Container b8Container1 
               Height          =   1875
               Left            =   60
               TabIndex        =   35
               Top             =   180
               Width           =   6540
               _ExtentX        =   11536
               _ExtentY        =   3307
               BorderColor     =   12307149
               BackColor       =   16185592
               ShadowColor1    =   13427430
               ShadowColor2    =   14215660
               Begin VB.PictureBox bgEntranceInfo 
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000005&
                  ForeColor       =   &H80000008&
                  Height          =   1755
                  Left            =   60
                  ScaleHeight     =   1725
                  ScaleWidth      =   6375
                  TabIndex        =   50
                  Top             =   60
                  Visible         =   0   'False
                  Width           =   6405
                  Begin VB.Label lblEntranceInfo 
                     Alignment       =   2  'Center
                     Appearance      =   0  'Flat
                     BackColor       =   &H80000005&
                     BackStyle       =   0  'Transparent
                     Caption         =   "No Yet Enrolled"
                     BeginProperty Font 
                        Name            =   "Tahoma"
                        Size            =   9.75
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H80000008&
                     Height          =   240
                     Left            =   30
                     TabIndex        =   51
                     Top             =   720
                     Width           =   6300
                  End
               End
               Begin MSComctlLib.ListView listPrevEnrolments 
                  Height          =   1755
                  Left            =   60
                  TabIndex        =   36
                  Top             =   60
                  Width           =   6435
                  _ExtentX        =   11351
                  _ExtentY        =   3096
                  View            =   3
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
                     Object.Width           =   5292
                  EndProperty
               End
            End
         End
      End
      Begin HSES.b8ChildTitleBar b8ChildTitleBar2 
         Height          =   285
         Left            =   0
         TabIndex        =   48
         Top             =   0
         Width           =   9675
         _ExtentX        =   17066
         _ExtentY        =   503
         BackColor       =   12835550
         Caption         =   "  Student"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontBold        =   -1  'True
         FontName        =   "Tahoma"
         FontSize        =   8.25
         ForeColor       =   8421504
         CloseButton     =   0   'False
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ID"
         Height          =   195
         Left            =   60
         TabIndex        =   47
         Top             =   390
         Width           =   165
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00D8E9EC&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2985
      Left            =   60
      ScaleHeight     =   2985
      ScaleWidth      =   9705
      TabIndex        =   6
      Top             =   4110
      Width           =   9705
      Begin VB.CommandButton cmdGetDepartment 
         BackColor       =   &H00D8E9EC&
         Height          =   330
         Left            =   2325
         Picture         =   "frmAddEnroment.frx":0E8C
         Style           =   1  'Graphical
         TabIndex        =   54
         Top             =   1560
         Width           =   345
      End
      Begin HSES.b8ChildTitleBar b8ChildTitleBar1 
         Height          =   285
         Left            =   0
         TabIndex        =   31
         Top             =   0
         Width           =   9675
         _ExtentX        =   17066
         _ExtentY        =   503
         BackColor       =   12835550
         Caption         =   " School Year / Year Level / Section"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontBold        =   -1  'True
         FontName        =   "Tahoma"
         FontSize        =   8.25
         ForeColor       =   8421504
         CloseButton     =   0   'False
      End
      Begin VB.CommandButton cmdGetSectionOfferingID 
         BackColor       =   &H00D8E9EC&
         Height          =   330
         Left            =   2325
         Picture         =   "frmAddEnroment.frx":1416
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   2370
         Width           =   345
      End
      Begin VB.CommandButton cmdGetSchoolYearTitle 
         BackColor       =   &H00D8E9EC&
         Height          =   330
         Left            =   2325
         Picture         =   "frmAddEnroment.frx":19A0
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   750
         Width           =   345
      End
      Begin VB.TextBox txtSchoolYearTitle 
         Height          =   375
         Left            =   60
         MaxLength       =   20
         TabIndex        =   25
         Top             =   720
         Width           =   2625
      End
      Begin TabDlg.SSTab stabSYSDetail 
         Height          =   2655
         Left            =   2910
         TabIndex        =   7
         Top             =   300
         Visible         =   0   'False
         Width           =   6645
         _ExtentX        =   11721
         _ExtentY        =   4683
         _Version        =   393216
         TabOrientation  =   1
         Style           =   1
         TabHeight       =   520
         BackColor       =   14215660
         TabCaption(0)   =   "Section Detail"
         TabPicture(0)   =   "frmAddEnroment.frx":1F2A
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "frameDetail(0)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Subjects"
         TabPicture(1)   =   "frmAddEnroment.frx":1F46
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "frameDetail(1)"
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "Fees"
         TabPicture(2)   =   "frmAddEnroment.frx":1F62
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "frameDetail(2)"
         Tab(2).ControlCount=   1
         Begin VB.Frame frameDetail 
            BackColor       =   &H00D8E9EC&
            Caption         =   "Fees"
            Height          =   2355
            Index           =   2
            Left            =   -75000
            TabIndex        =   22
            Top             =   0
            Width           =   6645
            Begin HSES.b8Container bgSYSDetail 
               Height          =   2100
               Index           =   2
               Left            =   60
               TabIndex        =   23
               Top             =   210
               Width           =   6510
               _ExtentX        =   11483
               _ExtentY        =   3704
               BorderColor     =   12307149
               BackColor       =   16185592
               ShadowColor1    =   13427430
               ShadowColor2    =   14215660
               Begin MSComctlLib.ListView listFee 
                  Height          =   1980
                  Left            =   60
                  TabIndex        =   24
                  Top             =   60
                  Width           =   6375
                  _ExtentX        =   11245
                  _ExtentY        =   3493
                  View            =   3
                  LabelWrap       =   -1  'True
                  HideSelection   =   0   'False
                  Checkboxes      =   -1  'True
                  FullRowSelect   =   -1  'True
                  GridLines       =   -1  'True
                  _Version        =   393217
                  ForeColor       =   -2147483640
                  BackColor       =   -2147483643
                  Appearance      =   0
                  NumItems        =   1
                  BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     Text            =   "Title"
                     Object.Width           =   5292
                  EndProperty
               End
            End
         End
         Begin VB.Frame frameDetail 
            BackColor       =   &H00D8E9EC&
            Caption         =   "Subjects"
            Height          =   2355
            Index           =   1
            Left            =   -75000
            TabIndex        =   20
            Top             =   0
            Width           =   6645
            Begin HSES.b8Container bgSYSDetail 
               Height          =   2100
               Index           =   1
               Left            =   60
               TabIndex        =   21
               Top             =   210
               Width           =   6510
               _ExtentX        =   11483
               _ExtentY        =   3704
               BorderColor     =   12307149
               BackColor       =   16185592
               ShadowColor1    =   13427430
               ShadowColor2    =   14215660
               Begin MSComctlLib.ListView listSubject 
                  Height          =   1980
                  Left            =   60
                  TabIndex        =   49
                  Top             =   60
                  Width           =   6375
                  _ExtentX        =   11245
                  _ExtentY        =   3493
                  View            =   3
                  LabelWrap       =   -1  'True
                  HideSelection   =   0   'False
                  FullRowSelect   =   -1  'True
                  GridLines       =   -1  'True
                  _Version        =   393217
                  ForeColor       =   -2147483640
                  BackColor       =   -2147483643
                  Appearance      =   0
                  NumItems        =   5
                  BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     Text            =   "Title"
                     Object.Width           =   4762
                  EndProperty
                  BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     SubItemIndex    =   1
                     Text            =   "Days"
                     Object.Width           =   2117
                  EndProperty
                  BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     SubItemIndex    =   2
                     Text            =   "Start"
                     Object.Width           =   2117
                  EndProperty
                  BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     SubItemIndex    =   3
                     Text            =   "End"
                     Object.Width           =   2117
                  EndProperty
                  BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     SubItemIndex    =   4
                     Text            =   "Teacher"
                     Object.Width           =   4762
                  EndProperty
               End
            End
         End
         Begin VB.Frame frameDetail 
            BackColor       =   &H00D8E9EC&
            Height          =   2355
            Index           =   0
            Left            =   0
            TabIndex        =   8
            Top             =   0
            Visible         =   0   'False
            Width           =   6645
            Begin HSES.b8Container bgSYSDetail 
               Height          =   2100
               Index           =   0
               Left            =   60
               TabIndex        =   9
               Top             =   210
               Width           =   6510
               _ExtentX        =   11483
               _ExtentY        =   3704
               BorderColor     =   12307149
               BackColor       =   16185592
               ShadowColor1    =   13427430
               ShadowColor2    =   14215660
               Begin VB.Label lblSectionFullTitle 
                  Appearance      =   0  'Flat
                  AutoSize        =   -1  'True
                  BackColor       =   &H80000005&
                  BackStyle       =   0  'Transparent
                  Caption         =   "Section Title"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   14.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00C25418&
                  Height          =   345
                  Left            =   555
                  TabIndex        =   19
                  Top             =   105
                  Width           =   1800
               End
               Begin VB.Label Label8 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Title:"
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
                  Left            =   135
                  TabIndex        =   18
                  Top             =   195
                  Width           =   360
               End
               Begin VB.Label Label11 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Student Count:"
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
                  Left            =   150
                  TabIndex        =   17
                  Top             =   975
                  Width           =   1110
               End
               Begin VB.Label lblStudentCOunt 
                  Appearance      =   0  'Flat
                  AutoSize        =   -1  'True
                  BackColor       =   &H80000005&
                  BackStyle       =   0  'Transparent
                  Caption         =   "100/100"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Left            =   1380
                  TabIndex        =   16
                  Top             =   945
                  Width           =   840
               End
               Begin VB.Label Label13 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Average Grade Allowed:"
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
                  Left            =   3180
                  TabIndex        =   15
                  Top             =   960
                  Width           =   1755
               End
               Begin VB.Label lblGradeAllowed 
                  Appearance      =   0  'Flat
                  AutoSize        =   -1  'True
                  BackColor       =   &H80000005&
                  BackStyle       =   0  'Transparent
                  Caption         =   "99.99 - 99.99"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Left            =   5055
                  TabIndex        =   14
                  Top             =   930
                  Width           =   1290
               End
               Begin VB.Label Label12 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Note:"
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
                  Left            =   150
                  TabIndex        =   13
                  Top             =   1275
                  Width           =   405
               End
               Begin VB.Label lblNote 
                  BackStyle       =   0  'Transparent
                  Caption         =   "This is Note"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   705
                  Left            =   585
                  TabIndex        =   12
                  Top             =   1305
                  Width           =   5700
                  WordWrap        =   -1  'True
               End
               Begin VB.Label lblDepartmentTitle 
                  Appearance      =   0  'Flat
                  AutoSize        =   -1  'True
                  BackColor       =   &H80000005&
                  BackStyle       =   0  'Transparent
                  Caption         =   "Department Title"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Left            =   1245
                  TabIndex        =   11
                  Top             =   645
                  Width           =   1635
               End
               Begin VB.Label Label15 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Department:"
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
                  Left            =   150
                  TabIndex        =   10
                  Top             =   675
                  Width           =   915
               End
            End
         End
      End
      Begin VB.TextBox txtSectionOfferingID 
         Height          =   375
         Left            =   60
         MaxLength       =   20
         TabIndex        =   28
         Top             =   2340
         Width           =   2625
      End
      Begin VB.TextBox txtDepartmentID 
         Height          =   375
         Left            =   60
         MaxLength       =   20
         TabIndex        =   55
         Top             =   1530
         Width           =   2625
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "(Optional) Department ID"
         Height          =   195
         Left            =   75
         TabIndex        =   56
         Top             =   1320
         Width           =   1755
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Section Offering ID"
         Height          =   195
         Left            =   75
         TabIndex        =   30
         Top             =   2130
         Width           =   1350
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "School Year"
         Height          =   195
         Left            =   75
         TabIndex        =   29
         Top             =   510
         Width           =   870
      End
   End
   Begin VB.TextBox txtEnrolmentID 
      BackColor       =   &H00D8E9EC&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3870
      Locked          =   -1  'True
      MaxLength       =   17
      TabIndex        =   3
      Top             =   690
      Width           =   2370
   End
   Begin HSES.b8Line b8Line1 
      Height          =   60
      Left            =   720
      TabIndex        =   0
      Top             =   510
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   106
   End
   Begin HSES.b8Line b8Line2 
      Height          =   60
      Left            =   -90
      TabIndex        =   1
      Top             =   7095
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   106
   End
   Begin HSES.b8Line b8Line3 
      Height          =   60
      Left            =   0
      TabIndex        =   5
      Top             =   1170
      Width           =   9825
      _ExtentX        =   17330
      _ExtentY        =   106
   End
   Begin lvButton.lvButtons_H cmdSave 
      Default         =   -1  'True
      Height          =   360
      Left            =   8280
      TabIndex        =   52
      Top             =   7200
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
      Left            =   6750
      TabIndex        =   53
      Top             =   7200
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
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Date of Enrollment:"
      Height          =   195
      Left            =   6780
      TabIndex        =   58
      Top             =   765
      Width           =   1350
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Enrolment ID:"
      Height          =   195
      Left            =   2805
      TabIndex        =   4
      Top             =   750
      Width           =   960
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   30
      Picture         =   "frmAddEnroment.frx":1F7E
      Top             =   30
      Width           =   720
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "New Enrolment"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H002F556A&
      Height          =   375
      Left            =   810
      TabIndex        =   2
      Top             =   120
      Width           =   2415
   End
   Begin VB.Image Image2 
      Height          =   525
      Left            =   0
      Picture         =   "frmAddEnroment.frx":2E48
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9825
   End
   Begin VB.Image Image4 
      Height          =   105
      Left            =   0
      Picture         =   "frmAddEnroment.frx":2EE5
      Stretch         =   -1  'True
      Top             =   1080
      Width           =   9765
   End
   Begin VB.Image Image3 
      Height          =   345
      Left            =   60
      Picture         =   "frmAddEnroment.frx":2F82
      Stretch         =   -1  'True
      Top             =   570
      Width           =   9705
   End
End
Attribute VB_Name = "frmAddEnrolment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim RecordAdded As Boolean


Dim selStudentPrevYL As Integer
Dim selStudentPrevSY As String
Dim selStudentPrevDepertmentID As String

Dim selStudentAveGrade As Double
Dim selStudentPassed As Boolean


Dim selSection As tSection
Dim selSectionStudentCount As Long

Dim curEnrolmentID As String

Dim curStudent As tStudent

Public Function ShowForm(Optional sStudentID As String = "", Optional sSchoolYear As String = "", Optional sSectionOfferingID As String = "") As Boolean
    
    '-------------------------------------------------------
    'check user access
    '-------------------------------------------------------
    If UserAllowedTo(CurrentUser.UserName, sCanAddEnrolment) = False Then
        MsgBox "Unable to continue adding Enrolment entry." & vbNewLine & vbNewLine & _
                "You are not permitted to do this. Please contact your administrator for more information.", vbExclamation
        Exit Function
    End If
    '-------------------------------------------------------

    
    mdiMain.MousePointer = vbHourglass
    Me.MousePointer = vbHourglass
    DoEvents
    On Error Resume Next
    CenterForm Me
    Me.Show
    
    'set defaults
    selStudentPrevYL = -1
    selStudentPrevSY = "0000"
    selStudentAveGrade = 0
    selStudentPassed = False

    selSectionStudentCount = 0
    
    'set parameter
    If Len(sStudentID) > 0 Then
        txtStudentID.Text = sStudentID
        txtStudentID.Locked = True
    End If
    
    If Len(sSchoolYear) > 0 Then
        txtSchoolYearTitle.Text = sSchoolYear
        txtSchoolYearTitle.Locked = True
    End If
 
    If sSectionOfferingID <> "" Then
        txtSectionOfferingID.Text = sSectionOfferingID
        txtSchoolYearTitle.Text = Left(sSectionOfferingID, 9)
        txtSectionOfferingID.Locked = True
        txtSchoolYearTitle.Locked = True
    End If
    '-----------------------------------------------------------------
    
    'show form
    Me.Hide
    Me.Show vbModal
    
    'return
    ShowForm = RecordAdded
End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdGetDepartment_Click()
    Dim sDepartmentID As String

    sDepartmentID = PickDepartment.GetItem(, txtDepartmentID)
    If sDepartmentID = "" Then
        Exit Sub
    End If
        
        txtDepartmentID.Text = sDepartmentID
        GenerateAutoSOID
        
    
End Sub

Private Function GenerateAutoSOID()
    
    Dim sSOID As String
    Dim sPYL As Integer
    
    If Len(txtStudentID.Text) < 1 Then
        Exit Function
    End If
    
    If Len(txtSchoolYearTitle.Text) < 1 Then
        Exit Function
    End If
    
    If Len(txtDepartmentID.Text) < 1 Then
        Exit Function
    End If
    
    If curStudent.Transferee = True Then
        GetAutoSectionOffering txtSchoolYearTitle.Text, txtDepartmentID.Text, curStudent.TransfereeYL + 1, curStudent.OldAveGrade, sSOID
    Else
        GetAutoSectionOffering txtSchoolYearTitle.Text, txtDepartmentID.Text, selStudentPrevYL + 1, curStudent.OldAveGrade, sSOID
    End If
    txtSectionOfferingID.Text = sSOID
End Function

Private Sub cmdGetSchoolYearTitle_Click()
    Dim sSchoolYearTitle As String
    
    If txtSchoolYearTitle.Locked = True Then
        MsgBox "School Year was locked." & vbNewLine & _
            "To Add Enrolment with different School Year, please close this window and open 'Add Enrolment' window from main 'Record Menu'", vbInformation
        Exit Sub
    End If


    sSchoolYearTitle = PickSchoolYear.GetItem(txtSchoolYearTitle, , selStudentPrevSY, True)
    
    If sSchoolYearTitle = "" Then
        Exit Sub
    End If
    
    txtSchoolYearTitle.Text = sSchoolYearTitle
    
    GenerateAutoSOID
End Sub

Private Sub cmdGetSectionOfferingID_Click()
    
    Dim sSectionOfferingID As String
    Dim iYLID As Integer
    
    If txtSectionOfferingID.Locked = True Then
    
        MsgBox "Section Offering ID was locked." & vbNewLine & _
            "To Add Enrolment with different Section Offering ID, please close this window and open 'Add Enrolment' window from main 'Record Menu'", vbInformation
        Exit Sub
    End If



    If curStudent.Transferee = True Then
        iYLID = curStudent.TransfereeYL
        sSectionOfferingID = PickSectionOffering.GetSectionOfferingID(txtSectionOfferingID, , txtSchoolYearTitle.Text, iYLID)
    Else
        sSectionOfferingID = PickSectionOffering.GetSectionOfferingID(txtSectionOfferingID, , txtSchoolYearTitle.Text, selStudentPrevYL + 1)
    End If
    If sSectionOfferingID <> "" Then
        txtSectionOfferingID.Text = sSectionOfferingID
    End If
    
End Sub

Private Sub cmdGetStudentID_Click()
    Dim sStudentID As String
    
    If txtStudentID.Locked = True Then
        MsgBox "Student ID was locked." & vbNewLine & _
        "To Add Enrolment with different Student ID, please close this window and open 'Add Enrolment' window from main 'Record Menu'", vbInformation
        Exit Sub
    End If
    
    
    sStudentID = PickStudent.GetStudentID(txtStudentID, , , True, True)
    
    If sStudentID = "" Then
        Exit Sub
    End If
    
        txtStudentID.Text = sStudentID
        
        GenerateAutoSOID
 
    
End Sub


Private Function ShowStudentDetail()
    
    
    Dim NewSY As String
    
    
    'show detail
    stabStudentDetail.Visible = True
    DoEvents
    
    If GetStudentByID(txtStudentID.Text, curStudent) = Success Then
        'Student found
        'set student detail
        '-------------------------------------------------
        
        lblStudentFullName.Caption = curStudent.LastName & ", " & curStudent.FirstName & " " & curStudent.MiddleName
        lblStudentAge.Caption = (Now - curStudent.BirthDate) \ 365
        lblOldAveGrade.Caption = curStudent.OldAveGrade
        

    
        If GetLatestSchoolYearYearLevel(txtStudentID.Text, selStudentPrevSY, selStudentPrevYL) <> Success Then
            'error accessing last SY and YL
            CatchError "frmAddEerolment", "ShowStudentDetail", "error accessing last SY and YL"
            Exit Function
        End If
        
         
        
        
        If curStudent.Transferee = True Then
            'Transferee student
            bgEntranceInfo.Visible = True
            lblEntranceInfo.Caption = "Transferee Student, Year Level: " & YLIDtoTitle(curStudent.TransfereeYL)
        
        ElseIf selStudentPrevYL < 1 Then
            'new student
            bgEntranceInfo.Visible = True
            lblEntranceInfo.Caption = "New Student, Entrance Grade:  " & curStudent.OldAveGrade
        Else
            'old student
            GetAcademicRecord txtStudentID.Text, selStudentPrevYL, selStudentAveGrade, selStudentPassed, selStudentPrevDepertmentID
            txtDepartmentID.Text = selStudentPrevDepertmentID
            If selStudentPassed = False Then
                'temp
                MsgBox "WARNING:" & vbNewLine & vbNewLine & "Student does not passed its previous academic record", vbExclamation
            End If
            bgEntranceInfo.Visible = False
        End If
        '--------------------------------------------------------
    Else
        
        'student not found
        
        'hide details
        stabStudentDetail.Visible = False
        

        txtEnrolmentID.Text = ""
        
        Exit Function
    End If
    
    '-------------------------------------------------------------------
    'auto generate school year
    If txtSchoolYearTitle.Locked = True Then
        
        MsgBox "School Year was locked." & vbNewLine & _
            "To Enrol Student in another School Year, please close this window and Click Record>Add New Enrolment.", vbExclamation

    Else
        If Val(Left(CurrentSchoolYear.SchoolYearTitle, 4)) > Val(Left(selStudentPrevSY, 4)) Then
            txtSchoolYearTitle.Text = CurrentSchoolYear.SchoolYearTitle
        Else
            GetNextSchoolYear selStudentPrevSY, NewSY
            txtSchoolYearTitle.Text = NewSY
        End If
    End If
    
    'generate enrolment id
    GenerateEnrolmentID
    
    FillPrevEnrolment
    
End Function



Public Function FillPrevEnrolment()
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    sSQL = "SELECT tblGrade.EnrolmentID, [tblYearLevel]![YearLevelTitle] & ' - ' & [tblSection]![SectionTitle] AS SectionFullTitle, Avg(tblGrade.GradeValue) AS AvgOfGradeValue, tblEnrolment.SchoolYear, tblDepartment.DepartmentTitle, IIf(Avg([tblGrade].[GradeValue])<75 Or Min([tblGrade].[GradeValue])<75,'Failed','Passed') AS Remark" & _
            " FROM tblYearLevel INNER JOIN (tblDepartment INNER JOIN (tblStudent INNER JOIN (tblSection INNER JOIN (tblSectionOffering INNER JOIN (tblEnrolment INNER JOIN tblGrade ON tblEnrolment.EnrolmentID = tblGrade.EnrolmentID) ON tblSectionOffering.SectionOfferingID = tblEnrolment.SectionOfferingID) ON tblSection.SectionID = tblSectionOffering.SectionID) ON tblStudent.StudentID = tblEnrolment.StudentID) ON tblDepartment.DepartmentID = tblSection.DepartmentID) ON tblYearLevel.YearLevelID = tblSection.YearLevelID" & _
            " Where (((tblStudent.StudentID) = '" & txtStudentID.Text & "'))" & _
            " GROUP BY tblGrade.EnrolmentID, [tblYearLevel]![YearLevelTitle] & ' - ' & [tblSection]![SectionTitle], tblEnrolment.SchoolYear, tblDepartment.DepartmentTitle" & _
            " ORDER BY [tblYearLevel]![YearLevelTitle] & ' - ' & [tblSection]![SectionTitle];"
    
    listPrevEnrolments.ListItems.Clear
    
    If ConnectRS(HSESDB, vRS, sSQL) = True Then
        If AnyRecordExisted(vRS) = True Then
            
            FillRecordToList vRS, listPrevEnrolments, KeyEnrolment
            
            listPrevEnrolments.Visible = True
        Else
            listPrevEnrolments.Visible = False
        End If
    Else
        'fatal error
        CatchError "frmAddEnrolment", "FillPrevEnrolment", "Connection RS - Previous Enrolment Info"
    End If
               
              
    Set vRS = Nothing
End Function
Private Function ShowSectionDetail()
    
    Dim vSectionOffering As tSectionOffering
    Dim vSection As tSection
    Dim vDepartment As tDepartment
    
    'set cursor
    mdiMain.MousePointer = vbHourglass
    
    If GetSectionOfferingByID(txtSectionOfferingID.Text, vSectionOffering) = Success Then
        If GetSectionByID(vSectionOffering.SectionID, vSection) = Success Then
            
            'show bg section
            stabSYSDetail.Visible = True
            
            DoEvents
            
            
            lblSectionFullTitle.Caption = YLIDtoTitle(vSection.YearLevelID) & " - " & vSection.SectionTitle
            
            If GetEnrolmentCountBySectionOfferingID(txtSectionOfferingID.Text, selSectionStudentCount) <> Success Then
                'fatal error
                CatchError "frmAddEnrolment", "ShowSectionDetail", "GetEnrolmentCountBySectionOfferingID(txtSectionOfferingID.Text, selSectionStudentCount) went failed"
            End If
            
            If GetDepartmentByID(vSection.DepartmentID, vDepartment) <> Success Then
                'fatal error
                CatchError "frmAddEnrolment", "ShowSectionDetail", "GetDepartmentByID(vSection.DepartmentID, vdepartment) went failed"
            End If
            
            'set details
            lblStudentCOunt.Caption = selSectionStudentCount & " / " & vSectionOffering.MaxStudentCount
            lblGradeAllowed.Caption = vSectionOffering.MinGrade & " - " & vSectionOffering.MaxGrade
            lblDepartmentTitle.Caption = vDepartment.DepartmentTitle
            lblNote.Caption = vSectionOffering.Note

                        
        Else
            'section not found
            'fatal ERROR
            stabSYSDetail.Visible = False
        End If
        
    Else
        'section not found
        stabSYSDetail.Visible = False
    End If
    
    'restore cursor
    mdiMain.MousePointer = vbDefault
End Function

Private Function ShowSubjectDetail() As Boolean
    
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    
    'set mouse pointer
    mdiMain.MousePointer = vbHourglass
    
    'default
    ShowSubjectDetail = False
    
    listSubject.Enabled = False
    
    sSQL = "SELECT tblSubjectOffering.SubjectOfferingID, tblSubject.SubjectTitle, tblSubjectOffering.Days, tblSubjectOffering.SchedTimeStart,tblSubjectOffering.SchedTimeEnd, [LastName] & ', ' & [FirstName] & ' ' & [MiddleName] AS TeacherFullName" & _
                " FROM tblTeacher INNER JOIN (tblSubject INNER JOIN (tblSectionOffering INNER JOIN tblSubjectOffering ON tblSectionOffering.SectionOfferingID = tblSubjectOffering.SectionOfferingID) ON tblSubject.SubjectID = tblSubjectOffering.SubjectID) ON tblTeacher.TeacherID = tblSubjectOffering.TeacherID" & _
                " WHERE (((tblSectionOffering.SectionOfferingID)='" & Trim(txtSectionOfferingID.Text) & "'))" & _
                " ORDER BY tblSubjectOffering.SchedTimeStart;"

    If ConnectRS(HSESDB, vRS, sSQL) = True Then
        If AnyRecordExisted(vRS) Then
            
            FillRecordToList vRS, listSubject, KeySubject
            listSubject.Enabled = True
            
            ShowSubjectDetail = True
        Else
        
            ShowSubjectDetail = False
        End If
    End If
    Set vRS = Nothing
    
    'restore cursor
    mdiMain.MousePointer = vbDefault
End Function


Private Function ShowFeeDetail()
    
    
    Dim vRS As New ADODB.Recordset
    
    Dim sSQL As String
    
    
    'set mouse pointer
    mdiMain.MousePointer = vbHourglass
    
    listFee.Enabled = False
    
    
    
    sSQL = "SELECT tblFee.FeeID, tblFee.Title, tblFee.Amount, IIf(Len([tblFee]![SchoolYear])<1,'ALL',[tblFee]![SchoolYear]) AS SchoolYear, IIf(Len([tblFee]![DepartmentID])<1,'ALL',[tblDepartment]![DepartmentTitle]) AS Department, IIf([tblFee]![YearLevelID]=0,'ALL',[tblYearLevel]![YearLevelTitle]) AS YearLevel, tblFee.Description, tblFee.CreationDate, tblDepartment.DepartmentID" & _
            " FROM (tblFee LEFT JOIN tblDepartment ON tblFee.DepartmentID = tblDepartment.DepartmentID) LEFT JOIN tblYearLevel ON tblFee.YearLevelID = tblYearLevel.YearLevelID"
    
    If ConnectRS(HSESDB, vRS, sSQL) = True Then
        FillRecordToList vRS, listFee, KeyFee
        'set checked
        SetFeeChecked
        
        listFee.Enabled = True
    End If
    
ReleaseAndExit:
    Set vRS = Nothing
    
    'restore cursor
    mdiMain.MousePointer = vbDefault
End Function

Private Function SetFeeChecked()
    
    Dim vSection As tSection
    Dim vDepartment As tDepartment
    
    Dim lv As ListItem
    Dim bUnchecked As Boolean
    
    If GetSectionBySectionOfferingID(txtSectionOfferingID.Text, vSection) <> Success Then
        'temp
        'fatal error
        Exit Function
    End If
    
    If GetDepartmentByID(vSection.DepartmentID, vDepartment) <> Success Then
        'temp
        'fatal error
        Exit Function
    End If
    
    For Each lv In listFee.ListItems
        'restore
        bUnchecked = False
        
        'check school year
        If lv.SubItems(2) <> "ALL" Then
            
            
            If lv.SubItems(2) = Left(txtSectionOfferingID.Text, 9) Then
                lv.Checked = True
            Else
                bUnchecked = True
                lv.Checked = False
            End If
            
        Else
            
            lv.Checked = True
        End If
                
        'checked department
        If bUnchecked = False Then
            If lv.SubItems(3) <> "ALL" Then
                If lv.SubItems(3) = vDepartment.DepartmentTitle Then
                    lv.Checked = True
                Else
                    bUnchecked = True
                    lv.Checked = False
                End If
            Else
                lv.Checked = True
            End If
        End If
        
        'checked year level
        If bUnchecked = False Then
            If lv.SubItems(4) <> "ALL" Then
                If lv.SubItems(4) = YLIDtoTitle(vSection.YearLevelID) Then
                    lv.Checked = True
                Else
                    bUnchecked = True
                    lv.Checked = False
                End If
            Else
                lv.Checked = True
            End If
        End If
        
    Next
    
End Function

Private Sub cmdSave_Click()

    SaveData

        
End Sub







Private Sub SSTab1_Click(PreviousTab As Integer)
    
End Sub

Private Sub SSTab1_DblClick()

End Sub

Private Sub Form_Activate()
    frameDetail(stabSYSDetail.Tab).Visible = True
    frameStudentDetail(stabStudentDetail.Tab).Visible = True

    mdiMain.MousePointer = vbDefault
    Me.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
    dtDateEnroled.Value = Now
End Sub

Private Sub stabStudentDetail_Click(PreviousTab As Integer)
    frameStudentDetail(stabStudentDetail.Tab).Visible = True
End Sub

Private Sub stabSYSDetail_Click(PreviousTab As Integer)
    frameDetail(stabSYSDetail.Tab).Visible = True
End Sub

Private Sub txtSchoolYearTitle_Change()
    'clear Section Offering ID
    

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

    

    If GenerateEnrolmentID = False Then

        txtEnrolmentID.Text = ""

    End If
    
    Dim vSY As tSchoolYear
    If GetSchoolYearByTitle(txtSchoolYearTitle.Text, vSY) <> Success Then
        Exit Sub
    End If
    
    If vSY.Locked = True Then
        MsgBox "Selected School Year was LOCKED. This entry cannot be used.", vbExclamation
        
        txtSchoolYearTitle.Text = ""
        Exit Sub
    End If
    
    
    
    If Left(txtSectionOfferingID.Text, Len(txtSchoolYearTitle.Text)) <> txtSchoolYearTitle.Text Then
        txtSectionOfferingID.Text = ""
    End If
    
End Sub

    

Private Function GenerateEnrolmentID() As Boolean
    
    Dim IsEnroled As Boolean
    
    'default
    GenerateEnrolmentID = False
    
    txtEnrolmentID.Text = ""
    
    If SchoolYearExistByTitle(txtSchoolYearTitle.Text) <> Success Then
        Exit Function
    End If
    If StudentExistByID(txtStudentID.Text) = Failed Then
        Exit Function
    End If
    
    'check if student is already enroled
    If StudentEnroledBySchoolYear(txtStudentID.Text, txtSchoolYearTitle.Text, IsEnroled) = Success Then
        If IsEnroled = True Then
            MsgBox "Already enroled", vbExclamation

            Exit Function
        End If
    Else
        'fatal error
        CatchError "AddEnrolment", "GenerateEnrolmentID", "Error: StudentEnroledBySchoolYear"
    End If

    'generate id
    txtEnrolmentID.Text = Left(txtSchoolYearTitle.Text, 4) & "-" & txtStudentID.Text

    'return success
    GenerateEnrolmentID = True
End Function

Private Sub txtSchoolYearTitle_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub txtSectionOfferingID_Change()
    
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

    ShowSectionDetail
    
    ShowSubjectDetail
    
    ShowFeeDetail
End Sub

Private Sub txtSectionOfferingID_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub txtStudentID_Change()
    'clear
     selStudentPrevYL = -1
     selStudentPrevSY = "0000"
    
     selStudentAveGrade = 0
     selStudentPassed = False

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

    ShowStudentDetail
End Sub




Private Function SaveData() As Boolean
    
    Dim NewEnrolment As tEnrolment
    
    Dim vSectionOffering As tSectionOffering
    Dim vSection As tSection
    
    Dim vStudent As tStudent
    'default
    SaveData = False
    
    
    
    'check fields
    
    'check student
    '-------------------------------------------------------------------
    If GetStudentByID(txtStudentID.Text, vStudent) <> Success Then
        
        MsgBox "Invalid Student Entry!" & vbNewLine & vbNewLine & _
                "Unable to save Enrolment Entry. The selected Student ID does not exist in record. Please enter correct Student ID.", vbExclamation
                
        Exit Function
    End If
    
    'check if dropped
    If IsStudentDropped(txtStudentID.Text) = Success Then
        MsgBox "Invalid Student Entry!" & vbNewLine & vbNewLine & _
                "Unable to save Enrolment Entry. The selected Student is already Dropped.", vbExclamation
                
        Exit Function
    End If
    
    'check if student passed previous grades
    If selStudentPrevYL > 0 Then
        If selStudentPassed = False Then
            MsgBox "Invalid Student Entry!" & vbNewLine & vbNewLine & _
                    "Unable to save Enrolment Entry. The selected Student does not passed its previous academic record. Please enter another Student.", vbExclamation
    
            Exit Function
        End If
    End If
    
    'check schoolyear
    If SchoolYearExistByTitle(txtSchoolYearTitle.Text) = Failed Then
        MsgBox "Invalid School Year Entry!" & vbNewLine & vbNewLine & _
                "Unable to save Enrolment Entry. The selected School Year does not exist in record. Please enter correct School Year Title.", vbExclamation
                
        Exit Function
    End If
        
        
    'check section Offering
    '-------------------------------------------------------------------
    If GetSectionOfferingByID(txtSectionOfferingID.Text, vSectionOffering) = Failed Then
         
         MsgBox "Invalid Section Offering Entry!" & vbNewLine & vbNewLine & _
                "Unable to save Enrolment Entry. The selected Section Offering ID does not exist in record. Please enter correct Section Offering ID.", vbExclamation
          
          Exit Function
    End If
    
    If Left(txtSectionOfferingID.Text, Len(txtSchoolYearTitle.Text)) <> txtSchoolYearTitle.Text Then
        MsgBox "Invalid Section Offering Entry!" & vbNewLine & vbNewLine & _
                "Unable to save Enrolment Entry. The selected Section Offering Entry does not match the selected School Year Entry. Please enter Section under School Year '" & txtSchoolYearTitle.Text & "'.", vbExclamation
          
          Exit Function
    End If
    
    
    'check section
    '-------------------------------------------------------------------
    If GetSectionByID(vSectionOffering.SectionID, vSection) = Failed Then
         
         MsgBox "Invalid Section Entry!" & vbNewLine & vbNewLine & _
                "Unable to save Enrolment Entry. The selected Section ID does not exist in record. Please enter correct Section ID.", vbExclamation
          
          Exit Function
    End If
    
    'check if year level match for student's next year level
    If curStudent.Transferee = True Then
        
        MsgBox curStudent.TransfereeYL
        MsgBox vSection.YearLevelID
        
        If vSection.YearLevelID <> curStudent.TransfereeYL Then
            MsgBox "Invalid Section Entry!" & vbNewLine & vbNewLine & _
                    "Unable to save Enrolment Entry. The selected Section's Year Level is invalid. Please select Section with '" & YLIDtoTitle(selStudentPrevYL + 1) & "' as Year Level Title.", vbExclamation
              
            Exit Function
        End If
        
    Else
    
        If vSection.YearLevelID <> (selStudentPrevYL + 1) Then
            MsgBox "Invalid Section Entry!" & vbNewLine & vbNewLine & _
                    "Unable to save Enrolment Entry. The selected Section's Year Level is invalid. Please select Section with '" & YLIDtoTitle(selStudentPrevYL + 1) & "' as Year Level Title.", vbExclamation
              
            Exit Function
        End If
        
    End If
    'check student count
    If selSectionStudentCount >= vSectionOffering.MaxStudentCount Then
        MsgBox "Invalid Section Entry!" & vbNewLine & vbNewLine & _
                "Unable to save Enrolment Entry. The selected Section is already Full. Please select another Section.", vbExclamation
          
        Exit Function
    End If
    
    'check grade
    If selStudentPrevYL < 1 Then
        'new student
        If vStudent.OldAveGrade < vSectionOffering.MinGrade Or vStudent.OldAveGrade > vSectionOffering.MaxGrade Then
            MsgBox vStudent.OldAveGrade & " to " & vSectionOffering.MinGrade & "-" & vSectionOffering.MaxGrade
           
            'invalid grade
            MsgBox "Invalid Student Entry!" & vbNewLine & vbNewLine & _
                "Unable to save Enrolment Entry. The selected Student Entrance Grade does not match the selected Section's Minimum and Maxium Grade. Please enter another Student.", vbExclamation
            
            Exit Function
        End If
    Else
        'old student
    End If
    
    'check subjects
    If ShowSubjectDetail = False Then
        MsgBox "Invalid Section Entry!" & vbNewLine & vbNewLine & _
                "Unable to save Enrolment Entry. The selected Section does not contain any Subjects. Please select another Section.", vbExclamation
    
        Exit Function
    End If
        
    
    
    'set enrolment
    '-----------------------------------------------------------------------
    NewEnrolment.EnrolmentID = txtEnrolmentID.Text
    NewEnrolment.StudentID = txtStudentID.Text
    NewEnrolment.SchoolYear = txtSchoolYearTitle.Text
    NewEnrolment.SectionOfferingID = txtSectionOfferingID.Text
    NewEnrolment.DateEnroled = dtDateEnroled.Value
    
    NewEnrolment.CreationDate = Now
    NewEnrolment.CreatedBy = CurrentUser.UserName
    '-----------------------------------------------------------------------
    
    'Add enrolment
    Select Case AddEnrolment(NewEnrolment)
        Case TranDBResult.Success 'SUCCESS
            
            'save charges
            If SaveCharges(NewEnrolment.EnrolmentID) = False Then
                MsgBox "There was an error in saving Charges.", vbExclamation
            End If
            
            MsgBox "New Enrolment successfully added.", vbInformation
            
            SaveData = True
            
            curEnrolmentID = txtStudentID.Text
            Unload Me
            
            frmStudentRecord.ShowForm curEnrolmentID, Me.Name
            
            Unload Me

        Case Else 'unknown result, consider as failed
            'temp
            MsgBox "Enrolment Not added", vbExclamation
            SaveData = False
    End Select
    
End Function

Private Function SaveCharges(sEnrolmentID As String) As Boolean
    
    Dim lv As ListItem
    Dim errorFound As Boolean
    
    'default
    SaveCharges = True
    
    For Each lv In listFee.ListItems
        If lv.Checked = True Then
            If AddCharge(sEnrolmentID, GetLVKey(lv), "", Now, CurrentUser.UserName) <> Success Then
                'fatal error
                'temp
                
                'mark flag
                SaveCharges = False
            End If
        End If
    Next
    
End Function

Private Sub txtStudentID_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub
