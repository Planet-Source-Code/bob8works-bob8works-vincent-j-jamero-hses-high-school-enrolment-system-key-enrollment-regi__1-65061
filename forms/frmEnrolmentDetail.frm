VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmStudentRecord 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Student Record"
   ClientHeight    =   9330
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9810
   ControlBox      =   0   'False
   Icon            =   "frmEnrolmentDetail.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   622
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   654
   WindowState     =   2  'Maximized
   Begin VB.PictureBox bgMain 
      Appearance      =   0  'Flat
      BackColor       =   &H00D8E9EC&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6345
      Left            =   0
      ScaleHeight     =   423
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   651
      TabIndex        =   0
      Top             =   30
      Width           =   9765
      Begin HSES.b8Line b8Line2 
         Height          =   60
         Left            =   -30
         TabIndex        =   21
         Top             =   810
         Width           =   15360
         _ExtentX        =   27093
         _ExtentY        =   106
      End
      Begin TabDlg.SSTab tabMAin 
         Height          =   4785
         Left            =   0
         TabIndex        =   6
         Top             =   2040
         Width           =   9735
         _ExtentX        =   17171
         _ExtentY        =   8440
         _Version        =   393216
         Style           =   1
         Tab             =   1
         TabHeight       =   529
         Enabled         =   0   'False
         BackColor       =   14215660
         TabCaption(0)   =   "Personal Info"
         TabPicture(0)   =   "frmEnrolmentDetail.frx":0ECA
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "bgTabCon(0)"
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Records"
         TabPicture(1)   =   "frmEnrolmentDetail.frx":0EE6
         Tab(1).ControlEnabled=   -1  'True
         Tab(1).Control(0)=   "bgTabCon(1)"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "Credentials Passed"
         TabPicture(2)   =   "frmEnrolmentDetail.frx":0F02
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "bgTabCon(2)"
         Tab(2).ControlCount=   1
         Begin HSES.b8SContainer bgTabCon 
            Height          =   4905
            Index           =   0
            Left            =   -75000
            TabIndex        =   7
            Top             =   300
            Width           =   7860
            _ExtentX        =   13864
            _ExtentY        =   8652
            BorderColor     =   12307149
            Begin RichTextLib.RichTextBox rtbInfo 
               Height          =   2355
               Left            =   30
               TabIndex        =   9
               Top             =   240
               Width           =   4305
               _ExtentX        =   7594
               _ExtentY        =   4154
               _Version        =   393217
               BackColor       =   16777215
               BorderStyle     =   0
               ReadOnly        =   -1  'True
               ScrollBars      =   3
               TextRTF         =   $"frmEnrolmentDetail.frx":0F1E
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin VB.Image imgTop 
               Height          =   165
               Index           =   0
               Left            =   30
               Picture         =   "frmEnrolmentDetail.frx":0F98
               Stretch         =   -1  'True
               Top             =   90
               Width           =   15360
            End
         End
         Begin HSES.b8SContainer bgTabCon 
            Height          =   4995
            Index           =   1
            Left            =   0
            TabIndex        =   8
            Top             =   300
            Width           =   9690
            _ExtentX        =   17092
            _ExtentY        =   8811
            BorderColor     =   12307149
            Begin HSES.b8SideBar bgAllYL 
               Height          =   4455
               Left            =   30
               TabIndex        =   10
               Top             =   210
               Width           =   9435
               _ExtentX        =   16642
               _ExtentY        =   7858
               BackColor       =   16185592
               BackColor       =   16185592
               BorderColor1    =   16185592
               BorderColor2    =   16185592
               BorderColor3    =   16185592
               BorderColor4    =   16185592
               BorderColor5    =   16185592
               Begin VB.PictureBox bgYLRec 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00E0E0E0&
                  BorderStyle     =   0  'None
                  ForeColor       =   &H80000008&
                  Height          =   2895
                  Index           =   3
                  Left            =   0
                  ScaleHeight     =   193
                  ScaleMode       =   3  'Pixel
                  ScaleWidth      =   627
                  TabIndex        =   72
                  Top             =   0
                  Visible         =   0   'False
                  Width           =   9405
                  Begin VB.PictureBox bgYLDetail 
                     Appearance      =   0  'Flat
                     BackColor       =   &H00FFFFFF&
                     BorderStyle     =   0  'None
                     ForeColor       =   &H80000008&
                     Height          =   2865
                     Index           =   3
                     Left            =   15
                     ScaleHeight     =   2865
                     ScaleWidth      =   3195
                     TabIndex        =   73
                     Top             =   15
                     Width           =   3195
                     Begin HSES.b8Line b8Line1 
                        Height          =   60
                        Index           =   3
                        Left            =   -105
                        TabIndex        =   74
                        Top             =   2475
                        Width           =   3315
                        _ExtentX        =   5847
                        _ExtentY        =   106
                        BorderColor1    =   12632256
                        BorderColor2    =   16777215
                        BorderColor3    =   16777215
                     End
                     Begin lvButton.lvButtons_H cmdPrint 
                        Height          =   300
                        Index           =   3
                        Left            =   1890
                        TabIndex        =   75
                        Top             =   2535
                        Width           =   1290
                        _ExtentX        =   2275
                        _ExtentY        =   529
                        Caption         =   "Print"
                        CapAlign        =   2
                        BackStyle       =   4
                        BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                           Name            =   "MS Sans Serif"
                           Size            =   8.25
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        cBhover         =   14215660
                        cGradient       =   14215660
                        Gradient        =   3
                        Mode            =   0
                        Value           =   0   'False
                        ImgSize         =   32
                        cBack           =   16185592
                     End
                     Begin VB.Image Image6 
                        Height          =   720
                        Left            =   2445
                        Picture         =   "frmEnrolmentDetail.frx":1035
                        Top             =   1770
                        Width           =   720
                     End
                     Begin VB.Line Line1 
                        BorderColor     =   &H00E0E0E0&
                        BorderStyle     =   3  'Dot
                        Index           =   3
                        X1              =   3870
                        X2              =   0
                        Y1              =   360
                        Y2              =   360
                     End
                     Begin VB.Label lblSectionFullTitle 
                        AutoSize        =   -1  'True
                        BackStyle       =   0  'Transparent
                        Caption         =   "Transferee"
                        BeginProperty Font 
                           Name            =   "Tahoma"
                           Size            =   12
                           Charset         =   0
                           Weight          =   700
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        ForeColor       =   &H00FF0000&
                        Height          =   285
                        Index           =   3
                        Left            =   150
                        TabIndex        =   86
                        Top             =   60
                        Width           =   1335
                     End
                     Begin VB.Label lblSchoolYear 
                        AutoSize        =   -1  'True
                        BackStyle       =   0  'Transparent
                        Caption         =   " "
                        BeginProperty Font 
                           Name            =   "Tahoma"
                           Size            =   8.25
                           Charset         =   0
                           Weight          =   700
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        ForeColor       =   &H00808080&
                        Height          =   195
                        Index           =   3
                        Left            =   300
                        TabIndex        =   85
                        Top             =   600
                        Width           =   45
                     End
                     Begin VB.Label lblDepartmentTitle 
                        AutoSize        =   -1  'True
                        BackStyle       =   0  'Transparent
                        Caption         =   " "
                        BeginProperty Font 
                           Name            =   "Tahoma"
                           Size            =   8.25
                           Charset         =   0
                           Weight          =   700
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        ForeColor       =   &H00808080&
                        Height          =   195
                        Index           =   3
                        Left            =   285
                        TabIndex        =   84
                        Top             =   1035
                        Width           =   45
                     End
                     Begin VB.Label lblAdviser 
                        AutoSize        =   -1  'True
                        BackStyle       =   0  'Transparent
                        Caption         =   " "
                        BeginProperty Font 
                           Name            =   "Tahoma"
                           Size            =   8.25
                           Charset         =   0
                           Weight          =   700
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        ForeColor       =   &H00808080&
                        Height          =   195
                        Index           =   3
                        Left            =   285
                        TabIndex        =   83
                        Top             =   1425
                        Width           =   45
                     End
                     Begin VB.Label Label24 
                        Appearance      =   0  'Flat
                        AutoSize        =   -1  'True
                        BackColor       =   &H80000005&
                        BackStyle       =   0  'Transparent
                        Caption         =   "S.Y.:"
                        BeginProperty Font 
                           Name            =   "Tahoma"
                           Size            =   8.25
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        ForeColor       =   &H80000008&
                        Height          =   195
                        Left            =   300
                        TabIndex        =   82
                        Top             =   405
                        Width           =   360
                     End
                     Begin VB.Label Label23 
                        Appearance      =   0  'Flat
                        AutoSize        =   -1  'True
                        BackColor       =   &H80000005&
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
                        ForeColor       =   &H80000008&
                        Height          =   195
                        Left            =   300
                        TabIndex        =   81
                        Top             =   840
                        Width           =   915
                     End
                     Begin VB.Label Label22 
                        Appearance      =   0  'Flat
                        AutoSize        =   -1  'True
                        BackColor       =   &H80000005&
                        BackStyle       =   0  'Transparent
                        Caption         =   "Adviser:"
                        BeginProperty Font 
                           Name            =   "Tahoma"
                           Size            =   8.25
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        ForeColor       =   &H80000008&
                        Height          =   195
                        Left            =   300
                        TabIndex        =   80
                        Top             =   1245
                        Width           =   600
                     End
                     Begin VB.Label Label21 
                        Appearance      =   0  'Flat
                        AutoSize        =   -1  'True
                        BackColor       =   &H80000005&
                        BackStyle       =   0  'Transparent
                        Caption         =   "Date Enrolled:"
                        BeginProperty Font 
                           Name            =   "Tahoma"
                           Size            =   8.25
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        ForeColor       =   &H80000008&
                        Height          =   195
                        Left            =   300
                        TabIndex        =   79
                        Top             =   1650
                        Width           =   1020
                     End
                     Begin VB.Label lblDateEnrolled 
                        AutoSize        =   -1  'True
                        BackStyle       =   0  'Transparent
                        Caption         =   " "
                        BeginProperty Font 
                           Name            =   "Tahoma"
                           Size            =   8.25
                           Charset         =   0
                           Weight          =   700
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        ForeColor       =   &H00808080&
                        Height          =   195
                        Index           =   3
                        Left            =   285
                        TabIndex        =   78
                        Top             =   1830
                        Width           =   45
                     End
                     Begin VB.Label Label20 
                        Appearance      =   0  'Flat
                        AutoSize        =   -1  'True
                        BackColor       =   &H80000005&
                        BackStyle       =   0  'Transparent
                        Caption         =   "Gen. Average Grade"
                        BeginProperty Font 
                           Name            =   "Tahoma"
                           Size            =   8.25
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        ForeColor       =   &H80000008&
                        Height          =   195
                        Left            =   300
                        TabIndex        =   77
                        Top             =   2070
                        Width           =   1485
                     End
                     Begin VB.Label lblAG 
                        AutoSize        =   -1  'True
                        BackStyle       =   0  'Transparent
                        Caption         =   " "
                        BeginProperty Font 
                           Name            =   "Tahoma"
                           Size            =   8.25
                           Charset         =   0
                           Weight          =   700
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        ForeColor       =   &H00808080&
                        Height          =   195
                        Index           =   3
                        Left            =   285
                        TabIndex        =   76
                        Top             =   2250
                        Width           =   45
                     End
                  End
                  Begin MSComctlLib.ListView listSubject 
                     Height          =   2865
                     Index           =   3
                     Left            =   3225
                     TabIndex        =   87
                     Top             =   15
                     Width           =   5565
                     _ExtentX        =   9816
                     _ExtentY        =   5054
                     View            =   3
                     LabelEdit       =   1
                     LabelWrap       =   -1  'True
                     HideSelection   =   -1  'True
                     GridLines       =   -1  'True
                     _Version        =   393217
                     Icons           =   "ilSubject"
                     SmallIcons      =   "ilSubject"
                     ForeColor       =   -2147483640
                     BackColor       =   16777215
                     Appearance      =   0
                     NumItems        =   5
                     BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                        Text            =   "Title"
                        Object.Width           =   2646
                     EndProperty
                     BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                        SubItemIndex    =   1
                        Text            =   "Grade"
                        Object.Width           =   1323
                     EndProperty
                     BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                        SubItemIndex    =   2
                        Text            =   "Time"
                        Object.Width           =   1852
                     EndProperty
                     BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                        SubItemIndex    =   3
                        Text            =   "Days"
                        Object.Width           =   1587
                     EndProperty
                     BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                        SubItemIndex    =   4
                        Text            =   "Teacher"
                        Object.Width           =   3175
                     EndProperty
                  End
               End
               Begin VB.PictureBox bgYLRec 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00E0E0E0&
                  BorderStyle     =   0  'None
                  ForeColor       =   &H80000008&
                  Height          =   2895
                  Index           =   2
                  Left            =   0
                  ScaleHeight     =   193
                  ScaleMode       =   3  'Pixel
                  ScaleWidth      =   627
                  TabIndex        =   56
                  Top             =   0
                  Visible         =   0   'False
                  Width           =   9405
                  Begin VB.PictureBox bgYLDetail 
                     Appearance      =   0  'Flat
                     BackColor       =   &H00FFFFFF&
                     BorderStyle     =   0  'None
                     ForeColor       =   &H80000008&
                     Height          =   2865
                     Index           =   2
                     Left            =   15
                     ScaleHeight     =   2865
                     ScaleWidth      =   3195
                     TabIndex        =   57
                     Top             =   15
                     Width           =   3195
                     Begin HSES.b8Line b8Line1 
                        Height          =   60
                        Index           =   2
                        Left            =   -105
                        TabIndex        =   58
                        Top             =   2475
                        Width           =   3315
                        _ExtentX        =   5847
                        _ExtentY        =   106
                        BorderColor1    =   12632256
                        BorderColor2    =   16777215
                        BorderColor3    =   16777215
                     End
                     Begin lvButton.lvButtons_H cmdPrint 
                        Height          =   300
                        Index           =   2
                        Left            =   1890
                        TabIndex        =   59
                        Top             =   2535
                        Width           =   1290
                        _ExtentX        =   2275
                        _ExtentY        =   529
                        Caption         =   "Print"
                        CapAlign        =   2
                        BackStyle       =   4
                        BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                           Name            =   "MS Sans Serif"
                           Size            =   8.25
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        cBhover         =   14215660
                        cGradient       =   14215660
                        Gradient        =   3
                        Mode            =   0
                        Value           =   0   'False
                        ImgSize         =   32
                        cBack           =   16185592
                     End
                     Begin VB.Image Image5 
                        Height          =   720
                        Left            =   2445
                        Picture         =   "frmEnrolmentDetail.frx":1EFF
                        Top             =   1770
                        Width           =   720
                     End
                     Begin VB.Line Line1 
                        BorderColor     =   &H00E0E0E0&
                        BorderStyle     =   3  'Dot
                        Index           =   2
                        X1              =   3870
                        X2              =   0
                        Y1              =   360
                        Y2              =   360
                     End
                     Begin VB.Label lblSectionFullTitle 
                        AutoSize        =   -1  'True
                        BackStyle       =   0  'Transparent
                        Caption         =   "Transferee"
                        BeginProperty Font 
                           Name            =   "Tahoma"
                           Size            =   12
                           Charset         =   0
                           Weight          =   700
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        ForeColor       =   &H00FF0000&
                        Height          =   285
                        Index           =   2
                        Left            =   150
                        TabIndex        =   70
                        Top             =   60
                        Width           =   1335
                     End
                     Begin VB.Label lblSchoolYear 
                        AutoSize        =   -1  'True
                        BackStyle       =   0  'Transparent
                        Caption         =   " "
                        BeginProperty Font 
                           Name            =   "Tahoma"
                           Size            =   8.25
                           Charset         =   0
                           Weight          =   700
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        ForeColor       =   &H00808080&
                        Height          =   195
                        Index           =   2
                        Left            =   300
                        TabIndex        =   69
                        Top             =   600
                        Width           =   45
                     End
                     Begin VB.Label lblDepartmentTitle 
                        AutoSize        =   -1  'True
                        BackStyle       =   0  'Transparent
                        Caption         =   " "
                        BeginProperty Font 
                           Name            =   "Tahoma"
                           Size            =   8.25
                           Charset         =   0
                           Weight          =   700
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        ForeColor       =   &H00808080&
                        Height          =   195
                        Index           =   2
                        Left            =   285
                        TabIndex        =   68
                        Top             =   1035
                        Width           =   45
                     End
                     Begin VB.Label lblAdviser 
                        AutoSize        =   -1  'True
                        BackStyle       =   0  'Transparent
                        Caption         =   " "
                        BeginProperty Font 
                           Name            =   "Tahoma"
                           Size            =   8.25
                           Charset         =   0
                           Weight          =   700
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        ForeColor       =   &H00808080&
                        Height          =   195
                        Index           =   2
                        Left            =   285
                        TabIndex        =   67
                        Top             =   1425
                        Width           =   45
                     End
                     Begin VB.Label Label19 
                        Appearance      =   0  'Flat
                        AutoSize        =   -1  'True
                        BackColor       =   &H80000005&
                        BackStyle       =   0  'Transparent
                        Caption         =   "S.Y.:"
                        BeginProperty Font 
                           Name            =   "Tahoma"
                           Size            =   8.25
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        ForeColor       =   &H80000008&
                        Height          =   195
                        Left            =   300
                        TabIndex        =   66
                        Top             =   405
                        Width           =   360
                     End
                     Begin VB.Label Label18 
                        Appearance      =   0  'Flat
                        AutoSize        =   -1  'True
                        BackColor       =   &H80000005&
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
                        ForeColor       =   &H80000008&
                        Height          =   195
                        Left            =   300
                        TabIndex        =   65
                        Top             =   840
                        Width           =   915
                     End
                     Begin VB.Label Label17 
                        Appearance      =   0  'Flat
                        AutoSize        =   -1  'True
                        BackColor       =   &H80000005&
                        BackStyle       =   0  'Transparent
                        Caption         =   "Adviser:"
                        BeginProperty Font 
                           Name            =   "Tahoma"
                           Size            =   8.25
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        ForeColor       =   &H80000008&
                        Height          =   195
                        Left            =   300
                        TabIndex        =   64
                        Top             =   1245
                        Width           =   600
                     End
                     Begin VB.Label Label16 
                        Appearance      =   0  'Flat
                        AutoSize        =   -1  'True
                        BackColor       =   &H80000005&
                        BackStyle       =   0  'Transparent
                        Caption         =   "Date Enrolled:"
                        BeginProperty Font 
                           Name            =   "Tahoma"
                           Size            =   8.25
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        ForeColor       =   &H80000008&
                        Height          =   195
                        Left            =   300
                        TabIndex        =   63
                        Top             =   1650
                        Width           =   1020
                     End
                     Begin VB.Label lblDateEnrolled 
                        AutoSize        =   -1  'True
                        BackStyle       =   0  'Transparent
                        Caption         =   " "
                        BeginProperty Font 
                           Name            =   "Tahoma"
                           Size            =   8.25
                           Charset         =   0
                           Weight          =   700
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        ForeColor       =   &H00808080&
                        Height          =   195
                        Index           =   2
                        Left            =   285
                        TabIndex        =   62
                        Top             =   1830
                        Width           =   45
                     End
                     Begin VB.Label Label15 
                        Appearance      =   0  'Flat
                        AutoSize        =   -1  'True
                        BackColor       =   &H80000005&
                        BackStyle       =   0  'Transparent
                        Caption         =   "Gen. Average Grade"
                        BeginProperty Font 
                           Name            =   "Tahoma"
                           Size            =   8.25
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        ForeColor       =   &H80000008&
                        Height          =   195
                        Left            =   300
                        TabIndex        =   61
                        Top             =   2070
                        Width           =   1485
                     End
                     Begin VB.Label lblAG 
                        AutoSize        =   -1  'True
                        BackStyle       =   0  'Transparent
                        Caption         =   " "
                        BeginProperty Font 
                           Name            =   "Tahoma"
                           Size            =   8.25
                           Charset         =   0
                           Weight          =   700
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        ForeColor       =   &H00808080&
                        Height          =   195
                        Index           =   2
                        Left            =   285
                        TabIndex        =   60
                        Top             =   2250
                        Width           =   45
                     End
                  End
                  Begin MSComctlLib.ListView listSubject 
                     Height          =   2865
                     Index           =   2
                     Left            =   3225
                     TabIndex        =   71
                     Top             =   15
                     Width           =   5565
                     _ExtentX        =   9816
                     _ExtentY        =   5054
                     View            =   3
                     LabelEdit       =   1
                     LabelWrap       =   -1  'True
                     HideSelection   =   -1  'True
                     GridLines       =   -1  'True
                     _Version        =   393217
                     Icons           =   "ilSubject"
                     SmallIcons      =   "ilSubject"
                     ForeColor       =   -2147483640
                     BackColor       =   16777215
                     Appearance      =   0
                     NumItems        =   5
                     BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                        Text            =   "Title"
                        Object.Width           =   2646
                     EndProperty
                     BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                        SubItemIndex    =   1
                        Text            =   "Grade"
                        Object.Width           =   1323
                     EndProperty
                     BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                        SubItemIndex    =   2
                        Text            =   "Time"
                        Object.Width           =   1852
                     EndProperty
                     BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                        SubItemIndex    =   3
                        Text            =   "Days"
                        Object.Width           =   1587
                     EndProperty
                     BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                        SubItemIndex    =   4
                        Text            =   "Teacher"
                        Object.Width           =   3175
                     EndProperty
                  End
               End
               Begin VB.PictureBox bgYLRec 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00E0E0E0&
                  BorderStyle     =   0  'None
                  ForeColor       =   &H80000008&
                  Height          =   2895
                  Index           =   1
                  Left            =   0
                  ScaleHeight     =   193
                  ScaleMode       =   3  'Pixel
                  ScaleWidth      =   627
                  TabIndex        =   40
                  Top             =   0
                  Visible         =   0   'False
                  Width           =   9405
                  Begin VB.PictureBox bgYLDetail 
                     Appearance      =   0  'Flat
                     BackColor       =   &H00FFFFFF&
                     BorderStyle     =   0  'None
                     ForeColor       =   &H80000008&
                     Height          =   2865
                     Index           =   0
                     Left            =   15
                     ScaleHeight     =   2865
                     ScaleWidth      =   3195
                     TabIndex        =   41
                     Top             =   15
                     Width           =   3195
                     Begin HSES.b8Line b8Line1 
                        Height          =   60
                        Index           =   1
                        Left            =   -105
                        TabIndex        =   42
                        Top             =   2475
                        Width           =   3315
                        _ExtentX        =   5847
                        _ExtentY        =   106
                        BorderColor1    =   12632256
                        BorderColor2    =   16777215
                        BorderColor3    =   16777215
                     End
                     Begin lvButton.lvButtons_H cmdPrint 
                        Height          =   300
                        Index           =   1
                        Left            =   1890
                        TabIndex        =   43
                        Top             =   2535
                        Width           =   1290
                        _ExtentX        =   2275
                        _ExtentY        =   529
                        Caption         =   "Print"
                        CapAlign        =   2
                        BackStyle       =   4
                        BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                           Name            =   "MS Sans Serif"
                           Size            =   8.25
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        cBhover         =   14215660
                        cGradient       =   14215660
                        Gradient        =   3
                        Mode            =   0
                        Value           =   0   'False
                        ImgSize         =   32
                        cBack           =   16185592
                     End
                     Begin VB.Image Image2 
                        Height          =   720
                        Left            =   2445
                        Picture         =   "frmEnrolmentDetail.frx":2DC9
                        Top             =   1770
                        Width           =   720
                     End
                     Begin VB.Line Line1 
                        BorderColor     =   &H00E0E0E0&
                        BorderStyle     =   3  'Dot
                        Index           =   1
                        X1              =   3870
                        X2              =   0
                        Y1              =   360
                        Y2              =   360
                     End
                     Begin VB.Label lblSectionFullTitle 
                        AutoSize        =   -1  'True
                        BackStyle       =   0  'Transparent
                        Caption         =   "Transferee"
                        BeginProperty Font 
                           Name            =   "Tahoma"
                           Size            =   12
                           Charset         =   0
                           Weight          =   700
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        ForeColor       =   &H00FF0000&
                        Height          =   285
                        Index           =   1
                        Left            =   150
                        TabIndex        =   54
                        Top             =   60
                        Width           =   1335
                     End
                     Begin VB.Label lblSchoolYear 
                        AutoSize        =   -1  'True
                        BackStyle       =   0  'Transparent
                        Caption         =   " "
                        BeginProperty Font 
                           Name            =   "Tahoma"
                           Size            =   8.25
                           Charset         =   0
                           Weight          =   700
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        ForeColor       =   &H00808080&
                        Height          =   195
                        Index           =   1
                        Left            =   300
                        TabIndex        =   53
                        Top             =   600
                        Width           =   45
                     End
                     Begin VB.Label lblDepartmentTitle 
                        AutoSize        =   -1  'True
                        BackStyle       =   0  'Transparent
                        Caption         =   " "
                        BeginProperty Font 
                           Name            =   "Tahoma"
                           Size            =   8.25
                           Charset         =   0
                           Weight          =   700
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        ForeColor       =   &H00808080&
                        Height          =   195
                        Index           =   1
                        Left            =   285
                        TabIndex        =   52
                        Top             =   1035
                        Width           =   45
                     End
                     Begin VB.Label lblAdviser 
                        AutoSize        =   -1  'True
                        BackStyle       =   0  'Transparent
                        Caption         =   " "
                        BeginProperty Font 
                           Name            =   "Tahoma"
                           Size            =   8.25
                           Charset         =   0
                           Weight          =   700
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        ForeColor       =   &H00808080&
                        Height          =   195
                        Index           =   1
                        Left            =   285
                        TabIndex        =   51
                        Top             =   1425
                        Width           =   45
                     End
                     Begin VB.Label Label14 
                        Appearance      =   0  'Flat
                        AutoSize        =   -1  'True
                        BackColor       =   &H80000005&
                        BackStyle       =   0  'Transparent
                        Caption         =   "S.Y.:"
                        BeginProperty Font 
                           Name            =   "Tahoma"
                           Size            =   8.25
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        ForeColor       =   &H80000008&
                        Height          =   195
                        Left            =   300
                        TabIndex        =   50
                        Top             =   405
                        Width           =   360
                     End
                     Begin VB.Label Label13 
                        Appearance      =   0  'Flat
                        AutoSize        =   -1  'True
                        BackColor       =   &H80000005&
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
                        ForeColor       =   &H80000008&
                        Height          =   195
                        Left            =   300
                        TabIndex        =   49
                        Top             =   840
                        Width           =   915
                     End
                     Begin VB.Label Label12 
                        Appearance      =   0  'Flat
                        AutoSize        =   -1  'True
                        BackColor       =   &H80000005&
                        BackStyle       =   0  'Transparent
                        Caption         =   "Adviser:"
                        BeginProperty Font 
                           Name            =   "Tahoma"
                           Size            =   8.25
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        ForeColor       =   &H80000008&
                        Height          =   195
                        Left            =   300
                        TabIndex        =   48
                        Top             =   1245
                        Width           =   600
                     End
                     Begin VB.Label Label11 
                        Appearance      =   0  'Flat
                        AutoSize        =   -1  'True
                        BackColor       =   &H80000005&
                        BackStyle       =   0  'Transparent
                        Caption         =   "Date Enrolled:"
                        BeginProperty Font 
                           Name            =   "Tahoma"
                           Size            =   8.25
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        ForeColor       =   &H80000008&
                        Height          =   195
                        Left            =   300
                        TabIndex        =   47
                        Top             =   1650
                        Width           =   1020
                     End
                     Begin VB.Label lblDateEnrolled 
                        AutoSize        =   -1  'True
                        BackStyle       =   0  'Transparent
                        Caption         =   " "
                        BeginProperty Font 
                           Name            =   "Tahoma"
                           Size            =   8.25
                           Charset         =   0
                           Weight          =   700
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        ForeColor       =   &H00808080&
                        Height          =   195
                        Index           =   1
                        Left            =   285
                        TabIndex        =   46
                        Top             =   1830
                        Width           =   45
                     End
                     Begin VB.Label Label10 
                        Appearance      =   0  'Flat
                        AutoSize        =   -1  'True
                        BackColor       =   &H80000005&
                        BackStyle       =   0  'Transparent
                        Caption         =   "Gen. Average Grade"
                        BeginProperty Font 
                           Name            =   "Tahoma"
                           Size            =   8.25
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        ForeColor       =   &H80000008&
                        Height          =   195
                        Left            =   300
                        TabIndex        =   45
                        Top             =   2070
                        Width           =   1485
                     End
                     Begin VB.Label lblAG 
                        AutoSize        =   -1  'True
                        BackStyle       =   0  'Transparent
                        Caption         =   " "
                        BeginProperty Font 
                           Name            =   "Tahoma"
                           Size            =   8.25
                           Charset         =   0
                           Weight          =   700
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        ForeColor       =   &H00808080&
                        Height          =   195
                        Index           =   1
                        Left            =   285
                        TabIndex        =   44
                        Top             =   2250
                        Width           =   45
                     End
                  End
                  Begin MSComctlLib.ListView listSubject 
                     Height          =   2865
                     Index           =   1
                     Left            =   3225
                     TabIndex        =   55
                     Top             =   15
                     Width           =   5565
                     _ExtentX        =   9816
                     _ExtentY        =   5054
                     View            =   3
                     LabelEdit       =   1
                     LabelWrap       =   -1  'True
                     HideSelection   =   -1  'True
                     GridLines       =   -1  'True
                     _Version        =   393217
                     Icons           =   "ilSubject"
                     SmallIcons      =   "ilSubject"
                     ForeColor       =   -2147483640
                     BackColor       =   16777215
                     Appearance      =   0
                     NumItems        =   5
                     BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                        Text            =   "Title"
                        Object.Width           =   2646
                     EndProperty
                     BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                        SubItemIndex    =   1
                        Text            =   "Grade"
                        Object.Width           =   1323
                     EndProperty
                     BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                        SubItemIndex    =   2
                        Text            =   "Time"
                        Object.Width           =   1852
                     EndProperty
                     BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                        SubItemIndex    =   3
                        Text            =   "Days"
                        Object.Width           =   1587
                     EndProperty
                     BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                        SubItemIndex    =   4
                        Text            =   "Teacher"
                        Object.Width           =   3175
                     EndProperty
                  End
               End
               Begin VB.PictureBox bgYLRec 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00E0E0E0&
                  BorderStyle     =   0  'None
                  ForeColor       =   &H80000008&
                  Height          =   2895
                  Index           =   0
                  Left            =   450
                  ScaleHeight     =   193
                  ScaleMode       =   3  'Pixel
                  ScaleWidth      =   627
                  TabIndex        =   11
                  Top             =   180
                  Visible         =   0   'False
                  Width           =   9405
                  Begin VB.PictureBox bgYLDetail 
                     Appearance      =   0  'Flat
                     BackColor       =   &H00FFFFFF&
                     BorderStyle     =   0  'None
                     ForeColor       =   &H80000008&
                     Height          =   2865
                     Index           =   1
                     Left            =   15
                     ScaleHeight     =   2865
                     ScaleWidth      =   3195
                     TabIndex        =   14
                     Top             =   15
                     Width           =   3195
                     Begin HSES.b8Line b8Line1 
                        Height          =   60
                        Index           =   0
                        Left            =   -105
                        TabIndex        =   20
                        Top             =   2475
                        Width           =   3315
                        _ExtentX        =   5847
                        _ExtentY        =   106
                        BorderColor1    =   12632256
                        BorderColor2    =   16777215
                        BorderColor3    =   16777215
                     End
                     Begin lvButton.lvButtons_H cmdPrint 
                        Height          =   300
                        Index           =   0
                        Left            =   1890
                        TabIndex        =   19
                        Top             =   2535
                        Width           =   1290
                        _ExtentX        =   2275
                        _ExtentY        =   529
                        Caption         =   "Print"
                        CapAlign        =   2
                        BackStyle       =   4
                        BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                           Name            =   "MS Sans Serif"
                           Size            =   8.25
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        cBhover         =   14215660
                        cGradient       =   14215660
                        Gradient        =   3
                        Mode            =   0
                        Value           =   0   'False
                        ImgSize         =   32
                        cBack           =   16185592
                     End
                     Begin VB.Label lblAG 
                        AutoSize        =   -1  'True
                        BackStyle       =   0  'Transparent
                        Caption         =   " "
                        BeginProperty Font 
                           Name            =   "Tahoma"
                           Size            =   8.25
                           Charset         =   0
                           Weight          =   700
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        ForeColor       =   &H00808080&
                        Height          =   195
                        Index           =   0
                        Left            =   285
                        TabIndex        =   39
                        Top             =   2250
                        Width           =   45
                     End
                     Begin VB.Label Label9 
                        Appearance      =   0  'Flat
                        AutoSize        =   -1  'True
                        BackColor       =   &H80000005&
                        BackStyle       =   0  'Transparent
                        Caption         =   "Gen. Average Grade"
                        BeginProperty Font 
                           Name            =   "Tahoma"
                           Size            =   8.25
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        ForeColor       =   &H80000008&
                        Height          =   195
                        Left            =   300
                        TabIndex        =   38
                        Top             =   2070
                        Width           =   1485
                     End
                     Begin VB.Label lblDateEnrolled 
                        AutoSize        =   -1  'True
                        BackStyle       =   0  'Transparent
                        Caption         =   " "
                        BeginProperty Font 
                           Name            =   "Tahoma"
                           Size            =   8.25
                           Charset         =   0
                           Weight          =   700
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        ForeColor       =   &H00808080&
                        Height          =   195
                        Index           =   0
                        Left            =   285
                        TabIndex        =   37
                        Top             =   1830
                        Width           =   45
                     End
                     Begin VB.Label Label8 
                        Appearance      =   0  'Flat
                        AutoSize        =   -1  'True
                        BackColor       =   &H80000005&
                        BackStyle       =   0  'Transparent
                        Caption         =   "Date Enrolled:"
                        BeginProperty Font 
                           Name            =   "Tahoma"
                           Size            =   8.25
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        ForeColor       =   &H80000008&
                        Height          =   195
                        Left            =   300
                        TabIndex        =   36
                        Top             =   1650
                        Width           =   1020
                     End
                     Begin VB.Label Label7 
                        Appearance      =   0  'Flat
                        AutoSize        =   -1  'True
                        BackColor       =   &H80000005&
                        BackStyle       =   0  'Transparent
                        Caption         =   "Adviser:"
                        BeginProperty Font 
                           Name            =   "Tahoma"
                           Size            =   8.25
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        ForeColor       =   &H80000008&
                        Height          =   195
                        Left            =   300
                        TabIndex        =   35
                        Top             =   1245
                        Width           =   600
                     End
                     Begin VB.Label Label6 
                        Appearance      =   0  'Flat
                        AutoSize        =   -1  'True
                        BackColor       =   &H80000005&
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
                        ForeColor       =   &H80000008&
                        Height          =   195
                        Left            =   300
                        TabIndex        =   34
                        Top             =   840
                        Width           =   915
                     End
                     Begin VB.Label Label5 
                        Appearance      =   0  'Flat
                        AutoSize        =   -1  'True
                        BackColor       =   &H80000005&
                        BackStyle       =   0  'Transparent
                        Caption         =   "S.Y.:"
                        BeginProperty Font 
                           Name            =   "Tahoma"
                           Size            =   8.25
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        ForeColor       =   &H80000008&
                        Height          =   195
                        Left            =   300
                        TabIndex        =   33
                        Top             =   405
                        Width           =   360
                     End
                     Begin VB.Label lblAdviser 
                        AutoSize        =   -1  'True
                        BackStyle       =   0  'Transparent
                        Caption         =   " "
                        BeginProperty Font 
                           Name            =   "Tahoma"
                           Size            =   8.25
                           Charset         =   0
                           Weight          =   700
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        ForeColor       =   &H00808080&
                        Height          =   195
                        Index           =   0
                        Left            =   285
                        TabIndex        =   18
                        Top             =   1425
                        Width           =   45
                     End
                     Begin VB.Label lblDepartmentTitle 
                        AutoSize        =   -1  'True
                        BackStyle       =   0  'Transparent
                        Caption         =   " "
                        BeginProperty Font 
                           Name            =   "Tahoma"
                           Size            =   8.25
                           Charset         =   0
                           Weight          =   700
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        ForeColor       =   &H00808080&
                        Height          =   195
                        Index           =   0
                        Left            =   285
                        TabIndex        =   17
                        Top             =   1035
                        Width           =   45
                     End
                     Begin VB.Label lblSchoolYear 
                        AutoSize        =   -1  'True
                        BackStyle       =   0  'Transparent
                        Caption         =   " "
                        BeginProperty Font 
                           Name            =   "Tahoma"
                           Size            =   8.25
                           Charset         =   0
                           Weight          =   700
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        ForeColor       =   &H00808080&
                        Height          =   195
                        Index           =   0
                        Left            =   300
                        TabIndex        =   16
                        Top             =   600
                        Width           =   45
                     End
                     Begin VB.Label lblSectionFullTitle 
                        AutoSize        =   -1  'True
                        BackStyle       =   0  'Transparent
                        Caption         =   "Transferee"
                        BeginProperty Font 
                           Name            =   "Tahoma"
                           Size            =   12
                           Charset         =   0
                           Weight          =   700
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        ForeColor       =   &H00FF0000&
                        Height          =   285
                        Index           =   0
                        Left            =   150
                        TabIndex        =   15
                        Top             =   60
                        Width           =   1335
                     End
                     Begin VB.Line Line1 
                        BorderColor     =   &H00E0E0E0&
                        BorderStyle     =   3  'Dot
                        Index           =   0
                        X1              =   3870
                        X2              =   0
                        Y1              =   360
                        Y2              =   360
                     End
                     Begin VB.Image Image1 
                        Height          =   720
                        Left            =   2445
                        Picture         =   "frmEnrolmentDetail.frx":3C93
                        Top             =   1770
                        Width           =   720
                     End
                  End
                  Begin MSComctlLib.ListView listSubject 
                     Height          =   2865
                     Index           =   0
                     Left            =   3225
                     TabIndex        =   12
                     Top             =   15
                     Width           =   5565
                     _ExtentX        =   9816
                     _ExtentY        =   5054
                     View            =   3
                     LabelEdit       =   1
                     LabelWrap       =   -1  'True
                     HideSelection   =   -1  'True
                     GridLines       =   -1  'True
                     _Version        =   393217
                     Icons           =   "ilSubject"
                     SmallIcons      =   "ilSubject"
                     ForeColor       =   -2147483640
                     BackColor       =   16777215
                     Appearance      =   0
                     NumItems        =   5
                     BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                        Text            =   "Title"
                        Object.Width           =   2646
                     EndProperty
                     BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                        SubItemIndex    =   1
                        Text            =   "Grade"
                        Object.Width           =   1323
                     EndProperty
                     BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                        SubItemIndex    =   2
                        Text            =   "Time"
                        Object.Width           =   1852
                     EndProperty
                     BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                        SubItemIndex    =   3
                        Text            =   "Days"
                        Object.Width           =   1587
                     EndProperty
                     BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                        SubItemIndex    =   4
                        Text            =   "Teacher"
                        Object.Width           =   3175
                     EndProperty
                  End
               End
               Begin VB.Label lblNE 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Student not yet Enroled"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00C0C0C0&
                  Height          =   240
                  Left            =   180
                  TabIndex        =   13
                  Top             =   180
                  Visible         =   0   'False
                  Width           =   2055
               End
            End
            Begin VB.Image imgTop 
               Height          =   165
               Index           =   1
               Left            =   0
               Picture         =   "frmEnrolmentDetail.frx":4B5D
               Stretch         =   -1  'True
               Top             =   60
               Width           =   15360
            End
         End
         Begin HSES.b8SContainer bgTabCon 
            Height          =   4905
            Index           =   2
            Left            =   -74970
            TabIndex        =   28
            Top             =   300
            Width           =   7860
            _ExtentX        =   13864
            _ExtentY        =   8652
            BorderColor     =   12307149
            Begin VB.CommandButton cmdAddCredential 
               BackColor       =   &H00D8E9EC&
               Height          =   315
               Left            =   6990
               Picture         =   "frmEnrolmentDetail.frx":4BFA
               Style           =   1  'Graphical
               TabIndex        =   30
               Top             =   30
               Width           =   315
            End
            Begin MSComctlLib.ImageList ilRecordIco 
               Left            =   0
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
                     Picture         =   "frmEnrolmentDetail.frx":5184
                     Key             =   ""
                  EndProperty
               EndProperty
            End
            Begin MSComctlLib.ListView listCredentials 
               Height          =   3975
               Left            =   60
               TabIndex        =   29
               Top             =   360
               Width           =   7305
               _ExtentX        =   12885
               _ExtentY        =   7011
               View            =   3
               LabelWrap       =   -1  'True
               HideSelection   =   -1  'True
               FullRowSelect   =   -1  'True
               GridLines       =   -1  'True
               _Version        =   393217
               Icons           =   "ilRecordIco"
               SmallIcons      =   "ilRecordIco"
               ForeColor       =   -2147483640
               BackColor       =   -2147483643
               Appearance      =   1
               NumItems        =   4
               BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Text            =   "Credential"
                  Object.Width           =   4233
               EndProperty
               BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   1
                  Text            =   "Remarks"
                  Object.Width           =   5821
               EndProperty
               BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   2
                  Text            =   "Creation Date"
                  Object.Width           =   2646
               EndProperty
               BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   3
                  Text            =   "Created By"
                  Object.Width           =   2646
               EndProperty
            End
            Begin VB.CommandButton cmdDeleteCredential 
               BackColor       =   &H00D8E9EC&
               Height          =   315
               Left            =   6690
               Picture         =   "frmEnrolmentDetail.frx":571E
               Style           =   1  'Graphical
               TabIndex        =   31
               Top             =   30
               Width           =   315
            End
            Begin VB.CommandButton cmdReloadCredential 
               BackColor       =   &H00D8E9EC&
               Height          =   315
               Left            =   6300
               Picture         =   "frmEnrolmentDetail.frx":5CA8
               Style           =   1  'Graphical
               TabIndex        =   32
               Top             =   30
               Width           =   405
            End
         End
      End
      Begin VB.CommandButton cmdGetStudent 
         BackColor       =   &H00D8E9EC&
         Height          =   285
         Left            =   6240
         Picture         =   "frmEnrolmentDetail.frx":6232
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   450
         Width           =   345
      End
      Begin VB.TextBox txtStudentFullName 
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
         Height          =   345
         Left            =   870
         Locked          =   -1  'True
         MaxLength       =   17
         TabIndex        =   3
         Top             =   420
         Width           =   5745
      End
      Begin HSES.b8ChildTitleBar b8Title 
         Height          =   345
         Left            =   15
         TabIndex        =   1
         Top             =   0
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   609
         BackColor       =   14215660
         Caption         =   "Student Record"
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
      Begin HSES.b8Line b8Line1 
         Height          =   60
         Index           =   4
         Left            =   0
         TabIndex        =   22
         Top             =   1950
         Width           =   15360
         _ExtentX        =   27093
         _ExtentY        =   106
      End
      Begin HSES.b8Container bgDetail 
         Height          =   1080
         Left            =   0
         TabIndex        =   23
         Top             =   870
         Width           =   10005
         _ExtentX        =   17648
         _ExtentY        =   1905
         BorderColor     =   12307149
         BackColor       =   16185592
         ShadowColor1    =   13427430
         ShadowColor2    =   14215660
         Begin VB.CheckBox chkDropped 
            BackColor       =   &H00F6F8F8&
            Caption         =   "Dropped"
            Enabled         =   0   'False
            Height          =   285
            Left            =   6315
            TabIndex        =   89
            Top             =   630
            Width           =   2460
         End
         Begin VB.CheckBox chkGraduated 
            BackColor       =   &H00F6F8F8&
            Caption         =   "Graduated"
            Enabled         =   0   'False
            Height          =   285
            Left            =   3315
            TabIndex        =   88
            Top             =   615
            Width           =   2460
         End
         Begin VB.Label Label1 
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
            Left            =   150
            TabIndex        =   27
            Top             =   120
            Width           =   465
         End
         Begin VB.Label lblStudentName 
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
            ForeColor       =   &H00C25418&
            Height          =   345
            Left            =   690
            TabIndex        =   26
            Top             =   60
            Width           =   240
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Level (Latest):"
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
            TabIndex        =   25
            Top             =   630
            Width           =   1050
         End
         Begin VB.Label lblLevel 
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
            ForeColor       =   &H00C25418&
            Height          =   195
            Left            =   1290
            TabIndex        =   24
            Top             =   630
            Width           =   150
         End
      End
      Begin MSComctlLib.ImageList ilSubject 
         Left            =   0
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
               Picture         =   "frmEnrolmentDetail.frx":67BC
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.Image Image4 
         Height          =   105
         Left            =   -90
         Picture         =   "frmEnrolmentDetail.frx":6D56
         Stretch         =   -1  'True
         Top             =   720
         Width           =   30000
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Student:"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   180
         TabIndex        =   4
         Top             =   480
         Width           =   600
      End
      Begin VB.Image Image3 
         Height          =   345
         Left            =   0
         Picture         =   "frmEnrolmentDetail.frx":6DF3
         Stretch         =   -1  'True
         Top             =   360
         Width           =   30000
      End
   End
   Begin HSES.b8Container b8cMain 
      Height          =   3420
      Left            =   390
      TabIndex        =   2
      Top             =   360
      Width           =   7035
      _ExtentX        =   12409
      _ExtentY        =   6033
      BorderColor     =   12632256
      BackColor       =   16777215
      InsideBorderColor=   14215660
      ShadowColor1    =   16777215
      ShadowColor2    =   16777215
   End
End
Attribute VB_Name = "frmStudentRecord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim curStudent As tStudent
Dim curStudentSchoolYearTitle As String
Dim curStudentYearLevelID As Integer

Dim curEnrolmentID(3) As String


Public Function ShowForm(Optional sStudentID As String = "", Optional FormName As String = "")
    
    
    
    On Error Resume Next
    'show form
    mdiMain.MousePointer = vbHourglass
    Me.Show
    Me.SetFocus
    
    DoEvents
    
    If sStudentID <> "" Then
         
        lblStudentName.Caption = ""
        lblLevel.Caption = ""
        chkGraduated.Caption = "Graduated"
        chkGraduated.Value = vbUnchecked
        
        
        If GetStudentByID(sStudentID, curStudent) <> Success Then
            Exit Function
        End If
        
        
        
        
        DoEvents
        tabMAin.Enabled = True
        
        RefreshStudentPersonalInfo
        RefreshRecord
        RefreshCredentials
        
        'refresh parent form
        
        Form_Activate

    End If
    
    
    mdiMain.MousePointer = vbDefault
    

End Function
Public Function IsStudentGraduated(sStudentID As String, ByRef sSYGraduated As String) As Boolean
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    sSQL = " SELECT tblGraduate.StudentID,SchoolYear" & _
            " From tblGraduate" & _
            " WHERE (((tblGraduate.StudentID)='" & sStudentID & "'))"

    'default
    IsStudentGraduated = False
    
    If ConnectRS(HSESDB, vRS, sSQL) = False Then
        'temp
        MsgBox "Unable to connect Recordset.", vbExclamation
        GoTo ReleaseAndExit
    End If
    
    If AnyRecordExisted(vRS) = True Then
        IsStudentGraduated = True
        sSYGraduated = ReadField(vRS.Fields("SchoolYear"))
    Else
        IsStudentGraduated = False
    End If
    
    
ReleaseAndExit:
    Set vRS = Nothing
End Function




Private Sub cmdAddCredential_Click()
    frmStudentCredential.ShowAdd curStudent.StudentID
    RefreshCredentials
End Sub

Private Sub cmdDeleteCredential_Click()
    Dim lvKey As String
    
    
    '-------------------------------------------------------
    'check user access
    '-------------------------------------------------------
    If UserAllowedTo(CurrentUser.UserName, sCanDeleteStudentCredential) = False Then
        MsgBox "Unable to continue deleting Student Credential entry." & vbNewLine & vbNewLine & _
                "You are not permitted to do this. Please contact your administrator for more information.", vbExclamation
        Exit Sub
    End If
    '-------------------------------------------------------

    
    If listCredentials.ListItems.Count < 1 Then
        MsgBox "There is no Student Credential to delete.", vbExclamation
        Exit Sub
    End If
    
    lvKey = GetLVKey(listCredentials.SelectedItem)
    
    If MsgBox("Student Credential will be delete permanently" & vbNewLine & vbNewLine & _
        "Do you want to delete it anaway?", vbQuestion + vbOKCancel) <> vbOK Then
        Exit Sub
    End If
    
    If DeleteStudentCredential(lvKey, curStudent.StudentID) = True Then
        MsgBox "Student Credential deleted.", vbInformation
        RefreshCredentials
    End If
    
End Sub

Private Function DeleteStudentCredential(sCredentialID As String, sStudentID As String) As Boolean

    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    '-------------------------------------------------------
    'check user access
    '-------------------------------------------------------
    If UserAllowedTo(CurrentUser.UserName, sCanDeleteStudentCredential) = False Then
        MsgBox "Unable to continue deleting Student Credential entry." & vbNewLine & vbNewLine & _
                "You are not permitted to do this. Please contact your administrator for more information.", vbExclamation
        GoTo ReleaseAndExit
    End If
    '-------------------------------------------------------

    
    
    sSQL = "DELETE * FROM tblStudentCredential " & _
        " WHERE tblStudentCredential.StudentID='" & sStudentID & "' AND tblStudentCredential.CredentialID='" & sCredentialID & "'"
    
    If ConnectRS(HSESDB, vRS, sSQL) = False Then
        'fatal
        'error
        CatchError "frmStudentRecord", "DeleteStudentCredential", "Unable to connect Recordset with SQL expresstion: '" & sSQL & "'"
        DeleteStudentCredential = False
        GoTo ReleaseAndExit
    End If
    
    DeleteStudentCredential = True
    
ReleaseAndExit:
    Set vRS = Nothing
End Function

Private Sub cmdGetStudent_Click()
    Dim sStudentID As String
    Dim sStudentFullName As String

    
    sStudentID = PickStudent.GetStudentID(txtStudentFullName, sStudentFullName, , False, False)
    
    If sStudentID = "" Then
        
        Exit Sub
    End If
    
    lblStudentName.Caption = ""
    lblLevel.Caption = ""
    chkGraduated.Caption = "Graduated"
    chkGraduated.Value = vbUnchecked
    
    If GetStudentByID(sStudentID, curStudent) <> Success Then
        Exit Sub
    End If
    
    'set name
    lblStudentName.Caption = curStudent.FirstName & " " & curStudent.MiddleName & " " & curStudent.LastName
        
    curStudent.StudentID = sStudentID
    txtStudentFullName.Text = sStudentFullName
    
    DoEvents
    tabMAin.Enabled = True
    
    RefreshStudentPersonalInfo
    RefreshRecord
    RefreshCredentials
    
    'refresh parent form
    
    Form_Activate
End Sub

Private Sub RefreshCredentials()
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    sSQL = " SELECT tblCredential.CredentialID as lvKey, tblCredential.Title, tblStudentCredential.Remarks, tblStudentCredential.CreationDate, tblStudentCredential.CreatedBy" & _
            " FROM tblCredential INNER JOIN tblStudentCredential ON tblCredential.CredentialID = tblStudentCredential.CredentialID" & _
            " WHERE (((tblStudentCredential.StudentID)='" & curStudent.StudentID & "'))" & _
            " ORDER BY tblCredential.Title;"


    If ConnectRS(HSESDB, vRS, sSQL) = False Then
        'temp
        'fatal
        MsgBox "Fatal Error." & vbNewLine & _
            "Detail: Unable to connect Recordset.", vbCritical
        CatchError "frmStudentRecord", "RefreshCredentials", "Uable to connect Recordset with SQL expression '" & sSQL & "'"
        GoTo ReleaseAndExit
    End If
    
    FillRecordToList vRS, listCredentials, "cred", , , , True
    
            
ReleaseAndExit:
    Set vRS = Nothing
End Sub


Private Function RefreshRecord()
    
    Dim i As Integer
    Dim sSYGraduated As String
    
    'set name
        curStudent.StudentID = curStudent.StudentID
        lblStudentName.Caption = curStudent.FirstName & " " & curStudent.MiddleName & " " & curStudent.LastName
        txtStudentFullName.Text = curStudent.FirstName & " " & curStudent.MiddleName & " " & curStudent.LastName
        
        'checck if graduate
        If IsStudentGraduated(curStudent.StudentID, sSYGraduated) = True Then
            chkGraduated.Caption = "Graduated (" & sSYGraduated & ")"
            chkGraduated.Value = vbChecked
        Else
            chkGraduated.Caption = "Graduated"
            chkGraduated.Value = vbUnchecked
        End If
        
        'check if dropped
        If IsStudentDropped(curStudent.StudentID) = Success Then
            chkDropped.Value = vbChecked
        Else
            chkDropped.Value = vbUnchecked
        End If
        
        
        
    
    For i = 0 To bgYLRec.UBound
        bgYLRec(i).Visible = False
    Next
    
    If GetLatestSchoolYearYearLevel(curStudent.StudentID, curStudentSchoolYearTitle, curStudentYearLevelID) <> Success Then
        'temp
        'fatal error
        MsgBox "error"
        
        GoTo ReleaseAndExit
    End If
    
    If curStudentYearLevelID < 1 Then
        tabMAin.Tab = 0
        lblNE.Visible = True
        GoTo ReleaseAndExit
    End If
    
    For i = 1 To curStudentYearLevelID
    
        bgYLRec(i - 1).Visible = True
        DoEvents
        
        ShowStudentYLRecord i
        
        
    Next
    
    bgAllYL.CheckExceedControl
    
    lblLevel = YLIDtoTitle(curStudentYearLevelID) & "  (" & curStudentSchoolYearTitle & ")"

        
ReleaseAndExit:
End Function

Private Function ShowStudentYLRecord(iYearLevelID As Integer)
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    Dim dStudentAveGrade As Double
    Dim bStudentPassed As Boolean
    Dim tmpStudentPrevDepertmentID As String
    
    sSQL = "SELECT tblEnrolment.EnrolmentID, [tblEnrolment]![SchoolYear] AS SY, tblDepartment.DepartmentTitle, [tblYearLevel]![YearLevelTitle] & ' - ' & [tblSection]![SectionTitle] AS SectionFullTitle, [tblTeacher_1]![LastName] & ', ' & [tblTeacher_1]![FirstName] & ' ' & [tblTeacher_1]![MiddleName] AS Adviser, tblEnrolment.CreationDate" & _
            " FROM tblYearLevel INNER JOIN ((tblDepartment INNER JOIN tblSection ON tblDepartment.DepartmentID = tblSection.DepartmentID) INNER JOIN ((tblTeacher AS tblTeacher_1 INNER JOIN tblSectionOffering ON tblTeacher_1.TeacherID = tblSectionOffering.TeacherID) INNER JOIN tblEnrolment ON tblSectionOffering.SectionOfferingID = tblEnrolment.SectionOfferingID) ON tblSection.SectionID = tblSectionOffering.SectionID) ON tblYearLevel.YearLevelID = tblSection.YearLevelID" & _
            " WHERE (((tblEnrolment.StudentID)='" & curStudent.StudentID & "') AND ((tblSection.YearLevelID)=" & iYearLevelID & "));"

    
    If ConnectRS(HSESDB, vRS, sSQL) = False Then
        'temp
        'fatal error
        MsgBox "error"
        
        GoTo ReleaseAndExit
    End If
    
    If AnyRecordExisted(vRS) = False Then
        GoTo ReleaseAndExit
    End If
    
    GetAcademicRecord curStudent.StudentID, iYearLevelID, dStudentAveGrade, bStudentPassed, tmpStudentPrevDepertmentID
    
    
    'set form's field controls
    On Error Resume Next
    
    lblAG(iYearLevelID - 1).Caption = FormatNumber(dStudentAveGrade, 2) & IIf(bStudentPassed = True, " (Passed)", " (Failed)")
    lblDateEnrolled(iYearLevelID - 1).Caption = ReadField(vRS.Fields("CreationDate"))
    lblSectionFullTitle(iYearLevelID - 1).Caption = ReadField(vRS.Fields("SectionFullTitle"))
    lblSchoolYear(iYearLevelID - 1).Caption = ReadField(vRS.Fields("SY"))
    lblDepartmentTitle(iYearLevelID - 1).Caption = ReadField(vRS.Fields("DepartmentTitle"))
    lblAdviser(iYearLevelID - 1).Caption = ReadField(vRS.Fields("Adviser"))
    
    curEnrolmentID(iYearLevelID - 1) = ReadField(vRS.Fields("EnrolmentID"))
    ShowSubjectsByEnrolmentID ReadField(vRS.Fields("EnrolmentID")), iYearLevelID

ReleaseAndExit:
    Set vRS = Nothing
End Function

Private Sub ShowSubjectsByEnrolmentID(sEnrolmentID As String, iYearLevelID As Integer)
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    sSQL = "SELECT tblSubject.SubjectID AS lvKey, tblSubject.SubjectTitle, tblGrade.GradeValue, [tblSubjectOffering]![SchedTimeStart] & '-' & [tblSubjectOffering]![SchedTimeEnd] AS TimeSched, tblSubjectOffering.Days,[tblTeacher]![LastName] & ', ' & [tblTeacher]![FirstName] & ' ' & [tblTeacher]![MiddleName] AS TeacherFullName" & _
            " FROM tblTeacher INNER JOIN (tblSubject INNER JOIN (tblSubjectOffering INNER JOIN (tblEnrolment INNER JOIN tblGrade ON tblEnrolment.EnrolmentID = tblGrade.EnrolmentID) ON tblSubjectOffering.SubjectOfferingID = tblGrade.SubjectOfferingID) ON tblSubject.SubjectID = tblSubjectOffering.SubjectID) ON tblTeacher.TeacherID = tblSubjectOffering.TeacherID" & _
            " Where (((tblEnrolment.EnrolmentID) = '" & sEnrolmentID & "'))" & _
            " ORDER BY tblSubject.SubjectTitle;"
            


    If ConnectRS(HSESDB, vRS, sSQL) = False Then
        'temp
        'fatal error
        MsgBox "error"
        
        GoTo ReleaseAndExit
    End If
    
    FillRecordToList vRS, listSubject(iYearLevelID - 1), KeyGraduate, , 32767, , True
    

ReleaseAndExit:
    
    Set vRS = Nothing
End Sub


Private Sub RefreshStudentPersonalInfo()
    
    Dim i As Integer
    Dim selLength As Integer
    Dim selStart As Integer
    Dim smFound As Boolean
    Dim fn As Boolean
    
    rtbInfo.Text = ""
    
    rtbInfo.Text = rtbInfo.Text & _
    "Name: " & vbTab & vbTab & vbTab & vbTab & txtStudentFullName.Text & vbNewLine

    rtbInfo.Text = rtbInfo.Text & _
    "Gender: " & vbTab & vbTab & vbTab & curStudent.Gender & vbNewLine
    
    rtbInfo.Text = rtbInfo.Text & _
    "Status: " & vbTab & vbTab & vbTab & vbTab & curStudent.Status & vbNewLine
    
    rtbInfo.Text = rtbInfo.Text & _
    "Citizenship: " & vbTab & vbTab & vbTab & curStudent.Citizenship & vbNewLine

    rtbInfo.Text = rtbInfo.Text & _
    "Religion: " & vbTab & vbTab & vbTab & curStudent.Religion & vbNewLine

    rtbInfo.Text = rtbInfo.Text & _
    "Blood Type: " & vbTab & vbTab & vbTab & curStudent.BloodType & vbNewLine
    
    rtbInfo.Text = rtbInfo.Text & _
    "Birth Date: " & vbTab & vbTab & vbTab & curStudent.BirthDate & vbNewLine

    rtbInfo.Text = rtbInfo.Text & _
    "Place Of Birth: " & vbTab & vbTab & vbTab & curStudent.PlaceOfBirth & vbNewLine
    
    rtbInfo.Text = rtbInfo.Text & _
    "City Address: " & vbTab & vbTab & vbTab & curStudent.CityAddress & vbNewLine

    rtbInfo.Text = rtbInfo.Text & _
    "Home Address: " & vbTab & vbTab & vbTab & curStudent.HomeAddress & vbNewLine

    rtbInfo.Text = rtbInfo.Text & _
    vbNewLine
    
    rtbInfo.Text = rtbInfo.Text & _
    "Last School Attended: " & vbTab & vbTab & curStudent.LastSchoolName & vbNewLine

    rtbInfo.Text = rtbInfo.Text & _
    "   Contact Number: " & vbTab & vbTab & curStudent.LastSchoolContactNumber & vbNewLine
    
    rtbInfo.Text = rtbInfo.Text & _
    "   Address: " & vbTab & vbTab & vbTab & curStudent.LastSchoolAddress & vbNewLine
    
    rtbInfo.Text = rtbInfo.Text & _
    vbNewLine
    
    rtbInfo.Text = rtbInfo.Text & _
    "Parents " & vbNewLine
    
    rtbInfo.Text = rtbInfo.Text & _
    "Mother Name: " & vbTab & vbTab & vbTab & curStudent.MotherName & vbNewLine

    rtbInfo.Text = rtbInfo.Text & _
    "   Occupation: " & vbTab & vbTab & vbTab & curStudent.MotherOccupation & vbNewLine

    rtbInfo.Text = rtbInfo.Text & _
    "Father Name: " & vbTab & vbTab & vbTab & curStudent.FatherName & vbNewLine
    
    rtbInfo.Text = rtbInfo.Text & _
    "   Occupation: " & vbTab & vbTab & vbTab & curStudent.FatherOccupation & vbNewLine
    
    rtbInfo.Text = rtbInfo.Text & _
    "   Parents Contact Number: " & vbTab & curStudent.ParentsContactNumber & vbNewLine
    
    rtbInfo.Text = rtbInfo.Text & _
    "   Parents Address: " & vbTab & vbTab & curStudent.ParentsAddress & vbNewLine
    
    rtbInfo.Text = rtbInfo.Text & _
    vbNewLine
    
    rtbInfo.Text = rtbInfo.Text & _
    "Guardian " & vbNewLine
    
    rtbInfo.Text = rtbInfo.Text & _
    "   Name: " & vbTab & vbTab & vbTab & curStudent.GuardianName & vbNewLine
    
    rtbInfo.Text = rtbInfo.Text & _
    "   Contact Number: " & vbTab & vbTab & curStudent.GuardianContactNumber & vbNewLine
    
    rtbInfo.Text = rtbInfo.Text & _
    "   Address: " & vbTab & vbTab & vbTab & curStudent.GuardianAddress & vbNewLine
    
    rtbInfo.Text = rtbInfo.Text & _
    vbNewLine
    
    rtbInfo.Text = rtbInfo.Text & _
    "Entrance Ave. Grade: " & vbTab & vbTab & curStudent.OldAveGrade & vbNewLine
    
    
    'set color
    rtbInfo.selStart = 0
    rtbInfo.selLength = Len(rtbInfo.Text)
    rtbInfo.SelColor = &H808080
    rtbInfo.SelBold = False
    
    For i = 1 To Len(rtbInfo.Text) + 1
    
        If Mid(rtbInfo.Text, i, 1) = ":" Then
            smFound = True
            selStart = i
            selLength = 0
        End If
        
        If smFound = True Then
            selLength = selLength + 1

            If Mid(rtbInfo.Text, i, 2) = vbNewLine Then
                
                rtbInfo.selStart = selStart
                rtbInfo.selLength = selLength
                rtbInfo.SelFontSize = 10
                rtbInfo.SelColor = &H0&
                rtbInfo.SelBold = True
                
                If fn = False Then
                    rtbInfo.SelFontSize = 12
                    fn = True
                End If
                
                rtbInfo.selLength = 0
                
                smFound = False
            End If
            
        End If
        

    Next
    
End Sub

Private Sub cmdPrint_Click(Index As Integer)

    frmPrintStudent.ShowStudentCopyByEnrolment curEnrolmentID(Index)

End Sub

Private Sub cmdReloadCredential_Click()
    RefreshCredentials
End Sub

Private Sub Form_Activate()
    mdiMain.RegMDIChild Me
End Sub

Private Sub Form_Resize()
    ReArrangeControls
End Sub
Private Sub ReArrangeControls()
On Error Resume Next
    
    
    Me.ScaleMode = vbPixels
    b8cMain.Move Form_LeftMargin - 3, Form_TopMargin - 3, Me.ScaleWidth - (Form_LeftMargin - 3) * 2, Me.ScaleHeight - (Form_TopMargin - 3) * 2
    
    bgMain.Move Form_LeftMargin, Form_TopMargin, Me.ScaleWidth - Form_LeftMargin * 2, Me.ScaleHeight - Form_TopMargin * 2
    
    b8Title.Move 0, 0, bgMain.Width
    
    bgDetail.Move bgDetail.Left, bgDetail.Top, bgMain.Width - bgDetail.Left * 2
    
    tabMAin.Move 0, tabMAin.Top, bgMain.Width, bgMain.Height - tabMAin.Top
    
    ArrangeBgTabCon tabMAin.Tab
    
End Sub

Private Function ArrangeBgTabCon(Index As Integer)

    Dim i As Integer
    
    On Error Resume Next
    
    bgTabCon(Index).Move 0, bgTabCon(Index).Top, Screen.TwipsPerPixelX * tabMAin.Width, Screen.TwipsPerPixelY * tabMAin.Height - bgTabCon(Index).Top
    
    Select Case Index
        
        Case 0 'student personal info
            imgTop(Index).Move 15, imgTop(Index).Top, bgTabCon(Index).Width - 30
            rtbInfo.Move rtbInfo.Left, rtbInfo.Top, bgTabCon(Index).Width - rtbInfo.Left - 60, bgTabCon(Index).Height - rtbInfo.Top - 60
        
        Case 1 'record
            
            imgTop(Index).Move 15, imgTop(Index).Top, bgTabCon(Index).Width - 30

            bgAllYL.Move 15, bgAllYL.Top, bgTabCon(Index).Width - 30, bgTabCon(Index).Height - bgAllYL.Top - 30
                    
            For i = 0 To bgYLRec.UBound
                bgYLRec(i).Move 0, (i * bgYLRec(0).Height), bgAllYL.Width, bgYLRec(0).Height
            Next
            
            For i = 0 To listSubject.UBound
                listSubject(i).Move listSubject(i).Left, listSubject(i).Top, (bgYLRec(i).Width / Screen.TwipsPerPixelX) - listSubject(i).Left - 2
            Next
        
        Case 2 'credentials
        
            listCredentials.Move 45, listCredentials.Top, bgTabCon(Index).Width - 90, bgTabCon(Index).Height - listCredentials.Top - 45
            
            cmdAddCredential.Move bgTabCon(Index).Width - cmdAddCredential.Width - 30
            cmdDeleteCredential.Move cmdAddCredential.Left - cmdDeleteCredential.Width + 15
            cmdReloadCredential.Move cmdDeleteCredential.Left - cmdReloadCredential.Width + 15
            
           
            
    End Select
    
End Function



Private Sub tabMAin_Click(PreviousTab As Integer)
    ArrangeBgTabCon tabMAin.Tab
    bgTabCon(tabMAin.Tab).Visible = True
    
    'refresh parent form
    Form_Activate
End Sub




















Public Function Form_CanPrint() As Boolean
    
    'default
    Form_CanPrint = False
    
    If Len(curStudent.StudentID) < 1 Then
        Exit Function
    End If
    
    
    Select Case tabMAin.Tab
        Case 0
            Form_CanPrint = True
        Case 1
            Form_CanPrint = False
    End Select
    
End Function

Public Function Form_Print()

    Select Case tabMAin.Tab
    
        Case 0
            frmPrintStudent.ShowStudentAccountDetailByStudent curStudent.StudentID
        
           
    End Select
    
End Function

