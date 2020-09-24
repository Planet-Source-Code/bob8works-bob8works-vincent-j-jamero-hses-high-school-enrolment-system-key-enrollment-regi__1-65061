VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.MDIForm mdiMain 
   Appearance      =   0  'Flat
   AutoShowChildren=   0   'False
   BackColor       =   &H00FFFFFF&
   Caption         =   "HSES"
   ClientHeight    =   6240
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   11850
   Icon            =   "mdiMain.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Timer timerVote 
      Interval        =   10000
      Left            =   5340
      Top             =   2880
   End
   Begin VB.Timer timerFormTab 
      Interval        =   1
      Left            =   4530
      Top             =   1290
   End
   Begin VB.Timer timerWatchCursor 
      Interval        =   1000
      Left            =   3270
      Top             =   1290
   End
   Begin VB.PictureBox tbMain 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00D8E9EC&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   0
      ScaleHeight     =   1095
      ScaleWidth      =   11850
      TabIndex        =   15
      Top             =   0
      Width           =   11850
      Begin VB.Timer timerUpdateDate 
         Interval        =   1000
         Left            =   1260
         Top             =   480
      End
      Begin VB.PictureBox bgTool 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   705
         Left            =   2160
         ScaleHeight     =   705
         ScaleWidth      =   6705
         TabIndex        =   20
         Top             =   0
         Width           =   6705
         Begin HSES.b8ToolButton cmdToolEdit 
            Height          =   615
            Left            =   1800
            TabIndex        =   21
            Top             =   30
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   1085
            Picture         =   "mdiMain.frx":08CA
            BackColor       =   -2147483643
            Caption         =   "Edit"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            FontName        =   "MS Sans Serif"
            FontSize        =   8.25
            ForeColor       =   0
            DisabledPicture =   "mdiMain.frx":11A4
         End
         Begin HSES.b8ToolButton cmdToolAdd 
            Height          =   615
            Left            =   570
            TabIndex        =   22
            Top             =   30
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   1085
            Picture         =   "mdiMain.frx":1A7E
            BackColor       =   -2147483643
            Caption         =   "New"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            FontName        =   "MS Sans Serif"
            FontSize        =   8.25
            ForeColor       =   0
            DisabledPicture =   "mdiMain.frx":2358
         End
         Begin HSES.b8ToolButton cmdToolDelete 
            Height          =   615
            Left            =   2970
            TabIndex        =   23
            Top             =   30
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   1085
            Picture         =   "mdiMain.frx":2C32
            BackColor       =   -2147483643
            Caption         =   "Delete"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            FontName        =   "MS Sans Serif"
            FontSize        =   8.25
            ForeColor       =   0
            DisabledPicture =   "mdiMain.frx":350C
         End
         Begin HSES.b8ToolButton cmdToolReload 
            Height          =   615
            Left            =   4170
            TabIndex        =   24
            Top             =   30
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   1085
            Picture         =   "mdiMain.frx":3DE6
            BackColor       =   -2147483643
            Caption         =   "Reload"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            FontName        =   "MS Sans Serif"
            FontSize        =   8.25
            ForeColor       =   0
            DisabledPicture =   "mdiMain.frx":46C0
         End
         Begin HSES.b8ToolButton cmdToolPrint 
            Height          =   615
            Left            =   5340
            TabIndex        =   25
            Top             =   30
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   1085
            Picture         =   "mdiMain.frx":4F9A
            BackColor       =   -2147483643
            Caption         =   "Print"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            FontName        =   "MS Sans Serif"
            FontSize        =   8.25
            ForeColor       =   0
            DisabledPicture =   "mdiMain.frx":5874
         End
         Begin VB.Line Line3 
            BorderColor     =   &H00E0E0E0&
            X1              =   30
            X2              =   21700
            Y1              =   0
            Y2              =   0
         End
         Begin VB.Line Line2 
            BorderColor     =   &H00E0E0E0&
            X1              =   330
            X2              =   22000
            Y1              =   690
            Y2              =   690
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00E0E0E0&
            X1              =   330
            X2              =   30
            Y1              =   720
            Y2              =   0
         End
         Begin VB.Image Image2 
            Height          =   720
            Left            =   0
            Picture         =   "mdiMain.frx":614E
            Top             =   0
            Width           =   615
         End
         Begin VB.Image Image4 
            Height          =   735
            Left            =   0
            Picture         =   "mdiMain.frx":78D0
            Stretch         =   -1  'True
            Top             =   0
            Width           =   19995
         End
      End
      Begin VB.PictureBox bgTabBack 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   435
         Left            =   3150
         ScaleHeight     =   435
         ScaleWidth      =   19995
         TabIndex        =   16
         Top             =   720
         Width           =   20000
         Begin VB.PictureBox bgTab 
            Appearance      =   0  'Flat
            BackColor       =   &H00D8E9EC&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   435
            Left            =   15
            ScaleHeight     =   435
            ScaleWidth      =   22005
            TabIndex        =   17
            Top             =   15
            Width           =   22000
            Begin lvButton.lvButtons_H cmdOpenForms 
               Height          =   420
               Index           =   0
               Left            =   0
               TabIndex        =   36
               Top             =   0
               Visible         =   0   'False
               Width           =   3420
               _ExtentX        =   6033
               _ExtentY        =   741
               Caption         =   "Quick Launch"
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
               cFore           =   8421504
               cFHover         =   8421504
               Focus           =   0   'False
               cGradient       =   16777215
               Mode            =   0
               Value           =   0   'False
               cBack           =   16777215
            End
         End
      End
      Begin VB.Label lblDate 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "---"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   195
         Left            =   390
         TabIndex        =   27
         Top             =   630
         Width           =   180
      End
      Begin VB.Label lblCurSchoolYear 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "User Name"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   195
         Left            =   390
         TabIndex        =   26
         Top             =   420
         Width           =   780
      End
      Begin VB.Image Image1 
         Height          =   225
         Left            =   90
         Picture         =   "mdiMain.frx":79D2
         Top             =   870
         Width           =   2970
      End
      Begin VB.Label lblCurrentUserName 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "User Name"
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
         Left            =   315
         TabIndex        =   18
         Top             =   180
         Width           =   915
      End
      Begin VB.Image Image3 
         Height          =   720
         Left            =   0
         Picture         =   "mdiMain.frx":7BCA
         Top             =   90
         Width           =   720
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   3  'Align Left
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5145
      Left            =   3150
      ScaleHeight     =   343
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   1
      TabIndex        =   14
      Top             =   1095
      Width           =   15
   End
   Begin VB.Timer timerMonChild 
      Interval        =   1
      Left            =   3690
      Top             =   1290
   End
   Begin VB.Timer timerWritePreLogOut 
      Interval        =   10000
      Left            =   4110
      Top             =   1290
   End
   Begin MSComctlLib.ImageList imgListOption 
      Left            =   4980
      Top             =   2460
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
            Picture         =   "mdiMain.frx":809E
            Key             =   "ListOption"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":8638
            Key             =   "ChangeFont"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ilMainDisable 
      Left            =   5550
      Top             =   2370
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ilMainNormal 
      Left            =   6240
      Top             =   2430
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":8BD2
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":91DF
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgTBList 
      Left            =   3510
      Top             =   2370
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":97D2
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":A0AC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":A986
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":B260
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":BB3A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":C414
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":CCEE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin HSES.b8SideBar SideBar 
      Align           =   3  'Align Left
      Height          =   5145
      Left            =   0
      TabIndex        =   0
      Top             =   1095
      Width           =   3150
      _ExtentX        =   5556
      _ExtentY        =   9075
      BackColor       =   14215660
      BackColor       =   14215660
      BorderColor1    =   14215660
      BorderColor2    =   14215660
      BorderColor3    =   14215660
      BorderColor4    =   14215660
      BorderColor5    =   14215660
      Begin HSES.b8SideTab b8tListOption 
         Height          =   2700
         Index           =   0
         Left            =   90
         TabIndex        =   2
         Top             =   0
         Width           =   2970
         _ExtentX        =   5239
         _ExtentY        =   4763
         Caption         =   "Info"
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
         ForeColor       =   16442835
         MaxHeight       =   2805
         BorderColor     =   13724971
         AutoExpand      =   0   'False
         Begin VB.Image imgInfoIcon 
            Height          =   645
            Left            =   90
            Top             =   390
            Width           =   615
         End
         Begin VB.Label lblFormDescription 
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
            ForeColor       =   &H00802C22&
            Height          =   645
            Left            =   960
            TabIndex        =   4
            Top             =   405
            Width           =   1965
         End
         Begin VB.Label lblAFTip 
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
            ForeColor       =   &H000000FB&
            Height          =   2100
            Left            =   60
            TabIndex        =   3
            Top             =   1110
            Width           =   2865
            WordWrap        =   -1  'True
         End
      End
      Begin HSES.b8SideTab b8tListOption 
         Height          =   5640
         Index           =   3
         Left            =   90
         TabIndex        =   1
         Top             =   1200
         Width           =   2970
         _ExtentX        =   5239
         _ExtentY        =   9948
         Caption         =   "Quick Launch"
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
         ForeColor       =   16442835
         MaxHeight       =   4760
         BorderColor     =   13724971
         Begin MSComctlLib.ImageList imgQuickLaunch 
            Left            =   1320
            Top             =   2910
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   -2147483643
            ImageWidth      =   32
            ImageHeight     =   32
            MaskColor       =   12632256
            _Version        =   393216
            BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
               NumListImages   =   9
               BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "mdiMain.frx":D5C8
                  Key             =   "department"
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "mdiMain.frx":DEA2
                  Key             =   "schoolyear"
               EndProperty
               BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "mdiMain.frx":E77C
                  Key             =   "yearlevel"
               EndProperty
               BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "mdiMain.frx":F056
                  Key             =   "sectionoffering"
               EndProperty
               BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "mdiMain.frx":F930
                  Key             =   "section"
               EndProperty
               BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "mdiMain.frx":1020A
                  Key             =   "teacher"
               EndProperty
               BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "mdiMain.frx":10AE4
                  Key             =   "student"
               EndProperty
               BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "mdiMain.frx":113BE
                  Key             =   "enrolment"
               EndProperty
               BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "mdiMain.frx":11C98
                  Key             =   "subject"
               EndProperty
            EndProperty
         End
         Begin MSComctlLib.ListView listQuickLaunch 
            Height          =   4365
            Left            =   30
            TabIndex        =   37
            Top             =   630
            Width           =   2895
            _ExtentX        =   5106
            _ExtentY        =   7699
            LabelEdit       =   1
            LabelWrap       =   0   'False
            HideSelection   =   -1  'True
            _Version        =   393217
            Icons           =   "imgQuickLaunch"
            SmallIcons      =   "imgQuickLaunch"
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            Appearance      =   0
            MousePointer    =   99
            MouseIcon       =   "mdiMain.frx":12B72
            NumItems        =   0
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Select a task..."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   210
            TabIndex        =   19
            Top             =   390
            Width           =   1710
         End
      End
      Begin HSES.b8SideTab b8tListOption 
         Height          =   345
         Index           =   2
         Left            =   90
         TabIndex        =   10
         Top             =   660
         Width           =   2970
         _ExtentX        =   5239
         _ExtentY        =   609
         Caption         =   "Filter List"
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
         ForeColor       =   16442835
         MaxHeight       =   2810
         BorderColor     =   13724971
         Begin HSES.b8Line b8Line1 
            Height          =   60
            Left            =   60
            TabIndex        =   35
            Top             =   2310
            Width           =   2805
            _ExtentX        =   4948
            _ExtentY        =   106
         End
         Begin VB.PictureBox bgSideFilter 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   1935
            Left            =   30
            ScaleHeight     =   1935
            ScaleWidth      =   2910
            TabIndex        =   28
            Top             =   360
            Width           =   2910
            Begin VB.ComboBox cmbSideFIlter 
               Height          =   315
               Left            =   75
               TabIndex        =   31
               Top             =   975
               Width           =   2790
            End
            Begin VB.TextBox txtSideFilter 
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   60
               TabIndex        =   30
               Top             =   330
               Width           =   2805
            End
            Begin lvButton.lvButtons_H cmdSideFilter 
               Height          =   375
               Left            =   1620
               TabIndex        =   29
               Top             =   1470
               Width           =   1140
               _ExtentX        =   2011
               _ExtentY        =   661
               Caption         =   "Filter"
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
               cGradient       =   14215660
               Gradient        =   3
               Mode            =   0
               Value           =   0   'False
               cBack           =   16185592
            End
            Begin lvButton.lvButtons_H cmdSideReload 
               Height          =   375
               Left            =   150
               TabIndex        =   32
               Top             =   1470
               Width           =   1155
               _ExtentX        =   2037
               _ExtentY        =   661
               Caption         =   "Reload All"
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
               cGradient       =   14215660
               Gradient        =   3
               Mode            =   0
               Value           =   0   'False
               cBack           =   16185592
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Look In:"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H0030A0B8&
               Height          =   195
               Left            =   90
               TabIndex        =   34
               Top             =   735
               Width           =   585
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Find What:"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H0030A0B8&
               Height          =   195
               Left            =   75
               TabIndex        =   33
               Top             =   120
               Width           =   795
            End
         End
         Begin lvButton.lvButtons_H cmdSideAdvanceFilter 
            Height          =   375
            Left            =   1650
            TabIndex        =   11
            Top             =   2460
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   661
            Caption         =   "More"
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
            cGradient       =   14215660
            Gradient        =   3
            Mode            =   0
            Value           =   0   'False
            cBack           =   16185592
         End
      End
      Begin HSES.b8SideTab b8tListOption 
         Height          =   360
         Index           =   4
         Left            =   90
         TabIndex        =   12
         Top             =   1005
         Width           =   2970
         _ExtentX        =   5239
         _ExtentY        =   635
         Caption         =   "Record Explorer"
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
         ForeColor       =   16442835
         MaxHeight       =   5200
         BorderColor     =   13724971
         Begin HSES.HSESDataFolder RecordTree 
            Height          =   5115
            Left            =   30
            TabIndex        =   13
            Top             =   360
            Width           =   2895
            _ExtentX        =   5106
            _ExtentY        =   9022
         End
      End
      Begin HSES.b8SideTab b8tListOption 
         Height          =   360
         Index           =   1
         Left            =   90
         TabIndex        =   5
         Top             =   330
         Width           =   2970
         _ExtentX        =   5239
         _ExtentY        =   635
         Caption         =   "Find List Item"
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
         ForeColor       =   16442835
         MaxHeight       =   2105
         BorderColor     =   13724971
         Begin VB.TextBox txtSideFind 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   420
            TabIndex        =   8
            Top             =   780
            Width           =   2445
         End
         Begin lvButton.lvButtons_H cmdSideFind 
            Height          =   345
            Left            =   1920
            TabIndex        =   7
            Top             =   1290
            Width           =   945
            _ExtentX        =   1667
            _ExtentY        =   609
            Caption         =   "Find"
            CapAlign        =   2
            BackStyle       =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cGradient       =   0
            Mode            =   0
            Value           =   0   'False
            Image           =   "mdiMain.frx":1344C
            cBack           =   -2147483633
         End
         Begin lvButton.lvButtons_H cmdSideFindNext 
            Height          =   345
            Left            =   720
            TabIndex        =   6
            Top             =   1290
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   609
            Caption         =   "Find Next"
            CapAlign        =   2
            BackStyle       =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cGradient       =   0
            Mode            =   0
            Value           =   0   'False
            Image           =   "mdiMain.frx":139E6
            cBack           =   -2147483633
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Find:"
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
            Left            =   30
            TabIndex        =   9
            Top             =   825
            Width           =   360
         End
         Begin VB.Image Image25 
            Height          =   1770
            Left            =   30
            Picture         =   "mdiMain.frx":13F80
            Top             =   330
            Width           =   1740
         End
      End
      Begin VB.Image imgSideBarBottom 
         Height          =   2970
         Left            =   0
         Picture         =   "mdiMain.frx":15520
         Top             =   4080
         Width           =   3150
      End
   End
   Begin MSComctlLib.ImageList ilMainHot 
      Left            =   4440
      Top             =   2340
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":18DAB
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":19501
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuUsers 
      Caption         =   "&Users"
      Begin VB.Menu mnuAddNewUser 
         Caption         =   "&Add New Entry"
      End
      Begin VB.Menu mnuViewAllUser 
         Caption         =   "&Manage Entries"
      End
      Begin VB.Menu mnuSeparator 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLockHSES 
         Caption         =   "Lock HSES"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Action"
      Begin VB.Menu mnuFormMenu 
         Caption         =   ""
         Index           =   0
         Visible         =   0   'False
      End
      Begin VB.Menu mnuSeparatorEdit3 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuKeyAdd 
         Caption         =   "&Add New Entry"
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnuKeyEdit 
         Caption         =   "&Modify Entry"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuKeyDelete 
         Caption         =   "&Delete Entry"
         Shortcut        =   {F4}
      End
      Begin VB.Menu mnuSeparatorEdit1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFindListItem 
         Caption         =   "Find Item"
         Shortcut        =   ^G
      End
      Begin VB.Menu mnuFilter 
         Caption         =   "Filter Records"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuSeparatorEdit2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuQuickLaunch 
         Caption         =   "Quick Launch"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuRecords 
      Caption         =   "&Records"
      Begin VB.Menu mnuRecordExplorer 
         Caption         =   "R&ecord Explorer"
         Shortcut        =   ^E
      End
      Begin VB.Menu mnuStatisticBySchoolYear 
         Caption         =   "&Statistic By School Year"
      End
      Begin VB.Menu mnuRecordSeparator1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSchoolYear 
         Caption         =   "&School Year"
         Begin VB.Menu mnuAddSchoolYear 
            Caption         =   "&Add New Entry"
         End
         Begin VB.Menu mnuDeleteSchoolYear 
            Caption         =   "&Delete Entry"
         End
         Begin VB.Menu mnuViewAllSchoolYear 
            Caption         =   "&View Entries"
         End
         Begin VB.Menu mnuSchoolYearSeparator1 
            Caption         =   "-"
         End
         Begin VB.Menu mnuLockUnlockSchoolYear 
            Caption         =   "&Lock / Unlock Entry"
         End
         Begin VB.Menu mnuSetActiveSchoolYear 
            Caption         =   "&Set Active School Year"
         End
      End
      Begin VB.Menu mnuDepartment 
         Caption         =   "Department"
         Begin VB.Menu mnuAddDepartment 
            Caption         =   "Add New Entry"
         End
         Begin VB.Menu mnuEditDepartment 
            Caption         =   "Edit Entry"
         End
         Begin VB.Menu mnuDeleteDepartment 
            Caption         =   "Delete Entry"
         End
         Begin VB.Menu mnuViewDepartment 
            Caption         =   "View Entries"
         End
      End
      Begin VB.Menu mnuYearLevel 
         Caption         =   "Year Level"
         Begin VB.Menu mnuViewAllYearLevel 
            Caption         =   "View Entries"
         End
      End
      Begin VB.Menu mnuSection 
         Caption         =   "Sections"
         Begin VB.Menu mnuAddSection 
            Caption         =   "Add New Entry"
         End
         Begin VB.Menu mnuDeleteSection 
            Caption         =   "Delete Entry"
         End
         Begin VB.Menu mnuViewAllSection 
            Caption         =   "View Entries"
         End
         Begin VB.Menu mnuSeparatorSection1 
            Caption         =   "-"
         End
         Begin VB.Menu mnuSectionOfferings 
            Caption         =   "Offerings"
            Begin VB.Menu mnuAddSectionOffering 
               Caption         =   "Add"
            End
            Begin VB.Menu mnuViewAllSectionOffering 
               Caption         =   "View Entries"
            End
            Begin VB.Menu mnuSectionOfferingViewByCriteria 
               Caption         =   "View Entries By Criteria"
            End
            Begin VB.Menu mnuVewSectionDetail 
               Caption         =   "Vew Individual Detail"
            End
         End
      End
      Begin VB.Menu mnuFees 
         Caption         =   "Fees"
         Visible         =   0   'False
         Begin VB.Menu mnuManageFees 
            Caption         =   "Manage Fees"
         End
      End
      Begin VB.Menu mnuSubjects 
         Caption         =   "&Subjects"
         Begin VB.Menu mnuAddSubject 
            Caption         =   "&Add New Entry"
         End
         Begin VB.Menu mnuEditSubject 
            Caption         =   "Edit  Entry"
         End
         Begin VB.Menu mnuDeleteSubject 
            Caption         =   "Delete Entry"
         End
         Begin VB.Menu mnuViewAllSubject 
            Caption         =   "View Entries"
         End
      End
      Begin VB.Menu mnuRoom 
         Caption         =   "&Rooms"
         Begin VB.Menu mnuViewAllRoom 
            Caption         =   "View Entries"
         End
      End
      Begin VB.Menu mnuCredentials 
         Caption         =   "Credentials"
         Begin VB.Menu mnuAddCredential 
            Caption         =   "Add Entry"
         End
         Begin VB.Menu mnuEditCredential 
            Caption         =   "Edit Entry"
         End
         Begin VB.Menu mnuViewCredentials 
            Caption         =   "View Entries"
         End
      End
      Begin VB.Menu mnuSpearator4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuStudents 
         Caption         =   "&Students"
         Begin VB.Menu mnuAddNewStudentAccount 
            Caption         =   "Add Entry"
         End
         Begin VB.Menu mnuEditStudent 
            Caption         =   "Edit Entry"
         End
         Begin VB.Menu mnuDeleteStudent 
            Caption         =   "Delete Entry"
         End
         Begin VB.Menu mnuSeparator2 
            Caption         =   "-"
         End
         Begin VB.Menu mnuViewStudentDetail 
            Caption         =   "View Student's Record"
         End
         Begin VB.Menu mnuViewAllStudent 
            Caption         =   "View Entries"
         End
      End
      Begin VB.Menu mnuEnrolment 
         Caption         =   "Enrolment Entries"
         Begin VB.Menu mnuAddEnrolment 
            Caption         =   "Add Entry"
         End
         Begin VB.Menu mnuViewAllEnrolment 
            Caption         =   "View Entries"
         End
         Begin VB.Menu mnuViewEnrolmentDetail 
            Caption         =   "View Entry Details"
         End
         Begin VB.Menu mnuViewEntriesByCriteria 
            Caption         =   "View Entries By Criteria"
         End
      End
      Begin VB.Menu mnuStudentCredentials 
         Caption         =   "Student Credentials"
         Begin VB.Menu mnuAddStudentCredential 
            Caption         =   "Add Entry"
         End
      End
      Begin VB.Menu mnuSeparatorRecord2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGraduates 
         Caption         =   "Graduates"
         Begin VB.Menu mnuAddGraduate 
            Caption         =   "Add Graduate"
         End
         Begin VB.Menu mnuManageGraduates 
            Caption         =   "Manage Graduates"
         End
      End
      Begin VB.Menu mnuDropped 
         Caption         =   "Dropped"
         Begin VB.Menu mnuDroppedStudent 
            Caption         =   "Dropped Student"
         End
         Begin VB.Menu mnuManageDropped 
            Caption         =   "Manage Dropped"
         End
      End
      Begin VB.Menu mnuLeavedStudents 
         Caption         =   "Leaved Students"
         Begin VB.Menu mnuAddLeavedStudents 
            Caption         =   "Add Entry"
         End
         Begin VB.Menu mnuManageLeavedStudentEntries 
            Caption         =   "Manage Leaved Student Entries"
         End
      End
      Begin VB.Menu mnuSeparatorRecord3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCashier 
         Caption         =   "Cashier"
         Visible         =   0   'False
         Begin VB.Menu mnuManageCashier 
            Caption         =   "&Manage Cashier"
         End
      End
      Begin VB.Menu mnuTeacher 
         Caption         =   "Teachers"
         Begin VB.Menu mnuAddTeacher 
            Caption         =   "Add Entry"
         End
         Begin VB.Menu mnuEditTeacher 
            Caption         =   "Edit Entry"
         End
         Begin VB.Menu mnuDeleteTeacher 
            Caption         =   "Delete Eetry"
         End
         Begin VB.Menu mnuViewAllTeacher 
            Caption         =   "View Entries"
         End
         Begin VB.Menu mnuViewTeacherRecord 
            Caption         =   "View Teacher's Record"
         End
      End
      Begin VB.Menu mnuSeparator5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSchoonInfo 
         Caption         =   "Schoon Info"
         Begin VB.Menu mnuModifySchoolInfo 
            Caption         =   "Modify School Info"
         End
      End
   End
   Begin VB.Menu mnuReports 
      Caption         =   "Repo&rts"
      Begin VB.Menu mnuReportsWizard 
         Caption         =   "View All Reports"
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "O&ptions"
      Begin VB.Menu mnuApplicationSettings 
         Caption         =   "Application Settings"
      End
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "&Window"
      WindowList      =   -1  'True
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuAboutHSES 
         Caption         =   "About HSES"
      End
   End
End
Attribute VB_Name = "mdiMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim LastSideTabOnFocus As Integer

Dim defCmdOpenFormsTop As Integer
Dim defCmdOpenFormsLeft As Integer
Dim defCmdOpenFormsWidth As Integer



Dim VisibleTabCount As Integer

Dim fn(100) As String


Private Sub b8tListOption_CaptionMouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Select Case Index
        Case 0 'info
            b8tListOption(Index).HideExpand
        Case 1
            
    End Select
End Sub

Private Sub b8tListOption_CompleteContract(Index As Integer)
    If b8tListOption(0).Top < 0 Then
        SideBar.MoveDownControls (0 - b8tListOption(0).Top)
    End If
    
    Select Case Index
        Case 0 'info
             b8tListOption(Index).ForeColor = b8tListOption(Index).ContractedForeColor
        Case 4
            RecordTree.Release
    End Select
End Sub

Private Sub b8tListOption_CompleteExpand(Index As Integer)
    Select Case Index
        Case 0
            lblFormDescription.Caption = GetActiveFormDescription
        Case 1
            b8tListOption(Index).Enabled = CanAFFindListItem
            If b8tListOption(Index).Enabled = True Then
                txtSideFind.SetFocus
            End If
        Case 4
            RecordTree.Start
    End Select
    
    LastSideTabOnFocus = Index
End Sub

Private Sub b8tListOption_Resize(Index As Integer)
    Dim i As Integer
    Dim iSpaceExceed As Integer
    
    'MsgBox 0
    
    If b8tListOption(Index).AutoExpand = True Then
        For i = 0 To b8tListOption.UBound
            If Index <> i And b8tListOption(i).AutoExpand = True Then
                b8tListOption(i).Expanded = False
            End If
        Next
    End If
    
    If Index > 0 Then
        b8tListOption(Index).Top = b8tListOption(Index - 1).Top + b8tListOption(Index - 1).Height - Screen.TwipsPerPixelY
    End If
    
    
            
        For i = 1 To b8tListOption.UBound
       
            b8tListOption(i).Top = b8tListOption(i - 1).Top + b8tListOption(i - 1).Height - Screen.TwipsPerPixelY
        Next

        iSpaceExceed = (b8tListOption(Index).Top + b8tListOption(Index).Height) - SideBar.Height

        If iSpaceExceed > 0 Then
            If iSpaceExceed - b8tListOption(Index).Top > 0 Then
                iSpaceExceed = b8tListOption(Index).Top
            End If

            SideBar.MoveUpControls iSpaceExceed
        End If

    SideBar.CheckExceedControl
End Sub







Private Sub cmdOpenForms_Click(Index As Integer)
    On Error Resume Next
    Dim frm As Form
    
    

    For Each frm In Forms
        If frm.Name = fn(Index) Then
            If Me.ActiveForm.Name <> fn(Index) Then
            frm.SetFocus
            End If
            Exit For
        End If
    Next
    
    RefreshTabButtonFace Index
End Sub


Private Function RefreshTabButtonFace(Index As Integer)
    Dim i As Integer
    
    cmdOpenForms(Index).ButtonStyle = lv_Flat
    Set cmdOpenForms(Index).Picture = Me.ActiveForm.Icon
    cmdOpenForms(Index).BackColor = &HFFFFFF
    cmdOpenForms(Index).CaptionAlign = vbLeftJustify
    cmdOpenForms(Index).FontStyle = lv_Bold
    
    
    For i = 0 To cmdOpenForms.UBound
        If Index <> i Then
            cmdOpenForms(i).FontStyle = lv_PlainStyle
            cmdOpenForms(i).BackColor = &HD8E9EC
            cmdOpenForms(Index).CaptionAlign = vbLeftJustify
            cmdOpenForms(i).ButtonStyle = lv_hover
        End If
    Next
    
End Function

Private Sub cmdQuickLaunch_Click(Index As Integer)
    Select Case Index
        Case 0 'add student
            Call mnuAddNewStudentAccount_Click
        Case 1 'add enrolment
            Call mnuAddEnrolment_Click
        Case 2 ' view student detail
            mnuViewStudentDetail_Click
        Case 3 'manage sections
            mnuViewAllSection_Click
        Case 4 'view section detail
            Call mnuVewSectionDetail_Click
        Case 5 'manage school year
            Call mnuViewAllSchoolYear_Click
        Case 6 'view all deparment
            Call mnuViewDepartment_Click
        Case 7 'manage teacher
            Call mnuViewAllTeacher_Click
        Case 8 'manage cashier
            Call mnuManageCashier_Click
        Case 9 'manage fees
            Call mnuManageFees_Click
        Case 10 'manage subjects
            Call mnuViewAllSubject_Click
        Case 11
                
            LockApp
                
    
    End Select
End Sub

Public Function LockApp()
    frmLock.ShowForm
End Function

Private Sub cmdSideAdvanceFilter_Click()
    On Error Resume Next
    
    Me.ActiveForm.Form_AdvanceFilter

End Sub

Private Sub cmdSideFilter_Click()
    On Error Resume Next
    

    Me.ActiveForm.Form_Filter txtSideFilter.Text, cmbSideFIlter.Text
    
End Sub

Private Sub cmdSideFind_Click()
    Dim tmpMultiSelect As Boolean
    Dim tmpInverseSelection As Boolean
    

    On Error Resume Next
    
    'trim
    txtSideFind.Text = Trim(txtSideFind.Text)
    
    'check length
    If Len(txtSideFind.Text) < 1 Then
        HLTxt txtSideFind
        Exit Sub
    End If
    
    'set values for searching
    'If chkMultiSelect.Value = vbChecked Then
    '    tmpMultiSelect = True
    'Else
    '    tmpMultiSelect = False
    'End If
    
    'If chkInverseSelection.Value = vbChecked Then
    '    tmpInverseSelection = True
    'Else
    '    tmpInverseSelection = False
    'End If
    
    'execute find
    FindLVItem Me.ActiveForm.listRecord, txtSideFind.Text ', , tmpMultiSelect, tmpInverseSelection

End Sub

Private Sub cmdSideFindNext_Click()
    Dim tmpMultiSelect As Boolean
    Dim tmpInverseSelection As Boolean
    

    On Error Resume Next
    
    'trim
    txtSideFind.Text = Trim(txtSideFind.Text)
    
    'check length
    If Len(txtSideFind.Text) < 1 Then
        HLTxt txtSideFind
        Exit Sub
    End If
    
    'set values for searching
    'If chkMultiSelect.Value = vbChecked Then
    '    tmpMultiSelect = True
    'Else
    '    tmpMultiSelect = False
    'End If
    
    'If chkInverseSelection.Value = vbChecked Then
    '    tmpInverseSelection = True
    'Else
    '    tmpInverseSelection = False
    'End If
    
    'execute find
    FindLVItem Me.ActiveForm.listRecord, txtSideFind.Text, , , , True ', , tmpMultiSelect, tmpInverseSelection

End Sub



Private Sub cmdSideReload_Click()
    Me.ActiveForm.Form_Filter "", cmbSideFIlter.Text

End Sub



Private Sub cmdToolAdd_Click()
    On Error Resume Next
    
    Me.ActiveForm.Form_Add
End Sub

Private Sub cmdToolDelete_Click()
    On Error Resume Next
    
    Me.ActiveForm.Form_Delete
End Sub

Private Sub cmdToolEdit_Click()
    On Error Resume Next
    
    Me.ActiveForm.Form_Edit
End Sub



Private Sub cmdToolPrint_Click()
    On Error Resume Next
    
    Me.ActiveForm.Form_Print
End Sub



Private Sub cmdToolReload_Click()
    On Error Resume Next
    
    Me.ActiveForm.Form_Reload
End Sub

Private Sub Image20_Click()
End Sub


Private Sub listQuickLaunch_DblClick()


    Select Case listQuickLaunch.SelectedItem.Key
    
        Case "AddEnrolment"
        
            frmAddEnrolment.ShowForm
            
        Case "AddStudent"
        
            frmAddStudent.ShowForm
        
        Case "ManageSchoolYear"
            frmAllSchoolYear.ShowFormList
    
        Case "Managedepartment"
            
            frmAllDepartment.ShowFormList
        
        Case "Manageyearlevel"
            
            frmAllYearLevel.ShowFormList
        
        Case "Managesection"
        
            frmAllSection.ShowFormList
        
        
        Case "Managesectionoffering"
        
            frmAllSectionOffering.ShowFormList
        
        Case "Manageteacher", "Teachers"
        
            frmAllTeacher.ShowFormList
    
        
        Case "ManageStudent", "Students"
        
            frmAllStudent.ShowFormList
        
        Case "ManageEnrolment"
        
            frmAllEnrolment.ShowFormList
        
        Case "ManageSubject"
        
            frmAllSubject.ShowFormList
    End Select

End Sub


Private Sub AddQuickLaunchItems()
    
    listQuickLaunch.ListItems.Clear
    
    listQuickLaunch.ListItems.Add _
    , "AddEnrolment", "New Enrolment", "enrolment", "enrolment"

    listQuickLaunch.ListItems.Add _
    , "AddStudent", "New Student", "student", "student"

 
    listQuickLaunch.ListItems.Add _
    , "ManageSchoolYear", "School Year", "schoolyear", "schoolyear"
    
    listQuickLaunch.ListItems.Add _
    , "Managedepartment", "Departments", "department", "department"
    
    listQuickLaunch.ListItems.Add _
    , "Manageyearlevel", "Year Levels", "yearlevel", "yearlevel"
    
    listQuickLaunch.ListItems.Add _
    , "Managesection", "Sections", "section", "section"
    
    listQuickLaunch.ListItems.Add _
    , "Managesectionoffering", "Section Offerings", "sectionoffering", "sectionoffering"
    
    listQuickLaunch.ListItems.Add _
    , "Manageteacher", "Teachers", "teacher", "teacher"



    listQuickLaunch.ListItems.Add _
    , "ManageStudent", "Students", "student", "student"
    
    listQuickLaunch.ListItems.Add _
    , "ManageEnrolment", "Enrolments", "enrolment", "enrolment"
    
    listQuickLaunch.ListItems.Add _
    , "ManageSubject", "Subjects", "subject", "subject"


End Sub

Private Sub MDIForm_Activate()
    On Error Resume Next
    SideBar_Resize
    
    Me.ActiveForm.Form_Refresh
End Sub

Public Function RefreshActiveForm()
    On Error Resume Next
    Me.ActiveForm.Form_Refresh
End Function






















Private Sub MDIForm_Load()
    Dim i As Integer
    
    'defaults
    defCmdOpenFormsTop = cmdOpenForms(0).Top
    defCmdOpenFormsLeft = cmdOpenForms(0).Left
    defCmdOpenFormsWidth = cmdOpenForms(0).Width
    
    
    frmQuickLaunch.Show
    
    'set sidebar
    For i = b8tListOption.UBound To 0 Step -1
        b8tListOption(i).Expanded = False
        b8tListOption(i).ZOrder 0
    Next
    
    
    'set user
    
    lblCurrentUserName.Caption = "Welcome, " & CurrentUser.UserName
    If CurrentSchoolYear.SchoolYearTitle = "0000" Then
        lblCurSchoolYear.Caption = "S.Y.: ---"
    Else
        lblCurSchoolYear.Caption = "S.Y.: " & CurrentSchoolYear.SchoolYearTitle
    End If
    lblDate.Caption = "Today is: " & FormatDateTime(Now, vbGeneralDate)
    
    'menu
    mnuUsers.Visible = IIf(LCase(CurrentUser.UserType) = "administrator", True, False)


    AddQuickLaunchItems
    
End Sub

















Private Sub MDIForm_Resize()

    On Error Resume Next
    
    bgTool.Left = mdiMain.Width - bgTool.Width
    'b8cRecOpt.Width = Me.Width - b8cRecOpt.Left
    'imgToolGrad1.Move 0, imgToolGrad1.Top, Me.Width
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    
    On Error Resume Next

    If UserLogOut(CurrentUser.UserName, Now, True) <> Success Then
        CatchError "mdiMain", "Private Sub MDIForm_Unload(Cancel As Integer)", "Unabled to save logout"
    End If
    
    Dim frm As Form
    
    For Each frm In Forms
        If frm.Name <> Me.Name Then
            Unload frm
        End If
    Next
    
    
    On Error Resume Next
    HSESDB.Close
    Set HSESDB = Nothing
End Sub

Private Sub mnuAbout_Click()
    frmAbout.ShowForm
End Sub

Private Sub mnuAboutHSES_Click()
    frmSplash.ShowAbout
End Sub

Private Sub mnuAddCredential_Click()
    frmCredential.ShowAdd
End Sub

Private Sub mnuAddDepartment_Click()
    frmAddDepartment.ShowForm
End Sub

Private Sub mnuAddEnrolment_Click()
    'show enrolment form
    frmAddEnrolment.ShowForm
End Sub

Private Sub mnuAddGraduate_Click()
    frmAddGraduate.ShowForm
End Sub

Private Sub mnuAddLeavedStudents_Click()
    frmAddLeaved.ShowForm
End Sub

'menu SCHOOL YEAR
'---------------------------------------------------
Private Sub mnuAddSchoolYear_Click()

    On Error Resume Next
    
    If frmAddSchoolYear.ShowForm = True Then
        Me.ActiveForm.Form_Reload
    End If
    
End Sub



Private Sub mnuAddSection_Click()
    frmAddSection.ShowForm
End Sub

Private Sub mnuAddSectionOffering_Click()
    frmAddSectionOffering.ShowForm
End Sub

Private Sub mnuAddStudentCredential_Click()
    frmStudentCredential.ShowAdd
End Sub

Private Sub mnuAddSubject_Click()
    frmAddSubject.ShowForm
End Sub

Private Sub mnuAddTeacher_Click()
    frmAddTeacher.ShowForm
End Sub


Private Sub mnuApplicationSettings_Click()
    frmSetting.ShowForm
End Sub

Private Sub mnuDeleteDepartment_Click()
    Dim sDepartmentTitle As String


    
    sDepartmentTitle = PickDepartment.GetItem
    

    
    If sDepartmentTitle <> "" Then
        frmDeleteDepartment.ShowForm sDepartmentTitle
    End If
End Sub

Private Sub mnuDeleteSection_Click()
    Dim sSectionTitle As String
    Dim vSection As tSection
    
    
    PickSection.GetSectionID , , , , sSectionTitle
        
    If sSectionTitle <> "" Then
        Select Case GetSectionByTitle(sSectionTitle, vSection)
            Case TranDBResult.Success
                
                'ask user
                
                If MsgBox("WARNING: Deleting Section may affect Enrolment Entries." & vbNewLine & "Delete this record anyway?", vbQuestion + vbYesNo) = vbYes Then
                    Select Case DeleteSection(vSection.SectionID)
                        Case TranDBResult.Success
                            MsgBox "Section record deleted.", vbInformation
                            
                        Case Else
                            MsgBox "Deleting Section went failed", vbCritical
                    End Select
                End If
                
            Case TranDBResult.InvalidTitle
                MsgBox "Title not found"
            
            Case Else
                MsgBox "unknown error"
        End Select
    End If
End Sub

Private Sub mnuDeleteStudent_Click()
    Dim sStudentID As String
    
    sStudentID = PickStudent.GetStudentID
    
    If sStudentID <> "" Then
        If MsgBox("Delete this file?", vbQuestion + vbOKCancel) = vbOK Then
            If DeleteStudent(sStudentID) = 1 Then
                MsgBox "Deleted"
            Else
                MsgBox "not deleted"
            End If
        End If
    End If
End Sub

Private Sub mnuDeleteSubject_Click()
    Dim sSubjectTitle As String
    Dim vSubject As tSubject
    
    
    sSubjectTitle = PickSubject.GetSubjectTitle
        
    If sSubjectTitle <> "" Then
        Select Case GetSubjectByTitle(sSubjectTitle, vSubject)
            Case TranDBResult.Success
                
                'ask user
                
                If MsgBox("WARNING: Deleting Subject may affect Enrolment Entries." & vbNewLine & "Delete this record anyway?", vbQuestion + vbYesNo) = vbYes Then
                    Select Case DeleteSubject(vSubject.SubjectID)
                        Case TranDBResult.Success
                            MsgBox "Subject record deleted.", vbInformation
                            
                        Case Else
                            MsgBox "Deleting Subject went failed", vbCritical
                    End Select
                End If
                
            Case TranDBResult.InvalidTitle
                MsgBox "Title not found"
            
            Case Else
                MsgBox "unknown error"
        End Select
    End If
End Sub

Private Sub mnuDeleteTeacher_Click()
    Dim sTeacherTitle As String
    Dim vTeacher As tTeacher
    
    sTeacherTitle = PickTeacher.GetTeacherID
    
    If sTeacherTitle <> "" Then
        If GetTeacherByTitle(sTeacherTitle, vTeacher) = 1 Then
            ExecDeleteTeacher vTeacher.TeacherID
        End If
                
    End If
End Sub







Private Sub mnuDroppedStudent_Click()
    frmAddDropped.ShowForm
End Sub



Private Sub mnuEditDepartment_Click()
    Dim sDepartmentTitle As String
    Dim vDepartment As tDepartment
    Dim GetInfoResult As Integer
    
    sDepartmentTitle = PickDepartment.GetItem
    

    
    If sDepartmentTitle <> "" Then

        If GetDepartmentByTitle(sDepartmentTitle, vDepartment) = Success Then
            frmEditDepartment.ShowEdit vDepartment.DepartmentID
        Else
            MsgBox "Title not Found", vbCritical
        End If
    End If
End Sub

Private Sub mnuDeleteSchoolYear_Click()
    
    On Error Resume Next
    
    If frmDeleteSchoolYear.ShowForm = True Then
        Me.ActiveForm.Form_Reload
    End If
    
End Sub



'---------------------------------------------------
'end school year









Private Sub mnuEditStudent_Click()
    Dim sStudentID As String
    'temp
    sStudentID = PickStudent.GetStudentID
    
    If sStudentID <> "" Then
        'show edit student form
        'temp
        frmEditStudent.ShowEdit sStudentID
    End If
End Sub

Private Sub mnuEditSubject_Click()
    Dim sSubjectTitle As String
    Dim vSubject As tSubject
    
    
    sSubjectTitle = PickSubject.GetSubjectTitle
        
    If sSubjectTitle <> "" Then
        If GetSubjectByTitle(sSubjectTitle, vSubject) = Success Then
            frmEditSubject.ShowEdit (vSubject.SubjectID)
        Else
            MsgBox "Unable to continue this operation." & vbNewLine & "The selected Subject ID not found in record.", vbCritical
        End If
    End If
End Sub

Private Sub mnuEditTeacher_Click()
    Dim sTeacherID As String
    
    sTeacherID = PickTeacher.GetTeacherID
        
    If sTeacherID <> "" Then
        frmEditTeacher.ShowEdit sTeacherID
    End If
    
    
End Sub




'---------------------------------------------------
'end menu year level


















'level 2
'Student > AddNewStudentAccount
Private Sub mnuAddNewStudentAccount_Click()
    frmAddStudent.ShowForm
End Sub






Private Sub mnuFilter_Click()
    b8tListOption(2).HideExpand
End Sub

Private Sub mnuFindListItem_Click()
    b8tListOption(1).HideExpand
End Sub

Private Sub mnuFormMenu_Click(Index As Integer)
    On Error Resume Next
    
    Me.ActiveForm.Form_MenuClick mnuFormMenu(Index).Caption
    
End Sub

Private Sub mnuKeyAdd_Click()
    cmdToolAdd_Click
End Sub

Private Sub mnuKeyDelete_Click()
    cmdToolDelete_Click
End Sub







Private Sub mnuKeyEdit_Click()
    cmdToolEdit_Click
End Sub

Private Sub mnuLockHSES_Click()
    frmLock.ShowForm
End Sub



Private Sub mnuLockUnlockSchoolYear_Click()
    frmSchoolYearLock.ShowForm
End Sub

Private Sub mnuManageCashier_Click()
    frmAllCashier.ShowFormList
End Sub

Private Sub mnuManageDropped_Click()
    frmAllDropped.ShowFormList
End Sub

Private Sub mnuManageFees_Click()
    frmAllFee.ShowForm
End Sub

Private Sub mnuManageGraduates_Click()
    frmAllGraduate.ShowFormList
End Sub

Private Sub mnuModifySchoolInfo_Click()
    frmSchoolAccount.ShowForm
End Sub

Private Sub mnuQuickLaunch_Click()
    b8tListOption(3).HideExpand
End Sub

Private Sub mnuRecordExplorer_Click()

    'show Explorer tab
    b8tListOption(4).HideExpand
End Sub

Private Sub mnuReportsWizard_Click()
    frmReports.ShowForm
End Sub

Private Sub mnuSectionOfferingViewByCriteria_Click()
    frmASSectionOffering.ShowForm
End Sub

Private Sub mnuSetActiveSchoolYear_Click()
    frmCurrentSchoolYear.setSchoolYear
End Sub



Private Sub mnuStatisticBySchoolYear_Click()
    frmSYstat.ShowForm CurrentSchoolYear.SchoolYearTitle
    
End Sub

'level 1
Private Sub mnuUsers_Click()
    'check settings
    'mnuAddNewUser.Enabled = CurrentUser.canAddUser
    'mnuEditAccountInformation.Enabled = CurrentUser.canEditUser
    'mnuEditUserAccessSettings.Enabled = CurrentUser.canEditUser
    'mnuDeleteUser.Enabled = CurrentUser.canDeleteUser
End Sub


'level 2
'Users > Add New User
Private Sub mnuAddNewUser_Click()
    If CurrentUser.UserType = sAdministratortitle Then
        frmAddUser.ShowForm
    Else
        MsgBox "Unable to show Add Users window." & vbNewLine & _
                "You are not permitted to aceess it. Please contact your Administrator.", vbExclamation

    End If
    
End Sub










Private Sub mnuVewSectionDetail_Click()
    frmSectionDetail.ShowForm
End Sub

Private Sub mnuView_Click()

End Sub

Private Sub mnuViewAllEnrolment_Click()
    frmAllEnrolment.ShowFormList "", "", "", "", ""
End Sub

Private Sub mnuViewAllRoom_Click()
    frmAllRoom.ShowFormList
End Sub

Private Sub mnuViewAllSchoolYear_Click()
    frmAllSchoolYear.ShowFormList
End Sub

Private Sub mnuViewAllSection_Click()
    'show section list form
    frmAllSection.ShowFormList
End Sub

Private Sub mnuViewAllSectionOffering_Click()
    frmAllSectionOffering.ShowFormList
End Sub

Private Sub mnuViewAllStudent_Click()
    'show all student account
    frmAllStudent.ShowFormList
End Sub

Private Sub mnuViewAllSubject_Click()
    'show Subject list form
    frmAllSubject.ShowFormList
End Sub

Private Sub mnuViewAllTeacher_Click()
    'show all Teacher account
    frmAllTeacher.ShowFormList
End Sub

'level 2
'Users > View All User
Private Sub mnuViewAllUser_Click()
    frmUserAccount.ShowForm
End Sub


'Level 2
'Year Level
Private Sub mnuViewAllYearLevel_Click()
    frmAllYearLevel.ShowFormList
End Sub

Private Sub mnuViewCredentials_Click()
    frmAllCredentials.ShowFormList 100
End Sub

Private Sub mnuViewDepartment_Click()
    frmAllDepartment.ShowFormList
End Sub

Private Sub Option2_Click()

End Sub

Private Sub mnuViewEnrolmentDetail_Click()
    frmStudentRecord.ShowForm
End Sub

Private Sub mnuViewEntriesByCriteria_Click()
    frmASEnrolment.ShowForm
End Sub

Private Sub mnuViewStudentDetail_Click()
    frmStudentRecord.ShowForm
End Sub

Private Sub mnuViewTeacherRecord_Click()
    frmTeacherRecord.ShowForm
End Sub

Private Sub RecordTree_FolderClick(fNode As MSComctlLib.Node, sRecordType As String)
    Dim splitKey() As String
    Dim sKey() As String
    Dim sText() As String
    
    
    splitKey = Split(fNode.Key, ";")
    
    Select Case splitKey(0)
        Case keySchoolYear
            RecordTree.GetSchoolYearChilds fNode.Text, sKey, sText
            frmREDepartment.ShowForm fNode.Text, sKey, sText
        Case KeyDepartment
            RecordTree.GetDepartmentChilds splitKey(1), splitKey(2), sKey, sText
            frmREYearLevel.ShowForm splitKey(1), splitKey(2), sKey, sText
        Case KeyYearLevel
            RecordTree.GetYearLevelChilds splitKey(1), splitKey(2), splitKey(3), sKey, sText
            frmRESection.ShowForm splitKey(1), splitKey(2), splitKey(3), sKey, sText
        Case KeySectionOffering
            frmSectionDetail.ShowForm splitKey(4)
    End Select
End Sub

Private Sub RecordTree_Started()
    Dim sKey() As String
    Dim sText() As String
    Dim i As Integer
    
    RecordTree.GetSchoolYearList sText
    
    ReDim sKey(UBound(sText))
    
    For i = 0 To UBound(sText)
        sKey(i) = keySchoolYear & ";" & sText(i)
    Next
    
    If b8tListOption(4).Expanded = False Then
        b8tListOption(4).HideExpand
        
    End If
    
    frmRESchoolYear.ShowForm sKey, sText
    
End Sub

Private Sub SideBar_CtlsScroll()
    imgSideBarBottom.Top = SideBar.Height - imgSideBarBottom.Height
End Sub

Private Sub SideBar_Resize()
    Dim iSpaceExceed As Integer
    
    iSpaceExceed = (b8tListOption(LastSideTabOnFocus).Top + b8tListOption(LastSideTabOnFocus).Height) - SideBar.Height

        If iSpaceExceed > 0 Then
            If iSpaceExceed - b8tListOption(LastSideTabOnFocus).Top > 0 Then
                iSpaceExceed = b8tListOption(LastSideTabOnFocus).Top
            End If

            SideBar.MoveUpControls iSpaceExceed
        Else
        
            iSpaceExceed = SideBar.Height - (b8tListOption(b8tListOption.UBound).Top + b8tListOption(b8tListOption.UBound).Height)

            If iSpaceExceed > 0 And b8tListOption(0).Top < 0 Then
                If b8tListOption(0).Top + iSpaceExceed > 0 Then
                    iSpaceExceed = 0 - b8tListOption(0).Top
                End If
                SideBar.MoveDownControls iSpaceExceed
            End If
        End If
        
    On Error Resume Next
    
    imgSideBarBottom.Top = SideBar.Height - imgSideBarBottom.Height
End Sub



Private Sub tbMain_ButtonClick(ByVal Button As MSComctlLib.Button)
    On Error Resume Next

    Select Case Button.Key
        Case "student"
            Me.PopupMenu mnuStudents
    End Select
End Sub
















'ACTIVE FORM
'----------------------------------------------------------

Private Function GetActiveFormDescription() As String
    On Error Resume Next
    
    GetActiveFormDescription = ""
    GetActiveFormDescription = Me.ActiveForm.Form_Description
        
End Function

Private Function GetAFTip() As String
    On Error Resume Next
    
    GetAFTip = ""
    GetAFTip = Me.ActiveForm.Form_Tip
        
End Function

Private Function CanAFFindListItem() As Boolean
    On Error Resume Next
    
    CanAFFindListItem = False
    CanAFFindListItem = Me.ActiveForm.Form_CanFind
End Function

Private Function AForm_CanFilter()
    
    If Form_CanFilter = False And AForm_CanAdvanceFilter = False Then
        b8tListOption(2).Enabled = False
    Else
    
        bgSideFilter.Enabled = Form_CanFilter
        cmdSideFilter.Enabled = Form_CanFilter
        cmdSideReload.Enabled = Form_CanFilter
        cmbSideFIlter.Enabled = Form_CanFilter
        txtSideFilter.Enabled = Form_CanFilter
        
        cmdSideAdvanceFilter.Enabled = AForm_CanAdvanceFilter
    End If
    
End Function

Private Function AForm_CanAdvanceFilter() As Boolean

    AForm_CanAdvanceFilter = False
    
    On Error Resume Next
    
    AForm_CanAdvanceFilter = Me.ActiveForm.Form_CanAdvanceFilter
    
    
End Function
Private Function Form_CanFilter() As Boolean
    Dim FieldList() As String
    Dim oldCMBIndex As Integer
    Dim i As Integer
    
    On Error GoTo errh
    
    Form_CanFilter = False
    
    Form_CanFilter = Me.ActiveForm.Form_CanFilter(FieldList)
    
    If Form_CanFilter = True Then
        oldCMBIndex = cmbSideFIlter.ListIndex
        cmbSideFIlter.Clear
        If UBound(FieldList) > 0 Then
            
            For i = 0 To UBound(FieldList)
                cmbSideFIlter.AddItem FieldList(i)
            Next
            
            If oldCMBIndex >= 0 And oldCMBIndex < cmbSideFIlter.ListCount Then
                cmbSideFIlter.ListIndex = oldCMBIndex
            Else
                cmbSideFIlter.ListIndex = 0
            End If
            
            Form_CanFilter = True
        
            cmdSideAdvanceFilter.Enabled = Me.ActiveForm.Form_CanAdvanceFilter
                
           
        Else
            Form_CanFilter = False
        End If
    End If
    
    Exit Function
    
errh:
    Debug.Print Err.Description
    Resume Next
End Function









Public Function RegMDIChild(ByRef AForm As Form)
    Dim i As Integer

    
On Error Resume Next

    'update button
    
    
    Refresh_FormTabButtons
    

    For i = 0 To cmdOpenForms.UBound
        If fn(i) = AForm.Name Then
            RefreshTabButtonFace i
            Exit For
        End If
    Next
    
    b8tListOption(1).Enabled = CanAFFindListItem
    Call GetACInfo
    
    
    
    Call AForm_CanFilter
    Call AForm_CanAddEntry 'from active form
    Call AForm_CanEditEntry 'from active form
    Call AForm_CanDeleteEntry 'from active form
    Call AForm_Can_Reload 'from active form
    Call AForm_CanShowOption 'from active form
    
    Call AForm_CanShowListOption 'from active form
    
    Call AForm_CanResizeListFont 'from active form
    Call AForm_CanChangeListFont 'from active form
    Call AForm_CanExplore
    Call AForm_CanPrint 'print
    
    Call Aform_SetFormMenu
    
    
    

End Function


Public Function GetACInfo()
    On Error Resume Next
    
    
    Set imgInfoIcon.Picture = Nothing
    Set imgInfoIcon.Picture = Me.ActiveForm.Icon
    
    lblFormDescription.Caption = GetActiveFormDescription
    lblAFTip.Caption = GetAFTip
End Function

Private Function Aform_SetFormMenu()
    
    Dim sMenu() As String
    Dim bHasMenu As Boolean
    Dim i As Integer
    
    On Error Resume Next
    
    'default
    bHasMenu = False
    
    'hide old menus
    For i = 0 To mnuFormMenu.UBound
        mnuFormMenu(i).Visible = False
        Unload mnuFormMenu(i)
    Next
    mnuSeparatorEdit3.Visible = False
    
    
    bHasMenu = Me.ActiveForm.Form_GetMenu(sMenu)
    
    If bHasMenu = False Then
        Exit Function
    End If
    
    mnuSeparatorEdit3.Visible = True
    
    For i = 0 To UBound(sMenu)
        Load mnuFormMenu(i)
        mnuFormMenu(i).Caption = sMenu(i)
        mnuFormMenu(i).Visible = True
    Next
    

End Function




Private Function Refresh_FormTabButtons()
    Static sfn As String
    Dim i As Integer
    Dim X As Integer

    Dim frm As Form
    Dim lv As lvButtons_H
    Dim tLeft As Integer
    Dim tWidth As Integer
    
    On Error Resume Next
    bgTab.Width = mdiMain.Width - bgTabBack.Left

    i = 0
    For Each frm In Forms
        If LCase(Trim(frm.Name)) <> LCase(Trim(mdiMain.Name)) Then
        If frm.MDIChild = True Then
            
            Load cmdOpenForms(i)
            cmdOpenForms(i).Caption = frm.Caption
            cmdOpenForms(i).Visible = True
            fn(i) = frm.Name
            
            i = i + 1
        End If
        End If
    Next
    

    While i <= cmdOpenForms.UBound
        cmdOpenForms(i).Visible = False
        Unload cmdOpenForms(i)
        i = i + 1
    Wend
    
    tLeft = 0
    tWidth = IIf((bgTab.Width / i) < defCmdOpenFormsWidth, (bgTab.Width / i), defCmdOpenFormsWidth)
    
    For i = 0 To cmdOpenForms.UBound
        If cmdOpenForms(i).Visible = True Then
        
            cmdOpenForms(i).Left = tLeft
            cmdOpenForms(i).Width = tWidth
            tLeft = cmdOpenForms(i).Left + cmdOpenForms(i).Width
            
        End If
    Next
    

    If sfn <> Me.ActiveForm.Name Then
        For i = 0 To cmdOpenForms.UBound
        If fn(i) = Me.ActiveForm.Name Then
            RefreshTabButtonFace i
            Exit For
        End If
    Next
    
    End If
    sfn = Me.ActiveForm.Name
End Function



'Record Operations
Public Function AForm_CanAddEntry() As Boolean
    
    On Error Resume Next
    
    'default
    AForm_CanAddEntry = False
    
    AForm_CanAddEntry = Me.ActiveForm.Form_CanAddEntry

    If AForm_CanAddEntry = True Then
        cmdToolAdd.Enabled = True
        mnuKeyAdd.Enabled = True
    Else
        cmdToolAdd.Enabled = False
        mnuKeyAdd.Enabled = False
    End If
End Function

Public Function AForm_CanEditEntry() As Boolean
    
    On Error Resume Next
    
    'default
    AForm_CanEditEntry = False
    
    AForm_CanEditEntry = Me.ActiveForm.Form_CanEditEntry

    If AForm_CanEditEntry = True Then
        cmdToolEdit.Enabled = True
        mnuKeyEdit.Enabled = True
    Else
        cmdToolEdit.Enabled = False
        mnuKeyEdit.Enabled = False
    End If
End Function

Public Function AForm_CanDeleteEntry() As Boolean
    
    
    On Error Resume Next
    
    'default
    AForm_CanDeleteEntry = False
    
    AForm_CanDeleteEntry = Me.ActiveForm.Form_CanDeleteEntry

    If AForm_CanDeleteEntry = True Then
        cmdToolDelete.Enabled = True
        mnuKeyDelete.Enabled = True
    Else
        cmdToolDelete.Enabled = False
        mnuKeyDelete.Enabled = False
    End If
End Function

Public Function AForm_Can_Reload() As Boolean
    
    On Error Resume Next
    
    
    'default
    AForm_Can_Reload = False
    
    AForm_Can_Reload = Me.ActiveForm.Form_Can_Reload

    If AForm_Can_Reload = True Then
        cmdToolReload.Enabled = True
    Else
        cmdToolReload.Enabled = False
    End If
End Function

Public Function AForm_CanShowOption() As Boolean
    
    On Error Resume Next
    
    
    'default
    AForm_CanShowOption = False
    
    AForm_CanShowOption = Me.ActiveForm.Form_CanShowOption

    If AForm_CanShowOption = True Then
        'cmdToolOption.Enabled = True
    Else
        'cmdToolOption.Enabled = False
    End If
End Function

Public Function AForm_CanExplore() As Boolean
    On Error Resume Next
    
    'default
    AForm_CanExplore = False
    
    AForm_CanExplore = Me.ActiveForm.Form_CanExplore
    
    If AForm_CanExplore = True Then
        
    Else
        b8tListOption(4).Expanded = False
    End If
End Function


'---------------------------------------------------------------
'List options
'---------------------------------------------------------------
Public Function AForm_CanShowListOption() As Boolean
    
    On Error Resume Next
    
    'default
    AForm_CanShowListOption = False
    
    AForm_CanShowListOption = Me.ActiveForm.Form_CanShowListOption
    
    
    If AForm_CanShowListOption = True Then
        'tbListOption.Enabled = True
    Else
        'tbListOption.Enabled = False
    End If
    
End Function

Public Function AForm_CanPrint() As Boolean
    On Error Resume Next
    
    'default
    AForm_CanPrint = False
    
    AForm_CanPrint = Me.ActiveForm.Form_CanPrint
    
    
    If AForm_CanPrint = True Then
        
        cmdToolPrint.Enabled = True
    Else
        cmdToolPrint.Enabled = False
    End If
End Function

Public Function AForm_CanResizeListFont() As Boolean

End Function
Public Function AForm_CanChangeListFont() As Boolean

End Function


Private Sub timerFormTab_Timer()
    Refresh_FormTabButtons
End Sub

Private Sub timerMonChild_Timer()
    On Error GoTo ErrShowWelcomeScreen
    Dim s As String
    s = Me.ActiveForm.Name
    
    Exit Sub
    
ErrShowWelcomeScreen:
    frmQuickLaunch.Show
    timerMonChild.Enabled = False
End Sub

Private Sub timerUpdateDate_Timer()
    lblDate.Caption = "Today is: " & FormatDateTime(Now, vbGeneralDate)
End Sub

Private Sub timerVote_Timer()
    'temp for vote
    'frmVote.ShowForm
    timerVote.Enabled = False
End Sub

Private Sub timerWatchCursor_Timer()

    Static ic As Integer
    Static op As POINTAPI
    
    Dim p As POINTAPI
    
    GetCursorPos p
    If (p.X < (op.X + 5) And p.X > op.X - 5) And (p.Y < (op.Y + 5) And p.Y > op.Y - 5) Then
        ic = ic + 1
    Else
        ic = 0
    End If
    
    op.X = p.X
    op.Y = p.Y

    If ic > AppSet_LockTimeOut Then
        ic = 0
        Call LockApp
    End If
End Sub

Private Sub timerWritePreLogOut_Timer()
    If UserLogOut(CurrentUser.UserName, Now, False) <> Success Then
        CatchError "mdiMain", "Private Sub MDIForm_Unload(Cancel As Integer)", "Unabled to save logout"
    End If
End Sub

'----------------------------------------------------------
'end ACTIVE FORM



Private Sub txtSideFilter_KeyPress(KeyAscii As Integer)
    

    If KeyAscii = vbKeyReturn Then
        Call cmdSideFilter_Click
    End If
End Sub

Private Sub txtSideFind_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        Call cmdSideFind_Click
    End If
End Sub

