VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmAddStudent 
   BackColor       =   &H00D8E9EC&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add New Student Account"
   ClientHeight    =   6600
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9240
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAddStudent.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   440
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   616
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WhatsThisHelp   =   -1  'True
   Begin lvButton.lvButtons_H cmdNext 
      Height          =   405
      Left            =   7710
      TabIndex        =   0
      Top             =   6150
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   714
      Caption         =   "&Next"
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
      cGradient       =   14215660
      Gradient        =   3
      Mode            =   0
      Value           =   0   'False
      cBack           =   16777215
   End
   Begin lvButton.lvButtons_H cmdTabButton 
      Height          =   375
      Index           =   0
      Left            =   -30
      TabIndex        =   3
      Top             =   870
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   661
      Caption         =   "[ 1 ] Student ID"
      CapAlign        =   2
      BackStyle       =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFore           =   8399906
      cFHover         =   8399906
      cBhover         =   16185592
      cGradient       =   16185592
      Gradient        =   3
      Mode            =   0
      Value           =   0   'False
      cBack           =   14215660
   End
   Begin lvButton.lvButtons_H cmdTabButton 
      Height          =   375
      Index           =   1
      Left            =   1140
      TabIndex        =   4
      Top             =   870
      Width           =   1545
      _ExtentX        =   2725
      _ExtentY        =   661
      Caption         =   "[ 2 ] Information Part 1"
      CapAlign        =   2
      BackStyle       =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFore           =   8399906
      cFHover         =   8399906
      cBhover         =   16185592
      cGradient       =   16185592
      Gradient        =   3
      Mode            =   0
      Value           =   0   'False
      cBack           =   14215660
   End
   Begin lvButton.lvButtons_H cmdPrevious 
      Height          =   405
      Left            =   6150
      TabIndex        =   9
      Top             =   6150
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   714
      Caption         =   "&Previous"
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
      cBack           =   16777215
   End
   Begin lvButton.lvButtons_H cmdCancel 
      Height          =   405
      Left            =   4380
      TabIndex        =   10
      Top             =   6150
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   714
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
      cGradient       =   14215660
      Gradient        =   3
      Mode            =   0
      Value           =   0   'False
      cBack           =   16777215
   End
   Begin lvButton.lvButtons_H cmdTabButton 
      Height          =   375
      Index           =   4
      Left            =   5670
      TabIndex        =   14
      Top             =   870
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   661
      Caption         =   "[ 5 ] Options"
      CapAlign        =   2
      BackStyle       =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFore           =   8399906
      cFHover         =   8399906
      cBhover         =   16185592
      cGradient       =   16185592
      Gradient        =   3
      Mode            =   0
      Value           =   0   'False
      cBack           =   14215660
   End
   Begin lvButton.lvButtons_H cmdTabButton 
      Height          =   375
      Index           =   3
      Left            =   4500
      TabIndex        =   6
      Top             =   870
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   661
      Caption         =   "[ 4 ] Summary"
      CapAlign        =   2
      BackStyle       =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFore           =   8399906
      cFHover         =   8399906
      cBhover         =   16185592
      cGradient       =   16185592
      Gradient        =   3
      Mode            =   0
      Value           =   0   'False
      cBack           =   14215660
   End
   Begin lvButton.lvButtons_H cmdTabButton 
      Height          =   375
      Index           =   2
      Left            =   2700
      TabIndex        =   5
      Top             =   870
      Width           =   1785
      _ExtentX        =   3149
      _ExtentY        =   661
      Caption         =   "[ 3 ] Information Part 2"
      CapAlign        =   2
      BackStyle       =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFore           =   8399906
      cFHover         =   8399906
      cBhover         =   16185592
      cGradient       =   16185592
      Gradient        =   3
      Mode            =   0
      Value           =   0   'False
      cBack           =   14215660
   End
   Begin HSES.b8Line b8Line5 
      Height          =   60
      Left            =   660
      TabIndex        =   24
      Top             =   510
      Width           =   9240
      _extentx        =   9499
      _extenty        =   106
   End
   Begin HSES.b8Line b8Line6 
      Height          =   60
      Left            =   0
      TabIndex        =   26
      Top             =   6060
      Width           =   9240
      _extentx        =   16298
      _extenty        =   106
   End
   Begin HSES.b8Container conData 
      Height          =   4695
      Index           =   1
      Left            =   105
      TabIndex        =   11
      Top             =   1350
      Width           =   9045
      _extentx        =   15954
      _extenty        =   8281
      backcolor       =   16185592
      Begin VB.ComboBox cmbBloodType 
         Height          =   315
         Left            =   6000
         TabIndex        =   86
         Top             =   1740
         Width           =   2745
      End
      Begin VB.ComboBox cmbReligion 
         Height          =   315
         Left            =   1290
         TabIndex        =   85
         Top             =   1740
         Width           =   2775
      End
      Begin VB.ComboBox cmbTransfereeYL 
         Enabled         =   0   'False
         Height          =   315
         Left            =   5520
         TabIndex        =   82
         Top             =   3960
         Width           =   735
      End
      Begin VB.CheckBox chkTransferee 
         BackColor       =   &H00F6F8F8&
         Caption         =   "Transferee"
         Height          =   255
         Left            =   4170
         TabIndex        =   81
         Top             =   3990
         Width           =   1305
      End
      Begin VB.ComboBox cmbStatus 
         Height          =   315
         Left            =   5970
         TabIndex        =   49
         Top             =   630
         Width           =   2775
      End
      Begin VB.ComboBox cmbGender 
         Height          =   315
         Left            =   5970
         TabIndex        =   48
         Top             =   270
         Width           =   2775
      End
      Begin VB.TextBox txtOldAveGrade 
         Height          =   330
         Left            =   1710
         MaxLength       =   50
         TabIndex        =   35
         Top             =   3990
         Width           =   1305
      End
      Begin VB.TextBox txtCityAddress 
         Height          =   330
         Left            =   1290
         MaxLength       =   50
         TabIndex        =   34
         Top             =   3360
         Width           =   7425
      End
      Begin VB.TextBox txtHomeAddress 
         Height          =   330
         Left            =   1290
         MaxLength       =   50
         TabIndex        =   33
         Top             =   3000
         Width           =   7425
      End
      Begin VB.TextBox txtPlaceOfBirth 
         Height          =   330
         Left            =   1290
         MaxLength       =   50
         TabIndex        =   32
         Top             =   2640
         Width           =   7425
      End
      Begin VB.TextBox txtCitizenship 
         Height          =   330
         Left            =   6000
         MaxLength       =   50
         TabIndex        =   31
         Text            =   "Filipino"
         Top             =   990
         Width           =   2745
      End
      Begin VB.TextBox txtAge 
         Height          =   330
         Left            =   1290
         MaxLength       =   50
         TabIndex        =   30
         Top             =   1350
         Width           =   2745
      End
      Begin VB.TextBox txtLastName 
         Height          =   330
         Left            =   1290
         MaxLength       =   50
         TabIndex        =   29
         Top             =   990
         Width           =   2745
      End
      Begin VB.TextBox txtMiddleName 
         Height          =   330
         Left            =   1290
         MaxLength       =   50
         TabIndex        =   28
         Top             =   630
         Width           =   2745
      End
      Begin VB.TextBox txtFirstName 
         Height          =   330
         Left            =   1290
         MaxLength       =   50
         TabIndex        =   27
         Top             =   270
         Width           =   2745
      End
      Begin HSES.b8Line b8Line3 
         Height          =   60
         Left            =   240
         TabIndex        =   22
         Top             =   2310
         Width           =   8475
         _extentx        =   14949
         _extenty        =   106
         bordercolor1    =   12307149
         borderstyle1    =   3
      End
      Begin MSComCtl2.DTPicker txtBirthDate 
         Height          =   330
         Left            =   6000
         TabIndex        =   1
         Top             =   1365
         Width           =   2745
         _ExtentX        =   4842
         _ExtentY        =   582
         _Version        =   393216
         Format          =   62849025
         CurrentDate     =   36892
         MinDate         =   2
      End
      Begin HSES.b8Line b8Line4 
         Height          =   60
         Left            =   150
         TabIndex        =   23
         Top             =   3750
         Width           =   8805
         _extentx        =   15531
         _extenty        =   106
         bordercolor2    =   16777215
         bordercolor3    =   16777215
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Blood Type"
         Height          =   195
         Left            =   4890
         TabIndex        =   51
         Top             =   1800
         Width           =   795
      End
      Begin VB.Label Religion 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Religion"
         Height          =   195
         Left            =   180
         TabIndex        =   50
         Top             =   1800
         Width           =   555
      End
      Begin VB.Label Label40 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Previous Ave. Grade"
         Height          =   195
         Left            =   180
         TabIndex        =   47
         Top             =   4020
         Width           =   1485
      End
      Begin VB.Label Label39 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "City Address"
         Height          =   195
         Left            =   180
         TabIndex        =   46
         Top             =   3480
         Width           =   915
      End
      Begin VB.Label Label38 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Home Address"
         Height          =   195
         Left            =   180
         TabIndex        =   45
         Top             =   3090
         Width           =   1035
      End
      Begin VB.Label Label37 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Place Of Birth"
         Height          =   195
         Left            =   180
         TabIndex        =   44
         Top             =   2730
         Width           =   975
      End
      Begin VB.Label Label36 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Birth Date"
         Height          =   195
         Left            =   4890
         TabIndex        =   43
         Top             =   1410
         Width           =   720
      End
      Begin VB.Label Label35 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Citizenship"
         Height          =   195
         Left            =   4860
         TabIndex        =   42
         Top             =   1050
         Width           =   765
      End
      Begin VB.Label Label34 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Status"
         Height          =   195
         Left            =   4860
         TabIndex        =   41
         Top             =   720
         Width           =   465
      End
      Begin VB.Label Label33 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Gender"
         Height          =   195
         Left            =   4860
         TabIndex        =   40
         Top             =   330
         Width           =   525
      End
      Begin VB.Label Label32 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Age"
         Height          =   195
         Left            =   180
         TabIndex        =   39
         Top             =   1410
         Width           =   285
      End
      Begin VB.Label Label31 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Last Name"
         Height          =   195
         Left            =   180
         TabIndex        =   38
         Top             =   1050
         Width           =   750
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Middle Name"
         Height          =   195
         Left            =   180
         TabIndex        =   37
         Top             =   690
         Width           =   900
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "First Name"
         Height          =   195
         Left            =   180
         TabIndex        =   36
         Top             =   300
         Width           =   765
      End
   End
   Begin HSES.b8Container conData 
      Height          =   4695
      Index           =   0
      Left            =   90
      TabIndex        =   2
      Top             =   1350
      Width           =   9045
      _extentx        =   15954
      _extenty        =   8281
      backcolor       =   16185592
      Begin VB.TextBox txtStudentID 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   1
         Left            =   1170
         MaxLength       =   7
         TabIndex        =   80
         Top             =   600
         Width           =   1605
      End
      Begin VB.TextBox txtStudentID 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   0
         Left            =   210
         MaxLength       =   4
         TabIndex        =   79
         Top             =   600
         Width           =   885
      End
      Begin lvButton.lvButtons_H cmdCreateCustom 
         Height          =   345
         Left            =   2880
         TabIndex        =   8
         Top             =   570
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   609
         Caption         =   "Create Custom"
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
         cBack           =   -2147483633
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Student ID"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   210
         TabIndex        =   7
         Top             =   270
         Width           =   900
      End
   End
   Begin HSES.b8Container conData 
      Height          =   4695
      Index           =   4
      Left            =   90
      TabIndex        =   15
      Top             =   1350
      Width           =   9045
      _extentx        =   15954
      _extenty        =   8281
      backcolor       =   16185592
      Begin HSES.b8ChildTitleBar b8ChildTitleBar1 
         Height          =   315
         Left            =   120
         TabIndex        =   84
         Top             =   2250
         Width           =   8805
         _extentx        =   15531
         _extenty        =   556
         backcolor       =   14215660
         caption         =   "Pick A Task"
         font            =   "frmAddStudent.frx":058A
         fontbold        =   -1  'True
         fontname        =   "Tahoma"
         fontsize        =   8.25
         forecolor       =   4210752
         closebutton     =   0   'False
      End
      Begin lvButton.lvButtons_H cmdEnrol 
         Height          =   435
         Left            =   900
         TabIndex        =   18
         Top             =   2760
         Width           =   2685
         _ExtentX        =   4736
         _ExtentY        =   767
         Caption         =   "Enrol This Student"
         CapAlign        =   2
         BackStyle       =   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cBhover         =   12640511
         LockHover       =   1
         cGradient       =   14215660
         Gradient        =   3
         Mode            =   0
         Value           =   0   'False
         Image           =   "frmAddStudent.frx":05B2
         cBack           =   16185592
      End
      Begin lvButton.lvButtons_H cmdAddNewEntry 
         Height          =   435
         Left            =   900
         TabIndex        =   19
         Top             =   3300
         Width           =   2685
         _ExtentX        =   4736
         _ExtentY        =   767
         Caption         =   "Add Another New Entry"
         CapAlign        =   2
         BackStyle       =   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cBhover         =   12640511
         LockHover       =   1
         cGradient       =   14215660
         Gradient        =   3
         Mode            =   0
         Value           =   0   'False
         Image           =   "frmAddStudent.frx":0B4C
         cBack           =   16185592
      End
      Begin lvButton.lvButtons_H cmdViewAllStudent 
         Height          =   435
         Left            =   900
         TabIndex        =   20
         Top             =   3900
         Width           =   2685
         _ExtentX        =   4736
         _ExtentY        =   767
         Caption         =   "View All Student List"
         CapAlign        =   2
         BackStyle       =   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cBhover         =   12640511
         LockHover       =   1
         cGradient       =   14215660
         Gradient        =   3
         Mode            =   0
         Value           =   0   'False
         Image           =   "frmAddStudent.frx":10E6
         cBack           =   16185592
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "New Student entry successfull created!!!"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   345
         Left            =   2580
         TabIndex        =   83
         Top             =   720
         Width           =   4245
      End
   End
   Begin HSES.b8Container conData 
      Height          =   4695
      Index           =   3
      Left            =   90
      TabIndex        =   13
      Top             =   1350
      Width           =   9045
      _extentx        =   15954
      _extenty        =   8281
      backcolor       =   16185592
      Begin VB.TextBox txtSummary 
         Appearance      =   0  'Flat
         BackColor       =   &H00F6F8F8&
         BorderStyle     =   0  'None
         Height          =   3915
         Left            =   210
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   21
         Top             =   510
         Width           =   6045
      End
      Begin lvButton.lvButtons_H cmdSave 
         Height          =   375
         Left            =   6540
         TabIndex        =   17
         Top             =   4050
         Width           =   2265
         _ExtentX        =   3995
         _ExtentY        =   661
         Caption         =   "&Save This Entry"
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
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Summary"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   210
         TabIndex        =   16
         Top             =   90
         Width           =   900
      End
   End
   Begin HSES.b8Container conData 
      Height          =   4695
      Index           =   2
      Left            =   90
      TabIndex        =   12
      Top             =   1350
      Width           =   9045
      _extentx        =   15954
      _extenty        =   8281
      backcolor       =   16185592
      Begin VB.Frame Frame3 
         BackColor       =   &H00F6F8F8&
         Caption         =   "School Last Attended"
         ForeColor       =   &H00C25418&
         Height          =   1365
         Left            =   150
         TabIndex        =   72
         Top             =   3180
         Width           =   8715
         Begin VB.TextBox txtLastSchoolContactNumber 
            Height          =   330
            Left            =   1470
            MaxLength       =   50
            TabIndex        =   77
            Top             =   630
            Width           =   2745
         End
         Begin VB.TextBox txtLastSchoolName 
            Height          =   330
            Left            =   1470
            MaxLength       =   50
            TabIndex        =   75
            Top             =   270
            Width           =   2745
         End
         Begin VB.TextBox txtLastSchoolAddress 
            Height          =   675
            Left            =   5850
            MaxLength       =   50
            TabIndex        =   73
            Top             =   240
            Width           =   2715
         End
         Begin VB.Label Label41 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Contact Number"
            Height          =   195
            Left            =   150
            TabIndex        =   78
            Top             =   660
            Width           =   1170
         End
         Begin VB.Label Label27 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Name"
            Height          =   195
            Left            =   150
            TabIndex        =   76
            Top             =   300
            Width           =   405
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Address"
            Height          =   195
            Left            =   4620
            TabIndex        =   74
            Top             =   300
            Width           =   585
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00F6F8F8&
         Caption         =   "Guardian"
         ForeColor       =   &H00C25418&
         Height          =   1875
         Left            =   4590
         TabIndex        =   65
         Top             =   120
         Width           =   4305
         Begin VB.TextBox txtGuardianAddress 
            Height          =   705
            Left            =   1410
            MaxLength       =   70
            TabIndex        =   70
            Top             =   990
            Width           =   2745
         End
         Begin VB.TextBox txtGuardianContactNumber 
            Height          =   330
            Left            =   1410
            MaxLength       =   20
            TabIndex        =   68
            Top             =   630
            Width           =   2745
         End
         Begin VB.TextBox txtGuardianName 
            Height          =   330
            Left            =   1410
            MaxLength       =   50
            TabIndex        =   66
            Top             =   270
            Width           =   2745
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Address"
            Height          =   195
            Left            =   210
            TabIndex        =   71
            Top             =   1020
            Width           =   585
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Contact Number"
            Height          =   195
            Left            =   210
            TabIndex        =   69
            Top             =   660
            Width           =   1170
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Name"
            Height          =   195
            Left            =   210
            TabIndex        =   67
            Top             =   330
            Width           =   405
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00F6F8F8&
         Caption         =   "Parents"
         ForeColor       =   &H00C25418&
         Height          =   2955
         Left            =   180
         TabIndex        =   52
         Top             =   120
         Width           =   4305
         Begin VB.TextBox txtParentsContactNumber 
            Height          =   330
            Left            =   1410
            MaxLength       =   20
            TabIndex        =   63
            Top             =   1710
            Width           =   2745
         End
         Begin VB.TextBox txtParentsAddress 
            Height          =   690
            Left            =   1410
            MaxLength       =   70
            TabIndex        =   61
            Top             =   2070
            Width           =   2745
         End
         Begin VB.TextBox txtFatherOccupation 
            Height          =   330
            Left            =   1410
            MaxLength       =   20
            TabIndex        =   59
            Top             =   1350
            Width           =   2745
         End
         Begin VB.TextBox txtFatherName 
            Height          =   330
            Left            =   1410
            MaxLength       =   50
            TabIndex        =   57
            Top             =   990
            Width           =   2745
         End
         Begin VB.TextBox txtMotherOccupation 
            Height          =   330
            Left            =   1410
            MaxLength       =   50
            TabIndex        =   55
            Top             =   630
            Width           =   2745
         End
         Begin VB.TextBox txtMotherName 
            Height          =   330
            Left            =   1410
            MaxLength       =   50
            TabIndex        =   53
            Top             =   270
            Width           =   2745
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Contact Number"
            Height          =   195
            Left            =   180
            TabIndex        =   64
            Top             =   1770
            Width           =   1170
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Address"
            Height          =   195
            Left            =   180
            TabIndex        =   62
            Top             =   2040
            Width           =   585
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Occupation"
            Height          =   195
            Left            =   180
            TabIndex        =   60
            Top             =   1410
            Width           =   810
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Father's Name"
            Height          =   195
            Left            =   180
            TabIndex        =   58
            Top             =   1050
            Width           =   1035
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Ocuupation"
            Height          =   195
            Left            =   180
            TabIndex        =   56
            Top             =   690
            Width           =   825
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Mothe's Name"
            Height          =   195
            Left            =   180
            TabIndex        =   54
            Top             =   330
            Width           =   1005
         End
      End
   End
   Begin VB.Image Image5 
      Height          =   720
      Left            =   0
      Picture         =   "frmAddStudent.frx":1680
      Top             =   30
      Width           =   720
   End
   Begin VB.Label Label28 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Add Student Entry"
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
      Left            =   840
      TabIndex        =   25
      Top             =   120
      Width           =   2850
   End
   Begin VB.Image Image2 
      Height          =   135
      Left            =   -240
      Picture         =   "frmAddStudent.frx":254A
      Stretch         =   -1  'True
      Top             =   1215
      Width           =   9675
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00F6F8F8&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00BFCED0&
      Height          =   495
      Left            =   0
      Top             =   840
      Width           =   6765
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00F6F8F8&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00BFCED0&
      Height          =   105
      Left            =   3960
      Top             =   1200
      Width           =   5745
   End
   Begin VB.Image Image6 
      Height          =   525
      Left            =   0
      Picture         =   "frmAddStudent.frx":25E7
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9240
   End
End
Attribute VB_Name = "frmAddStudent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Dim RecordSaved As Boolean

Dim CurrentTab As Integer

Dim CurrentStudent As tStudent

Dim conDataSize As RECT

Dim oldGender As Integer
Dim oldStatus As Integer


Dim vfrmUseTitleCase As Boolean

Public Function ShowForm() As Boolean
    Dim i As Integer
    
    '-------------------------------------------------------
    'check user access
    '-------------------------------------------------------
    If UserAllowedTo(CurrentUser.UserName, sCanAddStudent) = False Then
        MsgBox "Unable to continue adding Student entry." & vbNewLine & vbNewLine & _
                "You are not permitted to do this. Please contact your administrator for more information.", vbExclamation
        Exit Function
    End If
    '-------------------------------------------------------


    Me.MousePointer = vbHourglass
    DoEvents

    'set default
    RecordSaved = False
    
    'get tab size
    conDataSize.Top = conData(0).Top
    conDataSize.Left = conData(0).Left
    conDataSize.Right = conData(0).Width
    conDataSize.Bottom = conData(0).Height
    
    
    'set combo boxes
    cmbGender.AddItem "Male"
    cmbGender.AddItem "Female"
    
    cmbStatus.AddItem "Single"
    cmbStatus.AddItem "Married"
    cmbStatus.AddItem "Separated"
    cmbStatus.AddItem "Widowed"
    
    'enable controls
    'disabled controls
    cmdCreateCustom.Enabled = True
    cmdSave.Enabled = True
    cmdNext.Enabled = True
    cmdPrevious.Enabled = True
    For i = 0 To cmdTabButton.UBound
    cmdTabButton(i).Enabled = True
    Next
    
    txtBirthDate.MaxDate = Now - 1
    '415 = 11 years that is assigned to birth date current value which means it started at 11 yrs old
    txtBirthDate.Value = CDate(Now) - 4015
    txtAge.Text = (Now - txtBirthDate.Value) \ 365
    
    'set tabs
    SelectTab 0
    
    'set textboxes
    'Call AutoGenID
    
    'show form
    Me.MousePointer = vbDefault
    Me.Show vbModal
    
    ShowForm = RecordSaved
End Function


Private Function AllowShowTab(Index As Integer) As Boolean
    
    Dim sFullname As String


    'default
    AllowShowTab = False
    
    If Index >= 1 And RecordSaved = False Then
        'check student id
        If Len(CurrentStudent.StudentID) <> 12 Then
            MsgBox "You cannot move to this part unless you have created New Student ID", vbExclamation
            Exit Function
        End If
    End If
    
    
    
    If Index >= 2 And RecordSaved = False Then
        If CheckTextBox(txtFirstName, "Invalid 'FIRST NAME' value. This field is required, please enter a value.") = False Then
            HLTxt txtFirstName
            Exit Function
        End If
        
        If CheckTextBox(txtMiddleName, "Invalid 'Middle Name' value. This field is required, please enter a value.") = False Then
            HLTxt txtMiddleName
            Exit Function
        End If
        
        If CheckTextBox(txtLastName, "Invalid 'LAST NAME' value. This field is required, please enter a value.") = False Then
            HLTxt txtLastName
            Exit Function
        End If
        
        If CheckTextBox(cmbGender, "Invalid 'GENDER' value. This field is required, please enter a value.") = False Then
            HLTxt cmbGender
            Exit Function
        End If
        
        If CheckTextBox(cmbStatus, "Invalid 'STATUS' value. This field is required, please enter a value.") = False Then
            HLTxt cmbStatus
            Exit Function
        End If
        
        If CheckTextBox(txtCitizenship, "Invalid 'Citizenship' value. This field is required, please enter a value.") = False Then
            HLTxt txtCitizenship
            Exit Function
        End If
        
        If CheckTextBox(txtPlaceOfBirth, "Invalid 'PLACE OF BIRTH' value. This field is required, please enter a value.") = False Then
            HLTxt txtPlaceOfBirth
            Exit Function
        End If
        
        If CheckTextBox(txtHomeAddress, "Invalid 'HOME ADDRESS' value. This field is required, please enter a value.") = False Then
            HLTxt txtHomeAddress
            Exit Function
        End If
        
        If CheckTextBox(txtCityAddress, "Invalid 'CITY ADDRESS' value. This field is required, please enter a value.") = False Then
            HLTxt txtCityAddress
            Exit Function
        End If
        
        'check duplicate full name
        sFullname = LCase(Trim(txtFirstName.Text) & Trim(txtMiddleName.Text) & Trim(txtLastName.Text))
        

        If FindDuplicateFullName(sFullname) = Success Then
            MsgBox "Invalid Name value. The Student name the you entered is already existed, please enter another value.", vbExclamation
            HLTxt txtFirstName
            Exit Function
        End If

        
        
        'check old ave grade
        If IsNumeric(txtOldAveGrade.Text) Then
            If Val(txtOldAveGrade.Text) < 75 Or Val(txtOldAveGrade) > 100 Then
                MsgBox "Invalid Old Average Grade!" & vbNewLine & "It mus be Numeric and must be 75 - 100", vbCritical
                HLTxt txtOldAveGrade
                Exit Function
            End If
        Else
            MsgBox "Invalid Old Average Grade!" & vbNewLine & "It mus be Numeric and must be 75 - 100", vbCritical
            HLTxt txtOldAveGrade
            Exit Function
        End If
    
    
    End If
    
    
    
    
    
    
    If Index >= 3 And RecordSaved = False Then
    
        If CheckTextBox(txtMotherName, "Invalid 'MOTHER NAME' value. This field is required, please enter a value.") = False Then
            HLTxt txtMotherName
            Exit Function
        End If
        
        If CheckTextBox(txtMotherOccupation, "Invalid 'MOTHER OCCUPATION' value. This field is required, please enter a value.") = False Then
            HLTxt txtMotherOccupation
            Exit Function
        End If
        
        If CheckTextBox(txtFatherName, "Invalid 'FATHER NAME' value. This field is required, please enter a value.") = False Then
            HLTxt txtFatherName
            Exit Function
        End If
        
        If CheckTextBox(txtFatherOccupation, "Invalid 'FATHER OCCUPATION' value. This field is required, please enter a value.") = False Then
            HLTxt txtFatherOccupation
            Exit Function
        End If
        
        If CheckTextBox(txtParentsAddress, "Invalid 'PARENTS ADDRESS' value. This field is required, please enter a value.") = False Then
            HLTxt txtParentsAddress
            Exit Function
        End If
        
        If CheckTextBox(txtParentsContactNumber, "Invalid 'PARENTS CONTACT NUMBER' value. This field is required, please enter a value.") = False Then
            HLTxt txtParentsContactNumber
            Exit Function
        End If
        
        If CheckTextBox(txtGuardianName, "Invalid 'GUARDIAN NAME' value. This field is required, please enter a value.") = False Then
            HLTxt txtGuardianName
            Exit Function
        End If
        
        If CheckTextBox(txtGuardianContactNumber, "Invalid 'GUARDIAN CONTACT NUMBER' value. This field is required, please enter a value.") = False Then
            HLTxt txtGuardianContactNumber
            Exit Function
        End If
        
        If CheckTextBox(txtGuardianAddress, "Invalid 'GUARDIAN ADDRESS' value. This field is required, please enter a value.") = False Then
            HLTxt txtGuardianAddress
            Exit Function
        End If
        
        If CheckTextBox(txtLastSchoolName, "Invalid 'LAST SCHOOL ATTENDED' value. This field is required, please enter a value.") = False Then
            HLTxt txtLastSchoolName
            Exit Function
        End If
        
        If CheckTextBox(txtLastSchoolContactNumber, "Invalid 'LAST SCHOOL ATTENDED CONTACT NUMBER' value. This field is required, please enter a value.") = False Then
            HLTxt txtLastSchoolContactNumber
            Exit Function
        End If
        
        If CheckTextBox(txtLastSchoolAddress, "Invalid 'LAST SCHOOL ATTENDED ADDRESS' value. This field is required, please enter a value.") = False Then
            HLTxt txtLastSchoolAddress
            Exit Function
        End If
        
        
        'set student
        With CurrentStudent
            'ignore student id, it was set already
            '.StudentID = txtStudentID
            .FirstName = txtFirstName
            .MiddleName = txtMiddleName
            .LastName = txtLastName
            
            .CityAddress = txtCityAddress
            .HomeAddress = txtHomeAddress
            .BirthDate = FormatDateTime(txtBirthDate.Value, vbShortDate)
            .PlaceOfBirth = txtPlaceOfBirth
            .Gender = cmbGender
            .Status = cmbStatus
            .Citizenship = txtCitizenship
            
            .BloodType = cmbBloodType
            .Religion = cmbReligion.Text
            
            
            .LastSchoolName = txtLastSchoolName
            .LastSchoolContactNumber = txtLastSchoolContactNumber
            .LastSchoolAddress = txtLastSchoolAddress
               
            'parents
            .MotherName = txtMotherName
            .MotherOccupation = txtMotherOccupation
            .FatherName = txtFatherName
            .FatherOccupation = txtFatherOccupation
            .ParentsContactNumber = txtParentsContactNumber
            .ParentsAddress = txtParentsAddress
            
            .GuardianName = txtGuardianName
            .GuardianAddress = txtGuardianAddress
            .GuardianContactNumber = txtGuardianContactNumber
            .OldAveGrade = Val(txtOldAveGrade.Text)
            
            .Transferee = chkTransferee.Value
            .TransfereeYL = cmbTransfereeYL.ListIndex + 2
            
            .CreationDate = FormatDateTime(Now, vbShortDate)
            .CreatedBy = CurrentUser.UserName
        
        End With
        
        'show summary
           txtSummary.Text = "Student ID: " & CurrentStudent.StudentID
           txtSummary.Text = txtSummary.Text & vbNewLine & "First Name: " & CurrentStudent.FirstName
           txtSummary.Text = txtSummary.Text & vbNewLine & "Middle Name: " & CurrentStudent.MiddleName
           txtSummary.Text = txtSummary.Text & vbNewLine & "Last Name: " & CurrentStudent.LastName
           txtSummary.Text = txtSummary.Text & vbNewLine & "City Address: " & CurrentStudent.CityAddress
           txtSummary.Text = txtSummary.Text & vbNewLine & "Home Address: " & CurrentStudent.HomeAddress
           txtSummary.Text = txtSummary.Text & vbNewLine & "Birth Date: " & CurrentStudent.BirthDate
           txtSummary.Text = txtSummary.Text & vbNewLine & "Place Of Birth: " & CurrentStudent.PlaceOfBirth
           txtSummary.Text = txtSummary.Text & vbNewLine & "Gender: " & CurrentStudent.Gender
           txtSummary.Text = txtSummary.Text & vbNewLine & "Status: " & CurrentStudent.Status
           txtSummary.Text = txtSummary.Text & vbNewLine & "Citizenship: " & CurrentStudent.Citizenship
            
           txtSummary.Text = txtSummary.Text & vbNewLine & "Last School Attended: " & CurrentStudent.LastSchoolName
           txtSummary.Text = txtSummary.Text & vbNewLine & "Last School Attended Contact Number: " & CurrentStudent.LastSchoolContactNumber
           txtSummary.Text = txtSummary.Text & vbNewLine & "Last School Attended Address: " & CurrentStudent.LastSchoolAddress
               
            'parents
           txtSummary.Text = txtSummary.Text & vbNewLine & "Mother Name: " & CurrentStudent.MotherName
           txtSummary.Text = txtSummary.Text & vbNewLine & "Mother Occupation: " & CurrentStudent.MotherOccupation
           txtSummary.Text = txtSummary.Text & vbNewLine & "Father Name: " & CurrentStudent.FatherName
           txtSummary.Text = txtSummary.Text & vbNewLine & "Father Occupation: " & CurrentStudent.FatherOccupation
           txtSummary.Text = txtSummary.Text & vbNewLine & "Parents COntact Number: " & CurrentStudent.ParentsContactNumber
           txtSummary.Text = txtSummary.Text & vbNewLine & "ParentsAddress: " & CurrentStudent.ParentsAddress
            
           txtSummary.Text = txtSummary.Text & vbNewLine & "Guardian Name: " & CurrentStudent.GuardianName
           txtSummary.Text = txtSummary.Text & vbNewLine & "Guardian Contact Number: " & CurrentStudent.GuardianAddress
           txtSummary.Text = txtSummary.Text & vbNewLine & "Guardian Address: " & CurrentStudent.GuardianContactNumber
            
            txtSummary.Text = txtSummary.Text & vbNewLine & "Creation Date: " & CurrentStudent.CreationDate
    
        
    
    End If
    
    
    
    
    
    
    If Index >= 4 Then
        If RecordSaved = False Then
            MsgBox "Please Save this entry first.", vbExclamation
            Exit Function
        End If
    End If
    
    AllowShowTab = True
End Function


Private Function SelectTab(Index As Integer)
    
    Dim i As Integer
    
    
    If AllowShowTab(Index) = False Then
        Exit Function
    End If
        
    
    CurrentTab = Index

    
    For i = 0 To conData.UBound
        If Index <> i Then
            conData(i).Visible = False
            cmdTabButton(i).GradientColor = &HD8E9EC
        End If
    Next

    cmdTabButton(Index).GradientColor = &HF6F8F8
    conData(Index).Move conDataSize.Left, conDataSize.Top, conDataSize.Right, conDataSize.Bottom
    conData(Index).Visible = True
    
    If CurrentTab < 1 Then
        cmdPrevious.Enabled = False
    Else
        cmdPrevious.Enabled = True
    End If
    
    If CurrentTab >= conData.UBound Then
        cmdNext.Enabled = False
    Else
        cmdNext.Enabled = True
    End If
    
On Error Resume Next
    'set onfocus control
    Select Case Index
        Case 0 'student id
            cmdNext.SetFocus
        Case 1 'info 1
            HLTxt txtFirstName
        Case 2 'info 3
            HLTxt txtMotherName
    End Select
End Function



Private Sub AutoGenID()
    Dim newID As String
    
    newID = Trim(GetNewStudentID)

    If Len(newID) = 12 Then
        CurrentStudent.StudentID = newID
        txtStudentID(0) = Left(CurrentStudent.StudentID, 4)
        txtStudentID(1) = Right(CurrentStudent.StudentID, 7)
    End If
End Sub

Private Sub chkUseSentenceCase_Click()

End Sub

Private Sub chkTransferee_Click()
    cmbTransfereeYL.Enabled = chkTransferee.Value
End Sub

Private Sub cmbGender_GotFocus()
    oldGender = cmbGender.ListIndex
End Sub

Private Sub cmbGender_LostFocus()
    If cmbGender.ListIndex < 0 Then
        cmbGender.ListIndex = oldGender
    End If
End Sub

Private Sub cmbStatus_GotFocus()
    oldStatus = cmbStatus.ListIndex
End Sub

Private Sub cmbStatus_LostFocus()
    If cmbStatus.ListIndex < 0 Then
        cmbStatus.ListIndex = oldStatus
    End If
End Sub

Private Sub cmbTransfereeYL_LostFocus()
    If cmbTransfereeYL.ListIndex < 0 Then
        cmbTransfereeYL.ListIndex = 0
    End If
End Sub

Private Sub cmdAddNewEntry_Click()
    'unload this form
    Unload Me
    
    'reset and show this form
    frmAddStudent.ShowForm
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdCreateCustom_Click()
    Dim newID As String
    
    newID = frmStudentID.CreateNewStudentID
    DoEvents
    
    'MsgBox newID
    If Len(Trim(newID)) = 12 Then
        CurrentStudent.StudentID = newID
        txtStudentID(0) = Left(CurrentStudent.StudentID, 4)
        txtStudentID(1) = Right(CurrentStudent.StudentID, 7)
    End If
    
End Sub

Private Sub cmdEnrol_Click()
    'unload this form
    Unload Me
    'show add enrolment
    frmAddEnrolment.ShowForm CurrentStudent.StudentID
    
End Sub

Private Sub cmdNext_Click()
    SelectTab CurrentTab + 1
End Sub

Private Sub cmdPrevious_Click()
    SelectTab CurrentTab - 1
End Sub

Private Sub cmdSave_Click()
    
    Dim i  As Integer
    
    'check if entry is already saved
        If RecordSaved = False Then
            'save record
            Select Case AddStudent(CurrentStudent)
                Case TranDBResult.Success
                    RecordSaved = True
                    txtSummary.Text = "(New Entry Saved. CLick Next To Proceed."
                    
                    MsgBox "New Student entry successfully saved to the record.", vbInformation
                    
                    'goto option
                    SelectTab (4)
                    
                    'disabled controls
                    cmdCreateCustom.Enabled = False
                    cmdSave.Enabled = False
                    cmdNext.Enabled = False
                    cmdPrevious.Enabled = False
                    For i = 0 To cmdTabButton.UBound
                    cmdTabButton(i).Enabled = False
                    Next
                    

                Case Else ' failed
                    'fatal
                    'temp
                    MsgBox "Error! Unabled to saved entry.", vbExclamation
                    txtSummary.Text = "Error! Unabled to saved entry."
            End Select
        End If
End Sub

Private Sub cmdTabButton_Click(Index As Integer)
    SelectTab Index
End Sub

Private Sub cmdViewAllStudent_Click()
    'close this form
    Unload Me
    'show all student list
    frmAllStudent.ShowFormList
    
End Sub

Private Sub Form_Activate()
    'set textboxes
    Call AutoGenID
End Sub

Private Sub Form_Load()

    'defaults
    
    'title case
    vfrmUseTitleCase = True
    
    cmbTransfereeYL.Clear
    'cmbTransfereeYL.AddItem "I"
    cmbTransfereeYL.AddItem "II"
    cmbTransfereeYL.AddItem "III"
    cmbTransfereeYL.AddItem "IV"
    cmbTransfereeYL.ListIndex = 0
    
    'set religion
    cmbReligion.Clear
    'default
    cmbReligion.AddItem "Catholic"
    cmbReligion.AddItem "Islam"
    cmbReligion.AddItem "Budhist"
    cmbReligion.ListIndex = 0
    
    cmbBloodType.Clear
    cmbBloodType.AddItem "O"
    cmbBloodType.AddItem "A"
    cmbBloodType.AddItem "AB"
    cmbBloodType.ListIndex = 0
    
End Sub



Private Sub Form_Unload(Cancel As Integer)
    'clear temp new id
    DeleteUserStudentID CurrentUser.UserName
End Sub

Private Sub txtAge_GotFocus()
    txtBirthDate.SetFocus
End Sub

Private Sub txtBirthDate_Change()
On Error Resume Next
    txtAge.Text = (Now - txtBirthDate.Value) \ 365
End Sub







Private Sub txtCitizenship_LostFocus()
    If vfrmUseTitleCase = True Then
        txtCitizenship.Text = cSentenceCase(txtCitizenship.Text)
    End If
End Sub

Private Sub txtFirstName_LostFocus()
    If vfrmUseTitleCase = True Then
        txtFirstName.Text = cSentenceCase(txtFirstName.Text)
    End If
End Sub

Private Sub txtMiddleName_LostFocus()
    If vfrmUseTitleCase = True Then
        txtMiddleName.Text = cSentenceCase(txtMiddleName.Text)
    End If
End Sub
Private Sub txtLastName_LostFocus()
    If vfrmUseTitleCase = True Then
        txtLastName.Text = cSentenceCase(txtLastName.Text)
    End If
End Sub


Private Sub txtOldAveGrade_LostFocus()
    If IsNumeric(txtOldAveGrade.Text) Then
        If Val(txtOldAveGrade.Text) < 75 Or Val(txtOldAveGrade) > 100 Then
            MsgBox "Invalid Old Average Grade!" & vbNewLine & "It mus be Numeric and must be 75 - 100", vbCritical
            HLTxt txtOldAveGrade
        End If
    Else
        MsgBox "Invalid Old Average Grade!" & vbNewLine & "It mus be Numeric and must be 75 - 100", vbCritical
        HLTxt txtOldAveGrade
    End If
End Sub

Private Sub txtPlaceOfBirth_LostFocus()
    If vfrmUseTitleCase = True Then
        txtPlaceOfBirth.Text = cSentenceCase(txtPlaceOfBirth.Text)
    End If
End Sub
Private Sub txtHomeAddress_LostFocus()
    If vfrmUseTitleCase = True Then
        txtHomeAddress.Text = cSentenceCase(txtHomeAddress.Text)
    End If
End Sub
Private Sub txtCityAddress_LostFocus()
    If vfrmUseTitleCase = True Then
        txtCityAddress.Text = cSentenceCase(txtCityAddress.Text)
    End If
End Sub



'info part 2 LOST FOCUS subroutines
        Private Sub txtMotherName_LostFocus()
            If vfrmUseTitleCase = True Then
            txtMotherName.Text = cSentenceCase(txtMotherName)
            End If
        End Sub
        
        Private Sub txtMotherOccupation_LostFocus()
            If vfrmUseTitleCase = True Then
            txtMotherOccupation.Text = cSentenceCase(txtMotherOccupation)
            End If
        End Sub
        
        Private Sub txtFatherName_LostFocus()
            If vfrmUseTitleCase = True Then
            txtFatherName.Text = cSentenceCase(txtFatherName)
            End If
        End Sub
        
        Private Sub txtFatherOccupation_LostFocus()
            If vfrmUseTitleCase = True Then
            txtFatherOccupation.Text = cSentenceCase(txtFatherOccupation)
            End If
        End Sub
        
        Private Sub txtParentsAddress_LostFocus()
            If vfrmUseTitleCase = True Then
            txtParentsAddress.Text = cSentenceCase(txtParentsAddress)
            End If
        End Sub
        
        Private Sub txtParentsContactNumber_LostFocus()
            If vfrmUseTitleCase = True Then
            txtParentsContactNumber.Text = cSentenceCase(txtParentsContactNumber)
            End If
        End Sub
        
        Private Sub txtGuardianName_LostFocus()
            If vfrmUseTitleCase = True Then
            txtGuardianName.Text = cSentenceCase(txtGuardianName)
            End If
        End Sub
        
        Private Sub txtGuardianContactNumber_LostFocus()
            If vfrmUseTitleCase = True Then
            txtGuardianContactNumber.Text = cSentenceCase(txtGuardianContactNumber)
            End If
        End Sub
        
        Private Sub txtGuardianAddress_LostFocus()
            If vfrmUseTitleCase = True Then
            txtGuardianAddress.Text = cSentenceCase(txtGuardianAddress)
            End If
        End Sub
        
        Private Sub txtLastSchoolName_LostFocus()
            If vfrmUseTitleCase = True Then
            txtLastSchoolName.Text = cSentenceCase(txtLastSchoolName)
            End If
        End Sub
        
        Private Sub txtLastSchoolContactNumber_LostFocus()
            If vfrmUseTitleCase = True Then
            txtLastSchoolContactNumber.Text = cSentenceCase(txtLastSchoolContactNumber)
            End If
        End Sub
        
        Private Sub txtLastSchoolAddress_LostFocus()
            If vfrmUseTitleCase = True Then
            txtLastSchoolAddress.Text = cSentenceCase(txtLastSchoolAddress)
            End If
        End Sub
'end info part 2 sub routines



Private Sub txtStudentID_KeyPress(Index As Integer, KeyAscii As Integer)
    If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8 Or KeyAscii = 45) Then KeyAscii = 0
End Sub
