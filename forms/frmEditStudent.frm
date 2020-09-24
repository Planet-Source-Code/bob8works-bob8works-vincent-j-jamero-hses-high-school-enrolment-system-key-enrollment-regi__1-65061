VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmEditStudent 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00D8E9EC&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Edit Student Account"
   ClientHeight    =   6540
   ClientLeft      =   150
   ClientTop       =   480
   ClientWidth     =   9210
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmEditStudent.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   436
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   614
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin lvButton.lvButtons_H cmdTabButton 
      Height          =   375
      Index           =   0
      Left            =   -30
      TabIndex        =   4
      Top             =   780
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
      TabIndex        =   5
      Top             =   780
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
      cBhover         =   16185592
      cGradient       =   16185592
      Gradient        =   3
      Mode            =   0
      Value           =   0   'False
      cBack           =   14215660
   End
   Begin lvButton.lvButtons_H cmdTabButton 
      Height          =   375
      Index           =   4
      Left            =   5670
      TabIndex        =   6
      Top             =   780
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
      TabIndex        =   7
      Top             =   780
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
      TabIndex        =   8
      Top             =   780
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
      cBhover         =   16185592
      cGradient       =   16185592
      Gradient        =   3
      Mode            =   0
      Value           =   0   'False
      cBack           =   14215660
   End
   Begin HSES.b8Container conData 
      Height          =   4695
      Index           =   3
      Left            =   90
      TabIndex        =   9
      Top             =   1260
      Width           =   9045
      _ExtentX        =   15954
      _ExtentY        =   8281
      BackColor       =   16185592
      Begin lvButton.lvButtons_H cmdSave 
         Height          =   375
         Left            =   6600
         TabIndex        =   10
         Top             =   270
         Width           =   2265
         _ExtentX        =   3995
         _ExtentY        =   661
         Caption         =   "&Update Entry"
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
         cBack           =   -2147483633
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
         TabIndex        =   13
         Top             =   90
         Width           =   900
      End
      Begin VB.Label lblsummaryinfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Record Not Saved. (Click [Save This Entry] button)"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   1500
         TabIndex        =   11
         Top             =   120
         Width           =   3660
      End
      Begin VB.Label lblSummary 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4215
         Left            =   330
         TabIndex        =   12
         Top             =   360
         Width           =   5955
      End
   End
   Begin HSES.b8Container conData 
      Height          =   4695
      Index           =   2
      Left            =   90
      TabIndex        =   47
      Top             =   1290
      Width           =   9045
      _ExtentX        =   15954
      _ExtentY        =   8281
      BackColor       =   16185592
      Begin VB.Frame Frame1 
         BackColor       =   &H00F6F8F8&
         Caption         =   "Parents"
         ForeColor       =   &H00C25418&
         Height          =   2955
         Left            =   180
         TabIndex        =   62
         Top             =   120
         Width           =   4305
         Begin VB.TextBox txtMotherName 
            Height          =   330
            Left            =   1410
            MaxLength       =   50
            TabIndex        =   68
            Top             =   270
            Width           =   2745
         End
         Begin VB.TextBox txtMotherOccupation 
            Height          =   330
            Left            =   1410
            MaxLength       =   50
            TabIndex        =   67
            Top             =   630
            Width           =   2745
         End
         Begin VB.TextBox txtFatherName 
            Height          =   330
            Left            =   1410
            MaxLength       =   50
            TabIndex        =   66
            Top             =   990
            Width           =   2745
         End
         Begin VB.TextBox txtFatherOccupation 
            Height          =   330
            Left            =   1410
            MaxLength       =   20
            TabIndex        =   65
            Top             =   1350
            Width           =   2745
         End
         Begin VB.TextBox txtParentsAddress 
            Height          =   690
            Left            =   1410
            MaxLength       =   70
            TabIndex        =   64
            Top             =   2070
            Width           =   2745
         End
         Begin VB.TextBox txtParentsContactNumber 
            Height          =   330
            Left            =   1410
            MaxLength       =   20
            TabIndex        =   63
            Top             =   1710
            Width           =   2745
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Mothe's Name"
            Height          =   195
            Left            =   180
            TabIndex        =   74
            Top             =   330
            Width           =   1005
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Ocuupation"
            Height          =   195
            Left            =   180
            TabIndex        =   73
            Top             =   690
            Width           =   825
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Father's Name"
            Height          =   195
            Left            =   180
            TabIndex        =   72
            Top             =   1050
            Width           =   1035
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Occupation"
            Height          =   195
            Left            =   180
            TabIndex        =   71
            Top             =   1410
            Width           =   810
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Address"
            Height          =   195
            Left            =   180
            TabIndex        =   70
            Top             =   2040
            Width           =   585
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Contact Number"
            Height          =   195
            Left            =   180
            TabIndex        =   69
            Top             =   1770
            Width           =   1170
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00F6F8F8&
         Caption         =   "Guardian"
         ForeColor       =   &H00C25418&
         Height          =   1875
         Left            =   4590
         TabIndex        =   55
         Top             =   120
         Width           =   4305
         Begin VB.TextBox txtGuardianName 
            Height          =   330
            Left            =   1410
            MaxLength       =   50
            TabIndex        =   58
            Top             =   270
            Width           =   2745
         End
         Begin VB.TextBox txtGuardianContactNumber 
            Height          =   330
            Left            =   1410
            MaxLength       =   20
            TabIndex        =   57
            Top             =   630
            Width           =   2745
         End
         Begin VB.TextBox txtGuardianAddress 
            Height          =   705
            Left            =   1410
            MaxLength       =   70
            TabIndex        =   56
            Top             =   990
            Width           =   2745
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Name"
            Height          =   195
            Left            =   210
            TabIndex        =   61
            Top             =   330
            Width           =   405
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Contact Number"
            Height          =   195
            Left            =   210
            TabIndex        =   60
            Top             =   660
            Width           =   1170
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Address"
            Height          =   195
            Left            =   210
            TabIndex        =   59
            Top             =   1020
            Width           =   585
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00F6F8F8&
         Caption         =   "School Last Attended"
         ForeColor       =   &H00C25418&
         Height          =   1365
         Left            =   150
         TabIndex        =   48
         Top             =   3180
         Width           =   8715
         Begin VB.TextBox txtLastSchoolAddress 
            Height          =   675
            Left            =   5850
            MaxLength       =   50
            TabIndex        =   51
            Top             =   240
            Width           =   2715
         End
         Begin VB.TextBox txtLastSchoolName 
            Height          =   330
            Left            =   1470
            MaxLength       =   50
            TabIndex        =   50
            Top             =   270
            Width           =   2745
         End
         Begin VB.TextBox txtLastSchoolContactNumber 
            Height          =   330
            Left            =   1470
            MaxLength       =   50
            TabIndex        =   49
            Top             =   630
            Width           =   2745
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Address"
            Height          =   195
            Left            =   4620
            TabIndex        =   54
            Top             =   300
            Width           =   585
         End
         Begin VB.Label Label27 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Name"
            Height          =   195
            Left            =   150
            TabIndex        =   53
            Top             =   300
            Width           =   405
         End
         Begin VB.Label Label41 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Contact Number"
            Height          =   195
            Left            =   150
            TabIndex        =   52
            Top             =   660
            Width           =   1170
         End
      End
   End
   Begin HSES.b8Container conData 
      Height          =   4695
      Index           =   1
      Left            =   90
      TabIndex        =   16
      Top             =   1260
      Width           =   9045
      _ExtentX        =   15954
      _ExtentY        =   8281
      BackColor       =   16185592
      Begin VB.TextBox txtFirstName 
         Height          =   330
         Left            =   1290
         MaxLength       =   50
         TabIndex        =   29
         Top             =   270
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
      Begin VB.TextBox txtLastName 
         Height          =   330
         Left            =   1290
         MaxLength       =   50
         TabIndex        =   27
         Top             =   990
         Width           =   2745
      End
      Begin VB.TextBox txtAge 
         Height          =   330
         Left            =   1290
         MaxLength       =   50
         TabIndex        =   26
         Top             =   1350
         Width           =   2745
      End
      Begin VB.TextBox txtCitizenship 
         Height          =   330
         Left            =   6000
         MaxLength       =   50
         TabIndex        =   25
         Text            =   "Filipino"
         Top             =   990
         Width           =   2745
      End
      Begin VB.TextBox txtPlaceOfBirth 
         Height          =   330
         Left            =   1290
         MaxLength       =   50
         TabIndex        =   24
         Top             =   2910
         Width           =   7425
      End
      Begin VB.TextBox txtHomeAddress 
         Height          =   330
         Left            =   1290
         MaxLength       =   50
         TabIndex        =   23
         Top             =   3270
         Width           =   7425
      End
      Begin VB.TextBox txtCityAddress 
         Height          =   330
         Left            =   1290
         MaxLength       =   50
         TabIndex        =   22
         Top             =   3630
         Width           =   7425
      End
      Begin VB.TextBox txtOldAveGrade 
         Height          =   330
         Left            =   1860
         MaxLength       =   50
         TabIndex        =   21
         Top             =   4230
         Width           =   1305
      End
      Begin VB.ComboBox cmbGender 
         Height          =   315
         Left            =   5970
         TabIndex        =   20
         Top             =   270
         Width           =   2775
      End
      Begin VB.ComboBox cmbStatus 
         Height          =   315
         Left            =   5970
         TabIndex        =   19
         Top             =   630
         Width           =   2775
      End
      Begin VB.TextBox txtReligion 
         Height          =   330
         Left            =   1290
         MaxLength       =   50
         TabIndex        =   18
         Top             =   1710
         Width           =   2745
      End
      Begin VB.TextBox txtBloodType 
         Height          =   330
         Left            =   6000
         MaxLength       =   50
         TabIndex        =   17
         Top             =   1710
         Width           =   2745
      End
      Begin HSES.b8Line b8Line3 
         Height          =   60
         Left            =   240
         TabIndex        =   30
         Top             =   2580
         Width           =   8475
         _ExtentX        =   14949
         _ExtentY        =   106
         BorderColor1    =   12307149
         BorderStyle1    =   3
      End
      Begin MSComCtl2.DTPicker txtBirthDate 
         Height          =   330
         Left            =   6000
         TabIndex        =   31
         Top             =   1350
         Width           =   2745
         _ExtentX        =   4842
         _ExtentY        =   582
         _Version        =   393216
         Format          =   61734913
         CurrentDate     =   36892
         MinDate         =   2
      End
      Begin HSES.b8Line b8Line4 
         Height          =   60
         Left            =   150
         TabIndex        =   32
         Top             =   4020
         Width           =   8805
         _ExtentX        =   15531
         _ExtentY        =   106
         BorderColor2    =   16777215
         BorderColor3    =   16777215
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "First Name"
         Height          =   195
         Left            =   180
         TabIndex        =   46
         Top             =   300
         Width           =   765
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Middle Name"
         Height          =   195
         Left            =   180
         TabIndex        =   45
         Top             =   690
         Width           =   900
      End
      Begin VB.Label Label31 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Last Name"
         Height          =   195
         Left            =   180
         TabIndex        =   44
         Top             =   1050
         Width           =   750
      End
      Begin VB.Label Label32 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Age"
         Height          =   195
         Left            =   180
         TabIndex        =   43
         Top             =   1410
         Width           =   285
      End
      Begin VB.Label Label33 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Gender"
         Height          =   195
         Left            =   4860
         TabIndex        =   42
         Top             =   330
         Width           =   525
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
      Begin VB.Label Label35 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Citizenship"
         Height          =   195
         Left            =   4860
         TabIndex        =   40
         Top             =   1050
         Width           =   765
      End
      Begin VB.Label Label36 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Birth Date"
         Height          =   195
         Left            =   4890
         TabIndex        =   39
         Top             =   1410
         Width           =   720
      End
      Begin VB.Label Label37 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Place Of Birth"
         Height          =   195
         Left            =   180
         TabIndex        =   38
         Top             =   3000
         Width           =   975
      End
      Begin VB.Label Label38 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Home Address"
         Height          =   195
         Left            =   180
         TabIndex        =   37
         Top             =   3360
         Width           =   1035
      End
      Begin VB.Label Label39 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "City Address"
         Height          =   195
         Left            =   180
         TabIndex        =   36
         Top             =   3750
         Width           =   915
      End
      Begin VB.Label Label40 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Previous Ave. Grade"
         Height          =   195
         Left            =   330
         TabIndex        =   35
         Top             =   4260
         Width           =   1485
      End
      Begin VB.Label Religion 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Religion"
         Height          =   195
         Left            =   180
         TabIndex        =   34
         Top             =   1800
         Width           =   555
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Blood Type"
         Height          =   195
         Left            =   4890
         TabIndex        =   33
         Top             =   1800
         Width           =   795
      End
   End
   Begin lvButton.lvButtons_H cmdNext 
      Height          =   405
      Left            =   7680
      TabIndex        =   77
      Top             =   6090
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
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdPrevious 
      Height          =   405
      Left            =   6120
      TabIndex        =   78
      Top             =   6090
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
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdCancel 
      Height          =   405
      Left            =   4350
      TabIndex        =   79
      Top             =   6090
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
      cBack           =   -2147483633
   End
   Begin HSES.b8Line b8Line6 
      Height          =   60
      Left            =   0
      TabIndex        =   80
      Top             =   6000
      Width           =   9240
      _ExtentX        =   16298
      _ExtentY        =   106
   End
   Begin HSES.b8Line b8Line5 
      Height          =   60
      Left            =   660
      TabIndex        =   81
      Top             =   480
      Width           =   9240
      _ExtentX        =   9499
      _ExtentY        =   106
   End
   Begin HSES.b8Container conData 
      Height          =   4695
      Index           =   0
      Left            =   90
      TabIndex        =   14
      Top             =   1260
      Width           =   9045
      _ExtentX        =   15954
      _ExtentY        =   8281
      BackColor       =   16185592
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
         Left            =   2850
         MaxLength       =   50
         TabIndex        =   76
         Top             =   1830
         Width           =   885
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
         Index           =   1
         Left            =   3810
         MaxLength       =   50
         TabIndex        =   75
         Top             =   1830
         Width           =   1605
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Student ID"
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
         Left            =   2850
         TabIndex        =   15
         Top             =   1500
         Width           =   1050
      End
   End
   Begin HSES.b8Container conData 
      Height          =   4695
      Index           =   4
      Left            =   90
      TabIndex        =   0
      Top             =   1260
      Width           =   9045
      _ExtentX        =   15954
      _ExtentY        =   8281
      BackColor       =   16185592
      Begin lvButton.lvButtons_H cmdEnrol 
         Height          =   375
         Left            =   1050
         TabIndex        =   1
         Top             =   1110
         Width           =   2685
         _ExtentX        =   4736
         _ExtentY        =   661
         Caption         =   "Enrol This Student"
         CapAlign        =   2
         BackStyle       =   5
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
         Mode            =   0
         Value           =   0   'False
         cBack           =   16185592
      End
      Begin lvButton.lvButtons_H cmdViewAllStudent 
         Height          =   375
         Left            =   1050
         TabIndex        =   2
         Top             =   1650
         Width           =   2685
         _ExtentX        =   4736
         _ExtentY        =   661
         Caption         =   "View All Student List"
         CapAlign        =   2
         BackStyle       =   5
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
         Mode            =   0
         Value           =   0   'False
         cBack           =   16185592
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pick A Task"
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
         Left            =   270
         TabIndex        =   3
         Top             =   450
         Width           =   1080
      End
   End
   Begin VB.Label Label28 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Edit Student Entry"
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
      TabIndex        =   82
      Top             =   90
      Width           =   2865
   End
   Begin VB.Image Image5 
      Height          =   720
      Left            =   0
      Picture         =   "frmEditStudent.frx":08CA
      Top             =   0
      Width           =   720
   End
   Begin VB.Image Image2 
      Height          =   135
      Left            =   -240
      Picture         =   "frmEditStudent.frx":1794
      Stretch         =   -1  'True
      Top             =   1125
      Width           =   9675
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00F6F8F8&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00BFCED0&
      Height          =   105
      Left            =   3960
      Top             =   1110
      Width           =   5745
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00F6F8F8&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00BFCED0&
      Height          =   495
      Left            =   0
      Top             =   750
      Width           =   6765
   End
End
Attribute VB_Name = "frmEditStudent"
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

Public Function ShowEdit(sStudentID As String) As Boolean
    
    'set dafaults
    RecordSaved = False
    
    'get student, if failed then exit
    If GetStudentByID(sStudentID, CurrentStudent) <> Success Then
        MsgBox "Unable to continue Editing Student!" & vbNewLine & "The Selected Student ID: " & sStudentID & " cannot be found in record.", vbExclamation
        Unload Me
        Exit Function
    End If
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
        
    
    
    'set textboxes..
    'get value from record
    With CurrentStudent
            'ignore student id, it was set already
            txtStudentID(0) = Left(.StudentID, 4)
            txtStudentID(1) = Right(.StudentID, 7)
            txtFirstName = .FirstName
            txtMiddleName = .MiddleName
            txtLastName = .LastName
            
            txtCityAddress = .CityAddress
            txtHomeAddress = .HomeAddress
            txtBirthDate.Value = .BirthDate
            txtPlaceOfBirth = .PlaceOfBirth
            cmbGender = .Gender
            cmbStatus = .Status
            txtCitizenship = .Citizenship
            txtReligion = .Religion
            txtBloodType = .BloodType
            
            txtLastSchoolName = .LastSchoolName
            txtLastSchoolContactNumber = .LastSchoolContactNumber
            txtLastSchoolAddress = .LastSchoolAddress
               
            'parents
            txtMotherName = .MotherName
            txtMotherOccupation = .MotherOccupation
            txtFatherName = .FatherName
            txtFatherOccupation = .FatherOccupation
            txtParentsContactNumber = .ParentsContactNumber
            txtParentsAddress = .ParentsAddress
            
            txtGuardianName = .GuardianName
            txtGuardianAddress = .GuardianAddress
            txtGuardianContactNumber = .GuardianContactNumber
            
            'txtCreationDate = .CreationDate
        
        End With
        
        
        
        
        
    
    
    
    
    
    
    txtBirthDate.MaxDate = Now - 1
    '415 = 11 years that is assigned to birth date current value which means it started at 11 yrs old
    txtAge.Text = (Now - txtBirthDate.Value) \ 365
    
    'set tabs
    SelectTab 1
    

    
    'show form
    
    
    Me.Show vbModal
    
    ShowEdit = RecordSaved
End Function


Private Function AllowShowTab(Index As Integer) As Boolean
    
    Dim sFullname As String
    Dim sOldFullName As String

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
        sOldFullName = LCase(Trim(CurrentStudent.FirstName) & Trim(CurrentStudent.MiddleName) & Trim(CurrentStudent.LastName))
        
        'if full name was changed
        If sFullname <> sOldFullName Then
            If FindDuplicateFullName(sFullname) = Success Then
                MsgBox "Invalid Name value. The Student name the you entered is already existed, please enter another value.", vbExclamation
                HLTxt txtFirstName
                Exit Function
            End If
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
            
            .ModifiedDate = FormatDateTime(Now, vbShortDate)
            .CreatedBy = CurrentUser.UserName
            
        End With

        
        'show summary
           lblSummary.Caption = "Student ID: " & CurrentStudent.StudentID
           lblSummary.Caption = lblSummary.Caption & vbNewLine & "First Name: " & CurrentStudent.FirstName
           lblSummary.Caption = lblSummary.Caption & vbNewLine & "Middle Name: " & CurrentStudent.MiddleName
           lblSummary.Caption = lblSummary.Caption & vbNewLine & "Last Name: " & CurrentStudent.LastName
           lblSummary.Caption = lblSummary.Caption & vbNewLine & "City Address: " & CurrentStudent.CityAddress
           lblSummary.Caption = lblSummary.Caption & vbNewLine & "Home Address: " & CurrentStudent.HomeAddress
           lblSummary.Caption = lblSummary.Caption & vbNewLine & "Birth Date: " & CurrentStudent.BirthDate
           lblSummary.Caption = lblSummary.Caption & vbNewLine & "Place Of Birth: " & CurrentStudent.PlaceOfBirth
           lblSummary.Caption = lblSummary.Caption & vbNewLine & "Gender: " & CurrentStudent.Gender
           lblSummary.Caption = lblSummary.Caption & vbNewLine & "Status: " & CurrentStudent.Status
           lblSummary.Caption = lblSummary.Caption & vbNewLine & "Citizenship: " & CurrentStudent.Citizenship
            
           lblSummary.Caption = lblSummary.Caption & vbNewLine & "Last School Attended: " & CurrentStudent.LastSchoolName
           lblSummary.Caption = lblSummary.Caption & vbNewLine & "Last School Attended Contact Number: " & CurrentStudent.LastSchoolContactNumber
           lblSummary.Caption = lblSummary.Caption & vbNewLine & "Last School Attended Address: " & CurrentStudent.LastSchoolAddress
               
            'parents
           lblSummary.Caption = lblSummary.Caption & vbNewLine & "Mother Name: " & CurrentStudent.MotherName
           lblSummary.Caption = lblSummary.Caption & vbNewLine & "Mother Occupation: " & CurrentStudent.MotherOccupation
           lblSummary.Caption = lblSummary.Caption & vbNewLine & "Father Name: " & CurrentStudent.FatherName
           lblSummary.Caption = lblSummary.Caption & vbNewLine & "Father Occupation: " & CurrentStudent.FatherOccupation
           lblSummary.Caption = lblSummary.Caption & vbNewLine & "Parents COntact Number: " & CurrentStudent.ParentsContactNumber
           lblSummary.Caption = lblSummary.Caption & vbNewLine & "ParentsAddress: " & CurrentStudent.ParentsAddress
            
           lblSummary.Caption = lblSummary.Caption & vbNewLine & "Guardian Name: " & CurrentStudent.GuardianName
           lblSummary.Caption = lblSummary.Caption & vbNewLine & "Guardian Contact Number: " & CurrentStudent.GuardianAddress
           lblSummary.Caption = lblSummary.Caption & vbNewLine & "Guardian Address: " & CurrentStudent.GuardianContactNumber
            
            lblSummary.Caption = lblSummary.Caption & vbNewLine & "Creation Date: " & CurrentStudent.CreationDate
    
        
    
    End If
    
    
    
    
    
    
    If Index >= 4 Then
        If RecordSaved = False Then Exit Function
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


Private Sub cmdCancel_Click()
    Unload Me
End Sub



Private Sub cmdEnrol_Click()
    'show add enrolment
    frmAddEnrolment.ShowForm CurrentStudent.StudentID
    'unload this form
    Unload Me
End Sub

Private Sub cmdNext_Click()
    SelectTab CurrentTab + 1
End Sub

Private Sub cmdPrevious_Click()
    SelectTab CurrentTab - 1
End Sub


Private Function CheckFields() As Boolean

    CheckFields = False
    
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
            
            .ModifiedDate = FormatDateTime(Now, vbShortDate)
            .ModifiedBy = CurrentUser.UserName
            
        End With

        CheckFields = True
        
End Function
Private Sub cmdSave_Click()
            If CheckFields = False Then
                Exit Sub
            End If
    
            'save record
            Select Case EditStudent(CurrentStudent)
                Case TranDBResult.Success
                    
                    MsgBox "Student entry successfully edited.", vbInformation
                    RecordSaved = True
                    lblsummaryinfo.Caption = "(New Entry Saved. CLick Next To Proceed."
                        

                    'goto option
                    SelectTab (4)
                    
                    
                Case Else ' failed
                    'fatal
                    'temp
                    MsgBox "Error! Unabled to saved entry.", vbExclamation
                    lblsummaryinfo.Caption = "Error! Unabled to saved entry."
            End Select
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

Private Sub Form_Load()
    'defaults
    
    'title case
    vfrmUseTitleCase = True
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

