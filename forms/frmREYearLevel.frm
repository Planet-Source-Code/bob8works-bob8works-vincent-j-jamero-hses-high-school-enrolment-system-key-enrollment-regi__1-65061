VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmREYearLevel 
   Caption         =   "Explore: Year Level"
   ClientHeight    =   5070
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6300
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmREYearLevel.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   338
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   420
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList imgListIco32 
      Left            =   2550
      Top             =   2835
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
            Picture         =   "frmREYearLevel.frx":058A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox bgInfo 
      Appearance      =   0  'Flat
      BackColor       =   &H00F6F8F8&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1740
      Left            =   30
      ScaleHeight     =   1740
      ScaleWidth      =   8910
      TabIndex        =   2
      Top             =   360
      Width           =   8910
      Begin lvButton.lvButtons_H cmdShowEnrolmentList 
         Height          =   390
         Left            =   675
         TabIndex        =   3
         Top             =   480
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   688
         Caption         =   "View All Enrolment Entries In This Department"
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
         cFore           =   4210752
         cFHover         =   4210752
         Focus           =   0   'False
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         Image           =   "frmREYearLevel.frx":0E64
         cBack           =   16185592
      End
      Begin lvButton.lvButtons_H lvButtons_H1 
         Height          =   390
         Left            =   675
         TabIndex        =   4
         Top             =   855
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   688
         Caption         =   "Print All Enrolment Entries In This Department"
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
         cFore           =   4210752
         cFHover         =   4210752
         Focus           =   0   'False
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         Image           =   "frmREYearLevel.frx":0FBE
         cBack           =   16185592
      End
      Begin VB.Label lblSectionTItle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Pick a task"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C25418&
         Height          =   285
         Left            =   105
         TabIndex        =   6
         Top             =   60
         Width           =   1335
      End
      Begin VB.Image Image7 
         Height          =   30
         Left            =   0
         Picture         =   "frmREYearLevel.frx":1118
         Stretch         =   -1  'True
         Top             =   0
         Width           =   8925
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "or Select a folder"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C25418&
         Height          =   285
         Left            =   105
         TabIndex        =   5
         Top             =   1365
         Width           =   2100
      End
      Begin VB.Image Image6 
         Height          =   90
         Left            =   -30
         Picture         =   "frmREYearLevel.frx":11B5
         Stretch         =   -1  'True
         Top             =   1680
         Width           =   8925
      End
   End
   Begin MSComctlLib.ImageList imgListIco16 
      Left            =   4125
      Top             =   2460
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
            Picture         =   "frmREYearLevel.frx":1252
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView listRecord 
      Height          =   2880
      Left            =   60
      TabIndex        =   0
      Top             =   2100
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   5080
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      Icons           =   "imgListIco32"
      SmallIcons      =   "imgListIco16"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   0
   End
   Begin HSES.b8ChildTitleBar TitleBar 
      Height          =   360
      Left            =   1785
      TabIndex        =   1
      Top             =   15
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   635
      BackColor       =   13724971
      Caption         =   "Department"
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
      ForeColor       =   16777215
      GradTheme       =   1
   End
End
Attribute VB_Name = "frmREYearLevel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim curSchoolYearTitle As String
Dim curDepartmentTitle As String


Public Sub ShowForm(sSchoolYearTitle As String, sDepartmentTitle As String, sKey() As String, sText() As String)
    Dim i As Integer
    
    On Error Resume Next
    
    curSchoolYearTitle = sSchoolYearTitle
    curDepartmentTitle = sDepartmentTitle
    
    TitleBar.Caption = "S.Y.: " & sSchoolYearTitle & " - " & sDepartmentTitle
    
    Me.Show
    DoEvents
    
    listRecord.ListItems.Clear
    
    'error may occured here
    If UBound(sText) < 1 Then
        Exit Sub
    End If
    
    For i = 0 To UBound(sText)
        listRecord.ListItems.Add , sKey(i), sText(i), 1, 1
    Next
End Sub

Private Sub cmdShowEnrolmentList_Click()
    frmAllEnrolment.ShowFormList curSchoolYearTitle, curDepartmentTitle
End Sub

Private Sub Form_Activate()
    mdiMain.RegMDIChild Me
End Sub
Private Sub Form_Resize()
    On Error Resume Next
        
    TitleBar.Move 0, 0, Me.ScaleWidth
    bgInfo.Move 0, TitleBar.Height, Me.ScaleWidth
    
    listRecord.Move 0, listRecord.Top, Me.ScaleWidth, Me.ScaleHeight
End Sub

Private Sub Form_Deactivate()
    Unload Me
End Sub


Private Sub listRecord_DblClick()

    If listRecord.ListItems.Count < 1 Then
        Exit Sub
    End If
    
    
    mdiMain.RecordTree.SelectNode listRecord.SelectedItem.Key

End Sub

Private Sub listRecord_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Call listRecord_DblClick
    End If
End Sub

Public Function Form_CloseExplore()
    Unload Me
End Function


Public Function Form_CanExplore() As Boolean
    Form_CanExplore = True
End Function

