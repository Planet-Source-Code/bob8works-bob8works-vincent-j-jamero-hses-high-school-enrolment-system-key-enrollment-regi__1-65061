VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form PickTeacher 
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5070
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4575
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   338
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   305
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00D8E9EC&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5055
      Left            =   15
      ScaleHeight     =   337
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   303
      TabIndex        =   0
      Top             =   15
      Width           =   4545
      Begin lvButton.lvButtons_H cmdCancel 
         Cancel          =   -1  'True
         Height          =   375
         Left            =   2250
         TabIndex        =   1
         Top             =   4680
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         Caption         =   "&Cancel"
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
         cFore           =   8421504
         cFHover         =   8421504
         cBhover         =   16185592
         cGradient       =   16185592
         Gradient        =   3
         Mode            =   0
         Value           =   0   'False
         cBack           =   14215660
      End
      Begin lvButton.lvButtons_H cmdSelect 
         Default         =   -1  'True
         Height          =   375
         Left            =   3360
         TabIndex        =   2
         Top             =   4680
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   661
         Caption         =   "&Select"
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
         cFore           =   8421504
         cFHover         =   8421504
         cBhover         =   16185592
         Focus           =   0   'False
         cGradient       =   16185592
         Gradient        =   3
         Mode            =   0
         Value           =   0   'False
         cBack           =   14215660
      End
      Begin HSES.b8Container b8Container1 
         Height          =   4395
         Left            =   30
         TabIndex        =   3
         Top             =   270
         Width           =   4485
         _ExtentX        =   7911
         _ExtentY        =   7752
         BackColor       =   16185592
         Begin MSComctlLib.ListView listRecord 
            Height          =   4290
            Left            =   60
            TabIndex        =   4
            Top             =   60
            Width           =   4365
            _ExtentX        =   7699
            _ExtentY        =   7567
            View            =   3
            LabelEdit       =   1
            MultiSelect     =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            HideColumnHeaders=   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            Icons           =   "imgListStudent"
            SmallIcons      =   "imgListStudent"
            ForeColor       =   -2147483640
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
            NumItems        =   2
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Object.Width           =   3175
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Object.Width           =   8440
            EndProperty
         End
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Select Teacher"
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
         Left            =   390
         TabIndex        =   5
         Top             =   30
         Width           =   1260
      End
      Begin VB.Image imgIcon 
         Height          =   240
         Left            =   30
         Picture         =   "PickTeacher.frx":0000
         Stretch         =   -1  'True
         Top             =   30
         Width           =   240
      End
      Begin VB.Image Image2 
         Height          =   405
         Left            =   0
         Picture         =   "PickTeacher.frx":058A
         Stretch         =   -1  'True
         Top             =   4650
         Width           =   6495
      End
      Begin VB.Image Image4 
         Height          =   135
         Left            =   30
         Picture         =   "PickTeacher.frx":0627
         Stretch         =   -1  'True
         Top             =   0
         Width           =   5085
      End
   End
End
Attribute VB_Name = "PickTeacher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim tmpTeacherID As String
Dim tmpTeacherFullName As String

Public Function GetTeacherID(Optional ByRef sTeacherFullName As String) As String
            
    'set fail to default
    GetTeacherID = ""
        
    'add yr to list
    If Not FillList Then
        MsgBox "There is no TEACHER entries to display.", vbExclamation
        Unload Me
        Exit Function
    End If

    Me.Show vbModal
    
    'return
    sTeacherFullName = IIf(IsMissing(sTeacherFullName), "", tmpTeacherFullName)
    GetTeacherID = tmpTeacherID
End Function


Private Sub ReturnGetTeacherID()
    If Len(GetLVKey(listRecord.SelectedItem)) > 0 Then
    
        tmpTeacherFullName = listRecord.SelectedItem.Text
        tmpTeacherID = GetLVKey(listRecord.SelectedItem)
        
        'call return
        Unload Me
    
    End If
End Sub
Private Sub CancelGetTeacherID()
    tmpTeacherID = ""
    Unload Me
End Sub

Private Sub cmdCancel_Click()
    CancelGetTeacherID
End Sub

Private Sub cmdSelect_Click()
    ReturnGetTeacherID
End Sub



Private Sub listRecord_DblClick()
    ReturnGetTeacherID
End Sub

Private Function FillList() As Boolean
    Dim vRS As New ADODB.Recordset
    
    If ConnectRS(HSESDB, vRS, "SELECT tblTeacher.TeacherID AS lvKEY, [tblTeacher]![LastName]+', '+[tblTeacher]![FirstName]+' '+[tblTeacher]![MiddleName] AS [Full Name] FROM tblTeacher;") = True Then
        If AnyRecordExisted(vRS) Then
            
            FillRecordToList vRS, listRecord, KeyTeacher
            
            FillList = True
        Else
            FillList = False
        End If
    Else
        FillList = False
    End If
    Set vRS = Nothing
End Function
