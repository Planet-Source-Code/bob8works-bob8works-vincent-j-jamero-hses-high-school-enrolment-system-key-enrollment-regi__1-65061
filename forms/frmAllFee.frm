VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmAllFee 
   BackColor       =   &H00D8E9EC&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Fees"
   ClientHeight    =   6180
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   8955
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   412
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   597
   Begin MSComctlLib.ListView listRecord 
      Height          =   4485
      Left            =   60
      TabIndex        =   11
      Top             =   1080
      Width           =   8835
      _ExtentX        =   15584
      _ExtentY        =   7911
      View            =   3
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Title"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Amount"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "School Year"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Department"
         Object.Width           =   4762
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Year Level"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Descrition"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Creation Date"
         Object.Width           =   2117
      EndProperty
   End
   Begin VB.CommandButton cmdAddSubject 
      Appearance      =   0  'Flat
      BackColor       =   &H00D8E9EC&
      Height          =   285
      Left            =   9420
      Picture         =   "frmAllFee.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1170
      Width           =   255
   End
   Begin lvButton.lvButtons_H cmdClose 
      Cancel          =   -1  'True
      Height          =   360
      Left            =   7800
      TabIndex        =   1
      Top             =   5730
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   635
      Caption         =   "&Close"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin HSES.b8Line b8Line1 
      Height          =   60
      Left            =   0
      TabIndex        =   2
      Top             =   510
      Width           =   10065
      _ExtentX        =   17754
      _ExtentY        =   106
   End
   Begin HSES.b8Line b8Line2 
      Height          =   60
      Left            =   -1140
      TabIndex        =   3
      Top             =   5610
      Width           =   10065
      _ExtentX        =   17754
      _ExtentY        =   106
   End
   Begin HSES.b8Line b8Line3 
      Height          =   60
      Left            =   0
      TabIndex        =   7
      Top             =   960
      Width           =   10065
      _ExtentX        =   17754
      _ExtentY        =   106
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00D8E9EC&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   405
      Left            =   0
      ScaleHeight     =   27
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   583
      TabIndex        =   5
      Top             =   570
      Width           =   8745
      Begin lvButton.lvButtons_H cmdNew 
         Height          =   375
         Left            =   60
         TabIndex        =   6
         Top             =   0
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   661
         Caption         =   "New"
         CapAlign        =   2
         BackStyle       =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
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
         Image           =   "frmAllFee.frx":058A
         cBack           =   14215660
      End
      Begin lvButton.lvButtons_H cmdDelete 
         Height          =   375
         Left            =   1200
         TabIndex        =   8
         Top             =   0
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   661
         Caption         =   "Delete"
         CapAlign        =   2
         BackStyle       =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
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
         Image           =   "frmAllFee.frx":0B24
         cBack           =   14215660
      End
      Begin lvButton.lvButtons_H cmdFind 
         Height          =   375
         Left            =   2340
         TabIndex        =   9
         Top             =   0
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   661
         Caption         =   "Find"
         CapAlign        =   2
         BackStyle       =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
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
         Image           =   "frmAllFee.frx":10BE
         cBack           =   14215660
      End
      Begin lvButton.lvButtons_H cmdReload 
         Height          =   375
         Left            =   3480
         TabIndex        =   10
         Top             =   0
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   661
         Caption         =   "Reload"
         CapAlign        =   2
         BackStyle       =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   12307149
         Mode            =   0
         Value           =   0   'False
         Image           =   "frmAllFee.frx":1658
         cBack           =   14215660
      End
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   60
      Picture         =   "frmAllFee.frx":1F32
      Top             =   30
      Width           =   480
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "All Fees"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H002F556A&
      Height          =   240
      Left            =   630
      TabIndex        =   4
      Top             =   120
      Width           =   735
   End
   Begin VB.Image Image2 
      Height          =   525
      Left            =   0
      Picture         =   "frmAllFee.frx":27FC
      Stretch         =   -1  'True
      Top             =   0
      Width           =   10065
   End
   Begin VB.Menu mnuRecord 
      Caption         =   "&Record"
      Begin VB.Menu mnuAddEntry 
         Caption         =   "&Add Entry"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "&Delete Sected"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnuFind 
         Caption         =   "&Find"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuReload 
         Caption         =   "&Reload"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuSeparator1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "frmAllFee"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Function ShowForm()
    
    
    'show form
    Me.Show vbModal
End Function

Private Sub cmdCLose_Click()
    Unload Me
End Sub

Private Sub cmdDelete_Click()
    Form_Delete
End Sub

Private Sub cmdFind_Click()
    Form_Find
End Sub

Private Sub cmdNew_Click()
    Form_Add
End Sub

Private Function Form_Add()
    If frmAddFee.ShowForm = True Then
        Form_FillRecord
    End If
End Function

Private Function Form_Delete()
    If frmDeleteFee.ShowForm(GetLVKey(listRecord.SelectedItem)) = True Then
        Form_FillRecord
    End If
End Function
Private Sub Form_Find()
    frmFindListItem.ShowFind listRecord
End Sub
Private Sub Form_RefreshButtons()

    If listRecord.ListItems.Count > 0 Then
        cmdFind.Enabled = True
        cmdDelete.Enabled = True
        mnuFind.Enabled = True
        mnuDelete.Enabled = True
    Else
        cmdFind.Enabled = False
        cmdDelete.Enabled = False
        mnuFind.Enabled = False
        mnuDelete.Enabled = False
    End If
End Sub

Private Function Form_FillRecord()

    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    sSQL = "SELECT tblFee.FeeID, tblFee.Title, tblFee.Amount, IIf(Len([tblFee]![SchoolYear])<1,'ALL',[tblFee]![SchoolYear]) AS SchoolYear, IIf(Len([tblFee]![DepartmentID])<1,'ALL',[tblDepartment]![DepartmentTitle]) AS Departmentt, IIf([tblFee]![YearLevelID]=0,'ALL',[tblYearLevel]![YearLevelTitle]) AS YearLevel, tblFee.Description, tblFee.CreationDate" & _
            " FROM (tblFee LEFT JOIN tblDepartment ON tblFee.DepartmentID = tblDepartment.DepartmentID) LEFT JOIN tblYearLevel ON tblFee.YearLevelID = tblYearLevel.YearLevelID;"

    'clear
    Me.MousePointer = vbHourglass
    listRecord.Enabled = True
    listRecord.ListItems.Clear
    
    If ConnectRS(HSESDB, vRS, sSQL) = True Then
        FillRecordToList vRS, listRecord, KeyFee
        listRecord.Enabled = True
    Else
        'fatal Error
        CatchError "AllFee", "Form_fillRecord", "Fee record not connected"
    End If
    
    'restore
    
    Me.MousePointer = vbDefault
    
    'release
    Set vRS = Nothing
    
    Form_RefreshButtons
End Function

Private Sub cmdReload_Click()
    'reload
    Form_FillRecord
End Sub

Private Sub Form_Activate()
    'refresh list
    Form_FillRecord
End Sub


Private Sub listRecord_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
        Me.PopupMenu mnuRecord
    End If
End Sub

Private Sub mnuAddEntry_Click()
    Form_Add
End Sub

Private Sub mnuDelete_Click()
    Form_Delete
End Sub

Private Sub mnuExit_Click()
    Unload Me
End Sub

Private Sub mnuFind_Click()
    Form_Find
End Sub

Private Sub mnuReload_Click()
    Form_FillRecord
End Sub
