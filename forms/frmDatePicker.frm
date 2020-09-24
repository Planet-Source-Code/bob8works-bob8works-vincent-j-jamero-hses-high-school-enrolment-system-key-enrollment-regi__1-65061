VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmDatePicker 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Select Date"
   ClientHeight    =   3120
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   2790
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   208
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   186
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   2805
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2805
      _ExtentX        =   4948
      _ExtentY        =   4948
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      BackColor       =   14215660
      TabCaption(0)   =   "Date Picker"
      TabPicture(0)   =   "frmDatePicker.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "calendarNewDate"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      Begin MSComCtl2.MonthView calendarNewDate 
         Height          =   2370
         Left            =   45
         TabIndex        =   3
         Top             =   360
         Width           =   2700
         _ExtentX        =   4763
         _ExtentY        =   4180
         _Version        =   393216
         ForeColor       =   -2147483630
         BackColor       =   -2147483633
         Appearance      =   1
         StartOfWeek     =   225312769
         TitleBackColor  =   33023
         TitleForeColor  =   16777215
         CurrentDate     =   38064
      End
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00D8E9EC&
      Cancel          =   -1  'True
      Caption         =   "&Cancel (ESC)"
      Height          =   375
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2790
      Width           =   1365
   End
   Begin VB.CommandButton cmdSelect 
      BackColor       =   &H00D8E9EC&
      Caption         =   "&Select (Enter)"
      Default         =   -1  'True
      Height          =   375
      Left            =   1350
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2790
      Width           =   1455
   End
End
Attribute VB_Name = "frmDatePicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim vGetDate As Boolean
Dim vNewDate As Date


Public Function GetDate(ByRef newDate As Date) As Boolean
    
    

    Me.Show vbModal
    
    'return value
    newDate = vNewDate
    GetDate = vGetDate
End Function


Private Sub ReturnDate()

    vGetDate = True
    vNewDate = CDate(calendarNewDate.Value)
    Unload Me
End Sub
Private Sub CancelDate()
    'return value
    vGetDate = False
    Unload Me
End Sub


Private Sub calendarNewDate_DateDblClick(ByVal DateDblClicked As Date)
    Call ReturnDate
End Sub

Private Sub cmdCancel_Click()
    Call CancelDate
End Sub

Private Sub cmdSelect_Click()
    Call ReturnDate
End Sub




Private Sub Form_Load()
    calendarNewDate.Value = Now
End Sub
