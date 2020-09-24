VERSION 5.00
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmPrintSections 
   BackColor       =   &H00D8E9EC&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Print Section"
   ClientHeight    =   3555
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7125
   Icon            =   "frmPrintSections.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3555
   ScaleWidth      =   7125
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00D8E9EC&
      Caption         =   "Options"
      ForeColor       =   &H00C25418&
      Height          =   1920
      Left            =   180
      TabIndex        =   5
      Top             =   690
      Width           =   6690
      Begin VB.CheckBox chkDepartment 
         BackColor       =   &H00D8E9EC&
         Caption         =   "Department"
         Height          =   315
         Left            =   300
         TabIndex        =   11
         Top             =   1530
         Value           =   1  'Checked
         Width           =   2235
      End
      Begin VB.CheckBox chkCreatedBy 
         BackColor       =   &H00D8E9EC&
         Caption         =   "Created By"
         Height          =   345
         Left            =   3270
         TabIndex        =   10
         Top             =   1200
         Value           =   1  'Checked
         Width           =   1515
      End
      Begin VB.CheckBox chkCreationDate 
         BackColor       =   &H00D8E9EC&
         Caption         =   "Creation Date"
         Height          =   345
         Left            =   3270
         TabIndex        =   8
         Top             =   825
         Value           =   1  'Checked
         Width           =   1515
      End
      Begin VB.CheckBox chkSectionFullTitle 
         BackColor       =   &H00D8E9EC&
         Caption         =   "Section Complete Title"
         Height          =   315
         Left            =   315
         TabIndex        =   7
         Top             =   1155
         Value           =   1  'Checked
         Width           =   2235
      End
      Begin VB.CheckBox chkSectionID 
         BackColor       =   &H00D8E9EC&
         Caption         =   "Section ID"
         Height          =   345
         Left            =   315
         TabIndex        =   6
         Top             =   810
         Width           =   1155
      End
      Begin VB.Label Label1 
         BackColor       =   &H00D8E9EC&
         Caption         =   "Include/Remove Field"
         Height          =   300
         Left            =   210
         TabIndex        =   9
         Top             =   495
         Width           =   2925
      End
   End
   Begin HSES.b8Line b8Line1 
      Height          =   60
      Left            =   0
      TabIndex        =   0
      Top             =   510
      Width           =   7410
      _ExtentX        =   13070
      _ExtentY        =   106
   End
   Begin lvButton.lvButtons_H cmdPrint 
      Default         =   -1  'True
      Height          =   360
      Left            =   5445
      TabIndex        =   1
      Top             =   3030
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   635
      Caption         =   "&Print/Preview"
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
   Begin lvButton.lvButtons_H cmdCancel 
      Cancel          =   -1  'True
      Height          =   360
      Left            =   3795
      TabIndex        =   2
      Top             =   3030
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   635
      Caption         =   "&Cancel"
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
   Begin HSES.b8Line b8Line2 
      Height          =   60
      Left            =   -15
      TabIndex        =   3
      Top             =   2805
      Width           =   7380
      _ExtentX        =   13018
      _ExtentY        =   106
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   30
      Picture         =   "frmPrintSections.frx":08CA
      Top             =   30
      Width           =   480
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Print Section"
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
      Left            =   690
      TabIndex        =   4
      Top             =   120
      Width           =   1245
   End
   Begin VB.Image Image2 
      Height          =   525
      Left            =   0
      Picture         =   "frmPrintSections.frx":1194
      Stretch         =   -1  'True
      Top             =   0
      Width           =   7410
   End
End
Attribute VB_Name = "frmPrintSections"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim curRS As ADODB.Recordset
Public Function ShowForm(ByRef vRS As ADODB.Recordset)

    Set curRS = vRS
    
    Me.Show vbModal
End Function


Private Sub chkSectionFullTitle_Click()
    chkSectionFullTitle.Value = vbChecked
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()

    Set drSectionList.DataSource = curRS
    
    If chkSectionID.Value <> vbChecked Then
        drSectionList.Sections("Section1").Controls("text1").Visible = False
        drSectionList.Sections("Section2").Controls("label1").Visible = False

    End If
    If chkSectionFullTitle.Value <> vbChecked Then
        drSectionList.Sections("Section1").Controls("text2").Visible = False
        drSectionList.Sections("Section2").Controls("label2").Visible = False

    End If
    If chkDepartment.Value <> vbChecked Then
        drSectionList.Sections("Section1").Controls("text3").Visible = False
        drSectionList.Sections("Section2").Controls("label3").Visible = False

    End If
    If chkCreationDate.Value <> vbChecked Then
        drSectionList.Sections("Section1").Controls("text4").Visible = False
    End If
    If chkCreatedBy.Value <> vbChecked Then
        drSectionList.Sections("Section1").Controls("text5").Visible = False
    End If
    If chkCreatedBy.Value <> vbChecked And chkCreationDate.Value <> vbChecked Then
        drSectionList.Sections("Section2").Controls("label4").Visible = False
    End If
    
    
    drSectionList.Show vbModal
End Sub



Private Sub Form_Unload(Cancel As Integer)
    Set curRS = Nothing
End Sub

