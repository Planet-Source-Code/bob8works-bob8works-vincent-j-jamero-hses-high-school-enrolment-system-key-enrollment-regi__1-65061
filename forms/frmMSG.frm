VERSION 5.00
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmMSG 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Author Message"
   ClientHeight    =   6810
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5970
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6810
   ScaleWidth      =   5970
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin lvButton.lvButtons_H cmdVote 
      Height          =   345
      Left            =   60
      TabIndex        =   5
      Top             =   3450
      Width           =   2955
      _ExtentX        =   5212
      _ExtentY        =   609
      Caption         =   "Click Me To Vote"
      CapAlign        =   2
      BackStyle       =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFore           =   16512
      cFHover         =   16512
      cBhover         =   33023
      cGradient       =   33023
      Gradient        =   3
      Mode            =   0
      Value           =   0   'False
      cBack           =   12640511
   End
   Begin VB.CheckBox chkCls 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Clear User Entries To Show HSES Welcome Screen"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   30
      TabIndex        =   3
      Top             =   6540
      Width           =   4185
   End
   Begin lvButton.lvButtons_H cmdCLose 
      Cancel          =   -1  'True
      Default         =   -1  'True
      Height          =   315
      Left            =   4560
      TabIndex        =   2
      Top             =   6480
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   556
      Caption         =   "Close"
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
      cBhover         =   8421504
      cGradient       =   8421504
      Gradient        =   3
      Mode            =   0
      Value           =   0   'False
      cBack           =   16777215
   End
   Begin lvButton.lvButtons_H cmdVisit 
      Height          =   345
      Left            =   60
      TabIndex        =   6
      Top             =   3870
      Width           =   2955
      _ExtentX        =   5212
      _ExtentY        =   609
      Caption         =   "Visit bob8works.com"
      CapAlign        =   2
      BackStyle       =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFore           =   16512
      cFHover         =   16512
      cBhover         =   33023
      cGradient       =   33023
      Gradient        =   3
      Mode            =   0
      Value           =   0   'False
      cBack           =   12640511
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Kung Gusto nyo magpabuhat ug Program, just contact me.   CEL. # 09069223213"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0FF&
      Height          =   945
      Left            =   120
      TabIndex        =   7
      Top             =   4320
      Width           =   2175
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Please E-Mail me if you this HSES Teacher Application                   -->"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   525
      Left            =   4200
      TabIndex        =   4
      Top             =   3840
      Width           =   1905
   End
   Begin VB.Image Image3 
      Height          =   2085
      Left            =   2400
      Picture         =   "frmMSG.frx":0000
      Top             =   4320
      Width           =   3510
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "I hope this one will helps you. Happy coding!"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0FF&
      Height          =   465
      Left            =   3120
      TabIndex        =   1
      Top             =   3330
      Width           =   2745
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmMSG.frx":2179
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0FF&
      Height          =   945
      Left            =   3150
      TabIndex        =   0
      Top             =   2370
      Width           =   2775
   End
   Begin VB.Image Image2 
      Height          =   3195
      Left            =   30
      Picture         =   "frmMSG.frx":221F
      Stretch         =   -1  'True
      Top             =   3270
      Width           =   6015
   End
   Begin VB.Image Image1 
      Height          =   3270
      Left            =   0
      Picture         =   "frmMSG.frx":2715
      Top             =   0
      Width           =   6015
   End
End
Attribute VB_Name = "frmMSG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub ShowForm()

    Me.Show vbModal
End Sub

Private Sub cmdCLose_Click()
    Dim vUser As User
    
    If chkCls.Value = vbChecked Then
        DeleteAllUser
    End If
    
    
    Unload Me
End Sub

Private Function DeleteAllUser()

    Dim vRS As New ADODB.Recordset
    Dim DB As New ADODB.Connection
    
    Dim sSQL As String
    
    sSQL = "Delete * from tblUser"
    
    ConnectDB DB, App.Path & "/HSESData.mdb"
    ConnectRS DB, vRS, sSQL
    
    Set vRS = Nothing
    Set DB = Nothing
End Function

Private Sub cmdVisit_Click()
    OpenURL "www.bob8works.cjb.net", Me.hwnd

End Sub

Private Sub cmdVote_Click()
    OpenURL "http://www.pscode.com/vb/scripts/ShowCode.asp?lngWId=1&txtCodeId=65061", Me.hwnd
End Sub

