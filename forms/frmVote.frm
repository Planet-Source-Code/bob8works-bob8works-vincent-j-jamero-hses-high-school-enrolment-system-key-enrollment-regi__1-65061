VERSION 5.00
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmVote 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFEFE1&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Please Vote"
   ClientHeight    =   1230
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4590
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1230
   ScaleWidth      =   4590
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin lvButton.lvButtons_H cmdVote 
      Height          =   495
      Left            =   1290
      TabIndex        =   0
      Top             =   660
      Width           =   3165
      _ExtentX        =   5583
      _ExtentY        =   873
      Caption         =   "Give some Vote this Code"
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
      cBhover         =   16777215
      cGradient       =   16777215
      Gradient        =   3
      Mode            =   0
      Value           =   0   'False
      cBack           =   16773089
   End
   Begin lvButton.lvButtons_H lvButtons_H1 
      Height          =   495
      Left            =   150
      TabIndex        =   2
      Top             =   660
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   873
      Caption         =   "No, don't vote."
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
      cBhover         =   16777215
      cGradient       =   16777215
      Gradient        =   3
      Mode            =   0
      Value           =   0   'False
      cBack           =   16773089
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Please Give Some Vote To This Code"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   150
      TabIndex        =   1
      Top             =   90
      Width           =   4215
   End
End
Attribute VB_Name = "frmVote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Function ShowForm()
    Me.Show vbModal
End Function

Private Sub cmdVote_Click()
    OpenURL "http://www.pscode.com/vb/scripts/ShowCode.asp?lngWId=1&txtCodeId=65061", Me.hwnd
End Sub

Private Sub lvButtons_H1_Click()
    Unload Me
    mdiMain.timerVote.Enabled = False
End Sub
