VERSION 5.00
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmWelcome 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "HSES 1 Setup"
   ClientHeight    =   4995
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8445
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmWelcome.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   333
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   563
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin HSES.b8SContainer b8SContainer1 
      Height          =   615
      Left            =   -60
      TabIndex        =   1
      Top             =   4410
      Width           =   8595
      _ExtentX        =   15161
      _ExtentY        =   1085
      BorderColor     =   12632256
      Begin lvButton.lvButtons_H cmdNext 
         Default         =   -1  'True
         Height          =   405
         Left            =   7230
         TabIndex        =   2
         Top             =   90
         Width           =   1155
         _ExtentX        =   2037
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
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00F6F8F8&
      BorderStyle     =   0  'None
      Height          =   4695
      Left            =   0
      ScaleHeight     =   4695
      ScaleWidth      =   2025
      TabIndex        =   0
      Top             =   0
      Width           =   2025
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "CLick Next to Continue"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2100
      TabIndex        =   3
      Top             =   4080
      Width           =   1785
   End
   Begin VB.Image imgPic 
      Height          =   3615
      Left            =   2010
      Top             =   0
      Width           =   4695
   End
End
Attribute VB_Name = "frmWelcome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdNext_Click()

    CurrentUser.UserName = "system"
    
    Unload Me
    frmAddUser.ShowForm True
End Sub

Private Sub Form_Load()
    Load frmSplash
    Set imgPic.Picture = frmSplash.Picture
    Unload frmSplash
End Sub
