VERSION 5.00
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "COMCT232.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmAddYearLevel 
   BackColor       =   &H00D8E9EC&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add New Year Level"
   ClientHeight    =   3420
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4395
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   228
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   293
   StartUpPosition =   1  'CenterOwner
   Begin HSES.b8Container b8Container1 
      Height          =   2055
      Left            =   90
      TabIndex        =   2
      Top             =   660
      Width           =   4155
      _ExtentX        =   7329
      _ExtentY        =   3625
      BackColor       =   16185592
      Begin ComCtl2.UpDown UpDownYearLevel 
         Height          =   330
         Left            =   2130
         TabIndex        =   3
         Top             =   510
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   582
         _Version        =   327681
         Value           =   1
         OrigLeft        =   2160
         OrigTop         =   675
         OrigRight       =   2400
         OrigBottom      =   1005
         Min             =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   -517
         Enabled         =   -1  'True
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Yearl Level Title"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   240
         TabIndex        =   7
         Top             =   1170
         Width           =   1350
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Year Level"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   240
         TabIndex        =   6
         Top             =   570
         Width           =   885
      End
      Begin MSForms.TextBox txtYearLevelTitle 
         Height          =   330
         Left            =   1650
         TabIndex        =   5
         Top             =   1110
         Width           =   2190
         VariousPropertyBits=   746604571
         ForeColor       =   128
         MaxLength       =   20
         BorderStyle     =   1
         Size            =   "3863;582"
         BorderColor     =   11366490
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin MSForms.TextBox txtYearLevelID 
         Height          =   330
         Left            =   1680
         TabIndex        =   4
         Top             =   510
         Width           =   420
         VariousPropertyBits=   746604571
         ForeColor       =   128
         MaxLength       =   20
         BorderStyle     =   1
         Size            =   "741;582"
         Value           =   "1"
         BorderColor     =   11366490
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
   End
   Begin lvButton.lvButtons_H cmdSave 
      Default         =   -1  'True
      Height          =   405
      Left            =   2790
      TabIndex        =   0
      Top             =   2850
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   714
      Caption         =   "&Save"
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
      Height          =   405
      Left            =   1200
      TabIndex        =   1
      Top             =   2850
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   714
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
   Begin VB.Image Image4 
      Height          =   135
      Left            =   0
      Picture         =   "frmAddYearLevel.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5445
   End
   Begin VB.Image Image2 
      Height          =   1155
      Left            =   0
      Picture         =   "frmAddYearLevel.frx":009D
      Stretch         =   -1  'True
      Top             =   2310
      Width           =   5535
   End
End
Attribute VB_Name = "frmAddYearLevel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Dim RecordAdded As Boolean

Public Function ShowForm() As Boolean


    Me.Show vbModal
    
    'return
    ShowForm = RecordAdded
End Function






Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Function SaveRecord()


    If IsNumeric(txtYearLevelID.Text) Then
        If Val(txtYearLevelID.Text) < 1 Then
            'temp
            MsgBox "Invalid Yearl Level. It must be equal or grater than 1", vbExclamation
            
            HLTxt txtYearLevelID
            
            Exit Function
        End If
    Else
        Exit Function
    End If
    

    If Not CheckTextBox(txtYearLevelTitle, "Please fill in Year Level Title") Then
        Exit Function
    End If
    
    'save record
    Dim newYearLevel As tYearLevel
    
    
    
    newYearLevel.YearLevelID = Val(txtYearLevelID.Text)
    newYearLevel.YearLevelTitle = Trim(txtYearLevelTitle.Text)
    
    
    
    
    Select Case AddYearLevel(newYearLevel)
        Case TranDBResult.Success
            'success
            MsgBox "Creating New Year has been successfull.", vbInformation
            
            'set flag
            RecordAdded = True
            'close this form
            Unload Me
        Case TranDBResult.DuplicateID
            MsgBox "Year Level ID already existed.", vbCritical
            HLTxt txtYearLevelID
        Case TranDBResult.DuplicateTitle
            MsgBox "Year Level Title already existed", vbCritical
            HLTxt txtYearLevelTitle
        Case Else
            'temp
            MsgBox "Unknown Error", vbCritical
    End Select
    
End Function

Private Sub cmdSave_Click()
    Call SaveRecord
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyEscape
            Unload Me
    End Select
End Sub

Private Sub txtYearLevel_KeyPress(KeyAscii As Integer)
    If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8 Or KeyAscii = 45) Then KeyAscii = 0
End Sub

