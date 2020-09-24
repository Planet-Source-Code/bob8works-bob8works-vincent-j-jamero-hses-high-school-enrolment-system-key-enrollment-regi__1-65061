VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmEditYearLevel 
   BackColor       =   &H00D8E9EC&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Edit Year Level"
   ClientHeight    =   3375
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4290
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
   ScaleHeight     =   225
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   286
   StartUpPosition =   1  'CenterOwner
   Begin HSES.b8Container b8Container1 
      Height          =   1935
      Left            =   90
      TabIndex        =   2
      Top             =   660
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   3413
      BackColor       =   16185592
      Begin MSForms.TextBox txtYearLevelID 
         Height          =   330
         Left            =   1650
         TabIndex        =   6
         Top             =   480
         Width           =   420
         VariousPropertyBits=   746604575
         BackColor       =   16777215
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
      Begin MSForms.TextBox txtYearLevelTitle 
         Height          =   330
         Left            =   1650
         TabIndex        =   5
         Top             =   1050
         Width           =   2190
         VariousPropertyBits=   746604571
         BackColor       =   16777215
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
         ForeColor       =   &H00B14801&
         Height          =   195
         Left            =   240
         TabIndex        =   4
         Top             =   510
         Width           =   885
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
         ForeColor       =   &H00B14801&
         Height          =   195
         Left            =   240
         TabIndex        =   3
         Top             =   1110
         Width           =   1350
      End
   End
   Begin lvButton.lvButtons_H cmdUpdate 
      Default         =   -1  'True
      Height          =   405
      Left            =   2760
      TabIndex        =   0
      Top             =   2790
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   714
      Caption         =   "&Update"
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
      Left            =   1170
      TabIndex        =   1
      Top             =   2790
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
   Begin VB.Image Image2 
      Height          =   1155
      Left            =   0
      Picture         =   "frmEditYearLevel.frx":0000
      Stretch         =   -1  'True
      Top             =   2220
      Width           =   5535
   End
   Begin VB.Image Image4 
      Height          =   135
      Left            =   0
      Picture         =   "frmEditYearLevel.frx":009D
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5445
   End
End
Attribute VB_Name = "frmEditYearLevel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vCurrentYearLevel As tYearLevel
Dim RecordEdited As Boolean


Public Function ShowEdit(lYearLevelID As Integer) As Boolean


     Dim vYearlevel As tYearLevel
     
     
    If GetYearLevelByID(lYearLevelID, vYearlevel) <> Success Then
        MsgBox "The selected Year Level does not exist. Unable to continue this operation.", vbExclamation
        Unload Me
        Exit Function
    End If
     

        vCurrentYearLevel = vYearlevel
        txtYearLevelID.Text = vYearlevel.YearLevelID
        txtYearLevelTitle.Text = vYearlevel.YearLevelTitle
        
        'show
        Me.Show vbModal
        
        'return
        ShowEdit = RecordEdited

    
End Function






Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Function UpdateRecord()

    'check id
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

    Select Case EditYearLevel(newYearLevel)
        
        Case TranDBResult.Success
            'went success full
            MsgBox "Editing Year Level has been successfull.", vbInformation
            'close this form
            RecordEdited = True
            Unload Me
        Case TranDBResult.DuplicateID
            'temp
            MsgBox "The YEAR LEVEL ID you have entered already exist. Please enter another.", vbExclamation
            HLTxt txtYearLevelID
        Case TranDBResult.DuplicateTitle
            'temp
            MsgBox "The YEAR LEVEL TITLE you have entered already exist. Please enter another.", vbExclamation
            HLTxt txtYearLevelTitle
        Case Else
            'fatal
            MsgBox "Unknown Error", vbExclamation
    
    End Select
    
End Function


Private Sub cmdUpdate_Click()
    Call UpdateRecord
End Sub


Private Sub txtYearLevel_KeyPress(KeyAscii As Integer)
    If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8 Or KeyAscii = 45) Then KeyAscii = 0
End Sub


