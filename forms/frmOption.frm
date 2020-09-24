VERSION 5.00
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmOption 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Options"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   206
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkUseTitleCase 
      Caption         =   "Use Sentense Case"
      Height          =   255
      Left            =   690
      TabIndex        =   0
      Top             =   240
      Width           =   1815
   End
   Begin lvButton.lvButtons_H cmdOK 
      Default         =   -1  'True
      Height          =   405
      Left            =   3150
      TabIndex        =   1
      Top             =   2580
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   714
      Caption         =   "&OK"
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
      Left            =   1560
      TabIndex        =   2
      Top             =   2580
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
End
Attribute VB_Name = "frmOption"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Function ShowOption(Frm As Form)
    
    
    Me.Show vbModal
    
End Function


Private Function ReturnOption()

    On Error Resume Next
    
    
    If chkUseTitleCase.Value = vbcheked Then
        Frm.vfrmUseTitleCase = True
    Else
        Frm.vfrmUseTitleCase = False
    End If
    
    Unload mr
End Function

Private Function CancelOption()
    Unload Me
End Function

Private Sub cmdCancel_Click()
    CancelOption
End Sub

Private Sub cmdOK_Click()
    
    ReturnOption
End Sub
