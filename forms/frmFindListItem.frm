VERSION 5.00
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmFindListItem 
   BackColor       =   &H00D8E9EC&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "   Find"
   ClientHeight    =   1695
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5625
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   113
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   375
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtFindWhat 
      Height          =   315
      Left            =   1050
      TabIndex        =   6
      Top             =   150
      Width           =   2535
   End
   Begin VB.CheckBox chkWholeWord 
      BackColor       =   &H00D8E9EC&
      Caption         =   "Find Whole Field Value Only"
      ForeColor       =   &H00000000&
      Height          =   345
      Left            =   420
      TabIndex        =   4
      Top             =   660
      Width           =   2655
   End
   Begin VB.CheckBox chkMultiSelect 
      BackColor       =   &H00D8E9EC&
      Caption         =   "Select All"
      ForeColor       =   &H00000000&
      Height          =   345
      Left            =   420
      TabIndex        =   3
      Top             =   960
      Width           =   2655
   End
   Begin VB.CheckBox chkInverseSelection 
      BackColor       =   &H00D8E9EC&
      Caption         =   "Inverse Selection"
      ForeColor       =   &H00000000&
      Height          =   345
      Left            =   420
      TabIndex        =   2
      Top             =   1260
      Width           =   2655
   End
   Begin lvButton.lvButtons_H cmdFind 
      Default         =   -1  'True
      Height          =   375
      Left            =   4200
      TabIndex        =   0
      Top             =   180
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   661
      Caption         =   "&Find"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
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
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdCancel 
      Cancel          =   -1  'True
      Height          =   345
      Left            =   4200
      TabIndex        =   1
      Top             =   660
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   609
      Caption         =   "&Cancel"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
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
      cBack           =   -2147483633
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Find What:"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   240
      TabIndex        =   5
      Top             =   180
      Width           =   795
   End
End
Attribute VB_Name = "frmFindListItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim tmpLV As ListView

Public Function ShowFind(lv As ListView)

    'check list count
    If lv.ListItems.Count < 1 Then
        MsgBox "There is no necord in list to find.", vbExclamation
        Unload Me
    End If
    
    'get lv
    Set tmpLV = lv
    
    Me.Show vbModal
    
End Function



Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdFind_Click()

    Dim tmpMultiSelect As Boolean
    Dim tmpInverseSelection As Boolean
    
    'trim
    txtFindWhat.Text = Trim(txtFindWhat.Text)
    
    'check length
    If Len(txtFindWhat.Text) < 1 Then
        HLTxt txtFindWhat
        Exit Sub
    End If
    
    'set values for searching
    If chkMultiSelect.Value = vbChecked Then
        tmpMultiSelect = True
    Else
        tmpMultiSelect = False
    End If
    
    If chkInverseSelection.Value = vbChecked Then
        tmpInverseSelection = True
    Else
        tmpInverseSelection = False
    End If
    
    'execute find
    FindLVItem tmpLV, txtFindWhat.Text, , tmpMultiSelect, tmpInverseSelection

    
    'close this form to call return
    Unload Me
End Sub

