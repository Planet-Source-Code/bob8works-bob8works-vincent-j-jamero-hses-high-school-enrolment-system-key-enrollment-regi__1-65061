VERSION 5.00
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmDeleteSection 
   BackColor       =   &H00D8E9EC&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Section"
   ClientHeight    =   2835
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   5055
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   189
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   337
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtSection 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1110
      MaxLength       =   20
      TabIndex        =   1
      Top             =   735
      Width           =   3375
   End
   Begin VB.CommandButton cmdGetSection 
      BackColor       =   &H00D8E9EC&
      Height          =   330
      Left            =   4470
      Picture         =   "frmDeleteSection.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   735
      Width           =   345
   End
   Begin HSES.b8Line b8Line1 
      Height          =   60
      Left            =   0
      TabIndex        =   2
      Top             =   510
      Width           =   5385
      _ExtentX        =   9499
      _ExtentY        =   106
   End
   Begin HSES.b8Line b8Line2 
      Height          =   60
      Left            =   0
      TabIndex        =   3
      Top             =   2340
      Width           =   5385
      _ExtentX        =   9499
      _ExtentY        =   106
   End
   Begin HSES.b8Line b8Line3 
      Height          =   60
      Left            =   0
      TabIndex        =   4
      Top             =   1140
      Width           =   5385
      _ExtentX        =   9499
      _ExtentY        =   106
      BorderColor1    =   12307149
      BorderColor2    =   14215660
      BorderColor3    =   14215660
      BorderStyle1    =   3
   End
   Begin lvButton.lvButtons_H cmdDelete 
      Default         =   -1  'True
      Height          =   360
      Left            =   3570
      TabIndex        =   9
      Top             =   2430
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   635
      Caption         =   "&Delete"
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
      Focus           =   0   'False
      cGradient       =   14215660
      Gradient        =   3
      Mode            =   0
      Value           =   0   'False
      Enabled         =   0   'False
      cBack           =   16185592
   End
   Begin lvButton.lvButtons_H cmdCancel 
      Cancel          =   -1  'True
      Height          =   360
      Left            =   2040
      TabIndex        =   10
      Top             =   2430
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   635
      Caption         =   "&Cancel"
      CapAlign        =   2
      BackStyle       =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cBhover         =   14215660
      Focus           =   0   'False
      cGradient       =   14215660
      Gradient        =   3
      Mode            =   0
      Value           =   0   'False
      cBack           =   16185592
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Title:"
      Height          =   195
      Left            =   210
      TabIndex        =   8
      Top             =   765
      Width           =   360
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Delete Section"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   345
      Left            =   90
      TabIndex        =   7
      Top             =   150
      Width           =   2070
   End
   Begin VB.Label lblAskMsg 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Are You Sure Want To Delete This Record?"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Left            =   195
      TabIndex        =   6
      Top             =   2055
      Visible         =   0   'False
      Width           =   3705
   End
   Begin VB.Label lblDetail 
      BackStyle       =   0  'Transparent
      Caption         =   "Section not found!"
      ForeColor       =   &H00000000&
      Height          =   690
      Left            =   195
      TabIndex        =   5
      Top             =   1290
      Width           =   4755
   End
   Begin VB.Image Image2 
      Height          =   525
      Left            =   30
      Picture         =   "frmDeleteSection.frx":058A
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5295
   End
End
Attribute VB_Name = "frmDeleteSection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim RecordDeleted As Boolean

Dim curSection As String


Public Function ShowForm(Optional sSectionTitle As String = "") As Boolean
    
    
    '-------------------------------------------------------
    'check user access
    '-------------------------------------------------------
    If UserAllowedTo(CurrentUser.UserName, sCanDeleteSection) = False Then
        MsgBox "Unable to continue Deleteing Section entry." & vbNewLine & vbNewLine & _
                "You are not permitted to do this. Please contact your administrator for more information.", vbExclamation
        Exit Function
    End If
    '-------------------------------------------------------
    
    
    curSection = sSectionTitle
    
    Me.Show vbModal
    
    ShowForm = RecordDeleted
End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdDelete_Click()
    
    Dim vSection As tSection
    
    If GetSectionByFullTitle(txtSection.Text, vSection) = Success Then
        If DeleteSection(vSection.SectionID) = Success Then
            MsgBox "Section record successfully deleted.", vbInformation
            
            RecordDeleted = True
            Unload Me
        Else
            MsgBox "Unable to delete Section Record!", vbExclamation
        End If
    Else
        MsgBox "Section ID Not found.", vbCritical
    End If
        
End Sub

Private Sub cmdGetSection_Click()
    Dim sSection As String
    
    PickSection.GetSectionID txtSection, , , , sSection
    
    If sSection <> "" Then
        txtSection.Text = sSection
    End If
    
End Sub


Private Sub Form_Activate()
    txtSection.Text = curSection
End Sub

Private Sub txtSection_Change()
    
    cmdDelete.Enabled = False: lblAskMsg.Visible = False

    'delay 0.3 second
    'code by: VIncent J. Jamero
    '------------------------------------------------
    Static DelayStart As Single
    Static notFirst As Boolean
    DelayStart = GetTickCount + 300
    If notFirst = True Then Exit Sub
    notFirst = True
    While GetTickCount < DelayStart
        DoEvents
    Wend
    notFirst = False
    '------------------------------------------------
    'the next line will be if executed if user pause typing in 0.3 second


    ShowSectionDetail
End Sub

Private Sub ShowSectionDetail()
    
    Dim sSQL As String
    Dim vRS As New ADODB.Recordset
    
    sSQL = "SELECT Count(tblSectionOffering.SectionOfferingID) AS CountOfSectionOfferingID" & _
            " FROM tblYearLevel INNER JOIN (tblSection LEFT JOIN tblSectionOffering ON tblSection.SectionID = tblSectionOffering.SectionID) ON tblYearLevel.YearLevelID = tblSection.YearLevelID" & _
            " WHERE ((([tblYearLevel]![YearLevelTitle] & ' - ' & [tblSection]![SectionTitle])='" & txtSection.Text & "'));"

    If SectionExistByFullTitle(txtSection.Text) <> Success Then
        lblDetail.Caption = "Section not found!"
        cmdDelete.Enabled = False: lblAskMsg.Visible = False

        Exit Sub
    End If
    
    If ConnectRS(HSESDB, vRS, sSQL) = True Then
        If AnyRecordExisted(vRS) = True Then
            If ReadField(vRS.Fields("CountOfSectionOfferingID")) < 1 Then
                
                lblDetail.Caption = "Ready to delete this record."
                
                cmdDelete.Enabled = True: lblAskMsg.Visible = True
            Else
            
                lblDetail.Caption = "This Section Record cannot be deleted." & vbNewLine & _
                "Reason: This record contain " & ReadField(vRS.Fields("CountOfSectionOfferingID")) & " Section Offering record/s."
                
                cmdDelete.Enabled = False: lblAskMsg.Visible = False
            End If
        Else
            lblDetail.Caption = "Section not found."

            'Section not found
            cmdDelete.Enabled = False: lblAskMsg.Visible = False
        End If
    Else
        'fatal error
        CatchError "frmDeleteSection", "ShowSectionDetail", "Error connecting Section RS"
    End If
    
    Set vRS = Nothing
End Sub



