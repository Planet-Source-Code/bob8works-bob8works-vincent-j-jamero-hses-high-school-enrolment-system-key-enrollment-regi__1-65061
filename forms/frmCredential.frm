VERSION 5.00
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmCredential 
   BackColor       =   &H00D8E9EC&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Credential"
   ClientHeight    =   3000
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   5100
   Icon            =   "frmCredential.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   200
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   340
   StartUpPosition =   2  'CenterScreen
   Begin lvButton.lvButtons_H cmdHistory 
      Height          =   345
      Left            =   30
      TabIndex        =   11
      Top             =   2610
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   609
      Caption         =   "Entry History"
      CapAlign        =   2
      BackStyle       =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFore           =   8421504
      cFHover         =   8421504
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      Enabled         =   0   'False
      cBack           =   14215660
   End
   Begin VB.TextBox txtDescription 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      IMEMode         =   3  'DISABLE
      Left            =   1455
      MaxLength       =   100
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   1665
      Width           =   3135
   End
   Begin VB.TextBox txtCredentialID 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1470
      Locked          =   -1  'True
      MaxLength       =   4
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   900
      Width           =   3105
   End
   Begin VB.TextBox txtTitle 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1455
      MaxLength       =   50
      TabIndex        =   1
      Top             =   1290
      Width           =   3135
   End
   Begin HSES.b8Line b8Line1 
      Height          =   60
      Left            =   0
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   510
      Width           =   5385
      _ExtentX        =   9499
      _ExtentY        =   106
   End
   Begin HSES.b8Line b8Line2 
      Height          =   60
      Left            =   -45
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   2520
      Width           =   5385
      _ExtentX        =   9499
      _ExtentY        =   106
   End
   Begin lvButton.lvButtons_H cmdSave 
      Default         =   -1  'True
      Height          =   360
      Left            =   3600
      TabIndex        =   4
      Top             =   2610
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   635
      Caption         =   "&Save"
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
      cBack           =   16185592
   End
   Begin lvButton.lvButtons_H cmdCancel 
      Cancel          =   -1  'True
      Height          =   360
      Left            =   2070
      TabIndex        =   5
      Top             =   2610
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
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      Height          =   195
      Left            =   360
      TabIndex        =   2
      Top             =   1665
      Width           =   795
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Title"
      Height          =   195
      Left            =   360
      TabIndex        =   0
      Top             =   1305
      Width           =   300
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ID"
      Height          =   195
      Left            =   345
      TabIndex        =   10
      Top             =   930
      Width           =   165
   End
   Begin VB.Label lblFormTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "New Credential"
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
      TabIndex        =   9
      Top             =   180
      Width           =   2205
   End
   Begin VB.Image Image2 
      Height          =   525
      Left            =   0
      Picture         =   "frmCredential.frx":08CA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5295
   End
End
Attribute VB_Name = "frmCredential"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim RecordSaved As Boolean
Dim sFormState As String

Dim curCredentialID As String
Dim curTitle As String

Dim curCreationDate As Date
Dim curCreatedBy As String
Dim curModifiedDate As Date
Dim curModifiedBy As String


Public Function ShowAdd() As Boolean
    
    '-------------------------------------------------------
    'check user access
    '-------------------------------------------------------
    If UserAllowedTo(CurrentUser.UserName, sCanAddCredential) = False Then
        MsgBox "Unable to continue adding Credential entry." & vbNewLine & vbNewLine & _
                "You are not permitted to do this. Please contact your administrator for more information.", vbExclamation
        Exit Function
    End If
    '-------------------------------------------------------
    
    'set ui & var
    sFormState = "add"
    
    lblFormTitle = "New Credential"
    cmdSave.Caption = "&Save"
    
    
    
    'show form
    Me.Show vbModal
    
    'return
    ShowAdd = RecordSaved
End Function

Public Function ShowEdit(sCredentialID As String) As Boolean
    
    '-------------------------------------------------------
    'check user access
    '-------------------------------------------------------
    If UserAllowedTo(CurrentUser.UserName, sCanEditCredential) = False Then
        MsgBox "Unable to continue Editing Credential entry." & vbNewLine & vbNewLine & _
                "You are not permitted to do this. Please contact your administrator for more information.", vbExclamation
        Exit Function
    End If
    '-------------------------------------------------------
    
    'set ui & var
    curCredentialID = sCredentialID
    
    sFormState = "edit"
    lblFormTitle = "Edit Credential"
    cmdSave.Caption = "&Update"
    
    
    'show form
    Me.Show vbModal
    
    'return
    ShowEdit = RecordSaved
End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Function GenerateAutoID() As String
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    Dim sNewID As String
    
    sSQL = "SELECT Max(tblCredential.CredentialID) AS MaxOfCredentialID" & _
            " FROM tblCredential"

    If ConnectRS(HSESDB, vRS, sSQL) = False Then
        'temp
        'fatal
        MsgBox "error"
        GoTo ReleaseAndExit
    End If
    
    On Error Resume Next
    sNewID = "0"
    sNewID = Val(ReadField(vRS.Fields("MaxOfCredentialID"))) + 1
    If Not (Val(sNewID) > 0) Then
        sNewID = "1"
    End If
    sNewID = Left("0000", 4 - Len(Trim(sNewID))) & sNewID
    
    GenerateAutoID = sNewID
    
ReleaseAndExit:
    Set vRS = Nothing
End Function

Private Sub cmdHistory_Click()
    MsgBox "Record History" & vbNewLine & vbNewLine & _
    "   Created:  " & curCreationDate & vbNewLine & _
    "   By:       " & curCreatedBy & vbNewLine & vbNewLine & _
    "   Modified: " & IIf(Len(curModifiedBy) > 0, curModifiedDate, "") & vbNewLine & _
    "   By:       " & curModifiedBy, vbInformation, "Record History"
    
End Sub

Private Sub cmdSave_Click()
    'check form state
    Select Case sFormState
    
        Case "add"
            SaveAdd
            
        Case "edit"
            SaveEdit
    End Select
End Sub

Private Sub Form_Activate()

    'check form state
    Select Case sFormState
    
        Case "add"
            SetUIAdd
            
        Case "edit"
            SetUIEdit
    End Select
    
End Sub

Private Sub SetUIAdd()

    txtCredentialID.Text = GenerateAutoID
    
End Sub


Private Function SetUIEdit()

    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    
    sSQL = "SELECT tblCredential.CredentialID, tblCredential.Title, tblCredential.Description, tblCredential.CreationDate, tblCredential.CreatedBy, tblCredential.ModifiedDate, tblCredential.ModifiedBy" & _
            " From tblCredential" & _
            " WHERE (((tblCredential.CredentialID)='" & curCredentialID & "'))"


    If ConnectRS(HSESDB, vRS, sSQL) = False Then
        CatchError "frmCredential", "SetUIEdit", "Unable to connect Recordset with Sql Expression '" & sSQL & "'"
        MsgBox "Unable to connect Recordset.", vbExclamation
        GoTo ReleaseAndExit
    End If
    
    If AnyRecordExisted(vRS) = False Then
        CatchError "frmCredential", "SetUIEdit", "No record found when editing record with id: " & curCredentialID & " and with Sql Expression '" & sSQL & "'"
        MsgBox "Record not found.", vbExclamation
        GoTo ReleaseAndExit
    End If
    
    'set ui
    curTitle = ReadField(vRS.Fields("Title"))
    txtCredentialID.Text = ReadField(vRS.Fields("CredentialID"))
    txtTitle.Text = ReadField(vRS.Fields("Title"))
    txtDescription.Text = ReadField(vRS.Fields("Description"))
    
    curCreationDate = ReadField(vRS.Fields("CreationDate"))
    curCreatedBy = ReadField(vRS.Fields("CreatedBy"))
    curModifiedDate = ReadField(vRS.Fields("ModifiedDate"))
    curModifiedBy = ReadField(vRS.Fields("ModifiedBy"))
    
    cmdHistory.Enabled = True
    
ReleaseAndExit:
    Set vRS = Nothing
End Function






Private Sub SaveAdd()
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    
    'check
    If Len(txtTitle.Text) < 1 Then
        MsgBox "Please fill up Title.", vbExclamation
        HLTxt txtTitle
        GoTo ReleaseAndExit
    End If
    
    sSQL = "SELECT * FROM tblCredential WHERE tblCredential.CredentialID='" & txtCredentialID.Text & "' OR tblCredential.Title='" & txtTitle.Text & "'"
    
    If ConnectRS(HSESDB, vRS, sSQL) = False Then
        CatchError "frmCredential", "SaveAdd", "Unable to connect Recordset with Sql Expression '" & sSQL & "'"
        MsgBox "Unable to connect Recordset.", vbExclamation
        GoTo ReleaseAndExit
    End If
    
    If AnyRecordExisted(vRS) = True Then
        MsgBox "The Credential Title that you have entered was already in records." & vbNewLine & _
            "Please enter different Title.", vbExclamation
        HLTxt txtTitle
        GoTo ReleaseAndExit
    End If
    
    
    
    'save record
    On Error GoTo ErrSaving
    
    vRS.AddNew
        
    vRS.Fields("CredentialID").Value = txtCredentialID.Text
    vRS.Fields("Title").Value = txtTitle.Text
    vRS.Fields("Description").Value = txtDescription.Text
    
    vRS.Fields("CreatedBy").Value = CurrentUser.UserName
    vRS.Fields("CreationDate").Value = Now
        
    vRS.Update
    
    'show success message
    MsgBox "New Credential entry successfully created.", vbInformation
    'set flag
    RecordSaved = True
    'close this form
    Unload Me
    
    
ReleaseAndExit:
    Set vRS = Nothing
    Exit Sub

ErrSaving:
    CatchError "frmCredential", "SaveAdd", "Unable to update Record with Sql Expression '" & sSQL & "'"
    GoTo ReleaseAndExit
End Sub

Private Sub SaveEdit()
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    
    'check
    If Len(txtTitle.Text) < 1 Then
        MsgBox "Please fill up Title.", vbExclamation
        HLTxt txtTitle
        GoTo ReleaseAndExit
    End If
    
     'check duplicate
    If LCase(txtTitle.Text) <> LCase(curTitle) Then
        
        sSQL = "SELECT * FROM tblCredential WHERE tblCredential.Title='" & txtTitle.Text & "'"
        
        If ConnectRS(HSESDB, vRS, sSQL) = False Then
            CatchError "frmCredential", "SaveEdit", "Unable to connect Recordset with Sql Expression '" & sSQL & "'"
            MsgBox "Unable to connect Recordset.", vbExclamation
            GoTo ReleaseAndExit
        End If
        
        If AnyRecordExisted(vRS) = True Then
            MsgBox "The Credential Title that you have entered was already in records." & vbNewLine & _
                "Please enter different Title.", vbExclamation
            HLTxt txtTitle
            GoTo ReleaseAndExit
        End If
    
    End If
    
    
    sSQL = "SELECT * FROM tblCredential WHERE tblCredential.CredentialID='" & txtCredentialID.Text & "'"
    On Error Resume Next
    vRS.Close
    If ConnectRS(HSESDB, vRS, sSQL) = False Then
        CatchError "frmCredential", "SaveEdit", "Unable to connect Recordset with Sql Expression '" & sSQL & "'"
        MsgBox "Unable to connect Recordset.", vbExclamation
        GoTo ReleaseAndExit
    End If
    
    If AnyRecordExisted(vRS) = False Then
        MsgBox "Unable to update record." & vbNewLine & _
            "Record not found.", vbExclamation
        HLTxt txtTitle
        GoTo ReleaseAndExit
    End If
    
    
    
    'save record
    On Error GoTo ErrSaving
    
    'edit
        
    'vRS.Fields("CredentialID").Value = txtCredentialID.Text
    vRS.Fields("Title").Value = txtTitle.Text
    vRS.Fields("Description").Value = txtDescription.Text
    
    vRS.Fields("ModifiedBy").Value = CurrentUser.UserName
    vRS.Fields("ModifiedDate").Value = Now
        
    vRS.Update
    
    'show success message
    MsgBox "Credential entry successfully updated.", vbInformation
    'set flag
    RecordSaved = True
    'close this form
    Unload Me
    
    
ReleaseAndExit:
    Set vRS = Nothing
    Exit Sub

ErrSaving:
    CatchError "frmCredential", "SaveEdit", "Unable to update Record with Sql Expression '" & sSQL & "'"
    GoTo ReleaseAndExit
End Sub
