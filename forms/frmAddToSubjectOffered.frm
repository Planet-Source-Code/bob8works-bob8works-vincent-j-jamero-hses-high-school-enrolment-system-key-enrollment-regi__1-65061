VERSION 5.00
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmAddToSubjectOffered 
   BackColor       =   &H00D8E9EC&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Add Subject Offered"
   ClientHeight    =   2610
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4410
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   174
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   294
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cmbDays 
      Height          =   315
      Left            =   1110
      TabIndex        =   13
      Top             =   1110
      Width           =   3075
   End
   Begin VB.TextBox txtTimeEnd 
      Height          =   330
      Left            =   1920
      MaxLength       =   4
      TabIndex        =   6
      Text            =   "0800"
      Top             =   1470
      Width           =   615
   End
   Begin VB.TextBox txtTimeStart 
      Height          =   330
      Left            =   1110
      MaxLength       =   4
      TabIndex        =   5
      Text            =   "0700"
      Top             =   1470
      Width           =   615
   End
   Begin VB.CommandButton cmdGenTeacher 
      BackColor       =   &H00D8E9EC&
      Height          =   270
      Left            =   3810
      Picture         =   "frmAddToSubjectOffered.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   750
      Width           =   345
   End
   Begin VB.TextBox txtTeacherFullName 
      Height          =   330
      Left            =   1110
      MaxLength       =   50
      TabIndex        =   2
      Top             =   720
      Width           =   3075
   End
   Begin VB.TextBox txtSubjectTitle 
      Height          =   330
      Left            =   1110
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   0
      Top             =   330
      Width           =   3075
   End
   Begin lvButton.lvButtons_H cmdOk 
      Default         =   -1  'True
      Height          =   360
      Left            =   2940
      TabIndex        =   9
      Top             =   2220
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   635
      Caption         =   "&Ok"
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
   Begin lvButton.lvButtons_H cmdCancel 
      Cancel          =   -1  'True
      Height          =   360
      Left            =   1470
      TabIndex        =   10
      Top             =   2220
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
   Begin HSES.b8Line b8Line2 
      Height          =   60
      Left            =   -240
      TabIndex        =   11
      Top             =   2130
      Width           =   5385
      _ExtentX        =   9499
      _ExtentY        =   106
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1770
      TabIndex        =   12
      Top             =   1470
      Width           =   105
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Schedule"
      Height          =   195
      Left            =   180
      TabIndex        =   8
      Top             =   1140
      Width           =   675
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Time"
      Height          =   195
      Left            =   180
      TabIndex        =   7
      Top             =   1530
      Width           =   345
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Teacher"
      Height          =   195
      Left            =   150
      TabIndex        =   3
      Top             =   750
      Width           =   600
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Subject Title"
      Height          =   195
      Left            =   150
      TabIndex        =   1
      Top             =   360
      Width           =   885
   End
End
Attribute VB_Name = "frmAddToSubjectOffered"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim RecordSaved As Boolean


Dim curTeacherID As String
Dim curDays As String
Dim curSchedTimeEnd As String
Dim curSchedTimeStart As String
Dim curTeacherName As String

Dim curTime() As String

Public Function ShowForm(sSectionID As String, sSubjectID As String, sSubjectTitle As String, curTimeAc() As String, _
        ByRef sSubjectOfferingID As String, _
        ByRef sSchedTimeStart As String, _
        ByRef sSchedTimeEnd As String, _
        ByRef sTeacherID As String, _
        ByRef sDays As String, _
        ByRef sTeacherName As String) As Boolean
    
    
    Dim i As Integer
    Dim splitTime() As String
    Dim cTimeOut As Integer
    
    'set default
    RecordSaved = False
    sTeacherID = ""
    
    'set parameter
    
    
    txtSubjectTitle.Text = sSubjectTitle
    
    ReDim curTime(UBound(curTimeAc))
    
    cTimeOut = 700
    For i = 0 To UBound(curTimeAc)
        curTime(i) = curTimeAc(i)
        
        splitTime = Split(curTime(i), "-")
        If IsNumeric(splitTime(2)) = True Then
            If cTimeOut < Val(splitTime(2)) Then
                cTimeOut = Val(splitTime(2))
            End If
        End If
        
    Next
    

    
    cTimeOut = cTimeOut Mod 2400
    
    txtTimeStart.Text = Left("0000", 4 - Len(Trim(Str(cTimeOut)))) & cTimeOut
    cTimeOut = cTimeOut + 100
    cTimeOut = cTimeOut Mod 2400
    txtTimeEnd.Text = Left("0000", 4 - Len(Trim(Str(cTimeOut)))) & cTimeOut
    
    'show form
    Me.Show vbModal
    
    'return
    sTeacherID = curTeacherID
    sSchedTimeStart = curSchedTimeStart
    sSchedTimeEnd = curSchedTimeEnd
    sDays = curDays
    sTeacherName = curTeacherName
    
    
    ShowForm = RecordSaved
End Function

Private Sub cmbDays_LostFocus()
    If cmbDays.ListIndex < 0 Then
        cmbDays.ListIndex = 0
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdGenTeacher_Click()
    Dim sTeacherFullName As String
    Dim sTeacherID As String
    
    sTeacherID = PickTeacher.GetTeacherID(sTeacherFullName)
    
    If sTeacherFullName <> "" Then
        curTeacherID = sTeacherID
        curTeacherName = sTeacherFullName
        txtTeacherFullName.Text = sTeacherFullName
    End If
    
End Sub




Private Sub cmdOK_Click()

    

    If CheckTimeStart = False Then
        Exit Sub
    End If
    If CheckTimeEnd = False Then
        Exit Sub
    End If
    
    If CheckTime = False Then
        Exit Sub
    End If
    
    
    If Len(curTeacherID) < 1 Then
        MsgBox "Please enter Teacher", vbExclamation
        Exit Sub
    End If

    
    curDays = cmbDays.Text
    curSchedTimeStart = txtTimeStart.Text
    curSchedTimeEnd = txtTimeEnd.Text
    
    'set flag
    RecordSaved = True
    
    Unload Me
End Sub

Private Function CheckTime() As Boolean
    
    Dim i As Integer
    Dim iTimeIn As Integer
    Dim iTimeOut As Integer
    
    Dim splitTime() As String
    Dim cTimeIn As Integer
    Dim cTimeOut As Integer
    Dim cDays As String
    
    CheckTime = False
    
    If IsNumeric(txtTimeStart.Text) Then
        iTimeIn = Val(txtTimeStart.Text)
    Else
        Exit Function
    End If
    If IsNumeric(txtTimeEnd.Text) Then
        iTimeOut = Val(txtTimeEnd.Text)
    Else
        Exit Function
    End If
    
    
    For i = 0 To UBound(curTime)
        splitTime = Split(curTime(i), "-")
        If IsNumeric(splitTime(1)) And IsNumeric(splitTime(2)) = True Then
            
            cDays = splitTime(0)
            cTimeIn = Val(splitTime(1))
            cTimeOut = Val(splitTime(2))
            
            'check if the same time
            If FindNoCharMatch(cmbDays.Text, cDays) = False Then
            
                If cTimeIn <= iTimeIn And iTimeIn < cTimeOut Then
                    MsgBox "Invalid Time Schedule." & vbNewLine & "There is a conflict in Time Start.", vbExclamation
                    HLTxt txtTimeStart
                    Exit Function
                End If
                
                If cTimeIn < iTimeOut And iTimeOut < cTimeOut Then
                    MsgBox "Invalid Time Schedule." & vbNewLine & "There is a conflict in Time End.", vbExclamation
                    HLTxt txtTimeEnd
                    Exit Function
                End If
                
                If cTimeIn <= iTimeIn And cTimeOut >= iTimeOut Then
                    MsgBox "Invalid Time Schedule." & vbNewLine & "There is a conflict.", vbExclamation
                    HLTxt txtTimeStart
                    Exit Function
                End If
                
                If iTimeIn <= cTimeIn And iTimeOut >= cTimeOut Then
                    MsgBox "Invalid Time Schedule." & vbNewLine & "There is a conflict.", vbExclamation
                    HLTxt txtTimeStart
                    Exit Function
                End If
                
            End If
            
        Else
            MsgBox "Invalid TIme"
            Exit Function
        End If
    Next
    
    CheckTime = True
End Function


Private Function FindNoCharMatch(Str1 As String, Str2 As String) As Boolean
    
    Dim i As Integer
    Dim sC As String
    
    'default
    FindNoCharMatch = False

    'check the first stirng
    For i = 1 To Len(Str1)
        sC = Mid(Str1, i, 1)
    
        If InStr(1, Str2, sC) > 0 Then
            'found
            Exit Function
        End If
    Next
    
    'check the second stirng
    For i = 1 To Len(Str2)
        sC = Mid(Str2, i, 1)
    
        If InStr(1, Str1, sC) > 0 Then
            'found
            Exit Function
        End If
    Next
    
    
    'return success
    FindNoCharMatch = True
End Function


Private Sub Form_Load()
    
    cmbDays.Clear
    
    cmbDays.AddItem "MTWHF"
    cmbDays.AddItem "M"
    cmbDays.AddItem "T"
    cmbDays.AddItem "W"
    cmbDays.AddItem "H"
    cmbDays.AddItem "F"
    cmbDays.AddItem "S"
    cmbDays.AddItem "U"
    
    cmbDays.AddItem "MTW"
    cmbDays.AddItem "WHF"
    cmbDays.AddItem "MT"
    cmbDays.AddItem "WH"
    cmbDays.AddItem "FS"
    
    
    'set first item as default active item
    cmbDays.ListIndex = 0
End Sub


Private Sub txtTeacherFullName_KeyPress(KeyAscii As Integer)
    Call cmdGenTeacher_Click
End Sub

Private Sub txtTimeEnd_LostFocus()

    CheckTimeEnd
    
End Sub

Public Function CheckTimeEnd() As Boolean
    
    'default
    CheckTimeEnd = False

    If IsNumeric(txtTimeEnd.Text) = False Then
        MsgBox "Time Start must be in Millitary Time (ex: 1700,0100) format.", vbExclamation
        HLTxt txtTimeEnd
        
        Exit Function
    End If
    
    If Val(txtTimeEnd.Text) < 0 Or Val(txtTimeEnd.Text) >= 2400 Then
        MsgBox "Time Start must be in this range: 0001 to 2399.", vbExclamation
        HLTxt txtTimeEnd
        
        Exit Function
    End If
    
    If Not (Val(txtTimeStart.Text) < Val(txtTimeEnd.Text)) Then
        
        MsgBox "Time End must be less than Time Start"
        HLTxt txtTimeEnd
        
        Exit Function
    End If
    
    'return true
    CheckTimeEnd = True
End Function

Private Sub txtTimeStart_LostFocus()
    
    CheckTimeStart
End Sub

Public Function CheckTimeStart() As Boolean
    
    Dim iNewTimeEnd As Integer
    'default
    CheckTimeStart = False
    
    If IsNumeric(txtTimeStart.Text) = False Then
        MsgBox "Time Start must be in Millitary Time (ex: 1700,0100) format.", vbExclamation
        HLTxt txtTimeStart
        
        Exit Function
    End If
    
    If Val(txtTimeStart.Text) < 0 Or Val(txtTimeStart.Text) >= 2400 Then
        MsgBox "Time Start must be in this range: 0001 to 2399.", vbExclamation
        HLTxt txtTimeStart
        
        Exit Function
    End If
    
    If IsNumeric(txtTimeEnd.Text) = True Then
        If Val(txtTimeStart.Text) >= Val(txtTimeEnd.Text) Then
            iNewTimeEnd = Val(txtTimeStart.Text) + 100
            txtTimeEnd.Text = Left("0000", 4 - Len(Trim(Str(iNewTimeEnd)))) & iNewTimeEnd
            Exit Function
        End If
    End If
    
    'return true
    CheckTimeStart = True
    
    CheckTimeEnd
End Function
