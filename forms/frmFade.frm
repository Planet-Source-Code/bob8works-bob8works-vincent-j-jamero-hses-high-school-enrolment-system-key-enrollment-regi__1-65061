VERSION 5.00
Object = "{24F3B0F0-7086-439E-8A82-F7E7B7CA5762}#1.0#0"; "XandersTool.ocx"
Begin VB.Form frmFade 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin XandersTool.XandersXPTransparency XandersXPTransparency1 
      Height          =   465
      Left            =   1860
      TabIndex        =   0
      Top             =   1710
      Visible         =   0   'False
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   820
      TransparencyLevel=   80
   End
End
Attribute VB_Name = "frmFade"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim curFormOwner As Form
Dim KILLME As Boolean

Public Function ShowForm(ByRef FormOwner As Form, Optional iTransparency As Integer = 20, Optional BGColor As OLE_COLOR = &H0&)

    KILLME = False
    
    Set curFormOwner = FormOwner
    
    Me.BackColor = BGColor
    XandersXPTransparency1.TransparencyLevel = iTransparency
    Me.Move curFormOwner.Left - 100, curFormOwner.Top - 100, curFormOwner.Width + 200, curFormOwner.Height + 200
    
    Me.Show ' vbModal
    
    Set curFormOwner = Nothing
    
End Function

Private Sub Form_Activate()
    
    Static isFirst As Boolean
    
    If isFirst = False Then
        curFormOwner.Show ' vbModal
        isFirst = True
        WathForm
    Else
        KILLME = True
        Unload Me
        isFirst = False
    End If
End Sub

Private Sub WathForm()
    While KILLME = False
    
    Static R As RECT
    Static i As Integer

    
    If R.Left <> curFormOwner.Left Or _
            R.Right <> curFormOwner.Left + curFormOwner.Width Or _
            R.Top <> curFormOwner.Top Or _
            R.Bottom <> curFormOwner.Top + curFormOwner.Height Then
    
        i = 100
    End If

    If i > 80 Then
        i = i - 2
    End If
    
    

    XandersXPTransparency1.MakeSemiTransparent Me.hwnd, i
    
    Me.Move curFormOwner.Left - 100, curFormOwner.Top - 100, curFormOwner.Width + 200, curFormOwner.Height + 200
    DoEvents
    
    R.Left = curFormOwner.Left
    R.Right = curFormOwner.Left + curFormOwner.Width
    R.Top = curFormOwner.Top
    R.Bottom = curFormOwner.Top + curFormOwner.Height

    Wend
End Sub
