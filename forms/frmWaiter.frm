VERSION 5.00
Begin VB.Form frmWaiter 
   BorderStyle     =   0  'None
   ClientHeight    =   3090
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
End
Attribute VB_Name = "frmWaiter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim FRM As Form


Public Sub StartWaiting()
    
    
    Set FRM = mdiMain.ActiveForm
    
    Me.Show
    
    FRM.WindowState = vbMaximized
    FRM.SetFocus
End Sub

Private Sub Form_Activate()
    MsgBox mdiMain.Count
End Sub

