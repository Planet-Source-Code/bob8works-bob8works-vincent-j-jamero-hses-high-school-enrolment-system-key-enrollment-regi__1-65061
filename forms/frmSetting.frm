VERSION 5.00
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "COMCT232.OCX"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmSetting 
   BackColor       =   &H00D8E9EC&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Settings"
   ClientHeight    =   2415
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5445
   Icon            =   "frmSetting.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   161
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   363
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00D8E9EC&
      Caption         =   "Security"
      ForeColor       =   &H00C00000&
      Height          =   1050
      Left            =   75
      TabIndex        =   5
      Top             =   675
      Width           =   5220
      Begin ComCtl2.UpDown udAppSet_LockTimeOut 
         Height          =   330
         Left            =   526
         TabIndex        =   6
         Top             =   525
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   582
         _Version        =   327681
         Value           =   1
         AutoBuddy       =   -1  'True
         BuddyControl    =   "lblAppSet_LockTimeOut"
         BuddyDispid     =   196619
         OrigLeft        =   780
         OrigTop         =   525
         OrigRight       =   1035
         OrigBottom      =   840
         Max             =   1440
         Min             =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65537
         Enabled         =   -1  'True
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "minute/s"
         Height          =   330
         Left            =   930
         TabIndex        =   9
         Top             =   555
         Width           =   840
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Lock HSES when I did not move the mouse within"
         Height          =   330
         Left            =   195
         TabIndex        =   8
         Top             =   285
         Width           =   3705
      End
      Begin VB.Label lblAppSet_LockTimeOut 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "1"
         Height          =   330
         Left            =   195
         TabIndex        =   7
         Top             =   525
         Width           =   330
      End
   End
   Begin HSES.b8Line b8Line1 
      Height          =   60
      Left            =   15
      TabIndex        =   0
      Top             =   510
      Width           =   5385
      _ExtentX        =   9499
      _ExtentY        =   106
   End
   Begin lvButton.lvButtons_H cmdSave 
      Default         =   -1  'True
      Height          =   360
      Left            =   3975
      TabIndex        =   1
      Top             =   2025
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
      cGradient       =   14215660
      Gradient        =   3
      Mode            =   0
      Value           =   0   'False
      cBack           =   16777215
   End
   Begin lvButton.lvButtons_H cmdCancel 
      Cancel          =   -1  'True
      Height          =   360
      Left            =   2490
      TabIndex        =   2
      Top             =   2025
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
      Left            =   30
      TabIndex        =   3
      Top             =   1935
      Width           =   5385
      _ExtentX        =   9499
      _ExtentY        =   106
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Application Settings"
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
      TabIndex        =   4
      Top             =   120
      Width           =   2895
   End
   Begin VB.Image Image2 
      Height          =   525
      Left            =   15
      Picture         =   "frmSetting.frx":000C
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5385
   End
End
Attribute VB_Name = "frmSetting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Function ShowForm()
    
    Me.Show vbModal
End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    'save settings
    AppSet_SetLockTimeOut (udAppSet_LockTimeOut.Value * 60)
    
    'refresh setting variables
    GetAppSettings
End Sub

Private Sub Form_Activate()
    Call GetAllSettings
End Sub

Private Function GetAllSettings()
    Dim iAppSet_GetLockTimeOut As Integer
    
    iAppSet_GetLockTimeOut = AppSet_GetLockTimeOut
    If iAppSet_GetLockTimeOut < 60 Then
        iAppSet_GetLockTimeOut = 60
    End If
    udAppSet_LockTimeOut.Value = iAppSet_GetLockTimeOut \ 60
End Function


