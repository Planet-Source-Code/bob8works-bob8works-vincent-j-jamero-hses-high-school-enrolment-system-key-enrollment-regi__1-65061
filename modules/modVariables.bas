Attribute VB_Name = "modVariables"

Public Const defTxtBoxBorderColor = &HAD705A
Public Const defTxtBoxBackColor = vbWhite

Public Const hlTxtBoxBorderColor = &HFF&
Public Const hlTxtBoxBackColor = &HE3E3FF



Public DBPathFileName As String

Public UserExisted As Boolean
Public SchoolExisted As Boolean




Public CurrentUser As User
Public CurrentSchool As School



Public CurrentSchoolYear As tSchoolYear

Public kSelectListKey  As Integer

Public SYOriginalFilePath As String

Public Enum tFormState
    Ready = 0
    Searching = 1
    ReadingRecord = 2
    Adding = 12
    Editing = 13
    Deleting = 14
End Enum


Public Const Form_TopMargin = 4
Public Const Form_LeftMargin = 4



Public MinGradeForI As Double
Public MinGradeForII As Double
Public MinGradeForIII As Double
Public MinGradeForIV As Double
Public MinGradeForGraduate As Double



'settings
Public AppSet_LockTimeOut As Integer


Public currentUserLog As LogRec

