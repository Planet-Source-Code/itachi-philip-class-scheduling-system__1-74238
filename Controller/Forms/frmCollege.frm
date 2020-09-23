VERSION 5.00
Begin VB.Form frmCourse 
   BackColor       =   &H00D8E9EC&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Course"
   ClientHeight    =   5625
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   6735
   Icon            =   "frmCollege.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5625
   ScaleWidth      =   6735
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtCurriculum 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   240
      MaxLength       =   255
      TabIndex        =   12
      Top             =   3480
      Width           =   6225
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00D8E9EC&
      Caption         =   "Cancel"
      Height          =   375
      Left            =   5400
      TabIndex        =   9
      Top             =   5160
      Width           =   1215
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00D8E9EC&
      Caption         =   "Save"
      Height          =   375
      Left            =   3960
      TabIndex        =   8
      Top             =   5160
      Width           =   1335
   End
   Begin VB.CheckBox chkDiploma 
      BackColor       =   &H00D8E9EC&
      Caption         =   "Diploma"
      Height          =   255
      Left            =   1920
      TabIndex        =   7
      Top             =   4680
      Width           =   1455
   End
   Begin VB.Frame frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00D8E9EC&
      ForeColor       =   &H80000008&
      Height          =   5055
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   6495
      Begin VB.TextBox txtCourseID 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   120
         MaxLength       =   255
         TabIndex        =   19
         Top             =   480
         Width           =   2505
      End
      Begin VB.TextBox txtYears 
         Height          =   345
         Left            =   120
         MaxLength       =   50
         TabIndex        =   17
         Top             =   4080
         Width           =   1185
      End
      Begin VB.CommandButton cmdGetCollege 
         Appearance      =   0  'Flat
         BackColor       =   &H00D8E9EC&
         Height          =   345
         Left            =   6000
         Picture         =   "frmCollege.frx":492A
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   2280
         Width           =   345
      End
      Begin VB.TextBox txtCollege 
         Appearance      =   0  'Flat
         Height          =   345
         Left            =   120
         MaxLength       =   255
         TabIndex        =   15
         Top             =   2280
         Width           =   5865
      End
      Begin VB.CommandButton cmdGetItem 
         Appearance      =   0  'Flat
         BackColor       =   &H00D8E9EC&
         Height          =   345
         Left            =   6000
         Picture         =   "frmCollege.frx":4EB4
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   2880
         Width           =   345
      End
      Begin VB.TextBox txtDepartmentTitle 
         Appearance      =   0  'Flat
         Height          =   345
         Left            =   120
         MaxLength       =   255
         TabIndex        =   10
         Top             =   2880
         Width           =   5865
      End
      Begin VB.CheckBox chkOffered 
         BackColor       =   &H00D8E9EC&
         Caption         =   "Currently offered"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   4680
         Width           =   1455
      End
      Begin VB.TextBox txtMajor 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   120
         MaxLength       =   255
         TabIndex        =   4
         Top             =   1680
         Width           =   6225
      End
      Begin VB.TextBox txtCourse 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   120
         MaxLength       =   255
         TabIndex        =   2
         Top             =   1080
         Width           =   6225
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Course Reference #:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   2535
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Total Years"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   18
         Top             =   3840
         Width           =   1095
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "College"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   14
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Curriculum"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   13
         Top             =   3240
         Width           =   1095
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Department"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   5
         Top             =   2640
         Width           =   1095
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Major"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   3
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Course"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   1
         Top             =   840
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmCourse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim RecordAdded As Boolean
Dim curCourse As tCourse
Public mFormState As String

Dim CurrentCourse As tCourse

Dim RecordEdited As Boolean
Dim CourseID, curDepartment, CollegeID As String



Private Sub cmdGetCollege_Click()
 Dim sCollegeTitle As String
    Dim sCollegeID As String
    
    sCollegeID = frmPickCollege.GetItem(txtCollege, sCollegeTitle)
    If sCollegeID <> "" Then
        txtCollege.Text = sCollegeTitle
        CollegeID = sCollegeID
    End If
End Sub

Private Sub cmdGetItem_Click()
  Dim sDepartmentTitle As String
    Dim sDepartmentID As String
    
    sDepartmentID = frmPickDepartment.GetItem(txtDepartmentTitle, sDepartmentTitle)
    If sDepartmentID <> "" Then
        txtDepartmentTitle = sDepartmentTitle
        curDepartment = sDepartmentID
    End If
End Sub

Public Function ShowEdit(sCourseID As String) As Boolean
Dim vDepartment As tDepartment
Dim vCollege As tCollege
    
        If GetCourseByTitle(sCourseID, curCourse) <> Success Then
            MsgBox "Unable to continue editing Course Information: Course Reference # not found!", vbCritical
            Exit Function
        End If
        
        If GetDepartmentByID(curCourse.DepartmentID, vDepartment) <> Success Then
            Exit Function
        End If
        
        If GetCollegeByID(curCourse.CollegeID, vCollege) <> Success Then
            Exit Function
        End If
    
        
    txtCourseID.Text = curCourse.CourseID
    txtCourse.Text = curCourse.CourseTitle
    txtMajor.Text = curCourse.Major
    txtCurriculum.Text = curCourse.Curriculum
    chkOffered.Value = Abs(curCourse.CurrentOffered)
    chkDiploma.Value = Abs(curCourse.Diploma)
    CollegeID = curCourse.CollegeID
    txtYears.Text = curCourse.Years
    txtCollege.Text = vCollege.CollegeTitle
    txtDepartmentTitle = vDepartment.DepartmentTitle
    curDepartment = vDepartment.DepartmentID
    
    mFormState = "EDIT"
    
    Me.Show vbModal

    ShowEdit = RecordEdited
    
    
End Function

Public Function ShowProperties(sCourse As String) As Boolean
Dim vDepartment As tDepartment
Dim vCollege As tCollege
    
        If GetCourseByID(txtCourseID.Text, curCourse) <> Success Then
            MsgBox "Unable to continue editing Course Information: Course Reference # not found!", vbCritical
            Exit Function
        End If
        
        If GetDepartmentByID(curCourse.DepartmentID, vDepartment) <> Success Then
            Exit Function
        End If
        
        If GetCollegeByID(curCourse.CollegeID, vCollege) <> Success Then
            Exit Function
        End If
        
        
    txtCourseID.Text = curCourse.CourseID
    txtCourse.Text = curCourse.CourseTitle
    txtMajor.Text = curCourse.Major
    txtCurriculum.Text = curCourse.Curriculum
    chkOffered.Value = Abs(curCourse.CurrentOffered)
    chkDiploma.Value = Abs(curCourse.Diploma)
    CollegeID = curCourse.CollegeID
    txtYears.Text = curCourse.Years
    txtCollege.Text = vCollege.CollegeTitle
    txtDepartmentTitle = vDepartment.DepartmentTitle
    curDepartment = vDepartment.DepartmentID
    
    mFormState = "EDIT"
    
    Me.Show vbModal

    ShowProperties = RecordEdited
      
End Function
Public Function ShowForm() As Boolean
    
    Dim sNewID As String
    
    If GetNewCourseID(sNewID) = Failed Then
        CatchError "Course", "ShowForm()", "GetNewCourseID(sNewID) = Failed"
        Exit Function
    End If

    CourseID = sNewID
    
    txtCourseID.Text = CourseID
    
    mFormState = "ADD"

    Me.Show vbModal
    'return
    ShowForm = RecordAdded
    
End Function


Private Sub cmdCancel_Click()
    Unload Me
End Sub


Private Function SaveNewDepartment()
    
    If Not CheckTextBox(txtDepartmentTitle, "Please enter Department Title") Then
        Exit Function
    End If
    
    
    'save
    Dim newCourse As tCourse
    
    newCourse.CourseID = txtCourseID.Text
    newCourse.CollegeID = CollegeID
    newCourse.CourseTitle = txtCourse.Text
    newCourse.CurrentOffered = chkOffered.Value
    newCourse.Curriculum = txtCurriculum.Text
    newCourse.DepartmentID = curDepartment
    newCourse.Diploma = chkDiploma.Value
    newCourse.Major = txtMajor.Text
    newCourse.Years = txtYears.Text
    
    Select Case AddCourse(newCourse)
        Case TranDBResult.Success  'success
            MsgBox "New Course successfully added", vbInformation
            RecordAdded = True
            Unload Me
            
        Case TranDBResult.DuplicateID
            MsgBox "Invalid Course ID!" & vbNewLine & "The Course ID that you have entered is already existed. Enter another Department ID.", vbExclamation
            txtCourseID.SetFocus
            
        Case TranDBResult.DuplicateTitle
        
            MsgBox "Invalid Course Title!" & vbNewLine & "The Course Title that you have entered is already existed. Enter another Department Title.", vbExclamation
            HLTxt txtDepartmentTitle
            
        Case Else
            MsgBox "Unknown Error", vbExclamation
            CatchError "frmAddCourse", "SaveNewCourse", "Unknown result in Add New Department"
    End Select
    
    
End Function

Private Sub cmdSave_Click()
If mFormState = "ADD" Then
    SaveNewDepartment
Else
    UpdateData
End If
End Sub

Private Function UpdateData()
    
    If Not CheckTextBox(txtCourse, "Enter Course Title." & vbNewLine & " This field is required") Then
        Exit Function
    End If
    
    If Not CheckTextBox(txtMajor, "Enter Course Major." & vbNewLine & " This field is required") Then
        Exit Function
    End If
    
    If Not CheckTextBox(txtDepartmentTitle, "Enter Department Title." & vbNewLine & " This field is required") Then
        Exit Function
    End If
    
    If Not CheckTextBox(txtCollege, "Enter College Title." & vbNewLine & " This field is required") Then
        Exit Function
    End If
    
    If Not CheckTextBox(txtCurriculum, "Enter Curriculum ." & vbNewLine & " This field is required") Then
        Exit Function
    End If
    
    If Not CheckTextBox(txtYears, "Enter Years ." & vbNewLine & " This field is required") Then
        Exit Function
    End If
    
    Dim newCourse As tCourse
    
    Dim EditResult As Integer
    
    newCourse.CourseID = txtCourseID.Text
    newCourse.CollegeID = CollegeID
    newCourse.CourseTitle = txtCourse.Text
    newCourse.CurrentOffered = chkOffered.Value
    newCourse.Curriculum = txtCurriculum.Text
    newCourse.DepartmentID = curDepartment
    newCourse.Diploma = chkDiploma.Value
    newCourse.Major = txtMajor.Text
    newCourse.Years = txtYears.Text
    newCourse.DepartmentID = curDepartment
    
    Select Case EditCourse(newCourse)
        Case Success
        
            MsgBox "Course Information was successfully edited", vbInformation
            
            RecordEdited = True
        
            Unload Me
            
        Case DuplicateTitle
            MsgBox "The Course Title that you have entered was already existed." & vbNewLine & " Enter another Duplicate Title", vbExclamation
            HLTxt txtCourse

        Case Else
            MsgBox "UNKNOWN: Editing Course", vbCritical
    End Select
End Function

