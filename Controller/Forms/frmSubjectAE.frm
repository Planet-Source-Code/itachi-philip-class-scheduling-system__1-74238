VERSION 5.00
Begin VB.Form frmSubjectAE 
   BackColor       =   &H00D8E9EC&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Subject"
   ClientHeight    =   6135
   ClientLeft      =   4515
   ClientTop       =   1830
   ClientWidth     =   6030
   Icon            =   "frmSubjectAE.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6135
   ScaleWidth      =   6030
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   4560
      TabIndex        =   27
      Top             =   5640
      Width           =   1335
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   375
      Left            =   3120
      TabIndex        =   26
      Top             =   5640
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00D8E9EC&
      Height          =   5535
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   5775
      Begin VB.TextBox txtSubject 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1800
         TabIndex        =   28
         Top             =   720
         Width           =   3225
      End
      Begin VB.TextBox txtRepeatFee 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1800
         TabIndex        =   24
         Top             =   5040
         Width           =   3225
      End
      Begin VB.TextBox txtLaboratoryFee 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1800
         TabIndex        =   22
         Top             =   4680
         Width           =   3225
      End
      Begin VB.TextBox txtSubjectFee 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1800
         TabIndex        =   20
         Top             =   4320
         Width           =   3225
      End
      Begin VB.TextBox txtCategory 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1800
         TabIndex        =   18
         Top             =   3960
         Width           =   3225
      End
      Begin VB.TextBox txtFacultyCredit 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1800
         TabIndex        =   16
         Top             =   3600
         Width           =   3225
      End
      Begin VB.TextBox txtLabUnit 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1800
         MaxLength       =   50
         TabIndex        =   14
         Top             =   2520
         Width           =   705
      End
      Begin VB.TextBox txtSubjectID 
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
         Height          =   345
         Left            =   1800
         MaxLength       =   20
         TabIndex        =   12
         Top             =   360
         Width           =   3225
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
         Height          =   675
         Left            =   1800
         MaxLength       =   100
         MultiLine       =   -1  'True
         TabIndex        =   6
         Top             =   1080
         Width           =   3225
      End
      Begin VB.TextBox txtDepartmentTitle 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1800
         TabIndex        =   5
         Top             =   2880
         Width           =   3225
      End
      Begin VB.TextBox txtStudentCredit 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1800
         TabIndex        =   4
         Top             =   3240
         Width           =   3225
      End
      Begin VB.CommandButton cmdGetDepartmentTitle 
         BackColor       =   &H00D8E9EC&
         Height          =   300
         Left            =   5040
         Picture         =   "frmSubjectAE.frx":492A
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   2880
         Width           =   345
      End
      Begin VB.TextBox txtunits 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1800
         MaxLength       =   50
         TabIndex        =   2
         Top             =   1800
         Width           =   705
      End
      Begin VB.TextBox txtLecUnit 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1800
         MaxLength       =   50
         TabIndex        =   1
         Top             =   2160
         Width           =   705
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Subject"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   29
         Top             =   720
         Width           =   540
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Repeat Fee"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   25
         Top             =   5040
         Width           =   840
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Laboratory Fee"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   23
         Top             =   4680
         Width           =   1110
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Subject Fee"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   21
         Top             =   4320
         Width           =   855
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Category"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   19
         Top             =   3960
         Width           =   675
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Faculty Credit"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   17
         Top             =   3600
         Width           =   1005
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Lab(Unit)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   15
         Top             =   2520
         Width           =   660
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Subject Code"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   13
         Top             =   360
         Width           =   960
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Descriptive Title"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   11
         Top             =   1080
         Width           =   1140
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Department"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   10
         Top             =   2880
         Width           =   855
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Student Credit"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   9
         Top             =   3240
         Width           =   1050
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Units"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   8
         Top             =   1800
         Width           =   360
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Lec(Unit)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   7
         Top             =   2160
         Width           =   645
      End
   End
End
Attribute VB_Name = "frmSubjectAE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim vRS As New ADODB.Recordset

Dim RecordAdded As Boolean
Dim mFormState As String
Dim RecordEdited As Boolean

Dim CurrentSubject As tSubject

Public Function ShowForm(Optional sDepartmentTitle As String = "", Optional sYearLevelTitle As String = "") As Boolean
    
    Dim sNewSubjectID As String
    
    'set defaults
    ShowForm = False
    RecordAdded = False
    
    
    'check if other related recordset rentry exist
    If DepartmentRecordExist <> Success Then
        MsgBox "Unable to continue Adding Subject." & vbNewLine & "Department entries not exist", vbExclamation
        Unload Me
        Exit Function
    End If
        
    
    txtDepartmentTitle.Text = sDepartmentTitle
    If GetNewSubjectID(sNewSubjectID) = Success Then
        txtSubjectID.Text = sNewSubjectID
    End If
    
    mFormState = "ADD"

    Me.Show vbModal
    ShowForm = RecordAdded
End Function
Private Function SaveData() As Boolean
    
    Dim vSubject As tSubject
    Dim vDepartment As tDepartment

    SaveData = False

    If Not ValidateData Then Exit Function

    If GetDepartmentByTitle(txtDepartmentTitle.Text, vDepartment) <> Success Then
        MsgBox "Invalid Department Title", vbExclamation
        HLTxt txtDepartmentTitle
        Exit Function
    End If
    
        vSubject.SubjectID = Me.txtSubjectID.Text
        vSubject.SubjectTitle = txtSubject.Text
        vSubject.Description = Me.txtDescription.Text
        vSubject.DepartmentID = vDepartment.DepartmentID
        vSubject.Units = txtunits.Text
        vSubject.Category = txtCategory.Text
        vSubject.FacultyCredit = txtFacultyCredit.Text
        vSubject.LaboratoryFee = txtLaboratoryFee.Text
        vSubject.LaboratoryUnits = txtLabUnit.Text
        vSubject.LectureUnits = txtLecUnit.Text
        vSubject.RepeatFee = txtRepeatFee.Text
        vSubject.StudentCredit = txtStudentCredit.Text
        vSubject.SubjectFee = txtSubjectFee.Text


    Select Case AddSubject(vSubject)
        Case TranDBResult.Success
            SaveData = True
            
        Case TranDBResult.DuplicateID
            MsgBox "ID already existed.", vbExclamation
            HLTxt txtSubjectID
            SaveData = False
        
        Case TranDBResult.DuplicateTitle
            MsgBox "Title already existed.", vbExclamation
            HLTxt txtSubject
            SaveData = False
            
        Case TranDBResult.InvalidSubjectDepartmentID
            MsgBox "Invalid Department.", vbExclamation
            HLTxt txtDepartmentTitle
            SaveData = False

        Case TranDBResult.InvalidSubjectDescription
            MsgBox "Invalid Description.", vbExclamation
            HLTxt txtDescription
            SaveData = False
            
        Case Else
            MsgBox "Unknown Error.", vbExclamation
            SaveData = False
    End Select
End Function
Private Function ValidateData() As Boolean
    ValidateData = False

    If Not CheckTextBox(txtSubjectID, "Please Enter Subject ID") Then
        Exit Function
    End If
    
    If Not CheckTextBox(txtSubject, "Please Enter Subject Title") Then
        Exit Function
    End If
    
    If Not CheckTextBox(txtDescription, "Please Enter Description") Then
        Exit Function
    End If

    If Not CheckTextBox(txtDescription, "Please Enter Description") Then
        Exit Function
    End If
    
    If Not CheckTextBox(txtunits, "Please Enter Subject Unit") Then
        Exit Function
    End If
    
    If Not CheckTextBox(Me.txtCategory, "Please Enter Category") Then
        Exit Function
    End If
    
    If Not CheckTextBox(Me.txtFacultyCredit, "Please Enter Faculty Credit") Then
        Exit Function
    End If
    
    If Not CheckTextBox(Me.txtLaboratoryFee, "Please Enter Laboratory Fee") Then
        Exit Function
    End If

    If Not CheckTextBox(Me.txtLabUnit, "Please Enter Laboratory Unit") Then
        Exit Function
    End If
    
    If Not CheckTextBox(Me.txtLecUnit, "Please Enter Lecture Unit") Then
        Exit Function
    End If
    
    If Not CheckTextBox(Me.txtRepeatFee, "Please Enter Repeat Fee") Then
        Exit Function
    End If
    
    If Not CheckTextBox(Me.txtStudentCredit, "Please Enter Student Credit") Then
        Exit Function
    End If
    
    If Not CheckTextBox(Me.txtSubjectFee, "Please Enter Subject Fee") Then
        Exit Function
    End If
    
    
    
    If DepartmentExistByTitle(txtDepartmentTitle.Text) <> Success Then
        MsgBox "Invalid Department Title", vbExclamation
        HLTxt txtDepartmentTitle
        Exit Function
    End If
    
    ValidateData = True
End Function
Private Sub cmdCancel_Click()
    Unload Me
End Sub
Private Sub cmdGetDepartmentTitle_Click()
    Dim sDepartmentTitle As String

     frmPickDepartment.GetItem txtDepartmentTitle, sDepartmentTitle
    If sDepartmentTitle <> "" Then
        txtDepartmentTitle = sDepartmentTitle
    End If
End Sub
Private Sub cmdSave_Click()
    Select Case mFormState
        Case "ADD"
            If SaveData Then 'added
                MsgBox "Subject Entry successfully added.", vbInformation
                RecordAdded = True
                Unload Me
            End If
        Case Else
            If UpdateData Then 'updated
                MsgBox "Subject Entry successfully updated.", vbInformation
                RecordEdited = True
                Unload Me
            End If
    End Select
End Sub
Public Function ShowEdit(sSubjectID As String) As Boolean
    Dim vDepartment As tDepartment

    ShowEdit = False
    
    If GetSubjectByID(sSubjectID, CurrentSubject) = Success Then
        
        txtSubjectID.Text = CurrentSubject.SubjectID
        txtSubject.Text = CurrentSubject.SubjectTitle
        txtDescription.Text = CurrentSubject.Description
        txtunits.Text = CurrentSubject.Units
        txtLecUnit.Text = CurrentSubject.LectureUnits
        txtLabUnit.Text = CurrentSubject.LaboratoryUnits
        txtStudentCredit.Text = CurrentSubject.StudentCredit
        txtFacultyCredit.Text = CurrentSubject.FacultyCredit
        txtLaboratoryFee.Text = CurrentSubject.LaboratoryFee
        txtRepeatFee.Text = CurrentSubject.RepeatFee
        txtSubjectFee.Text = CurrentSubject.SubjectFee
        txtCategory.Text = CurrentSubject.Category
        
        If GetDepartmentByID(CurrentSubject.DepartmentID, vDepartment) = Success Then
            txtDepartmentTitle = vDepartment.DepartmentTitle
        End If
        
        mFormState = "EDIT"
        
    Else
        
        Unload Me
        Exit Function
    
    End If
    Me.Show vbModal
    
    ShowEdit = RecordEdited
End Function
Private Function UpdateData() As Boolean
    
    Dim newSubject As tSubject
    Dim vDepartment As tDepartment

    UpdateData = False

    If Not ValidateData Then Exit Function
    
    If GetDepartmentByTitle(txtDepartmentTitle.Text, vDepartment) <> Success Then
        MsgBox "Invalid Department Title", vbExclamation
        HLTxt txtDepartmentTitle
        Exit Function
    End If
    
        newSubject.SubjectID = txtSubjectID.Text
        newSubject.SubjectTitle = txtSubject.Text
        newSubject.Description = Me.txtDescription.Text
        newSubject.DepartmentID = vDepartment.DepartmentID
        newSubject.Units = txtunits.Text
        newSubject.Category = txtCategory.Text
        newSubject.FacultyCredit = txtFacultyCredit.Text
        newSubject.LaboratoryFee = txtLaboratoryFee.Text
        newSubject.LaboratoryUnits = txtLabUnit.Text
        newSubject.LectureUnits = txtLecUnit.Text
        newSubject.RepeatFee = txtRepeatFee.Text
        newSubject.StudentCredit = txtStudentCredit.Text
        newSubject.SubjectFee = txtSubjectFee.Text

    Select Case EditSubject(newSubject)
        Case TranDBResult.Success
            UpdateData = True
            
        Case TranDBResult.DuplicateTitle
            MsgBox "Title already existed.", vbExclamation
            HLTxt txtSubject
            UpdateData = False
            
        Case TranDBResult.InvalidSubjectDepartmentID
            MsgBox "Invalid Department.", vbExclamation
            HLTxt txtDepartmentTitle
            UpdateData = False
            
        Case TranDBResult.InvalidSubjectDescription
            MsgBox "Invalid Description.", vbExclamation
            HLTxt txtDescription
            UpdateData = False
    End Select
End Function

