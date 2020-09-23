VERSION 5.00
Begin VB.Form frmProspectusAE 
   BackColor       =   &H00D8E9EC&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Prospectus"
   ClientHeight    =   3015
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   7905
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   7905
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox CourseID 
      Height          =   315
      Left            =   1440
      TabIndex        =   14
      Top             =   5760
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.ComboBox cboSemesterID 
      Height          =   315
      Left            =   480
      TabIndex        =   12
      Top             =   5760
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4200
      TabIndex        =   11
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   2640
      TabIndex        =   10
      Top             =   2520
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00D8E9EC&
      Height          =   2415
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   7695
      Begin VB.ComboBox cboCourse 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   1080
         Width           =   7455
      End
      Begin VB.TextBox txtSubject 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   120
         MaxLength       =   50
         TabIndex        =   5
         Top             =   480
         Width           =   7065
      End
      Begin VB.ComboBox cboSemester 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "frmProspectusAE.frx":0000
         Left            =   4080
         List            =   "frmProspectusAE.frx":000D
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1680
         Width           =   3495
      End
      Begin VB.CommandButton cmdGetSubject 
         Appearance      =   0  'Flat
         BackColor       =   &H00D8E9EC&
         Height          =   345
         Left            =   7200
         Picture         =   "frmProspectusAE.frx":0035
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   480
         Width           =   345
      End
      Begin VB.CommandButton cmdGetYearLevel 
         BackColor       =   &H00D8E9EC&
         Height          =   330
         Left            =   3360
         Picture         =   "frmProspectusAE.frx":05BF
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   1680
         Width           =   345
      End
      Begin VB.TextBox txtYearLevel 
         Appearance      =   0  'Flat
         Height          =   345
         Left            =   120
         MaxLength       =   50
         TabIndex        =   1
         Top             =   1680
         Width           =   3225
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Course"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   210
         Left            =   120
         TabIndex        =   9
         Top             =   840
         Width           =   630
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Subject"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   210
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   705
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Semester"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   4080
         TabIndex        =   7
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Year Level"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   6
         Top             =   1440
         Width           =   915
      End
   End
End
Attribute VB_Name = "frmProspectusAE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim listRoomID() As String
Dim curSubjectID As String
Dim SectionOfferingID As String
Dim curRoomID As String

Dim vSubject As tSubject

Private Sub cboCourse_Change()
    CourseID.ListIndex = cboCourse.ListIndex
End Sub

Private Sub cboCourse_Click()
    cboCourse_Change
End Sub

Private Sub cboSemester_Change()
    cboSemesterID.ListIndex = cboSemester.ListIndex
End Sub

Private Sub cboSemester_Click()
    cboSemester_Change
End Sub


Private Sub cmdGetSubject_Click()
Dim sSubjectTitle As String

    sSubjectTitle = frmPickSubject.GetSubjectTitle
        
    If sSubjectTitle <> "" Then
        If GetSubjectByTitle(sSubjectTitle, vSubject) = Success Then
            curSubjectID = vSubject.SubjectID
        Else
            MsgBox "Unable to continue this operation." & vbNewLine & "The selected Subject ID not found in record.", vbCritical
        End If
    End If
End Sub

Private Sub cmdGetYearLevel_Click()
    Dim sYearLevelTitle As String
    sYearLevelTitle = PickYearLevel.GetYearLevelTitle
    txtYearLevel.Text = sYearLevelTitle
End Sub



Private Sub cmdSave_Click()
    AddToProspectus
End Sub

Private Sub AddToProspectus()
Dim sProspectusID As String
Dim vRS As New ADODB.Recordset

If ConnectRS(con, vRS, "Select * From tblProspectus Where SubjectID='" & curSubjectID & "' and CourseID='" & Me.CourseID.Text & "' and SemesterID='" & cboSemesterID.Text & "' and YearLevel='" & txtYearLevel.Text & "'") = True Then
    If GetNewProspectusID(sProspectusID) = Failed Then
        Exit Sub
    End If
   
    If vRS.RecordCount > 0 Then
        MsgBox "Subject is already added", vbInformation
    Else
        vRS.AddNew
        vRS.Fields("ProspectusID") = sProspectusID
        vRS.Fields("SubjectID") = curSubjectID
        vRS.Fields("CourseID") = CourseID.Text
        vRS.Fields("SemesterID") = cboSemesterID.Text
        vRS.Fields("YearLevel") = txtYearLevel.Text
        vRS.Update
        
        Unload Me
        
        frmCurriculum.Refresh_Prospectus
    End If
End If
End Sub
Private Sub CourseList()
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    
    sSQL = "SELECT tblCourse.CourseID as lvKey,tblCourse.CourseID,tblCourse.Course,tblCourse.Curriculum" & _
            " FROM tblCourse" & _
            " ORDER BY tblCourse.Course"
     
    If ConnectRS(con, vRS, sSQL) = True Then
        CourseID.Clear
        cboCourse.Clear
        Do Until vRS.EOF
            cboCourse.AddItem (vRS.Fields("Course"))
            CourseID.AddItem (vRS.Fields("CourseID"))
            vRS.MoveNext
        Loop
        
    End If

ReleaseAndExit:
    Set vRS = Nothing
End Sub
Private Sub SemesterList()
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    
    sSQL = "SELECT tblSemester.SemesterID as lvKey,tblSemester.SemesterID ,tblSemester.Semester" & _
            " FROM tblSemester" & _
            " ORDER BY tblSemester.Semester"
            
    If ConnectRS(con, vRS, sSQL) = True Then
        cboSemester.Clear
        cboSemesterID.Clear
        Do Until vRS.EOF
            cboSemester.AddItem (vRS.Fields("Semester"))
            cboSemesterID.AddItem (vRS.Fields("SemesterID"))
            vRS.MoveNext
        Loop
    End If

ReleaseAndExit:
    Set vRS = Nothing
End Sub

Public Sub ShowForm(Optional sSubject As String = "")
On Error Resume Next
       
    txtSubject.Text = sSubject

    CourseList
    SemesterList
    
     cboCourse.Text = frmCurriculum.cboCourse.Text
     cboSemester.Text = CurrentSemester.Semester
        
    Me.Show vbModal
End Sub

Private Sub txtSubject_Change()
    curSubjectID = frmCurriculum.lsvSubject.SelectedItem.SubItems(1)
End Sub
