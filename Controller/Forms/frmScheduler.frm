VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmScheduler 
   BackColor       =   &H00D8E9EC&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Scheduler"
   ClientHeight    =   8010
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   12210
   Icon            =   "frmScheduler.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8010
   ScaleWidth      =   12210
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      BackColor       =   &H00D8E9EC&
      Caption         =   "List of Conflict"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   2535
      Left            =   3840
      TabIndex        =   31
      Top             =   5400
      Width           =   8295
      Begin MSComctlLib.ListView lsvConflicts 
         Height          =   2175
         Left            =   120
         TabIndex        =   32
         Top             =   240
         Width           =   8055
         _ExtentX        =   14208
         _ExtentY        =   3836
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Subject"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Section"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Enrolled"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Schedule"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Lecturer"
            Object.Width           =   4410
         EndProperty
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00D8E9EC&
      Caption         =   "Room Usage"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   2895
      Left            =   3840
      TabIndex        =   29
      Top             =   2640
      Width           =   8295
      Begin MSComctlLib.ListView lsvRoomSchedule 
         Height          =   2535
         Left            =   120
         TabIndex        =   30
         Top             =   240
         Width           =   8055
         _ExtentX        =   14208
         _ExtentY        =   4471
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Subject"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Section"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Enrolled"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Schedule"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Lecturer"
            Object.Width           =   4410
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00D8E9EC&
      Caption         =   "Input Area"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   7935
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   3855
      Begin VB.ComboBox cboDay 
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
         ItemData        =   "frmScheduler.frx":492A
         Left            =   120
         List            =   "frmScheduler.frx":4952
         Style           =   2  'Dropdown List
         TabIndex        =   27
         Top             =   5640
         Width           =   3255
      End
      Begin VB.CommandButton cmdGetSubject 
         Appearance      =   0  'Flat
         BackColor       =   &H00D8E9EC&
         Height          =   345
         Left            =   3360
         Picture         =   "frmScheduler.frx":4982
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   1200
         Width           =   345
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
         ItemData        =   "frmScheduler.frx":4F0C
         Left            =   120
         List            =   "frmScheduler.frx":4F19
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   3600
         Width           =   3255
      End
      Begin VB.TextBox txtSchoolYearTitle 
         Appearance      =   0  'Flat
         Height          =   345
         Left            =   120
         MaxLength       =   50
         TabIndex        =   21
         Top             =   3000
         Width           =   3225
      End
      Begin VB.CommandButton cmdGetSchoolYear 
         BackColor       =   &H00D8E9EC&
         Height          =   345
         Left            =   3360
         Picture         =   "frmScheduler.frx":4F41
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   3000
         Width           =   345
      End
      Begin VB.TextBox txtDepartmentTitle 
         Appearance      =   0  'Flat
         Height          =   345
         Left            =   120
         MaxLength       =   50
         TabIndex        =   18
         Top             =   2400
         Width           =   3225
      End
      Begin VB.CommandButton cmdGetItem 
         Appearance      =   0  'Flat
         BackColor       =   &H00D8E9EC&
         Height          =   345
         Left            =   3360
         Picture         =   "frmScheduler.frx":54CB
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   2400
         Width           =   345
      End
      Begin VB.TextBox txtSlots 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         MaxLength       =   20
         TabIndex        =   15
         Text            =   "40"
         Top             =   4320
         Width           =   3225
      End
      Begin VB.CommandButton cmdGetTeacher 
         Appearance      =   0  'Flat
         BackColor       =   &H00D8E9EC&
         Height          =   345
         Left            =   3360
         Picture         =   "frmScheduler.frx":5A55
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   1800
         Width           =   345
      End
      Begin VB.TextBox txtTeacherFullName 
         Appearance      =   0  'Flat
         Height          =   345
         Left            =   120
         MaxLength       =   20
         TabIndex        =   13
         Top             =   1800
         Width           =   3225
      End
      Begin VB.TextBox txtSubject 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   120
         MaxLength       =   50
         TabIndex        =   12
         Top             =   1200
         Width           =   3225
      End
      Begin VB.TextBox txtSection 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   120
         MaxLength       =   50
         TabIndex        =   11
         Top             =   600
         Width           =   3225
      End
      Begin VB.ComboBox dtpTo 
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
         ItemData        =   "frmScheduler.frx":5FDF
         Left            =   2040
         List            =   "frmScheduler.frx":603D
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   6240
         Width           =   1455
      End
      Begin VB.ComboBox dtpFrom 
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
         ItemData        =   "frmScheduler.frx":6155
         Left            =   120
         List            =   "frmScheduler.frx":61B3
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   6240
         Width           =   1455
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
         Height          =   495
         Left            =   360
         TabIndex        =   7
         Top             =   6720
         Width           =   2655
      End
      Begin VB.ComboBox cboRoom 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "frmScheduler.frx":62CB
         Left            =   120
         List            =   "frmScheduler.frx":62CD
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   4920
         Width           =   3255
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Day "
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
         TabIndex        =   28
         Top             =   5400
         Width           =   405
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
         Left            =   120
         TabIndex        =   24
         Top             =   3360
         Width           =   855
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "School Year"
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
         TabIndex        =   22
         Top             =   2760
         Width           =   1065
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Department"
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
         TabIndex        =   19
         Top             =   2160
         Width           =   1110
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Slots"
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
         TabIndex        =   16
         Top             =   4080
         Width           =   465
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "To"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1680
         TabIndex        =   8
         Top             =   6240
         Width           =   255
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Faculty"
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
         TabIndex        =   6
         Top             =   1560
         Width           =   645
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
         TabIndex        =   5
         Top             =   960
         Width           =   705
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Room"
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
         TabIndex        =   4
         Top             =   4680
         Width           =   540
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Section"
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
         TabIndex        =   3
         Top             =   360
         Width           =   690
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00D8E9EC&
      Caption         =   "Faculty Schedules"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   2775
      Left            =   3840
      TabIndex        =   0
      Top             =   0
      Width           =   8295
      Begin MSComctlLib.ListView lsvSchedule 
         Height          =   2415
         Left            =   120
         TabIndex        =   25
         Top             =   240
         Width           =   8055
         _ExtentX        =   14208
         _ExtentY        =   4260
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Subject"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Section"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Enrolled"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Schedule"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Lecturer"
            Object.Width           =   4410
         EndProperty
      End
   End
End
Attribute VB_Name = "frmScheduler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim RecordAdded As Boolean
Dim cIRowCount, cIRowCount1 As Integer

Dim curSectionID As String
Dim curTeacherID As String

Dim curSubjectID As String
Dim SectionOfferingID As String
Dim curRoomID As String

Dim curDepartment As String
Dim curTeacher As tTeacher

Dim sSubjectTitle As String
Dim vSubject As tSubject
    
Dim listRoomID() As String

Private Sub cboRoom_Change()
On Error Resume Next
    curRoomID = listRoomID(cboRoom.ListIndex)
    Call ShowRoomSchedule(curRoomID, lsvRoomSchedule)
End Sub

Private Sub cboRoom_Click()
    cboRoom_Change
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

Private Sub cmdGetSchoolYear_Click()
     Dim sSchoolYearTitle As String
    
    sSchoolYearTitle = frmPickSchoolYear.GetItem(txtSchoolYearTitle, , , True)
    
    If sSchoolYearTitle <> "" Then
        txtSchoolYearTitle.Text = sSchoolYearTitle
    End If
End Sub

Private Sub cmdGetSubject_Click()
    sSubjectTitle = frmPickSubject.GetSubjectTitle(txtSubject)
        
    If sSubjectTitle <> "" Then
        If GetSubjectByTitle(sSubjectTitle, vSubject) = Success Then
            curSubjectID = vSubject.SubjectID
        Else
            MsgBox "Unable to continue this operation." & vbNewLine & "The selected Subject ID not found in record.", vbCritical
        End If
    End If
End Sub

Private Sub cmdGetTeacher_Click()
Dim sTeacherID As String
    Dim sTeacherFullName As String
    
    sTeacherID = frmPickTeacher.GetTeacherID(sTeacherFullName)
    
    If sTeacherID <> "" Then
        curTeacherID = sTeacherID
        txtTeacherFullName.Text = sTeacherFullName
            
        ShowTeacherSchedule curTeacherID, lsvSchedule
        
    End If
End Sub
Private Function LoadRoom() As Boolean
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    LoadRoom = False
    
    sSQL = "SELECT tblRoom.RoomID, tblRoom.Room" & _
            " From tblRoom" & _
            " ORDER BY tblRoom.Room"
            
    If ConnectRS(con, vRS, sSQL) = False Then
        CatchError "AddSectionOffering", "RefreshRommList", "Unable to connect Recordset with SQL Expression : '" & sSQL & "'"
        GoTo ReleaseAndExit
    End If
    
    If AnyRecordExisted(vRS) = False Then
        GoTo ReleaseAndExit
    End If
    
    ReDim listRoomID(getRecordCount(vRS) - 1)
    cboRoom.Clear
    vRS.MoveFirst
    While vRS.EOF = False
        cboRoom.AddItem (vRS("Room"))
        listRoomID(cboRoom.ListCount - 1) = (vRS("RoomID"))
        vRS.MoveNext
    Wend
    LoadRoom = True
    
ReleaseAndExit:
    Set vRS = Nothing
End Function

Public Sub ShowForm(ByRef sSubject As String, Optional sSchoolYear As String = "", Optional sSemester As String = "", Optional sDepartment As String = "")
    
    If sSchoolYear = "" Then
        txtSchoolYearTitle.Text = CurrentSchoolYear.SchoolYearTitle
    Else
        txtSchoolYearTitle.Text = sSchoolYear
    End If
    
    If sSemester = "" Then
        cboSemester.Text = CurrentSemester.Semester
    Else
        cboSemester.Text = sSemester
    End If
    
    txtSubject.Text = sSubject
    GenerateSectionOfferingID
    
    Me.Show vbModal
End Sub

Private Sub Form_SaveData()

    Dim newSectionOffering As tSectionOffering
    Dim vSection As tSection
    Dim vTeacher As tTeacher
    Dim vDepartment As tDepartment
    
    Dim sNewID As String
    
    Dim lvItem As ListItem
    
    Dim i As Integer
    Dim newSubjectOffering As tSubjectOffering
    Dim ErrMSG As String
    
    
    
    If ValidateData = False Then Exit Sub
        
    If Len(curTeacherID) < 1 Then
        MsgBox "Invalid Teahcer entry!" & vbNewLine & "Please Enter valid Teacher Full Name.", vbExclamation
        cmdGetTeacher.SetFocus
        Exit Sub
    End If
        

    newSectionOffering.SectionID = SectionOfferingID
    newSectionOffering.SectionTitle = txtSection.Text
    newSectionOffering.SchoolYear = txtSchoolYearTitle.Text
    newSectionOffering.DepartmentID = curDepartment
    newSectionOffering.Semester = cboSemester.Text
    newSectionOffering.Slots = Val(txtSlots.Text)
    newSectionOffering.CreationDate = Now
    newSectionOffering.CreatedBy = CurrentUser.Username
        
        
    Select Case AddSectionOffering(curSectionID, newSectionOffering)
        Case TranDBResult.Success
            ErrMSG = ""
                
                curRoomID = listRoomID(cboRoom.ListIndex)
                
                newSubjectOffering.SubjectOfferingID = newSectionOffering.SectionID & "-" & Left(CurrentSchoolYear.SchoolYearTitle, 4) & "-" & Left(CurrentSemester.Semester, 3)
                newSubjectOffering.SectionOfferingID = SectionOfferingID
                newSubjectOffering.SubjectID = curSubjectID
                newSubjectOffering.RoomID = curRoomID
                newSubjectOffering.TeacherID = curTeacherID
                newSubjectOffering.Days = cboDay.Text
                newSubjectOffering.TimeIn = dtpFrom.Text
                newSubjectOffering.TimeOut = dtpTo.Text
                newSubjectOffering.Semester = CurrentSemester.Semester
                newSubjectOffering.SchoolYear = CurrentSchoolYear.SchoolYearTitle
                newSubjectOffering.CreationDate = Now
                newSubjectOffering.CreatedBy = CurrentUser.Fullname
                                
                If AddSubjectOffering(newSubjectOffering) <> TranDBResult.Success Then
                    ErrMSG = vbNewLine & ErrMSG & "Error Adding Subject [ID: " & newSubjectOffering.SubjectID & "]  To Section [ID : " & newSectionOffering.SectionID & "]"
                End If
                
            If ErrMSG <> "" Then
                MsgBox "FATAL ERROR:" & ErrMSG, vbCritical
            End If
            
            MsgBox "New Section Offering entry successfully added.", vbInformation
            RecordAdded = True
            Unload Me
        Case TranDBResult.DuplicateID
            MsgBox "This Section Offering Entry is already exist in record." & vbNewLine & "Please change Section or School Year.", vbExclamation
            HLTxt txtSection
        Case Else
            CatchError "frmAddsection", "Form_savedata", "AddSection Unknown result"
    End Select
End Sub

Private Function ValidateData() As Boolean

    Dim sSubjects() As String
    curRoomID = listRoomID(cboRoom.ListIndex)
    ValidateData = False
        
    If SchoolYearExistByTitle(txtSchoolYearTitle.Text) <> Success Then
        MsgBox "Please enter valid School Year Title", vbExclamation
        HLTxt txtSchoolYearTitle
        Exit Function
    End If
    
    If SectionOfferingExistByID(SectionOfferingID) = Success Then
        MsgBox "This Section Offering Entry is already exist in record." & vbNewLine & "Please change Section or School Year.", vbExclamation
        HLTxt txtSection
        Exit Function
    End If
    
    If Len(curTeacherID) < 1 Then
        MsgBox "Please enter valid Teacher Full Name", vbExclamation
        Exit Function
    End If
    
    If Len(cboDay.Text) < 1 Then
        MsgBox "Please Select Day", vbExclamation
        Exit Function
    End If
    
    If Len(txtSection.Text) < 1 Then
         MsgBox "Please enter valid Section title", vbExclamation
        Exit Function
    End If
    
    If Len(txtDepartmentTitle.Text) < 1 Then
         MsgBox "Please enter valid Department title", vbExclamation
        Exit Function
    End If
    
    If Len(txtSchoolYearTitle.Text) < 1 Then
         MsgBox "Please enter valid Schoolyear ", vbExclamation
        Exit Function
    End If
    
    
    If IsNumeric(txtSlots.Text) Then
        If Val(txtSlots.Text) < 1 Or Val(txtSlots.Text) > 100 Then
            MsgBox "Invalid Entry!" & vbNewLine & "Max. Student # must be range 1-100", vbExclamation
            HLTxt txtSlots
            Exit Function
        End If
    Else
        MsgBox "Invalid Entry!" & vbNewLine & "Max. Student # must be range 1-100", vbExclamation
        HLTxt txtSlots
        Exit Function
    End If
    
    If Len(cboRoom.Text) < 1 Then
         MsgBox "Please select room", vbExclamation
        Exit Function
    End If
    
    If dtpFrom.ListIndex >= dtpTo.ListIndex Then
        MsgBox "Invalid Schedule.", vbInformation, "ERROR"
        Exit Function
    End If
    
    If FacultyInUse(TimeValue(dtpFrom.Text), TimeValue(dtpTo.Text), cboDay.Text, curTeacherID) = True Then
        modSchedule.ShowConflictFaculty dtpFrom.Text, dtpTo.Text, cboDay.Text, curTeacherID, lsvConflicts
        MsgBox "Faculty already in use.", vbInformation, "Conflict"
        Exit Function
    End If
    
    If RoomInUse(TimeValue(dtpFrom.Text), TimeValue(dtpTo.Text), cboDay.Text, curRoomID) = True Then
        modSchedule.ShowConflictRoom dtpFrom.Text, dtpTo.Text, cboDay.Text, curTeacherID, curRoomID, lsvConflicts
        MsgBox "Room already in use.", vbInformation, "Conflict"
        Exit Function
    End If
    
    ValidateData = True
End Function

Private Sub cmdSave_Click()
    Form_SaveData
End Sub

Private Sub Form_Activate()
    If LoadRoom = False Then
        MsgBox "There are no available Room to create Section Offering." & vbNewLine & _
            "Please add Room entry first.", vbExclamation
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    LoadRoom
    
    dtpFrom.ListIndex = 0
    dtpTo.ListIndex = 2
End Sub

Private Function RefreshSubjects()
        
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    lsvSchedule.ListItems.Clear
    
    On Error GoTo ReleaseAndExit
    
                sSQL = "SELECT tblSubjectOffering.SubjectOfferingID, tblSection.SchoolYear, tblSubject.SubjectTitle, tblSubjectOffering.SchedTimeStart, tblSubjectOffering.SchedTimeEnd, tblSubjectOffering.Days, tblDepartment.DepartmentTitle, tblYearLevel.YearLevelTitle, tblSection.SectionTitle, tblRoom.Room" & _
                        " FROM tblDepartment INNER JOIN (tblSubject INNER JOIN (tblTeacher INNER JOIN (tblRoom INNER JOIN (tblSubjectOffering INNER JOIN (tblYearLevel INNER JOIN tblSection ON tblYearLevel.YearLevelID = tblSection.YearLevelID) ON tblSubjectOffering.SectionID = tblSection.SectionID) ON tblRoom.RoomID = tblSubjectOffering.RoomID) ON tblTeacher.TeacherID = tblSubjectOffering.TeacherID) ON tblSubject.SubjectID = tblSubjectOffering.SubjectID) ON tblDepartment.DepartmentID = tblSection.DepartmentID" & _
                        " WHERE (((tblTeacher.TeacherID)='" & curTeacher.TeacherID & "') AND tblSection.SchoolYear='" & txtSchoolYearTitle.Text & "')" & _
                        " ORDER BY tblSection.SchoolYear DESC , tblSubject.SubjectTitle DESC"
     
    If ConnectRS(con, vRS, sSQL) = False Then
        MsgBox "Unable to connect Teacher Recordset.", vbCritical
        GoTo ReleaseAndExit
    End If
    
    If AnyRecordExisted(vRS) = True Then
        FillRecordToList vRS, lsvSchedule, KeySubjectOffering, , 32767, , True
    End If

ReleaseAndExit:
    Set vRS = Nothing
End Function

Private Sub txtSchoolYearTitle_Change()
    GenerateSectionOfferingID
End Sub

Private Sub txtSection_Change()
    GenerateSectionOfferingID
End Sub

Private Sub txtSlots_KeyPress(KeyAscii As Integer)
If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8 Or KeyAscii = 45) Then KeyAscii = 0
End Sub

Private Sub txtSlots_LostFocus()
If Len(txtSlots.Text) > 0 Then
        If IsNumeric(txtSlots.Text) Then
            If Val(txtSlots.Text) < 1 Or Val(txtSlots.Text) > 100 Then
                MsgBox "Invalid Entry!" & vbNewLine & "Max. Slots # must be range 1-100", vbExclamation
                HLTxt txtSlots
            End If
        Else
            MsgBox "Invalid Entry!" & vbNewLine & "Max. Slots # must be range 1-100", vbExclamation
            HLTxt txtSlots
        End If
    End If
End Sub

Private Sub GenerateSectionOfferingID()
    Dim vSection As tSection
    
    SectionOfferingID = ""
    
    If Len(txtSection.Text) < 1 Or Len(txtSchoolYearTitle.Text) < 1 Then
        curSectionID = ""
        Exit Sub
    End If

    
    SectionOfferingID = (txtSchoolYearTitle.Text) & "-" & txtSection.Text & "-" & txtSubject.Text
    
    Me.Caption = SectionOfferingID
    
    If SectionOfferingExistByID(SectionOfferingID) = Success Then
       MsgBox "This Section Offering Entry is already exist in record." & vbNewLine & "Please change Section or School Year.", vbExclamation
        HLTxt txtSection
    Else
        
    End If
End Sub

Private Sub txtSubject_Change()
     curSubjectID = GetLVKey(frmEnrollment.lsvSubject.SelectedItem)
End Sub




