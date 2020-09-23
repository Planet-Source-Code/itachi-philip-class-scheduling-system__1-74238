VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmSection 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Section Properties"
   ClientHeight    =   7800
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   8970
   Icon            =   "frmSection.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7800
   ScaleWidth      =   8970
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   7815
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   13785
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "General"
      TabPicture(0)   =   "frmSection.frx":492A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "ilRecordIco"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lsvInstructor"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Grid"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lsvRoomSchedule"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Enrollees"
      TabPicture(1)   =   "frmSection.frx":4946
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lsvEnrollees"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "StatusBar1"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      Begin MSComctlLib.StatusBar StatusBar1 
         Height          =   375
         Left            =   -74940
         TabIndex        =   5
         Top             =   7390
         Width           =   9015
         _ExtentX        =   15901
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
            NumPanels       =   3
            BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               Object.Width           =   3528
               MinWidth        =   3528
            EndProperty
            BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               Object.Width           =   3528
               MinWidth        =   3528
            EndProperty
            BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               Object.Width           =   8466
               MinWidth        =   8466
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ListView lsvEnrollees 
         Height          =   7050
         Left            =   -75000
         TabIndex        =   4
         Top             =   360
         Width           =   8950
         _ExtentX        =   15796
         _ExtentY        =   12435
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         Icons           =   "ilRecordIco"
         SmallIcons      =   "ilRecordIco"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Student Name"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "ID Number"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Course"
            Object.Width           =   7056
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Date Enrolled"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Enrolled By"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ListView lsvRoomSchedule 
         Height          =   3375
         Left            =   60
         TabIndex        =   2
         Top             =   2160
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   5953
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         Icons           =   "ilRecordIco"
         SmallIcons      =   "ilRecordIco"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Room Assignment"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Schedule"
            Object.Width           =   6174
         EndProperty
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid Grid 
         Height          =   1770
         Left            =   60
         TabIndex        =   1
         Top             =   360
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   3122
         _Version        =   393216
         Rows            =   6
         FixedRows       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin MSComctlLib.ListView lsvInstructor 
         Height          =   2175
         Left            =   60
         TabIndex        =   3
         Top             =   5520
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   3836
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         Icons           =   "ilRecordIco"
         SmallIcons      =   "ilRecordIco"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Instructor"
            Object.Width           =   10583
         EndProperty
      End
      Begin MSComctlLib.ImageList ilRecordIco 
         Left            =   0
         Top             =   480
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   1
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmSection.frx":4962
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.Menu mnuroomschedule 
      Caption         =   "Room Schedue"
      Visible         =   0   'False
      Begin VB.Menu mnuaddroom 
         Caption         =   "Add room schedule..."
      End
      Begin VB.Menu mnuedit 
         Caption         =   "Edit room schedule..."
      End
      Begin VB.Menu mnudelete 
         Caption         =   "Delete"
      End
      Begin VB.Menu mnusep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnurefresh 
         Caption         =   "Refresh"
      End
   End
   Begin VB.Menu mnuInstructor 
      Caption         =   "Instructor"
      Visible         =   0   'False
      Begin VB.Menu mnudeleteIns 
         Caption         =   "Delete instructor"
      End
      Begin VB.Menu s 
         Caption         =   "-"
      End
      Begin VB.Menu mnurefreshIns 
         Caption         =   "Refresh"
      End
   End
   Begin VB.Menu mnuEnrollees 
      Caption         =   "Enrollees"
      Visible         =   0   'False
      Begin VB.Menu mnuViewInfo 
         Caption         =   "View student info..."
         Visible         =   0   'False
      End
      Begin VB.Menu mnuExportTo 
         Caption         =   "Export to"
      End
      Begin VB.Menu mnusep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnurefreshenrollees 
         Caption         =   "Refresh"
      End
   End
End
Attribute VB_Name = "frmSection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub ShowForm(sSectionID As String)
Dim vRS As New ADODB.Recordset
Dim sSQL As String

On Error GoTo err

sSQL = "SELECT tblSection.SectionID AS lvKey, tblSection.SectionTitle, tblSubject.SubjectTitle,tblSubject.Description,tblSection.SectionID, tblDepartment.DepartmentTitle AS Gender, tblSection.Slots, tblRoom.Building & '- ' & tblRoom.Room AS Room, tblSubjectOffering.Days & ' ' & tblSubjectOffering.TimeIn & '- ' & tblSubjectOffering.TimeOut AS Schedule " & _
        "FROM tblSubject INNER JOIN ((tblDepartment INNER JOIN tblSection ON tblDepartment.DepartmentID = tblSection.DepartmentID) INNER JOIN (tblRoom INNER JOIN tblSubjectOffering ON tblRoom.RoomID = tblSubjectOffering.RoomID) ON tblSection.SectionID = tblSubjectOffering.SectionID) ON tblSubject.SubjectID = tblSubjectOffering.SubjectID " & _
        "WHERE tblSection.SectionID='" & sSectionID & "'"

If ConnectRS(con, vRS, sSQL) = True Then
    With Grid
        .Clear
        .ClearStructure
        .Rows = 6
        .FixedCols = 1
        
        .ColWidth(0) = 1800
        .ColWidth(1) = 6900

        .TextMatrix(0, 0) = "Subject"
        .TextMatrix(1, 0) = "Section"
        .TextMatrix(2, 0) = "Slots"
        .TextMatrix(3, 0) = "Gender"
        .TextMatrix(4, 0) = "Remarks"
        .TextMatrix(5, 0) = "Locked"
        
        .ColAlignment(0) = vbLeftJustify
        .ColAlignment(1) = vbLeftJustify
        
        .TextMatrix(0, 1) = vRS.Fields("SubjectTitle") & " " & vRS.Fields("Description")
        .Row = 0
        .Col = 1
        .CellBackColor = &H80C0FF
        .TextMatrix(1, 1) = vRS.Fields("SectionTitle")
        .TextMatrix(2, 1) = vRS.Fields("Slots")
        .TextMatrix(3, 1) = vRS.Fields("Gender")
        '.TextMatrix(4, 1) = "Remarks"
        '.TextMatrix(5, 1) = "Locked"
        
    End With
    
    ShowSectionEnrollees sSectionID
    ShowSectionRoom sSectionID
End If
    
    
    
    Me.Show 1
Exit Sub
err:
    MsgBox err.Description, vbCritical
End Sub

Private Sub ShowSectionEnrollees(sSectionID As String)
Dim vRS As New ADODB.Recordset
Dim sSQL As String
Dim lv As ListItem

    sSQL = "SELECT tblStudent.StudentID, tblSection.SectionTitle, [tblStudent]![LastName] & ', ' & [tblStudent]![FirstName] & ' ' & [tblStudent]![MiddleName] AS StudentFullName, tblEnrolment.CreatedBy, tblCourse.Course & ' major in ' & tblCourse.Major AS Course, tblEnrolment.DateEnroled " & _
            "FROM (tblSubject INNER JOIN (tblSection INNER JOIN ((tblTeacher INNER JOIN tblSubjectOffering ON tblTeacher.TeacherID = tblSubjectOffering.TeacherID) INNER JOIN (tblStudent INNER JOIN (tblEnrolment INNER JOIN tblGrade ON tblEnrolment.EnrollmentID = tblGrade.EnrolmentID) ON tblStudent.StudentID = tblEnrolment.StudentID) ON (tblSubjectOffering.SubjectOfferingID = tblGrade.SubjectOfferingID) AND (tblSubjectOffering.SubjectOfferingID = tblGrade.SubjectOfferingID)) ON tblSection.SectionID = tblGrade.SectionID) ON tblSubject.SubjectID = tblSubjectOffering.SubjectID) INNER JOIN (tblCourse INNER JOIN tblStudentStatus ON tblCourse.CourseID = tblStudentStatus.CourseID) ON tblStudent.StudentID = tblStudentStatus.StudentID " & _
            "WHERE tblSection.SectionID ='" & sSectionID & "'"
            
    If ConnectRS(con, vRS, sSQL) = True Then
        lsvEnrollees.ListItems.Clear
        Do Until vRS.EOF
            Set lv = lsvEnrollees.ListItems.Add(, , vRS.Fields("StudentFullName"), 1, 1)
                    lv.SubItems(1) = vRS.Fields("StudentID")
                    lv.SubItems(2) = vRS.Fields("Course")
                    lv.SubItems(3) = vRS.Fields("DateEnroled")
                    lv.SubItems(4) = vRS.Fields("CreatedBy")
            vRS.MoveNext
        Loop
    End If
    Set vRS = Nothing
End Sub


Private Sub ShowSectionRoom(sSectionID As String)
Dim vRS As New ADODB.Recordset
Dim sSQL As String
Dim lv As ListItem

    sSQL = "SELECT tblSubjectOffering.SectionID, tblRoom.Building &' '&tblRoom.Room as Room, tblSubjectOffering.Days&' '& tblSubjectOffering.TimeIn&' - '&tblSubjectOffering.TimeOut as Schedule, tblSubjectOffering.SchoolYear, tblSubjectOffering.Semester " & _
            "FROM tblRoom INNER JOIN tblSubjectOffering ON tblRoom.RoomID = tblSubjectOffering.RoomID " & _
            "WHERE tblSubjectOffering.SectionID ='" & sSectionID & "'"
            
    If ConnectRS(con, vRS, sSQL) = True Then
        lsvRoomSchedule.ListItems.Clear
        Do Until vRS.EOF
            Set lv = lsvRoomSchedule.ListItems.Add(, , vRS.Fields("Room"), 1, 1)
                    lv.SubItems(1) = vRS.Fields("Schedule")
            vRS.MoveNext
        Loop
    End If
    Set vRS = Nothing
End Sub
