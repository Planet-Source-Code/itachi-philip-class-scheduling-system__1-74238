VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmSubject 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Subject"
   ClientHeight    =   9465
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   13650
   Icon            =   "frmSubject.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9465
   ScaleWidth      =   13650
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   5000
      Left            =   0
      TabIndex        =   3
      Top             =   4080
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   8811
      _Version        =   393216
      Style           =   1
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Prerequisite"
      TabPicture(0)   =   "frmSubject.frx":492A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lvPreq"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Transcript"
      TabPicture(1)   =   "frmSubject.frx":4946
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lsvTranscript"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Course"
      TabPicture(2)   =   "frmSubject.frx":4962
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "imgListEnrolment"
      Tab(2).Control(1)=   "lsvCourse"
      Tab(2).ControlCount=   2
      Begin MSComctlLib.ListView lsvTranscript 
         Height          =   4575
         Left            =   -74940
         TabIndex        =   8
         Top             =   360
         Width           =   8655
         _ExtentX        =   15266
         _ExtentY        =   8070
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         Icons           =   "ImageList1"
         SmallIcons      =   "ImageList1"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Student Name"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "ID Number"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ListView lsvCourse 
         Height          =   4575
         Left            =   -74940
         TabIndex        =   9
         Top             =   360
         Width           =   8655
         _ExtentX        =   15266
         _ExtentY        =   8070
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         Icons           =   "ImageList1"
         SmallIcons      =   "ImageList1"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Course"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Curriculum"
            Object.Width           =   3528
         EndProperty
      End
      Begin MSComctlLib.ImageList imgListEnrolment 
         Left            =   -75000
         Top             =   360
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
               Picture         =   "frmSubject.frx":497E
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ListView lvPreq 
         Height          =   4575
         Left            =   60
         TabIndex        =   10
         Top             =   360
         Width           =   8655
         _ExtentX        =   15266
         _ExtentY        =   8070
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         OLEDragMode     =   1
         OLEDropMode     =   1
         FullRowSelect   =   -1  'True
         _Version        =   393217
         Icons           =   "imgListEnrolment"
         SmallIcons      =   "imgListEnrolment"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         OLEDragMode     =   1
         OLEDropMode     =   1
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Subject Title"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "ID"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Descriptive Title"
            Object.Width           =   7056
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Units"
            Object.Width           =   1411
         EndProperty
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   9090
      Width           =   13650
      _ExtentX        =   24077
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   17639
            MinWidth        =   17639
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Picture1 
      Align           =   4  'Align Right
      Height          =   9090
      Left            =   8715
      ScaleHeight     =   9030
      ScaleWidth      =   4875
      TabIndex        =   1
      Top             =   0
      Width           =   4935
      Begin VB.CheckBox chkMatch 
         Caption         =   "Match whole word"
         Height          =   255
         Left            =   2880
         TabIndex        =   6
         Top             =   80
         Value           =   1  'Checked
         Width           =   1695
      End
      Begin VB.TextBox txtSearch 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   0
         TabIndex        =   5
         Top             =   0
         Width           =   2775
      End
      Begin MSComctlLib.ListView lsvSearch 
         Height          =   4215
         Left            =   0
         TabIndex        =   4
         Top             =   360
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   7435
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         OLEDragMode     =   1
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         OLEDragMode     =   1
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Subject"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "ID"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Units"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Descriptive Title"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Offering Department"
            Object.Width           =   5292
         EndProperty
      End
      Begin MSComctlLib.ListView lsvSelect 
         Height          =   4455
         Left            =   0
         TabIndex        =   7
         Top             =   4560
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   7858
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         Icons           =   "ImageList1"
         SmallIcons      =   "ImageList1"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Course"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Curriculum"
            Object.Width           =   3528
         EndProperty
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6360
      Top             =   2640
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSubject.frx":4F18
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSubject.frx":592A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Grid 
      Height          =   4050
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   7144
      _Version        =   393216
      Rows            =   14
      FixedRows       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
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
End
Attribute VB_Name = "frmSubject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim sPreqID As String
Dim SubjectID As String

Dim MaxEntryCount As Long
Dim CurRecPos As Long
Dim CurStudentCount As Long

Public Sub ShowForm(sSubjectID As String)
Dim vRS As New ADODB.Recordset
Dim sSQL As String


sSQL = "SELECT tblSubject.SubjectID AS lvKey, tblSubject.SubjectID, tblSubject.SubjectTitle, tblSubject.Description, tblSubject.StudentCredit, tblSubject.FacultyCredit, tblSubject.Category, tblSubject.Units, tblSubject.LectureUnits, tblSubject.LaboratoryUnits, tblSubject.SubjectFee, tblSubject.LaboratoryFee, tblSubject.RepeatFee, tblDepartment.DepartmentTitle " & _
"FROM tblDepartment INNER JOIN tblSubject ON tblDepartment.DepartmentID = tblSubject.DepartmentID " & _
"WHERE tblSubject.SubjectID ='" & sSubjectID & "'"

If ConnectRS(con, vRS, sSQL) = True Then
        With Grid
            .Clear
        .ClearStructure
        .FixedCols = 1
        
        .ColWidth(0) = 2000
        .ColWidth(1) = 6650

        .TextMatrix(0, 0) = "Subject Code"
        .TextMatrix(1, 0) = "Descriptive Title"
        .TextMatrix(2, 0) = "Units"
        .TextMatrix(3, 0) = "Department"
        .TextMatrix(4, 0) = "Category"
        .TextMatrix(5, 0) = "Student Credit"
        .TextMatrix(6, 0) = "Faculty Credit"
        .TextMatrix(7, 0) = "Subject Fee"
        .TextMatrix(8, 0) = "Laboratory Fee"
        .TextMatrix(9, 0) = "Repeat Fee"
        .TextMatrix(10, 0) = "Lecture Unit"
        .TextMatrix(11, 0) = "Laboratory Unit"
        
        .ColAlignment(0) = vbLeftJustify
        .ColAlignment(1) = vbLeftJustify
        
            .TextMatrix(0, 1) = vRS.Fields("SubjectTitle")
            .Col = 1
            .Row = 0
            .CellBackColor = &H80C0FF
            .TextMatrix(1, 1) = vRS.Fields("Description")
            .TextMatrix(2, 1) = vRS.Fields("Units")
            .TextMatrix(3, 1) = vRS.Fields("DepartmentTitle")
            .TextMatrix(4, 1) = vRS.Fields("Category")
            .TextMatrix(5, 1) = vRS.Fields("StudentCredit")
            .TextMatrix(6, 1) = vRS.Fields("FacultyCredit")
            .TextMatrix(7, 1) = vRS.Fields("SubjectFee")
            .TextMatrix(8, 1) = vRS.Fields("LaboratoryFee")
            .TextMatrix(9, 1) = vRS.Fields("RepeatFee")
            .TextMatrix(10, 1) = vRS.Fields("LectureUnits")
            .TextMatrix(11, 1) = vRS.Fields("LaboratoryUnits")
        End With
        
        SubjectID = vRS.Fields("SubjectID")
    
        RefreshPrereq (SubjectID)
        CourseReq (SubjectID)
        TranscriptReq (SubjectID)
    
Else
     MsgBox "Unable to show Subject properties", vbCritical
     Exit Sub
End If
    Me.Show 1
End Sub

Private Sub lsvSearch_Click()
On Error Resume Next
 CourseSearch (lsvSearch.SelectedItem.SubItems(1))
End Sub

Private Sub lvPreq_OLEDragDrop(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    AddToPrerequisite
    RefreshPrereq (SubjectID)
End Sub

Private Sub txtSearch_Change()
    SearchItem (txtSearch.Text)
End Sub

Private Sub SearchItem(sSubjectID As String)
Dim vRS As New ADODB.Recordset
Dim sSQL As String
Dim lv As ListItem

    sSQL = "SELECT tblSubject.SubjectID AS lvKey, tblSubject.SubjectID, tblSubject.SubjectTitle, tblSubject.Description, tblSubject.StudentCredit, tblSubject.FacultyCredit, tblSubject.Category, tblSubject.Units, tblSubject.LectureUnits, tblSubject.LaboratoryUnits, tblSubject.SubjectFee, tblSubject.LaboratoryFee, tblSubject.RepeatFee, tblDepartment.DepartmentTitle " & _
    "FROM tblDepartment INNER JOIN tblSubject ON tblDepartment.DepartmentID = tblSubject.DepartmentID " & _
    "WHERE tblSubject.SubjectTitle like '%" & sSubjectID & "%'"
    
    If ConnectRS(con, vRS, sSQL) = True Then
            lsvSearch.ListItems.Clear
            Do Until vRS.EOF
                Set lv = lsvSearch.ListItems.Add(, , vRS.Fields("SubjectTitle"))
                    lv.SubItems(1) = vRS.Fields("lvKey")
                    lv.SubItems(2) = vRS.Fields("Units")
                    lv.SubItems(3) = vRS.Fields("Description")
                    lv.SubItems(4) = vRS.Fields("DepartmentTitle")
                vRS.MoveNext
            Loop
    End If
    
Set vRS = Nothing

End Sub
Public Sub Properties(sSubject As String)
Dim vRS As New ADODB.Recordset
Dim sSQL As String


sSQL = "SELECT tblSubject.SubjectID AS lvKey, tblSubject.SubjectID, tblSubject.SubjectTitle, tblSubject.Description, tblSubject.StudentCredit, tblSubject.FacultyCredit, tblSubject.Category, tblSubject.Units, tblSubject.LectureUnits, tblSubject.LaboratoryUnits, tblSubject.SubjectFee, tblSubject.LaboratoryFee, tblSubject.RepeatFee, tblDepartment.DepartmentTitle " & _
"FROM tblDepartment INNER JOIN tblSubject ON tblDepartment.DepartmentID = tblSubject.DepartmentID " & _
"WHERE tblSubject.SubjectTitle ='" & sSubject & "'"

If ConnectRS(con, vRS, sSQL) = True Then
        With Grid
            .Clear
        .ClearStructure
        .FixedCols = 1
        
        .ColWidth(0) = 2000
        .ColWidth(1) = 6650

        .TextMatrix(0, 0) = "Subject Code"
        .TextMatrix(1, 0) = "Descriptive Title"
        .TextMatrix(2, 0) = "Units"
        .TextMatrix(3, 0) = "Department"
        .TextMatrix(4, 0) = "Category"
        .TextMatrix(5, 0) = "Student Credit"
        .TextMatrix(6, 0) = "Faculty Credit"
        .TextMatrix(7, 0) = "Subject Fee"
        .TextMatrix(8, 0) = "Laboratory Fee"
        .TextMatrix(9, 0) = "Repeat Fee"
        .TextMatrix(10, 0) = "Lecture Unit"
        .TextMatrix(11, 0) = "Laboratory Unit"
        
        .ColAlignment(0) = vbLeftJustify
        .ColAlignment(1) = vbLeftJustify
        
            .TextMatrix(0, 1) = vRS.Fields("SubjectTitle")
            .Col = 1
            .Row = 0
            .CellBackColor = &H80C0FF
            
            
            
            .TextMatrix(1, 1) = vRS.Fields("Description")
            .TextMatrix(2, 1) = vRS.Fields("Units")
            .TextMatrix(3, 1) = vRS.Fields("DepartmentTitle")
            .TextMatrix(4, 1) = vRS.Fields("Category")
            .TextMatrix(5, 1) = vRS.Fields("StudentCredit")
            .TextMatrix(6, 1) = vRS.Fields("FacultyCredit")
            .TextMatrix(7, 1) = vRS.Fields("SubjectFee")
            .TextMatrix(8, 1) = vRS.Fields("LaboratoryFee")
            .TextMatrix(9, 1) = vRS.Fields("RepeatFee")
            .TextMatrix(10, 1) = vRS.Fields("LectureUnits")
            .TextMatrix(11, 1) = vRS.Fields("LaboratoryUnits")
        End With

        SubjectID = vRS.Fields("SubjectID")
        
    RefreshPrereq (SubjectID)
    CourseReq (SubjectID)
    TranscriptReq (SubjectID)
Else
     MsgBox "Unable to show Subject properties", vbCritical
     Exit Sub
End If
    Me.Show 1
End Sub

Public Function GetNewPreqID(ByRef sNewPreqID As String) As TranDBResult
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    Dim NewDNumber As Integer
    
    GetNewPreqID = Failed
    
    sSQL = "SELECT 'D-' & String$(2-Len(Count(*)+1),'0') & Count(*)+1 AS NewID" & _
            " FROM tblPrerequisite;"
            
    If ConnectRS(con, vRS, sSQL) = True Then
        
        sNewPreqID = vRS.Fields("NewID").Value
        
        While PreqExistByID(sNewPreqID) = Success
            NewDNumber = Val(Right(sNewPreqID, 2)) + 1
            sNewPreqID = "D-" & String(2 - Len(NewDNumber), "0") & NewDNumber
        Wend
        
        GetNewPreqID = Success
    
    Else
    
        GetNewPreqID = Failed
    End If
    Set vRS = Nothing
End Function
Public Function PreqExistByID(sPreqID As String) As TranDBResult
    
    Dim vRS As New ADODB.Recordset
        
    If ConnectRS(con, vRS, "SELECT * From tblPrerequisite WHERE (((tblPrerequisite.PrerequisiteID)='" & sPreqID & "'));") Then
        If vRS.RecordCount > 0 Then
            PreqExistByID = Success
        Else
            PreqExistByID = Failed
        End If
    Else
        PreqExistByID = Failed
       
    End If
    
    'release
    Set vRS = Nothing
End Function
Private Sub AddToPrerequisite()
On Error GoTo err
Dim sProspectusID As String
Dim vRS As New ADODB.Recordset

If ConnectRS(con, vRS, "Select * From tblPrerequisite") = True Then
    If GetNewPreqID(sPreqID) = Failed Then
        Exit Sub
    End If

        vRS.AddNew
        vRS.Fields("PrerequisiteID") = sPreqID
        vRS.Fields("SubjectID") = SubjectID
        vRS.Fields("RequisiteSubjectID") = lsvSearch.SelectedItem.SubItems(1)
        vRS.Update
End If

err:
    Set vRS = Nothing
End Sub

Private Sub RefreshPrereq(sSubjectID As String)
Dim rs As New ADODB.Recordset
Dim sSQL As String
Dim lv As ListItem

    sSQL = "SELECT tblSubject_1.SubjectID as lvKey, tblSubject.SubjectTitle, tblSubject_1.SubjectTitle AS Prerequisite, tblSubject_1.Description, tblSubject_1.Units " & _
            "FROM tblSubject AS tblSubject_1 INNER JOIN (tblSubject INNER JOIN tblPrerequisite ON tblSubject.SubjectID = tblPrerequisite.SubjectID) ON tblSubject_1.SubjectID = tblPrerequisite.RequisiteSubjectID " & _
            "WHERE tblSubject.SubjectID = '" & sSubjectID & "'"
 
 On Error GoTo err
    
    lvPreq.ListItems.Clear
    
    If ConnectRS(con, rs, sSQL) = True Then
    Do Until rs.EOF
        Set lv = lvPreq.ListItems.Add(, , rs.Fields("Prerequisite"), , 1)
                lv.SubItems(1) = rs.Fields("lvKey")
                lv.SubItems(2) = rs.Fields("Description")
                lv.SubItems(3) = rs.Fields("Units")
        rs.MoveNext
    Loop
    End If
Exit Sub
err:
    Set rs = Nothing
End Sub
Private Sub CourseReq(sSubjectID As String)
Dim rs As New ADODB.Recordset
Dim mySQL As String

    mySQL = "SELECT tblSubject.SubjectID as lvKey, tblCourse.Course, tblCourse.Curriculum " & _
            "FROM tblSubject INNER JOIN (tblSubjectOffering INNER JOIN (tblCourse INNER JOIN (tblEnrolment INNER JOIN tblGrade ON tblEnrolment.EnrollmentID = tblGrade.EnrolmentID) ON tblCourse.CourseID = tblEnrolment.CourseID) ON (tblSubjectOffering.SubjectOfferingID = tblGrade.SubjectOfferingID) AND (tblSubjectOffering.SubjectOfferingID = tblGrade.SubjectOfferingID)) ON tblSubject.SubjectID = tblSubjectOffering.SubjectID " & _
            "WHERE tblSubject.SubjectID='" & sSubjectID & "'"

If ConnectRS(con, rs, mySQL) = True Then
        UnSortLV lsvCourse
        
        FillRecordToList rs, lsvCourse, KeyStudent, CurRecPos, MaxEntryCount, 3, True
        
        SortLV lsvCourse, lsvCourse.SortKey, lsvCourse.SortOrder, False
End If
Set rs = Nothing
End Sub
Private Sub TranscriptReq(sSubjectID As String)
Dim rs As New ADODB.Recordset
Dim mySQL As String

    mySQL = "SELECT tblSubject.SubjectID, tblStudent.LastName&', '&tblStudent.FirstName&' '&tblStudent.MiddleName as Fullname, tblStudent.StudentID " & _
            "FROM tblStudent INNER JOIN (tblSubject INNER JOIN (tblSubjectOffering INNER JOIN (tblEnrolment INNER JOIN tblGrade ON tblEnrolment.EnrollmentID = tblGrade.EnrolmentID) ON (tblSubjectOffering.SubjectOfferingID = tblGrade.SubjectOfferingID) AND (tblSubjectOffering.SubjectOfferingID = tblGrade.SubjectOfferingID)) ON tblSubject.SubjectID = tblSubjectOffering.SubjectID) ON tblStudent.StudentID = tblEnrolment.StudentID " & _
            "WHERE tblSubject.SubjectID='" & sSubjectID & "'"

If ConnectRS(con, rs, mySQL) = True Then
        UnSortLV lsvTranscript
        
        FillRecordToList rs, lsvTranscript, KeyStudent, CurRecPos, MaxEntryCount, 3, True
        
        SortLV lsvTranscript, lsvTranscript.SortKey, lsvTranscript.SortOrder, False
End If
Set rs = Nothing
End Sub


Private Sub CourseSearch(sSubjectID As String)
Dim rs As New ADODB.Recordset
Dim mySQL As String

    mySQL = "SELECT tblSubject.SubjectID as lvKey, tblCourse.Course, tblCourse.Curriculum " & _
            "FROM tblSubject INNER JOIN (tblSubjectOffering INNER JOIN (tblCourse INNER JOIN (tblEnrolment INNER JOIN tblGrade ON tblEnrolment.EnrollmentID = tblGrade.EnrolmentID) ON tblCourse.CourseID = tblEnrolment.CourseID) ON (tblSubjectOffering.SubjectOfferingID = tblGrade.SubjectOfferingID) AND (tblSubjectOffering.SubjectOfferingID = tblGrade.SubjectOfferingID)) ON tblSubject.SubjectID = tblSubjectOffering.SubjectID " & _
            "WHERE tblSubject.SubjectID='" & sSubjectID & "'"

If ConnectRS(con, rs, mySQL) = True Then
        UnSortLV lsvSelect
        
        FillRecordToList rs, lsvSelect, KeyStudent, CurRecPos, MaxEntryCount, 3, True
        
        SortLV lsvSelect, lsvSelect.SortKey, lsvSelect.SortOrder, False
End If
Set rs = Nothing
End Sub

