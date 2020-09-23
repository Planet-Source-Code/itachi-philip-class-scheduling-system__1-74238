VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmCurriculum 
   Caption         =   "Curriculum"
   ClientHeight    =   8520
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   11400
   Icon            =   "frmCurriculum.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8520
   ScaleWidth      =   11400
   WindowState     =   2  'Maximized
   Begin TabDlg.SSTab SSTab1 
      Height          =   11175
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15255
      _ExtentX        =   26908
      _ExtentY        =   19711
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   2
      TabHeight       =   529
      BackColor       =   14215660
      TabCaption(0)   =   "Prospectus"
      TabPicture(0)   =   "frmCurriculum.frx":492A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "imgListEnrolment"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "ilRecordIco"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "icoHeader"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "tvCurriculum"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lsvSubject"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cboSemesterID"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "CourseID"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "FrameProspectus"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).ControlCount=   8
      Begin VB.Frame FrameProspectus 
         Height          =   615
         Left            =   0
         TabIndex        =   3
         Top             =   360
         Width           =   15135
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
            Left            =   840
            TabIndex        =   9
            Top             =   160
            Width           =   2415
         End
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
            Left            =   3360
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   160
            Width           =   5535
         End
         Begin VB.ComboBox cboCurriculum 
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
            Left            =   8880
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   160
            Width           =   1575
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
            ItemData        =   "frmCurriculum.frx":4946
            Left            =   10440
            List            =   "frmCurriculum.frx":4953
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   160
            Width           =   1935
         End
         Begin VB.CommandButton cmdShowCurriculum 
            Appearance      =   0  'Flat
            Height          =   375
            Left            =   12405
            Picture         =   "frmCurriculum.frx":497B
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Search Curriculum"
            Top             =   160
            Width           =   390
         End
         Begin VB.CommandButton cmdNewCurriculum 
            Appearance      =   0  'Flat
            Height          =   375
            Left            =   12840
            Picture         =   "frmCurriculum.frx":537D
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "New Curriculum"
            Top             =   160
            Width           =   1815
         End
         Begin VB.Label Label1 
            Caption         =   "Search"
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
            Left            =   120
            TabIndex        =   10
            Top             =   240
            Width           =   735
         End
      End
      Begin VB.ComboBox CourseID 
         Height          =   315
         Left            =   720
         TabIndex        =   2
         Top             =   6960
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.ComboBox cboSemesterID 
         Height          =   315
         Left            =   1560
         TabIndex        =   1
         Top             =   6960
         Visible         =   0   'False
         Width           =   855
      End
      Begin MSComctlLib.ListView lsvSubject 
         Height          =   4575
         Left            =   0
         TabIndex        =   11
         Top             =   960
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   8070
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         OLEDragMode     =   1
         OLEDropMode     =   1
         FullRowSelect   =   -1  'True
         _Version        =   393217
         Icons           =   "ilRecordIco"
         SmallIcons      =   "ilRecordIco"
         ColHdrIcons     =   "icoHeader"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         OLEDragMode     =   1
         OLEDropMode     =   1
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Subject"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Subject"
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
            Object.Width           =   7056
         EndProperty
      End
      Begin MSComctlLib.TreeView tvCurriculum 
         Height          =   6495
         Left            =   5880
         TabIndex        =   12
         Top             =   960
         Width           =   8295
         _ExtentX        =   14631
         _ExtentY        =   11456
         _Version        =   393217
         LabelEdit       =   1
         Style           =   7
         FullRowSelect   =   -1  'True
         ImageList       =   "imgListEnrolment"
         Appearance      =   1
         Enabled         =   0   'False
         OLEDropMode     =   1
      End
      Begin MSComctlLib.ImageList icoHeader 
         Left            =   1560
         Top             =   6000
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
               Picture         =   "frmCurriculum.frx":5D7F
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCurriculum.frx":6319
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageList ilRecordIco 
         Left            =   2280
         Top             =   6000
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
               Picture         =   "frmCurriculum.frx":68B3
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageList imgListEnrolment 
         Left            =   3480
         Top             =   6240
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
               Picture         =   "frmCurriculum.frx":6E4D
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "frmCurriculum"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SubjRs As New ADODB.Recordset

Dim MaxEntryCount As Long
Dim CurRecPos As Long
Dim CurStudentCount As Long

Dim SelectedItem As String
Dim IsStarted As Boolean

Private Const KeyYearLevel = "year"
Private Const keySemester = "Sem"
Private Const KeySubject = "subj"


Dim slSemesterTitle() As String
Dim slYearLevelTitle() As String
Dim slSubjectTitle() As String

Dim curYearLevelTitle As String
Dim curDepartmentTitle As String
Dim curSubjectTitle As String

Public Sub ShowFormList(Optional iMaxEntryCount As Long = 21, Optional iCurRecPos As Long = 0)
    Dim SubjSQL As String
    

    'MaxEntryCount = iMaxEntryCount
    'CurRecPos = iCurRecPos
    
    SubjSQL = "SELECT tblSubject.SubjectID AS lvKey, tblSubject.SubjectTitle AS Title,tblSubject.SubjectID,tblSubject.Units, tblSubject.Description, tblDepartment.DepartmentTitle AS Department FROM tblDepartment INNER JOIN tblSubject ON tblDepartment.DepartmentID = tblSubject.DepartmentID"
       
       
      If ConnectRS(con, SubjRs, SubjSQL) = True Then
        FillSubjectList SubjRs
        SemesterList
        CourseList
        cboSemester.Text = CurrentSemester.Semester
        cboSemester_Change
        Me.Show
     Else
    
        MsgBox "Unable to show Teacher list.", vbCritical
        Unload Me
    End If
End Sub
Public Sub FormSubject_Refresh()
    SubjRs.Requery
    FillSubjectList SubjRs
End Sub


Private Function FillSubjectList(ByRef vRS As ADODB.Recordset)
        mdiController.MousePointer = vbHourglass
        
        UnSortLV lsvSubject
        FillRecordToList vRS, lsvSubject, KeyStudent, , , , True
        SortLV lsvSubject, lsvSubject.SortKey, lsvSubject.SortOrder, False
          
        mdiController.MousePointer = vbDefault
End Function

Public Sub FormDepartment_Refresh()
    SubjRs.Requery
    FillSubjectList SubjRs
End Sub

Private Sub cboCourse_Change()
    CourseID.ListIndex = cboCourse.ListIndex
    CourseCurriculumList (cboCourse.Text)
End Sub

Private Sub cboCourse_Click()
    cboCourse_Change
End Sub

Private Sub cboSemester_Change()
    cboSemesterID.ListIndex = cboSemester.ListIndex
     Refresh_Prospectus
End Sub

Private Sub cboSemester_Click()
    cboSemester_Change
End Sub

Private Sub cmdNewCurriculum_Click()
    frmCourse.ShowForm
End Sub

Private Sub cmdShowCurriculum_Click()
    If Len(cboCourse.Text) < 0 And Len(cboCurriculum.Text) < 0 Then
        MsgBox "Please select course and curriculum", vbCritical
        Exit Sub
    End If

    Refresh_Prospectus
    
    tvCurriculum.Enabled = True
    lsvSubject.Enabled = True
End Sub

Private Sub Form_Resize()
    SSTab1.Width = ScaleWidth
    SSTab1.Height = ScaleHeight
    
    FrameProspectus.Width = SSTab1.Width
    lsvSubject.Height = SSTab1.Height - (FrameProspectus.Height + 400)
    
    tvCurriculum.Height = SSTab1.Height - (FrameProspectus.Height + 400)
    tvCurriculum.Width = SSTab1.Width - lsvSubject.Width
End Sub
Private Sub lsvSubject_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
      PopupMenu mdiController.popSubjectCur
    End If
End Sub

Private Sub tvCurriculum_Click()
    'Call FolderClick(tvCurriculum.SelectedItem, Left(tvCurriculum.SelectedItem.Key, 4))
End Sub

Private Sub tvCurriculum_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
       PopupMenu mdiController.popCurriculum
    End If
    
On Error Resume Next
    SelectedItem = tvCurriculum.SelectedItem.Text
End Sub
Public Sub FolderClick(fNode As Node, sRecordType As String)
    Dim splitKey() As String
    Dim sKey() As String
    Dim sText() As String
   
    splitKey = Split(fNode.Key, ";")
    
    Select Case splitKey(0)
        Case KeySubject
            ShowSubjectProperties splitKey(3)
    End Select
End Sub
Private Sub FolderDelete(fNode As MSComctlLib.Node, sRecordType As String)
    Dim splitKey() As String
    Dim sKey() As String
    Dim sText() As String
    
    
    splitKey = Split(fNode.Key, ";")
    
    Select Case splitKey(0)
        Case KeySubject
            
    End Select
End Sub
Public Function SelectNode(sKey As String) As Variant

    Dim tNode As Node
    
    If IsStarted = False Then
        Refresh_Prospectus
    End If
    
    For Each tNode In tvCurriculum.Nodes
        If tNode.Key = sKey Then

            tNode.Selected = True
            tvCurriculum_Click
        End If
    Next

End Function

Public Function GetSchoolYearChilds(sSchoolYearTitle As String, ByRef sKey() As String, ByRef sText() As String) As Boolean
    Dim tNode As Node
    Dim NodeCount As Integer
    Dim i As Integer
    Dim splitKey() As String
    
    NodeCount = 0
    
    For Each tNode In tvCurriculum.Nodes
        splitKey = Split(tNode.Key, ";")
        If splitKey(0) = keySemester And splitKey(1) = sSchoolYearTitle Then
            NodeCount = NodeCount + 1
        End If
    Next
    
    If NodeCount < 1 Then
        GetSchoolYearChilds = False
        Exit Function
    End If
    
    ReDim sKey(NodeCount - 1)
    ReDim sText(NodeCount - 1)

    i = 0
    For Each tNode In tvCurriculum.Nodes
        splitKey = Split(tNode.Key, ";")
        If splitKey(0) = keySemester And splitKey(1) = sSchoolYearTitle Then
             sKey(i) = tNode.Key
             sText(i) = tNode.Text
            i = i + 1
        End If
    Next
    
    GetSchoolYearChilds = True
End Function
Public Function GetDepartmentChilds(sSchoolYearTitle As String, sDepartmentTitle As String, ByRef sKey() As String, ByRef sText() As String) As Boolean
    Dim tNode As Node
    Dim NodeCount As Integer
    Dim i As Integer
    Dim splitKey() As String
    
    NodeCount = 0
    
    For Each tNode In tvCurriculum.Nodes
        splitKey = Split(tNode.Key, ";")
        If splitKey(0) = KeyYearLevel Then
            If splitKey(1) = sSchoolYearTitle And splitKey(2) = sDepartmentTitle Then
                NodeCount = NodeCount + 1
            End If
        End If
    Next
    
    If NodeCount < 1 Then
        GetDepartmentChilds = False
        Exit Function
    End If
    
    ReDim sKey(NodeCount - 1)
    ReDim sText(NodeCount - 1)
    
    i = 0
    For Each tNode In tvCurriculum.Nodes
        splitKey = Split(tNode.Key, ";")
        If splitKey(0) = KeyYearLevel Then
            If splitKey(1) = sSchoolYearTitle And splitKey(2) = sDepartmentTitle Then
                sKey(i) = tNode.Key
             sText(i) = tNode.Text
                i = i + 1
            End If
        End If
    Next
    
    GetDepartmentChilds = True
End Function
Public Function GetYearLevelChilds(sSchoolYearTitle As String, sDepartmentTitle As String, sYearLevelTitle As String, ByRef sKey() As String, ByRef sText() As String) As Boolean
    Dim tNode As Node
    Dim NodeCount As Integer
    Dim i As Integer
    Dim splitKey() As String
    
    NodeCount = 0
    
    For Each tNode In tvCurriculum.Nodes
        splitKey = Split(tNode.Key, ";")
        If splitKey(0) = KeySectionOffering Then
            If splitKey(1) = sSchoolYearTitle And splitKey(2) = sDepartmentTitle And splitKey(3) = sYearLevelTitle Then
                NodeCount = NodeCount + 1
            End If
        End If
    Next
    If NodeCount < 1 Then
        GetYearLevelChilds = False
        Exit Function
    End If
    
    ReDim sKey(NodeCount - 1)
    ReDim sText(NodeCount - 1)
    
    i = 0
    For Each tNode In tvCurriculum.Nodes
        splitKey = Split(tNode.Key, ";")
        If splitKey(0) = KeySectionOffering Then
            If splitKey(1) = sSchoolYearTitle And splitKey(2) = sDepartmentTitle And splitKey(3) = sYearLevelTitle Then
                sKey(i) = tNode.Key
             sText(i) = tNode.Text
                i = i + 1
            End If
        End If
    Next
    GetYearLevelChilds = True
End Function
Public Sub ShowSubjectProperties(sSubjectTitle As String)
    curSubjectTitle = sSubjectTitle
    frmSubject.Properties curSubjectTitle
End Sub
Private Function SetSelectedSection(sSectionTitle As String, Optional sSchoolYearTitle As String = "")
    Dim tNode As Node
    Dim splitKey() As String


    For Each tNode In tvCurriculum.Nodes
    
        If tNode.Text = sSectionTitle And Left(tNode.Key, 4) = keySection Then
            
            splitKey = Split(tNode.Key, ";")
            
            If sSchoolYearTitle = "" Then
                tNode.Selected = True
                tNode.EnsureVisible
                Exit For
            Else
            
                If splitKey(1) = sSchoolYearTitle Then
                    tNode.Selected = True
                    tNode.EnsureVisible
                    Exit For
                End If
            End If
            
        End If
        
    Next
End Function

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

Private Sub CourseList()
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    
    sSQL = "SELECT tblCourse.Course,tblCourse.CourseID" & _
            " FROM tblCourse"
     
    If ConnectRS(con, vRS, sSQL) = True Then
        cboCurriculum.Clear
        Do Until vRS.EOF
            cboCourse.AddItem (vRS.Fields("Course"))
            CourseID.AddItem (vRS.Fields("CourseID"))
            vRS.MoveNext
        Loop
    End If
    
ReleaseAndExit:
    Set vRS = Nothing
End Sub

Private Sub CourseCurriculumList(sCourse As String)
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    
    sSQL = "SELECT tblCourse.Curriculum" & _
            " FROM tblCourse" & _
            " WHERE tblCourse.Course='" & sCourse & "'"
     
    If ConnectRS(con, vRS, sSQL) = True Then
        cboCurriculum.Clear
        Do Until vRS.EOF
            cboCurriculum.AddItem (vRS.Fields("Curriculum"))
            vRS.MoveNext
        Loop
    End If
    
ReleaseAndExit:
    Set vRS = Nothing
End Sub


Private Function Refresh_YearLevel()
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    Dim i As Integer
    
    'clear tree
    tvCurriculum.Nodes.Clear
    
    sSQL = "SELECT tblYearLevel.YearLevelTitle" & _
            " FROM tblYearLevel;"
    
    If ConnectRS(con, vRS, sSQL) <> True Then
        GoTo RealeaseAndExit
    End If
    
    If AnyRecordExisted(vRS) = False Then
        GoTo RealeaseAndExit
    End If
    
    ReDim slYearLevelTitle(getRecordCount(vRS) - 1)
    
    i = 0
    vRS.MoveFirst
    
    While vRS.EOF = False
        
        slYearLevelTitle(i) = vRS.Fields("YearLevelTitle")
        AddYearLevelToTree slYearLevelTitle(i)
        
        i = i + 1
        vRS.MoveNext
    Wend
    
RealeaseAndExit:
    Set vRS = Nothing
End Function

Private Function Refresh_Department()
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    Dim i As Integer
    Dim ii As Integer
    
    sSQL = "SELECT tblSemester.Semester" & _
            " FROM tblSemester"

    If ConnectRS(con, vRS, sSQL) <> True Then
        GoTo RealeaseAndExit
    End If
    
    If AnyRecordExisted(vRS) = False Then
        GoTo RealeaseAndExit
    End If
        
    ReDim slSemesterTitle(getRecordCount(vRS) - 1)
    i = 0
    vRS.MoveFirst
    
    While vRS.EOF = False
    
        slSemesterTitle(i) = vRS.Fields("Semester")
        
        For ii = 0 To UBound(slYearLevelTitle)
            AddDepartmentToTree slYearLevelTitle(ii), slSemesterTitle(i)
        Next
        
        i = i + 1
        vRS.MoveNext
    Wend
RealeaseAndExit:
    Set vRS = Nothing
End Function

Private Function Refresh_Subject(sCourseID As String)

    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    Dim i As Integer
    Dim ii As Integer
    Dim iii As Integer

    sSQL = "SELECT tblSubject.SubjectTitle, tblCourse.Course & ' major in ' & tblCourse.Major AS Course, tblCourse.Curriculum, tblProspectus.YearLevel, tblSemester.Semester,tblSubject.SubjectID " & _
            "FROM tblSemester INNER JOIN (tblSubject INNER JOIN (tblCourse INNER JOIN tblProspectus ON tblCourse.CourseID = tblProspectus.CourseID) ON tblSubject.SubjectID = tblProspectus.SubjectID) ON tblSemester.SemesterID = tblProspectus.SemesterID " & _
            "WHERE tblCourse.CourseID='" & sCourseID & "'"

    If ConnectRS(con, vRS, sSQL) <> True Then
        GoTo RealeaseAndExit
    End If
    
    If AnyRecordExisted(vRS) = False Then
        GoTo RealeaseAndExit
    End If
        
    ReDim slSubjectTitle(getRecordCount(vRS))
    i = 0
    vRS.MoveFirst
    
    While vRS.EOF = False
        
        slSubjectTitle(i) = vRS.Fields("SubjectTitle")
        
        For ii = 0 To UBound(slYearLevelTitle)
            For iii = 0 To UBound(slSemesterTitle)
                AddSubjectToTree vRS.Fields("YearLevel"), vRS.Fields("Semester"), slSubjectTitle(i)
            Next
        Next
        i = i + 1
        vRS.MoveNext
    Wend

RealeaseAndExit:
    Set vRS = Nothing
End Function

Private Function AddSubjectToTree(sSchoolYearTitle As String, sDepartmentTitle As String, sYearLevelTitle As String)
    Dim tNode As Node
    
    For Each tNode In tvCurriculum.Nodes
        If tNode.Key = KeySubject & ";" & sSchoolYearTitle & ";" & sDepartmentTitle & ";" & Trim(sYearLevelTitle) Then
            Exit Function
        End If
    Next
    
    tvCurriculum.Nodes.Add keySemester & ";" & sSchoolYearTitle & ";" & sDepartmentTitle, tvwChild, KeySubject & ";" & sSchoolYearTitle & ";" & sDepartmentTitle & ";" & Trim(sYearLevelTitle), Trim(sYearLevelTitle), 1

End Function
Private Function AddYearLevelToTree(sLevelYearTitle As String)
    Dim tNode As Node
    
    For Each tNode In tvCurriculum.Nodes
        If tNode.Key = KeyYearLevel & ";" & sLevelYearTitle Then
            Exit Function
        End If
    Next
    
    tvCurriculum.Nodes.Add , , KeyYearLevel & ";" & sLevelYearTitle, sLevelYearTitle, 1
End Function

Private Function AddDepartmentToTree(sSchoolYearTitle As String, sDepartmentTitle As String)
    Dim tNode As Node
    
    For Each tNode In tvCurriculum.Nodes
        If tNode.Key = keySemester & ";" & sSchoolYearTitle & ";" & sDepartmentTitle Then
            Exit Function
        End If
    Next
    
    tvCurriculum.Nodes.Add KeyYearLevel & ";" & sSchoolYearTitle, tvwChild, keySemester & ";" & sSchoolYearTitle & ";" & sDepartmentTitle, sDepartmentTitle, 1
End Function
Private Sub tvCurriculum_OLEDragDrop(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)

If tvCurriculum.Nodes.count < 1 Then
    MsgBox "Please select COURSE first before adding subject", vbCritical
    Exit Sub
End If

    frmProspectusAE.ShowForm lsvSubject.SelectedItem
End Sub
Public Sub Refresh_Prospectus()
    Dim expAll As Integer
    
    Refresh_YearLevel
    Refresh_Department
    Refresh_Subject (CourseID.Text)
    
        For expAll = 1 To tvCurriculum.Nodes.count
            tvCurriculum.Nodes(expAll).Expanded = True
        Next
End Sub

Private Sub txtSearch_Change()
    If Len(txtSearch) < 1 Then
        Search
    Else
        SearchItem (txtSearch.Text)
    End If
End Sub
Private Sub SearchItem(sSubjectID As String)
Dim vRS As New ADODB.Recordset
Dim sSQL As String
Dim lv As ListItem

    sSQL = "SELECT tblSubject.SubjectID AS lvKey, tblSubject.SubjectTitle AS Title,tblSubject.SubjectID,tblSubject.Units, tblSubject.Description, tblDepartment.DepartmentTitle AS Department FROM tblDepartment INNER JOIN tblSubject ON tblDepartment.DepartmentID = tblSubject.DepartmentID" & _
            " WHERE tblSubject.SubjectTitle like '%" & sSubjectID & "%'"
    
    If ConnectRS(con, vRS, sSQL) = True Then
        FillSubjectList vRS
    End If
Set vRS = Nothing
End Sub

Private Sub Search()
Dim vRS As New ADODB.Recordset
Dim sSQL As String
Dim lv As ListItem

    sSQL = "SELECT tblSubject.SubjectID AS lvKey, tblSubject.SubjectTitle AS Title,tblSubject.SubjectID,tblSubject.Units, tblSubject.Description, tblDepartment.DepartmentTitle AS Department FROM tblDepartment INNER JOIN tblSubject ON tblDepartment.DepartmentID = tblSubject.DepartmentID"
    
    If ConnectRS(con, vRS, sSQL) = True Then
        FillSubjectList vRS
    End If
Set vRS = Nothing
End Sub

Public Sub Subject_Delete()
    Dim vSubject As tSubject
    Dim sMSG As String
    Dim sLVItemKey As String
    
    If tvCurriculum.Nodes.count < 1 Then
        MsgBox "No selected entry to delete." & _
            vbNewLine & "Please select it first in the list.", vbExclamation
        Exit Sub
    End If
    

        sMSG = "WARNING:" & vbNewLine & "You are about to delete this Subject Entries and you cannot Undo this operation." & vbNewLine & _
        " Are you sure to delete it?"
        
    If MsgBox(sMSG, vbQuestion + vbYesNo) = vbYes Then
        
            sLVItemKey = tvCurriculum.SelectedItem
                        
                        If GetSubjectByTitle(sLVItemKey, vSubject) <> Success Then
                            Exit Sub
                        End If
                        
                        If DeleteProspectus(vSubject.SubjectID, CourseID.Text) = Success Then
                            MsgBox "Subject record succesfully deleted.", vbInformation
                           Me.Refresh_Prospectus
                        Else
                            MsgBox "Unable to delete Subject Entry with ID: " & sLVItemKey
                        End If
    End If
End Sub

Public Function DeleteProspectus(sSubjectID As String, sCourseID As String) As TranDBResult
    
    Dim vRS As New ADODB.Recordset
    DeleteProspectus = Failed
    
    If ConnectRS(con, vRS, "Delete * From tblProspectus WHERE tblProspectus.SubjectID='" & sSubjectID & "' and tblProspectus.CourseID='" & sCourseID & "'") Then
        DeleteProspectus = Success
    Else
        DeleteProspectus = Success
    End If
    Set vRS = Nothing
End Function

