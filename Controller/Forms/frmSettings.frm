VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmSettings 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Options"
   ClientHeight    =   8175
   ClientLeft      =   2565
   ClientTop       =   1500
   ClientWidth     =   12405
   Icon            =   "frmSettings.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8175
   ScaleWidth      =   12405
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   8175
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   12375
      _ExtentX        =   21828
      _ExtentY        =   14420
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Faculty"
      TabPicture(0)   =   "frmSettings.frx":492A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lsvFaculty"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "ilRecordIco"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "icoHeader"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Subject"
      TabPicture(1)   =   "frmSettings.frx":4946
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lsvSubject"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "imgSubject"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "imgDepartment"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "lsvDepartment"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "txtDepartmentID"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).ControlCount=   5
      TabCaption(2)   =   "Class Room"
      TabPicture(2)   =   "frmSettings.frx":4962
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "lsvDepartmentRoom"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "imgRoom"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "lsvClassroom"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).ControlCount=   3
      TabCaption(3)   =   "Directory"
      TabPicture(3)   =   "frmSettings.frx":497E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "imgListEnrolment"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "tvDirectory"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).ControlCount=   2
      Begin VB.TextBox txtDepartmentID 
         Height          =   285
         Left            =   -66480
         TabIndex        =   14
         Top             =   3600
         Visible         =   0   'False
         Width           =   615
      End
      Begin MSComctlLib.ListView lsvClassroom 
         Height          =   7335
         Left            =   -75000
         TabIndex        =   8
         Top             =   360
         Width           =   4815
         _ExtentX        =   8493
         _ExtentY        =   12938
         View            =   3
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         Icons           =   "imgRoom"
         SmallIcons      =   "imgRoom"
         ColHdrIcons     =   "icoHeader"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Building"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Room"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Capacity"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Object.Width           =   0
         EndProperty
      End
      Begin MSComctlLib.ListView lsvDepartment 
         Height          =   7335
         Left            =   -75000
         TabIndex        =   7
         Top             =   360
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   12938
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         Icons           =   "imgDepartment"
         SmallIcons      =   "imgDepartment"
         ColHdrIcons     =   "icoHeader"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Department"
            Object.Width           =   7223
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "DepartmentID"
            Object.Width           =   0
         EndProperty
      End
      Begin MSComctlLib.ImageList icoHeader 
         Left            =   5160
         Top             =   6060
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
               Picture         =   "frmSettings.frx":499A
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmSettings.frx":4F34
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageList ilRecordIco 
         Left            =   2760
         Top             =   6420
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
               Picture         =   "frmSettings.frx":54CE
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageList imgDepartment 
         Left            =   -66240
         Top             =   660
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
               Picture         =   "frmSettings.frx":5A68
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageList imgSubject 
         Left            =   -66360
         Top             =   1380
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
               Picture         =   "frmSettings.frx":6002
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageList imgRoom 
         Left            =   -64560
         Top             =   300
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
               Picture         =   "frmSettings.frx":659C
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ListView lsvSubject 
         Height          =   7335
         Left            =   -70920
         TabIndex        =   9
         Top             =   360
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   12938
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         Icons           =   "imgSubject"
         SmallIcons      =   "imgSubject"
         ColHdrIcons     =   "icoHeader"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Subject"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Units"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Descriptive Title"
            Object.Width           =   5292
         EndProperty
      End
      Begin MSComctlLib.ListView lsvFaculty 
         Height          =   5895
         Left            =   0
         TabIndex        =   10
         Top             =   360
         Width           =   9375
         _ExtentX        =   16536
         _ExtentY        =   10398
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         Icons           =   "ilRecordIco"
         SmallIcons      =   "ilRecordIco"
         ColHdrIcons     =   "icoHeader"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Fullname"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "ID Number"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Gender"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Active"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ImageList imgListEnrolment 
         Left            =   -70560
         Top             =   5880
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
               Picture         =   "frmSettings.frx":6B36
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.TreeView tvDirectory 
         Height          =   7695
         Left            =   -75000
         TabIndex        =   12
         Top             =   360
         Width           =   12240
         _ExtentX        =   21590
         _ExtentY        =   13573
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   423
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         ImageList       =   "imgListEnrolment"
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSComctlLib.ListView lsvDepartmentRoom 
         Height          =   7335
         Left            =   -70200
         TabIndex        =   13
         Top             =   360
         Width           =   5535
         _ExtentX        =   9763
         _ExtentY        =   12938
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         Icons           =   "imgDepartment"
         SmallIcons      =   "imgDepartment"
         ColHdrIcons     =   "icoHeader"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Department"
            Object.Width           =   9763
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "DepartmentID"
            Object.Width           =   0
         EndProperty
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   3
      Left            =   -20000
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample4 
         Caption         =   "Sample 4"
         Height          =   1785
         Left            =   2100
         TabIndex        =   5
         Top             =   840
         Width           =   2055
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   2
      Left            =   -20000
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample3 
         Caption         =   "Sample 3"
         Height          =   1785
         Left            =   1545
         TabIndex        =   4
         Top             =   675
         Width           =   2055
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   1
      Left            =   -20000
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample2 
         Caption         =   "Sample 2"
         Height          =   1785
         Left            =   645
         TabIndex        =   3
         Top             =   300
         Width           =   2055
         Begin MSComctlLib.TreeView tvFolder 
            Height          =   7575
            Left            =   0
            TabIndex        =   11
            Top             =   0
            Width           =   3000
            _ExtentX        =   5292
            _ExtentY        =   13361
            _Version        =   393217
            HideSelection   =   0   'False
            Indentation     =   423
            LabelEdit       =   1
            LineStyle       =   1
            Style           =   7
            ImageList       =   "imgListEnrolment"
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MSComctlLib.ImageList imgListIco32 
            Left            =   0
            Top             =   0
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   -2147483643
            ImageWidth      =   32
            ImageHeight     =   32
            MaskColor       =   12632256
            _Version        =   393216
            BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
               NumListImages   =   1
               BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmSettings.frx":70D0
                  Key             =   ""
               EndProperty
            EndProperty
         End
      End
   End
   Begin VB.Menu popFaculty 
      Caption         =   "Faculty"
      Visible         =   0   'False
      Begin VB.Menu mnuNewFaculty 
         Caption         =   "New faculty..."
      End
      Begin VB.Menu mnuEditFaculty 
         Caption         =   "Edit faculty..."
      End
      Begin VB.Menu mnuDeleteFaculty 
         Caption         =   "Delete"
      End
      Begin VB.Menu mnusep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuActivate 
         Caption         =   "Activate"
      End
      Begin VB.Menu mnusep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRefresh 
         Caption         =   "Refresh"
      End
   End
   Begin VB.Menu popDepartment 
      Caption         =   "Department"
      Visible         =   0   'False
      Begin VB.Menu mnuprintsubjects 
         Caption         =   "Print subjects..."
      End
      Begin VB.Menu mnusep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDepartmentRefresh 
         Caption         =   "Refresh"
      End
   End
   Begin VB.Menu popClassroom 
      Caption         =   "Classroom"
      Visible         =   0   'False
      Begin VB.Menu mnuNewClassroom 
         Caption         =   "New classroom..."
      End
      Begin VB.Menu mnuEditClassroom 
         Caption         =   "Edit classroom..."
      End
      Begin VB.Menu mnuDeleteClassroom 
         Caption         =   "Delete"
      End
      Begin VB.Menu mnusep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRoomRefresh 
         Caption         =   "Refresh"
      End
   End
   Begin VB.Menu popSubject 
      Caption         =   "Subject"
      Visible         =   0   'False
      Begin VB.Menu mnuNewSubject 
         Caption         =   "New subject.."
      End
      Begin VB.Menu mnuEditSubject 
         Caption         =   "Edit Subject..."
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "Delete"
      End
      Begin VB.Menu mnusep5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSubjectRefresh 
         Caption         =   "Refresh"
      End
   End
   Begin VB.Menu popDepartmentAE 
      Caption         =   "DepartmentAE"
      Visible         =   0   'False
      Begin VB.Menu mnuNewCollege 
         Caption         =   "New college..."
      End
      Begin VB.Menu mnuEditCollege 
         Caption         =   "Edit college..."
      End
      Begin VB.Menu mnuDeleteCollege 
         Caption         =   "Delete"
      End
      Begin VB.Menu mnusep7 
         Caption         =   "-"
      End
      Begin VB.Menu mnuNewDepartment 
         Caption         =   "New department..."
      End
      Begin VB.Menu mnuEditDepartment 
         Caption         =   "Edit department..."
      End
      Begin VB.Menu mnuDeleteDepartment 
         Caption         =   "Delete"
      End
      Begin VB.Menu mnusep8 
         Caption         =   "-"
      End
      Begin VB.Menu mnuNewCourse 
         Caption         =   "New course..."
      End
      Begin VB.Menu mnuEditCourse 
         Caption         =   "Edit course..."
      End
      Begin VB.Menu mnuDeleteCourse 
         Caption         =   "Delete"
      End
      Begin VB.Menu mnusep9 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDepartmentAERefresh 
         Caption         =   "Refresh"
      End
   End
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vRS As New ADODB.Recordset
Dim roomRs As New ADODB.Recordset
Dim DeptRs As New ADODB.Recordset
Dim SubjRs As New ADODB.Recordset

Dim sDefaultSQL As String

Dim MaxEntryCount As Long
Dim CurRecPos As Long
Dim CurStudentCount As Long

Dim DepartmentID As String


Dim IsStarted As Boolean

Private Const keyCollege = "col"
Private Const keyDepartment = "dept"
Private Const KeyCourse = "cour"

Dim curCollegeTitle As String
Dim curDepartmentTitle As String
Dim curCourseTitle As String

Dim slCollegeTitle() As String
Dim slDepartmentTitle() As String
Dim slCourseTitle() As String

Private Function Refresh_College()
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    Dim i As Integer
    
    'clear tree
    tvDirectory.Nodes.Clear
    
    sSQL = "SELECT tblCollege.CollegeName" & _
            " FROM tblCollege;"
    
    'vRS.Open sSQL, con, 2, 3
    
    If ConnectRS(con, vRS, sSQL) <> True Then
        GoTo RealeaseAndExit
    End If
    
    If AnyRecordExisted(vRS) = False Then
        GoTo RealeaseAndExit
    End If
    
    ReDim slCollegeTitle(getRecordCount(vRS) - 1)
    i = 0
    vRS.MoveFirst
    
    While vRS.EOF = False
        slCollegeTitle(i) = vRS.Fields("CollegeName")
        AddSchoolYearToTree slCollegeTitle(i)
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
    
     sSQL = "SELECT tblCollege.CollegeID, tblDepartment.DepartmentID, tblCollege.CollegeName, tblDepartment.DepartmentTitle " & _
    "FROM tblCollege INNER JOIN tblDepartment ON tblCollege.CollegeID = tblDepartment.CollegeID "
    
    
    'vRS.Open sSQL, con, 2, 3
    
    If ConnectRS(con, vRS, sSQL) <> True Then
        GoTo RealeaseAndExit
    End If
    
    If AnyRecordExisted(vRS) = False Then
        GoTo RealeaseAndExit
    End If
        
    ReDim slDepartmentTitle(getRecordCount(vRS) - 1)
    
    i = 0
    vRS.MoveFirst
    
    While vRS.EOF = False
        slDepartmentTitle(i) = (vRS.Fields("DepartmentTitle"))
        For ii = 0 To UBound(slCollegeTitle)
            AddDepartmentToTree vRS.Fields("CollegeName"), slDepartmentTitle(i)
        Next
        i = i + 1
        vRS.MoveNext
    Wend

RealeaseAndExit:
    Set vRS = Nothing
End Function
Private Function Refresh_Course()

    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    Dim i As Integer
    Dim ii As Integer
    Dim iii As Integer

    sSQL = "SELECT tblCollege.CollegeName, tblDepartment.DepartmentTitle, tblCourse.Course" & _
            " FROM (tblCollege INNER JOIN tblDepartment ON tblCollege.CollegeID = tblDepartment.CollegeID) INNER JOIN tblCourse ON (tblDepartment.DepartmentID = tblCourse.DepartmentID) AND (tblCollege.CollegeID = tblCourse.CollegeID);"
    
    'vRS.Open sSQL, con, 2, 3
    
    If ConnectRS(con, vRS, sSQL) <> True Then
        GoTo RealeaseAndExit
    End If
    
    If AnyRecordExisted(vRS) = False Then
        GoTo RealeaseAndExit
    End If
        
    ReDim slCourseTitle(getRecordCount(vRS))
    
    i = 0
    vRS.MoveFirst
    
    While vRS.EOF = False
        
        slCourseTitle(i) = (vRS.Fields("Course"))
        
        For ii = 0 To UBound(slCollegeTitle)
            For iii = 0 To UBound(slDepartmentTitle)
                AddYearLevelToTree vRS.Fields("CollegeName"), vRS.Fields("DepartmentTitle"), slCourseTitle(i)
            Next
        Next
        i = i + 1
        vRS.MoveNext
    Wend
RealeaseAndExit:
    Set vRS = Nothing
End Function

Private Function AddSchoolYearToTree(sSchoolYearTitle As String)
    Dim tNode As Node
    
    For Each tNode In tvDirectory.Nodes
        If tNode.Key = keyCollege & ";" & sSchoolYearTitle Then
            Exit Function
        End If
    Next
    
    tvDirectory.Nodes.Add , , keyCollege & ";" & sSchoolYearTitle, sSchoolYearTitle, 1
End Function
Private Function AddDepartmentToTree(sSchoolYearTitle As String, sDepartmentTitle As String)
    Dim tNode As Node
    
    For Each tNode In tvDirectory.Nodes
        If tNode.Key = keyDepartment & ";" & sSchoolYearTitle & ";" & sDepartmentTitle Then
            Exit Function
        End If
    Next
    
    tvDirectory.Nodes.Add keyCollege & ";" & sSchoolYearTitle, tvwChild, keyDepartment & ";" & sSchoolYearTitle & ";" & sDepartmentTitle, sDepartmentTitle, 1
End Function
Private Function AddYearLevelToTree(sSchoolYearTitle As String, sDepartmentTitle As String, sYearLevelTitle As String)
    Dim tNode As Node
    
    For Each tNode In tvDirectory.Nodes
        If tNode.Key = KeyCourse & ";" & sSchoolYearTitle & ";" & sDepartmentTitle & ";" & sYearLevelTitle Then
            Exit Function
        End If
    Next
    
    tvDirectory.Nodes.Add keyDepartment & ";" & sSchoolYearTitle & ";" & sDepartmentTitle, tvwChild, KeyCourse & ";" & sSchoolYearTitle & ";" & sDepartmentTitle & ";" & sYearLevelTitle, sYearLevelTitle, 1

End Function
Private Function Refresh_Tree()
    Dim tv As Nodes
    Dim i As Integer
    
    Refresh_College

    Refresh_Department

    Refresh_Course
    
        For i = 1 To tvDirectory.Nodes.count
            tvDirectory.Nodes(i).Expanded = True
        Next i
End Function
Public Sub FolderClick(fNode As Node, sRecordType As String)
    Dim splitKey() As String
    Dim sKey() As String
    Dim sText() As String
   
    splitKey = Split(fNode.Key, ";")
    Select Case splitKey(0)
        
        Case keyCollege
            ShowCollege splitKey(1)
        Case keyDepartment
            ShowDepartment splitKey(2)
        Case KeyCourse
            ShowCourse splitKey(3)
    End Select
End Sub
Public Sub FolderDisable(fNode As Node, sRecordType As String)
    Dim splitKey() As String
    Dim sKey() As String
    Dim sText() As String
   
    splitKey = Split(fNode.Key, ";")
    Select Case splitKey(0)
        
        Case keyCollege
            EnableCollege
            DisableDepartment
            DisableCourse
        Case keyDepartment
            EnableDepartment
            DisableCourse
            DisableCollege
        Case KeyCourse
            EnableCourse
            DisableDepartment
            DisableCollege
    End Select
End Sub
Public Sub FolderDelete(fNode As Node, sRecordType As String)
    Dim splitKey() As String
    Dim sKey() As String
    Dim sText() As String
   
    splitKey = Split(fNode.Key, ";")
    Select Case splitKey(0)
        Case keyCollege
            Delete_College
        Case keyDepartment
            Delete_Department
        Case KeyCourse
            Delete_Course
    End Select
End Sub
Private Sub DisableCollege()
    mnuNewCollege.Enabled = False
    mnuEditCollege.Enabled = False
    mnuDeleteCollege.Enabled = False
End Sub
Private Sub DisableDepartment()
    mnuEditDepartment.Enabled = False
    mnuDeleteDepartment.Enabled = False
End Sub
Private Sub DisableCourse()
    mnuEditCourse.Enabled = False
    mnuDeleteCourse.Enabled = False
End Sub

Private Sub EnableCollege()
    mnuNewDepartment.Enabled = True
    mnuNewCollege.Enabled = True
    mnuEditCollege.Enabled = True
    mnuDeleteCollege.Enabled = True
End Sub
Private Sub EnableDepartment()
    mnuNewCourse.Enabled = True
    mnuNewDepartment.Enabled = True
    mnuEditDepartment.Enabled = True
    mnuDeleteDepartment.Enabled = True
End Sub
Private Sub EnableCourse()
    mnuNewCourse.Enabled = True
    mnuEditCourse.Enabled = True
    mnuDeleteCourse.Enabled = True
End Sub

Public Sub ShowFormList(Optional iMaxEntryCount As Long = 21, Optional iCurRecPos As Long = 0)
    Dim sSQL As String
    Dim sDepartmentSQL As String
    Dim SubjSQL As String
    

    'MaxEntryCount = iMaxEntryCount
    'CurRecPos = iCurRecPos

    sDefaultSQL = "SELECT RoomID AS lvKey,Building,Room,Capacity,RoomID" & _
                    " From tblRoom"

    sSQL = "SELECT tblTeacher.TeacherID AS lvKey,[TblTeacher]![LastName]+', '+[tblTeacher]![FirstName]+' '+[tblTeacher]![MiddleName] AS FullName, tblTeacher.TeacherID, tblTeacher.Gender " & _
                    " FROM tblTeacher;"
    
    sDepartmentSQL = "SELECT tblDepartment.DepartmentID as lvKey, tblDepartment.DepartmentTitle,DepartmentID From tblDepartment"
    
    
    Refresh_Tree
    
    
     If ConnectRS(con, roomRs, sDefaultSQL & " ORDER BY tblRoom.Room") = True Then
        FillRoomList roomRs
     End If
     
     If ConnectRS(con, DeptRs, sDepartmentSQL) = True Then
        FillDepartmentList DeptRs
     End If
    
    If ConnectRS(con, vRS, sSQL) = True Then
        FillList vRS
        
        Me.Show 1
    Else
        MsgBox "Unable to show Teacher list.", vbCritical
        Unload Me
    End If

End Sub
Private Function FillList(ByRef vRS As ADODB.Recordset)
        
        mdiController.MousePointer = vbHourglass

        UnSortLV lsvFaculty
        
        FillRecordToList vRS, lsvFaculty, KeyStudent, , , , True
        
        SortLV lsvFaculty, lsvFaculty.SortKey, lsvFaculty.SortOrder, False
        
        mdiController.MousePointer = vbDefault
End Function

Private Sub Form_Resize()
SSTab1.Height = ScaleHeight
SSTab1.Width = ScaleWidth

lsvFaculty.Width = SSTab1.Width
lsvFaculty.Height = SSTab1.Height - 400

lsvClassroom.Height = SSTab1.Height - 400
lsvDepartmentRoom.Height = SSTab1.Height - 400
lsvDepartmentRoom.Width = SSTab1.Width - lsvClassroom.Width
lsvDepartmentRoom.Left = lsvClassroom.Width

lsvSubject.Height = SSTab1.Height - 400
lsvSubject.Width = SSTab1.Width - lsvDepartment.Width
lsvDepartment.Height = SSTab1.Height - 400

tvDirectory.Height = SSTab1.Height - 400

End Sub

Private Sub lsvClassroom_Click()
Dim lv As ListItem
Dim rs As New ADODB.Recordset
Dim mySQL As String

    mySQL = "SELECT tblRoom.RoomID, tblDepartment.DepartmentTitle " & _
            "FROM tblDepartment INNER JOIN (tblRoom INNER JOIN tblRoomDepartment ON tblRoom.RoomID = tblRoomDepartment.RoomID) ON tblDepartment.DepartmentID = tblRoomDepartment.DepartmentID " & _
            "WHERE tblRoom.RoomID='" & GetLVKey(lsvClassroom.SelectedItem) & "'"

If ConnectRS(con, rs, mySQL) = True Then
        UnSortLV lsvDepartmentRoom
        lsvDepartmentRoom.ListItems.Clear
        
        Do Until rs.EOF
            Set lv = lsvDepartmentRoom.ListItems.Add(, , rs.Fields("DepartmentTitle"), 1, 1)
            rs.MoveNext
        Loop
        SortLV lsvDepartmentRoom, lsvDepartmentRoom.SortKey, lsvDepartmentRoom.SortOrder, False
End If
Set rs = Nothing
End Sub

Private Sub lsvClassroom_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    SortLV lsvClassroom, ColumnHeader.Index - 1
End Sub

Private Sub lsvClassroom_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        PopupMenu popClassroom
    End If
End Sub

Private Sub lsvDepartment_Click()
    txtDepartmentID.Text = lsvDepartment.SelectedItem.SubItems(1)
    LoadSubjectByDepartment txtDepartmentID.Text
End Sub

Private Sub lsvDepartment_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    SortLV lsvDepartment, ColumnHeader.Index - 1
End Sub

Private Sub lsvDepartment_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        PopupMenu popDepartment
    End If
End Sub
Private Sub lsvDepartmentList_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        PopupMenu Me.popDepartmentAE
    End If
End Sub


Private Sub lsvDepartmentRoom_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    SortLV lsvDepartmentRoom, ColumnHeader.Index - 1
End Sub

Private Sub lsvFaculty_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        PopupMenu popFaculty
    End If
End Sub
Private Sub lsvSubject_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    SortLV lsvSubject, ColumnHeader.Index - 1
End Sub

Private Sub lsvSubject_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        PopupMenu Me.popSubject
    End If
End Sub

Private Sub mnuActivate_Click()
    TeacherOnService (lsvFaculty.SelectedItem.SubItems(1))
End Sub
Public Sub TeacherOnService(sTeacherID As String)
    Dim vRS As New ADODB.Recordset
    
    If ConnectRS(con, vRS, "SELECT * From tblTeacher WHERE (((tblTeacher.TeacherID)='" & sTeacherID & "'));") Then
            vRS.MoveFirst
            vRS.Fields("Status").Value = "1"
            vRS.Update
            MsgBox "Teacher is now on service", vbInformation
    End If
    Set vRS = Nothing
End Sub

Private Sub mnuDelete_Click()
    FormSubject_Delete
End Sub

Private Sub mnuDeleteClassroom_Click()
    FormRoom_Delete
End Sub

Private Sub mnuDeleteCollege_Click()
On Error GoTo err
    Call FolderDelete(tvDirectory.SelectedItem, Left(tvDirectory.SelectedItem.Key, 4))
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub mnuDeleteCourse_Click()
On Error GoTo err
    Call FolderDelete(tvDirectory.SelectedItem, Left(tvDirectory.SelectedItem.Key, 4))
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub mnuDeleteDepartment_Click()
On Error GoTo err
    Call FolderDelete(tvDirectory.SelectedItem, Left(tvDirectory.SelectedItem.Key, 4))
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub mnuDeleteFaculty_Click()
    Form_TeacherDelete
End Sub

Private Sub mnuDepartmentAERefresh_Click()
    Refresh_Tree
End Sub

Private Sub mnuDepartmentRefresh_Click()
    FormDepartment_Refresh
End Sub

Private Sub mnuEditClassroom_Click()
    If lsvClassroom.ListItems.count < 1 Then Exit Sub
    If Len(GetLVKey(lsvClassroom.SelectedItem)) < 1 Then Exit Sub
    
    If frmRoomAE.ShowEdit(lsvClassroom.SelectedItem.SubItems(3)) = True Then
        FormRoom_Refresh
    End If
End Sub

Private Sub mnuEditCollege_Click()
    Call FolderClick(tvDirectory.SelectedItem, Left(tvDirectory.SelectedItem.Key, 4))
End Sub

Private Sub mnuEditCourse_Click()
    Call FolderClick(tvDirectory.SelectedItem, Left(tvDirectory.SelectedItem.Key, 4))
End Sub

Private Sub mnuEditDepartment_Click()
    Call FolderClick(tvDirectory.SelectedItem, Left(tvDirectory.SelectedItem.Key, 4))
End Sub

Private Sub mnuEditFaculty_Click()
    frmFaculty.ShowEdit (lsvFaculty.SelectedItem.SubItems(1))
End Sub

Private Sub mnuEditSubject_Click()
    frmSubjectAE.ShowEdit GetLVKey(lsvSubject.SelectedItem)
End Sub

Private Sub mnuNewClassroom_Click()
    frmRoomAE.ShowForm
End Sub
Private Sub mnuNewCollege_Click()
    frmCollege.ShowForm
End Sub

Private Sub mnuNewCourse_Click()
    frmCourse.ShowForm
End Sub

Private Sub mnuNewDepartment_Click()
    frmDepartmentAE.ShowForm
End Sub

Private Sub mnuNewFaculty_Click()
    frmFaculty.ShowForm
End Sub
Public Sub Form_TeacherDelete()
    Dim i As Integer
    Dim iSelectedCount As Integer
    Dim sMSG As String
    Dim sLVItemKey As String
    Dim EntryDeletedCount As Integer
    
    EntryDeletedCount = 0
    
    iSelectedCount = GetLVSelectedCount(lsvFaculty)
    
    'check if there is a record in the list
    If iSelectedCount < 1 Then
        MsgBox "No selected entry to delete." & _
            vbNewLine & "Please select it first in the list.", vbExclamation
        
        Exit Sub
    End If
    
    If iSelectedCount = 1 Then
        sMSG = "WARNING:" & vbNewLine & "You are about to delete 1 Teacher Entry and you cannot Undo this operation." & vbNewLine & _
        " Are you sure to delete it?"
    Else
        sMSG = "WARNING:" & vbNewLine & "You are about to delete " & iSelectedCount & " Teacher Entries and you cannot Undo this operation." & vbNewLine & _
        " Are you sure to delete it?"
    End If
    
    If MsgBox(sMSG, vbQuestion + vbYesNo) = vbYes Then
        
        For i = 1 To lsvFaculty.ListItems.count
            
            'get key
            sLVItemKey = GetLVKey(lsvFaculty.ListItems(i))
            If lsvFaculty.ListItems(i).Selected = True And Len(sLVItemKey) > 0 Then
                   
                If DeleteTeacher(sLVItemKey) = Success Then
                    EntryDeletedCount = EntryDeletedCount + 1
                Else
                    MsgBox "Unable to delete Teacher Entry with ID: " & sLVItemKey
                End If
            End If
        Next
        If EntryDeletedCount > 0 Then
            MsgBox EntryDeletedCount & " Entry/s deleted.", vbInformation
            Me.Form_FacultyRefresh
        End If
    
    End If
End Sub
Public Sub FormRoom_Delete()
    Dim i As Integer
    Dim iSelectedCount As Integer
    Dim sMSG As String
    Dim sLVItemKey As String
    Dim EntryDeletedCount As Integer
    
    Dim lEnrolmentCount As Long

    EntryDeletedCount = 0
    iSelectedCount = GetLVSelectedCount(lsvClassroom)
    
    If iSelectedCount < 1 Then
        MsgBox "No selected entry to delete." & _
            vbNewLine & "Please select it first in the list.", vbExclamation
        
        Exit Sub
    End If
    
    If iSelectedCount = 1 Then
        sMSG = "WARNING:" & vbNewLine & "You are about to delete 1 Room Entry and you cannot Undo this operation." & vbNewLine & _
        " Are you sure to delete it?"
    Else
        sMSG = "WARNING:" & vbNewLine & "You are about to delete " & iSelectedCount & " Room Entries and you cannot Undo this operation." & vbNewLine & _
        " Are you sure to delete it?"
    End If
    
    If MsgBox(sMSG, vbQuestion + vbYesNo) = vbYes Then
        
        For i = 1 To lsvClassroom.ListItems.count
            sLVItemKey = GetLVKey(lsvClassroom.ListItems(i))
            If lsvClassroom.ListItems(i).Selected = True And Len(sLVItemKey) > 0 Then
                
                        If DeleteRoom(sLVItemKey) = Success Then
                            EntryDeletedCount = EntryDeletedCount + 1
                        Else
                            MsgBox "Unable to delete Room Entry with ID: " & sLVItemKey
                        End If
            End If
        Next
        If EntryDeletedCount > 0 Then
            MsgBox EntryDeletedCount & " Entry/s deleted.", vbInformation

            FormRoom_Refresh
        End If
    
    End If
End Sub
Public Sub FormSubject_Delete()
    Dim i As Integer
    Dim iSelectedCount As Integer
    Dim sMSG As String
    Dim sLVItemKey As String
    Dim EntryDeletedCount As Integer
    
    Dim lEnrolmentCount As Long

    EntryDeletedCount = 0

    iSelectedCount = GetLVSelectedCount(lsvSubject)

    If iSelectedCount < 1 Then
        MsgBox "No selected entry to delete." & _
            vbNewLine & "Please select it first in the list.", vbExclamation
        
        Exit Sub
    End If
    
    If iSelectedCount = 1 Then
        sMSG = "WARNING:" & vbNewLine & "You are about to delete 1 Subject Entry and you cannot Undo this operation." & vbNewLine & _
        " Are you sure to delete it?"
    Else
        sMSG = "WARNING:" & vbNewLine & "You are about to delete " & iSelectedCount & " Subject Entries and you cannot Undo this operation." & vbNewLine & _
        " Are you sure to delete it?"
    End If
    
    If MsgBox(sMSG, vbQuestion + vbYesNo) = vbYes Then
        
        For i = 1 To lsvSubject.ListItems.count
            sLVItemKey = GetLVKey(lsvSubject.ListItems(i))
            If lsvSubject.ListItems(i).Selected = True And Len(sLVItemKey) > 0 Then

                        If DeleteSubject(sLVItemKey) = Success Then
                            EntryDeletedCount = EntryDeletedCount + 1
                        Else
                            MsgBox "Unable to delete Subject Entry with ID: " & sLVItemKey
                        End If
            End If
        Next

        If EntryDeletedCount > 0 Then
            MsgBox EntryDeletedCount & " Entry/s deleted.", vbInformation
            Me.FormSubject_Refresh
        End If
    
    End If
End Sub

Public Sub Form_FacultyRefresh()
    vRS.Requery
    FillList vRS
End Sub


Private Sub mnuNewSubject_Click()
    frmSubjectAE.ShowForm
End Sub

Private Sub mnuprintsubjects_Click()
    '''
End Sub

Private Sub mnuRefresh_Click()
    Form_FacultyRefresh
End Sub

Public Sub FormRoom_Refresh()
    roomRs.Requery
    FillRoomList roomRs
End Sub

Public Sub FormDepartment_Refresh()
    DeptRs.Requery
    FillDepartmentList DeptRs
End Sub

Public Sub FormSubject_Refresh()
On Error Resume Next
    LoadSubjectByDepartment txtDepartmentID.Text
End Sub
Private Function FillRoomList(ByRef vRS As ADODB.Recordset)
        
        mdiController.MousePointer = vbHourglass

        
        UnSortLV lsvClassroom
        
        FillRecordToList vRS, lsvClassroom, KeyStudent, , , , True
        
        SortLV lsvClassroom, lsvClassroom.SortKey, lsvClassroom.SortOrder, False
        
        
        mdiController.MousePointer = vbDefault
End Function

Private Function FillDepartmentList(ByRef vRS As ADODB.Recordset)
        
        mdiController.MousePointer = vbHourglass
                
        UnSortLV lsvDepartment
        FillRecordToList vRS, lsvDepartment, KeyStudent, , , , True
        SortLV lsvDepartment, lsvDepartment.SortKey, lsvDepartment.SortOrder, False

        mdiController.MousePointer = vbDefault
End Function

Private Function FillSubjectList(ByRef vRS As ADODB.Recordset)
        mdiController.MousePointer = vbHourglass
        UnSortLV lsvSubject
        FillRecordToList vRS, lsvSubject, KeyStudent, , , , True
        SortLV lsvSubject, lsvSubject.SortKey, lsvSubject.SortOrder, False
        mdiController.MousePointer = vbDefault
End Function

Private Sub mnuRoomRefresh_Click()
    FormRoom_Refresh
End Sub

Private Sub mnuSubjectRefresh_Click()
    FormSubject_Refresh
End Sub
Private Function DeleteRoom(sRoomID As String) As Boolean

End Function
Private Sub DeleteDepartment(sDepartmentTitle As String)

    Dim vDepartment As tDepartment
    If GetDepartmentByTitle(sDepartmentTitle, vDepartment) = Success Then
        If modDBDepartment.DeleteDepartment(vDepartment.DepartmentID) = Success Then
            MsgBox "Department record successfully deleted.", vbInformation
            FormDepartment_Refresh
        Else
            MsgBox "Unable to delete Department Record!", vbExclamation
        End If
    End If
        
End Sub

Private Sub tvDirectory_Click()
On Error GoTo err
    Call FolderDisable(tvDirectory.SelectedItem, Left(tvDirectory.SelectedItem.Key, 4))
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub tvDirectory_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        PopupMenu Me.popDepartmentAE
    End If
End Sub

Private Sub LoadSubjectByDepartment(sSubjectID As String)
Dim vRS As New ADODB.Recordset
Dim sSQL  As String

     sSQL = "SELECT tblSubject.SubjectID AS lvKey, tblSubject.SubjectTitle AS Title, tblSubject.Units, tblSubject.Description FROM tblDepartment INNER JOIN tblSubject ON tblDepartment.DepartmentID = tblSubject.DepartmentID " & _
            "WHERE tblDepartment.DepartmentID ='" & sSubjectID & "'"
            
If ConnectRS(con, vRS, sSQL) = True Then
    FillSubjectList vRS
End If

End Sub
Public Function SelectNode(sKey As String) As Variant

    Dim tNode As Node
    
    If IsStarted = False Then
        Refresh_Tree
    End If
    
    For Each tNode In tvFolder.Nodes
        If tNode.Key = sKey Then

            tNode.Selected = True
            'tvDirectory_Click
        End If
    Next

End Function
Public Function GetSchoolYearChilds(sSchoolYearTitle As String, ByRef sKey() As String, ByRef sText() As String) As Boolean
    Dim tNode As Node
    Dim NodeCount As Integer
    Dim i As Integer
    Dim splitKey() As String
    
    NodeCount = 0
    
    For Each tNode In tvFolder.Nodes
        splitKey = Split(tNode.Key, ";")
        If splitKey(0) = keyDepartment And splitKey(1) = sSchoolYearTitle Then
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
    For Each tNode In tvFolder.Nodes
        splitKey = Split(tNode.Key, ";")
        If splitKey(0) = keyDepartment And splitKey(1) = sSchoolYearTitle Then
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
    
    For Each tNode In tvFolder.Nodes
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
    For Each tNode In tvFolder.Nodes
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
    
    For Each tNode In tvFolder.Nodes
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
    For Each tNode In tvFolder.Nodes
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
Private Function ShowCourse(sCourse As String)
    frmCourse.ShowEdit sCourse
End Function
Public Sub ShowCollege(sCollegeTitle As String)
    curCollegeTitle = sCollegeTitle
    frmCollege.ShowProperties (curCollegeTitle)
End Sub
Private Function ShowDepartment(sDepartmentTitle As String)
    frmDepartmentAE.ShowProperties sDepartmentTitle
End Function
Private Function SetSelectedSection(sSectionTitle As String, Optional sSchoolYearTitle As String = "")
    Dim tNode As Node
    Dim splitKey() As String
    For Each tNode In tvFolder.Nodes
    
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
Public Function GetSchoolYearList(ByRef sList() As String) As Boolean
    If UBound(slDepartmentTitle) < 0 Then
        GetSchoolYearList = False
    Else
        sList = slDepartmentTitle
        GetSchoolYearList = True
    End If
End Function
Public Function GetYearLevelList(ByRef sList() As String) As Variant
    sList = slCourseTitle
End Function

Public Sub Delete_College()
    Dim vCollege As tCollege
    Dim sLVItemKey As String
    Dim sMSG As String
    
     sMSG = "WARNING:" & vbNewLine & _
        "Deleting this COLLEGE entry will affect all other record" & vbNewLine & vbNewLine & _
        "Delete this record anyway?"
        
    If MsgBox(sMSG, vbQuestion + vbYesNo) = vbYes Then
        
            sLVItemKey = tvDirectory.SelectedItem
                        
                        If GetCollegeByTitle(sLVItemKey, vCollege) <> Success Then
                            Exit Sub
                        End If
                        
                        If DeleteCollege(vCollege.CollegeID) = Success Then
                            MsgBox "COLLEGE entry and other related record succesfully deleted.", vbInformation
                            Refresh_Tree
                        Else
                            MsgBox "Deleting COLLEGE entry went failed.", vbExclamation
                        End If
    End If
        
End Sub
Public Sub Delete_Department()
    Dim vDepartment As tDepartment
    Dim sLVItemKey As String
    Dim sMSG As String
    
     sMSG = "WARNING:" & vbNewLine & _
        "Deleting this DEPARTMENT entry will affect all other record" & vbNewLine & vbNewLine & _
        "Delete this record anyway?"
        
    If MsgBox(sMSG, vbQuestion + vbYesNo) = vbYes Then
        
            sLVItemKey = tvDirectory.SelectedItem
                        
                        If GetDepartmentByTitle(sLVItemKey, vDepartment) <> Success Then
                            Exit Sub
                        End If
                        
                        If modDBDepartment.DeleteDepartment(vDepartment.DepartmentID) = Success Then
                            MsgBox "DEPARTMENT entry and other related record succesfully deleted.", vbInformation
                           Refresh_Tree
                        Else
                           MsgBox "Deleting DEPARTMENT entry went failed.", vbExclamation
                        End If
    End If
End Sub

Public Sub Delete_Course()
    Dim vCourse As tCourse
    Dim sLVItemKey As String
    Dim sMSG As String
    
     sMSG = "WARNING:" & vbNewLine & _
        "Deleting this COURSE entry will affect all other record" & vbNewLine & vbNewLine & _
        "Delete this record anyway?"
        
        If MsgBox(sMSG, vbQuestion + vbYesNo) = vbYes Then
        
            sLVItemKey = tvDirectory.SelectedItem
                        
                        If GetCourseByTitle(sLVItemKey, vCourse) <> Success Then
                            Exit Sub
                        End If
                        
                        If DeleteCourse(vCourse.CourseID) = Success Then
                            MsgBox "COURSE entry and other related record succesfully deleted.", vbInformation
                           Refresh_Tree
                        Else
                            MsgBox "Deleting COURSE entry went failed.", vbExclamation
                        End If
    End If
End Sub
