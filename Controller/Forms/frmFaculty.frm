VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmFaculty 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Faculty Properties"
   ClientHeight    =   8205
   ClientLeft      =   150
   ClientTop       =   510
   ClientWidth     =   8865
   Icon            =   "frmFaculty.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8205
   ScaleWidth      =   8865
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   8175
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   14420
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      TabCaption(0)   =   "Personal Info"
      TabPicture(0)   =   "frmFaculty.frx":492A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Advisee"
      TabPicture(1)   =   "frmFaculty.frx":4946
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "StatusBar"
      Tab(1).Control(1)=   "lsvAdvisee"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Faculty Load"
      TabPicture(2)   =   "frmFaculty.frx":4962
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "ImageList1"
      Tab(2).Control(1)=   "Toolbar1"
      Tab(2).Control(2)=   "lsvLoad"
      Tab(2).Control(3)=   "StatusBar1"
      Tab(2).Control(4)=   "lsvRoomSched"
      Tab(2).Control(5)=   "lsvClasslist"
      Tab(2).Control(6)=   "imgSubject"
      Tab(2).Control(7)=   "ilRecordIco"
      Tab(2).Control(8)=   "icoHeader"
      Tab(2).ControlCount=   9
      Begin VB.Frame Frame1 
         Height          =   7695
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   8655
         Begin VB.TextBox txtTeacherID 
            Height          =   345
            Left            =   1440
            TabIndex        =   35
            Top             =   240
            Width           =   1965
         End
         Begin VB.TextBox txtUserName 
            Height          =   345
            Left            =   120
            TabIndex        =   32
            Top             =   3720
            Width           =   3645
         End
         Begin VB.TextBox txtPassword 
            Height          =   345
            IMEMode         =   3  'DISABLE
            Left            =   120
            PasswordChar    =   "*"
            TabIndex        =   31
            Top             =   4440
            Width           =   3645
         End
         Begin VB.PictureBox picImageSetter 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Enabled         =   0   'False
            ForeColor       =   &H80000008&
            Height          =   1215
            Left            =   4680
            ScaleHeight     =   1185
            ScaleWidth      =   3825
            TabIndex        =   27
            Top             =   3960
            Visible         =   0   'False
            Width           =   3855
            Begin VB.TextBox txtPic 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   120
               TabIndex        =   28
               Text            =   "txtPic"
               Top             =   120
               Width           =   3615
            End
            Begin RichTextLib.RichTextBox rtbPic 
               Height          =   255
               Left            =   120
               TabIndex        =   29
               Top             =   840
               Width           =   3015
               _ExtentX        =   5318
               _ExtentY        =   450
               _Version        =   393217
               Appearance      =   0
               TextRTF         =   $"frmFaculty.frx":497E
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
            Begin MSComDlg.CommonDialog cdbPic 
               Left            =   3240
               Top             =   600
               _ExtentX        =   847
               _ExtentY        =   847
               _Version        =   393216
            End
            Begin VB.Label lblPic 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "lblPic"
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   120
               TabIndex        =   30
               Top             =   480
               Width           =   3015
            End
         End
         Begin VB.TextBox txtObjects 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            DataField       =   "Image"
            ForeColor       =   &H00000080&
            Height          =   285
            Index           =   11
            Left            =   6360
            TabIndex        =   26
            Top             =   3240
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.TextBox txtLName 
            Height          =   345
            Left            =   120
            TabIndex        =   18
            Top             =   1080
            Width           =   1965
         End
         Begin VB.TextBox txtFName 
            Height          =   345
            Left            =   2160
            TabIndex        =   17
            Top             =   1080
            Width           =   1965
         End
         Begin VB.TextBox txtMName 
            Height          =   345
            Left            =   4200
            TabIndex        =   16
            Top             =   1080
            Width           =   1965
         End
         Begin VB.ComboBox cboGender 
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
            ItemData        =   "frmFaculty.frx":49FC
            Left            =   120
            List            =   "frmFaculty.frx":4A06
            Style           =   2  'Dropdown List
            TabIndex        =   15
            Top             =   2040
            Width           =   1575
         End
         Begin VB.CheckBox chkService 
            Caption         =   "Active service"
            Height          =   375
            Left            =   1800
            TabIndex        =   14
            Top             =   2040
            Width           =   1575
         End
         Begin VB.CommandButton cmdBrowse 
            Caption         =   "Browse"
            Height          =   375
            Left            =   6240
            TabIndex        =   13
            Top             =   2640
            Width           =   855
         End
         Begin VB.CommandButton cmdClear 
            Caption         =   "Clear"
            Height          =   375
            Left            =   7680
            TabIndex        =   12
            Top             =   2640
            Width           =   855
         End
         Begin VB.CommandButton cmdSave 
            Caption         =   "Save"
            Height          =   375
            Left            =   6600
            TabIndex        =   11
            Top             =   6960
            Width           =   1575
         End
         Begin VB.CommandButton cmdGetItem 
            Appearance      =   0  'Flat
            BackColor       =   &H00D8E9EC&
            Height          =   345
            Left            =   3840
            Picture         =   "frmFaculty.frx":4A18
            Style           =   1  'Graphical
            TabIndex        =   10
            Top             =   2880
            Width           =   345
         End
         Begin VB.TextBox txtDepartmentTitle 
            Height          =   345
            Left            =   120
            MaxLength       =   50
            TabIndex        =   9
            Top             =   2880
            Width           =   3705
         End
         Begin VB.Label lblLabels 
            Caption         =   "Teacher Ref. No:"
            Height          =   270
            Index           =   7
            Left            =   120
            TabIndex        =   36
            Top             =   240
            Width           =   1560
         End
         Begin VB.Label lblLabels 
            BackStyle       =   0  'Transparent
            Caption         =   "&User Name:"
            Height          =   270
            Index           =   6
            Left            =   120
            TabIndex        =   34
            Top             =   3480
            Width           =   1080
         End
         Begin VB.Label lblLabels 
            BackStyle       =   0  'Transparent
            Caption         =   "&Password:"
            Height          =   270
            Index           =   5
            Left            =   120
            TabIndex        =   33
            Top             =   4200
            Width           =   1080
         End
         Begin VB.Label lblLabels 
            Caption         =   "Last name:"
            Height          =   270
            Index           =   0
            Left            =   120
            TabIndex        =   23
            Top             =   840
            Width           =   1080
         End
         Begin VB.Label lblLabels 
            Caption         =   "First name:"
            Height          =   270
            Index           =   1
            Left            =   2160
            TabIndex        =   22
            Top             =   840
            Width           =   1080
         End
         Begin VB.Label lblLabels 
            Caption         =   "Middle name:"
            Height          =   270
            Index           =   2
            Left            =   4200
            TabIndex        =   21
            Top             =   840
            Width           =   1080
         End
         Begin VB.Image imgMain 
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Height          =   2295
            Left            =   6240
            Stretch         =   -1  'True
            Top             =   240
            Width           =   2295
         End
         Begin VB.Label lblLabels 
            Caption         =   "Gender:"
            Height          =   270
            Index           =   3
            Left            =   120
            TabIndex        =   20
            Top             =   1800
            Width           =   1080
         End
         Begin VB.Label lblLabels 
            Caption         =   "Assigned Department:"
            Height          =   270
            Index           =   4
            Left            =   120
            TabIndex        =   19
            Top             =   2640
            Width           =   1920
         End
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   -68280
         Top             =   1080
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   3
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmFaculty.frx":4FA2
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmFaculty.frx":A794
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmFaculty.frx":B1A6
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   330
         Left            =   -75000
         TabIndex        =   1
         Top             =   360
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         Style           =   1
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   4
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               ImageIndex      =   1
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   4
               Object.Width           =   3500
            EndProperty
         EndProperty
         Begin VB.ComboBox cboSY 
            Height          =   315
            Left            =   600
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   0
            Width           =   1695
         End
         Begin VB.ComboBox cboSemester 
            Height          =   315
            Left            =   2400
            TabIndex        =   2
            Text            =   "cboSemester"
            Top             =   0
            Width           =   1695
         End
      End
      Begin MSComctlLib.ListView lsvLoad 
         Height          =   2775
         Left            =   -74940
         TabIndex        =   4
         Top             =   720
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   4895
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         Icons           =   "imgSubject"
         SmallIcons      =   "imgSubject"
         ColHdrIcons     =   "ilRecordIco"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Subject"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Section"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Units"
            Object.Width           =   1411
         EndProperty
      End
      Begin MSComctlLib.StatusBar StatusBar1 
         Height          =   375
         Left            =   -74940
         TabIndex        =   5
         Top             =   7740
         Width           =   8775
         _ExtentX        =   15478
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
            NumPanels       =   2
            BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               Object.Width           =   7938
               MinWidth        =   7938
            EndProperty
            BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               Object.Width           =   7938
               MinWidth        =   7938
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.StatusBar StatusBar 
         Height          =   375
         Left            =   -74940
         TabIndex        =   6
         Top             =   7750
         Width           =   8775
         _ExtentX        =   15478
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
               Object.Width           =   8819
               MinWidth        =   8819
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ListView lsvAdvisee 
         Height          =   7390
         Left            =   -74940
         TabIndex        =   7
         Top             =   360
         Width           =   8775
         _ExtentX        =   15478
         _ExtentY        =   13044
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
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Student Name"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "ID Number"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Course"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ListView lsvRoomSched 
         Height          =   2775
         Left            =   -70640
         TabIndex        =   24
         Top             =   720
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   4895
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
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Room Assignment"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Schedule"
            Object.Width           =   4057
         EndProperty
      End
      Begin MSComctlLib.ListView lsvClasslist 
         Height          =   4275
         Left            =   -74940
         TabIndex        =   25
         Top             =   3480
         Width           =   8775
         _ExtentX        =   15478
         _ExtentY        =   7541
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
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Student Name"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Final Grade"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Comp Grade"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Submitted By"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Date Submitted"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ImageList imgSubject 
         Left            =   -74880
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
               Picture         =   "frmFaculty.frx":BBB8
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageList ilRecordIco 
         Left            =   -74160
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
               Picture         =   "frmFaculty.frx":C152
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageList icoHeader 
         Left            =   -74520
         Top             =   360
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
               Picture         =   "frmFaculty.frx":C6EC
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmFaculty.frx":CC86
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.Menu mnup 
      Caption         =   "Print"
      Visible         =   0   'False
      Begin VB.Menu mnuprintclass 
         Caption         =   "Print class list..."
      End
   End
End
Attribute VB_Name = "frmFaculty"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim vRS As New ADODB.Recordset
Dim FacultyID As String
Dim mFormState As String
Dim curDepartment As String

Dim CurrentTeacher As tTeacher
Dim RecordAdded As Boolean
Dim RecordEdited As Boolean
Dim sDepartmentID As String

Private Function ValidateData() As Boolean
    ValidateData = False
    If Not CheckTextBox(txtLName, "Teacher Last Name") Then Exit Function
    If Not CheckTextBox(txtFName, "Teacher First Name") Then Exit Function
    If Not CheckTextBox(txtMName, "Teacher Middle Name") Then Exit Function
    If Not CheckTextBox(txtDepartmentTitle, "Assign Department") Then Exit Function
    If Not CheckTextBox(txtUserName, "UserName") Then Exit Function
    If Not CheckTextBox(txtPassword, "Password") Then Exit Function
    ValidateData = True
End Function

Private Sub cboSemester_Change()
   Call RefreshSubjects(CurrentTeacher.TeacherID, cboSY.Text, cboSemester.Text)
End Sub

Private Sub cboSemester_Click()
    cboSemester_Change
End Sub

Private Sub cboSY_Change()
     Call RefreshSubjects(CurrentTeacher.TeacherID, cboSY.Text, cboSemester.Text)
End Sub

Private Sub cboSY_Click()
    cboSY_Change
End Sub
Private Sub saveimage()
Dim vDirectory, vDatabase
vDirectory = App.Path & "\Picture"
If Right(vDirectory, 1) = Chr(92) Then vDirectory = Left(vDirectory, (Len(vDirectory) - 1))
    rtbPic.FileName = cdbPic.FileName
    vDatabase = FreeFile
    Open vDirectory + "\" + txtTeacherID + ".jpg" For Output As vDatabase
    Print #1, rtbPic.Text
    Close vDatabase
End Sub
Private Sub cmdBrowse_Click()
Dim vI As Integer

    On Error Resume Next
    With cdbPic
        .DialogTitle = "Search Employee picture"
        .Filter = "Bitmap (*.bmp)|*.bmp|Jpeg (*.jpg)|*.jpg|Gif (*.gif)|*.gif|All Files (*.*)|*.*"
        .Flags = cdlOFNHideReadOnly + cdlOFNOverwritePrompt + cdlOFNPathMustExist
        .ShowOpen
        .FilterIndex = 1
        .CancelError = True
        imgMain.Picture = LoadPicture(.FileName)
        rtbPic.FileName = .FileName
        lblPic = .FileName
    End With
    For vI = 1 To Len(lblPic)
        If Mid(lblPic, vI, 1) = "\" Then txtPic = Mid(lblPic, vI + 1, Len(lblPic))
    Next vI
    If imgMain.Picture = LoadPicture("") Then
        imgMain.Picture = LoadPicture(App.Path & "\Picture\default.jpg")
        txtObjects(11) = "default.jpg"
        rtbPic.FileName = App.Path & "\Picture\default.jpg"
    Else
        txtObjects(11) = FacultyID
    End If
End Sub

Private Sub cmdGetItem_Click()
    Dim sDepartmentTitle As String

    sDepartmentID = frmPickDepartment.GetItem(txtDepartmentTitle, sDepartmentTitle)
    If sDepartmentID <> "" Then
        txtDepartmentTitle = sDepartmentTitle
        curDepartment = sDepartmentID
    End If
End Sub

Private Sub cmdSave_Click()
    Select Case cmdSave.Caption
        Case "&Save"
            saveimage
            If SaveData Then
                RecordAdded = True
                Unload Me
            End If
        Case "&Update"
            saveimage
            If UpdateData Then
                RecordEdited = True
                Unload Me
            End If
    End Select
End Sub
Private Sub RefreshSYList()
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    
    sSQL = "SELECT tblSchoolYear.SchoolYear" & _
            " FROM tblSchoolYear" & _
            " ORDER BY tblSchoolYear.SchoolYear"
     
    If ConnectRS(con, vRS, sSQL) = True Then
        cboSY.Clear
        Do Until vRS.EOF
            cboSY.AddItem (vRS.Fields("SchoolYear"))
            vRS.MoveNext
        Loop
        cboSY.Text = CurrentSchoolYear.SchoolYearTitle
        cboSemester.Text = CurrentSemester.Semester
    End If
    
ReleaseAndExit:
    Set vRS = Nothing
End Sub
Private Sub RefreshSemesterList()
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    
    sSQL = "SELECT tblSemester.Semester" & _
            " FROM tblSemester" & _
            " ORDER BY tblSemester.Semester"
     
    If ConnectRS(con, vRS, sSQL) = True Then
        cboSemester.Clear
        Do Until vRS.EOF
            cboSemester.AddItem (vRS.Fields("Semester"))
            vRS.MoveNext
        Loop
        cboSemester.Text = CurrentSemester.Semester
    End If
    
ReleaseAndExit:
    Set vRS = Nothing
End Sub

Public Function ShowForm() As Boolean
    Dim sNewID As String
    
     RefreshSYList
    RefreshSemesterList
    
    GetNewTeacherID sNewID
    
    txtTeacherID.Text = sNewID
    
    cmdSave.Caption = "&Save"
    Me.Show vbModal
    
End Function

Public Sub ShowEdit(sTeacherID As String)
 Dim vDepartment As tDepartment
    
     cmdSave.Caption = "&Update"
    
    If GetTeacherByID(sTeacherID, CurrentTeacher) <> Success Then
        MsgBox "Unable to continue proccess." & vbNewLine & "The selected TEACHER entry was not found in record.", vbExclamation
        Unload Me
        Exit Sub
    End If
    
           
    If GetDepartmentByID(CurrentTeacher.Department, vDepartment) = Success Then
            txtDepartmentTitle = vDepartment.DepartmentTitle
    End If
        
    Call SetTextField
    
    Call RefreshSubjects(sTeacherID, CurrentSchoolYear.SchoolYearTitle, CurrentSemester.Semester)
    GenerateAdviseeList sTeacherID
    
On Error GoTo err
            Set imgMain.Picture = LoadPicture(App.Path & "\Picture\" & sTeacherID & ".jpg")
            txtObjects(11) = CurrentTeacher.TeacherID

    Me.Show vbModal
Exit Sub

err:
            imgMain.Picture = LoadPicture(App.Path & "\Picture\default.jpg")
            txtObjects(11) = "default.jpg"
            rtbPic.FileName = App.Path & "\Picture\default.jpg"
            
     Me.Show vbModal
     
End Sub
    
    


Private Function SetTextField()
    txtTeacherID.Text = CurrentTeacher.TeacherID
    txtFName.Text = CurrentTeacher.FirstName
    txtMName.Text = CurrentTeacher.MiddleName
    txtLName.Text = CurrentTeacher.LastName
    cboGender.Text = CurrentTeacher.Gender
    chkService.Value = CurrentTeacher.OnService
    curDepartment = CurrentTeacher.Department
    txtUserName.Text = CurrentTeacher.Username
    txtPassword.Text = CurrentTeacher.Password
End Function
Private Function SaveData() As Boolean
    
    Dim newTeacher As tTeacher
    SaveData = False

    If Not ValidateData Then Exit Function
    newTeacher.TeacherID = Trim(txtTeacherID)
    newTeacher.FirstName = Trim(txtFName)
    newTeacher.MiddleName = Trim(txtMName)
    newTeacher.LastName = Trim(txtLName)
    newTeacher.Gender = Trim(cboGender)
    newTeacher.OnService = Me.chkService.Value
    newTeacher.CreationDate = Now
    newTeacher.Department = curDepartment
    newTeacher.Username = txtUserName.Text
    newTeacher.Password = txtPassword.Text
    
    Select Case AddTeacher(newTeacher)
            
            Case TranDBResult.Success
                MsgBox "TEACHER entry successfull Added.", vbInformation
                'Unload Me

            Case TranDBResult.DuplicateID
                MsgBox "The TEACHER ID you have entered is already existed." & vbNewLine & "Please enter a different value.", vbExclamation
                txtTeacherID.SetFocus
                
            Case TranDBResult.InvalidTeacherFirstName
                MsgBox "Invalid FIRST NAME.", vbExclamation
                txtFName.SetFocus

            Case TranDBResult.InvalidTeacherMiddleName
                MsgBox "Invalid MIDDLE NAME.", vbExclamation
                txtMName.SetFocus

            Case TranDBResult.InvalidTeacherLastName
                MsgBox "Invalid LAST NAME.", vbExclamation
                txtLName.SetFocus
                
            Case Else
                MsgBox "Unknown result: Adding Teacher entry", vbCritical
        End Select
End Function

Private Function UpdateData() As Boolean

    UpdateData = False
        
    If Not ValidateData Then Exit Function
    
    'CurrentTeacher.TeacherID = txtTeacherID.Text
    CurrentTeacher.FirstName = Trim(txtFName)
    CurrentTeacher.MiddleName = Trim(txtMName)
    CurrentTeacher.LastName = Trim(txtLName)
    CurrentTeacher.Gender = cboGender.Text
    CurrentTeacher.OnService = chkService.Value
    CurrentTeacher.Department = curDepartment
    CurrentTeacher.Username = txtUserName.Text
    CurrentTeacher.Password = txtPassword.Text
    
        Select Case EditTeacher(CurrentTeacher)
            
            Case TranDBResult.Success
                MsgBox "TEACHER entry successfull edited.", vbInformation
                UpdateData = True
                'Unload Me
                
            Case TranDBResult.DuplicateID
                MsgBox "The TEACHER ID you have entered is already existed." & vbNewLine & "Please enter a different value.", vbExclamation
                txtTeacherID.SetFocus
                UpdateData = False
            Case TranDBResult.InvalidTeacherFirstName
                MsgBox "Invalid FIRST NAME.", vbExclamation
                txtFName.SetFocus
                UpdateData = False
            Case TranDBResult.InvalidTeacherMiddleName
                MsgBox "Invalid MIDDLE NAME.", vbExclamation
                txtMName.SetFocus
                UpdateData = False
            Case TranDBResult.InvalidTeacherLastName
                MsgBox "Invalid LAST NAME.", vbExclamation
                txtLName.SetFocus
                UpdateData = False
        End Select
End Function

Private Function RefreshSubjects(sTeacherID As String, sSchoolYear As String, sSemester As String)
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    lsvLoad.ListItems.Clear
    
    On Error GoTo ReleaseAndExit
    
   sSQL = "SELECT tblSubjectOffering.SubjectOfferingID, tblSubject.SubjectTitle, tblSection.SectionTitle, tblSubject.Units " & _
            "FROM tblTeacher INNER JOIN (tblSubject INNER JOIN (tblSection INNER JOIN tblSubjectOffering ON tblSection.SectionID = tblSubjectOffering.SectionID) ON tblSubject.SubjectID = tblSubjectOffering.SubjectID) ON tblTeacher.TeacherID = tblSubjectOffering.TeacherID " & _
            "WHERE tblTeacher.TeacherID='" & sTeacherID & "' and tblSection.SchoolYear='" & sSchoolYear & "' and tblSection.Semester ='" & sSemester & "'"
           

    If ConnectRS(con, vRS, sSQL) = False Then
        MsgBox "Unable to conect Teacher Recordset.", vbCritical
        GoTo ReleaseAndExit
    End If
    
    If AnyRecordExisted(vRS) = True Then
        FillRecordToList vRS, lsvLoad, KeySubjectOffering, , 32767, , True
    End If

ReleaseAndExit:
    Set vRS = Nothing
End Function
Private Function GenerateStudentList(sSubjectID As String)
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    Dim lv As ListItem

    sSQL = "SELECT tblSubjectOffering.SubjectOfferingID as lvKey, [tblStudent]![LastName] & ', ' & [tblStudent]![FirstName] & ' ' & [tblStudent]![MiddleName] AS FullName, tblGrade.FinalGrade, tblGrade.CompGrade " & _
            "FROM tblSubjectOffering INNER JOIN (tblStudent INNER JOIN (tblEnrolment INNER JOIN tblGrade ON tblEnrolment.EnrollmentID = tblGrade.EnrolmentID) ON tblStudent.StudentID = tblEnrolment.StudentID) ON (tblSubjectOffering.SubjectOfferingID = tblGrade.SubjectOfferingID) AND (tblSubjectOffering.SubjectOfferingID = tblGrade.SubjectOfferingID) " & _
            "WHERE tblSubjectOffering.SubjectOfferingID='" & sSubjectID & "'" & _
            "GROUP BY tblSubjectOffering.SubjectOfferingID, [tblStudent]![LastName] & ', ' & [tblStudent]![FirstName] & ' ' & [tblStudent]![MiddleName], tblGrade.FinalGrade, tblGrade.CompGrade " & _
            "ORDER BY [tblStudent]![LastName] & ', ' & [tblStudent]![FirstName] & ' ' & [tblStudent]![MiddleName];"

    lsvClasslist.ListItems.Clear

    If ConnectRS(con, vRS, sSQL) = True Then
        If AnyRecordExisted(vRS) = True Then
            Do Until vRS.EOF
            Set lv = lsvClasslist.ListItems.Add(, , vRS.Fields("Fullname"), 1, 1)
                    lv.SubItems(1) = vRS.Fields("FinalGrade")
                    lv.SubItems(2) = vRS.Fields("CompGrade")
                vRS.MoveNext
            Loop
        Else
            'no records
        End If
    Else
        CatchError "frmFaculty", "GenerateClassList", "Error connecting Enrolments"
    End If
    Set vRS = Nothing
End Function

Private Function GenerateAdviseeList(sTeacherID As String)
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String

     sSQL = "SELECT tblEnrolment.EnrollmentID, [tblStudent]![LastName] & ', ' & [tblStudent]![FirstName] & ' ' & [tblStudent]![MiddleName] AS FullName, tblStudent.StudentID, tblCourse.Course & ' major in ' & tblCourse.Major AS Course " & _
            "FROM (tblSubjectOffering INNER JOIN (tblStudent INNER JOIN (tblEnrolment INNER JOIN tblGrade ON tblEnrolment.EnrollmentID = tblGrade.EnrolmentID) ON tblStudent.StudentID = tblEnrolment.StudentID) ON (tblSubjectOffering.SubjectOfferingID = tblGrade.SubjectOfferingID) AND (tblSubjectOffering.SubjectOfferingID = tblGrade.SubjectOfferingID)) INNER JOIN (tblCourse INNER JOIN tblStudentStatus ON tblCourse.CourseID = tblStudentStatus.CourseID) ON tblStudent.StudentID = tblStudentStatus.StudentID " & _
            "WHERE tblSubjectOffering.TeacherID='" & sTeacherID & "'" & _
            "GROUP BY tblEnrolment.EnrollmentID, [tblStudent]![LastName] & ', ' & [tblStudent]![FirstName] & ' ' & [tblStudent]![MiddleName], tblStudent.StudentID, tblCourse.Course & ' major in ' & tblCourse.Major, tblSubjectOffering.TeacherID"

    
    lsvAdvisee.ListItems.Clear

    If ConnectRS(con, vRS, sSQL) = True Then
        If AnyRecordExisted(vRS) = True Then
            FillRecordToList vRS, lsvAdvisee, "class", , , , True
        Else
            'no records
        End If
    Else
        CatchError "frmFaculty", "GenerateAdviseeList", "Error connecting Enrolments"
    End If
    Set vRS = Nothing
End Function
Private Function GenerateRoom(sSubjectID As String)
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    Dim lv As ListItem
     
sSQL = "SELECT tblSubjectOffering.SubjectOfferingID AS lvKey, tblRoom.Building & ' - ' & tblRoom.Room AS Room, tblSubjectOffering.Days &' '& tblSubjectOffering.TimeIn&' - '&tblSubjectOffering.TimeOut as Schedule " & _
        "FROM tblRoom INNER JOIN (tblTeacher INNER JOIN tblSubjectOffering ON tblTeacher.TeacherID = tblSubjectOffering.TeacherID) ON tblRoom.RoomID = tblSubjectOffering.RoomID " & _
        "WHERE tblSubjectOffering.SubjectOfferingID ='" & sSubjectID & "'"
    
    lsvRoomSched.ListItems.Clear
    
    If ConnectRS(con, vRS, sSQL) = True Then
        If AnyRecordExisted(vRS) = True Then
            Do Until vRS.EOF
                Set lv = lsvRoomSched.ListItems.Add(, , vRS.Fields("Room"), 1, 1)
                    lv.SubItems(1) = vRS.Fields("Schedule")
                vRS.MoveNext
            Loop
        Else
            'no records
        End If
    Else
        CatchError "frmFaculty", "GenerateRoomList", "Error connecting Enrolments"
    End If
Set vRS = Nothing
End Function

Private Sub lsvLoad_Click()
    GenerateStudentList (GetLVKey(lsvLoad.SelectedItem))
    GenerateRoom (GetLVKey(lsvLoad.SelectedItem))
End Sub
