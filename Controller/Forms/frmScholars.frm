VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmScholars 
   BackColor       =   &H00D8E9EC&
   Caption         =   "Scholars"
   ClientHeight    =   7755
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   9615
   Icon            =   "frmScholars.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7755
   ScaleWidth      =   9615
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ListView lsvScholarship 
      Height          =   6375
      Left            =   0
      TabIndex        =   4
      Top             =   480
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   11245
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Scholarship"
         Object.Width           =   5953
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Object.Width           =   0
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Top             =   7380
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   6165
            MinWidth        =   6165
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   14111
            MinWidth        =   14111
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00D8E9EC&
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9135
      Begin VB.ComboBox cboSemester 
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
         ItemData        =   "frmScholars.frx":492A
         Left            =   1860
         List            =   "frmScholars.frx":4937
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   120
         Width           =   1815
      End
      Begin VB.ComboBox cboSY 
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
         Left            =   60
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   120
         Width           =   1815
      End
   End
   Begin MSComctlLib.ListView lsvScholar 
      Height          =   6375
      Left            =   3480
      TabIndex        =   5
      Top             =   480
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   11245
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   9
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Name"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "ID Number"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Gender"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Course/Curriculum"
         Object.Width           =   7056
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Load"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "GPA"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "CGPA"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Year Level"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Attended"
         Object.Width           =   3528
      EndProperty
   End
End
Attribute VB_Name = "frmScholars"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    RefreshSYList
    RefreshScholarship
End Sub

Private Sub Form_Resize()
    Frame1.Width = ScaleWidth
    
    lsvScholarship.Height = ScaleHeight - (Frame1.Height + StatusBar.Height)
    lsvScholar.Height = ScaleHeight - (Frame1.Height + StatusBar.Height)
    
    lsvScholar.Width = ScaleWidth - (lsvScholarship.Width)
    
    StatusBar.Panels(2).Width = lsvScholar.Width
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
Public Sub RefreshScholarship()
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    Dim lv As ListItem
    
    sSQL = "SELECT *" & _
            " FROM tblScholarship"
     
    If ConnectRS(con, vRS, sSQL) = True Then
        lsvScholarship.ListItems.Clear
        Do Until vRS.EOF
            Set lv = lsvScholarship.ListItems.Add(, , vRS.Fields("Scholarship"))
                lv.SubItems(1) = vRS.Fields("ScholarshipID")
            vRS.MoveNext
        Loop
    End If
    
ReleaseAndExit:
    Set vRS = Nothing
End Sub

Private Sub lsvScholar_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        PopupMenu mdiController.popScholar
    End If
End Sub

Private Sub lsvScholarship_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        PopupMenu mdiController.popScholarship
    End If
End Sub
