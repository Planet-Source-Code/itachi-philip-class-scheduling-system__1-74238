VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmUsage 
   Caption         =   "Class Room"
   ClientHeight    =   8520
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   8880
   Icon            =   "frmUsage.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8520
   ScaleWidth      =   8880
   WindowState     =   2  'Maximized
   Begin TabDlg.SSTab sSTab 
      Height          =   10095
      Left            =   0
      TabIndex        =   6
      Top             =   600
      Width           =   14895
      _ExtentX        =   26273
      _ExtentY        =   17806
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Utilization"
      TabPicture(0)   =   "frmUsage.frx":492A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "StatusBar1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lsvClassroom"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lsvUtilization"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "RoomStat"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lsvDepartmentRoom"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "ilRecordIco"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "icoHeader"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).ControlCount=   7
      Begin MSComctlLib.ImageList icoHeader 
         Left            =   3720
         Top             =   0
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
               Picture         =   "frmUsage.frx":4946
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmUsage.frx":4EE0
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageList ilRecordIco 
         Left            =   4440
         Top             =   0
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
               Picture         =   "frmUsage.frx":547A
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ListView lsvDepartmentRoom 
         Height          =   2895
         Left            =   6120
         TabIndex        =   7
         Top             =   5280
         Width           =   9015
         _ExtentX        =   15901
         _ExtentY        =   5106
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
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Assigned Department"
            Object.Width           =   15901
         EndProperty
      End
      Begin MSComctlLib.StatusBar RoomStat 
         Height          =   375
         Left            =   0
         TabIndex        =   8
         Top             =   8280
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
            NumPanels       =   2
            BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            EndProperty
            BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               Object.Width           =   8113
               MinWidth        =   8113
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ListView lsvUtilization 
         Height          =   4455
         Left            =   6120
         TabIndex        =   9
         Top             =   360
         Width           =   9135
         _ExtentX        =   16113
         _ExtentY        =   7858
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
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   5
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
            Text            =   "Enrolled"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Schedule"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Lecturer"
            Object.Width           =   3528
         EndProperty
      End
      Begin MSComctlLib.ListView lsvClassroom 
         Height          =   7935
         Left            =   0
         TabIndex        =   10
         Top             =   360
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   13996
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
            Text            =   "Building"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Room"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Capacity"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.StatusBar StatusBar1 
         Height          =   375
         Left            =   6120
         TabIndex        =   11
         Top             =   9480
         Width           =   9015
         _ExtentX        =   15901
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
            NumPanels       =   3
            BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            EndProperty
            BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               Object.Width           =   10583
               MinWidth        =   10583
            EndProperty
            BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            EndProperty
         EndProperty
      End
   End
   Begin VB.PictureBox picLine 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      Height          =   45
      Index           =   1
      Left            =   0
      ScaleHeight     =   45
      ScaleWidth      =   8880
      TabIndex        =   1
      Top             =   585
      Width           =   8880
   End
   Begin VB.PictureBox picLine 
      Align           =   1  'Align Top
      BackColor       =   &H80000010&
      BorderStyle     =   0  'None
      Height          =   10
      Index           =   5
      Left            =   0
      ScaleHeight     =   15
      ScaleWidth      =   8880
      TabIndex        =   0
      Top             =   570
      Width           =   8880
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3600
      Top             =   1080
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   41
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUsage.frx":5A14
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUsage.frx":62EE
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUsage.frx":7C80
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUsage.frx":9612
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUsage.frx":AFA4
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUsage.frx":C936
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUsage.frx":E2C8
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUsage.frx":FC5A
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUsage.frx":115EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUsage.frx":12F7E
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUsage.frx":14912
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUsage.frx":15964
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUsage.frx":16244
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUsage.frx":17296
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUsage.frx":182E8
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUsage.frx":1933A
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUsage.frx":1A38C
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUsage.frx":1B3DE
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUsage.frx":1BCBA
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUsage.frx":1C594
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUsage.frx":1CE6E
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUsage.frx":1D748
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUsage.frx":1E022
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUsage.frx":1E8FC
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUsage.frx":1F1D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUsage.frx":1FAB0
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUsage.frx":2038A
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUsage.frx":20C64
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUsage.frx":2153E
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUsage.frx":21E18
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUsage.frx":226F2
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUsage.frx":22FCC
            Key             =   ""
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUsage.frx":238A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUsage.frx":24180
            Key             =   ""
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUsage.frx":24A5A
            Key             =   ""
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUsage.frx":25334
            Key             =   ""
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUsage.frx":25C0E
            Key             =   ""
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUsage.frx":264E8
            Key             =   ""
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUsage.frx":26DC2
            Key             =   ""
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUsage.frx":2769C
            Key             =   ""
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUsage.frx":27F76
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   8880
      _ExtentX        =   15663
      _ExtentY        =   1005
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Print"
            Object.ToolTipText     =   "Print room usage"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
      EndProperty
      Begin VB.ComboBox cboSY 
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
         ItemData        =   "frmUsage.frx":28850
         Left            =   840
         List            =   "frmUsage.frx":2885D
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   120
         Width           =   1935
      End
      Begin VB.ComboBox cboDepartment 
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
         Left            =   4680
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   120
         Width           =   6375
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
         ItemData        =   "frmUsage.frx":28885
         Left            =   2760
         List            =   "frmUsage.frx":28892
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   120
         Width           =   1935
      End
   End
End
Attribute VB_Name = "frmUsage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim cIRowCount              As Integer
Dim vRS As New ADODB.Recordset
Dim sDefaultSQL As String

Private Const sAllFields = "ALL FIELDS"
Public Conflict As Boolean

Dim MaxEntryCount As Long
Dim CurRecPos As Long
Dim CurStudentCount As Long
Private Sub DepartmentList()
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    sSQL = "SELECT DepartmentTitle" & _
            " FROM tblDepartment" & _
            " ORDER BY DepartmentTitle"
     
    If ConnectRS(con, vRS, sSQL) = False Then
        MsgBox "ERROR"
        GoTo ReleaseAndExit
    End If
    
    cboDepartment.Clear
    While vRS.EOF = False
        
        cboDepartment.AddItem (vRS.Fields("DepartmentTitle"))
        vRS.MoveNext
    
    Wend
ReleaseAndExit:
    Set vRS = Nothing
End Sub
Private Sub RefreshSYList()
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    'On Error GoTo ReleaseAndExit
    
    sSQL = "SELECT tblSchoolYear.SchoolYear" & _
            " FROM tblSchoolYear" & _
            " ORDER BY tblSchoolYear.SchoolYear"
     
    If ConnectRS(con, vRS, sSQL) = False Then
        MsgBox "ERROR"
        GoTo ReleaseAndExit
    End If
    
    cboSY.Clear
    While vRS.EOF = False
        
        cboSY.AddItem (vRS.Fields("SchoolYear"))
        vRS.MoveNext
    
    Wend
    
    cboSY.Text = CurrentSchoolYear.SchoolYearTitle
    cboSemester.Text = CurrentSemester.Semester
    
ReleaseAndExit:
    Set vRS = Nothing
End Sub
Public Sub ShowFormList(Optional iMaxEntryCount As Long = 21, Optional iCurRecPos As Long = 0)

    MaxEntryCount = iMaxEntryCount
    CurRecPos = iCurRecPos

    sDefaultSQL = "SELECT tblRoom.RoomID,tblRoom.Building,tblRoom.Room,tblRoom.Capacity " & _
                    " From tblRoom"

    If ConnectRS(con, vRS, sDefaultSQL & " ORDER BY tblRoom.Room") Then
        FillList vRS
        
        RefreshSYList
        DepartmentList
            
        Me.Show
        Me.SetFocus
        
    Else
        MsgBox "Unable to show Room List.", vbCritical
        Unload Me
    End If
    
End Sub
Private Function FillList(ByRef vRS As ADODB.Recordset)
        
        mdiController.MousePointer = vbHourglass

        
        UnSortLV lsvClassroom
        
        FillRecordToList vRS, lsvClassroom, KeyStudent, CurRecPos, MaxEntryCount, , True
        
        SortLV lsvClassroom, lsvClassroom.SortKey, lsvClassroom.SortOrder, False

        lsvClassroom_Click
        
         
        
        mdiController.MousePointer = vbDefault
End Function
Public Sub Form_Refresh()
    vRS.Requery
    FillList vRS
End Sub

Private Sub Form_Load()
     RefreshSYList
    DepartmentList
End Sub

Private Sub Form_Resize()
    sSTab.Height = ScaleHeight - Toolbar1.Height
    sSTab.Width = ScaleWidth
    
    
    lsvClassroom.Height = sSTab.Height - (RoomStat.Height + 400)
    
    lsvUtilization.Width = sSTab.Width - (lsvClassroom.Width + 50)
    lsvUtilization.Height = sSTab.Height - (lsvDepartmentRoom.Height + 400)

    RoomStat.Width = lsvClassroom.Width
    RoomStat.Top = lsvClassroom.Height + 350
    
    lsvDepartmentRoom.Width = sSTab.Width - lsvClassroom.Width
    lsvDepartmentRoom.Top = lsvUtilization.Height + 350
    lsvDepartmentRoom.Height = sSTab.Height - (lsvUtilization.Height + StatusBar1.Height + 400)
    
    StatusBar1.Width = lsvDepartmentRoom.Width
    StatusBar1.Top = lsvClassroom.Height + 350
    
End Sub

Private Sub lsvClassroom_Click()
    Dim totalPage As Long
    Dim curPage As Long
    If lsvClassroom.ListItems.count < 1 Then
        RoomStat.Panels(2).Text = "No Record"
    Else
    RoomStat.Panels(2).Text = "Selected Entry: " & lsvClassroom.SelectedItem.Index + CurRecPos & "/" & CurStudentCount
    FillRoomUsage GetLVKey(lsvClassroom.SelectedItem)
    Call RoomDepartment
    End If
End Sub
Private Sub RoomDepartment()
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
Private Sub FillRoomUsage(sRoom As String)
Dim sSQL As String
Dim lst As ListItem
Dim lvItem As ListItem
Dim num1, num2 As Integer
Dim dTimeIn, dTimeOut As String
Dim vRS As New ADODB.Recordset

Dim time As Integer

sSQL = "SELECT tblSubjectOffering.SubjectOfferingID, tblSubject.SubjectTitle, tblSubjectOffering.Days & ' ' & tblSubjectOffering.TimeIn & '- ' & tblSubjectOffering.TimeOut AS Schedule, [LastName] & ', ' & [FirstName] & ' ' & [MiddleName] AS TeacherFullName, tblSection.SectionID, tblSection.SectionTitle, tblRoom.Building & '- ' & tblRoom.Room AS Room, tblRoom.RoomID, tblSubjectOffering.TimeIn, tblSubjectOffering.TimeOut, tblSubjectOffering.Days " & _
"FROM tblSubject INNER JOIN (tblSection INNER JOIN (tblRoom INNER JOIN (tblTeacher INNER JOIN tblSubjectOffering ON tblTeacher.TeacherID = tblSubjectOffering.TeacherID) ON tblRoom.RoomID = tblSubjectOffering.RoomID) ON tblSection.SectionID = tblSubjectOffering.SectionID) ON tblSubject.SubjectID = tblSubjectOffering.SubjectID " & _
"WHERE tblRoom.RoomID ='" & sRoom & "' and tblSubjectOffering.SchoolYear='" & cboSY.Text & "' and tblSubjectOffering.Semester='" & cboSemester.Text & "'" & _
"ORDER BY tblSubjectOffering.TimeIn;"

If ConnectRS(con, vRS, sSQL) = True Then
lsvUtilization.ListItems.Clear
Do Until vRS.EOF
    Set lst = lsvUtilization.ListItems.Add(, , vRS.Fields("SubjectTitle"), 1, 1)
            lst.SubItems(1) = vRS.Fields("SectionTitle")
            lst.SubItems(2) = vRS.Fields("Room")
            lst.SubItems(3) = vRS.Fields("Schedule")
            lst.SubItems(4) = vRS.Fields("TeacherFullname")
     vRS.MoveNext
Loop
End If
Set vRS = Nothing
End Sub

