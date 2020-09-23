VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmBrowse 
   BackColor       =   &H00D8E9EC&
   Caption         =   "Browse Student"
   ClientHeight    =   8610
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   11400
   Icon            =   "frmBrowse.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8610
   ScaleWidth      =   11400
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ListView lsvStudent 
      Height          =   6495
      Left            =   0
      TabIndex        =   7
      Top             =   600
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   11456
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
         Text            =   "ID Number"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Student Name"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Gender"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "CourseID"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Yearlvl"
         Object.Width           =   0
      EndProperty
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00D8E9EC&
      Height          =   615
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   12855
      Begin VB.CommandButton cmdSearch 
         BackColor       =   &H00D8E9EC&
         Caption         =   "&Search"
         Height          =   360
         Left            =   8160
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Search"
         Top             =   160
         Width           =   975
      End
      Begin VB.ComboBox cboCriteria 
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
         ItemData        =   "frmBrowse.frx":492A
         Left            =   6000
         List            =   "frmBrowse.frx":4934
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   160
         Width           =   2055
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
         Height          =   360
         Left            =   840
         TabIndex        =   3
         Top             =   160
         Width           =   4695
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "as"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5640
         TabIndex        =   4
         Top             =   200
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Search"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   200
         Width           =   855
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   8235
      Width           =   11400
      _ExtentX        =   20108
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3528
            MinWidth        =   3528
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   10583
            MinWidth        =   10583
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ilRecordIco 
      Left            =   8760
      Top             =   2760
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
            Picture         =   "frmBrowse.frx":494D
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu popAction 
      Caption         =   "Action"
      Visible         =   0   'False
      Begin VB.Menu mnuOpen 
         Caption         =   "Open..."
      End
   End
End
Attribute VB_Name = "frmBrowse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sRS As New ADODB.Recordset
Private Sub cmdSearch_Click()
    If Me.cboCriteria.Text = "ID Number" Then
        GenerateListStudentID
    Else
        GenerateListLastName
    End If
End Sub

Private Sub Form_Activate()
    cboCriteria.ListIndex = 0
End Sub

Private Sub Form_Resize()
    lsvStudent.Height = ScaleHeight - (Frame1.Height + StatusBar1.Height)
    lsvStudent.Width = ScaleWidth
End Sub
Public Function GenerateListLastName()
Dim x As ListItem
Dim count As Integer
Dim str As String

str = "Select * From tblStudent"

count = 0

If ConnectRS(con, sRS, str) = True Then
    lsvStudent.ListItems.Clear
On Error Resume Next
    sRS.Filter = "LastName LIKE '" & txtSearch.Text & "%'"
    Do Until sRS.EOF
         Set x = lsvStudent.ListItems.Add(, , sRS!StudentID, 1, 1)
                   x.SubItems(1) = sRS!LastName & ", " & sRS!FirstName & " " & sRS!MiddleName & "."
                   x.SubItems(2) = sRS!Gender
                   x.SubItems(3) = sRS!CourseID

       sRS.MoveNext
    Loop
    
     If lsvStudent.ListItems.count = 0 Then
                MsgBox "Sorry, No Record matched" & vbCrLf & "Student May Not Yet Enrolled!" & vbCrLf & vbCrLf & _
                "Please check the information you provided.", vbInformation + vbOKOnly, ""
    End If
    
End If
Set sRS = Nothing
End Function
Public Function GenerateListStudentID()
Dim x As ListItem
Dim count As Integer
Dim str As String

str = "Select * From tblStudent"

count = 0

If ConnectRS(con, sRS, str) = True Then
    lsvStudent.ListItems.Clear
    On Error Resume Next
    sRS.Filter = "StudentID LIKE '" & txtSearch.Text & "%'"
    Do Until sRS.EOF
         Set x = lsvStudent.ListItems.Add(, , sRS!StudentID)
                   x.SubItems(1) = sRS!LastName & ", " & sRS!FirstName & " " & sRS!MiddleName & "."
                   x.SubItems(2) = sRS!Gender
                   x.SubItems(3) = sRS!CourseID
       sRS.MoveNext
    Loop
    If lsvStudent.ListItems.count = 0 Then
                MsgBox "Sorry, No Record matched" & vbCrLf & "Student May Not Yet Enrolled!" & vbCrLf & vbCrLf & _
                "Please check the information you provided.", vbInformation + vbOKOnly, ""
    End If
    
End If
Set sRS = Nothing
End Function
Private Sub lsvStudent_DblClick()
    mnuOpen_Click
End Sub

Private Sub lsvStudent_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        PopupMenu Me.popAction
    End If
End Sub

Private Sub mnuOpen_Click()
    Dim sKey As String
    Dim sCourseKey As String
    Dim sYearKey As Integer
    Dim Key As String
    
    sKey = lsvStudent.SelectedItem
    
    If sKey <> "" Then
        frmStudent.ShowStudentDetail (sKey)
        Unload Me
    End If
End Sub

Private Sub txtSearch_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
    Call cmdSearch_Click
End If
End Sub
