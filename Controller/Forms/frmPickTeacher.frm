VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPickTeacher 
   BackColor       =   &H00D8E9EC&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Faculty/Lecturer"
   ClientHeight    =   4905
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   4575
   Icon            =   "frmPickTeacher.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4905
   ScaleWidth      =   4575
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00D8E9EC&
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2640
      TabIndex        =   1
      Top             =   4440
      Width           =   855
   End
   Begin VB.CommandButton cmdSelect 
      BackColor       =   &H00D8E9EC&
      Caption         =   "Select"
      Height          =   375
      Left            =   3600
      TabIndex        =   0
      Top             =   4440
      Width           =   855
   End
   Begin MSComctlLib.ImageList ilRecordIco 
      Left            =   0
      Top             =   5400
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPickTeacher.frx":492A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPickTeacher.frx":4EC4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPickTeacher.frx":545E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPickTeacher.frx":59F8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView listRecord 
      Height          =   4290
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   4365
      _ExtentX        =   7699
      _ExtentY        =   7567
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "imgListStudent"
      SmallIcons      =   "imgListStudent"
      ForeColor       =   -2147483640
      BackColor       =   16777215
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
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Object.Width           =   8440
      EndProperty
   End
End
Attribute VB_Name = "frmPickTeacher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim tmpTeacherID As String
Dim tmpTeacherFullName As String

Public Function GetTeacherID(Optional ByRef sTeacherFullName As String) As String
            
    'set fail to default
    GetTeacherID = ""
        
    'add yr to list
    If Not FillList Then
        MsgBox "There is no TEACHER entries to display.", vbExclamation
        Unload Me
        Exit Function
    End If

    Me.Show vbModal
    
    'return
    sTeacherFullName = IIf(IsMissing(sTeacherFullName), "", tmpTeacherFullName)
    GetTeacherID = tmpTeacherID
End Function


Private Sub ReturnGetTeacherID()
    If Len(GetLVKey(listRecord.SelectedItem)) > 0 Then
    
        tmpTeacherFullName = listRecord.SelectedItem.Text
        tmpTeacherID = GetLVKey(listRecord.SelectedItem)
        
        'call return
        Unload Me
    
    End If
End Sub
Private Sub CancelGetTeacherID()
    tmpTeacherID = ""
    Unload Me
End Sub

Private Sub cmdCancel_Click()
    CancelGetTeacherID
End Sub

Private Sub cmdSelect_Click()
    ReturnGetTeacherID
End Sub



Private Sub listRecord_DblClick()
    ReturnGetTeacherID
End Sub

Private Function FillList() As Boolean
    Dim vRS As New ADODB.Recordset
    
    If ConnectRS(con, vRS, "SELECT tblTeacher.TeacherID AS lvKEY, [tblTeacher]![LastName]+', '+[tblTeacher]![FirstName]+' '+[tblTeacher]![MiddleName] AS [Full Name] FROM tblTeacher;") = True Then
        If AnyRecordExisted(vRS) Then
            
            FillRecordToList vRS, listRecord, KeyTeacher
            
            FillList = True
        Else
            FillList = False
        End If
    Else
        FillList = False
    End If
    Set vRS = Nothing
End Function

