VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPickSubject 
   BackColor       =   &H00D8E9EC&
   Caption         =   "Subject"
   ClientHeight    =   4905
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   4590
   Icon            =   "frmPickSubject.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4905
   ScaleWidth      =   4590
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSelect 
      BackColor       =   &H00D8E9EC&
      Caption         =   "Select"
      Height          =   375
      Left            =   3600
      TabIndex        =   1
      Top             =   4440
      Width           =   855
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00D8E9EC&
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2640
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
            Picture         =   "frmPickSubject.frx":492A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPickSubject.frx":4EC4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPickSubject.frx":545E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPickSubject.frx":59F8
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
Attribute VB_Name = "frmPickSubject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

Private Type POINTAPI
        x As Long
        y As Long
End Type

Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type
Dim r As RECT
Dim Alignable As Boolean

Dim vRS As New ADODB.Recordset

Dim sGetSubjectTitle As String


Public Function GetSubjectTitle(Optional ByRef TextObject) As String
    Dim r As RECT
    Dim p As POINTAPI
    Dim vSubject As tSubject
    
    If SubjectRecordExist <> Success Then
        MsgBox "There are no record yet in Subject Entries", vbExclamation
        CancelGetSubjectTitle
        Exit Function
    End If
    
    FillList
    
   
    'show
    Me.Show vbModal
    
    'return
    GetSubjectTitle = sGetSubjectTitle
End Function



Private Sub CancelGetSubjectTitle()
    sGetSubjectTitle = ""
    Unload Me
End Sub
Private Sub ReturnGetSubjectTitle()
    sGetSubjectTitle = GetLVKey(listRecord.SelectedItem)
    Unload Me
End Sub


Private Sub FillList()
    
        If ConnectRS(con, vRS, "SELECT tblSubject.SubjectTitle as lvKey,tblSubject.SubjectTitle FROM tblSubject;") Then
            FillRecordToList vRS, listRecord, KeySubject
        End If
    
End Sub

Private Sub cmdCancel_Click()
    CancelGetSubjectTitle
End Sub

Private Sub cmdSelect_Click()
    ReturnGetSubjectTitle
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyEscape
            CancelGetSubjectTitle
        Case vbKeyReturn
            ReturnGetSubjectTitle
    End Select
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set vRS = Nothing
End Sub

Private Sub listRecord_DblClick()
    ReturnGetSubjectTitle
End Sub


