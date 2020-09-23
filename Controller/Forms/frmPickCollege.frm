VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmPickCollege 
   BackColor       =   &H00808080&
   BorderStyle     =   0  'None
   Caption         =   "frmPickCollege"
   ClientHeight    =   4560
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4650
   LinkTopic       =   "Form1"
   ScaleHeight     =   4560
   ScaleWidth      =   4650
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00D8E9EC&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4455
      Index           =   0
      Left            =   40
      ScaleHeight     =   297
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   303
      TabIndex        =   0
      Top             =   40
      Width           =   4545
      Begin VB.PictureBox b8Container1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   3675
         Left            =   30
         ScaleHeight     =   3645
         ScaleWidth      =   4395
         TabIndex        =   3
         Top             =   270
         Width           =   4425
         Begin MSComctlLib.ListView listRecord 
            Height          =   3570
            Left            =   0
            TabIndex        =   4
            Top             =   0
            Width           =   4365
            _ExtentX        =   7699
            _ExtentY        =   6297
            View            =   3
            LabelEdit       =   1
            MultiSelect     =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            HideColumnHeaders=   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            Icons           =   "ilRecordIco"
            SmallIcons      =   "ilRecordIco"
            ForeColor       =   -2147483640
            BackColor       =   16777215
            Appearance      =   0
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
               Object.Width           =   3528
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Object.Width           =   10583
            EndProperty
         End
      End
      Begin VB.CommandButton cmdCancel 
         BackColor       =   &H00D8E9EC&
         Caption         =   "Cancel"
         Height          =   375
         Left            =   2640
         TabIndex        =   2
         Top             =   3960
         Width           =   855
      End
      Begin VB.CommandButton cmdSelect 
         BackColor       =   &H00D8E9EC&
         Caption         =   "Select"
         Height          =   375
         Left            =   3480
         TabIndex        =   1
         Top             =   3960
         Width           =   855
      End
      Begin VB.Image imgIcon 
         Height          =   240
         Left            =   120
         Picture         =   "frmPickCollege.frx":0000
         Stretch         =   -1  'True
         Top             =   0
         Width           =   240
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Select College"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   195
         Left            =   480
         TabIndex        =   5
         Top             =   45
         Width           =   1185
      End
      Begin VB.Image Image4 
         Height          =   135
         Left            =   0
         Picture         =   "frmPickCollege.frx":058A
         Stretch         =   -1  'True
         Top             =   0
         Width           =   5085
      End
      Begin VB.Image Image2 
         Height          =   405
         Left            =   0
         Picture         =   "frmPickCollege.frx":0627
         Stretch         =   -1  'True
         Top             =   3960
         Width           =   6495
      End
   End
   Begin MSComctlLib.ImageList ilRecordIco 
      Left            =   240
      Top             =   5280
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
            Picture         =   "frmPickCollege.frx":06C4
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPickCollege.frx":0C5E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPickCollege.frx":11F8
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPickCollege.frx":1792
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmPickCollege"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type
Dim r As RECT
Dim Alignable As Boolean


Dim tmpCollege As String
Dim vRS As New ADODB.Recordset

Dim MaxEntryCount As Long
Dim CurRecPos As Long
Dim CurRecordCount As Long

Dim sOldCollege As String

Dim sGetCollegeTitle As String
Dim sCollegeID As String


Public Function GetItem(Optional TextObject As Variant, Optional ByRef sCollegeTitle As String, Optional lMaxEntryCount As Long = 15, Optional OldCollege As String = "0000", Optional ExcludeClosed As Boolean = False) As String
    
    Dim sSQL As String
    Dim vCollege As tCollege
    
    'set fail to default
    GetItem = ""
    tmpCollege = ""
    
    
    MaxEntryCount = lMaxEntryCount
    CurRecPos = 0
    
    sCollegeID = ""
    sGetCollegeTitle = ""
    
    If CollegeRecordExist <> Success Then
        MsgBox "There are no record yet in College Entries", vbExclamation
        Exit Function
    End If
    
    
    sSQL = "SELECT tblCollege.CollegeID as lvKey,tblCollege.CollegeName" & _
            " FROM tblCollege" & _
            " ORDER BY tblCollege.CollegeName"
            
    If ConnectRS(con, vRS, sSQL) = True Then
        
        If vRS.RecordCount > 0 Then
            FillList CurRecPos, MaxEntryCount
        Else
            MsgBox "No College  to be selected." & vbNewLine & "Please Add New College  first.", vbExclamation
            Unload Me
            Exit Function
        End If
    Else
        'error
    End If

    'get pos
    If Not IsMissing(TextObject) Then
        GetWindowRect TextObject.hwnd, r
        Alignable = True
        Form_Activate
    Else
        Alignable = False
    End If
    
    'show form
    Me.Show vbModal
    
    'this next line will not execute unload this for will be unloaded
    sCollegeTitle = sGetCollegeTitle
    GetItem = tmpCollege
End Function


Private Sub ReturnGetStudentID()
    If Len(GetLVKey(listRecord.SelectedItem)) > 0 Then
        sGetCollegeTitle = listRecord.SelectedItem.Text
        tmpCollege = GetLVKey(listRecord.SelectedItem)
        Unload Me
    End If
End Sub

Private Sub CancelGetStudentID()
    tmpCollege = ""
    Unload Me
End Sub
Private Sub cmdCancel_Click()
    CancelGetStudentID
End Sub

Private Sub cmdSelect_Click()
    ReturnGetStudentID
End Sub



Private Sub Form_Activate()
    Dim NewLeft As Long
    Dim NewTop As Long
    
    If Alignable = True Then
        If (r.Left * Screen.TwipsPerPixelX + Me.Width) > Screen.Width Then
            NewLeft = (r.Right * Screen.TwipsPerPixelX) - Me.Width
        Else
            NewLeft = r.Left * Screen.TwipsPerPixelX
        End If
        
        If (r.Bottom * Screen.TwipsPerPixelY + Me.Height) > Screen.Height Then
            NewTop = (r.Top * Screen.TwipsPerPixelY) - Me.Height
            If NewTop < 0 Then NewTop = 0
        Else
            NewTop = r.Bottom * Screen.TwipsPerPixelY
        End If
        
        Me.Left = NewLeft
        Me.Top = NewTop
        
    Else
    
        CenterForm Me
        
    End If
End Sub
Private Sub listRecord_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    SortLV listRecord, ColumnHeader.Index - 1
End Sub

Private Sub listRecord_DblClick()
    ReturnGetStudentID
End Sub

Private Function FillList(lStart As Long, dCount As Long) As Boolean
    FillRecordToList vRS, listRecord, KeyStudent, lStart, dCount, , True
End Function



