VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form PickYearLevel 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Select Year Level"
   ClientHeight    =   5085
   ClientLeft      =   0
   ClientTop       =   -30
   ClientWidth     =   4575
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   -1  'True
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   339
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   305
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00D8E9EC&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5055
      Left            =   15
      ScaleHeight     =   337
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   303
      TabIndex        =   0
      Top             =   15
      Width           =   4545
      Begin VB.CommandButton cmdSelect 
         BackColor       =   &H00D8E9EC&
         Caption         =   "Select"
         Height          =   375
         Left            =   3600
         TabIndex        =   5
         Top             =   4680
         Width           =   855
      End
      Begin VB.CommandButton cmdCancel 
         BackColor       =   &H00D8E9EC&
         Caption         =   "Cancel"
         Height          =   375
         Left            =   2760
         TabIndex        =   4
         Top             =   4680
         Width           =   855
      End
      Begin VB.PictureBox b8Container1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   4395
         Left            =   30
         ScaleHeight     =   4365
         ScaleWidth      =   4455
         TabIndex        =   1
         Top             =   270
         Width           =   4485
         Begin MSComctlLib.ListView listRecord 
            Height          =   4290
            Left            =   60
            TabIndex        =   2
            Top             =   60
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
               Object.Width           =   3175
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Object.Width           =   8440
            EndProperty
         End
      End
      Begin VB.Image Image2 
         Height          =   405
         Left            =   0
         Picture         =   "frmYearLevelPicker.frx":0000
         Stretch         =   -1  'True
         Top             =   4650
         Width           =   6495
      End
      Begin VB.Image imgIcon 
         Height          =   240
         Left            =   30
         Picture         =   "frmYearLevelPicker.frx":009D
         Stretch         =   -1  'True
         Top             =   30
         Width           =   240
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Select Year Level"
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
         Left            =   390
         TabIndex        =   3
         Top             =   30
         Width           =   1455
      End
      Begin VB.Image Image4 
         Height          =   135
         Left            =   0
         Picture         =   "frmYearLevelPicker.frx":0627
         Stretch         =   -1  'True
         Top             =   0
         Width           =   5085
      End
   End
End
Attribute VB_Name = "PickYearLevel"
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


Dim sYearLevel As String

Public Function GetYearLevelTitle() As String
    
    Dim r As RECT
    Dim p As POINTAPI
        
    
    If Not YearLevelRecordExisted Then
        MsgBox "There are no revords yeat in Year Level.", vbExclamation
        Unload Me
        Exit Function
    End If
    
    
    'add yr to list
    If Not FillList Then
        'temp
        MsgBox "There are no revords yeat in Year Level.", vbExclamation
        Unload Me
        Exit Function
    End If
    


    Me.Show vbModal
    
    
    'return
    GetYearLevelTitle = sYearLevel
End Function





Private Function FillList() As Boolean

    Dim vRS As New ADODB.Recordset
        
        If ConnectRS(con, vRS, "SELECT tblYearLevel.YearLevelTitle as lvKey, tblYearLevel.YearLevelTitle FROM tblYearLevel;") Then
            If AnyRecordExisted(vRS) Then
                FillRecordToList vRS, listRecord, KeyYearLevel
                FillList = True
            Else
                FillList = False
            End If
        Else
            FillList = False
        End If
    Set vRS = Nothing
End Function


Private Sub ReturnYearLevel()
    If Len(GetLVKey(listRecord.SelectedItem)) < 0 Then
        MsgBox "Please Select in the list.", vbInformation
        listRecord.SetFocus
    Else
        sYearLevel = GetLVKey(listRecord.SelectedItem)
        Unload Me
    End If
End Sub
Private Sub CancelYearLevel()
    sYearLevel = ""
    Unload Me
End Sub







Private Sub lblCancel_Click()
    CancelYearLevel
End Sub

Private Sub lblSelect_Click()
    Call ReturnYearLevel
End Sub

Private Sub cmdCancel_Click()
    CancelYearLevel
End Sub

Private Sub cmdSelect_Click()
    ReturnYearLevel
End Sub

Private Sub listRecord_DblClick()
    Call ReturnYearLevel
End Sub



