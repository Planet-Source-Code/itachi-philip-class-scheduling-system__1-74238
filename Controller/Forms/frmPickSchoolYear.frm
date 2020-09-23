VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPickSchoolYear 
   BackColor       =   &H00D8E9EC&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "School Year"
   ClientHeight    =   3915
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5520
   Icon            =   "frmPickSchoolYear.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3915
   ScaleWidth      =   5520
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSelect 
      BackColor       =   &H00D8E9EC&
      Caption         =   "Select"
      Height          =   375
      Left            =   4560
      TabIndex        =   2
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00D8E9EC&
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3720
      TabIndex        =   1
      Top             =   120
      Width           =   855
   End
   Begin VB.TextBox txtFind 
      Height          =   315
      Left            =   480
      TabIndex        =   0
      Top             =   120
      Width           =   3180
   End
   Begin MSComctlLib.ListView listRecord 
      Height          =   3240
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   5325
      _ExtentX        =   9393
      _ExtentY        =   5715
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      PictureAlignment=   4
      _Version        =   393217
      Icons           =   "ilRecordIco"
      SmallIcons      =   "ilRecordIco"
      ForeColor       =   8399906
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
      MouseIcon       =   "frmPickSchoolYear.frx":492A
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Deparment"
         Object.Width           =   9366
      EndProperty
   End
   Begin MSComctlLib.ImageList ilRecordIco 
      Left            =   0
      Top             =   3840
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
            Picture         =   "frmPickSchoolYear.frx":5204
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPickSchoolYear.frx":579E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPickSchoolYear.frx":5D38
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPickSchoolYear.frx":62D2
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Find"
      ForeColor       =   &H00404040&
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   300
   End
End
Attribute VB_Name = "frmPickSchoolYear"
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


Dim tmpSchoolYear As String
Dim vRS As New ADODB.Recordset

Dim MaxEntryCount As Long
Dim CurRecPos As Long
Dim CurRecordCount As Long

Dim sOldSchoolYear As String
Dim sWC As String

Public Function GetItem(Optional TextObject As Variant, Optional lMaxEntryCount As Long = 15, Optional OldSchoolYear As String = "0000", Optional ExcludeClosed As Boolean = False) As String
    
    Dim sSQL As String
    
    
    'set fail to default
    GetItem = ""
    tmpSchoolYear = ""
    
    
    MaxEntryCount = lMaxEntryCount
    CurRecPos = 0
    sOldSchoolYear = OldSchoolYear
    
    If ExcludeClosed = True Then
        sWC = "AND ((tblSchoolYear.Locked)=No)"
    Else
        sWC = ""
    End If
    
    sSQL = "SELECT tblSchoolYear.SchoolYear AS lvKey, tblSchoolYear.SchoolYear" & _
            " From tblSchoolYear" & _
            " WHERE (((Val(Left([SchoolYear],4)))>" & Left(sOldSchoolYear, 4) & ")) " & sWC
    
    'add yr to list
    If ConnectRS(con, vRS, sSQL) = True Then
                
        If vRS.RecordCount > 0 Then
            FillList CurRecPos, MaxEntryCount

        Else
            'error
            MsgBox "No School Year  to be selected." & vbNewLine & "Please Add New School Year  first.", vbExclamation
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
    GetItem = tmpSchoolYear
End Function


Private Sub ReturnGetStudentID()
    If Len(GetLVKey(listRecord.SelectedItem)) > 0 Then
        tmpSchoolYear = GetLVKey(listRecord.SelectedItem)
        Unload Me
    End If
End Sub
Private Sub CancelGetStudentID()
    tmpSchoolYear = ""
    Unload Me
End Sub
Private Sub cmdCancel_Click()
    CancelGetStudentID
End Sub

Private Sub cmdFind_Click()
    Dim sSQL As String
    
    sSQL = "SELECT tblSchoolYear.SchoolYearTitle AS lvKey, tblSchoolYear.SchoolYearTitle" & _
            " From tblSchoolYear" & _
            " WHERE (((Val(Left([SchoolYearTitle],4)))>" & Left(sOldSchoolYear, 4) & ") AND ((SchoolYearTitle) like '%" & txtFind.Text & "%')) " & sWC


    If ConnectRS(con, vRS, sSQL) = True Then
        
        CurRecPos = 0
        
        
        If CurRecordCount > 0 Then
            
            FillList CurRecPos, MaxEntryCount

        Else
            'no result
            listRecord.ListItems.Clear

        End If
    Else
        MsgBox "FATAL ERROR: PickStudent.cmdFind_Click - Connectrs"
    End If
    
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



Private Sub listRecord_KeyDown(keycode As Integer, Shift As Integer)
    Dim curPos As Long
    If keycode = vbKeyDown Then
        If listRecord.SelectedItem.Index = listRecord.ListItems.count Then
                        
            keycode = 0
        End If
    End If
    
    If keycode = vbKeyUp Then
        If listRecord.SelectedItem.Index = 1 Then
            curPos = CurRecPos
            
            If curPos <> CurRecPos Then
                listRecord.SelectedItem.Selected = False
                listRecord.ListItems(listRecord.ListItems.count).Selected = True
            End If
            
            keycode = 0
        End If
    End If
End Sub

Private Sub listRecord_KeyUp(keycode As Integer, Shift As Integer)
       
    If keycode = vbKeyReturn Then ReturnGetStudentID
    
End Sub

Private Sub txtFind_Change()
    'delay 0.3 second
    'code by: VIncent J. Jamero
    '------------------------------------------------
    Static DelayStart As Single
    Static notFirst As Boolean
    DelayStart = GetTickCount + 300
    If notFirst = True Then Exit Sub
    notFirst = True
    While GetTickCount < DelayStart
        DoEvents
    Wend
    notFirst = False
    '------------------------------------------------
    'the next line will be if executed if user pause typing in 0.3 second
    
    'If Len(Trim(txtFind.Text)) > 0 Then
        cmdFind_Click
    'End If
End Sub

Private Sub txtFind_KeyDown(keycode As Integer, Shift As Integer)
    If keycode = vbKeyDown Then
        listRecord.SetFocus
    End If
End Sub


'end p---


