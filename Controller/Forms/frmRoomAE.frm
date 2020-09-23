VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmRoomAE 
   BackColor       =   &H00D8E9EC&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Classroom"
   ClientHeight    =   5520
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   5055
   Icon            =   "frmRoomAE.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5520
   ScaleWidth      =   5055
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   2640
      TabIndex        =   10
      Top             =   5040
      Width           =   1215
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   375
      Left            =   1200
      TabIndex        =   9
      Top             =   5040
      Width           =   1215
   End
   Begin VB.TextBox txtRoomID 
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
      Left            =   1320
      MaxLength       =   50
      TabIndex        =   8
      Top             =   6240
      Width           =   825
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00D8E9EC&
      Height          =   4815
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4815
      Begin MSComctlLib.ListView lsvDepartment 
         Height          =   3375
         Left            =   120
         TabIndex        =   11
         Top             =   1320
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   5953
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         Icons           =   "ilRecordIco"
         SmallIcons      =   "ilRecordIco"
         ColHdrIcons     =   "icoHeader"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Department"
            Object.Width           =   8070
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "DepartmentID"
            Object.Width           =   0
         EndProperty
      End
      Begin VB.TextBox txtCapacity 
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
         Left            =   3840
         MaxLength       =   50
         TabIndex        =   7
         Top             =   480
         Width           =   825
      End
      Begin VB.TextBox txtRoom 
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
         Left            =   2040
         MaxLength       =   50
         TabIndex        =   6
         Top             =   480
         Width           =   1785
      End
      Begin VB.TextBox txtBuilding 
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
         Left            =   120
         MaxLength       =   50
         TabIndex        =   5
         Top             =   480
         Width           =   1905
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Assigned Department"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   4
         Top             =   1080
         Width           =   2295
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Building"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Capacity"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   3840
         TabIndex        =   2
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Room"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2040
         TabIndex        =   1
         Top             =   240
         Width           =   405
      End
   End
   Begin MSComctlLib.ImageList icoHeader 
      Left            =   270
      Top             =   6015
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
            Picture         =   "frmRoomAE.frx":492A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRoomAE.frx":4EC4
            Key             =   ""
         EndProperty
      EndProperty
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
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRoomAE.frx":545E
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmRoomAE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rs As New ADODB.Recordset

Dim RoomID As String
Dim mFormState As String
Dim RecordAdded As Boolean
Dim RecordEdited As Boolean

Dim sDepartmentID As String

Dim CurrentRoom As vRoom

Dim MaxEntryCount As Long
Dim CurRecPos As Long
Dim CurStudentCount As Long

Public Function ShowForm(Optional sRoomTitle As String = "", Optional sYearLevelTitle As String = "", Optional sTeacherTitle As String = "") As Boolean
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    Dim sNewRoomID As String
    ShowForm = False
    RecordAdded = False
    
    Dim lv As ListItem
    
    
    sSQL = "SELECT tblDepartment.DepartmentID as lvKey, tblDepartment.DepartmentTitle,DepartmentID From tblDepartment"
    
       
    If GetNewRoomID(sNewRoomID) = Failed Then
        CatchError "Room", "ShowForm()", "GetNewRoomID(sNewID) = Failed"
        Exit Function
    End If

    RoomID = sNewRoomID
        
    If ConnectRS(con, rs, sSQL) = True Then
        lsvDepartment.ListItems.Clear
        Do Until rs.EOF
            Set lv = lsvDepartment.ListItems.Add(, , rs.Fields("DepartmentTitle"), 1, 1)
                lv.SubItems(1) = rs.Fields("lvKey")
            rs.MoveNext
        Loop
     End If
        
    
    mFormState = "ADD"
    
    Me.Show vbModal

    ShowForm = RecordAdded
End Function

Public Function ShowEdit(sRoomID As String) As Boolean
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    Dim vDepartment As tDepartment
    Dim lv As ListItem
    Dim bUnchecked As Boolean
    
    ShowEdit = False
    
    sSQL = "SELECT tblDepartment.DepartmentID as lvKey, tblDepartment.DepartmentTitle,DepartmentID From tblDepartment"
    
    If GetRoomByID(sRoomID, CurrentRoom) <> Success Then
        Exit Function
    End If
    
    If GetRoomDepartmentByID(sRoomID, CurrentRoom) <> Success Then
        Exit Function
    End If
    
    If GetDepartmentByID(CurrentRoom.Department, vDepartment) <> Success Then
                Exit Function
    End If
        
    If ConnectRS(con, rs, sSQL) = True Then
        lsvDepartment.ListItems.Clear
        Do Until rs.EOF
            Set lv = lsvDepartment.ListItems.Add(, , rs.Fields("DepartmentTitle"), 1, 1)
                lv.SubItems(1) = rs.Fields("lvKey")
            rs.MoveNext
        Loop
     End If
    
        RoomID = CurrentRoom.RoomID
        txtBuilding.Text = CurrentRoom.Building
        txtRoom.Text = CurrentRoom.Roomname
        txtCapacity.Text = CurrentRoom.Capacity
    

     For Each lv In lsvDepartment.ListItems
        bUnchecked = False
        If bUnchecked = False Then
                If lv.Text = vDepartment.DepartmentTitle Then
                    lv.Checked = True
                Else
                    bUnchecked = True
                    lv.Checked = False
                End If
        End If
     Next
    
    
    mFormState = "EDIT"
    
    Me.Show vbModal
    ShowEdit = RecordEdited
End Function
Private Function SaveData()

    Dim lv As ListItem
    
    If Not CheckTextBox(txtRoom, "Please Enter Room name") Then
        Exit Function
    End If
    
    If Not CheckTextBox(txtBuilding, "Please Enter Building name") Then
        Exit Function
    End If
    
    If Not CheckTextBox(txtCapacity, "Please Enter Room capacity") Then
        Exit Function
    End If
    
    Dim newRoom As vRoom
    
    newRoom.RoomID = RoomID
    newRoom.Roomname = txtRoom.Text
    newRoom.Building = txtBuilding.Text
    newRoom.Capacity = txtCapacity.Text
    
    
    Select Case AddRoom(newRoom)
        Case TranDBResult.Success
            SaveRoomDepartment (RoomID)
             MsgBox "Room Entry successfully added.", vbInformation
            RecordAdded = True
            Unload Me

        Case TranDBResult.DuplicateID
            MsgBox "ID already existed.", vbExclamation
            HLTxt txtRoomID
            SaveData = False
        
        Case TranDBResult.DuplicateTitle
            MsgBox "Title already existed.", vbExclamation
            HLTxt txtRoom
            SaveData = False
            
        Case Else
            MsgBox "Unknown Error.", vbExclamation
    End Select
End Function

Private Function SaveRoomDepartment(sRoomID As String) As Boolean
    
    Dim lv As ListItem
    Dim errorFound As Boolean
    
    SaveRoomDepartment = True
    
    For Each lv In lsvDepartment.ListItems
        If lv.Checked = True Then
            If AddRoomDepartment(sRoomID, lv.SubItems(1)) <> Success Then
                SaveRoomDepartment = False
            End If
        End If
    Next
    
End Function


Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    Select Case mFormState
        Case "ADD"
            SaveData
        Case "EDIT"
           UpdateData
    End Select
End Sub
Private Function UpdateData() As Boolean
    
    Dim newRoom As vRoom
    UpdateData = False

    If Not ValidateData Then Exit Function
    
    
    newRoom.RoomID = RoomID
    newRoom.Roomname = txtRoom.Text
    newRoom.Capacity = txtCapacity.Text
    newRoom.Building = txtBuilding.Text

    Select Case EditRoom(newRoom)
        Case TranDBResult.Success
            UpdateData = True
            MsgBox "Subject Entry successfully Edited.", vbInformation
            Unload Me
            
        Case TranDBResult.DuplicateTitle
            MsgBox "Title already existed.", vbExclamation
            HLTxt txtRoom
            UpdateData = False

        Case Else
            MsgBox "Unknown Error.", vbExclamation
            UpdateData = False
    End Select
End Function


Private Function ValidateData() As Boolean
 Dim lv As ListItem
    ValidateData = False

    If Not CheckTextBox(txtBuilding, "Please Enter Building") Then
        Exit Function
    End If

    If Not CheckTextBox(txtRoom, "Please Enter Room title") Then
        Exit Function
    End If
   
    If Not CheckTextBox(txtCapacity, "Please Enter Room capacity") Then
        Exit Function
    End If
    
    For Each lv In lsvDepartment.ListItems
        If lv.Checked = False Then
           MsgBox "Please select department to assign", vbCritical
           Exit Function
        End If
    Next
    
    ValidateData = True
End Function
Public Function AddRoomDepartment(sRoomsID As String, DepartmentID As String) As TranDBResult
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    Dim sChargeID As String
    
    AddRoomDepartment = Failed
    
    sChargeID = sRoomsID & "-" & String$(10 - Len(Trim(DepartmentID)), "0") & DepartmentID
    
    sSQL = "SELECT * FROM tblRoomDepartment"
    
    If ConnectRS(con, vRS, sSQL) = False Then
        AddRoomDepartment = NotConnected
        GoTo ReleaseAndExit
    End If
    
    vRS.AddNew
    vRS.Fields("RoomDeptID").Value = sChargeID
    vRS.Fields("RoomID").Value = sRoomsID
    vRS.Fields("DepartmentID").Value = DepartmentID
    vRS.Update
    
    AddRoomDepartment = Success
    
ReleaseAndExit:
    Set vRS = Nothing
End Function

