VERSION 5.00
Begin VB.Form frmSectionAE 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Section"
   ClientHeight    =   2190
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   4335
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2190
   ScaleWidth      =   4335
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2640
      TabIndex        =   8
      Top             =   1680
      Width           =   1455
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   1080
      TabIndex        =   7
      Top             =   1680
      Width           =   1455
   End
   Begin VB.TextBox txtDepartmentTitle 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   120
      MaxLength       =   50
      TabIndex        =   3
      Top             =   1080
      Width           =   3705
   End
   Begin VB.TextBox txtSectionTitle 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   120
      MaxLength       =   50
      TabIndex        =   2
      Top             =   360
      Width           =   3705
   End
   Begin VB.TextBox txtSectionID 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4860
      Locked          =   -1  'True
      MaxLength       =   10
      TabIndex        =   1
      Top             =   4230
      Width           =   945
   End
   Begin VB.CommandButton cmdGetItem 
      Appearance      =   0  'Flat
      BackColor       =   &H00D8E9EC&
      Height          =   345
      Left            =   3840
      Picture         =   "frmSectionAE.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1080
      Width           =   345
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Section ID"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   3600
      TabIndex        =   6
      Top             =   4320
      Width           =   960
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Department"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   120
      TabIndex        =   5
      Top             =   840
      Width           =   1110
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Section Title"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   120
      TabIndex        =   4
      Top             =   135
      Width           =   1140
   End
End
Attribute VB_Name = "frmSectionAE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim RecordAdded As Boolean
Public Function ShowForm(Optional sDepartmentTitle As String = "", Optional sYearLevelTitle As String = "", Optional sTeacherTitle As String = "") As Boolean
    
    Dim sNewSectionID As String
    
    ShowForm = False

    If DepartmentRecordExist <> Success Then
        MsgBox "Unable to continue Adding Section." & vbNewLine & "Department entries not exist", vbExclamation
        Unload Me
        Exit Function
    End If
    
    If YearLevelRecordExist <> Success Then
        MsgBox "Unable to continue Adding Section." & vbNewLine & "Year Level entries not exist", vbExclamation
        Unload Me
        Exit Function
    End If

    txtDepartmentTitle.Text = sDepartmentTitle

    If GetNewSectionID(sNewSectionID) = Success Then
        txtSectionID.Text = sNewSectionID
    Else
    End If

    Me.Show vbModal
    
    ShowForm = RecordAdded
End Function

Private Function SaveData() As Boolean
    
    Dim newSection As tSection
    Dim vDepartment As tDepartment
    Dim vYearlevel As tYearLevel
    Dim vTeacher As tTeacher
    
    'set default
    SaveData = False
    
    'validate date
    If Not ValidateData Then Exit Function
    
    
    
    'set/check departmentid
    If GetDepartmentByTitle(txtDepartmentTitle.Text, vDepartment) <> Success Then
        MsgBox "Invalid Department Title", vbExclamation
        HLTxt txtDepartmentTitle
        Exit Function
    End If
    

    newSection.SectionID = txtSectionID
    newSection.SectionTitle = txtSectionTitle
    newSection.DepartmentID = vDepartment.DepartmentID
    newSection.CreationDate = Now
    newSection.CreatedBy = CurrentUser.USERNAME
    

    Dim ir As Integer
    ir = AddSection(newSection)
    Select Case ir
        Case TranDBResult.Success
            SaveData = True
        
        
        Case TranDBResult.DuplicateID
            MsgBox "ID already existed.", vbExclamation
            HLTxt txtSectionID
            SaveData = False
        
        Case TranDBResult.DuplicateTitle
            MsgBox "Title already existed.", vbExclamation
            HLTxt txtSectionTitle
            SaveData = False

        Case Else
            CatchError "frmAddSetion", "SaveData", "Saving Section"
            SaveData = False
    End Select
End Function



Private Function ValidateData() As Boolean
    ValidateData = False

    If Not CheckTextBox(txtSectionID, "Please Enter Section ID") Then
        Exit Function
    End If

    If Not CheckTextBox(txtSectionTitle, "Please Enter Section Title") Then
        Exit Function
    End If

    If DepartmentExistByTitle(txtDepartmentTitle.Text) <> Success Then
        MsgBox "Invalid Department Title", vbExclamation
        HLTxt txtDepartmentTitle
        Exit Function
    End If
    
    ValidateData = True
End Function
Private Sub cmdCancel_Click()
    Unload Me
End Sub
Private Sub cmdGetItem_Click()
    Dim sDepartmentID As String
    Dim sDepartmentTitle As String

    sDepartmentID = frmPickDepartment.GetItem(txtDepartmentTitle, sDepartmentTitle)
    If sDepartmentID <> "" Then
        txtDepartmentTitle = sDepartmentTitle
    End If
End Sub

Private Sub cmdSave_Click()
    If SaveData Then
    
        MsgBox "SECTION Entry successfully added.", vbInformation
        RecordAdded = True
        
        If MsgBox("Do you want to add schedule for this section?", vbQuestion + vbYesNo) = vbYes Then
            'frmAddSectionOffering.ShowForm txtYearLevelTitle.Text & " - " & txtSectionTitle.Text
        End If
        
        Unload Me
    End If
End Sub

