VERSION 5.00
Begin VB.Form frmDepartmentAE 
   BackColor       =   &H00D8E9EC&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Department"
   ClientHeight    =   2595
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   5850
   Icon            =   "frmDepartmentAE.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1533.213
   ScaleMode       =   0  'User
   ScaleWidth      =   5492.833
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtCollege 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   120
      MaxLength       =   255
      TabIndex        =   6
      Top             =   1560
      Width           =   5265
   End
   Begin VB.CommandButton cmdGetCollege 
      Appearance      =   0  'Flat
      BackColor       =   &H00D8E9EC&
      Height          =   345
      Left            =   5400
      Picture         =   "frmDepartmentAE.frx":492A
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1560
      Width           =   345
   End
   Begin VB.TextBox txtDepartmentID 
      BackColor       =   &H00C0FFFF&
      Height          =   345
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   5565
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Default         =   -1  'True
      Height          =   390
      Left            =   3120
      TabIndex        =   2
      Top             =   2040
      Width           =   1260
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   390
      Left            =   4500
      TabIndex        =   3
      Top             =   2040
      Width           =   1140
   End
   Begin VB.TextBox txtDepartmentTitle 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   120
      MaxLength       =   255
      TabIndex        =   1
      Top             =   960
      Width           =   5565
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Department Ref. #:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "College"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   7
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Department Name"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   1575
   End
End
Attribute VB_Name = "frmDepartmentAE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim RecordAdded As Boolean
Dim DepartmentID, CollegeID As String
Dim currentDepartment As tDepartment
Dim curCollege As tCollege
Public mFormState As String

Dim RecordEdited As Boolean

Public Function ShowEdit(sDepartmentID As String) As Boolean
        
    
        If GetDepartmentByID(sDepartmentID, currentDepartment) <> Success Then
            MsgBox "Unable to continue editing Department Information: Department ID not found!", vbCritical
            Exit Function
        End If
        
        If GetCollegeByID(currentDepartment.CollegeID, curCollege) <> Success Then
            MsgBox "Unable to continue editing Department Information: Department ID not found!", vbCritical
            Exit Function
        End If
        

    txtDepartmentID.Text = currentDepartment.DepartmentID
    txtDepartmentTitle.Text = currentDepartment.DepartmentTitle
    txtCollege.Text = curCollege.CollegeTitle
    CollegeID = curCollege.CollegeID
    
    mFormState = "EDIT"
    
    Me.Show vbModal

    ShowEdit = RecordEdited
    
    
End Function
Public Function ShowProperties(sDepartment As String) As Boolean
        
    
        If GetDepartmentByTitle(sDepartment, currentDepartment) <> Success Then
            MsgBox "Unable to continue editing Department Information: Department Reference # not found!", vbCritical
            Exit Function
        End If
        
        If GetCollegeByID(currentDepartment.CollegeID, curCollege) <> Success Then
            MsgBox "Unable to continue editing Department Information: College Reference # not found!", vbCritical
            Exit Function
        End If


    txtDepartmentID.Text = currentDepartment.DepartmentID
    txtDepartmentTitle.Text = currentDepartment.DepartmentTitle
    txtCollege.Text = curCollege.CollegeTitle
    CollegeID = curCollege.CollegeID
    
    mFormState = "EDIT"
    
    Me.Show vbModal

    ShowProperties = RecordEdited
    
    
End Function

Public Function ShowForm() As Boolean
    
    Dim sNewID As String
    
    If GetNewDepartmentID(sNewID) = Failed Then
        CatchError "Department", "ShowForm()", "GetNewDepartmentID(sNewID) = Failed"
        Exit Function
    End If

    txtDepartmentID.Text = sNewID
    
    mFormState = "ADD"

    Me.Show vbModal
    'return
    ShowForm = RecordAdded
    
End Function


Private Sub cmdCancel_Click()
    Unload Me
End Sub


Private Function SaveNewDepartment()
    
    If Not CheckTextBox(txtDepartmentTitle, "Please enter Department Title") Then
        Exit Function
    End If
    
    If Not CheckTextBox(txtCollege, "Please enter College Title") Then
        Exit Function
    End If
    
    'save
    Dim newDepartment As tDepartment
    
    newDepartment.DepartmentID = txtDepartmentID.Text
    newDepartment.DepartmentTitle = txtDepartmentTitle.Text
    newDepartment.CollegeID = CollegeID

    Select Case AddDepartment(newDepartment)
        Case TranDBResult.Success
            MsgBox "New Department succesfully added", vbInformation
            RecordAdded = True
            Unload Me
            
        Case TranDBResult.DuplicateID
            MsgBox "Invalid Department ID!" & vbNewLine & "The Department ID that you have entered is already existed. Enter another Department ID.", vbExclamation
            HLTxt txtDepartmentID
            
        Case TranDBResult.DuplicateTitle
        
            MsgBox "Invalid Department Title!" & vbNewLine & "The Department Title that you have entered is already existed. Enter another Department Title.", vbExclamation
            HLTxt txtDepartmentTitle
            
        Case Else
            MsgBox "Unknown Error", vbExclamation
            CatchError "frmAddDepartment", "SaveNewDepartment", "Unknown result in Add New Department"
    End Select
    
    
End Function

Private Sub cmdGetCollege_Click()
Dim sCollegeTitle As String
    Dim sCollegeID As String
    
    sCollegeID = frmPickCollege.GetItem(txtCollege, sCollegeTitle)
    If sCollegeID <> "" Then
        txtCollege.Text = sCollegeTitle
        CollegeID = sCollegeID
    End If
End Sub

Private Sub cmdSave_Click()
If mFormState = "ADD" Then
    SaveNewDepartment
Else
    UpdateData
End If
End Sub

Private Function UpdateData()
    
    If Not CheckTextBox(txtDepartmentTitle, "Enter Department Title." & vbNewLine & " This field is required") Then
        Exit Function
    End If
    
    If Not CheckTextBox(txtCollege, "Enter College Title." & vbNewLine & " This field is required") Then
        Exit Function
    End If

    Dim newDepartment As tDepartment
    Dim EditResult As Integer
    
    newDepartment.DepartmentID = txtDepartmentID.Text
    newDepartment.DepartmentTitle = txtDepartmentTitle.Text
    newDepartment.CollegeID = CollegeID
    
    Select Case EditDepartment(newDepartment)
        Case Success
        
            MsgBox "Department Information was successfully edited", vbInformation
            
            RecordEdited = True
        
            Unload Me
            
        Case DuplicateTitle
            MsgBox "The Department Title that you have enetered was already existed." & vbNewLine & " Enter another Duplicate Title", vbExclamation
            HLTxt txtDepartmentTitle

        Case Else
            MsgBox "UNKNOWN: Editing Department", vbCritical
    End Select
End Function


