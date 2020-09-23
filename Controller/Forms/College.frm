VERSION 5.00
Begin VB.Form frmCollege 
   BackColor       =   &H00D8E9EC&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "College"
   ClientHeight    =   2415
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   6720
   Icon            =   "College.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2415
   ScaleWidth      =   6720
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00D8E9EC&
      ForeColor       =   &H80000008&
      Height          =   1815
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   6495
      Begin VB.TextBox txtCollegeID 
         BackColor       =   &H00C0FFFF&
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
         MaxLength       =   255
         TabIndex        =   5
         Top             =   480
         Width           =   6225
      End
      Begin VB.TextBox txtCollege 
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
         MaxLength       =   255
         TabIndex        =   3
         Top             =   1080
         Width           =   6225
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "College Reference #:"
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
         TabIndex        =   6
         Top             =   270
         Width           =   1815
      End
      Begin VB.Label Label1 
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
         TabIndex        =   4
         Top             =   840
         Width           =   1095
      End
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00D8E9EC&
      Caption         =   "Save"
      Height          =   375
      Left            =   3960
      TabIndex        =   1
      Top             =   1920
      Width           =   1335
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00D8E9EC&
      Caption         =   "Cancel"
      Height          =   375
      Left            =   5400
      TabIndex        =   0
      Top             =   1920
      Width           =   1215
   End
End
Attribute VB_Name = "frmCollege"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim RecordAdded As Boolean
Dim CollegeID As String
Dim curCollege As tCollege
Public mFormState As String

Dim RecordEdited As Boolean

Public Function ShowEdit(sCollegeID As String) As Boolean
        
    
        If GetCollegeByID(sCollegeID, curCollege) <> Success Then
            MsgBox "Unable to continue editing COLLGE Information: COLLEGE ID not found!", vbCritical
            Exit Function
        End If

    CollegeID = curCollege.CollegeID
    txtCollege.Text = curCollege.CollegeTitle
    txtCollegeID.Text = CollegeID
    
    
    mFormState = "EDIT"
    
    Me.Show vbModal

    ShowEdit = RecordEdited
    
    
End Function

Public Function ShowProperties(sCollege As String) As Boolean
        
        If GetCollegeByTitle(sCollege, curCollege) <> Success Then
            MsgBox "Unable to continue editing COLLGE Information: COLLEGE ID not found!", vbCritical
            Exit Function
        End If

    CollegeID = curCollege.CollegeID
    txtCollege.Text = curCollege.CollegeTitle
    txtCollegeID.Text = CollegeID
    
    mFormState = "EDIT"
    
    Me.Show vbModal

    ShowProperties = RecordEdited
    
    
End Function

Public Function ShowForm() As Boolean
    
    Dim sNewID As String
    
    If GetNewCollegeID(sNewID) = Failed Then
        CatchError "College", "ShowForm()", "GetNewCollegeID(sNewID) = Failed"
        Exit Function
    End If

    CollegeID = sNewID
    txtCollegeID.Text = CollegeID
    
    mFormState = "ADD"

    Me.Show vbModal
    'return
    ShowForm = RecordAdded
    
End Function


Private Sub cmdCancel_Click()
    Unload Me
End Sub


Private Function SaveNewDepartment()
    
    If Not CheckTextBox(txtCollege, "Please enter COLLEGE NAME") Then
        Exit Function
    End If
    
    Dim newCollege As tCollege
    
    newCollege.CollegeID = txtCollegeID.Text
    newCollege.CollegeTitle = txtCollege.Text
    
    Select Case AddCollege(newCollege)
        Case TranDBResult.Success
            MsgBox "New COLLEGE succesfully added", vbInformation
            RecordAdded = True
            Unload Me
            
        Case TranDBResult.DuplicateID
            MsgBox "Invalid COLLEGE ID!" & vbNewLine & "The COLLEGE ID that you have entered is already existed. Enter another COLLEGE ID.", vbExclamation
            
        Case TranDBResult.DuplicateTitle
            MsgBox "Invalid COLLEGE Name!" & vbNewLine & "The COLLEGE Name that you have entered is already existed. Enter another COLLEGE Name.", vbExclamation
            HLTxt txtCollege
            
        Case Else
            MsgBox "Unknown Error", vbExclamation
            CatchError "frmCollege", "SaveNewCollege", "Unknown result in Add New College"
    End Select
    
    
End Function

Private Sub cmdSave_Click()
If mFormState = "ADD" Then
    SaveNewDepartment
Else
    UpdateData
End If
End Sub

Private Function UpdateData()
    
    If Not CheckTextBox(txtCollege, "Enter COLLEGE Name." & vbNewLine & " This field is required") Then
        Exit Function
    End If

    Dim newCollege As tCollege
    Dim EditResult As Integer
    
    newCollege.CollegeID = txtCollegeID.Text
    newCollege.CollegeTitle = txtCollege.Text

    Select Case EditCollege(newCollege)
        Case Success
        
            MsgBox "College Information was successfully edited", vbInformation
            
            RecordEdited = True
        
            Unload Me
            
        Case DuplicateTitle
            MsgBox "The College name that you have enetered was already existed." & vbNewLine & " Enter another Duplicate Title", vbExclamation
            HLTxt txtCollege

        Case Else
            MsgBox "UNKNOWN: Editing College", vbCritical
    End Select
End Function


