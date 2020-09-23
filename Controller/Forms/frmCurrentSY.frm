VERSION 5.00
Begin VB.Form frmCurrentSY 
   BackColor       =   &H00D8E9EC&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Current School Year"
   ClientHeight    =   1560
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   4680
   Icon            =   "frmCurrentSY.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   921.7
   ScaleMode       =   0  'User
   ScaleWidth      =   4394.266
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cboSemester 
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
      ItemData        =   "frmCurrentSY.frx":492A
      Left            =   1320
      List            =   "frmCurrentSY.frx":4937
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   480
      Width           =   2715
   End
   Begin VB.CommandButton cmdGetSchoolYear 
      BackColor       =   &H00D8E9EC&
      Height          =   315
      Left            =   4080
      Picture         =   "frmCurrentSY.frx":495F
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   120
      Width           =   345
   End
   Begin VB.TextBox txtUserName 
      Height          =   345
      Left            =   1320
      TabIndex        =   1
      Top             =   120
      Width           =   2715
   End
   Begin VB.CommandButton cmdSelect 
      BackColor       =   &H00D8E9EC&
      Caption         =   "Select"
      Default         =   -1  'True
      Height          =   390
      Left            =   2160
      TabIndex        =   2
      Top             =   1020
      Width           =   1140
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00D8E9EC&
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   390
      Left            =   3360
      TabIndex        =   3
      Top             =   1020
      Width           =   1140
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "&Semester"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   1
      Left            =   120
      TabIndex        =   6
      Top             =   480
      Width           =   1080
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "&School Year:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   0
      Left            =   105
      TabIndex        =   0
      Top             =   150
      Width           =   1290
   End
End
Attribute VB_Name = "frmCurrentSY"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub setSchoolYear()
    
    If SchoolYearRecordExisted <> Success Then
        MsgBox "There are no records yet in School Year.", vbInformation
        Unload Me
        Exit Sub
    End If

    If Len(CurrentSchoolYear.SchoolYearTitle) > 0 Then
        Me.txtUserName.Text = CurrentSchoolYear.SchoolYearTitle
    End If

    Me.Show vbModal
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdGetSchoolYear_Click()
    Dim sSchoolYearTitle As String
    
    sSchoolYearTitle = frmPickSchoolYear.GetItem(txtUserName)
    
    If sSchoolYearTitle <> "" Then
        txtUserName.Text = sSchoolYearTitle
    End If
End Sub


Private Sub cmdSelect_Click()
    If SchoolYearExistByTitle(txtUserName.Text) = Success Then
        SaveActiveSchoolYear txtUserName.Text
        SaveActiveSemester cboSemester.Text
        
        CurrentSchoolYear.SchoolYearTitle = txtUserName.Text
        CurrentSemester.Semester = cboSemester.Text
        
        mdiController.StatusBar1.Panels(4).Text = CurrentSchoolYear.SchoolYearTitle
        mdiController.StatusBar1.Panels(5).Text = CurrentSemester.Semester
        Unload Me
    Else
        MsgBox "The selected School Year does not exist in record!" & vbNewLine & _
        "Please enter valid School Year.", vbExclamation
    End If
End Sub
Private Sub txtSchoolYear_Change()
    If Len(txtSchoolYear) < 1 Then
        cmdSelect.Enabled = False
    Else
        cmdSelect.Enabled = True
    End If
End Sub

