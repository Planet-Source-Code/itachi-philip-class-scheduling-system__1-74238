VERSION 5.00
Begin VB.Form frmSchoolYearAE 
   BackColor       =   &H00D8E9EC&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "School Year"
   ClientHeight    =   2295
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   4965
   Icon            =   "frmSchoolYear.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2295
   ScaleWidth      =   4965
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   1800
      TabIndex        =   7
      Top             =   1800
      Width           =   1455
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3360
      TabIndex        =   6
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00D8E9EC&
      Height          =   1695
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   4695
      Begin VB.TextBox txtTo 
         BackColor       =   &H00D8E9EC&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2880
         MaxLength       =   4
         TabIndex        =   3
         Top             =   1020
         Width           =   1095
      End
      Begin VB.TextBox txtFrom 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1770
         MaxLength       =   4
         TabIndex        =   2
         Top             =   1020
         Width           =   1095
      End
      Begin VB.TextBox txtSchoolYear 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1770
         Locked          =   -1  'True
         MaxLength       =   9
         TabIndex        =   1
         Top             =   360
         Width           =   2205
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "From"
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
         Left            =   360
         TabIndex        =   5
         Top             =   1080
         Width           =   450
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "School Year"
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
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Width           =   1020
      End
   End
End
Attribute VB_Name = "frmSchoolYearAE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RecordSaved As Boolean

Public Function ShowForm(Optional newSchoolYearTitleFrom As String = "") As Boolean
    
    If newSchoolYearTitleFrom <> "" Then
        txtFrom = newSchoolYearTitleFrom
    End If
    
    'show form
    Me.Show vbModal
    
    ShowForm = RecordSaved
End Function


Private Sub cmdSave_Click()
'check if filled
    If Len(txtSchoolYear.Text) < 1 Then
        'temp
        MsgBox "Fill 'From Year' Text Field First", vbInformation
        Exit Sub
    End If
    
    'save
    Dim newSchoolYear As tSchoolYear
    
    'set object
    newSchoolYear.SchoolYearTitle = txtSchoolYear.Text
    
    Select Case AddSchoolYear(newSchoolYear)
        Case TranDBResult.Success
        
            'ADD success
            '------------------------------------------------------
                        
            'temp
            MsgBox "School Year created.", vbInformation
            
            'return true
            RecordSaved = True
            
            'close this form
            
            Unload Me
        
        Case TranDBResult.DuplicateTitle
            MsgBox "The Entry you have entered is already existed. Enter another entry.", vbExclamation
            HLTxt txtFrom

            
        Case Else
            'temp
            MsgBox "Error: Creating School Year", vbCritical
    End Select
End Sub

Private Sub txtFrom_Change()
     If Len(txtFrom) = 4 And Val(txtFrom) > 1000 Then
            'auto fill
            txtTo.Text = Val(txtFrom) + 1
            txtSchoolYear.Text = txtFrom.Text & "-" & txtTo.Text
    Else
        txtTo.Text = ""
        txtSchoolYear.Text = ""
    End If
End Sub

Private Sub txtFrom_KeyPress(KeyAscii As Integer)
If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8 Or KeyAscii = 45) Then KeyAscii = 0
End Sub

Private Sub txtSchoolYear_GotFocus()
 txtFrom.SetFocus
End Sub
Private Sub txtTo_GotFocus()
    txtFrom.SetFocus
End Sub

