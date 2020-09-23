VERSION 5.00
Begin VB.Form frmLogin 
   BackColor       =   &H00D8E9EC&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Login"
   ClientHeight    =   3090
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   3720
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1825.674
   ScaleMode       =   0  'User
   ScaleWidth      =   3492.879
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   390
      Left            =   2115
      TabIndex        =   4
      Top             =   2565
      Width           =   1140
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   390
      Left            =   510
      TabIndex        =   3
      Top             =   2565
      Width           =   1140
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00D8E9EC&
      Height          =   2295
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   3495
      Begin VB.TextBox txtPassword 
         Height          =   345
         IMEMode         =   3  'DISABLE
         Left            =   240
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   1200
         Width           =   3045
      End
      Begin VB.TextBox txtUserName 
         Height          =   345
         Left            =   240
         TabIndex        =   0
         Top             =   480
         Width           =   3045
      End
      Begin VB.Label lblLabels 
         BackStyle       =   0  'Transparent
         Caption         =   "&Password:"
         Height          =   270
         Index           =   1
         Left            =   240
         TabIndex        =   6
         Top             =   960
         Width           =   1080
      End
      Begin VB.Label lblLabels 
         BackStyle       =   0  'Transparent
         Caption         =   "&User Name:"
         Height          =   270
         Index           =   0
         Left            =   240
         TabIndex        =   5
         Top             =   240
         Width           =   1080
      End
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rs As New ADODB.Recordset
Dim passattemp As Integer
Dim CurrentUser As CurrentUser


Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If txtUserName.Text = "" Then txtUserName.SetFocus: Exit Sub
    If txtPassword.Text = "" Then txtPassword.SetFocus: Exit Sub
    
    With rs
            Dim strSql As String
            strSql = "Select * From tblUser Where UserName='" & txtUserName.Text & "'"
            .Open strSql, con, adOpenStatic, adLockOptimistic
            
            If .RecordCount >= 1 Then
                If .Fields("Password") = txtPassword.Text Then
                    CurrentUser.Fullname = .Fields("Fullname").Value
                    
                    Call Enableds
                    
                    Unload Me
                    
                   Else
                    passattemp = passattemp + 1
                        If passattemp = 3 Then
                            'sndPlaySound App.Path & "\Mad.wav", 0 + 17
                            MsgBox "You are not an authorized user", vbInformation + vbCritical, "Log In Error"
                            End
                        Else
                            MsgBox "Password incorrect. Please check the CAPS LOCK" & vbCrLf & " Attempt left " & 3 - passattemp & "", vbExclamation, "Log In Error"
                            txtPassword.Text = ""
                            txtPassword.SetFocus
                        End If
                End If
            Else
                MsgBox "This user does not exist", vbCritical, "Log In Error"
                txtUserName.Text = ""
                txtUserName.SetFocus
            End If
            .Close
        End With
        
Set rs = Nothing

Exit Sub

err:
MsgBox err.Description, vbCritical
Set rs = Nothing
End Sub


Private Sub Enableds()
With mdiController
    .mnulogin.Caption = "Logout..."
    .mnuchangepass.Enabled = True
    .mnubrowse.Enabled = True
    
    .mnuclassusage.Enabled = True
    .mnuCurriculum.Enabled = True
    .mnuEnrollment.Enabled = True
    .mnuscholars.Enabled = True
    .mnuSetActive.Enabled = True
    
    .mnusettings.Enabled = True
    .mnuDbase.Enabled = True
    .mnuSY.Enabled = True
    
    .Tool.Enabled = True
    
    .StatusBar1.Panels(3).Text = CurrentUser.Fullname
End With
End Sub
