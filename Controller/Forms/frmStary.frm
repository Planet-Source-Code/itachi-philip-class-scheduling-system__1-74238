VERSION 5.00
Begin VB.Form frmStart 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Welcome to College Enrollment System"
   ClientHeight    =   3705
   ClientLeft      =   2370
   ClientTop       =   2400
   ClientWidth     =   8385
   ControlBox      =   0   'False
   Icon            =   "frmStary.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3705
   ScaleWidth      =   8385
   StartUpPosition =   2  'CenterScreen
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.CommandButton cmdOK 
      Caption         =   "Continue >>"
      Height          =   375
      Left            =   6120
      TabIndex        =   0
      Top             =   3240
      Width           =   2175
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Developed By: "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      TabIndex        =   9
      Top             =   3360
      Width           =   1845
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "NIT Software Lab"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   285
      Left            =   2040
      TabIndex        =   8
      Top             =   3360
      Width           =   2325
   End
   Begin VB.Label lblDay 
      BackStyle       =   0  'Transparent
      Caption         =   "20"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   5520
      TabIndex        =   7
      Top             =   2400
      Width           =   495
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Run times left:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   3600
      TabIndex        =   6
      Top             =   2400
      Width           =   2175
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "philipgaray2@gmail.com"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   5400
      TabIndex        =   5
      Top             =   1800
      Width           =   2415
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "To buy College Enrollment System send us your purchase request on:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3600
      TabIndex        =   4
      Top             =   1560
      Width           =   4695
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "To continue working trial version of College Enrollment System click ""Continue""."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3600
      TabIndex        =   3
      Top             =   960
      Width           =   4695
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "This program is not free. It is an evaluation version of copyrighted software."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3600
      TabIndex        =   2
      Top             =   360
      Width           =   4455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Reminder"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   3600
      TabIndex        =   1
      Top             =   0
      Width           =   1575
   End
   Begin VB.Image Image1 
      Height          =   3615
      Left            =   0
      Picture         =   "frmStary.frx":492A
      Stretch         =   -1  'True
      Top             =   0
      Width           =   3375
   End
End
Attribute VB_Name = "frmStart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const ExpireNo = "10"
Dim INSTALLDATE As Date
Private Sub cmdOK_Click()
    If TrialRemains > 0 Then
        AllowedOpen
        
        Unload Me
        
        mdiController.StatusBar1.Panels(4).Text = CurrentSchoolYear.SchoolYearTitle
        mdiController.StatusBar1.Panels(5).Text = CurrentSemester.Semester
        mdiController.Show
    Else
        End
    End If
End Sub
Public Sub SetTrial()
Dim vRS As New ADODB.Recordset
    If ConnectRS(con, vRS, "Select * FROM tblExpiration") = True Then
            vRS.AddNew
            vRS.Fields("InstallationDate") = Date
            vRS.Fields("LASTRUNDATE") = Date
            vRS.Fields("CURRENTDATE") = Date
            vRS.Fields("ExpirationAmount") = ExpireNo
            vRS.Fields("AllowOpen") = "1"
            vRS.Fields("Remaining") = ExpireNo
            vRS.Update
    End If
    Me.Show 1
End Sub
Private Function InstallationDate() As Date
    Dim vRS As New ADODB.Recordset
    If ConnectRS(con, vRS, "Select * FROM tblExpiration") = True Then
            If vRS.RecordCount > 0 Then
                vRS.MoveFirst
                InstallationDate = vRS.Fields("InstallationDate")
            End If
    End If
End Function
Private Function TrialRemains() As Integer
    Dim vRS As New ADODB.Recordset
    If ConnectRS(con, vRS, "Select * FROM tblExpiration") = True Then
            If vRS.RecordCount > 0 Then
                vRS.MoveFirst
                TrialRemains = vRS.Fields("Remaining")
            End If
    End If
End Function

Private Function RunTime() As Integer
    Dim vRS As New ADODB.Recordset
    If ConnectRS(con, vRS, "Select * FROM tblExpiration") = True Then
            If vRS.RecordCount > 0 Then
                vRS.MoveFirst
                RunTime = vRS.Fields("ExpirationAmount")
            End If
    End If
End Function

Private Function PreviousDate() As Date
    Dim vRS As New ADODB.Recordset
    If ConnectRS(con, vRS, "Select * FROM tblExpiration") = True Then
            If vRS.RecordCount > 0 Then
                vRS.MoveFirst
                PreviousDate = vRS.Fields("LASTRUNDATE")
            End If
    End If
End Function

Public Sub AllowedOpen()
Dim vRS As New ADODB.Recordset

    If ConnectRS(con, vRS, "Select * FROM tblExpiration WHERE Remaining > 1") = True Then
        If vRS.RecordCount > 0 Then
            vRS.MoveFirst
            vRS.Fields("LASTRUNDATE") = PreviousDate
            vRS.Fields("CURRENTDATE") = Date
            vRS.Fields("AllowOpen") = "1"
            vRS.Fields("Remaining") = vRS.Fields("Remaining") - 1
            vRS.Update
        End If
    End If
  
End Sub

Private Sub Form_Load()
  SetupSettings
End Sub

Private Sub SetupSettings()
    InstallationDate
    TrialRemains
    
    lblDay.Caption = TrialRemains
    
    RunTime
    PreviousDate
End Sub
