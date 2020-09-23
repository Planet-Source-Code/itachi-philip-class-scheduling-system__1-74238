VERSION 5.00
Begin VB.Form frmScholarship 
   BackColor       =   &H00D8E9EC&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Scholarship"
   ClientHeight    =   3375
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5880
   Icon            =   "frmScholarship.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3375
   ScaleWidth      =   5880
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00D8E9EC&
      Height          =   2775
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   5655
      Begin VB.TextBox txtScholarshipID 
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
         TabIndex        =   9
         Top             =   360
         Width           =   2025
      End
      Begin VB.TextBox txtScholarship 
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
         Top             =   960
         Width           =   5385
      End
      Begin VB.TextBox txtAllocation 
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
         TabIndex        =   4
         Top             =   1560
         Width           =   2025
      End
      Begin VB.TextBox txtBenefactor 
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
         Top             =   2160
         Width           =   5385
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Scholarship Referrence #"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   10
         Top             =   120
         Width           =   2055
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Allocation"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   8
         Top             =   1320
         Width           =   780
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Benefactor"
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
         Top             =   1920
         Width           =   1695
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Scholarship"
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
         Top             =   720
         Width           =   1095
      End
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4440
      TabIndex        =   1
      Top             =   2880
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   3120
      TabIndex        =   0
      Top             =   2880
      Width           =   1215
   End
End
Attribute VB_Name = "frmScholarship"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Public FormState As String

Public Function GetNewScholarID(ByRef sNewScholarID As String) As TranDBResult
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    Dim NewDNumber As Integer
    
    GetNewScholarID = Failed
    
    sSQL = "SELECT 'Scholar-' & String$(2-Len(Count(*)+1),'0') & Count(*)+1 AS NewID" & _
            " FROM tblScholarship;"
            
    If ConnectRS(con, vRS, sSQL) = True Then
        
        sNewScholarID = vRS.Fields("NewID").Value
        
        While DepartmentExistByID(sNewScholarID) = Success
            NewDNumber = Val(Right(sNewScholarID, 2)) + 1
            sNewScholarID = "D-" & String(2 - Len(NewDNumber), "0") & NewDNumber
        Wend
        GetNewScholarID = Success
    Else
        GetNewScholarID = Failed
    End If
    Set vRS = Nothing

End Function

Public Function ScholarExistByID(sScholarID As String) As TranDBResult
    
    Dim vRS As New ADODB.Recordset
        
    If ConnectRS(con, vRS, "SELECT * From tblDepartment WHERE (((tblDepartment.DepartmentID)='" & sScholarID & "'));") Then
        If vRS.RecordCount > 0 Then
            ScholarExistByID = Success
        Else
            ScholarExistByID = Failed
        End If
    Else
        ScholarExistByID = Failed
       
    End If
    
    'release
    Set vRS = Nothing
End Function

Public Function ShowForm()
    Dim sNewID As String
    
    FormState = "ADD"
    
    If GetNewScholarID(sNewID) = Failed Then
        CatchError "Scholarship", "ShowForm()", "GetNewSholarID(sNewID) = Failed"
        Exit Function
    End If
        
    txtScholarshipID.Text = sNewID
    
    Me.Show 1
End Function
Public Function ShowEdit(sScholarshipID As String)
    Dim rs As New ADODB.Recordset

     FormState = "EDIT"
    
    If ScholarExistByID(sScholarshipID) = Failed Then
        CatchError "Scholarship", "ShowEdit()", "ScholarExistByID(sScholarshipID) = Failed"
        Exit Function
    End If
    
    If ConnectRS(con, rs, "SELECT * FROM tblScholarship WHERE ScholarshipID ='" & sScholarshipID & "'") = True Then
        If rs.RecordCount > 0 Then
            txtScholarshipID.Text = sScholarshipID
            txtScholarshipID.Text = rs.Fields("ScholarshipID")
            txtScholarship.Text = rs.Fields("Scholarship")
            txtAllocation.Text = rs.Fields("Allocation")
            txtBenefactor.Text = rs.Fields("Benefactor")
        End If
    End If
    
    Me.Show 1
End Function


Private Sub SaveData()
    Dim rs As New ADODB.Recordset
    
    If ConnectRS(con, rs, "Select * FROM tblScholarship") = True Then
        rs.AddNew
        rs.Fields("ScholarshipID") = txtScholarshipID.Text
        rs.Fields("Scholarship") = txtScholarship.Text
        rs.Fields("Allocation") = txtAllocation.Text
        rs.Fields("Benefactor") = txtBenefactor.Text
        rs.Fields("CreatedBy") = CurrentUser.Fullname
        rs.Fields("CreationDate") = Now
        rs.Update
        
        MsgBox "Record Successfully saved...", vbInformation
        Unload Me
    End If
End Sub

Private Sub OKButton_Click()
    Select Case FormState
        Case "ADD"
            SaveData
        Case Else
    End Select
End Sub
