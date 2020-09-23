VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About College Enrollment System"
   ClientHeight    =   3390
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   5580
   ClipControls    =   0   'False
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2339.838
   ScaleMode       =   0  'User
   ScaleWidth      =   5239.907
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Implemented At"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2760
      TabIndex        =   3
      Top             =   720
      Width           =   2085
   End
   Begin VB.Label Label6 
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
      Left            =   2520
      TabIndex        =   2
      Top             =   3000
      Width           =   2325
   End
   Begin VB.Label Label4 
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
      Left            =   600
      TabIndex        =   1
      Top             =   3000
      Width           =   1845
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "College Enrollment System"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   405
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5205
   End
   Begin VB.Image Image1 
      Height          =   3255
      Left            =   0
      Picture         =   "frmAbout.frx":F172
      Stretch         =   -1  'True
      Top             =   0
      Width           =   2655
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Click()
     Unload Me
End Sub

Private Sub Form_LostFocus()
    Unload Me
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then: Unload Me
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
     If Button = 2 Then: Unload Me
End Sub

