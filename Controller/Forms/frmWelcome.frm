VERSION 5.00
Begin VB.Form frmWelcome 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Welcome"
   ClientHeight    =   5235
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   8190
   Icon            =   "frmWelcome.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5235
   ScaleWidth      =   8190
   WindowState     =   2  'Maximized
   Begin VB.CheckBox Check1 
      Caption         =   "Show this message again"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   4800
      Width           =   3015
   End
   Begin VB.Image Image1 
      Height          =   3015
      Left            =   120
      Top             =   600
      Width           =   2175
   End
End
Attribute VB_Name = "frmWelcome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub bgPic_Click()

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If UnloadMode = 0 Then
    Cancel = 1
    Beep
End If
End Sub

Private Sub Form_Resize()
    bgPic.Height = ScaleHeight
    bgPic.Width = ScaleWidth
End Sub
