VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.MDIForm mdiController 
   BackColor       =   &H00D8E9EC&
   Caption         =   "Controller"
   ClientHeight    =   8220
   ClientLeft      =   165
   ClientTop       =   825
   ClientWidth     =   8880
   Icon            =   "mdiController.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picBackdrop 
      Align           =   1  'Align Top
      AutoRedraw      =   -1  'True
      Height          =   4335
      Left            =   0
      ScaleHeight     =   4275
      ScaleWidth      =   8820
      TabIndex        =   2
      Top             =   570
      Visible         =   0   'False
      Width           =   8880
      Begin VB.PictureBox picOriginal 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   9090
         Left            =   240
         Picture         =   "mdiController.frx":492A
         ScaleHeight     =   9090
         ScaleWidth      =   15360
         TabIndex        =   4
         Top             =   240
         Width           =   15360
      End
      Begin VB.PictureBox picStretched 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   7260
         Left            =   2160
         ScaleHeight     =   7260
         ScaleWidth      =   4095
         TabIndex        =   3
         Top             =   360
         Width           =   4095
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3360
      Top             =   5520
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   56
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiController.frx":10B8A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiController.frx":11464
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiController.frx":12DF6
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiController.frx":14788
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiController.frx":1611A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiController.frx":17AAC
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiController.frx":1943E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiController.frx":1ADD0
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiController.frx":1C762
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiController.frx":1E0F4
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiController.frx":1FA88
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiController.frx":20ADA
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiController.frx":213BA
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiController.frx":2240C
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiController.frx":2345E
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiController.frx":244B0
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiController.frx":25502
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiController.frx":26554
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiController.frx":26E30
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiController.frx":2770A
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiController.frx":27FE4
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiController.frx":288BE
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiController.frx":29198
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiController.frx":29A72
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiController.frx":2A34C
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiController.frx":2AC26
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiController.frx":2B500
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiController.frx":2BDDA
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiController.frx":2C6B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiController.frx":2CF8E
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiController.frx":2D868
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiController.frx":2E142
            Key             =   ""
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiController.frx":2EA1C
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiController.frx":2F2F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiController.frx":2FBD0
            Key             =   ""
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiController.frx":304AA
            Key             =   ""
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiController.frx":30D84
            Key             =   ""
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiController.frx":3165E
            Key             =   ""
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiController.frx":31F38
            Key             =   ""
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiController.frx":32812
            Key             =   ""
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiController.frx":330EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiController.frx":339C6
            Key             =   ""
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiController.frx":35E18
            Key             =   ""
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiController.frx":377AA
            Key             =   ""
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiController.frx":3913C
            Key             =   ""
         EndProperty
         BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiController.frx":3AACE
            Key             =   ""
         EndProperty
         BeginProperty ListImage47 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiController.frx":3BB20
            Key             =   ""
         EndProperty
         BeginProperty ListImage48 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiController.frx":3C3FC
            Key             =   ""
         EndProperty
         BeginProperty ListImage49 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiController.frx":3D44E
            Key             =   ""
         EndProperty
         BeginProperty ListImage50 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiController.frx":3E4A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage51 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiController.frx":3ED84
            Key             =   ""
         EndProperty
         BeginProperty ListImage52 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiController.frx":3FDD6
            Key             =   ""
         EndProperty
         BeginProperty ListImage53 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiController.frx":406B2
            Key             =   ""
         EndProperty
         BeginProperty ListImage54 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiController.frx":42046
            Key             =   ""
         EndProperty
         BeginProperty ListImage55 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiController.frx":439DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage56 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiController.frx":442B6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Tool 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8880
      _ExtentX        =   15663
      _ExtentY        =   1005
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   8
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Browse"
            Object.ToolTipText     =   "Browse Student [F3]"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Classroom"
            Object.ToolTipText     =   "Classroom Usage [F5]"
            ImageIndex      =   54
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Curriculum"
            Object.ToolTipText     =   "Curriculum [F6]"
            ImageIndex      =   55
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Enrollment"
            Object.ToolTipText     =   "Enrollment [F7]"
            ImageIndex      =   47
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Settings"
            Object.ToolTipText     =   "Settings [F8]"
            ImageIndex      =   9
         EndProperty
      EndProperty
      Enabled         =   0   'False
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   1
      Top             =   7920
      Width           =   8880
      _ExtentX        =   15663
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   13
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   442
            MinWidth        =   442
            Picture         =   "mdiController.frx":448E0
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   1587
            MinWidth        =   1587
            Text            =   "User Name:"
            TextSave        =   "User Name:"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   2646
            MinWidth        =   2646
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   529
            MinWidth        =   529
            Picture         =   "mdiController.frx":44C7C
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "Today:"
            TextSave        =   "Today:"
         EndProperty
         BeginProperty Panel8 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "17/12/2011"
         EndProperty
         BeginProperty Panel9 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "10:06 PM"
         EndProperty
         BeginProperty Panel10 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   3
            Enabled         =   0   'False
            Object.Width           =   882
            MinWidth        =   882
            TextSave        =   "INS"
         EndProperty
         BeginProperty Panel11 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Enabled         =   0   'False
            Object.Width           =   882
            MinWidth        =   882
            TextSave        =   "NUM"
         EndProperty
         BeginProperty Panel12 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Enabled         =   0   'False
            Object.Width           =   882
            MinWidth        =   882
            TextSave        =   "CAPS"
         EndProperty
         BeginProperty Panel13 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   4
            Enabled         =   0   'False
            Object.Width           =   882
            MinWidth        =   882
            TextSave        =   "SCRL"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu mnuaction 
      Caption         =   "&Action"
      Begin VB.Menu mnulogin 
         Caption         =   "Login..."
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuchangepass 
         Caption         =   "Change password..."
         Enabled         =   0   'False
      End
      Begin VB.Menu mnus1 
         Caption         =   "-"
      End
      Begin VB.Menu mnubrowse 
         Caption         =   "Browse Student..."
         Enabled         =   0   'False
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnuSetActive 
         Caption         =   "Set active schoolyear"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnus2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuexit 
         Caption         =   "Exit Module"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "View"
      Begin VB.Menu mnuclassusage 
         Caption         =   "Class Room usage..."
         Enabled         =   0   'False
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnusep4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCurriculum 
         Caption         =   "Curriculum..."
         Enabled         =   0   'False
         Shortcut        =   {F6}
      End
      Begin VB.Menu mnuEnrollment 
         Caption         =   "Enrollment..."
         Enabled         =   0   'False
         Shortcut        =   {F7}
      End
      Begin VB.Menu mnuscholars 
         Caption         =   "Scholars"
         Enabled         =   0   'False
         Shortcut        =   {F8}
      End
   End
   Begin VB.Menu mnutools 
      Caption         =   "&Tools"
      Begin VB.Menu mnusettings 
         Caption         =   "Settings..."
         Enabled         =   0   'False
         Shortcut        =   {F9}
      End
      Begin VB.Menu mnusee 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDbase 
         Caption         =   "Database Settings..."
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuSY 
         Caption         =   "School Year Settings..."
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuabout 
         Caption         =   "About..."
      End
   End
   Begin VB.Menu popSubject 
      Caption         =   "Subject"
      Visible         =   0   'False
      Begin VB.Menu mnuNewSubject 
         Caption         =   "New Subject..."
      End
      Begin VB.Menu mnuNewSection 
         Caption         =   "New Section..."
      End
      Begin VB.Menu mnuProperties 
         Caption         =   "Properties..."
      End
      Begin VB.Menu mnusep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRefresh 
         Caption         =   "Refresh"
      End
   End
   Begin VB.Menu popSection 
      Caption         =   "Section"
      Visible         =   0   'False
      Begin VB.Menu mnuSectionProperties 
         Caption         =   "Properties..."
      End
      Begin VB.Menu mnuDeleteSection 
         Caption         =   "Delete section"
      End
      Begin VB.Menu mnusep5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPrintClasslist 
         Caption         =   "Print classlist"
      End
      Begin VB.Menu mnusep6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRefreshSection 
         Caption         =   "Refresh"
      End
   End
   Begin VB.Menu popFaculty 
      Caption         =   "Faculty"
      Visible         =   0   'False
      Begin VB.Menu mnuviewfacultyinfo 
         Caption         =   "View Faculty info..."
      End
      Begin VB.Menu mnuprintteachingload 
         Caption         =   "Print teaching load..."
      End
      Begin VB.Menu mnusep8 
         Caption         =   "-"
      End
      Begin VB.Menu mnufacultyRefresh 
         Caption         =   "Refresh"
      End
   End
   Begin VB.Menu popCurriculum 
      Caption         =   "Curriculum"
      Visible         =   0   'False
      Begin VB.Menu mnuCurProperties 
         Caption         =   "Properties..."
      End
      Begin VB.Menu mnusep9 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDeleteCur 
         Caption         =   "Delete"
      End
      Begin VB.Menu mnusep10 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRefreshCur 
         Caption         =   "Refresh"
      End
   End
   Begin VB.Menu popSubjectCur 
      Caption         =   "Subject Curriculum"
      Visible         =   0   'False
      Begin VB.Menu mnuNewSubjectCur 
         Caption         =   "Add to Prospectus..."
      End
      Begin VB.Menu mnuSubjectPropCur 
         Caption         =   "Properties..."
      End
      Begin VB.Menu mnusep16 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRefreshSubjectCur 
         Caption         =   "Refresh"
      End
   End
   Begin VB.Menu popScholarship 
      Caption         =   "Scholarship"
      Visible         =   0   'False
      Begin VB.Menu mnuViewScholars 
         Caption         =   "View scholars"
      End
      Begin VB.Menu a 
         Caption         =   "-"
      End
      Begin VB.Menu mnuNewScholarship 
         Caption         =   "New scholarship..."
      End
      Begin VB.Menu mnueditScholarship 
         Caption         =   "Edit scholarship"
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "Delete"
      End
      Begin VB.Menu b 
         Caption         =   "-"
      End
      Begin VB.Menu mnuprintscholarmasterlist 
         Caption         =   "Print masterlist..."
      End
      Begin VB.Menu c 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRefreshScholarship 
         Caption         =   "Refresh"
      End
   End
   Begin VB.Menu popScholar 
      Caption         =   "Scholar"
      Visible         =   0   'False
      Begin VB.Menu mnuViewDetails 
         Caption         =   "View student info..."
      End
      Begin VB.Menu e 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRefreshDetail 
         Caption         =   "Refresh"
      End
   End
End
Attribute VB_Name = "mdiController"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Private Declare Function InvalidateRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT, ByVal bErase As Long) As Long

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public CloseMe  As Boolean
Private Sub MDIForm_Activate()
    MDIForm_Resize
End Sub
Private Sub Disabled()
    With mdiController
        .mnulogin.Caption = "Login..."
        .mnuchangepass.Enabled = False
        .mnubrowse.Enabled = False
        
        .mnuclassusage.Enabled = False
        .mnuCurriculum.Enabled = False
        .mnuEnrollment.Enabled = False
        .mnuscholars.Enabled = False
        .mnuSetActive.Enabled = False
        
        .mnusettings.Enabled = False
        .mnuDbase.Enabled = False
        .mnuSY.Enabled = False
        
        .Tool.Enabled = False
        
        .StatusBar1.Panels(3).Text = ""
    End With
End Sub
Private Sub MDIForm_Load()
If CurrentSchoolYear.SchoolYearTitle = "0000" Or CurrentSemester.Semester = "0000" Then
    MsgBox "Please set a current school year and current semester"
    frmCurrentSY.setSchoolYear
End If

    StatusBar1.Panels(4).Text = CurrentSchoolYear.SchoolYearTitle
    StatusBar1.Panels(5).Text = CurrentSemester.Semester
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If MsgBox("Are you sure to Exit?", vbQuestion + vbYesNo) = vbYes Then
            End
    Else
        Cancel = 1
End If
End Sub

Private Sub MDIForm_Resize()
    Dim client_rect As RECT
    Dim client_hwnd As Long

    picStretched.Move 0, 0, _
        ScaleWidth, ScaleHeight
        
    picStretched.PaintPicture _
        picOriginal.Picture, _
        0, 0, _
        picStretched.ScaleWidth, _
        picStretched.ScaleHeight, _
        0, 0, _
        picOriginal.ScaleWidth, _
        picOriginal.ScaleHeight

    Picture = picStretched.Image

    client_hwnd = FindWindowEx(Me.hwnd, 0, "MDIClient", vbNullChar)
    GetClientRect client_hwnd, client_rect
    InvalidateRect client_hwnd, client_rect, 1
End Sub

Private Sub mnuabout_Click()
    frmAbout.Show 1
End Sub

Private Sub mnubrowse_Click()
    frmBrowse.Show 1
End Sub

Private Sub mnuclassusage_Click()
    frmUsage.ShowFormList
End Sub

Private Sub mnuCurProperties_Click()
On Error Resume Next
Call frmCurriculum.FolderClick(frmCurriculum.tvCurriculum.SelectedItem, Left(frmCurriculum.tvCurriculum.SelectedItem.Key, 4))
End Sub

Private Sub mnuCurriculum_Click()
    frmCurriculum.ShowFormList
End Sub

Private Sub mnuDbase_Click()
    frmODBCLogon.Show 1
End Sub

Private Sub mnuDeleteCur_Click()
    frmCurriculum.Subject_Delete
End Sub

Private Sub mnuDeleteSection_Click()
    frmEnrollment.ShowSectionOfferingDetail
End Sub

Private Sub mnuEnrollment_Click()
    frmEnrollment.ShowFormList
End Sub

Private Sub mnuexit_Click()
    Unload Me
End Sub

Private Sub mnulogin_Click()
Select Case mnulogin.Caption
    Case "Login..."
        frmLogin.Show 1
    Case Else
        Disabled
End Select
End Sub

Private Sub mnuNewScholarship_Click()
    frmScholarship.ShowForm
End Sub

Private Sub mnuNewSection_Click()
    frmScheduler.ShowForm frmEnrollment.lsvSubject.SelectedItem.Text, , , frmEnrollment.cboDepartment.Text
End Sub

Private Sub mnuNewSubject_Click()
    frmSubjectAE.ShowForm (frmEnrollment.cboDepartment.Text)
End Sub

Private Sub mnuNewSubjectCur_Click()
    frmProspectusAE.ShowForm frmCurriculum.lsvSubject.SelectedItem
End Sub

Private Sub mnuProperties_Click()
On Error GoTo err
    frmSubject.ShowForm frmEnrollment.lsvSubject.SelectedItem.SubItems(1)
Exit Sub
err:
MsgBox err.Description, vbCritical
End Sub

Private Sub mnuRefresh_Click()
On Error Resume Next
    frmEnrollment.LoadSubjectByDepartment (frmEnrollment.cboDepartment.Text)
End Sub

Private Sub mnuRefreshCur_Click()
    frmCurriculum.Refresh_Prospectus
End Sub

Private Sub mnuRefreshScholarship_Click()
    frmScholars.RefreshScholarship
End Sub

Private Sub mnuRefreshSection_Click()
    Call frmEnrollment.LoadSectionBySubject(frmEnrollment.txtSubjectID.Text, frmEnrollment.cmbSchoolYear.Text, frmEnrollment.cboSemester.Text)
End Sub

Private Sub mnuRefreshSubjectCur_Click()
    frmCurriculum.FormSubject_Refresh
End Sub

Private Sub mnuscholars_Click()
    frmScholars.Show
End Sub

Private Sub mnuSectionProperties_Click()
On Error GoTo err
    frmSection.ShowForm frmEnrollment.lsvSection.SelectedItem.SubItems(1)
Exit Sub
err:
MsgBox err.Description, vbCritical
End Sub

Private Sub mnuSetActive_Click()
    frmCurrentSY.Show 1
End Sub

Private Sub mnusettings_Click()
    frmSettings.ShowFormList
End Sub

Private Sub mnuSubjectPropCur_Click()
On Error GoTo err
    frmSubject.ShowForm frmCurriculum.lsvSubject.SelectedItem.SubItems(1)
Exit Sub
err:
MsgBox err.Description, vbCritical
End Sub

Private Sub mnuSY_Click()
    frmSchoolYearAE.ShowForm
End Sub

Private Sub mnuviewfacultyinfo_Click()
On Error GoTo err
    If frmEnrollment.lsvFaculty.ListItems.count < 1 Then Exit Sub

    If Len(GetLVKey(frmEnrollment.lsvFaculty.SelectedItem)) < 1 Then Exit Sub
        frmFaculty.ShowEdit GetLVKey(frmEnrollment.lsvFaculty.SelectedItem)
    
Exit Sub
err:
MsgBox err.Description, vbCritical
End Sub
Private Sub Tool_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
        Case "Browse"
            mnubrowse_Click
        Case "Classroom"
            mnuclassusage_Click
        Case "Curriculum"
            mnuCurriculum_Click
        Case "Enrollment"
            mnuEnrollment_Click
        Case "Settings"
            mnusettings_Click
    End Select
End Sub
