VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmEnrollment 
   Caption         =   "Enrollment"
   ClientHeight    =   8520
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   8880
   Icon            =   "frmEnrollment.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8520
   ScaleWidth      =   8880
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picLine 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      Height          =   45
      Index           =   1
      Left            =   0
      ScaleHeight     =   45
      ScaleWidth      =   8880
      TabIndex        =   23
      Top             =   570
      Width           =   8880
   End
   Begin VB.PictureBox picSection 
      Align           =   3  'Align Left
      BorderStyle     =   0  'None
      Height          =   7890
      Left            =   4815
      ScaleHeight     =   7890
      ScaleWidth      =   3600
      TabIndex        =   12
      Top             =   630
      Width           =   3600
      Begin VB.Frame frameSection 
         Height          =   465
         Left            =   0
         TabIndex        =   14
         Top             =   0
         Width           =   1890
         Begin VB.Image Image2 
            Height          =   240
            Left            =   75
            Picture         =   "frmEnrollment.frx":492A
            Top             =   150
            Width           =   240
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Sections"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   480
            TabIndex        =   15
            Top             =   240
            Width           =   1290
         End
      End
      Begin MSComctlLib.ListView lsvSection 
         Height          =   5895
         Left            =   0
         TabIndex        =   13
         Top             =   480
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   10398
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         Icons           =   "imgSection"
         SmallIcons      =   "imgSection"
         ColHdrIcons     =   "icoHeader"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   7
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Section"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "ID"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Gender"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Slots"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Status/Building/Room"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Schedule"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Enrolled"
            Object.Width           =   1587
         EndProperty
      End
      Begin MSComctlLib.ImageList imgSection 
         Left            =   1080
         Top             =   6600
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   1
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEnrollment.frx":532C
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.PictureBox picLine 
      Align           =   1  'Align Top
      BackColor       =   &H80000010&
      BorderStyle     =   0  'None
      Height          =   10
      Index           =   0
      Left            =   0
      ScaleHeight     =   15
      ScaleWidth      =   8880
      TabIndex        =   10
      Top             =   615
      Width           =   8880
   End
   Begin VB.PictureBox picSubjects 
      Align           =   3  'Align Left
      BorderStyle     =   0  'None
      Height          =   7890
      Left            =   0
      ScaleHeight     =   7890
      ScaleWidth      =   4815
      TabIndex        =   7
      Top             =   630
      Width           =   4815
      Begin VB.TextBox txtSubjectID 
         Height          =   285
         Left            =   3600
         TabIndex        =   25
         Top             =   6600
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox txtSectionID 
         Height          =   285
         Left            =   3600
         TabIndex        =   24
         Top             =   6840
         Visible         =   0   'False
         Width           =   975
      End
      Begin MSComctlLib.ListView lsvSubject 
         Height          =   5895
         Left            =   0
         TabIndex        =   11
         Top             =   480
         Width           =   4815
         _ExtentX        =   8493
         _ExtentY        =   10398
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         Icons           =   "imgSubject"
         SmallIcons      =   "imgSubject"
         ColHdrIcons     =   "icoHeader"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Subject"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "SubjectID"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Units"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Descriptive Title"
            Object.Width           =   3528
         EndProperty
      End
      Begin VB.Frame frameSubject 
         Height          =   465
         Left            =   0
         TabIndex        =   8
         Top             =   0
         Width           =   1650
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Subjects"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   480
            TabIndex        =   9
            Top             =   240
            Width           =   1290
         End
         Begin VB.Image Image3 
            Height          =   240
            Left            =   75
            Picture         =   "frmEnrollment.frx":58C6
            Top             =   150
            Width           =   240
         End
      End
      Begin MSComctlLib.ImageList icoHeader 
         Left            =   840
         Top             =   6960
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   2
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEnrollment.frx":62C8
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEnrollment.frx":6862
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageList imgSubject 
         Left            =   1440
         Top             =   6960
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   1
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEnrollment.frx":6DFC
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageList ilRecordIco 
         Left            =   2520
         Top             =   6840
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   1
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEnrollment.frx":7396
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   3840
         Top             =   7320
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   32
         ImageHeight     =   32
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   41
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEnrollment.frx":7930
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEnrollment.frx":820A
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEnrollment.frx":9B9C
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEnrollment.frx":B52E
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEnrollment.frx":CEC0
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEnrollment.frx":E852
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEnrollment.frx":101E4
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEnrollment.frx":11B76
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEnrollment.frx":13508
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEnrollment.frx":14E9A
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEnrollment.frx":1682E
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEnrollment.frx":17880
               Key             =   ""
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEnrollment.frx":18160
               Key             =   ""
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEnrollment.frx":191B2
               Key             =   ""
            EndProperty
            BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEnrollment.frx":1A204
               Key             =   ""
            EndProperty
            BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEnrollment.frx":1B256
               Key             =   ""
            EndProperty
            BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEnrollment.frx":1C2A8
               Key             =   ""
            EndProperty
            BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEnrollment.frx":1D2FA
               Key             =   ""
            EndProperty
            BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEnrollment.frx":1DBD6
               Key             =   ""
            EndProperty
            BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEnrollment.frx":1E4B0
               Key             =   ""
            EndProperty
            BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEnrollment.frx":1ED8A
               Key             =   ""
            EndProperty
            BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEnrollment.frx":1F664
               Key             =   ""
            EndProperty
            BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEnrollment.frx":1FF3E
               Key             =   ""
            EndProperty
            BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEnrollment.frx":20818
               Key             =   ""
            EndProperty
            BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEnrollment.frx":210F2
               Key             =   ""
            EndProperty
            BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEnrollment.frx":219CC
               Key             =   ""
            EndProperty
            BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEnrollment.frx":222A6
               Key             =   ""
            EndProperty
            BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEnrollment.frx":22B80
               Key             =   ""
            EndProperty
            BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEnrollment.frx":2345A
               Key             =   ""
            EndProperty
            BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEnrollment.frx":23D34
               Key             =   ""
            EndProperty
            BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEnrollment.frx":2460E
               Key             =   ""
            EndProperty
            BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEnrollment.frx":24EE8
               Key             =   ""
            EndProperty
            BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEnrollment.frx":257C2
               Key             =   ""
            EndProperty
            BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEnrollment.frx":2609C
               Key             =   ""
            EndProperty
            BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEnrollment.frx":26976
               Key             =   ""
            EndProperty
            BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEnrollment.frx":27250
               Key             =   ""
            EndProperty
            BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEnrollment.frx":27B2A
               Key             =   ""
            EndProperty
            BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEnrollment.frx":28404
               Key             =   ""
            EndProperty
            BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEnrollment.frx":28CDE
               Key             =   ""
            EndProperty
            BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEnrollment.frx":295B8
               Key             =   ""
            EndProperty
            BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEnrollment.frx":29E92
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.PictureBox picFaculty 
      Align           =   4  'Align Right
      BorderStyle     =   0  'None
      Height          =   7890
      Left            =   5490
      ScaleHeight     =   7890
      ScaleWidth      =   3390
      TabIndex        =   4
      Top             =   630
      Width           =   3390
      Begin VB.PictureBox picTeachingLoad 
         BorderStyle     =   0  'None
         Height          =   4935
         Left            =   0
         ScaleHeight     =   4935
         ScaleWidth      =   4215
         TabIndex        =   16
         Top             =   5400
         Width           =   4215
         Begin VB.PictureBox picLine 
            BorderStyle     =   0  'None
            Height          =   50
            Index           =   2
            Left            =   0
            ScaleHeight     =   45
            ScaleWidth      =   13950
            TabIndex        =   22
            Top             =   0
            Width           =   13950
         End
         Begin VB.Frame frameTeaching 
            Height          =   465
            Left            =   0
            TabIndex        =   17
            Top             =   0
            Width           =   2250
            Begin VB.Image Image4 
               Height          =   240
               Left            =   75
               Picture         =   "frmEnrollment.frx":2A76C
               Top             =   150
               Width           =   240
            End
            Begin VB.Label Label4 
               BackStyle       =   0  'Transparent
               Caption         =   "Teaching Load"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Left            =   480
               TabIndex        =   18
               Top             =   240
               Width           =   1290
            End
         End
         Begin MSComctlLib.StatusBar StatusBar3 
            Height          =   375
            Left            =   0
            TabIndex        =   19
            Top             =   10080
            Width           =   15240
            _ExtentX        =   26882
            _ExtentY        =   661
            _Version        =   393216
            BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
               NumPanels       =   2
               BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               EndProperty
               BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               EndProperty
            EndProperty
         End
         Begin MSComctlLib.ListView lsvLoad 
            Height          =   3495
            Left            =   0
            TabIndex        =   21
            Top             =   480
            Width           =   3375
            _ExtentX        =   5953
            _ExtentY        =   6165
            View            =   3
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            _Version        =   393217
            Icons           =   "imgSubject"
            SmallIcons      =   "imgSubject"
            ColHdrIcons     =   "icoHeader"
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   3
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Subject"
               Object.Width           =   2822
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Section"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Units"
               Object.Width           =   1376
            EndProperty
         End
      End
      Begin VB.Frame frameFaculty 
         Height          =   465
         Left            =   0
         TabIndex        =   5
         Top             =   0
         Width           =   1650
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Faculty"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   360
            TabIndex        =   6
            Top             =   195
            Width           =   1290
         End
         Begin VB.Image Image1 
            Height          =   240
            Left            =   75
            Picture         =   "frmEnrollment.frx":2B16E
            Top             =   150
            Width           =   240
         End
      End
      Begin MSComctlLib.ListView lsvFaculty 
         Height          =   4935
         Left            =   0
         TabIndex        =   20
         Top             =   480
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   8705
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         Icons           =   "ilRecordIco"
         SmallIcons      =   "ilRecordIco"
         ColHdrIcons     =   "icoHeader"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Faculty Name"
            Object.Width           =   5953
         EndProperty
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
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
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   7
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
      EndProperty
      Begin VB.ComboBox cboDepartment 
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
         Left            =   4920
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   110
         Width           =   6375
      End
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
         ItemData        =   "frmEnrollment.frx":2BB70
         Left            =   2910
         List            =   "frmEnrollment.frx":2BB7D
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   110
         Width           =   1935
      End
      Begin VB.ComboBox cmbSchoolYear 
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
         Left            =   840
         TabIndex        =   1
         Text            =   "cmbSchoolYear"
         Top             =   110
         Width           =   2055
      End
   End
End
Attribute VB_Name = "frmEnrollment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim SubjRs As New ADODB.Recordset
Dim MaxEntryCount As Long
Dim CurRecPos As Long
Dim CurStudentCount As Long

Public curSectionOfferingID As String

Public Sub ShowFormList(Optional iMaxEntryCount As Long = 21, Optional iCurRecPos As Long = 0)
    'MaxEntryCount = iMaxEntryCount
    'CurRecPos = iCurRecPos

        DepartmentList
        RefreshSYList
        
    Me.Show
End Sub
Private Sub RefreshSYList()
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    
    sSQL = "SELECT tblSchoolYear.SchoolYear" & _
            " FROM tblSchoolYear" & _
            " ORDER BY tblSchoolYear.SchoolYear"
     
    If ConnectRS(con, vRS, sSQL) = True Then
        cmbSchoolYear.Clear
        Do Until vRS.EOF
            cmbSchoolYear.AddItem (vRS.Fields("SchoolYear"))
            vRS.MoveNext
        Loop
        cmbSchoolYear.Text = CurrentSchoolYear.SchoolYearTitle
        cboSemester.Text = CurrentSemester.Semester
    End If
    
ReleaseAndExit:
    Set vRS = Nothing
End Sub
Public Sub LoadSubjectByDepartment(sDepartmentTitle As String)
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    sSQL = "SELECT tblSubject.SubjectID AS lvKey, tblSubject.SubjectTitle AS Title,tblSubject.SubjectID,tblSubject.Units, tblSubject.Description FROM tblDepartment INNER JOIN tblSubject ON tblDepartment.DepartmentID = tblSubject.DepartmentID" & _
                " Where DepartmentTitle Like '%" & sDepartmentTitle & "%'"
    
    If ConnectRS(con, vRS, sSQL) = True Then
         FillSubjectList vRS
     Else
        MsgBox "Unable to show Subject list.", vbCritical
    End If
    
End Sub

Public Sub LoadFacultyByDepartment(sDepartmentTitle As String)
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    sSQL = "SELECT tblTeacher.TeacherID as lvKey, [tblTeacher].[LastName] & ', ' & [tblTeacher].[FirstName] & ' ' & [tblTeacher].[MiddleName] AS TeacherFullName " & _
            "FROM tblDepartment INNER JOIN tblTeacher ON tblDepartment.DepartmentID = tblTeacher.DepartmentID " & _
            " Where DepartmentTitle = '" & sDepartmentTitle & "'"
    
    If ConnectRS(con, vRS, sSQL) = True Then
         FillFacultyList vRS
     Else
        MsgBox "Unable to show Teacher list.", vbCritical
    End If
    
End Sub
Private Sub DepartmentList()
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    sSQL = "SELECT DepartmentTitle" & _
            " FROM tblDepartment" & _
            " ORDER BY DepartmentTitle"
     
    If ConnectRS(con, vRS, sSQL) = False Then
        MsgBox "ERROR"
        GoTo ReleaseAndExit
    End If
    
    cmbSchoolYear.Clear
    While vRS.EOF = False
        
        cboDepartment.AddItem (vRS.Fields("DepartmentTitle"))
        vRS.MoveNext
    
    Wend
ReleaseAndExit:
    Set vRS = Nothing
End Sub
Public Sub LoadSectionBySubject(sSubject As String, sSchoolYear As String, sSemester As String)
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    sSQL = "SELECT tblSection.SectionID AS lvKey, tblSection.SectionTitle, tblSection.SectionID, tblDepartment.DepartmentTitle AS Gender, tblSection.Slots, tblRoom.Building & '- ' & tblRoom.Room AS Room, tblSubjectOffering.Days & ' ' & tblSubjectOffering.TimeIn & '- ' & tblSubjectOffering.TimeOut AS Schedule " & _
            "FROM tblSubject INNER JOIN ((tblDepartment INNER JOIN tblSection ON tblDepartment.DepartmentID = tblSection.DepartmentID) INNER JOIN (tblRoom INNER JOIN tblSubjectOffering ON tblRoom.RoomID = tblSubjectOffering.RoomID) ON tblSection.SectionID = tblSubjectOffering.SectionID) ON tblSubject.SubjectID = tblSubjectOffering.SubjectID " & _
            "Where tblSubject.SubjectID = '" & sSubject & "' and tblSubjectOffering.SchoolYear ='" & sSchoolYear & "' and tblSubjectOffering.Semester='" & sSemester & "'"
            
    If ConnectRS(con, vRS, sSQL) = True Then
         FillSectionList vRS
     Else
        MsgBox "Unable to show Section list.", vbCritical
    End If
    
End Sub
Private Function FillSectionList(ByRef vRS As ADODB.Recordset)
Dim lv As ListItem

        mdiController.MousePointer = vbHourglass
        
        UnSortLV lsvSection
        
        lsvSection.ListItems.Clear
        Do Until vRS.EOF
        Set lv = lsvSection.ListItems.Add(, , vRS.Fields("SectionTitle"), 1, 1)
                lv.SubItems(1) = vRS.Fields("SectionID")
                lv.SubItems(2) = vRS.Fields("Gender")
                lv.SubItems(3) = vRS.Fields("Slots")
                lv.SubItems(4) = vRS.Fields("Room")
                lv.SubItems(5) = vRS.Fields("Schedule")
            vRS.MoveNext
        Loop
        
        SortLV lsvSection, lsvSection.SortKey, lsvSection.SortOrder, False
        mdiController.MousePointer = vbDefault
Set vRS = Nothing
End Function
Private Function FillSubjectList(ByRef vRS As ADODB.Recordset)
        
        mdiController.MousePointer = vbHourglass
        
        UnSortLV lsvSubject
        FillRecordToList vRS, lsvSubject, KeyStudent, , , , True
        SortLV lsvSubject, lsvSubject.SortKey, lsvSubject.SortOrder, False
                
        mdiController.MousePointer = vbDefault
End Function
Private Function FillFacultyList(ByRef vRS As ADODB.Recordset)
        
        mdiController.MousePointer = vbHourglass
        
        UnSortLV lsvFaculty
        FillRecordToList vRS, lsvFaculty, KeyStudent, , , , True
        SortLV lsvFaculty, lsvFaculty.SortKey, lsvFaculty.SortOrder, False
                
        mdiController.MousePointer = vbDefault
End Function

Public Sub FormDepartment_Refresh()
    SubjRs.Requery
    FillSubjectList SubjRs
End Sub

Private Sub cboDepartment_Change()
    LoadSubjectByDepartment (cboDepartment.Text)
    LoadFacultyByDepartment (cboDepartment.Text)
End Sub

Private Sub cboDepartment_Click()
 cboDepartment_Change
End Sub

Private Sub Form_Resize()

picSection.Width = ScaleWidth - (picSubjects.Width + picFaculty.Width)

lsvSubject.Width = picSubjects.ScaleWidth
lsvSubject.Height = picSubjects.ScaleHeight - (frameSubject.Height)

lsvSection.Width = picSection.ScaleWidth
lsvSection.Height = picSection.ScaleHeight - (frameSection.Height)

lsvFaculty.Width = picFaculty.Width
lsvFaculty.Height = picFaculty.ScaleHeight - (picTeachingLoad.Height)

picTeachingLoad.Top = lsvFaculty.Height
picTeachingLoad.ScaleWidth = picFaculty.Width

lsvLoad.Width = picTeachingLoad.ScaleWidth
lsvLoad.Height = picTeachingLoad.ScaleHeight - (frameTeaching.Height)
End Sub

Private Sub lsvFaculty_Click()
   Call RefreshSubjects(GetLVKey(lsvFaculty.SelectedItem), cmbSchoolYear.Text, cboSemester.Text)
End Sub

Private Sub lsvFaculty_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        PopupMenu mdiController.popFaculty
    End If
End Sub

Private Sub lsvSection_Click()
On Error Resume Next
    txtSectionID.Text = lsvSection.SelectedItem.SubItems(1)
End Sub

Private Sub lsvSection_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        PopupMenu mdiController.popSection
    End If
End Sub

Private Sub lsvSubject_Click()
On Error GoTo err
    txtSubjectID.Text = lsvSubject.SelectedItem.SubItems(1)
    Call LoadSectionBySubject(lsvSubject.SelectedItem.SubItems(1), cmbSchoolYear.Text, cboSemester.Text)
    Exit Sub
err:
MsgBox err.Description, vbCritical
End Sub

Private Sub lsvSubject_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        PopupMenu mdiController.popSubject
    End If
End Sub


Public Sub ShowSectionOfferingDetail()
On Error GoTo err
    
    Dim sSQL As String
    Dim vRS As New ADODB.Recordset
    
    sSQL = "SELECT Count([tblEnrolment].[EnrollmentID]) AS CountOfEnrolmentID, tblSection.SectionTitle, tblSection.SectionID " & _
            "FROM tblSubject INNER JOIN (tblSubjectOffering INNER JOIN (tblSection INNER JOIN (tblEnrolment INNER JOIN tblGrade ON tblEnrolment.EnrollmentID = tblGrade.EnrolmentID) ON tblSection.SectionID = tblGrade.SectionID) ON (tblSubjectOffering.SubjectOfferingID = tblGrade.SubjectOfferingID) AND (tblSection.SectionID = tblSubjectOffering.SectionID) AND (tblSubjectOffering.SubjectOfferingID = tblGrade.SubjectOfferingID)) ON tblSubject.SubjectID = tblSubjectOffering.SubjectID " & _
            "WHERE (((tblSection.SectionID)='" & txtSectionID.Text & "')) " & _
            "GROUP BY tblSection.SectionTitle, tblSection.SectionID"

    If ConnectRS(con, vRS, sSQL) = True Then
            If (vRS.Fields("CountOfEnrolmentID")) < 1 Then
                If MsgBox("WARNING:" & vbNewLine & "You are about to delete Section Entries and you cannot Undo this operation." & vbNewLine & " Are you sure to delete it?", vbCritical + vbYesNo) = vbYes Then
                     Select Case DeleteSectionOffering(lsvSection.SelectedItem.SubItems(1))
                        Case TranDBResult.Success
                            MsgBox "Section record successfully deleted.", vbInformation
                            Call LoadSectionBySubject(txtSubjectID.Text, cmbSchoolYear.Text, cboSemester.Text)
                        Case Else
                            MsgBox "Unable to delete Section Record!", vbExclamation
                        End Select
                End If
            Else
                MsgBox "This Section Offering Record cannot be deleted." & vbNewLine & _
                "Reason: This record contain " & (vRS.Fields("CountOfEnrolmentID")) & " Enrolment  entry/s.", vbInformation
            End If
    Else
        CatchError "frmDeleteSection", "ShowSectionOfferingDetail", "Error connecting Section RS"
    End If
    
    Set vRS = Nothing
    Exit Sub
err:
  MsgBox err.Description, vbCritical

End Sub
Private Function RefreshSubjects(sTeacherID As String, sSchoolYear As String, sSemester As String)
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    lsvLoad.ListItems.Clear
    
    On Error GoTo ReleaseAndExit
    
   sSQL = "SELECT tblSubjectOffering.SubjectOfferingID, tblSubject.SubjectTitle, tblSection.SectionTitle, tblSubject.Units " & _
            "FROM tblTeacher INNER JOIN (tblSubject INNER JOIN (tblSection INNER JOIN tblSubjectOffering ON tblSection.SectionID = tblSubjectOffering.SectionID) ON tblSubject.SubjectID = tblSubjectOffering.SubjectID) ON tblTeacher.TeacherID = tblSubjectOffering.TeacherID " & _
            "WHERE tblTeacher.TeacherID='" & sTeacherID & "' and tblSection.SchoolYear='" & sSchoolYear & "' and tblSection.Semester ='" & sSemester & "'"
           

    If ConnectRS(con, vRS, sSQL) = False Then
        MsgBox "Unable to conect Teacher Recordset.", vbCritical
        GoTo ReleaseAndExit
    End If
    
    If AnyRecordExisted(vRS) = True Then
        FillRecordToList vRS, lsvLoad, KeySubjectOffering, , 32767, , True
    End If

ReleaseAndExit:
    Set vRS = Nothing
End Function
