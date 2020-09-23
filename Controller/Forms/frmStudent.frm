VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmStudent 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Student Info"
   ClientHeight    =   9510
   ClientLeft      =   150
   ClientTop       =   210
   ClientWidth     =   14535
   Icon            =   "frmStudent.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9510
   ScaleWidth      =   14535
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab sTab 
      Height          =   9495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   14535
      _ExtentX        =   25638
      _ExtentY        =   16748
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "General Info"
      TabPicture(0)   =   "frmStudent.frx":492A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "imgPicture"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lsvHistory"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "GridInfo"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "GridRemarks"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lsvScholarShip"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "Transcript"
      TabPicture(1)   =   "frmStudent.frx":4946
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "statTranscript"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "lsvTranscript"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Evaluation"
      TabPicture(2)   =   "frmStudent.frx":4962
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "txtCourseID"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "statEval"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "lsvEvaluation"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).ControlCount=   3
      TabCaption(3)   =   "Billing and Payments"
      TabPicture(3)   =   "frmStudent.frx":497E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "PaymentTab"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "BillingTab"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "MiscTab"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "tabTuition"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).ControlCount=   4
      Begin VB.TextBox txtCourseID 
         Height          =   285
         Left            =   -73920
         TabIndex        =   19
         Top             =   6600
         Visible         =   0   'False
         Width           =   855
      End
      Begin MSComctlLib.ListView lsvScholarShip 
         Height          =   2055
         Left            =   0
         TabIndex        =   17
         Top             =   7395
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   3625
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Scholarship"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Academic Year"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Semester"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Change By"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.StatusBar statEval 
         Height          =   375
         Left            =   -75000
         TabIndex        =   3
         Top             =   8160
         Width           =   10815
         _ExtentX        =   19076
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
            NumPanels       =   4
            BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               Object.Width           =   3528
               MinWidth        =   3528
            EndProperty
            BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               Object.Width           =   3528
               MinWidth        =   3528
            EndProperty
            BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               Object.Width           =   3528
               MinWidth        =   3528
            EndProperty
            BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               Object.Width           =   8113
               MinWidth        =   8113
               Text            =   "Subjects marked in red are not offered in the current semester."
               TextSave        =   "Subjects marked in red are not offered in the current semester."
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.StatusBar statTranscript 
         Height          =   375
         Left            =   -75000
         TabIndex        =   2
         Top             =   6480
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
            NumPanels       =   2
            BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            EndProperty
            BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               Object.Width           =   6174
               MinWidth        =   6174
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ListView lsvTranscript 
         Height          =   5295
         Left            =   -75000
         TabIndex        =   1
         Top             =   360
         Width           =   7215
         _ExtentX        =   12726
         _ExtentY        =   9340
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "SubjectID"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Subject"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Desciptive Title"
            Object.Width           =   7056
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Units"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Final Grade"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Comp Grade"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ListView lsvEvaluation 
         Height          =   5295
         Left            =   -75000
         TabIndex        =   4
         Top             =   360
         Width           =   7215
         _ExtentX        =   12726
         _ExtentY        =   9340
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "SubjectID"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Subject"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Desciptive Title"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Grade"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Units"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Take"
            Object.Width           =   2540
         EndProperty
      End
      Begin TabDlg.SSTab PaymentTab 
         Height          =   3375
         Left            =   -75000
         TabIndex        =   5
         Top             =   6120
         Width           =   14535
         _ExtentX        =   25638
         _ExtentY        =   5953
         _Version        =   393216
         Style           =   1
         Tabs            =   1
         TabsPerRow      =   1
         TabHeight       =   520
         TabCaption(0)   =   "Payment Accounts"
         TabPicture(0)   =   "frmStudent.frx":499A
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "StatusBar1"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "lsvPayment"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "lsvPaymentItem"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).ControlCount=   3
         Begin MSComctlLib.ListView lsvPaymentItem 
            Height          =   2595
            Left            =   7080
            TabIndex        =   6
            Top             =   360
            Width           =   7425
            _ExtentX        =   13097
            _ExtentY        =   4577
            View            =   3
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   2
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Description"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Amount"
               Object.Width           =   2540
            EndProperty
         End
         Begin MSComctlLib.ListView lsvPayment 
            Height          =   2595
            Left            =   0
            TabIndex        =   7
            Top             =   360
            Width           =   7095
            _ExtentX        =   12515
            _ExtentY        =   4577
            View            =   3
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   4
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Payment Date"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "OR No."
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Recieved By"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Total Amount"
               Object.Width           =   2540
            EndProperty
         End
         Begin MSComctlLib.StatusBar StatusBar1 
            Height          =   375
            Left            =   60
            TabIndex        =   16
            Top             =   2940
            Width           =   13935
            _ExtentX        =   24580
            _ExtentY        =   661
            _Version        =   393216
            BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
               NumPanels       =   6
               BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
                  Object.Width           =   3528
                  MinWidth        =   3528
               EndProperty
               BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
                  Object.Width           =   7056
                  MinWidth        =   7056
               EndProperty
               BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               EndProperty
               BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               EndProperty
               BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
                  Object.Width           =   5292
                  MinWidth        =   5292
               EndProperty
               BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
                  Object.Width           =   5292
                  MinWidth        =   5292
               EndProperty
            EndProperty
         End
      End
      Begin TabDlg.SSTab BillingTab 
         Height          =   5775
         Left            =   -75000
         TabIndex        =   8
         Top             =   360
         Width           =   7095
         _ExtentX        =   12515
         _ExtentY        =   10186
         _Version        =   393216
         Style           =   1
         Tabs            =   1
         TabsPerRow      =   1
         TabHeight       =   520
         TabCaption(0)   =   "Billing of Accounts"
         TabPicture(0)   =   "frmStudent.frx":49B6
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "lsvBilling"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "BillingStat"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).ControlCount=   2
         Begin MSComctlLib.StatusBar BillingStat 
            Height          =   375
            Left            =   60
            TabIndex        =   9
            Top             =   5340
            Width           =   6975
            _ExtentX        =   12303
            _ExtentY        =   661
            _Version        =   393216
            BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
               NumPanels       =   4
               BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               EndProperty
               BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               EndProperty
               BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
                  Object.Width           =   2646
                  MinWidth        =   2646
               EndProperty
               BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
                  Object.Width           =   5292
                  MinWidth        =   5292
               EndProperty
            EndProperty
         End
         Begin MSComctlLib.ListView lsvBilling 
            Height          =   4980
            Left            =   30
            TabIndex        =   10
            Top             =   360
            Width           =   7020
            _ExtentX        =   12383
            _ExtentY        =   8784
            View            =   3
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   7
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Year"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Semester"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Date Assessed"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "Total Amt."
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Text            =   "Paid Amt."
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   6
               Text            =   "Balance"
               Object.Width           =   2540
            EndProperty
         End
      End
      Begin TabDlg.SSTab MiscTab 
         Height          =   3015
         Left            =   -67920
         TabIndex        =   11
         Top             =   360
         Width           =   7455
         _ExtentX        =   13150
         _ExtentY        =   5318
         _Version        =   393216
         Style           =   1
         Tabs            =   1
         TabsPerRow      =   1
         TabHeight       =   520
         TabCaption(0)   =   "Miscellaneous Fee"
         TabPicture(0)   =   "frmStudent.frx":49D2
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "lsvBillingItem"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "MiscStat"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).ControlCount=   2
         Begin MSComctlLib.StatusBar MiscStat 
            Height          =   375
            Left            =   60
            TabIndex        =   12
            Top             =   2590
            Width           =   7335
            _ExtentX        =   12938
            _ExtentY        =   661
            _Version        =   393216
            BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
               NumPanels       =   4
               BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
                  Object.Width           =   3528
                  MinWidth        =   3528
               EndProperty
               BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               EndProperty
               BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
                  Object.Width           =   3528
                  MinWidth        =   3528
               EndProperty
               BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
                  Object.Width           =   3528
                  MinWidth        =   3528
               EndProperty
            EndProperty
         End
         Begin MSComctlLib.ListView lsvBillingItem 
            Height          =   2250
            Left            =   60
            TabIndex        =   13
            Top             =   360
            Width           =   7335
            _ExtentX        =   12938
            _ExtentY        =   3969
            View            =   3
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   4
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Description"
               Object.Width           =   3528
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Amount"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Paid Amt."
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Balance"
               Object.Width           =   2540
            EndProperty
         End
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid GridRemarks 
         Height          =   2355
         Left            =   0
         TabIndex        =   14
         Top             =   5040
         Width           =   12135
         _ExtentX        =   21405
         _ExtentY        =   4154
         _Version        =   393216
         Rows            =   8
         FixedRows       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid GridInfo 
         Height          =   4650
         Left            =   0
         TabIndex        =   15
         Top             =   360
         Width           =   12135
         _ExtentX        =   21405
         _ExtentY        =   8202
         _Version        =   393216
         Rows            =   16
         FixedRows       =   0
         BackColorSel    =   8438015
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin MSComctlLib.ListView lsvHistory 
         Height          =   2055
         Left            =   7320
         TabIndex        =   18
         Top             =   7395
         Width           =   7200
         _ExtentX        =   12700
         _ExtentY        =   3625
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Previous Course"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "New Course"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Year/Semester"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Change By"
            Object.Width           =   2540
         EndProperty
      End
      Begin TabDlg.SSTab tabTuition 
         Height          =   2775
         Left            =   -67920
         TabIndex        =   20
         Top             =   3360
         Width           =   7455
         _ExtentX        =   13150
         _ExtentY        =   4895
         _Version        =   393216
         Style           =   1
         Tabs            =   2
         TabHeight       =   520
         TabCaption(0)   =   "Tuition Fee"
         TabPicture(0)   =   "frmStudent.frx":49EE
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "lvDetails"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "statTuition"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).ControlCount=   2
         TabCaption(1)   =   "Laboratory Fee"
         TabPicture(1)   =   "frmStudent.frx":4A0A
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "lvLab"
         Tab(1).Control(1)=   "statLab"
         Tab(1).ControlCount=   2
         Begin MSComctlLib.StatusBar statTuition 
            Height          =   375
            Left            =   60
            TabIndex        =   21
            Top             =   2280
            Width           =   7335
            _ExtentX        =   12938
            _ExtentY        =   661
            _Version        =   393216
            BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
               NumPanels       =   4
               BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
                  Object.Width           =   5292
                  MinWidth        =   5292
               EndProperty
               BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               EndProperty
               BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               EndProperty
               BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               EndProperty
            EndProperty
         End
         Begin MSComctlLib.StatusBar statLab 
            Height          =   375
            Left            =   -74940
            TabIndex        =   22
            Top             =   2280
            Width           =   7335
            _ExtentX        =   12938
            _ExtentY        =   661
            _Version        =   393216
            BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
               NumPanels       =   4
               BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
                  Object.Width           =   5292
                  MinWidth        =   5292
               EndProperty
               BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               EndProperty
               BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               EndProperty
               BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               EndProperty
            EndProperty
         End
         Begin MSComctlLib.ListView lvDetails 
            Height          =   1935
            Left            =   60
            TabIndex        =   23
            Top             =   360
            Width           =   7335
            _ExtentX        =   12938
            _ExtentY        =   3413
            View            =   3
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   5
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Subject"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Units"
               Object.Width           =   1235
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Amount"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Paid Amt."
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "Balance"
               Object.Width           =   1764
            EndProperty
         End
         Begin MSComctlLib.ListView lvLab 
            Height          =   1935
            Left            =   -74940
            TabIndex        =   24
            Top             =   360
            Width           =   7335
            _ExtentX        =   12938
            _ExtentY        =   3413
            View            =   3
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   5
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Subject"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Units"
               Object.Width           =   1235
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Amount"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Paid Amt."
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "Balance"
               Object.Width           =   1764
            EndProperty
         End
      End
      Begin VB.Image imgPicture 
         BorderStyle     =   1  'Fixed Single
         Height          =   2295
         Left            =   12160
         Picture         =   "frmStudent.frx":4A26
         Stretch         =   -1  'True
         Top             =   360
         Width           =   2295
      End
   End
   Begin VB.Menu popRefresh 
      Caption         =   "Refresh"
      Visible         =   0   'False
      Begin VB.Menu mnuRefreshProp 
         Caption         =   "Properties..."
      End
      Begin VB.Menu mnusep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRefRefres 
         Caption         =   "Refresh"
      End
   End
End
Attribute VB_Name = "frmStudent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cIRowCount, cIRowCount1, cIRowCount2 As Integer
Private Sub InitGridInfo()
    cIRowCount = 0
    With GridInfo
        .Clear
        .ClearStructure
        .FixedCols = 1
        
        .ColWidth(0) = 1800
        .ColWidth(1) = 12135 - 1800

        .TextMatrix(0, 0) = "ID Number"
        .TextMatrix(1, 0) = "Full name"
        .TextMatrix(2, 0) = "Gender"
        .TextMatrix(3, 0) = "Course"
        .TextMatrix(4, 0) = "Curriculum"
        .TextMatrix(5, 0) = "Scholarship"
        .TextMatrix(6, 0) = "Year Level"
        .TextMatrix(7, 0) = "Attended"
        .TextMatrix(8, 0) = "Year Admitted"
        .TextMatrix(9, 0) = "Semester admitted"
        .TextMatrix(10, 0) = "Religion"
        .TextMatrix(11, 0) = "Civil Status"
        .TextMatrix(12, 0) = "Birth date"
        .TextMatrix(13, 0) = "Birth place"
        .TextMatrix(14, 0) = "Home Address"
        .TextMatrix(15, 0) = "High School"
        
        
        .ColAlignment(0) = vbLeftJustify
        .ColAlignment(1) = vbLeftJustify
        
    End With
End Sub

Private Sub InitGridRemarks()
    cIRowCount1 = 0
    With GridRemarks
        .Clear
        .ClearStructure
        .Rows = 8
        .FixedCols = 1
        
        .ColWidth(0) = 1800
        .ColWidth(1) = 12135 - 1800

        .TextMatrix(0, 0) = "Acad. Remarks"
        .TextMatrix(1, 0) = "Total units"
        .TextMatrix(2, 0) = "Max units"
        .TextMatrix(3, 0) = "Required units"
        .TextMatrix(4, 0) = "Earned units"
        .TextMatrix(5, 0) = "Remaining units"
        .TextMatrix(6, 0) = "Percent passing"
        .TextMatrix(7, 0) = "Acad. adviser"
        
        .ColAlignment(0) = vbLeftJustify
        .ColAlignment(1) = vbLeftJustify
        
    End With
End Sub

Private Sub Form_Load()
    InitGridInfo
    InitGridRemarks
End Sub

Private Sub Form_Resize()
    GridInfo.ColWidth(1) = GridInfo.Width - GridInfo.ColWidth(0)
    GridRemarks.ColWidth(1) = GridRemarks.Width - GridRemarks.ColWidth(0)
    
    sTab.Height = ScaleHeight
    sTab.Width = ScaleWidth
    
    GridInfo.Left = 0
    GridInfo.Width = ScaleWidth - (imgPicture.Width + 100)
    imgPicture.Left = GridInfo.Width + 30
    
    GridRemarks.Top = GridInfo.Height + 400
    GridRemarks.Width = ScaleWidth - (imgPicture.Width + 100)
    

    lsvTranscript.Height = sTab.Height - (statTranscript.Height + 400)
    lsvTranscript.Width = sTab.Width - 50
    statTranscript.Width = lsvTranscript.Width
    statTranscript.Top = lsvTranscript.Height + 350
    
    lsvEvaluation.Height = sTab.Height - (statEval.Height + 400)
    lsvEvaluation.Width = sTab.Width - 50
    statEval.Top = lsvEvaluation.Height + 350
    statEval.Width = lsvEvaluation.Width
End Sub

Public Function ShowStudentDetail(sStudentID As String)
On Error GoTo err
Dim vRS As New ADODB.Recordset
Dim sSQL As String
Dim NewSY As String
    
    'sSQL = "SELECT tblStudent.StudentID, tblStudent.LastName & ', ' & tblStudent.FirstName & ' ' & tblStudent.MiddleName AS Fullname, tblStudent.Gender, tblCourse.Curriculum, tblCourse.Course, tblCourse.Major " &
    '        "FROM (tblCourse RIGHT JOIN (tblStudent LEFT JOIN tblEnrolment ON tblStudent.StudentID = tblEnrolment.StudentID) ON tblCourse.CourseID = tblEnrolment.CourseID) LEFT JOIN tblScholar ON tblStudent.StudentID = tblScholar.StudentID, tblScholar.Scholarship,tblStudent.LastSchoolName, tblStudent.Religion, tblStudent.Address, tblStudent.PlaceOfBirth, tblStudent.BirthDate, tblStudent.Status " & _

    sSQL = "SELECT tblStudent.StudentID, tblStudent.LastName & ', ' & tblStudent.FirstName & ' ' & tblStudent.MiddleName AS Fullname, tblStudent.Gender, tblStudent.LastSchoolName, tblStudent.Religion, tblStudent.Address, tblStudent.PlaceOfBirth, tblStudent.BirthDate, tblStudent.Status " & _
            "FROM tblStudent " & _
            "WHERE tblStudent.StudentID='" & sStudentID & "'"
    
    If ConnectRS(con, vRS, sSQL) = True Then
            With GridInfo
                .Clear
                .ClearStructure
                .FixedCols = 1
                
                .ColWidth(0) = 1800
                .ColWidth(1) = 3000
                
                .TextMatrix(0, 0) = "ID Number"
                .TextMatrix(1, 0) = "Full name"
                .TextMatrix(2, 0) = "Gender"
                .TextMatrix(3, 0) = "Course"
                .TextMatrix(4, 0) = "Curriculum"
                .TextMatrix(5, 0) = "Scholarship"
                .TextMatrix(6, 0) = "Year Level"
                .TextMatrix(7, 0) = "Attended"
                .TextMatrix(8, 0) = "Year Admitted"
                .TextMatrix(9, 0) = "Semester admitted"
                .TextMatrix(10, 0) = "Religion"
                .TextMatrix(11, 0) = "Civil Status"
                .TextMatrix(12, 0) = "Birth date"
                .TextMatrix(13, 0) = "Birth place"
                .TextMatrix(14, 0) = "Home Address"
                .TextMatrix(15, 0) = "High School"
        
                .ColAlignment(0) = vbLeftJustify
                .ColAlignment(1) = vbLeftJustify
                
                .TextMatrix(0, 1) = vRS.Fields("StudentID").Value
                .TextMatrix(1, 1) = vRS.Fields("Fullname").Value
                .TextMatrix(2, 1) = vRS.Fields("Gender").Value
    
                
                .TextMatrix(10, 1) = vRS.Fields("Religion").Value
                .TextMatrix(11, 1) = vRS.Fields("Status").Value
                .TextMatrix(12, 1) = vRS.Fields("BirthDate").Value
                .TextMatrix(13, 1) = vRS.Fields("PlaceOfBirth").Value
                .TextMatrix(14, 1) = vRS.Fields("Address").Value
                .TextMatrix(15, 1) = vRS.Fields("LastSchoolName").Value
                
                Set imgPicture.Picture = LoadPicture(App.Path & "\myPic\" & .TextMatrix(0, 1) & ".jpg")
             End With
    End If
    
                 
    LoadCourse (sStudentID)
    CheckScholarship (sStudentID)
    CheckCurrentYearLevel (sStudentID)
    ShowTotalEnrolledUnits (sStudentID)
    ShowEarnedUnits (sStudentID)
    LoadTranscript (sStudentID)
    LoadEvaluation (sStudentID)
    ShowBilledAccount (sStudentID)
    
    
    Me.Show vbModal
             
    Exit Function
    
err:
    MsgBox err.Description, vbCritical
    Exit Function
End Function
    
Private Sub LoadCourse(sStudentID As String)
Dim vRS As New ADODB.Recordset
Dim sSQL As String
    sSQL = "SELECT tblStudent.StudentID, tblStudent.LastName & ', ' & tblStudent.FirstName & ' ' & tblStudent.MiddleName AS Fullname,tblCourse.CourseID, tblCourse.Course, tblCourse.Major, tblCourse.Curriculum, tblStudentStatus.PreviousCourseID " & _
            "FROM tblStudent INNER JOIN (tblCourse INNER JOIN tblStudentStatus ON tblCourse.CourseID = tblStudentStatus.CourseID) ON tblStudent.StudentID = tblStudentStatus.StudentID " & _
            "WHERE tblStudent.StudentID='" & sStudentID & "'"
            
    If ConnectRS(con, vRS, sSQL) = True Then
        If vRS.RecordCount > 0 Then
                With GridInfo
                    .TextMatrix(3, 1) = vRS.Fields("Course").Value
                    .TextMatrix(4, 1) = vRS.Fields("Curriculum").Value
                End With
                txtCourseID.Text = vRS.Fields("CourseID").Value
        End If
    End If
End Sub
Private Sub ShowTotalEnrolledUnits(sStudentID As String)
Dim rs As New ADODB.Recordset
Dim mySQL As String

mySQL = "SELECT tblEnrolment.EnrollmentID, Sum(tblSubject.Units) AS SumOfUnits, tblEnrolment.Semester, tblStudentStatus.YearLevel " & _
"FROM (tblStudent INNER JOIN (tblEnrolment INNER JOIN ((tblSubject INNER JOIN tblSubjectOffering ON tblSubject.SubjectID = tblSubjectOffering.SubjectID) INNER JOIN tblGrade ON (tblSubjectOffering.SubjectOfferingID = tblGrade.SubjectOfferingID) AND (tblSubjectOffering.SubjectOfferingID = tblGrade.SubjectOfferingID)) ON tblEnrolment.EnrollmentID = tblGrade.EnrolmentID) ON tblStudent.StudentID = tblEnrolment.StudentID) INNER JOIN tblStudentStatus ON tblStudent.StudentID = tblStudentStatus.StudentID " & _
"WHERE tblEnrolment.EnrollmentID='" & sStudentID & "'" & _
"GROUP BY tblEnrolment.EnrollmentID, tblEnrolment.Semester, tblStudentStatus.YearLevel;"

    If ConnectRS(con, rs, mySQL) = True Then
         With GridRemarks
            .TextMatrix(1, 1) = rs.Fields("SumOfUnits")
         End With
    End If
End Sub
Private Sub CheckCurrentYearLevel(sStudentID As String)
Dim vRS As New ADODB.Recordset
Dim sSQL As String
If ConnectRS(con, vRS, "Select * from tblStudentStatus WHERE StudentID='" & sStudentID & "'") = True Then
    vRS.MoveLast
    With GridInfo
        .TextMatrix(6, 1) = vRS.Fields("YearLevel").Value
        .TextMatrix(7, 1) = vRS.Fields("YearLevel").Value & " - " & vRS.Fields("Semester").Value
        .TextMatrix(8, 1) = vRS.Fields("SchoolYear").Value
        .TextMatrix(9, 1) = vRS.Fields("Semester").Value
    End With
End If
End Sub
Private Sub CheckScholarship(sStudentID As String)
Dim vRS As New ADODB.Recordset
Dim sSQL As String

sSQL = "SELECT tblStudent.StudentID, tblStudent.LastName & ', ' & tblStudent.FirstName & ' ' & tblStudent.MiddleName AS Fullname, tblScholar.Scholarship " & _
        "FROM tblStudent INNER JOIN tblScholar ON tblStudent.StudentID = tblScholar.StudentID " & _
        "WHERE tblStudent.StudentID ='" & sStudentID & "'"

If ConnectRS(con, vRS, sSQL) = True Then
    If vRS.RecordCount > 0 Then
        With GridInfo
            .TextMatrix(5, 1) = vRS.Fields("Scholarship")
        End With
    Else
        With GridInfo
            .TextMatrix(5, 1) = "None"
        End With
    End If
End If
    Set vRS = Nothing
End Sub
Private Sub LoadTranscript(sStudentID As String)
Dim vRS As New ADODB.Recordset
Dim sSQL As String
Dim lv As ListItem
        
    sSQL = "SELECT tblStudent.StudentID,tblSubject.SubjectID, tblSubject.SubjectTitle, tblSubject.Description, tblSubject.Units, tblGrade.FinalGrade, tblGrade.CompGrade, tblCourse.Course " & _
            "FROM tblCourse INNER JOIN (tblStudent INNER JOIN ((tblSubject INNER JOIN tblProspectus ON tblSubject.SubjectID = tblProspectus.SubjectID) INNER JOIN (tblSubjectOffering INNER JOIN (tblEnrolment INNER JOIN tblGrade ON tblEnrolment.EnrollmentID = tblGrade.EnrolmentID) ON (tblSubjectOffering.SubjectOfferingID = tblGrade.SubjectOfferingID) AND (tblSubjectOffering.SubjectOfferingID = tblGrade.SubjectOfferingID)) ON tblSubject.SubjectID = tblSubjectOffering.SubjectID) ON tblStudent.StudentID = tblEnrolment.StudentID) ON (tblCourse.CourseID = tblProspectus.CourseID) AND (tblCourse.CourseID = tblEnrolment.CourseID) " & _
            "WHERE tblStudent.StudentID='" & sStudentID & "'"

If ConnectRS(con, vRS, sSQL) = True Then
    lsvTranscript.ListItems.Clear
    Do Until vRS.EOF
        Set lv = lsvTranscript.ListItems.Add(, , vRS.Fields("SubjectID"))
                lv.SubItems(1) = vRS.Fields("SubjectTitle")
                lv.SubItems(2) = vRS.Fields("Description")
                lv.SubItems(3) = vRS.Fields("Units")
                lv.SubItems(4) = vRS.Fields("FinalGrade")
                lv.SubItems(5) = vRS.Fields("CompGrade")
        vRS.MoveNext
    Loop
End If
    Set vRS = Nothing
End Sub
Private Sub LoadEvaluation(sStudentID As String)
Dim vRS As New ADODB.Recordset
Dim sSQL As String
Dim lv As ListItem
        
    sSQL = "SELECT tblStudent.StudentID, tblSubject.SubjectTitle, tblSubject.Description, tblSubject.Units, tblGrade.FinalGrade, tblSubject.SubjectID " & _
            "FROM tblSubject INNER JOIN (((tblCourse INNER JOIN (tblStudent INNER JOIN tblEnrolment ON tblStudent.StudentID = tblEnrolment.StudentID) ON tblCourse.CourseID = tblEnrolment.CourseID) INNER JOIN tblProspectus ON tblCourse.CourseID = tblProspectus.CourseID) INNER JOIN (tblSubjectOffering INNER JOIN tblGrade ON (tblSubjectOffering.SubjectOfferingID = tblGrade.SubjectOfferingID) AND (tblSubjectOffering.SubjectOfferingID = tblGrade.SubjectOfferingID)) ON tblEnrolment.EnrollmentID = tblGrade.EnrolmentID) ON (tblSubject.SubjectID = tblSubjectOffering.SubjectID) AND (tblSubject.SubjectID = tblProspectus.SubjectID) " & _
            "WHERE tblStudent.StudentID='" & sStudentID & "'"

If ConnectRS(con, vRS, sSQL) = True Then
    lsvEvaluation.ListItems.Clear
    Do Until vRS.EOF
        Set lv = lsvEvaluation.ListItems.Add(, , vRS.Fields("SubjectID"))
                lv.SubItems(1) = vRS.Fields("SubjectTitle")
                lv.SubItems(2) = vRS.Fields("Description")
                lv.SubItems(3) = vRS.Fields("FinalGrade")
                lv.SubItems(4) = vRS.Fields("Units")
                lv.SubItems(5) = ""
        vRS.MoveNext
    Loop
End If
    Set vRS = Nothing
End Sub

Private Sub ShowEarnedUnits(sStudentID As String)
Dim vRS As New ADODB.Recordset
Dim sSQL As String

    sSQL = "SELECT tblStudent.StudentID, Sum(tblSubject.Units) AS EarnedUnits " & _
            "FROM tblStudent INNER JOIN (tblEnrolment INNER JOIN (tblSubject INNER JOIN (tblSubjectOffering INNER JOIN tblGrade ON (tblSubjectOffering.SubjectOfferingID = tblGrade.SubjectOfferingID) AND (tblSubjectOffering.SubjectOfferingID = tblGrade.SubjectOfferingID)) ON tblSubject.SubjectID = tblSubjectOffering.SubjectID) ON tblEnrolment.EnrollmentID = tblGrade.EnrolmentID) ON tblStudent.StudentID = tblEnrolment.StudentID " & _
            "WHERE tblStudent.StudentID ='" & sStudentID & "'" & _
            "GROUP BY tblStudent.StudentID"
    If ConnectRS(con, vRS, sSQL) = True Then
         With GridRemarks
            .TextMatrix(4, 1) = vRS.Fields("EarnedUnits")
         End With
    End If
End Sub

Private Sub txtCourseID_Change()
    ShowMaxUnits
    ShowRequiredUnits
End Sub
Private Sub ShowMaxUnits()
Dim rs As New ADODB.Recordset
Dim mySQL As String

mySQL = "SELECT tblCourse.CourseID, Sum(tblSubject.Units) AS MaxUnits, tblProspectus.YearLevel, tblSemester.Semester " & _
            "FROM tblSubject INNER JOIN (tblSemester INNER JOIN (tblCourse INNER JOIN tblProspectus ON tblCourse.CourseID = tblProspectus.CourseID) ON tblSemester.SemesterID = tblProspectus.SemesterID) ON tblSubject.SubjectID = tblProspectus.SubjectID " & _
            "WHERE tblCourse.CourseID='" & txtCourseID.Text & "'" & _
            "GROUP BY tblCourse.CourseID, tblProspectus.YearLevel, tblSemester.Semester;"

 If ConnectRS(con, rs, mySQL) = True Then
         With GridRemarks
            .TextMatrix(2, 1) = rs.Fields("MaxUnits")
         End With
End If
End Sub
Private Sub ShowRequiredUnits()
Dim rs As New ADODB.Recordset
Dim mySQL As String

mySQL = "SELECT tblCourse.Course, Sum(tblSubject.Units) AS RequiredUnits, tblProspectus.CourseID " & _
        "FROM tblSubject INNER JOIN (tblSemester INNER JOIN (tblCourse INNER JOIN tblProspectus ON tblCourse.CourseID = tblProspectus.CourseID) ON tblSemester.SemesterID = tblProspectus.SemesterID) ON tblSubject.SubjectID = tblProspectus.SubjectID " & _
        "WHERE tblCourse.CourseID='" & txtCourseID.Text & "'" & _
        "GROUP BY tblCourse.Course, tblProspectus.CourseID;"

 If ConnectRS(con, rs, mySQL) = True Then
         With GridRemarks
            .TextMatrix(3, 1) = rs.Fields("RequiredUnits")
         End With
End If
End Sub
Public Sub ShowBilledAccount(sStudentID As String)
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    Dim lv As ListItem
    
    sSQL = "SELECT tblEnrolment.EnrollmentID, Sum(tblCharge.Amount) AS SumAmount, tblStudent.StudentID, tblStudent.LastName, tblEnrolment.SchoolYear, tblEnrolment.Semester, tblCharge.CreationDate " & _
            "FROM (tblStudent INNER JOIN tblEnrolment ON tblStudent.StudentID = tblEnrolment.StudentID) INNER JOIN tblCharge ON tblEnrolment.EnrollmentID = tblCharge.EnrolmentID " & _
            "WHERE tblStudent.StudentID ='" & sStudentID & "' and tblEnrolment.SchoolYear='" & CurrentSchoolYear.SchoolYearTitle & "' and tblEnrolment.Semester='" & CurrentSemester.Semester & "'" & _
            "GROUP BY tblEnrolment.EnrollmentID, tblStudent.StudentID, tblStudent.LastName, tblEnrolment.SchoolYear, tblEnrolment.Semester, tblCharge.CreationDate "

    
    lsvBilling.ListItems.Clear
    
    If ConnectRS(con, vRS, sSQL) = True Then
        Do Until vRS.EOF
            Set lv = lsvBilling.ListItems.Add(, , vRS.Fields("SchoolYear"))
            lv.SubItems(1) = vRS.Fields("Semester")
            lv.SubItems(2) = vRS.Fields("EnrollmentID")
            lv.SubItems(3) = DateValue(vRS.Fields("CreationDate"))
            lv.SubItems(4) = FormatNumber(vRS.Fields("SumAmount"), 2, vbTrue, vbTrue, vbTrue)
            vRS.MoveNext
        Loop
    End If
Set vRS = Nothing
End Sub
Public Function ShowSubjectDetail(sEnrollmentID As String, sSchoolYear As String, sSemester As String)
    Dim vRS As New ADODB.Recordset
    Dim rs As New ADODB.Recordset
    Dim sSQL As String
    Dim lv As ListItem
    Dim lv1 As ListItem
    
    Dim curTotal As Currency

    mdiController.MousePointer = vbHourglass
    
    
    sSQL = "SELECT tblEnrolment.EnrollmentID, tblEnrolment.StudentID, tblSubject.SubjectID, tblSubject.SubjectTitle, tblSubject.Units, tblSubject.SubjectFee, ([tblSubject.Units]*[tblSubject.SubjectFee]) AS Tuition, tblSubject.LaboratoryUnits, tblSubject.LaboratoryFee, ([tblSubject.LaboratoryUnits]*[tblSubject.LaboratoryFee]) AS LabFee, tblEnrolment.SchoolYear, tblEnrolment.Semester " & _
            "FROM tblSubject INNER JOIN (tblEnrolment INNER JOIN (tblSubjectOffering INNER JOIN tblGrade ON (tblSubjectOffering.SubjectOfferingID = tblGrade.SubjectOfferingID) AND (tblSubjectOffering.SubjectOfferingID = tblGrade.SubjectOfferingID)) ON tblEnrolment.EnrollmentID = tblGrade.EnrolmentID) ON tblSubject.SubjectID = tblSubjectOffering.SubjectID " & _
            "WHERE tblEnrolment.EnrollmentID='" & sEnrollmentID & "' and tblEnrolment.Semester ='" & sSemester & "' and tblEnrolment.SchoolYear='" & sSchoolYear & "'"
        
    lvDetails.ListItems.Clear
    
    If ConnectRS(con, vRS, sSQL) = True Then
        If AnyRecordExisted(vRS) Then
            Do Until vRS.EOF
                Set lv = lvDetails.ListItems.Add(, , vRS.Fields("SubjectTitle"))
                        lv.SubItems(1) = vRS.Fields("Units")
                        lv.SubItems(2) = FormatNumber(vRS.Fields("Tuition"), 2, vbTrue, vbTrue, vbTrue)
                    vRS.MoveNext
             Loop
        End If
    End If
    
    Set vRS = Nothing
    
    mdiController.MousePointer = vbDefault
End Function
Public Function ShowLabSubject(sEnrollmentID As String, sSchoolYear As String, sSemester As String)
    Dim rs As New ADODB.Recordset
    Dim sSQL As String
    Dim lv As ListItem

    mdiController.MousePointer = vbHourglass
    
    
    sSQL = "SELECT tblEnrolment.EnrollmentID, tblSubjectOffering.SubjectOfferingID, tblSubject.SubjectID, tblSubject.SubjectTitle, tblSubject.Description, tblSubject.LaboratoryUnits, tblSubject.LaboratoryFee, tblEnrolment.Semester, tblEnrolment.SchoolYear, ([tblSubject!LaboratoryUnits]*[tblSubject!LaboratoryFee]) AS Laboratory " & _
            "FROM tblSubject INNER JOIN (tblSubjectOffering INNER JOIN (tblEnrolment INNER JOIN tblGrade ON tblEnrolment.EnrollmentID = tblGrade.EnrolmentID) ON (tblSubjectOffering.SubjectOfferingID = tblGrade.SubjectOfferingID) AND (tblSubjectOffering.SubjectOfferingID = tblGrade.SubjectOfferingID)) ON tblSubject.SubjectID = tblSubjectOffering.SubjectID " & _
            "WHERE  (((([tblSubject!LaboratoryUnits]*[tblSubject!LaboratoryFee]))>0)) and tblEnrolment.EnrollmentID='" & sEnrollmentID & "' and tblEnrolment.Semester ='" & sSemester & "' and tblEnrolment.SchoolYear='" & sSchoolYear & "'"
        
    lvLab.ListItems.Clear
    
   If ConnectRS(con, rs, sSQL) = True Then
        Do Until rs.EOF
            Set lv = lvLab.ListItems.Add(, , rs.Fields("SubjectTitle"))
            lv.SubItems(1) = rs.Fields("LaboratoryUnits")
            lv.SubItems(2) = FormatNumber(rs.Fields("Laboratory"), 2, vbTrue, vbTrue, vbTrue)
            rs.MoveNext
        Loop
    End If
    
    Set rs = Nothing
        
    mdiController.MousePointer = vbDefault
End Function

