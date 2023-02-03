VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmBankRec 
   Caption         =   "Bank Reconciliation"
   ClientHeight    =   8610
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13710
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmBankRec.frx":0000
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   ScaleHeight     =   8610
   ScaleWidth      =   13710
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdClosePeriod 
      Caption         =   "Close Period"
      Height          =   360
      Left            =   3120
      TabIndex        =   101
      Top             =   8040
      Width           =   1395
   End
   Begin VB.CommandButton cmdAmount 
      Caption         =   "Commit"
      Height          =   360
      Left            =   1560
      TabIndex        =   98
      Top             =   8025
      Width           =   1395
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Print Recon Report"
      Height          =   345
      Left            =   4800
      TabIndex        =   96
      Top             =   8040
      Width           =   1755
   End
   Begin VB.CommandButton cmdFindAcc 
      Caption         =   "<>"
      Height          =   315
      Left            =   1455
      TabIndex        =   88
      Top             =   270
      Width           =   390
   End
   Begin VB.CommandButton cmdOffset 
      Caption         =   "Print Cash Book"
      Height          =   345
      Left            =   6600
      TabIndex        =   79
      Top             =   8085
      Width           =   1770
   End
   Begin VB.CommandButton cmdTransferFunds 
      Caption         =   "Transfer Funds"
      Height          =   345
      Left            =   8340
      TabIndex        =   70
      Top             =   8085
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.TextBox txtBankName 
      Height          =   300
      Left            =   1845
      TabIndex        =   69
      Top             =   300
      Width           =   3450
   End
   Begin VB.ComboBox cboBank 
      Height          =   315
      Left            =   75
      TabIndex        =   66
      Top             =   300
      Width           =   1365
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3150
      Top             =   6690
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdExport 
      Caption         =   "&Export"
      Height          =   345
      Left            =   0
      TabIndex        =   65
      Top             =   8040
      Width           =   1395
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6840
      Left            =   120
      TabIndex        =   11
      Top             =   1065
      Width           =   13515
      _ExtentX        =   23839
      _ExtentY        =   12065
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      Tab             =   1
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Cash Book Transactions"
      TabPicture(0)   =   "frmBankRec.frx":0442
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label27(1)"
      Tab(0).Control(1)=   "ListView1"
      Tab(0).Control(2)=   "fraTransfer"
      Tab(0).Control(3)=   "chkCheck"
      Tab(0).Control(4)=   "Frame1"
      Tab(0).Control(5)=   "cmdLoad"
      Tab(0).Control(6)=   "txtDifference(1)"
      Tab(0).ControlCount=   7
      TabCaption(1)   =   "Bank Debits"
      TabPicture(1)   =   "frmBankRec.frx":045E
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label15"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label16"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label17"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label18"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Label24"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Label25"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "lvwDebits"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "txtCrAccNo"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "txtCrAccName"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "txtDrNarration"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "txtDrAmount"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "txtDrDocumentNo"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "dtpDrTransDate"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "cmdAddDebits"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "lvwCrAccounts"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "cmdDrPost"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).ControlCount=   16
      TabCaption(2)   =   "Bank Credits"
      TabPicture(2)   =   "frmBankRec.frx":047A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label19"
      Tab(2).Control(1)=   "Label20"
      Tab(2).Control(2)=   "Label21"
      Tab(2).Control(3)=   "Label22"
      Tab(2).Control(4)=   "Label23"
      Tab(2).Control(5)=   "Label26"
      Tab(2).Control(6)=   "dtpCrTransDate"
      Tab(2).Control(7)=   "lvwCredits"
      Tab(2).Control(8)=   "txtDrAccNo"
      Tab(2).Control(9)=   "txtDrAccName"
      Tab(2).Control(10)=   "txtCrNarration"
      Tab(2).Control(11)=   "txtCrAmount"
      Tab(2).Control(12)=   "lvwDrAccounts"
      Tab(2).Control(13)=   "txtCrDocumentNo"
      Tab(2).Control(14)=   "cmdAddCredit"
      Tab(2).Control(15)=   "cmdCrPost"
      Tab(2).ControlCount=   16
      TabCaption(3)   =   "Reconciliation Report"
      TabPicture(3)   =   "frmBankRec.frx":0496
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "txtDifference(0)"
      Tab(3).Control(1)=   "txtOpeningBal"
      Tab(3).Control(2)=   "txtBankBal"
      Tab(3).Control(3)=   "txtBankDebits"
      Tab(3).Control(4)=   "txtBankCredits"
      Tab(3).Control(5)=   "txtDeposits"
      Tab(3).Control(6)=   "txtUnpresentedChq"
      Tab(3).Control(7)=   "txtCBBalance"
      Tab(3).Control(8)=   "txtPayments"
      Tab(3).Control(9)=   "txtReceipts"
      Tab(3).Control(10)=   "Label27(0)"
      Tab(3).Control(11)=   "Label6"
      Tab(3).Control(12)=   "Label14"
      Tab(3).Control(13)=   "Label13"
      Tab(3).Control(14)=   "Label12"
      Tab(3).Control(15)=   "Label11"
      Tab(3).Control(16)=   "Label10"
      Tab(3).Control(17)=   "Label9"
      Tab(3).Control(18)=   "Label8"
      Tab(3).Control(19)=   "Label7"
      Tab(3).ControlCount=   20
      Begin VB.TextBox txtDifference 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   -69990
         Locked          =   -1  'True
         TabIndex        =   99
         Text            =   "0"
         Top             =   6360
         Width           =   1950
      End
      Begin VB.CommandButton cmdLoad 
         Caption         =   "Load"
         Height          =   435
         Left            =   -74880
         TabIndex        =   97
         Top             =   6285
         Width           =   1395
      End
      Begin VB.Frame Frame1 
         Caption         =   "GL TRANSACTIONS"
         Height          =   3495
         Left            =   -70320
         TabIndex        =   90
         Top             =   720
         Width           =   6015
         Begin VB.CommandButton cmdUpdate 
            Caption         =   "&Update"
            Height          =   375
            Left            =   3480
            TabIndex        =   95
            Top             =   240
            Width           =   1095
         End
         Begin VB.TextBox txtDocNo 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1440
            Locked          =   -1  'True
            TabIndex        =   94
            Top             =   360
            Width           =   1815
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Close"
            Height          =   375
            Left            =   4680
            TabIndex        =   92
            Top             =   240
            Width           =   975
         End
         Begin MSComctlLib.ListView ListView3 
            Height          =   2535
            Left            =   240
            TabIndex        =   91
            Top             =   720
            Width           =   5655
            _ExtentX        =   9975
            _ExtentY        =   4471
            View            =   3
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   4
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "DRACCNO"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "CRACCNO"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "AMOUNT"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "MEMBERNO"
               Object.Width           =   2540
            EndProperty
         End
         Begin VB.Label Label30 
            Caption         =   "DocumentNo"
            Height          =   255
            Left            =   240
            TabIndex        =   93
            Top             =   360
            Width           =   1095
         End
      End
      Begin VB.CheckBox chkCheck 
         Caption         =   "Select All"
         Height          =   210
         Left            =   -74925
         TabIndex        =   89
         Top             =   360
         Width           =   1290
      End
      Begin VB.Frame fraTransfer 
         Caption         =   "Transfer Funds"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2130
         Left            =   -73230
         TabIndex        =   71
         Top             =   2520
         Visible         =   0   'False
         Width           =   6210
         Begin MSComCtl2.DTPicker dtpTransDate 
            Height          =   315
            Left            =   60
            TabIndex        =   86
            Top             =   1080
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   " dd-MM-yyyy"
            Format          =   94830595
            CurrentDate     =   39549
         End
         Begin VB.TextBox txtDocumentNo 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   4575
            TabIndex        =   84
            Top             =   1080
            Width           =   1590
         End
         Begin VB.TextBox txtMemberNo 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2745
            TabIndex        =   82
            Top             =   1080
            Width           =   1800
         End
         Begin VB.TextBox txtAmount 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1380
            TabIndex        =   80
            Top             =   1080
            Width           =   1335
         End
         Begin MSComctlLib.ListView ListView2 
            Height          =   1245
            Left            =   1545
            TabIndex        =   72
            Top             =   765
            Visible         =   0   'False
            Width           =   4290
            _ExtentX        =   7567
            _ExtentY        =   2196
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            HideColumnHeaders=   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   2
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "AccNo"
               Object.Width           =   18
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "AccName"
               Object.Width           =   10583
            EndProperty
         End
         Begin VB.TextBox txtAccNo 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   45
            TabIndex        =   76
            Top             =   480
            Width           =   1335
         End
         Begin VB.TextBox txtAccName 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1530
            TabIndex        =   75
            Top             =   480
            Width           =   4305
         End
         Begin VB.CommandButton cmdTransfer 
            Caption         =   "Transfer"
            Height          =   390
            Left            =   1635
            TabIndex        =   74
            Top             =   1650
            Width           =   1425
         End
         Begin VB.CommandButton cmdCancel 
            Caption         =   "&Cancel"
            Height          =   390
            Left            =   3075
            TabIndex        =   73
            Top             =   1650
            Width           =   1425
         End
         Begin VB.Label lblTransDate 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Transaction Date"
            Height          =   195
            Left            =   90
            TabIndex        =   87
            Top             =   840
            Width           =   1230
         End
         Begin VB.Label lblDocumentNo 
            AutoSize        =   -1  'True
            Caption         =   "Document No"
            Height          =   195
            Left            =   4590
            TabIndex        =   85
            Top             =   840
            Width           =   960
         End
         Begin VB.Label lblMemberNo 
            AutoSize        =   -1  'True
            Caption         =   "MemberNo/Organisation"
            Height          =   195
            Left            =   2775
            TabIndex        =   83
            Top             =   855
            Width           =   1740
         End
         Begin VB.Label lblAmount 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "New Amount"
            Height          =   195
            Left            =   1725
            TabIndex        =   81
            Top             =   840
            Width           =   915
         End
         Begin VB.Label lblAccNo 
            AutoSize        =   -1  'True
            Caption         =   "Account No"
            Height          =   210
            Left            =   60
            TabIndex        =   78
            Top             =   240
            Width           =   975
         End
         Begin VB.Label lblAccName 
            AutoSize        =   -1  'True
            Caption         =   "Account Name"
            Height          =   195
            Left            =   1545
            TabIndex        =   77
            Top             =   240
            Width           =   1035
         End
      End
      Begin VB.CommandButton cmdCrPost 
         Caption         =   "&Post"
         Height          =   375
         Left            =   -68835
         TabIndex        =   64
         Top             =   4785
         Width           =   1380
      End
      Begin VB.CommandButton cmdDrPost 
         Caption         =   "&Post"
         Height          =   375
         Left            =   6165
         TabIndex        =   63
         Top             =   4785
         Width           =   1380
      End
      Begin VB.TextBox txtDifference 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   -71055
         Locked          =   -1  'True
         TabIndex        =   61
         Text            =   "0"
         Top             =   4560
         Width           =   1950
      End
      Begin MSComctlLib.ListView lvwCrAccounts 
         Height          =   1320
         Left            =   1650
         TabIndex        =   60
         Top             =   4470
         Visible         =   0   'False
         Width           =   2805
         _ExtentX        =   4948
         _ExtentY        =   2328
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         HideColumnHeaders=   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "AccountName"
            Object.Width           =   8819
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "AccNo"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.CommandButton cmdAddDebits 
         Caption         =   "&Add"
         Height          =   375
         Left            =   4620
         TabIndex        =   59
         Top             =   4785
         Width           =   1380
      End
      Begin VB.CommandButton cmdAddCredit 
         Caption         =   "&Add"
         Height          =   375
         Left            =   -70380
         TabIndex        =   58
         Top             =   4785
         Width           =   1380
      End
      Begin MSComCtl2.DTPicker dtpDrTransDate 
         Height          =   330
         Left            =   300
         TabIndex        =   54
         Top             =   4845
         Width           =   1440
         _ExtentX        =   2540
         _ExtentY        =   582
         _Version        =   393216
         CustomFormat    =   " dd-MM-yyyy"
         Format          =   94830595
         CurrentDate     =   39407
      End
      Begin VB.TextBox txtDrDocumentNo 
         Height          =   300
         Left            =   9045
         TabIndex        =   52
         Top             =   4170
         Width           =   1590
      End
      Begin VB.TextBox txtCrDocumentNo 
         Height          =   300
         Left            =   -65955
         TabIndex        =   50
         Top             =   4170
         Width           =   1590
      End
      Begin MSComctlLib.ListView lvwDrAccounts 
         Height          =   1320
         Left            =   -73365
         TabIndex        =   49
         Top             =   4470
         Visible         =   0   'False
         Width           =   2805
         _ExtentX        =   4948
         _ExtentY        =   2328
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         HideColumnHeaders=   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "AccountName"
            Object.Width           =   8819
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "AccNo"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.TextBox txtCrAmount 
         Alignment       =   1  'Right Justify
         Height          =   300
         Left            =   -67575
         TabIndex        =   47
         Top             =   4170
         Width           =   1605
      End
      Begin VB.TextBox txtCrNarration 
         Height          =   300
         Left            =   -70605
         TabIndex        =   45
         Top             =   4170
         Width           =   3045
      End
      Begin VB.TextBox txtDrAccName 
         Height          =   300
         Left            =   -73350
         TabIndex        =   43
         Top             =   4170
         Width           =   2760
      End
      Begin VB.TextBox txtDrAccNo 
         Height          =   300
         Left            =   -74700
         TabIndex        =   41
         Top             =   4170
         Width           =   1350
      End
      Begin VB.TextBox txtDrAmount 
         Alignment       =   1  'Right Justify
         Height          =   300
         Left            =   7425
         TabIndex        =   39
         Top             =   4170
         Width           =   1605
      End
      Begin VB.TextBox txtDrNarration 
         Height          =   300
         Left            =   4395
         TabIndex        =   37
         Top             =   4170
         Width           =   3045
      End
      Begin VB.TextBox txtCrAccName 
         Height          =   300
         Left            =   1650
         TabIndex        =   35
         Top             =   4170
         Width           =   2760
      End
      Begin VB.TextBox txtCrAccNo 
         Height          =   300
         Left            =   300
         TabIndex        =   33
         Top             =   4170
         Width           =   1350
      End
      Begin VB.TextBox txtOpeningBal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -71055
         TabIndex        =   32
         Text            =   "0"
         Top             =   682
         Width           =   1950
      End
      Begin VB.TextBox txtBankBal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -71055
         Locked          =   -1  'True
         TabIndex        =   30
         Text            =   "0"
         Top             =   4162
         Width           =   1950
      End
      Begin VB.TextBox txtBankDebits 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -71055
         Locked          =   -1  'True
         TabIndex        =   29
         Text            =   "0"
         Top             =   3757
         Width           =   1950
      End
      Begin VB.TextBox txtBankCredits 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -71055
         Locked          =   -1  'True
         TabIndex        =   28
         Text            =   "0"
         Top             =   3367
         Width           =   1950
      End
      Begin VB.TextBox txtDeposits 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -71055
         Locked          =   -1  'True
         TabIndex        =   27
         Text            =   "0"
         Top             =   2962
         Width           =   1950
      End
      Begin VB.TextBox txtUnpresentedChq 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -71055
         Locked          =   -1  'True
         TabIndex        =   26
         Text            =   "0"
         Top             =   2572
         Width           =   1950
      End
      Begin VB.TextBox txtCBBalance 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -71055
         Locked          =   -1  'True
         TabIndex        =   25
         Text            =   "0"
         Top             =   1792
         Width           =   1950
      End
      Begin VB.TextBox txtPayments 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -71055
         Locked          =   -1  'True
         TabIndex        =   24
         Text            =   "0"
         Top             =   1417
         Width           =   1950
      End
      Begin VB.TextBox txtReceipts 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -71055
         Locked          =   -1  'True
         TabIndex        =   23
         Text            =   "0"
         Top             =   1057
         Width           =   1950
      End
      Begin MSComctlLib.ListView lvwCredits 
         Height          =   3060
         Left            =   -74715
         TabIndex        =   13
         Top             =   765
         Width           =   10110
         _ExtentX        =   17833
         _ExtentY        =   5398
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "TransDate"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Narration"
            Object.Width           =   6703
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "AccNo to Debit"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Amount"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Document No"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   5685
         Left            =   -75000
         TabIndex        =   12
         Top             =   600
         Width           =   13500
         _ExtentX        =   23813
         _ExtentY        =   10028
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   8
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "TransDate"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "TransDescription"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Receipts"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Payments"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Balance"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Document No"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "TransNo"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "ID"
            Object.Width           =   0
         EndProperty
      End
      Begin MSComctlLib.ListView lvwDebits 
         Height          =   3060
         Left            =   285
         TabIndex        =   14
         Top             =   765
         Width           =   10110
         _ExtentX        =   17833
         _ExtentY        =   5398
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "TransDate"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Narration"
            Object.Width           =   6703
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "AccNo to Credit"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Amount"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Document No"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComCtl2.DTPicker dtpCrTransDate 
         Height          =   330
         Left            =   -74700
         TabIndex        =   56
         Top             =   4845
         Width           =   1440
         _ExtentX        =   2540
         _ExtentY        =   582
         _Version        =   393216
         CustomFormat    =   " dd-MM-yyyy"
         Format          =   94830595
         CurrentDate     =   39407
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         Caption         =   "Difference ......................................................"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   -73320
         TabIndex        =   100
         Top             =   6405
         Width           =   3345
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         Caption         =   "Difference ......................................................"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   -74385
         TabIndex        =   62
         Top             =   4620
         Width           =   3345
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         Caption         =   "Trans Date"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -74580
         TabIndex        =   57
         Top             =   4620
         Width           =   930
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         Caption         =   "Trans Date"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   420
         TabIndex        =   55
         Top             =   4620
         Width           =   930
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         Caption         =   "Document No"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   9060
         TabIndex        =   53
         Top             =   3945
         Width           =   1125
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "Document No"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -65940
         TabIndex        =   51
         Top             =   3945
         Width           =   1125
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "Amount"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -66675
         TabIndex        =   48
         Top             =   3945
         Width           =   675
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "Narration"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -70545
         TabIndex        =   46
         Top             =   3945
         Width           =   795
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "Acc Name"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -72750
         TabIndex        =   44
         Top             =   3945
         Width           =   825
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "Acc No to Credit"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -74685
         TabIndex        =   42
         Top             =   3945
         Width           =   1335
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "Amount"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   8325
         TabIndex        =   40
         Top             =   3945
         Width           =   675
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Narration"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   4455
         TabIndex        =   38
         Top             =   3945
         Width           =   795
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Acc Name"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2250
         TabIndex        =   36
         Top             =   3945
         Width           =   825
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Acc No to Credit"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   315
         TabIndex        =   34
         Top             =   3945
         Width           =   1335
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Opening Balance .........................................."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -74385
         TabIndex        =   31
         Top             =   735
         Width           =   3330
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Balance as per Bank Statement .............."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -74385
         TabIndex        =   22
         Top             =   4215
         Width           =   3330
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "LESS: Bank Debits not in Cash Book ........"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -74385
         TabIndex        =   21
         Top             =   3810
         Width           =   3330
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "ADD: Bank Credits not in Cash Book ........"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -74385
         TabIndex        =   20
         Top             =   3420
         Width           =   3360
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "LESS: Deposits not yet credited ..............."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -74385
         TabIndex        =   19
         Top             =   3015
         Width           =   3345
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "ADD: Unprecented Cheques ......................"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -74385
         TabIndex        =   18
         Top             =   2625
         Width           =   3345
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Cash Book Balance ......................................"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -74385
         TabIndex        =   17
         Top             =   1845
         Width           =   3330
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Payments ......................................................"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -74385
         TabIndex        =   16
         Top             =   1470
         Width           =   3330
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Receipts ........................................................."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -74385
         TabIndex        =   15
         Top             =   1110
         Width           =   3345
      End
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   405
      Left            =   9975
      TabIndex        =   6
      Top             =   8010
      Width           =   1335
   End
   Begin VB.TextBox txtBankBalance 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   6720
      TabIndex        =   5
      Text            =   "0"
      Top             =   795
      Width           =   1575
   End
   Begin VB.TextBox txtOpBalance 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   6720
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   270
      Width           =   1500
   End
   Begin MSComCtl2.DTPicker dtpStatDate 
      Height          =   300
      Left            =   5340
      TabIndex        =   1
      Top             =   300
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   529
      _Version        =   393216
      CustomFormat    =   " dd-MM-yyyy"
      Format          =   94830595
      CurrentDate     =   39406
   End
   Begin MSComCtl2.DTPicker dtpStartDate 
      Height          =   300
      Left            =   8370
      TabIndex        =   7
      ToolTipText     =   "The last bank Recon Date."
      Top             =   300
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   529
      _Version        =   393216
      CustomFormat    =   " dd-MM-yyyy"
      Format          =   94830595
      CurrentDate     =   39406
   End
   Begin MSComCtl2.DTPicker dtpFinishDate 
      Height          =   300
      Left            =   9660
      TabIndex        =   9
      Top             =   300
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   529
      _Version        =   393216
      CustomFormat    =   " dd-MM-yyyy"
      Format          =   94830595
      CurrentDate     =   40353
   End
   Begin VB.Label Label29 
      AutoSize        =   -1  'True
      Caption         =   "Bank Code"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   68
      Top             =   75
      Width           =   885
   End
   Begin VB.Label Label28 
      AutoSize        =   -1  'True
      Caption         =   "Bank Name"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1815
      TabIndex        =   67
      Top             =   90
      Width           =   945
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Finish Date"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   9675
      TabIndex        =   10
      Top             =   75
      Width           =   930
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Start  Date"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   8430
      TabIndex        =   8
      Top             =   90
      Width           =   930
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Bank Balance"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   6885
      TabIndex        =   4
      Top             =   570
      Width           =   1125
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Opening Balance"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   6720
      TabIndex        =   2
      Top             =   45
      Width           =   1395
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Statement Date"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   5235
      TabIndex        =   0
      Top             =   90
      Width           =   1365
   End
End
Attribute VB_Name = "frmBankRec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim li As ListItem
Dim bCredits As Double, bDebits As Double
Dim F As Boolean

Dim cnn As New ADODB.Connection
Private Sub cboBank_Change()
Dim toLess As Double
    Set rst = oSaccoMaster.GetRecordset("select GLACCName from GLSETUP where ACCNO='" & cboBank & "'")
    If Not rst.EOF Then
       txtBankName.Text = rst(0)


'
'
''        'sql = "select gl.openingbal,cub.availablebalance,NewGLOpeningBal,NewGLOpeningBaldate from glsetup gl inner join cub on gl.accno=cub.accno where gl.accno='" & cboBank & "'"
'      ' sql = "select gl.openingbal,cub.balance,NewGLOpeningBal,NewGLOpeningBaldate from glsetup gl inner join cub on gl.accno=cub.accno where gl.accno='" & cboBank & "'"
     sql = "select openingbal,NewGLOpeningBal,NewGLOpeningBaldate from glsetup where accno='" & cboBank & "'"

         Set rst = oSaccoMaster.GetRecordset(sql)
          If Not rst.EOF Then
        txtOpBalance.Text = rst(0)
          Else
           txtOpBalance.Text = 0#
          End If
        Set rs = oSaccoMaster.GetRecordset("set dateformat dmy exec getPeriodicGlBalance '" & cboBank & "','" & dtpFinishDate & "','" & Get_Server_Date & "'")
        If Not rs.EOF Then
            toLess = IIf(IsNull(rs(0)) = True, 0, rs(0))
        Else
            toLess = 0
        End If
        txtCBBalance.Text = rst(0) - toLess
       
       ' txtOpeningBal.Text = rst(0)
        txtOpBalance.Text = rst(1)
         txtOpeningBal = txtOpBalance
        dtpStartDate.value = rst(2)
        txtBankBalance_Click
    End If
End Sub

Private Sub cboBank_Click()
    cboBank_Change
End Sub

Private Sub chkCheck_Click()
    With ListView1
            If .ListItems.Count > 0 Then
                If chkCheck.value = vbChecked Then
                    For I = 1 To .ListItems.Count
                        .ListItems(I).Checked = True
                    Next I
                Else
                    For I = 1 To .ListItems.Count
                        .ListItems(I).Checked = False
                    Next I

                End If
            End If
    End With
    Load_Statement

End Sub

Private Sub cmdAddCredit_Click()
    On Error GoTo SysError
    If Trim$(txtDrAccNo) = "" Then
        MsgBox "Please select an account to Debit", vbInformation, Me.Caption
        txtDrAccName.SetFocus
        Exit Sub
    End If
    If Trim$(txtDrAccName) = "" Then
        MsgBox "Please select an account to Debit", vbInformation, Me.Caption
        txtDrAccName.SetFocus
        Exit Sub
    End If
    If Trim$(txtCrNarration) = "" Then
        MsgBox "Please enter a description for the Transaction", vbInformation, Me.Caption
        txtCrNarration.SetFocus
        Exit Sub
    End If
    If Trim$(txtCrAmount) = "" Then
        MsgBox "Please enter an Amount.", vbInformation, Me.Caption
        txtCrAmount.SetFocus
        Exit Sub
    End If
    If Trim$(txtCrDocumentNo) = "" Then
        MsgBox "Please enter a Document No", vbInformation, Me.Caption
        txtCrDocumentNo.SetFocus
        Exit Sub
    End If
    Set li = lvwCredits.ListItems.Add(, , dtpCrTransDate)
    li.SubItems(1) = txtCrNarration
    li.SubItems(2) = txtDrAccNo
    li.SubItems(3) = Format(txtCrAmount, Cfmt)
    li.SubItems(4) = txtCrDocumentNo
    txtCrNarration = ""
    txtCrAmount = ""
    txtCrDocumentNo = ""
    txtDrAccName = ""
    txtDrAccNo = ""
    txtDrAccNo.SetFocus
    SendKeys "{Home}+{End}"
    Load_Statement
    Exit Sub
SysError:
    MsgBox err.description, vbInformation, Me.Caption
End Sub

Private Sub cmdAddDebits_Click()
    On Error GoTo SysError
    On Error GoTo SysError
    If Trim$(txtCrAccNo) = "" Then
        MsgBox "Please select an account to Credit", vbInformation, Me.Caption
        txtCrAccName.SetFocus
        Exit Sub
    End If
    If Trim$(txtCrAccName) = "" Then
        MsgBox "Please select an account to Credit", vbInformation, Me.Caption
        txtDrAccName.SetFocus
        Exit Sub
    End If
    If Trim$(txtDrNarration) = "" Then
        MsgBox "Please enter a description for the Transaction", vbInformation, Me.Caption
        txtDrNarration.SetFocus
        Exit Sub
    End If
    If Trim$(txtDrAmount) = "" Then
        MsgBox "Please enter an Amount.", vbInformation, Me.Caption
        txtDrAmount.SetFocus
        Exit Sub
    End If
    If Trim$(txtDrDocumentNo) = "" Then
        MsgBox "Please enter a Document No", vbInformation, Me.Caption
        txtDrDocumentNo.SetFocus
        Exit Sub
    End If
    Set li = lvwDebits.ListItems.Add(, , dtpDrTransDate)
    li.SubItems(1) = txtDrNarration
    li.SubItems(2) = txtCrAccNo
    li.SubItems(3) = Format(txtDrAmount, Cfmt)
    li.SubItems(4) = txtDrDocumentNo
    txtDrNarration = ""
    txtDrAmount = ""
    txtDrDocumentNo = ""
    txtCrAccName = ""
    txtCrAccNo = ""
    txtCrAccNo.SetFocus
    SendKeys "{Home}+{End}"
    Load_Statement
    Exit Sub
SysError:
    MsgBox err.description, vbInformation, Me.Caption
End Sub

Private Sub cmdAmount_Click()
    On Error GoTo SysError
    Dim reconid As Integer
    If MsgBox("Confirm this Activity", vbQuestion + vbYesNo) = vbNo Then
        Exit Sub
    End If
    Dim Id As Double
    Dim prevdocumentno As String
    DocumentNo = ""
    With ListView1
        If .ListItems.Count > 0 Then
            Set cnn = New ADODB.Connection
            cnn.Open "MAZIWA"
            'get the last reconid
            Set rst = oSaccoMaster.GetRecordset("Select isnull(max(reconid),0)+1 reconid from bankrecon")
            If Not rst.EOF Then
                reconid = rst(0)
            End If
            
            For I = 1 To .ListItems.Count
                DocumentNo = CStr(.ListItems(I).ListSubItems(5))
                If prevdocumentno <> DocumentNo Then
                    prevdocumentno = DocumentNo
                    If .ListItems(I).Checked = True Then
                        sql = "update gltransactions set recon=1,reconid=" & reconid & " where documentno='" & DocumentNo & "'"
                    Else
                        sql = "update gltransactions set recon=0,reconid=" & reconid & " where documentno='" & DocumentNo & "'"
                    End If
                    cnn.Execute sql
                End If
            Next I
        End If
    End With
        MsgBox "Recon Done!"
    Exit Sub
SysError:
    oSaccoMaster.goConn.RollbackTrans
    MsgBox err.description, vbInformation, Me.Caption
End Sub

Private Sub cmdCancel_Click()
    fraTransfer.Visible = False
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdClosePeriod_Click()
    On Error GoTo Capture
    Dim reconid As Integer
    'If there are no unreconciled items, close the statement period.
    
    If MsgBox("Close the statement period with the last date being " & dtpFinishDate & "?", vbQuestion + vbYesNo, vbCrLf) = vbYes Then
        oSaccoMaster.ExecuteThis ("set dateformat dmy update glsetup set newGlOpeningBal=" & CDbl(txtCBBalance.Text) & ",newglopeningbaldate='" & dtpFinishDate.value & "'" _
        & " where accno='" & cboBank.Text & "'")
        
        
         sql = "set dateformat dmy INSERT INTO BankRecon ( AccNo,ReconDate,OpeningBalDate,OpeningBal,Receipts,Payments,Unpresented," _
        & "UnCredited,DirectCredits,DirectDebits,StatementBal)" _
        & " VALUES ( '" & cboBank & "','" & dtpFinishDate & "','" & dtpStartDate & "'," & txtOpBalance & "," & txtReceipts & "," & txtPayments & ", " & txtUnpresentedChq & "," _
        & " " & txtDeposits & "," & txtBankCredits & "," & txtBankDebits & ", " & txtBankBal & ")"

        If Not oSaccoMaster.Execute(sql) Then
            GoTo Capture
        Else
        
            Set rst = oSaccoMaster.GetRecordset("Select isnull(max(reconid),0) reconid from bankrecon")
            If Not rst.EOF Then
                reconid = rst(0)
            End If
            With ListView1
                For I = 1 To .ListItems.Count
                    If ListView1.ListItems(I).Checked = False Then
                    
                        If CDbl(ListView1.ListItems(I).SubItems(2)) <> 0 Then
                            sql = "insert into ReconDocs (accno,DocumentNo,DrAmount,ReconId)" _
                            & " Values ('" & cboBank.Text & "','" & .ListItems(I).SubItems(5) & "'," & CDbl(.ListItems(I).SubItems(2)) & "," & reconid & ")"
                        End If
                        
                        If CDbl(.ListItems(I).SubItems(3)) <> 0 Then
                            sql = "insert into ReconDocs (accno,DocumentNo,CrAmount,ReconId)" _
                            & " Values ('" & cboBank.Text & "','" & .ListItems(I).SubItems(5) & "'," & CDbl(.ListItems(I).SubItems(3)) & "," & reconid & ")"
                        End If
                        
                        If Not oSaccoMaster.Execute(sql) Then
                            MsgBox ErrorMessage
                        End If
                        
                    End If
                Next I
            End With
        
        End If
        MsgBox "Recon Done!"
        
    End If
    Exit Sub
Capture:
    MsgBox IIf(ErrorMessage = "", err.description, ErrorMessage)
End Sub

Private Sub cmdCrPost_Click()
    On Error GoTo SysError
    Dim dAmount As Double, DRaccno As String, DocNo As String, sNarration As String
    For I = 1 To lvwCredits.ListItems.Count
        dAmount = CDbl(lvwCredits.ListItems(I).SubItems(3))
        DRaccno = lvwCredits.ListItems(I).SubItems(2)
        DocNo = lvwCredits.ListItems(I).SubItems(4)
        sNarration = lvwCredits.ListItems(I).SubItems(1)
        mTransDate = lvwCredits.ListItems(I)
        If Not SaveGLTRANSACTION(Date, dAmount, DRaccno, Craccno, DocNo, sNarration, sNarration, auditid, transactionNo) Then
            If ErrorMessage <> "" Then
                MsgBox ErrorMessage, vbInformation, Me.Caption
                ErrorMessage = ""
            End If
        End If
    Next
    Exit Sub
SysError:
    MsgBox err.description, vbInformation, Me.Caption
End Sub

Private Sub cmdDrPost_Click()
    On Error GoTo SysError
    Dim dAmount As Double, DRaccno As String, DocNo As String, sNarration As String
    For I = 1 To lvwDebits.ListItems.Count
        dAmount = CDbl(lvwDebits.ListItems(I).SubItems(3))
        DRaccno = lvwDebits.ListItems(I).SubItems(2)
        DocNo = lvwDebits.ListItems(I).SubItems(4)
        sNarration = lvwDebits.ListItems(I).SubItems(1)
        mTransDate = lvwDebits.ListItems(I)
        If Not SaveGLTRANSACTION(Date, dAmount, DRaccno, Craccno, DocNo, sNarration, sNarration, auditid, transactionNo) Then
            If ErrorMessage <> "" Then
                MsgBox ErrorMessage, vbInformation, Me.Caption
                ErrorMessage = ""
            End If
        End If
    Next
    Exit Sub
SysError:
    MsgBox err.description, vbInformation, Me.Caption
End Sub

Private Sub cmdExport_Click()
    On Error GoTo SsyError
    Dim MyFso As New FileSystemObject, strData As String, MFile As TextStream, _
    FileName As String, I As Long, li As ListItem
    If ListView1.ListItems.Count > 0 Then
        With CommonDialog1
            .Filter = "Comma Seperated Values|*.csv"
            .ShowSave
            If .FileName <> "" Then
                FileName = .FileName
            End If
            .FileName = ""
        End With
        Set MFile = MyFso.OpenTextFile(FileName, ForWriting, True)
        strData = "TransDate,MemberNo,Names,Receipts,Payments,TotalAmount,Document No"
        MFile.WriteLine strData
        strData = ""
        For I = 1 To ListView1.ListItems.Count
            Set li = ListView1.ListItems(I)
            strData = li & "," & li.SubItems(6) & "," & li.SubItems(1) & "," & CDbl(li.SubItems(2)) _
            & "," & CDbl(li.SubItems(3)) & "," & CDbl(li.SubItems(4)) & "," & li.SubItems(5)
            MFile.WriteLine strData
            strData = ""
        Next I
    Else
        MsgBox "There are no records to be exported", vbInformation, Me.Caption
    End If
    Exit Sub
SsyError:
    MsgBox err.description, vbInformation, Me.Caption
End Sub

Private Sub cmdFindacc_Click()
'    On Error GoTo SysError
'    frmsearchBanks.Show vbModal, Me
'    If Continue Then
'        If SearchValue <> "" Then
'            cboBank = SearchValue
'            SearchValue = ""
'        End If
'    End If
'    Exit Sub
'SysError:
'    MsgBox err.description, vbInformation, Me.Caption
On Error Resume Next
    frmAcctsSearch.Show vbModal, Me
    If Continue Then
        If SearchValue <> "" Then
            cboBank = SearchValue
            SearchValue = ""
            Continue = False
        End If
    End If
End Sub

Private Sub cmdLoad_Click()
    On Error GoTo hell
    bCredits = 0
    bDebits = 0
    LoadTransactions
    'Form_Load
    Exit Sub
hell:
    MsgBox err.description
End Sub

Private Sub Load_Statement()
    On Error GoTo SysEror
    Dim UnprecCheques As Double, BankDebits As Double, BankCredits As Double, _
    UnDebs As Double
    bDebits = 0
    bCredits = 0
    For I = 1 To ListView1.ListItems.Count
        If ListView1.ListItems(I).Checked = False Then
            If CDbl(ListView1.ListItems(I).SubItems(3)) <> 0 Then
                UnprecCheques = UnprecCheques + CDbl(ListView1.ListItems(I).SubItems(3))
            End If
            If CDbl(ListView1.ListItems(I).SubItems(2)) <> 0 Then
                UnDebs = UnDebs + CDbl(ListView1.ListItems(I).SubItems(2))
            End If
        Else
            bDebits = bDebits + ListView1.ListItems(I).SubItems(2)
            bCredits = bCredits + ListView1.ListItems(I).SubItems(3)
        End If
    Next I
    For I = 1 To lvwDebits.ListItems.Count
        If CDbl(lvwDebits.ListItems(I).SubItems(3)) <> 0 Then
            BankDebits = BankDebits + CDbl(lvwDebits.ListItems(I).SubItems(3))
        End If
    Next
    For I = 1 To lvwCredits.ListItems.Count
        If CDbl(lvwCredits.ListItems(I).SubItems(3)) <> 0 Then
            BankCredits = BankCredits + CDbl(lvwCredits.ListItems(I).SubItems(3))
        End If
    Next
    txtUnpresentedChq = UnprecCheques
    txtBankCredits = BankCredits
    txtBankDebits = BankDebits
    txtDeposits = UnDebs
    txtPayments = bCredits
    txtReceipts = bDebits
    txtCBBalance = txtOpeningBal + CDbl(txtReceipts - txtPayments)
    txtBankBalance_Change
    'DoTheDrCr
    Exit Sub
SysEror:
    MsgBox err.description, vbInformation, Me.Caption
End Sub
Sub DoTheDrCr()
    With ListView1
        If .SelectedItem.Checked = True Then
            bDebits = bDebits + .ListItems(.SelectedItem.index).ListSubItems(2)
            bCredits = bCredits + .ListItems(.SelectedItem.index).ListSubItems(3)
        Else
            bDebits = bDebits - .ListItems(.SelectedItem.index).ListSubItems(2)
            bCredits = bCredits - .ListItems(.SelectedItem.index).ListSubItems(3)
        End If
    End With
    txtReceipts.Text = bDebits
    txtPayments.Text = bCredits


End Sub


Private Sub cmdrefresh_Click()
    On Error GoTo SysError
    Dim UnprecCheques As Double
    For I = 1 To ListView1.ListItems.Count
        If ListView1.ListItems(I).Checked = False Then
            If CDbl(ListView1.ListItems(I).SubItems(2)) <> 0 Then
                UnprecCheques = UnprecCheques + CDbl(ListView1.ListItems(I).SubItems(2))
            End If
        Else
        End If
    Next I
    txtUnpresentedChq = Format(UnprecCheques, Cfmt)
    Exit Sub
SysError:
    MsgBox err.description, vbInformation, Me.Caption
End Sub



Private Sub Command1_Click()
    Frame1.Visible = False
End Sub

Private Sub cmdOffset_Click()
reportname = "CashBookTrans.rpt"
STRFORMULA = ""
Show_Sales_Crystal_Report STRFORMULA, reportname, CompanyName
'    On Error GoTo sysError
'    Dim FromAcc As String, ToAcc As String, DocumentNo As String, TransT As String, _
'    transdate As Date, mCredit As String
'    Dim rsTransfer As New Recordset
'    If MsgBox("Do you want to Change the selected Transaction to an Offset?", _
'    vbQuestion + vbYesNo, Me.Caption) = vbNo Then
'        Exit Sub
'    End If
'    If MsgBox("This process is not reversible. Do you want to continue?", _
'    vbExclamation + vbYesNo, Me.Caption) = vbNo Then
'        Exit Sub
'    End If
'    If ListView1.ListItems.Count > 0 Then
'        Set li = ListView1.SelectedItem
'        transdate = li
'        DocumentNo = CStr(li.SubItems(5))
'        TransT = CStr(li.SubItems(6))
'        Select Case CDbl(li.SubItems(3)) 'PAYMENTS
'            Case 0 'XXXXXXXXXXX RECEIPTS XXXXXXXXX
'            mCredit = "DR"
'            Case Else 'XXXXXXXX PAYMENTS XXXXXXXXX
'            mCredit = "CR"
'        End Select
'        Set rsTransfer = oSaccoMaster.GetRecordSet("Set DateFormat DMY Update CUSTOMERBALANCE" _
'        & " Set IDNo='2' where AccNo='" & cboBank & "' and VNo='" & TransT & _
'        "' and ChequeNo='" & DocumentNo & "' and TransDate='" & transdate & "' and TransType='" _
'        & mCredit & "'")
'        cmdLoad_Click
'    End If
'    Exit Sub
'sysError:
'    MsgBox Err.Description, vbInformation, Me.Caption
End Sub

Private Sub cmdPostRecon_Click()
    Dim rsRecon As New Recordset
    On Error GoTo SysError
    'XXXXXXXXXXX Update CashBook Entries XXXXXXXXXXXXXXXXXXXX'
    For I = 1 To ListView1.ListItems.Count
        mTransDate = ListView1.ListItems(I)
        If ListView1.ListItems(I).Checked = True Then
            Set rsRecon = oSaccoMaster.GetRecordset("Set DateFormat DMY Update " _
            & "CUSTOMERBALANCE Set Reconcile=1 where AccNo='A001' and TransDate='" _
            & mTransDate & "' and VNo='" & ListView1.ListItems(I).SubItems(1) _
            & "' and ChequeNo='" & ListView1.ListItems(I).SubItems(5) & "'")
        End If
    Next
    Exit Sub
SysError:
    MsgBox err.description, vbInformation, Me.Caption
End Sub

Private Sub cmdTransfer_Click()
    On Error GoTo SysError
    Dim FromAcc As String, ToAcc As String, DocumentNo As String, TransT As String, _
    transdate As Date, mCredit As String, mAmount As Double
    Dim rsTransfer As New Recordset
    Select Case SearchForm
        Case "Transfer"
        If ListView1.ListItems.Count > 0 Then
            Set li = ListView1.SelectedItem
            transdate = li
            DocumentNo = CStr(li.SubItems(5))
            TransT = CStr(li.SubItems(6))
            Select Case CDbl(li.SubItems(3)) 'PAYMENTS
                Case 0 'RECEIPTS
                mCredit = "DR"
                mAmount = CDbl(li.SubItems(2))
                Case Else 'PAYMENTS
                mCredit = "CR"
                mAmount = CDbl(li.SubItems(3))
            End Select
            Set rsTransfer = oSaccoMaster.GetRecordset("Set DateFormat DMY Update CUSTOMERBALANCE" _
            & " Set AccNo='" & txtAccNo & "' where AccNo='" & cboBank & "' and VNo='" & TransT & _
            "' and ChequeNo='" & DocumentNo & "' and TransDate='" & transdate & "' and TransType='" _
            & mCredit & "'")
            fraTransfer.Visible = False
            cmdLoad_Click
        End If
        Case "Update Amount"
        If ListView1.ListItems.Count > 0 Then
            Set li = ListView1.SelectedItem
            dtpTransDate = Format(li, " dd-MM-yyyy")
            transdate = Format(li, "dd-MM-yyyy")
            DocumentNo = CStr(li.SubItems(5))
            TransT = CStr(li.SubItems(6))
            Select Case CDbl(li.SubItems(3)) 'PAYMENTS
                Case 0 'RECEIPTS
                mCredit = "DR"
                mAmount = CDbl(li.SubItems(2))
                'txtAmount = li.SubItems(4)
                Case Else 'PAYMENTS
                mCredit = "CR"
                mAmount = CDbl(li.SubItems(3))
                'txtAmount = li.SubItems(3)
            End Select
            Set rsTransfer = oSaccoMaster.GetRecordset("Set DateFormat DMY Update CUSTOMERBALANCE" _
            & " Set Amount='" & txtAmount & "' where VNo='" & TransT & "' and ChequeNo='" & DocumentNo _
            & "' and TransDate='" & dtpTransDate & "' and Amount=" & mAmount)
            fraTransfer.Visible = False
            cmdLoad_Click
        End If
    End Select
    lblAmount.Visible = False
    txtAmount.Visible = False
    txtDocumentNo.Visible = False
    txtMemberNo.Visible = False
    lblMemberNo.Visible = False
    lblDocumentNo.Visible = False
    Exit Sub
SysError:
    MsgBox err.description, vbInformation, Me.Caption
End Sub

Private Sub Command2_Click()
reportname = "ReconReport.rpt"
STRFORMULA = ""
Show_Sales_Crystal_Report STRFORMULA, reportname, CompanyName
End Sub

'Private Sub cmdTransferFunds_Click()
'    On Error GoTo sysError
'    'frmTransferFunds.Show , Me
'    SearchForm = "Transfer"
'    txtAccName.Visible = True
'    txtAccNo.Visible = True
'    lblAccName.Visible = True
'    lblAccNo.Visible = True
'    txtAmount.Visible = False
'    txtDocumentNo.Visible = False
'    txtMemberNo.Visible = False
'    lblMemberNo.Visible = False
'    lblDocumentNo.Visible = False
'    lblAmount.Visible = False
'    fraTransfer.Visible = True
'    Exit Sub
'sysError:
'    MsgBox Err.Description, vbInformation, Me.Caption
'End Sub

Private Sub Form_Load()
    On Error GoTo SysError
        dtpFinishDate = Format(Get_Server_Date, " dd-MM-yyyy")
        dtpDrTransDate = dtpFinishDate
        dtpCrTransDate = dtpFinishDate
        Frame1.Visible = False
    Exit Sub
SysError:
    MsgBox err.description, vbInformation, Me.Caption
End Sub
Private Sub Load_Ledgers(DocNo As String, AccNo As String, Optional mdate As Date)
    On Error GoTo SysError
    Dim rsLedger As New Recordset
    Frame1.Visible = True
    ListView3.ListItems.Clear
    Set rsLedger = oSaccoMaster.GetRecordset("Select * From gltransactions where " _
    & " documentno='" & _
    DocNo & "' and (drAccNo='" & AccNo & "'or crAccNo='" & AccNo & "') and transdate='" & mdate & "' order by ID")
    With rsLedger
        If .State = adStateOpen Then
            While Not .EOF
                Set li = ListView3.ListItems.Add(, , IIf(IsNull(!DRaccno), "", !DRaccno))
                li.SubItems(1) = IIf(IsNull(!Craccno), "", !Craccno)
                li.SubItems(2) = IIf(IsNull(!amount), 0, !amount)
                li.SubItems(3) = IIf(IsNull(!Source), "", !Source)
                .MoveNext
            Wend
        End If
    End With
    Exit Sub
SysError:
    MsgBox err.description, vbInformation, Me.Caption
End Sub
Private Sub LoadTransactions()
    'On Error GoTo SysError
    Dim rsRecon As New Recordset, BankBal As Double, bCredits As Double, bDebits As Double, _
    RsDesc As New Recordset
    Dim unreconcile As Double
    If Trim(cboBank) = "" Then
        MsgBox "Please indicate the Bank Account to Reconcile", vbInformation, Me.Caption
        txtBankName.SetFocus
        SendKeys "{Home}+{End}"
        Exit Sub
    End If
    If dtpStartDate > dtpFinishDate Then
        MsgBox "The StartDate should be Earlier than the FinishDate", vbInformation, Me.Caption
        Exit Sub
    End If
    Get_GL_AccDetails cboBank
    If dtpStartDate < EarliestTransDate Then
        MsgBox "The StartDate should  not be earlier than " & Format(EarliestTransDate, _
        "dd-MM-yyyy"), vbInformation, Me.Caption
        dtpStartDate.SetFocus
        Exit Sub
    End If
    
    BankBal = GlAccBalance
    'Set rsRecon = oSaccoMaster.GetRecordSet("SET DATEFORMAT DMY EXEC GETgLtRANSACTIONS '" & cboBank & "','" & dtpStartDate.Value & "','" & dtpFinishDate.Value & "'")
    If Not oSaccoMaster.Execute("SET DATEFORMAT DMY EXEC getReconGlTransactions '" & cboBank & "','" & dtpStartDate.value & "','" & dtpFinishDate.value & "'") Then
        GoTo SysError
    End If
    Set rsRecon = oSaccoMaster.GetRecordset("select sum(amount)Amount,transdate,transtype,documentno,chequeno,recon,TransDescript from tempGlTransactions where documentno in (select documentno from gltransactions where (draccno='" & cboBank.Text & "' or craccno='" & cboBank.Text & "'  ))   group by transdate,documentno,transtype,chequeno,recon,TransDescript")
    
    ListView1.ListItems.Clear
    With rsRecon
        While Not .EOF
            DoEvents
            Set li = ListView1.ListItems.Add(, , !transdate)
Jump:       If !DocumentNo = "1005697" Then MsgBox "HERE"
            li.SubItems(1) = !TransDescript
            li.SubItems(2) = Format(IIf(!transtype <> "DR", 0, !amount), Cfmt)
            

'            If li.Checked = True Then
                bDebits = bDebits + CDbl(li.SubItems(2))
'            Else
'                bDebits = bDebits + 0
'            End If
            
            li.SubItems(3) = Format(IIf(!transtype <> "CR", 0, !amount), Cfmt)
            
            'If li.Checked = True Then
                bCredits = bCredits + CDbl(li.SubItems(3))
'            Else
'                bCredits = bCredits + 0
'            End If
            
            If !transtype = "DR" And !amount < 0 Then
            li.SubItems(3) = (!amount * -1)
            li.SubItems(2) = 0
            End If
            If !transtype = "CR" And !amount < 0 Then
            li.SubItems(2) = (!amount * -1)
            li.SubItems(3) = 0
            End If
            
            BankBal = BankBal + CDbl(li.SubItems(2)) - CDbl(li.SubItems(3))
            li.SubItems(4) = Format(BankBal, Cfmt)
            li.SubItems(5) = IIf(IsNull(!DocumentNo), "", !DocumentNo)
            If !DocumentNo = "RC NO 66777" Then MsgBox !amount
            li.SubItems(6) = IIf(IsNull(!chequeno), "", !chequeno)
            'li.SubItems(7) = IIf(IsNull(![id]), "", ![id])
            li.Checked = IIf(!recon = True, True, False)
            .MoveNext
        Wend
    End With
    
    '// remove the unchecked items from the
        
          F = True
    '
    sql = "select (select sum(amount) from gltransactions where draccno='" & cboBank & "' and transdate between '" & dtpStartDate & "' and '" & dtpFinishDate.value & "')  as Receipts,(select sum(amount) from gltransactions where craccno='" & cboBank & "' and transdate between '" & dtpStartDate & "' and '" & dtpFinishDate.value & "')  as Payments"
    Set rst = oSaccoMaster.GetRecordset(sql)
    
    
    txtReceipts = Format(rst!Receipts, Cfmt)
    txtPayments = Format(rst!Payments, Cfmt)
    'txtCBBalance = Format(BankBal, Cfmt) + CDbl(IIf(txtOpBalance = "", 0, txtOpBalance))
    Load_Statement
    Exit Sub
SysError:
    MsgBox IIf(ErrorMessage = "", err.description, ErrorMessage)
End Sub

Private Sub ListView1_Click()
    Load_Statement
End Sub

Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
  With ListView1
    .Sorted = True
      .SortKey = ColumnHeader.SubItemIndex
      If .SortOrder = lvwAscending Then
          .SortOrder = lvwDescending
      Else
          .SortOrder = lvwAscending
      End If
    End With

End Sub

Private Sub listview1_DblClick()
    If ListView1.ListItems.Count > 0 Then
        Set li = ListView1.SelectedItem
        mDocNo = li.SubItems(5)
        txtDocNo.Text = mDocNo
        Load_Ledgers mDocNo, cboBank.Text, li.Text
    End If
End Sub


Private Sub ListView2_Click()
    On Error GoTo SysError
    If ListView2.ListItems.Count > 0 Then
        Set li = ListView2.SelectedItem
        txtAccNo = li
        txtAccName = li.SubItems(1)
'        lvwAccounts.ListItems.Clear
'        lvwAccounts.Visible = False
    End If
    Exit Sub
SysError:
    MsgBox err.description, vbInformation, Me.Caption
End Sub

Private Sub lvwAccounts_Click()
    On Error GoTo SysError
'    If lvwAccounts.ListItems.count > 0 Then
'        Set li = lvwAccounts.SelectedItem
'        cboBank = li
'        txtBankName = li.SubItems(1)
'        lvwAccounts.ListItems.Clear
'        lvwAccounts.Visible = False
'    End If
    Exit Sub
SysError:
    MsgBox err.description, vbInformation, Me.Caption
End Sub

Private Sub lvwCrAccounts_Click()
    On Error GoTo SysError
    If lvwCrAccounts.ListItems.Count > 0 Then
        txtCrAccName = lvwCrAccounts.SelectedItem
        txtCrAccNo = lvwCrAccounts.SelectedItem.SubItems(1)
    End If
    lvwCrAccounts.Visible = False
    Exit Sub
SysError:
    MsgBox err.description, vbInformation, Me.Caption
End Sub

Private Sub lvwDrAccounts_Click()
    On Error GoTo SysError
    If lvwDrAccounts.ListItems.Count > 0 Then
        txtDrAccName = lvwDrAccounts.SelectedItem
        txtDrAccNo = lvwDrAccounts.SelectedItem.SubItems(1)
    End If
    lvwDrAccounts.Visible = False
    Exit Sub
SysError:
    MsgBox err.description, vbInformation, Me.Caption
End Sub

Private Sub txtAccName_Change()
    On Error GoTo SysError
    Dim rsAccount As New Recordset
    ListView2.ListItems.Clear
    If Trim$(txtAccName) <> "" Then
        If Not Editing Then
            Set rsAccount = oSaccoMaster.GetRecordset("Select * From GLSETUP where " _
            & "GLAccName like '%" & txtAccName & "%'")
            With rsAccount
                If .State = adStateOpen Then
                    If Not .EOF Then
                        ListView2.Visible = True
                        While Not .EOF
                            Set li = ListView2.ListItems.Add(, , IIf(IsNull(!AccNo), "", !AccNo))
                            li.SubItems(1) = IIf(IsNull(!GlAccName), "", !GlAccName)
                            .MoveNext
                        Wend
                        If ListView2.ListItems.Count = 1 Then
                            txtAccNo = li
                            txtAccName = li.SubItems(1)
                            ListView2.ListItems.Clear
                            ListView2.Visible = False
                        End If
                    Else
                        ListView2.Visible = False
                    End If
                End If
            End With
        End If
    Else
        ListView2.Visible = False
    End If
    Exit Sub
SysError:
    MsgBox err.description, vbInformation, Me.Caption
End Sub

Private Sub txtAmount_KeyPress(KeyAscii As Integer)
    On Error GoTo SysError
    Select Case KeyAscii
        Case 48 To 57
        Case Is = 46
        Case Is = 8
        Case Else
        KeyAscii = 0
    End Select
    Exit Sub
SysError:
    MsgBox err.description, vbInformation, Me.Caption
End Sub

Private Sub txtBankBal_Change()
        txtBankBal = txtBankBal
End Sub

Private Sub txtBankBalance_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
            txtBankBalance = txtBankBalance
    End If
End Sub

Private Sub txtBankBalance_LostFocus()
    txtBankBalance_KeyPress 13
End Sub

Private Sub txtBankName_Change()
    'On Error GoTo SysError
'    Dim rsAccount As New Recordset
'    lvwAccounts.ListItems.Clear
'    If Trim$(txtBankName) <> "" Then
'        If Not Editing Then
'            Set rsAccount = oSaccoMaster.GetRecordSet("Select * From GLSETUP where " _
'            & "GLAccName like '%" & txtBankName & "%'")
'            With rsAccount
'                If .State = adStateOpen Then
'                    If Not .EOF Then
'                        lvwAccounts.Visible = True
'                        While Not .EOF
'                            Set li = lvwAccounts.ListItems.Add(, , IIf(IsNull(!Accno), "", !Accno))
'                            li.SubItems(1) = IIf(IsNull(!GlAccName), "", !GlAccName)
'                            .MoveNext
'                        Wend
'                        If lvwAccounts.ListItems.count = 1 Then
'                            cboBank = li
'                            txtBankName = li.SubItems(1)
'                            lvwAccounts.ListItems.Clear
'                            lvwAccounts.Visible = False
'                        End If
'                    Else
'                        lvwAccounts.Visible = False
'                    End If
'                End If
'            End With
'        End If
'    Else
'        lvwAccounts.Visible = False
'    End If
    Exit Sub
SysError:
    MsgBox err.description, vbInformation, Me.Caption
End Sub

Private Sub txtCBBalance_Change()
    txtCBBalance.Text = txtCBBalance
End Sub

Private Sub txtBankBalance_Change()
    On Error Resume Next
    txtBankBal = txtBankBalance
    txtDifference(0) = CDbl(IIf(txtBankBalance.Text = "", 0, txtBankBalance.Text)) - CDbl(IIf(txtCBBalance.Text = "", 0, txtCBBalance.Text))
    txtDifference(1) = CDbl(IIf(txtBankBalance.Text = "", 0, txtBankBalance.Text)) - CDbl(IIf(txtCBBalance.Text = "", 0, txtCBBalance.Text))
End Sub

Private Sub txtBankBalance_Click()
    txtBankBalance_Change
End Sub

Private Sub txtCrAccName_Change()
    On Error GoTo SysError
    Dim rsAcc As New Recordset
    lvwCrAccounts.ListItems.Clear
    If Trim$(txtCrAccName) <> "" Then
        Set rsAcc = oSaccoMaster.GetRecordset("Select AccNo,GLAccName From GLSETUP " _
        & "where GLACCName like '%" & txtCrAccName & "%'")
        With rsAcc
            If Not .EOF Then
                lvwCrAccounts.Visible = True
                While Not .EOF
                    Set li = lvwCrAccounts.ListItems.Add(, , !GlAccName)
                    li.SubItems(1) = IIf(IsNull(!AccNo), "", !AccNo)
                    .MoveNext
                Wend
            Else
                lvwCrAccounts.Visible = False
            End If
        End With
    Else
        lvwCrAccounts.Visible = False
    End If
    Exit Sub
SysError:
    MsgBox err.description, vbInformation, Me.Caption
End Sub

Private Sub txtCrAmount_KeyPress(KeyAscii As Integer)
    On Error GoTo SysError
    Select Case KeyAscii
        Case 48 To 57
        Case Is = 46
        Case Is = 8
        Case 13
        If Trim$(txtCrAmount) Then
            Set li = lvwCredits.ListItems.Add(, , dtpCrTransDate)
            li.SubItems(1) = txtCrNarration
            li.SubItems(2) = txtDrAccNo
            li.SubItems(3) = Format(txtCrAmount, Cfmt)
            li.SubItems(4) = txtCrDocumentNo
        End If
        Case Else
        KeyAscii = 0
    End Select
    Exit Sub
SysError:
    MsgBox err.description, vbInformation, Me.Caption
End Sub

Private Sub txtDrAccName_Change()
    On Error GoTo SysError
    Dim rsAcc As New Recordset
    lvwDrAccounts.ListItems.Clear
    If Trim$(txtDrAccName) <> "" Then
        Set rsAcc = oSaccoMaster.GetRecordset("Select AccNo,GLAccName From GLSETUP " _
        & "where GLACCName like '%" & txtDrAccName & "%'")
        With rsAcc
            If Not .EOF Then
                lvwDrAccounts.Visible = True
                While Not .EOF
                    Set li = lvwDrAccounts.ListItems.Add(, , !GlAccName)
                    li.SubItems(1) = IIf(IsNull(!AccNo), "", !AccNo)
                    .MoveNext
                Wend
            Else
                lvwDrAccounts.Visible = False
            End If
        End With
    Else
        lvwDrAccounts.Visible = False
    End If
    Exit Sub
SysError:
    MsgBox err.description, vbInformation, Me.Caption
End Sub

Private Sub txtDrAmount_KeyPress(KeyAscii As Integer)
    On Error GoTo SysError
    Select Case KeyAscii
        Case 48 To 57
        Case Is = 46
        Case Is = 8
        Case 13
'        If Trim(txtDrAmount) <> "" Then
'            Set li = lvwCredits.ListItems.Add(, , dtpDrTransDate)
'            li.SubItems(1) = txtDrNarration
'            li.SubItems(2) = txtDrAccNo
'            li.SubItems(3) = Format(txtDrAmount, CfMt)
'            li.SubItems(4) = txtDrDocumentNo
'        End If
        Case Else
        KeyAscii = 0
    End Select
    Exit Sub
SysError:
    MsgBox err.description, vbInformation, Me.Caption
End Sub

Private Sub txtOpBalance_Change()
    txtOpBalance = txtOpBalance
End Sub

Private Sub txtOpBalance_Click()
    txtOpBalance_Change
End Sub
