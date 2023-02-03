VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmNominals 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "NOMINAL MODULE"
   ClientHeight    =   8070
   ClientLeft      =   45
   ClientTop       =   630
   ClientWidth     =   10605
   Icon            =   "frmNominals.frx":0000
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8070
   ScaleWidth      =   10605
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton frmglshow 
      Caption         =   "GL Inquiry"
      Height          =   375
      Left            =   8040
      TabIndex        =   74
      Top             =   6840
      Width           =   1455
   End
   Begin MSComDlg.CommonDialog dlg9 
      Left            =   4680
      Top             =   7200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7455
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   9735
      _ExtentX        =   17171
      _ExtentY        =   13150
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "RECEIPTS"
      TabPicture(0)   =   "frmNominals.frx":0442
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame2(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "PAYMENTS"
      TabPicture(1)   =   "frmNominals.frx":045E
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Frame2(1)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame2 
         Height          =   6960
         Index           =   0
         Left            =   -74880
         TabIndex        =   31
         Top             =   360
         Width           =   9510
         Begin VB.PictureBox Picture5 
            Height          =   255
            Left            =   5760
            Picture         =   "frmNominals.frx":047A
            ScaleHeight     =   195
            ScaleWidth      =   195
            TabIndex        =   69
            Top             =   1440
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.TextBox txtDNames 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            Height          =   285
            Left            =   6000
            TabIndex        =   68
            Top             =   1440
            Visible         =   0   'False
            Width           =   3375
         End
         Begin VB.TextBox txtTCode 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            Height          =   285
            Left            =   4920
            TabIndex        =   67
            Top             =   1440
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.CheckBox chkDebtno 
            Caption         =   "Paid by Debtor"
            Height          =   195
            Left            =   5880
            TabIndex        =   66
            Top             =   1200
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.CommandButton cmddebitcreditnote 
            Caption         =   "Debit/Credit "
            Height          =   315
            Left            =   2880
            TabIndex        =   65
            Top             =   6480
            Width           =   1215
         End
         Begin VB.CheckBox chkCreditors 
            Caption         =   "Debtors"
            Height          =   255
            Index           =   0
            Left            =   480
            TabIndex        =   63
            Top             =   2820
            Width           =   1095
         End
         Begin VB.CommandButton cmdClear 
            Caption         =   "Clear<<<"
            Height          =   345
            Index           =   0
            Left            =   7200
            TabIndex        =   49
            Top             =   3360
            Width           =   930
         End
         Begin VB.CommandButton cmdAcctsSearch 
            Height          =   300
            Index           =   0
            Left            =   1485
            Picture         =   "frmNominals.frx":073C
            Style           =   1  'Graphical
            TabIndex        =   48
            Top             =   3330
            Width           =   330
         End
         Begin VB.CommandButton cmdRemove 
            Caption         =   "<<Remove"
            Height          =   345
            Index           =   0
            Left            =   6240
            TabIndex        =   47
            Top             =   3360
            Width           =   930
         End
         Begin VB.CommandButton cmdAdd 
            Caption         =   "Add>>"
            Height          =   345
            Index           =   0
            Left            =   5280
            TabIndex        =   46
            Top             =   3360
            Width           =   930
         End
         Begin VB.TextBox txtAccNames 
            Height          =   315
            Index           =   0
            Left            =   1815
            TabIndex        =   45
            Top             =   3330
            Width           =   3225
         End
         Begin VB.ComboBox cboAccno 
            Height          =   315
            Index           =   0
            Left            =   315
            Style           =   2  'Dropdown List
            TabIndex        =   44
            Top             =   3330
            Width           =   1200
         End
         Begin VB.ComboBox cboBanks 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   0
            ItemData        =   "frmNominals.frx":083E
            Left            =   2160
            List            =   "frmNominals.frx":0840
            TabIndex        =   43
            Top             =   750
            Width           =   1350
         End
         Begin VB.CommandButton cmdBank 
            Caption         =   "<>"
            Height          =   300
            Index           =   0
            Left            =   3465
            TabIndex        =   42
            Top             =   765
            Width           =   345
         End
         Begin VB.CommandButton cmdupdatereceipt 
            Caption         =   "&Post"
            Height          =   375
            Index           =   0
            Left            =   255
            TabIndex        =   41
            Top             =   6480
            Width           =   1425
         End
         Begin VB.TextBox txtmode 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   0
            Left            =   3180
            TabIndex        =   40
            Top             =   1410
            Width           =   1380
         End
         Begin VB.CommandButton cmdReceipt 
            Caption         =   "<>"
            Height          =   300
            Index           =   0
            Left            =   7515
            TabIndex        =   39
            Top             =   1800
            Width           =   345
         End
         Begin VB.ComboBox cboMode 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   0
            ItemData        =   "frmNominals.frx":0842
            Left            =   1680
            List            =   "frmNominals.frx":0855
            TabIndex        =   38
            Text            =   "Cash"
            Top             =   1395
            Width           =   1425
         End
         Begin VB.TextBox txtAmountPaid 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   0
            Left            =   6360
            MaxLength       =   9
            TabIndex        =   37
            Text            =   "0"
            Top             =   2280
            Width           =   1380
         End
         Begin VB.TextBox txtReceiptsno 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   0
            Left            =   6120
            TabIndex        =   36
            Top             =   1800
            Width           =   1380
         End
         Begin VB.TextBox txtBalance 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   0
            Left            =   7950
            Locked          =   -1  'True
            TabIndex        =   35
            Text            =   "0"
            Top             =   2775
            Width           =   1380
         End
         Begin VB.TextBox txtDistributed 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   0
            Left            =   7920
            Locked          =   -1  'True
            MaxLength       =   20
            TabIndex        =   34
            Text            =   "0"
            Top             =   2295
            Width           =   1380
         End
         Begin VB.TextBox txtParticulars 
            Appearance      =   0  'Flat
            Height          =   375
            Index           =   0
            Left            =   1680
            TabIndex        =   33
            Top             =   1725
            Width           =   3225
         End
         Begin VB.TextBox txtPayee 
            Appearance      =   0  'Flat
            Height          =   375
            Index           =   0
            Left            =   1680
            TabIndex        =   32
            Top             =   2280
            Width           =   3225
         End
         Begin MSComctlLib.ListView lvwNtrans 
            Height          =   2580
            Index           =   0
            Left            =   270
            TabIndex        =   50
            Top             =   3795
            Width           =   9060
            _ExtentX        =   15981
            _ExtentY        =   4551
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
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   4
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Code"
               Object.Width           =   3528
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Trans Description"
               Object.Width           =   5292
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   2
               Text            =   "Amount"
               Object.Width           =   3175
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Accno"
               Object.Width           =   2540
            EndProperty
         End
         Begin MSComCtl2.DTPicker dtpTransDate 
            Height          =   375
            Index           =   0
            Left            =   6195
            TabIndex        =   51
            Top             =   240
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   661
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   176685057
            CurrentDate     =   40421
         End
         Begin VB.Label Label4 
            Caption         =   "DR A/C"
            Height          =   255
            Left            =   1800
            TabIndex        =   72
            Top             =   3120
            Width           =   975
         End
         Begin VB.Label Label3 
            Caption         =   "CR A/C"
            Height          =   255
            Left            =   2400
            TabIndex        =   71
            Top             =   480
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "Code"
            Height          =   255
            Left            =   5040
            TabIndex        =   70
            Top             =   1200
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.Label lblSupplier 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   0
            Left            =   1680
            TabIndex        =   64
            Top             =   2760
            Width           =   3165
         End
         Begin VB.Label lblbankname 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Index           =   0
            Left            =   3810
            TabIndex        =   60
            Top             =   765
            Width           =   4095
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "Bank A/C (Source)"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   450
            TabIndex        =   59
            Top             =   810
            Width           =   1485
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "Payment Mode"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   420
            TabIndex        =   58
            Top             =   1440
            Width           =   1230
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Received Amount"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   6360
            TabIndex        =   57
            Top             =   2085
            Width           =   1455
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Receipt No"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   0
            Left            =   5085
            TabIndex        =   56
            Top             =   1800
            Width           =   870
         End
         Begin VB.Label Label11 
            Caption         =   "Distributed"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   0
            Left            =   8085
            TabIndex        =   55
            Top             =   2085
            Width           =   1005
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Particulars"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   0
            Left            =   375
            TabIndex        =   54
            Top             =   1875
            Width           =   885
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "Payee"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   0
            Left            =   465
            TabIndex        =   53
            Top             =   2400
            Width           =   495
         End
         Begin VB.Label Label8 
            Caption         =   "Transaction Date"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   0
            Left            =   4560
            TabIndex        =   52
            Top             =   360
            Width           =   1590
         End
      End
      Begin VB.Frame Frame2 
         Height          =   6960
         Index           =   1
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   9510
         Begin VB.CheckBox chkCreditors 
            Caption         =   "Creditors"
            Height          =   255
            Index           =   1
            Left            =   480
            TabIndex        =   61
            Top             =   2820
            Width           =   1095
         End
         Begin VB.CommandButton cmdRemove 
            Caption         =   "<<Remove"
            Height          =   345
            Index           =   1
            Left            =   6240
            TabIndex        =   19
            Top             =   3360
            Width           =   930
         End
         Begin VB.TextBox txtPayee 
            Appearance      =   0  'Flat
            Height          =   375
            Index           =   1
            Left            =   1680
            TabIndex        =   18
            Top             =   2280
            Width           =   3225
         End
         Begin VB.TextBox txtParticulars 
            Appearance      =   0  'Flat
            Height          =   375
            Index           =   1
            Left            =   1680
            TabIndex        =   17
            Top             =   1725
            Width           =   3225
         End
         Begin VB.TextBox txtDistributed 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   1
            Left            =   7920
            TabIndex        =   16
            Text            =   "0"
            Top             =   2295
            Width           =   1380
         End
         Begin VB.TextBox txtBalance 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   1
            Left            =   7950
            TabIndex        =   15
            Text            =   "0"
            Top             =   2775
            Width           =   1380
         End
         Begin VB.TextBox txtReceiptsno 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   1
            Left            =   6120
            TabIndex        =   14
            Top             =   1425
            Width           =   1380
         End
         Begin VB.TextBox txtAmountPaid 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   1
            Left            =   6375
            TabIndex        =   13
            Text            =   "0"
            Top             =   2310
            Width           =   1380
         End
         Begin VB.ComboBox cboMode 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   1
            ItemData        =   "frmNominals.frx":0878
            Left            =   1680
            List            =   "frmNominals.frx":088B
            TabIndex        =   12
            Text            =   "Cash"
            Top             =   1395
            Width           =   1425
         End
         Begin VB.CommandButton cmdReceipt 
            Caption         =   "<>"
            Height          =   300
            Index           =   1
            Left            =   7515
            TabIndex        =   11
            Top             =   1410
            Width           =   345
         End
         Begin VB.TextBox txtmode 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   1
            Left            =   3180
            TabIndex        =   10
            Top             =   1410
            Width           =   1380
         End
         Begin VB.CommandButton cmdupdatereceipt 
            Caption         =   "&Post"
            Height          =   375
            Index           =   1
            Left            =   255
            TabIndex        =   9
            Top             =   6465
            Width           =   1425
         End
         Begin VB.CommandButton cmdBank 
            Caption         =   "<>"
            Height          =   300
            Index           =   1
            Left            =   3465
            TabIndex        =   8
            Top             =   765
            Width           =   345
         End
         Begin VB.ComboBox cboBanks 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   1
            ItemData        =   "frmNominals.frx":08AE
            Left            =   2160
            List            =   "frmNominals.frx":08B0
            TabIndex        =   7
            Top             =   750
            Width           =   1350
         End
         Begin VB.ComboBox cboAccno 
            Height          =   315
            Index           =   1
            Left            =   315
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   3330
            Width           =   1200
         End
         Begin VB.TextBox txtAccNames 
            Height          =   315
            Index           =   1
            Left            =   1815
            TabIndex        =   5
            Top             =   3330
            Width           =   3225
         End
         Begin VB.CommandButton cmdAdd 
            Caption         =   "Add>>"
            Height          =   345
            Index           =   1
            Left            =   5280
            TabIndex        =   4
            Top             =   3360
            Width           =   930
         End
         Begin VB.CommandButton cmdAcctsSearch 
            Height          =   300
            Index           =   1
            Left            =   1485
            Picture         =   "frmNominals.frx":08B2
            Style           =   1  'Graphical
            TabIndex        =   3
            Top             =   3330
            Width           =   330
         End
         Begin VB.CommandButton cmdClear 
            Caption         =   "Clear<<<"
            Height          =   345
            Index           =   1
            Left            =   7200
            TabIndex        =   2
            Top             =   3360
            Width           =   930
         End
         Begin MSComctlLib.ListView lvwNtrans 
            Height          =   2580
            Index           =   1
            Left            =   270
            TabIndex        =   20
            Top             =   3795
            Width           =   9060
            _ExtentX        =   15981
            _ExtentY        =   4551
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
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   4
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Code"
               Object.Width           =   3528
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Trans Description"
               Object.Width           =   5292
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   2
               Text            =   "Amount"
               Object.Width           =   3175
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Accno"
               Object.Width           =   2540
            EndProperty
         End
         Begin MSComCtl2.DTPicker dtpTransDate 
            Height          =   375
            Index           =   1
            Left            =   6195
            TabIndex        =   21
            Top             =   240
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   661
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   176750593
            CurrentDate     =   40421
         End
         Begin VB.Label Label5 
            Caption         =   "DR A/C"
            Height          =   375
            Left            =   1920
            TabIndex        =   73
            Top             =   3120
            Width           =   1095
         End
         Begin VB.Label lblSupplier 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   1
            Left            =   1680
            TabIndex        =   62
            Top             =   2760
            Width           =   3165
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "Payee"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   1
            Left            =   465
            TabIndex        =   30
            Top             =   2400
            Width           =   495
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Particulars"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   375
            TabIndex        =   29
            Top             =   1875
            Width           =   885
         End
         Begin VB.Label Label7 
            Caption         =   "Distributed"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Left            =   8085
            TabIndex        =   28
            Top             =   2085
            Width           =   1005
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Voucher No"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   1
            Left            =   5085
            TabIndex        =   27
            Top             =   1455
            Width           =   960
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Paid Amount"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   1
            Left            =   6360
            TabIndex        =   26
            Top             =   2085
            Width           =   1050
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Payment Mode"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   1
            Left            =   420
            TabIndex        =   25
            Top             =   1440
            Width           =   1230
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Bank (CR)"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   450
            TabIndex        =   24
            Top             =   810
            Width           =   795
         End
         Begin VB.Label lblbankname 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Index           =   1
            Left            =   3810
            TabIndex        =   23
            Top             =   765
            Width           =   4095
         End
         Begin VB.Label Label8 
            Caption         =   "Transaction Date"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   1
            Left            =   4560
            TabIndex        =   22
            Top             =   360
            Width           =   1590
         End
      End
   End
   Begin VB.Menu mnuReports 
      Caption         =   "Reports"
      Begin VB.Menu mnuCustStatements 
         Caption         =   "Customer Statements"
      End
      Begin VB.Menu mnuSupStatement 
         Caption         =   "Supplier Statements"
      End
      Begin VB.Menu mnuSalesDayBook 
         Caption         =   "Sales Day Book"
      End
      Begin VB.Menu mnuPurDayBook 
         Caption         =   "Purchase Day Book"
      End
      Begin VB.Menu mnureceiptslist 
         Caption         =   "Receipts"
      End
   End
End
Attribute VB_Name = "frmNominals"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Totalamount As Currency
Dim pushed As Currency
'Dim objLabelEdit As LabelEdit
'Dim objLabelEdit2 As LabelEdit
'Dim objLabelEdit3 As LabelEdit
Dim interestAcc As String, LoanAcc As String
Dim IsBatch As Boolean, IsGroup As Boolean, isfixed As Boolean
Dim k As Integer
Dim shareBal As Double, balance As Double
Dim daysIntoTheMonth As Integer
Dim Posted As Boolean
Dim ref As String, RefNo As String
Dim penalise As Boolean
Dim commit As Boolean
Dim rsDefaults As ADODB.Recordset
Dim fSeed As Integer
Dim newMember As Boolean



Private Sub cboAccno_Change(Index As Integer)
    Dim ACCNO As String
    ACCNO = cboAccno(Index).Text
    sql = "select GLACCNAME from glsetup where accno='" & ACCNO & "'"
    Set rs = oSaccoMaster.GetRecordset(sql)
    If Not rs.EOF Then
        txtAccNames(Index).Text = rs(0)
    Else
        txtAccNames(Index).Text = ""
    End If
End Sub

Private Sub cboAccno_Click(Index As Integer)
    cboAccno_Change (Index)
End Sub



Private Sub cboBanks_Change(Index As Integer)
    On Error GoTo SysError
    Dim rsGlaccname As New ADODB.Recordset
'    sql = "select bankname from banks where Accno='" & cboBanks(index).Text & "'"
'    Set rst = oSaccoMaster.GetRecordset(sql)
'    If Not rst.EOF Then
'        lblbankname(index).Caption = rst(0)
'    Else
'        lblbankname(index).Caption = ""
'    End If
'    Exit Sub
    Dim ACCNO As String
    ACCNO = cboBanks(Index).Text
    sql = "select Glaccname from glsetup where accno='" & ACCNO & "'"
    Set rsGlaccname = oSaccoMaster.GetRecordset(sql)
    If Not rsGlaccname.EOF Then
        lblbankname(Index).Caption = rsGlaccname("GLACCNAME")
    Else
       lblbankname(Index).Caption = ""
    End If
    Exit Sub

SysError:
    MsgBox err.description, vbInformation, Me.Caption
End Sub

Private Sub cboBanks_Click(Index As Integer)
    cboBanks_Change (Index)
End Sub




Private Sub cboMode_Change(Index As Integer)
'On Error Resume Next

    Select Case cboMode(Index).Text
        Case "Cash"
            txtmode(Index).Visible = False
            txtReceiptsno(Index).SetFocus
        Case "Cheque"
            txtmode(Index).Visible = True
            txtmode(Index).SetFocus
        Case "Direct Deposit"
            txtmode(Index).Visible = True
            txtmode(Index).SetFocus
        Case Else

    End Select
End Sub

Private Sub cboMode_Click(Index As Integer)
    cboMode_Change (Index)
End Sub





Private Sub chkCreditors_Click(Index As Integer)
    With chkCreditors(Index)
        If Index = 0 Then
            If .value = 1 Then
                frmSearchDebtors.Show vbModal
                sql = "Select dname companyName from d_debtors where dcode='" & sel & "'"
            End If
        Else
            If .value = 1 Then
                frmSearchVendor.Show vbModal
                sql = "Select companyName from ag_supplier1 where supplierId='" & sel & "'"
            End If
        End If
        
        
        If .value = 1 Then
            Set rst = oSaccoMaster.GetRecordset(sql)
            If Not rst.EOF Then
                lblSupplier(Index).Caption = rst!CompanyName
            Else
                lblSupplier(Index).Caption = ""
            End If
        Else
            InvoiceNo = ""
            InvoiceBal = 0#
            lblSupplier(Index).Caption = ""
        End If
    End With
End Sub

'Private Sub chkDebtno_Click()
'If chkDebtno.value = vbChecked Then
'    txtTCode.Visible = True
'    Picture5.Visible = True
'    txtDNames.Visible = True
'    Label1.Visible = True
'    txtDNames = "<select Debtors Code here>"
'
'
''    Set rsy = New ADODB.Recordset
''   ' dtpTransDate = frmMilkControl.ListView2.SelectedItem(3)
''    txtTCode = frmMilkControl.ListView2.SelectedItem
''    txtTCode_Validate True
''    'txtDNames = frmMilkControl.ListView2.SelectedItem(0)
''    sql = ""
''    sql = "set dateformat dmy SELECT     d.AccNo, d.GlAccName FROM GLSETUP AS d INNER JOIN d_MilkControl AS m ON d.AccNo = m.CreditAcc WHERE     (DCode = '" & txtTCode & "')"
''    Set rsy = oSaccoMaster.GetRecordset(sql)
''     If Not rsy.EOF Then
''     cboAccno(0).Text = rsy(0)
''     cboBanks(0).Text = "A004"
''     cboMode(0).Text = "Cash"
''     txtParticulars(0).Text = "CASH MILK PAYMENTS"
''    ' txtTCode(0).Text = rsy(0)
''     Else
''     End If
'
'Else
'    txtTCode.Visible = True
'    Picture5.Visible = True
'    txtDNames.Visible = True
'    Label1.Visible = False
''    txtDebt.Visible = False
'
'End If
'
'End Sub

Private Sub cmdAcctsSearch_Click(Index As Integer)
    On Error Resume Next
    frmAcctsSearch.Show vbModal, Me
    If Continue Then
        If SearchValue <> "" Then
            cboAccno(Index) = SearchValue
            SearchValue = ""
            Continue = False
        End If
    End If
End Sub

Private Sub cmdAdd_Click(Index As Integer)
    On Error GoTo SysError
    Dim ans As String
    ans = MsgBox("Is the payment date Correct ?", vbYesNo)
    If ans = vbNo Then
      Exit Sub
    Else
    End If
    If cboAccno(Index).Text = "" Then
        Exit Sub
    End If

    Set li = lvwNtrans(Index).ListItems.Add(, , cboAccno(Index))
    li.SubItems(1) = txtAccNames(Index)
    li.SubItems(2) = InvoiceBal
   
    Recalculate (Index)
    Exit Sub
   
SysError:
    MsgBox err.description, vbInformation, Me.Caption
End Sub

Private Sub cmdBank_Click(Index As Integer)
    On Error Resume Next
    frmAcctsSearch.Show vbModal, Me
    If Continue Then
        If SearchValue <> "" Then
            cboBanks(Index) = SearchValue
            SearchValue = ""
            Continue = False
        End If
    End If
 End Sub

Private Sub cmdclose_Click()
    Unload Me
End Sub

Private Sub PrintReceipt(Index As Integer)
    On Error GoTo SysError
    
    
    If Index = 0 Then
        reportname = "Receipt.rpt"
        STRFORMULA = "{VwReceipt.ReceiptNo}='" & txtReceiptsno(Index) & "'"
    Else
        reportname = "Voucher.rpt"
        STRFORMULA = "{paymentBooking.VoucherNo}='" & txtReceiptsno(Index) & "'"
    End If
    Show_Sales_Crystal_Report STRFORMULA, reportname, CompanyName
    
    Exit Sub

    Dim pay, tot, disc As Currency
    Dim Z, X As Integer
    'for number of copies
    Dim a As Integer
    Dim b As Integer
    Dim C As Integer
    dlg9.CancelError = True
    dlg9.FontName = "Garamond"
    Dim j As Printer

    a = dlg9.Copies
    Printer.CurrentY = 500
    Printer.CurrentX = 9000
    Printer.FontSize = 11
    Printer.CurrentY = 600
    Printer.CurrentX = 1000
    Printer.Print Tab(4); "TRANSACTION VOUCHER"
    Printer.Print Tab(0); "---------------------------------------"
    Printer.Print Tab(0); CompanyName
    'Printer.Print Tab(0); CompanyPhone
    'Printer.Print Tab(0); CompanyTown
    Printer.Print Tab(0); "---------------------------------------"
    Printer.Print Tab(0); "Document No"; Tab(20); txtReceiptsno(Index).Text
    Printer.Print Tab(0); "Date :  "; dtpTransDate(0).value
    Printer.Print Tab(0); "Source"; Tab(10); lblbankname(Index)

    Printer.CurrentX = 500#
    Printer.FontSize = 10
    Printer.CurrentX = 500
    Printer.FontSize = 8
    
    Printer.CurrentX = 500
    Printer.FontSize = 8
    Printer.Print
    Printer.CurrentX = 500
    Printer.FontSize = 8
    Printer.Print
    Printer.Print Tab(2); "Payment Type"; Tab(30); cboMode(0).Text
    Printer.Print Tab(0); "-------------------------------------------------"
    If Index = 0 Then
        Printer.Print "Total Received :"; Tab(20); txtDistributed(Index)
    Else
        Printer.Print "Total Paid :"; Tab(20); txtDistributed(Index)
    End If
    Printer.Print "Particulars :"; Tab(10); txtParticulars(Index)
    Printer.Print Tab(0); "------------------------------------------------"
    Printer.Print Tab(2); "ITEM"; Tab(30); "AMOUNT"
    Printer.Print Tab(0); "------------------------------------------------"
    Printer.FontSize = 8
    For I = 1 To lvwNtrans(Index).ListItems.Count
        Printer.Print lvwNtrans(Index).ListItems(I).ListSubItems(1); Tab(30); lvwNtrans(Index).ListItems(I).ListSubItems(2)
    Next I

    'Printer.Print "Your Balance is :"; Tab(20); asa
    Printer.Print Tab(0); "------------------------------------------------"
    Printer.Print Tab(0); "Document No: "; txtReceiptsno(0).Text
    Printer.Print Tab(2); "You were Served by: " & User
    Printer.Print Tab(2); "Receipient Signature   /Thumb Print"
    Printer.Print
    Printer.Print Tab(0); "Remarks"
    Printer.Print
    Printer.Print Tab(0); "---------------------------------------"
    'Printer.Print Tab(0); CompanyTagLine
    Printer.Print
    Printer.Print
    Printer.Print Tab(0); "POWERED BY EASYMA"
    Printer.Print
    Printer.EndDoc

'//-------------------------
'mysql = ""
'mysql = "set dateformat dmy Insert into Receiptno(Receiptno,Auditdate,auditid,memberno)values('" & txtReceiptsno & "','" & Format(Get_Server_Date, "dd/MM/yyyy") & "','" & User & "','" & txtMemberNo & "')"
'oSaccoMaster.ExecuteThis (mysql)
    Exit Sub
SysError:
    MsgBox err.description, vbInformation
End Sub
Private Sub cmddebitcreditnote_Click()
frmdebitcreditnote.Show vbModal, Me
End Sub

Private Sub cmdRemove_Click(Index As Integer)
    On Error GoTo SysError
    With lvwNtrans(Index)
        If .ListItems.Count > 0 Then
            If MsgBox("Do you want to remove " & lvwNtrans(Index).SelectedItem.SubItems(1) & _
            " From the list?", vbQuestion + vbYesNo, Me.Caption) = vbNo Then
                Exit Sub
            End If
            pushed = pushed - .SelectedItem.ListSubItems(2)
            .ListItems.Remove (.SelectedItem.Index)
        End If
    End With
    Recalculate (Index)
    Exit Sub
SysError:
    MsgBox err.description, vbInformation, Me.Caption
End Sub
Private Sub cmdupdatereceipt_Click(Index As Integer)
     On Error GoTo SysError
    Dim j As Integer
    Dim amt As Double, interest As Double
    Dim code As String, rsLoan As New Recordset
    Dim k As Long, repaymethod As String
    Dim sharebals As Double, amount As Double
    Dim ptype As String, principal As Double, chequeno As String
    Dim intWaive As Boolean
    'Recalculate
    If lvwNtrans(Index).ListItems.Count = 0 Then
        Exit Sub
    End If
    
    If txtReceiptsno(Index) = "" Then
        MsgBox "Please enter the receiptno", vbCritical
        Exit Sub
    End If
    If txtAmountPaid(Index) <= 0 Then
        MsgBox "Amount should be greater than zero", vbCritical
        Exit Sub
    End If
    If txtAmountPaid(Index) = "" Then
        MsgBox "Amount should be have a figure on it", vbCritical
        Exit Sub
    End If
    If cboBanks(Index).Text = "" Then
        MsgBox ("You Must Select the Bank Control Account"), vbCritical
        Exit Sub
    End If
    If cboMode(Index) = "Cheque" Then
        If txtmode(Index) = "" Then
            MsgBox ("Cheque Receipt No Required"), vbInformation
            Exit Sub
        End If
    End If
    If cboMode(Index) = "Cash" Then
        If txtReceiptsno(Index) = "" Then
            MsgBox ("Cash Receipt No Required"), vbInformation
            Exit Sub
        End If
    End If
    If cboMode(Index) = "EFT" Then
        If txtmode(Index) = "" Then
            MsgBox ("EFT Receipt No Required"), vbInformation
            Exit Sub
        End If
    End If
    If cboMode(Index) = "Mpesa" Then
        If txtmode(Index) = "" Then
            MsgBox ("Mpesa Receipt No Required"), vbInformation
            Exit Sub
        End If
    End If
    If cboMode(Index) = "Zap" Then
        If txtmode(Index) = "" Then
            MsgBox ("Zap Receipt No Required"), vbInformation
            Exit Sub
        End If
    End If
    If CDbl(txtBalance(Index)) <> 0 Then
        MsgBox "The Amount Received should be equal to the Amount Distributed", vbInformation, Me.Caption
        Exit Sub
    End If
    

    
    If Index = 0 Then
        sql = "select receiptno,chequeno from receiptbooking where receiptno in ('" & txtReceiptsno(Index) & "')"
        Set rst = oSaccoMaster.GetRecordset(sql)
        If Not rst.EOF Then
            MsgBox "Either the Receiptno or Chequeno is already used!", vbCritical
            Exit Sub
        End If
        
        If MsgBox("Do you want to post this Receipt: " & txtReceiptsno(0) & "?", vbQuestion + vbYesNo, Me.Caption) = vbNo Then
            Exit Sub
        End If
    
    Else
        sql = "select voucherno,chequeno from paymentbooking where voucherno in ('" & txtReceiptsno(Index) & "')"
        Set rst = oSaccoMaster.GetRecordset(sql)
        If Not rst.EOF Then
            MsgBox "Either the VoucherNo or Chequeno is already used!", vbCritical
            Exit Sub
        End If
        
        If MsgBox("Do you want to post this Voucher: " & txtReceiptsno(1) & "?", vbQuestion + vbYesNo, Me.Caption) = vbNo Then
            Exit Sub
        End If
        
    End If
'    ''...................insert the amount to debtor if available................................
'      If chkDebtno.value = vbChecked Then
'          If txtTCode = "" Then
'            MsgBox "Please enter the Debtors Code", vbCritical
'           Exit Sub
'          End If
'
'       Dim Amount1 As Integer
'       Set rs = New ADODB.Recordset
'       sql = ""
'       sql = "SET dateformat dmy Select Amount,status  from d_MilkControl  where DCode ='" & txtTCode & "' and DispDate='" & DTPTransdate(Index) & "'"
'       Set rs = oSaccoMaster.GetRecordset(sql)
'
'       If Not rs.EOF Then
''        sql = ""
''        sql = "set dateformat dmy insert into  d_MilkControl(Amount) values('" & CDbl(txtDistributed(Index)) & "') where DCode ='" & txtTCode & "' and DispDate='" & DTPTransdate(Index) & "'"
''        oSaccoMaster.ExecuteThis (sql)
''
''       Else
'         sql = ""
'         sql = "set dateformat DMY update d_MilkControl set Amount=" & rs.Fields("amount") + CDbl(txtDistributed(Index)) & " , status=1 where DCode ='" & txtTCode & "' and DispDate='" & DTPTransdate(Index) & "' "
'         oSaccoMaster.ExecuteThis (sql)
'       End If
'     Else
'     End If
'
'    '''..................end of debtor...........................................................
    
    
    Select Case txtmode(Index).Text
        Case "Cheque"
            chequeno = txtmode(Index)
        Case Else
            chequeno = txtmode(Index).Text
    End Select
    
    
    
    '***********************************BEGINNING OF POSTING**************************************

    ref = "Nominal"
    RefNo = ""

    
    With lvwNtrans(Index)
        If .ListItems.Count < 0 Then
            Exit Sub
        End If
        I = 0
        j = 0
        I = .ListItems.Count
        
        oSaccoMaster.goConn.BeginTrans
        On Error GoTo TransError
        'save TransactionNo
        transactionTotal = CDbl(txtAmountPaid(Index).Text)
        NewTransaction transactionTotal, dtpTransDate(Index), "ReceiptPosting -Receiptno " & txtReceiptsno(Index).Text
            
            Dim fperiod As Integer, intRate As Double

            If Index = 0 Then
                If Not saveReceipt(txtReceiptsno(Index), ref, RefNo, cboBanks(Index).Text, cboBanks(Index).Text, txtParticulars(Index), dtpTransDate(Index).value, CDbl(txtAmountPaid(Index).Text), txtmode(Index), cboMode(Index).Text, "Deposit") Then
                    GoTo TransError
                Else
                    For I = 1 To lvwNtrans(Index).ListItems.Count
                    
                        If Save_GLTRANSACTION(dtpTransDate(Index), lvwNtrans(Index).ListItems(I).SubItems(2), cboBanks(Index), lvwNtrans(Index).ListItems(I), _
                        txtReceiptsno(Index), mMemberNo, User, ErrorMessage, txtParticulars(Index) + "-" + Replace(lblSupplier(Index), ",", ""), 0, 1, txtmode(Index), transactionNo) = False Then
                            GoTo TransError
                        End If

                    Next I
                    
                    'Invoice
                    
                    If InvoiceNo <> "" And InvoiceBal > 0 And chkCreditors(0).value = 1 Then
                    
                        sql = "update SalesOrder set balance=balance-" & txtAmountPaid(1) & " where orderno='" & InvoiceNo & "'"
                        If Not oSaccoMaster.Execute(sql) Then
                            GoTo TransError
                        End If
                    
                        'Statement
                        
                        sql = "insert into CustomerStmt (invId,TransDate,Refno,Amount,TransType,Balance,Auditid,Remarks,Transactionno) " _
                        & " Values ('" & InvoiceNo & "','" & dtpTransDate(Index) & "','" & txtReceiptsno(Index) & "'," & txtAmountPaid(Index) & ",'CR',0,'" & User & "','" & txtParticulars(Index) & "','" & transactionNo & "')"
                        
                        If Not oSaccoMaster.Execute(sql) Then
                            GoTo TransError
                        End If
                        
                    End If
                    
                    
                End If
            Else
            
            
                    For I = 1 To lvwNtrans(Index).ListItems.Count
                    
                        If Save_GLTRANSACTION(dtpTransDate(Index), lvwNtrans(Index).ListItems(I).SubItems(2), lvwNtrans(Index).ListItems(I), cboBanks(Index), txtReceiptsno(Index), mMemberNo, User, ErrorMessage, txtParticulars(Index), 1, 1, txtmode(Index), TransNo) = False Then

                            GoTo TransError
                        End If
'           If Save_GLTRANSACTION(dtpTransDate(index), lvwNtrans(index).ListItems(I).SubItems(2), cboBanks(index), lvwNtrans(index).ListItems(I), _
'                        txtReceiptsno(index), mMemberNo, User, ErrorMessage, txtParticulars(index) + "-" + Replace(lblSupplier(index), ",", ""), 0, 1, txtmode(index), transactionNo) = False Then
'                            GoTo TransError
'                        End If
                    Next I
                    'Save the Voucher

                    sql = "set dateformat dmy INSERT INTO PaymentBooking (VoucherNo,Memberno,PayeeDesc,Ccode,Name,Transdate," _
                    & "Amount, Chequeno, Ptype, auditid,Transactionno) VALUES ('" & txtReceiptsno(Index) & "','" & _
                    cboBanks(Index) & "','" & lblbankname(Index) & "','" & cboBanks(Index) & "','" & lblbankname(Index).Caption & "','" & dtpTransDate(Index) & "'," & CDbl(txtDistributed(Index)) & ",'" & _
                    txtmode(Index).Text & "','" & cboMode(Index).Text & "','" & User & "','" & transactionNo & "')"
                    
                    If Not oSaccoMaster.Execute(sql) Then
                        GoTo TransError
                    End If
                    
                    
                    'Invoice
                    
                    If InvoiceNo <> "" And InvoiceBal > 0 And chkCreditors(Index).value = 1 Then
                    
                        sql = "update d_Invoice set balance=balance-" & txtAmountPaid(1) & " where invid='" & InvoiceNo & "'"
                        If Not oSaccoMaster.Execute(sql) Then
                            GoTo TransError
                        End If
                    
                        'Statement
                        
                        sql = "insert into supplierStmt (invId,TransDate,Refno,Amount,TransType,Balance,Auditid,Remarks,Transactionno) " _
                        & " Values ('" & InvoiceNo & "','" & dtpTransDate(Index) & "','" & txtReceiptsno(Index) & "'," & txtAmountPaid(Index) & ",'DR',0,'" & User & "','" & txtParticulars(Index) & "','" & transactionNo & "')"
                        
                        If Not oSaccoMaster.Execute(sql) Then
                            GoTo TransError
                        End If
                        
                    End If
                    
                        
                    
                    
            End If
        End With
    oSaccoMaster.goConn.CommitTrans
    
'    If MsgBox("Receipt updated successfully, Print Receipt?", vbQuestion + vbYesNo) = vbYes Then
'        PrintReceipt Index
'    End If
    MsgBox "Record Posted succesfully", vbCritical
    chkDebtno.value = vbUnchecked
    
    lvwNtrans(Index).ListItems.Clear
    
    txtReceiptsno(0) = NewRefno(0) 'Generate_ReceiptNo("Receipt")
    txtReceiptsno(1) = NewRefno(1) 'Generate_ReceiptNo("Voucher")
    
    Exit Sub
SysError:
    MsgBox IIf(ErrorMessage = "", err.description, ErrorMessage), vbCritical, Me.Caption
    GoTo TransError
    Exit Sub
TransError:
    If err.number = 35600 Then
        Resume Next
    End If
    If ErrorMessage = "" And err.description = "" Then
        oSaccoMaster.goConn.RollbackTrans
    Else
        MsgBox IIf(ErrorMessage = "", err.description, ErrorMessage) & vbNewLine & "Action Therefore Aborted. ", vbCritical, Me.Caption
        oSaccoMaster.goConn.RollbackTrans
    End If
End Sub
Private Sub cmdClear_Click(Index As Integer)
    On Error Resume Next
    Dim vote As String
    'Dim msgb As Response
    Dim rsponse As String

        
    With lvwNtrans(Index)
        If .ListItems.Count > 0 Then
            If MsgBox("Are you sure you want to clear the entire list?", vbQuestion + vbYesNo) = vbYes Then
                .ListItems.Clear
            End If
        End If
    End With

    Recalculate (Index)
        
End Sub

Private Sub Form_Load()
    On Error Resume Next
    dtpTransDate(0) = Format(Get_Server_Date, "dd/mm/yyyy")
    dtpTransDate(1) = Format(Get_Server_Date, "dd/mm/yyyy")
    
    Dim rscompany As New ADODB.Recordset
    InvoiceBal = 0

   'Load Gl's
    sql = "Select accno from glsetup order by accno asc"
    Set rst = oSaccoMaster.GetRecordset(sql)
    While Not rst.EOF
        cboAccno(0).AddItem (rst(0))
        cboAccno(1).AddItem (rst(0))
        rst.MoveNext
    Wend
    'load the banks
    cboBanks(0).Clear
    cboBanks(1).Clear
    sql = "select Accno from banks where accno not in (select assigngl from useraccounts where UserLoginIDs not in('" & User & "') and assigngl<>'' and assigngl is not null )"
    Set rst = oSaccoMaster.GetRecordset(sql)
    While Not rst.EOF
        cboBanks(0).AddItem rst(0)
        cboBanks(1).AddItem rst(0)
        rst.MoveNext
    Wend
    
    cboBanks(0).List(0) = currentUser.tellerGlAcc
    cboBanks(0).Text = cboBanks(0).List(0)
    cboBanks(1).List(0) = currentUser.tellerGlAcc
    cboBanks(1).Text = cboBanks(1).List(0)
    
   
    Totalamount = 0
    pushed = 0

    txtReceiptsno(0) = NewRefno(0) 'Generate_ReceiptNo("Receipt")
    txtReceiptsno(1) = NewRefno(1) 'Generate_ReceiptNo("Voucher")
 
'    lvwJuniors.Visible = False
    IsBatch = False
    commit = False 'will be used to determine whether a loaded interest is saved as a transaction
''    ACCNO = "O014"
''    cmdAcctsSearch_Click True
    'cboAccno(Index).Text = "O014"
End Sub

Public Function NewRefno(Index As Integer) As String
    
    Dim Rno As String
    Dim prefix As String
    Dim rcount As Integer
    If Index = 0 Then
        sql = "select count(distinct receiptno)ccount from receiptbooking"
        prefix = "MCR-"
    Else
        sql = "select count(distinct voucherno)ccount from paymentbooking"
        prefix = "MCV-"
    End If
        
    
    Set rst = oSaccoMaster.GetRecordset(sql)
    If Not rst.EOF Then
        rcount = rst(0) + 1
    Else
        rcount = 1
    End If
    
    'thisday = Get_Server_Date
    
    Rno = Format(CStr(rcount), "000000")
    
    
    NewRefno = prefix & "-" & CStr(Rno)
    
End Function
Private Sub Form_Unload(Cancel As Integer)
    'Stop subclassing
    CloseSubClass
    'Clean up by setting the classes to Nothing
    'Set objLabelEdit = Nothing
    'Set objLabelEdit2 = Nothing
End Sub






Private Sub frmglshow_Click()
GlinqueryTransaction.Show vbModal
End Sub

Private Sub lvwNTrans_DblClick(Index As Integer)
    Dim total As Double, amt As Double
    Dim ccount As Integer
    On Error Resume Next
    total = 0
      
    With lvwNtrans(Index)
        If .ListItems.Count > 0 Then
            amt = CDbl(txtAmountPaid(Index).Text)
            If amt = 0 Then
            amt = InputBox("Enter the amount", "AMOUNT", .SelectedItem.ListSubItems(2))
            End If
            .SelectedItem.ListSubItems(2) = amt
        End If
    End With
    Recalculate (Index)
End Sub









Private Sub Recalculate(Index As Integer)
    On Error Resume Next
    Dim balance As Double
    Dim thislistview As ListView
    
    If lvwNtrans(Index).ListItems.Count > 0 Then
        For I = 1 To lvwNtrans(Index).ListItems.Count
            balance = balance + CDbl(lvwNtrans(Index).ListItems(I).SubItems(2))
        Next I
    End If

    txtDistributed(Index) = Format(balance, Cfmt)
    txtBalance(Index) = Format(CDbl(txtAmountPaid(Index)) - CDbl(txtDistributed(Index)), Cfmt)
End Sub


Private Sub lvwNtrans_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
      lvwNTrans_DblClick Index
    End If
End Sub

Private Sub mnuCustStatements_Click()
    reportname = "CustomerStatement.rpt"
    Show_Sales_Crystal_Report "", reportname, CompanyName
End Sub

Private Sub mnuPurDayBook_Click()
    reportname = "PeriodicPurchases.rpt"
    Show_Sales_Crystal_Report "", reportname, CompanyName
End Sub

Private Sub mnureceiptslist_Click()
  reportname = "cashandbanklists.rpt"
    Show_Sales_Crystal_Report "", reportname, CompanyName
End Sub

Private Sub mnuSalesDayBook_Click()
    reportname = "PeriodicSales.rpt"
    Show_Sales_Crystal_Report "", reportname, CompanyName
End Sub

Private Sub mnuSupStatement_Click()
    reportname = "SupplierStatement.rpt"
    Show_Sales_Crystal_Report "", reportname, CompanyName
End Sub

Private Sub Picture5_Click()
Me.MousePointer = vbHourglass
        frmSearchDebtors.Show vbModal
        txtTCode = sel
        txtTCode_Validate True
        Me.MousePointer = 0
End Sub

Private Sub txtAmountPaid_Change(Index As Integer)
On Error Resume Next
''    Dim ans As String
''    ans = MsgBox("Is the payment date Correct ?", vbYesNo)
''    If ans = vbNo Then
''      Exit Sub
''    Else
''    End If
    If txtAmountPaid(Index).Text = "" Then txtAmountPaid(Index).Text = 0
    Totalamount = CDbl(txtAmountPaid(Index).Text)
    pushed = 0
    txtBalance(Index).Text = Totalamount - CDbl(txtDistributed(Index).Text)
   ' Recalculate
   'Recalculate (Index)
End Sub




Private Sub txtAmountPaid_KeyPress(Index As Integer, KeyAscii As Integer)
    If keyIsValid(KeyAscii, 1) = False Then
        Beep
        KeyAscii = 0
        Exit Sub
    End If
End Sub

Private Sub txtTCode_LinkOpen(Cancel As Integer)
        txtTCode_Validate True
        Me.MousePointer = 0
End Sub

Private Sub txtTCode_Validate(Cancel As Boolean)
Dim a As Boolean, b As Integer
Set rs = New ADODB.Recordset
sql = "d_sp_Selectdebtors '" & txtTCode & "'"
Set rs = oSaccoMaster.GetRecordset(sql)
If Not rs.EOF Then
If Not IsNull(rs.Fields(0)) Then txtDNames = rs.Fields(0)
'If Not IsNull(rs.Fields(1)) Then txtId = rs.Fields(1)
'If Not IsNull(rs.Fields(2)) Then cboLocation = rs.Fields(2)
'If Not IsNull(rs.Fields(3)) Then DTPRegDate = rs.Fields(3)
'If Not IsNull(rs.Fields(4)) Then txtEMail = rs.Fields(4)
'If Not IsNull(rs.Fields(5)) Then txtPhone = rs.Fields(5)
'If Not IsNull(rs.Fields(6)) Then txtTown = rs.Fields(6)
'If Not IsNull(rs.Fields(7)) Then txtPAddress = rs.Fields(7)
'If Not IsNull(rs.Fields(8)) Then txtsubsidy = Format(rs.Fields(8), "#0.00")
'If Not IsNull(rs.Fields(9)) Then Txtaccno = rs.Fields(9)
'If Not IsNull(rs.Fields(10)) Then cboBName = rs.Fields(10)
'If Not IsNull(rs.Fields(11)) Then cboBBranch = rs.Fields(11)
'If Not IsNull(rs.Fields(12)) Then a = rs.Fields(12)
'If Not IsNull(rs.Fields(13)) Then cbobranch = rs.Fields(13)
'If Not IsNull(rs.Fields(14)) Then txtPrice = Format(rs.Fields(14), "#0.00")
'If Not IsNull(rs.Fields(15)) Then txtDrAccNo = rs.Fields(15)
'If Not IsNull(rs.Fields(16)) Then txtCrAccNo = rs.Fields(16)
'If Not IsNull(rs.Fields(17)) Then txtcessrate = rs.Fields(17)
'If Not IsNull(rs.Fields(18)) Then txtcessdebit = rs.Fields(18)
'If Not IsNull(rs.Fields(19)) Then txtcesscredit = rs.Fields(19)
'If Not IsNull(rs.Fields(20)) Then b = rs.Fields(20)

End If
End Sub


















