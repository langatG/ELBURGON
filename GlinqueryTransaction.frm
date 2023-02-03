VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "Mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form GlinqueryTransaction 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "General Ledger Inquiry"
   ClientHeight    =   7980
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12495
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7980
   ScaleWidth      =   12495
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdexpef 
      Caption         =   "Expenses Report"
      Height          =   375
      Left            =   9120
      TabIndex        =   46
      Top             =   600
      Width           =   1575
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   2520
      TabIndex        =   19
      Top             =   7560
      Visible         =   0   'False
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Frame Frame1 
      Caption         =   "GL TRANSACTIONS"
      Height          =   4575
      Left            =   2880
      TabIndex        =   14
      Top             =   2160
      Visible         =   0   'False
      Width           =   6735
      Begin VB.Frame fraAlteradv 
         Height          =   3975
         Left            =   2040
         TabIndex        =   20
         Top             =   120
         Visible         =   0   'False
         Width           =   4695
         Begin VB.CommandButton Cmddelete 
            Caption         =   "DELETE"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   1800
            TabIndex        =   28
            Top             =   3120
            Width           =   1215
         End
         Begin VB.TextBox txttransno 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   285
            Left            =   1440
            TabIndex        =   27
            Top             =   2400
            Width           =   1335
         End
         Begin VB.CommandButton Cmdclose 
            Caption         =   "Close"
            DisabledPicture =   "GlinqueryTransaction.frx":0000
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   3360
            Picture         =   "GlinqueryTransaction.frx":0442
            TabIndex        =   26
            Top             =   3120
            Width           =   855
         End
         Begin VB.TextBox txtdracc 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1440
            TabIndex        =   25
            Text            =   "0"
            Top             =   1200
            Width           =   1575
         End
         Begin VB.CommandButton cmdOk 
            Caption         =   "CHANGE!"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   120
            TabIndex        =   24
            Top             =   3120
            Width           =   1215
         End
         Begin VB.TextBox txtcracc 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1440
            TabIndex        =   23
            Top             =   1560
            Width           =   1335
         End
         Begin VB.TextBox txtamt 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1440
            TabIndex        =   22
            Top             =   1920
            Width           =   1335
         End
         Begin VB.TextBox txtdesc 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1440
            TabIndex        =   21
            Top             =   2760
            Width           =   3135
         End
         Begin MSComCtl2.DTPicker dtpDateIssued 
            Height          =   255
            Left            =   1680
            TabIndex        =   29
            Top             =   720
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   450
            _Version        =   393216
            Format          =   154861569
            CurrentDate     =   41437
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Height          =   195
            Left            =   120
            TabIndex        =   45
            Top             =   3720
            Width           =   45
         End
         Begin VB.Label Label6 
            Height          =   255
            Left            =   120
            TabIndex        =   44
            Top             =   4200
            Width           =   735
         End
         Begin VB.Label Label10 
            Caption         =   "Transno"
            Height          =   15
            Left            =   120
            TabIndex        =   43
            Top             =   2640
            Width           =   615
         End
         Begin VB.Label Label11 
            Height          =   375
            Left            =   120
            TabIndex        =   42
            Top             =   3360
            Width           =   1215
         End
         Begin VB.Label Label12 
            Height          =   255
            Left            =   3000
            TabIndex        =   41
            Top             =   1320
            Width           =   615
         End
         Begin VB.Label Label13 
            Height          =   255
            Left            =   1200
            TabIndex        =   40
            Top             =   1920
            Width           =   1695
         End
         Begin VB.Label Label14 
            Height          =   255
            Left            =   1560
            TabIndex        =   39
            Top             =   2520
            Width           =   1695
         End
         Begin VB.Label Label15 
            Caption         =   "Transdate"
            Height          =   375
            Left            =   120
            TabIndex        =   38
            Top             =   720
            Width           =   1335
         End
         Begin VB.Label Label28 
            Caption         =   "DR Acc"
            Height          =   375
            Left            =   120
            TabIndex        =   37
            Top             =   1200
            Width           =   1215
         End
         Begin VB.Label lbldocno 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   1560
            TabIndex        =   36
            Top             =   240
            Width           =   1575
         End
         Begin VB.Label Label27 
            Caption         =   "Docno"
            Height          =   375
            Left            =   120
            TabIndex        =   35
            Top             =   240
            Width           =   735
         End
         Begin VB.Label lblaprinamt 
            Caption         =   "CR Acc"
            Height          =   255
            Left            =   120
            TabIndex        =   34
            Top             =   1560
            Width           =   1095
         End
         Begin VB.Label lblintamt 
            AutoSize        =   -1  'True
            Caption         =   "Amount"
            Height          =   195
            Left            =   120
            TabIndex        =   33
            Top             =   1920
            Width           =   540
         End
         Begin VB.Label lblcurrprinamt 
            Height          =   375
            Left            =   120
            TabIndex        =   32
            Top             =   3000
            Width           =   1215
         End
         Begin VB.Label Label5 
            Caption         =   "Transno"
            Height          =   255
            Left            =   120
            TabIndex        =   31
            Top             =   2400
            Width           =   615
         End
         Begin VB.Label Label16 
            Caption         =   "Descriptions"
            Height          =   255
            Left            =   120
            TabIndex        =   30
            Top             =   2760
            Width           =   1215
         End
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Close"
         Height          =   375
         Left            =   4680
         TabIndex        =   16
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox txtDocNo 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   360
         Width           =   1815
      End
      Begin MSComctlLib.ListView ListView3 
         Height          =   2535
         Left            =   240
         TabIndex        =   17
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
         NumItems        =   6
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
            Text            =   "Transdate"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Transno"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Transdes"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label Label30 
         Caption         =   "DocumentNo"
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.CommandButton frmAccountStatement 
      Caption         =   "View Statement"
      Height          =   375
      Left            =   7200
      TabIndex        =   13
      Top             =   600
      Width           =   1695
   End
   Begin VB.CommandButton cmdFindacc 
      Caption         =   "<>"
      Height          =   315
      Left            =   4260
      TabIndex        =   9
      Top             =   90
      Width           =   465
   End
   Begin MSComCtl2.DTPicker dtpFromdate 
      Height          =   315
      Left            =   2160
      TabIndex        =   6
      Top             =   600
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   556
      _Version        =   393216
      Format          =   155385857
      CurrentDate     =   40179
   End
   Begin VB.TextBox txtAccno 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2160
      TabIndex        =   1
      Top             =   90
      Width           =   1995
   End
   Begin MSComCtl2.DTPicker dtpTodate 
      Height          =   315
      Left            =   5280
      TabIndex        =   8
      Top             =   600
      Width           =   1620
      _ExtentX        =   2858
      _ExtentY        =   556
      _Version        =   393216
      Format          =   155385857
      CurrentDate     =   40364
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   5325
      Left            =   240
      TabIndex        =   12
      Top             =   2040
      Width           =   9420
      _ExtentX        =   16616
      _ExtentY        =   9393
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
      NumItems        =   8
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "TransDate"
         Object.Width           =   2293
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "TransDescription"
         Object.Width           =   4057
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "Debits"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Credits"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Balance"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Document No"
         Object.Width           =   2117
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
   Begin VB.Label Label4 
      Caption         =   "As by the Start of Range"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   240
      TabIndex        =   11
      Top             =   1080
      Width           =   2625
   End
   Begin VB.Label txtBalByRange 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   2880
      TabIndex        =   10
      Top             =   1080
      Width           =   2235
   End
   Begin VB.Label Label8 
      Caption         =   "To Date"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4200
      TabIndex        =   7
      Top             =   600
      Width           =   915
   End
   Begin VB.Label Label7 
      Caption         =   "From Date"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   240
      TabIndex        =   5
      Top             =   660
      Width           =   1155
   End
   Begin VB.Label lblCurrentbalance 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   7200
      TabIndex        =   4
      Top             =   1080
      Width           =   2235
   End
   Begin VB.Label Label2 
      Caption         =   "Book Balance"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   5520
      TabIndex        =   3
      Top             =   1080
      Width           =   1545
   End
   Begin VB.Label lblGlname 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   4845
      TabIndex        =   2
      Top             =   90
      Width           =   4590
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Account No"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   285
      TabIndex        =   0
      Top             =   165
      Width           =   1200
   End
End
Attribute VB_Name = "GlinqueryTransaction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim NormalBal As String
Dim RangeOpeningBal As Double
Private Sub rebuld_Gl(glaccno As String)
'//REFRESH BALANCES FIRST HERE
  Dim OpeningBal  As Double
  Dim CurrentBal2 As Double
Dim rsbal As New ADODB.Recordset, bal As Double
sql = "SELECT     *  FROM         GLSETUP   WHERE  accno='" & glaccno & "'     order by accno "
Set rs = oSaccoMaster.GetRecordset(sql)
If Not rs.EOF Then
'get the clusters and move to the next stuff
'scluster = IIf(IsNull(rs.Fields("tcluster")), 0, rs.Fields("tcluster"))
ACCNO = rs.Fields("accno")
OpeningBal = IIf(IsNull(rs.Fields("OpeningBal")), 0, rs.Fields("OpeningBal"))
NormalBal = rs.Fields("Normalbal")
'sql = "SET DATEFORMAT DMY Select " _
'        & " (select ISNULL(sum(amount),0) " _
'        & " from gltransactions " _
'        & " where Draccno='" & AccNo & "' and DATEADD(dd, DATEDIFF(dd, 0, TRANSDATE), 0)>='" & dtpFromdate & "') and DATEADD(dd, DATEDIFF(dd, 0, TRANSDATE), 0)<='" & Get_Server_Date & "')DR," _
'        & " (select ISNULL(sum(amount),0) " _
'        & " from gltransactions " _
'        & " where Craccno='" & AccNo & "' DATEADD(dd, DATEDIFF(dd, 0, TRANSDATE), 0)>='" & dtpFromdate & "') and DATEADD(dd, DATEDIFF(dd, 0, TRANSDATE), 0)<='" & Get_Server_Date & "')CR"
   sql = "SET DATEFORMAT DMY Select " _
        & " (select ISNULL(sum(amount),0) " _
        & " from gltransactions " _
        & " where Draccno='" & ACCNO & "'   and DATEADD(dd, DATEDIFF(dd, 0, TRANSDATE), 0)>='" & dtpFromdate & "' and DATEADD(dd, DATEDIFF(dd, 0, TRANSDATE), 0)<='" & dtpTodate & "')DR," _
        & " (select ISNULL(sum(amount),0) " _
        & " from gltransactions " _
        & " where Craccno='" & ACCNO & "'  and DATEADD(dd, DATEDIFF(dd, 0, TRANSDATE), 0)>='" & dtpFromdate & "' and DATEADD(dd, DATEDIFF(dd, 0, TRANSDATE), 0)<='" & dtpTodate & "')CR"
Set rst = oSaccoMaster.GetRecordset(sql)
If Not rst.EOF Then
  If NormalBal = "Debit" Then
            CurrentBal2 = OpeningBal + rst("DR") - rst("CR")
        Else
            CurrentBal2 = OpeningBal + rst("CR") - rst("DR")
        End If
    End If
    oSaccoMaster.Execute ("update GLSETUP set CurrentBal='" & CurrentBal2 & "' WHERE  accno='" & txtAccno & "'  ")
End If
End Sub
Private Sub cmdCorrect_Click()
 On Error GoTo Capture
    Set rst = oSaccoMaster.GetRecordset("SET DATEFORMAT DMY select trAnsactionno,MEMBERNO from vwglissues where transdate>='" & dtpFromdate.value & "' and transdate<='" & dtpTodate.value & "' ORDER BY TRANSDATE")
    If rst.EOF Then
        Exit Sub
    End If
    ProgressBar1.max = 100
    ProgressBar1.Visible = True
    While Not rst.EOF
        ProgressBar1.value = (rst.AbsolutePosition / rst.RecordCount) * 100
        oSaccoMaster.ExecuteThis ("UPDATE GLTRANSACTIONS SET SOURCE='" & rst!memberno & "' WHERE TRANSACTIONNO='" & rst!transactionNo & "'")
        If success = False Then
            GoTo Capture
        End If
    rst.MoveNext
    Wend
    MsgBox "Update Done!"
 Exit Sub
Capture:
 MsgBox IIf(ErrorMessage = "", err.description, ErrorMessage)
End Sub

Private Sub cmdFind_Click()
    frmSearchMembers.Show vbModal
    mno = SearchValue
    If mno <> "" Then
        txtMemberNo.Text = SearchValue
    End If
End Sub

Private Sub cmdclose_Click()
Unload Me
End Sub

Private Sub cmddelete_Click()
Dim groupId  As String
 Dim status As String
    If Cmddelete.Caption = "DELETE" Then
'    frmAuthorize.Show vbModal
'    If Authorize = False Then
'            MsgBox "Transaction denied/differed", vbOKOnly + vbExclamation
'            Exit Sub
'         End If
'
'        If MsgBox("You want to alter the Loan Details?", vbQuestion + vbYesNo) = vbYes Then
'            'frmAuthorize.Show vbModal
'             Set Rst = oSaccoMaster.GetRecordset(" SELECT GroupId FROM   UserAccounts WHERE UserLoginIDs='" & User & "'  ")
'              If Not Rst.EOF Then
'                groupId = Rst("GroupId")
'                Else
'              End If
'            If UCase(groupId) <> "ADM" And UCase(groupId) <> "MAN" Then
'            MsgBox "Transaction denied/differed", vbOKOnly + vbExclamation
'            Exit Sub
'            End If
'        Else
'            Exit Sub
'        End If
       
    Else
        If txtamt = "" Then
            MsgBox "Current amount should not be empty", vbInformation
            
            Exit Sub
        End If

      End If
      
'     If Not oSaccoMaster.Execute(" Set dateformat dmy DELETE FROM  REPAY WHERE     Repayid = '" & txttransno & "' AND LoanNo = '" & txtvno & "' and Principal = " & txtAccno & " and Interest = " & txttranstype & "") Then
'
'            GoTo Capture
'        End If
        If Not oSaccoMaster.Execute(" Set dateformat dmy Delete From GLTRANSACTIONS WHERE     (DocumentNo = '" & txtDocNo & "') And (id = " & txttransno & ")") Then
      
'     If Not oSaccoMaster.Execute(" set dateformat dmy Update Reversals set Transno  = '" & txttransno & "', LoanNo = '" & txtvno & "' ,Principal = " & txtAccno & " , Interest = " & txttranstype & ",Datereceived = " & dtpDateIssued & ",Auditid='" & User & "',Machine='" & Mach & "'") Then
'       GoTo Capture
    End If
    
    txtcurramt = ""
        'cmdDelete.Caption = "DELETE!"
        fraAlteradv.Visible = False
    'End If
    Exit Sub
Capture:
    ShowErrorMessage err.description
End Sub

Private Sub cmdexpef_Click()
        reportname = "expensesre.rpt"
        Show_Sales_Crystal_Report STRFORMULA, reportname, CompanyName
End Sub

Private Sub cmdFindacc_Click()
    frmAcctsSearch.Show vbModal, Me
    If Continue Then
        If SearchValue <> "" Then
            txtAccno.Text = SearchValue
            SearchValue = ""
            Continue = False
        End If
    End If
End Sub

Private Sub cmdOK_Click()
Dim groupId  As String
 Dim status As String
    If cmdOk.Caption = "CHANGE!" Then
    'frmAuthorize.Show vbModal
'    If Authorize = False Then
'            MsgBox "Transaction denied/differed", vbOKOnly + vbExclamation
'            Exit Sub
'         End If
'        If MsgBox("You want to alter the Loan Details?", vbQuestion + vbYesNo) = vbYes Then
'
'             Set Rst = oSaccoMaster.GetRecordset(" SELECT GroupId FROM   UserAccounts WHERE UserLoginIDs='" & User & "'  ")
'              If Not Rst.EOF Then
'                groupId = Rst("GroupId")
'                Else
'              End If
'            If UCase(groupId) <> "ADM" And UCase(groupId) <> "MAN" Then
'            MsgBox "Transaction denied/differed", vbOKOnly + vbExclamation
'            Exit Sub
'            End If
'        Else
'            Exit Sub
'        End If
        
    
        cmdOk.Caption = "COMMIT!"
    Else
        If txtamt = "" Then
            MsgBox "Current amount should not be empty", vbInformation
            
            Exit Sub
        End If

      
      
    
'        If Not oSaccoMaster.Execute("set dateformat dmy  Update GLTRANSACTIONS Set Principal = " & txtcurrprinamt & ",Interest = " & txtcurrint & ",IntrCharged = " & txtcurrintcharge & ",Amount = " & txttotamt & " " _
'        & "WHERE     loanno = '" & txtvno & "' AND Repayid = '" & txttransno & "'") Then
'            GoTo Capture
'        End If
        
If Not oSaccoMaster.Execute("set dateformat dmy Update GLTRANSACTIONS" _
& " SET  TransDate = '" & dtpDateIssued & "', Amount = " & txtamt & ", DrAccNo = '" & txtdracc & "', CrAccNo = '" & txtcracc & "', DocumentNo = '" & txtDocNo & "',TransDescript='" & txtdesc & "' WHERE     (DocumentNo = '" & txtDocNo & "')And (id = " & txttransno & ")") Then
            GoTo Capture
        End If
    txtcurramt = ""
        cmdOk.Caption = "CHANGE!"
        fraAlteradv.Visible = False
    End If
    Exit Sub
Capture:
    ShowErrorMessage err.description
End Sub

Private Sub Command1_Click()
    Frame1.Visible = False
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub dtpFromdate_Change()
    txtAccno_Change
End Sub

Private Sub dtpTodate_Change()
    'txtAccNo_KeyPress 13
     txtAccno_Change
End Sub

Private Sub Form_Load()
    dtpFromdate = DateSerial(year(Get_Server_Date), 1, 1)
    dtpTodate = Get_Server_Date
End Sub

Private Sub frmAccountStatement_Click()
        'reportname = "AccountStatement.rpt"
        reportname = "GLedgers.rpt"
        Show_Sales_Crystal_Report STRFORMULA, reportname, CompanyName
' reportname = "GLedgers.rpt"
'        Show_Sales_Crystal_Report STRFORMULA, reportname, CompanyName
End Sub

Private Sub lblCurrentbalance_Change()
    lblCurrentbalance = Format(lblCurrentbalance, Cfmt)

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
        tdate = ListView1.SelectedItem
         amt = li.SubItems(3) Or li.SubItems(2)
        mDocNo = li.SubItems(6)
        txtDocNo.Text = mDocNo
        Load_Ledgers mDocNo, txtAccno.Text, tdate
    End If
End Sub



Private Sub ListView3_DblClick()
'    On Error GoTo Capture
'    Dim rsDr As ADODB.Recordset, rsCr As ADODB.Recordset
'    If ListView3.ListItems.Count = 0 Then
'        Exit Sub
'    End If
'    If MsgBox("You want to switch the accounts?", vbQuestion + vbYesNo) = vbNo Then
'        Exit Sub
'    End If
'    Dim newDrAcc As String
'    Dim NewCrAcc As String
'    DRAcc = ListView3.SelectedItem
'    CRAcc = ListView3.SelectedItem.ListSubItems(1)
'
'    newDrAcc = InputBox("Enter the New DRAccNo ", "New Debit Acc", "")
'    NewCrAcc = InputBox("Enter the New CRAccNo ", "New Crebit Acc", "")
'
'
'
'    If newDrAcc = "" And NewCrAcc = "" Then
'        MsgBox "You chose not to make any change!"
'        Exit Sub
'    ElseIf newDrAcc = "" And NewCrAcc <> "" Then
'        Set rs = oSaccoMaster.GetRecordset("SELECT * FROM GLSETUP WHERE ACCNO='" & NewCrAcc & "'")
'        If rs.EOF Then
'            MsgBox "The new Credit account is not a valid Gl Account", vbCritical
'            Exit Sub
'        End If
'        sql = "Update Gltransactions set CRAccNo='" & NewCrAcc & "' where documentno='" & txtDocNo & "' and DRAccNo='" & DRAcc & "' and CRAccno='" & CRAcc & "'"
'    ElseIf newDrAcc <> "" And NewCrAcc = "" Then
'        Set rst = oSaccoMaster.GetRecordset("SELECT * FROM GLSETUP WHERE ACCNO='" & newDrAcc & "'")
'        If rs.EOF Then
'            MsgBox "The new Debit account is not a valid Gl Account", vbCritical
'            Exit Sub
'        End If
'        sql = "Update Gltransactions set DRAccNo='" & newDrAcc & "' where documentno='" & txtDocNo & "' and DRAccNo='" & DRAcc & "' and CRAccno='" & CRAcc & "'"
'    Else
'        If NewCrAcc <> "" Then
'            Set rsCr = oSaccoMaster.GetRecordset("SELECT * FROM GLSETUP WHERE ACCNO='" & NewCrAcc & "'")
'        End If
'        If newDrAcc <> "" Then
'            Set rsDr = oSaccoMaster.GetRecordset("SELECT * FROM GLSETUP WHERE ACCNO='" & newDrAcc & "'")
'        End If
'        If rsCr.EOF Then
'            MsgBox "The new Credit account is not a valid Gl Account", vbCritical
'            Exit Sub
'        End If
'        If rsDr.EOF Then
'            MsgBox "The new Debit account is not a valid Gl Account", vbCritical
'            Exit Sub
'        End If
'        sql = "Update Gltransactions set DRAccNo='" & newDrAcc & "',CRAccNo='" & NewCrAcc & "' where documentno='" & txtDocNo & "' and DRAccNo='" & DRAcc & "' and CRAccno='" & CRAcc & "'"
'    End If
'    oSaccoMaster.ExecuteThis (sql)
'    If success = False Then
'        MsgBox ErrorMessage
'    Else
'        MsgBox "Operation Successfull", vbInformation
'    End If
'    Exit Sub
'Capture:
'    MsgBox err.description
On Error GoTo Capture
'    Dim rsDr As ADODB.Recordset, rsCr As ADODB.Recordset
'    If ListView3.ListItems.Count = 0 Then
'        Exit Sub
'    End If
'    If MsgBox("You want to switch the accounts?", vbQuestion + vbYesNo) = vbNo Then
'        Exit Sub
'    End If
'    Dim newDrAcc As String
'    Dim NewCrAcc As String
'    dracc = ListView3.SelectedItem
'    CRAcc = ListView3.SelectedItem.ListSubItems(1)
'
'    newDrAcc = InputBox("Enter the New DRAccNo ", "New Debit Acc", "")
'    NewCrAcc = InputBox("Enter the New CRAccNo ", "New Crebit Acc", "")
'
'
'
'    If newDrAcc = "" And NewCrAcc = "" Then
'        MsgBox "You chose not to make any change!"
'        Exit Sub
'    ElseIf newDrAcc = "" And NewCrAcc <> "" Then
'        Set rs = oSaccoMaster.GetRecordset("SELECT * FROM GLSETUP WHERE ACCNO='" & NewCrAcc & "'")
'        If rs.EOF Then
'            MsgBox "The new Credit account is not a valid Gl Account", vbCritical
'            Exit Sub
'        End If
'        sql = "Update Gltransactions set CRAccNo='" & NewCrAcc & "' where documentno='" & txtDocNo & "' and DRAccNo='" & dracc & "' and CRAccno='" & CRAcc & "'"
'    ElseIf newDrAcc <> "" And NewCrAcc = "" Then
'        Set Rst = oSaccoMaster.GetRecordset("SELECT * FROM GLSETUP WHERE ACCNO='" & newDrAcc & "'")
'        If rs.EOF Then
'            MsgBox "The new Debit account is not a valid Gl Account", vbCritical
'            Exit Sub
'        End If
'        sql = "Update Gltransactions set DRAccNo='" & newDrAcc & "' where documentno='" & txtDocNo & "' and DRAccNo='" & dracc & "' and CRAccno='" & CRAcc & "'"
'    Else
'        If NewCrAcc <> "" Then
'            Set rsCr = oSaccoMaster.GetRecordset("SELECT * FROM GLSETUP WHERE ACCNO='" & NewCrAcc & "'")
'        End If
'        If newDrAcc <> "" Then
'            Set rs = oSaccoMaster.GetRecordset("SELECT * FROM GLSETUP WHERE ACCNO='" & newDrAcc & "'")
'        End If
'        If rsCr.EOF Then
'            MsgBox "The new Credit account is not a valid Gl Account", vbCritical
'            Exit Sub
'        End If
'        If rsDr.EOF Then
'            MsgBox "The new Debit account is not a valid Gl Account", vbCritical
'            Exit Sub
'        End If
'        sql = "Update Gltransactions set DRAccNo='" & newDrAcc & "',CRAccNo='" & NewCrAcc & "' where documentno='" & txtDocNo & "' and DRAccNo='" & dracc & "' and CRAccno='" & CRAcc & "'"
'    End If
'    oSaccoMaster.ExecuteThis (sql)
'    If success = False Then
'        MsgBox ErrorMessage
'    Else
'        MsgBox "Operation Successfull", vbInformation
'    End If
'    Exit Sub
If ListView3.ListItems.Count = 0 Then
        Exit Sub
    End If
    With ListView3
'        lblLoanno = .SelectedItem.Text
'        URUMANDI = getRperoid(lblLoanno)
'            If URUMANDI = "Monthly" Then
'             karaba = 1
'            Else
'             karaba = 0
'            End If
'          Calculate_Loan_Repayment .SelectedItem.Text
            
        
       'flow
       'lbldocno = .SelectedItem.Text
       lbldocno = txtDocNo
        txtdracc = .SelectedItem.Text
        dtpDateIssued.value = .SelectedItem.ListSubItems(3)
        txtcracc = .SelectedItem.ListSubItems(1)
        txtamt = .SelectedItem.ListSubItems(2)
        txttransno = .SelectedItem.ListSubItems(4)
        txtdesc = .SelectedItem.ListSubItems(5)
        'txtintcharge = .SelectedItem.ListSubItems(7)
        'txtperiod = Rperiod
        'cboLoanCode.Text = LoanCode
        fraAlteradv.Visible = True
        'Label5.Caption = mMemberno
    End With




Capture:
End Sub

Private Sub txtAccno_Change()
    If Trim(txtAccno) = "" Then
            Exit Sub
        Else
       rebuld_Gl (txtAccno)
    End If
    If Trim(txtAccno) = "" Then
        Exit Sub
    End If
    If dtpFromdate > dtpTodate Then
        dtpFromdate = dtpTodate
    End If
    Get_GL_AccDetails txtAccno
    If GlAccName = "" Then
        ListView1.ListItems.Clear
        lblGlname.Caption = ""
        lblCurrentbalance.Caption = 0
        txtBalByRange.Caption = 0
        Exit Sub
    Else
        lblGlname.Caption = GlAccName
        NormalBal = GlAccNBal
        'dtpFromdate.Value = OpeningBalDate
        lblCurrentbalance.Caption = CurrentBal
        'RangeOpeningBal = getGlBalance(Txtaccno, dtpFromdate, dtpFromdate)
        RangeOpeningBal = getGlBalance1(txtAccno, dtpFromdate, dtpFromdate)
        'txtBalByRange.Caption = getGlBalance(Txtaccno, dtpFromdate, dtpTodate)
        oSaccoMaster.Execute ("Truncate table gledgers")
        LoadTransactions
    End If
End Sub


Private Sub LoadTransactions()
    On Error GoTo SysError
    Dim rsRecon As New Recordset, BankBal As Double, bCredits As Double, bDebits As Double, _
    RsDesc As New Recordset
    'BankBal = RangeOpeningBal
    BankBal = 0
    
'    If lblfullnames.Caption = "ANY" Then
'        Set rsRecon = oSaccoMaster.GetRecordset("SET DATEFORMAT DMY EXEC getGlTransactions '" & txtAccno & "','" & dtpFromdate.value & "','" & Format(dtpTodate.value, "DD/MM/YYYY") & "','ANY'")
'    Else
        Set rsRecon = oSaccoMaster.GetRecordset("SET DATEFORMAT DMY EXEC getGlTransactions '" & txtAccno & "','" & dtpFromdate.value & "','" & Format(dtpTodate.value, "DD/MM/YYYY") & "'")
    'End If
    ListView1.ListItems.Clear
    With rsRecon
        While Not .EOF
            'DoEvents
            Set li = ListView1.ListItems.Add(, , !transdate)
Jump:
            li.SubItems(1) = !tDescription
            li.SubItems(2) = Format(IIf(!transtype = "DR", !amount, 0), Cfmt)
                'bDebits = bDebits + CDbl(li.SubItems(2))
                 bDebits = CDbl(li.SubItems(2))
            li.SubItems(3) = Format(IIf(!transtype = "CR", !amount, 0), Cfmt)
                'bCredits = bCredits + CDbl(li.SubItems(3))
                 bCredits = CDbl(li.SubItems(3))
            If UCase(NormalBal) = UCase("Dr") Then
                BankBal = BankBal + bDebits - bCredits
                'BankBal = bDebits - bCredits
            Else
                BankBal = BankBal + bCredits - bDebits
                'BankBal = bCredits - bDebits
            End If
            li.SubItems(4) = Format(BankBal, Cfmt)
            li.SubItems(5) = IIf(IsNull(!DocumentNo), "", !DocumentNo)
            'If !DocumentNo = "" Then MsgBox "JV-00013"
            li.SubItems(6) = IIf(IsNull(!DocumentNo), "", !DocumentNo)
            oSaccoMaster.Execute ("set dateformat dmy insert into GLedgers(Transdate,source,Debits,Credits,AccBal,Chequeno,Description,glname) values('" & rsRecon!transdate & "','" & rsRecon!Source & "'," & CDbl(li.SubItems(2)) & "," & CDbl(li.SubItems(3)) & "," & CDbl(li.SubItems(4)) & ",'" & rsRecon!DocumentNo & "','" & li.SubItems(1) & "','" & lblGlname & "')")
            oSaccoMaster.Execute ("update GLSETUP set CurrentBal='" & CDbl(li.SubItems(4)) & "' WHERE  accno='" & txtAccno & "'  ")
            
            lblCurrentbalance.Caption = CDbl(li.SubItems(4))
            
            .MoveNext
        Wend
    End With
    lblCurrentbalance.Caption = CDbl(li.SubItems(4))
    txtBalByRange.Caption = CDbl(li.SubItems(4))
    Exit Sub
SysError:
    MsgBox err.description, vbInformation, Me.Caption
End Sub

Private Sub txtBalByRange_Change()
    txtBalByRange = Format(txtBalByRange, Cfmt)
End Sub
Private Sub Load_Ledgers(DocNo As String, ACCNO As String, tdate As Date)
    On Error GoTo SysError
    Dim rsLedger As New Recordset
    Frame1.Visible = True
    ListView3.ListItems.Clear
    Set rsLedger = oSaccoMaster.GetRecordset("Select * From gltransactions where " _
    & " documentno='" & _
    DocNo & "' and (drAccNo='" & ACCNO & "'or crAccNo='" & ACCNO & "') and transdate='" & tdate & "'order by ID")
    With rsLedger
        If .State = adStateOpen Then
            While Not .EOF
                Set li = ListView3.ListItems.Add(, , IIf(IsNull(!DRaccno), "", !DRaccno))
                li.SubItems(1) = IIf(IsNull(!Craccno), "", !Craccno)
                li.SubItems(2) = IIf(IsNull(!amount), 0, !amount)
                li.SubItems(3) = IIf(IsNull(!transdate), "", !transdate)
                li.SubItems(4) = IIf(IsNull(!id), "", !id)
                li.SubItems(5) = IIf(IsNull(!TransDescript), "", !TransDescript)
                .MoveNext
            Wend
        End If
    End With
    Exit Sub
SysError:
    MsgBox err.description, vbInformation, Me.Caption
End Sub

Private Sub txtMemberNo_Change()
    Set rs = oSaccoMaster.GetRecordset("Select * from members where memberno= '" & txtMemberNo.Text & "'")
    With rs

        If .EOF Then
            lblfullnames.Caption = "ANY"
        Else
            lblfullnames = IIf(IsNull(!surname), "", !surname)
        End If
        LoadTransactions
    End With
End Sub
