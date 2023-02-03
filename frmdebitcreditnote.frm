VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmdebitcreditnote 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "DEBIT/CREDIT NOTE"
   ClientHeight    =   6060
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9690
   LinkTopic       =   "Form7"
   ScaleHeight     =   6060
   ScaleWidth      =   9690
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Wrong Posted Amounts"
      Height          =   1455
      Left            =   360
      TabIndex        =   17
      Top             =   4320
      Width           =   8055
      Begin VB.CommandButton cmdreverse 
         Caption         =   "Reverse"
         Height          =   255
         Left            =   4560
         TabIndex        =   20
         Top             =   360
         Width           =   1335
      End
      Begin VB.TextBox txtreceiptno 
         Height          =   285
         Left            =   2160
         TabIndex        =   19
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label Label4 
         Caption         =   "Enter Receipt No."
         Height          =   375
         Left            =   360
         TabIndex        =   18
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Save debit note"
      Height          =   255
      Left            =   2520
      TabIndex        =   16
      Top             =   3720
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Caption         =   "Customer Details"
      Height          =   1575
      Left            =   120
      TabIndex        =   10
      Top             =   1320
      Width           =   6375
      Begin VB.ComboBox cboDCode 
         Height          =   315
         Left            =   240
         TabIndex        =   13
         Text            =   "Combo1"
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox txtNames 
         Appearance      =   0  'Flat
         Height          =   405
         Left            =   2040
         TabIndex        =   12
         Top             =   360
         Width           =   4215
      End
      Begin VB.CommandButton cmdsearch 
         Caption         =   "F"
         Height          =   405
         Left            =   1440
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   360
         Width           =   495
      End
      Begin VB.Label ACCNO 
         Caption         =   "ACCNO"
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label lblGl 
         Caption         =   "Gl Control Acc:"
         Height          =   375
         Left            =   1560
         TabIndex        =   14
         Top             =   960
         Width           =   4695
      End
   End
   Begin VB.CommandButton cmdsave 
      Caption         =   "Save credit note"
      Height          =   255
      Left            =   600
      TabIndex        =   9
      Top             =   3720
      Width           =   1335
   End
   Begin VB.ComboBox cboAccno 
      Height          =   315
      Left            =   165
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   780
      Width           =   1200
   End
   Begin VB.TextBox txtAccNames 
      Height          =   315
      Left            =   1740
      TabIndex        =   3
      Top             =   765
      Width           =   3300
   End
   Begin VB.CommandButton cmdAcctsSearch 
      Height          =   300
      Left            =   1365
      Picture         =   "frmdebitcreditnote.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   765
      Width           =   330
   End
   Begin VB.TextBox txtDr 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1620
      TabIndex        =   1
      Text            =   "0"
      Top             =   3240
      Width           =   1275
   End
   Begin VB.TextBox txtJournaNo 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   1260
      TabIndex        =   0
      Top             =   255
      Width           =   1230
   End
   Begin MSComCtl2.DTPicker dtpReceiptDate 
      Height          =   300
      Left            =   7440
      TabIndex        =   5
      Top             =   315
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "  dd-MM-yyyy"
      Format          =   43057153
      CurrentDate     =   40336
   End
   Begin VB.Label Label1 
      Caption         =   "Transaction Date"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5760
      TabIndex        =   8
      Top             =   360
      Width           =   1575
   End
   Begin VB.Label Label3 
      Caption         =   "Amount"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   795
      TabIndex        =   7
      Top             =   3255
      Width           =   765
   End
   Begin VB.Label Label2 
      Caption         =   "Journal No"
      Height          =   270
      Left            =   120
      TabIndex        =   6
      Top             =   270
      Width           =   885
   End
End
Attribute VB_Name = "frmdebitcreditnote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Totalamount As Currency
Dim pushed As Currency
'Dim objLabelEdit As LabelEdit
Dim interestAcc As String, LoanAcc As String
Dim k As Integer
Dim shareBal As Double, balance As Double
Dim sharesCode As String, VoucherNo As String
Dim isMember As Boolean
Dim memberid As String
Dim subType As String
Dim TotalCr As Double, TotalDr As Double
Dim DRAcc As String, CRAcc As String, ContraAcc As String, paymentno As String

Private Sub cboAccno_Change()
    Dim ACCNO As String
    ACCNO = cboAccno.Text
    sql = "select GLACCNAME,TYPE,SUBTYPE from glsetup where accno='" & ACCNO & "'"
    Set rs = oSaccoMaster.GetRecordset(sql)
    If Not rs.EOF Then
        txtAccNames.Text = rs(0)
        subType = UCase(rs.Fields("SubType"))
        If UCase(rs.Fields(1)) = UCase("MEMBER") Then
            isMember = True
         
           
            cboAccno.Locked = False
          
            cboAccno_KeyPress 13
        Else
            isMember = False
           
            'cboShareType.Text = " "
            'cboLoanno.clear
            'cboLoanno.Text = " "
         
          
        End If
    Else
        txtAccNames.Text = ""
    End If
End Sub

Private Sub cboAccno_Click()
    cboAccno_Change
End Sub

Private Sub cboAccno_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        
    End If
End Sub


Private Sub cmdPrint_Click()

End Sub

Private Sub cboDCode_Change()
    Set Rst = oSaccoMaster.GetRecordset("select p.dname,p.accdr,isnull(gl.glaccname,'')GlName " _
    & " from d_debtors p left outer join glSetup gl on p.accdr=gl.accno " _
    & " where p.dcode='" & cboDCode & "'")
    If Not Rst.EOF Then
        txtNames.Text = Rst("dname")
        DRaccno = Rst("accdr")
        ACCNO = DRaccno
        lblGl = Rst("GlName")
    Else
        txtNames.Text = ""
        DRaccno = ""
        lblGl = ""
    End If

End Sub

Private Sub cboDCode_Click()
    cboDCode_Change
End Sub


Private Sub cmdAcctsSearch_Click()
 On Error Resume Next
    frmAcctsSearch.Show vbModal, Me
    If Continue Then
        If SearchValue <> "" Then
            cboAccno.List(0) = SearchValue
            cboAccno.Text = cboAccno.List(0)
            SearchValue = ""
            Continue = False
        End If
    End If
End Sub

Private Sub cmdreverse_Click()
'//check if the receits exists
sql = ""
sql = "select * from CustomerStmt WHERE REFNO='" & txtreceiptno & "'"
Set rs = oSaccoMaster.GetRecordset(sql)
If rs.EOF Then
MsgBox "The receipt doest not exist"
Exit Sub
Else
'reverse it in the gl
'//get the details from the database before reversals

sql = ""
sql = " SELECT * FROM GLTRANSACTIONS WHERE  DOCUMENTNO='" & txtreceiptno & "'"
Set Rst = oSaccoMaster.GetRecordset(sql)
If Not Rst.EOF Then
Dim dr As String, cr As String, tdate As String, amount As Double
dr = Rst.Fields("draccno")
cr = Rst.Fields("craccno")
amount = Rst.Fields("amount")

NewTransaction txtDr, dtpReceiptDate, "Credit Note"
If Not SaveGLTRANSACTION(dtpReceiptDate.value, amount, cr, dr, txtreceiptno, "..", "Wrong Entry", user, transactionNo) Then
GoTo Capture
End If
'//delete from the customer entry
sql = ""
sql = " delete from CustomerStmt WHERE REFNO='" & txtreceiptno & "'"
oSaccoMaster.ExecuteThis (sql)
End If
End If

MsgBox "Records successfully  updated", vbInformation

    Exit Sub
Capture:
    MsgBox err.description
End Sub

Private Sub cmdsave_Click()
NewTransaction txtDr, dtpReceiptDate, "Credit Note"
 If Not SaveGLTRANSACTION(dtpReceiptDate.value, txtDr, cboAccno, ACCNO, txtJournaNo, cboDCode, "Milk Sales -Creditnote", user, transactionNo) Then
                    GoTo Capture
                End If
                
sql = "set dateformat dmy insert into SalesOrder (OrderNo,Dcode,orderDate,OrderAmount,Balance,Auditid,Remarks,Transactionno,warehouse) " _
        & " Values ('" & txtJournaNo & "','" & cboDCode & "','" & dtpReceiptDate.value & "'," & txtDr & "," & txtDr & ",'" & user & "','ok','" & transactionNo & "','CreditNotes')"
        If Not oSaccoMaster.Execute(sql) Then
            GoTo Capture
        End If
        
sql = "insert into CustomerStmt (invId,TransDate,Refno,Amount,TransType,Balance,Auditid,Remarks,Transactionno) " _
        & " Values ('" & txtJournaNo & "','" & dtpReceiptDate.value & "','" & cboDCode & "'," & txtDr & ",'CR',0,'" & user & "','Credit Note','" & transactionNo & "')"
        
        If Not oSaccoMaster.Execute(sql) Then
            GoTo TransError
        End If
        MsgBox "Records update successfully"
        Exit Sub
TransError:
    'DispatchTrans.RollbackTrans
    MsgBox IIf(ErrorMessage = "", err.description, ErrorMessage)
    Exit Sub
Capture:
    MsgBox err.description
End Sub

Private Sub cmdsearch_Click()
Me.MousePointer = vbHourglass
        frmSearchDebtors.Show vbModal
        cboDCode = sel
        cboDCode_Change
        'txtTCode_Validate True
        Me.MousePointer = 0
End Sub

Private Sub Command1_Click()
NewTransaction txtDr, dtpReceiptDate, "Credit Note"
 If Not SaveGLTRANSACTION(dtpReceiptDate.value, txtDr, ACCNO, cboAccno, txtJournaNo, cboDCode, "Milk Sales -Creditnote", user, transactionNo) Then
                    GoTo Capture
                End If
                
sql = "set dateformat dmy insert into SalesOrder (OrderNo,Dcode,orderDate,OrderAmount,Balance,Auditid,Remarks,Transactionno,warehouse) " _
        & " Values ('" & txtJournaNo & "','" & cboDCode & "','" & dtpReceiptDate.value & "'," & txtDr & "," & txtDr & ",'" & user & "','ok','" & transactionNo & "','CreditNotes')"
        If Not oSaccoMaster.Execute(sql) Then
            GoTo Capture
        End If
        
sql = "insert into CustomerStmt (invId,TransDate,Refno,Amount,TransType,Balance,Auditid,Remarks,Transactionno) " _
        & " Values ('" & txtJournaNo & "','" & dtpReceiptDate.value & "','" & cboDCode & "',0,'DR'," & txtDr & ",'" & user & "','Credit Note','" & transactionNo & "')"
        
        If Not oSaccoMaster.Execute(sql) Then
            GoTo TransError
        End If
        MsgBox "Records update successfully"
        Exit Sub
TransError:
    'DispatchTrans.RollbackTrans
    MsgBox IIf(ErrorMessage = "", err.description, ErrorMessage)
    Exit Sub
Capture:
    MsgBox err.description
End Sub

Private Sub Form_Load()
txtJournaNo = JVnumber
dtpReceiptDate = Date
End Sub
Private Function JVnumber()
Dim jvid
    Set rs = oSaccoMaster.GetRecordset("select COUNT (distinct orderno) from SalesOrder ")
    If Not rs.EOF Then
        JVnumber = "DCR-" & Format(CStr(rs(0) + 1), "000")
    End If
'    Select Case jvid
'        Case Is < 10
'            JVnumber = "MCJ-0000" & CStr(jvid)
'        Case Is < 100
'            JVnumber = "MCJ-000" & CStr(jvid)
'        Case Is < 1000
'            JVnumber = "MCJ-00" & CStr(jvid)
'        Case Is < 10000
'            JVnumber = "MCJ-0" & CStr(jvid)
'        Case Else
'            JVnumber = "MCJ-" & CStr(jvid)
'    End Select


End Function
