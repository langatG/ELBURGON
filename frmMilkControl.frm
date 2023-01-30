VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmMilkControl 
   BackColor       =   &H00C0FFFF&
   Caption         =   "FRME"
   ClientHeight    =   7080
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   5295
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   7080
   ScaleWidth      =   5295
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdstatement 
      Caption         =   "Debtors Statement"
      Height          =   375
      Left            =   960
      TabIndex        =   32
      Top             =   6600
      Width           =   2415
   End
   Begin VB.CommandButton cmdreprint 
      Caption         =   "Reprint"
      Height          =   375
      Left            =   3960
      TabIndex        =   31
      Top             =   6120
      Width           =   1335
   End
   Begin VB.CommandButton cmdnewsearch 
      Caption         =   "New "
      Height          =   285
      Left            =   4080
      TabIndex        =   30
      Top             =   120
      Width           =   615
   End
   Begin VB.CheckBox chkapp 
      Caption         =   "Cess Applicable"
      Height          =   285
      Left            =   3000
      TabIndex        =   29
      Top             =   4440
      Width           =   1815
   End
   Begin VB.PictureBox Picture1 
      Height          =   285
      Left            =   3840
      Picture         =   "frmMilkControl.frx":0000
      ScaleHeight     =   225
      ScaleWidth      =   195
      TabIndex        =   19
      Top             =   120
      Width           =   255
   End
   Begin VB.PictureBox Picture3 
      Height          =   285
      Left            =   2760
      Picture         =   "frmMilkControl.frx":02C2
      ScaleHeight     =   225
      ScaleWidth      =   195
      TabIndex        =   18
      Top             =   3000
      Width           =   255
   End
   Begin VB.TextBox txtdcode 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1560
      TabIndex        =   17
      Top             =   3000
      Width           =   1215
   End
   Begin VB.TextBox txtRefNo 
      Appearance      =   0  'Flat
      Height          =   405
      Left            =   1560
      TabIndex        =   13
      Top             =   120
      Width           =   2295
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "Edit"
      Height          =   375
      Left            =   1320
      TabIndex        =   12
      Top             =   6120
      Width           =   1335
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   2640
      TabIndex        =   11
      Top             =   6120
      Width           =   1335
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "New"
      Height          =   375
      Left            =   0
      TabIndex        =   10
      Top             =   6120
      Width           =   1335
   End
   Begin MSComCtl2.DTPicker DTPDispatchDate 
      Height          =   375
      Left            =   1560
      TabIndex        =   9
      Top             =   600
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   661
      _Version        =   393216
      Format          =   102432769
      CurrentDate     =   40105
   End
   Begin VB.TextBox txtVariance 
      Appearance      =   0  'Flat
      Height          =   405
      Left            =   1560
      TabIndex        =   8
      Top             =   2520
      Width           =   2295
   End
   Begin VB.TextBox txtIntake 
      Appearance      =   0  'Flat
      Height          =   405
      Left            =   1560
      TabIndex        =   7
      Top             =   2040
      Width           =   2295
   End
   Begin VB.TextBox txtDipping 
      Appearance      =   0  'Flat
      Height          =   405
      Left            =   1560
      TabIndex        =   6
      Top             =   1560
      Width           =   2295
   End
   Begin VB.TextBox txtDispatch 
      Appearance      =   0  'Flat
      Height          =   405
      Left            =   1560
      TabIndex        =   5
      Top             =   1080
      Width           =   2295
   End
   Begin VB.Label Label16 
      Caption         =   "Cess Acc Dr"
      Height          =   255
      Left            =   360
      TabIndex        =   28
      Top             =   4920
      Width           =   2415
   End
   Begin VB.Label Label15 
      Caption         =   "Debtors Acc Cr"
      Height          =   255
      Left            =   360
      TabIndex        =   27
      Top             =   5400
      Width           =   2415
   End
   Begin VB.Label cessdr 
      Height          =   255
      Left            =   3000
      TabIndex        =   26
      Top             =   4920
      Width           =   1575
   End
   Begin VB.Label cesscr 
      Height          =   255
      Left            =   3000
      TabIndex        =   25
      Top             =   5400
      Width           =   1575
   End
   Begin VB.Label Label12 
      Caption         =   "CESS ACCOUNTS"
      Height          =   255
      Left            =   360
      TabIndex        =   24
      Top             =   4440
      Width           =   2175
   End
   Begin VB.Label Label11 
      Height          =   255
      Left            =   3000
      TabIndex        =   23
      Top             =   3960
      Width           =   1575
   End
   Begin VB.Label Label10 
      Height          =   255
      Left            =   3000
      TabIndex        =   22
      Top             =   3480
      Width           =   1575
   End
   Begin VB.Label Label9 
      Caption         =   "Acc Cr"
      Height          =   255
      Left            =   360
      TabIndex        =   21
      Top             =   3960
      Width           =   2415
   End
   Begin VB.Label Label6 
      Caption         =   "Acc Dr"
      Height          =   255
      Left            =   360
      TabIndex        =   20
      Top             =   3480
      Width           =   2415
   End
   Begin VB.Label lblDebtors 
      AutoSize        =   -1  'True
      Height          =   285
      Left            =   3000
      TabIndex        =   16
      Top             =   3000
      Width           =   60
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      Caption         =   "Debtors Code :"
      Height          =   285
      Left            =   120
      TabIndex        =   15
      Top             =   3000
      Width           =   1410
   End
   Begin VB.Label Label7 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      Caption         =   "Reference No. :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   120
      TabIndex        =   14
      Top             =   120
      Width           =   1395
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      Caption         =   "Variance :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   120
      TabIndex        =   4
      Top             =   2640
      Width           =   900
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      Caption         =   "Intake :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   120
      TabIndex        =   3
      Top             =   2160
      Width           =   630
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      Caption         =   "Dispatch : "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   930
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      Caption         =   "Dispatch Date :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   1365
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      Caption         =   "Dipping :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   120
      TabIndex        =   0
      Top             =   1680
      Width           =   795
   End
   Begin VB.Menu mnuinvoice 
      Caption         =   "Invoice"
   End
End
Attribute VB_Name = "frmMilkControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Price As Currency
Dim capp As Integer
Dim crate As Double
Private Sub cmdEdit_Click()
txtRefNo.Locked = True

    txtDipping.Locked = False
    txtDispatch.Locked = False
    txtIntake.Locked = False
    txtVariance.Locked = False

    Cmdnew.Enabled = False
    cmdsave.Enabled = True
    Cmdedit.Enabled = False
    
End Sub

Private Sub cmdNew_Click()
    'txtDipping.Locked = False
    txtDispatch.Locked = False
    txtIntake.Locked = True
    txtVariance.Locked = False
    txtDipping.Locked = True
    txtDispatch = ""
    txtVariance = ""
    txtdcode = ""
    lblDebtors = ""
    DTPDispatchDate = Get_Server_Date
  
    
    

    Cmdnew.Enabled = False
    cmdsave.Enabled = True
    Cmdedit.Enabled = False
    
End Sub

Private Sub cmdnewsearch_Click()
Dim rsr As New ADODB.Recordset
Dim rsg As New ADODB.Recordset
Dim I As Object
Dim Mylength As Integer
'//if this record is new then look for receipts no

''//clear all textboxes





mysql = ""
mysql = "select GenerateReceiptno from param"

Set rsg = oSaccoMaster.GetRecordset(mysql)
If Not rsg.EOF Then
    ''''check check
    If rsg!GenerateReceiptno = True Then
    
        mysql = ""
        mysql = "select * from Receiptno where receiptno like 'RF-%' order by Receipthnoid desc"
        
        Set rsr = oSaccoMaster.GetRecordset(mysql)
        
        If Not rsr.EOF Then
            Mylength = CInt(Mid(rsr!ReceiptNo, 5, 10))
            Mylength = Mylength + 1
            txtRefNo = Padding(Mylength)
            txtRefNo = "RF-" & txtRefNo
        Else
            Mylength = 1
            txtRefNo = "RF-" & Padding(Mylength)
            
        End If
Else
    ''//receiptno  will be keyed in
End If
End If
End Sub

Private Sub cmdreprint_Click()
STRFORMULA = "{d_MilkControl.RefNo}='" & txtRefNo & "'"
    reportname = "milkinvoice.rpt"
    Show_Sales_Crystal_Report STRFORMULA, reportname, title
End Sub

Private Sub cmdsave_Click()
If txtdcode = "" Then
MsgBox "Debtors code cannot be blank; input an existing one", vbCritical
Exit Sub
End If
If txtDispatch = "" Then
    MsgBox "Please enter the dispatch quantity."
        txtDispatch.SetFocus
    Exit Sub
End If

If txtDipping = "" Then
    MsgBox "Please enter the dipping quantity."
        txtDipping.SetFocus
    Exit Sub
End If

If txtIntake = "" Then
    MsgBox "Please enter the intake quantity."
        txtIntake.SetFocus
    Exit Sub
End If

If txtVariance = "" Then
    MsgBox "Please enter the variance quantity."
        txtVariance.SetFocus
    Exit Sub
End If



If txtRefNo = "" Then
    MsgBox "Please enter the reference number."
        txtRefNo.SetFocus
    Exit Sub
End If
'//check if the dispatch is greater than the dipping
If CDbl(txtDipping) < CDbl(txtDispatch) Then 'raiise an alarm
MsgBox "You cannot take more you have in the tank", vbCritical
Exit Sub
End If
Dim Debit As String
Dim Credit As String
'Dim Price As Currency

'Set rs = oSaccoMaster.GetRecordset("d_sp_getAccName '" & lblDebtors & "'")
'If IsNull(rs.Fields(0)) Then
'    MsgBox "The debtors account not set. " & vbNewLine & "Please contact the accountant to set GL for " & lblDebtors
'        Exit Sub
'End If
'
Debit = Label10
'
'Set rs = oSaccoMaster.GetRecordset("d_sp_getAccName 'Milk sale'")
'If IsNull(rs.Fields(0)) Then
'    MsgBox "The Creditors account not set. " & vbNewLine & "Please contact the accountant to set GL for milk sales"
'        Exit Sub
'End If
'
Credit = Label11

    

    If Not Save_GLTRANSACTION(Format(DTPDispatchDate, "dd/mm/yyyy"), (CCur(Price) * CCur(txtDispatch)), Debit, Credit, "Milk Sales ", txtRefNo, User, ErrorMessage, "Milk Sales", 1, 1, txtRefNo, transactionNo, "", "", 0) Then
            If ErrorMessage <> "" Then
                MsgBox ErrorMessage, vbInformation, Me.Caption
                ErrorMessage = ""
            End If
    End If
    
    If capp = 1 Then
    
    If Not Save_GLTRANSACTION(Format(DTPDispatchDate, "dd/mm/yyyy"), (CCur(crate) * CCur(txtDispatch)), cessdr, cesscr, "Cess Deductions ", txtRefNo, User, ErrorMessage, "Cess Deductions", 1, 1, txtRefNo, transactionNo, "", "", 0) Then
            If ErrorMessage <> "" Then
                MsgBox ErrorMessage, vbInformation, Me.Caption
                ErrorMessage = ""
            End If
    End If
    
    End If
        
'd_sp_MilkControl @DispDate char(10), @DipsQnty float,@DipQnty float,@InQnty float,@VarQnty float,@Price char(10),@RefNo varchar(35),@CreditAcc varchar(35),@DebitAcc varchar(35),@AuditID varchar (50)
Set rs = New ADODB.Recordset
sql = "d_sp_MilkControl  '" & DTPDispatchDate & "'," & txtDispatch & "," & txtDipping & "," & txtIntake & "," & txtVariance & "," & Price & ",'" & txtRefNo & "','" & Credit & "','" & Debit & "','" & User & "','" & txtdcode & "'"
oSaccoMaster.ExecuteThis (sql)

'//subtract from the dispatch table

    sql = ""
    sql = "SET      dateformat dmy     SELECT     ID, Intake,transdate     FROM         d_dispatch    WHERE     transdate = '" & DTPDispatchDate & "'"
    Set rs = oSaccoMaster.GetRecordset(sql)
    If rs.EOF Then
sql = ""
sql = "set dateformat dmy INSERT INTO d_dispatch (Transdate, descrip, Intake, dipping, dispatch, auditid, auditdate)values ('" & DTPDispatchDate.value & "','Dispatch',0," & CDbl(txtDipping) - CDbl(txtDispatch) & "," & CDbl(txtDispatch) & ",'" & User & "','" & Get_Server_Date & "')"
oSaccoMaster.ExecuteThis (sql)
Else
sql = ""
sql = "set dateformat dmy UPDATE    d_dispatch  SET   dipping =" & CDbl(txtDipping) - CDbl(txtDispatch) & ",dispatch=" & txtDispatch & "  WHERE     (Transdate = '" & DTPDispatchDate & "')"
oSaccoMaster.ExecuteThis (sql)
End If
mysql = "set dateformat dmy Insert into Receiptno(Receiptno,Auditdate,auditid)values('" & txtRefNo & "','" & Format(Get_Server_Date, "dd/MM/yyyy") & "','" & User & "')"
oSaccoMaster.ExecuteThis (mysql)

MsgBox "Records saved successifully."
'//PRINT THE REPORT HERE
'milkinvoice

'd_MilkControl."RefNo"
    STRFORMULA = "{d_MilkControl.RefNo}='" & txtRefNo & "'"
    reportname = "milkinvoice.rpt"
    Show_Sales_Crystal_Report STRFORMULA, reportname, title
    Form_Load
End Sub

Private Sub cmdstatement_Click()
'milkstatement


    reportname = "milkstatement.rpt"
    Show_Sales_Crystal_Report STRFORMULA, reportname, title
End Sub

Private Sub DTPDispatchDate_Change()
    Set rs = New ADODB.Recordset
    sql = ""
    'sql = "set dateformat dmy sp_dispatch '" & DTPDispatchDate & "'"
    sql = "set dateformat dmy select intake,dipping from  d_dispatch where transdate='" & DTPDispatchDate & "'"
    Set rs = oSaccoMaster.GetRecordset(sql)
    If Not rs.EOF Then
    txtIntake = (rs.Fields(0))
    txtDipping = rs.Fields(1)
    Else
    txtIntake = "0.00"
    txtDipping = 0#
    End If
    
    
    
End Sub

Private Sub DTPDispatchDate_Click()
DTPDispatchDate_Change
End Sub

Private Sub DTPDispatchDate_Validate(Cancel As Boolean)
DTPDispatchDate_Change
End Sub

Private Sub Form_Load()
    DTPDispatchDate = Format(Get_Server_Date, "dd/mm/yyyy")
    DTPDispatchDate.MaxDate = Format(Get_Server_Date, "dd/mm/yyyy")
    cmdnewsearch_Click
'    txtCreditAcc.Locked = True
'    txtCreditAccName.Locked = True
'    txtDebitAcc.Locked = True
'    txtDebitAccName.Locked = True
    txtDipping.Locked = True
    txtDispatch.Locked = True
    txtIntake.Locked = True
    txtVariance.Locked = True
    
'    txtCreditAcc = ""
'    txtCreditAccName = ""
'    txtDebitAcc = ""
'    txtDebitAccName = ""
    txtDipping = ""
    txtDispatch = ""
    txtIntake = ""
    txtVariance = ""
    

    Cmdnew.Enabled = True
    cmdsave.Enabled = True
    Cmdedit.Enabled = False
    
    
    
    Set rs = New ADODB.Recordset
    sql = "d_sp_DailyTotal '" & DTPDispatchDate & "'"
    Set rs = oSaccoMaster.GetRecordset(sql)
    If Not IsNull(rs.Fields(0)) Then
    txtIntake = Format(rs.Fields(0), "#0.00")
    'End If
    Else
    txtIntake = "0.00"
    End If
    
            
End Sub

Private Sub lvwCreditAcc_DblClick()
'Dim rsAccount As New ADODB.Recordset
'txtCreditAcc = lvwCreditAcc.SelectedItem
'Set rsAccount = oSaccoMaster.GetRecordset("Select * From GLSETUP where " _
'            & "accno= '" & txtCreditAcc & "'")
'If Not rsAccount.EOF Then
'   txtCreditAccName = IIf(IsNull(rsAccount!GlAccName), "", rsAccount!GlAccName)
'
 
'End If


'lvwCreditAcc.Visible = False

End Sub


Private Sub lvwDebitAcc_DblClick()
Dim rsAccount As New ADODB.Recordset
'txtDebitAcc = lvwDebitAcc.SelectedItem
'Set rsAccount = oSaccoMaster.GetRecordset("Select * From GLSETUP where " _
'            & "accno= '" & txtDebitAcc & "'")
If Not rsAccount.EOF Then
'   txtDebitAccName = IIf(IsNull(rsAccount!GlAccName), "", rsAccount!GlAccName)
  
 
End If


'lvwDebitAcc.Visible = False

End Sub



Private Sub mnuinvoice_Click()
frmmilkinvoice.Show vbModal
End Sub

Private Sub Picture1_Click()
Me.MousePointer = vbHourglass
         frmSearchMilkControl.Show vbModal
        txtRefNo = sel
        txtRefNo_Validate True
        Me.MousePointer = 0
End Sub

Private Sub Picture3_Click()
Me.MousePointer = vbHourglass
         frmSearchDebtors.Show vbModal
        txtdcode = sel
        txtdcode_Validate True
        Me.MousePointer = 0
End Sub

Private Sub txtCreditAccName_Change()
'On Error GoTo SysError
    Dim rsAccount As New Recordset
'    lvwCreditAcc.ListItems.Clear
    
'    If Trim$(txtCreditAccName) <> "" Then
'        'If Editing = True Then
'            Set rsAccount = oSaccoMaster.GetRecordset("Select * From GLSETUP where " _
'            & "GLAccName Like '%" & txtCreditAccName & "%'")
'            With rsAccount
'                If .State = adStateOpen Then
'                    If Not .EOF Then
'                        'lvwContraAcc.Visible = True
'                        If .RecordCount = 1 Then
'                            txtCreditAcc = IIf(IsNull(!accno), "", !accno)
'                            Editing = True
'                            txtCreditAccName = IIf(IsNull(!GlAccName), "", !GlAccName)
'                            lvwCreditAcc.Visible = False
'                            Else
'                            lvwCreditAcc.Visible = False
'
'                        End If
'                    Else
'                        lvwCreditAcc.Visible = False
'                    End If
'                    'lvwDeductionAcc.Visible = True
'                    While Not .EOF
'                        lvwCreditAcc.Visible = True
'                        Set li = lvwCreditAcc.ListItems.Add(, , IIf(IsNull(!accno), "", !accno))
'                        li.SubItems(1) = IIf(IsNull(!GlAccName), "", !GlAccName)
'                        .MoveNext
'                    Wend
'                    'lvwDeductionAcc.Visible = False
'                End If
'            End With
'        'End If
'    End If
'    Exit Sub
'SysError:
'    MsgBox err.description, vbInformation, Me.Caption
'
'End Sub
'
'
'
'Private Sub txtdcode_Validate(Cancel As Boolean)
'sql = "select dname,Price from d_debtors where dcode='" & txtdcode & "'"
'Set rs = oSaccoMaster.GetRecordset(sql)
'If Not rs.EOF Then
'If Not IsNull(rs.Fields(0)) Then lblDebtors = rs.Fields(0)
'If Not IsNull(rs.Fields(1)) Then Price = rs.Fields(1)
'End If
'End Sub
'
'Private Sub txtDebitAccName_Change()
'On Error GoTo SysError
'    Dim rsAccount As New Recordset
'    lvwDebitAcc.ListItems.Clear
'
'    If Trim$(txtDebitAccName) <> "" Then
'        'If Editing = True Then
'            Set rsAccount = oSaccoMaster.GetRecordset("Select * From GLSETUP where " _
'            & "GLAccName Like '%" & txtDebitAccName & "%'")
'            With rsAccount
'                If .State = adStateOpen Then
'                    If Not .EOF Then
'                        'lvwContraAcc.Visible = True
'                        If .RecordCount = 1 Then
'                            txtDebitAcc = IIf(IsNull(!accno), "", !accno)
'                            Editing = True
'                            txtDebitAccName = IIf(IsNull(!GlAccName), "", !GlAccName)
'                            lvwDebitAcc.Visible = False
'                            Else
'                            lvwDebitAcc.Visible = False
'
'                        End If
'                    Else
'                        lvwDebitAcc.Visible = False
'                    End If
'                    'lvwDeductionAcc.Visible = True
'                    While Not .EOF
'                        lvwDebitAcc.Visible = True
'                        Set li = lvwDebitAcc.ListItems.Add(, , IIf(IsNull(!accno), "", !accno))
'                        li.SubItems(1) = IIf(IsNull(!GlAccName), "", !GlAccName)
'                        .MoveNext
'                    Wend
'                    'lvwDeductionAcc.Visible = False
'                End If
'            End With
'        'End If
'    End If
'    Exit Sub
'SysError:
'    MsgBox err.description, vbInformation, Me.Caption

End Sub

Private Sub txtdcode_Validate(Cancel As Boolean)
Set rs = oSaccoMaster.GetRecordset("SELECT dname,Price,accdr,acccr,drcess,crcess,capp,crate FROM d_Debtors WHERE DCode = '" & txtdcode & "'")
If rs.RecordCount > 0 Then
If Not IsNull(rs.Fields(1)) Then Price = rs.Fields(1)
If Not IsNull(rs.Fields(0)) Then lblDebtors = rs.Fields(0)
If Not IsNull(rs.Fields(2)) Then Label10 = rs.Fields(2)
If Not IsNull(rs.Fields(3)) Then Label11 = rs.Fields(3)
If Not IsNull(rs.Fields(4)) Then cessdr = rs.Fields(4)
If Not IsNull(rs.Fields(5)) Then cesscr = rs.Fields(5)
If Not IsNull(rs.Fields(6)) Then capp = Abs(rs.Fields(6))
If Not IsNull(rs.Fields(7)) Then crate = rs.Fields(7)
If capp = 1 Then
chkapp = vbChecked
Else
chkapp = vbUnchecked
End If
Else
lblDebtors = ""
End If
End Sub

Private Sub txtDipping_Change()
If txtIntake = "" Then
txtIntake = "0"
End If
If txtDipping = "" Then
txtDipping = "0"
End If
txtVariance = Format(txtDipping - txtDispatch, "#0.00")
End Sub

Private Sub txtDipping_Validate(Cancel As Boolean)
txtDispatch_Change
End Sub

Private Sub txtDispatch_Change()
'txtDipping = txtDispatch
If txtDispatch = "" Then
txtDispatch = "0"
End If
If txtDipping = "" Then
txtDipping = "0"
End If
txtVariance = Format(txtDipping - txtDispatch, "#0.00")
End Sub



Private Sub txtDispatch_Validate(Cancel As Boolean)
txtDipping_Change
End Sub

Private Sub txtIntake_Change()
txtDispatch_Change
End Sub

Private Sub txtIntake_Validate(Cancel As Boolean)
txtDispatch_Change
End Sub

Private Sub txtRefNo_Validate(Cancel As Boolean)
On Error GoTo ErrorHandler
'SELECT TOP 300 DispDate,RefNo,dcode,DispQnty,Price,InQnty,Variance FROM dbo.d_MilkControl"
If Trim(txtRefNo) = "" Then
Exit Sub
End If
 Set rs = oSaccoMaster.GetRecordset("SELECT DispDate,dcode,DispQnty,Price,InQnty,Variance FROM d_MilkControl WHERE RefNo = '" & txtRefNo & "'")
 
 If rs.RecordCount > 0 Then
    DTPDispatchDate = rs.Fields(0)
    txtDispatch = rs.Fields(2)
    txtDipping = txtDispatch
    txtIntake = rs.Fields(4)
    txtVariance = rs.Fields(5)
    txtdcode = rs.Fields(1)
    
    Cmdedit.Enabled = True
Else
    Cmdedit.Enabled = False
    
End If
txtdcode_Validate True
Exit Sub
ErrorHandler:
MsgBox err.description
End Sub
