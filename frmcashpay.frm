VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmcashpay 
   BackColor       =   &H000080FF&
   Caption         =   "Agrovet cash payments updates"
   ClientHeight    =   3960
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8505
   LinkTopic       =   "Form1"
   ScaleHeight     =   3960
   ScaleWidth      =   8505
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture5 
      Height          =   255
      Left            =   1560
      Picture         =   "frmcashpay.frx":0000
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   10
      Top             =   720
      Width           =   255
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "New"
      Height          =   375
      Left            =   1320
      TabIndex        =   9
      Top             =   3000
      Width           =   855
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   375
      Left            =   3840
      TabIndex        =   8
      Top             =   3000
      Width           =   855
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Default         =   -1  'True
      Height          =   375
      Left            =   3000
      TabIndex        =   7
      Top             =   3000
      Width           =   855
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "Edit"
      Height          =   375
      Left            =   2160
      TabIndex        =   6
      Top             =   3000
      Width           =   855
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00FFFF80&
      Caption         =   "Close"
      Height          =   375
      Left            =   4680
      TabIndex        =   5
      Top             =   3000
      Width           =   855
   End
   Begin VB.TextBox txtSNames 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1800
      TabIndex        =   4
      Top             =   720
      Width           =   5775
   End
   Begin VB.TextBox txtAmount 
      Appearance      =   0  'Flat
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """Kshs ""#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2057
         SubFormatType   =   0
      EndProperty
      Height          =   285
      Left            =   1920
      TabIndex        =   3
      Text            =   "0"
      Top             =   1560
      Width           =   1575
   End
   Begin VB.TextBox txtSNo 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   0
      TabIndex        =   2
      Top             =   720
      Width           =   1575
   End
   Begin VB.TextBox txtRemarks 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   1800
      TabIndex        =   1
      Top             =   2160
      Width           =   1335
   End
   Begin VB.ComboBox cboDeductionType 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      ItemData        =   "frmcashpay.frx":02C2
      Left            =   0
      List            =   "frmcashpay.frx":02C4
      TabIndex        =   0
      Top             =   1560
      Width           =   1815
   End
   Begin MSComCtl2.DTPicker DTPEndDate 
      Height          =   255
      Left            =   5640
      TabIndex        =   11
      Top             =   1560
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   450
      _Version        =   393216
      Format          =   122355713
      CurrentDate     =   40096
   End
   Begin MSComCtl2.DTPicker DTPStartDate 
      Height          =   255
      Left            =   3600
      TabIndex        =   12
      Top             =   1560
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   450
      _Version        =   393216
      Format          =   122355713
      CurrentDate     =   40096
   End
   Begin MSComCtl2.DTPicker DTPDDeduction 
      Height          =   255
      Left            =   5760
      TabIndex        =   13
      Top             =   240
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   450
      _Version        =   393216
      Format          =   122355713
      CurrentDate     =   40096
   End
   Begin VB.Label Label7 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H000080FF&
      Caption         =   "End Date"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   5640
      TabIndex        =   21
      Top             =   1320
      Width           =   675
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H000080FF&
      Caption         =   "Start Date"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   3600
      TabIndex        =   20
      Top             =   1320
      Width           =   720
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H000080FF&
      Caption         =   "Type of Deduction"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   0
      TabIndex        =   19
      Top             =   1200
      Width           =   1320
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H000080FF&
      Caption         =   "Supplier Name"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   1800
      TabIndex        =   18
      Top             =   360
      Width           =   1035
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H000080FF&
      Caption         =   "Date of Payment"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   5760
      TabIndex        =   17
      Top             =   0
      Width           =   1185
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H000080FF&
      Caption         =   "Supplier Numer"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   0
      TabIndex        =   16
      Top             =   360
      Width           =   1080
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H000080FF&
      Caption         =   "Amount"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   1920
      TabIndex        =   15
      Top             =   1200
      Width           =   540
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackColor       =   &H000080FF&
      Caption         =   "Receipt No"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   240
      TabIndex        =   14
      Top             =   2160
      Width           =   1275
   End
End
Attribute VB_Name = "frmcashpay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Rs1 As New ADODB.Recordset, rs2 As New ADODB.Recordset, rs3 As New ADODB.Recordset
Dim rs4 As New ADODB.Recordset
Dim myclass As cdbase
Dim Transport As Currency, agrovet As Currency, AI As Currency, TMShares As Currency, FSA As Currency, HShares As Currency, Advance As Currency, Others As Currency

Private Sub cmdclose_Click()
Unload Me
End Sub

Private Sub cmdNew_Click()
txtamount = ""
txtSNames = ""
txtSNo = ""
cboDeductionType = ""

txtamount.Locked = False
txtSNo.Locked = False
cboDeductionType.Locked = False

cmdsave.Enabled = True
cmdnew.Enabled = False

DTPDDeduction = Format(Get_Server_Date, "dd/mm/yyyy")
dtpStartDate = Format(Get_Server_Date, "dd/mm/yyyy")
DTPEndDate = DateSerial(year(dtpStartDate), month(dtpStartDate) + 1, 1 - 1)

End Sub

Private Sub cmdsave_Click()
On Error GoTo ErrorHandler

'//Validation
'Set cn = New ADODB.Connection
'sql = "SET dateformat dmy SELECT NPay,GPay,TDeductions from d_Payroll WHERE SNo=" & txtSNo & " AND EndofPeriod ='" & DateSerial(Year(DTPDDeduction), month(DTPDDeduction) + 1, 1 - 1) & "'"
'Set rs = oSaccoMaster.GetRecordset(sql)
Dim ans As String
Dim NetP As Currency
'If rs.RecordCount > 0 Then
'If Not IsNull(rs.Fields("NPay")) Then
'NetP = rs.Fields("NPay")
'End If
'If Not IsNull(rs.Fields("GPay")) Then
'NetP = rs.Fields("GPay")
'End If

Startdate = DateSerial(year(DTPDDeduction), month(DTPDDeduction), 1)
Enddate = DateSerial(year(DTPDDeduction), month(DTPDDeduction) + 1, 1 - 1)
If txtSNo = "" Then Exit Sub
Set rs = oSaccoMaster.GetRecordset("d_sp_SupNet " & txtSNo & ",'" & Startdate & "','" & Enddate & "', 0")
If Not rs.EOF Then
If Not IsNull(rs.Fields(1)) Then
If Not IsNull(rs.Fields(1)) Then
NetP = IIf(IsNull(rs.Fields(1)), 0, rs.Fields(1))
Else
NetP = "0.00"
End If

Set rs = oSaccoMaster.GetRecordset("d_sp_SupNet " & txtSNo & ",'" & Startdate & "','" & Enddate & "', 1")
If Not IsNull(rs.Fields(0)) Then
NetP = rs.Fields(0) - NetP
Else
NetP = NetP - 0
End If
Set cn = New ADODB.Connection
sql = "d_sp_SupplierDeduct " & txtSNo & ",'" & DTPDDeduction & "','Agrovet'," & -1 * txtamount & ",'" & Startdate & "','" & Enddate & "'," & year(DTPDDeduction) & ",'" & User & "','Cash',''"
oSaccoMaster.ExecuteThis (sql)
If NetP < CCur(txtamount) Then
'ans = MsgBox("The supplier number " & txtSNo & " has; " & vbNewLine & "Gross pay of " & Format((NetP + rs.Fields(0)), "#,##0.00") & vbNewLine & " Total Deductios " & Format(rs.Fields(0), "#,##0.00") & vbNewLine & "NetPay " & Format(NetP, "#,##0.00") & "." & vbNewLine & "You Can Continue anyway?", vbInformation, "LESS NET AMOUNT")

'If ans = vbNo Then
'Exit Sub
'End If

'If ans = vbYes Then
'MsgBox "Please let the supplier apply an amount less or equal to " & Format(rs.Fields("NPay"), "#,##0.00") & ""
'txtamount.SetFocus
'Exit Sub
'End If
End If
End If
Else
MsgBox "There is no record for supplier number " & txtSNo & " for period ending " & DateSerial(year(DTPDDeduction), month(DTPDDeduction) + 1, 1 - 1) & ""
txtSNo.SetFocus
Exit Sub
End If
'End If


If cboDeductionType.Text = "Others" And txtremarks = "" Then
MsgBox "Please enter the remarks."
txtremarks.SetFocus
Exit Sub
End If
If txtremarks = "" Then
txtremarks = " "
End If

Dim DESCR As String

DESCR = cboDeductionType.Text & " cash payment "

'If Trim(cboDeductionType.Text) = "Shares" Then
'DESCR = "HShares"
'End If
'If Trim(cboDeductionType.Text) = "Registration" Then
'DESCR = "TMShares"
'End If

Startdate = DateSerial(year(DTPDDeduction), month(DTPDDeduction), 1)
Enddate = DateSerial(year(DTPDDeduction), month(DTPDDeduction) + 1, 1 - 1)

Set rs = oSaccoMaster.GetRecordset("d_sp_IsClosed '" & Enddate & "'")
If Not rs.EOF Then
    MsgBox "The period " & Enddate & " has been closed by " & rs.Fields(0)
    Exit Sub
End If
'//Update deductions
'Set cn = New ADODB.Connection
'sql = "d_sp_SupplierDeduct " & txtSNo & ",'" & DTPDDeduction & "','" & DESCR & "'," & txtAmount * -1 & ",'" & DTPStartDate & "','" & DTPEndDate & "'," & year(DTPEndDate) & ",'" & User & "','" & txtRemarks & "'"
'oSaccoMaster.ExecuteThis (sql)

'UPDATE Shares Chekoff
If UCase$(DESCR) = "SHARES" Then
Set rs = New ADODB.Recordset
sql = "d_sp_SelectSuppliers '" & txtSNo & "'"
Set rs = oSaccoMaster.GetRecordset(sql)
With rs
 If Not rs.EOF Then
 Dim idno As String, sex As String, Location As String
 idno = IIf(IsNull(!idno), "", !idno)
 sex = IIf(IsNull(!Type), "", !Type)
 Location = IIf(IsNull(!Location), "", !Location)
 End If
End With
sex = Left(sex, 1)
strSQL = "set dateformat dmy insert into [d_Shares]([IdNo],[SNO],[Code],[Name],[Sex],[Loc],[Type],[TransDate],[pmode],[Period],[Amnt],[amount],[AuditId], [AuditDateTime])"
strSQL = strSQL & " values( '" & Trim$(idno) & "','" & txtSNo & "','" & txtSNo & "','" & txtSNames & "','" & sex & "','" & Location & "','" & DESCR & "','"
strSQL = strSQL & Enddate & "',' 0','" & Enddate & "'," & txtamount & "," & txtamount & ",'" & User & "','" & Get_Server_Date & "')"
oSaccoMaster.ExecuteThis (strSQL)

sql = ""
sql = "set dateformat dmy insert into d_sconribution(sno, transdate, amount, bal, transdescription, auditid)"
sql = sql & " values('" & txtSNo & "','" & Enddate & "'," & txtamount & "," & txtamount & ",'" & sex & "','" & User & "') "
oSaccoMaster.ExecuteThis (sql)
End If

'//Update payroll
'Dim Startdate As String, Enddate As String
Set rs2 = New ADODB.Recordset
Dim qnty As Currency, GPay As Currency
'Startdate = DateSerial(DTPMilkDate, cboMonth, 1)

sql = "d_sp_UpdateGPAYQnty '" & Startdate & "','" & Enddate & "'," & txtSNo & ""
Set rs2 = oSaccoMaster.GetRecordset(sql)
If Not rs2.EOF Then
If Not IsNull(rs2.Fields(0)) Then qnty = rs2.Fields(0)
If Not IsNull(rs2.Fields(1)) Then GPay = rs2.Fields(1)
End If


Set Rs1 = New ADODB.Recordset
sql = "d_sp_TotalDeduct " & txtSNo & "," & month(DTPDDeduction) & "," & year(DTPDDeduction) & ""
Set Rs1 = oSaccoMaster.GetRecordset(sql)
If Not Rs1.EOF Then
Dim TotalDed As Currency
If Not IsNull(Rs1.Fields(0)) Then TotalDed = Rs1.Fields(0)
End If
'//Update payroll -- @SNo bigint,@EndPeriod varchar(15),@Kgs float,@GPay money,@NPay money,@TDeductions money,@auditid  varchar(35)
Set cn = New ADODB.Connection
sql = "d_sp_UpdatePayroll  " & txtSNo & ",'" & Enddate & "'," & qnty & "," & GPay & "," & GPay - TotalDed & "," & TotalDed & ",'" & User & "'"
oSaccoMaster.ExecuteThis (sql)



Set rs3 = New ADODB.Recordset
'Dim Startdate As String, Enddate As String
Dim desc As String
Dim Amnt As Currency
'Startdate = DateSerial(Year(DTPDDeduction), month(DTPDDeduction), 1)
'Enddate = DateSerial(Year(DTPMilkDate), month(DTPMilkDate) + 1, 1 - 1)
sql = "d_sp_SupDed " & txtSNo & ",'" & Startdate & "','" & Enddate & "'"
Set rs3 = oSaccoMaster.GetRecordset(sql)
If Not rs3.EOF Then
While Not rs3.EOF
If Not IsNull(rs3.Fields(0)) Then desc = Trim(rs3.Fields(0))
Amnt = 0
If Not IsNull(rs3.Fields(1)) Then Amnt = rs3.Fields(1)
sql = "SET dateformat DMY SELECT     Transport, Agrovet, AI, TMShares, FSA, HShares, Advance, Others FROM d_Payroll WHERE SNo=" & txtSNo & " AND EndofPeriod ='" & Enddate & "'"
Set rs4 = oSaccoMaster.GetRecordset(sql)
If UCase(rs4.Fields(0).name) = UCase(desc) Then
Transport = Amnt
End If
If UCase(rs4.Fields(1).name) = UCase(desc) Then
agrovet = Amnt
End If
If UCase(rs4.Fields(2).name) = UCase(desc) Then
AI = Amnt
End If
If UCase(rs4.Fields(3).name) = UCase(desc) Then
TMShares = Amnt
End If
If UCase(rs4.Fields(4).name) = UCase(desc) Then
FSA = Amnt
End If
If UCase(rs4.Fields(5).name) = UCase(desc) Then
HShares = Amnt
End If
If UCase(rs4.Fields(6).name) = UCase(desc) Then
Advance = Amnt
End If
If UCase(rs4.Fields(7).name) = UCase(desc) Then
Others = Amnt
End If

'//d_sp_UpdatePayrollDed  @SNo bigint,@EndPeriod varchar(15),@Transport money,@Agrovet money,@AI money,@TMShares money,@FSA money,@HShares money,@Advance money,@Others
rs3.MoveNext
Wend
'//Update Deductions -- d_sp_UpdatePayrollDed  @SNo bigint,@EndPeriod varchar(15),@Transport money,@Agrovet money,@AI money,@TMShares money,@FSA money,@HShares money,@Advance money,@Others money
Set cn = New ADODB.Connection
sql = "d_sp_UpdatePayrollDed  " & txtSNo & ",'" & Enddate & "'," & Transport & "," & agrovet & "," & AI & "," & TMShares & "," & FSA & "," & HShares & "," & Advance & "," & Others & ""
oSaccoMaster.ExecuteThis (sql)
End If

Dim txtTCHPBalances As Double
If UCase(Trim(cboDeductionType)) = UCase("Shares") Then

Set rst = New ADODB.Recordset
sql = "select bal from d_shares where sno= '" & txtSNo & "'"
Set rst = oSaccoMaster.GetRecordset(sql)
If Not rst.EOF Then
'txtTCHPBalances = Rst.Fields(0)

 '//get the balance

sql = "SELECT     bal   FROM         d_sconribution  WHERE     sno ='" & txtSNo & "'  ORDER BY transdate DESC, id DESC "
Dim rr As New ADODB.Recordset
Set rr = oSaccoMaster.GetRecordset(sql)
If Not rr.EOF Then
txtTCHPBalances = txtTCHPBalances + CCur(txtamount)
',[sno],[transdate],[amount],[bal],[transdescription],[auditid],[auditdate],[mno]
  'From [EASYTEA].[dbo].[d_sconribution]
  sql = ""
  sql = "set dateformat dmy insert into d_sconribution([sno],[transdate],[amount],[bal],[transdescription],[auditid])"
  sql = sql & " values ('" & txtSNo & "','" & DTPDDeduction & "'," & txtamount & "," & txtTCHPBalances & ",'Shares','" & User & "') "
  oSaccoMaster.ExecuteThis (sql)
  
  sql = ""
  sql = "update d_shares set bal=" & txtTCHPBalances & " where sno='" & txtSNo & "' "
  oSaccoMaster.ExecuteThis (sql)
'txtTCHPBALANCE = rr.Fields(0)
End If
Else
'//add new one
txtTCHPBalances = 0
sql = "insert into d_Shares(sno, Cash,bal,auditid)"
sql = sql & " values('" & txtSNo & "',1," & txtamount & ",'" & User & "')"
oSaccoMaster.ExecuteThis (sql)
sql = ""
sql = "set dateformat dmy insert into d_sconribution([sno],[transdate],[amount],[bal],[transdescription],[auditid])"
sql = sql & " values ('" & txtSNo & "','" & DTPDDeduction & "'," & txtamount & "," & txtamount & ",'Shares','" & User & "') "
oSaccoMaster.ExecuteThis (sql)

End If
End If


Transport = 0
agrovet = 0
AI = 0
TMShares = 0
FSA = 0
HShares = 0
Advance = 0
Others = 0

'Dim Yr As Integer

'Yr = Year(DTPDDeduction)
'vbHourglass
'Fixed deductions update
'oSaccoMaster.ExecuteThis ("d_sp_PresetDeductAssign_99 '" & DTPStartDate & "','" & DTPEndDate & "'," & Yr & ",'" & User & "', " & txtSNo)

'Payroll update
'd_sp_GDedNet @StartDate varchar(10) , @endPeriod varchar(10)
'oSaccoMaster.ExecuteThis ("d_sp_GDedNet_99 '" & DTPStartDate & "','" & DTPEndDate & "'," & txtSNo)

'Update transporters
'd_sp_TransUpdate @StartDate varchar(10),@EndPeriod varchar(10),@User varchar(35) AS
'oSaccoMaster.ExecuteThis ("d_sp_TransUpdate_99 '" & DTPStartDate & "','" & DTPEndDate & "','" & User & "'," & txtSNo)


'oSaccoMaster.ExecuteThis ("d_sp_TransPRoll '" & DTPStartDate & "','" & DTPEndDate & "','" & User & "'")
'Lock period

txtamount = ""
txtSNo = ""
txtSNo_Validate True

txtSNo.SetFocus
'Form_Load
MsgBox "Records successively updated."
Exit Sub
ErrorHandler:
MsgBox err.description

End Sub

Private Sub DTPDDeduction_Change()
dtpStartDate = DateSerial(year(DTPDDeduction), month(DTPDDeduction), 1)
DTPEndDate = DateSerial(year(DTPDDeduction), month(DTPDDeduction) + 1, 1 - 1)

End Sub



Private Sub Form_Load()
txtamount = ""
txtSNames = ""
txtSNo = ""
txtremarks = ""

cboDeductionType = ""

txtamount.Locked = True
txtSNames.Locked = True
txtSNo.Locked = True
cboDeductionType.Locked = True

cmdnew.Enabled = True
cmdsave.Enabled = False
cmdEdit.Enabled = False
cmdDelete.Enabled = False

DTPDDeduction = Format(Get_Server_Date, "dd/mm/yyyy")
dtpStartDate = DateSerial(year(DTPDDeduction), month(DTPDDeduction), 1)
'DTPStartDate = Format(Get_Server_Date, "dd/mm/yyyy")
DTPEndDate = DateSerial(year(dtpStartDate), month(dtpStartDate) + 1, 1 - 1)

    cboDeductionType.Clear
    Set myclass = New cdbase

    Provider = myclass.OpenCon

    Set cn = CreateObject("adodb.connection")

   cn.Open Provider, "atm", "atm"

    Set rs = CreateObject("adodb.recordset")

    rs.Open "SELECT Description FROM d_DCodes", cn

    If rs.EOF Then Exit Sub

    With rs

        While Not .EOF

         cboDeductionType.AddItem rs.Fields("Description")

         .MoveNext

        Wend

    End With
    


End Sub

Private Sub Form_LostFocus()
'txtAmount.DataFormat = FormatCurrency("'Kshs '#,##0.00", Val(txtAmount))
End Sub

Private Sub Form_Unload(Cancel As Integer)
'oSaccoMaster.ExecuteThis ("d_sp_GDedNet '" & DTPStartDate & "', '" & DTPEndDate & "'")
End Sub

Private Sub Picture5_Click()
        Me.MousePointer = vbHourglass
        frmSearchSupplier.Show vbModal
        txtSNo = sel
        txtSNo_Validate True
        Me.MousePointer = 0
End Sub

Private Sub txtSNo_Validate(Cancel As Boolean)
Dim a, t As Boolean
Set rs = New ADODB.Recordset
sql = "d_sp_SelectSuppliers '" & txtSNo & "'"
Set rs = oSaccoMaster.GetRecordset(sql)
If Not rs.EOF Then
If Not IsNull(rs.Fields(0)) Then txtSNames = rs.Fields(2)
Else
txtSNames = ""
End If
End Sub

