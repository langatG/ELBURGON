VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmstandingorders 
   BackColor       =   &H00FFFF80&
   Caption         =   "STANDING ORDER SET UP"
   ClientHeight    =   4740
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8955
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4740
   ScaleWidth      =   8955
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFF80&
      Caption         =   "Stop Standing Order"
      Height          =   375
      Left            =   4440
      TabIndex        =   25
      Top             =   3120
      Width           =   1695
   End
   Begin VB.CommandButton cmdpostall 
      BackColor       =   &H00FFFF80&
      Caption         =   "Post All Suppliers"
      Height          =   375
      Left            =   6600
      TabIndex        =   24
      Top             =   3120
      Width           =   1335
   End
   Begin VB.TextBox txtmaximumamount 
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
      Left            =   3600
      TabIndex        =   23
      Text            =   "0"
      Top             =   1560
      Width           =   1575
   End
   Begin VB.PictureBox Picture5 
      Height          =   255
      Left            =   1560
      Picture         =   "frmstandingorders.frx":0000
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   10
      Top             =   720
      Width           =   255
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "New"
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   3120
      Width           =   855
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   375
      Left            =   2640
      TabIndex        =   8
      Top             =   3120
      Width           =   855
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Default         =   -1  'True
      Height          =   375
      Left            =   1800
      TabIndex        =   7
      Top             =   3120
      Width           =   855
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "Edit"
      Height          =   375
      Left            =   960
      TabIndex        =   6
      Top             =   3120
      Width           =   855
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00FFFF80&
      Caption         =   "Close"
      Height          =   375
      Left            =   3480
      TabIndex        =   5
      Top             =   3120
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
      Left            =   1320
      TabIndex        =   1
      Top             =   2160
      Width           =   3135
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
      ItemData        =   "frmstandingorders.frx":02C2
      Left            =   0
      List            =   "frmstandingorders.frx":02DB
      TabIndex        =   0
      Top             =   1560
      Width           =   1815
   End
   Begin MSComCtl2.DTPicker DTPEndDate 
      Height          =   255
      Left            =   5640
      TabIndex        =   11
      Top             =   2280
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   450
      _Version        =   393216
      Format          =   104071169
      CurrentDate     =   40096
   End
   Begin MSComCtl2.DTPicker DTPStartDate 
      Height          =   255
      Left            =   5640
      TabIndex        =   12
      Top             =   1560
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   450
      _Version        =   393216
      Format          =   104071169
      CurrentDate     =   40096
   End
   Begin MSComCtl2.DTPicker DTPDDeduction 
      Height          =   255
      Left            =   5760
      TabIndex        =   13
      Top             =   360
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   450
      _Version        =   393216
      Format          =   104071169
      CurrentDate     =   40096
   End
   Begin VB.Label Label9 
      BackColor       =   &H00FFFF80&
      Caption         =   "Maximum Amount"
      Height          =   255
      Left            =   3600
      TabIndex        =   22
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Label Label7 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFF80&
      Caption         =   "End Date"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   5640
      TabIndex        =   21
      Top             =   2040
      Width           =   675
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFF80&
      Caption         =   "Start Date"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   5640
      TabIndex        =   20
      Top             =   1320
      Width           =   720
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFF80&
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
      BackColor       =   &H00FFFF80&
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
      BackColor       =   &H00FFFF80&
      Caption         =   "Date of Standing Order"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   3720
      TabIndex        =   17
      Top             =   360
      Width           =   1635
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFF80&
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
      BackColor       =   &H00FFFF80&
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
      BackColor       =   &H00FFFF80&
      Caption         =   "Remarks"
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
      Width           =   990
   End
End
Attribute VB_Name = "frmstandingorders"
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
Cmdnew.Enabled = False

DTPDDeduction = Format(Get_Server_Date, "dd/mm/yyyy")
DTPstartdate = Format(Get_Server_Date, "dd/mm/yyyy")
DTPenddate = DateSerial(year(DTPstartdate), month(DTPstartdate) + 1, 1 - 1)

End Sub

Private Sub cmdpostall_Click()


Dim ans As String
Dim NetP As Currency
Dim rshast As New ADODB.Recordset

Dim DESCR As String

DESCR = cboDeductionType.Text


Dim sno As Long
Startdate = DateSerial(year(DTPDDeduction), month(DTPDDeduction), 1)
Enddate = DateSerial(year(DTPDDeduction), month(DTPDDeduction) + 1, 1 - 1)
 Set rshast = oSaccoMaster.GetRecordset("select sno from d_suppliers order by sno")
While Not rshast.EOF
sno = rshast.Fields(0)
sql = ""
sql = "select description from d_supplier_standingorder where sno='" & sno & "' and description ='" & DESCR & "'"
Set rst = oSaccoMaster.GetRecordset(sql)
        If rst.EOF Then
        '//Update deductions
            Set cn = New ADODB.Connection
            sql = "set dateformat dmy INSERT INTO d_supplier_standingorder"
            sql = sql & "           (SNo, Date_Deduc, Description, Amount, MaxAmount, Period, StartDate, EndDate, auditid,  yyear, Remarks)"
            sql = sql & "  VALUES     (" & sno & ",'" & DTPDDeduction & "','" & DESCR & "'," & txtamount & "," & txtmaximumamount & ",'" & Format(DTPDDeduction, "mmm-YYYY") & "','" & DTPstartdate & "','" & DTPenddate & "','" & User & "'," & year(DTPenddate) & ",'" & txtremarks & "')"
            oSaccoMaster.ExecuteThis (sql)
            
            Else
            
        End If
rshast.MoveNext
Wend

MsgBox "Records Successfully Updated", vbInformation

End Sub

Private Sub cmdsave_Click()
On Error GoTo ErrorHandler

'//Validation

Dim ans As String
Dim NetP As Currency


Startdate = DateSerial(year(DTPDDeduction), month(DTPDDeduction), 1)
Enddate = DateSerial(year(DTPDDeduction), month(DTPDDeduction) + 1, 1 - 1)




Dim DESCR As String

DESCR = cboDeductionType.Text



Startdate = DateSerial(year(DTPDDeduction), month(DTPDDeduction), 1)
Enddate = DateSerial(year(DTPDDeduction), month(DTPDDeduction) + 1, 1 - 1)

'check if the supplier is bringing milk

Set rs = oSaccoMaster.GetRecordset("d_sp_SupNet " & txtSNo & ",'" & Startdate & "','" & Enddate & "', 0")

If Not IsNull(rs.Fields(1)) Then
NetP = rs.Fields(1)
Else
NetP = "0.00"
End If

Set rs = oSaccoMaster.GetRecordset("d_sp_SupNet " & txtSNo & ",'" & Startdate & "','" & Enddate & "', 1")
If Not IsNull(rs.Fields(0)) Then
NetP = NetP - rs.Fields(0)
Else
NetP = NetP - 0
End If


If NetP < CCur(txtamount) Then
ans = MsgBox("The supplier number " & txtSNo & " has; " & vbNewLine & "Gross pay of " & Format((NetP + rs.Fields(0)), "#,##0.00") & vbNewLine & " Total Deductions " & Format(rs.Fields(0), "#,##0.00") & vbNewLine & "NetPay " & Format(NetP, "#,##0.00") & "." & vbNewLine & "Continue anyway?", vbYesNo, "LESS NET AMOUNT")

If ans = vbNo Then
Exit Sub
End If
End If

Set rs = oSaccoMaster.GetRecordset("d_sp_IsClosed '" & Enddate & "'")
If Not rs.EOF Then
    MsgBox "The period " & Enddate & " has been closed by " & rs.Fields(0)
    Exit Sub
End If
'//check if it had another deductions of the same nature.
sql = ""
sql = "select description from d_supplier_standingorder where sno='" & txtSNo & "' and description ='" & DESCR & "'"
Set rst = oSaccoMaster.GetRecordset(sql)
If rst.EOF Then
'//Update deductions
    Set cn = New ADODB.Connection
    sql = "set dateformat dmy INSERT INTO d_supplier_standingorder"
    sql = sql & "           (SNo, Date_Deduc, Description, Amount, MaxAmount, Period, StartDate, EndDate, auditid,  yyear, Remarks)"
    sql = sql & "  VALUES     (" & txtSNo & ",'" & DTPDDeduction & "','" & DESCR & "'," & txtamount & "," & txtmaximumamount & ",'" & Format(DTPDDeduction, "mmm-YYYY") & "','" & DTPstartdate & "','" & DTPenddate & "','" & User & "'," & year(DTPenddate) & ",'" & txtremarks & "')"
    oSaccoMaster.ExecuteThis (sql)
    
    'sql = "d_sp_SupplierDeduct " & txtSNo & ",'" & DTPDDeduction & "','" & DESCR & "'," & txtAmount & ",'" & DTPStartDate & "','" & DTPEndDate & "'," & Year(DTPEndDate) & ",'" & User & "','" & txtRemarks & "'"
    'oSaccoMaster.ExecuteThis (sql)
    Else
    MsgBox "The Deduction Code Has Been Defined for this Member", vbInformation, "Standing Order Set Up"
    Exit Sub
End If


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

Private Sub Command1_Click()
Dim DESCR As String

DESCR = cboDeductionType.Text
sql = ""
sql = "select description from d_supplier_standingorder where sno='" & txtSNo & "' and description ='" & DESCR & "'"
Set rst = oSaccoMaster.GetRecordset(sql)
If Not rst.EOF Then
sql = ""
sql = "update  d_supplier_standingorder set active=1 where sno='" & txtSNo & "' and description='" & Trim(DESCR) & "'"
'update  d_supplier_standingorder set active=1 where sno=5 and description='CBO'
oSaccoMaster.ExecuteThis (sql)
End If
MsgBox "Standing Order Successfully Stopped", vbInformation, "Standing Order Set Up."
End Sub

Private Sub DTPDDeduction_Change()
DTPstartdate = DateSerial(year(DTPDDeduction), month(DTPDDeduction), 1)
DTPenddate = DateSerial(year(DTPDDeduction), month(DTPDDeduction) + 1, 1 - 1)

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

Cmdnew.Enabled = True
cmdsave.Enabled = False
Cmdedit.Enabled = False
cmdDelete.Enabled = False

DTPDDeduction = Format(Get_Server_Date, "dd/mm/yyyy")
DTPstartdate = DateSerial(year(DTPDDeduction), month(DTPDDeduction), 1)
'DTPStartDate = Format(Get_Server_Date, "dd/mm/yyyy")
DTPenddate = DateSerial(year(DTPstartdate), month(DTPstartdate) + 1, 1 - 1)

    cboDeductionType.Clear
    Set myclass = New cdbase

    Provider = myclass.OpenCon

    Set cn = CreateObject("adodb.connection")

    cn.Open Provider, "bi"

    Set rs = CreateObject("adodb.recordset")

    rs.Open "SELECT Description FROM d_DCodes order by 1 ", cn

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

