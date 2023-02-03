VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmMembership 
   Caption         =   "Membership"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   3945
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   3945
   StartUpPosition =   3  'Windows Default
   Begin MSComCtl2.DTPicker DTPenddate 
      Height          =   375
      Left            =   1920
      TabIndex        =   3
      Top             =   600
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393216
      Format          =   51773441
      CurrentDate     =   43784
   End
   Begin MSComCtl2.DTPicker DTPstartdate 
      Height          =   375
      Left            =   1920
      TabIndex        =   2
      Top             =   120
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393216
      Format          =   51773441
      CurrentDate     =   43784
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Detailedreport"
      Height          =   495
      Left            =   960
      TabIndex        =   1
      Top             =   1320
      Width           =   1815
   End
   Begin VB.CommandButton cmdMemberregister 
      Caption         =   "Memberregister"
      Height          =   495
      Left            =   960
      TabIndex        =   0
      Top             =   2280
      Width           =   2055
   End
   Begin VB.Label Label2 
      Caption         =   "End date"
      Height          =   375
      Left            =   480
      TabIndex        =   5
      Top             =   600
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Start date"
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "frmMembership"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdMemberregister_Click()
reportname = "suppliersregister.rpt"

 Show_Sales_Crystal_Report "", reportname, ""

End Sub

Private Sub Command1_Click()
Dim lastdate As Date
Dim lastdateofsale As Date
Dim pcode As String
Dim pname As String
Dim dy As Integer
Dim grade As String
Dim curamt As Double
Dim quantity As Double
Dim id As String

Dim rsd As New ADODB.Recordset
sql = ""
'sql = "DELETE FROM ag_sales"


sql = "delete From Detailed"
Set rs = oSaccoMaster.GetRecordset(sql)
sql = ""
sql = "set dateformat dmy SELECT     SUM(m.QSupplied) AS Total, m.SNo, s.Names, s.IdNo FROM         dbo.d_Milkintake AS m INNER JOIN dbo.d_Suppliers AS s ON s.SNo = m.SNo WHERE     (m.TransDate >= '" & DTPstartdate & "') AND (m.TransDate <= '" & DTPenddate & "') GROUP BY m.SNo, s.Names, s.IdNo"




'sql = "set dateformat dmy SELECT     SUM(changeinstock) AS Quantity, p_code,ProductName From ag_stockbalance WHERE   transdate <= '" & DTPedate & "' GROUP BY p_code,ProductName"
'sql = "set dateformat dmy SELECT     r.P_code,p.p_name, SUM(Qua) AS Quantity From ag_Receipts r inner join ag_products p on r.p_code=p.p_code WHERE   r.T_Date >= '" & DTPstdate & "' and r.T_Date <= '" & DTPedate & "'  GROUP BY r.P_code, p.P_name ORDER BY r.P_code asc"
'sql = "set dateformat dmy SELECT     r.P_code,p.productname, SUM(changeinstock) AS Quantity From ag_stockbalance p inner join ag_Receipts r on r.p_code=p.p_code WHERE   r.T_Date >= '" & DTPstdate & "' and r.T_Date <= '" & DTPedate & "'  GROUP BY r.P_code, p.productname ORDER BY r.P_code asc"

Set rs = oSaccoMaster.GetRecordset(sql)
While Not rs.EOF
pcode = rs!sno
quantity = rs!total
pname = Trim(rs!NAMES)
id = rs!idno
'lastdate = rs!transdate
'
'sql = "set dateformat dmy SELECT     SUM(Qua) AS Qty, p_code From ag_Receipts WHERE   T_Date <= '" & DTPedate & "' and p_code=" & rs!p_code & "  GROUP BY p_code"
'Set rsd = oSaccoMaster.GetRecordset(sql)
'If Not rsd.EOF Then
'curamt = IIf(IsNull(rs!Quantity), 0, rs!Quantity) - IIf(IsNull(rsd!qty), 0, rsd!qty)
'Else
'curamt = IIf(IsNull(rs!Quantity), 0, rs!Quantity)
'End If
'curamt = IIf(IsNull(rs!Quantity), 0, rs!Quantity) - IIf(IsNull(rsd!qty), 0, rsd!qty)
'select pcode,ldate,dy,auditdate,audit,grade from ag_paging
sql = ""
sql = "set dateformat dmy insert into Detailed (Sno,   Idno, Total)"
sql = sql & "values('" & pcode & "','" & id & "','" & quantity & "') "
oSaccoMaster.ExecuteThis (sql)


'sql = "set dateformat dmy insert into  ag_sales (pcode,pname,Quantity)"
'sql = sql & "values('" & pcode & "','" & pname & "','" & curamt & "') "
'oSaccoMaster.ExecuteThis (sql)


rs.MoveNext
Wend
MsgBox "Records successfully done", vbInformation

'//give him the report here
'agrovetagingreport
'reportname = "cummulative sales.rpt"
'reportname = "evans.rpt"

 
 Show_Sales_Crystal_Report STRFORMULA, reportname, ""
'//we look for receipts tables
'//get the number of days
'/// insert into the number of days
'//give us a report

End Sub
