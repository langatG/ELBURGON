VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmbonusprocess 
   Caption         =   "Bonus Processing"
   ClientHeight    =   2910
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6645
   LinkTopic       =   "Form1"
   ScaleHeight     =   2910
   ScaleWidth      =   6645
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdindre 
      Caption         =   "Print Individual Report"
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
      Left            =   4080
      TabIndex        =   9
      Top             =   120
      Width           =   2415
   End
   Begin VB.CommandButton Cmdshares 
      Caption         =   "Process Shares"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2520
      TabIndex        =   7
      Top             =   2040
      Width           =   1695
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4680
      TabIndex        =   6
      Top             =   2040
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "General Bonus report"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4080
      TabIndex        =   5
      Top             =   840
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Process Bonuses"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   1
      Top             =   2040
      Width           =   1935
   End
   Begin MSComCtl2.DTPicker DTPstdate 
      Height          =   375
      Left            =   1680
      TabIndex        =   0
      Top             =   480
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
      _Version        =   393216
      Format          =   120389633
      CurrentDate     =   42680
   End
   Begin MSComCtl2.DTPicker DTPedate 
      Height          =   375
      Left            =   1680
      TabIndex        =   2
      Top             =   1200
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
      _Version        =   393216
      Format          =   120389633
      CurrentDate     =   42680
   End
   Begin MSComctlLib.ProgressBar prgStatus 
      Height          =   255
      Left            =   0
      TabIndex        =   8
      Top             =   1680
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Label lbledate 
      Caption         =   "End date"
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
      Left            =   480
      TabIndex        =   4
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label lblstartdate 
      Caption         =   "Start date"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   480
      Width           =   1095
   End
End
Attribute VB_Name = "frmbonusprocess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdindre_Click()
frmSupplierStmtBonus.Show vbModal
End Sub

Private Sub Cmdshares_Click()
Dim lastdate, mon As Date
Dim lastdateofsale As Date
Dim pcode As String
Dim NetPay As Double
Dim dy, a As Integer
Dim grade As String
Dim bank As String
Dim bcode As String
Dim BBranch As String
Dim rsd, rskk, rsk, rsg As New ADODB.Recordset
sql = ""
sql = "DELETE FROM d_Bonus"
Set rs = oSaccoMaster.GetRecordset(sql)
sql = ""
sql = "set dateformat dmy SELECT s.SNo,s.Names,s.AccNo,s.Bcode,s.BBranch,d.Remarks, sum(Amount) AS Netpay From d_supplier_deduc d inner join d_Suppliers s on d.sno=s.sno WHERE d.Date_Deduc >= '" & DTPstdate & "' and d.Date_Deduc <= '" & DTPedate & "' and d.Remarks LIKE '%bonus%' GROUP BY s.sno, s.names,s.AccNo,s.Bcode,s.BBranch,d.Remarks ORDER BY s.sno asc"
Set rs = oSaccoMaster.GetRecordset(sql)
 While Not rs.EOF
    pcode = rs!sno
    NetPay = rs!NetPay
    pname = rs!NAMES
    bank = rs!ACCNO
    bcode = rs!bcode
    BBranch = rs!BBranch
    'select pcode,ldate,dy,auditdate,audit,grade from ag_paging
    sql = ""
    sql = "set dateformat dmy insert into  d_Bonus (Sno, Name,bank,bcode,branch, Startdate, Enddate, Amount,Pby)"
    sql = sql & "values('" & pcode & "','" & pname & "','" & bank & "','" & bcode & "','" & BBranch & "','" & DTPstdate & "','" & DTPedate & "','" & NetPay & "','" & User & "') "
    oSaccoMaster.ExecuteThis (sql)

rs.MoveNext
Wend
'sharesload
MsgBox "Records successfully done", vbInformation
End Sub
Private Sub sharesload()
Dim lastdate As Date
Dim lastdateofsale As Date
Dim pcode As String
Dim NetPay As Double
Dim dy, a As Integer
Dim grade As String
Dim bank As String
Dim bcode As String
Dim BBranch As String
Dim mon As Integer
Dim rsd, rskk, rsk, rsg As New ADODB.Recordset
sql = ""
sql = "set dateformat dmy DELETE FROM d_Bonus2 where  Date >= '" & DTPstdate & "' and Date <= '" & DTPedate & "'"
Set rs = oSaccoMaster.GetRecordset(sql)
sql = ""
sql = "set dateformat dmy SELECT count(distinct(SNo)) From d_supplier_deduc WHERE   Date_Deduc >= '" & DTPstdate & "' and Date_Deduc <= '" & DTPedate & "' and Remarks LIKE '%bonus%'"
Set rsk = oSaccoMaster.GetRecordset(sql)
sql = ""
sql = "set dateformat dmy SELECT distinct(SNo) From d_supplier_deduc WHERE   Date_Deduc >= '" & DTPstdate & "' and Date_Deduc <= '" & DTPedate & "' and Remarks LIKE '%bonus%'"
Set rskk = oSaccoMaster.GetRecordset(sql)

'sql = ""
'sql = "set dateformat dmy SELECT    count( SNo) From d_supplier_deduc  WHERE Date_Deduc >= '" & DTPstdate & "' and Date_Deduc <= '" & DTPedate & "' and Remarks LIKE '%bonus%' "
'Set rsj = cn.Execute(sql)
Dim b As Double
b = rsk.Fields(0)

prgStatus.max = 100
prgStatus.Min = 0
I = 0
While Not rskk.EOF
a = rskk.Fields(0)
 sql = ""
 sql = "set dateformat dmy insert into  d_Bonus2 (Sno,Date)"
 sql = sql & "values('" & a & "','" & DTPedate & "') "
 oSaccoMaster.ExecuteThis (sql)
     I = I + 1
prgStatus = Round((I / b) * 100, 0)
    
sql = ""
sql = "set dateformat dmy SELECT s.SNo,s.Names,s.AccNo,s.Bcode,s.Location,d.Remarks, d.Amount AS Netpay,d.Date_Deduc From d_supplier_deduc d inner join d_Suppliers s on d.sno=s.sno WHERE  d.sno = '" & a & "' and d.Date_Deduc >= '" & DTPstdate & "' and d.Date_Deduc <= '" & DTPedate & "' and d.Remarks LIKE '%bonus%' GROUP BY s.sno, s.names,s.AccNo,s.Bcode,s.Location,d.Remarks,d.Amount,d.Date_Deduc ORDER BY d.Date_Deduc asc"
Set rs = oSaccoMaster.GetRecordset(sql)
 Do While Not rs.EOF
    mon = month(rs.Fields(7))
            Select Case mon
             Case "1"
              sql = ""
              sql = "set dateformat dmy Update d_Bonus2 SET Mon1 ='" & rs.Fields(6) & "' WHERE Sno='" & rs.Fields(0) & "' and Date >= '" & DTPstdate & "'And Date<='" & DTPedate & "'"
              cn.Execute sql
             Case "2"
              sql = ""
              sql = "set dateformat dmy Update d_Bonus2 SET Mon2 ='" & rs.Fields(6) & "' WHERE Sno='" & rs.Fields(0) & "' and Date >= '" & DTPstdate & "'And Date<='" & DTPedate & "'"
              cn.Execute sql
             Case "3"
              sql = ""
              sql = "set dateformat dmy Update d_Bonus2 SET Mon3 ='" & rs.Fields(6) & "' WHERE Sno='" & rs.Fields(0) & "' and Date >= '" & DTPstdate & "'And Date<='" & DTPedate & "'"
              cn.Execute sql
             Case "4"
              sql = ""
              sql = "set dateformat dmy Update d_Bonus2 SET Mon4 ='" & rs.Fields(6) & "' WHERE Sno='" & rs.Fields(0) & "' and Date >= '" & DTPstdate & "'And Date<='" & DTPedate & "'"
              cn.Execute sql
             Case "5"
              sql = ""
              sql = "set dateformat dmy Update d_Bonus2 SET Mon5 ='" & rs.Fields(6) & "' WHERE Sno='" & rs.Fields(0) & "' and Date >= '" & DTPstdate & "'And Date<='" & DTPedate & "'"
              cn.Execute sql
             Case "6"
              sql = ""
              sql = "set dateformat dmy Update d_Bonus2 SET Mon6 ='" & rs.Fields(6) & "' WHERE Sno='" & rs.Fields(0) & "' and Date >= '" & DTPstdate & "'And Date<='" & DTPedate & "'"
              cn.Execute sql
             Case "7"
              sql = ""
              sql = "set dateformat dmy Update d_Bonus2 SET Mon7 ='" & rs.Fields(6) & "' WHERE Sno='" & rs.Fields(0) & "' and Date >= '" & DTPstdate & "'And Date<='" & DTPedate & "'"
              cn.Execute sql
             Case "8"
              sql = ""
              sql = "set dateformat dmy Update d_Bonus2 SET Mon8 ='" & rs.Fields(6) & "' WHERE Sno='" & rs.Fields(0) & "' and Date >= '" & DTPstdate & "'And Date<='" & DTPedate & "'"
              cn.Execute sql
             Case "9"
              sql = ""
              sql = "set dateformat dmy Update d_Bonus2 SET Mon9 ='" & rs.Fields(6) & "' WHERE Sno='" & rs.Fields(0) & "' and Date >= '" & DTPstdate & "'And Date<='" & DTPedate & "'"
              cn.Execute sql
             Case "10"
              sql = ""
              sql = "set dateformat dmy Update d_Bonus2 SET Mon10 ='" & rs.Fields(6) & "' WHERE Sno='" & rs.Fields(0) & "' and Date >= '" & DTPstdate & "'And Date<='" & DTPedate & "'"
              cn.Execute sql
             Case "11"
              sql = ""
              sql = "set dateformat dmy Update d_Bonus2 SET Mon11 ='" & rs.Fields(6) & "' WHERE Sno='" & rs.Fields(0) & "' and Date >= '" & DTPstdate & "'And Date<='" & DTPedate & "'"
              cn.Execute sql
             Case "12"
              sql = ""
              sql = "set dateformat dmy Update d_Bonus2 SET Mon12 ='" & rs.Fields(6) & "' WHERE Sno='" & rs.Fields(0) & "' and Date >= '" & DTPstdate & "'And Date<='" & DTPedate & "'"
              cn.Execute sql

             Case Else
            End Select
  rs.MoveNext
 Loop
rskk.MoveNext
Wend
End Sub
Private Sub Command1_Click()
Dim lastdate As Date
Dim lastdateofsale As Date
Dim pcode As String
Dim NetPay As Double
Dim dy As Integer
Dim grade As String
Dim bank As String
Dim bcode As String
Dim BBranch As String
Dim rsd, rsk As New ADODB.Recordset
sql = ""
sql = "set dateformat dmy DELETE FROM d_Bonus where Startdate >= '" & DTPstdate & "' and Enddate <= '" & DTPedate & "'"
Set rs = oSaccoMaster.GetRecordset(sql)
'sql = ""
'sql = "set dateformat dmy SELECT    count( SNo) From d_supplier_deduc  WHERE Date_Deduc >= '" & DTPstdate & "' and Date_Deduc <= '" & DTPedate & "' and Remarks LIKE '%bonus%' "
'Set rsk = cn.Execute(sql)
'Dim a As Double
'a = rsk.Fields(0)
'
'prgStatus.max = 100
'prgStatus.Min = 0
'I = 0
sql = ""
sql = "set dateformat dmy SELECT s.SNo,s.Names,s.AccNo,s.Bcode,s.Location,d.Remarks, SUM(Amount) AS Netpay From d_supplier_deduc d inner join d_Suppliers s on d.sno=s.sno WHERE   d.Date_Deduc >= '" & DTPstdate & "' and d.Date_Deduc <= '" & DTPedate & "' and d.Remarks LIKE '%bonus%' GROUP BY s.sno, s.names,s.AccNo,s.Bcode,s.Location,d.Remarks ORDER BY s.sno asc"

Set rs = oSaccoMaster.GetRecordset(sql)
While Not rs.EOF
 ' I = I + 1
'prgStatus = Round((I / a) * 100, 0)
pcode = rs!sno
NetPay = rs!NetPay
pname = rs!NAMES
bank = rs!ACCNO
bcode = rs!bcode
BBranch = rs!Location

'select pcode,ldate,dy,auditdate,audit,grade from ag_paging
sql = ""
sql = "set dateformat dmy insert into  d_Bonus (Sno, Name,bank,bcode,branch, Startdate, Enddate, Amount,Pby)"
sql = sql & "values('" & pcode & "','" & pname & "','" & bank & "','" & bcode & "','" & BBranch & "','" & DTPstdate & "','" & DTPedate & "','" & NetPay & "','" & User & "') "
oSaccoMaster.ExecuteThis (sql)

rs.MoveNext
Wend
sharesload
MsgBox "Records successfully done", vbInformation

'//give him the report here
'agrovetagingreport
'reportname = "Bonus Report.rpt"

 
 'Show_Sales_Crystal_Report STRFORMULA, reportname, ""
'//we look for receipts tables
'//get the number of days
'/// insert into the number of days
'//give us a report

End Sub


Private Sub Command2_Click()
'reportname = "Bonus Report.rpt"
reportname = "bonusyear.rpt"
    Show_Sales_Crystal_Report STRFORMULA, reportname, ""
End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Form_Load()
DTPstdate = Format(Get_Server_Date, "dd/mm/yyyy")
DTPstdate = DateSerial(Year(DTPstdate) - 1, month(-2), 1)
DTPedate = DateSerial(Year(DTPstdate) + 1, month(-1), 1 - 1)

End Sub

