VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmbonusprocess 
   Caption         =   "Bonus Processing"
   ClientHeight    =   3075
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6615
   LinkTopic       =   "Form1"
   ScaleHeight     =   3075
   ScaleWidth      =   6615
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "General Shares report"
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
      TabIndex        =   10
      Top             =   1320
      Width           =   2415
   End
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
      Height          =   495
      Left            =   2520
      TabIndex        =   7
      Top             =   2400
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
      Height          =   495
      Left            =   5640
      TabIndex        =   6
      Top             =   2400
      Visible         =   0   'False
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
      Height          =   495
      Left            =   4080
      TabIndex        =   5
      Top             =   720
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
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   2400
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
      Format          =   120455169
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
      Format          =   120455169
      CurrentDate     =   42680
   End
   Begin MSComctlLib.ProgressBar prgStatus 
      Height          =   255
      Left            =   0
      TabIndex        =   8
      Top             =   1920
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
 On Error GoTo SysError

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
    sql = "DELETE FROM d_SharesReport"
    Set rs = oSaccoMaster.GetRecordset(sql)
    prgStatus.value = 0
    sql = ""
    'sql = "set dateformat dmy SELECT s.SNo,s.Names,s.IdNo,s.PhoneNo,s.AccNo,s.Bcode,s.Location From d_supplier_deduc d inner join d_Suppliers s on d.sno=s.sno inner join d_sconribution k on k.sno=Convert(varchar,s.SNo) WHERE (d.Description LIKE '%SHARES%' or k.transdescription LIKE '%SHARES%' ) GROUP BY s.SNo,s.Names,s.IdNo,s.PhoneNo,s.AccNo,s.Bcode,s.Location ORDER BY s.sno asc"
    sql = "set dateformat dmy SELECT SNo,Names,IdNo,PhoneNo,AccNo,Bcode,Location From d_Suppliers ORDER BY sno asc"
    Set rs = oSaccoMaster.GetRecordset(sql)
     While Not rs.EOF
        prgStatus.Max = rs.RecordCount
        prgStatus.value = rs.AbsolutePosition
        
        PhoneNo = rs!PhoneNo
        idno = rs!idno
        pcode = rs!sno
        
        'NetPay = rs!NetPay
        pname = Replace(rs!NAMES, "'", "")
        bank = rs!ACCNO
        bcode = rs!bcode
        BBranch = rs!Location
        

        
        Dim rss As New Recordset, amt As Double, rsts As New Recordset, shareamt As Double, TXTshares As Double
        Set rsts = oSaccoMaster.GetRecordset("SELECT isnull(SUM(Amount),0) AS amtt From d_sconribution WHERE (transdescription LIKE '%shares%') AND (SNo = '" & pcode & "')")
        If Not rsts.EOF Then
            shareamt = IIf(IsNull(rsts!amtt), 0, rsts!amtt)
        End If
        
        Set rss = oSaccoMaster.GetRecordset("SELECT    isnull(SUM(Amount),0) AS amt From d_supplier_deduc WHERE     (Description LIKE '%shares%') AND (SNo = '" & pcode & "')")
        If Not rss.EOF Then
        TXTshares = IIf(IsNull(rss!amt), 0, rss!amt) + shareamt
        End If
        
        If TXTshares > 0 Then
            Set Rst = oSaccoMaster.GetRecordset("SELECT SNo FROM d_SharesReport WHERE Sno = '" & pcode & "'")
            If Rst.EOF Then
               Dim namecheck As String
               namecheck = Replace(rs!NAMES, "'", "")
               sql = ""
               sql = "set dateformat dmy insert into d_SharesReport(Sno, Name, IDNo, Type, Amount)"
               sql = sql & " values ('" & pcode & "','" & pname & "','" & idno & "','SHARES','0') "
               oSaccoMaster.ExecuteThis (sql)
            End If
        
            sql = ""
            sql = "update d_SharesReport set Amount='" & TXTshares & "' where sno='" & pcode & "' "
            oSaccoMaster.ExecuteThis (sql)
        
        End If
    
    rs.MoveNext
    Wend
    'sharesload
    MsgBox "Records successfully done", vbInformation
    Exit Sub
SysError:
    MsgBox err.description, vbInformation, Me.Caption
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

prgStatus.Max = 100
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
 On Error GoTo SysError
    Dim lastdate As Date
    Dim lastdateofsale As Date
    Dim pcode As String
    Dim NetPay As Double
    Dim Price As Double
    Dim dy As Integer
    Dim grade As String
    Dim bank As String
    Dim bcode As String
    Dim BBranch As String
    Dim idno As String
    Dim PhoneNo As String
    Dim rsd, rsk As New ADODB.Recordset
    sql = ""
    sql = "set dateformat dmy DELETE FROM d_Bonus "
    Set rs = oSaccoMaster.GetRecordset(sql)
    
    
    sql = ""
    sql = "set dateformat dmy SELECT isnull(Price,0) as Price  FROM d_PriceBonus "
    Set rs = oSaccoMaster.GetRecordset(sql)
    Price = rs!Price
    
    prgStatus.value = 0
        sql = ""
        sql = "set dateformat dmy SELECT s.SNo,s.Names,s.IdNo,s.PhoneNo,s.AccNo,s.Bcode,s.Location, isnull(SUM(QSupplied),0) AS Netpay From d_Milkintake d inner join d_Suppliers s on d.sno=s.sno WHERE   d.TransDate >= '" & DTPstdate & "' and d.TransDate <= '" & DTPedate & "' GROUP BY s.sno, s.names,s.AccNo,s.Bcode,s.Location,s.IdNo,s.PhoneNo ORDER BY s.sno asc"
        
        Set rs = oSaccoMaster.GetRecordset(sql)
        While Not rs.EOF
            prgStatus.Max = rs.RecordCount
            prgStatus.value = rs.AbsolutePosition
            
            PhoneNo = rs!PhoneNo
            idno = rs!idno
            pcode = rs!sno
            NetPay = rs!NetPay
            pname = Replace(rs!NAMES, "'", "")
            bank = rs!ACCNO
            bcode = rs!bcode
            BBranch = rs!Location
            
            'select pcode,ldate,dy,auditdate,audit,grade from ag_paging
            sql = ""
            sql = "set dateformat dmy insert into  d_Bonus (Sno, Name,IdNo, PhoneNo,bank,bcode,branch, Startdate, Enddate, Gross ,Pby,Price,Amount)"
            sql = sql & "values('" & pcode & "','" & pname & "','" & idno & "','" & PhoneNo & "','" & bank & "','" & bcode & "','" & BBranch & "','" & DTPstdate & "','" & DTPedate & "','" & NetPay & "','" & User & "','" & Price & "','" & NetPay * Price & "') "
            oSaccoMaster.ExecuteThis (sql)
        
        rs.MoveNext
        Wend
    'sharesload
    MsgBox "Records successfully done", vbInformation
    Exit Sub
SysError:
    MsgBox err.description, vbInformation, Me.Caption

End Sub


Private Sub Command2_Click()
reportname = "Bonus Report.rpt"
'reportname = "bonusyear.rpt"
    Show_Sales_Crystal_Report STRFORMULA, reportname, ""
End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Command4_Click()
    reportname = "memberssharesreport.rpt"
    Show_Sales_Crystal_Report "", reportname, ""
End Sub

Private Sub Form_Load()
DTPstdate = Format(Get_Server_Date, "dd/mm/yyyy")
DTPstdate = DateSerial(year(DTPstdate) - 1, month(-2), 1)
DTPedate = DateSerial(year(DTPstdate) + 1, month(-1), 1 - 1)

End Sub

