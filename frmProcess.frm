VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmProcess 
   BackColor       =   &H00C0FFFF&
   Caption         =   "Process Payroll"
   ClientHeight    =   8685
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   8655
   LinkTopic       =   "Form1"
   ScaleHeight     =   8685
   ScaleWidth      =   8655
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdupdatecurrforw 
      Caption         =   "Update Carryforward"
      Height          =   375
      Left            =   5880
      TabIndex        =   33
      Top             =   120
      Width           =   1695
   End
   Begin VB.TextBox Txtcreditedac 
      Height          =   285
      Left            =   2520
      TabIndex        =   30
      ToolTipText     =   "a/c to be credited"
      Top             =   5520
      Width           =   975
   End
   Begin VB.TextBox lblcreditedac 
      Enabled         =   0   'False
      Height          =   285
      Left            =   3720
      TabIndex        =   29
      Top             =   5520
      Width           =   2775
   End
   Begin VB.TextBox Txtdebitedac 
      Height          =   285
      Left            =   2520
      TabIndex        =   28
      ToolTipText     =   "a/c to be debited"
      Top             =   6000
      Width           =   975
   End
   Begin VB.TextBox lbldebitedac 
      Enabled         =   0   'False
      Height          =   285
      Left            =   3720
      TabIndex        =   27
      Top             =   6000
      Width           =   2775
   End
   Begin VB.CommandButton Cmds1 
      Height          =   255
      Left            =   3480
      TabIndex        =   26
      Top             =   5550
      Width           =   255
   End
   Begin VB.CommandButton Cmds2 
      Height          =   255
      Left            =   3480
      TabIndex        =   25
      Top             =   6030
      Width           =   255
   End
   Begin VB.CommandButton CMDCFN 
      Caption         =   "Carry Forward Transport Deductions"
      Height          =   375
      Left            =   4560
      TabIndex        =   24
      Top             =   3600
      Width           =   3015
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Include Compulsory Deductions"
      Height          =   375
      Left            =   4440
      TabIndex        =   23
      Top             =   1560
      Width           =   2415
   End
   Begin VB.CommandButton cmdupdatebr 
      Caption         =   "Payroll Update"
      Height          =   375
      Left            =   7080
      TabIndex        =   22
      Top             =   1560
      Width           =   1335
   End
   Begin VB.TextBox txtsubsidy 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   7440
      TabIndex        =   21
      Top             =   720
      Width           =   1215
   End
   Begin VB.CheckBox chksubsidyc 
      Caption         =   "Add Subsidy Based on Current Month on Self Only "
      Height          =   255
      Left            =   0
      TabIndex        =   19
      Top             =   1200
      Width           =   5175
   End
   Begin VB.CheckBox chksubsidyprev 
      Caption         =   "Add Subsidy Based on Previous Month on Self Only "
      Height          =   255
      Left            =   0
      TabIndex        =   18
      Top             =   720
      Width           =   5175
   End
   Begin VB.TextBox txttotal 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   1080
      TabIndex        =   17
      Top             =   8280
      Width           =   2415
   End
   Begin MSComCtl2.DTPicker DTPto 
      Height          =   255
      Left            =   2880
      TabIndex        =   16
      Top             =   7320
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   450
      _Version        =   393216
      Format          =   111083521
      CurrentDate     =   40555
   End
   Begin MSComCtl2.DTPicker DTPfrom 
      Height          =   255
      Left            =   360
      TabIndex        =   15
      Top             =   7320
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   450
      _Version        =   393216
      Format          =   111083521
      CurrentDate     =   40555
   End
   Begin VB.CommandButton cmdtotalmonthlyq 
      Caption         =   "Get The Kgs Periods Total"
      Height          =   375
      Left            =   960
      TabIndex        =   14
      Top             =   7800
      Width           =   2775
   End
   Begin VB.CommandButton cmdcompare 
      Caption         =   "Compare"
      Height          =   375
      Left            =   8760
      TabIndex        =   13
      Top             =   2760
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   8760
      TabIndex        =   12
      Top             =   3600
      Visible         =   0   'False
      Width           =   855
   End
   Begin MSComCtl2.DTPicker DTPEOD 
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   5400
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      Format          =   111083521
      CurrentDate     =   40440
   End
   Begin VB.CommandButton cmdendofday 
      Caption         =   "End Of Day"
      Height          =   375
      Left            =   0
      TabIndex        =   9
      Top             =   6000
      Width           =   1815
   End
   Begin VB.CommandButton cmdCarry 
      Caption         =   "Carry Forward Suppliers Deductions"
      Height          =   375
      Left            =   480
      TabIndex        =   5
      Top             =   3600
      Width           =   2775
   End
   Begin VB.CommandButton cmdProcess 
      Caption         =   "Process"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1560
      Width           =   975
   End
   Begin MSComCtl2.DTPicker dtpProcess 
      Height          =   375
      Left            =   3600
      TabIndex        =   2
      Top             =   120
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   111083521
      CurrentDate     =   40214
   End
   Begin VB.CheckBox chkStop 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Stop Further Updates"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      TabIndex        =   1
      Top             =   1560
      UseMaskColor    =   -1  'True
      Width           =   2415
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   135
      Left            =   0
      TabIndex        =   4
      Top             =   2040
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   238
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin MSComCtl2.DTPicker dtpCarry 
      Height          =   375
      Left            =   1800
      TabIndex        =   6
      Top             =   3000
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CalendarBackColor=   16384
      Format          =   111083521
      CurrentDate     =   40214
   End
   Begin MSComctlLib.ProgressBar ProgressBar2 
      Height          =   135
      Left            =   0
      TabIndex        =   8
      Top             =   4200
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   238
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin MSComCtl2.DTPicker previousp 
      Height          =   375
      Left            =   5280
      TabIndex        =   20
      Top             =   720
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   111083521
      CurrentDate     =   40214
   End
   Begin VB.Label Label51 
      Caption         =   "CR:"
      Height          =   255
      Index           =   1
      Left            =   1920
      TabIndex        =   32
      Top             =   5550
      Width           =   735
   End
   Begin VB.Label Label101 
      Caption         =   "DR:"
      Height          =   255
      Index           =   1
      Left            =   1920
      TabIndex        =   31
      Top             =   6030
      Width           =   735
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFF80&
      Caption         =   "Process the total kilo for the day for seliing to the processor or any debtor."
      Height          =   495
      Left            =   0
      TabIndex        =   10
      Top             =   4440
      Width           =   8895
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFC0&
      Caption         =   "Carry Forward Deductions For Negative Net Pay For Period Ending"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   0
      TabIndex        =   7
      Top             =   2760
      Width           =   8895
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      Caption         =   "Process Payrolls For the Period ending :"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   3630
   End
End
Attribute VB_Name = "frmProcess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCarry_Click()


dtpCarry_Validate True

If dtpCarry > Get_Server_Date Then
    MsgBox "The records for the period ending " & dtpCarry & " has not been processed."
        dtpCarry.SetFocus
    Exit Sub
End If
   
Startdate = DateSerial(year(dtpCarry), month(dtpCarry), 1)
Enddate = DateSerial(year(dtpCarry), month(dtpCarry) + 1, 1 - 1)

ProgressBar1.value = 0
Dim Startdate1 As String
Dim Enddate1 As String
Dim sno As String
Startdate1 = DateSerial(year(dtpCarry), month(dtpCarry) + 1, 1)
Enddate1 = DateSerial(year(dtpCarry), month(dtpCarry) + 1, 28)
'Startdate1 = DateSerial(year(Get_Server_Date), month(Get_Server_Date) - 1, 1)
'Enddate1 = DateSerial(year(Get_Server_Date), month(Get_Server_Date) - 1, 28)
'd_sp_BroughtForward  '01/04/2010','30/04/2010','01/05/2010','31/05/2010','birgen'
sql = ""
'sql = "d_sp_BroughtForward '" & Startdate & "','" & Enddate & "','"
'sql = sql & Startdate1 & "','" & Enddate1 & "','" & user & "'"

'oSaccoMaster.ExecuteThis (sql)

'//start of procedure


Dim NPay  As Currency


Dim RsLessAmount As New ADODB.Recordset

Set RsLessAmount = oSaccoMaster.GetRecordset("set dateformat dmy SELECT *  From d_Payroll Where (NPay < 0) And endofperiod = '" & Enddate & "' order by sno")


Do Until RsLessAmount.EOF
DoEvents
frmProcess.Caption = RsLessAmount.Fields("sno")


    Dim desc As String
    Dim Id  As Double
    Dim Amnt As Currency
    Dim Flag As Double
    Dim TotalDed As Currency
    Dim TotalNetpay As Currency
    Dim TDeductions As Currency
    Dim NetPay As Currency
    Dim Others As Currency
    Dim RsDescription As New ADODB.Recordset
    
    NetPay = IIf(IsNull(RsLessAmount!NPay), 0, RsLessAmount!NPay)
    Others = IIf(IsNull(RsLessAmount!Others), 0, RsLessAmount!Others)
    Dim DeductCusor  As New ADODB.Recordset
     'insert the current next month
    oSaccoMaster.ExecuteThis ("SET DATEFORMAT DMY INSERT INTO d_Supplier_Deduc (SNo, Date_Deduc,[Description],Amount,StartDate,enddate, AuditID,Remarks) values ('" & RsLessAmount!sno & "', '" & Startdate & "','Others',('" & NetPay & "'),'" & Startdate & "','" & Enddate & "','" & User & "','" & UCase(Format(Startdate, "MMMM YYYY")) & "'+' Arrears CF')")
     'insert next month
    oSaccoMaster.ExecuteThis ("SET DATEFORMAT DMY INSERT INTO d_Supplier_Deduc (SNo, Date_Deduc,[Description],Amount,StartDate,enddate, AuditID,Remarks) values ('" & RsLessAmount!sno & "', '" & Startdate1 & "','Others',('" & NetPay * -1 & "'),'" & Startdate1 & "','" & Enddate1 & "','" & User & "','" & UCase(Format(Startdate, "MMMM YYYY")) & "'+' Arrears')")
     'update payroll
    oSaccoMaster.ExecuteThis ("SET DATEFORMAT DMY UPDATE d_Payroll SET NPay=0,Others= Others -('" & NetPay * -1 & "'),TDeductions=TDeductions -('" & NetPay * -1 & "') WHERE SNo='" & RsLessAmount!sno & "' AND endofperiod = '" & Enddate & "'")
     
     '//loop next description
    ' RsDescription.MoveNext
     
     'Loop
maritim:
 
    RsLessAmount.MoveNext
Loop

    'Set DeductCusor = oSaccoMaster.GetRecordset("SELECT     [Description], Amount,[Id] From d_Supplier_Deduc WHERE      (Date_Deduc  BETWEEN '" & Startdate & "' AND '" & Enddate & "' AND SNo = '" & RsLessAmount!sno & "' AND [Description] <> 'Transport' AND [Description] <> 'HShares'  AND [Description] <> 'TMShares') AND Amount > 0  ORDER BY [Id] DESC")
    'Do Until DeductCusor.EOF
    
    
       '  TotalDed = 0
       '  Dim RsTotalDed As New ADODB.Recordset
         
       ' Set RsTotalDed = oSaccoMaster.GetRecordset("select isnull(SUM(Amount),0) from d_supplier_deduc  where (Date_Deduc  BETWEEN '" & Startdate & "'  AND '" & Enddate & "' AND SNo ='" & RsLessAmount!sno & "' AND [Description]='" & DeductCusor!description & "' and id =" & DeductCusor!Id & ") ")
        
       '  TotalNetpay = RsLessAmount!NPay + DeductCusor!amount
       '  NPay = NPay + RsTotalDed.Fields(0)
       ' NetPay = NetPay + DeductCusor!amount
        
       ' If NetPay > 0 Then
            '//you exit from here
        
       '      GoTo maritim
    
       ' Else
    
            ' oSaccoMaster.ExecuteThis ("UPDATE d_supplier_Deduc SET Amount=0,Remarks='C/F '+CONVERT(VARCHAR(150), (" & DeductCusor!amount & "))  WHERE SNo='" & RsLessAmount!sno & "' AND [Id]=" & DeductCusor!Id & "")
       '     oSaccoMaster.ExecuteThis ("UPDATE d_supplier_Deduc SET Amount=0,Remarks='Arrears '+CONVERT(VARCHAR(150), (" & DeductCusor!amount & "))  WHERE SNo='" & RsLessAmount!sno & "' AND [Id]=" & DeductCusor!Id & "")
       '     oSaccoMaster.ExecuteThis ("INSERT INTO d_Supplier_Deduc (SNo, Date_Deduc,[Description],Amount,StartDate,enddate, AuditID,Remarks) values ('" & RsLessAmount!sno & "', '" & Startdate1 & "','" & DeductCusor!description & "',(" & DeductCusor!amount & "),'" & Startdate1 & "','" & Enddate1 & "','" & User & "','Arrears')")
       '
       '     If UCase(DeductCusor!description) = "OTHERS" Then
       '         oSaccoMaster.ExecuteThis ("UPDATE d_Payroll SET NPay=0,TDeductions=TDeductions -" & DeductCusor!amount & ",Others=Others -( " & DeductCusor!amount & ") WHERE SNo='" & RsLessAmount!sno & "' AND endofperiod = '" & Enddate & "'")
       '     End If
    
       ' End If

    'DeductCusor.MoveNext
   'Loop



oSaccoMaster.ExecuteThis ("DELETE FROM D_Supplier_Deduc WHERE  [Description]='' AND Amount=0")



'//end of procedure


MsgBox "Records saved successful!"

'dtpCarry_Validate True
'
'If dtpCarry > Get_Server_Date Then
'    MsgBox "The records for the period ending " & dtpCarry & " has not been processed."
'        dtpCarry.SetFocus
'    Exit Sub
'End If
'
'Startdate = DateSerial(Year(dtpCarry), month(dtpCarry), 1)
'Enddate = DateSerial(Year(dtpCarry), month(dtpCarry) + 1, 1 - 1)
'
'ProgressBar1.value = 0
'Dim Startdate1 As String
'Dim Enddate1 As String
'
'Startdate1 = DateSerial(Year(Get_Server_Date), month(Get_Server_Date), 1)
'Enddate1 = DateSerial(Year(Get_Server_Date), month(Get_Server_Date) + 1, 1 - 1)
''d_sp_BroughtForward  '01/04/2010','30/04/2010','01/05/2010','31/05/2010','birgen'
'sql = ""
'sql = "d_sp_BroughtForward '" & Startdate & "','" & Enddate & "','"
'sql = sql & Startdate1 & "','" & Enddate1 & "','" & User & "'"
'
'oSaccoMaster.ExecuteThis (sql)
'
''//DO IT FOR TRANSPORTERS ALSO
'
'sql = ""
'sql = "d_sp_BroughtForward_transporters '" & Startdate & "','" & Enddate & "','"
'sql = sql & Startdate1 & "','" & Enddate1 & "','" & User & "'"
'
'oSaccoMaster.ExecuteThis (sql)
'
'MsgBox "Records saved successful!"
'
End Sub

Private Sub CMDCFN_Click()
dtpCarry_Validate True

If dtpCarry > Get_Server_Date Then
    MsgBox "The records for the period ending " & dtpCarry & " has not been processed."
        dtpCarry.SetFocus
    Exit Sub
End If
   
Startdate = DateSerial(year(dtpCarry), month(dtpCarry), 1)
Enddate = DateSerial(year(dtpCarry), month(dtpCarry) + 1, 1 - 1)

ProgressBar1.value = 0
Dim Startdate1 As String
Dim Enddate1 As String
Dim damoun As Double
Startdate1 = DateSerial(year(Get_Server_Date), month(Get_Server_Date), 1)
Enddate1 = DateSerial(year(Get_Server_Date), month(Get_Server_Date) + 1, 1 - 1)
'd_sp_BroughtForward  '01/04/2010','30/04/2010','01/05/2010','31/05/2010','birgen'
sql = ""
sql = "SET dateformat DMY SELECT     distinct d_TransportersPayRoll.code, d_TransportersPayRoll.NetPay      "
sql = sql & " FROM         d_TransportersPayRoll   inner join d_transport on "
 sql = sql & " d_TransportersPayRoll.code=d_transport.trans_code WHERE     (d_TransportersPayRoll.NetPay < 0)"
sql = sql & " AND d_TransportersPayRoll.endperiod = '" & Enddate & "'    and  d_transport.active=1 "
sql = sql & "  order by code "
Set rs = oSaccoMaster.GetRecordset(sql)
Dim Rs1 As New ADODB.Recordset

Dim NetPay As Double
NetPay = 0
    While Not rs.EOF
    DoEvents
    frmProcess.Caption = rs.Fields("code")
        If Not IsNull(rs.Fields(0)) And rs.Fields(0) <> "" Then
            'sql = "SEt dateformat dmy SELECT Description, Amount, id FROM  d_Transport_Deduc WHERE (TDate_Deduc BETWEEN '" & Startdate & "' AND '" & Enddate & "') AND transcode = '" & Trim(rs.Fields(0)) & "' AND (Amount > 0)  ORDER BY Description DESC"
            'Set Rs1 = oSaccoMaster.GetRecordset(sql)
            NetPay = rs.Fields(1)
        End If
    'Totaldeductions
    'insert the current next month
    sql = "INSERT INTO d_Transport_Deduc (transcode, tdate_deduc,[Description],Amount,StartDate,enddate, AuditID,remarks) values "
    sql = sql & "('" & rs.Fields(0) & "', '" & Startdate & "','Others','" & NetPay & "','" & Startdate & "','" & Enddate & "','" & User & "','" & UCase(Format(Startdate, "MMMM YYYY")) & "'+' Arrears CF')" '
    oSaccoMaster.ExecuteThis (sql)
    
    'insert next month
     
    sql = "INSERT INTO d_Transport_Deduc (transcode, tdate_deduc,[Description],Amount,StartDate,enddate, AuditID,remarks) values "
    sql = sql & "('" & rs.Fields(0) & "', '" & Startdate1 & "','Others','" & NetPay * -1 & "','" & Startdate1 & "','" & Enddate1 & "','" & User & "','" & UCase(Format(Startdate, "MMMM YYYY")) & "'+' Arrears')" '
    oSaccoMaster.ExecuteThis (sql)
    
    'update payroll
    oSaccoMaster.ExecuteThis ("SET DATEFORMAT DMY UPDATE d_TransportersPayRoll SET NetPay=0,Others= Others -('" & NetPay * -1 & "'),Totaldeductions= Totaldeductions-('" & NetPay * -1 & "') WHERE code='" & rs.Fields(0) & "' AND '" & Enddate & "' = '" & Enddate & "'")
     
    

'
'            If Not Rs1.EOF Then
'            While NetPay < 0
'            DoEvents
'
'                 NetPay = NetPay + CDbl(Rs1.Fields(1))
'
'                    If NetPay > 0 Then
'                    sql = "UPDATE d_Transport_Deduc SET Amount=" & NetPay & ",description='C/F '+CONVERT(VARCHAR(150), (" & CDbl(Rs1.Fields(1)) - NetPay & "))  WHERE transcode='" & rs.Fields(0) & "' AND [Id]=" & Rs1.Fields(2) & " "
'                    oSaccoMaster.ExecuteThis (sql)
'                    damoun = (CDbl(Rs1.Fields(1)) - NetPay)
'                        If damoun <> 0 Then
'                            sql = " INSERT INTO d_Transport_Deduc (transcode,  tdate_deduc,[Description],Amount,StartDate,enddate, AuditID,remarks)"
'                            sql = sql & " values ('" & rs.Fields(0) & "', '" & Startdate1 & "','" & Rs1.Fields(0) & "'," & (CDbl(Rs1.Fields(1)) - NetPay) & ",'" & Startdate1 & "','" & Enddate1 & "','" & User & "','Brought Forward')"
'                        End If
'                    oSaccoMaster.ExecuteThis (sql)
'                    End If
'
'                    If UCase(Rs1.Fields(0)) = "OTHERS" Then
'                    'UPDATE d_TransportersPayRoll SET NetPay=0,Totaldeductions=Totaldeductions-(Amnt-NPay),Others=NPay WHERE code=SNo AND '" & Enddate & "' = Enddate
'                    sql = "UPDATE d_TransportersPayRoll SET NetPay=0,Totaldeductions=Totaldeductions-" & (CDbl(Rs1.Fields(1)) - NetPay) & ",Others=" & NetPay & " WHERE code='" & rs.Fields(0) & "' AND '" & Enddate & "' = '" & Enddate & "'"
'                    End If
'
'                    sql = "UPDATE d_Transport_Deduc SET Amount=0,description='C/F '+CONVERT(VARCHAR(150), " & CDbl(Rs1.Fields(1)) & ")  WHERE transcode='" & rs.Fields(0) & "' AND [Id]=" & Rs1.Fields(2)
'                    oSaccoMaster.ExecuteThis (sql)
'                    damoun = Rs1.Fields(1)
'                    If damoun <> 0 Then
'                        sql = "INSERT INTO d_Transport_Deduc (transcode, tdate_deduc,[Description],Amount,StartDate,enddate, AuditID,remarks) values "
'                        sql = sql & "('" & rs.Fields(0) & "', '" & Startdate1 & "','" & Rs1.Fields(0) & "'," & Rs1.Fields(1) & ",'" & Startdate1 & "','" & Enddate1 & "','" & User & "','Brought Forward')" '
'                        oSaccoMaster.ExecuteThis (sql)
'                    End If
'
'                    If UCase(Rs1.Fields(0)) = "OTHERS" Then
'                    sql = "UPDATE d_TransportersPayRoll SET NetPay=0,Totaldeductions=Totaldeductions-" & (CDbl(Rs1.Fields(1))) & ",Others=" & Rs1.Fields(1) & " WHERE code='" & rs.Fields(0) & "' AND '" & Enddate & "' = '" & Enddate & "'"
'
'                    oSaccoMaster.ExecuteThis (sql)
'
'                    End If
'            Rs1.MoveNext
'
'            Wend
'    End If
    rs.MoveNext
    Wend

Set rs = New ADODB.Recordset
Set rs = oSaccoMaster.GetRecordset("select id from d_Transport_Deduc where amount=0")
While Not rs.EOF
oSaccoMaster.ExecuteThis ("DELETE FROM d_Transport_Deduc WHERE   Amount=0 and id=" & rs.Fields(0) & " ")
frmProcess.Caption = rs.Fields(0)
rs.MoveNext
Wend

MsgBox "Records saved successful!"
End Sub

Private Sub cmdcompare_Click()
On Error GoTo ErrorHandler

Set rs = oSaccoMaster.GetRecordset("SELECT     SNo, AccNo, Bcode, BBranch  FROM         d_Suppliers  where sno>=4242 and sno<=4469 ORDER BY SNo")
While Not rs.EOF
DoEvents

Set rst = oSaccoMaster.GetRecordset("SELECT     sno,accno,bank,branch,idno  FROM         Sheet11 where sno=" & rs.Fields(0) & "")
If Not rst.EOF Then

If Trim(rs.Fields(1)) <> Trim(rst.Fields(1)) Then
sql = ""
sql = "update d_suppliers set ACCNO='" & rst.Fields(1) & "',BCODE='" & rst.Fields(2) & "',BBRANCH='" & rst.Fields(3) & "',IDNO='" & rst.Fields(4) & "' where sno=" & rs.Fields(0) & ""
oSaccoMaster.ExecuteThis (sql)
End If
End If
rs.MoveNext
Wend


Exit Sub
ErrorHandler:
MsgBox err.description
End Sub

Private Sub cmdendofday_Click()
On Error GoTo ErrorHandler
Dim totalkilo As Double
Dim dipping As Double

'get the total kilo for the day
  Set rs = New ADODB.Recordset
    sql = "d_sp_DailyTotal '" & DTPEOD & "'"
    Set rs = oSaccoMaster.GetRecordset(sql)
    If Not rs.EOF Then
    totalkilo = IIf(IsNull(rs.Fields(0)), 0, rs.Fields(0))
    Else
    totalkilo = 0
    End If
    '//check if milk is available
    If totalkilo = 0 Then
    MsgBox ("No milk has been received for this day; kindly choose another date"), vbInformation, "EASYMA=END OF DAY"
    Exit Sub
    End If
        If Txtdebitedac = "" Then
            MsgBox "please input the account to be debited"
            Txtdebitedac.SetFocus
        Exit Sub
        End If
        If Txtcreditedac = "" Then
            MsgBox "please input the account to be credited"
            Txtcreditedac.SetFocus
        Exit Sub
        End If
    '//get dipping
    sql = "SELECT     TOP 1 dipping  From d_dispatch ORDER BY Transdate DESC, ID DESC"
    Set rs = oSaccoMaster.GetRecordset(sql)
    If Not rs.EOF Then
    dipping = rs.Fields(0) + totalkilo
    Else
    dipping = totalkilo
    End If
    'validate the milk available for intake
    sql = ""
    sql = "SET  dateformat dmy  SELECT ID, Intake FROM d_dispatch    WHERE     transdate = '" & DTPEOD & "'"
    Set rs = oSaccoMaster.GetRecordset(sql)
    If rs.EOF Then
sql = ""
sql = "set dateformat dmy INSERT INTO d_dispatch (Transdate, descrip, Intake, dipping, dispatch, auditid, auditdate)values ('" & DTPEOD.value & "','Intake'," & totalkilo & "," & dipping & ",0,'" & User & "','" & Get_Server_Date & "')"
'sql = "set dateformat dmy insert_d_dispatch '" & DTPEOD.value & "','Intake'," & totalkilo & "," & dipping + totalkilo & ",0,'" & User & "','" & Get_Server_Date & "'"
oSaccoMaster.ExecuteThis (sql)
Dim Price As Double
Price = 0
sql = "select price from d_price"
Set rs = oSaccoMaster.GetRecordset(sql)
Price = rs!Price
If Not Save_GLTRANSACTION(DTPEOD, totalkilo * Price, Txtdebitedac, Txtcreditedac, "milk purchase", "eod", User, ErrorMessage, "close of day", 1, 1, "intake & Get_Server_Date", transactionNo, "", "", 0) Then
If ErrorMessage <> "" Then
MsgBox err.description, vbInformation, "end of day"
End If
End If
MsgBox "Close of day sucessfully updated"
Exit Sub
Else
sql = ""
sql = "set dateformat dmy UPDATE d_dispatch  SET Intake =" & totalkilo & ", dipping =" & totalkilo & "  WHERE     (Transdate = '" & DTPEOD & "')"
oSaccoMaster.ExecuteThis (sql)
MsgBox "Close of day sucessfully updated"
Exit Sub
End If
Exit Sub
ErrorHandler:
MsgBox err.description
End Sub
Private Sub cmdprocess_Click()

'addomittedentried
'd_sp_PresetDeductAssign  StartDate varchar(10),'" & Enddate & "' varchar(10) , Year bigint, User varchar(35) AS
Dim Yr As Integer, ym As Integer
Dim currdate As Date
Dim rsself As New ADODB.Recordset
Dim rshast As New ADODB.Recordset
Dim rsquality As New ADODB.Recordset
Dim rsts As New Recordset
Dim rsts2 As New Recordset
Dim shareamt As Double
Dim TShares As Double
Dim qual As String
Dim rss As New Recordset
Dim cann As String
Dim rsrate As New ADODB.Recordset
Dim netamt As Double
Dim samount As Double
Dim snum As Long
Dim rskilo As New ADODB.Recordset
Dim qsupp As Double
Dim qrate As Double
currdate = Format(Get_Server_Date, "dd/mm/yyyy")
Startdate = DateSerial(year(dtpProcess), month(dtpProcess), 1)
Enddate = DateSerial(year(dtpProcess), month(dtpProcess) + 1, 1 - 1)
ProgressBar1.value = 0
Yr = year(dtpProcess)
ym = DateDiff("d", Startdate, currdate)
'vbHourglass
'Fixed deductions update
'//update deduction before anything else before running the payroll start here
Dim rshast1 As New ADODB.Recordset, rstq As New Recordset, rshast3 As New ADODB.Recordset, rshast4 As New ADODB.Recordset, descrip As String, remark As String, TransCode As String, amt As Double, sno As Long
Set rshast1 = oSaccoMaster.GetRecordset("set dateformat dmy select * from d_supplier_standingorder  where enddate>='" & Enddate & "'and description ='SHARES' order by sno")
While Not rshast1.EOF
DoEvents
sno = rshast1.Fields("sno")
remark = Trim(rshast1.Fields("remarks"))
'remark = rshast.Fields("remarks")
amt = rshast1.Fields("amount")
sql = ""
sql = "select description,remarks from d_supplier_deduc where sno='" & sno & "' and description ='SHARES' and remarks='" & remark & "' and month(date_deduc)=" & month(Enddate) & " and year(date_deduc)=" & year(Enddate) & ""
Set rst = oSaccoMaster.GetRecordset(sql)
        If rst.EOF Then
        frmProcess.Caption = sno
        sql = "SET dateformat dmy SELECT     SUM(QSupplied) AS QNTY, SUM(QSupplied * (PPU)) AS GrossPay From d_Milkintake " _
        & " WHERE     (TransDate BETWEEN '" & Startdate & "'  AND '" & Enddate & "' AND SNo ='" & sno & "')"
        Set rstq = oSaccoMaster.GetRecordset(sql)
        If rstq!qnty > 50 Then
            If amt > 0 Then
            sql = "d_sp_SupplierDeduct " & sno & ",'" & Enddate & "','SHARES'," & amt & ",'" & Startdate & "','" & Enddate & "'," & year(Enddate) & ",'" & User & "','" & remark & "',''"
            oSaccoMaster.ExecuteThis (sql)
            End If
            Else
            End If
            Else
        End If
        amt = 0
rshast1.MoveNext
Wend
''**********************DAIRYBOARD**********************'
'sql = ""
'sql = "set dateformat dmy SELECT     distinct sno  FROM   d_milkintake where month(transdate)=" & month(Enddate) & " and year(transdate)=" & year(Enddate) & " ORDER BY sno"
'Set rst = oSaccoMaster.GetRecordset(sql)
'While Not rst.EOF
'DoEvents
'sno = rst.Fields("sno")
'sql = "delete from d_supplier_deduc where sno='" & sno & "' and description ='KDBC' and remarks='' and month(date_deduc)=" & month(Enddate) & " and year(date_deduc)=" & year(Enddate) & ""
'
'        Set rst2 = oSaccoMaster.GetRecordset(sql)
'sql = "select description,remarks from d_supplier_deduc where sno='" & sno & "' and description ='KDBC' and remarks='' and month(date_deduc)=" & month(Enddate) & " and year(date_deduc)=" & year(Enddate) & ""
'
'        Set rst2 = oSaccoMaster.GetRecordset(sql)
'If rst2.EOF Then
'            'If rst.EOF Then
'            frmProcess.Caption = sno
'            sql = "SET dateformat dmy SELECT     SUM(QSupplied) AS QNTY, SUM(QSupplied * 0.7) AS GrossPay From d_Milkintake " _
'            & " WHERE     (TransDate BETWEEN '" & Startdate & "'  AND '" & Enddate & "' AND SNo ='" & sno & "')"
'            Set rstq = oSaccoMaster.GetRecordset(sql)
'            If rstq!GrossPay > 0 Then
'             amt = rstq!GrossPay
'           ' If amt > 0 Then
'            sql = "d_sp_SupplierDeduct " & sno & ",'" & Enddate & "','KDBC'," & amt & ",'" & Startdate & "','" & Enddate & "'," & year(Enddate) & ",'" & User & "','',''"
'            oSaccoMaster.ExecuteThis (sql)
'            End If
'            'Else
'            'End If
'            Else
'        End If
'        amt = 0
'rst.MoveNext
'Wend

'********************END DAIRY BOARD***********************'
''**********************SHARES**********************'
sql = ""

sql = "set dateformat dmy SELECT     distinct sno  FROM   d_milkintake where month(transdate)=" & month(Enddate) & " and year(transdate)=" & year(Enddate) & " ORDER BY sno"
Set rst = oSaccoMaster.GetRecordset(sql)
While Not rst.EOF
DoEvents
    sno = rst.Fields("sno")
    '********CHECK SHARE AMOUNT********'
    'If sno = 28 Then
    'MsgBox "here"
    'End If
    sql = "set dateformat dmy select * from d_suppliers where sno='" & sno & "' and shares=1"
    Set rsts2 = oSaccoMaster.GetRecordset(sql)
    If Not rsts2.EOF Then
    
    
    Set rsts = oSaccoMaster.GetRecordset("SELECT    SUM(Amount) AS amtt From d_sconribution WHERE     (transdescription LIKE '%shares%') AND (SNo = '" & sno & "')")
    If Not rsts.EOF Then
    shareamt = IIf(IsNull(rsts!amtt), 0, rsts!amtt)
    End If
    Set rss = oSaccoMaster.GetRecordset("SELECT    SUM(Amount) AS amt From d_supplier_deduc WHERE     (Description LIKE '%shares%') AND (SNo = '" & sno & "')")
    If Not rss.EOF Then
    TShares = IIf(IsNull(rss!amt), 0, rss!amt) + shareamt
    
    End If
    
 Set RShares1 = oSaccoMaster.GetRecordset("SELECT SNo FROM d_SharesReport WHERE Sno = '" & sno & "'")
 If RShares1.EOF Then
 Dim namecheck As String
 namecheck = Replace(rsts2!NAMES, "'", "")
    sql = ""
    sql = "set dateformat dmy insert into d_SharesReport(Sno, Name, IDNo, Type, Amount)"
    sql = sql & " values ('" & sno & "','" & namecheck & "','" & rsts2!idno & "','SHARES','0') "
    oSaccoMaster.ExecuteThis (sql)
 End If
    
    If TShares >= 20000 Then
    Else
    
    '***************END*******************'
    sql = "delete from d_supplier_deduc where sno='" & sno & "' and description ='SHARES' and remarks='' and month(date_deduc)=" & month(Enddate) & " and year(date_deduc)=" & year(Enddate) & ""
    
    Set rst2 = oSaccoMaster.GetRecordset(sql)
    sql = "select description,remarks from d_supplier_deduc where sno='" & sno & "' and description ='SHARES' and remarks='' and month(date_deduc)=" & month(Enddate) & " and year(date_deduc)=" & year(Enddate) & ""
    
     Set rst2 = oSaccoMaster.GetRecordset(sql)
     If rst2.EOF Then
                'If rst.EOF Then
                frmProcess.Caption = sno
                sql = "SET dateformat dmy SELECT     SUM(QSupplied) AS QNTY, SUM(QSupplied * 1) AS GrossPay From d_Milkintake " _
                & " WHERE     (TransDate BETWEEN '" & Startdate & "'  AND '" & Enddate & "' AND SNo ='" & sno & "')"
                Set rstq = oSaccoMaster.GetRecordset(sql)
                If rstq!GrossPay > 0 Then
                 amt = rstq!GrossPay
               ' If amt > 0 Then
                sql = "d_sp_SupplierDeduct " & sno & ",'" & Enddate & "','SHARES'," & amt & ",'" & Startdate & "','" & Enddate & "'," & year(Enddate) & ",'" & User & "','',''"
                oSaccoMaster.ExecuteThis (sql)
                End If
                'Else
                'End If
                Else
            End If
            End If
       End If
       
    sql = ""
    sql = "set dateformat dmy update d_SharesReport set Amount='" & TShares + amt & "' where sno='" & sno & "'"
    oSaccoMaster.ExecuteThis (sql)
       
        amt = 0
    rst.MoveNext
Wend

'********************END SHARES***********************'

'start other suppliers stos'
Set rshast3 = oSaccoMaster.GetRecordset("set dateformat dmy select * from d_supplier_standingorder  where enddate>='" & Enddate & "' and description ='Agrovet' order by sno")
While Not rshast3.EOF
DoEvents
sno = rshast3.Fields("sno")
remark = Trim(rshast3.Fields("remarks"))
'remark = rshast.Fields("remarks")
amt = rshast3.Fields("amount")
sql = ""
sql = "select description,remarks from d_supplier_deduc where sno='" & sno & "' and description ='Agrovet' and remarks='" & remark & "' and month(date_deduc)=" & month(Enddate) & " and year(date_deduc)=" & year(Enddate) & ""
Set rst = oSaccoMaster.GetRecordset(sql)
        If rst.EOF Then
        frmProcess.Caption = sno
            If amt > 0 Then
            sql = "d_sp_SupplierDeduct " & sno & ",'" & Enddate & "','Agrovet'," & amt & ",'" & Startdate & "','" & Enddate & "'," & year(Enddate) & ",'" & User & "','" & remark & "',''"
            oSaccoMaster.ExecuteThis (sql)
            End If
            Else
            
        End If
        amt = 0
rshast3.MoveNext
Wend
'end'

'start transporters'
Set rshast4 = oSaccoMaster.GetRecordset("set dateformat dmy select * from d_transport_standingorder  where enddate>='" & Enddate & "' and description ='Agrovet'")
While Not rshast4.EOF
DoEvents
TransCode = rshast4.Fields("TransCode")
remark = Trim(rshast4.Fields("description"))
'remark = rshast.Fields("remarks")
amt = rshast4.Fields("amount")
sql = ""
sql = "select description,remarks from d_Transport_Deduc where TransCode='" & TransCode & "' and description ='Agrovet' and remarks='" & remark & "' and month(tdate_deduc)=" & month(Enddate) & " and year(tdate_deduc)=" & year(Enddate) & ""
Set rst = oSaccoMaster.GetRecordset(sql)
        If rst.EOF Then
        frmProcess.Caption = TransCode
            If amt > 0 Then
            sql = "d_sp_TransDeduct " & TransCode & ",'" & Enddate & "','Agrovet'," & amt & ",'" & Startdate & "','" & Enddate & "','" & User & "'"
            oSaccoMaster.ExecuteThis (sql)
            End If
            Else
            
        End If
        amt = 0
rshast4.MoveNext
Wend


'end'
'**************end here
ProgressBar1.value = 20

sql = ("d_sp_PresetDeductAssign '" & Startdate & "','" & Enddate & "'," & Yr & ",'" & User & "'")
oSaccoMaster.ExecuteThis (sql)
ProgressBar1.value = 30

sql = " SET      DATEFORMAT DMY SELECT     Sno  From d_Payroll WHERE     EndofPeriod = '" & Enddate & "'  order by sno"
Set rs = oSaccoMaster.GetRecordset(sql)
With rs
    If Not .EOF Then
    While Not .EOF
    DoEvents
    sno = !sno
    frmProcess.Caption = sno
    'If sno = "14" Then MsgBox "here"
        Dim RsShares As New Recordset
        Dim rspu As New Recordset

        
      'update the shares amt
        sql = "select amount,premium from d_shares where sno='" & sno & "' "
        Set rspu = oSaccoMaster.GetRecordset(sql)
        While Not rspu.EOF
        DoEvents
            Dim amt1 As Double
            Dim amt2 As Double
            Dim sp As Double
            amt1 = rspu.Fields("amount") / rspu.Fields("premium")
            oSaccoMaster.ExecuteThis ("update d_shares set spu=" & amt1 & " where amount=" & rspu.Fields("amount") & " and premium='" & rspu.Fields("premium") & "' and sno='" & sno & "'")
        rspu.MoveNext
        Wend

        
        sql = "select sum(spu) as shares from d_shares where sno='" & sno & "'"
        Set RsShares = oSaccoMaster.GetRecordset(sql)
        
            

        
        If Not RsShares.EOF Then
        Shares = IIf(IsNull(RsShares!Shares), 0, RsShares!Shares)
                      If Shares > 900000000000# Then
                      sql = "SET dateformat dmy SELECT     SUM(QSupplied) AS QNTY, SUM(QSupplied * (PPU +1)) AS GrossPay From d_Milkintake " _
                     & " WHERE     (TransDate BETWEEN '" & Startdate & "'  AND '" & Enddate & "' AND SNo ='" & sno & "')"
                     Else
                     sql = "SET dateformat dmy " _
 _
                     & "SELECT SUM(QSupplied) AS QNTY, SUM(pAmount) AS GrossPay " _
                     & "From d_Milkintake WHERE (TransDate BETWEEN '" & Startdate & "'  AND '" & Enddate & "' AND SNo ='" & sno & "')"
                     End If
               Set rs2 = oSaccoMaster.GetRecordset(sql)
               Dim GPay As Double
               GPay = 0
            
                    If Not rs2.EOF Then
                    GPay = IIf(IsNull(rs2!GrossPay), 0, rs2!GrossPay)
                    'rs2!qnty
                    sql = "SET dateformat DMY UPDATE d_Payroll SET GPay=" & GPay & ", KgsSupplied = " & IIf(IsNull(rs2!qnty), 0, rs2!qnty) & " WHERE EndofPeriod = '" & Enddate & "' AND SNo= '" & sno & "'"
                    oSaccoMaster.GetRecordset (sql)
                    Else
                    sql = "SET dateformat DMY UPDATE d_Payroll SET GPay=0, KgsSupplied =0 WHERE EndofPeriod = '" & Enddate & "' AND SNo= '" & !sno & "'"
                    oSaccoMaster.GetRecordset (sql)
                    End If
            Else
            
                        sql = "SET dateformat dmy SELECT     SUM(QSupplied) AS QNTY, SUM(QSupplied * (PPU)) AS GrossPay From d_Milkintake " _
                        & " WHERE     (TransDate BETWEEN '" & Startdate & "'  AND '" & Enddate & "' AND SNo ='" & sno & "')"
                        
                        Set rs2 = oSaccoMaster.GetRecordset(sql)
                       
                        GPay = 0
                        
                        If Not rs2.EOF Then
                        GPay = IIf(IsNull(rs2!GrossPay), 0, rs2!GrossPay)
                        sql = "SET dateformat DMY UPDATE d_Payroll SET GPay=" & GPay & ", KgsSupplied = " & IIf(IsNull(rs2!qnty), 0, rs2!qnty) & " WHERE EndofPeriod = '" & Enddate & "' AND SNo= '" & sno & "'"
                        oSaccoMaster.GetRecordset (sql)
                        Else
                        sql = "SET dateformat DMY UPDATE d_Payroll SET GPay=0, KgsSupplied =0 WHERE EndofPeriod = '" & Enddate & "' AND SNo= '" & !sno & "'"
                        oSaccoMaster.GetRecordset (sql)
                        End If
            
        End If
        Dim agrovet As Double
        Dim FSA As Double
        Dim AI As Double
        Dim Others As Double
        Dim TMShares As Double
        Dim HShares As Double
        Dim Advance As Double
        Dim Transport As Double
        Dim TCHP As Double
        Dim CBO As Double
        Dim MILKREJECTS As Double
        Dim variance As Double
            agrovet = 0
            FSA = 0
            AI = 0
            Others = 0
            TMShares = 0
            HShares = 0
            Advance = 0
            Transport = 0
           TCHP = 0
            CBO = 0
            MILKREJECTS = 0
            variance = 0
            
           Dim rsdeduction As New Recordset
'             If sno = "70" Then
'           MsgBox "here"
'           End If
         
          sql = " set dateformat dmy SELECT  [Description], SUM(Amount) AS Amount " _
          & "From d_Supplier_deduc WHERE  (startDate >= '" & Startdate & "' AND EndDate <= '" & Enddate & "') AND SNo='" & sno & "'" _
                           & "GROUP BY [Description]"
         Set rsdeduction = oSaccoMaster.GetRecordset(sql)
         With rsdeduction
            If Not .EOF Then
                While Not .EOF
                DoEvents
                Dim description  As String
                Dim deduction, TotalDed As Double
                        description = ""
                        deduction = 0
                        TotalDed = 0
                    description = IIf(IsNull(rsdeduction!description), "others", rsdeduction!description)
                    deduction = IIf(IsNull(rsdeduction!amount), 0, rsdeduction!amount)
                    If UCase(description) = "AGROVET" Then
                     agrovet = agrovet + deduction
                    End If
                    If UCase(description) = "NIABA LOAN" Then
                     FSA = FSA + deduction
                     End If
                    If UCase(description) = "AI" Then
                     AI = AI + deduction
                     End If
                    If UCase(description) = "CLINICAL" Then
                     TMShares = TMShares + deduction
                     End If
                    If UCase(description) = "OTHERS" Then
                     Others = Others + deduction
                     End If
                    If UCase(description) = "SHARES" Then
                     HShares = HShares + deduction
                     End If
                    If UCase(description) = "ADVANCE" Then
                     Advance = Advance + deduction
                     End If
                    If UCase(description) = "TRANSPORT" Then
                     Transport = Transport + deduction
                    End If
                    If UCase(description) = "MILKREJECTS" Then
                     TCHP = TCHP + deduction
                    End If
'                    If UCase(description) = "KDBC" Then
'                     KDBC = KDBC + deduction
'                    End If
                    If UCase(description) = "VARIANCE" Then
                    CBO = CBO + deduction
                    End If
                .MoveNext
                 frmProcess.Caption = sno
                Wend
             End If
            TotalDed = agrovet + FSA + AI + Others + TMShares + HShares + Advance + Transport + TCHP + CBO
                
                sql = "SET DATEFORMAT DMY UPDATE    d_Payroll " _
                 & " SET  Transport = " & Transport & ", Agrovet = " & agrovet & ", AI = " & AI & ", TMShares = " & TMShares & ", FSA = " & FSA & ", HShares =" & HShares & ", Advance = " & Advance & ", " _
                                      & " Others = " & Others & ", TDeductions =" & TotalDed & ", NPay = " & GPay - TotalDed & " ,TCHP=" & TCHP & ",CBO=" & CBO & " " _
                & "Where sno = '" & sno & "' And EndofPeriod =  '" & Enddate & "'"
                oSaccoMaster.GetRecordset (sql)
        End With
    .MoveNext
    frmProcess.Caption = sno
    Wend
    End If
End With

'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
ProgressBar1.value = 50
oSaccoMaster.ExecuteThis ("set dateformat dmy UPDATE d_Payroll SET Transport=0,TDeductions=(TDeductions -Transport),NPay=(NPay + Transport) WHERE Endofperiod='" & Enddate & "'")
oSaccoMaster.ExecuteThis ("set dateformat dmy     UPDATE    d_supplier_deduc   SET  amount=0 where [Description] = 'Transport'and  EndDate ='" & Enddate & "'")
'Update transporters
'd_sp_TransUpdate StartDate varchar(10),'" & Enddate & "' varchar(10),User varchar(35) AS
Set rst = oSaccoMaster.GetRecordset("select transcode from d_transporters  order by transcode asc")
While Not rst.EOF
DoEvents
frmProcess.Caption = rst.Fields(0)
oSaccoMaster.ExecuteThis ("d_sp_TransUpdate '" & Startdate & "','" & Enddate & "','" & User & "','" & Trim(rst.Fields(0)) & "'")
rst.MoveNext
Wend
Set rst = Nothing
ProgressBar1.value = 70
'oSaccoMaster.ExecuteThis ("delete from  d_TransportersPayroll WHERE EndPeriod ='" & Enddate & "'")

'oSaccoMaster.ExecuteThis (" exec d_sp_TransPRoll '" & Startdate & "','" & Enddate & "','" & User & "'")
    updatepayroll_deduc
    transproll

'//work on flat rate guys.
sql = ""
sql = "set dateformat dmy delete from d_trans_frate where period='" & dtpProcess & "' "
oSaccoMaster.ExecuteThis (sql)
'get the flat rate guys on board first

Dim tcode As String, Period As Date, rate As Currency, amount As Currency, D As Integer, samson As Integer, NPay As Double
D = Days_In_Month(month(Enddate), year(Enddate))
sql = "SELECT     TransCode, Active, isfrate, rate  FROM         d_Transporters  WHERE     (isfrate = '1') AND (Active = 1)"
Set rs = oSaccoMaster.GetRecordset(sql)
While Not rs.EOF
DoEvents
If Not IsNull(rs.Fields(0)) Then tcode = rs.Fields(0)
If Not IsNull(rs.Fields(3)) Then rate = rs.Fields(3)
samson = Days_In_Month(month(dtpProcess), month(dtpProcess))
amount = rate * samson
'//get total deduction for transporter
Dim rstt As New ADODB.Recordset, tot As Double
Set rstt = oSaccoMaster.GetRecordset("SET              dateformat dmy  SELECT     Totaldeductions    FROM         d_TransportersPayRoll  WHERE     (Code = '" & tcode & "') AND ('" & Enddate & "' = '" & dtpProcess & "')")
If Not rstt.EOF Then
tot = IIf(IsNull(rstt.Fields(0)), 0, rstt.Fields(0))
Else
tot = 0
End If
'delete from this table until all the runnings are done

NPay = amount - tot
sql = ""
sql = "SELECT     Trans_code, Period, rate, days, Amount, auditid, auditdatetime From d_trans_frate where period='" & dtpProcess & "' and trans_code='" & tcode & "'"
Set rst = oSaccoMaster.GetRecordset(sql)
If rst.EOF Then
sql = ""
sql = "set dateformat dmy INSERT INTO d_trans_frate"
sql = sql & " (Trans_code, Period, rate, days, Amount, auditid, auditdatetime,total)"
sql = sql & " VALUES     ('" & tcode & "','" & dtpProcess & "'," & rate & "," & samson & "," & amount & ",'" & User & "','" & Get_Server_Date & "'," & tot & ")"
oSaccoMaster.ExecuteThis (sql)

''update the payroll
sql = ""
sql = "set dateformat dmy update d_TransportersPayRoll set frate=1,grosspay=" & amount & ",netpay=" & NPay & " where code='" & tcode & "' and '" & Enddate & "'='" & dtpProcess & "' "
oSaccoMaster.ExecuteThis (sql)
Else

'sql = ""
'sql="set dateformat dmy update d_trans_frate set
End If

'//update the payroll



rs.MoveNext
Wend

'//do the subsidy for self farmers
    If chksubsidyprev = vbChecked Then
            
            sql = "set dateformat dmy SELECT     distinct sno  FROM   d_milkintake where month(transdate)=" & month(Enddate) & " and year(transdate)=" & year(Enddate) & " ORDER BY sno"
            Set rsself = oSaccoMaster.GetRecordset(sql)
            
            While Not rsself.EOF
            DoEvents
            '//check if it is a self person
            snum = rsself.Fields("sno")
            Set rshast = oSaccoMaster.GetRecordset("select sno from d_suppliers where sno='" & snum & "' and hast=0")
            If Not rshast.EOF Then
            '//get total kilos for the previous months
            Set rskilo = oSaccoMaster.GetRecordset("select sum(qsupplied) from d_milkintake where sno=" & snum & " and month(transdate)=" & month(previousp) & " and year(transdate)=" & year(previousp) & "")
                If Not rskilo.EOF Then
                qsupp = IIf(IsNull(rskilo.Fields(0)), 0, rskilo.Fields(0))
                    If qsupp = 0 Then GoTo sargoi
                    If txtsubsidy = "" Then txtsubsidy = 0
                    samount = CCur(txtsubsidy) * qsupp 'rsself.Fields("amount")
                    Set rs = oSaccoMaster.GetRecordset(" set dateformat dmy SELECT     *   FROM         d_Payroll  WHERE     sno = " & snum & " AND endofperiod='" & Enddate & "'")
                    If Not rs.EOF Then
                        netamt = IIf(IsNull(rs.Fields("npay")), 0, rs.Fields("npay")) + samount
                        '//update payroll net and subsidy
                        sql = "set dateformat dmy update d_payroll set npay=" & netamt & ",subsidy=" & samount & " where sno=" & snum & " and endofperiod='" & Enddate & "'"
                        oSaccoMaster.ExecuteThis (sql)
                    End If
                End If
            End If
            
            
sargoi:
            samount = 0
            qsupp = 0
            rsself.MoveNext
            Wend
    End If
    '//QUALITY CHECK'
            sql = "set dateformat dmy SELECT     distinct sno  FROM   d_milkintake where month(transdate)=" & month(Enddate) & " and year(transdate)=" & year(Enddate) & " ORDER BY sno"
                    Set rsself = oSaccoMaster.GetRecordset(sql)
            
            While Not rsself.EOF
            DoEvents
            snum = rsself.Fields("sno")
            
            'Set rsquality = oSaccoMaster.GetRecordset("select sno,rate from d_quality where sno='" & snum & "'")
            Set rsquality = oSaccoMaster.GetRecordset("select s.sno,s.canno,q.remarks from QBMPS q inner join d_Suppliers s on s.canno= q.canno  where s.sno='" & snum & "'")
            If Not rsquality.EOF Then
            If rsquality.Fields(1) = "" Then
            rsquality.Fields(1) = 0
            End If
            qual = rsquality.Fields(2)
            cann = rsquality.Fields(1)
            Set rsrate = oSaccoMaster.GetRecordset("SELECT     Quality, irate FROM   Qsetup where Quality = '" & qual & "'")
            If Not rsrate.EOF Then
            qrate = IIf(IsNull(rsrate.Fields(1)), 0, rsrate.Fields(1))
            End If
            '//get total kilos for the previous months
            Set rskilo = oSaccoMaster.GetRecordset("select sum(qsupplied) from d_milkintake where sno=" & snum & " and month(transdate)=" & month(previousp) & " and year(transdate)=" & year(previousp) & "")
                If Not rskilo.EOF Then
                'If sno = "1377" Then MsgBox "here"
                qsupp = IIf(IsNull(rskilo.Fields(0)), 0, rskilo.Fields(0))
                    If qsupp = 0 Then GoTo home
                    If qual = "Penalty" Then
                    samount = qrate * qsupp 'rsself.Fields("amount")
                    Else
                    
                    samount = qrate * qsupp 'rsself.Fields("amount")
                    End If
                    Set rs = oSaccoMaster.GetRecordset(" set dateformat dmy SELECT     *   FROM         d_Payroll  WHERE     sno = " & snum & " AND endofperiod='" & Enddate & "'")
                    If Not rs.EOF Then
                        netamt = IIf(IsNull(rs.Fields("npay")), 0, rs.Fields("npay")) + samount
                        '//update payroll net and subsidy
                        sql = "set dateformat dmy update d_payroll set npay=" & netamt & ",Tchp=" & samount & ",Trader='" & qual & "',otheraccno='" & cann & "' where sno=" & snum & " and endofperiod='" & Enddate & "'"
                        oSaccoMaster.ExecuteThis (sql)
                    End If
                End If
            End If
home:
            rsself.MoveNext
            Wend
            'END '
    '//standing orders
    Dim rg As New ADODB.Recordset, snoo As Long, cboamount As Double, cbonetamnt As Double
    Set rg = New ADODB.Recordset
'sql = "set dateformat dmy select * from d_supplier_standingorder  where enddate>='" & Enddate & "' order by sno"
'sql = "SELECT     *  FROM         NewDeductions n INNER JOIN  OtherDed O ON n.ODCode = O.OCode  WHERE     (n.PayrollNo = '" & empcode & "')AND n.ODCode<>'PEN' AND (O.ContribDeduction = 1 and month(n.edate)>=" & mMonth & " and year(n.edate)>=" & yyear & ")"
'Set rg = oSaccoMaster.GetRecordset(sql)
'While Not rg.EOF
'DoEvents
'snoo = rg.Fields("sno")
'cboamount = rg.Fields("amount")
'    sql = ""
    'sql = "SELECT "
    'Set rs = oSaccoMaster.GetRecordset(" set dateformat dmy SELECT     *   FROM         d_Payroll  WHERE     sno = " & snoo & " AND endofperiod='" & Enddate & "'")
'                    If Not rs.EOF Then
'                        'cbonetamnt = IIf(IsNull(rs.Fields("npay")), 0, rs.Fields("npay")) '- cboamount
'                        '//update payroll net and cbo
'                        sql = "set dateformat dmy update d_payroll set cbo=" & cboamount & " where sno=" & snoo & " and endofperiod='" & Enddate & "'"
'                        oSaccoMaster.ExecuteThis (sql)
'                    End If
 'rg.MoveNext
' Wend
    '*****8END IF STANDING ORDERS
Dim Va As Integer

If chkStop.value = vbChecked Then
Va = 1
Else
Va = 0
End If

oSaccoMaster.ExecuteThis ("d_sp_Periods '" & Enddate & "'," & Va & ",'" & User & "'")
ProgressBar1.value = 100

MsgBox "Completed Payroll"
'vbDefault

End Sub



Private Sub Cmds1_Click()
frmSearchGLAccounts.Show vbModal
If SearchValue <> "" Then
Txtcreditedac = SearchValue
sql = ""
sql = "select * from glsetup where accno='" & SearchValue & "'"
Set rs = oSaccoMaster.GetRecordset(sql)
lblcreditedac = rs!GlAccName
End If
End Sub

Private Sub Cmds2_Click()
frmSearchGLAccounts.Show vbModal
If SearchValue <> "" Then
Txtdebitedac = SearchValue
sql = ""
sql = "select * from glsetup where accno='" & SearchValue & "'"
Set rs = oSaccoMaster.GetRecordset(sql)
'lblcreditedac = rs!GlAccName
End If
End Sub

Private Sub cmdtotalmonthlyq_Click()

sql = ""
sql = "SET              dateformat dmy SELECT     SUM(qsupplied) From d_Milkintake WHERE     transdate BETWEEN '" & DTPfrom & "' AND '" & DTPto & "'"
Set rs = oSaccoMaster.GetRecordset(sql)
If Not rs.EOF Then
txttotal = IIf(IsNull(rs.Fields(0)), 0, rs.Fields(0))
End If
End Sub

Private Sub cmdupdatebr_Click()
Startdate = DateSerial(year(dtpProcess), month(dtpProcess), 1)
Enddate = DateSerial(year(dtpProcess), month(dtpProcess) + 1, 1 - 1)
Dim sno As String, speriod As Date, eperiod As Date
sql = ""
sql = "delete from d_payroll where mmonth=" & month(dtpProcess) & " and yyear=" & year(dtpProcess) & ""
oSaccoMaster.ExecuteThis (sql)
sql = ""
sql = "delete from   d_TransportersPayRoll where mmonth=" & month(dtpProcess) & " and yyear=" & year(dtpProcess) & ""
oSaccoMaster.ExecuteThis (sql)
sql = "set dateformat dmy select distinct sno from d_milkintake where transdate<='" & Enddate & "' and transdate>='" & Startdate & "'  order by sno"
Set rs = oSaccoMaster.GetRecordset(sql)
While Not rs.EOF
DoEvents
sno = rs.Fields(0)

sql = "select sno from d_payroll where mmonth=" & month(dtpProcess) & " and yyear=" & year(dtpProcess) & " and sno=" & sno & " order by sno"
Set rst = oSaccoMaster.GetRecordset(sql)
    If rst.EOF Then
        sql = ""
        sql = "insert into d_Payroll (SNo,EndofPeriod,auditid ) "
        sql = sql & " values (" & sno & ",'" & dtpProcess & "','" & User & "' )"
        oSaccoMaster.ExecuteThis (sql)
        
    End If
    frmProcess.Caption = sno
    rs.MoveNext
Wend
'gaa:
MsgBox "records updated successfully", vbInformation

End Sub

Private Sub cmdupdatecurrforw_Click()
'Startdate = DateSerial(year(dtpProcess), month(dtpProcess), 1)
'Enddate = DateSerial(year(dtpProcess), month(dtpProcess) + 1, 1 - 1)
'Dim sno As String, speriod As Date, eperiod As Date
'sql = ""
''sql = "set dateformat dmy select distinct sno from d_milkintake where transdate<='" & Enddate & "' and transdate>='" & Startdate & "'  order by sno"
'sql = "select sno  FROM         d_Suppliers ORDER BY SNo"
'Set rs = oSaccoMaster.GetRecordset(sql)
'While Not rs.EOF
'DoEvents
''If sno = "" Then
''GoTo gaa
'
'If rs!sno <> "" Then
'sno = rs.Fields(0)
'
''sql = "select sno from d_payroll where mmonth=" & month(dtpProcess) & " and yyear=" & year(dtpProcess) & " and sno=" & sno & " order by sno"
''Set rst = oSaccoMaster.GetRecordset(sql)
'   ' If rst.EOF Then
''        sql = ""
''        sql = "insert into d_Payroll (SNo,EndofPeriod,auditid ) "
''        sql = sql & " values (" & sno & ",'" & dtpProcess & "','" & User & "' )"
'        sql = "d_sp_MilkIntake " & sno & ",'" & Startdate & "',0,0,0,'" & Time & "','" & User & "','update'"
'        oSaccoMaster.ExecuteThis (sql)
'
'    'End If
'    frmProcess.Caption = sno
'    rs.MoveNext
'
'   End If
''End If
'Wend
''gaa:
'MsgBox "records updated successfully", vbInformation
End Sub

Private Sub Command1_Click()
On Error GoTo ErrorHandler
Dim sno As String
sno = 1
'process_payroll (sno)
'UPDATE THE ACCOUNTS FOR KIPKAREN FSA
sql = "SELECT     *  FROM         FSA_ACCS  WHERE     (Payrollno <> 'NA')  ORDER BY 4"
Set rs = oSaccoMaster.GetRecordset(sql)
While Not rs.EOF
DoEvents
sql = "UPDATE    d_Suppliers   SET              accno='" & rs.Fields(0) & "' where sno='" & Trim(rs.Fields(3)) & "'"
oSaccoMaster.ExecuteThis (sql)
frmProcess.Caption = rs.Fields(0)
rs.MoveNext
Wend
 MsgBox "Done"
 Exit Sub
ErrorHandler:
 MsgBox err.description
End Sub

Private Sub Command2_Click()
Startdate = DateSerial(year(dtpProcess), month(dtpProcess), 1)
Enddate = DateSerial(year(dtpProcess), month(dtpProcess) + 1, 1 - 1)
Dim sno As String, speriod As Date, eperiod As Date
Dim Others As String, Remarks As String, rate As Double, Rated As Integer
Dim deduction As String
Dim Stopped As Integer
sql = ""
sql = "set dateformat dmy select distinct sno from d_milkintake where transdate<='" & Enddate & "' and transdate>='" & Startdate & "'  order by sno"
Set rs = oSaccoMaster.GetRecordset(sql)
While Not rs.EOF
DoEvents
sno = rs.Fields(0)

sql = "select * from d_PreSets where  sno=" & sno & " order by Deduction"
Set rst = oSaccoMaster.GetRecordset(sql)
    If rst.EOF Then
    'SELECT     SNo, Deduction, Remark, StartDate, Rate, Stopped, Auditdatetime, AuditId, Rated
'From d_PreSets
'//OPERATIONS ON OTHERS
        Stopped = 0
        'sno = Rst.Fields("sno")
'        deduction = "Others"
'        Remarks = "OPERATIONS"
        Rated = 1
        'rate = 1.7
        Dim rst1 As New ADODB.Recordset
'        sql = "select sno from d_PreSets where  sno=" & sno & " and deduction='Others' and remark='OPERATIONS' order by Deduction"
'        Set rst1 = oSaccoMaster.GetRecordset(sql)
'        If rst1.EOF Then
'            sql = ""
'            sql = "set dateformat dmy insert into d_PreSets (SNo, Deduction, Remark, StartDate, Rate, Stopped, Auditdatetime, AuditId, Rated) "
'            sql = sql & " values (" & sno & ", '" & deduction & "','" & Remarks & "', '" & Startdate & "', " & rate & ", " & Stopped & ", '" & Get_Server_Date & "', '" & User & "', " & Rated & ")"
'            oSaccoMaster.ExecuteThis (sql)
'        End If

        'HSARES
        Stopped = 0
'        deduction = "HShares"
'        Remarks = ""
        Rated = 1
        'rate = 0.3

        sql = "select sno from d_PreSets where sno=" & sno & " and deduction='HShares' order by Deduction"
        Set rst1 = oSaccoMaster.GetRecordset(sql)
        If rst1.EOF Then
            sql = ""
            sql = "set dateformat dmy insert into d_PreSets (SNo, Deduction, Remark, StartDate, Rate, Stopped, Auditdatetime, AuditId, Rated) "
            sql = sql & " values (" & sno & ", '" & deduction & "','" & Remarks & "', '" & Startdate & "', " & rate & ", " & Stopped & ", '" & Get_Server_Date & "', '" & User & "', " & Rated & ")"
            oSaccoMaster.ExecuteThis (sql)
        End If


    End If
    '//cbo fees

'            sql = "select description from d_supplier_standingorder where sno='" & sno & "' and description ='CBO'"
'            Set rst = oSaccoMaster.GetRecordset(sql)
'            If rst.EOF Then
'            '//Update deductions
'                Set cn = New ADODB.Connection
'                sql = "set dateformat dmy INSERT INTO d_supplier_standingorder"
'                sql = sql & "           (SNo, Date_Deduc, Description, Amount, MaxAmount, Period, StartDate, EndDate, auditid,  yyear, Remarks)"
'                sql = sql & "  VALUES     (" & sno & ",'" & Startdate & "','CBO',50,0,'" & Format(Startdate, "mmm-YYYY") & "','" & Startdate & "','31/05/2015','" & User & "'," & year(Enddate) & ",'" & Remarks & "')"
'                oSaccoMaster.ExecuteThis (sql)
''
''
''
'            End If
''
'
    Stopped = 0
        deduction = ""
        Remarks = ""
        Rated = 1
        rate = 0
'
    frmProcess.Caption = sno
  rs.MoveNext
Wend

End Sub

Private Sub dtpCarry_Validate(Cancel As Boolean)
dtpCarry = DateSerial(year(dtpCarry), month(dtpCarry) + 1, 1 - 1)

End Sub

Private Sub Form_Load()
dtpProcess = DateSerial(year(Get_Server_Date), month(Get_Server_Date) + 1, 1 - 1)
dtpCarry = DateSerial(year(Get_Server_Date), month(Get_Server_Date) + 1, 1 - 1)
DTPEOD = Format(Get_Server_Date, "dd/mm/yyyy")
DTPfrom = DTPEOD
DTPto = DTPEOD
previousp = DTPfrom
'chkStop.Caption = False

Txtcreditedac = "060101"
Txtdebitedac = "080101"
cmdupdatecurrforw.Visible = False

End Sub
Public Sub transproll()
'On Error Resume Next
On Error GoTo milgo

Dim GPay As Double
Dim qnty As Double
Dim tcode As String
Dim Amnt As Double
Dim subsidy As Double
'tcode = "T20"
Set rst = oSaccoMaster.GetRecordset(" SELECT SUM(dbo.d_TransDetailed.qnty) AS QNTY, dbo.d_TransDetailed.Trans_Code AS Code, SUM(dbo.d_TransDetailed.Amount) AS Amount,SUM(d_TransDetailed.Subsidy) As Subsidy From d_TransDetailed WHERE     EndPeriod = '" & Enddate & "'  GROUP BY d_TransDetailed.Trans_Code")
While Not rst.EOF
'TCode = "T308"
    DoEvents
    tcode = rst.Fields("code")
    frmProcess.Caption = tcode
    'If Trim$(tcode) = "T20" Then MsgBox "HERE"
    subsidy = IIf(IsNull(rst.Fields("subsidy")), 0, rst.Fields("subsidy"))
    Amnt = IIf(IsNull(rst.Fields("amount")), 0, rst.Fields("amount"))
    GPay = IIf(IsNull(Amnt + subsidy), 0, (Amnt + subsidy))
    qnty = IIf(IsNull(rst.Fields("qnty")), 0, rst.Fields("qnty"))
    tcode = IIf(IsNull(rst.Fields("Code")), 0, rst.Fields("Code"))
    
    oSaccoMaster.ExecuteThis ("exec d_sp_UpdateTransPay '" & tcode & "', '" & qnty & "'," & Amnt & "," & subsidy & "," & GPay & ", '" & Enddate & "','" & User & "'")
    
Dim agrovet As Double
Dim FSA  As Double
Dim AI As Double
Dim Others As Double
Dim TMShares As Double
Dim HShares As Double
Dim Advance As Double
Dim TotalDed As Double
Dim MILKREJECTS As Double
Dim variance As Double

agrovet = 0
FSA = 0
AI = 0
Others = 0
TMShares = 0
HShares = 0
Advance = 0
variance = 0
MILKREJECTS = 0

Dim desc As String
Dim deduction As Double

    Set rst2 = oSaccoMaster.GetRecordset(" set dateformat dmy SELECT  [Description], SUM(Amount) AS Amount From d_Transport_Deduc WHERE  (startdate>='" & Startdate & "'and enddate<='" & Enddate & "') and TransCode='" & tcode & "'  GROUP BY [Description]")

 
    While Not rst2.EOF
    DoEvents
    'frmProcess.Caption = rst2.Fields("TransCode")
        desc = IIf(IsNull(rst2.Fields("description")), "others", rst2.Fields("description"))
        deduction = IIf(IsNull(rst2.Fields("amount")), 0, rst2.Fields("amount"))
        
        If UCase(desc) = "AGROVET" Then
        agrovet = deduction
        
        ElseIf UCase(desc) = "NIABA LOAN" Then
        FSA = deduction
        
        ElseIf UCase(desc) = "AI" Then
        AI = deduction
        
        ElseIf UCase(desc) = "TMSHARES" Then
        TMShares = deduction
        
        ElseIf UCase(desc) = "OTHERS" Then
        Others = deduction
        
        ElseIf UCase(desc) = "CLINICAL" Then
        HShares = deduction
        
        ElseIf UCase(desc) = "ADVANCE" Then
        Advance = deduction
        ElseIf UCase(desc) = "VARIANCE" Then
        variance = deduction
        ElseIf UCase(desc) = "MILKREJECTS" Then
        MILKREJECTS = deduction
        
        End If
        
        TotalDed = agrovet + FSA + AI + TMShares + Others + HShares + Advance + variance + MILKREJECTS
        
        oSaccoMaster.ExecuteThis ("exec d_sp_UpdateTransDed  '" & tcode & "',' " & Enddate & "'," & TotalDed & "," & agrovet & "," & AI & "," & TMShares & "," & FSA & "," & HShares & "," & Advance & "," & Others & "," & variance & "," & MILKREJECTS & "")
 
    
    rst2.MoveNext
    Wend
    TotalDed = agrovet + FSA + AI + TMShares + Others + HShares + Advance + variance + MILKREJECTS
        
     oSaccoMaster.ExecuteThis ("exec d_sp_UpdateTransDed  '" & tcode & "',' " & Enddate & "'," & TotalDed & "," & agrovet & "," & AI & "," & TMShares & "," & FSA & "," & HShares & "," & Advance & "," & Others & "," & variance & "," & MILKREJECTS & "")

    rst.MoveNext
    Wend


Exit Sub

milgo:
    MsgBox err.description, vbInformation, tcode
    Exit Sub
    
End Sub

Public Sub Transportertransporter()
Dim ptrans As String
Dim tt As String
Dim rate As Integer

Set rst = oSaccoMaster.GetRecordset("select transcode,ptransporter,tt from d_transporters where tt=1")
While Not rst.EOF
DoEvents
rst.MoveNext
Wend

2
End Sub

Public Sub transupdate()

End Sub

Public Sub updatepayroll_deduc()

  

Set rst = oSaccoMaster.GetRecordset("select sno,sum(amount)as amount from d_transdetailed where endperiod='" & dtpProcess & "' and amount>0 group by sno order by sno asc")
While Not rst.EOF
DoEvents
frmProcess.Caption = rst.Fields("sno")
'If Rst.Fields("sno") = "432" Then MsgBox "here"
    
    oSaccoMaster.ExecuteThis ("set dateformat dmy UPDATE d_Payroll SET Transport=" & rst.Fields("amount") & " WHERE SNo='" & rst.Fields("sno") & "' AND Endofperiod='" & dtpProcess & "' ")
    oSaccoMaster.ExecuteThis ("set dateformat dmy update d_payroll set TDeductions=TDeductions +" & rst.Fields("amount") & ",NPay=gpay - tdeductions  WHERE SNo='" & rst.Fields("sno") & "' AND Endofperiod='" & dtpProcess & "'  ")
    oSaccoMaster.ExecuteThis ("set dateformat dmy update d_payroll set NPay=gpay - tdeductions  WHERE SNo='" & rst.Fields("sno") & "' AND Endofperiod='" & dtpProcess & "'  ")
    oSaccoMaster.ExecuteThis ("set dateformat dmy   insert into  d_supplier_deduc  (sno,Date_Deduc,[Description],Amount,StartDate,EndDate,auditid) values('" & rst.Fields("sno") & "','" & Startdate & "','Transport'," & rst.Fields("amount") & ",'" & Startdate & "','" & Enddate & "' ,'" & User & "')  ")
    rst.MoveNext
Wend

'Set Rst1 = oSaccoMaster.GetRecordset("set dateformat dmy select code from d_transporterspayroll where endperiod='" & dtpProcess & "' order by code asc")
'While Not Rst1.EOF
'DoEvents
'    frmProcess.Caption = Rst1.Fields(0)
'    If Trim$(Rst1.Fields(0)) = "T008" Then MsgBox ""
'    oSaccoMaster.ExecuteThis ("update d_transporterspayroll set netpay=grosspay-totaldeductions where code='" & Rst1.Fields(0) & "'")
'Rst1.MoveNext
'Wend
End Sub

Public Sub Refresh_TranportersPayroll()
Dim year As Integer
Dim month As Integer


oSaccoMaster.ExecuteThis ("delete from d_TransportersPayRoll where  yyear=year('" & dtpProcess & "') and mmonth =month('" & dtpProcess & "')")

Set rst = oSaccoMaster.GetRecordset("set dateformat dmy select distinct Trans_Code from d_Transport where sno in (select sno from d_milkintake where (month(transdate)=month('" & dtpProcess & "')) and (year(transdate)=year('" & dtpProcess & "')))" _
  & "  and Trans_Code not in(select code from d_TransportersPayRoll where yyear=year('" & dtpProcess & "') and mmonth =month('" & dtpProcess & "')) order by Trans_Code asc")
While Not rst.EOF
DoEvents
frmProcess.Caption = rst.Fields("Trans_Code")

    oSaccoMaster.ExecuteThis ("insert into d_TransportersPayRoll (code,EndPeriod) values('" & Trim$(rst.Fields("Trans_Code")) & "','" & dtpProcess & "')")
rst.MoveNext
Wend
End Sub
Public Sub addomittedentried()
Dim year As Integer
Dim month As Integer

Set rst = oSaccoMaster.GetRecordset("set dateformat dmy select distinct sno from d_milkintake where (month(transdate)=month('" & dtpProcess & "')) and (year(transdate)=year('" & dtpProcess & "')) and (sno not in(select sno from d_payroll where yyear=year('" & dtpProcess & "') and mmonth =month('" & dtpProcess & "'))) order by sno asc")
While Not rst.EOF
DoEvents
frmProcess.Caption = rst.Fields("sno")

    oSaccoMaster.ExecuteThis ("insert into d_payroll (sno,endofperiod) values('" & rst.Fields("sno") & "','" & dtpProcess & "')")
rst.MoveNext
Wend
End Sub

Private Sub Txtcreditedac_Change()
Set rst = oSaccoMaster.GetRecordset("select glaccname from glsetup where accno='" & Txtcreditedac & "'")
If Not rst.EOF Then
    lblcreditedac = rst.Fields(0)
End If
End Sub

Private Sub Txtdebitedac_Change()
Set rst = oSaccoMaster.GetRecordset("select glaccname from glsetup where accno='" & Txtdebitedac & "'")
If Not rst.EOF Then
    lbldebitedac = rst.Fields(0)
End If



End Sub


