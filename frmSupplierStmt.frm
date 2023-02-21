VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmSupplierStmt 
   Caption         =   "Print Suppliers'/Farmers' Statement"
   ClientHeight    =   3615
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8775
   Icon            =   "frmSupplierStmt.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3615
   ScaleWidth      =   8775
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdViewreport 
      Caption         =   "View Report"
      Height          =   375
      Left            =   6120
      TabIndex        =   16
      Top             =   3120
      Width           =   1455
   End
   Begin VB.CheckBox chkLPT 
      Caption         =   "Use LPT Port"
      Height          =   345
      Left            =   6120
      TabIndex        =   15
      Top             =   240
      Value           =   1  'Checked
      Width           =   1575
   End
   Begin VB.ComboBox Ports 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      ItemData        =   "frmSupplierStmt.frx":030A
      Left            =   6120
      List            =   "frmSupplierStmt.frx":031D
      TabIndex        =   14
      Text            =   "\\127.0.0.1\E-PoS 80mm Thermal Printer"
      Top             =   720
      Width           =   2535
   End
   Begin VB.CheckBox chkNotepad 
      Caption         =   "To Notepad"
      Height          =   255
      Left            =   4440
      TabIndex        =   13
      Top             =   600
      Width           =   1455
   End
   Begin VB.CommandButton cmdroute 
      Caption         =   "Routes"
      Height          =   375
      Left            =   3720
      TabIndex        =   12
      Top             =   3240
      Width           =   1335
   End
   Begin VB.CommandButton cmClose 
      Appearance      =   0  'Flat
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2640
      TabIndex        =   10
      Top             =   3240
      Width           =   735
   End
   Begin VB.CommandButton cmdPrint 
      Appearance      =   0  'Flat
      Caption         =   "Print"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1920
      TabIndex        =   9
      Top             =   3240
      Width           =   735
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Supplier Statements"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   1815
      Left            =   120
      TabIndex        =   5
      Top             =   1320
      Width           =   5775
      Begin VB.OptionButton optAdvanceSlip 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Caption         =   "Print Advance Slip"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   480
         TabIndex        =   11
         Top             =   1440
         Width           =   4095
      End
      Begin VB.OptionButton OptNormalA4 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Caption         =   "Normal Statement (Use Normal Printer (A4))"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   480
         TabIndex        =   8
         Top             =   1080
         Width           =   4095
      End
      Begin VB.OptionButton OptDetailedStmt 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Caption         =   "Detailed Statement (Use POS Printer)"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   480
         TabIndex        =   7
         Top             =   720
         Width           =   4095
      End
      Begin VB.OptionButton OptNormalStmt 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Caption         =   "Normal Statement (Use POS Printer)"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   480
         TabIndex        =   6
         Top             =   360
         Value           =   -1  'True
         Width           =   4095
      End
   End
   Begin VB.TextBox txtSNo 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1920
      TabIndex        =   3
      Top             =   480
      Width           =   735
   End
   Begin MSComCtl2.DTPicker DTPStmts 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "dd/MM/yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   3
      EndProperty
      Height          =   255
      Left            =   1920
      TabIndex        =   2
      Top             =   840
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   450
      _Version        =   393216
      Format          =   51118081
      CurrentDate     =   40109
   End
   Begin MSComDlg.CommonDialog cdgPrint 
      Left            =   5400
      Top             =   3120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      FileName        =   "c:\receipt.txt"
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0E0FF&
      Caption         =   "Enter supplier number and select end of period to print statement"
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
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   5835
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0E0FF&
      Caption         =   "End of Period :"
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
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   1365
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0E0FF&
      Caption         =   "Supplier Number :"
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
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   1665
   End
End
Attribute VB_Name = "frmSupplierStmt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Enddate As Date
Dim Startdate As Date
Dim TRANSPORTER As String

Private Sub chkLPT_Click()
ports.Clear
ports = ""
'//If the drivers are installed it won't matter whether the Port is indicated
' or not it will just work.

If chkLPT.value = vbChecked Then
ports.AddItem ports
ports = ports
ports.AddItem "LPT2"
ports.AddItem "LPT3"
ports.AddItem "LPT4"
ports.AddItem "LPT5"
Else
'Share the printer first the use of 127.0.0.1 which is
'standard IP address for a loopback network connection
'instead of getting the computer name or IP Address
'
Dim prnPrinter As Printer
Dim pr As String
ports.Clear

For Each prnPrinter In Printers
   If InStr(prnPrinter.DeviceName, "\\") Then
    ports.AddItem prnPrinter.DeviceName
    If InStr(prnPrinter.DeviceName, "G") Then
    ports.Text = prnPrinter.DeviceName
    End If
    Else
    ports.AddItem "\\127.0.0.1\" & prnPrinter.DeviceName
    If InStr(prnPrinter.DeviceName, "G") Then
    ports.Text = "\\127.0.0.1\" & prnPrinter.DeviceName
    End If
    End If
   
   
Next
End If
'This code will work only if there is a connection e.g LAN or modem.
'It is not a must that it is an internet connection because
'computer's network interface card has to be functional
End Sub

Private Sub cmClose_Click()
Unload Me
End Sub

Private Sub cmdPrint_Click()
On Error GoTo errorhandler22
'Dim fso, chkPrinter, txtFile, GPay As Currency, TotDeduction As Double
'GPay = 0
Dim fso, chkPrinter, txtFile, GPay As Currency, TotDeduction As Double, rss As New Recordset, rsts As New Recordset, shareamt As Double, amtt As Double
GPay = 0
    Dim ttt
     Dim escFeedAndCut As String
     Dim escNewLine As String
     Dim escUnerLineON As String
     Dim escUnerLineOnX2 As String
     Dim escUnerLineOff As String
     Dim escBoldOn As String
     Dim escBoldOff As String
     Dim escNegativeOn As String
     Dim escNegativeOff As String
     Dim esc8CpiOn As String
     Dim esc8CPiOff As String
     Dim esc16Cpi As String
     Dim esc20Cpi As String
     Dim escAlignLeft As String
     Dim escAlignCenter As String
     Dim escAlignRight As String
    Dim CummulKgs As Double
     
        escNewLine = Chr(10) '//New Line (LF Line Feed)
        escUnerLineON = Chr(27) + Chr(45) + Chr(1) '//Unerline on
        escUnerLineOnX2 = Chr(27) + Chr(45) + Chr(1) '//Unerline on X2
        escUnerLineOff = Chr(27) + Chr(45) + Chr(0) '//unerline off
        escBoldOn = Chr(27) + Chr(69) + Chr(1) '//Bold on
        escBoldOff = Chr(27) + Chr(69) + Chr(0) '//Bold off
        escNegativeOn = Chr(29) + Chr(66) + Chr(1) '//White on Black on
        escNegativeOff = Chr(29) + Chr(66) + Chr(0) '//white on
        esc8CpiOn = Chr(29) + Chr(33) + Chr(16) '//Font Size X2 on
        esc8CPiOff = Chr(29) + Chr(33) + Chr(0) '//Font size X2 off
        esc16Cpi = Chr(27) + Chr(77) + Chr(48) '//Font A - Normal Size
        esc20Cpi = Chr(27) + Chr(77) + Chr(49) '//Font B - Small Font
        escAlignLeft = Chr(27) + Chr(97) + Chr(48) '//Align text to the left
        escAlignCenter = Chr(27) + Chr(97) + Chr(49) '//Align text to the center
        escAlignRight = Chr(27) + Chr(97) + Chr(50) '//Align text to the right
        escFeedAndCut = Chr(29) + Chr(86) + Chr(65) '//Partial cut and feed
       
If txtSNo = "" Then
    MsgBox "Please enter supplier number.", vbCritical
        txtSNo.SetFocus
    Exit Sub
End If

If Not IsNumeric(txtSNo) Then
    MsgBox "Please enter number. '" & UCase(txtSNo) & "' is not a number", vbCritical
        txtSNo.SetFocus
    Exit Sub
End If

Startdate = DateSerial(year(DTPStmts), month(DTPStmts), 1)
Enddate = DateSerial(year(DTPStmts), month(DTPStmts) + 1, 1 - 1)
'oSaccoMaster.ExecuteThis ("d_sp_TransUpdate '" & Startdate & "','" & Enddate & "','" & User & "','" & Trim(rst.Fields(0)) & "'")
DTPStmts = Enddate

If optAdvanceSlip.value = True Then
'--Net amount as at date
'    Startdate = DateSerial(Year(txtransdate), month(txtransdate), 1)
'Enddate = DateSerial(Year(txtransdate), month(txtransdate) + 1, 1 - 1)

Set rs = oSaccoMaster.GetRecordset("d_sp_SupNet " & txtSNo & ",'" & Startdate & "','" & Enddate & "', 0")

Dim Kgs As Double
If Not IsNull(rs.Fields(0)) Then
Kgs = rs.Fields(0)
Else
Kgs = "0.00"
End If

Dim Gross As Double

If Not IsNull(rs.Fields(1)) Then
Gross = rs.Fields(1)
Else
Gross = "0.00"
End If
Dim Kainet As String
If Not IsNull(rs.Fields(2)) Then
Kainet = rs.Fields(2)
Else
Kainet = "XXXXX XXXX"
End If

Dim Ded As Double

Set rs = oSaccoMaster.GetRecordset("d_sp_SupNet " & txtSNo & ",'" & Startdate & "','" & Enddate & "', 1")
If Not rs.EOF Then
If Not IsNull(rs.Fields(0)) Then
Ded = rs.Fields(0)
Else
Ded = "0.00"
End If
End If
 Dim Net As Double
Net = Format((CCur(Gross) - CCur(Ded)), "#,##0.00")
 Dim PORT As String
        PORT = ports
        'ttt = "LPT1" 'LPT1
        ttt = PORT
        'ttt = "LPT1" 'LPT1,LPT2....
        Set fso = CreateObject("Scripting.FileSystemObject")
        On Error GoTo err
        
        'Set chkPrinter = fso.GetFile(ttt)
       
        
'    Set txtFile = fso.CreateTextFile(ttt, True)
'    txtFile.WriteLine escAlignCenter
'    txtFile.WriteLine "Advance Slip"
'    txtFile.WriteLine "" & cname & ""
'    txtFile.WriteLine "........................................"
'    txtFile.WriteLine escAlignLeft
'    txtFile.WriteLine "SNo. : " & txtSNo
'    txtFile.WriteLine "Names : " & Kainet
'    txtFile.WriteLine "Issue Items/Services worth not more than"
'    txtFile.WriteLine "Kshs. : " & Format(Net, "#,##0.00") & ""
'    txtFile.WriteLine "Sign"
'    txtFile.WriteLine "___________________________"
'    txtFile.WriteLine UCase(username)
'    txtFile.Write "Date " & Format(Get_Server_Date, "dd/mm/yyyy")
'    txtFile.WriteLine ", Time : " & Time
'    txtFile.WriteLine "........................................"
'    txtFile.WriteLine escFeedAndCut
    
    
    
End If
    
'----d_sp_PrintStmt @SNo bigint,@EndPeriod varchar(10)

If OptNormalStmt.value = True Then
    Set rst = New ADODB.Recordset
        sql = "d_sp_PrintStmt " & txtSNo & ",'" & Enddate & "'"
    Set rst = oSaccoMaster.GetRecordset(sql)
If rst.EOF Then
    MsgBox "There is no record in the payroll for supplier number " & txtSNo, vbInformation
        txtSNo.SetFocus
    Exit Sub
End If


 'Dim PORT As String
        PORT = ports
        'ttt = "LPT1" 'LPT1
        ttt = PORT
        'ttt = "LPT1" 'LPT1,LPT2....
        Set fso = CreateObject("Scripting.FileSystemObject")
        On Error GoTo err
        
'        Set chkPrinter = fso.GetFile(ttt)
       
        
    Set txtFile = fso.CreateTextFile(ttt, True)
    txtFile.WriteLine escAlignCenter
    txtFile.WriteLine "" & cname & ""
    txtFile.WriteLine "" & paddress & ""
    txtFile.WriteLine "" & town & ""
    txtFile.WriteLine "MILK STATEMENT FOR " & UCase(Format(DTPStmts, "MMMM/YYYY"))
    txtFile.WriteLine escAlignLeft
    txtFile.WriteLine "........................................"
    txtFile.WriteLine "SNo :" & txtSNo
    txtFile.WriteLine "Name :" & rst!NAMES
    txtFile.WriteLine "........................................"
    'startdate = DateSerial(Year(DTPStmts), month(DTPStmts) - 1, 1)
    Set rs = New ADODB.Recordset
    sql = "d_sp_TotalMonth " & txtSNo & ",'" & Startdate & "','" & Enddate & "'"
    Set rs = oSaccoMaster.GetRecordset(sql)
    If Not rs.EOF Then
    If Not IsNull(rs.Fields(0)) Then CummulKgs = rs.Fields(0)
    Else
    CummulKgs = "0.00"
    End If
    
    txtFile.WriteLine "Total Kgs :" & Format(CummulKgs, "#,##0.00" & " Kgs")
    
    txtFile.WriteLine "Gross Amount               Kshs: " & Format(rst!GPay, "#,##0.00") & ""
    GPay = Format(rst!GPay, "#,##0.00")
    txtFile.Write escBoldOn
    txtFile.WriteLine "DEDUCTIONS"
    txtFile.WriteLine "-------------"
    txtFile.Write escBoldOff
    Set rst = New ADODB.Recordset
    sql = "d_sp_PrintDeductStmt " & txtSNo & ",'" & Startdate & "','" & Enddate & "'"
    Set rst = oSaccoMaster.GetRecordset(sql)
    
    txtFile.WriteLine "........................................"
    txtFile.WriteLine "DATE " & vbTab & "" & vbTab & "AMOUNT" & vbTab & "DESCRIPTION"
    txtFile.WriteLine "........................................"
   ' Dim TotDeduction As Double
    TotDeduction = 0
    While Not rst.EOF
        'MsgBox rs!QSupplied
        txtFile.WriteLine rst!date_deduc & " " & vbTab & " " & Format(rst!amount, "#,##0.00" & vbTab & " " & rst!description & " " & rst!Remarks & " ")
        TotDeduction = TotDeduction + rst!amount
        'txtfile.WriteLine rs!PPU
         rst.MoveNext
        
        Wend
    Set rst1 = New ADODB.Recordset
        sql = "d_sp_PrintStmt " & txtSNo & ",'" & Enddate & "'"
    Set rst1 = oSaccoMaster.GetRecordset(sql)
    If Not IsNull(rst1!Transport) Then
              txtFile.WriteLine Enddate & " " & vbTab & " " & Format(rst1!Transport, "#,##0.00" & vbTab & " " & "Transport ")
              TotDeduction = TotDeduction + rst1!Transport
    End If
     'shares'
            Set rsts = oSaccoMaster.GetRecordset("SELECT    SUM(Amount) AS amtt From d_sconribution WHERE     (transdescription LIKE '%shares%') AND (SNo = '" & txtSNo & "')AND (transdate <=  '" & DTPStmts & "')")
            If Not rsts.EOF Then
            shareamt = IIf(IsNull(rsts!amtt), 0, rsts!amtt)
            End If
            'Set rss = oSaccoMaster.GetRecordset("SELECT    SUM(Amount) AS amt From d_supplier_deduc WHERE     (Description LIKE '%shares%') AND (SNo = '" & txtsno & "')")
            'If Not rss.EOF Then
            'TXTshares = IIf(IsNull(rss!amt), 0, rss!amt) + shareamt
            'End If
            Set rss = oSaccoMaster.GetRecordset("SELECT    SUM(Amount) AS amt From d_supplier_deduc WHERE     (Description LIKE '%shares%') AND (SNo = '" & txtSNo & "')AND (Date_Deduc <=  '" & DTPStmts & "')")
                If Not rss.EOF Then
            Shares = IIf(IsNull(rss!amt), 0, rss!amt) + shareamt
            End If
            
            'end shares'
    txtFile.WriteLine "Total Shares: " & Format(IIf(IsNull(Shares), 0, Shares), "#,##0.00") & ""
    txtFile.WriteLine "........................................"
    txtFile.WriteLine "Total Deductions Kshs: " & Format(TotDeduction, "#,##0.00") & ""
    txtFile.WriteLine "........................................"
    txtFile.WriteLine "NET PAY                   Kshs :" & Format((GPay - TotDeduction), "#,##0.00") & ""
    txtFile.WriteLine "........................................"
    txtFile.WriteLine "BANK DETAILS"
    txtFile.WriteLine "-------------"
    txtFile.WriteLine "Bank Name :" & rst!bank & ""
    txtFile.WriteLine "Bank Branch :" & rst!BBranch
    txtFile.WriteLine "Account Number :" & rst!accountnumber
'    txtfile.WriteLine "........................................"

'    sql = "d_sp_TransName '" & txtSNo & "'"
'    Set rs = oSaccoMaster.GetRecordset(sql)
'    If Not rs.EOF Then
'    If Not IsNull(rs.Fields(0)) Then TRANSPORTER = rs.Fields(0)
'    Else
'
'    End If
    txtFile.WriteLine "---------------------------------------"
    txtFile.WriteLine "        Date :" & Format(Get_Server_Date, "dd/mm/yyyy HH:MM:SS AM/PM")
    txtFile.WriteLine "         " & motto & ""
    txtFile.WriteLine "---------------------------------------"
    txtFile.WriteLine escFeedAndCut
    txtFile.Close
'End If
Exit Sub
err: MsgBox err.description & " or There is no printer connected."
End If
'/print detail statement in the notepad
If chkNotepad = vbChecked Then
                       
                            
                           
                        escFeedAndCut = Chr(29) + Chr(86) + Chr(65) '//Partial cut and feed
                        cdgPrint.Filter = "*.csv|*.txt"
                        cdgPrint.ShowSave
                        ttt = cdgPrint.FileName
                        If ttt = "" Then
                        MsgBox "File should not be blank", vbCritical, "Data transfer"
                        Exit Sub
                        End If
                        'ttt = "D:\PROJECTS\FOSA\DAILY" & Date & ""
                        Set fso = CreateObject("Scripting.FileSystemObject")
                        Set txtFile = fso.CreateTextFile(ttt, True)
                        txtFile.WriteLine
                       'PORT = ttt
                      
                       'ttt = PORT
                
                       'Set fso = CreateObject("Scripting.FileSystemObject")
                       On Error GoTo err
                            
                            
                           
                            
                        'Set txtfile = fso.CreateTextFile(ttt, True)
                        txtFile.WriteLine escAlignCenter
                        txtFile.WriteLine "" & cname & ""
                        txtFile.WriteLine "" & paddress & ""
                        txtFile.WriteLine "" & town & ""
                        txtFile.WriteLine "DETAILED STATEMENT FOR " & UCase(Format(DTPStmts, "MMMM/YYYY"))
                        txtFile.WriteLine escAlignLeft
                        '//PUT HERE THE TRANSPORTER
                        Dim rtg As New ADODB.Recordset, sno3 As String
                        Set rtg = oSaccoMaster.GetRecordset("SELECT     TOP 1 Trans_Code, Sno   FROM         d_Transport WHERE     (Sno = " & txtSNo & ")  ORDER BY auditdatetime DESC")
                        If Not rtg.EOF Then
                        sno3 = IIf(IsNull(Trim(rtg.Fields(0))), "Self", Trim(rtg.Fields(0)))
                        Else
                        sno3 = "Self"
                        End If
                        txtFile.WriteLine "Transporter :" & sno3
                        txtFile.WriteLine "........................................"
                        txtFile.WriteLine "SNo :" & txtSNo
                        
                        Set rs = New ADODB.Recordset
                        sql = "d_sp_PrintDedStmt " & txtSNo & ",'" & Startdate & "','" & Enddate & "'"
                        Set rs = oSaccoMaster.GetRecordset(sql)
                        If rs.EOF Then
                        MsgBox "The supplier did not supplier for the month specified", vbInformation
                        
                        txtFile.WriteLine "---------------------------------------"
                        txtFile.WriteLine escFeedAndCut
                        txtFile.Close
                        Exit Sub
                        End If
                        
                        txtFile.WriteLine "Name :" & rs!NAMES
                        txtFile.WriteLine "........................................"
                        txtFile.WriteLine "DATE " & vbTab & "" & vbTab & "QNTY" & vbTab & "PRICE" & vbTab & "PAYABLE"
                        txtFile.WriteLine "........................................"
                        sql = ""
                        sql = "SELECT SUM(d_Shares.Amnt) AS TotalShares FROM d_Shares where d_Shares.Code = CONVERT(varchar(35)," & txtSNo & ")"
                        Set rs2 = oSaccoMaster.GetRecordset(sql)
                        Dim qnty As Currency
                             qnty = 0
                             GPay = 0
                             
                             
                            While Not rs.EOF
                            Dim Pric As Currency
                            Pric = rs!ppu
                            If Not IsNull(rs2.Fields(0)) Then
                            If rs2.Fields(0) > 99999999999999# Then
                              Pric = (rs!ppu) + 1

                            End If
                            End If
                            
                            'MsgBox rs!QSupplied
                            
                            txtFile.WriteLine rs!transdate & " " & vbTab & " " & Format(rs!QSupplied, "#,##0.0#") & " " & vbTab & " " & Format(Pric, "#,##0.00") & " " & vbTab & " " & Format(((Pric) * rs!QSupplied), "#,##0.00")
                            'txtfile.WriteLine rs!PPU
                            qnty = qnty + rs!QSupplied
                            GPay = GPay + (Pric * rs!QSupplied)
                             rs.MoveNext
                            
                            Wend
                    Set rs2 = New ADODB.Recordset
                    'Dim Startdate As String, Enddate As String
                    
                    'Startdate = DateSerial(DTPMilkDate, cboMonth, 1)
                    sql = "d_sp_UpdateGPAYQnty '" & Startdate & "','" & Enddate & "'," & txtSNo & ""
                    Set rs2 = oSaccoMaster.GetRecordset(sql)
                    If Not rs2.EOF Then
                    If Not IsNull(rs2.Fields(0)) Then qnty = rs2.Fields(0)
                    '-If Not IsNull(rs2.Fields(1)) Then GPay = rs2.Fields(1)
                    End If
                    
                        Dim subsidy As Double
                        
                        txtFile.WriteLine "........................................"
                        txtFile.WriteLine "Total Kgs :" & Format(qnty, "#,##0.00" & " Kgs")
                        txtFile.WriteLine "Gross Pay Kshs :" & Format(GPay, "#,##0.00" & "")
                        txtFile.WriteLine "........................................"
'                        Set rs = oSaccoMaster.GetRecordset(" set dateformat dmy SELECT     subsidy   FROM         d_Payroll  WHERE     sno = " & txtSNo & " AND endofperiod='" & DTPStmts & "'")
'                                        If Not rs.EOF Then
'                                            subsidy = IIf(IsNull(rs.Fields(0)), 0, rs.Fields(0))
'                                        End If
Set rs = oSaccoMaster.GetRecordset(" set dateformat dmy SELECT     REWARD   FROM         Bonus  WHERE     sno = " & txtSNo & " and transdate='" & DTPStmts & "'")
                                        If Not rs.EOF Then
                                            subsidy = IIf(IsNull(rs.Fields(0)), 0, rs.Fields(0))
                                        End If
                                        
                        'txtFile.WriteLine "Other Income:" & Format(subsidy, "#,##0.00" & " Kshs")
                        txtFile.WriteLine "Gross Pay Kshs :" & Format(GPay + subsidy, "#,##0.00" & "")
                        txtFile.WriteLine "........................................"
                        txtFile.WriteLine escBoldOn
                        txtFile.WriteLine "DEDUCTIONS"
                        txtFile.WriteLine "........................................"
                        txtFile.WriteLine escBoldOff
                        GPay = GPay + subsidy
                    Set rst = New ADODB.Recordset
                    sql = "d_sp_PrintDeductStmt " & txtSNo & ",'" & Startdate & "','" & Enddate & "'"
                    Set rst = oSaccoMaster.GetRecordset(sql)
                        
                        txtFile.WriteLine "........................................"
                        txtFile.WriteLine "DATE " & vbTab & "" & vbTab & "AMOUNT" & vbTab & "DESCRIPTION"
                        txtFile.WriteLine "........................................"
                        
                        TotDeduction = 0
                        While Not rst.EOF
                            'MsgBox rs!QSupplied
                            txtFile.WriteLine rst!date_deduc & " " & vbTab & " " & Format(rst!amount, "#,##0.00" & vbTab & " " & rst!description & " " & rst!Remarks & " ")
                            TotDeduction = TotDeduction + rst!amount
                            'txtfile.WriteLine rs!PPU
                             rst.MoveNext
                            
                            Wend
                        Set rst1 = New ADODB.Recordset
                            sql = "d_sp_PrintStmt " & txtSNo & ",'" & Enddate & "'"
                        Set rst1 = oSaccoMaster.GetRecordset(sql)
                        If Not IsNull(rst1!Transport) Then
                                  txtFile.WriteLine Enddate & " " & vbTab & " " & Format(rst1!Transport, "#,##0.00" & vbTab & " " & "Transport ")
                                  TotDeduction = TotDeduction + rst1!Transport
                        End If
                       ' txtFile.WriteLine "Total Shares: " & Format(IIf(IsNull(Shares), 0, Shares), "#,##0.00") & ""
                        'txtFile.WriteLine "........................................"
                        txtFile.WriteLine "........................................"
                        txtFile.WriteLine "Total Deductions Kshs: " & Format(TotDeduction, "#,##0.00") & ""
                        txtFile.WriteLine "........................................"
                        txtFile.WriteLine "NET PAY                   Kshs :" & Format((GPay - TotDeduction), "#,##0.00") & ""
                        txtFile.WriteLine "-----------------------------------------"
                        txtFile.WriteLine "BANK DETAILS"
                        txtFile.WriteLine "-------------"
                        txtFile.WriteLine "Bank Name :" & rst1!bank & ""
                        txtFile.WriteLine "Bank Branch :" & rst1!BBranch
                        txtFile.WriteLine "Account Number :" & rst1!accountnumber
                        txtFile.WriteLine "---------------------------------------"
                        txtFile.WriteLine "        Date :" & Format(Get_Server_Date, "dd/mm/yyyy HH:MM:SS AM/PM")
                        txtFile.WriteLine "         " & motto & ""
                        txtFile.WriteLine "---------------------------------------"
                        txtFile.WriteLine "DEVELOP BY: AMTECH TECHNOLOGIES LIMITED"
                        txtFile.WriteLine "---------------------------------------"
                        txtFile.WriteLine escFeedAndCut
                        txtFile.Close
End If
'--Print detailed statement
If OptDetailedStmt.value = True And chkNotepad = vbUnchecked Then
                     'Dim PORT As String
                            PORT = ports
                            'ttt = "LPT1" 'LPT1
                            ttt = PORT
                      'ttt = "LPT1" 'LPT1,LPT2....
                            Set fso = CreateObject("Scripting.FileSystemObject")
                            On Error GoTo err
                            
                            'Set chkPrinter = fso.GetFile(ttt)
                           
                            
                        Set txtFile = fso.CreateTextFile(ttt, True)
                        txtFile.WriteLine escAlignCenter
                        txtFile.WriteLine "" & cname & ""
                        txtFile.WriteLine "" & paddress & ""
                        txtFile.WriteLine "" & town & ""
                        txtFile.WriteLine "DETAILED STATEMENT FOR " & UCase(Format(DTPStmts, "MMMM/YYYY"))
                        txtFile.WriteLine escAlignLeft
                        '//PUT HERE THE TRANSPORTER
                        'Dim rtg As New ADODB.Recordset, sno3 As String
                        Set rtg = oSaccoMaster.GetRecordset("SELECT     TOP 1 Trans_Code, Sno   FROM         d_Transport WHERE     (Sno = " & txtSNo & ")  ORDER BY auditdatetime DESC")
                        If Not rtg.EOF Then
                        sno3 = IIf(IsNull(Trim(rtg.Fields(0))), "Self", Trim(rtg.Fields(0)))
                        Else
                        sno3 = "Self"
                        End If
                        txtFile.WriteLine "Transporter :" & sno3
                        txtFile.WriteLine "........................................"
                        txtFile.WriteLine "SNo :" & txtSNo
                        
                        Set rs = New ADODB.Recordset
                        sql = "d_sp_PrintDedStmt " & txtSNo & ",'" & Startdate & "','" & Enddate & "'"
                        Set rs = oSaccoMaster.GetRecordset(sql)
                        If rs.EOF Then
                        MsgBox "The supplier did not supplier for the month specified", vbInformation
                        
                        txtFile.WriteLine "---------------------------------------"
                        txtFile.WriteLine escFeedAndCut
                        txtFile.Close
                        Exit Sub
                        End If
                        
                        txtFile.WriteLine "Name :" & rs!NAMES
                        txtFile.WriteLine "........................................"
                        txtFile.WriteLine "DATE " & vbTab & "" & vbTab & "QNTY" & vbTab & "PRICE" & vbTab & "PAYABLE"
                        txtFile.WriteLine "........................................"
                        sql = ""
'                        sql = "SELECT SUM(d_Shares.Amnt) AS TotalShares FROM d_Shares where d_Shares.Code = CONVERT(varchar(35)," & txtSNo & ")"
'                        Set rs2 = oSaccoMaster.GetRecordset(sql)
                        sql = "select sum(spu) as shares from d_shares where sno='" & txtSNo & "'"
                        Set rs2 = oSaccoMaster.GetRecordset(sql)
                        'Dim qnty As Currency, GPay As Currency
                             qnty = 0
                             GPay = 0
                             
                             
                            While Not rs.EOF
                           ' Dim Pric As Currency
                            Pric = rs!ppu
                            If Not IsNull(rs2.Fields(0)) Then
                            If rs2.Fields(0) > 999999999999# Then
                              Pric = (rs!ppu) + 1

                            End If
                            End If
                            
                            'MsgBox rs!QSupplied
                            
                            txtFile.WriteLine rs!transdate & " " & vbTab & " " & Format(rs!QSupplied, "#,##0.0#") & " " & vbTab & " " & Format(Pric, "#,##0.00") & " " & vbTab & " " & Format(((Pric) * rs!QSupplied), "#,##0.00")
                            'txtfile.WriteLine rs!PPU
                            qnty = qnty + rs!QSupplied
                            GPay = GPay + (Pric * rs!QSupplied)
                             rs.MoveNext
                            
                            Wend
                    Set rs2 = New ADODB.Recordset
                    'Dim Startdate As String, Enddate As String
                    
                    'Startdate = DateSerial(DTPMilkDate, cboMonth, 1)
                    sql = "d_sp_UpdateGPAYQnty '" & Startdate & "','" & Enddate & "'," & txtSNo & ""
                    Set rs2 = oSaccoMaster.GetRecordset(sql)
                    If Not rs2.EOF Then
                    If Not IsNull(rs2.Fields(0)) Then qnty = rs2.Fields(0)
                    '-If Not IsNull(rs2.Fields(1)) Then GPay = rs2.Fields(1)
                    End If
                    
                        'Dim subsidy As Double
                        
                        txtFile.WriteLine "........................................"
                        txtFile.WriteLine "Total Kgs :" & Format(qnty, "#,##0.00" & " Kgs")
                        txtFile.WriteLine "Gross Pay Kshs :" & Format(GPay, "#,##0.00" & "")
                        txtFile.WriteLine "........................................"
'                        Set rs = oSaccoMaster.GetRecordset(" set dateformat dmy SELECT     subsidy   FROM         d_Payroll  WHERE     sno = " & txtSNo & " AND endofperiod='" & DTPStmts & "'")
'                                        If Not rs.EOF Then
'                                            subsidy = IIf(IsNull(rs.Fields(0)), 0, rs.Fields(0))
'                                        End If
                    Set rs = oSaccoMaster.GetRecordset(" set dateformat dmy SELECT     REWARD   FROM         Bonus  WHERE     sno = " & txtSNo & " and transdate='" & DTPStmts & "'")
                        If Not rs.EOF Then
                            subsidy = IIf(IsNull(rs.Fields(0)), 0, rs.Fields(0))
                        End If
                                        
                        txtFile.WriteLine "Other Income(Reward) :" & Format(subsidy, "#,##0.00" & " Kshs")
                        txtFile.WriteLine "Gross  Pay Kshs :" & Format(GPay + subsidy, "#,##0.00" & "")
                        txtFile.WriteLine "........................................"
                        txtFile.WriteLine escBoldOn
                        txtFile.WriteLine "DEDUCTIONS"
                        txtFile.WriteLine "........................................"
                        txtFile.WriteLine escBoldOff
                        GPay = GPay + subsidy
                    Set rst = New ADODB.Recordset
                    sql = "d_sp_PrintDeductStmt " & txtSNo & ",'" & Startdate & "','" & Enddate & "'"
                    Set rst = oSaccoMaster.GetRecordset(sql)
                        
                        txtFile.WriteLine "........................................"
                        txtFile.WriteLine "DATE " & vbTab & "" & vbTab & "AMOUNT" & vbTab & "DESCRIPTION"
                        txtFile.WriteLine "........................................"
                       ' Dim TotDeduction As Double
                        TotDeduction = 0
                        While Not rst.EOF
                            'MsgBox rs!QSupplied
                            txtFile.WriteLine rst!date_deduc & " " & vbTab & " " & Format(rst!amount, "#,##0.00" & vbTab & " " & rst!description & " " & rst!Remarks & " ")
                            TotDeduction = TotDeduction + rst!amount
                            'txtfile.WriteLine rs!PPU
                             rst.MoveNext
                            
                            Wend
                        Set rst1 = New ADODB.Recordset
                            sql = "d_sp_PrintStmt " & txtSNo & ",'" & Enddate & "'"
                        Set rst1 = oSaccoMaster.GetRecordset(sql)
                        If Not IsNull(rst1!Transport) Then
                                  txtFile.WriteLine Enddate & " " & vbTab & " " & Format(rst1!Transport, "#,##0.00" & vbTab & " " & "Transport ")
                                  TotDeduction = TotDeduction + rst1!Transport
                        End If
                         'shares'
                        Set rsts = oSaccoMaster.GetRecordset("SELECT    SUM(Amount) AS amtt From d_sconribution WHERE     (transdescription LIKE '%shares%') AND (SNo = '" & txtSNo & "')AND (transdate <=  '" & DTPStmts & "')")
                        If Not rsts.EOF Then
                        shareamt = IIf(IsNull(rsts!amtt), 0, rsts!amtt)
                        End If
                        'Set rss = oSaccoMaster.GetRecordset("SELECT    SUM(Amount) AS amt From d_supplier_deduc WHERE     (Description LIKE '%shares%') AND (SNo = '" & txtsno & "')")
                        'If Not rss.EOF Then
                        'TXTshares = IIf(IsNull(rss!amt), 0, rss!amt) + shareamt
                        'End If
                        Set rss = oSaccoMaster.GetRecordset("SELECT    SUM(Amount) AS amt From d_supplier_deduc WHERE     (Description LIKE '%shares%') AND (SNo = '" & txtSNo & "')AND (Date_Deduc <=  '" & DTPStmts & "')")
                            If Not rss.EOF Then
                        Shares = IIf(IsNull(rss!amt), 0, rss!amt) + shareamt
                        End If
                        
                        'end shares'
                        txtFile.WriteLine "Total Shares: " & Format(IIf(IsNull(Shares), 0, Shares), "#,##0.00") & ""
                        'txtFile.WriteLine "........................................"
                        txtFile.WriteLine "........................................"
                        txtFile.WriteLine "Total Deductions Kshs: " & Format(TotDeduction, "#,##0.00") & ""
                        txtFile.WriteLine "........................................"
                        txtFile.WriteLine "NET PAY                   Kshs :" & Format((GPay - TotDeduction), "#,##0.00") & ""
                        txtFile.WriteLine "-----------------------------------------"
                        txtFile.WriteLine "BANK DETAILS"
                        txtFile.WriteLine "-------------"
                        txtFile.WriteLine "Bank Name :" & rst1!bank & ""
                        txtFile.WriteLine "Bank Branch :" & rst1!BBranch
                        txtFile.WriteLine "Account Number :" & rst1!accountnumber
                        txtFile.WriteLine "---------------------------------------"
                        txtFile.WriteLine "        Date :" & Format(Get_Server_Date, "dd/mm/yyyy HH:MM:SS AM/PM")
                        txtFile.WriteLine "         " & motto & ""
                        txtFile.WriteLine "---------------------------------------"
                        txtFile.WriteLine "DEVELOP BY: AMTECH TECHNOLOGIES LIMITED"
                        txtFile.WriteLine "---------------------------------------"
                        txtFile.WriteLine escFeedAndCut
                        txtFile.Close
    End If
            
    If OptNormalA4.value = True Then
    reportname = "d_StmtA4.rpt"
    '{d_Payroll.NPay} > 0 and {d_Payroll.Bank} <> '' and month({d_Payroll.EndofPeriod})= month(30/09/2010)  AND year({d_Payroll.EndofPeriod}) = Year(30/09/2010)
    STRFORMULA = "{d_Payroll.SNo}= " & txtSNo & " and month({d_Payroll.EndofPeriod})=" & month(DTPStmts) & " AND year({d_Payroll.EndofPeriod}) =" & year(DTPStmts)
    Show_Sales_Crystal_Report STRFORMULA, reportname, ""

'    d_StmtA4
    End If
    txtSNo = ""
    Exit Sub
errorhandler22:
    MsgBox err.description
End Sub

Private Sub cmdroute_Click()
On Error GoTo errorhandler22
Dim fso, chkPrinter, txtFile
    Dim ttt
     Dim escFeedAndCut As String
     Dim escNewLine As String
     Dim escUnerLineON As String
     Dim escUnerLineOnX2 As String
     Dim escUnerLineOff As String
     Dim escBoldOn As String
     Dim escBoldOff As String
     Dim escNegativeOn As String
     Dim escNegativeOff As String
     Dim esc8CpiOn As String
     Dim esc8CPiOff As String
     Dim esc16Cpi As String
     Dim esc20Cpi As String
     Dim escAlignLeft As String
     Dim escAlignCenter As String
     Dim escAlignRight As String
    Dim CummulKgs As Double
     
        escNewLine = Chr(10) '//New Line (LF Line Feed)
        escUnerLineON = Chr(27) + Chr(45) + Chr(1) '//Unerline on
        escUnerLineOnX2 = Chr(27) + Chr(45) + Chr(1) '//Unerline on X2
        escUnerLineOff = Chr(27) + Chr(45) + Chr(0) '//unerline off
        escBoldOn = Chr(27) + Chr(69) + Chr(1) '//Bold on
        escBoldOff = Chr(27) + Chr(69) + Chr(0) '//Bold off
        escNegativeOn = Chr(29) + Chr(66) + Chr(1) '//White on Black on
        escNegativeOff = Chr(29) + Chr(66) + Chr(0) '//white on
        esc8CpiOn = Chr(29) + Chr(33) + Chr(16) '//Font Size X2 on
        esc8CPiOff = Chr(29) + Chr(33) + Chr(0) '//Font size X2 off
        esc16Cpi = Chr(27) + Chr(77) + Chr(48) '//Font A - Normal Size
        esc20Cpi = Chr(27) + Chr(77) + Chr(49) '//Font B - Small Font
        escAlignLeft = Chr(27) + Chr(97) + Chr(48) '//Align text to the left
        escAlignCenter = Chr(27) + Chr(97) + Chr(49) '//Align text to the center
        escAlignRight = Chr(27) + Chr(97) + Chr(50) '//Align text to the right
        escFeedAndCut = Chr(29) + Chr(86) + Chr(65) '//Partial cut and feed
       




Startdate = DateSerial(year(DTPStmts), month(DTPStmts), 1)
Enddate = DateSerial(year(DTPStmts), month(DTPStmts) + 1, 1 - 1)
DTPStmts = Enddate
'********************************************to notepad
If chkNotepad.value = vbChecked Then

  
'     Dim escFeedAndCut As String
     escFeedAndCut = Chr(29) + Chr(86) + Chr(65) '//Partial cut and feed
       cdgPrint.Filter = "*.csv|*.txt"
        cdgPrint.ShowSave
        ttt = cdgPrint.FileName
        If ttt = "" Then
        MsgBox "File should not be blank", vbCritical, "Data transfer"
        Exit Sub
        End If
        'ttt = "D:\PROJECTS\FOSA\DAILY" & Date & ""
        Set fso = CreateObject("Scripting.FileSystemObject")
        Set txtFile = fso.CreateTextFile(ttt, True)
        txtFile.WriteLine
        
    txtFile.WriteLine "---------------------------------------"
    txtFile.WriteLine "" & cname & ""
   ' Printer.Print Tab(0); "Kimathi House Branch"
    txtFile.WriteLine " " & paddress & " "
    txtFile.WriteLine "" & town & ""
    txtFile.WriteLine "Milk Receipt"
    txtFile.WriteLine "---------------------------------------"
'    If cbomemtrans = "Shares" Then
'    DESC = bosanames & " -Member No " & memberno
    txtFile.WriteLine "SNo :" & txtSNo
    txtFile.WriteLine "Name :" & name
'    Else
    txtFile.WriteLine "Quantity Supplied :" & CummulKgs & " Kgs"
    Startdate = DateSerial(year(DTPStmts), month(DTPStmts) - 1, 1)
    'sql = "d_sp_TotalMonth " & txtSNo & ",'" & StartDate & "','" & DTPMilkDate & "'"
    Set rs = New ADODB.Recordset
    sql = "d_sp_TotalMonth " & txtSNo & ",'" & Startdate & "','" & DTPStmts & "'"
    Set rs = oSaccoMaster.GetRecordset(sql)
    If Not rs.EOF Then
    If Not IsNull(rs.Fields(0)) Then CummulKgs = rs.Fields(0)
    Else
    CummulKgs = "0.00"
    End If
    txtFile.WriteLine "Cummulative This Month " & Format(CummulKgs, "#,##0.00" & " Kgs")
'    End If
    Set rs = New ADODB.Recordset
    sql = "d_sp_TransName '" & txtSNo & "'"
    Set rs = oSaccoMaster.GetRecordset(sql)
    If Not rs.EOF Then
    If Not IsNull(rs.Fields(0)) Then TRANSPORTER = rs.Fields(0)
    Else
    TRANSPORTER = "Self"
    End If
    txtFile.WriteLine "---------------------------------------"
    txtFile.WriteLine "Transporter :" & TRANSPORTER
    txtFile.WriteLine "Received by :" & username
    txtFile.WriteLine "---------------------------------------"
    txtFile.WriteLine "Date :" & Format(Get_Server_Date, "dd/mm/yyyy HH:MM:SS AM/PM")
    txtFile.WriteLine "     " & motto & ""
    txtFile.WriteLine "---------------------------------------"
    txtFile.WriteLine escFeedAndCut

txtFile.Close
End If

'**********************************endtonotepad
If optAdvanceSlip.value = True Then
'--Net amount as at date
'    Startdate = DateSerial(Year(txtransdate), month(txtransdate), 1)
'Enddate = DateSerial(Year(txtransdate), month(txtransdate) + 1, 1 - 1)

Set rs = oSaccoMaster.GetRecordset("d_sp_SupNet " & txtSNo & ",'" & Startdate & "','" & Enddate & "', 0")

Dim Kgs As Double
If Not IsNull(rs.Fields(0)) Then
Kgs = rs.Fields(0)
Else
Kgs = "0.00"
End If

Dim Gross As Double

If Not IsNull(rs.Fields(1)) Then
Gross = rs.Fields(1)
Else
Gross = "0.00"
End If
Dim Kainet As String
If Not IsNull(rs.Fields(2)) Then
Kainet = rs.Fields(2)
Else
Kainet = "XXXXX XXXX"
End If

Dim Ded As Double

Set rs = oSaccoMaster.GetRecordset("d_sp_SupNet " & txtSNo & ",'" & Startdate & "','" & Enddate & "', 1")
If Not rs.EOF Then
If Not IsNull(rs.Fields(0)) Then
Ded = rs.Fields(0)
Else
Ded = "0.00"
End If
End If
 Dim Net As Double
Net = Format((CCur(Gross) - CCur(Ded)), "#,##0.00")

        ttt = ports 'LPT1,LPT2....
        Set fso = CreateObject("Scripting.FileSystemObject")
        On Error GoTo err
        
        'Set chkPrinter = fso.GetFile(ttt)
       
        
    Set txtFile = fso.CreateTextFile(ttt, True)
    txtFile.WriteLine escAlignCenter
    txtFile.WriteLine "Advance Slip"
    txtFile.WriteLine "" & cname & ""
    txtFile.WriteLine "........................................"
    txtFile.WriteLine escAlignLeft
    txtFile.WriteLine "SNo. : " & txtSNo
    txtFile.WriteLine "Names : " & Kainet
    txtFile.WriteLine "Issue Items/Services worth not more than"
    txtFile.WriteLine "Kshs. : " & Format(Net, "#,##0.00") & ""
    txtFile.WriteLine "Sign"
    txtFile.WriteLine "___________________________"
    txtFile.WriteLine UCase(username)
    txtFile.Write "Date " & Format(Get_Server_Date, "dd/mm/yyyy")
    txtFile.WriteLine ", Time : " & Time
    txtFile.WriteLine "........................................"
    txtFile.WriteLine escFeedAndCut
    
    
    
End If
    
'----d_sp_PrintStmt @SNo bigint,@EndPeriod varchar(10)

If OptNormalStmt.value = True Then
Dim rsnorm As New ADODB.Recordset, sno1 As Long
Set rsnorm = oSaccoMaster.GetRecordset("select sno  from d_transport where active=1  order by sno ")
While Not rsnorm.EOF
sno1 = IIf(IsNull(rsnorm.Fields(0)), 0, rsnorm.Fields(0))
    Set rst = New ADODB.Recordset
        sql = "d_sp_PrintStmt " & sno1 & ",'" & Enddate & "'"
    Set rst = oSaccoMaster.GetRecordset(sql)
If rst.EOF Then
    MsgBox "There is no record in the payroll for supplier number " & sno1, vbInformation
        txtSNo.SetFocus
    Exit Sub
End If



        ttt = ports 'LPT1,LPT2....
        Set fso = CreateObject("Scripting.FileSystemObject")
        On Error GoTo err
        
'        Set chkPrinter = fso.GetFile(ttt)
       
        
    Set txtFile = fso.CreateTextFile(ttt, True)
    txtFile.WriteLine escAlignCenter
    txtFile.WriteLine "" & cname & ""
    txtFile.WriteLine "" & paddress & ""
    txtFile.WriteLine "" & town & ""
    txtFile.WriteLine "MILK STATEMENT FOR " & UCase(Format(DTPStmts, "MMMM/YYYY"))
    txtFile.WriteLine escAlignLeft
    txtFile.WriteLine "........................................"
    txtFile.WriteLine "SNo :" & sno1
    txtFile.WriteLine "Name :" & rst!NAMES
    txtFile.WriteLine "........................................"
    'startdate = DateSerial(Year(DTPStmts), month(DTPStmts) - 1, 1)
    Set rs = New ADODB.Recordset
    sql = "d_sp_TotalMonth " & sno1 & ",'" & Startdate & "','" & Enddate & "'"
    Set rs = oSaccoMaster.GetRecordset(sql)
    If Not rs.EOF Then
    If Not IsNull(rs.Fields(0)) Then CummulKgs = rs.Fields(0)
    Else
    CummulKgs = "0.00"
    End If
    
    txtFile.WriteLine "Total Kgs :" & Format(CummulKgs, "#,##0.00" & " Kgs")
    
    txtFile.WriteLine "Gross Amount               Kshs: " & Format(rst!GPay, "#,##0.00") & ""
    txtFile.Write escBoldOn
    txtFile.WriteLine "DEDUCTIONS"
    txtFile.WriteLine "-------------"
    txtFile.Write escBoldOff
    txtFile.WriteLine "Transport        Kshs: " & Format(rst!Transport, "#,##0.00") & ""
    txtFile.WriteLine "Agrovet          Kshs: " & Format(rst!agrovet, "#,##0.00") & ""
    txtFile.WriteLine "TM Shares        Kshs: " & Format(rst!TMShares, "#,##0.00") & ""
    txtFile.WriteLine "H Shares         Kshs: " & Format(rst!HShares, "#,##0.00") & ""
    txtFile.WriteLine "Advances         Kshs: " & Format(rst!Advance, "#,##0.00") & ""
    txtFile.WriteLine "FSA              Kshs: " & Format(rst!FSA, "#,##0.00") & ""
    txtFile.WriteLine "AI               Kshs: " & Format(rst!AI, "#,##0.00") & ""
    txtFile.WriteLine "Others           Kshs: " & Format(rst!Others, "#,##0.00") & ""
    txtFile.WriteLine "Total Deductions Kshs: " & Format(rst!TDeductions, "#,##0.00") & ""
    txtFile.WriteLine "........................................"
    txtFile.WriteLine "NET PAY                    Kshs: " & Format(rst!NPay, "#,##0.00") & ""
    txtFile.WriteLine "........................................"
    txtFile.WriteLine "BANK DETAILS"
    txtFile.WriteLine "-------------"
    txtFile.WriteLine "Bank Name :" & rst!bank & ""
    txtFile.WriteLine "Bank Branch :" & rst!BBranch
    txtFile.WriteLine "Account Number :" & rst!accountnumber

    txtFile.WriteLine "---------------------------------------"
    txtFile.WriteLine "        Date :" & Format(Get_Server_Date, "dd/mm/yyyy HH:MM:SS AM/PM")
    txtFile.WriteLine "         " & motto & ""
    txtFile.WriteLine "---------------------------------------"
    txtFile.WriteLine escFeedAndCut
    txtFile.Close
'End If
Exit Sub
err: MsgBox err.description & " or There is no printer connected."


rsnorm.MoveNext
Wend
End If
'--Print detailed statement
If OptDetailedStmt.value = True Then

  ttt = ports 'LPT1,LPT2....
        Set fso = CreateObject("Scripting.FileSystemObject")
        On Error GoTo err
        
        'Set chkPrinter = fso.GetFile(ttt)
       
        
        Set txtFile = fso.CreateTextFile(ttt, True)
    txtFile.WriteLine escAlignCenter
    txtFile.WriteLine "" & cname & ""
    txtFile.WriteLine "" & paddress & ""
    txtFile.WriteLine "" & town & ""
    txtFile.WriteLine "DETAILED STATEMENT FOR " & UCase(Format(DTPStmts, "MMMM/YYYY"))
    txtFile.WriteLine escAlignLeft
    '//PUT HERE THE TRANSPORTER
    Dim rtg As New ADODB.Recordset, sno3 As String
    Set rtg = oSaccoMaster.GetRecordset("SELECT     TOP 1 Trans_Code, Sno   FROM         d_Transport WHERE     (Sno = " & txtSNo & ")  ORDER BY auditdatetime DESC")
    If Not rtg.EOF Then
    sno3 = IIf(IsNull(Trim(rtg.Fields(0))), "Self", Trim(rtg.Fields(0)))
    Else
    sno3 = "Self"
    End If
    txtFile.WriteLine "Transporter :" & sno3
    txtFile.WriteLine "........................................"
    txtFile.WriteLine "SNo :" & txtSNo
    
    Set rs = New ADODB.Recordset
    sql = "d_sp_PrintDedStmt " & txtSNo & ",'" & Startdate & "','" & Enddate & "'"
    Set rs = oSaccoMaster.GetRecordset(sql)
    If rs.EOF Then
    MsgBox "The supplier did not supplier for the month specified", vbInformation
    
    txtFile.WriteLine "---------------------------------------"
    txtFile.WriteLine escFeedAndCut
    txtFile.Close
    Exit Sub
    End If
    
    txtFile.WriteLine "Name :" & rs!NAMES
    txtFile.WriteLine "........................................"
    txtFile.WriteLine "DATE " & vbTab & "" & vbTab & "QNTY" & vbTab & "PRICE" & vbTab & "PAYABLE"
    txtFile.WriteLine "........................................"
    sql = ""
    sql = "SELECT SUM(d_Shares.Amnt) AS TotalShares FROM d_Shares where d_Shares.Code = CONVERT(varchar(35)," & txtSNo & ")"
    Set rs2 = oSaccoMaster.GetRecordset(sql)
    Dim qnty As Currency, GPay As Currency
         qnty = 0
         GPay = 0
         
         
        While Not rs.EOF
        Dim Pric As Currency
        Pric = rs!ppu
        If Not IsNull(rs2.Fields(0)) Then
        If rs2.Fields(0) > 4999 Then
          Pric = (rs!ppu) + 1
        
        End If
        End If
        
        'MsgBox rs!QSupplied
        
        txtFile.WriteLine rs!transdate & " " & vbTab & " " & Format(rs!QSupplied, "#,##0.0#") & " " & vbTab & " " & Format(Pric, "#,##0.00") & " " & vbTab & " " & Format(((Pric) * rs!QSupplied), "#,##0.00")
        'txtfile.WriteLine rs!PPU
        qnty = qnty + rs!QSupplied
        GPay = GPay + (Pric * rs!QSupplied)
         rs.MoveNext
        
        Wend
Set rs2 = New ADODB.Recordset
'Dim Startdate As String, Enddate As String

'Startdate = DateSerial(DTPMilkDate, cboMonth, 1)
sql = "d_sp_UpdateGPAYQnty '" & Startdate & "','" & Enddate & "'," & txtSNo & ""
Set rs2 = oSaccoMaster.GetRecordset(sql)
If Not rs2.EOF Then
If Not IsNull(rs2.Fields(0)) Then qnty = rs2.Fields(0)
'-If Not IsNull(rs2.Fields(1)) Then GPay = rs2.Fields(1)
End If

    
    txtFile.WriteLine "........................................"
    txtFile.WriteLine "Total Kgs :" & Format(qnty, "#,##0.00" & " Kgs")
    txtFile.WriteLine "Gross Pay Kshs :" & Format(GPay, "#,##0.00" & "")
    txtFile.WriteLine "........................................"
    txtFile.WriteLine escBoldOn
    txtFile.WriteLine "DEDUCTIONS"
    txtFile.WriteLine "........................................"
    txtFile.WriteLine escBoldOff
    
Set rst = New ADODB.Recordset
sql = "d_sp_PrintDeductStmt " & txtSNo & ",'" & Startdate & "','" & Enddate & "'"
Set rst = oSaccoMaster.GetRecordset(sql)
    
    txtFile.WriteLine "........................................"
    txtFile.WriteLine "DATE " & vbTab & "" & vbTab & "AMOUNT" & vbTab & "DESCRIPTION"
    txtFile.WriteLine "........................................"
    Dim TotDeduction As Double
    TotDeduction = 0
    While Not rst.EOF
        'MsgBox rs!QSupplied
        txtFile.WriteLine rst!date_deduc & " " & vbTab & " " & Format(rst!amount, "#,##0.00" & vbTab & " " & rst!description & " " & rst!Remarks & " ")
        TotDeduction = TotDeduction + rst!amount
        'txtfile.WriteLine rs!PPU
         rst.MoveNext
        
        Wend
    Set rst1 = New ADODB.Recordset
        sql = "d_sp_PrintStmt " & txtSNo & ",'" & Enddate & "'"
    Set rst1 = oSaccoMaster.GetRecordset(sql)
    If Not IsNull(rst1!Transport) Then
              txtFile.WriteLine Enddate & " " & vbTab & " " & Format(rst1!Transport, "#,##0.00" & vbTab & " " & "Transport ")
              TotDeduction = TotDeduction + rst1!Transport
    End If
    txtFile.WriteLine "........................................"
    txtFile.WriteLine "Total Deductions Kshs: " & Format(TotDeduction, "#,##0.00") & ""
    txtFile.WriteLine "........................................"
    txtFile.WriteLine "NET PAY                   Kshs :" & Format((GPay - TotDeduction), "#,##0.00") & ""
    txtFile.WriteLine "-----------------------------------------"
    txtFile.WriteLine "BANK DETAILS"
    txtFile.WriteLine "-------------"
    txtFile.WriteLine "Bank Name :" & rst1!bank & ""
    txtFile.WriteLine "Bank Branch :" & rst1!BBranch
    txtFile.WriteLine "Account Number :" & rst1!accountnumber
    txtFile.WriteLine "---------------------------------------"
    txtFile.WriteLine "        Date :" & Format(Get_Server_Date, "dd/mm/yyyy HH:MM:SS AM/PM")
    txtFile.WriteLine "         " & motto & ""
    txtFile.WriteLine "---------------------------------------"
    txtFile.WriteLine escFeedAndCut
    txtFile.Close
    End If
    
    If OptNormalA4.value = True Then
    reportname = "d_StmtA4.rpt"
    '{d_Payroll.NPay} > 0 and {d_Payroll.Bank} <> '' and month({d_Payroll.EndofPeriod})= month(30/09/2010)  AND year({d_Payroll.EndofPeriod}) = Year(30/09/2010)
    STRFORMULA = "{d_Payroll.SNo}= " & txtSNo & " and month({d_Payroll.EndofPeriod})=" & month(DTPStmts) & " AND year({d_Payroll.EndofPeriod}) =" & year(DTPStmts)
    Show_Sales_Crystal_Report STRFORMULA, reportname, ""

'    d_StmtA4
    End If
    txtSNo = ""
    Exit Sub
errorhandler22:
    MsgBox err.description

End Sub

Private Sub cmdViewreport_Click()
    reportname = "suppliersstatement.rpt"
    STRFORMULA = "{D_SUPPLIERS.SNO}=" & txtSNo & ""
    Show_Sales_Crystal_Report STRFORMULA, reportname, ""
End Sub



Private Sub eward_Click()

End Sub

Private Sub Form_Load()
DTPStmts = Format(Get_Server_Date, "dd/mm/yyyy")
Enddate = DateSerial(year(DTPStmts), month(DTPStmts) + 1, 1 - 1)
DTPStmts = Enddate
End Sub


Private Sub txtSNo_KeyPress(KeyAscii As Integer)
If (KeyAscii = 8) Or (KeyAscii = 48) Or (KeyAscii = 49) Or (KeyAscii = 50) Or (KeyAscii = 51) Or (KeyAscii = 52) Or (KeyAscii = 53) Or (KeyAscii = 54) Or (KeyAscii = 55) Or (KeyAscii = 56) Or (KeyAscii = 57) Then
KeyAscii = KeyAscii
Else
KeyAscii = 0
MsgBox "Please enter a number "
End If
End Sub
