VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmSupplierStmtBonus 
   Caption         =   "Suppliers Bonus statement"
   ClientHeight    =   3510
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   10110
   LinkTopic       =   "Form1"
   ScaleHeight     =   3510
   ScaleWidth      =   10110
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtSNo 
      Appearance      =   0  'Flat
      Height          =   405
      Left            =   1800
      TabIndex        =   18
      Top             =   360
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Supplier Bonus Statements"
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
      Height          =   1095
      Left            =   0
      TabIndex        =   13
      Top             =   1800
      Width           =   9255
      Begin VB.OptionButton OptNormalStmt 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Caption         =   "Normal Statement (Use POS Printer)"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   480
         TabIndex        =   17
         Top             =   360
         Visible         =   0   'False
         Width           =   4095
      End
      Begin VB.OptionButton OptDetailedStmt 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Caption         =   "Bonus Statement (Use POS Printer)"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   480
         TabIndex        =   16
         Top             =   720
         Value           =   -1  'True
         Width           =   4095
      End
      Begin VB.OptionButton OptNormalA4 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Caption         =   "Normal Statement (Use Normal Printer (A4))"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   480
         TabIndex        =   15
         Top             =   1080
         Visible         =   0   'False
         Width           =   4095
      End
      Begin VB.OptionButton optAdvanceSlip 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Caption         =   "Print Advance Slip"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   480
         TabIndex        =   14
         Top             =   1440
         Visible         =   0   'False
         Width           =   4095
      End
   End
   Begin VB.CommandButton cmdPrint 
      Appearance      =   0  'Flat
      Caption         =   "Print"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1800
      TabIndex        =   12
      Top             =   3120
      Width           =   735
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
      Left            =   2520
      TabIndex        =   11
      Top             =   3120
      Width           =   735
   End
   Begin VB.CommandButton cmdroute 
      Caption         =   "Routes"
      Height          =   375
      Left            =   3600
      TabIndex        =   10
      Top             =   3120
      Width           =   1335
   End
   Begin VB.CheckBox chkNotepad 
      Caption         =   "To Notepad"
      Height          =   255
      Left            =   9960
      TabIndex        =   9
      Top             =   480
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.ComboBox ports 
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      ItemData        =   "frmSupplierStmtBonus.frx":0000
      Left            =   5880
      List            =   "frmSupplierStmtBonus.frx":0010
      TabIndex        =   8
      Text            =   "\\127.0.0.1\E-PoS 80mm Thermal Printer"
      Top             =   480
      Width           =   4215
   End
   Begin VB.CheckBox chprint 
      Caption         =   "Use LPT1 Printer"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   6120
      TabIndex        =   7
      Top             =   0
      Width           =   2175
   End
   Begin VB.Frame Frame2 
      Caption         =   "Print as per Branch"
      BeginProperty Font 
         Name            =   "Century"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   2055
      Left            =   9960
      TabIndex        =   2
      Top             =   1080
      Visible         =   0   'False
      Width           =   3255
      Begin VB.ComboBox cbobranch 
         Height          =   315
         Left            =   840
         TabIndex        =   5
         Top             =   1080
         Width           =   2295
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Print All"
         Height          =   375
         Left            =   1200
         TabIndex        =   4
         Top             =   1560
         Width           =   1695
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Print Range?"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   375
         Left            =   720
         TabIndex        =   3
         Top             =   240
         Width           =   2295
      End
      Begin VB.Label Label4 
         Caption         =   "Branch"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   1080
         Width           =   975
      End
   End
   Begin VB.TextBox txtstart 
      Height          =   495
      Left            =   9600
      TabIndex        =   1
      Top             =   1320
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtend 
      Height          =   495
      Left            =   9600
      TabIndex        =   0
      Top             =   2160
      Visible         =   0   'False
      Width           =   855
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
      Height          =   375
      Left            =   1800
      TabIndex        =   19
      Top             =   1320
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      Format          =   121110529
      CurrentDate     =   40109
   End
   Begin MSComDlg.CommonDialog cdgPrint 
      Left            =   5280
      Top             =   3000
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      FileName        =   "c:\receipt.txt"
   End
   Begin MSComCtl2.DTPicker DTPStmts1 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "dd/MM/yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   3
      EndProperty
      Height          =   375
      Left            =   1800
      TabIndex        =   26
      Top             =   840
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      Format          =   121110529
      CurrentDate     =   40109
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0E0FF&
      Caption         =   "Starting Month:"
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
      TabIndex        =   27
      Top             =   840
      Width           =   1395
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
      Left            =   0
      TabIndex        =   25
      Top             =   360
      Width           =   1665
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0E0FF&
      Caption         =   "Ending Month:"
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
      TabIndex        =   24
      Top             =   1320
      Width           =   1320
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
      Left            =   0
      TabIndex        =   23
      Top             =   0
      Width           =   5835
   End
   Begin VB.Label Label6 
      Caption         =   "Printer Port"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4680
      TabIndex        =   22
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label Label5 
      Caption         =   "TO"
      Height          =   375
      Left            =   9720
      TabIndex        =   21
      Top             =   1920
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label7 
      Caption         =   "FROM"
      Height          =   495
      Left            =   9720
      TabIndex        =   20
      Top             =   1080
      Visible         =   0   'False
      Width           =   495
   End
End
Attribute VB_Name = "frmSupplierStmtBonus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Enddate As Date
Dim Startdate As Date
Dim TRANSPORTER As String

Private Sub Check1_Click()
If Check1.value = vbChecked Then
 Check1 = 1
 txtstart.Visible = True
 txtend.Visible = True
 Label5.Visible = True
 Label7.Visible = True
Else
Check1 = 0
 txtstart = ""
 txtend = ""
 Label5 = ""
 Label7 = ""
 txtstart.Visible = False
 txtend.Visible = False
 Label5.Visible = False
 Label7.Visible = False
End If
End Sub

Private Sub chprint_Click()
Ports.Clear
Ports = ""
'//If the drivers are installed it won't matter whether the Port is indicated
' or not it will just work.

If chprint.value = vbChecked Then
Ports.AddItem "LPT1"
Ports = "LPT1"
Ports.AddItem "LPT2"
Ports.AddItem "LPT3"
Ports.AddItem "LPT4"
Ports.AddItem "LPT5"
Else
'Share the printer first the use of 127.0.0.1 which is
'standard IP address for a loopback network connection
'instead of getting the computer name or IP Address
'
Dim prnPrinter As Printer
Dim pr As String
Ports.Clear

For Each prnPrinter In Printers
   If InStr(prnPrinter.DeviceName, "\\") Then
    Ports.AddItem prnPrinter.DeviceName
    If InStr(prnPrinter.DeviceName, "G") Then
    Ports.Text = prnPrinter.DeviceName
    End If
    Else
    Ports.AddItem "\\127.0.0.1\" & prnPrinter.DeviceName
    If InStr(prnPrinter.DeviceName, "G") Then
    Ports.Text = "\\127.0.0.1\" & prnPrinter.DeviceName
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

If DTPStmts1 > DTPStmts Then
    MsgBox "Startdate cannot be Grater than End Date", vbCritical
    Exit Sub
End If

Startdate = DateSerial(year(DTPStmts), month(2), 1)
Enddate = DateSerial(year(DTPStmts), month(DTPStmts) + 1, 1 - 1)
'DTPStmts = Enddate

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
        PORT = Ports
        'ttt = "LPT1" 'LPT1
        ttt = PORT
        'ttt = "LPT1" 'LPT1,LPT2....
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
    Set Rst = New ADODB.Recordset
        sql = "d_sp_PrintStmt " & txtSNo & ",'" & Enddate & "'"
    Set Rst = oSaccoMaster.GetRecordset(sql)
    If Rst.EOF Then
        MsgBox "There is no record in the payroll for supplier number " & txtSNo, vbInformation
            txtSNo.SetFocus
        Exit Sub
    End If


 'Dim PORT As String
        PORT = Ports
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
    txtFile.WriteLine "Name :" & Rst!NAMES
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
    
    txtFile.WriteLine "Gross Amount               Kshs: " & Format(Rst!GPay, "#,##0.00") & ""
    GPay = Format(Rst!GPay, "#,##0.00")
    txtFile.Write escBoldOn
    txtFile.WriteLine "DEDUCTIONS"
    txtFile.WriteLine "-------------"
    txtFile.Write escBoldOff
    Set Rst = New ADODB.Recordset
    sql = "d_sp_PrintDeductStmt " & txtSNo & ",'" & Startdate & "','" & Enddate & "'"
    Set Rst = oSaccoMaster.GetRecordset(sql)
    
    txtFile.WriteLine "........................................"
    txtFile.WriteLine "DATE " & vbTab & "" & vbTab & "AMOUNT" & vbTab & "DESCRIPTION"
    txtFile.WriteLine "........................................"
   ' Dim TotDeduction As Double
    TotDeduction = 0
    While Not Rst.EOF
        'MsgBox rs!QSupplied
        txtFile.WriteLine Rst!Date_Deduc & " " & vbTab & " " & Format(Rst!amount, "#,##0.00" & vbTab & " " & Rst!description & " " & Rst!Remarks & " ")
        TotDeduction = TotDeduction + Rst!amount
        'txtfile.WriteLine rs!PPU
         Rst.MoveNext
        
        Wend
    Set Rst1 = New ADODB.Recordset
        sql = "d_sp_PrintStmt " & txtSNo & ",'" & Enddate & "'"
    Set Rst1 = oSaccoMaster.GetRecordset(sql)
    If Not IsNull(Rst1!Transport) Then
              txtFile.WriteLine Enddate & " " & vbTab & " " & Format(Rst1!Transport, "#,##0.00" & vbTab & " " & "Transport ")
              TotDeduction = TotDeduction + Rst1!Transport
    End If
    txtFile.WriteLine "Quality Type: " & Format(IIf(IsNull(Rst1!Trader), 0, Rst1!Trader), "#,##0.00") & ""
    txtFile.WriteLine "Quality Bonus Kshs: " & Format(IIf(IsNull(Rst1!TCHP), 0, Rst1!TCHP), "#,##0.00") & ""
    txtFile.WriteLine "Can Number: " & Format(IIf(IsNull(Rst1!otheraccno), 0, Rst1!otheraccno)) & ""
    txtFile.WriteLine "Total Deductions Kshs: " & Format(TotDeduction, "#,##0.00") & ""
    txtFile.WriteLine "........................................"
    txtFile.WriteLine "NET PAY                   Kshs :" & Format((GPay - TotDeduction + IIf(IsNull(Rst1!TCHP), 0, Rst1!TCHP)), "#,##0.00") & ""
    txtFile.WriteLine "........................................"
    txtFile.WriteLine "BANK DETAILS"
    txtFile.WriteLine "-------------"
    txtFile.WriteLine "Bank Name :" & Rst!bank & ""
    txtFile.WriteLine "Bank Branch :" & Rst!BBranch
    txtFile.WriteLine "Account Number :" & Rst!accountnumber
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
 Exit Sub
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
                        txtFile.WriteLine "BONUS STATEMENT FOR THE YEAR " & UCase(Format(DTPStmts, "YYYY"))
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
                        sql = "set dateformat dmy select * from d_Bonus2 where Sno='" & txtSNo & "' and Date BETWEEN '" & Startdate & "' AND '" & Enddate & "'"
                        'sql = "d_sp_PrintDedStmt " & txtSNo & ",'" & Startdate & "','" & Enddate & "'"
                        Set rs = oSaccoMaster.GetRecordset(sql)
                        If rs.EOF Then
                        MsgBox "The supplier do not have bonus for the year specified", vbInformation
                        
                        txtFile.WriteLine "---------------------------------------"
                        txtFile.WriteLine escFeedAndCut
                        txtFile.Close
                        Exit Sub
                        End If
                        
                        Set Rst1 = New ADODB.Recordset
                        sql = "set dateformat dmy select Name, Bank, Bcode, Branch from d_Bonus where Sno='" & txtSNo & "' and Enddate BETWEEN '" & Startdate & "' AND '" & Enddate & "'"
                        Set Rst1 = oSaccoMaster.GetRecordset(sql)
                        
                        txtFile.WriteLine "Name :" & Rst1!name
                        txtFile.WriteLine "........................................"
                        txtFile.WriteLine "DATE " & vbTab & "" & vbTab & "AMOUNT"
                        txtFile.WriteLine "........................................"
                        sql = ""
'                        sql = "SELECT SUM(d_Shares.Amnt) AS TotalShares FROM d_Shares where d_Shares.Code = CONVERT(varchar(35)," & txtSNo & ")"
'                        Set rs2 = oSaccoMaster.GetRecordset(sql)
                        'sql = "select sum(spu) as shares from d_shares where sno='" & txtSNo & "'"
                        'Set rs2 = oSaccoMaster.GetRecordset(sql)
                        'Dim qnty As Currency, GPay As Currency
                        Set rs2 = New ADODB.Recordset
                        Dim qnty As Currency
                             qnty = 0
                             GPay = 0
                             Dim m As Integer
                             Dim X As String
                             Dim D As Date
                                m = Format(DateSerial(year(DTPStmts), month(DTPStmts) + 1, 1 - 1), "MM")
                                D = Startdate
                                  For I = 1 To m
                                   X = "Mon" + Trim(I)
                                   sql = "set dateformat dmy select " & X & " AS MON from d_Bonus2 where Sno='" & txtSNo & "' and Date BETWEEN '" & Startdate & "' AND '" & Enddate & "'"
                                   Set rs2 = oSaccoMaster.GetRecordset(sql)
                                   'D = Format(DateSerial(Year(D), month(D) + 1, 1 - 1), "MM-YYYY")
                                    txtFile.WriteLine D & " " & vbTab & " " & Format(rs2!mon, "#,##0.0#")
                                    qnty = qnty + rs2!mon
                                    D = DateSerial(year(D), month(D) + 1, 1)
                                  Next I
                             
                           ' While Not rs.EOF
                           ' Dim Pric As Currency
                           'Pric = rs!ppu
                            'txtFile.WriteLine rs!transdate & " " & vbTab & " " & Format(rs!QSupplied, "#,##0.0#")
                            '& " " & vbTab & " " & Format(Pric, "#,##0.00") & " " & vbTab & " " & Format(((Pric) * rs!QSupplied), "#,##0.00")
                            'txtfile.WriteLine rs!PPU
                            'qnty = qnty + rs!QSupplied
                            'GPay = GPay + (Pric * rs!QSupplied)
                            ' rs.MoveNext
                            
                            'Wend
                    
                        
                        'txtFile.WriteLine "........................................"
                        'txtFile.WriteLine "Total Amount :" & Format(qnty, "#,##0.00" & " kshs")
                        ''txtFile.WriteLine "Gross Pay Kshs :" & Format(GPay, "#,##0.00" & "")
                        'txtFile.WriteLine "........................................"
                        'Set rst1 = New ADODB.Recordset
                        'sql = "d_sp_PrintStmt " & txtSNo & ",'" & Enddate & "'"
                        'Set rst1 = oSaccoMaster.GetRecordset(sql)
                        'If Not IsNull(rst1!Transport) Then
                                  'txtFile.WriteLine Enddate & " " & vbTab & " " & Format(rst1!Transport, "#,##0.00" & vbTab & " " & "Transport ")
                                  'TotDeduction = TotDeduction + rst1!Transport
                        'End If
                        txtFile.WriteLine "........................................"
                        txtFile.WriteLine "NET PAY                   Kshs : " & Format(qnty, "#,##0.00") & ""
                        '" & Format((GPay - TotDeduction + IIf(IsNull(rst1!TCHP), 0, rst1!TCHP)), "#,##0.00") & ""
                        txtFile.WriteLine "-----------------------------------------"
                        txtFile.WriteLine "BANK DETAILS"  'Bank, Bcode, Branch
                        txtFile.WriteLine "-------------"
                        txtFile.WriteLine "Bank Name :" & Rst1!bcode & ""
                        txtFile.WriteLine "Bank Branch :" & Rst1!Branch
                        txtFile.WriteLine "Account Number :" & Rst1!bank
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
                     Dim Bonusprice, amountopay, totalkgs, Totalamount As Double
                     Bonusprice = 2
                     amountopay = 0
                     totalkgs = 0
                     Totalamount = 0
                            PORT = Ports
                            'ttt = "LPT1" 'LPT1
                            ttt = PORT
                      'ttt = "LPT1" 'LPT1,LPT2....
                            Set fso = CreateObject("Scripting.FileSystemObject")
                            On Error GoTo err
                            
                        Set txtFile = fso.CreateTextFile(ttt, True)
                        txtFile.WriteLine escAlignCenter
                        txtFile.WriteLine "" & cname & ""
                        txtFile.WriteLine "" & paddress & ""
                        txtFile.WriteLine "" & town & ""
                        txtFile.WriteLine "BONUS STATEMENT FOR " & UCase(Format(DTPStmts1, "MMMM/YYYY"))
                        txtFile.WriteLine "TO " & UCase(Format(DTPStmts, "MMMM/YYYY"))
                        txtFile.WriteLine escAlignLeft
                        ''get bonus rate
                        sql = "set dateformat dmy select Price from d_PriceBonus "
                        Set rs2 = oSaccoMaster.GetRecordset(sql)
                        If Not rs2.EOF Then
                          Bonusprice = rs2!Price
                        End If
                        '//PUT HERE THE TRANSPORTER
                        'Dim rtg As New ADODB.Recordset, sno3 As String
                        Set rtg = oSaccoMaster.GetRecordset("SELECT TOP 1 Trans_Code, Sno FROM d_Transport WHERE (Sno = " & txtSNo & ")  ORDER BY auditdatetime DESC")
                        If Not rtg.EOF Then
                        sno3 = IIf(IsNull(Trim(rtg.Fields(0))), "Self", Trim(rtg.Fields(0)))
                        Else
                        sno3 = "Self"
                        End If
                        txtFile.WriteLine "Transporter :" & sno3
                        txtFile.WriteLine "........................................"
                        txtFile.WriteLine "SNo :" & txtSNo
                        
                        Set rs = New ADODB.Recordset
                        sql = "set dateformat dmy select SUM(QSupplied) from d_Milkintake where Sno='" & txtSNo & "' and TransDate BETWEEN '" & DTPStmts1 & "' AND '" & DTPStmts & "'"
                        Set rs = oSaccoMaster.GetRecordset(sql)
                            If rs.EOF Then
                            MsgBox "The supplier do not have bonus for the Months specified", vbInformation
                            
                            txtFile.WriteLine "---------------------------------------"
                            txtFile.WriteLine escFeedAndCut
                            txtFile.Close
                            Exit Sub
                            End If
                            
                        Set Rst1 = New ADODB.Recordset
                            sql = "set dateformat dmy select SNo, Names, AccNo, Bcode, BBranch from d_Suppliers where Sno='" & txtSNo & "'"
                            Set Rst1 = oSaccoMaster.GetRecordset(sql)
                            txtFile.WriteLine "Name :" & Rst1!NAMES
                            txtFile.WriteLine "........................................"
                            txtFile.WriteLine "Month " & vbTab & "Kgs " & vbTab & " Rate " & vbTab & "  AMOUNT"
                            txtFile.WriteLine "........................................"
                            
                            Dim EEnddate As Date
                            Enddate = DateSerial(year(DTPStmts1), month(DTPStmts1) + 1, 1 - 1)
                            EEnddate = DateSerial(year(DTPStmts), month(DTPStmts) + 1, 1 - 1)
                        ''sum kgs per month
                        ''sql = "set dateformat dmy select distinct( format(TransDate,'MM/yyy')) as ddate from d_Milkintake where Sno='" & txtSNo & "' AND TransDate BETWEEN '" & DTPStmts1 & "' AND '" & DTPStmts & "' order by ddate asc"
                        ''Set rtg = oSaccoMaster.GetRecordset(sql)
                        'While Not rtg.EOF
                        While Enddate <= EEnddate
                          Dim dateget As String
                            'dateget = "01/" + rtg!ddate
                            dateget = Enddate
                            Startdate = DateSerial(year(dateget), month(dateget), 1)
                            Enddate = DateSerial(year(dateget), month(dateget) + 1, 1 - 1)
                            
                            Set rs = New ADODB.Recordset
                            sql = "set dateformat dmy select isnull(SUM(QSupplied),0) as QSupplied from d_Milkintake where Sno='" & txtSNo & "' and TransDate BETWEEN '" & Startdate & "' AND '" & Enddate & "'"
                            Set rs = oSaccoMaster.GetRecordset(sql)
                            
                            
                            sql = ""
                            Set rs2 = New ADODB.Recordset
                             qnty = 0
                             GPay = 0
                             Dim da As String
                                da = UCase(Format(Startdate, "MM/YYYY"))
                                  
                                amountopay = Bonusprice * rs!QSupplied
                                txtFile.WriteLine da & " " & vbTab & "" & rs!QSupplied & " " & vbTab & "" & Bonusprice & " " & vbTab & "" & Format(amountopay, "#,##0.0#")
                                qnty = qnty + rs!QSupplied
                                totalkgs = totalkgs + qnty
                                Totalamount = Totalamount + amountopay
                                
                                Enddate = Enddate + 1
                                
                          'rtg.MoveNext
                         Wend
                                
                        txtFile.WriteLine "-----------------------------------------"
                        txtFile.WriteLine "Total Kgs                      : " & Format(totalkgs, "#,##0.00") & ""
                        txtFile.WriteLine "Rate Per Kgs              Kshs : " & Format(Bonusprice, "#,##0.00") & ""
                        txtFile.WriteLine "NET PAY                   Kshs : " & Format(Totalamount, "#,##0.00") & ""
                        txtFile.WriteLine "-----------------------------------------"
                        txtFile.WriteLine "BANK DETAILS"
                        txtFile.WriteLine "-------------"
                        txtFile.WriteLine "Bank Name :" & Rst1!bcode & ""
                        txtFile.WriteLine "Bank Branch :" & Rst1!BBranch
                        txtFile.WriteLine "Account Number :" & Rst1!ACCNO
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

        ttt = "LPT1" 'LPT1,LPT2....
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
    Set Rst = New ADODB.Recordset
        sql = "d_sp_PrintStmt " & sno1 & ",'" & Enddate & "'"
    Set Rst = oSaccoMaster.GetRecordset(sql)
If Rst.EOF Then
    MsgBox "There is no record in the payroll for supplier number " & sno1, vbInformation
        txtSNo.SetFocus
    Exit Sub
End If



        ttt = "LPT1" 'LPT1,LPT2....
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
    txtFile.WriteLine "Name :" & Rst!NAMES
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
    
    txtFile.WriteLine "Gross Amount               Kshs: " & Format(Rst!GPay, "#,##0.00") & ""
    txtFile.Write escBoldOn
    txtFile.WriteLine "DEDUCTIONS"
    txtFile.WriteLine "-------------"
    txtFile.Write escBoldOff
    txtFile.WriteLine "Transport        Kshs: " & Format(Rst!Transport, "#,##0.00") & ""
    txtFile.WriteLine "Agrovet          Kshs: " & Format(Rst!agrovet, "#,##0.00") & ""
    txtFile.WriteLine "TM Shares        Kshs: " & Format(Rst!TMShares, "#,##0.00") & ""
    txtFile.WriteLine "H Shares         Kshs: " & Format(Rst!HShares, "#,##0.00") & ""
    txtFile.WriteLine "Advances         Kshs: " & Format(Rst!Advance, "#,##0.00") & ""
    txtFile.WriteLine "FSA              Kshs: " & Format(Rst!FSA, "#,##0.00") & ""
    txtFile.WriteLine "AI               Kshs: " & Format(Rst!AI, "#,##0.00") & ""
    txtFile.WriteLine "Others           Kshs: " & Format(Rst!Others, "#,##0.00") & ""
    txtFile.WriteLine "Total Deductions Kshs: " & Format(Rst!TDeductions, "#,##0.00") & ""
    txtFile.WriteLine "........................................"
    txtFile.WriteLine "NET PAY                    Kshs: " & Format(Rst!NPay, "#,##0.00") & ""
    txtFile.WriteLine "........................................"
    txtFile.WriteLine "BANK DETAILS"
    txtFile.WriteLine "-------------"
    txtFile.WriteLine "Bank Name :" & Rst!bank & ""
    txtFile.WriteLine "Bank Branch :" & Rst!BBranch
    txtFile.WriteLine "Account Number :" & Rst!accountnumber

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

  ttt = "LPT1" 'LPT1,LPT2....
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
        If rs2.Fields(0) > 19999 Then
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
    
Set Rst = New ADODB.Recordset
sql = "d_sp_PrintDeductStmt " & txtSNo & ",'" & Startdate & "','" & Enddate & "'"
Set Rst = oSaccoMaster.GetRecordset(sql)
    
    txtFile.WriteLine "........................................"
    txtFile.WriteLine "DATE " & vbTab & "" & vbTab & "AMOUNT" & vbTab & "DESCRIPTION"
    txtFile.WriteLine "........................................"
    Dim TotDeduction As Double
    TotDeduction = 0
    While Not Rst.EOF
        'MsgBox rs!QSupplied
        txtFile.WriteLine Rst!Date_Deduc & " " & vbTab & " " & Format(Rst!amount, "#,##0.00" & vbTab & " " & Rst!description & " " & Rst!Remarks & " ")
        TotDeduction = TotDeduction + Rst!amount
        'txtfile.WriteLine rs!PPU
         Rst.MoveNext
        
        Wend
    Set Rst1 = New ADODB.Recordset
        sql = "d_sp_PrintStmt " & txtSNo & ",'" & Enddate & "'"
    Set Rst1 = oSaccoMaster.GetRecordset(sql)
    If Not IsNull(Rst1!Transport) Then
              txtFile.WriteLine Enddate & " " & vbTab & " " & Format(Rst1!Transport, "#,##0.00" & vbTab & " " & "Transport ")
              TotDeduction = TotDeduction + Rst1!Transport
    End If
    txtFile.WriteLine "........................................"
    txtFile.WriteLine "Total Deductions Kshs: " & Format(TotDeduction, "#,##0.00") & ""
    txtFile.WriteLine "........................................"
    txtFile.WriteLine "NET PAY                   Kshs :" & Format((GPay - TotDeduction), "#,##0.00") & ""
    txtFile.WriteLine "-----------------------------------------"
    txtFile.WriteLine "BANK DETAILS"
    txtFile.WriteLine "-------------"
    txtFile.WriteLine "Bank Name :" & Rst1!bank & ""
    txtFile.WriteLine "Bank Branch :" & Rst1!BBranch
    txtFile.WriteLine "Account Number :" & Rst1!accountnumber
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

Private Sub Command1_Click()
If cbobranch = "" Then
  MsgBox "Please Enter Branch for suppliers To print statements", vbInformation, Me.Caption
 Exit Sub
End If
Startdate = DateSerial(year(DTPStmts), month(2), 1)
Enddate = DateSerial(year(DTPStmts), month(DTPStmts) + 1, 1 - 1)

If Check1 = 0 Then
Set rst8 = oSaccoMaster.GetRecordset("set dateformat dmy select distinct SNo from d_Bonus2  where SNo" _
   & " in(select SNo from d_Bonus where Branch='" & cbobranch & "') and Date BETWEEN '" & Startdate & "' AND '" & Enddate & "'")
Else
If txtstart = "" Then
  MsgBox "Please Enter Starting Supplier To print statements", vbInformation, Me.Caption
 Exit Sub
End If
If txtend = "" Then
  MsgBox "Please Enter Ending supplier To print statements", vbInformation, Me.Caption
 Exit Sub
End If
Set rst8 = oSaccoMaster.GetRecordset("set dateformat dmy select distinct SNo from d_Bonus2 where SNo" _
   & " in(select SNo from d_Bonus where SNo >='" & txtstart & "' and SNo <='" & txtend & "' and Branch='" & cbobranch & "') and Date BETWEEN '" & Startdate & "' AND '" & Enddate & "'")
'Set rst8 = oSaccoMaster.GetRecordset("select SNo from d_Payroll where SNo >='" & txtstart & "' and SNo <='" & txtend & "' and location='" & cbobranch & "' and EndofPeriod ='" & Enddate & "'")
End If
   'txtSNo
   With rst8
      While Not .EOF
          txtSNo = .Fields(0)
        
          cmdPrint_Click
        .MoveNext
      Wend
'     MsgBox "This supplier '" & txtSNo & "' didnt supply milk this month "
   End With
   MsgBox "You have succesfully print all suppliers for this Branch: '" & cbobranch & "' "
   Check1.value = vbUnchecked
   Exit Sub


End Sub
Private Sub branames()
    cbobranch.Clear
    Set Rst = New Recordset
    Dim cn As Connection
    Set cn = New ADODB.Connection
    Provider = "MAZIWA"
    'Provider = cnn
    cn.Open Provider, "atm", "atm"
    Set Rst = New Recordset
    sql = "Select distinct(location) from d_suppliers ORDER BY location"
    Rst.Open sql, cn, adOpenKeyset, adLockOptimistic
    While Not Rst.EOF
    cbobranch.AddItem Rst.Fields(0)
    Rst.MoveNext
    Wend

End Sub

Private Sub Form_Load()
DTPStmts = Format(Get_Server_Date, "dd/mm/yyyy")
DTPStmts1 = DateSerial(year(DTPStmts), month(2), 1)
Enddate = DateSerial(year(DTPStmts), month(DTPStmts) + 1, 1 - 1)
DTPStmts1 = DTPStmts1
DTPStmts = Enddate
branames
Check1.value = vbUnchecked
 txtstart.Visible = False
 txtend.Visible = False
 Label5.Visible = False
 Label7.Visible = False
End Sub


Private Sub txtSNo_KeyPress(KeyAscii As Integer)
If (KeyAscii = 8) Or (KeyAscii = 48) Or (KeyAscii = 49) Or (KeyAscii = 50) Or (KeyAscii = 51) Or (KeyAscii = 52) Or (KeyAscii = 53) Or (KeyAscii = 54) Or (KeyAscii = 55) Or (KeyAscii = 56) Or (KeyAscii = 57) Then
KeyAscii = KeyAscii
Else
KeyAscii = 0
MsgBox "Please enter a number "
End If
End Sub



