VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmshares 
   BackColor       =   &H00FF00FF&
   Caption         =   "Shares "
   ClientHeight    =   6390
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12450
   LinkTopic       =   "Form1"
   ScaleHeight     =   6390
   ScaleWidth      =   12450
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cboType 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      ItemData        =   "frmshares.frx":0000
      Left            =   1440
      List            =   "frmshares.frx":000A
      TabIndex        =   22
      Text            =   "Shares"
      Top             =   3600
      Width           =   2055
   End
   Begin VB.TextBox txtSNo 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   1440
      TabIndex        =   21
      Top             =   1080
      Width           =   2055
   End
   Begin VB.OptionButton optSupplier 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0FF&
      Caption         =   "Supplier"
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   0
      TabIndex        =   20
      Top             =   0
      Value           =   -1  'True
      Width           =   1335
   End
   Begin VB.OptionButton optTransport 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0FF&
      Caption         =   "Transporter"
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   1440
      TabIndex        =   19
      Top             =   0
      Width           =   1695
   End
   Begin VB.TextBox txtIdNo 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   1440
      TabIndex        =   17
      Top             =   480
      Width           =   2055
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   435
      Left            =   4560
      TabIndex        =   16
      Top             =   5640
      Width           =   975
   End
   Begin VB.TextBox txtAmnt 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1440
      TabIndex        =   15
      Top             =   4320
      Width           =   2055
   End
   Begin VB.TextBox txtName 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   4800
      TabIndex        =   14
      Top             =   480
      Width           =   5175
   End
   Begin VB.OptionButton optNon 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0FF&
      Caption         =   "Non Supplier/Transporter"
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   3240
      TabIndex        =   13
      Top             =   0
      Width           =   3255
   End
   Begin VB.ComboBox cboSex 
      Appearance      =   0  'Flat
      Height          =   315
      ItemData        =   "frmshares.frx":0024
      Left            =   1440
      List            =   "frmshares.frx":0031
      TabIndex        =   12
      Text            =   "M"
      Top             =   2880
      Width           =   735
   End
   Begin VB.ComboBox cboLocation 
      Appearance      =   0  'Flat
      Height          =   315
      ItemData        =   "frmshares.frx":003E
      Left            =   4920
      List            =   "frmshares.frx":0040
      TabIndex        =   11
      Top             =   1200
      Width           =   3255
   End
   Begin VB.CheckBox Check3 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      Caption         =   "Print Receipt"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   3840
      TabIndex        =   10
      Top             =   4920
      Width           =   2175
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Payment Mode"
      Height          =   1215
      Left            =   7320
      TabIndex        =   6
      Top             =   3600
      Width           =   3135
      Begin VB.OptionButton optCash 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Cash"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.OptionButton optCheckOff 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Check Off"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1080
         TabIndex        =   8
         Top             =   360
         Width           =   1455
      End
      Begin VB.CheckBox chkMonthlyD 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         Caption         =   "Deduct Every Month"
         BeginProperty Font 
            Name            =   "Sylfaen"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   720
         Visible         =   0   'False
         Width           =   2895
      End
   End
   Begin VB.CommandButton cmdclose 
      Caption         =   "Close"
      Height          =   435
      Left            =   5880
      TabIndex        =   5
      Top             =   5640
      Width           =   975
   End
   Begin VB.TextBox txtmemberno 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   6600
      TabIndex        =   3
      Top             =   1800
      Width           =   2295
   End
   Begin VB.TextBox txtopeningbal 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   10680
      TabIndex        =   2
      Top             =   1200
      Width           =   1575
   End
   Begin VB.TextBox lblMaxAmnt 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   10440
      TabIndex        =   1
      Top             =   2880
      Width           =   1815
   End
   Begin VB.CheckBox chkop 
      Caption         =   "Open"
      Height          =   315
      Left            =   10560
      TabIndex        =   0
      Top             =   600
      Width           =   1335
   End
   Begin MSComCtl2.DTPicker DTPregdate 
      Height          =   375
      Left            =   2520
      TabIndex        =   4
      Top             =   1800
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      _Version        =   393216
      Format          =   119537665
      CurrentDate     =   40637
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   5280
      TabIndex        =   18
      Top             =   3600
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393216
      Format          =   119537665
      CurrentDate     =   40442
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   5280
      TabIndex        =   23
      Top             =   4320
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393216
      Format          =   119537665
      CurrentDate     =   40442
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0FF&
      Caption         =   "SNo"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   38
      Top             =   1080
      Width           =   495
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0FF&
      Caption         =   "ID Number"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   37
      Top             =   480
      Width           =   1410
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0FF&
      Caption         =   "Type"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   36
      Top             =   3600
      Width           =   615
   End
   Begin VB.Label Period 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0FF&
      Caption         =   "Trans Date"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3840
      TabIndex        =   35
      Top             =   3600
      Width           =   1305
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0FF&
      Caption         =   "Period"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3840
      TabIndex        =   34
      Top             =   4320
      Width           =   795
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0FF&
      Caption         =   "Amount"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   33
      Top             =   4320
      Width           =   1005
   End
   Begin VB.Label lblShares 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   405
      Left            =   5160
      TabIndex        =   32
      Top             =   2880
      Width           =   1515
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0FF&
      Caption         =   "Name :"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3840
      TabIndex        =   31
      Top             =   480
      Width           =   855
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0FF&
      Caption         =   "Location"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3840
      TabIndex        =   30
      Top             =   1200
      Width           =   1035
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0FF&
      Caption         =   "Sex"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   29
      Top             =   2880
      Width           =   405
   End
   Begin VB.Label Label10 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Shares Account : Kshs "
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   405
      Left            =   2400
      TabIndex        =   28
      Top             =   2880
      Width           =   2745
   End
   Begin VB.Label Label11 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Maximum Amount : Kshs "
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   405
      Left            =   7080
      TabIndex        =   27
      Top             =   2880
      Width           =   3225
   End
   Begin VB.Label Label4 
      Caption         =   "Registration Date"
      Height          =   255
      Left            =   120
      TabIndex        =   26
      Top             =   1800
      Width           =   2175
   End
   Begin VB.Label Label12 
      Caption         =   "Member No"
      Height          =   375
      Left            =   4920
      TabIndex        =   25
      Top             =   1800
      Width           =   1575
   End
   Begin VB.Label Label13 
      Caption         =   "Openning Bal"
      Height          =   375
      Left            =   8640
      TabIndex        =   24
      Top             =   1200
      Width           =   1815
   End
End
Attribute VB_Name = "frmshares"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cboSex_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub cmdclose_Click()
Unload Me
End Sub

Private Sub cmdsave_Click()
If Trim(txtName) = "" Then
    MsgBox "Please enter name."
    txtName.SetFocus
    Exit Sub
End If

 If optSupplier.value = True Then
    If txtSNo = "" Then
        MsgBox "Please enter supplier number."
        Exit Sub
        txtSNo.SetFocus
    End If
    
If Trim(cboLocation) = "" Then
    MsgBox "Please select location"
        cboLocation.SetFocus
    Exit Sub
End If
    
    If Trim(txtIdNo) = "" Then
    MsgBox "Please select location"
        txtIdNo.SetFocus
    Exit Sub
End If
    
    
Set Rst = oSaccoMaster.GetRecordset("SELECT SNo FROM d_Suppliers WHERE SNo = '" & txtSNo & "'")

If Rst.RecordCount = 0 Then
 MsgBox "Please enter a valid supplier number."
 txtSNo.SetFocus
 Exit Sub
End If
End If

If optTransport.value = True Then
    If txtSNo = "" Then
        MsgBox "Please enter transporter code."
        Exit Sub
        txtSNo.SetFocus
    End If
    Set rs = oSaccoMaster.GetRecordset("SELECT TransCode FROM d_Transporters WHERE TransCode = '" + txtSNo + "'")

If rs.RecordCount = 0 Then
 MsgBox "Please enter a valid transporter code."
 txtSNo.SetFocus
 Exit Sub
End If
End If
Dim cash As Integer

If optCash.value = True Then
cash = 1
Else
cash = 0
End If
If Trim(lblShares) = "" Then
lblShares = "0.00"
End If


Dim desc As String
desc = cboType
 Set Rst = oSaccoMaster.GetRecordset("SELECT * FROM d_Suppliers WHERE SNo = '" & txtSNo & "'")
 Set RShares1 = oSaccoMaster.GetRecordset("SELECT SNo FROM d_SharesReport WHERE Sno = '" & txtSNo & "'")
 If RShares1.EOF Then
  Dim namecheck As String
 namecheck = Replace(Rst!NAMES, "'", "")
    sql = ""
    sql = "set dateformat dmy insert into d_SharesReport(Sno, Name, IDNo, Type, Amount)"
    sql = sql & " values ('" & txtSNo & "','" & namecheck & "','" & Rst!idno & "','SHARES','0') "
    oSaccoMaster.ExecuteThis (sql)
 End If

Dim rss As New Recordset, amt As Double, rsts As New Recordset, shareamt As Double, TXTshares As Double
Set rsts = oSaccoMaster.GetRecordset("SELECT    SUM(Amount) AS amtt From d_sconribution WHERE     (transdescription LIKE '%shares%') AND (SNo = '" & txtSNo & "')")
If Not rsts.EOF Then
shareamt = IIf(IsNull(rsts!amtt), 0, rsts!amtt)
End If
Set rss = oSaccoMaster.GetRecordset("SELECT    SUM(Amount) AS amt From d_supplier_deduc WHERE     (Description LIKE '%shares%') AND (SNo = '" & txtSNo & "')")
If Not rss.EOF Then
TXTshares = IIf(IsNull(rss!amt), 0, rss!amt) + shareamt
End If


If chkop = vbChecked Then GoTo HEREEE
    If optCash.value = True Then
    Enddate = DateSerial(year(DTPicker2), month(DTPicker2) + 1, 1 - 1)
    
    '//update the regdate and memberno
    sql = ""
    sql = "set dateformat dmy update d_SharesReport set Amount='" & TXTshares + txtAmnt & "' where sno='" & txtSNo & "'"
    oSaccoMaster.ExecuteThis (sql)
    
    'insert inot d_contrib
    Dim rsamount As Double
    'SELECT     idno, transdate, amount, bal, transdescription, auditid  FROM         d_sconribution
    sql = ""
    sql = "set dateformat dmy insert into d_sconribution(sno, transdate, amount, bal, transdescription, auditid,toledgers,datepostedtoledger,Remarks)"
    sql = sql & " values('" & txtSNo & "','" & DTPicker1 & "','" & txtAmnt & "','" & txtAmnt & "','" & cboType & "','" & User & "','0','" & DTPicker1 & "','" & cash & "') "
    oSaccoMaster.ExecuteThis (sql)
    
    End If

    If optCheckOff.value = True Then
        Startdate = DateSerial(year(DTPicker2), month(DTPicker2), 1)
        Enddate = DateSerial(year(DTPicker2), month(DTPicker2) + 1, 1 - 1)
        
        strSQL = "d_sp_Shares '" & txtIdNo & "','" & txtSNo & "','" & txtName & "','" & cboSex.Text & "','" & cboLocation.Text & "','" & cboType.Text & "','"
        strSQL = strSQL & DTPicker1 & "','" & cash & "','" & Enddate & "'," & txtAmnt & "," & txtAmnt & ",'" & Enddate & "','" & User & "'"
        oSaccoMaster.ExecuteThis (strSQL)
        
        sql = ""
        sql = "set dateformat dmy insert into d_sconribution(sno, transdate, amount, bal, transdescription, auditid,toledgers,datepostedtoledger,Remarks)"
        sql = sql & " values('" & txtSNo & "','" & DTPicker1 & "','" & txtAmnt & "','" & txtAmnt & "','" & cboType & "','" & User & "','0','" & DTPicker1 & "','" & cash & "') "
        oSaccoMaster.ExecuteThis (sql)
        
        sql = "d_SP_PreSets '" & txtSNo & "','" & desc & "','','" & Startdate & "','" & txtAmnt & "',0,'" & User & "',0,0"
        oSaccoMaster.ExecuteThis (sql)
        sql = ""
        sql = "update d_shares set regdate='" & DTPregdate & "',mno='" & txtmemberno & "',bal=" & IIf(txtopeningbal = "", 0, txtopeningbal) & " where sno='" & txtSNo & "' and idno='" & txtIdNo & "'"
        oSaccoMaster.ExecuteThis (sql)
    End If
'/////do it here
Dim txtTCHPBalances As Double
If chkop = vbChecked Then

HEREEE:

    Set Rst = New ADODB.Recordset
    sql = "select bal from d_shares where sno= '" & txtSNo & "'"
    Set Rst = oSaccoMaster.GetRecordset(sql)
    If Not Rst.EOF Then
    txtTCHPBalances = Rst.Fields(0)
    
     '//get the balance
    
        sql = "SELECT bal FROM d_sconribution  WHERE sno ='" & txtSNo & "'  ORDER BY transdate DESC, id DESC "
        Dim rr As New ADODB.Recordset
        Set rr = oSaccoMaster.GetRecordset(sql)
        If Not rr.EOF Then
        txtTCHPBalances = txtTCHPBalances + CCur(txtAmnt)
        ',[sno],[transdate],[amount],[bal],[transdescription],[auditid],[auditdate],[mno]
          'From [EASYTEA].[dbo].[d_sconribution]
          sql = ""
          sql = "set dateformat dmy insert into d_sconribution([sno],[transdate],[amount],[bal],[transdescription],[auditid])"
          sql = sql & " values ('" & txtSNo & "','" & DTPicker1 & "'," & txtAmnt & "," & txtTCHPBalances & ",'Shares-Openning Bal','" & User & "') "
          oSaccoMaster.ExecuteThis (sql)
          
          sql = ""
          sql = "update d_shares set bal=" & txtTCHPBalances & ",amount=" & txtopeningbal & " where sno='" & txtSNo & "' "
          oSaccoMaster.ExecuteThis (sql)
        'txtTCHPBALANCE = rr.Fields(0)
        End If
    Else
        '//add new one
        txtTCHPBalances = 0
        sql = "insert into d_Shares(sno,idno, Cash,bal,auditid)"
        sql = sql & " values('" & txtSNo & "','" & txtIdNo & "',1,'" & txtAmnt & "','" & User & "')"
        oSaccoMaster.ExecuteThis (sql)
        sql = ""
        sql = "set dateformat dmy insert into d_sconribution([sno],[transdate],[amount],[bal],[transdescription],[auditid])"
        sql = sql & " values ('" & txtSNo & "','" & DTPicker1 & "','" & txtAmnt & "','" & txtAmnt & "','Shares-Openning Bal','" & User & "') "
        oSaccoMaster.ExecuteThis (sql)
        
          sql = ""
          sql = "update d_shares set amount='" & txtopeningbal & "' where sno='" & txtSNo & "' "
          oSaccoMaster.ExecuteThis (sql)
    
    End If
End If

MsgBox "Records saved successfully!"
txtSNo = ""
txtAmnt = ""
txtName = ""

End Sub
Private Sub Form_Load()
 Set rs = CreateObject("adodb.recordset")
    
    DTPicker1 = Format(Get_Server_Date, "dd/mm/yyyy")
    DTPregdate = DTPicker1
    DTPicker2 = DTPicker1
    Set rs = oSaccoMaster.GetRecordset("SELECT LName FROM d_Location")
    
    If rs.EOF Then Exit Sub
    
    With rs
        
        While Not .EOF
         If Not IsNull(rs.Fields("LName")) Then
         cboLocation.AddItem rs.Fields("LName")
         End If
         
         .MoveNext
        
        Wend
    
    End With
    Enddate = DateSerial(year(DTPicker2), month(DTPicker2) + 1, 1 - 1)

End Sub

Private Sub Optcash_Click()
chkMonthlyD.Visible = False
End Sub

Private Sub optCheckOff_Click()
chkMonthlyD.Visible = True
chkMonthlyD.value = False
End Sub

Private Sub optNon_Click()
Optcash_Click
optCash.value = True

optCheckOff.Enabled = False
txtSNo.Visible = False
Label1.Visible = False
End Sub

Private Sub optSupplier_Click()
txtSNo.Visible = True
Label1.Visible = True
optCheckOff.Enabled = True

txtName = ""
txtSNo = ""
Label1.Caption = "SNo"
End Sub

Private Sub opttransport_Click()
txtSNo.Visible = True
Label1.Visible = True
optCheckOff.Enabled = True

txtName = ""
txtSNo = ""
Label1.Caption = "Code"
End Sub



Private Sub txtIdNo_Validate(Cancel As Boolean)
If Trim(txtIdNo) = "" Then
    Exit Sub
End If
Enddate = DateSerial(year(DTPicker2), month(DTPicker2) + 1, 1 - 1)


Set rs = oSaccoMaster.GetRecordset("SELECT SUM(Amnt) AS Shares From d_Shares WHERE sno = '" & txtSNo & "'")

If rs.RecordCount > 0 Then
lblShares = Format(rs.Fields(0), "0.00")
Else
lblShares = "0.00"
End If

Set Rst = oSaccoMaster.GetRecordset("SELECT MaxAmnt From d_MaxShares WHERE IdNo = '" & txtIdNo & "'")

If Rst.RecordCount > 0 Then
lblMaxAmnt = Format(Rst.Fields(0), "0.00")
Else
lblMaxAmnt = 20000
End If

sql = "SET dateformat dmy SELECT     Code, Name, Sex, Loc, Type, TransDate, Cash, Amnt"
sql = sql & " From d_Shares WHERE  sNo = '" & txtSNo & "'" 'Period = '" & Enddate & "' AND

Set rs2 = oSaccoMaster.GetRecordset(sql)
If rs2.RecordCount > 0 Then
txtSNo = rs2.Fields(0)
txtName = rs2.Fields(1)
cboSex = rs2.Fields(2)
cboLocation = rs2.Fields(3)
cboType = "Shares"
'If rs2.Fields(4) = "HShares" Then
'cboType = "Shares"
'Else
'cboType = "TMShares"
'End If

DTPicker1 = rs2.Fields(5)
optCash.value = IIf(IsNull(rs2.Fields(6).value), 0, rs2.Fields(6).value)
txtAmnt = rs2.Fields(7)
DTPicker2 = Enddate
End If
End Sub



Private Sub txtSNo_KeyPress(KeyAscii As Integer)
'If optSupplier.value = True Then
'If (KeyAscii = 8) Or (KeyAscii = 48) Or (KeyAscii = 49) Or (KeyAscii = 50) Or (KeyAscii = 51) Or (KeyAscii = 52) Or (KeyAscii = 53) Or (KeyAscii = 54) Or (KeyAscii = 55) Or (KeyAscii = 56) Or (KeyAscii = 57) Then
'KeyAscii = KeyAscii
'Else
'KeyAscii = 0
'MsgBox "Please enter a number "
'End If
'End If
End Sub

Private Sub txtSNo_Validate(Cancel As Boolean)

If Trim(txtSNo) = "" Then
Exit Sub
End If

If optSupplier.value = True Then
txtName = ""
Set rs = oSaccoMaster.GetRecordset("SELECT Names,idno,Location,Type,Regdate FROM d_Suppliers WHERE SNo = '" & txtSNo & "'")
If Not rs.EOF Then
txtName = rs.Fields(0).value
txtIdNo = rs.Fields(1).value
cboSex.Text = IIf(IsNull(rs.Fields(3)), "G", Left(rs.Fields(3), 1))
cboLocation.Text = IIf(IsNull(rs.Fields(2)), "", rs.Fields(2))
DTPregdate = IIf(IsNull(rs.Fields(4)), Format(Get_Server_Date, "dd/mm/yyyy"), Format(rs.Fields(4), "dd/mm/yyyy"))
End If
End If

If optTransport.value = True Then
txtName = ""
Set rs = oSaccoMaster.GetRecordset("SELECT TransName FROM d_Transporters WHERE TransCode = '" + txtSNo + "'")
If Not rs.EOF Then txtName = rs.Fields(0).value
End If

End Sub

