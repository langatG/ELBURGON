VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmsalesreturn 
   Caption         =   "Sales Return"
   ClientHeight    =   8955
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11220
   LinkTopic       =   "Form1"
   ScaleHeight     =   8955
   ScaleWidth      =   11220
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton Optothers 
      Caption         =   "Institutions"
      Height          =   255
      Left            =   7200
      TabIndex        =   64
      Top             =   3720
      Width           =   1455
   End
   Begin VB.TextBox txtcomment 
      Height          =   615
      Left            =   8400
      TabIndex        =   62
      Top             =   4800
      Width           =   2775
   End
   Begin VB.OptionButton optrecieve 
      Caption         =   "Receive From Staion"
      Height          =   375
      Left            =   8760
      TabIndex        =   61
      Top             =   2160
      Width           =   1815
   End
   Begin VB.CommandButton cmdsalesreturn 
      Caption         =   "Sales "
      Height          =   495
      Left            =   9360
      TabIndex        =   60
      Top             =   2880
      Width           =   1335
   End
   Begin VB.TextBox txtrno 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   1560
      TabIndex        =   29
      Top             =   120
      Width           =   2535
   End
   Begin VB.TextBox txtamount 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   1560
      TabIndex        =   27
      Top             =   2520
      Width           =   1935
   End
   Begin VB.TextBox txtquantity 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   1560
      TabIndex        =   26
      Top             =   2040
      Width           =   1935
   End
   Begin VB.TextBox txtpcode 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   1560
      TabIndex        =   25
      Top             =   600
      Width           =   2535
   End
   Begin VB.CommandButton cmdnew 
      Caption         =   "New"
      Height          =   375
      Left            =   0
      TabIndex        =   24
      Top             =   5400
      Width           =   855
   End
   Begin VB.CommandButton cmdsave 
      Caption         =   "Save"
      Height          =   375
      Left            =   0
      TabIndex        =   23
      Top             =   6000
      Width           =   855
   End
   Begin VB.CommandButton cmdclose 
      Caption         =   "Close"
      Height          =   375
      Left            =   0
      TabIndex        =   22
      Top             =   6600
      Width           =   855
   End
   Begin VB.CommandButton cmdnextitem 
      Caption         =   "Next item"
      Default         =   -1  'True
      Height          =   375
      Left            =   3960
      TabIndex        =   21
      Top             =   2520
      Width           =   1455
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      DrawStyle       =   2  'Dot
      DrawWidth       =   17015
      Height          =   360
      Left            =   4080
      Picture         =   "frmsalesreturn.frx":0000
      ScaleHeight     =   360
      ScaleWidth      =   240
      TabIndex        =   20
      Top             =   120
      Width           =   240
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      DrawStyle       =   2  'Dot
      DrawWidth       =   17015
      Height          =   360
      Left            =   4080
      Picture         =   "frmsalesreturn.frx":0182
      ScaleHeight     =   360
      ScaleWidth      =   240
      TabIndex        =   19
      Top             =   600
      Width           =   240
   End
   Begin VB.OptionButton optCash 
      Caption         =   "Cash"
      Height          =   255
      Left            =   0
      TabIndex        =   18
      Top             =   3240
      Value           =   -1  'True
      Width           =   1335
   End
   Begin VB.OptionButton lblCheckOff 
      Caption         =   "Check Off"
      Height          =   255
      Left            =   1440
      TabIndex        =   17
      Top             =   3240
      Width           =   1335
   End
   Begin VB.TextBox txtSNo 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3600
      TabIndex        =   16
      Top             =   3240
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.ComboBox cboproductname 
      Height          =   315
      Left            =   1560
      TabIndex        =   15
      Top             =   1080
      Width           =   4215
   End
   Begin VB.CheckBox chkPrint 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Print Receipt"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   0
      TabIndex        =   14
      Top             =   4440
      Value           =   2  'Grayed
      Width           =   1935
   End
   Begin VB.OptionButton opttransport 
      Caption         =   "Transporters"
      Height          =   255
      Left            =   7680
      TabIndex        =   13
      Top             =   1440
      Width           =   1335
   End
   Begin VB.TextBox txttranscode 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3840
      TabIndex        =   12
      Top             =   4680
      Width           =   1335
   End
   Begin VB.ComboBox ports 
      Height          =   315
      ItemData        =   "frmsalesreturn.frx":0304
      Left            =   840
      List            =   "frmsalesreturn.frx":0314
      TabIndex        =   10
      Text            =   "\\127.0.0.1\E-PoS 80mm Thermal Printer"
      Top             =   8040
      Width           =   3735
   End
   Begin VB.ComboBox Cmbstation 
      Height          =   315
      ItemData        =   "frmsalesreturn.frx":0330
      Left            =   7320
      List            =   "frmsalesreturn.frx":033A
      TabIndex        =   9
      Top             =   1800
      Width           =   1455
   End
   Begin VB.OptionButton Optbranch 
      Caption         =   "Station"
      Height          =   255
      Left            =   6000
      TabIndex        =   8
      Top             =   1440
      Width           =   1455
   End
   Begin VB.OptionButton optstaff 
      Caption         =   "Staff"
      Height          =   255
      Left            =   7200
      TabIndex        =   7
      Top             =   4080
      Width           =   1335
   End
   Begin VB.TextBox txtstaffno 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   9840
      TabIndex        =   6
      Top             =   4080
      Width           =   1095
   End
   Begin VB.CommandButton cmdsagroded 
      Caption         =   "Staff Agrovet Deductions"
      Height          =   375
      Left            =   0
      TabIndex        =   5
      Top             =   7200
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H008080FF&
      Caption         =   "Add New Product"
      Height          =   375
      Left            =   1680
      TabIndex        =   4
      Top             =   1560
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Remove"
      Height          =   495
      Left            =   6000
      TabIndex        =   3
      Top             =   4920
      Width           =   2175
   End
   Begin VB.TextBox txtamtreceived 
      Height          =   495
      Left            =   8520
      TabIndex        =   2
      Text            =   "0"
      Top             =   6960
      Width           =   2415
   End
   Begin VB.TextBox TXTCHANGE 
      Height          =   495
      Left            =   8520
      TabIndex        =   1
      Text            =   "0"
      Top             =   8160
      Width           =   2415
   End
   Begin VB.TextBox TXTTOTAL 
      Enabled         =   0   'False
      Height          =   495
      Left            =   8520
      TabIndex        =   0
      Text            =   "0"
      Top             =   6000
      Width           =   2535
   End
   Begin MSComCtl2.DTPicker DTPto 
      Height          =   255
      Left            =   7080
      TabIndex        =   11
      Top             =   2520
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   450
      _Version        =   393216
      Format          =   119799809
      CurrentDate     =   40588
   End
   Begin MSComCtl2.DTPicker txtransdate 
      Height          =   375
      Left            =   6480
      TabIndex        =   28
      Top             =   360
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   393216
      Format          =   119799809
      CurrentDate     =   40265
   End
   Begin MSComctlLib.ListView Lvwitems 
      Height          =   2535
      Left            =   2040
      TabIndex        =   30
      Top             =   5400
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   4471
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      MousePointer    =   4
      MouseIcon       =   "frmsalesreturn.frx":0353
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ITEM"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "QNTY"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "PRICE"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "AMOUNT"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Cash"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label19 
      Caption         =   "Comment"
      Height          =   255
      Left            =   9600
      TabIndex        =   63
      Top             =   4560
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "Receipt No."
      Height          =   255
      Left            =   120
      TabIndex        =   59
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Product Name"
      Height          =   255
      Left            =   120
      TabIndex        =   58
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Trans_Date"
      Height          =   255
      Left            =   6480
      TabIndex        =   57
      Top             =   0
      Width           =   1095
   End
   Begin VB.Label Label4 
      Caption         =   "Amount"
      Height          =   255
      Left            =   120
      TabIndex        =   56
      Top             =   2640
      Width           =   855
   End
   Begin VB.Label Label6 
      Caption         =   "Quantity"
      Height          =   255
      Left            =   120
      TabIndex        =   55
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label Label7 
      Caption         =   "Product Code"
      Height          =   255
      Left            =   120
      TabIndex        =   54
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label lblbalance 
      BackColor       =   &H00FF00FF&
      Height          =   255
      Left            =   4800
      TabIndex        =   53
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label Label9 
      Caption         =   "Balance"
      Height          =   255
      Left            =   3960
      TabIndex        =   52
      Top             =   2040
      Width           =   615
   End
   Begin VB.Label lblSNo 
      Caption         =   "SNo :"
      Height          =   255
      Left            =   3000
      TabIndex        =   51
      Top             =   3240
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label lblNetPay 
      BackColor       =   &H0000FF00&
      Caption         =   "NetPay:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2760
      TabIndex        =   50
      Top             =   4080
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label lblNPay 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00"
      DragMode        =   1  'Automatic
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   4080
      TabIndex        =   49
      Top             =   4080
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Label lblDed 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   7080
      TabIndex        =   48
      Top             =   2880
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label Label8 
      Caption         =   "Deductions :"
      Height          =   255
      Left            =   5880
      TabIndex        =   47
      Top             =   2880
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label lblGPay 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3360
      TabIndex        =   46
      Top             =   3600
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label Label11 
      Caption         =   "Gross Pay:"
      Height          =   255
      Left            =   2400
      TabIndex        =   45
      Top             =   3600
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label lblTKgs 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   960
      TabIndex        =   44
      Top             =   3600
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label13 
      Caption         =   "Total Kgs :"
      Height          =   255
      Left            =   0
      TabIndex        =   43
      Top             =   3600
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label lblSNames 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF00FF&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   4560
      TabIndex        =   42
      Top             =   3240
      Width           =   60
   End
   Begin VB.Label Label5 
      Caption         =   "Transport Code"
      Height          =   255
      Left            =   2640
      TabIndex        =   41
      Top             =   4680
      Width           =   1095
   End
   Begin VB.Label lbltransportername 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF00FF&
      Height          =   195
      Left            =   5400
      TabIndex        =   40
      Top             =   4680
      Width           =   165
   End
   Begin VB.Label Label10 
      Caption         =   "Net Pay:"
      Height          =   255
      Left            =   2640
      TabIndex        =   39
      Top             =   5040
      Width           =   855
   End
   Begin VB.Label lbltransnetpay 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3840
      TabIndex        =   38
      Top             =   5040
      Width           =   1695
   End
   Begin VB.Label Label16 
      Caption         =   "Period Ending"
      Height          =   255
      Left            =   5880
      TabIndex        =   37
      Top             =   2520
      Width           =   1095
   End
   Begin VB.Label Label12 
      Caption         =   "Printer Port"
      Height          =   375
      Left            =   0
      TabIndex        =   36
      Top             =   8040
      Width           =   855
   End
   Begin VB.Label lblstation 
      Caption         =   "Agrovet Branch"
      Height          =   255
      Left            =   6000
      TabIndex        =   35
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label Label14 
      Caption         =   "Staff No"
      Height          =   255
      Left            =   8880
      TabIndex        =   34
      Top             =   4080
      Width           =   735
   End
   Begin VB.Label Label15 
      Caption         =   "AMOUNT RECEIVED"
      Height          =   255
      Left            =   8280
      TabIndex        =   33
      Top             =   6600
      Width           =   1695
   End
   Begin VB.Label Label17 
      Caption         =   "CHANGE"
      Height          =   255
      Left            =   8280
      TabIndex        =   32
      Top             =   7800
      Width           =   975
   End
   Begin VB.Label Label18 
      Caption         =   "TOTAL"
      Height          =   255
      Left            =   8400
      TabIndex        =   31
      Top             =   5640
      Width           =   1575
   End
End
Attribute VB_Name = "frmsalesreturn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cboproductname_Change()
Set rst = oSaccoMaster.GetRecordset("select p_code from ag_products where p_name ='" & cboproductname & "'")
If Not rst.EOF Then
txtpcode = rst.Fields("p_code")
End If
End Sub

Private Sub cboproductname_Click()
Set rst = oSaccoMaster.GetRecordset("select p_code from ag_products where p_name ='" & cboproductname & "'")
If Not rst.EOF Then
txtpcode = rst.Fields("p_code")
End If
End Sub

Private Sub cboproductname_Validate(Cancel As Boolean)
cmdNew_Click

Provider = cn
Set cn = New ADODB.Connection
Dim p As Integer
'cn.Open Provider, "bi"
'If reportpath = "" Then reportpath = GetSetting("payroll", "AppName", "rptPath", rptPath)
'Dim rst As New ADODB.Recordset
sql = ""
'SELECT p_code, p_name, S_No, Qout, sprice FROM   ag_Products
'sql = "select p_code, S_No, Qout, sprice from ag_products where p_name='" & cboproductname & "'"
sql = "select p_code, S_No, Qout, sprice from ag_products where p_name='" & cboproductname & "' and p_code='" & txtpcode & "'"
Set rs = New ADODB.Recordset
Set rs = oSaccoMaster.GetRecordset(sql)
If Not rs.EOF Then
txtpcode = rs.Fields(0)
lblbalance = rs.Fields(2)
'txtserialno = rs.Fields(1)
txtamount = rs.Fields(3)

End If
End Sub

Private Sub cmdclose_Click()
Unload Me
End Sub

Private Sub cmdNew_Click()
'Set rs = oSaccoMaster.GetRecordset("d_sp_NextReceipt")
Set rs = oSaccoMaster.GetRecordset("select rcpno from rcpno")
If Not (rs.EOF) Then
txtrno = rs.Fields(0) + 1
Else
txtrno = 1
End If

 'txtpcode = ""
 'txtserialno = ""
 txtquantity = 0
 txtamount = 0
 txtamtreceived = 0
 TXTCHANGE = 0
 TXTTOTAL = 0
End Sub

Private Sub cmdnextitem_Click()
Dim cash As Integer
Dim total As Double
'check the user
sql = "SELECT     UserLoginID,levels, UserGroup, SUPERUSER From UserAccounts where UserLoginID='" & User & "'"
Set rs = oSaccoMaster.GetRecordset(sql)
If Not rs.EOF Then
If rs!Levels <> "Agrovet" Then
MsgBox "You are not allowed to sell", vbInformation
Exit Sub

End If
End If
If Trim(txtquantity) = "" Then
MsgBox "Quantity cannot be Zero", vbInformation
Exit Sub
End If

    If Optbranch = True Then
    If Trim(Cmbstation.Text) = "" Then
        MsgBox "Please enter the Agrovet Station."
            Cmbstation.SetFocus
    Exit Sub
    End If
    
    If optrecieve = True Then
    If Trim(Cmbstation.Text) = "" Then
        MsgBox "Please enter the Agrovet Station."
            Cmbstation.SetFocus
    Exit Sub
    End If
    End If
    Set rst = oSaccoMaster.GetRecordset("select pprice from ag_products where p_code='" & txtpcode & "'")
    If Not rst.EOF Then
    'txtAmount = Rst.Fields("pprice")
    End If
    End If
    
    
    
    If opttransport = True Then
    If Trim(txttranscode) = "" Then
        MsgBox "Please enter the Transporter"
    
    Exit Sub
    End If
    End If
    
    If txtpcode = "" Then
        MsgBox "Please Enter the Product CODE before You Proceed!", vbCritical
        Exit Sub
    End If
    If txtrno = "" Then
        MsgBox "Please Enter Receipt Number before you Proceed!", vbCritical
        Exit Sub
    End If
    
If txtamount = "" Then
txtamount = 0
End If
Provider = "maziwa"
Set cn = New ADODB.Connection
cn.Open Provider, "bi"
'// check if they are in stock.
Dim rsinstock As Recordset
sql = ""
sql = "select P_CODE,qout from ag_products where p_code='" & txtpcode & "'"
Set rsinstock = New ADODB.Recordset
rsinstock.Open sql, cn

'// check the quanttity being sold versus the balance
Dim piu As Double
piu = rsinstock.Fields(1) + CInt(txtquantity)



If optCash.value = True Then
cash = 1
Else
cash = 0
End If

Dim j, Coun As Integer
j = 1




'Check if same item is in the list
   Do While Not j > (Coun)
         Lvwitems.ListItems.Item(j).selected = True
            
    If Lvwitems.SelectedItem = txtpcode Then
        txtquantity = (CCur(txtquantity) + CCur(Lvwitems.SelectedItem.ListSubItems(2)))
        Lvwitems.ListItems.Remove (Lvwitems.SelectedItem.index)
                        
        Set li = Lvwitems.ListItems.Add(, , txtpcode)
                        li.SubItems(1) = cboproductname & ""
                        li.SubItems(2) = txtquantity & ""
                        li.SubItems(3) = txtamount & ""
                        li.SubItems(4) = CCur(txtamount) * CCur(txtquantity) & ""
                        li.SubItems(5) = cash
                        'Total = CCur(Total + li.SubItems(4))
                        TXTTOTAL = total
                                                
        j = Coun + 1
        
        lblbalance = CCur(lblbalance) - CCur(txtquantity)

        txtpcode = ""
        txtquantity = ""
       ' txtserialno = ""
        txtpcode.SetFocus
        Exit Sub
         
    
   
'   lvwItems.ListItems.Item(J).selected = True
   End If
   j = j + 1
    Loop
    
     If j > 1 Then
   
    Set li = Lvwitems.ListItems.Add(, , txtpcode)
                        li.SubItems(1) = cboproductname & ""
                        li.SubItems(2) = txtquantity & ""
                        li.SubItems(3) = txtamount & ""
                        li.SubItems(4) = CCur(txtamount) * (CCur(txtquantity)) & ""
                        li.SubItems(5) = cash
                        'Total = CCur(Total + li.SubItems(4))
                        TXTTOTAL = total
                        
        lblbalance = CCur(lblbalance) - CCur(txtquantity)
        txtpcode = ""
        txtquantity = ""
        'txtserialno = ""
        txtpcode.SetFocus
        Exit Sub
    End If
     If Coun = 0 Then
     Set li = Lvwitems.ListItems.Add(, , txtpcode)
                        li.SubItems(1) = cboproductname & ""
                        li.SubItems(2) = txtquantity & ""
                        li.SubItems(3) = txtamount & ""
                        li.SubItems(4) = CCur(txtamount) * (CCur(txtquantity)) & ""
                        li.SubItems(5) = cash
                        'Total = CCur(Total + li.SubItems(4))
                        TXTTOTAL = total
    End If

lblbalance = CCur(lblbalance) + CCur(txtquantity)
TXTTOTAL = 0
'Coun = Lvwitems.ListItems.Count
'For j = 1 To Lvwitems.ListItems.Count
'    Total = CCur(Total + li.SubItems(4))
'    txttotal = Total
'
'Next j
Do While Not j > (Lvwitems.ListItems.Count)
'For j = 1 To Lvwitems.ListItems.Count
 Lvwitems.ListItems.Item(j).selected = True
 total = total + CCur(Lvwitems.SelectedItem.SubItems(4))
 TXTTOTAL = total
j = j + 1
Loop

txtpcode = ""
txtquantity = ""
'txtserialno = ""
txtpcode.SetFocus

End Sub

Private Sub cmdsalesreturn_Click()

Unload Me
frmreceipts.Show vbModal
'Unload Me
End Sub

Private Sub cmdsave_Click()
On Error GoTo HEREEE

If opttransport = True Then
savetransporters
Exit Sub
End If

If Optbranch = True Then
savestation
Exit Sub
End If

If optrecieve = True Then
ReceiveFromStaion
Exit Sub
End If


If lblCheckOff = True Then
   If txtSNo = "" Then
    MsgBox "Enter the SupplierNo  ", vbInformation, "CheckOff"
     Exit Sub
    End If
savesno
Exit Sub
End If

If optCash = True Then
    If TXTCHANGE < 0 Then
         If MsgBox("Insufficient Amount Received,do you want to transfer balance to check off ", vbYesNo) = vbYes Then
            lblCheckOff_Click
            lblCheckOff.value = True
            optCash.value = False
         Exit Sub
        Else
       Exit Sub
       End If
    End If
    savecash
   Exit Sub
End If
If optstaff = True Then
savestaff
Exit Sub
End If
If Optothers = True Then
saveothers
Exit Sub
End If
HEREEE:
MsgBox err.description & " error occured."
End Sub
Sub savestation()
On Error GoTo kiparu

If Optbranch = True Then
Dim centre As String
centre = Cmbstation.Text
If Trim(Cmbstation.Text) = "" Then
 MsgBox "Please enter the Agrovet Station."
   Cmbstation.SetFocus
Exit Sub
End If


Dim j As Integer
If Lvwitems.ListItems.Count = 0 Then
MsgBox "There are no items sold."
Exit Sub
End If
j = 1

Dim total As Currency
total = 0
Do While Not j > (Lvwitems.ListItems.Count)
 Lvwitems.ListItems.Item(j).selected = True
 total = total + CCur(Lvwitems.SelectedItem.SubItems(4))
j = j + 1
Loop


Startdate = DateSerial(year(txtransdate), month(txtransdate), 1)
Enddate = DateSerial(year(txtransdate), month(txtransdate) + 1, 1 - 1)


Set rs = oSaccoMaster.GetRecordset("d_sp_IsClosed '" & Enddate & "'")
If Not rs.EOF Then
    MsgBox "The period " & Enddate & " has been closed by " & rs.Fields(0)
    Exit Sub
End If
'End If
'End If
For j = 1 To Lvwitems.ListItems.Count
'Do While Not j > (Lvwitems.ListItems.Count)
 Lvwitems.ListItems.Item(j).selected = True



Dim rsinstock As Recordset
sql = ""
sql = "select P_CODE,qout,Pprice from ag_products where p_code='" & Lvwitems.SelectedItem & "'"

Set rsinstock = oSaccoMaster.GetRecordset(sql)

'// check the quanttity being sold versus the balance
Dim Remain, Pprice, amount As Double
Remain = rsinstock.Fields(1) + CInt(Lvwitems.SelectedItem.SubItems(2))
Pprice = IIf(IsNull(rsinstock.Fields(2)), CInt(Lvwitems.SelectedItem.SubItems(3)), rsinstock.Fields(2))
amount = CInt(Lvwitems.SelectedItem.SubItems(3)) * CInt(Lvwitems.SelectedItem.SubItems(2))

sql = ""
sql = sql & "SET dateformat DMY INSERT INTO ag_Receipts1(R_No, P_code, T_Date, Amount, Qua, S_Bal, user_id, Cash, SNo,Remarks) VALUES ("
sql = sql & txtrno & ",'" & Lvwitems.SelectedItem & "','" & txtransdate & "'," & -1 * amount
sql = sql & "," & Lvwitems.SelectedItem.SubItems(2) & "," & Remain & ",'" & User & "'," & Lvwitems.SelectedItem.SubItems(5) & ",'" & Cmbstation & "','Sales Return')"

oSaccoMaster.ExecuteThis (sql)
oSaccoMaster.ExecuteThis ("Update ag_Products SET Qout =" & CCur(Remain) & " WHERE p_code= '" & Lvwitems.SelectedItem & "'")

'********************************************************************************
pcode = Lvwitems.SelectedItem
'sql = ""
'sql = "set dateformat DMY update ag_products set qin=" & Lvwitems.ListItems(j).SubItems(2) & ",qout=" & Remain & ",o_bal=" & Remain & ",last_d_updated='" & txtransdate & "',user_id='" & User & "',audit_date='" & Get_Server_Date & "',unserialized=0,SERIA=0,pprice=" & rs.Fields("pprice") & ",sprice=" & rs.Fields("sprice") & " where p_code='" & pcode & "'"
'cn.Execute sql
'
'Dim rsst As Recordset
'sql = ""
'sql = "set dateformat DMY select top 1 * from ag_stockbalance where p_code='" & pcode & "' order by trackid desc "
'Set rsst = New ADODB.Recordset
'rsst.Open sql, cn
'If Not rsst.EOF Then
'sql = ""
'sql = "set dateformat DMY INSERT INTO ag_stockbalance"
'sql = sql & " (p_code, productname, openningstock, changeinstock, stockbalance, transdate,companyid)"
'sql = sql & " VALUES     ('" & pcode & "', '" & Lvwitems.ListItems(j).SubItems(1) & "', '" & rs.Fields("o_bal") & "', '" & Lvwitems.ListItems(j).SubItems(2) & "', '" & Remain & "', '" & Format(Get_Server_Date, "dd/mm/yyyy") & "',1)"
'cn.Execute sql
'End If

Dim DRaccno As String
Dim Craccno As String
If centre = "SANGALO" Then
    Craccno = "A008"
    DRaccno = "I004"
ElseIf centre = "OLMAROROI" Then
    Craccno = "A010"
    DRaccno = "I005"
ElseIf centre = "KABISAGA" Then
    Craccno = "A012"
    DRaccno = "I006"
ElseIf centre = "KOISOLIK" Then
    Craccno = "A009"
    DRaccno = "I007"
ElseIf centre = "CHEMUSWO" Then
    Craccno = "A011"
    DRaccno = "I008"
ElseIf centre = "BELEKENYA" Then
    Craccno = "A013"
End If
'XXXXXXXXXXX SAVE TO GL
    sql = ""
    sql = "set dateformat dmy insert into gltransactions(transdate,amount,draccno,craccno,documentno,source,transdescript,auditid,cash,doc_posted) values('" & txtransdate & "'," & amount & ",'" & DRaccno & "','" & Craccno & "','" & Lvwitems.SelectedItem & "','" & cboproductname & "' ,'SALES RETURN ','" & User & "',0,0)"
    oSaccoMaster.ExecuteThis (sql)


'XXXXXXXXXXXXXXXXXXXXXX


Next j



If chkPrint.value = vbChecked Then
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
             Dim fso, chkPrinter, txtFile
        'ttt = "LPT1" 'LPT1
         Dim PORT As String
        PORT = ports
        'ttt = "LPT1" 'LPT1
        ttt = PORT
        Set fso = CreateObject("Scripting.FileSystemObject")
        Dim strReceipts As String
        Set txtFile = fso.CreateTextFile(ttt, True)
    
        txtFile.WriteLine "  " & cname & ""
        txtFile.WriteLine "     " & paddress & ""
        txtFile.WriteLine "     " & town & ""
        txtFile.WriteLine "     " & Phone & ""
        txtFile.WriteLine "     " & Email & ""
        txtFile.WriteLine "---------------------------------------"
        txtFile.WriteLine "    AGROVET RECEIPT"
        txtFile.WriteLine "  STOCK RETURNED FROM " & centre & ""
        txtFile.WriteLine "---------------------------------------"
        txtFile.WriteLine "QNTY" & vbTab & vbTab & "PRICE" & vbTab & vbTab & "AMOUNT"
        txtFile.WriteLine "........................................"
        j = 1
        strReceipts = ""
        Do While Not j > (Lvwitems.ListItems.Count)
            Lvwitems.ListItems.Item(j).selected = True
            strReceipts = strReceipts & Lvwitems.SelectedItem.SubItems(1) & vbNewLine & Lvwitems.SelectedItem.SubItems(2) & vbTab & vbTab
            strReceipts = strReceipts & Format(Lvwitems.SelectedItem.SubItems(3), "#,##0.00") & vbTab & vbTab & Format(Lvwitems.SelectedItem.SubItems(4), "#,##0.00") & vbNewLine
            txtFile.WriteLine strReceipts
            j = j + 1
        Loop
    txtFile.WriteLine "TOTAL" & vbTab & vbTab & vbTab & vbTab & Format(total, "#,##0.00") & vbNewLine
    txtFile.WriteLine
    txtFile.WriteLine
    txtFile.WriteLine "YOU WERE SERVED By " & UCase(username)
    txtFile.WriteLine "    Date :" & Format(Get_Server_Date, "dd/mm/yyyy HH:MM:SS AM/PM")
    txtFile.WriteLine " Stock Returned From " & centre & "at the selling price"
    txtFile.WriteLine "     THANK YOU AND WELCOME "
    txtFile.WriteLine "****************************************"
    txtFile.WriteLine escFeedAndCut
    txtFile.Close
    End If
End If

Lvwitems.ListItems.Clear
txtrno = ""
txtpcode.Text = ""
txtquantity = ""
txtamount = ""
Cmbstation.Text = ""

MsgBox "Record saved Successfully"
Exit Sub
kiparu:
MsgBox err.description & " error occured."
End Sub

Private Sub Command2_Click()
Dim total As Double
Dim j, Coun As Integer
j = 1
On Error GoTo ErrorHandler
TXTTOTAL = 0
'If Lvwitems.ListItems.Count > 0 Then
''Total = CCur(txttotal - li.SubItems(4))
Lvwitems.ListItems.Remove (Lvwitems.SelectedItem.index)  '// removes the selected item

Do While Not j > (Lvwitems.ListItems.Count)
'For j = 1 To Lvwitems.ListItems.Count
 Lvwitems.ListItems.Item(j).selected = True
 total = total + CCur(Lvwitems.SelectedItem.SubItems(4))
 TXTTOTAL = total
j = j + 1
Loop

'End If
Exit Sub
ErrorHandler:
MsgBox err.description
End Sub

Private Sub Form_Load()
Label5.Visible = False
txttranscode.Visible = False
lbltransportername.Visible = False
Label10.Visible = False
lbltransnetpay.Visible = False
txtransdate = Format(Date, "dd/mm/yyyy")

Provider = "MAZIWA"
Set cn = New ADODB.Connection
cn.Open Provider, "bi"
'If reportpath = "" Then reportpath = GetSetting("payroll", "AppName", "rptPath", rptPath)
sql = "select P_NAME  from ag_products ORDER BY P_NAME ASC"
Set rs = New ADODB.Recordset
rs.Open sql, cn

While Not rs.EOF
cboproductname.AddItem rs.Fields(0)
rs.MoveNext
Wend
cboproductname.Enabled = True
chkPrint.value = vbChecked
End Sub

Private Sub lblCheckOff_Click()
lblSNo.Visible = True
txtSNo.Visible = True
lblNetPay.Visible = True
lblNPay.Visible = True
lblDed.Visible = True
lblTKgs.Visible = True
lblGPay.Visible = True
Label11.Visible = True
Label13.Visible = True
Label8.Visible = True
txttranscode.Visible = False
Label5.Visible = False
Label10.Visible = False
lbltransnetpay.Visible = False
lbltransportername.Visible = False
End Sub

Private Sub Optbranch_Click()
lblSNo.Visible = False
txtSNo.Visible = False
Label13.Visible = False
lblTKgs.Visible = False
Label11.Visible = False
lblGPay.Visible = False
Label8.Visible = False
lblDed.Visible = False
lblNetPay.Visible = False
lblNPay.Visible = False
Label5.Visible = False
txttranscode.Visible = False
Label10.Visible = False
lbltransnetpay.Visible = False
lbltransportername.Visible = False

End Sub

Private Sub Optcash_Click()
lblSNo.Visible = False
txtSNo.Visible = False
lblNetPay.Visible = False
lblNPay.Visible = False

lblDed.Visible = False
lblTKgs.Visible = False
lblGPay.Visible = False
Label11.Visible = False
Label13.Visible = False
Label8.Visible = False

End Sub

Private Sub Optothers_Click()
lblSNo.Visible = False
txtSNo.Visible = False
lblNetPay.Visible = False
lblNPay.Visible = False

lblDed.Visible = False
lblTKgs.Visible = False
lblGPay.Visible = False
Label11.Visible = False
Label13.Visible = False
Label8.Visible = False
End Sub

Private Sub optrecieve_Click()
lblSNo.Visible = False
txtSNo.Visible = False
Label13.Visible = False
lblTKgs.Visible = False
Label11.Visible = False
lblGPay.Visible = False
Label8.Visible = False
lblDed.Visible = False
lblNetPay.Visible = False
lblNPay.Visible = False
Label5.Visible = False
txttranscode.Visible = False
Label10.Visible = False
lbltransnetpay.Visible = False
lbltransportername.Visible = False
End Sub

Private Sub optstaff_Click()
lblSNo.Visible = False
txtSNo.Visible = False
lblNetPay.Visible = False
lblNPay.Visible = False

lblDed.Visible = False
lblTKgs.Visible = False
lblGPay.Visible = False
Label11.Visible = False
Label13.Visible = False
Label8.Visible = False
End Sub

Private Sub opttransport_Click()
If opttransport = True Then
Label5.Visible = True
txttranscode.Visible = True
lbltransportername.Visible = True
Label10.Visible = True
lbltransnetpay.Visible = True
lblSNames.Visible = False

lblSNo.Visible = False
txtSNo.Visible = False
Label13.Visible = False
lblTKgs.Visible = False
Label11.Visible = False
lblGPay.Visible = False
Label8.Visible = False
lblDed.Visible = False
lblNetPay.Visible = False
lblNPay.Visible = False
lblSNames.Visible = False

Else
Label5.Visible = False
txttranscode.Visible = False
lbltransportername.Visible = False
Label10.Visible = False
lbltransnetpay.Visible = False
End If

End Sub

Private Sub Picture1_Click()
frmSearch.Show vbModal
Dim Y As String
Y = sel
Dim p As Integer
If Y <> "" Then
'Provider = cn
Set cn = New ADODB.Connection
'cn.Open Provider, "bi"
'If reportpath = "" Then reportpath = GetSetting("payroll", "AppName", "rptPath", rptPath)
sql = "select P_CODE,P_NAME,S_NO,QOUT,seria,s_no from ag_products where p_code='" & Y & "'"
Set rs = New ADODB.Recordset
Set rs = oSaccoMaster.GetRecordset(sql)
If Not rs.EOF Then
If Not IsNull(rs.Fields(0)) Then txtpcode.Text = (rs.Fields(0))
If Not IsNull(rs.Fields(4)) Then p = (rs.Fields(4))
If p = 1 Then
If Not IsNull(rs.Fields(5)) Then 'txtserialno = (rs.Fields(5))
'lblserialno.Visible = True
'txtserialno.Visible = True
Else
'lblserialno.Visible = False
'txtserialno.Visible = False
End If
End If

If Not IsNull(rs.Fields(1)) Then cboproductname = (rs.Fields(1))
If Not IsNull(rs.Fields(0)) Then txtpcode = (rs.Fields(0))

'If Not IsNull(rs.Fields(3)) Then lblbalance = (rs.Fields(3))
'// check if it has the serial numbers
'get_serialno Y
End If

'// check if the product have the serial then show the ag_receipts details
cboproductname_Validate True

End If
End Sub

Private Sub txtpcode_Change()

If KeyAscii = 13 Then
Provider = "MAZIWA"
Set cn = New ADODB.Connection
cn.Open Provider, "bi"
sql = "select P_CODE,P_NAME,S_NO,QOUT,supplierid from ag_products where p_code='" & txtpcode & "'"
Set rs = New ADODB.Recordset
rs.Open sql, cn
If Not rs.EOF Then
 
If Not IsNull(rs.Fields(1)) Then cboproductname = (rs.Fields(1))



End If
End If

End Sub
Private Sub txtquantity_Validate(Cancel As Boolean)
If Not IsNumeric(txtquantity) Then
MsgBox "Enter a value please", vbCritical
txtquantity = ""
txtquantity.SetFocus
Exit Sub
End If
End Sub

Private Sub txtSNo_Change()
On Error GoTo ErrorHandler
Set rs = New ADODB.Recordset
sql = "d_sp_SelectSuppliers '" & txtSNo & "'"
Set rs = oSaccoMaster.GetRecordset(sql)
If Not rs.EOF Then
If Not IsNull(rs.Fields(0)) Then lblSNames = rs.Fields(2)
Else
lblSNames = ""
End If

Startdate = DateSerial(year(txtransdate), month(txtransdate), 1)
Enddate = DateSerial(year(txtransdate), month(txtransdate) + 1, 1 - 1)

Set rs = oSaccoMaster.GetRecordset("d_sp_SupNet " & txtSNo & ",'" & Startdate & "','" & Enddate & "', 0")

If Not rs.EOF Then
lblTKgs = IIf(IsNull(rs.Fields(0)), 0, rs.Fields(0))
lblGPay = IIf(IsNull(rs.Fields(1)), 0, rs.Fields(1))
Else
lblTKgs = "0.00"
End If



Set rs = oSaccoMaster.GetRecordset("d_sp_SupNet " & txtSNo & ",'" & Startdate & "','" & Enddate & "', 1")
If Not IsNull(rs.Fields(0)) Then
lblDed = rs.Fields(0)
Else
lblDed = "0.00"
End If

lblNPay = Format((CCur(lblGPay) - CCur(lblDed)), "#,##0.00")

Exit Sub
ErrorHandler:
MsgBox err.description

End Sub
Private Sub txttranscode_Change()
Set rs = New ADODB.Recordset
Dim DTPfrom As Date
sql = "d_sp_TransEnquiry  '" & txttranscode & "'"
Set rs = oSaccoMaster.GetRecordset(sql)
If Not rs.EOF Then
If Not IsNull(rs.Fields(0)) Then lbltransportername = rs.Fields(0)
End If
DTPfrom = DateSerial(year(txtransdate), month(txtransdate), 1)
DTPto = DateSerial(year(txtransdate), month(txtransdate) + 1, 1 - 1)

' get transporter netpay
   Dim mMonth, yyear As Integer
   mMonth = month(txtransdate)
   yyear = year(txtransdate)
   
  sql = " Select(Select isnull(SUM(Amount + Subsidy),0) from d_TransDetailed where Trans_Code='" & txttranscode & "' and MMonth= " & mMonth & " and YYear=" & yyear & "),"
  sql = sql & " (Select isnull(SUM(Amount),0) from d_Transport_Deduc where TransCode='" & txttranscode & "' and MONTH(TDate_Deduc)=" & mMonth & " and YEAR(TDate_Deduc)= " & yyear & ")"
   Set rs2 = oSaccoMaster.GetRecordset(sql)
   If Not rs2.EOF Then
   lbltransnetpay = Format(rs2.Fields(0) - rs2.Fields(1), Cfmt)
  
   Else
   lbltransnetpay = "0.00"
   End If
End Sub
Private Sub savetransporters()
On Error GoTo kiparu2

If opttransport = True Then
    If txttranscode = "" Then
      MsgBox "Please enter the transporter"
    Exit Sub
End If

Set rst = New Recordset
Dim a, b, X
DIA = 0
Dim U As Double, S As Double
Dim cn As Connection
Dim j As Integer

If DIA = 1 Then
Exit Sub
End If
If Lvwitems.ListItems.Count = 0 Then
MsgBox "There are no items sold."
Exit Sub
End If
j = 1

Dim total As Currency
total = 0
Do While Not j > (Lvwitems.ListItems.Count)
 Lvwitems.ListItems.Item(j).selected = True
 total = total + CCur(Lvwitems.SelectedItem.SubItems(3))
j = j + 1
Loop

If optCash.value = False Then
   If total > CCur(lbltransnetpay) Then
         If MsgBox("Transporter number " & txttranscode & " has a netpay of " & lblNPay & " do you wish to proceed?", vbYesNo) = vbYes Then
     Else
   Exit Sub
  End If
End If


Startdate = DateSerial(year(txtransdate), month(txtransdate), 1)
Enddate = DateSerial(year(txtransdate), month(txtransdate) + 1, 1 - 1)


Set rs = oSaccoMaster.GetRecordset("d_sp_IsClosed '" & Enddate & "'")
If Not rs.EOF Then
    MsgBox "The period " & Enddate & " has been closed by " & rs.Fields(0)
    Exit Sub
End If
'End If
End If
j = 1
For j = 1 To Lvwitems.ListItems.Count
'Do While Not j > (Lvwitems.ListItems.Count)
 Lvwitems.ListItems.Item(j).selected = True
If Trim(txttranscode) = "" Then
txttranscode = "0"
End If
'// check if they are in stock.

Dim rsinstock As Recordset
sql = ""
sql = "select P_CODE,qout from ag_products where p_code='" & Lvwitems.SelectedItem & "'"

Set rsinstock = oSaccoMaster.GetRecordset(sql)


Dim Remain As Double
Remain = rsinstock.Fields(1) + CInt(Lvwitems.SelectedItem.SubItems(2))


sql = ""
sql = sql & "SET dateformat DMY INSERT INTO ag_Receipts(R_No, P_code, T_Date, Amount, Qua, S_Bal, user_id, Cash, SNo,Branch) VALUES ("
sql = sql & txtrno & ",'" & Lvwitems.SelectedItem & "','" & txtransdate & "'," & -1 * Lvwitems.SelectedItem.SubItems(4)
sql = sql & "," & Lvwitems.SelectedItem.SubItems(2) & "," & Remain & ",'" & User & "'," & Lvwitems.SelectedItem.SubItems(5) & ",'" & txttranscode & "','" & Cmbstation & "')"

oSaccoMaster.ExecuteThis (sql)
sql = ""
sql = sql & "SET dateformat DMY INSERT INTO ag_Receipts3(R_No, P_code, T_Date, Amount, Qua, S_Bal, user_id, Cash, SNo) VALUES ("
sql = sql & txtrno & ",'" & Lvwitems.SelectedItem & "','" & txtransdate & "'," & -1 * Lvwitems.SelectedItem.SubItems(4)
sql = sql & "," & Lvwitems.SelectedItem.SubItems(2) & "," & Remain & ",'" & User & "'," & Lvwitems.SelectedItem.SubItems(5) & ",'" & txttranscode & "')"

oSaccoMaster.ExecuteThis (sql)
oSaccoMaster.ExecuteThis ("Update ag_Products SET Qout =" & CCur(Remain) & " WHERE p_code= '" & Lvwitems.SelectedItem & "'")
'j = j + 1
'Loop
    sql = ""
    sql = "set dateformat dmy insert into gltransactions(transdate,amount,draccno,craccno,documentno,source,transdescript,auditid,cash,doc_posted) values('" & txtransdate & "'," & Lvwitems.SelectedItem.SubItems(4) & ",'I005','A007','" & Lvwitems.SelectedItem & "','" & cboproductname & "' ,'TRANSPORTERS SALES','" & User & "',0,0)"
    oSaccoMaster.ExecuteThis (sql)

Next j
'//Update deductions
If optCash.value = False Then
Set cn = New ADODB.Connection
sql = "d_sp_TransDeduct '" & txttranscode & "','" & txtransdate & "','Agrovet'," & -1 * total & ",'" & Startdate & "','" & Enddate & "','" & User & "','" & txtcomment & "'"
oSaccoMaster.ExecuteThis (sql)
End If

If CDbl(txtamtreceived) > 0 Then
amount = 0
amount = CDbl(txtamtreceived) * -1
Set cn = New ADODB.Connection
sql = "d_sp_TransDeduct '" & txttranscode & "','" & txtransdate & "','Agrovet'," & amount & ",'" & Startdate & "','" & Enddate & "','" & User & "','Cash'"
oSaccoMaster.ExecuteThis (sql)
End If

If chkPrint.value = vbChecked Then
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
       Dim fso, chkPrinter, txtFile
        Dim PORT As String
        PORT = ports
        ttt = PORT 'LPT1
        
        Set fso = CreateObject("Scripting.FileSystemObject")
        Dim strReceipts As String
        j = 1
        strReceipts = ""
        Do While Not j > (Lvwitems.ListItems.Count)
            Lvwitems.ListItems.Item(j).selected = True
            strReceipts = strReceipts & Lvwitems.SelectedItem.SubItems(1) & vbNewLine & Lvwitems.SelectedItem.SubItems(2) & vbTab & vbTab
            strReceipts = strReceipts & Format(Lvwitems.SelectedItem.SubItems(3), "#,##0.00") & vbTab & vbTab & Format(Lvwitems.SelectedItem.SubItems(4), "#,##0.00") & vbNewLine
            j = j + 1
        Loop

        'MsgBox strReceipts
        strReceipts = strReceipts & vbNewLine & "---------------------------------------" & vbNewLine
        strReceipts = strReceipts & "TOTAL" & vbTab & vbTab & vbTab & vbTab & Format(total, "#,##0.00") & vbNewLine
        strReceipts = strReceipts & "======================================="
                 
        Set txtFile = fso.CreateTextFile(ttt, True)
    txtFile.WriteLine "      " & cname & ""
    txtFile.WriteLine "       " & paddress & ""
    txtFile.WriteLine "---------------------------------------"
    txtFile.WriteLine "      AGROVET RECEIPT"
    txtFile.WriteLine "          Check-off"
    txtFile.WriteLine "---------------------------------------"
    txtFile.WriteLine "TCode:" & txttranscode
    txtFile.WriteLine "Name:" & lbltransportername
    
'NAME QNTY PRICE AMOUNT
    txtFile.WriteLine "QNTY" & vbTab & vbTab & "PRICE" & vbTab & vbTab & "AMOUNT"
    txtFile.WriteLine "......................................."
    txtFile.WriteLine strReceipts
    txtFile.WriteLine
    txtFile.WriteLine "TOTAL" & TXTTOTAL
    txtFile.WriteLine
    txtFile.WriteLine "AMOUNT RECEVED" & txtamtreceived
    txtFile.WriteLine
    txtFile.WriteLine "CHANGE" & TXTCHANGE
    txtFile.WriteLine
    txtFile.WriteLine "YOU WERE SERVED By " & UCase(username)
    txtFile.WriteLine "    Date :" & Format(Get_Server_Date, "dd/mm/yyyy HH:MM:SS AM/PM")
    txtFile.WriteLine "     THANK YOU AND WELCOME "
    txtFile.WriteLine "****************************************"
    txtFile.WriteLine escFeedAndCut
    txtFile.Close
End If
End If


Lvwitems.ListItems.Clear
txttranscode = ""
txtrno.Text = ""
txtpcode.Text = ""
'txtserialno = ""
lbltransnetpay = ""
txtquantity = ""
txtamount = ""

MsgBox "Records saved"
Exit Sub
kiparu2:
MsgBox err.description & " error occured."

End Sub
Sub savecash()
On Error GoTo olkalou

If optCash = True Then
Dim C As String
C = "Cash"

Dim j As Integer
If Lvwitems.ListItems.Count = 0 Then
MsgBox "There are no items sold."
Exit Sub
End If
j = 1

Dim total As Currency
total = 0
Do While Not j > (Lvwitems.ListItems.Count)
 Lvwitems.ListItems.Item(j).selected = True
 total = total + CCur(Lvwitems.SelectedItem.SubItems(4))
j = j + 1
Loop



Startdate = DateSerial(year(txtransdate), month(txtransdate), 1)
Enddate = DateSerial(year(txtransdate), month(txtransdate) + 1, 1 - 1)

Set rs = oSaccoMaster.GetRecordset("d_sp_IsClosed '" & Enddate & "'")
If Not rs.EOF Then
    MsgBox "The period " & Enddate & " has been closed by " & rs.Fields(0)
    Exit Sub
End If



For j = 1 To Lvwitems.ListItems.Count
 Lvwitems.ListItems.Item(j).selected = True


Dim rsinstock As Recordset
sql = ""
sql = "select P_CODE,qout from ag_products where p_code='" & Lvwitems.SelectedItem & "'"

Set rsinstock = oSaccoMaster.GetRecordset(sql)

Dim Remain As Double
Remain = rsinstock.Fields(1) + CInt(Lvwitems.SelectedItem.SubItems(2))


sql = ""
sql = sql & "SET dateformat DMY INSERT INTO ag_Receipts(R_No, P_code, T_Date, Amount, Qua, S_Bal, user_id, Cash, SNo,Remarks,Branch) VALUES ("
sql = sql & txtrno & ",'" & Lvwitems.SelectedItem & "','" & txtransdate & "'," & -1 * Lvwitems.SelectedItem.SubItems(4)
sql = sql & "," & Lvwitems.SelectedItem.SubItems(2) & "," & Remain & ",'" & User & "'," & Lvwitems.SelectedItem.SubItems(5) & ",'" & C & "','Sales Return','" & Cmbstation & "')"

oSaccoMaster.ExecuteThis (sql)
sql = ""
sql = sql & "SET dateformat DMY INSERT INTO ag_Receipts3(R_No, P_code, T_Date, Amount, Qua, S_Bal, user_id, Cash, SNo,Remarks) VALUES ("
sql = sql & txtrno & ",'" & Lvwitems.SelectedItem & "','" & txtransdate & "'," & -1 * Lvwitems.SelectedItem.SubItems(4)
sql = sql & "," & Lvwitems.SelectedItem.SubItems(2) & "," & Remain & ",'" & User & "'," & Lvwitems.SelectedItem.SubItems(5) & ",'" & C & "','Sales Return')"

oSaccoMaster.ExecuteThis (sql)
oSaccoMaster.ExecuteThis ("Update ag_Products SET Qout =" & CCur(Remain) & " WHERE p_code= '" & Lvwitems.SelectedItem & "'")

'\\ save to gl


    sql = ""
    sql = "set dateformat dmy insert into gltransactions(transdate,amount,draccno,craccno,documentno,source,transdescript,auditid,cash,doc_posted) values('" & txtransdate & "'," & Lvwitems.SelectedItem.SubItems(4) & ",'I001','A004','" & Lvwitems.SelectedItem & "','" & cboproductname & "' ,'Sales Return','" & User & "',1,0)"
    oSaccoMaster.ExecuteThis (sql)
'

'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
Next j


If chkPrint.value = vbChecked Then
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
       Dim fso, chkPrinter, txtFile
        'ttt = "LPT1" 'LPT1
         Dim PORT As String
        PORT = ports
        'ttt = "LPT1" 'LPT1
        ttt = PORT
        Set fso = CreateObject("Scripting.FileSystemObject")
        Dim strReceipts As String
        Set txtFile = fso.CreateTextFile(ttt, True)
        
        txtFile.WriteLine "      " & cname & ""
        txtFile.WriteLine "      AGROVET"
        txtFile.WriteLine "      " & paddress & ""
        txtFile.WriteLine "      " & town & ""
        txtFile.WriteLine "      " & Phone & ""
        txtFile.WriteLine "      " & Email & ""
        
        txtFile.WriteLine "---------------------------------------"
        txtFile.WriteLine "  AGROVET RECEIPT"
        txtFile.WriteLine "     CASH SALES"
        txtFile.WriteLine "---------------------------------------"
        txtFile.WriteLine "QNTY" & vbTab & vbTab & "PRICE" & vbTab & vbTab & "AMOUNT"
        txtFile.WriteLine "........................................"
        j = 1
        strReceipts = ""
        Do While Not j > (Lvwitems.ListItems.Count)
            Lvwitems.ListItems.Item(j).selected = True
            strReceipts = strReceipts & Lvwitems.SelectedItem.SubItems(1) & vbNewLine & Lvwitems.SelectedItem.SubItems(2) & vbTab & vbTab
            strReceipts = strReceipts & Format(Lvwitems.SelectedItem.SubItems(3), "#,##0.00") & vbTab & vbTab & Format(Lvwitems.SelectedItem.SubItems(4), "#,##0.00") & vbNewLine
            txtFile.WriteLine strReceipts
            j = j + 1
        Loop
       
    txtFile.WriteLine "TOTAL" & vbTab & vbTab & vbTab & vbTab & Format(total, "#,##0.00") & vbNewLine
    txtFile.WriteLine
    txtFile.WriteLine "TOTAL" & vbTab & TXTTOTAL
    txtFile.WriteLine
    txtFile.WriteLine "AMOUNT RECEVED" & vbTab & txtamtreceived
    txtFile.WriteLine
    txtFile.WriteLine "CHANGE" & vbTab & TXTCHANGE
    txtFile.WriteLine
    txtFile.WriteLine "YOU WERE SERVED By " & UCase(username)
    txtFile.WriteLine "    Date :" & Format(Get_Server_Date, "dd/mm/yyyy HH:MM:SS AM/PM")
    txtFile.WriteLine "     THANK YOU AND WELCOME "
    
    txtFile.WriteLine " GOODS ONCE SOLD WILL NOT BE RE-ACCEPTED"
    txtFile.WriteLine "****************************************"
    txtFile.WriteLine escFeedAndCut
    txtFile.Close
    End If
End If
'//Update deductions
'If optCash.value = False Then
'Set cn = New ADODB.Connection
'sql = "d_sp_TransDeduct '" & txttranscode & "','" & txtransdate & "','Agrovet'," & Total & ",'" & Startdate & "','" & Enddate & "','" & User & "'"
'oSaccoMaster.ExecuteThis (sql)

'Set cn = New ADODB.Connection
'sql = "d_sp_TransDeduct '" & txtTCode & "','" & DTPDDate & "','" & cboDeductionType & "'," & txtamount & ",'" & DTPStartDate & "','" & DTPEndDate & "','" & User & "'"
'oSaccoMaster.ExecuteThis (sql)

'XXXXXXXXXXXXXXXXXXXXXXXxx


Lvwitems.ListItems.Clear
txtrno = ""
txtpcode.Text = ""
txtquantity = ""
txtamount = ""

MsgBox "Record saved Successfully"
Exit Sub
olkalou:
MsgBox err.description & " error occured."


End Sub
Sub savesno()
On Error GoTo ebraim

If lblCheckOff = True Then

Dim a, b, X
DIA = 0
Dim U As Double, S As Double
Dim cn As Connection
Dim j As Integer

If DIA = 1 Then
Exit Sub
End If
If Lvwitems.ListItems.Count = 0 Then
MsgBox "There are no items sold."
Exit Sub
End If
j = 1

Dim total As Currency
total = 0
Do While Not j > (Lvwitems.ListItems.Count)
'For j = 1 To Lvwitems.ListItems.Count
 Lvwitems.ListItems.Item(j).selected = True
 total = total + CCur(Lvwitems.SelectedItem.SubItems(4))
j = j + 1
Loop

If optCash.value = False Then

Startdate = DateSerial(year(txtransdate), month(txtransdate), 1)
Enddate = DateSerial(year(txtransdate), month(txtransdate) + 1, 1 - 1)


Set rs = oSaccoMaster.GetRecordset("d_sp_IsClosed '" & Enddate & "'")
If Not rs.EOF Then
    MsgBox "The period " & Enddate & " has been closed by " & rs.Fields(0)
    Exit Sub
End If
'End If
End If
j = 1
For j = 1 To Lvwitems.ListItems.Count
'Do While Not j > (Lvwitems.ListItems.Count)
 Lvwitems.ListItems.Item(j).selected = True
If Trim(txtSNo) = "" Then
txtSNo = "0"
End If
'// check if they are in stock.

Dim rsinstock As Recordset
sql = ""
sql = "select P_CODE,qout from ag_products where p_code='" & Lvwitems.SelectedItem & "'"

Set rsinstock = oSaccoMaster.GetRecordset(sql)

Dim Remain As Double
Remain = rsinstock.Fields(1) + CInt(Lvwitems.SelectedItem.SubItems(2))


sql = ""
sql = sql & "SET dateformat DMY INSERT INTO ag_Receipts(R_No, P_code, T_Date, Amount, Qua, S_Bal, user_id, Cash, SNo,Remarks,Branch) VALUES ("
sql = sql & txtrno & ",'" & Lvwitems.SelectedItem & "','" & txtransdate & "'," & -1 * Lvwitems.SelectedItem.SubItems(4)
sql = sql & "," & Lvwitems.SelectedItem.SubItems(2) & "," & Remain & ",'" & User & "'," & Lvwitems.SelectedItem.SubItems(5) & "," & txtSNo & ",'Sales Return','" & Cmbstation & "')"

oSaccoMaster.ExecuteThis (sql)
sql = ""
sql = sql & "SET dateformat DMY INSERT INTO ag_Receipts3(R_No, P_code, T_Date, Amount, Qua, S_Bal, user_id, Cash, SNo,Remarks) VALUES ("
sql = sql & txtrno & ",'" & Lvwitems.SelectedItem & "','" & txtransdate & "'," & -1 * Lvwitems.SelectedItem.SubItems(4)
sql = sql & "," & Lvwitems.SelectedItem.SubItems(2) & "," & Remain & ",'" & User & "'," & Lvwitems.SelectedItem.SubItems(5) & "," & txtSNo & ",'Sales Return')"

oSaccoMaster.ExecuteThis (sql)
oSaccoMaster.ExecuteThis ("Update ag_Products SET Qout =" & CCur(Remain) & " WHERE p_code= '" & Lvwitems.SelectedItem & "'")
'//XXXXXXXXXXXXXXX
    sql = ""
    sql = "set dateformat dmy insert into gltransactions(transdate,amount,draccno,craccno,documentno,source,transdescript,auditid,cash,doc_posted) values('" & txtransdate & "'," & Lvwitems.SelectedItem.SubItems(4) & ",'A005','I002','" & Lvwitems.SelectedItem & "','" & cboproductname & "' ,' CHECK OFF SALES ','" & User & "',0,0)"
    oSaccoMaster.ExecuteThis (sql)




'XXXXXXXXXXXXXXXXXXXXXX
Next j
'j = j + 1
'Loop

If optCash.value = False Then
Set cn = New ADODB.Connection
sql = "d_sp_SupplierDeduct " & txtSNo & ",'" & txtransdate & "','Agrovet'," & -1 * total & ",'" & Startdate & "','" & Enddate & "'," & year(txtransdate) & ",'" & User & "','Receipt " & txtrno & "','" & Cmbstation & "'"
oSaccoMaster.ExecuteThis (sql)
End If

If CDbl(txtamtreceived) > 0 Then
    '******Deduct Amount paid in cash
   
    amount = 0
    amount = CDbl(txtamtreceived)
    sql = "d_sp_SupplierDeduct " & txtSNo & ",'" & txtransdate & "','Agrovet'," & -1 * amount & ",'" & Startdate & "','" & Enddate & "'," & year(txtransdate) & ",'" & User & "','Cash','" & Cmbstation & "'"
oSaccoMaster.ExecuteThis (sql)
End If

'//Update deductions
If chkPrint.value = vbChecked Then
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
       Dim fso, chkPrinter, txtFile
        'ttt = "LPT1" 'LPT1
         Dim PORT As String
        PORT = ports
        'ttt = "LPT1" 'LPT1
        ttt = PORT
        Set fso = CreateObject("Scripting.FileSystemObject")
        Dim strReceipts As String
        j = 1
        strReceipts = ""
        Do While Not j > (Lvwitems.ListItems.Count)
            Lvwitems.ListItems.Item(j).selected = True
            strReceipts = strReceipts & Lvwitems.SelectedItem.SubItems(1) & vbNewLine & Lvwitems.SelectedItem.SubItems(2) & vbTab & vbTab
            strReceipts = strReceipts & Format(Lvwitems.SelectedItem.SubItems(3), "#,##0.00") & vbTab & vbTab & Format(Lvwitems.SelectedItem.SubItems(4), "#,##0.00") & vbNewLine
            j = j + 1
        Loop
        
        strReceipts = strReceipts & vbNewLine & "---------------------------------------" & vbNewLine
        strReceipts = strReceipts & "TOTAL" & vbTab & vbTab & vbTab & vbTab & Format(total, "#,##0.00") & vbNewLine
        strReceipts = strReceipts & "======================================="
        Set txtFile = fso.CreateTextFile(ttt, True)
    txtFile.WriteLine "     " & cname & ""
    txtFile.WriteLine "      " & paddress & ""
    txtFile.WriteLine "---------------------------------------"
    txtFile.WriteLine "    AGROVET RECEIPT"
    txtFile.WriteLine "     Check-off"
    txtFile.WriteLine "......................................."
    If lblCheckOff = True Then
    txtFile.WriteLine "SNo:" & txtSNo
    txtFile.WriteLine "Name:" & lblSNames
    End If
    txtFile.WriteLine "---------------------------------------"

    txtFile.WriteLine "QNTY" & vbTab & vbTab & "PRICE" & vbTab & vbTab & "AMOUNT"
    txtFile.WriteLine "......................................."
    txtFile.WriteLine strReceipts
    txtFile.WriteLine
    txtFile.WriteLine "SALES RETURNED"
    txtFile.WriteLine "AMOUNT RECEIVED" & vbTab & txtamtreceived
    txtFile.WriteLine
    txtFile.WriteLine "CHANGE" & vbTab & vbTab & IIf(CDbl(TXTCHANGE) < 0, 0, CDbl(TXTCHANGE))
    txtFile.WriteLine
    txtFile.WriteLine "YOU WERE SERVED By " & UCase(username)
    txtFile.WriteLine "    Date :" & Format(Get_Server_Date, "dd/mm/yyyy HH:MM:SS AM/PM")
    txtFile.WriteLine "     THANK YOU AND WELCOME "
    txtFile.WriteLine "****************************************"
    txtFile.WriteLine escFeedAndCut
    txtFile.Close
End If
End If

Lvwitems.ListItems.Clear
txtpcode.Text = ""
txtquantity = ""
txtamount = ""
cboproductname = ""
txtrno = ""
'txtSNo = ""
lblTKgs = ""
lblGPay = ""
lblDed = ""
lblNPay = ""
lblSNames = ""
cmdNew_Click
MsgBox "Records saved"
Exit Sub
ebraim:
MsgBox err.description & " error occured."
End Sub
Private Sub savestaff()
On Error GoTo olkalou

If optstaff = True Then
Dim C As String
C = "Staff" & txtstaffno

Dim j As Integer
If Lvwitems.ListItems.Count = 0 Then
MsgBox "There are no items sold."
Exit Sub
End If
If txtstaffno = "" Then
MsgBox "Enter Staff Number before you continue", vbCritical, "Maziwa"

Exit Sub
End If
j = 1

Dim total As Currency
total = 0
Do While Not j > (Lvwitems.ListItems.Count)
 Lvwitems.ListItems.Item(j).selected = True
 total = total + CCur(Lvwitems.SelectedItem.SubItems(4))
j = j + 1
Loop



Startdate = DateSerial(year(txtransdate), month(txtransdate), 1)
Enddate = DateSerial(year(txtransdate), month(txtransdate) + 1, 1 - 1)

Set rs = oSaccoMaster.GetRecordset("d_sp_IsClosed '" & Enddate & "'")
If Not rs.EOF Then
    MsgBox "The period " & Enddate & " has been closed by " & rs.Fields(0)
    Exit Sub
End If


'// check if they are in stock.
For j = 1 To Lvwitems.ListItems.Count
 Lvwitems.ListItems.Item(j).selected = True


Dim rsinstock As Recordset
sql = ""
sql = "select P_CODE,qout from ag_products where p_code='" & Lvwitems.SelectedItem & "'"
'//Set rsinstock = New ADODB.Recordset
Set rsinstock = oSaccoMaster.GetRecordset(sql)
'// check the stock if it is less than zero
If rsinstock.Fields(1) <= 0 Then
MsgBox "Sorry Stock is Zero for item " & Lvwitems.SelectedItem.SubItems(1) & " please re-stock before your proceed", vbInformation
Exit Sub
End If
'// check the quanttity being sold versus the balance
Dim Remain As Double
Remain = rsinstock.Fields(1) + CInt(Lvwitems.SelectedItem.SubItems(2))
If Remain < 0 Then
MsgBox "Stock will be negative " & Remain & " please re-stock before you proceed", vbInformation
Exit Sub
End If

sql = ""
sql = sql & "SET dateformat DMY INSERT INTO ag_Receipts(R_No, P_code, T_Date, Amount, Qua, S_Bal, user_id, Cash, SNo,Branch) VALUES ("
sql = sql & txtrno & ",'" & Lvwitems.SelectedItem & "','" & txtransdate & "'," & -1 * Lvwitems.SelectedItem.SubItems(4)
sql = sql & "," & Lvwitems.SelectedItem.SubItems(2) & "," & Remain & ",'" & User & "'," & Lvwitems.SelectedItem.SubItems(5) & ",'" & C & "','" & Cmbstation & "')"

oSaccoMaster.ExecuteThis (sql)
oSaccoMaster.ExecuteThis ("Update ag_Products SET Qout =" & CCur(Remain) & " WHERE p_code= '" & Lvwitems.SelectedItem & "'")


    sql = ""
    sql = "set dateformat dmy insert into gltransactions(transdate,amount,draccno,craccno,documentno,source,transdescript,auditid,cash,doc_posted) values('" & txtransdate & "'," & Lvwitems.SelectedItem & ",'A006','I004','" & Lvwitems.SelectedItem & "','" & cboproductname & "' ,'" & C & "','" & User & "',0,0)"
    oSaccoMaster.ExecuteThis (sql)
Next j

If chkPrint.value = vbChecked Then
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
       Dim fso, chkPrinter, txtFile
        'ttt = "LPT1" 'LPT1
         Dim PORT As String
        PORT = ports
        'ttt = "LPT1" 'LPT1
        ttt = PORT
        Set fso = CreateObject("Scripting.FileSystemObject")
        Dim strReceipts As String
        j = 1
        strReceipts = ""
        Do While Not j > (Lvwitems.ListItems.Count)
            Lvwitems.ListItems.Item(j).selected = True
            strReceipts = strReceipts & Lvwitems.SelectedItem.SubItems(1) & vbNewLine & Lvwitems.SelectedItem.SubItems(2) & vbTab & vbTab
            strReceipts = strReceipts & Format(Lvwitems.SelectedItem.SubItems(3), "#,##0.00") & vbTab & vbTab & Format(Lvwitems.SelectedItem.SubItems(4), "#,##0.00") & vbNewLine
            j = j + 1
        Loop

        'MsgBox strReceipts
        strReceipts = strReceipts & vbNewLine & "---------------------------------------" & vbNewLine
        strReceipts = strReceipts & "TOTAL" & vbTab & vbTab & vbTab & vbTab & Format(total, "#,##0.00") & vbNewLine
        strReceipts = strReceipts & "======================================="
        Set txtFile = fso.CreateTextFile(ttt, True)
        
        If optCash = True Then
        Set rs = New ADODB.Recordset
        Dim a As String
        sql = "select Adress from d_company"
        Set rs = oSaccoMaster.GetRecordset(sql)
        If Not rs.EOF Then
        If Not IsNull(rs.Fields(0)) Then a = rs.Fields(0)
        End If
    txtFile.WriteLine "  " & cname & ""
    txtFile.WriteLine "     " & a & ""
    txtFile.WriteLine "---------------------------------------"
    txtFile.WriteLine "  AGROVET RECEIPT"
    txtFile.WriteLine "     STAFF SALES"
    txtFile.WriteLine "---------------------------------------"
'nAME QNTY PRICE AMOUNT
    txtFile.WriteLine "QNTY" & vbTab & vbTab & "PRICE" & vbTab & vbTab & "AMOUNT"
    txtFile.WriteLine "........................................"
    txtFile.WriteLine strReceipts
    txtFile.WriteLine
    txtFile.WriteLine "TOTAL" & TXTTOTAL
    txtFile.WriteLine
    txtFile.WriteLine "AMOUNT RECEVED" & txtamtreceived
    txtFile.WriteLine
    txtFile.WriteLine "CHANGE" & TXTCHANGE
    txtFile.WriteLine
    txtFile.WriteLine "YOU WERE SERVED By " & UCase(username)
    txtFile.WriteLine "    Date :" & Format(Get_Server_Date, "dd/mm/yyyy HH:MM:SS AM/PM")
    txtFile.WriteLine "     THANK YOU AND WELCOME "
    txtFile.WriteLine "****************************************"
    txtFile.WriteLine escFeedAndCut
    txtFile.Close
    End If
End If
'//Update deductions
'If optCash.value = False Then
'Set cn = New ADODB.Connection
'sql = "d_sp_TransDeduct '" & txttranscode & "','" & txtransdate & "','Agrovet'," & Total & ",'" & Startdate & "','" & Enddate & "','" & User & "'"
'oSaccoMaster.ExecuteThis (sql)

'Set cn = New ADODB.Connection
'sql = "d_sp_TransDeduct '" & txtTCode & "','" & DTPDDate & "','" & cboDeductionType & "'," & txtamount & ",'" & DTPStartDate & "','" & DTPEndDate & "','" & User & "'"
'oSaccoMaster.ExecuteThis (sql)





Lvwitems.ListItems.Clear
txtrno = ""
txtpcode.Text = ""
txtquantity = ""
txtamount = ""

MsgBox "Record saved Successfully"
Exit Sub
End If
olkalou:
MsgBox err.description & " error occured."

End Sub
Private Sub saveothers()
On Error GoTo olkalou

If Optothers = True Then
Dim C As String
Dim D As String
C = "Other"
D = "Non suppliers"
Dim j As Integer
If Lvwitems.ListItems.Count = 0 Then
MsgBox "There are no items sold."
Exit Sub
End If
'If txtstaffno = "" Then
'MsgBox "Enter Staff Number before you continue", vbCritical, "Maziwa"
'
'Exit Sub
'End If
j = 1

Dim total As Currency
total = 0
Do While Not j > (Lvwitems.ListItems.Count)
 Lvwitems.ListItems.Item(j).selected = True
 total = total + CCur(Lvwitems.SelectedItem.SubItems(4))
j = j + 1
Loop



Startdate = DateSerial(year(txtransdate), month(txtransdate), 1)
Enddate = DateSerial(year(txtransdate), month(txtransdate) + 1, 1 - 1)

Set rs = oSaccoMaster.GetRecordset("d_sp_IsClosed '" & Enddate & "'")
If Not rs.EOF Then
    MsgBox "The period " & Enddate & " has been closed by " & rs.Fields(0)
    Exit Sub
End If


'// check if they are in stock.
For j = 1 To Lvwitems.ListItems.Count
 Lvwitems.ListItems.Item(j).selected = True


Dim rsinstock As Recordset
sql = ""
sql = "select P_CODE,qin,qout from ag_products where p_code='" & Lvwitems.SelectedItem & "'"
'//Set rsinstock = New ADODB.Recordset
Set rsinstock = oSaccoMaster.GetRecordset(sql)
'// check the stock if it is less than zero
If rsinstock.Fields(1) <= 0 Then
MsgBox "Sorry Stock is Zero for item " & Lvwitems.SelectedItem.SubItems(1) & " please re-stock before your proceed", vbInformation
Exit Sub
End If
'// check the quanttity being sold versus the balance
Dim Remain As Double
Remain = rsinstock.Fields(1) + CInt(Lvwitems.SelectedItem.SubItems(2))
If Remain < 0 Then
MsgBox "Stock will be negative " & Remain & " please re-stock before you proceed", vbInformation
Exit Sub
End If

sql = ""
sql = sql & "SET dateformat DMY INSERT INTO ag_Receipts(R_No, P_code, T_Date, Amount, Qua, S_Bal, user_id, Cash, SNo,Branch) VALUES ("
sql = sql & txtrno & ",'" & Lvwitems.SelectedItem & "','" & txtransdate & "'," & -1 * Lvwitems.SelectedItem.SubItems(4)
sql = sql & "," & Lvwitems.SelectedItem.SubItems(2) & "," & Remain & ",'" & User & "'," & Lvwitems.SelectedItem.SubItems(5) & ",'" & C & "','" & Cmbstation & "')"

oSaccoMaster.ExecuteThis (sql)
oSaccoMaster.ExecuteThis ("Update ag_Products SET Qout =" & CCur(Remain) & " WHERE p_code= '" & Lvwitems.SelectedItem & "'")


    sql = ""
    sql = "set dateformat dmy insert into gltransactions(transdate,amount,draccno,craccno,documentno,source,transdescript,auditid,cash,doc_posted) values('" & txtransdate & "'," & Lvwitems.SelectedItem & ",'A006','I004','" & Lvwitems.SelectedItem & "','" & cboproductname & "' ,'" & C & "','" & User & "',0,0)"
    oSaccoMaster.ExecuteThis (sql)
Next j

If chkPrint.value = vbChecked Then
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
       Dim fso, chkPrinter, txtFile
        'ttt = "LPT1" 'LPT1
         Dim PORT As String
        PORT = ports
        'ttt = "LPT1" 'LPT1
        ttt = PORT
        Set fso = CreateObject("Scripting.FileSystemObject")
        Dim strReceipts As String
        j = 1
        strReceipts = ""
        Do While Not j > (Lvwitems.ListItems.Count)
            Lvwitems.ListItems.Item(j).selected = True
            strReceipts = strReceipts & Lvwitems.SelectedItem.SubItems(1) & vbNewLine & Lvwitems.SelectedItem.SubItems(2) & vbTab & vbTab
            strReceipts = strReceipts & Format(Lvwitems.SelectedItem.SubItems(3), "#,##0.00") & vbTab & vbTab & Format(Lvwitems.SelectedItem.SubItems(4), "#,##0.00") & vbNewLine
            j = j + 1
        Loop

        'MsgBox strReceipts
        strReceipts = strReceipts & vbNewLine & "---------------------------------------" & vbNewLine
        strReceipts = strReceipts & "TOTAL" & vbTab & vbTab & vbTab & vbTab & Format(total, "#,##0.00") & vbNewLine
        strReceipts = strReceipts & "======================================="
        Set txtFile = fso.CreateTextFile(ttt, True)
        
        If optCash = True Then
        Set rs = New ADODB.Recordset
        Dim a As String
        sql = "select Adress from d_company"
        Set rs = oSaccoMaster.GetRecordset(sql)
        If Not rs.EOF Then
        If Not IsNull(rs.Fields(0)) Then a = rs.Fields(0)
        End If
    txtFile.WriteLine "  " & cname & ""
    txtFile.WriteLine "     " & a & ""
    txtFile.WriteLine "---------------------------------------"
    txtFile.WriteLine "  AGROVET RECEIPT"
    txtFile.WriteLine "     NON SUPPLIERS SALES"
    txtFile.WriteLine "---------------------------------------"
'nAME QNTY PRICE AMOUNT
    txtFile.WriteLine "QNTY" & vbTab & vbTab & "PRICE" & vbTab & vbTab & "AMOUNT"
    txtFile.WriteLine "........................................"
    txtFile.WriteLine strReceipts
    txtFile.WriteLine
    txtFile.WriteLine "TOTAL" & TXTTOTAL
    txtFile.WriteLine
    txtFile.WriteLine "AMOUNT RECEVED" & txtamtreceived
    txtFile.WriteLine
    txtFile.WriteLine "CHANGE" & TXTCHANGE
    txtFile.WriteLine
    txtFile.WriteLine "YOU WERE SERVED By " & UCase(username)
    txtFile.WriteLine "    Date :" & Format(Get_Server_Date, "dd/mm/yyyy HH:MM:SS AM/PM")
    txtFile.WriteLine "     THANK YOU AND WELCOME "
    txtFile.WriteLine "****************************************"
    txtFile.WriteLine escFeedAndCut
    txtFile.Close
    End If
End If
'//Update deductions
'If optCash.value = False Then
'Set cn = New ADODB.Connection
'sql = "d_sp_TransDeduct '" & txttranscode & "','" & txtransdate & "','Agrovet'," & Total & ",'" & Startdate & "','" & Enddate & "','" & User & "'"
'oSaccoMaster.ExecuteThis (sql)

'Set cn = New ADODB.Connection
'sql = "d_sp_TransDeduct '" & txtTCode & "','" & DTPDDate & "','" & cboDeductionType & "'," & txtamount & ",'" & DTPStartDate & "','" & DTPEndDate & "','" & User & "'"
'oSaccoMaster.ExecuteThis (sql)





Lvwitems.ListItems.Clear
txtrno = ""
txtpcode.Text = ""
txtquantity = 1
txtamount = ""

MsgBox "Record saved Successfully"
Exit Sub
End If
olkalou:
MsgBox err.description & " error occured."

End Sub
Sub ReceiveFromStaion()
On Error GoTo kiparu

If optrecieve = True Then
Dim centre As String
centre = Cmbstation.Text
If Trim(Cmbstation.Text) = "" Then
 MsgBox "Please enter the Agrovet Station."
   Cmbstation.SetFocus
Exit Sub
End If


Dim j As Integer
If Lvwitems.ListItems.Count = 0 Then
MsgBox "There are no items Received."
Exit Sub
End If
j = 1

Dim total As Currency
total = 0
Do While Not j > (Lvwitems.ListItems.Count)
 Lvwitems.ListItems.Item(j).selected = True
 total = total + CCur(Lvwitems.SelectedItem.SubItems(4))
j = j + 1
Loop


Startdate = DateSerial(year(txtransdate), month(txtransdate), 1)
Enddate = DateSerial(year(txtransdate), month(txtransdate) + 1, 1 - 1)


Set rs = oSaccoMaster.GetRecordset("d_sp_IsClosed '" & Enddate & "'")
If Not rs.EOF Then
    MsgBox "The period " & Enddate & " has been closed by " & rs.Fields(0)
    Exit Sub
End If
'End If
'End If
For j = 1 To Lvwitems.ListItems.Count
'Do While Not j > (Lvwitems.ListItems.Count)
 Lvwitems.ListItems.Item(j).selected = True



Dim rsinstock As Recordset
sql = ""
sql = "select P_CODE,qout,Pprice,Sprice from ag_products where p_code='" & Lvwitems.SelectedItem & "'"

Set rsinstock = oSaccoMaster.GetRecordset(sql)

'// check the quanttity being sold versus the balance
Dim Remain, Pprice, amount As Double
Dim pcode As String
Remain = rsinstock.Fields(1) + CInt(Lvwitems.SelectedItem.SubItems(2))
Pprice = IIf(IsNull(rsinstock.Fields(2)), CInt(Lvwitems.SelectedItem.SubItems(3)), rsinstock.Fields(2))
amount = CInt(Lvwitems.SelectedItem.SubItems(3)) * CInt(Lvwitems.SelectedItem.SubItems(2))

'sql = ""
'sql = sql & "SET dateformat DMY INSERT INTO ag_Receipts(R_No, P_code, T_Date, Amount, Qua, S_Bal, user_id, Cash, SNo,Remarks) VALUES ("
'sql = sql & txtrno & ",'" & Lvwitems.SelectedItem & "','" & txtransdate & "'," & Amount
'sql = sql & "," & Lvwitems.SelectedItem.SubItems(2) & "," & Remain & ",'" & User & "'," & Lvwitems.SelectedItem.SubItems(5) & ",'" & Cmbstation & "','Sales Return')"

'oSaccoMaster.ExecuteThis (sql)
'oSaccoMaster.ExecuteThis ("Update ag_Products SET Qout =" & CCur(Remain) & " WHERE p_code= '" & Lvwitems.SelectedItem & "'")

'********************************************************************************
pcode = Lvwitems.SelectedItem
sql = "select P_CODE,qout,unserialized,pprice,sprice,o_bal from ag_products where p_code='" & pcode & "'"
Set rs = New ADODB.Recordset
rs.Open sql, cn
pcode = rs.Fields("p_code")
sql = ""
sql = "set dateformat DMY update ag_products set qin=" & Lvwitems.ListItems(j).SubItems(2) & ",qout=" & Remain & ",o_bal=" & Remain & ",last_d_updated='" & txtransdate & "',user_id='" & User & "',audit_date='" & Get_Server_Date & "',unserialized=0,SERIA=0,pprice=" & rs.Fields("pprice") & ",sprice=" & rs.Fields("sprice") & " where p_code='" & pcode & "'"
cn.Execute sql




Dim rsst As Recordset
sql = ""
sql = "set dateformat DMY select top 1 * from ag_stockbalance where p_code='" & pcode & "' order by trackid desc "
Set rsst = New ADODB.Recordset
rsst.Open sql, cn
If Not rsst.EOF Then
sql = ""
sql = "set dateformat DMY INSERT INTO ag_stockbalance"
sql = sql & " (p_code, productname, openningstock, changeinstock, stockbalance, transdate,companyid)"
sql = sql & " VALUES     ('" & pcode & "', '" & Lvwitems.ListItems(j).SubItems(1) & "', '" & rs.Fields("o_bal") & "', '" & Lvwitems.ListItems(j).SubItems(2) & "', '" & Remain & "', '" & Format(Get_Server_Date, "dd/mm/yyyy") & "',1)"
cn.Execute sql
End If

'Dim DRaccno As String
'Dim Craccno As String
'If centre = "SANGALO" Then
'    Craccno = "A008"
'    DRaccno = "I004"
'ElseIf centre = "OLMAROROI" Then
'    Craccno = "A010"
'    DRaccno = "I005"
'ElseIf centre = "KABISAGA" Then
'    Craccno = "A012"
'    DRaccno = "I006"
'ElseIf centre = "KOISOLIK" Then
'    Craccno = "A009"
'    DRaccno = "I007"
'ElseIf centre = "CHEMUSWO" Then
'    Craccno = "A011"
'    DRaccno = "I008"
'ElseIf centre = "BELEKENYA" Then
'    Craccno = "A013"
'End If
''XXXXXXXXXXX SAVE TO GL
'    sql = ""
'    sql = "set dateformat dmy insert into gltransactions(transdate,amount,draccno,craccno,documentno,source,transdescript,auditid,cash,doc_posted) values('" & txtransdate & "'," & Amount & ",'" & DRaccno & "','" & Craccno & "','" & lvwItems.SelectedItem & "','" & cboproductname & "' ,'SALES RETURN ','" & User & "',0,0)"
'    oSaccoMaster.ExecuteThis (sql)


'XXXXXXXXXXXXXXXXXXXXXX


Next j
If chkPrint.value = vbChecked Then
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
       Dim fso, chkPrinter, txtFile
        'ttt = "LPT1" 'LPT1
         Dim PORT As String
        PORT = ports
        'ttt = "LPT1" 'LPT1
        ttt = PORT
        Set fso = CreateObject("Scripting.FileSystemObject")
        Dim strReceipts As String
         
        Set txtFile = fso.CreateTextFile(ttt, True)
        txtFile.WriteLine "      " & cname & ""
        txtFile.WriteLine "      AGROVET"
        txtFile.WriteLine "      " & paddress & ""
        txtFile.WriteLine "      " & town & ""
        txtFile.WriteLine "      " & Phone & ""
        txtFile.WriteLine "      " & Email & ""
        txtFile.WriteLine "---------------------------------------"
        txtFile.WriteLine "    RECEIVING VOUCHER"
        txtFile.WriteLine
        txtFile.WriteLine "    INVOICE NO:" & "-------------------"""
        txtFile.WriteLine "......................................."
        txtFile.WriteLine "Station :" & Cmbstation
        txtFile.WriteLine
        txtFile.WriteLine "---------------------------------------"
        txtFile.WriteLine "ITEM" & vbTab & vbTab & "QNTY" & vbTab & "PRICE" & vbTab & "AMOUNT"
        txtFile.WriteLine "......................................."
       
        j = 1
        strReceipts = ""
        Do While Not j > (Lvwitems.ListItems.Count)
            Lvwitems.ListItems.Item(j).selected = True
            strReceipts = Lvwitems.SelectedItem.SubItems(1)
            If Len(strReceipts) > 14 Then
            strReceipts = strReceipts & "-"
            Else
            strReceipts = strReceipts & vbTab
            End If
            strReceipts = strReceipts & CDbl(Lvwitems.SelectedItem.SubItems(2)) & vbTab & Format(Lvwitems.SelectedItem.SubItems(3), "#,##0.00") & vbTab & Format(Lvwitems.SelectedItem.SubItems(4), "#,##0.00") & vbNewLine
            txtFile.WriteLine strReceipts
            j = j + 1
        Loop
      
        txtFile.WriteLine "---------------------------------------" & vbNewLine
        txtFile.WriteLine "RECEIPT TOTAL" & vbTab & vbTab & Format(TXTTOTAL, "#,##0.00") & vbNewLine
        txtFile.WriteLine "======================================="
        txtFile.WriteLine
        txtFile.WriteLine "Remarks" & vbTab & txtcomment
        txtFile.WriteLine
        txtFile.WriteLine "YOU WERE SERVED By " & UCase(username)
        txtFile.WriteLine "    Date :" & Format(Get_Server_Date, "dd/mm/yyyy HH:MM:SS AM/PM")
        txtFile.WriteLine "     THANK YOU AND WELCOME "
        txtFile.WriteLine "****************************************"
        txtFile.WriteLine escFeedAndCut
        txtFile.Close
End If
End If
Lvwitems.ListItems.Clear
txtrno = ""
txtpcode.Text = ""
txtquantity = 1
txtamount = ""
Cmbstation.Text = ""
txtcomment = ""
MsgBox "Record saved Successfully"
Exit Sub
kiparu:
MsgBox err.description & " error occured."

End Sub

