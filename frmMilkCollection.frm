VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmMilkCollection 
   BackColor       =   &H00FFFFC0&
   Caption         =   "Milk Collection"
   ClientHeight    =   9495
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11100
   BeginProperty Font 
      Name            =   "Sylfaen"
      Size            =   15.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMilkCollection.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "frmMilkCollection.frx":164A
   ScaleHeight     =   9495
   ScaleWidth      =   11100
   StartUpPosition =   1  'CenterOwner
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
      Left            =   7680
      TabIndex        =   33
      Top             =   1320
      Width           =   3255
   End
   Begin MSComctlLib.ListView lvwpreviousmilk 
      Height          =   2415
      Left            =   4440
      TabIndex        =   32
      ToolTipText     =   """milk supply trend"""
      Top             =   600
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   4260
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   255
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
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
      ItemData        =   "frmMilkCollection.frx":190C
      Left            =   8760
      List            =   "frmMilkCollection.frx":191C
      TabIndex        =   30
      Text            =   "\\127.0.0.1\E-PoS 80mm Thermal Printer"
      Top             =   1920
      Width           =   2175
   End
   Begin VB.CommandButton cmdreprintreceipt 
      Caption         =   "Reprint"
      Height          =   405
      Left            =   9600
      TabIndex        =   29
      ToolTipText     =   "Click "
      Top             =   3120
      Width           =   1335
   End
   Begin VB.CommandButton cmdget 
      Caption         =   "Get DD"
      Height          =   405
      Left            =   12360
      TabIndex        =   26
      ToolTipText     =   "Click to receive the milk"
      Top             =   600
      Width           =   1335
   End
   Begin VB.CommandButton cmdfind 
      Caption         =   "F"
      Height          =   405
      Left            =   3480
      TabIndex        =   25
      Top             =   240
      Width           =   375
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   24
      Top             =   9240
      Width           =   11100
      _ExtentX        =   19579
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   2
            Object.Width           =   7937
            MinWidth        =   7937
            Text            =   "USER : Birgen Gideon K."
            TextSave        =   "USER : Birgen Gideon K."
            Object.ToolTipText     =   "EASYMA User"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   5292
            MinWidth        =   5292
            Picture         =   "frmMilkCollection.frx":1938
            Text            =   "DATE : 07/12/2009"
            TextSave        =   "DATE : 07/12/2009"
            Object.ToolTipText     =   "Today's Date"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "17:20"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Enabled         =   0   'False
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "CAPS"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Enabled         =   0   'False
            TextSave        =   "NUM"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CheckBox chkNotepad 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      Caption         =   "To Notepad"
      ForeColor       =   &H80000008&
      Height          =   405
      Left            =   2400
      TabIndex        =   23
      Top             =   2520
      Width           =   1935
   End
   Begin MSComDlg.CommonDialog cdgPrint 
      Left            =   7560
      Top             =   3000
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      FileName        =   "c:\receipt.txt"
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   405
      Left            =   5160
      TabIndex        =   22
      Top             =   8640
      Width           =   1095
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "Remove"
      Height          =   405
      Left            =   12360
      TabIndex        =   19
      ToolTipText     =   "Click "
      Top             =   1200
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton cmdReceive 
      Caption         =   "Receive"
      Default         =   -1  'True
      Height          =   405
      Left            =   8160
      TabIndex        =   10
      ToolTipText     =   "Click to receive the milk"
      Top             =   3120
      Width           =   1335
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Today's collection"
      Height          =   3015
      Left            =   120
      TabIndex        =   17
      Top             =   5520
      Width           =   10815
      Begin MSComctlLib.ListView lvwMilkCollected 
         Height          =   2415
         Left            =   120
         TabIndex        =   18
         Top             =   480
         Width           =   10215
         _ExtentX        =   18018
         _ExtentY        =   4260
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   12582912
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Sylfaen"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Transporter's Receipt"
      Height          =   1695
      Left            =   120
      TabIndex        =   12
      Top             =   3600
      Width           =   10815
      Begin VB.TextBox txttotalsup 
         Appearance      =   0  'Flat
         ForeColor       =   &H000000FF&
         Height          =   525
         Left            =   8640
         TabIndex        =   35
         Top             =   480
         Width           =   1455
      End
      Begin VB.CommandButton cmdprintself 
         Caption         =   "Print Self"
         Height          =   405
         Left            =   8280
         TabIndex        =   28
         Top             =   1200
         Width           =   2295
      End
      Begin VB.CommandButton cmdprintlist 
         Caption         =   "Print Transporter"
         Height          =   405
         Left            =   2520
         TabIndex        =   27
         Top             =   1200
         Width           =   2655
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "Print Summary"
         Height          =   405
         Left            =   5400
         TabIndex        =   16
         Top             =   1200
         Width           =   2655
      End
      Begin VB.TextBox txtTCode 
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
         Height          =   405
         Left            =   2280
         TabIndex        =   15
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFC0&
         Caption         =   "Total  suppliers"
         Height          =   405
         Left            =   6360
         TabIndex        =   34
         Top             =   480
         Width           =   2070
      End
      Begin VB.Label lblTName 
         AutoSize        =   -1  'True
         BackColor       =   &H0080FF80&
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
         Left            =   3360
         TabIndex        =   14
         Top             =   600
         Width           =   75
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Transporter code"
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
         TabIndex        =   13
         Top             =   600
         Width           =   2175
      End
   End
   Begin VB.TextBox txtQnty 
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
      Left            =   3240
      TabIndex        =   9
      Top             =   1080
      Width           =   1095
   End
   Begin VB.TextBox txtSNo 
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
      Left            =   2400
      TabIndex        =   8
      Top             =   240
      Width           =   975
   End
   Begin MSComCtl2.DTPicker DTPMilkDate 
      Height          =   495
      Left            =   8640
      TabIndex        =   7
      Top             =   600
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   873
      _Version        =   393216
      MouseIcon       =   "frmMilkCollection.frx":1ACC
      CalendarBackColor=   8454016
      Format          =   108920833
      CurrentDate     =   40095
   End
   Begin VB.CheckBox chkPrint 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      Caption         =   "Print Receipt"
      ForeColor       =   &H80000008&
      Height          =   405
      Left            =   120
      TabIndex        =   6
      Top             =   2520
      Width           =   2295
   End
   Begin VB.CheckBox chkComment 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      Caption         =   "Add Comment"
      ForeColor       =   &H80000008&
      Height          =   405
      Left            =   8280
      TabIndex        =   5
      Top             =   2520
      Width           =   2535
   End
   Begin VB.TextBox txtComment 
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
      Left            =   1560
      TabIndex        =   4
      Text            =   "<Put your comment here>"
      Top             =   3120
      Visible         =   0   'False
      Width           =   5655
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
      Left            =   7680
      TabIndex        =   31
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Label lblDTotal 
      AutoSize        =   -1  'True
      BackColor       =   &H00004040&
      Caption         =   "0"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#,##0.00;(#,##0.00)"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   525
      Left            =   2640
      TabIndex        =   21
      Top             =   1680
      Width           =   270
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Today's Total (Kgs)"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   20
      Top             =   1800
      Width           =   2415
   End
   Begin VB.Label lblNames 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFC0&
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
      Left            =   3960
      TabIndex        =   11
      Top             =   240
      Width           =   315
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Milk Date"
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
      Left            =   7320
      TabIndex        =   3
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label lblComment 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Comment"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   3120
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Quantity Supplied (Kgs)"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   3015
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Supplier Number"
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
      TabIndex        =   0
      Top             =   240
      Width           =   2175
   End
End
Attribute VB_Name = "frmMilkCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Rs1 As New ADODB.Recordset, rs2 As New ADODB.Recordset, rs3 As New ADODB.Recordset
Dim rs4 As New ADODB.Recordset, CummulKgs As Double, TRANSPORTER As String
Dim Transport As Currency, agrovet As Currency, AI As Currency, TMShares As Currency, FSA As Currency, HShares As Currency, Advance As Currency, Others As Currency
Private Sub chkComment_Click()
If chkComment.value = vbChecked Then
    lblComment.Visible = True
    txtComment = "<Put your comment here>"
    txtComment.Visible = True
    txtComment.SetFocus
Else
    lblComment.Visible = False
    txtComment.Visible = False

End If
End Sub

Private Sub Text1_GotFocus()

End Sub

Private Sub chprint_Click()

ports.Clear
ports = ""
'//If the drivers are installed it won't matter whether the Port is indicated
' or not it will just work.

If chprint.value = vbChecked Then
ports.AddItem "LPT1"
ports = "LPT1"
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

Private Sub cmdclose_Click()
Unload Me
End Sub

Private Sub cmdFind_Click()
        Me.MousePointer = vbHourglass
        frmSearchSupplier.Show vbModal
        txtSNo = sel
        txtSNo_Validate True
        Me.MousePointer = 0

End Sub

Private Sub cmdget_Click()
On Error GoTo ErrorHandler
Dim Price As Currency
Dim Startdate, CummulKgs, TRANSPORTER As String
Dim transdate As Date

'//open the database of serby
Dim PROVIDER2 As String
PROVIDER2 = "SERBRY"
Dim cn2 As New ADODB.Connection
Set cn2 = New ADODB.Connection
Dim RN2 As New ADODB.Recordset
cn2.Open PROVIDER2, "", "10FLAT"
sql = ""
sql = "select [DATE],ACNO,TIME,TRANSPORTER,KG,USER FROM DELIVERY  ORDER BY TIME"
Set RN2 = New ADODB.Recordset
RN2.Open sql, cn2
If Not RN2.EOF Then
txtSNo = Replace(RN2.Fields(1), "0", "")
txtQnty = RN2.Fields(4)
User = RN2.Fields(5)
Else
GoTo HEL
End If
HEL:
txtSNo_Validate True

If lblNames.Caption = "" Then
MsgBox "Please enter a valid supplier number."
txtSNo.SetFocus
Exit Sub
End If

If Not IsNumeric(txtQnty) Then
MsgBox "Please enter a number. " & txtQnty & " is not a number", vbExclamation
txtQnty.SetFocus
Exit Sub
End If

Startdate = DateSerial(year(DTPMilkDate), month(DTPMilkDate), 1)
Enddate = DateSerial(year(DTPMilkDate), month(DTPMilkDate) + 1, 1 - 1)

Set rs = oSaccoMaster.GetRecordset("d_sp_IsClosed '" & Enddate & "'")
If Not rs.EOF Then
    MsgBox "The period " & Enddate & " has been closed by " & rs.Fields(0)
    Exit Sub
End If

Dim ans As String
Set rs = New ADODB.Recordset
 sql = "SET dateformat dmy SELECT d_Milkintake.[ID],d_Milkintake.SNo, d_Milkintake.QSupplied, "
 sql = sql & " d_Milkintake.TransTime , d_Suppliers.[Names] FROM  d_Milkintake INNER JOIN "
 sql = sql & " d_Suppliers ON d_Milkintake.SNo = d_Suppliers.SNo AND  d_Milkintake.SNo= " & txtSNo & " AND (d_milkintake.TransDate = '" & DTPMilkDate & "')"
 sql = sql & " ORDER BY d_Milkintake.Id DESC"

 
Set rs = oSaccoMaster.GetRecordset(sql)
If rs.RecordCount > 0 Then

With lvwMilkCollected
    
       .ListItems.Clear
    
        .ColumnHeaders.Clear

  End With

    With lvwMilkCollected
        
        
        .ColumnHeaders.Add , , "SNo", 2000
        .ColumnHeaders.Add , , "Names", 3000
        .ColumnHeaders.Add , , "QNTY"
        .ColumnHeaders.Add , , "Time"
        .ColumnHeaders.Add , , "Receipt No.", 2000
    
        While Not rs.EOF
        
        If Not IsNull(rs.Fields("SNo")) Then
        
            Set li = .ListItems.Add(, , Trim(rs.Fields("SNo")))
            End If
            If Not IsNull(rs.Fields("Names")) Then
            li.ListSubItems.Add , , Trim(rs.Fields("Names"))
            End If
            If Not IsNull(rs.Fields("QSupplied")) Then
            li.ListSubItems.Add , , Trim(rs.Fields("QSupplied"))
            End If
            If Not IsNull(rs.Fields("TransTime")) Then
            li.ListSubItems.Add , , Trim(rs.Fields("TransTime"))
            End If
            If Not IsNull(rs.Fields("ID")) Then
             li.ListSubItems.Add , , Trim(rs.Fields("ID"))
            End If

            
                    rs.MoveNext
        
        Wend
        
    End With
    
    rs.Close
    
    Set rs = Nothing
    
lvwMilkCollected.View = lvwReport
ans = MsgBox("Supplier number " & txtSNo & " had supplied milk today. Add this?", vbYesNo, "MILK REPEAT")
If ans = vbNo Then
txtSNo.SetFocus
Exit Sub
End If
End If


Set rs = New ADODB.Recordset
sql = "SELECT Price from d_Price"
Set rs = oSaccoMaster.GetRecordset(sql)
If Not rs.EOF Then
Price = rs!Price
End If



'//Update Milk Intake


    
Set cn = New ADODB.Connection
sql = "d_sp_MilkIntake " & txtSNo & ",'" & DTPMilkDate & "'," & txtQnty & "," & Price & "," & Price * CCur(txtQnty) & ",'" & Time & "','" & User & "','Intake'"
oSaccoMaster.ExecuteThis (sql)




Set rs = New ADODB.Recordset
    sql = "d_sp_TransActive " & txtSNo & ",'" & DTPMilkDate & "'"
        Set rs = oSaccoMaster.GetRecordset(sql)
    If Not rs.EOF Then
    
        transdate = rs!rate
        If transdate < Startdate Then
            transdate = Startdate
        End If
        
        
        '---d_sp_SelTransDet @SNo bigint,@StartDate varchar(10), @Enddate varchar(10)
        Set rst = New ADODB.Recordset
            sql = "d_sp_SelTransDet " & txtSNo & ",'" & transdate & "','" & Enddate & "'"
        Set rst = oSaccoMaster.GetRecordset(sql)
        
    '--d_sp_UpdateDetTrans @SNo bigint,@QNTY float,@Amnt money,@Code varchar(35),@Subsidy money,@EPeriod varchar(10),@user varchar(35)
        If Not rst.EOF Then
            Set cn = New ADODB.Connection
                sql = "d_sp_UpdateDetTrans " & txtSNo & "," & rst!qnty & "," & rst!amount & ",'" & rst!code & "'," & rst!subsidy & ",'" & Enddate & "','" & User & "'"
            oSaccoMaster.ExecuteThis (sql)
            
        
        
        
        '----d_sp_SelTransGPayQnty @EP varchar(10), @Code varchar(35)
        Set rst2 = New ADODB.Recordset
            sql = "d_sp_SelTransGPayQnty '" & Enddate & "','" & rst!code & "'"
        Set rst2 = oSaccoMaster.GetRecordset(sql)
        
    '-- d_sp_UpdateTransPay @Code varchar(35), @Qnty float,@Amnt money,@Subsidy money,@GrossPay money, @EndDate varchar(10)  AS
    If Not rst2.EOF Then
            Set cn = New ADODB.Connection
                sql = "d_sp_UpdateTransPay '" & rst!code & "'," & rst2!qnty & "," & rst2!Amnt & "," & rst2!subsidy & "," & rst2!GPay & ",'" & Enddate & "','" & User & "'"
            oSaccoMaster.ExecuteThis (sql)
            
        End If
 
        '---Get transporters Total Deductions//d_sp_TotalTransDeduct @Code varchar(35),@Month bigint,@Year bigint
        Set Rs1 = New ADODB.Recordset
            sql = "d_sp_TotalTransDeduct '" & rst!code & "'," & month(DTPMilkDate) & "," & year(DTPMilkDate) & ""
        Set Rs1 = oSaccoMaster.GetRecordset(sql)
    If Not Rs1.EOF Then
    Dim TransTotalDed As Currency
    If Not IsNull(Rs1.Fields(0)) Then TransTotalDed = Rs1.Fields(0)
    End If
    
    Set rs3 = New ADODB.Recordset
    Dim DESCR As String
    Dim amount As Currency
    
    '--d_sp_SelTransDed @Code varchar(35), @startdate varchar(10),@enddate varchar(10)
    sql = "d_sp_SelTransDed '" & rst!code & "','" & Startdate & "','" & Enddate & "'"
    Set rs3 = oSaccoMaster.GetRecordset(sql)
    
    agrovet = 0
    AI = 0
    TMShares = 0
    FSA = 0
    HShares = 0
    Advance = 0
    Others = 0
    TransTotalDed = 0
    
    
If Not rs3.EOF Then
    While Not rs3.EOF
    DESCR = Trim(rs3.Fields(0))
    amount = 0
    amount = rs3.Fields(1)
    sql = "SELECT     Agrovet, AI, TMShares, FSA, HShares, Advance, Others FROM d_TransportersPayroll WHERE Code='" & rst!code & "' AND EndPeriod ='" & Enddate & "'"
    Set rs4 = oSaccoMaster.GetRecordset(sql)
     If UCase(rs4.Fields(0).name) = UCase(DESCR) Then
        agrovet = amount
    End If
    If UCase(rs4.Fields(1).name) = UCase(DESCR) Then
        AI = amount
    End If
    If UCase(rs4.Fields(2).name) = UCase(DESCR) Then
        TMShares = amount
    End If
    If UCase(rs4.Fields(3).name) = UCase(DESCR) Then
        FSA = amount
    End If
    If UCase(rs4.Fields(4).name) = UCase(DESCR) Then
        HShares = amount
    End If
    If UCase(rs4.Fields(5).name) = UCase(DESCR) Then
        Advance = amount
    End If
    If UCase(rs4.Fields(6).name) = UCase(DESCR) Then
        Others = amount
    End If

    '//d_sp_UpdatePayrollDed  @SNo bigint,@EndPeriod varchar(15),@Transport money,@Agrovet money,@AI money,@TMShares money,@FSA money,@HShares money,@Advance money,@Others
    rs3.MoveNext
    Wend
    
End If
     ' d_sp_UpdateTransDed  @Code varchar(35),@EndPeriod varchar(15),@TotalDed money,@Agrovet money,@AI money,@TMShares money,@FSA money,@HShares money,@Advance money,@Others money
   
    Set cn = New ADODB.Connection
    sql = "d_sp_UpdateTransDed  '" & rst!code & "','" & Enddate & "'," & TransTotalDed & "," & agrovet & "," & AI & "," & TMShares & "," & FSA & "," & HShares & "," & Advance & "," & Others & ""
    oSaccoMaster.ExecuteThis (sql)
    
 
        '//Update supplier deductions
            Set cn = New ADODB.Connection
            Dim r2  As String
            r2 = "  "
                sql = "d_sp_SupplierDeduct " & txtSNo & ",'" & DTPMilkDate & "','Transport'," & rs!rate * CCur(txtQnty) & ",'" & Startdate & "','" & Enddate & "'," & year(Enddate) & ",'" & User & "','" & r2 & "',''"
            oSaccoMaster.ExecuteThis (sql)
            



End If
'Check inactive Transporter
Set rs = New ADODB.Recordset
sql = "d_sp_TransInActive " & txtSNo & ",'" & Startdate & "'"
Set rs = oSaccoMaster.GetRecordset(sql)
If Not rs.EOF Then

'--Update Detailed Trasport
    transdate = rs!Startdate
        If transdate < Startdate Then
            transdate = Startdate
        End If
        
    '---d_sp_SelTransDet @SNo bigint,@StartDate varchar(10), @Enddate varchar(10)
        Set rst = New ADODB.Recordset
            sql = "d_sp_SelTransDet " & txtSNo & ",'" & transdate & "','" & rs!dateinactivate & "'"
        Set rst = oSaccoMaster.GetRecordset(sql)
        
    '--d_sp_UpdateDetTrans @SNo bigint,@QNTY float,@Amnt money,@Code varchar(35),@Subsidy money,@EPeriod varchar(10),@user varchar(35)
        If Not rst.EOF Then
            Set cn = New ADODB.Connection
                sql = "d_sp_UpdateDetTrans " & txtSNo & "," & rst!qnty & "," & rst!amount & ",'" & rst!code & "'," & rst!subsidy & ",'" & Enddate & "','" & User & "'"
            oSaccoMaster.ExecuteThis (sql)
            
            '----d_sp_SelTransGPayQnty @EP varchar(10), @Code varchar(35)
        Set rst2 = New ADODB.Recordset
            sql = "d_sp_SelTransGPayQnty '" & Enddate & "','" & rst!code & "'"
        Set rst2 = oSaccoMaster.GetRecordset(sql)
            
        
        
        
        
        
    '-- d_sp_UpdateTransPay @Code varchar(35), @Qnty float,@Amnt money,@Subsidy money,@GrossPay money, @EndDate varchar(10)  AS
    If Not rst2.EOF Then
            Set cn = New ADODB.Connection
                sql = "d_sp_UpdateTransPay '" & rst!code & "'," & rst2!qnty & "," & rst2!Amnt & "," & rst2!subsidy & "," & rst2!GPay & ",'" & Enddate & "','" & User & "'"
            oSaccoMaster.ExecuteThis (sql)
            
        End If
     
        '---Get transporters Total Deductions//d_sp_TotalTransDeduct @Code varchar(35),@Month bigint,@Year bigint
        Set Rs1 = New ADODB.Recordset
            sql = "d_sp_TotalTransDeduct '" & rst!code & "'," & month(DTPMilkDate) & "," & year(DTPMilkDate) & ""
        Set Rs1 = oSaccoMaster.GetRecordset(sql)
    If Not Rs1.EOF Then
    'Dim TransTotalDed As Currency
    If Not IsNull(Rs1.Fields(0)) Then TransTotalDed = Rs1.Fields(0)
    End If
    
    Set rs3 = New ADODB.Recordset
    'Dim Startdate As String, Enddate As String
    'Dim DESCR As String
   ' Dim amount As Currency
    
    '--d_sp_SelTransDed @Code varchar(35), @startdate varchar(10),@enddate varchar(10)
    sql = "d_sp_SelTransDed '" & rst!code & "','" & Startdate & "','" & Enddate & "'"
    Set rs3 = oSaccoMaster.GetRecordset(sql)
    
    agrovet = 0
    AI = 0
    TMShares = 0
    FSA = 0
    HShares = 0
    Advance = 0
    Others = 0
    TransTotalDed = 0
If Not rs3.EOF Then
    While Not rs3.EOF
    DESCR = Trim(rs3.Fields(0))
    amount = 0
    amount = rs3.Fields(1)
    sql = "SELECT     Agrovet, AI, TMShares, FSA, HShares, Advance, Others FROM d_TransportersPayroll WHERE Code='" & rst!code & "' AND EndPeriod ='" & Enddate & "'"
    Set rs4 = oSaccoMaster.GetRecordset(sql)
     If UCase(rs4.Fields(0).name) = UCase(DESCR) Then
        agrovet = amount
    End If
    If UCase(rs4.Fields(1).name) = UCase(DESCR) Then
        AI = amount
    End If
    If UCase(rs4.Fields(2).name) = UCase(DESCR) Then
        TMShares = amount
    End If
    If UCase(rs4.Fields(3).name) = UCase(DESCR) Then
        FSA = amount
    End If
    If UCase(rs4.Fields(4).name) = UCase(DESCR) Then
        HShares = amount
    End If
    If UCase(rs4.Fields(5).name) = UCase(DESCR) Then
        Advance = amount
    End If
    If UCase(rs4.Fields(6).name) = UCase(DESCR) Then
        Others = amount
    End If

    '//d_sp_UpdatePayrollDed  @SNo bigint,@EndPeriod varchar(15),@Transport money,@Agrovet money,@AI money,@TMShares money,@FSA money,@HShares money,@Advance money,@Others
    rs3.MoveNext
    Wend
    ' d_sp_UpdateTransDed  @Code varchar(35),@EndPeriod varchar(15),@TotalDed money,@Agrovet money,@AI money,@TMShares money,@FSA money,@HShares money,@Advance money,@Others money
   
End If
 Set cn = New ADODB.Connection
    sql = "d_sp_UpdateTransDed  '" & rst!code & "','" & Enddate & "'," & TransTotalDed & "," & agrovet & "," & AI & "," & TMShares & "," & FSA & "," & HShares & "," & Advance & "," & Others & ""
    oSaccoMaster.ExecuteThis (sql)
    End If

End If

 
End If



'//d_sp_TotalDeduct-Total Deductions
'//d_sp_UpdateGPAYQnty - Total Grosspay and Quantity
'//d_sp_SupDed - Supply Deductions



Set rs2 = New ADODB.Recordset
'Dim Startdate As String, Enddate As String
Dim qnty As Currency, GPay As Currency
'Startdate = DateSerial(DTPMilkDate, cboMonth, 1)
sql = "d_sp_UpdateGPAYQnty '" & Startdate & "','" & Enddate & "'," & txtSNo & ""
Set rs2 = oSaccoMaster.GetRecordset(sql)
If Not rs2.EOF Then
If Not IsNull(rs2.Fields(0)) Then qnty = rs2.Fields(0)
If Not IsNull(rs2.Fields(1)) Then GPay = rs2.Fields(1)
End If


Set Rs1 = New ADODB.Recordset
sql = "d_sp_TotalDeduct " & txtSNo & "," & month(DTPMilkDate) & "," & year(DTPMilkDate) & ""
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
Startdate = DateSerial(year(DTPMilkDate), month(DTPMilkDate), 1)
Enddate = DateSerial(year(DTPMilkDate), month(DTPMilkDate) + 1, 1 - 1)
sql = "d_sp_SupDed " & txtSNo & ",'" & Startdate & "','" & Enddate & "'"
Set rs3 = oSaccoMaster.GetRecordset(sql)
If Not rs3.EOF Then
While Not rs3.EOF
desc = Trim(rs3.Fields(0))
Amnt = 0
Amnt = rs3.Fields(1)
sql = "SELECT     Transport, Agrovet, AI, TMShares, FSA, HShares, Advance, Others FROM d_Payroll WHERE SNo=" & txtSNo & " AND EndofPeriod ='" & Enddate & "'"
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

Transport = 0
agrovet = 0
AI = 0
TMShares = 0
FSA = 0
HShares = 0
Advance = 0
Others = 0
'//Print Receipt
    If chkPrint = vbChecked Then
    
'/*Print out
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
       'cdgPrint.Filter = "*.csv|*.txt"
        'cdgPrint.ShowSave
        Dim PORT As String
        PORT = ports
        'ttt = "LPT1" 'LPT1
        ttt = PORT
        'ttt = "LPT1"
'        ttt = cdgPrint.PrinterDefault
        'ttt = "D:\PROJECTS\FOSA\DAILY" & Date & ""
        Set fso = CreateObject("Scripting.FileSystemObject")
        'Set chkPrinter = fso.GetFile(ttt)
       
        
        Set txtFile = fso.CreateTextFile(ttt, True)
    txtFile.WriteLine "         " & cname & ""
    txtFile.WriteLine "            Milk Receipt"
    txtFile.WriteLine "---------------------------------------"
        
    Set rs2 = New ADODB.Recordset
    sql = "d_sp_ReceiptNumber"
    Set rs2 = oSaccoMaster.GetRecordset(sql)
    
    Dim RNumber As String
    'RNumber = rs2.Fields(0)
    If Not IsNull(rs2.Fields(0)) Then RNumber = rs2.Fields(0)
'    Else
'    RNumber = "0"
'    End If
    
    txtFile.WriteLine "SNo :" & txtSNo
    txtFile.WriteLine "Name :" & lblNames
    txtFile.WriteLine "Quantity Supplied :" & txtQnty & " Kgs"
    Set rs = New ADODB.Recordset
    sql = "d_sp_TotalMonth " & txtSNo & ",'" & Startdate & "','" & DTPMilkDate & "'"
    Set rs = oSaccoMaster.GetRecordset(sql)
    If Not rs.EOF Then
    If Not IsNull(rs.Fields(0)) Then CummulKgs = rs.Fields(0)
    Else
    CummulKgs = "0.00"
    End If
    txtFile.WriteLine "Cummulative This Month : " & Format(CummulKgs, "#,##0.00" & " Kgs")
    Set rs = New ADODB.Recordset
    sql = "d_sp_TransName '" & txtSNo & "'"
    Set rs = oSaccoMaster.GetRecordset(sql)
    If Not rs.EOF Then
    If Not IsNull(rs.Fields(0)) Then TRANSPORTER = rs.Fields(0)
    Else
        TRANSPORTER = "Self"
    End If
    txtFile.WriteLine "---------------------------------------"
    txtFile.WriteLine "Receipt Number :" & RNumber
    txtFile.WriteLine "TRANSPORTER :" & TRANSPORTER
    txtFile.WriteLine "Received by :" & username
    txtFile.WriteLine "---------------------------------------"
    txtFile.WriteLine "  Date :" & Format(DTPMilkDate, "dd/mm/yyyy") & " ,Time : " & Format(Time, "hh:mm:ss AM/PM")
    txtFile.WriteLine "         " & motto & ""
    txtFile.WriteLine "---------------------------------------"
    If chkComment.value = vbChecked Then
        txtFile.WriteLine txtComment
        txtFile.WriteLine "---------------------------------------"
    End If
    txtFile.WriteLine escFeedAndCut
    
 txtFile.Close
 Reset
End If


'* writing to notepad
If chkNotepad.value = vbChecked Then

'    Dim fso, txtfile
'    Dim ttt
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
    txtFile.WriteLine "Name :" & lblNames
'    Else
    txtFile.WriteLine "Quantity Supplied :" & txtQnty & " Kgs"
    Startdate = DateSerial(year(DTPMilkDate), month(DTPMilkDate) - 1, 1)
    'sql = "d_sp_TotalMonth " & txtSNo & ",'" & StartDate & "','" & DTPMilkDate & "'"
    Set rs = New ADODB.Recordset
    sql = "d_sp_TotalMonth " & txtSNo & ",'" & Startdate & "','" & DTPMilkDate & "'"
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
    'Printer.Print
    'Printer.CurrentX = 500#
    'Printer.FontSize = 10
    'Printer.CurrentX = 500
    'Printer.FontSize = 8
'    Printer.Print Tab(0); "Date :"; Tab(10); Format(Get_Server_Date, "dd/mm/yyyy HH:MM:SS AM/PM")
'    Printer.Print Tab(0); "     TDL - Improves Your Value "
'    Printer.Print Tab(0); "---------------------------------------"
'    Printer.CurrentX = 500
'    Printer.FontSize = 8
'    Printer.Print
'    Printer.CurrentX = 500
'    Printer.FontSize = 8
'    Printer.Print
txtFile.Close
End If

loadMilk

txtSNo = ""
txtQnty = ""
'txtSNo_Validate True
txtSNo.SetFocus
Exit Sub
ErrorHandler:

MsgBox err.description




End Sub

Private Sub cmdPrint_Click()
On Error GoTo ErrorHandler
If txtTCode = "" Then
MsgBox "Please enter the transporter code.", vbInformation
txtTCode.SetFocus
Exit Sub
End If

txtTCode_Validate True

If lblTName.Caption = "" Then
MsgBox "Please enter code for a valid transporter. Transporter with code " & txtTCode & " does not exist.", vbInformation
txtTCode.SetFocus
Exit Sub
End If

'//Print Receipt
    Set rs = oSaccoMaster.GetRecordset("d_sp_TransDTotal '" & txtTCode & "','" & DTPMilkDate & "'")
    If IsNull(rs.Fields(0)) Then
        MsgBox "There is no milk supplied by Code " & txtTCode
        txtTCode.SetFocus
        Exit Sub
    End If
    
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
       
  '      ttt = "LPT1" 'LPT1
         ttt = ports
        
        Set fso = CreateObject("Scripting.FileSystemObject")

        
        Set txtFile = fso.CreateTextFile(ttt, True)
    txtFile.WriteLine "             " & cname & ""
    txtFile.WriteLine "           Transporter's Receipt"
    txtFile.WriteLine "---------------------------------------"
    
    txtFile.WriteLine "Transporter Code :" & txtTCode
    txtFile.WriteLine "Transporter Name :" & lblTName
    
    txtFile.WriteLine "Quantity Transported :" & Format(CCur(rs.Fields(0)), "#,##0.00") & " Kgs "
    
    Set rs = oSaccoMaster.GetRecordset("d_sp_TransTotal '" & txtTCode & "'," & month(DTPMilkDate) & ", " & year(DTPMilkDate))
    CummulKgs = 0
    If Not IsNull(rs.Fields(0)) Then CummulKgs = rs.Fields(0)
   
    txtFile.WriteLine "Cummulative This Month " & Format(CummulKgs, "#,##0.00" & " Kgs")
    txtFile.WriteLine "---------------------------------------"
    txtFile.WriteLine " Printed By " & UCase(username)
    txtFile.WriteLine "    Date :" & Format(Get_Server_Date, "dd/mm/yyyy HH:MM:SS AM/PM")
    txtFile.WriteLine "     " & motto & ""
    txtFile.WriteLine "****************************************"
    txtFile.WriteLine escFeedAndCut
    txtFile.Close
    Exit Sub
ErrorHandler:
    MsgBox err.description
    
End Sub

Private Sub cmdprintlist_Click()
'On Error GoTo errorhandler
'If txtTCode = "" Then
'MsgBox "Please enter the transporter code.", vbInformation
'txtTCode.SetFocus
'Exit Sub
'End If
'
'txtTCode_Validate True
'
'If lblTName.Caption = "" Then
'MsgBox "Please enter code for a valid transporter. Transporter with code " & txtTCode & " does not exist.", vbInformation
'txtTCode.SetFocus
'Exit Sub
'End If
'
''//Print Receipt
'    Set rs = oSaccoMaster.GetRecordset("d_sp_TransDTotal '" & txtTCode & "','" & DTPMilkDate & "'")
'    If IsNull(rs.Fields(0)) Then
'        MsgBox "There is no milk supplied by Code " & txtTCode
'        txtTCode.SetFocus
'        Exit Sub
'    End If
'
'     Dim fso, chkPrinter, txtfile
'    Dim ttt
'     Dim escFeedAndCut As String
'     Dim escNewLine As String
'     Dim escUnerLineON As String
'     Dim escUnerLineOnX2 As String
'     Dim escUnerLineOff As String
'     Dim escBoldOn As String
'     Dim escBoldOff As String
'     Dim escNegativeOn As String
'     Dim escNegativeOff As String
'     Dim esc8CpiOn As String
'     Dim esc8CPiOff As String
'     Dim esc16Cpi As String
'     Dim esc20Cpi As String
'     Dim escAlignLeft As String
'     Dim escAlignCenter As String
'     Dim escAlignRight As String
'
'
'        escNewLine = Chr(10) '//New Line (LF Line Feed)
'        escUnerLineON = Chr(27) + Chr(45) + Chr(1) '//Unerline on
'        escUnerLineOnX2 = Chr(27) + Chr(45) + Chr(1) '//Unerline on X2
'        escUnerLineOff = Chr(27) + Chr(45) + Chr(0) '//unerline off
'        escBoldOn = Chr(27) + Chr(69) + Chr(1) '//Bold on
'        escBoldOff = Chr(27) + Chr(69) + Chr(0) '//Bold off
'        escNegativeOn = Chr(29) + Chr(66) + Chr(1) '//White on Black on
'        escNegativeOff = Chr(29) + Chr(66) + Chr(0) '//white on
'        esc8CpiOn = Chr(29) + Chr(33) + Chr(16) '//Font Size X2 on
'        esc8CPiOff = Chr(29) + Chr(33) + Chr(0) '//Font size X2 off
'        esc16Cpi = Chr(27) + Chr(77) + Chr(48) '//Font A - Normal Size
'        esc20Cpi = Chr(27) + Chr(77) + Chr(49) '//Font B - Small Font
'        escAlignLeft = Chr(27) + Chr(97) + Chr(48) '//Align text to the left
'        escAlignCenter = Chr(27) + Chr(97) + Chr(49) '//Align text to the center
'        escAlignRight = Chr(27) + Chr(97) + Chr(50) '//Align text to the right
'        escFeedAndCut = Chr(29) + Chr(86) + Chr(65) '//Partial cut and feed
'
'        'ttt = "LPT1" 'LPT1
'
'        Dim PORT As String
'   '     PORT = ports
'   '     ttt = PORT
'        ttt = ports
'        'ttt = "LPT1" 'LPT1
'        Set fso = CreateObject("Scripting.FileSystemObject")
'
'
'    Set txtfile = fso.CreateTextFile(ttt, True)
'    txtfile.WriteLine " " & cname & ""
'    txtfile.WriteLine " Transporter's Receipt"
'    txtfile.WriteLine "------------------------------------------------"
'
'    txtfile.WriteLine "Transporter Code :" & txtTCode
'    txtfile.WriteLine "Transporter Name :" & lblTName
'
'    txtfile.WriteLine "Quantity Transported :" & Format(CCur(rs.Fields(0)), "#,##0.00") & " Kgs "
'    txtfile.WriteLine "***************************************************" 'START
'    txtfile.WriteLine "SNO" & vbTab & vbTab & "NAME" & vbTab & vbTab & vbTab & "KGS"
'    txtfile.WriteLine "------------------------------------------------"
'    txtfile.WriteLine "****************************************" 'START
'    Dim rsTrans As New ADODB.Recordset, sno As Long, qs As Double, name As String
'            sql = "set dateformat dmy SELECT     d_Transport.Trans_Code, d_Milkintake.SNo, d_Milkintake.TransDate, d_Milkintake.QSupplied, d_Transporters.TransName, d_Suppliers.[Names]"
'            sql = sql & " FROM         d_Transport d_Transport INNER JOIN"
'            sql = sql & " d_Transporters d_Transporters ON d_Transport.Trans_Code = d_Transporters.TransCode INNER JOIN"
'            sql = sql & "  d_Milkintake d_Milkintake ON d_Transport.Sno = d_Milkintake.SNo INNER JOIN"
'            sql = sql & " d_Suppliers d_Suppliers ON d_Milkintake.SNo = d_Suppliers.SNo"
'            sql = sql & " WHERE     d_Milkintake.TransDate ='" & DTPMilkDate & "'"
'            sql = sql & "  AND d_Transport.Trans_Code = '" & txtTCode & "' AND d_Transport.active =1  "
'            sql = sql & "ORDER BY  d_Milkintake.SNo"
'            Set rsTrans = oSaccoMaster.GetRecordset(sql)
'            While Not rsTrans.EOF
'            sno = IIf(IsNull(rsTrans.Fields(1)), "NA", rsTrans.Fields(1))
'            qs = IIf(IsNull(rsTrans.Fields(3)), 0, rsTrans.Fields(3))
'            name = IIf(IsNull(rsTrans.Fields(5)), "NA", rsTrans.Fields(5))
'    txtfile.WriteLine sno & vbTab & name & vbTab & vbTab & vbTab & qs
'            rsTrans.MoveNext
'            Wend
'    txtfile.WriteLine "****************************************" 'END
'    Set rs = oSaccoMaster.GetRecordset("d_sp_TransTotal '" & txtTCode & "'," & month(DTPMilkDate) & ", " & year(DTPMilkDate))
'    CummulKgs = 0
'    If Not IsNull(rs.Fields(0)) Then CummulKgs = rs.Fields(0)
'
'    txtfile.WriteLine "Cummulative This Month " & Format(CummulKgs, "#,##0.00" & " Kgs")
'    txtfile.WriteLine "---------------------------------------"
'    txtfile.WriteLine " Printed By " & UCase(username)
'    txtfile.WriteLine "    Date :" & Format(Get_Server_Date, "dd/mm/yyyy HH:MM:SS AM/PM")
'    txtfile.WriteLine "     " & motto & ""
'    txtfile.WriteLine "****************************************"
'    txtfile.WriteLine escFeedAndCut
'    txtfile.Close
'Exit Sub
'errorhandler:
'MsgBox err.description
On Error GoTo ErrorHandler
If txtTCode = "" Then
MsgBox "Please enter the transporter code.", vbInformation
txtTCode.SetFocus
Exit Sub
End If

txtTCode_Validate True

If lblTName.Caption = "" Then
MsgBox "Please enter code for a valid transporter. Transporter with code " & txtTCode & " does not exist.", vbInformation
txtTCode.SetFocus
Exit Sub
End If

'//Print Receipt
    Set rs = oSaccoMaster.GetRecordset("d_sp_TransDTotal '" & txtTCode & "','" & DTPMilkDate & "'")
    If IsNull(rs.Fields(0)) Then
        MsgBox "There is no milk supplied by Code " & txtTCode
        txtTCode.SetFocus
        Exit Sub
    End If
    
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
       
        'ttt = "LPT1" 'LPT1
        
         Dim PORT As String
        PORT = ports
        'ttt = "LPT1" 'LPT1
        ttt = PORT
        Set fso = CreateObject("Scripting.FileSystemObject")

        
    Set txtFile = fso.CreateTextFile(ttt, True)
    txtFile.WriteLine " " & cname & ""
    txtFile.WriteLine " Transporter's Receipt"
    txtFile.WriteLine "------------------------------------------------"
    
    txtFile.WriteLine "Transporter Code :" & txtTCode
    txtFile.WriteLine "Transporter Name :" & lblTName
    
    txtFile.WriteLine "Quantity Transported :" & Format(CCur(rs.Fields(0)), "#,##0.00") & " Kgs "
    txtFile.WriteLine "***************************************************" 'START
    txtFile.WriteLine "SNO" & vbTab & vbTab & "NAME" & vbTab & vbTab & vbTab & "KGS"
    txtFile.WriteLine "------------------------------------------------"
    txtFile.WriteLine "****************************************" 'START
    Dim rsTrans As New ADODB.Recordset, sno As Long, qs As Double, name As String
            sql = "set dateformat dmy SELECT     d_Transport.Trans_Code, d_Milkintake.SNo, d_Milkintake.TransDate, d_Milkintake.QSupplied, d_Transporters.TransName, d_Suppliers.[Names]"
            sql = sql & " FROM         d_Transport d_Transport INNER JOIN"
            sql = sql & " d_Transporters d_Transporters ON d_Transport.Trans_Code = d_Transporters.TransCode INNER JOIN"
            sql = sql & "  d_Milkintake d_Milkintake ON d_Transport.Sno = d_Milkintake.SNo INNER JOIN"
            sql = sql & " d_Suppliers d_Suppliers ON d_Milkintake.SNo = d_Suppliers.SNo"
            sql = sql & " WHERE     d_Milkintake.TransDate ='" & DTPMilkDate & "'"
            sql = sql & "  AND d_Transport.Trans_Code = '" & txtTCode & "' AND d_Transport.active =1  "
            sql = sql & "ORDER BY  d_Milkintake.SNo"
            Set rsTrans = oSaccoMaster.GetRecordset(sql)
            While Not rsTrans.EOF
            sno = IIf(IsNull(rsTrans.Fields(1)), "NA", rsTrans.Fields(1))
            qs = IIf(IsNull(rsTrans.Fields(3)), 0, rsTrans.Fields(3))
            name = IIf(IsNull(rsTrans.Fields(5)), "NA", rsTrans.Fields(5))
    txtFile.WriteLine sno & vbTab & name & vbTab & vbTab & vbTab & qs
            rsTrans.MoveNext
            Wend
    txtFile.WriteLine "****************************************" 'END
    Set rs = oSaccoMaster.GetRecordset("d_sp_TransTotal '" & txtTCode & "'," & month(DTPMilkDate) & ", " & year(DTPMilkDate))
    CummulKgs = 0
    If Not IsNull(rs.Fields(0)) Then CummulKgs = rs.Fields(0)
   
    txtFile.WriteLine "Cummulative This Month " & Format(CummulKgs, "#,##0.00" & " Kgs")
    txtFile.WriteLine "---------------------------------------"
    txtFile.WriteLine " Printed By " & UCase(username)
    txtFile.WriteLine "    Date :" & Format(Get_Server_Date, "dd/mm/yyyy HH:MM:SS AM/PM")
    txtFile.WriteLine "     " & motto & ""
    txtFile.WriteLine "****************************************"
    txtFile.WriteLine escFeedAndCut
    txtFile.Close
Exit Sub
ErrorHandler:
MsgBox err.description
End Sub

Private Sub cmdprintself_Click()
'get the members who are self

Set rs = oSaccoMaster.GetRecordset("SELECT DISTINCT Sno  FROM         d_Transport  where active=1  ORDER BY Sno")
While Not rs.EOF

Set rst = oSaccoMaster.GetRecordset("SELECT * FROM D_SUPPLIERS WHERE SNO='" & rs.Fields(0) & "' AND HAST=1 ")
If rst.EOF Then
sql = ""
sql = "update d_suppliers set hast=1 where sno='" & Trim(rs.Fields(0)) & "'"
oSaccoMaster.ExecuteThis (sql)
End If
rs.MoveNext
Wend
'//update also the one that have been deactivated
Set rs = oSaccoMaster.GetRecordset("SELECT DISTINCT Sno  FROM         d_Transport  where active=0  ORDER BY Sno")
While Not rs.EOF

Set rst = oSaccoMaster.GetRecordset("SELECT * FROM D_SUPPLIERS WHERE SNO='" & rs.Fields(0) & "' AND HAST=1 ")
If rst.EOF Then
sql = ""
sql = "update d_suppliers set hast=1 where sno='" & Trim(rs.Fields(0)) & "'"
oSaccoMaster.ExecuteThis (sql)
Else
sql = ""
sql = "update d_suppliers set hast=0 where sno='" & Trim(rs.Fields(0)) & "'"
oSaccoMaster.ExecuteThis (sql)
End If
rs.MoveNext
Wend
'self

 reportname = "self.rpt"

 
 Show_Sales_Crystal_Report STRFORMULA, reportname, ""
End Sub

Private Sub cmdReceive_Click()

On Error GoTo ErrorHandler
Dim Price As Currency
Dim Startdate, CummulKgs, TRANSPORTER As String
Dim transdate As Date
'check the user
sql = "SELECT     UserLoginID,levels, UserGroup, SUPERUSER From UserAccounts where UserLoginID='" & User & "'"
Set rs = oSaccoMaster.GetRecordset(sql)
If Not rs.EOF Then
If rs!Levels <> "Intake" Then
MsgBox "You are not allowed to Intake", vbInformation
Exit Sub

End If
End If
If Trim(txtSNo) = "" Then
    MsgBox "Please enter the supplier number."
        txtSNo.SetFocus
    Exit Sub
End If

If Trim(txtQnty) = "" Then
    MsgBox "Please enter the quantity supplied."
        txtQnty.SetFocus
Exit Sub

End If

txtSNo_Validate True

If Trim(lblNames.Caption) = "" Then
MsgBox "Please enter a valid supplier number."
txtSNo.SetFocus
Exit Sub
End If

If Not IsNumeric(txtQnty) Then
MsgBox "Please enter a number. " & txtQnty & " is not a number", vbExclamation
txtQnty.SetFocus
Exit Sub
End If

Startdate = DateSerial(year(DTPMilkDate), month(DTPMilkDate), 1)
Enddate = DateSerial(year(DTPMilkDate), month(DTPMilkDate) + 1, 1 - 1)
DTPMilkDate = Format(DTPMilkDate, "DD/MM/YYYY hh:mm:ss")

Set rs = oSaccoMaster.GetRecordset("d_sp_IsClosed '" & Enddate & "'")
If Not rs.EOF Then
    MsgBox "The period " & Enddate & " has been closed by " & rs.Fields(0)
    Exit Sub
End If

Dim ans As String
Set rs = New ADODB.Recordset
 sql = "SET dateformat dmy SELECT d_Milkintake.[ID],d_Milkintake.SNo, d_Milkintake.QSupplied, "
 sql = sql & " d_Milkintake.TransTime , d_Suppliers.[Names] FROM  d_Milkintake INNER JOIN "
 sql = sql & " d_Suppliers ON d_Milkintake.SNo = d_Suppliers.SNo AND  d_Milkintake.SNo= " & txtSNo & " AND (d_milkintake.TransDate = '" & DTPMilkDate & "')"
 sql = sql & " ORDER BY d_Milkintake.Id DESC"

 
Set rs = oSaccoMaster.GetRecordset(sql)
If rs.RecordCount > 0 Then

With lvwMilkCollected
    
       .ListItems.Clear
    
        .ColumnHeaders.Clear

  End With

    With lvwMilkCollected
        
        
        .ColumnHeaders.Add , , "SNo", 2000
        .ColumnHeaders.Add , , "Names", 3000
        .ColumnHeaders.Add , , "QNTY"
        .ColumnHeaders.Add , , "Time"
        .ColumnHeaders.Add , , "Receipt No.", 2000
    
        While Not rs.EOF
        
        If Not IsNull(rs.Fields("SNo")) Then
        
            Set li = .ListItems.Add(, , Trim(rs.Fields("SNo")))
            End If
            If Not IsNull(rs.Fields("Names")) Then
            li.ListSubItems.Add , , Trim(rs.Fields("Names"))
            End If
            If Not IsNull(rs.Fields("QSupplied")) Then
            li.ListSubItems.Add , , Trim(rs.Fields("QSupplied"))
            End If
            If Not IsNull(rs.Fields("TransTime")) Then
            li.ListSubItems.Add , , Trim(rs.Fields("TransTime"))
            End If
            If Not IsNull(rs.Fields("ID")) Then
             li.ListSubItems.Add , , Trim(rs.Fields("ID"))
            End If

            
                    rs.MoveNext
        
        Wend
        
    End With
    
    rs.Close
    
    Set rs = Nothing
    
lvwMilkCollected.View = lvwReport
ans = MsgBox("Supplier number " & txtSNo & " had supplied milk today. Add this?", vbYesNo, "MILK REPEAT")
If ans = vbNo Then
txtSNo.SetFocus
Exit Sub
End If
End If
Dim rsrepeat As New ADODB.Recordset
Dim nnn As Integer
Set rsrepeat = oSaccoMaster.GetRecordset("SET              dateformat dmy   SELECT     COUNT(*)   FROM         d_Milkintake WHERE     transdate = '" & DTPMilkDate & "' AND sno = " & txtSNo & "")
If Not rsrepeat.EOF Then
If CLng(rsrepeat.Fields(0)) >= 1 Then
MsgBox "You have supplied twice today. You are not allowed to supply for the third time in a day", vbCritical
Exit Sub
End If
End If


Set rs = New ADODB.Recordset
sql = "SELECT Price from d_Price"
Set rs = oSaccoMaster.GetRecordset(sql)
If Not rs.EOF Then
Price = rs!Price
End If
'//Update trend analysis
'//how do we add =+- 5

Dim rstrend As New ADODB.Recordset
Dim rstrend1 As New ADODB.Recordset
Dim avgkg As Double, ag As Double, fd As Integer, cd As Date, fdate As Date, rt As New ADODB.Recordset, ala As Integer, mornin As String, FINADATE As String

Dim rrt As New ADODB.Recordset, qs As Double
cd = DTPMilkDate
mornin = " 00:00:00 AM"
FINADATE = cd & mornin
fdate = CDate(CLng(DTPMilkDate) - 5)
DTPMilkDate = (Get_Server_Date)
avgkg = 0
sql = ""
sql = "set dateformat dmy SELECT      *   FROM         d_Milkintake   WHERE     (SNo = " & txtSNo & ") and transdate>='" & fdate & "' and transdate<>'" & DTPMilkDate & "'    ORDER BY TransDate DESC, ID DESC"
Set rstrend = oSaccoMaster.GetRecordset(sql)
If rstrend.RecordCount < 5 Then GoTo ebrahim
While Not rstrend.EOF

avgkg = avgkg + IIf(IsNull(rstrend.Fields("qsupplied")), 0, rstrend.Fields("qsupplied"))

rstrend.MoveNext
Wend
'//compare with the actual supplied
ag = avgkg / 5
'//get to know if he has supplied the milk before on the same date
sql = ""
sql = "set dateformat dmy SELECT      sum(QSupplied)   FROM         d_Milkintake   WHERE     (SNo = " & txtSNo & ") and transdate='" & DTPMilkDate & "'"
Set rt = oSaccoMaster.GetRecordset(sql)
If Not rt.EOF Then
qs = IIf(IsNull(rt.Fields(0)), 0, rt.Fields(0))
qs = qs + CDbl(txtQnty)
ala = 1
Else
ala = 0
End If

sql = ""
sql = "SET              dateformat dmy  SELECT     *  FROM         d_tbl_trends  WHERE     auditdatetime >= '" & FINADATE & "' AND sno = " & txtSNo & ""
Set rrt = oSaccoMaster.GetRecordset(sql)
If Not rrt.EOF Then
sql = ""
sql = "delete from d_tbl_trends WHERE     auditdatetime >= '" & FINADATE & "' AND sno = " & txtSNo & " "
oSaccoMaster.ExecuteThis (sql)

If ag > CDbl(txtQnty) + 5 Then
sql = "insert_d_tbl_trends  '" & txtSNo & "'," & qs & "," & ag & ",'" & User & "'"
oSaccoMaster.ExecuteThis (sql)
End If
If ag < CDbl(txtQnty) - 5 Then
sql = "insert_d_tbl_trends  '" & txtSNo & "'," & qs & "," & ag & ",'" & User & "'"
oSaccoMaster.ExecuteThis (sql)
End If
'//sum the total maziwa to be inputed to the trend table.

Else

If ag > CDbl(txtQnty) + 5 Then
sql = "insert_d_tbl_trends  '" & txtSNo & "'," & txtQnty & "," & ag & ",'" & User & "'"
oSaccoMaster.ExecuteThis (sql)
End If
If ag < CDbl(txtQnty) - 5 Then
sql = "insert_d_tbl_trends  '" & txtSNo & "'," & txtQnty & "," & ag & ",'" & User & "'"
oSaccoMaster.ExecuteThis (sql)
End If
End If




ebrahim:


'//Update Milk Intake

'DTPMilkDate = Format(DTPMilkDate, "DD/MM/YYYY hh:mm:ss")
DTPMilkDate = Format(Get_Server_Date, "DD/MM/YYYY hh:mm:ss")
Set cn = New ADODB.Connection
sql = "d_sp_MilkIntake " & txtSNo & ",'" & DTPMilkDate & "'," & txtQnty & "," & Price & "," & Price * CCur(txtQnty) & ",'" & Time & "','" & User & "','Intake','" & txtComment & "'"
oSaccoMaster.ExecuteThis (sql)



'//Check Transporter if active

Set rs = New ADODB.Recordset
    sql = "d_sp_TransActive " & txtSNo & ",'" & DTPMilkDate & "'"
        Set rs = oSaccoMaster.GetRecordset(sql)
    If Not rs.EOF Then

        transdate = rs!rate
        If transdate < Startdate Then
            transdate = Startdate
        End If


        '---d_sp_SelTransDet @SNo bigint,@StartDate varchar(10), @Enddate varchar(10)
        Set rst = New ADODB.Recordset
            sql = "d_sp_SelTransDet " & txtSNo & ",'" & transdate & "','" & Enddate & "'"
        Set rst = oSaccoMaster.GetRecordset(sql)

    '--d_sp_UpdateDetTrans @SNo bigint,@QNTY float,@Amnt money,@Code varchar(35),@Subsidy money,@EPeriod varchar(10),@user varchar(35)
        If Not rst.EOF Then
            Set cn = New ADODB.Connection
                sql = "d_sp_UpdateDetTrans " & txtSNo & "," & rst!qnty & "," & rst!amount & ",'" & rst!code & "'," & rst!subsidy & ",'" & Enddate & "','" & User & "'"
            oSaccoMaster.ExecuteThis (sql)




        '----d_sp_SelTransGPayQnty @EP varchar(10), @Code varchar(35)
        Set rst2 = New ADODB.Recordset
            sql = "d_sp_SelTransGPayQnty '" & Enddate & "','" & rst!code & "'"
        Set rst2 = oSaccoMaster.GetRecordset(sql)

    '-- d_sp_UpdateTransPay @Code varchar(35), @Qnty float,@Amnt money,@Subsidy money,@GrossPay money, @EndDate varchar(10)  AS
    If Not rst2.EOF Then
            Set cn = New ADODB.Connection
                sql = "d_sp_UpdateTransPay '" & rst!code & "'," & rst2!qnty & "," & rst2!Amnt & "," & rst2!subsidy & "," & rst2!GPay & ",'" & Enddate & "','" & User & "'"
            oSaccoMaster.ExecuteThis (sql)

        End If

        '---Get transporters Total Deductions//d_sp_TotalTransDeduct @Code varchar(35),@Month bigint,@Year bigint
        Set Rs1 = New ADODB.Recordset
            sql = "d_sp_TotalTransDeduct '" & rst!code & "'," & month(DTPMilkDate) & "," & year(DTPMilkDate) & ""
        Set Rs1 = oSaccoMaster.GetRecordset(sql)
    If Not Rs1.EOF Then
    Dim TransTotalDed As Currency
    If Not IsNull(Rs1.Fields(0)) Then TransTotalDed = Rs1.Fields(0)
    End If

    Set rs3 = New ADODB.Recordset
    Dim DESCR As String
    Dim amount As Currency

    '--d_sp_SelTransDed @Code varchar(35), @startdate varchar(10),@enddate varchar(10)
    sql = "d_sp_SelTransDed '" & rst!code & "','" & Startdate & "','" & Enddate & "'"
    Set rs3 = oSaccoMaster.GetRecordset(sql)

    agrovet = 0
    AI = 0
    TMShares = 0
    FSA = 0
    HShares = 0
    Advance = 0
    Others = 0
    TransTotalDed = 0


If Not rs3.EOF Then
    While Not rs3.EOF
    DESCR = Trim(rs3.Fields(0))
    amount = 0
    amount = rs3.Fields(1)
    sql = "SELECT     Agrovet, AI, TMShares, FSA, HShares, Advance, Others FROM d_TransportersPayroll WHERE Code='" & rst!code & "' AND EndPeriod ='" & Enddate & "'"
    Set rs4 = oSaccoMaster.GetRecordset(sql)
     If UCase(rs4.Fields(0).name) = UCase(DESCR) Then
        agrovet = amount
    End If
    If UCase(rs4.Fields(1).name) = UCase(DESCR) Then
        AI = amount
    End If
    If UCase(rs4.Fields(2).name) = UCase(DESCR) Then
        TMShares = amount
    End If
    If UCase(rs4.Fields(3).name) = UCase(DESCR) Then
        FSA = amount
    End If
    If UCase(rs4.Fields(4).name) = UCase(DESCR) Then
        HShares = amount
    End If
    If UCase(rs4.Fields(5).name) = UCase(DESCR) Then
        Advance = amount
    End If
    If UCase(rs4.Fields(6).name) = UCase(DESCR) Then
        Others = amount
    End If

    '//d_sp_UpdatePayrollDed  @SNo bigint,@EndPeriod varchar(15),@Transport money,@Agrovet money,@AI money,@TMShares money,@FSA money,@HShares money,@Advance money,@Others
    rs3.MoveNext
    Wend

End If
     ' d_sp_UpdateTransDed  @Code varchar(35),@EndPeriod varchar(15),@TotalDed money,@Agrovet money,@AI money,@TMShares money,@FSA money,@HShares money,@Advance money,@Others money

    Set cn = New ADODB.Connection
    sql = "d_sp_UpdateTransDed  '" & rst!code & "','" & Enddate & "'," & TransTotalDed & "," & agrovet & "," & AI & "," & TMShares & "," & FSA & "," & HShares & "," & Advance & "," & Others & ",0,0"
    oSaccoMaster.ExecuteThis (sql)


        '//Update supplier deductions
            Set cn = New ADODB.Connection
            Dim r2  As String
            r2 = "  "
                sql = "d_sp_SupplierDeduct " & txtSNo & ",'" & DTPMilkDate & "','Transport'," & rs!rate * CCur(txtQnty) & ",'" & Startdate & "','" & Enddate & "'," & year(Enddate) & ",'" & User & "','" & r2 & "',''"
            oSaccoMaster.ExecuteThis (sql)




End If
'Check inactive Transporter
Set rs = New ADODB.Recordset
sql = "d_sp_TransInActive " & txtSNo & ",'" & Startdate & "'"
Set rs = oSaccoMaster.GetRecordset(sql)
If Not rs.EOF Then

'--Update Detailed Trasport
    transdate = rs!Startdate
        If transdate < Startdate Then
            transdate = Startdate
        End If

    '---d_sp_SelTransDet @SNo bigint,@StartDate varchar(10), @Enddate varchar(10)
        Set rst = New ADODB.Recordset
            sql = "d_sp_SelTransDet " & txtSNo & ",'" & transdate & "','" & rs!dateinactivate & "'"
        Set rst = oSaccoMaster.GetRecordset(sql)

    '--d_sp_UpdateDetTrans @SNo bigint,@QNTY float,@Amnt money,@Code varchar(35),@Subsidy money,@EPeriod varchar(10),@user varchar(35)
        If Not rst.EOF Then
            Set cn = New ADODB.Connection
                sql = "d_sp_UpdateDetTrans " & txtSNo & "," & rst!qnty & "," & rst!amount & ",'" & rst!code & "'," & rst!subsidy & ",'" & Enddate & "','" & User & "'"
            oSaccoMaster.ExecuteThis (sql)

            '----d_sp_SelTransGPayQnty @EP varchar(10), @Code varchar(35)
        Set rst2 = New ADODB.Recordset
            sql = "d_sp_SelTransGPayQnty '" & Enddate & "','" & rst!code & "'"
        Set rst2 = oSaccoMaster.GetRecordset(sql)






    '-- d_sp_UpdateTransPay @Code varchar(35), @Qnty float,@Amnt money,@Subsidy money,@GrossPay money, @EndDate varchar(10)  AS
    If Not rst2.EOF Then
            Set cn = New ADODB.Connection
                sql = "d_sp_UpdateTransPay '" & rst!code & "'," & rst2!qnty & "," & rst2!Amnt & "," & rst2!subsidy & "," & rst2!GPay & ",'" & Enddate & "','" & User & "'"
            oSaccoMaster.ExecuteThis (sql)

        End If

        '---Get transporters Total Deductions//d_sp_TotalTransDeduct @Code varchar(35),@Month bigint,@Year bigint
        Set Rs1 = New ADODB.Recordset
            sql = "d_sp_TotalTransDeduct '" & rst!code & "'," & month(DTPMilkDate) & "," & year(DTPMilkDate) & ""
        Set Rs1 = oSaccoMaster.GetRecordset(sql)
    If Not Rs1.EOF Then
    'Dim TransTotalDed As Currency
    If Not IsNull(Rs1.Fields(0)) Then TransTotalDed = Rs1.Fields(0)
    End If

    Set rs3 = New ADODB.Recordset
    'Dim Startdate As String, Enddate As String
    'Dim DESCR As String
   ' Dim amount As Currency

    '--d_sp_SelTransDed @Code varchar(35), @startdate varchar(10),@enddate varchar(10)
    sql = "d_sp_SelTransDed '" & rst!code & "','" & Startdate & "','" & Enddate & "'"
    Set rs3 = oSaccoMaster.GetRecordset(sql)

    agrovet = 0
    AI = 0
    TMShares = 0
    FSA = 0
    HShares = 0
    Advance = 0
    Others = 0
    TransTotalDed = 0
If Not rs3.EOF Then
    While Not rs3.EOF
    DESCR = Trim(rs3.Fields(0))
    amount = 0
    amount = rs3.Fields(1)
    sql = "SELECT     Agrovet, AI, TMShares, FSA, HShares, Advance, Others FROM d_TransportersPayroll WHERE Code='" & rst!code & "' AND EndPeriod ='" & Enddate & "'"
    Set rs4 = oSaccoMaster.GetRecordset(sql)
     If UCase(rs4.Fields(0).name) = UCase(DESCR) Then
        agrovet = amount
    End If
    If UCase(rs4.Fields(1).name) = UCase(DESCR) Then
        AI = amount
    End If
    If UCase(rs4.Fields(2).name) = UCase(DESCR) Then
        TMShares = amount
    End If
    If UCase(rs4.Fields(3).name) = UCase(DESCR) Then
        FSA = amount
    End If
    If UCase(rs4.Fields(4).name) = UCase(DESCR) Then
        HShares = amount
    End If
    If UCase(rs4.Fields(5).name) = UCase(DESCR) Then
        Advance = amount
    End If
    If UCase(rs4.Fields(6).name) = UCase(DESCR) Then
        Others = amount
    End If

    '//d_sp_UpdatePayrollDed  @SNo bigint,@EndPeriod varchar(15),@Transport money,@Agrovet money,@AI money,@TMShares money,@FSA money,@HShares money,@Advance money,@Others
    rs3.MoveNext
    Wend
    ' d_sp_UpdateTransDed  @Code varchar(35),@EndPeriod varchar(15),@TotalDed money,@Agrovet money,@AI money,@TMShares money,@FSA money,@HShares money,@Advance money,@Others money

End If
 Set cn = New ADODB.Connection
    sql = "d_sp_UpdateTransDed  '" & rst!code & "','" & Enddate & "'," & TransTotalDed & "," & agrovet & "," & AI & "," & TMShares & "," & FSA & "," & HShares & "," & Advance & "," & Others & ""
    oSaccoMaster.ExecuteThis (sql)
    End If

End If


End If



'//d_sp_TotalDeduct-Total Deductions
'//d_sp_UpdateGPAYQnty - Total Grosspay and Quantity
'//d_sp_SupDed - Supply Deductions



Set rs2 = New ADODB.Recordset
'Dim Startdate As String, Enddate As String
Dim qnty As Currency, GPay As Currency
'Startdate = DateSerial(DTPMilkDate, cboMonth, 1)
sql = "d_sp_UpdateGPAYQnty '" & Startdate & "','" & Enddate & "'," & txtSNo & ""
Set rs2 = oSaccoMaster.GetRecordset(sql)
If Not rs2.EOF Then
If Not IsNull(rs2.Fields(0)) Then qnty = rs2.Fields(0)
If Not IsNull(rs2.Fields(1)) Then GPay = rs2.Fields(1)
End If


Set Rs1 = New ADODB.Recordset
sql = "d_sp_TotalDeduct " & txtSNo & "," & month(DTPMilkDate) & "," & year(DTPMilkDate) & ""
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
Startdate = DateSerial(year(DTPMilkDate), month(DTPMilkDate), 1)
Enddate = DateSerial(year(DTPMilkDate), month(DTPMilkDate) + 1, 1 - 1)
sql = "d_sp_SupDed " & txtSNo & ",'" & Startdate & "','" & Enddate & "'"
Set rs3 = oSaccoMaster.GetRecordset(sql)
If Not rs3.EOF Then
While Not rs3.EOF
desc = Trim(rs3.Fields(0))
Amnt = 0
Amnt = rs3.Fields(1)
sql = "SELECT     Transport, Agrovet, AI, TMShares, FSA, HShares, Advance, Others FROM d_Payroll WHERE SNo=" & txtSNo & " AND EndofPeriod ='" & Enddate & "'"
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
 '************message***************'
              '//check settings  for sms alert
        'If SendSmsOnSalary = True Then
        'Dim phone As String
        Dim RsPhoneNumbers As New ADODB.Recordset
        Dim MsgContent As String
        Set RsPhoneNumbers = oSaccoMaster.GetRecordset("select  PhoneNo  from d_Suppliers  where SNo ='" & txtSNo & "' and PhoneNo  is not null and PhoneNo <>'' and LEN(PhoneNo) =10 ")

        If Not RsPhoneNumbers.EOF Then
            Phone = Trim(RsPhoneNumbers.Fields!PhoneNo & "")
        Else
            Phone = ""
        End If


            '//insert into sms alert
        Set rs = New ADODB.Recordset
        sql = "d_sp_TotalMonth " & txtSNo & ",'" & Startdate & "','" & DTPMilkDate & "'"
        Set rs = oSaccoMaster.GetRecordset(sql)
        If Not rs.EOF Then
        If Not IsNull(rs.Fields(0)) Then CummulKgs = rs.Fields(0)
        Else
        CummulKgs = "0.00"
        End If

            If Trim(Phone) <> "" Then
            MsgContent = ""
            MsgContent = "Dear supplier, You have supplied," & txtQnty & " Kgs Today,Cummulative This Month :  " & Format(CummulKgs, "#,##0.00" & "Kgs") ' From Soitaran Mcs Dairy"
            strSQL = ""
            strSQL = "INSERT INTO Swift_Messages(SaccoCode,PhoneNumber,Message, msgstatus)"
            strSQL = strSQL & "Values (16,'" & Phone & "','" & MsgContent & "',0)"

            oSaccoMaster.ExecuteThis (strSQL)
            End If

Transport = 0
agrovet = 0
AI = 0
TMShares = 0
FSA = 0
HShares = 0
Advance = 0
Others = 0

'
'Dim Yr As Integer
'
'Yr = Year(DTPMilkDate)
''vbHourglass
''Fixed deductions update
'oSaccoMaster.ExecuteThis ("d_sp_PresetDeductAssign_99 '" & Startdate & "','" & Enddate & "'," & Yr & ",'" & User & "', " & txtSNo)
''Payroll update
''d_sp_GDedNet @StartDate varchar(10) , @endPeriod varchar(10)
'oSaccoMaster.ExecuteThis ("d_sp_GDedNet_99 '" & Startdate & "','" & Enddate & "'," & txtSNo)
'
''Update transporters
''d_sp_TransUpdate @StartDate varchar(10),@EndPeriod varchar(10),@User varchar(35) AS
'oSaccoMaster.ExecuteThis ("d_sp_TransUpdate_99 '" & Startdate & "','" & Enddate & "','" & User & "'," & txtSNo)
'
'
'oSaccoMaster.ExecuteThis ("d_sp_TransPRoll '" & Startdate & "','" & Enddate & "','" & User & "'")
''Lock period
''[d_sp_Periods]

'//Print Receipt
'**************PHONE NO**********'
Set RsPhoneNumbers = oSaccoMaster.GetRecordset("select  PhoneNo  from d_company ")

If Not RsPhoneNumbers.EOF Then
    Phone = Trim(RsPhoneNumbers.Fields!PhoneNo & "")
Else
    Phone = ""
End If




'**************END***************'


If chkPrint = vbChecked Then
    
'/*Print out
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
       'cdgPrint.Filter = "*.csv|*.txt"
        'cdgPrint.ShowSave
        Dim PORT As String
   '     PORT = ports
        'ttt = "LPT1" 'LPT1
        ttt = ports
        'ttt = "D:\PROJECTS\FOSA\DAILY" & Date & ""
        Set fso = CreateObject("Scripting.FileSystemObject")
        'Set chkPrinter = fso.GetFile(ttt)
        
    Set txtFile = fso.CreateTextFile(ttt, True)
    txtFile.WriteLine "         " & cname & ""
    txtFile.WriteLine "         Address :" & paddress & ""
    txtFile.WriteLine "         Phone :" & Phone & ""
    txtFile.WriteLine "         Email :" & Email & ""
    'txtfile.WriteLine " " & txtSNo
    
    txtFile.WriteLine "          Milk Receipt"
    txtFile.WriteLine "**********************************************"
        
    Set rs2 = New ADODB.Recordset
    sql = "d_sp_ReceiptNumber"
    Set rs2 = oSaccoMaster.GetRecordset(sql)
    
    Dim RNumber As String
    'RNumber = rs2.Fields(0)
    If Not IsNull(rs2.Fields(0)) Then RNumber = rs2.Fields(0)
    'Else
    'RNumber = "0"
    'End If
    
    txtFile.WriteLine "SNo :" & txtSNo
    txtFile.WriteLine "Name :" & lblNames
    txtFile.WriteLine "Quantity Supplied :" & txtQnty & " Kgs"
    Set rs = New ADODB.Recordset
    sql = "d_sp_TotalMonth " & txtSNo & ",'" & Startdate & "','" & DTPMilkDate & "'"
    Set rs = oSaccoMaster.GetRecordset(sql)
    If Not rs.EOF Then
    If Not IsNull(rs.Fields(0)) Then CummulKgs = rs.Fields(0)
    Else
    CummulKgs = "0.00"
    End If
    txtFile.WriteLine "Cummulative This Month : " & Format(CummulKgs, "#,##0.00" & " Kgs")
'     '************message***************'
'              '//check settings  for sms alert
'        'If SendSmsOnSalary = True Then
'        'Dim phone As String
'        Dim RsPhoneNumbers As New ADODB.Recordset
'        Dim MsgContent As String
'        Set RsPhoneNumbers = oSaccoMaster.GetRecordset("select  PhoneNo  from d_Suppliers  where SNo ='" & txtSNo & "' and MobileNo  is not null and PhoneNo <>'' and LEN(PhoneNo) =10 ")
'
'        If Not RsPhoneNumbers.EOF Then
'            phone = Trim(RsPhoneNumbers.Fields!PhoneNo & "")
'        Else
'            phone = ""
'        End If
'
'
'            '//insert into sms alert
'
'            If Trim(phone) <> "" Then
'            MsgContent = ""
'            MsgContent = "Dear supplier, You have supplied," & txtQnty & " Kgs Today,Cummulative This Month :  " & Format(CummulKgs, "#,##0.00" & "Kgs") ' From Soitaran Mcs Dairy"
'            strSQL = ""
'            strSQL = "INSERT INTO Swift_Messages(SaccoCode,PhoneNumber,Message, msgstatus)"
'            strSQL = strSQL & "Values (5,'" & phone & "','" & MsgContent & "',0)"
'
'            oSaccoMaster.ExecuteThis (strSQL)
'            End If
    
    
    
    Set rs = New ADODB.Recordset
    sql = "d_sp_TransName '" & txtSNo & "'"
    Set rs = oSaccoMaster.GetRecordset(sql)
    If Not rs.EOF Then
    If Not IsNull(rs.Fields(0)) Then TRANSPORTER = rs.Fields(0)
    Else
        TRANSPORTER = "Self"
    End If
    txtFile.WriteLine "---------------------------------------"
    txtFile.WriteLine "Receipt Number :" & RNumber
    txtFile.WriteLine "TRANSPORTER :" & TRANSPORTER
    txtFile.WriteLine "Received by :" & username
    txtFile.WriteLine "---------------------------------------"
    txtFile.WriteLine "     Date :" & Format(DTPMilkDate, "dd/mm/yyyy") & " ,Time : " & Format(Time, "hh:mm:ss AM/PM")
    txtFile.WriteLine "" & motto & ""
    txtFile.WriteLine "---------------------------------------"
    If chkComment.value = vbChecked Then
        txtFile.WriteLine txtComment
        txtFile.WriteLine "---------------------------------------"
        txtFile.WriteLine "********POWERED BY EASYMA***************"
    End If
    txtFile.WriteLine escFeedAndCut
    
 txtFile.Close
 Reset
End If


'* writing to notepad

If chkNotepad.value = vbChecked Then

'    Dim fso, txtfile
'    Dim ttt
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
    txtFile.WriteLine " " & paddress & ""
    txtFile.WriteLine " " & Phone & ""
   ' Printer.Print Tab(0); "Kimathi House Branch"
    txtFile.WriteLine " " & paddress & " "
    txtFile.WriteLine "" & town & ""
    txtFile.WriteLine "Milk Receipt"
    txtFile.WriteLine "---------------------------------------"
'    If cbomemtrans = "Shares" Then
'    DESC = bosanames & " -Member No " & memberno
    txtFile.WriteLine "SNo :" & txtSNo
    txtFile.WriteLine "Name :" & lblNames
'    Else
    txtFile.WriteLine "Quantity Supplied :" & txtQnty & " Kgs"
    Startdate = DateSerial(year(DTPMilkDate), month(DTPMilkDate) - 1, 1)
    'sql = "d_sp_TotalMonth " & txtSNo & ",'" & StartDate & "','" & DTPMilkDate & "'"
    Set rs = New ADODB.Recordset
    sql = "d_sp_TotalMonth " & txtSNo & ",'" & Startdate & "','" & DTPMilkDate & "'"
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
    txtFile.WriteLine "     " & motto & " "
    txtFile.WriteLine "---------------------------------------"
    txtFile.WriteLine escFeedAndCut
    'Printer.Print
    'Printer.CurrentX = 500#
    'Printer.FontSize = 10
    'Printer.CurrentX = 500
    'Printer.FontSize = 8
'    Printer.Print Tab(0); "Date :"; Tab(10); Format(Get_Server_Date, "dd/mm/yyyy HH:MM:SS AM/PM")
'    Printer.Print Tab(0); "     TDL - Improves Your Value "
'    Printer.Print Tab(0); "---------------------------------------"
'    Printer.CurrentX = 500
'    Printer.FontSize = 8
'    Printer.Print
'    Printer.CurrentX = 500
'    Printer.FontSize = 8
'    Printer.Print
txtFile.Close
End If

loadMilk

txtSNo = ""
txtQnty = ""
'txtSNo_Validate True
txtSNo.SetFocus

lvwpreviousmilk.Visible = False

Exit Sub
ErrorHandler:

MsgBox err.description

End Sub

Public Sub loadMilk()

    lblDTotal.Caption = "0"
    Set rs = New ADODB.Recordset
    Dim rstt As New Recordset
    sql = "d_sp_DailyTotal '" & DTPMilkDate & "'"
    Set rs = oSaccoMaster.GetRecordset(sql)
    If Not rs.EOF Then
    If Not IsNull(rs.Fields(0)) Then lblDTotal.Caption = rs.Fields(0)
    Else
    lblDTotal.Caption = "0"
    End If
    sql = "set dateformat dmy SELECT     COUNT(SNo) AS Expr1 From d_Milkintake WHERE     (TransDate >= '" & DTPMilkDate & "') AND (QSupplied > 0)"
    Set rstt = oSaccoMaster.GetRecordset(sql)
    If Not rstt.EOF Then
    If Not IsNull(rstt.Fields(0)) Then txttotalsup.Text = rstt.Fields(0)
    Else
    txttotalsup.Text = "0"
    End If
    With lvwMilkCollected
    
       .ListItems.Clear
    
        .ColumnHeaders.Clear

  End With

    Set rs = CreateObject("adodb.recordset")
  
    sql = "d_sp_CurrentList '" & Format(DTPMilkDate, "dd/mm/yyyy") & "','" & User & "'"
    
   Set rs = oSaccoMaster.GetRecordset(sql)
    
    With lvwMilkCollected
        
        
        .ColumnHeaders.Add , , "SNo", 2000
        .ColumnHeaders.Add , , "Names", 3000
        .ColumnHeaders.Add , , "QNTY"
        .ColumnHeaders.Add , , "Time"
        .ColumnHeaders.Add , , "Receipt No.", 2000
    
        While Not rs.EOF
            If Not IsNull(rs.Fields("SNo")) Then
            Set li = .ListItems.Add(, , Trim(rs.Fields("SNo")))
            End If
            If Not IsNull(rs.Fields("Names")) Then
            li.ListSubItems.Add , , Trim(rs.Fields("Names"))
            End If
            If Not IsNull(rs.Fields("QSupplied")) Then
            li.ListSubItems.Add , , Trim(rs.Fields("QSupplied"))
            End If
            If Not IsNull(rs.Fields("TransTime")) Then
            li.ListSubItems.Add , , Trim(rs.Fields("TransTime"))
            End If
            If Not IsNull(rs.Fields("ID")) Then
            li.ListSubItems.Add , , Trim(rs.Fields("ID"))
            End If

            
                    rs.MoveNext
        
        Wend
        
    End With
    
    rs.Close
    
    Set rs = Nothing
    
lvwMilkCollected.View = lvwReport

End Sub

Private Sub cmdRemove_Click()
Dim Valu As String
Dim TransTotalDed As Currency
Dim DESCR As String
Dim amount As Currency
    

'InputBox
Valu = InputBox("Please enter the receipt number", "REMOVE INTAKE", "<Enter receipt number here>")

If Valu = "" Then
    MsgBox "Please enter receipt number "
    Exit Sub
End If
    
If Valu = "<Enter receipt number here>" Then
    MsgBox "Please enter receipt number "
    Exit Sub
End If
    
If Not IsNumeric(Valu) Then
    MsgBox "Please enter receipt number." & Valu & " is not a number", vbCritical
    Exit Sub
End If

Set rs = New ADODB.Recordset
sql = "d_sp_SpecificReceipt " & Valu & ""
Set rs = oSaccoMaster.GetRecordset(sql)

Startdate = DateSerial(year(rs.Fields("TransDate")), month(rs.Fields("TransDate")), 1)
Enddate = DateSerial(year(rs.Fields("TransDate")), month(rs.Fields("TransDate")) + 1, 1 - 1)

sql = "d_sp_DeleteMIn " & Valu & ""
oSaccoMaster.ExecuteThis (sql)

'//Delete Daily Intake

Set rs2 = New ADODB.Recordset
sql = "SET dateformat DMY SELECT [" & Day(rs.Fields("TransDate")) & "] AS Milk from d_DailySummary WHERE SNo =" & rs.Fields("SNo") & " AND Endperiod ='" & Enddate & "'"
Set rs2 = oSaccoMaster.GetRecordset(sql)
If Not rs2.EOF Then
If ((rs2.Fields(0)) > 0) Then
sql = "SET dateformat DMY Update  d_DailySummary SET [" & Day(DTPMilkDate) & "] = " & CCur(rs2!milk) - CCur(rs.Fields("QSupplied")) & " WHERE SNo =" & rs.Fields("SNo") & " AND Endperiod ='" & Enddate & "'"
oSaccoMaster.ExecuteThis (sql)
Else
sql = "SET dateformat DMY Update  d_DailySummary SET [" & Day(DTPMilkDate) & "] = '' WHERE SNo =" & rs.Fields("SNo") & " AND Endperiod ='" & Enddate & "'"
oSaccoMaster.ExecuteThis (sql)
End If
End If


Set Rs1 = New ADODB.Recordset
sql = "d_sp_TransActive " & rs.Fields("SNo") & ",'" & rs.Fields("TransDate") & "'"
Set Rs1 = oSaccoMaster.GetRecordset(sql)
    If Not Rs1.EOF Then
        '//Update deductions
        Set cn = New ADODB.Connection

        Dim r2  As String
        r2 = "Transport"
            sql = "d_sp_DeleteSupDeduc " & rs.Fields("SNo") & "," & Rs1!rate * rs.Fields("QSupplied") & ",'" & r2 & "','" & rs.Fields("TransDate") & "'"
        oSaccoMaster.ExecuteThis (sql)
        
        transdate = Rs1!Startdate
        If transdate < Startdate Then
            transdate = Startdate
        End If
        
        
        '---d_sp_SelTransDet @SNo bigint,@StartDate varchar(10), @Enddate varchar(10)
        Set rst = New ADODB.Recordset
            sql = "d_sp_SelTransDet " & rs.Fields("SNo") & ",'" & transdate & "','" & Enddate & "'"
        Set rst = oSaccoMaster.GetRecordset(sql)
        
    '--d_sp_UpdateDetTrans @SNo bigint,@QNTY float,@Amnt money,@Code varchar(35),@Subsidy money,@EPeriod varchar(10),@user varchar(35)
        If Not rst.EOF Then
            Set cn = New ADODB.Connection
                sql = "d_sp_UpdateDetTrans " & rs.Fields("SNo") & "," & rst!qnty & "," & rst!amount & ",'" & rst!code & "'," & rst!subsidy & ",'" & Enddate & "','" & User & "'"
            oSaccoMaster.ExecuteThis (sql)
            
        
        Else
            Set cn = New ADODB.Connection
                sql = "delete d_TransDetailed WHERE SNO = " & rs.Fields("SNo") & " AND EndPeriod ='" & Enddate & "'"
            oSaccoMaster.ExecuteThis (sql)
        End If
        '----d_sp_SelTransGPayQnty @EP varchar(10), @Code varchar(35)
        Set rst2 = New ADODB.Recordset
            sql = "d_sp_SelTransGPayQnty '" & Enddate & "','" & rst!code & "'"
        Set rst2 = oSaccoMaster.GetRecordset(sql)
        
    '-- d_sp_UpdateTransPay @Code varchar(35), @Qnty float,@Amnt money,@Subsidy money,@GrossPay money, @EndDate varchar(10)  AS
    If Not rst2.EOF Then
            Set cn = New ADODB.Connection
                sql = "d_sp_UpdateTransPay '" & rst!code & "'," & rst2!qnty & "," & rst2!Amnt & "," & rst2!subsidy & "," & rst2!GPay & ",'" & Enddate & "','" & User & "'"
            oSaccoMaster.ExecuteThis (sql)
            
        End If
 
        '---Get transporters Total Deductions//d_sp_TotalTransDeduct @Code varchar(35),@Month bigint,@Year bigint
        Set Rs1 = New ADODB.Recordset
            sql = "d_sp_TotalTransDeduct '" & rst!code & "'," & month(rs.Fields("transdate")) & "," & year(rs.Fields("transdate")) & ""
        Set Rs1 = oSaccoMaster.GetRecordset(sql)
    If Not Rs1.EOF Then
    'Dim TransTotalDed As Currency
    If Not IsNull(Rs1.Fields(0)) Then TransTotalDed = Rs1.Fields(0)
    End If
    
    Set rs3 = New ADODB.Recordset
    'Dim Startdate As String, Enddate As String
    'Dim DESCR As String
   ' Dim amount As Currency
    
    '--d_sp_SelTransDed @Code varchar(35), @startdate varchar(10),@enddate varchar(10)
    sql = "d_sp_SelTransDed '" & rst!code & "','" & Startdate & "','" & Enddate & "'"
    Set rs3 = oSaccoMaster.GetRecordset(sql)
If Not rs3.EOF Then
    While Not rs3.EOF
    DESCR = Trim(rs3.Fields(0))
    amount = 0
    amount = rs3.Fields(1)
    sql = "SELECT     Agrovet, AI, TMShares, FSA, HShares, Advance, Others FROM d_TransportersPayroll WHERE Code='" & rst!code & "' AND EndPeriod ='" & Enddate & "'"
    Set rs4 = oSaccoMaster.GetRecordset(sql)
     If UCase(rs4.Fields(0).name) = UCase(DESCR) Then
        agrovet = amount
    End If
    If UCase(rs4.Fields(1).name) = UCase(DESCR) Then
        AI = amount
    End If
    If UCase(rs4.Fields(2).name) = UCase(DESCR) Then
        TMShares = amount
    End If
    If UCase(rs4.Fields(3).name) = UCase(DESCR) Then
        FSA = amount
    End If
    If UCase(rs4.Fields(4).name) = UCase(DESCR) Then
        HShares = amount
    End If
    If UCase(rs4.Fields(5).name) = UCase(DESCR) Then
        Advance = amount
    End If
    If UCase(rs4.Fields(6).name) = UCase(DESCR) Then
        Others = amount
    End If

    '//d_sp_UpdatePayrollDed  @SNo bigint,@EndPeriod varchar(15),@Transport money,@Agrovet money,@AI money,@TMShares money,@FSA money,@HShares money,@Advance money,@Others
    rs3.MoveNext
    Wend
    ' d_sp_UpdateTransDed  @Code varchar(35),@EndPeriod varchar(15),@TotalDed money,@Agrovet money,@AI money,@TMShares money,@FSA money,@HShares money,@Advance money,@Others money
    Set cn = New ADODB.Connection
    sql = "d_sp_UpdateTransDed  '" & rst!code & "','" & Enddate & "'," & TransTotalDed & "," & agrovet & "," & AI & "," & TMShares & "," & FSA & "," & HShares & "," & Advance & "," & Others & ""
    oSaccoMaster.ExecuteThis (sql)
End If

    agrovet = 0
    AI = 0
    TMShares = 0
    FSA = 0
    HShares = 0
    Advance = 0
    Others = 0
    TransTotalDed = 0
 
    End If

Set Rs1 = New ADODB.Recordset
sql = "d_sp_TransInActive " & rs.Fields("SNo") & ",'" & rs.Fields("TransDate") & "'"
Set Rs1 = oSaccoMaster.GetRecordset(sql)
If Not Rs1.EOF Then
''//Update deductions
sql = "d_sp_DeleteSupDeduc " & rs.Fields("SNo") & "," & Rs1!rate * rs.Fields("QSupplied") & ",'Transport','" & rs.Fields("TransDate") & "'"
oSaccoMaster.ExecuteThis (sql)


        '---d_sp_SelTransDet @SNo bigint,@StartDate varchar(10), @Enddate varchar(10)
        Set rst = New ADODB.Recordset
            sql = "d_sp_SelTransDet " & rs.Fields("SNo") & ",'" & transdate & "','" & Enddate & "'"
        Set rst = oSaccoMaster.GetRecordset(sql)
        
    '--d_sp_UpdateDetTrans @SNo bigint,@QNTY float,@Amnt money,@Code varchar(35),@Subsidy money,@EPeriod varchar(10),@user varchar(35)
        If Not rst.EOF Then
            Set cn = New ADODB.Connection
                sql = "d_sp_UpdateDetTrans " & rs.Fields("SNo") & "," & rst!qnty & "," & rst!amount & ",'" & rst!code & "'," & rst!subsidy & ",'" & Enddate & "','" & User & "'"
            oSaccoMaster.ExecuteThis (sql)
            
        Else
            Set cn = New ADODB.Connection
                sql = "delete d_TransDetailed WHERE SNO = " & rs.Fields("SNo") & " AND EndPeriod ='" & Enddate & "'"
            oSaccoMaster.ExecuteThis (sql)
            
        End If
        
                '----d_sp_SelTransGPayQnty @EP varchar(10), @Code varchar(35)
        Set rst2 = New ADODB.Recordset
            sql = "d_sp_SelTransGPayQnty '" & Enddate & "','" & rst!code & "'"
        Set rst2 = oSaccoMaster.GetRecordset(sql)
        
    '-- d_sp_UpdateTransPay @Code varchar(35), @Qnty float,@Amnt money,@Subsidy money,@GrossPay money, @EndDate varchar(10)  AS
    If Not rst2.EOF Then
            Set cn = New ADODB.Connection
                sql = "d_sp_UpdateTransPay '" & rst!code & "'," & rst2!qnty & "," & rst2!Amnt & "," & rst2!subsidy & "," & rst2!GPay & ",'" & Enddate & "','" & User & "'"
            oSaccoMaster.ExecuteThis (sql)
            
        End If
 
        '---Get transporters Total Deductions//d_sp_TotalTransDeduct @Code varchar(35),@Month bigint,@Year bigint
        Set Rs1 = New ADODB.Recordset
            sql = "d_sp_TotalTransDeduct '" & rst!code & "'," & month(rs.Fields("transdate")) & "," & year(rs.Fields("transdate")) & ""
        Set Rs1 = oSaccoMaster.GetRecordset(sql)
    If Not Rs1.EOF Then
    'Dim TransTotalDed As Currency
    If Not IsNull(Rs1.Fields(0)) Then TransTotalDed = Rs1.Fields(0)
    End If
    
    Set rs3 = New ADODB.Recordset
    'Dim Startdate As String, Enddate As String
    'Dim DESCR As String
   ' Dim amount As Currency
    
    '--d_sp_SelTransDed @Code varchar(35), @startdate varchar(10),@enddate varchar(10)
    sql = "d_sp_SelTransDed '" & rst!code & "','" & Startdate & "','" & Enddate & "'"
    Set rs3 = oSaccoMaster.GetRecordset(sql)
If Not rs3.EOF Then
    While Not rs3.EOF
    DESCR = Trim(rs3.Fields(0))
    amount = 0
    amount = rs3.Fields(1)
    sql = "SELECT     Agrovet, AI, TMShares, FSA, HShares, Advance, Others FROM d_TransportersPayroll WHERE Code='" & rst!code & "' AND EndPeriod ='" & Enddate & "'"
    Set rs4 = oSaccoMaster.GetRecordset(sql)
     If UCase(rs4.Fields(0).name) = UCase(DESCR) Then
        agrovet = amount
    End If
    If UCase(rs4.Fields(1).name) = UCase(DESCR) Then
        AI = amount
    End If
    If UCase(rs4.Fields(2).name) = UCase(DESCR) Then
        TMShares = amount
    End If
    If UCase(rs4.Fields(3).name) = UCase(DESCR) Then
        FSA = amount
    End If
    If UCase(rs4.Fields(4).name) = UCase(DESCR) Then
        HShares = amount
    End If
    If UCase(rs4.Fields(5).name) = UCase(DESCR) Then
        Advance = amount
    End If
    If UCase(rs4.Fields(6).name) = UCase(DESCR) Then
        Others = amount
    End If

    '//d_sp_UpdatePayrollDed  @SNo bigint,@EndPeriod varchar(15),@Transport money,@Agrovet money,@AI money,@TMShares money,@FSA money,@HShares money,@Advance money,@Others
    rs3.MoveNext
    Wend
    ' d_sp_UpdateTransDed  @Code varchar(35),@EndPeriod varchar(15),@TotalDed money,@Agrovet money,@AI money,@TMShares money,@FSA money,@HShares money,@Advance money,@Others money
    Set cn = New ADODB.Connection
    sql = "d_sp_UpdateTransDed  '" & rst!code & "','" & Enddate & "'," & TransTotalDed & "," & agrovet & "," & AI & "," & TMShares & "," & FSA & "," & HShares & "," & Advance & "," & Others & ""
    oSaccoMaster.ExecuteThis (sql)
End If

    agrovet = 0
    AI = 0
    TMShares = 0
    FSA = 0
    HShares = 0
    Advance = 0
    Others = 0
    TransTotalDed = 0
End If


'//d_sp_TotalDeduct-Total Deductions
'//d_sp_UpdateGPAYQnty - Total Grosspay and Quantity
'//d_sp_SupDed - Supply Deductions



Set rs2 = New ADODB.Recordset
'Dim Startdate As String, Enddate As String
Dim qnty As Currency, GPay As Currency
'Startdate = DateSerial(DTPMilkDate, cboMonth, 1)
sql = "d_sp_UpdateGPAYQnty '" & Startdate & "','" & Enddate & "'," & rs.Fields("SNo") & ""
Set rs2 = oSaccoMaster.GetRecordset(sql)
If Not rs2.EOF Then
If Not IsNull(rs2.Fields(0)) Then qnty = rs2.Fields(0)
If Not IsNull(rs2.Fields(1)) Then GPay = rs2.Fields(1)
End If


Set Rs1 = New ADODB.Recordset
sql = "d_sp_TotalDeduct " & rs.Fields("SNo") & "," & month(DTPMilkDate) & "," & year(DTPMilkDate) & ""
Set Rs1 = oSaccoMaster.GetRecordset(sql)
If Not Rs1.EOF Then
Dim TotalDed As Currency
If Not IsNull(Rs1.Fields(0)) Then TotalDed = Rs1.Fields(0)
End If
'//Update payroll -- @SNo bigint,@EndPeriod varchar(15),@Kgs float,@GPay money,@NPay money,@TDeductions money,@auditid  varchar(35)
Set cn = New ADODB.Connection
sql = "d_sp_UpdatePayroll  " & rs.Fields("SNo") & ",'" & Enddate & "'," & qnty & "," & GPay & "," & GPay - TotalDed & "," & TotalDed & ",'" & User & "'"
oSaccoMaster.ExecuteThis (sql)



Set rs3 = New ADODB.Recordset
''Dim Startdate As String, Enddate As String
Dim desc As String
Dim Amnt As Currency
'Startdate = DateSerial(Year(DTPMilkDate), month(DTPMilkDate), 1)
'Enddate = DateSerial(Year(DTPMilkDate), month(DTPMilkDate) + 1, 1 - 1)
sql = "d_sp_SupDed " & rs.Fields("SNo") & ",'" & Startdate & "','" & Enddate & "'"
Set rs3 = oSaccoMaster.GetRecordset(sql)
If Not rs3.EOF Then
While Not rs3.EOF
desc = Trim(rs3.Fields(0))
Amnt = 0
Amnt = rs3.Fields(1)
sql = "SELECT     Transport, Agrovet, AI, TMShares, FSA, HShares, Advance, Others FROM d_Payroll WHERE SNo=" & rs.Fields("SNo") & " AND EndofPeriod ='" & Enddate & "'"
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
Else
Transport = 0
'//Update Deductions -- d_sp_UpdatePayrollDed  @SNo bigint,@EndPeriod varchar(15),@Transport money,@Agrovet money,@AI money,@TMShares money,@FSA money,@HShares money,@Advance money,@Others money
Set cn = New ADODB.Connection
sql = "d_sp_UpdatePayrollDed  " & rs.Fields("SNo") & ",'" & Enddate & "'," & Transport & "," & agrovet & "," & AI & "," & TMShares & "," & FSA & "," & HShares & "," & Advance & "," & Others & ""
oSaccoMaster.ExecuteThis (sql)
End If

Dim Yr As Integer

Yr = year(DTPMilkDate)
'vbHourglass
'Fixed deductions update
oSaccoMaster.ExecuteThis ("d_sp_PresetDeductAssign_99 '" & Startdate & "','" & Enddate & "'," & Yr & ",'" & User & "', " & rs.Fields("SNo"))

'Payroll update
'd_sp_GDedNet @StartDate varchar(10) , @endPeriod varchar(10)
oSaccoMaster.ExecuteThis ("d_sp_GDedNet_99 '" & Startdate & "','" & Enddate & "'," & rs.Fields("SNo"))

'Update transporters
'd_sp_TransUpdate @StartDate varchar(10),@EndPeriod varchar(10),@User varchar(35) AS
oSaccoMaster.ExecuteThis ("d_sp_TransUpdate_99 '" & Startdate & "','" & Enddate & "','" & User & "'," & rs.Fields("SNo"))


oSaccoMaster.ExecuteThis ("d_sp_TransPRoll '" & Startdate & "','" & Enddate & "','" & User & "'")
'Lock period
'[d_sp_Periods]

loadMilk
MsgBox "Record removed successfully."

Transport = 0
agrovet = 0
AI = 0
TMShares = 0
FSA = 0
HShares = 0
Advance = 0
Others = 0

    
End Sub

Private Sub Command1_Click()

End Sub

Private Sub cmdreprintreceipt_Click()
On Error GoTo ErrorHandler
Dim Price As Currency
Dim Startdate, CummulKgs, TRANSPORTER As String
Dim transdate As Date
Startdate = DateSerial(year(DTPMilkDate), month(DTPMilkDate), 1)
Enddate = DateSerial(year(DTPMilkDate), month(DTPMilkDate) + 1, 1 - 1)
    
'/*Print out
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
       'cdgPrint.Filter = "*.csv|*.txt"
        'cdgPrint.ShowSave
        Dim PORT As String
 '      PORT = ports
        'ttt = "LPT1" 'LPT1
 '      ttt = PORT
        ttt = ports
        'ttt = "LPT1" 'LPT1
        'ttt = "D:\PROJECTS\FOSA\DAILY" & Date & ""
        Set fso = CreateObject("Scripting.FileSystemObject")
        'Set chkPrinter = fso.GetFile(ttt)
       
        
        Set txtFile = fso.CreateTextFile(ttt, True)
    txtFile.WriteLine "      " & cname & ""
    txtFile.WriteLine "             Milk Receipt"
    txtFile.WriteLine "---------------------------------------"
        
    Set rs2 = New ADODB.Recordset
    sql = "d_sp_ReceiptNumber"
    Set rs2 = oSaccoMaster.GetRecordset(sql)
    
    Dim RNumber As String
    'RNumber = rs2.Fields(0)
    If Not IsNull(rs2.Fields(0)) Then RNumber = rs2.Fields(0)
    'Else
    'RNumber = "0"
    'End If
    
    txtFile.WriteLine "SNo :" & txtSNo
    txtFile.WriteLine "Name :" & lblNames
    '//get the kilo supplied
      Set rs = New ADODB.Recordset
    sql = "d_sp_Totalday  " & txtSNo & ",'" & DTPMilkDate & "'"
    Set rs = oSaccoMaster.GetRecordset(sql)
    If Not rs.EOF Then
    If Not IsNull(rs.Fields(0)) Then txtQnty = rs.Fields(0)
    Else
    txtQnty = "0.00"
    End If
    txtFile.WriteLine "Quantity Supplied :" & txtQnty & " Kgs"
    Set rs = New ADODB.Recordset
    sql = "d_sp_TotalMonth " & txtSNo & ",'" & Startdate & "','" & DTPMilkDate & "'"
    Set rs = oSaccoMaster.GetRecordset(sql)
    If Not rs.EOF Then
    If Not IsNull(rs.Fields(0)) Then CummulKgs = rs.Fields(0)
    Else
    CummulKgs = "0.00"
    End If
    txtFile.WriteLine "Cummulative This Month : " & Format(CummulKgs, "#,##0.00" & " Kgs")
    Set rs = New ADODB.Recordset
    sql = "d_sp_TransName '" & txtSNo & "'"
    Set rs = oSaccoMaster.GetRecordset(sql)
    If Not rs.EOF Then
    If Not IsNull(rs.Fields(0)) Then TRANSPORTER = rs.Fields(0)
    Else
        TRANSPORTER = "Self"
    End If
    txtFile.WriteLine "---------------------------------------"
    txtFile.WriteLine "Receipt Number :" & RNumber
    txtFile.WriteLine "TRANSPORTER :" & TRANSPORTER
    txtFile.WriteLine "Received by :" & username
    txtFile.WriteLine "---------------------------------------"
    txtFile.WriteLine "  Date :" & Format(Get_Server_Date, "dd/mm/yyyy") & " ,Time : " & Format(Time, "hh:mm:ss AM/PM")
    txtFile.WriteLine "       " & motto & ""
    txtFile.WriteLine "---------------------------------------"
     txtFile.WriteLine "********POWERED BY EASYMA***************"
    If chkComment.value = vbChecked Then
        txtFile.WriteLine txtComment
        txtFile.WriteLine "---------------------------------------"
    End If
    txtFile.WriteLine escFeedAndCut
    
 txtFile.Close
 Reset
Exit Sub
ErrorHandler:
MsgBox err.description, vbCritical
End Sub

Private Sub DTPMilkDate_Change()
loadMilk
End Sub

Private Sub DTPMilkDate_Click()
DTPMilkDate = Format(Get_Server_Date, "dd/mm/yyyy")
loadMilk
End Sub

Private Sub DTPMilkDate_GotFocus()
DTPMilkDate = Format(Get_Server_Date, "dd/mm/yyyy")
loadMilk
End Sub

Private Sub DTPMilkDate_KeyDown(KeyCode As Integer, Shift As Integer)
DTPMilkDate = Format(Get_Server_Date, "dd/mm/yyyy")
loadMilk
End Sub

Private Sub DTPMilkDate_KeyPress(KeyAscii As Integer)
DTPMilkDate = Format(Get_Server_Date, "dd/mm/yyyy")
loadMilk
End Sub

Private Sub DTPMilkDate_Validate(Cancel As Boolean)
DTPMilkDate = Format(Get_Server_Date, "dd/mm/yyyy")
loadMilk
End Sub

Private Sub Form_Load()

DTPMilkDate = Format(Get_Server_Date, "dd/mm/yyyy")
DTPMilkDate.MinDate = Format(Get_Server_Date, "dd/mm/yyyy")
DTPMilkDate.MaxDate = Format(Get_Server_Date, "dd/mm/yyyy")
With StatusBar1.Panels
    .Item(1).Text = "USER : " & username
    .Item(2).Text = "DATE : " & Format(Get_Server_Date, "dd/mm/yyyy")

End With

loadMilk
lvwpreviousmilk.Visible = False

chkPrint.value = vbChecked
End Sub
Private Sub Form_click()
lvwpreviousmilk.Visible = False
'chkPrint.value = vbChecked
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'lblWait.Visible = True
'Timer1.Enabled = True
Startdate = DateSerial(year(DTPMilkDate), month(DTPMilkDate), 1)
Enddate = DateSerial(year(DTPMilkDate), month(DTPMilkDate) + 1, 1 - 1)
'
'oSaccoMaster.ExecuteThis ("d_sp_TransUpdate '" & Startdate & "','" & Enddate & "','" & User & "'")
'oSaccoMaster.ExecuteThis ("d_sp_TransPRoll '" & Startdate & "','" & Enddate & "','" & User & "'")

MainForm.Caption = "EasyMa "
End Sub

Private Sub Timer1_Timer()
End Sub

Private Sub Listmilktrend_BeforeLabelEdit(Cancel As Integer)

End Sub

Private Sub txtComment_Click()
If txtComment = "<Put your comment here>" Then
txtComment = ""
End If
End Sub

Private Sub txtQnty_KeyPress(KeyAscii As Integer)

If (KeyAscii = 8) Or (KeyAscii = 46) Or (KeyAscii = 48) Or (KeyAscii = 49) Or (KeyAscii = 50) Or (KeyAscii = 51) Or (KeyAscii = 52) Or (KeyAscii = 53) Or (KeyAscii = 54) Or (KeyAscii = 55) Or (KeyAscii = 56) Or (KeyAscii = 57) Then
KeyAscii = KeyAscii
Else
KeyAscii = 0
MsgBox "Please enter a number "
End If

'If KeyAscii = 45 Then
'MsgBox "Please enter a positive number "
'txtQnty = ""
'End If

lvwpreviousmilk.Visible = False

End Sub
Private Sub txtQnty_Validate(Cancel As Boolean)
txtQnty = Format(txtQnty, "####0.00")

lvwpreviousmilk.Visible = False

End Sub

Private Sub txtSNo_KeyPress(KeyAscii As Integer)
If (KeyAscii = 8) Or (KeyAscii = 48) Or (KeyAscii = 49) Or (KeyAscii = 50) Or (KeyAscii = 51) Or (KeyAscii = 52) Or (KeyAscii = 53) Or (KeyAscii = 54) Or (KeyAscii = 55) Or (KeyAscii = 56) Or (KeyAscii = 57) Then
KeyAscii = KeyAscii
Else
KeyAscii = 0
MsgBox "Please enter a number "

End If
End Sub
Private Sub txtSNo_Validate(Cancel As Boolean)
On Error GoTo ErrorHandler

Set rs = New ADODB.Recordset
sql = "d_sp_SelectSuppliers '" & txtSNo & "'"
Set rs = oSaccoMaster.GetRecordset(sql)
If Not rs.EOF Then
If Not IsNull(rs.Fields(2)) Then lblNames.Caption = rs.Fields(2)
Else
lblNames.Caption = ""
End If
If rs.RecordCount = 0 Then
lblNames.Caption = ""
End If
'//CHECK THE TRANSPORTER
If txtSNo = "" Then Exit Sub
sql = ""
sql = "SELECT     TOP 1 Trans_Code  FROM         d_Transport  WHERE     (Sno = " & txtSNo & ") AND (Active = 1)  ORDER BY ID DESC"
Set rs = oSaccoMaster.GetRecordset(sql)
If Not rs.EOF Then
txtTCode = rs.Fields(0)
'//PUT THE NAME OF THE TRANSPORTER
Set rst = New ADODB.Recordset
sql = "d_sp_SelectTrans '" & txtTCode & "'"
Set rst = oSaccoMaster.GetRecordset(sql)
If Not rst.EOF Then
If Not IsNull(rst.Fields(0)) Then lblTName = rst.Fields(0)
Else
lblTName = ""
End If
Else
txtTCode = ""
End If


If Trim$(txtSNo) <> "" Then

'Dim ruttoh As Long
'Dim rr As Integer
Dim pdate As Date
Dim rsmilk As New Recordset
  
        
Startdate = DateSerial(year(DTPMilkDate), month(DTPMilkDate), 1)
Enddate = DateSerial(year(DTPMilkDate), month(DTPMilkDate) + 1, 1 - 1)
        
pdate = CDate(CLng(DTPMilkDate) - 30)
        
sql = "set dateformat dmy SELECT * FROM d_Milkintake WHERE (SNo = " & txtSNo & ") and transdate>='" & pdate & "' and transdate<>'" & DTPMilkDate & "'    ORDER BY TransDate DESC, ID DESC"
Set rsmilk = oSaccoMaster.GetRecordset(sql)
                      
With rsmilk
If Not .EOF Then
lvwpreviousmilk.Visible = True
    With lvwpreviousmilk
         .ListItems.Clear
         .ColumnHeaders.Clear
         .ColumnHeaders.Add , , "KGS", 1000
         .ColumnHeaders.Add , , "TDATE", 2000
         .View = lvwReport
         .GridLines = True
    End With

End If
End With
     
With rsmilk
Do While Not .EOF
Set li = lvwpreviousmilk.ListItems.Add(, , rsmilk!QSupplied)
li.SubItems(1) = rsmilk!transdate & ""

rsmilk.MoveNext
Loop
End With
Exit Sub
End If

ErrorHandler:
MsgBox err.description
End Sub

Private Sub txtTCode_Validate(Cancel As Boolean)
Set rs = New ADODB.Recordset
sql = "d_sp_SelectTrans '" & txtTCode & "'"
Set rs = oSaccoMaster.GetRecordset(sql)
If Not rs.EOF Then
If Not IsNull(rs.Fields(0)) Then lblTName = rs.Fields(0)
Else
lblTName = ""
End If

End Sub

