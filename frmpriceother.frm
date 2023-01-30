VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmprice 
   Caption         =   "Other Prices"
   ClientHeight    =   3915
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3915
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H80000018&
      Height          =   3495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4815
      Begin VB.CommandButton Command1 
         Caption         =   "Update Suppliers"
         Height          =   375
         Left            =   1200
         TabIndex        =   9
         Top             =   2040
         Width           =   2055
      End
      Begin VB.CommandButton cmdcLOSE 
         Caption         =   "Close"
         Height          =   375
         Left            =   2880
         TabIndex        =   4
         Top             =   3000
         Width           =   1335
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "Update"
         Height          =   375
         Left            =   360
         TabIndex        =   3
         Top             =   3000
         Width           =   1335
      End
      Begin VB.TextBox txtNewPrice 
         Height          =   285
         Left            =   1320
         TabIndex        =   2
         Top             =   840
         Width           =   1935
      End
      Begin VB.TextBox txtCurrentPrice 
         Height          =   285
         Left            =   1320
         TabIndex        =   1
         Top             =   360
         Width           =   1815
      End
      Begin MSComCtl2.DTPicker DTPStartFrom 
         Height          =   375
         Left            =   1320
         TabIndex        =   5
         Top             =   1320
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         _Version        =   393216
         Format          =   130809857
         CurrentDate     =   40095
      End
      Begin VB.Label Label3 
         BackColor       =   &H80000018&
         Caption         =   "Start From"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label Label2 
         BackColor       =   &H80000018&
         Caption         =   "New Price:"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000018&
         Caption         =   "Current Price"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   360
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmprice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub cmdclose_Click()
Unload Me
End Sub

Private Sub cmdupdate_Click()
On Error GoTo ErrorHandler

If Trim(txtNewPrice) = "" Then
MsgBox "Enter the new price."
txtNewPrice.SetFocus
Exit Sub
End If

If Not IsNumeric(txtNewPrice) Then
MsgBox "Please enter a number." & txtNewPrice & " is not a number", vbExclamation
txtNewPrice.SetFocus
Exit Sub
End If
sql = "Save_Price '" & DTPStartFrom & "'," & txtNewPrice & ""
oSaccoMaster.ExecuteThis (sql)

txtCurrentPrice = txtNewPrice
txtNewPrice = ""
'//select

Set rs = New ADODB.Recordset
sql = "MilkPrice '" & DTPStartFrom & "'"
Set rs = oSaccoMaster.GetRecordset(sql)
While Not rs.EOF

Set rst = New ADODB.Recordset
sql = "ChangePrice1 " & rs.Fields(0) & ",'" & rs.Fields(1) & "'," & CCur(txtCurrentPrice) & "," & rs.Fields(3) & "," & rs.Fields(4) & "," & rs.Fields(5) & ""
oSaccoMaster.ExecuteThis (sql)
frmPricing.Caption = "UPDATING SUPPLIER NUMBER "
frmPricing.Caption = frmPricing.Caption & " " & rs.Fields(0)


rs.MoveNext
Wend
frmPricing.Caption = rs.RecordCount & " Records Updated."
MsgBox "Records successively updated."
frmPricing.Caption = "Pricing Updates"
Exit Sub
ErrorHandler:
MsgBox err.description

End Sub

Private Sub Command1_Click()
Dim rsts2 As New Recordset
Dim sno As String
''**********************SHARES**********************'
sql = ""

sql = "set dateformat dmy select * from d_suppliers where Processor=1"
Set rsts2 = oSaccoMaster.GetRecordset(sql)
While Not rsts2.EOF
DoEvents
sno = rsts2.Fields("sno")


sql = "set dateformat dmy update  d_milkintake set processor=1 where  sno='" & sno & "'"
oSaccoMaster.ExecuteThis (sql)
'While Not rst.EOF
'DoEvents
'sno = rst.Fields("sno")
'
'sql = "set dateformat dmy select * from d_suppliers where sno='" & sno & "' and shares=1"
'Set rsts2 = oSaccoMaster.GetRecordset(sql)
'If Not rsts2.EOF Then



rsts2.MoveNext
Wend
MsgBox "Records successively updated."
End Sub

Private Sub Form_Load()
DTPStartFrom = Format(Get_Server_Date, "dd/mm/yyyy")
DTPStartFrom.MaxDate = Format(Get_Server_Date, "dd/mm/yyyy")
Set rs = New ADODB.Recordset
sql = "Pick_Current_Price"
Set rs = oSaccoMaster.GetRecordset(sql)
If Not rs.EOF Then
txtCurrentPrice = rs.Fields(0)
Else
txtCurrentPrice = 0
End If
txtCurrentPrice = Format(txtCurrentPrice, "####0.00")
End Sub

Private Sub txtCurrentPrice_Validate(Cancel As Boolean)
txtCurrentPrice = Format(txtCurrentPrice, "####0.00")
End Sub

Private Sub txtNewPrice_Validate(Cancel As Boolean)
txtNewPrice = Format(txtNewPrice, "####0.00")
End Sub

