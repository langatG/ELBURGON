VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmproductdisposal 
   Caption         =   "Product Removal"
   ClientHeight    =   6045
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8085
   LinkTopic       =   "Form1"
   ScaleHeight     =   6045
   ScaleWidth      =   8085
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   855
      Left            =   1560
      MultiLine       =   -1  'True
      TabIndex        =   24
      Top             =   4440
      Width           =   5775
   End
   Begin VB.TextBox txtpprice 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   1560
      TabIndex        =   20
      Top             =   2160
      Width           =   1455
   End
   Begin VB.TextBox txtsellingprice 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   4680
      TabIndex        =   19
      Top             =   2040
      Width           =   1335
   End
   Begin VB.CommandButton cmdnew 
      Caption         =   "New"
      Height          =   375
      Left            =   240
      TabIndex        =   16
      Top             =   5520
      Width           =   975
   End
   Begin VB.CommandButton mm 
      Caption         =   "&Save"
      Height          =   375
      Left            =   1680
      TabIndex        =   15
      Top             =   5520
      Width           =   975
   End
   Begin VB.CommandButton cmdclose 
      Caption         =   "Close"
      Height          =   375
      Left            =   4320
      TabIndex        =   14
      Top             =   5520
      Width           =   975
   End
   Begin VB.CommandButton cmddelete 
      Caption         =   "&Delete"
      Height          =   375
      Left            =   3000
      TabIndex        =   13
      Top             =   5520
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ComboBox cbobranch 
      Height          =   315
      ItemData        =   "frmproductdisposal.frx":0000
      Left            =   1560
      List            =   "frmproductdisposal.frx":0007
      TabIndex        =   11
      Top             =   3840
      Width           =   1935
   End
   Begin VB.TextBox txtquantity 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   1560
      TabIndex        =   9
      Top             =   3000
      Width           =   1455
   End
   Begin VB.TextBox txtbalance 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   375
      Left            =   4680
      TabIndex        =   8
      Top             =   3000
      Width           =   1335
   End
   Begin VB.TextBox txtpassit 
      Appearance      =   0  'Flat
      Height          =   405
      IMEMode         =   3  'DISABLE
      Left            =   6960
      PasswordChar    =   "*"
      TabIndex        =   7
      Top             =   2400
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtpcode 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   1560
      TabIndex        =   3
      Top             =   840
      Width           =   1215
   End
   Begin VB.TextBox txtpname 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   1560
      TabIndex        =   2
      Top             =   1440
      Width           =   4695
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      DrawStyle       =   2  'Dot
      DrawWidth       =   17015
      Height          =   360
      Left            =   2760
      Picture         =   "frmproductdisposal.frx":000F
      ScaleHeight     =   360
      ScaleWidth      =   255
      TabIndex        =   0
      Top             =   840
      Width           =   255
   End
   Begin MSComCtl2.DTPicker txtdateenterered 
      Height          =   375
      Left            =   4920
      TabIndex        =   1
      Top             =   960
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   393216
      Format          =   121765889
      CurrentDate     =   38814
   End
   Begin VB.Label Label7 
      Caption         =   "Remarks:"
      Height          =   255
      Left            =   120
      TabIndex        =   23
      Top             =   4680
      Width           =   975
   End
   Begin VB.Label Label9 
      Caption         =   "Purchase Price "
      Height          =   375
      Left            =   120
      TabIndex        =   22
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label Label10 
      Caption         =   "Selling Price "
      Height          =   255
      Left            =   3240
      TabIndex        =   21
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0FFC0&
      Caption         =   "PRODUCT REMOVAL FORM."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   600
      TabIndex        =   18
      Top             =   120
      Width           =   4455
   End
   Begin VB.Label Label5 
      Caption         =   "Balance In Store"
      Height          =   255
      Left            =   3240
      TabIndex        =   17
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Label Label14 
      Caption         =   "Branch"
      Height          =   255
      Left            =   240
      TabIndex        =   12
      Top             =   3840
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "Quantity"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Product Code"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Product Name"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label Label6 
      Caption         =   "Date Entered"
      Height          =   255
      Left            =   3720
      TabIndex        =   4
      Top             =   960
      Width           =   1095
   End
End
Attribute VB_Name = "frmproductdisposal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    txtdateenterered = Date
    Set rst = New Recordset
    sql = "Select BName from d_Branch order by BName asc "
    Set rst = oSaccoMaster.GetRecordset(sql)
    While Not rst.EOF
        cbobranch.AddItem rst.Fields(0)
        rst.MoveNext
    Wend
End Sub

Private Sub mm_Click()
If Text1 = "" Then
 MsgBox "Please provide the Reason", vbInformation
 Exit Sub
End If
If txtquantity = "" Then txtquantity = 0

If txtquantity < 1 Then
 MsgBox "Please provide the Quantity", vbInformation
 Exit Sub
End If

If txtquantity > txtbalance Then
 MsgBox "Please quantity to be remove should not be more than the balance", vbInformation
 Exit Sub
End If

If txtpcode = "" Then
 MsgBox "Please provide the Product", vbInformation
 Exit Sub
End If

If cbobranch = "" Then
 MsgBox "Please provide the Product", vbInformation
 Exit Sub
End If

   sql = "SELECT UserLoginID, UserGroup, SUPERUSER From UserAccounts where UserLoginID='" & User & "'"
    Set rs = oSaccoMaster.GetRecordset(sql)
    If Not rs.EOF Then
        If rs!UserGroup <> "Manager" Then
            MsgBox "Only Manager allowed to Remove stock", vbInformation
            Exit Sub
        End If
    End If

'// insert into ag_products4
    sql = ""
    sql = "set dateformat dmy insert into ag_products4(p_code,p_name,s_no,qin,qout,date_entered,last_d_updated,user_id,audit_date,o_bal,branch,serialized,unserialized,seria,Remarks )"
    sql = sql & "  values('" & txtpcode & "','" & txtpname.Text & "','0','" & txtquantity.Text & "','" & txtquantity.Text & "','" & txtdateenterered.value & "','" & txtdateenterered.value & "','Admin','" & Date & "','" & txtquantity.Text & "','" & cbobranch & "','0','0','0','" & Text1 & "')"
    cn.Execute sql


    Dim quantity, qntbal As Double
    Set rst = New Recordset
    sql = "select * from ag_products where p_code='" & txtpcode & "' and Branch='" & cbobranch & "'"
    Set rst = oSaccoMaster.GetRecordset(sql)
    If rst.EOF Then
    quantity = rst!Qin - txtquantity
    qntbal = rst!o_bal - txtquantity
        sql = ""
        sql = "set dateformat DMY update ag_products set qin='" & quantity & "',qout='" & quantity & "',o_bal='" & qntbal & "' where p_code='" & txtpcode.Text & "' and branch='" & cbobranch & "'"
        cn.Execute sql
    End If

MsgBox "Product Remove Successfully", vbInformation

End Sub

Private Sub Picture2_Click()

If cbobranch = "" Then
MsgBox "Please select branch", vbInformation
Exit Sub
End If

frmSearch.Show vbModal
Dim Y As String
Y = sel

If Y <> "" Then

Provider = "MAZIWA"

Set cn = New ADODB.Connection
cn.Open Provider, "bi"
'If reportpath = "" Then reportpath = GetSetting("payroll", "AppName", "rptPath", rptPath)
sql = "select P_CODE,P_NAME,S_NO,QOUT,supplierID,pprice,sprice,QIN from ag_products where p_code='" & Y & "'AND Branch='" & cbobranch & "'"
Set rs = New ADODB.Recordset
rs.Open sql, cn
If Not rs.EOF Then
If Not IsNull(rs.Fields(0)) Then txtpcode = (rs.Fields(0))
If Not IsNull(rs.Fields(1)) Then txtpname = (rs.Fields(1))
If Not IsNull(rs.Fields(4)) Then cbosupplier = (rs.Fields(4))
If Not IsNull(rs.Fields(5)) Then txtpprice = (rs.Fields(5))
If Not IsNull(rs.Fields(6)) Then txtsellingprice = (rs.Fields(6))
If Not IsNull(rs.Fields(7)) Then txtreceived = (rs.Fields(7))
If Not IsNull(rs.Fields(3)) Then txtbalance = (rs.Fields(3))

If txtbalance <= 0 Then
MsgBox "Your stock is below zero please reorder", vbInformation
End If
'// check with serial no if it exist


End If
End If
End Sub

Private Sub txtpcode_KeyPress(KeyAscii As Integer)

If cbobranch = "" Then
MsgBox "Please select branch", vbInformation
Exit Sub
End If

If KeyAscii = 13 Then
txtpcode11_Change
'txtpcode11_KeyPress
Else
Exit Sub
End If
End Sub
Private Sub txtpcode11_Change()
'//TWNG001
Provider = "MAZIWA"
Set cn = New ADODB.Connection
cn.Open Provider, "bi"
'If reportpath = "" Then reportpath = GetSetting("payroll", "AppName", "rptPath", rptPath)
sql = ""
sql = "select P_CODE,P_NAME,S_NO,QOUT,supplierid,pprice, sprice,QIN from ag_products where p_code='" & txtpcode & "'AND Branch='" & cbobranch & "' "
Set rs = New ADODB.Recordset
rs.Open sql, cn
If Not rs.EOF Then
 txtpcode = (rs.Fields(0))
If Not IsNull(rs.Fields(1)) Then txtpname = (rs.Fields(1))
If Not IsNull(rs.Fields(3)) Then txtbalance = (rs.Fields(3))
If Not IsNull(rs.Fields(5)) Then txtpprice = (rs.Fields(5))
If Not IsNull(rs.Fields(6)) Then txtsellingprice = (rs.Fields(6))
If txtbalance <= 0 Then
MsgBox "Warning:Your stock is below zero please reorder", vbInformation
Else

End If
End If


'// check with serial no if it exist
End Sub
