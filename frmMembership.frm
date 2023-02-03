VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmMembership 
   Caption         =   "Membership"
   ClientHeight    =   3360
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8520
   LinkTopic       =   "Form1"
   ScaleHeight     =   3360
   ScaleWidth      =   8520
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdnewmembe 
      Caption         =   "New Members"
      Height          =   495
      Left            =   6480
      TabIndex        =   8
      Top             =   2280
      Width           =   1815
   End
   Begin VB.CommandButton cmdinactivemem 
      Caption         =   "InActive Members"
      Height          =   495
      Left            =   4680
      TabIndex        =   7
      Top             =   2280
      Width           =   1575
   End
   Begin VB.CommandButton cmdactivemem 
      Caption         =   "Active Members"
      Height          =   495
      Left            =   2520
      TabIndex        =   6
      Top             =   2280
      Width           =   1815
   End
   Begin MSComCtl2.DTPicker DTPenddate 
      Height          =   375
      Left            =   4800
      TabIndex        =   3
      Top             =   600
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393216
      Format          =   131006465
      CurrentDate     =   43784
   End
   Begin MSComCtl2.DTPicker DTPstartdate 
      Height          =   375
      Left            =   1200
      TabIndex        =   2
      Top             =   600
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393216
      Format          =   131006465
      CurrentDate     =   43784
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Load Records"
      Height          =   495
      Left            =   2760
      TabIndex        =   1
      Top             =   1200
      Width           =   1815
   End
   Begin VB.CommandButton cmdMemberregister 
      Caption         =   "Members Register"
      Height          =   495
      Left            =   600
      TabIndex        =   0
      Top             =   2280
      Width           =   1695
   End
   Begin VB.Label Label3 
      Caption         =   "Please wait to complete"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   1920
      TabIndex        =   9
      Top             =   1800
      Width           =   3855
   End
   Begin VB.Label Label2 
      Caption         =   "End date"
      Height          =   375
      Left            =   4680
      TabIndex        =   5
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Start date"
      Height          =   255
      Left            =   1560
      TabIndex        =   4
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "frmMembership"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdactivemem_Click()
reportname = "suppliersregisterActive.rpt"
Show_Sales_Crystal_Report "", reportname, ""
End Sub

Private Sub cmdinactivemem_Click()
reportname = "suppliersregisterInActive.rpt"
Show_Sales_Crystal_Report "", reportname, ""
End Sub

Private Sub cmdMemberregister_Click()
reportname = "suppliersregister.rpt"

 Show_Sales_Crystal_Report "", reportname, ""

End Sub

Private Sub cmdnewmembe_Click()
reportname = "suppliersregisternew.rpt"
Show_Sales_Crystal_Report "", reportname, ""
End Sub

Private Sub Command1_Click()
Dim lastdate As Date
Dim lastdateofsale As Date
Dim sno As String
Dim pname As String
Dim dy As Integer
Dim grade As String
Dim curamt As Double
Dim quantity As Double
Dim id As String

Dim rsd As New ADODB.Recordset

Label3.Visible = True

sql = ""
sql = "set dateformat dmy delete From d_SuppliersDetails " 'where Startdate>='" & DTPstartdate & "' and Enddate<='" & DTPenddate & "'
Set rs = oSaccoMaster.GetRecordset(sql)
sql = ""
'sql = "set dateformat dmy SELECT     SUM(m.QSupplied) AS Total, m.SNo, s.Names, s.IdNo FROM         dbo.d_Milkintake AS m INNER JOIN dbo.d_Suppliers AS s ON s.SNo = m.SNo WHERE     (m.TransDate >= '" & DTPstartdate & "') AND (m.TransDate <= '" & DTPenddate & "') GROUP BY m.SNo, s.Names, s.IdNo"
sql = "set dateformat dmy SELECT distinct(sno) from d_Suppliers where Regdate<='" & DTPenddate & "'"
Set rs = oSaccoMaster.GetRecordset(sql)
While Not rs.EOF
     sno = rs!sno
    sql = "set dateformat dmy SELECT top(1)* from d_Milkintake where SNo='" & sno & "' and TransDate>='" & DTPstartdate & "' and TransDate<='" & DTPenddate & "'"
    Set Rst = oSaccoMaster.GetRecordset(sql)
    If Not Rst.EOF Then
       sql = "set dateformat dmy insert into d_SuppliersDetails (sno , Startdate, Enddate, Remarks)"
       sql = sql & "values('" & sno & "','" & DTPstartdate & "','" & DTPenddate & "','Active') "
       oSaccoMaster.ExecuteThis (sql)
    Else
        sql = "set dateformat dmy insert into d_SuppliersDetails (sno , Startdate, Enddate, Remarks)"
        sql = sql & "values('" & sno & "','" & DTPstartdate & "','" & DTPenddate & "','InActive') "
        oSaccoMaster.ExecuteThis (sql)
    End If
rs.MoveNext
Wend
MsgBox "Records successfully done", vbInformation

Label3.Visible = False


End Sub

Private Sub Form_Load()
DTPenddate = Date
DTPstartdate = Date
Label3.Visible = False
End Sub
