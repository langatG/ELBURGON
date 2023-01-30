VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmEnquery 
   Caption         =   "Farmers Details "
   ClientHeight    =   10515
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   17820
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H0000FFFF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   10515
   ScaleWidth      =   17820
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin VB.TextBox TXTshares 
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   12600
      Locked          =   -1  'True
      TabIndex        =   42
      Top             =   1440
      Width           =   1695
   End
   Begin VB.TextBox txtcanno 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   4680
      TabIndex        =   41
      Top             =   1320
      Width           =   2415
   End
   Begin VB.CommandButton cmdExport 
      Caption         =   "&Export To Excel"
      Height          =   345
      Left            =   5475
      TabIndex        =   39
      Top             =   9855
      Width           =   2220
   End
   Begin VB.TextBox TXTIDNO 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   840
      TabIndex        =   27
      Top             =   1200
      Width           =   2175
   End
   Begin VB.TextBox txtLocation 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   720
      TabIndex        =   24
      Top             =   480
      Width           =   1575
   End
   Begin VB.PictureBox Picture5 
      Height          =   255
      Left            =   1920
      Picture         =   "frmEnquiry.frx":0000
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   23
      Top             =   120
      Width           =   255
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   120
      TabIndex        =   20
      Top             =   1560
      Width           =   7455
      Begin VB.CommandButton cmdShow 
         Caption         =   "Show"
         Default         =   -1  'True
         Height          =   375
         Left            =   6000
         TabIndex        =   4
         Top             =   240
         Width           =   1335
      End
      Begin MSComCtl2.DTPicker dtpFrom 
         Height          =   375
         Left            =   1320
         TabIndex        =   2
         Top             =   240
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         Format          =   130875393
         CurrentDate     =   40157
      End
      Begin MSComCtl2.DTPicker dtpTo 
         Height          =   375
         Left            =   4080
         TabIndex        =   3
         Top             =   240
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         Format          =   130875393
         CurrentDate     =   40157
      End
      Begin VB.Label Label10 
         Caption         =   "Date To"
         Height          =   255
         Left            =   3120
         TabIndex        =   22
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label9 
         Caption         =   "Date From"
         Height          =   375
         Left            =   240
         TabIndex        =   21
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.TextBox txtAccNo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   9000
      TabIndex        =   19
      Top             =   960
      Width           =   3015
   End
   Begin VB.TextBox txtBBranch 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   4200
      TabIndex        =   17
      Top             =   960
      Width           =   2175
   End
   Begin VB.TextBox txtSNo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   840
      TabIndex        =   1
      Top             =   120
      Width           =   1095
   End
   Begin VB.TextBox txtTransport 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   9000
      TabIndex        =   14
      Top             =   120
      Width           =   3015
   End
   Begin VB.TextBox txtBox 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   3120
      TabIndex        =   13
      Top             =   480
      Width           =   4575
   End
   Begin VB.TextBox txtTelNo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   9000
      TabIndex        =   12
      Top             =   480
      Width           =   3015
   End
   Begin VB.TextBox txtBank 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   840
      TabIndex        =   11
      Top             =   840
      Width           =   2175
   End
   Begin VB.TextBox txtName 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   3120
      TabIndex        =   10
      Top             =   120
      Width           =   4575
   End
   Begin MSComctlLib.ListView lvwEnguery 
      Height          =   6615
      Left            =   120
      TabIndex        =   0
      Top             =   2400
      Width           =   14535
      _ExtentX        =   25638
      _ExtentY        =   11668
      View            =   3
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HotTracking     =   -1  'True
      HoverSelection  =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Description"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Milk Intake (Kgs)"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "CR"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "DR"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Balance"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2880
      Top             =   9840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label21 
      Caption         =   "TOTAL SHARES"
      Height          =   615
      Left            =   12480
      TabIndex        =   43
      Top             =   960
      Width           =   2055
   End
   Begin VB.Label Label20 
      Caption         =   "Can Number"
      Height          =   255
      Left            =   3240
      TabIndex        =   40
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Label Label19 
      Caption         =   "Deductions"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7680
      TabIndex        =   38
      Top             =   9240
      Width           =   855
   End
   Begin VB.Label Label18 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   8640
      TabIndex        =   37
      Top             =   9240
      Width           =   2055
   End
   Begin VB.Label Label17 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   5280
      TabIndex        =   36
      Top             =   9240
      Width           =   1935
   End
   Begin VB.Label Label16 
      Caption         =   "Gross"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4440
      TabIndex        =   35
      Top             =   9240
      Width           =   735
   End
   Begin VB.Label Label13 
      Caption         =   "Total Kgs"
      Height          =   375
      Left            =   240
      TabIndex        =   34
      Top             =   9240
      Width           =   1335
   End
   Begin VB.Label lblTKgs 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "0.00"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1560
      TabIndex        =   33
      Top             =   9240
      Width           =   2415
   End
   Begin VB.Label Label15 
      Caption         =   "Deductions"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   -3000
      TabIndex        =   32
      Top             =   10080
      Width           =   855
   End
   Begin VB.Label lblDeductions 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   -2040
      TabIndex        =   31
      Top             =   10080
      Width           =   2055
   End
   Begin VB.Label lblGross 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   -5400
      TabIndex        =   30
      Top             =   10080
      Width           =   1935
   End
   Begin VB.Label Label14 
      Caption         =   "Gross"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   -6240
      TabIndex        =   29
      Top             =   10080
      Width           =   735
   End
   Begin VB.Label Label12 
      Caption         =   "ID:"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   28
      Top             =   1200
      Width           =   495
   End
   Begin VB.Label lblNPay 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C0C0&
      Height          =   615
      Left            =   7920
      TabIndex        =   26
      Top             =   1440
      Width           =   4095
   End
   Begin VB.Label Label11 
      Caption         =   "Loc :"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   25
      Top             =   480
      Width           =   615
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Account Number :"
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
      Left            =   7080
      TabIndex        =   18
      Top             =   960
      Width           =   1800
   End
   Begin VB.Label Label7 
      Caption         =   "Branch :"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3240
      TabIndex        =   16
      Top             =   960
      Width           =   855
   End
   Begin VB.Label Label6 
      Caption         =   "SNo :"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label5 
      Caption         =   "Transport :"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7800
      TabIndex        =   9
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label4 
      Caption         =   "Telephone :"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7800
      TabIndex        =   8
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Box :"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2400
      TabIndex        =   7
      Top             =   480
      Width           =   495
   End
   Begin VB.Label Label2 
      Caption         =   "Name :"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2280
      TabIndex        =   6
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Bank :"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   840
      Width           =   1215
   End
End
Attribute VB_Name = "frmEnquery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdExport_Click()
  On Error GoTo SsyError
  Dim sno As String
  Dim NAMES As String
  sno = txtSNo
  NAMES = txtName
    Dim MyFso As New FileSystemObject, strData As String, MFile As TextStream, _
    FileName As String, I As Long, li As ListItem
    If lvwEnguery.ListItems.Count > 0 Then
        With CommonDialog1
            .Filter = "Comma Seperated Values|*.csv"
            .FileName = "STATEMENT"
            .ShowSave
            
            If .FileName <> "" Then
                FileName = .FileName
            End If
            .FileName = ""
        End With
        Set MFile = MyFso.OpenTextFile(FileName, ForWriting, True)
        strData = "Supplier No  :" & sno
        MFile.WriteLine strData
        strData = "Supplier Name:" & NAMES
        MFile.WriteLine strData
        strData = "---------------------------------------------"
        MFile.WriteLine strData
        strData = ""
        'strData = "Period for" - "& dtpFrom &" & "to" & "& dtpto &"
        strData = "Transdate    ,Description         ,Intake     ,DEBIT,CREDIT,Balance"
        MFile.WriteLine strData
        strData = ""
        For I = 1 To lvwEnguery.ListItems.Count
            Set li = lvwEnguery.ListItems(I)
            strData = li & "," & li.SubItems(1) & "," & CDbl(li.SubItems(2)) & "," & CDbl(li.SubItems(3)) _
            & "," & (li.SubItems(4)) & "," & li.SubItems(5)
            MFile.WriteLine strData
            strData = ""
        Next I
    Else
        MsgBox "There are no records to be exported", vbInformation, Me.Caption
    End If
    MsgBox "Items Successfully Imported Into CSV file", vbOKOnly
    Exit Sub
SsyError:
    MsgBox err.description, vbInformation, Me.Caption

End Sub

Private Sub cmdShow_Click()
txtSNo_Validate True
End Sub

Private Sub Form_Activate()
txtSNo.SetFocus
End Sub

Private Sub Form_Load()
dtpFrom = Format(Get_Server_Date, "dd/mm/yyyy")
dtpFrom = DateSerial(year(dtpFrom), month(dtpFrom), 1)
dtpTo = DateSerial(year(dtpFrom), month(dtpFrom) + 1, 1 - 1)

 
End Sub

Private Sub Picture5_Click()
    
Me.MousePointer = vbHourglass
        frmSearchSupplier.Show vbModal
        txtSNo = sel
        txtSNo_Validate True
        Me.MousePointer = 0
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
On Error GoTo errmsg
Dim a, t As Boolean
If Trim(txtSNo) = "" Then
            txtSNo.SetFocus
        Exit Sub
    End If

txtName = ""
txtAccNo = ""
txtBank = ""
txtBBranch = ""
txtBox = ""
TXTIDNO = ""
txtLocation = ""
txtTelNo = ""
txtTransport = ""
lvwEnguery.ListItems.Clear
lblNPay = "0.00"
    
Set rs = New ADODB.Recordset
sql = "d_sp_SupplierEnquiry " & txtSNo & ""
Set rs = oSaccoMaster.GetRecordset(sql)
If rs.EOF Then
    MsgBox "There is no supplier with number " & txtSNo
    Exit Sub
End If
If Not rs.EOF Then
' [Names], AccNo, Bcode, BBranch, Location, PhoneNo, Address + ' ' + Town AS ADDRESS
'FROM         d_Suppliers WHERE SNo= @SNo
If Not IsNull(rs.Fields(0)) Then txtName = rs.Fields(0)
If Not IsNull(rs.Fields(1)) Then txtAccNo = rs.Fields(1)
If Not IsNull(rs.Fields(2)) Then txtBank = rs.Fields(2)
If Not IsNull(rs.Fields(3)) Then txtBBranch = rs.Fields(3)
If Not IsNull(rs.Fields(4)) Then txtLocation = rs.Fields(4)
If Not IsNull(rs.Fields(5)) Then txtTelNo = rs.Fields(5)
If Not IsNull(rs.Fields(6)) Then txtBox = rs.Fields(6)
If Not IsNull(rs.Fields(7)) Then TXTIDNO = rs.Fields(7)
'If Not IsNull(rs.Fields(8)) Then txtcanno = rs.Fields(8)


 
Set rs = New ADODB.Recordset
sql = "d_sp_TransName " & txtSNo & ""
Set rs = oSaccoMaster.GetRecordset(sql)
If Not rs.EOF Then
If Not IsNull(rs.Fields(0)) Then txtTransport = rs.Fields(0)
End If
getshares
LoadData


  'SNo
End If
Exit Sub
errmsg:
MsgBox txtName & " did not supply milk between " & dtpFrom & " and " & dtpTo

End Sub
Private Sub getshares()
Dim bal As Double, rss As New Recordset, amt As Double, rsts As New Recordset, shareamt As Double
Set rsts = oSaccoMaster.GetRecordset("SELECT    SUM(Amount) AS amtt From d_sconribution WHERE     (transdescription LIKE '%shares%') AND (SNo = '" & txtSNo & "')")
If Not rsts.EOF Then
shareamt = IIf(IsNull(rsts!amtt), 0, rsts!amtt)
End If
Set rss = oSaccoMaster.GetRecordset("SELECT    SUM(Amount) AS amt From d_supplier_deduc WHERE     (Description LIKE '%shares%') AND (SNo = '" & txtSNo & "')")
If Not rss.EOF Then
TXTshares = IIf(IsNull(rss!amt), 0, rss!amt) + shareamt
End If
End Sub
Private Sub LoadData()
Dim bal As Double, rss As New Recordset, amt As Double, rsts As New Recordset, shareamt As Double
Set rs = oSaccoMaster.GetRecordset("d_sp_SupNet " & txtSNo & ",'" & dtpFrom & "','" & dtpTo & "', 0")

If Not IsNull(rs.Fields(0)) Then
lblTKgs = rs.Fields(0)
Else
lblTKgs = "0.00"
End If

If Not IsNull(rs.Fields(1)) Then
Label17 = rs.Fields(1)
Else
Label17 = "0.00"
End If

Set rs = oSaccoMaster.GetRecordset("d_sp_SupNet " & txtSNo & ",'" & dtpFrom & "','" & dtpTo & "', 1")
If rs.RecordCount > 0 Then
If Not IsNull(rs.Fields(0)) Then
Label18 = rs.Fields(0)
Else
Label18 = "0.00"
End If
End If

bal = 0
lvwEnguery.ListItems.Clear
oSaccoMaster.ExecuteThis ("DELETE FROM d_tmpEnquery")

oSaccoMaster.ExecuteThis ("d_sp_UpdatetmpEnquery " & txtSNo & ",'" & dtpFrom & "','" & dtpTo & "'")
oSaccoMaster.ExecuteThis ("d_sp_UpdatetmpEnqueryDed " & txtSNo & ",'" & dtpFrom & "','" & dtpTo & "'")
Dim Descrption

Set rs = oSaccoMaster.GetRecordset("SELECT TransDate, Description,Intake,CR,DR From d_tmpEnquery WHERE SNo=" & txtSNo & " ORDER BY TransDate")
With rs
While Not rs.EOF
   Descrption = !description
   If Trim(!description) = "HShares" Then
   Descrption = "Shares"
   End If
   
   If Trim(!description) = "TMShares" Then
   Descrption = "Registration"
   End If
   
   Set li = lvwEnguery.ListItems.Add(, , IIf(IsNull(!transdate), "", !transdate))
   li.SubItems(1) = IIf(IsNull(!description), "", Descrption)
   li.SubItems(2) = IIf(IsNull(!Intake), "", !Intake)
   li.SubItems(3) = IIf(IsNull(!cr), 0, !cr)
   li.SubItems(4) = IIf(IsNull(!dr), 0, !dr)
   bal = Format(bal + li.SubItems(3) - li.SubItems(4), "#,##0.00")
   li.SubItems(5) = bal
   .MoveNext

Wend
End With
lblNPay = "Net Pay :" & Format(bal, "#,##0.00")

End Sub
