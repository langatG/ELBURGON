VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Login"
   ClientHeight    =   2145
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   3930
   ForeColor       =   &H80000007&
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1267.337
   ScaleMode       =   0  'User
   ScaleWidth      =   3690.057
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkchangepassword 
      Caption         =   "Change"
      Height          =   255
      Left            =   3000
      TabIndex        =   17
      Top             =   1680
      Width           =   975
   End
   Begin VB.Frame Frame2 
      Height          =   1215
      Left            =   180
      TabIndex        =   13
      Top             =   180
      Width           =   3615
      Begin VB.TextBox txtPassword 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         IMEMode         =   3  'DISABLE
         Left            =   1215
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   743
         Width           =   2325
      End
      Begin VB.TextBox txtUserName 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1215
         TabIndex        =   0
         Top             =   240
         Width           =   2325
      End
      Begin VB.Label lblLabels 
         Caption         =   "&Password:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   1
         Left            =   120
         TabIndex        =   15
         Top             =   780
         Width           =   1080
      End
      Begin VB.Label lblLabels 
         Caption         =   "&User Name:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   0
         Left            =   120
         TabIndex        =   14
         Top             =   277
         Width           =   1080
      End
   End
   Begin VB.Frame Frame1 
      Height          =   15
      Left            =   0
      TabIndex        =   12
      Top             =   1440
      Width           =   3735
   End
   Begin VB.TextBox txtExpire 
      Height          =   285
      Left            =   5640
      TabIndex        =   6
      Top             =   1080
      Width           =   1695
   End
   Begin VB.CommandButton cmdChangePassword 
      Caption         =   "Change"
      Height          =   375
      Left            =   6240
      TabIndex        =   7
      Top             =   1560
      Width           =   1095
   End
   Begin VB.TextBox txtConfirmPassword 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   5640
      PasswordChar    =   "*"
      TabIndex        =   5
      Top             =   600
      Width           =   1695
   End
   Begin VB.TextBox txtNewPassword 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   5640
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   120
      Width           =   1695
   End
   Begin VB.CommandButton cmdOK 
      Appearance      =   0  'Flat
      Caption         =   "OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   195
      TabIndex        =   2
      Top             =   1590
      Width           =   1200
   End
   Begin VB.CommandButton cmdCancel 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   1680
      TabIndex        =   3
      Top             =   1620
      Width           =   1215
   End
   Begin VB.Label Label5 
      BackColor       =   &H80000015&
      Height          =   1335
      Left            =   120
      TabIndex        =   16
      Top             =   120
      Width           =   3735
   End
   Begin VB.Label Label4 
      Caption         =   "Days"
      Height          =   255
      Left            =   7440
      TabIndex        =   11
      Top             =   1080
      Width           =   375
   End
   Begin VB.Label Label3 
      Caption         =   "Password expires after"
      Height          =   255
      Left            =   3960
      TabIndex        =   10
      Top             =   1080
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "Confirm Password"
      Height          =   255
      Left            =   3960
      TabIndex        =   9
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "New Password"
      Height          =   255
      Left            =   3960
      TabIndex        =   8
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Group As String
Public LoginSucceeded As Boolean
Dim ExpiryDate As Integer
Dim results As String
Dim sql As String
Dim DateCreated As String
Dim PassStatus As String
Dim WithEvents myclass As cdbase
Attribute myclass.VB_VarHelpID = -1
Private Sub dismenu()
Dim I As Control
Dim intIncrement As Integer

For Each I In Controls
If TypeOf I Is Menu Then I.enable = False
Next
'

'
End Sub
Private Sub GetUserRights()
    On Error Resume Next
   
    Dim Provider As String
    Dim rs As Object
    Dim clsClass As cdbase
    Set clsClass = New cdbase
    Set cn = CreateObject("adodb.connection")
    If Provider = "" Then
        Provider = clsClass.OpenCon
    End If
     cn.Open Provider, "bi"
     
    '//disable all menus
    dismenu
    sql = ""
    sql = "select * from UserAccounts where UserName='" & results & "'"
    Set rs = CreateObject("adodb.recordset")
    Dim bbname As String
    rs.Open sql, cn
    Dim waldate As Integer
    If Not rs.EOF Then
    If Not IsNull(rs.Fields("username")) Then username = rs.Fields("username")
        If Not IsNull(rs!BranchCode) Then
            bcode = Trim(rs.Fields("branchcode"))
        End If
        Dim rsbr As Recordset
        sql = "select branchname from branches where branchcode='" & bcode & "'"
        Set rsbr = New ADODB.Recordset
        rsbr.Open sql, cn
        If rsbr.EOF Then
            bbname = "HEAD OFFICE"
        Else
            If Not IsNull(rsbr.Fields("branchname")) Then
                bbname = rsbr.Fields("branchname")
            Else
                bbname = "HEAD OFFICE"
            End If
        End If
       ' MainForm.StatusBar1.Panels(5).Text = bbname
       'MainForm.Show
    End If
    If rs.EOF Then
        MsgBox "User " & results & " has no login rights!", vbExclamation
        Exit Sub
    End If
    MainForm.StatusBar1.Panels(1).Text = rs!username
    MainForm.DTPPeriod.value = Format(Get_Server_Date, "dd/mm/yyyy")
    If rs!passexpire <> "" Then
        ExpiryDate = rs!passexpire
    End If
    Dim MyMonth As Integer
    waldate = DateDiff("d", Format(rs!DateCreated, "dd/mm/yyyy"), Format(Get_Server_Date, "dd/mm/yyyy"))
    If rs!DateCreated <> "" Then
        DateCreated = month(Format(rs!DateCreated, "dd/mm/yyyy"))
    End If
    MyMonth = month(Format(rs!DateCreated, "dd/mm/yyyy"))
    If waldate > ExpiryDate Then
        GoTo Expired
    End If
    sql = ""
    Dim rsg As Recordset
    sql = "select * from UserGroups where GroupName='" & rs.Fields("UserGroup") & "'"
    Set rsg = New ADODB.Recordset
    rsg.Open sql, cn
    Dim intMonth As Integer
    intMonth = CCur(month(Date) - CCur(MyMonth))
'    If intMonth <> 0 Then
        '// Means that there is a difference in months so automatically
        '//the password has expired
''        GoTo Expired
'    End If
    PassStatus = CCur(Day(Date)) - CCur(DateCreated)
    If PassStatus < 0 Then
        PassStatus = PassStatus * -1
    End If
    If CCur(PassStatus) > CCur(ExpiryDate) Then '// Password has expired
Expired:
        MsgBox "Password has expired.You are required to change your old password", vbExclamation: txtNewPassword.SetFocus
        Me.Width = 7920 '// Strech Size
        Me.Move 2600
        cmdOk.Default = False
        cmdChangePassword.Default = True
        Exit Sub
    Else '// Password has not expired

    End If
    Set rs = CreateObject("adodb.recordset")
    'rs.Open sql, cn
    'If rs.EOF Then
      '  Exit Sub
    'End If
    Dim myclass As cdbase
    Set myclass = New cdbase
    With myclass
        sql = "select * from UserGroups where GroupName='" & rs.Fields("UserGroup") & "'"
        Set rsg = New ADODB.Recordset
        rsg.Open sql, cn
        If Not rsg.EOF Then
        MainForm.Show
'        If rsg.Fields("CashBook") = True Then
'            'MainForm.mnucashbook1.Enabled = True
'        Else
'            'MainForm.mnucashbook1.Enabled = False
'        End If
'        ' If MDIACCOUNTSANDCASH.Progress1.MyProgressBar <> 20 Then MDIACCOUNTSANDCASH.Progress1.MyProgressBar = 16.66
'        If rsg.Fields("Transactions") = True Then
'            MainForm.mnutransactions.Enabled = True
'        Else
'            MainForm.mnutransactions.Enabled = False
'        End If
'        '        If MDIACCOUNTSANDCASH.Progress1.MyProgressBar <> 20 Then MDIACCOUNTSANDCASH.Progress1.MyProgressBar = 33.33
'        If rsg.Fields("activity") = True Then
'            MainForm.mnuActivities.Enabled = True
'        Else
'            MainForm.mnuActivities.Enabled = False
'        End If
'        If rsg.Fields("files") = True Then
'            MainForm.mnuFiles.Enabled = True
'        Else
'            MainForm.mnuFiles.Enabled = False
'        End If
'        If rsg.Fields("Reports") = True Then
'            MainForm.mnuReports.Enabled = True
'        Else
'            MainForm.mnuReports.Enabled = False
'        End If
'
'        If rsg.Fields("FixedAssets") = True Then
'            MainForm.mnufixedassetlistings.Enabled = True
'        Else
'            MainForm.mnufixedassetlistings.Enabled = False
'        End If
'
'        If rsg.Fields("Setup") = True Then
'            MainForm.mnuSetUp.Enabled = True
'        Else
'            MainForm.mnuSetUp.Enabled = False
'        End If
'        ' If RSG.Fields("Setup") = True Then
'        ' MDIACCOUNTSANDCASH.mnusetup1.Enabled = True
'        ' End If
'        If rsg.Fields("Accounts") = True Then
'        MainForm.mnuaccounts.Enabled = True
'        Else
'        MainForm.mnuaccounts.Enabled = False
'        End If
'
'       'Else
'            'MainForm.mnuOtherSchemes.Enabled = False
'            '  MainForm.mnurefunds.Enabled = False
'        'End If
'        If rsg.Fields("AccountsPay") = True Then
'        MainForm.mnuAccountpayable.Enabled = True
'        Else
'        MainForm.mnuAccountpayable.Enabled = False
'        End If
'
'        If rsg.Fields("FixedAssets") = True Then
'        MainForm.mnuassets.Enabled = True
'        Else
'        MainForm.mnuassets.Enabled = False
'        End If
   Unload Me
       MainForm.Show
'        Else
''        MainForm.mnuglaccounts.Enabled = False
''        MainForm.mnuOtherSchemes.Enabled = False
''        MainForm.mnurefunds.Enabled = False
''        MainForm.mnuSetUp.Enabled = False
''        MainForm.mnuReports.Enabled = False
''        MainForm.mnufile.Enabled = False
'        MainForm.mnumembership.Enabled = False
        End If
    End With
  
         
        
      

     
    '//get the company name
    Set myclass = Nothing
End Sub

Private Sub chkchangepassword_Click()
If chkchangepassword = vbChecked Then
Me.Width = 7920 '// Strech Size
        Me.Move 2600
        cmdOk.Default = False
        cmdChangePassword.Default = True
        Else
    Me.Width = 3780   '// Normal Size
    cmdOk.Default = True
    cmdChangePassword.Default = False
        End If
End Sub

Private Sub cmdCancel_Click()
    
    LoginSucceeded = False

    Unload Me
    
    End
    
End Sub

Private Sub cmdChangePassword_Click()
    'Dim Pass As EncryptDecrypt
    'Set Pass = New EncryptDecrypt
    'txtPassword = Pass.en(txtPassword)
    On Error Resume Next
    If txtExpire = "" Then
        MsgBox "Specify when password expires", vbExclamation
        txtExpire.SetFocus
        Exit Sub
    End If
    If txtNewPassword <> txtConfirmPassword Then
        MsgBox "Invalid confirmation password.Confirm again", vbExclamation
        txtConfirmPassword.SetFocus
        Exit Sub
    End If
    'txtNewPassword = Pass.Encrypt(txtNewPassword)
    txtPassword = modsecurity.Encript_String(txtPassword)
    txtNewPassword = modsecurity.Encript_String(txtNewPassword)
    sql = ""
    sql = "Set DateFormat DMY Update UserAccounts Set passwords='" & txtNewPassword _
    & "', PassExpire='" & txtExpire & "', DateCreated='" & Format(Get_Server_Date, _
    "dd-MM-yyyy") & "' where UserLoginID='" & txtUserName & "' and passwords='" & txtPassword & "'"
    Dim Kwenda As cdbase
    Set Kwenda = New cdbase
    Kwenda.save sql
    Kwenda.CloseCon
    Set Kwenda = Nothing
    Me.Width = 3780   '// Normal Size
    cmdOk.Default = True
    cmdChangePassword.Default = False
    Unload Me
    frmLogin.Show
End Sub

Private Sub cmdOk_Click()
    On Error GoTo ErrorHandler
    Dim myclass As cdbase
    Set myclass = New cdbase
    Dim Password As String
    Dim username As String
    Password = txtPassword
    username = txtUserName
    results = myclass.GetUsers(Password, username)
    User = username
    If results <> "" Then
        SaveSetting "1003", "BackUp", "ID", txtUserName
        GetUserRights
        Set rs = oSaccoMaster.GetRecordset("SELECT     CompanyName  FROM         SYSPARAM")
    If Not rs.EOF Then
        MainForm.StatusBar1.Panels(5).Text = rs.Fields(0)
    End If
    Else
        MsgBox "Invalid Password, try again!", , "Login"
        On Error Resume Next
        txtPassword.SetFocus
        SendKeys "{Home}+{End}"
    End If
    myclass.CloseCon
    Set myclass = Nothing
    Exit Sub
ErrorHandler:
    MsgBox err.description
End Sub

Private Sub Form_Load()
    
    On Error GoTo SysError
    txtUserName = GetSetting("1003", "BackUp", "ID")
    On Error Resume Next
    If Trim$(txtUserName) <> "" Then
        txtPassword.SetFocus
    Else
        txtUserName.SetFocus
        SendKeys "{Home}+{End}"
    End If
    If Trim$(txtUserName) <> "" Then
        txtPassword.TabIndex = 0
        SendKeys "{Home}+{End}"
    Else
        txtUserName.SetFocus
        SendKeys "{Home}+{End}"
    End If
    Exit Sub
SysError:
    MsgBox err.description, vbInformation, Me.Caption
End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
txtPassword_Change
End Sub

Private Sub txtPassword_Change()
If sbrCaps Then
    txtPassword.ToolTipText = "Num Lock is on"
    Else
    txtPassword.ToolTipText = ""
    End If
End Sub

Private Sub txtPassword_GotFocus()
txtPassword_Change
End Sub

Private Sub txtUserName_Change()
On Error Resume Next
If txtUserName <> "" Then
'TXTPASSWORD.SetFocus
End If
End Sub
