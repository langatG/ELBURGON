VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmImportLedgers 
   Caption         =   "Import Ledgers"
   ClientHeight    =   3990
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6735
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmImportLedgers.frx":0000
   LinkTopic       =   "Form6"
   ScaleHeight     =   3990
   ScaleWidth      =   6735
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   225
      Left            =   105
      TabIndex        =   6
      Top             =   3525
      Visible         =   0   'False
      Width           =   3510
      _ExtentX        =   6191
      _ExtentY        =   397
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3105
      Top             =   2565
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ListView lvwFields 
      Height          =   2310
      Left            =   105
      TabIndex        =   4
      Top             =   840
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   4075
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      HideColumnHeaders=   -1  'True
      Checkboxes      =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "m"
         Object.Width           =   8819
      EndProperty
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "..."
      Height          =   345
      Left            =   6165
      TabIndex        =   3
      Top             =   480
      Width           =   420
   End
   Begin VB.TextBox txtFileName 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   495
      Width           =   6015
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   390
      Left            =   5145
      TabIndex        =   1
      Top             =   3465
      Width           =   1425
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Default         =   -1  'True
      Height          =   390
      Left            =   3660
      TabIndex        =   0
      Top             =   3465
      Width           =   1425
   End
   Begin VB.Label Label1 
      Caption         =   $"frmImportLedgers.frx":0442
      Height          =   1500
      Left            =   3120
      TabIndex        =   5
      Top             =   960
      Width           =   3135
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmImportLedgers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOk_Click()
    On Error GoTo SysError
    If Trim$(txtFileName) <> "" Then
        Dim MyFso As New FileSystemObject, MFile As TextStream, strData As String, _
        dAmount As Double, sAccNo As String, iRecs As Long, mPos As Long, Cnn As _
        New Connection, TheAmount As Double, theBalance As Double
        With Cnn
            If .State = adStateClosed Then
                .Open SelectedDsn
            End If
        End With
        Set MFile = MyFso.OpenTextFile(txtFileName, ForReading, False)
        Do Until MFile.AtEndOfStream
            strData = MFile.ReadLine
            If strData = "" Then
                Exit Do
            End If
            iRecs = iRecs + 1
        Loop
        MsgBox "Amount Is " & Format(dAmount, CfMt)
        If iRecs > 0 Then
            ProgressBar1.Visible = True
            ProgressBar1.max = iRecs
        End If
        MFile.Close
        'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
        Set rs = oSaccoMaster.GetRecordSet("Delete From CustomerBalance")
        DoEvents
        Set MFile = MyFso.OpenTextFile(txtFileName, ForReading, False)
        Do Until MFile.AtEndOfStream
            mPos = mPos + 1
            ProgressBar1.Value = mPos
            DoEvents
            strData = MFile.ReadLine
            If strData <> "" Then
                sAccNo = Left(strData, InStr(1, strData, Chr(9), vbTextCompare) - 1)
                strData = Right(strData, Len(strData) - Len(sAccNo) - 1)
                dAmount = strData
                If Not GLAccount_Exists(sAccNo, ErrorMessage) Then
                    MsgBox ErrorMessage, vbInformation, Me.Caption
                    ErrorMessage = ""
                    ProgressBar1.Visible = False
                    Exit Do
                Else
                    Get_GL_AccDetails (sAccNo)
'                    If dAmount < 0 Then
'                        MsgBox ""
'                    End If
                    Select Case GlAccNBal
                        Case "DR" 'Debit Balance
                        If dAmount < 0 Then
                            TheAmount = dAmount * (-1)
                        Else
                            TheAmount = dAmount
                        End If
                        theBalance = dAmount
                        Case "CR" 'Credit Balance
                        If dAmount < 0 Then
                            TheAmount = dAmount * (-1)
                        Else
                            TheAmount = dAmount
                        End If
                        theBalance = dAmount * (-1)
                    End Select
                    If Not Save_CustBalance("", "", "", GlAccName, TheAmount, theBalance, _
                    sAccNo, "Opening Bal", "30/06/2007", 0, "", "06", 0, 0, IIf(dAmount > _
                    0, "DR", "CR"), 1, "Open Bal", User, "Ledger Import", "", "30/06/2007", _
                    theBalance, 0, "", 0, Cnn, ErrorMessage) Then
                        If ErrorMessage <> "" Then
                            MsgBox ErrorMessage, vbInformation, Me.Caption
                            ErrorMessage = ""
                        End If
                    End If
                    dAmount = 0
                End If
            End If
            strData = ""
        Loop
    End If
    ProgressBar1.Visible = False
    Exit Sub
SysError:
    ProgressBar1.Visible = False
    iRecs = 0
    mPos = 0
    MsgBox Err.description, vbInformation, Me.Caption
End Sub

Private Sub cmdopen_Click()
    On Error GoTo SysError
    With CommonDialog1
        .Filter = "Text Files|*.txt"
        .ShowOpen
        If .FileName <> "" Then
            txtFileName = .FileName
        End If
    End With
    Exit Sub
SysError:
    MsgBox Err.description, vbInformation, Me.Caption
End Sub

Private Sub Form_Load()
    On Error GoTo SysError
    With lvwFields
        .ListItems.Clear
        .ListItems.Add , , "Account No"
        .ListItems.Add , , "Amount"
    End With
    
    Exit Sub
SysError:
    MsgBox Err.description, vbInformation, Me.Caption
End Sub
