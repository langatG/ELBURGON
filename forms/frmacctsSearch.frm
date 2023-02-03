VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAcctsSearch 
   Caption         =   "Search Accounts"
   ClientHeight    =   4965
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8025
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form7"
   MaxButton       =   0   'False
   ScaleHeight     =   4965
   ScaleWidth      =   8025
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdSelect 
      Caption         =   "&Select"
      Default         =   -1  'True
      Height          =   390
      Left            =   2445
      TabIndex        =   6
      Top             =   4320
      Width           =   1455
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   390
      Left            =   4215
      TabIndex        =   5
      Top             =   4320
      Width           =   1455
   End
   Begin VB.TextBox txtValue 
      Height          =   330
      Left            =   2385
      TabIndex        =   0
      Top             =   480
      Width           =   3405
   End
   Begin VB.ComboBox cboSearchField 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      ItemData        =   "frmacctsSearch.frx":0000
      Left            =   135
      List            =   "frmacctsSearch.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   480
      Width           =   2160
   End
   Begin MSComctlLib.ListView lvwAccounts 
      Height          =   3195
      Left            =   105
      TabIndex        =   1
      Top             =   900
      Width           =   7830
      _ExtentX        =   13811
      _ExtentY        =   5636
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Account No"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Account Name"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Account Group"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Account Sub Group"
         Object.Width           =   0
      EndProperty
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Search Value"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   2430
      TabIndex        =   4
      Top             =   225
      Width           =   1140
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Search Field"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   150
      TabIndex        =   3
      Top             =   225
      Width           =   1050
   End
End
Attribute VB_Name = "frmAcctsSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdcancel_Click()
    continue = False
    Unload Me
End Sub

Private Sub cmdSelect_Click()
    On Error GoTo SysError
    If lvwAccounts.ListItems.Count > 0 Then
        continue = True
        SearchValue = lvwAccounts.SelectedItem
        'AccountName = lvwAccounts.SelectedItem.SubItems(1)
    End If
    Unload Me
    Exit Sub
SysError:
    MsgBox err.description, vbInformation, Me.Caption
End Sub

Private Sub Form_Load()
    cboSearchField.ListIndex = 0
    txtvalue_Change
End Sub

Private Sub lvwAccounts_DblClick()
    cmdSelect_Click
End Sub

Private Sub txtvalue_Change()
    On Error Resume Next
    Dim rsAccts As New Recordset
    lvwAccounts.ListItems.clear
    Select Case cboSearchField.ListIndex
        Case 0 'Account Name
        Set rsAccts = oSaccoMaster.GetRecordset("Select AccNo,GLAccName,GLAccGroup,GLAccGroup From GLSETUP " _
        & "Where GLAccName Like '%" & txtvalue & "%' Order By GLAccName")
        Case 1 'Account No
        Set rsAccts = oSaccoMaster.GetRecordset("Select AccNo,GLAccName,GLAccGroup From GLSETUP " _
        & "Where AccNo Like '%" & txtvalue & "%' Order By AccNo")
    End Select
    With rsAccts
        While Not .EOF
            Set li = lvwAccounts.ListItems.Add(, , IIf(IsNull(!AccNo), "", !AccNo))
            li.SubItems(1) = IIf(IsNull(!GlAccName), "", !GlAccName)
            li.SubItems(2) = IIf(IsNull(!GlAccGroup), "", !GlAccGroup)
            li.SubItems(3) = IIf(IsNull(!GlAccGroup), "", !GlAccGroup)
            .MoveNext
        Wend
    End With
    Exit Sub
SysError:
    MsgBox err.description, vbInformation, Me.Caption
End Sub
