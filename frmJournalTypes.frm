VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmJournalTypes 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "JOURNAL TYPES"
   ClientHeight    =   2205
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5640
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2205
   ScaleWidth      =   5640
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ListView ListView1 
      Height          =   1575
      Left            =   240
      TabIndex        =   6
      Top             =   120
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   2778
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
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ID"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Type"
         Object.Width           =   3528
      EndProperty
   End
   Begin VB.TextBox txtType 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1920
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   840
      Width           =   2415
   End
   Begin VB.TextBox txtJId 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1920
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   360
      Width           =   975
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   2040
      TabIndex        =   1
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "New"
      Height          =   375
      Left            =   720
      TabIndex        =   0
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Product Stage"
      Height          =   255
      Left            =   480
      TabIndex        =   4
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Product Stage"
      Height          =   255
      Left            =   480
      TabIndex        =   2
      Top             =   360
      Width           =   1215
   End
End
Attribute VB_Name = "frmJournalTypes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim edit As Boolean
Private Sub cmdNew_Click()
    On Error GoTo Capture
        ListView1.Visible = False
        txtJId.Text = ""
        txtType.Text = ""
    Exit Sub
Capture:
    MsgBox err.description
End Sub

Private Sub cmdSave_Click()
    On Error GoTo Capture
        If txtJId.Text = "" Then
            MsgBox "Enter the type id", vbCritical
            Exit Sub
        ElseIf txtType.Text = "" Then
            MsgBox "enter the name", vbCritical
            Exit Sub
        End If
        
        If Not edit Then
            Set rst = oSaccoMaster.GetRecordset("select * from JournalTypes where jid='" & txtJId & "'")
            If Not rst.EOF Then
                MsgBox "The type is already registered, or id is already used", vbCritical
                Exit Sub
            End If
        
            If Not oSaccoMaster.Execute("Insert into JournalTypes (Jid,Type)" _
            & " Values('" & txtJId & "','" & txtType & "')") Then
                MsgBox err.description
                Exit Sub
            End If
        Else
            If Not oSaccoMaster.Execute("update JournalTypes set Type='" & txtType & "' where jid='" & txtJId & "'") Then
                MsgBox err.description
                Exit Sub
            End If
        End If
        edit = False
        MsgBox "Record Saved successfully!", vbInformation
        LoadStages
        ListView1.Visible = True
        Exit Sub
        
    Exit Sub
Capture:
    MsgBox err.description
End Sub
Sub LoadStages()
        ListView1.ListItems.clear
        Set rst = oSaccoMaster.GetRecordset("select * from JournalTypes order by jid")
        While Not rst.EOF
            Set li = ListView1.ListItems.Add(, , rst!jid)
            li.ListSubItems.Add , , rst![Type]
            rst.MoveNext
        Wend
End Sub


Private Sub Form_Load()
    LoadStages
End Sub

Private Sub listview1_DblClick()
    On Error GoTo Capture
        With ListView1
            If .ListItems.Count = 0 Then
                Exit Sub
            End If
            
            txtJId.Text = .SelectedItem.Text
            txtType.Text = .SelectedItem.ListSubItems(1)
            
            edit = True
            ListView1.Visible = False
            
        End With
    Exit Sub
Capture:
    MsgBox err.description
End Sub

