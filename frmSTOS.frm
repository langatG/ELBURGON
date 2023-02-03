VERSION 5.00
Begin VB.Form frmSTOS 
   Caption         =   "Standing orders"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6390
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   6390
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame FraSTO 
      Caption         =   "STOS"
      Height          =   2895
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6255
      Begin VB.CommandButton Command2 
         Caption         =   "Transporters Standing order"
         Height          =   735
         Left            =   3360
         TabIndex        =   2
         Top             =   960
         Width           =   2655
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Supplier Standing order"
         Height          =   735
         Left            =   360
         TabIndex        =   1
         Top             =   960
         Width           =   2295
      End
   End
End
Attribute VB_Name = "frmSTOS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
frmstandingorders.Show vbModal
End Sub

Private Sub Command2_Click()
frmtransportersstos.Show vbModal
End Sub
