VERSION 5.00
Begin VB.Form frmLog 
   Caption         =   "The Odyssey Online Classic [Message Logs]"
   ClientHeight    =   5175
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   6885
   LinkTopic       =   "Form1"
   ScaleHeight     =   5175
   ScaleWidth      =   6885
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Done"
      Height          =   495
      Left            =   4200
      TabIndex        =   2
      Top             =   4560
      Width           =   1335
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "&Clear"
      Height          =   495
      Left            =   1200
      TabIndex        =   1
      Top             =   4560
      Width           =   1455
   End
   Begin VB.TextBox txtLog 
      Height          =   4455
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   0
      Width           =   6855
   End
End
Attribute VB_Name = "frmLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClear_Click()
txtLog.Text = ""
End Sub

Private Sub cmdOk_Click()
Me.Hide
End Sub
