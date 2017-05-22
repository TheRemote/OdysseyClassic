VERSION 5.00
Begin VB.Form frmScan 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Scan"
   ClientHeight    =   6210
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4725
   Icon            =   "frmScan.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6210
   ScaleWidth      =   4725
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lstBank 
      Height          =   4155
      Left            =   2400
      TabIndex        =   7
      Top             =   1920
      Width           =   2175
   End
   Begin VB.ListBox lstInventory 
      Height          =   4155
      Left            =   120
      TabIndex        =   5
      Top             =   1920
      Width           =   2175
   End
   Begin VB.Label lblMaxMana 
      Height          =   255
      Left            =   4200
      TabIndex        =   15
      Top             =   1560
      Width           =   495
   End
   Begin VB.Label lblMaxEnergy 
      Height          =   255
      Left            =   2520
      TabIndex        =   14
      Top             =   1200
      Width           =   495
   End
   Begin VB.Label lblMaxHP 
      Height          =   255
      Left            =   960
      TabIndex        =   13
      Top             =   1200
      Width           =   495
   End
   Begin VB.Label cptMaxMana 
      Caption         =   "MaxMana:"
      Height          =   255
      Left            =   3240
      TabIndex        =   12
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label cptMaxEnergy 
      Caption         =   "MaxEnergy:"
      Height          =   255
      Left            =   1560
      TabIndex        =   11
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label cptMaxHP 
      Caption         =   "MaxHP:"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   1200
      Width           =   735
   End
   Begin VB.Label lblClass 
      Height          =   255
      Left            =   840
      TabIndex        =   9
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label cptClass 
      Caption         =   "Class:"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   840
      Width           =   615
   End
   Begin VB.Label cptBank 
      Alignment       =   2  'Center
      Caption         =   "Bank"
      Height          =   255
      Left            =   2400
      TabIndex        =   6
      Top             =   1560
      Width           =   2175
   End
   Begin VB.Label cptItems 
      Alignment       =   2  'Center
      Caption         =   "Inventory"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1560
      Width           =   2175
   End
   Begin VB.Label lblLevel 
      Height          =   255
      Left            =   960
      TabIndex        =   3
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label cptLevel 
      Caption         =   "Level:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   735
   End
   Begin VB.Label lblPlayer 
      Height          =   255
      Left            =   960
      TabIndex        =   1
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label cptPlayer 
      Caption         =   "Player:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "frmScan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Terminate()
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload Me
End Sub
