VERSION 5.00
Begin VB.Form frmMessages 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "The Odyssey Online Classic [Messages]"
   ClientHeight    =   4755
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5475
   ControlBox      =   0   'False
   Icon            =   "frmMessages.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4755
   ScaleWidth      =   5475
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnNew 
      Cancel          =   -1  'True
      Caption         =   "New Message"
      Height          =   375
      Left            =   2760
      TabIndex        =   4
      Top             =   4320
      Width           =   1215
   End
   Begin VB.CommandButton btnReply 
      Caption         =   "Reply"
      Height          =   375
      Left            =   1440
      TabIndex        =   3
      Top             =   4320
      Width           =   1215
   End
   Begin VB.CommandButton btnDelete 
      Caption         =   "Delete"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   4320
      Width           =   1215
   End
   Begin VB.CommandButton btnOk 
      Caption         =   "Close"
      Height          =   375
      Left            =   4200
      TabIndex        =   5
      Top             =   4320
      Width           =   1215
   End
   Begin VB.TextBox txtMessages 
      Height          =   2175
      Left            =   120
      Locked          =   -1  'True
      MaxLength       =   1024
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   2040
      Width           =   5295
   End
   Begin VB.ListBox lstMessages 
      Height          =   1425
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   5295
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   5295
   End
End
Attribute VB_Name = "frmMessages"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub btnNew_Click()
    frmMessage.Show 1
End Sub

Private Sub btnOk_Click()
    Unload Me
End Sub


