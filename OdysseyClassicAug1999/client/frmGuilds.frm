VERSION 5.00
Begin VB.Form frmGuilds 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "The Odyssey Online Classic [Guilds]"
   ClientHeight    =   4365
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4395
   ControlBox      =   0   'False
   Icon            =   "frmGuilds.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4365
   ScaleWidth      =   4395
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnView 
      Cancel          =   -1  'True
      Caption         =   "View"
      Enabled         =   0   'False
      Height          =   495
      Left            =   2880
      TabIndex        =   2
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton btnClose 
      Caption         =   "Close"
      Height          =   495
      Left            =   2880
      TabIndex        =   1
      Top             =   720
      Width           =   1455
   End
   Begin VB.ListBox lstGuilds 
      Height          =   4155
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2655
   End
End
Attribute VB_Name = "frmGuilds"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub btnClose_Click()
    Unload Me
End Sub


Private Sub btnNew_Click()
    Me.Hide
End Sub


Private Sub btnView_Click()
    btnView.Enabled = False
    btnClose.Enabled = False
    lstGuilds.Enabled = False
    SendSocket Chr$(39) + Chr$(lstGuilds.ItemData(lstGuilds.ListIndex))
End Sub


Private Sub Form_Load()
    Dim A As Long
    
    For A = 1 To 255
        If Guild(A).Name <> "" Then
            lstGuilds.AddItem Guild(A).Name
            lstGuilds.ItemData(lstGuilds.ListCount - 1) = A
        End If
    Next A
    
    frmGuilds_Loaded = True
End Sub


Private Sub Form_LostFocus()
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmGuilds_Loaded = False
End Sub


Private Sub lstGuilds_Click()
    btnView.Enabled = True
End Sub


Private Sub lstGuilds_DblClick()
    btnView_Click
End Sub


