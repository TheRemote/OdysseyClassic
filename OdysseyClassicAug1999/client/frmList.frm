VERSION 5.00
Begin VB.Form frmList 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "The Odyssey Online Classic"
   ClientHeight    =   5400
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3225
   ControlBox      =   0   'False
   Icon            =   "frmList.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5400
   ScaleWidth      =   3225
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lstHalls 
      Height          =   4545
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.ListBox lstBans 
      Height          =   4545
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.ListBox lstNPCs 
      Height          =   4545
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.ListBox lstObjects 
      Height          =   4545
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.ListBox lstMonsters 
      Height          =   4545
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.CommandButton btnOk 
      Caption         =   "Edit"
      Enabled         =   0   'False
      Height          =   495
      Left            =   1680
      TabIndex        =   2
      Top             =   4800
      Width           =   1455
   End
   Begin VB.CommandButton btnCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   4800
      Width           =   1455
   End
End
Attribute VB_Name = "frmList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub btnCancel_Click()
    Me.Hide
End Sub

Private Sub btnOk_Click()
    If lstObjects.Visible Then
        SendSocket Chr$(19) + Chr$(lstObjects.ListIndex + 1)
    ElseIf lstMonsters.Visible Then
        SendSocket Chr$(20) + Chr$(lstMonsters.ListIndex + 1)
    ElseIf lstNPCs.Visible Then
        SendSocket Chr$(50) + Chr$(lstNPCs.ListIndex + 1)
    ElseIf lstHalls.Visible = True Then
        SendSocket Chr$(48) + Chr$(lstHalls.ListIndex + 1)
    ElseIf lstBans.Visible Then
        SendSocket Chr$(57) + Chr$(lstBans.ItemData(lstBans.ListIndex))
    End If
    Me.Hide
End Sub

Private Sub Form_Load()
    Dim A As Long
    For A = 1 To 255
        lstMonsters.AddItem CStr(A) + ": " + Monster(A).Name
        lstObjects.AddItem CStr(A) + ": " + Object(A).Name
        lstHalls.AddItem CStr(A) + ": " + Hall(A).Name
        lstNPCs.AddItem CStr(A) + ": " + NPC(A).Name
    Next A
    frmList_Loaded = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmList_Loaded = False
End Sub


Private Sub lstBans_Click()
    btnOk.Enabled = True
End Sub

Private Sub lstBans_DblClick()
    btnOk_Click
End Sub


Private Sub lstHalls_Click()
    btnOk.Enabled = True
End Sub

Private Sub lstHalls_DblClick()
    btnOk_Click
End Sub


Private Sub lstMonsters_Click()
    btnOk.Enabled = True
End Sub

Private Sub lstMonsters_DblClick()
    btnOk_Click
End Sub


Private Sub lstNPCs_Click()
    btnOk.Enabled = True
End Sub

Private Sub lstNPCs_DblClick()
    btnOk_Click
End Sub


Private Sub lstObjects_Click()
    btnOk.Enabled = True
End Sub


Private Sub lstObjects_DblClick()
    btnOk_Click
End Sub


