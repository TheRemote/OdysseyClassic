VERSION 5.00
Begin VB.Form frmGuild 
   Caption         =   "The Odyssey Online Classic [Guild]"
   ClientHeight    =   4140
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7185
   ControlBox      =   0   'False
   Icon            =   "frmGuild.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4140
   ScaleWidth      =   7185
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnRank 
      Caption         =   "Founder"
      Enabled         =   0   'False
      Height          =   495
      Index           =   3
      Left            =   6360
      TabIndex        =   8
      Top             =   3000
      Width           =   735
   End
   Begin VB.CommandButton btnMoveOut 
      Caption         =   "<-- Move Out"
      Enabled         =   0   'False
      Height          =   495
      Left            =   2520
      TabIndex        =   17
      Top             =   360
      Width           =   1215
   End
   Begin VB.CommandButton btnRemoveDeclaration 
      Caption         =   "Remove Declaration"
      Enabled         =   0   'False
      Height          =   495
      Left            =   2040
      TabIndex        =   4
      Top             =   3000
      Width           =   1695
   End
   Begin VB.CommandButton btnAddDeclaration 
      Caption         =   "Add Declaration"
      Enabled         =   0   'False
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   3000
      Width           =   1695
   End
   Begin VB.CommandButton btnDisband 
      Caption         =   "Disband"
      Enabled         =   0   'False
      Height          =   495
      Left            =   120
      TabIndex        =   9
      Top             =   3600
      Width           =   1695
   End
   Begin VB.CommandButton btnRank 
      Caption         =   "Lord"
      Enabled         =   0   'False
      Height          =   495
      Index           =   2
      Left            =   5520
      TabIndex        =   7
      Top             =   3000
      Width           =   735
   End
   Begin VB.CommandButton btnRank 
      Caption         =   "Member"
      Enabled         =   0   'False
      Height          =   495
      Index           =   1
      Left            =   4680
      TabIndex        =   6
      Top             =   3000
      Width           =   735
   End
   Begin VB.CommandButton btnRank 
      Caption         =   "Initiate"
      Enabled         =   0   'False
      Height          =   495
      Index           =   0
      Left            =   3840
      TabIndex        =   5
      Top             =   3000
      Width           =   735
   End
   Begin VB.CommandButton btnRemoveMember 
      Caption         =   "Remove -->"
      Enabled         =   0   'False
      Height          =   495
      Left            =   2520
      TabIndex        =   0
      Top             =   960
      Width           =   1215
   End
   Begin VB.ListBox lstDeclarations 
      Height          =   1230
      Left            =   120
      TabIndex        =   1
      Top             =   1560
      Width           =   3615
   End
   Begin VB.CommandButton btnOk 
      Caption         =   "Close"
      Height          =   495
      Left            =   2040
      TabIndex        =   10
      Top             =   3600
      Width           =   5055
   End
   Begin VB.ListBox lstMembers 
      Height          =   2400
      Left            =   3840
      TabIndex        =   2
      Top             =   540
      Width           =   3255
   End
   Begin VB.Label lblHall 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   840
      TabIndex        =   16
      Top             =   480
      Width           =   1575
   End
   Begin VB.Label Label4 
      Caption         =   "Hall:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   480
      Width           =   615
   End
   Begin VB.Label lblName 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   840
      TabIndex        =   14
      Top             =   120
      Width           =   2535
   End
   Begin VB.Label Label3 
      Caption         =   "Name:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "Members:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3840
      TabIndex        =   12
      Top             =   240
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Declarations:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   1200
      Width           =   1575
   End
End
Attribute VB_Name = "frmGuild"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Sub UpdateForm()
    SendSocket Chr$(39) + Chr$(TempVar1)
    btnOk.Enabled = False
    btnRank(0).Enabled = False
    btnRank(1).Enabled = False
    btnRank(2).Enabled = False
    btnRank(3).Enabled = False
    btnRemoveMember.Enabled = False
    btnAddDeclaration.Enabled = False
    btnRemoveDeclaration.Enabled = False
    btnDisband.Enabled = False
End Sub


Private Sub btnAddDeclaration_Click()
    frmDeclaration.Show 1
    If TempVar3 > 0 Then
        SendSocket Chr$(37) + Chr$(TempVar3) + Chr$(TempVar2)
        UpdateForm
    End If
End Sub

Private Sub btnDisband_Click()
    If MsgBox("Are you sure you wish to disband your guild?  Your guild will be delete, and you will not get refunded for any of the guild fees nor will you get the money in your guild's bank account.  Continue?", vbYesNo + vbQuestion, TitleString) = vbYes Then
        SendSocket Chr$(42)
        Unload Me
    End If
End Sub

Private Sub btnMoveOut_Click()
    If MsgBox("Are you sure you wish to leave your guild hall?", vbYesNo + vbQuestion, TitleString) = vbYes Then
        SendSocket Chr$(44)
        UpdateForm
    End If
End Sub

Private Sub btnOk_Click()
    Unload Me
End Sub


Private Sub btnRank_Click(Index As Integer)
    SendSocket Chr$(36) + Chr$(lstMembers.ItemData(lstMembers.ListIndex)) + Chr$(Index)
    UpdateForm
End Sub

Private Sub btnRemoveDeclaration_Click()
    SendSocket Chr$(38) + Chr$(lstDeclarations.ItemData(lstDeclarations.ListIndex))
    UpdateForm
End Sub
Private Sub btnRemoveMember_Click()
    Dim St As String
    If lstMembers.ListCount <= 3 Then
        St = " Because there are only " + CStr(lstMembers.ListCount) + " members in your guild, removing this member will cause your guild to be deleted.  Are you SURE you wish to continue?"
    End If
    If MsgBox("Are you sure you wish to kick " + Chr$(34) + lstMembers.List(lstMembers.ListIndex) + Chr$(34) + " out of the guild?" + St, vbYesNo + vbQuestion, TitleString) = vbYes Then
        SendSocket Chr$(35) + Chr$(lstMembers.ItemData(lstMembers.ListIndex))
        UpdateForm
    End If
End Sub
Private Sub Form_Load()
    frmGuild_Loaded = True
End Sub


Private Sub Form_LostFocus()
    Unload Me
End Sub


Private Sub Form_Unload(Cancel As Integer)
    frmGuild_Loaded = False
End Sub


Private Sub lstDeclarations_Click()
    If TempVar1 = Character.Guild And Character.GuildRank >= 2 Then
        btnRemoveDeclaration.Enabled = True
    End If
End Sub

Private Sub lstMembers_Click()
    If TempVar1 = Character.Guild And Character.GuildRank >= 2 Then
        btnRemoveMember.Enabled = True
        btnRank(0).Enabled = True
        btnRank(1).Enabled = True
        btnRank(2).Enabled = True
        If Character.GuildRank = 3 Then
            btnRank(3).Enabled = True
        End If
    End If
End Sub


