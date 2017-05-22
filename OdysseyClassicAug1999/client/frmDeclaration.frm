VERSION 5.00
Begin VB.Form frmDeclaration 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "The Odyssey Online Classic [Delcaration]"
   ClientHeight    =   1635
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4305
   ControlBox      =   0   'False
   Icon            =   "frmDeclaration.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1635
   ScaleWidth      =   4305
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   495
      Left            =   1560
      TabIndex        =   5
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton btnOk 
      Caption         =   "Ok"
      Height          =   495
      Left            =   2880
      TabIndex        =   4
      Top             =   1080
      Width           =   1215
   End
   Begin VB.ComboBox cmbGuild 
      Height          =   315
      Left            =   1560
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   600
      Width           =   2535
   End
   Begin VB.ComboBox cmbType 
      Height          =   315
      Left            =   1560
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   120
      Width           =   2535
   End
   Begin VB.Label Label2 
      Caption         =   "Declaration Guild:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Declaration Type:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "frmDeclaration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub btnCancel_Click()
    TempVar2 = 0
    TempVar3 = 0
    Unload Me
End Sub

Private Sub btnOk_Click()
    TempVar2 = cmbType.ListIndex
    TempVar3 = cmbGuild.ItemData(cmbGuild.ListIndex)
    Unload Me
End Sub

Private Sub Form_Load()
    Dim A As Long
    
    cmbType.AddItem "Alliance"
    cmbType.AddItem "War"
    cmbType.ListIndex = 0
    
    For A = 1 To 255
        If Guild(A).Name <> "" And A <> Character.Guild Then
            cmbGuild.AddItem Guild(A).Name
            cmbGuild.ItemData(cmbGuild.ListCount - 1) = A
        End If
    Next A
    
    If cmbGuild.ListCount > 0 Then
        cmbGuild.ListIndex = 0
    Else
        btnOk.Enabled = False
    End If
End Sub
