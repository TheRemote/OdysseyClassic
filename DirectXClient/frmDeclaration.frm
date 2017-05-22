VERSION 5.00
Begin VB.Form frmDeclaration 
   BackColor       =   &H0061514B&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "The Odyssey Online Classic [Delcaration]"
   ClientHeight    =   1635
   ClientLeft      =   90
   ClientTop       =   345
   ClientWidth     =   4305
   ControlBox      =   0   'False
   Icon            =   "frmDeclaration.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1635
   ScaleWidth      =   4305
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cmbGuild 
      BackColor       =   &H0044342E&
      ForeColor       =   &H009AADC2&
      Height          =   315
      Left            =   1560
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   600
      Width           =   2535
   End
   Begin VB.ComboBox cmbType 
      BackColor       =   &H0044342E&
      ForeColor       =   &H009AADC2&
      Height          =   315
      Left            =   1560
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   120
      Width           =   2535
   End
   Begin VB.Label btnOk 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0044342E&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Declare"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H009AADC2&
      Height          =   375
      Left            =   2400
      TabIndex        =   5
      Top             =   1080
      Width           =   1740
   End
   Begin VB.Label btnCancel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0044342E&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H009AADC2&
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   1860
   End
   Begin VB.Label Label2 
      BackColor       =   &H0061514B&
      Caption         =   "Declaration Guild:"
      ForeColor       =   &H009AADC2&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H0061514B&
      Caption         =   "Declaration Type:"
      ForeColor       =   &H009AADC2&
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

    For A = 1 To MaxGuilds
        If Guild(A).name <> "" And A <> Character.Guild Then
            cmbGuild.AddItem Guild(A).name
            cmbGuild.ItemData(cmbGuild.ListCount - 1) = A
        End If
    Next A

    If cmbGuild.ListCount > 0 Then
        cmbGuild.ListIndex = 0
    Else
        btnOk.Enabled = False
    End If
End Sub
