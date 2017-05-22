VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "The Odyssey Online Classic [Login]"
   ClientHeight    =   2115
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   ControlBox      =   0   'False
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2115
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkSavePassword 
      Caption         =   "Remember Password"
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
      Left            =   1560
      TabIndex        =   6
      Top             =   1080
      Width           =   2175
   End
   Begin VB.CommandButton btnCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   495
      Left            =   1560
      TabIndex        =   2
      Top             =   1560
      Width           =   1455
   End
   Begin VB.CommandButton btnOk 
      Caption         =   "Login"
      Enabled         =   0   'False
      Height          =   495
      Left            =   3120
      TabIndex        =   3
      Top             =   1560
      Width           =   1455
   End
   Begin VB.TextBox txtPass 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1560
      MaxLength       =   15
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   600
      Width           =   3015
   End
   Begin VB.TextBox txtUser 
      Height          =   375
      Left            =   1560
      MaxLength       =   15
      TabIndex        =   0
      Top             =   120
      Width           =   3015
   End
   Begin VB.Label Label2 
      Caption         =   "Password:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "User:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub btnCancel_Click()
    Unload Me
    frmMenu.Show
End Sub


Private Sub btnOk_Click()
    User = txtUser
    Pass = txtPass
    
    WriteString "Login", "User", User
    If chkSavePassword = 1 Then
        WriteString "Login", "Password", Pass
    Else
        DelItem "Login", "Password"
    End If
    WriteString "Login", "SavePassword", chkSavePassword
    
    NewAccount = False
    
    ConnectClient
    
    Me.Hide
End Sub

Private Sub Form_Load()
    frmLogin_Loaded = True
    
    If ReadInt("Login", "SavePassword") > 0 Then
        chkSavePassword = 1
        txtPass = ReadString("Login", "Password")
    Else
        chkSavePassword = 0
    End If
    
    txtUser = ReadString("Login", "User")
    
    Me.Show
    
    If txtUser = "" Then
        txtUser.SetFocus
    Else
        txtPass.SetFocus
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmLogin_Loaded = False
End Sub


Private Sub txtPass_Change()
    If txtUser <> "" And txtPass <> "" Then
        If btnOk.Enabled = False Then
            btnOk.Enabled = True
        End If
    Else
        If btnOk.Enabled = True Then
            btnOk.Enabled = False
        End If
    End If
End Sub

Private Sub txtPass_GotFocus()
    txtPass.SelStart = 0
    txtPass.SelLength = Len(txtPass)
End Sub


Private Sub txtPass_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And btnOk.Enabled = True Then btnOk_Click
End Sub

Private Sub txtUser_Change()
    If txtUser <> "" And txtPass <> "" Then
        If btnOk.Enabled = False Then
            btnOk.Enabled = True
        End If
    Else
        If btnOk.Enabled = True Then
            btnOk.Enabled = False
        End If
    End If
End Sub
Private Sub txtUser_GotFocus()
    txtUser.SelStart = 0
    txtUser.SelLength = Len(txtUser)
End Sub
