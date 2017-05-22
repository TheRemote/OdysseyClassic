VERSION 5.00
Begin VB.Form frmLogin 
   Appearance      =   0  'Flat
   BackColor       =   &H0061514B&
   BorderStyle     =   0  'None
   Caption         =   "The Odyssey Online Classic [Login]"
   ClientHeight    =   3645
   ClientLeft      =   0
   ClientTop       =   -15
   ClientWidth     =   4500
   ControlBox      =   0   'False
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3645
   ScaleWidth      =   4500
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtUser 
      Appearance      =   0  'Flat
      BackColor       =   &H0044342E&
      ForeColor       =   &H009AADC2&
      Height          =   285
      Left            =   1320
      MaxLength       =   15
      TabIndex        =   2
      Top             =   840
      Width           =   3015
   End
   Begin VB.TextBox txtPass 
      Appearance      =   0  'Flat
      BackColor       =   &H0044342E&
      ForeColor       =   &H009AADC2&
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1320
      MaxLength       =   15
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   1200
      Width           =   3015
   End
   Begin VB.Label lblPasswordWarning 
      BackColor       =   &H0044342E&
      BackStyle       =   0  'Transparent
      Caption         =   $"frmLogin.frx":1CFA
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H009AADC2&
      Height          =   1455
      Left            =   120
      TabIndex        =   7
      Top             =   1560
      Width           =   4215
   End
   Begin VB.Label lblCurrentServer 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H009AADC2&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   4215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
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
      ForeColor       =   &H009AADC2&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   615
   End
   Begin VB.Label Label2 
      BackColor       =   &H0044342E&
      BackStyle       =   0  'Transparent
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
      ForeColor       =   &H009AADC2&
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label btnOk 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0044342E&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Login"
      Enabled         =   0   'False
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
      Height          =   420
      Left            =   2280
      TabIndex        =   6
      Top             =   3120
      Width           =   2055
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
      Height          =   420
      Left            =   120
      TabIndex        =   5
      Top             =   3120
      Width           =   2100
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
    RequestedMap = False

    User = txtUser
    Pass = txtPass

    WriteString "Login", "User", User
    DelItem "Login", "Password"

    NewAccount = False

    ConnectClient

    Me.Hide
End Sub

Private Sub Form_Load()
    frmLogin_Loaded = True

    Me.Picture = frmMenu.Picture

    txtUser = ReadString$("Login", "User")

    Me.Show

    If txtUser = vbNullString Then
        txtUser.SetFocus
    Else
        txtPass.SetFocus
    End If

    lblCurrentServer = ServerDescription
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
