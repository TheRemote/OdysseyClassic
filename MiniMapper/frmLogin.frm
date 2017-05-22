VERSION 5.00
Begin VB.Form frmLogin 
   BackColor       =   &H0061514B&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Odyssey Final Minimapper"
   ClientHeight    =   2625
   ClientLeft      =   120
   ClientTop       =   360
   ClientWidth     =   4485
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2625
   ScaleWidth      =   4485
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtServerPort 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   1440
      TabIndex        =   10
      Text            =   "5750"
      Top             =   1200
      Width           =   615
   End
   Begin VB.TextBox txtServerIP 
      Appearance      =   0  'Flat
      BackColor       =   &H0061514B&
      ForeColor       =   &H009AADC2&
      Height          =   285
      Left            =   1320
      MaxLength       =   50
      TabIndex        =   7
      Text            =   "72.167.33.97"
      Top             =   840
      Width           =   3015
   End
   Begin VB.TextBox txtVersion 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   1440
      TabIndex        =   6
      Text            =   "197"
      Top             =   1560
      Width           =   615
   End
   Begin VB.Timer TimeDeinitialize 
      Left            =   0
      Top             =   0
   End
   Begin VB.TextBox txtUser 
      Appearance      =   0  'Flat
      BackColor       =   &H0061514B&
      ForeColor       =   &H009AADC2&
      Height          =   285
      Left            =   1320
      MaxLength       =   15
      TabIndex        =   0
      Top             =   120
      Width           =   3015
   End
   Begin VB.TextBox txtPass 
      Appearance      =   0  'Flat
      BackColor       =   &H0061514B&
      ForeColor       =   &H009AADC2&
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1320
      MaxLength       =   15
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   480
      Width           =   3015
   End
   Begin VB.Label Label3 
      BackColor       =   &H0044342E&
      BackStyle       =   0  'Transparent
      Caption         =   "Server Port:"
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
      TabIndex        =   11
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label lblVersion 
      BackColor       =   &H0044342E&
      BackStyle       =   0  'Transparent
      Caption         =   "Version:"
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
      TabIndex        =   9
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label lblServerIP 
      BackColor       =   &H0044342E&
      BackStyle       =   0  'Transparent
      Caption         =   "Server IP:"
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
      TabIndex        =   8
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H0044342E&
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
      TabIndex        =   5
      Top             =   120
      Width           =   1215
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
      TabIndex        =   4
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label btnOk 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0044342E&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Map It"
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
      Height          =   495
      Left            =   2280
      TabIndex        =   3
      Top             =   2040
      Width           =   2100
   End
   Begin VB.Label btnCancel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0044342E&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Quit"
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
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   2040
      Width           =   2100
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnCancel_Click()
TimeDeinitialize.Interval = 50
End Sub

Private Sub btnOk_Click()
    User = txtUser
    Pass = txtPass
    
    Dim St As String
    CMap = 1
    
    ServerIP = txtServerIP.Text
    
    PacketOrder = 0
    ServerPacketOrder = 0
    
    SocketData = ""
    ClientSocket = ConnectSock(ServerIP, txtServerPort, St, gHW, True)
    
    btnOk.Enabled = False
End Sub

Private Sub Form_Load()
    gHW = Me.hwnd
    'SaveMap
End Sub

Private Sub Form_Unload(Cancel As Integer)
    TimeDeinitialize.Interval = 50
End Sub

Private Sub TimeDeinitialize_Timer()
    TimeDeinitialize.Interval = 0
    DeInitialize
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
