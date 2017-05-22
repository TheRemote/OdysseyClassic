VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmChat 
   Caption         =   "Remote Admin - Chat"
   ClientHeight    =   4575
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6375
   Icon            =   "frmChat.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4575
   ScaleWidth      =   6375
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close Chat"
      Height          =   375
      Left            =   5040
      TabIndex        =   3
      Top             =   4080
      Width           =   1215
   End
   Begin MSWinsockLib.Winsock sockChat 
      Left            =   240
      Top             =   4080
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemotePort      =   6969
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "Send"
      Height          =   375
      Left            =   2520
      TabIndex        =   2
      Top             =   4080
      Width           =   1335
   End
   Begin VB.TextBox txtTyping 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Text            =   "txtTyping"
      Top             =   3600
      Width           =   6135
   End
   Begin VB.TextBox txtChat 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   120
      Width           =   6135
   End
End
Attribute VB_Name = "frmChat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
frmMain.lblStatus.Caption = " Status:  Closing Remote Chat"
frmMain.ForeColor = vbGreen
frmMain.MessageTimer.Enabled = True
frmMain.SockMain.SendData "|CLOSECHAT|"
Pause 10
frmMain.lblState(3).Caption = "Not Connected"
sockChat.Close
Unload Me
End Sub

Private Sub cmdSend_Click()
Dim MSG As String
MSG = "|CHAT|" & "Admin: " & txtTyping.Text & Chr(13) & Chr(10)
sockChat.SendData ENCRYPT(MSG, Len(MSG))
txtChat.Text = txtChat.Text & "Admin: " & txtTyping.Text & Chr(13) & Chr(10)
txtTyping.Text = ""
txtTyping.SetFocus
End Sub

Private Sub Form_Load()
txtTyping.Text = ""
Me.Show
sockChat.Close
sockChat.RemoteHost = frmMain.SockMain.RemoteHost
sockChat.Connect
End Sub

Private Sub sockChat_Close()
sockChat.Close
frmMain.lblState(3).Caption = "Not Connected"
Unload Me
End Sub

Private Sub sockChat_Connect()
Me.Show
frmMain.lblState(3).Caption = "Connected"
End Sub

Private Sub sockChat_DataArrival(ByVal bytesTotal As Long)
Dim Data As String
sockChat.GetData Data
Data = DECRYPT(Data, Len(Data))

If InStr(1, Data, "|CHAT|") <> 0 Then
    txtChat.Text = txtChat.Text & Mid(Data, 7, Len(Data))
    Exit Sub
End If

End Sub

Private Sub txtChat_Change()
txtChat.SelStart = Len(txtChat.Text)
End Sub

Private Sub txtTyping_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then cmdSend = True
End Sub
