VERSION 5.00
Begin VB.Form frmWait 
   BackColor       =   &H00004080&
   BorderStyle     =   0  'None
   Caption         =   "The Odyssey Online Classic"
   ClientHeight    =   585
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4500
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmWait.frx":0000
   ScaleHeight     =   39
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   300
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnCancel 
      BackColor       =   &H00808000&
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   330
      Left            =   3750
      TabIndex        =   1
      Top             =   135
      Visible         =   0   'False
      Width           =   660
   End
   Begin VB.Timer ConnectTimer 
      Enabled         =   0   'False
      Interval        =   15000
      Left            =   0
      Top             =   0
   End
   Begin VB.Label lblStatus 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   3495
   End
End
Attribute VB_Name = "frmWait"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnCancel_Click()
    CloseClientSocket 1
End Sub

Private Sub ConnectTimer_Timer()
    ConnectTimer.Enabled = False
    ConnectClient
End Sub

Private Sub Form_Load()
    frmWait_Loaded = True
End Sub


Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 And Me.WindowState = 0 Then
       Dim ReturnVal As Long
       ReleaseCapture
       ReturnVal = SendMessage(Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0)
    End If
End Sub


Private Sub Form_Unload(Cancel As Integer)
    frmWait_Loaded = False
End Sub


Private Sub lblStatus_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 And Me.WindowState = 0 Then
       Dim ReturnVal As Long
       ReleaseCapture
       ReturnVal = SendMessage(Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0)
    End If
End Sub


