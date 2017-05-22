VERSION 5.00
Begin VB.Form frmMenu 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   ClientHeight    =   2580
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4500
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmMenu.frx":0000
   ScaleHeight     =   2580
   ScaleWidth      =   4500
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label lblMenu 
      Alignment       =   2  'Center
      BackColor       =   &H0080C0FF&
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
      ForeColor       =   &H00000000&
      Height          =   495
      Index           =   4
      Left            =   120
      TabIndex        =   4
      Top             =   2000
      Width           =   4260
   End
   Begin VB.Label lblMenu 
      Alignment       =   2  'Center
      BackColor       =   &H0080C0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Credits"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Index           =   3
      Left            =   120
      TabIndex        =   3
      Top             =   1420
      Width           =   2100
   End
   Begin VB.Label lblMenu 
      Alignment       =   2  'Center
      BackColor       =   &H0080C0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Options"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Index           =   2
      Left            =   2280
      TabIndex        =   2
      Top             =   1420
      Width           =   2100
   End
   Begin VB.Label lblMenu 
      Alignment       =   2  'Center
      BackColor       =   &H0080C0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "New Account"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Index           =   1
      Left            =   2280
      TabIndex        =   1
      Top             =   840
      Width           =   2100
   End
   Begin VB.Label lblMenu 
      Alignment       =   2  'Center
      BackColor       =   &H0080C0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Login"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   2100
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    gHW = Me.hWnd
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 And Me.WindowState = 0 Then
       Dim ReturnVal As Long
       ReleaseCapture
       ReturnVal = SendMessage(Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0)
    End If
End Sub
Private Sub lblMenu_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblMenu(Index).BackColor = QBColor(7)
End Sub


Private Sub lblMenu_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblMenu(Index).BackColor = &H80C0FF
    If X >= 0 And X <= lblMenu(Index).Width And Y >= 0 And Y <= lblMenu(Index).Height Then
        Select Case Index
            Case 0 'Login
                frmLogin.Show
                Me.Hide
            Case 1 'New Account
                frmAccount.Show
                Me.Hide
            Case 2 'Options
                frmOptions.Show
                Me.Hide
            Case 3 'Credits
                frmCredits.Show
                Me.Hide
            Case 4 'Quit
                blnEnd = True
        End Select
    End If
End Sub


