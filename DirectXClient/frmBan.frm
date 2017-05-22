VERSION 5.00
Begin VB.Form frmBan 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "The Odyssey Online Classic [Ban]"
   ClientHeight    =   4290
   ClientLeft      =   150
   ClientTop       =   345
   ClientWidth     =   6900
   ControlBox      =   0   'False
   Icon            =   "frmBan.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4290
   ScaleWidth      =   6900
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnOk 
      Caption         =   "Ok"
      Height          =   495
      Left            =   4320
      TabIndex        =   17
      Top             =   3720
      Width           =   1455
   End
   Begin VB.CommandButton btnCancel 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   2760
      TabIndex        =   16
      Top             =   3720
      Width           =   1455
   End
   Begin VB.CommandButton btnClear 
      Caption         =   "Clear"
      Height          =   495
      Left            =   1200
      TabIndex        =   15
      Top             =   3720
      Width           =   1455
   End
   Begin VB.TextBox txtName 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1320
      MaxLength       =   15
      TabIndex        =   3
      Top             =   480
      Width           =   2055
   End
   Begin VB.TextBox txtReason 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1320
      MaxLength       =   255
      MultiLine       =   -1  'True
      TabIndex        =   11
      Top             =   2400
      Width           =   5415
   End
   Begin VB.TextBox txtBanner 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1320
      MaxLength       =   15
      TabIndex        =   5
      Top             =   960
      Width           =   2055
   End
   Begin VB.HScrollBar sclUnban 
      Height          =   255
      Left            =   1200
      Max             =   255
      TabIndex        =   13
      Top             =   3360
      Width           =   4815
   End
   Begin VB.TextBox txtComputerID 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1320
      MaxLength       =   255
      TabIndex        =   9
      Top             =   1920
      Width           =   5415
   End
   Begin VB.TextBox txtIPAddress 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1320
      MaxLength       =   15
      TabIndex        =   7
      Top             =   1440
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "Name:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   480
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Reason:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   10
      Top             =   2520
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "Banner:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   4
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Days Left:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   12
      Top             =   3360
      Width           =   1095
   End
   Begin VB.Label lblUnban 
      Alignment       =   2  'Center
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6240
      TabIndex        =   14
      Top             =   3360
      Width           =   495
   End
   Begin VB.Label Label5 
      Caption         =   "Number:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   855
   End
   Begin VB.Label lblNumber 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1320
      TabIndex        =   1
      Top             =   0
      Width           =   2055
   End
   Begin VB.Label Label6 
      Caption         =   "ComputerID:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   8
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label Label7 
      Caption         =   "IP Address:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   6
      Top             =   1440
      Width           =   1215
   End
End
Attribute VB_Name = "frmBan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub btnClear_Click()
    SendSocket Chr$(18) + Chr$(11) + Chr$(lblNumber)
    Unload Me
End Sub

Private Sub btnOk_Click()
    SendSocket Chr$(58) + Chr$(lblNumber) + Chr$(sclUnban) + txtName + vbNullChar + txtBanner + vbNullChar + txtReason + vbNullChar + txtComputerID + vbNullChar + txtIPAddress
    Unload Me
End Sub

Private Sub Form_Load()
    frmBan_Loaded = True
End Sub

Private Sub Form_LostFocus()
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmBan_Loaded = False
End Sub

Private Sub sclUnban_Change()
    lblUnban = sclUnban
End Sub

Private Sub sclUnban_Scroll()
    sclUnban_Change
End Sub

Private Sub txtReason_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Or KeyAscii = 10 Then
        KeyAscii = 0
        Beep
    End If
End Sub


