VERSION 5.00
Begin VB.Form frmBan 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "The Odyssey Online Classic [Ban]"
   ClientHeight    =   3225
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5955
   ControlBox      =   0   'False
   Icon            =   "frmBan.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3225
   ScaleWidth      =   5955
   StartUpPosition =   2  'CenterScreen
   Begin VB.HScrollBar sclUnban 
      Height          =   255
      Left            =   1440
      Max             =   100
      TabIndex        =   3
      Top             =   2280
      Width           =   3855
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
      Left            =   1440
      Locked          =   -1  'True
      MaxLength       =   15
      TabIndex        =   1
      Top             =   960
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
      Left            =   1440
      MaxLength       =   255
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   1440
      Width           =   4575
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
      Left            =   1440
      MaxLength       =   15
      TabIndex        =   0
      Top             =   480
      Width           =   2055
   End
   Begin VB.CommandButton btnClear 
      Caption         =   "Clear"
      Height          =   495
      Left            =   1320
      TabIndex        =   4
      Top             =   2640
      Width           =   1455
   End
   Begin VB.CommandButton btnCancel 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   2880
      TabIndex        =   5
      Top             =   2640
      Width           =   1455
   End
   Begin VB.CommandButton btnOk 
      Caption         =   "Ok"
      Height          =   495
      Left            =   4440
      TabIndex        =   6
      Top             =   2640
      Width           =   1455
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
      Left            =   1440
      TabIndex        =   13
      Top             =   0
      Width           =   2055
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
      Left            =   120
      TabIndex        =   12
      Top             =   0
      Width           =   855
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
      Left            =   5400
      TabIndex        =   11
      Top             =   2280
      Width           =   495
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
      Left            =   120
      TabIndex        =   10
      Top             =   2280
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
      Left            =   120
      TabIndex        =   9
      Top             =   960
      Width           =   1215
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
      Left            =   120
      TabIndex        =   8
      Top             =   1440
      Width           =   1095
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
      Left            =   120
      TabIndex        =   7
      Top             =   480
      Width           =   855
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
    txtName = ""
    txtBanner = ""
    txtReason = ""
    sclUnban = 0
End Sub

Private Sub btnOk_Click()
    SendSocket Chr$(58) + Chr$(lblNumber) + Chr$(sclUnban) + txtName + Chr$(0) + txtBanner + Chr$(0) + txtReason
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


