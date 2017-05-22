VERSION 5.00
Begin VB.Form frmNewGuild 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "The Odyssey Online Classic [New Guild]"
   ClientHeight    =   2265
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4620
   ControlBox      =   0   'False
   Icon            =   "frmNewGuild.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2265
   ScaleWidth      =   4620
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtName 
      Height          =   375
      Left            =   1440
      MaxLength       =   15
      TabIndex        =   0
      Top             =   120
      Width           =   3135
   End
   Begin VB.CommandButton btnOk 
      Caption         =   "Ok"
      Enabled         =   0   'False
      Height          =   495
      Left            =   3120
      TabIndex        =   2
      Top             =   1680
      Width           =   1455
   End
   Begin VB.CommandButton btnCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   495
      Left            =   1560
      TabIndex        =   1
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   $"frmNewGuild.frx":000C
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   4455
   End
   Begin VB.Label Label1 
      Caption         =   "Guild Name:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmNewGuild"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub btnOk_Click()
    txtName = Trim$(txtName)
    If Len(txtName) >= 3 Then
        SendSocket Chr$(33) + txtName
        Unload Me
    Else
        MsgBox "Guild name must be atleast 3 characters long!", vbOKOnly + vbExclamation, TitleString
    End If
End Sub

Private Sub Form_Load()
    frmNewGuild_Loaded = True
End Sub

Private Sub Form_LostFocus()
    Unload Me
End Sub


Private Sub Form_Unload(Cancel As Integer)
    frmNewGuild_Loaded = False
End Sub

Private Sub txtName_Change()
    If txtName <> "" Then
        If btnOk.Enabled = False Then
            btnOk.Enabled = True
        End If
    Else
        If btnOk.Enabled = True Then
            btnOk.Enabled = False
        End If
    End If
End Sub


Private Sub txtName_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Or (KeyAscii >= 48 And KeyAscii <= 57) Or (KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii >= 97 And KeyAscii <= 122) Or KeyAscii = 95 Or KeyAscii = 32 Then
        'Valid Key
    Else
        KeyAscii = 0
        Beep
    End If
End Sub

