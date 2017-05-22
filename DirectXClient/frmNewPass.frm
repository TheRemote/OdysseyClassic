VERSION 5.00
Begin VB.Form frmNewPass 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "The Odyssey Online Classic [Change Password]"
   ClientHeight    =   1785
   ClientLeft      =   75
   ClientTop       =   300
   ClientWidth     =   5490
   ControlBox      =   0   'False
   Icon            =   "frmNewPass.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   119
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   366
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtPass1 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   2400
      MaxLength       =   15
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   120
      Width           =   3015
   End
   Begin VB.TextBox txtPass2 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   2400
      MaxLength       =   15
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   600
      Width           =   3015
   End
   Begin VB.CommandButton btnOk 
      Caption         =   "Ok"
      Enabled         =   0   'False
      Height          =   495
      Left            =   3960
      TabIndex        =   3
      Top             =   1200
      Width           =   1455
   End
   Begin VB.CommandButton btnCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   495
      Left            =   2400
      TabIndex        =   2
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "New Password:"
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
      Top             =   120
      Width           =   2175
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Re-Type Password:"
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
      Top             =   600
      Width           =   2175
   End
End
Attribute VB_Name = "frmNewPass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub btnCancel_Click()
    Pass = vbNullString
    Unload Me
End Sub

Private Sub btnOk_Click()
    If Len(txtPass1) >= 3 Then
        If UCase$(txtPass1) = UCase$(txtPass2) Then
            SendSocket Chr$(3) + UCase$(txtPass1)
            PrintChat "Password changed", 15
            Unload Me
        Else
            MsgBox "Your two passwords do not match, please re-enter!", vbOKOnly, TitleString
        End If
    Else
        MsgBox "Your password must be atleast 3 characters long!", vbOKOnly + vbExclamation, TitleString
    End If
End Sub
Private Sub Form_Load()
    frmNewPass_Loaded = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmNewPass_Loaded = False
End Sub


Private Sub txtPass1_Change()
    If txtPass1 <> "" And txtPass2 <> "" Then
        If btnOk.Enabled = False Then
            btnOk.Enabled = True
        End If
    Else
        If btnOk.Enabled = True Then
            btnOk.Enabled = False
        End If
    End If
End Sub

Private Sub txtPass1_GotFocus()
    txtPass1.SelStart = 0
    txtPass1.SelLength = Len(txtPass1)
End Sub


Private Sub txtPass2_Change()
    If txtPass1 <> "" And txtPass2 <> "" Then
        If btnOk.Enabled = False Then
            btnOk.Enabled = True
        End If
    Else
        If btnOk.Enabled = True Then
            btnOk.Enabled = False
        End If
    End If
End Sub


Private Sub txtPass2_GotFocus()
    txtPass2.SelStart = 0
    txtPass2.SelLength = Len(txtPass2)
End Sub


