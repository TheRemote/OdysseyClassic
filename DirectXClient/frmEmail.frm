VERSION 5.00
Begin VB.Form frmEmail 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Enter an e-mail address"
   ClientHeight    =   2910
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   5550
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2910
   ScaleWidth      =   5550
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   495
      Left            =   2400
      TabIndex        =   3
      Top             =   2280
      Width           =   1455
   End
   Begin VB.CommandButton btnOk 
      Caption         =   "Ok"
      Enabled         =   0   'False
      Height          =   495
      Left            =   3960
      TabIndex        =   2
      Top             =   2280
      Width           =   1455
   End
   Begin VB.TextBox txtEmail2 
      Height          =   375
      HideSelection   =   0   'False
      IMEMode         =   3  'DISABLE
      Left            =   2400
      MaxLength       =   150
      TabIndex        =   1
      Top             =   600
      Width           =   3015
   End
   Begin VB.TextBox txtEmail1 
      Height          =   375
      HideSelection   =   0   'False
      IMEMode         =   3  'DISABLE
      Left            =   2400
      MaxLength       =   150
      TabIndex        =   0
      Top             =   120
      Width           =   3015
   End
   Begin VB.Label Label1 
      Caption         =   $"frmEmail.frx":0000
      Height          =   1095
      Left            =   120
      TabIndex        =   6
      Top             =   1080
      Width           =   5295
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Re-Type E-mail:"
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
      Width           =   2175
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "New E-mail:"
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
      Width           =   2175
   End
End
Attribute VB_Name = "frmEmail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub btnOk_Click()
    If Len(txtEmail1) >= 6 Then
        If UCase$(txtEmail1) = UCase$(txtEmail2) Then
            SendSocket Chr$(103) + txtEmail1
            Unload Me
        Else
            MsgBox "Your two email addresses do not match, please re-enter!", vbOKOnly, TitleString
        End If
    Else
        MsgBox "Your e-mail must be atleast 6 characters long!", vbOKOnly + vbExclamation, TitleString
    End If
End Sub

Private Sub Form_Load()
    frmEmail_Loaded = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmEmail_Loaded = False
End Sub

Private Sub txtEmail1_Change()
    If txtEmail1 <> "" And txtEmail2 <> "" Then
        If btnOk.Enabled = False Then
            btnOk.Enabled = True
        End If
    Else
        If btnOk.Enabled = True Then
            btnOk.Enabled = False
        End If
    End If
End Sub

Private Sub txtEmail1_GotFocus()
    txtEmail1.SelStart = 0
    txtEmail1.SelLength = Len(txtEmail1)
End Sub


Private Sub txtEmail2_Change()
    If txtEmail1 <> "" And txtEmail2 <> "" Then
        If btnOk.Enabled = False Then
            btnOk.Enabled = True
        End If
    Else
        If btnOk.Enabled = True Then
            btnOk.Enabled = False
        End If
    End If
End Sub


Private Sub txtEmail2_GotFocus()
    txtEmail2.SelStart = 0
    txtEmail2.SelLength = Len(txtEmail2)
End Sub



