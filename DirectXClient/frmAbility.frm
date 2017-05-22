VERSION 5.00
Begin VB.Form frmAbility 
   Caption         =   "Edit Ability"
   ClientHeight    =   6510
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4950
   ControlBox      =   0   'False
   Icon            =   "frmAbility.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6510
   ScaleWidth      =   4950
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtDescription 
      Height          =   1335
      Left            =   1320
      MaxLength       =   100
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   1560
      Width           =   2895
   End
   Begin VB.TextBox txtScript 
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
      TabIndex        =   1
      Top             =   1080
      Width           =   2895
   End
   Begin VB.OptionButton optClass 
      Caption         =   "Cleric"
      Height          =   255
      Index           =   4
      Left            =   1320
      TabIndex        =   16
      Top             =   5040
      Width           =   1455
   End
   Begin VB.OptionButton optClass 
      Caption         =   "Thief"
      Height          =   255
      Index           =   3
      Left            =   1320
      TabIndex        =   15
      Top             =   4680
      Width           =   1455
   End
   Begin VB.OptionButton optClass 
      Caption         =   "Paladin"
      Height          =   255
      Index           =   2
      Left            =   1320
      TabIndex        =   14
      Top             =   4320
      Width           =   1455
   End
   Begin VB.OptionButton optClass 
      Caption         =   "Mage"
      Height          =   255
      Index           =   1
      Left            =   1320
      TabIndex        =   13
      Top             =   3960
      Width           =   1455
   End
   Begin VB.OptionButton optClass 
      Caption         =   "Knight"
      Height          =   255
      Index           =   0
      Left            =   1320
      TabIndex        =   5
      Top             =   3600
      Width           =   1455
   End
   Begin VB.CommandButton btnOk 
      Caption         =   "Ok"
      Height          =   495
      Left            =   2520
      TabIndex        =   6
      Top             =   5520
      Width           =   1455
   End
   Begin VB.CommandButton btnCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   495
      Left            =   960
      TabIndex        =   3
      Top             =   5520
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
      TabIndex        =   0
      Top             =   600
      Width           =   2895
   End
   Begin VB.HScrollBar sclLevel 
      Height          =   255
      Left            =   1320
      Max             =   255
      Min             =   1
      TabIndex        =   4
      Top             =   3120
      Value           =   1
      Width           =   2895
   End
   Begin VB.Label cptScript 
      Caption         =   "Description:"
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
      Index           =   1
      Left            =   120
      TabIndex        =   18
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label cptScript 
      Caption         =   "Script:"
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
      Index           =   0
      Left            =   120
      TabIndex        =   17
      Top             =   1080
      Width           =   855
   End
   Begin VB.Label cptName 
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
      TabIndex        =   12
      Top             =   600
      Width           =   855
   End
   Begin VB.Label cptLevel 
      Caption         =   "Level:"
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
      TabIndex        =   11
      Top             =   3120
      Width           =   855
   End
   Begin VB.Label cptNumber 
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
      TabIndex        =   10
      Top             =   120
      Width           =   855
   End
   Begin VB.Label lblNumber 
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
      Height          =   375
      Left            =   1560
      TabIndex        =   9
      Top             =   120
      Width           =   2655
   End
   Begin VB.Label cptClass 
      Caption         =   "Class:"
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
      Top             =   3480
      Width           =   855
   End
   Begin VB.Label lblLevel 
      Alignment       =   2  'Center
      Caption         =   "1"
      Height          =   255
      Left            =   4320
      TabIndex        =   7
      Top             =   3120
      Width           =   495
   End
End
Attribute VB_Name = "frmAbility"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnCancel_Click()
    Me.Hide
End Sub

Private Sub btnOk_Click()
    Dim a As Long, Class As Byte
    For a = 0 To 4
        If optClass(a).value = True Then Class = a + 1
    Next a
    If Class > 0 And Len(txtName) > 0 And Len(txtScript) > 0 Then
        SendSocket Chr$(83) + Chr$(lblNumber) + Chr$(lblLevel) + Chr$(Class) + txtName + Chr$(0) + txtScript
        SendSocket Chr$(85) + Chr$(lblNumber) + txtDescription
        Me.Hide
    Else
        MsgBox "Invalid parameters!"
    End If
End Sub

Private Sub Form_Load()
    frmAbility_Loaded = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmAbility_Loaded = False
End Sub

Private Sub sclLevel_Change()
    lblLevel = sclLevel
End Sub
