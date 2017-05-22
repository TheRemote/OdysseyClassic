VERSION 5.00
Begin VB.Form frmHall 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "The Odyssey Online Classic [Guild Hall]"
   ClientHeight    =   3960
   ClientLeft      =   90
   ClientTop       =   345
   ClientWidth     =   5820
   ControlBox      =   0   'False
   Icon            =   "frmHall.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3960
   ScaleWidth      =   5820
   StartUpPosition =   2  'CenterScreen
   Begin VB.HScrollBar sclStartMap 
      Height          =   255
      LargeChange     =   25
      Left            =   1200
      Max             =   3000
      Min             =   1
      TabIndex        =   12
      Top             =   2040
      Value           =   1
      Width           =   3975
   End
   Begin VB.HScrollBar sclStartX 
      Height          =   255
      Left            =   1200
      Max             =   11
      TabIndex        =   11
      Top             =   2400
      Width           =   3975
   End
   Begin VB.HScrollBar sclStartY 
      Height          =   255
      Left            =   1200
      Max             =   11
      TabIndex        =   10
      Top             =   2760
      Width           =   3975
   End
   Begin VB.TextBox txtUpkeep 
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
      Left            =   1200
      MaxLength       =   9
      TabIndex        =   3
      Top             =   1560
      Width           =   1575
   End
   Begin VB.TextBox txtPrice 
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
      Left            =   1200
      MaxLength       =   9
      TabIndex        =   2
      Top             =   1080
      Width           =   1575
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
      Left            =   1200
      MaxLength       =   15
      TabIndex        =   1
      Top             =   600
      Width           =   2895
   End
   Begin VB.CommandButton btnOk 
      Caption         =   "Ok"
      Height          =   495
      Left            =   4320
      TabIndex        =   5
      Top             =   3360
      Width           =   1455
   End
   Begin VB.CommandButton btnCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   495
      Left            =   2760
      TabIndex        =   4
      Top             =   3360
      Width           =   1455
   End
   Begin VB.Label Label7 
      Caption         =   "Start Y:"
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
      TabIndex        =   18
      Top             =   2760
      Width           =   975
   End
   Begin VB.Label Label6 
      Caption         =   "Start X:"
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
      TabIndex        =   17
      Top             =   2400
      Width           =   975
   End
   Begin VB.Label Label5 
      Caption         =   "Start Map:"
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
      TabIndex        =   16
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label lblStartMap 
      Alignment       =   2  'Center
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5280
      TabIndex        =   15
      Top             =   2040
      Width           =   495
   End
   Begin VB.Label lblStartX 
      Alignment       =   2  'Center
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5280
      TabIndex        =   14
      Top             =   2400
      Width           =   495
   End
   Begin VB.Label lblStartY 
      Alignment       =   2  'Center
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5280
      TabIndex        =   13
      Top             =   2760
      Width           =   495
   End
   Begin VB.Label Label3 
      Caption         =   "Upkeep:"
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
      Top             =   1560
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Price:"
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
      Top             =   1080
      Width           =   855
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
      Top             =   600
      Width           =   855
   End
   Begin VB.Label Label4 
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
      TabIndex        =   0
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
      Left            =   1200
      TabIndex        =   6
      Top             =   120
      Width           =   2895
   End
End
Attribute VB_Name = "frmHall"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub btnCancel_Click()
    Unload Me
End Sub


Private Sub btnOk_Click()
    Dim A As Long, B As Long

    A = Int(Val(txtPrice))
    B = Int(Val(txtUpkeep))
    If A > 0 And A < 1000000000 And B > 0 And B < 1000000000 Then
        SendSocket Chr$(49) + Chr$(lblNumber) + QuadChar(A) + QuadChar(B) + DoubleChar(sclStartMap) + Chr$(sclStartX) + Chr$(sclStartY) + txtName
        Unload Me
    Else
        MsgBox "Invalid Price or Upkeep!", vbOKOnly + vbExclamation, TitleString
    End If
End Sub


Private Sub Form_Load()
    frmHall_Loaded = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmHall_Loaded = False
End Sub


Private Sub sclStartMap_Change()
    lblStartMap = sclStartMap
End Sub


Private Sub sclStartMap_Scroll()
    sclStartMap_Change
End Sub


Private Sub sclStartX_Change()
    lblStartX = sclStartX
End Sub


Private Sub sclStartX_Scroll()
    sclStartX_Change
End Sub


Private Sub sclStartY_Change()
    lblStartY = sclStartY
End Sub


Private Sub sclStartY_Scroll()
    sclStartY_Change
End Sub


