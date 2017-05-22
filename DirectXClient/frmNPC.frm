VERSION 5.00
Begin VB.Form frmNPC 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "The Odyssey Online Classic [Editing NPC]"
   ClientHeight    =   6540
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9225
   ControlBox      =   0   'False
   Icon            =   "frmNPC.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6540
   ScaleWidth      =   9225
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkFlag 
      Caption         =   "Can Sell To"
      Height          =   255
      Index           =   2
      Left            =   6600
      TabIndex        =   34
      Top             =   360
      Width           =   1215
   End
   Begin VB.CheckBox chkFlag 
      Caption         =   "Can Repair"
      Height          =   255
      Index           =   1
      Left            =   7920
      TabIndex        =   33
      Top             =   120
      Width           =   1215
   End
   Begin VB.CheckBox chkFlag 
      Caption         =   "Banker"
      Height          =   255
      Index           =   0
      Left            =   6600
      TabIndex        =   31
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton btnCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   495
      Left            =   6120
      TabIndex        =   14
      Top             =   6000
      Width           =   1455
   End
   Begin VB.CommandButton btnOk 
      Caption         =   "Ok"
      Height          =   495
      Left            =   7680
      TabIndex        =   15
      Top             =   6000
      Width           =   1455
   End
   Begin VB.CommandButton btnUpdate 
      Caption         =   "<-- Update"
      Height          =   375
      Left            =   5280
      TabIndex        =   13
      Top             =   5520
      Width           =   3855
   End
   Begin VB.TextBox txtTakeValue 
      Height          =   315
      Left            =   8160
      MaxLength       =   9
      TabIndex        =   12
      Top             =   5160
      Width           =   975
   End
   Begin VB.ComboBox cmbTakeObject 
      Height          =   315
      Left            =   6360
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   5160
      Width           =   1695
   End
   Begin VB.TextBox txtGiveValue 
      Height          =   315
      Left            =   8160
      MaxLength       =   9
      TabIndex        =   10
      Top             =   4800
      Width           =   975
   End
   Begin VB.ComboBox cmbGiveObject 
      Height          =   315
      Left            =   6360
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   4800
      Width           =   1695
   End
   Begin VB.ListBox lstSaleItems 
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1875
      Left            =   1320
      TabIndex        =   8
      Top             =   4440
      Width           =   3855
   End
   Begin VB.TextBox txtSayText5 
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
      MaxLength       =   255
      TabIndex        =   7
      Top             =   3960
      Width           =   7815
   End
   Begin VB.TextBox txtSayText4 
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
      MaxLength       =   255
      TabIndex        =   6
      Top             =   3480
      Width           =   7815
   End
   Begin VB.TextBox txtSayText3 
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
      MaxLength       =   255
      TabIndex        =   5
      Top             =   3000
      Width           =   7815
   End
   Begin VB.TextBox txtSayText2 
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
      MaxLength       =   255
      TabIndex        =   4
      Top             =   2520
      Width           =   7815
   End
   Begin VB.TextBox txtSayText1 
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
      MaxLength       =   255
      TabIndex        =   3
      Top             =   2040
      Width           =   7815
   End
   Begin VB.TextBox txtLeaveText 
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
      MaxLength       =   255
      TabIndex        =   2
      Top             =   1560
      Width           =   7815
   End
   Begin VB.TextBox txtJoinText 
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
      MaxLength       =   255
      TabIndex        =   1
      Top             =   1080
      Width           =   7815
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
      Width           =   3855
   End
   Begin VB.Label Label14 
      Caption         =   "Flags:"
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
      Left            =   5400
      TabIndex        =   32
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label lblItemNumber 
      Height          =   375
      Left            =   6360
      TabIndex        =   30
      Top             =   4440
      Width           =   1935
   End
   Begin VB.Label Label13 
      Caption         =   "Item Number:"
      Height          =   375
      Left            =   5280
      TabIndex        =   29
      Top             =   4440
      Width           =   975
   End
   Begin VB.Label Label12 
      Caption         =   "TakeObject:"
      Height          =   375
      Left            =   5280
      TabIndex        =   28
      Top             =   5160
      Width           =   975
   End
   Begin VB.Label Label11 
      Caption         =   "GiveObject:"
      Height          =   375
      Left            =   5280
      TabIndex        =   27
      Top             =   4800
      Width           =   975
   End
   Begin VB.Label Label10 
      Caption         =   "Sale Items:"
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
      TabIndex        =   26
      Top             =   4440
      Width           =   1095
   End
   Begin VB.Label Label9 
      Caption         =   "SayText5:"
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
      TabIndex        =   25
      Top             =   3960
      Width           =   1095
   End
   Begin VB.Label Label8 
      Caption         =   "SayText4:"
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
      TabIndex        =   24
      Top             =   3480
      Width           =   1095
   End
   Begin VB.Label Label7 
      Caption         =   "SayText3:"
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
      TabIndex        =   23
      Top             =   3000
      Width           =   1095
   End
   Begin VB.Label Label6 
      Caption         =   "SayText2:"
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
      TabIndex        =   22
      Top             =   2520
      Width           =   1095
   End
   Begin VB.Label Label5 
      Caption         =   "SayText1:"
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
      TabIndex        =   21
      Top             =   2040
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "LeaveText:"
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
      TabIndex        =   20
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "JoinText:"
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
      TabIndex        =   19
      Top             =   1080
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
      TabIndex        =   18
      Top             =   600
      Width           =   1095
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
      TabIndex        =   17
      Top             =   120
      Width           =   1095
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
      Left            =   1320
      TabIndex        =   16
      Top             =   120
      Width           =   3855
   End
End
Attribute VB_Name = "frmNPC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub btnCancel_Click()
    Me.Hide
End Sub

Private Sub btnOk_Click()
    Dim A As Long, St As String
    Dim flags As Byte

    For A = 0 To 2
        If chkFlag(A) = 1 Then
            SetBit flags, CByte(A)
        Else
            ClearBit flags, CByte(A)
        End If
    Next A

    St = Chr$(51) + DoubleChar$(lblNumber) + Chr$(flags)
    For A = 0 To 9
        With NPC(lblNumber).SaleItem(A)
            St = St + DoubleChar$(CLng(.GiveObject)) + QuadChar(.GiveValue) + DoubleChar$(CLng(.TakeObject)) + QuadChar(.TakeValue)
        End With
    Next A
    St = St + txtName + vbNullChar + txtJoinText + vbNullChar + txtLeaveText + vbNullChar + txtSayText1 + vbNullChar + txtSayText2 + vbNullChar + txtSayText3 + vbNullChar + txtSayText4 + vbNullChar + txtSayText5
    SendSocket St
    Me.Hide
End Sub
Private Sub btnUpdate_Click()
    Dim A As Long, B As Long, C As Long
    B = Int(Val(txtGiveValue))
    C = Int(Val(txtTakeValue))
    If B >= 0 And C >= 0 Then
        A = lstSaleItems.ListIndex
        With NPC(lblNumber).SaleItem(A)
            .GiveObject = cmbGiveObject.ListIndex
            .GiveValue = B
            .TakeObject = cmbTakeObject.ListIndex
            .TakeValue = C
        End With
        UpdateSaleItem lblNumber, A
    Else
        MsgBox "Invalid Give or Take Values!", vbOKOnly + vbExclamation, TitleString
    End If
End Sub
Private Sub Form_Load()
    frmNPC_Loaded = True
End Sub
Private Sub Form_Unload(Cancel As Integer)
    frmNPC_Loaded = False
End Sub

Private Sub lstSaleItems_Click()
    With NPC(lblNumber).SaleItem(lstSaleItems.ListIndex)
        cmbGiveObject.ListIndex = .GiveObject
        txtGiveValue = .GiveValue
        cmbTakeObject.ListIndex = .TakeObject
        txtTakeValue = .TakeValue
    End With
End Sub

Sub UpdateList()
    Dim A As Long
    For A = 0 To 9
        lstSaleItems.AddItem CStr(A) + ":"
        UpdateSaleItem lblNumber, A
    Next A
    cmbGiveObject.AddItem "<nothing>"
    cmbTakeObject.AddItem "<nothing>"
    For A = 1 To MaxObjects
        cmbGiveObject.AddItem CStr(A) + ": " + Object(A).name
        cmbTakeObject.AddItem CStr(A) + ": " + Object(A).name
    Next A
    lstSaleItems.ListIndex = 0
End Sub
