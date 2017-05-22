VERSION 5.00
Begin VB.Form frmPrefix 
   Caption         =   "The Odyssey Online Classic [Editing Prefix]"
   ClientHeight    =   5235
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5925
   ControlBox      =   0   'False
   DrawMode        =   15  'Merge Pen Not
   Icon            =   "frmPrefix.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5235
   ScaleWidth      =   5925
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   3240
      TabIndex        =   22
      Top             =   4800
      Width           =   1575
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1200
      TabIndex        =   21
      Top             =   4800
      Width           =   1575
   End
   Begin VB.HScrollBar sclValue 
      Height          =   255
      Left            =   1680
      Max             =   255
      TabIndex        =   19
      Top             =   4320
      Width           =   3495
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   1680
      MaxLength       =   15
      TabIndex        =   17
      Top             =   480
      Width           =   2295
   End
   Begin VB.CheckBox chkOccursNaturally 
      Caption         =   "Occurs Naturally"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   1320
      Width           =   1575
   End
   Begin VB.OptionButton optModType 
      Caption         =   "Magic Defense"
      Height          =   255
      Index           =   13
      Left            =   1680
      TabIndex        =   15
      Top             =   3960
      Width           =   1455
   End
   Begin VB.OptionButton optModType 
      Caption         =   "Defense"
      Height          =   255
      Index           =   12
      Left            =   1680
      TabIndex        =   14
      Top             =   3720
      Width           =   1215
   End
   Begin VB.OptionButton optModType 
      Caption         =   "Damage"
      Height          =   255
      Index           =   11
      Left            =   1680
      TabIndex        =   13
      Top             =   3480
      Width           =   1215
   End
   Begin VB.OptionButton optModType 
      Caption         =   "Max Mana"
      Height          =   255
      Index           =   10
      Left            =   1680
      TabIndex        =   12
      Top             =   3240
      Width           =   1215
   End
   Begin VB.OptionButton optModType 
      Caption         =   "Max Energy"
      Height          =   255
      Index           =   9
      Left            =   1680
      TabIndex        =   11
      Top             =   3000
      Width           =   1215
   End
   Begin VB.OptionButton optModType 
      Caption         =   "Max HP"
      Height          =   255
      Index           =   8
      Left            =   1680
      TabIndex        =   10
      Top             =   2760
      Width           =   1215
   End
   Begin VB.OptionButton optModType 
      Caption         =   "Wisdom"
      Height          =   255
      Index           =   7
      Left            =   1680
      TabIndex        =   9
      Top             =   2520
      Width           =   1215
   End
   Begin VB.OptionButton optModType 
      Caption         =   "Stamina"
      Height          =   255
      Index           =   6
      Left            =   1680
      TabIndex        =   8
      Top             =   2280
      Width           =   1215
   End
   Begin VB.OptionButton optModType 
      Caption         =   "Constitution"
      Height          =   255
      Index           =   5
      Left            =   1680
      TabIndex        =   7
      Top             =   2040
      Width           =   1215
   End
   Begin VB.OptionButton optModType 
      Caption         =   "Concentration"
      Height          =   255
      Index           =   4
      Left            =   1680
      TabIndex        =   6
      Top             =   1800
      Width           =   1335
   End
   Begin VB.OptionButton optModType 
      Caption         =   "Intelligence"
      Height          =   255
      Index           =   3
      Left            =   1680
      TabIndex        =   5
      Top             =   1560
      Width           =   1215
   End
   Begin VB.OptionButton optModType 
      Caption         =   "Agility"
      Height          =   255
      Index           =   2
      Left            =   1680
      TabIndex        =   4
      Top             =   1320
      Width           =   1215
   End
   Begin VB.OptionButton optModType 
      Caption         =   "Endurance"
      Height          =   255
      Index           =   1
      Left            =   1680
      TabIndex        =   3
      Top             =   1080
      Width           =   1215
   End
   Begin VB.OptionButton optModType 
      Caption         =   "Strength"
      Height          =   255
      Index           =   0
      Left            =   1680
      TabIndex        =   2
      Top             =   840
      Width           =   975
   End
   Begin VB.Label cptlblNumber 
      Alignment       =   1  'Right Justify
      Caption         =   "Number:"
      Height          =   255
      Left            =   120
      TabIndex        =   24
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label lblNumber 
      Alignment       =   2  'Center
      Caption         =   "0"
      Height          =   255
      Left            =   1680
      TabIndex        =   23
      Top             =   120
      Width           =   615
   End
   Begin VB.Label lblModValue 
      Alignment       =   2  'Center
      Caption         =   "0"
      Height          =   255
      Left            =   5280
      TabIndex        =   20
      Top             =   4320
      Width           =   615
   End
   Begin VB.Label lblModificationValue 
      Caption         =   "Modification Value:"
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   4320
      Width           =   1455
   End
   Begin VB.Label lblName 
      Alignment       =   1  'Right Justify
      Caption         =   "Name:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label lblModificationType 
      Caption         =   "Modification Type:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   1335
   End
End
Attribute VB_Name = "frmPrefix"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    Me.Hide
End Sub

Private Sub cmdSave_Click()
    Dim A As Long, TheType As Byte
    For A = 0 To 13
        If optModType(A).value = True Then TheType = A
    Next A
    SendSocket Chr$(87) + Chr$(lblNumber) + Chr$(TheType) + Chr$(lblModValue) + Chr$(chkOccursNaturally.value) + txtName
    Me.Hide
End Sub

Private Sub Form_Load()
    frmPrefix_Loaded = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmPrefix_Loaded = False
End Sub

Private Sub sclValue_Change()
    lblModValue = sclValue
End Sub
