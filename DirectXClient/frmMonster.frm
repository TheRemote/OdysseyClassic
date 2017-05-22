VERSION 5.00
Begin VB.Form frmMonster 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "The Odyssey Online Classic [Edit Monster]"
   ClientHeight    =   7215
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6105
   ControlBox      =   0   'False
   Icon            =   "frmMonster.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   481
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   407
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkFlag 
      Height          =   255
      Index           =   7
      Left            =   3480
      TabIndex        =   51
      Top             =   5880
      Width           =   1095
   End
   Begin VB.CheckBox chkFlag 
      Caption         =   "Attacks Monsters"
      Height          =   255
      Index           =   6
      Left            =   3480
      TabIndex        =   50
      Top             =   5640
      Width           =   1575
   End
   Begin VB.CheckBox chkFlag 
      Caption         =   "Fast Hit"
      Height          =   255
      Index           =   5
      Left            =   2280
      TabIndex        =   49
      Top             =   6120
      Width           =   1095
   End
   Begin VB.CheckBox chkFlag 
      Caption         =   "Elite"
      Height          =   255
      Index           =   4
      Left            =   2280
      TabIndex        =   48
      Top             =   5880
      Width           =   1095
   End
   Begin VB.HScrollBar sclMagicDefense 
      Height          =   255
      Left            =   1200
      Max             =   255
      TabIndex        =   46
      Top             =   3960
      Value           =   1
      Width           =   4215
   End
   Begin VB.HScrollBar sclExperience 
      Height          =   255
      Left            =   1200
      Max             =   255
      TabIndex        =   42
      Top             =   3600
      Width           =   4215
   End
   Begin VB.CheckBox chkFlag 
      Caption         =   "Runner"
      Height          =   255
      Index           =   2
      Left            =   1080
      TabIndex        =   41
      Top             =   6120
      Width           =   1095
   End
   Begin VB.CheckBox chkFlag 
      Caption         =   "Supersight"
      Height          =   255
      Index           =   1
      Left            =   1080
      TabIndex        =   40
      Top             =   5880
      Width           =   1095
   End
   Begin VB.CheckBox chkFlag 
      Caption         =   "Friendly"
      Height          =   255
      Index           =   3
      Left            =   2280
      TabIndex        =   38
      Top             =   5640
      Width           =   1095
   End
   Begin VB.CheckBox chkFlag 
      Caption         =   "Guard"
      Height          =   255
      Index           =   0
      Left            =   1080
      TabIndex        =   37
      Top             =   5640
      Width           =   1095
   End
   Begin VB.HScrollBar sclAgility 
      Height          =   255
      Left            =   1200
      Max             =   255
      TabIndex        =   7
      Top             =   3240
      Width           =   4215
   End
   Begin VB.HScrollBar sclSight 
      Height          =   255
      Left            =   1200
      Max             =   255
      Min             =   1
      TabIndex        =   6
      Top             =   2880
      Value           =   1
      Width           =   4215
   End
   Begin VB.TextBox txtValue 
      Height          =   335
      Index           =   2
      Left            =   3480
      MaxLength       =   4
      TabIndex        =   13
      Top             =   5160
      Width           =   1095
   End
   Begin VB.TextBox txtValue 
      Height          =   335
      Index           =   1
      Left            =   3480
      MaxLength       =   4
      TabIndex        =   11
      Top             =   4800
      Width           =   1095
   End
   Begin VB.TextBox txtValue 
      Height          =   335
      Index           =   0
      Left            =   3480
      MaxLength       =   4
      TabIndex        =   9
      Top             =   4440
      Width           =   1095
   End
   Begin VB.HScrollBar sclArmor 
      Height          =   255
      Left            =   1200
      Max             =   255
      TabIndex        =   4
      Top             =   2160
      Width           =   4215
   End
   Begin VB.ComboBox cmbObject 
      Height          =   315
      Index           =   1
      Left            =   1080
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   4800
      Width           =   2295
   End
   Begin VB.ComboBox cmbObject 
      Height          =   315
      Index           =   2
      Left            =   1080
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   5160
      Width           =   2295
   End
   Begin VB.ComboBox cmbObject 
      Height          =   315
      Index           =   0
      Left            =   1080
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   4440
      Width           =   2295
   End
   Begin VB.HScrollBar sclSpeed 
      Height          =   255
      Left            =   1200
      Max             =   255
      Min             =   1
      TabIndex        =   5
      Top             =   2520
      Value           =   1
      Width           =   4215
   End
   Begin VB.HScrollBar sclStrength 
      Height          =   255
      Left            =   1200
      Max             =   255
      Min             =   1
      TabIndex        =   3
      Top             =   1800
      Value           =   1
      Width           =   4215
   End
   Begin VB.HScrollBar sclHP 
      Height          =   255
      Left            =   1200
      Max             =   32000
      Min             =   1
      TabIndex        =   2
      Top             =   1440
      Value           =   1
      Width           =   4215
   End
   Begin VB.PictureBox picSprite 
      AutoRedraw      =   -1  'True
      Height          =   540
      Left            =   4200
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   18
      Top             =   0
      Width           =   540
   End
   Begin VB.HScrollBar sclSprite 
      Height          =   255
      Left            =   1200
      Max             =   255
      Min             =   1
      TabIndex        =   1
      Top             =   1080
      Value           =   1
      Width           =   4215
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
      TabIndex        =   0
      Top             =   600
      Width           =   2895
   End
   Begin VB.CommandButton btnCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   495
      Left            =   1560
      TabIndex        =   14
      Top             =   6600
      Width           =   1455
   End
   Begin VB.CommandButton btnOk 
      Caption         =   "Ok"
      Height          =   495
      Left            =   3120
      TabIndex        =   15
      Top             =   6600
      Width           =   1455
   End
   Begin VB.Label lblMagicDefense 
      Alignment       =   2  'Center
      Caption         =   "1"
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
      Left            =   5520
      TabIndex        =   47
      Top             =   3960
      Width           =   495
   End
   Begin VB.Label Label15 
      Caption         =   "Magic Df:"
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
      TabIndex        =   45
      Top             =   3960
      Width           =   1095
   End
   Begin VB.Label lblExperience 
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
      Left            =   5520
      TabIndex        =   44
      Top             =   3600
      Width           =   495
   End
   Begin VB.Label Label14 
      Caption         =   "Experience:"
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
      TabIndex        =   43
      Top             =   3600
      Width           =   1095
   End
   Begin VB.Label lblSprite 
      Alignment       =   2  'Center
      Caption         =   "1"
      Height          =   255
      Left            =   5520
      TabIndex        =   39
      Top             =   1080
      Width           =   495
   End
   Begin VB.Label Label10 
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
      Left            =   120
      TabIndex        =   36
      Top             =   5640
      Width           =   855
   End
   Begin VB.Label Label13 
      Caption         =   "Agility"
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
      TabIndex        =   35
      Top             =   3240
      Width           =   855
   End
   Begin VB.Label lblAgility 
      Alignment       =   2  'Center
      Caption         =   "1"
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
      Left            =   5520
      TabIndex        =   34
      Top             =   3240
      Width           =   495
   End
   Begin VB.Label Label12 
      Caption         =   "Sight:"
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
      TabIndex        =   33
      Top             =   2880
      Width           =   855
   End
   Begin VB.Label lblSight 
      Alignment       =   2  'Center
      Caption         =   "1"
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
      Left            =   5520
      TabIndex        =   32
      Top             =   2880
      Width           =   495
   End
   Begin VB.Label Label11 
      Caption         =   "Armor:"
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
      TabIndex        =   31
      Top             =   2160
      Width           =   855
   End
   Begin VB.Label lblArmor 
      Alignment       =   2  'Center
      Caption         =   "1"
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
      Left            =   5520
      TabIndex        =   30
      Top             =   2160
      Width           =   495
   End
   Begin VB.Label Label9 
      Caption         =   "Object 3:"
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
      TabIndex        =   29
      Top             =   5160
      Width           =   855
   End
   Begin VB.Label Label8 
      Caption         =   "Object 2:"
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
      TabIndex        =   28
      Top             =   4800
      Width           =   855
   End
   Begin VB.Label Label6 
      Caption         =   "Object 1:"
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
      TabIndex        =   27
      Top             =   4440
      Width           =   855
   End
   Begin VB.Label Label7 
      Caption         =   "Speed:"
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
      Top             =   2520
      Width           =   855
   End
   Begin VB.Label lblSpeed 
      Alignment       =   2  'Center
      Caption         =   "1"
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
      Left            =   5520
      TabIndex        =   25
      Top             =   2520
      Width           =   495
   End
   Begin VB.Label lblStrength 
      Alignment       =   2  'Center
      Caption         =   "1"
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
      Left            =   5520
      TabIndex        =   24
      Top             =   1800
      Width           =   495
   End
   Begin VB.Label lblHP 
      Alignment       =   2  'Center
      Caption         =   "1"
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
      Left            =   5520
      TabIndex        =   23
      Top             =   1440
      Width           =   495
   End
   Begin VB.Label Label5 
      Caption         =   "Strength:"
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
      Top             =   1800
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "HP:"
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
      Top             =   1440
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
      Left            =   1080
      TabIndex        =   20
      Top             =   120
      Width           =   2655
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
      TabIndex        =   19
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "Sprite:"
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
      TabIndex        =   16
      Top             =   600
      Width           =   855
   End
End
Attribute VB_Name = "frmMonster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub btnCancel_Click()
    Me.Hide
End Sub

Private Sub btnOk_Click()
    Dim V1 As Long, V2 As Long, V3 As Long
    Dim A As Long, flags As Byte

    For A = 0 To 7
        If chkFlag(A) = 1 Then
            SetBit flags, CByte(A)
        Else
            ClearBit flags, CByte(A)
        End If
    Next A

    V1 = Val(txtValue(0))
    If V1 < 0 Then V1 = 0
    V2 = Val(txtValue(1))
    If V2 < 0 Then V2 = 0
    V3 = Val(txtValue(2))
    If V3 < 0 Then V3 = 0
    SendSocket Chr$(22) + DoubleChar$(lblNumber) + DoubleChar$(sclSprite) + DoubleChar$(sclHP) + Chr$(sclStrength) + Chr$(sclArmor) + Chr$(sclSpeed) + Chr$(sclSight) + Chr$(sclAgility) + Chr$(flags) + DoubleChar$(cmbObject(0).ListIndex) + Chr$(V1) + DoubleChar$(cmbObject(1).ListIndex) + Chr$(V2) + DoubleChar$(cmbObject(2).ListIndex) + Chr$(V3) + Chr$(sclExperience.value) + Chr$(sclMagicDefense) + txtName
    Me.Hide
End Sub
Private Sub Form_Load()
    Dim A As Long
    cmbObject(0).AddItem "<None>"
    cmbObject(1).AddItem "<None>"
    cmbObject(2).AddItem "<None>"
    For A = 1 To MaxObjects
        cmbObject(0).AddItem CStr(A) + ": " + Object(A).name
        cmbObject(1).AddItem CStr(A) + ": " + Object(A).name
        cmbObject(2).AddItem CStr(A) + ": " + Object(A).name
    Next A
    sclSprite.max = MaxSprite

    frmMonster_Loaded = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmMonster_Loaded = False
End Sub

Private Sub sclAgility_Change()
    lblAgility = sclAgility
End Sub

Private Sub sclAgility_Scroll()
    sclAgility_Change
End Sub

Private Sub sclArmor_Change()
    lblArmor = sclArmor
End Sub

Private Sub sclArmor_Scroll()
    sclArmor_Change
End Sub

Private Sub sclExperience_Change()
    lblExperience = sclExperience * 10
End Sub

Private Sub sclExperience_Scroll()
    sclExperience_Change
End Sub

Private Sub sclHP_Change()
    lblHP = sclHP
End Sub

Private Sub sclHP_Scroll()
    sclHP_Change
End Sub

Private Sub sclMagicDefense_Change()
    lblMagicDefense = sclMagicDefense
End Sub

Private Sub sclSight_Change()
    lblSight = sclSight
End Sub

Private Sub sclSight_Scroll()
    sclSight_Change
End Sub

Private Sub sclSpeed_Change()
    lblSpeed = sclSpeed
End Sub

Private Sub sclSpeed_Scroll()
    sclSpeed_Change
End Sub


Private Sub sclSprite_Change()
    lblSprite.Caption = sclSprite
    picSprite.Height = 36
    picSprite.Width = 36
    'BitBlt picSprite.hDC, 0, 0, 32, 32, hdcSprites, 96, (sclSprite - 1) * 32, SRCCOPY
    DrawToDC 0, 0, 32, 32, picSprite.hDC, DDSSprites, 96, (sclSprite - 1) * 32
    picSprite.Refresh
End Sub

Private Sub sclSprite_Scroll()
    sclSprite_Change
End Sub


Private Sub sclStrength_Change()
    lblStrength = sclStrength
End Sub


Private Sub sclStrength_Scroll()
    sclStrength_Change
End Sub


