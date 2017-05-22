VERSION 5.00
Begin VB.Form frmMapAtt 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "The Odyssey Online Classic [Map Attribute]"
   ClientHeight    =   2745
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4305
   ControlBox      =   0   'False
   Icon            =   "frmMapAtt.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2745
   ScaleWidth      =   4305
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picAtt9 
      Height          =   1455
      Left            =   120
      ScaleHeight     =   1395
      ScaleWidth      =   4035
      TabIndex        =   38
      Top             =   600
      Visible         =   0   'False
      Width           =   4095
      Begin VB.HScrollBar sclAtt9Damage 
         Height          =   255
         Left            =   1080
         Max             =   50
         Min             =   1
         TabIndex        =   39
         Top             =   120
         Value           =   1
         Width           =   2295
      End
      Begin VB.Label Label14 
         Caption         =   "Damage:"
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
         TabIndex        =   41
         Top             =   120
         Width           =   855
      End
      Begin VB.Label lblAtt9Damage 
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
         Left            =   3480
         TabIndex        =   40
         Top             =   120
         Width           =   495
      End
   End
   Begin VB.PictureBox picAtt8 
      Height          =   1455
      Left            =   120
      ScaleHeight     =   1395
      ScaleWidth      =   4035
      TabIndex        =   28
      Top             =   600
      Visible         =   0   'False
      Width           =   4095
      Begin VB.HScrollBar sclAtt8Hall 
         Height          =   255
         LargeChange     =   10
         Left            =   720
         Max             =   255
         TabIndex        =   35
         Top             =   840
         Width           =   2415
      End
      Begin VB.HScrollBar sclAtt8X 
         Height          =   255
         Left            =   720
         Max             =   11
         TabIndex        =   30
         Top             =   120
         Width           =   2415
      End
      Begin VB.HScrollBar sclAtt8Y 
         Height          =   255
         Left            =   720
         Max             =   11
         TabIndex        =   29
         Top             =   480
         Width           =   2415
      End
      Begin VB.Label lblAtt8Hall 
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
         Left            =   3240
         TabIndex        =   37
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label12 
         Caption         =   "Hall:"
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
         Top             =   840
         Width           =   495
      End
      Begin VB.Label lblAtt8X 
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
         Left            =   3240
         TabIndex        =   34
         Top             =   120
         Width           =   735
      End
      Begin VB.Label Label10 
         Caption         =   "X:"
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
         Top             =   120
         Width           =   495
      End
      Begin VB.Label lblAtt8Y 
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
         Left            =   3240
         TabIndex        =   32
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label5 
         Caption         =   "Y:"
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
         Top             =   480
         Width           =   495
      End
   End
   Begin VB.PictureBox picAtt7 
      Height          =   1455
      Left            =   120
      ScaleHeight     =   1395
      ScaleWidth      =   4035
      TabIndex        =   21
      Top             =   600
      Visible         =   0   'False
      Width           =   4095
      Begin VB.HScrollBar sclAtt7Val 
         Height          =   255
         Left            =   720
         Max             =   32000
         TabIndex        =   25
         Top             =   480
         Width           =   2415
      End
      Begin VB.HScrollBar sclAtt7Obj 
         Height          =   255
         Left            =   720
         Max             =   255
         Min             =   1
         TabIndex        =   22
         Top             =   120
         Value           =   1
         Width           =   2415
      End
      Begin VB.Label Label8 
         Caption         =   "Val:"
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
         Top             =   480
         Width           =   495
      End
      Begin VB.Label lblAtt7Val 
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
         Left            =   3240
         TabIndex        =   26
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label7 
         Caption         =   "Obj:"
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
         Top             =   120
         Width           =   495
      End
      Begin VB.Label lblAtt7Obj 
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
         Height          =   375
         Left            =   3240
         TabIndex        =   23
         Top             =   120
         Width           =   735
      End
   End
   Begin VB.PictureBox picAtt6 
      Height          =   1455
      Left            =   120
      ScaleHeight     =   1395
      ScaleWidth      =   4035
      TabIndex        =   17
      Top             =   600
      Visible         =   0   'False
      Width           =   4095
      Begin VB.HScrollBar sclAtt6Num 
         Height          =   255
         Left            =   720
         Max             =   255
         Min             =   1
         TabIndex        =   18
         Top             =   120
         Value           =   1
         Width           =   2415
      End
      Begin VB.Label lblAtt6Num 
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
         Height          =   375
         Left            =   3240
         TabIndex        =   20
         Top             =   120
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "Num:"
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
         Width           =   495
      End
   End
   Begin VB.PictureBox picAtt3 
      Height          =   1455
      Left            =   120
      ScaleHeight     =   1395
      ScaleWidth      =   4035
      TabIndex        =   13
      Top             =   600
      Visible         =   0   'False
      Width           =   4095
      Begin VB.HScrollBar sclAtt3Key 
         Height          =   255
         Left            =   720
         Max             =   255
         Min             =   1
         TabIndex        =   14
         Top             =   120
         Value           =   1
         Width           =   2415
      End
      Begin VB.Label Label9 
         Caption         =   "Key:"
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
         Top             =   120
         Width           =   495
      End
      Begin VB.Label lblAtt3Key 
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
         Height          =   375
         Left            =   3240
         TabIndex        =   15
         Top             =   120
         Width           =   735
      End
   End
   Begin VB.PictureBox picAtt2 
      Height          =   1455
      Left            =   120
      ScaleHeight     =   1395
      ScaleWidth      =   4035
      TabIndex        =   3
      Top             =   600
      Visible         =   0   'False
      Width           =   4095
      Begin VB.HScrollBar sclAtt2Y 
         Height          =   255
         Left            =   720
         Max             =   11
         TabIndex        =   9
         Top             =   840
         Width           =   2415
      End
      Begin VB.HScrollBar sclAtt2X 
         Height          =   255
         Left            =   720
         Max             =   11
         TabIndex        =   8
         Top             =   480
         Width           =   2415
      End
      Begin VB.HScrollBar sclAtt2Map 
         Height          =   255
         LargeChange     =   25
         Left            =   720
         Max             =   2000
         Min             =   1
         TabIndex        =   7
         Top             =   120
         Value           =   1
         Width           =   2415
      End
      Begin VB.Label lblAtt2Y 
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
         Left            =   3240
         TabIndex        =   12
         Top             =   840
         Width           =   735
      End
      Begin VB.Label lblAtt2X 
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
         Left            =   3240
         TabIndex        =   11
         Top             =   480
         Width           =   735
      End
      Begin VB.Label lblAtt2Map 
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
         Height          =   375
         Left            =   3240
         TabIndex        =   10
         Top             =   120
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "Y:"
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
         TabIndex        =   6
         Top             =   840
         Width           =   495
      End
      Begin VB.Label Label2 
         Caption         =   "X:"
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
         TabIndex        =   5
         Top             =   480
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "Map:"
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
         TabIndex        =   4
         Top             =   120
         Width           =   495
      End
   End
   Begin VB.CommandButton btnCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   495
      Left            =   1200
      TabIndex        =   1
      Top             =   2160
      Width           =   1455
   End
   Begin VB.CommandButton btnOk 
      Caption         =   "Ok"
      Height          =   495
      Left            =   2760
      TabIndex        =   0
      Top             =   2160
      Width           =   1455
   End
   Begin VB.Label lblAtt 
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
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   4095
   End
End
Attribute VB_Name = "frmMapAtt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub btnOk_Click()
    CurAtt = NewAtt
    Select Case CurAtt
        Case 2 'Warp
            CurAttData(0) = Int(sclAtt2Map / 256)
            CurAttData(1) = sclAtt2Map Mod 256
            CurAttData(2) = sclAtt2X
            CurAttData(3) = sclAtt2Y
        Case 3 'Key
            CurAttData(0) = sclAtt3Key
            CurAttData(1) = 0
            CurAttData(2) = 0
            CurAttData(3) = 0
        Case 6 'News
            CurAttData(0) = sclAtt6Num
            CurAttData(1) = 0
            CurAttData(2) = 0
            CurAttData(3) = 0
        Case 7 'Object
            CurAttData(0) = sclAtt7Obj
            CurAttData(1) = Int(sclAtt7Val / 65536)
            CurAttData(2) = Int(sclAtt7Val / 256) Mod 256
            CurAttData(3) = sclAtt7Val Mod 256
        Case 8 'Touch Plate
            CurAttData(0) = sclAtt8X
            CurAttData(1) = sclAtt8Y
            CurAttData(2) = sclAtt8Hall
            CurAttData(3) = 0
        Case 9 'Damage
            CurAttData(0) = sclAtt9Damage
            CurAttData(1) = 0
            CurAttData(2) = 0
            CurAttData(3) = 0
    End Select
    Unload Me
End Sub


Private Sub Form_Load()
    Select Case NewAtt
        Case 2 'Warp
            lblAtt = "2 - Warp"
            picAtt2.Visible = True
        Case 3 'Key
            lblAtt = "3 - Key"
            picAtt3.Visible = True
        Case 6 'News
            lblAtt = "6 - News"
            picAtt6.Visible = True
        Case 7 'Obj
            lblAtt = "7 - Object"
            picAtt7.Visible = True
        Case 8 'Touch Plate
            lblAtt = "8 - Touch Plate"
            picAtt8.Visible = True
        Case 9 'Damage
            lblAtt = "9 - Damage"
            picAtt9.Visible = True
    End Select
End Sub


Private Sub Label6_Click()

End Sub

Private Sub sclAtt2Map_Change()
    lblAtt2Map = sclAtt2Map
End Sub


Private Sub sclAtt2Map_Scroll()
    sclAtt2Map_Change
End Sub


Private Sub sclAtt2X_Change()
    lblAtt2X = sclAtt2X
End Sub


Private Sub sclAtt2X_Scroll()
    sclAtt2X_Change
End Sub


Private Sub sclAtt2Y_Change()
    lblAtt2Y = sclAtt2Y
End Sub


Private Sub sclAtt2Y_Scroll()
    sclAtt2Y_Change
End Sub


Private Sub sclAtt3Key_Change()
    lblAtt3Key = sclAtt3Key
End Sub

Private Sub sclAtt3Key_Scroll()
    sclAtt3Key_Change
End Sub


Private Sub sclAtt6Num_Change()
    lblAtt6Num = sclAtt6Num
End Sub

Private Sub sclAtt6Num_Scroll()
    sclAtt6Num_Change
End Sub


Private Sub sclAtt7Obj_Change()
    lblAtt7Obj = sclAtt7Obj
End Sub


Private Sub sclAtt7Obj_Scroll()
    sclAtt7Obj_Change
End Sub


Private Sub sclAtt7Val_Change()
    lblAtt7Val = sclAtt7Val
End Sub


Private Sub sclAtt7Val_Scroll()
    sclAtt7Val_Change
End Sub


Private Sub sclAtt8Hall_Change()
    lblAtt8Hall = sclAtt8Hall
End Sub

Private Sub sclAtt8Hall_Scroll()
    sclAtt8Hall_Change
End Sub


Private Sub sclAtt8X_Change()
    lblAtt8X = sclAtt8X
End Sub


Private Sub sclAtt8X_Scroll()
    sclAtt8X_Change
End Sub


Private Sub sclAtt8Y_Change()
    lblAtt8Y = sclAtt8Y
End Sub


Private Sub sclAtt8Y_Scroll()
    sclAtt8Y_Change
End Sub


Private Sub sclAtt9Damage_Change()
    lblAtt9Damage = sclAtt9Damage
End Sub


Private Sub sclAtt9Damage_Scroll()
    sclAtt9Damage_Change
End Sub


