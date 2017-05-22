VERSION 5.00
Begin VB.Form frmMapProperties 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "The Odyssey Online Classic [Map Properties]"
   ClientHeight    =   8880
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9720
   ControlBox      =   0   'False
   Icon            =   "frmMapProperties.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8880
   ScaleWidth      =   9720
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkFlag2 
      Height          =   195
      Index           =   7
      Left            =   6360
      TabIndex        =   79
      Top             =   8040
      Width           =   1215
   End
   Begin VB.CheckBox chkFlag2 
      Height          =   195
      Index           =   6
      Left            =   4920
      TabIndex        =   78
      Top             =   8040
      Width           =   1215
   End
   Begin VB.CheckBox chkFlag2 
      Height          =   195
      Index           =   5
      Left            =   3480
      TabIndex        =   77
      Top             =   8040
      Width           =   1215
   End
   Begin VB.CheckBox chkFlag2 
      Height          =   195
      Index           =   4
      Left            =   2400
      TabIndex        =   76
      Top             =   8040
      Width           =   1215
   End
   Begin VB.CheckBox chkFlag2 
      Height          =   195
      Index           =   3
      Left            =   1200
      TabIndex        =   75
      Top             =   8040
      Width           =   1215
   End
   Begin VB.CheckBox chkFlag2 
      Caption         =   "Don't Reset"
      Height          =   195
      Index           =   2
      Left            =   4920
      TabIndex        =   74
      Top             =   7680
      Width           =   1215
   End
   Begin VB.CheckBox chkFlag2 
      Caption         =   "Snowing"
      Height          =   195
      Index           =   1
      Left            =   3480
      TabIndex        =   73
      Top             =   7680
      Width           =   1215
   End
   Begin VB.CheckBox chkFlag2 
      Caption         =   "Raining"
      Height          =   195
      Index           =   0
      Left            =   2400
      TabIndex        =   72
      Top             =   7680
      Width           =   1215
   End
   Begin VB.TextBox txtDeathMap 
      Height          =   375
      Left            =   2640
      MaxLength       =   4
      TabIndex        =   67
      Top             =   5880
      Width           =   975
   End
   Begin VB.TextBox txtDeathX 
      Height          =   375
      Left            =   4320
      MaxLength       =   2
      TabIndex        =   66
      Top             =   5880
      Width           =   975
   End
   Begin VB.TextBox txtDeathY 
      Height          =   375
      Left            =   5880
      MaxLength       =   2
      TabIndex        =   65
      Top             =   5880
      Width           =   975
   End
   Begin VB.HScrollBar sclRate 
      Height          =   255
      Index           =   9
      Left            =   3960
      Max             =   255
      TabIndex        =   63
      Top             =   3600
      Width           =   2175
   End
   Begin VB.ComboBox cmbMonster 
      Height          =   315
      Index           =   9
      Left            =   3960
      Style           =   2  'Dropdown List
      TabIndex        =   62
      Top             =   3240
      Width           =   2175
   End
   Begin VB.HScrollBar sclRate 
      Height          =   255
      Index           =   8
      Left            =   6840
      Max             =   255
      TabIndex        =   60
      Top             =   2880
      Width           =   2175
   End
   Begin VB.ComboBox cmbMonster 
      Height          =   315
      Index           =   8
      Left            =   6840
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   2520
      Width           =   2175
   End
   Begin VB.HScrollBar sclRate 
      Height          =   255
      Index           =   7
      Left            =   3960
      Max             =   255
      TabIndex        =   58
      Top             =   2880
      Width           =   2175
   End
   Begin VB.ComboBox cmbMonster 
      Height          =   315
      Index           =   7
      Left            =   3960
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   2520
      Width           =   2175
   End
   Begin VB.HScrollBar sclRate 
      Height          =   255
      Index           =   6
      Left            =   1080
      Max             =   255
      TabIndex        =   56
      Top             =   2880
      Width           =   2175
   End
   Begin VB.ComboBox cmbMonster 
      Height          =   315
      Index           =   6
      Left            =   1080
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   2520
      Width           =   2175
   End
   Begin VB.HScrollBar sclRate 
      Height          =   255
      Index           =   5
      Left            =   6840
      Max             =   255
      TabIndex        =   54
      Top             =   2160
      Width           =   2175
   End
   Begin VB.ComboBox cmbMonster 
      Height          =   315
      Index           =   5
      Left            =   6840
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   1800
      Width           =   2175
   End
   Begin VB.HScrollBar sclRate 
      Height          =   255
      Index           =   4
      Left            =   3960
      Max             =   255
      TabIndex        =   52
      Top             =   2160
      Width           =   2175
   End
   Begin VB.ComboBox cmbMonster 
      Height          =   315
      Index           =   4
      Left            =   3960
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   1800
      Width           =   2175
   End
   Begin VB.HScrollBar sclRate 
      Height          =   255
      Index           =   3
      Left            =   1080
      Max             =   255
      TabIndex        =   50
      Top             =   2160
      Width           =   2175
   End
   Begin VB.ComboBox cmbMonster 
      Height          =   315
      Index           =   3
      Left            =   1080
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   1800
      Width           =   2175
   End
   Begin VB.CheckBox chkFlag 
      Caption         =   "Arena"
      Height          =   195
      Index           =   7
      Left            =   1200
      TabIndex        =   49
      Top             =   7680
      Width           =   975
   End
   Begin VB.TextBox txtBootY 
      Height          =   375
      Left            =   5880
      MaxLength       =   2
      TabIndex        =   21
      Top             =   5400
      Width           =   975
   End
   Begin VB.TextBox txtBootX 
      Height          =   375
      Left            =   4320
      MaxLength       =   2
      TabIndex        =   20
      Top             =   5400
      Width           =   975
   End
   Begin VB.TextBox txtBootMap 
      Height          =   375
      Left            =   2640
      MaxLength       =   4
      TabIndex        =   19
      Top             =   5400
      Width           =   975
   End
   Begin VB.TextBox txtRight 
      Height          =   375
      Left            =   4200
      MaxLength       =   4
      TabIndex        =   18
      Top             =   4800
      Width           =   975
   End
   Begin VB.TextBox txtLeft 
      Height          =   375
      Left            =   2040
      MaxLength       =   4
      TabIndex        =   17
      Top             =   4800
      Width           =   975
   End
   Begin VB.TextBox txtDown 
      Height          =   375
      Left            =   4200
      MaxLength       =   4
      TabIndex        =   16
      Top             =   4320
      Width           =   975
   End
   Begin VB.TextBox txtUp 
      Height          =   375
      Left            =   2040
      MaxLength       =   4
      TabIndex        =   15
      Top             =   4320
      Width           =   975
   End
   Begin VB.ComboBox cmbNPC 
      Height          =   315
      Left            =   240
      Style           =   2  'Dropdown List
      TabIndex        =   14
      Top             =   3840
      Width           =   2175
   End
   Begin VB.CheckBox chkFlag 
      Caption         =   "Anyone can fight"
      Height          =   195
      Index           =   6
      Left            =   4920
      TabIndex        =   43
      Top             =   7320
      Width           =   1935
   End
   Begin VB.CheckBox chkFlag 
      Caption         =   "Can't Attack Monsters"
      Height          =   195
      Index           =   5
      Left            =   2880
      TabIndex        =   42
      Top             =   7320
      Width           =   1935
   End
   Begin VB.CheckBox chkFlag 
      Caption         =   "Double Monsters"
      Height          =   195
      Index           =   4
      Left            =   1200
      TabIndex        =   26
      Top             =   7320
      Width           =   1695
   End
   Begin VB.CheckBox chkFlag 
      Caption         =   "Monsters Start on Map"
      Height          =   195
      Index           =   3
      Left            =   4920
      TabIndex        =   25
      Top             =   6960
      Width           =   1935
   End
   Begin VB.ComboBox cmbMonster 
      Height          =   315
      Index           =   2
      Left            =   6840
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   1080
      Width           =   2175
   End
   Begin VB.ComboBox cmbMonster 
      Height          =   315
      Index           =   1
      Left            =   3960
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1080
      Width           =   2175
   End
   Begin VB.ComboBox cmbMonster 
      Height          =   315
      Index           =   0
      Left            =   1080
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1080
      Width           =   2175
   End
   Begin VB.HScrollBar sclMIDI 
      Height          =   255
      Left            =   1080
      Max             =   255
      TabIndex        =   1
      Top             =   600
      Width           =   3855
   End
   Begin VB.HScrollBar sclRate 
      Height          =   255
      Index           =   2
      Left            =   6840
      Max             =   255
      TabIndex        =   12
      Top             =   1440
      Width           =   2175
   End
   Begin VB.HScrollBar sclRate 
      Height          =   255
      Index           =   1
      Left            =   3960
      Max             =   255
      TabIndex        =   13
      Top             =   1440
      Width           =   2175
   End
   Begin VB.HScrollBar sclRate 
      Height          =   255
      Index           =   0
      Left            =   1080
      Max             =   255
      TabIndex        =   11
      Top             =   1440
      Width           =   2175
   End
   Begin VB.CheckBox chkFlag 
      Caption         =   "Always Dark"
      Height          =   195
      Index           =   2
      Left            =   3480
      TabIndex        =   24
      Top             =   6960
      Width           =   1215
   End
   Begin VB.CheckBox chkFlag 
      Caption         =   "Indoors"
      Height          =   195
      Index           =   1
      Left            =   2400
      TabIndex        =   23
      Top             =   6960
      Width           =   975
   End
   Begin VB.CheckBox chkFlag 
      Caption         =   "Friendly"
      Height          =   195
      Index           =   0
      Left            =   1200
      TabIndex        =   22
      Top             =   6960
      Width           =   1215
   End
   Begin VB.TextBox txtName 
      Height          =   375
      Left            =   1080
      MaxLength       =   30
      TabIndex        =   0
      Top             =   120
      Width           =   3855
   End
   Begin VB.CommandButton btnOk 
      Caption         =   "Ok"
      Height          =   495
      Left            =   4800
      TabIndex        =   28
      Top             =   8280
      Width           =   1455
   End
   Begin VB.CommandButton btnCancel 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   3240
      TabIndex        =   27
      Top             =   8280
      Width           =   1455
   End
   Begin VB.Label lblDeathLocation 
      Caption         =   "Death Location:"
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
      Left            =   240
      TabIndex        =   71
      Top             =   6000
      Width           =   1335
   End
   Begin VB.Label Label17 
      Caption         =   "Map:"
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
      Left            =   1920
      TabIndex        =   70
      Top             =   6000
      Width           =   735
   End
   Begin VB.Label Label16 
      Caption         =   "X:"
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
      Left            =   3840
      TabIndex        =   69
      Top             =   6000
      Width           =   375
   End
   Begin VB.Label Label15 
      Caption         =   "Y:"
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
      Left            =   5520
      TabIndex        =   68
      Top             =   6000
      Width           =   255
   End
   Begin VB.Label lblRate 
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
      Index           =   9
      Left            =   6240
      TabIndex        =   64
      Top             =   3600
      Width           =   495
   End
   Begin VB.Label lblRate 
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
      Index           =   8
      Left            =   9120
      TabIndex        =   61
      Top             =   2880
      Width           =   495
   End
   Begin VB.Label lblRate 
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
      Index           =   7
      Left            =   6240
      TabIndex        =   59
      Top             =   2880
      Width           =   495
   End
   Begin VB.Label lblRate 
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
      Index           =   6
      Left            =   3360
      TabIndex        =   57
      Top             =   2880
      Width           =   495
   End
   Begin VB.Label lblRate 
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
      Index           =   5
      Left            =   9120
      TabIndex        =   55
      Top             =   2160
      Width           =   495
   End
   Begin VB.Label lblRate 
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
      Index           =   4
      Left            =   6240
      TabIndex        =   53
      Top             =   2160
      Width           =   495
   End
   Begin VB.Label lblRate 
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
      Index           =   3
      Left            =   3360
      TabIndex        =   51
      Top             =   2160
      Width           =   495
   End
   Begin VB.Label Label14 
      Caption         =   "Y:"
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
      Left            =   5520
      TabIndex        =   48
      Top             =   5520
      Width           =   255
   End
   Begin VB.Label Label13 
      Caption         =   "X:"
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
      Left            =   3840
      TabIndex        =   47
      Top             =   5520
      Width           =   375
   End
   Begin VB.Label Label12 
      Caption         =   "Map:"
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
      Left            =   1920
      TabIndex        =   46
      Top             =   5520
      Width           =   735
   End
   Begin VB.Label Label11 
      Caption         =   "NPC:"
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
      Left            =   240
      TabIndex        =   45
      Top             =   3480
      Width           =   855
   End
   Begin VB.Label Label10 
      Caption         =   "Boot Location:"
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
      Left            =   240
      TabIndex        =   44
      Top             =   5520
      Width           =   1335
   End
   Begin VB.Label Label9 
      Caption         =   "Exits:"
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
      Left            =   240
      TabIndex        =   41
      Top             =   4320
      Width           =   855
   End
   Begin VB.Label Label8 
      Caption         =   "Up:"
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
      Left            =   1200
      TabIndex        =   40
      Top             =   4320
      Width           =   735
   End
   Begin VB.Label Label7 
      Caption         =   "Down:"
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
      Left            =   3360
      TabIndex        =   39
      Top             =   4320
      Width           =   735
   End
   Begin VB.Label Label5 
      Caption         =   "Left:"
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
      Left            =   1200
      TabIndex        =   38
      Top             =   4920
      Width           =   735
   End
   Begin VB.Label Label6 
      Caption         =   "Right:"
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
      Left            =   3360
      TabIndex        =   37
      Top             =   4800
      Width           =   735
   End
   Begin VB.Label lblMidi 
      Alignment       =   2  'Center
      Caption         =   "<None>"
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
      Left            =   5160
      TabIndex        =   36
      Top             =   600
      Width           =   1575
   End
   Begin VB.Label Label4 
      Caption         =   "Midi:"
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
      Left            =   120
      TabIndex        =   35
      Top             =   600
      Width           =   855
   End
   Begin VB.Label lblRate 
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
      Index           =   2
      Left            =   9120
      TabIndex        =   34
      Top             =   1440
      Width           =   495
   End
   Begin VB.Label lblRate 
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
      Index           =   1
      Left            =   6240
      TabIndex        =   33
      Top             =   1440
      Width           =   495
   End
   Begin VB.Label lblRate 
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
      Index           =   0
      Left            =   3360
      TabIndex        =   32
      Top             =   1440
      Width           =   495
   End
   Begin VB.Label Label3 
      Caption         =   "Monsters:"
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
      Left            =   120
      TabIndex        =   31
      Top             =   1080
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Flags:"
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
      Left            =   240
      TabIndex        =   30
      Top             =   6840
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Name:"
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
      Left            =   120
      TabIndex        =   29
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "frmMapProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub btnCancel_Click()
    Me.Hide
End Sub

Private Sub btnOk_Click()
    Dim A As Long
    With EditMap
        .name = txtName
        .MIDI = sclMIDI
        .ExitUp = Int(Val(txtUp))
        .ExitDown = Int(Val(txtDown))
        .ExitLeft = Int(Val(txtLeft))
        .ExitRight = Int(Val(txtRight))
        .BootLocation.Map = Int(Val(txtBootMap))
        .BootLocation.X = Int(Val(txtBootX))
        .BootLocation.Y = Int(Val(txtBootY))
        .DeathLocation.Map = Int(Val(txtDeathMap))
        .DeathLocation.X = Int(Val(txtDeathX))
        .DeathLocation.Y = Int(Val(txtDeathY))
        .NPC = cmbNPC.ListIndex
        For A = 0 To 9
            .MonsterSpawn(A).Monster = cmbMonster(A).ListIndex
            .MonsterSpawn(A).Rate = sclRate(A)
        Next A
        For A = 0 To 7
            If chkFlag(A) = 1 Then
                SetBit .flags, CByte(A)
            Else
                ClearBit .flags, CByte(A)
            End If
            If chkFlag2(A) = 1 Then
                SetBit .Flags2, CByte(A)
            Else
                ClearBit .Flags2, CByte(A)
            End If
        Next A
    End With
    
    Me.Hide
End Sub

Private Sub Form_Load()
    Dim A As Long, B As Long
    
    For A = 0 To 9
        cmbMonster(A).AddItem "<None>"
    Next A
    cmbNPC.AddItem "<None>"

    For A = 1 To MaxNPCs
        cmbNPC.AddItem CStr(A) + ": " + NPC(A).name
    Next A
    
    For A = 1 To MaxTotalMonsters
        For B = 0 To 9
            cmbMonster(B).AddItem CStr(A) + ": " + Monster(A).name
        Next B
    Next A

    frmMapProperties_Loaded = True
End Sub
Private Sub Form_Unload(Cancel As Integer)
    frmMapProperties_Loaded = False
End Sub


Private Sub sclMIDI_Change()
    If sclMIDI = 0 Then
        lblMidi = "<None>"
    Else
        lblMidi = sclMIDI
    End If
End Sub
Private Sub sclMIDI_Scroll()
    sclMIDI_Change
End Sub


Private Sub sclRate_Change(index As Integer)
    lblRate(index) = sclRate(index)
End Sub


Private Sub sclRate_Scroll(index As Integer)
    sclRate_Change (index)
End Sub

Private Sub txtBootMap_LostFocus()
    Dim A As Double
    A = Int(Val(txtBootMap))
    If A > MaxMaps Then A = MaxMaps
    If A < 0 Then A = 0
    txtBootMap = CStr(A)
End Sub


Private Sub txtBootX_LostFocus()
    Dim A As Double
    A = Int(Val(txtBootX))
    If A > 11 Then A = 11
    If A < 0 Then A = 0
    txtBootX = CStr(A)
End Sub


Private Sub txtBootY_LostFocus()
    Dim A As Double
    A = Int(Val(txtBootY))
    If A > 11 Then A = 11
    If A < 0 Then A = 0
    txtBootY = CStr(A)
End Sub

Private Sub txtDeathMap_LostFocus()
    Dim A As Double
    A = Int(Val(txtDeathMap))
    If A > MaxMaps Then A = MaxMaps
    If A < 0 Then A = 0
    txtDeathMap = CStr(A)
End Sub


Private Sub txtDeathX_LostFocus()
    Dim A As Double
    A = Int(Val(txtDeathX))
    If A > 11 Then A = 11
    If A < 0 Then A = 0
    txtDeathX = CStr(A)
End Sub


Private Sub txtDeathY_LostFocus()
    Dim A As Double
    A = Int(Val(txtDeathY))
    If A > 11 Then A = 11
    If A < 0 Then A = 0
    txtDeathY = CStr(A)
End Sub


Private Sub txtDown_LostFocus()
    Dim A As Double
    A = Int(Val(txtDown))
    If A > MaxMaps Then A = MaxMaps
    If A < 0 Then A = 0
    txtDown = CStr(A)
End Sub


Private Sub txtLeft_LostFocus()
    Dim A As Double
    A = Int(Val(txtLeft))
    If A > MaxMaps Then A = MaxMaps
    If A < 0 Then A = 0
    txtLeft = CStr(A)
End Sub


Private Sub txtRight_LostFocus()
    Dim A As Double
    A = Int(Val(txtRight))
    If A > MaxMaps Then A = MaxMaps
    If A < 0 Then A = 0
    txtRight = CStr(A)
End Sub


Private Sub txtUp_LostFocus()
    Dim A As Double
    A = Int(Val(txtUp))
    If A > MaxMaps Then A = MaxMaps
    If A < 0 Then A = 0
    txtUp = CStr(A)
End Sub
