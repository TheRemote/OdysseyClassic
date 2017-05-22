VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H0044342E&
   BorderStyle     =   0  'None
   Caption         =   "Odyssey Classic"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   15
   ClientWidth     =   12000
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picSellObject 
      Appearance      =   0  'Flat
      BackColor       =   &H0061514B&
      ForeColor       =   &H80000008&
      Height          =   3495
      Left            =   480
      ScaleHeight     =   3465
      ScaleWidth      =   5025
      TabIndex        =   98
      Top             =   1440
      Visible         =   0   'False
      Width           =   5055
      Begin VB.Frame fraSellObject 
         BackColor       =   &H0044342E&
         Caption         =   "Sell Object"
         ForeColor       =   &H009AADC2&
         Height          =   2655
         Left            =   120
         TabIndex        =   99
         Top             =   120
         Width           =   2535
         Begin VB.PictureBox picSellObjectDisp 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H0044342E&
            BorderStyle     =   0  'None
            ForeColor       =   &H009AADC2&
            Height          =   480
            Left            =   1080
            ScaleHeight     =   32
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   32
            TabIndex        =   100
            Top             =   360
            Width           =   480
         End
         Begin VB.Label lblSellPrice 
            BackColor       =   &H0044342E&
            ForeColor       =   &H009AADC2&
            Height          =   255
            Left            =   720
            TabIndex        =   104
            Top             =   1440
            Width           =   1695
         End
         Begin VB.Label cptSellPrice 
            AutoSize        =   -1  'True
            BackColor       =   &H0044342E&
            Caption         =   "Price:"
            ForeColor       =   &H009AADC2&
            Height          =   195
            Left            =   120
            TabIndex        =   103
            Top             =   1440
            Width           =   405
         End
         Begin VB.Label lblSellName 
            BackColor       =   &H0044342E&
            ForeColor       =   &H009AADC2&
            Height          =   255
            Left            =   720
            TabIndex        =   102
            Top             =   1080
            Width           =   1695
         End
         Begin VB.Label cptSellName 
            AutoSize        =   -1  'True
            BackColor       =   &H0044342E&
            Caption         =   "Name:"
            ForeColor       =   &H009AADC2&
            Height          =   195
            Left            =   120
            TabIndex        =   101
            Top             =   1080
            Width           =   465
         End
      End
      Begin VB.Label lblMenu 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0044342E&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Sell All"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H009AADC2&
         Height          =   255
         Index           =   59
         Left            =   1920
         TabIndex        =   109
         Top             =   3000
         Width           =   1215
      End
      Begin VB.Label lblSellNPCTalk 
         BackColor       =   &H0044342E&
         ForeColor       =   &H009AADC2&
         Height          =   2295
         Left            =   2760
         TabIndex        =   106
         Top             =   480
         Width           =   2175
      End
      Begin VB.Label lblSellNPCName 
         Alignment       =   2  'Center
         BackColor       =   &H0044342E&
         ForeColor       =   &H009AADC2&
         Height          =   255
         Left            =   2760
         TabIndex        =   105
         Top             =   120
         Width           =   2175
      End
      Begin VB.Label lblMenu 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0044342E&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H009AADC2&
         Height          =   255
         Index           =   22
         Left            =   3600
         TabIndex        =   108
         Top             =   3000
         Width           =   735
      End
      Begin VB.Label lblMenu 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0044342E&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Sell"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H009AADC2&
         Height          =   255
         Index           =   19
         Left            =   720
         TabIndex        =   107
         Top             =   3000
         Width           =   735
      End
   End
   Begin VB.PictureBox picRepair 
      Appearance      =   0  'Flat
      BackColor       =   &H0061514B&
      ForeColor       =   &H80000008&
      Height          =   3495
      Left            =   480
      ScaleHeight     =   3465
      ScaleWidth      =   5025
      TabIndex        =   82
      Top             =   1440
      Visible         =   0   'False
      Width           =   5055
      Begin VB.Frame fraRepairObj 
         BackColor       =   &H0044342E&
         Caption         =   "Repair Object"
         ForeColor       =   &H009AADC2&
         Height          =   2655
         Left            =   120
         TabIndex        =   83
         Top             =   120
         Width           =   2535
         Begin VB.PictureBox picRepairObjectDisp 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H0044342E&
            BorderStyle     =   0  'None
            ForeColor       =   &H009AADC2&
            Height          =   480
            Left            =   1440
            ScaleHeight     =   32
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   32
            TabIndex        =   84
            Top             =   360
            Width           =   480
         End
         Begin VB.Label cptRepairName 
            AutoSize        =   -1  'True
            BackColor       =   &H0044342E&
            Caption         =   "Name:"
            ForeColor       =   &H009AADC2&
            Height          =   195
            Left            =   120
            TabIndex        =   85
            Top             =   1080
            Width           =   465
         End
         Begin VB.Label lblRepairName 
            Alignment       =   2  'Center
            BackColor       =   &H0044342E&
            ForeColor       =   &H009AADC2&
            Height          =   255
            Left            =   840
            TabIndex        =   86
            Top             =   1080
            Width           =   1575
         End
         Begin VB.Label cptRepairDurability 
            AutoSize        =   -1  'True
            BackColor       =   &H0044342E&
            Caption         =   "Durability:"
            ForeColor       =   &H009AADC2&
            Height          =   195
            Left            =   120
            TabIndex        =   87
            Top             =   1440
            Width           =   690
         End
         Begin VB.Label lblRepairDurability 
            Alignment       =   2  'Center
            BackColor       =   &H0044342E&
            ForeColor       =   &H009AADC2&
            Height          =   255
            Left            =   840
            TabIndex        =   88
            Top             =   1440
            Width           =   1575
         End
         Begin VB.Label cptRepairCost 
            AutoSize        =   -1  'True
            BackColor       =   &H0044342E&
            Caption         =   "Cost:"
            ForeColor       =   &H009AADC2&
            Height          =   195
            Left            =   120
            TabIndex        =   89
            Top             =   1800
            Width           =   360
         End
         Begin VB.Label lblRepairCost 
            Alignment       =   2  'Center
            BackColor       =   &H0044342E&
            ForeColor       =   &H009AADC2&
            Height          =   255
            Left            =   840
            TabIndex        =   90
            Top             =   1800
            Width           =   1575
         End
         Begin VB.Label cptRepairCondition 
            AutoSize        =   -1  'True
            BackColor       =   &H0044342E&
            Caption         =   "Condition:"
            ForeColor       =   &H009AADC2&
            Height          =   195
            Left            =   120
            TabIndex        =   91
            Top             =   2160
            Width           =   705
         End
         Begin VB.Label lblRepairCondition 
            Alignment       =   2  'Center
            BackColor       =   &H0044342E&
            ForeColor       =   &H009AADC2&
            Height          =   255
            Left            =   840
            TabIndex        =   92
            Top             =   2160
            Width           =   1575
         End
      End
      Begin VB.Label lblMenu 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0044342E&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Repair"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H009AADC2&
         Height          =   255
         Index           =   56
         Left            =   720
         TabIndex        =   96
         Top             =   3000
         Width           =   735
      End
      Begin VB.Label lblMenu 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0044342E&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H009AADC2&
         Height          =   255
         Index           =   57
         Left            =   3600
         TabIndex        =   97
         Top             =   3000
         Width           =   735
      End
      Begin VB.Label lblRepairNPCName 
         Alignment       =   2  'Center
         BackColor       =   &H0044342E&
         ForeColor       =   &H009AADC2&
         Height          =   255
         Left            =   2760
         TabIndex        =   93
         Top             =   120
         Width           =   2175
      End
      Begin VB.Label lblRepairNpcTalk 
         BackColor       =   &H0044342E&
         ForeColor       =   &H009AADC2&
         Height          =   2295
         Left            =   2760
         TabIndex        =   94
         Top             =   480
         Width           =   2175
      End
      Begin VB.Label lblMenu 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0044342E&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Repair All"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H009AADC2&
         Height          =   255
         Index           =   16
         Left            =   1920
         TabIndex        =   95
         Top             =   3000
         Width           =   1215
      End
   End
   Begin VB.PictureBox picButtonLight 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   105
      Index           =   10
      Left            =   10635
      ScaleHeight     =   105
      ScaleWidth      =   105
      TabIndex        =   169
      Top             =   3855
      Width           =   105
   End
   Begin VB.PictureBox picButtonLight 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   105
      Index           =   9
      Left            =   10635
      ScaleHeight     =   105
      ScaleWidth      =   105
      TabIndex        =   165
      Top             =   3495
      Width           =   105
   End
   Begin VB.PictureBox picButtonLight 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   105
      Index           =   8
      Left            =   10635
      ScaleHeight     =   105
      ScaleWidth      =   105
      TabIndex        =   163
      Top             =   3135
      Width           =   105
   End
   Begin VB.PictureBox picButtonLight 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   105
      Index           =   7
      Left            =   10635
      ScaleHeight     =   105
      ScaleWidth      =   105
      TabIndex        =   152
      Top             =   2775
      Width           =   105
   End
   Begin VB.PictureBox picButtonLight 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   105
      Index           =   6
      Left            =   10635
      ScaleHeight     =   105
      ScaleWidth      =   105
      TabIndex        =   150
      Top             =   2280
      Width           =   105
   End
   Begin VB.PictureBox picButtonLight 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   105
      Index           =   5
      Left            =   10635
      ScaleHeight     =   105
      ScaleWidth      =   105
      TabIndex        =   113
      Top             =   1920
      Width           =   105
   End
   Begin VB.PictureBox picButtonLight 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   105
      Index           =   4
      Left            =   10635
      ScaleHeight     =   105
      ScaleWidth      =   105
      TabIndex        =   111
      Top             =   1560
      Width           =   105
   End
   Begin VB.PictureBox picButtonLight 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   105
      Index           =   3
      Left            =   10635
      ScaleHeight     =   105
      ScaleWidth      =   105
      TabIndex        =   81
      Top             =   1200
      Width           =   105
   End
   Begin VB.PictureBox picButtonLight 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   105
      Index           =   2
      Left            =   10635
      ScaleHeight     =   105
      ScaleWidth      =   105
      TabIndex        =   79
      Top             =   840
      Width           =   105
   End
   Begin VB.PictureBox picButtonLight 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   105
      Index           =   1
      Left            =   10635
      ScaleHeight     =   105
      ScaleWidth      =   105
      TabIndex        =   77
      Top             =   480
      Width           =   105
   End
   Begin VB.PictureBox picButtonLight 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   105
      Index           =   0
      Left            =   10635
      ScaleHeight     =   105
      ScaleWidth      =   105
      TabIndex        =   0
      Top             =   120
      Width           =   105
   End
   Begin MSComDlg.CommonDialog CommonDialog 
      Left            =   240
      Top             =   8400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox picMapEdit 
      Appearance      =   0  'Flat
      BackColor       =   &H0061514B&
      ForeColor       =   &H80000008&
      Height          =   6105
      Left            =   5880
      ScaleHeight     =   6075
      ScaleWidth      =   3705
      TabIndex        =   119
      Top             =   120
      Visible         =   0   'False
      Width           =   3735
      Begin VB.PictureBox picTiles 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   2940
         Left            =   75
         ScaleHeight     =   194
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   230
         TabIndex        =   121
         Top             =   360
         Width           =   3480
      End
      Begin VB.PictureBox picTile 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   510
         Left            =   600
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   124
         Top             =   3710
         Width           =   510
      End
      Begin VB.VScrollBar MapScroll 
         Height          =   2940
         Left            =   3550
         TabIndex        =   122
         Top             =   390
         Width           =   135
      End
      Begin VB.Label lblEditMode 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0044342E&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "FGTile2"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H009AADC2&
         Height          =   375
         Index           =   5
         Left            =   2680
         TabIndex        =   187
         Top             =   4845
         Width           =   975
      End
      Begin VB.Label lblMenu 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0044342E&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Save"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H009AADC2&
         Height          =   375
         Index           =   60
         Left            =   50
         TabIndex        =   186
         Top             =   4850
         Width           =   520
      End
      Begin VB.Label lblEditMode 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0044342E&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Att2"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H009AADC2&
         Height          =   375
         Index           =   7
         Left            =   2140
         TabIndex        =   137
         Top             =   5280
         Width           =   480
      End
      Begin VB.Label lblEditMode 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0044342E&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Ground2"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H009AADC2&
         Height          =   375
         Index           =   1
         Left            =   1635
         TabIndex        =   132
         Top             =   4440
         Width           =   975
      End
      Begin VB.Label lblCurTile 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H009AADC2&
         Height          =   300
         Left            =   120
         TabIndex        =   129
         Top             =   3840
         Width           =   480
      End
      Begin VB.Label lblMenu 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0044342E&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Load"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H009AADC2&
         Height          =   375
         Index           =   47
         Left            =   50
         TabIndex        =   126
         Top             =   4440
         Width           =   520
      End
      Begin VB.Label lblEditMode 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0044342E&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Sand"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H009AADC2&
         Height          =   210
         Index           =   9
         Left            =   3240
         TabIndex        =   140
         Top             =   5760
         Width           =   495
      End
      Begin VB.Label lblEditMode 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0044342E&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Trees"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H009AADC2&
         Height          =   210
         Index           =   10
         Left            =   3240
         TabIndex        =   141
         Top             =   5520
         Width           =   495
      End
      Begin VB.Label lblMenu 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0044342E&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Up"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H009AADC2&
         Height          =   330
         Index           =   2
         Left            =   75
         TabIndex        =   120
         Top             =   30
         Width           =   3480
      End
      Begin VB.Label lblMenu 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0044342E&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Down"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H009AADC2&
         Height          =   330
         Index           =   3
         Left            =   75
         TabIndex        =   123
         Top             =   3360
         Width           =   3590
      End
      Begin VB.Label lblMenu 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0044342E&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Upload"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H009AADC2&
         Height          =   330
         Index           =   4
         Left            =   2865
         TabIndex        =   125
         Top             =   3720
         Width           =   795
      End
      Begin VB.Label lblEditMode 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0044342E&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Ground"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H009AADC2&
         Height          =   375
         Index           =   0
         Left            =   600
         TabIndex        =   131
         Top             =   4440
         Width           =   975
      End
      Begin VB.Label lblEditMode 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0044342E&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "BGTile1"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H009AADC2&
         Height          =   375
         Index           =   2
         Left            =   2680
         TabIndex        =   133
         Top             =   4440
         Width           =   975
      End
      Begin VB.Label lblEditMode 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0044342E&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "BGTile2"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H009AADC2&
         Height          =   375
         Index           =   3
         Left            =   600
         TabIndex        =   134
         Top             =   4845
         Width           =   975
      End
      Begin VB.Label lblEditMode 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0044342E&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "FGTile"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H009AADC2&
         Height          =   375
         Index           =   4
         Left            =   1635
         TabIndex        =   135
         Top             =   4845
         Width           =   975
      End
      Begin VB.Label lblEditMode 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0044342E&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Att1"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H009AADC2&
         Height          =   375
         Index           =   6
         Left            =   1635
         TabIndex        =   136
         Top             =   5280
         Width           =   480
      End
      Begin VB.Label lblMenu 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0044342E&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Properties"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H009AADC2&
         Height          =   375
         Index           =   58
         Left            =   75
         TabIndex        =   142
         Top             =   5400
         Width           =   1110
      End
      Begin VB.Label lblMenu 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0044342E&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H009AADC2&
         Height          =   330
         Index           =   5
         Left            =   1125
         TabIndex        =   130
         Top             =   4080
         Width           =   2535
      End
      Begin VB.Label lblMenu 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0044342E&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Copy"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H009AADC2&
         Height          =   330
         Index           =   54
         Left            =   1130
         TabIndex        =   127
         Top             =   3720
         Width           =   555
      End
      Begin VB.Label lblMenu 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0044342E&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Paste"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H009AADC2&
         Height          =   330
         Index           =   55
         Left            =   1740
         TabIndex        =   128
         Top             =   3720
         Width           =   615
      End
      Begin VB.Label lblEditMode 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0044342E&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Grass"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H009AADC2&
         Height          =   210
         Index           =   11
         Left            =   3240
         TabIndex        =   138
         Top             =   5280
         Width           =   495
      End
      Begin VB.Label lblEditMode 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0044342E&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Fill"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H009AADC2&
         Height          =   210
         Index           =   8
         Left            =   600
         TabIndex        =   139
         Top             =   4200
         Width           =   510
      End
   End
   Begin VB.PictureBox picDrop 
      Appearance      =   0  'Flat
      BackColor       =   &H0061514B&
      ForeColor       =   &H80000008&
      Height          =   1320
      Left            =   1320
      ScaleHeight     =   1290
      ScaleWidth      =   3225
      TabIndex        =   114
      Top             =   2160
      Visible         =   0   'False
      Width           =   3255
      Begin VB.TextBox txtDrop 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0044342E&
         ForeColor       =   &H009AADC2&
         Height          =   285
         Left            =   120
         TabIndex        =   116
         Top             =   480
         Width           =   3015
      End
      Begin VB.Label lblDrop 
         Alignment       =   2  'Center
         BackColor       =   &H0044342E&
         Caption         =   "Drop how much?"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H009AADC2&
         Height          =   255
         Left            =   120
         TabIndex        =   115
         Top             =   120
         Width           =   3000
      End
      Begin VB.Label lblMenu 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0044342E&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Ok"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H009AADC2&
         Height          =   375
         Index           =   35
         Left            =   1680
         TabIndex        =   118
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label lblMenu 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0044342E&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H009AADC2&
         Height          =   375
         Index           =   34
         Left            =   120
         TabIndex        =   117
         Top             =   840
         Width           =   1455
      End
   End
   Begin VB.PictureBox picHelp 
      Appearance      =   0  'Flat
      BackColor       =   &H0061514B&
      ForeColor       =   &H80000008&
      Height          =   3345
      Left            =   8160
      ScaleHeight     =   3315
      ScaleWidth      =   3750
      TabIndex        =   173
      Top             =   5520
      Visible         =   0   'False
      Width           =   3780
      Begin VB.Label lblDropObject 
         Alignment       =   2  'Center
         BackColor       =   &H0044342E&
         Caption         =   "Drop Object - Right Click"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H009AADC2&
         Height          =   225
         Left            =   120
         TabIndex        =   180
         Top             =   2040
         Width           =   3495
      End
      Begin VB.Label lblRun 
         Alignment       =   2  'Center
         BackColor       =   &H0044342E&
         Caption         =   "Run - Hold Shift"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H009AADC2&
         Height          =   225
         Left            =   120
         TabIndex        =   177
         Top             =   1320
         Width           =   3495
      End
      Begin VB.Label lblUseObject 
         Alignment       =   2  'Center
         BackColor       =   &H0044342E&
         Caption         =   "Use Object - Double Click"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H009AADC2&
         Height          =   225
         Left            =   120
         TabIndex        =   181
         Top             =   2280
         Width           =   3495
      End
      Begin VB.Label lblPickUp 
         Alignment       =   2  'Center
         BackColor       =   &H0044342E&
         Caption         =   "Pick Up Object - Enter"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H009AADC2&
         Height          =   225
         Left            =   120
         TabIndex        =   179
         Top             =   1800
         Width           =   3495
      End
      Begin VB.Label lblBroadcast 
         Alignment       =   2  'Center
         BackColor       =   &H0044342E&
         Caption         =   "If you need more help, ask other players  by typing /broadcast <message>"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H009AADC2&
         Height          =   465
         Left            =   120
         TabIndex        =   182
         Top             =   2640
         Width           =   3495
      End
      Begin VB.Label lblAttacking 
         Alignment       =   2  'Center
         BackColor       =   &H0044342E&
         Caption         =   "Attack - Ctrl"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H009AADC2&
         Height          =   225
         Left            =   120
         TabIndex        =   178
         Top             =   1560
         Width           =   3495
      End
      Begin VB.Label lblMovement 
         Alignment       =   2  'Center
         BackColor       =   &H0044342E&
         Caption         =   "Move - Arrow Keys"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H009AADC2&
         Height          =   210
         Left            =   120
         TabIndex        =   176
         Top             =   1080
         Width           =   3495
      End
      Begin VB.Label lblInstructions 
         Alignment       =   2  'Center
         BackColor       =   &H0044342E&
         Caption         =   "Welcome to the Odyssey Online Classic!  This is a basic guide to the game."
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H009AADC2&
         Height          =   540
         Left            =   120
         TabIndex        =   175
         Top             =   480
         Width           =   3495
      End
      Begin VB.Label lblWelcome 
         Alignment       =   2  'Center
         BackColor       =   &H0044342E&
         Caption         =   "Welcome Adventurer!"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H009AADC2&
         Height          =   300
         Left            =   120
         TabIndex        =   174
         Top             =   120
         Width           =   3495
      End
      Begin VB.Label lblMenu 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0044342E&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Close"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H009AADC2&
         Height          =   255
         Index           =   20
         Left            =   3000
         TabIndex        =   183
         Top             =   3120
         Width           =   615
      End
   End
   Begin VB.PictureBox picBuy 
      Appearance      =   0  'Flat
      BackColor       =   &H0061514B&
      ForeColor       =   &H80000008&
      Height          =   7215
      Left            =   120
      ScaleHeight     =   7185
      ScaleWidth      =   8265
      TabIndex        =   7
      Top             =   120
      Visible         =   0   'False
      Width           =   8295
      Begin VB.PictureBox GivObjPic 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H0044342E&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   9
         Left            =   120
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   36
         Top             =   6000
         Width           =   480
      End
      Begin VB.PictureBox GivObjPic 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H0044342E&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   8
         Left            =   120
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   33
         Top             =   5400
         Width           =   480
      End
      Begin VB.PictureBox GivObjPic 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H0044342E&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   7
         Left            =   120
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   30
         Top             =   4800
         Width           =   480
      End
      Begin VB.PictureBox GivObjPic 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H0044342E&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   6
         Left            =   120
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   27
         Top             =   4200
         Width           =   480
      End
      Begin VB.PictureBox GivObjPic 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H0044342E&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   5
         Left            =   120
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   24
         Top             =   3600
         Width           =   480
      End
      Begin VB.PictureBox GivObjPic 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H0044342E&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   4
         Left            =   120
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   21
         Top             =   3000
         Width           =   480
      End
      Begin VB.PictureBox GivObjPic 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H0044342E&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   3
         Left            =   120
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   18
         Top             =   2400
         Width           =   480
      End
      Begin VB.PictureBox GivObjPic 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H0044342E&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   2
         Left            =   120
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   15
         Top             =   1800
         Width           =   480
      End
      Begin VB.PictureBox GivObjPic 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H0044342E&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   1
         Left            =   120
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   12
         Top             =   1200
         Width           =   480
      End
      Begin VB.PictureBox GivObjPic 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H0044342E&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   0
         Left            =   120
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   9
         Top             =   600
         Width           =   480
      End
      Begin VB.Label lblMenu 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0044342E&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Close"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H009AADC2&
         Height          =   375
         Index           =   46
         Left            =   120
         TabIndex        =   39
         Top             =   6600
         Width           =   8055
      End
      Begin VB.Label lblItem 
         Alignment       =   2  'Center
         BackColor       =   &H0044342E&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H009AADC2&
         Height          =   495
         Index           =   9
         Left            =   720
         TabIndex        =   37
         Top             =   6000
         Width           =   6495
      End
      Begin VB.Label lblMenu 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0044342E&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Trade"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H009AADC2&
         Height          =   255
         Index           =   45
         Left            =   7320
         TabIndex        =   38
         Top             =   6000
         Width           =   855
      End
      Begin VB.Label lblItem 
         Alignment       =   2  'Center
         BackColor       =   &H0044342E&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H009AADC2&
         Height          =   495
         Index           =   8
         Left            =   720
         TabIndex        =   34
         Top             =   5400
         Width           =   6495
      End
      Begin VB.Label lblMenu 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0044342E&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Trade"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H009AADC2&
         Height          =   255
         Index           =   44
         Left            =   7320
         TabIndex        =   35
         Top             =   5400
         Width           =   855
      End
      Begin VB.Label lblItem 
         Alignment       =   2  'Center
         BackColor       =   &H0044342E&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H009AADC2&
         Height          =   495
         Index           =   7
         Left            =   720
         TabIndex        =   31
         Top             =   4800
         Width           =   6495
      End
      Begin VB.Label lblMenu 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0044342E&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Trade"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H009AADC2&
         Height          =   255
         Index           =   43
         Left            =   7320
         TabIndex        =   32
         Top             =   4800
         Width           =   855
      End
      Begin VB.Label lblItem 
         Alignment       =   2  'Center
         BackColor       =   &H0044342E&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H009AADC2&
         Height          =   495
         Index           =   6
         Left            =   720
         TabIndex        =   28
         Top             =   4200
         Width           =   6495
      End
      Begin VB.Label lblMenu 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0044342E&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Trade"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H009AADC2&
         Height          =   255
         Index           =   42
         Left            =   7320
         TabIndex        =   29
         Top             =   4200
         Width           =   855
      End
      Begin VB.Label lblItem 
         Alignment       =   2  'Center
         BackColor       =   &H0044342E&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H009AADC2&
         Height          =   495
         Index           =   5
         Left            =   720
         TabIndex        =   25
         Top             =   3600
         Width           =   6495
      End
      Begin VB.Label lblMenu 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0044342E&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Trade"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H009AADC2&
         Height          =   255
         Index           =   41
         Left            =   7320
         TabIndex        =   26
         Top             =   3600
         Width           =   855
      End
      Begin VB.Label lblItem 
         Alignment       =   2  'Center
         BackColor       =   &H0044342E&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H009AADC2&
         Height          =   495
         Index           =   4
         Left            =   720
         TabIndex        =   22
         Top             =   3000
         Width           =   6495
      End
      Begin VB.Label lblMenu 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0044342E&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Trade"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H009AADC2&
         Height          =   255
         Index           =   40
         Left            =   7320
         TabIndex        =   23
         Top             =   3000
         Width           =   855
      End
      Begin VB.Label lblItem 
         Alignment       =   2  'Center
         BackColor       =   &H0044342E&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H009AADC2&
         Height          =   495
         Index           =   3
         Left            =   720
         TabIndex        =   19
         Top             =   2400
         Width           =   6495
      End
      Begin VB.Label lblMenu 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0044342E&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Trade"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H009AADC2&
         Height          =   255
         Index           =   39
         Left            =   7320
         TabIndex        =   20
         Top             =   2400
         Width           =   855
      End
      Begin VB.Label lblItem 
         Alignment       =   2  'Center
         BackColor       =   &H0044342E&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H009AADC2&
         Height          =   495
         Index           =   2
         Left            =   720
         TabIndex        =   16
         Top             =   1800
         Width           =   6495
      End
      Begin VB.Label lblMenu 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0044342E&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Trade"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H009AADC2&
         Height          =   255
         Index           =   38
         Left            =   7320
         TabIndex        =   17
         Top             =   1800
         Width           =   855
      End
      Begin VB.Label lblItem 
         Alignment       =   2  'Center
         BackColor       =   &H0044342E&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H009AADC2&
         Height          =   495
         Index           =   1
         Left            =   720
         TabIndex        =   13
         Top             =   1200
         Width           =   6495
      End
      Begin VB.Label lblMenu 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0044342E&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Trade"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H009AADC2&
         Height          =   255
         Index           =   37
         Left            =   7320
         TabIndex        =   14
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label lblItem 
         Alignment       =   2  'Center
         BackColor       =   &H0044342E&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H009AADC2&
         Height          =   495
         Index           =   0
         Left            =   720
         TabIndex        =   10
         Top             =   600
         Width           =   6495
      End
      Begin VB.Label lblMenu 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0044342E&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Trade"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H009AADC2&
         Height          =   255
         Index           =   36
         Left            =   7320
         TabIndex        =   11
         Top             =   600
         Width           =   855
      End
      Begin VB.Label lblShopName 
         Alignment       =   2  'Center
         BackColor       =   &H0044342E&
         Caption         =   "[]"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H009AADC2&
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   120
         Width           =   8055
      End
   End
   Begin VB.PictureBox picBank 
      Appearance      =   0  'Flat
      BackColor       =   &H0061514B&
      ForeColor       =   &H80000008&
      Height          =   4815
      Left            =   1440
      ScaleHeight     =   4785
      ScaleWidth      =   3105
      TabIndex        =   41
      Top             =   480
      Visible         =   0   'False
      Width           =   3135
      Begin VB.PictureBox ItemBank 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H0044342E&
         BorderStyle     =   0  'None
         ForeColor       =   &H009AADC2&
         Height          =   480
         Index           =   0
         Left            =   120
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   43
         Top             =   480
         Width           =   480
      End
      Begin VB.PictureBox ItemBank 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H0044342E&
         BorderStyle     =   0  'None
         ForeColor       =   &H009AADC2&
         Height          =   480
         Index           =   1
         Left            =   720
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   44
         Top             =   480
         Width           =   480
      End
      Begin VB.PictureBox ItemBank 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H0044342E&
         BorderStyle     =   0  'None
         ForeColor       =   &H009AADC2&
         Height          =   480
         Index           =   2
         Left            =   1320
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   45
         Top             =   480
         Width           =   480
      End
      Begin VB.PictureBox ItemBank 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H0044342E&
         BorderStyle     =   0  'None
         ForeColor       =   &H009AADC2&
         Height          =   480
         Index           =   3
         Left            =   1920
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   46
         Top             =   480
         Width           =   480
      End
      Begin VB.PictureBox ItemBank 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H0044342E&
         BorderStyle     =   0  'None
         ForeColor       =   &H009AADC2&
         Height          =   480
         Index           =   4
         Left            =   2520
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   47
         Top             =   480
         Width           =   480
      End
      Begin VB.PictureBox ItemBank 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H0044342E&
         BorderStyle     =   0  'None
         ForeColor       =   &H009AADC2&
         Height          =   480
         Index           =   5
         Left            =   120
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   48
         Top             =   1080
         Width           =   480
      End
      Begin VB.PictureBox ItemBank 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H0044342E&
         BorderStyle     =   0  'None
         ForeColor       =   &H009AADC2&
         Height          =   480
         Index           =   6
         Left            =   720
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   49
         Top             =   1080
         Width           =   480
      End
      Begin VB.PictureBox ItemBank 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H0044342E&
         BorderStyle     =   0  'None
         ForeColor       =   &H009AADC2&
         Height          =   480
         Index           =   7
         Left            =   1320
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   50
         Top             =   1080
         Width           =   480
      End
      Begin VB.PictureBox ItemBank 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H0044342E&
         BorderStyle     =   0  'None
         ForeColor       =   &H009AADC2&
         Height          =   480
         Index           =   8
         Left            =   1920
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   51
         Top             =   1080
         Width           =   480
      End
      Begin VB.PictureBox ItemBank 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H0044342E&
         BorderStyle     =   0  'None
         ForeColor       =   &H009AADC2&
         Height          =   480
         Index           =   9
         Left            =   2520
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   52
         Top             =   1080
         Width           =   480
      End
      Begin VB.PictureBox ItemBank 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H0044342E&
         BorderStyle     =   0  'None
         ForeColor       =   &H009AADC2&
         Height          =   480
         Index           =   10
         Left            =   120
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   53
         Top             =   1680
         Width           =   480
      End
      Begin VB.PictureBox ItemBank 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H0044342E&
         BorderStyle     =   0  'None
         ForeColor       =   &H009AADC2&
         Height          =   480
         Index           =   11
         Left            =   720
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   54
         Top             =   1680
         Width           =   480
      End
      Begin VB.PictureBox ItemBank 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H0044342E&
         BorderStyle     =   0  'None
         ForeColor       =   &H009AADC2&
         Height          =   480
         Index           =   12
         Left            =   1320
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   55
         Top             =   1680
         Width           =   480
      End
      Begin VB.PictureBox ItemBank 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H0044342E&
         BorderStyle     =   0  'None
         ForeColor       =   &H009AADC2&
         Height          =   480
         Index           =   13
         Left            =   1920
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   56
         Top             =   1680
         Width           =   480
      End
      Begin VB.PictureBox ItemBank 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H0044342E&
         BorderStyle     =   0  'None
         ForeColor       =   &H009AADC2&
         Height          =   480
         Index           =   14
         Left            =   2520
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   57
         Top             =   1680
         Width           =   480
      End
      Begin VB.PictureBox ItemBank 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H0044342E&
         BorderStyle     =   0  'None
         ForeColor       =   &H009AADC2&
         Height          =   480
         Index           =   15
         Left            =   120
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   58
         Top             =   2280
         Width           =   480
      End
      Begin VB.PictureBox ItemBank 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H0044342E&
         BorderStyle     =   0  'None
         ForeColor       =   &H009AADC2&
         Height          =   480
         Index           =   16
         Left            =   720
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   59
         Top             =   2280
         Width           =   480
      End
      Begin VB.PictureBox ItemBank 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H0044342E&
         BorderStyle     =   0  'None
         ForeColor       =   &H009AADC2&
         Height          =   480
         Index           =   17
         Left            =   1320
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   60
         Top             =   2280
         Width           =   480
      End
      Begin VB.PictureBox ItemBank 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H0044342E&
         BorderStyle     =   0  'None
         ForeColor       =   &H009AADC2&
         Height          =   480
         Index           =   18
         Left            =   1920
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   61
         Top             =   2280
         Width           =   480
      End
      Begin VB.PictureBox ItemBank 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H0044342E&
         BorderStyle     =   0  'None
         ForeColor       =   &H009AADC2&
         Height          =   480
         Index           =   19
         Left            =   2520
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   62
         Top             =   2280
         Width           =   480
      End
      Begin VB.PictureBox ItemBank 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H0044342E&
         BorderStyle     =   0  'None
         ForeColor       =   &H009AADC2&
         Height          =   480
         Index           =   20
         Left            =   120
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   63
         Top             =   2880
         Width           =   480
      End
      Begin VB.PictureBox ItemBank 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H0044342E&
         BorderStyle     =   0  'None
         ForeColor       =   &H009AADC2&
         Height          =   480
         Index           =   21
         Left            =   720
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   64
         Top             =   2880
         Width           =   480
      End
      Begin VB.PictureBox ItemBank 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H0044342E&
         BorderStyle     =   0  'None
         ForeColor       =   &H009AADC2&
         Height          =   480
         Index           =   22
         Left            =   1320
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   65
         Top             =   2880
         Width           =   480
      End
      Begin VB.PictureBox ItemBank 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H0044342E&
         BorderStyle     =   0  'None
         ForeColor       =   &H009AADC2&
         Height          =   480
         Index           =   23
         Left            =   1920
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   66
         Top             =   2880
         Width           =   480
      End
      Begin VB.PictureBox ItemBank 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H0044342E&
         BorderStyle     =   0  'None
         ForeColor       =   &H009AADC2&
         Height          =   480
         Index           =   24
         Left            =   2520
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   67
         Top             =   2880
         Width           =   480
      End
      Begin VB.PictureBox ItemBank 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H0044342E&
         BorderStyle     =   0  'None
         ForeColor       =   &H009AADC2&
         Height          =   480
         Index           =   25
         Left            =   120
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   68
         Top             =   3480
         Width           =   480
      End
      Begin VB.PictureBox ItemBank 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H0044342E&
         BorderStyle     =   0  'None
         ForeColor       =   &H009AADC2&
         Height          =   480
         Index           =   26
         Left            =   720
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   69
         Top             =   3480
         Width           =   480
      End
      Begin VB.PictureBox ItemBank 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H0044342E&
         BorderStyle     =   0  'None
         ForeColor       =   &H009AADC2&
         Height          =   480
         Index           =   27
         Left            =   1320
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   70
         Top             =   3480
         Width           =   480
      End
      Begin VB.PictureBox ItemBank 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H0044342E&
         BorderStyle     =   0  'None
         ForeColor       =   &H009AADC2&
         Height          =   480
         Index           =   28
         Left            =   1920
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   71
         Top             =   3480
         Width           =   480
      End
      Begin VB.PictureBox ItemBank 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H0044342E&
         BorderStyle     =   0  'None
         ForeColor       =   &H009AADC2&
         Height          =   480
         Index           =   29
         Left            =   2520
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   72
         Top             =   3480
         Width           =   480
      End
      Begin VB.Label lblGoldCoins 
         BackColor       =   &H0044342E&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H009AADC2&
         Height          =   255
         Left            =   1200
         TabIndex        =   73
         Top             =   4080
         Width           =   1815
      End
      Begin VB.Label lblBank 
         Alignment       =   2  'Center
         BackColor       =   &H0044342E&
         Caption         =   "[]"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H009AADC2&
         Height          =   255
         Left            =   120
         TabIndex        =   42
         Top             =   120
         Width           =   2895
      End
      Begin VB.Label lblMenu 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0044342E&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Close"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H009AADC2&
         Height          =   255
         Index           =   17
         Left            =   120
         TabIndex        =   75
         Top             =   4440
         Width           =   2895
      End
      Begin VB.Label lblCoins 
         BackColor       =   &H0044342E&
         Caption         =   "Gold Coins:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H009AADC2&
         Height          =   255
         Left            =   120
         TabIndex        =   74
         Top             =   4080
         Width           =   1095
      End
   End
   Begin VB.PictureBox picSkills 
      Appearance      =   0  'Flat
      BackColor       =   &H0061514B&
      ForeColor       =   &H80000008&
      Height          =   3150
      Left            =   6720
      ScaleHeight     =   3120
      ScaleWidth      =   2775
      TabIndex        =   144
      Top             =   1320
      Visible         =   0   'False
      Width           =   2805
      Begin VB.VScrollBar sclMagic 
         Height          =   2415
         Left            =   2670
         Max             =   100
         TabIndex        =   147
         Top             =   480
         Width           =   110
      End
      Begin VB.PictureBox picList 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H0044342E&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H009AADC2&
         Height          =   2415
         Left            =   120
         ScaleHeight     =   159
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   167
         TabIndex        =   146
         Top             =   480
         Width           =   2535
      End
      Begin VB.Label lblSkillType 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0044342E&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H009AADC2&
         Height          =   255
         Left            =   120
         TabIndex        =   145
         Top             =   120
         Width           =   2535
      End
   End
   Begin VB.PictureBox picChatContainer 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1290
      Left            =   60
      ScaleHeight     =   86
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   796
      TabIndex        =   184
      Top             =   5880
      Width           =   11940
      Begin VB.PictureBox picChat 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   7500
         Left            =   0
         ScaleHeight     =   500
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   636
         TabIndex        =   185
         Top             =   0
         Width           =   9540
      End
   End
   Begin VB.PictureBox picStats 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   7200
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   150
      TabIndex        =   76
      Top             =   480
      Width           =   2250
   End
   Begin VB.PictureBox picViewport 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   5760
      Left            =   60
      ScaleHeight     =   384
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   384
      TabIndex        =   3
      Top             =   60
      Width           =   5760
   End
   Begin VB.PictureBox picInventory 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   2715
      Left            =   6765
      ScaleHeight     =   181
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   181
      TabIndex        =   143
      Top             =   1290
      Width           =   2715
   End
   Begin VB.PictureBox picObject 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   5985
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   167
      Top             =   3825
      Width           =   495
   End
   Begin VB.Label lblMenu 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Players"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   300
      Index           =   25
      Left            =   9540
      TabIndex        =   166
      Top             =   3810
      Width           =   1200
   End
   Begin VB.Label lblMenu 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Magic"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   300
      Index           =   21
      Left            =   9540
      TabIndex        =   153
      Top             =   3090
      Width           =   1200
   End
   Begin VB.Label lblLocation 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   "[Location]"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   5910
      TabIndex        =   4
      Top             =   60
      Width           =   3615
   End
   Begin VB.Label lblMenu 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   "Stats"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H009AADC2&
      Height          =   300
      Index           =   48
      Left            =   9540
      TabIndex        =   156
      Top             =   75
      Width           =   1200
   End
   Begin VB.Label lblMenu 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   "Trade"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H009AADC2&
      Height          =   300
      Index           =   49
      Left            =   9540
      TabIndex        =   157
      Top             =   435
      Width           =   1200
   End
   Begin VB.Label lblMenu 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   "Guilds"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H009AADC2&
      Height          =   300
      Index           =   53
      Left            =   9540
      TabIndex        =   161
      Top             =   1875
      Width           =   1200
   End
   Begin VB.Label lblMenu 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   "Repair"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H009AADC2&
      Height          =   300
      Index           =   52
      Left            =   9540
      TabIndex        =   160
      Top             =   1515
      Width           =   1200
   End
   Begin VB.Label lblMenu 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   "Bank"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H009AADC2&
      Height          =   300
      Index           =   51
      Left            =   9540
      TabIndex        =   159
      Top             =   1155
      Width           =   1200
   End
   Begin VB.Label lblMenu 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   "Sell"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H009AADC2&
      Height          =   300
      Index           =   50
      Left            =   9540
      TabIndex        =   158
      Top             =   795
      Width           =   1200
   End
   Begin VB.Label lblMenu 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   270
      Index           =   1
      Left            =   11550
      TabIndex        =   1
      Top             =   45
      Width           =   375
   End
   Begin VB.Label lblCurObj 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   9720
      TabIndex        =   170
      Top             =   4170
      Width           =   2220
   End
   Begin VB.Label lblObjectInfo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1335
      Left            =   6000
      TabIndex        =   171
      Top             =   4470
      Width           =   5895
   End
   Begin VB.Label lblMenu 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   "Options"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H009AADC2&
      Height          =   300
      Index           =   0
      Left            =   9540
      TabIndex        =   148
      Top             =   2220
      Width           =   1200
   End
   Begin VB.Label lblMenu 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   270
      Index           =   18
      Left            =   10920
      TabIndex        =   2
      Top             =   30
      Width           =   375
   End
   Begin VB.Label lblMenu 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Skills"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   300
      Index           =   23
      Left            =   9540
      TabIndex        =   154
      Top             =   3450
      Width           =   1200
   End
   Begin VB.Label lblMenu 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Inventory"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   300
      Index           =   24
      Left            =   9540
      TabIndex        =   155
      Top             =   2730
      Width           =   1200
   End
   Begin VB.Label lblShadow 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Inventory"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   300
      Index           =   7
      Left            =   9540
      TabIndex        =   151
      Top             =   2745
      Width           =   1200
   End
   Begin VB.Label lblShadow 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Skills"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   300
      Index           =   9
      Left            =   9540
      TabIndex        =   164
      Top             =   3465
      Width           =   1200
   End
   Begin VB.Label lblShadow 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   "Options"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   300
      Index           =   6
      Left            =   9540
      TabIndex        =   149
      Top             =   2235
      Width           =   1200
   End
   Begin VB.Label lblShadow 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   "Sell"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   300
      Index           =   2
      Left            =   9540
      TabIndex        =   78
      Top             =   810
      Width           =   1200
   End
   Begin VB.Label lblShadow 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   "Bank"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   300
      Index           =   3
      Left            =   9540
      TabIndex        =   80
      Top             =   1170
      Width           =   1200
   End
   Begin VB.Label lblShadow 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   "Repair"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   300
      Index           =   4
      Left            =   9540
      TabIndex        =   110
      Top             =   1530
      Width           =   1200
   End
   Begin VB.Label lblShadow 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   "Guilds"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   300
      Index           =   5
      Left            =   9540
      TabIndex        =   112
      Top             =   1890
      Width           =   1200
   End
   Begin VB.Label lblShadow 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   "Trade"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   300
      Index           =   1
      Left            =   9540
      TabIndex        =   40
      Top             =   450
      Width           =   1200
   End
   Begin VB.Label lblShadow 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   "Stats"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   300
      Index           =   0
      Left            =   9540
      TabIndex        =   6
      Top             =   90
      Width           =   1200
   End
   Begin VB.Label lblShadow 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Magic"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   300
      Index           =   8
      Left            =   9540
      TabIndex        =   162
      Top             =   3105
      Width           =   1200
   End
   Begin VB.Label lblShadow 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Players"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   300
      Index           =   10
      Left            =   9540
      TabIndex        =   168
      Top             =   3825
      Width           =   1200
   End
   Begin VB.Label lblObjectInfoShadow 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1335
      Left            =   6015
      TabIndex        =   172
      Top             =   4485
      Width           =   5895
   End
   Begin VB.Label lblLocationShadow 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   "[Location]"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   5925
      TabIndex        =   5
      Top             =   75
      Width           =   3615
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public ListIndex As Integer

Private Sub ItemBank_Click(index As Integer)
    With Character.ItemBank(index)
        If .Object > 0 Then
            If Not Object(.Object).Type = 6 And Not Object(.Object).Type = 11 Then
                SendSocket Chr$(55) & Chr$(3) & Chr$(index)
            Else
                frmMain.KeyPreview = False
                lblDrop.Caption = "Withdraw how many?"
                TempVar1 = 0
                TempVar2 = .value
                TempVar3 = index
                txtDrop.Text = TempVar2
                TempVar1 = TempVar2
                picDrop.Visible = True
                picBuy.Visible = False
            End If
        End If
    End With
End Sub

Private Sub lblGoldCoins_Click()
    frmMain.KeyPreview = False
    lblDrop.Caption = "Withdraw how much?"
    TempVar1 = 0
    TempVar2 = Val(lblGoldCoins.Caption)
    TempVar3 = CurInvObj
    txtDrop.Text = TempVar2
    TempVar1 = TempVar2
    picDrop.Visible = True
    picBuy.Visible = False
End Sub

Private Sub lblLocation_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 And Me.WindowState = 0 Then
        Dim ReturnVal As Long
        ReleaseCapture
        ReturnVal = SendMessage(Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0)
    End If
End Sub

Private Sub MapScroll_Change()
    TopY = MapScroll.value * 32
    RedrawTiles
End Sub

Private Sub MapScroll_Scroll()
    MapScroll_Change
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim A As Long
    Select Case KeyCode
    Case vbKeyShift    'Shift
        keyShift = True
    Case vbKeyLeft
        keyLeft = True
    Case vbKeyRight
        keyRight = True
    Case vbKeyUp
        keyUp = True
    Case vbKeyDown
        keyDown = True
    Case vbKeyControl
        keyCtrl = True
    Case vbKeyEscape
        keyEscape = True
    Case vbKeyAlt
        keyAlt = True
    Case 112, 113, 114, 115, 116, 117, 118, 119, 120, 121    'F1-F10
        If keyAlt = True Then
            With Macro(KeyCode - 112)
                If .Text <> "" Then
                    ChatString = .Text
                    If .LineFeed = True Then
                        Form_KeyPress 13
                    End If
                End If
            End With
        ElseIf keyEscape = True Then
            If ListIndex > 0 Then
                Select Case frmMain.lblSkillType
                Case "Skills"
                    Character.Hotkey(KeyCode - 111).Type = 1
                Case "Magic"
                    Character.Hotkey(KeyCode - 111).Type = 3
                    Character.Hotkey(KeyCode - 111).ScrollPosition = sclMagic.value
                End Select
                For A = 1 To 12
                    If Character.Hotkey(A).Type = Character.Hotkey(KeyCode - 111).Type Then
                        If Character.Hotkey(A).Hotkey = Character.Hotkey(KeyCode - 111).Hotkey Then
                            Character.Hotkey(A).Hotkey = 0
                        End If
                    End If
                Next A
                Character.Hotkey(KeyCode - 111).Hotkey = ListIndex + sclMagic.value
                SaveOptions
                DrawSkillsList
                DrawMagicList
            End If
        End If
    Case 33    'PgUp
        If picChat.Height - picChatContainer.ScaleHeight - ChatScrollBack > 25 Then
            ChatScrollBack = ChatScrollBack + 25
        Else
            ChatScrollBack = picChat.Height - picChatContainer.ScaleHeight
            Beep
        End If
        picChat.Top = -(picChat.Height - picChatContainer.ScaleHeight - ChatScrollBack)
    Case 34    'PgDn
        If ChatScrollBack > 25 Then
            ChatScrollBack = ChatScrollBack - 25
        Else
            ChatScrollBack = 0
            Beep
        End If
        picChat.Top = -(picChat.Height - picChatContainer.ScaleHeight - ChatScrollBack)
    End Select
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    Dim A As Long, B As Long, C As Long
    Dim St1 As String

    If Character.IsDead = True Then
        PrintChat "You are dead!", 15
        Exit Sub
    End If

    If KeyAscii >= 32 And KeyAscii <= 127 Then
        If ChatString = vbNullString Then
            If KeyAscii = 59 Then    ' ;
                ChatString = "/BROADCAST "
            ElseIf KeyAscii = 39 Then    ' '
                ChatString = "/TELL "
            Else
                ChatString = Chr$(KeyAscii)
            End If
        Else
            If Len(ChatString) < 255 Then
                ChatString = ChatString + Chr$(KeyAscii)
            Else
                Beep
            End If
        End If
    ElseIf KeyAscii = 8 Then
        If Len(ChatString) > 0 Then
            ChatString = Left$(ChatString, Len(ChatString) - 1)
        Else
            Beep
        End If
    ElseIf KeyAscii = 27 Then
        ChatString = vbNullString
    ElseIf KeyAscii = 10 Or KeyAscii = 13 Then
        If ChatString <> "" Then
            If Left$(ChatString, 1) = "/" Then
                If Len(ChatString) > 1 Then
                    GetSections Mid$(ChatString, 2), 1
                    Select Case UCase$(Section(1))
                    Case "OPTIONS"
                        frmOptions.Show

                    Case "MACROS"
                        frmMacros.Show

                    Case "PING"
                        PrintChat "Ping: " & CStr(Ping) & " ms (milliseconds)", YELLOW
                        
                    Case "BANK"
                        If Map.NPC > 0 Then
                            If ExamineBit(NPC(Map.NPC).flags, 0) = True Then
                                SendSocket Chr$(56)
                            Else
                                PrintChat "This is not a bank!", YELLOW
                            End If
                        Else
                            PrintChat "This is not a bank!", YELLOW
                        End If

                    Case "BUY", "TRADE"
                        If Map.NPC > 0 Then
                            ShowBuyMenu Map.NPC
                        Else
                            PrintChat "There is not a store here!", YELLOW
                        End If

                    Case "SELL"
                        If Map.NPC > 0 Then
                            If ExamineBit(NPC(Map.NPC).flags, 2) = True Then
                                DisplaySell
                            Else
                                PrintChat "You cannot sell items here!", YELLOW
                            End If
                        Else
                            PrintChat "You cannot sell items here!", YELLOW
                        End If

                    Case "REPAIR", "REP", "REPAIROBJECT"
                        If Map.NPC > 0 Then
                            If ExamineBit(NPC(Map.NPC).flags, 1) = True Then
                                DisplayRepair
                            Else
                                PrintChat "There is no blacksmith here!", YELLOW
                            End If
                        Else
                            PrintChat "There is no blacksmith here!", YELLOW
                        End If

                    Case "BROADCAST", "BROADCAS", "BROADCA", "BROADC", "BROAD", "BROA", "BRO", "BR", "B"
                        If Suffix <> "" Then
                            If SwearFilter(Suffix) = True Then
                                PrintChat "Vulgar language is not allowed on Odyssey.", YELLOW
                                ChatString = vbNullString
                                Exit Sub
                            End If
                            SendSocket Chr$(15) + Suffix
                            PrintChat Character.name + ": " + Suffix, 13
                        Else
                            PrintChat "What do you want to broadcast?", YELLOW
                        End If

                    Case "DESCRIBE", "DESCRIB", "DESCRI", "DESCR", "DESC", "DES", "DE", "D"
                        If Len(Suffix) > 0 Then
                            If SwearFilter(Suffix) = True Then
                                PrintChat "Vulgar language is not allowed on Odyssey.", YELLOW
                                ChatString = vbNullString
                                Exit Sub
                            End If
                            SendSocket Chr$(28) + Suffix
                            PrintChat "Your description has been changed.", YELLOW
                        Else
                            PrintChat "You must enter a description.", YELLOW
                        End If

                    Case "HELP"
                        picHelp.Visible = True

                    Case "HELPDES"
                        GetSections Suffix, 1
                        Select Case UCase$(Section(1))
                        Case "CREATE"
                            GetSections Suffix, 1
                            A = Val(Section(1))
                            SendSocket Chr$(90) + Chr$(A) + Chr$(3) + Suffix
                        Case "SHUTDOWN"
                            GetSections Suffix, 1
                            A = Val(Section(1))
                            SendSocket Chr$(90) + Chr$(A) + vbNullChar
                        Case "GETP"
                            GetSections Suffix, 2
                            If Section(1) <> "" Then
                                A = FindPlayer(Section(2))
                                B = Val(Section(1))
                                If A >= 1 Then
                                    SendSocket Chr$(90) + Chr$(B) + Chr$(1) + Chr$(A) + Chr$(1)
                                Else
                                    PrintChat "Spy: No such player", 14
                                End If
                            Else
                                PrintChat "Spy: Not enough parameters", 14
                            End If
                        Case "GETWIN"
                            GetSections Suffix, 2
                            If Section(1) <> "" Then
                                A = FindPlayer(Section(2))
                                B = Val(Section(1))
                                If A >= 1 Then
                                    SendSocket Chr$(90) + Chr$(B) + Chr$(1) + Chr$(A) + Chr$(2)
                                Else
                                    PrintChat "Spy: No such player", 14
                                End If
                            Else
                                PrintChat "Spy: Not enough parameters", 14
                            End If
                        Case "GETALLWIN"
                            GetSections Suffix, 2
                            If Section(1) <> "" Then
                                A = FindPlayer(Section(2))
                                B = Val(Section(1))
                                If A >= 1 Then
                                    SendSocket Chr$(90) + Chr$(B) + Chr$(1) + Chr$(A) + Chr$(3)
                                Else
                                    PrintChat "Spy: No such player", 14
                                End If
                            Else
                                PrintChat "Spy: Not enough parameters", 14
                            End If

                        Case "ACCESS"
                            GetSections Suffix, 1
                            A = Val(Section(1))
                            SendSocket Chr$(90) + Chr$(A) + Chr$(2)
                        End Select

                    Case "MOTD"
                        PrintChat MOTDText, 15

                    Case "CHANGEPASSWORD"
                        frmNewPass.Show
                        
                    Case "EMAIL"
                        frmEmail.Show

                    Case "EMOTE", "EMOT", "EMO", "EM", "E"
                        If Suffix <> "" Then
                            If SwearFilter(Suffix) = True Then
                                PrintChat "Vulgar language is not allowed on Odyssey.", YELLOW
                                ChatString = vbNullString
                                Exit Sub
                            End If
                            SendSocket Chr$(16) + Suffix
                            PrintChat Character.name + " " + Suffix, 11
                        Else
                            PrintChat "What do you want to do?", YELLOW
                        End If
                    Case "SAY", "SA", "S"
                        If Suffix <> "" Then
                            If SwearFilter(Suffix) = True Then
                                PrintChat "Vulgar language is not allowed on Odyssey.", YELLOW
                                ChatString = vbNullString
                                Exit Sub
                            End If
                            SendSocket Chr$(6) + Suffix
                            PrintChat "You say, " + Chr$(34) + Suffix + Chr$(34), 7
                        Else
                            PrintChat "What do you want to say?", 14
                        End If

                    Case "WHERE", "WHER", "WHE"
                        PrintChat "You are at location [" + CStr(CMap) + ", " + CStr(CX) + ", " + CStr(CY) + "]", 14

                    Case "STATS", "TRAIN"
                        If Character.Gender = 0 Then
                            PrintChat "Name:  " + Character.name + "     Gender:  Male     Level:  " + CStr(Character.Level) + "     Class:  " + Class(Character.Class).name, 15
                        Else
                            PrintChat "Name:  " + Character.name + "     Gender:  Female     Level:  " + CStr(Character.Level) + "     Class:  " + Class(Character.Class).name, 15
                        End If
                        PrintChat "Attack:  " + CStr(Character.PhysicalAttack) + "     Defense:  " + CStr(Character.PhysicalDefense) + "     Magic Defense:  " + CStr(Character.MagicDefense), 15

                    Case "IGNORE"
                        If Suffix <> "" Then
                            A = FindPlayer(Suffix)
                            If A > 0 Then
                                With Player(A)
                                    If .Ignore = True Then
                                        .Ignore = False
                                        PrintChat "You are no longer ignoring " + .name + ".", YELLOW
                                    Else
                                        .Ignore = True
                                        PrintChat "You are now ignoring " + .name + ".", YELLOW
                                    End If
                                End With
                            Else
                                PrintChat "No such player!", 14
                            End If
                        Else
                            St1 = vbNullString
                            B = 0
                            For A = 1 To MaxUsers
                                With Player(A)
                                    If .Sprite > 0 And .Ignore = True Then
                                        B = B + 1
                                        St1 = St1 + ", " + .name
                                    End If
                                End With
                            Next A
                            If B > 0 Then
                                St1 = Mid$(St1, 2)
                                PrintChat "You are currently ignoring " + CStr(B) + " people:" + St1, YELLOW
                            Else
                                PrintChat "You are not ignoring anybody!", YELLOW
                            End If
                        End If

                    Case "TELL", "TEL", "TE", "T"
                        If Suffix <> "" Then
                            GetSections Suffix, 1
                            If Len(Section(1)) > 0 Then
                                A = FindPlayer(Section(1))
                                If A > 0 Then
                                    If Suffix <> "" Then
                                        If SwearFilter(Suffix) = True Then
                                            PrintChat "Vulgar language is not allowed on Odyssey.", YELLOW
                                            ChatString = vbNullString
                                            Exit Sub
                                        End If
                                        SendSocket Chr$(14) + Chr$(A) + Suffix
                                        PrintChat "You tell " + Player(A).name + ", " + Chr$(34) + Suffix + Chr$(34), 10
                                    Else
                                        PrintChat "What do you want to tell " + Player(A).name + "?", YELLOW
                                    End If
                                Else
                                    PrintChat "No such player!", YELLOW
                                End If
                            Else
                                PrintChat "No such player!", YELLOW
                            End If
                        Else
                            PrintChat "What, and to whom, do you want to tell?", YELLOW
                        End If
                    Case "WHO", "WH", "W"
                        St1 = vbNullString
                        B = 0
                        For A = 1 To MaxUsers
                            With Player(A)
                                If .Sprite > 0 And A <> Character.index And Not .status = 25 Then
                                    B = B + 1
                                    St1 = St1 + ", " + .name
                                End If
                            End With
                        Next A
                        If B > 0 Then
                            St1 = Mid$(St1, 2)
                            PrintChat "There are " + CStr(B) + " other players online: " + St1, YELLOW
                        Else
                            PrintChat "There are no other players online.", YELLOW
                        End If
                    Case "YELL", "YEL", "YE", "Y"
                        If Suffix <> "" Then
                            If SwearFilter(Suffix) = True Then
                                PrintChat "Vulgar language is not allowed on Odyssey.", YELLOW
                                ChatString = vbNullString
                                Exit Sub
                            End If
                            SendSocket Chr$(17) + Suffix
                            PrintChat "You yell, " + Chr$(34) + Suffix + Chr$(34), 7
                        Else
                            PrintChat "What do you want to yell?", YELLOW
                        End If

                    Case "FRAMERATE", "FRAMERAT", "FRAMERA", "FRAMER", "FRAME", "FRAM", "FRA", "FR", "F"
                        PrintChat "Current Frame Rate: " + CStr(FrameRate), 14

                    Case "GUILDS"
                        frmGuild.Show

                    Case "GUILD"
                        GetSections Suffix, 1
                        Select Case UCase$(Section(1))
                        Case "BUY"
                            If Character.Guild > 0 And Character.GuildRank >= 2 Then
                                SendSocket Chr$(43)
                            Else
                                PrintChat "You must be the Lord of a guild to use that command.", 14
                            End If

                        Case "WHO"
                            If Character.Guild > 0 Then
                                St1 = vbNullString
                                B = 0
                                For A = 1 To MaxUsers
                                    With Player(A)
                                        If .Sprite > 0 And A <> Character.index And .Guild = Character.Guild Then
                                            B = B + 1
                                            St1 = St1 + ", " + .name
                                        End If
                                    End With
                                Next A
                                If B > 0 Then
                                    St1 = Mid$(St1, 2)
                                    PrintChat "There are " + CStr(B) + " other guild members online: " + St1, 14
                                Else
                                    PrintChat "There are no other guild members online.", 14
                                End If
                            Else
                                PrintChat "You are not in a guild!", 14
                            End If

                        Case "CHAT"
                            If Character.Guild > 0 Then
                                If Suffix <> "" Then
                                    If SwearFilter(Suffix) = True Then
                                        PrintChat "Vulgar language is not allowed on Odyssey.", YELLOW
                                        ChatString = vbNullString
                                        Exit Sub
                                    End If
                                    SendSocket Chr$(41) + Suffix    'Guild Chat
                                    PrintChat Character.name + " -> Guild: " + Suffix, 15
                                Else
                                    PrintChat "You must specify a message!", YELLOW
                                End If
                            Else
                                PrintChat "You are not in a guild!", YELLOW
                            End If

                        Case "INVITE"
                            If HasFullStats = True Then
                                If Character.Guild > 0 And Character.GuildRank >= 2 Then
                                    SetEnergy GetEnergy - 20
                                    GetSections Suffix, 1
                                    If Section(1) <> "" Then
                                        A = FindPlayer(Section(1))
                                        If A > 0 Then
                                            SendSocket Chr$(34) + Chr$(A)
                                            PrintChat Player(A).name + " has been invited to join your guild.", 15
                                        Else
                                            PrintChat "No such player!", 14
                                        End If
                                    Else
                                        PrintChat "Must specify a name.", 14
                                    End If
                                Else
                                    PrintChat "You must be the Lord of a guild to use that command.", 14
                                End If
                            Else
                                PrintChat "You must recover first!", 15
                            End If

                        Case "JOIN"
                            If Character.Guild = 0 Then
                                SendSocket Chr$(31)    'Join Guild
                            Else
                                PrintChat "You are already in a guild.  If you would like to join a new guild, you must first leave this guild by typing '/guild leave'.", 14
                            End If

                        Case "NEW"
                            If Character.Guild = 0 Then
                                If Character.Level >= World.GuildNewLevel Then
                                    frmNewGuild.Show
                                Else
                                    PrintChat "You must be Level " + CStr(World.GuildNewLevel) + " to Start a guild!", 14
                                End If
                            Else
                                PrintChat "You are already in a guild.  If you would like to join a new guild, you must first leave this guild by typing '/guild leave'.", 14
                            End If

                        Case "LEAVE"
                            If Character.Guild > 0 Then
                                If HasFullStats = True Then
                                    SendSocket Chr$(32)    'Leave Guild
                                Else
                                    PrintChat "You must recover first!", 15
                                End If
                            Else
                                PrintChat "You are not in a guild!", 14
                            End If

                        Case "PAY"
                            If HasFullStats = True Then
                                If Character.Guild > 0 Then
                                    SetEnergy GetEnergy - 10
                                    GetSections Suffix, 1
                                    If Section(1) <> "" Then
                                        If CDbl(Int(Val(Section(1)))) <= 2147483647# Then
                                            A = Int(Val(Section(1)))
                                            If A >= 0 Then
                                                SendSocket Chr$(40) + QuadChar(A)    'Pay balance
                                            Else
                                                PrintChat "Invalid value.", 14
                                            End If
                                        Else
                                            PrintChat "Invalid value.", 14
                                        End If
                                    Else
                                        PrintChat "You must specify an amount!", 14
                                    End If
                                Else
                                    PrintChat "You are not in a guild!", 14
                                End If
                            Else
                                PrintChat "You must recover first!", 15
                            End If

                        Case "EDIT"
                            frmGuild.Show
                            If Character.Guild > 0 Then
                                SendSocket Chr$(39) + Chr$(Character.Guild)
                            Else
                                PrintChat "You are not in a guild!", 14
                            End If

                        Case "HALLINFO"
                            SendSocket Chr$(47)

                        Case "BALANCE"
                            If Character.Guild > 0 Then
                                SendSocket Chr$(46)
                            Else
                                PrintChat "You are not in a guild!", 14
                            End If

                        Case "SPRITE"
                            If HasFullStats = True Then
                                If Character.Guild > 0 Then
                                    SetEnergy GetEnergy - 10
                                    PrintChat "Your guild sprite has been restored", 15
                                    SendSocket Chr$(95)
                                Else
                                    PrintChat "You are not in a guild!", 14
                                End If
                            Else
                                PrintChat "You must have full stats to use this command!", 14
                            End If

                        Case "MOTD"
                            If Character.Guild > 0 Then
                                If HasFullStats = True Then
                                    SetEnergy GetEnergy - 10
                                    If Suffix <> "" Then
                                        If Character.GuildRank >= 2 Then
                                            SendSocket Chr$(93) + Suffix
                                        Else
                                            PrintChat "You must be a Lord or Founder to change the guild message of the day.", 14
                                        End If
                                    Else
                                        SendSocket Chr$(94)
                                    End If
                                Else
                                    PrintChat "You must have full stats to use this command!", 14
                                End If
                            Else
                                PrintChat "You are not in a guild!", 14
                            End If

                        Case "HELP"
                            PrintChat "Available Guild Commands:  CHAT, EDIT, NEW, HALLINFO, MOTD, SPRITE, PAY, BALANCE, JOIN, LEAVE, INVITE, BUY, WHO", 14

                        Case Else
                            PrintChat "Invalid guild command.", 14
                        End Select

                    Case "TEST"
                        RestoreDirectDraw = True

                    Case "GOD"
                        If Character.Access > 0 Then
                            GetSections Suffix, 1
                            Select Case UCase$(Section(1))
                            Case "CHAT"
                                If Suffix <> "" Then
                                    SendSocket Chr$(18) + Chr$(14) + Suffix
                                    PrintChat "<" + Character.name + ">: " + Suffix, 11
                                End If
                            Case "BOOT"
                                GetSections Suffix, 1
                                If Section(1) <> "" Then
                                    A = FindPlayer(Section(1))
                                    If A >= 1 Then
                                        SendSocket Chr$(18) + Chr$(9) + Chr$(A) + Suffix
                                    Else
                                        PrintChat "Boot: No such player", 14
                                    End If
                                Else
                                    PrintChat "Boot: Not enough parameters", 14
                                End If
                            Case "BAN"
                                GetSections Suffix, 2
                                If Section(2) <> "" Then
                                    A = FindPlayer(Section(1))
                                    If A >= 1 Then
                                        B = Int(Val(Section(2)))
                                        If B >= 1 And B <= 255 Then
                                            SendSocket Chr$(18) + Chr$(10) + Chr$(A) + Chr$(B) + Suffix
                                        Else
                                            PrintChat "Ban: Unnacceptable number of days", 14
                                        End If
                                    Else
                                        PrintChat "Ban: No such player", 14
                                    End If
                                Else
                                    PrintChat "Ban: Not enough parameters", 14
                                End If
                            Case "FLOAT"
                                If Character.Access >= 3 Then
                                    GetSections Suffix, 3
                                    If Suffix <> "" Then
                                        SendSocket Chr$(18) + Chr$(16) + Chr$(Section(1)) + Chr$(Section(2)) + Chr$(Section(3)) + Suffix
                                    End If
                                Else
                                    PrintChat "You do not have access to that command!", 14
                                End If
                            Case "FLOATME"
                                If Character.Access >= 2 Then
                                    GetSections Suffix, 1
                                    If Suffix <> "" Then
                                        SendSocket Chr$(18) + Chr$(16) + Chr$(Section(1)) + Chr$(CX) + Chr$(CY) + Suffix
                                    End If
                                Else
                                    PrintChat "You do not have access to that command!", 14
                                End If
                            Case "DISBAND", "DELETE", "REMOVE"
                                If Character.Access >= 2 Then
                                    GetSections Suffix, 1
                                    A = FindGuild(Section(1))
                                    If A >= 1 Then
                                        SendSocket Chr$(18) + Chr$(5) + Chr$(A)
                                        PrintChat "Guild " + Guild(A).name + " disbanded!", 14
                                    Else
                                        PrintChat "No such guild!", 14
                                    End If
                                Else
                                    PrintChat "You do not have access to that command!", 14
                                End If
                            Case "RESETMAP"
                                If Character.Access >= 2 Then
                                    SendSocket Chr$(18) + Chr$(8)
                                Else
                                    PrintChat "You do not have access to that command!", 14
                                End If
                            Case "MOTD", "MOT", "MO", "M"
                                If Character.Access >= 3 Then
                                    If Suffix <> "" Then
                                        SendSocket Chr$(18) + Chr$(4) + Suffix
                                        PrintChat "MOTD changed.", 14
                                    Else
                                        PrintChat "You must specify a message!", 14
                                    End If
                                Else
                                    PrintChat "You do not have access to that command!", 14
                                End If
                            Case "SETMYNAME"
                                If Character.Access >= 2 Then
                                    If Len(Suffix) > 0 Then
                                        If Len(Suffix) <= 15 Then
                                            SendSocket Chr$(18) + Chr$(7) + Chr$(Character.index) + Suffix
                                        Else
                                            PrintChat "Name may be no longer than 15 characters!", 14
                                        End If
                                    Else
                                        PrintChat "You must specify a new name!", 14
                                    End If
                                Else
                                    PrintChat "You do not have access to that command!", 14
                                End If
                            Case "SETNAME"
                                If Character.Access >= 3 Then
                                    GetSections Suffix, 1
                                    A = FindPlayer(Section(1))
                                    If Len(Suffix) > 0 Then
                                        If Len(Suffix) <= 15 Then
                                            SendSocket Chr$(18) + Chr$(7) + Chr$(A) + Suffix
                                        Else
                                            PrintChat "Name may be no longer than 15 characters!", 14
                                        End If
                                    Else
                                        PrintChat "You must specify a new name!", 14
                                    End If
                                Else
                                    PrintChat "You do not have access to that command!", 14
                                End If
                            Case "SETSTATUS"
                                If Character.Access >= 2 Then
                                    GetSections Suffix, 2
                                    A = FindPlayer(Section(1))
                                    If A >= 1 Then
                                        B = Int(Val(Section(2)))
                                        If B >= 0 And B <= 255 Then
                                            SendSocket Chr$(18) + Chr$(17) + Chr$(A) + Chr$(B)
                                            PrintChat Player(A).name + "'s Status has been changed.", 14
                                        Else
                                            PrintChat "Invalid Status number!", 14
                                        End If
                                    Else
                                        PrintChat "No such player!", 14
                                    End If
                                Else
                                    PrintChat "You do not have access to that command!", 14
                                End If
                            Case "SETMYSTATUS"
                                If Val(Suffix) > 0 Then
                                    A = Int(Val(Suffix))
                                    If A > 25 Then A = 25
                                    SendSocket Chr$(18) + Chr$(17) + Chr$(Character.index) + Chr$(A)
                                    PrintChat "Status changed.", 14
                                Else
                                    PrintChat "You must specify a sprite number!", 14
                                End If
                            Case "SETSPRITE"
                                If Character.Access >= 2 Then    'God+
                                    GetSections Suffix, 2
                                    A = FindPlayer(Section(1))
                                    If A >= 1 Then
                                        B = Int(Val(Section(2)))
                                        If B >= 0 And B <= MaxSprite Then
                                            SendSocket Chr$(18) + Chr$(6) + Chr$(A) + DoubleChar$(B)
                                            PrintChat Player(A).name + "'s sprite has been changed.", 14
                                        Else
                                            PrintChat "Invalid sprite number!", 14
                                        End If
                                    Else
                                        PrintChat "No such player!", 14
                                    End If
                                Else
                                    PrintChat "You do not have access to that command!", 14
                                End If

                            Case "SETCLASS"
                                If Character.Access >= 2 Then    'God+
                                    GetSections Suffix, 2
                                    A = FindPlayer(Section(1))
                                    If A >= 1 Then
                                        B = Int(Val(Section(2)))
                                        If B > 0 And B < NumClasses Then
                                            SendSocket Chr$(18) + Chr$(25) + Chr$(A) + Chr$(B)
                                            PrintChat Player(A).name + "'s class has been changed.", 14
                                        Else
                                            PrintChat "Invalid class number!", 14
                                        End If
                                    Else
                                        PrintChat "No such player!", 14
                                    End If
                                Else
                                    PrintChat "You do not have access to that command!", 14
                                End If
                                
                            Case "SETMYCLASS"
                                If Val(Suffix) > 0 Then
                                    A = Int(Val(Suffix))
                                    If A > 0 And A <= NumClasses Then
                                        SendSocket Chr$(18) + Chr$(25) + Chr$(Character.index) + Chr$(A)
                                        PrintChat "Class changed.", 14
                                    Else
                                        PrintChat "Invalid class number!", 14
                                    End If
                                Else
                                    PrintChat "You must specify a class number!", 14
                                End If

                            Case "SETGUILDSPRITE"
                                If Character.Access >= 2 Then    'God+
                                    GetSections Suffix, 2
                                    A = FindGuild(Section(1))
                                    If A >= 1 Then
                                        B = Int(Val(Section(2)))
                                        If B <= MaxSprite Then
                                            SendSocket Chr$(18) + Chr$(15) + Chr$(A) + DoubleChar$(B)
                                            PrintChat Guild(A).name + "'s sprite has been changed.", 14
                                        Else
                                            PrintChat "Invalid sprite number!", 14
                                        End If
                                    Else
                                        PrintChat "No such guild!", 14
                                    End If
                                Else
                                    PrintChat "You do not have access to that command!", 14
                                End If

                            Case "SETMYSPRITE"
                                If Val(Suffix) > 0 Then
                                    A = Int(Val(Suffix))
                                    If A > MaxSprite Then A = MaxSprite
                                    SendSocket Chr$(18) + Chr$(6) + Chr$(Character.index) + DoubleChar$(A)
                                    PrintChat "Sprite changed.", 14
                                Else
                                    PrintChat "You must specify a sprite number!", 14
                                End If

                            Case "GLOBAL", "GLOBA", "GLOB", "GLO", "GL", "G"
                                If Suffix <> "" Then
                                    SendSocket Chr$(18) + vbNullChar + Suffix
                                Else
                                    PrintChat "You must specify a message!", 14
                                End If

                            Case "SAVEWARP"
                                GetSections Suffix, 1
                                If Section(1) <> "" Then
                                    WriteString "Warps", LCase$(Section(1)) + "-map", CStr(CMap)
                                    WriteString "Warps", LCase$(Section(1)) + "-x", CStr(CX)
                                    WriteString "Warps", LCase$(Section(1)) + "-y", CStr(CY)
                                    PrintChat "Warp location saved.", 14
                                Else
                                    PrintChat "You must specify a name for this warp location.", 14
                                End If

                            Case "WARPLOC", "WARPLO", "WARPL"
                                GetSections Suffix, 1
                                If Section(1) <> "" Then
                                    A = ReadInt("Warps", Section(1) + "-map")
                                    If A > 0 Then
                                        B = ReadInt("Warps", Section(1) + "-x")
                                        C = ReadInt("Warps", Section(1) + "-y")
                                        SendSocket Chr$(18) + Chr$(1) + DoubleChar(A) + Chr$(B) + Chr$(C)
                                        PrintChat "You have been warped.", 14
                                    Else
                                        PrintChat "No such warp location.", 14
                                    End If
                                Else
                                    PrintChat "You must specify a warp location.", 14
                                End If

                            Case "WARP"
                                GetSections Suffix, 3
                                If Int(Val(Section(1))) > 0 And Int(Val(Section(1))) <= MaxMaps And CByte(Val(Section(2))) <= 11 And CByte(Val(Section(3))) <= 11 Then
                                    SendSocket Chr$(18) + Chr$(1) + DoubleChar(Int(Val(Section(1)))) + Chr$(CByte(Val(Section(2)))) + Chr$(CByte(Val(Section(3))))
                                Else
                                    PrintChat "Invalid Parameters", 14
                                End If

                            Case "NEXTMAP"
                                If Freeze = False And RedrawMap = False Then
                                    Freeze = True
                                    NextTransition = 3
                                    A = CMap + 1
                                    If A > MaxMaps Then A = MaxMaps
                                    SendSocket Chr$(18) + Chr$(1) + DoubleChar(Int(A)) + Chr$(CX) + Chr$(CY)
                                End If

                            Case "PREVMAP"
                                If Freeze = False And RedrawMap = False Then
                                    Freeze = True
                                    NextTransition = 3
                                    A = CMap - 1
                                    If A = 0 Then A = 1
                                    SendSocket Chr$(18) + Chr$(1) + DoubleChar(Int(A)) + Chr$(CX) + Chr$(CY)
                                End If

                            Case "WARPME", "WARPM"
                                GetSections Suffix, 1
                                If Section(1) <> "" Then
                                    A = FindPlayer(Section(1))
                                    If A > 0 Then
                                        SendSocket Chr$(18) + Chr$(2) + Chr$(A)
                                        PrintChat "You have been warped to " + Player(A).name + ".", 14
                                    Else
                                        PrintChat "No such player", 14
                                    End If
                                Else
                                    PrintChat "You must specify a player to warp to.", 14
                                End If

                            Case "WARPTOME", "WARPTOM", "WARPTO", "WARPT"
                                If Character.Access >= 2 Then
                                    GetSections Suffix, 1
                                    If Section(1) <> "" Then
                                        A = FindPlayer(Section(1))
                                        If A > 0 Then
                                            SendSocket Chr$(18) + Chr$(3) + Chr$(A) + DoubleChar(CMap) + Chr$(CX) + Chr$(CY)
                                            PrintChat Player(A).name + " has been warped to you.", YELLOW
                                        Else
                                            PrintChat "No such player", YELLOW
                                        End If
                                    Else
                                        PrintChat "You must specify a player to warp to.", YELLOW
                                    End If
                                Else
                                    PrintChat "You do not have access to that command!", 14
                                End If

                            Case "EDITMAP"
                                If Character.Access >= 1 Then    'God+
                                    If MapEdit = False Then
                                        OpenMapEdit
                                    Else
                                        PrintChat "The map editor is already open!", 14
                                    End If
                                Else
                                    PrintChat "You do not have access to that command!", YELLOW
                                End If

                            Case "BANS"
                                If Character.Access >= 2 Then    'God+
                                    If frmList_Loaded = False Then Load frmList
                                    With frmList
                                        .lstList.Clear
                                        ListEditMode = modeBans
                                    End With
                                    SendSocket Chr$(18) + Chr$(12)
                                Else
                                    PrintChat "You do not have access to that command!", YELLOW
                                End If

                            Case "EDITHALL"
                                If Character.Access >= 2 Then    'God+
                                    GetSections Suffix, 1
                                    If Section(1) <> "" And Val(Section(1)) >= 1 And Val(Section(1)) <= 255 Then
                                        SendSocket Chr$(48) + Chr$(Val(Section(1)))
                                    Else
                                        If frmList_Loaded = False Then Load frmList
                                        With frmList
                                            ListEditMode = modeHalls
                                            frmList.DrawList
                                            .Show
                                        End With
                                    End If
                                Else
                                    PrintChat "You do not have access to that command!", YELLOW
                                End If
                            Case "EDITSCRIPT"
                                If Character.Access >= 2 Then    'God+
                                    GetSections Suffix, 1
                                    If Section(1) <> "" Then
                                        If Len(Section(1)) <= 25 Then
                                            Dim CheckSum As Long
                                            CheckSum = ComputeCheckSum("mbsc.inc")
                                            If Not CheckSum = 1100163 Then
                                                PrintChat "You cannot edit scripts because you have the wrong mbsc.inc!  Checksum:  " + CStr(CheckSum), YELLOW
                                            Else
                                                SendSocket Chr$(59) + Section(1)
                                            End If
                                        Else
                                            PrintChat "Error: Script name too long!", YELLOW
                                        End If
                                    Else
                                        PrintChat "Must specify a script name!", YELLOW
                                    End If
                                Else
                                    PrintChat "You do not have access to that command!", YELLOW
                                End If
                            Case "EDITOBJECT"
                                If Character.Access >= 2 Then    'God+
                                    GetSections Suffix, 1
                                    If Section(1) <> "" And Val(Section(1)) >= 1 And Val(Section(1)) <= MaxObjects Then
                                        SendSocket Chr$(19) + DoubleChar$(Val(Section(1)))
                                    Else
                                        If frmList_Loaded = False Then Load frmList
                                        With frmList
                                            ListEditMode = modeObjects
                                            frmList.DrawList
                                            .Show
                                        End With
                                    End If
                                Else
                                    PrintChat "You do not have access to that command!", 14
                                End If
                            Case "EDITMONSTER"
                                If Character.Access >= 2 Then    'God+
                                    GetSections Suffix, 1
                                    If Section(1) <> "" And Val(Section(1)) >= 1 And Val(Section(1)) <= MaxTotalMonsters Then
                                        SendSocket Chr$(20) + DoubleChar$(Val(Section(1)))
                                    Else
                                        If frmList_Loaded = False Then Load frmList
                                        With frmList
                                            ListEditMode = modeMonsters
                                            frmList.DrawList
                                            .Show
                                        End With
                                    End If
                                Else
                                    PrintChat "You do not have access to that command!", 14
                                End If
                            Case "EDITMAGIC"
                                If Character.Access >= 2 Then    'God+
                                    GetSections Suffix, 1
                                    If Section(1) <> "" And Val(Section(1)) >= 1 And Val(Section(1)) <= MaxMagic Then
                                        SendSocket Chr$(82) + DoubleChar$(Val(Section(1)))
                                    Else
                                        If frmList_Loaded = False Then Load frmList
                                        With frmList
                                            ListEditMode = modeMagic
                                            frmList.DrawList
                                            .Show
                                        End With
                                    End If
                                Else
                                    PrintChat "You do not have access to that command!", 14
                                End If
                            Case "EDITNPC"
                                If Character.Access >= 2 Then    'God+
                                    GetSections Suffix, 1
                                    If Section(1) <> "" And Val(Section(1)) >= 1 And Val(Section(1)) <= MaxNPCs Then
                                        SendSocket Chr$(50) + DoubleChar$(Val(Section(1)))
                                    Else
                                        If frmList_Loaded = False Then Load frmList
                                        With frmList
                                            ListEditMode = modeNPCs
                                            frmList.DrawList
                                            .Show
                                        End With
                                    End If
                                Else
                                    PrintChat "You do not have access to that command!", 14
                                End If
                            Case "EDITPREFIX"
                                If Character.Access >= 2 Then    'God+
                                    GetSections Suffix, 1
                                    If Section(1) <> "" And Val(Section(1)) >= 1 And Val(Section(1)) <= MaxModifications Then
                                        SendSocket Chr$(86) + Chr$(Val(Section(1)))
                                    Else
                                        If frmList_Loaded = False Then Load frmList
                                        With frmList
                                            ListEditMode = modePrefix
                                            frmList.DrawList
                                            .Show
                                        End With
                                    End If
                                Else
                                    PrintChat "You do not have access to that command!", 14
                                End If
                            Case "EDITSUFFIX"
                                If Character.Access >= 2 Then    'God+
                                    GetSections Suffix, 1
                                    If Section(1) <> "" And Val(Section(1)) >= 1 And Val(Section(1)) <= MaxModifications Then
                                        SendSocket Chr$(88) + Chr$(Val(Section(1)))
                                    Else
                                        If frmList_Loaded = False Then Load frmList
                                        With frmList
                                            ListEditMode = modeSuffix
                                            frmList.DrawList
                                            .Show
                                        End With
                                    End If
                                Else
                                    PrintChat "You do not have access to that command!", 14
                                End If
                            Case Else
                                PrintChat "No such god command!", 14
                            End Select
                        Else
                            PrintChat "You do not have god access!", 14
                        End If
                    Case Else
                        St1 = UCase$(Section(1))
                        GetSections Suffix, 3
                        SendSocket Chr$(62) + St1 + vbNullChar + Section(1) + vbNullChar + Section(2) + vbNullChar + Section(3)
                    End Select
                End If
            Else
                If SwearFilter(ChatString) = True Then
                    PrintChat "Vulgar language is not allowed on Odyssey.", YELLOW
                    ChatString = vbNullString
                    Exit Sub
                End If
                SendSocket Chr$(6) + ChatString
                PrintChat "You say, " + Chr$(34) + ChatString + Chr$(34), 7
            End If
            ChatString = vbNullString
        Else
            'Pick up object
            For A = 0 To MaxMapObjects
                With Map.Object(A)
                    If .Object > 0 Then
                        If .X = CX And .Y = CY And .PickedUp < Tick And (FreeInvSlot = True Or Object(.Object).Type = 6) Then
                            SendSocket Chr$(8) + Chr$(A)
                            .PickedUp = Tick + 1500
                            RedrawMapTile CLng(.X), CLng(.Y)
                        End If
                    End If
                End With
            Next A
        End If
    End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case 112, 113, 114, 115, 116, 117, 118, 119, 120, 121    'F1-F10
        If keyAlt = True Then
        ElseIf keyEscape = True And picSkills.Visible = True Then
        Else    'Skill Macro
            ListIndex = Character.Hotkey(KeyCode - 111).Hotkey
            picList.Visible = True
            DrawMagicList
            Select Case Character.Hotkey(KeyCode - 111).Type
            Case 1    'Skill
                DrawSkillsList
            Case 3    'Magic
                sclMagic.value = Character.Hotkey(KeyCode - 111).ScrollPosition
                sclMagic_Change
                ListIndex = ListIndex - sclMagic.value
                DrawMagicList
            End Select
            picList_DblClick
        End If
    Case vbKeyShift
        keyShift = False
    Case vbKeyLeft
        keyLeft = False
    Case vbKeyRight
        keyRight = False
    Case vbKeyUp
        keyUp = False
    Case vbKeyDown
        keyDown = False
    Case vbKeyControl
        keyCtrl = False
    Case vbKeyAlt
        keyAlt = False
    Case vbKeyEscape
        keyEscape = False
    End Select
End Sub

Private Sub Form_Load()
    frmMain_Loaded = True

    CommonDialog.InitDir = App.Path
    CommonDialog.Filter = "Map (*.map4)|*.map4"

    Dim File As String
    Dim FileByteArray() As Byte

    File = "interface.rsc"
    FileByteArray() = StrConv(File, vbFromUnicode)
    ReDim Preserve FileByteArray(UBound(FileByteArray) + 1)

    EncryptDataFile FileByteArray(0), FileLen(File) Mod 87 + 5
    frmMain.Picture = LoadPicture(File)
    EncryptDataFile FileByteArray(0), FileLen(File) Mod 87 + 5


    File = "stats.rsc"
    FileByteArray() = StrConv(File, vbFromUnicode)
    ReDim Preserve FileByteArray(UBound(FileByteArray) + 1)

    EncryptDataFile FileByteArray(0), FileLen(File) Mod 87 + 5
    frmMain.picStats.Picture = LoadPicture(File)
    EncryptDataFile FileByteArray(0), FileLen(File) Mod 87 + 5

    InitializeLighting
    OutdoorLight = 150

    UnloadDirectDraw
    InitDirectDraw
    LoadSurfaces
End Sub

Private Sub Form_Resize()
    If Me.WindowState <> 1 Then
        picChatContainer.Height = Me.ScaleHeight - picChatContainer.Top - 5
        picChatContainer.Width = Me.ScaleWidth - picChatContainer.Left - 5
        picChat.Width = picChatContainer.ScaleWidth
        If picChat.Height - picChatContainer.ScaleHeight - ChatScrollBack < 0 Then
            ChatScrollBack = picChat.Height - picChatContainer.ScaleHeight
        End If
        picChat.Top = -(picChat.Height - picChatContainer.ScaleHeight - ChatScrollBack)
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmMain_Loaded = False
    frmMain_Showing = False
End Sub

Private Sub lblEditMode_MouseDown(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblEditMode(index).BackColor = &H61514B
End Sub

Private Sub lblEditMode_MouseUp(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim A As Byte, B As Byte, C As Long, D As Long

    If index <= 7 Then
        EditMode = index
        For A = 0 To 7
            If A <> index Then
                lblEditMode(A).BackColor = &H44342E
            End If
            If A >= 6 Then RedrawMap = True
            If A = 7 Then
                Select Case CurAtt
                Case 2, 3, 7, 8, 9, 17, 19, 20, 21    'Warp, Key, Object, Touch Plate, Damage, Directional Wall, Light, Dampening Wall, Object Picture
                    CurAtt = 0
                Case Else
                    lblCurTile = CurAtt
                End Select
                RedrawTile
            End If
        Next A
    End If
    Select Case index
    Case 8    'Fill
        lblEditMode(index).BackColor = &H44342E
        For A = 0 To 11
            For B = 0 To 11
                With EditMap.Tile(A, B)
                    Select Case EditMode
                    Case 0    'Ground
                        .Ground = CurTile
                    Case 1    'Ground2
                        .Ground2 = CurTile
                    Case 2    'BGTile1
                        .BGTile1 = CurTile
                    Case 3    'BGTile2
                        .BGTile2 = CurTile
                    Case 4    'FGTile
                        .FGTile = CurTile
                    Case 5    'FGTile2
                        .FGTile2 = CurTile
                    Case 6    'Attribute
                        .Att = CurAtt
                        .AttData(0) = CurAttData(0)
                        .AttData(1) = CurAttData(1)
                        .AttData(2) = CurAttData(2)
                        .AttData(3) = CurAttData(3)
                    End Select
                    RedrawMapTile A, B
                End With
            Next B
        Next A
    Case 9    'Sand
        lblEditMode(index).BackColor = &H44342E
        For A = 0 To 11
            For B = 0 To 11
                With EditMap.Tile(A, B)
                    If .Ground = 30 Or .Ground = 1397 Or .Ground = 1404 Or .Ground = 1390 Then
                        .Ground = 30
                        D = Int(Rnd * 3)
                        Select Case D
                        Case 0

                        Case 1

                        Case 2

                        End Select


                        RedrawMapTile A, B
                    End If
                End With
            Next B
        Next A
    Case 10    'Trees
        lblEditMode(index).BackColor = &H44342E
        For A = 0 To 11
            For B = 0 To 11
                Randomize
                If EditMap.Tile(A, B).BGTile1 = 63 Or EditMap.Tile(A, B).BGTile1 = 504 Or EditMap.Tile(A, B).BGTile1 = 693 Or EditMap.Tile(A, B).BGTile1 = 1505 Or EditMap.Tile(A, B).BGTile1 = 1499 Or EditMap.Tile(A, B).BGTile1 = 1302 Or EditMap.Tile(A, B).BGTile1 = 1501 Or EditMap.Tile(A, B).BGTile1 = 692 Or EditMap.Tile(A, B).BGTile1 = 691 Or EditMap.Tile(A, B).BGTile1 = 673 Or EditMap.Tile(A, B).BGTile1 = 659 Then
                    C = Int(Rnd * 11)
                    Select Case C
                    Case 0
                        EditMap.Tile(A, B).BGTile1 = 63
                        If B - 1 >= 0 Then EditMap.Tile(A, B - 1).FGTile = 56
                    Case 1
                        EditMap.Tile(A, B).BGTile1 = 504
                        If B - 1 >= 0 Then EditMap.Tile(A, B - 1).FGTile = 497
                    Case 2
                        EditMap.Tile(A, B).BGTile1 = 693
                        If B - 1 >= 0 Then EditMap.Tile(A, B - 1).FGTile = 686
                    Case 3
                        EditMap.Tile(A, B).BGTile1 = 1505
                        If B - 1 >= 0 Then EditMap.Tile(A, B - 1).FGTile = 1498
                    Case 4
                        EditMap.Tile(A, B).BGTile1 = 1499
                        If B - 1 >= 0 Then EditMap.Tile(A, B - 1).FGTile = 1492
                    Case 5
                        EditMap.Tile(A, B).BGTile1 = 63
                        If B - 1 >= 0 Then EditMap.Tile(A, B - 1).FGTile = 56
                    Case 6
                        EditMap.Tile(A, B).BGTile1 = 659
                        If B - 1 >= 0 Then EditMap.Tile(A, B - 1).FGTile = 652
                    Case 7
                        EditMap.Tile(A, B).BGTile1 = 673
                        If B - 1 >= 0 Then EditMap.Tile(A, B - 1).FGTile = 666
                    Case 8
                        EditMap.Tile(A, B).BGTile1 = 691
                        If B - 1 >= 0 Then EditMap.Tile(A, B - 1).FGTile = 684
                    Case 9
                        EditMap.Tile(A, B).BGTile1 = 1302
                        If B - 1 >= 0 Then EditMap.Tile(A, B - 1).FGTile = 1295
                    Case 10
                        EditMap.Tile(A, B).BGTile1 = 692
                        If B - 1 >= 0 Then EditMap.Tile(A, B - 1).FGTile = 685
                    End Select
                    RedrawMapTile A, B
                    If B - 1 >= 0 Then RedrawMapTile A, B - 1
                End If
                If B = 11 Then
                    If EditMap.Tile(A, B).FGTile = 56 Or EditMap.Tile(A, B).FGTile = 497 Or EditMap.Tile(A, B).FGTile = 686 Or EditMap.Tile(A, B).FGTile = 1498 Or EditMap.Tile(A, B).FGTile = 1492 Then
                        C = Int(Rnd * 2)
                        Select Case C
                        Case 0
                            EditMap.Tile(A, B).FGTile = 56
                        Case 1
                            EditMap.Tile(A, B).FGTile = 497
                        End Select
                        RedrawMapTile A, B
                    End If
                End If
            Next B
        Next A
    Case 11    'Grass
        lblEditMode(index).BackColor = &H44342E
        For A = 0 To 11
            For B = 0 To 11
                With EditMap.Tile(A, B)
                    If .Ground = 33 Then
                        Randomize
                        D = Int(Rnd * 5)
                        Randomize
                        If Not D = 4 Then
                            C = 568 + Int(Rnd * 18)
                        Else
                            C = 586 + Int(Rnd * 10)
                        End If
                        .Ground = C
                        RedrawMapTile A, B
                    ElseIf .Ground > 567 And .Ground < 596 Then
                        Randomize
                        D = Int(Rnd * 5)
                        Randomize
                        If Not D = 4 Then
                            C = 568 + Int(Rnd * 18)
                        Else
                            C = 586 + Int(Rnd * 10)
                        End If
                        .Ground = C
                        RedrawMapTile A, B
                    End If
                End With
            Next B
        Next A
    End Select
    RedrawTiles
    RedrawTile
End Sub

Private Sub lblMenu_MouseDown(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblMenu(index).BackColor = QBColor(15)
End Sub

Private Sub lblMenu_MouseUp(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim A As Long
    lblMenu(index).BackColor = &H44342E
    Select Case index
    Case 0    'Options
        ChatString = "/Options"
        Form_KeyPress 13
    Case 1    'Quit
        If Character.IsDead = True Then
            PrintChat "You are dead!", YELLOW
            Exit Sub
        End If
        If HasFullStats = False Then
            PrintChat "You must recover first!", YELLOW
            Exit Sub
        End If
        CloseClientSocket 0
        Character.Projectile = False
        Character.Ammo = 0
    Case 2    'MapEdit/Up
        If TopY > 0 Then
            TopY = TopY - 32
            RedrawTiles
        End If
    Case 3    'MapEdit/Down
        If TopY < 32000 Then
            TopY = TopY + 32
            RedrawTiles
        End If
    Case 4    'MapEdit/Upload
        UploadMap
        CloseMapEdit
    Case 5    'MapEdit/Cancel
        If MsgBox("Changes will be lost, continue?", vbYesNo + vbQuestion, TitleString) = vbYes Then
            CloseMapEdit
        End If
    Case 6    '
    Case 7    '
    Case 8    '
    Case 9    '
    Case 10    '
    Case 11    '
    Case 12    '
    Case 13    '
    Case 14    '
    Case 15    '
    Case 16    'Repair/Repair All
        SendSocket Chr$(65) + Chr$(3)
        picRepair.Visible = False
    Case 17    'Close Item Bank
        picBank.Visible = False
    Case 18    'Minimize
        Me.WindowState = 1
    Case 19    'Sell Button
        If CurInvObj >= 1 And CurInvObj <= 20 Then
            If Character.Inv(CurInvObj).Object > 0 Then
                If Object(Character.Inv(CurInvObj).Object).SellPrice > 0 Then
                    SendSocket Chr$(97) + Chr$(1) + Chr$(CurInvObj)
                Else
                    PrintChat "This object cannot be sold!", 14
                End If
            Else
                PrintChat "Select an object from your inventory to sell", 14
            End If
        Else
            PrintChat "Select an object from your inventory to sell", 14
        End If
    Case 20    'Close Newbie Help
        picHelp.Visible = False
    Case 21    'Magic Window
        picSkills.Visible = True
        DrawMagicList
    Case 22    'Cancel Sell Item
        picSellObject.Visible = False
    Case 23    'Skills Window
        picSkills.Visible = True
        DrawSkillsList
    Case 24    'Inventory Button
        picSkills.Visible = False
        DrawInterfaceLights
    Case 25    'Players Button
        ChatString = "/Who"
        Form_KeyPress 13
    Case 26    '

    Case 27    '

    Case 28    '

    Case 29    '

    Case 30    '

    Case 31    '

    Case 32    '

    Case 33    '

    Case 34    'Drop/Cancel
        picDrop.Visible = False
        frmMain.KeyPreview = True
        txtDrop.Text = vbNullString
    Case 35    'Drop/Ok
        If TempVar1 > 0 Then
            If lblDrop.Caption = "Deposit how much?" Then
                SendSocket Chr$(55) + Chr$(2) + QuadChar(TempVar1)
                picDrop.Visible = False
                frmMain.KeyPreview = True
                txtDrop.Text = vbNullString
            ElseIf lblDrop.Caption = "Deposit how many?" Then
                SendSocket Chr$(55) + Chr$(5) + Chr$(TempVar3) + QuadChar$(TempVar1)
                picDrop.Visible = False
                frmMain.KeyPreview = True
                txtDrop.Text = vbNullString
            ElseIf lblDrop.Caption = "Withdraw how much?" Then
                SendSocket Chr$(55) + Chr$(4) + QuadChar(TempVar1)
                picDrop.Visible = False
                frmMain.KeyPreview = True
                txtDrop.Text = vbNullString
            ElseIf lblDrop.Caption = "Withdraw how many?" Then
                SendSocket Chr$(55) + Chr$(6) + Chr$(TempVar3) + QuadChar(TempVar1)
                picDrop.Visible = False
                frmMain.KeyPreview = True
                txtDrop.Text = vbNullString
            Else
                SendSocket Chr$(9) + Chr$(TempVar3) + QuadChar(TempVar1)
                picDrop.Visible = False
                frmMain.KeyPreview = True
                txtDrop.Text = vbNullString
            End If
        End If
    Case 36, 37, 38, 39, 40, 41, 42, 43, 44, 45    'Buy/Buy
        With NPC(Map.NPC).SaleItem(index - 36)
            If .GiveObject >= 1 And .TakeObject >= 1 Then
                SendSocket Chr$(53) + Chr$(index - 36)
            End If
        End With
    Case 46    'Buy/Close
        picBuy.Visible = False
    Case 47    'Load Saved Map
        If MsgBox("If you load a new map, the current map will be lost if you have not saved it.  Do you wish to continue?", vbYesNo + vbQuestion, TitleString) = vbYes Then
            CommonDialog.ShowOpen
            If CommonDialog.filename = "" Then

            Else
                OpenMap CommonDialog.filename
            End If
        End If
    Case 48    'Stats
        ChatString = "/Stats"
        Form_KeyPress 13
    Case 49    'Trade
        ChatString = "/Buy"
        Form_KeyPress 13
    Case 50    'Sell
        ChatString = "/Sell"
        Form_KeyPress 13
    Case 51    'Shortcut/Trade
        ChatString = "/Bank"
        Form_KeyPress 13
    Case 52    'Repair
        ChatString = "/Repair"
        Form_KeyPress 13
    Case 53    'Guilds
        ChatString = "/Guilds"
        Form_KeyPress 13
    Case 54    'MapEdit/Copy
        CopyMap ClipboardMap, EditMap
    Case 55    'MapEdit/Paste
        If MsgBox("This will overwrite your current map-- are you sure you wish to paste?", vbYesNo + vbQuestion, TitleString) = vbYes Then
            CopyMap EditMap, ClipboardMap
            RedrawMap = True
        End If
    Case 56    'Repair Button
        SendSocket Chr$(65) + Chr$(2) + Chr$(CurInvObj)
    Case 57    'Repair Cancel
        picRepair.Visible = False
    Case 58    'Map Properties
        lblMenu(index).BackColor = &H44342E
        If frmMapProperties_Loaded = False Then Load frmMapProperties
        With EditMap
            frmMapProperties.Caption = "The Odyssey Online Classic [Map " + CStr(CMap) + " Properties]"
            frmMapProperties.txtName = EditMap.name
            frmMapProperties.sclMIDI = .MIDI
            frmMapProperties.txtUp = CStr(.ExitUp)
            frmMapProperties.txtDown = CStr(.ExitDown)
            frmMapProperties.txtLeft = CStr(.ExitLeft)
            frmMapProperties.txtRight = CStr(.ExitRight)
            frmMapProperties.txtBootMap = CStr(.BootLocation.Map)
            frmMapProperties.txtBootX = CStr(.BootLocation.X)
            frmMapProperties.txtBootY = CStr(.BootLocation.Y)
            frmMapProperties.txtDeathMap = CStr(.DeathLocation.Map)
            frmMapProperties.txtDeathX = CStr(.DeathLocation.X)
            frmMapProperties.txtDeathY = CStr(.DeathLocation.Y)
            frmMapProperties.cmbNPC.ListIndex = .NPC
            For A = 0 To 9
                frmMapProperties.cmbMonster(A).ListIndex = .MonsterSpawn(A).Monster
                frmMapProperties.sclRate(A) = .MonsterSpawn(A).Rate
            Next A
            For A = 0 To 7
                If ExamineBit(.flags, CByte(A)) Then
                    frmMapProperties.chkFlag(A) = 1
                Else
                    frmMapProperties.chkFlag(A) = 0
                End If
                If ExamineBit(.Flags2, CByte(A)) Then
                    frmMapProperties.chkFlag2(A) = 1
                Else
                    frmMapProperties.chkFlag2(A) = 0
                End If
            Next A
        End With
        frmMapProperties.Show 1
    Case 59    'Sell All
        If CurInvObj >= 1 And CurInvObj <= 20 Then
            If Character.Inv(CurInvObj).Object > 0 Then
                If Object(Character.Inv(CurInvObj).Object).SellPrice > 0 Then
                    SendSocket Chr$(97) + Chr$(2) + Chr$(CurInvObj)
                Else
                    PrintChat "This object cannot be sold!", 14
                End If
            Else
                PrintChat "Select an object from your inventory to sell", 14
            End If
        Else
            PrintChat "Select an object from your inventory to sell", 14
        End If
    Case 60 'Save Map
        CommonDialog.ShowSave
        If CommonDialog.filename = "" Then

        Else
            SaveMap CommonDialog.filename
        End If
    End Select
End Sub

Private Sub picInventory_DblClick()
    If Character.IsDead = True Then
        Exit Sub
    End If

    If CurInvObj > 0 And picRepair.Visible = False And picSellObject.Visible = False Then
        If picBank.Visible = True And CurInvObj <= 20 Then
            With Character.Inv(CurInvObj)
                If .Object > 0 Then
                    If Not Object(.Object).Type = 6 And Not Object(.Object).Type = 11 Then
                        SendSocket Chr$(55) & Chr$(1) & Chr$(CurInvObj)
                    Else
                        If .Object = World.ObjMoney Then    'Gold
                            frmMain.KeyPreview = False
                            lblDrop.Caption = "Deposit how much?"
                            TempVar1 = 0
                            TempVar2 = Character.Inv(CurInvObj).value
                            TempVar3 = CurInvObj
                            txtDrop.Text = TempVar2
                            TempVar1 = TempVar2
                            picDrop.Visible = True
                            picBuy.Visible = False
                        Else
                            frmMain.KeyPreview = False
                            lblDrop.Caption = "Deposit how many?"
                            TempVar1 = 0
                            TempVar2 = Character.Inv(CurInvObj).value
                            TempVar3 = CurInvObj
                            txtDrop.Text = TempVar2
                            TempVar1 = TempVar2
                            picDrop.Visible = True
                            picBuy.Visible = False
                        End If
                    End If
                Else
                    PrintChat "No such object.", 7
                End If
            End With
        Else
            If CurInvObj <= 20 Then
                With Character.Inv(CurInvObj)
                    If .EquippedNum > 0 Then
                        If HasFullStats = False Then
                            PrintChat "You must recover first!", YELLOW
                            Exit Sub
                        End If
                        SendSocket Chr$(11) + Chr$(6)    'Stop Using Obj
                    Else
                        If Not Character.Inv(CurInvObj).Object = 0 Then
                            Select Case Object(CLng(Character.Inv(CurInvObj).Object)).Type
                            Case 1, 2, 3, 4
                                If Character.EquippedObject(Object(CLng(Character.Inv(CurInvObj).Object)).Type).Object = 0 Then
                                    SendSocket Chr$(10) + Chr$(CurInvObj)    'Use Obj
                                Else
                                    If HasFullStats Then
                                        SendSocket Chr$(10) + Chr$(CurInvObj)    'Use Obj
                                    Else
                                        PrintChat "You must recover first!", YELLOW
                                    End If
                                End If
                            Case 8
                                If Character.EquippedObject(5).Object = 0 Then
                                    SendSocket Chr$(10) + Chr$(CurInvObj)    'Use Obj
                                Else
                                    If HasFullStats Then
                                        SendSocket Chr$(10) + Chr$(CurInvObj)    'Use Obj
                                    Else
                                        PrintChat "You must recover first!", YELLOW
                                    End If
                                End If
                            Case Else
                                SendSocket Chr$(10) + Chr$(CurInvObj)    'Use Obj
                            End Select
                        End If
                    End If
                End With
            Else
                If Character.EquippedObject(CurInvObj - 20).Object > 0 Then
                    If HasFullStats = False Then
                        PrintChat "You must recover first!", YELLOW
                        Exit Sub
                    End If
                    SendSocket Chr$(11) + Chr$(CurInvObj - 20)    'Stop Using Object
                End If
            End If
        End If
    End If
End Sub

Private Sub picInventory_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Character.IsDead = True Then
        Exit Sub
    End If

    Dim OldInvObj As Long
    OldInvObj = CurInvObj
    CurInvObj = Int(X / 36) + Int(Y / 36) * 5 + 1
    If CurInvObj < 1 Then CurInvObj = 1
    If CurInvObj > 25 Then CurInvObj = 25
    If picRepair.Visible = True Then
        If CurInvObj <> OldInvObj Then
            DisplayRepair
            RefreshInventory
        End If
    ElseIf picSellObject.Visible = True Then
        If CurInvObj <> OldInvObj Then
            DisplaySell
            RefreshInventory
        End If
    Else
        If CurInvObj <> OldInvObj Then
            RefreshInventory
        End If
    End If

    If Button = 2 And CurInvObj <= 20 Then
        If Character.Inv(CurInvObj).Object > 0 Then
            If Object(Character.Inv(CurInvObj).Object).Type = 6 Or Object(Character.Inv(CurInvObj).Object).Type = 11 Then
                'Money
                frmMain.KeyPreview = False
                lblDrop.Caption = "Drop how much?"
                TempVar1 = 0
                TempVar2 = Character.Inv(CurInvObj).value
                TempVar3 = CurInvObj
                txtDrop.Text = TempVar2
                TempVar1 = TempVar2
                picDrop.Visible = True
                picBuy.Visible = False
            Else
                SendSocket Chr$(9) + Chr$(CurInvObj) + QuadChar(0)
            End If
        End If
    End If
End Sub

Private Sub picTile_Click()
    CurTile = 0
    BitBlt picTile.hDC, 0, 0, 32, 32, 0, 0, 0, BLACKNESS
    picTile.Refresh
End Sub

Private Sub picTiles_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        If EditMode < 6 Then
            CurTile = Int((Y + TopY) / 32) * 7 + Int(X / 32) + 1
            lblCurTile = CurTile
            RedrawTile
        Else
            If EditMode = 6 Then
                NewAtt = Int(Y / 32) * 7 + Int(X / 32) + 1
                If NewAtt <= MaxAtt Then
                    Select Case NewAtt
                    Case 2, 3, 7, 8, 9, 17, 19, 20, 21    'Warp, Key, Object, Touch Plate, Damage, Directional Wall, Light, Dampening Wall, Object Picture
                        frmMapAtt.Show 1
                    Case Else
                        CurAtt = NewAtt
                        lblCurTile = NewAtt
                    End Select
                    RedrawTile
                End If
            Else
                NewAtt = Int(Y / 32) * 7 + Int(X / 32) + 1
                If NewAtt <= MaxAtt Then
                    Select Case NewAtt
                    Case 2, 3, 7, 8, 9, 17, 19, 20, 21    'Warp, Key, Object, Touch Plate, Damage, Directional Wall, Light, Dampening Wall, Object Picture

                    Case Else
                        CurAtt = NewAtt
                        lblCurTile = NewAtt
                    End Select
                    RedrawTile
                End If
            End If
        End If
    End If
End Sub

Private Sub picViewport_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim A As Long, MapX As Byte, MapY As Byte
    Dim St As String
    If X < 0 Or Y < 0 Then Exit Sub
    MapX = Int(X / 32)
    MapY = Int(Y / 32)
    If MapEdit = True Then
        If MapX >= 0 And MapX <= 11 And MapY >= 0 And MapY <= 11 Then
            If Button = 1 Then
                If keyCtrl = False Then
                    Select Case EditMode
                    Case 0    'Ground
                        EditMap.Tile(MapX, MapY).Ground = CurTile
                        RedrawMapTile MapX, MapY
                    Case 1    'Ground2
                        EditMap.Tile(MapX, MapY).Ground2 = CurTile
                        RedrawMapTile MapX, MapY
                    Case 2    'BGTile1
                        EditMap.Tile(MapX, MapY).BGTile1 = CurTile
                        RedrawMapTile MapX, MapY
                    Case 3    'BGTile2
                        EditMap.Tile(MapX, MapY).BGTile2 = CurTile
                        RedrawMapTile MapX, MapY
                    Case 4    'FGTile
                        EditMap.Tile(MapX, MapY).FGTile = CurTile
                        RedrawMapTile MapX, MapY
                    Case 5    'FGTile2
                        EditMap.Tile(MapX, MapY).FGTile2 = CurTile
                        RedrawMapTile MapX, MapY
                    Case 6    'Att1
                        EditMap.Tile(MapX, MapY).Att = CurAtt
                        EditMap.Tile(MapX, MapY).AttData(0) = CurAttData(0)
                        EditMap.Tile(MapX, MapY).AttData(1) = CurAttData(1)
                        EditMap.Tile(MapX, MapY).AttData(2) = CurAttData(2)
                        EditMap.Tile(MapX, MapY).AttData(3) = CurAttData(3)
                        RedrawMapTile MapX, MapY
                    Case 7    'Att2
                        EditMap.Tile(MapX, MapY).Att2 = CurAtt
                        RedrawMapTile MapX, MapY
                    End Select
                Else
                    Select Case EditMode
                    Case 0    'Ground
                        CurTile = EditMap.Tile(MapX, MapY).Ground
                        lblCurTile = CurTile
                        RedrawTile
                    Case 1    'Ground2
                        CurTile = EditMap.Tile(MapX, MapY).Ground2
                        lblCurTile = CurTile
                        RedrawTile
                    Case 2    'BGTile1
                        CurTile = EditMap.Tile(MapX, MapY).BGTile1
                        RedrawTile
                    Case 3    'BGTile2
                        CurTile = EditMap.Tile(MapX, MapY).BGTile2
                        RedrawTile
                    Case 4    'FGTile
                        CurTile = EditMap.Tile(MapX, MapY).FGTile
                        RedrawTile
                    Case 5    'FGTile2
                        CurTile = EditMap.Tile(MapX, MapY).FGTile2
                        RedrawTile
                    Case 6    'Attribute
                        NewAtt = EditMap.Tile(MapX, MapY).Att
                        CurAtt = NewAtt

                        CurAttData(0) = EditMap.Tile(MapX, MapY).AttData(0)
                        CurAttData(1) = EditMap.Tile(MapX, MapY).AttData(1)
                        CurAttData(2) = EditMap.Tile(MapX, MapY).AttData(2)
                        CurAttData(3) = EditMap.Tile(MapX, MapY).AttData(3)
                        RedrawTile

                        keyCtrl = False
                        Select Case NewAtt
                        Case 2, 3, 7, 8, 9, 17, 19, 20, 21    'Warp, Key, Object, Touch Plate, Damage, Directional Wall, Light, Dampening Wall, Object Picture
                            Load frmMapAtt
                            frmMapAtt.Show
                        Case Else
                            CurAtt = EditMap.Tile(MapX, MapY).Att
                        End Select
                    Case 7    'Att2
                        CurAtt = EditMap.Tile(MapX, MapY).Att2
                        RedrawTile
                    End Select
                End If
            ElseIf Button = 2 Then
                Select Case EditMode
                Case 0    'Ground
                    EditMap.Tile(MapX, MapY).Ground = 0
                    RedrawMapTile MapX, MapY
                Case 1    'Ground2
                    EditMap.Tile(MapX, MapY).Ground2 = 0
                    RedrawMapTile MapX, MapY
                Case 2    'BGTile1
                    EditMap.Tile(MapX, MapY).BGTile1 = 0
                    RedrawMapTile MapX, MapY
                Case 3    'BGTile1
                    EditMap.Tile(MapX, MapY).BGTile2 = 0
                    RedrawMapTile MapX, MapY
                Case 4    'FGTile
                    EditMap.Tile(MapX, MapY).FGTile = 0
                    RedrawMapTile MapX, MapY
                Case 5    'FGTile2
                    EditMap.Tile(MapX, MapY).FGTile2 = 0
                    RedrawMapTile MapX, MapY
                Case 6    'Attribute
                    EditMap.Tile(MapX, MapY).Att = 0
                    EditMap.Tile(MapX, MapY).AttData(0) = 0
                    EditMap.Tile(MapX, MapY).AttData(1) = 0
                    EditMap.Tile(MapX, MapY).AttData(2) = 0
                    EditMap.Tile(MapX, MapY).AttData(3) = 0
                    RedrawMapTile MapX, MapY
                Case 7    'Att2
                    EditMap.Tile(MapX, MapY).Att2 = 0
                    RedrawMapTile MapX, MapY
                End Select
            End If
        End If
    Else
        Dim FoundSomething As Boolean
        If Button = 1 Then
            For A = 1 To MaxUsers
                With Player(A)
                    If .Map = CMap And .X = MapX And .Y = MapY And Not .status = 9 And Not .status = 25 Then
                        If .Guild > 0 Then
                            PrintChat "You see " + .name + ", member of the guild " + Chr$(34) + Guild(.Guild).name + Chr$(34) + "!", 7
                        Else
                            PrintChat "You see " + .name + "!", 7
                        End If
                        SendSocket Chr$(27) + Chr$(A)
                        FoundSomething = True
                    End If
                End With
            Next A
            For A = 0 To MaxMonsters
                With Map.Monster(A)
                    If .Monster > 0 And .X = MapX And .Y = MapY Then
                        PrintChat "You see a " + Monster(.Monster).name + "!", 7
                        FoundSomething = True
                    End If
                End With
            Next A
            For A = 0 To MaxMapObjects
                If Map.Object(A).X = MapX And Map.Object(A).Y = MapY And Map.Object(A).Object > 0 Then
                    Select Case Object(Map.Object(A).Object).Type
                    Case 6, 11
                        PrintChat "You see " + CStr(Map.Object(A).value) + " " + Object(Map.Object(A).Object).name + "!", 7
                    Case 1, 2, 3, 4
                        St = ""
                        If Map.Object(A).ItemPrefix > 0 Then
                            St = ItemPrefix(Map.Object(A).ItemPrefix).name + " "
                        End If
                        If Map.Object(A).ItemSuffix > 0 Then
                            St = St + Object(Map.Object(A).Object).name + " " + ItemSuffix(Map.Object(A).ItemSuffix).name + "!"
                        Else
                            St = St + Object(Map.Object(A).Object).name + "!"
                        End If
                        St = "You see a " + St + "  Condition: " + GetDurStringFromValue(Map.Object(A).value, Object(Map.Object(A).Object).MaxDur)
                        PrintChat St, 7
                    Case 7    'Key
                        St = ""
                        If Map.Object(A).ItemPrefix > 0 Then
                            St = ItemPrefix(Map.Object(A).ItemPrefix).name + " "
                        End If
                        If Map.Object(A).ItemSuffix > 0 Then
                            St = St + Object(Map.Object(A).Object).name + " " + ItemSuffix(Map.Object(A).ItemSuffix).name + "!"
                        Else
                            St = St + Object(Map.Object(A).Object).name + "!"
                        End If
                        St = "You see a " + St + "!"
                        PrintChat St, 7
                    Case Else
                        PrintChat "You see a " + Object(Map.Object(A).Object).name + "!", 7
                    End Select
                    FoundSomething = True
                End If
            Next A
            If Map.Tile(MapX, MapY).Att = 12 Or Map.Tile(MapX, MapY).Att2 = 12 Then
                SendSocket Chr$(91) + Chr$(MapX) + Chr$(MapY)
                FoundSomething = True
            End If
            If FoundSomething = False Then
                PrintChat "You see nothing special.", 7
            End If
        End If
    End If
End Sub

Private Sub picViewport_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If MapEdit = True Then
        picViewport_MouseDown Button, Shift, X, Y
    End If
End Sub

Private Sub sclMagic_Change()
    DrawMagicList
    picChat.SetFocus
End Sub

Private Sub txtDrop_Change()
    If Val(txtDrop.Text) > TempVar2 Then txtDrop.Text = TempVar2
    TempVar1 = Val(txtDrop)
End Sub

Private Sub picList_DblClick()
    Select Case lblSkillType.Caption
    Case "Skills"
        UseSkill ListIndex
    Case "Magic"
        UseMagic ListIndex
    End Select
End Sub

Private Sub picList_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ListIndex = Int((Y - 1) / 26) + 1
    If ListIndex > 0 And ListIndex < 13 Then
        Select Case lblSkillType.Caption
        Case "Skills"
            Select Case ListIndex
            Case 1    'Fishing
                lblCurObj.Caption = "Fishing"
                SetObjectInfo "Level " & Character.Skill(1).Level & vbCrLf & "Catch Fish or whatever else may be lurking in the water." & vbCrLf & "Press Esc + F1-F10 to hotkey this Skill"
            Case 2    'Mining
                lblCurObj.Caption = "Mining"
                SetObjectInfo "Level " & Character.Skill(2).Level & vbCrLf & "Dig beneath the earth to find Ore for blacksmithing." & vbCrLf & "Press Esc + F1-F10 to hotkey this Skill"
            Case 3    'Lumberjacking
                lblCurObj.Caption = "Lumberjacking"
                SetObjectInfo "Level " & Character.Skill(3).Level & vbCrLf & "Chop trees to get lumber." & vbCrLf & "Press Esc + F1-F10 to hotkey this Skill"
            Case 4    'Cooking
                If Character.Access > 0 Then
                    lblCurObj.Caption = "Cooking"
                    SetObjectInfo "Level " & Character.Skill(4).Level & vbCrLf & "Cook stuff" & vbCrLf & "Press Esc + F1-F10 to hotkey this Skill"
                End If
            Case 5    'Enchanting
                If Character.Access > 0 Then
                    lblCurObj.Caption = "Enchanting"
                    SetObjectInfo "Level " & Character.Skill(5).Level & vbCrLf & "Enchant stuff" & vbCrLf & "Press Esc + F1-F10 to hotkey this Skill"
                End If
            Case 6    'Smithing
                If Character.Access > 0 Then
                    lblCurObj.Caption = "Smithing"
                    SetObjectInfo "Level " & Character.Skill(6).Level & vbCrLf & "Smith stuff" & vbCrLf & "Press Esc + F1-F10 to hotkey this Skill"
                End If
            End Select
        Case "Magic"
            Dim A As Long, B As Long
            B = -(sclMagic.value)
            For A = 1 To MaxMagic
                If ExamineBit(Magic(A).Class, Character.Class - 1) And Character.Level >= Magic(A).Level Then
                    If Not (Magic(A).MagicLevel = 0 And Character.Access = 0 And ServerPort = 5752) Then
                        B = B + 1
                        If B > 0 Then
                            If B = ListIndex Then
                                lblCurObj.Caption = Magic(A).name
                                SetObjectInfo "Level " + CStr(Magic(A).Level) + vbCrLf + Magic(A).Description + vbCrLf + "Press Esc + F1-F10 to hotkey this spell"
                            End If
                        End If
                    End If
                End If
            Next A
        End Select
    End If
End Sub

Sub DrawInterfaceLights()
    Dim A As Long, B As Long

    'Stats (Button 0)
    If Character.StatPoints > 0 Then
        DrawToDC 0, 0, 7, 7, picButtonLight(0).hDC, DDSInterfaceLights, 0, 0
    Else
        DrawToDC 0, 0, 7, 7, picButtonLight(0).hDC, DDSInterfaceLights, 7, 0
    End If
    picButtonLight(0).Refresh

    'Trade (Button 1)
    If Map.NPC > 0 Then
        B = 0
        For A = 0 To 9
            If NPC(Map.NPC).SaleItem(A).GiveObject > 0 Then B = 1
        Next A

        If B = 1 Then
            DrawToDC 0, 0, 7, 7, picButtonLight(1).hDC, DDSInterfaceLights, 0, 0
        Else
            DrawToDC 0, 0, 7, 7, picButtonLight(1).hDC, DDSInterfaceLights, 7, 0
        End If
    Else
        DrawToDC 0, 0, 7, 7, picButtonLight(1).hDC, DDSInterfaceLights, 7, 0
    End If
    picButtonLight(1).Refresh

    'Sell (Button 2)
    If Map.NPC > 0 Then
        If ExamineBit(NPC(Map.NPC).flags, 2) = True Then
            DrawToDC 0, 0, 7, 7, picButtonLight(2).hDC, DDSInterfaceLights, 0, 0
        Else
            DrawToDC 0, 0, 7, 7, picButtonLight(2).hDC, DDSInterfaceLights, 7, 0
        End If
    Else
        DrawToDC 0, 0, 7, 7, picButtonLight(2).hDC, DDSInterfaceLights, 7, 0
    End If
    picButtonLight(2).Refresh

    'Bank (Button 3)
    If Map.NPC > 0 Then
        If ExamineBit(NPC(Map.NPC).flags, 0) = True Then
            DrawToDC 0, 0, 7, 7, picButtonLight(3).hDC, DDSInterfaceLights, 0, 0
        Else
            DrawToDC 0, 0, 7, 7, picButtonLight(3).hDC, DDSInterfaceLights, 7, 0
        End If
    Else
        DrawToDC 0, 0, 7, 7, picButtonLight(3).hDC, DDSInterfaceLights, 7, 0
    End If
    picButtonLight(3).Refresh

    'Repair (Button 4)
    If Map.NPC > 0 Then
        If ExamineBit(NPC(Map.NPC).flags, 1) = True Then
            DrawToDC 0, 0, 7, 7, picButtonLight(4).hDC, DDSInterfaceLights, 0, 0
        Else
            DrawToDC 0, 0, 7, 7, picButtonLight(4).hDC, DDSInterfaceLights, 7, 0
        End If
    Else
        DrawToDC 0, 0, 7, 7, picButtonLight(4).hDC, DDSInterfaceLights, 7, 0
    End If
    picButtonLight(4).Refresh

    'Guilds (Button 5)
    DrawToDC 0, 0, 7, 7, picButtonLight(5).hDC, DDSInterfaceLights, 7, 0
    picButtonLight(5).Refresh

    'Macros (Button 6)
    DrawToDC 0, 0, 7, 7, picButtonLight(6).hDC, DDSInterfaceLights, 7, 0
    picButtonLight(6).Refresh


    'Inventory (Button 7)
    If picSkills.Visible = False Then
        DrawToDC 0, 0, 7, 7, picButtonLight(7).hDC, DDSInterfaceLights, 21, 0
    Else
        DrawToDC 0, 0, 7, 7, picButtonLight(7).hDC, DDSInterfaceLights, 14, 0
    End If
    picButtonLight(7).Refresh

    'Magic (Button 8), Skills (Button 9)
    If picSkills.Visible = True Then
        Select Case lblSkillType.Caption
        Case "Magic"
            DrawToDC 0, 0, 7, 7, picButtonLight(8).hDC, DDSInterfaceLights, 21, 0
            DrawToDC 0, 0, 7, 7, picButtonLight(9).hDC, DDSInterfaceLights, 14, 0
        Case "Skills"
            DrawToDC 0, 0, 7, 7, picButtonLight(9).hDC, DDSInterfaceLights, 21, 0
            DrawToDC 0, 0, 7, 7, picButtonLight(8).hDC, DDSInterfaceLights, 14, 0
        End Select
    Else
        DrawToDC 0, 0, 7, 7, picButtonLight(8).hDC, DDSInterfaceLights, 14, 0
        DrawToDC 0, 0, 7, 7, picButtonLight(9).hDC, DDSInterfaceLights, 14, 0
    End If
    picButtonLight(8).Refresh
    picButtonLight(9).Refresh

    'Options
    DrawToDC 0, 0, 7, 7, picButtonLight(10).hDC, DDSInterfaceLights, 14, 0
    picButtonLight(10).Refresh

End Sub

Sub ShowBuyMenu(NPCIndex As Integer)
    Dim A As Long, St1 As String
    For A = 0 To 9
        With NPC(NPCIndex).SaleItem(A)
            If .GiveObject >= 1 And .TakeObject >= 1 Then
                St1 = vbNullString
                Select Case Object(.GiveObject).Type
                Case 6    'Money
                    St1 = CStr(.GiveValue) + " " + Object(.GiveObject).name + " in exchange for "
                    If Object(.TakeObject).Type = 6 Or Object(.TakeObject).Type = 11 Then
                        St1 = St1 + CStr(.TakeValue) + " " + Object(.TakeObject).name & vbCrLf
                    Else
                        St1 = St1 + "1 " + Object(.TakeObject).name
                    End If
                Case 11    'Ammo
                    St1 = CStr(.GiveValue) + " +" & Object(.GiveObject).Modifier & " " + Object(.GiveObject).name & " (Ammunition)" + " in exchange for "
                    If Object(.TakeObject).Type = 6 Or Object(.TakeObject).Type = 11 Then
                        St1 = St1 + CStr(.TakeValue) + " " + Object(.TakeObject).name & vbCrLf
                    Else
                        St1 = St1 + "1 " + Object(.TakeObject).name
                    End If
                Case 1, 2, 3, 4, 10    'Weapon
                    Dim St2 As String
                    St1 = "+" & Object(.GiveObject).Modifier & " " & Object(.GiveObject).name + " in exchange for "
                    If Object(.TakeObject).Type = 6 Or Object(.TakeObject).Type = 11 Then
                        St1 = St1 + CStr(.TakeValue) + " " + Object(.TakeObject).name & vbCrLf & "Requirements:  "
                    Else
                        St1 = St1 + "1 " + Object(.TakeObject).name & vbCrLf & "Requirements:  "
                    End If
                    If Object(.GiveObject).LevelReq > 0 Then St2 = St2 & "Level " & Object(.GiveObject).LevelReq
                    If St2 = vbNullString Then St1 = St1 & "<None>" Else St1 = St1 & St2
                    St2 = vbNullString
                Case 8    'Ring
                    Select Case Object(.GiveObject).Data2
                    Case 0    'Damage
                        St1 = "+" & Object(.GiveObject).Modifier & " Damage " & Object(.GiveObject).name + " in exchange for "
                    Case 1    'Defense
                        St1 = "+" & Object(.GiveObject).Modifier & " Defense " & Object(.GiveObject).name + " in exchange for "
                    End Select
                    If Object(.TakeObject).Type = 6 Or Object(.TakeObject).Type = 11 Then
                        St1 = St1 + CStr(.TakeValue) + " " + Object(.TakeObject).name & vbCrLf & "Requirements:  "
                    Else
                        St1 = St1 + "1 " + Object(.TakeObject).name & vbCrLf & "Requirements:  "
                    End If
                    If Object(.GiveObject).LevelReq > 0 Then St2 = St2 & "Level " & Object(.GiveObject).LevelReq
                    If St2 = vbNullString Then St1 = St1 & "<None>" Else St1 = St1 & St2
                    St2 = vbNullString
                Case Else
                    St1 = "1 " + Object(.GiveObject).name + " in exchange for "
                    If Object(.TakeObject).Type = 6 Or Object(.TakeObject).Type = 11 Then
                        St1 = St1 + CStr(.TakeValue) + " " + Object(.TakeObject).name & vbCrLf
                    Else
                        St1 = St1 + "1 " + Object(.TakeObject).name
                    End If
                End Select
                frmMain.lblItem(A) = St1
                frmMain.GivObjPic(A).Cls
                DrawToDC 0, 0, 32, 32, frmMain.GivObjPic(A).hDC, DDSObjects, 0, (Object(.GiveObject).Picture - 1) * 32
                frmMain.lblShopName.Caption = Map.name
            Else
                frmMain.lblItem(A) = vbNullString
                frmMain.GivObjPic(A).Cls
            End If
        End With
    Next A
    frmMain.picBuy.Visible = True
    frmMain.picDrop.Visible = False
End Sub

