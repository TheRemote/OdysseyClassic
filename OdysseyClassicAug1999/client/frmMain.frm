VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   0  'None
   Caption         =   "The Odyssey Online Classic"
   ClientHeight    =   7200
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9600
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   480
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   640
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picRepair 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3495
      Left            =   3360
      ScaleHeight     =   3465
      ScaleWidth      =   5025
      TabIndex        =   109
      Top             =   1800
      Visible         =   0   'False
      Width           =   5055
      Begin VB.Frame fraObjInfo 
         BackColor       =   &H80000004&
         Caption         =   "Object Info"
         Height          =   2655
         Left            =   120
         TabIndex        =   112
         Top             =   120
         Width           =   2535
         Begin VB.PictureBox RepairObjPic 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   480
            Left            =   1440
            ScaleHeight     =   32
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   32
            TabIndex        =   118
            Top             =   360
            Width           =   480
         End
         Begin VB.Label lblObjCond 
            Alignment       =   2  'Center
            Height          =   255
            Left            =   840
            TabIndex        =   124
            Top             =   2160
            Width           =   1575
         End
         Begin VB.Label lblObjectCond 
            AutoSize        =   -1  'True
            Caption         =   "Condition:"
            Height          =   195
            Left            =   120
            TabIndex        =   123
            Top             =   2160
            Width           =   705
         End
         Begin VB.Label lblRepairCst 
            Alignment       =   2  'Center
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   840
            TabIndex        =   122
            Top             =   1800
            Width           =   1575
         End
         Begin VB.Label lblRepairCost 
            AutoSize        =   -1  'True
            Caption         =   "Cost:"
            Height          =   195
            Left            =   120
            TabIndex        =   121
            Top             =   1800
            Width           =   360
         End
         Begin VB.Label lblObjectPic 
            AutoSize        =   -1  'True
            Caption         =   "Picture:"
            Height          =   195
            Left            =   120
            TabIndex        =   117
            Top             =   480
            Width           =   540
         End
         Begin VB.Label lblObjDur 
            Alignment       =   2  'Center
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   840
            TabIndex        =   116
            Top             =   1440
            Width           =   1575
         End
         Begin VB.Label lblObjectDur 
            AutoSize        =   -1  'True
            Caption         =   "Durability:"
            Height          =   195
            Left            =   120
            TabIndex        =   115
            Top             =   1440
            Width           =   690
         End
         Begin VB.Label lblObjName 
            Alignment       =   2  'Center
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   840
            TabIndex        =   114
            Top             =   1080
            Width           =   1575
         End
         Begin VB.Label lblObjectName 
            AutoSize        =   -1  'True
            Caption         =   "Name:"
            Height          =   195
            Left            =   120
            TabIndex        =   113
            Top             =   1080
            Width           =   465
         End
      End
      Begin VB.Label lblRepairNpcTalk 
         Height          =   2295
         Left            =   2760
         TabIndex        =   120
         Top             =   480
         Width           =   2175
      End
      Begin VB.Label lblRepairNPCName 
         Alignment       =   2  'Center
         Height          =   255
         Left            =   2760
         TabIndex        =   119
         Top             =   120
         Width           =   2175
      End
      Begin VB.Label lblMenu 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
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
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   57
         Left            =   3120
         TabIndex        =   111
         Top             =   3000
         Width           =   735
      End
      Begin VB.Label lblMenu 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
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
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   56
         Left            =   720
         TabIndex        =   110
         Top             =   3000
         Width           =   735
      End
   End
   Begin VB.PictureBox picBuy 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   6135
      Left            =   2760
      ScaleHeight     =   6105
      ScaleWidth      =   6225
      TabIndex        =   76
      Top             =   360
      Visible         =   0   'False
      Width           =   6255
      Begin VB.PictureBox GivObjPic 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   0
         Left            =   120
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   86
         Top             =   600
         Width           =   480
      End
      Begin VB.PictureBox GivObjPic 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   1
         Left            =   120
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   85
         Top             =   1080
         Width           =   480
      End
      Begin VB.PictureBox GivObjPic 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   2
         Left            =   120
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   84
         Top             =   1560
         Width           =   480
      End
      Begin VB.PictureBox GivObjPic 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   3
         Left            =   120
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   83
         Top             =   2040
         Width           =   480
      End
      Begin VB.PictureBox GivObjPic 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   4
         Left            =   120
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   82
         Top             =   2520
         Width           =   480
      End
      Begin VB.PictureBox GivObjPic 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   5
         Left            =   120
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   81
         Top             =   3000
         Width           =   480
      End
      Begin VB.PictureBox GivObjPic 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   6
         Left            =   120
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   80
         Top             =   3480
         Width           =   480
      End
      Begin VB.PictureBox GivObjPic 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   7
         Left            =   120
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   79
         Top             =   3960
         Width           =   480
      End
      Begin VB.PictureBox GivObjPic 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   8
         Left            =   120
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   78
         Top             =   4440
         Width           =   480
      End
      Begin VB.PictureBox GivObjPic 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   9
         Left            =   120
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   77
         Top             =   4920
         Width           =   480
      End
      Begin VB.Label lblShopName 
         Alignment       =   2  'Center
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
         Height          =   375
         Left            =   120
         TabIndex        =   108
         Top             =   120
         Width           =   5895
      End
      Begin VB.Label lblMenu 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
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
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   36
         Left            =   5280
         TabIndex        =   107
         Top             =   720
         Width           =   735
      End
      Begin VB.Label lblItem 
         Alignment       =   2  'Center
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
         Index           =   0
         Left            =   720
         TabIndex        =   106
         Top             =   720
         Width           =   4455
      End
      Begin VB.Label lblMenu 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
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
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   37
         Left            =   5280
         TabIndex        =   105
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label lblItem 
         Alignment       =   2  'Center
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
         Index           =   1
         Left            =   720
         TabIndex        =   104
         Top             =   1200
         Width           =   4455
      End
      Begin VB.Label lblMenu 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
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
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   38
         Left            =   5280
         TabIndex        =   103
         Top             =   1680
         Width           =   735
      End
      Begin VB.Label lblItem 
         Alignment       =   2  'Center
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
         Index           =   2
         Left            =   720
         TabIndex        =   102
         Top             =   1680
         Width           =   4455
      End
      Begin VB.Label lblMenu 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
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
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   39
         Left            =   5280
         TabIndex        =   101
         Top             =   2160
         Width           =   735
      End
      Begin VB.Label lblItem 
         Alignment       =   2  'Center
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
         Index           =   3
         Left            =   720
         TabIndex        =   100
         Top             =   2160
         Width           =   4455
      End
      Begin VB.Label lblMenu 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
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
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   40
         Left            =   5280
         TabIndex        =   99
         Top             =   2640
         Width           =   735
      End
      Begin VB.Label lblItem 
         Alignment       =   2  'Center
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
         Index           =   4
         Left            =   720
         TabIndex        =   98
         Top             =   2640
         Width           =   4455
      End
      Begin VB.Label lblMenu 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
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
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   41
         Left            =   5280
         TabIndex        =   97
         Top             =   3120
         Width           =   735
      End
      Begin VB.Label lblItem 
         Alignment       =   2  'Center
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
         Index           =   5
         Left            =   720
         TabIndex        =   96
         Top             =   3120
         Width           =   4455
      End
      Begin VB.Label lblMenu 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
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
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   42
         Left            =   5280
         TabIndex        =   95
         Top             =   3600
         Width           =   735
      End
      Begin VB.Label lblItem 
         Alignment       =   2  'Center
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
         Index           =   6
         Left            =   720
         TabIndex        =   94
         Top             =   3600
         Width           =   4455
      End
      Begin VB.Label lblMenu 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
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
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   43
         Left            =   5280
         TabIndex        =   93
         Top             =   4080
         Width           =   735
      End
      Begin VB.Label lblItem 
         Alignment       =   2  'Center
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
         Index           =   7
         Left            =   720
         TabIndex        =   92
         Top             =   4080
         Width           =   4455
      End
      Begin VB.Label lblMenu 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
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
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   44
         Left            =   5280
         TabIndex        =   91
         Top             =   4560
         Width           =   735
      End
      Begin VB.Label lblItem 
         Alignment       =   2  'Center
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
         Index           =   8
         Left            =   720
         TabIndex        =   90
         Top             =   4560
         Width           =   4455
      End
      Begin VB.Label lblMenu 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
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
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   45
         Left            =   5280
         TabIndex        =   89
         Top             =   5040
         Width           =   735
      End
      Begin VB.Label lblItem 
         Alignment       =   2  'Center
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
         Index           =   9
         Left            =   720
         TabIndex        =   88
         Top             =   5040
         Width           =   4455
      End
      Begin VB.Label lblMenu 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
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
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   46
         Left            =   120
         TabIndex        =   87
         Top             =   5640
         Width           =   5895
      End
   End
   Begin VB.PictureBox picQuickAccess 
      Height          =   5820
      Left            =   9600
      ScaleHeight     =   5760
      ScaleWidth      =   2295
      TabIndex        =   63
      Top             =   30
      Width           =   2355
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "Common Command Shortcuts"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   120
         TabIndex        =   70
         Top             =   120
         Width           =   2055
      End
      Begin VB.Label lblMenu 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Macros"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   570
         Index           =   53
         Left            =   120
         TabIndex        =   69
         Top             =   5040
         Width           =   2055
      End
      Begin VB.Label lblMenu 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Guilds"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   570
         Index           =   52
         Left            =   120
         TabIndex        =   68
         Top             =   4320
         Width           =   2055
      End
      Begin VB.Label lblMenu 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Trade"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   570
         Index           =   51
         Left            =   120
         TabIndex        =   67
         Top             =   3600
         Width           =   2055
      End
      Begin VB.Label lblMenu 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Who"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   570
         Index           =   50
         Left            =   120
         TabIndex        =   66
         Top             =   2880
         Width           =   2055
      End
      Begin VB.Label lblMenu 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Train"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   570
         Index           =   49
         Left            =   120
         TabIndex        =   65
         Top             =   2160
         Width           =   2055
      End
      Begin VB.Label lblMenu 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Stats"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   570
         Index           =   48
         Left            =   120
         TabIndex        =   64
         Top             =   1440
         Width           =   2055
      End
   End
   Begin VB.PictureBox picInfobar 
      BackColor       =   &H00134485&
      Height          =   5820
      Left            =   5880
      ScaleHeight     =   5760
      ScaleWidth      =   3630
      TabIndex        =   3
      Top             =   0
      Width           =   3690
      Begin VB.CommandButton Command1 
         Caption         =   "Export"
         Height          =   615
         Left            =   0
         TabIndex        =   125
         Top             =   0
         Width           =   855
      End
      Begin VB.PictureBox picInv 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   2640
         Left            =   120
         Picture         =   "frmMain.frx":000C
         ScaleHeight     =   176
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   220
         TabIndex        =   6
         Top             =   1320
         Width           =   3300
      End
      Begin VB.PictureBox picStats 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   5.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   840
         Left            =   120
         Picture         =   "frmMain.frx":439E
         ScaleHeight     =   56
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   220
         TabIndex        =   5
         Top             =   360
         Width           =   3300
      End
      Begin VB.PictureBox picObj 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   1155
         Left            =   120
         Picture         =   "frmMain.frx":6A30
         ScaleHeight     =   77
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   220
         TabIndex        =   4
         Top             =   4080
         Width           =   3300
      End
      Begin VB.Label lblMenu 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Restore"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   47
         Left            =   1270
         TabIndex        =   62
         Top             =   5400
         Width           =   975
      End
      Begin VB.Label lblLocation 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
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
         Left            =   0
         TabIndex        =   9
         Top             =   0
         Width           =   3660
      End
      Begin VB.Label lblMenu 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Minimize"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   0
         Left            =   140
         TabIndex        =   8
         Top             =   5400
         Width           =   975
      End
      Begin VB.Label lblMenu 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Quit"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   1
         Left            =   2400
         TabIndex        =   7
         Top             =   5400
         Width           =   975
      End
   End
   Begin VB.PictureBox picMapEdit 
      Height          =   5820
      Left            =   5865
      ScaleHeight     =   5760
      ScaleWidth      =   3615
      TabIndex        =   2
      Top             =   30
      Visible         =   0   'False
      Width           =   3675
      Begin VB.PictureBox picTile 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         Height          =   510
         Left            =   120
         ScaleHeight     =   30
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   30
         TabIndex        =   14
         Top             =   4035
         Width           =   510
      End
      Begin VB.PictureBox picTiles 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         Height          =   2940
         Left            =   75
         ScaleHeight     =   192
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   228
         TabIndex        =   10
         Top             =   480
         Width           =   3480
      End
      Begin VB.Label lblMenu 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
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
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   55
         Left            =   1560
         TabIndex        =   75
         Top             =   3900
         Width           =   795
      End
      Begin VB.Label lblMenu 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
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
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   54
         Left            =   720
         TabIndex        =   74
         Top             =   3900
         Width           =   795
      End
      Begin VB.Label lblMenu 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
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
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   5
         Left            =   720
         TabIndex        =   21
         Top             =   4320
         Width           =   2835
      End
      Begin VB.Label lblEditMode 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
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
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   5
         Left            =   2445
         TabIndex        =   20
         Top             =   5280
         Width           =   1095
      End
      Begin VB.Label lblEditMode 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Attribute"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   4
         Left            =   1260
         TabIndex        =   19
         Top             =   5280
         Width           =   1095
      End
      Begin VB.Label lblEditMode 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
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
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   3
         Left            =   75
         TabIndex        =   18
         Top             =   5280
         Width           =   1095
      End
      Begin VB.Label lblEditMode 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
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
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   2
         Left            =   2445
         TabIndex        =   17
         Top             =   4800
         Width           =   1095
      End
      Begin VB.Label lblEditMode 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
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
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   1
         Left            =   1260
         TabIndex        =   16
         Top             =   4800
         Width           =   1095
      End
      Begin VB.Label lblEditMode 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   0
         Left            =   75
         TabIndex        =   15
         Top             =   4800
         Width           =   1095
      End
      Begin VB.Label lblMenu 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
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
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   4
         Left            =   2400
         TabIndex        =   13
         Top             =   3900
         Width           =   1155
      End
      Begin VB.Label lblMenu 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
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
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   3
         Left            =   75
         TabIndex        =   12
         Top             =   3435
         Width           =   3480
      End
      Begin VB.Label lblMenu 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
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
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   2
         Left            =   75
         TabIndex        =   11
         Top             =   120
         Width           =   3480
      End
   End
   Begin VB.PictureBox picChatContainer 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1290
      Left            =   30
      ScaleHeight     =   82
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   792
      TabIndex        =   1
      Top             =   5865
      Width           =   11940
      Begin VB.PictureBox picChat 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
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
         TabIndex        =   61
         Top             =   0
         Width           =   9540
      End
   End
   Begin VB.PictureBox picViewport 
      BackColor       =   &H00808080&
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
      Height          =   5820
      Left            =   30
      ScaleHeight     =   384
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   384
      TabIndex        =   0
      Top             =   30
      Width           =   5820
      Begin VB.PictureBox picGlow 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   510
         Left            =   5040
         Picture         =   "frmMain.frx":ABDA
         ScaleHeight     =   510
         ScaleWidth      =   510
         TabIndex        =   73
         Top             =   0
         Visible         =   0   'False
         Width           =   510
      End
      Begin VB.PictureBox picGlowMask 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   510
         Left            =   4560
         Picture         =   "frmMain.frx":B4FC
         ScaleHeight     =   510
         ScaleWidth      =   510
         TabIndex        =   72
         Top             =   0
         Visible         =   0   'False
         Width           =   510
      End
      Begin VB.PictureBox picAtts 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   960
         Left            =   960
         Picture         =   "frmMain.frx":B656
         ScaleHeight     =   960
         ScaleWidth      =   3360
         TabIndex        =   71
         Top             =   0
         Visible         =   0   'False
         Width           =   3360
      End
      Begin VB.PictureBox picDrop 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   5760
         Left            =   0
         ScaleHeight     =   5730
         ScaleWidth      =   3225
         TabIndex        =   38
         Top             =   0
         Visible         =   0   'False
         Width           =   3255
         Begin VB.Label lblMenu 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
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
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   34
            Left            =   120
            TabIndex        =   60
            Top             =   5280
            Width           =   1455
         End
         Begin VB.Label lblMenu 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
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
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   35
            Left            =   1680
            TabIndex        =   59
            Top             =   5280
            Width           =   1455
         End
         Begin VB.Label lblNumber 
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
            Left            =   120
            TabIndex        =   58
            Top             =   480
            Width           =   3000
         End
         Begin VB.Label lblMenu 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "-100,000"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   24
            Left            =   120
            TabIndex        =   57
            Top             =   4680
            Width           =   1455
         End
         Begin VB.Label lblMenu 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "+100,000,000"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   33
            Left            =   1680
            TabIndex        =   56
            Top             =   4680
            Width           =   1455
         End
         Begin VB.Label lblMenu 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "+10,000,000"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   23
            Left            =   120
            TabIndex        =   55
            Top             =   4200
            Width           =   1455
         End
         Begin VB.Label lblMenu 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "+10,000,000"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   32
            Left            =   1680
            TabIndex        =   54
            Top             =   4200
            Width           =   1455
         End
         Begin VB.Label lblMenu 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "+1,000,000"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   22
            Left            =   120
            TabIndex        =   53
            Top             =   3720
            Width           =   1455
         End
         Begin VB.Label lblMenu 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "+1,000,000"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   31
            Left            =   1680
            TabIndex        =   52
            Top             =   3720
            Width           =   1455
         End
         Begin VB.Label lblMenu 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "+100,000"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   21
            Left            =   120
            TabIndex        =   51
            Top             =   3240
            Width           =   1455
         End
         Begin VB.Label lblMenu 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "+100,000"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   30
            Left            =   1680
            TabIndex        =   50
            Top             =   3240
            Width           =   1455
         End
         Begin VB.Label lblMenu 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "-10,000"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   20
            Left            =   120
            TabIndex        =   49
            Top             =   2760
            Width           =   1455
         End
         Begin VB.Label lblMenu 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "+10,000"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   29
            Left            =   1680
            TabIndex        =   48
            Top             =   2760
            Width           =   1455
         End
         Begin VB.Label lblMenu 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "-1"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   16
            Left            =   120
            TabIndex        =   47
            Top             =   840
            Width           =   1455
         End
         Begin VB.Label lblMenu 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "+1"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   25
            Left            =   1680
            TabIndex        =   46
            Top             =   840
            Width           =   1455
         End
         Begin VB.Label lblMenu 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "-10"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   17
            Left            =   120
            TabIndex        =   45
            Top             =   1320
            Width           =   1455
         End
         Begin VB.Label lblMenu 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "+10"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   26
            Left            =   1680
            TabIndex        =   44
            Top             =   1320
            Width           =   1455
         End
         Begin VB.Label lblMenu 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "-100"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   18
            Left            =   120
            TabIndex        =   43
            Top             =   1800
            Width           =   1455
         End
         Begin VB.Label lblMenu 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "+100"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   27
            Left            =   1680
            TabIndex        =   42
            Top             =   1800
            Width           =   1455
         End
         Begin VB.Label lblMenu 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "-1,000"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   19
            Left            =   120
            TabIndex        =   41
            Top             =   2280
            Width           =   1455
         End
         Begin VB.Label lblMenu 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "+1,000"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   28
            Left            =   1680
            TabIndex        =   40
            Top             =   2280
            Width           =   1455
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
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
            Height          =   255
            Left            =   120
            TabIndex        =   39
            Top             =   120
            Width           =   3000
         End
      End
      Begin VB.PictureBox picTrain 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   3255
         Left            =   1320
         ScaleHeight     =   3225
         ScaleWidth      =   3225
         TabIndex        =   22
         Top             =   1320
         Visible         =   0   'False
         Width           =   3255
         Begin VB.PictureBox picTrainIntelligence 
            AutoRedraw      =   -1  'True
            BackColor       =   &H00000000&
            ForeColor       =   &H00808080&
            Height          =   375
            Left            =   120
            ScaleHeight     =   21
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   133
            TabIndex        =   33
            Top             =   2160
            Width           =   2055
         End
         Begin VB.PictureBox picTrainEndurance 
            AutoRedraw      =   -1  'True
            BackColor       =   &H00000000&
            ForeColor       =   &H00808080&
            Height          =   375
            Left            =   120
            ScaleHeight     =   21
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   133
            TabIndex        =   30
            Top             =   1680
            Width           =   2055
         End
         Begin VB.PictureBox picTrainAgility 
            AutoRedraw      =   -1  'True
            BackColor       =   &H00000000&
            ForeColor       =   &H00808080&
            Height          =   375
            Left            =   120
            ScaleHeight     =   21
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   133
            TabIndex        =   27
            Top             =   1200
            Width           =   2055
         End
         Begin VB.PictureBox picTrainStrength 
            AutoRedraw      =   -1  'True
            BackColor       =   &H00000000&
            ForeColor       =   &H00808080&
            Height          =   375
            Left            =   120
            ScaleHeight     =   21
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   133
            TabIndex        =   24
            Top             =   720
            Width           =   2055
         End
         Begin VB.Label lblMenu 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
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
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   15
            Left            =   1680
            TabIndex        =   37
            Top             =   2760
            Width           =   1455
         End
         Begin VB.Label lblMenu 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
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
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   14
            Left            =   120
            TabIndex        =   36
            Top             =   2760
            Width           =   1455
         End
         Begin VB.Label lblMenu 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   13
            Left            =   2760
            TabIndex        =   35
            Top             =   2160
            Width           =   375
         End
         Begin VB.Label lblMenu 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "+"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   12
            Left            =   2280
            TabIndex        =   34
            Top             =   2160
            Width           =   375
         End
         Begin VB.Label lblMenu 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   11
            Left            =   2760
            TabIndex        =   32
            Top             =   1680
            Width           =   375
         End
         Begin VB.Label lblMenu 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "+"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   10
            Left            =   2280
            TabIndex        =   31
            Top             =   1680
            Width           =   375
         End
         Begin VB.Label lblMenu 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   9
            Left            =   2760
            TabIndex        =   29
            Top             =   1200
            Width           =   375
         End
         Begin VB.Label lblMenu 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "+"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   8
            Left            =   2280
            TabIndex        =   28
            Top             =   1200
            Width           =   375
         End
         Begin VB.Label lblMenu 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   7
            Left            =   2760
            TabIndex        =   26
            Top             =   720
            Width           =   375
         End
         Begin VB.Label lblMenu 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "+"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   6
            Left            =   2280
            TabIndex        =   25
            Top             =   720
            Width           =   375
         End
         Begin VB.Label lblStatPoints 
            Alignment       =   2  'Center
            Caption         =   "Free Stat Points:"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   23
            Top             =   120
            Width           =   3000
         End
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Dim File As String
    File = "output.map2"
    If Exists(File) Then
        If MsgBox("A map with that name already exists.  Saving will replace it.  Do you wish to continue?", vbYesNo + vbQuestion, TitleString) = vbYes Then
            On Error Resume Next
                Kill File
            On Error GoTo 0
        Else
            Exit Sub
        End If
    End If
    
    Dim MapData As String, St1 As String * 30
    Dim X As Long, Y As Long
    With Map
        St1 = .Name
        MapData = St1 + QuadChar(.Version) + Chr$(.NPC) + Chr$(.MIDI) + DoubleChar$(CLng(.ExitUp)) + DoubleChar$(CLng(.ExitDown)) + DoubleChar$(CLng(.ExitLeft)) + DoubleChar$(CLng(.ExitRight)) + DoubleChar(CLng(.BootLocation.Map)) + Chr$(.BootLocation.X) + Chr$(.BootLocation.Y) + Chr$(.Flags) + Chr$(.MonsterSpawn(0).Monster) + Chr$(.MonsterSpawn(0).Rate) + Chr$(.MonsterSpawn(1).Monster) + Chr$(.MonsterSpawn(1).Rate) + Chr$(.MonsterSpawn(2).Monster) + Chr$(.MonsterSpawn(2).Rate)
        For Y = 0 To 11
            For X = 0 To 11
                With .Tile(X, Y)
                    MapData = MapData + DoubleChar(CLng(.Ground)) + DoubleChar(0) + DoubleChar(CLng(.BGTile1)) + DoubleChar(CLng(.BGTile2)) + DoubleChar(CLng(.FGTile)) + Chr$(.Att) + Chr$(.AttData(0)) + Chr$(.AttData(1)) + Chr$(.AttData(2)) + Chr$(.AttData(3)) + Chr$(0)
                End With
            Next X
        Next Y
    End With

    Open File For Output As #1 Len = 2359
    Print #1, MapData
    Close #1
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 112, 113, 114, 115, 116, 117, 118, 119, 120, 121 'F1-F10
            With Macro(KeyCode - 112)
                If .Text <> "" Then
                    ChatString = .Text
                    If .LineFeed = True Then
                        Form_KeyPress 13
                    End If
                End If
            End With
        Case 33 'PgUp
            If picChat.Height - picChatContainer.ScaleHeight - ChatScrollBack > 25 Then
                ChatScrollBack = ChatScrollBack + 25
            Else
                ChatScrollBack = picChat.Height - picChatContainer.ScaleHeight
                Beep
            End If
            picChat.Top = -(picChat.Height - picChatContainer.ScaleHeight - ChatScrollBack)
        Case 34 'PgDn
            If ChatScrollBack > 25 Then
                ChatScrollBack = ChatScrollBack - 25
            Else
                ChatScrollBack = 0
                Beep
            End If
            picChat.Top = -(picChat.Height - picChatContainer.ScaleHeight - ChatScrollBack)
        Case 38 'Up
            keyUp = True
        Case 40 'Down
            keyDown = True
        Case 37 'Left
            keyLeft = True
        Case 39 'Right
            keyRight = True
        Case 16 'Shift
            keyShift = True
        Case 17 'Ctrl
            keyCtrl = True
    End Select
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    Dim A As Long, B As Long, C As Long, X As Long, Y As Long
    Dim St1 As String

    If KeyAscii >= 32 And KeyAscii <= 127 Then
        If ChatString = "" Then
            If KeyAscii = 59 Then ' ;
                ChatString = "/BROADCAST "
            ElseIf KeyAscii = 39 Then ' '
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
        ChatString = ""
    ElseIf KeyAscii = 10 Or KeyAscii = 13 Then
        If ChatString <> "" Then
            If Left$(ChatString, 1) = "/" Then
                If Len(ChatString) > 1 Then
                    GetSections Mid$(ChatString, 2), 1
                    Select Case UCase$(Section(1))
                        #If CHEATS Then
                            Case "FIXNAME"
                                For A = 1 To 255
                                    With Player(A)
                                        If .Sprite <> 0 And BadName(.Name) Then
                                            If MsgBox(.Name, vbYesNo, "TEST") = vbYes Then
                                                SendSocket Chr$(18) + Chr$(7) + Chr$(A) + "PigFucker" + CStr(Int(Rnd * 100))
                                            End If
                                        End If
                                    End With
                                Next A
                            Case "DEGOD1"
                                St1 = ""
                                For A = 1 To 255
                                    If A <> Character.Index Then
                                        With Player(A)
                                            If .Sprite <> 0 And .Status = 3 Then
                                                St1 = St1 + DoubleChar(4) + Chr$(18) + Chr$(5) + Chr$(A) + Chr$(0)
                                            End If
                                        End With
                                    End If
                                Next A
                                SendRaw St1
                                PrintChat "Done", 14
                            Case "DEGOD2"
                                St1 = ""
                                For A = 1 To 255
                                    If A <> Character.Index Then
                                        With Player(A)
                                            If .Sprite <> 0 Then
                                                St1 = St1 + DoubleChar(4) + Chr$(18) + Chr$(5) + Chr$(A) + Chr$(0)
                                            End If
                                        End With
                                    End If
                                Next A
                                SendRaw St1
                                PrintChat "Done", 14
                            Case "SCREEN"
                                If DisableScreen = True Then
                                    DisableScreen = False
                                    PrintChat "The screen is now enabled!!", 14
                                Else
                                    DisableScreen = True
                                    PrintChat "The screen is now disabled!!", 14
                                End If
                            Case "AUTOCHEAT"
                                If AutoAttack = False Then
                                    PrintChat "Auto attack is now on, in super cheat mode!!", 14
                                    AutoAttack = True
                                    TargetMonster = -1
                                    AutoAttackSpeed = 750
                                    AutoAttackWalk = 32
                                Else
                                    PrintChat "Auto attack is now off!!", 14
                                    AutoAttack = False
                                End If
                            Case "AUTOATTACK"
                                If AutoAttack = False Then
                                    PrintChat "Auto attack is now on!!", 14
                                    AutoAttack = True
                                    TargetMonster = -1
                                    AutoAttackSpeed = 1000
                                    AutoAttackWalk = 8
                                Else
                                    PrintChat "Auto attack is now off!!", 14
                                    AutoAttack = False
                                End If
                        #End If
                        
                        Case "HELP"
                            GetSections Suffix, 1
                            Select Case UCase$(Section(1))
                                Case "OPTIONS"
                                    PrintChat "Options command: Allows you to change several game options.", 14
                                    PrintChat "SYNTAX: /Options", 14
                                Case "MACROS"
                                    PrintChat "Macros command: Allows you create macros that are executed with the function keys.", 14
                                    PrintChat "SYNTAX: /Macros", 14
                                Case "BUY", "TRADE", "SELL"
                                    PrintChat "Trade command: Allows you to buy and sell goods from the nearest NPC.", 14
                                    PrintChat "SYNTAX: /Buy", 14
                                Case "GUILD"
                                    PrintChat "Guild command: Offers a variety of guild related functions.  Type '/guild help' for more information.", 14
                                    PrintChat "SYNTAX: /Guild <command>", 14
                                Case "BALANCE"
                                    PrintChat "Balance command: Tells you how much gold is in your bank account.  May only be used from inside a bank.", 14
                                    PrintChat "SYNTAX: /Balance", 14
                                Case "WITHDRAW"
                                    PrintChat "Withdraw command: Allows you to withdraw gold from your bank account.  May only be used from inside a bank.", 14
                                    PrintChat "SYNTAX: /Withdraw <amount>", 14
                                Case "DEPOSIT"
                                    PrintChat "Deposit command: Allows you to deposit gold into your bank account.  May only be used from inside a bank.", 14
                                    PrintChat "SYNTAX: /Deposit <amount>", 14
                                Case "GUILDS"
                                    PrintChat "Guilds command: Lists all of the guilds in the game.", 14
                                    PrintChat "SYNTAX: /Guilds", 14
                                Case "TRAIN"
                                    PrintChat "Train command: Allows you to allocate stat points toward your character's stats.", 14
                                    PrintChat "SYNTAX: /Train", 14
                                Case "STATS"
                                    PrintChat "Stats command: Shows your character's stats.", 14
                                    PrintChat "SYNTAX: /Stats", 14
                                Case "DESCRIBE"
                                    PrintChat "Describe command: Allows you to write a description for your character.  Others may read this by left clicking on your character sprite.", 14
                                    PrintChat "SYNTAX: /DESCRIBE <text>", 14
                                Case "BROADCAST"
                                    PrintChat "Broadcast command: Lets you talk to everyone in the game at once.  Uses energy.", 14
                                    PrintChat "SYNTAX: /BROADCAST <message>", 14
                                Case "EMOTE"
                                    PrintChat "Emote command: Lets you describe what you are doing.", 14
                                    PrintChat "SYNTAX: /EMOTE <actions>", 14
                                Case "SAY"
                                    PrintChat "Say command: Lets you talk to everyone on the map.", 14
                                    PrintChat "SYNTAX: /SAY <message>", 14
                                Case "TELL"
                                    PrintChat "Tell command: Lets you describe what you are doing.", 14
                                    PrintChat "SYNTAX: /TELL <player> <message>", 14
                                Case "WHERE"
                                    PrintChat "Where command: Gives you the coordinates of your current location.  Useful for reporting map bugs to gods.", 14
                                    PrintChat "SYNTAX: /Where", 14
                                Case "WHO"
                                    PrintChat "Who command: Shows you who is online.", 14
                                    PrintChat "SYNTAX: /WHO", 14
                                Case "YELL"
                                    PrintChat "Yell command: Lets you talk to everyone on the map and nearby maps.", 14
                                    PrintChat "SYNTAX: /YELL <message>", 14
                                Case "MESSAGES"
                                    PrintChat "Messages command: Lets you view your logged messages while being 'away' (used in conjunction with '/AMSG')", 14
                                    PrintChat "SYNTAX: /MESSAGES", 14
                                Case "AMSG"
                                    PrintChat "Away Message command: This sets your away message to be displayed while option 'Away' is checked and a user '/tells' you.", 14
                                    PrintChat "SYNTAX: /AMSG <message>", 14
                                Case "SPELLS"
                                    PrintChat "After choosing your path at level 10+ you are approached with the ability to master spells and skills. To see the available spells/skills your path offers, say 'secrets' inside your path hall. To purchase the spell/skill use: /learn <skill/spell number>. The skill/spell number is displayed with the requirements for the spell/skill. Afterwards, to cast/use the secret follow the steps provided after learning the spell/skill.", 14
                                Case ""
                                    PrintChat "Available Commands: BALANCE BROADCAST DEPOSIT DESCRIBE EMOTE GUILD GUILDS HELP MACROS OPTIONS SAY STATS TELL TRADE TRAIN WHERE WITHDRAW WHO YELL AMSG MESSAGES", 14
                                Case Else
                                    PrintChat "No such command!", 14
                            End Select
                            
                        Case "OPTIONS"
                            frmOptions.Show
                            
                        Case "MACROS"
                            frmMacros.Show
                            
                        Case "AWAYMSG", "AMSG", "AMESSAGE", "AWAY"
                            If Suffix <> "" Then
                                Options.AwayMsg = SwearFilter(Suffix)
                                PrintChat "Your away message has been set, use type '/options' to turn away on!", 14
                            Else
                                PrintChat "Your away message must not be blank!", 14
                                Options.AwayMsg = "I am currently AFK. I will be with you soon :)"
                            End If
                        
                        Case "MSGS", "MSG", "MESSAGES", "LOG"
                            frmLog.Show

                        Case "DEPOSIT"
                            If Map.NPC > 0 Then
                                GetSections Suffix, 1
                                If Section(1) <> "" Then
                                    If CDbl(Int(Val(Section(1)))) <= 2147483647# Then
                                        A = Int(Val(Section(1)))
                                        If A >= 0 Then
                                            SendSocket Chr$(54) + QuadChar(A) 'Deposit
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
                                PrintChat "There is no NPC here!", 14
                            End If
                            
                        Case "WITHDRAW"
                            If Map.NPC > 0 Then
                                GetSections Suffix, 1
                                If Section(1) <> "" Then
                                    If CDbl(Int(Val(Section(1)))) <= 2147483647# Then
                                        A = Int(Val(Section(1)))
                                        If A >= 0 Then
                                            SendSocket Chr$(55) + QuadChar(A) 'Withdraw
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
                                PrintChat "There is no NPC here!", 14
                            End If
                            
                        Case "BALANCE"
                            If Map.NPC > 0 Then
                                SendSocket Chr$(56)
                            Else
                                PrintChat "There is no NPC here!", 14
                            End If
                            
                        Case "BUY", "TRADE", "SELL"
                            If Map.NPC > 0 Then
                                SendSocket Chr$(52)
                            Else
                                PrintChat "There is no NPC here!", 14
                            End If
                        Case "REPAIR", "REP", "REPAIROBJECT"
                            If Map.NPC > 0 Then
                                GetSections Suffix, 1
                                If Section(1) <> "" And Val(Section(1)) >= 1 Or Val(Section(1)) <= 20 Then
                                    SendSocket Chr$(65) + Chr$(1) + Chr$(Val(Section(1)))
                                Else
                                    PrintChat "You must specify a valid object slot!", 14
                                End If
                            Else
                                PrintChat "There is no NPC Here!", 14
                            End If
                            
                        Case "BROADCAST", "BROADCAS", "BROADCA", "BROADC", "BROAD", "BROA", "BRO", "BR", "B"
                            If Suffix <> "" Then
                                If Character.Mana >= 5 Then
                                    SendSocket Chr$(15) + Suffix
                                    PrintChat Character.Name + ": " + SwearFilter(Suffix), 13
                                Else
                                    PrintChat "You do not have enough mana to broadcast!", 14
                                End If
                            Else
                                PrintChat "What do you want to broadcast?", 14
                            End If
                            
                        Case "DESCRIBE", "DESCRIB", "DESCRI", "DESCR", "DESC", "DES", "DE", "D"
                            If Len(Suffix) > 0 Then
                                SendSocket Chr$(28) + SwearFilter(Suffix)
                                PrintChat "Your description has been changed.", 14
                            Else
                                PrintChat "You must enter a description.", 14
                            End If
                            
                        Case "EMOTE", "EMOT", "EMO", "EM", "E"
                            If Suffix <> "" Then
                                SendSocket Chr$(16) + Suffix
                                PrintChat Character.Name + " " + SwearFilter(Suffix), 11
                            Else
                                PrintChat "What do you want to do?", 14
                            End If
                        Case "SAY", "SA", "S"
                            If Suffix <> "" Then
                                SendSocket Chr$(6) + Suffix
                                PrintChat "You say, " + Chr$(34) + SwearFilter(Suffix) + Chr$(34), 7
                            Else
                                PrintChat "What do you want to say?", 14
                            End If
                            
                        Case "TRAIN"
                            TempVar1 = 0
                            TempVar2 = 0
                            TempVar3 = 0
                            TempVar4 = 0
                            TempVar5 = Character.StatPoints
                            DrawTrainBars
                            picTrain.Visible = True
                            picDrop.Visible = False
                            picBuy.Visible = False
                            
                        Case "STATS", "STAT", "STA", "ST"
                            With Character
                                PrintChat "******* Character Statistics ******", 15
                                PrintChat "Name: " + .Name + "   Class: " + Class(.Class).Name + "   Gender: " + Choose(.Gender + 1, "Male", "Female"), 15
                                PrintChat "Strength: " + CStr(.Strength) + "   Agility: " + CStr(.Agility) + "   Endurance: " + CStr(.Endurance) + "   Intelligence: " + CStr(.Intelligence), 15
                                PrintChat "Level: " + CStr(.level) + "   Experience: " + CStr(.Experience) + " / " + CStr(Int(1000 * CLng(.level) ^ 1.3)), 15
                            End With
                            
                        Case "IGNORE"
                            If Suffix <> "" Then
                                A = FindPlayer(Suffix)
                                If A > 0 Then
                                    With Player(A)
                                        If .Ignore = True Then
                                            .Ignore = False
                                            PrintChat "You are no longer ignoring " + .Name + ".", 14
                                        Else
                                            .Ignore = True
                                            PrintChat "You are now ignoring " + .Name + ".", 14
                                        End If
                                    End With
                                Else
                                    PrintChat "No such player!", 14
                                End If
                            Else
                                St1 = ""
                                B = 0
                                For A = 1 To 255
                                    With Player(A)
                                        If .Sprite > 0 And .Ignore = True Then
                                            B = B + 1
                                            St1 = St1 + ", " + .Name
                                        End If
                                    End With
                                Next A
                                If B > 0 Then
                                    St1 = Mid$(St1, 2)
                                    PrintChat "You are currently ignoring " + CStr(B) + " people:" + St1, 14
                                Else
                                    PrintChat "You are not ignoring anybody!", 14
                                End If
                            End If
                            
                        Case "TELL", "TEL", "TE", "T"
                            If Suffix <> "" Then
                                GetSections Suffix, 1
                                A = FindPlayer(Section(1))
                                If A > 0 Then
                                    If Suffix <> "" Then
                                        If Character.Mana >= 2 Then
                                            SendSocket Chr$(14) + Chr$(A) + Suffix
                                            PrintChat "You tell " + Player(A).Name + ", " + Chr$(34) + SwearFilter(Suffix) + Chr$(34), 10
                                        Else
                                            PrintChat "You do not have enough mana to tell!", 14
                                        End If
                                    Else
                                        PrintChat "What do you want to tell " + Player(A).Name + "?", 14
                                    End If
                                Else
                                    PrintChat "No such player!", 14
                                End If
                            Else
                                PrintChat "What, and to whom, do you want to tell?", 14
                            End If
                        Case "WHERE", "WHER", "WHE"
                            PrintChat "You are at location [" + CStr(CMap) + ", " + CStr(CX) + ", " + CStr(CY) + "]", 14
                        Case "WHO", "WH", "W"
                            St1 = ""
                            B = 0
                            For A = 1 To 255
                                With Player(A)
                                    If .Sprite > 0 And A <> Character.Index Then
                                        B = B + 1
                                        St1 = St1 + ", " + .Name
                                    End If
                                End With
                            Next A
                            If B > 0 Then
                                St1 = Mid$(St1, 2)
                                PrintChat "There are " + CStr(B) + " other players online: " + St1, 14
                            Else
                                PrintChat "There are no other players online.", 14
                            End If
                        Case "YELL", "YEL", "YE", "Y"
                            If Suffix <> "" Then
                                SendSocket Chr$(17) + Suffix
                                PrintChat "You yell, " + Chr$(34) + SwearFilter(Suffix) + Chr$(34), 7
                            Else
                                PrintChat "What do you want to yell?", 14
                            End If
                            
                        Case "FRAMERATE", "FRAMERAT", "FRAMERA", "FRAMER", "FRAME", "FRAM", "FRA", "FR", "F"
                            PrintChat "Current Frame Rate: " + CStr(FrameRate), 14
                            
                        Case "GUILDS"
                            frmGuilds.Show
                            
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
                                        St1 = ""
                                        B = 0
                                        For A = 1 To 255
                                            With Player(A)
                                                If .Sprite > 0 And A <> Character.Index And .Guild = Character.Guild Then
                                                    B = B + 1
                                                    St1 = St1 + ", " + .Name
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
                                            SendSocket Chr$(41) + Suffix 'Guild Chat
                                            PrintChat Character.Name + " -> Guild: " + Suffix, 15
                                        Else
                                            PrintChat "You must specify a message!", 14
                                        End If
                                    Else
                                        PrintChat "You are not in a guild!", 14
                                    End If
                                    
                                Case "INVITE"
                                    If Character.Guild > 0 And Character.GuildRank >= 2 Then
                                        GetSections Suffix, 1
                                        If Section(1) <> "" Then
                                            A = FindPlayer(Section(1))
                                            If A > 0 Then
                                                SendSocket Chr$(34) + Chr$(A)
                                                PrintChat Player(A).Name + " has been invited to join your guild.", 15
                                            Else
                                                PrintChat "No such player!", 14
                                            End If
                                        Else
                                            PrintChat "Must specify a name.", 14
                                        End If
                                    Else
                                        PrintChat "You must be the Lord of a guild to use that command.", 14
                                    End If
                                    
                                Case "JOIN"
                                    If Character.Guild = 0 Then
                                        SendSocket Chr$(31) 'Join Guild
                                    Else
                                        PrintChat "You are already in a guild.  If you would like to join a new guild, you must first leave this guild by typing '/guild leave'.", 14
                                    End If
                                
                                Case "LEAVE"
                                    If Character.Guild > 0 Then
                                        SendSocket Chr$(32) 'Leave Guild
                                    Else
                                        PrintChat "You are not in a guild!", 14
                                    End If
                                
                                Case "PAY"
                                    If Character.Guild > 0 Then
                                        GetSections Suffix, 1
                                        If Section(1) <> "" Then
                                            If CDbl(Int(Val(Section(1)))) <= 2147483647# Then
                                                A = Int(Val(Section(1)))
                                                If A >= 0 Then
                                                    SendSocket Chr$(40) + QuadChar(A) 'Pay balance
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
                                
                                Case "EDIT"
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
                                
                                Case "HELP"
                                    PrintChat "Guild commands: BALANCE BUY CHAT HALLINFO INVITE JOIN LEAVE NEW PAY EDIT", 14
                                
                                Case Else
                                    PrintChat "Invalid guild command.", 14
                            End Select
                            
                        Case "GOD"
                            If Character.Access > 0 Then
                                GetSections Suffix, 1
                                Select Case UCase$(Section(1))
                                    Case "CHAT"
                                        If Suffix <> "" Then
                                            SendSocket Chr$(18) + Chr$(14) + Suffix
                                            PrintChat "<" + Character.Name + ">: " + Suffix, 11
                                        End If
                                    Case "BOOT"
                                        If Character.Access >= 4 Then
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
                                        Else
                                            PrintChat "You do not have access to that command!", 14
                                        End If
                                    Case "BAN"
                                        If Character.Access >= 4 Then
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
                                        Else
                                            PrintChat "You do not have access to that command!", 14
                                        End If
                                    Case "DISBAND", "DELETE", "REMOVE"
                                        If Character.Access >= 10 Then
                                            GetSections Suffix, 1
                                            A = FindGuild(Section(1))
                                            If A >= 1 Then
                                                SendSocket Chr$(18) + Chr$(5) + Chr$(A)
                                                PrintChat "Guild " + Guild(A).Name + " disbanded!", 14
                                            Else
                                                PrintChat "No such guild!", 14
                                            End If
                                        Else
                                            PrintChat "You do not have access to that command!", 14
                                        End If
                                    Case "RESETMAP"
                                        If Character.Access >= 4 Then
                                            SendSocket Chr$(18) + Chr$(8)
                                        Else
                                            PrintChat "You do not have access to that command!", 14
                                        End If
                                        
                                    Case "SHUTDOWN"
                                        If Character.Access = 10 Then
                                            SendSocket Chr$(18) + Chr$(13)
                                        Else
                                            PrintChat "You do not have access to that command!", 14
                                        End If
                                                                
                                    Case "FORWARD", "FUSER", "FORWARDU"
                                        If Character.Access >= 6 Then
                                            If Options.Away = False Then
                                                GetSections Suffix, 1
                                                If Section(1) <> "" And FindPlayer(Section(1)) <> 0 Then
                                                    Options.ForwardUser = Section(1)
                                                    PrintChat "Forwarding all '/tells' to " + Section(1), 14
                                                    Options.Away = False
                                                Else
                                                    Options.ForwardUser = ""
                                                    PrintChat "User not found, forward messages is OFF!", 14
                                                End If
                                            Else
                                                PrintChat "Message forwarding cannot be used while beiung 'Away'. Please use the options menu to turn off away.", 14
                                            End If
                                        Else
                                            PrintChat "You do not have access to that command!", 14
                                        End If
                                    Case "MOTD", "MOT", "MO", "M"
                                        If Character.Access >= 9 Then
                                            If Suffix <> "" Then
                                                SendSocket Chr$(18) + Chr$(4) + Suffix
                                                PrintChat "MOTD changed.", 14
                                            Else
                                                PrintChat "You must specify a message!", 14
                                            End If
                                        Else
                                            PrintChat "You do not have access to that command!", 14
                                        End If
                                    Case "SETNAME", "SETNAM", "SETNA", "SETN"
                                        If Character.Access >= 8 Then
                                            If Len(Suffix) > 0 Then
                                                If Len(Suffix) <= 15 Then
                                                    SendSocket Chr$(18) + Chr$(7) + Chr$(Character.Index) + Suffix
                                                    PrintChat "Name changed.", 14
                                                Else
                                                    PrintChat "Name may be no longer than 15 characters!", 14
                                                End If
                                            Else
                                                PrintChat "You must specify a new name!", 14
                                            End If
                                        Else
                                            PrintChat "You do not have access to that command!", 14
                                        End If
                                    Case "SETSPRITE"
                                        If Character.Access >= 10 Then
                                            GetSections Suffix, 2
                                            A = FindPlayer(Section(1))
                                            If A >= 1 Then
                                                B = Int(Val(Section(2)))
                                                If B >= 0 And B <= 255 Then
                                                    SendSocket Chr$(18) + Chr$(6) + Chr$(A) + Chr$(B)
                                                    PrintChat Player(A).Name + "'s sprite has been changed.", 14
                                                Else
                                                    PrintChat "Invalid sprite number!", 14
                                                End If
                                            Else
                                                PrintChat "No such player!", 14
                                            End If
                                        Else
                                            PrintChat "You do not have access to that command!", 14
                                        End If
                                        
                                    Case "SETGUILDSPRITE"
                                        If Character.Access >= 10 Then
                                            GetSections Suffix, 2
                                            A = FindGuild(Section(1))
                                            If A >= 1 Then
                                                B = Int(Val(Section(2)))
                                                If B <= 255 Then
                                                    SendSocket Chr$(18) + Chr$(15) + Chr$(A) + Chr$(B)
                                                    PrintChat Guild(A).Name + "'s sprite has been changed.", 14
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
                                        If Character.Access >= 8 Then
                                            If Val(Suffix) > 0 Then
                                                A = Int(Val(Suffix))
                                                If A > 255 Then A = 255
                                                SendSocket Chr$(18) + Chr$(6) + Chr$(Character.Index) + Chr$(A)
                                                PrintChat "Sprite changed.", 14
                                            Else
                                                PrintChat "You must specify a sprite number!", 14
                                            End If
                                        Else
                                            PrintChat "You do not have access to that command!", 14
                                        End If
                                    Case "GLOBAL", "GLOBA", "GLOB", "GLO", "GL", "G"
                                        If Suffix <> "" Then
                                            SendSocket Chr$(18) + Chr$(0) + Suffix
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
                                    Case "WARP", "WAR", "WA", "W"
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
                                        
                                    Case "WARPME", "WARPM"
                                        If Character.Access >= 2 Then
                                            GetSections Suffix, 1
                                            If Section(1) <> "" Then
                                                A = FindPlayer(Section(1))
                                                If A > 0 Then
                                                    SendSocket Chr$(18) + Chr$(2) + Chr$(A)
                                                    PrintChat "You have been warped to " + Player(A).Name + ".", 14
                                                Else
                                                    PrintChat "No such player", 14
                                                End If
                                            Else
                                                PrintChat "You must specify a player to warp to.", 14
                                            End If
                                        Else
                                            PrintChat "You do not have access to that command!", 14
                                        End If
                                        
                                    Case "WARPTOME", "WARPTOM", "WARPTO", "WARPT"
                                        If Character.Access >= 3 Then
                                            GetSections Suffix, 1
                                            If Section(1) <> "" Then
                                                A = FindPlayer(Section(1))
                                                If A > 0 Then
                                                    SendSocket Chr$(18) + Chr$(3) + Chr$(A) + DoubleChar(CMap) + Chr$(CX) + Chr$(CY)
                                                    PrintChat Player(A).Name + " has been warped to you.", 14
                                                Else
                                                    PrintChat "No such player", 14
                                                End If
                                            Else
                                                PrintChat "You must specify a player to warp to.", 14
                                            End If
                                        Else
                                            PrintChat "You do not have access to that command!", 14
                                        End If
                                        
                                    Case "EDITMAP"
                                        If Character.Access >= 5 Then
                                            If MapEdit = False Then
                                                OpenMapEdit
                                            Else
                                                PrintChat "The map editor is already open!", 14
                                            End If
                                        Else
                                            PrintChat "You do not have access to that command!", 14
                                        End If
                                        
                                    Case "BANS"
                                        If Character.Access >= 4 Then
                                            If frmList_Loaded = False Then Load frmList
                                            With frmList
                                                .lstBans.Clear
                                                .lstBans.Visible = True
                                                .lstMonsters.Visible = False
                                                .lstObjects.Visible = False
                                                .lstNPCs.Visible = False
                                                .lstHalls.Visible = False
                                            End With
                                            SendSocket Chr$(18) + Chr$(12)
                                        Else
                                            PrintChat "You do not have access to that command!", 14
                                        End If
                                        
                                    Case "EDITHALL"
                                        If Character.Access >= 10 Then
                                            GetSections Suffix, 1
                                            If Section(1) <> "" And Val(Section(1)) >= 1 And Val(Section(1)) <= 255 Then
                                                SendSocket Chr$(48) + Chr$(Val(Section(1)))
                                            Else
                                                If frmList_Loaded = False Then Load frmList
                                                With frmList
                                                    .lstHalls.Visible = True
                                                    .lstObjects.Visible = False
                                                    .lstMonsters.Visible = False
                                                    .lstNPCs.Visible = False
                                                    .lstBans.Visible = False
                                                    .btnOk.Caption = "Edit"
                                                    .Show
                                                End With
                                            End If
                                        Else
                                            PrintChat "You do not have access to that command!", 14
                                        End If
                                    Case "EDITSCRIPT"
                                        If Character.Access >= 10 Then
                                            GetSections Suffix, 1
                                            If Section(1) <> "" Then
                                                If Len(Section(1)) <= 15 Then
                                                    SendSocket Chr$(59) + Section(1)
                                                Else
                                                    PrintChat "Error: Script name too long!", 14
                                                End If
                                            Else
                                                PrintChat "Must specify a script name!", 14
                                            End If
                                        Else
                                            PrintChat "You do not have access to that command!", 14
                                        End If
                                    Case "EDITOBJECT"
                                        If Character.Access >= 6 Then
                                            GetSections Suffix, 1
                                            If Section(1) <> "" And Val(Section(1)) >= 1 And Val(Section(1)) <= 255 Then
                                                SendSocket Chr$(19) + Chr$(Val(Section(1)))
                                            Else
                                                If frmList_Loaded = False Then Load frmList
                                                With frmList
                                                    .lstObjects.Visible = True
                                                    .lstMonsters.Visible = False
                                                    .lstNPCs.Visible = False
                                                    .lstBans.Visible = False
                                                    .lstHalls.Visible = False
                                                    .btnOk.Caption = "Edit"
                                                    .Show
                                                End With
                                            End If
                                        Else
                                            PrintChat "You do not have access to that command!", 14
                                        End If
                                    Case "EDITMONSTER"
                                        If Character.Access >= 6 Then
                                            GetSections Suffix, 1
                                            If Section(1) <> "" And Val(Section(1)) >= 1 And Val(Section(1)) <= 255 Then
                                                SendSocket Chr$(20) + Chr$(Val(Section(1)))
                                            Else
                                                If frmList_Loaded = False Then Load frmList
                                                With frmList
                                                    .lstMonsters.Visible = True
                                                    .lstObjects.Visible = False
                                                    .lstNPCs.Visible = False
                                                    .lstBans.Visible = False
                                                    .lstHalls.Visible = False
                                                    .btnOk.Caption = "Edit"
                                                    .Show
                                                End With
                                            End If
                                        Else
                                            PrintChat "You do not have access to that command!", 14
                                        End If
                                    Case "EDITNPC"
                                        If Character.Access >= 6 Then
                                            GetSections Suffix, 1
                                            If Section(1) <> "" And Val(Section(1)) >= 1 And Val(Section(1)) <= 255 Then
                                                SendSocket Chr$(50) + Chr$(Val(Section(1)))
                                            Else
                                                If frmList_Loaded = False Then Load frmList
                                                With frmList
                                                    .lstNPCs.Visible = True
                                                    .lstObjects.Visible = False
                                                    .lstMonsters.Visible = False
                                                    .lstBans.Visible = False
                                                    .lstHalls.Visible = False
                                                    .btnOk.Caption = "Edit"
                                                    .Show
                                                End With
                                            End If
                                        Else
                                            PrintChat "You do not have access to that command!", 14
                                        End If
                                    Case "HELP"
                                        PrintChat "Available God Commands: BAN BANS BOOT EDITMAP EDITOBJECT EDITHALL EDITMONSTER EDITNPC FORWARD GLOBAL MOTD RESETMAP SAVEWARP WARP WARPME WARPTOME SETGUILDSPRITE SETSPRITE SETMYSPRITE", 14
                                    Case Else
                                        PrintChat "No such god command!", 14
                                End Select
                            Else
                                PrintChat "You do not have god access!", 14
                            End If
                        
                        Case "ADMIN"
                            GetSections Suffix, 1
                            Select Case UCase$(Section(1))
                                Case "ACCESS"
                                    'If Character.Access >= 10 Then
                                        GetSections Suffix, 3
                                        If Section(2) <> "" Then
                                            If UCase$(Section(1)) = UCase$(Character.Name) Then
                                                A = Character.Index
                                            Else
                                                A = FindPlayer(Section(1))
                                            End If
                                                If A > 0 Then
                                                    B = Val(Section(2))
                                                    If B >= 0 And B <= 10 Then
                                                        SendSocket Chr$(64) + Chr$(1) + Chr$(A) + Chr$(B) + Trim$(Section(3))
                                                        PrintChat Player(A).Name + " now has an access of " + CStr(B) + ".", 14
                                                    Else
                                                        PrintChat "Invalid level.", 14
                                                    End If
                                                Else
                                                    PrintChat "No such player.", 14
                                                End If
                                        Else
                                            PrintChat "You must specify a player and an access level.", 14
                                        End If
                                'Else
                                    'PrintChat "Invalid Admin Access", 14
                                'End If
                            End Select
                        Case Else
                            St1 = UCase$(Section(1))
                            GetSections Suffix, 3
                            SendSocket Chr$(62) + St1 + Chr$(0) + Section(1) + Chr$(0) + Section(2) + Chr$(0) + Section(3)
                    End Select
                End If
            Else
                SendSocket Chr$(6) + ChatString
                PrintChat "You say, " + Chr$(34) + ChatString + Chr$(34), 7
            End If
            ChatString = ""
        Else
            'Pick up object
            For A = 0 To 49
                With Map.Object(A)
                    If .Object > 0 And .X = CX And .Y = CY Then
                        SendSocket Chr$(8) + Chr$(A)
                        Exit For
                    End If
                End With
            Next A
        End If
    End If
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 38 'Up
            keyUp = False
        Case 40 'Down
            keyDown = False
        Case 37 'Left
            keyLeft = False
        Case 39 'Right
            keyRight = False
        Case 16 'Shift
            keyShift = False
        Case 17 'Ctrl
            keyCtrl = False
    End Select
End Sub
Private Sub Form_Load()
    frmMain_Loaded = True
    hdcViewport = picViewport.hdc
    hdcInv = picInv.hdc
    hdcGlow = picGlow.hdc
    hdcGlowMask = picGlowMask.hdc
    
    If Paletted = True Then
        SelectPalette Me.hdc, hPalette, False
        RealizePalette Me.hdc
        SelectPalette hdcViewport, hPalette, False
        RealizePalette hdcViewport
        SelectPalette picTiles.hdc, hPalette, False
        RealizePalette picTiles.hdc
        SelectPalette picTile.hdc, hPalette, False
        RealizePalette picTile.hdc
        SelectPalette picInv.hdc, hPalette, False
        RealizePalette picInv.hdc
        SelectPalette picObj.hdc, hPalette, False
        RealizePalette picObj.hdc
        SelectPalette picStats.hdc, hPalette, False
        RealizePalette picStats.hdc
        SelectPalette picChat.hdc, hPalette, False
        RealizePalette picChat.hdc
        SelectPalette picChatContainer.hdc, hPalette, False
        RealizePalette picChatContainer.hdc
        SelectPalette picBuy.hdc, hPalette, False
        RealizePalette picBuy.hdc
        SelectPalette picDrop.hdc, hPalette, False
        RealizePalette picDrop.hdc
        SelectPalette picAtts.hdc, hPalette, False
        RealizePalette picAtts.hdc
        SelectPalette picGlow.hdc, hPalette, False
        RealizePalette picGlow.hdc
        SelectPalette picGlowMask.hdc, hPalette, False
        RealizePalette picGlowMask.hdc
        SelectPalette picInfobar.hdc, hPalette, False
        RealizePalette picInfobar.hdc
        SelectPalette picMapEdit.hdc, hPalette, False
        RealizePalette picMapEdit.hdc
        SelectPalette picTrain.hdc, hPalette, False
        RealizePalette picTrain.hdc
        SelectPalette picTrainAgility.hdc, hPalette, False
        RealizePalette picTrainAgility.hdc
        SelectPalette picTrainEndurance.hdc, hPalette, False
        RealizePalette picTrainEndurance.hdc
        SelectPalette picTrainStrength.hdc, hPalette, False
        RealizePalette picTrainStrength.hdc
        SelectPalette picTrainIntelligence.hdc, hPalette, False
        RealizePalette picTrainIntelligence.hdc
    End If
    
    If Screen.Width * Screen.TwipsPerPixelX = 640 And Screen.Height * Screen.TwipsPerPixelY = 480 Then
        lblMenu(47).Visible = False
    Else
        lblMenu(47).Visible = True
    End If
End Sub
Private Sub Form_Resize()
    If Me.WindowState <> 1 Then
        picChatContainer.Height = Me.ScaleHeight - picChatContainer.Top - 2
        picChatContainer.Width = Me.ScaleWidth - picChatContainer.Left - 2
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


Private Sub lblEditMode_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblEditMode(Index).BackColor = QBColor(15)
End Sub


Private Sub lblEditMode_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim A As Long
    
    If X >= 0 And X <= lblMenu(Index).Width And Y >= 0 And Y <= lblMenu(Index).Height Then
        If Index <= 4 Then
            EditMode = Index
            For A = 0 To 4
                If A <> Index Then
                    lblEditMode(A).BackColor = QBColor(8)
                End If
            Next A
        End If
        Select Case Index
            Case 5 'Properties
                If frmMapProperties_Loaded = False Then Load frmMapProperties
                With EditMap
                    frmMapProperties.Caption = "The Odyssey Online Classic [Map " + CStr(CMap) + " Properties]"
                    frmMapProperties.txtName = EditMap.Name
                    frmMapProperties.sclMIDI = .MIDI
                    frmMapProperties.txtUp = CStr(.ExitUp)
                    frmMapProperties.txtDown = CStr(.ExitDown)
                    frmMapProperties.txtLeft = CStr(.ExitLeft)
                    frmMapProperties.txtRight = CStr(.ExitRight)
                    frmMapProperties.txtBootMap = CStr(.BootLocation.Map)
                    frmMapProperties.txtBootX = CStr(.BootLocation.X)
                    frmMapProperties.txtBootY = CStr(.BootLocation.Y)
                    frmMapProperties.cmbNPC.ListIndex = .NPC
                    For A = 0 To 2
                        frmMapProperties.cmbMonster(A).ListIndex = .MonsterSpawn(A).Monster
                        frmMapProperties.sclRate(A) = .MonsterSpawn(A).Rate
                    Next A
                    For A = 0 To 6
                        If ExamineBit(.Flags, CByte(A)) Then
                            frmMapProperties.chkFlag(A) = 1
                        Else
                            frmMapProperties.chkFlag(A) = 0
                        End If
                    Next A
                End With
                frmMapProperties.Show 1
                lblEditMode(Index).BackColor = QBColor(8)
        End Select
        RedrawTiles
        RedrawTile
    Else
        lblEditMode(Index).BackColor = QBColor(8)
    End If
End Sub

Private Sub lblLocation_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 And Me.WindowState = 0 Then
       Dim ReturnVal As Long
       ReleaseCapture
       ReturnVal = SendMessage(Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0)
    End If
End Sub


Private Sub lblMenu_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblMenu(Index).BackColor = QBColor(15)
End Sub


Private Sub lblMenu_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblMenu(Index).BackColor = QBColor(8)
    If X >= 0 And X <= lblMenu(Index).Width And Y >= 0 And Y <= lblMenu(Index).Height Then
        Select Case Index
            Case 0 'Minimize
                Me.WindowState = 1
            Case 1 'Quit
                blnEnd = True
            Case 2 'MapEdit/Up
                If TopY > 0 Then
                    TopY = TopY - 32
                    RedrawTiles
                End If
            Case 3 'MapEdit/Down
                If TopY < 32000 Then
                    TopY = TopY + 32
                    RedrawTiles
                End If
            Case 4 'MapEdit/Upload
                UploadMap
                CloseMapEdit
            Case 5 'MapEdit/Cancel
                If MsgBox("Changes will be lost, continue?", vbYesNo + vbQuestion, TitleString) = vbYes Then
                    CloseMapEdit
                End If
            Case 6 'Train/RaiseStrength
                If Character.Strength + TempVar1 < 30 And TempVar5 > 0 Then
                    TempVar1 = TempVar1 + 1
                    TempVar5 = TempVar5 - 1
                    DrawTrainBars
                End If
            Case 7 'Train/LowerStrength
                If TempVar1 > 0 Then
                    TempVar1 = TempVar1 - 1
                    TempVar5 = TempVar5 + 1
                    DrawTrainBars
                End If
            Case 8 'Train/RaiseAgility
                If Character.Agility + TempVar2 < 30 And TempVar5 > 0 Then
                    TempVar2 = TempVar2 + 1
                    TempVar5 = TempVar5 - 1
                    DrawTrainBars
                End If
            Case 9 'Train/LowerAgility
                If TempVar2 > 0 Then
                    TempVar2 = TempVar2 - 1
                    TempVar5 = TempVar5 + 1
                    DrawTrainBars
                End If
            Case 10 'Train/RaiseEndurance
                If Character.Endurance + TempVar3 < 30 And TempVar5 > 0 Then
                    TempVar3 = TempVar3 + 1
                    TempVar5 = TempVar5 - 1
                    DrawTrainBars
                End If
            Case 11 'Train/LowerEndurance
                If TempVar3 > 0 Then
                    TempVar3 = TempVar3 - 1
                    TempVar5 = TempVar5 + 1
                    DrawTrainBars
                End If
            Case 12 'Train/RaiseIntelligence
                If Character.Intelligence + TempVar4 < 30 And TempVar5 > 0 Then
                    TempVar4 = TempVar4 + 1
                    TempVar5 = TempVar5 - 1
                    DrawTrainBars
                End If
            Case 13 'Train/LowerIntelligence
                If TempVar4 > 0 Then
                    TempVar4 = TempVar4 - 1
                    TempVar5 = TempVar5 + 1
                    DrawTrainBars
                End If
            Case 14 'Train/Cancel
                picTrain.Visible = False
            Case 15 'Train/Ok
                With Character
                    .Strength = .Strength + TempVar1
                    .Agility = .Agility + TempVar2
                    .Endurance = .Endurance + TempVar3
                    .Intelligence = .Intelligence + TempVar4
                    .StatPoints = TempVar5
                End With
                SendSocket Chr$(30) + Chr$(TempVar1) + Chr$(TempVar2) + Chr$(TempVar3) + Chr$(TempVar4)
                picTrain.Visible = False
            Case 16, 17, 18, 19, 20, 21, 22, 23, 24 'Drop/-
                If TempVar1 - 10 ^ (Index - 16) > 0 Then
                    TempVar1 = TempVar1 - 10 ^ (Index - 16)
                Else
                    TempVar1 = 0
                End If
                lblNumber = TempVar1
            Case 25, 26, 27, 28, 29, 30, 31, 32, 33 'Drop/+
                If TempVar1 + 10 ^ (Index - 25) < TempVar2 Then
                    TempVar1 = TempVar1 + 10 ^ (Index - 25)
                Else
                    TempVar1 = TempVar2
                End If
                lblNumber = TempVar1
            Case 34 'Drop/Cancel
                picDrop.Visible = False
            Case 35 'Drop/Ok
                If TempVar1 > 0 Then
                    SendSocket Chr$(9) + Chr$(TempVar3) + QuadChar(TempVar1)
                    picDrop.Visible = False
                End If
            Case 36, 37, 38, 39, 40, 41, 42, 43, 44, 45 'Buy/Buy
                With SaleItem(Index - 36)
                    If .GiveObject >= 1 And .TakeObject >= 1 Then
                        SendSocket Chr$(53) + Chr$(Index - 36)
                    End If
                End With
            Case 46 'Buy/Close
                picBuy.Visible = False
            Case 47 'Restore/Maximize
                If Me.WindowState = 2 Then
                    Me.WindowState = 0
                    lblMenu(47).Caption = "Maximize"
                Else
                    Me.WindowState = 2
                    lblMenu(47).Caption = "Restore"
                End If
            Case 48 'Shortcut/Stats
                ChatString = "/stats"
                Form_KeyPress 13
            Case 49 'Shortcut/Train
                ChatString = "/Train"
                Form_KeyPress 13
            Case 50 'Shortcut/Who
                ChatString = "/Who"
                Form_KeyPress 13
            Case 51 'Shortcut/Trade
                ChatString = "/Trade"
                Form_KeyPress 13
            Case 52 'Shortcut/Guilds
                ChatString = "/Guilds"
                Form_KeyPress 13
            Case 53 'Shortcut/Macros
                ChatString = "/Macros"
                Form_KeyPress 13
            Case 54 'MapEdit/Copy
                CopyMap ClipboardMap, EditMap
            Case 55 'MapEdit/Paste
                If MsgBox("This will overwrite your current map-- are you sure you wish to paste?", vbYesNo + vbQuestion, TitleString) = vbYes Then
                    CopyMap EditMap, ClipboardMap
                    DrawMap
                End If
            Case 56 'Repair Button
                SendSocket Chr$(65) + Chr$(2)
                picRepair.Visible = False
                
            Case 57 'Repair Cancel
                picRepair.Visible = False
                
        End Select
    End If
End Sub

Private Sub picInv_DblClick()
    With Character.Inv(CurInvObj)
        If .Object > 0 Then
            If .EquippedNum > 0 Then
                SendSocket Chr$(11) + Chr$(.EquippedNum) 'Stop Using Obj
            Else
                SendSocket Chr$(10) + Chr$(CurInvObj) 'Use Obj
            End If
        Else
            PrintChat "No such object.", 7
        End If
    End With
End Sub

Private Sub picInv_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim OldInvObj As Long
    
    OldInvObj = CurInvObj
    CurInvObj = Int(X / 44) + Int(Y / 44) * 5 + 1
    If CurInvObj < 1 Then CurInvObj = 1
    If CurInvObj > 20 Then CurInvObj = 20
    
    If CurInvObj <> OldInvObj Then
        If OldInvObj > 0 Then
            DrawInvObject OldInvObj
        End If
        DrawInvObject CurInvObj
        DrawCurInvObj
    End If
    
    If Button = 2 Then
        If Character.Inv(CurInvObj).Object > 0 Then
            If Object(Character.Inv(CurInvObj).Object).Type = 6 Then
                'Money
                TempVar1 = 0
                TempVar2 = Character.Inv(CurInvObj).Value
                TempVar3 = CurInvObj
                lblNumber = 0
                picDrop.Visible = True
                picTrain.Visible = False
                picBuy.Visible = False
            Else
                SendSocket Chr$(9) + Chr$(CurInvObj) + QuadChar(0)
            End If
        End If
    End If
End Sub

Private Sub picObj_Paint()
    DrawCurInvObj
End Sub


Private Sub picTile_Click()
    CurTile = 0
    BitBlt picTile.hdc, 0, 0, 32, 32, 0, 0, 0, BLACKNESS
    picTile.Refresh
End Sub

Private Sub picTiles_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        If EditMode < 4 Then
            CurTile = Int((Y + TopY) / 32) * 7 + Int(X / 32) + 1
            RedrawTile
        Else
            NewAtt = Int(Y / 32) * 7 + Int(X / 32) + 1
            If NewAtt <= 11 Then
                Select Case NewAtt
                    Case 2, 3, 6, 7, 8, 9 'Warp, Key, News, Object, Touch Plate, Damage
                        frmMapAtt.Show 1
                    Case Else
                        CurAtt = NewAtt
                End Select
                RedrawTile
            End If
        End If
    End If
End Sub
Private Sub picViewport_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim A As Long, MapX As Long, MapY As Long
    MapX = Int(X / 32)
    MapY = Int(Y / 32)
    If MapEdit = True Then
        If MapX >= 0 And MapX <= 11 And MapY >= 0 And MapY <= 11 Then
            If Button = 1 Then
                Select Case EditMode
                    Case 0 'Ground
                        EditMap.Tile(MapX, MapY).Ground = CurTile
                        RedrawMapTile MapX, MapY
                    Case 1 'BGTile1
                        EditMap.Tile(MapX, MapY).BGTile1 = CurTile
                        RedrawMapTile MapX, MapY
                    Case 2 'BGTile2
                        EditMap.Tile(MapX, MapY).BGTile2 = CurTile
                        RedrawMapTile MapX, MapY
                    Case 3 'FGTile
                        EditMap.Tile(MapX, MapY).FGTile = CurTile
                        RedrawMapTile MapX, MapY
                    Case 4 'Attribute
                        EditMap.Tile(MapX, MapY).Att = CurAtt
                        EditMap.Tile(MapX, MapY).AttData(0) = CurAttData(0)
                        EditMap.Tile(MapX, MapY).AttData(1) = CurAttData(1)
                        EditMap.Tile(MapX, MapY).AttData(2) = CurAttData(2)
                        EditMap.Tile(MapX, MapY).AttData(3) = CurAttData(3)
                        RedrawMapTile MapX, MapY
                End Select
            ElseIf Button = 2 Then
                Select Case EditMode
                    Case 0 'Ground
                        EditMap.Tile(MapX, MapY).Ground = 0
                        RedrawMapTile MapX, MapY
                    Case 1 'BGTile1
                        EditMap.Tile(MapX, MapY).BGTile1 = 0
                        RedrawMapTile MapX, MapY
                    Case 2 'BGTile1
                        EditMap.Tile(MapX, MapY).BGTile2 = 0
                        RedrawMapTile MapX, MapY
                    Case 3 'FGTile
                        EditMap.Tile(MapX, MapY).FGTile = 0
                        RedrawMapTile MapX, MapY
                    Case 4 'Attribute
                        EditMap.Tile(MapX, MapY).Att = 0
                        EditMap.Tile(MapX, MapY).AttData(0) = 0
                        EditMap.Tile(MapX, MapY).AttData(1) = 0
                        EditMap.Tile(MapX, MapY).AttData(2) = 0
                        EditMap.Tile(MapX, MapY).AttData(3) = 0
                        RedrawMapTile MapX, MapY
                End Select
            End If
        End If
    Else
        For A = 1 To 255
            With Player(A)
                If .Map = CMap And .X = MapX And .Y = MapY Then
                    If .Guild > 0 Then
                        PrintChat "You see " + .Name + ", member of the guild " + Chr$(34) + Guild(.Guild).Name + Chr$(34) + "!", 7
                    Else
                        PrintChat "You see " + .Name + "!", 7
                    End If
                    'LastProjectile = LastProjectile Mod 10 + 1
                    'With Projectile(LastProjectile)
                    '    .Sprite = 66
                    '    .TargetType = pttPlayer
                    '    .TargetNum = A
                    '    .D = CDir
                    '    .X = CX
                    '    .Y = CY
                    '    .Frame = 0
                    'End With
                    SendSocket Chr$(27) + Chr$(A)
                    Exit Sub
                End If
            End With
        Next A
        If A = 256 Then
            For A = 0 To 5
                With Map.Monster(A)
                    If .Monster > 0 And .X = MapX And .Y = MapY Then
                        'LastProjectile = LastProjectile Mod 10 + 1
                        'With Projectile(LastProjectile)
                        '    .Sprite = 66
                        '    .TargetType = pttMonster
                        '    .TargetNum = A
                        '    .D = CDir
                        '    .X = CX
                        '    .Y = CY
                        '    .Frame = 0
                        'End With
                        PrintChat "You see a " + Monster(.Monster).Name + "!", 7
                        Exit Sub
                    End If
                End With
            Next A
            If A = 6 Then
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
