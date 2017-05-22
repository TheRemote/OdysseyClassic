VERSION 5.00
Begin VB.Form frmCredits 
   BackColor       =   &H0044342E&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "The Odyssey Online Classic [Credits]"
   ClientHeight    =   7095
   ClientLeft      =   105
   ClientTop       =   330
   ClientWidth     =   5850
   ControlBox      =   0   'False
   Icon            =   "frmCredits.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   473
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   390
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label lblBorfshwitz 
      Alignment       =   2  'Center
      BackColor       =   &H0044342E&
      Caption         =   "Patrick Bukowski (Borfshwitz) - Three Wishes and Hero's quest (Hero's maps originally by Velius), mapping"
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
      Height          =   495
      Left            =   120
      TabIndex        =   14
      Top             =   4920
      Width           =   5535
   End
   Begin VB.Label lblFankadore 
      Alignment       =   2  'Center
      BackColor       =   &H0044342E&
      Caption         =   "Fankadore - Object, magic, and monster balance, mini events, mapping"
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
      Height          =   495
      Left            =   120
      TabIndex        =   13
      Top             =   4440
      Width           =   5535
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackColor       =   &H0044342E&
      Caption         =   "Server:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H009AADC2&
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Top             =   3960
      Width           =   5535
   End
   Begin VB.Label lblSmithy 
      Alignment       =   2  'Center
      BackColor       =   &H0044342E&
      Caption         =   "Original lighting and weather effects by Clay Rance (Smithy)"
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
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   1800
      Width           =   5535
   End
   Begin VB.Label btnOk 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0044342E&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Ok"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H009AADC2&
      Height          =   495
      Left            =   1560
      TabIndex        =   10
      Top             =   6480
      Width           =   2775
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackColor       =   &H0044342E&
      Caption         =   "Vivi"
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
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   3480
      Width           =   5535
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H0044342E&
      Caption         =   "Judy Shmidt (Gecky)"
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
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   3120
      Width           =   5535
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H0044342E&
      Caption         =   "Art:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H009AADC2&
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   2280
      Width           =   5535
   End
   Begin VB.Label lblBugaboo 
      Alignment       =   2  'Center
      BackColor       =   &H0044342E&
      Caption         =   "Original game by Justin E. Schumacher (Bugaboo)"
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
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   1440
      Width           =   5535
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H0044342E&
      Caption         =   "James Chambers (remote)"
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
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   5535
   End
   Begin VB.Label lblVer 
      AutoSize        =   -1  'True
      BackColor       =   &H0044342E&
      Caption         =   "Version A5"
      ForeColor       =   &H009AADC2&
      Height          =   195
      Left            =   4800
      TabIndex        =   5
      Top             =   6720
      Width           =   885
   End
   Begin VB.Label lblSpecialThanks 
      Alignment       =   2  'Center
      BackColor       =   &H0044342E&
      Caption         =   "Special thanks to all those who donated their time or money to make this game possible."
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
      Height          =   615
      Left            =   120
      TabIndex        =   4
      Top             =   5760
      Width           =   5535
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H0044342E&
      Caption         =   "Greg Dorando (Archbane)"
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
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   2760
      Width           =   5535
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H0044342E&
      Caption         =   "Programming:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H009AADC2&
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   5535
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H0044342E&
      Caption         =   "The Odyssey Online Classic Edition"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H009AADC2&
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5535
   End
End
Attribute VB_Name = "frmCredits"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub btnOk_Click()
    Unload Me
    frmMenu.Show
End Sub

Private Sub Form_Load()
    lblVer.Caption = "Version A" + CStr(ClientVer)
End Sub
