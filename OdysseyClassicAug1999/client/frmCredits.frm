VERSION 5.00
Begin VB.Form frmCredits 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "The Odyssey Online Classic [Credits]"
   ClientHeight    =   5295
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5775
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   353
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   385
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnOk 
      Caption         =   "Ok"
      Height          =   495
      Left            =   2160
      TabIndex        =   0
      Top             =   4560
      Width           =   1455
   End
   Begin VB.Label lblVer 
      AutoSize        =   -1  'True
      Caption         =   "Version A5"
      Height          =   195
      Left            =   4920
      TabIndex        =   7
      Top             =   4920
      Width           =   765
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Caption         =   "Assisstant Programming by Matthew T. Beutel"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   2760
      Width           =   5535
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "Special thanks to all those who donated their time to make this game possible."
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   120
      TabIndex        =   5
      Top             =   3720
      Width           =   5535
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "Character and Monster Graphics by Greg Dorando"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   3240
      Width           =   5535
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Programming by Justin E. Schumacher"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   2400
      Width           =   5535
   End
   Begin VB.Label Label2 
      Caption         =   $"frmCredits.frx":0000
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1815
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   5535
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
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
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   120
      TabIndex        =   1
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

