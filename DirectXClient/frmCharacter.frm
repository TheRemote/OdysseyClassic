VERSION 5.00
Begin VB.Form frmCharacter 
   BackColor       =   &H0061514B&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "The Odyssey Online Classic [Connected]"
   ClientHeight    =   7860
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8085
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7860
   ScaleWidth      =   8085
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frameCharacter 
      Appearance      =   0  'Flat
      BackColor       =   &H0044342E&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   7575
      Left            =   4560
      TabIndex        =   2
      Top             =   120
      Width           =   3495
      Begin VB.TextBox txtDesc 
         Appearance      =   0  'Flat
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
         Height          =   1815
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   27
         Top             =   5640
         Width           =   3255
      End
      Begin VB.PictureBox picSprite 
         AutoRedraw      =   -1  'True
         Height          =   540
         Left            =   2760
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   24
         Top             =   1200
         Width           =   540
      End
      Begin VB.Label lblMagicDefense 
         BackStyle       =   0  'Transparent
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
         Left            =   1200
         TabIndex        =   52
         Top             =   4920
         Width           =   375
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Magic Defense:"
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
         Left            =   120
         TabIndex        =   51
         Top             =   4800
         Width           =   975
      End
      Begin VB.Label lblDefense 
         BackStyle       =   0  'Transparent
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
         Left            =   1200
         TabIndex        =   50
         Top             =   4440
         Width           =   375
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Defense:"
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
         Left            =   120
         TabIndex        =   49
         Top             =   4440
         Width           =   975
      End
      Begin VB.Label lblDamage 
         BackStyle       =   0  'Transparent
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
         Left            =   1200
         TabIndex        =   48
         Top             =   4080
         Width           =   375
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Damage:"
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
         Left            =   120
         TabIndex        =   47
         Top             =   4080
         Width           =   975
      End
      Begin VB.Label Label24 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Concentration"
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
         Left            =   1560
         TabIndex        =   46
         Top             =   4080
         Width           =   1215
      End
      Begin VB.Label Label23 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Constitution:"
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
         Left            =   1560
         TabIndex        =   45
         Top             =   4440
         Width           =   1215
      End
      Begin VB.Label Label22 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Stamina:"
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
         Left            =   1800
         TabIndex        =   44
         Top             =   4800
         Width           =   975
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Wisdom:"
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
         Left            =   1800
         TabIndex        =   43
         Top             =   5160
         Width           =   975
      End
      Begin VB.Label lblConcentration 
         BackStyle       =   0  'Transparent
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
         Left            =   2880
         TabIndex        =   42
         Top             =   4080
         Width           =   375
      End
      Begin VB.Label lblConstitution 
         BackStyle       =   0  'Transparent
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
         Left            =   2880
         TabIndex        =   41
         Top             =   4440
         Width           =   375
      End
      Begin VB.Label lblStamina 
         BackStyle       =   0  'Transparent
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
         Left            =   2880
         TabIndex        =   40
         Top             =   4800
         Width           =   375
      End
      Begin VB.Label lblWisdom 
         BackStyle       =   0  'Transparent
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
         Left            =   2880
         TabIndex        =   39
         Top             =   5160
         Width           =   375
      End
      Begin VB.Label lblGuildRank 
         BackStyle       =   0  'Transparent
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
         Left            =   1320
         TabIndex        =   31
         Top             =   2040
         Width           =   1935
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Guild Rank:"
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
         Left            =   240
         TabIndex        =   30
         Top             =   2040
         Width           =   975
      End
      Begin VB.Label lblGuild 
         BackStyle       =   0  'Transparent
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
         Left            =   1320
         TabIndex        =   29
         Top             =   1680
         Width           =   1935
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Guild:"
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
         Left            =   240
         TabIndex        =   28
         Top             =   1680
         Width           =   975
      End
      Begin VB.Shape Shape2 
         Height          =   1935
         Left            =   120
         Top             =   480
         Width           =   3255
      End
      Begin VB.Shape Shape1 
         Height          =   3015
         Left            =   120
         Top             =   2520
         Width           =   3255
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Level:"
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
         Left            =   120
         TabIndex        =   26
         Top             =   3720
         Width           =   975
      End
      Begin VB.Label lblLevel 
         BackStyle       =   0  'Transparent
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
         Left            =   1200
         TabIndex        =   25
         Top             =   3720
         Width           =   375
      End
      Begin VB.Label lblIntelligence 
         BackStyle       =   0  'Transparent
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
         Left            =   2880
         TabIndex        =   23
         Top             =   3720
         Width           =   375
      End
      Begin VB.Label lblEndurance 
         BackStyle       =   0  'Transparent
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
         Left            =   2880
         TabIndex        =   22
         Top             =   3360
         Width           =   375
      End
      Begin VB.Label lblAgility 
         BackStyle       =   0  'Transparent
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
         Left            =   2880
         TabIndex        =   21
         Top             =   3000
         Width           =   375
      End
      Begin VB.Label lblStrength 
         BackStyle       =   0  'Transparent
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
         Left            =   2880
         TabIndex        =   20
         Top             =   2640
         Width           =   375
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Intelligence:"
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
         Left            =   1800
         TabIndex        =   19
         Top             =   3720
         Width           =   975
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Endurance:"
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
         Left            =   1800
         TabIndex        =   18
         Top             =   3360
         Width           =   975
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Agility:"
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
         Left            =   1800
         TabIndex        =   17
         Top             =   3000
         Width           =   975
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Strength:"
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
         Left            =   1800
         TabIndex        =   16
         Top             =   2640
         Width           =   975
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H0044342E&
         Caption         =   "Character Data"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H009AADC2&
         Height          =   375
         Left            =   360
         TabIndex        =   15
         Top             =   120
         Width           =   2775
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H009AADC2&
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   600
         Width           =   975
      End
      Begin VB.Label lblName 
         BackStyle       =   0  'Transparent
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
         Left            =   1320
         TabIndex        =   13
         Top             =   600
         Width           =   2055
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Class:"
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
         Left            =   240
         TabIndex        =   12
         Top             =   960
         Width           =   975
      End
      Begin VB.Label lblClass 
         BackStyle       =   0  'Transparent
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
         Left            =   1320
         TabIndex        =   11
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label lblGender 
         BackStyle       =   0  'Transparent
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
         Left            =   1320
         TabIndex        =   10
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Gender:"
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
         Left            =   240
         TabIndex        =   9
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label lblHP 
         BackStyle       =   0  'Transparent
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
         Left            =   1200
         TabIndex        =   8
         Top             =   2640
         Width           =   375
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "HP:"
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
         Left            =   120
         TabIndex        =   7
         Top             =   2640
         Width           =   975
      End
      Begin VB.Label lblEnergy 
         BackStyle       =   0  'Transparent
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
         Left            =   1200
         TabIndex        =   6
         Top             =   3000
         Width           =   375
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Energy:"
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
         Left            =   120
         TabIndex        =   5
         Top             =   3000
         Width           =   975
      End
      Begin VB.Label lblMana 
         BackStyle       =   0  'Transparent
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
         Left            =   1200
         TabIndex        =   4
         Top             =   3360
         Width           =   375
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Mana:"
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
         Left            =   120
         TabIndex        =   3
         Top             =   3360
         Width           =   975
      End
   End
   Begin VB.Frame frameNoCharacter 
      Appearance      =   0  'Flat
      BackColor       =   &H0044342E&
      ForeColor       =   &H80000008&
      Height          =   5895
      Left            =   4560
      TabIndex        =   36
      Top             =   120
      Width           =   3495
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H0044342E&
         Caption         =   "You have not yet created a character.  Click 'New Character' to do so."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H009AADC2&
         Height          =   1095
         Left            =   240
         TabIndex        =   37
         Top             =   360
         Width           =   3135
      End
   End
   Begin VB.TextBox txtMotd 
      Appearance      =   0  'Flat
      BackColor       =   &H0044342E&
      Enabled         =   0   'False
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
      Height          =   4815
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   2880
      Width           =   4335
   End
   Begin VB.Label lblMenu 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0044342E&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Delete Account"
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
      Index           =   3
      Left            =   2640
      TabIndex        =   38
      Top             =   7680
      Visible         =   0   'False
      Width           =   4335
   End
   Begin VB.Label lblMenu 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0044342E&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Change Password"
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
      Index           =   2
      Left            =   120
      TabIndex        =   35
      Top             =   1320
      Width           =   4335
   End
   Begin VB.Label lblMenu 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0044342E&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "New Character"
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
      Index           =   1
      Left            =   120
      TabIndex        =   34
      Top             =   720
      Width           =   4335
   End
   Begin VB.Label lblMenu 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0044342E&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Play"
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
      Index           =   0
      Left            =   120
      TabIndex        =   33
      Top             =   120
      Width           =   4335
   End
   Begin VB.Label lblMenu 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0044342E&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Disconnect"
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
      Index           =   4
      Left            =   120
      TabIndex        =   32
      Top             =   1920
      Width           =   4335
   End
   Begin VB.Label Label1 
      BackColor       =   &H0061514B&
      Caption         =   "Message of the day:"
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
      TabIndex        =   1
      Top             =   2520
      Width           =   1935
   End
End
Attribute VB_Name = "frmCharacter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim AHits As Long

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 65 Then
        AHits = AHits + 1
        If AHits > 3 Then
            lblMenu(3).Visible = True
        End If
    End If
End Sub

Private Sub Form_Load()
    frmCharacter_Loaded = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmCharacter_Loaded = False
End Sub

Private Sub lblMenu_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblMenu(Index).BackColor = &H61514B
End Sub


Private Sub lblMenu_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim A As Long
    
    lblMenu(Index).BackColor = &H44342E
    If X >= 0 And X <= lblMenu(Index).Width And Y >= 0 And Y <= lblMenu(Index).Height Then
        Select Case Index
            Case 0 'Play
                If Character.Class > 0 Then
                    SetMap 0
                    For A = 1 To 80
                        Guild(A).Name = vbNullString
                        With Player(A)
                            .Sprite = 0
                            .Map = 0
                        End With
                    Next A
                    With Character
                        For A = 1 To 20
                            With .Inv(A)
                                .Object = 0
                                .EquippedNum = 0
                                .Value = 0
                            End With
                        Next A
                    End With
                    keyLeft = False
                    keyRight = False
                    keyUp = False
                    keyDown = False
                    keyCtrl = False
                    keyShift = False
                    frmWait.Show
                    frmWait.lblStatus = "Receiving Game Data ..."
                    frmWait.Refresh
                    SendSocket Chr$(7) + Chr$(1) 'I wanna play
                    Unload Me
                Else
                    MsgBox "You must create a character first!", vbOKOnly + vbExclamation, TitleString
                End If
            Case 1 'New Character
                If Character.Class > 0 Then
                    MsgBox "You already have a character!"
                    'If MsgBox("Creating a new character will erase your current character, continue?", vbYesNo + vbQuestion, TitleString) = vbYes Then
                    '    frmNewCharacter.Show
                    '    Me.Hide
                    'End If
                Else
                    frmNewCharacter.Show
                    Me.Hide
                End If
            Case 2 'Change Password
                frmNewPass.Show
                Me.Hide
            Case 3 'Delete Account
                If MsgBox("This will permanently erase your character and account, are you *sure* you want to continue?", vbYesNo + vbExclamation, TitleString) = vbYes Then
                    If MsgBox("Last chance to back out -- are you *sure* you wish to delete you account?", vbYesNo + vbExclamation, TitleString) = vbYes Then
                        SendSocket Chr$(4)
                        'CloseClientSocket 0
                    End If
                End If
            Case 4 'Disconnect
                CloseClientSocket 0
        End Select
    End If
End Sub
