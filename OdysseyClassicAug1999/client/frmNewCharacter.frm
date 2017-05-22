VERSION 5.00
Begin VB.Form frmNewCharacter 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "The Odyssey Online Classic [New Character]"
   ClientHeight    =   6195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4380
   ControlBox      =   0   'False
   Icon            =   "frmNewCharacter.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   413
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   292
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtDesc 
      Height          =   1215
      Left            =   120
      MaxLength       =   255
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   4320
      Width           =   4215
   End
   Begin VB.Timer SpriteTimer 
      Interval        =   250
      Left            =   0
      Top             =   0
   End
   Begin VB.OptionButton optGender 
      Caption         =   "Female"
      Height          =   255
      Index           =   1
      Left            =   2160
      TabIndex        =   3
      Top             =   2040
      Width           =   975
   End
   Begin VB.OptionButton optGender 
      Caption         =   "Male"
      Height          =   255
      Index           =   0
      Left            =   1080
      TabIndex        =   2
      Top             =   2040
      Value           =   -1  'True
      Width           =   975
   End
   Begin VB.PictureBox picSprite 
      AutoRedraw      =   -1  'True
      Height          =   540
      Left            =   3795
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   9
      Top             =   1920
      Width           =   540
   End
   Begin VB.ComboBox cmbClass 
      Height          =   315
      Left            =   1080
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1440
      Width           =   3255
   End
   Begin VB.TextBox txtName 
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
      Left            =   1080
      MaxLength       =   15
      TabIndex        =   0
      Top             =   120
      Width           =   3255
   End
   Begin VB.CommandButton btnCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   495
      Left            =   1320
      TabIndex        =   5
      Top             =   5640
      Width           =   1455
   End
   Begin VB.CommandButton btnOk 
      Caption         =   "Create"
      Enabled         =   0   'False
      Height          =   495
      Left            =   2880
      TabIndex        =   6
      Top             =   5640
      Width           =   1455
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      Caption         =   $"frmNewCharacter.frx":000C
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   0
      TabIndex        =   27
      Top             =   480
      Width           =   4335
   End
   Begin VB.Label lblIntelligence 
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
      ForeColor       =   &H00004080&
      Height          =   255
      Left            =   3120
      TabIndex        =   26
      Top             =   3600
      Width           =   735
   End
   Begin VB.Label lblEndurance 
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
      ForeColor       =   &H00004080&
      Height          =   255
      Left            =   3120
      TabIndex        =   25
      Top             =   3360
      Width           =   735
   End
   Begin VB.Label lblAgility 
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
      ForeColor       =   &H00004080&
      Height          =   255
      Left            =   3120
      TabIndex        =   24
      Top             =   3120
      Width           =   735
   End
   Begin VB.Label lblStrength 
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
      ForeColor       =   &H00004080&
      Height          =   255
      Left            =   3120
      TabIndex        =   23
      Top             =   2880
      Width           =   735
   End
   Begin VB.Label lblMana 
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
      ForeColor       =   &H00004080&
      Height          =   255
      Left            =   1080
      TabIndex        =   22
      Top             =   3360
      Width           =   735
   End
   Begin VB.Label lblEnergy 
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
      ForeColor       =   &H00004080&
      Height          =   255
      Left            =   1080
      TabIndex        =   21
      Top             =   3120
      Width           =   735
   End
   Begin VB.Label lblHP 
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
      ForeColor       =   &H00004080&
      Height          =   255
      Left            =   1080
      TabIndex        =   20
      Top             =   2880
      Width           =   735
   End
   Begin VB.Label Label12 
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
      Height          =   255
      Left            =   1920
      TabIndex        =   19
      Top             =   3600
      Width           =   1215
   End
   Begin VB.Label Label11 
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
      Height          =   255
      Left            =   1920
      TabIndex        =   18
      Top             =   3360
      Width           =   1215
   End
   Begin VB.Label Label10 
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
      Height          =   255
      Left            =   1920
      TabIndex        =   17
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Label Label9 
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
      Height          =   255
      Left            =   1920
      TabIndex        =   16
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Label Label8 
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
      Height          =   255
      Left            =   240
      TabIndex        =   15
      Top             =   3360
      Width           =   735
   End
   Begin VB.Label Label7 
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
      Height          =   255
      Left            =   240
      TabIndex        =   14
      Top             =   3120
      Width           =   735
   End
   Begin VB.Label Label6 
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
      Height          =   255
      Left            =   240
      TabIndex        =   13
      Top             =   2880
      Width           =   735
   End
   Begin VB.Label Label5 
      Caption         =   "Description (optional):"
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
      TabIndex        =   12
      Top             =   3960
      Width           =   1935
   End
   Begin VB.Label Label4 
      Caption         =   "Stats (based on class):"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   2520
      Width           =   1935
   End
   Begin VB.Label Label3 
      Caption         =   "Gender:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   2040
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Class:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   1440
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Name:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "frmNewCharacter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim D As Byte, W As Byte, A As Byte
Private Sub btnCancel_Click()
    Unload Me
    frmCharacter.Show
End Sub

Private Sub btnOk_Click()
    Dim Gender As Byte, A As Long
    
    If Len(txtName) >= 3 Then
        A = Asc(Mid$(txtName, 1, 1))
        If (A >= 65 And A <= 90) Or (A >= 97 And A <= 122) Then
            If optGender(0) = True Then Gender = 0 Else Gender = 1
                
            frmWait.Show
            frmWait.Caption = "Creating new character ..."
            
            SendSocket Chr$(2) + Chr$(cmbClass.ListIndex + 1) + Chr$(Gender) + txtName + Chr$(0) + txtDesc
            
            Me.Hide
        Else
            MsgBox "Name must start with a letter!", vbOKOnly + vbExclamation, TitleString
        End If
    Else
        MsgBox "Name must be atleast 3 characters long!", vbOKOnly + vbExclamation, TitleString
    End If
End Sub

Private Sub cmbClass_Click()
    With Class(cmbClass.ListIndex + 1)
        lblHP = .StartHP
        lblEnergy = .StartEnergy
        lblMana = .StartMana
        lblStrength = .StartStrength
        lblAgility = .StartAgility
        lblEndurance = .StartEndurance
        lblIntelligence = .StartIntelligence
    End With
End Sub


Private Sub Form_Load()
    Dim A As Long
    frmNewCharacter_Loaded = True
    For A = 1 To NumClasses
        cmbClass.AddItem Class(A).Name
    Next A
    cmbClass.ListIndex = 0
End Sub
Private Sub Form_Unload(Cancel As Integer)
    frmNewCharacter_Loaded = False
End Sub


Private Sub SpriteTimer_Timer()
    Dim Frame As Byte
    If Int(Rnd * 10) = 0 Then
        A = 2
    End If
    
    If A > 0 Then
        A = A - 1
        Frame = D * 3 + 2
    Else
        Frame = D * 3 + W
        W = 1 - W
        If Int(Rnd * 10) = 0 Then
            D = (D + 1) Mod 4
        End If
    End If
    
    If optGender(0) = True Then
        BitBlt picSprite.hdc, 0, 0, 32, 32, hdcSprites, Frame * 32, cmbClass.ListIndex * 64, SRCCOPY
    Else
        BitBlt picSprite.hdc, 0, 0, 32, 32, hdcSprites, Frame * 32, cmbClass.ListIndex * 64 + 32, SRCCOPY
    End If
    
    picSprite.Refresh
End Sub

Private Sub txtDesc_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Or (KeyAscii >= 32 And KeyAscii <= 127) Then
        'Valid Key
    Else
        KeyAscii = 0
        Beep
    End If
End Sub
Private Sub txtName_Change()
    If txtName <> "" Then
        If btnOk.Enabled = False Then
            btnOk.Enabled = True
        End If
    Else
        If btnOk.Enabled = True Then
            btnOk.Enabled = False
        End If
    End If
End Sub
Private Sub txtName_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Or (KeyAscii >= 48 And KeyAscii <= 57) Or (KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii >= 97 And KeyAscii <= 122) Or KeyAscii = 32 Then
        'Valid Key
    Else
        KeyAscii = 0
        Beep
    End If
End Sub
