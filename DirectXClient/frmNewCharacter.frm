VERSION 5.00
Begin VB.Form frmNewCharacter 
   BackColor       =   &H0061514B&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "The Odyssey Online Classic [New Character]"
   ClientHeight    =   6240
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4380
   ControlBox      =   0   'False
   Icon            =   "frmNewCharacter.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   416
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   292
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtDesc 
      Appearance      =   0  'Flat
      BackColor       =   &H0044342E&
      ForeColor       =   &H009AADC2&
      Height          =   1215
      Left            =   120
      MaxLength       =   255
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   4440
      Width           =   4215
   End
   Begin VB.Timer SpriteTimer 
      Interval        =   250
      Left            =   0
      Top             =   0
   End
   Begin VB.OptionButton optGender 
      BackColor       =   &H0061514B&
      Caption         =   "Female"
      ForeColor       =   &H009AADC2&
      Height          =   255
      Index           =   1
      Left            =   2160
      TabIndex        =   3
      Top             =   2040
      Width           =   975
   End
   Begin VB.OptionButton optGender 
      BackColor       =   &H0061514B&
      Caption         =   "Male"
      ForeColor       =   &H009AADC2&
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
      TabIndex        =   7
      Top             =   1920
      Width           =   540
   End
   Begin VB.ComboBox cmbClass 
      Appearance      =   0  'Flat
      BackColor       =   &H0044342E&
      ForeColor       =   &H009AADC2&
      Height          =   315
      Left            =   1080
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1440
      Width           =   3255
   End
   Begin VB.TextBox txtName 
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
      Height          =   375
      Left            =   1080
      MaxLength       =   15
      TabIndex        =   0
      Top             =   120
      Width           =   3255
   End
   Begin VB.Label lblClassDescription 
      Alignment       =   2  'Center
      BackColor       =   &H0061514B&
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
      Height          =   1095
      Left            =   1560
      TabIndex        =   20
      Top             =   2850
      Width           =   2775
   End
   Begin VB.Label btnCancel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0044342E&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H009AADC2&
      Height          =   375
      Left            =   120
      TabIndex        =   19
      Top             =   5760
      Width           =   1980
   End
   Begin VB.Label btnOk 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0044342E&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Create"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H009AADC2&
      Height          =   375
      Left            =   2280
      TabIndex        =   18
      Top             =   5760
      Width           =   1980
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      BackColor       =   &H0061514B&
      Caption         =   $"frmNewCharacter.frx":1CFA
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
      Height          =   855
      Left            =   0
      TabIndex        =   17
      Top             =   480
      Width           =   4335
   End
   Begin VB.Label lblMana 
      BackColor       =   &H0061514B&
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
      ForeColor       =   &H009AADC2&
      Height          =   255
      Left            =   1080
      TabIndex        =   16
      Top             =   3600
      Width           =   735
   End
   Begin VB.Label lblEnergy 
      BackColor       =   &H0061514B&
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
      ForeColor       =   &H009AADC2&
      Height          =   255
      Left            =   1080
      TabIndex        =   15
      Top             =   3240
      Width           =   735
   End
   Begin VB.Label lblHP 
      BackColor       =   &H0061514B&
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
      ForeColor       =   &H009AADC2&
      Height          =   255
      Left            =   1080
      TabIndex        =   14
      Top             =   2880
      Width           =   735
   End
   Begin VB.Label Label8 
      BackColor       =   &H0061514B&
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
      Left            =   240
      TabIndex        =   13
      Top             =   3600
      Width           =   735
   End
   Begin VB.Label Label7 
      BackColor       =   &H0061514B&
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
      Left            =   240
      TabIndex        =   12
      Top             =   3240
      Width           =   735
   End
   Begin VB.Label Label6 
      BackColor       =   &H0061514B&
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
      Left            =   240
      TabIndex        =   11
      Top             =   2880
      Width           =   735
   End
   Begin VB.Label Label5 
      BackColor       =   &H0061514B&
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
      ForeColor       =   &H009AADC2&
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   4080
      Width           =   1935
   End
   Begin VB.Label Label4 
      BackColor       =   &H0061514B&
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
      ForeColor       =   &H009AADC2&
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   2520
      Width           =   1935
   End
   Begin VB.Label Label3 
      BackColor       =   &H0061514B&
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
      ForeColor       =   &H009AADC2&
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   2040
      Width           =   855
   End
   Begin VB.Label Label2 
      BackColor       =   &H0061514B&
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
      ForeColor       =   &H009AADC2&
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   1440
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H0061514B&
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
      ForeColor       =   &H009AADC2&
      Height          =   375
      Left            =   120
      TabIndex        =   5
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
Attribute W.VB_VarUserMemId = 1073938432
Attribute A.VB_VarUserMemId = 1073938432
Private Sub btnCancel_Click()
    Unload Me
    CloseClientSocket 0
End Sub

Private Sub btnOk_Click()
    Dim Gender As Byte, A As Long

    If Len(txtName) >= 3 Then
        A = Asc(Mid$(txtName, 1, 1))
        If (A >= 65 And A <= 90) Or (A >= 97 And A <= 122) Then
            If SwearFilter(txtName) = False And ValidName(txtName) Then
                If optGender(0) = True Then Gender = 0 Else Gender = 1

                frmWait.Show
                frmWait.lblStatus.Caption = "Creating new character ..."

                Dim ClassIndex As Integer
                ClassIndex = cmbClass.ListIndex + 1

                SendSocket Chr$(2) + Chr$(ClassIndex) + Chr$(Gender) + txtName + vbNullChar + txtDesc

                Me.Hide
            Else
                MsgBox "Invalid name!"
            End If
        Else
            MsgBox "Name must Start with a letter!", vbOKOnly + vbExclamation, TitleString
        End If
    Else
        MsgBox "Name must be atleast 3 characters long!", vbOKOnly + vbExclamation, TitleString
    End If
End Sub

Private Sub cmbClass_Click()
    Dim ClassIndex As Integer

    ClassIndex = cmbClass.ListIndex + 1

    With Class(ClassIndex)
        lblHP = .StartHP
        lblEnergy = .StartEnergy
        lblMana = .StartMana
    End With

    Select Case ClassIndex
    Case 1    'Knight
        lblClassDescription = "The most powerful fighting class.  Knights are able to use the strongest equipment and focus on hand-to-hand combat.  Very highly recommended for new players."
    Case 2    'Mage
        lblClassDescription = "Focuses on using magic to accomplish tasks instead of brute force.  Mages are frail physically compared to the other classes, and are difficult to start out with if you have never played the game before."
    Case 3    'Rogue
        lblClassDescription = "Relies on stealth, speed, and cunning to make a mark on the world.  While moderately strong physically, Rogues have much more access to magic than Knights."
    Case 4    'Cleric
        lblClassDescription = "Focused on healing and support for questing groups and guilds.  Clerics are poor fighters, but have access to a wide variety of healing and support magic.  *NOT* recommended for new players."
    End Select
End Sub


Private Sub Form_Load()
    Dim A As Long
    frmNewCharacter_Loaded = True
    For A = 1 To NumClasses
        cmbClass.AddItem Class(A).name
    Next A
    cmbClass.ListIndex = 0
End Sub
Private Sub Form_Unload(Cancel As Integer)
    frmNewCharacter_Loaded = False
End Sub


Private Sub SpriteTimer_Timer()
    Dim ClassIndex As Integer

    ClassIndex = cmbClass.ListIndex

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
        'BitBlt picSprite.hDC, 0, 0, 32, 32, hdcSprites, Frame * 32, ClassIndex * 64, SRCCOPY
        DrawToDC 0, 0, 32, 32, picSprite.hDC, DDSSprites, Frame * 32, ClassIndex * 64
    Else
        'BitBlt picSprite.hDC, 0, 0, 32, 32, hdcSprites, Frame * 32, ClassIndex * 64 + 32, SRCCOPY
        DrawToDC 0, 0, 32, 32, picSprite.hDC, DDSSprites, Frame * 32, ClassIndex * 64 + 32
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
    Dim OldString As String
    OldString = txtName.Text

    If ValidName(txtName) = False Then
        txtName.Text = ""
        Beep
    End If

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
