VERSION 5.00
Object = "{E910F8E1-8996-4EE9-90F1-3E7C64FA9829}#1.1#0"; "vbaListView6.ocx"
Begin VB.Form frmGuild 
   BackColor       =   &H0061514B&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "The Odyssey Online Classic [Guild]"
   ClientHeight    =   7440
   ClientLeft      =   90
   ClientTop       =   345
   ClientWidth     =   10065
   ControlBox      =   0   'False
   Icon            =   "frmGuild.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   496
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   671
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picGuilds 
      BackColor       =   &H0044342E&
      Height          =   1620
      Left            =   3360
      ScaleHeight     =   1560
      ScaleWidth      =   6555
      TabIndex        =   12
      Top             =   120
      Width           =   6615
      Begin VB.HScrollBar sclSprite 
         Height          =   135
         Left            =   720
         Max             =   1
         Min             =   1
         TabIndex        =   14
         Top             =   1380
         Value           =   1
         Width           =   1935
      End
      Begin VB.PictureBox picSprite 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         Height          =   540
         Left            =   1440
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   13
         Top             =   780
         Width           =   540
      End
      Begin VB.Label cmdResetStats 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0044342E&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Reset Stats"
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
         Left            =   5280
         TabIndex        =   32
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label lblMembers 
         BackColor       =   &H0061514B&
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
         Left            =   4320
         TabIndex        =   31
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label cptMembers 
         BackColor       =   &H0061514B&
         BackStyle       =   0  'Transparent
         Caption         =   "Members:"
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
         Left            =   3360
         TabIndex        =   30
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label cmdSellSprite 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0044342E&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Clear Sprite"
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
         Left            =   2280
         TabIndex        =   29
         Top             =   1080
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label lblDeaths 
         BackColor       =   &H0061514B&
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
         Left            =   4320
         TabIndex        =   28
         Top             =   840
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label lblKills 
         BackColor       =   &H0061514B&
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
         Left            =   4320
         TabIndex        =   27
         Top             =   480
         Width           =   855
      End
      Begin VB.Label lblCreated 
         BackColor       =   &H0061514B&
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
         Left            =   4320
         TabIndex        =   26
         Top             =   120
         Width           =   1335
      End
      Begin VB.Label cptDeaths 
         BackColor       =   &H0061514B&
         BackStyle       =   0  'Transparent
         Caption         =   "Deaths:"
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
         Left            =   3360
         TabIndex        =   25
         Top             =   840
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label cptKills 
         BackColor       =   &H0061514B&
         BackStyle       =   0  'Transparent
         Caption         =   "Kills:"
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
         Left            =   3360
         TabIndex        =   24
         Top             =   480
         Width           =   735
      End
      Begin VB.Label cptCreated 
         BackColor       =   &H0061514B&
         BackStyle       =   0  'Transparent
         Caption         =   "Created:"
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
         Left            =   3360
         TabIndex        =   23
         Top             =   120
         Width           =   735
      End
      Begin VB.Label cmdBuySprite 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0044342E&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Buy Sprite"
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
         Left            =   2280
         TabIndex        =   22
         Top             =   840
         Width           =   975
      End
      Begin VB.Label cptSprite 
         BackColor       =   &H0061514B&
         BackStyle       =   0  'Transparent
         Caption         =   "Sprite:"
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
         TabIndex        =   21
         Top             =   930
         Width           =   615
      End
      Begin VB.Label cptName 
         BackColor       =   &H0061514B&
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
         Left            =   120
         TabIndex        =   20
         Top             =   120
         Width           =   615
      End
      Begin VB.Label lblName 
         BackColor       =   &H0061514B&
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
         Left            =   840
         TabIndex        =   19
         Top             =   120
         Width           =   2535
      End
      Begin VB.Label cptHall 
         BackColor       =   &H0061514B&
         BackStyle       =   0  'Transparent
         Caption         =   "Hall:"
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
         TabIndex        =   18
         Top             =   480
         Width           =   615
      End
      Begin VB.Label lblHall 
         BackColor       =   &H0061514B&
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
         Left            =   840
         TabIndex        =   17
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label btnMoveOut 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0044342E&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Leave Hall"
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
         Left            =   2280
         TabIndex        =   16
         Top             =   480
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label btnDisband 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0044342E&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Disband"
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
         Left            =   5280
         TabIndex        =   15
         Top             =   1200
         Width           =   1215
      End
   End
   Begin vbalListViewLib6.vbalListViewCtl lstViewMembers 
      Height          =   2415
      Left            =   3360
      TabIndex        =   11
      Top             =   1815
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   4260
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   10137026
      BackColor       =   4469806
      MultiSelect     =   -1  'True
      LabelEdit       =   0   'False
      AutoArrange     =   0   'False
      HeaderButtons   =   0   'False
      HeaderTrackSelect=   0   'False
      HideSelection   =   0   'False
      InfoTips        =   0   'False
      ScaleMode       =   3
   End
   Begin VB.Timer timeReEnable 
      Interval        =   2000
      Left            =   4920
      Top             =   4680
   End
   Begin VB.Timer SpriteTimer 
      Interval        =   250
      Left            =   5400
      Top             =   4680
   End
   Begin VB.ListBox lstDeclarations 
      BackColor       =   &H0044342E&
      ForeColor       =   &H009AADC2&
      Height          =   1425
      Left            =   3360
      TabIndex        =   1
      Top             =   5040
      Width           =   6615
   End
   Begin vbalListViewLib6.vbalListViewCtl lstGuilds 
      Height          =   7215
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   12726
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   10137026
      BackColor       =   4469806
      MultiSelect     =   -1  'True
      LabelEdit       =   0   'False
      AutoArrange     =   0   'False
      HeaderButtons   =   0   'False
      HeaderTrackSelect=   0   'False
      HideSelection   =   0   'False
      InfoTips        =   0   'False
      ScaleMode       =   3
   End
   Begin VB.Label cptDeclarations 
      BackColor       =   &H0061514B&
      Caption         =   "Declarations:"
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
      Left            =   3360
      TabIndex        =   10
      Top             =   4680
      Width           =   1575
   End
   Begin VB.Label btnRemoveMember 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0044342E&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Remove"
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
      Height          =   300
      Left            =   9000
      TabIndex        =   9
      Top             =   4320
      Width           =   975
   End
   Begin VB.Label btnAddDeclaration 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0044342E&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Add Declaration"
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
      Left            =   3360
      TabIndex        =   8
      Top             =   6600
      Width           =   1815
   End
   Begin VB.Label btnRemoveDeclaration 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0044342E&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Remove Declaration"
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
      Left            =   7680
      TabIndex        =   7
      Top             =   6600
      Width           =   2295
   End
   Begin VB.Label btnOk 
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
      Left            =   3360
      TabIndex        =   6
      Top             =   6960
      Width           =   6615
   End
   Begin VB.Label btnRank 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0044342E&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Initiate"
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
      Height          =   300
      Index           =   0
      Left            =   3360
      TabIndex        =   5
      Top             =   4320
      Width           =   855
   End
   Begin VB.Label btnRank 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0044342E&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Member"
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
      Height          =   300
      Index           =   1
      Left            =   4320
      TabIndex        =   4
      Top             =   4320
      Width           =   855
   End
   Begin VB.Label btnRank 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0044342E&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Lord"
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
      Height          =   300
      Index           =   2
      Left            =   5280
      TabIndex        =   3
      Top             =   4320
      Width           =   855
   End
   Begin VB.Label btnRank 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0044342E&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Founder"
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
      Height          =   300
      Index           =   3
      Left            =   6240
      TabIndex        =   2
      Top             =   4320
      Width           =   855
   End
End
Attribute VB_Name = "frmGuild"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public A As Long, D As Long, W As Long
Attribute D.VB_VarUserMemId = 1073938432
Attribute W.VB_VarUserMemId = 1073938432
Public CurrentGuild As Long
Attribute CurrentGuild.VB_VarUserMemId = 1073938435
Public CurrentSprite As Long
Attribute CurrentSprite.VB_VarUserMemId = 1073938436
Public ReEnable As Boolean
Attribute ReEnable.VB_VarUserMemId = 1073938437
Public MemberCount As Long
Attribute MemberCount.VB_VarUserMemId = 1073938438

Sub UpdateForm()
    If CurrentGuild > 0 Then
        SendSocket Chr$(39) + Chr$(CurrentGuild)
    End If

    btnRank(0).Visible = False
    btnRank(1).Visible = False
    btnRank(2).Visible = False
    btnRank(3).Visible = False
    btnRemoveMember.Visible = False
    btnAddDeclaration.Visible = False
    btnRemoveDeclaration.Visible = False
    btnDisband.Visible = False
    cmdBuySprite.Visible = False
    cmdSellSprite.Visible = False
    cmdResetStats.Visible = False
    sclSprite.Visible = False
End Sub


Private Sub btnAddDeclaration_Click()
    frmDeclaration.Show 1
    If Character.GuildRank >= 2 Then
        SendSocket Chr$(37) + Chr$(TempVar3) + Chr$(TempVar2)
        UpdateForm
    End If
End Sub

Private Sub btnDisband_Click()
    If MsgBox("Are you sure you wish to disband your guild?  Your guild will be delete, and you will not get refunded for any of the guild fees nor will you get the money in your guild's bank account.  Continue?", vbYesNo + vbQuestion, TitleString) = vbYes Then
        SendSocket Chr$(42)
        Unload Me
    End If
End Sub

Private Sub btnMoveOut_Click()
    If MsgBox("Are you sure you wish to leave your guild hall?", vbYesNo + vbQuestion, TitleString) = vbYes Then
        SendSocket Chr$(44)
        UpdateForm
    End If
End Sub

Private Sub btnOk_Click()
    Unload Me
End Sub

Private Sub btnRank_Click(index As Integer)
    Dim SelectedIndex As Long
    SelectedIndex = -1
    On Error Resume Next
    SelectedIndex = Val(lstViewMembers.SelectedItem.key)
    On Error GoTo 0

    If SelectedIndex > -1 And Character.GuildRank >= 2 Then
        If Character.name = lstViewMembers.SelectedItem Then
            Exit Sub
        End If

        If index = 3 Then
            If Character.GuildRank = 3 Then
                SendSocket Chr$(36) + Chr$(SelectedIndex) + Chr$(index)
            End If
        Else
            SendSocket Chr$(36) + Chr$(SelectedIndex) + Chr$(index)
        End If

        UpdateForm
    End If
End Sub

Private Sub btnRemoveDeclaration_Click()
    If lstDeclarations.ListIndex > -1 Then
        SendSocket Chr$(38) + Chr$(lstDeclarations.ItemData(lstDeclarations.ListIndex))
        UpdateForm
    End If
End Sub
Private Sub btnRemoveMember_Click()
    Dim SelectedIndex As Long
    SelectedIndex = -1
    On Error Resume Next
    SelectedIndex = Val(lstViewMembers.SelectedItem.key)
    On Error GoTo 0

    If SelectedIndex > -1 And Character.GuildRank >= 2 Then
        If Character.name = lstViewMembers.SelectedItem Then
            Exit Sub
        End If

        Dim St As String
        If lstViewMembers.ListItems.count <= 3 Then
            MsgBox "Because there are only " + CStr(lstViewMembers.ListItems.count) + " members in your guild, removing this member would cause your guild to be deleted."
        Else
            If MsgBox("Are you sure you wish to kick " + Chr$(34) + lstViewMembers.SelectedItem + Chr$(34) + " out of the guild?" + St, vbYesNo + vbQuestion, TitleString) = vbYes Then
                SendSocket Chr$(35) + Chr$(SelectedIndex)
                UpdateForm
            End If
        End If
    End If
End Sub

Private Sub cmdBuySprite_Click()
    If Not Me.lblHall = "<none>" Then
        If MsgBox("Are you sure you wish to buy this guild sprite?", vbYesNo + vbQuestion, TitleString) = vbYes Then
            SendSocket Chr$(81) + DoubleChar$(CurrentSprite)
        End If
    Else
        MsgBox "You must own a hall to buy a sprite!"
    End If
End Sub

Private Sub cmdResetStats_Click()
    If MsgBox("Are you sure you wish to reset your guild's main stats?  This will not reset personal stats.", vbYesNo + vbQuestion, TitleString) = vbYes Then
        SendSocket Chr$(102)
    End If
End Sub

Private Sub cmdSellSprite_Click()
    If MsgBox("Are you sure you wish to clear your guild's sprite?", vbYesNo + vbQuestion, TitleString) = vbYes Then
        SendSocket Chr$(101)
    End If
End Sub

Private Sub Form_Load()
    frmGuild_Loaded = True

    Dim Column As cColumn

    With lstViewMembers
        .View = eViewDetails
        .FullRowSelect = True

        Set Column = .Columns.Add(, "NAME", "Name")
        Column.Tag = "The guild member's name"
        Column.Width = 105
        Set Column = .Columns.Add(, "RANK", "Rank")
        Column.Tag = "The guild member's rank in the guild"
        Column.Width = 70
        Column.SortType = eLVSortItemData
        Column.SortOrder = eSortOrderDescending
        Set Column = .Columns.Add(, "JOINDATE", "Joined")
        Column.Tag = "The date the member joined the guild"
        Column.Width = 70
        Set Column = .Columns.Add(, "KILLS", "Kills")
        Column.Tag = "The amount of kills the player has got since joining the guild"
        Column.Width = 50
        
        If Character.Access > 0 Then
            Set Column = .Columns.Add(, "DEATHS", "Deaths")
            Column.Tag = "The amount of deaths the player has got since joining the guild"
            Column.Width = 50
            Set Column = .Columns.Add(, "KDRATIO", "K/D Ratio")
            Column.Tag = "The ratio of kills to deaths this player has.  Higher is better."
            Column.Width = 70
            
            frmGuild.lblDeaths.Visible = True
            frmGuild.cptDeaths.Visible = True
        End If
    End With

    With lstGuilds
        .View = eViewDetails
        .FullRowSelect = True

        Set Column = .Columns.Add(, "GUILDINDEX", "GuildIndex")
        Column.Tag = "The index of the guild"
        Column.Width = 0
        Set Column = .Columns.Add(, "NAME", "Name")
        Column.Tag = "The guild's name"
        Column.Width = 135
        Set Column = .Columns.Add(, "MEMBERS", "Members")
        Column.Tag = "The number of members in the guild"
        Column.Width = 60
        Column.SortType = eLVSortNumeric
        Column.SortOrder = eSortOrderDescending
    End With

    CurrentGuild = -1

    If Character.Guild > 0 Then
        CurrentGuild = Character.Guild
    End If

    Dim A As Long
    For A = 1 To MaxGuilds
        If Guild(A).name <> "" Then
            AddGuild A, Guild(A).name, Guild(A).MemberCount, False
            If CurrentGuild = A Then
                lstGuilds_ItemClick lstGuilds.ListItems(lstGuilds.ListItems.count)
            End If
        End If
    Next A

    lstGuilds.ListItems.SortItems

    If lstGuilds.ListItems.count > 0 Then
        If CurrentGuild = -1 Then
            lstGuilds_ItemClick lstGuilds.ListItems(1)
        End If
    End If

    sclSprite.max = MaxSprite
    CurrentSprite = 1
    UpdateForm
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmGuild_Loaded = False
End Sub

Private Sub lstDeclarations_Click()
    If CurrentGuild = Character.Guild And Character.GuildRank >= 2 Then
        btnRemoveDeclaration.Enabled = True
    End If
End Sub

Private Sub lstGuilds_ItemClick(Item As vbalListViewLib6.cListItem)
    On Error GoTo DoNothing

    If ReEnable = False Then
        If Not CurrentGuild = Val(Item.key) Then
            CurrentGuild = Val(Item.key)
            SendSocket Chr$(39) + Chr$(Val(Item.key))
            lstGuilds.ForeColor = &HFFFFFF
            lstGuilds.Refresh
            ReEnable = True
        End If
    Else
        Beep
    End If

DoNothing:

End Sub

Private Sub sclSprite_Change()
    If BannedSprite(sclSprite.value) = False Then
        CurrentSprite = sclSprite.value
    End If
End Sub

Private Sub SpriteTimer_Timer()
    If CurrentSprite > 0 Then
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

        DrawToDC 0, 0, 32, 32, picSprite.hDC, DDSSprites, Frame * 32, (CurrentSprite - 1) * 32
    Else
        picSprite.Cls
    End If

    picSprite.Refresh
End Sub

Private Sub timeReEnable_Timer()
    If ReEnable = True Then
        ReEnable = False
        lstGuilds.ForeColor = &H9AADC2
        lstGuilds.Refresh
    End If
End Sub

Public Sub AddMember(name As String, GuildIndex As Long, Rank As Long, Kills As Long, Deaths As Long, JoinedDate As Long)
    If frmGuild.Visible = False Then Exit Sub

    Dim Item As cListItem

    Dim RankString As String
    RankString = Choose(Rank + 1, "Initiate", "Member", "Lord", "Founder")

    Set Item = lstViewMembers.ListItems.Add(, GuildIndex, name)

    With Item.SubItems(1)
        .Caption = RankString
    End With
    With Item.SubItems(2)
        .Caption = CDate(JoinedDate)
    End With
    With Item.SubItems(3)
        .Caption = Kills
    End With
    Item.ItemData = Rank
    If Character.Access > 0 Then
        With Item.SubItems(4)
            .Caption = Deaths
        End With
        With Item.SubItems(5)
            Dim KDR As Double
            If Deaths > 0 Then
                KDR = Round((Kills / Deaths), 2)
            Else
                KDR = 0
            End If
            .Caption = KDR
        End With
    End If

    lstViewMembers.ListItems.SortItems

    MemberCount = MemberCount + 1
    lblMembers.Caption = CStr(MemberCount)
End Sub

Public Sub AddGuild(GuildIndex As Long, name As String, Members As Byte, Optional Resort As Boolean = True)
    Dim Item As cListItem

    Set Item = lstGuilds.ListItems.Add(, GuildIndex, name)

    With Item.SubItems(1)
        .Caption = name
    End With
    With Item.SubItems(2)
        .Caption = Members
    End With

    If Resort = True Then
        lstGuilds.ListItems.SortItems
    End If
End Sub

Private Sub lstViewMembers_ColumnClick(Column As cColumn)
' Sort according to the column type:
    Select Case Column.key
    Case "NAME"
        Column.SortType = eLVSortString
        Column.SortOrder = NewSortOrder(Column.SortOrder)
    Case "JOINDATE"
        Column.SortType = eLVSortDate
        Column.SortOrder = NewSortOrder(Column.SortOrder)
    Case "RANK"
        Column.SortType = eLVSortItemData
        Column.SortOrder = NewSortOrder(Column.SortOrder)
    Case "KILLS"
        Column.SortType = eLVSortNumeric
        Column.SortOrder = NewSortOrder(Column.SortOrder)
    Case "DEATHS"
        Column.SortType = eLVSortNumeric
        Column.SortOrder = NewSortOrder(Column.SortOrder)
    Case "KDRATIO"
        Column.SortType = eLVSortNumeric
        Column.SortOrder = NewSortOrder(Column.SortOrder)
    End Select
    lstViewMembers.ListItems.SortItems
End Sub

Private Function NewSortOrder(ByVal SortOrder As ESortOrderConstants) As ESortTypeConstants
    Select Case SortOrder
    Case eSortOrderNone, eSortOrderDescending
        NewSortOrder = eSortOrderAscending
    Case eSortOrderAscending
        NewSortOrder = eSortOrderDescending
    End Select
End Function

Private Sub lstGuilds_ColumnClick(Column As cColumn)
' Sort according to the column type:
    Select Case Column.key
    Case "NAME"
        Column.SortType = eLVSortString
        Column.SortOrder = NewSortOrder(Column.SortOrder)
    Case "MEMBERS"
        Column.SortType = eLVSortNumeric
        Column.SortOrder = NewSortOrder(Column.SortOrder)
    End Select
    lstGuilds.ListItems.SortItems
End Sub
