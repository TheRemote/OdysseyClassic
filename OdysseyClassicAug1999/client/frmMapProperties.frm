VERSION 5.00
Begin VB.Form frmMapProperties 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "The Odyssey Online Classic [Map Properties]"
   ClientHeight    =   5985
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6840
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5985
   ScaleWidth      =   6840
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtBootY 
      Height          =   375
      Left            =   5760
      MaxLength       =   2
      TabIndex        =   15
      Top             =   3840
      Width           =   975
   End
   Begin VB.TextBox txtBootX 
      Height          =   375
      Left            =   4200
      MaxLength       =   2
      TabIndex        =   14
      Top             =   3840
      Width           =   975
   End
   Begin VB.TextBox txtBootMap 
      Height          =   375
      Left            =   2520
      MaxLength       =   4
      TabIndex        =   13
      Top             =   3840
      Width           =   975
   End
   Begin VB.TextBox txtRight 
      Height          =   375
      Left            =   4080
      MaxLength       =   4
      TabIndex        =   12
      Top             =   3120
      Width           =   975
   End
   Begin VB.TextBox txtLeft 
      Height          =   375
      Left            =   1920
      MaxLength       =   4
      TabIndex        =   11
      Top             =   3120
      Width           =   975
   End
   Begin VB.TextBox txtDown 
      Height          =   375
      Left            =   4080
      MaxLength       =   4
      TabIndex        =   10
      Top             =   2640
      Width           =   975
   End
   Begin VB.TextBox txtUp 
      Height          =   375
      Left            =   1920
      MaxLength       =   4
      TabIndex        =   9
      Top             =   2640
      Width           =   975
   End
   Begin VB.ComboBox cmbNPC 
      Height          =   315
      Left            =   4080
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   2160
      Width           =   2175
   End
   Begin VB.CheckBox chkFlag 
      Caption         =   "Anyone can fight"
      Height          =   195
      Index           =   6
      Left            =   4800
      TabIndex        =   37
      Top             =   5040
      Width           =   1935
   End
   Begin VB.CheckBox chkFlag 
      Caption         =   "Can't Attack Monsters"
      Height          =   195
      Index           =   5
      Left            =   2760
      TabIndex        =   36
      Top             =   5040
      Width           =   1935
   End
   Begin VB.CheckBox chkFlag 
      Caption         =   "Double Monsters"
      Height          =   195
      Index           =   4
      Left            =   1080
      TabIndex        =   20
      Top             =   5040
      Width           =   1695
   End
   Begin VB.CheckBox chkFlag 
      Caption         =   "Monsters Start on Map"
      Height          =   195
      Index           =   3
      Left            =   4800
      TabIndex        =   19
      Top             =   4680
      Width           =   1935
   End
   Begin VB.ComboBox cmbMonster 
      Height          =   315
      Index           =   2
      Left            =   1080
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   1800
      Width           =   2175
   End
   Begin VB.ComboBox cmbMonster 
      Height          =   315
      Index           =   1
      Left            =   4080
      Style           =   2  'Dropdown List
      TabIndex        =   6
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
      Left            =   1080
      Max             =   255
      TabIndex        =   5
      Top             =   2160
      Width           =   2175
   End
   Begin VB.HScrollBar sclRate 
      Height          =   255
      Index           =   1
      Left            =   4080
      Max             =   255
      TabIndex        =   7
      Top             =   1440
      Width           =   2175
   End
   Begin VB.HScrollBar sclRate 
      Height          =   255
      Index           =   0
      Left            =   1080
      Max             =   255
      TabIndex        =   3
      Top             =   1440
      Width           =   2175
   End
   Begin VB.CheckBox chkFlag 
      Caption         =   "Always Dark"
      Height          =   195
      Index           =   2
      Left            =   3360
      TabIndex        =   18
      Top             =   4680
      Width           =   1215
   End
   Begin VB.CheckBox chkFlag 
      Caption         =   "Indoors"
      Height          =   195
      Index           =   1
      Left            =   2280
      TabIndex        =   17
      Top             =   4680
      Width           =   975
   End
   Begin VB.CheckBox chkFlag 
      Caption         =   "Friendly"
      Height          =   195
      Index           =   0
      Left            =   1080
      TabIndex        =   16
      Top             =   4680
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
      Left            =   5280
      TabIndex        =   22
      Top             =   5400
      Width           =   1455
   End
   Begin VB.CommandButton btnCancel 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   3720
      TabIndex        =   21
      Top             =   5400
      Width           =   1455
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
      Left            =   5400
      TabIndex        =   42
      Top             =   3960
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
      Left            =   3720
      TabIndex        =   41
      Top             =   3960
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
      Left            =   1800
      TabIndex        =   40
      Top             =   3960
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
      Left            =   4080
      TabIndex        =   39
      Top             =   1800
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
      Left            =   120
      TabIndex        =   38
      Top             =   3960
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
      Left            =   120
      TabIndex        =   35
      Top             =   2640
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
      Left            =   1080
      TabIndex        =   34
      Top             =   2640
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
      Left            =   3240
      TabIndex        =   33
      Top             =   2640
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
      Left            =   1080
      TabIndex        =   32
      Top             =   3240
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
      Left            =   3240
      TabIndex        =   31
      Top             =   3120
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
      TabIndex        =   30
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
      TabIndex        =   29
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
      Left            =   3360
      TabIndex        =   28
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
      Index           =   1
      Left            =   6360
      TabIndex        =   27
      Top             =   1440
      Width           =   375
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
      TabIndex        =   26
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
      TabIndex        =   25
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
      Left            =   120
      TabIndex        =   24
      Top             =   4560
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
      TabIndex        =   23
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
        .Name = txtName
        .MIDI = sclMIDI
        .ExitUp = Int(Val(txtUp))
        .ExitDown = Int(Val(txtDown))
        .ExitLeft = Int(Val(txtLeft))
        .ExitRight = Int(Val(txtRight))
        .BootLocation.Map = Int(Val(txtBootMap))
        .BootLocation.X = Int(Val(txtBootX))
        .BootLocation.Y = Int(Val(txtBootY))
        .NPC = cmbNPC.ListIndex
        For A = 0 To 2
            .MonsterSpawn(A).Monster = cmbMonster(A).ListIndex
            .MonsterSpawn(A).Rate = sclRate(A)
        Next A
        For A = 0 To 6
            If chkFlag(A) = 1 Then
                SetBit .Flags, CByte(A)
            Else
                ClearBit .Flags, CByte(A)
            End If
        Next A
    End With
    Me.Hide
End Sub

Private Sub Form_Load()
    Dim A As Long

    cmbMonster(0).AddItem "<None>"
    cmbMonster(1).AddItem "<None>"
    cmbMonster(2).AddItem "<None>"
    cmbNPC.AddItem "<None>"
    
    For A = 1 To 255
        cmbMonster(0).AddItem CStr(A) + ": " + Monster(A).Name
        cmbMonster(1).AddItem CStr(A) + ": " + Monster(A).Name
        cmbMonster(2).AddItem CStr(A) + ": " + Monster(A).Name
        cmbNPC.AddItem CStr(A) + ": " + NPC(A).Name
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


Private Sub sclRate_Change(Index As Integer)
    lblRate(Index) = sclRate(Index)
End Sub


Private Sub sclRate_Scroll(Index As Integer)
    sclRate_Change (Index)
End Sub

Private Sub txtBootMap_LostFocus()
    Dim A As Double
    A = Int(Val(txtBootMap))
    If A > 2000 Then A = 2000
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


Private Sub txtDown_LostFocus()
    Dim A As Double
    A = Int(Val(txtDown))
    If A > 2000 Then A = 2000
    If A < 0 Then A = 0
    txtDown = CStr(A)
End Sub


Private Sub txtLeft_LostFocus()
    Dim A As Double
    A = Int(Val(txtLeft))
    If A > 2000 Then A = 2000
    If A < 0 Then A = 0
    txtLeft = CStr(A)
End Sub


Private Sub txtRight_LostFocus()
    Dim A As Double
    A = Int(Val(txtRight))
    If A > 2000 Then A = 2000
    If A < 0 Then A = 0
    txtRight = CStr(A)
End Sub


Private Sub txtUp_LostFocus()
    Dim A As Double
    A = Int(Val(txtUp))
    If A > 2000 Then A = 2000
    If A < 0 Then A = 0
    txtUp = CStr(A)
End Sub
