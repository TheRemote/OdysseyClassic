VERSION 5.00
Begin VB.Form frmMapAtt 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "The Odyssey Online Classic [Map Attribute]"
   ClientHeight    =   3090
   ClientLeft      =   150
   ClientTop       =   225
   ClientWidth     =   4860
   ControlBox      =   0   'False
   Icon            =   "frmMapAtt.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   4860
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picAtt7 
      Height          =   1455
      Left            =   120
      ScaleHeight     =   1395
      ScaleWidth      =   4035
      TabIndex        =   15
      Top             =   600
      Visible         =   0   'False
      Width           =   4095
      Begin VB.TextBox lblAtt7Val 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   3240
         TabIndex        =   36
         Text            =   "0"
         Top             =   480
         Width           =   735
      End
      Begin VB.TextBox lblAtt7Obj 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   3240
         TabIndex        =   35
         Text            =   "0"
         Top             =   120
         Width           =   735
      End
      Begin VB.HScrollBar sclAtt7Val 
         Height          =   255
         Left            =   720
         Max             =   32000
         TabIndex        =   17
         Top             =   480
         Width           =   2415
      End
      Begin VB.HScrollBar sclAtt7Obj 
         Height          =   255
         Left            =   720
         Max             =   1000
         Min             =   1
         TabIndex        =   16
         Top             =   120
         Value           =   1
         Width           =   2415
      End
      Begin VB.Label lblValue 
         Caption         =   "Value"
         Height          =   255
         Left            =   120
         TabIndex        =   38
         Top             =   480
         Width           =   615
      End
      Begin VB.Label lblObj 
         Caption         =   "Object:"
         Height          =   255
         Left            =   120
         TabIndex        =   37
         Top             =   120
         Width           =   615
      End
      Begin VB.Label lblObjName 
         Caption         =   "Name:"
         Height          =   255
         Left            =   120
         TabIndex        =   33
         Top             =   840
         Width           =   495
      End
      Begin VB.Label lblName 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   720
         TabIndex        =   32
         Top             =   840
         Width           =   3015
      End
   End
   Begin VB.PictureBox picAtt20 
      Height          =   1815
      Left            =   120
      ScaleHeight     =   1755
      ScaleWidth      =   4035
      TabIndex        =   61
      Top             =   600
      Visible         =   0   'False
      Width           =   4095
      Begin VB.CheckBox chkAtt20Direction 
         Caption         =   "Right"
         Height          =   375
         Index           =   3
         Left            =   3240
         TabIndex        =   74
         Top             =   120
         Width           =   735
      End
      Begin VB.CheckBox chkAtt20Direction 
         Caption         =   "Left"
         Height          =   375
         Index           =   2
         Left            =   2640
         TabIndex        =   73
         Top             =   120
         Width           =   735
      End
      Begin VB.CheckBox chkAtt20Direction 
         Caption         =   "Down"
         Height          =   375
         Index           =   1
         Left            =   1920
         TabIndex        =   72
         Top             =   120
         Width           =   735
      End
      Begin VB.CheckBox chkAtt20Direction 
         Caption         =   "Up"
         Height          =   375
         Index           =   0
         Left            =   1320
         TabIndex        =   71
         Top             =   120
         Width           =   735
      End
      Begin VB.CheckBox chkAtt20Indoor 
         Caption         =   "Indoor"
         Height          =   375
         Left            =   2640
         TabIndex        =   70
         Top             =   1320
         Width           =   1095
      End
      Begin VB.CheckBox chkAtt20Outdoor 
         Caption         =   "Outdoor"
         Height          =   375
         Left            =   1680
         TabIndex        =   69
         Top             =   1320
         Width           =   1095
      End
      Begin VB.CheckBox chkAtt20Wall 
         Caption         =   "Wall"
         Height          =   375
         Left            =   960
         TabIndex        =   64
         Top             =   1320
         Width           =   735
      End
      Begin VB.HScrollBar sclAtt20X 
         Height          =   255
         Left            =   960
         Max             =   32
         TabIndex        =   63
         Top             =   600
         Width           =   2415
      End
      Begin VB.HScrollBar sclAtt20Y 
         Height          =   255
         Left            =   960
         Max             =   32
         TabIndex        =   62
         Top             =   960
         Width           =   2415
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Split Direction:"
         Height          =   255
         Left            =   120
         TabIndex        =   75
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label cptAtt20Y 
         Alignment       =   1  'Right Justify
         Caption         =   "Y:"
         Height          =   255
         Left            =   120
         TabIndex        =   68
         Top             =   960
         Width           =   615
      End
      Begin VB.Label cptAtt20X 
         Alignment       =   1  'Right Justify
         Caption         =   "X:"
         Height          =   255
         Left            =   120
         TabIndex        =   67
         Top             =   600
         Width           =   615
      End
      Begin VB.Label lblAtt20X 
         Alignment       =   2  'Center
         Caption         =   "0"
         Height          =   255
         Left            =   3480
         TabIndex        =   66
         Top             =   600
         Width           =   375
      End
      Begin VB.Label lblAtt20Y 
         Alignment       =   2  'Center
         Caption         =   "0"
         Height          =   255
         Left            =   3480
         TabIndex        =   65
         Top             =   960
         Width           =   375
      End
   End
   Begin VB.PictureBox picAtt19 
      Height          =   1815
      Left            =   120
      ScaleHeight     =   1755
      ScaleWidth      =   4035
      TabIndex        =   53
      Top             =   600
      Visible         =   0   'False
      Width           =   4095
      Begin VB.HScrollBar sclIntensity 
         Height          =   255
         Left            =   960
         Max             =   255
         TabIndex        =   58
         Top             =   600
         Width           =   2415
      End
      Begin VB.HScrollBar sclRadius 
         Height          =   255
         Left            =   960
         Max             =   255
         Min             =   1
         TabIndex        =   57
         Top             =   240
         Value           =   100
         Width           =   2415
      End
      Begin VB.CheckBox chkWall 
         Caption         =   "Wall"
         Height          =   375
         Left            =   120
         TabIndex        =   56
         Top             =   960
         Width           =   735
      End
      Begin VB.Label lblIntensity 
         Alignment       =   2  'Center
         Caption         =   "0"
         Height          =   255
         Left            =   3480
         TabIndex        =   60
         Top             =   600
         Width           =   375
      End
      Begin VB.Label lblRadius 
         Alignment       =   2  'Center
         Caption         =   "100"
         Height          =   255
         Left            =   3480
         TabIndex        =   59
         Top             =   240
         Width           =   375
      End
      Begin VB.Label cptRadius 
         Caption         =   "Radius"
         Height          =   255
         Left            =   120
         TabIndex        =   55
         Top             =   240
         Width           =   735
      End
      Begin VB.Label cptIntensity 
         Caption         =   "Intensity:"
         Height          =   255
         Left            =   120
         TabIndex        =   54
         Top             =   600
         Width           =   735
      End
   End
   Begin VB.PictureBox picAtt17 
      Height          =   1815
      Left            =   120
      ScaleHeight     =   1755
      ScaleWidth      =   4035
      TabIndex        =   41
      Top             =   600
      Visible         =   0   'False
      Width           =   4095
      Begin VB.CheckBox chkNoMonster 
         Caption         =   "No Monster"
         Height          =   255
         Left            =   2880
         TabIndex        =   52
         Top             =   1440
         Width           =   1215
      End
      Begin VB.CheckBox chkDirections 
         Caption         =   "From Left"
         Height          =   255
         Index           =   5
         Left            =   240
         TabIndex        =   51
         Top             =   720
         Width           =   1695
      End
      Begin VB.CheckBox chkDirections 
         Caption         =   "To Left"
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   50
         Top             =   480
         Width           =   975
      End
      Begin VB.CheckBox chkDirections 
         Caption         =   "From Below"
         Height          =   255
         Index           =   3
         Left            =   1440
         TabIndex        =   49
         Top             =   1200
         Width           =   1695
      End
      Begin VB.CheckBox chkDirections 
         Caption         =   "To Below"
         Height          =   255
         Index           =   2
         Left            =   1440
         TabIndex        =   48
         Top             =   960
         Width           =   1695
      End
      Begin VB.CheckBox chkKeep 
         Caption         =   "Keep"
         Height          =   255
         Left            =   120
         TabIndex        =   47
         Top             =   1440
         Width           =   1215
      End
      Begin VB.CheckBox chkDirections 
         Caption         =   "To Right"
         Height          =   255
         Index           =   7
         Left            =   2880
         TabIndex        =   46
         Top             =   720
         Width           =   1695
      End
      Begin VB.CheckBox chkDirections 
         Caption         =   "From Right"
         Height          =   255
         Index           =   6
         Left            =   2880
         TabIndex        =   45
         Top             =   480
         Width           =   1695
      End
      Begin VB.CheckBox chkDirections 
         Caption         =   "To Above"
         Height          =   255
         Index           =   1
         Left            =   1440
         TabIndex        =   44
         Top             =   360
         Width           =   1695
      End
      Begin VB.CheckBox chkDirections 
         Caption         =   "From Above"
         Height          =   255
         Index           =   0
         Left            =   1440
         TabIndex        =   43
         Top             =   120
         Width           =   1695
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Height          =   255
         Left            =   840
         TabIndex        =   42
         Top             =   480
         Width           =   2415
      End
   End
   Begin VB.PictureBox picAtt3 
      Height          =   1455
      Left            =   120
      ScaleHeight     =   1395
      ScaleWidth      =   4035
      TabIndex        =   13
      Top             =   600
      Visible         =   0   'False
      Width           =   4095
      Begin VB.TextBox lblAtt3Key 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   3360
         TabIndex        =   40
         Text            =   "1"
         Top             =   120
         Width           =   615
      End
      Begin VB.HScrollBar sclAtt3Key 
         Height          =   255
         Left            =   840
         Max             =   510
         Min             =   1
         TabIndex        =   14
         Top             =   120
         Value           =   1
         Width           =   2415
      End
      Begin VB.Label lblKey 
         Caption         =   "Key:"
         Height          =   255
         Left            =   240
         TabIndex        =   39
         Top             =   120
         Width           =   495
      End
      Begin VB.Label lblObject 
         Alignment       =   2  'Center
         Height          =   255
         Left            =   840
         TabIndex        =   34
         Top             =   480
         Width           =   2415
      End
   End
   Begin VB.PictureBox picAtt9 
      Height          =   1455
      Left            =   120
      ScaleHeight     =   1395
      ScaleWidth      =   4035
      TabIndex        =   28
      Top             =   600
      Visible         =   0   'False
      Width           =   4095
      Begin VB.HScrollBar sclAtt9Damage 
         Height          =   255
         Left            =   1080
         Max             =   50
         Min             =   1
         TabIndex        =   29
         Top             =   120
         Value           =   1
         Width           =   2295
      End
      Begin VB.Label Label14 
         Caption         =   "Damage:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   31
         Top             =   120
         Width           =   855
      End
      Begin VB.Label lblAtt9Damage 
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
         Height          =   375
         Left            =   3480
         TabIndex        =   30
         Top             =   120
         Width           =   495
      End
   End
   Begin VB.PictureBox picAtt8 
      Height          =   1455
      Left            =   120
      ScaleHeight     =   1395
      ScaleWidth      =   4035
      TabIndex        =   18
      Top             =   600
      Visible         =   0   'False
      Width           =   4095
      Begin VB.HScrollBar sclAtt8Hall 
         Height          =   255
         LargeChange     =   10
         Left            =   720
         Max             =   255
         TabIndex        =   25
         Top             =   840
         Width           =   2415
      End
      Begin VB.HScrollBar sclAtt8X 
         Height          =   255
         Left            =   720
         Max             =   11
         TabIndex        =   20
         Top             =   120
         Width           =   2415
      End
      Begin VB.HScrollBar sclAtt8Y 
         Height          =   255
         Left            =   720
         Max             =   11
         TabIndex        =   19
         Top             =   480
         Width           =   2415
      End
      Begin VB.Label lblAtt8Hall 
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
         Height          =   375
         Left            =   3240
         TabIndex        =   27
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label12 
         Caption         =   "Hall:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   26
         Top             =   840
         Width           =   495
      End
      Begin VB.Label lblAtt8X 
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
         Height          =   375
         Left            =   3240
         TabIndex        =   24
         Top             =   120
         Width           =   735
      End
      Begin VB.Label Label10 
         Caption         =   "X:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   23
         Top             =   120
         Width           =   495
      End
      Begin VB.Label lblAtt8Y 
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
         Height          =   375
         Left            =   3240
         TabIndex        =   22
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label5 
         Caption         =   "Y:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   21
         Top             =   480
         Width           =   495
      End
   End
   Begin VB.PictureBox picAtt2 
      Height          =   1455
      Left            =   120
      ScaleHeight     =   1395
      ScaleWidth      =   4035
      TabIndex        =   3
      Top             =   600
      Visible         =   0   'False
      Width           =   4095
      Begin VB.HScrollBar sclAtt2Y 
         Height          =   255
         Left            =   720
         Max             =   11
         TabIndex        =   9
         Top             =   840
         Width           =   2415
      End
      Begin VB.HScrollBar sclAtt2X 
         Height          =   255
         Left            =   720
         Max             =   11
         TabIndex        =   8
         Top             =   480
         Width           =   2415
      End
      Begin VB.HScrollBar sclAtt2Map 
         Height          =   255
         LargeChange     =   25
         Left            =   720
         Max             =   2000
         Min             =   1
         TabIndex        =   7
         Top             =   120
         Value           =   1
         Width           =   2415
      End
      Begin VB.Label lblAtt2Y 
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
         Height          =   375
         Left            =   3240
         TabIndex        =   12
         Top             =   840
         Width           =   735
      End
      Begin VB.Label lblAtt2X 
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
         Height          =   375
         Left            =   3240
         TabIndex        =   11
         Top             =   480
         Width           =   735
      End
      Begin VB.Label lblAtt2Map 
         Alignment       =   2  'Center
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3240
         TabIndex        =   10
         Top             =   120
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "Y:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   840
         Width           =   495
      End
      Begin VB.Label Label2 
         Caption         =   "X:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   480
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "Map:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   120
         Width           =   495
      End
   End
   Begin VB.CommandButton btnCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   495
      Left            =   1680
      TabIndex        =   1
      Top             =   2520
      Width           =   1455
   End
   Begin VB.CommandButton btnOk 
      Caption         =   "Ok"
      Height          =   495
      Left            =   3240
      TabIndex        =   0
      Top             =   2520
      Width           =   1455
   End
   Begin VB.Label lblAtt 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   4095
   End
End
Attribute VB_Name = "frmMapAtt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub btnOk_Click()
    Dim A As Long
    CurAtt = NewAtt
    Select Case CurAtt
    Case 2    'Warp
        CurAttData(0) = Int(sclAtt2Map / 256)
        CurAttData(1) = sclAtt2Map Mod 256
        CurAttData(2) = sclAtt2X
        CurAttData(3) = sclAtt2Y
    Case 3    'Key
        CurAttData(0) = sclAtt3Key Mod 256
        CurAttData(1) = Int(sclAtt3Key / 256)
        CurAttData(2) = 0
        CurAttData(3) = 0
    Case 7    'Object
        CurAttData(0) = sclAtt7Obj Mod 256
        CurAttData(1) = Int(sclAtt7Obj / 256)
        CurAttData(2) = Int(sclAtt7Val / 256)
        CurAttData(3) = sclAtt7Val Mod 256
    Case 8    'Touch Plate
        CurAttData(0) = sclAtt8X
        CurAttData(1) = sclAtt8Y
        CurAttData(2) = sclAtt8Hall
        CurAttData(3) = 0
    Case 9    'Damage
        CurAttData(0) = sclAtt9Damage
        CurAttData(1) = 0
        CurAttData(2) = 0
        CurAttData(3) = 0
    Case 17    'Directional Wall
        CurAttData(0) = 0
        For A = 0 To 7
            If chkDirections(A).Value = 1 Then SetBit CurAttData(0), CByte(A)
        Next A
        CurAttData(1) = chkKeep.Value
        CurAttData(2) = chkNoMonster.Value
        CurAttData(3) = 0
    Case 19    'Light Tile
        CurAttData(0) = sclRadius
        CurAttData(1) = sclIntensity
        If chkWall = 1 Then
            SetBit CurAttData(2), 0
        Else
            ClearBit CurAttData(2), 0
        End If
        ClearBit CurAttData(2), 1
        ClearBit CurAttData(2), 2
        ClearBit CurAttData(2), 3
        ClearBit CurAttData(2), 4
        ClearBit CurAttData(2), 5
        ClearBit CurAttData(2), 6
        ClearBit CurAttData(2), 7
    Case 20    'Dampening Tile
        If chkAtt20Direction(0).Value = 1 Then
            SetBit CurAttData(0), 0
        Else
            ClearBit CurAttData(0), 0
        End If
        If chkAtt20Direction(1).Value = 1 Then
            SetBit CurAttData(0), 1
        Else
            ClearBit CurAttData(0), 1
        End If
        If chkAtt20Direction(2).Value = 1 Then
            SetBit CurAttData(0), 2
        Else
            ClearBit CurAttData(0), 2
        End If
        If chkAtt20Direction(3).Value = 1 Then
            SetBit CurAttData(0), 3
        Else
            ClearBit CurAttData(0), 3
        End If
        ClearBit CurAttData(0), 4
        ClearBit CurAttData(0), 5
        ClearBit CurAttData(0), 6
        ClearBit CurAttData(0), 7
        CurAttData(1) = sclAtt20X.Value
        CurAttData(2) = sclAtt20Y.Value
        If chkAtt20Wall = 1 Then
            SetBit CurAttData(3), 0
        Else
            ClearBit CurAttData(3), 0
        End If
        If chkAtt20Outdoor = 1 Then
            SetBit CurAttData(3), 1
        Else
            ClearBit CurAttData(3), 1
        End If
        If chkAtt20Indoor = 1 Then
            SetBit CurAttData(3), 2
        Else
            ClearBit CurAttData(3), 2
        End If
        ClearBit CurAttData(3), 3
        ClearBit CurAttData(3), 4
        ClearBit CurAttData(3), 5
        ClearBit CurAttData(3), 6
        ClearBit CurAttData(3), 7
    End Select
    Unload Me
End Sub

Private Sub Form_Load()

    sclAtt3Key.max = MaxObjects
    sclAtt7Obj.max = MaxObjects
    sclAtt2Map.max = MaxMaps
            
    Dim A As Long
    Select Case NewAtt
    Case 2    'Warp
        lblAtt = "2 - Warp"
        picAtt2.Visible = True
        If CurAtt = 2 Then
            sclAtt2Map.Value = CurAttData(0) * 256 + CurAttData(1)
            sclAtt2X.Value = CurAttData(2)
            sclAtt2Y.Value = CurAttData(3)
        End If
    Case 3    'Key
        lblAtt = "3 - Key"
        picAtt3.Visible = True
        If CurAtt = 3 Then
            sclAtt3Key.Value = CurAttData(1) * 256 + CurAttData(0)
        End If
    Case 7    'Obj
        lblAtt = "7 - Object"
        picAtt7.Visible = True
        If CurAtt = 7 Then
            sclAtt7Obj.Value = CurAttData(1) * 256 + CurAttData(0)
            sclAtt7Val.Value = CurAttData(2) * 256 + CurAttData(3)
        End If
    Case 8    'Touch Plate
        lblAtt = "8 - Touch Plate"
        picAtt8.Visible = True
        If CurAtt = 8 Then
            sclAtt8X = CurAttData(0)
            sclAtt8Y = CurAttData(1)
            sclAtt8Hall = CurAttData(2)
        End If
    Case 9    'Damage
        lblAtt = "9 - Damage"
        picAtt9.Visible = True
        If CurAtt = 9 Then
            sclAtt9Damage = CurAttData(0)
        End If
    Case 17    'Directional Wall
        lblAtt = "17 - Directional Wall"
        picAtt17.Visible = True
        If CurAtt = 17 Then
            For A = 0 To 7
                If ExamineBit(CurAttData(0), CByte(A)) = True Then chkDirections(A).Value = 1
            Next A
            chkKeep.Value = CurAttData(1)
            chkNoMonster.Value = CurAttData(2)
        End If
    Case 19    'Light
        lblAtt = "19 - Light"
        picAtt19.Visible = True
        If CurAtt = 19 Then
            sclRadius = CurAttData(0)
            sclIntensity = CurAttData(1)
            If ExamineBit(CurAttData(2), 0) = True Then chkWall.Value = 1
        End If
    Case 20    'Light Dampening
        lblAtt = "20 - Light Dampening"
        picAtt20.Visible = True
        If CurAtt = 20 Then
            If ExamineBit(CurAttData(0), 0) = True Then chkAtt20Direction(0).Value = 1
            If ExamineBit(CurAttData(0), 1) = True Then chkAtt20Direction(1).Value = 1
            If ExamineBit(CurAttData(0), 2) = True Then chkAtt20Direction(2).Value = 1
            If ExamineBit(CurAttData(0), 3) = True Then chkAtt20Direction(3).Value = 1
            sclAtt20X = CurAttData(1)
            sclAtt20Y = CurAttData(2)
            If ExamineBit(CurAttData(3), 0) = True Then chkAtt20Wall.Value = 1
            If ExamineBit(CurAttData(3), 1) = True Then chkAtt20Outdoor.Value = 1
            If ExamineBit(CurAttData(3), 2) = True Then chkAtt20Indoor.Value = 1
        End If
    End Select
End Sub

Private Sub lblAtt3Key_Change()
    If Val(lblAtt3Key) > 0 And Val(lblAtt3Key) <= MaxObjects Then sclAtt3Key = Val(lblAtt3Key)
End Sub

Private Sub lblAtt7Obj_Change()
    If Val(lblAtt7Obj) > 0 And Val(lblAtt7Obj) <= MaxObjects Then sclAtt7Obj = Val(lblAtt7Obj)
End Sub

Private Sub lblAtt7Val_Change()
    If Val(lblAtt7Val) <= 32000 Then sclAtt7Val = Val(lblAtt7Val)
End Sub

Private Sub sclAtt20X_Change()
    lblAtt20X = sclAtt20X
End Sub

Private Sub sclAtt20Y_Change()
    lblAtt20Y = sclAtt20Y
End Sub

Private Sub sclAtt2Map_Change()
    lblAtt2Map = sclAtt2Map
End Sub

Private Sub sclAtt2Map_Scroll()
    sclAtt2Map_Change
End Sub

Private Sub sclAtt2X_Change()
    lblAtt2X = sclAtt2X
End Sub

Private Sub sclAtt2X_Scroll()
    sclAtt2X_Change
End Sub

Private Sub sclAtt2Y_Change()
    lblAtt2Y = sclAtt2Y
End Sub

Private Sub sclAtt2Y_Scroll()
    sclAtt2Y_Change
End Sub

Private Sub sclAtt3Key_Change()
    lblAtt3Key = sclAtt3Key
    lblObject.Caption = Object(sclAtt3Key).name
End Sub

Private Sub sclAtt3Key_Scroll()
    sclAtt3Key_Change
End Sub

Private Sub sclAtt7Obj_Change()
    lblAtt7Obj = sclAtt7Obj
    lblName.Caption = Object(sclAtt7Obj).name
End Sub

Private Sub sclAtt7Obj_Scroll()
    sclAtt7Obj_Change
End Sub

Private Sub sclAtt7Val_Change()
    lblAtt7Val = sclAtt7Val
End Sub

Private Sub sclAtt7Val_Scroll()
    sclAtt7Val_Change
End Sub

Private Sub sclAtt8Hall_Change()
    lblAtt8Hall = sclAtt8Hall
End Sub

Private Sub sclAtt8Hall_Scroll()
    sclAtt8Hall_Change
End Sub

Private Sub sclAtt8X_Change()
    lblAtt8X = sclAtt8X
End Sub

Private Sub sclAtt8X_Scroll()
    sclAtt8X_Change
End Sub

Private Sub sclAtt8Y_Change()
    lblAtt8Y = sclAtt8Y
End Sub

Private Sub sclAtt8Y_Scroll()
    sclAtt8Y_Change
End Sub

Private Sub sclAtt9Damage_Change()
    lblAtt9Damage = sclAtt9Damage
End Sub

Private Sub sclAtt9Damage_Scroll()
    sclAtt9Damage_Change
End Sub

Private Sub sclIntensity_Change()
    lblIntensity = Str(sclIntensity)
End Sub

Private Sub sclIntensity_Scroll()
    sclIntensity_Change
End Sub

Private Sub sclRadius_Change()
    lblRadius = Str(sclRadius)
End Sub

Private Sub sclRadius_Scroll()
    sclRadius_Change
End Sub
