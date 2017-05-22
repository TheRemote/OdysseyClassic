VERSION 5.00
Begin VB.Form frmMagic 
   Caption         =   "Edit Magic"
   ClientHeight    =   6345
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4950
   ControlBox      =   0   'False
   Icon            =   "frmMagic.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6345
   ScaleWidth      =   4950
   StartUpPosition =   3  'Windows Default
   Begin VB.HScrollBar sclCastTimer 
      Height          =   255
      Left            =   1320
      Max             =   30000
      TabIndex        =   24
      Top             =   5040
      Value           =   1
      Width           =   2895
   End
   Begin VB.OptionButton optIconType 
      Caption         =   "Object"
      Height          =   255
      Index           =   1
      Left            =   2160
      TabIndex        =   21
      Top             =   4560
      Width           =   855
   End
   Begin VB.PictureBox picSprite 
      AutoRedraw      =   -1  'True
      Height          =   540
      Left            =   120
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   20
      Top             =   4200
      Width           =   540
   End
   Begin VB.HScrollBar sclFrame 
      Height          =   255
      Left            =   1320
      Max             =   9
      TabIndex        =   19
      Top             =   4200
      Width           =   2895
   End
   Begin VB.OptionButton optIconType 
      Caption         =   "Effect"
      Height          =   255
      Index           =   0
      Left            =   1320
      TabIndex        =   18
      Top             =   4560
      Value           =   -1  'True
      Width           =   855
   End
   Begin VB.HScrollBar sclIcon 
      Height          =   255
      Left            =   1320
      Max             =   255
      Min             =   1
      TabIndex        =   17
      Top             =   3840
      Value           =   1
      Width           =   2895
   End
   Begin VB.CheckBox chkClass 
      Caption         =   "Class4"
      Height          =   255
      Index           =   3
      Left            =   2760
      TabIndex        =   15
      Top             =   3360
      Width           =   1335
   End
   Begin VB.CheckBox chkClass 
      Caption         =   "Class3"
      Height          =   255
      Index           =   2
      Left            =   1320
      TabIndex        =   14
      Top             =   3360
      Width           =   1335
   End
   Begin VB.CheckBox chkClass 
      Caption         =   "Class2"
      Height          =   255
      Index           =   1
      Left            =   2760
      TabIndex        =   13
      Top             =   3000
      Width           =   1335
   End
   Begin VB.CheckBox chkClass 
      Caption         =   "Class1"
      Height          =   255
      Index           =   0
      Left            =   1320
      TabIndex        =   12
      Top             =   3000
      Width           =   1335
   End
   Begin VB.TextBox txtDescription 
      Height          =   1335
      Left            =   1320
      MaxLength       =   100
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   1080
      Width           =   2895
   End
   Begin VB.CommandButton btnOk 
      Caption         =   "Ok"
      Height          =   495
      Left            =   2520
      TabIndex        =   4
      Top             =   5640
      Width           =   1455
   End
   Begin VB.CommandButton btnCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   495
      Left            =   960
      TabIndex        =   2
      Top             =   5640
      Width           =   1455
   End
   Begin VB.TextBox txtName 
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
      Left            =   1320
      MaxLength       =   15
      TabIndex        =   0
      Text            =   " "
      Top             =   600
      Width           =   2895
   End
   Begin VB.HScrollBar sclLevel 
      Height          =   255
      Left            =   1320
      Max             =   255
      Min             =   1
      TabIndex        =   3
      Top             =   2520
      Value           =   1
      Width           =   2895
   End
   Begin VB.Label Label1 
      Caption         =   "Cast Timer:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   26
      Top             =   4920
      Width           =   855
   End
   Begin VB.Label lblCastTimer 
      Alignment       =   2  'Center
      Caption         =   "0"
      Height          =   255
      Left            =   4320
      TabIndex        =   25
      Top             =   5040
      Width           =   495
   End
   Begin VB.Label lblFrame 
      Alignment       =   2  'Center
      Caption         =   "0"
      Height          =   255
      Left            =   4320
      TabIndex        =   23
      Top             =   4200
      Width           =   495
   End
   Begin VB.Label lblIcon 
      Alignment       =   2  'Center
      Caption         =   "1"
      Height          =   255
      Left            =   4320
      TabIndex        =   22
      Top             =   3840
      Width           =   495
   End
   Begin VB.Label cptIcon 
      Caption         =   "Icon:"
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
      TabIndex        =   16
      Top             =   3840
      Width           =   855
   End
   Begin VB.Label cptScript 
      Caption         =   "Description:"
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
      Index           =   1
      Left            =   120
      TabIndex        =   11
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label cptName 
      Caption         =   "Name:"
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
      TabIndex        =   10
      Top             =   600
      Width           =   855
   End
   Begin VB.Label cptLevel 
      Caption         =   "Level:"
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
      TabIndex        =   9
      Top             =   2520
      Width           =   855
   End
   Begin VB.Label cptNumber 
      Caption         =   "Number:"
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
      TabIndex        =   8
      Top             =   120
      Width           =   855
   End
   Begin VB.Label lblNumber 
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
      Left            =   1560
      TabIndex        =   7
      Top             =   120
      Width           =   2655
   End
   Begin VB.Label cptClass 
      Caption         =   "Class:"
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
      Top             =   3000
      Width           =   855
   End
   Begin VB.Label lblLevel 
      Alignment       =   2  'Center
      Caption         =   "1"
      Height          =   255
      Left            =   4320
      TabIndex        =   5
      Top             =   2520
      Width           =   495
   End
End
Attribute VB_Name = "frmMagic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub btnOk_Click()
    Dim A As Long, Class As Byte, IconType As Byte

    For A = 0 To NumClasses - 1
        If chkClass(A).value = 1 Then
            SetBit Class, CByte(A)
        Else
            ClearBit Class, CByte(A)
        End If
    Next A

    For A = 0 To 1
        If optIconType(A).value = True Then
            IconType = A
        End If
    Next A

    SendSocket Chr$(83) + DoubleChar$(lblNumber) + Chr$(lblLevel) + Chr$(Class) + DoubleChar$(lblIcon) + Chr$(IconType) + DoubleChar$(lblCastTimer) + txtName + Chr$(0) + txtDescription
    Unload Me
End Sub

Private Sub Form_Load()
    frmMagic_Loaded = True

    Dim A As Long
    For A = 0 To NumClasses - 1
        chkClass(A).Caption = Class(A + 1).name
    Next A

    RefreshIcon
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmMagic_Loaded = False
End Sub

Private Sub optIconType_Click(index As Integer)
    RefreshIcon
End Sub

Private Sub sclCastTimer_Change()
    lblCastTimer = sclCastTimer.value
End Sub

Private Sub sclFrame_Change()
    lblFrame = sclFrame.value
    RefreshIcon
End Sub

Private Sub sclIcon_Change()
    lblIcon = sclIcon.value
    RefreshIcon
End Sub

Private Sub sclLevel_Change()
    lblLevel = sclLevel
End Sub

Sub RefreshIcon()
    Dim A As Long, IconType As Byte

    For A = 0 To 1
        If optIconType(A).value = True Then
            IconType = A
        End If
    Next A

    Select Case IconType
    Case 0:    'Effect
        DrawToDC 0, 0, 32, 32, picSprite.hDC, DDSEffects, sclFrame * 32, (sclIcon - 1) * 32
        picSprite.Refresh
    Case 1:    'Object
        DrawToDC 0, 0, 32, 32, picSprite.hDC, DDSObjects, 0, (sclIcon - 1) * 32
        picSprite.Refresh
    End Select
End Sub
