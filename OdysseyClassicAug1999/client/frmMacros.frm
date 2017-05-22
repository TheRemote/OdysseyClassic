VERSION 5.00
Begin VB.Form frmMacros 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "The Odyssey Online Classic [Macros]"
   ClientHeight    =   5580
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8700
   Icon            =   "frmMacros.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5580
   ScaleWidth      =   8700
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkLineFeed 
      Caption         =   "LineFeed"
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
      Index           =   9
      Left            =   7440
      TabIndex        =   19
      Top             =   4440
      Width           =   1215
   End
   Begin VB.TextBox txtMacro 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   9
      Left            =   720
      MaxLength       =   255
      TabIndex        =   18
      Top             =   4440
      Width           =   6615
   End
   Begin VB.CheckBox chkLineFeed 
      Caption         =   "LineFeed"
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
      Index           =   8
      Left            =   7440
      TabIndex        =   17
      Top             =   3960
      Width           =   1215
   End
   Begin VB.TextBox txtMacro 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   8
      Left            =   720
      MaxLength       =   255
      TabIndex        =   16
      Top             =   3960
      Width           =   6615
   End
   Begin VB.CheckBox chkLineFeed 
      Caption         =   "LineFeed"
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
      Index           =   7
      Left            =   7440
      TabIndex        =   15
      Top             =   3480
      Width           =   1215
   End
   Begin VB.TextBox txtMacro 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   7
      Left            =   720
      MaxLength       =   255
      TabIndex        =   14
      Top             =   3480
      Width           =   6615
   End
   Begin VB.CheckBox chkLineFeed 
      Caption         =   "LineFeed"
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
      Index           =   6
      Left            =   7440
      TabIndex        =   13
      Top             =   3000
      Width           =   1215
   End
   Begin VB.TextBox txtMacro 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   6
      Left            =   720
      MaxLength       =   255
      TabIndex        =   12
      Top             =   3000
      Width           =   6615
   End
   Begin VB.CheckBox chkLineFeed 
      Caption         =   "LineFeed"
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
      Index           =   5
      Left            =   7440
      TabIndex        =   11
      Top             =   2520
      Width           =   1215
   End
   Begin VB.TextBox txtMacro 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   5
      Left            =   720
      MaxLength       =   255
      TabIndex        =   10
      Top             =   2520
      Width           =   6615
   End
   Begin VB.CheckBox chkLineFeed 
      Caption         =   "LineFeed"
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
      Index           =   4
      Left            =   7440
      TabIndex        =   9
      Top             =   2040
      Width           =   1215
   End
   Begin VB.TextBox txtMacro 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   4
      Left            =   720
      MaxLength       =   255
      TabIndex        =   8
      Top             =   2040
      Width           =   6615
   End
   Begin VB.CheckBox chkLineFeed 
      Caption         =   "LineFeed"
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
      Index           =   3
      Left            =   7440
      TabIndex        =   7
      Top             =   1560
      Width           =   1215
   End
   Begin VB.TextBox txtMacro 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   3
      Left            =   720
      MaxLength       =   255
      TabIndex        =   6
      Top             =   1560
      Width           =   6615
   End
   Begin VB.CheckBox chkLineFeed 
      Caption         =   "LineFeed"
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
      Index           =   2
      Left            =   7440
      TabIndex        =   5
      Top             =   1080
      Width           =   1215
   End
   Begin VB.TextBox txtMacro 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   2
      Left            =   720
      MaxLength       =   255
      TabIndex        =   4
      Top             =   1080
      Width           =   6615
   End
   Begin VB.CheckBox chkLineFeed 
      Caption         =   "LineFeed"
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
      Index           =   1
      Left            =   7440
      TabIndex        =   3
      Top             =   600
      Width           =   1215
   End
   Begin VB.TextBox txtMacro 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   1
      Left            =   720
      MaxLength       =   255
      TabIndex        =   2
      Top             =   600
      Width           =   6615
   End
   Begin VB.CheckBox chkLineFeed 
      Caption         =   "LineFeed"
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
      Index           =   0
      Left            =   7440
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox txtMacro 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   0
      Left            =   720
      MaxLength       =   255
      TabIndex        =   0
      Top             =   120
      Width           =   6615
   End
   Begin VB.CommandButton btnCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   495
      Left            =   4320
      TabIndex        =   20
      Top             =   5040
      Width           =   1455
   End
   Begin VB.CommandButton btnOk 
      Caption         =   "Ok"
      Height          =   495
      Left            =   5880
      TabIndex        =   21
      Top             =   5040
      Width           =   1455
   End
   Begin VB.Label Label10 
      Caption         =   "F10:"
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
      Top             =   4440
      Width           =   495
   End
   Begin VB.Label Label9 
      Caption         =   "F9:"
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
      TabIndex        =   30
      Top             =   3960
      Width           =   495
   End
   Begin VB.Label Label8 
      Caption         =   "F8:"
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
      TabIndex        =   29
      Top             =   3480
      Width           =   495
   End
   Begin VB.Label Label7 
      Caption         =   "F7:"
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
      TabIndex        =   28
      Top             =   3000
      Width           =   495
   End
   Begin VB.Label Label6 
      Caption         =   "F6:"
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
      TabIndex        =   27
      Top             =   2520
      Width           =   495
   End
   Begin VB.Label Label5 
      Caption         =   "F5:"
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
      Top             =   2040
      Width           =   495
   End
   Begin VB.Label Label4 
      Caption         =   "F4:"
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
      TabIndex        =   25
      Top             =   1560
      Width           =   495
   End
   Begin VB.Label Label3 
      Caption         =   "F3:"
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
      TabIndex        =   24
      Top             =   1080
      Width           =   495
   End
   Begin VB.Label Label2 
      Caption         =   "F2:"
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
      Top             =   600
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "F1:"
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
      TabIndex        =   22
      Top             =   120
      Width           =   495
   End
End
Attribute VB_Name = "frmMacros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnCancel_Click()
    Unload Me
End Sub


Private Sub btnOk_Click()
    Dim A As Long
    For A = 0 To 9
        With Macro(A)
            .Text = txtMacro(A)
            .LineFeed = Choose(chkLineFeed(A) + 1, False, True)
        End With
        WriteString "Macros", "Text" + CStr(A + 1), txtMacro(A)
        WriteString "Macros", "LineFeed" + CStr(A + 1), CStr(chkLineFeed(A))
    Next A
    Unload Me
End Sub

Private Sub Form_Load()
    Dim A As Long
    For A = 0 To 9
        With Macro(A)
            txtMacro(A) = .Text
            If .LineFeed = True Then
                chkLineFeed(A) = 1
            Else
                chkLineFeed(A) = 0
            End If
        End With
    Next A
    frmMacros_Loaded = True
End Sub

Private Sub Form_LostFocus()
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmMacros_Loaded = False
End Sub


Private Sub txtMacro_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 8 Or (KeyAscii >= 32 And KeyAscii <= 127) Then
        'Valid Key
    Else
        KeyAscii = 0
        Beep
    End If
End Sub

