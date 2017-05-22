VERSION 5.00
Begin VB.Form frmOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Odyssey Server Options"
   ClientHeight    =   8580
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7410
   ControlBox      =   0   'False
   Icon            =   "frmOptions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8580
   ScaleWidth      =   7410
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtVal 
      Height          =   375
      Index           =   8
      Left            =   2760
      MaxLength       =   5
      TabIndex        =   47
      Top             =   7680
      Width           =   975
   End
   Begin VB.TextBox txtVal 
      Height          =   375
      Index           =   7
      Left            =   2760
      MaxLength       =   5
      TabIndex        =   46
      Top             =   7320
      Width           =   975
   End
   Begin VB.TextBox txtVal 
      Height          =   375
      Index           =   6
      Left            =   2760
      MaxLength       =   5
      TabIndex        =   45
      Top             =   6960
      Width           =   975
   End
   Begin VB.TextBox txtVal 
      Height          =   375
      Index           =   5
      Left            =   2760
      MaxLength       =   5
      TabIndex        =   44
      Top             =   6600
      Width           =   975
   End
   Begin VB.TextBox txtVal 
      Height          =   375
      Index           =   4
      Left            =   2760
      MaxLength       =   5
      TabIndex        =   43
      Top             =   6240
      Width           =   975
   End
   Begin VB.TextBox txtVal 
      Height          =   375
      Index           =   3
      Left            =   2760
      MaxLength       =   5
      TabIndex        =   42
      Top             =   5880
      Width           =   975
   End
   Begin VB.TextBox txtVal 
      Height          =   375
      Index           =   2
      Left            =   2760
      MaxLength       =   5
      TabIndex        =   41
      Top             =   5520
      Width           =   975
   End
   Begin VB.TextBox txtVal 
      Height          =   375
      Index           =   1
      Left            =   2760
      MaxLength       =   5
      TabIndex        =   40
      Top             =   5160
      Width           =   975
   End
   Begin VB.TextBox txtObj 
      Height          =   375
      Index           =   8
      Left            =   1680
      MaxLength       =   3
      TabIndex        =   39
      Top             =   7680
      Width           =   735
   End
   Begin VB.TextBox txtObj 
      Height          =   375
      Index           =   7
      Left            =   1680
      MaxLength       =   3
      TabIndex        =   38
      Top             =   7320
      Width           =   735
   End
   Begin VB.TextBox txtObj 
      Height          =   375
      Index           =   6
      Left            =   1680
      MaxLength       =   3
      TabIndex        =   37
      Top             =   6960
      Width           =   735
   End
   Begin VB.TextBox txtObj 
      Height          =   375
      Index           =   5
      Left            =   1680
      MaxLength       =   3
      TabIndex        =   36
      Top             =   6600
      Width           =   735
   End
   Begin VB.TextBox txtObj 
      Height          =   375
      Index           =   4
      Left            =   1680
      MaxLength       =   3
      TabIndex        =   35
      Top             =   6240
      Width           =   735
   End
   Begin VB.TextBox txtObj 
      Height          =   375
      Index           =   3
      Left            =   1680
      MaxLength       =   3
      TabIndex        =   34
      Top             =   5880
      Width           =   735
   End
   Begin VB.TextBox txtObj 
      Height          =   375
      Index           =   2
      Left            =   1680
      MaxLength       =   3
      TabIndex        =   33
      Top             =   5520
      Width           =   735
   End
   Begin VB.TextBox txtObj 
      Height          =   375
      Index           =   1
      Left            =   1680
      MaxLength       =   3
      TabIndex        =   32
      Top             =   5160
      Width           =   735
   End
   Begin VB.TextBox txtText 
      Height          =   375
      Index           =   4
      Left            =   3960
      MaxLength       =   255
      TabIndex        =   20
      Top             =   4320
      Width           =   3375
   End
   Begin VB.TextBox txtY 
      Height          =   375
      Index           =   4
      Left            =   3120
      MaxLength       =   2
      TabIndex        =   19
      Top             =   4320
      Width           =   735
   End
   Begin VB.TextBox txtX 
      Height          =   375
      Index           =   4
      Left            =   2280
      MaxLength       =   2
      TabIndex        =   18
      Top             =   4320
      Width           =   735
   End
   Begin VB.TextBox txtMap 
      Height          =   375
      Index           =   4
      Left            =   1440
      MaxLength       =   4
      TabIndex        =   17
      Top             =   4320
      Width           =   735
   End
   Begin VB.TextBox txtText 
      Height          =   375
      Index           =   3
      Left            =   3960
      MaxLength       =   255
      TabIndex        =   16
      Top             =   3840
      Width           =   3375
   End
   Begin VB.TextBox txtY 
      Height          =   375
      Index           =   3
      Left            =   3120
      MaxLength       =   2
      TabIndex        =   15
      Top             =   3840
      Width           =   735
   End
   Begin VB.TextBox txtX 
      Height          =   375
      Index           =   3
      Left            =   2280
      MaxLength       =   2
      TabIndex        =   14
      Top             =   3840
      Width           =   735
   End
   Begin VB.TextBox txtMap 
      Height          =   375
      Index           =   3
      Left            =   1440
      MaxLength       =   4
      TabIndex        =   13
      Top             =   3840
      Width           =   735
   End
   Begin VB.TextBox txtText 
      Height          =   375
      Index           =   2
      Left            =   3960
      MaxLength       =   255
      TabIndex        =   12
      Top             =   3360
      Width           =   3375
   End
   Begin VB.TextBox txtY 
      Height          =   375
      Index           =   2
      Left            =   3120
      MaxLength       =   2
      TabIndex        =   11
      Top             =   3360
      Width           =   735
   End
   Begin VB.TextBox txtX 
      Height          =   375
      Index           =   2
      Left            =   2280
      MaxLength       =   2
      TabIndex        =   10
      Top             =   3360
      Width           =   735
   End
   Begin VB.TextBox txtMap 
      Height          =   375
      Index           =   2
      Left            =   1440
      MaxLength       =   4
      TabIndex        =   9
      Top             =   3360
      Width           =   735
   End
   Begin VB.TextBox txtText 
      Height          =   375
      Index           =   1
      Left            =   3960
      MaxLength       =   255
      TabIndex        =   8
      Top             =   2880
      Width           =   3375
   End
   Begin VB.TextBox txtY 
      Height          =   375
      Index           =   1
      Left            =   3120
      MaxLength       =   2
      TabIndex        =   7
      Top             =   2880
      Width           =   735
   End
   Begin VB.TextBox txtX 
      Height          =   375
      Index           =   1
      Left            =   2280
      MaxLength       =   2
      TabIndex        =   6
      Top             =   2880
      Width           =   735
   End
   Begin VB.TextBox txtMap 
      Height          =   375
      Index           =   1
      Left            =   1440
      MaxLength       =   4
      TabIndex        =   5
      Top             =   2880
      Width           =   735
   End
   Begin VB.TextBox txtText 
      Height          =   375
      Index           =   0
      Left            =   3960
      MaxLength       =   255
      TabIndex        =   4
      Top             =   2400
      Width           =   3375
   End
   Begin VB.TextBox txtY 
      Height          =   375
      Index           =   0
      Left            =   3120
      MaxLength       =   2
      TabIndex        =   3
      Top             =   2400
      Width           =   735
   End
   Begin VB.TextBox txtX 
      Height          =   375
      Index           =   0
      Left            =   2280
      MaxLength       =   2
      TabIndex        =   2
      Top             =   2400
      Width           =   735
   End
   Begin VB.TextBox txtMap 
      Height          =   375
      Index           =   0
      Left            =   1440
      MaxLength       =   4
      TabIndex        =   1
      Top             =   2400
      Width           =   735
   End
   Begin VB.TextBox txtMOTD 
      Height          =   1455
      Left            =   1440
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   240
      Width           =   5895
   End
   Begin VB.CommandButton btnCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   495
      Left            =   6120
      TabIndex        =   21
      Top             =   7920
      Width           =   1215
   End
   Begin VB.CommandButton btnOk 
      Caption         =   "Ok"
      Height          =   495
      Left            =   4800
      TabIndex        =   22
      Top             =   7920
      Width           =   1215
   End
   Begin VB.Label Label9 
      Caption         =   "Ammount"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2760
      TabIndex        =   31
      Top             =   4800
      Width           =   735
   End
   Begin VB.Label Label8 
      Caption         =   "Object"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1800
      TabIndex        =   30
      Top             =   4800
      Width           =   615
   End
   Begin VB.Label Label7 
      Caption         =   "Starting Objects:"
      Height          =   255
      Left            =   120
      TabIndex        =   29
      Top             =   4800
      Width           =   1215
   End
   Begin VB.Label Label6 
      Caption         =   "Text"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3960
      TabIndex        =   28
      Top             =   1920
      Width           =   735
   End
   Begin VB.Label Label5 
      Caption         =   "Y"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3120
      TabIndex        =   27
      Top             =   1920
      Width           =   735
   End
   Begin VB.Label Label4 
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      TabIndex        =   26
      Top             =   1920
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "Map"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      TabIndex        =   25
      Top             =   1920
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "Start Locations:"
      Height          =   375
      Left            =   120
      TabIndex        =   24
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "MOTD:"
      Height          =   375
      Left            =   120
      TabIndex        =   23
      Top             =   240
      Width           =   855
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub btnOk_Click()
    Dim B As Long
    
    DataRS.Edit
    
    World.MOTD = txtMOTD
    DataRS!MOTD = txtMOTD
    
    For A = 0 To 4
        With World.StartLocation(A)
            B = Val(txtMap(A))
            If B < 1 Then B = 1
            If B > 2000 Then B = 2000
            .Map = B
            B = Val(txtX(A))
            If B < 0 Then B = 0
            If B > 11 Then B = 11
            .X = B
            B = Val(txtY(A))
            If B < 0 Then B = 0
            If B > 11 Then B = 11
            .Y = B
            .Message = txtText(A)
            
            DataRS.Fields("StartLocationX" + CStr(A)) = .X
            DataRS.Fields("StartLocationY" + CStr(A)) = .Y
            DataRS.Fields("StartLocationMap" + CStr(A)) = .Map
            DataRS.Fields("StartLocationMessage" + CStr(A)) = .Message
        End With
    Next A

            For A = 1 To 8
                If Val(txtObj(A).Text) < 0 Then txtObj(A).Text = 0
                    If Val(txtObj(A).Text) > 255 Then txtObj(A).Text = 255
                        DataRS.Fields("StartingObj" + CStr(A)) = Val(txtObj(A).Text)
                        World.StartObjects(A) = Val(txtObj(A).Text)
            Next A
            
            For A = 1 To 8
                If Val(txtVal(A).Text) < 0 Then txtVal(A).Text = 0
                    If Val(txtVal(A).Text) > 32000 Then txtVal(A).Text = 32000
                        DataRS.Fields("StartingObjVal" + CStr(A)) = Val(txtVal(A).Text)
                        World.StartObjValues(A) = Val(txtVal(A).Text)
            Next A
    DataRS.Update
    Unload Me
End Sub

Private Sub Form_Load()
    Dim A As Long
    
    txtMOTD = World.MOTD
    For A = 0 To 4
        With World.StartLocation(A)
            txtMap(A) = .Map
            txtX(A) = .X
            txtY(A) = .Y
            txtText(A) = .Message
        End With
    Next A
    
    For A = 1 To 8
        txtObj(A).Text = World.StartObjects(A)
        txtVal(A).Text = World.StartObjValues(A)
    Next A
End Sub

Private Sub txtMOTD_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Or KeyAscii = 10 Then
        KeyAscii = 0
        Beep
    End If
End Sub


