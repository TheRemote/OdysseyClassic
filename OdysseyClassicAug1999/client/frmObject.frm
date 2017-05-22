VERSION 5.00
Begin VB.Form frmObject 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "The Odyssey Online Classic [Editing Object]"
   ClientHeight    =   4080
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5865
   ControlBox      =   0   'False
   Icon            =   "frmObject.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4080
   ScaleWidth      =   5865
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cmbData 
      Height          =   315
      Index           =   3
      Left            =   1320
      Style           =   2  'Dropdown List
      TabIndex        =   22
      Top             =   3000
      Width           =   3495
   End
   Begin VB.CheckBox chkData 
      Height          =   315
      Index           =   3
      Left            =   1320
      TabIndex        =   21
      Top             =   3000
      Width           =   3495
   End
   Begin VB.HScrollBar sclData 
      Height          =   255
      Index           =   3
      Left            =   1320
      Max             =   255
      TabIndex        =   20
      Top             =   3000
      Width           =   3495
   End
   Begin VB.ComboBox cmbData 
      Height          =   315
      Index           =   0
      Left            =   1320
      Style           =   2  'Dropdown List
      TabIndex        =   19
      Top             =   1920
      Width           =   3495
   End
   Begin VB.CheckBox chkData 
      Height          =   315
      Index           =   0
      Left            =   1320
      TabIndex        =   18
      Top             =   1920
      Width           =   3495
   End
   Begin VB.ComboBox cmbData 
      Height          =   315
      Index           =   2
      Left            =   1320
      Style           =   2  'Dropdown List
      TabIndex        =   17
      Top             =   2640
      Width           =   3495
   End
   Begin VB.ComboBox cmbData 
      Height          =   315
      Index           =   1
      Left            =   1320
      Style           =   2  'Dropdown List
      TabIndex        =   16
      Top             =   2280
      Width           =   3495
   End
   Begin VB.CheckBox chkData 
      Height          =   315
      Index           =   2
      Left            =   1320
      TabIndex        =   15
      Top             =   2640
      Width           =   3495
   End
   Begin VB.CheckBox chkData 
      Height          =   315
      Index           =   1
      Left            =   1320
      TabIndex        =   14
      Top             =   2280
      Width           =   3495
   End
   Begin VB.HScrollBar sclData 
      Height          =   255
      Index           =   2
      Left            =   1320
      Max             =   255
      TabIndex        =   13
      Top             =   2640
      Width           =   3495
   End
   Begin VB.HScrollBar sclData 
      Height          =   255
      Index           =   1
      Left            =   1320
      Max             =   255
      TabIndex        =   12
      Top             =   2280
      Width           =   3495
   End
   Begin VB.HScrollBar sclData 
      Height          =   255
      Index           =   0
      Left            =   1320
      Max             =   255
      TabIndex        =   11
      Top             =   1920
      Width           =   3495
   End
   Begin VB.PictureBox picPicture 
      AutoRedraw      =   -1  'True
      Height          =   540
      Left            =   5280
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   8
      Top             =   600
      Width           =   540
   End
   Begin VB.HScrollBar sclPicture 
      Height          =   255
      Left            =   1320
      Max             =   255
      Min             =   1
      TabIndex        =   1
      Top             =   1200
      Value           =   1
      Width           =   3495
   End
   Begin VB.ComboBox cmbType 
      Height          =   315
      Left            =   1320
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1560
      Width           =   3495
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
      Top             =   600
      Width           =   3855
   End
   Begin VB.CommandButton btnOk 
      Caption         =   "Ok"
      Height          =   495
      Left            =   4320
      TabIndex        =   4
      Top             =   3480
      Width           =   1455
   End
   Begin VB.CommandButton btnCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   495
      Left            =   2760
      TabIndex        =   3
      Top             =   3480
      Width           =   1455
   End
   Begin VB.Label ObjData 
      Height          =   255
      Index           =   3
      Left            =   0
      TabIndex        =   34
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label ObjData 
      Height          =   255
      Index           =   2
      Left            =   0
      TabIndex        =   33
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label ObjData 
      Height          =   255
      Index           =   1
      Left            =   0
      TabIndex        =   32
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label ObjData 
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   31
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblData 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   255
      Index           =   3
      Left            =   4920
      TabIndex        =   30
      Top             =   3000
      Width           =   855
   End
   Begin VB.Label lblCaption 
      Caption         =   "Data4:"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   29
      Top             =   3000
      Width           =   1095
   End
   Begin VB.Label lblData 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   255
      Index           =   2
      Left            =   4920
      TabIndex        =   28
      Top             =   2640
      Width           =   855
   End
   Begin VB.Label lblCaption 
      Caption         =   "Data3:"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   27
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Label lblData 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   255
      Index           =   1
      Left            =   4920
      TabIndex        =   26
      Top             =   2280
      Width           =   855
   End
   Begin VB.Label lblCaption 
      Caption         =   "Data2:"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   25
      Top             =   2280
      Width           =   1095
   End
   Begin VB.Label lblData 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   255
      Index           =   0
      Left            =   4920
      TabIndex        =   24
      Top             =   1920
      Width           =   855
   End
   Begin VB.Label lblCaption 
      Caption         =   "Data1:"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   23
      Top             =   1920
      Width           =   1095
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
      Left            =   1320
      TabIndex        =   10
      Top             =   120
      Width           =   4455
   End
   Begin VB.Label Label4 
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
      TabIndex        =   9
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "Picture:"
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
      TabIndex        =   7
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Type:"
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
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label Label1 
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
      TabIndex        =   5
      Top             =   600
      Width           =   1095
   End
End
Attribute VB_Name = "frmObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ChangingTypes As Boolean
Private Sub btnCancel_Click()
    Me.Hide
End Sub


Private Sub btnOk_Click()
    SendSocket Chr$(21) + Chr$(lblNumber) + Chr$(sclPicture) + Chr$(cmbType.ListIndex) + Chr$(0) + Chr$(Val(ObjData(0))) + Chr$(Val(ObjData(1))) + Chr$(Val(ObjData(2))) + Chr$(Val(ObjData(3))) + txtName
    Me.Hide
End Sub


Private Sub chkData_Click(Index As Integer)
    lblData(Index) = chkData(Index)
End Sub

Private Sub cmbData_Click(Index As Integer)
    lblData(Index) = cmbData(Index).ListIndex
End Sub


Private Sub cmbType_Click()
    ChangingTypes = True
    
    Select Case cmbType.ListIndex
        Case 0 'None
            SetData 0, 0, ""
            SetData 1, 0, ""
            SetData 2, 0, ""
            SetData 3, 0, ""
            
        Case 1 'Weapon
            SetData 0, 1, "HP / 10"
            SetData 1, 1, "Strength"
            SetData 2, 0, ""
            SetData 3, 2, "Newbified"
            
        Case 2 'Shield
            SetData 0, 1, "HP / 10"
            SetData 1, 1, "Strength"
            SetData 2, 0, ""
            SetData 3, 2, "Newbified"
        
        Case 3 'Armor
            SetData 0, 1, "HP / 10"
            SetData 1, 1, "Strength"
            SetData 2, 0, ""
            SetData 3, 2, "Newbified"
        
        Case 4 'Helmut
            SetData 0, 1, "HP / 10"
            SetData 1, 1, "Strength"
            SetData 2, 0, ""
            SetData 3, 2, "Newbified"
            
        Case 5 'Potion
            With cmbData(0)
                .Clear
                .AddItem "Gives HP"
                .AddItem "Takes HP"
                .AddItem "Give Mana"
                .AddItem "Takes Mana"
                .AddItem "Gives Energy"
                .AddItem "Takes Energy"
            End With
            SetData 0, 3, "Potion Type"
            SetData 1, 1, "Value"
            SetData 2, 0, ""
            SetData 3, 2, "Newbified"
        
        Case 6 'Money
            SetData 0, 0, ""
            SetData 1, 0, ""
            SetData 2, 0, ""
            SetData 3, 0, ""
            
        Case 7 'Key
            SetData 0, 2, "Unlimited Use"
            SetData 1, 0, ""
            SetData 2, 0, ""
            SetData 3, 2, "Newbified"

        Case 8 'Ring
            With cmbData(0)
                .Clear
                .AddItem "Modifies Attack"
                .AddItem "Modifies Defense"
            End With
            SetData 0, 3, "Ring Type"
            SetData 1, 1, "Durability / 10"
            SetData 2, 1, "Modifier"
            SetData 3, 2, "Newbified"
        
        Case 9 'Guild
            SetData 0, 0, ""
            SetData 1, 0, ""
            SetData 2, 0, ""
            SetData 3, 0, ""
        End Select
        
    ChangingTypes = False
End Sub
Private Sub Form_Load()
    BitBlt picPicture.hdc, 0, 0, 32, 32, hdcObjects, 0, (sclPicture - 1) * 32, SRCCOPY
    picPicture.Refresh
    cmbType.AddItem "<None>"
    cmbType.AddItem "Weapon"
    cmbType.AddItem "Shield"
    cmbType.AddItem "Armor"
    cmbType.AddItem "Helmut"
    cmbType.AddItem "Potion"
    cmbType.AddItem "Money"
    cmbType.AddItem "Key"
    cmbType.AddItem "Ring"
    cmbType.AddItem "Guild Deed"
    
    frmObject_Loaded = True
End Sub
Sub SetData(Index As Integer, DataType As Integer, St As String)
    Select Case DataType
        Case 0
            lblData(Index).Visible = False
            sclData(Index).Visible = False
            chkData(Index).Visible = False
            cmbData(Index).Visible = False
        Case 1
            lblData(Index).Visible = True
            sclData(Index).Visible = True
            chkData(Index).Visible = False
            cmbData(Index).Visible = False
            sclData(Index) = Val(ObjData(Index))
            lblData(Index) = ObjData(Index)
        Case 2
            lblData(Index).Visible = True
            sclData(Index).Visible = False
            chkData(Index).Visible = True
            cmbData(Index).Visible = False
            If Val(ObjData(Index)) = 0 Then
                chkData(Index) = 0
                lblData(Index) = 0
            Else
                chkData(Index) = 1
                lblData(Index) = 1
            End If
        Case 3
            lblData(Index).Visible = True
            sclData(Index).Visible = False
            chkData(Index).Visible = False
            cmbData(Index).Visible = True
            If Val(ObjData(Index)) < cmbData(Index).ListCount Then
                cmbData(Index).ListIndex = Val(ObjData(Index))
                lblData(Index) = ObjData(Index)
            Else
                cmbData(Index).ListIndex = cmbData(Index).ListCount - 1
                lblData(Index) = cmbData(Index).ListCount - 1
            End If
    End Select
    lblCaption(Index) = St
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmObject_Loaded = False
End Sub

Private Sub lblData_Change(Index As Integer)
    If ChangingTypes = False Then
        ObjData(Index) = lblData(Index)
    End If
End Sub

Private Sub sclData_Change(Index As Integer)
    lblData(Index) = sclData(Index)
End Sub

Private Sub sclData_Scroll(Index As Integer)
    sclData_Change (Index)
End Sub


Private Sub sclPicture_Change()
    BitBlt picPicture.hdc, 0, 0, 32, 32, hdcObjects, 0, (sclPicture - 1) * 32, SRCCOPY
    picPicture.Refresh
End Sub


Private Sub sclPicture_Scroll()
    sclPicture_Change
End Sub


