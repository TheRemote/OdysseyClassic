VERSION 5.00
Begin VB.Form frmObject 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "The Odyssey Online Classic [Editing Object]"
   ClientHeight    =   6015
   ClientLeft      =   105
   ClientTop       =   330
   ClientWidth     =   5865
   ControlBox      =   0   'False
   Icon            =   "frmObject.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6015
   ScaleWidth      =   5865
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtSellPrice 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   4920
      TabIndex        =   52
      Text            =   "0"
      Top             =   4920
      Width           =   855
   End
   Begin VB.CheckBox chkFlags 
      Caption         =   "Can't Trade"
      Height          =   255
      Index           =   6
      Left            =   3000
      TabIndex        =   51
      Top             =   3600
      Width           =   1215
   End
   Begin VB.HScrollBar sclSellPrice 
      Height          =   255
      Left            =   960
      Max             =   30000
      TabIndex        =   50
      Top             =   4920
      Value           =   1
      Width           =   3855
   End
   Begin VB.HScrollBar sclLevel 
      Height          =   255
      Left            =   1320
      Max             =   255
      TabIndex        =   47
      Top             =   3960
      Value           =   1
      Width           =   3495
   End
   Begin VB.CheckBox chkFlags 
      Caption         =   "Locked"
      Height          =   255
      Index           =   5
      Left            =   1560
      TabIndex        =   45
      Top             =   3600
      Width           =   975
   End
   Begin VB.CheckBox chkClass 
      Caption         =   "Cleric"
      Height          =   255
      Index           =   3
      Left            =   3840
      TabIndex        =   44
      Top             =   4320
      Width           =   855
   End
   Begin VB.CheckBox chkClass 
      Caption         =   "Rogue"
      Height          =   255
      Index           =   2
      Left            =   3000
      TabIndex        =   43
      Top             =   4320
      Width           =   855
   End
   Begin VB.CheckBox chkClass 
      Caption         =   "Knight"
      Height          =   255
      Index           =   0
      Left            =   1320
      TabIndex        =   42
      Top             =   4320
      Width           =   855
   End
   Begin VB.CheckBox chkClass 
      Caption         =   "Mage"
      Height          =   255
      Index           =   1
      Left            =   2160
      TabIndex        =   41
      Top             =   4320
      Width           =   735
   End
   Begin VB.CommandButton btnCancel 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   2760
      TabIndex        =   22
      Top             =   5400
      Width           =   1455
   End
   Begin VB.CommandButton btnOk 
      Caption         =   "Ok"
      Height          =   495
      Left            =   4320
      TabIndex        =   21
      Top             =   5400
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
      TabIndex        =   20
      Top             =   600
      Width           =   3855
   End
   Begin VB.ComboBox cmbType 
      Height          =   315
      ItemData        =   "frmObject.frx":1CFA
      Left            =   1320
      List            =   "frmObject.frx":1CFC
      Style           =   2  'Dropdown List
      TabIndex        =   19
      Top             =   1560
      Width           =   3495
   End
   Begin VB.HScrollBar sclPicture 
      Height          =   255
      Left            =   1320
      Max             =   977
      Min             =   1
      TabIndex        =   18
      Top             =   1200
      Value           =   1
      Width           =   3495
   End
   Begin VB.PictureBox picPicture 
      AutoRedraw      =   -1  'True
      Height          =   540
      Left            =   5280
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   17
      Top             =   600
      Width           =   540
   End
   Begin VB.HScrollBar sclData 
      Height          =   255
      Index           =   0
      Left            =   1320
      Max             =   255
      TabIndex        =   16
      Top             =   1920
      Width           =   3495
   End
   Begin VB.HScrollBar sclData 
      Height          =   255
      Index           =   1
      Left            =   1320
      Max             =   255
      TabIndex        =   15
      Top             =   2280
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
   Begin VB.CheckBox chkData 
      Height          =   315
      Index           =   2
      Left            =   1320
      TabIndex        =   13
      Top             =   2640
      Width           =   3495
   End
   Begin VB.ComboBox cmbData 
      Height          =   315
      Index           =   1
      Left            =   1320
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   2280
      Width           =   3495
   End
   Begin VB.ComboBox cmbData 
      Height          =   315
      Index           =   2
      Left            =   1320
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   2640
      Width           =   3495
   End
   Begin VB.CheckBox chkData 
      Height          =   315
      Index           =   0
      Left            =   1320
      TabIndex        =   10
      Top             =   1920
      Width           =   3495
   End
   Begin VB.ComboBox cmbData 
      Height          =   315
      Index           =   0
      Left            =   1320
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   1920
      Width           =   3495
   End
   Begin VB.HScrollBar sclData 
      Height          =   255
      Index           =   3
      Left            =   1320
      Max             =   255
      TabIndex        =   8
      Top             =   3000
      Width           =   3495
   End
   Begin VB.CheckBox chkData 
      Height          =   315
      Index           =   3
      Left            =   1320
      TabIndex        =   7
      Top             =   3000
      Width           =   3495
   End
   Begin VB.ComboBox cmbData 
      Height          =   315
      Index           =   3
      Left            =   1320
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   3000
      Width           =   3495
   End
   Begin VB.HScrollBar sclData 
      Height          =   255
      Index           =   2
      Left            =   1320
      Max             =   255
      TabIndex        =   5
      Top             =   2640
      Width           =   3495
   End
   Begin VB.CheckBox chkFlags 
      Caption         =   "Can't Repair"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   3360
      Width           =   1215
   End
   Begin VB.CheckBox chkFlags 
      Caption         =   "Indestructible"
      Height          =   255
      Index           =   1
      Left            =   1560
      TabIndex        =   3
      Top             =   3360
      Width           =   1335
   End
   Begin VB.CheckBox chkFlags 
      Caption         =   "Undroppable"
      Height          =   255
      Index           =   2
      Left            =   3000
      TabIndex        =   2
      Top             =   3360
      Width           =   1335
   End
   Begin VB.CheckBox chkFlags 
      Caption         =   "Two Handed"
      Height          =   255
      Index           =   3
      Left            =   4440
      TabIndex        =   1
      Top             =   3360
      Width           =   1335
   End
   Begin VB.CheckBox chkFlags 
      Caption         =   "Can't Bank"
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   0
      Top             =   3600
      Width           =   1215
   End
   Begin VB.Label Label6 
      Caption         =   "Sell Price:"
      Height          =   255
      Left            =   120
      TabIndex        =   49
      Top             =   4920
      Width           =   735
   End
   Begin VB.Label lblLevel 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "1"
      Height          =   255
      Left            =   4920
      TabIndex        =   48
      Top             =   3960
      Width           =   855
   End
   Begin VB.Label lblLevelReq 
      Caption         =   "Level:"
      Height          =   255
      Left            =   120
      TabIndex        =   46
      Top             =   3960
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
      TabIndex        =   40
      Top             =   600
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
      TabIndex        =   39
      Top             =   1560
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
      TabIndex        =   38
      Top             =   1200
      Width           =   1095
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
      TabIndex        =   37
      Top             =   120
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
      TabIndex        =   36
      Top             =   120
      Width           =   4455
   End
   Begin VB.Label lblCaption 
      Caption         =   "Data1:"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   35
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Label lblData 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   255
      Index           =   0
      Left            =   4920
      TabIndex        =   34
      Top             =   1920
      Width           =   855
   End
   Begin VB.Label lblCaption 
      Caption         =   "Data2:"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   33
      Top             =   2280
      Width           =   1095
   End
   Begin VB.Label lblData 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   255
      Index           =   1
      Left            =   4920
      TabIndex        =   32
      Top             =   2280
      Width           =   855
   End
   Begin VB.Label lblCaption 
      Caption         =   "Data3:"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   31
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Label lblData 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   255
      Index           =   2
      Left            =   4920
      TabIndex        =   30
      Top             =   2640
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
      Index           =   3
      Left            =   4920
      TabIndex        =   28
      Top             =   3000
      Width           =   855
   End
   Begin VB.Label ObjData 
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   27
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label ObjData 
      Height          =   255
      Index           =   1
      Left            =   0
      TabIndex        =   26
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label ObjData 
      Height          =   255
      Index           =   2
      Left            =   0
      TabIndex        =   25
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label ObjData 
      Height          =   255
      Index           =   3
      Left            =   0
      TabIndex        =   24
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Label5 
      Caption         =   "Cannot be used by:"
      Height          =   495
      Left            =   120
      TabIndex        =   23
      Top             =   4320
      Width           =   855
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
    Dim flags As Byte, A As Long, ClassReq As Byte

    For A = 0 To 6
        If chkFlags(A) = 1 Then
            SetBit flags, CByte(A)
        Else
            ClearBit flags, CByte(A)
        End If
    Next A

    For A = 0 To 3
        If chkClass(A) = 1 Then
            SetBit ClassReq, CByte(A)
        Else
            ClearBit ClassReq, CByte(A)
        End If
    Next A

    SendSocket Chr$(21) + DoubleChar$(lblNumber) + DoubleChar$(sclPicture) + Chr$(cmbType.ListIndex) + Chr$(flags) + Chr$(Val(ObjData(0))) + Chr$(Val(ObjData(1))) + Chr$(Val(ObjData(2))) + Chr$(Val(ObjData(3))) + Chr$(ClassReq) + Chr$(lblLevel) + DoubleChar$(sclSellPrice) + txtName
    Me.Hide
End Sub

Private Sub chkData_Click(index As Integer)
    lblData(index) = chkData(index)
End Sub

Private Sub cmbData_Click(index As Integer)
    lblData(index) = cmbData(index).ListIndex
End Sub


Private Sub cmbType_Click()
    ChangingTypes = True
    Select Case cmbType.ListIndex
    Case 0    'None
        SetData 0, 0, ""
        SetData 1, 0, ""
        SetData 2, 0, ""
        SetData 3, 0, ""
    Case 1    'Weapon
        SetData 0, 1, "HP / 10"
        SetData 1, 1, "Strength"
        SetData 2, 0, ""
        SetData 3, 0, ""
    Case 2    'Shield
        SetData 0, 1, "HP / 10"
        SetData 1, 1, "+Defense"
        SetData 2, 1, "+Magic Def."
        SetData 3, 0, ""
    Case 3    'Armor
        SetData 0, 1, "HP / 10"
        SetData 1, 1, "+Defense"
        SetData 2, 1, "+Magic Def."
        SetData 3, 0, ""
    Case 4    'Helmut
        SetData 0, 1, "HP / 10"
        SetData 1, 1, "+Defense"
        SetData 2, 1, "+Magic Def."
        SetData 3, 0, ""
    Case 5    'Potion
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
        SetData 3, 0, ""
    Case 6    'Money
        SetData 0, 0, ""
        SetData 1, 0, ""
        SetData 2, 0, ""
        SetData 3, 0, ""
    Case 7    'Key
        SetData 0, 2, "Unlimited Use"
        SetData 1, 0, ""
        SetData 2, 0, ""
        SetData 3, 0, ""
    Case 8    'Ring
        With cmbData(0)
            .Clear
            .AddItem "Modifies Attack"
            .AddItem "Modifies Defense"
            .AddItem "Modifies Magic Defense"
        End With
        SetData 0, 3, "Ring Type"
        SetData 1, 1, "Durability / 10"
        SetData 2, 1, "Modifier"
        SetData 3, 0, ""
    Case 9    'Guild
        SetData 0, 0, ""
        SetData 1, 0, ""
        SetData 2, 0, ""
        SetData 3, 0, ""
    Case 10    'Projectile
        SetData 0, 1, "Strength"
        SetData 1, 1, "Ammo"
        SetData 2, 1, "Ammo2"
        SetData 3, 1, "Ammo3"
    Case 11    'Ammo
        SetData 0, 1, "Bonus"
        SetData 1, 1, "Max"
        SetData 2, 1, "Type"
        SetData 3, 0, ""
    End Select

    UpdateFlags
    ChangingTypes = False
End Sub
Private Sub Form_Load()
'BitBlt picPicture.hDC, 0, 0, 32, 32, hdcObjects, 0, (sclPicture - 1) * 32, SRCCOPY
    DrawToDC 0, 0, 32, 32, picPicture.hDC, DDSObjects, 0, (sclPicture - 1) * 32
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
    cmbType.AddItem "Projectile"
    cmbType.AddItem "Ammo"

    frmObject_Loaded = True
End Sub
Sub SetData(index As Integer, DataType As Integer, St As String)
    Select Case DataType
    Case 0
        lblData(index).Visible = False
        sclData(index).Visible = False
        chkData(index).Visible = False
        cmbData(index).Visible = False
    Case 1
        lblData(index).Visible = True
        sclData(index).Visible = True
        chkData(index).Visible = False
        cmbData(index).Visible = False
        sclData(index) = Val(ObjData(index))
        lblData(index) = ObjData(index)
    Case 2
        lblData(index).Visible = True
        sclData(index).Visible = False
        chkData(index).Visible = True
        cmbData(index).Visible = False
        If Val(ObjData(index)) = 0 Then
            chkData(index) = 0
            lblData(index) = 0
        Else
            chkData(index) = 1
            lblData(index) = 1
        End If
    Case 3
        lblData(index).Visible = True
        sclData(index).Visible = False
        chkData(index).Visible = False
        cmbData(index).Visible = True
        If Val(ObjData(index)) < cmbData(index).ListCount Then
            cmbData(index).ListIndex = Val(ObjData(index))
            lblData(index) = ObjData(index)
        Else
            cmbData(index).ListIndex = cmbData(index).ListCount - 1
            lblData(index) = cmbData(index).ListCount - 1
        End If
    End Select
    lblCaption(index) = St
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmObject_Loaded = False
End Sub

Private Sub lblData_Change(index As Integer)
    If ChangingTypes = False Then
        ObjData(index) = lblData(index)
    End If
End Sub

Private Sub sclData_Change(index As Integer)
    If cmbType.ListIndex = 11 And index = 2 Then    'Its AMMO!
        Select Case sclData(index).value
        Case 0    'None
            lblCaption(2) = "None"
        Case 1    'Bow and Arrow
            lblCaption(2) = "Bow/Arrow"
        Case 2    'Fireball
            lblCaption(2) = "Fireball"
        Case 3    'Ninja Star
            lblCaption(2) = "Ninja Star"
        Case 4    'Snowball
            lblCaption(2) = "Snowball"
        Case 5    'Throwing Axe
            lblCaption(2) = "Throwing Axe"
        Case 6    'Throwing Knife
            lblCaption(2) = "Throwing Knife"
        Case 7    'Fireball2
            lblCaption(2) = "Fireball V2"
        Case 8    'Blue Thing
            lblCaption(2) = "Blue Thing"
        Case 9    'Blue Energy
            lblCaption(2) = "Energy Ball"
        Case 10    'Lightning Ball
            lblCaption(2) = "Lightning Zap"
        Case 11    'Web
            lblCaption(2) = "Web"
        Case 12    'White Ball
            lblCaption(2) = "White Ball (30)"
        Case 13    'Slime
            lblCaption(2) = "Slime (35)"
        Case 14    'Twirly
            lblCaption(2) = "Twirly (12)"
        Case 15    'Death Head
            lblCaption(2) = "Death Head (20)"
        Case 16    'Yellow Wave
            lblCaption(2) = "Yellow Wave (22)"
        Case 17    'Orange Flame
            lblCaption(2) = "Orange Flame (23)"
        Case 18    'Pink Ball
            lblCaption(2) = "Pink Ball (24)"
        Case 19    'Flash
            lblCaption(2) = "Flash (37)"
        Case 20    'Red Line
            lblCaption(2) = "Red Line (38)"
        Case 21    'Grey Ball
            lblCaption(2) = "Grey Ball (53)"
        Case 22    'Zombie
            lblCaption(2) = "Zombie (44)"
        Case 23    'Purple Spooge
            lblCaption(2) = "Purple Spooge (48)"
        Case 24    'Fire Pillar
            lblCaption(2) = "Fire Pillar (6)"
        End Select
    End If
    lblData(index) = sclData(index)
End Sub

Private Sub sclData_Scroll(index As Integer)
    sclData_Change (index)
End Sub

Private Sub sclLevel_Change()
    lblLevel = sclLevel
End Sub

Private Sub sclLevel_Scroll()
    sclLevel_Change
End Sub

Private Sub sclPicture_Change()
    DrawToDC 0, 0, 32, 32, picPicture.hDC, DDSObjects, 0, (sclPicture - 1) * 32
    picPicture.Refresh
End Sub

Private Sub sclPicture_Scroll()
    sclPicture_Change
End Sub

Sub UpdateFlags()
    Dim A As Long
    For A = 0 To 4
        chkFlags(A).Visible = False
    Next A

    Select Case cmbType.ListIndex
    Case 0    'None
        chkFlags(2).Visible = True
    Case 1    'Weapon
        For A = 0 To 4
            chkFlags(A).Visible = True
        Next A
    Case 2, 3, 4, 8, 10
        For A = 0 To 4
            chkFlags(A).Visible = True
        Next A
        chkFlags(3).Visible = False
    Case 5, 6, 7, 9, 11    'Stuff
        For A = 0 To 4
            chkFlags(A).Visible = False
        Next A
        chkFlags(2).Visible = True
    End Select
End Sub

Private Sub sclSellPrice_Change()
    txtSellPrice = sclSellPrice
End Sub

Private Sub sclSellPrice_Scroll()
    sclSellPrice_Change
End Sub

Private Sub txtSellPrice_Change()
    If Val(txtSellPrice) <= 30000 Then
        sclSellPrice.value = Val(txtSellPrice)
    Else
        sclSellPrice.value = 0
    End If
End Sub
