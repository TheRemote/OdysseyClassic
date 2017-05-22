VERSION 5.00
Begin VB.Form frmMenu 
   BackColor       =   &H0061514B&
   BorderStyle     =   0  'None
   Caption         =   "The Odyssey Online Classic"
   ClientHeight    =   2820
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4485
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmMenu.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2820
   ScaleWidth      =   4485
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label lblMenu 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0044342E&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Map Editor"
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
      Height          =   420
      Index           =   6
      Left            =   2280
      TabIndex        =   7
      Top             =   1815
      Width           =   2100
   End
   Begin VB.Label lblMenu 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0044342E&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Home Page"
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
      Height          =   420
      Index           =   5
      Left            =   120
      TabIndex        =   6
      Top             =   1815
      Width           =   2100
   End
   Begin VB.Label lblMenu 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0044342E&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Login"
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
      Height          =   420
      Index           =   0
      Left            =   120
      TabIndex        =   5
      Top             =   840
      Width           =   2100
   End
   Begin VB.Label lblMenu 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0044342E&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "New Account"
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
      Height          =   420
      Index           =   1
      Left            =   2280
      TabIndex        =   4
      Top             =   840
      Width           =   2100
   End
   Begin VB.Label lblMenu 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0044342E&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Options"
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
      Height          =   420
      Index           =   2
      Left            =   2280
      TabIndex        =   3
      Top             =   1335
      Width           =   2100
   End
   Begin VB.Label lblMenu 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0044342E&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Credits"
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
      Height          =   420
      Index           =   3
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   2100
   End
   Begin VB.Label lblMenu 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0044342E&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Quit"
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
      Height          =   420
      Index           =   4
      Left            =   120
      TabIndex        =   1
      Top             =   2295
      Width           =   4260
   End
   Begin VB.Label lblCurrentServer 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
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
      TabIndex        =   0
      Top             =   240
      Width           =   4215
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_GotFocus()
    lblCurrentServer = "Main Menu"
End Sub

Private Sub Form_Load()
    gHW = Me.hwnd


    Dim File As String
    Dim FileByteArray() As Byte

    File = "menu.rsc"
    FileByteArray() = StrConv(File, vbFromUnicode)
    ReDim Preserve FileByteArray(UBound(FileByteArray) + 1)

    EncryptDataFile FileByteArray(0), FileLen(File) Mod 87 + 5
    frmMenu.Picture = LoadPicture(File)
    EncryptDataFile FileByteArray(0), FileLen(File) Mod 87 + 5

    lblCurrentServer = ServerDescription
End Sub

Private Sub lblMenu_MouseDown(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblMenu(index).BackColor = &H61514B
End Sub

Private Sub lblMenu_MouseUp(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblMenu(index).BackColor = &H44342E
    If X >= 0 And X <= lblMenu(index).Width And Y >= 0 And Y <= lblMenu(index).Height Then
        Select Case index
        Case 0    'Login
            frmLogin.Show
            Me.Hide
        Case 1    'New Account
            frmAccount.Show
            Me.Hide
        Case 2    'Options
            frmOptions.Show
            Me.Hide
        Case 3    'Credits
            frmCredits.Show
            Me.Hide
        Case 4    'Quit
            blnEnd = True
        Case 5    'Home Page
            Dim sTopic As String
            Dim sFile As String
            Dim sParams As String
            Dim sDirectory As String
            sTopic = "Open"
            sFile = TheWebSite
            sParams = 0&
            sDirectory = 0&

            RunShellExecute sTopic, sFile, sParams, sDirectory, 1
        Case 6    'Map Editor
            sTopic = "Open"
            sFile = "mapeditor.exe"
            sParams = 0&
            sDirectory = 0&

            RunShellExecute sTopic, sFile, sParams, sDirectory, 1
        End Select
    End If
End Sub

Sub RunShellExecute(sTopic As String, sFile As Variant, _
                    sParams As Variant, sDirectory As Variant, _
                    nShowCmd As Long)

    Dim hWndDesk As Long
    Dim SUCCESS As Long

    'the desktop will be the
    'default for error messages
    hWndDesk = GetDesktopWindow()

    'execute the passed operation
    SUCCESS = ShellExecute(hWndDesk, sTopic, sFile, sParams, sDirectory, nShowCmd)

    'This is optional. comment the three lines
    'below to not have the "Open With.." dialog appear
    'when the ShellExecute API call fails
    If SUCCESS = 31 Then
        MsgBox "Couldn't load the default application"    'shouldn't happen
        'but if it does, try to get the user to make an association...
        Call Shell("rundll32.exe shell32.dll,OpenAs_RunDLL " & sFile, vbNormalFocus)
    End If
End Sub
