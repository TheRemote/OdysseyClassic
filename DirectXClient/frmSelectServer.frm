VERSION 5.00
Begin VB.Form frmSelectServer 
   BackColor       =   &H0061514B&
   BorderStyle     =   0  'None
   Caption         =   "Select Odyssey Server"
   ClientHeight    =   2535
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3735
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2535
   ScaleWidth      =   3735
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   2280
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.ListBox lstServers 
      BackColor       =   &H0044342E&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H009AADC2&
      Height          =   1230
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   3495
   End
   Begin VB.Label btnCancel 
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
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   2040
      Width           =   1500
   End
   Begin VB.Label btnOk 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0044342E&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Continue"
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
      Left            =   1680
      TabIndex        =   1
      Top             =   2040
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackColor       =   &H0044342E&
      Caption         =   "Choose which Odyssey server you would like to connect to:"
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
      Height          =   615
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   3495
   End
End
Attribute VB_Name = "frmSelectServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim NumPlayers(5) As Integer

Private Sub btnCancel_Click()
    Unload Me
    frmMenu.Show
End Sub

Private Sub btnOk_Click()
    Dim St As String
    
    If Exists("servers.dat") Then
        Open "servers.dat" For Input As #1
        
        Dim ListLoop As Long
        While Not EOF(1)
            Line Input #1, St
            GetSections3 DecipherString(St), ","
            If Section(1) <> "" And Section(2) <> "" And Section(3) <> "" Then
                If ListLoop = lstServers.ListIndex Then
                    ChosenServer = True
                    ServerDescription = Section(1)
                    CacheDirectory = Section(3)
                    ServerIP = Section(2)
                    ServerPort = Section(4)
                End If
                
                ListLoop = ListLoop + 1
            End If
        Wend
        
        Close #1
    Else
        MsgBox "Servers.dat not found!  Redownload the client!"
    End If
    
    CheckCache
    
    frmMenu.Show
    Unload Me
End Sub

Function EncryptString(St As String) As String
Dim TempStr As String, TempStr2 As String
Dim A As Integer, TmpNum As Integer

TempStr = ""
TempStr2 = ""

For A = 1 To Len(St)
    TempStr = Mid$(St, A, 1)
    TmpNum = Asc(TempStr)
    TempStr2 = TempStr2 + Chr$(TmpNum + 40)
Next A

EncryptString = Trim$(TempStr2)
End Function
Function DecipherString(St As String) As String
Dim TempStr As String, TempStr2 As String
Dim A As Integer, TmpNum As Integer

TempStr = ""
TempStr2 = ""

For A = 1 To Len(St)
    TempStr = Mid$(St, A, 1)
    TmpNum = Asc(TempStr)
    TempStr2 = TempStr2 + Chr$(TmpNum - 40)
Next A

DecipherString = Trim$(TempStr2)
End Function

Private Sub Command1_Click()
    Dim St As String
    
    If Exists("realservers.dat") Then
        Open "realservers.dat" For Input As #1
        Open "servers.dat" For Output As #2
        Dim ListLoop As Long
        
        While Not EOF(1)
            Line Input #1, St
            Print #2, EncryptString(St)
            
            ListLoop = ListLoop + 1
        Wend
        Close #1
        Close #2
    End If
End Sub

Private Sub Form_Load()
Command1_Click
    Dim A As Long
    For A = 0 To 5
        NumPlayers(A) = -1
    Next A
    Dim St As String
    
    If Exists("servers.dat") Then
        Open "servers.dat" For Input As #1
        Dim ListLoop As Long
        
        While Not EOF(1)
            Line Input #1, St
            GetSections3 DecipherString(St), ","
            If Section(1) <> "" And Section(2) <> "" And Section(3) <> "" Then
                lstServers.AddItem Section(1)
            End If
            
            ListLoop = ListLoop + 1
        Wend
        Close #1
    Else
        MsgBox "Servers.dat not found!  Redownload the client!"
        End
    End If
    
    lstServers.ListIndex = 0
End Sub

Private Sub UpdateList()
    Dim St As String
    lstServers.Clear
    
    If Exists("servers.dat") Then
        Open "servers.dat" For Input As #1
        Dim ListLoop As Long
        While Not EOF(1)
            Line Input #1, St
            GetSections3 DecipherString(St), ","
            If Section(1) <> "" And Section(2) <> "" And Section(3) <> "" Then
                lstServers.AddItem Section(1)
            End If
            ListLoop = ListLoop + 1
        Wend
        Close #1
    Else
        MsgBox "Servers.dat not found!  Redownload the client!"
        End
    End If
    
    lstServers.ListIndex = 0
End Sub
