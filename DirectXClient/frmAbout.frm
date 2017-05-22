VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About Remote Administration"
   ClientHeight    =   3375
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4575
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3375
   ScaleWidth      =   4575
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "System Info"
      Height          =   375
      Left            =   3480
      TabIndex        =   7
      Top             =   2880
      Width           =   975
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Height          =   375
      Left            =   3480
      TabIndex        =   6
      Top             =   2400
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Height          =   2175
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4335
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Caption         =   "(Client)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2160
         TabIndex        =   8
         Top             =   720
         Width           =   2055
      End
      Begin VB.Label Label4 
         Caption         =   "All rights reserved"
         Height          =   255
         Left            =   2160
         TabIndex        =   4
         Top             =   1800
         Width           =   2055
      End
      Begin VB.Label Label3 
         Caption         =   "Copyright© Usssssy"
         Height          =   255
         Left            =   2160
         TabIndex        =   3
         Top             =   1440
         Width           =   2055
      End
      Begin VB.Label Label2 
         Caption         =   "Version:  5.0.1"
         Height          =   255
         Left            =   2160
         TabIndex        =   2
         Top             =   1080
         Width           =   2055
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Remote Administration"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2160
         TabIndex        =   1
         Top             =   240
         Width           =   2055
      End
      Begin VB.Image Image1 
         Height          =   1785
         Left            =   120
         Picture         =   "frmAbout.frx":0442
         Top             =   240
         Width           =   1905
      End
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "For any help or bug reports, Please contact me at:  usssssy@yahoo.com"
      Height          =   435
      Left            =   120
      TabIndex        =   5
      Top             =   2520
      Width           =   3210
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const READ_CONTROL = &H20000
Const KEY_QUERY_VALUE = &H1
Const KEY_SET_VALUE = &H2
Const KEY_CREATE_SUB_KEY = &H4
Const KEY_ENUMERATE_SUB_KEYS = &H8
Const KEY_NOTIFY = &H10
Const KEY_CREATE_LINK = &H20
Const KEY_ALL_ACCESS = KEY_QUERY_VALUE + KEY_SET_VALUE + _
                       KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS + _
                       KEY_NOTIFY + KEY_CREATE_LINK + READ_CONTROL
                     
'// Reg Key ROOT Types...
Const HKEY_LOCAL_MACHINE = &H80000002
Const ERROR_SUCCESS = 0
Const REG_SZ = 1                         '// Unicode nul terminated string
Const REG_DWORD = 4                      '// 32-bit number

Const gREGKEYSYSINFOLOC = "SOFTWARE\Microsoft\Shared Tools Location"
Const gREGVALSYSINFOLOC = "MSINFO"
Const gREGKEYSYSINFO = "SOFTWARE\Microsoft\Shared Tools\MSINFO"
Const gREGVALSYSINFO = "PATH"

Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long

Private Sub cmdOK_Click()
Unload Me
End Sub

Private Sub Command2_Click()
frmSysInfo.Show
'Call StartSysInfo
End Sub
Public Sub StartSysInfo()
On Error GoTo SysInfoErr
Dim rc As Long
Dim SysInfoPath As String
If GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFO, gREGVALSYSINFO, SysInfoPath) Then
ElseIf GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFOLOC, gREGVALSYSINFOLOC, SysInfoPath) Then
    If (Dir(SysInfoPath & "\MSINFO32.EXE") <> "") Then
        SysInfoPath = SysInfoPath & "\MSINFO32.EXE"
    Else
        GoTo SysInfoErr
    End If
Else
    GoTo SysInfoErr
End If
Call Shell(SysInfoPath, vbNormalFocus)
Exit Sub
SysInfoErr:
    MsgBox "System Information Is Unavailable At This Time", vbOKOnly
End Sub
Public Function GetKeyValue(KeyRoot As Long, Keyname As String, SubKeyRef As String, ByRef KeyVal As String) As Boolean
Dim i As Long                                           '// Loop Counter
Dim rc As Long                                          '// Return Code
Dim hKey As Long                                        '// Handle To An Open Registry Key
Dim hDepth As Long                                      '//
Dim KeyValType As Long                                  '// Data Type Of A Registry Key
Dim tmpVal As String                                    '// Tempory Storage For A Registry Key Value
Dim KeyValSize As Long                                  '// Size Of Registry Key Variable
rc = RegOpenKeyEx(KeyRoot, Keyname, 0, KEY_ALL_ACCESS, hKey) '// Open Registry Key
If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError               '// Handle Error...
tmpVal = String$(1024, 0)                             '// Allocate Variable Space
KeyValSize = 1024                                       '// Mark Variable Size
rc = RegQueryValueEx(hKey, SubKeyRef, 0, _
    KeyValType, tmpVal, KeyValSize)
If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError
If (Asc(Mid(tmpVal, KeyValSize, 1)) = 0) Then           '// Win95 Adds Null Terminated String...
    tmpVal = Left(tmpVal, KeyValSize - 1)
Else                                                    '// WinNT Does NOT Null Terminate String...
    tmpVal = Left(tmpVal, KeyValSize)                   '// Null Not Found, Extract String Only
End If
Select Case KeyValType
    Case REG_SZ
        KeyVal = tmpVal
    Case REG_DWORD
        For i = Len(tmpVal) To 1 Step -1
            KeyVal = KeyVal + Hex(Asc(Mid(tmpVal, i, 1)))
        Next
        KeyVal = Format$("&h" + KeyVal)
End Select
GetKeyValue = True
rc = RegCloseKey(hKey)
Exit Function
GetKeyError:
    KeyVal = ""
    GetKeyValue = False
    rc = RegCloseKey(hKey)
End Function

