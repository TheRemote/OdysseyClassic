VERSION 5.00
Begin VB.Form frmOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "The Odyssey Online Classic [Options]"
   ClientHeight    =   2790
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3465
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2790
   ScaleWidth      =   3465
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkAway 
      Caption         =   "Away"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   1560
      Width           =   2415
   End
   Begin VB.CheckBox chkBroadcasts 
      Caption         =   "Display Broadcasts"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   2415
   End
   Begin VB.CheckBox chkWAV 
      Caption         =   "WAV Sound Effects"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   2415
   End
   Begin VB.CheckBox chkMidi 
      Caption         =   "MIDI Music"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
   Begin VB.CommandButton btnOk 
      Caption         =   "Ok"
      Height          =   495
      Left            =   1920
      TabIndex        =   3
      Top             =   2160
      Width           =   1455
   End
   Begin VB.CommandButton btnCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   495
      Left            =   360
      TabIndex        =   2
      Top             =   2160
      Width           =   1455
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub btnCancel_Click()
    Unload Me
    If blnPlaying = False Then frmMenu.Show
End Sub


Private Sub btnOk_Click()
    With Options
        If chkMidi = 1 Then
            .MIDI = True
        Else
            If .MIDI = True Then
                StopMidi
            End If
            .MIDI = False
        End If
        If chkWAV = 1 Then
            .Wav = True
        Else
            .Wav = False
        End If
        If chkBroadcasts = 1 Then
            .Broadcasts = True
        Else
            .Broadcasts = False
        End If
        If chkAway = 1 Then
            If .ForwardUser <> "" Then
                .ForwardUser = ""
                .Away = True
                MsgBox "Since forwarding was on it is now set to off, away mode is active!", vbCritical + vbOKOnly, "Changing Status"
            Else
                .Away = True
            End If
            PrintChat "Away mode is active, please use /AMSG <message> to set your away message, it is set to default now.", 14
        Else
            .Away = False
        End If
    End With
    SaveOptions
    Unload Me
    If blnPlaying = False Then frmMenu.Show
End Sub


Private Sub Form_Load()
    With Options
        If .MIDI = True Then
            chkMidi = 1
        Else
            chkMidi = 0
        End If
        If .Wav = True Then
            chkWAV = 1
        Else
            chkWAV = 0
        End If
        If .Broadcasts = True Then
            chkBroadcasts = 1
        Else
            chkBroadcasts = 0
        End If
        If .Away = True Then
            chkAway = 1
        Else
            chkAway = 0
        End If
    End With
    frmOptions_Loaded = True
End Sub
Private Sub Form_LostFocus()
    Unload Me
End Sub


Private Sub Form_Unload(Cancel As Integer)
    frmOptions_Loaded = False
End Sub


