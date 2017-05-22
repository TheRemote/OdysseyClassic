VERSION 5.00
Begin VB.Form frmOptions 
   BackColor       =   &H0061514B&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "The Odyssey Online Classic [Options]"
   ClientHeight    =   3720
   ClientLeft      =   60
   ClientTop       =   315
   ClientWidth     =   4320
   ControlBox      =   0   'False
   Icon            =   "frmOptions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3720
   ScaleWidth      =   4320
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraOptions 
      Appearance      =   0  'Flat
      BackColor       =   &H0044342E&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4095
      Begin VB.CheckBox chkDisableLighting 
         BackColor       =   &H0044342E&
         Caption         =   "Disable Lighting and Weather"
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
         Height          =   420
         Left            =   120
         TabIndex        =   14
         Top             =   840
         Width           =   1935
      End
      Begin VB.CheckBox chkDisablePlayerLights 
         BackColor       =   &H0044342E&
         Caption         =   "Disable Other Player's Lights"
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
         Height          =   300
         Left            =   120
         TabIndex        =   13
         Top             =   2520
         Width           =   2895
      End
      Begin VB.CheckBox chkHighPriority 
         BackColor       =   &H0044342E&
         Caption         =   "High Priority"
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
         Height          =   300
         Left            =   2160
         TabIndex        =   12
         Top             =   840
         Width           =   1935
      End
      Begin VB.OptionButton optLighting 
         BackColor       =   &H0044342E&
         Caption         =   "Low"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H009AADC2&
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   9
         Top             =   1680
         Width           =   1455
      End
      Begin VB.OptionButton optLighting 
         BackColor       =   &H0044342E&
         Caption         =   "Medium"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H009AADC2&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   8
         Top             =   1920
         Width           =   1455
      End
      Begin VB.OptionButton optLighting 
         BackColor       =   &H0044342E&
         Caption         =   "High"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H009AADC2&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   7
         Top             =   2160
         Width           =   1455
      End
      Begin VB.CheckBox chkBroadcasts 
         BackColor       =   &H0044342E&
         Caption         =   "Display Broadcasts"
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
         Height          =   300
         Left            =   2160
         TabIndex        =   4
         Top             =   120
         Width           =   1935
      End
      Begin VB.CheckBox chkWAV 
         BackColor       =   &H0044342E&
         Caption         =   "WAV Sound Effects"
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
         Height          =   300
         Left            =   120
         TabIndex        =   3
         Top             =   480
         Width           =   1935
      End
      Begin VB.CheckBox chkMidi 
         BackColor       =   &H0044342E&
         Caption         =   "MIDI Music"
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
         Height          =   300
         Left            =   120
         TabIndex        =   2
         Top             =   120
         Width           =   1935
      End
      Begin VB.CheckBox chkWindowed 
         BackColor       =   &H0044342E&
         Caption         =   "Windowed Mode"
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
         Height          =   300
         Left            =   2160
         TabIndex        =   1
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label cmdMacros 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0044342E&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Macros"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H009AADC2&
         Height          =   255
         Left            =   2160
         TabIndex        =   15
         Top             =   1800
         Width           =   1620
      End
      Begin VB.Label cmdChangePassword 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0044342E&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Change Password"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H009AADC2&
         Height          =   255
         Left            =   2160
         TabIndex        =   11
         Top             =   1440
         Width           =   1620
      End
      Begin VB.Label lblLightingQuality 
         BackColor       =   &H0044342E&
         Caption         =   "Lighting Quality:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H009AADC2&
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   1440
         Width           =   1935
      End
      Begin VB.Label btnCancel 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0044342E&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H009AADC2&
         Height          =   255
         Left            =   480
         TabIndex        =   6
         Top             =   3000
         Width           =   1260
      End
      Begin VB.Label btnOk 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0044342E&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Save"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H009AADC2&
         Height          =   255
         Left            =   2400
         TabIndex        =   5
         Top             =   3000
         Width           =   1185
      End
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
    With options
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
        If chkWindowed = 1 Then
            .Windowed = True
        Else
            .Windowed = False
        End If
        If chkHighPriority = 1 Then
            .HighPriority = True
            SetPriority HIGH_PRIORITY_CLASS
        Else
            .HighPriority = False
            SetPriority NORMAL_PRIORITY_CLASS
        End If
        If optLighting(0) = True Then
            .LightingQuality = 0
        ElseIf optLighting(1) = True Then
            .LightingQuality = 1
        ElseIf optLighting(2) = True Then
            .LightingQuality = 2
        End If
        If chkDisablePlayerLights = 1 Then
            .DisablePlayerLights = True
        Else
            .DisablePlayerLights = False
        End If
        If chkDisableLighting = 1 Then
            .DisableLighting = True
        Else
            .DisableLighting = False
        End If
    End With
    SaveOptions
    If blnPlaying = True Then
        RedrawMap = True
    Else
        frmMenu.Show
    End If
    Unload Me
End Sub

Private Sub cmdChangePassword_Click()
    frmNewPass.Show
    Unload Me
End Sub

Private Sub cmdMacros_Click()
    frmMacros.Show
End Sub

Private Sub Form_Load()
    With options
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
        If .Windowed = True Then
            chkWindowed = 1
        Else
            chkWindowed = 0
        End If
        If .HighPriority = True Then
            chkHighPriority = 1
        Else
            chkHighPriority = 0
        End If
        If .LightingQuality = 0 Then
            optLighting(0) = True
        ElseIf .LightingQuality = 1 Then
            optLighting(1) = True
        ElseIf .LightingQuality = 2 Then
            optLighting(2) = True
        End If
        If .DisablePlayerLights = True Then
            chkDisablePlayerLights = 1
        Else
            chkDisablePlayerLights = 0
        End If
        If .DisableLighting = True Then
            chkDisableLighting = 1
        Else
            chkDisableLighting = 0
        End If
    End With
    frmOptions_Loaded = True
    If blnPlaying = False Then
        cmdChangePassword.Visible = False
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmOptions_Loaded = False
End Sub

