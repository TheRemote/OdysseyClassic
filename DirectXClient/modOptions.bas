Attribute VB_Name = "modOptions"
Option Explicit

Type OptionsData
    MIDI As Boolean
    Wav As Boolean
    Broadcasts As Boolean
    Windowed As Boolean
    LightingQuality As Byte
    Bit32 As Boolean
    HighPriority As Boolean
    DisablePlayerLights As Boolean
    DisableLighting As Boolean
End Type

Public options As OptionsData

Sub SaveOptions()
    With options
        WriteString "Options", "Saved", "1"
        If .MIDI = True Then
            WriteString "Options", "MIDI", "1"
        Else
            WriteString "Options", "MIDI", "0"
        End If
        If .Wav = True Then
            WriteString "Options", "Wav", "1"
        Else
            WriteString "Options", "Wav", "0"
        End If
        If .Broadcasts = True Then
            WriteString "Options", "Broadcasts", "1"
        Else
            WriteString "Options", "Broadcasts", "0"
        End If
        If .Windowed = True Then
            WriteString "Options", "Windowed", "1"
        Else
            WriteString "Options", "Windowed", "0"
        End If
        If .HighPriority = True Then
            WriteString "Options", "HighPriority", "1"
        Else
            WriteString "Options", "HighPriority", "0"
        End If
        If .DisablePlayerLights = True Then
            WriteString "Options", "DisablePlayerLights", "1"
        Else
            WriteString "Options", "DisablePlayerLights", "0"
        End If
        If .DisableLighting = True Then
            WriteString "Options", "DisableLighting", "1"
        Else
            WriteString "Options", "DisableLighting", "0"
        End If
        WriteString "Options", "LightingQuality", CStr(.LightingQuality)
        If Character.name <> "" Then
            Dim A As Long
            For A = 1 To 12
                WriteString Character.name, "Hotkey" + CStr(A), Character.Hotkey(A).Hotkey
                WriteString Character.name, "Hotkey" + CStr(A) + "Type", Character.Hotkey(A).Type
                WriteString Character.name, "Hotkey" + CStr(A) + "ScrollPosition", Character.Hotkey(A).ScrollPosition
            Next A
        End If
        If blnPlaying = True Then
            RedrawMap = True
        End If
    End With
End Sub
Sub LoadOptions()
    With options
        If ReadInt("Options", "Saved") = 1 Then
            If ReadInt("Options", "MIDI") = 1 Then
                .MIDI = True
            Else
                .MIDI = False
            End If
            If ReadInt("Options", "Wav") = 1 Then
                .Wav = True
            Else
                .Wav = False
            End If
            If ReadInt("Options", "Broadcasts") = 1 Then
                .Broadcasts = True
            Else
                .Broadcasts = False
            End If
            If ReadInt("Options", "Windowed") = 1 Then
                .Windowed = True
            Else
                .Windowed = False
            End If
            If ReadInt("Options", "HighPriority") = 1 Then
                .HighPriority = True
            Else
                .HighPriority = False
            End If
            If ReadInt("Options", "DisablePlayerLights") = 1 Then
                .DisablePlayerLights = True
            Else
                .DisablePlayerLights = False
            End If
            If ReadInt("Options", "DisableLighting") = 1 Then
                .DisableLighting = True
            Else
                .DisableLighting = False
            End If
            .LightingQuality = ReadInt("Options", "LightingQuality")
            If Character.name <> "" Then
                Dim A As Long
                For A = 1 To 12
                    Character.Hotkey(A).Hotkey = ReadInt(Character.name, "Hotkey" + CStr(A))
                    Character.Hotkey(A).Type = ReadInt(Character.name, "Hotkey" + CStr(A) + "Type")
                    Character.Hotkey(A).ScrollPosition = ReadInt(Character.name, "Hotkey" + CStr(A) + "ScrollPosition")
                Next A
            End If
        Else
            .MIDI = True
            .Wav = True
            .Broadcasts = True
            .Windowed = True
            .LightingQuality = 1
            .HighPriority = False
            .DisablePlayerLights = False
            .DisableLighting = False
            SaveOptions
        End If
    End With
End Sub
