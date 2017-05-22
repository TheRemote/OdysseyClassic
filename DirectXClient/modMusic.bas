Attribute VB_Name = "modMusic"
Option Explicit

Public mus(1 To 28) As Long

Sub LoadMusic()
    Dim i As Long
    For i = 1 To 28
        If Exists("mus" + CStr(i) + ".mid") Then
            mus(i) = FMUSIC_LoadSong(App.Path + "\mus" + CStr(i) + ".mid")
            If mus(i) > 0 Then FMUSIC_SetLooping mus(i), 1
        End If
    Next i
End Sub

Sub UnloadMusic()
    StopMidi

    Dim i As Long
    For i = 1 To 28
        If Exists("mus" + CStr(i) + ".mid") Then
            If mus(i) > 0 Then
                FMUSIC_FreeSong (mus(i))
            End If
        End If
    Next i
End Sub

Sub StopMidi()
    If options.MIDI = True Then
        If CurrentMIDI > 0 Then
            FMUSIC_StopAllSongs
            CurrentMIDI = 0
        End If
    End If
End Sub

Sub PlayMidi(number As Long)
    If options.MIDI = True Then
        If CurrentMIDI = number Then Exit Sub
        If CurrentMIDI > 0 Then
            StopMidi
        End If
        CurrentMIDI = number

        If (number > 0 And number < 29) Then
            If mus(number) > 0 Then
                FMUSIC_SetMasterVolume mus(number), 100
                FMUSIC_PlaySong (mus(number))
            End If
        End If
    End If
End Sub
