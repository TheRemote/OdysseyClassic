Attribute VB_Name = "modLighting"
Option Explicit

Public Declare Function InitializeLighting Lib "odysseydll" () As Long
Public Declare Function ShadeMap16 Lib "odysseydll" (ByRef Surface As Byte) As Long
Public Declare Function ShadeMap32 Lib "odysseydll" (ByRef Surface As Byte) As Long
Public Declare Function CreateLightMap Lib "odysseydll" (ByRef LightStruct As LightSource, ByVal Alpha As Byte, ByRef MapData As Byte, ByVal OutdoorLight As Byte) As Long
Public Declare Function UpdateLightMap Lib "odysseydll" (ByRef LightStruct As LightSource) As Long

'Lighting
Type LightSource
    X As Integer
    Y As Integer
    Radius As Byte
    Intensity As Byte
    Permanent As Byte
End Type

Type LightingData
    InUse As Boolean
    Player As Byte
    Monster As Byte
    Map As Boolean
End Type

Public Darkness As Byte
Public Lighting(0 To 29) As LightSource
Public LightingData(0 To 29) As LightingData

Public ShadeMapFrame As Byte

Public AlwaysDark As Boolean
Public Indoors As Boolean

Public OutdoorLight As Byte
Public Const IndoorLight = 40
Public Const AlwaysDarkLight = 130

Public Const Radius = 80

Sub ClearLighting()
    AlwaysDark = False
    Indoors = False

    Dim i As Long
    For i = 1 To 29
        If LightingData(i).Map = False Then
            Lighting(i).Intensity = 0
            Lighting(i).Radius = 0
            Lighting(i).X = 0
            Lighting(i).Y = 0
            Lighting(i).Permanent = 0

            LightingData(i).InUse = False
            LightingData(i).Monster = 0
            LightingData(i).Player = 0
            LightingData(i).Map = False
        End If
    Next i
End Sub

Sub ClearMapLights()
    Dim i As Long
    For i = 1 To 29
        If LightingData(i).Map = True Then
            Lighting(i).Intensity = 0
            Lighting(i).Radius = 0
            Lighting(i).X = 0
            Lighting(i).Y = 0
            Lighting(i).Permanent = 0

            LightingData(i).InUse = False
            LightingData(i).Monster = 0
            LightingData(i).Player = 0
            LightingData(i).Map = False
        End If
    Next i
End Sub

Sub UpdateLights()
    If Indoors = True Then
        If AlwaysDark = True Then
            Darkness = AlwaysDarkLight
        Else
            Darkness = IndoorLight
        End If
    ElseIf AlwaysDark = True Then
        Darkness = AlwaysDarkLight
    Else
        Darkness = OutdoorLight
    End If

    Dim i As Long
    Lighting(0).X = CXO + 16
    Lighting(0).Y = CYO
    Lighting(0).Intensity = Darkness
    Lighting(0).Radius = Radius
    For i = 1 To 29
        If LightingData(i).Player > 0 Then
            If Player(LightingData(i).Player).status = 9 Or Player(LightingData(i).Player).status = 25 Then
                Lighting(i).Radius = 0
                Lighting(i).Intensity = 0
            Else
                Lighting(i).Radius = Radius
                Lighting(i).Intensity = Darkness
                Lighting(i).X = Player(LightingData(i).Player).XO + 16
                Lighting(i).Y = Player(LightingData(i).Player).YO
            End If
        End If
    Next i
End Sub

Function FreeLight() As Byte
    Dim i As Byte
    For i = 1 To 29
        If LightingData(i).InUse = False Then
            FreeLight = i
            Exit Function
        End If
    Next i
    FreeLight = 0
End Function

Sub AddPlayerLight(index As Long)
    If options.DisablePlayerLights = True Then Exit Sub
    Dim i As Byte
    i = FreeLight
    If i > 0 Then
        Lighting(i).X = Player(index).XO + 16
        Lighting(i).Y = Player(index).YO
        Lighting(i).Intensity = Darkness
        Lighting(i).Radius = Radius
        Lighting(i).Permanent = 0
        LightingData(i).Player = index
        LightingData(i).InUse = True
    End If
End Sub

Sub RemovePlayerLight(index As Long)
    Dim i As Byte
    For i = 1 To 29
        If LightingData(i).Player = index Then
            LightingData(i).Player = 0
            LightingData(i).InUse = False

            Lighting(i).X = 0
            Lighting(i).Y = 0
            Lighting(i).Radius = 0
            Lighting(i).Intensity = 0
            Lighting(i).Permanent = 0
        End If
    Next i
End Sub

Sub AddMapLight(X As Integer, Y As Integer, Radius As Byte, Intensity As Byte)
    If Indoors = True Then
        If AlwaysDark = True Then
            Darkness = AlwaysDarkLight
        Else
            Darkness = IndoorLight
        End If
    ElseIf AlwaysDark = True Then
        Darkness = AlwaysDarkLight
    Else
        Darkness = OutdoorLight
    End If

    Dim i As Byte
    i = FreeLight
    If i > 0 Then
        Lighting(i).X = X
        Lighting(i).Y = Y
        If Intensity = 0 Then
            Lighting(i).Intensity = Darkness
        Else
            Lighting(i).Intensity = Intensity
        End If
        Lighting(i).Radius = Radius
        LightingData(i).Map = True
        LightingData(i).InUse = True
        Lighting(i).Permanent = 1
    End If
End Sub
