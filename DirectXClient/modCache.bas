Attribute VB_Name = "modCache"
Option Explicit

Sub LoadObject(TheObj As Integer)
    Dim St As String * 27
    Open CacheDirectory + "/ocache.dat" For Random As #1 Len = 27
    Get #1, TheObj, St
    Close #1
    With Object(TheObj)
        .name = ClipString$(Mid$(St, 1, 15))
        .Picture = Asc(Mid$(St, 16, 1)) * 256 + Asc(Mid$(St, 17, 1))
        .Type = Asc(Mid$(St, 18, 1))
        Select Case .Type
        Case 8    'Ring
            .MaxDur = Asc(Mid$(St, 20, 1))
            .Data2 = Asc(Mid$(St, 19, 1))
            .Modifier = Asc(Mid$(St, 21, 1))
        Case 10, 11    'Projectile Weapon
            .Modifier = Asc(Mid$(St, 19, 1))
            .Data2 = Asc(Mid$(St, 21, 1))
        Case Else
            .MaxDur = Asc(Mid$(St, 19, 1))
            .Modifier = Asc(Mid$(St, 20, 1))
            .Data2 = Asc(Mid$(St, 21, 1))
        End Select
        .flags = Asc(Mid$(St, 22, 1))
        .ClassReq = Asc(Mid$(St, 23, 1))
        .LevelReq = Asc(Mid$(St, 24, 1))
        .Version = Asc(Mid$(St, 25, 1))
        .SellPrice = Asc(Mid$(St, 26, 1)) * 256 + Asc(Mid$(St, 27, 1))
    End With
End Sub

Sub SaveObject(TheObj As Long)
    Dim St1 As String * 15
    Dim WritableData As String * 27
    With Object(TheObj)
        St1 = .name
        Select Case .Type
        Case 8    'Ring
            WritableData = St1 + DoubleChar$(CLng(.Picture)) + Chr$(.Type) + Chr$(.Data2) + Chr$(.MaxDur) + Chr$(.Modifier) + Chr$(.flags) + Chr$(.ClassReq) + Chr$(.LevelReq) + Chr$(.Version) + DoubleChar$(CLng(.SellPrice))
        Case 10, 11    'Projectile Weapon
            WritableData = St1 + DoubleChar$(CLng(.Picture)) + Chr$(.Type) + Chr$(.Modifier) + vbNullChar + Chr$(.Data2) + Chr$(.flags) + Chr$(.ClassReq) + Chr$(.LevelReq) + Chr$(.Version) + DoubleChar$(CLng(.SellPrice))
        Case Else
            WritableData = St1 + DoubleChar$(CLng(.Picture)) + Chr$(.Type) + Chr$(.MaxDur) + Chr$(.Modifier) + Chr$(.Data2) + Chr$(.flags) + Chr$(.ClassReq) + Chr$(.LevelReq) + Chr$(.Version) + DoubleChar$(CLng(.SellPrice))
        End Select
    End With
    Open CacheDirectory + "/ocache.dat" For Random As #1 Len = 27
    Put #1, TheObj, WritableData
    Close #1
End Sub

Sub CreateMapCache()
    Dim St1 As String * 2677, A As Long
    St1 = String$(2677, 0)
    Open CacheDirectory + "/cache1.dat" For Random As #1 Len = 2677
    For A = 1 To MaxMaps
        Put #1, , St1
    Next A
    Close #1
End Sub

Sub CreateObjectCache()
    If Exists(CacheDirectory + "/ocache.dat") Then Kill CacheDirectory + "/ocache.dat"
    Dim St1 As String * 27, A As Long
    St1 = String$(27, 0)
    Open CacheDirectory + "/ocache.dat" For Random As #1 Len = 27
    For A = 1 To MaxObjects
        Put #1, , St1
    Next A
    Close #1
End Sub

Sub CreateNPCCache()
    If Exists(CacheDirectory + "/ncache.dat") Then Kill CacheDirectory + "/ncache.dat"
    Dim St1 As String * 137, A As Long
    St1 = String$(137, 0)
    Open CacheDirectory + "/ncache.dat" For Random As #1 Len = 137
    For A = 1 To MaxNPCs
        Put #1, , St1
    Next A
    Close #1
End Sub

Sub CreateHallCache()
    If Exists(CacheDirectory + "/hcache.dat") Then Kill CacheDirectory + "/hcache.dat"
    Dim St1 As String * 16, A As Long
    St1 = String$(16, 0)
    Open CacheDirectory + "/hcache.dat" For Random As #1 Len = 16
    For A = 1 To MaxHalls
        Put #1, , St1
    Next A
    Close #1
End Sub

Sub CreateMonsterCache()
    If Exists(CacheDirectory + "/moncache.dat") Then Kill CacheDirectory + "/moncache.dat"
    Dim St1 As String * 21, A As Long
    St1 = String$(21, 0)
    Open CacheDirectory + "/moncache.dat" For Random As #1 Len = 21
    For A = 1 To MaxTotalMonsters
        Put #1, , St1
    Next A
    Close #1
End Sub

Sub CreateMagicCache()
    If Exists(CacheDirectory + "/magcache.dat") Then Kill CacheDirectory + "/magcache.dat"
    Dim St1 As String * 124, A As Long
    St1 = String$(124, 0)
    Open CacheDirectory + "/magcache.dat" For Random As #1 Len = 124
    For A = 1 To MaxMagic
        Put #1, , St1
    Next A
    Close #1
End Sub

Sub CreatePrefixCache()
    If Exists(CacheDirectory + "/itemprecache.dat") Then Kill CacheDirectory + "/itemprecache.dat"
    Dim St1 As String * 19, A As Long
    St1 = String$(19, 0)
    Open CacheDirectory + "/itemprecache.dat" For Random As #1 Len = 19
    For A = 1 To MaxModifications
        Put #1, , St1
    Next A
    Close #1
End Sub

Sub CreateSuffixCache()
    If Exists(CacheDirectory + "/itemsufcache.dat") Then Kill CacheDirectory + "/itemsufcache.dat"
    Dim St1 As String * 19, A As Long
    St1 = String$(19, 0)
    Open CacheDirectory + "/itemsufcache.dat" For Random As #1 Len = 19
    For A = 1 To MaxModifications
        Put #1, , St1
    Next A
    Close #1
End Sub

Sub SaveNPC(TheNPC As Long)
    Dim St1 As String * 15
    Dim WritableData As String * 137
    Dim OutputString As String
    Dim A As Long

    With NPC(TheNPC)
        St1 = .name
        OutputString = ""
        For A = 0 To 9
            OutputString = OutputString + DoubleChar$(CLng(.SaleItem(A).GiveObject)) + QuadChar$(.SaleItem(A).GiveValue) + DoubleChar$(CLng(.SaleItem(A).TakeObject)) + QuadChar$(.SaleItem(A).TakeValue)
        Next A
        WritableData = St1 + Chr$(.Version) + Chr$(.flags) + OutputString
    End With

    Open CacheDirectory + "/ncache.dat" For Random As #1 Len = 137
    Put #1, TheNPC, WritableData
    Close #1
End Sub

Sub SaveHall(TheHall As Long)
    Dim St1 As String * 15
    Dim WritableData As String * 16
    With Hall(TheHall)
        St1 = .name
        WritableData = St1 + Chr$(.Version)
    End With
    Open CacheDirectory + "/hcache.dat" For Random As #1 Len = 16
    Put #1, TheHall, WritableData
    Close #1
End Sub

Sub SaveMonster(TheMonster As Long)
    Dim St1 As String * 15
    Dim WritableData As String * 21
    With Monster(TheMonster)
        St1 = .name
        WritableData = St1 + DoubleChar$(CLng(.Sprite)) + Chr$(.Version) + DoubleChar$(CLng(.MaxLife)) + Chr$(.flags)
    End With
    Open CacheDirectory + "/moncache.dat" For Random As #1 Len = 21
    Put #1, TheMonster, WritableData
    Close #1
End Sub

Sub LoadNPC(TheNPC As Integer)
    Dim St As String * 137
    Dim A As Long, B As Long
    Open CacheDirectory + "/ncache.dat" For Random As #1 Len = 137
    Get #1, TheNPC, St
    Close #1
    With NPC(TheNPC)
        .name = ClipString$(Mid$(St, 1, 15))
        .Version = Asc(Mid$(St, 16, 1))
        .flags = Asc(Mid$(St, 17, 1))
        For A = 0 To 9
            B = 18 + A * 12
            .SaleItem(A).GiveObject = Asc(Mid$(St, B, 1)) * 256& + Asc(Mid$(St, B + 1, 1))
            .SaleItem(A).GiveValue = Asc(Mid$(St, B + 2, 1)) * 16777216 + Asc(Mid$(St, B + 3, 1)) * 65536 + Asc(Mid$(St, B + 4, 1)) * 256& + Asc(Mid$(St, B + 5, 1))
            .SaleItem(A).TakeObject = Asc(Mid$(St, B + 6, 1)) * 256& + Asc(Mid$(St, B + 7, 1))
            .SaleItem(A).TakeValue = Asc(Mid$(St, B + 8, 1)) * 16777216 + Asc(Mid$(St, B + 9, 1)) * 65536 + Asc(Mid$(St, B + 10, 1)) * 256& + Asc(Mid$(St, B + 11, 1))
        Next A
    End With
End Sub

Sub LoadHall(TheHall As Integer)
    Dim St As String * 16
    Open CacheDirectory + "/hcache.dat" For Random As #1 Len = 16
    Get #1, TheHall, St
    Close #1
    With Hall(TheHall)
        .name = ClipString$(Mid$(St, 1, 15))
        .Version = Asc(Mid$(St, 16, 1))
    End With
End Sub

Sub LoadMonster(TheMonster As Integer)
    Dim St As String * 21
    Open CacheDirectory + "/moncache.dat" For Random As #1 Len = 21
    Get #1, TheMonster, St
    Close #1
    With Monster(TheMonster)
        .name = ClipString$(Mid$(St, 1, 15))
        .Sprite = Asc(Mid$(St, 16, 1)) * 256 + Asc(Mid$(St, 17, 1))
        .Version = Asc(Mid$(St, 18, 1))
        .MaxLife = GetInt(Mid$(St, 19, 2))
        .flags = Asc(Mid$(St, 21, 1))
    End With
End Sub

Sub LoadMagic(TheMagic As Integer)
    Dim St As String * 124
    Open CacheDirectory + "/magcache.dat" For Random As #1 Len = 124
    Get #1, TheMagic, St
    Close #1
    With Magic(TheMagic)
        .name = ClipString$(Mid$(St, 1, 15))
        .Version = Asc(Mid$(St, 16, 1))
        If .Version > 0 Then
            .Level = Asc(Mid$(St, 17, 1))
            .Class = Asc(Mid$(St, 18, 1))
            .Icon = Asc(Mid$(St, 19, 1)) * 256 + Asc(Mid$(St, 20, 1))
            .IconType = Asc(Mid$(St, 21, 1))
            .CastTimer = Asc(Mid$(St, 22, 1)) * 256 + Asc(Mid$(St, 23, 1))
            .Description = ClipString$(Mid$(St, 24, 100))
        End If
    End With
End Sub

Sub SaveMagic(TheMagic As Integer)
    Dim St1 As String * 15
    Dim St2 As String * 100
    Dim WritableData As String * 124
    With Magic(TheMagic)
        St1 = .name
        St2 = .Description
        WritableData = St1 + Chr$(.Version) + Chr$(.Level) + Chr$(.Class) + DoubleChar$(CLng(.Icon)) + Chr$(.IconType) + DoubleChar$(CLng(.CastTimer)) + St2
    End With
    Open CacheDirectory + "/magcache.dat" For Random As #1 Len = 124
    Put #1, TheMagic, WritableData
    Close #1
End Sub

Sub LoadPrefix(ThePrefix As Integer)
    Dim St As String * 19
    Open CacheDirectory + "/itemprecache.dat" For Random As #1 Len = 19
    Get #1, ThePrefix, St
    Close #1
    With ItemPrefix(ThePrefix)
        .name = ClipString$(Mid$(St, 1, 15))
        .Version = Asc(Mid$(St, 16, 1))
        .ModificationType = Asc(Mid$(St, 17, 1))
        .ModificationValue = Asc(Mid$(St, 18, 1))
        .OccursNaturally = Asc(Mid$(St, 19, 1))
    End With
End Sub

Sub LoadSuffix(TheSuffix As Integer)
    Dim St As String * 19
    Open CacheDirectory + "/itemsufcache.dat" For Random As #1 Len = 19
    Get #1, TheSuffix, St
    Close #1
    With ItemSuffix(TheSuffix)
        .name = ClipString$(Mid$(St, 1, 15))
        .Version = Asc(Mid$(St, 16, 1))
        .ModificationType = Asc(Mid$(St, 17, 1))
        .ModificationValue = Asc(Mid$(St, 18, 1))
        .OccursNaturally = Asc(Mid$(St, 19, 1))
    End With
End Sub

Sub SavePrefix(ThePrefix As Byte)
    Dim St1 As String * 15
    Dim WritableData As String * 19
    With ItemPrefix(ThePrefix)
        St1 = .name
        WritableData = St1 + Chr$(.Version) + Chr$(.ModificationType) + Chr$(.ModificationValue) + Chr$(.OccursNaturally)
    End With
    Open CacheDirectory + "/itemprecache.dat" For Random As #1 Len = 19
    Put #1, ThePrefix, WritableData
    Close #1
End Sub

Sub SaveSuffix(TheSuffix As Byte)
    Dim St1 As String * 15
    Dim WritableData As String * 19
    With ItemSuffix(TheSuffix)
        St1 = .name
        WritableData = St1 + Chr$(.Version) + Chr$(.ModificationType) + Chr$(.ModificationValue) + Chr$(.OccursNaturally)
    End With
    Open CacheDirectory + "/itemsufcache.dat" For Random As #1 Len = 19
    Put #1, TheSuffix, WritableData
    Close #1
End Sub

Sub CheckCache()
    If Exists(CacheDirectory + "/cache1.dat") = False Then
        frmWait.lblStatus = "Creating Map Cache .."
        frmWait.Refresh
        CreateMapCache
    Else
        If FileLen(CacheDirectory + "/cache1.dat") <> 8031000 Then
            frmWait.lblStatus = "Creating Map Cache .."
            frmWait.Refresh
            CreateMapCache
        End If
    End If

    If Exists(CacheDirectory + "/ocache.dat") = False Then
        frmWait.lblStatus = "Creating Object Cache .."
        frmWait.Refresh
        CreateObjectCache
    Else
        If FileLen(CacheDirectory + "/ocache.dat") <> 27000 Then
            frmWait.lblStatus = "Creating Object Cache .."
            frmWait.Refresh
            CreateObjectCache
        End If
    End If

    If Exists(CacheDirectory + "/hcache.dat") = False Then
        frmWait.lblStatus = "Creating Hall Cache .."
        frmWait.Refresh
        CreateHallCache
    Else
        If FileLen(CacheDirectory + "/hcache.dat") <> 4080 Then
            frmWait.lblStatus = "Creating Hall Cache .."
            frmWait.Refresh
            CreateHallCache
        End If
    End If

    If Exists(CacheDirectory + "/ncache.dat") = False Then
        frmWait.lblStatus = "Creating NPC Cache .."
        frmWait.Refresh
        CreateNPCCache
    Else
        If FileLen(CacheDirectory + "/ncache.dat") <> 68500 Then
            frmWait.lblStatus = "Creating NPC Cache .."
            frmWait.Refresh
            CreateNPCCache
        End If
    End If

    If Exists(CacheDirectory + "/moncache.dat") = False Then
        frmWait.lblStatus = "Creating Monster Cache .."
        frmWait.Refresh
        CreateMonsterCache
    Else
        If FileLen(CacheDirectory + "/moncache.dat") <> 21000 Then
            frmWait.lblStatus = "Creating Monster Cache .."
            frmWait.Refresh
            CreateMonsterCache
        End If
    End If

    If Exists(CacheDirectory + "/magcache.dat") = False Then
        frmWait.lblStatus = "Creating Magic Cache .."
        frmWait.Refresh
        CreateMagicCache
    Else
        If FileLen(CacheDirectory + "/magcache.dat") <> 62000 Then
            frmWait.lblStatus = "Creating Magic Cache .."
            frmWait.Refresh
            CreateMagicCache
        End If
    End If

    If Exists(CacheDirectory + "/itemprecache.dat") = False Then
        frmWait.lblStatus = "Creating Prefix Cache .."
        frmWait.Refresh
        CreatePrefixCache
    Else
        If FileLen(CacheDirectory + "/itemprecache.dat") <> 4845 Then
            frmWait.lblStatus = "Creating Prefix Cache .."
            frmWait.Refresh
            CreatePrefixCache
        End If
    End If

    If Exists(CacheDirectory + "/itemsufcache.dat") = False Then
        frmWait.lblStatus = "Creating Suffix Cache .."
        frmWait.Refresh
        CreateSuffixCache
    Else
        If FileLen(CacheDirectory + "/itemsufcache.dat") <> 4845 Then
            frmWait.lblStatus = "Creating Suffix Cache .."
            frmWait.Refresh
            CreateSuffixCache
        End If
    End If
End Sub
