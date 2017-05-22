Attribute VB_Name = "modScript"
Option Explicit

Declare Function RunASMScript Lib "script.dll" Alias "RunScript" (Script As Any, FunctionTable As Any, Parameters As Any) As Long
Declare Function SysFreeString Lib "oleaut32.dll" (ByVal StringPointer As Long) As Long
Declare Function SysAllocStringByteLen Lib "oleaut32.dll" (ByVal St As String, ByVal Length As Long) As Long

Public FunctionTable(0 To 255) As Long
Public ScriptRunning As Boolean

Public Parameter(0 To 5) As Long
Public StringStack(0 To 1023) As Long
Public StringPointer As Long
Sub Boot_Player(ByVal Index As Long, ByVal Reason As String)
    BootPlayer Index, 0, StrConv(Reason, vbUnicode)
End Sub
Sub Ban_Player(ByVal Index As Long, ByVal NumDays As Long, ByVal Reason As String)
    BanPlayer Index, 0, NumDays, StrConv(Reason, vbUnicode), "Script Ban"
End Sub
Function Find_Player(ByVal Name As String) As Long
    Find_Player = FindPlayer(StrConv(Name, vbUnicode))
End Function

Function GetObjX(ByVal MapIndex As Long, ByVal ObjIndex As Long) As Long
    If MapIndex >= 1 And MapIndex <= 2000 And ObjIndex >= 0 And ObjIndex <= 49 Then
        With Map(MapIndex).Object(ObjIndex)
            GetObjX = .X
        End With
    End If
End Function
Function GetObjY(ByVal MapIndex As Long, ByVal ObjIndex As Long) As Long
    If MapIndex >= 1 And MapIndex <= 2000 And ObjIndex >= 0 And ObjIndex <= 49 Then
        With Map(MapIndex).Object(ObjIndex)
            GetObjY = .Y
        End With
    End If
End Function
Function GetObjNum(ByVal MapIndex As Long, ByVal ObjIndex As Long) As Long
    If MapIndex >= 1 And MapIndex <= 2000 And ObjIndex >= 0 And ObjIndex <= 49 Then
        With Map(MapIndex).Object(ObjIndex)
            GetObjNum = .Object
        End With
    End If
End Function
Function GetTileAtt(ByVal MapIndex As Long, ByVal X As Long, ByVal Y As Long) As Long
    If MapIndex >= 1 And MapIndex <= 2000 And X >= 0 And X <= 11 And Y >= 0 And Y <= 11 Then
        GetTileAtt = Map(MapIndex).Tile(X, Y).Att
    End If
End Function

Function DestroyObj(ByVal MapIndex As Long, ByVal ObjIndex As Long) As Long
    If MapIndex >= 1 And MapIndex <= 2000 And ObjIndex >= 0 And ObjIndex <= 49 Then
        With Map(MapIndex).Object(ObjIndex)
            .Object = 0
            SendToMap MapIndex, Chr$(15) + Chr$(ObjIndex) 'Erase Map Obj
        End With
    End If
End Function
Function GetObjVal(ByVal MapIndex As Long, ByVal ObjIndex As Long) As Long
    If MapIndex >= 1 And MapIndex <= 2000 And ObjIndex >= 0 And ObjIndex <= 49 Then
        With Map(MapIndex).Object(ObjIndex)
            GetObjVal = .Value
        End With
    End If
End Function


Function NewString(St As String) As Long
    Dim A As Long
    If StringPointer < 1024 Then
        A = SysAllocStringByteLen(St, Len(St))
        StringStack(StringPointer) = A
        StringPointer = StringPointer + 1
        NewString = A
    Else
        NewString = 0
    End If
End Function

Function AttackMonster(ByVal Index As Long, ByVal MonsterIndex As Long, ByVal Damage As Long) As Long
    Dim MapNum As Long, A As Long, B As Long, C As Long
    If Index >= 1 And Index <= MaxUsers And MonsterIndex >= 0 And MonsterIndex <= 5 Then
        If Player(Index).Mode = modePlaying Then
            MapNum = Player(Index).Map
            If Map(MapNum).Monster(MonsterIndex).Monster > 0 Then
                If Damage < 0 Then Damage = 0
                If Damage > 255 Then Damage = 255
                SendSocket Index, Chr$(44) + Chr$(0) + Chr$(MonsterIndex) + Chr$(Damage) 'Hit Monster
                SendToMapAllBut MapNum, Index, Chr$(42) + Chr$(Index)
                With Map(MapNum).Monster(MonsterIndex)
                    .Target = Index
                    If .HP > Damage Then
                        .HP = .HP - Damage
                    Else
                        'Monster Died
                        SendToMapAllBut MapNum, Index, Chr$(39) + Chr$(MonsterIndex) 'Monster Died
                        With Monster(.Monster)
                            A = CLng(.Strength) * 10 + CLng(.Armor) * 2 + CLng(.HP) * 2 + CLng(.Agility / 10)
                        End With
                        SendSocket Index, Chr$(51) + Chr$(MonsterIndex) + QuadChar(A) 'You killed monster
                        GainExp Index, A
                        B = Int(Rnd * 3)
                        C = Monster(.Monster).Object(B)
                        If C > 0 Then
                            NewMapObject MapNum, C, Monster(.Monster).Value(B), CLng(.X), CLng(.Y), False
                        End If
                        
                        Parameter(0) = Index
                        RunScript "MONSTERDIE" + CStr(.Monster)
                        .Monster = 0
                        
                        AttackMonster = True
                    End If
                End With
            End If
        End If
    End If
End Function
Function AttackPlayer(ByVal Index As Long, ByVal Target As Long, ByVal Damage As Long) As Long
    Dim A As Long
    
    If Index >= 1 And Index <= MaxUsers And Target >= 1 And Target <= MaxUsers Then
        If Player(Index).Mode = modePlaying And Player(Target).Mode = modePlaying Then
            If Damage < 0 Then Damage = 0
            If Damage > 255 Then Damage = 255
            With Player(Target)
                If .HP > Damage Then
                    .HP = .HP - Damage
                Else
                    .HP = 0
                End If
            End With
            SendSocket Target, Chr$(49) + Chr$(0) + Chr$(Index) + Chr$(Damage)
            SendSocket Index, Chr$(43) + Chr$(0) + Chr$(Target) + Chr$(Damage)
            SendToMapAllBut Player(Index).Map, Target, Chr$(42) + Chr$(Index)
            
            If Player(Target).HP = 0 Then
                'Player Died
                If Player(Target).Status <> 1 Then
                    Player(Target).Status = 1
                End If
                
                Parameter(0) = Index
                Parameter(1) = Target
                RunScript "KILLPLAYER"
        
                SendSocket Target, Chr$(52) + Chr$(Index) 'Player Killed You
                
                A = Player(Target).Experience
                PlayerDied Target
                A = A - Player(Target).Experience
                
                SendSocket Index, Chr$(45) + Chr$(Target) + QuadChar(A) 'You Killed Player
                GainExp Index, A
                SendAllButBut Index, Target, Chr$(61) + Chr$(Target) + Chr$(Index) 'Player was killed by player
                
                AttackPlayer = True
            End If
        End If
    End If
End Function
Function CanAttackMonster(ByVal Index As Long, ByVal MonsterIndex As Long) As Long
    Dim MapIndex As Long
    If Index >= 1 And Index <= MaxUsers And MonsterIndex >= 0 And MonsterIndex <= 5 Then
        If Player(Index).Mode = modePlaying Then
            MapIndex = Player(Index).Map
            If ExamineBit(Map(MapIndex).flags, 5) = False And Map(MapIndex).Monster(MonsterIndex).Monster > 0 Then
                CanAttackMonster = True
            End If
        End If
    End If
End Function
Function CanAttackPlayer(ByVal Player1 As Long, ByVal Player2 As Long) As Long
    Dim PKMap As Boolean
    If Player1 >= 1 And Player1 <= MaxUsers And Player2 >= 1 And Player2 <= MaxUsers Then
        With Player(Player1)
            If ExamineBit(Map(.Map).flags, 0) = False Then
                If .Mode = modePlaying And Player(Player2).Mode = modePlaying Then
                    If .Map = Player(Player2).Map Then
                        If .Access = 0 And Player(Player2).Access = 0 Then
                            PKMap = ExamineBit(Map(.Map).flags, 6)
                            If .Guild > 0 Or PKMap = True Then
                                If Player(Player2).Guild > 0 Or PKMap = True Then
                                    CanAttackPlayer = True
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End With
    End If
End Function
Function GetAbs(ByVal Value As Long) As Long
    GetAbs = Abs(Value)
End Function
Function GetMaxUsers() As Long
    GetMaxUsers = MaxUsers
End Function
Function GetMonsterType(ByVal MapIndex As Long, ByVal MonsterIndex As Long) As Long
    If MapIndex >= 1 And MapIndex <= 2000 And MonsterIndex >= 0 And MonsterIndex <= 5 Then
        GetMonsterType = Map(MapIndex).Monster(MonsterIndex).Monster
    End If
End Function
Function GetMonsterTarget(ByVal MapIndex As Long, ByVal MonsterIndex As Long) As Long
    If MapIndex >= 1 And MapIndex <= 2000 And MonsterIndex >= 0 And MonsterIndex <= 5 Then
        GetMonsterTarget = Map(MapIndex).Monster(MonsterIndex).Target
    End If
End Function
Sub NPCSay(ByVal MapIndex As Long, ByVal St As String)
    If MapIndex >= 1 And MapIndex <= 2000 Then
        If Map(MapIndex).NPC > 0 Then
            SendToMap MapIndex, Chr$(88) + Chr$(Map(MapIndex).NPC) + StrConv(St, vbUnicode)
        End If
    End If
End Sub
Sub NPCTell(ByVal Index As Long, ByVal St As String)
    Dim A As Long
    If Index >= 1 And Index <= MaxUsers Then
        With Player(Index)
            If .Mode = modePlaying Then
                A = Map(Player(Index).Map).NPC
                If A > 0 Then
                    SendSocket Index, Chr$(88) + Chr$(A) + StrConv(St, vbUnicode)
                End If
            End If
        End With
    End If
End Sub

Sub ResetPlayerFlag(ByVal FlagNum As Long)
    Dim A As Long
    If FlagNum >= 0 And FlagNum <= 127 Then
        A = World.PlayerFlagCounter(FlagNum)
        If A = 2147483647 Then
            A = 0
        Else
            A = A + 1
        End If
        World.PlayerFlagCounter(FlagNum) = A
    End If
End Sub
Sub ScriptTimer(ByVal Index As Long, ByVal Seconds As Long, ByVal Script As String)
Dim A As Long
    If Index >= 1 And Index <= MaxUsers Then
        If Seconds > 86400 Then Seconds = 86400
        If Seconds < 0 Then Seconds = 0
        With Player(Index)
            If .Mode = modePlaying Then
                For A = 1 To MaxPlayerTimers
                    If .ScriptTimer(A) = 0 Then
                        .Script(A) = StrConv(Script, vbUnicode)
                        .ScriptTimer(A) = GetTickCount + Seconds * 1000
                        Exit For
                        'Parameter(0) = Index
                        '.ScriptTimer = 0
                        'ScriptRunning = False
                        'RunScript .Script
                        'ScriptRunning = True
                    End If
                Next A
            End If
        End With
    End If
End Sub
Sub SetFlag(ByVal FlagNum As Long, ByVal Value As Long)
    If FlagNum >= 0 And FlagNum <= 255 Then
        World.Flag(FlagNum) = Value
    End If
End Sub
Function GetFlag(ByVal FlagNum As Long) As Long
    If FlagNum >= 0 And FlagNum <= 255 Then
        GetFlag = World.Flag(FlagNum)
    End If
End Function
Function SetMonsterTarget(ByVal MapIndex As Long, ByVal MonsterIndex As Long, ByVal Player As Long) As Long
    If MapIndex >= 1 And MapIndex <= 2000 And MonsterIndex >= 0 And MonsterIndex <= 5 And Player >= 1 And Player <= MaxUsers Then
        With Map(MapIndex).Monster(MonsterIndex)
            .Target = Player
            .Distance = 1
        End With
    End If
End Function

Function GetMonsterX(ByVal MapIndex As Long, ByVal MonsterIndex As Long) As Long
    If MapIndex >= 1 And MapIndex <= 2000 And MonsterIndex >= 0 And MonsterIndex <= 5 Then
        GetMonsterX = Map(MapIndex).Monster(MonsterIndex).X
    End If
End Function
Function GetMonsterY(ByVal MapIndex As Long, ByVal MonsterIndex As Long) As Long
    If MapIndex >= 1 And MapIndex <= 2000 And MonsterIndex >= 0 And MonsterIndex <= 5 Then
        GetMonsterY = Map(MapIndex).Monster(MonsterIndex).Y
    End If
End Function

Function GetSqr(ByVal Value As Long) As Long
    GetSqr = Sqr(Value)
End Function

Function GetTime() As Long
    GetTime = World.Hour
End Function

Function HasObj(ByVal Index As Long, ByVal ObjIndex As Long) As Long
    Dim A As Long, B As Long, C As Long
    If Index >= 1 And Index <= MaxUsers And ObjIndex >= 1 And ObjIndex <= 255 Then
        B = Object(ObjIndex).Type
        With Player(Index)
            For A = 1 To 20
                With .Inv(A)
                    If .Object = ObjIndex Then
                        If B = 6 Then
                            C = .Value
                            Exit For
                        Else
                            C = C + 1
                        End If
                    End If
                End With
            Next A
        End With
        HasObj = C
    End If
End Function
Function GetPlayerName(ByVal Index As Long) As Long
    If Index >= 1 And Index <= MaxUsers Then
        With Player(Index)
            GetPlayerName = NewString(.Name)
        End With
    End If
End Function
Function GetPlayerIP(ByVal Index As Long) As Long
    If Index >= 1 And Index <= MaxUsers Then
        With Player(Index)
            GetPlayerIP = NewString(.IP)
        End With
    End If
End Function

Function GetPlayerUser(ByVal Index As Long) As Long
    If Index >= 1 And Index <= MaxUsers Then
        With Player(Index)
            GetPlayerUser = NewString(.User)
        End With
    End If
End Function
Function GetPlayerDesc(ByVal Index As Long) As Long
    If Index >= 1 And Index <= MaxUsers Then
        With Player(Index)
            GetPlayerDesc = NewString(.desc)
        End With
    End If
End Function
Function GetGuildName(ByVal Index As Long) As Long
    If Index >= 1 And Index <= 255 Then
        With Guild(Index)
            GetGuildName = NewString(.Name)
        End With
    End If
End Function
Function GiveObj(ByVal Index As Long, ByVal ObjIndex As Long, ByVal Amount As Long) As Long
    Dim A As Long, B As Long, C As Long
    If Index >= 1 And Index <= MaxUsers And ObjIndex >= 1 And ObjIndex <= 255 Then
        With Player(Index)
            If .Mode = modePlaying Then
                B = Object(ObjIndex).Type
                If B = 6 Then
                    A = FindInvObject(Index, ObjIndex)
                    If A = 0 Then
                        A = FreeInvNum(Index)
                    Else
                        C = 1
                    End If
                Else
                    A = FreeInvNum(Index)
                End If
                If A > 0 Then
                    With .Inv(A)
                        .Object = ObjIndex
                        
                        Select Case B
                            Case 1, 2, 3, 4 'Weapon, Shield, Armor, Helmut
                                .Value = CLng(Object(ObjIndex).Data(0)) * 10
                            Case 6 'Money
                                If C = 1 Then
                                    .Value = .Value + Amount
                                Else
                                    .Value = Amount
                                End If
                            Case 8 'Ring
                                .Value = CLng(Object(ObjIndex).Data(1)) * 10
                            Case Else
                                .Value = 0
                        End Select
                        
                        SendSocket Index, Chr$(17) + Chr$(A) + Chr$(ObjIndex) + QuadChar(.Value) 'New Inv Obj
                    End With
                End If
            End If
        End With
    End If
End Function
Sub GlobalMessage(ByVal Message As String, ByVal MsgColor As Long)
    MsgColor = MsgColor Mod 16
    SendAll Chr$(56) + Chr$(MsgColor) + StrConv(Message, vbUnicode)
End Sub

Function IsPlaying(ByVal Index As Long) As Long
    If Index >= 1 And Index <= MaxUsers Then
        IsPlaying = Player(Index).Mode = modePlaying
    End If
End Function

Sub MapMessage(ByVal Index As Long, ByVal Message As String, ByVal MsgColor As Long)
    If Index >= 1 And Index <= 2000 Then
        MsgColor = MsgColor Mod 16
        SendToMap Index, Chr$(56) + Chr$(MsgColor) + StrConv(Message, vbUnicode)
    End If
End Sub
Sub MapMessageAllBut(ByVal MapIndex As Long, ByVal PlayerIndex As Long, ByVal Message As String, ByVal MsgColor As Long)
    If MapIndex >= 1 And MapIndex <= 2000 Then
        MsgColor = MsgColor Mod 16
        SendToMapAllBut MapIndex, PlayerIndex, Chr$(56) + Chr$(MsgColor) + StrConv(Message, vbUnicode)
    End If
End Sub
Function OpenDoor(ByVal MapNum As Long, ByVal X As Long, ByVal Y As Long) As Long
    Dim A As Long
    If MapNum >= 1 And MapNum <= 2000 And X >= 0 And X <= 11 And Y >= 0 And Y <= 11 Then
        A = FreeMapDoorNum(MapNum)
        If A >= 0 Then
            With Map(MapNum).Door(A)
                .Att = Map(MapNum).Tile(X, Y).Att
                .X = X
                .Y = Y
                .T = GetTickCount
            End With
            Map(MapNum).Tile(X, Y).Att = 0
            SendToMap MapNum, Chr$(36) + Chr$(A) + Chr$(X) + Chr$(Y)
            OpenDoor = 1
        End If
    End If
End Function
Sub PlayerMessage(ByVal Index As Long, ByVal Message As String, ByVal MsgColor As Long)
    If Index >= 1 And Index <= MaxUsers Then
        MsgColor = MsgColor Mod 16
        SendSocket Index, Chr$(56) + Chr$(MsgColor) + StrConv(Message, vbUnicode)
    End If
End Sub
Function RunScript0(ByVal Script As String) As Long
    ScriptRunning = False
    RunScript0 = RunScript(StrConv(Script, vbUnicode))
    ScriptRunning = True
End Function
Function RunScript1(ByVal Script As String, ByVal Parm1 As Long) As Long
    Parameter(0) = Parm1
    ScriptRunning = False
    RunScript1 = RunScript(StrConv(Script, vbUnicode))
    ScriptRunning = True
End Function
Function RunScript2(ByVal Script As String, ByVal Parm1 As Long, ByVal Parm2 As Long) As Long
    Parameter(0) = Parm1
    Parameter(1) = Parm2
    ScriptRunning = False
    RunScript2 = RunScript(StrConv(Script, vbUnicode))
    ScriptRunning = True
End Function
Function RunScript3(ByVal Script As String, ByVal Parm1 As Long, ByVal Parm2 As Long, ByVal Parm3 As Long) As Long
    Parameter(0) = Parm1
    Parameter(1) = Parm2
    Parameter(2) = Parm3
    ScriptRunning = False
    RunScript3 = RunScript(StrConv(Script, vbUnicode))
    ScriptRunning = True
End Function
Function RunScript4(ByVal Script As String, ByVal Parm1 As Long, ByVal Parm2 As Long, ByVal Parm3 As Long, ByVal Parm4 As Long) As Long
    Parameter(0) = Parm1
    Parameter(1) = Parm2
    Parameter(2) = Parm3
    Parameter(3) = Parm4
    ScriptRunning = False
    RunScript4 = RunScript(StrConv(Script, vbUnicode))
    ScriptRunning = True
End Function

Sub SetPlayerSprite(ByVal Index As Long, ByVal Sprite As Long)
    If Index >= 1 And Index <= MaxUsers And Sprite >= 0 And Sprite <= 255 Then
        With Player(Index)
            If Sprite = 0 Then
                If .Guild > 0 Then
                    If Guild(.Guild).Sprite > 0 Then
                        .Sprite = Guild(.Guild).Sprite
                    Else
                        .Sprite = .Class * 2 + .Gender - 1
                    End If
                Else
                    .Sprite = .Class * 2 + .Gender - 1
                End If
            Else
                .Sprite = Sprite
            End If
            SendAll Chr$(63) + Chr$(Index) + Chr$(.Sprite)
        End With
    End If
End Sub
Function SpawnMonster(ByVal MapIndex As Long, ByVal Monster As Long, ByVal X As Long, ByVal Y As Long) As Long
    Dim A As Long
    
    If MapIndex >= 1 And MapIndex <= 2000 And Monster >= 1 And Monster <= 255 And X >= 0 And X <= 11 And Y >= 0 And Y <= 11 Then
        With Map(MapIndex)
            For A = 0 To 5
                With .Monster(A)
                    If .Monster = 0 Then
                        SendToMapRaw MapIndex, SpawnMapMonster(MapIndex, A, Monster, X, Y)
                        SpawnMonster = 1
                        Exit Function
                    End If
                End With
            Next A
        End With
    End If
    
    SpawnMonster = 0
End Function
Function SpawnObject(ByVal MapIndex As Long, ByVal Object As Long, ByVal Value As Long, ByVal X As Long, ByVal Y As Long) As Long
    If MapIndex >= 1 And MapIndex <= 2000 And Object >= 1 And Object <= 255 And Value >= 0 And X >= 0 And X <= 11 And Y >= 0 And Y <= 11 Then
        SpawnObject = NewMapObject(MapIndex, Object, Value, X, Y, True)
    End If
End Function
Function Str(ByVal Value As Long) As Long
    Str = NewString(CStr(Value))
End Function
Function StrCat(ByVal String1 As String, ByVal String2 As String) As Long
    StrCat = NewString(StrConv(String1, vbUnicode) + StrConv(String2, vbUnicode))
End Function

Function StrCmp(ByVal String1 As String, ByVal String2 As String) As Long
    StrCmp = UCase$(StrConv(String1, vbUnicode)) = UCase$(StrConv(String2, vbUnicode))
End Function
Function GetInStr(ByVal String1 As String, ByVal String2 As String) As Long
    GetInStr = InStr(UCase$(StrConv(String1, vbUnicode)), UCase$(StrConv(String2, vbUnicode)))
End Function

Function StrFormat(ByVal String1 As String, ByVal String2 As String) As Long
    Dim St As String, St1 As String, St2 As String
    St1 = StrConv(String1, vbUnicode)
    St2 = StrConv(String2, vbUnicode)
    
    Dim A As Long, B As Byte
    For A = 1 To Len(St1)
        B = Asc(Mid$(St1, A, 1))
        If B = 42 Then
            St = St + St2
        Else
            St = St + Chr$(B)
        End If
    Next A
    
    StrFormat = NewString(St)
End Function
Function GetGuildHall(ByVal Index As Long) As Long
    If Index >= 1 And Index <= 255 Then
        GetGuildHall = Guild(Index).Hall
    End If
End Function
Function GetGuildBank(ByVal Index As Long) As Long
    If Index >= 1 And Index <= 255 Then
        GetGuildBank = Guild(Index).Bank
    End If
End Function
Function GetGuildSprite(ByVal Index As Long) As Long
    If Index >= 1 And Index <= 255 Then
        GetGuildSprite = Guild(Index).Sprite
    End If
End Function



Function GetGuildMemberCount(ByVal Index As Long) As Long
    If Index >= 1 And Index <= 255 Then
        GetGuildMemberCount = CountGuildMembers(Index)
    End If
End Function
Function GetMapPlayerCount(ByVal Index As Long)
    If Index >= 1 And Index <= 2000 Then
        GetMapPlayerCount = Map(Index).NumPlayers
    End If
End Function

Function GetPlayerAccess(ByVal Index As Long) As Long
    If Index >= 1 And Index <= MaxUsers Then
        GetPlayerAccess = Player(Index).Access
    End If
End Function
Function GetPlayerAgility(ByVal Index As Long) As Long
    If Index >= 1 And Index <= MaxUsers Then
        GetPlayerAgility = Player(Index).Agility
    End If
End Function

Function GetPlayerBank(ByVal Index As Long) As Long
    If Index >= 1 And Index <= MaxUsers Then
        GetPlayerBank = Player(Index).Bank
    End If
End Function
Function GetPlayerClass(ByVal Index As Long) As Long
    If Index >= 1 And Index <= MaxUsers Then
        GetPlayerClass = Player(Index).Class
    End If
End Function
Function GetPlayerEndurance(ByVal Index As Long) As Long
    If Index >= 1 And Index <= MaxUsers Then
        GetPlayerEndurance = Player(Index).Endurance
    End If
End Function
Function GetPlayerEnergy(ByVal Index As Long) As Long
    If Index >= 1 And Index <= MaxUsers Then
        GetPlayerEnergy = Player(Index).Energy
    End If
End Function
Function GetPlayerEquipped(ByVal Index As Long, ByVal EquippedIndex As Long) As Long
    If Index >= 1 And Index <= MaxUsers And EquippedIndex >= 1 And EquippedIndex <= 6 Then
        GetPlayerEquipped = Player(Index).EquippedObject(EquippedIndex)
    End If
End Function

Function GetPlayerExperience(ByVal Index As Long) As Long
    If Index >= 1 And Index <= MaxUsers Then
        GetPlayerExperience = Player(Index).Experience
    End If
End Function
Function GetPlayerGender(ByVal Index As Long) As Long
    If Index >= 1 And Index <= MaxUsers Then
        GetPlayerGender = Player(Index).Gender
    End If
End Function
Function GetPlayerGuild(ByVal Index As Long) As Long
    If Index >= 1 And Index <= MaxUsers Then
        GetPlayerGuild = Player(Index).Guild
    End If
End Function
Function GetPlayerHP(ByVal Index As Long) As Long
    If Index >= 1 And Index <= MaxUsers Then
        GetPlayerHP = Player(Index).HP
    End If
End Function
Function GetPlayerIntelligence(ByVal Index As Long) As Long
    If Index >= 1 And Index <= MaxUsers Then
        GetPlayerIntelligence = Player(Index).Intelligence
    End If
End Function
Function GetPlayerInvObject(ByVal Index As Long, ByVal InvIndex As Long) As Long
    If Index >= 1 And Index <= MaxUsers And InvIndex >= 1 And InvIndex <= 20 Then
        GetPlayerInvObject = Player(Index).Inv(InvIndex).Object
    End If
End Function
Function GetPlayerInvValue(ByVal Index As Long, ByVal InvIndex As Long) As Long
    If Index >= 1 And Index <= MaxUsers And InvIndex >= 1 And InvIndex <= 20 Then
        GetPlayerInvValue = Player(Index).Inv(InvIndex).Value
    End If
End Function
Function GetPlayerLevel(ByVal Index As Long) As Long
    If Index >= 1 And Index <= MaxUsers Then
        GetPlayerLevel = Player(Index).level
    End If
End Function
Function GetPlayerMana(ByVal Index As Long) As Long
    If Index >= 1 And Index <= MaxUsers Then
        GetPlayerMana = Player(Index).Mana
    End If
End Function
Function GetPlayerMap(ByVal Index As Long) As Long
    If Index >= 1 And Index <= MaxUsers Then
        With Player(Index)
            If .Mode = modePlaying Then
                GetPlayerMap = .Map
            End If
        End With
    End If
End Function

Function GetPlayerMaxEnergy(ByVal Index As Long) As Long
    If Index >= 1 And Index <= MaxUsers Then
        GetPlayerMaxEnergy = Player(Index).MaxEnergy
    End If
End Function
Function GetPlayerMaxHP(ByVal Index As Long) As Long
    If Index >= 1 And Index <= MaxUsers Then
        GetPlayerMaxHP = Player(Index).MaxHP
    End If
End Function
Function GetPlayerMaxMana(ByVal Index As Long) As Long
    If Index >= 1 And Index <= MaxUsers Then
        GetPlayerMaxMana = Player(Index).MaxMana
    End If
End Function
Function GetPlayerSprite(ByVal Index As Long) As Long
    If Index >= 1 And Index <= MaxUsers Then
        GetPlayerSprite = Player(Index).Sprite
    End If
End Function
Function GetPlayerFlag(ByVal Index As Long, ByVal FlagNum As Long) As Long
    If Index >= 1 And Index <= MaxUsers And FlagNum >= 0 And FlagNum <= 127 Then
        With Player(Index).Flag(FlagNum)
            If .ResetCounter = World.PlayerFlagCounter(FlagNum) Then
                GetPlayerFlag = .Value
            Else
                .Value = 0
                .ResetCounter = World.PlayerFlagCounter(FlagNum)
            End If
        End With
    End If
End Function
Sub SetPlayerFlag(ByVal Index As Long, ByVal FlagNum As Long, ByVal Value As Long)
    If Index >= 1 And Index <= MaxUsers And FlagNum >= 0 And FlagNum <= 127 Then
        With Player(Index).Flag(FlagNum)
            .Value = Value
            .ResetCounter = World.PlayerFlagCounter(FlagNum)
        End With
    End If
End Sub

Function GetPlayerStatus(ByVal Index As Long) As Long
    If Index >= 1 And Index <= MaxUsers Then
        GetPlayerStatus = Player(Index).Status
    End If
End Function
Function GetPlayerStrength(ByVal Index As Long) As Long
    If Index >= 1 And Index <= MaxUsers Then
        GetPlayerStrength = Player(Index).Strength
    End If
End Function
Function GetPlayerX(ByVal Index As Long) As Long
    If Index >= 1 And Index <= MaxUsers Then
        GetPlayerX = Player(Index).X
    End If
End Function
Function GetPlayerY(ByVal Index As Long) As Long
    If Index >= 1 And Index <= MaxUsers Then
        GetPlayerY = Player(Index).Y
    End If
End Function
Function GetValue(Value As Long) As Long
    GetValue = Value
End Function

Sub InitFunctionTable()
    FunctionTable(0) = GetValue(AddressOf DeleteString)
    FunctionTable(1) = GetValue(AddressOf StrCat)
    FunctionTable(2) = GetValue(AddressOf StrCmp)
    FunctionTable(3) = GetValue(AddressOf StrFormat)
    
    FunctionTable(4) = GetValue(AddressOf Random)
    
    FunctionTable(5) = GetValue(AddressOf GetPlayerAccess)
    FunctionTable(6) = GetValue(AddressOf GetPlayerMap)
    FunctionTable(7) = GetValue(AddressOf GetPlayerX)
    FunctionTable(8) = GetValue(AddressOf GetPlayerY)
    FunctionTable(9) = GetValue(AddressOf GetPlayerSprite)
    FunctionTable(10) = GetValue(AddressOf GetPlayerClass)
    FunctionTable(11) = GetValue(AddressOf GetPlayerGender)
    FunctionTable(12) = GetValue(AddressOf GetPlayerHP)
    FunctionTable(13) = GetValue(AddressOf GetPlayerEnergy)
    FunctionTable(14) = GetValue(AddressOf GetPlayerMana)
    FunctionTable(15) = GetValue(AddressOf GetPlayerMaxHP)
    FunctionTable(16) = GetValue(AddressOf GetPlayerMaxEnergy)
    FunctionTable(17) = GetValue(AddressOf GetPlayerMaxMana)
    FunctionTable(18) = GetValue(AddressOf GetPlayerStrength)
    FunctionTable(19) = GetValue(AddressOf GetPlayerEndurance)
    FunctionTable(20) = GetValue(AddressOf GetPlayerIntelligence)
    FunctionTable(21) = GetValue(AddressOf GetPlayerAgility)
    FunctionTable(22) = GetValue(AddressOf GetPlayerBank)
    FunctionTable(23) = GetValue(AddressOf GetPlayerExperience)
    FunctionTable(24) = GetValue(AddressOf GetPlayerLevel)
    FunctionTable(25) = GetValue(AddressOf GetPlayerStatus)
    FunctionTable(26) = GetValue(AddressOf GetPlayerGuild)
    
    FunctionTable(27) = GetValue(AddressOf GetPlayerInvObject)
    FunctionTable(28) = GetValue(AddressOf GetPlayerInvValue)
    FunctionTable(29) = GetValue(AddressOf GetPlayerEquipped)
    
    FunctionTable(30) = GetValue(AddressOf GetPlayerName)
    FunctionTable(31) = GetValue(AddressOf GetPlayerUser)
    FunctionTable(32) = GetValue(AddressOf GetPlayerDesc)
       
    FunctionTable(33) = GetValue(AddressOf SetPlayerHP)
    FunctionTable(34) = GetValue(AddressOf SetPlayerEnergy)
    FunctionTable(35) = GetValue(AddressOf SetPlayerMana)
    
       
    FunctionTable(36) = GetValue(AddressOf PlayerMessage)
    FunctionTable(37) = GetValue(AddressOf PlayerWarp)
    
    FunctionTable(38) = GetValue(AddressOf MapMessage)
    
    FunctionTable(39) = GetValue(AddressOf GlobalMessage)
    
    FunctionTable(40) = GetValue(AddressOf GetGuildHall)
    FunctionTable(41) = GetValue(AddressOf GetGuildBank)
    FunctionTable(42) = GetValue(AddressOf GetGuildMemberCount)
    FunctionTable(43) = GetValue(AddressOf GetGuildName)
    
    FunctionTable(44) = GetValue(AddressOf GetMapPlayerCount)
    
    FunctionTable(45) = GetValue(AddressOf MapMessageAllBut)
    
    FunctionTable(46) = GetValue(AddressOf HasObj)
    FunctionTable(47) = GetValue(AddressOf TakeObj)
    FunctionTable(48) = GetValue(AddressOf GiveObj)
    
    FunctionTable(49) = GetValue(AddressOf GetTime)
    
    FunctionTable(50) = GetValue(AddressOf GetMaxUsers)
    
    FunctionTable(51) = GetValue(AddressOf RunScript0)
    FunctionTable(52) = GetValue(AddressOf RunScript1)
    FunctionTable(53) = GetValue(AddressOf RunScript2)
    FunctionTable(54) = GetValue(AddressOf RunScript3)
    
    FunctionTable(55) = GetValue(AddressOf OpenDoor)
    
    FunctionTable(56) = GetValue(AddressOf Str)
    
    FunctionTable(57) = GetValue(AddressOf SetPlayerSprite)
    
    FunctionTable(58) = GetValue(AddressOf GetAbs)
    FunctionTable(59) = GetValue(AddressOf GetSqr)
    
    FunctionTable(60) = GetValue(AddressOf CanAttackPlayer)
    FunctionTable(61) = GetValue(AddressOf IsPlaying)
    
    FunctionTable(62) = GetValue(AddressOf AttackPlayer)
    FunctionTable(63) = GetValue(AddressOf AttackMonster)
    FunctionTable(64) = GetValue(AddressOf CanAttackMonster)
    
    FunctionTable(65) = GetValue(AddressOf GetMonsterType)
    FunctionTable(66) = GetValue(AddressOf GetMonsterX)
    FunctionTable(67) = GetValue(AddressOf GetMonsterY)
    FunctionTable(68) = GetValue(AddressOf GetMonsterTarget)
    FunctionTable(69) = GetValue(AddressOf SetMonsterTarget)
    
    FunctionTable(70) = GetValue(AddressOf GetInStr)
    FunctionTable(71) = GetValue(AddressOf SpawnObject)
    
    FunctionTable(72) = GetValue(AddressOf NPCSay)
    FunctionTable(73) = GetValue(AddressOf NPCTell)
    
    FunctionTable(74) = GetValue(AddressOf GetGuildSprite)
    
    FunctionTable(75) = GetValue(AddressOf ScriptTimer)
    
    FunctionTable(76) = GetValue(AddressOf SetPlayerGuild)
    
    FunctionTable(77) = GetValue(AddressOf GetFlag)
    FunctionTable(78) = GetValue(AddressOf SetFlag)
   
    FunctionTable(79) = GetValue(AddressOf GetPlayerFlag)
    FunctionTable(80) = GetValue(AddressOf SetPlayerFlag)
    
    FunctionTable(81) = GetValue(AddressOf ResetPlayerFlag)
    
    FunctionTable(82) = GetValue(AddressOf SetPlayerAccess)
    
    FunctionTable(83) = GetValue(AddressOf GetObjX)
    FunctionTable(84) = GetValue(AddressOf GetObjY)
    FunctionTable(85) = GetValue(AddressOf GetObjNum)
    FunctionTable(86) = GetValue(AddressOf GetObjVal)
    FunctionTable(87) = GetValue(AddressOf DestroyObj)
    FunctionTable(88) = GetValue(AddressOf Boot_Player)
    FunctionTable(89) = GetValue(AddressOf Ban_Player)
    FunctionTable(90) = GetValue(AddressOf SetPlayerName)
    FunctionTable(91) = GetValue(AddressOf SetPlayerBank)
    FunctionTable(92) = GetValue(AddressOf SetGuildBank)
    FunctionTable(93) = GetValue(AddressOf Find_Player)
    FunctionTable(94) = GetValue(AddressOf StrVal)
    FunctionTable(95) = GetValue(AddressOf GetTileAtt)
    FunctionTable(96) = GetValue(AddressOf GetPlayerIP)
    FunctionTable(97) = GetValue(AddressOf RunScript4)
    FunctionTable(98) = GetValue(AddressOf SpawnMonster)
    FunctionTable(99) = GetValue(AddressOf SetPlayerStatus)
    FunctionTable(100) = GetValue(AddressOf GivePlayerExp)
    'FunctionTable(101) = GetValue(AddressOf SetAttackingMod)
    'FunctionTable(102) = GetValue(AddressOf SetDefendingMod)
    FunctionTable(103) = GetValue(AddressOf GetObjectName)
    FunctionTable(104) = GetValue(AddressOf GetObjectData)
    FunctionTable(105) = GetValue(AddressOf GetObjectType)
    FunctionTable(106) = GetValue(AddressOf DisplayObjDur)
    FunctionTable(107) = GetValue(AddressOf SetInvObjectVal)
    FunctionTable(108) = GetValue(AddressOf PlayCustomWav)
    FunctionTable(109) = GetValue(AddressOf GetPlayerArmor)
    FunctionTable(110) = GetValue(AddressOf CreateTileEffect)
    FunctionTable(111) = GetValue(AddressOf CreateCharacterEffect)
    FunctionTable(112) = GetValue(AddressOf CreateMonsterEffect)
    FunctionTable(113) = GetValue(AddressOf CreatePlayerEffect)
End Sub
Function Random(ByVal Max As Long) As Long
    Random = Int(Rnd * Max)
End Function

Function RunScript(Name As String) As Long
    If ScriptRunning = False Then
        ScriptRS.Seek "=", Name
        If ScriptRS.NoMatch = False Then
            Dim A As Long, StringCount As Long
            StringCount = StringPointer
            Dim MCode() As Byte
            MCode() = StrConv(ScriptRS!Data, vbFromUnicode)
            ScriptRunning = True
            RunScript = RunASMScript(MCode(0), FunctionTable(0), Parameter(0))
            ScriptRunning = False
            For A = StringCount To StringPointer - 1
                SysFreeString StringStack(A)
            Next A
            StringPointer = StringCount
        End If
    End If
End Function
Sub SetPlayerEnergy(ByVal Index As Long, ByVal Energy As Long)
    If Index >= 1 And Index <= MaxUsers Then
        With Player(Index)
            If .Mode = modePlaying Then
                If Energy > 255 Then Energy = 255
                If Energy < 0 Then Energy = 0
                .Energy = Energy
                SendSocket Index, Chr$(47) + Chr$(Energy)
            End If
        End With
    End If
End Sub
Sub SetPlayerName(ByVal Index As Long, ByVal Name As String)
    If Index >= 1 And Index <= MaxUsers Then
        With Player(Index)
            If .Mode = modePlaying Then
                .Name = StrConv(Name, vbUnicode)
                SendAll Chr$(64) + Chr$(Index) + .Name
            End If
        End With
    End If
End Sub

Sub SetPlayerMana(ByVal Index As Long, ByVal Mana As Long)
    If Index >= 1 And Index <= MaxUsers Then
        With Player(Index)
            If .Mode = modePlaying Then
                If Mana > 255 Then Mana = 255
                If Mana < 0 Then Mana = 0
                .Mana = Mana
                SendSocket Index, Chr$(48) + Chr$(Mana)
            End If
        End With
    End If
End Sub


Sub SetPlayerHP(ByVal Index As Long, ByVal HP As Long)
    If Index >= 1 And Index <= MaxUsers Then
        With Player(Index)
            If .Mode = modePlaying Then
                If HP > 255 Then HP = 255
                If HP < 1 Then HP = 1
                .HP = HP
                SendSocket Index, Chr$(46) + Chr$(HP)
            End If
        End With
    End If
End Sub
Sub SetPlayerBank(ByVal Index As Long, ByVal Bank As Long)
    If Index >= 1 And Index <= MaxUsers Then
        With Player(Index)
            If .Mode = modePlaying Then
                .Bank = Bank
            End If
        End With
    End If
End Sub

Sub SetPlayerStatus(ByVal Index As Long, ByVal Status As Long)
If Index >= 1 And Index <= MaxUsers And Status >= 0 And Status <= 100 Then
    With Player(Index)
        If .Mode = modePlaying Then
            .Status = Status
            SendAll Chr$(91) + Chr$(Index) + Chr$(Status)
        End If
    End With
End If
End Sub
Sub SetPlayerGuild(ByVal Index As Long, ByVal GuildIndex As Long)
    If Index >= 1 And Index <= MaxUsers And GuildIndex >= 0 And GuildIndex <= 255 Then
        With Player(Index)
            If GuildIndex > 0 Then
                If .Guild <> GuildIndex Then
                    .Guild = GuildIndex
                    .GuildRank = 1
                    If Guild(GuildIndex).Sprite > 0 Then
                        .Sprite = Guild(GuildIndex).Sprite
                        SendAll Chr$(63) + Chr$(Index) + Chr$(.Sprite)
                    End If
                    SendSocket Index, Chr$(72) + Chr$(GuildIndex) 'Change guild
                    SendAllBut Index, Chr$(73) + Chr$(Index) + Chr$(GuildIndex) 'Player changed guild
                End If
            Else
                If .Guild > 0 Then
                    If Guild(.Guild).Sprite > 0 Then
                        .Sprite = .Class * 2 + .Gender - 1
                        SendAll Chr$(63) + Chr$(Index) + Chr$(.Sprite)
                    End If
                    .Guild = 0
                    SendSocket Index, Chr$(72) + Chr$(0) 'Change guild
                    SendAllBut Index, Chr$(73) + Chr$(Index) + Chr$(0) 'Player changed guild
                End If
            End If
        End With
    End If
End Sub

Sub SetGuildBank(ByVal Index As Long, ByVal Bank As Long)
    If Index >= 1 And Index <= 255 Then
        With Guild(Index)
            If .Name <> "" Then
                .Bank = Bank
                GuildRS.Bookmark = .Bookmark
                GuildRS.Edit
                GuildRS!Bank = Bank
                GuildRS.Update
            End If
        End With
    End If
End Sub
Sub PlayerWarp(ByVal Index As Long, ByVal Map As Long, ByVal X As Long, ByVal Y As Long)
    If Index >= 1 And Index <= MaxUsers And Map >= 1 And Map <= 2000 And X >= 0 And X <= 11 And Y >= 0 And Y <= 11 Then
        With Player(Index)
            If .Mode = modePlaying Then
                Partmap (Index)
                .Map = Map
                .X = X
                .Y = Y
                JoinMap (Index)
            End If
        End With
    End If
End Sub
Sub DeleteString(ByVal StPointer As Long)
End Sub
Function StrVal(ByVal String1 As String) As Long
    On Error Resume Next
    StrVal = Int(Val(StrConv(String1, vbUnicode)))
    On Error GoTo 0
End Function

Function TakeObj(ByVal Index As Long, ByVal ObjIndex As Long, ByVal Amount As Long) As Long
    Dim A As Long
    If Index >= 1 And Index <= MaxUsers And ObjIndex >= 1 And ObjIndex <= 255 Then
        With Player(Index)
            If .Mode = modePlaying Then
                A = FindInvObject(Index, ObjIndex)
                If A > 0 Then
                    With .Inv(A)
                        If Object(ObjIndex).Type = 6 Then
                            If .Value >= Amount Then
                                .Value = .Value - Amount
                                If .Value = 0 Then
                                    .Object = 0
                                    SendSocket Index, Chr$(18) + Chr$(A)
                                Else
                                    SendSocket Index, Chr$(17) + Chr$(A) + Chr$(ObjIndex) + QuadChar(.Value) 'New Inv Obj
                                End If
                                TakeObj = Amount
                            End If
                        Else
                            .Object = 0
                            TakeObj = 1
                            SendSocket Index, Chr$(18) + Chr$(A)
                        End If
                    End With
                End If
            End If
        End With
    End If
End Function

Sub SetPlayerAccess(ByVal Index As Long, ByVal Access As Long)
    'If Index >= 1 And Index <= MaxUsers Then
    '    With Player(Index)
    '        If .Mode = modePlaying And .Access < 11 Then
    '            If Access < 0 Then
    '                .Access = 0
    '            ElseIf Access > 10 Then
    '                .Access = 10
    '            Else
    '                .Access = Access
    '            End If
    '
    '            SendSocket Index, Chr$(65) + Chr$(Access)
    '            If Access > 0 Then
    '                SendAllBut Index, Chr$(91) + Chr$(Index) + Chr$(3)
    '                .Status = 3
    '            Else
    '                SendAllBut Index, Chr$(91) + Chr$(Index) + Chr$(0)
    '                .Status = 0
    '            End If
    '        End If
    '    End With
    'End If
End Sub

Sub GivePlayerExp(ByVal Index As Long, ByVal Experience As Long)
If Index >= 1 And Index <= MaxUsers And Experience <= 50000 Then
    GainExp Index, Experience
    SendSocket Index, Chr$(60) + QuadChar(Player(Index).Experience)
End If
End Sub
Function GetObjectName(ByVal ObjectNum As Long) As Long
If ObjectNum >= 1 And ObjectNum <= 255 Then
    GetObjectName = NewString(Object(ObjectNum).Name)
End If
End Function

Function GetObjectData(ByVal ObjectNum As Long, ByVal DataNum As Long) As Long
If ObjectNum >= 1 And ObjectNum <= 255 Then
    If DataNum >= 0 And DataNum <= 3 Then
        GetObjectData = Object(ObjectNum).Data(DataNum)
    End If
End If
End Function

Function GetObjectType(ByVal ObjectNum As Long) As Long
If ObjectNum >= 1 And ObjectNum <= 255 Then
    GetObjectType = Object(ObjectNum).Type
End If
End Function

Sub DisplayObjDur(ByVal Index As Long, ByVal ObjectNum As Long)
Dim Percent As Single, St As String, MsgColor As Long
Dim Display As Boolean
    Select Case Object(Player(Index).Inv(ObjectNum).Object).Type
    Case 1, 2, 3, 4
        Display = True
    Case Else
        Display = False
    End Select
If Display = True Then
    'Percent = Int((Player(Index).Inv(ObjectNum).Value * 5) / Object(Player(Index).Inv(ObjectNum).Object).Data(0))
    Percent = Player(Index).Inv(ObjectNum).Value / (Object(Player(Index).Inv(ObjectNum).Object).Data(0) * 10)
    Percent = Int(Percent * 100)
    If Percent > 100 Then Percent = 100
    If Percent <= 5 Then
        St = "Your " + Object(Player(Index).Inv(ObjectNum).Object).Name + " is about to break!"
        MsgColor = 2
    Else
        St = "Your " + Object(Player(Index).Inv(ObjectNum).Object).Name + " is at " + CStr(Percent) + "% durability."
        MsgColor = 14
    End If
    SendSocket Index, Chr$(56) + Chr$(MsgColor Mod 16) + St
Else
    St = "This is an invalid object or no object."
    MsgColor = 2
    SendSocket Index, Chr$(56) + Chr$(MsgColor Mod 16) + St
End If
End Sub

Sub SetInvObjectVal(ByVal Index As Long, ByVal InvSlot As Long, ByVal NewVal As Long)
If Index >= 1 And Index <= MaxUsers And InvSlot >= 1 And InvSlot <= 20 Then
    Player(Index).Inv(InvSlot).Value = NewVal
End If
End Sub

Sub PlayCustomWav(ByVal Index As Long, ByVal SoundNum As Long)
If Index >= 1 And Index <= MaxUsers And SoundNum <= 255 And SoundNum >= 1 Then
    SendSocket Index, Chr$(96) + Chr$(SoundNum)
End If
End Sub

Function GetPlayerArmor(ByVal Index As Long, ByVal Damage As Long) As Long
If Index >= 1 And Index <= MaxUsers And Damage <= 255 And Damage >= 1 Then
    GetPlayerArmor = PlayerArmor(Index, Damage)
End If
End Function
Sub CreateTileEffect(ByVal Index As Long, ByVal X As Long, ByVal Y As Long, ByVal Sprite As Long, ByVal Speed As Long, ByVal TotalFrames As Long, ByVal LoopCount As Long, ByVal EndSound As Long)
    Dim A As Integer
    If Index >= 1 And Index <= 2000 Then
        If X < 0 Then X = 0
        If X > 11 Then X = 11
        If Y < 0 Then Y = 0
        If Y > 11 Then Y = 11
        SendToMap Index, Chr$(99) + Chr$(1) + Chr$(X) + Chr$(Y) + Chr$(Sprite) + DoubleChar(Speed) + Chr$(TotalFrames) + Chr$(LoopCount) + Chr$(EndSound)
    End If
End Sub
Sub CreateCharacterEffect(ByVal Index As Long, ByVal Player As Long, ByVal Sprite As Long, ByVal Speed As Long, ByVal TotalFrames As Long, ByVal LoopCount As Long, ByVal EndSound As Long)
    If Index >= 1 And Index <= 2000 Then
        SendToMap Index, Chr$(99) + Chr$(2) + Chr$(Player) + Chr$(Sprite) + DoubleChar(Speed) + Chr$(TotalFrames) + Chr$(LoopCount) + Chr$(EndSound)
    End If
End Sub
Sub CreateMonsterEffect(ByVal Index As Long, ByVal Player As Long, ByVal Monster As Long, ByVal Sprite As Long, ByVal Speed As Long, ByVal TotalFrames As Long, ByVal EndSound As Long)
    If Index >= 1 And Index <= 2000 Then
        SendToMap Index, Chr$(99) + Chr$(3) + Chr$(Player) + Chr$(Monster) + Chr$(Sprite) + DoubleChar(Speed) + Chr$(TotalFrames) + Chr$(EndSound)
    End If
End Sub
Sub CreatePlayerEffect(ByVal Index As Long, ByVal SourcePlayer As Long, ByVal TargetPlayer As Long, ByVal Sprite As Long, ByVal Speed As Long, ByVal TotalFrames As Long, ByVal EndSound As Long)
    If Index >= 1 And Index <= 2000 Then
        SendToMap Index, Chr$(99) + Chr$(4) + Chr$(SourcePlayer) + Chr$(TargetPlayer) + Chr$(Sprite) + DoubleChar(Speed) + Chr$(TotalFrames) + Chr$(EndSound)
    End If
End Sub

