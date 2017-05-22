EXTERNAL FUNCTION StrCat(St1 AS STRING, St2 AS STRING) AS STRING = 1
EXTERNAL FUNCTION StrCmp(St1 AS STRING, St2 AS STRING) AS LONG = 2
EXTERNAL FUNCTION StrFormat(FmtString AS STRING, RplString AS STRING) AS STRING = 3
EXTERNAL FUNCTION Str(Value AS LONG) AS STRING = 56
EXTERNAL FUNCTION InStr(St1 AS STRING, St2 AS STRING) AS LONG = 70
EXTERNAL FUNCTION Val(St1 AS STRING) AS LONG = 94

EXTERNAL FUNCTION Random(NumChoices AS LONG) AS LONG = 4
EXTERNAL FUNCTION Abs(Value AS LONG) AS LONG = 58
EXTERNAL FUNCTION Sqr(Value AS LONG) AS LONG = 59

EXTERNAL FUNCTION FindPlayer(Name as STRING) AS LONG = 93

EXTERNAL FUNCTION GetPlayerIP(Player AS LONG) AS STRING = 96
EXTERNAL FUNCTION GetPlayerAccess(Player AS LONG) AS LONG = 5
EXTERNAL FUNCTION GetPlayerMap(Player AS LONG) AS LONG = 6
EXTERNAL FUNCTION GetPlayerX(Player AS LONG) AS LONG = 7
EXTERNAL FUNCTION GetPlayerY(Player AS LONG) AS LONG = 8
EXTERNAL FUNCTION GetPlayerSprite(Player AS LONG) AS LONG = 9
EXTERNAL FUNCTION GetPlayerClass(Player AS LONG) AS LONG = 10
EXTERNAL FUNCTION GetPlayerGender(Player AS LONG) AS LONG = 11
EXTERNAL FUNCTION GetPlayerHP(Player AS LONG) AS LONG = 12
EXTERNAL FUNCTION GetPlayerEnergy(Player AS LONG) AS LONG = 13
EXTERNAL FUNCTION GetPlayerMana(Player AS LONG) AS LONG = 14
EXTERNAL FUNCTION GetPlayerMaxHP(Player AS LONG) AS LONG = 15
EXTERNAL FUNCTION GetPlayerMaxEnergy(Player AS LONG) AS LONG = 16
EXTERNAL FUNCTION GetPlayerMaxMana(Player AS LONG) AS LONG = 17
EXTERNAL FUNCTION GetPlayerStrength(Player AS LONG) AS LONG = 18
EXTERNAL FUNCTION GetPlayerEndurance(Player AS LONG) AS LONG = 19
EXTERNAL FUNCTION GetPlayerIntelligence(Player AS LONG) AS LONG = 20
EXTERNAL FUNCTION GetPlayerAgility(Player AS LONG) AS LONG = 21
EXTERNAL FUNCTION GetPlayerBank(Player AS LONG) AS LONG = 22
EXTERNAL FUNCTION GetPlayerExperience(Player AS LONG) AS LONG = 23
EXTERNAL FUNCTION GetPlayerLevel(Player AS LONG) AS LONG = 24
EXTERNAL FUNCTION GetPlayerStatus(Player AS LONG) AS LONG = 25
EXTERNAL FUNCTION SetPlayerStatus(Player AS LONG, Status AS LONG) AS LONG = 99
EXTERNAL FUNCTION GetPlayerGuild(Player AS LONG) AS LONG = 26
EXTERNAL FUNCTION GetPlayerInvObject(Player AS LONG, InvIndex AS LONG) AS LONG = 27
EXTERNAL FUNCTION GetPlayerInvValue(Player AS LONG, InvIndex AS LONG) AS LONG = 28
EXTERNAL FUNCTION GetPlayerEquipped(Player AS LONG, EqIndex AS LONG) AS LONG = 29

EXTERNAL SUB GivePlayerExp(Player AS LONG, Experience as LONG) = 100
EXTERNAL FUNCTION GetPlayerArmor(Player AS LONG, Damage AS LONG) AS LONG = 109

EXTERNAL FUNCTION GetPlayerName(Player AS LONG) AS STRING = 30
EXTERNAL FUNCTION GetPlayerUser(Player AS LONG) AS STRING = 31
EXTERNAL FUNCTION GetPlayerDesc(Player AS LONG) AS STRING = 32

EXTERNAL FUNCTION HasObj(Player AS LONG, Object AS LONG) AS LONG = 46
EXTERNAL FUNCTION TakeObj(Player AS LONG, Object AS LONG, Amount AS LONG) AS LONG = 47
EXTERNAL FUNCTION GiveObj(Player AS LONG, Object AS LONG, Amount AS LONG) AS LONG = 48

EXTERNAL FUNCTION IsPlaying(Player AS LONG) AS LONG = 61

EXTERNAL FUNCTION CanAttackPlayer(Attacker AS LONG, Attackee AS LONG) AS LONG = 60
EXTERNAL FUNCTION CanAttackMonster(Attacker AS LONG, Monster AS LONG) AS LONG = 64
EXTERNAL FUNCTION AttackPlayer(Attacker AS LONG, Attackee AS LONG, Damage AS LONG) AS LONG = 62
EXTERNAL FUNCTION AttackMonster(Attacker AS LONG, Monster AS LONG, Damage AS LONG) AS LONG = 63

EXTERNAL SUB SetPlayerHP(Player AS LONG, HP AS LONG) = 33
EXTERNAL SUB SetPlayerEnergy(Player AS LONG, Energy AS LONG) = 34
EXTERNAL SUB SetPlayerMana(Player AS LONG, Mana AS LONG) = 35

EXTERNAL SUB SetPlayerAccess(Player AS LONG, Access AS LONG) = 82

EXTERNAL SUB SetPlayerSprite(Player AS LONG, Sprite AS LONG) = 57
EXTERNAL SUB SetPlayerGuild(Player AS LONG, Guild AS LONG) = 76

EXTERNAL SUB SetPlayerName(Player AS LONG, Name as STRING) = 90
EXTERNAL SUB SetPlayerBank(Player AS LONG, Bank as LONG) = 91

EXTERNAL SUB BootPlayer(Player AS LONG, Reason AS STRING)= 88
EXTERNAL SUB BanPlayer(Player AS LONG, Days as LONG, Reason AS STRING)= 89

EXTERNAL SUB PlayerMessage(Player AS LONG, Message AS STRING, MsgColor AS LONG) = 36
EXTERNAL SUB PlayerWarp(Player AS LONG, Map AS LONG, X AS LONG, Y AS LONG) = 37

EXTERNAL FUNCTION GetPlayerFlag(Player AS LONG, FlagNum AS LONG) AS LONG = 79
EXTERNAL SUB SetPlayerFlag(Player AS LONG, FlagNum AS LONG, Value AS LONG) = 80
EXTERNAL SUB ResetPlayerFlag(FlagNum AS LONG) = 81

EXTERNAL FUNCTION GetGuildHall(Guild AS LONG) AS LONG = 40
EXTERNAL FUNCTION GetGuildBank(Guild AS LONG) AS LONG = 41
EXTERNAL FUNCTION GetGuildMemberCount(Guild AS LONG) AS LONG = 42
EXTERNAL FUNCTION GetGuildName(Guild AS LONG) AS STRING = 43
EXTERNAL FUNCTION GetGuildSprite(Guild AS LONG) AS LONG = 74

EXTERNAL SUB SetGuildBank(Player AS LONG, Bank as LONG) = 92

EXTERNAL FUNCTION GetMapPlayerCount(Map AS LONG) AS LONG = 44

EXTERNAL FUNCTION OpenDoor(Map AS LONG, X AS LONG, Y AS LONG) AS LONG = 55
EXTERNAL SUB MapMessageAllBut(Map AS LONG, Player AS LONG, Message AS STRING, MsgColor AS LONG) = 45
EXTERNAL SUB MapMessage(Map AS LONG, Message AS STRING, MsgColor AS LONG) = 38
EXTERNAL SUB NPCSay(Map AS LONG, Message AS STRING) = 72
EXTERNAL SUB NPCTell(Player AS LONG, Message AS STRING) = 73

EXTERNAL FUNCTION SpawnMonster(Map As Long, Monster As Long, X As Long, Y As Long,Frozen as Long) AS LONG = 98

EXTERNAL FUNCTION SpawnObject(Map AS LONG, Object AS LONG, Value AS LONG, X AS LONG, Y AS LONG) AS LONG = 71
EXTERNAL SUB DestroyObject(Map AS LONG, Object AS LONG) = 87
EXTERNAL FUNCTION GetObjX(Map AS LONG, Object AS LONG) AS LONG = 83
EXTERNAL FUNCTION GetObjY(Map AS LONG, Object AS LONG) AS LONG = 84
EXTERNAL FUNCTION GetObjNum(Map AS LONG, Object AS LONG) AS LONG = 85
EXTERNAL FUNCTION GetObjVal(Map AS LONG, Object AS LONG) AS LONG = 86

EXTERNAL FUNCTION GetObjectName(ObjectNum AS LONG) AS STRING = 103
EXTERNAL FUNCTION GetObjectData(ObjectNum AS LONG, Data AS LONG) AS LONG = 104
EXTERNAL FUNCTION GetObjectType(ObjectNum AS LONG) AS LONG = 105
EXTERNAL SUB DisplayObjDur(Player AS LONG, InvSlot AS LONG) = 106
EXTERNAL SUB SetInvObjectVal(Player AS LONG, InvSlot AS LONG, NewVal AS LONG) = 107

EXTERNAL FUNCTION GetMonsterType(Map AS LONG, Monster AS LONG) AS LONG = 65
EXTERNAL FUNCTION GetMonsterX(Map AS LONG, Monster AS LONG) AS LONG = 66
EXTERNAL FUNCTION GetMonsterY(Map AS LONG, Monster AS LONG) AS LONG = 67
EXTERNAL FUNCTION GetMonsterTarget(Map AS LONG, Monster AS LONG) AS LONG = 68
EXTERNAL SUB SetMonsterTarget(Map AS LONG, Monster AS LONG, Player AS LONG) = 69

EXTERNAL FUNCTION GetFlag(FlagNum AS LONG) AS LONG = 77
EXTERNAL SUB SetFlag(FlagNum AS LONG, Value AS LONG) = 78

EXTERNAL SUB GlobalMessage(Message AS STRING, MsgColor AS LONG) = 39
EXTERNAL FUNCTION GetTime() AS LONG = 49
EXTERNAL FUNCTION GetMaxUsers() AS LONG = 50

EXTERNAL FUNCTION RunScript0(Script AS STRING) AS LONG = 51
EXTERNAL FUNCTION RunScript1(Script AS STRING, Parm1 AS LONG) AS LONG = 52
EXTERNAL FUNCTION RunScript2(Script AS STRING, Parm1 AS LONG, Parm2 AS LONG) AS LONG = 53
EXTERNAL FUNCTION RunScript3(Script AS STRING, Parm1 AS LONG, Parm2 AS LONG, Parm3 AS LONG) AS LONG = 54
EXTERNAL FUNCTION RunScript4(Script AS STRING, Parm1 AS LONG, Parm2 AS LONG, Parm3 AS LONG, Parm4 as LONG) AS LONG = 97

EXTERNAL FUNCTION GetTileAtt(Map as LONG, X as LONG, Y as LONG) AS LONG = 95

EXTERNAL SUB PlayCustomWav(Player AS LONG, SoundNum AS LONG) = 108
EXTERNAL SUB Timer(Player AS LONG, Seconds AS LONG, Script AS STRING) = 75

EXTERNAL SUB CreateTileEffect(Map AS LONG, X AS LONG, Y AS LONG, Sprite AS LONG, Speed AS LONG, TotalFrames AS LONG, LoopCount AS LONG, EndSound AS LONG) = 110
EXTERNAL SUB CreateCharacterEffect(Map AS LONG, Player AS LONG, Sprite AS LONG, Speed AS LONG, TotalFrames AS LONG, LoopCount AS LONG, EndSound AS LONG) = 111
EXTERNAL SUB CreateMonsterEffect(Map AS LONG, Player AS LONG, Monster AS LONG, Sprite AS LONG, Speed AS LONG, TotalFrames AS LONG, EndSound AS LONG) = 112
EXTERNAL SUB CreatePlayerEffect(Map AS LONG, SourcePlayer AS LONG, TargetPlayer AS LONG, Sprite AS LONG, Speed AS LONG, TotalFrames AS LONG, EndSound AS LONG) = 113

CONST BLACK = 0
CONST BLUE = 1
CONST GREEN = 2
CONST CYAN = 3
CONST RED = 4
CONST MAGENTA = 5
CONST BROWN = 6
CONST GREY = 7
CONST DARKGREY = 8
CONST BRIGHTBLUE = 9
CONST BRIGHTGREEN = 10
CONST BRIGHTCYAN = 11
CONST BRIGHTRED = 12
CONST BRIGHTMAGENTA = 13
CONST YELLOW = 14
CONST WHITE = 15

CONST CONTINUE = 0
CONST STOP = 1
FUNCTION Main(Player as LONG, Command as STRING, Parm1 as STRING, Parm2 as STRING, Parm3 as STRING) AS LONG
   Main = Continue
   If StrCmp(Command, "gods") Then
      Dim i as long
      PlayerMessage(Player,"Current Available Gods:",WHite)
      For i = 0 to getmaxusers
         If GetPlayerAccess(i) >0 & IsPlaying(i) Then
            PlayerMessage(Player,StrCat(StrCat(GetPlayerName(i),"   Access: "),Str(GetPlayerAccess(i))),White)
         End if
      Next i
      Main = Stop
   End if
   
   If StrCmp(Command, "so") & GetPlayerAccess(Player) >0 Then
      SpawnObject(GetPlayerMap(Player),Val(Parm1),Val(Parm2),GetPlayerX(Player),GetPlayerY(Player))
      Main = Stop
   End if
   
   If StrCmp(Command, "info") & GetPlayerAccess(Player) > 0 Then
      Dim PlayNum AS LONG
      PlayNum = FindPlayer(Parm1)
      If PlayNum >= 0 & IsPlaying(Player) Then
         RunScript2("INFO", Player, PlayNum)
      Else
         PlayerMessage(Player, "No such player!", WHITE)
      End if
     Main = Stop
   End If
   
   If GetPlayerAccess(Player) > 0 & StrCmp(command, "up") Then
      PlayerWarp(Player, GetPlayerMap(Player), GetPlayerX(Player), GetPlayerY(Player) - 1)
      Main = Stop
   End If
   
   If GetPlayerAccess(Player) > 0 & StrCmp(command, "down") Then
      PlayerWarp(Player, GetPlayerMap(Player), GetPlayerX(Player), GetPlayerY(Player) + 1)
      Main = Stop 
   End If
   
   If GetPlayerAccess(Player) > 0 & StrCmp(command, "left") Then
      PlayerWarp(Player, GetPlayerMap(Player), GetPlayerX(Player) - 1, GetPlayerY(Player))
      Main = Stop
   End If
 
   If GetPlayerAccess(Player) > 0 & StrCmp(command, "right") Then
      PlayerWarp(Player, GetPlayerMap(Player), GetPlayerX(Player) + 1, GetPlayerY(Player))
      Main = Stop
   End If
   
   If StrCmp(Command, "setstatus") & GetPlayerAccess(Player) =>10 Then
      SetPlayerStatus(FindPlayer(Parm1),Val(Parm2))
      Main = Stop
   End if
   
   If StrCmp(Command, "display") Then
      DisplayObjDur(Player,Val(Parm1))
      Main = Stop
   End if
   
   If StrCmp(Command, "warp") & GetPlayerAccess(Player) > 0 Then
      PlayNum = -1
      PlayNum = FindPlayer(Parm1)
      If PlayNum >= 0 & IsPlaying(PlayNum) Then
         RunScript2("WARP", PlayNum, Player)
      Else
         PlayerMessage(Player, "No such player!", WHITE)
      End if
     Main = Stop
   End If

   If StrCmp(Command, "invoke") & GetPlayerAccess(Player) >= 10 Then
      SpawnObject(GetPlayerMap(Player),Val(Parm1),Val(Parm2),GetPlayerX(Player),GetPlayerY(Player))
   End if
END FUNCTION
