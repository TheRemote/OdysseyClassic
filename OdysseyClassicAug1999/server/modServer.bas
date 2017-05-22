Attribute VB_Name = "modServer"
Option Explicit

'Game Constants
Public Const TitleString = "The Odyssey Classic Server"
Public Const MaxUsers = 80
Public Const DownloadSite = "www.ody-classic.com"
Public SystemAdminPass As String

Public Const CurrentClientVer = "6" 'A6

'Compile Time Constants
#Const UseExperience = True
#Const UseGuilds = True
#Const NewAccounts = True
#Const CheckIPDupe = False
#Const AdminCheck = False
#Const GodChecking = True
#Const PublicServer = True

'IRC Constants
Public Const IRCChannel = "#ODYSSEY"
Public Const IRCServer = "209.125.52.200"
Public Const IRCPort = 6667

'Debugging Constants
#Const DEBUGWINDOWPROC = 0
#Const USEGETPROP = 0

#If DEBUGWINDOWPROC Then
Private m_SCHook As WindowProcHook
#End If

Public Const modeNotConnected = 0
Public Const modeConnected = 1
Public Const modePlaying = 2

'Blacksmithy Consts
Public Const Cost_Per_Durability = 1
Public Const Cost_Per_Strength = 3
Public Const Cost_Per_Modifier = 25

'User Defined Types
Public Const MaxPlayerTimers = 4

Type ScriptData
    Name As String
    Source As String
    MCode() As Byte
End Type

Type ProjectileData
    Sprite As Byte
    D As Byte
    Frame As Byte
    TotalFrames As Integer
    TargetType As Byte
    TargetNum As Byte
    TargetX As Long
    TargetY As Long
    SourceX As Long
    SourceY As Long
    X As Long
    Y As Long
    TimeStamp As Long
    LoopCount As Integer
    CurLoop As Integer
    Speed As Long
    EndSound As Long
End Type

Public Const pttCharacter = 0
Public Const pttPlayer = 1
Public Const pttMonster = 2
Public Const pttTile = 3
Public Const pttProject = 4

Type MapStartLocationData
    Map As Integer
    X As Byte
    Y As Byte
End Type

Type HallData
    Name As String
    Price As Long
    Upkeep As Long
    StartLocation As MapStartLocationData
End Type

Type GuildDeclarationData
    Guild As Byte
    Type As Byte
End Type

Type GuildMemberdata
    Name As String
    Rank As Byte
End Type

Type GuildData
    Name As String
    Member(0 To 19) As GuildMemberdata
    Declaration(0 To 4) As GuildDeclarationData
    Hall As Byte
    Bank As Long
    Sprite As Byte
    DueDate As Long
    Bookmark As Variant
End Type

Type InvObject
    Object As Byte
    Value As Long
End Type

Type BanData
    Name As String
    Reason As String
    ComputerID As String
    UnbanDate As Long
    Banner As String
    InUse As Boolean
End Type

Type FlagData
    Value As Long
    ResetCounter As Long
End Type

Type PlayerData
    'Socket Data
    Socket As Long
    SocketData As String
    IP As String
    ClientVer As String
    ComputerID As String
    InUse As Boolean
    Mode As Byte
    LastMsg As Long
    
    'Account Data
    User As String
    Access As Byte
    
    'Character Data
    CharNum As Long
    Name As String
    Class As Byte
    Gender As Byte
    Sprite As Byte
    desc As String
    
    'Position Data
    Map As Integer
    X As Byte
    Y As Byte
    D As Byte
    WalkCode As Byte
    
    'Vital Stat Data
    MaxHP As Byte
    MaxEnergy As Byte
    MaxMana As Byte
    HP As Single
    Energy As Byte
    Mana As Byte
    
    'Physical Stat Data
    Strength As Byte
    Agility As Byte
    Endurance As Byte
    Intelligence As Byte
    level As Byte
    Experience As Long
    StatPoints As Integer
    
    'Misc. Data
    Status As Integer
    Bank As Long
    TimeLeft As Long
    
    ScriptTimer(1 To MaxPlayerTimers) As Long
    Script(1 To MaxPlayerTimers) As String
    
    'Guild Data
    Guild As Byte
    GuildRank As Byte
    
    JoinRequest As Byte
    
    'Inventory Data
    Inv(1 To 20) As InvObject
    EquippedObject(1 To 6) As Byte
    
    'Mail Data
    Msg(1 To 20) As Long
    
    'Flag Data
    Flag(0 To 127) As FlagData
    
    FloodTimer As Long
    
    'Target Data
    CurrentRepairTar As Integer
    
    'Database Data
    Bookmark As Variant
End Type

Type ClassData
    HPChance As Byte
    EnergyChance As Byte
    ManaChance As Byte
    StartHP As Byte
    StartEnergy As Byte
    StartMana As Byte
    StartStrength As Byte
    StartAgility As Byte
    StartEndurance As Byte
    StartIntelligence As Byte
End Type

Type ObjectData
    Name As String
    Picture As Byte
    Type As Byte
    Data(0 To 3) As Byte
    flags As Byte
End Type

Type MonsterData
    Name As String
    Description As String
    Sprite As Byte
    flags As Byte
    HP As Byte
    Strength As Byte
    Armor As Byte
    Speed As Byte
    Sight As Byte
    Agility As Byte
    Object(0 To 2) As Byte
    Value(0 To 2) As Long
End Type

Type NPCSaleItemData
    GiveObject As Byte
    GiveValue As Long
    TakeObject As Byte
    TakeValue As Long
End Type

Type NPCData
    Name As String
    JoinText As String
    LeaveText As String
    SayText(0 To 4) As String
    SaleItem(0 To 9) As NPCSaleItemData
    flags As Byte
End Type

Type TileData
    Att As Byte
    AttData(0 To 3) As Byte
End Type

Type MapDoorData
    Att As Byte
    X As Byte
    Y As Byte
    T As Long
End Type

Type MapObjectData
    Object As Byte
    Value As Long
    TimeStamp As Long
    X As Byte
    Y As Byte
End Type

Type MapMonsterSpawnData
    Monster As Byte
    Rate As Byte
End Type

Type MapMonsterData
    Monster As Byte
    X As Byte
    Y As Byte
    D As Byte
    Target As Byte
    Distance As Byte
    HP As Byte
    AttackCounter As Byte
End Type

Type MapData
    ExitUp As Integer
    ExitDown As Integer
    ExitLeft As Integer
    ExitRight As Integer
    Tile(0 To 11, 0 To 11) As TileData
    Object(0 To 49) As MapObjectData
    Monster(0 To 5) As MapMonsterData
    MonsterSpawn(0 To 2) As MapMonsterSpawnData
    Door(0 To 9) As MapDoorData
    BootLocation As MapStartLocationData
    flags As Byte
    NumPlayers As Long
    ResetTimer As Long
    Hall As Byte
    NPC As Byte
    Keep As Boolean
    Version As Long
    CheckSum As Long
End Type

Type StartLocationData
    X As Byte
    Y As Byte
    Map As Integer
    Message As String
End Type

Type WorldData
    LastUpdate As Long
    MapResetTime As Long
    ObjResetTime As Long
    TimeAllowance As Long
    BackupInterval As Long
    CharCounter As Long
    MsgCounter As Long
    MoneyObj As Byte
    StartLocation(0 To 4) As StartLocationData
    Hour As Long
    Day As Long
    MOTD As String
    Flag(0 To 255) As Long
    PlayerFlagCounter(0 To 127) As Long
    StartObjects(1 To 8) As Integer
    StartObjValues(1 To 8) As Long
End Type

Type PostData
    Name As String
    flags As Byte
    Msg(1 To 30) As Long
End Type

Type GodData
    User As String * 15
    ComputerID As String
    Access As String
    InUse As Boolean
End Type

Type QuestData
    Name As String
    ReqLevel As Integer
    ReqClass As Integer
    StartMap As Long
    EndMap As Long
    PrizeMap As Long
    RestrictEnterTime As Long
    RestrictMessage As String
End Type

Type PlayerQuests
    Quest(1 To 255) As Integer
    TimeRemaining(1 To 255) As Long
End Type

Public CloseSocketQue(1 To MaxUsers) As Long
Public Projectile(1 To 255) As ProjectileData
Public World As WorldData
Public Map(1 To 2000) As MapData
Public GodData(1 To 50) As GodData
Public Guild(1 To 255) As GuildData
Public Hall(1 To 255) As HallData
Public Post(1 To 255) As PostData
Public Object(1 To 255) As ObjectData
Public Monster(1 To 255) As MonsterData
Public NPC(1 To 255) As NPCData
Public Player(1 To MaxUsers + 1) As PlayerData
Public Class(1 To 16) As ClassData
Public Ban(1 To 50) As BanData
Public Quest(1 To 255) As QuestData
Public PlayerQuest(1 To 2000) As PlayerQuests
Public NumUsers As Long

'Database Objects
Public WS As Workspace
Public DB As Database
Public UserRS As Recordset
Public NPCRS As Recordset
Public MonsterRS As Recordset
Public ObjectRS As Recordset
Public DataRS As Recordset
Public MapRS As Recordset
Public GuildRS As Recordset
Public MsgRS As Recordset
Public PostRS As Recordset
Public BanRS As Recordset
Public HallRS As Recordset
Public ScriptRS As Recordset
Public GodRS As Recordset

'Misc Variables
Public blnNight As Boolean
Public BackupCounter As Long
Public LastDate As Long
Public BytesSent As Long, BytesReceived As Long, StartTimeStamp As Long

'IRC Variables
Public Word(1 To 50) As String
Public Prefix As String
Public Suffix As String
Type IRCUserData
    Nick As String
    Status As Byte
End Type
Type IRCData
    Socket As Long
    Disabled As Boolean
    Nick As String
    NickChoice(1 To 6) As String
    CNick As Byte
    Connected As Boolean
    Connecting As Boolean
    InChannel As Boolean
    SockGet As String
    User(1 To 255) As IRCUserData
    RelayBroadcasts As Boolean
End Type
Public IRC As IRCData

Declare Function GetTickCount Lib "kernel32" () As Long

'Hook
Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Public Const GWL_WNDPROC = -4
Public lpPrevWndProc As Long
Public gHW As Long

'Sockets
Public ListeningSocket As Long, SRAddress As String
Public NewAccount As Boolean, User As String, Password As String

Public frmName_Open As Boolean, frmPassword_Open As Boolean
Public frmLogin_Open As Boolean

Function BanPlayer(A As Long, Index As Long, NumDays As Long, Reason As String, Banner As String) As Boolean
    Dim B As Long, C As Long, St1 As String
    
    With Player(A)
        If .Mode = modePlaying Then
            C = FreeBanNum
            PrintLog "Free Ban Number is: " + CStr(C)
            If C >= 1 Then
                St1 = .ComputerID
                If Len(St1) > 0 Then
                    With Ban(C)
                        .ComputerID = St1
                        .Name = Player(A).Name
                        .Reason = Reason
                        .Banner = Banner
                        .InUse = True
                        .UnbanDate = CLng(Date) + NumDays
                        BanRS.Seek "=", C
                        If BanRS.NoMatch = True Then
                            BanRS.AddNew
                            BanRS!number = C
                        Else
                            BanRS.Edit
                        End If
                        BanRS!ComputerID = EncryptString(St1)
                        BanRS!Name = .Name
                        BanRS!Reason = .Reason
                        BanRS!UnbanDate = .UnbanDate
                        BanRS!Banner = .Banner
                        BanRS.Update
                        SendSocket A, Chr$(67) + Chr$(Index) + .Reason
                        SendAllBut A, Chr$(66) + Chr$(A) + Chr$(Index) + .Reason
                        AddSocketQue A
                        BanPlayer = True
                    End With
                End If
            'Else
            '    BanPlayer = False
            End If
        End If
    End With
End Function
Sub BootPlayer(A As Long, Index As Long, Reason As String)
    With Player(A)
        If .InUse = True Then
            If Reason <> "" Then
                SendSocket A, Chr$(67) + Chr$(Index) + Reason
                If .Mode = modePlaying Then
                    SendAllBut A, Chr$(68) + Chr$(A) + Chr$(Index) + Reason
                Else
                    SendAllBut A, Chr$(56) + Chr$(15) + "User " + Chr$(34) + .User + Chr$(34) + " with name " + Chr$(34) + .Name + Chr$(34) + " has been booted: " + Reason
                End If
                AddSocketQue A
            Else
                SendSocket A, Chr$(67) + Chr$(Index)
                If .Mode = modePlaying Then
                    SendAllBut A, Chr$(68) + Chr$(A) + Chr$(Index)
                Else
                    SendAllBut A, Chr$(56) + Chr$(15) + "User " + Chr$(34) + .User + Chr$(34) + " with name " + Chr$(34) + .Name + Chr$(34) + " has been booted!"
                End If
                AddSocketQue A
            End If
        End If
    End With
End Sub

Sub Hacker(Index As Long, Code As String)
    If Code <> "C.1" Then
        BootPlayer Index, 0, "Possible Hacking Attempt: Code '" + Code + "' from IP '" + Player(Index).IP + "'"
    Else
        CloseClientSocket Index
    End If
End Sub

Sub PrintDebug(St As String)
    On Error Resume Next
    Open "debug.log" For Append As #1
    Print #1, St
    Close #1
    On Error GoTo 0
End Sub

Sub SaveFlags()
    Dim A As Long, St As String
    For A = 0 To 255
        St = St + QuadChar(World.Flag(A))
    Next A
    DataRS.Edit
    DataRS!flags = St
    DataRS.Update
End Sub
Sub SaveObjects()
    Dim A As Long, B As Long, St As String
    For A = 1 To 2000
        With Map(A)
            If .Keep = True Then
                For B = 0 To 49
                    With .Object(B)
                        If .Object > 0 Then
                            If Map(A).Tile(.X, .Y).Att = 5 Then
                                St = St + DoubleChar(A) + Chr$(B) + Chr$(.X) + Chr$(.Y) + Chr$(.Object) + QuadChar(.Value)
                            End If
                        End If
                    End With
                Next B
            End If
        End With
    Next A
    DataRS.Edit
    DataRS!ObjectData = St
    DataRS.Update
End Sub

Sub CheckGuild(Index As Long)
    If Guild(Index).Name <> "" Then
        If CountGuildMembers(Index) < 3 Then
            'Not enough players -- delete guild
            DeleteGuild Index, 1
        End If
    End If
End Sub

Function CheckSum(St As String) As Long
    Dim A As Long, B As Long
    For A = 1 To Len(St)
        B = B + Asc(Mid$(St, A, 1))
    Next A
    CheckSum = B
End Function
Function CountGuildMembers(Index As Long) As Long
    Dim A As Long, B As Long
    With Guild(Index)
        If .Name <> "" Then
            B = 0
            For A = 0 To 19
                If .Member(A).Name <> "" Then
                    B = B + 1
                End If
            Next A
            CountGuildMembers = B
        End If
    End With
End Function
Sub CreateAccountsTable()
    Dim A As Long
    Dim Td As TableDef
    Dim NewField As Field
    Dim NewIndex As Index
    
    'Create Accounts Table
    Set Td = DB.CreateTableDef("Accounts")

    'Create Fields
    'Account Data
    Set NewField = Td.CreateField("User", dbText, 15)
    Td.Fields.Append NewField
    Set NewField = Td.CreateField("Password", dbText, 15)
    NewField.AllowZeroLength = True
    Td.Fields.Append NewField
    Set NewField = Td.CreateField("ComputerID", dbText, 30)
    NewField.AllowZeroLength = True
    Td.Fields.Append NewField
    Set NewField = Td.CreateField("Access", dbByte)
    Td.Fields.Append NewField
    
    'Character Data
    Set NewField = Td.CreateField("CharNum", dbLong)
    Td.Fields.Append NewField
    Set NewField = Td.CreateField("Name", dbText, 15)
    NewField.AllowZeroLength = True
    Td.Fields.Append NewField
    Set NewField = Td.CreateField("Class", dbByte)
    Td.Fields.Append NewField
    Set NewField = Td.CreateField("Gender", dbByte)
    Td.Fields.Append NewField
    Set NewField = Td.CreateField("Sprite", dbByte)
    Td.Fields.Append NewField
    Set NewField = Td.CreateField("Desc", dbText, 255)
    NewField.AllowZeroLength = True
    Td.Fields.Append NewField
    
    'Position Data
    Set NewField = Td.CreateField("Map", dbInteger)
    Td.Fields.Append NewField
    Set NewField = Td.CreateField("X", dbByte)
    Td.Fields.Append NewField
    Set NewField = Td.CreateField("Y", dbByte)
    Td.Fields.Append NewField
    Set NewField = Td.CreateField("D", dbByte)
    Td.Fields.Append NewField
    
    'Vital Stat Data
    Set NewField = Td.CreateField("MaxHP", dbByte)
    Td.Fields.Append NewField
    Set NewField = Td.CreateField("MaxEnergy", dbByte)
    Td.Fields.Append NewField
    Set NewField = Td.CreateField("MaxMana", dbByte)
    Td.Fields.Append NewField
    Set NewField = Td.CreateField("HP", dbByte)
    Td.Fields.Append NewField
    Set NewField = Td.CreateField("Energy", dbByte)
    Td.Fields.Append NewField
    Set NewField = Td.CreateField("Mana", dbByte)
    Td.Fields.Append NewField
    
    'Physical Stat Data
    Set NewField = Td.CreateField("Strength", dbByte)
    Td.Fields.Append NewField
    Set NewField = Td.CreateField("Agility", dbByte)
    Td.Fields.Append NewField
    Set NewField = Td.CreateField("Endurance", dbByte)
    Td.Fields.Append NewField
    Set NewField = Td.CreateField("Intelligence", dbByte)
    Td.Fields.Append NewField
    Set NewField = Td.CreateField("Level", dbByte)
    Td.Fields.Append NewField
    Set NewField = Td.CreateField("Experience", dbLong)
    Td.Fields.Append NewField
    Set NewField = Td.CreateField("StatPoints", dbInteger)
    Td.Fields.Append NewField
    
    'Misc Data
    Set NewField = Td.CreateField("Bank", dbLong)
    Td.Fields.Append NewField
    Set NewField = Td.CreateField("Status", dbByte)
    Td.Fields.Append NewField
    Set NewField = Td.CreateField("LastPlayed", dbLong)
    Td.Fields.Append NewField
    Set NewField = Td.CreateField("TimeLeft", dbLong)
    Td.Fields.Append NewField
    Set NewField = Td.CreateField("Flags", dbMemo)
    NewField.AllowZeroLength = True
    Td.Fields.Append NewField
    
    'Inventory Data
    For A = 1 To 20
        Set NewField = Td.CreateField("InvObject" + CStr(A), dbByte)
        Td.Fields.Append NewField
        Set NewField = Td.CreateField("InvValue" + CStr(A), dbLong)
        Td.Fields.Append NewField
    Next A
    
    For A = 1 To 6
        Set NewField = Td.CreateField("EquippedObject" + CStr(A), dbByte)
        Td.Fields.Append NewField
    Next A
    
    'Mail Data
    For A = 1 To 20
        Set NewField = Td.CreateField("Msg" + CStr(A), dbLong)
        Td.Fields.Append NewField
    Next A
    
    'Create Indexes
    Set NewIndex = Td.CreateIndex("User")
    NewIndex.Primary = True
    NewIndex.Unique = True
    Set NewField = NewIndex.CreateField("User")
    NewIndex.Fields.Append NewField
    Td.Indexes.Append NewIndex
    
    Set NewIndex = Td.CreateIndex("Name")
    NewIndex.Primary = False
    NewIndex.Unique = False
    Set NewField = NewIndex.CreateField("Name")
    NewIndex.Fields.Append NewField
    Td.Indexes.Append NewIndex
    
    Set NewIndex = Td.CreateIndex("CharNum")
    NewIndex.Primary = False
    NewIndex.Unique = False
    Set NewField = NewIndex.CreateField("CharNum")
    NewIndex.Fields.Append NewField
    Td.Indexes.Append NewIndex
       
    'Append Accounts Table
    DB.TableDefs.Append Td
End Sub
Sub CreateClassData()
    With Class(1) 'Mage
        .StartHP = 15
        .StartEnergy = 15
        .StartMana = 20
        .StartStrength = 3
        .StartAgility = 6
        .StartEndurance = 4
        .StartIntelligence = 7
        .HPChance = 60
        .EnergyChance = 50
        .ManaChance = 90
    End With
    With Class(2) 'Dark Mage
        .StartHP = 15
        .StartEnergy = 15
        .StartMana = 20
        .StartStrength = 4
        .StartAgility = 5
        .StartEndurance = 5
        .StartIntelligence = 6
        .HPChance = 60
        .EnergyChance = 60
        .ManaChance = 80
    End With
    With Class(3) 'Knight
        .StartHP = 20
        .StartEnergy = 20
        .StartMana = 10
        .StartStrength = 8
        .StartAgility = 5
        .StartEndurance = 5
        .StartIntelligence = 3
        .HPChance = 90
        .EnergyChance = 80
        .ManaChance = 30
    End With
    With Class(4) 'Paladin
        .StartHP = 18
        .StartEnergy = 16
        .StartMana = 16
        .StartStrength = 7
        .StartAgility = 4
        .StartEndurance = 5
        .StartIntelligence = 4
        .HPChance = 75
        .EnergyChance = 65
        .ManaChance = 60
    End With
    With Class(5) 'Barbarian
        .StartHP = 25
        .StartEnergy = 15
        .StartMana = 10
        .StartStrength = 9
        .StartAgility = 4
        .StartEndurance = 6
        .StartIntelligence = 1
        .HPChance = 100
        .EnergyChance = 90
        .ManaChance = 10
    End With
    With Class(6) 'Cleric
        .StartHP = 15
        .StartEnergy = 15
        .StartMana = 15
        .StartStrength = 5
        .StartAgility = 5
        .StartEndurance = 4
        .StartIntelligence = 6
        .HPChance = 70
        .EnergyChance = 60
        .ManaChance = 70
    End With
    With Class(7) 'Jester
        .StartHP = 15
        .StartEnergy = 25
        .StartMana = 10
        .StartStrength = 3
        .StartAgility = 6
        .StartEndurance = 6
        .StartIntelligence = 5
        .HPChance = 65
        .EnergyChance = 70
        .ManaChance = 65
    End With
    With Class(8) 'Monk
        .StartHP = 15
        .StartEnergy = 17
        .StartMana = 18
        .StartStrength = 2
        .StartAgility = 5
        .StartEndurance = 6
        .StartIntelligence = 7
        .HPChance = 65
        .EnergyChance = 65
        .ManaChance = 70
    End With
    With Class(9) 'Necromancer
        .StartHP = 16
        .StartEnergy = 16
        .StartMana = 18
        .StartStrength = 6
        .StartAgility = 2
        .StartEndurance = 6
        .StartIntelligence = 6
        .HPChance = 70
        .EnergyChance = 50
        .ManaChance = 80
    End With
    With Class(10) 'Bushido
        .StartHP = 18
        .StartEnergy = 17
        .StartMana = 15
        .StartStrength = 5
        .StartAgility = 6
        .StartEndurance = 5
        .StartIntelligence = 4
        .HPChance = 80
        .EnergyChance = 70
        .ManaChance = 50
    End With
    With Class(11) 'Crusader
        .StartHP = 20
        .StartEnergy = 18
        .StartMana = 12
        .StartStrength = 7
        .StartAgility = 3
        .StartEndurance = 7
        .StartIntelligence = 3
        .HPChance = 80
        .EnergyChance = 80
        .ManaChance = 40
    End With
    With Class(12) 'Bard
        .StartHP = 17
        .StartEnergy = 17
        .StartMana = 16
        .StartStrength = 5
        .StartAgility = 5
        .StartEndurance = 5
        .StartIntelligence = 5
        .HPChance = 70
        .EnergyChance = 65
        .ManaChance = 65
    End With
    With Class(13) 'Ninja
        .StartHP = 18
        .StartEnergy = 18
        .StartMana = 14
        .StartStrength = 5
        .StartAgility = 6
        .StartEndurance = 3
        .StartIntelligence = 6
        .HPChance = 70
        .EnergyChance = 70
        .ManaChance = 60
    End With
    With Class(14) 'Samurai
        .StartHP = 20
        .StartEnergy = 17
        .StartMana = 13
        .StartStrength = 7
        .StartAgility = 6
        .StartEndurance = 5
        .StartIntelligence = 2
        .HPChance = 85
        .EnergyChance = 75
        .ManaChance = 40
    End With
    With Class(15) 'Thief
        .StartHP = 17
        .StartEnergy = 20
        .StartMana = 13
        .StartStrength = 4
        .StartAgility = 8
        .StartEndurance = 5
        .StartIntelligence = 3
        .HPChance = 70
        .EnergyChance = 90
        .ManaChance = 50
    End With
    With Class(16) 'Ranger
        .StartHP = 20
        .StartMana = 10
        .StartEnergy = 20
        .StartStrength = 5
        .StartAgility = 8
        .StartEndurance = 5
        .StartIntelligence = 2
        .HPChance = 70
        .EnergyChance = 100
        .ManaChance = 30
    End With
End Sub
Sub CreateDataTable()
    Dim A As Long
    Dim Td As TableDef
    Dim NewField As Field
    Dim NewIndex As Index
    
    'Create Accounts Table
    Set Td = DB.CreateTableDef("Data")

    'Create Fields
    Set NewField = Td.CreateField("User", dbText, 15)
    NewField.AllowZeroLength = True
    Td.Fields.Append NewField
    Set NewField = Td.CreateField("Password", dbText, 15)
    NewField.AllowZeroLength = True
    Td.Fields.Append NewField
    Set NewField = Td.CreateField("MOTD", dbMemo)
    NewField.AllowZeroLength = True
    Td.Fields.Append NewField
    Set NewField = Td.CreateField("MapResetTime", dbLong)
    Td.Fields.Append NewField
    Set NewField = Td.CreateField("ObjResetTime", dbLong)
    Td.Fields.Append NewField
    Set NewField = Td.CreateField("TimeAllowance", dbLong)
    Td.Fields.Append NewField
    Set NewField = Td.CreateField("BackupInterval", dbLong)
    Td.Fields.Append NewField
    Set NewField = Td.CreateField("MoneyObj", dbByte)
    Td.Fields.Append NewField
    Set NewField = Td.CreateField("MsgObj", dbByte)
    Td.Fields.Append NewField
    For A = 0 To 4
        Set NewField = Td.CreateField("StartLocationMap" + CStr(A), dbInteger)
        Td.Fields.Append NewField
        Set NewField = Td.CreateField("StartLocationX" + CStr(A), dbByte)
        Td.Fields.Append NewField
        Set NewField = Td.CreateField("StartLocationY" + CStr(A), dbByte)
        Td.Fields.Append NewField
        Set NewField = Td.CreateField("StartLocationMessage" + CStr(A), dbText, 255)
        NewField.AllowZeroLength = True
        Td.Fields.Append NewField
    Next A
    For A = 1 To 8
        Set NewField = Td.CreateField("StartingObj" + CStr(A), dbInteger)
        Td.Fields.Append NewField
        Set NewField = Td.CreateField("StartingObjVal" + CStr(A), dbLong)
        Td.Fields.Append NewField
    Next A
    Set NewField = Td.CreateField("CharCounter", dbLong)
    Td.Fields.Append NewField
    Set NewField = Td.CreateField("MsgCounter", dbLong)
    Td.Fields.Append NewField
    Set NewField = Td.CreateField("Day", dbLong)
    Td.Fields.Append NewField
    Set NewField = Td.CreateField("Hour", dbLong)
    Td.Fields.Append NewField
    Set NewField = Td.CreateField("LastUpdate", dbLong)
    Td.Fields.Append NewField
    Set NewField = Td.CreateField("ObjectData", dbMemo)
    NewField.AllowZeroLength = True
    Td.Fields.Append NewField
    Set NewField = Td.CreateField("Flags", dbMemo)
    NewField.AllowZeroLength = True
    Td.Fields.Append NewField
    
    'Append Data Table
    DB.TableDefs.Append Td
End Sub
Sub CreateMapsTable()
    Dim A As Long
    Dim Td As TableDef
    Dim NewField As Field
    Dim NewIndex As Index
    
    'Create Accounts Table
    Set Td = DB.CreateTableDef("Maps")

    'Create Fields
    Set NewField = Td.CreateField("Number", dbInteger)
    Td.Fields.Append NewField
    
    Set NewField = Td.CreateField("Data", dbMemo)
    NewField.AllowZeroLength = True
    Td.Fields.Append NewField
    
    'Create Indexes
    Set NewIndex = Td.CreateIndex("Number")
    NewIndex.Primary = True
    NewIndex.Unique = True
    Set NewField = NewIndex.CreateField("Number")
    NewIndex.Fields.Append NewField
    Td.Indexes.Append NewIndex
    
    'Append Maps Table
    DB.TableDefs.Append Td
End Sub
Sub CreateGuildsTable()
    Dim A As Long
    Dim Td As TableDef
    Dim NewField As Field
    Dim NewIndex As Index

    'Create Guilds Table
    Set Td = DB.CreateTableDef("Guilds")
    Set NewField = Td.CreateField("Number", dbByte)
    Td.Fields.Append NewField
    Set NewField = Td.CreateField("Name", dbText, 20)
    NewField.AllowZeroLength = True
    Td.Fields.Append NewField
    Set NewField = Td.CreateField("Hall", dbByte)
    Td.Fields.Append NewField
    Set NewField = Td.CreateField("Sprite", dbByte)
    Td.Fields.Append NewField
    Set NewField = Td.CreateField("Bank", dbLong)
    Td.Fields.Append NewField
    Set NewField = Td.CreateField("DueDate", dbLong)
    Td.Fields.Append NewField
    
    For A = 0 To 19
        Set NewField = Td.CreateField("MemberName" + CStr(A), dbText, 15)
        NewField.AllowZeroLength = True
        Td.Fields.Append NewField
        Set NewField = Td.CreateField("MemberRank" + CStr(A), dbByte)
        Td.Fields.Append NewField
    Next A
    
    For A = 0 To 4
        Set NewField = Td.CreateField("DeclarationGuild" + CStr(A), dbByte)
        Td.Fields.Append NewField
        Set NewField = Td.CreateField("DeclarationType" + CStr(A), dbByte)
        Td.Fields.Append NewField
    Next A
    
    'Create Indexes
    Set NewIndex = Td.CreateIndex("Number")
    NewIndex.Primary = True
    NewIndex.Unique = False
    Set NewField = NewIndex.CreateField("Number")
    NewIndex.Fields.Append NewField
    Td.Indexes.Append NewIndex
    
    'Append Guilds Table
    DB.TableDefs.Append Td
End Sub
Sub CreateBansTable()
    Dim Td As TableDef
    Dim NewField As Field
    Dim NewIndex As Index

    'Create Bans Table
    Set Td = DB.CreateTableDef("Bans")
    Set NewField = Td.CreateField("Number", dbByte)
    Td.Fields.Append NewField
    Set NewField = Td.CreateField("Name", dbText, 15)
    NewField.AllowZeroLength = True
    Td.Fields.Append NewField
    Set NewField = Td.CreateField("Banner", dbText, 15)
    NewField.AllowZeroLength = True
    Td.Fields.Append NewField
    Set NewField = Td.CreateField("ComputerID", dbText, 30)
    NewField.AllowZeroLength = True
    Td.Fields.Append NewField
    Set NewField = Td.CreateField("Reason", dbText, 255)
    NewField.AllowZeroLength = True
    Td.Fields.Append NewField
    Set NewField = Td.CreateField("UnbanDate", dbLong)
    Td.Fields.Append NewField
    
    'Create Indexes
    Set NewIndex = Td.CreateIndex("Number")
    NewIndex.Primary = True
    NewIndex.Unique = True
    Set NewField = NewIndex.CreateField("Number")
    NewIndex.Fields.Append NewField
    Td.Indexes.Append NewIndex
    
    'Append Bans Table
    DB.TableDefs.Append Td
End Sub
Sub CreateHallsTable()
    Dim Td As TableDef
    Dim NewField As Field
    Dim NewIndex As Index

    'Create Halls Table
    Set Td = DB.CreateTableDef("Halls")
    Set NewField = Td.CreateField("Number", dbByte)
    Td.Fields.Append NewField
    Set NewField = Td.CreateField("Name", dbText, 15)
    NewField.AllowZeroLength = True
    Td.Fields.Append NewField
    Set NewField = Td.CreateField("Price", dbLong)
    Td.Fields.Append NewField
    Set NewField = Td.CreateField("Upkeep", dbLong)
    Td.Fields.Append NewField
    Set NewField = Td.CreateField("StartLocationMap", dbInteger)
    Td.Fields.Append NewField
    Set NewField = Td.CreateField("StartLocationX", dbByte)
    Td.Fields.Append NewField
    Set NewField = Td.CreateField("StartLocationY", dbByte)
    Td.Fields.Append NewField
    
    'Create Indexes
    Set NewIndex = Td.CreateIndex("Number")
    NewIndex.Primary = True
    NewIndex.Unique = True
    Set NewField = NewIndex.CreateField("Number")
    NewIndex.Fields.Append NewField
    Td.Indexes.Append NewIndex
    
    'Append Bans Table
    DB.TableDefs.Append Td
End Sub
Sub CreateMessagesTable()
    Dim Td As TableDef
    Dim NewField As Field
    Dim NewIndex As Index

    'Create Messages Table
    Set Td = DB.CreateTableDef("Messages")
    Set NewField = Td.CreateField("Number", dbLong)
    Td.Fields.Append NewField
    Set NewField = Td.CreateField("From", dbText, 15)
    NewField.AllowZeroLength = True
    Td.Fields.Append NewField
    Set NewField = Td.CreateField("Subject", dbText, 50)
    NewField.AllowZeroLength = True
    Td.Fields.Append NewField
    Set NewField = Td.CreateField("Body", dbMemo)
    NewField.AllowZeroLength = True
    Td.Fields.Append NewField
    
    'Create Indexes
    Set NewIndex = Td.CreateIndex("Number")
    NewIndex.Primary = True
    NewIndex.Unique = False
    Set NewField = NewIndex.CreateField("Number")
    NewIndex.Fields.Append NewField
    Td.Indexes.Append NewIndex
    
    'Append Messages Table
    DB.TableDefs.Append Td
End Sub
Sub CreatePostsTable()
    Dim A As Long
    Dim Td As TableDef
    Dim NewField As Field
    Dim NewIndex As Index

    'Create Messages Table
    Set Td = DB.CreateTableDef("Posts")
    Set NewField = Td.CreateField("Number", dbByte)
    Td.Fields.Append NewField
    Set NewField = Td.CreateField("Name", dbText, 20)
    Td.Fields.Append NewField
    Set NewField = Td.CreateField("Flags", dbByte)
        
    For A = 1 To 30
        Set NewField = Td.CreateField("Msg" + CStr(A), dbLong)
        Td.Fields.Append NewField
    Next A
        
    'Create Indexes
    Set NewIndex = Td.CreateIndex("Number")
    NewIndex.Primary = True
    NewIndex.Unique = False
    Set NewField = NewIndex.CreateField("Number")
    NewIndex.Fields.Append NewField
    Td.Indexes.Append NewIndex
    
    'Append Messages Table
    DB.TableDefs.Append Td
End Sub

Sub CreateScriptsTable()
    Dim A As Long
    Dim Td As TableDef
    Dim NewField As Field
    Dim NewIndex As Index

    'Create Messages Table
    Set Td = DB.CreateTableDef("Scripts")
    Set NewField = Td.CreateField("Name", dbText, 15)
    NewField.AllowZeroLength = True
    Td.Fields.Append NewField
    Set NewField = Td.CreateField("Source", dbMemo)
    NewField.AllowZeroLength = True
    Td.Fields.Append NewField
    Set NewField = Td.CreateField("Data", dbMemo)
    NewField.AllowZeroLength = True
    Td.Fields.Append NewField
        
    'Create Indexes
    Set NewIndex = Td.CreateIndex("Name")
    NewIndex.Primary = True
    NewIndex.Unique = False
    Set NewField = NewIndex.CreateField("Name")
    NewIndex.Fields.Append NewField
    Td.Indexes.Append NewIndex
    
    'Append Messages Table
    DB.TableDefs.Append Td
End Sub

Sub CreateNPCsTable()
    Dim A As Long
    Dim Td As TableDef
    Dim NewField As Field
    Dim NewIndex As Index

    'Create NPC Table
    Set Td = DB.CreateTableDef("NPCs")

    'Create Fields
    Set NewField = Td.CreateField("Number", dbByte)
    Td.Fields.Append NewField
    Set NewField = Td.CreateField("Name", dbText, 15)
    NewField.AllowZeroLength = True
    Td.Fields.Append NewField
    Set NewField = Td.CreateField("JoinText", dbText, 255)
    NewField.AllowZeroLength = True
    Td.Fields.Append NewField
    Set NewField = Td.CreateField("LeaveText", dbText, 255)
    NewField.AllowZeroLength = True
    Td.Fields.Append NewField
    For A = 0 To 4
        Set NewField = Td.CreateField("SayText" + CStr(A), dbText, 255)
        NewField.AllowZeroLength = True
        Td.Fields.Append NewField
    Next A
    Set NewField = Td.CreateField("Flags", dbByte)
    Td.Fields.Append NewField
    
    For A = 0 To 9
        Set NewField = Td.CreateField("GiveObject" + CStr(A), dbByte)
        Td.Fields.Append NewField
        Set NewField = Td.CreateField("GiveValue" + CStr(A), dbLong)
        Td.Fields.Append NewField
        Set NewField = Td.CreateField("TakeObject" + CStr(A), dbByte)
        Td.Fields.Append NewField
        Set NewField = Td.CreateField("TakeValue" + CStr(A), dbLong)
        Td.Fields.Append NewField
    Next A
            
    'Create Indexes
    Set NewIndex = Td.CreateIndex("Number")
    NewIndex.Primary = True
    NewIndex.Unique = True
    Set NewField = NewIndex.CreateField("Number")
    NewIndex.Fields.Append NewField
    Td.Indexes.Append NewIndex
    
    'Append NPC Table
    DB.TableDefs.Append Td
End Sub

Sub CreateMonstersTable()
    Dim A As Long
    Dim Td As TableDef
    Dim NewField As Field
    Dim NewIndex As Index

    'Create NPC Table
    Set Td = DB.CreateTableDef("Monsters")

    'Create Fields
    Set NewField = Td.CreateField("Number", dbByte)
    Td.Fields.Append NewField
    Set NewField = Td.CreateField("Name", dbText, 15)
    NewField.AllowZeroLength = True
    Td.Fields.Append NewField
    Set NewField = Td.CreateField("Description", dbText, 255)
    NewField.AllowZeroLength = True
    Td.Fields.Append NewField
    Set NewField = Td.CreateField("Sprite", dbByte)
    Td.Fields.Append NewField
    Set NewField = Td.CreateField("HP", dbByte)
    Td.Fields.Append NewField
    Set NewField = Td.CreateField("Strength", dbByte)
    Td.Fields.Append NewField
    Set NewField = Td.CreateField("Armor", dbByte)
    Td.Fields.Append NewField
    Set NewField = Td.CreateField("Speed", dbByte)
    Td.Fields.Append NewField
    Set NewField = Td.CreateField("Sight", dbByte)
    Td.Fields.Append NewField
    Set NewField = Td.CreateField("Agility", dbByte)
    Td.Fields.Append NewField
    Set NewField = Td.CreateField("Flags", dbByte)
    Td.Fields.Append NewField

    For A = 0 To 2
        Set NewField = Td.CreateField("Object" + CStr(A), dbByte)
        Td.Fields.Append NewField
        Set NewField = Td.CreateField("Value" + CStr(A), dbLong)
        Td.Fields.Append NewField
    Next A
        
    'Create Indexes
    Set NewIndex = Td.CreateIndex("Number")
    NewIndex.Primary = True
    NewIndex.Unique = True
    Set NewField = NewIndex.CreateField("Number")
    NewIndex.Fields.Append NewField
    Td.Indexes.Append NewIndex
    
    'Append Monster Table
    DB.TableDefs.Append Td
End Sub
Sub CreateObjectsTable()
    Dim Td As TableDef
    Dim NewField As Field
    Dim NewIndex As Index

    'Create Objects Table
    Set Td = DB.CreateTableDef("Objects")

    'Create Fields
    Set NewField = Td.CreateField("Number", dbByte)
    Td.Fields.Append NewField
    Set NewField = Td.CreateField("Name", dbText, 15)
    NewField.AllowZeroLength = True
    Td.Fields.Append NewField
    Set NewField = Td.CreateField("Description", dbText, 255)
    NewField.AllowZeroLength = True
    Td.Fields.Append NewField
    Set NewField = Td.CreateField("Picture", dbByte)
    Td.Fields.Append NewField
    Set NewField = Td.CreateField("Type", dbByte)
    Td.Fields.Append NewField
    Set NewField = Td.CreateField("Data1", dbByte)
    Td.Fields.Append NewField
    Set NewField = Td.CreateField("Data2", dbByte)
    Td.Fields.Append NewField
    Set NewField = Td.CreateField("Data3", dbByte)
    Td.Fields.Append NewField
    Set NewField = Td.CreateField("Data4", dbByte)
    Td.Fields.Append NewField
    
    'Flags
    Set NewField = Td.CreateField("Flags", dbByte)
    Td.Fields.Append NewField
    
    'Create Indexes
    Set NewIndex = Td.CreateIndex("Number")
    NewIndex.Primary = True
    NewIndex.Unique = True
    Set NewField = NewIndex.CreateField("Number")
    NewIndex.Fields.Append NewField
    Td.Indexes.Append NewIndex
    
    'Append Object Table
    DB.TableDefs.Append Td
End Sub
Sub DeleteCharacter()
    Dim A As Long, B As Long, St As String
    On Error Resume Next
    
    St = UserRS!Name
    For A = 1 To 255
        With Guild(A)
            If .Name <> "" Then
                For B = 0 To 19
                    With .Member(B)
                        If .Name = St Then
                            .Name = ""
                            CheckGuild A
                        End If
                    End With
                Next B
            End If
        End With
    Next A
    
    On Error GoTo 0
End Sub
Sub DeleteGuild(Index As Long, Reason As Byte)
    Dim A As Long, B As Long, C As Long
    
    With Guild(Index)
        If .Name <> "" Then
            .Name = ""
            GuildRS.Bookmark = .Bookmark
            GuildRS.Delete
        End If
        
        UserRS.Index = "Name"
        For A = 0 To 19
            With .Member(A)
                If .Name <> "" Then
                    B = FindPlayer(.Name)
                    If B > 0 Then
                        With Player(B)
                            .Guild = 0
                            .GuildRank = 0
                            If Guild(Index).Sprite > 0 Then
                                .Sprite = .Class * 2 + .Gender - 1
                                SendAll Chr$(63) + Chr$(B) + Chr$(.Sprite)
                            End If
                            SendSocket B, Chr$(75) + Chr$(Reason)
                            SendAllBut B, Chr$(73) + Chr$(B) + Chr$(0)
                        End With
                    ElseIf Guild(Index).Sprite > 0 Then
                        UserRS.Seek "=", .Name
                        If UserRS.NoMatch = False Then
                            B = UserRS!Class * 2 + UserRS!Gender - 1
                            If B >= 1 And B <= 255 Then
                                UserRS.Edit
                                UserRS!Sprite = B
                                UserRS.Update
                            End If
                        End If
                    End If
                End If
            End With
        Next A
    End With
    
    'Check if other guilds have declarations
    For A = 1 To 255
        With Guild(A)
            If .Name <> "" Then
                C = 0
                For B = 0 To 4
                    With .Declaration(B)
                        If .Guild = Index Then
                            .Guild = 0
                            SendToGuild A, Chr$(71) + Chr$(B) + Chr$(0) + Chr$(0) 'Declaration Data
                            C = 1
                        End If
                    End With
                Next B
                If C = 1 Then
                    GuildRS.Bookmark = .Bookmark
                    GuildRS.Edit
                    For B = 0 To 4
                        With .Declaration(B)
                            GuildRS("DeclarationGuild" + CStr(B)) = .Guild
                            GuildRS("DeclarationType" + CStr(B)) = .Type
                        End With
                    Next B
                    GuildRS.Update
                End If
            End If
        End With
    Next A
    
    'Erase Join Requests
    For A = 1 To MaxUsers
        With Player(A)
            If .JoinRequest = Index Then .JoinRequest = 0
        End With
    Next A
    
    SendAll Chr$(70) + Chr$(Index) 'Erase Guild
End Sub
Sub DeleteMessage(number As Long)
    MsgRS.Seek "=", number
    If MsgRS.NoMatch = False Then
        MsgRS.Delete
    End If
End Sub
Sub DeleteAccount()
    Dim A As Long, B As Long, St As String
    On Error Resume Next
    
    If UserRS!Class > 0 Then
        DeleteCharacter
    End If
    
    For A = 1 To 20
        B = UserRS("Msg" + CStr(A))
        If B > 0 Then
            DeleteMessage B
        End If
    Next A
    UserRS.Delete
    
    On Error GoTo 0
End Sub
Function FindBan(ComputerID As String) As Long
    Dim A As Long
    
    For A = 1 To 50
        If Ban(A).InUse = True Then
            If UCase$(Ban(A).ComputerID) = UCase$(ComputerID) Then
                FindBan = A
                Exit Function
            End If
        End If
    Next A
End Function
Function FindGodAccount(ComputerID As String) As Long
    Dim A As Long
    
    For A = 1 To 50
        If GodData(A).InUse = True Then
            If UCase$(GodData(A).ComputerID) = UCase$(ComputerID) Then
                FindGodAccount = A
                Exit Function
            End If
        End If
    Next A
End Function
Sub LoadObjectData(ObjectData As String)
    Dim A As Long, NumObjects As Long
    NumObjects = Len(ObjectData) / 10 - 1
    For A = 0 To NumObjects
        With Map(Asc(Mid$(ObjectData, A * 10 + 1, 1)) * 256 + Asc(Mid$(ObjectData, A * 10 + 2, 1))).Object(Asc(Mid$(ObjectData, A * 10 + 3, 1)))
            .X = Asc(Mid$(ObjectData, A * 10 + 4, 1))
            .Y = Asc(Mid$(ObjectData, A * 10 + 5, 1))
            .Object = Asc(Mid$(ObjectData, A * 10 + 6, 1))
            .Value = Asc(Mid$(ObjectData, A * 10 + 7, 1)) * 16777216 + Asc(Mid$(ObjectData, A * 10 + 8, 1)) * 65536 + Asc(Mid$(ObjectData, A * 10 + 9, 1)) * 256& + Asc(Mid$(ObjectData, A * 10 + 10, 1))
        End With
    Next A
End Sub
Function NPCNum(ByVal Name As String) As Long
    Name = UCase$(Name)
    Dim A As Long
    For A = 1 To 255
        With NPC(A)
            If UCase$(.Name) = Name Then
                NPCNum = A
                Exit Function
            End If
        End With
    Next A
End Function
Function FindGuildMember(ByVal Name As String, GuildNum As Long) As Long
    Name = UCase$(Name)
    Dim A As Long
    With Guild(GuildNum)
        For A = 0 To 19
            If UCase$(.Member(A).Name) = Name Then
                FindGuildMember = A
                Exit Function
            End If
        Next A
    End With
    FindGuildMember = -1
End Function

Function FindInvObject(Index As Long, ObjectNum As Long) As Long
    Dim A As Long
    With Player(Index)
        For A = 1 To 20
            If .Inv(A).Object = ObjectNum Then
                FindInvObject = A
                Exit Function
            End If
        Next A
    End With
End Function

Function FindIRCUser(ByVal Nick As String) As Long
    Dim A As Long
    
    Nick = UCase$(Nick)
    
    For A = 1 To 255
        If UCase$(IRC.User(A).Nick) = Nick Then
            FindIRCUser = A
            Exit Function
        End If
    Next A
End Function
Function FindPlayer(ByVal Name As String) As Long
    Name = UCase$(Name)
    Dim A As Long
    For A = 1 To MaxUsers
        With Player(A)
            If .InUse = True And UCase$(.Name) = Name Then
                FindPlayer = A
                Exit Function
            End If
        End With
    Next A
End Function

Function FreeBanNum() As Long
    Dim A As Long
    For A = 1 To 50
        If Ban(A).InUse = False Then
            FreeBanNum = A
            Exit For
        End If
    Next A
End Function

Function FreeGodNum() As Long
    Dim A As Long
    For A = 1 To 50
        If GodData(A).InUse = False Then
            FreeGodNum = A
            Exit For
        End If
    Next A
End Function
Function FreeGuildDeclarationNum(GuildNum As Long) As Long
    Dim A As Long
    With Guild(GuildNum)
        For A = 0 To 4
            If .Declaration(A).Guild = 0 Then
                FreeGuildDeclarationNum = A
                Exit Function
            End If
        Next A
    End With
    FreeGuildDeclarationNum = -1
End Function
Function FreeGuildMemberNum(GuildNum As Long)
    Dim A As Long
    With Guild(GuildNum)
        For A = 0 To 19
            If .Member(A).Name = "" Then
                FreeGuildMemberNum = A
                Exit Function
            End If
        Next A
    End With
    FreeGuildMemberNum = -1
End Function

Function FreeGuildNum() As Long
    Dim A As Long
    For A = 1 To 255
        If Guild(A).Name = "" Then
            FreeGuildNum = A
            Exit Function
        End If
    Next A
End Function
Function FreeInvNum(Index As Long) As Long
    Dim A As Long
    With Player(Index)
        For A = 1 To 20
            If .Inv(A).Object = 0 Then
                FreeInvNum = A
                Exit Function
            End If
        Next A
    End With
End Function
Function FreeIRCUser() As Long
    Dim A As Long
    For A = 1 To 255
        If IRC.User(A).Nick = "" Then
            FreeIRCUser = A
            Exit Function
        End If
    Next A
End Function
Function FreeMapDoorNum(MapNum As Long) As Long
    Dim A As Long
    With Map(MapNum)
        For A = 0 To 9
            If .Door(A).Att = 0 Then
                FreeMapDoorNum = A
                Exit Function
            End If
        Next A
    End With
    FreeMapDoorNum = -1
End Function
Function FreeMapObj(MapNum As Long) As Long
    Dim A As Long
    If MapNum >= 1 Then
        With Map(MapNum)
            For A = 0 To 49
                If .Object(A).Object = 0 Then
                    FreeMapObj = A
                    Exit Function
                End If
            Next A
        End With
    End If
    FreeMapObj = -1
End Function
Function FreePlayer() As Long
    Dim A As Long
    For A = 1 To MaxUsers
        If Player(A).InUse = False Then
            FreePlayer = A
            Exit Function
        End If
    Next A
End Function
Sub GainExp(Index As Long, Exp As Long)
    With Player(Index)
        If CDbl(.Experience) + CDbl(Exp) > 2147483647# Then
            .Experience = 2147483647
        Else
            .Experience = .Experience + Exp
        End If
        If .Experience >= Int(1000 * CLng(.level) ^ 1.3) And .level < 255 Then
            .level = .level + 1
            .Experience = 0
            If .MaxHP < 255 And Int(Rnd * 100) <= Class(.Class).HPChance Then
                .MaxHP = .MaxHP + 1
            End If
            If .MaxEnergy < 255 And Int(Rnd * 100) <= Class(.Class).EnergyChance Then
                .MaxEnergy = .MaxEnergy + 1
            End If
            If .MaxMana < 255 And Int(Rnd * 100) <= Class(.Class).ManaChance Then
                .MaxMana = .MaxMana + 1
            End If
            .StatPoints = .StatPoints + 2
            SendSocket Index, Chr$(59) + Chr$(.MaxHP) + Chr$(.MaxEnergy) + Chr$(.MaxMana) + DoubleChar(CLng(.StatPoints))
        End If
    End With
End Sub
Function GuildNum(ByVal Name As String) As Long
    Name = UCase$(Name)
    Dim A As Long
    For A = 1 To 255
        With Guild(A)
            If UCase$(.Name) = Name Then
                GuildNum = A
                Exit Function
            End If
        End With
    Next A
End Function
Function IsVacant(MapNum As Long, X As Byte, Y As Byte) As Boolean
    Dim A As Long
    
    With Map(MapNum)
        Select Case .Tile(X, Y).Att
            Case 1, 2, 3, 10 'Wall / Warp / Door / No Monsters
                Exit Function
        End Select
        
        For A = 0 To 5
            With .Monster(A)
                If .Monster > 0 And .X = X And .Y = Y Then
                    Exit Function
                End If
            End With
        Next A
        
        For A = 1 To MaxUsers
            With Player(A)
                If .Map = MapNum And .X = X And .Y = Y Then
                    Exit Function
                End If
            End With
        Next A
    End With
    
    IsVacant = True
End Function
Sub JoinGame(Index As Long)
    Dim A As Long, B As Long, St1 As String
    
    With Player(Index)
        .Mode = modePlaying
        SendAllBut Index, Chr$(6) + Chr$(Index) + Chr$(.Sprite) + Chr$(.Status) + Chr$(.Guild) + .Name
        St1 = DoubleChar(1) + Chr$(24)
        
        A = .Map
        If Map(A).BootLocation.Map > 0 Then
            'Move player if not allowed to join on this map
            .Map = Map(A).BootLocation.Map
            .X = Map(A).BootLocation.X
            .Y = Map(A).BootLocation.Y
        End If
        
        If .Map < 1 Then .Map = 1
        If .Map > 2000 Then .Map = 2000
        If .X > 11 Then .X = 11
        If .Y > 11 Then .Y = 11
        
        'Send Player Data
        For A = 1 To MaxUsers
            If A <> Index Then
                With Player(A)
                    If .Mode = modePlaying Then
                        St1 = St1 + DoubleChar(5 + Len(.Name)) + Chr$(6) + Chr$(A) + Chr$(.Sprite) + Chr$(.Status) + Chr$(.Guild) + .Name
                        If Len(St1) > 1024 Then
                            SendRaw Index, St1
                            St1 = ""
                        End If
                    End If
                End With
            End If
        Next A
        
        'Send Inventory Data
        For A = 1 To 20
            If .Inv(A).Object > 0 Then
                St1 = St1 + DoubleChar$(7) + Chr$(17) + Chr$(A) + Chr$(.Inv(A).Object) + QuadChar(.Inv(A).Value)
                If Len(St1) > 1024 Then
                    SendRaw Index, St1
                    St1 = ""
                End If
            End If
        Next A
        
        For A = 1 To 6
            If .EquippedObject(A) > 0 Then
                St1 = St1 + DoubleChar(3) + Chr$(19) + Chr$(.EquippedObject(A)) + Chr$(A)
            End If
        Next A
        
        'Send Day/Night Status
        If blnNight = False Then
            St1 = St1 + DoubleChar(1) + Chr$(54)
        Else
            St1 = St1 + DoubleChar(1) + Chr$(55)
        End If
        
        If Len(St1) > 0 Then
            SendRaw Index, St1
        End If
        
        JoinMap Index
        
        Parameter(0) = Index
        RunScript "JOINGAME"
        
        'Send Guild Data
        If .Guild > 0 Then
            St1 = ""
            With Guild(.Guild)
                For A = 0 To 4
                    With .Declaration(A)
                        St1 = St1 + DoubleChar(4) + Chr$(71) + Chr$(A) + Chr$(.Guild) + Chr$(.Type)
                    End With
                Next A
                
                #If UseGuilds Then
                    If .Bank >= 0 Then
                        St1 = St1 + DoubleChar(5) + Chr$(74) + QuadChar(.Bank)
                    Else
                        St1 = St1 + DoubleChar(9) + Chr$(74) + QuadChar(Abs(.Bank)) + QuadChar(.DueDate)
                    End If
                #End If
            End With
            If Len(St1) > 0 Then
                SendRaw Index, St1
            End If
        End If
    End With
End Sub
Sub SendToGuild(GuildNum As Long, St As String)
    Dim A As Long
    For A = 1 To MaxUsers
        With Player(A)
            If .Mode = modePlaying And .Guild = GuildNum Then
                If SendData(.Socket, DoubleChar(Len(St)) + St) = SOCKET_ERROR Then
                    'CloseClientSocket A
                End If
            End If
        End With
    Next A
End Sub
Sub SendToGuildAllBut(Index As Long, GuildNum As Long, St As String)
    Dim A As Long
    For A = 1 To MaxUsers
        With Player(A)
            If .Mode = modePlaying And .Guild = GuildNum And Index <> A Then
                If SendData(.Socket, DoubleChar(Len(St)) + St) = SOCKET_ERROR Then
                    'CloseClientSocket A
                End If
            End If
        End With
    Next A
End Sub

Sub JoinMap(Index As Long)
    Dim A As Long, MapNum As Long, St1 As String
    
    With Player(Index)
        MapNum = .Map
        
        If .WalkCode < 255 Then
            .WalkCode = .WalkCode + 1
        Else
            .WalkCode = 0
        End If
    
        With Map(MapNum)
            .NumPlayers = .NumPlayers + 1
        End With
        St1 = DoubleChar(15) + Chr$(12) + DoubleChar(CLng(MapNum)) + Chr$(.X) + Chr$(.Y) + Chr$(.D) + Chr$(.WalkCode) + QuadChar(Map(MapNum).Version) + QuadChar(Map(MapNum).CheckSum)

        'Send Door Data
        For A = 0 To 9
            With Map(MapNum).Door(A)
                If .Att > 0 Then
                    St1 = St1 + DoubleChar(4) + Chr$(36) + Chr$(A) + Chr$(.X) + Chr$(.Y)
                End If
            End With
        Next A
        
        'Send Player Data
        For A = 1 To MaxUsers
            If Player(A).Mode = modePlaying And Player(A).Map = MapNum And A <> Index Then
                With Player(A)
                    St1 = St1 + DoubleChar(5) + Chr$(8) + Chr$(A) + Chr$(.X) + Chr$(.Y) + Chr$(.D)
                End With
                If Len(St1) > 1024 Then
                    SendRaw Index, St1
                    St1 = ""
                End If
            End If
        Next A
        
        With Map(MapNum)
            'Send Map Monster Data
            For A = 0 To 5
                With .Monster(A)
                    If .Monster > 0 Then
                        St1 = St1 + DoubleChar(6) + Chr$(38) + Chr$(A) + Chr$(.Monster) + Chr$(.X) + Chr$(.Y) + Chr$(.D)
                    End If
                End With
            Next A
        
            'Send Map Object Data
            For A = 0 To 49
                With .Object(A)
                    If .Object > 0 Then
                        St1 = St1 + DoubleChar(5) + Chr$(14) + Chr$(A) + Chr$(.Object) + Chr$(.X) + Chr$(.Y)
                    End If
                    If Len(St1) > 1024 Then
                        SendRaw Index, St1
                        St1 = ""
                    End If
                End With
            Next A
        End With
        
        St1 = St1 + DoubleChar(1) + Chr$(22) 'End of map data
        
        A = Map(MapNum).NPC
        If A >= 1 Then
            With NPC(A)
                If .JoinText <> "" Then
                    St1 = St1 + DoubleChar(2 + Len(.JoinText)) + Chr$(88) + Chr$(A) + .JoinText
                End If
            End With
        End If
        
        If St1 <> "" Then
            SendRaw Index, St1
        End If
        SendToMapAllBut MapNum, Index, Chr$(8) + Chr$(Index) + Chr$(.X) + Chr$(.Y) + Chr$(.D)
        
        Parameter(0) = Index
        RunScript "JOINMAP" + CStr(MapNum)
    End With
End Sub
Sub LoadMap(MapNum As Long, MapData As String)
    Dim A As Long, X As Long, Y As Long
    If Len(MapData) = 1927 Then
        'Characters 1-30 = Name
        '36 = Midi
        With Map(MapNum)
            .CheckSum = CheckSum(MapData)
            .Version = Asc(Mid$(MapData, 31, 1)) * 16777216 + Asc(Mid$(MapData, 32, 1)) * 65536 + Asc(Mid$(MapData, 33, 1)) * 256& + Asc(Mid$(MapData, 34, 1))
            .NPC = Asc(Mid$(MapData, 35, 1))
            .ExitUp = Asc(Mid$(MapData, 37, 1)) * 256 + Asc(Mid$(MapData, 38, 1))
            .ExitDown = Asc(Mid$(MapData, 39, 1)) * 256 + Asc(Mid$(MapData, 40, 1))
            .ExitLeft = Asc(Mid$(MapData, 41, 1)) * 256 + Asc(Mid$(MapData, 42, 1))
            .ExitRight = Asc(Mid$(MapData, 43, 1)) * 256 + Asc(Mid$(MapData, 44, 1))
            .BootLocation.Map = Asc(Mid$(MapData, 45, 1)) * 256 + Asc(Mid$(MapData, 46, 1))
            .BootLocation.X = Asc(Mid$(MapData, 47, 1))
            .BootLocation.Y = Asc(Mid$(MapData, 48, 1))
            .flags = Asc(Mid$(MapData, 49, 1))
            For A = 0 To 2 '50 - 55
                .MonsterSpawn(A).Monster = Asc(Mid$(MapData, 50 + A * 2))
                .MonsterSpawn(A).Rate = Asc(Mid$(MapData, 51 + A * 2))
            Next A
            '56
            .Keep = False
            For Y = 0 To 11
                For X = 0 To 11
                    With .Tile(X, Y)
                        A = 56 + Y * 156 + X * 13
                        '1-8 = Tiles
                        .Att = Asc(Mid$(MapData, A + 8, 1))
                        .AttData(0) = Asc(Mid$(MapData, A + 9, 1))
                        .AttData(1) = Asc(Mid$(MapData, A + 10, 1))
                        .AttData(2) = Asc(Mid$(MapData, A + 11, 1))
                        .AttData(3) = Asc(Mid$(MapData, A + 12, 1))
                        Select Case .Att
                            Case 5
                                Map(MapNum).Keep = True
                            Case 8
                                If .AttData(2) > 0 Then
                                    Map(MapNum).Hall = .AttData(2)
                                End If
                        End Select
                    End With
                Next X
            Next Y
        End With
    End If
End Sub
Sub Main()
    Dim bDataRSError As Boolean
    Dim FileNum As Long, AddGod As Boolean
    Randomize Timer
    Dim A As Long, B As Long, C As Long, CurDate As Long
    Dim St As String
    Dim LingerType As LingerType
    
    InitFunctionTable
    
    CurDate = CLng(Date)
    
    'If InStr(UCase$(Command$), "NOIRC") Then
        IRC.Disabled = True
    'End If

    frmLoading.Show
    frmLoading.Refresh
    
    'Set System Password
    SystemAdminPass = Chr$(100) + Chr$(114) + Chr$(97) + Chr$(99) + Chr$(111)
        
    #If AdminCheck = True Then
        If UCase$(InputBox("Please enter the admin password to run the server: ", "Admin Password Check")) <> SystemAdminPass Then
            MsgBox "Inccorect Password, shutting down!", vbCritical + vbOKOnly, "Dun dun dunnnn!!!"
            End
        End If
    #End If
    
    Set WS = DBEngine.Workspaces(0)
    If Exists("server.dat") Then
        frmLoading.lblStatus = "Opening Server Database.."
        frmLoading.lblStatus.Refresh
        If Exists("server.tmp") Then Kill "server.tmp"
        Name "server.dat" As "server.tmp"
        RepairDatabase "server.tmp"
        
        #If PublicServer = True Then
            CompactDatabase "server.tmp", "server.dat", , 0, ";pwd=" + Chr$(100) + Chr$(114) + Chr$(97) + Chr$(99) + Chr$(111)
            Set DB = WS.OpenDatabase("server.dat", 0, False, ";pwd=" + Chr$(100) + Chr$(114) + Chr$(97) + Chr$(99) + Chr$(111))
        #Else
            CompactDatabase "server.tmp", "server.dat"
            Set DB = WS.OpenDatabase("server.dat")
        #End If
        Kill "server.tmp"
    Else
        frmLoading.lblStatus = "Creating Server Database.."
        frmLoading.lblStatus.Refresh
        CreateDatabase
        AddGod = True
    End If
    
    'DB.TableDefs.Delete ("Data")
    'CreateDataTable
    
    On Error Resume Next
    
    Err.Clear
    Set UserRS = DB.TableDefs("Accounts").OpenRecordset(dbOpenTable)
    If Err.number > 0 Then
        CreateAccountsTable
        Set UserRS = DB.TableDefs("Accounts").OpenRecordset(dbOpenTable)
    End If
    
    Err.Clear
    Set NPCRS = DB.TableDefs("NPCs").OpenRecordset(dbOpenTable)
    If Err.number > 0 Then
        CreateNPCsTable
        Set NPCRS = DB.TableDefs("NPCs").OpenRecordset(dbOpenTable)
    End If
    
    Err.Clear
    Set MonsterRS = DB.TableDefs("Monsters").OpenRecordset(dbOpenTable)
    If Err.number > 0 Then
        CreateMonstersTable
        Set MonsterRS = DB.TableDefs("Monsters").OpenRecordset(dbOpenTable)
    End If
    
    Err.Clear
    Set ObjectRS = DB.TableDefs("Objects").OpenRecordset(dbOpenTable)
    If Err.number > 0 Then
        CreateObjectsTable
        Set ObjectRS = DB.TableDefs("Objects").OpenRecordset(dbOpenTable)
    End If
    
    Err.Clear
    Set MapRS = DB.TableDefs("Maps").OpenRecordset(dbOpenTable)
    If Err.number > 0 Then
        CreateMapsTable
        Set MapRS = DB.TableDefs("Maps").OpenRecordset(dbOpenTable)
    End If
    
    Err.Clear
    Set MsgRS = DB.TableDefs("Messages").OpenRecordset(dbOpenTable)
    If Err.number > 0 Then
        CreateMessagesTable
        Set MsgRS = DB.TableDefs("Messages").OpenRecordset(dbOpenTable)
    End If
    
    Err.Clear
    Set PostRS = DB.TableDefs("Posts").OpenRecordset(dbOpenTable)
    If Err.number > 0 Then
        CreatePostsTable
        Set PostRS = DB.TableDefs("Posts").OpenRecordset(dbOpenTable)
    End If
    
    Err.Clear
    Set BanRS = DB.TableDefs("Bans").OpenRecordset(dbOpenTable)
    If Err.number > 0 Then
        CreateBansTable
        Set BanRS = DB.TableDefs("Bans").OpenRecordset(dbOpenTable)
    End If
    
    Err.Clear
    Set GuildRS = DB.TableDefs("Guilds").OpenRecordset(dbOpenTable)
    If Err.number > 0 Then
        CreateGuildsTable
        Set GuildRS = DB.TableDefs("Guilds").OpenRecordset(dbOpenTable)
    End If
    
    Err.Clear
    Set HallRS = DB.TableDefs("Halls").OpenRecordset(dbOpenTable)
    If Err.number > 0 Then
        CreateHallsTable
        Set HallRS = DB.TableDefs("Halls").OpenRecordset(dbOpenTable)
    End If
    
    Err.Clear
    Set ScriptRS = DB.TableDefs("Scripts").OpenRecordset(dbOpenTable)
    If Err.number > 0 Then
        CreateScriptsTable
        Set ScriptRS = DB.TableDefs("Scripts").OpenRecordset(dbOpenTable)
    End If
    
    Err.Clear
    Set GodRS = DB.TableDefs("Gods").OpenRecordset(dbOpenTable)
    If Err.number > 0 Then
        CreateGodTable
        Set GodRS = DB.TableDefs("Gods").OpenRecordset(dbOpenTable)
    End If
    
    On Error GoTo 0
    
    UserRS.Index = "User"
    ObjectRS.Index = "Number"
    NPCRS.Index = "Number"
    MonsterRS.Index = "Number"
    MapRS.Index = "Number"
    MsgRS.Index = "Number"
    PostRS.Index = "Number"
    BanRS.Index = "Number"
    GuildRS.Index = "Number"
    HallRS.Index = "Number"
    GodRS.Index = "Number"
    ScriptRS.Index = "Name"
    
    'If AddGod Then AddGodAccount "Bugaboo", "BUGGY", "BUGABOO"
    'If AddGod Then AddGodAccount "mbeutel", "MOZ", "NIGGA"
    
    'If ScriptRS.BOF = False Then
    '    ScriptRS.MoveFirst
    '    Dim bData As Byte
    '    While ScriptRS.EOF = False
    '        Open "SCRIPT.BAS" For Output As #1
    '        Print #1, ScriptRS!Source;
    '        Close #1
    '        MsgBox "Do your stuff!"
    '        St = ""
    '        Open "SCRIPT.BIN" For Binary As #1
    '        While Not EOF(1)
    '            Get #1, , bData
    '            If Not EOF(1) Then St = St + Chr$(bData)
    '        Wend
    '        Close #1
    '        ScriptRS.Edit
    '        ScriptRS!Data = St
    '        ScriptRS.Update
    '        ScriptRS.MoveNext
    '    Wend
    'End If
    
    CreateClassData
    
    frmLoading.lblStatus = "Loading World Data.."
    frmLoading.lblStatus.Refresh

ReloadData:
    Set DataRS = DB.TableDefs("Data").OpenRecordset(dbOpenTable)
    'Check if World Data exists
    If DataRS.RecordCount = 0 Then
        'Create default data
        With DataRS
            .AddNew
            !User = ""
            !Password = ""
            !BackupInterval = 5
            !MapResetTime = 120000
            !ObjResetTime = 300000
            !TimeAllowance = 0
            !MoneyObj = 1
            !MsgObj = 2
            !CharCounter = 0
            !MsgCounter = 0
            !MOTD = ""
            !Hour = 12
            !Day = 1
            !LastUpdate = CLng(Date)
            !ObjectData = ""
            !flags = String$(1024, 0)
            For A = 0 To 4
                .Fields("StartLocationX" + CStr(A)) = 5
                .Fields("StartLocationY" + CStr(A)) = 5
                .Fields("StartLocationMap" + CStr(A)) = 1
                .Fields("StartLocationMessage" + CStr(A)) = ""
            Next A
            For A = 1 To 8
                .Fields("StartingObj" + CStr(A)) = 0
                .Fields("StartingObjVal" + CStr(A)) = 0
            Next A
            .Update
            .MoveFirst
        End With
    End If
    
    On Error GoTo DataRSError
    
    'Load World Data
    LoadObjectData DataRS!ObjectData
    
    With World
        .BackupInterval = DataRS!BackupInterval
        .MapResetTime = DataRS!MapResetTime
        .ObjResetTime = DataRS!ObjResetTime
        .TimeAllowance = DataRS!TimeAllowance
        .MoneyObj = DataRS!MoneyObj
        .MsgCounter = DataRS!MsgCounter
        .CharCounter = DataRS!CharCounter
        .Hour = DataRS!Hour
        .LastUpdate = DataRS!LastUpdate
        .MOTD = DataRS!MOTD
        St = DataRS!flags
        For A = 0 To 255
            .Flag(A) = Asc(Mid$(St, A * 4 + 1, 1)) * 16777216 + Asc(Mid$(St, A * 4 + 2, 1)) * 65536 + Asc(Mid$(St, A * 4 + 3, 1)) * 256& + Asc(Mid$(St, A * 4 + 4, 1))
        Next A
        For A = 0 To 4
            .StartLocation(A).X = DataRS("StartLocationX" + CStr(A))
            .StartLocation(A).Y = DataRS("StartLocationY" + CStr(A))
            .StartLocation(A).Map = DataRS("StartLocationMap" + CStr(A))
            .StartLocation(A).Message = DataRS("StartLocationMessage" + CStr(A))
        Next A
        
        For A = 1 To 8
            .StartObjects(A) = DataRS("StartingObj" + CStr(A))
            .StartObjValues(A) = DataRS("StartingObjVal" + CStr(A))
        Next A
        
    End With
    User = DataRS!User
    Password = DataRS!Password
    
    On Error GoTo 0
    If bDataRSError Then
        MsgBox "There was an error loading the server options.  The database will be rebuilt, but some data may be lost.", vbOKOnly, TitleString
        DataRS.Close
        DB.TableDefs.Delete "Data"
        CreateDataTable
        bDataRSError = False
        GoTo ReloadData
    End If
    
    If World.LastUpdate > CurDate Or Abs(World.LastUpdate - CurDate) >= 30 Then
        If MsgBox("Please verify that your system date and time is set correctly -- click ok to go on.", vbOKCancel, TitleString) = vbCancel Then
            ShutdownServer
            End
        Else
            CurDate = CLng(Date)
        End If
    End If
    
    frmLoading.lblStatus = "Checking Accounts.."
    frmLoading.lblStatus.Refresh
    If UserRS.BOF = False Then
        UserRS.MoveFirst
        While UserRS.EOF = False
            'If UserRS!Bank >= 20000000 Then
            '    Debug.Print "User " + UserRS!User + ", name " + UserRS!Name + " has " + CStr(UserRS!Bank) + " gold in the bank."
            '    UserRS.Edit
            '    UserRS!Bank = 0
            '    UserRS.Update
            'End If
            'For A = 1 To 20
            '    If UserRS("InvObject" + CStr(A)) = 6 And UserRS("InvValue" + CStr(A)) >= 20000000 Then
            '        Debug.Print "User " + UserRS!User + ", name " + UserRS!Name + " has " + CStr(UserRS("InvValue" + CStr(A))) + " gold in hand."
            '        UserRS.Edit
            '        UserRS("InvObject" + CStr(A)) = 0
            '        UserRS("InvValue" + CStr(A)) = 0
            '        UserRS.Update
            '    End If
            'Next A
            'If UserRS!Name <> "" And UserRS!Class = 0 Then
            '    UserRS.Delete
            'End If
            If CurDate - UserRS!LastPlayed >= 30 Then
                DeleteAccount
            End If
            UserRS.MoveNext
        Wend
    End If
    
    frmLoading.lblStatus = "Loading Guilds.."
    frmLoading.lblStatus.Refresh
    If GuildRS.BOF = False Then
        GuildRS.MoveFirst
        While GuildRS.EOF = False
            A = GuildRS!number
            If A > 0 Then
                With Guild(A)
                    .Name = GuildRS!Name
                    .Bank = GuildRS!Bank
                    .DueDate = GuildRS!DueDate
                    .Hall = GuildRS!Hall
                    .Sprite = GuildRS!Sprite
                    For B = 0 To 19
                        .Member(B).Name = GuildRS("MemberName" + CStr(B))
                        .Member(B).Rank = GuildRS("MemberRank" + CStr(B))
                    Next B
                    For B = 0 To 4
                        .Declaration(B).Guild = GuildRS("DeclarationGuild" + CStr(B))
                        .Declaration(B).Type = GuildRS("DeclarationType" + CStr(B))
                    Next B
                    .Bookmark = GuildRS.Bookmark
                End With
            End If
            GuildRS.MoveNext
        Wend
    End If
    
    frmLoading.lblStatus = "Loading Halls.."
    frmLoading.lblStatus.Refresh
    If HallRS.BOF = False Then
        HallRS.MoveFirst
        While HallRS.EOF = False
            A = HallRS!number
            If A > 0 Then
                With Hall(A)
                    .Name = HallRS!Name
                    .Price = HallRS!Price
                    .Upkeep = HallRS!Upkeep
                    .StartLocation.Map = HallRS!StartLocationMap
                    .StartLocation.X = HallRS!StartLocationX
                    .StartLocation.Y = HallRS!StartLocationY
                End With
            End If
            HallRS.MoveNext
        Wend
    End If
    
    'Open "objects.txt" For Output As #1
    frmLoading.lblStatus = "Loading Objects.."
    frmLoading.lblStatus.Refresh
    If ObjectRS.BOF = False Then
        ObjectRS.MoveFirst
        While ObjectRS.EOF = False
            A = ObjectRS!number
            If A > 0 Then
                With Object(A)
                    .Name = ObjectRS!Name
                    .Picture = ObjectRS!Picture
                    .Type = ObjectRS!Type
                    .flags = ObjectRS!flags
                    .Data(0) = ObjectRS!Data1
                    .Data(1) = ObjectRS!Data2
                    .Data(2) = ObjectRS!Data3
                    .Data(3) = ObjectRS!Data4
                End With
            End If
            ObjectRS.MoveNext
        Wend
    End If
    'Close #1
    
    frmLoading.lblStatus = "Loading NPCs.."
    frmLoading.lblStatus.Refresh
    
    If NPCRS.BOF = False Then
        NPCRS.MoveFirst
        While NPCRS.EOF = False
            A = NPCRS!number
            If A > 0 Then
                With NPC(A)
                    .Name = NPCRS!Name
                    .JoinText = NPCRS!JoinText
                    .LeaveText = NPCRS!LeaveText
                    For B = 0 To 4
                        .SayText(B) = NPCRS("SayText" + CStr(B))
                    Next B
                    For B = 0 To 9
                        With .SaleItem(B)
                            .GiveObject = NPCRS("GiveObject" + CStr(B))
                            .GiveValue = NPCRS("GiveValue" + CStr(B))
                            .TakeObject = NPCRS("TakeObject" + CStr(B))
                            .TakeValue = NPCRS("TakeValue" + CStr(B))
                        End With
                    Next B
                    .flags = NPCRS!flags
                End With
            End If
            NPCRS.MoveNext
        Wend
    End If
    
    frmLoading.lblStatus = "Loading Monsters.."
    frmLoading.lblStatus.Refresh
    
    'Open "monsters.txt" For Output As #1
    If MonsterRS.BOF = False Then
        MonsterRS.MoveFirst
        While MonsterRS.EOF = False
            A = MonsterRS!number
            If A > 0 Then
                With Monster(A)
                    .Name = MonsterRS!Name
                    .Sprite = MonsterRS!Sprite
                    .HP = MonsterRS!HP
                    .Strength = MonsterRS!Strength
                    .Armor = MonsterRS!Armor
                    .Speed = MonsterRS!Speed
                    .Sight = MonsterRS!Sight
                    .Agility = MonsterRS!Agility
                    .flags = MonsterRS!flags
                    .Object(0) = MonsterRS!Object0
                    .Value(0) = MonsterRS!Value0
                    .Object(1) = MonsterRS!Object1
                    .Value(1) = MonsterRS!Value1
                    .Object(2) = MonsterRS!Object2
                    .Value(2) = MonsterRS!Value2
                    'Print #1, "Name=" + .Name + " HP=" + CStr(.HP) + " STRENGTH=" + CStr(.Strength) + " ARMOR=" + CStr(.Armor) + " AGILITY=" + CStr(.Agility)
                    'For A = 0 To 2
                    '    If .Object(A) > 0 Then
                    '        If Object(.Object(A)).Type = 6 Then
                    '            Print #1, "Drops " + CStr(.Value(A)) + " " + Object(.Object(A)).Name
                    '        Else
                    '            Print #1, "Drops " + Object(.Object(A)).Name
                    '        End If
                    '    End If
                    'Next A
                    'Print #1, ""
                End With
            End If
            MonsterRS.MoveNext
        Wend
    End If
    'Close #1
    
    frmLoading.lblStatus = "Loading Posts.."
    frmLoading.lblStatus.Refresh
    If PostRS.BOF = False Then
        PostRS.MoveFirst
        While PostRS.EOF = False
            A = PostRS!number
            If A > 0 Then
                With Post(A)
                    .Name = PostRS!Name
                    .flags = PostRS!flags
                    For B = 1 To 30
                        .Msg(B) = PostRS("Msg" + CStr(B))
                    Next B
                End With
            End If
            PostRS.MoveNext
        Wend
    End If
    
    frmLoading.lblStatus = "Loading Ban List.."
    frmLoading.lblStatus.Refresh
    If BanRS.BOF = False Then
        BanRS.MoveFirst
        While BanRS.EOF = False
            A = BanRS!number
            If A > 0 Then
                With Ban(A)
                    .Name = BanRS!Name
                    .ComputerID = DecipherString(BanRS!ComputerID)
                    .Reason = BanRS!Reason
                    .UnbanDate = BanRS!UnbanDate
                    .Banner = BanRS!Banner
                    .InUse = True
                End With
            End If
            BanRS.MoveNext
        Wend
    End If
    
    frmLoading.lblStatus = "Loading God Data.."
    frmLoading.lblStatus.Refresh
    If GodRS.BOF = False Then
        GodRS.MoveFirst
        While GodRS.EOF = False
            A = GodRS!number
            If A > 0 Then
                With GodData(A)
                    .User = GodRS!User
                    .ComputerID = DecipherString(GodRS!ComputerID)
                    .Access = GodRS!Access
                    .InUse = True
                End With
            End If
            GodRS.MoveNext
        Wend
    End If
    
    frmLoading.lblStatus = "Loading Maps.."
    frmLoading.lblStatus.Refresh
    If MapRS.BOF = False Then
        MapRS.MoveFirst
        While MapRS.EOF = False
            A = MapRS!number
            If A > 0 And A <= 2000 Then
                LoadMap A, MapRS!Data
                ResetMap A
            End If
            MapRS.MoveNext
        Wend
    End If
    
    frmLoading.lblStatus = "Initializing Sockets.."
    frmLoading.lblStatus.Refresh
    
    Load frmMain
    frmMain.Caption = TitleString + " [0]"
    Hook
    StartWinsock St
    
    'Listen for connections
    With LingerType
        .l_onoff = 1
        .l_linger = 0
    End With
                    
    ListeningSocket = ListenForConnect(5678, gHW, 1025)
    If ListeningSocket = INVALID_SOCKET Then
        MsgBox "Unable to create listening socket!1", vbOKOnly + vbExclamation, TitleString
        EndWinsock
        Unhook
        End
    End If
    If setsockopt(ListeningSocket, SOL_SOCKET, SO_LINGER, LingerType, 4) <> 0 Then
        MsgBox "Unable to create listening socket!", vbOKOnly + vbExclamation, TitleString
        EndWinsock
        Unhook
        End
    End If
    If setsockopt(ListeningSocket, IPPROTO_TCP, TCP_NODELAY, 1&, 4) <> 0 Then
        MsgBox "Unable to create listening socket!", vbOKOnly + vbExclamation, TitleString
        EndWinsock
        Unhook
        End
    End If
    If setsockopt(ListeningSocket, SOL_SOCKET, SO_RCVBUF, 4096&, 4) <> 0 Then
        MsgBox "Unable to create listening socket!", vbOKOnly + vbExclamation, TitleString
        EndWinsock
        Unhook
        End
    End If
    If setsockopt(ListeningSocket, SOL_SOCKET, SO_SNDBUF, 8192&, 4) <> 0 Then
        MsgBox "Unable to create listening socket!", vbOKOnly + vbExclamation, TitleString
        EndWinsock
        Unhook
        End
    End If
    
    Unload frmLoading
    
    'Connect to IRC Server
    With IRC
        .Socket = INVALID_SOCKET
        If .Disabled = False Then
            .NickChoice(1) = "ClassicServer"
            .NickChoice(2) = "Classic_Server"
            .NickChoice(3) = "_ClassicServer_"
            .NickChoice(4) = "_Classic_Server_"
            .NickChoice(5) = "_ClassicServer"
            .NickChoice(6) = "ClassicServer_"
            .Socket = ConnectSock(IRCServer, IRCPort, St, gHW, 1026, True)
            If Not .Socket = INVALID_SOCKET Then
                .Connecting = True
            Else
                .Connecting = False
            End If
        End If
    End With
    
    StartTimeStamp = GetTickCount
    
    frmMain.Show
    PrintLog ("The Odyssey Online Classic Server Version A" + CurrentClientVer + ".")
    Exit Sub
    
DataRSError:
    bDataRSError = True
    Resume Next
End Sub
Function NewMapMonster(MapNum As Long, MonsterNum As Long) As String
    Dim TX As Long, TY As Long, TriesLeft As Long
    Dim MonsterType As Long, MonsterFlags As Byte
    
    If Int(MonsterNum / 2) * 2 = MonsterNum Or ExamineBit(Map(MapNum).flags, 4) = True Then
        MonsterType = Map(MapNum).MonsterSpawn(Int(MonsterNum / 2)).Monster
        If MonsterType > 0 Then
            MonsterFlags = Monster(MonsterType).flags
            If (ExamineBit(MonsterFlags, 1) = False Or blnNight = False) And (ExamineBit(MonsterFlags, 2) = False Or blnNight = True) Then
                TX = Int(Rnd * 12)
                TY = Int(Rnd * 12)
                TriesLeft = 10
                While TriesLeft > 0 And Map(MapNum).Tile(TX, TY).Att > 0
                    TX = Int(Rnd) * 12
                    TY = Int(Rnd * 12)
                    TriesLeft = TriesLeft - 1
                Wend
                If TriesLeft > 0 Then
                    NewMapMonster = SpawnMapMonster(MapNum, MonsterNum, MonsterType, TX, TY)
                End If
            End If
        End If
    End If
End Function
Function NewMapObject(MapNum As Long, ObjectNum As Long, Value As Long, X As Long, Y As Long, Infinite As Boolean) As Long
    Dim A As Long
    If MapNum >= 1 Then
        A = FreeMapObj(MapNum)
        If A >= 0 Then
            With Map(MapNum).Object(A)
                .Object = ObjectNum
                .X = X
                .Y = Y
                If Infinite = True Then
                    .TimeStamp = 0
                Else
                    .TimeStamp = GetTickCount + Int(Rnd * 60000) - 30000
                End If
                Select Case Object(ObjectNum).Type
                    Case 1, 2, 3, 4 'Weapon, Shield, Armor, Helmut
                        .Value = CLng(Object(ObjectNum).Data(0)) * 10
                    Case 6 'Money
                        If Value < 1 Then Value = 1
                        .Value = Value
                    Case 8 'Ring
                        .Value = CLng(Object(ObjectNum).Data(1)) * 10
                    Case Else
                        .Value = 0
                End Select
                SendToMap MapNum, Chr$(14) + Chr$(A) + Chr$(ObjectNum) + Chr$(X) + Chr$(Y)
            End With
            NewMapObject = True
        End If
    End If
End Function
Function NewMessage(from As String, Subject As String, Body As String) As Long
    With World
        If .MsgCounter = 2147483647 Then
            .MsgCounter = 1
        Else
            .MsgCounter = .MsgCounter + 1
        End If
    End With
    With DataRS
        .Edit
        !MsgCounter = World.MsgCounter
        .Update
    End With
    With MsgRS
        .Seek "=", World.MsgCounter
        If .NoMatch Then
            .AddNew
            !number = World.MsgCounter
        Else
            .Edit
        End If
        !from = from
        !Subject = Subject
        !Body = Body
        .Update
    End With
End Function
Sub Partmap(Index As Long)
    Dim A As Long, MapNum As Long

    With Player(Index)
        MapNum = .Map
        If MapNum > 0 Then
            Parameter(0) = Index
            RunScript "PARTMAP" + CStr(MapNum)
            
            With Map(MapNum)
                .NumPlayers = .NumPlayers - 1
                For A = 0 To 5
                    With .Monster(A)
                        If .Target = Index And .Monster > 0 Then
                            .Target = 0
                            .Distance = Monster(.Monster).Sight
                        End If
                    End With
                Next A
                If .NumPlayers = 0 Then
                    .ResetTimer = GetTickCount
                End If
            End With
            SendToMapAllBut MapNum, Index, Chr$(9) + Chr$(Index)
            
            If .Socket <> INVALID_SOCKET Then
                A = Map(MapNum).NPC
                If A >= 1 Then
                    With NPC(A)
                        If .LeaveText <> "" Then
                            SendSocket Index, Chr$(88) + Chr$(A) + .LeaveText
                        End If
                    End With
                End If
            End If
            
            .Map = 0
        End If
    End With
End Sub
Function PlayerArmor(Index As Long, ByVal Damage As Long) As Long
    Dim A As Long, Armor As Long, ObjNum As Long
    With Player(Index)
        If .EquippedObject(2) > 0 Then
            'Has a shield
            If Int(Rnd * 100) < .Agility * 2 Then
                'Uses shield
                ObjNum = .EquippedObject(2)
                If .Inv(ObjNum).Object > 0 Then
                    Armor = Int((CSng(Object(.Inv(ObjNum).Object).Data(1)) / 255!) * CSng(Damage))
                    If Armor > Damage Then Armor = Damage
                    Damage = Damage - Armor
                    A = .Inv(ObjNum).Value - (Armor / 2)
                    If A <= 0 Then
                        'Object Is Destroyed
                        SendSocket Index, Chr$(57) + Chr$(ObjNum)
                        .Inv(ObjNum).Object = 0
                        .EquippedObject(2) = 0
                    Else
                        .Inv(ObjNum).Value = A
                    End If
                Else
                    .EquippedObject(2) = 0
                End If
            End If
        End If
        If Rnd > 0.2 Then
            'Body Shot
            If .EquippedObject(3) > 0 Then
                'Uses Armor
                ObjNum = .EquippedObject(3)
                If .Inv(ObjNum).Object > 0 Then
                    Armor = Object(.Inv(ObjNum).Object).Data(1)
                    If Armor > Damage Then Armor = Damage
                    Damage = Damage - Armor
                    A = .Inv(ObjNum).Value - Armor
                    If A <= 0 Then
                        'Object Is Destroyed
                        SendSocket Index, Chr$(57) + Chr$(ObjNum)
                        .Inv(ObjNum).Object = 0
                        .EquippedObject(3) = 0
                    Else
                        .Inv(ObjNum).Value = A
                    End If
                Else
                    .EquippedObject(3) = 0
                End If
            End If
        Else
            'Head Shot
            If .EquippedObject(4) > 0 Then
                'Uses Helmet
                ObjNum = .EquippedObject(4)
                If .Inv(ObjNum).Object > 0 Then
                    Armor = Object(.Inv(ObjNum).Object).Data(1)
                    A = .Inv(ObjNum).Value - Armor
                    If Armor > Damage Then Armor = Damage
                    Damage = Damage - Armor
                    If A <= 0 Then
                        'Object Is Destroyed
                        SendSocket Index, Chr$(57) + Chr$(ObjNum)
                        .Inv(ObjNum).Object = 0
                        .EquippedObject(4) = 0
                    Else
                        .Inv(ObjNum).Value = A
                    End If
                Else
                    .EquippedObject(4) = 0
                End If
            End If
        End If
    End With
    PlayerArmor = Damage
End Function
Function PlayerDamage(Index As Long) As Long
    Dim A As Long, Damage As Long, Weapon As Long, Modifier As Long, ObjNum As Long
    With Player(Index)
        If .EquippedObject(1) > 0 Then
            'Uses Weapon
            Weapon = .EquippedObject(1)
            If .Inv(Weapon).Object > 0 Then
                Damage = Int(.Strength / 5) + Object(.Inv(Weapon).Object).Data(1) + 1
                A = .Inv(Weapon).Value - (Damage / 2)
                If A <= 0 Then
                    'Object Is Destroyed
                    SendSocket Index, Chr$(57) + Chr$(Weapon)
                    .Inv(Weapon).Object = 0
                    .EquippedObject(1) = 0
                Else
                    .Inv(Weapon).Value = A
                End If
            Else
                .EquippedObject(1) = 0
            End If
        Else
            'Punches
            Damage = Int(.Strength / 5) + 1
        End If
        
        If .EquippedObject(5) > 0 Then
            'Has Ring
            ObjNum = .EquippedObject(5)
            If .Inv(ObjNum).Object > 0 Then
                If Object(.Inv(ObjNum).Object).Data(0) = 0 Then
                    Modifier = Object(.Inv(ObjNum).Object).Data(2)
                    A = .Inv(ObjNum).Value - Modifier
                    Damage = Damage + Modifier
                    If A <= 0 Then
                        'Object Is Destroyed
                        SendSocket Index, Chr$(57) + Chr$(ObjNum)
                        .Inv(ObjNum).Object = 0
                        .EquippedObject(5) = 0
                    Else
                        .Inv(ObjNum).Value = A
                    End If
                End If
            Else
                .EquippedObject(5) = 0
            End If
        End If
        
        If .EquippedObject(6) > 0 Then
            'Has Special EQ
            ObjNum = .EquippedObject(6)
            If .Inv(ObjNum).Object > 0 Then
                If Object(.Inv(ObjNum).Object).Data(0) = 0 Then
                    Modifier = Object(.Inv(ObjNum).Object).Data(2)
                    A = .Inv(ObjNum).Value - Modifier
                    Damage = Damage + Modifier
                    If A <= 0 Then
                        'Object Is Destroyed
                        SendSocket Index, Chr$(57) + Chr$(ObjNum)
                        .Inv(ObjNum).Object = 0
                        .EquippedObject(5) = 0
                    Else
                        .Inv(ObjNum).Value = A
                    End If
                End If
            Else
                .EquippedObject(5) = 0
            End If
        End If
        End With
    PlayerDamage = Damage
End Function
Sub PlayerDied(Index As Long)
    Dim A As Long, B As Long, C As Long, St1 As String, St2 As String
    Dim MapNum As Long
    
    With Player(Index)
        St1 = ""
        St2 = ""
        MapNum = .Map
        
        Parameter(0) = Index
        If RunScript("PLAYERDIE") = 0 Then
            For A = 1 To 20
                If .Inv(A).Object > 0 Then
                    C = 0
                    For B = 1 To 6
                        If .EquippedObject(B) = A Then
                            .EquippedObject(B) = 0
                            C = 1
                        End If
                    Next B
                    Parameter(0) = Index
                    If (C = 1 Or Rnd <= 0.3) And RunScript("DROPOBJ" + CStr(.Inv(A).Object)) = 0 Then
                        If Object(.Inv(A).Object).Data(3) = 0 Then '(Newbie Item )
                            B = FreeMapObj(MapNum)
                            If B >= 0 Then
                                Map(MapNum).Object(B).X = .X
                                Map(MapNum).Object(B).Y = .Y
                                Map(MapNum).Object(B).Object = .Inv(A).Object
                                Map(MapNum).Object(B).Value = .Inv(A).Value
                                Map(MapNum).Object(B).TimeStamp = GetTickCount + Int(Rnd * 60000) - 30000
                                St1 = St1 + DoubleChar(5) + Chr$(14) + Chr$(B) + Chr$(.Inv(A).Object) + Chr$(.X) + Chr$(.Y)
                            End If
                            .Inv(A).Object = 0
                            St2 = St2 + DoubleChar(2) + Chr$(18) + Chr$(A)
                        End If
                    End If
                End If
            Next A
            
            Partmap Index
            If St1 <> "" Then
                SendToMapRaw MapNum, St1
            End If
            
            If .Guild > 0 Then
                If Guild(.Guild).Hall >= 1 Then
                    A = 1
                Else
                    A = 0
                End If
            Else
                A = 0
            End If
            
            If A = 0 Then
                'Random Start Location
                A = Int(Rnd * 2)
                
                .Map = World.StartLocation(A).Map
                .X = World.StartLocation(A).X
                .Y = World.StartLocation(A).Y
                
                If World.StartLocation(A).Message <> "" Then
                    St2 = St2 + DoubleChar(2 + Len(World.StartLocation(A).Message)) + Chr$(56) + Chr$(15) + World.StartLocation(A).Message
                End If
            Else
                A = Guild(.Guild).Hall
                
                .Map = Hall(A).StartLocation.Map
                .X = Hall(A).StartLocation.X
                .Y = Hall(A).StartLocation.Y
            End If
    
            If St2 <> "" Then
                SendRaw Index, St2
            End If
            
            If .Status = 1 Then .Status = 0
    
            .Experience = Int((2 / 3) * .Experience)
            SendSocket Index, Chr$(60) + QuadChar(.Experience)
            
            If .Map < 1 Then .Map = 1
            If .Map > 2000 Then .Map = 2000
            If .Y > 11 Then .Y = 11
            If .X > 11 Then .X = 11
            
            JoinMap Index
        End If
    End With
End Sub
Sub ReadClientData(Index As Long)
    Dim St As String, SocketData As String, PacketLength As Long, PacketID As Long
    Dim A As Long, B As Long, C As Long, D As Long, E As Long, F As Long, G As Long, H As Long, I As Long, J As Long
    Dim Life As Long, Stun As Long
    Dim St1 As String, St2 As String
    Dim VarData As Variant
    Dim TInvNum As Byte, TInvVal As Long, MapNum As Long
    
    With Player(Index)
        MapNum = .Map
        SocketData = .SocketData + Receive(.Socket)
        .LastMsg = GetTickCount
        If GetTickCount > .FloodTimer Then
            .FloodTimer = GetTickCount
        End If
        .FloodTimer = .FloodTimer + 100 'MS Per Message
        If .FloodTimer - GetTickCount > 5000 Then
            BootPlayer Index, 0, "Flooding"
            Exit Sub
        End If
LoopRead:
        If Len(SocketData) >= 3 Then
            PacketLength = GetInt(Mid$(SocketData, 1, 2))
            If PacketLength >= 3072 And .Access < 10 Then
                Hacker Index, "C.1"
                Exit Sub
            End If
            If Len(SocketData) - 2 >= PacketLength Then
                St = Mid$(SocketData, 3, PacketLength)
                SocketData = Mid$(SocketData, PacketLength + 3)
                If PacketLength > 0 Then
                    PacketID = Asc(Mid$(St, 1, 1))
                    If Len(St) > 1 Then
                        St = Mid$(St, 2)
                    Else
                        St = ""
                    End If
                    Select Case .Mode
                        Case modeNotConnected
                            Select Case PacketID
                                Case 0 'New Account
                                    #If NewAccounts = False Then
                                        SendSocket Index, Chr$(0) + Chr$(0) + "Sorry, this is a private Odyssey server.  Please follow the instructions in the server description for obtaining an account."
                                        AddSocketQue Index
                                    #Else
                                        If .ClientVer = CurrentClientVer Then
                                            A = InStr(1, St, Chr$(0))
                                            If A > 1 And A < Len(St) Then
                                                St1 = Trim$(Mid$(St, 1, A - 1))
                                                B = Len(St1)
                                                If B >= 3 And B <= 15 And ValidName(St1) Then
                                                    UserRS.Index = "User"
                                                    UserRS.Seek "=", St1
                                                    If UserRS.NoMatch = True And GuildNum(St1) = 0 Then
                                                        UserRS.AddNew
                                                        UserRS!User = St1
                                                        'UserRS!ComputerID = EncryptString(.ComputerID)
                                                        .User = St1
                                                        St1 = Trim$(UCase$(Mid$(St, A + 1)))
                                                        If Len(St1) > 15 Then
                                                            UserRS!Password = DecipherString(Left$(St1, 15))
                                                        Else
                                                            UserRS!Password = DecipherString(St1)
                                                        End If
                                                        UserRS.Update
                                                        UserRS.Seek "=", .User
                                                        .Bookmark = UserRS.Bookmark
                                                        .Access = 0
                                                        .Class = 0
                                                        SavePlayerData Index
                                                        SendSocket Index, Chr$(2) 'New account created!
                                                        AddSocketQue Index
                                                    Else
                                                        SendSocket Index, Chr$(1) + Chr$(1) 'User Already Exists
                                                        AddSocketQue Index
                                                    End If
                                                Else
                                                    Hacker Index, "A.79"
                                                End If
                                            Else
                                                AddSocketQue Index
                                            End If
                                        Else
                                            SendSocket Index, Chr$(0) + Chr$(0) + "Your client is outdated, please visit " + DownloadSite + "! Download the newest update and unzip it into your Odyssey Online Classic folder."
                                            AddSocketQue Index
                                        End If
                                    #End If
                                    
                                Case 1 'Log on
                                    'SendSocket Index, Chr$(0) + Chr$(0) + "The server IP has changed -- please quit and re-load your client."
                                    'Exit Sub
                                    If .ClientVer = CurrentClientVer Then
                                        A = InStr(1, St, Chr$(0))
                                        If A > 1 And A < Len(St) Then
                                            .User = Mid$(St, 1, A - 1)
                                            UserRS.Index = "User"
                                            UserRS.Seek "=", .User
                                            If UserRS.NoMatch = False Then
                                                If UCase$(DecipherString(Mid$(St, A + 1))) = UCase$(DecipherString(UserRS!Password)) Then
                                                    B = 0
                                                    St1 = UCase$(.User)
                                                    For A = 1 To MaxUsers
                                                        If St1 = UCase$(Player(A).User) And A <> Index Then
                                                            B = 1
                                                            Exit For
                                                        End If
                                                    Next A
                                                    If B = 0 Then
                                                        'Account Data
                                                        .Access = UserRS!Access
                                                        .Bookmark = UserRS.Bookmark
                                                        
                                                        'Character Data
                                                        .CharNum = UserRS!CharNum
                                                        .Name = UserRS!Name
                                                        .Class = UserRS!Class
                                                        .Gender = UserRS!Gender
                                                        .Sprite = UserRS!Sprite
                                                        .Map = UserRS!Map
                                                        If .Map < 1 Then .Map = 1
                                                        If .Map > 2000 Then .Map = 2000
                                                        .X = UserRS!X
                                                        If .X > 11 Then .X = 11
                                                        .Y = UserRS!Y
                                                        If .Y > 11 Then .Y = 11
                                                        .D = UserRS!D
                                                        .desc = UserRS!desc
                                                        
                                                        'Character Vital Stats
                                                        .MaxHP = UserRS!MaxHP
                                                        .MaxEnergy = UserRS!MaxEnergy
                                                        .MaxMana = UserRS!MaxMana
                                                        .HP = UserRS!HP
                                                        .Energy = UserRS!Energy
                                                        .Mana = UserRS!Mana
                                                        
                                                        'Character Physical Stats
                                                        .Strength = UserRS!Strength
                                                        .Agility = UserRS!Agility
                                                        .Endurance = UserRS!Endurance
                                                        .Intelligence = UserRS!Intelligence
                                                        .level = UserRS!level
                                                        .Experience = UserRS!Experience
                                                        .StatPoints = UserRS!StatPoints
                                                        
                                                        'God Checking
                                                        #If GodChecking = True Then
                                                            If .Access > 0 Then
                                                                A = FindGodAccount(.ComputerID)
                                                                If A = 0 Then 'Doesn't Exist
                                                                    SendSocket Index, Chr$(0) + Chr$(5)
                                                                    CloseClientSocket Index
                                                                End If
                                                            End If
                                                        #End If
                                                        
                                                        'If it their ComputerID is Blank. Update it.
                                                        If UserRS!ComputerID = "" Then
                                                            UserRS.Edit
                                                            UserRS!ComputerID = EncryptString(.ComputerID)
                                                            UserRS.Update
                                                        End If
                                                        
                                                        'Inventory Data
                                                        For A = 1 To 20
                                                            .Inv(A).Object = UserRS.Fields("InvObject" + CStr(A))
                                                            .Inv(A).Value = UserRS.Fields("InvValue" + CStr(A))
                                                        Next A
                                                        For A = 1 To 6
                                                            .EquippedObject(A) = UserRS.Fields("EquippedObject" + CStr(A))
                                                        Next A
                                                        
                                                        'Mail Data
                                                        For A = 1 To 20
                                                            .Msg(A) = UserRS("Msg" + CStr(A))
                                                        Next A
                                                        
                                                        'Flags
                                                        St1 = UserRS!flags
                                                        For A = 0 To 127
                                                            With .Flag(A)
                                                                .Value = Asc(Mid$(St1, A * 8 + 1, 1)) * 16777216 + Asc(Mid$(St1, A * 8 + 2, 1)) * 65536 + Asc(Mid$(St1, A * 8 + 3, 1)) * 256& + Asc(Mid$(St1, A * 8 + 4, 1))
                                                                .ResetCounter = Asc(Mid$(St1, A * 8 + 5, 1)) * 16777216 + Asc(Mid$(St1, A * 8 + 6, 1)) * 65536 + Asc(Mid$(St1, A * 8 + 7, 1)) * 256& + Asc(Mid$(St1, A * 8 + 8, 1))
                                                            End With
                                                        Next A
                                                        
                                                        'Misc Data
                                                        .Bank = UserRS!Bank
                                                        .Status = UserRS!Status
                                                        If .Access > 0 Then .Status = 3
                                                        If .Access >= 10 Then .Status = 10
                                                        
                                                        .Guild = 0
                                                        .GuildRank = 0
                                                        
                                                        #If UseGuilds Then
                                                            'Find Guild
                                                            St1 = .Name
                                                            For A = 1 To 255
                                                                With Guild(A)
                                                                    If .Name <> "" Then
                                                                        For B = 0 To 19
                                                                            If .Member(B).Name = St1 Then
                                                                                Player(Index).Guild = A
                                                                                Player(Index).GuildRank = .Member(B).Rank
                                                                                Exit For
                                                                            End If
                                                                        Next B
                                                                    End If
                                                                End With
                                                            Next A
                                                        #End If
                                                        
                                                        For A = 1 To MaxPlayerTimers
                                                            .JoinRequest = 0
                                                            .ScriptTimer(A) = 0
                                                        Next A
                                                        
                                                        .Mode = modeConnected
                                                        
                                                        SendSocket Index, Chr$(23) + Chr$(Index) + Chr$(.Access) 'Send Misc Data
                                                        SendCharacterData Index
                                                        If World.MOTD <> "" Then
                                                            SendSocket Index, Chr$(4) + World.MOTD
                                                        End If
                                                    Else
                                                        SendSocket Index, Chr$(0) + Chr$(2) 'Account already in use
                                                        AddSocketQue Index
                                                    End If
                                                Else
                                                    SendSocket Index, Chr$(0) + Chr$(1) 'Invalid User/Password
                                                    AddSocketQue Index
                                                End If
                                            Else
                                                SendSocket Index, Chr$(0) + Chr$(1) 'Invalid User/Password
                                                AddSocketQue Index
                                            End If
                                        Else
                                            Hacker Index, "A.1"
                                        End If
                                    Else
                                        SendSocket Index, Chr$(0) + Chr$(0) + "Your client is outdated, please visit " + DownloadSite + "! Download the newest update and unzip it into your Odyssey Online Classic folder."
                                        AddSocketQue Index
                                    End If
                                    
                                Case 29 'Pong
                                    If Len(St) > 0 Then
                                        Hacker Index, "A.2"
                                    End If
                                    
                                Case 61 'Version and ComputerID
                                    If Len(St) >= 2 Then
                                        .ClientVer = Asc(Mid$(St, 1, 1))
                                        .ComputerID = Trim$(DecipherString(Mid$(St, 2, Len(St) - 1)))
                                        CheckBan Index, UCase$(.ComputerID)
                                    Else
                                        Hacker Index, "A.100"
                                    End If
                                    
                                    
                                Case Else
                                    Hacker Index, "B.1"
                            End Select
                        Case modeConnected
                            Select Case PacketID
                                Case 2 'Create New Character
                                    If Len(St) >= 4 Then
                                        A = InStr(3, St, Chr$(0))
                                        If A > 1 Then
                                            St1 = Trim$(Mid$(St, 3, A - 3))
                                            If Len(St1) >= 3 And Len(St1) <= 15 And ValidName(St1) Then
                                                UserRS.Index = "Name"
                                                UserRS.Seek "=", St1
                                                If (UserRS.NoMatch Or UCase$(.Name) = UCase$(St1)) And GuildNum(St1) = 0 And NPCNum(St1) = 0 Then
                                                    If .Class > 0 Then
                                                        UserRS.Bookmark = .Bookmark
                                                        DeleteCharacter
                                                    End If
                                                    .Class = Asc(Mid$(St, 1, 1))
                                                    If .Class < 1 Then .Class = 1
                                                    If .Class > 15 Then .Class = 15
                                                    .Gender = Asc(Mid$(St, 2, 1))
                                                    If .Gender > 1 Then .Gender = 1
                                                    .Sprite = .Class * 2 + .Gender - 1
                                                    .Name = St1
                                                    If A < Len(St) Then
                                                        St1 = Mid$(St, A + 1)
                                                        If Len(St1) > 255 Then
                                                            .desc = Left$(St1, 255)
                                                        Else
                                                            .desc = St1
                                                        End If
                                                    Else
                                                        .desc = ""
                                                    End If
                                                    .level = 1
                                                    .Bank = 0
                                                    .Status = 2
                                                    .MaxHP = Class(.Class).StartHP
                                                    .HP = .MaxHP
                                                    .MaxEnergy = Class(.Class).StartEnergy
                                                    .Energy = .MaxEnergy
                                                    .MaxMana = Class(.Class).StartMana
                                                    .Mana = .MaxMana
                                                    .Strength = Class(.Class).StartStrength
                                                    .Agility = Class(.Class).StartAgility
                                                    .Endurance = Class(.Class).StartEndurance
                                                    .Intelligence = Class(.Class).StartIntelligence
                                                    .Experience = 0
                                                    .StatPoints = 0
                                                    For A = 1 To 20
                                                        .Inv(A).Object = 0
                                                    Next A
                                                    For A = 1 To 6
                                                        .EquippedObject(A) = 0
                                                    Next A
                                                    For A = 0 To 127
                                                        With .Flag(A)
                                                            .Value = 0
                                                            .ResetCounter = 0
                                                        End With
                                                    Next A
                                                    .Map = World.StartLocation(0).Map
                                                    If .Map < 1 Then .Map = 1
                                                    If .Map > 2000 Then .Map = 2000
                                                    .X = World.StartLocation(0).X
                                                    If .X > 11 Then .X = 11
                                                    .Y = World.StartLocation(0).Y
                                                    If .Y > 11 Then .Y = 11
                                                    .Guild = 0
                                                    .GuildRank = 0
                                                    GiveStartingEQ Index
                                                    SavePlayerData Index
                                                    SendCharacterData Index
                                                Else
                                                    SendSocket Index, Chr$(13) 'Name already in use
                                                End If
                                            Else
                                                Hacker Index, "A.78"
                                            End If
                                        Else
                                            Hacker Index, "A.4"
                                        End If
                                    Else
                                        Hacker Index, "A.5"
                                    End If
                                    
                                Case 3 'Change Password
                                    If Len(St) > 0 Then
                                        UserRS.Bookmark = .Bookmark
                                        UserRS.Edit
                                        UserRS!Password = UCase$(EncryptString(St))
                                        UserRS.Update
                                        SendSocket Index, Chr$(5) 'Password Changed
                                    Else
                                        Hacker Index, "A.6"
                                    End If
                                    
                                Case 4 'Delete Account
                                    If Len(St) = 0 Then
                                        .Class = 0
                                        UserRS.Bookmark = .Bookmark
                                        DeleteAccount
                                        CloseClientSocket Index
                                    Else
                                        Hacker Index, "A.7"
                                    End If
                                    
                                Case 5 'Play
                                    If .Class > 0 Then
                                        If MapNum > 0 Then
                                            SendDataPacket Index, 1
                                        Else
                                            Hacker Index, "A.8"
                                        End If
                                    Else
                                        Hacker Index, "A.9"
                                    End If
                                    
                                Case 23 'Done receiving Data
                                    If Len(St) = 0 Then
                                        JoinGame Index
                                    Else
                                        Hacker Index, "A.10"
                                    End If
                                    
                                Case 24 'Send Next Packet
                                    If Len(St) = 1 Then
                                        SendDataPacket Index, Asc(Mid$(St, 1, 1))
                                    Else
                                        Hacker Index, "A.11"
                                    End If
                                    
                                Case 45 'Request Map
                                    If Len(St) = 2 And .Access > 0 Then
                                        A = Asc(Mid$(St, 1, 1)) * 256 + Asc(Mid$(St, 2, 1))
                                        MapRS.Seek "=", A
                                        If MapRS.NoMatch Then
                                            SendSocket Index, Chr$(21) + String$(1927, Chr$(0))
                                        Else
                                            SendSocket Index, Chr$(21) + MapRS!Data
                                        End If
                                    Else
                                        Hacker Index, "A.12"
                                    End If
                                    
                                    Case 29 'Pong
                                        If Len(St) > 0 Then
                                            Hacker Index, "A.13"
                                        End If
                                    
                                    Case Else
                                        Hacker Index, "B.2"
                            End Select
                        Case modePlaying
                            Select Case PacketID
                                Case 6 'Say
                                    .FloodTimer = .FloodTimer + 1000
                                    If Len(St) >= 1 And Len(St) <= 512 Then
                                        A = SysAllocStringByteLen(St, Len(St))
                                        Parameter(0) = Index
                                        Parameter(1) = A
                                        B = RunScript("MAPSAY" + CStr(MapNum))
                                        SysFreeString A
                                        
                                        If B = 0 Then
                                            SendToMapAllBut MapNum, Index, Chr$(11) + Chr$(Index) + St
                                            
                                            If Int(Rnd * 100) <= 8 Then
                                                A = Map(MapNum).NPC
                                                If A >= 1 Then
                                                    With NPC(A)
                                                        B = Int(Rnd * 5)
                                                        If .SayText(B) <> "" Then
                                                            SendToMap MapNum, Chr$(88) + Chr$(A) + .SayText(B)
                                                        End If
                                                    End With
                                                End If
                                            End If
                                        End If
                                    Else
                                        Hacker Index, "A.15"
                                    End If
                                    
                                Case 7 'Move
                                    If Len(St) = 5 Then
                                        If Asc(Mid$(St, 1, 1)) = .WalkCode Then
                                            I = .X
                                            J = .Y
                                            A = Asc(Mid$(St, 2, 1))
                                            B = Asc(Mid$(St, 3, 1))
                                            If Abs(A - CLng(.X)) + Abs(B - CLng(.Y)) <= 1 Then
                                                If .X <> A Or .Y <> B Then
                                                    .X = A
                                                    .Y = B
                                                    D = 1
                                                    If Asc(Mid$(St, 5, 1)) = 8 Then
                                                        'Use Energy
                                                        If .Energy > 0 Then .Energy = .Energy - 1
                                                        SendSocket Index, Chr$(47) + Chr$(.Energy)
                                                    End If
                                                Else
                                                    D = 0
                                                End If
                                                .D = Asc(Mid$(St, 4, 1))
                                                
                                                'Check if monsters notice
                                                If .Access = 0 Then
                                                    For C = 0 To 5
                                                        If Map(MapNum).Monster(C).Monster > 0 Then
                                                            If ExamineBit(Monster(Map(MapNum).Monster(C).Monster).flags, 3) = False Then
                                                                'Isn't Friendly
                                                                If ExamineBit(Monster(Map(MapNum).Monster(C).Monster).flags, 0) = False Or .Status = 1 Then
                                                                    With Map(MapNum).Monster(C)
                                                                        E = .X
                                                                        F = .Y
                                                                        G = .Distance
                                                                    End With
                                                                    H = Sqr((CLng(.X) - E) ^ 2 + (CLng(.Y) - F) ^ 2)
                                                                    If H <= G Then
                                                                        With Map(MapNum).Monster(C)
                                                                            If Index <> .Target Then
                                                                                Parameter(0) = Index
                                                                                If RunScript("MONSTERSEE" + CStr(.Monster)) = 0 Then
                                                                                    .Target = Index
                                                                                    .Distance = H
                                                                                End If
                                                                            Else
                                                                                .Distance = H
                                                                            End If
                                                                        End With
                                                                    End If
                                                                End If
                                                            End If
                                                        End If
                                                    Next C
                                                End If
                                                SendToMapAllBut MapNum, Index, Chr$(10) + Chr$(Index) + Chr$(.X) + Chr$(.Y) + Chr$(.D) + Mid$(St, 5, 1)
                                                Select Case Map(MapNum).Tile(.X, .Y).Att
                                                    Case 1 'Wall
                                                        If D = 1 Then
                                                            .X = I
                                                            .Y = J
                                                            Hacker Index, "E.1"
                                                        End If
                                                    Case 2 'Warp
                                                        A = Map(MapNum).Tile(.X, .Y).AttData(2)
                                                        B = Map(MapNum).Tile(.X, .Y).AttData(3)
                                                        C = CLng(Map(MapNum).Tile(.X, .Y).AttData(0)) * 256 + CLng(Map(MapNum).Tile(.X, .Y).AttData(1))
                                                        If A <= 11 And B <= 11 And C >= 1 And C <= 2000 Then
                                                            Partmap Index
                                                            .Map = C
                                                            .X = A
                                                            .Y = B
                                                            JoinMap Index
                                                        Else
                                                            AddSocketQue Index
                                                        End If
                                                    Case 3 'Key Door
                                                        Partmap Index
                                                        .Map = MapNum
                                                        .X = I
                                                        .Y = J
                                                        JoinMap Index
                                                    Case 4 'Door
                                                        C = FreeMapDoorNum(MapNum)
                                                        If C >= 0 Then
                                                            With Map(MapNum).Door(C)
                                                                .Att = 4
                                                                .X = A
                                                                .Y = B
                                                                .T = GetTickCount
                                                            End With
                                                            Map(MapNum).Tile(A, B).Att = 0
                                                            SendToMap MapNum, Chr$(36) + Chr$(C) + Chr$(A) + Chr$(B)
                                                        End If
                                                    Case 8 'Touch Plate
                                                        F = Map(MapNum).Tile(A, B).AttData(2)
                                                        If F > 0 Then
                                                            If .Guild > 0 Then
                                                                If .GuildRank >= 1 And Guild(.Guild).Hall = F Then
                                                                    G = 1
                                                                Else
                                                                    G = 0
                                                                End If
                                                            Else
                                                                G = 0
                                                            End If
                                                        Else
                                                            G = 1
                                                        End If
                                                        If G = 1 Then
                                                            D = Map(MapNum).Tile(A, B).AttData(0)
                                                            E = Map(MapNum).Tile(A, B).AttData(1)
                                                            If D <= 11 And E <= 11 Then
                                                                If Map(MapNum).Tile(D, E).Att > 0 Then
                                                                    C = FreeMapDoorNum(MapNum)
                                                                    If C >= 0 Then
                                                                        With Map(MapNum).Door(C)
                                                                            .Att = Map(MapNum).Tile(D, E).Att
                                                                            .X = D
                                                                            .Y = E
                                                                            .T = GetTickCount
                                                                        End With
                                                                        Map(MapNum).Tile(D, E).Att = 0
                                                                        SendToMap MapNum, Chr$(36) + Chr$(C) + Chr$(D) + Chr$(E)
                                                                    End If
                                                                End If
                                                            End If
                                                        End If
                                                    Case 9 'Damage Tile
                                                    Case 11 'Script
                                                        If D = 1 Then
                                                            Parameter(0) = Index
                                                            RunScript "MAP" + CStr(MapNum) + "_" + CStr(A) + "_" + CStr(B)
                                                        End If
                                                End Select
                                            Else
                                                Hacker Index, "D.1"
                                            End If
                                        End If
                                    Else
                                        Hacker Index, "A.16"
                                    End If
                                    
                                Case 8 'Pick up map object
                                    If Len(St) = 1 Then
                                        A = Asc(Mid$(St, 1, 1)) 'map object #
                                        If A <= 49 Then
                                            C = Map(MapNum).Object(A).Object
                                            If C > 0 Then
                                                If Map(MapNum).Object(A).X = .X And Map(MapNum).Object(A).Y = .Y Then
                                                    Parameter(0) = Index
                                                    If RunScript("GETOBJ" + CStr(C)) = 0 Then
                                                        If Object(C).Type = 6 Then
                                                            'Money
                                                            B = FindInvObject(Index, C)
                                                            If B = 0 Then
                                                                B = FreeInvNum(Index)
                                                                E = 0
                                                            Else
                                                                E = 1
                                                            End If
                                                        Else
                                                            B = FreeInvNum(Index)
                                                            E = 0
                                                        End If
                                                        If B > 0 Then
                                                            With .Inv(B)
                                                                .Object = C
                                                                If E = 1 Then
                                                                    If CDbl(.Value) + CDbl(Map(MapNum).Object(A).Value) > 2147483647# Then
                                                                        D = 2147483647
                                                                    Else
                                                                        D = .Value + Map(MapNum).Object(A).Value
                                                                    End If
                                                                Else
                                                                    D = Map(MapNum).Object(A).Value
                                                                End If
                                                                .Value = D
                                                            End With
                                                            Map(MapNum).Object(A).Object = 0
                                                            SendToMap MapNum, Chr$(15) + Chr$(A) 'Erase Map Obj
                                                            SendSocket Index, Chr$(17) + Chr$(B) + Chr$(C) + QuadChar(D) 'New Inv Obj
                                                        Else
                                                            SendSocket Index, Chr$(16) + Chr$(1) 'Inv Full
                                                        End If
                                                    End If
                                                Else
                                                    SendSocket Index, Chr$(16) + Chr$(3) 'No such object
                                                End If
                                            Else
                                                SendSocket Index, Chr$(16) + Chr$(3) 'No such object
                                            End If
                                        End If
                                    Else
                                        Hacker Index, "A.17"
                                    End If
                                    
                                Case 9 'Drop Object
                                    If Len(St) = 5 Then
                                        A = Asc(Mid$(St, 1, 1))
                                        If A >= 1 And A <= 20 Then
                                            B = .Inv(A).Object
                                            Parameter(0) = Index
                                            If RunScript("DROPOBJ" + CStr(B)) = 0 Then
                                                If B > 0 Then
                                                    C = FreeMapObj(MapNum)
                                                    If C >= 0 Then
                                                        For E = 1 To 6
                                                            If .EquippedObject(E) = A Then .EquippedObject(E) = 0
                                                        Next E
                                                        F = 0
                                                        If Object(B).Type = 6 Then
                                                            E = Asc(Mid$(St, 2, 1)) * 16777216 + Asc(Mid$(St, 3, 1)) * 65536 + Asc(Mid$(St, 4, 1)) * 256& + Asc(Mid$(St, 5, 1))
                                                            If E < .Inv(A).Value Then
                                                                D = E
                                                                .Inv(A).Value = .Inv(A).Value - E
                                                                F = 1
                                                            Else
                                                                D = .Inv(A).Value
                                                                .Inv(A).Object = 0
                                                            End If
                                                        Else
                                                            D = .Inv(A).Value
                                                            .Inv(A).Object = 0
                                                        End If
                                                        With Map(MapNum).Object(C)
                                                            .Object = B
                                                            .Value = D
                                                            .TimeStamp = GetTickCount() + Int(Rnd * 60000) - 30000
                                                        End With
                                                        Map(MapNum).Object(C).X = .X
                                                        Map(MapNum).Object(C).Y = .Y
                                                        SendToMap MapNum, Chr$(14) + Chr$(C) + Chr$(B) + Chr$(.X) + Chr$(.Y) 'New Map Obj
                                                        If F = 0 Then
                                                            SendSocket Index, Chr$(18) + Chr$(A) 'Erase Inv Obj
                                                        Else
                                                            SendSocket Index, Chr$(17) + Chr$(A) + Chr$(B) + QuadChar(.Inv(A).Value) 'Update inv obj
                                                        End If
                                                    Else
                                                        SendSocket Index, Chr$(16) + Chr$(2) 'Map full
                                                    End If
                                                Else
                                                    SendSocket Index, Chr$(16) + Chr$(3) 'No such object
                                                End If
                                            End If
                                        End If
                                    Else
                                        Hacker Index, "A.18"
                                    End If
                                    
                                Case 10 'Use Object
                                    If Len(St) = 1 Then
                                        A = Asc(Mid$(St, 1, 1))
                                        If A >= 1 And A <= 20 Then
                                            If .Inv(A).Object > 0 Then
                                                Parameter(0) = Index
                                                If RunScript("USEOBJ" + CStr(.Inv(A).Object)) = 0 Then
                                                    If .Inv(A).Object > 0 Then
                                                        Select Case Object(.Inv(A).Object).Type
                                                            Case 1 'Weapon
                                                                B = 1
                                                                C = 0
                                                            Case 2 'Shield
                                                                B = 2
                                                                C = 0
                                                            Case 3 'Armor
                                                                B = 3
                                                                C = 0
                                                            Case 4 'Helmut
                                                                B = 4
                                                                C = 0
                                                            Case 5 'Potion
                                                                B = Object(.Inv(A).Object).Data(1)
                                                                Select Case Object(.Inv(A).Object).Data(0)
                                                                    Case 0 'Gives HP
                                                                        If CLng(.HP) + B < .MaxHP Then
                                                                            .HP = .HP + B
                                                                        Else
                                                                            .HP = .MaxHP
                                                                        End If
                                                                        SendSocket Index, Chr$(46) + Chr$(.HP)
                                                                    Case 1 'Takes HP
                                                                        If CLng(.HP) - B > 0 Then
                                                                            .HP = .HP - B
                                                                        Else
                                                                            .HP = 0
                                                                        End If
                                                                        SendSocket Index, Chr$(46) + Chr$(.HP)
                                                                    Case 2 'Gives Mana
                                                                        If CLng(.Mana) + B < .MaxMana Then
                                                                            .Mana = .Mana + B
                                                                        Else
                                                                            .Mana = .MaxMana
                                                                        End If
                                                                        SendSocket Index, Chr$(48) + Chr$(.Mana)
                                                                    Case 3 'Takes Mana
                                                                        If CLng(.Mana) - B > 0 Then
                                                                            .Mana = .Mana - B
                                                                        Else
                                                                            .Mana = 0
                                                                        End If
                                                                        SendSocket Index, Chr$(48) + Chr$(.Mana)
                                                                    Case 4 'Gives Energy
                                                                        If CLng(.Energy) + B < .MaxEnergy Then
                                                                            .Energy = .Energy + B
                                                                        Else
                                                                            .Energy = .MaxEnergy
                                                                        End If
                                                                        SendSocket Index, Chr$(47) + Chr$(.Energy)
                                                                    Case 5 'Takes Energy
                                                                        If CLng(.Energy) - B > 0 Then
                                                                            .Energy = .Energy - B
                                                                        Else
                                                                            .Energy = 0
                                                                        End If
                                                                        SendSocket Index, Chr$(47) + Chr$(.Energy)
                                                                End Select
                                                                B = 0
                                                                C = 1
                                                            Case 7 'Key
                                                                Select Case .D
                                                                    Case 0 'Up
                                                                        C = .X
                                                                        D = CLng(.Y) - 1
                                                                    Case 1 'Down
                                                                        C = .X
                                                                        D = .Y + 1
                                                                    Case 2 'Left
                                                                        C = CLng(.X) - 1
                                                                        D = .Y
                                                                    Case 3 'Right
                                                                        C = .X + 1
                                                                        D = .Y
                                                                End Select
                                                                If C >= 0 And C <= 11 And D >= 0 And D <= 11 Then
                                                                    If Map(MapNum).Tile(C, D).Att = 3 And Map(MapNum).Tile(C, D).AttData(0) = .Inv(A).Object Then
                                                                        E = FreeMapDoorNum(MapNum)
                                                                        If E >= 0 Then
                                                                            With Map(MapNum).Door(E)
                                                                                .Att = 3
                                                                                .X = C
                                                                                .Y = D
                                                                                .T = GetTickCount
                                                                            End With
                                                                            Map(MapNum).Tile(C, D).Att = 0
                                                                            SendToMap MapNum, Chr$(36) + Chr$(E) + Chr$(C) + Chr$(D)
                                                                            If Object(.Inv(A).Object).Data(0) = 0 Then
                                                                                C = 1
                                                                            Else
                                                                                C = 0
                                                                            End If
                                                                        End If
                                                                    Else
                                                                        C = 0
                                                                    End If
                                                                Else
                                                                    C = 0
                                                                End If
                                                                B = 0
                                                              
                                                            Case 8 'Ring
                                                                B = 5
                                                                C = 0
                                                            
                                                            Case 9 'Guild Deed
                                                            #If UseGuilds = True Then
                                                                If Player(Index).Guild = 0 Then
                                                                    B = 0
                                                                    C = 1
                                                                    SendSocket Index, Chr$(97) + Chr$(1)
                                                                Else
                                                                    B = 0
                                                                    C = 0
                                                                    SendSocket Index, Chr$(97) + Chr$(0)
                                                                End If
                                                            #Else
                                                                SendSocket Index, Chr$(97) + Chr$(2)
                                                            #End If
                                                            
                                                            Case Else
                                                                B = 0
                                                                C = 0
                                                                SendSocket Index, Chr$(16) + Chr$(8) 'You cannot use that
                                                        End Select
                                                        If B > 0 Then
                                                            'Equip Item
                                                            If .EquippedObject(B) > 0 Then
                                                                SendSocket Index, Chr$(20) + Chr$(.EquippedObject(B)) 'Stop Using Object
                                                            End If
                                                            .EquippedObject(B) = A
                                                            SendSocket Index, Chr$(19) + Chr$(A) + Chr$(B) 'Use Object
                                                        End If
                                                        If C > 0 Then
                                                            'Destroy Item
                                                            .Inv(A).Object = 0
                                                            SendSocket Index, Chr$(18) + Chr$(A) 'Remove inv object
                                                        End If
                                                    End If
                                                End If
                                            Else
                                                SendSocket Index, Chr$(16) + Chr$(3) 'No such object
                                            End If
                                        Else
                                            Hacker Index, "A.19"
                                        End If
                                    Else
                                        Hacker Index, "A.20"
                                    End If
                                    
                                Case 11 'Stop Using Object
                                    If Len(St) = 1 Then
                                        A = Asc(Mid$(St, 1, 1))
                                        If A >= 1 And A <= 6 Then
                                            B = .EquippedObject(A)
                                            If B > 0 Then
                                                SendSocket Index, Chr$(20) + Chr$(B) 'Stop Using Object
                                                .EquippedObject(A) = 0
                                            Else
                                                SendSocket Index, Chr$(16) + Chr$(3) 'No such object
                                            End If
                                        Else
                                            Hacker Index, "A.21"
                                        End If
                                    End If
                                    
                                Case 12 'Upload Map
                                    If Len(St) = 1927 And .Access >= 5 Then
                                        MapRS.Seek "=", MapNum
                                        If MapRS.NoMatch Then
                                            MapRS.AddNew
                                            MapRS!number = MapNum
                                        Else
                                            MapRS.Edit
                                        End If
                                        MapRS!Data = St
                                        MapRS.Update
                                        LoadMap MapNum, St
                                        For A = 0 To 9
                                            Map(MapNum).Door(A).Att = 0
                                        Next A
                                        For A = 1 To MaxUsers
                                            With Player(A)
                                                If .Mode = modePlaying And .Map = MapNum Then
                                                    Partmap A
                                                    .Map = MapNum
                                                    JoinMap A
                                                End If
                                            End With
                                        Next A
                                    Else
                                        Hacker Index, "A.22"
                                    End If
                                    
                                Case 13 'Exit Map
                                    If Len(St) = 1 Then
                                        Select Case Asc(Mid$(St, 1, 1))
                                            Case 0
                                                If Map(MapNum).ExitUp > 0 Then
                                                    If .Y = 0 Then
                                                        Partmap Index
                                                        .Map = Map(MapNum).ExitUp
                                                        .Y = 11
                                                        JoinMap Index
                                                    Else
                                                        Hacker Index, "D.1"
                                                    End If
                                                Else
                                                    Partmap Index
                                                    .Map = MapNum
                                                    JoinMap Index
                                                End If
                                            Case 1
                                                If Map(MapNum).ExitDown > 0 Then
                                                    If .Y = 11 Then
                                                        Partmap Index
                                                        .Map = Map(MapNum).ExitDown
                                                        .Y = 0
                                                        JoinMap Index
                                                    Else
                                                        Hacker Index, "D.1"
                                                    End If
                                                Else
                                                    Partmap Index
                                                    .Map = MapNum
                                                    JoinMap Index
                                                End If
                                            Case 2
                                                If Map(MapNum).ExitLeft > 0 Then
                                                    If .X = 0 Then
                                                        Partmap Index
                                                        .Map = Map(MapNum).ExitLeft
                                                        .X = 11
                                                        JoinMap Index
                                                    Else
                                                        Hacker Index, "D.1"
                                                    End If
                                                Else
                                                    Partmap Index
                                                    .Map = MapNum
                                                    JoinMap Index
                                                End If
                                            Case 3
                                                If Map(MapNum).ExitRight > 0 Then
                                                    If .X = 11 Then
                                                        Partmap Index
                                                        .Map = Map(MapNum).ExitRight
                                                        .X = 0
                                                        JoinMap Index
                                                    Else
                                                        Hacker Index, "D.1"
                                                    End If
                                                Else
                                                    Partmap Index
                                                    .Map = MapNum
                                                    JoinMap Index
                                                End If
                                        End Select
                                    Else
                                        Hacker Index, "A.23"
                                    End If
                                    
                                Case 14 'Tell
                                    .FloodTimer = .FloodTimer + 1000
                                    If Len(St) >= 2 And Len(St) <= 513 Then
                                        A = Asc(Mid$(St, 1, 1))
                                        If A >= 1 And A <= MaxUsers Then
                                            If Player(A).Mode = modePlaying Then
                                                SendSocket A, Chr$(25) + Chr$(Index) + Mid$(St, 2)
                                                If .Mana > 2 Then
                                                    .Mana = .Mana - 2
                                                Else
                                                    .Mana = 0
                                                End If
                                                SendSocket Index, Chr$(48) + Chr$(.Mana)
                                            End If
                                        End If
                                    Else
                                        Hacker Index, "A.24"
                                    End If
                                    
                                Case 15 'Broadcast
                                    .FloodTimer = .FloodTimer + 1500
                                    If Len(St) >= 1 And Len(St) <= 512 Then
                                        A = SysAllocStringByteLen(St, Len(St))
                                        Parameter(0) = Index
                                        Parameter(1) = A
                                        SysFreeString A
                                        
                                        If RunScript("BROADCAST") = 0 Then
                                            SendAllBut Index, Chr$(26) + Chr$(Index) + St
                                            PrintLog .Name + ": " + St
                                            If .Mana > 5 Then
                                                .Mana = .Mana - 5
                                            Else
                                                .Mana = 0
                                            End If
                                            SendSocket Index, Chr$(48) + Chr$(.Mana)
                                        End If
                                    Else
                                        Hacker Index, "A.25"
                                    End If
                                    
                                Case 16 'Emote
                                    .FloodTimer = .FloodTimer + 1000
                                    If Len(St) >= 1 And Len(St) <= 512 Then
                                        A = SysAllocStringByteLen(St, Len(St))
                                        Parameter(0) = Index
                                        Parameter(1) = A
                                        If RunScript("MAPSAY" + CStr(MapNum)) = 0 Then
                                            SendToMapAllBut MapNum, Index, Chr$(27) + Chr$(Index) + St
                                        End If
                                        SysFreeString A
                                    Else
                                        Hacker Index, "A.26"
                                    End If
                                    
                                Case 17 'Yell
                                    .FloodTimer = .FloodTimer + 1200
                                    If Len(St) >= 1 And Len(St) <= 512 Then
                                        SendToMapAllBut MapNum, Index, Chr$(28) + Chr$(Index) + St
                                        A = MapNum
                                        With Map(MapNum)
                                            B = .ExitUp
                                            C = .ExitDown
                                            D = .ExitLeft
                                            E = .ExitRight
                                        End With
                                        If B <> MapNum Then SendToMap B, Chr$(28) + Chr$(Index) + St
                                        If C <> MapNum And C <> B Then SendToMap C, Chr$(28) + Chr$(Index) + St
                                        If D <> MapNum And D <> B And D <> C Then SendToMap D, Chr$(28) + Chr$(Index) + St
                                        If E <> MapNum And E <> B And E <> C And E <> D Then SendToMap E, Chr$(28) + Chr$(Index) + St
                                    Else
                                        Hacker Index, "A.27"
                                    End If
                                    
                                Case 18 'God Commands
                                    If .Access > 0 And Len(St) >= 1 Then
                                        Select Case Asc(Mid$(St, 1, 1))
                                            Case 0 'Server Message
                                                If Len(St) >= 2 Then
                                                    SendAll Chr$(30) + "[" + .Name + "] " + Mid$(St, 2)
                                                Else
                                                    Hacker Index, "A.28"
                                                End If
                                            
                                            Case 1 'Warp
                                                If Len(St) = 5 Then
                                                    A = Asc(Mid$(St, 2, 1)) * 256 + Asc(Mid$(St, 3, 1))
                                                    B = Asc(Mid$(St, 4, 1))
                                                    C = Asc(Mid$(St, 5, 1))
                                                    If A >= 1 And A <= 2000 And B <= 11 And C <= 11 Then
                                                        Partmap Index
                                                        .Map = A
                                                        .X = B
                                                        .Y = C
                                                        JoinMap Index
                                                    End If
                                                Else
                                                    Hacker Index, "A.29"
                                                End If
                                                
                                            Case 2 'WarpMe
                                                If Len(St) = 2 And .Access >= 2 Then
                                                    A = Asc(Mid$(St, 2, 1))
                                                    If A >= 1 And A <= MaxUsers And A <> Index Then
                                                        If Player(A).Mode = modePlaying Then
                                                            Partmap Index
                                                            .Map = Player(A).Map
                                                            .X = Player(A).X
                                                            .Y = Player(A).Y
                                                            JoinMap Index
                                                        End If
                                                    End If
                                                Else
                                                    Hacker Index, "A.30"
                                                End If
                                                
                                            Case 3 'WarpPlayer
                                                If Len(St) = 6 And .Access >= 3 Then
                                                    A = Asc(Mid$(St, 2, 1))
                                                    B = Asc(Mid$(St, 3, 1)) * 256 + Asc(Mid$(St, 4, 1))
                                                    C = Asc(Mid$(St, 5, 1))
                                                    D = Asc(Mid$(St, 6, 1))
                                                    If A >= 1 And A <= MaxUsers And B >= 1 And B <= 2000 And C <= 11 And D <= 11 Then
                                                        With Player(A)
                                                            If .Mode = modePlaying Then
                                                                Partmap A
                                                                .Map = B
                                                                .X = C
                                                                .Y = D
                                                                JoinMap A
                                                            End If
                                                        End With
                                                    End If
                                                Else
                                                    Hacker Index, "A.31"
                                                End If
                                            
                                            Case 4 'Set MOTD
                                                If Len(St) > 1 And .Access >= 9 Then
                                                    World.MOTD = Mid$(St, 2)
                                                    DataRS.Edit
                                                    DataRS!MOTD = World.MOTD
                                                    DataRS.Update
                                                Else
                                                    Hacker Index, "A.32"
                                                End If
                                                
                                            Case 5 'Disband Guild
                                                If Len(St) = 2 And .Access >= 10 Then
                                                    A = Asc(Mid$(St, 2, 1))
                                                    DeleteGuild A, 3
                                                Else
                                                    Hacker Index, "A.33"
                                                End If
                    
                                            Case 6 'Set sprite
                                                If Len(St) = 3 And .Access >= 8 Then
                                                    A = Asc(Mid$(St, 2, 1))
                                                    B = Asc(Mid$(St, 3, 1))
                                                    If A >= 1 And A <= MaxUsers And B <= 255 Then
                                                        With Player(A)
                                                            If .Mode = modePlaying Then
                                                                If B = 0 Then
                                                                    .Sprite = .Class * 2 + .Gender - 1
                                                                Else
                                                                    .Sprite = B
                                                                End If
                                                                SendAll Chr$(63) + Chr$(A) + Chr$(.Sprite)
                                                            End If
                                                        End With
                                                    End If
                                                Else
                                                    Hacker Index, "A.34"
                                                End If
                                                
                                            Case 7 'Set name
                                                If Len(St) >= 3 And Len(St) <= 17 And .Access >= 8 Then
                                                    A = Asc(Mid$(St, 2, 1))
                                                    If A >= 1 And A <= MaxUsers Then
                                                        With Player(A)
                                                            If .Mode = modePlaying Then
                                                                .Name = Mid$(St, 3)
                                                                SendAll Chr$(64) + Chr$(A) + .Name
                                                            End If
                                                        End With
                                                    End If
                                                Else
                                                    Hacker Index, "A.35"
                                                End If
                                            Case 8 'Resetmap
                                                If Len(St) = 1 Then
                                                    ResetMap CLng(.Map)
                                                Else
                                                    Hacker Index, "A.84"
                                                End If
                                            Case 9 'Boot
                                                If Len(St) >= 2 And .Access >= 4 Then
                                                    A = Asc(Mid$(St, 2, 1))
                                                    If A >= 1 And A <= MaxUsers Then
                                                        If Player(A).Access < 11 Then
                                                            BootPlayer A, Index, Mid$(St, 3)
                                                        End If
                                                    End If
                                                Else
                                                    Hacker Index, "A.36"
                                                End If
                                            Case 10 'Ban
                                                If Len(St) >= 3 And .Access >= 4 Then
                                                    A = Asc(Mid$(St, 2, 1))
                                                    If A >= 1 And A <= MaxUsers Then
                                                        If Player(A).Access < 11 Then
                                                            If BanPlayer(A, Index, Asc(Mid$(St, 3, 1)), Mid$(St, 4), Player(Index).Name) = False Then
                                                                SendSocket Index, Chr$(16) + Chr$(13) 'Ban list full
                                                            End If
                                                        End If
                                                    End If
                                                Else
                                                    Hacker Index, "A.37"
                                                End If
                                                
                                            Case 11 'Remove Ban
                                                If Len(St) = 2 And .Access >= 4 Then
                                                    A = Asc(Mid$(St, 2, 1))
                                                    If A >= 1 And A <= 50 Then
                                                        With Ban(A)
                                                            .ComputerID = ""
                                                            .InUse = False
                                                        End With
                                                        BanRS.Seek "=", A
                                                        If BanRS.NoMatch = False Then
                                                            BanRS.Delete
                                                        End If
                                                        SendAll Chr$(56) + Chr$(15) + Ban(A).Name + " has been unbanned by " + .Name + "."
                                                    End If
                                                Else
                                                    Hacker Index, "A.38"
                                                End If
                                                
                                            Case 12 'List Bans
                                                If Len(St) = 1 And .Access >= 4 Then
                                                    St1 = ""
                                                    For A = 1 To 50
                                                        With Ban(A)
                                                            If .InUse = True Then
                                                                St1 = St1 + DoubleChar(2 + Len(.Name)) + Chr$(69) + Chr$(A) + .Name
                                                            End If
                                                        End With
                                                    Next A
                                                    St1 = St1 + DoubleChar(1) + Chr$(69)
                                                    SendRaw Index, St1
                                                Else
                                                    Hacker Index, "A.39"
                                                End If
                                                
                                            Case 13 'Shutdown Server
                                                If Len(St) = 1 And .Access >= 10 Then
                                                    ShutdownServer
                                                    End
                                                Else
                                                    Hacker Index, "A.40"
                                                End If
                                                
                                            Case 14 'Chat
                                                If Len(St) >= 2 Then
                                                    SendToGodsAllBut Index, Chr$(90) + Chr$(Index) + Mid$(St, 2)
                                                Else
                                                    Hacker Index, "A.83"
                                                End If
                                                
                                            Case 15 'Set Guild Sprite
                                                If Len(St) = 3 And .Access >= 10 Then
                                                    A = Asc(Mid$(St, 2, 1))
                                                    B = Asc(Mid$(St, 3, 1))
                                                    If A >= 1 Then
                                                        UserRS.Index = "Name"
                                                        With Guild(A)
                                                            If .Name <> "" Then
                                                                .Sprite = B
                                                                GuildRS.Bookmark = .Bookmark
                                                                GuildRS.Edit
                                                                GuildRS!Sprite = B
                                                                GuildRS.Update
                                                                
                                                                For C = 0 To 19
                                                                    With .Member(C)
                                                                        If .Name <> "" Then
                                                                            D = FindPlayer(.Name)
                                                                            If D > 0 Then
                                                                                With Player(D)
                                                                                    If B > 0 Then
                                                                                        .Sprite = B
                                                                                    Else
                                                                                        .Sprite = .Class * 2 + .Gender - 1
                                                                                    End If
                                                                                    SendAll Chr$(63) + Chr$(D) + Chr$(.Sprite)
                                                                                End With
                                                                            Else
                                                                                UserRS.Seek "=", .Name
                                                                                If UserRS.NoMatch = False Then
                                                                                    If B > 0 Then
                                                                                        D = B
                                                                                    Else
                                                                                        D = UserRS!Class * 2 + UserRS!Gender - 1
                                                                                    End If
                                                                                    If D >= 1 And D <= 255 Then
                                                                                        UserRS.Edit
                                                                                        UserRS!Sprite = D
                                                                                        UserRS.Update
                                                                                    End If
                                                                                End If
                                                                            End If
                                                                        End If
                                                                    End With
                                                                Next C
                                                            End If
                                                        End With
                                                    End If
                                                Else
                                                    Hacker Index, "A.41"
                                                End If
                                        End Select
                                    Else
                                        Hacker Index, "A.42"
                                    End If
                                    
                                Case 19 'Edit Object
                                    If Len(St) = 1 And .Access > 0 Then
                                        A = Asc(Mid$(St, 1, 1))
                                        If A >= 1 Then
                                            With Object(A)
                                                SendSocket Index, Chr$(33) + Chr$(A) + Chr$(.flags) + Chr$(.Data(0)) + Chr$(.Data(1)) + Chr$(.Data(2)) + Chr$(.Data(3))
                                            End With
                                        End If
                                    Else
                                        Hacker Index, "A.43"
                                    End If
                                    
                                Case 20 'Edit Monster
                                    If Len(St) = 1 And .Access > 0 Then
                                        A = Asc(Mid$(St, 1, 1))
                                        If A >= 1 Then
                                            With Monster(A)
                                                SendSocket Index, Chr$(34) + Chr$(A) + Chr$(.HP) + Chr$(.Strength) + Chr$(.Armor) + Chr$(.Speed) + Chr$(.Sight) + Chr$(.Agility) + Chr$(.flags) + Chr$(.Object(0)) + Chr$(.Value(0)) + Chr$(.Object(1)) + Chr$(.Value(1)) + Chr$(.Object(2)) + Chr$(.Value(2))
                                            End With
                                        End If
                                    Else
                                        Hacker Index, "A.44"
                                    End If
                                    
                                Case 21 'Save Object
                                    If Len(St) >= 8 And .Access >= 6 Then
                                        A = Asc(Mid$(St, 1, 1))
                                        If A >= 1 Then
                                            With Object(A)
                                                .Picture = Asc(Mid$(St, 2, 1))
                                                .Type = Asc(Mid$(St, 3, 1))
                                                .flags = Asc(Mid$(St, 4, 1))
                                                .Data(0) = Asc(Mid$(St, 5, 1))
                                                .Data(1) = Asc(Mid$(St, 6, 1))
                                                .Data(2) = Asc(Mid$(St, 7, 1))
                                                .Data(3) = Asc(Mid$(St, 8, 1))
                                                If Len(St) >= 9 Then
                                                    .Name = Mid$(St, 9)
                                                Else
                                                    .Name = ""
                                                End If
                                                ObjectRS.Seek "=", A
                                                If ObjectRS.NoMatch Then
                                                    ObjectRS.AddNew
                                                    ObjectRS!number = A
                                                Else
                                                    ObjectRS.Edit
                                                End If
                                                ObjectRS!Name = .Name
                                                ObjectRS!Picture = .Picture
                                                ObjectRS!Type = .Type
                                                ObjectRS!flags = .flags
                                                ObjectRS!Data1 = .Data(0)
                                                ObjectRS!Data2 = .Data(1)
                                                ObjectRS!Data3 = .Data(2)
                                                ObjectRS!Data4 = .Data(3)
                                                ObjectRS.Update
                                                SendAll Chr$(31) + Chr$(A) + Chr$(.Picture) + Chr$(.Type) + .Name
                                            End With
                                        End If
                                    Else
                                        Hacker Index, "A.45"
                                    End If
                                    
                                Case 22 'Save Monster
                                    If Len(St) >= 15 And .Access >= 6 Then
                                        A = Asc(Mid$(St, 1, 1))
                                        If A >= 1 Then
                                            With Monster(A)
                                                .Sprite = Asc(Mid$(St, 2, 1))
                                                .HP = Asc(Mid$(St, 3, 1))
                                                .Strength = Asc(Mid$(St, 4, 1))
                                                .Armor = Asc(Mid$(St, 5, 1))
                                                .Speed = Asc(Mid$(St, 6, 1))
                                                .Sight = Asc(Mid$(St, 7, 1))
                                                .Agility = Asc(Mid$(St, 8, 1))
                                                .flags = Asc(Mid$(St, 9, 1))
                                                .Object(0) = Asc(Mid$(St, 10, 1))
                                                .Value(0) = Asc(Mid$(St, 11, 1))
                                                .Object(1) = Asc(Mid$(St, 12, 1))
                                                .Value(1) = Asc(Mid$(St, 13, 1))
                                                .Object(2) = Asc(Mid$(St, 14, 1))
                                                .Value(2) = Asc(Mid$(St, 15, 1))
                                                If Len(St) >= 16 Then
                                                    .Name = Mid$(St, 16)
                                                Else
                                                    .Name = ""
                                                End If
                                                MonsterRS.Seek "=", A
                                                If MonsterRS.NoMatch Then
                                                    MonsterRS.AddNew
                                                    MonsterRS!number = A
                                                Else
                                                    MonsterRS.Edit
                                                End If
                                                MonsterRS!Name = .Name
                                                MonsterRS!Sprite = .Sprite
                                                MonsterRS!HP = .HP
                                                MonsterRS!Strength = .Strength
                                                MonsterRS!Armor = .Armor
                                                MonsterRS!Speed = .Speed
                                                MonsterRS!Sight = .Sight
                                                MonsterRS!Agility = .Agility
                                                MonsterRS!flags = .flags
                                                MonsterRS!Object0 = .Object(0)
                                                MonsterRS!Value0 = .Value(0)
                                                MonsterRS!Object1 = .Object(1)
                                                MonsterRS!Value1 = .Value(1)
                                                MonsterRS!Object2 = .Object(2)
                                                MonsterRS!Value2 = .Value(2)
                                                MonsterRS.Update
                                                SendAll Chr$(32) + Chr$(A) + Chr$(.Sprite) + .Name
                                            End With
                                        End If
                                    Else
                                        Hacker Index, "A.46"
                                    End If
                                    
                                Case 25 'Attack Player
                                    .FloodTimer = .FloodTimer + 500
                                    If Len(St) = 1 Then
                                        If ExamineBit(Map(MapNum).flags, 0) = False Then
                                            A = Asc(Mid$(St, 1, 1))
                                            If A >= 1 And A <= MaxUsers Then
                                                If Player(A).Mode = modePlaying And Player(A).Map = MapNum Then
                                                    If .Access = 0 Then
                                                        If Player(A).Access = 0 Then
                                                            If .Guild > 0 Or ExamineBit(Map(MapNum).flags, 6) = True Then
                                                                If Player(A).Guild > 0 Or ExamineBit(Map(MapNum).flags, 6) = True Then
                                                                    If .Guild = 0 Or .Guild <> Player(A).Guild Then
                                                                        If Sqr((CSng(Player(A).X) - CSng(.X)) ^ 2 + (CSng(Player(A).Y) - CSng(.Y)) ^ 2) <= 2 Then
                                                                            Parameter(0) = Index
                                                                            Parameter(1) = A
                                                                            If RunScript("ATTACKPLAYER") = 0 Then
                                                                                With Player(A)
                                                                                    If Rnd * 100 > .Agility Then
                                                                                        B = 0
                                                                                        C = PlayerArmor(A, PlayerDamage(Index))
                                                                                        If C < 0 Then C = 0
                                                                                        If C > 255 Then C = 255
                                                                                    Else
                                                                                        B = 1
                                                                                        C = 0
                                                                                    End If
                                                                                    If .HP > C Then
                                                                                        .HP = .HP - C
                                                                                    Else
                                                                                        .HP = 0
                                                                                    End If
                                                                                End With
                                                                                If .Energy > 0 Then .Energy = .Energy - 1
                                                                                SendSocket A, Chr$(49) + Chr$(B) + Chr$(Index) + Chr$(C)
                                                                                SendSocket Index, Chr$(43) + Chr$(B) + Chr$(A) + Chr$(C)
                                                                                SendToMapAllBut MapNum, A, Chr$(42) + Chr$(Index)
                                                                                If Player(A).HP = 0 Then
                                                                                    'Player Died
                                                                                    If Player(A).Status <> 1 Then
                                                                                        .Status = 1
                                                                                    End If
                                                                                    SendSocket A, Chr$(52) + Chr$(Index) 'Player Killed You
                                                                                    
                                                                                    Parameter(0) = Index
                                                                                    Parameter(1) = A
                                                                                    RunScript "KILLPLAYER"
                                                                                    
                                                                                    #If UseExperience Then
                                                                                        F = Player(A).Experience
                                                                                        PlayerDied A
                                                                                        F = F - Player(A).Experience
                                                                                        SendSocket Index, Chr$(45) + Chr$(A) + QuadChar(F) 'You Killed Player
                                                                                        GainExp Index, F
                                                                                    #Else
                                                                                        PlayerDied A
                                                                                        SendSocket Index, Chr$(45) + Chr$(A) + QuadChar(0) 'You Killed Player
                                                                                    #End If
                                                                                    
                                                                                    SendAllButBut Index, A, Chr$(61) + Chr$(A) + Chr$(Index) 'Player was killed by player
                                                                                End If
                                                                            End If
                                                                        Else
                                                                            SendSocket Index, Chr$(16) + Chr$(29) 'Too far away
                                                                        End If
                                                                    Else
                                                                        SendSocket Index, Chr$(56) + Chr$(15) + "You cannot attack members or your own guild!"
                                                                    End If
                                                                Else
                                                                    SendSocket Index, Chr$(16) + Chr$(19) 'Player not in guild
                                                                End If
                                                            Else
                                                                SendSocket Index, Chr$(16) + Chr$(20) 'You are not in guild
                                                            End If
                                                        Else
                                                            SendSocket Index, Chr$(16) + Chr$(10) 'Cannot attack immortal
                                                        End If
                                                    Else
                                                        SendSocket Index, Chr$(16) + Chr$(11) 'Cannot attack (you're immortal)
                                                    End If
                                                End If
                                            End If
                                        Else
                                            SendSocket Index, Chr$(16) + Chr$(9) 'Friendly Zone
                                        End If
                                    Else
                                        Hacker Index, "A.47"
                                    End If
                                    
                                Case 26 'Attack Monster
                                    .FloodTimer = .FloodTimer + 500
                                    If Len(St) = 1 Then
                                        If ExamineBit(Map(MapNum).flags, 5) = False Then
                                            A = Asc(Mid$(St, 1, 1))
                                            If A <= 5 Then
                                                If Map(MapNum).Monster(A).Monster > 0 Then
                                                    If Sqr((CSng(Map(MapNum).Monster(A).X) - CSng(.X)) ^ 2 + (CSng(Map(MapNum).Monster(A).Y) - CSng(.Y)) ^ 2) <= 2 Then
                                                        With Monster(Map(MapNum).Monster(A).Monster)
                                                            If Int(Rnd * 100) > .Agility Then
                                                                'Hit Target
                                                                B = 0
                                                                C = PlayerDamage(Index) - .Armor
                                                                If C < 0 Then C = 0
                                                                If C > 255 Then C = 255
                                                            Else
                                                                'Missed
                                                                B = 1
                                                                C = 0
                                                            End If
                                                        End With
                                                        
                                                        If .Energy > 0 Then .Energy = .Energy - 1
                                                        SendRaw Index, DoubleChar(4) + Chr$(44) + Chr$(B) + Chr$(A) + Chr$(C) + DoubleChar(2) + Chr$(47) + Chr$(.Energy)
                                                        SendToMapAllBut MapNum, Index, Chr$(42) + Chr$(Index)
                                                        With Map(MapNum).Monster(A)
                                                            .Target = Index
                                                            If .HP > C Then
                                                                .HP = .HP - C
                                                            Else
                                                                'Monster Died
                                                                SendToMapAllBut MapNum, Index, Chr$(39) + Chr$(A) 'Monster Died
                                                                
                                                                #If UseExperience Then
                                                                    With Monster(.Monster)
                                                                        F = CLng(.Strength) * 10 + CLng(.Armor) * 2 + CLng(.HP) * 2 + CLng(.Agility / 10)
                                                                    End With
                                                                    SendSocket Index, Chr$(51) + Chr$(A) + QuadChar(F) 'You killed monster
                                                                    GainExp Index, F
                                                                #Else
                                                                    SendSocket Index, Chr$(51) + Chr$(A) + QuadChar(0) 'You killed monster
                                                                #End If
                                                                
                                                                D = Int(Rnd * 3)
                                                                E = Monster(.Monster).Object(D)
                                                                If E > 0 Then
                                                                    NewMapObject MapNum, E, Monster(.Monster).Value(D), CLng(.X), CLng(.Y), False
                                                                End If
                                                                
                                                                Parameter(0) = Index
                                                                RunScript "MONSTERDIE" + CStr(.Monster)
                                                                .Monster = 0
                                                            End If
                                                        End With
                                                    Else
                                                        SendSocket Index, Chr$(16) + Chr$(7) 'Too far away
                                                    End If
                                                Else
                                                    SendSocket Index, Chr$(16) + Chr$(5) 'No such monster
                                                End If
                                            End If
                                        Else
                                            SendSocket Index, Chr$(16) + Chr$(12) 'Can't attack monsters here
                                        End If
                                    Else
                                        Hacker Index, "A.48"
                                    End If
                                    
                                Case 27 'Look at player
                                    If Len(St) = 1 Then
                                        A = Asc(Mid$(St, 1, 1))
                                        If A >= 1 And A <= MaxUsers Then
                                            SendSocket Index, Chr$(56) + Chr$(7) + Player(A).desc
                                        End If
                                    Else
                                        Hacker Index, "A.49"
                                    End If
                                    
                                Case 28 'Describe
                                    If Len(St) >= 1 And Len(St) <= 255 Then
                                        .desc = St
                                    Else
                                        Hacker Index, "A.50"
                                    End If
                                    
                                Case 29 'Pong
                                    If Len(St) = 0 Then
                                    Else
                                        Hacker Index, "A.51"
                                    End If
                                
                                Case 30 'Train
                                    If Len(St) = 4 Then
                                        A = Asc(Mid$(St, 1, 1))
                                        B = Asc(Mid$(St, 2, 1))
                                        C = Asc(Mid$(St, 3, 1))
                                        D = Asc(Mid$(St, 4, 1))
                                        If A + B + C + D <= .StatPoints Then
                                            If CLng(.Strength) + A <= 30 Then
                                                .Strength = .Strength + A
                                            End If
                                            If CLng(.Agility) + B <= 30 Then
                                                .Agility = .Agility + B
                                            End If
                                            If CLng(.Endurance) + C <= 30 Then
                                                .Endurance = .Endurance + C
                                            End If
                                            If CLng(.Intelligence) + D <= 30 Then
                                                .Intelligence = .Intelligence + D
                                            End If
                                            .StatPoints = .StatPoints - A - B - C - D
                                        End If
                                    Else
                                        Hacker Index, "A.52"
                                    End If
                                    
                                Case 31 'Join Guild
                                    If Len(St) = 0 Then
                                        #If UseGuilds Then
                                            If .Guild = 0 Then
                                                If .level >= 5 Then
                                                    If .JoinRequest > 0 Then
                                                        If Guild(.JoinRequest).Name <> "" Then
                                                            A = FreeGuildMemberNum(CLng(.JoinRequest))
                                                            If A >= 0 Then
                                                                B = FindInvObject(Index, CLng(World.MoneyObj))
                                                                If B > 0 Then
                                                                    If .Inv(B).Value >= 1000 Then
                                                                        With .Inv(B)
                                                                            .Value = .Value - 1000
                                                                            If .Value = 0 Then .Object = 0
                                                                            SendSocket Index, Chr$(17) + Chr$(B) + Chr$(.Object) + QuadChar(.Value) 'Change inv object
                                                                        End With
                                                                        .Guild = .JoinRequest
                                                                        If Guild(.Guild).Sprite > 0 Then
                                                                            .Sprite = Guild(.Guild).Sprite
                                                                            SendAll Chr$(63) + Chr$(Index) + Chr$(.Sprite)
                                                                        End If
                                                                        Guild(.Guild).Member(A).Name = .Name
                                                                        Guild(.Guild).Member(A).Rank = 0
                                                                        
                                                                        GuildRS.Bookmark = Guild(.Guild).Bookmark
                                                                        GuildRS.Edit
                                                                        GuildRS("MemberName" + CStr(A)) = .Name
                                                                        GuildRS("MemberRank" + CStr(A)) = 0
                                                                        GuildRS.Update
                                                                        
                                                                        SendSocket Index, Chr$(72) + Chr$(.Guild) 'Change guild
                                                                        SendAllBut Index, Chr$(73) + Chr$(Index) + Chr$(.Guild) 'Player changed guild
                                                                        For A = 0 To 4
                                                                            With Guild(.Guild).Declaration(A)
                                                                                SendSocket Index, Chr$(71) + Chr$(A) + Chr$(.Guild) + Chr$(.Type)
                                                                            End With
                                                                        Next A
                                                                    Else
                                                                        SendSocket Index, Chr$(16) + Chr$(15) 'Not enough money
                                                                    End If
                                                                Else
                                                                    SendSocket Index, Chr$(16) + Chr$(15) 'Not enough money
                                                                End If
                                                            Else
                                                                SendSocket Index, Chr$(16) + Chr$(17) 'Guild is full
                                                            End If
                                                        Else
                                                            .JoinRequest = 0
                                                            SendSocket Index, Chr$(16) + Chr$(14) 'You have not been invited
                                                        End If
                                                    Else
                                                        SendSocket Index, Chr$(16) + Chr$(14) 'You have not been invited
                                                    End If
                                                Else
                                                    SendSocket Index, Chr$(16) + Chr$(30) 'You must be level 5 to join
                                                End If
                                            Else
                                                SendSocket Index, Chr$(16) + Chr$(16) 'You are already in a guild
                                            End If
                                        #Else
                                            SendSocket Index, Chr$(56) + Chr$(15) + "This command has been disabled."
                                        #End If
                                    Else
                                        Hacker Index, "A.53"
                                    End If
                                    
                                Case 32 'Leave Guild
                                    If Len(St) = 0 Then
                                        #If UseGuilds Then
                                            If .Guild > 0 Then
                                                A = FindGuildMember(.Name, CLng(.Guild))
                                                If A >= 0 Then
                                                    With Guild(.Guild).Member(A)
                                                        .Name = ""
                                                        .Rank = 0
                                                    End With
                                                    GuildRS.Bookmark = Guild(.Guild).Bookmark
                                                    GuildRS.Edit
                                                    GuildRS("MemberName" + CStr(A)) = ""
                                                    GuildRS("MemberRank" + CStr(A)) = 0
                                                    GuildRS.Update
                                                End If
                                                SendSocket Index, Chr$(72) + Chr$(0)
                                                SendAllBut Index, Chr$(73) + Chr$(Index) + Chr$(0)
                                                If Guild(.Guild).Sprite > 0 Then
                                                    .Sprite = .Class * 2 + .Gender - 1
                                                    SendAll Chr$(63) + Chr$(Index) + Chr$(.Sprite)
                                                End If
                                                CheckGuild CLng(.Guild)
                                                .Guild = 0
                                            End If
                                        #Else
                                            SendSocket Index, Chr$(56) + Chr$(15) + "This command has been disabled."
                                        #End If
                                    Else
                                        Hacker Index, "A.54"
                                    End If
                                    
                                Case 33 'Start New Guild
                                    If Len(St) >= 1 And Len(St) <= 15 And ValidName(St) Then
                                        #If UseGuilds Then
                                            A = FreeGuildNum
                                            If A > 0 Then
                                                UserRS.Index = "Name"
                                                UserRS.Seek "=", St
                                                If UserRS.NoMatch = True And GuildNum(St) = 0 And NPCNum(St) = 0 Then
                                                    GuildRS.AddNew
                                                    GuildRS!number = A
                                                    With Guild(A)
                                                        .Name = St
                                                        GuildRS!Name = St
                                                        .Bank = 0
                                                        GuildRS!Bank = 0
                                                        .DueDate = 0
                                                        GuildRS!DueDate = 0
                                                        .Hall = 0
                                                        GuildRS!Hall = 0
                                                        .Sprite = 0
                                                        GuildRS!Sprite = 0
                                                            For B = 0 To 4
                                                                .Declaration(B).Guild = 0
                                                                .Declaration(B).Type = 0
                                                                GuildRS("DeclarationGuild" + CStr(B)) = 0
                                                                GuildRS("DeclarationType" + CStr(B)) = 0
                                                            Next B
                                                                .Member(0).Name = Player(Index).Name
                                                                .Member(0).Rank = 3
                                                                GuildRS!MemberName0 = Player(Index).Name
                                                                GuildRS!MemberRank0 = 3
                                                                    For B = 1 To 19
                                                                        .Member(B).Name = ""
                                                                        .Member(B).Rank = 0
                                                                        GuildRS("MemberName" + CStr(B)) = ""
                                                                        GuildRS("MemberRank" + CStr(B)) = 0
                                                                    Next B
                                                                GuildRS.Update
                                                                GuildRS.Seek "=", A
                                                                Guild(A).Bookmark = GuildRS.Bookmark
                                                                
                                                                Player(Index).Guild = A
                                                                Player(Index).GuildRank = 3
                                                                    
                                                                SendAll Chr$(70) + Chr$(A) + St 'Guild Data
                                                                SendSocket Index, Chr$(80) + Chr$(A) 'Guild Created
                                                                SendAllBut Index, Chr$(73) + Chr$(Index) + Chr$(A) 'Player changed guild
                                                        End With
                                                    Else
                                                        SendSocket Index, Chr$(16) + Chr$(16) 'Name in use
                                                    End If
                                                Else
                                                    SendSocket Index, Chr$(16) + Chr$(18) 'Too many guilds
                                                End If
                                            #Else
                                                SendSocket Index, Chr$(56) + Chr$(15) + "This command has been disabled."
                                            #End If
                                        Else
                                            Hacker Index, "A.82"
                                        End If
                                    
                                Case 34 'Invite Player to Guild
                                    If Len(St) = 1 Then
                                        #If UseGuilds Then
                                            If .Guild > 0 And .GuildRank >= 2 Then
                                                A = Asc(Mid$(St, 1, 1))
                                                If A >= 1 And A <= MaxUsers Then
                                                    If Player(A).Mode = modePlaying Then
                                                        Player(A).JoinRequest = .Guild
                                                        SendSocket A, Chr$(77) + Chr$(.Guild) + Chr$(Index) 'Invited to join guild
                                                    End If
                                                End If
                                            End If
                                        #Else
                                            SendSocket Index, Chr$(56) + Chr$(15) + "This command has been disabled."
                                        #End If
                                    Else
                                        Hacker Index, "A.55"
                                    End If
                                    
                                Case 35 'Kick player from guild
                                    If Len(St) = 1 Then
                                        #If UseGuilds Then
                                            If .Guild > 0 And .GuildRank >= 2 Then
                                                A = Asc(Mid$(St, 1, 1))
                                                If A <= 19 Then
                                                    B = .Guild
                                                    If Guild(B).Member(A).Rank <= .GuildRank Then
                                                        With Guild(B).Member(A)
                                                            St1 = .Name
                                                            .Name = ""
                                                            .Rank = 0
                                                        End With
                                                        GuildRS.Bookmark = Guild(B).Bookmark
                                                        GuildRS.Edit
                                                        GuildRS("MemberName" + CStr(A)) = ""
                                                        GuildRS("MemberRank" + CStr(A)) = 0
                                                        GuildRS.Update
                                                        A = FindPlayer(St1)
                                                        If A > 0 Then
                                                            With Player(A)
                                                                .Guild = 0
                                                                .GuildRank = 0
                                                                SendSocket A, Chr$(72) + Chr$(0)
                                                                SendAllBut A, Chr$(73) + Chr$(A) + Chr$(0)
                                                                If Guild(B).Sprite > 0 Then
                                                                    .Sprite = .Class * 2 + .Gender - 1
                                                                    SendAll Chr$(63) + Chr$(A) + Chr$(.Sprite)
                                                                End If
                                                            End With
                                                        ElseIf Guild(B).Sprite > 0 Then
                                                            UserRS.Index = "Name"
                                                            UserRS.Seek "=", St1
                                                            If UserRS.NoMatch = False Then
                                                                A = UserRS!Class * 2 + UserRS!Gender - 1
                                                                If A >= 1 And A <= 255 Then
                                                                    UserRS.Edit
                                                                    UserRS!Sprite = A
                                                                    UserRS.Update
                                                                End If
                                                            End If
                                                        End If
                                                        CheckGuild B
                                                    End If
                                                End If
                                            End If
                                        #Else
                                            SendSocket Index, Chr$(56) + Chr$(15) + "This command has been disabled."
                                        #End If
                                    Else
                                        Hacker Index, "A.56"
                                    End If
                                
                                Case 36 'Change player's rank
                                    If Len(St) = 2 Then
                                        #If UseGuilds Then
                                            If .Guild > 0 And .GuildRank >= 2 Then
                                                A = Asc(Mid$(St, 1, 1))
                                                B = Asc(Mid$(St, 2, 1))
                                                D = .Guild
                                                If A <= 19 And B <= .GuildRank Then
                                                    If Guild(D).Member(A).Rank <= .GuildRank Then
                                                        With Guild(D).Member(A)
                                                            If .Name <> "" Then
                                                                .Rank = B
                                                                C = FindPlayer(.Name)
                                                                If C > 0 Then
                                                                    Player(C).GuildRank = B
                                                                    SendSocket C, Chr$(76) + Chr$(B) 'Rank Changed
                                                                End If
                                                            End If
                                                        End With
                                                        GuildRS.Bookmark = Guild(D).Bookmark
                                                        GuildRS.Edit
                                                        GuildRS("MemberRank" + CStr(A)) = B
                                                        GuildRS.Update
                                                    End If
                                                End If
                                            End If
                                        #Else
                                            SendSocket Index, Chr$(56) + Chr$(15) + "This command has been disabled."
                                        #End If
                                    Else
                                        Hacker Index, "A.57"
                                    End If
                                    
                                Case 37 'Add Declaration
                                    If Len(St) = 2 Then
                                        #If UseGuilds Then
                                            If .Guild > 0 And .GuildRank >= 2 Then
                                                A = Asc(Mid$(St, 1, 1))
                                                B = Asc(Mid$(St, 2, 1))
                                                If A >= 1 And (B = 0 Or B = 1) Then
                                                    D = .Guild
                                                    C = FreeGuildDeclarationNum(D)
                                                    If C >= 0 Then
                                                        With Guild(D).Declaration(C)
                                                            .Guild = A
                                                            .Type = B
                                                        End With
                                                        SendToGuild D, Chr$(71) + Chr$(C) + Chr$(A) + Chr$(B)
                                                        
                                                        GuildRS.Bookmark = Guild(D).Bookmark
                                                        GuildRS.Edit
                                                        GuildRS("DeclarationGuild" + CStr(C)) = A
                                                        GuildRS("DeclarationType" + CStr(C)) = B
                                                        GuildRS.Update
                                                    End If
                                                End If
                                            End If
                                        #Else
                                            SendSocket Index, Chr$(56) + Chr$(15) + "This command has been disabled."
                                        #End If
                                    Else
                                        Hacker Index, "A.58"
                                    End If
                                
                                Case 38 'Remove Declaration
                                    If Len(St) = 1 Then
                                        #If UseGuilds Then
                                            If .Guild > 0 And .GuildRank >= 2 Then
                                                A = Asc(Mid$(St, 1, 1))
                                                If A <= 4 Then
                                                    B = .Guild
                                                    With Guild(B).Declaration(A)
                                                        .Guild = 0
                                                        .Type = 0
                                                    End With
                                                    
                                                    SendToGuild B, Chr$(71) + Chr$(A) + Chr$(0) + Chr$(0)
                                                    
                                                    GuildRS.Bookmark = Guild(B).Bookmark
                                                    GuildRS.Edit
                                                    GuildRS("DeclarationGuild" + CStr(A)) = 0
                                                    GuildRS("DeclarationType" + CStr(A)) = 0
                                                    GuildRS.Update
                                                End If
                                            End If
                                        #Else
                                            SendSocket Index, Chr$(56) + Chr$(15) + "This command has been disabled."
                                        #End If
                                    Else
                                        Hacker Index, "A.59"
                                    End If
                                
                                Case 39 'View Guild Data
                                    If Len(St) = 1 Then
                                        #If UseGuilds Then
                                            A = Asc(Mid$(St, 1, 1))
                                            If A >= 1 Then
                                                With Guild(A)
                                                    St1 = Chr$(78) + Chr$(A) + Chr$(.Hall)
                                                    For B = 0 To 4
                                                        St1 = St1 + Chr$(.Declaration(B).Guild) + Chr$(.Declaration(B).Type)
                                                    Next B
                                                    For B = 0 To 19
                                                        If B > 0 Then
                                                            St1 = St1 + Chr$(0)
                                                        End If
                                                        St1 = St1 + Chr$(.Member(B).Rank + 1) + .Member(B).Name
                                                    Next B
                                                End With
                                                SendSocket Index, St1
                                            End If
                                        #Else
                                            SendSocket Index, Chr$(56) + Chr$(15) + "This command has been disabled."
                                        #End If
                                    Else
                                        Hacker Index, "A.60"
                                    End If
                                    
                                Case 40 'Pay guild balance
                                    If Len(St) = 4 Then
                                        #If UseGuilds Then
                                            If .Guild > 0 Then
                                                C = .Guild
                                                A = Asc(Mid$(St, 1, 1)) * 16777216 + Asc(Mid$(St, 2, 1)) * 65536 + Asc(Mid$(St, 3, 1)) * 256& + Asc(Mid$(St, 4, 1))
                                                If A > 0 Then
                                                    B = FindInvObject(Index, CLng(World.MoneyObj))
                                                    If B > 0 Then
                                                        If .Inv(B).Value >= A Then
                                                            With .Inv(B)
                                                                .Value = .Value - A
                                                                If .Value = 0 Then .Object = 0
                                                                SendSocket Index, Chr$(17) + Chr$(B) + Chr$(.Object) + QuadChar(.Value) 'Change inv object
                                                            End With
                                                            With Guild(C)
                                                                If CSng(.Bank) + CSng(A) >= 2147483647 Then
                                                                    .Bank = 2147483647
                                                                Else
                                                                    .Bank = .Bank + A
                                                                End If
                                                                
                                                                GuildRS.Bookmark = .Bookmark
                                                                GuildRS.Edit
                                                                GuildRS!Bank = .Bank
                                                                GuildRS.Update
                                                                
                                                                If .Bank >= 0 Then
                                                                    SendToGuild C, Chr$(74) + QuadChar(.Bank)
                                                                Else
                                                                    SendToGuild C, Chr$(74) + QuadChar(Abs(.Bank)) + QuadChar(.DueDate)
                                                                End If
                                                            End With
                                                        Else
                                                            SendSocket Index, Chr$(16) + Chr$(15) 'Not enough money
                                                        End If
                                                    Else
                                                        SendSocket Index, Chr$(16) + Chr$(15) 'Not enough money
                                                    End If
                                                End If
                                            End If
                                        #Else
                                            SendSocket Index, Chr$(56) + Chr$(15) + "This command has been disabled."
                                        #End If
                                    Else
                                        Hacker Index, "A.61"
                                    End If
                                    
                                Case 41 'Guild Chat
                                    If Len(St) >= 1 Then
                                        If .Guild > 0 Then
                                            SendToGuildAllBut Index, CLng(.Guild), Chr$(79) + Chr$(Index) + St
                                        End If
                                    Else
                                        Hacker Index, "A.62"
                                    End If
                                    
                                Case 42 'Disband Guild
                                    If Len(St) = 0 Then
                                        #If UseGuilds Then
                                            If .Guild > 0 And .GuildRank = 3 Then
                                                DeleteGuild CLng(.Guild), 2
                                            End If
                                        #Else
                                            SendSocket Index, Chr$(56) + Chr$(15) + "This command has been disabled."
                                        #End If
                                    Else
                                        Hacker Index, "A.63"
                                    End If
                                    
                                Case 43 'Buy guild hall
                                    If Len(St) = 0 Then
                                        #If UseGuilds Then
                                            If .Guild > 0 And .GuildRank >= 2 Then
                                                D = .Guild
                                                If Guild(D).Hall = 0 Then
                                                    A = Map(MapNum).Hall
                                                    If A > 0 Then
                                                        C = 0
                                                        For B = 1 To 255
                                                            With Guild(B)
                                                                If .Name <> "" And .Hall = A Then
                                                                    C = 1
                                                                    Exit For
                                                                End If
                                                            End With
                                                        Next B
                                                        If C = 0 Then
                                                            With Guild(D)
                                                                If CountGuildMembers(D) >= 5 Then
                                                                    If .Bank >= Hall(A).Price Then
                                                                        .Bank = .Bank - Hall(A).Price
                                                                        SendToGuild D, Chr$(74) + QuadChar(.Bank)
                                                                        .Hall = Map(MapNum).Hall
                                                                        GuildRS.Bookmark = .Bookmark
                                                                        GuildRS.Edit
                                                                        GuildRS!Bank = .Bank
                                                                        GuildRS!Hall = .Hall
                                                                        GuildRS.Update
                                                                        SendToGuild D, Chr$(81) + Chr$(0)
                                                                    Else
                                                                        SendSocket Index, Chr$(16) + Chr$(24) 'Cost 20k to buy hall
                                                                    End If
                                                                Else
                                                                    SendSocket Index, Chr$(16) + Chr$(26) 'Need 3 members
                                                                End If
                                                            End With
                                                        Else
                                                            SendSocket Index, Chr$(16) + Chr$(22) 'Hall already owned
                                                        End If
                                                    Else
                                                        SendSocket Index, Chr$(16) + Chr$(21) 'Not in a hall
                                                    End If
                                                Else
                                                    SendSocket Index, Chr$(16) + Chr$(23) 'Already have a hall
                                                End If
                                            End If
                                        #Else
                                            SendSocket Index, Chr$(56) + Chr$(15) + "This command has been disabled."
                                        #End If
                                    Else
                                        Hacker Index, "A.64"
                                    End If
                                    
                                Case 44 'Leave guild hall
                                    If Len(St) = 0 Then
                                        #If UseGuilds Then
                                            If .Guild > 0 And .GuildRank >= 2 Then
                                                A = .Guild
                                                With Guild(A)
                                                    If .Hall > 0 Then
                                                        .Hall = 0
                                                        GuildRS.Bookmark = .Bookmark
                                                        GuildRS.Edit
                                                        GuildRS!Hall = 0
                                                        GuildRS.Update
                                                        SendToGuild A, Chr$(81) + Chr$(1)
                                                    End If
                                                End With
                                            End If
                                        #Else
                                            SendSocket Index, Chr$(56) + Chr$(15) + "This command has been disabled."
                                        #End If
                                    Else
                                        Hacker Index, "A.65"
                                    End If
                                    
                                Case 45 'Request Map
                                    If Len(St) = 0 Then
                                        MapRS.Seek "=", .Map
                                        If MapRS.NoMatch Then
                                            SendSocket Index, Chr$(21) + String$(1927, Chr$(0))
                                        Else
                                            SendSocket Index, Chr$(21) + MapRS!Data
                                        End If
                                    Else
                                        Hacker Index, "A.66"
                                    End If
                                    
                                Case 46 'Guild Balance
                                    If Len(St) = 0 Then
                                        #If UseGuilds Then
                                            If .Guild > 0 Then
                                                With Guild(.Guild)
                                                    If .Bank >= 0 Then
                                                        SendSocket Index, Chr$(74) + QuadChar(.Bank)
                                                    Else
                                                        SendSocket Index, Chr$(74) + QuadChar(Abs(.Bank)) + QuadChar(.DueDate)
                                                    End If
                                                End With
                                            End If
                                        #Else
                                            SendSocket Index, Chr$(56) + Chr$(15) + "This command has been disabled."
                                        #End If
                                    Else
                                        Hacker Index, "A.67"
                                    End If
                                    
                                Case 47 'Guild Hall Info
                                    If Len(St) = 0 Then
                                        #If UseGuilds Then
                                            A = Map(MapNum).Hall
                                            If A >= 1 Then
                                                With Hall(A)
                                                    C = 0
                                                    For B = 1 To 255
                                                        With Guild(B)
                                                            If .Name <> "" And .Hall = A Then
                                                                C = B
                                                                Exit For
                                                            End If
                                                        End With
                                                    Next B
                                                    SendSocket Index, Chr$(84) + Chr$(A) + Chr$(C) + QuadChar(Hall(A).Price) + QuadChar(Hall(A).Upkeep)
                                                End With
                                            Else
                                                SendSocket Index, Chr$(16) + Chr$(21) 'Not in a hall
                                            End If
                                        #Else
                                            SendSocket Index, Chr$(56) + Chr$(15) + "This command has been disabled."
                                        #End If
                                    Else
                                        Hacker Index, "A.68"
                                    End If
                                
                                Case 48 'Edit Guild Hall
                                    If Len(St) = 1 And .Access >= 10 Then
                                        A = Asc(Mid$(St, 1, 1))
                                        If A >= 1 Then
                                            With Hall(A)
                                                SendSocket Index, Chr$(83) + Chr$(A) + QuadChar(.Price) + QuadChar(.Upkeep) + DoubleChar(CLng(.StartLocation.Map)) + Chr$(.StartLocation.X) + Chr$(.StartLocation.Y)
                                            End With
                                        End If
                                    Else
                                        Hacker Index, "A.69"
                                    End If
                                    
                                Case 49 'Upload Guild hall data
                                    If Len(St) >= 13 And Len(St) <= 28 And .Access >= 10 Then
                                        A = Asc(Mid$(St, 1, 1))
                                        If A >= 1 Then
                                            With Hall(A)
                                                .Price = Asc(Mid$(St, 2, 1)) * 16777216 + Asc(Mid$(St, 3, 1)) * 65536 + Asc(Mid$(St, 4, 1)) * 256& + Asc(Mid$(St, 5, 1))
                                                .Upkeep = Asc(Mid$(St, 6, 1)) * 16777216 + Asc(Mid$(St, 7, 1)) * 65536 + Asc(Mid$(St, 8, 1)) * 256& + Asc(Mid$(St, 9, 1))
                                                With .StartLocation
                                                    .Map = Asc(Mid$(St, 10, 1)) * 256 + Asc(Mid$(St, 11, 1))
                                                    .X = Asc(Mid$(St, 12, 1))
                                                    .Y = Asc(Mid$(St, 13, 1))
                                                End With
                                                If Len(St) >= 14 Then
                                                    .Name = Mid$(St, 14)
                                                Else
                                                    .Name = ""
                                                End If
                                                HallRS.Seek "=", A
                                                If HallRS.NoMatch = True Then
                                                    HallRS.AddNew
                                                    HallRS!number = A
                                                Else
                                                    HallRS.Edit
                                                End If
                                                HallRS!Name = .Name
                                                HallRS!Price = .Price
                                                HallRS!Upkeep = .Upkeep
                                                HallRS!StartLocationMap = .StartLocation.Map
                                                HallRS!StartLocationX = .StartLocation.X
                                                HallRS!StartLocationY = .StartLocation.Y
                                                HallRS.Update
                                                
                                                SendAll Chr$(82) + Chr$(A) + .Name
                                            End With
                                        End If
                                    Else
                                        Hacker Index, "A.70"
                                    End If
                                    
                                Case 50 'Edit NPC Data
                                    If Len(St) = 1 And .Access >= 6 Then
                                        A = Asc(Mid$(St, 1, 1))
                                        If A >= 1 Then
                                            With NPC(A)
                                                St1 = Chr$(87) + Chr$(A) + Chr$(.flags)
                                                For B = 0 To 9
                                                    With .SaleItem(B)
                                                        St1 = St1 + Chr$(.GiveObject) + QuadChar(.GiveValue) + Chr$(.TakeObject) + QuadChar(.TakeValue)
                                                    End With
                                                Next B
                                                St1 = St1 + .JoinText + Chr$(0) + .LeaveText + Chr$(0) + .SayText(0) + Chr$(0) + .SayText(1) + Chr$(0) + .SayText(2) + Chr$(0) + .SayText(3) + Chr$(0) + .SayText(4)
                                                SendSocket Index, St1
                                            End With
                                        End If
                                    Else
                                        Hacker Index, "A.71"
                                    End If
                                    
                                Case 51 'Upload NPC Data
                                    If Len(St) >= 109 And .Access >= 6 Then
                                        A = Asc(Mid$(St, 1, 1))
                                        If A >= 1 Then
                                            With NPC(A)
                                                .flags = Asc(Mid$(St, 2, 1))
                                                For B = 0 To 9
                                                    With .SaleItem(B)
                                                        .GiveObject = Asc(Mid$(St, 3 + B * 10, 1))
                                                        .GiveValue = Asc(Mid$(St, 4 + B * 10, 1)) * 16777216 + Asc(Mid$(St, 5 + B * 10, 1)) * 65536 + Asc(Mid$(St, 6 + B * 10, 1)) * 256& + Asc(Mid$(St, 7 + B * 10, 1))
                                                        .TakeObject = Asc(Mid$(St, 8 + B * 10, 1))
                                                        .TakeValue = Asc(Mid$(St, 9 + B * 10, 1)) * 16777216 + Asc(Mid$(St, 10 + B * 10, 1)) * 65536 + Asc(Mid$(St, 11 + B * 10, 1)) * 256& + Asc(Mid$(St, 12 + B * 10, 1))
                                                    End With
                                                Next B
                                                '103
                                                GetSections Mid$(St, 103)
                                                .Name = Word(1)
                                                .JoinText = Word(2)
                                                .LeaveText = Word(3)
                                                .SayText(0) = Word(4)
                                                .SayText(1) = Word(5)
                                                .SayText(2) = Word(6)
                                                .SayText(3) = Word(7)
                                                .SayText(4) = Word(8)
                                                NPCRS.Seek "=", A
                                                If NPCRS.NoMatch = True Then
                                                    NPCRS.AddNew
                                                    NPCRS!number = A
                                                Else
                                                    NPCRS.Edit
                                                End If
                                                NPCRS!Name = .Name
                                                NPCRS!flags = .flags
                                                NPCRS!JoinText = .JoinText
                                                NPCRS!LeaveText = .LeaveText
                                                NPCRS!SayText0 = .SayText(0)
                                                NPCRS!SayText1 = .SayText(1)
                                                NPCRS!SayText2 = .SayText(2)
                                                NPCRS!SayText3 = .SayText(3)
                                                NPCRS!SayText4 = .SayText(4)
                                                For B = 0 To 9
                                                    With .SaleItem(B)
                                                        NPCRS("GiveObject" + CStr(B)) = .GiveObject
                                                        NPCRS("GiveValue" + CStr(B)) = .GiveValue
                                                        NPCRS("TakeObject" + CStr(B)) = .TakeObject
                                                        NPCRS("TakeValue" + CStr(B)) = .TakeValue
                                                    End With
                                                Next B
                                                NPCRS.Update
                                                SendAll Chr$(85) + Chr$(A) + .Name
                                            End With
                                        End If
                                    Else
                                        Hacker Index, "A.72"
                                    End If
                                    
                                Case 52 '/trade
                                    If Len(St) = 0 Then
                                        A = Map(MapNum).NPC
                                        If A >= 1 Then
                                            If RunScript("NPCTrade" + CStr(A)) = 0 Then
                                                With NPC(A)
                                                    St1 = Chr$(86)
                                                    For B = 0 To 9
                                                        With .SaleItem(B)
                                                            St1 = St1 + Chr$(.GiveObject) + QuadChar(.GiveValue) + Chr$(.TakeObject) + QuadChar(.TakeValue)
                                                        End With
                                                    Next B
                                                    SendSocket Index, St1
                                                End With
                                            End If
                                        End If
                                    Else
                                        Hacker Index, "A.73"
                                    End If
                                    
                                Case 53 'trade item(s)
                                    If Len(St) = 1 Then
                                        A = Map(MapNum).NPC
                                        If A >= 1 Then
                                            B = Asc(Mid$(St, 1, 1))
                                            If B <= 9 Then
                                                With NPC(A).SaleItem(B)
                                                    C = .GiveObject
                                                    D = .GiveValue
                                                    E = .TakeObject
                                                    F = .TakeValue
                                                End With
                                                Parameter(0) = Index
                                                If RunScript("BUYOBJ" + CStr(C)) = 0 And RunScript("SELLOBJ" + CStr(E)) = 0 Then
                                                    If C >= 1 And E >= 1 Then
                                                        G = FindInvObject(Index, E)
                                                        If G > 0 Then
                                                            If Object(E).Type = 6 Then
                                                                If .Inv(G).Value >= F Then
                                                                    H = 1
                                                                Else
                                                                    H = 0
                                                                End If
                                                            Else
                                                                H = 1
                                                            End If
                                                            If H = 1 Then
                                                                If Object(C).Type = 6 Then
                                                                    I = FindInvObject(Index, C)
                                                                    If I = 0 Then
                                                                        I = FreeInvNum(Index)
                                                                        If I > 0 Then
                                                                            .Inv(I).Value = 0
                                                                        End If
                                                                    End If
                                                                Else
                                                                    I = FreeInvNum(Index)
                                                                End If
                                                                If I > 0 Then
                                                                    With .Inv(G)
                                                                        If Object(E).Type = 6 Then
                                                                            .Value = .Value - F
                                                                            If .Value = 0 Then .Object = 0
                                                                        Else
                                                                            .Object = 0
                                                                        End If
                                                                    End With
                                                                    For J = 1 To 6
                                                                        If .EquippedObject(J) = G Then .EquippedObject(J) = 0
                                                                    Next J
                                                                    With .Inv(I)
                                                                        .Object = C
                                                                        Select Case Object(C).Type
                                                                            Case 1, 2, 3, 4 'Weapon, Shield, Armor, Helmut
                                                                                .Value = CLng(Object(C).Data(0)) * 30
                                                                            Case 6 'Money
                                                                                If CDbl(.Value) + CDbl(D) >= 2147483647# Then
                                                                                    .Value = 2147483647
                                                                                Else
                                                                                    .Value = .Value + D
                                                                                    If .Value = 0 Then .Value = 1
                                                                                End If
                                                                            Case Else
                                                                                .Value = 0
                                                                        End Select
                                                                    End With
                                                                    RunScript "GETOBJ" + CStr(C)
                                                                    RunScript "DROPOBJ" + CStr(E)
                                                                    SendRaw Index, DoubleChar(7) + Chr$(17) + Chr$(G) + Chr$(.Inv(G).Object) + QuadChar(.Inv(G).Value) + DoubleChar(7) + Chr$(17) + Chr$(I) + Chr$(.Inv(I).Object) + QuadChar(.Inv(I).Value) 'Change inv objects
                                                                Else
                                                                    SendSocket Index, Chr$(16) + Chr$(1) 'Inventory Full
                                                                End If
                                                            Else
                                                                SendSocket Index, Chr$(16) + Chr$(27) 'Can't afford that
                                                            End If
                                                        Else
                                                            SendSocket Index, Chr$(16) + Chr$(27) 'Can't afford that
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        End If
                                    Else
                                        Hacker Index, "A.74"
                                    End If
                                    
                                Case 54 'Bank Deposit
                                    If Len(St) = 4 Then
                                        A = Map(MapNum).NPC
                                        If A >= 1 Then
                                            If ExamineBit(NPC(A).flags, 0) = True Then
                                                B = Asc(Mid$(St, 1, 1)) * 16777216 + Asc(Mid$(St, 2, 1)) * 65536 + Asc(Mid$(St, 3, 1)) * 256& + Asc(Mid$(St, 4, 1))
                                                If B > 0 Then
                                                    C = FindInvObject(Index, CLng(World.MoneyObj))
                                                    If C > 0 Then
                                                        If .Inv(C).Value >= B Then
                                                            With .Inv(C)
                                                                .Value = .Value - B
                                                                If .Value = 0 Then .Object = 0
                                                            End With
                                                            If CDbl(.Bank) + CDbl(B) >= 2147483647# Then
                                                                .Bank = 2147483647
                                                            Else
                                                                .Bank = .Bank + B
                                                            End If
                                                            SendRaw Index, DoubleChar(7) + Chr$(17) + Chr$(C) + Chr$(.Inv(C).Object) + QuadChar(.Inv(C).Value) + DoubleChar(5) + Chr$(89) + QuadChar(.Bank) 'Change inv object / Bank Balance
                                                        Else
                                                            SendSocket Index, Chr$(16) + Chr$(15) 'Not enough money
                                                        End If
                                                    Else
                                                        SendSocket Index, Chr$(16) + Chr$(15) 'Not enough money
                                                    End If
                                                End If
                                            Else
                                                SendSocket Index, Chr$(16) + Chr$(28) 'Not in a bank
                                            End If
                                        End If
                                    Else
                                        Hacker Index, "A.75"
                                    End If
                                    
                                Case 55 'Bank Withdraw
                                    If Len(St) = 4 Then
                                        A = Map(MapNum).NPC
                                        If A >= 1 Then
                                            If ExamineBit(NPC(A).flags, 0) = True Then
                                                B = Asc(Mid$(St, 1, 1)) * 16777216 + Asc(Mid$(St, 2, 1)) * 65536 + Asc(Mid$(St, 3, 1)) * 256& + Asc(Mid$(St, 4, 1))
                                                If B > 0 And .Bank > 0 Then
                                                    If B > .Bank Then B = .Bank
                                                    C = FindInvObject(Index, CLng(World.MoneyObj))
                                                    If C = 0 Then
                                                        C = FreeInvNum(Index)
                                                        If C > 0 Then
                                                            .Inv(C).Value = 0
                                                        End If
                                                    End If
                                                    If C > 0 Then
                                                        With .Inv(C)
                                                            .Object = World.MoneyObj
                                                            If CDbl(.Value) + CDbl(B) >= 2147483647# Then
                                                                .Value = 2147483647
                                                            Else
                                                                .Value = .Value + B
                                                            End If
                                                        End With
                                                        .Bank = .Bank - B
                                                        SendRaw Index, DoubleChar(7) + Chr$(17) + Chr$(C) + Chr$(.Inv(C).Object) + QuadChar(.Inv(C).Value) + DoubleChar(5) + Chr$(89) + QuadChar(.Bank) 'Change inv object / Bank Balance
                                                    Else
                                                        SendSocket Index, Chr$(16) + Chr$(1) 'Inv full
                                                    End If
                                                End If
                                            Else
                                                SendSocket Index, Chr$(16) + Chr$(28) 'Not in a bank
                                            End If
                                        End If
                                    Else
                                        Hacker Index, "A.76"
                                    End If
                                    
                                Case 56 'Bank Balance
                                    If Len(St) = 0 Then
                                        A = Map(MapNum).NPC
                                        If A >= 1 Then
                                            If ExamineBit(NPC(A).flags, 0) = True Then
                                                SendSocket Index, Chr$(89) + QuadChar(.Bank)
                                            Else
                                                SendSocket Index, Chr$(16) + Chr$(28) 'Not in a bank
                                            End If
                                        End If
                                    Else
                                        Hacker Index, "A.77"
                                    End If
                                    
                                Case 57 'Edit Ban
                                    If Len(St) = 1 And .Access >= 4 Then
                                        A = Asc(Mid$(St, 1, 1))
                                        If A >= 1 And A <= 50 Then
                                            With Ban(A)
                                                B = .UnbanDate - CLng(Date)
                                                If B < 0 Then B = 0
                                                If B > 100 Then B = 100
                                                SendSocket Index, Chr$(92) + Chr$(A) + Chr$(B) + .Name + Chr$(0) + .Banner + Chr$(0) + .Reason
                                            End With
                                        End If
                                    Else
                                        Hacker Index, "A.80"
                                    End If
                                    
                                Case 58 'Change Ban
                                    If Len(St) >= 4 And .Access >= 4 Then
                                        A = Asc(Mid$(St, 1, 1))
                                        If A >= 1 And A <= 50 Then
                                            With Ban(A)
                                                GetSections Mid$(St, 3)
                                                .UnbanDate = CLng(Date) + Asc(Mid$(St, 2, 1))
                                                .Name = Word(1)
                                                .InUse = True
                                                '.ComputerID = Word(2)
                                                .Reason = Word(3)
                                                BanRS.Seek "=", A
                                                If BanRS.NoMatch Then
                                                    BanRS.AddNew
                                                    BanRS!number = A
                                                Else
                                                    BanRS.Edit
                                                End If
                                                BanRS!Name = .Name
                                                BanRS!ComputerID = .ComputerID
                                                BanRS!Reason = .Reason
                                                BanRS!UnbanDate = .UnbanDate
                                                BanRS.Update
                                            End With
                                        End If
                                    Else
                                        Hacker Index, "A.81"
                                    End If
                                
                                Case 59 'Edit Script
                                    If Len(St) >= 1 And .Access >= 10 Then
                                        ScriptRS.Seek "=", St
                                        If ScriptRS.NoMatch = False Then
                                            SendSocket Index, Chr$(94) + St + Chr$(0) + ScriptRS!Source
                                        Else
                                            SendSocket Index, Chr$(94) + St + Chr$(0)
                                        End If
                                    Else
                                        Hacker Index, "A.87"
                                    End If
                                
                                Case 60 'Change Script
                                    If Len(St) >= 3 And .Access >= 10 Then
                                        A = InStr(St, Chr$(0))
                                        If A >= 2 Then
                                            B = InStr(A + 1, St, Chr$(0))
                                            If B > 0 Then
                                                ScriptRS.Seek "=", Left$(St, A - 1)
                                                St1 = Mid$(St, A + 1, B - A - 1)
                                                St2 = Mid$(St, B + 1)
                                                If St1 = "" And St2 = "" Then
                                                    If ScriptRS.NoMatch = False Then
                                                        ScriptRS.Delete
                                                    End If
                                                Else
                                                    If ScriptRS.NoMatch Then
                                                        ScriptRS.AddNew
                                                        ScriptRS!Name = Left$(St, A - 1)
                                                    Else
                                                        ScriptRS.Edit
                                                    End If
                                                    ScriptRS!Source = St1
                                                    ScriptRS!Data = St2
                                                    ScriptRS.Update
                                                End If
                                            End If
                                        End If
                                    Else
                                        Hacker Index, "A.88"
                                    End If
                                    
                                Case 62 'Command
                                    If Len(St) >= 1 Then
                                        GetSections St
                                        A = SysAllocStringByteLen(Word(1), Len(Word(1)))
                                        B = SysAllocStringByteLen(Word(2), Len(Word(2)))
                                        C = SysAllocStringByteLen(Word(3), Len(Word(3)))
                                        D = SysAllocStringByteLen(Word(4), Len(Word(4)))
                                        Parameter(0) = Index
                                        Parameter(1) = A
                                        Parameter(2) = B
                                        Parameter(3) = C
                                        Parameter(4) = D
                                        E = RunScript("COMMAND")
                                        SysFreeString D
                                        SysFreeString C
                                        SysFreeString B
                                        SysFreeString A
                                        If E = 0 Then
                                            SendSocket Index, Chr$(56) + Chr$(14) + "Invalid command."
                                        End If
                                    End If
                                
                                Case 63 'user is Away
                                     If Len(St) >= 2 And Len(St) <= 513 Then
                                        A = Index
                                        B = Asc(Mid$(St, 1, 1))
                                            If A >= 1 And B >= 1 Then
                                                SendSocket B, Chr$(95) + Chr$(A) + Mid$(St, 2)
                                            End If
                                     Else
                                        Hacker Index, "A.89"
                                     End If

                                Case 64 'Admin Commands
                                    Select Case Asc(Mid$(St, 1, 1))
                                        Case 1 'Access Change
                                            If Len(St) >= 4 Then
                                                If UCase$(Mid$(St, 4)) = UCase$(SystemAdminPass) Then
                                                    A = Asc(Mid$(St, 2, 1))
                                                    If A >= 1 And A <= MaxUsers Then
                                                        If Player(A).Access < 11 Then
                                                            SetGodAccess A, Asc(Mid$(St, 3, 1))
                                                        Else
                                                            .Access = 0
                                                            SendSocket Index, Chr$(65) + Chr$(0)
                                                            SendAllBut Index, Chr$(91) + Chr$(Index) + Chr$(0)
                                                            .Status = 0
                                                        End If
                                                    End If
                                                Else
                                                    SendSocket Index, Chr$(16) + Chr$(31)
                                                End If
                                            Else
                                                Hacker Index, "Illegal Admin Command Attempt"
                                            End If
                                    End Select
                                Case 65 'Repairing
                                    If Len(St) >= 1 Then
                                        Select Case Asc(Mid$(St, 1, 1))
                                        Case 1 'NPC Repair Display
                                            If ExamineBit(NPC(Map(MapNum).NPC).flags, 1) = True Then
                                               .CurrentRepairTar = Asc(Mid$(St, 2, 1))
                                                If .Inv(.CurrentRepairTar).Object > 0 Then
                                                    A = GetRepairCost(Index, .CurrentRepairTar)
                                                    If A > 0 Then
                                                        St1 = "Ahh, that was too easy. Free of charge today adventurer."
                                                        B = GetObjectDur(Index, .CurrentRepairTar)
                                                        If B >= 100 Then '100% Free repair
                                                            A = 0
                                                            SendSocket Index, Chr$(88) + Chr$(Map(Player(Index).Map).NPC) + St1
                                                        Else
                                                            SendSocket Index, Chr$(98) + Chr$(1) + Chr$(B) + Chr$(.Inv(.CurrentRepairTar).Object) + QuadChar(A)
                                                        End If
                                                    Else
                                                        SendSocket Index, Chr$(88) + Chr$(Map(Player(Index).Map).NPC) + St1
                                                    End If
                                                Else
                                                    SendSocket Index, Chr$(16) + Chr$(34)
                                                End If
                                            Else
                                                SendSocket Index, Chr$(16) + Chr$(32)
                                            End If

                                        Case 2 'NPC Repair the Object
                                            If ExamineBit(NPC(Map(MapNum).NPC).flags, 1) = True Then
                                                B = .Inv(.CurrentRepairTar).Object 'Object
                                                If B > 0 Then 'Slot isn't empty
                                                    A = GetRepairCost(Index, .CurrentRepairTar) 'Cost
                                                    C = FindInvObject(Index, 6) 'Money Slot
                                                    If C > 0 Then 'Has money
                                                        If .Inv(C).Value >= A Then 'Has the Cash
                                                            TakeObj Index, 6, A 'Take Cash
                                                            TakeObj Index, B, 1 'Take Old Item
                                                            GiveObj Index, B, 1 'Give New Item
                                                            SendSocket Index, Chr$(98) + Chr$(2) + Chr$(B)
                                                        Else
                                                            SendSocket Index, Chr$(16) + Chr$(33)
                                                        End If
                                                    Else
                                                        SendSocket Index, Chr$(16) + Chr$(33)
                                                    End If
                                                Else
                                                    SendSocket Index, Chr$(16) + Chr$(34)
                                                End If
                                            End If
                                                
                                        End Select
                                    Else
                                        Hacker Index, "R.1"
                                    End If
                                Case Else
                                    Hacker Index, "B.3"
                                End Select
                            End Select
                End If
                GoTo LoopRead
            End If
        End If
        .SocketData = SocketData
    End With
End Sub

Private Function GetString(St As String, Num As Long) As String
    Dim A As Long, B As Long, C As Long
    
    B = 1
    If Num > 0 Then
        For A = 1 To Num
GetAgain:
                C = B
                B = InStr(B + 1, St, " ")
                If B = C + 1 Then GoTo GetAgain
                
            If B = 0 Then Exit For
        Next A
        If B <> 0 Then GetString = Mid$(St, B + 1, Len(St) - B)
    End If
End Function

Sub ResetMap(MapNum As Long)
    Dim A As Long, X As Long, Y As Long
    Dim NumPlayers As Long
    Dim St1 As String
    
    With Map(MapNum)
        NumPlayers = .NumPlayers
        For A = 0 To 49
            With .Object(A)
                If .Object > 0 Then
                    If Map(MapNum).Tile(.X, .Y).Att <> 5 Then
                        .Object = 0
                        If NumPlayers > 0 Then
                            St1 = St1 + DoubleChar(2) + Chr$(15) + Chr$(A)
                        End If
                    End If
                End If
            End With
        Next A
        For A = 0 To 9
            With .Door(A)
                If .Att > 0 Then
                    Map(MapNum).Tile(.X, .Y).Att = .Att
                    If NumPlayers > 0 Then
                        St1 = St1 + DoubleChar(2) + Chr$(37) + Chr$(A)
                    End If
                    .Att = 0
                End If
            End With
        Next A
        If ExamineBit(.flags, 3) = True Then
            'Create Monsters
            For A = 0 To 5
                St1 = St1 + NewMapMonster(MapNum, A)
            Next A
        Else
            'Clear Monsters
            For A = 0 To 5
                If .Monster(A).Monster > 0 Then
                    .Monster(A).Monster = 0
                    If NumPlayers > 0 Then
                        St1 = St1 + DoubleChar(2) + Chr$(39) + Chr$(A)
                    End If
                End If
            Next A
        End If
        If NumPlayers > 0 Then
            SendToMapRaw MapNum, St1
        End If
        For Y = 0 To 11
            For X = 0 To 11
                With Map(MapNum).Tile(X, Y)
                    If .Att = 7 Then
                        NewMapObject MapNum, CLng(.AttData(0)), CLng(.AttData(1)) * 65536 + CLng(.AttData(2)) * 256& + CLng(.AttData(3)), X, Y, True
                    End If
                End With
            Next X
        Next Y
        .ResetTimer = 0
    End With
End Sub
Sub SendCharacterData(Index As Long)
    With Player(Index)
        If .Class > 0 Then
            SendSocket Index, Chr$(3) + Chr$(.Class) + Chr$(.Gender) + Chr$(.Sprite) + Chr$(.HP) + Chr$(.Energy) + Chr$(.Mana) + Chr$(.MaxHP) + Chr$(.MaxEnergy) + Chr$(.MaxMana) + Chr$(.Strength) + Chr$(.Agility) + Chr$(.Endurance) + Chr$(.Intelligence) + Chr$(.level) + Chr$(.Status) + Chr$(.Guild) + Chr$(.GuildRank) + Chr$(.Access) + Chr$(Index) + QuadChar(.Experience) + DoubleChar(CLng(.StatPoints)) + .Name + Chr$(0) + .desc
        Else
            SendSocket Index, Chr$(3)
        End If
    End With
End Sub

Sub SendDataPacket(Index As Long, StartNum As Long)
    Dim A As Long, St1 As String
    
    For A = StartNum To 255
        If NPC(A).Name <> "" Then
            With NPC(A)
                St1 = St1 + DoubleChar(2 + Len(.Name)) + Chr$(85) + Chr$(A) + .Name
            End With
        End If
        If Hall(A).Name <> "" Then
            With Hall(A)
                St1 = St1 + DoubleChar(2 + Len(.Name)) + Chr$(82) + Chr$(A) + .Name
            End With
        End If
        If Guild(A).Name <> "" Then
            With Guild(A)
                St1 = St1 + DoubleChar(2 + Len(.Name)) + Chr$(70) + Chr$(A) + .Name
            End With
        End If
        If Object(A).Picture > 0 Then
            With Object(A)
                St1 = St1 + DoubleChar(4 + Len(.Name)) + Chr$(31) + Chr$(A) + Chr$(.Picture) + Chr$(.Type) + .Name
            End With
        End If
        If Monster(A).Sprite > 0 Then
            With Monster(A)
                St1 = St1 + DoubleChar(3 + Len(.Name)) + Chr$(32) + Chr$(A) + Chr$(.Sprite) + .Name
            End With
        End If
        If Len(St1) >= 1024 Then
            If A < 255 Then
                St1 = St1 + DoubleChar(3) + Chr$(35) + Chr$(24) + Chr$(A + 1)
            Else
                St1 = St1 + DoubleChar(2) + Chr$(35) + Chr$(23)
            End If
            SendRaw Index, St1
            Exit Sub
        End If
    Next A
    St1 = St1 + DoubleChar(2) + Chr$(35) + Chr$(23)
    SendRaw Index, St1
End Sub
Sub SetGodAccess(Index As Long, Access As Long)
    Dim GodAccountNum As Long
        
    If Index >= 1 And Index <= MaxUsers And Access >= 0 And Access <= 10 Then
        With Player(Index)
            If .Mode = modePlaying Then
                .Access = Access
                SendSocket Index, Chr$(65) + Chr$(Access)
                
                If .Access > 0 Then
                    SendAllBut Index, Chr$(91) + Chr$(Index) + Chr$(3)
                    .Status = 3
                    GodAccountNum = FindGodAccount(.ComputerID)
                    
                    If GodAccountNum > 0 Then
                        GodRS.Index = "number"
                        GodRS.Seek "=", GodAccountNum
                        If GodRS.NoMatch = False Then
                            GodRS.Edit
                            GodRS!User = .User
                            GodRS!ComputerID = EncryptString(.ComputerID)
                            GodRS!Access = Access
                            GodRS.Update
                        End If
                        
                        With GodData(GodAccountNum)
                            .User = Player(Index).User
                            .ComputerID = Trim$(Player(Index).ComputerID)
                            .Access = Access
                            .InUse = True
                        End With
                    Else
                        GodAccountNum = FreeGodNum()
                        If GodAccountNum <= 50 Then
                            GodRS.AddNew
                            GodRS!number = GodAccountNum
                            GodRS!User = .User
                            GodRS!ComputerID = EncryptString(.ComputerID)
                            GodRS!Access = Access
                            GodRS.Update
                        End If
                        With GodData(GodAccountNum)
                            .User = Player(Index).User
                            .ComputerID = Trim$(Player(Index).ComputerID)
                            .Access = Access
                            .InUse = True
                        End With
                    End If
                Else
                    GodAccountNum = FindGodAccount(.ComputerID)
                    If GodAccountNum > 0 Then
                        GodData(GodAccountNum).InUse = False
                        GodRS.Index = "number"
                        GodRS.Seek "=", GodAccountNum
                        If GodRS.NoMatch = False Then
                            GodRS.Delete
                        End If
                    End If
                    SendAllBut Index, Chr$(91) + Chr$(Index) + Chr$(0)
                    .Status = 0
                End If
            End If
        End With
    End If
End Sub

Function SpawnMapMonster(MapNum As Long, MonsterNum As Long, MonsterType As Long, TX As Long, TY As Long)
    With Map(MapNum).Monster(MonsterNum)
        .Monster = MonsterType
        .X = TX
        .Y = TY
        .HP = Monster(.Monster).HP
        .Distance = Monster(.Monster).Sight
        .Target = 0
        .D = Int(Rnd * 4)
        SpawnMapMonster = DoubleChar(6) + Chr$(38) + Chr$(MonsterNum) + Chr$(.Monster) + Chr$(.X) + Chr$(.Y) + Chr$(.D)
    End With
End Function

Function ValidName(St As String) As Boolean
    Dim A As Long, B As Long
    If Len(St) > 0 Then
        For A = 1 To Len(St)
            B = Asc(Mid$(St, A, 1))
            If (B < 48 Or B > 57) And (B < 65 Or B > 90) And (B < 97 Or B > 122) And B <> 32 And B <> 95 Then
                ValidName = False
                Exit Function
            End If
        Next
    End If
    ValidName = True
End Function
Function WindowProc(ByVal hw As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim TempSocket As Long
    
    If uMsg >= 1029 And uMsg < 1029 + MaxUsers Then
        Dim Index As Long
        Index = uMsg - 1028
        Select Case lParam And 255
            Case FD_CLOSE
                AddSocketQue Index
            Case FD_READ
                ReadClientData Index
        End Select
    End If
    Select Case uMsg
        Case 1025 'Listening Socket
            Select Case lParam And 255
                Case FD_ACCEPT
                    If lParam = FD_ACCEPT Then
                        Dim NewPlayer As Long, Address As sockaddr
                        Dim ClientIP As String, BanNum As Long, Banned As Boolean
                        Dim A As Long, DupIP As Boolean
                        
                        NewPlayer = FreePlayer()
                        If NewPlayer > 0 Then
                            With Player(NewPlayer)
                                .Socket = accept(ListeningSocket, Address, sockaddr_size)
                                If Not .Socket = INVALID_SOCKET Then
                                    ClientIP = GetPeerAddress(.Socket)
                                    DupIP = False
                                    #If CheckIPDupe = True Then
                                        For A = 1 To MaxUsers
                                            With Player(A)
                                                If .InUse = True And .IP = ClientIP Then
                                                    DupIP = True
                                                    Exit For
                                                End If
                                            End With
                                        Next A
                                    #End If
                                    
                                    If DupIP = True Then
                                        'Duplicate IP
                                        SendData .Socket, DoubleChar(59) + Chr$(0) + Chr$(0) + "You may not log in multiple times from the same computer!"
                                        closesocket .Socket
                                    Else
                                        If WSAAsyncSelect(.Socket, gHW, ByVal 1028 + NewPlayer, ByVal FD_READ Or FD_WRITE Or FD_CONNECT Or FD_CLOSE) = 0 Then
                                            .InUse = True
                                            .Mode = modeNotConnected
                                            .IP = ClientIP
                                            .Class = 0
                                            .SocketData = ""
                                            .LastMsg = GetTickCount - 50
                                            .ClientVer = ""
                                            .FloodTimer = GetTickCount
                                            PrintLog ("Connection accepted from " + .IP)
                                            NumUsers = NumUsers + 1
                                            frmMain.mnuDatabase.Enabled = False
                                            frmMain.Caption = TitleString + " [" + CStr(NumUsers) + "]"
                                        Else
                                            closesocket .Socket
                                        End If
                                    End If
                                Else
                                    closesocket .Socket
                                End If
                            End With
                        Else
                            TempSocket = accept(ListeningSocket, Address, sockaddr_size)
                            SendData TempSocket, DoubleChar(2) + Chr$(0) + Chr$(4)
                            closesocket TempSocket
                        End If
                    End If
            End Select
    End Select
    WindowProc = CallWindowProc(lpPrevWndProc, hw, uMsg, wParam, lParam)
End Function
Sub ParseString(St1)
    Dim St As String, A As Long, B As Long, C As Long
    If Mid$(St1, Len(St1), 1) = Chr$(13) Or Mid$(St1, Len(St1), 1) = Chr$(10) Then
        St1 = Mid$(St1, 1, Len(St1) - 1)
    End If
    If Mid$(St1, Len(St1), 1) = Chr$(13) Or Mid$(St1, Len(St1), 1) = Chr$(10) Then
        St1 = Mid$(St1, 1, Len(St1) - 1)
    End If
    For A = 1 To Len(St1)
        If Asc(Mid$(St1, 1, 1)) < 32 Then
            St1 = Mid$(St1, 2)
        Else
            Exit For
        End If
    Next A
    St = St1
    Suffix = ""
    Prefix = ""
    If Mid$(St, 1, 1) = ":" Then
        A = InStr(St, " ")
        Prefix = Mid$(St, 2, A - 2)
        St = Mid$(St, A + 1)
    End If
    St1 = St
    A = InStr(St, ":")
    If A > 0 Then
        Suffix = Mid$(St, A + 1, Len(St) - A)
        St = Mid$(St, 1, A - 1)
    End If
    B = 1
    Erase Word
    For A = 1 To 10
TryAgain9:
        C = InStr(B, St, " ")
        If C - B = 0 Then B = B + 1: GoTo TryAgain9
        If C <> 0 Then
                Word(A) = Mid$(St, B, C - B)
        Else
                Word(A) = Mid$(St, B, Len(St) - B + 1)
                Exit For
        End If
        B = C + 1
    Next A
End Sub
Sub CloseIRC()
    With IRC
        If Not .Socket = INVALID_SOCKET Then
            closesocket .Socket
            .Socket = INVALID_SOCKET
        End If
        .Connected = False
        .Connecting = False
    End With
End Sub
Function ExamineBit(bytByte As Byte, Bit As Byte) As Byte
    ExamineBit = ((bytByte And (2 ^ Bit)) > 0)
End Function
Sub CloseClientSocket(Index As Long)
Dim A As Long
    With Player(Index)
        If .InUse = True Then
            'Decrement User Num
            NumUsers = NumUsers - 1
            If NumUsers = 0 Then frmMain.mnuDatabase.Enabled = True
            frmMain.Caption = TitleString + " [" + CStr(NumUsers) + "]"
            
            For A = 1 To MaxPlayerTimers
                If .ScriptTimer(A) > 0 Then
                    Parameter(0) = Index
                    .ScriptTimer(A) = 0
                    RunScript .Script(A)
                End If
            Next A
            
            If .Mode = modePlaying Then
                Parameter(0) = Index
                RunScript "PARTGAME"
            End If
            
            'Close Socket
            If Not .Socket = INVALID_SOCKET Then
                closesocket .Socket
                .Socket = INVALID_SOCKET
            End If
            
            If Not .Class = 0 Then
                If .Status = 2 Then .Status = 0
                SavePlayerData Index
            End If
            
            PrintLog "Connection closed from " + .IP + " [" + Player(Index).Name + "]"
            
            'Clear Socket Data
            .InUse = False
            .SocketData = ""
            .Class = 0
            .User = ""
            .Name = ""
            
            'Send Quit Message
            If .Mode = modePlaying Then
                .Mode = modeNotConnected
                SendAll Chr$(7) + Chr$(Index)
                If .Map > 0 Then
                    Partmap Index
                    .Map = 0
                End If
            Else
                .Mode = modeNotConnected
            End If
        End If
    End With
End Sub
Sub CreateDatabase()
    Dim A As Long
    Dim Td As TableDef
    Dim NewField As Field
    Dim NewIndex As Index
        
    'Create Database
    #If PublicServer = True Then
        Set DB = WS.CreateDatabase("server.dat", ";pwd=" + Chr$(100) + Chr$(114) + Chr$(97) + Chr$(99) + Chr$(111) + dbLangGeneral, dbEncrypt + dbVersion30)
    #Else
        Set DB = WS.CreateDatabase("server.dat", dbLangGeneral, dbVersion30)
    #End If
    
    CreateAccountsTable
    CreateGuildsTable
    CreateNPCsTable
    CreateMonstersTable
    CreateObjectsTable
    CreateDataTable
    CreateMapsTable
    CreateMessagesTable
    CreatePostsTable
    CreateBansTable
    CreateHallsTable
    CreateScriptsTable
    CreateGodTable
End Sub
Function DoubleChar(Num As Long) As String
    DoubleChar = Chr$(Int(Num / 256)) + Chr$(Num Mod 256)
End Function
Function TripleChar(Num As Long) As String
    TripleChar = Chr$(Int(Num / 65536)) + Chr$(Int((Num Mod 65536) / 256)) + Chr$(Num Mod 256)
End Function
Function QuadChar(Num As Long) As String
    QuadChar = Chr$(Int(Num / 16777216) Mod 256) + Chr$(Int(Num / 65536) Mod 256) + Chr$(Int(Num / 256) Mod 256) + Chr$(Num Mod 256)
End Function
Function Exists(Filename As String) As Boolean
     Exists = (Dir(Filename) <> "")
End Function
Function GetInt(Chars As String) As Long
    GetInt = CLng(Asc(Mid$(Chars, 1, 1))) * 256& + CLng(Asc(Mid$(Chars, 2, 1)))
End Function
Sub GetWords(St As String)
    Dim A As Long, B As Long, C As Long
    B = 1
    Erase Word
    For A = 1 To 50
TryAgain:
        C = InStr(B, St, " ")
        If C - B = 0 Then B = B + 1: GoTo TryAgain
        If C <> 0 Then
                Word(A) = Mid$(St, B, C - B)
        Else
                Word(A) = Mid$(St, B, Len(St) - B + 1)
                Exit For
        End If
        B = C + 1
    Next A
End Sub
Sub GetSections(St)
    Dim A As Long, B As Long, C As Long
    B = 1
    Erase Word
    For A = 1 To 10
        C = InStr(B, St, Chr$(0))
        If C - B = 0 Then
            Word(A) = ""
        ElseIf C <> 0 Then
            Word(A) = Mid$(St, B, C - B)
        Else
            Word(A) = Mid$(St, B, Len(St) - B + 1)
            Exit For
        End If
        B = C + 1
    Next A
End Sub
Function Nick(UserHost As String) As String
    Dim A As Long
    
    A = InStr(UserHost, "!")
    If A > 0 Then
        Nick = Mid$(UserHost, 1, A - 1)
    Else
        Nick = UserHost
    End If
End Function
Public Sub Hook()
#If DEBUGWINDOWPROC Then
    On Error Resume Next
    Set m_SCHook = CreateWindowProcHook
    If Err Then
        MsgBox Err.Description
        Err.Clear
        Unhook
        Exit Sub
    End If
    On Error GoTo 0
    With m_SCHook
        .SetMainProc AddressOf WindowProc
        lpPrevWndProc = SetWindowLong(gHW, GWL_WNDPROC, .ProcAddress)
        .SetDebugProc lpPrevWndProc
    End With
#Else
    lpPrevWndProc = SetWindowLong(gHW, GWL_WNDPROC, AddressOf WindowProc)
#End If
End Sub
Public Sub Unhook()
    SetWindowLong gHW, GWL_WNDPROC, lpPrevWndProc
End Sub
Sub SavePlayerData(Index)
    Dim A As Long, St As String

    With Player(Index)
        UserRS.Index = "User"
        UserRS.Seek "=", .User
        UserRS.Edit
        
        .Bookmark = UserRS.Bookmark
        UserRS!Access = .Access
        UserRS!ComputerID = EncryptString(.ComputerID)
        
        'Character Data
        UserRS!CharNum = .CharNum
        UserRS!Name = .Name
        UserRS!Class = .Class
        UserRS!Gender = .Gender
        UserRS!Sprite = .Sprite
        UserRS!desc = .desc
        
        'Position Data
        UserRS!Map = .Map
        UserRS!X = .X
        UserRS!Y = .Y
        UserRS!D = .D
        
        'Character Vital Stats
        UserRS!MaxHP = .MaxHP
        UserRS!MaxEnergy = .MaxEnergy
        UserRS!MaxMana = .MaxMana
        UserRS!HP = .HP
        UserRS!Energy = .Energy
        UserRS!Mana = .Mana
        
        'Character Physical Stats
        UserRS!Strength = .Strength
        UserRS!Agility = .Agility
        UserRS!Endurance = .Endurance
        UserRS!Intelligence = .Intelligence
        UserRS!level = .level
        UserRS!Experience = .Experience
        UserRS!StatPoints = .StatPoints
        
        'Misc. Data
        UserRS!Bank = .Bank
        UserRS!Status = .Status
        UserRS!LastPlayed = CLng(Date)
        UserRS!TimeLeft = .TimeLeft
        
        'Inventory Data
        For A = 1 To 20
            UserRS.Fields("InvObject" + CStr(A)).Value = .Inv(A).Object
            UserRS.Fields("InvValue" + CStr(A)).Value = .Inv(A).Value
        Next A
        
        'Equipped Objects
        For A = 1 To 6
            UserRS.Fields("EquippedObject" + CStr(A)).Value = .EquippedObject(A)
        Next A
        
        'Mail
        For A = 1 To 20
            UserRS.Fields("Msg" + CStr(A)) = .Msg(A)
        Next A
        
        'Flags
        St = ""
        For A = 0 To 127
            With .Flag(A)
                St = St + QuadChar(.Value) + QuadChar(.ResetCounter)
            End With
        Next A
        UserRS!flags = St
        
        UserRS.Update
    End With
End Sub
Sub SendAll(ByVal St As String)
    Dim A As Long
    For A = 1 To MaxUsers
        With Player(A)
            If .Mode = modePlaying Then
                If SendData(.Socket, DoubleChar(Len(St)) + St) = SOCKET_ERROR Then
                    'CloseClientSocket A
                End If
            End If
        End With
    Next A
End Sub
Sub SendToConnected(ByVal St As String)
    Dim A As Long
    For A = 1 To MaxUsers
        With Player(A)
            If .Mode > 0 Then
                If SendData(.Socket, DoubleChar(Len(St)) + St) = SOCKET_ERROR Then
                    'CloseClientSocket A
                End If
            End If
        End With
    Next A
End Sub
Sub SendAllBut(ByVal Index As Long, ByVal St As String)
    Dim A As Long
    For A = 1 To MaxUsers
        With Player(A)
            If .Mode = modePlaying And A <> Index Then
                If SendData(.Socket, DoubleChar(Len(St)) + St) = SOCKET_ERROR Then
                    'CloseClientSocket A
                End If
            End If
        End With
    Next A
End Sub
Sub SendAllButRaw(ByVal Index As Long, ByVal St As String)
    Dim A As Long
    For A = 1 To MaxUsers
        With Player(A)
            If .Mode = modePlaying And A <> Index Then
                If SendData(.Socket, St) = SOCKET_ERROR Then
                    'CloseClientSocket A
                End If
            End If
        End With
    Next A
End Sub
Sub SendAllButBut(ByVal Index1 As Long, ByVal Index2 As Long, ByVal St As String)
    Dim A As Long
    For A = 1 To MaxUsers
        With Player(A)
            If .Mode = modePlaying And A <> Index1 And A <> Index2 Then
                If SendData(.Socket, DoubleChar(Len(St)) + St) = SOCKET_ERROR Then
                    'CloseClientSocket A
                End If
            End If
        End With
    Next A
End Sub

Sub SendToGods(ByVal St As String)
    Dim A As Long
    For A = 1 To MaxUsers
        With Player(A)
            If .Mode = modePlaying And .Access > 0 Then
                If SendData(.Socket, DoubleChar(Len(St)) + St) = SOCKET_ERROR Then
                    'CloseClientSocket A
                End If
            End If
        End With
    Next A
End Sub
Sub SendToGodsAllBut(Index As Long, ByVal St As String)
    Dim A As Long
    For A = 1 To MaxUsers
        With Player(A)
            If .Mode = modePlaying And .Access > 0 And Index <> A Then
                If SendData(.Socket, DoubleChar(Len(St)) + St) = SOCKET_ERROR Then
                    'CloseClientSocket A
                End If
            End If
        End With
    Next A
End Sub

Sub SendToMap(ByVal MapNum As Long, ByVal St As String)
    Dim A As Long
    For A = 1 To MaxUsers
        With Player(A)
            If .Mode = modePlaying And .Map = MapNum Then
                If SendData(.Socket, DoubleChar(Len(St)) + St) = SOCKET_ERROR Then
                    'CloseClientSocket A
                End If
            End If
        End With
    Next A
End Sub
Sub SendToMapRaw(ByVal MapNum As Long, ByVal St As String)
    Dim A As Long
    For A = 1 To MaxUsers
        With Player(A)
            If .Mode = modePlaying And .Map = MapNum Then
                If SendData(.Socket, St) = SOCKET_ERROR Then
                    'CloseClientSocket A
                End If
            End If
        End With
    Next A
End Sub
Sub ShutdownServer()
    Dim A As Long
    For A = 1 To MaxUsers
        If Player(A).InUse = True Then
            AddSocketQue A
        End If
    Next A
    
    SaveFlags
    SaveObjects
    
    UserRS.Close
    GuildRS.Close
    NPCRS.Close
    MonsterRS.Close
    ObjectRS.Close
    MsgRS.Close
    DataRS.Close
    PostRS.Close
    MapRS.Close
    BanRS.Close
    DB.Close
    WS.Close
    If ListeningSocket <> INVALID_SOCKET Then
        closesocket ListeningSocket
    End If
    EndWinsock
    Unhook
End Sub
Sub SendToMapAllBut(ByVal MapNum As Long, ByVal Index As Long, ByVal St As String)
    Dim A As Long
    For A = 1 To MaxUsers
        With Player(A)
            If .Mode = modePlaying And .Map = MapNum And Index <> A Then
                If SendData(.Socket, DoubleChar(Len(St)) + St) = SOCKET_ERROR Then
                    'CloseClientSocket A
                End If
            End If
        End With
    Next A
End Sub
Sub SendSocket(ByVal Index As Long, ByVal St As String)
    With Player(Index)
        If .InUse = True Then
            If SendData(.Socket, DoubleChar(Len(St)) + St) = SOCKET_ERROR Then
                'CloseClientSocket Index
            End If
        End If
    End With
End Sub
Sub SendRaw(ByVal Index As Long, ByVal St As String)
    With Player(Index)
        If .InUse = True Then
            If SendData(.Socket, St) = SOCKET_ERROR Then
                'CloseClientSocket Index
            End If
        End If
    End With
End Sub
Sub PrintLog(St)
    With frmMain.lstLog
        .AddItem St
        If .ListCount > 30 Then .RemoveItem 0
        If .ListIndex = .ListCount - 2 Then .ListIndex = .ListCount - 1
    End With
End Sub

Sub CheckBan(Index As Long, ComputerID As String)
Dim BanNum As Long, Banned As Boolean

BanNum = FindBan(ComputerID)
If BanNum > 0 Then
    With Ban(BanNum)
        If CLng(Date) >= .UnbanDate Then
            'Unban
            .ComputerID = ""
                BanRS.Seek "=", BanNum
                If BanRS.NoMatch = False Then
                    BanRS.Delete
                End If
                Banned = False
                Exit Sub
        End If
    End With
    SendSocket Index, Chr$(0) + Chr$(3) + QuadChar(Ban(BanNum).UnbanDate) + Ban(BanNum).Reason
    CloseClientSocket Index
End If
End Sub

Sub CreateGodTable()
    Dim Td As TableDef
    Dim NewField As Field
    Dim NewIndex As Index

    'Create Gods Table
    Set Td = DB.CreateTableDef("Gods")
    Set NewField = Td.CreateField("Number", dbByte)
    Td.Fields.Append NewField
    Set NewField = Td.CreateField("User", dbText, 15)
    NewField.AllowZeroLength = True
    Td.Fields.Append NewField
    Set NewField = Td.CreateField("ComputerID", dbText, 30)
    NewField.AllowZeroLength = True
    Td.Fields.Append NewField
    Set NewField = Td.CreateField("Access", dbByte)
    Td.Fields.Append NewField
    
    'Create Indexes
    Set NewIndex = Td.CreateIndex("Number")
    NewIndex.Primary = True
    NewIndex.Unique = True
    Set NewField = NewIndex.CreateField("Number")
    NewIndex.Fields.Append NewField
    Td.Indexes.Append NewIndex
    
    'Append Bans Table
    DB.TableDefs.Append Td
End Sub
Function EncryptString(St As String) As String
Dim TempStr As String, TempStr2 As String
Dim A As Integer, TmpNum As Integer

TempStr = ""
TempStr2 = ""

For A = 1 To Len(St)
    TempStr = Mid$(St, A, 1)
    TmpNum = Asc(TempStr)
    TempStr2 = TempStr2 + Chr$(TmpNum + 3 - 10)
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
    TempStr2 = TempStr2 + Chr$(TmpNum - 3 + 10)
Next A

DecipherString = Trim$(TempStr2)
End Function

Function AddSocketQue(Index As Long) As Integer
Dim A As Integer

For A = 1 To MaxUsers
    If CloseSocketQue(A) = Index Then
        Exit Function
    End If
Next A

For A = 1 To MaxUsers
    If CloseSocketQue(A) = 0 Then
        CloseSocketQue(A) = Index
        Exit For
    End If
Next A
End Function

Sub GiveStartingEQ(Index As Long)
    Dim A As Long, B As Long, C As Long
    If Index >= 1 And Index <= MaxUsers Then
    
    With Player(Index)
    For A = 1 To 8
        If World.StartObjects(A) > 0 Then
            B = World.StartObjects(A)
            C = World.StartObjValues(A)
            .Inv(A).Object = B

        Select Case Object(B).Type
            Case 1, 2, 3, 4 'Weapon, Shield, Armor, Helmut
                .Inv(A).Value = CLng(Object(B).Data(0)) * 10
            Case 6 'Money
                .Inv(A).Value = C
            Case 8 'Ring
                .Inv(A).Value = CLng(Object(B).Data(1)) * 10
            Case Else
                .Inv(A).Value = 0
        End Select
        End If
    Next A
    End With
    End If
End Sub
Function GetRepairCost(Index As Long, Slot As Integer) As Long
Dim A As Long, B As Long, C As Long
If Index >= 1 And Index <= MaxUsers Then
    If Slot >= 0 And Slot <= 20 Then
        Select Case Object(Player(Index).Inv(Slot).Object).Type
            Case 1, 2, 3, 4, 8 'Weapon, Shield, Armor, Helmet, Ring
                A = Object(Player(Index).Inv(Slot).Object).Type
            Case Else
                A = 0
        End Select
    End If
        
        If A > 0 Then
            Select Case A
                Case 1, 2, 3, 4 'Weapon, Shield, Armor, Helmet
                    C = Object(Player(Index).Inv(Slot).Object).Data(0) - (Player(Index).Inv(Slot).Value / 10)
                    B = B + (C * Cost_Per_Durability)
                    B = B + (Object(Player(Index).Inv(Slot).Object).Data(1) * Cost_Per_Strength)
                    GetRepairCost = B
                    Exit Function
                Case 8 'Ring
                    C = Object(Player(Index).Inv(Slot).Object).Data(1) - (Player(Index).Inv(Slot).Value / 10)
                    B = B + (C * Cost_Per_Durability)
                    B = B + (Object(Player(Index).Inv(Slot).Object).Data(2) * Cost_Per_Modifier)
                    GetRepairCost = B
                    Exit Function
            End Select
        Else
            GetRepairCost = 0
        End If
End If
End Function
Function GetObjectDur(ByVal Index As Long, ByVal Slot As Long) As Long
Dim Percent As Single
Dim Display As Boolean
    Select Case Object(Player(Index).Inv(Slot).Object).Type
    Case 1, 2, 3, 4
        Display = True
    Case Else
        Display = False
    End Select
If Display = True Then
    'Percent = Int((Player(Index).Inv(ObjectNum).Value * 5) / Object(Player(Index).Inv(ObjectNum).Object).Data(0))
    Percent = Player(Index).Inv(Slot).Value / (Object(Player(Index).Inv(Slot).Object).Data(0) * 10)
    Percent = Int(Percent * 100)
    If Percent > 100 Then Percent = 100
    GetObjectDur = Percent
Else
    GetObjectDur = 0
End If
End Function
