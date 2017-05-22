Attribute VB_Name = "modArray"
Option Explicit

Public ReceiveArray(0 To 255) As Long

Type NPCSaleItemData
    GiveObject As Integer
    GiveValue As Long
    TakeObject As Integer
    TakeValue As Long
End Type

Type ItemBankData
    Object As Integer
    value As Long
    ItemPrefix As Byte
    ItemSuffix As Byte
End Type

Type SkillData
    Level As Byte
    Experience As Long
End Type

Type HotkeyData
    Hotkey As Byte
Type As Byte
    ScrollPosition As Byte
End Type

Type ProjectileData
    Sprite As Integer
    Type As Byte
    Frame As Byte
    TotalFrames As Integer
    TargetType As Byte
    TargetNum As Byte
    TargetX As Long
    TargetY As Long
    SourceX As Byte
    SourceY As Byte
    X As Long
    Y As Long
    TimeStamp As Long
    LoopCount As Integer
    CurLoop As Integer
    speed As Long
    EndSound As Long
    offset As Byte
    Creator As Byte
    Magic As Byte
    Damage As Byte
    Alternate As Boolean
End Type

Type PlayerData
    name As String
    LastMessage As String
    RepCount As Long
    Map As Long
    Sprite As Integer
    status As Byte
    X As Byte
    Y As Byte
    XO As Long
    YO As Long
    D As Byte
    A As Byte
    W As Long
    WalkStep As Long
    Guild As Byte
    Color As Long
    Ignore As Boolean
    IsDead As Boolean
    HP As Byte
    MaxHP As Byte
End Type

Type MacroData
    Text As String
    LineFeed As Boolean
End Type

Type ObjectData
    name As String
    Type As Byte
    flags As Byte
    Picture As Integer
    Modifier As Byte
    Data2 As Byte
    MaxDur As Byte
    ClassReq As Byte
    LevelReq As Byte
    Version As Byte
    SellPrice As Integer
End Type

Type MonsterData
    name As String
    Sprite As Integer
    Version As Byte
    MaxLife As Integer
    flags As Byte
End Type

Type MapStartLocationData
    Map As Integer
    X As Byte
    Y As Byte
End Type

Type MapDoorData
    Att As Byte
    BGTile1 As Integer
    X As Byte
    Y As Byte
End Type

Type TileData
    Ground As Integer
    Ground2 As Integer
    BGTile1 As Integer
    BGTile2 As Integer
    FGTile As Integer
    FGTile2 As Integer
    Att As Byte
    AttData(0 To 3) As Byte
    Att2 As Byte
End Type

Type MapMonsterData
    Monster As Integer
    Life As Integer
    X As Byte
    Y As Byte
    XO As Long
    YO As Long
    D As Byte
    A As Byte
    W As Byte
    HPBar As Boolean
End Type

Type MapObjectData
    Object As Integer
    X As Byte
    Y As Byte
    XOffset As Byte
    YOffset As Byte
    ItemPrefix As Byte
    ItemSuffix As Byte
    value As Long
    PickedUp As Long
End Type

Type MapMonsterSpawnData
    Monster As Integer
    Rate As Byte
End Type

Type MapData
    name As String
    ExitUp As Integer
    ExitDown As Integer
    ExitLeft As Integer
    ExitRight As Integer
    Tile(0 To 11, 0 To 11) As TileData
    Object(0 To MaxMapObjects) As MapObjectData
    Monster(0 To MaxMonsters) As MapMonsterData
    MonsterSpawn(0 To 9) As MapMonsterSpawnData
    Door(0 To 9) As MapDoorData
    BootLocation As MapStartLocationData
    DeathLocation As MapStartLocationData
    NPC As Integer
    MIDI As Byte
    flags As Byte
    Flags2 As Byte
    Version As Long
End Type

Type InvObjData
    Object As Integer
    value As Long
    EquippedNum As Byte
    ItemPrefix As Byte
    ItemSuffix As Byte
End Type

Type GuildData
    name As String
    MemberCount As Byte
End Type

Type MagicData
    name As String
    MagicLevel As Byte
    MagicExperience As Long
    Level As Byte
    Class As Byte
    Version As Byte
    Description As String
    Icon As Integer
    IconType As Byte
    CastTimer As Integer
End Type

Type HallData
    name As String
    Version As Byte
End Type

Type NPCData
    name As String
    Version As Byte
    flags As Byte
    SaleItem(0 To 9) As NPCSaleItemData
End Type

Type EquippedObjectData
    Object As Integer
    value As Long
    ItemPrefix As Byte
    ItemSuffix As Byte
End Type

Type GuildDeclarationData
    Guild As Byte
    Type As Byte
End Type

Type CharacterData
    name As String
    Class As Byte
    Gender As Byte
    Sprite As Integer

    MaxHP As Byte
    MaxEnergy As Byte
    MaxMana As Byte

    Experience As Long
    Level As Byte
    StatPoints As Integer

    PhysicalAttack As Byte
    PhysicalDefense As Byte
    MagicDefense As Byte

    Skill(1 To 10) As SkillData
    Hotkey(1 To 12) As HotkeyData
    ItemBank(0 To 29) As ItemBankData

    status As Byte
    Access As Byte
    index As Byte
    Guild As Byte
    GuildRank As Byte

    Projectile As Boolean
    Ammo As Byte

    IsDead As Boolean
    LastMove As Long
    SpeedTick As Long
    EnergyTick As Long

    GuildDeclaration(0 To 4) As GuildDeclarationData
    EquippedObject(1 To 5) As EquippedObjectData

    Desc As String
    Inv(1 To MaxInvObjects) As InvObjData
End Type

Type ClassData
    name As String
    StartHP As Byte
    StartEnergy As Byte
    StartMana As Byte
End Type

Type FloatTextData
    Text As String
    Color As Byte
    Static As Boolean
    X As Byte
    Y As Byte
    FloatY As Integer
    InUse As Boolean
End Type

Type PrefixData
    name As String
    ModificationType As Byte
    ModificationValue As Byte
    OccursNaturally As Byte
    Version As Byte
End Type

Type WorldProperties
    StatStrength As Byte
    StatEndurance As Byte
    StatIntelligence As Byte
    StatConcentration As Byte
    StatConstitution As Byte
    StatStamina As Byte
    StatWisdom As Byte
    ObjMoney As Byte
    Cost_Per_Durability As Integer
    Cost_Per_Strength As Integer
    Cost_Per_Modifier As Integer
    GuildJoinLevel As Byte
    GuildNewLevel As Byte
    GuildJoinCost As Long
    GuildNewCost As Long
End Type

Public Map As MapData, EditMap As MapData, ClipboardMap As MapData
Attribute EditMap.VB_VarUserMemId = 1073741825
Attribute ClipboardMap.VB_VarUserMemId = 1073741825
Public StatusColors(0 To 100) As Long
Attribute StatusColors.VB_VarUserMemId = 1073741828
Public Player(1 To MaxUsers) As PlayerData
Attribute Player.VB_VarUserMemId = 1073741829
Public Magic(1 To MaxMagic) As MagicData
Public Projectile(1 To MaxProjectiles) As ProjectileData
Attribute Projectile.VB_VarUserMemId = 1073741830
Public Monster(1 To MaxTotalMonsters) As MonsterData
Attribute Monster.VB_VarUserMemId = 1073741831
Public Object(1 To MaxObjects) As ObjectData
Attribute Object.VB_VarUserMemId = 1073741832
Public Class(1 To NumClasses) As ClassData
Attribute Class.VB_VarUserMemId = 1073741833
Public Guild(1 To MaxGuilds) As GuildData
Attribute Guild.VB_VarUserMemId = 1073741834
Public Hall(1 To MaxHalls) As HallData
Attribute Hall.VB_VarUserMemId = 1073741835
Public NPC(1 To MaxNPCs) As NPCData
Attribute NPC.VB_VarUserMemId = 1073741836
Public ItemPrefix(1 To MaxModifications) As PrefixData
Attribute ItemPrefix.VB_VarUserMemId = 1073741838
Public ItemSuffix(1 To MaxModifications) As PrefixData
Attribute ItemSuffix.VB_VarUserMemId = 1073741839
Public Character As CharacterData
Attribute Character.VB_VarUserMemId = 1073741840
Public Macro(0 To 9) As MacroData
Attribute Macro.VB_VarUserMemId = 1073741841
Public World As WorldProperties
Attribute World.VB_VarUserMemId = 1073741842

