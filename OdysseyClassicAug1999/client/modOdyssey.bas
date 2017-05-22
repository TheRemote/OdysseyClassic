Attribute VB_Name = "modOdyssey"
Option Explicit

Public Const TitleString = "The Odyssey Online Classic"

Public Const ClientVer = 6 'A6

#Const DEBUGWINDOWPROC = 0
#Const USEGETPROP = 0
#Const CHEATS = 0

#If DEBUGWINDOWPROC Then
Private m_SCHook As WindowProcHook
#End If

Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

'INI File Related
Declare Function GetPrivateProfileInt Lib "kernel32" Alias "GetPrivateProfileIntA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal nDefault As Long, ByVal lpFileName As String) As Long
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Declare Function MessageBox Lib "user32" Alias "MessageBoxA" (ByVal hWnd As Long, ByVal lpText As String, ByVal lpCaption As String, ByVal wType As Long) As Long

Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long

Declare Function LoadImage Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As String, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
Declare Function RealizePalette Lib "gdi32" (ByVal hdc As Long) As Long
Declare Function CreatePalette Lib "gdi32" (lpLogPalette As LogPalette) As Long
Declare Function SelectPalette Lib "gdi32" (ByVal hdc As Long, ByVal hPalette As Long, ByVal bForceBackground As Long) As Long
Declare Function SetPaletteEntries Lib "gdi32" (ByVal hPalette As Long, ByVal wStartIndex As Long, ByVal wNumEntries As Long, lpPaletteEntries As PaletteEntry) As Long
Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Declare Function GetTickCount Lib "kernel32" () As Long
Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, ByVal nBkMode As Long) As Long
Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Declare Function CreateBitmap Lib "gdi32" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, lpBits As Any) As Long
Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long

'Lighting Effect Related
Public Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long

'System Registry Related
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegQueryValueExString Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
Private Declare Function RegQueryValueExLong Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Long, lpcbData As Long) As Long
Private Declare Function RegQueryValueExNULL Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As Long, lpcbData As Long) As Long

'System Registry Constants
Const REG_SZ As Long = 1
Const REG_DWORD As Long = 4

Const HKEY_CLASSES_ROOT = &H80000000
Const HKEY_CURRENT_USER = &H80000001
Const HKEY_LOCAL_MACHINE = &H80000002
Const HKEY_USERS = &H80000003

Const ERROR_NONE = 0
Const ERROR_BADDB = 1
Const ERROR_BADKEY = 2
Const ERROR_CANTOPEN = 3
Const ERROR_CANTREAD = 4
Const ERROR_CANTWRITE = 5
Const ERROR_OUTOFMEMORY = 6
Const ERROR_INVALID_PARAMETER = 7
Const ERROR_ACCESS_DENIED = 8
Const ERROR_INVALID_PARAMETERS = 87
Const ERROR_NO_MORE_ITEMS = 259

Const KEY_ALL_ACCESS = &H3F
Const REG_OPTION_NON_VOLATILE = 0


'SendMessage Constants
Public Const EM_LINEINDEX = &HBB
Public Const EM_LINELENGTH = &HC1
Public Const EM_LINEFROMCHAR = &HC9

'EM_SCROLL EM_LINELENGTH EM_SETRECTNP EM_GETPASSWORDCHAR EM_LINEFROMCHAR EM_SETHANDLE EM_UNDO EM_REPLACESEL EM_GETWORDBREAKPROC EM_GETTHUMB

'SetBkMode Constants
Public Const TRANSPARENT = 1

'WaitForTerm
Public Const SYNCHRONIZE = &H100000
Public Const INFINITE = &HFFFFFFFF

'BitBlt Constants
Public Const BLACKNESS = &H42
Public Const WHITENESS = &HFF0062
Public Const SRCCOPY = &HCC0020
Public Const SRCPAINT = &HEE0086
Public Const SRCAND = &H8800C6
Public Const NOTSRCCOPY = &H330008
Public Const SRCINVERT = &H660046
Public Const DSTINVERT = &H550009

'DrawText Constants
Public Const DT_TOP = &H0
Public Const DT_LEFT = &H0
Public Const DT_CENTER = &H1
Public Const DT_RIGHT = &H2
Public Const DT_VCENTER = &H4
Public Const DT_BOTTOM = &H8
Public Const DT_WORDBREAK = &H10
Public Const DT_SINGLELINE = &H20
Public Const DT_EXPANDTABS = &H40
Public Const DT_TABSTOP = &H80
Public Const DT_NOCLIP = &H100
Public Const DT_EXTERNALLEADING = &H200
Public Const DT_CALCRECT = &H400
Public Const DT_NOPREFIX = &H800
Public Const DT_INTERNAL = &H1000

'Hook
Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Public Const GWL_WNDPROC = -4
Public lpPrevWndProc As Long
Public gHW As Long

Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Declare Function ReleaseCapture Lib "user32" () As Long
Public Const WM_NCLBUTTONDOWN = &HA1
Public Const HTCAPTION = 2

Public hdcTiles As Long, hbmpTiles As Long, obmpTiles As Long
Public hdcTilesMask As Long, hbmpTilesMask As Long, obmpTilesMask As Long

Public hdcObjects As Long, hbmpObjects As Long, obmpObjects As Long
Public hdcObjectsMask As Long, hbmpObjectsMask As Long, obmpObjectsMask As Long

Public hdcSprites As Long, hbmpSprites As Long, obmpSprites As Long
Public hdcSpritesMask As Long, hbmpSpritesMask As Long, obmpSpritesMask As Long

Public hdcLSprites As Long, hbmpLSprites As Long, obmpLSprites As Long
Public hdcLSpritesMask As Long, hbmpLSpritesMask As Long, obmpLSpritesMask As Long

Public hdcEffects As Long, hbmpEffects As Long, obmpEffects As Long
Public hdcEffectsMask As Long, hbmpEffectsMask As Long, obmpEffectsMask As Long

Public hdcTempMask As Long, hbmpTempMask As Long, obmpTempMask As Long

Public hdcBack(0 To 1) As Long, hbmpBack(0 To 1) As Long, obmpBack(0 To 1) As Long

Public hdcBuffer As Long, hbmpBuffer As Long, obmpBuffer As Long

Public hdcFront As Long, hbmpFront As Long, obmpFront As Long
Public hdcFrontMask As Long, hbmpFrontMask As Long, obmpFrontMask As Long
Public hdcNight As Long, hbmpNight As Long, obmpNight As Long
Public hdcNightMask As Long, hbmpNightMask As Long, obmpNightMask As Long
Public hdcViewport As Long, hdcInv As Long
Public hdcGlow As Long, hdcGlowMask As Long
Public hPalette As Long

Public ClientSocket As Long, SocketData As String

Type Bitmap
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type

Private Type PaletteEntry
        peRed As Byte
        peGreen As Byte
        peBlue As Byte
        peFlags As Byte
End Type

Type LogPalette
    palVersion As Integer
    palNumEntries As Integer
    palPalEntry(0 To 255) As PaletteEntry
End Type

Public Palette(0 To 255) As PaletteEntry
Public Bitmap As Bitmap

Type NPCSaleItemData
    GiveObject As Byte
    GiveValue As Long
    TakeObject As Byte
    TakeValue As Long
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

Type PlayerData
    Name As String
    LastMessage As String
    RepCount As Long
    Map As Long
    Sprite As Byte
    Status As Byte
    X As Long
    Y As Long
    XO As Long
    YO As Long
    D As Byte
    A As Byte
    W As Long
    WalkStep As Long
    Guild As Byte
    Color As Long
    Ignore As Boolean
End Type

Type MacroData
    Text As String
    LineFeed As Boolean
End Type

Type ObjectData
    Name As String
    Type As Byte
    Picture As Byte
End Type

Type MonsterData
    Name As String
    Sprite As Byte
    Large As Boolean
End Type

Type MapStartLocationData
    Map As Integer
    X As Byte
    Y As Byte
End Type

Type MapDoorData
    Att As Byte
    BGTile1 As Integer
    X As Long
    Y As Long
End Type

Type TileData
    Ground As Integer
    BGTile1 As Integer
    BGTile2 As Integer
    FGTile As Integer
    Att As Byte
    AttData(0 To 3) As Byte
End Type

Type MapMonsterData
    Monster As Byte
    X As Long
    Y As Long
    XO As Long
    YO As Long
    D As Byte
    A As Byte
    W As Byte
End Type

Type MapObjectData
    Object As Byte
    X As Byte
    Y As Byte
    XOffset As Byte
    YOffset As Byte
End Type

Type MapMonsterSpawnData
    Monster As Byte
    Rate As Byte
End Type

Type MapData
    Name As String
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
    NPC As Byte
    MIDI As Byte
    Flags As Byte
    Version As Long
End Type

Type InvObjData
    Object As Byte
    Value As Long
    EquippedNum As Byte
End Type

Type GuildData
    Name As String
End Type

Type HallData
    Name As String
End Type

Type NPCData
    Name As String
End Type

Type GuildDeclarationData
    Guild As Byte
    Type As Byte
End Type

Type CharacterData
    Name As String
    Class As Byte
    Gender As Byte
    Sprite As Byte
    
    MaxHP As Byte
    MaxEnergy As Byte
    MaxMana As Byte
    HP As Byte
    Energy As Byte
    Mana As Byte
    
    Strength As Byte
    Agility As Byte
    Endurance As Byte
    Intelligence As Byte
    Experience As Long
    level As Byte
    StatPoints As Integer
    
    Status As Byte
    Access As Byte
    Index As Byte
    Guild As Byte
    GuildRank As Byte
    
    GuildDeclaration(0 To 4) As GuildDeclarationData
    
    desc As String
    Inv(1 To 20) As InvObjData
End Type

Type ClassData
    Name As String
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

Type OptionsData
    MIDI As Boolean
    Wav As Boolean
    Broadcasts As Boolean
    Away As Boolean
    AwayMsg As String
    ForwardUser As String
End Type

'Map Editing
Public MapEdit As Boolean, EditMode As Byte
Public CurTile As Integer, TopY As Long
Public NewAtt As Integer, CurAtt As Integer, CurAttData(0 To 3) As Byte

Public Map As MapData, EditMap As MapData, ClipboardMap As MapData
Public StatusColors(0 To 100) As Long
Public Player(1 To 255) As PlayerData
Public Projectile(1 To 255) As ProjectileData
Public Monster(1 To 255) As MonsterData
Public Object(1 To 255) As ObjectData
Public Class(1 To 16) As ClassData
Public Guild(1 To 255) As GuildData
Public Hall(1 To 255) As HallData
Public NPC(1 To 255) As NPCData
Public SaleItem(0 To 9) As NPCSaleItemData
Public Const NumClasses = 16
Public Character As CharacterData
Public Options As OptionsData
Public Macro(0 To 9) As MacroData

Public CX As Long, CY As Long, CMap As Long 'Your Location
Public CWalkCode As Byte
Public CXO As Long, CYO As Long, CDir As Long, CWalkStep As Long
Public CAttack As Long, CWalk As Long

Public NewAccount As Boolean
Public User As String, Pass As String

Public frmWait_Loaded As Boolean
Public frmMain_Loaded As Boolean, frmMain_Showing As Boolean
Public frmLogin_Loaded As Boolean
Public frmAccount_Loaded As Boolean
Public frmCharacter_Loaded As Boolean
Public frmNewCharacter_Loaded As Boolean
Public frmNewPass_Loaded As Boolean
Public frmMonster_Loaded As Boolean
Public frmObject_Loaded As Boolean
Public frmList_Loaded As Boolean
Public frmMapProperties_Loaded As Boolean
Public frmGuilds_Loaded As Boolean
Public frmGuild_Loaded As Boolean
Public frmNPC_Loaded As Boolean
Public frmMacros_Loaded As Boolean
Public frmOptions_Loaded As Boolean
Public frmNewGuild_Loaded As Boolean
Public frmBan_Loaded As Boolean
Public frmHall_Loaded As Boolean

Public blnEnd As Boolean, blnPlaying As Boolean

Public keyUp As Boolean, keyDown As Boolean
Public keyLeft As Boolean, keyRight As Boolean
Public keyCtrl As Boolean, keyShift As Boolean

'Misc Variables
Public MapCacheFile As String
Public AttackTimer As Long
Public ChatString As String
Public blnNight As Boolean
Public Const ServerIP As String = "76.178.167.74"
Public CurInvObj As Long
Public Freeze As Boolean
Public NextTransition As Long
Public CurrentMIDI As Long
Public SecondCounter As Long, FrameCounter As Long, FrameTimer As Long, FrameRate As Long
Public CurFrame As Long
Public TempVar1 As Long, TempVar2 As Long, TempVar3 As Long, TempVar4 As Long, TempVar5 As Long
Public ChatScrollBack As Long
Public RequestedMap As Boolean
Public Paletted As Boolean
Public LastProjectile As Long
Public ComputerID As String

Public Section(1 To 30) As String, Suffix As String

Public AutoAttack As Boolean, TargetMonster As Integer 'Temporary
Public AutoAttackSpeed As Integer, AutoAttackWalk As Integer
Public DisableScreen As Boolean

'Info Text
Public InfoText(0 To 3) As String
Public InfoTextTimer As Long

Public MapData As String * 1927

Public ServerID As String
Function BadName(St As String) As Boolean
    Dim A As Long, B As Long
    For A = 1 To Len(St)
        B = Asc(Mid$(St, A, 1))
        If B < 32 Or B > 126 Then
            BadName = True
        End If
    Next A
End Function

Function GetFilename(St As String) As String
    If Exists(ServerID + "\" + St) Then
        GetFilename = ServerID + "\" + St
    Else
        GetFilename = St
    End If
End Function

Sub SendRaw(ByVal St As String)
    If SendData(ClientSocket, St) = SOCKET_ERROR Then
        CloseClientSocket 0
    End If
End Sub

Sub WaitForTerm(pid As Long)
    Dim phnd As Long
    phnd = OpenProcess(SYNCHRONIZE, 0, pid)
    If phnd <> 0 Then
        Call WaitForSingleObject(phnd, INFINITE)
        Call CloseHandle(phnd)
    End If
End Sub

Sub CreateMapCache()
    Dim St1 As String * 1927, A As Long
    St1 = String$(1927, 0)
    Open MapCacheFile For Random As #1 Len = 1927
    For A = 1 To 2000
        Put #1, , St1
    Next A
    Close #1
End Sub

Sub DelItem(lpAppName As String, lpKeyName As String)
    WritePrivateProfileString lpAppName, lpKeyName, 0&, App.Path + "\odyssey.ini"
End Sub
Sub LoadMacros()
    Dim A As Long
    For A = 0 To 9
        With Macro(A)
            .Text = ReadString("Macros", "Text" + CStr(A + 1))
            If Len(.Text) > 255 Then .Text = Left$(.Text, 255)
            If ReadInt("Macros", "LineFeed" + CStr(A + 1)) = 0 Then
                .LineFeed = False
            Else
                .LineFeed = True
            End If
        End With
    Next A
End Sub
Sub MonsterDied(Index As Long)
    Dim A As Long
    
    For A = 1 To 100
        With Projectile(A)
            If .Sprite > 0 Then
                If .TargetType = pttMonster Then
                    If .TargetNum = Index Then
                        DestroyEffect (A)
                    End If
                End If
            End If
        End With
    Next A
End Sub
Sub PlayerLeftMap(Index As Long)
    Dim A As Long
    
    For A = 1 To 100
        With Projectile(A)
            If .Sprite > 0 Then
                If .TargetType = pttCharacter Then
                    If .TargetNum = Index Then
                        DestroyEffect (Index)
                    End If
                End If
            End If
        End With
    Next A
    
    Player(Index).Map = 0
End Sub
Function QuadChar(Num As Long) As String
    QuadChar = Chr$(Int(Num / 16777216) Mod 256) + Chr$(Int(Num / 65536) Mod 256) + Chr$(Int(Num / 256) Mod 256) + Chr$(Num Mod 256)
End Function

Sub CreateClassData()
    With Class(1) 'Mage
        .Name = "Mage"
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
        .Name = "Dark Mage"
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
        .Name = "Knight"
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
        .Name = "Paladin"
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
        .Name = "Barbarian"
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
        .Name = "Cleric"
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
        .Name = "Jester"
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
        .Name = "Monk"
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
        .Name = "Necromancer"
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
        .Name = "Bushido"
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
        .Name = "Crusader"
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
        .Name = "Bard"
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
        .Name = "Ninja"
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
        .Name = "Samurai"
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
        .Name = "Thief"
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
        .Name = "Ranger"
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
Sub DrawCurInvObj()
    frmMain.picObj.Cls
    If CurInvObj > 0 Then
        If Character.Inv(CurInvObj).Object > 0 Then
            BitBlt frmMain.picObj.hdc, 6, 6, 32, 32, hdcObjects, 0, (CLng(Object(Character.Inv(CurInvObj).Object).Picture) - 1) * 32, SRCCOPY
            TextOut frmMain.picObj.hdc, 50, 6, Object(Character.Inv(CurInvObj).Object).Name, Len(Object(Character.Inv(CurInvObj).Object).Name)
            If Object(Character.Inv(CurInvObj).Object).Type = 6 Then
                'Money
                TextOut frmMain.picObj.hdc, 50, 30, "[" + CStr(Character.Inv(CurInvObj).Value) + "]", Len(CStr(Character.Inv(CurInvObj).Value)) + 2
            End If
        End If
    End If
End Sub

Sub DrawInfoText()
    Dim A As Long
    If InfoTextTimer > 0 Then
        If GetTickCount - InfoTextTimer <= 10000 Then
            For A = 0 To 3
                If InfoText(A) <> "" Then
                    SetTextColor hdcBuffer, QBColor(0)
                    TextOut hdcBuffer, 5, 332 + 12 * A, InfoText(A), Len(InfoText(A))
                    SetTextColor hdcBuffer, QBColor(12)
                    TextOut hdcBuffer, 3, 330 + 12 * A, InfoText(A), Len(InfoText(A))
                End If
            Next A
        Else
            For A = 0 To 3
                InfoText(A) = ""
            Next A
            InfoTextTimer = 0
        End If
    End If
End Sub
Sub DrawTrainBars()
    Dim X1 As Long, X2 As Long, St As String
    frmMain.lblStatPoints = "Free Stat Points: " + CStr(TempVar5)
    With frmMain.picTrainStrength
        X1 = Int((CSng(Character.Strength) / 30!) * CSng(.ScaleWidth))
        X2 = Int((CSng(Character.Strength + TempVar1) / 30!) * CSng(.ScaleWidth))
        If X1 > 0 Then frmMain.picTrainStrength.Line (0, 0)-(X1, .ScaleHeight), QBColor(7), BF
        If X2 > X1 Then frmMain.picTrainStrength.Line (X1, 0)-(X2, .ScaleHeight), QBColor(15), BF
        If X2 < .ScaleWidth Then frmMain.picTrainStrength.Line (X2, 0)-(.ScaleWidth, .ScaleHeight), 0, BF
        If TempVar1 > 0 Then
            St = "Strength +" + CStr(TempVar1)
        Else
            St = "Strength"
        End If
        TextOut .hdc, 5, (.ScaleHeight - .TextHeight(St)) / 2, St, Len(St)
        .Refresh
    End With
    With frmMain.picTrainAgility
        X1 = Int((CSng(Character.Agility) / 30!) * CSng(.ScaleWidth))
        X2 = Int((CSng(Character.Agility + TempVar2) / 30!) * CSng(.ScaleWidth))
        If X1 > 0 Then frmMain.picTrainAgility.Line (0, 0)-(X1, .ScaleHeight), QBColor(7), BF
        If X2 > X1 Then frmMain.picTrainAgility.Line (X1, 0)-(X2, .ScaleHeight), QBColor(15), BF
        If X2 < .ScaleWidth Then frmMain.picTrainAgility.Line (X2, 0)-(.ScaleWidth, .ScaleHeight), 0, BF
        If TempVar2 > 0 Then
            St = "Agility +" + CStr(TempVar2)
        Else
            St = "Agility"
        End If
        TextOut .hdc, 5, (.ScaleHeight - .TextHeight(St)) / 2, St, Len(St)
        .Refresh
    End With
    With frmMain.picTrainEndurance
        X1 = Int((CSng(Character.Endurance) / 30!) * CSng(.ScaleWidth))
        X2 = Int((CSng(Character.Endurance + TempVar3) / 30!) * CSng(.ScaleWidth))
        If X1 > 0 Then frmMain.picTrainEndurance.Line (0, 0)-(X1, .ScaleHeight), QBColor(7), BF
        If X2 > X1 Then frmMain.picTrainEndurance.Line (X1, 0)-(X2, .ScaleHeight), QBColor(15), BF
        If X2 < .ScaleWidth Then frmMain.picTrainEndurance.Line (X2, 0)-(.ScaleWidth, .ScaleHeight), 0, BF
        If TempVar3 > 0 Then
            St = "Endurance +" + CStr(TempVar3)
        Else
            St = "Endurance"
        End If
        TextOut .hdc, 5, (.ScaleHeight - .TextHeight(St)) / 2, St, Len(St)
        .Refresh
    End With
    With frmMain.picTrainIntelligence
        X1 = Int((CSng(Character.Intelligence) / 30!) * CSng(.ScaleWidth))
        X2 = Int((CSng(Character.Intelligence + TempVar4) / 30!) * CSng(.ScaleWidth))
        If X1 > 0 Then frmMain.picTrainIntelligence.Line (0, 0)-(X1, .ScaleHeight), QBColor(7), BF
        If X2 > X1 Then frmMain.picTrainIntelligence.Line (X1, 0)-(X2, .ScaleHeight), QBColor(15), BF
        If X2 < .ScaleWidth Then frmMain.picTrainIntelligence.Line (X2, 0)-(.ScaleWidth, .ScaleHeight), 0, BF
        If TempVar4 > 0 Then
            St = "Intelligence +" + CStr(TempVar4)
        Else
            St = "Intelligence"
        End If
        TextOut .hdc, 5, (.ScaleHeight - .TextHeight(St)) / 2, St, Len(St)
        .Refresh
    End With
End Sub
Function mmSend(ByVal sCommand As String) As String
    Dim iLen As Long, msReturn As String, Success As Long
    msReturn = Space$(255)
    iLen = Len(msReturn)
    Success = mciSendString(sCommand, msReturn, iLen, 0)
    If Success = 0 Then
        msReturn = Trim$(msReturn)
        msReturn = Left$(msReturn, Len(msReturn) - 1)
    Else
        msReturn = ""
    End If
    mmSend = msReturn
End Function
Sub MoveToTile()
    SendSocket Chr$(7) + Chr$(CWalkCode) + Chr$(CX) + Chr$(CY) + Chr$(CDir) + Chr$(CWalkStep)
    If Map.Tile(CX, CY).Att = 2 Then
        Freeze = True
        NextTransition = 5
    End If
End Sub

Sub PrintInfoText(St As String)
    Dim A As Long
    InfoTextTimer = GetTickCount
    For A = 0 To 2
        InfoText(A) = InfoText(A + 1)
    Next A
    InfoText(3) = St
End Sub
Sub SaveOptions()
    With Options
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
        If .Away = True Then
            WriteString "Options", "Away", "1"
        Else
            WriteString "Options", "Away", "0"
        End If
    End With
End Sub
Sub LoadOptions()
    With Options
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
            If ReadInt("Options", "Away") = 1 Then
                .Away = True
                .AwayMsg = "Sorry, I'm not here to answer you. Will be back soon. :)"
            Else
                .Away = False
            End If
        Else
            .MIDI = True
            .Wav = True
            .Broadcasts = True
            .Away = True
            SaveOptions
        End If
    End With
End Sub
Sub ShowMap()
    frmMain.lblLocation = "[" + Map.Name + "]"
    If ExamineBit(Map.Flags, 0) = True Then
        frmMain.lblLocation.ForeColor = QBColor(11)
    ElseIf ExamineBit(Map.Flags, 6) = True Then
        frmMain.lblLocation.ForeColor = QBColor(12)
    Else
        frmMain.lblLocation.ForeColor = QBColor(15)
    End If
    If Map.MIDI > 0 Then
        PlayMidi CLng(Map.MIDI)
    Else
        StopMidi
    End If
    DrawMap
    If frmWait_Loaded = True Then
        Unload frmWait
    End If
    If frmMain_Showing = False Then
        frmMain.Show
        frmMain_Showing = True
    End If
    Transition
    Freeze = False
End Sub
Sub StopMidi()
    If CurrentMIDI > 0 Then
        mciSendString "stop midi wait", 0&, 0, 0
        mciSendString "close midi wait", 0&, 0, 0
        CurrentMIDI = 0
    End If
End Sub

Sub PlayMidi(number As Long)
    If Options.MIDI = True Then
        If CurrentMIDI = number Then Exit Sub
        If CurrentMIDI > 0 Then
            StopMidi
        End If
        CurrentMIDI = number
        mciSendString "open " + GetFilename("mus" + CStr(number) + ".mid") + " type sequencer alias midi wait", 0&, 0, 0
        mciSendString "play midi", 0&, 0, 0
    End If
End Sub
Sub ClearBit(bytByte As Byte, Bit As Byte)
      bytByte = bytByte And Not (2 ^ Bit)
End Sub

Sub PlayWav(number As Long)
    If Options.Wav = True Then
        sndPlaySound GetFilename("sound" + CStr(number) + ".wav"), 1
    End If
End Sub


Sub RedrawTile()
    BitBlt frmMain.picTile.hdc, 0, 0, 32, 32, 0, 0, 0, BLACKNESS
    If EditMode < 4 Then
        If CurTile > 0 Then
            BitBlt frmMain.picTile.hdc, 0, 0, 32, 32, hdcTiles, ((CurTile - 1) Mod 7) * 32, Int((CurTile - 1) / 7) * 32, SRCCOPY
        End If
    Else
        If CurAtt > 0 Then
            BitBlt frmMain.picTile.hdc, 0, 0, 32, 32, frmMain.picAtts.hdc, ((CurAtt - 1) Mod 7) * 32, Int((CurAtt - 1) / 7) * 32, SRCCOPY
        End If
    End If
    frmMain.picTile.Refresh
End Sub
Sub RedrawTiles()
    BitBlt frmMain.picTiles.hdc, 0, 0, 224, 192, 0, 0, 0, BLACKNESS
    If EditMode < 4 Then
        BitBlt frmMain.picTiles.hdc, 0, 0, 224, 192, hdcTiles, 0, TopY, SRCCOPY
    Else
        BitBlt frmMain.picTiles.hdc, 0, 0, 224, 192, frmMain.picAtts.hdc, 0, 0, SRCCOPY
    End If
    frmMain.picTiles.Refresh
End Sub
Sub SetBit(bytByte As Byte, Bit As Byte)
    bytByte = bytByte Or (2 ^ Bit)
End Sub

Function ExamineBit(bytByte As Byte, Bit As Byte) As Byte
    ExamineBit = ((bytByte And (2 ^ Bit)) > 0)
End Function

Sub CloseMapEdit()
    MapEdit = False
    frmMain.picInfobar.Visible = True
    frmMain.picMapEdit.Visible = False
    DrawMap
End Sub

Sub DrawChatString()
    Dim R As RECT
    
    If ChatString <> "" Then
        With R
            .Left = 5
            .Top = 5
            .Right = 379
            .Bottom = 50
        End With
        SetTextColor hdcBuffer, QBColor(15)
        DrawText hdcBuffer, ChatString, Len(ChatString), R, DT_CENTER Or DT_NOCLIP Or DT_WORDBREAK
    End If
End Sub

Sub GetSections(ByVal St As String, NumSections)
    Dim A As Integer, W As Integer, Q As Boolean
    Dim CurChar As String * 1, LastChar As String * 1
    Erase Section
    Suffix = ""
    If Len(St) = 0 Then Exit Sub
    
    W = 1
    For A = 1 To Len(St)
        CurChar = Mid$(St, A, 1)
        Select Case Asc(CurChar)
            Case 32
                If Q = False Then
                    If Not LastChar = Chr$(32) Then W = W + 1
                    If W > NumSections Then Exit For
                Else
                    Section(W) = Section(W) + CurChar
                End If
            Case 34
                If Q = False Then Q = True Else Q = False
            Case Else
                Section(W) = Section(W) + CurChar
        End Select
        LastChar = CurChar
    Next A
    If A < Len(St) Then
        Suffix = Mid$(St, A + 1)
    Else
        Suffix = ""
    End If
End Sub
Sub GetSections2(St)
    Dim A As Long, B As Long, C As Long
    B = 1
    Erase Section
    For A = 1 To 30
        C = InStr(B, St, Chr$(0))
        If C - B = 0 Then
            Section(A) = ""
        ElseIf C <> 0 Then
            Section(A) = Mid$(St, B, C - B)
        Else
            Section(A) = Mid$(St, B, Len(St) - B + 1)
            Exit For
        End If
        B = C + 1
    Next A
End Sub
Function ClipString(St As String) As String
    Dim A As Long
    For A = Len(St) To 1 Step -1
        If Mid$(St, A, 1) <> Chr$(32) Then
            ClipString = Mid$(St, 1, A)
            Exit Function
        End If
    Next A
End Function

Sub CopyMap(DestMap As MapData, SourceMap As MapData)
    Dim A As Long, X As Long, Y As Long
    
    With DestMap
        .Name = SourceMap.Name
        .MIDI = SourceMap.MIDI
        .NPC = SourceMap.NPC
        .ExitUp = SourceMap.ExitUp
        .ExitDown = SourceMap.ExitDown
        .ExitLeft = SourceMap.ExitLeft
        .ExitRight = SourceMap.ExitRight
        .BootLocation.Map = SourceMap.BootLocation.Map
        .BootLocation.X = SourceMap.BootLocation.X
        .BootLocation.Y = SourceMap.BootLocation.Y
        .Flags = SourceMap.Flags
        For A = 0 To 2
            .MonsterSpawn(A).Monster = SourceMap.MonsterSpawn(A).Monster
            .MonsterSpawn(A).Rate = SourceMap.MonsterSpawn(A).Rate
        Next A
        For Y = 0 To 11
            For X = 0 To 11
                With .Tile(X, Y)
                    .Ground = SourceMap.Tile(X, Y).Ground
                    .BGTile1 = SourceMap.Tile(X, Y).BGTile1
                    .BGTile2 = SourceMap.Tile(X, Y).BGTile2
                    .FGTile = SourceMap.Tile(X, Y).FGTile
                    .Att = SourceMap.Tile(X, Y).Att
                    .AttData(0) = SourceMap.Tile(X, Y).AttData(0)
                    .AttData(1) = SourceMap.Tile(X, Y).AttData(1)
                    .AttData(2) = SourceMap.Tile(X, Y).AttData(2)
                    .AttData(3) = SourceMap.Tile(X, Y).AttData(3)
                End With
            Next X
        Next Y
        For A = 0 To 9
            With SourceMap.Door(A)
                If .Att > 0 Then
                    DestMap.Tile(.X, .Y).BGTile1 = .BGTile1
                    DestMap.Tile(.X, .Y).Att = .Att
                End If
            End With
            .Door(A).Att = 0
            .Door(A).BGTile1 = 0
        Next A
    End With
End Sub
Sub LoadMap(MapData As String)
    Dim A As Long, X As Long, Y As Long
    If Len(MapData) = 1927 Then
        With Map
            .Name = ClipString$(Mid$(MapData, 1, 30))
            .Version = Asc(Mid$(MapData, 31, 1)) * 16777216 + Asc(Mid$(MapData, 32, 1)) * 65536 + Asc(Mid$(MapData, 33, 1)) * 256& + Asc(Mid$(MapData, 34, 1))
            .NPC = Asc(Mid$(MapData, 35, 1))
            .MIDI = Asc(Mid$(MapData, 36, 1))
            .ExitUp = Asc(Mid$(MapData, 37, 1)) * 256 + Asc(Mid$(MapData, 38, 1))
            .ExitDown = Asc(Mid$(MapData, 39, 1)) * 256 + Asc(Mid$(MapData, 40, 1))
            .ExitLeft = Asc(Mid$(MapData, 41, 1)) * 256 + Asc(Mid$(MapData, 42, 1))
            .ExitRight = Asc(Mid$(MapData, 43, 1)) * 256 + Asc(Mid$(MapData, 44, 1))
            .BootLocation.Map = Asc(Mid$(MapData, 45, 1)) * 256 + Asc(Mid$(MapData, 46, 1))
            .BootLocation.X = Asc(Mid$(MapData, 47, 1))
            .BootLocation.Y = Asc(Mid$(MapData, 48, 1))
            .Flags = Asc(Mid$(MapData, 49, 1))
            For A = 0 To 2 '50 - 55
                .MonsterSpawn(A).Monster = Asc(Mid$(MapData, 50 + A * 2))
                .MonsterSpawn(A).Rate = Asc(Mid$(MapData, 51 + A * 2))
            Next A
            '56
            For Y = 0 To 11
                For X = 0 To 11
                    With .Tile(X, Y)
                        A = 56 + Y * 156 + X * 13
                        .Ground = Asc(Mid$(MapData, A, 1)) * 256 + Asc(Mid$(MapData, A + 1, 1))
                        .BGTile1 = Asc(Mid$(MapData, A + 2, 1)) * 256 + Asc(Mid$(MapData, A + 3, 1))
                        .BGTile2 = Asc(Mid$(MapData, A + 4, 1)) * 256 + Asc(Mid$(MapData, A + 5, 1))
                        .FGTile = Asc(Mid$(MapData, A + 6, 1)) * 256 + Asc(Mid$(MapData, A + 7, 1))
                        .Att = Asc(Mid$(MapData, A + 8, 1))
                        .AttData(0) = Asc(Mid$(MapData, A + 9, 1))
                        .AttData(1) = Asc(Mid$(MapData, A + 10, 1))
                        .AttData(2) = Asc(Mid$(MapData, A + 11, 1))
                        .AttData(3) = Asc(Mid$(MapData, A + 12, 1))
                    End With
                Next X
            Next Y
        End With
    End If
End Sub
Sub Draw3dText(DC As Long, TargetRect As RECT, St As String, lngColor As Long, Height As Integer)
    Dim ShadowRect As RECT
    With ShadowRect
        .Top = TargetRect.Top + Height
        .Left = TargetRect.Left + Height
        .Bottom = TargetRect.Bottom + Height
        .Right = TargetRect.Right + Height
    End With
    SetTextColor DC, RGB(10, 10, 10)
    DrawText DC, St, Len(St), ShadowRect, DT_CENTER Or DT_NOCLIP Or DT_WORDBREAK
    SetTextColor DC, lngColor
    DrawText DC, St, Len(St), TargetRect, DT_CENTER Or DT_NOCLIP Or DT_WORDBREAK
End Sub
Sub DrawHP()
    Dim Percent As Single, St As String
    
    If Character.MaxHP > 0 Then
        Percent = Character.HP / Character.MaxHP
    Else
        Percent = 0
    End If
    St = CStr(Int(Percent * 100)) + "% " + CStr(Character.HP) + "/" + CStr(Character.MaxHP)
    With frmMain.picStats
        frmMain.picStats.Line (44, 9)-(44 + 167 * Percent, 15), QBColor(15), BF
        frmMain.picStats.Line (44 + 167 * Percent, 9)-(211, 15), QBColor(8), BF
        TextOut .hdc, 44 + (167 - .TextWidth(St)) / 2, 9 + (7 - .TextHeight(St)) / 2, St, Len(St)
        .Refresh
    End With
End Sub
Sub DrawEnergy()
    Dim Percent As Single, St As String
    
    If Character.MaxEnergy > 0 Then
        Percent = Character.Energy / Character.MaxEnergy
    Else
        Percent = 0
    End If
    St = CStr(Int(Percent * 100)) + "% " + CStr(Character.Energy) + "/" + CStr(Character.MaxEnergy)
    With frmMain.picStats
        frmMain.picStats.Line (44, 25)-(44 + 167 * Percent, 31), QBColor(15), BF
        frmMain.picStats.Line (44 + 167 * Percent, 25)-(211, 31), QBColor(8), BF
        TextOut .hdc, 44 + (167 - .TextWidth(St)) / 2, 25 + (7 - .TextHeight(St)) / 2, St, Len(St)
        .Refresh
    End With
End Sub
Sub DrawInvObject(InvNum As Long)
    Dim A As Long, X As Long, Y As Long
    
    X = 6 + 44 * ((InvNum - 1) Mod 5)
    Y = 6 + 44 * Int((InvNum - 1) / 5)
    
    With Character.Inv(InvNum)
        If .Object > 0 Then
            A = Object(.Object).Picture
            If A > 0 Then
                If CurInvObj = InvNum Then
                    BitBlt hdcInv, X - 1, Y - 1, 34, 34, 0, 0, 0, WHITENESS
                    If .EquippedNum > 0 Then
                        TransparentBlt hdcInv, X - 1, Y - 1, 34, 34, hdcGlow, 0, 0, hdcGlowMask
                    End If
                    TransparentBlt hdcInv, X, Y, 32, 32, hdcObjects, 0, CLng(A - 1) * 32, hdcObjectsMask
                Else
                    If .EquippedNum > 0 Then
                        BitBlt hdcInv, X - 1, Y - 1, 34, 34, hdcGlow, 0, 0, SRCCOPY
                        TransparentBlt hdcInv, X, Y, 32, 32, hdcObjects, 0, CLng(A - 1) * 32, hdcObjectsMask
                    Else
                        BitBlt hdcInv, X - 1, Y - 1, 34, 34, 0, 0, 0, BLACKNESS
                        BitBlt hdcInv, X, Y, 32, 32, hdcObjects, 0, CLng(A - 1) * 32, SRCCOPY
                    End If
                End If
            Else
                If CurInvObj = InvNum Then
                    BitBlt hdcInv, X - 1, Y - 1, 34, 34, 0, 0, 0, WHITENESS
                Else
                    BitBlt hdcInv, X - 1, Y - 1, 34, 34, 0, 0, 0, BLACKNESS
                End If
            End If
        Else
            If CurInvObj = InvNum Then
                BitBlt hdcInv, X - 1, Y - 1, 34, 34, 0, 0, 0, WHITENESS
            Else
                BitBlt hdcInv, X - 1, Y - 1, 34, 34, 0, 0, 0, BLACKNESS
            End If
        End If
    End With
    
    frmMain.picInv.Refresh
    
    If InvNum = CurInvObj Then DrawCurInvObj
End Sub
Sub DrawMana()
    Dim Percent As Single, St As String
    
    If Character.MaxMana > 0 Then
        Percent = Character.Mana / Character.MaxMana
    Else
        Percent = 0
    End If
    St = CStr(Int(Percent * 100)) + "% " + CStr(Character.Mana) + "/" + CStr(Character.MaxMana)
    With frmMain.picStats
        frmMain.picStats.Line (44, 41)-(44 + 167 * Percent, 47), QBColor(15), BF
        frmMain.picStats.Line (44 + 167 * Percent, 41)-(211, 47), QBColor(8), BF
        TextOut .hdc, 44 + (167 - .TextWidth(St)) / 2, 41 + (7 - .TextHeight(St)) / 2, St, Len(St)
        .Refresh
    End With
End Sub

Function IsVacant(X As Long, Y As Long) As Boolean
    Dim A As Long
    Select Case Map.Tile(X, Y).Att
        Case 1, 3 'Wall / Key Door
            Exit Function
        Case 2 'Warp
            If AutoAttack = True Then Exit Function
    End Select
    For A = 0 To 5
        With Map.Monster(A)
            If .Monster > 0 And .X = X And .Y = Y Then
                Exit Function
            End If
        End With
    Next A
    For A = 1 To 255
        With Player(A)
            If .Map = CMap And .X = X And .Y = Y Then
                Exit Function
            End If
        End With
    Next A
    IsVacant = True
End Function
Sub OpenMapEdit()
    MapEdit = True
    CopyMap EditMap, Map
    frmMain.picMapEdit.Visible = True
    frmMain.picInfobar.Visible = False
    RedrawTiles
    RedrawTile
    DrawMap
End Sub
Sub PrepSourceDC(DC As Long)
    SetBkColor DC, 0
End Sub
Sub PrepTargetDC(DC As Long)
    SetBkColor DC, RGB(255, 255, 255)
    SetTextColor DC, 0
End Sub
Sub DrawMap()
    Dim A As Long, B As Long, X As Long, Y As Long
    BitBlt hdcBack(0), 0, 0, 384, 384, 0, 0, 0, BLACKNESS
    BitBlt hdcBack(1), 0, 0, 384, 384, 0, 0, 0, BLACKNESS
    BitBlt hdcFront, 0, 0, 384, 384, 0, 0, 0, BLACKNESS
    BitBlt hdcFrontMask, 0, 0, 384, 384, 0, 0, 0, WHITENESS
    
    If MapEdit = False Then
        For X = 0 To 11
            For Y = 0 To 11
                With Map.Tile(X, Y)
                    If .Ground > 0 Then
                        BitBlt hdcBack(0), X * 32, Y * 32, 32, 32, hdcTiles, ((.Ground - 1) Mod 7) * 32, Int((.Ground - 1) / 7) * 32, SRCCOPY
                        BitBlt hdcBack(1), X * 32, Y * 32, 32, 32, hdcTiles, ((.Ground - 1) Mod 7) * 32, Int((.Ground - 1) / 7) * 32, SRCCOPY
                    End If
                    If .BGTile1 > 0 Then
                        TransparentBlt hdcBack(0), X * 32, Y * 32, 32, 32, hdcTiles, ((.BGTile1 - 1) Mod 7) * 32, Int((.BGTile1 - 1) / 7) * 32, hdcTilesMask
                    End If
                    If .BGTile2 > 0 Then
                        TransparentBlt hdcBack(1), X * 32, Y * 32, 32, 32, hdcTiles, ((.BGTile2 - 1) Mod 7) * 32, Int((.BGTile2 - 1) / 7) * 32, hdcTilesMask
                    ElseIf .BGTile1 > 0 Then
                        TransparentBlt hdcBack(1), X * 32, Y * 32, 32, 32, hdcTiles, ((.BGTile1 - 1) Mod 7) * 32, Int((.BGTile1 - 1) / 7) * 32, hdcTilesMask
                    End If
                    If .FGTile > 0 Then
                        BitBlt hdcFront, X * 32, Y * 32, 32, 32, hdcTiles, ((.FGTile - 1) Mod 7) * 32, Int((.FGTile - 1) / 7) * 32, SRCCOPY
                        BitBlt hdcFrontMask, X * 32, Y * 32, 32, 32, hdcTilesMask, ((.FGTile - 1) Mod 7) * 32, Int((.FGTile - 1) / 7) * 32, SRCCOPY
                    End If
                End With
            Next Y
        Next X
    Else
        For X = 0 To 11
            For Y = 0 To 11
                With EditMap.Tile(X, Y)
                    If .Ground > 0 Then
                        BitBlt hdcBack(0), X * 32, Y * 32, 32, 32, hdcTiles, ((.Ground - 1) Mod 7) * 32, Int((.Ground - 1) / 7) * 32, SRCCOPY
                        BitBlt hdcBack(1), X * 32, Y * 32, 32, 32, hdcTiles, ((.Ground - 1) Mod 7) * 32, Int((.Ground - 1) / 7) * 32, SRCCOPY
                    End If
                    If .BGTile1 > 0 Then
                        TransparentBlt hdcBack(0), X * 32, Y * 32, 32, 32, hdcTiles, ((.BGTile1 - 1) Mod 7) * 32, Int((.BGTile1 - 1) / 7) * 32, hdcTilesMask
                    End If
                    If .BGTile2 > 0 Then
                        TransparentBlt hdcBack(1), X * 32, Y * 32, 32, 32, hdcTiles, ((.BGTile2 - 1) Mod 7) * 32, Int((.BGTile2 - 1) / 7) * 32, hdcTilesMask
                    ElseIf .BGTile1 > 0 Then
                        TransparentBlt hdcBack(1), X * 32, Y * 32, 32, 32, hdcTiles, ((.BGTile1 - 1) Mod 7) * 32, Int((.BGTile1 - 1) / 7) * 32, hdcTilesMask
                    End If
                    If .FGTile > 0 Then
                        BitBlt hdcFront, X * 32, Y * 32, 32, 32, hdcTiles, ((.FGTile - 1) Mod 7) * 32, Int((.FGTile - 1) / 7) * 32, SRCCOPY
                        BitBlt hdcFrontMask, X * 32, Y * 32, 32, 32, hdcTilesMask, ((.FGTile - 1) Mod 7) * 32, Int((.FGTile - 1) / 7) * 32, SRCCOPY
                    End If
                    If .Att > 0 Then
                        BitBlt hdcFront, X * 32 + 8, Y * 32 + 8, 16, 16, frmMain.picAtts.hdc, ((.Att - 1) Mod 7) * 32 + 8, Int((.Att - 1) / 7) * 32 + 8, SRCCOPY
                        BitBlt hdcFrontMask, X * 32 + 8, Y * 32 + 8, 16, 16, 0, 0, 0, BLACKNESS
                    End If
                End With
            Next Y
        Next X
    End If
    For A = 0 To 49
        With Map.Object(A)
            If .Object > 0 Then
                B = Object(.Object).Picture
                If B > 0 Then
                    TransparentBlt hdcBack(0), .X * 32, .Y * 32, 32, 32, hdcObjects, 0, (B - 1) * 32, hdcObjectsMask
                    TransparentBlt hdcBack(1), .X * 32, .Y * 32, 32, 32, hdcObjects, 0, (B - 1) * 32, hdcObjectsMask
                    'StretchBlt hdcBack(0), .X * 32 + .XOffset, .Y * 32 + .YOffset, 16, 16, hdcObjectsMask, 0, (B - 1) * 32, 32, 32, SRCAND
                    'StretchBlt hdcBack(0), .X * 32 + .XOffset, .Y * 32 + .YOffset, 16, 16, hdcObjects, 0, (B - 1) * 32, 32, 32, SRCPAINT
                    'StretchBlt hdcBack(1), .X * 32 + .XOffset, .Y * 32 + .YOffset, 16, 16, hdcObjectsMask, 0, (B - 1) * 32, 32, 32, SRCAND
                    'StretchBlt hdcBack(1), .X * 32 + .XOffset, .Y * 32 + .YOffset, 16, 16, hdcObjects, 0, (B - 1) * 32, 32, 32, SRCPAINT
                End If
            End If
        End With
    Next A
    
    If MapEdit = False Then
        If (ExamineBit(Map.Flags, 1) = False And blnNight = True) Or (ExamineBit(Map.Flags, 1) = True And ExamineBit(Map.Flags, 2) = True) Then
            TransparentBlt hdcFront, 0, 0, 384, 384, hdcNight, 0, 0, hdcNightMask
            TransparentBlt hdcFrontMask, 0, 0, 384, 384, hdcNight, 0, 0, hdcNightMask
        End If
    End If
End Sub
Function ReadInt(lpAppName, lpKeyName$) As Integer
    ReadInt = GetPrivateProfileInt&(lpAppName, lpKeyName$, 0, App.Path + "\odyssey.ini")
End Function

Function ReadString(lpAppName, lpKeyName As String) As String
    Dim lpReturnedString As String, Valid As Long
    lpReturnedString = Space$(256)
    Valid = GetPrivateProfileString&(lpAppName, lpKeyName, "", lpReturnedString, 256, App.Path + "\odyssey.ini")
    ReadString = Left$(lpReturnedString, Valid)
End Function
Sub RedrawMapTile(X As Long, Y As Long)
    Dim A As Long, B As Long
    If X >= 0 And X <= 11 And Y >= 0 And Y <= 11 Then
        If MapEdit = False Then
            With Map.Tile(X, Y)
                If .Ground > 0 Then
                    BitBlt hdcBack(0), X * 32, Y * 32, 32, 32, hdcTiles, ((.Ground - 1) Mod 7) * 32, Int((.Ground - 1) / 7) * 32, SRCCOPY
                    BitBlt hdcBack(1), X * 32, Y * 32, 32, 32, hdcTiles, ((.Ground - 1) Mod 7) * 32, Int((.Ground - 1) / 7) * 32, SRCCOPY
                Else
                    BitBlt hdcBack(0), X * 32, Y * 32, 32, 32, 0, 0, 0, BLACKNESS
                    BitBlt hdcBack(1), X * 32, Y * 32, 32, 32, 0, 0, 0, BLACKNESS
                End If
                If .BGTile1 > 0 Then
                    TransparentBlt hdcBack(0), X * 32, Y * 32, 32, 32, hdcTiles, ((.BGTile1 - 1) Mod 7) * 32, Int((.BGTile1 - 1) / 7) * 32, hdcTilesMask
                End If
                If .BGTile2 > 0 Then
                    TransparentBlt hdcBack(1), X * 32, Y * 32, 32, 32, hdcTiles, ((.BGTile2 - 1) Mod 7) * 32, Int((.BGTile2 - 1) / 7) * 32, hdcTilesMask
                ElseIf .BGTile1 > 0 Then
                    TransparentBlt hdcBack(1), X * 32, Y * 32, 32, 32, hdcTiles, ((.BGTile1 - 1) Mod 7) * 32, Int((.BGTile1 - 1) / 7) * 32, hdcTilesMask
                End If
            End With
        Else
            With EditMap.Tile(X, Y)
                If .Ground > 0 Then
                    BitBlt hdcBack(0), X * 32, Y * 32, 32, 32, hdcTiles, ((.Ground - 1) Mod 7) * 32, Int((.Ground - 1) / 7) * 32, SRCCOPY
                    BitBlt hdcBack(1), X * 32, Y * 32, 32, 32, hdcTiles, ((.Ground - 1) Mod 7) * 32, Int((.Ground - 1) / 7) * 32, SRCCOPY
                Else
                    BitBlt hdcBack(0), X * 32, Y * 32, 32, 32, 0, 0, 0, BLACKNESS
                    BitBlt hdcBack(1), X * 32, Y * 32, 32, 32, 0, 0, 0, BLACKNESS
                End If
                If .BGTile1 > 0 Then
                    TransparentBlt hdcBack(0), X * 32, Y * 32, 32, 32, hdcTiles, ((.BGTile1 - 1) Mod 7) * 32, Int((.BGTile1 - 1) / 7) * 32, hdcTilesMask
                End If
                If .BGTile2 > 0 Then
                    TransparentBlt hdcBack(1), X * 32, Y * 32, 32, 32, hdcTiles, ((.BGTile2 - 1) Mod 7) * 32, Int((.BGTile2 - 1) / 7) * 32, hdcTilesMask
                ElseIf .BGTile1 > 0 Then
                    TransparentBlt hdcBack(1), X * 32, Y * 32, 32, 32, hdcTiles, ((.BGTile1 - 1) Mod 7) * 32, Int((.BGTile1 - 1) / 7) * 32, hdcTilesMask
                End If
                If .FGTile > 0 Then
                    BitBlt hdcFront, X * 32, Y * 32, 32, 32, hdcTiles, ((.FGTile - 1) Mod 7) * 32, Int((.FGTile - 1) / 7) * 32, SRCCOPY
                    BitBlt hdcFrontMask, X * 32, Y * 32, 32, 32, hdcTilesMask, ((.FGTile - 1) Mod 7) * 32, Int((.FGTile - 1) / 7) * 32, SRCCOPY
                Else
                    BitBlt hdcFront, X * 32, Y * 32, 32, 32, 0, 0, 0, BLACKNESS
                    BitBlt hdcFrontMask, X * 32, Y * 32, 32, 32, 0, 0, 0, WHITENESS
                End If
                If .Att > 0 Then
                    BitBlt hdcFront, X * 32 + 8, Y * 32 + 8, 16, 16, frmMain.picAtts.hdc, ((.Att - 1) Mod 7) * 32 + 8, Int((.Att - 1) / 7) * 32 + 8, SRCCOPY
                    BitBlt hdcFrontMask, X * 32 + 8, Y * 32 + 8, 16, 16, 0, 0, 0, BLACKNESS
                End If
            End With
        End If
        For A = 0 To 49
            With Map.Object(A)
                If .Object > 0 And .X = X And .Y = Y Then
                    B = Object(.Object).Picture
                    If B > 0 Then
                        TransparentBlt hdcBack(0), .X * 32, .Y * 32, 32, 32, hdcObjects, 0, (B - 1) * 32, hdcObjectsMask
                        TransparentBlt hdcBack(1), .X * 32, .Y * 32, 32, 32, hdcObjects, 0, (B - 1) * 32, hdcObjectsMask
                        'StretchBlt hdcBack(0), .X * 32 + .XOffset, .Y * 32 + .YOffset, 16, 16, hdcObjectsMask, 0, (B - 1) * 32, 32, 32, SRCAND
                        'StretchBlt hdcBack(0), .X * 32 + .XOffset, .Y * 32 + .YOffset, 16, 16, hdcObjects, 0, (B - 1) * 32, 32, 32, SRCPAINT
                        'StretchBlt hdcBack(1), .X * 32 + .XOffset, .Y * 32 + .YOffset, 16, 16, hdcObjectsMask, 0, (B - 1) * 32, 32, 32, SRCAND
                        'StretchBlt hdcBack(1), .X * 32 + .XOffset, .Y * 32 + .YOffset, 16, 16, hdcObjects, 0, (B - 1) * 32, 32, 32, SRCPAINT
                    End If
                End If
            End With
        Next A
    End If
End Sub
Function SwearFilter(ByVal St As String) As String
    Dim A As Long
    A = InStr(UCase$(St), "FUCK")
    While A > 0
        If Mid$(St, A, 4) = "fuck" Then
            St = Mid$(St, 1, A - 1) + "frick" + Mid$(St, A + 4)
        ElseIf Mid$(St, A, 4) = "Fuck" Then
            St = Mid$(St, 1, A - 1) + "Frick" + Mid$(St, A + 4)
        Else
            St = Mid$(St, 1, A - 1) + "FRICK" + Mid$(St, A + 4)
        End If
        A = InStr(UCase$(St), "FUCK")
    Wend
    A = InStr(UCase$(St), "SHIT")
    While A > 0
        If Mid$(St, A, 4) = "shit" Then
            St = Mid$(St, 1, A - 1) + "poop" + Mid$(St, A + 4)
        ElseIf Mid$(St, A, 4) = "Shit" Then
            St = Mid$(St, 1, A - 1) + "Poop" + Mid$(St, A + 4)
        Else
            St = Mid$(St, 1, A - 1) + "POOP" + Mid$(St, A + 4)
        End If
        A = InStr(UCase$(St), "SHIT")
    Wend
    A = InStr(UCase$(St), "ASSHOLE")
    While A > 0
        If Mid$(St, A, 7) = "asshole" Then
            St = Mid$(St, 1, A - 1) + "bunghole" + Mid$(St, A + 7)
        ElseIf Mid$(St, A, 7) = "Asshole" Then
            St = Mid$(St, 1, A - 1) + "Bunghole" + Mid$(St, A + 7)
        Else
            St = Mid$(St, 1, A - 1) + "BUNGHOLE" + Mid$(St, A + 7)
        End If
        A = InStr(UCase$(St), "ASSHOLE")
    Wend
    A = InStr(UCase$(St), "DAMNIT")
    While A > 0
        If Mid$(St, A, 6) = "damnit" Then
            St = Mid$(St, 1, A - 1) + "darn" + Mid$(St, A + 6)
        ElseIf Mid$(St, A, 6) = "Damnit" Then
            St = Mid$(St, 1, A - 1) + "Darn" + Mid$(St, A + 6)
        Else
            St = Mid$(St, 1, A - 1) + "DARN" + Mid$(St, A + 6)
        End If
        A = InStr(UCase$(St), "DAMNIT")
    Wend
    A = InStr(UCase$(St), "BITCH")
    While A > 0
        If Mid$(St, A, 5) = "bitch" Then
            St = Mid$(St, 1, A - 1) + "wench" + Mid$(St, A + 5)
        ElseIf Mid$(St, A, 5) = "Bitch" Then
            St = Mid$(St, 1, A - 1) + "Wench" + Mid$(St, A + 5)
        Else
            St = Mid$(St, 1, A - 1) + "WENCH" + Mid$(St, A + 5)
        End If
        A = InStr(UCase$(St), "BITCH")
    Wend
    SwearFilter = St
End Function
Sub Transition()
    Dim A As Long, T As Long
    
    DrawNextFrame
    
    Select Case NextTransition
        Case 0 'None
            PlayWav 6
        Case 1 'Player Moved Up
            PlayWav 1
            For A = 4 To 384 Step 8
                T = GetTickCount
                BitBlt hdcViewport, 0, 8, 384, 376, hdcViewport, 0, 0, SRCCOPY
                BitBlt hdcViewport, 0, 0, 384, 8, hdcBuffer, 0, 384 - A, SRCCOPY
                While GetTickCount - T < 1
                Wend
            Next A
        Case 2 'Player Moved Down
            PlayWav 1
            For A = 4 To 384 Step 8
                T = GetTickCount
                BitBlt hdcViewport, 0, 0, 384, 376, hdcViewport, 0, 8, SRCCOPY
                BitBlt hdcViewport, 0, 376, 384, 8, hdcBuffer, 0, A, SRCCOPY
                While GetTickCount - T < 1
                Wend
            Next A
        Case 3 'Player Moved Left
            PlayWav 1
            For A = 4 To 384 Step 8
                T = GetTickCount
                BitBlt hdcViewport, 8, 0, 376, 384, hdcViewport, 0, 0, SRCCOPY
                BitBlt hdcViewport, 0, 0, 8, 384, hdcBuffer, 384 - A, 0, SRCCOPY
                While GetTickCount - T < 1
                Wend
            Next A
        Case 4 'Player Moved Right
            PlayWav 1
            For A = 4 To 384 Step 8
                T = GetTickCount
                BitBlt hdcViewport, 0, 0, 376, 384, hdcViewport, 8, 0, SRCCOPY
                BitBlt hdcViewport, 376, 0, 8, 384, hdcBuffer, A, 0, SRCCOPY
                While GetTickCount - T < 1
                Wend
            Next A
        Case 5 'Warp
            PlayWav 6
            Select Case Int(Rnd * 5)
                Case 0
                    For A = 0 To 192 Step 4
                        T = GetTickCount
                        BitBlt hdcViewport, A, A, 383 - 2 * A, 4, hdcBuffer, A, A, SRCCOPY
                        BitBlt hdcViewport, A, 379 - A, 383 - 2 * A, 4, hdcBuffer, A, 379 - A, SRCCOPY
                        BitBlt hdcViewport, A, A, 4, 383 - 2 * A, hdcBuffer, A, A, SRCCOPY
                        BitBlt hdcViewport, 379 - A, A, 4, 383 - 2 * A, hdcBuffer, 379 - A, A, SRCCOPY
                        While GetTickCount - T < 2
                        Wend
                    Next A
                Case 1
                    For A = 192 To 0 Step -4
                        T = GetTickCount
                        BitBlt hdcViewport, A, A, 383 - 2 * A, 4, hdcBuffer, A, A, SRCCOPY
                        BitBlt hdcViewport, A, 379 - A, 383 - 2 * A, 4, hdcBuffer, A, 379 - A, SRCCOPY
                        BitBlt hdcViewport, A, A, 4, 383 - 2 * A, hdcBuffer, A, A, SRCCOPY
                        BitBlt hdcViewport, 379 - A, A, 4, 383 - 2 * A, hdcBuffer, 379 - A, A, SRCCOPY
                        While GetTickCount - T < 2
                        Wend
                    Next A
                Case 2
                    For A = 0 To 48
                        T = GetTickCount
                        BitBlt hdcViewport, 0, A, 384, 1, hdcBuffer, 0, A, SRCCOPY
                        BitBlt hdcViewport, 0, A + 48, 384, 1, hdcBuffer, 0, A + 48, SRCCOPY
                        BitBlt hdcViewport, 0, A + 96, 384, 1, hdcBuffer, 0, A + 96, SRCCOPY
                        BitBlt hdcViewport, 0, A + 144, 384, 1, hdcBuffer, 0, A + 144, SRCCOPY
                        BitBlt hdcViewport, 0, A + 192, 384, 1, hdcBuffer, 0, A + 192, SRCCOPY
                        BitBlt hdcViewport, 0, A + 240, 384, 1, hdcBuffer, 0, A + 240, SRCCOPY
                        BitBlt hdcViewport, 0, A + 288, 384, 1, hdcBuffer, 0, A + 288, SRCCOPY
                        BitBlt hdcViewport, 0, A + 336, 384, 1, hdcBuffer, 0, A + 336, SRCCOPY
                        While GetTickCount - T < 2
                        Wend
                    Next A
                Case 3
                    For A = 0 To 48
                        T = GetTickCount
                        BitBlt hdcViewport, A, 0, 1, 384, hdcBuffer, A, 0, SRCCOPY
                        BitBlt hdcViewport, A + 48, 0, 1, 384, hdcBuffer, A + 48, 0, SRCCOPY
                        BitBlt hdcViewport, A + 96, 0, 1, 384, hdcBuffer, A + 96, 0, SRCCOPY
                        BitBlt hdcViewport, A + 144, 0, 1, 384, hdcBuffer, A + 144, 0, SRCCOPY
                        BitBlt hdcViewport, A + 192, 0, 1, 384, hdcBuffer, A + 192, 0, SRCCOPY
                        BitBlt hdcViewport, A + 240, 0, 1, 384, hdcBuffer, A + 240, 0, SRCCOPY
                        BitBlt hdcViewport, A + 288, 0, 1, 384, hdcBuffer, A + 288, 0, SRCCOPY
                        BitBlt hdcViewport, A + 336, 0, 1, 384, hdcBuffer, A + 336, 0, SRCCOPY
                        While GetTickCount - T < 2
                        Wend
                    Next A
                Case 4
                    For A = 0 To 188 Step 4
                        T = GetTickCount
                        BitBlt hdcViewport, 192 + A, 0, 4, 192, hdcBuffer, 192 + A, 0, SRCCOPY
                        BitBlt hdcViewport, 192, 192 + A, 192, 4, hdcBuffer, 192, 192 + A, SRCCOPY
                        BitBlt hdcViewport, 188 - A, 192, 4, 192, hdcBuffer, 188 - A, 192, SRCCOPY
                        BitBlt hdcViewport, 0, 188 - A, 192, 4, hdcBuffer, 0, 188 - A, SRCCOPY
                        While GetTickCount - T < 2
                        Wend
                    Next A
            End Select
        Case 6 'Death
            For A = 0 To 192 Step 4
                T = GetTickCount
                BitBlt hdcViewport, A, A, 383 - 2 * A, 4, 0, 0, 0, BLACKNESS
                BitBlt hdcViewport, A, 379 - A, 383 - 2 * A, 4, 0, 0, 0, BLACKNESS
                BitBlt hdcViewport, A, A, 4, 383 - 2 * A, 0, 0, 0, BLACKNESS
                BitBlt hdcViewport, 379 - A, A, 4, 383 - 2 * A, 0, 0, 0, BLACKNESS
                While GetTickCount - T < 2
                Wend
            Next A
            For A = 192 To 0 Step -4
                T = GetTickCount
                BitBlt hdcViewport, A, A, 383 - 2 * A, 4, hdcBuffer, A, A, SRCCOPY
                BitBlt hdcViewport, A, 379 - A, 383 - 2 * A, 4, hdcBuffer, A, 379 - A, SRCCOPY
                BitBlt hdcViewport, A, A, 4, 383 - 2 * A, hdcBuffer, A, A, SRCCOPY
                BitBlt hdcViewport, 379 - A, A, 4, 383 - 2 * A, hdcBuffer, 379 - A, A, SRCCOPY
                While GetTickCount - T < 2
                Wend
            Next A
        Case Else
            PlayWav 6
        End Select
    NextTransition = 0
End Sub
Sub UpdatePlayerColor(Index As Long)
    Dim A As Long
    With Player(Index)
        If .Guild > 0 Then
            If .Guild = Character.Guild Then
                .Color = 11
            Else
                .Color = 15
                If Character.Guild > 0 Then
                    For A = 0 To 4
                        If Character.GuildDeclaration(A).Guild = .Guild Then
                            If Character.GuildDeclaration(A).Type = 0 Then
                                .Color = 10
                            Else
                                .Color = 12
                            End If
                        End If
                    Next A
                End If
            End If
        Else
            .Color = 7
        End If
    End With
End Sub
Sub UpdatePlayersColors()
    Dim A As Long, B As Long, C As Long
    For A = 1 To 255
        UpdatePlayerColor A
    Next A
End Sub
Sub UpdateSaleItem(A As Long)
    Dim St As String
    With SaleItem(A)
        If .GiveObject >= 1 And .TakeObject >= 1 Then
            St = CStr(A) + ": "
            If Object(.GiveObject).Type = 6 Then
                'Money
                St = St + CStr(.GiveValue) + " " + Object(.GiveObject).Name
            Else
                St = St + "1 " + Object(.GiveObject).Name
            End If
            St = St + " in exchange for "
            If Object(.TakeObject).Type = 6 Then
                'Money
                St = St + CStr(.TakeValue) + " " + Object(.TakeObject).Name
            Else
                St = St + "1 " + Object(.TakeObject).Name
            End If
            frmNPC.lstSaleItems.List(A) = St
        Else
            frmNPC.lstSaleItems.List(A) = CStr(A) + ":"
        End If
    End With
End Sub
Sub WriteString(lpAppName, lpKeyName As String, A)
    Dim lpString As String, Valid As Long
    lpString = A
    Valid = WritePrivateProfileString&(lpAppName, lpKeyName, lpString, App.Path + "\odyssey.ini")
End Sub

Sub CheckKeys()
    Dim A As Long, B As Long, C As Long, D As Long
    
    If keyCtrl = True Then
        If GetTickCount - AttackTimer >= 1000 Then
            AttackTimer = GetTickCount
            Dim TX As Long, TY As Long
            Select Case CDir
                Case 0
                    TX = CX
                    TY = CY - 1
                Case 1
                    TX = CX
                    TY = CY + 1
                Case 2
                    TX = CX - 1
                    TY = CY
                Case 3
                    TX = CX + 1
                    TY = CY
            End Select
            If TX >= 0 And TX <= 11 And TY >= 0 And TY <= 11 Then
                For A = 0 To 5
                    With Map.Monster(A)
                        If .Monster > 0 And .X = TX And .Y = TY Then
                            SendSocket Chr$(26) + Chr$(A)
                            Exit For
                        End If
                    End With
                Next A
                If A = 6 Then
                    For A = 1 To 255
                        With Player(A)
                            If .Map = CMap And .Sprite > 0 And .X = TX And .Y = TY Then
                                SendSocket Chr$(25) + Chr$(A)
                                Exit For
                            End If
                        End With
                    Next A
                    If A = 256 Then
                        PrintInfoText "There is nobody there to attack!"
                    End If
                End If
            Else
                PrintInfoText "There is nobody there to attack!"
            End If
        End If
    End If
    If CX * 32 = CXO And CY * 32 = CYO Then
        #If CHEATS Then
        If AutoAttack = False Then
        #End If
            If keyShift = True And Character.Energy > 0 Then
                CWalkStep = 8
            Else
                CWalkStep = 4
            End If
            If keyUp = True Then
                If CDir = 0 Then
                    If CY > 0 Then
                        If IsVacant(CX, CY - 1) Then
                            CY = CY - 1
                            MoveToTile
                        End If
                    Else
                        If Map.ExitUp > 0 Then
                            SendSocket Chr$(13) + Chr$(0)
                            Freeze = True
                            NextTransition = 1
                        End If
                    End If
                Else
                    CDir = 0
                    SendSocket Chr$(7) + Chr$(CWalkCode) + Chr$(CX) + Chr$(CY) + Chr$(CDir) + Chr$(CWalkStep)
                End If
            ElseIf keyDown = True Then
                If CDir = 1 Then
                    If CY < 11 Then
                        If IsVacant(CX, CY + 1) Then
                            CY = CY + 1
                            MoveToTile
                        End If
                    Else
                        If Map.ExitDown > 0 Then
                            SendSocket Chr$(13) + Chr$(1)
                            Freeze = True
                            NextTransition = 2
                        End If
                    End If
                Else
                    CDir = 1
                    SendSocket Chr$(7) + Chr$(CWalkCode) + Chr$(CX) + Chr$(CY) + Chr$(CDir) + Chr$(CWalkStep)
                End If
            ElseIf keyLeft = True Then
                If CDir = 2 Then
                    If CX > 0 Then
                        If IsVacant(CX - 1, CY) Then
                            CX = CX - 1
                            MoveToTile
                        End If
                    Else
                        If Map.ExitLeft > 0 Then
                            SendSocket Chr$(13) + Chr$(2)
                            Freeze = True
                            NextTransition = 3
                        End If
                    End If
                Else
                    CDir = 2
                    SendSocket Chr$(7) + Chr$(CWalkCode) + Chr$(CX) + Chr$(CY) + Chr$(CDir) + Chr$(CWalkStep)
                End If
            ElseIf keyRight = True Then
                If CDir = 3 Then
                    If CX < 11 Then
                        If IsVacant(CX + 1, CY) Then
                            CX = CX + 1
                            MoveToTile
                        End If
                    Else
                        If Map.ExitRight > 0 Then
                            SendSocket Chr$(13) + Chr$(3)
                            Freeze = True
                            NextTransition = 4
                        End If
                    End If
                Else
                    CDir = 3
                    SendSocket Chr$(7) + Chr$(CWalkCode) + Chr$(CX) + Chr$(CY) + Chr$(CDir) + Chr$(CWalkStep)
                End If
            End If
        #If CHEATS Then
        Else
            If TargetMonster >= 0 Then
                With Map.Monster(TargetMonster)
                    If .Monster > 0 Then
                        If (Abs(CX - .X) + Abs(CY - .Y)) > 1 Then
                            'Pick up any gold
                            For A = 0 To 49
                                With Map.Object(A)
                                    If .Object = 1 And .X = CX And .Y = CY Then
                                        SendSocket Chr$(8) + Chr$(A)
                                    End If
                                End With
                            Next A
                            
                            'Move
                            CWalkStep = AutoAttackWalk
                            If CX < .X Then
                                If IsVacant(CX + 1, CY) Then
                                    CX = CX + 1
                                    CDir = 3
                                End If
                            ElseIf CX > .X Then
                                If IsVacant(CX - 1, CY) Then
                                    CX = CX - 1
                                    CDir = 2
                                End If
                            End If
                            If CX * 32 = CXO And CY * 32 = CYO Then
                                If CY < .Y Then
                                    If IsVacant(CX, CY + 1) Then
                                        CY = CY + 1
                                        CDir = 1
                                    End If
                                ElseIf CY > .Y Then
                                    If IsVacant(CX, CY - 1) Then
                                        CY = CY - 1
                                        CDir = 0
                                    End If
                                End If
                            End If
                            SendSocket Chr$(7) + Chr$(CWalkCode) + Chr$(CX) + Chr$(CY) + Chr$(CDir) + Chr$(CWalkStep)
                        Else
                            If CX < .X Then
                                A = 3
                            ElseIf CX > .X Then
                                A = 2
                            ElseIf CY < .Y Then
                                A = 1
                            Else
                                A = 0
                            End If
                            
                            If CDir <> A Then
                                CDir = A
                                SendSocket Chr$(7) + Chr$(CWalkCode) + Chr$(CX) + Chr$(CY) + Chr$(CDir) + Chr$(CWalkStep)
                            End If
                            
                            'Attack
                            If GetTickCount - AttackTimer >= AutoAttackSpeed Then
                                AttackTimer = GetTickCount
                                SendSocket Chr$(26) + Chr$(TargetMonster)
                            End If
                        End If
                    Else
                        TargetMonster = -1
                    End If
                End With
            Else
                B = 99999
                For A = 0 To 5
                    With Map.Monster(A)
                        If .Monster > 0 Then
                            C = (CX - .X) ^ 2 + (CY - .Y) ^ 2
                            If C < B Then
                                TargetMonster = A
                                B = C
                            End If
                        End If
                    End With
                Next A
            End If
        End If
        #End If
    End If
End Sub
Sub DrawNextFrame()
    Dim A As Long, B As Long, C As Long, R As RECT
    
    'Copy Back Buffer to Viewport
    If DisableScreen = False Then BitBlt hdcBuffer, 0, 0, 384, 384, hdcBack(CurFrame), 0, 0, SRCCOPY
    
    'Move You
    If CXO < CX * 32 Then
        CXO = CXO + CWalkStep
        If Int(CXO / 16) * 16 = CXO Then
            CWalk = 1 - CWalk
            If CWalk = 0 Then PlayWav 4
        End If
    ElseIf CXO > CX * 32 Then
        CXO = CXO - CWalkStep
        If Int(CXO / 16) * 16 = CXO Then
            CWalk = 1 - CWalk
            If CWalk = 0 Then PlayWav 4
        End If
    End If
    If CYO < CY * 32 Then
        CYO = CYO + CWalkStep
        If Int(CYO / 16) * 16 = CYO Then
            CWalk = 1 - CWalk
            If CWalk = 0 Then PlayWav 4
        End If
    ElseIf CYO > CY * 32 Then
        CYO = CYO - CWalkStep
        If Int(CYO / 16) * 16 = CYO Then
            CWalk = 1 - CWalk
            If CWalk = 0 Then PlayWav 4
        End If
    End If
    
    If DisableScreen = True Then Exit Sub
    
    'Draw You
    If CAttack > 0 Then
        B = CDir * 3 + 2
        CAttack = CAttack - 1
    Else
        B = CDir * 3 + CWalk
    End If
    If Character.Sprite >= 200 Then
        TransparentBlt hdcBuffer, CXO, CYO - 32, 64, 64, hdcLSprites, B * 64, ((Character.Sprite - 200) - 1) * 64, hdcLSpritesMask
    Else
        TransparentBlt hdcBuffer, CXO, CYO - 16, 32, 32, hdcSprites, B * 32, (Character.Sprite - 1) * 32, hdcSpritesMask
    End If
    
    For A = 0 To 5
        With Map.Monster(A)
            If .Monster > 0 Then
                C = Monster(.Monster).Sprite
                If C > 0 Then
                    If .XO < .X * 32 Then
                        .XO = .XO + 4
                        If Int(.XO / 16) * 16 = .XO Then .W = 1 - .W
                    ElseIf .XO > .X * 32 Then
                        .XO = .XO - 4
                        If Int(.XO / 16) * 16 = .XO Then .W = 1 - .W
                    End If
                    If .YO < .Y * 32 Then
                        .YO = .YO + 4
                        If Int(.YO / 16) * 16 = .YO Then .W = 1 - .W
                    ElseIf .YO > .Y * 32 Then
                        .YO = .YO - 4
                        If Int(.YO / 16) * 16 = .YO Then .W = 1 - .W
                    End If
                    
                    'Draw Monster
                    If .A > 0 Then
                        B = .D * 3 + 2
                        .A = .A - 1
                    Else
                        B = .D * 3 + .W
                    End If
                    If C >= 200 Then
                        TransparentBlt hdcBuffer, .XO, .YO - 32, 64, 64, hdcLSprites, B * 64, ((C - 200) - 1) * 64, hdcLSpritesMask
                    Else
                        TransparentBlt hdcBuffer, .XO, .YO - 16, 32, 32, hdcSprites, B * 32, (C - 1) * 32, hdcSpritesMask
                    End If
                End If
            End If
        End With
    Next A
    
    For A = 1 To 255
        With Player(A)
            If .Map = CMap Then
                'Move Player
                If .XO < .X * 32 Then
                    .XO = .XO + .WalkStep
                    If Int(.XO / 16) * 16 = .XO Then .W = 1 - .W
                ElseIf .XO > .X * 32 Then
                    .XO = .XO - .WalkStep
                    If Int(.XO / 16) * 16 = .XO Then .W = 1 - .W
                End If
                If .YO < .Y * 32 Then
                    .YO = .YO + .WalkStep
                    If Int(.YO / 16) * 16 = .YO Then .W = 1 - .W
                ElseIf .YO > .Y * 32 Then
                    .YO = .YO - .WalkStep
                    If Int(.YO / 16) * 16 = .YO Then .W = 1 - .W
                End If
                
                'Draw Player
                If .A > 0 Then
                    B = .D * 3 + 2
                    .A = .A - 1
                Else
                    B = .D * 3 + .W
                End If
                If .Sprite >= 200 Then
                    TransparentBlt hdcBuffer, .XO, .YO - 32, 64, 64, hdcLSprites, B * 64, ((.Sprite - 200) - 1) * 64, hdcLSpritesMask
                Else
                    TransparentBlt hdcBuffer, .XO, .YO - 16, 32, 32, hdcSprites, B * 32, (.Sprite - 1) * 32, hdcSpritesMask
                End If
            End If
        End With
    Next A
    
    For A = 1 To 255
        With Projectile(A)
            If .Sprite > 0 Then
                Select Case .TargetType
                    Case pttCharacter
                        If .TargetNum = Character.Index Then
                            .X = CXO
                            .Y = CYO
                        Else
                            .X = Player(.TargetNum).XO
                            .Y = Player(.TargetNum).YO
                        End If
                        
                        If GetTickCount - .TimeStamp >= .Speed Then
                            If .Frame < .TotalFrames Then
                                .Frame = .Frame + 1
                            Else
                                If .CurLoop = .LoopCount Then
                                    If .EndSound > 0 Then
                                        PlayWav .EndSound
                                    End If
                                    DestroyEffect (A)
                                Else
                                    .CurLoop = .CurLoop + 1
                                    .Frame = 0
                                End If
                            End If
                            .TimeStamp = GetTickCount
                        End If
                    Case pttPlayer
                        If .TargetNum = Character.Index Then
                            .TargetX = CXO
                            .TargetY = CYO
                        Else
                            .TargetX = Player(.TargetNum).XO
                            .TargetY = Player(.TargetNum).YO
                        End If
                        If .X < .TargetX Then .X = .X + 16
                        If .X > .TargetX Then .X = .X - 16
                        If .Y < .TargetY Then .Y = .Y + 16
                        If .Y > .TargetY Then .Y = .Y - 16
                        
                        If GetTickCount - .TimeStamp >= .Speed Then
                            If .X = .TargetX And .Y = .TargetY Then
                                If .Frame < .TotalFrames Then
                                    .Frame = .Frame + 1
                                Else
                                    DestroyEffect (A)
                                End If
                            End If
                            .TimeStamp = GetTickCount
                        End If
                    Case pttMonster
                        .TargetX = Map.Monster(.TargetNum).XO
                        .TargetY = Map.Monster(.TargetNum).YO
                        If .X < .TargetX Then .X = .X + 16
                        If .X > .TargetX Then .X = .X - 16
                        If .Y < .TargetY Then .Y = .Y + 16
                        If .Y > .TargetY Then .Y = .Y - 16
                        
                        If GetTickCount - .TimeStamp >= .Speed Then
                            If .X = .TargetX And .Y = .TargetY Then
                                If .Frame < .TotalFrames Then
                                    .Frame = .Frame + 1
                                Else
                                    DestroyEffect (A)
                                End If
                            End If
                            .TimeStamp = GetTickCount
                        End If
                    Case pttTile
                        If GetTickCount - .TimeStamp >= .Speed Then
                            If .Frame < .TotalFrames Then
                                .Frame = .Frame + 1
                            Else
                                If .CurLoop = .LoopCount Then
                                    If .EndSound > 0 Then
                                        PlayWav .EndSound
                                    End If
                                    DestroyEffect (A)
                                Else
                                    .CurLoop = .CurLoop + 1
                                    .Frame = 0
                                End If
                            End If
                            .TimeStamp = GetTickCount
                        End If
                    End Select
                TransparentBlt hdcBuffer, .X, .Y - 16, 32, 32, hdcEffects, .Frame * 32, (.Sprite - 1) * 32, hdcEffectsMask
            End If
        End With
    Next A
    
    TransparentBlt hdcBuffer, 0, 0, 384, 384, hdcFront, 0, 0, hdcFrontMask
    
    With R
        .Left = CXO - 32
        .Right = CXO + 64
        .Top = CYO - 32
        .Bottom = CYO - 16
    End With

    If Character.Guild > 0 Then
        If Character.Status = 1 And CurFrame = 0 Then
            Draw3dText hdcBuffer, R, Character.Name, QBColor(4), 2
        Else
            Draw3dText hdcBuffer, R, Character.Name, QBColor(11), 2
        End If
    Else
        If Character.Status = 2 Then
            Draw3dText hdcBuffer, R, Character.Name, QBColor(14), 2
        ElseIf Character.Status = 3 Then
            Draw3dText hdcBuffer, R, Character.Name, QBColor(9), 2
        ElseIf Character.Status = 1 And CurFrame = 0 Then
            Draw3dText hdcBuffer, R, Character.Name, QBColor(4), 2
        Else
            Draw3dText hdcBuffer, R, Character.Name, StatusColors(Character.Status), 2
        End If
    End If
    
    If Character.Status = 21 Then 'Rainbow
        Draw3dText hdcBuffer, R, Character.Name, StatusColors((Int(Rnd * 20))), 2
    End If
    
    For A = 1 To 255
        With Player(A)
            If .Map = CMap Then
                R.Left = .XO - 32
                R.Right = .XO + 64
                R.Top = .YO - 32
                R.Bottom = .YO - 16
                If .Status = 9 Then  'Invisible Name
                Else
                    If .Status = 1 And CurFrame = 0 Then
                        Draw3dText hdcBuffer, R, .Name, QBColor(4), 2
                    ElseIf .Status = 1 And CurFrame = 1 Then
                        Draw3dText hdcBuffer, R, .Name, QBColor(.Color), 2
                    ElseIf .Status = 0 Then
                        Draw3dText hdcBuffer, R, .Name, QBColor(.Color), 2
                    Else
                        Draw3dText hdcBuffer, R, .Name, StatusColors(.Status), 2
                    End If
                    If .Status = 21 Then  'Rainbow
                        Draw3dText hdcBuffer, R, .Name, StatusColors((Int(Rnd * 20))), 2
                    End If
                End If
            End If
        End With
    Next A

    DrawChatString
    DrawInfoText
End Sub
Function Exists(Filename As String) As Boolean
     Exists = (Dir(Filename) <> "")
End Function
Sub CheckFile(Filename As String)
    If Exists(Filename) = False Then
        MsgBox "Error: File " + Chr$(34) + Filename + Chr$(34) + " not found!", vbOKOnly + vbExclamation, TitleString
        End
    End If
End Sub

Sub CloseClientSocket(Action As Byte)

    closesocket ClientSocket
    ClientSocket = INVALID_SOCKET
    
    If frmMain_Loaded = True Then Unload frmMain
    If frmWait_Loaded = True Then Unload frmWait
    If frmCharacter_Loaded = True Then Unload frmCharacter
    If frmNewCharacter_Loaded = True Then Unload frmNewCharacter
    If frmNewPass_Loaded = True Then Unload frmNewPass
    If frmMonster_Loaded = True Then Unload frmMonster
    If frmObject_Loaded = True Then Unload frmObject
    If frmList_Loaded = True Then Unload frmList
    If frmGuilds_Loaded = True Then Unload frmGuilds
    If frmGuild_Loaded = True Then Unload frmGuild
    If frmBan_Loaded = True Then Unload frmBan
    If frmHall_Loaded = True Then Unload frmHall
    If frmOptions_Loaded = True Then Unload frmOptions
    If frmNewGuild_Loaded = True Then Unload frmNewGuild
    If frmLogin_Loaded = True And Action <> 1 Then Unload frmLogin
    If frmAccount_Loaded = True And Action <> 2 Then Unload frmAccount
    
    Select Case Action
        Case 0
            frmMenu.Show
        Case 1
            frmMenu.Show
        Case 2
            frmAccount.Show
        Case 3
            blnEnd = True
        Case Else
            frmMenu.Show
    End Select
End Sub
Sub DeInitialize()
    StopMidi
    
    If ClientSocket <> INVALID_SOCKET Then
        closesocket ClientSocket
    End If
    
    'Unload Graphics
    If hdcTiles <> 0 Then
        SelectObject hdcTiles, obmpTiles
        DeleteObject hbmpTiles
        DeleteDC hdcTiles
    End If
    
    If hdcTilesMask <> 0 Then
        SelectObject hdcTilesMask, obmpTilesMask
        DeleteObject hbmpTilesMask
        DeleteDC hdcTilesMask
    End If
    
    If hdcObjects <> 0 Then
        SelectObject hdcObjects, obmpObjects
        DeleteObject hbmpObjects
        DeleteDC hdcObjects
    End If
    
    If hdcObjectsMask <> 0 Then
        SelectObject hdcObjectsMask, obmpObjectsMask
        DeleteObject hbmpObjectsMask
        DeleteDC hdcObjectsMask
    End If
    
    If hdcSprites <> 0 Then
        SelectObject hdcSprites, obmpSprites
        DeleteObject hbmpSprites
        DeleteDC hdcSprites
    End If
    
    If hdcSpritesMask <> 0 Then
        SelectObject hdcSpritesMask, obmpSpritesMask
        DeleteObject hbmpSpritesMask
        DeleteDC hdcSpritesMask
    End If
    
    If hdcLSprites <> 0 Then
        SelectObject hdcLSprites, obmpLSprites
        DeleteObject hbmpLSprites
        DeleteDC hdcLSprites
    End If
    
    If hdcLSpritesMask <> 0 Then
        SelectObject hdcLSpritesMask, obmpLSpritesMask
        DeleteObject hbmpLSpritesMask
        DeleteDC hdcLSpritesMask
    End If
    
    If hdcEffects <> 0 Then
        SelectObject hdcEffects, obmpEffects
        DeleteObject hbmpEffects
        DeleteDC hdcEffects
    End If
    
    If hdcEffectsMask <> 0 Then
        SelectObject hdcEffectsMask, obmpEffectsMask
        DeleteObject hbmpEffectsMask
        DeleteDC hdcEffectsMask
    End If
       
    If hdcBuffer <> 0 Then
        SelectObject hdcBuffer, obmpBuffer
        DeleteObject hbmpBuffer
        DeleteDC hdcBuffer
    End If
    
    If hdcBack(0) <> 0 Then
        SelectObject hdcBack(0), obmpBack(0)
        DeleteObject hbmpBack(0)
        DeleteDC hdcBack(0)
    End If
    
    If hdcBack(1) <> 0 Then
        SelectObject hdcBack(1), obmpBack(1)
        DeleteObject hbmpBack(1)
        DeleteDC hdcBack(1)
    End If
    
    If hdcFront <> 0 Then
        SelectObject hdcFront, obmpFront
        DeleteObject hbmpFront
        DeleteDC hdcFront
    End If
    
    If hdcFrontMask <> 0 Then
        SelectObject hdcFrontMask, obmpFrontMask
        DeleteObject hbmpFrontMask
        DeleteDC hdcFrontMask
    End If
    
    If hdcNight <> 0 Then
        SelectObject hdcNight, obmpNight
        DeleteObject hbmpNight
        DeleteDC hdcNight
    End If
    
    If hdcNightMask <> 0 Then
        SelectObject hdcNightMask, obmpNightMask
        DeleteObject hbmpNightMask
        DeleteDC hdcNightMask
    End If
    
    If hPalette <> 0 Then
        DeleteObject hPalette
    End If
    
    'Unload Winsock
    EndWinsock
    
    'Unhook Form
    Unhook
    
    End
End Sub
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
Function WindowProc(ByVal hw As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Dim St1 As String
    If uMsg = 1025 Then
        'Client Socket
        Select Case lParam And 255
            Case FD_CLOSE
                CloseClientSocket 1
                MessageBox frmMenu.hWnd, "You have been disconnected from the server!", TitleString, vbOKOnly + vbExclamation
            Case FD_CONNECT
                If lParam = FD_CONNECT Then
                    St1 = EncryptString(ComputerID)
                    SendSocket Chr$(61) + Chr$(ClientVer) + Trim$(St1) + "poo"
                    If NewAccount = True Then
                        frmWait.lblStatus = "Sending Account Information ..."
                        SendSocket Chr$(0) + User + Chr$(0) + EncryptString(Pass)
                    Else
                        frmWait.lblStatus = "Sending Login Information ..."
                        SendSocket Chr$(1) + User + Chr$(0) + EncryptString(Pass)
                    End If
                Else
                    CloseClientSocket 0
                    WaitForConnect "Error Connecting - Waiting"
                End If
            Case FD_READ
                If lParam = FD_READ Then ReceiveData
        End Select
    End If
    WindowProc = CallWindowProc(lpPrevWndProc, hw, uMsg, wParam, lParam)
End Function

Function CheckSum(St As String) As Long
    Dim A As Long, B As Long
    For A = 1 To Len(St)
        B = B + Asc(Mid$(St, A, 1))
    Next A
    CheckSum = B
End Function

Function FindPlayer(ByVal St As String) As Long
    Dim A As Long, StLen As Long
    St = UCase$(St)
    
    'Search for exact match
    For A = 1 To 255
        With Player(A)
            If .Sprite > 0 And UCase$(.Name) = St Then
                FindPlayer = A
                Exit Function
            End If
        End With
    Next A
    
    'Search for partial match
    StLen = Len(St)
    For A = 1 To 255
        With Player(A)
            If .Sprite > 0 Then
                If Len(.Name) >= StLen Then
                    If UCase$(Left$(.Name, StLen)) = St Then
                        FindPlayer = A
                        Exit Function
                    End If
                End If
            End If
        End With
    Next A
End Function

Function FindGuild(ByVal St As String) As Long
    Dim A As Long, StLen As Long
    St = UCase$(St)
    
    'Search for exact match
    For A = 1 To 255
        With Guild(A)
            If UCase$(.Name) = St Then
                FindGuild = A
                Exit Function
            End If
        End With
    Next A
    
    'Search for partial match
    StLen = Len(St)
    For A = 1 To 255
        With Guild(A)
            If Len(.Name) >= StLen Then
                If UCase$(Left$(.Name, StLen)) = St Then
                    FindGuild = A
                    Exit Function
                End If
            End If
        End With
    Next A
End Function
Function GetInt(Chars As String) As Long
    GetInt = Asc(Mid$(Chars, 1, 1)) * 256 + Asc(Mid$(Chars, 2, 1))
End Function

Sub SendSocket(ByVal St As String)
    If SendData(ClientSocket, DoubleChar(Len(St)) + St) = SOCKET_ERROR Then
        CloseClientSocket 0
    End If
End Sub
Function DoubleChar(Num As Long) As String
    DoubleChar = Chr$(Int(Num / 256)) + Chr$(Num Mod 256)
End Function
Sub UploadMap()
    Dim MapData As String, St1 As String * 30
    Dim X As Long, Y As Long
    With EditMap
        If .Version < 2147483647 Then
            .Version = .Version + 1
        Else
            .Version = 1
        End If
        St1 = .Name
        MapData = St1 + QuadChar(.Version) + Chr$(.NPC) + Chr$(.MIDI) + DoubleChar$(CLng(.ExitUp)) + DoubleChar$(CLng(.ExitDown)) + DoubleChar$(CLng(.ExitLeft)) + DoubleChar$(CLng(.ExitRight)) + DoubleChar(CLng(.BootLocation.Map)) + Chr$(.BootLocation.X) + Chr$(.BootLocation.Y) + Chr$(.Flags) + Chr$(.MonsterSpawn(0).Monster) + Chr$(.MonsterSpawn(0).Rate) + Chr$(.MonsterSpawn(1).Monster) + Chr$(.MonsterSpawn(1).Rate) + Chr$(.MonsterSpawn(2).Monster) + Chr$(.MonsterSpawn(2).Rate)
        For Y = 0 To 11
            For X = 0 To 11
                With .Tile(X, Y)
                    MapData = MapData + DoubleChar(CLng(.Ground)) + DoubleChar(CLng(.BGTile1)) + DoubleChar(CLng(.BGTile2)) + DoubleChar(CLng(.FGTile)) + Chr$(.Att) + Chr$(.AttData(0)) + Chr$(.AttData(1)) + Chr$(.AttData(2)) + Chr$(.AttData(3))
                End With
            Next X
        Next Y
    End With
    SendSocket Chr$(12) + MapData
End Sub
Sub Main()
    Dim A As Long
    Dim LogPalette As LogPalette
    
    ServerID = "OldMain"
    MapCacheFile = "mcache.dat"
    
    'Get Computer ID
    ComputerID = Trim$(GetWindowsKey())
    
    LoadOptions
    LoadMacros
    
    CreateClassData

    Dim St As String
    Dim Pic As StdPicture
    Dim T As Long
    
    'Check Files
    CheckFile GetFilename("tiles.rsc")
    CheckFile GetFilename("tilesm.rsc")
    CheckFile GetFilename("objects.rsc")
    CheckFile GetFilename("objectsm.rsc")
    CheckFile GetFilename("sprites.rsc")
    CheckFile GetFilename("spritesm.rsc")
    CheckFile GetFilename("Lsprites.rsc")
    CheckFile GetFilename("Lspritesm.rsc")
    CheckFile GetFilename("palette.dat")
    CheckFile GetFilename("Night.rsc")
    CheckFile GetFilename("Nightm.rsc")
    
    frmWait.Show
    frmWait.Refresh
    
    If Exists(MapCacheFile) = False Then
        frmWait.lblStatus = "Creating Map Cache .."
        frmWait.Refresh
        CreateMapCache
    Else
        If FileLen(MapCacheFile) <> 3854000 Then
            frmWait.lblStatus = "Creating Map Cache .."
            frmWait.Refresh
            CreateMapCache
        End If
    End If
    
    frmWait.lblStatus = "Creating Buffers .."
    frmWait.Refresh
    
    hdcBuffer = CreateCompatibleDC(0&)
    hbmpBuffer = CreateCompatibleBitmap(frmWait.hdc, 384, 384)
    obmpBuffer = SelectObject(hdcBuffer, hbmpBuffer)
    
    SetBkMode hdcBuffer, TRANSPARENT
    
    hdcBack(0) = CreateCompatibleDC(0&)
    hbmpBack(0) = CreateCompatibleBitmap(frmWait.hdc, 384, 384)
    obmpBack(0) = SelectObject(hdcBack(0), hbmpBack(0))
    
    hdcBack(1) = CreateCompatibleDC(0&)
    hbmpBack(1) = CreateCompatibleBitmap(frmWait.hdc, 384, 384)
    obmpBack(1) = SelectObject(hdcBack(1), hbmpBack(1))
    
    hdcFront = CreateCompatibleDC(0&)
    hbmpFront = CreateCompatibleBitmap(frmWait.hdc, 384, 384)
    obmpFront = SelectObject(hdcFront, hbmpFront)
    
    hdcFrontMask = CreateCompatibleDC(0&)
    hbmpFrontMask = CreateCompatibleBitmap(frmWait.hdc, 384, 384)
    obmpFrontMask = SelectObject(hdcFrontMask, hbmpFrontMask)
    
    GetObject hbmpBack(0), Len(Bitmap), Bitmap
    
    If Bitmap.bmBitsPixel = 8 Then
        LogPalette.palVersion = 768
        LogPalette.palNumEntries = 256
        Open GetFilename("palette.dat") For Random As #1 Len = 4
        For A = 0 To 255
            Get #1, , LogPalette.palPalEntry(A)
        Next A
        Close #1
        Paletted = True
    Else
        Paletted = False
    End If
    
    frmWait.lblStatus = "Loading Map Tiles .."
    frmWait.Refresh
    
    hdcTiles = CreateCompatibleDC(0&)
    Set Pic = LoadPicture(GetFilename("tiles.rsc"))
    hbmpTiles = Pic.Handle
    obmpTiles = SelectObject(hdcTiles, hbmpTiles)
    
    hdcTilesMask = CreateCompatibleDC(0&)
    Set Pic = LoadPicture(GetFilename("tilesm.rsc"))
    hbmpTilesMask = Pic.Handle
    obmpTilesMask = SelectObject(hdcTilesMask, hbmpTilesMask)
    
    frmWait.lblStatus = "Loading Object Tiles .."
    frmWait.Refresh
    
    hdcObjects = CreateCompatibleDC(0&)
    Set Pic = LoadPicture(GetFilename("objects.rsc"))
    hbmpObjects = Pic.Handle
    obmpObjects = SelectObject(hdcObjects, hbmpObjects)
    
    hdcObjectsMask = CreateCompatibleDC(0&)
    Set Pic = LoadPicture(GetFilename("objectsm.rsc"))
    hbmpObjectsMask = Pic.Handle
    obmpObjectsMask = SelectObject(hdcObjectsMask, hbmpObjectsMask)
    
    frmWait.lblStatus = "Loading Sprites .."
    frmWait.Refresh
    
    'Load Sprites
    hdcSprites = CreateCompatibleDC(0&)
    Set Pic = LoadPicture(GetFilename("sprites.rsc"))
    hbmpSprites = Pic.Handle
    obmpSprites = SelectObject(hdcSprites, hbmpSprites)
   
    hdcSpritesMask = CreateCompatibleDC(0&)
    Set Pic = LoadPicture(GetFilename("spritesm.rsc"))
    hbmpSpritesMask = Pic.Handle
    obmpSpritesMask = SelectObject(hdcSpritesMask, hbmpSpritesMask)

    hdcLSprites = CreateCompatibleDC(0&)
    Set Pic = LoadPicture(GetFilename("Lsprites.rsc"))
    hbmpLSprites = Pic.Handle
    obmpLSprites = SelectObject(hdcLSprites, hbmpLSprites)
   
    hdcLSpritesMask = CreateCompatibleDC(0&)
    Set Pic = LoadPicture(GetFilename("Lspritesm.rsc"))
    hbmpLSpritesMask = Pic.Handle
    obmpLSpritesMask = SelectObject(hdcLSpritesMask, hbmpLSpritesMask)

    frmWait.lblStatus = "Loading Spell Effects..."
    frmWait.Refresh
    
    hdcEffects = CreateCompatibleDC(0&)
    Set Pic = LoadPicture(GetFilename("Effects.rsc"))
    hbmpEffects = Pic.Handle
    obmpEffects = SelectObject(hdcEffects, hbmpEffects)
   
    hdcEffectsMask = CreateCompatibleDC(0&)
    Set Pic = LoadPicture(GetFilename("Effectsm.rsc"))
    hbmpEffectsMask = Pic.Handle
    obmpEffectsMask = SelectObject(hdcEffectsMask, hbmpEffectsMask)
    
    frmWait.lblStatus = "Loading Weather Effects..."
    frmWait.Refresh
    
    'Night
    hdcNight = CreateCompatibleDC(0&)
    Set Pic = LoadPicture(GetFilename("night.rsc"))
    hbmpNight = Pic.Handle
    obmpNight = SelectObject(hdcNight, hbmpNight)

    hdcNightMask = CreateCompatibleDC(0&)
    Set Pic = LoadPicture(GetFilename("nightm.rsc"))
    hbmpNightMask = Pic.Handle
    obmpNightMask = SelectObject(hdcNightMask, hbmpNightMask)

    'Create Colors
    CreateStatusColors
    
    If Paletted = True Then
        hPalette = CreatePalette(LogPalette)
        SelectPalette hdcBack(0), hPalette, False
        RealizePalette hdcBack(0)
        SelectPalette hdcBack(1), hPalette, False
        RealizePalette hdcBack(1)
        SelectPalette hdcFront, hPalette, False
        RealizePalette hdcFront
        SelectPalette hdcBuffer, hPalette, False
        RealizePalette hdcBuffer
        'SelectPalette hdcTiles, hPalette, False
        'RealizePalette hdcTiles
        'SelectPalette hdcTilesMask, hPalette, False
        'RealizePalette hdcTilesMask
        'SelectPalette hdcSprites, hPalette, False
        'RealizePalette hdcSprites
        'SelectPalette hdcSpritesMask, hPalette, False
        'RealizePalette hdcSpritesMask
        'SelectPalette hdcObjects, hPalette, False
        'RealizePalette hdcObjects
        'SelectPalette hdcObjects, hPalette, False
        'RealizePalette hdcObjectsMask
        'SelectPalette hdcNight, hPalette, False
        'RealizePalette hdcNight
        'SelectPalette hdcNightMask, hPalette, False
        'RealizePalette hdcNightMask
    End If
    frmWait.Refresh
    
    PlayWav 7
    PlayMidi 19
    
    Unload frmWait
    
    Load frmMenu
    
    'Hook Form
    Hook
    
    'Load Winsock
    StartWinsock (St)
    frmMenu.Show
    
    FrameCounter = 0
    FrameTimer = GetTickCount
    blnEnd = False
    
    While blnEnd = False
        FrameCounter = FrameCounter + 1
        If FrameCounter >= 5 Then
            FrameCounter = 0
            CurFrame = 1 - CurFrame
        End If
        SecondCounter = SecondCounter + 1
        If SecondCounter >= 20 Then
            If GetTickCount - FrameTimer > 0 Then
                FrameRate = 20000 / (GetTickCount - FrameTimer)
                FrameTimer = GetTickCount
            End If
            SecondCounter = 0
            If CurrentMIDI > 0 Then
                Dim ReturnString As String
                ReturnString = mmSend("status midi mode")
                If UCase$(ReturnString) = "STOPPED" Then
                    mciSendString "play midi from 0", 0&, 0, 0
                End If
            End If
        End If
        If blnPlaying = True Then
            T = GetTickCount
            If Freeze = False Then
                CheckKeys
                DrawNextFrame
                If DisableScreen = False Then BitBlt hdcViewport, 0, 0, 384, 384, hdcBuffer, 0, 0, SRCCOPY
            End If
        End If
        While GetTickCount - T < 46
            DoEvents
        Wend
        DoEvents
    Wend
    DeInitialize
End Sub
Sub TransparentBlt(hdc As Long, ByVal destX As Long, ByVal destY As Long, destWidth As Long, destHeight As Long, srcDC As Long, srcX As Long, srcY As Long, maskDC As Long)
    BitBlt hdc, destX, destY, destWidth, destHeight, maskDC, srcX, srcY, SRCAND
    BitBlt hdc, destX, destY, destWidth, destHeight, srcDC, srcX, srcY, SRCPAINT
End Sub
Sub PrintChat(ByVal St As String, Color As Byte)
    Dim A As Long, B As Long, FoundLine As Boolean
    Dim Text As String, TextHeight As Long, TextWidth As Long
    
    With frmMain.picChat
        .ForeColor = QBColor(Color)
        TextHeight = .TextHeight("A")
        MoveUp
        While St <> ""
            B = 0
            FoundLine = False
            For A = 1 To Len(St)
                If .TextWidth(Left$(St, A)) > .ScaleWidth - .CurrentX Then
                    FoundLine = True
                    If B = 0 Then
                        B = A - 1
                    End If
                    If B > 0 Then
                        Text = Left$(St, B)
                        St = Mid$(St, B + 1)
                    Else
                        Text = ""
                    End If
                    Exit For
                End If
                If Mid$(St, A, 1) = " " Then B = A
            Next A
            If FoundLine = False Then
                Text = St
                St = ""
            End If
            If Text <> "" Then
                TextWidth = .TextWidth(Text)
                TextOut .hdc, .CurrentX, .ScaleHeight - TextHeight, Text, Len(Text)
                If FoundLine = True Then
                    MoveUp
                Else
                    .CurrentX = .CurrentX + TextWidth
                End If
            Else
                If St <> "" Then
                    MoveUp
                End If
            End If
        Wend
    End With
    frmMain.picChat.Refresh
End Sub
Sub MoveUp()
    Dim TextHeight As Long
    Dim A As Long
    With frmMain.picChat
        A = .TextHeight("A")
        .CurrentX = 0
        BitBlt .hdc, 0, 0, .ScaleWidth, .ScaleHeight - A, .hdc, 0, A, SRCCOPY
        BitBlt .hdc, 0, .ScaleHeight - A, .ScaleWidth, A, 0, 0, 0, 0
    End With
End Sub
Sub YouDied()
    Dim A As Long
    
    With Character
        'Reset Stat Bars
        .HP = .MaxHP
        .Energy = .MaxEnergy
        .Mana = .MaxMana
        DrawHP
        DrawEnergy
        DrawMana
    End With
    
    Freeze = True
    NextTransition = 6
    PlayWav 8
End Sub

Sub AddLog(Message As String)
frmLog.txtLog.Text = frmLog.txtLog.Text + vbCrLf + Trim$(Message)
End Sub

Sub SendAwayMsg(Player As Long, Message As String)
If Message <> "" Then
    If Player > 0 Then
        SendSocket Chr$(14) + Chr$(Player) + Message
    End If
End If
End Sub
Sub ConnectClient()
    Dim St As String
    
    frmWait.Show
    frmWait.lblStatus = "Connecting ..."
    frmWait.btnCancel.Visible = True
    frmWait.Refresh
    
    ClientSocket = ConnectSock(ServerIP, 5678, St, gHW, True)
End Sub
Sub WaitForConnect(Message As String)
    frmWait.Show
    frmWait.lblStatus.Caption = Message
    frmWait.btnCancel.Visible = True
    frmWait.Refresh
    
    frmWait.ConnectTimer.Enabled = True
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
Sub CreateStatusColors()
Dim A As Long

StatusColors(0) = QBColor(7)
StatusColors(1) = QBColor(7)
StatusColors(2) = QBColor(14)
StatusColors(3) = QBColor(9)
StatusColors(4) = &HC0C0FF
StatusColors(5) = &H404000
StatusColors(6) = &H400000
StatusColors(7) = QBColor(6)
StatusColors(8) = QBColor(6)
StatusColors(9) = &H808080
StatusColors(10) = QBColor(0)
StatusColors(11) = &H4000&
StatusColors(12) = &H808000
StatusColors(13) = &HFFC0C0
StatusColors(14) = &H404080
StatusColors(15) = &H80FF&
StatusColors(16) = &HC0FFC0
StatusColors(17) = &HC0&
StatusColors(18) = &H404000
StatusColors(19) = &H800080
StatusColors(20) = &HC0FFFF

For A = 21 To 100
    StatusColors(A) = QBColor(7)
Next A
End Sub
Function QueryValueEx(ByVal lhKey As Long, ByVal szValueName As String, vValue As Variant) As Long
    'This Function Copied from MS Knowledge Base
    'Reads a value from the registry
    Dim cch As Long
    Dim lrc As Long
    Dim lType As Long
    Dim lValue As Long
    Dim sValue As String

    On Error GoTo QueryValueExError

    ' Determine the size and type of data to be read
    lrc = RegQueryValueExNULL(lhKey, szValueName, 0&, lType, 0&, cch)
    If lrc <> ERROR_NONE Then Error 5

    Select Case lType
        Case REG_SZ: ' For strings
            sValue = String(cch, 0)
            lrc = RegQueryValueExString(lhKey, szValueName, 0&, lType, sValue, cch)
            If lrc = ERROR_NONE Then
                vValue = Left$(sValue, cch - 1)
            Else
                vValue = Empty
            End If
        Case REG_DWORD: ' For DWORDS
            lrc = RegQueryValueExLong(lhKey, szValueName, 0&, lType, lValue, cch)
            If lrc = ERROR_NONE Then vValue = lValue
        Case Else 'all other data types not supported
            lrc = -1
        End Select

QueryValueExExit:
    QueryValueEx = lrc
    Exit Function
QueryValueExError:
    Resume QueryValueExExit
End Function
Function GetWindowsKey() As String
    Dim lRetVal As Long, hKey As Long, KeyValue As Variant
    Dim A As Integer, ExecString As String
    
    lRetVal = RegOpenKeyEx(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion", 0, KEY_ALL_ACCESS, hKey)
    If lRetVal = 0 Then
        lRetVal = QueryValueEx(hKey, "ProductKey", KeyValue)
        If lRetVal = 0 Then
            GetWindowsKey = KeyValue
        Else
            'Error Finding Windows Key
        End If
        RegCloseKey (hKey)
    Else
        'Error Finding Windows Key
    End If
End Function
Sub CreateTileEffect(X As Long, Y As Long, Sprite As Long, Speed As Long, TotalFrames As Long, LoopCount As Integer, EndSound As Long)
Dim A As Long

For A = 1 To 255
    With Projectile(A)
        If .Sprite = 0 Then
            .Sprite = Sprite
            .TargetType = 3
            .Frame = 0
            .TotalFrames = TotalFrames
            .Speed = Speed
            .EndSound = EndSound
            .LoopCount = LoopCount
            .X = X * 32
            .Y = Y * 32
            Exit For
        End If
    End With
Next A
End Sub
Sub CreateMonsterEffect(TargetNum As Byte, Sprite As Long, Speed As Long, TotalFrames As Long, SourceX As Long, SourceY As Long, EndSound As Long)
Dim A As Long

For A = 1 To 255
    With Projectile(A)
        If .Sprite = 0 Then
            .Sprite = Sprite
            .TargetType = 2
            .TargetNum = TargetNum
            .Frame = 0
            .TotalFrames = TotalFrames
            .Speed = Speed
            .EndSound = EndSound
            .SourceX = SourceX
            .SourceY = SourceY
            .X = SourceX * 32
            .Y = SourceY * 32
            Exit For
        End If
    End With
Next A
End Sub
Sub CreatePlayerEffect(TargetNum As Byte, Sprite As Long, Speed As Long, TotalFrames As Long, SourceX As Long, SourceY As Long, EndSound As Long)
Dim A As Long

For A = 1 To 255
    With Projectile(A)
        If .Sprite = 0 Then
            .Sprite = Sprite
            .TargetType = 1
            .TargetNum = TargetNum
            .Frame = 0
            .TotalFrames = TotalFrames
            .Speed = Speed
            .EndSound = EndSound
            .SourceX = SourceX
            .SourceY = SourceY
            .X = SourceX * 32
            .Y = SourceY * 32
            Exit For
        End If
    End With
Next A
End Sub
Sub CreateCharacterEffect(TargetNum As Long, Sprite As Long, Speed As Long, TotalFrames As Long, LoopCount As Integer, EndSound As Long)
Dim A As Long

For A = 1 To 255
    With Projectile(A)
        If .Sprite = 0 Then
            .Sprite = Sprite
            .TargetType = 0
            .Frame = 0
            .TotalFrames = TotalFrames
            .Speed = Speed
            .EndSound = EndSound
            .LoopCount = LoopCount
            .TargetNum = TargetNum
            Exit For
        End If
    End With
Next A
End Sub
Sub DestroyEffect(number As Integer)
With Projectile(number)
    .CurLoop = 0
    .D = 0
    .EndSound = 0
    .Frame = 0
    .LoopCount = 0
    .SourceX = 0
    .SourceY = 0
    .Speed = 0
    .Sprite = 0
    .TargetNum = 0
    .TargetType = 0
    .TargetX = 0
    .TargetY = 0
    .TimeStamp = 0
    .TotalFrames = 0
    .X = 0
    .Y = 0
End With
End Sub

