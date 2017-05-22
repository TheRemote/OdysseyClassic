Attribute VB_Name = "modConstants"
Option Explicit

'Editing Mode Constants
Public ListEditMode As Byte
Public Const modeObjects = 1
Public Const modeMonsters = 2
Public Const modeNPCs = 3
Public Const modeHalls = 4
Public Const modeMagic = 5
Public Const modeBans = 6
Public Const modePrefix = 7
Public Const modeSuffix = 8
Public Const modeAbility = 9
Public Const modeSkill = 10

'Maximum Constants
Public Const MaxSprite = 643
Public Const MaxMonsters = 20
Public Const MaxMaps = 3000
Public Const MaxNPCs = 500
Public Const MaxMagic = 500
Public Const MaxGuilds = 255
Public Const MaxHalls = 255
Public Const MaxModifications = 255
Public Const MaxTotalMonsters = 1000
Public Const MaxObjects = 1000
Public Const MaxAtt = 22
Public Const NumClasses = 4
Public Const MaxUsers = 80
Public Const MaxProjectiles = 20
Public Const MaxFloatText = 20
Public Const MaxInvObjects = 20
Public Const MaxMapObjects = 79
Public Const MaxSkill = 10
Public Const MaxRequestLength = 200

'Hacking Code Constant
Public Const HackCode = 98

'Projectile Type Constants
Public Const pttCharacter = 0
Public Const pttPlayer = 1
Public Const pttMonster = 2
Public Const pttTile = 3
Public Const pttProject = 4

'Color Constants
Public Const BLACK = 0
Public Const BLUE = 1
Public Const GREEN = 2
Public Const CYAN = 3
Public Const RED = 4
Public Const MAGENTA = 5
Public Const BROWN = 6
Public Const GREY = 7
Public Const DARKGREY = 8
Public Const BRIGHTBLUE = 9
Public Const BRIGHTGREEN = 10
Public Const BRIGHTCYAN = 11
Public Const BRIGHTRED = 12
Public Const BRIGHTMAGENTA = 13
Public Const YELLOW = 14
Public Const WHITE = 15

'API CONSTANTS

'SetBkMode Constants
Public Const Transparent = 1

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


'SendMessage Constants
Public Const EM_LINEINDEX = &HBB
Public Const EM_LINELENGTH = &HC1

Public Const WM_NCLBUTTONDOWN = &HA1
Public Const HTCAPTION = 2


'WaitForTerm (Used by Scripting)
Public Const SYNCHRONIZE = &H100000
Public Const INFINITE = &HFFFFFFFF

'Bitmap File Header
Public Type BITMAPINFOHEADER2    '40 bytes
    biSize As Long
    biWidth As Long
    biHeight As Long
    biPlanes As Integer
    biBitCount As Integer
    biCompression As Long
    biSizeImage As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed As Long
    biClrImportant As Long
End Type

Public Type BITMAPFILEHEADER
    bfType As Integer
    bfSize As Long
    bfReserved1 As Integer
    bfReserved2 As Integer
    bfOhFileBits As Long
End Type

'GetQueueStatus constants
Public Const QS_HOTKEY = &H80
Public Const QS_KEY = &H1
Public Const QS_MOUSEBUTTON = &H4
Public Const QS_MOUSEMOVE = &H2
Public Const QS_PAINT = &H20
Public Const QS_POSTMESSAGE = &H8
Public Const QS_SENDMESSAGE = &H40
Public Const QS_TIMER = &H10
Public Const QS_ALLINPUT = (QS_SENDMESSAGE Or QS_PAINT Or _
                            QS_TIMER Or QS_POSTMESSAGE Or QS_MOUSEBUTTON Or QS_MOUSEMOVE Or _
                            QS_HOTKEY Or QS_KEY)
Public Const QS_MOUSE = (QS_MOUSEMOVE Or QS_MOUSEBUTTON)
Public Const QS_INPUT = (QS_MOUSE Or QS_KEY)
Public Const QS_ALLEVENTS = (QS_INPUT Or QS_POSTMESSAGE Or QS_TIMER Or QS_PAINT Or QS_HOTKEY)

'Key Constants
Public Const vbKeyAlt = vbKeyControl + 1
