Attribute VB_Name = "modMinimapper"
Option Explicit

Public ServerIP As String
Public SocketData As String
Public ClientSocket As Long
Public User As String, Pass As String, CMap As Long

Public MapDataArray() As Byte
Public MapData As String * 2677
Public RequestedMap As Boolean
Public ReceivedMap As Boolean
Public MapSize As Long
Public OriginalStartLoc As Long

Public Const ClientVer = 34

Public PacketOrder As Integer
Public ServerPacketOrder As Integer

'Hook
Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Public Const GWL_WNDPROC = -4
Public lpPrevWndProc As Long
Public gHW As Long

Public Declare Function EncryptDataString Lib "odysseydll" (ByRef File As Any, ByVal XorValue As Byte) As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

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

Public hPalette As Long

Public Palette(0 To 255) As PaletteEntry
Public Bitmap As Bitmap

Public XOffset As Long, YOffset As Long

Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Declare Function RealizePalette Lib "gdi32" (ByVal hdc As Long) As Long
Declare Function CreatePalette Lib "gdi32" (lpLogPalette As LogPalette) As Long
Declare Function SelectPalette Lib "gdi32" (ByVal hdc As Long, ByVal hPalette As Long, ByVal bForceBackground As Long) As Long
Declare Function SetPaletteEntries Lib "gdi32" (ByVal hPalette As Long, ByVal wStartIndex As Long, ByVal wNumEntries As Long, lpPaletteEntries As PaletteEntry) As Long
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
Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Declare Function SetStretchBltMode Lib "gdi32" (ByVal hdc As Long, ByVal hStretchMode As Long) As Long

Public hdcTiles As Long, hbmpTiles As Long, obmpTiles As Long
Public hdcTilesMask As Long, hbmpTilesMask As Long, obmpTilesMask As Long

Public hdcTempMask As Long, hbmpTempMask As Long, obmpTempMask As Long

Public hdcBack(0 To 1) As Long, hbmpBack(0 To 1) As Long, obmpBack(0 To 1) As Long

Public hdcBuffer As Long, hbmpBuffer As Long, obmpBuffer As Long

Public hdcFront As Long, hbmpFront As Long, obmpFront As Long
Public hdcFrontMask As Long, hbmpFrontMask As Long, obmpFrontMask As Long

'SetBkMode Constants
Public Const TRANSPARENT = 1

'BitBlt Constants
Public Const BLACKNESS = &H42
Public Const WHITENESS = &HFF0062
Public Const SRCCOPY = &HCC0020
Public Const SRCPAINT = &HEE0086
Public Const SRCAND = &H8800C6
Public Const NOTSRCCOPY = &H330008
Public Const SRCINVERT = &H660046
Public Const DSTINVERT = &H550009

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
    Ground2 As Integer
    BGTile1 As Integer
    BGTile2 As Integer
    FGTile As Integer
    Att As Byte
    AttData(0 To 3) As Byte
    Att2 As Byte
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
    MonsterSpawn(0 To 9) As MapMonsterSpawnData
    Door(0 To 9) As MapDoorData
    BootLocation As MapStartLocationData
    DeathLocation As MapStartLocationData
    NPC As Byte
    MIDI As Byte
    flags As Byte
    Flags2 As Byte
    Version As Long
End Type

Type TheMapData
    X As Long
    Y As Long
    Drawn As Boolean
    Received As Boolean
End Type

Public Map As MapData
Public TheMap(0 To 3000) As TheMapData

Public MapArray(0 To 25, 0 To 25) As Integer
Public TileDataArray(0 To 125, 0 To 125) As TileData
Public TileDataArray2(0 To 125, 125 To 255) As TileData
Public TileDataArray3(125 To 255, 0 To 125) As TileData
Public TileDataArray4(125 To 255, 125 To 255) As TileData

Sub Main()
    Dim LogPalette As LogPalette, A As Long, Paletted As Boolean
    Dim Pic As StdPicture, St As String

    CheckFile "tiles.rsc"
    CheckFile "tilesm.rsc"
    
    If Exists("cache1.dat") = False Then
        CreateMapCache
    Else
        If FileLen("cache1.dat") <> 7164000 Then
            CreateMapCache
        End If
    End If
    
    hdcBuffer = CreateCompatibleDC(0&)
    hbmpBuffer = CreateCompatibleBitmap(frmLogin.hdc, 384, 384)
    obmpBuffer = SelectObject(hdcBuffer, hbmpBuffer)
    
    SetBkMode hdcBuffer, TRANSPARENT
    
    hdcBack(0) = CreateCompatibleDC(0&)
    hbmpBack(0) = CreateCompatibleBitmap(frmLogin.hdc, 384, 384)
    obmpBack(0) = SelectObject(hdcBack(0), hbmpBack(0))
    
    hdcBack(1) = CreateCompatibleDC(0&)
    hbmpBack(1) = CreateCompatibleBitmap(frmLogin.hdc, 384, 384)
    obmpBack(1) = SelectObject(hdcBack(1), hbmpBack(1))
    
    hdcFront = CreateCompatibleDC(0&)
    hbmpFront = CreateCompatibleBitmap(frmLogin.hdc, 384, 384)
    obmpFront = SelectObject(hdcFront, hbmpFront)
    
    hdcFrontMask = CreateCompatibleDC(0&)
    hbmpFrontMask = CreateCompatibleBitmap(frmLogin.hdc, 384, 384)
    obmpFrontMask = SelectObject(hdcFrontMask, hbmpFrontMask)
    
    If Bitmap.bmBitsPixel = 8 Then
        LogPalette.palVersion = 768
        LogPalette.palNumEntries = 256
        Open "palette.dat" For Random As #1 Len = 4
        For A = 0 To 255
            Get #1, , LogPalette.palPalEntry(A)
        Next A
        Close #1
        Paletted = True
    Else
        Paletted = False
    End If
    
    hdcTiles = CreateCompatibleDC(0&)
    Set Pic = LoadPicture("tiles.rsc")
    hbmpTiles = Pic.Handle
    obmpTiles = SelectObject(hdcTiles, hbmpTiles)
    
    hdcTilesMask = CreateCompatibleDC(0&)
    Set Pic = LoadPicture("tilesm.rsc")
    hbmpTilesMask = Pic.Handle
    obmpTilesMask = SelectObject(hdcTilesMask, hbmpTilesMask)
    
    Load frmLogin
    
    'Hook Form
    Hook
    
    'Load Winsock
    StartWinsock (St)
    frmLogin.Show
End Sub

Sub CheckFile(FileName As String)
    If Exists(FileName) = False Then
        MsgBox "Error: File " + Chr$(34) + FileName + Chr$(34) + " not found!", vbOKOnly + vbExclamation
        End
    End If
End Sub

Function Exists(FileName As String) As Boolean
     Exists = (Dir(FileName) <> "")
End Function

Sub CreateMapCache()
    Dim St1 As String * 2677, A As Long
    St1 = String$(2677, 0)
    Open "cache1.dat" For Random As #1 Len = 2677
    For A = 1 To 3000
        Put #1, , St1
    Next A
    Close #1
End Sub

Sub DeInitialize()
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
    
    If hPalette <> 0 Then
        DeleteObject hPalette
    End If
    
    'Unload Winsock
    EndWinsock
    
    'Unhook Form
    Unhook
    
    DoEvents
    
    End
End Sub

Public Sub Hook()
    lpPrevWndProc = SetWindowLong(gHW, GWL_WNDPROC, AddressOf WindowProc)
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
            Case FD_CONNECT
                If lParam = FD_CONNECT Then
                    SendSocket Chr$(61) + Chr$(frmLogin.txtVersion) + Chr$(CheckSum("minimapper") Mod 256) + "minimapper"
                    SendSocket Chr$(1) + User + Chr$(0) + Pass
                End If
            Case FD_READ
                If lParam = FD_READ Then ReceiveData
        End Select
    End If
    WindowProc = CallWindowProc(lpPrevWndProc, hw, uMsg, wParam, lParam)
End Function

Sub ReceiveData()
    On Error Resume Next
    Dim PacketLength As Integer, PacketID As Integer, PacketChecksum As Integer, ServPacketOrder As Integer
    Dim St As String, St1 As String
    Dim A As Long, B As Long, C As Long, D As Long
    
    SocketData = SocketData + Receive(ClientSocket)
LoopRead:
    If Len(SocketData) >= 5 Then
        PacketLength = GetInt(Mid$(SocketData, 1, 2))
        PacketChecksum = Asc(Mid$(SocketData, 3, 1))
        ServPacketOrder = Asc(Mid$(SocketData, 4, 1))
        If Len(SocketData) - 4 >= PacketLength Then
            St = Mid$(SocketData, 5, PacketLength)
            SocketData = Mid$(SocketData, PacketLength + 5)
            
            If PacketLength > 0 Then
                PacketID = Asc(Mid$(St, 1, 1))
                If Len(St) > 1 Then
                    St = Mid$(St, 2)
                Else
                    St = ""
                End If
                Select Case PacketID
                    Case 0 'Error Logging On
                        If Len(St) >= 1 Then
                            Select Case Asc(Mid$(St, 1, 1))
                                Case 0 'Custom Message
                                    If Len(St) >= 2 Then
                                        MsgBox Mid$(St, 2), vbOKOnly + vbExclamation
                                    End If
                                Case 1 'Invalid User/Pass
                                    MsgBox "Invalid user name/password!", vbOKOnly + vbExclamation
                                Case 2 'Account already in use
                                    MsgBox "Someone is already using that account!", vbOKOnly + vbExclamation
                                Case 3 'Banned
                                    If Len(St) >= 5 Then
                                        A = Asc(Mid$(St, 2, 1)) * 16777216 + Asc(Mid$(St, 3, 1)) * 65536 + Asc(Mid$(St, 4, 1)) * 256& + Asc(Mid$(St, 5, 1))
                                        If Len(St) > 5 Then
                                            MsgBox "You are banned from The Odyssey Online Classic until " + CStr(CDate(A)) + " (" + Mid$(St, 6) + ")!", vbOKOnly
                                        Else
                                            MsgBox "You are banned from The Odyssey Online Classic until " + CStr(CDate(A)) + "!", vbOKOnly
                                        End If
                                        CloseClientSocket 3
                                    End If
                                Case 4 'Server Full
                                    MsgBox "The server is full, please try again in a few minutes!", vbOKOnly + vbExclamation
                                Case 5 'Multiple Login
                                    MsgBox "You may not log in multiple times from the same computer!", vbOKOnly + vbExclamation
                            End Select
                        End If
                        CloseClientSocket 1
            
                    Case 3 'Logged On / Character Data
                        SendSocket Chr$(23)
                        frmMain.Show
                        frmLogin.Hide
                    
                    Case 10 'Player moved
                    
                    Case 12 'Joined map
                    
                    Case 21 'Map Data
                        If CMap > 0 Then
                            RequestedMap = False

                            MapData = UncompressString$(St)
                            
                            MapDataArray = StrConv(MapData, vbFromUnicode)
                            ReDim Preserve MapDataArray(UBound(MapDataArray) + 1)
                            EncryptDataString MapDataArray(0), CMap * 16 Mod 50 + 5
                            MapData = StrConv(MapDataArray, vbUnicode)
                    
                            On Error Resume Next
                            Close #1
                            On Error GoTo 0
                            
                            Open "cache1.dat" For Random As #1 Len = 2677
                            Put #1, CMap, MapData
                            Close #1
                            
                            LoadMapFromCache CMap
                            If RequestedMap = False Then
                                TheMap(CMap).Received = True
                                ReceivedMap = True
                            End If
                        End If
                        
                    Case 26 'Broadcast
                        
                    Case 130 'Stat Update
                        
                    Case 170 'Raw Data
                    
                    Case 171 'Temp - remove later
                    
                    Case Else
                        'MsgBox PacketID
                End Select
            End If
            GoTo LoopRead
        End If
    End If
End Sub

Sub LoadMapFromCache(LoadMap As Long)
On Error Resume Next
    Close #1
On Error GoTo LoadError

    Open "cache1.dat" For Random As #1 Len = 2677
    Get #1, LoadMap, MapData
    Close #1
    
    If Asc(Mid$(MapData, 1, 1)) > 0 Then
        MapDataArray = StrConv(MapData, vbFromUnicode)
        ReDim Preserve MapDataArray(UBound(MapDataArray) + 1)
        EncryptDataString MapDataArray(0), LoadMap * 16 Mod 50 + 5
        MapData = StrConv(MapDataArray, vbUnicode)
    End If
    
    LoadMapData MapData
    
    Exit Sub
    
LoadError:
    RequestedMap = True
    SendSocket Chr$(45)
End Sub
Sub LoadMapData(LoadMapData As String)
    On Error GoTo LoadError

    Dim A As Long, X As Long, Y As Long
    If Len(LoadMapData) = 2677 Then
        MapData = LoadMapData
        With Map
            .Name = ClipString$(Mid$(MapData, 1, 30))
            .Version = Asc(Mid$(MapData, 31, 1)) * 16777216 + Asc(Mid$(MapData, 32, 1)) * 65536 + Asc(Mid$(MapData, 33, 1)) * 256& + Asc(Mid$(MapData, 34, 1))
            .NPC = Asc(Mid$(MapData, 35, 1)) * 256 + Asc(Mid$(MapData, 36, 1))
            .MIDI = Asc(Mid$(MapData, 37, 1))
            .ExitUp = Asc(Mid$(MapData, 38, 1)) * 256 + Asc(Mid$(MapData, 39, 1))
            .ExitDown = Asc(Mid$(MapData, 40, 1)) * 256 + Asc(Mid$(MapData, 41, 1))
            .ExitLeft = Asc(Mid$(MapData, 42, 1)) * 256 + Asc(Mid$(MapData, 43, 1))
            .ExitRight = Asc(Mid$(MapData, 44, 1)) * 256 + Asc(Mid$(MapData, 45, 1))
            .BootLocation.Map = Asc(Mid$(MapData, 46, 1)) * 256 + Asc(Mid$(MapData, 47, 1))
            .BootLocation.X = Asc(Mid$(MapData, 48, 1))
            .BootLocation.Y = Asc(Mid$(MapData, 49, 1))
            .DeathLocation.Map = Asc(Mid$(MapData, 50, 1)) * 256 + Asc(Mid$(MapData, 51, 1))
            .DeathLocation.X = Asc(Mid$(MapData, 52, 1))
            .DeathLocation.Y = Asc(Mid$(MapData, 53, 1))
            .flags = Asc(Mid$(MapData, 54, 1))
            .Flags2 = Asc(Mid$(MapData, 55, 1))
            For A = 0 To 9    '56 - 86
                .MonsterSpawn(A).Monster = Asc(Mid$(MapData, 56 + A * 3)) * 256 + Asc(Mid$(MapData, 57 + A * 3))
                .MonsterSpawn(A).Rate = Asc(Mid$(MapData, 58 + A * 3))
            Next A
            '86
            For Y = 0 To 11
                For X = 0 To 11
                    With .Tile(X, Y)
                        A = 86 + Y * 216 + X * 18
                        '1-10 = Tiles
                        .Ground = Asc(Mid$(MapData, A, 1)) * 256 + Asc(Mid$(MapData, A + 1, 1))
                        .Ground2 = Asc(Mid$(MapData, A + 2, 1)) * 256 + Asc(Mid$(MapData, A + 3, 1))
                        .BGTile1 = Asc(Mid$(MapData, A + 4, 1)) * 256 + Asc(Mid$(MapData, A + 5, 1))
                        .BGTile2 = Asc(Mid$(MapData, A + 6, 1)) * 256 + Asc(Mid$(MapData, A + 7, 1))
                        .FGTile = Asc(Mid$(MapData, A + 8, 1)) * 256 + Asc(Mid$(MapData, A + 9, 1))
                        '.FGTile2 = Asc(Mid$(MapData, A + 10, 1)) * 256 + Asc(Mid$(MapData, A + 11, 1))
                        .Att = Asc(Mid$(MapData, A + 12, 1))
                        .AttData(0) = Asc(Mid$(MapData, A + 13, 1))
                        .AttData(1) = Asc(Mid$(MapData, A + 14, 1))
                        .AttData(2) = Asc(Mid$(MapData, A + 15, 1))
                        .AttData(3) = Asc(Mid$(MapData, A + 16, 1))
                        .Att2 = Asc(Mid$(MapData, A + 17, 1))
                    End With
                Next X
            Next Y
        End With
    End If

    Exit Sub
    
LoadError:
    RequestedMap = True
    SendSocket Chr$(45)
End Sub

Sub SaveFinalMap()
    Dim I As Integer, J As Integer
    Dim X As Integer, Y As Integer
    
    For X = 0 To 24
        For Y = 0 To 24
            For I = 0 To 3000
                If TheMap(I).X / MapSize = X And TheMap(I).Y / MapSize = Y Then
                    MapArray(X, Y) = I
                End If
            Next I
        Next Y
    Next X
    
    For I = 0 To 12
        For J = 0 To 12
            If MapArray(I, J) > 0 Then
                LoadMapFromCache CLng(MapArray(I, J))
                For X = 0 To 11
                    For Y = 0 To 11
                        If I * 12 + X > 125 Then
                            If J * 12 + Y > 125 Then
                                With Map.Tile(X, Y)
                                    TileDataArray4(I * 12 + X, J * 12 + Y).Ground = .Ground
                                    TileDataArray4(I * 12 + X, J * 12 + Y).Ground2 = .Ground2
                                    TileDataArray4(I * 12 + X, J * 12 + Y).BGTile1 = .BGTile1
                                    TileDataArray4(I * 12 + X, J * 12 + Y).BGTile2 = .BGTile2
                                    TileDataArray4(I * 12 + X, J * 12 + Y).FGTile = .FGTile
                                    TileDataArray4(I * 12 + X, J * 12 + Y).Att = .Att
                                    TileDataArray4(I * 12 + X, J * 12 + Y).AttData(0) = .AttData(0)
                                    TileDataArray4(I * 12 + X, J * 12 + Y).AttData(1) = .AttData(1)
                                    TileDataArray4(I * 12 + X, J * 12 + Y).AttData(2) = .AttData(2)
                                    TileDataArray4(I * 12 + X, J * 12 + Y).AttData(3) = .AttData(3)
                                    TileDataArray4(I * 12 + X, J * 12 + Y).Att2 = .Att2
                                End With
                            Else
                                With Map.Tile(X, Y)
                                    TileDataArray3(I * 12 + X, J * 12 + Y).Ground = .Ground
                                    TileDataArray3(I * 12 + X, J * 12 + Y).Ground2 = .Ground2
                                    TileDataArray3(I * 12 + X, J * 12 + Y).BGTile1 = .BGTile1
                                    TileDataArray3(I * 12 + X, J * 12 + Y).BGTile2 = .BGTile2
                                    TileDataArray3(I * 12 + X, J * 12 + Y).FGTile = .FGTile
                                    TileDataArray3(I * 12 + X, J * 12 + Y).Att = .Att
                                    TileDataArray3(I * 12 + X, J * 12 + Y).AttData(0) = .AttData(0)
                                    TileDataArray3(I * 12 + X, J * 12 + Y).AttData(1) = .AttData(1)
                                    TileDataArray3(I * 12 + X, J * 12 + Y).AttData(2) = .AttData(2)
                                    TileDataArray3(I * 12 + X, J * 12 + Y).AttData(3) = .AttData(3)
                                    TileDataArray3(I * 12 + X, J * 12 + Y).Att2 = .Att2
                                End With
                            End If
                        Else
                            If J * 12 + Y > 125 Then
                                With Map.Tile(X, Y)
                                    TileDataArray2(I * 12 + X, J * 12 + Y).Ground = .Ground
                                    TileDataArray2(I * 12 + X, J * 12 + Y).Ground2 = .Ground2
                                    TileDataArray2(I * 12 + X, J * 12 + Y).BGTile1 = .BGTile1
                                    TileDataArray2(I * 12 + X, J * 12 + Y).BGTile2 = .BGTile2
                                    TileDataArray2(I * 12 + X, J * 12 + Y).FGTile = .FGTile
                                    TileDataArray2(I * 12 + X, J * 12 + Y).Att = .Att
                                    TileDataArray2(I * 12 + X, J * 12 + Y).AttData(0) = .AttData(0)
                                    TileDataArray2(I * 12 + X, J * 12 + Y).AttData(1) = .AttData(1)
                                    TileDataArray2(I * 12 + X, J * 12 + Y).AttData(2) = .AttData(2)
                                    TileDataArray2(I * 12 + X, J * 12 + Y).AttData(3) = .AttData(3)
                                    TileDataArray2(I * 12 + X, J * 12 + Y).Att2 = .Att2
                                End With
                            Else
                                With Map.Tile(X, Y)
                                    TileDataArray(I * 12 + X, J * 12 + Y).Ground = .Ground
                                    TileDataArray(I * 12 + X, J * 12 + Y).Ground2 = .Ground2
                                    TileDataArray(I * 12 + X, J * 12 + Y).BGTile1 = .BGTile1
                                    TileDataArray(I * 12 + X, J * 12 + Y).BGTile2 = .BGTile2
                                    TileDataArray(I * 12 + X, J * 12 + Y).FGTile = .FGTile
                                    TileDataArray(I * 12 + X, J * 12 + Y).Att = .Att
                                    TileDataArray(I * 12 + X, J * 12 + Y).AttData(0) = .AttData(0)
                                    TileDataArray(I * 12 + X, J * 12 + Y).AttData(1) = .AttData(1)
                                    TileDataArray(I * 12 + X, J * 12 + Y).AttData(2) = .AttData(2)
                                    TileDataArray(I * 12 + X, J * 12 + Y).AttData(3) = .AttData(3)
                                    TileDataArray(I * 12 + X, J * 12 + Y).Att2 = .Att2
                                End With
                            End If
                        End If
                    Next Y
                Next X
            End If
        Next J
    Next I
    
    SaveMap
End Sub

Sub SaveMap()
    Dim MapData As String, St1 As String * 30
    Dim X As Long, Y As Long
    With Map
        If .Version < 2147483647 Then
            .Version = .Version + 1
        Else
            .Version = 1
        End If
        St1 = .Name
        MapData = St1 + QuadChar(.Version) + Chr$(.NPC) + Chr$(.MIDI) + DoubleChar$(CLng(.ExitUp)) + DoubleChar$(CLng(.ExitDown)) + DoubleChar$(CLng(.ExitLeft)) + DoubleChar$(CLng(.ExitRight)) + DoubleChar(CLng(.BootLocation.Map)) + Chr$(.BootLocation.X) + Chr$(.BootLocation.Y) + Chr$(.flags) + Chr$(.MonsterSpawn(0).Monster) + Chr$(.MonsterSpawn(0).Rate) + Chr$(.MonsterSpawn(1).Monster) + Chr$(.MonsterSpawn(1).Rate) + Chr$(.MonsterSpawn(2).Monster) + Chr$(.MonsterSpawn(2).Rate)
        For Y = 0 To 200
            For X = 0 To 200
                If X > 125 Then
                    If Y > 125 Then
                        With TileDataArray4(X, Y)
                            MapData = MapData + DoubleChar(CLng(.Ground)) + DoubleChar(CLng(.Ground2)) + DoubleChar(CLng(.BGTile1)) + DoubleChar(CLng(.BGTile2)) + DoubleChar(CLng(.FGTile)) + Chr$(.Att) + Chr$(.AttData(0)) + Chr$(.AttData(1)) + Chr$(.AttData(2)) + Chr$(.AttData(3)) + Chr$(.Att2)
                        End With
                    Else
                        With TileDataArray3(X, Y)
                            MapData = MapData + DoubleChar(CLng(.Ground)) + DoubleChar(CLng(.Ground2)) + DoubleChar(CLng(.BGTile1)) + DoubleChar(CLng(.BGTile2)) + DoubleChar(CLng(.FGTile)) + Chr$(.Att) + Chr$(.AttData(0)) + Chr$(.AttData(1)) + Chr$(.AttData(2)) + Chr$(.AttData(3)) + Chr$(.Att2)
                        End With
                    End If
                Else
                    If Y > 125 Then
                        With TileDataArray2(X, Y)
                            MapData = MapData + DoubleChar(CLng(.Ground)) + DoubleChar(CLng(.Ground2)) + DoubleChar(CLng(.BGTile1)) + DoubleChar(CLng(.BGTile2)) + DoubleChar(CLng(.FGTile)) + Chr$(.Att) + Chr$(.AttData(0)) + Chr$(.AttData(1)) + Chr$(.AttData(2)) + Chr$(.AttData(3)) + Chr$(.Att2)
                        End With
                    Else
                        With TileDataArray(X, Y)
                            MapData = MapData + DoubleChar(CLng(.Ground)) + DoubleChar(CLng(.Ground2)) + DoubleChar(CLng(.BGTile1)) + DoubleChar(CLng(.BGTile2)) + DoubleChar(CLng(.FGTile)) + Chr$(.Att) + Chr$(.AttData(0)) + Chr$(.AttData(1)) + Chr$(.AttData(2)) + Chr$(.AttData(3)) + Chr$(.Att2)
                        End With
                    End If
                End If
            Next X
        Next Y
    End With
    
    Open "merples" For Output As #1
    Print #1, MapData
    Close #1
End Sub

Function QuadChar(Num As Long) As String
    QuadChar = Chr$(Int(Num / 16777216) Mod 256) + Chr$(Int(Num / 65536) Mod 256) + Chr$(Int(Num / 256) Mod 256) + Chr$(Num Mod 256)
End Function

Sub CloseClientSocket(Action As Byte)
    ShutDown ClientSocket, 2
    closesocket ClientSocket
    ClientSocket = INVALID_SOCKET
    
    frmLogin.btnOk.Enabled = True
End Sub

Function CheckSum(St As String) As Long
    Dim A As Long, B As Long
    For A = 1 To Len(St)
        B = B + Asc(Mid$(St, A, 1))
    Next A
    CheckSum = B
End Function

Sub SendSocket(ByVal St As String)
    If SendData(ClientSocket, DoubleChar$(Len(St)) + Chr$(CheckSum(St) * 20 Mod 194) + Chr$(PacketOrder) + St) = SOCKET_ERROR Then
        'CloseClientSocket 0
    End If
    PacketOrder = PacketOrder + 1
    If PacketOrder > 250 Then PacketOrder = 0
End Sub

Function GetInt(Chars As String) As Long
    GetInt = Asc(Mid$(Chars, 1, 1)) * 256 + Asc(Mid$(Chars, 2, 1))
End Function

Function DoubleChar(Num As Long) As String
    DoubleChar = Chr$(Int(Num / 256)) + Chr$(Num Mod 256)
End Function

Function ClipString(St As String) As String
    Dim A As Long
    For A = Len(St) To 1 Step -1
        If Mid$(St, A, 1) <> Chr$(32) Then
            ClipString = Mid$(St, 1, A)
            Exit Function
        End If
    Next A
End Function
