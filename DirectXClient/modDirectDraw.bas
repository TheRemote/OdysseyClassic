Attribute VB_Name = "modDirectDraw"
Option Explicit

Public DDraw As DirectDraw4

Public DDCK As DDCOLORKEY

'Surfaces
Public BackBufferSurf As DirectDrawSurface4
Public PrimarySurf As DirectDrawSurface4

Public BGTile1Buffer As DirectDrawSurface4
Public BGTile2Buffer As DirectDrawSurface4
Public FGTileBuffer As DirectDrawSurface4
Public FGTile2Buffer As DirectDrawSurface4
Public InventoryBuffer As DirectDrawSurface4

Public DDSTiles As DirectDrawSurface4
Public DDSObjects As DirectDrawSurface4
Public DDSSprites As DirectDrawSurface4
Public DDSEffects As DirectDrawSurface4
Public DDSAtts As DirectDrawSurface4
Public DDSHPBar As DirectDrawSurface4
Public DDSInventory As DirectDrawSurface4
Public DDSInterfaceLights As DirectDrawSurface4
Public DDSStats As DirectDrawSurface4

Public PrimaryClip As DirectDrawClipper

Public FontInfo As New StdFont

Public DDSDPrimary As DDSURFACEDESC2
Public DDSDBackBuffer As DDSURFACEDESC2

Public BackBufferRect As RECT
Public MapRect As RECT
Public FullMapRect As RECT
Public InventoryRect As RECT
Public EmptyRect As RECT
Public SrcRect As RECT
Public DestRect As RECT

Public ddsBufferArray() As Byte

Public Initialized As Boolean
Public RestoreDirectDraw As Boolean

Public LastDrawReturn As Long

Public Sub InitDirectDraw()
    DDraw.SetCooperativeLevel frmMain.hwnd, DDSCL_NORMAL

    'Create Primary Surface
    With DDSDPrimary
        .lFlags = DDSD_CAPS
        .ddsCaps.lCaps = DDSCAPS_PRIMARYSURFACE
    End With

    Set PrimarySurf = Nothing
    Set PrimarySurf = DDraw.CreateSurface(DDSDPrimary)

    Set PrimaryClip = Nothing
    Set PrimaryClip = DDraw.CreateClipper(0)

    PrimaryClip.SetHWnd frmMain.picViewport.hwnd
    PrimarySurf.SetClipper PrimaryClip

    FullMapRect.Top = 0
    FullMapRect.Bottom = 384
    FullMapRect.Left = 0
    FullMapRect.Right = 384

    InventoryRect.Left = 0
    InventoryRect.Top = 0
    InventoryRect.Right = 181
    InventoryRect.Bottom = 181

    LoadSurfaces
End Sub

Public Sub LoadSurfaces()
'Create Back Buffer Surface
    With BackBufferRect
        .Left = 0
        .Top = 0
        .Right = 384
        .Bottom = 384
    End With

    With DDSDBackBuffer
        .lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
        .ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
        .lHeight = BackBufferRect.Bottom
        .lWidth = BackBufferRect.Right
    End With

    Set BackBufferSurf = Nothing
    Set BackBufferSurf = DDraw.CreateSurface(DDSDBackBuffer)
    Set BGTile1Buffer = Nothing
    Set BGTile1Buffer = DDraw.CreateSurface(DDSDBackBuffer)
    Set BGTile2Buffer = Nothing
    Set BGTile2Buffer = DDraw.CreateSurface(DDSDBackBuffer)
    Set FGTileBuffer = Nothing
    Set FGTileBuffer = DDraw.CreateSurface(DDSDBackBuffer)
    Set FGTile2Buffer = Nothing
    Set FGTile2Buffer = DDraw.CreateSurface(DDSDBackBuffer)

    BackBufferSurf.GetSurfaceDesc DDSDBackBuffer

    If DDSDBackBuffer.ddpfPixelFormat.lRGBBitCount = 16 Then
        options.Bit32 = False
    ElseIf DDSDBackBuffer.ddpfPixelFormat.lRGBBitCount = 32 Then
        options.Bit32 = True
    Else
        If options.DisableLighting = False Then
            MsgBox "Odyssey's lighting and weather effects require either 16 or 32 bit color display mode.  These effects have been disabled on your machine!"
            options.DisableLighting = True
            SaveOptions
        End If
    End If

    Set InventoryBuffer = Nothing
    Set InventoryBuffer = DDraw.CreateSurface(DDSDBackBuffer)

    'Set color Key
    DDCK.low = 0
    DDCK.high = 0
    BackBufferSurf.SetColorKey DDCKEY_SRCBLT, DDCK
    BGTile1Buffer.SetColorKey DDCKEY_SRCBLT, DDCK
    BGTile2Buffer.SetColorKey DDCKEY_SRCBLT, DDCK
    FGTileBuffer.SetColorKey DDCKEY_SRCBLT, DDCK
    FGTile2Buffer.SetColorKey DDCKEY_SRCBLT, DDCK
    InventoryBuffer.SetColorKey DDCKEY_SRCBLT, DDCK

    LoadProtectedSurface DDSEffects, App.Path + "\effects.rsc"
    LoadProtectedSurface DDSObjects, App.Path + "\objects.rsc"
    LoadProtectedSurface DDSSprites, App.Path + "\sprites.rsc"
    LoadProtectedSurface DDSTiles, App.Path + "\tiles.rsc"
    LoadProtectedSurface DDSAtts, App.Path + "\atts.rsc"
    LoadProtectedSurface DDSHPBar, App.Path + "\hpbar.rsc"
    LoadProtectedSurface DDSInventory, App.Path + "\inventory.rsc"
    LoadProtectedSurface DDSInterfaceLights, App.Path + "\InterfaceLights.rsc"
    LoadProtectedSurface DDSStats, App.Path + "\stats.rsc"

    Call BackBufferSurf.BltColorFill(FullMapRect, RGB(0, 0, 0))
    Call BGTile1Buffer.BltColorFill(FullMapRect, RGB(0, 0, 0))
    Call BGTile2Buffer.BltColorFill(FullMapRect, RGB(0, 0, 0))
    Call FGTileBuffer.BltColorFill(FullMapRect, RGB(0, 0, 0))
    Call FGTile2Buffer.BltColorFill(FullMapRect, RGB(0, 0, 0))
    Call InventoryBuffer.BltColorFill(InventoryRect, RGB(0, 0, 0))

    FontInfo.Bold = False
    FontInfo.Size = 11
    FontInfo.name = "System"
    BackBufferSurf.SetFontTransparency True
    BackBufferSurf.SetFont FontInfo
End Sub

Public Function UnloadDirectDraw()
    Set BackBufferSurf = Nothing
    Set PrimarySurf = Nothing

    Set BGTile1Buffer = Nothing
    Set BGTile2Buffer = Nothing
    Set FGTileBuffer = Nothing
    Set FGTile2Buffer = Nothing
    Set InventoryBuffer = Nothing

    Set DDSTiles = Nothing
    Set DDSObjects = Nothing
    Set DDSSprites = Nothing
    Set DDSEffects = Nothing
    Set DDSAtts = Nothing
    Set DDSHPBar = Nothing
    Set DDSInventory = Nothing
    Set DDSInterfaceLights = Nothing
    Set DDSStats = Nothing

    Set PrimaryClip = Nothing

    Set FontInfo = Nothing
End Function

Public Sub LoadSurface(Surface As DirectDrawSurface4, File As String)
    With DDSDBackBuffer
        .lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
        .ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
    End With
    GetBitmapDimensions File, DDSDBackBuffer.lWidth, DDSDBackBuffer.lHeight
    Set Surface = Nothing
    Set Surface = DDraw.CreateSurfaceFromFile(File, DDSDBackBuffer)
    DDCK.low = 0
    DDCK.high = 0
    Surface.SetColorKey DDCKEY_SRCBLT, DDCK
End Sub

Public Sub LoadProtectedSurface(Surface As DirectDrawSurface4, File As String)
    If Exists(File) Then
        Dim FileByteArray() As Byte

        FileByteArray() = StrConv(File, vbFromUnicode)
        ReDim Preserve FileByteArray(UBound(FileByteArray) + 1)

        EncryptDataFile FileByteArray(0), FileLen(File) Mod 87 + 5

        LoadSurface Surface, File

        EncryptDataFile FileByteArray(0), FileLen(File) Mod 87 + 5
    End If
End Sub

Function GetBitmapDimensions(File As String, ByRef bmWidth As Long, ByRef bmHeight As Long)
'Gets the dimensions of a file
    Dim BMHeader As BITMAPFILEHEADER
    Dim BINFOHeader As BITMAPINFOHEADER2

    Open File For Binary Access Read As #1
    Get #1, , BMHeader
    Get #1, , BINFOHeader
    Close #1

    bmWidth = BINFOHeader.biWidth
    bmHeight = BINFOHeader.biHeight
End Function

Public Sub Draw(X As Long, Y As Long, Width As Integer, Height As Integer, Surface As DirectDrawSurface4, SrcX As Integer, SrcY As Integer, Transparent As Boolean)
    Dim DrawRect As RECT
    DrawRect.Left = SrcX
    DrawRect.Top = SrcY
    DrawRect.Right = SrcX + Width
    DrawRect.Bottom = SrcY + Height

    If Y < 0 Then
        DrawRect.Top = DrawRect.Top - Y
        Y = 0
    End If

    If Transparent = True Then
        Call BackBufferSurf.BltFast(X, Y, Surface, DrawRect, DDBLTFAST_SRCCOLORKEY)
    Else
        Call BackBufferSurf.BltFast(X, Y, Surface, DrawRect, DDBLTFAST_NOCOLORKEY)
    End If
End Sub

Public Sub FillRect(X As Long, Y As Long, Width As Long, Height As Long, Surface As DirectDrawSurface4, Color As Long)
    Dim DrawRect As RECT
    DrawRect.Left = X
    DrawRect.Top = Y
    DrawRect.Right = X + Width
    DrawRect.Bottom = Y + Height

    If Y < 0 Then
        DrawRect.Top = DrawRect.Top - Y
        Y = 0
    End If

    Call Surface.BltColorFill(DrawRect, Color)
End Sub

Public Sub RestoreSurfaces()
    On Error Resume Next
    DDraw.RestoreAllSurfaces
End Sub

Sub DrawToDC(X As Long, Y As Long, Width As Integer, Height As Integer, hDC As Long, Surface As DirectDrawSurface4, SrcX As Integer, SrcY As Integer)
    SrcRect.Left = SrcX
    SrcRect.Top = SrcY
    SrcRect.Right = SrcRect.Left + Width
    SrcRect.Bottom = SrcRect.Top + Height
    DestRect.Left = X
    DestRect.Top = Y
    DestRect.Right = DestRect.Left + Width
    DestRect.Bottom = DestRect.Top + Height
    Call Surface.BltToDC(hDC, SrcRect, DestRect)
End Sub
