VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Minimapper"
   ClientHeight    =   9540
   ClientLeft      =   390
   ClientTop       =   885
   ClientWidth     =   12195
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   636
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   813
   Begin VB.Timer TimeUpdateAll 
      Left            =   11280
      Top             =   2040
   End
   Begin VB.Timer TimeUpdate 
      Left            =   11280
      Top             =   1560
   End
   Begin VB.Timer TimePing 
      Interval        =   25000
      Left            =   11280
      Top             =   1080
   End
   Begin VB.Timer TimeRequestAll 
      Left            =   11280
      Top             =   600
   End
   Begin VB.PictureBox picSave 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   120
      ScaleHeight     =   12
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   12
      TabIndex        =   2
      Top             =   240
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.PictureBox picView 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   120
      ScaleHeight     =   12
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   12
      TabIndex        =   0
      Top             =   240
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.Timer TimeRequest 
      Left            =   11280
      Top             =   120
   End
   Begin VB.Label lblDetails 
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   4215
   End
   Begin VB.Menu cmdRequestArea 
      Caption         =   "Request Area"
   End
   Begin VB.Menu cmdRequestAll 
      Caption         =   "Request All Maps"
   End
   Begin VB.Menu menuSaveMap 
      Caption         =   "Save Map"
   End
   Begin VB.Menu cmdMapSize 
      Caption         =   "Redraw Map"
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdMapSize_Click()
    Dim A As Long
    Dim B As Long
    A = CLng(Val(InputBox("How many pixels wide do you want each map?", , 48)))
    B = CLng(Val(InputBox("Start from which map?", , 1)))
    If A > 0 And B > 0 Then
        MapSize = A
        DrawItAll B
    End If
End Sub

Private Sub cmdRequestAll_Click()
    SendSocket Chr$(4) + DoubleChar$(1)
    ReceivedMap = False
    TimeRequestAll.Interval = 1
    TimeRequest.Interval = 0
End Sub

Private Sub cmdRequestArea_Click()
    Dim A As Long
    Dim B As Long
    A = CLng(Val(InputBox("How many pixels wide do you want each map?", , 48)))
    B = CLng(Val(InputBox("Start from which map?", , 1)))
    If A > 0 And B > 0 Then
        MapSize = A
        CMap = B
        OriginalStartLoc = CMap
        For A = 1 To 3000
            TheMap(A).Received = False
        Next A
        TimeRequest.Interval = 1
        SendSocket Chr$(4) + DoubleChar$(CMap)
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 38 'Up
            Form_KeyPress 38
        Case 40 'Down
            Form_KeyPress 40
        Case 37 'Left
            Form_KeyPress 37
        Case 39 'Right
            Form_KeyPress 39
    End Select
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 38 'Up
            YOffset = YOffset + MapSize
            DrawView
        Case 40 'Down
            YOffset = YOffset - MapSize
            DrawView
        Case 37 'Left
            XOffset = XOffset + MapSize
            DrawView
        Case 39 'Right
            XOffset = XOffset - MapSize
            DrawView
    End Select
End Sub

Private Sub Form_Load()
    SetStretchBltMode picView.hdc, vbPaletteModeNone
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmLogin.TimeDeinitialize.Interval = 50
End Sub

Private Sub menuSaveMap_Click()
    SaveFinalMap
End Sub

Private Sub TimePing_Timer()
    SendSocket Chr$(29)
End Sub

Private Sub TimeRequest_Timer()
    Dim FoundMap As Boolean
    If ReceivedMap = True Then
        Do Until FoundMap = True
            frmMain.lblDetails = "Now loading map " & CMap
            frmMain.Refresh
            LoadMapFromCache CMap
            If Map.Version = 0 Then 'Found Map
                FoundMap = True
                frmMain.lblDetails = "Requesting Map " & CMap
                ReceivedMap = False
                frmMain.Refresh
                SendSocket Chr$(4) + DoubleChar$(CMap)
            Else
                TheMap(CMap).Received = True
                If Map.ExitRight > 0 And TheMap(Map.ExitRight).Received = False Then
                    CMap = Map.ExitRight
                    FoundMap = True
                ElseIf Map.ExitDown > 0 And TheMap(Map.ExitDown).Received = False Then
                    CMap = Map.ExitDown
                    FoundMap = True
                ElseIf Map.ExitUp > 0 And TheMap(Map.ExitUp).Received = False Then
                    CMap = Map.ExitUp
                    FoundMap = True
                ElseIf Map.ExitLeft > 0 And TheMap(Map.ExitLeft).Received = False Then
                    CMap = Map.ExitLeft
                    FoundMap = True
                Else
                    CMap = FindReceivableMap
                    If CMap = 0 Then
                        FoundMap = True
                        TimeRequest.Interval = 0
                        DrawItAll OriginalStartLoc
                    Else
                        FoundMap = True
                    End If
                End If
            End If
        Loop
    End If
End Sub

Sub DrawMap(DrawX As Long, DrawY As Long)
    Dim A As Long, B As Long, X As Long, Y As Long
    BitBlt hdcBack(0), 0, 0, 384, 384, 0, 0, 0, BLACKNESS
    BitBlt hdcBack(1), 0, 0, 384, 384, 0, 0, 0, BLACKNESS
    BitBlt hdcFront, 0, 0, 384, 384, 0, 0, 0, BLACKNESS
    BitBlt hdcFrontMask, 0, 0, 384, 384, 0, 0, 0, WHITENESS
    
    For X = 0 To 11
        For Y = 0 To 11
            With Map.Tile(X, Y)
                If .Ground > 0 Then
                    BitBlt hdcBack(0), X * 32, Y * 32, 32, 32, hdcTiles, ((.Ground - 1) Mod 7) * 32, Int((.Ground - 1) / 7) * 32, SRCCOPY
                    BitBlt hdcBack(1), X * 32, Y * 32, 32, 32, hdcTiles, ((.Ground - 1) Mod 7) * 32, Int((.Ground - 1) / 7) * 32, SRCCOPY
                End If
                If .Ground2 > 0 Then
                    TransparentBlt hdcBack(0), X * 32, Y * 32, 32, 32, hdcTiles, ((.Ground2 - 1) Mod 7) * 32, Int((.Ground2 - 1) / 7) * 32, SRCCOPY
                    TransparentBlt hdcBack(1), X * 32, Y * 32, 32, 32, hdcTiles, ((.Ground2 - 1) Mod 7) * 32, Int((.Ground2 - 1) / 7) * 32, SRCCOPY
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
    
    BitBlt hdcBuffer, 0, 0, 384, 384, hdcBack(0), 0, 0, SRCCOPY
    TransparentBlt hdcBuffer, 0, 0, 384, 384, hdcFront, 0, 0, hdcFrontMask
    SetStretchBltMode picView.hdc, vbPaletteModeNone
    StretchBlt picView.hdc, DrawX, DrawY, MapSize, MapSize, hdcBuffer, 0, 0, 384, 384, vbSrcCopy
End Sub

Sub TransparentBlt(hdc As Long, ByVal destX As Long, ByVal destY As Long, destWidth As Long, destHeight As Long, srcDC As Long, srcX As Long, srcY As Long, maskDC As Long)
    BitBlt hdc, destX, destY, destWidth, destHeight, maskDC, srcX, srcY, SRCAND
    BitBlt hdc, destX, destY, destWidth, destHeight, srcDC, srcX, srcY, SRCPAINT
End Sub

Sub DrawItAll(StartLoc As Long)
    picView.Cls
    picSave.Cls
    TimeRequest.Interval = 0
    TimeRequestAll.Interval = 0
    CMap = StartLoc
    picView.Visible = True
    picSave.Width = MapSize
    picSave.Height = MapSize
    picView.Width = MapSize
    picView.Height = MapSize
    Dim A As Long
    For A = 1 To 3000
        TheMap(A).Drawn = False
        TheMap(A).X = 0
        TheMap(A).Y = 0
    Next A
    
    Dim CX As Long, CY As Long
    XOffset = 0
    YOffset = 0
    CX = 0
    CY = 0
    Dim Drawing As Boolean
    Drawing = True
    Do While Drawing = True
        LoadMapFromCache CMap
        If CX < 0 Then
            picView.Width = picView.Width + MapSize
            picSave.Width = picView.Width
            BitBlt picSave.hdc, 0, 0, picView.Width, picView.Height, picView.hdc, 0, 0, SRCCOPY
            picView.Cls
            BitBlt picView.hdc, MapSize, 0, picSave.Width, picSave.Height, picSave.hdc, 0, 0, SRCCOPY
            picSave.Cls
            IncrementX MapSize
            picView.Refresh
            CX = 0
        End If
        If CY < 0 Then
            picView.Height = picView.Height + MapSize
            picSave.Height = picView.Height
            BitBlt picSave.hdc, 0, 0, picView.Width, picView.Height, picView.hdc, 0, 0, SRCCOPY
            picView.Cls
            BitBlt picView.hdc, 0, MapSize, picSave.Width, picSave.Height, picSave.hdc, 0, 0, SRCCOPY
            picSave.Cls
            IncrementY MapSize
            picView.Refresh
            CY = 0
        End If
        If CX >= picView.Width Then
            picView.Width = CX + MapSize
            picSave.Width = picView.Width
            picView.Refresh
        End If
        If CY >= picView.Height Then
            picView.Height = CY + MapSize
            picSave.Height = picView.Height
            picView.Refresh
        End If
        DrawMap CX, CY
        CheckForLinkerErrors CMap, CX, CY
        TheMap(CMap).Drawn = True
        TheMap(CMap).X = CX
        TheMap(CMap).Y = CY
        If Map.ExitRight > 0 And TheMap(Map.ExitRight).Drawn = False Then
            CX = CX + MapSize
            CMap = Map.ExitRight
        ElseIf Map.ExitDown > 0 And TheMap(Map.ExitDown).Drawn = False Then
            CY = CY + MapSize
            CMap = Map.ExitDown
        ElseIf Map.ExitUp > 0 And TheMap(Map.ExitUp).Drawn = False Then
            CY = CY - MapSize
            CMap = Map.ExitUp
        ElseIf Map.ExitLeft > 0 And TheMap(Map.ExitLeft).Drawn = False Then
            CX = CX - MapSize
            CMap = Map.ExitLeft
        Else
            CMap = FindDrawableMap
            If CMap = 0 Then
                Drawing = False
            Else
                CX = TheMap(CMap).X
                CY = TheMap(CMap).Y
            End If
        End If
        
        'picView.Refresh
        'DoEvents
    Loop
    BitBlt picSave.hdc, 0, 0, picView.Width, picView.Height, picView.hdc, 0, 0, SRCCOPY
    picView.Refresh
    'SaveFinalMap
End Sub

Function FindDrawableMap() As Long
    Dim A As Long
    For A = 1 To 3000
        If TheMap(A).Drawn = True Then
            LoadMapFromCache CLng(A)
            If Map.ExitRight > 0 And TheMap(Map.ExitRight).Drawn = False Then
                FindDrawableMap = A
                Exit Function
            ElseIf Map.ExitDown > 0 And TheMap(Map.ExitDown).Drawn = False Then
                FindDrawableMap = A
                Exit Function
            ElseIf Map.ExitUp > 0 And TheMap(Map.ExitUp).Drawn = False Then
                FindDrawableMap = A
                Exit Function
            ElseIf Map.ExitLeft > 0 And TheMap(Map.ExitLeft).Drawn = False Then
                FindDrawableMap = A
                Exit Function
            End If
        End If
    Next A
End Function

Function FindReceivableMap() As Long
    Dim A As Long
    For A = 1 To 3000
        If TheMap(A).Received = True Then
            LoadMapFromCache A
            If Map.ExitRight > 0 And TheMap(Map.ExitRight).Received = False Then
                FindReceivableMap = A
                Exit Function
            ElseIf Map.ExitDown > 0 And TheMap(Map.ExitDown).Received = False Then
                FindReceivableMap = A
                Exit Function
            ElseIf Map.ExitUp > 0 And TheMap(Map.ExitUp).Received = False Then
                FindReceivableMap = A
                Exit Function
            ElseIf Map.ExitLeft > 0 And TheMap(Map.ExitLeft).Received = False Then
                FindReceivableMap = A
                Exit Function
            End If
        End If
    Next A
End Function

Sub DrawView()
    picView.Cls
    BitBlt picView.hdc, XOffset, YOffset, picSave.Width, picSave.Height, picSave.hdc, 0, 0, SRCCOPY
    picView.Refresh
End Sub

Sub IncrementX(Offset As Long)
    Dim A As Long
    For A = 1 To 3000
        If TheMap(A).Drawn = True Then
            TheMap(A).X = TheMap(A).X + Offset
            picView.Refresh
        End If
    Next A
End Sub

Sub IncrementY(Offset As Long)
    Dim A As Long
    For A = 1 To 3000
        If TheMap(A).Drawn = True Then
            TheMap(A).Y = TheMap(A).Y + Offset
            picView.Refresh
        End If
    Next A
End Sub

Private Sub TimeRequestAll_Timer()
    Dim FoundMap As Boolean
    If ReceivedMap = True Then
        If Not CMap = 3000 Then
            Do Until FoundMap = True
                If Not CMap = 3000 Then
                    CMap = CMap + 1
                    frmMain.lblDetails = "Now loading map " & CMap
                    frmMain.Refresh
                    LoadMapFromCache CMap
                    If Map.Version = 0 Then 'Found Map
                        FoundMap = True
                        frmMain.lblDetails = "Requesting Map " & CMap
                        ReceivedMap = False
                        frmMain.Refresh
                        SendSocket Chr$(4) + DoubleChar$(CMap)
                    End If
                Else
                    FoundMap = True
                    MapSize = 48
                    DrawItAll 1
                End If
            Loop
        Else
            FoundMap = True
            MapSize = 48
            DrawItAll 1
        End If
    End If
End Sub

Sub CheckForLinkerErrors(TheMapVal As Long, X As Long, Y As Long)
    Dim A As Long
    For A = 1 To 3000
        If Not A = TheMapVal Then
            If TheMap(A).Drawn = True Then
                If TheMap(A).X = X And TheMap(A).Y = Y Then
                    MsgBox "Linker error:  " + CStr(A)
                End If
            End If
        End If
    Next A
End Sub
