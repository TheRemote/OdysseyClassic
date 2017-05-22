VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "The Odyssey Classic Odyssey Server"
   ClientHeight    =   1605
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   7500
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1605
   ScaleWidth      =   7500
   StartUpPosition =   3  'Windows Default
   WindowState     =   1  'Minimized
   Begin VB.Timer tmrCloseScks 
      Interval        =   4000
      Left            =   1920
      Top             =   0
   End
   Begin VB.Timer ObjectTimer 
      Interval        =   10000
      Left            =   1440
      Top             =   0
   End
   Begin VB.Timer PlayerTimer 
      Interval        =   2000
      Left            =   960
      Top             =   0
   End
   Begin VB.Timer MinuteTimer 
      Interval        =   60000
      Left            =   480
      Top             =   0
   End
   Begin VB.Timer MapTimer 
      Interval        =   500
      Left            =   0
      Top             =   0
   End
   Begin VB.ListBox lstLog 
      Height          =   1230
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   7455
   End
   Begin VB.TextBox txtMessage 
      Height          =   285
      Left            =   0
      TabIndex        =   0
      Top             =   1320
      Width           =   7455
   End
   Begin VB.Menu mnuServer 
      Caption         =   "&Server"
      Begin VB.Menu mnuServerOptions 
         Caption         =   "&Options"
      End
      Begin VB.Menu mnuServerBandwidth 
         Caption         =   "&Bandwidth Usage"
      End
   End
   Begin VB.Menu mnuReports 
      Caption         =   "&Reports"
      Begin VB.Menu mnuReportsMaps 
         Caption         =   "&Free Maps"
      End
      Begin VB.Menu mnuReportsGods 
         Caption         =   "&Gods"
      End
   End
   Begin VB.Menu mnuDatabase 
      Caption         =   "&Database"
      Begin VB.Menu mnuReset 
         Caption         =   "&Reset"
         Begin VB.Menu mnuDatabaseResetAccounts 
            Caption         =   "Accounts"
         End
         Begin VB.Menu mnuDatabaseResetObjects 
            Caption         =   "Objects"
         End
         Begin VB.Menu mnuDatabaseResetMonsters 
            Caption         =   "Monsters"
         End
         Begin VB.Menu mnuDatabaseResetNPCs 
            Caption         =   "NPCs"
         End
         Begin VB.Menu mnuDatabaseResetMessages 
            Caption         =   "Messages"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuDatabaseResetPosts 
            Caption         =   "Posts"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuDatabaseResetGuilds 
            Caption         =   "Guilds"
         End
         Begin VB.Menu mnuDatabaseResetGods 
            Caption         =   "Gods"
         End
         Begin VB.Menu mnuDatabaseResetBans 
            Caption         =   "Bans"
         End
      End
   End
   Begin VB.Menu mnuAddGod 
      Caption         =   "Add InGame God"
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    gHW = Me.hWnd
End Sub
Private Sub Form_Unload(Cancel As Integer)
    ShutdownServer
End Sub
Private Sub Form_Resize()
    If Not Me.WindowState = 1 Then
        lstLog.Width = Me.ScaleWidth
        lstLog.Height = Me.ScaleHeight - txtMessage.Height
        txtMessage.Top = lstLog.Height
        txtMessage.Width = Me.ScaleWidth
    End If
End Sub

Private Sub MapTimer_Timer()
    Dim A As Long, B As Long, C As Long, D As Long, E As Long
    Dim MapNum As Long
    Dim St1 As String
    
    For MapNum = 1 To 2000
        With Map(MapNum)
            St1 = ""
            For A = 0 To 9
                If .Door(A).Att > 0 Then
                    If GetTickCount - .Door(A).T > 10000 Then
                        .Tile(.Door(A).X, .Door(A).Y).Att = .Door(A).Att
                        .Door(A).Att = 0
                        St1 = St1 + DoubleChar(2) + Chr$(37) + Chr$(A)
                    End If
                End If
            Next A
            If .NumPlayers > 0 Then
                For A = 0 To 5
                    With .Monster(A)
                        If .Monster > 0 Then
                            If .Target = 0 Then
                                'Random Movement
                                If Rnd < 0.3 Then
                                    .D = Int(Rnd * 4)
                                    Select Case .D
                                        Case 0 'Up
                                            If .Y > 0 Then
                                                If IsVacant(MapNum, .X, .Y - 1) Then
                                                    .Y = .Y - 1
                                                    St1 = St1 + DoubleChar(5) + Chr$(40) + Chr$(A) + Chr$(.X) + Chr$(.Y) + Chr$(.D)
                                                End If
                                            End If
                                        Case 1 'Down
                                            If .Y < 11 Then
                                                If IsVacant(MapNum, .X, .Y + 1) Then
                                                    .Y = .Y + 1
                                                    St1 = St1 + DoubleChar(5) + Chr$(40) + Chr$(A) + Chr$(.X) + Chr$(.Y) + Chr$(.D)
                                                End If
                                            End If
                                        Case 2 'Left
                                            If .X > 0 Then
                                                If IsVacant(MapNum, .X - 1, .Y) Then
                                                    .X = .X - 1
                                                    St1 = St1 + DoubleChar(5) + Chr$(40) + Chr$(A) + Chr$(.X) + Chr$(.Y) + Chr$(.D)
                                                End If
                                            End If
                                        Case 3 'Right
                                            If .X < 11 Then
                                                If IsVacant(MapNum, .X + 1, .Y) Then
                                                    .X = .X + 1
                                                    St1 = St1 + DoubleChar(5) + Chr$(40) + Chr$(A) + Chr$(.X) + Chr$(.Y) + Chr$(.D)
                                                End If
                                            End If
                                    End Select
                                End If
                            Else
                                'Move Toward Target
                                B = .Target
                                If Player(B).Mode = modePlaying And Player(B).Map = MapNum Then
                                    C = .X
                                    D = .Y
                                    E = .D
                                    If Sqr(CSng(CLng(Player(B).X) - C) ^ 2 + CSng(CLng(Player(B).Y) - D) ^ 2) > 1 Then
                                        .AttackCounter = 0
                                        If Rnd < 0.5 Then
                                            If C < Player(B).X Then
                                                If IsVacant(MapNum, C + 1, CByte(D)) Then
                                                    C = C + 1
                                                    E = 3
                                                End If
                                            ElseIf C > Player(B).X Then
                                                If IsVacant(MapNum, C - 1, CByte(D)) Then
                                                    C = C - 1
                                                    E = 2
                                                End If
                                            End If
                                            If C = .X And D = .Y Then
                                                If D < Player(B).Y Then
                                                    If IsVacant(MapNum, CByte(C), D + 1) Then
                                                        D = D + 1
                                                        E = 1
                                                    ElseIf Rnd < 0.2 Then
                                                        If Rnd < 0.5 Then
                                                            If C > 0 Then
                                                                If IsVacant(MapNum, C - 1, CByte(D)) Then
                                                                    C = C - 1
                                                                    E = 2
                                                                End If
                                                            End If
                                                        Else
                                                            If C < 11 Then
                                                                If IsVacant(MapNum, C + 1, CByte(D)) Then
                                                                    C = C + 1
                                                                    E = 3
                                                                End If
                                                            End If
                                                        End If
                                                    End If
                                                ElseIf D > Player(B).Y Then
                                                    If IsVacant(MapNum, CByte(C), D - 1) Then
                                                        D = D - 1
                                                        E = 0
                                                    ElseIf Rnd < 0.2 Then
                                                        If Rnd < 0.5 Then
                                                            If C > 0 Then
                                                                If IsVacant(MapNum, C - 1, CByte(D)) Then
                                                                    C = C - 1
                                                                    E = 2
                                                                End If
                                                            End If
                                                        Else
                                                            If C < 11 Then
                                                                If IsVacant(MapNum, C + 1, CByte(D)) Then
                                                                    C = C + 1
                                                                    E = 3
                                                                End If
                                                            End If
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        Else
                                            If D < Player(B).Y Then
                                                If IsVacant(MapNum, CByte(C), D + 1) Then
                                                    D = D + 1
                                                    E = 1
                                                End If
                                            ElseIf D > Player(B).Y Then
                                                If IsVacant(MapNum, CByte(C), D - 1) Then
                                                    D = D - 1
                                                    E = 0
                                                End If
                                            End If
                                            If C = .X And D = .Y Then
                                                If C < Player(B).X Then
                                                    If IsVacant(MapNum, C + 1, CByte(D)) Then
                                                        C = C + 1
                                                        E = 3
                                                    ElseIf Rnd < 0.2 Then
                                                        If Rnd < 0.5 Then
                                                            If D > 0 Then
                                                                If IsVacant(MapNum, CByte(C), D - 1) Then
                                                                    D = D - 1
                                                                    E = 0
                                                                End If
                                                            End If
                                                        Else
                                                            If D < 11 Then
                                                                If IsVacant(MapNum, CByte(C), D + 1) Then
                                                                    D = D + 1
                                                                    E = 1
                                                                End If
                                                            End If
                                                        End If
                                                    End If
                                                ElseIf C > Player(B).X Then
                                                    If IsVacant(MapNum, C - 1, CByte(D)) Then
                                                        C = C - 1
                                                        E = 2
                                                    ElseIf Rnd < 0.2 Then
                                                        If Rnd < 0.5 Then
                                                            If D > 0 Then
                                                                If IsVacant(MapNum, CByte(C), D - 1) Then
                                                                    D = D - 1
                                                                    E = 0
                                                                End If
                                                            End If
                                                        Else
                                                            If D < 11 Then
                                                                If IsVacant(MapNum, CByte(C), D + 1) Then
                                                                    D = D + 1
                                                                    E = 1
                                                                End If
                                                            End If
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        End If
                                        If C <> .X Or D <> .Y Or E <> .D Then
                                            .X = C
                                            .Y = D
                                            .D = E
                                            St1 = St1 + DoubleChar(5) + Chr$(40) + Chr$(A) + Chr$(C) + Chr$(D) + Chr$(.D)
                                        End If
                                    Else
                                        'Attack Player
                                        If .AttackCounter = 0 Then
                                            .AttackCounter = 1
                                            C = Player(B).X
                                            D = Player(B).Y
                                            If .X < C And .D <> 3 Then
                                                .D = 3
                                                St1 = St1 + DoubleChar(5) + Chr$(40) + Chr$(A) + Chr$(.X) + Chr$(.Y) + Chr$(3)
                                            ElseIf .X > C And .D <> 2 Then
                                                .D = 2
                                                St1 = St1 + DoubleChar(5) + Chr$(40) + Chr$(A) + Chr$(.X) + Chr$(.Y) + Chr$(2)
                                            ElseIf .Y < D And .D <> 1 Then
                                                .D = 1
                                                St1 = St1 + DoubleChar(5) + Chr$(40) + Chr$(A) + Chr$(.X) + Chr$(.Y) + Chr$(1)
                                            ElseIf .Y > D And .D <> 0 Then
                                                .D = 0
                                                St1 = St1 + DoubleChar(5) + Chr$(40) + Chr$(A) + Chr$(.X) + Chr$(.Y) + Chr$(0)
                                            End If
                                            If Int(Rnd * 100) > Player(B).Agility Then
                                                C = PlayerArmor(B, Monster(.Monster).Strength)
                                                If C < 0 Then C = 0
                                                If C > 255 Then C = 255
                                                SendSocket B, Chr$(50) + Chr$(0) + Chr$(A) + Chr$(C)
                                                St1 = St1 + DoubleChar(2) + Chr$(41) + Chr$(A)
                                                With Player(B)
                                                    If C >= .HP Then
                                                        'Player Died
                                                        SendSocket B, Chr$(53) + Chr$(Map(MapNum).Monster(A).Monster) 'Monster Killed You
                                                        SendAllBut B, Chr$(62) + Chr$(B) + Chr$(Map(MapNum).Monster(A).Monster) 'Player was killed by monster
                                                        PlayerDied B
                                                    Else
                                                        .HP = .HP - C
                                                    End If
                                                End With
                                            Else
                                                SendSocket B, Chr$(50) + Chr$(1) + Chr$(A) + Chr$(0)
                                            End If
                                        Else
                                            .AttackCounter = 0
                                        End If
                                    End If
                                Else
                                    .Target = 0
                                    .Distance = Monster(.Monster).Sight
                                End If
                            End If
                        Else
                            With Map(MapNum).MonsterSpawn(Int(A / 2))
                                If .Monster > 0 Then
                                    If Int(Rnd * .Rate) = 0 Then
                                        St1 = St1 + NewMapMonster(MapNum, A)
                                    End If
                                End If
                            End With
                        End If
                    End With
                Next A
                If St1 <> "" Then
                    SendToMapRaw MapNum, St1
                End If
            Else
                'Map is Empty
                If .ResetTimer > 0 And GetTickCount - .ResetTimer >= World.MapResetTime Then
                    ResetMap MapNum
                End If
            End If
        End With
    Next MapNum
End Sub
Private Sub MinuteTimer_Timer()
    Dim St As String, A As Long, B As Long, C As Long
 
    World.Hour = World.Hour + 1
    If World.Hour > 24 Then World.Hour = 1
    If World.Hour = 7 Then
        blnNight = False
        SendAll Chr$(54)
    End If
    If World.Hour = 22 Then
        blnNight = True
        SendAll Chr$(55)
    End If
    
    If World.BackupInterval > 0 Then
        BackupCounter = BackupCounter + 1
        If BackupCounter >= World.BackupInterval Then
            BackupCounter = 0
            'Backup Server Data
            For A = 1 To MaxUsers
                If Player(A).Mode = modePlaying Then
                    SavePlayerData A
                End If
            Next A
            SaveFlags
            SaveObjects
        End If
    End If
    
    If World.LastUpdate <> CLng(Date) Then
        World.LastUpdate = CLng(Date)
        DataRS.Edit
        DataRS!LastUpdate = World.LastUpdate
        DataRS.Update
        
        #If UseGuilds Then
            'Update Guilds
            For A = 1 To 255
                With Guild(A)
                    If .Name <> "" Then
                        If .Bank < 0 And World.LastUpdate >= .DueDate Then
                            'Debt not payed, delete guild
                            DeleteGuild A, 0
                        ElseIf CountGuildMembers(A) < 3 Then
                            'Not enough members, guild deleted
                            DeleteGuild A, 1
                        Else
                            If .Bank >= 0 Then
                                C = 0
                            Else
                                C = 1
                            End If
                            
                            'Pay bill
                            .Bank = .Bank - 200
                            For B = 0 To 19
                                If .Member(B).Name <> "" Then
                                    .Bank = .Bank - 50
                                End If
                            Next B
                            If .Hall > 0 Then
                                .Bank = .Bank - Hall(.Hall).Upkeep
                            End If
                            If C = 0 And .Bank < 0 Then
                                .DueDate = CLng(Date) + 2
                            End If
                            If .Bank >= 0 Then
                                SendToGuild A, Chr$(74) + Chr$(74) + QuadChar(.Bank)
                            Else
                                SendToGuild A, Chr$(74) + QuadChar(Abs(.Bank)) + QuadChar(.DueDate)
                            End If
                            GuildRS.Seek "=", A
                            If GuildRS.NoMatch = False Then
                                GuildRS.Edit
                                GuildRS!Bank = .Bank
                                GuildRS!DueDate = .DueDate
                                GuildRS.Update
                            End If
                        End If
                    End If
                End With
            Next A
        #End If
    End If
    
    RunScript ("MINUTETIMER")
End Sub

Private Sub mnuAddGod_Click()
Dim A As Long, B As Long, C As Long, D As Long
A = FindPlayer(InputBox$("Enter the name of a person ingame in which you would like to change their access: ", "Enter Player Name"))
B = Val(InputBox$("Enter a numerical number between 0 and 11 in which the selected player's new access shall be: ", "Enter Access"))
    If A >= 1 And A <= MaxUsers And B >= 0 And B <= 11 Then
        With Player(A)
            If .Access <= 11 Then
                .Access = B
                SendSocket A, Chr$(65) + Chr$(B)
                If .Access > 0 Then
                    SendAllBut A, Chr$(91) + Chr$(A) + Chr$(3)
                    .Status = 3
                        
                    D = FindGodAccount(.ComputerID)
                    C = FreeGodNum
                    If D > 0 Then
                        GodRS.Index = "number"
                        GodRS.Seek "=", D
                        If GodRS.NoMatch = False Then
                            GodRS.Edit
                            GodRS!User = .User
                            GodRS!ComputerID = EncryptString(.ComputerID)
                            GodRS!Access = B
                            GodRS.Update
                        End If
                                                                            
                        With GodData(D)
                            .User = Player(A).User
                            .ComputerID = Player(A).ComputerID
                            .Access = B
                            .InUse = True
                        End With
                    Else
                        If C <= 50 Then
                            GodRS.AddNew
                            GodRS!number = C
                            GodRS!User = .User
                            GodRS!ComputerID = EncryptString(.ComputerID)
                            GodRS!Access = B
                            GodRS.Update
                            
                            With GodData(C)
                                .User = Player(A).User
                                .ComputerID = Player(A).ComputerID
                                .Access = B
                                .InUse = True
                            End With
                        Else
                            MsgBox "There are already 50 active god accounts. Please remove one to add one!", vbOKOnly, "Cannot add a god"
                        End If
                    End If
                Else
                    D = FindGodAccount(.ComputerID)
                    GodRS.Index = "number"
                    GodRS.Seek "=", D
                    If GodRS.NoMatch = False Then
                        GodRS.Delete
                    End If
                    SendAllBut A, Chr$(91) + Chr$(A) + Chr$(0)
                    Player(A).Status = 0
                    MsgBox "Successfully removed god account!", vbOKOnly, "Removed God"
                End If
            Else
                With Player(A)
                    .Access = 0
                    SendSocket A, Chr$(65) + Chr$(0)
                    SendAllBut A, Chr$(91) + Chr$(A) + Chr$(0)
                    .Status = 0
                End With
            End If
        End With
        MsgBox "Added God Successfully!", vbExclamation + vbOKOnly, "Added Successfully"
    Else
        MsgBox "You have entered an invalid selection. Please choose 'Add InGame God' again supplying valid information!", vbExclamation + vbOKOnly, "Invalid Information"
    End If
End Sub

Private Sub mnuDatabaseResetAccounts_Click()
    If MsgBox("Are you *sure* you wish to delete every account?", vbYesNo) = vbYes Then
        If MsgBox("About to reset every account -- continue?", vbYesNo) = vbYes Then
            If Not UserRS.BOF Then
                UserRS.MoveFirst
                While Not UserRS.EOF
                    DeleteAccount
                    UserRS.MoveNext
                Wend
            End If
            UserRS.Close
            Set UserRS = Nothing
            DB.TableDefs.Delete "Accounts"
            CreateAccountsTable
            Set UserRS = DB.TableDefs("Accounts").OpenRecordset(dbOpenTable)
            UserRS.Index = "User"
        End If
    End If
End Sub

Private Sub mnuDatabaseResetBans_Click()
    Dim A As Long
    
    If MsgBox("Are you *sure* you wish to delete every Ban?", vbYesNo) = vbYes Then
        If MsgBox("About to reset every Ban -- continue?", vbYesNo) = vbYes Then
            BanRS.Close
            Set BanRS = Nothing
            DB.TableDefs.Delete "Bans"
            CreateBansTable
            Set BanRS = DB.TableDefs("Bans").OpenRecordset(dbOpenTable)
            BanRS.Index = "Number"
            For A = 1 To 50
                With Ban(A)
                    .Banner = ""
                    .ComputerID = ""
                    .Reason = ""
                    .Name = ""
                    .UnbanDate = 0
                End With
            Next A
        End If
    End If
End Sub

Private Sub mnuDatabaseResetGods_Click()
    Dim A As Long
    
    If MsgBox("Are you *sure* you wish to delete every God?", vbYesNo) = vbYes Then
        If MsgBox("About to reset every God -- continue?", vbYesNo) = vbYes Then
            GodRS.Close
            Set GodRS = Nothing
            DB.TableDefs.Delete "Gods"
            CreateGodTable
            Set GodRS = DB.TableDefs("Gods").OpenRecordset(dbOpenTable)
            GodRS.Index = "Number"
            For A = 1 To 50
                With GodData(A)
                    .Access = 0
                    .ComputerID = ""
                    .InUse = False
                    .User = ""
                End With
            Next A
        End If
        
        If UserRS.BOF = False Then
        UserRS.MoveFirst
        While UserRS.EOF = False
            If UserRS!Access > 0 Then
                UserRS.Edit
                UserRS!Access = 0
                UserRS.Update
            End If
            UserRS.MoveNext
        Wend
    End If
    End If
End Sub

Private Sub mnuDatabaseResetGuilds_Click()
    Dim A As Long, B As Long
    
    If MsgBox("Are you *sure* you wish to delete every guild?", vbYesNo) = vbYes Then
        If MsgBox("About to reset every guild -- continue?", vbYesNo) = vbYes Then
            GuildRS.Close
            Set GuildRS = Nothing
            DB.TableDefs.Delete "Guilds"
            CreateGuildsTable
            Set GuildRS = DB.TableDefs("Guilds").OpenRecordset(dbOpenTable)
            GuildRS.Index = "Number"
            For A = 1 To 255
                With Guild(A)
                    .Bank = 0
                    .Bookmark = 0
                    .Name = ""
                    .Hall = 0
                    .DueDate = 0
                    .Sprite = 0
                    
                    For B = 0 To 4
                        With .Declaration(B)
                            .Guild = 0
                            .Type = 0
                        End With
                    Next B
                    For B = 0 To 19
                        With .Member(B)
                            .Name = ""
                            .Rank = 0
                        End With
                    Next B
                End With
            Next A
            
            'For A = 1 To 2
            '    GuildRS.AddNew
            '    GuildRS!number = A
            '    If A = 1 Then
            '        GuildRS!Name = "Angelic Ones"
            '        GuildRS!Hall = 1
            '        GuildRS!Sprite = 78
            '    Else
            '        GuildRS!Name = "Devil's Own"
            '        GuildRS!Hall = 2
            '        GuildRS!Sprite = 79
            '    End If
            '    'GuildRS!Hall = 0
            '    'GuildRS!Sprite = 0
            '    'Select Case A
            '    '    Case 1
            '    '        GuildRS!Name = "Bone"
            '    '    Case 2
            '    '        GuildRS!Name = "Fire"
            '    '    Case 3
            '    '        GuildRS!Name = "Ice"
            '    '    Case 4
            '    '        GuildRS!Name = "Rock"
            '    'End Select
            '    GuildRS!Bank = 0
            '    GuildRS!DueDate = 0
            '    For B = 0 To 4
            '        GuildRS("DeclarationGuild" + CStr(B)) = 0
            '        GuildRS("DeclarationType" + CStr(B)) = 0
            '    Next B
            '    For B = 0 To 19
            '        GuildRS("MemberName" + CStr(B)) = ""
            '        GuildRS("MemberRank" + CStr(B)) = 0
            '    Next B
            '    GuildRS.Update
            'Next A
        End If
    End If
End Sub
Private Sub mnuDatabaseResetMessages_Click()
    Dim A As Long, B As Long
    
    If MsgBox("Are you *sure* you wish to delete every message?", vbYesNo) = vbYes Then
        If MsgBox("About to reset every message -- continue?", vbYesNo) = vbYes Then
            If Not UserRS.BOF Then
                UserRS.MoveFirst
                While Not UserRS.EOF
                    For A = 1 To 20
                        UserRS("Msg" + CStr(A)) = 0
                    Next A
                    UserRS.MoveNext
                Wend
            End If
            
            If Not PostRS.BOF Then
                PostRS.MoveFirst
                While Not PostRS.EOF
                    For A = 1 To 20
                        PostRS("Msg" + CStr(A)) = 0
                    Next A
                    PostRS.MoveNext
                Wend
            End If
            
            For A = 1 To 255
                With Post(A)
                    For B = 1 To 30
                        .Msg(B) = 0
                    Next B
                End With
            Next A
            
            World.MsgCounter = 0
            With DataRS
                .Edit
                !MsgCounter = 0
                .Update
            End With
            
            MsgRS.Close
            Set MsgRS = Nothing
            DB.TableDefs.Delete "Messages"
            CreateMessagesTable
            Set MsgRS = DB.TableDefs("Messages").OpenRecordset(dbOpenTable)
            MsgRS.Index = "Number"
        End If
    End If
End Sub
Private Sub mnuDatabaseResetMonsters_Click()
    Dim A As Long
    
    If MsgBox("Are you *sure* you wish to delete every monster?", vbYesNo) = vbYes Then
        If MsgBox("About to reset every monster -- continue?", vbYesNo) = vbYes Then
            MonsterRS.Close
            Set MonsterRS = Nothing
            DB.TableDefs.Delete "Monsters"
            CreateMonstersTable
            Set MonsterRS = DB.TableDefs("Monsters").OpenRecordset(dbOpenTable)
            MonsterRS.Index = "Number"
            For A = 1 To 255
                With Monster(A)
                    .Armor = 0
                    .Agility = 0
                    .Description = ""
                    .flags = 0
                    .HP = 0
                    .Name = ""
                    .Sight = 0
                    .Speed = 0
                    .Sprite = 0
                    .Strength = 0
                    .Object(0) = 0
                    .Object(1) = 0
                    .Object(2) = 0
                    .Value(0) = 0
                    .Value(1) = 0
                    .Value(2) = 0
                End With
            Next A
        End If
    End If
End Sub

Private Sub mnuDatabaseResetNPCs_Click()
    Dim A As Long, B As Long
    
    If MsgBox("Are you *sure* you wish to delete every NPC?", vbYesNo) = vbYes Then
        If MsgBox("About to reset every NPC -- continue?", vbYesNo) = vbYes Then
            NPCRS.Close
            Set NPCRS = Nothing
            DB.TableDefs.Delete "NPCS"
            CreateNPCsTable
            Set NPCRS = DB.TableDefs("NPCS").OpenRecordset(dbOpenTable)
            NPCRS.Index = "Number"
            For A = 1 To 255
                With NPC(A)
                    .Name = ""
                    .JoinText = ""
                    .LeaveText = ""
                    .flags = 0
                    For B = 0 To 9
                        With .SaleItem(B)
                            .GiveObject = 0
                            .GiveValue = 0
                            .TakeObject = 0
                            .TakeValue = 0
                        End With
                    Next B
                End With
            Next A
        End If
    End If
End Sub
Private Sub mnuDatabaseResetObjects_Click()
    Dim A As Long
    
    If MsgBox("Are you *sure* you wish to delete every object?", vbYesNo) = vbYes Then
        If MsgBox("About to reset every object -- continue?", vbYesNo) = vbYes Then
            ObjectRS.Close
            Set ObjectRS = Nothing
            DB.TableDefs.Delete "Objects"
            CreateObjectsTable
            Set ObjectRS = DB.TableDefs("Objects").OpenRecordset(dbOpenTable)
            ObjectRS.Index = "Number"
            
            For A = 1 To 255
                With Object(A)
                    .Data(0) = 0
                    .Data(1) = 0
                    .Data(2) = 0
                    .Data(3) = 0
                    .flags = 0
                    .Name = ""
                    .Picture = 0
                    .Type = 0
                End With
            Next A
        End If
    End If
End Sub
Private Sub mnuDatabaseResetPosts_Click()
    Dim A As Long, B As Long
    
    If MsgBox("Are you *sure* you wish to delete every message post?", vbYesNo) = vbYes Then
        If MsgBox("About to reset every message post -- continue?", vbYesNo) = vbYes Then
            If Not PostRS.BOF Then
                PostRS.MoveFirst
                While Not PostRS.EOF
                    For A = 1 To 30
                        B = PostRS("Msg" + CStr(A))
                        If B > 0 Then
                            DeleteMessage B
                        End If
                    Next A
                    PostRS.MoveNext
                Wend
            End If
            PostRS.Close
            Set PostRS = Nothing
            DB.TableDefs.Delete "Posts"
            CreatePostsTable
            Set PostRS = DB.TableDefs("Posts").OpenRecordset(dbOpenTable)
            PostRS.Index = "Number"
            
            For A = 1 To 255
                With Post(A)
                    .flags = 0
                    .Name = ""
                    For B = 1 To 30
                        .Msg(B) = 0
                    Next B
                End With
            Next A
        End If
    End If
End Sub

Private Sub mnuReportsGods_Click()
    Dim A As Long, StartFree As Long, IsFree As Boolean
    Open "report.txt" For Output As #1
    
    Print #1, "***Odyssey God Report ***"
    Print #1, ""
    
    If UserRS.BOF = False Then
        UserRS.MoveFirst
        While UserRS.EOF = False
            If UserRS!Access > 0 Then
                Print #1, "User=" + UserRS!User + " Name=" + UserRS!Name + " Access=" + CStr(UserRS!Access)
            End If
            UserRS.MoveNext
        Wend
    End If
    
    Close #1
    
    Shell "notepad.exe report.txt", vbNormalFocus
End Sub
Private Sub mnuReportsMaps_Click()
    Dim A As Long, StartFree As Long, IsFree As Boolean
    Open "report.txt" For Output As #1
    
    Print #1, "***Odyssey Free Map Report ***"
    Print #1, ""
    
    IsFree = False
    
    For A = 1 To 2000
        MapRS.Seek "=", A
        If MapRS.NoMatch = False Then
            If IsFree = True Then
                If StartFree < A - 1 Then
                    Print #1, CStr(StartFree) + "-" + CStr(A - 1)
                Else
                    Print #1, CStr(A - 1)
                End If
                IsFree = False
            End If
        Else
            If IsFree = False Then
                StartFree = A
                IsFree = True
            End If
        End If
    Next A
    
    If IsFree = True Then
        If StartFree < 2000 Then
            Print #1, CStr(StartFree) + "-2000"
        Else
            Print #1, "2000"
        End If
    End If
    Close #1
    
    Shell "notepad.exe report.txt", vbNormalFocus
End Sub

Private Sub mnuServerBandwidth_Click()
    Dim T As Long
    T = Abs(GetTickCount - StartTimeStamp)
    
    If BytesReceived < 2000000000 Then
        PrintLog "Total bytes received: " + CStr(BytesReceived) + ", " + CStr(Int(BytesReceived / T)) + " k/sec (average)"
    Else
        PrintLog "Error -- Bytes received exceeded 2 billion!"
    End If
    If BytesSent < 2000000000 Then
        PrintLog "Total bytes Sent: " + CStr(BytesSent) + ", " + CStr(Int(BytesSent / T)) + " k/sec (average)"
    Else
        PrintLog "Error -- Bytes sent exceeded 2 billion!"
    End If
End Sub

Private Sub mnuServerOptions_Click()
    frmOptions.Show 1
End Sub

Private Sub ObjectTimer_Timer()
   Dim MapNum As Long, A As Long, St1 As String
   
   For MapNum = 1 To 2000
        With Map(MapNum)
            If .NumPlayers > 0 Then
                St1 = ""
                For A = 0 To 49
                    With .Object(A)
                        If .Object > 0 Then
                            If Map(MapNum).Tile(.X, .Y).Att <> 5 And .TimeStamp > 0 Then
                                If GetTickCount - .TimeStamp >= World.ObjResetTime Then
                                    .Object = 0
                                    St1 = St1 + DoubleChar(2) + Chr$(15) + Chr$(A)
                                End If
                            End If
                        End If
                    End With
                Next A
                If St1 <> "" Then
                    SendToMapRaw MapNum, St1
                End If
            End If
        End With
    Next MapNum
End Sub
Private Sub PlayerTimer_Timer()
    Dim St1 As String, A As Long
    Dim PlayerNum As Long
    
    For PlayerNum = 1 To MaxUsers
        With Player(PlayerNum)
            If .InUse = True Then
                St1 = ""
                If GetTickCount - .LastMsg >= 30000 And GetTickCount - .LastMsg <= 35000 Then
                    If .Mode <> modeNotConnected Then
                        'Send ping
                        St1 = St1 + DoubleChar(1) + Chr$(58)
                        .LastMsg = GetTickCount - 35000
                    Else
                        CloseClientSocket PlayerNum
                    End If
                End If
                If GetTickCount - .LastMsg >= 60000 Then
                    'Lag time out
                    CloseClientSocket PlayerNum
                End If
                If .Mode = modePlaying Then
                    If .HP < .MaxHP Then
                        A = .HP
                        A = A + Int(.Endurance / 5) + 1
                        If A > .MaxHP Then A = .MaxHP
                        .HP = A
                        St1 = St1 + DoubleChar(2) + Chr$(46) + Chr$(.HP)
                    End If
                    If .Energy < .MaxEnergy Then
                        A = .Energy
                        If .HP > 0 Then
                            A = A + Int((CSng(.HP) / CSng(.MaxHP)) * 4)
                        End If
                        If A > .MaxEnergy Then A = .MaxEnergy
                        .Energy = A
                        St1 = St1 + DoubleChar(2) + Chr$(47) + Chr$(.Energy)
                    End If
                    If .Mana < .MaxMana Then
                        A = .Mana
                        A = A + Int(.Intelligence / 5) + 1
                        If A > .MaxMana Then A = .MaxMana
                        .Mana = A
                        St1 = St1 + DoubleChar(2) + Chr$(48) + Chr$(.Mana)
                    End If
                    If St1 <> "" Then
                        SendRaw PlayerNum, St1
                    End If
                    For A = 1 To MaxPlayerTimers
                        If .ScriptTimer(A) > 0 And GetTickCount() >= .ScriptTimer(A) Then
                            Parameter(0) = PlayerNum
                            .ScriptTimer(A) = 0
                            RunScript .Script(A)
                        End If
                    Next A
                End If
            End If
        End With
    Next PlayerNum
End Sub

Private Sub tmrCloseScks_Timer()
Dim A As Long
    'Wait Procedure for Sockets
    For A = 1 To MaxUsers
        If CloseSocketQue(A) > 0 Then
            CloseClientSocket CloseSocketQue(A)
            CloseSocketQue(A) = 0
        End If
    Next A
End Sub

Private Sub txtMessage_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Or KeyAscii = 10 Then
        SendAll Chr$(30) + txtMessage
        PrintLog "Server Message: " + txtMessage
        txtMessage = ""
    End If
End Sub


