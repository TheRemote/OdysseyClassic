Attribute VB_Name = "modProcess"
Option Explicit

Public Sub ReceiveData()
    Dim PacketLength As Integer, PacketID As Integer
    Dim St As String, St1 As String
    Dim A As Long, B As Long, C As Long, D As Long
    
    SocketData = SocketData + Receive(ClientSocket)
LoopRead:
    If Len(SocketData) >= 3 Then
        PacketLength = GetInt(Mid$(SocketData, 1, 2))
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
                Select Case PacketID
                    Case 0 'Error Logging On
                        If Len(St) >= 1 Then
                            Select Case Asc(Mid$(St, 1, 1))
                                Case 0 'Custom Message
                                    If Len(St) >= 2 Then
                                        MessageBox frmMenu.hWnd, Mid$(St, 2), TitleString, vbOKOnly + vbExclamation
                                    End If
                                Case 1 'Invalid User/Pass
                                    MessageBox frmMenu.hWnd, "Invalid user name/password!", TitleString, vbOKOnly + vbExclamation
                                Case 2 'Account already in use
                                    MessageBox frmMenu.hWnd, "Someone is already using that account!", TitleString, vbOKOnly + vbExclamation
                                Case 3 'Banned
                                    If Len(St) >= 5 Then
                                        A = Asc(Mid$(St, 2, 1)) * 16777216 + Asc(Mid$(St, 3, 1)) * 65536 + Asc(Mid$(St, 4, 1)) * 256& + Asc(Mid$(St, 5, 1))
                                        If Len(St) > 5 Then
                                            MessageBox frmMenu.hWnd, "You are banned from The Odyssey Online Classic until " + CStr(CDate(A)) + " (" + Mid$(St, 6) + ")!", TitleString, vbOKOnly
                                        Else
                                            MessageBox frmMenu.hWnd, "You are banned from The Odyssey Online Classic until " + CStr(CDate(A)) + "!", TitleString, vbOKOnly
                                        End If
                                        CloseClientSocket 0
                                    End If
                                Case 4 'Server Full
                                    MessageBox frmMenu.hWnd, "The server is full, please try again in a few minutes!", TitleString, vbOKOnly + vbExclamation
                                Case 5 'Logging in with a god account
                                    MessageBox frmMenu.hWnd, "You are not permitted to use this god account. Your access is now set at 0!", TitleString, vbOKOnly + vbExclamation
                                    CloseClientSocket 0
                            End Select
                        End If
                        CloseClientSocket 1
                        
                    Case 1 'Error Creating New Account
                        If Len(St) >= 1 Then
                            Select Case Asc(Mid$(St, 1, 1))
                                Case 0 'Custom Message
                                    If Len(St) >= 2 Then
                                        MessageBox frmAccount.hWnd, Mid$(St, 2), TitleString, vbOKOnly + vbExclamation
                                    End If
                                Case 1 'User name already in use
                                    MessageBox frmAccount.hWnd, "That user name is already in use.  Please try another.", TitleString, vbOKOnly + vbExclamation
                            End Select
                        End If
                        CloseClientSocket 2
                        
                    Case 2 'Account Created
                        CloseClientSocket 0
                        MessageBox frmMenu.hWnd, "Your account has been created successfully!  Please write down your user name and password somewhere safe so that you do not loose them.  Click Login to connect to the game server.", TitleString, vbOKOnly + vbExclamation
                    
                    Case 3 'Logged On / Character Data
                        If frmWait_Loaded = True Then Unload frmWait
                        If frmLogin_Loaded = True Then Unload frmLogin
                        If frmNewCharacter_Loaded = True Then Unload frmNewCharacter
                        If frmCharacter_Loaded = False Then Load frmCharacter
                        If Len(St) >= 26 Then
                            With Character
                                .Name = ""
                                .Class = Asc(Mid$(St, 1, 1))
                                .Gender = Asc(Mid$(St, 2, 1))
                                .Sprite = Asc(Mid$(St, 3, 1))
                                .HP = Asc(Mid$(St, 4, 1))
                                .Energy = Asc(Mid$(St, 5, 1))
                                .Mana = Asc(Mid$(St, 6, 1))
                                .MaxHP = Asc(Mid$(St, 7, 1))
                                .MaxEnergy = Asc(Mid$(St, 8, 1))
                                .MaxMana = Asc(Mid$(St, 9, 1))
                                .Strength = Asc(Mid$(St, 10, 1))
                                .Agility = Asc(Mid$(St, 11, 1))
                                .Endurance = Asc(Mid$(St, 12, 1))
                                .Intelligence = Asc(Mid$(St, 13, 1))
                                .level = Asc(Mid$(St, 14, 1))
                                .Status = Asc(Mid$(St, 15, 1))
                                .Guild = Asc(Mid$(St, 16, 1))
                                .GuildRank = Asc(Mid$(St, 17, 1))
                                .Access = Asc(Mid$(St, 18, 1))
                                .Index = Asc(Mid$(St, 19, 1))
                                .Experience = Asc(Mid$(St, 20, 1)) * 16777216 + Asc(Mid$(St, 21, 1)) * 65536 + Asc(Mid$(St, 22, 1)) * 256& + Asc(Mid$(St, 23, 1))
                                .StatPoints = Asc(Mid$(St, 24, 1)) * 256& + Asc(Mid$(St, 25, 1))
                                St = Mid$(St, 26)
                                A = InStr(St, Chr$(0))
                                If A > 1 Then
                                    .Name = Mid$(St, 1, A - 1)
                                    If A < Len(St) Then
                                        .desc = Mid$(St, A + 1)
                                    End If
                                End If
                            End With
                            With frmCharacter
                                .frameCharacter.Visible = True
                                .lblName = Character.Name
                                .lblClass = Class(Character.Class).Name
                                If Character.Gender = 0 Then
                                    .lblGender = "Male"
                                Else
                                    .lblGender = "Female"
                                End If
                                .lblHP = Character.MaxHP
                                .lblEnergy = Character.MaxEnergy
                                .lblMana = Character.MaxMana
                                .lblStrength = Character.Strength
                                .lblAgility = Character.Agility
                                .lblEndurance = Character.Endurance
                                .lblIntelligence = Character.Intelligence
                                .lblLevel = Character.level
                                If Character.Guild > 0 Then
                                    .lblGuild = "Yes"
                                    .lblGuildRank = Choose(Character.GuildRank + 1, "Initiate", "Member", "Lord", "Founder")
                                Else
                                    .lblGuild = "No"
                                    .lblGuildRank = ""
                                End If
                                .txtDesc = Character.desc
                                BitBlt .picSprite.hdc, 0, 0, 32, 32, hdcSprites, 128, (Character.Sprite - 1) * 32, SRCCOPY
                            End With
                        Else
                            Character.Class = 0
                            frmCharacter.frameCharacter.Visible = False
                        End If
                        frmCharacter.Show
                        
                    Case 4 'Motd
                        If frmCharacter_Loaded = True Then
                            With frmCharacter.txtMotd
                                .Enabled = True
                                .BackColor = QBColor(15)
                                .Text = St
                            End With
                        End If
                        
                    Case 5 'Password Changed
                        If frmWait_Loaded = True Then Unload frmWait
                        frmCharacter.Show
                        
                    Case 6 'Player Joined Game
                        If Len(St) >= 5 Then
                            A = Asc(Mid$(St, 1, 1))
                            With Player(A)
                                .Ignore = False
                                .Sprite = Asc(Mid$(St, 2, 1))
                                .Status = Asc(Mid$(St, 3, 1))
                                .Guild = Asc(Mid$(St, 4, 1))
                                .Name = Mid$(St, 5)
                                If CMap > 0 Then
                                    If .Status = 2 Then
                                        PrintChat "All hail " + .Name + ", a new adventurer in this land!", 3
                                    Else
                                        PrintChat .Name + " has joined the game!", 3
                                    End If
                                End If
                                UpdatePlayerColor A
                            End With
                        End If
                        
                    Case 7 'Player Left Game
                        If Len(St) = 1 Then
                            A = Asc(Mid$(St, 1, 1))
                            If A >= 1 Then
                                With Player(A)
                                    PlayerLeftMap A
                                    .Sprite = 0
                                    PrintChat .Name + " has left the game!", 3
                                End With
                            End If
                        End If
                    
                    Case 8 'Player joined map
                        If Len(St) = 4 Then
                            A = Asc(Mid$(St, 1, 1))
                            With Player(A)
                                .Map = CMap
                                .X = Asc(Mid$(St, 2, 1))
                                .Y = Asc(Mid$(St, 3, 1))
                                .D = Asc(Mid$(St, 4, 1))
                                .XO = .X * 32
                                .YO = .Y * 32
                                .A = 0
                            End With
                        End If
                        
                    Case 9 'Player left map
                        If Len(St) = 1 Then
                            A = Asc(Mid$(St, 1, 1))
                            If A >= 1 Then
                                PlayerLeftMap (A)
                            End If
                        End If
                        
                    Case 10 'Player moved
                        If Len(St) = 5 Then
                            With Player(Asc(Mid$(St, 1, 1)))
                                If .X * 32 = .XO And .Y * 32 = .YO Then
                                    .X = Asc(Mid$(St, 2, 1))
                                    .Y = Asc(Mid$(St, 3, 1))
                                Else
                                    .XO = .X * 32
                                    .YO = .Y * 32
                                    .X = Asc(Mid$(St, 2, 1))
                                    .Y = Asc(Mid$(St, 3, 1))
                                End If
                                .D = Asc(Mid$(St, 4, 1))
                                .WalkStep = Asc(Mid$(St, 5, 1))
                            End With
                        End If
                        
                    Case 11 'Say
                        If Len(St) >= 2 Then
                            A = Asc(Mid$(St, 1, 1))
                            PrintChat Player(A).Name + " says, " + Chr$(34) + SwearFilter(Mid$(St, 2)) + Chr$(34), 7
                        End If
                        
                    Case 12 'You joined map
                        If Len(St) = 14 Then
                            If MapEdit = True Then CloseMapEdit
                            'Destroy Projectiles
                            For A = 1 To 100
                                DestroyEffect (A)
                            Next A
                            If CMap = 0 Then
                                St1 = ""
                                B = 0
                                For A = 1 To 255
                                    With Player(A)
                                        If .Sprite > 0 And A <> Character.Index Then
                                            B = B + 1
                                            St1 = St1 + ", " + .Name
                                        End If
                                    End With
                                Next A
                                If B > 0 Then
                                    St1 = Mid$(St1, 2)
                                    PrintChat "Welcome to the Odyssey Online Classic!  There are " + CStr(B) + " other players online:" + St1, 15
                                Else
                                    PrintChat "Welcome to the Odyssey Online Classic!  There are no other users currently online.", 15
                                End If
                                Load frmMain
                            End If
                            CMap = Asc(Mid$(St, 1, 1)) * 256 + Asc(Mid$(St, 2, 1))
                            CX = Asc(Mid$(St, 3, 1))
                            CY = Asc(Mid$(St, 4, 1))
                            CDir = Asc(Mid$(St, 5, 1))
                            CWalkCode = Asc(Mid$(St, 6, 1))
                            CXO = CX * 32
                            CYO = CY * 32
                            For A = 0 To 5
                                Map.Monster(A).Monster = 0
                            Next A
                            For A = 0 To 49
                                Map.Object(A).Object = 0
                            Next A
                            For A = 0 To 9
                                Map.Door(A).Att = 0
                            Next A
                            For A = 1 To 255
                                Player(A).Map = 0
                            Next A
                            Freeze = True
                            Open MapCacheFile For Random As #1 Len = 1927
                            Get #1, CMap, MapData
                            Close #1
                            LoadMap MapData
                            If Map.Version <> Asc(Mid$(St, 7, 1)) * 16777216 + Asc(Mid$(St, 8, 1)) * 65536 + Asc(Mid$(St, 9, 1)) * 256& + Asc(Mid$(St, 10, 1)) Or CheckSum(MapData) <> Asc(Mid$(St, 11, 1)) * 16777216 + Asc(Mid$(St, 12, 1)) * 65536 + Asc(Mid$(St, 13, 1)) * 256& + Asc(Mid$(St, 14, 1)) Then
                                SendSocket Chr$(45)
                                RequestedMap = True
                            Else
                                RequestedMap = False
                            End If
                        End If
                        
                    Case 13 'Error creating character
                        If frmWait_Loaded = True Then Unload frmWait
                        frmNewCharacter.Show
                        MessageBox frmNewCharacter.hWnd, "That name is already in use, please try another!", TitleString, vbOKOnly + vbExclamation
                        
                    Case 14 'New Map Object
                        If Len(St) = 4 Then
                            A = Asc(Mid$(St, 1, 1))
                            If A <= 49 Then
                                With Map.Object(A)
                                    .Object = Asc(Mid$(St, 2, 1))
                                    .X = Asc(Mid$(St, 3, 1))
                                    .Y = Asc(Mid$(St, 4, 1))
                                    .XOffset = Int(Rnd * 16)
                                    .YOffset = Int(Rnd * 16)
                                    RedrawMapTile CLng(.X), CLng(.Y)
                                End With
                            End If
                        End If
                        
                    Case 15 'Erase Map Object
                        If Len(St) = 1 Then
                            A = Asc(Mid$(St, 1, 1))
                            If A <= 49 Then
                                With Map.Object(A)
                                    .Object = 0
                                    RedrawMapTile CLng(.X), CLng(.Y)
                                End With
                            End If
                        End If
                        
                    Case 16 'Error messages
                        If Len(St) >= 1 Then
                            Select Case Asc(Mid$(St, 1, 1))
                                Case 0 'Custom
                                    If Len(St) > 2 Then
                                        PrintChat Mid$(St, 2), 7
                                    End If
                                Case 1 'Inv full
                                    PrintChat "Your inventory is full!", 7
                                Case 2 'Map full
                                    PrintChat "There is too much already on the ground here to drop that.", 7
                                Case 3 'No such object
                                    PrintChat "No such object.", 7
                                Case 4 'No such player
                                    PrintChat "No such player.", 7
                                Case 5 'No such monster
                                    PrintChat "No such monster.", 7
                                Case 6 'Player is too far away
                                    PrintChat "Player is too far away.", 7
                                Case 7 'Monster is too far away
                                    PrintChat "Monster is too far away.", 7
                                Case 8 'You cannot use that
                                    PrintChat "You cannot use that object.", 7
                                Case 9 'Friendly Zone - can't attack
                                    PrintChat "This is a friendly area, you cannot attack here!", 7
                                Case 10 'Cannot attack immortal
                                    PrintChat "You may not attack an immortal!", 7
                                Case 11 'You are an immortal
                                    PrintChat "Immortals may not attack other players!", 7
                                Case 12 'Can't attack monsters here
                                    PrintChat "You cannot attack these monsters!", 7
                                Case 13 'Ban list full
                                    PrintChat "The ban list is full!", 7
                                Case 14 'Not invited to join
                                    PrintChat "You have not been invited to join any guild.", 7
                                Case 15 'Not enough cash
                                    PrintChat "You do not have enough gold to do that!", 7
                                Case 16 'Guild name in use
                                    PrintChat "That name is already used either by another player or guild.  Please try another.", 7
                                Case 17 'Guild full
                                    PrintChat "That guild is full!", 7
                                Case 18 'too many guilds
                                    PrintChat "Too many guilds already exist.  You may join another guild or try again later.", 7
                                Case 19 'cannot attack player -- he is not in guild
                                    PrintChat "That player is not in a guild -- you may not attack non-guild players.", 7
                                Case 20 'cannot attack player -- you are not in guild
                                    PrintChat "You must be a member of a guild to attack other players.", 7
                                Case 21 'not in a hall
                                    PrintChat "You are not in a guild hall!", 7
                                Case 22 'hall already owned
                                    PrintChat "This hall is already owned by another guild.", 7
                                Case 23 'already have hall
                                    PrintChat "Your guild already owns a hall.  You must move out of your old hall before you may purchase a new one.", 7
                                Case 24 'don't have enough money to buy hall
                                    PrintChat "Your guild does not have enough money in its bank account to buy this hall.  Type /guild hallinfo for the price information of this hall.", 7
                                Case 25 'do not own a guild hall
                                    PrintChat "Your guild does not own a hall.", 7
                                Case 26 'need 5 members
                                    PrintChat "You must have atleast 5 members in your guild before you may do that.", 7
                                Case 27 'Can't afford that
                                    PrintChat "You do not have the items required to purchase that!", 7
                                Case 28 'Not in a bank
                                    PrintChat "You are not in a bank!", 7
                                Case 29 'too far away
                                    PrintChat "That player is too far away to hit!", 7
                                Case 30 'must be level 5 to join guild
                                    PrintChat "You must be at least level 5 to join a guild!", 7
                                Case 31 'Invalid Admin Password
                                    PrintChat "Incorrect Administration Password to perform action!", 7
                                Case 32 'Must be in a smithy shop
                                    PrintChat "You are not in a blacksmithy shop!", 7
                                Case 33 'Do not have enough money
                                    PrintChat "You do not have enough money to repair this object!", 7
                                Case 34 'Do not have specified object
                                    PrintChat "You do not have the object to be repaired!", 7
                            End Select
                        End If
                        
                    Case 17 'New Inv Object
                        If Len(St) = 6 Then
                            A = Asc(Mid$(St, 1, 1))
                            If A >= 1 And A <= 20 Then
                                With Character.Inv(A)
                                    .Object = Asc(Mid$(St, 2, 1))
                                    .Value = Asc(Mid$(St, 3, 1)) * 16777216 + Asc(Mid$(St, 4, 1)) * 65536 + Asc(Mid$(St, 5, 1)) * 256& + Asc(Mid$(St, 6, 1))
                                    .EquippedNum = 0
                                End With
                                DrawInvObject A
                            End If
                        End If
                        
                    Case 18 'Erase Inv Object
                        If Len(St) = 1 Then
                            A = Asc(Mid$(St, 1, 1))
                            If A >= 1 And A <= 20 Then
                                With Character.Inv(A)
                                    .Object = 0
                                    .EquippedNum = 0
                                End With
                                DrawInvObject A
                            End If
                        End If
                        
                    Case 19 'Use Object
                        If Len(St) = 2 Then
                            A = Asc(Mid$(St, 1, 1))
                            If A >= 1 And A <= 20 Then
                                Character.Inv(A).EquippedNum = Asc(Mid$(St, 2, 1))
                                DrawInvObject A
                            End If
                        End If
                        
                    Case 20 'Stop using object
                        If Len(St) = 1 Then
                            A = Asc(Mid$(St, 1, 1))
                            If A >= 1 And A <= 20 Then
                                Character.Inv(A).EquippedNum = False
                                DrawInvObject A
                            End If
                        End If
                        
                    Case 21 'Map Data
                        If Len(St) = 1927 Then
                            MapData = St
                            Open MapCacheFile For Random As #1 Len = 1927
                            Put #1, CMap, MapData
                            Close #1
                            LoadMap MapData
                            ShowMap
                        End If
                        
                    Case 22 'Done Sending Map
                        If RequestedMap = False Then
                            ShowMap
                        End If
                        
                    Case 24 'Joined Game
                        If frmWait_Loaded = True Then
                            frmWait.lblStatus = "Loading Game ..."
                            frmWait.lblStatus.Refresh
                        End If
                        For A = 1 To 255
                            Player(A).Map = 0
                        Next A
                        Load frmMain
                        DrawHP
                        DrawEnergy
                        DrawMana
                        DrawMap
                        blnPlaying = True
                        
                    Case 25 'Tell
                        If Len(St) >= 2 Then
                            A = Asc(Mid$(St, 1, 1))
                            If A >= 1 Then
                                With Player(A)
                                    If .Ignore = False Then
                                        PrintChat .Name + " tells you, " + Chr$(34) + SwearFilter(Mid$(St, 2)) + Chr$(34), 10
                                    End If
                                End With
                            End If
                        End If
                        
                    Case 26 'Broadcast
                        If Len(St) >= 2 And Options.Broadcasts = True Then
                            A = Asc(Mid$(St, 1, 1))
                            If A >= 1 Then
                                If Mid$(St, 2) = Player(A).LastMessage Then
                                    Player(A).RepCount = Player(A).RepCount + 1
                                Else
                                    Player(A).RepCount = 0
                                    Player(A).LastMessage = Mid$(St, 2)
                                End If
                                
                                PrintChat Player(A).Name + ": " + SwearFilter(Mid$(St, 2)), 13
                            End If
                        End If
                        
                    Case 27 'Emote
                        If Len(St) >= 2 Then
                            A = Asc(Mid$(St, 1, 1))
                            PrintChat Player(A).Name + " " + SwearFilter(Mid$(St, 2)), 11
                        End If
                        
                    Case 28 'Yell
                        If Len(St) >= 2 Then
                            A = Asc(Mid$(St, 1, 1))
                            PrintChat Player(A).Name + " yells, " + Chr$(34) + SwearFilter(Mid$(St, 2)) + Chr$(34), 7
                        End If
                        
                    Case 29 'Map Changed
                        If Len(St) = 1 Then
                            PrintChat "This map has been altered by " + Player(A).Name + ".", 14
                        End If
                        
                    Case 30 'Server Message
                        If Len(St) > 0 Then
                            PrintChat "Server Message: " + St, 9
                        End If
                        
                    Case 31 'Object Data
                        If Len(St) >= 3 Then
                            A = Asc(Mid$(St, 1, 1))
                            If A >= 1 Then
                                With Object(A)
                                    .Picture = Asc(Mid$(St, 2, 1))
                                    .Type = Asc(Mid$(St, 3, 1))
                                    If Len(St) >= 4 Then
                                        .Name = Mid$(St, 4)
                                    Else
                                        .Name = ""
                                    End If
                                    If frmMonster_Loaded = True Then
                                        frmMonster.cmbObject(0).List(A) = CStr(A) + ": " + .Name
                                        frmMonster.cmbObject(1).List(A) = CStr(A) + ": " + .Name
                                        frmMonster.cmbObject(2).List(A) = CStr(A) + ": " + .Name
                                    End If
                                    If frmNPC_Loaded = True Then
                                        frmNPC.cmbGiveObject.List(A) = CStr(A) + ": " + .Name
                                        frmNPC.cmbTakeObject.List(A) = CStr(A) + ": " + .Name
                                    End If
                                    If frmList_Loaded = True Then
                                        frmList.lstObjects.List(A - 1) = CStr(A) + ": " + .Name
                                    End If
                                End With
                                For B = 1 To 20
                                    If Character.Inv(B).Object = A Then
                                        DrawInvObject B
                                    End If
                                Next B
                                For B = 0 To 49
                                    With Map.Object(B)
                                        If .Object = A Then
                                            RedrawMapTile CLng(.X), CLng(.Y)
                                        End If
                                    End With
                                Next B
                            End If
                        End If
                        
                    Case 32 'Monster Data
                        If Len(St) >= 2 Then
                            A = Asc(Mid$(St, 1, 1))
                            If A >= 1 Then
                                With Monster(A)
                                    .Sprite = Asc(Mid$(St, 2, 1))
                                    If Len(St) >= 3 Then
                                        .Name = Mid$(St, 3)
                                    Else
                                        .Name = ""
                                    End If
                                    If frmList_Loaded = True Then
                                        frmList.lstMonsters.List(A - 1) = CStr(A) + ": " + .Name
                                    End If
                                    If frmMapProperties_Loaded = True Then
                                        frmMapProperties.cmbMonster(0).List(A) = CStr(A) + ": " + .Name
                                        frmMapProperties.cmbMonster(1).List(A) = CStr(A) + ": " + .Name
                                        frmMapProperties.cmbMonster(2).List(A) = CStr(A) + ": " + .Name
                                    End If
                                End With
                            End If
                        End If
                        
                    Case 33 'Edit Object Data
                        If Len(St) = 6 Then
                            A = Asc(Mid$(St, 1, 1))
                            If frmObject_Loaded = False Then Load frmObject
                            With frmObject
                                .lblNumber = A
                                .txtName = Object(A).Name
                                If Object(A).Picture > 0 Then
                                    .sclPicture = Object(A).Picture
                                Else
                                    .sclPicture = 1
                                End If
                                .ObjData(0) = Asc(Mid$(St, 3, 1))
                                .ObjData(1) = Asc(Mid$(St, 4, 1))
                                .ObjData(2) = Asc(Mid$(St, 5, 1))
                                .ObjData(3) = Asc(Mid$(St, 6, 1))
                                .lblData(0) = .ObjData(0)
                                .lblData(1) = .ObjData(1)
                                .lblData(2) = .ObjData(2)
                                .lblData(3) = .ObjData(3)
                                If Object(A).Type < .cmbType.ListCount Then
                                    .cmbType.ListIndex = 0
                                    .cmbType.ListIndex = Object(A).Type
                                Else
                                    .cmbType.ListIndex = 0
                                End If
                                .Show 1
                            End With
                        End If
                        
                    Case 34 'Edit Monster Data
                        If Len(St) = 14 Then
                            A = Asc(Mid$(St, 1, 1))
                            If frmMonster_Loaded = False Then Load frmMonster
                            With frmMonster
                                .lblNumber = A
                                .txtName = Monster(A).Name
                                If Monster(A).Sprite > 0 Then
                                    .sclSprite = Monster(A).Sprite
                                Else
                                    .sclSprite = 1
                                End If
                                B = Asc(Mid$(St, 2, 1))
                                If B > 0 Then .sclHP = B Else .sclHP = 1
                                B = Asc(Mid$(St, 3, 1))
                                If B > 0 Then .sclStrength = B Else .sclStrength = 1
                                B = Asc(Mid$(St, 4, 1))
                                .sclArmor = B
                                B = Asc(Mid$(St, 5, 1))
                                If B > 0 Then .sclSpeed = B Else .sclSpeed = 1
                                B = Asc(Mid$(St, 6, 1))
                                If B > 0 Then .sclSight = B Else .sclSight = 1
                                B = Asc(Mid$(St, 7, 1))
                                If B <= 100 Then .sclAgility = B Else .sclAgility = 100
                                B = Asc(Mid$(St, 8, 1))
                                For C = 0 To 3
                                    If ExamineBit(CByte(B), CByte(C)) = True Then
                                        .chkFlag(C) = 1
                                    Else
                                        .chkFlag(C) = 0
                                    End If
                                Next C
                                .cmbObject(0).ListIndex = Asc(Mid$(St, 9, 1))
                                .txtValue(0) = Asc(Mid$(St, 10, 1))
                                .cmbObject(1).ListIndex = Asc(Mid$(St, 11, 1))
                                .txtValue(1) = Asc(Mid$(St, 12, 1))
                                .cmbObject(2).ListIndex = Asc(Mid$(St, 13, 1))
                                .txtValue(2) = Asc(Mid$(St, 14, 1))
                                .Show 1
                            End With
                        End If
                    Case 35 'Repeat
                        If Len(St) >= 1 Then
                            SendSocket St
                        End If
                        
                    Case 36 'Door Open
                        If Len(St) = 3 Then
                            A = Asc(Mid$(St, 1, 1))
                            If A <= 9 Then
                                With Map.Door(A)
                                    .X = Asc(Mid$(St, 2, 1))
                                    .Y = Asc(Mid$(St, 3, 1))
                                    .Att = Map.Tile(.X, .Y).Att
                                    .BGTile1 = Map.Tile(.X, .Y).BGTile1
                                    Map.Tile(.X, .Y).Att = 0
                                    Map.Tile(.X, .Y).BGTile1 = 0
                                    RedrawMapTile .X, .Y
                                End With
                            End If
                        End If
                        
                    Case 37 'Close Door
                        If Len(St) = 1 Then
                            A = Asc(Mid$(St, 1, 1))
                            If A <= 9 Then
                                With Map.Door(A)
                                    Map.Tile(.X, .Y).Att = .Att
                                    Map.Tile(.X, .Y).BGTile1 = .BGTile1
                                    .Att = 0
                                    RedrawMapTile .X, .Y
                                End With
                            End If
                        End If
                        
                    Case 38 'New Map Monster
                        If Len(St) = 5 Then
                            A = Asc(Mid$(St, 1, 1))
                            If A <= 5 Then
                                With Map.Monster(A)
                                    .Monster = Asc(Mid$(St, 2, 1))
                                    .X = Asc(Mid$(St, 3, 1))
                                    .Y = Asc(Mid$(St, 4, 1))
                                    .D = Asc(Mid$(St, 5, 1))
                                    .XO = .X * 32
                                    .YO = .Y * 32
                                    .A = 0
                                End With
                            End If
                        End If
                        
                    Case 39 'Monster Die
                        If Len(St) = 1 Then
                            A = Asc(Mid$(St, 1, 1))
                            If A <= 5 Then
                                PlayWav 9
                                Map.Monster(A).Monster = 0
                                MonsterDied A
                            End If
                        End If
                        
                    Case 40 'Monster Move
                        If Len(St) = 4 Then
                            A = Asc(Mid$(St, 1, 1))
                            If A <= 5 Then
                                With Map.Monster(A)
                                    If CLng(.X) * 32 <> .XO Or CLng(.Y) * 32 <> .YO Then
                                        .X = Asc(Mid$(St, 2, 1))
                                        .Y = Asc(Mid$(St, 3, 1))
                                        .XO = .X * 32
                                        .YO = .Y * 32
                                    Else
                                        .X = Asc(Mid$(St, 2, 1))
                                        .Y = Asc(Mid$(St, 3, 1))
                                    End If
                                    .D = Asc(Mid$(St, 4, 1))
                                End With
                            End If
                        End If
                        
                    Case 41 'Monster Attack
                        If Len(St) = 1 Then
                            A = Asc(Mid$(St, 1, 1))
                            If A <= 5 Then
                                Map.Monster(A).A = 5
                            End If
                        End If
                        
                    Case 42 'Player Attack
                        If Len(St) = 1 Then
                            A = Asc(Mid$(St, 1, 1))
                            If A >= 1 Then
                                Player(A).A = 5
                            End If
                        End If
                        
                    Case 43 'You hit player
                        If Len(St) = 3 Then
                            A = Asc(Mid$(St, 2, 1))
                            If A >= 1 Then
                                B = Asc(Mid$(St, 3, 1))
                                Select Case Asc(Mid$(St, 1, 1))
                                    Case 0
                                        PlayWav 2
                                        If B > 0 Then
                                            PrintInfoText "You hit " + Player(A).Name + " for " + CStr(B) + " hitpoints."
                                        Else
                                            PrintInfoText "Your attack does nothing against " + Player(A).Name + "."
                                        End If
                                        CAttack = 5
                                    Case 1
                                        PlayWav 3
                                        PrintInfoText "You missed " + Player(A).Name + "."
                                End Select
                            End If
                        End If
                        
                    Case 44 'You hit monster
                        If Len(St) = 3 Then
                            A = Asc(Mid$(St, 2, 1))
                            If A >= 0 And A <= 5 Then
                                If Map.Monster(A).Monster > 0 Then
                                    B = Asc(Mid$(St, 3, 1))
                                    Select Case Asc(Mid$(St, 1, 1))
                                        Case 0
                                            PlayWav 2
                                            If B > 0 Then
                                                PrintInfoText "You hit the " + Monster(Map.Monster(A).Monster).Name + " for " + CStr(B) + " hitpoints."
                                            Else
                                                PrintInfoText "Your attack does nothing against the " + Monster(Map.Monster(A).Monster).Name + "."
                                            End If
                                            CAttack = 5
                                        Case 1
                                            PlayWav 3
                                            PrintInfoText "You missed the " + Monster(Map.Monster(A).Monster).Name + "."
                                    End Select
                                End If
                            End If
                        End If
                        
                    Case 45 'You killed player
                        If Len(St) = 5 Then
                            A = Asc(Mid$(St, 1, 1))
                            If A >= 1 Then
                                PlayWav 8
                                For C = 1 To 100
                                    If Projectile(C).TargetNum = A Then
                                        DestroyEffect (C)
                                    End If
                                Next C
                                With Player(A)
                                    If .Status = 1 Then
                                        PrintChat "You have put the evil murderer " + .Name + " to justice!", 12
                                        .Status = 0
                                    Else
                                        PrintChat "You have murdered " + .Name + " in cold blood!", 12
                                        Character.Status = 1
                                    End If
                                    B = Asc(Mid$(St, 2, 1)) * 16777216 + Asc(Mid$(St, 3, 1)) * 65536 + Asc(Mid$(St, 4, 1)) * 256& + Asc(Mid$(St, 5, 1))
                                    PrintInfoText "You have killed " + .Name + "!"
                                    PrintInfoText "You have gained " + CStr(B) + " exp."
                                    Character.Experience = Character.Experience + B
                                End With
                            End If
                        End If
                        
                    Case 46 'Change HP
                        If Len(St) = 1 Then
                            Character.HP = Asc(Mid$(St, 1, 1))
                            DrawHP
                        End If
                        
                    Case 47 'Change Energy
                        If Len(St) = 1 Then
                            Character.Energy = Asc(Mid$(St, 1, 1))
                            DrawEnergy
                        End If
                        
                    Case 48 'Change Mana
                        If Len(St) = 1 Then
                            Character.Mana = Asc(Mid$(St, 1, 1))
                            DrawMana
                        End If
                        
                    Case 49 'Player Hit You
                        If Len(St) = 3 Then
                            A = Asc(Mid$(St, 2, 1))
                            If A >= 1 Then
                                B = Asc(Mid$(St, 3, 1))
                                Select Case Asc(Mid$(St, 1, 1))
                                    Case 0
                                        PlayWav 2
                                        If B > 0 Then
                                            PrintInfoText Player(A).Name + " hit you for " + CStr(B) + " hitpoints."
                                            If Character.HP > B Then
                                                Character.HP = Character.HP - B
                                            Else
                                                Character.HP = 0
                                            End If
                                            DrawHP
                                        Else
                                            PrintInfoText Player(A).Name + "'s attack does not even scratch you."
                                        End If
                                        Player(A).A = 5
                                    Case 1
                                        PlayWav 3
                                        PrintInfoText Player(A).Name + " misses you."
                                End Select
                            End If
                        End If
                        
                    Case 50 'Monster Hit You
                        If Len(St) = 3 Then
                            A = Asc(Mid$(St, 2, 1))
                            If A <= 5 Then
                                B = Map.Monster(A).Monster
                                C = Asc(Mid$(St, 3, 1))
                                Select Case Asc(Mid$(St, 1, 1))
                                    Case 0
                                        PlayWav 2
                                        If C > 0 Then
                                            If B > 0 Then
                                                PrintInfoText "The " + Monster(B).Name + " hit you for " + CStr(C) + " hitpoints."
                                            Else
                                                PrintInfoText "A monster hit you for " + CStr(C) + " hitpoints."
                                            End If
                                            If Character.HP > C Then
                                                Character.HP = Character.HP - C
                                            Else
                                                Character.HP = 0
                                            End If
                                            DrawHP
                                        Else
                                            If B > 0 Then
                                                PrintInfoText "The " + Monster(B).Name + "'s attack does not even scratch you."
                                            Else
                                                PrintInfoText "The monster's attack does not even scratch you."
                                            End If
                                        End If
                                        Map.Monster(A).A = 5
                                    Case 1
                                        PlayWav 3
                                        If B > 0 Then
                                            PrintInfoText "The " + Monster(B).Name + " misses you."
                                        Else
                                            PrintInfoText "The monster misses you."
                                        End If
                                End Select
                            End If
                        End If
                        
                    Case 51 'You killed the monster
                        If Len(St) = 5 Then
                            A = Asc(Mid$(St, 1, 1))
                            If A <= 5 Then
                                PlayWav 9
                                With Map.Monster(A)
                                    If .Monster > 0 Then
                                        PrintInfoText "You have killed the " + Monster(.Monster).Name + "!"
                                    Else
                                        PrintInfoText "You have killed the monster!"
                                    End If
                                    .Monster = 0
                                End With
                                B = Asc(Mid$(St, 2, 1)) * 16777216 + Asc(Mid$(St, 3, 1)) * 65536 + Asc(Mid$(St, 4, 1)) * 256& + Asc(Mid$(St, 5, 1))
                                PrintInfoText "You have gained " + CStr(B) + " exp."
                                Character.Experience = Character.Experience + B
                                MonsterDied (A)
                            End If
                        End If
                        
                    Case 52 'Player Killed You
                        If Len(St) = 1 Then
                            A = Asc(Mid$(St, 1, 1))
                            If A >= 1 Then
                                With Player(A)
                                    If Character.Status = 1 Then
                                        PrintChat .Name + " has put you to justice!  You have lost 1/3 of your experience!", 12
                                        Character.Status = 0
                                    Else
                                        PrintChat .Name + " has murdered you in cold blood!  You have lost 1/3 of your experience!", 12
                                        .Status = 1
                                    End If
                                End With
                                YouDied
                            End If
                        End If
                    
                    Case 53 'Monster Killed You
                        If Len(St) = 1 Then
                            A = Asc(Mid$(St, 1, 1))
                            If A >= 1 Then
                                NextTransition = 6
                                PlayWav 8
                                With Monster(A)
                                    PrintChat "The " + .Name + " has killed you!  You have lost 1/3 of your experience!", 12
                                End With
                                If Character.Status = 1 Then Character.Status = 0
                                YouDied
                            End If
                        End If
                        
                    Case 54 'Day Time
                        blnNight = False
                        If CMap > 0 Then
                            PrintChat "It is now day time...", 13
                            If ExamineBit(Map.Flags, 1) = False Then DrawMap
                        End If
                        
                    Case 55 'Night Time
                        blnNight = True
                        If CMap > 0 Then
                            PrintChat "It is now night time...", 13
                            If ExamineBit(Map.Flags, 1) = False Then DrawMap
                        End If
                        
                    Case 56 'Text
                        If Len(St) >= 2 Then
                            PrintChat Mid$(St, 2), Asc(Mid$(St, 1, 1))
                        End If
                    
                    Case 57 'Object Breaks
                        If Len(St) = 1 Then
                            A = Asc(Mid$(St, 1, 1))
                            If A >= 1 And A <= 20 Then
                                With Character.Inv(A)
                                    If .Object > 0 Then
                                        PrintInfoText "Your " + Object(.Object).Name + " breaks."
                                    End If
                                    .Object = 0
                                    .EquippedNum = 0
                                End With
                                DrawInvObject A
                            End If
                        End If
                    
                    Case 58 'Ping
                        SendSocket Chr$(29) 'Pong
                        
                    Case 59 'Level Up
                        If Len(St) = 5 Then
                            With Character
                                .level = .level + 1
                                .MaxHP = Asc(Mid$(St, 1, 1))
                                .MaxEnergy = Asc(Mid$(St, 2, 1))
                                .MaxMana = Asc(Mid$(St, 3, 1))
                                .StatPoints = Asc(Mid$(St, 4, 1)) * 256 + Asc(Mid$(St, 5, 1))
                                .Experience = 0
                                PrintChat "Level Up!  You are now level " + CStr(.level) + ".  You have " + CStr(.StatPoints) + " stat points.  Type /train to spend them.", 12
                                If .level = 5 Then
                                    PrintChat "You are no longer protected from other players.", 12
                                End If
                                DrawHP
                                DrawEnergy
                                DrawMana
                            End With
                        End If
                        
                    Case 60 'Experience Change
                        If Len(St) = 4 Then
                            Character.Experience = Asc(Mid$(St, 1, 1)) * 16777216 + Asc(Mid$(St, 2, 1)) * 65536 + Asc(Mid$(St, 3, 1)) * 256& + Asc(Mid$(St, 4, 1))
                        End If
                        
                    Case 61 'Player killed by player
                        If Len(St) = 2 Then
                            A = Asc(Mid$(St, 1, 1))
                            B = Asc(Mid$(St, 2, 1))
                            If A >= 1 And B >= 1 Then
                                With Player(A)
                                    If .Status = 1 Then
                                        If .Map = CMap Then
                                            PlayWav 8
                                            PrintChat Player(B).Name + " has put " + .Name + " to justice!", 12
                                        Else
                                            PrintChat .Name + " has been put to justice!", 12
                                        End If
                                        .Status = 0
                                    Else
                                        If .Map = CMap Then
                                            PlayWav 8
                                            PrintChat Player(B).Name + " has murdered " + .Name + " in cold blood!", 12
                                        Else
                                            PrintChat .Name + " has been murdered in cold blood!", 12
                                        End If
                                        Player(B).Status = 1
                                    End If
                                End With
                            End If
                        End If
                        
                    Case 62 'Player killed by monster
                        If Len(St) = 2 Then
                            A = Asc(Mid$(St, 1, 1))
                            B = Asc(Mid$(St, 2, 1))
                            If A >= 1 And B >= 1 Then
                                With Player(A)
                                    If .Map = CMap Then
                                        PlayWav 8
                                        PrintChat .Name + " has been killed by a " + Monster(B).Name + "!", 12
                                    End If
                                    If .Status = 1 Then .Status = 0
                                End With
                            End If
                        End If
                    
                    Case 63 'Player Sprite Changed
                        If Len(St) = 2 Then
                            A = Asc(Mid$(St, 1, 1))
                            B = Asc(Mid$(St, 2, 1))
                            If A >= 1 And B >= 1 Then
                                If A = Character.Index Then
                                    Character.Sprite = B
                                Else
                                    If Player(A).Sprite > 0 Then
                                        Player(A).Sprite = B
                                    End If
                                End If
                            End If
                        End If
                        
                    Case 64 'Player Name Change
                        If Len(St) >= 2 Then
                            A = Asc(Mid$(St, 1, 1))
                            If A >= 1 Then
                                If A = Character.Index Then
                                    Character.Name = Mid$(St, 2)
                                Else
                                    If Player(A).Sprite > 0 Then
                                        Player(A).Name = Mid$(St, 2)
                                    End If
                                End If
                            End If
                        End If
                        
                    Case 65 'Changed access
                        If Len(St) = 1 Then
                            Character.Access = Asc(Mid$(St, 1, 1))
                            If Character.Access > 0 Then
                                Character.Status = 3
                            Else
                                If Character.Status = 0 Then
                                    Character.Status = 0
                                End If
                            End If
                        End If
                        
                    Case 66 'Player banned
                        If Len(St) >= 2 Then
                            B = Asc(Mid$(St, 1, 1))
                            C = Asc(Mid$(St, 2, 1))
                            If B >= 1 And C >= 1 Then
                                If C >= 1 Then
                                    If Len(St) > 2 Then
                                        PrintChat Player(B).Name + " has been banned by " + Player(C).Name + ": " + Mid$(St, 3), 15
                                    Else
                                        PrintChat Player(B).Name + " has been banned by " + Player(C).Name + "!", 15
                                    End If
                                Else
                                    If Len(St) > 2 Then
                                        PrintChat Player(B).Name + " has been banned: " + Mid$(St, 3), 15
                                    Else
                                        PrintChat Player(B).Name + " has been banned!", 15
                                    End If
                                End If
                            End If
                        End If
                        
                    Case 67 'Booted
                        If Len(St) >= 1 Then
                            A = Asc(Mid$(St, 1, 1))
                            If A >= 1 Then
                                If Len(St) > 1 Then
                                    MessageBox frmMenu.hWnd, "You have been booted from The Odyssey by " + Player(A).Name + ": " + Mid$(St, 2), TitleString, vbOKOnly + vbExclamation
                                Else
                                    MessageBox frmMenu.hWnd, "You have been booted from The Odyssey by " + Player(A).Name + "!", TitleString, vbOKOnly + vbExclamation
                                End If
                            Else
                                If Len(St) > 1 Then
                                    MessageBox frmMenu.hWnd, "You have been booted from The Odyssey: " + Mid$(St, 2), TitleString, vbOKOnly + vbExclamation
                                Else
                                    MessageBox frmMenu.hWnd, "You have been booted from The Odyssey!", TitleString, vbOKOnly + vbExclamation
                                End If
                            End If
                            CloseClientSocket 0
                        End If
                        
                    Case 68 'Player Booted
                        If Len(St) >= 2 Then
                            A = Asc(Mid$(St, 1, 1))
                            B = Asc(Mid$(St, 2, 1))
                            If A >= 1 Then
                                If B >= 1 Then
                                    If Len(St) > 2 Then
                                        PrintChat Player(A).Name + " has been booted by " + Player(B).Name + ": " + Mid$(St, 3), 15
                                    Else
                                        PrintChat Player(A).Name + " has been booted by " + Player(B).Name + "!", 15
                                    End If
                                Else
                                    If Len(St) > 2 Then
                                        PrintChat Player(A).Name + " has been booted: " + Mid$(St, 3), 15
                                    Else
                                        PrintChat Player(A).Name + " has been booted!", 15
                                    End If
                                End If
                            End If
                        End If
                        
                    Case 69 'Ban List
                        If Len(St) >= 2 Then
                            A = Asc(Mid$(St, 1, 1))
                            With frmList.lstBans
                                .AddItem CStr(A) + ": " + Mid$(St, 2)
                                .ItemData(.ListCount - 1) = A
                            End With
                        Else
                            frmList.Show
                        End If
                        
                    Case 70 'Guild Data
                        If Len(St) >= 1 Then
                            A = Asc(Mid$(St, 1, 1))
                            If A >= 1 Then
                                If Len(St) > 1 Then
                                    Guild(A).Name = Mid$(St, 2)
                                Else
                                    Guild(A).Name = ""
                                End If
                            End If
                        End If
                        
                    Case 71 'Guild Dec. Data
                        If Len(St) = 3 Then
                            A = Asc(Mid$(St, 1, 1))
                            If A <= 4 Then
                                With Character.GuildDeclaration(A)
                                    .Guild = Asc(Mid$(St, 2, 1))
                                    .Type = Asc(Mid$(St, 3, 1))
                                End With
                                UpdatePlayersColors
                            End If
                        End If
                        
                    Case 72 'Guild Change
                        If Len(St) = 1 Then
                            A = Asc(Mid$(St, 1, 1))
                            If A > 0 Then
                                PrintChat "You are now a member of " + Chr$(34) + Guild(A).Name + Chr$(34), 15
                            Else
                                If Character.Guild > 0 Then
                                    PrintChat "You are no longer a member of " + Chr$(34) + Guild(Character.Guild).Name + Chr$(34), 15
                                End If
                            End If
                            Character.Guild = A
                            Character.GuildRank = 0
                            UpdatePlayersColors
                        End If
                        
                    Case 73 'Player Changed Guild
                        If Len(St) = 2 Then
                            A = Asc(Mid$(St, 1, 1))
                            B = Asc(Mid$(St, 2, 1))
                            If A >= 1 Then
                                If Player(A).Guild = Character.Guild And Character.Guild > 0 Then
                                    PrintChat Player(A).Name + " is no longer a member of your guild.", 15
                                End If
                                Player(A).Guild = B
                                If B > 0 And B = Character.Guild Then
                                    PrintChat Player(A).Name + " is now a member of your guild.", 15
                                End If
                            End If
                            UpdatePlayerColor A
                        End If
                        
                    Case 74 'Guild Account Status
                        If Len(St) = 4 Then
                            A = Asc(Mid$(St, 1, 1)) * 16777216 + Asc(Mid$(St, 2, 1)) * 65536 + Asc(Mid$(St, 3, 1)) * 256& + Asc(Mid$(St, 4, 1))
                            PrintChat "Your guild has " + CStr(A) + " gold in the bank.", 15
                        ElseIf Len(St) = 8 Then
                            A = Asc(Mid$(St, 1, 1)) * 16777216 + Asc(Mid$(St, 2, 1)) * 65536 + Asc(Mid$(St, 3, 1)) * 256& + Asc(Mid$(St, 4, 1))
                            B = Asc(Mid$(St, 5, 1)) * 16777216 + Asc(Mid$(St, 6, 1)) * 65536 + Asc(Mid$(St, 7, 1)) * 256& + Asc(Mid$(St, 8, 1))
                            PrintChat "Your guild owes " + CStr(A) + " gold.  This must be payed before " + CStr(CDate(B)) + " or your guild will be disbanded.  Type '/guild pay <amount>' to pay toward the debt.", 15
                        End If
                        
                    Case 75 'Guild Deleted
                        If Len(St) = 1 Then
                            Select Case Asc(Mid$(St, 1, 1))
                                Case 0
                                    PrintChat "Your guild has failed to pay its debt in time and has been disbanded!", 15
                                Case 1
                                    PrintChat "Your guild member count has fallen below three -- your guild has been disbanded!", 15
                                Case 2
                                    PrintChat "Your guild has been disbanded!", 15
                                Case 3
                                    PrintChat "Your guild has been disbanded by a god!", 15
                            End Select
                            Character.Guild = 0
                            Character.GuildRank = 0
                            UpdatePlayersColors
                        End If
                        
                    Case 76 'Rank Changed
                        If Len(St) = 1 Then
                            Character.GuildRank = Asc(Mid$(St, 1, 1))
                            PrintChat "Your guild rank has been changed to " + Chr$(34) + Choose(Character.GuildRank + 1, "Initiate", "Member", "Lord", "Founder") + Chr$(34) + ".", 15
                        End If
                        
                    Case 77 'Invited to join guild
                        If Len(St) = 2 Then
                            A = Asc(Mid$(St, 1, 1))
                            B = Asc(Mid$(St, 2, 1))
                            If A >= 1 And B >= 1 Then
                                PrintChat "You have been invited to join the guild " + Chr$(34) + Guild(A).Name + Chr$(34) + " by " + Player(B).Name + ".  If you wish to join, type /guild join.  It will cost 1000 gold to join this guild.", 15
                            End If
                        End If
                        
                    Case 78 'View Guild Data
                        If Len(St) >= 12 Then
                            frmMain.picBuy.Visible = False
                            frmMain.picDrop.Visible = False
                            frmMain.picTrain.Visible = False
                            TempVar1 = Asc(Mid$(St, 1, 1))
                            If frmGuild_Loaded = False Then Load frmGuild
                            With frmGuild
                                A = Asc(Mid$(St, 2, 1))
                                If A > 0 Then
                                    .lblHall = Hall(A).Name
                                Else
                                    .lblHall = "<none>"
                                End If
                                .lstDeclarations.Clear
                                For A = 0 To 4
                                    B = Asc(Mid$(St, 3 + 2 * A))
                                    If B > 0 Then
                                        If Asc(Mid$(St, 4 + 2 * A)) = 0 Then
                                            .lstDeclarations.AddItem "Declaration of Alliance with " + Guild(B).Name
                                        Else
                                            .lstDeclarations.AddItem "Declaration of War with " + Guild(B).Name
                                        End If
                                        .lstDeclarations.ItemData(.lstDeclarations.ListCount - 1) = A
                                    End If
                                Next A
                                If Len(St) >= 13 Then
                                    GetSections2 Mid$(St, 13)
                                    .lblName = Guild(TempVar1).Name
                                    .lstMembers.Clear
                                    For A = 0 To 19
                                        If Len(Section(A + 1)) >= 2 Then
                                            B = Asc(Mid$(Section(A + 1), 1, 1)) - 1
                                            If B <= 3 Then
                                                .lstMembers.AddItem Mid$(Section(A + 1), 2) + " - " + Choose(B + 1, "Initiate", "Member", "Lord", "Founder")
                                                .lstMembers.ItemData(.lstMembers.ListCount - 1) = A
                                            End If
                                        End If
                                    Next A
                                End If
                                If Character.Guild = TempVar1 And Character.GuildRank >= 2 Then
                                    If Character.GuildRank = 3 Then
                                        .btnDisband.Enabled = True
                                    Else
                                        .btnDisband.Enabled = False
                                    End If
                                    If .lstDeclarations.ListCount < 5 Then
                                        .btnAddDeclaration.Enabled = True
                                    Else
                                        .btnAddDeclaration.Enabled = False
                                    End If
                                    If .lblHall = "<none>" Then
                                        .btnMoveOut.Enabled = False
                                    Else
                                        .btnMoveOut.Enabled = True
                                    End If
                                Else
                                    .btnDisband.Enabled = False
                                    .btnAddDeclaration.Enabled = False
                                    .btnMoveOut.Enabled = False
                                End If
                                .btnRemoveMember.Enabled = False
                                .btnRemoveDeclaration.Enabled = False
                                .btnRank(0).Enabled = False
                                .btnRank(1).Enabled = False
                                .btnRank(2).Enabled = False
                                .btnOk.Enabled = True
                                .Show
                            End With
                        End If
                        
                    Case 79 'Guild Chat
                        If Len(St) >= 2 Then
                            A = Asc(Mid$(St, 1, 1))
                            If A >= 1 Then
                                PrintChat Player(A).Name + " -> Guild: " + Mid$(St, 2), 15
                            End If
                        End If
                    
                    Case 80 'Created Guild
                        If Len(St) = 1 Then
                            A = Asc(Mid$(St, 1, 1))
                            Character.Guild = A
                            Character.GuildRank = 3
                            If A > 0 Then
                                PrintChat "You have created a new guild called " + Chr$(34) + Guild(A).Name + Chr$(34) + ".  To invite other players to your guild, type '/guild invite <player>'.  You must get atleast two other players to join your guild today or your guild will be disbanded.", 15
                            End If
                        End If
                        
                    Case 81 'Guild hall change
                        If Len(St) = 1 Then
                            If Asc(Mid$(St, 1, 1)) = 0 Then
                                PrintChat "Your guild now owns a hall!", 15
                            Else
                                PrintChat "Your guild no longer owns a hall!", 15
                            End If
                        End If
                    
                    Case 82 'Guild hall data
                        If Len(St) >= 1 Then
                            A = Asc(Mid$(St, 1, 1))
                            If A >= 1 Then
                                If Len(St) >= 2 Then
                                    Hall(A).Name = Mid$(St, 2)
                                Else
                                    Hall(A).Name = ""
                                End If
                                If frmList_Loaded = True Then
                                    frmList.lstHalls.List(A - 1) = CStr(A) + ": " + Hall(A).Name
                                End If
                            End If
                        End If
                        
                    Case 83 'Guild Hall Edit Data
                        If Len(St) = 13 Then
                            A = Asc(Mid$(St, 1, 1))
                            If A >= 1 Then
                                If frmHall_Loaded = False Then Load frmHall
                                With frmHall
                                    .lblNumber = A
                                    .txtName = Hall(A).Name
                                    .txtPrice = CStr(Asc(Mid$(St, 2, 1)) * 16777216 + Asc(Mid$(St, 3, 1)) * 65536 + Asc(Mid$(St, 4, 1)) * 256& + Asc(Mid$(St, 5, 1)))
                                    .txtUpkeep = CStr(Asc(Mid$(St, 6, 1)) * 16777216 + Asc(Mid$(St, 7, 1)) * 65536 + Asc(Mid$(St, 8, 1)) * 256& + Asc(Mid$(St, 9, 1)))
                                    B = Asc(Mid$(St, 10, 1)) * 256 + Asc(Mid$(St, 11, 1))
                                    If B < 1 Then B = 1
                                    If B > 2000 Then B = 2000
                                    .sclStartMap = B
                                    B = Asc(Mid$(St, 12, 1))
                                    If B > 11 Then B = 11
                                    .sclStartX = B
                                    B = Asc(Mid$(St, 13, 1))
                                    If B > 11 Then B = 11
                                    .sclStartY = B
                                    .Show
                                End With
                            End If
                        End If
                        
                    Case 84 'Guild Hall Info
                        If Len(St) = 10 Then
                            A = Asc(Mid$(St, 1, 1))
                            If A >= 1 Then
                                B = Asc(Mid$(St, 2, 1))
                                If B > 0 Then
                                    PrintChat "Owned By: " + Guild(B).Name, 15
                                Else
                                    PrintChat "This guild hall is not yet owned!", 15
                                End If
                                A = Asc(Mid$(St, 3, 1)) * 16777216 + Asc(Mid$(St, 4, 1)) * 65536 + Asc(Mid$(St, 5, 1)) * 256& + Asc(Mid$(St, 6, 1))
                                PrintChat "Cost: " + CStr(A) + " gold coins", 15
                                A = Asc(Mid$(St, 7, 1)) * 16777216 + Asc(Mid$(St, 8, 1)) * 65536 + Asc(Mid$(St, 9, 1)) * 256& + Asc(Mid$(St, 10, 1))
                                PrintChat "Upkeep: " + CStr(A) + " gold coins per day", 15
                            End If
                        End If
                        
                    Case 85 'NPC Data
                        If Len(St) >= 1 Then
                            A = Asc(Mid$(St, 1, 1))
                            If A >= 1 Then
                                With NPC(A)
                                    If Len(St) >= 2 Then
                                        .Name = Mid$(St, 2)
                                    Else
                                        .Name = ""
                                    End If
                                    If frmList_Loaded = True Then
                                        frmList.lstNPCs.List(A - 1) = CStr(A) + ": " + .Name
                                    End If
                                    If frmMapProperties_Loaded = True Then
                                        frmMapProperties.cmbNPC.List(A) = CStr(A) + ": " + .Name
                                    End If
                                End With
                            End If
                        End If
                        
                    Case 86 'Buy Data
                        If Len(St) = 100 Then
                            For A = 0 To 9
                                With SaleItem(A)
                                    .GiveObject = Asc(Mid$(St, 1 + A * 10, 1))
                                    .GiveValue = Asc(Mid$(St, 2 + A * 10, 1)) * 16777216 + Asc(Mid$(St, 3 + A * 10, 1)) * 65536 + Asc(Mid$(St, 4 + A * 10, 1)) * 256& + Asc(Mid$(St, 5 + A * 10, 1))
                                    .TakeObject = Asc(Mid$(St, 6 + A * 10, 1))
                                    .TakeValue = Asc(Mid$(St, 7 + A * 10, 1)) * 16777216 + Asc(Mid$(St, 8 + A * 10, 1)) * 65536 + Asc(Mid$(St, 9 + A * 10, 1)) * 256& + Asc(Mid$(St, 10 + A * 10, 1))
                                    If .GiveObject >= 1 And .TakeObject >= 1 Then
                                        St1 = ""
                                        If Object(.GiveObject).Type = 6 Then
                                            St1 = CStr(.GiveValue) + " " + Object(.GiveObject).Name
                                        Else
                                            St1 = "1 " + Object(.GiveObject).Name
                                        End If
                                        St1 = St1 + " in exchange for "
                                        If Object(.TakeObject).Type = 6 Then
                                            St1 = St1 + CStr(.TakeValue) + " " + Object(.TakeObject).Name
                                        Else
                                            St1 = St1 + "1 " + Object(.TakeObject).Name
                                        End If
                                        frmMain.lblItem(A) = St1
                                        frmMain.GivObjPic(A).Cls
                                        TransparentBlt frmMain.GivObjPic(A).hdc, 0, 0, 32, 32, hdcObjects, 0, (Object(.GiveObject).Picture - 1) * 32, hdcObjectsMask
                                        frmMain.lblShopName.Caption = Map.Name
                                    Else
                                        frmMain.lblItem(A) = ""
                                        frmMain.GivObjPic(A).Cls
                                    End If
                                End With
                            Next A
                            frmMain.picBuy.Visible = True
                            frmMain.picTrain.Visible = False
                            frmMain.picDrop.Visible = False
                        End If
                    
                    Case 87 'Edit NPC Data
                        If Len(St) >= 108 Then
                            A = Asc(Mid$(St, 1, 1))
                            If A >= 1 Then
                                If frmNPC_Loaded = False Then Load frmNPC
                                B = Asc(Mid$(St, 2, 1))
                                For C = 0 To 1
                                    If ExamineBit(CByte(B), CByte(C)) = True Then
                                        frmNPC.chkFlag(C) = 1
                                    Else
                                        frmNPC.chkFlag(C) = 0
                                    End If
                                Next C
                                For B = 0 To 9
                                    With SaleItem(B)
                                        .GiveObject = Asc(Mid$(St, 3 + B * 10, 1))
                                        .GiveValue = Asc(Mid$(St, 4 + B * 10, 1)) * 16777216 + Asc(Mid$(St, 5 + B * 10, 1)) * 65536 + Asc(Mid$(St, 6 + B * 10, 1)) * 256& + Asc(Mid$(St, 7 + B * 10, 1))
                                        .TakeObject = Asc(Mid$(St, 8 + B * 10, 1))
                                        .TakeValue = Asc(Mid$(St, 9 + B * 10, 1)) * 16777216 + Asc(Mid$(St, 10 + B * 10, 1)) * 65536 + Asc(Mid$(St, 11 + B * 10, 1)) * 256& + Asc(Mid$(St, 12 + B * 10, 1))
                                    End With
                                    UpdateSaleItem B
                                Next B
                                '103
                                GetSections2 Mid$(St, 103)
                                With frmNPC
                                    .lblNumber = A
                                    .txtName = NPC(A).Name
                                    .txtJoinText = Section(1)
                                    .txtLeaveText = Section(2)
                                    .txtSayText1 = Section(3)
                                    .txtSayText2 = Section(4)
                                    .txtSayText3 = Section(5)
                                    .txtSayText4 = Section(6)
                                    .txtSayText5 = Section(7)
                                    .Show 1
                                End With
                            End If
                        End If
                    
                    Case 88 'NPC Talks
                        If Len(St) >= 2 Then
                            A = Asc(Mid$(St, 1, 1))
                            If A >= 1 Then
                                PrintChat NPC(A).Name + " says, " + Chr$(34) + Mid$(St, 2) + Chr$(34), 7
                            End If
                        End If
                        
                    Case 89 'Bank Balance
                        If Len(St) = 4 Then
                            If Map.NPC >= 1 Then
                                A = Asc(Mid$(St, 1, 1)) * 16777216 + Asc(Mid$(St, 2, 1)) * 65536 + Asc(Mid$(St, 3, 1)) * 256& + Asc(Mid$(St, 4, 1))
                                PrintChat NPC(Map.NPC).Name + " tells you, " + Chr$(34) + "You have " + CStr(A) + " gold coins in the bank." + Chr$(34), 7
                            End If
                        End If
                        
                    Case 90 'God Chat
                        If Len(St) >= 2 Then
                            A = Asc(Mid$(St, 1, 1))
                            If A >= 1 Then
                                PrintChat "<" + Player(A).Name + ">: " + Mid$(St, 2), 11
                            End If
                        End If
                        
                    Case 91 'Status Change
                        If Len(St) = 2 Then
                            A = Asc(Mid$(St, 1, 1))
                            If A >= 1 Then
                                If A = Character.Index Then
                                    Character.Status = Asc(Mid$(St, 2, 1))
                                Else
                                    Player(A).Status = Asc(Mid$(St, 2, 1))
                                End If
                                UpdatePlayerColor A
                            End If
                        End If
                        
                    Case 92 'Edit Ban Data
                        If Len(St) >= 4 Then
                            If frmBan_Loaded = False Then Load frmBan
                            GetSections2 Mid$(St, 3)
                            With frmBan
                                .lblNumber = Asc(Mid$(St, 1, 1))
                                .sclUnban = Asc(Mid$(St, 2, 1))
                                .txtName = Section(1)
                                .txtBanner = Section(2)
                                .txtReason = Section(3)
                                .Show
                            End With
                        End If
                        
                    Case 93 'Gained exp
                        If Len(St) = 4 Then
                            B = Asc(Mid$(St, 1, 1)) * 16777216 + Asc(Mid$(St, 2, 1)) * 65536 + Asc(Mid$(St, 3, 1)) * 256& + Asc(Mid$(St, 4, 1))
                            PrintChat "You have gained " + CStr(B) + " exp.", 12
                        End If
                        
                    Case 94 'Edit Script Data
                        If Len(St) >= 3 Then
                            A = InStr(St, Chr$(0))
                            If A >= 1 Then
                                Load frmScript
                                With frmScript
                                    .lblName = Left$(St, A - 1)
                                    .txtCode = Mid$(St, A + 1)
                                    If .txtCode = "" Then
                                        St1 = .lblName
                                        If St1 Like "MAPSAY*" Then
                                            .txtCode = "FUNCTION Main(Player AS LONG, Message AS STRING) AS LONG" + Chr$(13) + Chr$(10) + "   Main = Continue" + Chr$(13) + Chr$(10) + "END FUNCTION"
                                        ElseIf St1 Like "MAP*" Or St1 Like "MONSTERDIE*" Or St1 Like "JOINMAP*" Or St1 Like "PARTMAP*" Or St1 = "JOINGAME" Or St1 = "PARTGAME" Then
                                            .txtCode = "SUB Main(Player AS LONG)" + Chr$(13) + Chr$(10) + "END SUB"
                                        ElseIf St1 Like "USEOBJ*" Or St1 Like "GETOBJ*" Or St1 Like "DROPOBJ*" Or St1 = "PLAYERDIE" Or St1 Like "MONSTERSEE*" Then
                                            .txtCode = "FUNCTION Main(Player AS LONG) AS LONG" + Chr$(13) + Chr$(10) + "   Main = Continue" + Chr$(13) + Chr$(10) + "END FUNCTION"
                                        ElseIf St1 = "PLAYERKILL" Then
                                            .txtCode = "FUNCTION Main(Killer AS LONG, Killee AS LONG) AS LONG" + Chr$(13) + Chr$(10) + "   Main = Continue" + Chr$(13) + Chr$(10) + "END FUNCTION"
                                        ElseIf St1 = "BROADCAST" Then
                                            .txtCode = "FUNCTION Main(Player AS LONG, Message AS STRING) AS LONG" + Chr$(13) + Chr$(10) + "   Main = Continue" + Chr$(13) + Chr$(10) + "END FUNCTION"
                                        ElseIf St1 = "COMMAND" Then
                                            .txtCode = "FUNCTION Main(Player as LONG, Command as STRING, Parm1 as STRING, Parm2 as STRING, Parm3 as STRING) AS LONG" + Chr$(13) + Chr$(10) + "Main = Continue" + Chr$(13) + Chr$(10) + "END FUNCTION"
                                        End If
                                    End If
                                    .Show 1
                                End With
                            End If
                        End If
                    
                    Case 95 'User is Away
                    If Len(St) >= 1 Then
                        A = Asc(Mid$(St, 1, 1))
                            If A >= 1 Then
                                PrintChat Player(A).Name + " is currently away! ' " + Mid$(St, 2) + "'", 14
                            End If
                    End If
                    
                    Case 96 'Custom Wav Play
                    If Len(St) >= 1 Then
                        A = Asc(Mid$(St, 1, 1))
                            If Exists("Sound" + CStr(A) + ".wav") Then
                                PlayWav A
                            End If
                    End If
                    
                    Case 97 'New Guild Info
                    Select Case Asc(Mid$(St, 1, 1))
                        Case 0 'Already in a guild.
                            PrintChat "You are already in a guild.  If you would like to create a new guild, you must first leave this guild by typing '/guild leave'.", 14
                    
                        Case 1 'New Guild Show
                            frmNewGuild.Show
                    
                        Case 2 'Guilds are Disabled
                            PrintChat "Guilds have been disabled.", 14
                    
                        Case 3 'Need to be atleast level 5
                            PrintChat "You must be atleast level 5 to join a guild!", 14
                    End Select
                    
                    Case 98 'Repairing
                        Select Case Asc(Mid$(St, 1, 1))
                        
                            Case 1 'NPC Repair Display
                                A = Asc(Mid$(St, 4, 1)) * 16777216 + Asc(Mid$(St, 5, 1)) * 65536 + Asc(Mid$(St, 6, 1)) * 256& + Asc(Mid$(St, 7, 1)) 'Repair Ammount
                                B = Asc(Mid$(St, 2, 1)) 'Dur
                                C = QBColor(1)
                                D = Asc(Mid$(St, 3, 1))
                                Select Case B
                                    Case Is >= 99
                                        St = "Excellent"
                                        C = QBColor(11)
                                    Case Is >= 75
                                        St = "Fair"
                                        C = QBColor(14)
                                    Case Is >= 50
                                        St = "Average"
                                        C = QBColor(2)
                                    Case Is >= 25
                                        St = "Seen Better Days"
                                        C = QBColor(4)
                                    Case Is >= 0
                                        St = "About to Break"
                                        C = QBColor(5)
                                End Select
                                frmMain.lblObjDur = Str(B) + "%"
                                frmMain.lblObjCond = St
                                frmMain.lblObjCond.ForeColor = C
                                frmMain.lblObjName = Trim$(Object(D).Name)
                                frmMain.lblRepairNPCName = Trim$(NPC(Map.NPC).Name)
                                frmMain.RepairObjPic.Cls
                                TransparentBlt frmMain.RepairObjPic.hdc, 0, 0, 32, 32, hdcObjects, 0, (Object(D).Picture - 1) * 32, hdcObjectsMask
                            If A = 0 Then
                                frmMain.lblRepairCst = "Free"
                                St1 = "Free"
                            Else
                                frmMain.lblRepairCst = Str(A) + " gold pieces."
                                St1 = Str(A)
                            End If
                            frmMain.lblRepairNpcTalk = "Hail adventurer! I am able to assist you with this. I can repair your " + Trim$(Object(D).Name) + " for " + St1 + " gold pieces. My efforts are worth for what I charge."
                            frmMain.picRepair.Visible = True
                        
                        Case 2 'Done Repairing Object
                            A = Asc(Mid$(St, 2, 1))
                            PrintChat "Your " + Object(A).Name + " is now at 100% durability. You repaired it successfully.", 14
                        End Select
                    
                    Case 99 'Projectiles
                        Select Case Asc(Mid$(St, 1, 1))
                            
                            Case 1 'Tile Effect
                                CreateTileEffect Asc(Mid$(St, 2, 1)), Asc(Mid$(St, 3, 1)), Asc(Mid$(St, 4, 1)), Asc(Mid$(St, 5, 1)) * 256 + Asc(Mid$(St, 6, 1)), Asc(Mid$(St, 7, 1)), Asc(Mid$(St, 8, 1)), Asc(Mid$(St, 9, 1))
                            Case 2 'Character Effect
                                CreateCharacterEffect Asc(Mid$(St, 2, 1)), Asc(Mid$(St, 3, 1)), Asc(Mid$(St, 4, 1)) * 256 + Asc(Mid$(St, 5, 1)), Asc(Mid$(St, 6, 1)), Asc(Mid$(St, 7, 1)), Asc(Mid$(St, 8, 1))
                            Case 3 'Monster Effect
                                If Asc(Mid$(St, 2, 1)) = Character.Index Then
                                    A = CX
                                    B = CY
                                Else
                                    A = Player(Asc(Mid$(St, 2, 1))).X
                                    B = Player(Asc(Mid$(St, 2, 1))).Y
                                End If
                                CreateMonsterEffect Asc(Mid$(St, 3, 1)), Asc(Mid$(St, 4, 1)), Asc(Mid$(St, 5, 1)) * 256 + Asc(Mid$(St, 6, 1)), Asc(Mid$(St, 6, 1)), A, B, Asc(Mid$(St, 8, 1))
                            Case 4 'Player Effect
                                If Asc(Mid$(St, 2, 1)) = Character.Index Then
                                    A = CX
                                    B = CY
                                Else
                                    A = Player(Asc(Mid$(St, 2, 1))).X
                                    B = Player(Asc(Mid$(St, 2, 1))).Y
                                End If
                                CreatePlayerEffect Asc(Mid$(St, 3, 1)), Asc(Mid$(St, 4, 1)), Asc(Mid$(St, 5, 1)) * 256 + Asc(Mid$(St, 6, 1)), Asc(Mid$(St, 7, 1)), A, B, Asc(Mid$(St, 8, 1))
                        End Select
                End Select
            End If
            GoTo LoopRead
        End If
    End If
End Sub
