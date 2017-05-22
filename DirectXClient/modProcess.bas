Attribute VB_Name = "modProcess"
Option Explicit

Sub ProcessRawData(SockData As String)
    Dim St As String, PacketLength As Long, PacketID As Long
LoopRead:
    If Len(SockData) >= 3 Then
        PacketLength = GetInt(Mid$(SockData, 1, 2))
        If Len(SockData) - 2 >= PacketLength Then
            St = Mid$(SockData, 3, PacketLength)
            SockData = Mid$(SockData, PacketLength + 3)

            If PacketLength > 0 Then
                PacketID = Asc(Mid$(St, 1, 1))
                If Len(St) > 1 Then
                    St = Mid$(St, 2)
                Else
                    St = vbNullString
                End If
                ProcessString PacketID, St
            End If
            GoTo LoopRead
        End If
    End If
End Sub
Sub ReceiveData()
    On Error GoTo HandleError

    Dim PacketLength As Integer, PacketID As Long
    Dim PacketChecksum As Long, CurrentChecksum As Long
    Dim ServPacketOrder As Integer
    Dim St As String

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

                CurrentChecksum = CheckSum(St) * 20 Mod 194
                'If Not CurrentChecksum = PacketChecksum Then
                '    SocketData = vbNullString
                '    St = vbNullString
                'Else
                If Len(St) > 1 Then
                    St = Mid$(St, 2)
                Else
                    St = vbNullString
                End If

                Select Case PacketID
                Case 170    'Raw Data
                    ProcessRawData St
                Case Else
                    ProcessString PacketID, St
                End Select
                'End If
            End If
            GoTo LoopRead
        End If
    End If

    Exit Sub

HandleError:
    If PacketID > 0 Then
        SendSocket Chr$(100) + "Error " & CStr(PacketID) & "  -  " & Err.Description
        MsgBox "Error " & CStr(PacketID) & "  -  " & Err.Description
    End If
End Sub

Sub ProcessString(PacketID As Long, St As String)
    On Error GoTo HandleError

    Dim A As Long, B As Long, C As Long, D As Long, E As Long
    Dim St1 As String

    ReceiveArray(PacketID) = ReceiveArray(PacketID) + 1

    Select Case PacketID
    Case 0    'Error Logging On
        If Len(St) >= 1 Then
            Select Case Asc(Mid$(St, 1, 1))
            Case 0    'Custom Message
                If Len(St) >= 2 Then
                    MsgBox Mid$(St, 2), vbOKOnly + vbExclamation, TitleString
                End If
            Case 1    'Invalid User/Pass
                MsgBox "Invalid user name/password!", vbOKOnly + vbExclamation, TitleString
            Case 2    'Account already in use
                MsgBox "Someone is already using that account!", vbOKOnly + vbExclamation, TitleString
            Case 3    'Banned
                If Len(St) >= 5 Then
                    A = Asc(Mid$(St, 2, 1)) * 16777216 + Asc(Mid$(St, 3, 1)) * 65536 + Asc(Mid$(St, 4, 1)) * 256& + Asc(Mid$(St, 5, 1))
                    If Len(St) > 5 Then
                        MsgBox "You are banned from The Odyssey Online Classic until " + CStr(CDate(A)) + " (" + Mid$(St, 6) + ")!", vbOKOnly, TitleString
                    Else
                        MsgBox "You are banned from The Odyssey Online Classic until " + CStr(CDate(A)) + "!", vbOKOnly, TitleString
                    End If
                    CloseClientSocket 3
                End If
            Case 4    'Server Full
                MsgBox "The server is full, please try again in a few minutes!", vbOKOnly + vbExclamation, TitleString
            Case 5    'Multiple Login
                MsgBox "You may not log in multiple times from the same computer!", vbOKOnly + vbExclamation, TitleString
            End Select
        End If
        CloseClientSocket 4

    Case 1    'Error Creating New Account
        If Len(St) >= 1 Then
            Select Case Asc(Mid$(St, 1, 1))
            Case 0    'Custom Message
                If Len(St) >= 2 Then
                    MsgBox Mid$(St, 2), vbOKOnly + vbExclamation, TitleString
                End If
            Case 1    'User name already in use
                MsgBox "That user name is already in use.  Please try another.", vbOKOnly + vbExclamation, TitleString
            End Select
        End If
        CloseClientSocket 2

    Case 2    'Account Created
        CloseClientSocket 0
        MsgBox "Your account has been created successfully!  Please write down your user name and password somewhere safe so that you do not loose them.  Click Login to connect to the game server.", vbOKOnly + vbExclamation, TitleString

    Case 3    'Logged On / Character Data
        If frmWait_Loaded = True Then Unload frmWait
        If frmLogin_Loaded = True Then Unload frmLogin
        If frmNewCharacter_Loaded = True Then Unload frmNewCharacter
        CWalkStep = 0
        If Len(St) >= 10 Then
            With Character
                .name = vbNullString
                .Class = Asc(Mid$(St, 1, 1))
                .Gender = Asc(Mid$(St, 2, 1))
                .Sprite = Asc(Mid$(St, 3, 1)) * 256 + Asc(Mid$(St, 4, 1))
                .Level = Asc(Mid$(St, 5, 1))
                .status = Asc(Mid$(St, 6, 1))
                .Guild = Asc(Mid$(St, 7, 1))
                .GuildRank = Asc(Mid$(St, 8, 1))
                .Access = Asc(Mid$(St, 9, 1))
                .index = Asc(Mid$(St, 10, 1))
                .Experience = Asc(Mid$(St, 11, 1)) * 16777216 + Asc(Mid$(St, 12, 1)) * 65536 + Asc(Mid$(St, 13, 1)) * 256& + Asc(Mid$(St, 14, 1))

                St = Mid$(St, 15)
                A = InStr(St, vbNullChar)
                If A > 1 Then
                    .name = Mid$(St, 1, A - 1)
                    If A < Len(St) Then
                        .Desc = Mid$(St, A + 1)
                    End If
                End If
            End With
            SetMap 0
            For A = 1 To MaxUsers
                Guild(A).name = vbNullString
                With Player(A)
                    .Sprite = 0
                    .Map = 0
                End With
            Next A
            With Character
                For A = 1 To MaxInvObjects
                    With .Inv(A)
                        .Object = 0
                        .EquippedNum = 0
                        .value = 0
                        .ItemPrefix = 0
                        .ItemSuffix = 0
                    End With
                Next A
            End With
            frmWait.Show
            frmWait.lblStatus = "Receiving Game Data ..."
            frmWait.btnCancel.Visible = True
            frmWait.Refresh
            SendSocket Chr$(7) + Chr$(1)    'I wanna play
        Else
            Character.Class = 0
            frmNewCharacter.Show
            frmWait.Hide
        End If

    Case 4    'Motd
        MOTDText = St

    Case 5    'Password Changed
        If frmWait_Loaded = True Then Unload frmWait

    Case 6    'Player Joined Game
        If Len(St) >= 7 Then
            A = Asc(Mid$(St, 1, 1))
            With Player(A)
                .Ignore = False
                .IsDead = False
                .Sprite = Asc(Mid$(St, 2, 1)) * 256 + Asc(Mid$(St, 3, 1))
                .status = Asc(Mid$(St, 4, 1))
                .Guild = Asc(Mid$(St, 5, 1))
                .MaxHP = Asc(Mid$(St, 6, 1))
                .name = Mid$(St, 7)
                If CMap > 0 Then
                    If Not .status = 25 Then
                        If .status = 2 Then
                            PrintChat "All hail " + .name + ", a new adventurer in this land!", 3
                        Else
                            PrintChat .name + " has joined the game!", 3
                        End If
                    End If
                End If
                UpdatePlayerColor A
            End With
        End If

    Case 7    'Player Left Game
        If Len(St) = 1 Then
            A = Asc(Mid$(St, 1, 1))
            If A >= 1 Then
                With Player(A)
                    If Not .status = 25 Then
                        PrintChat .name + " has left the game!", 3
                    End If
                    PlayerLeftMap A
                    .Sprite = 0
                    .IsDead = False
                End With
            End If
        End If

    Case 8    'Player joined map
        If Len(St) = 7 Then
            A = Asc(Mid$(St, 1, 1))
            With Player(A)
                .Map = CMap
                .X = Asc(Mid$(St, 2, 1))
                .Y = Asc(Mid$(St, 3, 1))
                .D = Asc(Mid$(St, 4, 1))
                .Sprite = Asc(Mid$(St, 5, 1)) * 256 + Asc(Mid$(St, 6, 1))
                .status = Asc(Mid$(St, 7, 1))
                .XO = .X * 32
                .YO = .Y * 32
                AddPlayerLight A
                .A = 0
            End With
        End If

    Case 9    'Player left map
        If Len(St) = 1 Then
            A = Asc(Mid$(St, 1, 1))
            If A >= 1 Then
                PlayerLeftMap A
                RemovePlayerLight A
            End If
        End If

    Case 10    'Player moved
        If Len(St) = 5 Then
            A = Asc(Mid$(St, 1, 1))
            If A > 0 And A < MaxUsers Then
                With Player(A)
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
                    .IsDead = False
                End With
            End If
        End If

    Case 11    'Say
        If Len(St) >= 2 Then
            A = Asc(Mid$(St, 1, 1))
            If Player(A).Ignore = False Then PrintChat Player(A).name + " says, " + Chr$(34) + Mid$(St, 2) + Chr$(34), 7
        End If

    Case 12    'You joined map
        If Len(St) = 13 Then
            If MapEdit = True Then CloseMapEdit
            'Destroy Projectiles
            For A = 1 To MaxProjectiles
                DestroyEffect A
            Next A
            For A = 1 To MaxFloatText
                ClearFloatText A
            Next A
            For A = 1 To MaxUsers
                Player(A).HP = 0
            Next A
            If CMap = 0 Then
                Tick = timeGetTime
                Character.LastMove = Tick
                St1 = vbNullString
                B = 0
                For A = 1 To MaxUsers
                    With Player(A)
                        If .Sprite > 0 And A <> Character.index And Not .status = 25 Then
                            B = B + 1
                            St1 = St1 + ", " + .name
                        End If
                    End With
                Next A
                Character.IsDead = False
                SetHP GetMaxHP
                SetEnergy GetMaxEnergy
                SetMana GetMaxMana
                If B > 0 Then
                    St1 = Mid$(St1, 2)
                    PrintChat "Welcome to the Odyssey Online Classic!  There are " + CStr(B) + " other players online:" + St1, 15
                Else
                    PrintChat "Welcome to the Odyssey Online Classic!  There are no other users currently online.", 15
                End If
                LoadOptions
                PrintChat MOTDText, 11
                Load frmMain
            End If
            SetMap Asc(Mid$(St, 1, 1)) * 256 + Asc(Mid$(St, 2, 1))
            CX = Asc(Mid$(St, 3, 1))
            CY = Asc(Mid$(St, 4, 1))
            CDir = Asc(Mid$(St, 5, 1))
            CXO = CX * 32
            CYO = CY * 32
            For A = 0 To MaxMonsters
                Map.Monster(A).Monster = 0
            Next A
            For A = 0 To MaxMapObjects
                Map.Object(A).Object = 0
                Map.Object(A).ItemPrefix = 0
                Map.Object(A).ItemSuffix = 0
            Next A
            For A = 0 To 9
                Map.Door(A).Att = 0
            Next A
            For A = 1 To MaxUsers
                Player(A).Map = 0
            Next A
            ClearLighting
            Freeze = True
            LoadMapFromCache CMap
            If RequestedMap = False Then
                If Map.Version <> Asc(Mid$(St, 6, 1)) * 16777216 + Asc(Mid$(St, 7, 1)) * 65536 + Asc(Mid$(St, 8, 1)) * 256& + Asc(Mid$(St, 9, 1)) Or CheckSum(MapData) <> Asc(Mid$(St, 10, 1)) * 16777216 + Asc(Mid$(St, 11, 1)) * 65536 + Asc(Mid$(St, 12, 1)) * 256& + Asc(Mid$(St, 13, 1)) Then
                    RequestedMap = True
                    SendSocket Chr$(45)
                Else
                    ShowMap
                End If
            End If
        End If

    Case 13    'Error creating character
        If frmWait_Loaded = True Then Unload frmWait
        frmNewCharacter.Show
        MsgBox "That name is already in use, please try another!", vbOKOnly + vbExclamation, TitleString

    Case 14    'New Map Object
        If Len(St) = 11 Then
            A = Asc(Mid$(St, 1, 1))
            If A <= MaxMapObjects Then
                With Map.Object(A)
                    .Object = GetInt(Mid$(St, 2, 2))
                    .X = Asc(Mid$(St, 4, 1))
                    .Y = Asc(Mid$(St, 5, 1))
                    .ItemPrefix = Asc(Mid$(St, 6, 1))
                    .ItemSuffix = Asc(Mid$(St, 7, 1))
                    B = Int(Rnd * 9)
                    .XOffset = B
                    B = Int(Rnd * 9)
                    .YOffset = B
                    .value = Asc(Mid$(St, 8, 1)) * 16777216 + Asc(Mid$(St, 9, 1)) * 65536 + Asc(Mid$(St, 10, 1)) * 256& + Asc(Mid$(St, 11, 1))
                    .PickedUp = 0
                    RedrawMapTile CLng(.X), CLng(.Y)
                End With
            End If
        End If

    Case 15    'Erase Map Object
        If Len(St) = 1 Then
            A = Asc(Mid$(St, 1, 1))
            If A <= MaxMapObjects Then
                With Map.Object(A)
                    .Object = 0
                    .ItemPrefix = 0
                    .ItemSuffix = 0
                    .value = 0
                    .PickedUp = 0
                    RedrawMapTile CLng(.X), CLng(.Y)
                End With
            End If
        End If

    Case 16    'Error messages
        If Len(St) >= 1 Then
            Select Case Asc(Mid$(St, 1, 1))
            Case 0    'Custom
                If Len(St) > 2 Then
                    PrintChat Mid$(St, 2), 7
                End If
            Case 1    'Inv full
                PrintChat "Your inventory is full!", 7
            Case 2    'Map full
                PrintChat "There is too much already on the ground here to drop that.", 7
            Case 3    'No such object
                PrintChat "No such object.", 7
            Case 4    'No such player
                PrintChat "No such player.", 7
            Case 5    'No such monster
                PrintChat "No such monster.", 7
            Case 6    'Player is too far away
                PrintChat "Player is too far away.", 7
            Case 7    'Monster is too far away
                PrintChat "Monster is too far away.", 7
            Case 8    'You cannot use that
                PrintChat "You cannot use that object.", 7
            Case 9    'Friendly Zone - can't attack
                PrintChat "This is a friendly area, you cannot attack here!", 7
            Case 10    'Cannot attack immortal
                PrintChat "You may not attack an immortal!", 7
            Case 11    'You are an immortal
                PrintChat "Immortals may not attack other players!", 7
            Case 12    'Can't attack monsters here
                PrintChat "You cannot attack these monsters!", 7
            Case 13    'Ban list full
                PrintChat "The ban list is full!", 7
            Case 14    'Not invited to join
                PrintChat "You have not been invited to join any guild.", 7
            Case 15    'Not enough cash
                PrintChat "You do not have enough gold to do that!", 7
            Case 16    'Guild name in use
                PrintChat "That name is already used either by another player or guild.  Please try another.", 7
            Case 17    'Guild full
                PrintChat "That guild is full!", 7
            Case 18    'too many guilds
                PrintChat "Too many guilds already exist.  You may join another guild or try again later.", 7
            Case 19    'cannot attack player -- he is not in guild
                PrintChat "That player is not in a guild -- you may not attack non-guild players.", 7
            Case 20    'cannot attack player -- you are not in guild
                PrintChat "You must be a member of a guild to attack other players.", 7
            Case 21    'not in a hall
                PrintChat "You are not in a guild hall!", 7
            Case 22    'hall already owned
                PrintChat "This hall is already owned by another guild.", 7
            Case 23    'already have hall
                PrintChat "Your guild already owns a hall.  You must move out of your old hall before you may purchase a new one.", 7
            Case 24    'don't have enough money to buy hall
                PrintChat "Your guild does not have enough money in its bank account to buy this hall.  Type /guild hallinfo for the price information of this hall.", 7
            Case 25    'do not own a guild hall
                PrintChat "Your guild does not own a hall.", 7
            Case 26    'need 5 members
                PrintChat "You must have atleast 5 members in your guild before you may do that.", 7
            Case 27    'Can't afford that
                PrintChat "You do not have the items required to purchase that!", 7
            Case 28    'Not in a bank
                PrintChat "You are not in a bank!", 7
            Case 29    'too far away
                PrintChat "That player is too far away to hit!", 7
            Case 30    'must be Level # to join guild
                PrintChat "You must be at least Level " + CStr(World.GuildJoinLevel) + " to join a guild!", 7
            Case 32    'Must be in a smithy shop
                PrintChat "You are not in a blacksmithy shop!", 7
            Case 33    'Do not have enough money
                PrintChat "You do not have enough money to repair this object!", 7
            Case 34    'Do not have specified object
                PrintChat "You do not have the object to be repaired!", 7
            Case 35    'Not high enough stats
                PrintChat "You do not meet the requirements to use this item!", 7
            Case 36    'Ammo
                PrintChat "This ammo must be used with a projectile weapon!", 7
            Case 37    'Ammo
                PrintChat "This ammo is not used with this weapon!", 7
            Case 40    'Can't use
                PrintChat "Your class cannot use this object!", 7
            Case 41    'Squelched or Jailed
                PrintChat "Nobody hears your pitiful cries.", 7
            Case 42    'New Day
                PrintChat "A new day dawns in the land of Odyssey!", 7
            Case 43    'No Guild MOTD
                PrintChat "Your guild does not currently have a message of the day.  If you are a lord or a founder, you can set this by typing /guild motd <message>.", 7
            Case 44    'Bank is Full
                PrintChat "Your bank is full!", 7
            Case 45    'E-mail is already saved
                PrintChat "Your account already has an e-mail address saved!", 7
            Case 46    'E-mail updated
                PrintChat "Your e-mail address has been stored!", 7
            Case 47    'E-mail needed
                PrintChat "You have not entered an email address for your account.  Type /email to enter it.  This is the only way to recover your account if it is lost or stolen!", 7
            Case 48    'Server is being backed up
                PrintChat "Server is being backed up.", 7
            End Select
        End If

    Case 17    'New Inv Object
        If Len(St) = 9 Then
            A = Asc(Mid$(St, 1, 1))
            If A >= 1 And A <= 20 Then
                With Character.Inv(A)
                    .Object = GetInt(Mid$(St, 2, 2))
                    .value = Asc(Mid$(St, 4, 1)) * 16777216 + Asc(Mid$(St, 5, 1)) * 65536 + Asc(Mid$(St, 6, 1)) * 256& + Asc(Mid$(St, 7, 1))
                    .ItemPrefix = Asc(Mid$(St, 8, 1))
                    .ItemSuffix = Asc(Mid$(St, 9, 1))
                End With
                If frmMain.picRepair.Visible = True Then
                    DisplayRepair
                ElseIf frmMain.picSellObject.Visible = True Then
                    DisplaySell
                End If
                RefreshInventory
            End If
        End If

    Case 18    'Erase Inv Object
        If Len(St) = 1 Then
            A = Asc(Mid$(St, 1, 1))
            If A >= 1 And A <= 20 Then
                With Character.Inv(A)
                    If .EquippedNum > 0 And .Object > 0 Then
                        If Object(.Object).Type = 10 Then
                            Character.Projectile = False
                        ElseIf Object(.Object).Type = 11 Then
                            Character.Ammo = 0
                        End If
                    End If
                    .Object = 0
                    .ItemPrefix = 0
                    .ItemSuffix = 0
                    .value = 0
                    .EquippedNum = 0
                End With
                If frmMain.picRepair.Visible = True Then
                    DisplayRepair
                ElseIf frmMain.picSellObject.Visible = True Then
                    DisplaySell
                End If
                RefreshInventory
            End If
        End If

    Case 19    'Use Object
        If Len(St) >= 1 Then
            A = Asc(Mid$(St, 1, 1))
            If A >= 1 And A <= 20 Then
                If Character.Inv(A).Object > 0 Then
                    Select Case Object(Character.Inv(A).Object).Type
                    Case 2, 3, 4    'Armor pieces
                        Character.EquippedObject(Object(Character.Inv(A).Object).Type).Object = Character.Inv(A).Object
                        Character.EquippedObject(Object(Character.Inv(A).Object).Type).value = Character.Inv(A).value
                        Character.EquippedObject(Object(Character.Inv(A).Object).Type).ItemPrefix = Character.Inv(A).ItemPrefix
                        Character.EquippedObject(Object(Character.Inv(A).Object).Type).ItemSuffix = Character.Inv(A).ItemSuffix
                        Character.Inv(A).Object = 0
                        Character.Inv(A).value = 0
                        Character.Inv(A).ItemPrefix = 0
                        Character.Inv(A).ItemSuffix = 0
                    Case 8    'Ring
                        Character.EquippedObject(5).Object = Character.Inv(A).Object
                        Character.EquippedObject(5).value = Character.Inv(A).value
                        Character.EquippedObject(5).ItemPrefix = Character.Inv(A).ItemPrefix
                        Character.EquippedObject(5).ItemSuffix = Character.Inv(A).ItemSuffix
                        Character.Inv(A).Object = 0
                        Character.Inv(A).value = 0
                        Character.Inv(A).ItemPrefix = 0
                        Character.Inv(A).ItemSuffix = 0
                    Case 1    'Weapons
                        Character.EquippedObject(1).Object = Character.Inv(A).Object
                        Character.EquippedObject(1).value = Character.Inv(A).value
                        Character.EquippedObject(1).ItemPrefix = Character.Inv(A).ItemPrefix
                        Character.EquippedObject(1).ItemSuffix = Character.Inv(A).ItemSuffix
                        Character.Inv(A).Object = 0
                        Character.Inv(A).value = 0
                        Character.Inv(A).ItemPrefix = 0
                        Character.Inv(A).ItemSuffix = 0
                        Character.Projectile = False
                    Case 10    'Projectile Weapon
                        Character.EquippedObject(1).Object = Character.Inv(A).Object
                        Character.EquippedObject(1).value = Character.Inv(A).value
                        Character.EquippedObject(1).ItemPrefix = Character.Inv(A).ItemPrefix
                        Character.EquippedObject(1).ItemSuffix = Character.Inv(A).ItemSuffix
                        Character.Inv(A).Object = 0
                        Character.Inv(A).value = 0
                        Character.Inv(A).ItemPrefix = 0
                        Character.Inv(A).ItemSuffix = 0
                        Character.Projectile = True
                    Case 11    'Ammo (Stays in inventory)
                        Character.Ammo = A
                        Character.Inv(A).EquippedNum = A
                    End Select
                    RefreshInventory
                End If
            End If
        End If

    Case 20    'Stop using object
        If Len(St) = 1 Then
            A = Asc(Mid$(St, 1, 1))
            If A >= 1 And A <= 20 Then
                If Character.Inv(A).Object > 0 Then
                    Character.Inv(A).EquippedNum = False
                    If Object(Character.Inv(A).Object).Type = 10 Then Character.Projectile = False
                    If Object(Character.Inv(A).Object).Type = 11 Then Character.Ammo = 0
                    RefreshInventory
                End If
            Else
                If A >= 21 Then    'Ok
                    A = A - 20
                    If Character.EquippedObject(A).Object > 0 Then
                        If Object(Character.EquippedObject(A).Object).Type = 10 Then Character.Projectile = False
                        Character.EquippedObject(A).Object = 0
                        Character.EquippedObject(A).value = 0
                        Character.EquippedObject(A).ItemPrefix = 0
                        Character.EquippedObject(A).ItemSuffix = 0
                        RefreshInventory
                    End If
                End If
            End If
        End If

    Case 21    'Map Data
        ProcessReceivedMap St

    Case 24    'Joined Game
        If frmWait_Loaded = True Then
            frmWait.lblStatus = "Loading Game ..."
            frmWait.lblStatus.Refresh
        End If
        For A = 1 To MaxUsers
            Player(A).Map = 0
        Next A
        With Character
            For A = 1 To MaxInvObjects
                With .Inv(A)
                    .Object = 0
                    .EquippedNum = 0
                    .value = 0
                    .ItemPrefix = 0
                    .ItemSuffix = 0
                End With
            Next A
            For A = 1 To 5
                .EquippedObject(A).Object = 0
                .EquippedObject(A).value = 0
                .EquippedObject(A).ItemPrefix = 0
                .EquippedObject(A).ItemSuffix = 0
            Next A
        End With
        Load frmMain
        frmMain.WindowState = 0
        SetHP GetMaxHP
        SetMana GetMaxMana
        SetEnergy GetMaxEnergy
        DrawStats
        RedrawMap = True
        If Character.Level = 1 Then frmMain.picHelp.Visible = True
        blnPlaying = True
        ResetTimers

    Case 25    'Tell
        If Len(St) >= 2 Then
            A = Asc(Mid$(St, 1, 1))
            If A >= 1 Then
                With Player(A)
                    If .Ignore = False Then
                        PrintChat .name + " tells you, " + Chr$(34) + Mid$(St, 2) + Chr$(34), 10
                    End If
                End With
            End If
        End If

    Case 26    'Broadcast
        If Len(St) >= 2 And options.Broadcasts = True Then
            A = Asc(Mid$(St, 1, 1))
            If A >= 1 Then
                If Player(A).Ignore = False Then
                    PrintChat Player(A).name + ": " + Mid$(St, 2), 13
                End If
            End If
        End If

    Case 27    'Emote
        If Len(St) >= 2 Then
            A = Asc(Mid$(St, 1, 1))
            If Player(A).Ignore = False Then
                PrintChat Player(A).name + " " + Mid$(St, 2), 11
            End If
        End If

    Case 28    'Yell
        If Len(St) >= 2 Then
            A = Asc(Mid$(St, 1, 1))
            If Player(A).Ignore = False Then
                PrintChat Player(A).name + " yells, " + Chr$(34) + Mid$(St, 2) + Chr$(34), 7
            End If
        End If

    Case 30    'Server Message
        If Len(St) > 0 Then
            PrintChat "Server Message: " + St, 9
        End If

    Case 31    'Object Data
        If Len(St) >= 10 Then
            A = GetInt(Mid$(St, 1, 2))
            If A >= 1 Then
                With Object(A)
                    .Picture = Asc(Mid$(St, 3, 1)) * 256 + Asc(Mid$(St, 4, 1))
                    .Type = Asc(Mid$(St, 5, 1))
                    Select Case .Type
                    Case 8    'Ring
                        .MaxDur = Asc(Mid$(St, 7, 1))
                        .Data2 = Asc(Mid$(St, 6, 1))
                        .Modifier = Asc(Mid$(St, 8, 1))
                    Case 10, 11    'Projectile Weapon
                        .Modifier = Asc(Mid$(St, 6, 1))
                        .Data2 = Asc(Mid$(St, 8, 1))
                    Case Else
                        .MaxDur = Asc(Mid$(St, 6, 1))
                        .Modifier = Asc(Mid$(St, 7, 1))
                        .Data2 = Asc(Mid$(St, 8, 1))
                    End Select
                    .flags = Asc(Mid$(St, 9, 1))
                    .ClassReq = Asc(Mid$(St, 10, 1))
                    .LevelReq = Asc(Mid$(St, 11, 1))
                    .Version = Asc(Mid$(St, 12, 1))
                    .SellPrice = Asc(Mid$(St, 13, 1)) * 256 + Asc(Mid$(St, 14, 1))
                    If Len(St) >= 15 Then
                        .name = Mid$(St, 15)
                    Else
                        .name = vbNullString
                    End If
                    If frmMonster_Loaded = True Then
                        frmMonster.cmbObject(0).List(A) = CStr(A) + ": " + .name
                        frmMonster.cmbObject(1).List(A) = CStr(A) + ": " + .name
                        frmMonster.cmbObject(2).List(A) = CStr(A) + ": " + .name
                    End If
                    If frmNPC_Loaded = True Then
                        frmNPC.cmbGiveObject.List(A) = CStr(A) + ": " + .name
                        frmNPC.cmbTakeObject.List(A) = CStr(A) + ": " + .name
                    End If
                    If frmList_Loaded = True Then
                        frmList.DrawList
                    End If
                End With
                RefreshInventory
                For B = 0 To MaxMapObjects
                    With Map.Object(B)
                        If .Object = A Then
                            RedrawMapTile CLng(.X), CLng(.Y)
                        End If
                    End With
                Next B
                Debug.Print "Save Object " + CStr(A)
                SaveObject CInt(A)
            End If
        End If

    Case 32    'Monster Data
        If Len(St) >= 8 Then
            A = GetInt(Mid$(St, 1, 2))
            If A >= 1 Then
                With Monster(A)
                    .Sprite = Asc(Mid$(St, 3, 1)) * 256 + Asc(Mid$(St, 4, 1))
                    .Version = Asc(Mid$(St, 5, 1))
                    .MaxLife = GetInt(Mid$(St, 6, 2))
                    .flags = Asc(Mid$(St, 8, 1))
                    If Len(St) >= 9 Then
                        .name = Mid$(St, 9)
                    Else
                        .name = vbNullString
                    End If
                    If frmList_Loaded = True Then
                        frmList.DrawList
                    End If
                    If frmMapProperties_Loaded = True Then
                        For B = 0 To 9
                            frmMapProperties.cmbMonster(B).List(A) = CStr(A) + ": " + .name
                        Next B

                    End If
                End With
                Debug.Print "Save Monster " + CStr(A)
                SaveMonster A
            End If
        End If

    Case 33    'Edit Object Data
        If Len(St) = 9 Then
            A = GetInt(Mid$(St, 1, 2))
            If ExamineBit(Object(A).flags, 5) = 255 Then
                If Character.Access < 3 Then
                    PrintChat "This object is locked!", 7
                    Exit Sub
                End If
            End If
            If frmObject_Loaded = False Then Load frmObject
            With frmObject
                .lblNumber = A
                .txtName = Object(A).name
                If Object(A).Picture > 0 Then
                    .sclPicture = Object(A).Picture
                Else
                    .sclPicture = 1
                End If
                B = Asc(Mid$(St, 3, 1))
                For C = 0 To 6
                    If ExamineBit(CByte(B), CByte(C)) = True Then
                        frmObject.chkFlags(C) = 1
                    Else
                        frmObject.chkFlags(C) = 0
                    End If
                Next C
                .ObjData(0) = Asc(Mid$(St, 4, 1))
                .ObjData(1) = Asc(Mid$(St, 5, 1))
                .ObjData(2) = Asc(Mid$(St, 6, 1))
                .ObjData(3) = Asc(Mid$(St, 7, 1))
                .lblData(0) = .ObjData(0)
                .lblData(1) = .ObjData(1)
                .lblData(2) = .ObjData(2)
                .lblData(3) = .ObjData(3)
                .sclLevel.value = Asc(Mid$(St, 9, 1))
                B = Asc(Mid$(St, 8, 1))
                For C = 0 To 3
                    If ExamineBit(CByte(B), CByte(C)) = True Then
                        frmObject.chkClass(C) = 1
                    Else
                        frmObject.chkClass(C) = 0
                    End If
                Next C
                If Object(A).Type < .cmbType.ListCount Then
                    .cmbType.ListIndex = 0
                    .cmbType.ListIndex = Object(A).Type
                Else
                    .cmbType.ListIndex = 0
                End If
                .sclSellPrice = Object(A).SellPrice
                .txtSellPrice = .sclSellPrice
                .Show 1
            End With
        End If

    Case 34    'Edit Monster Data
        If Len(St) > 16 Then
            A = GetInt(Mid$(St, 1, 2))
            If frmMonster_Loaded = False Then Load frmMonster
            With frmMonster
                .lblNumber = A
                .txtName = Monster(A).name
                If Monster(A).Sprite > 0 Then
                    .sclSprite = Monster(A).Sprite
                Else
                    .sclSprite = 1
                End If
                B = GetInt(Mid$(St, 3, 2))
                If B > 0 Then .sclHP = B Else .sclHP = 1
                B = Asc(Mid$(St, 5, 1))
                If B > 0 Then .sclStrength = B Else .sclStrength = 1
                B = Asc(Mid$(St, 6, 1))
                .sclArmor = B
                B = Asc(Mid$(St, 7, 1))
                If B > 0 Then .sclSpeed = B Else .sclSpeed = 1
                B = Asc(Mid$(St, 8, 1))
                If B > 0 Then .sclSight = B Else .sclSight = 1
                B = Asc(Mid$(St, 9, 1))
                If B <= 100 Then .sclAgility = B Else .sclAgility = 100
                B = Asc(Mid$(St, 10, 1))
                For C = 0 To 7
                    If ExamineBit(CByte(B), CByte(C)) = True Then
                        .chkFlag(C) = 1
                    Else
                        .chkFlag(C) = 0
                    End If
                Next C
                .cmbObject(0).ListIndex = GetInt(Mid$(St, 11, 2))
                .txtValue(0) = Asc(Mid$(St, 13, 1))
                .cmbObject(1).ListIndex = GetInt(Mid$(St, 14, 2))
                .txtValue(1) = Asc(Mid$(St, 16, 1))
                .cmbObject(2).ListIndex = GetInt(Mid$(St, 17, 2))
                .txtValue(2) = Asc(Mid$(St, 19, 1))
                .sclExperience.value = Asc(Mid$(St, 20, 1))
                .sclMagicDefense.value = Asc(Mid$(St, 21, 1))
                .lblExperience = .sclExperience.value
                .lblMagicDefense = .sclMagicDefense.value
                .Show 1
            End With
        End If
    Case 35    'Repeat
        If Len(St) >= 1 Then
            SendSocket St
        End If

    Case 36    'Door Open
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

    Case 37    'Close Door
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

    Case 38    'New Map Monster
        If Len(St) = 8 Then
            A = Asc(Mid$(St, 1, 1))
            If A <= MaxMonsters Then
                With Map.Monster(A)
                    .Monster = GetInt(Mid$(St, 2, 2))
                    .X = Asc(Mid$(St, 4, 1))
                    .Y = Asc(Mid$(St, 5, 1))
                    .D = Asc(Mid$(St, 6, 1))
                    .Life = GetInt(Mid$(St, 7, 2))
                    .XO = .X * 32
                    .YO = .Y * 32
                    .A = 0
                    .HPBar = False
                End With
            End If
        End If

    Case 39    'Monster Die
        If Len(St) = 1 Then
            A = Asc(Mid$(St, 1, 1))
            If A <= MaxMonsters Then
                PlayWav 9
                Map.Monster(A).Monster = 0
                MonsterDied A
            End If
        End If

    Case 40    'Monster Move
        If Len(St) = 4 Then
            A = Asc(Mid$(St, 1, 1))
            If A <= MaxMonsters Then
                With Map.Monster(A)
                    If CLng(.X) * 32 <> .XO Then
                        .X = Asc(Mid$(St, 2, 1))
                        .XO = .X * 32
                    Else
                        .X = Asc(Mid$(St, 2, 1))
                    End If
                    If CLng(.Y) * 32 <> .YO Then
                        .Y = Asc(Mid$(St, 3, 1))
                        .YO = .Y * 32
                    Else
                        .Y = Asc(Mid$(St, 3, 1))
                    End If
                    .D = Asc(Mid$(St, 4, 1))
                End With
            End If
        End If

    Case 41    'Monster Attack
        If Len(St) = 1 Then
            A = Asc(Mid$(St, 1, 1))
            If A <= MaxMonsters Then
                Map.Monster(A).A = 5
            End If
        End If

    Case 42    'Player Attack
        If Len(St) = 1 Then
            A = Asc(Mid$(St, 1, 1))
            If A >= 1 Then
                Player(A).A = 5
            End If
        End If

    Case 43    'You hit player
        If Len(St) = 3 Then
            A = Asc(Mid$(St, 2, 1))
            If A >= 1 Then
                B = Asc(Mid$(St, 3, 1))
                Select Case Asc(Mid$(St, 1, 1))
                Case 0
                    PlayWav 2
                    CAttack = 5
                Case 1
                    PlayWav 3
                End Select
                WeaponLoss
            End If
        End If

    Case 44    'Player hit monster
        If Len(St) = 6 Then
            A = Asc(Mid$(St, 1, 1))    'Player Index
            B = Asc(Mid$(St, 2, 1))    'Event Type (missed/hit)
            C = Asc(Mid$(St, 3, 1))    'Monster Index
            D = Asc(Mid$(St, 4, 1))    'Damage
            E = GetInt(Mid$(St, 5, 2))    'Monster's HP

            If C >= 0 And C <= MaxMonsters Then
                If Map.Monster(C).Monster > 0 Then
                    Map.Monster(C).Life = CInt(E)
                    If A = Character.index Then
                        Select Case B
                        Case 0    'Hit
                            PlayWav 2
                            CAttack = 6
                            CreateFloatText CStr(D), BRIGHTRED, Map.Monster(C).X, Map.Monster(C).Y
                        Case 1    'Miss
                            PlayWav 3
                            CreateFloatText "Miss!", YELLOW, Map.Monster(C).X, Map.Monster(C).Y
                        End Select
                        Map.Monster(C).HPBar = True
                        WeaponLoss
                    Else
                        Select Case B
                        Case 0    'Hit
                            Player(A).A = 6
                            CreateFloatText CStr(D), BRIGHTRED, Map.Monster(C).X, Map.Monster(C).Y
                        Case 1    'Miss
                            CreateFloatText "Miss!", YELLOW, Map.Monster(C).X, Map.Monster(C).Y
                        End Select
                        Player(A).A = 6
                    End If
                End If
            End If
        End If

    Case 45    'You killed player
        If Len(St) = 5 Then
            A = Asc(Mid$(St, 1, 1))
            If A >= 1 Then
                With Player(A)
                    If .status = 1 Then
                        PrintChat "You have put the evil murderer " + .name + " to justice!", 12
                    Else
                        PrintChat "You have murdered " + .name + " in cold blood!", 12
                    End If
                    .IsDead = True
                    B = Asc(Mid$(St, 2, 1)) * 16777216 + Asc(Mid$(St, 3, 1)) * 65536 + Asc(Mid$(St, 4, 1)) * 256& + Asc(Mid$(St, 5, 1))
                    Character.Experience = B
                    DrawStats
                End With
                PlayWav 8
                For C = 1 To MaxProjectiles
                    If Projectile(C).TargetNum = A Then
                        DestroyEffect C
                    End If
                Next C
            End If
        ElseIf Len(St) = 1 Then
            A = Asc(Mid$(St, 1, 1))
            Player(A).A = 5
        End If

    Case 46    'Change HP
        If Len(St) = 1 Then
            SetHP Asc(Mid$(St, 1, 1))
            DrawStats
        End If

    Case 47    'Change Energy
        If Len(St) = 1 Then
            SetEnergy Asc(Mid$(St, 1, 1))
            DrawStats
        End If

    Case 48    'Change Mana
        If Len(St) = 1 Then
            SetMana Asc(Mid$(St, 1, 1))
            DrawStats
        End If

    Case 49    'Player Hit You
        If Len(St) = 2 Then
            A = Asc(Mid$(St, 1, 1))
            If A >= 1 Then
                B = Asc(Mid$(St, 2, 1))
                Character.LastMove = Tick

                PlayWav 2
                If B > 0 Then
                    If GetHP > B Then
                        SetHP GetHP - B
                    Else
                        SetHP 0
                    End If
                    DrawStats
                End If
                Player(A).A = 5
                ArmorLoss
            End If
        End If

    Case 50    'Monster Hit You
        If Len(St) = 3 Then
            A = Asc(Mid$(St, 2, 1))
            If A <= MaxMonsters Then
                B = Map.Monster(A).Monster
                C = Asc(Mid$(St, 3, 1))
                Character.LastMove = Tick
                Select Case Asc(Mid$(St, 1, 1))
                Case 0
                    PlayWav 2
                    If C > 0 Then
                        If GetHP > C Then
                            SetHP GetHP - C
                        Else
                            SetHP 0
                        End If
                        DrawStats
                    End If
                    Map.Monster(A).A = 5
                    ArmorLoss
                Case 1
                    PlayWav 3
                    If B > 0 Then
                        'PrintInfoText "The " + Monster(B).Name + " misses you."
                    Else
                        'PrintInfoText "The monster misses you."
                    End If
                End Select
            End If
        End If

    Case 51    'You killed the monster
        If Len(St) = 5 Then
            A = Asc(Mid$(St, 1, 1))
            If A <= MaxMonsters Then
                PlayWav 9
                With Map.Monster(A)
                    .Monster = 0
                End With
                B = Asc(Mid$(St, 2, 1)) * 16777216 + Asc(Mid$(St, 3, 1)) * 65536 + Asc(Mid$(St, 4, 1)) * 256& + Asc(Mid$(St, 5, 1))
                Character.Experience = B
                DrawStats
                MonsterDied (A)
            End If
        End If

    Case 52    'Player Killed You
        If Len(St) = 1 Then
            A = Asc(Mid$(St, 1, 1))
            If A >= 1 Then
                With Player(A)
                    If Character.status = 1 Then
                        PrintChat .name + " has put you to justice!", 12
                    Else
                        PrintChat .name + " has murdered you in cold blood!", 12
                    End If
                End With
                YouDied
            End If
        End If

    Case 53    'Monster Killed You
        If Len(St) = 2 Then
            A = GetInt(Mid$(St, 1, 2))
            If A >= 1 Then
                NextTransition = 6
                PlayWav 8
                With Monster(A)
                    PrintChat "The " + .name + " has killed you!", 12
                End With
                YouDied
            End If
        End If

    Case 56    'Text
        If Len(St) >= 2 Then
            PrintChat Mid$(St, 2), Asc(Mid$(St, 1, 1))
        End If

    Case 57    'Object Breaks
        If Len(St) = 1 Then
            A = Asc(Mid$(St, 1, 1))
            If A >= 1 And A <= 5 Then
                With Character.EquippedObject(A)
                    If .Object > 0 Then
                        PrintInfoText "Your " + Object(.Object).name + " breaks!"
                        PrintChat "Your " + Object(.Object).name + " breaks!", 14
                    End If
                    .Object = 0
                    .value = 0
                    .ItemPrefix = 0
                    .ItemSuffix = 0
                End With
                RefreshInventory
            End If
        End If

    Case 58    'Ping
        SendSocket Chr$(29)    'Pong

    Case 59    'Level Up
        With Character
            .Level = .Level + 1
            SetMaxHP Asc(Mid$(St, 1, 1))
            SetMaxEnergy Asc(Mid$(St, 2, 1))
            SetMaxMana Asc(Mid$(St, 3, 1))
            .Experience = 0
            PrintChat "Level Up!  You are now Level " + CStr(.Level) + "!", 12
            DrawStats
        End With

    Case 60    'Experience Change
        If Len(St) = 4 Then
            Character.Experience = Asc(Mid$(St, 1, 1)) * 16777216 + Asc(Mid$(St, 2, 1)) * 65536 + Asc(Mid$(St, 3, 1)) * 256& + Asc(Mid$(St, 4, 1))
        End If

    Case 61    'Player killed by player
        If Len(St) = 2 Then
            A = Asc(Mid$(St, 1, 1))
            B = Asc(Mid$(St, 2, 1))
            If A >= 1 And B >= 1 Then
                With Player(A)
                    .IsDead = True
                    If Not A = B Then
                        If .status = 1 Then
                            If .Map = CMap Then
                                PlayWav 8
                                CreateFloatText "Dead!", 12, .X, .Y
                            End If
                            PrintChat Player(B).name + " has put " + .name + " to justice!", 12
                        Else
                            If .Map = CMap Then
                                PlayWav 8
                                CreateFloatText "Dead!", 12, .X, .Y
                            End If
                            PrintChat Player(B).name + " has murdered " + .name + " in cold blood!", 12
                        End If
                    Else
                        PrintChat Player(B).name + " has died!", 12
                    End If
                End With
            End If
        End If

    Case 62    'Player killed by monster
        If Len(St) = 3 Then
            A = Asc(Mid$(St, 1, 1))
            B = GetInt(Mid$(St, 2, 2))
            If A >= 1 And B >= 1 Then
                With Player(A)
                    .IsDead = True
                    If .Map = CMap Then
                        PlayWav 8
                        CreateFloatText "Dead!", 12, .X, .Y
                    End If
                    PrintChat .name + " has been killed by a " + Monster(B).name + "!", 12
                End With
            End If
        End If

    Case 63    'Player Sprite Changed
        If Len(St) = 3 Then
            A = Asc(Mid$(St, 1, 1))
            B = Asc(Mid$(St, 2, 1)) * 256 + Asc(Mid$(St, 3, 1))
            If A >= 1 And B >= 1 Then
                If A = Character.index Then
                    Character.Sprite = B
                Else
                    If Player(A).Sprite > 0 Then
                        Player(A).Sprite = B
                    End If
                End If
            End If
        End If

    Case 64    'Player Name Change
        If Len(St) >= 2 Then
            A = Asc(Mid$(St, 1, 1))
            If A >= 1 Then
                If A = Character.index Then
                    Character.name = Mid$(St, 2)
                Else
                    If Player(A).Sprite > 0 Then
                        Player(A).name = Mid$(St, 2)
                    End If
                End If
            End If
        End If

    Case 65    'Changed access
        If Len(St) = 1 Then
            Character.Access = Asc(Mid$(St, 1, 1))
            If Character.Access > 0 Then
                Character.status = 3
            Else
                CWalkStep = 4
                Character.status = 0
            End If
        End If

    Case 66    'Player banned
        If Len(St) >= 2 Then
            B = Asc(Mid$(St, 1, 1))
            C = Asc(Mid$(St, 2, 1))
            If B >= 1 Then
                If C >= 1 Then
                    If Len(St) > 2 Then
                        PrintChat Player(B).name + " has been banned by " + Player(C).name + ": " + Mid$(St, 3), 15
                    Else
                        PrintChat Player(B).name + " has been banned by " + Player(C).name + "!", 15
                    End If
                Else
                    If Len(St) > 2 Then
                        PrintChat Player(B).name + " has been banned: " + Mid$(St, 3), 15
                    Else
                        PrintChat Player(B).name + " has been banned!", 15
                    End If
                End If
            End If
        End If

    Case 67    'Booted
        If Len(St) >= 1 Then
            A = Asc(Mid$(St, 1, 1))
            If A >= 1 Then
                If Len(St) > 1 Then
                    MsgBox "You have been booted from The Odyssey by " + Player(A).name + ": " + Mid$(St, 2), vbOKOnly + vbExclamation, TitleString
                Else
                    MsgBox "You have been booted from The Odyssey by " + Player(A).name + "!", vbOKOnly + vbExclamation, TitleString
                End If
            Else
                If Len(St) > 1 Then
                    MsgBox "You have been booted from The Odyssey: " + Mid$(St, 2), vbOKOnly + vbExclamation, TitleString
                Else
                    MsgBox "You have been booted from The Odyssey!", vbOKOnly + vbExclamation, TitleString
                End If
            End If
            CloseClientSocket 0
        End If

    Case 68    'Player Booted
        If Len(St) >= 2 Then
            A = Asc(Mid$(St, 1, 1))
            B = Asc(Mid$(St, 2, 1))
            If A >= 1 Then
                If B >= 1 Then
                    If Len(St) > 2 Then
                        PrintChat Player(A).name + " has been booted by " + Player(B).name + ": " + Mid$(St, 3), 15
                    Else
                        PrintChat Player(A).name + " has been booted by " + Player(B).name + "!", 15
                    End If
                Else
                    If Len(St) > 2 Then
                        PrintChat Player(A).name + " has been booted: " + Mid$(St, 3), 15
                    Else
                        PrintChat Player(A).name + " has been booted!", 15
                    End If
                End If
            End If
        End If

    Case 69    'Ban List
        If Len(St) >= 2 Then
            A = Asc(Mid$(St, 1, 1))
            With frmList.lstList
                .AddItem CStr(A) + ": " + Mid$(St, 2)
                .ItemData(.ListCount - 1) = A
            End With
        Else
            frmList.Show
        End If

    Case 70    'Guild Data
        If Len(St) >= 1 Then
            A = Asc(Mid$(St, 1, 1))
            If A >= 1 Then
                If Len(St) > 1 Then
                    B = Asc(Mid$(St, 2, 1))
                    Guild(A).name = Mid$(St, 3)
                    Guild(A).MemberCount = B
                Else
                    Guild(A).name = vbNullString
                End If
            End If
        End If

    Case 71    'Guild Dec. Data
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

    Case 72    'Guild Change
        If Len(St) = 1 Then
            A = Asc(Mid$(St, 1, 1))
            If A > 0 Then
                PrintChat "You are now a member of " + Chr$(34) + Guild(A).name + Chr$(34) + ".  For available guild commands, type /guild help.", 15
            Else
                If Character.Guild > 0 Then
                    PrintChat "You are no longer a member of " + Chr$(34) + Guild(Character.Guild).name + Chr$(34), 15
                End If
            End If
            Character.Guild = A
            Character.GuildRank = 0
            UpdatePlayersColors
        End If

    Case 73    'Player Changed Guild
        If Len(St) = 2 Then
            A = Asc(Mid$(St, 1, 1))
            B = Asc(Mid$(St, 2, 1))
            If A >= 1 Then
                If Player(A).Guild = Character.Guild And Character.Guild > 0 Then
                    PrintChat Player(A).name + " is no longer a member of your guild.", 15
                End If
                Player(A).Guild = B
                If B > 0 And B = Character.Guild Then
                    PrintChat Player(A).name + " is now a member of your guild.", 15
                End If
            End If
            UpdatePlayerColor A
        End If

    Case 74    'Guild Account Status
        If Len(St) = 8 Then
            A = Asc(Mid$(St, 1, 1)) * 16777216 + Asc(Mid$(St, 2, 1)) * 65536 + Asc(Mid$(St, 3, 1)) * 256& + Asc(Mid$(St, 4, 1))
            B = Asc(Mid$(St, 5, 1)) * 16777216 + Asc(Mid$(St, 6, 1)) * 65536 + Asc(Mid$(St, 7, 1)) * 256& + Asc(Mid$(St, 8, 1))
            PrintChat "Your guild owes " + CStr(A) + " gold.  This must be payed before " + CStr(CDate(B)) + " or your guild will be disbanded.  Type '/guild pay <amount>' to pay toward the debt.", 15
        ElseIf Len(St) = 13 Then
            A = Asc(Mid$(St, 1, 1)) * 16777216 + Asc(Mid$(St, 2, 1)) * 65536 + Asc(Mid$(St, 3, 1)) * 256& + Asc(Mid$(St, 4, 1))
            B = Asc(Mid$(St, 5, 1)) * 16777216 + Asc(Mid$(St, 6, 1)) * 65536 + Asc(Mid$(St, 7, 1)) * 256& + Asc(Mid$(St, 8, 1))
            C = Asc(Mid$(St, 9, 1))
            D = Asc(Mid$(St, 10, 1)) * 16777216 + Asc(Mid$(St, 11, 1)) * 65536 + Asc(Mid$(St, 12, 1)) * 256& + Asc(Mid$(St, 13, 1))
            If C = Character.index Then
                PrintChat "You have deposited " + CStr(D) + " gold.  Your guild owes " + CStr(A) + " gold.  This must be payed before " + CStr(CDate(B)) + " or your guild will be disbanded.  Type '/guild pay <amount>' to pay toward the debt.", 15
            Else
                PrintChat Player(C).name + " has deposited " + CStr(D) + " gold.  Your guild owes " + CStr(A) + " gold.  This must be payed before " + CStr(CDate(B)) + " or your guild will be disbanded.  Type '/guild pay <amount>' to pay toward the debt.", 15
            End If
        End If

    Case 75    'Guild Deleted
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

    Case 76    'Rank Changed
        If Len(St) = 1 Then
            Character.GuildRank = Asc(Mid$(St, 1, 1))
            PrintChat "Your guild rank has been changed to " + Chr$(34) + Choose(Character.GuildRank + 1, "Initiate", "Member", "Lord", "Founder") + Chr$(34) + ".", 15
        End If

    Case 77    'Invited to join guild
        If Len(St) = 2 Then
            A = Asc(Mid$(St, 1, 1))
            B = Asc(Mid$(St, 2, 1))
            If A >= 1 And B >= 1 Then
                PrintChat "You have been invited to join the guild " + Chr$(34) + Guild(A).name + Chr$(34) + " by " + Player(B).name + ".  If you wish to join, Type /guild join.  It will cost " + CStr(World.GuildJoinCost) + " gold to join this guild.", 15
            End If
        End If

    Case 78    'View Guild Data
        If Len(St) >= 12 Then
            frmMain.picBuy.Visible = False
            frmMain.picDrop.Visible = False
            frmGuild.MemberCount = 0
            frmGuild.CurrentGuild = Asc(Mid$(St, 1, 1))
            frmGuild.CurrentSprite = Asc(Mid$(St, 2, 1)) * 256 + Asc(Mid$(St, 3, 1))
            frmGuild.lblCreated = CDate(Asc(Mid$(St, 4, 1)) * 16777216 + Asc(Mid$(St, 5, 1)) * 65536 + Asc(Mid$(St, 6, 1)) * 256& + Asc(Mid$(St, 7, 1)))
            frmGuild.lblKills = Asc(Mid$(St, 8, 1)) * 16777216 + Asc(Mid$(St, 9, 1)) * 65536 + Asc(Mid$(St, 10, 1)) * 256& + Asc(Mid$(St, 11, 1))
            frmGuild.lblDeaths = Asc(Mid$(St, 12, 1)) * 16777216 + Asc(Mid$(St, 13, 1)) * 65536 + Asc(Mid$(St, 14, 1)) * 256& + Asc(Mid$(St, 15, 1))
            With frmGuild
                .lblName = Guild(frmGuild.CurrentGuild).name
                A = Asc(Mid$(St, 16, 1))
                If A > 0 Then
                    .lblHall = Hall(A).name
                Else
                    .lblHall = "<none>"
                End If
                .lstDeclarations.Clear
                .lstViewMembers.ListItems.Clear
                For A = 0 To 4
                    B = Asc(Mid$(St, 17 + 14 * A))
                    If B > 0 Then
                        C = Asc(Mid$(St, 19 + 14 * A, 1)) * 16777216 + Asc(Mid$(St, 20 + 14 * A, 1)) * 65536 + Asc(Mid$(St, 21 + 14 * A, 1)) * 256& + Asc(Mid$(St, 22 + 14 * A, 1))
                        D = Asc(Mid$(St, 23 + 14 * A, 1)) * 16777216 + Asc(Mid$(St, 24 + 14 * A, 1)) * 65536 + Asc(Mid$(St, 25 + 14 * A, 1)) * 256& + Asc(Mid$(St, 26 + 14 * A, 1))
                        E = Asc(Mid$(St, 27 + 14 * A, 1)) * 16777216 + Asc(Mid$(St, 28 + 14 * A, 1)) * 65536 + Asc(Mid$(St, 29 + 14 * A, 1)) * 256& + Asc(Mid$(St, 30 + 14 * A, 1))
                        If Asc(Mid$(St, 18 + 14 * A)) = 0 Then
                            .lstDeclarations.AddItem "Alliance with " + Guild(B).name + " (Forged " & CDate(C) & ")"
                        Else
                            If Character.Access > 0 Then
                                .lstDeclarations.AddItem "War with " + Guild(B).name + " (Began " & CDate(C) & ", " + CStr(D) + " kills, " + CStr(E) + " deaths)"
                            Else
                                .lstDeclarations.AddItem "War with " + Guild(B).name + " (Began " & CDate(C) & ")"
                            End If
                        End If
                        .lstDeclarations.ItemData(.lstDeclarations.ListCount - 1) = A
                    End If
                Next A
                If Character.Guild = frmGuild.CurrentGuild And Character.GuildRank >= 2 Then
                    If Character.GuildRank = 3 Then
                        .btnDisband.Visible = True
                        .cmdResetStats.Visible = True
                    Else
                        .btnDisband.Visible = False
                        .cmdResetStats.Visible = False
                    End If
                    If .lstDeclarations.ListCount < 5 Then
                        .btnAddDeclaration.Visible = True
                    Else
                        .btnAddDeclaration.Visible = False
                    End If
                    If .lblHall = "<none>" Then
                        .btnMoveOut.Visible = False
                    Else
                        .btnMoveOut.Visible = True
                    End If
                    .cmdBuySprite.Visible = True
                    .cmdSellSprite.Visible = True
                    .sclSprite.Visible = True
                    .btnRemoveDeclaration.Visible = True
                    .btnRemoveMember.Visible = True
                    .btnRank(0).Visible = True
                    .btnRank(1).Visible = True
                    .btnRank(2).Visible = True
                    .btnRank(3).Visible = True
                Else
                    .btnDisband.Visible = False
                    .btnAddDeclaration.Visible = False
                    .btnMoveOut.Visible = False
                    .cmdBuySprite.Visible = False
                    .cmdSellSprite.Visible = False
                    .cmdResetStats.Visible = False
                    .sclSprite.Visible = False
                    .btnRemoveMember.Visible = False
                    .btnRemoveDeclaration.Visible = False
                    .btnRank(0).Visible = False
                    .btnRank(1).Visible = False
                    .btnRank(2).Visible = False
                    .btnRank(3).Visible = False
                End If
            End With
        End If

    Case 79    'Guild Chat
        If Len(St) >= 2 Then
            A = Asc(Mid$(St, 1, 1))
            If A >= 1 Then
                PrintChat Player(A).name + " -> Guild: " + Mid$(St, 2), 15
            End If
        End If

    Case 80    'Created Guild
        If Len(St) = 1 Then
            A = Asc(Mid$(St, 1, 1))
            Character.Guild = A
            Character.GuildRank = 3
            If A > 0 Then
                PrintChat "You have created a new guild called " + Chr$(34) + Guild(A).name + Chr$(34) + ".  To invite other players to your guild, Type '/guild invite <player>'.  You must get atleast two other players to join your guild today or your guild will be disbanded.  For a listing of other available guild commands, type /guild help.", 15
            End If
        End If

    Case 81    'Guild hall change
        If Len(St) = 1 Then
            If Asc(Mid$(St, 1, 1)) = 0 Then
                PrintChat "Your guild now owns a hall!", 15
            Else
                PrintChat "Your guild no longer owns a hall!", 15
            End If
        End If

    Case 82    'Guild hall data
        If Len(St) >= 1 Then
            A = Asc(Mid$(St, 1, 1))
            If A >= 1 Then
                Hall(A).Version = Asc(Mid$(St, 2, 1))
                If Len(St) >= 3 Then
                    Hall(A).name = Mid$(St, 3)
                Else
                    Hall(A).name = vbNullString
                End If
                If frmList_Loaded = True Then
                    frmList.DrawList
                End If
                Debug.Print "Save Hall " + CStr(A)
                SaveHall CByte(A)
            End If
        End If

    Case 83    'Guild Hall Edit Data
        If Len(St) = 13 Then
            A = Asc(Mid$(St, 1, 1))
            If A >= 1 Then
                If frmHall_Loaded = False Then Load frmHall
                With frmHall
                    .lblNumber = A
                    .txtName = Hall(A).name
                    .txtPrice = CStr(Asc(Mid$(St, 2, 1)) * 16777216 + Asc(Mid$(St, 3, 1)) * 65536 + Asc(Mid$(St, 4, 1)) * 256& + Asc(Mid$(St, 5, 1)))
                    .txtUpkeep = CStr(Asc(Mid$(St, 6, 1)) * 16777216 + Asc(Mid$(St, 7, 1)) * 65536 + Asc(Mid$(St, 8, 1)) * 256& + Asc(Mid$(St, 9, 1)))
                    B = Asc(Mid$(St, 10, 1)) * 256 + Asc(Mid$(St, 11, 1))
                    If B < 1 Then B = 1
                    If B > MaxMaps Then B = MaxMaps
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

    Case 84    'Guild Hall Info
        If Len(St) = 10 Then
            A = Asc(Mid$(St, 1, 1))
            If A >= 1 Then
                B = Asc(Mid$(St, 2, 1))
                PrintChat Hall(A).name, 15
                If B > 0 Then
                    PrintChat "Owned By: " + Guild(B).name, 15
                Else
                    PrintChat "This guild hall is not yet owned!", 15
                End If
                A = Asc(Mid$(St, 3, 1)) * 16777216 + Asc(Mid$(St, 4, 1)) * 65536 + Asc(Mid$(St, 5, 1)) * 256& + Asc(Mid$(St, 6, 1))
                PrintChat "Cost: " + CStr(A) + " gold coins", 15
                A = Asc(Mid$(St, 7, 1)) * 16777216 + Asc(Mid$(St, 8, 1)) * 65536 + Asc(Mid$(St, 9, 1)) * 256& + Asc(Mid$(St, 10, 1))
                PrintChat "Upkeep: " + CStr(A) + " gold coins per day", 15
            End If
        End If

    Case 85    'NPC Data
        If Len(St) >= 1 Then
            A = Asc(Mid$(St, 1, 1)) * 256 + Asc(Mid$(St, 2, 1))
            If A >= 1 Then
                With NPC(A)
                    .Version = Asc(Mid$(St, 3, 1))
                    .flags = Asc(Mid$(St, 4, 1))
                    For B = 0 To 9
                        C = 5 + B * 12
                        .SaleItem(B).GiveObject = Asc(Mid$(St, C, 1)) * 256 + Asc(Mid$(St, C + 1, 1))
                        .SaleItem(B).GiveValue = Asc(Mid$(St, C + 2, 1)) * 16777216 + Asc(Mid$(St, C + 3, 1)) * 65536 + Asc(Mid$(St, C + 4, 1)) * 256& + Asc(Mid$(St, C + 5, 1))
                        .SaleItem(B).TakeObject = Asc(Mid$(St, C + 6, 1)) * 256 + Asc(Mid$(St, C + 7, 1))
                        .SaleItem(B).TakeValue = Asc(Mid$(St, C + 8, 1)) * 16777216 + Asc(Mid$(St, C + 9, 1)) * 65536 + Asc(Mid$(St, C + 10, 1)) * 256& + Asc(Mid$(St, C + 11, 1))
                    Next B
                    If Len(St) >= 125 Then
                        .name = Mid$(St, 125)
                    Else
                        .name = vbNullString
                    End If
                    If frmList_Loaded = True Then
                        frmList.DrawList
                    End If
                    If frmMapProperties_Loaded = True Then
                        frmMapProperties.cmbNPC.List(A) = CStr(A) + ": " + .name
                    End If
                    Debug.Print "Save NPC " + CStr(A)
                    SaveNPC CInt(A)
                End With
            End If
        End If

    Case 87    'Edit NPC Data
        If Len(St) >= 2 Then
            A = GetInt(Mid$(St, 1, 2))
            If A >= 1 Then
                B = Asc(Mid$(St, 3, 1))
                '123
                GetSections2 Mid$(St, 4)
                With frmNPC
                    For C = 0 To 2
                        If ExamineBit(CByte(B), CByte(C)) = True Then
                            .chkFlag(C) = 1
                        Else
                            .chkFlag(C) = 0
                        End If
                    Next C
                    .lblNumber = A
                    .txtName = NPC(A).name
                    .txtJoinText = Section(1)
                    .txtLeaveText = Section(2)
                    .txtSayText1 = Section(3)
                    .txtSayText2 = Section(4)
                    .txtSayText3 = Section(5)
                    .txtSayText4 = Section(6)
                    .txtSayText5 = Section(7)
                    .UpdateList
                    .Show
                End With
            End If
        End If

    Case 88    'NPC Talks
        If Len(St) >= 3 Then
            A = GetInt(Mid$(St, 1, 2))
            If A >= 1 Then
                PrintChat NPC(A).name + " says, " + Chr$(34) + Mid$(St, 3) + Chr$(34), 7
            End If
        End If

    Case 89    'Bank Gold
        If Len(St) = 4 Then
            If Map.NPC >= 1 Then
                A = Asc(Mid$(St, 1, 1)) * 16777216 + Asc(Mid$(St, 2, 1)) * 65536 + Asc(Mid$(St, 3, 1)) * 256& + Asc(Mid$(St, 4, 1))
                frmMain.picBank.Visible = True
                frmMain.lblBank = Map.name
                frmMain.lblGoldCoins = CStr(A)
            End If
        End If

    Case 90    'God Chat
        If Len(St) >= 2 Then
            A = Asc(Mid$(St, 1, 1))
            If A >= 1 Then
                PrintChat "<" + Player(A).name + ">: " + Mid$(St, 2), 11
            End If
        End If

    Case 91    'Status Change
        If Len(St) = 2 Then
            A = Asc(Mid$(St, 1, 1))
            If A >= 1 Then
                If A = Character.index Then
                    Character.status = Asc(Mid$(St, 2, 1))
                Else
                    Player(A).status = Asc(Mid$(St, 2, 1))
                End If
                UpdatePlayerColor A
            End If
        End If

    Case 92    'Edit Ban Data
        If Len(St) >= 4 Then
            If frmBan_Loaded = False Then Load frmBan
            GetSections2 Mid$(St, 3)
            With frmBan
                .lblNumber = Asc(Mid$(St, 1, 1))
                .sclUnban = Asc(Mid$(St, 2, 1))
                .txtName = Section(1)
                .txtBanner = Section(2)
                .txtReason = Section(3)
                .txtComputerID = Section(4)
                .txtIPAddress = Section(5)
                If Not Character.name = Section(2) And Not Character.Access >= 3 Then
                    .btnClear.Enabled = False
                    .btnOk.Enabled = False
                Else
                    .btnClear.Enabled = True
                    .btnOk.Enabled = True
                End If
                .Show
            End With
        End If

    Case 94    'Edit Script Data
        If Len(St) >= 3 Then
            A = InStr(St, vbNullChar)
            If A >= 1 Then
                On Error Resume Next
                Load frmScript
                With frmScript
                    .lblName = Left$(St, A - 1)
                    .txtCode.Text = Mid$(St, A + 1)
                    If .txtCode.Text = vbNullString Then
                        St1 = .lblName
                        If St1 Like "MAPSAY*" Then
                            .Scintilla.Text = "FUNCTION Main(Player AS LONG, Message AS String) AS LONG" + Chr$(13) + Chr$(10) + "   Main = Continue" + Chr$(13) + Chr$(10) + "END FUNCTION"
                        ElseIf St1 Like "MAP*" Or St1 Like "MONSTERDIE*" Or St1 Like "JOINMAP*" Or St1 Like "PARTMAP*" Or St1 = "JOINGAME" Or St1 = "PARTGAME" Then
                            .Scintilla.Text = "SUB Main(Player AS LONG)" + Chr$(13) + Chr$(10) + "END SUB"
                        ElseIf St1 Like "USEOBJ*" Or St1 Like "GETOBJ*" Or St1 Like "DROPOBJ*" Or St1 = "PLAYERDIE" Or St1 Like "MONSTERSEE*" Then
                            .Scintilla.Text = "FUNCTION Main(Player AS LONG) AS LONG" + Chr$(13) + Chr$(10) + "   Main = Continue" + Chr$(13) + Chr$(10) + "END FUNCTION"
                        ElseIf St1 Like "MONSTERSEE*" Then
                            .Scintilla.Text = "FUNCTION Main(Player AS LONG) AS LONG" + Chr$(13) + Chr$(10) + "   Main = Continue" + Chr$(13) + Chr$(10) + "END FUNCTION"
                        ElseIf St1 Like "MVMSEE*" Or St1 Like "MVMATTACK*" Or St1 Like "MVMDIE*" Then
                            .Scintilla.Text = "FUNCTION Main(Map AS LONG, Index AS LONG, TargetIndex AS LONG) AS LONG" + Chr$(13) + Chr$(10) + "   Main = Continue" + Chr$(13) + Chr$(10) + "END FUNCTION"
                        ElseIf St1 Like "CLICKMAP*" Then
                            .Scintilla.Text = "SUB Main()" + Chr$(13) + Chr$(10) + "END SUB"
                        ElseIf St1 = "KILLPLAYER" Then
                            .Scintilla.Text = "FUNCTION Main(Killer AS LONG, Killee AS LONG) AS LONG" + Chr$(13) + Chr$(10) + "   Main = Continue" + Chr$(13) + Chr$(10) + "END FUNCTION"
                        ElseIf St1 = "BROADCAST" Then
                            .Scintilla.Text = "FUNCTION Main(Player AS LONG, Message AS String) AS LONG" + Chr$(13) + Chr$(10) + "   Main = Continue" + Chr$(13) + Chr$(10) + "END FUNCTION"
                        ElseIf St1 = "COMMAND" Then
                            .Scintilla.Text = "FUNCTION Main(Player as LONG, Command as String, Parm1 as String, Parm2 as String, Parm3 as String) AS LONG" + Chr$(13) + Chr$(10) + "Main = Continue" + Chr$(13) + Chr$(10) + "END FUNCTION"
                        ElseIf St1 = "PLAYERRESURRECT" Then
                            .Scintilla.Text = "SUB Main(Player AS LONG)" + Chr$(13) + Chr$(10) + "END SUB"
                        ElseIf St1 = "CALCULATESTATS" Then
                            .Scintilla.Text = "SUB Main(Player AS LONG)" + Chr$(13) + Chr$(10) + "END SUB"
                        ElseIf St1 = "MINUTETIMER" Then
                            .Scintilla.Text = "SUB Main()" + Chr$(13) + Chr$(10) + "END SUB"
                        ElseIf St1 = "CHOPLUMBER" Then
                            .Scintilla.Text = "SUB Main(Player AS LONG, Amount AS LONG)" + Chr$(13) + Chr$(10) + "END SUB"
                        ElseIf St1 = "CATCHFISH" Then
                            .Scintilla.Text = "SUB Main(Player AS LONG, Grade AS LONG)" + Chr$(13) + Chr$(10) + "END SUB"
                        ElseIf St1 = "MINEORE" Then
                            .Scintilla.Text = "SUB Main(Player AS LONG, Grade AS LONG)" + Chr$(13) + Chr$(10) + "END SUB"
                        ElseIf St1 = "CLICKPLAYER" Then
                            .Scintilla.Text = "SUB Main(Player AS LONG, Clicked AS LONG)" + Chr$(13) + Chr$(10) + "END SUB"
                        ElseIf St1 = "MONSTERKILL" Then
                            .Scintilla.Text = "FUNCTION Main(Monster AS LONG, MonsterIndex AS LONG, Player AS LONG) AS LONG" + Chr$(13) + Chr$(10) + "   Main = Continue" + Chr$(13) + Chr$(10) + "END FUNCTION"
                        ElseIf St1 Like "ATTACKMONSTER*" Then
                            .Scintilla.Text = "FUNCTION Main(Player AS LONG) AS LONG" + Chr$(13) + Chr$(10) + "   Main = Continue" + Chr$(13) + Chr$(10) + "END FUNCTION"
                        ElseIf St1 Like "SPELL*" Then
                            .Scintilla.Text = "SUB Main(Player AS LONG)" + Chr$(13) + Chr$(10) + "END SUB"
                        ElseIf St1 = "DAYTIMER" Then
                            .Scintilla.Text = "SUB Main()" + Chr$(13) + Chr$(10) + "END SUB"
                        End If
                    Else
                        .Scintilla.Text = .txtCode.Text
                    End If
                    .Show
                End With
            End If
            On Error GoTo HandleError
        End If

    Case 96    'Custom Wav Play
        If Len(St) >= 1 Then
            A = Asc(Mid$(St, 1, 1))
            If Exists("Sound" + CStr(A) + ".wav") Then
                PlayWav A
            End If
        End If

    Case 97    'New Guild Info
        Select Case Asc(Mid$(St, 1, 1))
        Case 0    'Already in a guild.
            PrintChat "You are already in a guild.  If you would like to create a new guild, you must first leave this guild by typing '/guild leave'.", 14
        Case 2    'Guilds are Disabled
            PrintChat "Guilds have been disabled.", 14
        Case 3    'Need to be atleast Level 20
            PrintChat "You must be atleast Level 20 to join a guild!", 14
        Case 4    'Guild sprite taken!
            PrintChat "That guild sprite is taken!", 14
        Case 5    'Need 100k
            PrintChat "You must have 100,000 gold coins in your guild bank to buy a sprite!", 14
        Case 6    'Guild Sprite
            PrintChat "Your guild now has a guild sprite!", 14
        Case 7    'Guild Sprite Taken
            PrintChat "That guild sprite is already taken!", 14
        End Select

    Case 98    'Repairing
        Select Case Asc(Mid$(St, 1, 1))
        Case 2    'Done Repairing Object
            A = GetInt(Mid$(St, 2, 2))
            PrintChat "Your " + Object(A).name + " is now at 100% durability.", 14
            DisplayRepair
        End Select

    Case 99    'Projectiles
        Select Case Asc(Mid$(St, 1, 1))
        Case 1    'Tile Effect
            CreateTileEffect Asc(Mid$(St, 2, 1)), Asc(Mid$(St, 3, 1)), Asc(Mid$(St, 4, 1)), Asc(Mid$(St, 5, 1)) * 256 + Asc(Mid$(St, 6, 1)), Asc(Mid$(St, 7, 1)), Asc(Mid$(St, 8, 1)), Asc(Mid$(St, 9, 1))
        Case 2    'Character Effect
            CreateCharacterEffect Asc(Mid$(St, 2, 1)), Asc(Mid$(St, 3, 1)), Asc(Mid$(St, 4, 1)) * 256 + Asc(Mid$(St, 5, 1)), Asc(Mid$(St, 6, 1)), Asc(Mid$(St, 7, 1)), Asc(Mid$(St, 8, 1))
        Case 3    'Monster Effect
            If Asc(Mid$(St, 2, 1)) = Character.index Then
                A = CX
                B = CY
            Else
                A = Player(Asc(Mid$(St, 2, 1))).X
                B = Player(Asc(Mid$(St, 2, 1))).Y
            End If
            CreateMonsterEffect Asc(Mid$(St, 3, 1)), Asc(Mid$(St, 4, 1)), Asc(Mid$(St, 5, 1)) * 256 + Asc(Mid$(St, 6, 1)), Asc(Mid$(St, 6, 1)), A, B, Asc(Mid$(St, 8, 1))
        Case 4    'Player Effect
            If Asc(Mid$(St, 2, 1)) = Character.index Then
                A = CX
                B = CY
            Else
                A = Player(Asc(Mid$(St, 2, 1))).X
                B = Player(Asc(Mid$(St, 2, 1))).Y
            End If
            CreatePlayerEffect Asc(Mid$(St, 3, 1)), Asc(Mid$(St, 4, 1)), Asc(Mid$(St, 5, 1)) * 256 + Asc(Mid$(St, 6, 1)), Asc(Mid$(St, 7, 1)), A, B, Asc(Mid$(St, 8, 1))
        Case 5    'Projectile
            CreateProjectile Asc(Mid$(St, 2, 1)), Asc(Mid$(St, 3, 1)), Asc(Mid$(St, 4, 1)), Asc(Mid$(St, 5, 1)), Asc(Mid$(St, 6, 1))
        Case 6    'Script Projectile
            CreateProjectile Asc(Mid$(St, 2, 1)), Asc(Mid$(St, 3, 1)), Asc(Mid$(St, 4, 1)), Asc(Mid$(St, 5, 1)), Asc(Mid$(St, 6, 1)), Asc(Mid$(St, 7, 1))
        Case 7    'Magic Script Projectile
            CreateProjectile Asc(Mid$(St, 2, 1)), Asc(Mid$(St, 3, 1)), Asc(Mid$(St, 4, 1)), Asc(Mid$(St, 5, 1)), Asc(Mid$(St, 6, 1)), Asc(Mid$(St, 7, 1)), 1
        End Select
    Case 102    'Display
        St1 = UncompressString$(Mid$(St, 1))
        PrintChat St1, 15
    Case 104    'Scan
        A = Asc(Mid$(St, 1, 1))    'Player

        With frmScan
            .lblPlayer = Player(A).name
            .lblLevel = Asc(Mid$(St, 3, 1))
            .lblClass = Class(Asc(Mid$(St, 2, 1))).name
            .lblMaxHP = Asc(Mid$(St, 8, 1))
            .lblMaxEnergy = Asc(Mid$(St, 10, 1))
            .lblMaxMana = Asc(Mid$(St, 9, 1))

            Dim TempString As String
            For D = 1 To MaxInvObjects
                C = Asc(Mid$(St, D * 6 + 7, 1)) * 16777216 + Asc(Mid$(St, D * 6 + 8, 1)) * 65536 + Asc(Mid$(St, D * 6 + 9, 1)) * 256& + Asc(Mid$(St, D * 6 + 10, 1))
                If GetInt(Mid$(St, D * 6 + 5, 2)) > 0 Then
                    TempString = D & ":  " & Object(GetInt(Mid$(St, D * 6 + 5, 2))).name & " (" & GetInt(Mid$(St, D * 6 + 5, 2)) & ")" & " (" & C & ") "
                    .lstInventory.AddItem TempString
                End If
            Next D
            For D = 1 To 5
                C = Asc(Mid$(St, D * 6 + 127, 1)) * 16777216 + Asc(Mid$(St, D * 6 + 128, 1)) * 65536 + Asc(Mid$(St, D * 6 + 129, 1)) * 256& + Asc(Mid$(St, D * 6 + 130, 1))
                If GetInt(Mid$(St, D * 6 + 125, 2)) > 0 Then
                    TempString = "(E) " & Object(GetInt(Mid$(St, D * 6 + 125, 2))).name & " (" & GetInt(Mid$(St, D * 6 + 125, 2)) & ")" & " (" & C & ") "
                    .lstInventory.AddItem TempString
                End If
            Next D
            For D = 1 To 30
                C = Asc(Mid$(St, D * 6 + 157, 1)) * 16777216 + Asc(Mid$(St, D * 6 + 158, 1)) * 65536 + Asc(Mid$(St, D * 6 + 159, 1)) * 256& + Asc(Mid$(St, D * 6 + 160, 1))
                If GetInt(Mid$(St, D * 6 + 155, 2)) > 0 Then
                    TempString = D & ":  " & Object(GetInt(Mid$(St, D * 6 + 155, 2))).name & " (" & GetInt(Mid$(St, D * 6 + 155, 2)) & ")" & " (" & C & ") "
                    .lstBank.AddItem TempString
                End If
            Next D

            .Show

        End With
    Case 109    'You took damage
        A = Asc(Mid$(St, 1, 1))
        PlayWav 2
        PrintInfoText "You took " & A & " damage!"
        If GetHP > A Then
            SetHP GetHP - A
        Else
            SetHP 0
            YouDied
        End If
        DrawStats
    Case 110    'Someone died from damage tile
        A = Asc(Mid$(St, 1, 1))
        If A = Character.index Then
            PlayWav 8
            PrintChat "You have died!", 12
        Else
            PlayWav 8
            PrintChat Player(A).name & " has died!", 12
            Player(A).IsDead = True
        End If
    Case 111    'Floating Number
        A = Asc(Mid$(St, 1, 1))
        B = Asc(Mid$(St, 2, 1))
        C = Asc(Mid$(St, 3, 1))
        D = Asc(Mid$(St, 4, 1))
        CreateFloatText CStr(B), A, CByte(C), CByte(D)
    Case 112    'Floating Text
        If Len(St) > 4 Then
            A = Asc(Mid$(St, 1, 1))
            B = Asc(Mid$(St, 2, 1))
            C = Asc(Mid$(St, 3, 1))
            CreateFloatText Mid$(St, 4), A, CByte(B), CByte(C)
        End If
    Case 113    'Item Bank Item
        A = Asc(Mid$(St, 1, 1))    'Slot #
        B = GetInt(Mid$(St, 2, 2))
        C = Asc(Mid$(St, 4, 1)) * 16777216 + Asc(Mid$(St, 5, 1)) * 65536 + Asc(Mid$(St, 6, 1)) * 256& + Asc(Mid$(St, 7, 1))
        D = Asc(Mid$(St, 8, 1))
        E = Asc(Mid$(St, 9, 1))
        Character.ItemBank(A).Object = B
        Character.ItemBank(A).value = C
        Character.ItemBank(A).ItemPrefix = D
        Character.ItemBank(A).ItemSuffix = E
        If frmMain.Visible = True Then
            DrawToDC 0, 0, 32, 32, frmMain.ItemBank(A).hDC, DDSObjects, 0, (Object(B).Picture - 1) * 32
            frmMain.ItemBank(A).Refresh
        End If
    Case 114    'Clear Object
        A = Asc(Mid$(St, 1, 1))    'Slot #
        If frmMain.Visible = True Then
            frmMain.ItemBank(A).Picture = Nothing
            frmMain.ItemBank(A).Refresh
        End If
        Character.ItemBank(A).Object = 0
        Character.ItemBank(A).value = 0
        Character.ItemBank(A).ItemPrefix = 0
        Character.ItemBank(A).ItemSuffix = 0
    Case 115    'Equipped Object
        A = GetInt(Mid$(St, 1, 2))
        B = Asc(Mid$(St, 3, 1)) * 16777216 + Asc(Mid$(St, 4, 1)) * 65536 + Asc(Mid$(St, 5, 1)) * 256& + Asc(Mid$(St, 6, 1))
        C = Asc(Mid$(St, 7, 1))
        D = Asc(Mid$(St, 8, 1))
        Select Case Object(A).Type
        Case 1, 10    'Weapon
            Character.EquippedObject(1).Object = A
            Character.EquippedObject(1).value = B
            Character.EquippedObject(1).ItemPrefix = C
            Character.EquippedObject(1).ItemSuffix = D
            If Object(A).Type = 10 Then Character.Projectile = True
        Case 2, 3, 4
            Character.EquippedObject(Object(A).Type).Object = A
            Character.EquippedObject(Object(A).Type).value = B
            Character.EquippedObject(Object(A).Type).ItemPrefix = C
            Character.EquippedObject(Object(A).Type).ItemSuffix = D
        Case 8
            Character.EquippedObject(5).Object = A
            Character.EquippedObject(5).value = B
            Character.EquippedObject(5).ItemPrefix = C
            Character.EquippedObject(5).ItemSuffix = D
        End Select
        If frmMain.picRepair.Visible = True Then
            DisplayRepair
        ElseIf frmMain.picSellObject.Visible = True Then
            DisplaySell
        End If
        RefreshInventory
    Case 116    'Version is Outdated
        MsgBox "Your version does not match the version running on this server!  Try closing the client and running the updater."
        End
    Case 117    'Float Code
        A = Asc(Mid$(St, 1, 1))
        B = Asc(Mid$(St, 2, 1))
        C = Asc(Mid$(St, 3, 1))
        Select Case C
        Case 1    'Miss
            CreateFloatText "Miss!", 14, CByte(A), CByte(B)
        Case 2    'Dead
            CreateFloatText "Dead!", 12, CByte(A), CByte(B)
        Case 3    'Caught something
            D = Asc(Mid$(St, 4, 1))
            CreateFloatText "Caught " + Object(D).name, 11, CByte(A), CByte(B)
        Case 4    'Chopped something
            D = Asc(Mid$(St, 4, 1))
            CreateFloatText "Chopped " + CStr(D) + " lumber!", 11, CByte(A), CByte(B)
        Case 5    'Mined something
            D = Asc(Mid$(St, 4, 1))
            CreateFloatText "Mined " + Object(D).name, 11, CByte(A), CByte(B)
        End Select
    Case 119    'Skill Data
        ProcessSkillData St
    Case 120    'Player Revived
        A = Asc(Mid$(St, 1, 1))
        If Len(St) = 2 Then
            If Character.index = A Then
                Character.IsDead = True
                NextTransition = 6
            Else
                Player(A).IsDead = True
            End If
        Else
            If Character.index = A Then
                Character.IsDead = False
                NextTransition = 6
                SetHP GetMaxHP
                SetEnergy GetMaxEnergy
                SetMana GetMaxMana
                DrawStats
            Else
                Player(A).IsDead = False
            End If
        End If
    Case 122    'Object List
        St1 = ""
        frmWait.lblStatus = "Updating Objects ..."
        frmWait.Refresh
        For A = 1 To MaxObjects
            LoadObject CInt(A)
            If Not Object(A).Version = Asc(Mid$(St, A, 1)) Then
                St1 = St1 + vbNullChar + Chr$(3) + Chr$(79) + DoubleChar$(A)
                If Len(St1) > MaxRequestLength Then
                    SendRaw St1
                    St1 = ""
                    SendSocket Chr$(7) + Chr$(1)
                    Exit For
                End If
            End If
            If A = MaxObjects Then
                If Len(St1) > 0 Then SendRaw St1
                SendSocket Chr$(7) + Chr$(2)
            End If
        Next A
    Case 123    'NPC List
        St1 = ""
        frmWait.lblStatus = "Updating NPCs ..."
        frmWait.Refresh
        For A = 1 To MaxNPCs
            LoadNPC CInt(A)
            If Not NPC(A).Version = Asc(Mid$(St, A, 1)) Then
                St1 = St1 + vbNullChar + Chr$(3) + Chr$(80) + DoubleChar$(A)
                If Len(St1) > MaxRequestLength Then
                    SendRaw St1
                    St1 = ""
                    SendSocket Chr$(7) + Chr$(2)
                    Exit For
                End If
            End If
            If A = MaxNPCs Then
                If Len(St1) > 0 Then SendRaw St1
                SendSocket Chr$(7) + Chr$(3)
            End If
        Next A
    Case 124    'Hall List
        St1 = ""
        frmWait.lblStatus = "Updating Halls ..."
        frmWait.Refresh
        For A = 1 To MaxHalls
            LoadHall CInt(A)
            If Not Hall(A).Version = Asc(Mid$(St, A, 1)) Then
                St1 = St1 + vbNullChar + Chr$(2) + Chr$(81) + Chr$(A)
                If Len(St1) > MaxRequestLength Then
                    SendRaw St1
                    St1 = ""
                    SendSocket Chr$(7) + Chr$(3)
                    Exit For
                End If
            End If
            If A = MaxHalls Then
                If Len(St1) > 0 Then SendRaw St1
                SendSocket Chr$(7) + Chr$(4)
            End If
        Next A
    Case 125    'Monster List
        St1 = ""
        frmWait.lblStatus = "Updating Monsters ..."
        frmWait.Refresh
        For A = 1 To MaxTotalMonsters
            LoadMonster CInt(A)
            If Not Monster(A).Version = Asc(Mid$(St, A, 1)) Then
                St1 = St1 + vbNullChar + Chr$(3) + Chr$(82) + DoubleChar$(A)
                If Len(St1) > MaxRequestLength Then
                    SendRaw St1
                    St1 = ""
                    SendSocket Chr$(7) + Chr$(4)
                    Exit For
                End If
            End If
            If A = MaxTotalMonsters Then
                If Len(St1) > 0 Then SendRaw St1
                SendSocket Chr$(7) + Chr$(5)
            End If
        Next A
    Case 126    'Magic List
        St1 = ""
        frmWait.lblStatus = "Updating Magic ..."
        frmWait.Refresh
        For A = 1 To MaxMagic
            LoadMagic CInt(A)
            If Not Magic(A).Version = Asc(Mid$(St, A, 1)) Then
                St1 = St1 + vbNullChar + Chr$(3) + Chr$(83) + DoubleChar$(A)
                If Len(St1) > MaxRequestLength Then
                    SendRaw St1
                    St1 = ""
                    SendSocket Chr$(7) + Chr$(5)
                    Exit For
                End If
            End If
            If A = MaxMagic Then
                If Len(St1) > 0 Then SendRaw St1
                SendSocket Chr$(7) + Chr$(6)
            End If
        Next A
    Case 127    'Magic Data
        If Len(St) >= 3 Then
            A = GetInt(Mid$(St, 1, 2))
            If A >= 1 Then
                With Magic(A)
                    .Class = Asc(Mid$(St, 3, 1))
                    .Level = Asc(Mid$(St, 4, 1))
                    .Version = Asc(Mid$(St, 5, 1))
                    .Icon = Asc(Mid$(St, 6, 1)) * 256 + Asc(Mid$(St, 7, 1))
                    .IconType = Asc(Mid$(St, 8, 1))
                    .CastTimer = Asc(Mid$(St, 9, 1)) * 256 + Asc(Mid$(St, 10, 1))
                    B = InStr(11, St, Chr$(0))
                    If B > 11 And B <= Len(St) Then
                        .name = Mid$(St, 11, B - 11)
                        If Not B = Len(St) Then .Description = Mid$(St, B + 1)
                        If frmList_Loaded = True Then
                            frmList.DrawList
                        End If
                        DrawMagicList
                    End If
                End With
                Debug.Print "Save Magic " + CStr(A)
                SaveMagic CInt(A)
            End If
        End If
    Case 128    'Edit Magic
        A = GetInt(Mid$(St, 1, 2))
        If frmMagic_Loaded = False Then Load frmMagic
        With frmMagic
            If Magic(A).Level > 0 Then .sclLevel.value = Magic(A).Level
            .lblLevel = .sclLevel.value
            For B = 0 To NumClasses - 1
                If ExamineBit(Magic(A).Class, CByte(B)) = True Then
                    .chkClass(B).value = 1
                Else
                    .chkClass(B).value = 0
                End If
            Next B
            .txtName = Magic(A).name
            .txtDescription = Magic(A).Description
            '.sclIcon = Magic(A).Icon
            '.sclCastTimer = Magic(A).CastTimer
            '.optIconType(Magic(A).IconType).value = True
            .lblNumber = CStr(A)
            .Show
        End With
    Case 130    'Stats Update
        With Character
            SetMaxHP Asc(Mid$(St, 1, 1))
            SetMaxEnergy Asc(Mid$(St, 2, 1))
            SetMaxMana Asc(Mid$(St, 3, 1))
            .PhysicalAttack = Asc(Mid$(St, 4, 1))
            .PhysicalDefense = Asc(Mid$(St, 5, 1))
            .MagicDefense = Asc(Mid$(St, 6, 1))

            DrawStats
        End With
    Case 131    'Prefix List
        St1 = ""
        frmWait.lblStatus = "Updating Prefixes ..."
        frmWait.Refresh
        For A = 1 To MaxModifications
            LoadPrefix CInt(A)
            If Not ItemPrefix(A).Version = Asc(Mid$(St, A, 1)) Then
                St1 = St1 + vbNullChar + Chr$(2) + Chr$(84) + Chr$(A)
                If Len(St1) > MaxRequestLength Then
                    SendRaw St1
                    St1 = ""
                    SendSocket Chr$(7) + Chr$(6)
                    Exit For
                End If
            End If
            If A = MaxModifications Then
                If Len(St1) > 0 Then SendRaw St1
                SendSocket Chr$(7) + Chr$(7)
            End If
        Next A
    Case 132    'Suffix List
        St1 = ""
        frmWait.lblStatus = "Updating Suffixes ..."
        frmWait.Refresh
        For A = 1 To MaxModifications
            LoadSuffix CInt(A)
            If Not ItemSuffix(A).Version = Asc(Mid$(St, A, 1)) Then
                St1 = St1 + vbNullChar + Chr$(2) + Chr$(85) + Chr$(A)
                If Len(St1) > MaxRequestLength Then
                    SendRaw St1
                    St1 = ""
                    SendSocket Chr$(7) + Chr$(7)
                    Exit For
                End If
            End If
            If A = MaxModifications Then
                If Len(St1) > 0 Then SendRaw St1
                SendSocket Chr$(7) + Chr$(8)
            End If
        Next A
    Case 133    'Prefix
        A = Asc(Mid$(St, 1, 1))
        With ItemPrefix(A)
            .ModificationType = Asc(Mid$(St, 2, 1))
            .ModificationValue = Asc(Mid$(St, 3, 1))
            .OccursNaturally = Asc(Mid$(St, 4, 1))
            .Version = Asc(Mid$(St, 5, 1))
            If Len(St) > 5 Then .name = Mid$(St, 6)
            If frmList_Loaded = True Then
                frmList.DrawList
            End If
            Debug.Print "Save Prefix " + CStr(A)
            SavePrefix (CByte(A))
        End With
    Case 134    'Suffix
        A = Asc(Mid$(St, 1, 1))
        With ItemSuffix(A)
            .ModificationType = Asc(Mid$(St, 2, 1))
            .ModificationValue = Asc(Mid$(St, 3, 1))
            .OccursNaturally = Asc(Mid$(St, 4, 1))
            .Version = Asc(Mid$(St, 5, 1))
            If Len(St) > 5 Then .name = Mid$(St, 6)
            If frmList_Loaded = True Then
                frmList.DrawList
            End If
            Debug.Print "Save Suffix " + CStr(A)
            SaveSuffix (CByte(A))
        End With
    Case 135    'Edit Prefix
        A = Asc(Mid$(St, 1, 1))
        Load frmPrefix
        With frmPrefix
            .txtName = ItemPrefix(A).name
            .optModType(ItemPrefix(A).ModificationType).value = True
            .sclValue = ItemPrefix(A).ModificationValue
            .lblModValue = .sclValue
            .chkOccursNaturally.value = ItemPrefix(A).OccursNaturally
            .lblNumber = A
            .Show
        End With
    Case 136    'Edit Suffix
        A = Asc(Mid$(St, 1, 1))
        Load frmSuffix
        With frmSuffix
            .txtName = ItemSuffix(A).name
            .optModType(ItemSuffix(A).ModificationType).value = True
            .sclValue = ItemSuffix(A).ModificationValue
            .lblModValue = .sclValue
            .chkOccursNaturally.value = ItemSuffix(A).OccursNaturally
            .lblNumber = A
            .Show
        End With
    Case 139    'Server Stats
        World.StatStrength = Asc(Mid$(St, 1, 1))
        World.StatEndurance = Asc(Mid$(St, 2, 1))
        World.StatIntelligence = Asc(Mid$(St, 3, 1))
        World.StatConcentration = Asc(Mid$(St, 4, 1))
        World.StatConstitution = Asc(Mid$(St, 5, 1))
        World.StatStamina = Asc(Mid$(St, 6, 1))
        World.StatWisdom = Asc(Mid$(St, 7, 1))
        World.ObjMoney = Asc(Mid$(St, 8, 1))
        World.Cost_Per_Durability = Asc(Mid$(St, 9, 1)) * 256 + Asc(Mid$(St, 10, 1))
        World.Cost_Per_Strength = Asc(Mid$(St, 11, 1)) * 256 + Asc(Mid$(St, 12, 1))
        World.Cost_Per_Modifier = Asc(Mid$(St, 13, 1)) * 256 + Asc(Mid$(St, 14, 1))
        World.GuildJoinLevel = Asc(Mid$(St, 15, 1))
        World.GuildNewLevel = Asc(Mid$(St, 16, 1))
        World.GuildJoinCost = Asc(Mid$(St, 17, 1)) * 16777216 + Asc(Mid$(St, 18, 1)) * 65536 + Asc(Mid$(St, 19, 1)) * 256& + Asc(Mid$(St, 20, 1))
        World.GuildNewCost = Asc(Mid$(St, 21, 1)) * 16777216 + Asc(Mid$(St, 22, 1)) * 65536 + Asc(Mid$(St, 23, 1)) * 256& + Asc(Mid$(St, 24, 1))
    Case 140    'Done sending everything
        SendSocket Chr$(6)
    Case 142    'Monster HP
        A = Asc(Mid$(St, 1, 1))
        B = GetInt(Mid$(St, 2, 2))
        Map.Monster(A).Life = CInt(B)
    Case 143    'Outdoor Conditions
        A = Asc(Mid$(St, 1, 1))    'Light level
        OutdoorLight = A
        If options.DisableLighting = False Then
            UpdateLights
            CreateLightMap Lighting(0), Darkness, MapDataLoadingArray(0), OutdoorLight
        End If
    Case 144    'Add Guild Member
        A = Asc(Mid$(St, 1, 1))
        B = Asc(Mid$(St, 2, 1))
        C = Asc(Mid$(St, 3, 1)) * 16777216 + Asc(Mid$(St, 4, 1)) * 65536 + Asc(Mid$(St, 5, 1)) * 256& + Asc(Mid$(St, 6, 1))
        D = Asc(Mid$(St, 7, 1)) * 16777216 + Asc(Mid$(St, 8, 1)) * 65536 + Asc(Mid$(St, 9, 1)) * 256& + Asc(Mid$(St, 10, 1))
        E = Asc(Mid$(St, 11, 1)) * 16777216 + Asc(Mid$(St, 12, 1)) * 65536 + Asc(Mid$(St, 13, 1)) * 256& + Asc(Mid$(St, 14, 1))
        St1 = Mid$(St, 15)
        frmGuild.AddMember St1, A, B, C, D, E
    Case 145    'Class Changed
        A = Asc(Mid$(St, 1, 1))
        Character.Class = A
        DrawMagicList
    Case 146    'Change Direction
        A = Asc(Mid$(St, 1, 1))
        B = Asc(Mid$(St, 2, 1))
        If A = Character.index Then
            CDir = B
        Else
            Player(A).D = B
        End If
    Case 147    'Map Warp (same map)
        A = Asc(Mid$(St, 1, 1))
        B = Asc(Mid$(St, 2, 1))
        C = Asc(Mid$(St, 3, 1))
        CX = A
        CXO = CX * 32
        CY = B
        CYO = CY * 32
        CDir = C
        Freeze = False
    Case 148    'Static Text
        A = Asc(Mid$(St, 1, 1))
        B = Asc(Mid$(St, 2, 1))
        C = Asc(Mid$(St, 3, 1))
        CreateStaticText Mid$(St, 4), A, CByte(B), CByte(C)
    Case 149    'Pong
        Ping = Tick - PingSent
    Case 150    'Update HP
        A = Asc(Mid$(St, 1, 1))    'Index
        B = Asc(Mid$(St, 2, 1))    'HP
        Player(A).HP = B
    Case 151    'Update MaxHP
        A = Asc(Mid$(St, 1, 1))    'Index
        B = Asc(Mid$(St, 2, 1))    'MaxHP
        Player(A).MaxHP = B
    Case 152    'Guild Balance
        If Len(St) = 8 Then
            A = Asc(Mid$(St, 1, 1)) * 16777216 + Asc(Mid$(St, 2, 1)) * 65536 + Asc(Mid$(St, 3, 1)) * 256& + Asc(Mid$(St, 4, 1))
            B = Asc(Mid$(St, 5, 1)) * 16777216 + Asc(Mid$(St, 6, 1)) * 65536 + Asc(Mid$(St, 7, 1)) * 256& + Asc(Mid$(St, 8, 1))
            PrintChat "Your guild has " + CStr(A) + " gold in the bank.  Your guild upkeep is " + CStr(B) + " per day.", 15
        ElseIf Len(St) = 13 Then
            A = Asc(Mid$(St, 1, 1)) * 16777216 + Asc(Mid$(St, 2, 1)) * 65536 + Asc(Mid$(St, 3, 1)) * 256& + Asc(Mid$(St, 4, 1))
            B = Asc(Mid$(St, 5, 1)) * 16777216 + Asc(Mid$(St, 6, 1)) * 65536 + Asc(Mid$(St, 7, 1)) * 256& + Asc(Mid$(St, 8, 1))
            C = Asc(Mid$(St, 9, 1))
            D = Asc(Mid$(St, 10, 1)) * 16777216 + Asc(Mid$(St, 11, 1)) * 65536 + Asc(Mid$(St, 12, 1)) * 256& + Asc(Mid$(St, 13, 1))
            If C = Character.index Then
                PrintChat "You have deposited " + CStr(D) + " gold.  Your guild has " + CStr(A) + " gold in the bank.  Your guild upkeep is " + CStr(B) + " per day.", 15
            Else
                PrintChat Player(C).name + " has deposited " + CStr(D) + " gold.  Your guild has " + CStr(A) + " gold in the bank.  Your guild upkeep is " + CStr(B) + " per day.", 15
            End If
        End If
    Case 153    'Magic Level Data
        ProcessMagicData St
    Case 154    'Client Data Report
        A = Asc(Mid$(St, 1, 1))
        St1 = CompressString$("Ping: " + CStr(Ping) + "   Energy: " + CStr(GetEnergy()) + "/" + CStr(GetMaxEnergy()))
        SendSocket Chr$(24) + Chr$(A) + St1
    Case 155    'Monster hit monster
        If Len(St) = 5 Then
            A = Asc(Mid$(St, 1, 1)) 'Monster
            B = Asc(Mid$(St, 2, 1)) 'Victim
            D = Asc(Mid$(St, 3, 1)) 'Damage
            C = GetInt(Mid$(St, 4, 2)) 'Victim's HP
            
            If A <= MaxMonsters Then
                Map.Monster(A).A = 5
            End If
            
            If B <= MaxMonsters Then
                Map.Monster(B).Life = C
            End If
            
            CreateFloatText CStr(D), BRIGHTRED, Map.Monster(B).X, Map.Monster(B).Y
        End If
    Case 156 'Uncompressed Map Data
    
    Case 255    'Pong (Do Nothing)

    End Select

    Exit Sub

HandleError:
    If PacketID > 0 Then
        SendSocket Chr$(100) + "Error " & CStr(PacketID) & "  -  " & Err.Description
        MsgBox "Error " & CStr(PacketID) & "  -  " & Err.Description
    End If
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
            If frmMain_Loaded = True Then
                CloseClientSocket 1
            End If
        Case FD_CONNECT
            If lParam = FD_CONNECT Then
                St1 = GetComputerID
                SendSocket Chr$(61) + Chr$(ClientVer) + Chr$(CheckSum(St1) Mod 256) + St1
                If NewAccount = True Then
                    frmWait.lblStatus = "Sending Account Information ..."
                    SendSocket vbNullChar + User + vbNullChar + Pass
                Else
                    frmWait.lblStatus = "Sending Login Information ..."
                    SendSocket Chr$(1) + User + vbNullChar + Pass
                End If
            Else
                If frmWait.Visible = True Then
                    CloseClientSocket 5
                    WaitForConnect "Error Connecting - Waiting"
                End If
            End If
        Case FD_READ
            If lParam = FD_READ Then ReceiveData
        End Select
    End If
    WindowProc = CallWindowProc(lpPrevWndProc, hw, uMsg, wParam, lParam)
End Function
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

Sub ProcessReceivedMap(St As String)
    On Error GoTo FailedToLoad

    RequestedMap = False

    MapData = UncompressString$(St)

    Dim MapDataWorkingArray() As Byte
    MapDataWorkingArray() = StrConv(MapData, vbFromUnicode)
    EncryptDataString MapDataWorkingArray(0), CMap * 16 Mod 50 + 5
    MapData = StrConv(MapDataWorkingArray, vbUnicode)

    On Error Resume Next
    Close #1
    On Error GoTo FailedToLoad

    Open CacheDirectory + "/cache1.dat" For Random As #1 Len = 2677
    Put #1, CMap, MapData
    Close #1

    LoadMapFromCache CMap

    If RequestedMap = False Then ShowMap

    Exit Sub

FailedToLoad:
    PrintChat "Map failed to load - Showing Anyways", YELLOW
    ShowMap
    Freeze = False

End Sub
