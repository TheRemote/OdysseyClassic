Attribute VB_Name = "modOdyssey"
Option Explicit

Public Const TitleString = "The Odyssey Online Classic"
Public Const TheWebSite = "http://www.odysseyclassic.com/"
Public Const ClientVer = 201

Public ServerIP As String
Public ServerPort As Long

Sub SendRaw(ByVal St As String)
    SendSocket Chr$(170) + St
End Sub

Sub WaitForTerm(pid As Long)
    Dim phnd As Long
    phnd = OpenProcess(SYNCHRONIZE, 0, pid)
    If phnd <> 0 Then
        Call WaitForSingleObject(phnd, INFINITE)
        Call CloseHandle(phnd)
    End If
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
Sub MonsterDied(index As Long)
    Dim A As Long
    
    For A = 1 To MaxProjectiles
        With Projectile(A)
            If .Sprite > 0 Then
                If .TargetType = pttMonster Then
                    If .TargetNum = index Then
                        DestroyEffect A
                    End If
                End If
            End If
        End With
    Next A
End Sub

Sub DisplayRepair()
    If frmMain.Visible = False Then Exit Sub
    
    Dim A As Long, B As Long, C As Long, D As Long, St As String, St2 As String
    If Map.NPC > 0 Then
        If ExamineBit(NPC(Map.NPC).flags, 1) = False Then
            frmMain.picRepair.Visible = False
            Exit Sub
        End If
    Else
        frmMain.picRepair.Visible = False
        Exit Sub
    End If
    
    frmMain.picRepair.Visible = True
    frmMain.lblRepairNPCName = Trim$(NPC(Map.NPC).name)
    frmMain.picRepairObjectDisp.Cls

    If CurInvObj = 0 Then
        frmMain.lblRepairName.Caption = "None"
        frmMain.lblRepairCost.Caption = "None"
        frmMain.lblRepairCondition.Caption = "None"
        frmMain.lblRepairDurability.Caption = "None"
        frmMain.lblMenu(56).Visible = False
        frmMain.lblMenu(16).Visible = False
        St = "Hail, " + Character.name + "!  Please select the object you would like to repair."
        frmMain.lblSellNPCTalk = St
        Exit Sub
    End If

    If CurInvObj <= 20 Then
        D = Character.Inv(CurInvObj).Object
    Else
        D = Character.EquippedObject(CurInvObj - 20).Object
    End If

    If D = 0 Then
        frmMain.lblRepairName.Caption = "None"
        frmMain.lblRepairCost.Caption = "None"
        frmMain.lblRepairCondition.Caption = "None"
        frmMain.lblRepairDurability.Caption = "None"
        frmMain.lblMenu(56).Visible = False
        frmMain.lblMenu(16).Visible = False
        St = "Hail, " + Character.name + "!  Please select the object you would like to repair."
        frmMain.lblRepairNpcTalk = St
        Exit Sub
    End If

    frmMain.lblMenu(56).Visible = True
    frmMain.lblMenu(16).Visible = True
    frmMain.lblRepairName = Trim$(Object(D).name)
    DrawToDC 0, 0, 32, 32, frmMain.picRepairObjectDisp.hDC, DDSObjects, 0, (Object(D).Picture - 1) * 32

    A = GetRepairCost(CInt(CurInvObj))
    If A > 0 Then
        B = GetObjectDur(CurInvObj)
        Select Case B
        Case Is >= 85
            St = "Excellent"
            C = QBColor(11)
        Case Is >= 60
            St = "Fair"
            C = QBColor(14)
        Case Is >= 45
            St = "Average"
            C = QBColor(2)
        Case Is >= 25
            St = "Seen Better Days"
            C = QBColor(4)
        Case Is >= 0
            St = "About to Break"
            C = QBColor(5)
        End Select
    Else
        If ExamineBit(Object(D).flags, 1) = 255 Then
            St = "Indestructible"
            B = 100
            C = QBColor(11)
            A = 0
            frmMain.lblMenu(56).Visible = False
        ElseIf ExamineBit(Object(D).flags, 0) = 255 Or Object(D).SellPrice = 0 Then
            B = GetObjectDur(CurInvObj)
            St = "Cannot Repair"
            C = QBColor(4)
            A = 0
            frmMain.lblMenu(56).Visible = False
        Else
            St = "Perfect"
            B = 100
            C = QBColor(11)
            A = 0
            frmMain.lblMenu(56).Visible = False
        End If
    End If

    frmMain.lblRepairDurability = B & "%"
    frmMain.lblRepairCondition = St
    frmMain.lblRepairCondition.ForeColor = C
    frmMain.lblRepairCost = CStr(A) + " gold coins."
    If A > 0 Then
        St2 = "Hail, " + Character.name + "!  I can repair your " + Trim$(Object(D).name) + " for " + CStr(A) + " gold coins." + vbCrLf + vbCrLf
    Else
        St2 = "Hail, " + Character.name + "!  Your " + Trim$(Object(D).name) + " is in perfect condition and does not need to be repaired." + vbCrLf + vbCrLf
    End If

    If GetRepairAllCost > 0 Then
        St2 = St2 + "I can repair all your equipment for " & GetRepairAllCost & " gold pieces."
    Else
        frmMain.lblMenu(16).Visible = False
    End If
    frmMain.lblRepairNpcTalk = St2
    frmMain.picRepair.Visible = True
End Sub

Sub DisplaySell()
    If frmMain.Visible = False Then Exit Sub

    Dim St As String
    Dim TheObject As Long, SellPrice As Long, SellAllAmount As Long, SellAllOffer As Long

    If Map.NPC > 0 Then
        If ExamineBit(NPC(Map.NPC).flags, 2) = False Then
            frmMain.picSellObject.Visible = False
            Exit Sub
        End If
    Else
        frmMain.picSellObject.Visible = False
        Exit Sub
    End If

    frmMain.picRepair.Visible = False
    frmMain.picSellObject.Visible = True
    frmMain.lblSellNPCName = Trim$(NPC(Map.NPC).name)
    frmMain.picSellObjectDisp.Cls

    If CurInvObj = 0 Then
        frmMain.lblSellName.Caption = "None"
        frmMain.lblSellPrice.Caption = "None"
        St = "Hail, " + Character.name + "!  Please select the object you would like to sell."
        frmMain.lblSellNPCTalk = St
        Exit Sub
    End If

    If CurInvObj > 20 Then
        frmMain.lblSellName.Caption = "None"
        frmMain.lblSellPrice.Caption = "None"
        St = "Hail, " + Character.name + "!  You must unequip an item before you can sell it."
        frmMain.lblSellNPCTalk = St
        Exit Sub
    End If

    If Character.Inv(CurInvObj).Object = 0 Then
        frmMain.lblSellName.Caption = "None"
        frmMain.lblSellPrice.Caption = "None"
        St = "Hail, " + Character.name + "!  Please select the object you would like to sell."
        frmMain.lblSellNPCTalk = St
        Exit Sub
    End If

    TheObject = Character.Inv(CurInvObj).Object
    SellPrice = GetSellPrice(CurInvObj)
    frmMain.lblSellName.Caption = Object(TheObject).name
    DrawToDC 0, 0, 32, 32, frmMain.picSellObjectDisp.hDC, DDSObjects, 0, (Object(TheObject).Picture - 1) * 32

    If SellPrice > 0 Then
        frmMain.lblSellPrice = CStr(SellPrice) + " gold coins each."
        frmMain.lblMenu(19).Visible = True
    Else
        St = "Hail, " + Character.name + "!  I am afraid I have no interest in buying your " + Object(TheObject).name + "."
        frmMain.lblSellNPCTalk = St
        frmMain.lblSellPrice = "Cannot Sell"
        frmMain.lblMenu(19).Visible = False
        frmMain.lblSellNPCTalk = St
        Exit Sub
    End If

    Select Case Object(TheObject).Type
    Case 6, 11
        SellAllAmount = Character.Inv(CurInvObj).value
        SellAllOffer = SellAllAmount * Object(TheObject).SellPrice
        frmMain.lblMenu(59).Visible = True
    Case Else
        frmMain.lblMenu(59).Visible = False
    End Select

    St = "Hail, " + Character.name + "!  I will pay you " + CStr(SellPrice) + " gold coins for your " + Object(TheObject).name + "." + vbCrLf + vbCrLf
    If SellAllAmount > 0 Then
        St = St + "I will also buy all " + CStr(SellAllAmount) & " of your " + Object(TheObject).name + "s for " + CStr(SellAllOffer) + " gold coins."
    End If
    frmMain.lblSellNPCTalk = St
End Sub

Function GetObjectDur(ByVal Slot As Long) As Long
    Dim Percent As Single
    Dim Display As Boolean
    If Slot = 0 Then Exit Function

    If Slot <= 20 Then
        If Character.Inv(Slot).Object = 0 Then Exit Function
        If Object(Character.Inv(Slot).Object).MaxDur = 0 Then Exit Function

        Select Case Object(Character.Inv(Slot).Object).Type
        Case 1, 2, 3, 4
            Percent = Character.Inv(Slot).value / (Object(Character.Inv(Slot).Object).MaxDur * 10)
            Percent = Int(Percent * 100)
            If Percent > 100 Then Percent = 100
            Display = True
        Case 8    'Ring
            Percent = Character.Inv(Slot).value / (Object(Character.Inv(Slot).Object).MaxDur * 10)
            Percent = Int(Percent * 100)
            If Percent > 100 Then Percent = 100
            Display = True
        Case Else
            Display = False
        End Select
    Else
        Slot = Slot - 20
        If Character.EquippedObject(Slot).Object = 0 Then Exit Function
        If Object(Character.EquippedObject(Slot).Object).MaxDur = 0 Then Exit Function

        Select Case Object(Character.EquippedObject(Slot).Object).Type
        Case 1, 2, 3, 4
            Percent = Character.EquippedObject(Slot).value / (Object(Character.EquippedObject(Slot).Object).MaxDur * 10)
            Percent = Int(Percent * 100)
            If Percent > 100 Then Percent = 100
            Display = True
        Case 8    'Ring
            Percent = Character.EquippedObject(Slot).value / (Object(Character.EquippedObject(Slot).Object).MaxDur * 10)
            Percent = Int(Percent * 100)
            If Percent > 100 Then Percent = 100
            Display = True
        Case Else
            Display = False
        End Select
    End If

    If Display = True Then
        GetObjectDur = Percent
    Else
        GetObjectDur = 0
    End If
End Function

Function GetDurStringFromValue(ByVal value As Long, ByVal MaxValue As Long) As String
    Dim Percent As Single
    If MaxValue = 0 Then Exit Function
    Percent = value / (MaxValue * 10)
    Percent = Int(Percent * 100)
    If Percent > 100 Then Percent = 100

    Select Case Percent
    Case Is >= 85
        GetDurStringFromValue = "Excellent (" + CStr(CLng(Percent)) + "%)"
    Case Is >= 60
        GetDurStringFromValue = "Fair (" + CStr(CLng(Percent)) + "%)"
    Case Is >= 45
        GetDurStringFromValue = "Average (" + CStr(CLng(Percent)) + "%)"
    Case Is >= 20
        GetDurStringFromValue = "Seen Better Days (" + CStr(CLng(Percent)) + "%)"
    Case Is >= 0
        GetDurStringFromValue = "About to Break (" + CStr(CLng(Percent)) + "%)"
    End Select
End Function

Function GetRepairCost(Slot As Integer) As Long
    If Slot = 0 Then Exit Function

    If Slot <= 20 Then
        If Character.Inv(Slot).Object = 0 Then Exit Function
    Else
        If Character.EquippedObject(Slot - 20).Object = 0 Then Exit Function
    End If

    Dim A As Long, B As Long, C As Long
    If Slot >= 1 And Slot <= 20 Then
        Select Case Object(Character.Inv(Slot).Object).Type
        Case 1, 2, 3, 4, 8    'Weapon, Shield, Armor, Helmet, Ring
            A = Object(Character.Inv(Slot).Object).Type

            If ExamineBit(Object(Character.Inv(Slot).Object).flags, 0) Or ExamineBit(Object(Character.Inv(Slot).Object).flags, 1) Or Object(Character.Inv(Slot).Object).SellPrice = 0 Then
                A = 0
            End If

        Case Else
            A = 0
        End Select

        If A > 0 Then
            Select Case A
            Case 1, 2, 3, 4, 8    'Weapon, Shield, Armor, Helmet, Ring
                'C = Object(Character.Inv(Slot).Object).MaxDur * 10 - (Character.Inv(Slot).value)
                'B = B + (C * World.Cost_Per_Durability)
                'B = B + (C * (Object(Character.Inv(Slot).Object).Modifier * World.Cost_Per_Strength))
                'If B > 0 Then B = B / 100
                If Object(Character.Inv(Slot).Object).MaxDur * 10 > 0 Then
                    C = Object(Character.Inv(Slot).Object).SellPrice - ((Character.Inv(Slot).value / (Object(Character.Inv(Slot).Object).MaxDur * 10)) * Object(Character.Inv(Slot).Object).SellPrice)
                    If C >= 0 Then
                        GetRepairCost = C
                    Else
                        GetRepairCost = 0
                    End If
                Else
                    GetRepairCost = 0
                End If
                Exit Function
            End Select
        Else
            GetRepairCost = 0
        End If
    Else
        Slot = Slot - 20
        Select Case Object(Character.EquippedObject(Slot).Object).Type
        Case 1, 2, 3, 4, 8    'Weapon, Shield, Armor, Helmet, Ring
            A = Object(Character.EquippedObject(Slot).Object).Type

            If ExamineBit(Object(Character.EquippedObject(Slot).Object).flags, 0) Or ExamineBit(Object(Character.EquippedObject(Slot).Object).flags, 1) Then
                A = 0
            End If

        Case Else
            A = 0
        End Select

        If A > 0 Then
            Select Case A
            Case 1, 2, 3, 4, 8    'Weapon, Shield, Armor, Helmet, Ring
                'C = Object(Character.EquippedObject(Slot).Object).MaxDur * 10 - (Character.EquippedObject(Slot).value)
                'B = B + (C * World.Cost_Per_Durability)
                'B = B + (C * (Object(Character.EquippedObject(Slot).Object).Modifier * World.Cost_Per_Strength))
                'If B > 0 Then B = B / 100
                If Object(Character.EquippedObject(Slot).Object).MaxDur * 10 > 0 Then
                    C = Object(Character.EquippedObject(Slot).Object).SellPrice - (Character.EquippedObject(Slot).value / (Object(Character.EquippedObject(Slot).Object).MaxDur * 10) * Object(Character.EquippedObject(Slot).Object).SellPrice)
                    If C >= 0 Then
                        GetRepairCost = C
                    Else
                        GetRepairCost = 0
                    End If
                Else
                    GetRepairCost = 0
                End If
                Exit Function
            End Select
        Else
            GetRepairCost = 0
        End If
    End If
End Function
Function GetRepairAllCost() As Long
    Dim A As Long
    For A = 1 To 25
        GetRepairAllCost = GetRepairAllCost + GetRepairCost(CInt(A))
    Next A
End Function
Sub PlayerLeftMap(index As Long)

    Player(index).Map = 0
    Player(index).HP = 0

    Dim A As Long
    For A = 1 To MaxProjectiles
        With Projectile(A)
            If .Sprite > 0 Then
                If .TargetType = pttCharacter Then
                    If .TargetNum = index Then
                        DestroyEffect A
                    End If
                End If
            End If
        End With
    Next A
End Sub
Function QuadChar(num As Long) As String
    QuadChar = Chr$(Int(num / 16777216) Mod 256) + Chr$(Int(num / 65536) Mod 256) + Chr$(Int(num / 256) Mod 256) + Chr$(num Mod 256)
End Function

Sub CreateClassData()
    With Class(1)    'Knight
        .name = "Knight"
        .StartHP = 25
        .StartEnergy = 20
        .StartMana = 10
    End With
    With Class(2)    'Mage
        .name = "Mage"
        .StartHP = 25
        .StartEnergy = 20
        .StartMana = 25
    End With
    With Class(3)    'Rogue
        .name = "Rogue"
        .StartHP = 25
        .StartEnergy = 20
        .StartMana = 10
    End With
    With Class(4)    'Cleric
        .name = "Cleric"
        .StartHP = 25
        .StartEnergy = 20
        .StartMana = 20
    End With
End Sub

Function GetEquippedObjDur(ByVal Slot As Long) As Long
    Dim Percent As Single
    Dim Display As Boolean
    If Slot <= 0 Then Exit Function
    Slot = Slot - 20
    If Character.EquippedObject(Slot).value = 0 Then Exit Function
    If Character.EquippedObject(Slot).Object = 0 Then Exit Function
    Select Case Object(Character.EquippedObject(Slot).Object).Type
    Case 1, 2, 3, 4
        Percent = Character.EquippedObject(Slot).value / (Object(Character.EquippedObject(Slot).Object).MaxDur * 10)
        Percent = Int(Percent * 100)
        If Percent > 100 Then Percent = 100
        Display = True
    Case 8    'Ring
        Percent = Character.EquippedObject(Slot).value / (Object(Character.EquippedObject(Slot).Object).MaxDur * 10)
        Percent = Int(Percent * 100)
        If Percent > 100 Then Percent = 100
        Display = True
    Case Else
        Display = False
    End Select
    If Display = True Then
        GetEquippedObjDur = Percent
    Else
        GetEquippedObjDur = 0
    End If
End Function

Function DurString(Slot As Long) As String
    Dim B As Long
    If Slot > 20 Then
        If ExamineBit(Object(Character.EquippedObject(Slot - 20).Object).flags, 1) = 255 Then
            DurString = "Indestructible"
            Exit Function
        End If
        B = GetEquippedObjDur(Slot)
    Else
        If ExamineBit(Object(Character.Inv(Slot).Object).flags, 1) = 255 Then
            DurString = "Indestructible"
            Exit Function
        End If
        B = GetObjectDur(Slot)
    End If
    Select Case B
    Case Is >= 88
        DurString = "Excellent (" + CStr(B) + "%)"
    Case Is >= 65
        DurString = "Fair (" + CStr(B) + "%)"
    Case Is >= 45
        DurString = "Average (" + CStr(B) + "%)"
    Case Is >= 20
        DurString = "Seen Better Days (" + CStr(B) + "%)"
    Case Is >= 0
        DurString = "About to Break (" + CStr(B) + "%)"
    End Select
End Function

Sub DrawInfoText(hdcBuffer As Long)
    Dim A As Long
    If InfoTextTimer > 0 Then
        If Tick - InfoTextTimer <= 2000 Then
            For A = 0 To 1
                If InfoText(A) <> vbNullString Then
                    SetTextColor hdcBuffer, QBColor(0)
                    TextOut hdcBuffer, 5, 354 + 12 * A, InfoText(A), Len(InfoText(A))
                    SetTextColor hdcBuffer, QBColor(12)
                    TextOut hdcBuffer, 3, 352 + 12 * A, InfoText(A), Len(InfoText(A))
                End If
            Next A
        Else
            For A = 0 To 1
                InfoText(A) = vbNullString
            Next A
            InfoTextTimer = 0
        End If
    End If
End Sub

Sub MoveToTile()
    Character.LastMove = Tick
    SendSocket Chr$(7) + Chr$(CX) + Chr$(CY) + Chr$(CDir) + Chr$(CWalkStep)

    If CWalkStep = 4 Then
        If GetEnergy > 0 Then SetEnergy GetEnergy - 1
        DrawStats
    End If
    If Map.Tile(CX, CY).Att = 2 Then
        Freeze = True
        NextTransition = 5
    End If
End Sub

Sub PrintInfoText(St As String)
    Dim A As Long
    InfoTextTimer = Tick
    InfoText(0) = InfoText(1)
    InfoText(1) = St
End Sub

Sub ShowMap()
    SetLocation "[" + Map.name + "]"
    If ExamineBit(Map.flags, 0) = True Then
        frmMain.lblLocation.ForeColor = QBColor(11)
    ElseIf ExamineBit(Map.flags, 6) = True Then
        frmMain.lblLocation.ForeColor = QBColor(12)
    Else
        frmMain.lblLocation.ForeColor = QBColor(15)
    End If

    RedrawMap = True

    If frmWait_Loaded = True Then
        Unload frmWait
    End If
    If Map.MIDI > 0 Then
        PlayMidi CLng(Map.MIDI)
    Else
        StopMidi
    End If
    Transition

    frmMain.DrawInterfaceLights

    frmMain.picBuy.Visible = False
    frmMain.picSellObject.Visible = False
    frmMain.picRepair.Visible = False

    Freeze = False
End Sub

Sub ClearBit(bytByte As Byte, Bit As Byte)
    bytByte = bytByte And Not (2 ^ Bit)
End Sub

Sub RedrawTile()
    BitBlt frmMain.picTile.hDC, 0, 0, 32, 32, 0, 0, 0, BLACKNESS

    If EditMode < 6 Then
        If CurTile > 0 Then
            DrawToDC 0, 0, 32, 32, frmMain.picTile.hDC, DDSTiles, ((CurTile - 1) Mod 7) * 32, Int((CurTile - 1) / 7) * 32
        End If
    Else
        If CurAtt > 0 Then
            DrawToDC 0, 0, 32, 32, frmMain.picTile.hDC, DDSAtts, ((CurAtt - 1) Mod 7) * 32, Int((CurAtt - 1) / 7) * 32
        End If
    End If
    frmMain.picTile.Refresh
End Sub
Sub RedrawTiles()
    BitBlt frmMain.picTiles.hDC, 0, 0, 224, 192, 0, 0, 0, BLACKNESS

    If EditMode < 6 Then
        DrawToDC 0, 0, 224, 192, frmMain.picTiles.hDC, DDSTiles, 0, CInt(TopY)
    Else
        DrawToDC 0, 0, 224, 192, frmMain.picTiles.hDC, DDSAtts, 0, 0
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
    frmMain.picMapEdit.Visible = False
    RedrawMap = True
End Sub

Sub DrawChatString(hdcBuffer As Long)
    Dim r As RECT

    If ChatString <> vbNullString Then
        With r
            .Left = 7
            .Top = 7
            .Right = 381
            .Bottom = 52
        End With
        SetTextColor hdcBuffer, RGB(10, 10, 10)
        DrawText hdcBuffer, ChatString, Len(ChatString), r, DT_CENTER Or DT_NOCLIP Or DT_WORDBREAK
        With r
            .Left = 5
            .Top = 5
            .Right = 379
            .Bottom = 50
        End With
        SetTextColor hdcBuffer, QBColor(15)
        DrawText hdcBuffer, ChatString, Len(ChatString), r, DT_CENTER Or DT_NOCLIP Or DT_WORDBREAK

    End If
End Sub

Sub GetSections(ByVal St As String, NumSections)
    Dim A As Integer, W As Integer, Q As Boolean
    Dim CurChar As String * 1, LastChar As String * 1
    Erase Section
    Suffix = vbNullString
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
        Suffix = vbNullString
    End If
End Sub
Sub GetSections2(St)
    Dim A As Long, B As Long, C As Long
    B = 1
    Erase Section
    For A = 1 To 30
        C = InStr(B, St, vbNullChar)
        If C - B = 0 Then
            Section(A) = vbNullString
        ElseIf C <> 0 Then
            Section(A) = Mid$(St, B, C - B)
        Else
            Section(A) = Mid$(St, B, Len(St) - B + 1)
            Exit For
        End If
        B = C + 1
    Next A
End Sub
Sub GetSections3(St, Deliminator As String)
    Dim A As Integer, B As Integer, C As Integer
    B = 1
    Erase Section
    For A = 1 To 10
TryAgain:
        C = InStr(B, St, Deliminator)
        If C - B = 0 Then B = B + 1: GoTo TryAgain
        If C <> 0 Then
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
        If Mid$(St, A, 1) <> Chr$(32) And Mid$(St, A, 1) <> Chr$(0) Then
            ClipString = Mid$(St, 1, A)
            Exit Function
        End If
    Next A
End Function

Sub CopyMap(DestMap As MapData, SourceMap As MapData)
    Dim A As Long, X As Long, Y As Long

    With DestMap
        .name = SourceMap.name
        .MIDI = SourceMap.MIDI
        .NPC = SourceMap.NPC
        .ExitUp = SourceMap.ExitUp
        .ExitDown = SourceMap.ExitDown
        .ExitLeft = SourceMap.ExitLeft
        .ExitRight = SourceMap.ExitRight
        .BootLocation.Map = SourceMap.BootLocation.Map
        .BootLocation.X = SourceMap.BootLocation.X
        .BootLocation.Y = SourceMap.BootLocation.Y
        .DeathLocation.Map = SourceMap.DeathLocation.Map
        .DeathLocation.X = SourceMap.DeathLocation.X
        .DeathLocation.Y = SourceMap.DeathLocation.Y
        .flags = SourceMap.flags
        .Flags2 = SourceMap.Flags2
        For A = 0 To 9
            .MonsterSpawn(A).Monster = SourceMap.MonsterSpawn(A).Monster
            .MonsterSpawn(A).Rate = SourceMap.MonsterSpawn(A).Rate
        Next A
        For Y = 0 To 11
            For X = 0 To 11
                With .Tile(X, Y)
                    .Ground = SourceMap.Tile(X, Y).Ground
                    .Ground2 = SourceMap.Tile(X, Y).Ground2
                    .BGTile1 = SourceMap.Tile(X, Y).BGTile1
                    .BGTile2 = SourceMap.Tile(X, Y).BGTile2
                    .FGTile = SourceMap.Tile(X, Y).FGTile
                    .FGTile2 = SourceMap.Tile(X, Y).FGTile2
                    .Att = SourceMap.Tile(X, Y).Att
                    .AttData(0) = SourceMap.Tile(X, Y).AttData(0)
                    .AttData(1) = SourceMap.Tile(X, Y).AttData(1)
                    .AttData(2) = SourceMap.Tile(X, Y).AttData(2)
                    .AttData(3) = SourceMap.Tile(X, Y).AttData(3)
                    .Att2 = SourceMap.Tile(X, Y).Att2
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

Function IsVacant(X As Byte, Y As Byte) As Boolean
    If Character.Access > 0 And keyAlt = False Then
        IsVacant = True
        Exit Function
    End If
    Dim A As Long
    Select Case Map.Tile(X, Y).Att
    Case 1, 3, 13, 14, 15, 16    'Wall / Key Door
        Exit Function
    Case 2    'Warp
        If Tick > SwitchMapTimer Or Character.Access > 0 Then
            SwitchMapTimer = Tick + GetMapSwitchTime
        Else
            Exit Function
        End If
    Case 19    'Light
        If ExamineBit(Map.Tile(X, Y).AttData(2), 0) Then
            Exit Function
        End If
    Case 20    'Light Dampening
        If ExamineBit(Map.Tile(X, Y).AttData(3), 0) Then
            Exit Function
        End If
    End Select
    Select Case Map.Tile(X, Y).Att2
    Case 1, 3, 13, 14, 15, 16    'Wall / Key Door
        Exit Function
    End Select
    For A = 0 To MaxMonsters
        With Map.Monster(A)
            If .Monster > 0 And .X = X And .Y = Y Then
                Exit Function
            End If
        End With
    Next A
    For A = 1 To MaxUsers
        With Player(A)
            If .Map = CMap Then
                If .X = X Then
                    If .Y = Y Then
                        If Not .status = 25 Then
                            If .IsDead = False Then
                                If Character.Guild > 0 Then
                                    If Player(A).Guild = 0 Then
                                        If ExamineBit(Map.flags, 0) = False And ExamineBit(Map.flags, 6) = False Then

                                        Else
                                            Exit Function
                                        End If
                                    Else
                                        Exit Function
                                    End If
                                Else
                                    Exit Function
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End With
    Next A
    IsVacant = True
End Function
Sub OpenMapEdit()
    Dim Width As Long, Height As Long


    Dim File As String
    Dim FileByteArray() As Byte

    File = "tiles.rsc"
    FileByteArray() = StrConv(File, vbFromUnicode)
    ReDim Preserve FileByteArray(UBound(FileByteArray) + 1)

    EncryptDataFile FileByteArray(0), FileLen(File) Mod 87 + 5
    GetBitmapDimensions "tiles.rsc", Width, Height
    EncryptDataFile FileByteArray(0), FileLen(File) Mod 87 + 5



    frmMain.MapScroll.max = Int(Height / 32) - 7

    MapEdit = True
    CopyMap EditMap, Map
    frmMain.picMapEdit.Visible = True
    RedrawTiles
    RedrawTile
    RedrawMap = True
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

Function ReadUniqID() As String
    Dim lpReturnedString As String, Valid As Long
    lpReturnedString = Space$(256)
    Valid = GetPrivateProfileString&("UniqID", "ID", "", lpReturnedString, 256, "froogle")
    ReadUniqID = Left$(lpReturnedString, Valid)
End Function

Function WriteUniqID(UniqID As String) As String
    WritePrivateProfileString "UniqID", "ID", UniqID, "froogle"
End Function

Sub RedrawMapTile(X As Byte, Y As Byte)
    If frmMain.Visible = False Then Exit Sub

    Dim TileSource As RECT
    Dim A As Long, B As Long
    If X >= 0 And X <= 11 And Y >= 0 And Y <= 11 Then
        TileSource.Left = X * 32
        TileSource.Top = Y * 32
        TileSource.Right = TileSource.Left + 32
        TileSource.Bottom = TileSource.Top + 32
        Call BGTile1Buffer.BltColorFill(TileSource, RGB(0, 0, 0))
        Call BGTile2Buffer.BltColorFill(TileSource, RGB(0, 0, 0))
        Call FGTileBuffer.BltColorFill(TileSource, RGB(0, 0, 0))
        Call FGTile2Buffer.BltColorFill(TileSource, RGB(0, 0, 0))
        If MapEdit = False Then
            With Map.Tile(X, Y)
                If .Ground > 0 Then
                    TileSource.Left = ((.Ground - 1) Mod 7) * 32
                    TileSource.Top = Int((.Ground - 1) / 7) * 32
                    TileSource.Right = TileSource.Left + 32
                    TileSource.Bottom = TileSource.Top + 32
                    Call BGTile1Buffer.BltFast(X * 32, Y * 32, DDSTiles, TileSource, DDBLTFAST_SRCCOLORKEY)
                    Call BGTile2Buffer.BltFast(X * 32, Y * 32, DDSTiles, TileSource, DDBLTFAST_SRCCOLORKEY)
                End If
                If .Ground2 > 0 Then
                    TileSource.Left = ((.Ground2 - 1) Mod 7) * 32
                    TileSource.Top = Int((.Ground2 - 1) / 7) * 32
                    TileSource.Right = TileSource.Left + 32
                    TileSource.Bottom = TileSource.Top + 32
                    Call BGTile1Buffer.BltFast(X * 32, Y * 32, DDSTiles, TileSource, DDBLTFAST_SRCCOLORKEY)
                    Call BGTile2Buffer.BltFast(X * 32, Y * 32, DDSTiles, TileSource, DDBLTFAST_SRCCOLORKEY)
                End If
                If .BGTile1 > 0 Then
                    TileSource.Left = ((.BGTile1 - 1) Mod 7) * 32
                    TileSource.Top = Int((.BGTile1 - 1) / 7) * 32
                    TileSource.Right = TileSource.Left + 32
                    TileSource.Bottom = TileSource.Top + 32
                    Call BGTile1Buffer.BltFast(X * 32, Y * 32, DDSTiles, TileSource, DDBLTFAST_SRCCOLORKEY)
                End If
                If .BGTile2 > 0 Then
                    TileSource.Left = ((.BGTile2 - 1) Mod 7) * 32
                    TileSource.Top = Int((.BGTile2 - 1) / 7) * 32
                    TileSource.Right = TileSource.Left + 32
                    TileSource.Bottom = TileSource.Top + 32
                    Call BGTile2Buffer.BltFast(X * 32, Y * 32, DDSTiles, TileSource, DDBLTFAST_SRCCOLORKEY)
                ElseIf .BGTile1 > 0 Then
                    TileSource.Left = ((.BGTile1 - 1) Mod 7) * 32
                    TileSource.Top = Int((.BGTile1 - 1) / 7) * 32
                    TileSource.Right = TileSource.Left + 32
                    TileSource.Bottom = TileSource.Top + 32
                    Call BGTile2Buffer.BltFast(X * 32, Y * 32, DDSTiles, TileSource, DDBLTFAST_SRCCOLORKEY)
                End If
                If .FGTile > 0 Then
                    TileSource.Left = ((.FGTile - 1) Mod 7) * 32
                    TileSource.Top = Int((.FGTile - 1) / 7) * 32
                    TileSource.Right = TileSource.Left + 32
                    TileSource.Bottom = TileSource.Top + 32
                    Call FGTileBuffer.BltFast(X * 32, Y * 32, DDSTiles, TileSource, DDBLTFAST_SRCCOLORKEY)
                End If
                If .FGTile2 > 0 Then
                    TileSource.Left = ((.FGTile2 - 1) Mod 7) * 32
                    TileSource.Top = Int((.FGTile2 - 1) / 7) * 32
                    TileSource.Right = TileSource.Left + 32
                    TileSource.Bottom = TileSource.Top + 32
                    Call FGTile2Buffer.BltFast(X * 32, Y * 32, DDSTiles, TileSource, DDBLTFAST_SRCCOLORKEY)
                ElseIf .FGTile > 0 Then
                    TileSource.Left = ((.FGTile - 1) Mod 7) * 32
                    TileSource.Top = Int((.FGTile - 1) / 7) * 32
                    TileSource.Right = TileSource.Left + 32
                    TileSource.Bottom = TileSource.Top + 32
                    Call FGTile2Buffer.BltFast(X * 32, Y * 32, DDSTiles, TileSource, DDBLTFAST_SRCCOLORKEY)
                End If
            End With
        Else
            With EditMap.Tile(X, Y)
                If .Ground > 0 Then
                    TileSource.Left = ((.Ground - 1) Mod 7) * 32
                    TileSource.Top = Int((.Ground - 1) / 7) * 32
                    TileSource.Right = TileSource.Left + 32
                    TileSource.Bottom = TileSource.Top + 32
                    Call BGTile1Buffer.BltFast(X * 32, Y * 32, DDSTiles, TileSource, DDBLTFAST_SRCCOLORKEY)
                    Call BGTile2Buffer.BltFast(X * 32, Y * 32, DDSTiles, TileSource, DDBLTFAST_SRCCOLORKEY)
                End If
                If .Ground2 > 0 Then
                    TileSource.Left = ((.Ground2 - 1) Mod 7) * 32
                    TileSource.Top = Int((.Ground2 - 1) / 7) * 32
                    TileSource.Right = TileSource.Left + 32
                    TileSource.Bottom = TileSource.Top + 32
                    Call BGTile1Buffer.BltFast(X * 32, Y * 32, DDSTiles, TileSource, DDBLTFAST_SRCCOLORKEY)
                    Call BGTile2Buffer.BltFast(X * 32, Y * 32, DDSTiles, TileSource, DDBLTFAST_SRCCOLORKEY)
                End If
                If .BGTile1 > 0 Then
                    TileSource.Left = ((.BGTile1 - 1) Mod 7) * 32
                    TileSource.Top = Int((.BGTile1 - 1) / 7) * 32
                    TileSource.Right = TileSource.Left + 32
                    TileSource.Bottom = TileSource.Top + 32
                    Call BGTile1Buffer.BltFast(X * 32, Y * 32, DDSTiles, TileSource, DDBLTFAST_SRCCOLORKEY)
                End If
                If .BGTile2 > 0 Then
                    TileSource.Left = ((.BGTile2 - 1) Mod 7) * 32
                    TileSource.Top = Int((.BGTile2 - 1) / 7) * 32
                    TileSource.Right = TileSource.Left + 32
                    TileSource.Bottom = TileSource.Top + 32
                    Call BGTile2Buffer.BltFast(X * 32, Y * 32, DDSTiles, TileSource, DDBLTFAST_SRCCOLORKEY)
                ElseIf .BGTile1 > 0 Then
                    TileSource.Left = ((.BGTile1 - 1) Mod 7) * 32
                    TileSource.Top = Int((.BGTile1 - 1) / 7) * 32
                    TileSource.Right = TileSource.Left + 32
                    TileSource.Bottom = TileSource.Top + 32
                    Call BGTile2Buffer.BltFast(X * 32, Y * 32, DDSTiles, TileSource, DDBLTFAST_SRCCOLORKEY)
                End If
                If .FGTile > 0 Then
                    TileSource.Left = ((.FGTile - 1) Mod 7) * 32
                    TileSource.Top = Int((.FGTile - 1) / 7) * 32
                    TileSource.Right = TileSource.Left + 32
                    TileSource.Bottom = TileSource.Top + 32
                    Call FGTileBuffer.BltFast(X * 32, Y * 32, DDSTiles, TileSource, DDBLTFAST_SRCCOLORKEY)
                End If
                If .FGTile2 > 0 Then
                    TileSource.Left = ((.FGTile2 - 1) Mod 7) * 32
                    TileSource.Top = Int((.FGTile2 - 1) / 7) * 32
                    TileSource.Right = TileSource.Left + 32
                    TileSource.Bottom = TileSource.Top + 32
                    Call FGTile2Buffer.BltFast(X * 32, Y * 32, DDSTiles, TileSource, DDBLTFAST_SRCCOLORKEY)
                ElseIf .FGTile > 0 Then
                    TileSource.Left = ((.FGTile - 1) Mod 7) * 32
                    TileSource.Top = Int((.FGTile - 1) / 7) * 32
                    TileSource.Right = TileSource.Left + 32
                    TileSource.Bottom = TileSource.Top + 32
                    Call FGTile2Buffer.BltFast(X * 32, Y * 32, DDSTiles, TileSource, DDBLTFAST_SRCCOLORKEY)
                End If
                If EditMode >= 6 Then
                    If .Att2 > 0 Then
                        TileSource.Left = ((.Att2 - 1) Mod 7) * 32 + 8
                        TileSource.Top = Int((.Att2 - 1) / 7) * 32 + 8
                        TileSource.Right = TileSource.Left + 16
                        TileSource.Bottom = TileSource.Top + 16
                        Call FGTileBuffer.BltFast(X * 32 + 12, Y * 32 + 12, DDSAtts, TileSource, DDBLTFAST_SRCCOLORKEY)
                        Call FGTile2Buffer.BltFast(X * 32 + 12, Y * 32 + 12, DDSAtts, TileSource, DDBLTFAST_SRCCOLORKEY)
                    End If
                    If .Att > 0 Then
                        TileSource.Left = ((.Att - 1) Mod 7) * 32 + 8
                        TileSource.Top = Int((.Att - 1) / 7) * 32 + 8
                        TileSource.Right = TileSource.Left + 16
                        TileSource.Bottom = TileSource.Top + 16
                        Call FGTileBuffer.BltFast(X * 32 + 4, Y * 32 + 4, DDSAtts, TileSource, DDBLTFAST_SRCCOLORKEY)
                        Call FGTile2Buffer.BltFast(X * 32 + 4, Y * 32 + 4, DDSAtts, TileSource, DDBLTFAST_SRCCOLORKEY)
                    End If
                End If
            End With
        End If
        For A = 0 To MaxMapObjects
            With Map.Object(A)
                If .Object > 0 And .X = X And .Y = Y Then
                    B = Object(.Object).Picture
                    If B > 0 Then
                        TileSource.Left = 0
                        TileSource.Top = (B - 1) * 32
                        TileSource.Right = TileSource.Left + 32
                        TileSource.Bottom = TileSource.Top + 32
                        Call BGTile1Buffer.BltFast(.X * 32, .Y * 32, DDSObjects, TileSource, DDBLTFAST_SRCCOLORKEY)
                        Call BGTile2Buffer.BltFast(.X * 32, .Y * 32, DDSObjects, TileSource, DDBLTFAST_SRCCOLORKEY)
                    End If
                End If
            End With
        Next A
    End If
End Sub
Sub Transition()
    Select Case NextTransition
    Case 0    'None
        PlayWav 6
    Case 1    'Player Moved Up
        PlayWav 1
    Case 2    'Player Moved Down
        PlayWav 1
    Case 3    'Player Moved Left
        PlayWav 1
    Case 4    'Player Moved Right
        PlayWav 1
    Case 5    'Warp
        PlayWav 6
    Case 6    'Death

    Case Else
        PlayWav 6
    End Select
    NextTransition = 0
End Sub
Sub UpdatePlayerColor(index As Long)
    Dim A As Long
    With Player(index)
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
    Dim A As Long
    For A = 1 To MaxUsers
        UpdatePlayerColor A
    Next A
End Sub
Sub UpdateSaleItem(NPCIndex As Long, A As Long)
    Dim St As String
    With NPC(NPCIndex).SaleItem(A)
        If .GiveObject > MaxObjects Then
            .GiveObject = 0
            SaveNPC NPCIndex
        End If
        If .TakeObject > MaxObjects Then
            .TakeObject = 0
            SaveNPC NPCIndex
        End If
        If .GiveObject >= 1 And .TakeObject >= 1 And .GiveObject <= MaxObjects And .TakeObject <= MaxObjects Then
            St = CStr(A) + ": "
            If Object(.GiveObject).Type = 6 Then
                'Money
                St = St + CStr(.GiveValue) + " " + Object(.GiveObject).name
            Else
                St = St + "1 " + Object(.GiveObject).name
            End If
            St = St + " in exchange for "
            If Object(.TakeObject).Type = 6 Then
                'Money
                St = St + CStr(.TakeValue) + " " + Object(.TakeObject).name
            Else
                St = St + "1 " + Object(.TakeObject).name
            End If
            frmNPC.lstSaleItems.List(A) = St
        Else
            frmNPC.lstSaleItems.List(A) = CStr(A) + ":"
        End If
    End With
End Sub
Sub WriteString(lpAppName, lpKeyName As String, A, Optional FileString As String = "odyssey.ini")
    Dim lpString As String, Valid As Long
    lpString = A
    Valid = WritePrivateProfileString&(lpAppName, lpKeyName, lpString, App.Path + "\" + FileString)
End Sub

Sub CheckKeys()
    If Character.IsDead = True Then Exit Sub
    Dim A As Long

    If keyCtrl = True And Not Tick > Character.LastMove + 60000 And MapEdit = False Then
        If Tick >= AttackTimer Then
            If Character.Projectile = False Then
                AttackTimer = Tick + 1000
                If NoDirectionalWalls(CX, CY, CDir) Then
                    Dim tx As Long, ty As Long
                    Select Case CDir
                    Case 0
                        tx = CX
                        ty = CY - 1
                    Case 1
                        tx = CX
                        ty = CY + 1
                    Case 2
                        tx = CX - 1
                        ty = CY
                    Case 3
                        tx = CX + 1
                        ty = CY
                    End Select
                    If tx >= 0 And tx <= 11 And ty >= 0 And ty <= 11 Then
                        For A = 0 To MaxMonsters
                            With Map.Monster(A)
                                If .Monster > 0 And .X = tx And .Y = ty Then
                                    SendSocket Chr$(26) + Chr$(A)
                                    Exit For
                                End If
                            End With
                        Next A
                        If A = MaxMonsters + 1 Then
                            For A = 1 To MaxUsers
                                With Player(A)
                                    If .Map = CMap And .Sprite > 0 And .X = tx And .Y = ty And .IsDead = False And Not .status = 25 Then
                                        If Not Player(A).Guild = Character.Guild Or Character.Guild = 0 Then
                                            If Character.Guild > 0 Then
                                                If Player(A).Guild = 0 Then
                                                    If ExamineBit(Map.flags, 6) = False Then

                                                    Else
                                                        SendSocket Chr$(25) + Chr$(A)
                                                        Exit For
                                                    End If
                                                Else
                                                    SendSocket Chr$(25) + Chr$(A)
                                                    Exit For
                                                End If
                                            Else
                                                SendSocket Chr$(25) + Chr$(A)
                                                Exit For
                                            End If
                                        Else
                                            PrintChat "You cannot attack members of your own guild!", 7
                                        End If
                                    End If
                                End With
                            Next A
                            If A = 81 Then
                                PrintInfoText "There is nothing there to attack!"
                            End If
                        End If
                    Else
                        PrintInfoText "There is nothing there to attack!"
                    End If
                Else
                    PrintInfoText "You cannot attack here!"
                End If
            Else
                AttackTimer = Tick + 1000
                If Not Object(Character.EquippedObject(1).Object).Type = 10 Then Character.Projectile = False
                If Character.Ammo > 0 Then
                    If Character.Inv(Character.Ammo).value > 0 Then
                        Character.Inv(Character.Ammo).value = Character.Inv(Character.Ammo).value - 1
                        If CurInvObj = Character.Ammo Then RefreshInventory
                        SendSocket Chr$(72)
                    Else
                        PrintInfoText "Out of Ammunition!"
                    End If
                Else
                    PrintInfoText "Out of Ammunition!"
                End If
            End If
        End If
    End If
    If CX * 32 = CXO And CY * 32 = CYO Then
        If Character.Access = 0 Or keyAlt = True Then
            If CWalkStep > 4 And Character.Access = 0 Then
                SendSocket Chr$(68) + "Walk Hack"
                Exit Sub
            End If
            If keyShift = True And GetEnergy > 0 Then
                CWalkStep = 4
            Else
                CWalkStep = 2
            End If
        Else
            CWalkStep = 16
        End If
        If keyUp = True Then
            If CDir = 0 Then
                If CY > 0 Then
                    If IsVacant(CX, CY - 1) Then
                        If NoDirectionalWalls(CX, CY, 0) Then
                            CY = CY - 1
                            MoveToTile
                        End If
                    End If
                Else
                    If Map.ExitUp > 0 Then
                        If Tick > SwitchMapTimer Or Character.Access > 0 Then
                            SwitchMapTimer = Tick + GetMapSwitchTime
                            SendSocket Chr$(13) + Chr$(0)
                            Freeze = True
                            NextTransition = 1
                        End If
                    End If
                End If
            Else
                CDir = 0
                SendSocket Chr$(7) + Chr$(CX) + Chr$(CY) + Chr$(CDir) + Chr$(CWalkStep)
            End If
        ElseIf keyDown = True Then
            If CDir = 1 Then
                If CY < 11 Then
                    If IsVacant(CX, CY + 1) Then
                        If NoDirectionalWalls(CX, CY, 1) Then
                            CY = CY + 1
                            MoveToTile
                        End If
                    End If
                Else
                    If Map.ExitDown > 0 Then
                        If Tick > SwitchMapTimer Or Character.Access > 0 Then
                            SwitchMapTimer = Tick + GetMapSwitchTime
                            SendSocket Chr$(13) + Chr$(1)
                            Freeze = True
                            NextTransition = 2
                        End If
                    End If
                End If
            Else
                CDir = 1
                SendSocket Chr$(7) + Chr$(CX) + Chr$(CY) + Chr$(CDir) + Chr$(CWalkStep)
            End If
        ElseIf keyLeft = True Then
            If CDir = 2 Then
                If CX > 0 Then
                    If IsVacant(CX - 1, CY) Then
                        If NoDirectionalWalls(CX, CY, 2) Then
                            CX = CX - 1
                            MoveToTile
                        End If
                    End If
                Else
                    If Map.ExitLeft > 0 Then
                        If Tick > SwitchMapTimer Or Character.Access > 0 Then
                            SwitchMapTimer = Tick + GetMapSwitchTime
                            SendSocket Chr$(13) + Chr$(2)
                            Freeze = True
                            NextTransition = 3
                        End If
                    End If
                End If
            Else
                CDir = 2
                SendSocket Chr$(7) + Chr$(CX) + Chr$(CY) + Chr$(CDir) + Chr$(CWalkStep)
            End If
        ElseIf keyRight = True Then
            If CDir = 3 Then
                If CX < 11 Then
                    If IsVacant(CX + 1, CY) Then
                        If NoDirectionalWalls(CX, CY, 3) Then
                            CX = CX + 1
                            MoveToTile
                        End If
                    End If
                Else
                    If Map.ExitRight > 0 Then
                        If Tick > SwitchMapTimer Or Character.Access > 0 Then
                            SwitchMapTimer = Tick + GetMapSwitchTime
                            SendSocket Chr$(13) + Chr$(3)
                            Freeze = True
                            NextTransition = 4
                        End If
                    End If
                End If
            Else
                CDir = 3
                SendSocket Chr$(7) + Chr$(CX) + Chr$(CY) + Chr$(CDir) + Chr$(CWalkStep)
            End If
        End If
    End If
End Sub
Sub DrawNextFrame()
    Dim A As Long, B As Long, C As Long, D As Long, E As Double, r As RECT
    Dim TempVar As Byte, TempStr As String

    If RestoreDirectDraw = True Then
        On Error Resume Next
        UnloadDirectDraw
        InitDirectDraw
        LoadSurfaces
        RedrawMap = True
        On Error GoTo 0
        RestoreDirectDraw = False
    End If

    If RedrawMap = True Then
        DrawMap
        RedrawMap = False
    End If

    'Copy Back Buffer to Viewport
    If CurFrame = 0 Then
        BackBufferSurf.BltFast 0, 0, BGTile1Buffer, FullMapRect, DDBLTFAST_WAIT
    Else
        BackBufferSurf.BltFast 0, 0, BGTile2Buffer, FullMapRect, DDBLTFAST_WAIT
    End If

    For A = 1 To MaxUsers
        With Player(A)
            If .Map = CMap Then
                If Not .status = 25 Then
                    If .Sprite <= MaxSprite Then
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

                        If Player(A).IsDead Then
                            Draw .XO, .YO, 32, 32, DDSTiles, (623 Mod 7) * 32, (623 / 7) * 32, True
                        Else
                            Draw .XO, .YO - 16, 32, 32, DDSSprites, B * 32, (.Sprite - 1) * 32, True
                            If Player(A).HP > 0 Then
                                If Not Player(A).HP = Player(A).MaxHP Then
                                    Draw .XO + 3, .YO - 16, 2, 26, DDSHPBar, 0, 4, False
                                    Draw .XO + 3, .YO - 16, 2, 26 - (Player(A).HP / Player(A).MaxHP) * 26, DDSHPBar, 2, 4, False
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End With
    Next A

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

    'Draw You
    If CAttack > 0 Then
        B = CDir * 3 + 2
        CAttack = CAttack - 1
    Else
        B = CDir * 3 + CWalk
    End If

    If Character.IsDead = True Then
        Draw CXO, CYO, 32, 32, DDSTiles, (623 Mod 7) * 32, (623 / 7) * 32, True
    Else
        Draw CXO, CYO - 16, 32, 32, DDSSprites, B * 32, (Character.Sprite - 1) * 32, True
    End If

    For A = 0 To MaxMonsters
        With Map.Monster(A)
            If .Monster > 0 Then
                C = Monster(.Monster).Sprite
                If C > 0 And C <= MaxSprite Then
                    If .XO < .X * 32 Then
                        If ExamineBit(Monster(.Monster).flags, 2) = False Then 'Not runner
                            .XO = .XO + 2
                        Else
                            .XO = .XO + 4
                        End If
                        If Int(.XO / 16) * 16 = .XO Then .W = 1 - .W
                    ElseIf .XO > .X * 32 Then
                        If ExamineBit(Monster(.Monster).flags, 2) = False Then 'Not runner
                            .XO = .XO - 2
                        Else
                            .XO = .XO - 4
                        End If
                        If Int(.XO / 16) * 16 = .XO Then .W = 1 - .W
                    End If
                    If .YO < .Y * 32 Then
                        If ExamineBit(Monster(.Monster).flags, 2) = False Then 'Not runner
                            .YO = .YO + 2
                        Else
                            .YO = .YO + 4
                        End If
                        If Int(.YO / 16) * 16 = .YO Then .W = 1 - .W
                    ElseIf .YO > .Y * 32 Then
                        If ExamineBit(Monster(.Monster).flags, 2) = False Then 'Not runner
                            .YO = .YO - 2
                        Else
                            .YO = .YO - 4
                        End If
                        If Int(.YO / 16) * 16 = .YO Then .W = 1 - .W
                    End If

                    'Draw Monster
                    If .A > 0 Then
                        B = .D * 3 + 2
                        .A = .A - 1
                    Else
                        B = .D * 3 + .W
                    End If
                    Draw .XO, .YO - 16, 32, 32, DDSSprites, B * 32, (C - 1) * 32, True
                    
                    If .HPBar = True Or Character.Access > 0 Then
                        Draw .XO + 3, .YO - 20, 26, 2, DDSHPBar, 0, 0, False
                        E = (.Life / Monster(.Monster).MaxLife)
                        If E > 1 Then E = 1
                        Draw .XO + 3, .YO - 20, E * 26, 2, DDSHPBar, 0, 2, False
                    End If
                End If
            End If
        End With
    Next A

    For A = 1 To MaxProjectiles
        With Projectile(A)
            If .Sprite > 0 Then
                Select Case .TargetType
                Case pttCharacter
                    If .TargetNum = Character.index Then
                        .X = CXO
                        .Y = CYO
                    Else
                        .X = Player(.TargetNum).XO
                        .Y = Player(.TargetNum).YO
                    End If

                    If Tick - .TimeStamp >= .speed Then
                        If .Frame < .TotalFrames Then
                            .Frame = .Frame + 1
                        Else
                            If .CurLoop = .LoopCount Then
                                If .EndSound > 0 Then
                                    PlayWav .EndSound
                                End If
                                DestroyEffect A
                            Else
                                .CurLoop = .CurLoop + 1
                                .Frame = 0
                            End If
                        End If
                        .TimeStamp = Tick
                    End If
                Case pttPlayer
                    If .TargetNum = Character.index Then
                        .TargetX = CXO
                        .TargetY = CYO
                    Else
                        .TargetX = Player(.TargetNum).XO
                        .TargetY = Player(.TargetNum).YO
                    End If
                    If .X < .TargetX Then .X = .X + 8
                    If .X > .TargetX Then .X = .X - 8
                    If .Y < .TargetY Then .Y = .Y + 8
                    If .Y > .TargetY Then .Y = .Y - 8

                    If Tick - .TimeStamp >= .speed Then
                        If .X = .TargetX Then
                            If .Y = .TargetY Then
                                If .Frame < .TotalFrames Then
                                    .Frame = .Frame + 1
                                Else
                                    DestroyEffect A
                                End If
                            End If
                        End If
                        .TimeStamp = Tick
                    End If
                Case pttMonster
                    .TargetX = Map.Monster(.TargetNum).XO
                    .TargetY = Map.Monster(.TargetNum).YO
                    If .X < .TargetX Then .X = .X + 8
                    If .X > .TargetX Then .X = .X - 8
                    If .Y < .TargetY Then .Y = .Y + 8
                    If .Y > .TargetY Then .Y = .Y - 8

                    If Tick - .TimeStamp >= .speed Then
                        If .X = .TargetX Then
                            If .Y = .TargetY Then
                                If .Frame < .TotalFrames Then
                                    .Frame = .Frame + 1
                                Else
                                    If .EndSound > 0 Then PlayWav .EndSound
                                    DestroyEffect A
                                End If
                            End If
                        End If
                        .TimeStamp = Tick
                    End If
                Case pttTile
                    If Tick - .TimeStamp >= .speed Then
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
                        .TimeStamp = Tick
                    End If
                Case pttProject
                    If Tick - .TimeStamp >= .speed Then
                        If .X = .TargetX And .Y = .TargetY Then
                            If .TotalFrames > 0 Then
                                If .Frame < .TotalFrames Then
                                    .Frame = .Frame + 1
                                Else
                                    If .EndSound > 0 Then PlayWav .EndSound
                                    DestroyEffect A
                                End If
                            Else
                                If .EndSound > 0 Then PlayWav .EndSound
                                DestroyEffect A
                            End If
                        Else
                            If .X < .TargetX Then .X = .X + 8
                            If .X > .TargetX Then .X = .X - 8
                            If .Y < .TargetY Then .Y = .Y + 8
                            If .Y > .TargetY Then .Y = .Y - 8
                            If .Alternate = True Then
                                Select Case .Type
                                Case 2
                                    .offset = 1 - .offset
                                    .Frame = .offset
                                Case 4
                                    If .offset = 3 Then .offset = 0 Else .offset = .offset + 1
                                    .Frame = .offset
                                End Select
                            End If
                            C = (.X / 32)
                            D = (.Y / 32)
                            'Projectile Collision
                            Select Case Map.Tile(C, D).Att
                            Case 1, 2, 3, 14, 16
                                .TargetX = .X
                                .TargetY = .Y
                            Case 19    'Light
                                If ExamineBit(Map.Tile(C, D).AttData(2), 0) = 1 Then
                                    .TargetX = .X
                                    .TargetY = .Y
                                End If
                            Case 20    'Light Dampening
                                If ExamineBit(Map.Tile(C, D).AttData(3), 0) Then
                                    .TargetX = .X
                                    .TargetY = .Y
                                End If
                            End Select
                            Select Case Map.Tile(C, D).Att2
                            Case 1, 14, 16
                                .TargetX = .X
                                .TargetY = .Y
                            End Select
                            Dim Direction As Byte
                            If .X < .TargetX Then Direction = 3
                            If .X > .TargetX Then Direction = 2
                            If .Y < .TargetY Then Direction = 1
                            If .Y > .TargetY Then Direction = 0
                            If NoDirectionalWalls(CByte(.X / 32), CByte(.Y / 32), Direction) = False Then
                                .TargetX = .X
                                .TargetY = .Y
                            End If

                            For B = 0 To MaxMonsters
                                If Map.Monster(B).X = C Then
                                    If Map.Monster(B).Y = D Then
                                        If Map.Monster(B).Monster > 0 Then
                                            If .Creator = Character.index Then
                                                If .Damage > 0 Then
                                                    TempVar = (CMap + CX + CY) Mod 250
                                                    If .Magic > 0 Then
                                                        'Magic Projectile
                                                        TempStr = Chr$(TempVar) + Chr$(1) + Chr$(B) + Chr$(.Damage)
                                                        SendSocket Chr$(79) + Chr$(CheckSum(TempStr) Mod 256) + TempStr
                                                    Else
                                                        'Normal Projectile
                                                        TempStr = Chr$(TempVar) + Chr$(2) + Chr$(B) + Chr$(.Damage)
                                                        SendSocket Chr$(79) + Chr$(CheckSum(TempStr) Mod 256) + TempStr
                                                    End If
                                                Else
                                                    SendSocket Chr$(73) & Chr$(B)
                                                End If
                                            End If

                                            .TargetX = .X
                                            .TargetY = .Y
                                        End If
                                    End If
                                End If
                            Next B
                            For B = 1 To MaxUsers
                                If Player(B).X = C Then
                                    If Player(B).Y = D Then
                                        If Player(B).Map = CMap Then
                                            If Not B = .Creator Then
                                                If Player(B).IsDead = False Then
                                                    Dim Collide As Boolean
                                                    If Character.Guild > 0 Then
                                                        If Player(B).Guild = 0 Then
                                                            If ExamineBit(Map.flags, 0) = False And ExamineBit(Map.flags, 6) = False Then

                                                            Else
                                                                Collide = True
                                                            End If
                                                        Else
                                                            Collide = True
                                                        End If
                                                    Else
                                                        Collide = True
                                                    End If
                                                    If Collide = True Then
                                                        .TargetX = .X
                                                        .TargetY = .Y
                                                        If .Creator = Character.index Then
                                                            If .Damage > 0 Then
                                                                TempVar = CMap Mod 250

                                                                If .Magic > 0 Then
                                                                    'Magic Projectile
                                                                    TempStr = Chr$(TempVar) + Chr$(3) + Chr$(B) + Chr$(.Damage)
                                                                    SendSocket Chr$(79) + Chr$(CheckSum(TempStr) Mod 256) + TempStr
                                                                    Exit For
                                                                Else
                                                                    'Normal Projectile
                                                                    TempStr = Chr$(TempVar) + Chr$(4) + Chr$(B) + Chr$(.Damage)
                                                                    SendSocket Chr$(79) + Chr$(CheckSum(TempStr) Mod 256) + TempStr
                                                                    Exit For
                                                                End If
                                                            Else
                                                                SendSocket Chr$(74) & Chr$(B)
                                                            End If
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            Next B
                            If CX = C Then
                                If CY = D Then
                                    If Not .Creator = Character.index Then
                                        .TargetX = .X
                                        .TargetY = .Y
                                    End If
                                End If
                            End If
                        End If
                        .TimeStamp = Tick
                    End If
                End Select
                Draw .X, .Y - 16, 32, 32, DDSEffects, .Frame * 32, (.Sprite - 1) * 32, True
            End If
        End With
    Next A

    If CurFrame = 0 Then
        Call BackBufferSurf.BltFast(0, 0, FGTileBuffer, FullMapRect, DDBLTFAST_SRCCOLORKEY)
    Else
        Call BackBufferSurf.BltFast(0, 0, FGTile2Buffer, FullMapRect, DDBLTFAST_SRCCOLORKEY)
    End If

    Dim hdcBuffer As Long
    hdcBuffer = BackBufferSurf.GetDC
    SetBkMode hdcBuffer, Transparent

    With r
        .Left = CXO - 32
        .Right = CXO + 64
        .Top = CYO - 32
        .Bottom = CYO - 16
    End With

    If Character.Guild > 0 Then
        If Character.status = 1 And CurFrame = 0 Then
            Draw3dText hdcBuffer, r, Character.name, QBColor(4), 2
        Else
            If Character.status > 1 Then
                Draw3dText hdcBuffer, r, Character.name, StatusColors(Character.status), 2
            Else
                Draw3dText hdcBuffer, r, Character.name, QBColor(11), 2
            End If
        End If
    Else
        If Character.status = 2 Then
            Draw3dText hdcBuffer, r, Character.name, QBColor(14), 2
        ElseIf Character.status = 3 Then
            Draw3dText hdcBuffer, r, Character.name, QBColor(9), 2
        ElseIf Character.status = 1 And CurFrame = 0 Then
            Draw3dText hdcBuffer, r, Character.name, QBColor(4), 2
        Else
            Draw3dText hdcBuffer, r, Character.name, StatusColors(Character.status), 2
        End If
    End If

    If Character.status = 24 Then    'Rainbow
        Draw3dText hdcBuffer, r, Character.name, StatusColors((Int(Rnd * 23))), 2
    End If

    For A = 1 To MaxUsers
        With Player(A)
            If .Map = CMap Then
                If .IsDead = False Then
                    If .status = 9 Or .status = 25 Then

                    Else
                        r.Left = .XO - 32
                        r.Right = .XO + 64
                        r.Top = .YO - 32
                        r.Bottom = .YO - 16

                        If .status = 1 And CurFrame = 0 Then
                            Draw3dText hdcBuffer, r, .name, QBColor(4), 2
                        ElseIf .status = 1 And CurFrame = 1 Then
                            Draw3dText hdcBuffer, r, .name, QBColor(.Color), 2
                        ElseIf .status = 0 Then
                            Draw3dText hdcBuffer, r, .name, QBColor(.Color), 2
                        Else
                            If .status = 24 Then  'Rainbow
                                Draw3dText hdcBuffer, r, .name, StatusColors((Int(Rnd * 23))), 2
                            Else
                                Draw3dText hdcBuffer, r, .name, StatusColors(.status), 2
                            End If
                        End If
                    End If
                End If
            End If
        End With
    Next A

    For A = 1 To MaxFloatText    'Floating Text
        With FloatText(A)
            If .InUse = True Then
                With r
                    .Left = FloatText(A).X * 32 - 32
                    .Right = FloatText(A).X * 32 + 64
                    .Top = FloatText(A).Y * 32 - 32 + FloatText(A).FloatY
                    .Bottom = FloatText(A).Y * 32 - 16
                End With
                Draw3dText hdcBuffer, r, .Text, QBColor(.Color), 2
                If .Static = False Then
                    .FloatY = .FloatY - 1
                    If .FloatY <= -38 Then ClearFloatText CByte(A)
                End If
            End If
        End With
    Next A

    BackBufferSurf.ReleaseDC hdcBuffer

    If options.DisableLighting = False Then
        UpdateLights

        If ShadeMapFrame = 0 Then
            UpdateLightMap Lighting(0)
            ShadeMapFrame = options.LightingQuality
        Else
            ShadeMapFrame = ShadeMapFrame - 1
        End If

        BackBufferSurf.Lock EmptyRect, DDSDBackBuffer, DDLOCK_NOSYSLOCK Or DDLOCK_WAIT, 0

        BackBufferSurf.GetLockedArray ddsBufferArray()

        If Indoors = False Then
            If ExamineBit(Map.Flags2, 0) = True Then
                If options.Bit32 = True Then
                    Rain32 ddsBufferArray(0, 0), Tick
                Else
                    Rain16 ddsBufferArray(0, 0), Tick
                End If
            End If
            If ExamineBit(Map.Flags2, 1) = True Then
                If options.Bit32 = True Then
                    Snow32 ddsBufferArray(0, 0), Tick
                Else
                    Snow16 ddsBufferArray(0, 0), Tick
                End If
            End If
        End If

        If options.Bit32 = True Then
            ShadeMap32 ddsBufferArray(0, 0)
        Else
            ShadeMap16 ddsBufferArray(0, 0)
        End If

        BackBufferSurf.Unlock EmptyRect
    End If
    
    hdcBuffer = BackBufferSurf.GetDC
    SetBkMode hdcBuffer, Transparent
    
    DrawChatString hdcBuffer
    DrawInfoText hdcBuffer

    BackBufferSurf.ReleaseDC hdcBuffer

    If frmMain_Showing = False Then
        frmMain.Show
        RefreshInventory
        frmMain_Showing = True
    End If

    Call DX7.GetWindowRect(frmMain.picViewport.hwnd, MapRect)

    LastDrawReturn = -1
    On Error Resume Next
    LastDrawReturn = PrimarySurf.Blt(MapRect, BackBufferSurf, FullMapRect, DDBLT_WAIT)
    On Error GoTo 0

    If (LastDrawReturn <> 0) Then

        On Error Resume Next
        RestoreSurfaces

        LastDrawReturn = -1
        LastDrawReturn = PrimarySurf.Blt(MapRect, BackBufferSurf, FullMapRect, DDBLT_WAIT)

        On Error GoTo 0

        If LastDrawReturn <> 0 Then
            RestoreDirectDraw = True
        End If
    End If

End Sub
Function Exists(filename As String) As Boolean
    Exists = (Dir(filename) <> vbNullString)
End Function
Sub CheckFile(filename As String)
    If Exists(filename) = False Then
        MsgBox "Error: File " + Chr$(34) + filename + Chr$(34) + " not found!", vbOKOnly + vbExclamation, TitleString
        End
    End If
End Sub

Sub CloseClientSocket(Action As Byte)

    ShutDown ClientSocket, 2
    closesocket ClientSocket
    ClientSocket = INVALID_SOCKET

    blnPlaying = False

    If frmWait_Loaded = True Then Unload frmWait
    If frmMain_Loaded = True Then Unload frmMain
    frmMain_Showing = False
    If frmLogin_Loaded = True And Action <> 1 Then Unload frmLogin
    If frmAccount_Loaded = True And Action <> 2 Then Unload frmAccount
    If frmNewCharacter_Loaded = True Then Unload frmNewCharacter
    If frmNewPass_Loaded = True Then Unload frmNewPass
    If frmEmail_Loaded = True Then Unload frmEmail
    If frmMonster_Loaded = True Then Unload frmMonster
    If frmObject_Loaded = True Then Unload frmObject
    If frmList_Loaded = True Then Unload frmList
    If frmMapProperties_Loaded = True Then Unload frmMapProperties
    If frmGuild_Loaded = True Then Unload frmGuild
    If frmNPC_Loaded = True Then Unload frmNPC
    If frmMacros_Loaded = True Then Unload frmMacros
    If frmOptions_Loaded = True Then Unload frmOptions
    If frmNewGuild_Loaded = True Then Unload frmNewGuild
    If frmBan_Loaded = True Then Unload frmBan
    If frmHall_Loaded = True Then Unload frmHall
    If frmMagic_Loaded = True Then Unload frmMagic
    If frmPrefix_Loaded = True Then Unload frmPrefix
    If frmSuffix_Loaded = True Then Unload frmSuffix

    Select Case Action
    Case 0
        frmMenu.Show
    Case 1
        frmMenu.Show
    Case 2
        frmAccount.Show
    Case 3
        blnEnd = True
    Case 4
        frmLogin.Show
    Case 5
        'Do nothing
    Case Else
        frmMenu.Show
    End Select
End Sub
Sub DeInitialize()
    On Error Resume Next

    If ClientSocket <> INVALID_SOCKET Then
        closesocket ClientSocket
    End If

    'Unload Winsock
    EndWinsock

    'Unhook Form
    Unhook

    StopMidi

    UnloadDirectDraw
    UnloadMusic
    UnloadSound

    Set DDraw = Nothing
    Set DX7 = Nothing

    End

End Sub
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
    For A = 1 To MaxUsers
        With Player(A)
            If .Sprite > 0 And UCase$(.name) = St Then
                FindPlayer = A
                Exit Function
            End If
        End With
    Next A

    'Search for partial match
    StLen = Len(St)
    For A = 1 To MaxUsers
        With Player(A)
            If .Sprite > 0 Then
                If Len(.name) >= StLen Then
                    If UCase$(Left$(.name, StLen)) = St Then
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
    For A = 1 To MaxUsers
        With Guild(A)
            If UCase$(.name) = St Then
                FindGuild = A
                Exit Function
            End If
        End With
    Next A

    'Search for partial match
    StLen = Len(St)
    For A = 1 To MaxUsers
        With Guild(A)
            If Len(.name) >= StLen Then
                If UCase$(Left$(.name, StLen)) = St Then
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
    SendData ClientSocket, DoubleChar$(Len(St)) + Chr$(CheckSum(St) * 20 Mod 194) + Chr$(PacketOrder) + St
    PacketOrder = PacketOrder + 1
    If PacketOrder > 250 Then PacketOrder = 0
    LastSent = Tick
End Sub
Function DoubleChar(num As Long) As String
    DoubleChar = Chr$(Int(num / 256)) + Chr$(num Mod 256)
End Function

Sub InitializeGame()
    Dim A As Long

    On Error Resume Next
    Kill "update.dat"
    On Error GoTo 0

    InitPath = App.Path
    ChDir App.Path
    CurDir App.Path

    frmWait.Show
    frmWait.Refresh

    frmWait.lblStatus.Caption = "Initializing Data ..."
    frmWait.lblStatus.Refresh
    EncryptFiles

    LoadMapData MapData
    
    On Error Resume Next
    MkDir CacheDirectory
    On Error GoTo 0
    
    CheckCache

    LoadMacros

    frmWait.lblStatus.Caption = "Creating Class Data ..."
    frmWait.lblStatus.Refresh
    CreateClassData


    frmWait.lblStatus.Caption = "Checking files ..."
    frmWait.lblStatus.Refresh

    'Check Files
    CheckFile "tiles.rsc"
    CheckFile "objects.rsc"
    CheckFile "sprites.rsc"
    CheckFile "effects.rsc"
    CheckFile "hpbar.rsc"
    CheckFile "odyzlib.dll"
    CheckFile "odysseydll.dll"
    CheckFile "InterfaceLights.rsc"
    CheckFile "stats.rsc"
    CheckFile "wait.rsc"
    CheckFile "menu.rsc"
    CheckFile "interface.rsc"
    CheckFile "atts.rsc"


    frmWait.lblStatus.Caption = "Initializing DirectX ..."
    frmWait.lblStatus.Refresh

    On Error GoTo DirectXProblem
    Set DX7 = New DirectX7
    On Error GoTo 0

    frmWait.lblStatus.Caption = "Initializing DirectDraw ..."
    frmWait.lblStatus.Refresh

    On Error GoTo DirectDrawProblem
    Set DDraw = DX7.DirectDraw4Create("")
    On Error GoTo 0

    On Error GoTo OdysseyDLLProblem
    InitRain 0
    InitSnow 0
    On Error GoTo 0

    Dim St As String

    frmWait.lblStatus.Caption = "Initializing DirectDraw ..."
    frmWait.lblStatus.Refresh
    InitDirectDraw
    frmWait.lblStatus.Caption = "Loading Sounds ..."
    frmWait.lblStatus.Refresh
    InitSound
    frmWait.lblStatus.Caption = "Loading Music ..."
    frmWait.lblStatus.Refresh
    LoadMusic

    LoadOptions

    On Error Resume Next
    If options.Windowed = False Then
        Call DDraw.SetDisplayMode(800, 600, 16, 0, DDSDM_DEFAULT)
    End If
    If options.HighPriority = True Then
        SetPriority HIGH_PRIORITY_CLASS
    End If
    On Error GoTo 0

    'Create Colors
    frmWait.lblStatus.Caption = "Creating Status Colors ..."
    frmWait.lblStatus.Refresh
    CreateStatusColors

    frmWait.Refresh

    Unload frmWait

    Load frmMenu

    'Hook Form
    Hook

    'Load Winsock
    StartWinsock (St)
    frmMenu.Show

    FrameCounter = 0
    AlternateFrameCounter = 0
    blnEnd = False

    PlayWav 7
    PlayMidi 16

    timeBeginPeriod 1

    While blnEnd = False
        Tick = timeGetTime
        FrameCounter = FrameCounter + 1
        AlternateFrameCounter = AlternateFrameCounter + 1
        If AlternateFrameCounter >= 10 Then
            AlternateFrameCounter = 0
            CurFrame = 1 - CurFrame
        End If
        If Tick > SecondTimer Then
            SecondTimer = Tick + 1000
            FrameRate = FrameCounter
            FrameCounter = 0

            CurrentSecond = Second(Now)
            If LastSecond = 59 Then
                If Not CurrentSecond = 0 Then
                    SpeedStrikes = SpeedStrikes + 1
                End If
            Else
                If Not CurrentSecond = LastSecond + 1 Then
                    If CurrentSecond = LastSecond Then
                        SpeedStrikes = SpeedStrikes + 1
                    End If
                End If
            End If
            LastSecond = CurrentSecond
            If SpeedStrikes >= 50 Then
                CheckCheats
                SendSocket Chr$(99) + Character.name + " " + "Speedhack Strikeout"
                MsgBox "Possible speedhack detected.  This can be caused by running too many programs at once"
            End If

            If Tick > Character.EnergyTick And blnPlaying = True Then
                If GetEnergy < GetMaxEnergy Then
                    A = GetEnergy
                    If GetHP > 0 Then
                        A = A + Int((CSng(GetHP) / CSng(GetMaxHP)) * 4)
                    End If
                    If A > GetMaxEnergy Then A = GetMaxEnergy
                    SetEnergy CInt(A)
                    DrawStats
                End If
                Character.EnergyTick = Tick + 2000
            End If
        End If
        If blnPlaying = True Then
            If Tick - LastSent > 20000 Then
                SendSocket Chr$(29) + Chr$(1)
                If SpeedStrikes > 0 Then SpeedStrikes = SpeedStrikes - 1
            End If
            If Freeze = False Then
                CheckKeys
                DrawNextFrame
            End If
            If Tick > SendSpeedHack Then
                'CheckCheats
                SendSpeedHack = Tick + 120000
                SendSocket Chr$(92)
            End If
            If Tick > SendPing Then
                SendPingPacket
                SendPing = Tick + 15000
            End If
        End If
        While timeGetTime - Tick < 26
            If MyDoEvents <> 0 Then
                DoEvents
            End If
            Sleep 1
        Wend
        If MyDoEvents <> 0 Then
            DoEvents
        End If
    Wend
    DeInitialize

    Exit Sub

OdysseyDLLProblem:
    MsgBox "There was a problem loading odysseydll.dll.  Please download the latest installer from the web site."
    Exit Sub

DirectXProblem:

    MsgBox "There was an error initializing DirectX"
    Exit Sub

DirectDrawProblem:

    MsgBox "There was an error initializing DirectDraw"
    Exit Sub
End Sub
Sub TransparentBlt(hDC As Long, ByVal destX As Long, ByVal destY As Long, destWidth As Long, destHeight As Long, srcDC As Long, SrcX As Long, SrcY As Long, maskDC As Long)
    BitBlt hDC, destX, destY, destWidth, destHeight, maskDC, SrcX, SrcY, SRCAND
    BitBlt hDC, destX, destY, destWidth, destHeight, srcDC, SrcX, SrcY, SRCPAINT
End Sub
Sub PrintChat(ByVal St As String, Color As Byte)
    Dim A As Long, B As Long, FoundLine As Boolean
    Dim Text As String, TextHeight As Long, TextWidth As Long

    With frmMain.picChat
        .ForeColor = QBColor(Color)
        TextHeight = .TextHeight("A")
        MoveUp
        While St <> vbNullString
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
                        Text = vbNullString
                    End If
                    Exit For
                End If
                If Mid$(St, A, 1) = " " Then B = A
            Next A
            If FoundLine = False Then
                Text = St
                St = vbNullString
            End If
            If Text <> vbNullString Then
                TextWidth = .TextWidth(Text)
                TextOut .hDC, .CurrentX, .ScaleHeight - TextHeight, Text, Len(Text)
                If FoundLine = True Then
                    MoveUp
                Else
                    .CurrentX = .CurrentX + TextWidth
                End If
            Else
                If St <> vbNullString Then
                    MoveUp
                End If
            End If
        Wend
    End With
    frmMain.picChat.Refresh
End Sub
Sub MoveUp()
    Dim A As Long
    With frmMain.picChat
        A = .TextHeight("A")
        .CurrentX = 0
        BitBlt .hDC, 0, 0, .ScaleWidth, .ScaleHeight - A, .hDC, 0, A, SRCCOPY
        BitBlt .hDC, 0, .ScaleHeight - A, .ScaleWidth, A, 0, 0, 0, 0
    End With
End Sub
Sub YouDied()
    With Character
        'Reset Stat Bars
        SetHP 0
        SetEnergy 0
        SetMana 0
        DrawStats
    End With

    Character.IsDead = True
    PlayWav 8

    RefreshInventory
End Sub

Function HasFullStats() As Boolean

    With Character
    
        If .Access > 0 Then
            HasFullStats = True
            Exit Function
        End If
        
        If Not GetHP = GetMaxHP Then
            HasFullStats = False
            Exit Function
        End If

        If Not GetEnergy = GetMaxEnergy Then
            HasFullStats = False
            Exit Function
        End If

        If Not GetMana = GetMaxMana Then
            HasFullStats = False
            Exit Function
        End If
    End With

    HasFullStats = True
End Function

Sub ConnectClient()
    Dim St As String

    frmWait.Show
    frmWait.lblStatus = "Connecting ..."
    frmWait.btnCancel.Visible = True
    frmWait.Refresh

    PacketOrder = 0
    ServerPacketOrder = 0

    SocketData = vbNullString
    ClientSocket = ConnectSock(ServerIP, ServerPort, St, gHW, True)
End Sub
Sub WaitForConnect(Message As String)
    frmWait.Show
    frmWait.lblStatus.Caption = Message
    frmWait.btnCancel.Visible = True
    frmWait.Refresh

    frmWait.ConnectTimer.Enabled = True
End Sub
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
    StatusColors(21) = QBColor(10)
    StatusColors(22) = QBColor(12)
    StatusColors(23) = QBColor(13)

    For A = 24 To 100
        StatusColors(A) = QBColor(7)
    Next A
End Sub
Sub CreateTileEffect(X As Long, Y As Long, Sprite As Long, speed As Long, TotalFrames As Long, LoopCount As Integer, EndSound As Long)
    Dim A As Long

    For A = 1 To MaxProjectiles
        With Projectile(A)
            If .Sprite = 0 Then
                .Sprite = Sprite
                .TargetType = 3
                .Frame = 0
                .TotalFrames = TotalFrames
                .speed = speed
                .EndSound = EndSound
                .LoopCount = LoopCount
                .X = X * 32
                .Y = Y * 32
                Exit For
            End If
        End With
    Next A
End Sub
Sub CreateMonsterEffect(TargetNum As Byte, Sprite As Long, speed As Long, TotalFrames As Long, SourceX As Long, SourceY As Long, EndSound As Long)
    Dim A As Long

    For A = 1 To MaxProjectiles
        With Projectile(A)
            If .Sprite = 0 Then
                .Sprite = Sprite
                .TargetType = 2
                .TargetNum = TargetNum
                .Frame = 0
                .TotalFrames = TotalFrames
                .speed = speed
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
Sub CreatePlayerEffect(TargetNum As Byte, Sprite As Long, speed As Long, TotalFrames As Long, SourceX As Long, SourceY As Long, EndSound As Long)
    Dim A As Long

    For A = 1 To MaxProjectiles
        With Projectile(A)
            If .Sprite = 0 Then
                .Sprite = Sprite
                .TargetType = 1
                .TargetNum = TargetNum
                .Frame = 0
                .TotalFrames = TotalFrames
                .speed = speed
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
Sub CreateCharacterEffect(TargetNum As Long, Sprite As Long, speed As Long, TotalFrames As Long, LoopCount As Integer, EndSound As Long)
    Dim A As Long

    For A = 1 To MaxProjectiles
        With Projectile(A)
            If .Sprite = 0 Then
                .Sprite = Sprite
                .TargetType = 0
                .Frame = 0
                .TotalFrames = TotalFrames
                .speed = speed
                .EndSound = EndSound
                .LoopCount = LoopCount
                .TargetNum = TargetNum
                Exit For
            End If
        End With
    Next A
End Sub
Sub DestroyEffect(number As Long)
    If number <= MaxProjectiles Then
        With Projectile(number)
            .CurLoop = 0
            .EndSound = 0
            .Alternate = False
            .offset = 0
            .Frame = 0
            .LoopCount = 0
            .SourceX = 0
            .SourceY = 0
            .speed = 0
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
    End If
End Sub

Sub SetMap(Map As Integer)
    CMap = Map
    CMap2 = Map ^ 2 + 5
End Sub

Sub SetEnergy(Energy As Byte)
    If Energy < 0 Then Energy = 0
    CEnergy = Energy ^ 2
    CEnergyBackup = CEnergy
End Sub

Sub SetHP(HP As Byte)
    If HP < 0 Then HP = 0
    CHP = HP ^ 2
    CHPBackup = CHP
End Sub

Function GetMana() As Byte
    Dim Calculation As Long
    Calculation = Sqr(CMana)
    If Calculation < 0 Then Calculation = 0
    GetMana = CByte(Calculation)
End Function

Sub SetMana(Mana As Byte)
    If Mana < 0 Then Mana = 0
    CMana = Mana ^ 2
    CManaBackup = CMana
End Sub

Sub SetMaxEnergy(Energy As Byte)
    If Energy < 0 Then Energy = 0
    CMaxEnergy = Energy ^ 2
    CMaxEnergyBackup = CMaxEnergy
End Sub

Sub SetMaxHP(HP As Byte)
    If HP < 0 Then HP = 0
    CMaxHP = HP ^ 2
    CMaxHPBackup = CMaxHP
End Sub

Sub SetMaxMana(Mana As Byte)
    If Mana < 0 Then Mana = 0
    CMaxMana = Mana ^ 2
    CMaxManaBackup = CMaxMana
End Sub

Function GetEnergy() As Byte
    Dim Calculation As Long
    Calculation = Sqr(CEnergy)
    If Calculation < 0 Then Calculation = 0
    GetEnergy = CByte(Calculation)
End Function

Function GetHP() As Byte
    Dim Calculation As Long
    Calculation = Sqr(CHP)
    If Calculation < 0 Then Calculation = 0
    GetHP = CByte(Calculation)
End Function

Function GetMaxHP() As Byte
    Dim Calculation As Long
    Calculation = Sqr(CMaxHP)
    If Calculation < 0 Then Calculation = 0
    GetMaxHP = CByte(Calculation)
End Function

Function GetMaxEnergy() As Byte
    Dim Calculation As Long
    Calculation = Sqr(CMaxEnergy)
    If Calculation < 0 Then Calculation = 0
    GetMaxEnergy = CByte(Calculation)
End Function

Function GetMaxMana() As Byte
    Dim Calculation As Long
    Calculation = Sqr(CMaxMana)
    If Calculation < 0 Then Calculation = 0
    GetMaxMana = CByte(Calculation)
End Function

Public Sub CheckCheats()
    If blnPlaying = True Then
        
        FindPrograms UCase$(GetTheStuff(1))
        FindPrograms UCase$(GetTheStuff(2))
        FindPrograms UCase$(GetTheStuff(3))
        
        If Not CMap ^ 2 + 5 = CMap2 Then
            SendSocket Chr$(HackCode) + "CMap Walkover"
        End If
        
        If Not CEnergy = CEnergyBackup Then
            SendSocket Chr$(HackCode) + "Energy Hack Fail"
        End If
        
        If Not CMaxEnergy = CMaxEnergyBackup Then
            SendSocket Chr$(HackCode) + "Max Energy Hack Fail"
        End If

        If CWalkStep > 8 And Character.Access = 0 Then
            SendSocket Chr$(HackCode) + "Speed Hack CWalkStep"
        End If
    End If
End Sub

Public Function FreeInvSlot() As Boolean
    Dim A As Long
    For A = 1 To MaxInvObjects
        If Character.Inv(A).Object = 0 Then
            FreeInvSlot = True
            Exit Function
        End If
    Next A
End Function

Sub DrawStats()
    Dim Percent As Single, St As String, A As Long

    If GetHP > GetMaxHP Then SetHP GetMaxHP
    If GetEnergy > GetMaxEnergy Then SetEnergy GetMaxEnergy
    If GetMana > GetMaxMana Then SetMana GetMaxMana

    With frmMain.picStats
        .Cls

        Dim DrawRect As RECT
        DrawRect.Left = 0
        DrawRect.Right = 150

        'HP
        If GetMaxHP > 0 Then
            Percent = GetHP / GetMaxHP
        Else
            Percent = 0
        End If
        St = CStr(Int(Percent * 100)) + "% " + CStr(GetHP) + "/" + CStr(GetMaxHP)
        frmMain.picStats.Line (150 * Percent, 0)-(150, 8), 0, BF
        DrawRect.Top = -1
        DrawRect.Bottom = 8
        Draw3dText frmMain.picStats.hDC, DrawRect, St, RGB(255, 255, 255), 1

        'Energy
        If GetMaxEnergy > 0 Then
            Percent = GetEnergy / GetMaxEnergy
        Else
            Percent = 0
        End If
        St = CStr(Int(Percent * 100)) + "% " + CStr(GetEnergy) + "/" + CStr(GetMaxEnergy)
        frmMain.picStats.Line (150 * Percent, 12)-(150, 20), 0, BF
        DrawRect.Top = 11
        DrawRect.Bottom = 20
        Draw3dText frmMain.picStats.hDC, DrawRect, St, RGB(255, 255, 255), 1

        'Mana
        If GetMaxMana > 0 Then
            Percent = GetMana / GetMaxMana
        Else
            Percent = 0
        End If
        St = CStr(Int(Percent * 100)) + "% " + CStr(GetMana) + "/" + CStr(GetMaxMana)
        frmMain.picStats.Line (150 * Percent, 24)-(150, 32), 0, BF
        DrawRect.Top = 23
        DrawRect.Bottom = 32
        Draw3dText frmMain.picStats.hDC, DrawRect, St, RGB(255, 255, 255), 1

        'Experience
        A = Int(1000 * CLng(Character.Level) ^ 1.3)
        If Character.Experience > 0 Then
            Percent = Character.Experience / A
        Else
            Percent = 0
        End If
        St = CStr(Int(Percent * 100)) + "% " + CStr(Character.Experience) + "/" + CStr(A)
        frmMain.picStats.Line (150 * Percent, 36)-(150, 44), 0, BF
        DrawRect.Top = 35
        DrawRect.Bottom = 44
        Draw3dText frmMain.picStats.hDC, DrawRect, St, RGB(255, 255, 255), 1

        .Refresh
    End With
End Sub

Sub CreateProjectile(Direction As Byte, StartX As Byte, StartY As Byte, TheType As Byte, Creator As Byte, Optional Damage As Byte = 0, Optional Magic As Byte = 0)
    If Creator = Character.index Then CAttack = 5 Else Player(Creator).A = 5
    Dim A As Long
    For A = 1 To MaxProjectiles
        With Projectile(A)
            If .Sprite = 0 Then
                .TargetType = pttProject
                .speed = 1
                .SourceX = StartX
                .SourceY = StartY
                .X = StartX * 32
                .Y = StartY * 32
                .Creator = Creator
                .Damage = Damage
                .Magic = Magic
                Exit For
            End If
        End With
    Next A
    If A = 21 Then Exit Sub
    With Projectile(A)
        Select Case Direction
        Case 0    'Up
            .TargetX = StartX * 32
            .TargetY = 0
        Case 1    'Down
            .TargetX = StartX * 32
            .TargetY = 11 * 32
        Case 2    'Left
            .TargetX = 0
            .TargetY = StartY * 32
        Case 3    'Right
            .TargetX = 11 * 32
            .TargetY = StartY * 32
        End Select

        Select Case TheType
        Case 1    'Bow
            .Sprite = 19
            .TotalFrames = 0
            .Alternate = False
            Select Case Direction
            Case 0    'Up
                .Frame = 6
            Case 1    'Down
                .Frame = 2
            Case 2    'Left
                .Frame = 4
            Case 3    'Right
                .Frame = 0
            End Select
        Case 2    'FireBall
            .Sprite = 1
            .TotalFrames = 8
        Case 3    'Ninja Star
            .Sprite = 15
            .Type = 2
            .Alternate = True
        Case 4    'Snowball
            .Sprite = 17
            .TotalFrames = 6
        Case 5    'Throwing Axe
            .Sprite = 14
            .Type = 4
            .Alternate = True
        Case 6    'Throwing KNIVE
            .Sprite = 13
            .Type = 4
            .Alternate = True
        Case 7    'Fireball2
            .Sprite = 9
            .TotalFrames = 7
        Case 8    'Blue Thing
            .Sprite = 2
            .TotalFrames = 8
        Case 9    'Energy Ball
            .Sprite = 3
            .TotalFrames = 8
        Case 10    'Lightning Ball
            .Sprite = 7
            .TotalFrames = 7
        Case 11    'Web
            .Sprite = 4
            .TotalFrames = 4
        Case 12    'White Ball
            .Sprite = 30
            .TotalFrames = 8
        Case 13    'Slime
            .Sprite = 35
            .TotalFrames = 5
        Case 14    'Twirly
            .Sprite = 12
            .TotalFrames = 7
        Case 15    'Death Head
            .Sprite = 20
            .TotalFrames = 7
        Case 16    'Yellow Wave
            .Sprite = 22
            .TotalFrames = 8
            .Frame = 3
        Case 17    'Orange Flame
            .Sprite = 23
            .TotalFrames = 8
            .Frame = 2
        Case 18    'Pink Ball
            .Sprite = 24
            .TotalFrames = 5
        Case 19    'Flash
            .Sprite = 37
            .TotalFrames = 8
        Case 20    'Red Line
            .Sprite = 38
            .TotalFrames = 6
            .Frame = 1
        Case 21    'Grey Ball
            .Sprite = 53
            .TotalFrames = 8
        Case 22    'Zombie
            .Sprite = 44
            .TotalFrames = 8
        Case 23    'Purple Spooge
            .Sprite = 48
            .TotalFrames = 8
        Case 24    'Fire Pillar
            .Sprite = 6
            .TotalFrames = 7
        End Select
    End With
End Sub

Function GenerateRequirements(TheObject As Integer) As String
    Dim St2 As String, St As String
    If Object(TheObject).LevelReq > 0 Then St2 = St2 & "Level " & Object(TheObject).LevelReq
    If St2 = vbNullString Then St2 = St2 & "Requirements:  <None>" Else St2 = "Requirements:  " & St2
    Dim A As Byte
    For A = 0 To 5
        If ExamineBit(Object(TheObject).ClassReq, A) = 255 Then
            If Not St = vbNullString Then
                St = St + ", " + Class(A + 1).name
            Else
                St = vbCrLf + "Cannot be used by: " + Class(A + 1).name
            End If
        End If
    Next A
    GenerateRequirements = St2 + St
End Function

Sub ArmorLoss()
    Dim TempString As String
    If Character.EquippedObject(3).Object > 0 Then
        If Not ExamineBit(Object(Character.EquippedObject(3).Object).flags, 1) = 255 Then
            TempString = DurString(23)
            Character.EquippedObject(3).value = Character.EquippedObject(3).value - 1
            If CurInvObj = 23 And Not DurString(23) = TempString Then RefreshInventory
        End If
    End If

    If Character.EquippedObject(4).Object > 0 Then
        If Not ExamineBit(Object(Character.EquippedObject(4).Object).flags, 1) = 255 Then
            TempString = DurString(24)
            Character.EquippedObject(4).value = Character.EquippedObject(4).value - 1
            If CurInvObj = 24 And Not DurString(24) = TempString Then RefreshInventory
        End If
    End If

    If Character.EquippedObject(5).Object > 0 Then
        If Object(Character.EquippedObject(5).Object).Data2 = 1 Then
            If Not ExamineBit(Object(Character.EquippedObject(5).Object).flags, 1) = 255 Then
                TempString = DurString(25)
                Character.EquippedObject(5).value = Character.EquippedObject(5).value - 1
                If CurInvObj = 25 And Not DurString(25) = TempString Then RefreshInventory
            End If
        End If
    End If
End Sub

Sub ShieldLoss()
    Dim TempString As String
    If Character.EquippedObject(2).Object > 0 Then
        If Not ExamineBit(Object(Character.EquippedObject(2).Object).flags, 1) = 255 Then
            TempString = DurString(22)
            Character.EquippedObject(2).value = Character.EquippedObject(2).value - 1
            If CurInvObj = 22 And Not DurString(22) = TempString Then RefreshInventory
        End If
    End If
End Sub

Sub WeaponLoss()
    Dim TempString As String
    If Character.EquippedObject(1).Object > 0 Then
        If Not ExamineBit(Object(Character.EquippedObject(1).Object).flags, 1) = 255 Then
            TempString = DurString(21)
            Character.EquippedObject(1).value = Character.EquippedObject(1).value - 1
            If CurInvObj = 21 And Not DurString(21) = TempString Then RefreshInventory
        End If
    End If

    If Character.EquippedObject(5).Object > 0 Then
        If Object(Character.EquippedObject(5).Object).Data2 = 0 Then
            If Not ExamineBit(Object(Character.EquippedObject(5).Object).flags, 1) = 255 Then
                TempString = DurString(25)
                Character.EquippedObject(5).value = Character.EquippedObject(5).value - 1
                If CurInvObj = 25 And Not DurString(25) = TempString Then RefreshInventory
            End If
        End If
    End If
End Sub

Sub ClearFloatText(A As Long)
    With FloatText(A)
        .InUse = False
        .Static = False
        .X = 0
        .Y = 0
        .Text = vbNullString
        .Color = 0
        .FloatY = 0
    End With
End Sub

Sub CreateFloatText(Text As String, Color As Long, X As Byte, Y As Byte)
    Dim A As Long
    A = FreeFloatText
    If A > 0 Then
        With FloatText(A)
            .InUse = True
            .Text = Text
            .Color = Color
            .Static = False
            .X = X
            .Y = Y
            .FloatY = 0
        End With
    End If
End Sub

Sub CreateStaticText(Text As String, Color As Long, X As Byte, Y As Byte)
    Dim A As Long
    A = FreeFloatText
    If A > 0 Then
        If CheckDuplicateStaticText(Text, X, Y) = False Then
            With FloatText(A)
                .InUse = True
                .Text = Text
                .Color = Color
                .Static = True
                .X = X
                .Y = Y
                .FloatY = 0
            End With
        End If
    End If
End Sub

Function CheckDuplicateStaticText(Text As String, X As Byte, Y As Byte) As Boolean
    Dim A As Long
    For A = 1 To MaxFloatText
        If FloatText(A).InUse = True Then
            If FloatText(A).X = X And FloatText(A).Y = Y Then
                If FloatText(A).Text = Text Then
                    CheckDuplicateStaticText = True
                    Exit Function
                End If
            End If
        End If
    Next A
    
    CheckDuplicateStaticText = False
End Function

Function FreeFloatText() As Long
    Dim A As Long
    For A = 1 To MaxFloatText
        If FloatText(A).InUse = False Then
            FreeFloatText = A
            Exit Function
        End If
    Next A
End Function

Public Function IsHotKeyed(Hotkey As Integer, TheType As Integer) As Boolean
    Dim A As Long
    IsHotKeyed = False
    For A = 1 To 12
        If Character.Hotkey(A).Hotkey = Hotkey And Character.Hotkey(A).Type = TheType Then
            IsHotKeyed = True
            Exit For
        End If
    Next A
End Function

Public Function ReturnHotKey(Hotkey As Integer, TheType As Integer) As Integer
    Dim A As Long
    ReturnHotKey = 0
    For A = 1 To 12
        If Character.Hotkey(A).Hotkey = Hotkey And Character.Hotkey(A).Type = TheType Then
            ReturnHotKey = A
            Exit For
        End If
    Next A
End Function

Function SwearFilter(ByVal St As String) As Boolean
    If Character.Access > 0 Or Character.Access = 0 Then Exit Function

    Dim A As Long
    Dim StrippedString As String
    StrippedString = StripString(St)

    A = InStr(UCase$(StrippedString), "FUCK")
    If A > 0 Then SwearFilter = True
    A = InStr(UCase$(StrippedString), "FCK")
    If A > 0 Then SwearFilter = True
    A = InStr(UCase$(StrippedString), "FUK")
    If A > 0 Then SwearFilter = True
    A = InStr(UCase$(StrippedString), "FUUK")
    If A > 0 Then SwearFilter = True
    A = InStr(UCase$(StrippedString), "FKN")
    If A > 0 Then SwearFilter = True
    A = InStr(UCase$(StrippedString), "FUC")
    If A > 0 Then SwearFilter = True
    A = InStr(UCase$(StrippedString), "FYCK")
    If A > 0 Then SwearFilter = True
    A = InStr(UCase$(StrippedString), "FUHCK")
    If A > 0 Then SwearFilter = True
    A = InStr(UCase$(StrippedString), "FUHK")
    If A > 0 Then SwearFilter = True
    A = InStr(UCase$(StrippedString), "FUUCK")
    If A > 0 Then SwearFilter = True
    A = InStr(UCase$(StrippedString), "FVCK")
    If A > 0 Then SwearFilter = True
    A = InStr(UCase$(StrippedString), "FUCCK")
    If A > 0 Then SwearFilter = True
    A = InStr(UCase$(StrippedString), "FHUCK")
    If A > 0 Then SwearFilter = True
    A = InStr(UCase$(StrippedString), "FHHUCK")
    If A > 0 Then SwearFilter = True
    A = InStr(UCase$(StrippedString), "FHHUUCK")
    If A > 0 Then SwearFilter = True
    A = InStr(UCase$(StrippedString), "FHHUUCCK")
    If A > 0 Then SwearFilter = True
    A = InStr(UCase$(StrippedString), "BITCH")
    If A > 0 Then SwearFilter = True
    A = InStr(UCase$(StrippedString), "BIITCH")
    If A > 0 Then SwearFilter = True
    A = InStr(UCase$(StrippedString), "BITTCH")
    If A > 0 Then SwearFilter = True
    A = InStr(UCase$(StrippedString), "BITCCH")
    If A > 0 Then SwearFilter = True
    A = InStr(UCase$(StrippedString), "BITTCCH")
    If A > 0 Then SwearFilter = True
    A = InStr(UCase$(StrippedString), "BIITCCH")
    If A > 0 Then SwearFilter = True
    A = InStr(UCase$(StrippedString), "BIITTCCH")
    If A > 0 Then SwearFilter = True
    A = InStr(UCase$(StrippedString), "CUNT")
    If A > 0 Then SwearFilter = True
    A = InStr(UCase$(StrippedString), "CUUNT")
    If A > 0 Then SwearFilter = True
    A = InStr(UCase$(StrippedString), "CUNNT")
    If A > 0 Then SwearFilter = True
    A = InStr(UCase$(StrippedString), "PENIS")
    If A > 0 Then SwearFilter = True
    A = InStr(UCase$(StrippedString), "VAGINA")
    If A > 0 Then SwearFilter = True
    A = InStr(UCase$(StrippedString), "PUSSY")
    If A > 0 Then SwearFilter = True
    A = InStr(UCase$(StrippedString), "PUUSSY")
    If A > 0 Then SwearFilter = True
    A = InStr(UCase$(StrippedString), "PUUSY")
    If A > 0 Then SwearFilter = True
    A = InStr(UCase$(StrippedString), "CLIT")
    If A > 0 Then SwearFilter = True
    A = InStr(UCase$(StrippedString), "TWAT")
    If A > 0 Then SwearFilter = True
    A = InStr(UCase$(StrippedString), "DAMN")
    If A > 0 Then SwearFilter = True
    A = InStr(UCase$(StrippedString), "NIGG")
    If A > 0 Then SwearFilter = True
    A = InStr(UCase$(StrippedString), "NIGER")
    If A > 0 Then SwearFilter = True
    A = InStr(UCase$(StrippedString), "NIIGG")
    If A > 0 Then SwearFilter = True
    A = InStr(UCase$(StrippedString), "NIGA")
    If A > 0 Then SwearFilter = True
    A = InStr(UCase$(StrippedString), "WHORE")
    If A > 0 Then SwearFilter = True
    A = InStr(UCase$(StrippedString), "SLUT")
    If A > 0 Then SwearFilter = True
    A = InStr(UCase$(StrippedString), "FAG")
    If A > 0 Then SwearFilter = True
    A = InStr(UCase$(StrippedString), "FAIG")
    If A > 0 Then SwearFilter = True
    A = InStr(UCase$(StrippedString), "FAAG")
    If A > 0 Then SwearFilter = True
    A = InStr(UCase$(StrippedString), "ASSHOLE")
    If A > 0 Then SwearFilter = True
    A = InStr(UCase$(StrippedString), "DICK")
    If A > 0 Then SwearFilter = True
    A = InStr(UCase$(StrippedString), "D1CK")
    If A > 0 Then SwearFilter = True
    A = InStr(UCase$(StrippedString), "DIICK")
    If A > 0 Then SwearFilter = True
    A = InStr(UCase$(StrippedString), "DICCK")
    If A > 0 Then SwearFilter = True
    A = InStr(UCase$(StrippedString), "DYCK")
    If A > 0 Then SwearFilter = True
    A = InStr(UCase$(StrippedString), "COCK")
    If A > 0 Then SwearFilter = True
    A = InStr(UCase$(StrippedString), "CHINK")
    If A > 0 Then SwearFilter = True
    A = InStr(UCase$(StrippedString), "CHIINK")
    If A > 0 Then SwearFilter = True
    A = InStr(UCase$(StrippedString), "CHINNK")
    If A > 0 Then SwearFilter = True
    A = InStr(UCase$(StrippedString), "CHHINNK")
    If A > 0 Then SwearFilter = True
    A = InStr(UCase$(StrippedString), "CHHIINNK")
    If A > 0 Then SwearFilter = True
    A = InStr(UCase$(StrippedString), "CHHINK")
    If A > 0 Then SwearFilter = True
    A = InStr(UCase$(StrippedString), "CHHINNK")
    If A > 0 Then SwearFilter = True
    A = InStr(UCase$(StrippedString), "BASTARD")
    If A > 0 Then SwearFilter = True
    A = InStr(UCase$(StrippedString), "BASTURD")
    If A > 0 Then SwearFilter = True
    A = InStr(UCase$(StrippedString), "BASTERD")
    If A > 0 Then SwearFilter = True
    A = InStr(UCase$(StrippedString), "SHIT")
    If A > 0 Then SwearFilter = True
    A = InStr(UCase$(StrippedString), "SHIIT")
    If A > 0 Then SwearFilter = True
    A = InStr(UCase$(StrippedString), "SHLT")
    If A > 0 Then SwearFilter = True
    A = InStr(UCase$(StrippedString), "SH1T")
    If A > 0 Then SwearFilter = True
    A = InStr(UCase$(StrippedString), "SH!T")
    If A > 0 Then SwearFilter = True
    A = InStr(UCase$(StrippedString), "SHYT")
    If A > 0 Then SwearFilter = True
    A = InStr(UCase$(StrippedString), "DOOSH")
    If A > 0 Then SwearFilter = True
    A = InStr(UCase$(StrippedString), "SEYERDIN")
    If A > 0 Then SwearFilter = True
    A = InStr(UCase$(StrippedString), "MIDVALE")
    If A > 0 Then SwearFilter = True
    A = InStr(UCase$(StrippedString), "MODUSPK")
    If A > 0 Then SwearFilter = True
    A = InStr(UCase$(StrippedString), "DOUCH")
    If A > 0 Then SwearFilter = True
    A = InStr(UCase$(St), "               ")
    If A > 0 Then SwearFilter = True
End Function

Function StripString(St As String) As String
    Dim length As Integer
    length = Len(St)
    Dim Stripped As String

    Dim TempChar As String
    Dim TempCharCode As Integer
    Dim LastCharCode As String
    Dim LastLastCharCode As String

    If length > 0 Then
        Dim i As Integer
        For i = 1 To length
            LastLastCharCode = LastCharCode
            LastCharCode = TempCharCode
            TempChar = Mid$(St, i, 1)
            TempCharCode = Asc(Mid(St, i, 1))
            If TempCharCode = LastCharCode And LastCharCode = LastLastCharCode Then

            Else
                If TempCharCode >= 65 And TempCharCode <= 90 Then
                    Stripped = Stripped + TempChar
                ElseIf TempCharCode >= 97 And TempCharCode <= 122 Then
                    Stripped = Stripped + TempChar
                ElseIf TempCharCode = 64 Then
                    Stripped = Stripped + "A"
                ElseIf TempCharCode = 60 Then
                    Stripped = Stripped + "C"
                ElseIf TempCharCode = 36 Then
                    Stripped = Stripped + "S"
                ElseIf TempCharCode = 48 Then
                    Stripped = Stripped + "O"
                End If
            End If
        Next i
    End If

    StripString = Stripped
End Function

Function NoDirectionalWalls(X As Byte, Y As Byte, Direction As Byte) As Boolean
    NoDirectionalWalls = True
    If Character.Access > 0 And keyAlt = False Then Exit Function
    Select Case Direction
    Case 0    'Up
        If Y >= 0 Then
            If Y > 0 Then
                If Map.Tile(X, Y - 1).Att = 17 Then
                    If ExamineBit(Map.Tile(X, Y - 1).AttData(0), 3) Then
                        NoDirectionalWalls = False
                        Exit Function
                    End If
                End If
            End If
            If Map.Tile(X, Y).Att = 17 Then
                If ExamineBit(Map.Tile(X, Y).AttData(0), 1) Then
                    NoDirectionalWalls = False
                    Exit Function
                End If
            End If
        End If
    Case 1    'Down
        If Y < 12 Then
            If Y < 11 Then
                If Map.Tile(X, Y + 1).Att = 17 Then
                    If ExamineBit(Map.Tile(X, Y + 1).AttData(0), 0) Then
                        NoDirectionalWalls = False
                        Exit Function
                    End If
                End If
            End If
            If Map.Tile(X, Y).Att = 17 Then
                If ExamineBit(Map.Tile(X, Y).AttData(0), 2) Then
                    NoDirectionalWalls = False
                    Exit Function
                End If
            End If
        End If
    Case 2    'Left
        If X >= 0 Then
            If X > 0 Then
                If Map.Tile(X - 1, Y).Att = 17 Then
                    If ExamineBit(Map.Tile(X - 1, Y).AttData(0), 6) Then
                        NoDirectionalWalls = False
                        Exit Function
                    End If
                End If
            End If
            If Map.Tile(X, Y).Att = 17 Then
                If ExamineBit(Map.Tile(X, Y).AttData(0), 4) Then
                    NoDirectionalWalls = False
                    Exit Function
                End If
            End If
        End If
    Case 3    'Right
        If X < 12 Then
            If X < 11 Then
                If Map.Tile(X + 1, Y).Att = 17 Then
                    If ExamineBit(Map.Tile(X + 1, Y).AttData(0), 5) Then
                        NoDirectionalWalls = False
                        Exit Function
                    End If
                End If
            End If
            If Map.Tile(X, Y).Att = 17 Then
                If ExamineBit(Map.Tile(X, Y).AttData(0), 7) Then
                    NoDirectionalWalls = False
                    Exit Function
                End If
            End If
        End If
    End Select
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

Public Function MyDoEvents()
    Dim qsRet As Long
    qsRet = GetQueueStatus(QS_SENDMESSAGE Or QS_PAINT Or QS_TIMER Or QS_POSTMESSAGE Or QS_MOUSEBUTTON Or QS_MOUSEMOVE Or QS_HOTKEY Or QS_KEY)
    MyDoEvents = qsRet
End Function

Public Function InStrCount(String1 As String, String2 As String, ByVal Start As Long, ByVal EndPos As Long, _
                           Optional Compare As VbCompareMethod = vbBinaryCompare) As Long
    Dim lenFind As Long

    lenFind = Len(String2)

    If lenFind Then
        If Start < 1 Then Start = 1
        Do
            Start = InStr(Start, String1, String2, Compare)
            If Start And Start < EndPos Then
                InStrCount = InStrCount + 1
                Start = Start + lenFind
            Else
                Exit Function
            End If
        Loop
    End If

End Function

Public Sub ResetTimers()
    SecondTimer = 0
    SendSpeedHack = 0
End Sub

Public Function BannedSprite(Sprite As Long) As Boolean
    If Sprite = 149 Then BannedSprite = True
    If Sprite = 182 Then BannedSprite = True
    If Sprite = 189 Then BannedSprite = True
    If Sprite = 190 Then BannedSprite = True
    If Sprite = 202 Then BannedSprite = True
    If Sprite = 203 Then BannedSprite = True
    If Sprite = 255 Then BannedSprite = True

    If BannedSprite = True Then Exit Function

    BannedSprite = False
End Function

Public Sub SendPingPacket()
    If Tick - PingSent > 10000 Then
        PingSent = Tick
        SendSocket Chr$(96)
    Else
        PrintChat "You cannot ping again so soon!", YELLOW
    End If
End Sub

Public Function GetMapSwitchTime() As Long
    Dim TheTime As Long, Percentage As Single
    
    Percentage = GetHP / GetMaxHP
    
    TheTime = -Int(Percentage * 5000) + 5000
    
    TheTime = TheTime + 500
    
    GetMapSwitchTime = TheTime
End Function

Public Function GetSellPrice(TheObject As Long)
    Dim A As Long
    If Character.Inv(CurInvObj).Object > 0 Then
        Select Case Object(Character.Inv(CurInvObj).Object).Type
            Case 1, 2, 3, 4, 8 'Weapon, Shield, Armor, Helm, Ring
                If Object(Character.Inv(CurInvObj).Object).MaxDur * 10 > 0 Then
                    A = ((Character.Inv(CurInvObj).value / (Object(Character.Inv(CurInvObj).Object).MaxDur * 10)) * Object(Character.Inv(CurInvObj).Object).SellPrice)
                    If A >= 0 Then
                        GetSellPrice = A
                    Else
                        GetSellPrice = 0
                    End If
                Else
                    GetSellPrice = 0
                End If
                Exit Function
            Case Else
                GetSellPrice = Object(Character.Inv(CurInvObj).Object).SellPrice
                Exit Function
        End Select
    End If
    
    GetSellPrice = 0
End Function
