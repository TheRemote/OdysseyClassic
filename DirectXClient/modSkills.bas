Attribute VB_Name = "modSkills"
Option Explicit
Sub ProcessSkillData(St As String)
    Select Case Asc(Mid$(St, 1, 1))
    Case 1    'Used Skill
        ProcessUsedSkill St
    Case 2    'Skill Level Up
        ProcessSkillLevelUp St
    Case 3    'Skill Level
        ProcessSkillLevel St
    End Select
End Sub

Sub ProcessSkillLevel(St As String)
    Dim TheSkill As Byte
    TheSkill = Asc(Mid$(St, 2, 1))
    Character.Skill(TheSkill).Level = Asc(Mid$(St, 3, 1))
    Character.Skill(TheSkill).Experience = Asc(Mid$(St, 4, 1)) * 16777216 + Asc(Mid$(St, 5, 1)) * 65536 + Asc(Mid$(St, 6, 1)) * 256& + Asc(Mid$(St, 7, 1))
    DrawSkillsList
End Sub

Sub ProcessMagicData(St As String)
    Select Case Asc(Mid$(St, 1, 1))
    Case 2    'Skill Level up
        ProcessMagicLevelUp St
    Case 3    'Skill Level
        ProcessMagicLevel St
    End Select
End Sub

Sub ProcessMagicLevel(St As String)
    Dim TheMagic As Integer, TheLevel As Byte
    TheMagic = Int(Asc(Mid$(St, 2, 1)) * 256 + Asc(Mid$(St, 3, 1)))
    TheLevel = Asc(Mid$(St, 4, 1))
    Magic(TheMagic).MagicLevel = TheLevel
    Magic(TheMagic).MagicExperience = Asc(Mid$(St, 5, 1)) * 16777216 + Asc(Mid$(St, 6, 1)) * 65536 + Asc(Mid$(St, 7, 1)) * 256& + Asc(Mid$(St, 8, 1))
    DrawMagicList
End Sub

Sub ProcessUsedSkill(St As String)
    Dim A As Long, B As Long

    Select Case Asc(Mid$(St, 2, 1))
    Case 1    'Fishing
        A = GetInt(Mid$(St, 3, 2))
        If A > 0 Then
            PrintChat "You caught a " + Object(A).name + "!", 15
        Else
            PrintChat "You cast your line into the water, but do not catch anything.", YELLOW
        End If
    Case 2    'Mining
        A = GetInt(Mid$(St, 3, 2))
        If A > 0 Then
            PrintChat "You found some " + Object(A).name + "!", 15
        Else
            PrintChat "You hack away with your pick, but discover nothing.", YELLOW
        End If
    Case 3    'Lumberjacking
        B = Asc(Mid$(St, 3, 1))
        If B > 0 Then
            PrintChat "You chop " + CStr(B) + " lumber!", 15
        Else
            PrintChat "You chop at the tree's trunk but end up with nothing but splinters.", YELLOW
        End If
    End Select
End Sub
Sub ProcessSkillLevelUp(St As String)
    Dim TheSkill As Byte, TheLevel As Byte
    TheSkill = CByte(Asc(Mid$(St, 2, 1)))
    TheLevel = CByte(Asc(Mid$(St, 3, 1)))
    Character.Skill(TheSkill).Level = TheLevel
    Character.Skill(TheSkill).Experience = 0

    Select Case TheSkill
    Case 1    'Fishing
        PrintChat "Your fishing skill has increased to level " + CStr(Character.Skill(1).Level) + "!", 12
    Case 2    'Mining
        PrintChat "Your mining skill has increased to level " + CStr(Character.Skill(2).Level) + "!", 12
    Case 3    'Lumberjacking
        PrintChat "Your lumberjacking skill has increased to level " + CStr(Character.Skill(3).Level) + "!", 12
    Case 4    'Cooking
        PrintChat "Your cooking skill has increased to level " + CStr(Character.Skill(4).Level) + "!", 12
    Case 5    'Enchanting
        PrintChat "Your enchanting skill has increased to level " + CStr(Character.Skill(5).Level) + "!", 12
    Case 6    'Smithing
        PrintChat "Your smithing skill has increased to level " + CStr(Character.Skill(6).Level) + "!", 12
    End Select

    DrawSkillsList
End Sub

Sub ProcessMagicLevelUp(St As String)
    Dim TheMagic As Integer, TheLevel As Byte
    TheMagic = Int(Asc(Mid$(St, 2, 1)) * 256 + Asc(Mid$(St, 3, 1)))
    TheLevel = Asc(Mid$(St, 4, 1))
    Magic(TheMagic).MagicLevel = TheLevel
    Magic(TheMagic).MagicExperience = 0

    PrintChat "Your " + Magic(TheMagic).name + " has increased to level " + CStr(Magic(TheMagic).MagicLevel) + "!", 12

    DrawMagicList
End Sub

Sub DrawSkillsList()
    Dim St As String
    frmMain.picList.Cls
    frmMain.lblSkillType = "Skills"

    Dim A As Long, C As Long, DrawY As Long, Percent As Single
    Dim DontDraw As Boolean
    DontDraw = False

    For A = 1 To MaxSkill
        If A > 3 And Character.Access = 0 Then DontDraw = True Else DontDraw = False

        DrawY = A * 26 - 26
        Select Case A
        Case 1    'Fishing
            St = "Fishing - Level " + CStr(Character.Skill(1).Level)
            If IsHotKeyed(1, 1) Then St = St + " - F" + CStr(ReturnHotKey(1, 1))
        Case 2    'Mining
            St = "Mining - Level " + CStr(Character.Skill(2).Level)
            If IsHotKeyed(2, 1) Then St = St + " - F" + CStr(ReturnHotKey(2, 1))
        Case 3    'Lumberjacking
            St = "Lumberjacking - Level " + CStr(Character.Skill(3).Level)
            If IsHotKeyed(3, 1) Then St = St + " - F" + CStr(ReturnHotKey(3, 1))
        Case 4    'Cooking
            If Character.Access > 0 Then
                St = "Cooking - Level " + CStr(Character.Skill(4).Level)
                If IsHotKeyed(4, 1) Then St = St + " - F" + CStr(ReturnHotKey(4, 1))
            End If
        Case 5    'Enchanting
            If Character.Access > 0 Then
                St = "Enchanting - Level " + CStr(Character.Skill(5).Level)
                If IsHotKeyed(5, 1) Then St = St + " - F" + CStr(ReturnHotKey(5, 1))
            End If
        Case 6    'Smithing
            If Character.Access > 0 Then
                St = "Smithing - Level " + CStr(Character.Skill(6).Level)
                If IsHotKeyed(6, 1) Then St = St + " - F" + CStr(ReturnHotKey(6, 1))
            End If
        End Select

        If DontDraw = False Then

            TextOut frmMain.picList.hDC, 3, 1 + DrawY, St, Len(St)

            'Experience
            C = Int(5 * CLng(Character.Skill(A).Level) ^ 1.3)
            If Character.Skill(A).Level > 0 Then
                Percent = Character.Skill(A).Experience / C
            Else
                Percent = 0
            End If
            St = CStr(Int(Percent * 100)) + "% " + CStr(Character.Skill(A).Experience) + "/" + CStr(C)
            DrawToDC 0, DrawY + 14, 150, 12, frmMain.picList.hDC, DDSStats, 0, 34
            frmMain.picList.Line (150 * Percent, DrawY + 14)-(150, DrawY + 25), 0, BF
            Dim DrawRect As RECT
            DrawRect.Left = 0
            DrawRect.Right = 150
            DrawRect.Top = DrawY + 13
            DrawRect.Bottom = DrawY + 21
            Draw3dText frmMain.picList.hDC, DrawRect, St, frmMain.picList.ForeColor, 1

        End If
    Next A

    frmMain.DrawInterfaceLights
End Sub

Sub UseSkill(ListIndex As Integer)
    If Not Tick > Character.LastMove + 60000 And Character.IsDead = False Then
        If LastSkillUse + 500 < Tick Then
            LastSkillUse = Tick
            Dim CanUse As Boolean
            Select Case ListIndex
            Case 1    'Fishing
                CanUse = False
                Select Case CDir
                Case 0
                    If CY > 0 Then If Map.Tile(CX, CY - 1).Att = 13 Then CanUse = True
                Case 1
                    If CY < 11 Then If Map.Tile(CX, CY + 1).Att = 13 Then CanUse = True
                Case 2
                    If CX > 0 Then If Map.Tile(CX - 1, CY).Att = 13 Then CanUse = True
                Case 3
                    If CX < 11 Then If Map.Tile(CX + 1, CY).Att = 13 Then CanUse = True
                End Select
                If CanUse = True Then
                    If GetEnergy >= 5 Then
                        SendSocket Chr$(78) + Chr$(1)
                        SetEnergy GetEnergy - 5
                        DrawStats
                    Else
                        PrintChat "You are too tired to fish!", 14
                    End If
                Else
                    PrintChat "You cannot fish here!", 14
                End If
            Case 2    'Mining
                CanUse = False
                Select Case CDir
                Case 0
                    If CY > 0 Then If Map.Tile(CX, CY - 1).Att = 14 Then CanUse = True
                Case 1
                    If CY < 11 Then If Map.Tile(CX, CY + 1).Att = 14 Then CanUse = True
                Case 2
                    If CX > 0 Then If Map.Tile(CX - 1, CY).Att = 14 Then CanUse = True
                Case 3
                    If CX < 11 Then If Map.Tile(CX + 1, CY).Att = 14 Then CanUse = True
                End Select
                If CanUse = True Then
                    If GetEnergy >= 5 Then
                        SendSocket Chr$(78) + Chr$(2)
                        SetEnergy GetEnergy - 5
                        DrawStats
                    Else
                        PrintChat "You are too tired to mine!", 14
                    End If
                Else
                    PrintChat "You cannot mine here!", 14
                End If
            Case 3    'Lumberjacking
                CanUse = False
                Select Case CDir
                Case 0
                    If CY > 0 Then If Map.Tile(CX, CY - 1).Att = 16 Then CanUse = True
                Case 1
                    If CY < 11 Then If Map.Tile(CX, CY + 1).Att = 16 Then CanUse = True
                Case 2
                    If CX > 0 Then If Map.Tile(CX - 1, CY).Att = 16 Then CanUse = True
                Case 3
                    If CX < 11 Then If Map.Tile(CX + 1, CY).Att = 16 Then CanUse = True
                End Select
                If CanUse = True Then
                    If GetEnergy >= 5 Then
                        SendSocket Chr$(78) + Chr$(3)
                        SetEnergy GetEnergy - 5
                        DrawStats
                    Else
                        PrintChat "You are too tired to chop!", 14
                    End If
                Else
                    PrintChat "You cannot chop here!", 14
                End If
            Case Else
                SendSocket Chr$(78) + Chr$(ListIndex)
            End Select
        Else
            PrintChat "You are too tired!", 14
        End If
    End If
End Sub

Sub DrawMagicList()
    Dim DrawY As Long, A As Long, B As Long, C As Long, Percent As Single
    Dim St As String
    DrawY = 1

    frmMain.picList.Cls
    frmMain.lblSkillType = "Magic"

    B = -(frmMain.sclMagic.value)
    For A = 1 To MaxMagic
        If ExamineBit(Magic(A).Class, Character.Class - 1) = True And Character.Level >= Magic(A).Level Then
            If Not (Magic(A).MagicLevel = 0 And Character.Access = 0 And ServerPort = 5752) Then
                B = B + 1
                If B > 0 Then
                    St = Magic(A).name + " - Level: " + CStr(Magic(A).MagicLevel)
                    If IsHotKeyed(B + frmMain.sclMagic.value, 3) Then St = St + " - F" + CStr(ReturnHotKey(B + frmMain.sclMagic.value, 3))
                    TextOut frmMain.picList.hDC, 3, DrawY, St, Len(St)
    
                    'Experience
                    C = Int(5 * CLng(Magic(A).MagicLevel) ^ 1.3)
                    If Magic(A).MagicLevel > 0 Then
                        Percent = Magic(A).MagicExperience / C
                    Else
                        Percent = 0
                    End If
                    St = CStr(Int(Percent * 100)) + "% " + CStr(Magic(A).MagicExperience) + "/" + CStr(C)
                    DrawToDC 0, DrawY + 14, 150, 12, frmMain.picList.hDC, DDSStats, 0, 34
                    frmMain.picList.Line (150 * Percent, DrawY + 14)-(150, DrawY + 25), 0, BF
                    Dim DrawRect As RECT
                    DrawRect.Left = 0
                    DrawRect.Right = 150
                    DrawRect.Top = DrawY + 13
                    DrawRect.Bottom = DrawY + 21
                    Draw3dText frmMain.picList.hDC, DrawRect, St, frmMain.picList.ForeColor, 1
    
                    DrawY = DrawY + 26
                End If
            End If
        Else

        End If
    Next A
    
    frmMain.DrawInterfaceLights
End Sub

Sub UseMagic(ListIndex As Integer)
    If Not Tick > Character.LastMove + 60000 And Character.IsDead = False Then
        If Tick >= AttackTimer Then
            Dim A As Long, B As Long
            B = -(frmMain.sclMagic.value)
            For A = 1 To MaxMagic
                If ExamineBit(Magic(A).Class, Character.Class - 1) = True And Character.Level >= Magic(A).Level Then
                    If Not (Magic(A).MagicLevel = 0 And Character.Access = 0 And ServerPort = 5752) Then
                        B = B + 1
                        If B > 0 Then
                            If B = ListIndex Then
                                AttackTimer = Tick + 800
                                SendSocket Chr$(84) + Chr$(A)
                                Exit Sub
                            End If
                        End If
                    End If
                End If
            Next A
        End If
    End If
End Sub
