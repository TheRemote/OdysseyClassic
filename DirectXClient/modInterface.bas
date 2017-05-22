Attribute VB_Name = "modInterface"
Option Explicit

Public Sub SetObjectInfo(Info As String)
    If frmMain.Visible = False Then Exit Sub

    frmMain.lblObjectInfo.Caption = Info
    frmMain.lblObjectInfoShadow.Caption = Info
End Sub

Public Sub SetLocation(Location As String)
    If frmMain.Visible = False Then Exit Sub

    frmMain.lblLocation.Caption = Location
    frmMain.lblLocationShadow.Caption = Location
End Sub

Public Sub RefreshInventory()
    If frmMain.Visible = False Then Exit Sub

    DrawInventoryBackground
    DrawInventoryItems
    DrawToDC 0, 0, 181, 181, frmMain.picInventory.hDC, InventoryBuffer, 0, 0
    DrawSelection
    RealDrawCurInvObject
End Sub

Sub DrawInventoryBackground()
    If frmMain.Visible = False Then Exit Sub

    InventoryBuffer.BltFast 0, 0, DDSInventory, InventoryRect, DDBLTFAST_WAIT
End Sub

Sub DrawSelection()
    If frmMain.Visible = False Then Exit Sub

    Dim X As Long, Y As Long
    X = 2 + 36 * ((CurInvObj - 1) Mod 5)
    Y = 2 + 36 * Int((CurInvObj - 1) / 5)

    If CurInvObj > 20 Then Y = Y + 1

    BitBlt frmMain.picInventory.hDC, X - 1, Y - 1, 34, 1, 0, 0, 0, WHITENESS
    BitBlt frmMain.picInventory.hDC, X - 1, Y + 33, 34, 1, 0, 0, 0, WHITENESS
    BitBlt frmMain.picInventory.hDC, X - 1, Y - 1, 1, 34, 0, 0, 0, WHITENESS
    BitBlt frmMain.picInventory.hDC, X + 33, Y - 1, 1, 34, 0, 0, 0, WHITENESS

    frmMain.picInventory.Refresh
End Sub

Sub DrawInventoryItems()
    If frmMain.Visible = False Then Exit Sub

    Dim A As Long

    For A = 1 To MaxInvObjects
        RealDrawInvObject A
    Next A

    For A = 1 To 5
        RealDrawEquippedObject A
    Next A
End Sub

Private Sub RealDrawInvObject(InvNum As Long)
    If frmMain.Visible = False Then Exit Sub

    Dim A As Long, X As Long, Y As Long

    X = 2 + 36 * ((InvNum - 1) Mod 5)
    Y = 2 + 36 * Int((InvNum - 1) / 5)

    If InvNum <= 20 Then
        With Character.Inv(InvNum)
            If .Object > 0 Then
                A = Object(.Object).Picture
                SrcRect.Left = 0
                SrcRect.Top = CLng(A - 1) * 32
                SrcRect.Right = 32
                SrcRect.Bottom = SrcRect.Top + 32
                If A > 0 Then
                    If .EquippedNum > 0 Then
                        FillRect X + 4, Y + 4, 24, 24, InventoryBuffer, RGB(0, 255, 255)
                    End If
                    InventoryBuffer.BltFast X, Y, DDSObjects, SrcRect, DDBLTFAST_SRCCOLORKEY
                End If
            End If
        End With
    End If
End Sub

Private Sub RealDrawEquippedObject(InvNum As Long)
    If frmMain.Visible = False Then Exit Sub

    Dim A As Long, X As Long, Y As Long

    X = 2 + 36 * ((InvNum - 1) Mod 5)
    Y = 4 + 36 * Int((InvNum + 20 - 1) / 5)

    With Character.EquippedObject(InvNum)
        If .Object > 0 Then
            A = Object(.Object).Picture
            If A > 0 Then
                SrcRect.Left = 0
                SrcRect.Top = CLng(A - 1) * 32
                SrcRect.Right = 32
                SrcRect.Bottom = SrcRect.Top + 32
                If A > 0 Then
                    InventoryBuffer.BltFast X, Y, DDSObjects, SrcRect, DDBLTFAST_SRCCOLORKEY
                End If
            End If
        End If
    End With
End Sub

Private Sub RealDrawCurInvObject()
    If frmMain.Visible = False Then Exit Sub

    Dim St1 As String, TheObj As Byte

    BitBlt frmMain.picObject.hDC, 0, 0, 32, 32, 0, 0, 0, BLACKNESS

    If CurInvObj > 0 Then
        If CurInvObj <= 20 Then
            If Character.Inv(CurInvObj).Object > 0 Then
                DrawToDC 0, 0, 32, 32, frmMain.picObject.hDC, DDSObjects, 0, (Object(Character.Inv(CurInvObj).Object).Picture - 1) * 32
                frmMain.lblCurObj = Object(Character.Inv(CurInvObj).Object).name

                'First line (the name line)
                If Character.Inv(CurInvObj).ItemPrefix > 0 Then
                    If Len(ItemPrefix(Character.Inv(CurInvObj).ItemPrefix).name) > 0 Then
                        St1 = ItemPrefix(Character.Inv(CurInvObj).ItemPrefix).name + " " + Object(Character.Inv(CurInvObj).Object).name
                    Else
                        St1 = Object(Character.Inv(CurInvObj).Object).name
                    End If

                    If Character.Inv(CurInvObj).ItemSuffix > 0 Then
                        If Len(ItemSuffix(Character.Inv(CurInvObj).ItemSuffix).name) > 0 Then
                            St1 = St1 + " " + ItemSuffix(Character.Inv(CurInvObj).ItemSuffix).name + vbCrLf
                        Else
                            St1 = St1 + vbCrLf
                        End If
                    Else
                        St1 = St1 + vbCrLf
                    End If
                Else
                    St1 = Object(Character.Inv(CurInvObj).Object).name
                    If Character.Inv(CurInvObj).ItemSuffix > 0 Then
                        If Len(ItemSuffix(Character.Inv(CurInvObj).ItemSuffix).name) > 0 Then
                            St1 = St1 + " " + ItemSuffix(Character.Inv(CurInvObj).ItemSuffix).name + vbCrLf
                        Else
                            St1 = St1 + vbCrLf
                        End If
                    Else
                        St1 = St1 + vbCrLf
                    End If
                End If

                'Second line (the bonus line)
                If Character.Inv(CurInvObj).ItemPrefix > 0 Then
                    St1 = St1 + "Bonus (+"
                    Select Case ItemPrefix(Character.Inv(CurInvObj).ItemPrefix).ModificationType
                    Case 8    'Max HP
                        St1 = St1 + CStr(ItemPrefix(Character.Inv(CurInvObj).ItemPrefix).ModificationValue) + " HP"
                    Case 9    'Max Energy
                        St1 = St1 + CStr(ItemPrefix(Character.Inv(CurInvObj).ItemPrefix).ModificationValue) + " Energy"
                    Case 10    'Max Mana
                        St1 = St1 + CStr(ItemPrefix(Character.Inv(CurInvObj).ItemPrefix).ModificationValue) + " Mana"
                    Case 11    'Damage
                        St1 = St1 + CStr(ItemPrefix(Character.Inv(CurInvObj).ItemPrefix).ModificationValue) + " Damage"
                    Case 12    'Defense
                        St1 = St1 + CStr(ItemPrefix(Character.Inv(CurInvObj).ItemPrefix).ModificationValue) + " Defense"
                    Case 13    'Magic Defense
                        St1 = St1 + CStr(ItemPrefix(Character.Inv(CurInvObj).ItemPrefix).ModificationValue) + " Magic Defense"
                    End Select
                    If Character.Inv(CurInvObj).ItemSuffix > 0 Then
                        St1 = St1 + ", +"
                        Select Case ItemSuffix(Character.Inv(CurInvObj).ItemSuffix).ModificationType
                        Case 8    'Max HP
                            St1 = St1 + CStr(ItemSuffix(Character.Inv(CurInvObj).ItemSuffix).ModificationValue) + " HP"
                        Case 9    'Max Energy
                            St1 = St1 + CStr(ItemSuffix(Character.Inv(CurInvObj).ItemSuffix).ModificationValue) + " Energy"
                        Case 10    'Max Mana
                            St1 = St1 + CStr(ItemSuffix(Character.Inv(CurInvObj).ItemSuffix).ModificationValue) + " Mana"
                        Case 11    'Damage
                            St1 = St1 + CStr(ItemSuffix(Character.Inv(CurInvObj).ItemSuffix).ModificationValue) + " Damage"
                        Case 12    'Defense
                            St1 = St1 + CStr(ItemSuffix(Character.Inv(CurInvObj).ItemSuffix).ModificationValue) + " Defense"
                        Case 13    'Magic Defense
                            St1 = St1 + CStr(ItemSuffix(Character.Inv(CurInvObj).ItemSuffix).ModificationValue) + " Magic Defense"
                        End Select
                        St1 = St1 + ")" + vbCrLf
                    Else
                        St1 = St1 + ")" + vbCrLf
                    End If
                Else
                    If Character.Inv(CurInvObj).ItemSuffix > 0 Then
                        St1 = St1 + "Bonus (+"
                        Select Case ItemSuffix(Character.Inv(CurInvObj).ItemSuffix).ModificationType
                        Case 8    'Max HP
                            St1 = St1 + CStr(ItemSuffix(Character.Inv(CurInvObj).ItemSuffix).ModificationValue) + " HP"
                        Case 9    'Max Energy
                            St1 = St1 + CStr(ItemSuffix(Character.Inv(CurInvObj).ItemSuffix).ModificationValue) + " Energy"
                        Case 10    'Max Mana
                            St1 = St1 + CStr(ItemSuffix(Character.Inv(CurInvObj).ItemSuffix).ModificationValue) + " Mana"
                        Case 11    'Damage
                            St1 = St1 + CStr(ItemSuffix(Character.Inv(CurInvObj).ItemSuffix).ModificationValue) + " Damage"
                        Case 12    'Defense
                            St1 = St1 + CStr(ItemSuffix(Character.Inv(CurInvObj).ItemSuffix).ModificationValue) + " Defense"
                        Case 13    'Magic Defense
                            St1 = St1 + CStr(ItemSuffix(Character.Inv(CurInvObj).ItemSuffix).ModificationValue) + " Magic Defense"
                        End Select
                        St1 = St1 + ")" + vbCrLf
                    Else

                    End If
                End If

                Select Case Object(Character.Inv(CurInvObj).Object).Type
                Case 6  'Money
                    St1 = St1 + "[" + CStr(Character.Inv(CurInvObj).value) + "]"
                Case 11    'Ammo
                    St1 = St1 + "Ammunition" & vbCrLf & "[" + CStr(Character.Inv(CurInvObj).value) + "]" & vbCrLf & "+" & Object(Character.Inv(CurInvObj).Object).Modifier & " Damage" & vbCrLf & GenerateRequirements(Character.Inv(CurInvObj).Object)
                Case 1  'Weapon
                    St1 = St1 + "Weapon (+" & Object(Character.Inv(CurInvObj).Object).Modifier & " Damage)" & vbCrLf & "Condition: " & DurString(CurInvObj) & vbCrLf & GenerateRequirements(Character.Inv(CurInvObj).Object)
                Case 10    'Projectile Weapon
                    St1 = St1 + "Projectile Weapon (+" & Object(Character.Inv(CurInvObj).Object).Modifier & " Damage)" & vbCrLf & GenerateRequirements(Character.Inv(CurInvObj).Object)
                Case 2, 3, 4    'Shield, Helm, Armor
                    If Object(Character.Inv(CurInvObj).Object).Type = 3 Then
                        St1 = St1 + "Armor (+" & Object(Character.Inv(CurInvObj).Object).Modifier & " Defense, +" & Object(Character.Inv(CurInvObj).Object).Data2 & " Magic Defense)"
                    ElseIf Object(Character.Inv(CurInvObj).Object).Type = 4 Then
                        St1 = St1 + "Helm (+" & Object(Character.Inv(CurInvObj).Object).Modifier & " Defense, +" & Object(Character.Inv(CurInvObj).Object).Data2 & " Magic Defense)"
                    ElseIf Object(Character.Inv(CurInvObj).Object).Type = 2 Then
                        St1 = St1 + "Shield (+" & Object(Character.Inv(CurInvObj).Object).Modifier & " Defense, +" & Object(Character.Inv(CurInvObj).Object).Data2 & " Magic Defense)"
                    End If
                    St1 = St1 & vbCrLf & "Condition: " & DurString(CurInvObj) & vbCrLf & GenerateRequirements(Character.Inv(CurInvObj).Object)
                Case 8  'Ring
                    If Object(Character.Inv(CurInvObj).Object).Data2 = 0 Then
                        St1 = St1 + "(Ring) +" & Object(Character.Inv(CurInvObj).Object).Modifier & " Damage" & vbCrLf & "Condition: " & DurString(CurInvObj) & vbCrLf & GenerateRequirements(Character.Inv(CurInvObj).Object)
                    Else
                        St1 = St1 + "(Ring) +" & Object(Character.Inv(CurInvObj).Object).Modifier & " Defense" & vbCrLf & "Condition: " & DurString(CurInvObj) & vbCrLf & GenerateRequirements(Character.Inv(CurInvObj).Object)
                    End If
                End Select
                If ExamineBit(Object(Character.Inv(CurInvObj).Object).flags, 0) = 255 Then St1 = St1 + vbCrLf + "Cannot be repaired"
                If ExamineBit(Object(Character.Inv(CurInvObj).Object).flags, 2) = 255 Then St1 = St1 + vbCrLf + "Does not drop on death"
                If ExamineBit(Object(Character.Inv(CurInvObj).Object).flags, 3) = 255 Then St1 = St1 + vbCrLf + "Two Handed - Cannot use a shield"
                If ExamineBit(Object(Character.Inv(CurInvObj).Object).flags, 6) = 255 Then St1 = St1 + vbCrLf + "Cannot be traded"
                If Object(Character.Inv(CurInvObj).Object).SellPrice > 0 Then St1 = St1 + vbCrLf + "Sells for " + CStr(Object(Character.Inv(CurInvObj).Object).SellPrice) + " gold"
                SetObjectInfo St1
            Else
                frmMain.lblCurObj = vbNullString
                SetObjectInfo vbNullString
            End If
        Else
            TheObj = CurInvObj - 20
            If Character.EquippedObject(TheObj).Object > 0 Then
                frmMain.lblCurObj = Object(Character.EquippedObject(TheObj).Object).name
                DrawToDC 0, 0, 32, 32, frmMain.picObject.hDC, DDSObjects, 0, (Object(Character.EquippedObject(TheObj).Object).Picture - 1) * 32
                If Character.EquippedObject(TheObj).ItemPrefix > 0 Then
                    If Len(ItemPrefix(Character.EquippedObject(TheObj).ItemPrefix).name) > 0 Then
                        St1 = ItemPrefix(Character.EquippedObject(TheObj).ItemPrefix).name + " " + Object(Character.EquippedObject(TheObj).Object).name
                        If Character.EquippedObject(TheObj).ItemSuffix > 0 Then
                            St1 = St1 + " " + ItemSuffix(Character.EquippedObject(TheObj).ItemSuffix).name + vbCrLf
                        Else
                            St1 = St1 + vbCrLf
                        End If
                    Else
                        St1 = Object(Character.EquippedObject(TheObj).Object).name
                        If Character.EquippedObject(TheObj).ItemSuffix > 0 Then
                            If Len(ItemSuffix(Character.EquippedObject(TheObj).ItemSuffix).name) > 0 Then
                                St1 = St1 + " " + ItemSuffix(Character.EquippedObject(TheObj).ItemSuffix).name + vbCrLf
                            Else
                                St1 = St1 + vbCrLf
                            End If
                        Else
                            St1 = St1 + vbCrLf
                        End If
                    End If
                Else
                    St1 = Object(Character.EquippedObject(TheObj).Object).name
                    If Character.EquippedObject(TheObj).ItemSuffix > 0 Then
                        St1 = St1 + " " + ItemSuffix(Character.EquippedObject(TheObj).ItemSuffix).name + vbCrLf
                    Else
                        St1 = St1 + vbCrLf
                    End If
                End If
                If Character.EquippedObject(TheObj).ItemPrefix > 0 Then
                    St1 = St1 + "Bonus (+"
                    Select Case ItemPrefix(Character.EquippedObject(TheObj).ItemPrefix).ModificationType
                    Case 8    'Max HP
                        St1 = St1 + CStr(ItemPrefix(Character.EquippedObject(TheObj).ItemPrefix).ModificationValue) + " HP"
                    Case 9    'Max Energy
                        St1 = St1 + CStr(ItemPrefix(Character.EquippedObject(TheObj).ItemPrefix).ModificationValue) + " Energy"
                    Case 10    'Max Mana
                        St1 = St1 + CStr(ItemPrefix(Character.EquippedObject(TheObj).ItemPrefix).ModificationValue) + " Mana"
                    Case 11    'Damage
                        St1 = St1 + CStr(ItemPrefix(Character.EquippedObject(TheObj).ItemPrefix).ModificationValue) + " Damage"
                    Case 12    'Defense
                        St1 = St1 + CStr(ItemPrefix(Character.EquippedObject(TheObj).ItemPrefix).ModificationValue) + " Defense"
                    Case 13    'Magic Defense
                        St1 = St1 + CStr(ItemPrefix(Character.EquippedObject(TheObj).ItemPrefix).ModificationValue) + " Magic Defense"
                    End Select
                    If Character.EquippedObject(TheObj).ItemSuffix > 0 Then
                        St1 = St1 + ", +"
                        Select Case ItemSuffix(Character.EquippedObject(TheObj).ItemSuffix).ModificationType
                        Case 8    'Max HP
                            St1 = St1 + CStr(ItemSuffix(Character.EquippedObject(TheObj).ItemSuffix).ModificationValue) + " HP"
                        Case 9    'Max Energy
                            St1 = St1 + CStr(ItemSuffix(Character.EquippedObject(TheObj).ItemSuffix).ModificationValue) + " Energy"
                        Case 10    'Max Mana
                            St1 = St1 + CStr(ItemSuffix(Character.EquippedObject(TheObj).ItemSuffix).ModificationValue) + " Mana"
                        Case 11    'Damage
                            St1 = St1 + CStr(ItemSuffix(Character.EquippedObject(TheObj).ItemSuffix).ModificationValue) + " Damage"
                        Case 12    'Defense
                            St1 = St1 + CStr(ItemSuffix(Character.EquippedObject(TheObj).ItemSuffix).ModificationValue) + " Defense"
                        Case 13    'Magic Defense
                            St1 = St1 + CStr(ItemSuffix(Character.EquippedObject(TheObj).ItemSuffix).ModificationValue) + " Magic Defense"
                        End Select
                        St1 = St1 + ")" + vbCrLf
                    Else
                        St1 = St1 + ")" + vbCrLf
                    End If
                Else
                    If Character.EquippedObject(TheObj).ItemSuffix > 0 Then
                        St1 = St1 + "Bonus (+"
                        Select Case ItemSuffix(Character.EquippedObject(TheObj).ItemSuffix).ModificationType
                        Case 8    'Max HP
                            St1 = St1 + CStr(ItemSuffix(Character.EquippedObject(TheObj).ItemSuffix).ModificationValue) + " HP"
                        Case 9    'Max Energy
                            St1 = St1 + CStr(ItemSuffix(Character.EquippedObject(TheObj).ItemSuffix).ModificationValue) + " Energy"
                        Case 10    'Max Mana
                            St1 = St1 + CStr(ItemSuffix(Character.EquippedObject(TheObj).ItemSuffix).ModificationValue) + " Mana"
                        Case 11    'Damage
                            St1 = St1 + CStr(ItemSuffix(Character.EquippedObject(TheObj).ItemSuffix).ModificationValue) + " Damage"
                        Case 12    'Defense
                            St1 = St1 + CStr(ItemSuffix(Character.EquippedObject(TheObj).ItemSuffix).ModificationValue) + " Defense"
                        Case 13    'Magic Defense
                            St1 = St1 + CStr(ItemSuffix(Character.EquippedObject(TheObj).ItemSuffix).ModificationValue) + " Magic Defense"
                        End Select
                        St1 = St1 + ")" + vbCrLf
                    Else

                    End If
                End If
                Select Case Object(Character.EquippedObject(TheObj).Object).Type
                Case 6    'Money
                    St1 = St1 + "[" + CStr(Character.EquippedObject(TheObj).value) + "]"
                Case 11    'Ammo
                    St1 = St1 + "Ammunition" & vbCrLf & "[" + CStr(Character.EquippedObject(TheObj).value) + "]" & vbCrLf & "+" & Object(Character.EquippedObject(TheObj).Object).Modifier & " Damage"
                Case 1    'Weapon
                    St1 = St1 + "Weapon (+" & Object(Character.EquippedObject(TheObj).Object).Modifier & " Damage)" & vbCrLf & "Condition: " & DurString(CurInvObj) & vbCrLf & GenerateRequirements(Character.EquippedObject(TheObj).Object)
                Case 10    'Projectile Weapon
                    St1 = St1 + "Projectile Weapon (+" & Object(Character.EquippedObject(TheObj).Object).Modifier & " Damage)" & vbCrLf & GenerateRequirements(Character.EquippedObject(TheObj).Object)
                Case 2, 3, 4    'Shield, Helm, Armor
                    If Object(Character.EquippedObject(TheObj).Object).Type = 3 Then
                        St1 = St1 + "Armor (+" & Object(Character.EquippedObject(TheObj).Object).Modifier & " Defense, +" & Object(Character.EquippedObject(TheObj).Object).Data2 & " Magic Defense)"
                    ElseIf Object(Character.EquippedObject(TheObj).Object).Type = 4 Then
                        St1 = St1 + "Helm (+" & Object(Character.EquippedObject(TheObj).Object).Modifier & " Defense, +" & Object(Character.EquippedObject(TheObj).Object).Data2 & " Magic Defense)"
                    ElseIf Object(Character.EquippedObject(TheObj).Object).Type = 2 Then
                        St1 = St1 + "Shield (+" & Object(Character.EquippedObject(TheObj).Object).Modifier & " Defense, +" & Object(Character.EquippedObject(TheObj).Object).Data2 & " Magic Defense)"
                    End If
                    St1 = St1 & vbCrLf & "Condition: " & DurString(CurInvObj) & vbCrLf & GenerateRequirements(Character.EquippedObject(TheObj).Object)
                Case 8    'Ring
                    If Object(Character.EquippedObject(TheObj).Object).Data2 = 0 Then
                        St1 = St1 + "Ring (+" & Object(Character.EquippedObject(TheObj).Object).Modifier & " Damage)" & vbCrLf & "Condition: " & DurString(CurInvObj) & vbCrLf & GenerateRequirements(Character.EquippedObject(TheObj).Object)
                    ElseIf Object(Character.EquippedObject(TheObj).Object).Data2 = 1 Then
                        St1 = St1 + "Ring (+" & Object(Character.EquippedObject(TheObj).Object).Modifier & " Defense)" & vbCrLf & "Condition: " & DurString(CurInvObj) & vbCrLf & GenerateRequirements(Character.EquippedObject(TheObj).Object)
                    ElseIf Object(Character.EquippedObject(TheObj).Object).Data2 = 2 Then
                        St1 = St1 + "Ring (+" & Object(Character.EquippedObject(TheObj).Object).Modifier & " Magic Defense)" & vbCrLf & "Condition: " & DurString(CurInvObj) & vbCrLf & GenerateRequirements(Character.EquippedObject(TheObj).Object)
                    End If
                End Select
                If ExamineBit(Object(Character.EquippedObject(TheObj).Object).flags, 0) = 255 Then St1 = St1 + vbCrLf + "Cannot be repaired"
                If ExamineBit(Object(Character.EquippedObject(TheObj).Object).flags, 2) = 255 Then St1 = St1 + vbCrLf + "Does not drop on death"
                If ExamineBit(Object(Character.EquippedObject(TheObj).Object).flags, 3) = 255 Then St1 = St1 + vbCrLf + "Two Handed - Cannot use a shield"
                If ExamineBit(Object(Character.EquippedObject(TheObj).Object).flags, 6) = 255 Then St1 = St1 + vbCrLf + "Cannot be traded"
                If Object(Character.EquippedObject(TheObj).Object).SellPrice > 0 Then St1 = St1 + vbCrLf + "Sells for " + CStr(Object(Character.EquippedObject(TheObj).Object).SellPrice) + " gold"
                SetObjectInfo St1
            Else
                frmMain.lblCurObj = vbNullString
                SetObjectInfo vbNullString
            End If
        End If
    End If
End Sub
