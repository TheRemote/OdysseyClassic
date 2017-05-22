Attribute VB_Name = "modClasses"
Option Explicit

Sub CreateClassData()
    If World.ServerPort = 5752 Then 'PK Server Classes
        With Class(1)    'Knight
            .StartHP = 42
            .StartEnergy = 41
            .StartMana = 100
            .MaxHP = 200
            .MaxEnergy = 120
            .MaxMana = 100
        End With
        With Class(2)    'Mage
            .StartHP = 42
            .StartEnergy = 41
            .StartMana = 100
            .MaxHP = 200
            .MaxEnergy = 120
            .MaxMana = 100
        End With
        With Class(3)    'Rogue
            .StartHP = 42
            .StartEnergy = 41
            .StartMana = 100
            .MaxHP = 200
            .MaxEnergy = 120
            .MaxMana = 100
        End With
        With Class(4)    'Cleric
            .StartHP = 42
            .StartEnergy = 41
            .StartMana = 100
            .MaxHP = 200
            .MaxEnergy = 120
            .MaxMana = 100
        End With
        
        Exit Sub
    End If
    
    If World.ServerPort = 5750 Then 'Main Server Classes
        With Class(1)    'Knight
            .StartHP = 25
            .StartEnergy = 20
            .StartMana = 10
            .MaxHP = 395
            .MaxEnergy = 178
            .MaxMana = 51
        End With
        With Class(2)    'Mage
            .StartHP = 25
            .StartEnergy = 20
            .StartMana = 25
            .MaxHP = 293
            .MaxEnergy = 153
            .MaxMana = 357
        End With
        With Class(3)    'Rogue
            .StartHP = 25
            .StartEnergy = 20
            .StartMana = 10
            .MaxHP = 293
            .MaxEnergy = 204
            .MaxMana = 142
        End With
        With Class(4)    'Cleric
            .StartHP = 25
            .StartEnergy = 20
            .StartMana = 20
            .MaxHP = 255
            .MaxEnergy = 153
            .MaxMana = 267
        End With
        
        Exit Sub
    End If
    
    If World.ServerPort = 5756 Then 'Classic Main Stats
        With Class(1) 'Knight
            .StartHP = 25
            .StartEnergy = 15
            .StartMana = 10
            .MaxHP = 140
            .MaxEnergy = 100
            .MaxMana = 20
        End With
        With Class(2) 'Mage
            .StartHP = 10
            .StartEnergy = 15
            .StartMana = 25
            .MaxHP = 68
            .MaxEnergy = 106
            .MaxMana = 140
        End With
        With Class(3) 'Rogue
            .StartHP = 20
            .StartEnergy = 20
            .StartMana = 10
            .MaxHP = 100
            .MaxEnergy = 134
            .MaxMana = 56
        End With
        With Class(4) 'Cleric
            .StartHP = 15
            .StartEnergy = 15
            .StartMana = 20
            .MaxHP = 95
            .MaxEnergy = 78
            .MaxMana = 105
        End With
        Exit Sub
    End If
    
    With Class(1)    'Knight
        .StartHP = 25
        .StartEnergy = 20
        .StartMana = 10
        .MaxHP = 155
        .MaxEnergy = 70
        .MaxMana = 20
    End With
    With Class(2)    'Mage
        .StartHP = 25
        .StartEnergy = 20
        .StartMana = 25
        .MaxHP = 115
        .MaxEnergy = 60
        .MaxMana = 140
    End With
    With Class(3)    'Rogue
        .StartHP = 25
        .StartEnergy = 20
        .StartMana = 10
        .MaxHP = 115
        .MaxEnergy = 80
        .MaxMana = 56
    End With
    With Class(4)    'Cleric
        .StartHP = 25
        .StartEnergy = 20
        .StartMana = 20
        .MaxHP = 100
        .MaxEnergy = 60
        .MaxMana = 105
    End With
End Sub
