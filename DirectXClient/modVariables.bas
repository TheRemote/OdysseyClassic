Attribute VB_Name = "modVariables"
Option Explicit

'DirectX
Public DX7 As DirectX7

'Server Details
Public CacheDirectory As String
Public ServerDescription As String
Public MOTDText As String

'Winsock
Public PacketOrder As Integer
Public ServerPacketOrder As Integer
Public ClientSocket As Long
Public SocketData As String
Public LastSent As Long

Public Ping As Long
Public PingSent As Long

'Skills
Public LastSkillUse As Long

'Map Editing
Public MapEdit As Boolean, EditMode As Byte
Attribute EditMode.VB_VarUserMemId = 1073741836
Public CurTile As Integer, TopY As Long
Attribute CurTile.VB_VarUserMemId = 1073741838
Attribute TopY.VB_VarUserMemId = 1073741838
Public NewAtt As Integer, CurAtt As Integer, CurAttData(0 To 3) As Byte
Attribute NewAtt.VB_VarUserMemId = 1073741840
Attribute CurAtt.VB_VarUserMemId = 1073741840

'Player Location Variables
Public CX As Byte, CY As Byte, CMap As Long, CMap2 As Long
Attribute CX.VB_VarUserMemId = 1073741843
Attribute CY.VB_VarUserMemId = 1073741843
Attribute CMap.VB_VarUserMemId = 1073741843
Attribute CMap2.VB_VarUserMemId = 1073741843
Public CXO As Long, CYO As Long, CDir As Byte, CWalkStep As Long
Attribute CXO.VB_VarUserMemId = 1073741847
Attribute CYO.VB_VarUserMemId = 1073741847
Attribute CDir.VB_VarUserMemId = 1073741847
Attribute CWalkStep.VB_VarUserMemId = 1073741847
Public CAttack As Long, CWalk As Long
Attribute CAttack.VB_VarUserMemId = 1073741851
Attribute CWalk.VB_VarUserMemId = 1073741851
Public CHP As Long, CEnergy As Long, CMana As Long
Attribute CHP.VB_VarUserMemId = 1073741853
Attribute CEnergy.VB_VarUserMemId = 1073741853
Attribute CMana.VB_VarUserMemId = 1073741853
Public CHPBackup As Long, CEnergyBackup As Long, CManaBackup As Long
Attribute CHPBackup.VB_VarUserMemId = 1073741856
Attribute CEnergyBackup.VB_VarUserMemId = 1073741856
Attribute CManaBackup.VB_VarUserMemId = 1073741856
Public CMaxHP As Long, CMaxEnergy As Long, CMaxMana As Long
Public CMaxHPBackup As Long, CMaxEnergyBackup As Long, CMaxManaBackup As Long

'Login/New Account Variables
Public NewAccount As Boolean
Attribute NewAccount.VB_VarUserMemId = 1073741859
Public User As String, Pass As String
Attribute User.VB_VarUserMemId = 1073741860
Attribute Pass.VB_VarUserMemId = 1073741860

'Form Status Variables
Public frmWait_Loaded As Boolean
Attribute frmWait_Loaded.VB_VarUserMemId = 1073741862
Public frmMain_Loaded As Boolean, frmMain_Showing As Boolean
Attribute frmMain_Loaded.VB_VarUserMemId = 1073741863
Attribute frmMain_Showing.VB_VarUserMemId = 1073741863
Public frmLogin_Loaded As Boolean
Attribute frmLogin_Loaded.VB_VarUserMemId = 1073741865
Public frmAccount_Loaded As Boolean
Attribute frmAccount_Loaded.VB_VarUserMemId = 1073741866
Public frmCharacter_Loaded As Boolean
Attribute frmCharacter_Loaded.VB_VarUserMemId = 1073741867
Public frmNewCharacter_Loaded As Boolean
Attribute frmNewCharacter_Loaded.VB_VarUserMemId = 1073741868
Public frmNewPass_Loaded As Boolean
Attribute frmNewPass_Loaded.VB_VarUserMemId = 1073741869
Public frmEmail_Loaded As Boolean
Public frmMonster_Loaded As Boolean
Attribute frmMonster_Loaded.VB_VarUserMemId = 1073741870
Public frmObject_Loaded As Boolean
Attribute frmObject_Loaded.VB_VarUserMemId = 1073741871
Public frmList_Loaded As Boolean
Attribute frmList_Loaded.VB_VarUserMemId = 1073741872
Public frmMapProperties_Loaded As Boolean
Attribute frmMapProperties_Loaded.VB_VarUserMemId = 1073741873
Public frmGuild_Loaded As Boolean
Attribute frmGuild_Loaded.VB_VarUserMemId = 1073741874
Public frmNPC_Loaded As Boolean
Attribute frmNPC_Loaded.VB_VarUserMemId = 1073741875
Public frmMacros_Loaded As Boolean
Attribute frmMacros_Loaded.VB_VarUserMemId = 1073741876
Public frmOptions_Loaded As Boolean
Attribute frmOptions_Loaded.VB_VarUserMemId = 1073741877
Public frmNewGuild_Loaded As Boolean
Attribute frmNewGuild_Loaded.VB_VarUserMemId = 1073741878
Public frmBan_Loaded As Boolean
Attribute frmBan_Loaded.VB_VarUserMemId = 1073741879
Public frmHall_Loaded As Boolean
Attribute frmHall_Loaded.VB_VarUserMemId = 1073741880
Public frmMagic_Loaded As Boolean
Attribute frmMagic_Loaded.VB_VarUserMemId = 1073741881
Public frmPrefix_Loaded As Boolean
Attribute frmPrefix_Loaded.VB_VarUserMemId = 1073741882
Public frmSuffix_Loaded As Boolean
Attribute frmSuffix_Loaded.VB_VarUserMemId = 1073741883
Public frmSkill_Loaded As Boolean
Attribute frmSkill_Loaded.VB_VarUserMemId = 1073741884
Public frmAbility_Loaded As Boolean
Attribute frmAbility_Loaded.VB_VarUserMemId = 1073741885

'Game State Variables
Public blnEnd As Boolean, blnPlaying As Boolean
Attribute blnEnd.VB_VarUserMemId = 1073741886
Attribute blnPlaying.VB_VarUserMemId = 1073741886
Public Tick As Long
Attribute Tick.VB_VarUserMemId = 1073741888

'Timers
Public AttackTimer As Long
Attribute AttackTimer.VB_VarUserMemId = 1073741889
Public SwitchMapTimer As Long
Attribute SwitchMapTimer.VB_VarUserMemId = 1073741890

'Keyboard State Variables
Public keyUp As Boolean, keyDown As Boolean
Attribute keyUp.VB_VarUserMemId = 1073741891
Attribute keyDown.VB_VarUserMemId = 1073741891
Public keyLeft As Boolean, keyRight As Boolean
Attribute keyLeft.VB_VarUserMemId = 1073741893
Attribute keyRight.VB_VarUserMemId = 1073741893
Public keyCtrl As Boolean, keyShift As Boolean
Attribute keyCtrl.VB_VarUserMemId = 1073741895
Attribute keyShift.VB_VarUserMemId = 1073741895
Public keyAlt As Boolean, keyEscape As Boolean
Attribute keyAlt.VB_VarUserMemId = 1073741897

'Misc Variables
Public InitPath As String
Attribute InitPath.VB_VarUserMemId = 1073741898
Public ChatString As String
Attribute ChatString.VB_VarUserMemId = 1073741899
Public CurInvObj As Long
Attribute CurInvObj.VB_VarUserMemId = 1073741900
Public Freeze As Boolean
Attribute Freeze.VB_VarUserMemId = 1073741901
Public NextTransition As Long
Attribute NextTransition.VB_VarUserMemId = 1073741902
Public CurrentMIDI As Long
Attribute CurrentMIDI.VB_VarUserMemId = 1073741903
Public FrameCounter As Long, AlternateFrameCounter As Long, FrameRate As Long, SecondTimer As Long
Attribute FrameCounter.VB_VarUserMemId = 1073741904
Attribute AlternateFrameCounter.VB_VarUserMemId = 1073741904
Attribute FrameRate.VB_VarUserMemId = 1073741904
Attribute SecondTimer.VB_VarUserMemId = 1073741904
Public SendSpeedHack As Long, SendPing As Long, CurrentSecond As Byte, LastSecond As Byte, SpeedStrikes As Byte
Public CurFrame As Long
Attribute CurFrame.VB_VarUserMemId = 1073741908
Public TempVar1 As Long, TempVar2 As Long, TempVar3 As Long, TempVar4 As Long, TempVar5 As Long, TempVar6 As Long, TempVar7 As Long, TempVar8 As Long, TempVar9 As Long
Attribute TempVar1.VB_VarUserMemId = 1073741909
Attribute TempVar2.VB_VarUserMemId = 1073741909
Attribute TempVar3.VB_VarUserMemId = 1073741909
Attribute TempVar4.VB_VarUserMemId = 1073741909
Attribute TempVar5.VB_VarUserMemId = 1073741909
Attribute TempVar6.VB_VarUserMemId = 1073741909
Attribute TempVar7.VB_VarUserMemId = 1073741909
Attribute TempVar8.VB_VarUserMemId = 1073741909
Attribute TempVar9.VB_VarUserMemId = 1073741909
Public ChatScrollBack As Long
Attribute ChatScrollBack.VB_VarUserMemId = 1073741918
Public LastProjectile As Long
Attribute LastProjectile.VB_VarUserMemId = 1073741919
Public ComputerID As String
Attribute ComputerID.VB_VarUserMemId = 1073741920

Public Section(1 To 30) As String, Suffix As String
Attribute Section.VB_VarUserMemId = 1073741921
Attribute Suffix.VB_VarUserMemId = 1073741921

'Info Text
Public InfoText(0 To 1) As String
Attribute InfoText.VB_VarUserMemId = 1073741923
Public InfoTextTimer As Long
Attribute InfoTextTimer.VB_VarUserMemId = 1073741924
Public FloatText(1 To MaxFloatText) As FloatTextData
Attribute FloatText.VB_VarUserMemId = 1073741925

'Map Data
Public RequestedMap As Boolean
Attribute RequestedMap.VB_VarUserMemId = 1073741926
Public RedrawMap As Boolean
Attribute RedrawMap.VB_VarUserMemId = 1073741927

'WindowProc
Public Const GWL_WNDPROC = -4
Public lpPrevWndProc As Long
Attribute lpPrevWndProc.VB_VarUserMemId = 1073741930
Public gHW As Long
Attribute gHW.VB_VarUserMemId = 1073741931

Public MapData As String * 2677
Public MapDataLoadingArray() As Byte
