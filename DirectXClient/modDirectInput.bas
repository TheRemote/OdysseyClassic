Attribute VB_Name = "modDirectInput"
Option Explicit

Public DI As DirectInput
Public DevKeyboard As DirectInputDevice
Public DevMouse As DirectInputDevice

'Mouse Variables
Dim LastLeftClick As Long
Dim LastLeftRelease As Long
Public LeftIsDown As Boolean
Dim LastRightClick As Long
Dim LastRightRelease As Long
Public RightIsDown As Boolean

'Key Constants
Global Const DIK_ESCAPE = 1
Global Const DIK_UP = 200
Global Const DIK_LEFT = 203
Global Const DIK_RIGHT = 205
Global Const DIK_DOWN = 208
Global Const DIK_1 = 2
Global Const DIK_2 = 3
Global Const DIK_3 = 4
Global Const DIK_4 = 5
Global Const DIK_5 = 6
Global Const DIK_6 = 7
Global Const DIK_7 = 8
Global Const DIK_8 = 9
Global Const DIK_9 = 10
Global Const DIK_0 = 11
Global Const DIK_MINUS = 12
Global Const DIK_EQUALS = 13
Global Const DIK_BACKSPACE = 14
Global Const DIK_TAB = 15
Global Const DIK_Q = 16
Global Const DIK_W = 17
Global Const DIK_E = 18
Global Const DIK_R = 19
Global Const DIK_T = 20
Global Const DIK_Y = 21
Global Const DIK_U = 22
Global Const DIK_I = 23
Global Const DIK_O = 24
Global Const DIK_P = 25
Global Const DIK_LBRACKET = 26
Global Const DIK_RBRACKET = 27
Global Const DIK_RETURN = 28
Global Const DIK_LCONTROL = 29
Global Const DIK_A = 30
Global Const DIK_S = 31
Global Const DIK_D = 32
Global Const DIK_F = 33
Global Const DIK_G = 34
Global Const DIK_H = 35
Global Const DIK_J = 36
Global Const DIK_K = 37
Global Const DIK_L = 38
Global Const DIK_SEMICOLON = 39
Global Const DIK_APOSTROPHE = 40
Global Const DIK_GRAVE = 41
Global Const DIK_LSHIFT = 42
Global Const DIK_BACKSLASH = 43
Global Const DIK_Z = 44
Global Const DIK_X = 45
Global Const DIK_C = 46
Global Const DIK_V = 47
Global Const DIK_B = 48
Global Const DIK_N = 49
Global Const DIK_M = 50
Global Const DIK_COMMA = 51
Global Const DIK_PERIOD = 52
Global Const DIK_SLASH = 53
Global Const DIK_RSHIFT = 54
Global Const DIK_MULTIPLY = 55
Global Const DIK_LALT = 56
Global Const DIK_SPACE = 57
Global Const DIK_CAPSLOCK = 58
Global Const DIK_F1 = 59
Global Const DIK_F2 = 60
Global Const DIK_F3 = 61
Global Const DIK_F4 = 62
Global Const DIK_F5 = 63
Global Const DIK_F6 = 64
Global Const DIK_F7 = 65
Global Const DIK_F8 = 66
Global Const DIK_F9 = 67
Global Const DIK_F10 = 68
Global Const DIK_NUMLOCK = 69
Global Const DIK_SCROLL = 70
Global Const DIK_NUMPAD7 = 71
Global Const DIK_NUMPAD8 = 72
Global Const DIK_NUMPAD9 = 73
Global Const DIK_SUBTRACT = 74
Global Const DIK_NUMPAD4 = 75
Global Const DIK_NUMPAD5 = 76
Global Const DIK_NUMPAD6 = 77
Global Const DIK_ADD = 78
Global Const DIK_NUMPAD1 = 79
Global Const DIK_NUMPAD2 = 80
Global Const DIK_NUMPAD3 = 81
Global Const DIK_NUMPAD0 = 82
Global Const DIK_DECIMAL = 83
Global Const DIK_F11 = 87
Global Const DIK_F12 = 88
Global Const DIK_NUMPADENTER = 156
Global Const DIK_RCONTROL = 157
Global Const DIK_DIVIDE = 181
Global Const DIK_RALT = 184
Global Const DIK_HOME = 199
Global Const DIK_PAGEUP = 201
Global Const DIK_END = 207
Global Const DIK_PAGEDOWN = 209
Global Const DIK_INSERT = 210
Global Const DIK_DELETE = 211

Public Function GetKeyboardDevice(ByVal Window As Form) As DirectInputDevice
    Set GetKeyboardDevice = DI.CreateDevice("GUID_SysKeyboard")
    GetKeyboardDevice.SetCommonDataFormat DIFORMAT_KEYBOARD
End Function

Public Function GetMouseDevice(ByVal Window As Form) As DirectInputDevice
    Set GetMouseDevice = DI.CreateDevice("GUID_SysMouse")
    GetMouseDevice.SetCommonDataFormat DIFORMAT_MOUSE
End Function

Public Sub ReacquireMouse()
    On Error Resume Next
    DevMouse.SetCooperativeLevel frmMain.hwnd, DISCL_BACKGROUND Or DISCL_EXCLUSIVE
    DevMouse.Acquire
End Sub

Public Sub ReacquireKeyboard()
    On Error Resume Next
    DevKeyboard.SetCooperativeLevel frmMain.hwnd, DISCL_FOREGROUND Or DISCL_NONEXCLUSIVE
    DevKeyboard.Acquire
End Sub

Public Function GetKeyboardState() As Boolean()
    On Error GoTo Reacquire
    
    Dim TempKeyboardState As DIKEYBOARDSTATE
    Dim i As Integer
    Dim Keys(1 To 211) As Boolean
    DevKeyboard.GetDeviceStateKeyboard TempKeyboardState
    For i = 1 To 211
        If TempKeyboardState.key(i) <> 0 Then
            Keys(i) = True
        Else
            Keys(i) = False
        End If
    Next i
    Let GetKeyboardState = Keys
    
    Exit Function
    
Reacquire:
    ReacquireKeyboard
    
    Let GetKeyboardState = Keys
End Function

Public Function GetMouseState() As Boolean()
    On Error GoTo Reacquire
    
    Dim TempMouseState As DIMOUSESTATE
    Dim i As Integer
    Dim Buttons(0 To 2) As Boolean
    
    DevMouse.GetDeviceStateMouse TempMouseState
    
    For i = 0 To 2
        If TempMouseState.Buttons(i) <> 0 Then
            Buttons(i) = True
        End If
    Next i
    
    Let GetMouseState = Buttons
    
    Exit Function
    
Reacquire:
    ReacquireMouse
    
    Let GetMouseState = Buttons
End Function

Public Sub InitDirectInput()
    Set DevKeyboard = GetKeyboardDevice(frmMain)
    Set DevMouse = GetMouseDevice(frmMain)
End Sub

Public Sub UnloadDirectInput()
    DevKeyboard.Unacquire
    DevMouse.Unacquire
End Sub

Public Sub CheckDirectInputKeys()
    Dim Keys() As Boolean
    Keys = GetKeyboardState()
    
    keyUp = Keys(DIK_UP)
    keyDown = Keys(DIK_DOWN)
    keyLeft = Keys(DIK_LEFT)
    keyRight = Keys(DIK_RIGHT)
    If Keys(DIK_LALT) = True Or Keys(DIK_RALT) = True Then
        keyAlt = True
    Else
        keyAlt = False
    End If
    If Keys(DIK_LCONTROL) = True Or Keys(DIK_RCONTROL) = True Then
        keyCtrl = True
    Else
        keyCtrl = False
    End If
End Sub

Public Sub CheckDirectInputMouse()
    Dim Buttons() As Boolean
    Buttons = GetMouseState()
    
    If LeftIsDown = True Then
        If Buttons(0) = False Then
            LastLeftRelease = Tick
            LeftIsDown = False
        End If
    Else
        If Buttons(0) = True Then
            If Tick - LastLeftClick < 300 Then
                frmMain.DoubleClick 0
                LastLeftClick = 0
            Else
                frmMain.Click 0
                LastLeftClick = Tick
            End If
            LeftIsDown = True
        End If
    End If
    
    If RightIsDown = True Then
        If Buttons(1) = False Then
            LastRightRelease = Tick
            RightIsDown = False
        End If
    Else
        If Buttons(1) = True Then
            If Tick - LastRightClick < 300 Then
                frmMain.DoubleClick 1
                LastRightClick = 0
            Else
                frmMain.Click 1
                LastRightClick = Tick
            End If
            RightIsDown = True
        End If
    End If
End Sub
