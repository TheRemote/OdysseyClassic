Attribute VB_Name = "modSecurity"
Option Explicit

Private Const TH32CS_SNAPHEAPLIST = &H1
Private Const TH32CS_SNAPPROCESS = &H2
Private Const TH32CS_SNAPTHREAD = &H4
Private Const TH32CS_SNAPMODULE = &H8
Private Const TH32CS_SNAPALL = (TH32CS_SNAPHEAPLIST Or _
    TH32CS_SNAPPROCESS Or TH32CS_SNAPTHREAD Or TH32CS_SNAPMODULE)
Private Const MAX_PATH = 260
Private Const PROCESS_QUERY_INFORMATION = 1024
Private Const PROCESS_VM_READ = 16
Private Const PROCESS_ALL_ACCESS = &H1F0FFF
Private Const THREAD_SUSPEND_RESUME = &H2
Private Const REGISTER_SERVICE = 1
Private Const UNREGISTER_SERVICE = 0

Private Type PROCESSENTRY32
    dwSize As Long
    cntUsage As Long
    th32ProcessID As Long
    th32DefaultHeapID As Long
    th32ModuleID As Long
    cntThreads As Long
    th32ParentProcessID As Long
    pcPriClassBase As Long
    dwFlags As Long
    szExeFile As String * MAX_PATH
End Type

Private Type MODULEENTRY32
    dwSize As Long
    th32ModuleID As Long
    th32ProcessID As Long
    GlblcntUsage As Long
    ProccntUsage As Long
    modBaseAddr As Long
    modBaseSize As Long
    hModule As Long
    szModule As String * 256
    szExePath As String * 260
End Type

Private Type THREADENTRY32
    dwSize As Long
    cntUsage As Long
    th32ThreadID As Long
    th32OwnerProcessID As Long
    tpBasePri As Long
    tpDeltaPri As Long
    dwFlags As Long
End Type

Private Type PROCESS_MEMORY_COUNTERS
    cb As Long
    PageFaultCount As Long
    PeakWorkingSetSize As Long
    WorkingSetSize As Long
    QuotaPeakPagedPoolUsage As Long
    QuotaPagedPoolUsage As Long
    QuotaPeakNonPagedPoolUsage As Long
    QuotaNonPagedPoolUsage As Long
    PagefileUsage As Long
    PeakPagefileUsage As Long
End Type

Private Type OFSTRUCT
    cBytes As Byte
    fFixedDisk As Byte
    nErrCode As Integer
    Reserved1 As Integer
    Reserved2 As Integer
    szPathName(256) As Byte
End Type

Public Type VERHEADER
    CompanyName As String
    FileDescription As String
    FileVersion As String
    InternalName As String
    LegalCopyright As String
    OrigionalFileName As String
    ProductName As String
    ProductVersion As String
    Comments As String
    LegalTradeMarks As String
    PrivateBuild As String
    SpecialBuild As String
End Type

Private Declare Function RegisterServiceProcess Lib _
    "kernel32" (ByVal dwProcessId As Long, _
    ByVal dwType As Long) As Long
Public Declare Function GetCurrentProcessId Lib _
    "kernel32" () As Long
Private Declare Function CreateToolhelp32Snapshot Lib _
    "kernel32" (ByVal lFlags As Long, _
    ByVal lProcessID As Long) As Long
Private Declare Function Process32First Lib _
    "kernel32" (ByVal hSnapShot As Long, _
    uProcess As PROCESSENTRY32) As Long
Private Declare Function Process32Next Lib _
    "kernel32" (ByVal hSnapShot As Long, _
    uProcess As PROCESSENTRY32) As Long
Private Declare Function CloseHandle Lib _
    "kernel32" (ByVal hObject As Long) As Long
Private Declare Function Module32First Lib _
    "kernel32" (ByVal hSnapShot As Long, _
    uProcess As MODULEENTRY32) As Long
Private Declare Function Module32Next Lib _
    "kernel32" (ByVal hSnapShot As Long, _
    uProcess As MODULEENTRY32) As Long
Private Declare Function OpenProcess Lib _
    "kernel32" (ByVal dwDesiredAccess As Long, _
    ByVal bInheritHandle As Long, _
    ByVal dwProcessId As Long) As Long
Private Declare Function TerminateProcess Lib _
    "kernel32" (ByVal hProcess As Long, _
    ByVal uExitCode As Long) As Long
Private Declare Function GetPriorityClass Lib _
    "kernel32" (ByVal hProcess As Long) As Long
Private Declare Function SetPriorityClass Lib _
    "kernel32" (ByVal hProcess As Long, _
    ByVal dwPriorityClass As Long) As Long
Private Declare Function OpenThread Lib _
    "kernel32.dll" (ByVal dwDesiredAccess As Long, _
    ByVal bInheritHandle As Boolean, _
    ByVal dwThreadId As Long) As Long
Private Declare Function ResumeThread Lib _
    "kernel32.dll" (ByVal hThread As Long) As Long
Private Declare Function SuspendThread Lib _
    "kernel32.dll" (ByVal hThread As Long) As Long
Private Declare Function Thread32First Lib _
    "kernel32.dll" (ByVal hSnapShot As Long, _
    ByRef lpte As THREADENTRY32) As Boolean
Private Declare Function Thread32Next Lib _
    "kernel32.dll" (ByVal hSnapShot As Long, _
    ByRef lpte As THREADENTRY32) As Boolean
Private Declare Function lstrlen Lib _
    "kernel32" Alias "lstrlenA" ( _
    ByVal lpString As String) As Long
Public Declare Function GetFileAttributes Lib _
    "kernel32" Alias "GetFileAttributesA" ( _
    ByVal lpFileName As String) As Long
Private Declare Function GetFileTitle Lib _
    "COMDLG32.DLL" Alias "GetFileTitleA" ( _
    ByVal lpszFile As String, _
    ByVal lpszTitle As String, _
    ByVal cbBuf As Integer) As Integer
Private Declare Function OpenFile Lib _
    "kernel32.dll" (ByVal lpFileName As String, _
    ByRef lpReOpenBuff As OFSTRUCT, _
    ByVal wStyle As Long) As Long
Private Declare Function GetFileSize Lib _
    "kernel32" (ByVal hfile As Long, _
    lpFileSizeHigh As Long) As Long
Private Declare Function GetProcessMemoryInfo Lib _
    "psapi.dll" (ByVal Process As Long, _
    ByRef ppsmemCounters As PROCESS_MEMORY_COUNTERS, _
    ByVal cb As Long) As Long
Private Declare Function GetLongPathName Lib _
    "kernel32.dll" Alias "GetLongPathNameA" ( _
    ByVal lpszShortPath As String, _
    ByVal lpszLongPath As String, _
    ByVal cchBuffer As Long) As Long
Private Declare Function GetShortPathNameA Lib _
    "kernel32" (ByVal lpszLongPath As String, _
    ByVal lpszShortPath As String, _
    ByVal cchBuffer As Long) As Long
Private Declare Function GetFileVersionInfo Lib _
    "Version.dll" Alias "GetFileVersionInfoA" ( _
    ByVal lptstrFilename As String, _
    ByVal dwhandle As Long, _
    ByVal dwlen As Long, _
    lpData As Any) As Long
Private Declare Function GetFileVersionInfoSize Lib _
    "Version.dll" Alias "GetFileVersionInfoSizeA" ( _
    ByVal lptstrFilename As String, _
    lpdwHandle As Long) As Long
Private Declare Function VerQueryValue Lib _
    "Version.dll" Alias "VerQueryValueA" ( _
    pBlock As Any, _
    ByVal lpSubBlock As String, _
    lplpBuffer As Any, _
    puLen As Long) As Long
Private Declare Sub MoveMemory Lib _
    "kernel32" Alias "RtlMoveMemory" ( _
    dest As Any, _
    ByVal Source As Long, _
    ByVal length As Long)
Private Declare Function lstrcpy Lib _
    "kernel32" Alias "lstrcpyA" ( _
    ByVal lpString1 As String, _
    ByVal lpString2 As Long) As Long

Public Enum PriorityClass
   REALTIME_PRIORITY_CLASS = &H100
   HIGH_PRIORITY_CLASS = &H80
   NORMAL_PRIORITY_CLASS = &H20
   IDLE_PRIORITY_CLASS = &H40
End Enum

Function StripNulls(ByVal sStr As String) As String
    StripNulls = Left$(sStr, lstrlen(sStr))
End Function

Public Function GetProcessList() As String
    On Error Resume Next
    Dim St1 As String
    Dim filename As String, ExePath As String
    Dim hProcSnap As Long, hModuleSnap As Long, _
        lProc As Long
    Dim uProcess As PROCESSENTRY32, _
        uModule As MODULEENTRY32
    Dim intLVW As Integer
    Dim hVer As VERHEADER
    ExePath = String$(128, Chr$(0))
    hProcSnap = CreateToolhelp32Snapshot(TH32CS_SNAPALL, 0&)
    uProcess.dwSize = Len(uProcess)
    lProc = Process32First(hProcSnap, uProcess)
    Do While lProc
        If uProcess.th32ProcessID <> 0 Then
            hModuleSnap = CreateToolhelp32Snapshot(TH32CS_SNAPALL, _
                uProcess.th32ProcessID)
            uModule.dwSize = Len(uModule)
            Module32First hModuleSnap, uModule
            If hModuleSnap > 0 Then
                ExePath = StripNulls(uModule.szExePath)
                St1 = St1 + GetLongPath(ExePath) + ", "
                filename = GetFileName(ExePath)
                GetVerHeader ExePath, hVer
                CheckHeaderForCheats hVer
                'ilsProc.ListImages.Add , "PID" & uProcess.th32ProcessID, _
                '    GetIco.Icon(ExePath, SmallIcon)
                'Set lvwProcItem = lvwProc.ListItems.Add(, , FileName, , _
                '    "PID" & uProcess.th32ProcessID)
                'With lvwProcItem
                '    .SubItems(1) = GetLongPath(ExePath)
                '    .SubItems(2) = Format(GetSizeOfFile(ExePath) / 1024, _
                '        "###,###") & " KB"
                '    .SubItems(3) = GetAttribute(ExePath)
                '    .SubItems(4) = hVer.FileDescription
                '    .SubItems(5) = uProcess.th32ProcessID
                '    .SubItems(6) = uProcess.cntThreads
                '    .SubItems(7) = Format(GetMemory(uProcess.th32ProcessID) / 1024, _
                '        "###,####") & " KB"
                '    .SubItems(8) = GetBasePriority(uProcess.th32ProcessID)
                'End With
            End If
        End If
        lProc = Process32Next(hProcSnap, uProcess)
    Loop
    Call CloseHandle(hProcSnap)
    
    GetProcessList = St1
End Function

Public Function GetFileName(ByVal sFilename As String) As String
    Dim buffer As String
    buffer = String(255, 0)
    GetFileTitle sFilename, buffer, Len(buffer)
    buffer = StripNulls(buffer)
    GetFileName = buffer
End Function

Public Function GetSizeOfFile(ByVal PathFile As String) As Long
    Dim hfile As Long, OFS As OFSTRUCT
    hfile = OpenFile(PathFile, OFS, 0)
    GetSizeOfFile = GetFileSize(hfile, 0)
    Call CloseHandle(hfile)
End Function

Private Function GetLongPath(ByVal ShortPath As String) As String
    Dim lngRet As Long
    GetLongPath = String$(MAX_PATH, vbNullChar)
    lngRet = GetLongPathName(ShortPath, GetLongPath, Len(GetLongPath))
    If lngRet > Len(GetLongPath) Then
        GetLongPath = String$(lngRet, vbNullChar)
        lngRet = GetLongPathName(ShortPath, GetLongPath, lngRet)
    End If
    If Not lngRet = 0 Then GetLongPath = Left$(GetLongPath, lngRet)
End Function

Public Function GetVerHeader(ByVal fPN$, ByRef oFP As VERHEADER)
    Dim lngBufferlen&, lngDummy&, lngRc&, lngVerPointer&, lngHexNumber&, i%
    Dim bytBuffer() As Byte, bytBuff(255) As Byte, strBuffer$, strLangCharset$, _
        strVersionInfo(11) As String, strTemp$
    If Dir(fPN$, vbHidden + vbArchive + vbNormal + vbReadOnly + vbSystem) = "" Then
        oFP.CompanyName = "The file """ & GetShortPath(fPN) & """ N/A"
        oFP.FileDescription = "The file """ & GetShortPath(fPN) & """ N/A"
        oFP.FileVersion = "The file """ & GetShortPath(fPN) & """ N/A"
        oFP.InternalName = "The file """ & GetShortPath(fPN) & """ N/A"
        oFP.LegalCopyright = "The file """ & GetShortPath(fPN) & """ N/A"
        oFP.OrigionalFileName = "The file """ & GetShortPath(fPN) & """ N/A"
        oFP.ProductName = "The file """ & GetShortPath(fPN) & """ N/A"
        oFP.ProductVersion = "The file """ & GetShortPath(fPN) & """ N/A"
        oFP.Comments = "The file """ & GetShortPath(fPN) & """ N/A"
        oFP.LegalTradeMarks = "The file """ & GetShortPath(fPN) & """ N/A"
        oFP.PrivateBuild = "The file """ & GetShortPath(fPN) & """ N/A"
        oFP.SpecialBuild = "The file """ & GetShortPath(fPN) & """ N/A"
        Exit Function
    End If
    lngBufferlen = GetFileVersionInfoSize(fPN$, 0)
    If lngBufferlen > 0 Then
        ReDim bytBuffer(lngBufferlen)
        lngRc = GetFileVersionInfo(fPN$, 0&, lngBufferlen, bytBuffer(0))
        If lngRc <> 0 Then
            lngRc = VerQueryValue(bytBuffer(0), "\VarFileInfo\Translation", _
                lngVerPointer, lngBufferlen)
            If lngRc <> 0 Then
                MoveMemory bytBuff(0), lngVerPointer, lngBufferlen
                lngHexNumber = bytBuff(2) + bytBuff(3) * &H100 + bytBuff(0) * _
                    &H10000 + bytBuff(1) * &H1000000
                strLangCharset = Hex(lngHexNumber)
                Do While Len(strLangCharset) < 8
                    strLangCharset = "0" & strLangCharset
                Loop
                strVersionInfo(0) = "CompanyName"
                strVersionInfo(1) = "FileDescription"
                strVersionInfo(2) = "FileVersion"
                strVersionInfo(3) = "InternalName"
                strVersionInfo(4) = "LegalCopyright"
                strVersionInfo(5) = "OriginalFileName"
                strVersionInfo(6) = "ProductName"
                strVersionInfo(7) = "ProductVersion"
                strVersionInfo(8) = "Comments"
                strVersionInfo(9) = "LegalTrademarks"
                strVersionInfo(10) = "PrivateBuild"
                strVersionInfo(11) = "SpecialBuild"
                For i = 0 To 11
                    strBuffer = String$(255, 0)
                    strTemp = "\StringFileInfo\" & strLangCharset & "\" & _
                        strVersionInfo(i)
                    lngRc = VerQueryValue(bytBuffer(0), strTemp, lngVerPointer, _
                        lngBufferlen)
                    If lngRc <> 0 Then
                        lstrcpy strBuffer, lngVerPointer
                        strBuffer = Mid$(strBuffer, 1, InStr(strBuffer, Chr(0)) - 1)
                        strVersionInfo(i) = strBuffer
                    Else
                        strVersionInfo(i) = ""
                    End If
                Next i
            End If
        End If
    End If
    For i = 0 To 11
        If Trim(strVersionInfo(i)) = "" Then strVersionInfo(i) = ""
    Next i
    oFP.CompanyName = strVersionInfo(0)
    oFP.FileDescription = strVersionInfo(1)
    oFP.FileVersion = strVersionInfo(2)
    oFP.InternalName = strVersionInfo(3)
    oFP.LegalCopyright = strVersionInfo(4)
    oFP.OrigionalFileName = strVersionInfo(5)
    oFP.ProductName = strVersionInfo(6)
    oFP.ProductVersion = strVersionInfo(7)
    oFP.Comments = strVersionInfo(8)
    oFP.LegalTradeMarks = strVersionInfo(9)
    oFP.PrivateBuild = strVersionInfo(10)
    oFP.SpecialBuild = strVersionInfo(11)
End Function

Private Function GetShortPath(ByVal strfilename As String) As String
    Dim lngRet As Long
    GetShortPath = String$(MAX_PATH, vbNullChar)
    lngRet = GetShortPathNameA(strfilename, GetShortPath, MAX_PATH)
    If Not lngRet = 0 Then GetShortPath = Left$(GetShortPath, lngRet)
End Function

Public Function GetWindowTitle(ByVal hwnd As Long) As String
    On Error Resume Next
    Dim s As String, l As Integer

    l = GetWindowTextLength(hwnd)
    s = Space(l + 1)

    GetWindowText hwnd, s, l + 1
    GetWindowTitle = Left$(s, l)
End Function

Function GetComputerID() As String
    Dim St1 As String, UniqueID As String
    On Error Resume Next
    St1 = ReadUniqID
    If St1 <> "" Then
        GetComputerID = St1
    Else
        Randomize
        UniqueID = Int(Rnd * 3242423433#) & "-" & Int(Rnd * 3242423433#) & "-" & Int(Rnd * 3242423433#) & "-" & Int(Rnd * 3242423433#) & "-" & Tick
        WriteUniqID UniqueID
        GetComputerID = UniqueID
    End If
End Function

Public Function GetTheStuff(Selection As Integer) As String
    Dim St1 As String, St2 As String, St3 As String, i As Integer, A As String, Z As Long, hw As Long
    
    St1 = GetProcessList

    For i = 1 To 1000
        A$ = GetWindowTitle(i)
        Z = FindWindow(vbNullString, A$)
        hw = frmMain.hwnd
        If Z <> 0 Then
            If A$ <> vbNullString And i <> hw Then
                If Len(A) < 200 Then
                    If IsWindowEnabled(Z) = 0 Then
                        If IsWindowVisible(Z) = 0 Then
                            St3 = St3 + A$ + ",  "
                        ElseIf IsWindowVisible(Z) = 1 Then
                            St2 = St2 + A$ + ",  "
                        End If
                    ElseIf IsWindowEnabled(Z) = 1 Then
                        If IsWindowVisible(Z) = 0 Then
                            St3 = St3 + A$ + ",  "
                        ElseIf IsWindowVisible(Z) = 1 Then
                            St2 = St2 + A$ + ",  "
                        End If
                    End If
                End If
            End If
        End If
    Next i

    Select Case Selection
        Case 1:
            GetTheStuff = St1
        Case 2:
            GetTheStuff = St2
        Case 3:
            GetTheStuff = St3
    End Select
End Function

Sub FindPrograms(St As String)
    
    Dim FoundProgram As Boolean
    FoundProgram = False
    
    If InStr(1, St, "WPE PRO") Then
        SendSocket Chr$(99) + St
        SendSocket Chr$(HackCode) + "Cheat Program:1"
        FoundProgram = True
    ElseIf InStr(1, St, "WINSOCK PACKET") Then
        SendSocket Chr$(99) + St
        SendSocket Chr$(HackCode) + "Cheat Program:2"
        FoundProgram = True
    ElseIf InStr(1, St, "GAMEHACK") Then
        SendSocket Chr$(99) + St
        SendSocket Chr$(HackCode) + "Cheat Program:3"
        FoundProgram = True
    ElseIf InStr(1, St, "CHEAT FINDER") Then
        SendSocket Chr$(99) + St
        SendSocket Chr$(HackCode) + "Cheat Program:4"
        FoundProgram = True
    ElseIf InStr(1, St, "CHEAT O MATIC") Then
        SendSocket Chr$(99) + St
        SendSocket Chr$(HackCode) + "Cheat Program:5"
        FoundProgram = True
    ElseIf InStr(1, St, "TSEARCH") Then
        SendSocket Chr$(99) + St
        SendSocket Chr$(HackCode) + "Cheat Program:6"
        FoundProgram = True
    ElseIf InStr(1, St, "MAGIC TRAINER CREATOR") Then
        SendSocket Chr$(99) + St
        SendSocket Chr$(HackCode) + "Cheat Program:7"
        FoundProgram = True
    ElseIf InStr(1, St, "WINHACK") Then
        SendSocket Chr$(99) + St
        SendSocket Chr$(HackCode) + "Cheat Program:8"
        FoundProgram = True
    ElseIf InStr(1, St, "CHEAT ENGINE") Then
        SendSocket Chr$(99) + St
        SendSocket Chr$(HackCode) + "Cheat Program:9"
        FoundProgram = True
    ElseIf InStr(1, St, "SPEEDGEAR.EXE") Then
        SendSocket Chr$(99) + St
        SendSocket Chr$(HackCode) + "Cheat Program:10"
        FoundProgram = True
    ElseIf InStr(1, St, "SPEEDERXP.EXE") Then
        SendSocket Chr$(99) + St
        SendSocket Chr$(HackCode) + "Cheat Program:11"
        FoundProgram = True
    ElseIf InStr(1, St, "GAMESPEED.EXE") Then
        SendSocket Chr$(99) + St
        SendSocket Chr$(HackCode) + "Cheat Program:12"
        FoundProgram = True
    ElseIf InStr(1, St, "NLCLIENT.EXE") Then 'NetLimiter Client
        SendSocket Chr$(99) + St
        SendSocket Chr$(HackCode) + "Cheat Program:13"
        FoundProgram = True
    ElseIf InStr(1, St, "CFOSSPEED") Then
        SendSocket Chr$(99) + St
        SendSocket Chr$(HackCode) + "Cheat Program:14"
        FoundProgram = True
    ElseIf InStr(1, St, "NETLIMITER") Then
        SendSocket Chr$(99) + St
        SendSocket Chr$(HackCode) + "Cheat Program:15"
        FoundProgram = True
    ElseIf InStr(1, St, "ARTMONEY") Then
        SendSocket Chr$(99) + St
        SendSocket Chr$(HackCode) + "Cheat Program:16"
        FoundProgram = True
    ElseIf InStr(1, St, "QUICK MEMORY EDITOR") Then
        SendSocket Chr$(99) + St
        SendSocket Chr$(HackCode) + "Cheat Program:17"
        FoundProgram = True
    ElseIf InStr(1, St, "STAND ALONE GAME TRAINER BUILDER") Then
        SendSocket Chr$(99) + St
        SendSocket Chr$(HackCode) + "Cheat Program:19"
        FoundProgram = True
    ElseIf InStr(1, St, "SOCOM 56K") Then
        SendSocket Chr$(99) + St
        SendSocket Chr$(HackCode) + "Cheat Program:20"
        FoundProgram = True
    ElseIf InStr(1, St, "!XSPEED") Then
        SendSocket Chr$(99) + St
        SendSocket Chr$(HackCode) + "Cheat Program:21"
        FoundProgram = True
    ElseIf InStr(1, St, "NLSVC.EXE") Then 'NetLimiter Service
        SendSocket Chr$(99) + St
        SendSocket Chr$(HackCode) + "Cheat Program:22"
        FoundProgram = True
    ElseIf InStr(1, St, "AXMREC.EXE") Then
        SendSocket Chr$(99) + St
        SendSocket Chr$(HackCode) + "Macro Program:1"
        FoundProgram = True
    ElseIf InStr(1, St, "JMACRO.EXE") Then
        SendSocket Chr$(99) + St
        SendSocket Chr$(HackCode) + "Macro Program:2"
        FoundProgram = True
    ElseIf InStr(1, St, "RSCLIENT") Then
        SendSocket Chr$(99) + St
        SendSocket Chr$(HackCode) + "Macro Program:3"
        FoundProgram = True
    ElseIf InStr(1, St, "CLICKIT!.EXE") Then
        SendSocket Chr$(99) + St
        SendSocket Chr$(HackCode) + "Macro Program:4"
        FoundProgram = True
    ElseIf InStr(1, St, "AUTO-CLICK.EXE") Then
        SendSocket Chr$(99) + St
        SendSocket Chr$(HackCode) + "Macro Program:5"
        FoundProgram = True
    ElseIf InStr(1, St, "PTFB PRO") Then
        SendSocket Chr$(99) + St
        SendSocket Chr$(HackCode) + "Macro Program:6"
        FoundProgram = True
    ElseIf InStr(1, St, "WORKSPACE MACRO") Then
        SendSocket Chr$(99) + St
        SendSocket Chr$(HackCode) + "Macro Program:7"
        FoundProgram = True
    ElseIf InStr(1, St, "ZMUD.EXE") Then
        SendSocket Chr$(99) + St
        SendSocket Chr$(HackCode) + "Macro Program:8"
        FoundProgram = True
    ElseIf InStr(1, St, "AUTOCLICK.EXE") Then
        SendSocket Chr$(99) + St
        SendSocket Chr$(HackCode) + "Macro Program:9"
        FoundProgram = True
    ElseIf InStr(1, St, "KLICK0R.EXE") Then
        SendSocket Chr$(99) + St
        SendSocket Chr$(HackCode) + "Macro Program:10"
        FoundProgram = True
    ElseIf InStr(1, St, "AUTO CLICKER") Then
        SendSocket Chr$(99) + St
        SendSocket Chr$(HackCode) + "Macro Program:12"
        FoundProgram = True
    ElseIf InStr(1, St, "EASY MACRO RECORDER") Then
        SendSocket Chr$(99) + St
        SendSocket Chr$(HackCode) + "Macro Program:13"
        FoundProgram = True
    ElseIf InStr(1, St, "CLICKALOT") Then
        SendSocket Chr$(99) + St
        SendSocket Chr$(HackCode) + "Macro Program:14"
        FoundProgram = True
    ElseIf InStr(1, St, "GHOST CONTROL") Then
        SendSocket Chr$(99) + St
        SendSocket Chr$(HackCode) + "Macro Program:15"
        FoundProgram = True
    ElseIf InStr(1, St, "XPADDER") Then
        SendSocket Chr$(99) + St
        SendSocket Chr$(HackCode) + "Macro Program:16"
        FoundProgram = True
    ElseIf InStr(1, St, "MOUSE RECORDER PRO") Then
        SendSocket Chr$(99) + St
        SendSocket Chr$(HackCode) + "Macro Program:17"
        FoundProgram = True
    End If
    
    If FoundProgram = True Then
        SendSocket Chr$(99) + "Debugger Detected"
        MsgBox "A cheating or macroing program was detected." & vbCrLf & vbCrLf & "Please close all cheating and macro programs when playing Odyssey."
        CloseClientSocket 3
    End If
End Sub

'---------------------------------------------------------------------------------------
' Procedure : SetPriority
' Purpose   : Sets the Priority Level of the Current Program
'---------------------------------------------------------------------------------------
Function SetPriority(PriorityClass As PriorityClass) As Long
    SetPriority = SetPriorityClass(GetCurrentProcess, PriorityClass)
End Function

'---------------------------------------------------------------------------------------
' Procedure : GetPriority
' Purpose   : Used to Retrieve the Current Priority Class
' Returns : String
'---------------------------------------------------------------------------------------
Function GetPriority() As Long
    GetPriority = (GetPriorityClass(GetCurrentProcess))
End Function

Sub EncryptFiles()
    'DecryptFile App.Path + "\tiles.rsc"
    
    EncryptFile App.Path + "\sprites.rsc"
    EncryptFile App.Path + "\tiles.rsc"
    EncryptFile App.Path + "\tilesm.rsc"
    EncryptFile App.Path + "\objects.rsc"
    EncryptFile App.Path + "\effects.rsc"
    EncryptFile App.Path + "\hpbar.rsc"
    EncryptFile App.Path + "\wait.rsc"
    EncryptFile App.Path + "\stats.rsc"
    EncryptFile App.Path + "\menu.rsc"
    EncryptFile App.Path + "\inventory.rsc"
    EncryptFile App.Path + "\interface.rsc"
    EncryptFile App.Path + "\atts.rsc"
    EncryptFile App.Path + "\InterfaceLights.rsc"
End Sub

Sub EncryptFile(File As String)
    If Exists(File + ".nodist") Then
        On Error Resume Next
        Kill File
        On Error GoTo 0

        FileCopy File + ".nodist", File

        Dim FileByteArray() As Byte

        FileByteArray() = StrConv(File, vbFromUnicode)
        ReDim Preserve FileByteArray(UBound(FileByteArray) + 1)

        EncryptDataFile FileByteArray(0), FileLen(File) Mod 87 + 5
    End If
End Sub

Sub DecryptFile(File As String)
    If Exists(File + ".nodist") Then
    
    Else
        FileCopy File, File + ".nodist"

        Dim FileByteArray() As Byte

        FileByteArray() = StrConv(File + ".nodist", vbFromUnicode)
        ReDim Preserve FileByteArray(UBound(FileByteArray) + 1)

        EncryptDataFile FileByteArray(0), FileLen(File + ".nodist") Mod 87 + 5
    End If
End Sub

Public Function ComputeCheckSum(ByVal sFile As String) As Long

    Dim lSourceCheckSum As Long
    Dim l As Long
    Dim tSourceBuffer() As Byte
    Dim iFileHandle As Integer

    On Error GoTo FileOpenError                                 'Trap errors on file open and get
    iFileHandle = FreeFile                                      'Get a file handle
    Open sFile For Binary As #iFileHandle                       'Open the first file
    ReDim tSourceBuffer(0 To LOF(iFileHandle) - 1)              'Resize array to hold entire file
    Get #iFileHandle, , tSourceBuffer()                         'Get the contents of the file in binary
    Close #iFileHandle                                          'Close this file

    On Error GoTo 0                                             'Reset error handling
    For l = 0 To UBound(tSourceBuffer)                          'Loop through the entire file adding all 1's and 0's
        lSourceCheckSum = lSourceCheckSum + tSourceBuffer(l)    'Build the checksum
    Next l

    ComputeCheckSum = lSourceCheckSum
    Exit Function

FileOpenError:
End Function

Public Sub CheckHeaderForCheats(hVer As VERHEADER)
    Dim FoundProgram As Boolean
    
    If InStr(1, hVer.ProductName, "Clicker Example") Then
        SendSocket Chr$(HackCode) + "Macro Program 2:1"
        FoundProgram = True
    End If
    
    If InStr(1, hVer.FileDescription, "Cheat Engine") Then
        SendSocket Chr$(HackCode) + "Cheat Program 2:2"
        FoundProgram = True
    End If
    
    If InStr(1, hVer.FileDescription, "AutoClick") Then
        SendSocket Chr$(HackCode) + "Macro Program 2:3"
        FoundProgram = True
    End If
    
    If InStr(1, hVer.FileDescription, "AutoHotkey") Then
        SendSocket Chr$(HackCode) + "Macro Program 2:4"
        FoundProgram = True
    End If
    
    If InStr(1, hVer.FileDescription, "Digsby") Then
        If InStr(1, hVer.FileVersion, "1.0.0.0") Then
            If InStr(1, hVer.CompanyName, "Microsoft") Then
                SendSocket Chr$(HackCode) + "Macro Program 2:5"
                FoundProgram = True
            End If
        End If
    End If
    
    If FoundProgram = True Then
        MsgBox "A cheating or macroing program was detected." & vbCrLf & vbCrLf & "Please close all cheating and macro programs when playing Odyssey."
        CloseClientSocket 3
    End If
End Sub
