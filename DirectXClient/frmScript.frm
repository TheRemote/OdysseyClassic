VERSION 5.00
Object = "{871470D6-5AF6-4EE8-9C28-9F67DCB46490}#12.0#0"; "SCIVBX.OCX"
Begin VB.Form frmScript 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "The Odyssey Online Classic [Editing Script]"
   ClientHeight    =   8400
   ClientLeft      =   45
   ClientTop       =   630
   ClientWidth     =   10905
   ControlBox      =   0   'False
   Icon            =   "frmScript.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8400
   ScaleWidth      =   10905
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtCode 
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   9
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5055
      HideSelection   =   0   'False
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   5
      Top             =   120
      Visible         =   0   'False
      Width           =   8175
   End
   Begin SCIVBX.SCIHighlighter SCIHighlighter 
      Left            =   840
      Top             =   720
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin SCIVBX.SCIVB Scintilla 
      Left            =   240
      Top             =   720
      _ExtentX        =   847
      _ExtentY        =   847
      UseTabs         =   -1  'True
   End
   Begin VB.Label btnCancel 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6240
      TabIndex        =   4
      Top             =   7920
      Width           =   1455
   End
   Begin VB.Label btnClear 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7800
      TabIndex        =   3
      Top             =   7920
      Width           =   1455
   End
   Begin VB.Label btnOk 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Ok"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9360
      TabIndex        =   2
      Top             =   7920
      Width           =   1455
   End
   Begin VB.Label lblName 
      Height          =   255
      Left            =   840
      TabIndex        =   1
      Top             =   120
      Width           =   3735
   End
   Begin VB.Label cptName 
      Caption         =   "Name:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   615
   End
   Begin VB.Menu mnuSearch 
      Caption         =   "&Search"
      Begin VB.Menu mnuFind 
         Caption         =   "&Find"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuReplace 
         Caption         =   "&Replace"
         Shortcut        =   ^H
      End
      Begin VB.Menu mnuGoto 
         Caption         =   "&Goto"
         Shortcut        =   ^G
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFindNext 
         Caption         =   "Find &Next"
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnuFindPrevious 
         Caption         =   "Find &Previous"
         Shortcut        =   ^{F3}
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuCopy 
         Caption         =   "&Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuPaste 
         Caption         =   "&Paste"
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuSelectAll 
         Caption         =   "Select All"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuUndo 
         Caption         =   "&Undo"
         Shortcut        =   ^Z
      End
      Begin VB.Menu mnuRedo 
         Caption         =   "&Redo"
         Shortcut        =   ^Y
      End
   End
End
Attribute VB_Name = "frmScript"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub btnClear_Click()
    If MsgBox("This will clear the current script from the server-- are you sure you wish to continue?", vbYesNo, TitleString) = vbYes Then
        SendSocket Chr$(60) + lblName + vbNullChar + vbNullChar
        Unload Me
    End If
End Sub

Private Sub btnOk_Click()
    Dim St As String, A As Long, B As Long
    Dim IncludeLineCount As Long
    Dim Process As Long

    Me.Caption = "The Odyssey Online Classic [Compiling Script]"

    If Exists("MBSC.EXE") Then
        If Exists("MBSC.INC") Then
            txtCode.Text = Scintilla.Text

            If CheckScriptInvalid(txtCode.Text) = True Then
                Exit Sub
            End If

            Open "script.bas" For Output As #1

            Open "mbsc.inc" For Input As #2
            While Not EOF(2)
                Line Input #2, St
                Print #1, St
                IncludeLineCount = IncludeLineCount + 1
            Wend
            Close #2

            Print #1, txtCode.Text

            Close #1

            Open "COMPILE.BAT" For Output As #1
            Print #1, "MBSC SCRIPT.BAS SCRIPT.ASM SCRIPT.BIN SCRIPT.LOG"
            Print #1, "CLOSE.COM"
            Close #1

            Open "CLOSE.COM" For Binary As #1
            Put #1, , Chr$(184) + Chr$(64) + vbNullChar + Chr$(142) + Chr$(216) + Chr$(199) + Chr$(6) + Chr$(114) + vbNullChar + Chr$(52) + Chr$(18) + Chr$(234) + vbNullChar + vbNullChar + Chr$(255) + Chr$(255)
            Close #1

            Process = Shell("COMPILE.BAT", vbHide)
            If Process <> 0 Then
                WaitForTerm Process

                If Exists("SCRIPT.LOG") Then
                    St = vbNullString
                    Open "SCRIPT.LOG" For Input As #1
                    If Not EOF(1) Then Line Input #1, St
                    Close #1
                    If St = vbNullString Then
                        If Exists("SCRIPT.BIN") Then
                            Dim bData As Byte
                            St = vbNullString
                            Open "SCRIPT.BIN" For Binary As #1
                            While Not EOF(1)
                                Get #1, , bData
                                If Not EOF(1) Then St = St + Chr$(bData)
                            Wend
                            Close #1
                            SendSocket Chr$(60) + lblName + vbNullChar + txtCode.Text + vbNullChar + St
                            Unload Me
                        Else
                            MsgBox "An unknown error occurred when assembling script!", vbOKOnly, TitleString
                        End If
                    Else
                        A = InStr(St, ":")
                        If A > 1 Then
                            B = Int(Val(Left$(St, A - 1))) - IncludeLineCount - 1
                            St = Mid$(St, A + 1)
                            A = SendMessage(txtCode.hwnd, EM_LINEINDEX, B, 0)
                            If A >= 0 Then
                                B = SendMessage(txtCode.hwnd, EM_LINELENGTH, A, 0)
                                txtCode.SelStart = A
                                txtCode.SelLength = B
                                Scintilla.SelStart = A
                                Scintilla.SelEnd = A + B
                                Scintilla.GotoLine Scintilla.GetCurrentLine
                                Scintilla.SelectLine
                            End If
                        End If
                        MsgBox St, vbOKOnly + vbExclamation, TitleString
                    End If
                Else
                    MsgBox "Error: script.log not found!", vbOKOnly + vbExclamation, TitleString
                End If
            Else
                MsgBox "Unable to execute mbsc.exe!", vbOKOnly + vbExclamation, TitleString
            End If

            Kill "COMPILE.BAT"
            Kill "CLOSE.COM"
        Else
            MsgBox "File 'mbsc.inc' not found!", vbOKOnly + vbExclamation, TitleString
        End If
    Else
        MsgBox "Unable to execute mbsc.exe!", vbOKOnly + vbExclamation, TitleString
    End If

    Me.Caption = "The Odyssey Online Classic [Editing Script]"
End Sub

Private Sub Form_Load()
    Scintilla.InitScintilla Me.hwnd
    SCIHighlighter.LoadHighlighters App.Path
    SCIHighlighter.SetHighlighter frmScript.Scintilla, "VB"
    Scintilla.MoveSCI 0, 25, Me.ScaleWidth, Me.ScaleHeight - 1000
End Sub

Private Sub mnuCopy_Click()
    Scintilla.Copy
End Sub

Private Sub mnuFind_Click()
    Scintilla.DoFind
End Sub

Private Sub mnuFindNext_Click()
    Scintilla.FindNext
End Sub

Private Sub mnuFindPrevious_Click()
    Scintilla.FindPrev
End Sub

Private Sub mnuGoto_Click()
    Scintilla.DoGoto
End Sub

Private Sub mnuPaste_Click()
    Scintilla.Paste
End Sub

Private Sub mnuRedo_Click()
    Scintilla.Redo
End Sub

Private Sub mnuReplace_Click()
    Scintilla.DoReplace
End Sub

Private Sub mnuSelectAll_Click()
    Scintilla.SelectAll
End Sub

Private Sub mnuUndo_Click()
    Scintilla.Undo
End Sub

Private Function CheckScriptInvalid(ScriptText As String) As Boolean
    Dim ScriptLength As Long, A As Long, B As Long
    ScriptLength = Len(ScriptText)
    If ScriptLength < 8000 Then
        If UCase$(Left$(ScriptText, 3)) = "SUB" Then
            A = InStr(1, UCase$(ScriptText), "EXIT FUNCTION")
            If A > 0 Then
                B = SendMessage(txtCode.hwnd, EM_LINELENGTH, A, 0)
                txtCode.SelStart = A
                txtCode.SelLength = B
                Scintilla.SelStart = A
                Scintilla.SelEnd = A + B
                Scintilla.GotoLine Scintilla.GetCurrentLine - 1
                Scintilla.SelectLine
                MsgBox "You cannot have Exit Function inside a Sub!"
                CheckScriptInvalid = True
                Exit Function
            End If
        ElseIf UCase$(Left$(ScriptText, 8)) = "FUNCTION" Then
            A = InStr(1, UCase$(ScriptText), "EXIT SUB")
            If A > 0 Then
                B = SendMessage(txtCode.hwnd, EM_LINELENGTH, A, 0)
                txtCode.SelStart = A
                txtCode.SelLength = B
                Scintilla.SelStart = A
                Scintilla.SelEnd = A + B
                Scintilla.GotoLine Scintilla.GetCurrentLine - 1
                Scintilla.SelectLine
                MsgBox "You cannot have Exit Sub inside a Function!"
                CheckScriptInvalid = True
                Exit Function
            End If
        End If

        Dim Going As Boolean
        Dim TestString As String
        Dim FindString As String
        FindString = ScriptText
        TestString = Chr$(34)

        Going = True
        A = 1
        Do While Going = True
            A = InStr(A, FindString, "/")
            If A > 0 Then
                B = InStrCount(FindString, TestString, 1, A)
                If B = 0 Or Not B Mod 2 = 1 Then
                    B = SendMessage(txtCode.hwnd, EM_LINELENGTH, A, 0)
                    txtCode.SelStart = A
                    txtCode.SelLength = B
                    Scintilla.SelStart = A
                    Scintilla.SelEnd = A + B
                    Scintilla.GotoLine Scintilla.GetCurrentLine - 1
                    Scintilla.SelectLine
                    MsgBox "Use of the / symbol is not allowed.  Please use the Divide function instead."
                    CheckScriptInvalid = True
                    Exit Function
                End If
            Else
                Going = False
            End If
            A = A + 1
        Loop
    Else
        MsgBox "This script is too big!  The maximum size is 8000 characters, and this script is " + CStr(ScriptLength) + " characters!"
        CheckScriptInvalid = True
        Exit Function
    End If
End Function
